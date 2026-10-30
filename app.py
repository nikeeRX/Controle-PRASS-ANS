# -*- coding: utf-8 -*-
import os
import io
import time
import json
import sqlite3
import traceback
from datetime import datetime
from flask import (
    Flask, request, jsonify, send_file, render_template_string
)

# ==============================================================================
# IMPORTS BLINDADOS (GARANTEM QUE O WORKER DO GUNICORN NUNCA FALHE NO BOOT)
# ==============================================================================
try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
    HAS_PSYCOPG2 = True
except Exception as e:
    HAS_PSYCOPG2 = False
    print(f"[AVISO BOOT] psycopg2 indisponivel: {e}")

try:
    import pypdf
    HAS_PYPDF = True
except Exception as e:
    HAS_PYPDF = False
    print(f"[AVISO BOOT] pypdf indisponivel: {e}")

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, HRFlowable
    )
    HAS_REPORTLAB = True
except Exception as e:
    HAS_REPORTLAB = False
    print(f"[AVISO BOOT] reportlab indisponivel: {e}")

try:
    from PIL import Image
    HAS_PIL = True
except Exception as e:
    HAS_PIL = False
    print(f"[AVISO BOOT] PIL indisponivel: {e}")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # Limite de 64MB para uploads


# ==============================================================================
# DADOS FIXOS INTEGRAIS EXTRAÍDOS DA PLANILHA 'PRASS - Projeto - base inicial.xlsx'
# ==============================================================================
DEMANDAS_FIXAS_PADRAO = [
    "Status 08/2026",
    "Status 09/2026",
    "Status 10/2026",
    "Complemento ANS 26/10",
    "Status 11/2026",
    "Status 12/2026",
    "Status 01/2027",
    "Status 02/2027",
    "Status 03/2027",
    "Status 04/2027",
    "Status 05/2027",
    "Status 06/2027",
    "Status 07/2027",
    "Status 08/2027",
]

ITENS_PRASS_INICIAIS = [
    {
        "codigo": "8.1.2",
        "titulo": "8.1.2 Falhas no processo de fornecimento de reembolsos",
        "areas_atribuidas": ["CODER", "COCAP", "CONEC"],
        "apontamento_ans": (
            "Apontamento da ANS\n"
            "A Nota Técnica nº 39/2026 identificou, no âmbito da Anormalidade A, falhas relacionadas ao processo de "
            "fornecimento de reembolsos, abrangendo situações de demora na comunicação para solução de pendências "
            "documentais, extrapolação de prazos, pagamentos de valores incorretos e negativas indevidas, inclusive em "
            "situações relacionadas à insuficiência de rede prestadora.\n"
            "A análise da Agência também posicionou a Postal Saúde na 19ª colocação no indicador de Falhas de Reembolso, "
            "registrando aumento das demandas relacionadas ao tema no período analisado.\n\n"
            "Diagnóstico institucional\n"
            "A Postal Saúde dispõe de fluxo formal para o processamento das solicitações de reembolso, contemplando "
            "recebimento, conferência documental, análise, enquadramento, processamento e pagamento. O POP009 – Prover "
            "Reembolso Assistencial – Versão 4 estabelece as etapas do processo, os documentos necessários e os "
            "procedimentos aplicáveis às solicitações de reembolso.\n"
            "O procedimento estabelece prazo de até 30 dias corridos para o processamento do reembolso, contado do "
            "recebimento da documentação completa, além de prever tratamento específico para situações de documentação "
            "incompleta ou inconsistente.\n"
            "No acompanhamento operacional realizado pela área responsável, foram identificados protocolos pendentes, "
            "inclusive solicitações que ultrapassavam o prazo de referência. A situação foi formalmente monitorada e "
            "compartilhada com os gestores responsáveis e com a área financeira para adoção das providências pertinentes.\n"
            "A análise institucional também identificou situações em que determinados prestadores de serviços assistenciais "
            "solicitam indevidamente ao beneficiário o pagamento direto por serviços que deveriam observar o fluxo regular "
            "de cobertura e faturamento da Operadora. Essa prática pode contribuir para a geração de solicitações de "
            "reembolso e constitui fator de risco que demanda tratamento sob as perspectivas assistencial, operacional, "
            "ética e de compliance.\n"
            "Em resposta a esse cenário, a Postal Saúde instaurou internamente Comitê de Procedimento Ético, destinado à "
            "análise de situações que possam envolver condutas inadequadas ou indícios de fraude relacionados aos "
            "prestadores assistenciais, observados os procedimentos internos aplicáveis.\n"
            "Paralelamente, encontra-se em desenvolvimento, em conjunto com a área de Comunicação, campanha educativa "
            "dirigida aos prestadores assistenciais, destinada a esclarecer as condutas inadequadas relacionadas à cobrança "
            "direta do beneficiário e reforçar os procedimentos que devem ser observados na relação com a Operadora e seus "
            "beneficiários.\n"
            "Também está em desenvolvimento Código de Conduta Ética para Prestadores Assistenciais, com a finalidade de "
            "estabelecer, de forma objetiva, os padrões de conduta esperados da rede prestadora, fortalecendo os mecanismos "
            "de orientação, prevenção, integridade e responsabilização.\n"
            "As iniciativas integram o processo de fortalecimento do Compliance da Postal Saúde, ampliando a abordagem "
            "preventiva sobre os riscos relacionados à atuação da rede assistencial."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• A Postal Saúde vem realizando o saneamento das pendências identificadas no processo de reembolso, sob "
            "condução da área responsável, com acompanhamento dos protocolos e atuação direcionada sobre os casos que "
            "ultrapassaram o prazo estabelecido.\n"
            "• O processo encontra-se em fase de aperfeiçoamento integrado com a área Financeira, mediante revisão dos "
            "fluxos e das interfaces relacionadas ao processamento e pagamento dos reembolsos.\n"
            "• Em conjunto com a área de Tecnologia da Informação, estão sendo desenvolvidas medidas de automação "
            "destinadas a reduzir intervenções manuais, ampliar a rastreabilidade das etapas e fortalecer os controles do processo.\n"
            "• No âmbito da rede prestadora, encontram-se em implementação as iniciativas relacionadas ao Comitê de "
            "Procedimento Ético, à apuração de situações que possam envolver condutas inadequadas ou indícios de fraude e "
            "ao fortalecimento dos mecanismos de orientação e prevenção.\n"
            "• Também estão em desenvolvimento, em conjunto com a área de Comunicação, a campanha educativa dirigida aos "
            "prestadores assistenciais e o Código de Conduta Ética para Prestadores Assistenciais.\n"
            "• As pendências atualmente identificadas estão sendo tratadas pela área responsável, com previsão de "
            "saneamento até outubro de 2026, concomitantemente ao avanço das medidas de aperfeiçoamento processual e tecnológico."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "O tratamento da matéria será acompanhado por indicadores capazes de demonstrar a redução do passivo e a "
            "evolução do processo, contemplando:\n"
            "• estoque de solicitações pendentes;\n"
            "• quantidade e percentual de processos fora do prazo;\n"
            "• percentual de solicitações processadas dentro do prazo;\n"
            "• tempo médio de processamento;\n"
            "• pendências documentais;\n"
            "• negativas e respectivas reanálises;\n"
            "• divergências de valores;\n"
            "• tempo entre processamento e pagamento;\n"
            "• NIPs relacionadas a reembolso;\n"
            "• ocorrências relacionadas à conduta de prestadores;\n"
            "• evolução do indicador de Falhas de Reembolso.\n\n"
            "O acompanhamento será realizado pelas áreas responsáveis, com reporte à governança competente e tratamento "
            "dos desvios identificados. A conclusão do ciclo de saneamento previsto para outubro de 2026 será acompanhada "
            "da avaliação dos resultados alcançados e da efetividade das melhorias implementadas, especialmente quanto à "
            "tempestividade, rastreabilidade e prevenção de recorrências."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED\n"
                "Em atenção às informações contidas no item medidas implementadas e em implementação, seguem nossas considerações:\n"
                "• Monitoramento dos processos de Garantia de Atendimento e Reembolso: GERED/CODER é responsável pelo monitoramento // A execução é de responsabilidade da GEREL - UAR/NAR;\n"
                "• Processamento dos Reembolsos Assistenciais e efetivo pagamento aos beneficiários: Responsabilidade da GEFAC/COANF e GECOF/COPAG, fazendo-se necessária a revisão dos fluxos e normativos pelas respectivas áreas;\n"
                "• Recursos Tecnológicos: Em andamento a solicitação para que no APP, quando o beneficiário solicitar Reembolso Integral, seja obrigatório informar o protocolo de atendimento;\n"
                "• Rede Assistencial (Prestadores): Está sob a gestão da CONEC e os processos estão sendo revisados visando garantir o alinhamento com as diretrizes de compliance.\n\n"
                "Processo de Garantia de Atendimento e Reembolso (Agosto/26):\n"
                "- Garantia de Atendimento: Dentro do Prazo: 236 | Fora do Prazo: 1.631 | Total: 1.867 (+27,70%)\n"
                "- Reembolso: Dentro do Prazo: 361 | Fora do Prazo: 715 | Total: 1.076 (+2,09%)\n"
                "Infere-se que a variação no volume de processos em agosto/26 ocorre devido à sazonalidade do período de férias escolares. Paralelamente, a GERED/CODER tem atuado fortemente no monitoramento e auxílio às UAR/NAR com maior volume de processos."
            ),
            "Status 09/2026": (
                "GERED\n"
                "Em atenção às informações contidas no item medidas implementadas e em implementação, seguem nossas considerações:\n"
                "• Monitoramento dos processos de Garantia de Atendimento e Reembolso: GERED/CODER é responsável pelo monitoramento // Execução: GEREL - UAR/NAR;\n"
                "• Recursos Tecnológicos: Em andamento a solicitação para obrigatoriedade do protocolo de atendimento no APP para Reembolso Integral;\n"
                "• Rede de prestadores: Sob gestão da CONEC com processos em revisão.\n\n"
                "Evolução Ago/26 a Set/26 (Redução de pendências):\n"
                "- Garantia de Atendimento: Ago/26 Total: 1.867 -> Set/26: Dentro do Prazo: 200 | Fora do Prazo: 1.370 | Total: 1.570 (-15,91% de redução)\n"
                "- Reembolso: Ago/26 Total: 1.076 -> Set/26: Dentro do Prazo: 444 | Fora do Prazo: 442 | Total: 886 (-17,66% de redução)\n"
                "Decorrido o período de férias e como reflexo do monitoramento mais extensivo da GERED/CODER, é perceptível a redução dos protocolos de garantia, especialmente nos processos fora do prazo."
            ),
            "Status 10/2026": (
                "GERED/CODER\n"
                "Em atenção às informações contidas no item medidas implementadas e em implementação, seguem nossas considerações:\n"
                "• Monitoramento dos processos de Garantia de Atendimento e Reembolso: GERED/CODER é responsável pelo monitoramento // Execução: GEREL - UAR/NAR;\n"
                "• Recursos Tecnológicos: Foi implementada em 01/10/2026 a obrigatoriedade de o beneficiário informar o número do protocolo de atendimento quando solicitar Reembolso Integral no aplicativo. Essa medida visa a rastreabilidade dos processos de garantia de atendimento liberados como Reembolso Integral. Caso o beneficiário não informe o protocolo, poderá solicitar o Reembolso na modalidade Livre Escolha.\n\n"
                "Processo de Garantia de Atendimento e Reembolso Integral:\n"
                "- Garantia de Atendimento: Ago/26 (1.867) -> Set/26 (1.570, redução de -15,91%)\n"
                "- Reembolso: Ago/26 (1.076) -> Set/26 (886, redução de -17,66%)"
            ),
            "Complemento ANS 26/10": ""
        }
    },
    {
        "codigo": "8.3.3",
        "titulo": "8.3.3 Demora na liberação de procedimentos e consultas em conjunto em função de insuficiência de rede prestadora",
        "areas_atribuidas": ["CODER", "COCAP", "CONEC"],
        "apontamento_ans": (
            "Apontamento da ANS\n"
            "A Nota Técnica nº 39/2026 aponta como anormalidade administrativa grave a demora na liberação de procedimentos "
            "e consultas em conjunto em decorrência de insuficiência de rede prestadora.\n\n"
            "Diagnóstico institucional\n"
            "A Postal Saúde possui processo formal de Garantia de Atendimento destinado justamente ao tratamento de "
            "situações de inexistência ou indisponibilidade de rede assistencial.\n"
            "O Manual define indisponibilidade como a situação em que existe prestador para o serviço ou procedimento "
            "demandado, mas este não se encontra disponível dentro dos prazos estabelecidos pela ANS. Define, ainda, "
            "inexistência como a ausência de prestador apto a oferecer o serviço ou procedimento no município demandado.\n"
            "Nessas situações, a Postal Saúde deve disponibilizar a Garantia de Atendimento em município limítrofe ou "
            "Região de Saúde, observando as regras e os prazos regulatórios.\n"
            "O Manual determina que sejam esgotadas, de forma ordenada, as alternativas assistenciais:\n"
            "1. rede credenciada;\n"
            "2. contratação direta;\n"
            "3. Tratamento Fora do Domicílio;\n"
            "4. reembolso integral.\n"
            "Também determina que as negociações de contratação direta sejam formalizadas e anexadas ao sistema para fins "
            "de evidência.\n"
            "Portanto, a insuficiência de rede é reconhecida no próprio processo como evento de risco a ser tratado "
            "mediante alternativas assistenciais, e não como justificativa para postergação indefinida do atendimento."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• A Postal Saúde mantém processo estruturado para gestão da Garantia de Atendimento, com atribuição formal da "
            "área responsável pela gestão estratégica da rede para assegurar sua composição assistencial, identificar "
            "vulnerabilidades e estabelecer os processos de Contratação Direta, TFD e Reembolso Assistencial.\n"
            "• O Manual estabelece que a busca por atendimento considere rede direta, rede indireta, parcerias e "
            "Telemedicina, seguida, quando necessário, de contratação direta, TFD e reembolso integral.\n"
            "• A contratação direta exige formalização da negociação e manutenção das evidências no sistema.\n"
            "• Para TFD, o Manual determina que o atendimento seja garantido conforme o prazo da ANS e que sejam "
            "providenciadas as autorizações necessárias para as despesas relacionadas ao deslocamento.\n"
            "• Esses mecanismos constituem medidas concretas para impedir que uma deficiência localizada de rede resulte "
            "automaticamente em desassistência."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "A operacionalização da Garantia de Atendimento prevê:\n"
            "• acolhimento da solicitação;\n"
            "• análise da cobertura;\n"
            "• verificação da rede credenciada;\n"
            "• comprovação sistêmica das tentativas de agendamento;\n"
            "• definição da alternativa assistencial;\n"
            "• autorização de reembolso integral, quando aplicável; e\n"
            "• comunicação ao beneficiário.\n\n"
            "O Manual estabelece ainda que o retorno ao beneficiário observe os prazos máximos da ANS e que os registros "
            "dos contatos, tentativas de agendamento, negativas e protocolos sejam mantidos de forma integral e organizada.\n"
            "O monitoramento da rede deve permitir identificar vulnerabilidades e orientar ações de composição ou "
            "recomposição da rede assistencial.\n"
            "Eventual demora decorrente de insuficiência de rede será analisada sob duas perspectivas:\n"
            "1. causa assistencial: indisponibilidade ou inexistência efetiva do prestador;\n"
            "2. resposta operacional: tempo e adequação das providências adotadas pela Operadora para assegurar alternativa assistencial."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED\n"
                "Em atenção às informações contidas neste item, seguem nossas considerações:\n"
                "Abaixo, listamos os normativos vigentes da Garantia de Atendimento:\n"
                "• MAN 037 (versão 3) – publicado em 11/08/2026\n"
                "• POP 009 (versão 4) – publicado em 13/08/2026\n"
                "• POP 012 (versão 1) – publicado em 13/08/2026\n"
                "• POP 013 (versão 0) – publicado em 13/08/2026\n"
                "• IOT 222 (versão 1) – publicado em 02/08/2021\n"
                "• IOT 224 (versão 1) – publicado em 23/02/2021\n"
                "Atualmente, as instruções normativas IOT 222 e IOT 224 estão em processo de revisão para alinhamento ao cenário vigente. Paralelamente, solicitamos subsídios às UAR/NAR a fim de mapear e atualizar as informações sobre os prestadores ativos no sistema. Em complemento, o painel de análise de vulnerabilidade está em processo de revisão."
            ),
            "Status 09/2026": (
                "GERED\n"
                "Atualmente, as instruções normativas IOT 222 e IOT 224 estão em processo de revisão para alinhamento ao cenário vigente. Paralelamente, solicitamos subsídios às UAR/NAR a fim de mapear e atualizar as informações sobre os prestadores ativos no sistema. Em complemento à ação de atualização dos normativos, o painel de análise de vulnerabilidade está em processo de revisão, tanto na metodologia de extração dos dados quanto na sua elaboração."
            ),
            "Status 10/2026": (
                "GERED/CODER\n"
                "No que diz respeito às ações da GERED/CODER, registramos que o painel de análise de vulnerabilidade encontra-se em processo de revisão, abrangendo a metodologia de extração, tratamento e análise dos dados utilizados em sua construção. Em decorrência das avaliações realizadas, verificou-se a necessidade de reestruturação do modelo atualmente adotado, tendo sido formalizada, em 08/10/2026, demanda junto à área de Tecnologia da Informação para desenvolvimento e institucionalização do painel de vulnerabilidade."
            ),
            "Complemento ANS 26/10": (
                "GERED/CODER — Subitem: Vulnerabilidade de Rede\n"
                "No que diz respeito às ações da GERED/CODER, registramos que o painel de análise de vulnerabilidade encontra-se em processo de revisão, abrangendo a metodologia de extração, tratamento e análise dos dados utilizados em sua construção. Foi formalizada, em 08/10/2026, demanda junto à área de Tecnologia da Informação para desenvolvimento e institucionalização do painel de vulnerabilidade."
            )
        }
    },
    {
        "codigo": "8.4.1",
        "titulo": "8.4.1 Insuficiência/inexistência de rede prestadora para a realização de consultas e procedimentos de cobertura obrigatória, acarretando desassistência ou pagamento particular para posterior reembolso",
        "areas_atribuidas": ["CODER", "CONEC"],
        "apontamento_ans": (
            "Apontamento da ANS\n"
            "A Nota Técnica nº 39 aponta, como anormalidade administrativa grave de natureza assistencial, a insuficiência "
            "ou inexistência de rede prestadora para realização de consultas e procedimentos de cobertura obrigatória, "
            "quando essa condição acarreta desassistência ou obriga o beneficiário a realizar pagamento particular para "
            "posterior reembolso, inclusive diante de irregularidades na efetivação desse reembolso.\n\n"
            "Diagnóstico institucional\n"
            "A Postal Saúde reconhece que a disponibilidade e a suficiência da rede assistencial constituem fatores "
            "diretamente relacionados à garantia de acesso e à continuidade da assistência.\n"
            "A Operadora dispõe de processo formal de Garantia de Atendimento, que contempla acolhimento da demanda, "
            "análise de cobertura, verificação da rede disponível, comprovação das tentativas de agendamento, definição da "
            "alternativa assistencial e comunicação ao beneficiário. Nas situações de inexistência ou indisponibilidade de "
            "prestador, o processo prevê a utilização de alternativas assistenciais destinadas a preservar o atendimento, "
            "inclusive contratação direta, Tratamento Fora do Domicílio e, quando cabível, reembolso integral.\n"
            "No que se refere aos casos em que o beneficiário realiza pagamento particular, o tratamento do respectivo "
            "reembolso está diretamente relacionado ao processo analisado no item 8.1.2 deste PRASS."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• A Postal Saúde vem adotando medidas destinadas ao fortalecimento da rede e à redução das vulnerabilidades "
            "relacionadas à suficiência assistencial.\n"
            "• Entre essas medidas está a implementação do Compartilhamento de Gestão de Riscos e Rede Assistencial, "
            "concebido para ampliar a disponibilidade e a capilaridade da rede, especialmente em localidades nas quais "
            "foram identificadas vulnerabilidades ou necessidade de complementação da estrutura assistencial.\n"
            "• O modelo contempla etapas de identificação da necessidade, avaliação técnica, implantação, operação, "
            "monitoramento, avaliação de resultados e melhoria contínua.\n"
            "• Paralelamente, permanecem disponíveis os mecanismos previstos no processo de Garantia de Atendimento para "
            "situações de inexistência ou indisponibilidade de prestadores.\n"
            "• Quanto ao reembolso, as medidas de saneamento e aperfeiçoamento estão sendo conduzidas conforme descrito no item 8.1.2."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "O tratamento do risco ocorre em duas dimensões complementares:\n"
            "I – Gestão da rede: identificação de insuficiências e vulnerabilidades, recomposição e ampliação da rede e "
            "utilização de mecanismos de compartilhamento e contratação alternativa;\n"
            "II – Garantia individual de acesso: adoção das alternativas assistenciais previstas quando identificada "
            "inexistência ou indisponibilidade de prestador.\n\n"
            "O processo prevê registro das tentativas de agendamento e das providências adotadas, possibilitando "
            "rastreabilidade da atuação institucional.\n"
            "As ocorrências relacionadas a pagamento particular e posterior reembolso permanecem vinculadas ao "
            "monitoramento específico do processo de reembolso, permitindo a análise da cadeia completa:\n"
            "indisponibilidade de rede → alternativa assistencial → eventual pagamento particular → reembolso → encerramento da demanda."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED\n"
                "Em atenção às informações contidas neste item, ratificamos as informações constantes no Diagnóstico Institucional:\n"
                "1) Garantia de atendimento, reembolso, medidas de saneamento e aperfeiçoamento: conduzidas conforme descrito no item 8.1.2;\n"
                "2) Gestão da Rede Assistencial (Direta/Indireta): sob a gestão da CONEC, com processos em revisão para alinhamento às diretrizes de compliance;\n"
                "3) Compartilhamento de Gestão de Riscos (Encapsulamento): sob responsabilidade da GEREL."
            ),
            "Status 09/2026": (
                "GERED\n"
                "Ratificamos as informações constantes no Diagnóstico Institucional:\n"
                "1) Garantia de atendimento e reembolso conduzidos conforme descrito no item 8.1.2;\n"
                "2) Gestão da Rede Assistencial (Direta/Indireta) sob gestão da CONEC em revisão normativa;\n"
                "3) Compartilhamento de Gestão de Riscos sob responsabilidade da GEREL."
            ),
            "Status 10/2026": (
                "GERED/CODER\n"
                "Implementada em 01/10/2026 a obrigatoriedade de o beneficiário informar o número do protocolo de atendimento quando solicitar Reembolso Integral no aplicativo, garantindo rastreabilidade dos processos de Garantia de Atendimento. Redução acumulada de -15,91% nos processos de Garantia de Atendimento e de -17,66% nos processos de Reembolso entre agosto e setembro/2026."
            ),
            "Complemento ANS 26/10": ""
        }
    },
    {
        "codigo": "8.4.2",
        "titulo": "8.4.2 Relatos de suspensão de atendimentos, inclusive de prestadores hospitalares, em função de inadimplência por parte da Operadora",
        "areas_atribuidas": ["CODER", "COCAP", "CONEC"],
        "apontamento_ans": (
            "Apontamento da ANS\n"
            "A Nota Técnica nº 39 aponta relatos de suspensão de atendimento, inclusive por prestadores hospitalares, em "
            "decorrência de inadimplência da Operadora.\n\n"
            "Diagnóstico institucional\n"
            "A Postal Saúde reconhece que o período analisado foi marcado por relevante restrição financeira decorrente da "
            "irregularidade dos repasses da mantenedora, circunstância que produziu impactos sobre a capacidade de "
            "cumprimento tempestivo de obrigações financeiras perante parte da rede prestadora.\n"
            "A documentação institucional demonstra que a restrição financeira repercutiu sobre o relacionamento com "
            "prestadores, incluindo situações de inadimplência, suspensão de atendimento e descredenciamento.\n"
            "A Operadora, portanto, não trata o fato histórico como inexistente ou meramente pontual. O diagnóstico "
            "institucional considera que houve efetiva exposição ao risco de desassistência decorrente da combinação entre "
            "restrição financeira e dependência da rede prestadora.\n"
            "A questão central, entretanto, passou a ser a adoção de medidas capazes de reduzir a vulnerabilidade "
            "assistencial decorrente desse cenário."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• Como resposta estruturante, a Postal Saúde implementou o Compartilhamento de Gestão de Riscos e Rede "
            "Assistencial, utilizando redes disponibilizadas por operadoras parceiras.\n"
            "• A medida foi concebida para ampliar a disponibilidade de prestadores, aumentar a capilaridade da rede, "
            "reduzir vulnerabilidades assistenciais e proporcionar alternativas diante de eventual indisponibilidade da rede direta.\n"
            "• A documentação relativa ao projeto demonstra que sua implementação foi precedida de análises assistenciais, "
            "regulatórias, operacionais, atuariais, econômico-financeiras e jurídicas, permitindo avaliar a estratégia sob "
            "perspectiva multidimensional.\n"
            "• O modelo também contempla ciclo de implantação, operação, monitoramento, avaliação de resultados e melhoria contínua."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "O risco passou a ser tratado de forma integrada entre as dimensões financeira, assistencial, regulatória e de "
            "gestão da rede.\n"
            "O compartilhamento de rede e risco não é considerado substituto da responsabilidade da Postal Saúde perante "
            "seus beneficiários. Ao contrário, constitui mecanismo de mitigação destinado a ampliar a capacidade "
            "assistencial e reduzir a exposição decorrente de eventual descontinuidade de prestadores.\n"
            "O monitoramento contempla a suficiência da rede, a continuidade da assistência, os indicadores assistenciais, "
            "reclamações, NIPs e demais sinais de risco relacionados à disponibilidade de prestadores.\n"
            "A estratégia também incorpora a necessidade de planejamento de contingência para eventual descontinuidade das "
            "redes compartilhadas, evitando a criação de nova dependência sem mecanismo de mitigação."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED\n"
                "No que se refere aos itens de controle de rede, seguem nossos esclarecimentos:\n"
                "• Descredenciamento ou restrição de atendimento: Aplicação das IOTs 031, 036 e 068, avaliação de substituição/expansão da rede e fortalecimento da rastreabilidade por meio de painel gerencial;\n"
                "• Estudos de suficiência e capilaridade da rede: Contemplados no painel de análise de vulnerabilidade em revisão;\n"
                "• Mecanismos de monitoramento assistencial e regulatório: Conduzidos pelaGERED/CODER conforme item 8.1.2.\n"
                "GEREM: Ampliada a rede assistencial por meio de convênios de reciprocidade/compartilhamento de risco (AMIL RJ, UNIMED CNU e UNIMED CENTRO OESTE PAULISTA)."
            ),
            "Status 09/2026": (
                "GERED\n"
                "Manutenção do acompanhamento das ocorrências de suspensão, revisão metodológica do painel de vulnerabilidade pela CODER, aplicação das IOTs 031, 036 e 068 e vigência das parcerias de compartilhamento de rede (AMIL RJ, UNIMED CNU e UNIMED CENTRO OESTE PAULISTA)."
            ),
            "Status 10/2026": (
                "GERED/CODER\n"
                "Formalizada em 08/10/2026 demanda junto à área de Tecnologia da Informação para desenvolvimento e institucionalização do painel de vulnerabilidade, visando aprimorar a qualidade das informações e fortalecer a rastreabilidade das bases utilizadas."
            ),
            "Complemento ANS 26/10": (
                "GERED/CODER — Subitem: Garantia de Atendimento em virtude das suspensões de atendimento\n"
                "Foram adotadas as medidas imediatas necessárias para assegurar a continuidade da assistência aos beneficiários e mitigar os impactos das suspensões contratuais, viabilizando o direcionamento do fluxo assistencial para prestadores aptos da rede credenciada, disponibilização de prestadores (credenciados ou não) ou liberação de reembolso integral nos prazos da RN nº 566/2022 da ANS.\n\n"
                "GERED/CONEC — Subitem: Suspensão de atendimento por inadimplência\n"
                "Em atendimento ao Ofício nº 108/2026 da ANS, apresentamos os procedimentos adotados para:\n"
                "1. Registro e controle das suspensões de atendimento;\n"
                "2. Monitoramento gerencial das ocorrências;\n"
                "3. Atuação junto aos prestadores e formalização de acordos de quitação;\n"
                "4. Evidências documentais."
            )
        }
    },
    {
        "codigo": "8.5",
        "titulo": "8.5. NÃO DISPONIBILIZAÇÃO DE FORMA CLARA DA REDE AOS BENEFICIÁRIOS NOS TERMOS DA REGULAMENTAÇÃO EM VIGOR",
        "areas_atribuidas": ["CODER", "COCAP", "CONEC"],
        "apontamento_ans": (
            "Apontamento da ANS\n"
            "A ANS apontou situação relacionada à não disponibilização, de forma clara, da rede assistencial aos "
            "beneficiários, nos termos da regulamentação vigente, abrangendo aspectos relacionados à atualização das "
            "informações da rede, à sua divulgação no portal corporativo e à comunicação das substituições de prestadores "
            "de serviços de atenção à saúde.\n\n"
            "Diagnóstico institucional\n"
            "A Postal Saúde possui estrutura formal de governança e gestão da rede assistencial, sob responsabilidade da "
            "Gerência Executiva de Rede (GERED), com processos definidos para credenciamento, recredenciamento, "
            "descredenciamento, substituição, suspensão, higienização cadastral e monitoramento da vulnerabilidade da rede.\n"
            "A rede assistencial apresenta dinâmica permanente de expansão, alteração e reorganização. Em dezembro de 2024, "
            "a Operadora contava com 12.781 prestadores credenciados na rede direta. Entre janeiro de 2025 e julho de 2026, "
            "foram incorporados 305 novos prestadores, alcançando 13.086 prestadores na rede direta, além de 27.100 "
            "prestadores disponibilizados pela rede indireta (CNU), totalizando 40.186 prestadores.\n"
            "No mesmo período, foram registrados 580 descredenciamentos, decorrentes de diferentes situações, incluindo "
            "ausência de atendimento, solicitação do prestador, interesse da Operadora, alteração de CNPJ, encerramento "
            "das atividades e migração de tipo de pessoa.\n"
            "Nesse contexto, a Postal Saúde vem fortalecendo a integração entre a gestão operacional da rede, os controles "
            "cadastrais, os sistemas corporativos e os mecanismos de disponibilização das informações, com foco na "
            "tempestividade, consistência e confiabilidade dos dados apresentados aos beneficiários."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• Manutenção de estrutura centralizada de gestão da rede assistencial pela GERED, com acompanhamento das "
            "movimentações cadastrais e contratuais.\n"
            "• Utilização do Painel de Vulnerabilidade da Rede Prestadora, que consolida informações relacionadas à "
            "suficiência, distribuição geográfica, disponibilidade, cobertura assistencial e potenciais riscos de desassistência.\n"
            "• Monitoramento contínuo dos credenciamentos, descredenciamentos, reativações, suspensões e demais alterações da rede.\n"
            "• Aplicação de normativos e procedimentos internos específicos para credenciamento, descredenciamento, "
            "recontratualização, higienização cadastral e suspensão de atendimento.\n"
            "• Realização de análises de vulnerabilidade para subsidiar decisões de expansão, substituição e reorganização da rede assistencial.\n"
            "• Fortalecimento dos controles sistêmicos e automatizados, com iniciativas voltadas à rastreabilidade e integração das informações.\n"
            "• Revisão contínua dos processos operacionais relacionados à gestão da rede, com adoção do ciclo de melhoria "
            "contínua para identificação e tratamento de eventuais inconsistências."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "O apontamento está sendo tratado no âmbito do processo permanente de governança, monitoramento e melhoria "
            "contínua da rede assistencial.\n"
            "As movimentações da rede são acompanhadas pela GERED por meio dos sistemas corporativos e do Painel de "
            "Vulnerabilidade, permitindo identificar alterações relevantes na composição, disponibilidade e suficiência da "
            "rede e subsidiar as decisões de credenciamento, substituição, expansão ou reorganização.\n"
            "Os processos de descredenciamento e suspensão de atendimento também são submetidos a avaliação quanto aos "
            "seus potenciais impactos assistenciais.\n"
            "Paralelamente, a Operadora vem aprimorando os controles sistêmicos e os fluxos operacionais destinados a "
            "assegurar maior consistência entre a situação efetiva da rede, as bases corporativas e as informações "
            "disponibilizadas aos beneficiários."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED\n"
                "• Ações CODER: Painel de Vulnerabilidade da Rede Prestadora para subsidiar decisões de expansão, substituição e reorganização.\n"
                "• Ações CONEC: Revisão dos normativos de credenciamento, descredenciamento, recontratualização, higienização e suspensão.\n"
                "• Ações COCAP: Disponibilização da rede no portal refletindo a base oficial, tratamento de notificações de inconsistências e publicação de substituições.\n"
                "Composição da Rede (Ago/2026): Rede Direta: 12.399 | Rede Indireta: 30.194 | Total: 42.593 prestadores."
            ),
            "Status 09/2026": (
                "GERED (CODER, CONEC e COCAP)\n"
                "Continuidade da revisão do painel de vulnerabilidade (CODER), revisão normativa (CONEC) e atualização contínua do portal e tratamento de inconsistências (COCAP).\n"
                "Composição da Rede (Set/2026): Rede Direta: 12.357 | Rede Indireta: 30.268 | Total: 42.625 prestadores."
            ),
            "Status 10/2026": (
                "GERED/CODER\n"
                "Formalizada em 08/10/2026 demanda junto à área de Tecnologia da Informação para desenvolvimento e institucionalização do painel de vulnerabilidade, visando aprimorar a qualidade das informações e fortalecer a rastreabilidade das bases."
            ),
            "Complemento ANS 26/10": (
                "GERED/CODER — Subitem: Vulnerabilidade de Rede\n"
                "O painel de análise de vulnerabilidade encontra-se em processo de revisão metodológica de extração, tratamento e análise dos dados, com demanda formalizada em 08/10/2026 junto à TI para institucionalização da ferramenta."
            )
        }
    },
    {
        "codigo": "8.5.1",
        "titulo": "8.5.1. RELATOS DE DESATUALIZAÇÃO DA REDE EFETIVAMENTE DISPONIBILIZADA PELA OPERADORA",
        "areas_atribuidas": ["CODER", "COCAP", "CONEC"],
        "apontamento_ans": (
            "Apontamento da ANS\n"
            "Foram identificados relatos relacionados à desatualização da rede efetivamente disponibilizada pela Operadora, "
            "situação que pode comprometer a adequada orientação dos beneficiários quanto aos prestadores disponíveis para atendimento.\n\n"
            "Diagnóstico institucional\n"
            "A Postal Saúde possui processo formal de atualização e gestão da rede, entretanto, a dinâmica permanente de "
            "credenciamentos, descredenciamentos, alterações cadastrais, suspensões e reorganizações exige atualização "
            "contínua das informações.\n"
            "No período de janeiro de 2025 a julho de 2026, foram registrados 305 novos credenciamentos e 580 "
            "descredenciamentos, evidenciando a intensidade da movimentação da rede e, consequentemente, a necessidade de "
            "mecanismos permanentes de higienização e atualização cadastral.\n"
            "O diagnóstico institucional indica, portanto, que o ponto de atenção está relacionado à tempestividade e "
            "integração das atualizações, e não à inexistência de estrutura ou processo formal de gestão da rede."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• Monitoramento sistemático das movimentações da rede.\n"
            "• Utilização do Painel de Vulnerabilidade da Rede Prestadora.\n"
            "• Aplicação da IOT 068, relativa à higienização da rede credenciada.\n"
            "• Aplicação da IOT 031, relativa ao descredenciamento de prestadores.\n"
            "• Aplicação da IOT 075, relativa aos casos de suspensão de atendimento.\n"
            "• Aplicação da IOT 158, relativa ao monitoramento da proporção de prestadores com variação significativa no atendimento.\n"
            "• Aplicação da IOT 224, relativa às rotinas dos relatórios de vulnerabilidade da rede.\n"
            "• Aprimoramento dos controles sistêmicos e automatizados voltados à rastreabilidade das movimentações."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "As movimentações da rede são acompanhadas pela GERED e submetidas aos fluxos institucionais correspondentes.\n"
            "O Painel de Vulnerabilidade é utilizado para identificação de situações que possam representar risco de "
            "desassistência, subsidiando medidas de expansão, substituição ou reorganização da rede.\n"
            "As situações de descredenciamento e suspensão de atendimento são objeto de tratamento específico, inclusive "
            "com identificação de alternativas assistenciais quando necessário.\n"
            "O processo permanece em acompanhamento contínuo, com foco na redução de divergências entre a situação "
            "operacional da rede e as informações disponibilizadas aos beneficiários."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED/CODER: Painel de análise de vulnerabilidade em processo de revisão metodológica.\n"
                "COCAP: Executa diariamente análises focadas na qualificação contínua do cadastro e construção de painel de controle. A suspensão temporária (IOT 075) é aplicada estritamente nas hipóteses previstas na RN nº 585 da ANS (obras/reformas ou determinação sanitária/fiscal).\n"
                "CONEC: Normativos em revisão visando alinhamento com diretrizes de compliance."
            ),
            "Status 09/2026": (
                "GERED/CODER: Revisão metodológica do painel de vulnerabilidade.\n"
                "COCAP: Análises diárias de qualificação cadastral, desenvolvimento de painel de controle e observância estrita da IOT 075 e RN nº 585/ANS.\n"
                "CONEC: Revisão contínua dos normativos da área."
            ),
            "Status 10/2026": (
                "GERED/CODER: Formalizada em 08/10/2026 demanda junto à TI para desenvolvimento e institucionalização do painel de vulnerabilidade."
            ),
            "Complemento ANS 26/10": (
                "GERED/CODER — Subitem: Vulnerabilidade de Rede\n"
                "O painel de análise de vulnerabilidade encontra-se em processo de revisão, tendo sido formalizada, em 08/10/2026, demanda junto à área de Tecnologia da Informação para desenvolvimento e institucionalização da ferramenta.\n\n"
                "GERED/CONEC — Subitem: Atualização de normativos (MAN006 e IOT036)\n"
                "O Manual de Credenciamento da Assistência Médico-Hospitalar (MAN006) encontra-se em revisão na plataforma DocSpider. A IOT 036.07 teve última atualização em 14/08/2026, com nova revisão prevista para 01/12/2027.\n\n"
                "GERED/COCAP — Subitem: Bases de rede e controles de atualização\n"
                "A Postal Saúde vem desenvolvendo ações voltadas ao aprimoramento dos processos de extração, atualização, monitoramento e higienização das informações cadastrais da rede assistencial, em atendimento ao Ofício nº 108/2026 da ANS. Nesse contexto, foram disponibilizadas bases cadastrais para análise e identificação de inconsistências, além da formalização do chamado GLPI nº 31043, destinado à implementação de relatório de movimentações cadastrais com identificação das inclusões, alterações, exclusões, datas e responsáveis. Também são realizadas consultas ao sistema BENNER, por meio do relatório POSTAL 988, para identificação de prestadores sem atendimento nos últimos 12 meses, subsidiando os procedimentos de higienização da rede. Adicionalmente, foram apresentadas as IOTs 031, 068 e 075, que regulamentam os processos de descredenciamento, higienização e suspensão de atendimento."
            )
        }
    },
    {
        "codigo": "8.5.2",
        "titulo": "8.5.2. DIVULGAÇÃO DA REDE ASSISTENCIAL NO PORTAL CORPORATIVO DA OPERADORA, NÃO INTEGRALMENTE NOS TERMOS DA RN ANS Nº 486/2022",
        "areas_atribuidas": ["CODER", "COCAP"],
        "apontamento_ans": (
            "Apontamento da ANS\n"
            "A ANS apontou possível não disponibilização integral das informações da rede assistencial no portal "
            "corporativo da Operadora, nos termos estabelecidos pela regulamentação vigente.\n\n"
            "Diagnóstico institucional\n"
            "A Postal Saúde mantém base estruturada de gestão da rede assistencial e mecanismos formais para registro e "
            "acompanhamento de suas alterações.\n"
            "Considerando a elevada dinâmica de movimentação da rede, o principal ponto de atenção consiste em assegurar "
            "que as informações constantes das bases operacionais estejam refletidas, de forma tempestiva e consistente, "
            "nos mecanismos de disponibilização ao beneficiário.\n"
            "A Operadora vem, portanto, fortalecendo a integração entre os processos de gestão da rede, os sistemas "
            "corporativos e os canais de disponibilização das informações."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• Gestão centralizada das movimentações da rede pela GERED.\n"
            "• Monitoramento da composição e vulnerabilidade da rede.\n"
            "• Aplicação dos normativos internos de cadastro, credenciamento, descredenciamento e higienização.\n"
            "• Aprimoramento dos controles sistêmicos e automatizados.\n"
            "• Fortalecimento da rastreabilidade das alterações cadastrais e contratuais.\n"
            "• Revisão contínua dos fluxos operacionais relacionados à atualização das informações."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "A consistência das informações da rede é tratada como componente da governança assistencial.\n"
            "As movimentações são acompanhadas pelas áreas responsáveis e confrontadas com os instrumentos gerenciais "
            "disponíveis, permitindo a identificação de alterações que demandem atualização e tratamento.\n"
            "O processo permanece submetido a ciclo contínuo de monitoramento e aperfeiçoamento, com foco na integração "
            "das informações e na redução de eventuais divergências entre as bases internas e os canais de disponibilização ao beneficiário."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED/CODER: Painel de análise de vulnerabilidade em processo de revisão.\n"
                "GERED/COCAP: Em desenvolvimento painel de controle gerencial (dashboard) para sistematizar a gestão centralizada das movimentações da rede, aplicação dos normativos internos e fortalecimento da rastreabilidade das alterações cadastrais."
            ),
            "Status 09/2026": (
                "GERED/CODER e GERED/COCAP: Continuidade da revisão do painel de vulnerabilidade e desenvolvimento do painel de controle gerencial das movimentações e atualizações da rede."
            ),
            "Status 10/2026": (
                "GERED/CODER: Demanda formalizada em 08/10/2026 junto à TI para institucionalização do painel de vulnerabilidade."
            ),
            "Complemento ANS 26/10": (
                "GERED/CODER — Subitem: Vulnerabilidade de Rede\n"
                "Painel de análise de vulnerabilidade em revisão metodológica, com demanda formalizada em 08/10/2026 junto à TI.\n\n"
                "GERED/COCAP — Subitem: Rede credenciada no portal: RN 486/2022, art. 2º\n"
                "A Postal Saúde vem adotando medidas para aprimorar os mecanismos de controle, atualização e transparência das informações disponibilizadas no Guia de Saúde, em atendimento aos requisitos da RN ANS nº 486/2022. Nesse contexto, foram formalizados os chamados GLPI nº 30182 e nº 30999, contemplando o desenvolvimento de ferramentas para gerenciamento de inconsistências cadastrais, rastreabilidade das manifestações dos beneficiários e adequação das funcionalidades do portal. A primeira etapa do sistema de gerenciamento de reportes foi disponibilizada para homologação, enquanto o painel gerencial existente já permite o acompanhamento de demandas e prazos. Adicionalmente, foram realizadas análises documentais de conformidade e identificadas oportunidades de melhoria nas informações cadastrais dos prestadores e planos de saúde."
            )
        }
    },
    {
        "codigo": "8.5.3",
        "titulo": "8.5.3. NÃO DISPONIBILIZAÇÃO DE INFORMAÇÃO SOBRE A SUBSTITUIÇÃO DE PRESTADORES DE SERVIÇOS DE ATENÇÃO À SAÚDE NÃO HOSPITALARES, CONFORME DETERMINA A RN ANS Nº 567/2022",
        "areas_atribuidas": ["CODER", "COCAP", "CONEC"],
        "apontamento_ans": (
            "Apontamento da ANS\n"
            "A ANS apontou possível não disponibilização de informações relativas à substituição de prestadores de serviços "
            "de atenção à saúde não hospitalares, nos termos da regulamentação vigente.\n\n"
            "Diagnóstico institucional\n"
            "A Postal Saúde possui processo formal de gestão das movimentações da rede, contemplando credenciamento, "
            "descredenciamento, recontratualização, substituição, suspensão e reorganização.\n"
            "No período analisado, foram registrados 580 descredenciamentos, sendo 104 relacionados ao interesse da "
            "Operadora, incluindo situações de substituição ou redução da rede.\n"
            "A Operadora também mantém mecanismos de análise de vulnerabilidade para avaliar os impactos assistenciais "
            "decorrentes das movimentações da rede e subsidiar a adoção de medidas de recomposição, substituição ou expansão."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• Aplicação da IOT 031, relativa ao descredenciamento de prestadores.\n"
            "• Aplicação da IOT 036, relativa ao credenciamento de prestadores.\n"
            "• Aplicação da IOT 035, relativa à recontratualização de prestadores.\n"
            "• Aplicação da IOT 068, relativa à higienização da rede credenciada.\n"
            "• Monitoramento da vulnerabilidade assistencial decorrente das movimentações da rede.\n"
            "• Avaliação da necessidade de substituição ou expansão da rede.\n"
            "• Fortalecimento dos mecanismos de rastreabilidade das movimentações.\n"
            "• Revisão contínua dos processos relacionados à gestão e comunicação das alterações da rede."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "Os processos de descredenciamento e substituição são acompanhados pela GERED e avaliados quanto aos seus "
            "potenciais impactos sobre a continuidade da assistência.\n"
            "O Painel de Vulnerabilidade da Rede é utilizado como instrumento de apoio à identificação de riscos e à "
            "tomada de decisão quanto à reorganização da rede.\n"
            "As alterações são tratadas mediante os fluxos e normativos institucionais aplicáveis, com acompanhamento da "
            "necessidade de manutenção ou recomposição da oferta assistencial.\n"
            "O processo permanece submetido ao ciclo de melhoria contínua, com foco no fortalecimento da rastreabilidade "
            "das substituições e na adequada disponibilização das informações aos beneficiários."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED/CODER: Revisão do painel de análise de vulnerabilidade.\n"
                "COCAP: Desenvolvimento de painel de controle gerencial para sistematizar a aplicação das IOTs 031, 036 e 068, avaliação de substituição e fortalecimento da rastreabilidade.\n"
                "CONEC: Revisão contínua dos normativos da área."
            ),
            "Status 09/2026": (
                "GERED/CODER, COCAP e CONEC: Manutenção das frentes de revisão do painel de vulnerabilidade, evolução do dashboard de controle de movimentações da rede e revisão normativa."
            ),
            "Status 10/2026": (
                "GERED/CODER: Demanda formalizada em 08/10/2026 junto à TI para institucionalização do painel de vulnerabilidade."
            ),
            "Complemento ANS 26/10": (
                "GERED/CODER — Subitem: Vulnerabilidade de Rede\n"
                "Painel de análise de vulnerabilidade em reestruturação junto à TI (demanda de 08/10/2026).\n\n"
                "GERED/CONEC — Subitem: Atualização de normativos (MAN006 e IOT036)\n"
                "O Manual MAN006 encontra-se em revisão na plataforma DocSpider e a IOT 036.07 foi atualizada em 14/08/2026.\n\n"
                "GERED/COCAP — Subitem: Registros de movimentação e substituição da rede / Substituição de rede não hospitalar: RN 567/2022, art. 11\n"
                "A resposta da Postal Saúde ao Ofício nº 108/2026 da ANS consolida as iniciativas da operadora para adequar os processos de movimentação e substituição da rede assistencial às exigências regulatórias, em especial ao Art. 11 da RN nº 567/2022. No documento, a operadora comprova possuir normativos internos estruturados para a gestão da rede e apresenta evidências visuais atestando que a comunicação das substituições aos beneficiários já ocorre no portal dentro dos prazos legais. Visando aprimorar a transparência e a usabilidade do Guia de Saúde, foram formalizadas demandas tecnológicas (GLPI nº 30992, nº 31012 e nº 30182) para extração consolidada de substitutos no Benner, melhoria da busca isolada por prestadores substitutos e tratamento estruturado de inconsistências cadastrais, além de instrumento de análise de equivalência assistencial, CTI DIOPE/GERED 0020/2025 e checklist documental de conformidade."
            )
        }
    },
    {
        "codigo": "8.6.1",
        "titulo": "8.6.1. Insuficiência de rede prestadora que garanta o cumprimento dos prazos máximos de atendimento, conforme a RN ANS nº 566/2022",
        "areas_atribuidas": ["CODER", "CONEC"],
        "apontamento_ans": (
            "Apontamento da ANS\n"
            "A ANS apontou a existência de indícios de insuficiência de rede prestadora para garantir o cumprimento dos "
            "prazos máximos de atendimento estabelecidos pela RN ANS nº 566/2022, com reflexos na garantia de acesso dos "
            "beneficiários aos serviços e procedimentos de cobertura obrigatória. A análise apresentada pela Agência "
            "registra, ainda, relatos de dificuldade de acesso decorrentes de indisponibilidade ou insuficiência de "
            "prestadores, inclusive situações relacionadas à suspensão de atendimentos e à necessidade de utilização de "
            "alternativas assistenciais.\n\n"
            "Diagnóstico Institucional\n"
            "A Postal Saúde reconhece que, no período analisado pela ANS, foram identificadas situações pontuais de "
            "indisponibilidade, insuficiência ou inadequação da rede prestadora, especialmente em determinadas localidades, "
            "especialidades e serviços assistenciais, com potencial impacto sobre o cumprimento tempestivo dos prazos regulamentares.\n"
            "A análise interna também evidenciou que o risco relacionado à rede não deve ser tratado exclusivamente sob a "
            "perspectiva quantitativa de prestadores credenciados, mas considerando sua capacidade assistencial efetiva, "
            "disponibilidade para o serviço demandado, especialidade, município, abrangência territorial, modalidade de "
            "atendimento e efetiva capacidade de agendamento dentro do prazo regulamentar.\n"
            "Nesse sentido, a Postal Saúde passou a fortalecer o acompanhamento da vulnerabilidade da rede e a integração "
            "entre os mecanismos de monitoramento da Garantia de Atendimento e a gestão da rede credenciada."
        ),
        "medidas_implementadas": (
            "Medidas Implementadas e em Implementação\n\n"
            "A Postal Saúde implementou e vem aprimorando conjunto integrado de medidas destinado à prevenção, "
            "identificação tempestiva e tratamento das situações de insuficiência de rede, destacando-se:\n"
            "• Monitoramento diário da Garantia de Atendimento, mediante painel gerencial em BI, com acompanhamento das "
            "demandas recepcionadas, situação dos atendimentos e prazos regulatórios.\n"
            "• Monitoramento da vulnerabilidade da rede credenciada, permitindo identificar localidades e especialidades "
            "com maior exposição a risco assistencial.\n"
            "• Acompanhamento direcionado das unidades regionais com maior incidência de demandas de Garantia de Atendimento.\n"
            "• Emissão e encaminhamento de relatórios analíticos às Regionais, subsidiando a atuação dos gestores locais e "
            "a adoção de medidas corretivas.\n"
            "• Acompanhamento individualizado das demandas críticas e reincidentes, com priorização das situações com "
            "maior risco de extrapolação de prazo.\n"
            "• Utilização de Contratação Direta, Reembolso Integral e Tratamento Fora do Domicílio, quando configuradas as "
            "hipóteses aplicáveis e necessárias à garantia do atendimento.\n"
            "• Padronização dos fluxos operacionais por meio do Manual da Garantia de Atendimento e dos procedimentos "
            "operacionais associados.\n"
            "• Automação e melhoria dos formulários e fluxos eletrônicos, inclusive mediante migração para o E-doc, com "
            "objetivo de ampliar rastreabilidade, padronização e controle."
        ),
        "observacoes": (
            "Tratamento e Monitoramento\n\n"
            "O tratamento do risco de insuficiência de rede passou a ser conduzido sob perspectiva preventiva e contínua, "
            "mediante integração entre monitoramento da demanda, análise da capacidade assistencial da rede e atuação regional.\n"
            "O acompanhamento diário permite identificar demandas próximas do vencimento dos prazos regulamentares e "
            "direcionar tempestivamente a atuação das áreas responsáveis.\n"
            "Paralelamente, a análise de vulnerabilidade da rede possibilita a identificação de cenários que possam "
            "comprometer a capacidade de atendimento, permitindo adoção antecipada de medidas de mitigação.\n"
            "A GERED/CODER mantém, ainda, rotina de alinhamento com as Regionais, orientação operacional e "
            "disponibilização de ferramentas de BI para consulta da rede, beneficiários, eventos e demandas de Garantia de Atendimento.\n"
            "O monitoramento também contempla a classificação adequada dos protocolos, especialmente após a alteração "
            "processual implementada em outubro de 2025, de modo a evitar que demandas efetivamente caracterizadas como "
            "Garantia de Atendimento permaneçam indevidamente classificadas como Rede Credenciada."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED: As informações deste item estão contempladas no item 8.1.2 no que se refere ao monitoramento diário das UAR/NAR com maior incidência em processos de Garantia de Atendimento fora do prazo."
            ),
            "Status 09/2026": (
                "GERED: Monitoramento diário das UAR/NAR conduzido de forma integrada ao item 8.1.2, registrando queda de -15,91% no volume de processos de Garantia de Atendimento em setembro/2026."
            ),
            "Status 10/2026": (
                "GERED/CODER: Implementada em 01/10/2026 a obrigatoriedade de o beneficiário informar o número do protocolo de atendimento ao solicitar Reembolso Integral no aplicativo, fortalecendo a rastreabilidade dos processos de Garantia de Atendimento."
            ),
            "Complemento ANS 26/10": ""
        }
    }
]

CHECKLISTS_INICIAIS_COMPLEMENTO = {
    "8.4.2": {
        "CODER": [
            {
                "titulo": "Demandas de Garantia de Atendimento geradas por essas suspensões e sua solução",
                "analise": "Foram adotadas medidas imediatas pela GERED/CODER para assegurar a continuidade da assistência aos beneficiários impactados por suspensões contratuais, mediante direcionamento assistencial para prestadores aptos da rede direta/indireta, contratação direta ou liberação de reembolso integral nos prazos da RN nº 566/2022."
            },
            {
                "titulo": "Controle mensal de prestadores suspensos e restabelecidos (indicadores do PRASS)",
                "analise": "Monitoramento mensal consolidado pela GERED acerca da evolução das suspensões e restabelecimentos de atendimento na rede assistencial."
            }
        ],
        "CONEC": [
            {
                "titulo": "Relação dos prestadores que suspenderam ou ameaçaram suspender por inadimplência: prestador, UF, data, valor em aberto e situação",
                "analise": "Levantamento estruturado contendo a identificação dos prestadores com notificação ou registro de suspensão por inadimplência, unidade federativa, data da ocorrência e status atual das tratativas."
            },
            {
                "titulo": "Registros de negociação: e-mails, ofícios, notificações recebidas e atas de reunião com prestadores",
                "analise": "Documentação comprobatória das tratativas negociais conduzidas junto aos prestadores e unidades regionais para regularização dos atendimentos."
            },
            {
                "titulo": "Termos de acordo, parcelamento ou confissão de dívida e comprovantes de pagamento",
                "analise": "Instrumentos formais de negociação e comprovação de quitação/parcelamento articulados em conjunto com a área financeira (GECOF)."
            },
            {
                "titulo": "Comunicação de restabelecimento do atendimento, com data",
                "analise": "Registros formais de comunicação do restabelecimento do fluxo assistencial junto aos prestadores regularizados."
            }
        ]
    },
    "8.5.1": {
        "COCAP": [
            {
                "titulo": "Extração da base de rede com data de corte, por tipo, UF e especialidade, conciliada com o portal",
                "analise": "Disponibilizadas extrações cadastrais completas da rede de prestadores de serviços de saúde extraídas do sistema corporativo BENNER (.CSV com 12.360 e 12.403 registros), contemplando CPF/CNPJ, razão social, nome fantasia, categoria, tipo de prestador, datas de credenciamento/inclusão, endereço com IBGE/UF/CEP e especialidades."
            },
            {
                "titulo": "Relatório de alterações cadastrais do mês: inclusões, alterações e exclusões, com data e responsável",
                "analise": "Formalizada solicitação de extração junto à área técnica responsável por meio do chamado GLPI nº 31043 (COIDE > Dados Gerenciais), abrangendo os registros de inclusão, alteração e exclusão de prestadores realizados em agosto e setembro de 2026, com data, usuário responsável e valores anteriores/posteriores."
            },
            {
                "titulo": "Registros da rotina da GERED: checklist e amostra de validações",
                "analise": "Formalizado documento técnico da GERED/COCAP registrando o mapeamento e incorporação de checklist padronizado e metodologia de amostragem de validações às rotinas diárias de qualificação cadastral."
            },
            {
                "titulo": "Relatório de higienização da rede (IOT 068) do período",
                "analise": "Apresentado o procedimento de extração do relatório POSTAL 988 (Comprovante de Atendimento do Prestador no Período de 12 Meses) no sistema BENNER, exemplificado com amostra referente ao prestador Instituição Adventista Centro Oeste de Promoção à Saúde (CNPJ 24.823.958/0002-21), período de 27/10/2025 a 27/10/2026, em cumprimento à IOT 068."
            },
            {
                "titulo": "IOT 031, 068, 075, 158 e 224 revisadas ou cronograma formal de revisão",
                "analise": "Apresentadas na íntegra as Instruções Operacionais de Trabalho vigentes: IOT 075.02 (Suspensão de Atendimento - 8 págs.), IOT 031.07 (Descredenciamento de Prestadores - 12 págs.) e IOT 068.04 (Higienização da Rede Credenciada - 9 págs.), acompanhadas da avaliação de revisão normativa junto às IOTs 158 e 224."
            }
        ],
        "CODER": [
            {
                "titulo": "Revisão metodológica e reestruturação do Painel de Vulnerabilidade da Rede Prestadora",
                "analise": "Formalizada em 08/10/2026 demanda junto à área de Tecnologia da Informação para desenvolvimento e institucionalização do painel de vulnerabilidade, fortalecendo a rastreabilidade das bases utilizadas."
            }
        ],
        "CONEC": [
            {
                "titulo": "Atualização do Manual de Credenciamento (MAN006) e Instrução Operacional IOT 036.07",
                "analise": "Demonstrado o andamento da revisão do MAN006 na plataforma DocSpider e a vigência da IOT 036.07 atualizada em 14/08/2026."
            }
        ]
    },
    "8.5.2": {
        "COCAP": [
            {
                "titulo": "Relatório das notificações 'Encontrou inconsistência?' do portal e prazo de correção",
                "analise": "Formalizada demanda tecnológica sob o chamado GLPI nº 30182 (COSIS > Projetos). A 1ª Entrega (Registro e Atendimento de Reportes do Guia Saúde) foi disponibilizada em ambiente de homologação em 21/10/2026, contemplando geração de protocolo, envio de e-mail automático com prazo de 5 dias úteis, painel administrativo com acesso AD e filtros por Status/Ano/Mês/Tipo/UF, além do Manual Operacional."
            },
            {
                "titulo": "Cronograma do novo painel de controle e prints da versão atual",
                "analise": "Apresentado registro visual da versão atual do painel gerencial utilizado pela GERED/COCAP para acompanhamento de atividades operacionais, demandas, prazos e indicadores da rede assistencial."
            },
            {
                "titulo": "Bloco C1 (RN 486/2022, art. 2º) - Cronograma de Implantação",
                "analise": "Formalizada solicitação de adequação do Guia de Saúde aos requisitos do art. 2º da RN ANS nº 486/2022 por meio do chamado GLPI nº 30999 (COSIS > Minha Postal), contemplando a exibição individualizada de todos os planos atendidos por prestador e revisão dos mecanismos de busca."
            },
            {
                "titulo": "Bloco C1 (RN 486/2022, art. 2º) - Demanda de TI aberta (número e data) e especificação aprovada",
                "analise": "Comprovada a abertura do chamado GLPI nº 30999 em 26/10/2026 com a especificação detalhada de todos os requisitos do art. 2º, incisos I e II e alínea 'g', da RN ANS nº 486/2022."
            },
            {
                "titulo": "Bloco C1 (RN 486/2022, art. 2º) - Homologação e data de entrada em produção",
                "analise": "Apresentada a especificação técnica registrada no chamado GLPI nº 30999 para orientar o desenvolvimento, homologação e entrada em produção das melhorias no portal."
            },
            {
                "titulo": "Bloco C1 (RN 486/2022, art. 2º) - Prints datados com URL, antes e depois, em ao menos três tipos de prestador",
                "analise": "Apresentados registros visuais datados (27/10/2026, às 10h28) com URL visível do Guia de Saúde (guiasaude.postalsaude.com.br), demonstrando a exibição cadastral de diferentes estabelecimentos (clínicas médicas e centros de diagnóstico) e o mecanismo de reporte de inconsistências."
            },
            {
                "titulo": "Bloco C1 (RN 486/2022, art. 2º) - Checklist de conformidade do art. 2º assinado pela área",
                "analise": "Elaborado Checklist de Auditoria Documental do Art. 2º da RN 486/2022 (5 páginas), avaliando item a item (I-a a I-d, II-a a II-g e § 1º) com telas comprobatórias e vinculação ao chamado GLPI nº 30999."
            }
        ]
    },
    "8.5.3": {
        "COCAP": [
            {
                "titulo": "Bloco B6 - Relação do período: prestador, CNPJ, data, motivo, substituto e data de publicação no portal",
                "analise": "Formalizada solicitação de extração consolidada no sistema BENNER por meio do chamado GLPI nº 30992 (aberto em 26/10/2026 na fila COSIS > Benner - Saúde > Relatórios), com envio da planilha base de descredenciados de agosto e setembro/2026 para retorno estruturado de todos os campos de substituição e datas de publicação."
            },
            {
                "titulo": "Bloco B6 - Amostra de processos completos (IOT 031 e 036): solicitação, parecer, aprovação e aditivo ou distrato",
                "analise": "Apresentado compilado normativo de 59 páginas contendo a íntegra da IOT 031.07 (Descredenciamento de Prestadores - págs. 1 a 12) e da IOT 036.07 (Credenciamento de Prestador - págs. 13 a 59), incluindo os formulários FOP 056 e FOP 075."
            },
            {
                "titulo": "Bloco B6 - Prova da comunicação ao beneficiário: print datado e data de publicação",
                "analise": "Apresentada evidência comparativa contendo: (1) Captura da consulta pública de substituição no Guia de Saúde (Clínica Bella Vita / Corpus Clean) e (2) Captura da tela cadastral no sistema BENNER comprovando o registro do período de publicação no portal (Início: 01/06/2026 e Fim: 28/11/2026)."
            },
            {
                "titulo": "Bloco B6 - Indicador do PRASS: percentual de substituições comunicadas no prazo",
                "analise": "Vinculado ao levantamento primário solicitado no chamado GLPI nº 30992 junto à equipe do sistema BENNER para apuração auditável do numerador e denominador das substituições comunicadas no prazo regulatório."
            },
            {
                "titulo": "Bloco B6 - Análise de equivalência do substituto (mesmo serviço e município)",
                "analise": "Apresentado registro visual da ferramenta técnica de verificação de equivalência assistencial entre prestador excluído e substituto (conferência de município, porte e serviços contratados)."
            },
            {
                "titulo": "Bloco B6 - Termos aditivos de habilitação e CTI DIOPE/GERED 0020/2025, prometidos no PRASS",
                "analise": "Apresentada a íntegra da Comunicação Interna CTI DIOPE/GERED nº 0020/2025 (NUP 00000.020940/2025-81), de 03/07/2025, assinada digitalmente por João Victor Peres Moura (código verificador 73F10830), determinando às regionais o mapeamento de hospitais substitutos em observância à RN 566/2022 e ao MAN 015."
            },
            {
                "titulo": "Bloco C2 (RN 567/2022, art. 11) - Cronograma de Implantação",
                "analise": "Formalizada demanda evolutiva junto à TI por meio do chamado GLPI nº 31012 (aberto em 27/10/2026 na fila COSIS > Intranet / Site), solicitando a implementação da pesquisa isolada por prestador na página de substituições (art. 11, inciso II, da RN 567/2022) e respectivo cronograma."
            },
            {
                "titulo": "Bloco C2 (RN 567/2022, art. 11) - Demanda de TI, especificação e homologação",
                "analise": "Comprovado o registro do chamado GLPI nº 31012 contendo a especificação funcional detalhada para adequação da busca isolada e combinada de substituições no portal conforme o art. 11, inciso II, da RN 567/2022 c/c a RN 486/2022."
            },
            {
                "titulo": "Bloco C2 (RN 567/2022, art. 11) - Prints datados com URL, antes e depois, de ao menos três substituições",
                "analise": "Apresentado caderno comprobatório da CODER/GERED (5 páginas, datado de 27/10/2026) com capturas das telas de substituição para 3 prestadores distintos: 1) CR Serviços de Saúde Ltda (SC); 2) Hospital da Criança Urgil Urgência Infantil (AC); e 3) Ortoclínica de Sergipe Ltda (SE), abrangendo os planos CorreioSaúde, CorreioSaúde II e Viver Saúde."
            },
            {
                "titulo": "Bloco C2 (RN 567/2022, art. 11) - Saneamento dos cadastros com erro (datas e telefones) e relatório da correção",
                "analise": "Apresentado o fluxo estruturado de tratamento e saneamento de inconsistências cadastrais do Guia de Saúde disponibilizado em homologação por meio do chamado GLPI nº 30182, aliado à conferência cadastral direta no sistema BENNER."
            },
            {
                "titulo": "Bloco C2 (RN 567/2022, art. 11) - Checklist de conformidade do art. 11 assinado",
                "analise": "Elaborado documento de verificação de conformidade do art. 11 da RN 567/2022 (2 páginas), demonstrando por capturas de tela o atendimento às alíneas 'a', 'b', 'c' e 'd' do inciso I e o acompanhamento da busca isolada do inciso II via GLPI nº 31012."
            }
        ]
    }
}


# ==============================================================================
# CAMADA DE BANCO DE DADOS COM AUTO-FALLBACK E MIGRAÇÃO TRANSPARENTE
# ==============================================================================
def obter_url_postgres():
    """Busca a URL do PostgreSQL priorizando DATABASE_PUBLIC_URL para evitar falha IPv6 no railway.internal."""
    for nome_var in ("DATABASE_PUBLIC_URL", "DATABASE_URL", "POSTGRES_PUBLIC_URL", "POSTGRES_URL", "POSTGRESQL_URL"):
        val = os.environ.get(nome_var, "").strip()
        if val and not val.startswith("${{") and (val.startswith("postgres://") or val.startswith("postgresql://")):
            return val.replace("postgres://", "postgresql://", 1)

    pghost = os.environ.get("PGHOST", "").strip()
    pguser = os.environ.get("PGUSER", "").strip()
    pgpass = os.environ.get("PGPASSWORD", "").strip()
    pgdb = os.environ.get("PGDATABASE", "").strip()
    pgport = os.environ.get("PGPORT", "5432").strip()
    if pghost and pguser and pgdb and not pghost.startswith("${{"):
        return f"postgresql://{pguser}:{pgpass}@{pghost}:{pgport}/{pgdb}"

    return ""


DATABASE_URL = obter_url_postgres()
USE_POSTGRES = bool(DATABASE_URL and HAS_PSYCOPG2)
DB_INITIALIZED = False
LAST_PG_ERROR = ""
SQLITE_PATH = os.path.join(os.path.abspath(os.path.dirname(__file__)), "prass_local.db")


def get_db_conn():
    global USE_POSTGRES, LAST_PG_ERROR, DB_INITIALIZED
    if USE_POSTGRES:
        for _ in range(3):
            try:
                conn = psycopg2.connect(DATABASE_URL, connect_timeout=5)
                return conn, True
            except Exception as e:
                LAST_PG_ERROR = str(e).strip()
                time.sleep(0.5)
        print(f"[FALLBACK DB] PostgreSQL indisponivel ({LAST_PG_ERROR}). Usando SQLite persistente.")
        USE_POSTGRES = False
        DB_INITIALIZED = False

    conn = sqlite3.connect(SQLITE_PATH, timeout=20)
    conn.row_factory = sqlite3.Row
    return conn, False


def db_query(sql, params=(), fetch=True, commit=False):
    conn, is_pg = get_db_conn()
    try:
        if is_pg:
            cur = conn.cursor(cursor_factory=RealDictCursor)
            cur.execute(sql, params)
        else:
            sql_sqlite = sql.replace("%s", "?")
            cur = conn.cursor()
            cur.execute(sql_sqlite, params)

        res = None
        if fetch:
            rows = cur.fetchall()
            res = [dict(r) for r in rows]
        if commit:
            conn.commit()
        return res
    finally:
        conn.close()


def init_db():
    global DB_INITIALIZED
    conn, is_pg = get_db_conn()
    try:
        cur = conn.cursor()
        if is_pg:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS prass_items (
                    codigo VARCHAR(20) PRIMARY KEY,
                    titulo TEXT NOT NULL,
                    apontamento_ans TEXT NOT NULL,
                    medidas_implementadas TEXT NOT NULL,
                    observacoes TEXT NOT NULL,
                    historico_json TEXT NOT NULL,
                    areas_atribuidas TEXT NOT NULL,
                    oficio_rascunho TEXT DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS demandas (
                    id SERIAL PRIMARY KEY,
                    item_codigo VARCHAR(20) NOT NULL,
                    nome VARCHAR(180) NOT NULL,
                    tipo VARCHAR(20) NOT NULL,
                    ordem INT DEFAULT 100,
                    resumo_texto TEXT DEFAULT '',
                    oficio_texto TEXT DEFAULT '',
                    UNIQUE(item_codigo, nome)
                );
                CREATE TABLE IF NOT EXISTS checklist_evidencias (
                    id SERIAL PRIMARY KEY,
                    demanda_id INT NOT NULL REFERENCES demandas(id) ON DELETE CASCADE,
                    setor VARCHAR(20) NOT NULL,
                    titulo TEXT NOT NULL,
                    analise_evidencia TEXT DEFAULT '',
                    concluido BOOLEAN DEFAULT FALSE,
                    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS anexos (
                    id SERIAL PRIMARY KEY,
                    checklist_id INT NOT NULL REFERENCES checklist_evidencias(id) ON DELETE CASCADE,
                    nome_arquivo VARCHAR(255) NOT NULL,
                    mime_type VARCHAR(120) NOT NULL,
                    tamanho INT NOT NULL,
                    resumo_extraido TEXT DEFAULT '',
                    conteudo BYTEA NOT NULL,
                    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
            """)
            cur.execute("ALTER TABLE demandas ADD COLUMN IF NOT EXISTS resumo_texto TEXT DEFAULT '';")
        else:
            cur.executescript("""
                CREATE TABLE IF NOT EXISTS prass_items (
                    codigo TEXT PRIMARY KEY,
                    titulo TEXT NOT NULL,
                    apontamento_ans TEXT NOT NULL,
                    medidas_implementadas TEXT NOT NULL,
                    observacoes TEXT NOT NULL,
                    historico_json TEXT NOT NULL,
                    areas_atribuidas TEXT NOT NULL,
                    oficio_rascunho TEXT DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS demandas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    item_codigo TEXT NOT NULL,
                    nome TEXT NOT NULL,
                    tipo TEXT NOT NULL,
                    ordem INTEGER DEFAULT 100,
                    resumo_texto TEXT DEFAULT '',
                    oficio_texto TEXT DEFAULT '',
                    UNIQUE(item_codigo, nome)
                );
                CREATE TABLE IF NOT EXISTS checklist_evidencias (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    demanda_id INTEGER NOT NULL,
                    setor TEXT NOT NULL,
                    titulo TEXT NOT NULL,
                    analise_evidencia TEXT DEFAULT '',
                    concluido INTEGER DEFAULT 0,
                    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS anexos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    checklist_id INTEGER NOT NULL,
                    nome_arquivo TEXT NOT NULL,
                    mime_type TEXT NOT NULL,
                    tamanho INTEGER NOT NULL,
                    resumo_extraido TEXT DEFAULT '',
                    conteudo BLOB NOT NULL,
                    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
            """)
            cur.execute("PRAGMA table_info(demandas);")
            cols = [row[1] for row in cur.fetchall()]
            if "resumo_texto" not in cols:
                cur.execute("ALTER TABLE demandas ADD COLUMN resumo_texto TEXT DEFAULT '';")
        conn.commit()
    finally:
        conn.close()

    existentes = db_query("SELECT codigo FROM prass_items")
    codigos_existentes = {r["codigo"] for r in existentes}

    for item in ITENS_PRASS_INICIAIS:
        cod = item["codigo"]
        hist_dict = item.get("historico_planilha", {})
        if cod not in codigos_existentes:
            db_query(
                """INSERT INTO prass_items
                   (codigo, titulo, apontamento_ans, medidas_implementadas, observacoes, historico_json, areas_atribuidas)
                   VALUES (%s, %s, %s, %s, %s, %s, %s)""",
                (
                    cod,
                    item["titulo"],
                    item["apontamento_ans"],
                    item["medidas_implementadas"],
                    item["observacoes"],
                    json.dumps(hist_dict, ensure_ascii=False),
                    json.dumps(item.get("areas_atribuidas", ["CODER", "COCAP", "CONEC"]), ensure_ascii=False),
                ),
                fetch=False,
                commit=True,
            )
        else:
            # Atualiza o historico_json com os textos integrais da planilha
            db_query(
                """UPDATE prass_items
                   SET apontamento_ans = %s, medidas_implementadas = %s, observacoes = %s, historico_json = %s
                   WHERE codigo = %s""",
                (
                    item["apontamento_ans"],
                    item["medidas_implementadas"],
                    item["observacoes"],
                    json.dumps(hist_dict, ensure_ascii=False),
                    cod
                ),
                fetch=False,
                commit=True
            )

        for idx_d, nome_dem in enumerate(DEMANDAS_FIXAS_PADRAO, start=1):
            ja_tem = db_query(
                "SELECT id, resumo_texto FROM demandas WHERE item_codigo = %s AND nome = %s",
                (cod, nome_dem)
            )
            resumo_inicial = hist_dict.get(nome_dem, "").strip()
            if resumo_inicial in ("N/A", "Preencher", "Prencher"):
                resumo_inicial = ""

            if not ja_tem:
                db_query(
                    "INSERT INTO demandas (item_codigo, nome, tipo, ordem, resumo_texto) VALUES (%s, %s, %s, %s, %s)",
                    (cod, nome_dem, "FIXA", idx_d, resumo_inicial),
                    fetch=False,
                    commit=True
                )
            elif not (ja_tem[0].get("resumo_texto") or "").strip() and resumo_inicial:
                db_query(
                    "UPDATE demandas SET resumo_texto = %s WHERE id = %s",
                    (resumo_inicial, ja_tem[0]["id"]),
                    fetch=False,
                    commit=True
                )

        if cod in CHECKLISTS_INICIAIS_COMPLEMENTO:
            dem_row = db_query(
                "SELECT id FROM demandas WHERE item_codigo = %s AND nome = %s",
                (cod, "Complemento ANS 26/10")
            )
            if dem_row:
                dem_id = dem_row[0]["id"]
                exist_chk = db_query(
                    "SELECT COUNT(*) AS qtd FROM checklist_evidencias WHERE demanda_id = %s",
                    (dem_id,)
                )
                if exist_chk and int(exist_chk[0]["qtd"]) == 0:
                    for setor, lista_objs in CHECKLISTS_INICIAIS_COMPLEMENTO[cod].items():
                        for obj_ev in lista_objs:
                            tit = obj_ev["titulo"] if isinstance(obj_ev, dict) else str(obj_ev)
                            ana = obj_ev.get("analise", "") if isinstance(obj_ev, dict) else ""
                            db_query(
                                """INSERT INTO checklist_evidencias (demanda_id, setor, titulo, analise_evidencia, concluido)
                                   VALUES (%s, %s, %s, %s, %s)""",
                                (dem_id, setor, tit, ana, False),
                                fetch=False,
                                commit=True
                            )

    DB_INITIALIZED = True


@app.before_request
def garantir_banco_inicializado():
    global DB_INITIALIZED
    if not DB_INITIALIZED:
        try:
            init_db()
        except Exception as e:
            print(f"[ERRO INIT_DB]: {e}")
            traceback.print_exc()


# ==============================================================================
# FUNÇÕES AUXILIARES DE ANÁLISE DE ANEXOS E GERAÇÃO DE PDF / WORD
# ==============================================================================
def extrair_resumo_arquivo(nome_arquivo, raw_bytes):
    ext = os.path.splitext(nome_arquivo.lower())[1]
    try:
        if ext == ".pdf" and HAS_PYPDF:
            reader = pypdf.PdfReader(io.BytesIO(raw_bytes))
            num_pags = len(reader.pages)
            trechos = []
            for p in reader.pages[:3]:
                txt = (p.extract_text() or "").strip()
                if txt:
                    trechos.append(" ".join(txt.split())[:350])
            resumo_txt = " | ".join(trechos)[:600] if trechos else "Documento PDF composto por imagens/capturas de tela."
            return f"Documento PDF com {num_pags} página(s). Síntese extraída: {resumo_txt}"
        elif ext in (".xlsx", ".xls", ".csv"):
            if ext == ".csv":
                linhas = raw_bytes.decode("latin-1", errors="ignore").splitlines()
                cabecalho = linhas[0][:250] if linhas else ""
                return (
                    f"Planilha/Base de Dados em formato {ext.upper()} (referenciada no Ofício pelo nome '{nome_arquivo}'), "
                    f"contendo {max(0, len(linhas)-1)} registros. Colunas identificadas: {cabecalho}"
                )
            return (
                f"Planilha eletrônica Excel ({ext.upper()}) mantida como arquivo externo referenciado sob o nome "
                f"'{nome_arquivo}' ({len(raw_bytes)/1024:.1f} KB)."
            )
        elif ext in (".png", ".jpg", ".jpeg") and HAS_PIL:
            im = Image.open(io.BytesIO(raw_bytes))
            return f"Captura de tela / Evidência visual ({ext.upper()}, resolução {im.size[0]}x{im.size[1]}px)."
        else:
            return f"Documento comprobatório ({ext.upper()}, {len(raw_bytes)/1024:.1f} KB)."
    except Exception:
        return f"Arquivo comprobatório anexado ({nome_arquivo})."


def calcular_semaforo_setores(demanda_id, areas_atribuidas):
    rows = db_query(
        """SELECT c.id, c.setor, c.concluido, COUNT(a.id) AS qtd_anexos
           FROM checklist_evidencias c
           LEFT JOIN anexos a ON a.checklist_id = c.id
           WHERE c.demanda_id = %s
           GROUP BY c.id, c.setor, c.concluido""",
        (demanda_id,)
    )
    por_setor = {"CODER": [], "COCAP": [], "CONEC": []}
    for r in rows:
        s = r["setor"]
        if s in por_setor:
            por_setor[s].append(r)

    resultado = {}
    for setor in ["CODER", "COCAP", "CONEC"]:
        itens = por_setor[setor]
        atribuido = setor in areas_atribuidas
        total = len(itens)
        concluidos = sum(1 for x in itens if bool(x["concluido"]))
        com_anexo = sum(1 for x in itens if int(x["qtd_anexos"]) > 0)

        if total == 0:
            cor = "VERMELHO"
            label = "Não Respondido"
        elif concluidos == total and total > 0:
            cor = "VERDE"
            label = f"Respondido ({concluidos}/{total})"
        elif concluidos > 0 or com_anexo > 0:
            cor = "AMARELO"
            label = f"Respondendo ({concluidos}/{total})"
        else:
            cor = "VERMELHO"
            label = f"Não Respondido (0/{total})"

        resultado[setor] = {
            "atribuido": atribuido,
            "cor": cor,
            "label": label,
            "total": total,
            "concluidos": concluidos,
            "percentual": round((concluidos / total * 100), 1) if total > 0 else 0.0
        }
    return resultado


def construir_texto_oficio_automatico(item_row, demanda_row, checklists_com_anexos):
    """
    Monta o Ofício colocando PRIMEIRO o texto da Caixinha de Resumo (Seção 1)
    e DEPOIS a Análise Técnica de todas as evidências e anexos (Seção 2).
    """
    data_str = datetime.now().strftime("%d/%m/%Y")
    resumo_caixinha = (demanda_row.get("resumo_texto") or "").strip()
    hist = json.loads(item_row.get("historico_json") or "{}")
    resumo_planilha = hist.get(demanda_row["nome"], "").strip()

    texto_resumo_inicial = resumo_caixinha if resumo_caixinha else (
        resumo_planilha if resumo_planilha and resumo_planilha not in ("N/A", "Preencher") else (
            "Em atendimento às determinações da Agência Nacional de Saúde Suplementar (ANS) no âmbito do Plano de "
            "Recuperação Assistencial (PRASS - Processo ANS nº 33910.030426/2026-21), a Gerência Executiva de Rede (GERED) "
            "apresenta a seguir o resumo executivo e a análise técnica das evidências documentais referentes ao "
            f"item {item_row['codigo']}, consolidadas pelas coordenações técnicas vinculadas."
        )
    )

    linhas = [
        f"Brasília, {data_str}",
        "",
        "RESPOSTA AO OFÍCIO DA ANS — MONITORAMENTO PRASS POSTAL SAÚDE",
        f"ITEM DO PRASS: {item_row['titulo']}",
        f"REFERÊNCIA / DEMANDA: {demanda_row['nome']} ({demanda_row['tipo']})",
        f"COORDENAÇÕES ATRIBUÍDAS (GERED): {', '.join(json.loads(item_row['areas_atribuidas']))}",
        "",
        "1. RESUMO EXECUTIVO E CONSIDERAÇÕES INICIAIS",
        texto_resumo_inicial,
        "",
        "2. ANÁLISE TÉCNICA DAS EVIDÊNCIAS E DOCUMENTOS COMPROBATÓRIOS",
        ""
    ]

    if not checklists_com_anexos:
        linhas.append("Nenhum item de checklist cadastrado para esta demanda até o momento.")
    else:
        for idx, chk in enumerate(checklists_com_anexos, start=1):
            status_txt = "CONCLUÍDO" if chk["concluido"] else "EM ANDAMENTO / PENDENTE DE COMPLEMENTAÇÃO"
            linhas.append(f"2.{idx}. [{chk['setor']}] {chk['titulo']}")

            anexos = chk.get("anexos", [])
            if anexos:
                nomes_arqs = []
                notas_excel = []
                resumos_arqs = []
                for a in anexos:
                    n_arq = a["nome_arquivo"]
                    nomes_arqs.append(n_arq)
                    ext = os.path.splitext(n_arq.lower())[1]
                    if ext in (".xlsx", ".xls", ".csv"):
                        notas_excel.append(
                            f"• Planilha/Base Externa Referenciada: '{n_arq}' (mantida em formato original de planilha "
                            f"eletrônica e não mesclada ao corpo do PDF para preservação das linhas e colunas auditáveis)."
                        )
                    else:
                        resumos_arqs.append(f"• {n_arq}: {a.get('resumo_extraido', 'Anexado integralmente ao dossiê PDF.')}")

                linhas.append(f"Evidência(s) apresentada(s): {', '.join(nomes_arqs)}")
                linhas.append("")
                if chk.get("analise_evidencia") and chk["analise_evidencia"].strip():
                    linhas.append(chk["analise_evidencia"].strip())
                else:
                    linhas.append(
                        f"A coordenação {chk['setor']}/GERED disponibiliza a documentação comprobatória referente ao item "
                        f"'{chk['titulo']}', extraída dos sistemas corporativos da Postal Saúde (BENNER / GLPI / Portal Guia de Saúde) "
                        f"e validada conforme os procedimentos operacionais vigentes."
                    )
                if resumos_arqs:
                    linhas.append("Detalhamento dos documentos incorporados ao PDF consolidado:")
                    linhas.extend(resumos_arqs)
                if notas_excel:
                    linhas.append("Referência de Bases em Excel / CSV:")
                    linhas.extend(notas_excel)
            else:
                linhas.append("Evidência(s) apresentada(s): (Sem anexo vinculado até o momento)")
                if chk.get("analise_evidencia") and chk["analise_evidencia"].strip():
                    linhas.append(chk["analise_evidencia"].strip())
                else:
                    linhas.append(
                        f"Item registrado sob responsabilidade da coordenação {chk['setor']}, aguardando juntada do "
                        f"documento primário comprobatório."
                    )

            linhas.append(f"Situação: {status_txt}.")
            linhas.append("-" * 78)
            linhas.append("")

    linhas.extend([
        "",
        "Atenciosamente,",
        "",
        "João Victor Peres de Moura",
        "Gerente Executivo de Rede",
        "CODER / COCAP / CONEC — GERED / POSTAL SAÚDE"
    ])
    return "\n".join(linhas)


def gerar_pdf_capa_oficio(titulo_item, nome_demanda, texto_oficio, lista_arquivos_excel, lista_pdfs_anexados):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=2.2 * cm,
        rightMargin=2.2 * cm,
        topMargin=2.0 * cm,
        bottomMargin=2.0 * cm
    )

    styles = getSampleStyleSheet()
    st_header = ParagraphStyle(
        "HeaderPostal",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        textColor=colors.HexColor("#003366")
    )
    st_sub = ParagraphStyle(
        "SubPostal",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#475569")
    )
    st_secao = ParagraphStyle(
        "SecaoOficio",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10.5,
        leading=14,
        spaceBefore=8,
        spaceAfter=4,
        textColor=colors.HexColor("#003366")
    )
    st_corpo = ParagraphStyle(
        "CorpoOficio",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9.8,
        leading=13.5,
        spaceBefore=2,
        spaceAfter=5,
        textColor=colors.HexColor("#1E293B")
    )

    story = []
    story.append(Paragraph("POSTAL SAÚDE — CAIXA DE ASSISTÊNCIA E SAÚDE DOS EMPREGADOS DOS CORREIOS", st_header))
    story.append(Paragraph("PRASS — Plano de Recuperação Assistencial | Processo ANS nº 33910.030426/2026-21 | GERED", st_sub))
    story.append(Spacer(1, 6))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#003366")))
    story.append(Spacer(1, 10))

    for linha in texto_oficio.splitlines():
        l_strip = linha.strip()
        if not l_strip:
            story.append(Spacer(1, 6))
            continue
        if l_strip.startswith("---"):
            story.append(Spacer(1, 4))
            story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#CBD5E1")))
            story.append(Spacer(1, 4))
            continue

        l_safe = (
            l_strip.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

        if (
            l_strip.startswith("RESPOSTA AO OFÍCIO")
            or l_strip.startswith("ITEM DO PRASS:")
            or l_strip.startswith("1.")
            or l_strip.startswith("2.")
            or l_strip.startswith("3.")
            or l_strip.startswith("Evidência(s)")
            or l_strip.startswith("Situação:")
        ):
            story.append(Paragraph(l_safe, st_secao))
        else:
            story.append(Paragraph(l_safe, st_corpo))

    story.append(Spacer(1, 14))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#003366")))
    story.append(Spacer(1, 6))
    story.append(Paragraph("ÍNDICE DE ANEXOS CONSOLIDADOS NESTE DOSSIÊ", st_secao))

    if lista_pdfs_anexados:
        story.append(Paragraph("<b>A) Documentos PDF / Imagens unificados a seguir neste próprio arquivo PDF:</b>", st_corpo))
        for idx_p, info_p in enumerate(lista_pdfs_anexados, start=1):
            story.append(Paragraph(f"{idx_p}. [{info_p['setor']}] {info_p['nome_arquivo']} — ({info_p['titulo_chk']})", st_corpo))

    if lista_arquivos_excel:
        story.append(Spacer(1, 6))
        story.append(Paragraph("<b>B) Arquivos em formato Planilha Eletrônica (Excel / CSV) referenciados pelo nome original:</b>", st_corpo))
        for idx_x, info_x in enumerate(lista_arquivos_excel, start=1):
            story.append(Paragraph(
                f"{idx_x}. <b>{info_x['nome_arquivo']}</b> — Setor: {info_x['setor']} | Item: {info_x['titulo_chk']}",
                st_corpo
            ))

    doc.build(story)
    buffer.seek(0)
    return buffer


def converter_imagem_para_pdf_bytes(img_bytes):
    im = Image.open(io.BytesIO(img_bytes))
    if im.mode in ("RGBA", "P"):
        im = im.convert("RGB")
    out = io.BytesIO()
    im.save(out, format="PDF", resolution=100.0)
    out.seek(0)
    return out


# ==============================================================================
# ROTAS DE API E DIAGNÓSTICO (BACKEND FULL PYTHON)
# ==============================================================================
@app.route("/api/diagnostico", methods=["GET"])
def api_diagnostico():
    url_mascarada = ""
    if DATABASE_URL:
        partes = DATABASE_URL.split("@")
        url_mascarada = ("***@" + partes[-1]) if len(partes) > 1 else "definida"
    return jsonify({
        "postgres_ativo": USE_POSTGRES,
        "url_detectada": url_mascarada,
        "ultimo_erro_postgres": LAST_PG_ERROR,
        "sqlite_fallback_path": SQLITE_PATH
    })


@app.route("/api/bootstrap", methods=["GET"])
def api_bootstrap():
    try:
        itens = db_query("SELECT * FROM prass_items ORDER BY codigo")
        demanda_selecionada_nome = request.args.get("demanda_nome", "Complemento ANS 26/10")

        resultado_itens = []
        for it in itens:
            cod = it["codigo"]
            demandas = db_query(
                "SELECT * FROM demandas WHERE item_codigo = %s ORDER BY ordem ASC, id ASC",
                (cod,)
            )
            dem_ativa = next((d for d in demandas if d["nome"] == demanda_selecionada_nome), None)
            if not dem_ativa and demandas:
                dem_ativa = demandas[0]

            areas = json.loads(it["areas_atribuidas"] or '["CODER","COCAP","CONEC"]')
            semaforo = calcular_semaforo_setores(dem_ativa["id"], areas) if dem_ativa else {}

            resultado_itens.append({
                "codigo": cod,
                "titulo": it["titulo"],
                "apontamento_ans": it["apontamento_ans"],
                "medidas_implementadas": it["medidas_implementadas"],
                "observacoes": it["observacoes"],
                "historico_planilha": json.loads(it["historico_json"] or "{}"),
                "areas_atribuidas": areas,
                "demandas": demandas,
                "semaforo": semaforo
            })

        return jsonify({
            "itens": resultado_itens,
            "postgres_ativo": USE_POSTGRES,
            "ultimo_erro_pg": LAST_PG_ERROR
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({"erro": str(e)}), 500


@app.route("/api/item/<path:codigo>/areas", methods=["POST"])
def api_atualizar_areas(codigo):
    body = request.get_json(force=True)
    areas = body.get("areas_atribuidas", ["CODER", "COCAP", "CONEC"])
    db_query(
        "UPDATE prass_items SET areas_atribuidas = %s WHERE codigo = %s",
        (json.dumps(areas, ensure_ascii=False), codigo),
        fetch=False,
        commit=True
    )
    return jsonify({"ok": True, "areas_atribuidas": areas})


@app.route("/api/demandas", methods=["POST"])
def api_criar_demanda():
    body = request.get_json(force=True)
    item_codigo = body.get("item_codigo", "").strip()
    nome = body.get("nome", "").strip()
    replicar_todos = bool(body.get("replicar_todos", True))

    if not item_codigo or not nome:
        return jsonify({"erro": "Informe o código do item e o nome da demanda."}), 400

    codigos_alvo = [r["codigo"] for r in db_query("SELECT codigo FROM prass_items")] if replicar_todos else [item_codigo]
    for c in codigos_alvo:
        existe = db_query("SELECT id FROM demandas WHERE item_codigo = %s AND nome = %s", (c, nome))
        if not existe:
            db_query(
                "INSERT INTO demandas (item_codigo, nome, tipo, ordem, resumo_texto) VALUES (%s, %s, %s, %s, %s)",
                (c, nome, "AVULSA", 500, ""),
                fetch=False,
                commit=True
            )
    return jsonify({"ok": True, "nome": nome})


@app.route("/api/demanda/<int:demanda_id>/resumo", methods=["POST"])
def api_salvar_resumo_demanda(demanda_id):
    """Salva o texto da caixinha de Resumo (antes das evidências) no banco."""
    body = request.get_json(force=True)
    resumo_texto = body.get("resumo_texto", "").strip()
    db_query(
        "UPDATE demandas SET resumo_texto = %s WHERE id = %s",
        (resumo_texto, demanda_id),
        fetch=False,
        commit=True
    )
    return jsonify({"ok": True, "resumo_texto": resumo_texto})


@app.route("/api/checklist/<int:demanda_id>", methods=["GET"])
def api_listar_checklist(demanda_id):
    dem = db_query("SELECT * FROM demandas WHERE id = %s", (demanda_id,))
    if not dem:
        return jsonify({"erro": "Demanda não encontrada."}), 404
    dem_row = dem[0]
    it_row = db_query("SELECT * FROM prass_items WHERE codigo = %s", (dem_row["item_codigo"],))[0]

    chks = db_query(
        "SELECT * FROM checklist_evidencias WHERE demanda_id = %s ORDER BY id ASC",
        (demanda_id,)
    )
    for c in chks:
        c["concluido"] = bool(c["concluido"])
        anexos = db_query(
            """SELECT id, checklist_id, nome_arquivo, mime_type, tamanho, resumo_extraido, criado_em
               FROM anexos WHERE checklist_id = %s ORDER BY id ASC""",
            (c["id"],)
        )
        for a in anexos:
            a["criado_em"] = str(a["criado_em"])
        c["anexos"] = anexos
        c["criado_em"] = str(c["criado_em"])

    areas = json.loads(it_row["areas_atribuidas"] or '["CODER","COCAP","CONEC"]')
    semaforo = calcular_semaforo_setores(demanda_id, areas)

    return jsonify({
        "demanda": dem_row,
        "checklist": chks,
        "semaforo": semaforo
    })


@app.route("/api/checklist", methods=["POST"])
def api_adicionar_checklist():
    body = request.get_json(force=True)
    demanda_id = int(body["demanda_id"])
    setor = body["setor"].strip().upper()
    titulo = body["titulo"].strip()
    analise = body.get("analise_evidencia", "").strip()

    if not titulo:
        return jsonify({"erro": "Digite a descrição da evidência para o checklist."}), 400

    db_query(
        """INSERT INTO checklist_evidencias (demanda_id, setor, titulo, analise_evidencia, concluido)
           VALUES (%s, %s, %s, %s, %s)""",
        (demanda_id, setor, titulo, analise, False),
        fetch=False,
        commit=True
    )
    return jsonify({"ok": True})


@app.route("/api/checklist/<int:chk_id>", methods=["PUT"])
def api_editar_checklist(chk_id):
    body = request.get_json(force=True)
    if "titulo" in body:
        db_query(
            "UPDATE checklist_evidencias SET titulo = %s, analise_evidencia = %s WHERE id = %s",
            (body["titulo"].strip(), body.get("analise_evidencia", "").strip(), chk_id),
            fetch=False,
            commit=True
        )
    if "concluido" in body:
        db_query(
            "UPDATE checklist_evidencias SET concluido = %s WHERE id = %s",
            (bool(body["concluido"]), chk_id),
            fetch=False,
            commit=True
        )
    return jsonify({"ok": True})


@app.route("/api/checklist/<int:chk_id>", methods=["DELETE"])
def api_excluir_checklist(chk_id):
    db_query("DELETE FROM anexos WHERE checklist_id = %s", (chk_id,), fetch=False, commit=True)
    db_query("DELETE FROM checklist_evidencias WHERE id = %s", (chk_id,), fetch=False, commit=True)
    return jsonify({"ok": True})


@app.route("/api/checklist/<int:chk_id>/anexos", methods=["POST"])
def api_upload_anexos(chk_id):
    arquivos = request.files.getlist("arquivos")
    if not arquivos:
        return jsonify({"erro": "Nenhum arquivo enviado."}), 400

    conn, is_pg = get_db_conn()
    try:
        cur = conn.cursor()
        for f in arquivos:
            if not f.filename:
                continue
            raw_bytes = f.read()
            nome_arq = f.filename
            mime = f.mimetype or "application/octet-stream"
            tamanho = len(raw_bytes)
            resumo = extrair_resumo_arquivo(nome_arq, raw_bytes)

            if is_pg:
                cur.execute(
                    """INSERT INTO anexos (checklist_id, nome_arquivo, mime_type, tamanho, resumo_extraido, conteudo)
                       VALUES (%s, %s, %s, %s, %s, %s)""",
                    (chk_id, nome_arq, mime, tamanho, resumo, psycopg2.Binary(raw_bytes))
                )
            else:
                cur.execute(
                    """INSERT INTO anexos (checklist_id, nome_arquivo, mime_type, tamanho, resumo_extraido, conteudo)
                       VALUES (?, ?, ?, ?, ?, ?)""",
                    (chk_id, nome_arq, mime, tamanho, resumo, sqlite3.Binary(raw_bytes))
                )
        conn.commit()
    finally:
        conn.close()

    return jsonify({"ok": True})


@app.route("/api/anexos/<int:anexo_id>", methods=["DELETE"])
def api_remover_anexo(anexo_id):
    row = db_query("SELECT checklist_id FROM anexos WHERE id = %s", (anexo_id,))
    if row:
        chk_id = row[0]["checklist_id"]
        db_query("DELETE FROM anexos WHERE id = %s", (anexo_id,), fetch=False, commit=True)
        db_query(
            "UPDATE checklist_evidencias SET concluido = %s WHERE id = %s",
            (False, chk_id),
            fetch=False,
            commit=True
        )
    return jsonify({"ok": True})


@app.route("/api/anexos/<int:anexo_id>/view", methods=["GET"])
def api_visualizar_anexo(anexo_id):
    row = db_query("SELECT nome_arquivo, mime_type, conteudo FROM anexos WHERE id = %s", (anexo_id,))
    if not row:
        return "Arquivo não encontrado", 404
    r = row[0]
    dados = bytes(r["conteudo"])
    return send_file(
        io.BytesIO(dados),
        mimetype=r["mime_type"],
        as_attachment=False,
        download_name=r["nome_arquivo"]
    )


@app.route("/api/anexos/<int:anexo_id>/download", methods=["GET"])
def api_baixar_anexo(anexo_id):
    row = db_query("SELECT nome_arquivo, mime_type, conteudo FROM anexos WHERE id = %s", (anexo_id,))
    if not row:
        return "Arquivo não encontrado", 404
    r = row[0]
    dados = bytes(r["conteudo"])
    return send_file(
        io.BytesIO(dados),
        mimetype=r["mime_type"],
        as_attachment=True,
        download_name=r["nome_arquivo"]
    )


@app.route("/api/demanda/<int:demanda_id>/gerar-minuta-oficio", methods=["POST"])
def api_gerar_minuta_oficio(demanda_id):
    body = request.get_json(silent=True) or {}
    if "resumo_texto" in body:
        db_query(
            "UPDATE demandas SET resumo_texto = %s WHERE id = %s",
            (body["resumo_texto"].strip(), demanda_id),
            fetch=False,
            commit=True
        )

    dem = db_query("SELECT * FROM demandas WHERE id = %s", (demanda_id,))
    if not dem:
        return jsonify({"erro": "Demanda não encontrada."}), 404
    dem_row = dem[0]
    it_row = db_query("SELECT * FROM prass_items WHERE codigo = %s", (dem_row["item_codigo"],))[0]

    chks = db_query(
        "SELECT * FROM checklist_evidencias WHERE demanda_id = %s ORDER BY setor ASC, id ASC",
        (demanda_id,)
    )
    arquivos_pdf = []
    arquivos_excel = []
    for c in chks:
        c["concluido"] = bool(c["concluido"])
        anexos = db_query(
            "SELECT id, nome_arquivo, mime_type, tamanho, resumo_extraido FROM anexos WHERE checklist_id = %s ORDER BY id ASC",
            (c["id"],)
        )
        c["anexos"] = anexos
        for a in anexos:
            ext = os.path.splitext(a["nome_arquivo"].lower())[1]
            if ext in (".xlsx", ".xls", ".csv"):
                arquivos_excel.append({"setor": c["setor"], "nome_arquivo": a["nome_arquivo"], "titulo_chk": c["titulo"]})
            else:
                arquivos_pdf.append({"setor": c["setor"], "nome_arquivo": a["nome_arquivo"], "titulo_chk": c["titulo"]})

    forcar_regeneracao = bool(body.get("regenerar", False))

    if dem_row.get("oficio_texto") and dem_row["oficio_texto"].strip() and not forcar_regeneracao:
        texto_minuta = dem_row["oficio_texto"]
    else:
        texto_minuta = construir_texto_oficio_automatico(it_row, dem_row, chks)
        db_query(
            "UPDATE demandas SET oficio_texto = %s WHERE id = %s",
            (texto_minuta, demanda_id),
            fetch=False,
            commit=True
        )

    return jsonify({
        "texto_oficio": texto_minuta,
        "arquivos_pdf": arquivos_pdf,
        "arquivos_excel": arquivos_excel,
        "item_codigo": it_row["codigo"],
        "demanda_nome": dem_row["nome"]
    })


@app.route("/api/demanda/<int:demanda_id>/salvar-oficio", methods=["POST"])
def api_salvar_oficio(demanda_id):
    body = request.get_json(force=True)
    texto = body.get("texto_oficio", "")
    db_query(
        "UPDATE demandas SET oficio_texto = %s WHERE id = %s",
        (texto, demanda_id),
        fetch=False,
        commit=True
    )
    return jsonify({"ok": True})


@app.route("/api/demanda/<int:demanda_id>/exportar-word", methods=["POST"])
def api_exportar_word(demanda_id):
    """Permite baixar o Ofício aprovado também em formato Word (.doc) formatado."""
    body = request.get_json(force=True)
    texto_oficio = body.get("texto_oficio", "").strip()
    dem_row = db_query("SELECT * FROM demandas WHERE id = %s", (demanda_id,))[0]
    it_row = db_query("SELECT * FROM prass_items WHERE codigo = %s", (dem_row["item_codigo"],))[0]

    paragrafos_html = ""
    for linha in texto_oficio.splitlines():
        l_strip = linha.strip()
        if not l_strip:
            continue
        if l_strip.startswith("---"):
            paragrafos_html += "<hr style='border:0;border-top:1px solid #CBD5E1;margin:12pt 0;'>"
        elif l_strip.startswith(("RESPOSTA AO OFÍCIO", "ITEM DO PRASS:", "1.", "2.", "Evidência(s)", "Situação:")):
            paragrafos_html += f"<p style='font-weight:bold;color:#003366;margin-top:10pt;margin-bottom:4pt;'>{l_strip}</p>"
        else:
            paragrafos_html += f"<p style='text-align:justify;margin-top:2pt;margin-bottom:6pt;'>{l_strip}</p>"

    html_doc = f"""<html xmlns:o='urn:schemas-microsoft-com:office:office'
    xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head><meta charset='utf-8'><title>Ofício PRASS {it_row['codigo']}</title></head>
    <body style="font-family:'Calibri','Arial',sans-serif;font-size:11pt;line-height:1.45;">
        <div style="border-bottom:2px solid #003366;padding-bottom:6pt;margin-bottom:14pt;">
            <p style="font-size:13pt;font-weight:bold;color:#003366;margin:0;">POSTAL SAÚDE — CAIXA DE ASSISTÊNCIA E SAÚDE DOS EMPREGADOS DOS CORREIOS</p>
            <p style="font-size:9pt;color:#555555;margin:0;">PRASS — Processo ANS nº 33910.030426/2026-21 | GERED</p>
        </div>
        {paragrafos_html}
    </body></html>"""

    nome_doc = f"Oficio_PRASS_{it_row['codigo'].replace('.', '_')}_{dem_row['nome'].replace('/', '-').replace(' ', '_')}.doc"
    return send_file(
        io.BytesIO(("\ufeff" + html_doc).encode("utf-8")),
        mimetype="application/msword",
        as_attachment=True,
        download_name=nome_doc
    )


@app.route("/api/demanda/<int:demanda_id>/exportar-pdf-unificado", methods=["POST"])
def api_exportar_pdf_unificado(demanda_id):
    body = request.get_json(force=True)
    texto_oficio = body.get("texto_oficio", "").strip()

    dem = db_query("SELECT * FROM demandas WHERE id = %s", (demanda_id,))
    if not dem:
        return jsonify({"erro": "Demanda não encontrada."}), 404
    dem_row = dem[0]
    it_row = db_query("SELECT * FROM prass_items WHERE codigo = %s", (dem_row["item_codigo"],))[0]

    db_query(
        "UPDATE demandas SET oficio_texto = %s WHERE id = %s",
        (texto_oficio, demanda_id),
        fetch=False,
        commit=True
    )

    chks = db_query(
        "SELECT * FROM checklist_evidencias WHERE demanda_id = %s ORDER BY setor ASC, id ASC",
        (demanda_id,)
    )

    lista_pdfs_info = []
    lista_excel_info = []
    anexos_para_merge = []

    for c in chks:
        anexos = db_query(
            "SELECT id, nome_arquivo, mime_type, conteudo FROM anexos WHERE checklist_id = %s ORDER BY id ASC",
            (c["id"],)
        )
        for a in anexos:
            ext = os.path.splitext(a["nome_arquivo"].lower())[1]
            if ext in (".xlsx", ".xls", ".csv"):
                lista_excel_info.append({
                    "setor": c["setor"],
                    "nome_arquivo": a["nome_arquivo"],
                    "titulo_chk": c["titulo"]
                })
            elif ext in (".pdf", ".png", ".jpg", ".jpeg"):
                lista_pdfs_info.append({
                    "setor": c["setor"],
                    "nome_arquivo": a["nome_arquivo"],
                    "titulo_chk": c["titulo"]
                })
                anexos_para_merge.append((ext, a["nome_arquivo"], bytes(a["conteudo"])))

    capa_buffer = gerar_pdf_capa_oficio(
        it_row["titulo"],
        dem_row["nome"],
        texto_oficio,
        lista_excel_info,
        lista_pdfs_info
    )

    writer = pypdf.PdfWriter()
    capa_reader = pypdf.PdfReader(capa_buffer)
    for page in capa_reader.pages:
        writer.add_page(page)

    for ext, nome_arq, raw_bytes in anexos_para_merge:
        try:
            if ext == ".pdf":
                r_anexo = pypdf.PdfReader(io.BytesIO(raw_bytes))
                for p in r_anexo.pages:
                    writer.add_page(p)
            elif ext in (".png", ".jpg", ".jpeg"):
                pdf_img_buf = converter_imagem_para_pdf_bytes(raw_bytes)
                r_img = pypdf.PdfReader(pdf_img_buf)
                for p in r_img.pages:
                    writer.add_page(p)
        except Exception as err:
            print(f"Aviso ao mesclar {nome_arq}: {err}")

    pdf_final = io.BytesIO()
    writer.write(pdf_final)
    pdf_final.seek(0)

    nome_download = f"Dossie_Oficio_PRASS_{it_row['codigo'].replace('.', '_')}_{dem_row['nome'].replace('/', '-').replace(' ', '_')}.pdf"
    return send_file(
        io.BytesIO(pdf_final.getvalue()),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=nome_download
    )


# ==============================================================================
# FRONTEND WEB COMPLETO EMBUTIDO (HTML5 + CSS3 + JAVASCRIPT VANILLA)
# ==============================================================================
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Postal Saúde | Sistema de Automação PRASS — GERED (CODER • COCAP • CONEC)</title>
<style>
    :root {
        --azul-postal: #003366;
        --azul-escuro: #001F3F;
        --amarelo-postal: #F2A900;
        --amarelo-claro: #FFF8E1;
        --bg-geral: #F0F4F8;
        --card-bg: #FFFFFF;
        --texto: #0F172A;
        --cinza-borda: #CBD5E1;
        --verde: #198754;
        --verde-bg: #D1E7DD;
        --amarelo-status: #B45309;
        --amarelo-status-bg: #FEF3C7;
        --vermelho: #DC3545;
        --vermelho-bg: #F8D7DA;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
        font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
        background-color: var(--bg-geral);
        color: var(--texto);
        display: flex;
        flex-direction: column;
        height: 100vh;
        overflow: hidden;
    }
    header {
        background: linear-gradient(90deg, var(--azul-escuro) 0%, var(--azul-postal) 100%);
        color: #FFF;
        padding: 12px 24px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        border-bottom: 4px solid var(--amarelo-postal);
        flex-shrink: 0;
    }
    .brand-title { font-size: 18px; font-weight: 800; letter-spacing: 0.3px; }
    .brand-sub { font-size: 12px; color: #FFD100; margin-top: 2px; }
    .db-badge {
        background: rgba(255,255,255,0.12);
        border: 1px solid rgba(255,255,255,0.25);
        padding: 6px 12px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 600;
        cursor: pointer;
    }
    .main-container {
        display: flex;
        flex: 1;
        overflow: hidden;
    }
    .sidebar {
        width: 390px;
        background: #FFFFFF;
        border-right: 1px solid var(--cinza-borda);
        display: flex;
        flex-direction: column;
        flex-shrink: 0;
    }
    .sidebar-header {
        padding: 14px 16px;
        background: #F8FAFC;
        border-bottom: 1px solid var(--cinza-borda);
        font-weight: 700;
        color: var(--azul-postal);
        font-size: 13px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .items-list {
        flex: 1;
        overflow-y: auto;
        padding: 10px;
    }
    .prass-item-card {
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 12px;
        margin-bottom: 10px;
        cursor: pointer;
        transition: all 0.15s ease;
        background: #FFFFFF;
        position: relative;
    }
    .prass-item-card:hover {
        border-color: #94A3B8;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
    }
    .prass-item-card.active {
        border: 2px solid var(--azul-postal);
        background: #F0F7FF;
    }
    .item-code-badge {
        display: inline-block;
        background: var(--azul-postal);
        color: #FFF;
        font-weight: 800;
        font-size: 11px;
        padding: 2px 8px;
        border-radius: 4px;
        margin-bottom: 6px;
    }
    .prass-item-card.active .item-code-badge {
        background: var(--amarelo-postal);
        color: var(--azul-escuro);
    }
    .item-title-short {
        font-size: 12.5px;
        font-weight: 700;
        color: #1E293B;
        line-height: 1.35;
        margin-bottom: 8px;
    }
    .semaforo-row {
        display: flex;
        align-items: center;
        gap: 6px;
        flex-wrap: wrap;
        margin-top: 6px;
    }
    .area-pill {
        font-size: 10.5px;
        font-weight: 700;
        padding: 2px 8px;
        border-radius: 12px;
        display: inline-flex;
        align-items: center;
        gap: 4px;
    }
    .area-pill.VERMELHO { background: var(--vermelho-bg); color: #842029; border: 1px solid #F5C2C7; }
    .area-pill.AMARELO { background: var(--amarelo-status-bg); color: #92400E; border: 1px solid #FDE68A; }
    .area-pill.VERDE { background: var(--verde-bg); color: #0F5132; border: 1px solid #BADBCC; }
    .area-pill.INATIVO { background: #F1F5F9; color: #94A3B8; border: 1px dashed #CBD5E1; text-decoration: line-through; }
    .workspace {
        flex: 1;
        overflow-y: auto;
        padding: 18px 24px;
        display: flex;
        flex-direction: column;
        gap: 16px;
    }
    .top-item-banner {
        background: #FFFFFF;
        border: 1px solid var(--cinza-borda);
        border-radius: 10px;
        padding: 16px 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    .banner-header-row {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 16px;
    }
    .banner-title {
        font-size: 16px;
        font-weight: 800;
        color: var(--azul-postal);
        line-height: 1.35;
    }
    .banner-actions {
        display: flex;
        gap: 8px;
        flex-shrink: 0;
        align-items: center;
    }
    .dropdown-areas {
        position: relative;
        display: inline-block;
    }
    .dropdown-btn {
        background: #F8FAFC;
        border: 1px solid var(--cinza-borda);
        padding: 7px 12px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 700;
        color: var(--azul-postal);
        cursor: pointer;
    }
    .dropdown-menu {
        display: none;
        position: absolute;
        right: 0;
        top: 110%;
        background: #FFF;
        border: 1px solid var(--cinza-borda);
        border-radius: 8px;
        box-shadow: 0 8px 20px rgba(0,0,0,0.12);
        padding: 10px 14px;
        z-index: 100;
        min-width: 210px;
    }
    .dropdown-menu.show { display: block; }
    .dropdown-menu label {
        display: flex;
        align-items: center;
        gap: 8px;
        font-size: 12.5px;
        font-weight: 600;
        padding: 5px 0;
        cursor: pointer;
    }
    .planilha-buttons-row {
        display: flex;
        gap: 10px;
        margin-top: 14px;
        flex-wrap: wrap;
    }
    .btn-info-planilha {
        flex: 1;
        min-width: 220px;
        padding: 9px 14px;
        border-radius: 7px;
        border: 1px solid #CBD5E1;
        background: #F8FAFC;
        color: #1E293B;
        font-weight: 700;
        font-size: 12.5px;
        cursor: pointer;
        text-align: left;
        display: flex;
        align-items: center;
        justify-content: space-between;
        transition: all 0.15s;
    }
    .btn-info-planilha:hover { background: #E2E8F0; }
    .btn-info-planilha.active {
        background: var(--azul-postal);
        color: #FFFFFF;
        border-color: var(--azul-postal);
    }
    .info-drawer {
        display: none;
        margin-top: 12px;
        background: #F8FAFC;
        border-left: 4px solid var(--azul-postal);
        border-radius: 6px;
        padding: 14px 18px;
        font-size: 13px;
        line-height: 1.55;
        white-space: pre-wrap;
        color: #1E293B;
        max-height: 280px;
        overflow-y: auto;
    }
    .info-drawer.open { display: block; }
    .controls-bar {
        background: #FFFFFF;
        border: 1px solid var(--cinza-borda);
        border-radius: 10px;
        padding: 14px 20px;
        display: flex;
        flex-wrap: wrap;
        gap: 18px;
        align-items: center;
        justify-content: space-between;
    }
    .control-group {
        display: flex;
        align-items: center;
        gap: 8px;
        flex-wrap: wrap;
    }
    .control-label {
        font-size: 12px;
        font-weight: 800;
        text-transform: uppercase;
        color: #475569;
    }
    .seg-btn {
        padding: 7px 14px;
        border-radius: 6px;
        border: 1px solid var(--cinza-borda);
        background: #F8FAFC;
        font-size: 12.5px;
        font-weight: 700;
        cursor: pointer;
        color: #334155;
    }
    .seg-btn.active {
        background: var(--azul-postal);
        color: #FFFFFF;
        border-color: var(--azul-postal);
    }
    select.select-demanda {
        padding: 7px 12px;
        border-radius: 6px;
        border: 1px solid var(--azul-postal);
        font-size: 13px;
        font-weight: 700;
        color: var(--azul-postal);
        background: #FFF;
        min-width: 220px;
    }
    .setor-btn {
        padding: 8px 18px;
        border-radius: 7px;
        border: 2px solid transparent;
        font-size: 13px;
        font-weight: 800;
        cursor: pointer;
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #E2E8F0;
        color: #1E293B;
    }
    .setor-btn.active {
        background: var(--azul-postal);
        color: #FFFFFF;
        border-color: var(--amarelo-postal);
    }

    /* CAIXINHA DE RESUMO EM TEXTO (ANTES DAS EVIDÊNCIAS) */
    .resumo-before-box {
        background: #F8FAFC;
        border: 1px solid #93C5FD;
        border-left: 5px solid var(--azul-postal);
        border-radius: 8px;
        padding: 14px 16px;
        display: flex;
        flex-direction: column;
        gap: 8px;
    }
    .resumo-header-line {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 10px;
        flex-wrap: wrap;
    }
    .resumo-textarea {
        width: 100%;
        min-height: 110px;
        padding: 10px 12px;
        border: 1px solid #CBD5E1;
        border-radius: 6px;
        font-family: 'Segoe UI', system-ui, sans-serif;
        font-size: 13px;
        line-height: 1.5;
        color: #0F172A;
        resize: vertical;
        background: #FFFFFF;
    }
    .resumo-textarea:focus {
        outline: none;
        border-color: var(--azul-postal);
        box-shadow: 0 0 0 2px rgba(0, 51, 102, 0.15);
    }

    .checklist-panel {
        background: #FFFFFF;
        border: 1px solid var(--cinza-borda);
        border-radius: 10px;
        padding: 18px 20px;
        display: flex;
        flex-direction: column;
        gap: 14px;
    }
    .chk-header-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 12px;
        flex-wrap: wrap;
    }
    .progress-wrap {
        flex: 1;
        min-width: 260px;
    }
    .progress-bar-outer {
        width: 100%;
        height: 12px;
        background: #E2E8F0;
        border-radius: 6px;
        overflow: hidden;
        margin-top: 5px;
    }
    .progress-bar-inner {
        height: 100%;
        background: var(--verde);
        width: 0%;
        transition: width 0.25s ease;
    }
    .add-chk-row {
        display: flex;
        gap: 8px;
        background: #F8FAFC;
        padding: 10px;
        border-radius: 8px;
        border: 1px dashed #94A3B8;
    }
    .add-chk-row input {
        flex: 1;
        padding: 8px 12px;
        border: 1px solid #CBD5E1;
        border-radius: 6px;
        font-size: 13px;
    }
    .btn-primary {
        background: var(--azul-postal);
        color: #FFF;
        border: none;
        padding: 8px 14px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 12.5px;
        cursor: pointer;
    }
    .btn-success {
        background: var(--verde);
        color: #FFF;
        border: none;
        padding: 8px 14px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 12.5px;
        cursor: pointer;
    }
    .btn-warning {
        background: var(--amarelo-postal);
        color: var(--azul-escuro);
        border: none;
        padding: 8px 14px;
        border-radius: 6px;
        font-weight: 800;
        font-size: 12.5px;
        cursor: pointer;
    }
    .chk-item-card {
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 12px 14px;
        background: #FFFFFF;
        display: flex;
        flex-direction: column;
        gap: 8px;
    }
    .chk-item-card.concluido {
        border-left: 5px solid var(--verde);
        background: #F8FFF9;
    }
    .chk-item-card.pendente {
        border-left: 5px solid var(--vermelho);
    }
    .chk-top-line {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 10px;
    }
    .chk-title-group {
        display: flex;
        align-items: center;
        gap: 10px;
        flex: 1;
    }
    .chk-title-text {
        font-size: 13.5px;
        font-weight: 700;
        color: #0F172A;
    }
    .chk-tools {
        display: flex;
        align-items: center;
        gap: 6px;
        flex-wrap: wrap;
    }
    .tool-btn {
        border: 1px solid #CBD5E1;
        background: #F8FAFC;
        color: #1E293B;
        padding: 5px 9px;
        border-radius: 5px;
        font-size: 11.5px;
        font-weight: 700;
        cursor: pointer;
    }
    .tool-btn:hover { background: #E2E8F0; }
    .tool-btn.concluir-on { background: var(--verde); color: #FFF; border-color: var(--verde); }
    .tool-btn.danger { background: #FEE2E2; color: #991B1B; border-color: #FCA5A5; }
    .anexos-box {
        background: #F8FAFC;
        border-radius: 6px;
        padding: 8px 10px;
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        align-items: center;
    }
    .anexo-chip {
        background: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 6px;
        padding: 4px 8px;
        font-size: 11.5px;
        display: inline-flex;
        align-items: center;
        gap: 6px;
    }
    .anexo-chip a {
        text-decoration: none;
        font-weight: 700;
        color: var(--azul-postal);
    }
    .modal-backdrop {
        display: none;
        position: fixed;
        inset: 0;
        background: rgba(15, 23, 42, 0.65);
        z-index: 1000;
        align-items: center;
        justify-content: center;
        padding: 20px;
    }
    .modal-backdrop.open { display: flex; }
    .modal-window {
        background: #FFFFFF;
        width: 100%;
        max-width: 980px;
        max-height: 90vh;
        border-radius: 12px;
        display: flex;
        flex-direction: column;
        overflow: hidden;
        box-shadow: 0 20px 50px rgba(0,0,0,0.3);
    }
    .modal-header {
        background: var(--azul-postal);
        color: #FFF;
        padding: 14px 20px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .modal-body {
        padding: 18px 20px;
        overflow-y: auto;
        flex: 1;
        display: flex;
        flex-direction: column;
        gap: 12px;
    }
    .oficio-textarea {
        width: 100%;
        min-height: 360px;
        font-family: 'Consolas', 'Courier New', monospace;
        font-size: 12.5px;
        line-height: 1.5;
        padding: 14px;
        border: 1px solid #94A3B8;
        border-radius: 8px;
        resize: vertical;
    }
    .modal-footer {
        background: #F1F5F9;
        padding: 12px 20px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-top: 1px solid #CBD5E1;
    }
</style>
</head>
<body>

<header>
    <div>
        <div class="brand-title">POSTAL SAÚDE | SISTEMA DE AUTOMAÇÃO PRASS — GERED</div>
        <div class="brand-sub">Módulos 8.1.2 a 8.6.1 • Coordenações CODER, COCAP e CONEC • Resumo + Análise Automática de Ofício + Dossiê PDF Unificado</div>
    </div>
    <div class="db-badge" id="dbBadge" onclick="verDiagnosticoBanco()" title="Clique para ver detalhes da conexão com o banco">Conectando ao Banco...</div>
</header>

<div class="main-container">
    <aside class="sidebar">
        <div class="sidebar-header">
            <span>MÓDULOS DO PRASS (GERED)</span>
            <span id="lblDemandaSidebar" style="font-size:11px; color:#0D6EFD;"></span>
        </div>
        <div class="items-list" id="sidebarItems"></div>
    </aside>

    <main class="workspace">
        <section class="top-item-banner">
            <div class="banner-header-row">
                <div>
                    <div class="banner-title" id="mainItemTitulo">Carregando...</div>
                    <div id="mainSemaforoBadges" class="semaforo-row" style="margin-top:8px;"></div>
                </div>
                <div class="banner-actions">
                    <div class="dropdown-areas">
                        <button class="dropdown-btn" onclick="toggleAreasDropdown()">⚙️ Áreas Atribuídas ▼</button>
                        <div class="dropdown-menu" id="areasDropdownMenu">
                            <div style="font-size:11px; color:#64748B; margin-bottom:6px; font-weight:700;">Atribuir a este Item do PRASS:</div>
                            <label><input type="checkbox" value="CODER" onchange="salvarAreasAtribuidas()"> CODER</label>
                            <label><input type="checkbox" value="COCAP" onchange="salvarAreasAtribuidas()"> COCAP</label>
                            <label><input type="checkbox" value="CONEC" onchange="salvarAreasAtribuidas()"> CONEC</label>
                        </div>
                    </div>

                    <button class="btn-warning" onclick="abrirModalOficio()">
                        📄 Gerar Ofício + Unificar PDF
                    </button>
                </div>
            </div>

            <div class="planilha-buttons-row">
                <button class="btn-info-planilha" id="btnDrawerApontamento" onclick="toggleDrawer('apontamento')">
                    <span>📌 1. Apontamento da ANS & Diagnóstico (Col. B)</span>
                    <span>▼</span>
                </button>
                <button class="btn-info-planilha" id="btnDrawerMedidas" onclick="toggleDrawer('medidas')">
                    <span>🛠️ 2. Medidas Implementadas e em Implementação (Col. B)</span>
                    <span>▼</span>
                </button>
                <button class="btn-info-planilha" id="btnDrawerObs" onclick="toggleDrawer('obs')">
                    <span>👁️ 3. Observações — Tratamento e Monitoramento (Col. C)</span>
                    <span>▼</span>
                </button>
            </div>

            <div class="info-drawer" id="drawerContent"></div>
        </section>

        <section class="controls-bar">
            <div class="control-group">
                <span class="control-label">Tipo:</span>
                <button class="seg-btn active" id="btnTipoFixa" onclick="selecionarTipoDemanda('FIXA')">Demanda Fixa (08/26 a 08/27)</button>
                <button class="seg-btn" id="btnTipoAvulsa" onclick="selecionarTipoDemanda('AVULSA')">Demanda Avulsa</button>
            </div>

            <div class="control-group">
                <span class="control-label">Competência / Demanda:</span>
                <select class="select-demanda" id="selectDemanda" onchange="aoMudarDemanda()"></select>
                <button class="btn-primary" onclick="criarNovaDemandaAvulsa()">+ Nova Demanda</button>
            </div>

            <div class="control-group">
                <span class="control-label">Setor (GERED):</span>
                <button class="setor-btn" id="btnSetorCODER" onclick="selecionarSetor('CODER')">CODER</button>
                <button class="setor-btn active" id="btnSetorCOCAP" onclick="selecionarSetor('COCAP')">COCAP</button>
                <button class="setor-btn" id="btnSetorCONEC" onclick="selecionarSetor('CONEC')">CONEC</button>
            </div>
        </section>

        <section class="checklist-panel">
            <!-- CAIXINHA DE RESUMO EM TEXTO ANTES DAS EVIDÊNCIAS (ENTRA NO TOPO DO OFÍCIO ANTES DA ANÁLISE) -->
            <div class="resumo-before-box">
                <div class="resumo-header-line">
                    <div>
                        <span style="font-size:13.5px; font-weight:800; color:var(--azul-postal);">
                            📝 Resumo Executivo / Considerações Iniciais (Antes das Evidências)
                        </span>
                        <span style="font-size:11.5px; color:#475569; margin-left:6px;">
                            • Este texto sairá na Seção 1 do Ofício, logo antes da análise técnica das evidências abaixo.
                        </span>
                    </div>
                    <div style="display:flex; gap:6px; align-items:center;">
                        <span id="statusResumoSalvo" style="font-size:11.5px; color:var(--verde); font-weight:700;"></span>
                        <button class="tool-btn" onclick="importarResumoPlanilhaBase()" title="Carregar o texto integral desta competência registrado na planilha inicial">
                            📥 Puxar Texto da Planilha Base
                        </button>
                        <button class="btn-primary" style="padding:5px 12px; font-size:12px;" onclick="salvarResumoDemanda(true)">
                            💾 Salvar Resumo
                        </button>
                    </div>
                </div>
                <textarea
                    id="textareaResumoDemanda"
                    class="resumo-textarea"
                    placeholder="Digite aqui o resumo geral deste item do PRASS antes das evidências (ele sairá automaticamente no início do Ofício, antes da análise detalhada dos anexos)..."
                    onblur="salvarResumoDemanda(false)"
                ></textarea>
            </div>

            <div class="chk-header-row">
                <div>
                    <h3 id="tituloChecklistSetor" style="color:var(--azul-postal); font-size:15px;">Checklist de Evidências — COCAP</h3>
                </div>
                <div class="progress-wrap">
                    <div style="display:flex; justify-content:space-between; font-size:12px; font-weight:700;">
                        <span id="lblProgressoSetor">Progresso do Setor: 0/0 (0%)</span>
                        <span id="lblProgressoDemandaGeral" style="color:var(--azul-postal);">Total Geral da Demanda: 0%</span>
                    </div>
                    <div class="progress-bar-outer">
                        <div class="progress-bar-inner" id="barraProgressoSetor"></div>
                    </div>
                </div>
            </div>

            <div class="add-chk-row">
                <input type="text" id="inputNovaEvidencia" placeholder="Digite o título da nova evidência para o checklist deste setor e pressione Enter ou clique em (+)..." onkeydown="if(event.key==='Enter') adicionarEvidencia()">
                <button class="btn-success" onclick="adicionarEvidencia()">+ Adicionar Evidência</button>
            </div>

            <div id="listaChecklistContainer" style="display:flex; flex-direction:column; gap:10px;"></div>
        </section>
    </main>
</div>

<div class="modal-backdrop" id="modalOficio">
    <div class="modal-window">
        <div class="modal-header">
            <div>
                <div style="font-size:15px; font-weight:800;">📄 ANÁLISE AUTOMÁTICA DE EVIDÊNCIAS & MINUTA DO OFÍCIO</div>
                <div style="font-size:11.5px; color:#FFD100;" id="modalSubtitulo">O Resumo foi posicionado no início (Seção 1), seguido da Análise Técnica de todas as evidências (Seção 2).</div>
            </div>
            <button onclick="fecharModalOficio()" style="background:none; border:none; color:#FFF; font-size:20px; cursor:pointer;">✖</button>
        </div>
        <div class="modal-body">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:12.5px; font-weight:700; color:#334155;">Texto do Ofício (Editável pelo Analista):</span>
                <button class="tool-btn" onclick="regenerarTextoOficio()">🔄 Atualizar Ofício com Resumo + Anexos Atuais</button>
            </div>
            <textarea class="oficio-textarea" id="textareaOficio"></textarea>
            <div id="resumoArquivosModal" style="background:#F8FAFC; border:1px solid #CBD5E1; border-radius:6px; padding:10px 14px; font-size:12px;"></div>
        </div>
        <div class="modal-footer">
            <div style="display:flex; gap:8px;">
                <button class="tool-btn" onclick="salvarRascunhoOficio()">💾 Salvar Edição no Banco</button>
                <button class="btn-primary" onclick="baixarOficioEmWord()">📝 Baixar Ofício em Word (.doc)</button>
            </div>
            <div style="display:flex; gap:10px;">
                <button class="tool-btn" onclick="fecharModalOficio()">Cancelar</button>
                <button class="btn-success" onclick="aprovarEBaixarPDFUnificado()">✅ Aprovar Ofício e Baixar PDF Unificado (Ofício + Anexos)</button>
            </div>
        </div>
    </div>
</div>

<input type="file" id="hiddenFileInput" multiple style="display:none" onchange="executarUploadAnexos()">

<script>
let ESTADO_ITENS = [];
let ITEM_ATUAL = null;
let TIPO_DEMANDA_ATUAL = 'FIXA';
let DEMANDA_ATUAL_ID = null;
let DEMANDA_ATUAL_NOME = 'Complemento ANS 26/10';
let SETOR_ATUAL = 'COCAP';
let DRAWER_ABERTO = null;
let CHECKLIST_ATUAL = [];
let CHK_ID_PARA_UPLOAD = null;

async function verDiagnosticoBanco() {
    const r = await fetch('/api/diagnostico');
    const d = await r.json();
    alert(
        "DIAGNÓSTICO DE BANCO DE DADOS:\n\n" +
        "• PostgreSQL Ativo: " + (d.postgres_ativo ? "SIM" : "NÃO (Usando Banco Persistente Interno)") + "\n" +
        "• URL Detectada: " + (d.url_detectada || "(Nenhuma)") + "\n" +
        "• Aviso Postgres: " + (d.ultimo_erro_postgres || "Nenhum erro")
    );
}

async function carregarSistema() {
    try {
        const resp = await fetch(`/api/bootstrap?demanda_nome=${encodeURIComponent(DEMANDA_ATUAL_NOME)}`);
        const dados = await resp.json();
        if (!resp.ok || dados.erro) {
            throw new Error(dados.erro || 'Falha HTTP ' + resp.status);
        }
        ESTADO_ITENS = dados.itens;

        const badge = document.getElementById('dbBadge');
        if (dados.postgres_ativo) {
            badge.textContent = '🟢 PostgreSQL Railway Ativo';
        } else {
            badge.textContent = '🟢 Banco Persistente Ativo (Clique p/ Info)';
        }

        if (!ITEM_ATUAL && ESTADO_ITENS.length > 0) {
            ITEM_ATUAL = ESTADO_ITENS.find(i => i.codigo === '8.5.1') || ESTADO_ITENS[0];
        } else if (ITEM_ATUAL) {
            ITEM_ATUAL = ESTADO_ITENS.find(i => i.codigo === ITEM_ATUAL.codigo) || ESTADO_ITENS[0];
        }

        renderizarSidebar();
        renderizarItemAtual();
    } catch (e) {
        console.error(e);
        document.getElementById('dbBadge').textContent = '🔴 Erro: ' + e.message;
    }
}

function renderizarSidebar() {
    const container = document.getElementById('sidebarItems');
    document.getElementById('lblDemandaSidebar').textContent = DEMANDA_ATUAL_NOME;
    container.innerHTML = '';

    ESTADO_ITENS.forEach(item => {
        const card = document.createElement('div');
        card.className = 'prass-item-card' + (ITEM_ATUAL && ITEM_ATUAL.codigo === item.codigo ? ' active' : '');
        card.onclick = () => {
            ITEM_ATUAL = item;
            renderizarSidebar();
            renderizarItemAtual();
        };

        let badgesHtml = '';
        ['CODER', 'COCAP', 'CONEC'].forEach(setor => {
            const st = item.semaforo && item.semaforo[setor];
            if (!item.areas_atribuidas.includes(setor)) {
                badgesHtml += `<span class="area-pill INATIVO">${setor}</span>`;
            } else if (st) {
                const icone = st.cor === 'VERDE' ? '🟢' : (st.cor === 'AMARELO' ? '🟡' : '🔴');
                badgesHtml += `<span class="area-pill ${st.cor}">${icone} ${setor}</span>`;
            }
        });

        card.innerHTML = `
            <div class="item-code-badge">ITEM ${item.codigo}</div>
            <div class="item-title-short">${item.titulo}</div>
            <div class="semaforo-row">${badgesHtml}</div>
        `;
        container.appendChild(card);
    });
}

function renderizarItemAtual() {
    if (!ITEM_ATUAL) return;
    document.getElementById('mainItemTitulo').textContent = ITEM_ATUAL.titulo;

    const checks = document.querySelectorAll('#areasDropdownMenu input[type="checkbox"]');
    checks.forEach(chk => {
        chk.checked = ITEM_ATUAL.areas_atribuidas.includes(chk.value);
    });

    if (DRAWER_ABERTO) {
        atualizarTextoDrawer();
    }

    populaSelectDemandas();
}

function toggleAreasDropdown() {
    document.getElementById('areasDropdownMenu').classList.toggle('show');
}

async function salvarAreasAtribuidas() {
    const checks = document.querySelectorAll('#areasDropdownMenu input[type="checkbox"]');
    const selecionadas = [];
    checks.forEach(c => { if (c.checked) selecionadas.push(c.value); });
    if (selecionadas.length === 0) {
        alert('Pelo menos uma coordenação deve estar atribuída ao item.');
        renderizarItemAtual();
        return;
    }
    await fetch(`/api/item/${encodeURIComponent(ITEM_ATUAL.codigo)}/areas`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({areas_atribuidas: selecionadas})
    });
    await carregarSistema();
}

function toggleDrawer(qual) {
    const drawer = document.getElementById('drawerContent');
    if (DRAWER_ABERTO === qual) {
        DRAWER_ABERTO = null;
        drawer.classList.remove('open');
    } else {
        DRAWER_ABERTO = qual;
        drawer.classList.add('open');
        atualizarTextoDrawer();
    }
    document.getElementById('btnDrawerApontamento').classList.toggle('active', DRAWER_ABERTO === 'apontamento');
    document.getElementById('btnDrawerMedidas').classList.toggle('active', DRAWER_ABERTO === 'medidas');
    document.getElementById('btnDrawerObs').classList.toggle('active', DRAWER_ABERTO === 'obs');
}

function atualizarTextoDrawer() {
    const drawer = document.getElementById('drawerContent');
    if (!ITEM_ATUAL || !DRAWER_ABERTO) return;
    if (DRAWER_ABERTO === 'apontamento') {
        drawer.textContent = ITEM_ATUAL.apontamento_ans;
    } else if (DRAWER_ABERTO === 'medidas') {
        drawer.textContent = ITEM_ATUAL.medidas_implementadas;
    } else if (DRAWER_ABERTO === 'obs') {
        drawer.textContent = ITEM_ATUAL.observacoes;
    }
}

function selecionarTipoDemanda(tipo) {
    TIPO_DEMANDA_ATUAL = tipo;
    document.getElementById('btnTipoFixa').classList.toggle('active', tipo === 'FIXA');
    document.getElementById('btnTipoAvulsa').classList.toggle('active', tipo === 'AVULSA');
    populaSelectDemandas();
}

function populaSelectDemandas() {
    const sel = document.getElementById('selectDemanda');
    sel.innerHTML = '';
    const filtradas = ITEM_ATUAL.demandas.filter(d => d.tipo === TIPO_DEMANDA_ATUAL);

    if (filtradas.length === 0) {
        const opt = document.createElement('option');
        opt.value = '';
        opt.textContent = TIPO_DEMANDA_ATUAL === 'AVULSA' ? '(Nenhuma demanda avulsa — clique em + Nova Demanda)' : '(Sem demandas)';
        sel.appendChild(opt);
        DEMANDA_ATUAL_ID = null;
        document.getElementById('textareaResumoDemanda').value = '';
        document.getElementById('listaChecklistContainer').innerHTML = '';
        return;
    }

    filtradas.forEach(d => {
        const opt = document.createElement('option');
        opt.value = d.id;
        opt.textContent = d.nome;
        if (d.nome === DEMANDA_ATUAL_NOME) opt.selected = true;
        sel.appendChild(opt);
    });

    DEMANDA_ATUAL_ID = parseInt(sel.value, 10);
    const demObj = filtradas.find(d => d.id === DEMANDA_ATUAL_ID);
    if (demObj) DEMANDA_ATUAL_NOME = demObj.nome;

    carregarChecklistDemanda();
}

async function aoMudarDemanda() {
    const sel = document.getElementById('selectDemanda');
    if (!sel.value) return;
    DEMANDA_ATUAL_ID = parseInt(sel.value, 10);
    const demObj = ITEM_ATUAL.demandas.find(d => d.id === DEMANDA_ATUAL_ID);
    if (demObj) DEMANDA_ATUAL_NOME = demObj.nome;
    await carregarSistema();
}

async function criarNovaDemandaAvulsa() {
    const nome = prompt('Digite o nome da nova Demanda Avulsa (ex: Ofício 108/2026 - Revisão Extra):');
    if (!nome || !nome.trim()) return;
    await fetch('/api/demandas', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
            item_codigo: ITEM_ATUAL.codigo,
            nome: nome.trim(),
            replicar_todos: true
        })
    });
    TIPO_DEMANDA_ATUAL = 'AVULSA';
    DEMANDA_ATUAL_NOME = nome.trim();
    document.getElementById('btnTipoFixa').classList.remove('active');
    document.getElementById('btnTipoAvulsa').classList.add('active');
    await carregarSistema();
}

function selecionarSetor(setor) {
    SETOR_ATUAL = setor;
    ['CODER', 'COCAP', 'CONEC'].forEach(s => {
        document.getElementById('btnSetor' + s).classList.toggle('active', s === setor);
    });
    renderizarChecklistNaTela();
}

async function salvarResumoDemanda(mostrarAlerta) {
    if (!DEMANDA_ATUAL_ID) return;
    const texto = document.getElementById('textareaResumoDemanda').value;
    const resp = await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/resumo`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ resumo_texto: texto })
    });
    if (resp.ok) {
        const lbl = document.getElementById('statusResumoSalvo');
        lbl.textContent = '✔ Resumo salvo!';
        setTimeout(() => { lbl.textContent = ''; }, 3000);
        if (mostrarAlerta) {
            alert('Resumo executivo salvo com sucesso! Ele sairá na Seção 1 do Ofício antes da análise das evidências.');
        }
    }
}

async function importarResumoPlanilhaBase() {
    if (!ITEM_ATUAL || !DEMANDA_ATUAL_ID) return;
    const histTxt = (ITEM_ATUAL.historico_planilha && ITEM_ATUAL.historico_planilha[DEMANDA_ATUAL_NOME]) || '';
    if (!histTxt || histTxt === 'N/A' || histTxt === 'Preencher') {
        alert('Não há texto histórico pré-cadastrado na planilha inicial para a competência "' + DEMANDA_ATUAL_NOME + '".');
        return;
    }
    const atual = document.getElementById('textareaResumoDemanda').value.trim();
    if (atual && !confirm('Substituir o texto atual da caixinha pelo texto integral da planilha para "' + DEMANDA_ATUAL_NOME + '"?')) {
        return;
    }
    document.getElementById('textareaResumoDemanda').value = histTxt;
    await salvarResumoDemanda(false);
}

async function carregarChecklistDemanda() {
    if (!DEMANDA_ATUAL_ID) return;
    const resp = await fetch(`/api/checklist/${DEMANDA_ATUAL_ID}`);
    const dados = await resp.json();
    CHECKLIST_ATUAL = dados.checklist || [];
    ITEM_ATUAL.semaforo = dados.semaforo;

    const resumoSalvo = (dados.demanda && dados.demanda.resumo_texto) ? dados.demanda.resumo_texto : '';
    document.getElementById('textareaResumoDemanda').value = resumoSalvo;

    renderizarSemaforoTopo(dados.semaforo);
    renderizarChecklistNaTela();
}

function renderizarSemaforoTopo(semaforo) {
    const box = document.getElementById('mainSemaforoBadges');
    box.innerHTML = '';
    ['CODER', 'COCAP', 'CONEC'].forEach(setor => {
        const st = semaforo[setor];
        const btnSetor = document.getElementById('btnSetor' + setor);
        if (!ITEM_ATUAL.areas_atribuidas.includes(setor)) {
            box.innerHTML += `<span class="area-pill INATIVO">${setor}: Não atribuído</span>`;
            btnSetor.innerHTML = `${setor}`;
        } else if (st) {
            const icone = st.cor === 'VERDE' ? '🟢' : (st.cor === 'AMARELO' ? '🟡' : '🔴');
            box.innerHTML += `<span class="area-pill ${st.cor}">${icone} ${setor}: ${st.label}</span>`;
            btnSetor.innerHTML = `${icone} ${setor} (${st.concluidos}/${st.total})`;
        }
    });
}

function renderizarChecklistNaTela() {
    document.getElementById('tituloChecklistSetor').textContent =
        `Checklist de Evidências — Coordenação ${SETOR_ATUAL} | Demanda: ${DEMANDA_ATUAL_NOME}`;

    const itensSetor = CHECKLIST_ATUAL.filter(c => c.setor === SETOR_ATUAL);
    const concluidosSetor = itensSetor.filter(c => c.concluido).length;
    const percSetor = itensSetor.length > 0 ? Math.round((concluidosSetor / itensSetor.length) * 100) : 0;

    document.getElementById('lblProgressoSetor').textContent =
        `Progresso ${SETOR_ATUAL}: ${concluidosSetor} de ${itensSetor.length} concluídos (${percSetor}%)`;
    document.getElementById('barraProgressoSetor').style.width = `${percSetor}%`;

    const totalGeral = CHECKLIST_ATUAL.length;
    const concGeral = CHECKLIST_ATUAL.filter(c => c.concluido).length;
    const percGeral = totalGeral > 0 ? Math.round((concGeral / totalGeral) * 100) : 0;
    document.getElementById('lblProgressoDemandaGeral').textContent =
        `Total da Demanda (3 Coordenações): ${concGeral}/${totalGeral} (${percGeral}%)`;

    const container = document.getElementById('listaChecklistContainer');
    container.innerHTML = '';

    if (itensSetor.length === 0) {
        container.innerHTML = `
            <div style="padding:24px; text-align:center; color:#64748B; background:#F8FAFC; border-radius:8px;">
                Nenhuma evidência cadastrada para <b>${SETOR_ATUAL}</b> em <b>${DEMANDA_ATUAL_NOME}</b>.<br>
                Use o campo acima para adicionar itens ao checklist.
            </div>`;
        return;
    }

    itensSetor.forEach((chk, idx) => {
        const card = document.createElement('div');
        card.className = 'chk-item-card ' + (chk.concluido ? 'concluido' : 'pendente');

        let anexosHtml = '';
        if (chk.anexos && chk.anexos.length > 0) {
            chk.anexos.forEach(a => {
                anexosHtml += `
                    <div class="anexo-chip">
                        <span>📄 ${a.nome_arquivo}</span>
                        <a href="/api/anexos/${a.id}/view" target="_blank" title="Visualizar Anexo">👁️ Ver</a>
                        <a href="/api/anexos/${a.id}/download" title="Baixar Anexo">⬇️ Baixar</a>
                        <button onclick="removerAnexo(${a.id})" style="border:none; background:none; color:#DC3545; font-weight:800; cursor:pointer;" title="Remover Anexo (volta para Pendente)">✖</button>
                    </div>
                `;
            });
        } else {
            anexosHtml = `<span style="font-size:11.5px; color:#64748B;">Nenhum anexo incluído ainda.</span>`;
        }

        card.innerHTML = `
            <div class="chk-top-line">
                <div class="chk-title-group">
                    <span style="font-weight:800; font-size:12px; color:#003366;">#${idx + 1}</span>
                    <span class="chk-title-text">${chk.titulo}</span>
                </div>
                <div class="chk-tools">
                    <button class="tool-btn" onclick="acionarUpload(${chk.id})">📎 Incluir Anexo(s)</button>
                    <button class="tool-btn" onclick="editarEvidencia(${chk.id})">✏️ Editar</button>
                    <button class="tool-btn danger" onclick="excluirEvidencia(${chk.id})">🗑️ Excluir</button>
                    <button class="tool-btn ${chk.concluido ? 'concluir-on' : ''}" onclick="alternarConclusao(${chk.id}, ${!chk.concluido}, ${chk.anexos.length})">
                        ${chk.concluido ? '✔ Concluído (Reabrir)' : '✔ Concluir'}
                    </button>
                </div>
            </div>
            <div class="anexos-box">${anexosHtml}</div>
        `;
        container.appendChild(card);
    });
}

async function adicionarEvidencia() {
    if (!DEMANDA_ATUAL_ID) {
        alert('Selecione ou crie uma demanda primeiro.');
        return;
    }
    const input = document.getElementById('inputNovaEvidencia');
    const titulo = input.value.trim();
    if (!titulo) return;

    await fetch('/api/checklist', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
            demanda_id: DEMANDA_ATUAL_ID,
            setor: SETOR_ATUAL,
            titulo: titulo
        })
    });
    input.value = '';
    await carregarChecklistDemanda();
    await carregarSistema();
}

async function editarEvidencia(chkId) {
    const chk = CHECKLIST_ATUAL.find(c => c.id === chkId);
    if (!chk) return;
    const novoTitulo = prompt('Editar descrição da evidência:', chk.titulo);
    if (!novoTitulo || !novoTitulo.trim()) return;
    const novaAnalise = prompt('Observação/Explicação de como obteve esta evidência (para sair no Ofício):', chk.analise_evidencia || '');

    await fetch(`/api/checklist/${chkId}`, {
        method: 'PUT',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
            titulo: novoTitulo.trim(),
            analise_evidencia: novaAnalise !== null ? novaAnalise : (chk.analise_evidencia || '')
        })
    });
    await carregarChecklistDemanda();
}

async function excluirEvidencia(chkId) {
    if (!confirm('Deseja realmente excluir este item do checklist e seus anexos?')) return;
    await fetch(`/api/checklist/${chkId}`, { method: 'DELETE' });
    await carregarChecklistDemanda();
    await carregarSistema();
}

async function alternarConclusao(chkId, novoStatus, qtdAnexos) {
    if (novoStatus && qtdAnexos === 0) {
        const ok = confirm('Atenção: Este item ainda não possui anexo importado. Deseja marcá-lo como CONCLUÍDO mesmo assim?');
        if (!ok) return;
    }
    await fetch(`/api/checklist/${chkId}`, {
        method: 'PUT',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ concluido: novoStatus })
    });
    await carregarChecklistDemanda();
    await carregarSistema();
}

function acionarUpload(chkId) {
    CHK_ID_PARA_UPLOAD = chkId;
    const inp = document.getElementById('hiddenFileInput');
    inp.value = '';
    inp.click();
}

async function executarUploadAnexos() {
    const inp = document.getElementById('hiddenFileInput');
    if (!inp.files || inp.files.length === 0 || !CHK_ID_PARA_UPLOAD) return;

    const formData = new FormData();
    for (let i = 0; i < inp.files.length; i++) {
        formData.append('arquivos', inp.files[i]);
    }

    await fetch(`/api/checklist/${CHK_ID_PARA_UPLOAD}/anexos`, {
        method: 'POST',
        body: formData
    });
    await carregarChecklistDemanda();
    await carregarSistema();
}

async function removerAnexo(anexoId) {
    if (!confirm('Remover este anexo? O status do item voltará automaticamente para PENDENTE.')) return;
    await fetch(`/api/anexos/${anexoId}`, { method: 'DELETE' });
    await carregarChecklistDemanda();
    await carregarSistema();
}

async function abrirModalOficio() {
    if (!DEMANDA_ATUAL_ID) {
        alert('Selecione uma demanda primeiro.');
        return;
    }
    const resumoAtual = document.getElementById('textareaResumoDemanda').value;
    const resp = await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/gerar-minuta-oficio`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
            regenerar: true,
            resumo_texto: resumoAtual
        })
    });
    const dados = await resp.json();
    preencherModalOficio(dados);
    document.getElementById('modalOficio').classList.add('open');
}

async function regenerarTextoOficio() {
    const resumoAtual = document.getElementById('textareaResumoDemanda').value;
    const resp = await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/gerar-minuta-oficio`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
            regenerar: true,
            resumo_texto: resumoAtual
        })
    });
    const dados = await resp.json();
    preencherModalOficio(dados);
}

function preencherModalOficio(dados) {
    document.getElementById('textareaOficio').value = dados.texto_oficio || '';
    const box = document.getElementById('resumoArquivosModal');
    const qtdPdf = (dados.arquivos_pdf || []).length;
    const qtdXls = (dados.arquivos_excel || []).length;

    let html = `<b>📦 Composição do Dossiê Unificado que será gerado após sua aprovação:</b><br>`;
    html += `• <b>1º) Ofício de Resposta (acima)</b> com o <b>Resumo Executivo (Seção 1)</b> seguido da <b>Análise das Evidências (Seção 2)</b>;<br>`;
    html += `• <b>2º) ${qtdPdf} anexo(s) PDF/Imagem</b> mesclados integralmente logo abaixo do Ofício em um único PDF;<br>`;
    if (qtdXls > 0) {
        const nomesXls = dados.arquivos_excel.map(x => x.nome_arquivo).join(', ');
        html += `• <b>3º) ${qtdXls} planilha(s) Excel/CSV</b> citadas expressamente pelo nome no corpo e no índice do Ofício: <code>${nomesXls}</code>.`;
    } else {
        html += `• Nenhuma planilha Excel/CSV externa pendente de citação nominal.`;
    }
    box.innerHTML = html;
}

function fecharModalOficio() {
    document.getElementById('modalOficio').classList.remove('open');
}

async function salvarRascunhoOficio() {
    const texto = document.getElementById('textareaOficio').value;
    await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/salvar-oficio`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ texto_oficio: texto })
    });
    alert('Texto do Ofício salvo com sucesso no banco de dados!');
}

async function baixarOficioEmWord() {
    const texto = document.getElementById('textareaOficio').value;
    const resp = await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/exportar-word`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ texto_oficio: texto })
    });
    if (!resp.ok) {
        alert('Erro ao gerar documento Word.');
        return;
    }
    const blob = await resp.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Oficio_PRASS_${ITEM_ATUAL.codigo.replace(/\./g, '_')}_${DEMANDA_ATUAL_NOME.replace(/[^a-zA-Z0-9]/g, '_')}.doc`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
}

async function aprovarEBaixarPDFUnificado() {
    const texto = document.getElementById('textareaOficio').value;
    const resp = await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/exportar-pdf-unificado`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ texto_oficio: texto })
    });
    if (!resp.ok) {
        alert('Erro ao gerar o PDF unificado.');
        return;
    }
    const blob = await resp.blob();
    const nomeSugerido = `Oficio_e_Evidencias_PRASS_${ITEM_ATUAL.codigo.replace(/\./g, '_')}_${DEMANDA_ATUAL_NOME.replace(/[^a-zA-Z0-9]/g, '_')}.pdf`;

    if (window.showSaveFilePicker) {
        try {
            const handle = await window.showSaveFilePicker({
                suggestedName: nomeSugerido,
                types: [{
                    description: 'Documento PDF Unificado',
                    accept: {'application/pdf': ['.pdf']}
                }]
            });
            const writable = await handle.createWritable();
            await writable.write(blob);
            await writable.close();
            fecharModalOficio();
            return;
        } catch (err) {
            if (err.name === 'AbortError') return;
        }
    }

    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = nomeSugerido;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
    fecharModalOficio();
}

window.addEventListener('DOMContentLoaded', carregarSistema);
</script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=True)
