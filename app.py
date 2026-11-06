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

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except Exception as e:
    HAS_OPENPYXL = False
    print(f"[AVISO BOOT] openpyxl indisponivel: {e}")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # Limite de 64MB para uploads


# ==============================================================================
# OPÇÕES PADRONIZADAS DE PREENCHIMENTO DA PLANILHA PRASS (MÓDULO PLANILHA)
# ==============================================================================
OPCOES_STATUS_PLANILHA = [
    "Finalizado / Concluído",
    "Em Andamento / Em Implementação",
    "Parcialmente Atendido (Em Complementação)",
    "Aguardando Área Técnica (GETEC / GECOF / GEREL)",
    "Preencher",
    "N/A (Não Aplicável)"
]

SUBITENS_SUGERIDOS_POR_SETOR = {
    "CODER": [
        "Vulnerabilidade de Rede (Painel de Análise de Vulnerabilidade)",
        "Monitoramento de Garantia de Atendimento e Reembolso Integral",
        "Garantia de Atendimento em virtude das suspensões contratuais (RN 566/2022)",
        "Revisão das Instruções Normativas IOT 222 e IOT 224"
    ],
    "COCAP": [
        "Bases de rede e controles de atualização (BENNER / Relatório POSTAL 988)",
        "Rede credenciada no portal corporativo: RN 486/2022, art. 2º",
        "Registros de movimentação e substituição de rede não hospitalar: RN 567/2022, art. 11",
        "Gerenciamento de inconsistências do Guia de Saúde (GLPI 30182) e Higienização (IOT 068)"
    ],
    "CONEC": [
        "Atualização de normativos: Manual MAN006 e Instrução Operacional IOT 036.07",
        "Suspensão de atendimento por inadimplência (Registro, Negociação e Quitação)",
        "Gestão da Rede Assistencial Direta e Indireta (Compliance e Contratos)",
        "Aplicação das IOTs 031 (Descredenciamento), 035 (Recontratualização) e 036 (Credenciamento)"
    ]
}

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

TXT_CODER_OUT_26 = (
    "GERED/CODER\n"
    "No que diz respeito às ações da GERED/CODER, registramos que o painel de análise de vulnerabilidade "
    "encontra-se em processo de revisão, abrangendo a metodologia de extração, tratamento e análise dos dados "
    "utilizados em sua construção. A iniciativa tem por objetivo aprimorar a qualidade das informações, "
    "fortalecer a rastreabilidade das bases utilizadas e assegurar maior aderência às necessidades institucionais "
    "e operacionais.\n"
    "Em decorrência das avaliações realizadas, verificou-se a necessidade de reestruturação do modelo atualmente "
    "adotado, com vistas à institucionalização da ferramenta e ao estabelecimento de critérios formais para sua "
    "manutenção, atualização e governança.\n"
    "Nesse contexto, foi formalizada, em 08/10/2026, demanda junto à área de Tecnologia da Informação para "
    "desenvolvimento e institucionalização do painel de vulnerabilidade, conforme evidência constante do Anexo I."
)

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
            "situações relacionadas à insuficiência de rede prestadora. A análise da Agência também posicionou a Postal "
            "Saúde na 19ª colocação no indicador de Falhas de Reembolso.\n\n"
            "Diagnóstico institucional\n"
            "A Postal Saúde dispõe de fluxo formal para o processamento das solicitações de reembolso (POP009 – Prover "
            "Reembolso Assistencial – Versão 4), estabelecendo prazo de até 30 dias corridos contado do recebimento da "
            "documentação completa. No acompanhamento operacional, foram identificados protocolos pendentes além do prazo "
            "de referência, monitorados em conjunto com a área financeira. Identificaram-se também situações em que "
            "prestadores solicitam indevidamente pagamento direto ao beneficiário, motivando a instauração do Comitê de "
            "Procedimento Ético, campanha educativa e Código de Conduta Ética para Prestadores Assistenciais."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• Saneamento das pendências identificadas no processo de reembolso, com acompanhamento dos protocolos e "
            "atuação direcionada sobre os casos que ultrapassaram o prazo estabelecido.\n"
            "• Aperfeiçoamento integrado com a área Financeira, mediante revisão dos fluxos e interfaces de processamento e pagamento.\n"
            "• Desenvolvimento de medidas de automação junto à TI para reduzir intervenções manuais e ampliar a rastreabilidade.\n"
            "• Implementação do Comitê de Procedimento Ético, campanha educativa e Código de Conduta Ética para Prestadores.\n"
            "• Tratamento das pendências identificadas com previsão de saneamento até outubro de 2026."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "O tratamento da matéria é acompanhado por indicadores capazes de demonstrar a redução do passivo e a evolução do processo:\n"
            "• estoque de solicitações pendentes;\n"
            "• quantidade e percentual de processos fora do prazo;\n"
            "• percentual de solicitações processadas dentro do prazo;\n"
            "• tempo médio de processamento;\n"
            "• pendências documentais, negativas, reanálises e divergências de valores;\n"
            "• tempo entre processamento e pagamento;\n"
            "• NIPs relacionadas a reembolso e evolução do indicador de Falhas de Reembolso."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED\n"
                "• Monitoramento dos processos de Garantia de Atendimento e Reembolso: GERED/CODER é responsável pelo monitoramento // Execução: GEREL - UAR/NAR;\n"
                "• Processamento e pagamento: Responsabilidade da GEFAC/COANF e GECOF/COPAG;\n"
                "• Recursos Tecnológicos: Em andamento solicitação para obrigatoriedade do protocolo de atendimento no APP ao solicitar Reembolso Integral;\n"
                "• Rede Assistencial: Sob gestão da CONEC em revisão normativa.\n"
                "Indicadores Ago/26:\n"
                "- Garantia de Atendimento: Dentro do Prazo: 236 | Fora do Prazo: 1.631 | Total: 1.867 (+27,70%)\n"
                "- Reembolso: Dentro do Prazo: 361 | Fora do Prazo: 715 | Total: 1.076 (+2,09%)"
            ),
            "Status 09/2026": (
                "GERED\n"
                "Identificada redução dos processos de garantia e reembolso no período de agosto a setembro/2026:\n"
                "- Garantia de Atendimento: Ago/26 (1.867) -> Set/26: Dentro do Prazo: 200 | Fora do Prazo: 1.370 | Total: 1.570 (-15,91% de redução)\n"
                "- Reembolso: Ago/26 (1.076) -> Set/26: Dentro do Prazo: 444 | Fora do Prazo: 442 | Total: 886 (-17,66% de redução)\n"
                "Decorrido o período de férias e como reflexo do monitoramento extensivo da GERED/CODER, observa-se redução consistente dos protocolos fora do prazo."
            ),
            "Status 10/2026": (
                "GERED/CODER\n"
                "Implementada em 01/10/2026 a obrigatoriedade de o beneficiário informar o número do protocolo de atendimento quando solicitar Reembolso Integral no aplicativo, garantindo rastreabilidade dos processos de Garantia de Atendimento liberados como Reembolso Integral."
            ),
            "Complemento ANS 26/10": "N/A"
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
            "A Postal Saúde possui processo formal de Garantia de Atendimento destinado ao tratamento de situações de "
            "inexistência ou indisponibilidade de rede assistencial, determinando o esgotamento ordenado das alternativas:\n"
            "1. rede credenciada;\n"
            "2. contratação direta;\n"
            "3. Tratamento Fora do Domicílio (TFD);\n"
            "4. reembolso integral.\n"
            "As negociações de contratação direta são formalizadas e anexadas ao sistema para fins de evidência."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• Processo estruturado para gestão da Garantia de Atendimento, com atribuição formal da área estratégica de rede "
            "para identificar vulnerabilidades e estabelecer Contratação Direta, TFD e Reembolso Assistencial.\n"
            "• Busca por atendimento considerando rede direta, rede indireta, parcerias e Telemedicina, seguida de "
            "contratação direta, TFD e reembolso integral.\n"
            "• Formalização obrigatória das negociações de contratação direta e manutenção das evidências no sistema."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "A operacionalização da Garantia de Atendimento prevê: acolhimento da solicitação, análise da cobertura, "
            "verificação da rede credenciada, comprovação sistêmica das tentativas de agendamento, definição da alternativa "
            "assistencial, autorização de reembolso integral (quando aplicável) e comunicação ao beneficiário nos prazos da ANS."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED\n"
                "Normativos vigentes da Garantia de Atendimento:\n"
                "• MAN 037 (v.3 – 11/08/2026) | POP 009 (v.4 – 13/08/2026) | POP 012 (v.1 – 13/08/2026) | POP 013 (v.0 – 13/08/2026) | IOT 222 | IOT 224.\n"
                "As instruções normativas IOT 222 e IOT 224 e o painel de análise de vulnerabilidade encontram-se em revisão metodológica."
            ),
            "Status 09/2026": (
                "GERED\n"
                "Continuidade da revisão das instruções normativas IOT 222 e IOT 224, mapeamento de prestadores ativos junto às UAR/NAR e revisão metodológica do painel de análise de vulnerabilidade."
            ),
            "Status 10/2026": TXT_CODER_OUT_26,
            "Complemento ANS 26/10": f"Subitem: Vulnerabilidade de Rede\n{TXT_CODER_OUT_26}"
        }
    },
    {
        "codigo": "8.4.1",
        "titulo": "8.4.1 Insuficiência/inexistência de rede prestadora para a realização de consultas e procedimentos de cobertura obrigatória, acarretando desassistência ou pagamento particular para posterior reembolso",
        "areas_atribuidas": ["CODER", "CONEC"],
        "apontamento_ans": (
            "Apontamento da ANS\n"
            "A Nota Técnica nº 39 aponta insuficiência ou inexistência de rede prestadora para realização de consultas e "
            "procedimentos de cobertura obrigatória, acarretando desassistência ou pagamento particular para posterior reembolso.\n\n"
            "Diagnóstico institucional\n"
            "A Operadora dispõe de processo formal de Garantia de Atendimento que contempla acolhimento da demanda, análise "
            "de cobertura, verificação da rede disponível, comprovação das tentativas de agendamento e definição da "
            "alternativa assistencial (contratação direta, TFD ou reembolso integral). O tratamento do reembolso vincula-se ao item 8.1.2."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• Implementação do Compartilhamento de Gestão de Riscos e Rede Assistencial para ampliar a disponibilidade e capilaridade da rede.\n"
            "• Manutenção dos mecanismos previstos no processo de Garantia de Atendimento para situações de inexistência ou indisponibilidade.\n"
            "• Condução das medidas de saneamento e aperfeiçoamento do reembolso conforme item 8.1.2."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "O tratamento do risco ocorre em duas dimensões complementares:\n"
            "I – Gestão da rede: identificação de insuficiências e vulnerabilidades, recomposição e ampliação da rede;\n"
            "II – Garantia individual de acesso: adoção das alternativas assistenciais previstas quando identificada inexistência ou indisponibilidade."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED\n"
                "1) Garantia de atendimento e reembolso: conduzidos conforme item 8.1.2;\n"
                "2) Gestão da Rede Assistencial (Direta/Indireta): sob gestão da CONEC em revisão de compliance;\n"
                "3) Compartilhamento de Gestão de Riscos (Encapsulamento): sob responsabilidade da GEREL."
            ),
            "Status 09/2026": (
                "GERED\n"
                "Ratificado o acompanhamento integrado ao item 8.1.2 (CODER), revisão normativa da rede direta/indireta (CONEC) e gestão do compartilhamento de riscos (GEREL)."
            ),
            "Status 10/2026": (
                "GERED/CODER\n"
                "Implementada em 01/10/2026 a obrigatoriedade do protocolo de atendimento na solicitação de Reembolso Integral no APP, acompanhada da redução de -15,91% nos processos de Garantia de Atendimento e -17,66% nos processos de Reembolso."
            ),
            "Complemento ANS 26/10": "N/A"
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
            "A Postal Saúde reconhece que o período analisado foi marcado por restrição financeira decorrente da "
            "irregularidade dos repasses da mantenedora, impactando obrigações financeiras perante parte da rede prestadora "
            "e exigindo medidas estruturantes para reduzir a vulnerabilidade assistencial."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• Implementação do Compartilhamento de Gestão de Riscos e Rede Assistencial utilizando redes de operadoras parceiras.\n"
            "• Monitoramento e negociação direta com prestadores suspensos e direcionamento assistencial via Garantia de Atendimento."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "O risco é tratado de forma integrada entre as dimensões financeira, assistencial, regulatória e de gestão da rede, "
            "monitorando a suficiência da rede, continuidade da assistência, indicadores assistenciais, reclamações e NIPs."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED\n"
                "• Aplicação das IOTs 031 (descredenciamento), 036 (credenciamento) e 068 (higienização) e desenvolvimento de dashboard de controle;\n"
                "• Estudos de suficiência e capilaridade contemplados na revisão do painel de vulnerabilidade;\n"
                "• Monitoramento assistencial conduzido pela CODER conforme item 8.1.2;\n"
                "• Parcerias de compartilhamento de risco firmadas até ago/2026: AMIL RJ, UNIMED CNU e UNIMED CENTRO OESTE PAULISTA."
            ),
            "Status 09/2026": (
                "GERED\n"
                "Manutenção do monitoramento das suspensões, aplicação das IOTs 031, 036 e 068, evolução do painel de vulnerabilidade e vigência dos convênios de compartilhamento de risco."
            ),
            "Status 10/2026": TXT_CODER_OUT_26,
            "Complemento ANS 26/10": (
                "GERED/CODER — Subitem: Garantia de Atendimento em virtude das suspensões de atendimento\n"
                "Adotadas medidas imediatas para assegurar a continuidade da assistência aos beneficiários e mitigar os impactos das suspensões contratuais, mediante direcionamento do fluxo assistencial para prestadores aptos da rede credenciada, contratação direta ou reembolso integral nos prazos da RN nº 566/2022 da ANS.\n\n"
                "GERED/CONEC — Subitem: Suspensão de atendimento por inadimplência\n"
                "Em atendimento ao Ofício nº 108/2026 da ANS, apresentados os procedimentos adotados para: 1. Registro e controle das suspensões de atendimento; 2. Monitoramento gerencial das ocorrências; 3. Atuação junto aos prestadores e formalização de acordos de quitação; 4. Evidências documentais."
            )
        }
    },
    {
        "codigo": "8.5",
        "titulo": "8.5. NÃO DISPONIBILIZAÇÃO DE FORMA CLARA DA REDE AOS BENEFICIÁRIOS NOS TERMOS DA REGULAMENTAÇÃO EM VIGOR",
        "areas_atribuidas": ["CODER", "COCAP", "CONEC"],
        "apontamento_ans": (
            "Apontamento da ANS\n"
            "A ANS apontou situação relacionada à não disponibilização clara da rede assistencial aos beneficiários, "
            "abrangendo atualização das informações da rede, divulgação no portal corporativo e comunicação das substituições.\n\n"
            "Diagnóstico institucional\n"
            "A Postal Saúde possui estrutura formal de governança da rede sob responsabilidade da GERED. Entre janeiro/2025 "
            "e julho/2026, foram incorporados 305 novos prestadores (alcançando 13.086 na rede direta e 27.100 na rede "
            "indireta CNU, totalizando 40.186 prestadores) e registrados 580 descredenciamentos."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• Estrutura centralizada de gestão da rede assistencial pela GERED.\n"
            "• Utilização do Painel de Vulnerabilidade da Rede Prestadora.\n"
            "• Monitoramento contínuo dos credenciamentos, descredenciamentos, reativações e suspensões.\n"
            "• Aplicação das IOTs 031, 035, 036, 068 e 075 e fortalecimento dos controles sistêmicos."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "O apontamento é tratado no processo permanente de governança e melhoria contínua da rede assistencial, "
            "assegurando consistência entre a situação efetiva da rede, as bases corporativas e as informações divulgadas aos beneficiários."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED\n"
                "• Ações CODER: Painel de Vulnerabilidade da Rede Prestadora;\n"
                "• Ações CONEC: Revisão normativa de credenciamento, descredenciamento e suspensão;\n"
                "• Ações COCAP: Disponibilização da rede em tempo real no portal, funcionalidade de reporte de inconsistências e consulta de substitutos.\n"
                "Rede Assistencial (Ago/26): Rede Direta: 12.399 | Rede Indireta: 30.194 | Total: 42.593 prestadores."
            ),
            "Status 09/2026": (
                "GERED\n"
                "Continuidade das ações de revisão do Painel de Vulnerabilidade (CODER), revisão normativa (CONEC) e atualização contínua do Guia de Saúde (COCAP).\n"
                "Rede Assistencial (Set/26): Rede Direta: 12.357 | Rede Indireta: 30.268 | Total: 42.625 prestadores."
            ),
            "Status 10/2026": TXT_CODER_OUT_26,
            "Complemento ANS 26/10": f"Subitem: Vulnerabilidade de Rede\n{TXT_CODER_OUT_26}"
        }
    },
    {
        "codigo": "8.5.1",
        "titulo": "8.5.1. RELATOS DE DESATUALIZAÇÃO DA REDE EFETIVAMENTE DISPONIBILIZADA PELA OPERADORA",
        "areas_atribuidas": ["CODER", "COCAP", "CONEC"],
        "apontamento_ans": (
            "Apontamento da ANS\n"
            "Foram identificados relatos relacionados à desatualização da rede efetivamente disponibilizada pela Operadora.\n\n"
            "Diagnóstico institucional\n"
            "No período de janeiro/2025 a julho/2026, foram registrados 305 novos credenciamentos e 580 descredenciamentos, "
            "evidenciando a intensidade da movimentação da rede e a necessidade de mecanismos permanentes de higienização e "
            "atualização cadastral tempestiva."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• Monitoramento sistemático das movimentações da rede e utilização do Painel de Vulnerabilidade.\n"
            "• Aplicação da IOT 068 (higienização da rede), IOT 031 (descredenciamento), IOT 075 (suspensão de atendimento), "
            "IOT 158 e IOT 224.\n"
            "• Aprimoramento dos controles sistêmicos e automatizados voltados à rastreabilidade das movimentações."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "As movimentações da rede são acompanhadas pela GERED e submetidas aos fluxos institucionais correspondentes, "
            "com foco na redução de divergências entre a situação operacional da rede e as informações disponibilizadas aos beneficiários."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED/CODER: Revisão do painel de análise de vulnerabilidade.\n"
                "COCAP: Análises diárias focadas na qualificação contínua do cadastro, construção de painel de controle e aplicação estrita da IOT 075 nas hipóteses da RN nº 585 da ANS.\n"
                "CONEC: Revisão dos normativos da área."
            ),
            "Status 09/2026": (
                "GERED/CODER, COCAP e CONEC: Manutenção das rotinas diárias de qualificação cadastral, desenvolvimento do painel de controle e revisão normativa."
            ),
            "Status 10/2026": TXT_CODER_OUT_26,
            "Complemento ANS 26/10": (
                f"Subitem: Vulnerabilidade de Rede\n{TXT_CODER_OUT_26}\n\n"
                "GERED/CONEC — Subitem: Atualização de normativos (MAN006 e IOT036)\n"
                "O Manual de Credenciamento MAN006 encontra-se em revisão na plataforma DocSpider. A IOT 036.07 teve última atualização em 14/08/2026, com nova revisão prevista para 01/12/2027.\n\n"
                "GERED/COCAP — Subitem: Bases de rede e controles de atualização\n"
                "A Postal Saúde vem desenvolvendo ações voltadas ao aprimoramento dos processos de extração, atualização, monitoramento e higienização das informações cadastrais da rede assistencial, em atendimento ao Ofício nº 108/2026 da ANS. Foram disponibilizadas bases cadastrais (.CSV), formalizado o chamado GLPI nº 31043 para relatório de movimentações cadastrais com data e responsável, extraído o relatório POSTAL 988 no sistema BENNER para higienização da rede (IOT 068) e apresentadas as IOTs 031, 068 e 075."
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
            "corporativo da Operadora, nos termos da RN ANS nº 486/2022, art. 2º.\n\n"
            "Diagnóstico institucional\n"
            "Considerando a elevada dinâmica de movimentação da rede, o principal ponto de atenção consiste em assegurar "
            "que as informações constantes das bases operacionais estejam refletidas tempestivamente nos canais de "
            "disponibilização ao beneficiário."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• Gestão centralizada das movimentações da rede pela GERED.\n"
            "• Aprimoramento dos controles sistêmicos e automatizados (chamados GLPI nº 30182 e nº 30999).\n"
            "• Fortalecimento da rastreabilidade das alterações cadastrais e contratuais no Guia de Saúde."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "A consistência das informações da rede permanece submetida a ciclo contínuo de monitoramento e aperfeiçoamento "
            "para redução de eventuais divergências entre as bases internas e o portal corporativo."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED/CODER: Revisão do painel de vulnerabilidade.\n"
                "GERED/COCAP: Desenvolvimento de painel de controle gerencial (dashboard) para sistematizar a gestão centralizada das movimentações da rede e rastreabilidade das alterações cadastrais."
            ),
            "Status 09/2026": (
                "GERED/CODER e GERED/COCAP: Continuidade da revisão do painel de vulnerabilidade e evolução do módulo de gerenciamento de inconsistências do portal."
            ),
            "Status 10/2026": TXT_CODER_OUT_26,
            "Complemento ANS 26/10": (
                f"Subitem: Vulnerabilidade de Rede\n{TXT_CODER_OUT_26}\n\n"
                "GERED/COCAP — Subitem: Rede credenciada no portal: RN 486/2022, art. 2º\n"
                "A Postal Saúde vem adotando medidas para aprimorar os mecanismos de controle, atualização e transparência das informações disponibilizadas no Guia de Saúde, em atendimento aos requisitos da RN ANS nº 486/2022. Foram formalizados os chamados GLPI nº 30182 (1ª entrega do sistema de gerenciamento de reportes disponibilizada em homologação) e GLPI nº 30999 (adequação integral dos requisitos do art. 2º da RN 486/2022 no portal), acompanhados do painel gerencial, prints datados com URL e Checklist de Auditoria Documental do art. 2º."
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
            "de atenção à saúde não hospitalares, nos termos da RN ANS nº 567/2022, art. 11.\n\n"
            "Diagnóstico institucional\n"
            "A Postal Saúde possui processo formal de gestão das movimentações da rede, contemplando credenciamento, "
            "descredenciamento, recontratualização, substituição, suspensão e reorganização, apoiado por análise de vulnerabilidade."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• Aplicação das IOTs 031, 035, 036 e 068.\n"
            "• Avaliação da necessidade de substituição ou expansão da rede (CTI DIOPE/GERED 0020/2025).\n"
            "• Fortalecimento da rastreabilidade das movimentações e adequação do portal (GLPI nº 30992 e GLPI nº 31012)."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "Os processos de descredenciamento e substituição são acompanhados pela GERED e avaliados quanto aos impactos "
            "sobre a continuidade da assistência e adequada disponibilização das informações aos beneficiários no portal."
        ),
        "historico_planilha": {
            "Status 08/2026": (
                "GERED/CODER: Revisão do painel de vulnerabilidade.\n"
                "COCAP: Desenvolvimento de dashboard gerencial para sistematizar a aplicação das IOTs 031, 036 e 068 e rastreabilidade das substituições.\n"
                "CONEC: Revisão normativa contínua."
            ),
            "Status 09/2026": (
                "GERED/CODER, COCAP e CONEC: Evolução do painel de controle de movimentações e substituições da rede e revisão normativa."
            ),
            "Status 10/2026": TXT_CODER_OUT_26,
            "Complemento ANS 26/10": (
                f"Subitem: Vulnerabilidade de Rede\n{TXT_CODER_OUT_26}\n\n"
                "GERED/CONEC — Subitem: Atualização de normativos (MAN006 e IOT036)\n"
                "Manual MAN006 em revisão na plataforma DocSpider e IOT 036.07 atualizada em 14/08/2026.\n\n"
                "GERED/COCAP — Subitem: Registros de movimentação e substituição da rede / Substituição de rede não hospitalar: RN 567/2022, art. 11\n"
                "A resposta da Postal Saúde ao Ofício nº 108/2026 da ANS consolida as iniciativas para adequar os processos de movimentação e substituição da rede assistencial às exigências regulatórias, em especial ao Art. 11 da RN nº 567/2022. Foram formalizados os chamados GLPI nº 30992 (relatório consolidado de substitutos no Benner), GLPI nº 31012 (pesquisa isolada por prestador no portal) e GLPI nº 30182 (saneamento de inconsistências), além da apresentação das IOTs 031 e 036, prints comparativos Benner x Guia de Saúde, ferramenta de equivalência, CTI DIOPE/GERED 0020/2025 e Checklist de Conformidade do art. 11."
            )
        }
    },
    {
        "codigo": "8.6.1",
        "titulo": "8.6.1. Insuficiência de rede prestadora que garanta o cumprimento dos prazos máximos de atendimento, conforme a RN ANS nº 566/2022",
        "areas_atribuidas": ["CODER", "CONEC"],
        "apontamento_ans": (
            "Apontamento da ANS\n"
            "A ANS apontou indícios de insuficiência de rede prestadora para garantir o cumprimento dos prazos máximos de "
            "atendimento estabelecidos pela RN ANS nº 566/2022.\n\n"
            "Diagnóstico Institucional\n"
            "A Postal Saúde fortaleceu o acompanhamento da vulnerabilidade da rede considerando sua capacidade assistencial "
            "efetiva, disponibilidade por especialidade e município e integração entre o monitoramento da Garantia de "
            "Atendimento e a gestão da rede credenciada."
        ),
        "medidas_implementadas": (
            "Medidas Implementadas e em Implementação\n\n"
            "• Monitoramento diário da Garantia de Atendimento mediante painel gerencial em BI.\n"
            "• Monitoramento da vulnerabilidade da rede credenciada e emissão de relatórios analíticos às Regionais.\n"
            "• Utilização de Contratação Direta, Reembolso Integral e Tratamento Fora do Domicílio.\n"
            "• Padronização dos fluxos operacionais por meio do Manual da Garantia de Atendimento e migração para o E-doc."
        ),
        "observacoes": (
            "Tratamento e Monitoramento\n\n"
            "O tratamento do risco de insuficiência de rede é conduzido de forma preventiva e contínua, integrando "
            "monitoramento diário da demanda, análise da capacidade assistencial da rede e atuação junto às Regionais."
        ),
        "historico_planilha": {
            "Status 08/2026": "GERED: As informações deste item estão contempladas no item 8.1.2 no que se refere ao monitoramento diário das UAR/NAR com maior incidência em processos de GA fora do prazo.",
            "Status 09/2026": "GERED: Monitoramento diário das UAR/NAR conduzido de forma integrada ao item 8.1.2, registrando redução de -15,91% nos processos de Garantia de Atendimento em setembro/2026.",
            "Status 10/2026": "GERED/CODER: Implementada em 01/10/2026 a obrigatoriedade do protocolo de atendimento no aplicativo ao solicitar Reembolso Integral oriundo de Garantia de Atendimento.",
            "Complemento ANS 26/10": "N/A"
        }
    }
]

CHECKLISTS_INICIAIS_COMPLEMENTO = {
    "8.4.2": {
        "CODER": [
            {
                "titulo": "Demandas de Garantia de Atendimento geradas por essas suspensões e sua solução",
                "analise": "Adotadas medidas imediatas pela GERED/CODER para assegurar a continuidade da assistência aos beneficiários impactados por suspensões contratuais, mediante direcionamento assistencial para prestadores aptos da rede direta/indireta, contratação direta ou liberação de reembolso integral nos prazos da RN nº 566/2022."
            },
            {
                "titulo": "Controle mensal de prestadores suspensos e restabelecidos (indicadores do PRASS)",
                "analise": "Monitoramento mensal consolidado pela GERED acerca da evolução das suspensões e restabelecimentos de atendimento na rede assistencial."
            }
        ],
        "CONEC": [
            {
                "titulo": "Relação dos prestadores que suspenderam ou ameaçaram suspender por inadimplência: prestador, UF, data, valor em aberto e situação",
                "analise": "Levantamento estruturado contendo a identificação dos prestadores com notificação ou registro de suspensão por inadimplência, UF, data da ocorrência e status atual das tratativas."
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
                "analise": "Disponibilizadas extrações cadastrais completas da rede de prestadores extraídas do sistema BENNER (.CSV com 12.360 e 12.403 registros), contemplando CPF/CNPJ, razão social, categoria, tipo de prestador, datas de credenciamento/inclusão, endereço com IBGE/UF/CEP e especialidades."
            },
            {
                "titulo": "Relatório de alterações cadastrais do mês: inclusões, alterações e exclusões, com data e responsável",
                "analise": "Formalizada solicitação de extração junto à área técnica responsável por meio do chamado GLPI nº 31043 (COIDE > Dados Gerenciais), abrangendo os registros de inclusão, alteração e exclusão de prestadores realizados em agosto e setembro de 2026, com data e usuário responsável."
            },
            {
                "titulo": "Registros da rotina da GERED: checklist e amostra de validações",
                "analise": "Formalizado documento técnico da GERED/COCAP registrando o mapeamento e incorporação de checklist padronizado e metodologia de amostragem de validações às rotinas diárias de qualificação cadastral."
            },
            {
                "titulo": "Relatório de higienização da rede (IOT 068) do período",
                "analise": "Apresentado o procedimento de extração do relatório POSTAL 988 (Comprovante de Atendimento do Prestador no Período de 12 Meses) no sistema BENNER, em cumprimento à IOT 068."
            },
            {
                "titulo": "IOT 031, 068, 075, 158 e 224 revisadas ou cronograma formal de revisão",
                "analise": "Apresentadas na íntegra as Instruções Operacionais de Trabalho vigentes: IOT 075.02 (Suspensão de Atendimento), IOT 031.07 (Descredenciamento de Prestadores) e IOT 068.04 (Higienização da Rede Credenciada), acompanhadas do cronograma de revisão normativa."
            }
        ],
        "CODER": [
            {
                "titulo": "Revisão metodológica e reestruturação do Painel de Vulnerabilidade da Rede Prestadora",
                "analise": "Formalizada em 08/10/2026 demanda junto à TI para desenvolvimento e institucionalização do painel de vulnerabilidade, fortalecendo a rastreabilidade das bases utilizadas."
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
                "analise": "Formalizada demanda tecnológica sob o chamado GLPI nº 30182 (COSIS > Projetos). A 1ª Entrega (Registro e Atendimento de Reportes do Guia Saúde) foi disponibilizada em homologação em 21/10/2026, com geração de protocolo, e-mail automático com prazo de 5 dias úteis e painel administrativo."
            },
            {
                "titulo": "Cronograma do novo painel de controle e prints da versão atual",
                "analise": "Apresentado registro visual da versão atual do painel gerencial utilizado pela GERED/COCAP para acompanhamento de atividades operacionais, demandas, prazos e indicadores da rede assistencial."
            },
            {
                "titulo": "Bloco C1 (RN 486/2022, art. 2º) - Cronograma de Implantação",
                "analise": "Formalizada solicitação de adequação do Guia de Saúde aos requisitos do art. 2º da RN ANS nº 486/2022 por meio do chamado GLPI nº 30999 (COSIS > Minha Postal)."
            },
            {
                "titulo": "Bloco C1 (RN 486/2022, art. 2º) - Demanda de TI aberta (número e data) e especificação aprovada",
                "analise": "Comprovada a abertura do chamado GLPI nº 30999 em 26/10/2026 com a especificação detalhada de todos os requisitos do art. 2º da RN ANS nº 486/2022."
            },
            {
                "titulo": "Bloco C1 (RN 486/2022, art. 2º) - Homologação e data de entrada em produção",
                "analise": "Apresentada a especificação técnica registrada no chamado GLPI nº 30999 para orientar o desenvolvimento, homologação e entrada em produção das melhorias no portal."
            },
            {
                "titulo": "Bloco C1 (RN 486/2022, art. 2º) - Prints datados com URL, antes e depois, em ao menos três tipos de prestador",
                "analise": "Apresentados registros visuais datados com URL visível do Guia de Saúde (guiasaude.postalsaude.com.br), demonstrando a exibição cadastral de diferentes estabelecimentos e o mecanismo de reporte de inconsistências."
            },
            {
                "titulo": "Bloco C1 (RN 486/2022, art. 2º) - Checklist de conformidade do art. 2º assinado pela área",
                "analise": "Elaborado Checklist de Auditoria Documental do Art. 2º da RN 486/2022 (5 páginas), avaliando item a item com telas comprobatórias e vinculação ao chamado GLPI nº 30999."
            }
        ]
    },
    "8.5.3": {
        "COCAP": [
            {
                "titulo": "Bloco B6 - Relação do período: prestador, CNPJ, data, motivo, substituto e data de publicação no portal",
                "analise": "Formalizada solicitação de extração consolidada no sistema BENNER por meio do chamado GLPI nº 30992 (aberto em 26/10/2026 na fila COSIS > Benner - Saúde > Relatórios), com envio da planilha base de descredenciados de agosto e setembro/2026."
            },
            {
                "titulo": "Bloco B6 - Amostra de processos completos (IOT 031 e 036): solicitação, parecer, aprovação e aditivo ou distrato",
                "analise": "Apresentado compilado normativo de 59 páginas contendo a íntegra da IOT 031.07 (Descredenciamento de Prestadores) e da IOT 036.07 (Credenciamento de Prestador), incluindo os formulários FOP 056 e FOP 075."
            },
            {
                "titulo": "Bloco B6 - Prova da comunicação ao beneficiário: print datado e data de publicação",
                "analise": "Apresentada evidência comparativa contendo a captura da consulta pública de substituição no Guia de Saúde e a tela cadastral no sistema BENNER comprovando o registro do período de publicação no portal."
            },
            {
                "titulo": "Bloco B6 - Indicador do PRASS: percentual de substituições comunicadas no prazo",
                "analise": "Vinculado ao levantamento primário solicitado no chamado GLPI nº 30992 junto à equipe do sistema BENNER para apuração auditável das substituições comunicadas no prazo regulatório."
            },
            {
                "titulo": "Bloco B6 - Análise de equivalência do substituto (mesmo serviço e município)",
                "analise": "Apresentado registro visual da ferramenta técnica de verificação de equivalência assistencial entre prestador excluído e substituto (conferência de município, porte e serviços contratados)."
            },
            {
                "titulo": "Bloco B6 - Termos aditivos de habilitação e CTI DIOPE/GERED 0020/2025, prometidos no PRASS",
                "analise": "Apresentada a íntegra da Comunicação Interna CTI DIOPE/GERED nº 0020/2025 (NUP 00000.020940/2025-81), assinada digitalmente por João Victor Peres Moura, determinando às regionais o mapeamento de hospitais substitutos."
            },
            {
                "titulo": "Bloco C2 (RN 567/2022, art. 11) - Cronograma de Implantação",
                "analise": "Formalizada demanda evolutiva junto à TI por meio do chamado GLPI nº 31012 (COSIS > Intranet / Site), solicitando a implementação da pesquisa isolada por prestador na página de substituições."
            },
            {
                "titulo": "Bloco C2 (RN 567/2022, art. 11) - Demanda de TI, especificação e homologação",
                "analise": "Comprovado o registro do chamado GLPI nº 31012 contendo a especificação funcional detalhada para adequação da busca isolada e combinada de substituições no portal conforme o art. 11, inciso II, da RN 567/2022."
            },
            {
                "titulo": "Bloco C2 (RN 567/2022, art. 11) - Prints datados com URL, antes e depois, de ao menos três substituições",
                "analise": "Apresentado caderno comprobatório da CODER/GERED (5 páginas) com capturas das telas de substituição para 3 prestadores distintos (CR Serviços de Saúde Ltda - SC; Hospital da Criança Urgil - AC; e Ortoclínica de Sergipe Ltda - SE)."
            },
            {
                "titulo": "Bloco C2 (RN 567/2022, art. 11) - Saneamento dos cadastros com erro (datas e telefones) e relatório da correção",
                "analise": "Apresentado o fluxo estruturado de tratamento e saneamento de inconsistências cadastrais do Guia de Saúde disponibilizado em homologação por meio do chamado GLPI nº 30182, aliado à conferência no sistema BENNER."
            },
            {
                "titulo": "Bloco C2 (RN 567/2022, art. 11) - Checklist de conformidade do art. 11 assinado",
                "analise": "Elaborado documento de verificação de conformidade do art. 11 da RN 567/2022 (2 páginas), demonstrando o atendimento às alíneas 'a', 'b', 'c' e 'd' do inciso I e o acompanhamento da busca isolada do inciso II via GLPI nº 31012."
            }
        ]
    }
}


# ==============================================================================
# CAMADA DE BANCO DE DADOS COM AUTO-FALLBACK E MIGRAÇÃO AUTOMÁTICA
# ==============================================================================
def obter_url_postgres():
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


def inferir_status_inicial(nome_dem, resumo_inicial):
    if resumo_inicial == "N/A":
        return "N/A (Não Aplicável)"
    if nome_dem in ("Status 08/2026", "Status 09/2026", "Status 10/2026") and resumo_inicial:
        return "Finalizado / Concluído"
    if nome_dem == "Complemento ANS 26/10" and resumo_inicial:
        return "Em Andamento / Em Implementação"
    return "Preencher"


def init_db(forcar_reparo=False):
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
                    status_planilha VARCHAR(80) DEFAULT 'Preencher',
                    dados_planilha_json TEXT DEFAULT '{}',
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
            cur.execute("ALTER TABLE demandas ADD COLUMN IF NOT EXISTS oficio_texto TEXT DEFAULT '';")
            cur.execute("ALTER TABLE demandas ADD COLUMN IF NOT EXISTS status_planilha VARCHAR(80) DEFAULT 'Preencher';")
            cur.execute("ALTER TABLE demandas ADD COLUMN IF NOT EXISTS dados_planilha_json TEXT DEFAULT '{}';")
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
                    status_planilha TEXT DEFAULT 'Preencher',
                    dados_planilha_json TEXT DEFAULT '{}',
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
            if "oficio_texto" not in cols:
                cur.execute("ALTER TABLE demandas ADD COLUMN oficio_texto TEXT DEFAULT '';")
            if "status_planilha" not in cols:
                cur.execute("ALTER TABLE demandas ADD COLUMN status_planilha TEXT DEFAULT 'Preencher';")
            if "dados_planilha_json" not in cols:
                cur.execute("ALTER TABLE demandas ADD COLUMN dados_planilha_json TEXT DEFAULT '{}';")
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
                "SELECT id, resumo_texto, status_planilha, dados_planilha_json FROM demandas WHERE item_codigo = %s AND nome = %s",
                (cod, nome_dem)
            )
            bruto_hist = hist_dict.get(nome_dem, "").strip()
            st_ini = inferir_status_inicial(nome_dem, bruto_hist)
            resumo_inicial = "" if bruto_hist in ("N/A", "Preencher", "Prencher") else bruto_hist

            if not ja_tem:
                db_query(
                    """INSERT INTO demandas (item_codigo, nome, tipo, ordem, resumo_texto, status_planilha, dados_planilha_json)
                       VALUES (%s, %s, %s, %s, %s, %s, %s)""",
                    (cod, nome_dem, "FIXA", idx_d, resumo_inicial, st_ini, "{}"),
                    fetch=False,
                    commit=True
                )
            else:
                row_d = ja_tem[0]
                precisa_up = forcar_reparo or (not (row_d.get("resumo_texto") or "").strip() and resumo_inicial)
                if precisa_up:
                    novo_res = (row_d.get("resumo_texto") or "").strip() or resumo_inicial
                    novo_st = row_d.get("status_planilha") or st_ini
                    if novo_st == "Preencher" and st_ini != "Preencher":
                        novo_st = st_ini
                    db_query(
                        "UPDATE demandas SET resumo_texto = %s, status_planilha = %s WHERE id = %s",
                        (novo_res, novo_st, row_d["id"]),
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
# FUNÇÕES AUXILIARES DE ANÁLISE DE ANEXOS, OFÍCIO, PDF E EXPORTAÇÃO EXCEL
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
        f"STATUS NA PLANILHA PRASS: {demanda_row.get('status_planilha') or 'Em Andamento'}",
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


def atualizar_bloco_resumo_no_oficio_salvo(texto_oficio_salvo, novo_resumo):
    novo_resumo = (novo_resumo or "").strip()
    if not novo_resumo or not texto_oficio_salvo:
        return texto_oficio_salvo

    marcador_1 = "1. RESUMO EXECUTIVO E CONSIDERAÇÕES INICIAIS"
    marcador_2 = "2. ANÁLISE TÉCNICA DAS EVIDÊNCIAS E DOCUMENTOS COMPROBATÓRIOS"

    pos_1 = texto_oficio_salvo.find(marcador_1)
    pos_2 = texto_oficio_salvo.find(marcador_2)

    if pos_1 != -1 and pos_2 != -1 and pos_2 > pos_1:
        inicio_bloco_1 = pos_1 + len(marcador_1)
        return (
            texto_oficio_salvo[:inicio_bloco_1]
            + "\n"
            + novo_resumo
            + "\n\n"
            + texto_oficio_salvo[pos_2:]
        )
    return texto_oficio_salvo


def montar_celula_competencia_planilha(dem_row, chks_demanda, texto_hist_original):
    if not dem_row:
        return texto_hist_original if texto_hist_original else "Preencher"

    st_plan = (dem_row.get("status_planilha") or "Preencher").strip()
    if st_plan.startswith("N/A"):
        return "N/A"

    resumo = (dem_row.get("resumo_texto") or "").strip()
    dados_json = {}
    try:
        dados_json = json.loads(dem_row.get("dados_planilha_json") or "{}")
    except Exception:
        dados_json = {}

    blocos = []
    # Se o usuário preencheu blocos por setor no Módulo 2 da Planilha:
    blocos_setor = []
    for s in ("CODER", "COCAP", "CONEC"):
        info_s = dados_json.get(s) or {}
        txt_s = (info_s.get("texto") or "").strip()
        sub_s = (info_s.get("subitem") or "").strip()
        st_s = (info_s.get("status") or "").strip()
        ref_s = (info_s.get("referencias") or "").strip()
        if txt_s or sub_s or ref_s:
            cab_s = f"GERED/{s}"
            if st_s and st_s != "Preencher":
                cab_s += f" [{st_s}]"
            linhas_s = [cab_s]
            if sub_s:
                linhas_s.append(f"Subitem: {sub_s}")
            if ref_s:
                linhas_s.append(f"Referências/Chamados: {ref_s}")
            if txt_s:
                linhas_s.append(txt_s)
            blocos_setor.append("\n".join(linhas_s))

    if resumo:
        blocos.append(resumo)
    elif blocos_setor:
        blocos.append("\n\n".join(blocos_setor))
    elif texto_hist_original and texto_hist_original not in ("Preencher", "Prencher"):
        blocos.append(texto_hist_original)

    if chks_demanda:
        linhas_chk = ["", "[EVIDÊNCIAS E CHECKLIST NO SISTEMA]:"]
        for c in chks_demanda:
            st = "CONCLUÍDO" if bool(c.get("concluido")) else "PENDENTE"
            arqs = [a["nome_arquivo"] for a in c.get("anexos", [])]
            str_arqs = f" | Anexo(s): {', '.join(arqs)}" if arqs else ""
            linhas_chk.append(f"• [{c['setor']} - {st}] {c['titulo']}{str_arqs}")
            if (c.get("analise_evidencia") or "").strip():
                linhas_chk.append(f"  Análise: {c['analise_evidencia'].strip()}")
        blocos.append("\n".join(linhas_chk))

    texto_final = "\n".join(blocos).strip()
    if not texto_final:
        return "Preencher" if st_plan == "Preencher" else f"[{st_plan}]"
    if st_plan and st_plan != "Preencher" and not texto_final.startswith(f"[{st_plan}]"):
        return f"[STATUS GERED: {st_plan.upper()}]\n\n{texto_final}"
    return texto_final


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
            or l_strip.startswith("STATUS NA PLANILHA")
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


@app.route("/api/reset-reparo", methods=["POST"])
def api_reset_reparo():
    """Rota de emergência para sincronizar/reparar todas as tabelas e textos base sem perder anexos."""
    try:
        init_db(forcar_reparo=True)
        return jsonify({"ok": True, "mensagem": "Estrutura do banco e textos base sincronizados com sucesso!"})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"erro": str(e)}), 500


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
            for d in demandas:
                if not d.get("status_planilha"):
                    d["status_planilha"] = "Preencher"
                try:
                    d["dados_planilha"] = json.loads(d.get("dados_planilha_json") or "{}")
                except Exception:
                    d["dados_planilha"] = {}

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
            "ultimo_erro_pg": LAST_PG_ERROR,
            "opcoes_status_planilha": OPCOES_STATUS_PLANILHA,
            "subitens_sugeridos": SUBITENS_SUGERIDOS_POR_SETOR
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({"erro": str(e)}), 500


@app.route("/api/planilha/salvar-celula", methods=["POST"])
def api_salvar_celula_planilha():
    """Salva todos os campos do Módulo 2 (Preenchimento da Planilha PRASS) e sincroniza com a Demanda."""
    body = request.get_json(force=True)
    demanda_id = int(body["demanda_id"])
    status_planilha = body.get("status_planilha", "Preencher").strip()
    resumo_texto = body.get("resumo_texto", "").strip()
    dados_planilha = body.get("dados_planilha", {})

    dem = db_query("SELECT oficio_texto FROM demandas WHERE id = %s", (demanda_id,))
    oficio_atual = (dem[0].get("oficio_texto") or "").strip() if dem else ""

    if oficio_atual and resumo_texto:
        oficio_atual = atualizar_bloco_resumo_no_oficio_salvo(oficio_atual, resumo_texto)

    db_query(
        """UPDATE demandas
           SET status_planilha = %s,
               resumo_texto = %s,
               dados_planilha_json = %s,
               oficio_texto = %s
           WHERE id = %s""",
        (
            status_planilha,
            resumo_texto,
            json.dumps(dados_planilha, ensure_ascii=False),
            oficio_atual,
            demanda_id
        ),
        fetch=False,
        commit=True
    )
    return jsonify({"ok": True, "horario": datetime.now().strftime("%H:%M:%S")})


@app.route("/api/planilha/importar-xlsx", methods=["POST"])
def api_importar_planilha_xlsx():
    """Permite subir uma planilha .xlsx pelo navegador e sincronizar automaticamente os textos e colunas no banco."""
    if not HAS_OPENPYXL:
        return jsonify({"erro": "Biblioteca openpyxl não disponível no servidor."}), 400
    arq = request.files.get("arquivo")
    if not arq or not arq.filename:
        return jsonify({"erro": "Selecione um arquivo Excel (.xlsx)."}), 400

    try:
        raw = arq.read()
        wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
        ws = wb.active
        linhas = list(ws.iter_rows(values_only=True))
        if not linhas:
            return jsonify({"erro": "Planilha vazia."}), 400

        cabecalhos = [str(c).strip() if c is not None else "" for c in linhas[0]]
        itens_db = db_query("SELECT codigo, titulo, historico_json FROM prass_items")
        atualizados = 0

        for row in linhas[1:]:
            if not row or not row[0]:
                continue
            titulo_cel = str(row[0]).strip()
            item_match = next((it for it in itens_db if titulo_cel.startswith(it["codigo"])), None)
            if not item_match:
                continue

            cod = item_match["codigo"]
            hist = json.loads(item_match.get("historico_json") or "{}")

            for col_idx in range(3, len(cabecalhos)):
                nome_col = cabecalhos[col_idx]
                if not nome_col or col_idx >= len(row):
                    continue
                val_cel = str(row[col_idx]).strip() if row[col_idx] is not None else ""
                if not val_cel:
                    continue

                hist[nome_col] = val_cel
                dem_exist = db_query(
                    "SELECT id FROM demandas WHERE item_codigo = %s AND nome = %s",
                    (cod, nome_col)
                )
                st_inf = inferir_status_inicial(nome_col, val_cel)
                res_limpo = "" if val_cel in ("N/A", "Preencher", "Prencher") else val_cel

                if dem_exist:
                    db_query(
                        "UPDATE demandas SET resumo_texto = %s, status_planilha = %s WHERE id = %s",
                        (res_limpo, st_inf, dem_exist[0]["id"]),
                        fetch=False,
                        commit=True
                    )
                else:
                    db_query(
                        """INSERT INTO demandas (item_codigo, nome, tipo, ordem, resumo_texto, status_planilha)
                           VALUES (%s, %s, %s, %s, %s, %s)""",
                        (cod, nome_col, "FIXA", 100 + col_idx, res_limpo, st_inf),
                        fetch=False,
                        commit=True
                    )
            db_query(
                "UPDATE prass_items SET historico_json = %s WHERE codigo = %s",
                (json.dumps(hist, ensure_ascii=False), cod),
                fetch=False,
                commit=True
            )
            atualizados += 1

        return jsonify({"ok": True, "itens_atualizados": atualizados, "colunas": cabecalhos})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"erro": f"Falha ao processar planilha: {e}"}), 500


@app.route("/api/exportar-planilha", methods=["GET"])
def api_exportar_planilha():
    itens = db_query("SELECT * FROM prass_items ORDER BY codigo")
    todas_demandas = db_query("SELECT * FROM demandas ORDER BY ordem ASC, id ASC")
    todos_chks = db_query("SELECT * FROM checklist_evidencias ORDER BY id ASC")
    todos_anexos = db_query("SELECT id, checklist_id, nome_arquivo FROM anexos ORDER BY id ASC")

    anexos_por_chk = {}
    for a in todos_anexos:
        anexos_por_chk.setdefault(a["checklist_id"], []).append(a)

    chks_por_dem = {}
    for c in todos_chks:
        c["anexos"] = anexos_por_chk.get(c["id"], [])
        chks_por_dem.setdefault(c["demanda_id"], []).append(c)

    dems_por_item = {}
    nomes_avulsas = []
    for d in todas_demandas:
        dems_por_item.setdefault(d["item_codigo"], {})[d["nome"]] = d
        if d["tipo"] == "AVULSA" and d["nome"] not in DEMANDAS_FIXAS_PADRAO and d["nome"] not in nomes_avulsas:
            nomes_avulsas.append(d["nome"])

    colunas_competencias = DEMANDAS_FIXAS_PADRAO + nomes_avulsas
    cabecalhos = ["Título", "Descrição", "Observações"] + colunas_competencias

    if HAS_OPENPYXL:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PRASS - Projeto"

        fill_header = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
        fill_comp = PatternFill(start_color="F2A900", end_color="F2A900", fill_type="solid")
        fill_verde = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
        fill_amarelo = PatternFill(start_color="FFFDE7", end_color="FFFDE7", fill_type="solid")

        font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        font_comp = Font(name="Calibri", size=11, bold=True, color="001F3F")
        font_titulo = Font(name="Calibri", size=10, bold=True, color="003366")
        font_corpo = Font(name="Calibri", size=10, color="1E293B")
        borda_fina = Border(
            left=Side(style="thin", color="CBD5E1"),
            right=Side(style="thin", color="CBD5E1"),
            top=Side(style="thin", color="CBD5E1"),
            bottom=Side(style="thin", color="CBD5E1")
        )
        alinhamento_topo = Alignment(horizontal="left", vertical="top", wrap_text=True)
        alinhamento_header = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws.append(cabecalhos)
        ws.row_dimensions[1].height = 30
        for col_idx, nome_col in enumerate(cabecalhos, start=1):
            cell = ws.cell(row=1, column=col_idx)
            if nome_col == "Complemento ANS 26/10":
                cell.fill = fill_comp
                cell.font = font_comp
            else:
                cell.fill = fill_header
                cell.font = font_header
            cell.alignment = alinhamento_header
            cell.border = borda_fina
            letra = get_column_letter(col_idx)
            ws.column_dimensions[letra].width = 36 if col_idx == 1 else (52 if col_idx in (2, 3) else 46)

        for row_idx, it in enumerate(itens, start=2):
            cod = it["codigo"]
            hist = json.loads(it.get("historico_json") or "{}")
            col_b_completa = f"{it['apontamento_ans']}\n\n{it['medidas_implementadas']}".strip()
            linha_vals = [it["titulo"], col_b_completa, it["observacoes"]]

            mapa_dems = dems_por_item.get(cod, {})
            status_cols = []
            for nome_c in colunas_competencias:
                d_row = mapa_dems.get(nome_c)
                lista_c = chks_por_dem.get(d_row["id"], []) if d_row else []
                txt_celula = montar_celula_competencia_planilha(d_row, lista_c, hist.get(nome_c, ""))
                linha_vals.append(txt_celula)
                status_cols.append((d_row.get("status_planilha") or "Preencher") if d_row else "Preencher")

            ws.append(linha_vals)
            ws.row_dimensions[row_idx].height = 180
            for c_i in range(1, len(linha_vals) + 1):
                c_cell = ws.cell(row=row_idx, column=c_i)
                c_cell.font = font_titulo if c_i == 1 else font_corpo
                c_cell.alignment = alinhamento_topo
                c_cell.border = borda_fina
                if c_i >= 4:
                    st_c = status_cols[c_i - 4]
                    if st_c.startswith("Finalizado"):
                        c_cell.fill = fill_verde
                    elif st_c.startswith("Em Andamento") or st_c.startswith("Parcialmente"):
                        c_cell.fill = fill_amarelo

        # Aba 2: Resumo de Status por Área e Item
        ws2 = wb.create_sheet(title="Status GERED por Competência")
        cab2 = ["Item PRASS", "Competência / Demanda", "Status Geral da Entrega", "Subitem CODER", "Subitem COCAP", "Subitem CONEC", "Texto Consolidado da Célula"]
        ws2.append(cab2)
        ws2.row_dimensions[1].height = 26
        for c_i in range(1, len(cab2) + 1):
            c_cell = ws2.cell(row=1, column=c_i)
            c_cell.fill = fill_header
            c_cell.font = font_header
            c_cell.alignment = alinhamento_header
            ws2.column_dimensions[get_column_letter(c_i)].width = 26 if c_i < 7 else 55

        for d in todas_demandas:
            dp = {}
            try:
                dp = json.loads(d.get("dados_planilha_json") or "{}")
            except Exception:
                dp = {}
            ws2.append([
                d["item_codigo"],
                d["nome"],
                d.get("status_planilha") or "Preencher",
                (dp.get("CODER") or {}).get("subitem", ""),
                (dp.get("COCAP") or {}).get("subitem", ""),
                (dp.get("CONEC") or {}).get("subitem", ""),
                d.get("resumo_texto") or "Preencher"
            ])

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        nome_arq = f"PRASS_Projeto_GERED_Preenchido_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        return send_file(
            out,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=nome_arq
        )

    ths = "".join(f"<th style='background:#003366;color:#FFF;padding:8px;'>{h}</th>" for h in cabecalhos)
    trs = ""
    for it in itens:
        cod = it["codigo"]
        hist = json.loads(it.get("historico_json") or "{}")
        col_b = f"{it['apontamento_ans']}\n\n{it['medidas_implementadas']}".replace("\n", "<br>")
        obs = it["observacoes"].replace("\n", "<br>")
        tds = f"<td><b>{it['titulo']}</b></td><td>{col_b}</td><td>{obs}</td>"
        mapa_dems = dems_por_item.get(cod, {})
        for nome_c in colunas_competencias:
            d_row = mapa_dems.get(nome_c)
            lista_c = chks_por_dem.get(d_row["id"], []) if d_row else []
            txt_cel = montar_celula_competencia_planilha(d_row, lista_c, hist.get(nome_c, "")).replace("\n", "<br>")
            tds += f"<td>{txt_cel}</td>"
        trs += f"<tr>{tds}</tr>"

    html_xls = f"<html xmlns:x='urn:schemas-microsoft-com:office:excel'><head><meta charset='utf-8'></head><body><table border='1'>{ths}{trs}</table></body></html>"
    return send_file(
        io.BytesIO(("\ufeff" + html_xls).encode("utf-8")),
        mimetype="application/vnd.ms-excel",
        as_attachment=True,
        download_name="PRASS_Projeto_GERED_Preenchido.xls"
    )


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
                "INSERT INTO demandas (item_codigo, nome, tipo, ordem, resumo_texto, status_planilha) VALUES (%s, %s, %s, %s, %s, %s)",
                (c, nome, "AVULSA", 500, "", "Em Andamento / Em Implementação"),
                fetch=False,
                commit=True
            )
    return jsonify({"ok": True, "nome": nome})


@app.route("/api/demanda/<int:demanda_id>/resumo", methods=["POST"])
def api_salvar_resumo_demanda(demanda_id):
    body = request.get_json(force=True)
    resumo_texto = body.get("resumo_texto", "").strip()
    status_planilha = body.get("status_planilha")

    dem = db_query("SELECT oficio_texto, status_planilha FROM demandas WHERE id = %s", (demanda_id,))
    oficio_atual = (dem[0].get("oficio_texto") or "").strip() if dem else ""
    st_atual = status_planilha if status_planilha else ((dem[0].get("status_planilha") or "Preencher") if dem else "Preencher")

    if oficio_atual:
        oficio_sincronizado = atualizar_bloco_resumo_no_oficio_salvo(oficio_atual, resumo_texto)
        db_query(
            "UPDATE demandas SET resumo_texto = %s, oficio_texto = %s, status_planilha = %s WHERE id = %s",
            (resumo_texto, oficio_sincronizado, st_atual, demanda_id),
            fetch=False,
            commit=True
        )
    else:
        db_query(
            "UPDATE demandas SET resumo_texto = %s, status_planilha = %s WHERE id = %s",
            (resumo_texto, st_atual, demanda_id),
            fetch=False,
            commit=True
        )
    return jsonify({"ok": True, "resumo_texto": resumo_texto, "status_planilha": st_atual})


@app.route("/api/checklist/<int:demanda_id>", methods=["GET"])
def api_listar_checklist(demanda_id):
    dem = db_query("SELECT * FROM demandas WHERE id = %s", (demanda_id,))
    if not dem:
        return jsonify({"erro": "Demanda não encontrada."}), 404
    dem_row = dem[0]
    try:
        dem_row["dados_planilha"] = json.loads(dem_row.get("dados_planilha_json") or "{}")
    except Exception:
        dem_row["dados_planilha"] = {}

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
    forcar_regeneracao = bool(body.get("regenerar", False))

    dem = db_query("SELECT * FROM demandas WHERE id = %s", (demanda_id,))
    if not dem:
        return jsonify({"erro": "Demanda não encontrada."}), 404
    dem_row = dem[0]

    if "resumo_texto" in body:
        novo_resumo = body["resumo_texto"].strip()
        dem_row["resumo_texto"] = novo_resumo
        db_query(
            "UPDATE demandas SET resumo_texto = %s WHERE id = %s",
            (novo_resumo, demanda_id),
            fetch=False,
            commit=True
        )

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

    oficio_salvo = (dem_row.get("oficio_texto") or "").strip()

    if oficio_salvo and not forcar_regeneracao:
        if "resumo_texto" in body and body["resumo_texto"].strip():
            oficio_salvo = atualizar_bloco_resumo_no_oficio_salvo(oficio_salvo, body["resumo_texto"].strip())
            db_query(
                "UPDATE demandas SET oficio_texto = %s WHERE id = %s",
                (oficio_salvo, demanda_id),
                fetch=False,
                commit=True
            )
        texto_minuta = oficio_salvo
        carregou_salvo = True
    else:
        texto_minuta = construir_texto_oficio_automatico(it_row, dem_row, chks)
        db_query(
            "UPDATE demandas SET oficio_texto = %s WHERE id = %s",
            (texto_minuta, demanda_id),
            fetch=False,
            commit=True
        )
        carregou_salvo = False

    return jsonify({
        "texto_oficio": texto_minuta,
        "carregou_versao_salva": carregou_salvo,
        "arquivos_pdf": arquivos_pdf,
        "arquivos_excel": arquivos_excel,
        "item_codigo": it_row["codigo"],
        "demanda_nome": dem_row["nome"]
    })


@app.route("/api/demanda/<int:demanda_id>/salvar-oficio", methods=["POST"])
def api_salvar_oficio(demanda_id):
    body = request.get_json(force=True)
    texto = body.get("texto_oficio", "").strip()
    db_query(
        "UPDATE demandas SET oficio_texto = %s WHERE id = %s",
        (texto, demanda_id),
        fetch=False,
        commit=True
    )
    return jsonify({"ok": True, "horario": datetime.now().strftime("%H:%M:%S")})


@app.route("/api/demanda/<int:demanda_id>/exportar-word", methods=["POST"])
def api_exportar_word(demanda_id):
    body = request.get_json(force=True)
    texto_oficio = body.get("texto_oficio", "").strip()

    db_query(
        "UPDATE demandas SET oficio_texto = %s WHERE id = %s",
        (texto_oficio, demanda_id),
        fetch=False,
        commit=True
    )

    dem_row = db_query("SELECT * FROM demandas WHERE id = %s", (demanda_id,))[0]
    it_row = db_query("SELECT * FROM prass_items WHERE codigo = %s", (dem_row["item_codigo"],))[0]

    paragrafos_html = ""
    for linha in texto_oficio.splitlines():
        l_strip = linha.strip()
        if not l_strip:
            continue
        if l_strip.startswith("---"):
            paragrafos_html += "<hr style='border:0;border-top:1px solid #CBD5E1;margin:12pt 0;'>"
        elif l_strip.startswith(("RESPOSTA AO OFÍCIO", "ITEM DO PRASS:", "STATUS NA PLANILHA", "1.", "2.", "Evidência(s)", "Situação:")):
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
# FRONTEND WEB COMPLETO COM 2 MÓDULOS INTEGRADOS (EVIDÊNCIAS + PLANILHA PRASS)
# ==============================================================================
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Postal Saúde | Sistema de Automação PRASS — GERED</title>
<style>
    :root {
        --azul-postal: #003366;
        --azul-escuro: #001F3F;
        --amarelo-postal: #F2A900;
        --bg-geral: #F0F4F8;
        --texto: #0F172A;
        --cinza-borda: #CBD5E1;
        --verde: #198754;
        --verde-bg: #D1E7DD;
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
        padding: 10px 22px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        border-bottom: 4px solid var(--amarelo-postal);
        flex-shrink: 0;
        flex-wrap: wrap;
        gap: 12px;
    }
    .brand-title { font-size: 17px; font-weight: 800; letter-spacing: 0.3px; }
    .brand-sub { font-size: 11.5px; color: #FFD100; margin-top: 2px; }
    .header-right { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
    .module-tabs {
        display: flex;
        background: rgba(0,0,0,0.25);
        padding: 4px;
        border-radius: 8px;
        gap: 4px;
    }
    .tab-top-btn {
        border: none;
        background: transparent;
        color: #E2E8F0;
        padding: 7px 14px;
        border-radius: 6px;
        font-size: 12.5px;
        font-weight: 800;
        cursor: pointer;
        transition: all 0.15s;
    }
    .tab-top-btn.active {
        background: var(--amarelo-postal);
        color: var(--azul-escuro);
    }
    .btn-export-excel {
        background: #107C41;
        color: #FFFFFF;
        border: 1px solid #34D399;
        padding: 7px 13px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 800;
        cursor: pointer;
        text-decoration: none;
        display: inline-flex;
        align-items: center;
        gap: 5px;
    }
    .btn-export-excel:hover { background: #0E6A37; }
    .db-badge {
        background: rgba(255,255,255,0.12);
        border: 1px solid rgba(255,255,255,0.25);
        padding: 6px 10px;
        border-radius: 6px;
        font-size: 11.5px;
        font-weight: 600;
        cursor: pointer;
    }
    .main-container { display: flex; flex: 1; overflow: hidden; }
    .sidebar {
        width: 375px;
        background: #FFFFFF;
        border-right: 1px solid var(--cinza-borda);
        display: flex;
        flex-direction: column;
        flex-shrink: 0;
    }
    .sidebar-header {
        padding: 12px 16px;
        background: #F8FAFC;
        border-bottom: 1px solid var(--cinza-borda);
        font-weight: 700;
        color: var(--azul-postal);
        font-size: 12.5px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .items-list { flex: 1; overflow-y: auto; padding: 10px; }
    .prass-item-card {
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 11px;
        margin-bottom: 9px;
        cursor: pointer;
        transition: all 0.15s ease;
        background: #FFFFFF;
    }
    .prass-item-card:hover { border-color: #94A3B8; }
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
        margin-bottom: 5px;
    }
    .prass-item-card.active .item-code-badge {
        background: var(--amarelo-postal);
        color: var(--azul-escuro);
    }
    .item-title-short {
        font-size: 12px;
        font-weight: 700;
        color: #1E293B;
        line-height: 1.35;
        margin-bottom: 6px;
    }
    .semaforo-row { display: flex; align-items: center; gap: 5px; flex-wrap: wrap; margin-top: 5px; }
    .area-pill {
        font-size: 10px;
        font-weight: 700;
        padding: 2px 7px;
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
        padding: 16px 22px;
        display: flex;
        flex-direction: column;
        gap: 14px;
    }
    .top-item-banner {
        background: #FFFFFF;
        border: 1px solid var(--cinza-borda);
        border-radius: 10px;
        padding: 15px 18px;
    }
    .banner-header-row { display: flex; justify-content: space-between; align-items: flex-start; gap: 14px; }
    .banner-title { font-size: 15.5px; font-weight: 800; color: var(--azul-postal); line-height: 1.35; }
    .banner-actions { display: flex; gap: 8px; flex-shrink: 0; align-items: center; }
    .dropdown-areas { position: relative; display: inline-block; }
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
    .planilha-buttons-row { display: flex; gap: 10px; margin-top: 12px; flex-wrap: wrap; }
    .btn-info-planilha {
        flex: 1;
        min-width: 210px;
        padding: 8px 12px;
        border-radius: 7px;
        border: 1px solid #CBD5E1;
        background: #F8FAFC;
        color: #1E293B;
        font-weight: 700;
        font-size: 12px;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .btn-info-planilha.active {
        background: var(--azul-postal);
        color: #FFFFFF;
        border-color: var(--azul-postal);
    }
    .info-drawer {
        display: none;
        margin-top: 10px;
        background: #F8FAFC;
        border-left: 4px solid var(--azul-postal);
        border-radius: 6px;
        padding: 12px 16px;
        font-size: 12.5px;
        line-height: 1.55;
        white-space: pre-wrap;
        color: #1E293B;
        max-height: 250px;
        overflow-y: auto;
    }
    .info-drawer.open { display: block; }
    .controls-bar {
        background: #FFFFFF;
        border: 1px solid var(--cinza-borda);
        border-radius: 10px;
        padding: 12px 18px;
        display: flex;
        flex-wrap: wrap;
        gap: 14px;
        align-items: center;
        justify-content: space-between;
    }
    .control-group { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
    .control-label { font-size: 11.5px; font-weight: 800; text-transform: uppercase; color: #475569; }
    .seg-btn {
        padding: 6px 12px;
        border-radius: 6px;
        border: 1px solid var(--cinza-borda);
        background: #F8FAFC;
        font-size: 12px;
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
        padding: 6px 10px;
        border-radius: 6px;
        border: 1px solid var(--azul-postal);
        font-size: 12.5px;
        font-weight: 700;
        color: var(--azul-postal);
        background: #FFF;
        min-width: 210px;
    }
    .setor-btn {
        padding: 7px 15px;
        border-radius: 7px;
        border: 2px solid transparent;
        font-size: 12.5px;
        font-weight: 800;
        cursor: pointer;
        background: #E2E8F0;
        color: #1E293B;
    }
    .setor-btn.active {
        background: var(--azul-postal);
        color: #FFFFFF;
        border-color: var(--amarelo-postal);
    }
    .resumo-before-box {
        background: #F8FAFC;
        border: 1px solid #93C5FD;
        border-left: 5px solid var(--azul-postal);
        border-radius: 8px;
        padding: 12px 14px;
        display: flex;
        flex-direction: column;
        gap: 8px;
    }
    .resumo-header-line { display: flex; justify-content: space-between; align-items: center; gap: 10px; flex-wrap: wrap; }
    .resumo-textarea {
        width: 100%;
        min-height: 100px;
        padding: 10px 12px;
        border: 1px solid #CBD5E1;
        border-radius: 6px;
        font-family: inherit;
        font-size: 13px;
        line-height: 1.5;
        color: #0F172A;
        resize: vertical;
        background: #FFFFFF;
    }
    .checklist-panel {
        background: #FFFFFF;
        border: 1px solid var(--cinza-borda);
        border-radius: 10px;
        padding: 16px 18px;
        display: flex;
        flex-direction: column;
        gap: 14px;
    }
    .chk-header-row { display: flex; justify-content: space-between; align-items: center; gap: 12px; flex-wrap: wrap; }
    .progress-wrap { flex: 1; min-width: 260px; }
    .progress-bar-outer {
        width: 100%;
        height: 11px;
        background: #E2E8F0;
        border-radius: 6px;
        overflow: hidden;
        margin-top: 5px;
    }
    .progress-bar-inner { height: 100%; background: var(--verde); width: 0%; transition: width 0.25s ease; }
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
        padding: 7px 13px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 12px;
        cursor: pointer;
    }
    .btn-success {
        background: var(--verde);
        color: #FFF;
        border: none;
        padding: 7px 13px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 12px;
        cursor: pointer;
    }
    .btn-warning {
        background: var(--amarelo-postal);
        color: var(--azul-escuro);
        border: none;
        padding: 7px 13px;
        border-radius: 6px;
        font-weight: 800;
        font-size: 12px;
        cursor: pointer;
    }
    .chk-item-card {
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 11px 14px;
        background: #FFFFFF;
        display: flex;
        flex-direction: column;
        gap: 8px;
    }
    .chk-item-card.concluido { border-left: 5px solid var(--verde); background: #F8FFF9; }
    .chk-item-card.pendente { border-left: 5px solid var(--vermelho); }
    .chk-top-line { display: flex; justify-content: space-between; align-items: center; gap: 10px; }
    .chk-title-group { display: flex; align-items: center; gap: 10px; flex: 1; }
    .chk-title-text { font-size: 13px; font-weight: 700; color: #0F172A; }
    .chk-tools { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
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
    .tool-btn.concluir-on { background: var(--verde); color: #FFF; border-color: var(--verde); }
    .tool-btn.danger { background: #FEE2E2; color: #991B1B; border-color: #FCA5A5; }
    .anexos-box {
        background: #F8FAFC;
        border-radius: 6px;
        padding: 7px 10px;
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
    .anexo-chip a { text-decoration: none; font-weight: 700; color: var(--azul-postal); }

    /* ESTILOS DO MÓDULO 2: PREENCHIMENTO DA PLANILHA PRASS */
    .grid-3-setores {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(290px, 1fr));
        gap: 12px;
    }
    .setor-form-card {
        background: #F8FAFC;
        border: 1px solid #CBD5E1;
        border-top: 4px solid var(--azul-postal);
        border-radius: 8px;
        padding: 12px;
        display: flex;
        flex-direction: column;
        gap: 8px;
    }
    .field-lbl { font-size: 11.5px; font-weight: 700; color: #334155; }
    .input-std, .select-std {
        width: 100%;
        padding: 7px 10px;
        border: 1px solid #CBD5E1;
        border-radius: 6px;
        font-size: 12.5px;
        background: #FFF;
    }
    .matriz-table-wrap {
        overflow-x: auto;
        border: 1px solid #CBD5E1;
        border-radius: 8px;
    }
    table.matriz-prass {
        width: 100%;
        border-collapse: collapse;
        font-size: 12px;
    }
    table.matriz-prass th {
        background: var(--azul-postal);
        color: #FFF;
        padding: 8px 10px;
        text-align: left;
        border: 1px solid #CBD5E1;
        white-space: nowrap;
    }
    table.matriz-prass td {
        padding: 8px 10px;
        border: 1px solid #E2E8F0;
        vertical-align: top;
    }
    table.matriz-prass tr:hover { background: #F1F5F9; cursor: pointer; }

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
    }
    .modal-header {
        background: var(--azul-postal);
        color: #FFF;
        padding: 14px 20px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .modal-body { padding: 18px 20px; overflow-y: auto; flex: 1; display: flex; flex-direction: column; gap: 12px; }
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
        <div class="brand-sub">Módulos 8.1.2 a 8.6.1 • Coordenações CODER, COCAP e CONEC • Evidências + Módulo de Preenchimento da Planilha PRASS</div>
    </div>

    <div class="module-tabs">
        <button class="tab-top-btn active" id="tabBtnModulo1" onclick="alternarModulo('MODULO1')">
            📋 1. Checklist, Resumo & Ofício PDF
        </button>
        <button class="tab-top-btn" id="tabBtnModulo2" onclick="alternarModulo('MODULO2')">
            📊 2. Preenchimento da Planilha PRASS
        </button>
    </div>

    <div class="header-right">
        <button class="tool-btn" onclick="acionarImportacaoPlanilha()" title="Subir planilha .xlsx do computador para atualizar dados">
            📤 Importar .xlsx
        </button>
        <a href="/api/exportar-planilha" class="btn-export-excel">
            📊 Exportar Planilha PRASS (.xlsx)
        </a>
        <button class="tool-btn" onclick="executarResetReparo()" title="Sincronizar e reparar estrutura do banco sem apagar anexos">
            🛠️ Reset / Sincronizar
        </button>
        <div class="db-badge" id="dbBadge" onclick="verDiagnosticoBanco()">Conectando...</div>
    </div>
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
        <!-- BANNER FIXO DO ITEM SELECIONADO -->
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

        <!-- BARRA DE SELEÇÃO DE COMPETÊNCIA / DEMANDA E SETOR -->
        <section class="controls-bar">
            <div class="control-group">
                <span class="control-label">Tipo:</span>
                <button class="seg-btn active" id="btnTipoFixa" onclick="selecionarTipoDemanda('FIXA')">Demanda Fixa (08/26 a 08/27)</button>
                <button class="seg-btn" id="btnTipoAvulsa" onclick="selecionarTipoDemanda('AVULSA')">Demanda Avulsa</button>
            </div>

            <div class="control-group">
                <span class="control-label">Competência / Coluna da Planilha:</span>
                <select class="select-demanda" id="selectDemanda" onchange="aoMudarDemanda()"></select>
                <button class="btn-primary" onclick="criarNovaDemandaAvulsa()">+ Nova Demanda</button>
            </div>

            <div class="control-group" id="grupoBotoesSetor">
                <span class="control-label">Setor (GERED):</span>
                <button class="setor-btn" id="btnSetorCODER" onclick="selecionarSetor('CODER')">CODER</button>
                <button class="setor-btn active" id="btnSetorCOCAP" onclick="selecionarSetor('COCAP')">COCAP</button>
                <button class="setor-btn" id="btnSetorCONEC" onclick="selecionarSetor('CONEC')">CONEC</button>
            </div>
        </section>

        <!-- ================================================================== -->
        <!-- VISÃO DO MÓDULO 1: CAIXINHA DE RESUMO + CHECKLIST DE EVIDÊNCIAS    -->
        <!-- ================================================================== -->
        <section class="checklist-panel" id="painelModulo1">
            <div class="resumo-before-box">
                <div class="resumo-header-line">
                    <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
                        <span style="font-size:13.5px; font-weight:800; color:var(--azul-postal);">
                            📝 Resumo Executivo / Texto da Célula da Planilha (Antes das Evidências)
                        </span>
                        <select id="selectStatusRapidoM1" class="select-std" style="width:auto; font-weight:700; color:var(--azul-postal);" onchange="salvarResumoDemanda(false)"></select>
                    </div>
                    <div style="display:flex; gap:6px; align-items:center;">
                        <span id="statusResumoSalvo" style="font-size:11.5px; color:var(--verde); font-weight:700;"></span>
                        <button class="tool-btn" onclick="importarResumoPlanilhaBase()">📥 Puxar Texto Base da Planilha</button>
                        <button class="btn-primary" style="padding:5px 12px; font-size:12px;" onclick="salvarResumoDemanda(true)">💾 Salvar Resumo</button>
                    </div>
                </div>
                <textarea
                    id="textareaResumoDemanda"
                    class="resumo-textarea"
                    placeholder="Digite aqui o resumo geral deste item do PRASS antes das evidências (sai na Seção 1 do Ofício e na célula correspondente da Planilha PRASS)..."
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

        <!-- ================================================================== -->
        <!-- VISÃO DO MÓDULO 2: PREENCHIMENTO ESTRUTURADO DA PLANILHA PRASS     -->
        <!-- ================================================================== -->
        <section class="checklist-panel" id="painelModulo2" style="display:none;">
            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; border-bottom:1px solid #E2E8F0; padding-bottom:10px;">
                <div>
                    <h3 style="color:var(--azul-postal); font-size:15.5px;">
                        📊 Módulo de Preenchimento Direto da Planilha PRASS — <span id="lblItemCompetenciaM2"></span>
                    </h3>
                    <p style="font-size:12px; color:#475569; margin-top:2px;">
                        Preencha o status (ex.: Finalizado, Em Andamento, Preencher) e os campos de cada coordenação da GERED. Tudo será exportado formatado na planilha .xlsx.
                    </p>
                </div>
                <div style="display:flex; gap:8px; align-items:center;">
                    <span id="msgSalvoPlanilhaM2" style="font-size:12px; font-weight:700; color:var(--verde);"></span>
                    <button class="tool-btn" onclick="consolidarSetoresNaCelula()">⚡ Juntar CODER + COCAP + CONEC no Texto Final</button>
                    <button class="btn-success" onclick="salvarModuloPlanilha(true)">💾 Salvar Preenchimento da Planilha</button>
                </div>
            </div>

            <!-- STATUS GERAL DA ENTREGA NA COMPETÊNCIA -->
            <div style="display:flex; gap:14px; align-items:center; flex-wrap:wrap; background:#F8FAFC; padding:12px; border-radius:8px; border:1px solid #CBD5E1;">
                <div style="flex:1; min-width:240px;">
                    <label class="field-lbl">Status Geral desta Competência na Planilha PRASS:</label>
                    <select id="selStatusGeralM2" class="select-std" style="margin-top:4px; font-weight:800; color:var(--azul-postal);" onchange="salvarModuloPlanilha(false)"></select>
                </div>
                <div style="flex:2; min-width:300px;">
                    <label class="field-lbl">Atalhos Rápidos de Status:</label>
                    <div style="display:flex; gap:6px; margin-top:4px; flex-wrap:wrap;">
                        <button class="tool-btn concluir-on" onclick="definirStatusRapidoM2('Finalizado / Concluído')">✅ Marcar como Finalizado</button>
                        <button class="tool-btn" onclick="definirStatusRapidoM2('Em Andamento / Em Implementação')">🟡 Em Andamento</button>
                        <button class="tool-btn" onclick="definirStatusRapidoM2('Aguardando Área Técnica (GETEC / GECOF / GEREL)')">⏳ Aguardando TI / Outra Área</button>
                        <button class="tool-btn" onclick="definirStatusRapidoM2('Preencher')">📋 Preencher</button>
                        <button class="tool-btn" onclick="definirStatusRapidoM2('N/A (Não Aplicável)')">⚪ N/A</button>
                    </div>
                </div>
            </div>

            <!-- CARDS DAS 3 COORDENAÇÕES (CODER, COCAP, CONEC) -->
            <div class="grid-3-setores">
                <!-- CODER -->
                <div class="setor-form-card">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <b style="color:var(--azul-postal); font-size:13.5px;">🏢 GERED / CODER</b>
                        <select id="selStatus_CODER" class="select-std" style="width:auto; font-size:11.5px; font-weight:700;"></select>
                    </div>
                    <label class="field-lbl">Subitem / Frente CODER:</label>
                    <input type="text" id="inpSubitem_CODER" class="input-std" list="dlSub_CODER" placeholder="Ex: Vulnerabilidade de Rede / Garantia de Atendimento">
                    <datalist id="dlSub_CODER"></datalist>
                    <label class="field-lbl">Chamados GLPI / Normativos / BI:</label>
                    <input type="text" id="inpRef_CODER" class="input-std" placeholder="Ex: Painel BI, MAN 037, POP 009, IOT 222, IOT 224">
                    <label class="field-lbl">Considerações / Resposta da CODER:</label>
                    <textarea id="txtResp_CODER" class="resumo-textarea" style="min-height:90px;" placeholder="Digite o posicionamento da CODER para esta competência..."></textarea>
                </div>

                <!-- COCAP -->
                <div class="setor-form-card">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <b style="color:var(--azul-postal); font-size:13.5px;">🏢 GERED / COCAP</b>
                        <select id="selStatus_COCAP" class="select-std" style="width:auto; font-size:11.5px; font-weight:700;"></select>
                    </div>
                    <label class="field-lbl">Subitem / Frente COCAP:</label>
                    <input type="text" id="inpSubitem_COCAP" class="input-std" list="dlSub_COCAP" placeholder="Ex: Bases de rede / RN 486 art. 2º / RN 567 art. 11">
                    <datalist id="dlSub_COCAP"></datalist>
                    <label class="field-lbl">Chamados GLPI / Relatórios Benner / IOTs:</label>
                    <input type="text" id="inpRef_COCAP" class="input-std" placeholder="Ex: GLPI 31043, 30182, 30999, 30992, 31012, Benner 988">
                    <label class="field-lbl">Considerações / Resposta da COCAP:</label>
                    <textarea id="txtResp_COCAP" class="resumo-textarea" style="min-height:90px;" placeholder="Digite o posicionamento da COCAP para esta competência..."></textarea>
                </div>

                <!-- CONEC -->
                <div class="setor-form-card">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <b style="color:var(--azul-postal); font-size:13.5px;">🏢 GERED / CONEC</b>
                        <select id="selStatus_CONEC" class="select-std" style="width:auto; font-size:11.5px; font-weight:700;"></select>
                    </div>
                    <label class="field-lbl">Subitem / Frente CONEC:</label>
                    <input type="text" id="inpSubitem_CONEC" class="input-std" list="dlSub_CONEC" placeholder="Ex: Atualização MAN006 e IOT036 / Suspensões">
                    <datalist id="dlSub_CONEC"></datalist>
                    <label class="field-lbl">Normativos / Acordos / Processos:</label>
                    <input type="text" id="inpRef_CONEC" class="input-std" placeholder="Ex: MAN 006 (DocSpider), IOT 036.07, IOT 031, Acordos">
                    <label class="field-lbl">Considerações / Resposta da CONEC:</label>
                    <textarea id="txtResp_CONEC" class="resumo-textarea" style="min-height:90px;" placeholder="Digite o posicionamento da CONEC para esta competência..."></textarea>
                </div>
            </div>

            <!-- TEXTO CONSOLIDADO FINAL DA CÉLULA DA PLANILHA -->
            <div class="resumo-before-box">
                <div class="resumo-header-line">
                    <span style="font-size:13px; font-weight:800; color:var(--azul-postal);">
                        📄 Texto Consolidado da Célula (Como sairá na coluna "<span id="lblColunaExportM2"></span>" da Planilha Excel e no Resumo do Ofício):
                    </span>
                    <div style="display:flex; gap:6px;">
                        <button class="tool-btn" onclick="importarResumoPlanilhaBase()">📥 Restaurar Texto Original da Planilha</button>
                        <button class="btn-primary" onclick="salvarModuloPlanilha(true)">💾 Salvar Tudo</button>
                    </div>
                </div>
                <textarea id="textareaConsolidadoM2" class="resumo-textarea" style="min-height:120px;" onblur="salvarModuloPlanilha(false)"></textarea>
            </div>

            <!-- MATRIZ PANORÂMICA DE TODAS AS COMPETÊNCIAS DOS 9 ITENS -->
            <div style="margin-top:6px;">
                <h4 style="color:var(--azul-postal); font-size:14px; margin-bottom:8px;">
                    🗂️ Visão Panorâmica da Planilha PRASS (Clique em qualquer linha para abrir e editar o Item):
                </h4>
                <div class="matriz-table-wrap" id="tabelaMatrizContainer"></div>
            </div>
        </section>
    </main>
</div>

<!-- MODAL DO OFÍCIO -->
<div class="modal-backdrop" id="modalOficio">
    <div class="modal-window">
        <div class="modal-header">
            <div>
                <div style="font-size:15px; font-weight:800;">📄 ANÁLISE AUTOMÁTICA DE EVIDÊNCIAS & MINUTA DO OFÍCIO</div>
                <div style="font-size:11.5px; color:#FFD100;" id="modalSubtitulo">Carregando versão do Ofício...</div>
            </div>
            <button onclick="fecharModalOficio()" style="background:none; border:none; color:#FFF; font-size:20px; cursor:pointer;">✖</button>
        </div>
        <div class="modal-body">
            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;">
                <div>
                    <span style="font-size:12.5px; font-weight:700; color:#334155;">Texto do Ofício (Editável pelo Analista):</span>
                    <span id="badgeStatusOficioModal" style="font-size:11.5px; font-weight:700; color:var(--verde); margin-left:8px;"></span>
                </div>
                <button class="tool-btn" onclick="regenerarTextoOficio()">🔄 Recriar Ofício do Zero (Puxar Novos Anexos)</button>
            </div>
            <textarea class="oficio-textarea" id="textareaOficio" onblur="salvarRascunhoOficio(true)"></textarea>
            <div id="resumoArquivosModal" style="background:#F8FAFC; border:1px solid #CBD5E1; border-radius:6px; padding:10px 14px; font-size:12px;"></div>
        </div>
        <div class="modal-footer">
            <div style="display:flex; gap:8px;">
                <button class="tool-btn" onclick="salvarRascunhoOficio(false)">💾 Salvar Edição no Banco</button>
                <button class="btn-primary" onclick="baixarOficioEmWord()">📝 Baixar Ofício em Word (.doc)</button>
            </div>
            <div style="display:flex; gap:10px;">
                <button class="tool-btn" onclick="fecharModalOficio()">Fechar (Salva Automático)</button>
                <button class="btn-success" onclick="aprovarEBaixarPDFUnificado()">✅ Aprovar Ofício e Baixar PDF Unificado (Ofício + Anexos)</button>
            </div>
        </div>
    </div>
</div>

<input type="file" id="hiddenFileInput" multiple style="display:none" onchange="executarUploadAnexos()">
<input type="file" id="hiddenXlsxInput" accept=".xlsx,.xls" style="display:none" onchange="executarImportacaoXlsx()">

<script>
let ESTADO_ITENS = [];
let OPCOES_STATUS = [];
let SUBITENS_SUGERIDOS = {};
let MODULO_ATIVO = 'MODULO1';
let ITEM_ATUAL = null;
let TIPO_DEMANDA_ATUAL = 'FIXA';
let DEMANDA_ATUAL_ID = null;
let DEMANDA_ATUAL_NOME = 'Complemento ANS 26/10';
let SETOR_ATUAL = 'COCAP';
let DRAWER_ABERTO = null;
let CHECKLIST_ATUAL = [];
let CHK_ID_PARA_UPLOAD = null;

function alternarModulo(qual) {
    MODULO_ATIVO = qual;
    document.getElementById('tabBtnModulo1').classList.toggle('active', qual === 'MODULO1');
    document.getElementById('tabBtnModulo2').classList.toggle('active', qual === 'MODULO2');
    document.getElementById('painelModulo1').style.display = (qual === 'MODULO1') ? 'flex' : 'none';
    document.getElementById('painelModulo2').style.display = (qual === 'MODULO2') ? 'flex' : 'none';
    if (qual === 'MODULO2') {
        renderizarMatrizPanoramica();
    }
}

async function verDiagnosticoBanco() {
    const r = await fetch('/api/diagnostico', { cache: 'no-store' });
    const d = await r.json();
    alert(
        "DIAGNÓSTICO DE BANCO DE DADOS:\n\n" +
        "• PostgreSQL Ativo: " + (d.postgres_ativo ? "SIM" : "NÃO (Usando Banco Persistente Interno)") + "\n" +
        "• URL Detectada: " + (d.url_detectada || "(Nenhuma)") + "\n" +
        "• Aviso Postgres: " + (d.ultimo_erro_postgres || "Nenhum erro")
    );
}

async function executarResetReparo() {
    if (!confirm('Deseja sincronizar/reparar a estrutura do banco e restaurar os textos base vazios (sem apagar seus anexos)?')) return;
    const r = await fetch('/api/reset-reparo', { method: 'POST' });
    const d = await r.json();
    alert(d.mensagem || d.erro || 'Concluído');
    await carregarSistema();
}

function acionarImportacaoPlanilha() {
    const inp = document.getElementById('hiddenXlsxInput');
    inp.value = '';
    inp.click();
}

async function executarImportacaoXlsx() {
    const inp = document.getElementById('hiddenXlsxInput');
    if (!inp.files || !inp.files.length) return;
    const fd = new FormData();
    fd.append('arquivo', inp.files[0]);
    const r = await fetch('/api/planilha/importar-xlsx', { method: 'POST', body: fd });
    const d = await r.json();
    if (d.erro) {
        alert('Erro: ' + d.erro);
    } else {
        alert(`Planilha importada com sucesso! ${d.itens_atualizados} itens do PRASS sincronizados.`);
        await carregarSistema();
    }
}

function inicializarSelectsStatus() {
    const ids = ['selectStatusRapidoM1', 'selStatusGeralM2', 'selStatus_CODER', 'selStatus_COCAP', 'selStatus_CONEC'];
    ids.forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        el.innerHTML = '';
        OPCOES_STATUS.forEach(st => {
            const o = document.createElement('option');
            o.value = st;
            o.textContent = st;
            el.appendChild(o);
        });
    });

    ['CODER', 'COCAP', 'CONEC'].forEach(s => {
        const dl = document.getElementById('dlSub_' + s);
        if (!dl) return;
        dl.innerHTML = '';
        (SUBITENS_SUGERIDOS[s] || []).forEach(txt => {
            const o = document.createElement('option');
            o.value = txt;
            dl.appendChild(o);
        });
    });
}

async function carregarSistema() {
    try {
        const resp = await fetch(`/api/bootstrap?demanda_nome=${encodeURIComponent(DEMANDA_ATUAL_NOME)}`, { cache: 'no-store' });
        const dados = await resp.json();
        if (!resp.ok || dados.erro) throw new Error(dados.erro || 'Falha HTTP ' + resp.status);
        ESTADO_ITENS = dados.itens;
        OPCOES_STATUS = dados.opcoes_status_planilha || [];
        SUBITENS_SUGERIDOS = dados.subitens_sugeridos || {};

        inicializarSelectsStatus();

        const badge = document.getElementById('dbBadge');
        badge.textContent = dados.postgres_ativo ? '🟢 PostgreSQL Ativo' : '🟢 Banco Ativo';

        if (!ITEM_ATUAL && ESTADO_ITENS.length > 0) {
            ITEM_ATUAL = ESTADO_ITENS.find(i => i.codigo === '8.5.1') || ESTADO_ITENS[0];
        } else if (ITEM_ATUAL) {
            ITEM_ATUAL = ESTADO_ITENS.find(i => i.codigo === ITEM_ATUAL.codigo) || ESTADO_ITENS[0];
        }

        renderizarSidebar();
        renderizarItemAtual();
        if (MODULO_ATIVO === 'MODULO2') renderizarMatrizPanoramica();
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

        const demComp = (item.demandas || []).find(d => d.nome === DEMANDA_ATUAL_NOME);
        const stPlan = demComp ? (demComp.status_planilha || 'Preencher') : 'Preencher';
        const corSt = stPlan.startsWith('Finalizado') ? 'VERDE' : (stPlan === 'Preencher' ? 'VERMELHO' : 'AMARELO');

        let badgesHtml = `<span class="area-pill ${corSt}" title="Status na Planilha">📊 ${stPlan.split('/')[0].trim()}</span>`;
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

    if (DRAWER_ABERTO) atualizarTextoDrawer();
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
    if (DRAWER_ABERTO === 'apontamento') drawer.textContent = ITEM_ATUAL.apontamento_ans;
    else if (DRAWER_ABERTO === 'medidas') drawer.textContent = ITEM_ATUAL.medidas_implementadas;
    else if (DRAWER_ABERTO === 'obs') drawer.textContent = ITEM_ATUAL.observacoes;
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
        opt.textContent = `${d.nome} [${(d.status_planilha || 'Preencher').split('/')[0].trim()}]`;
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
    const stPlan = document.getElementById('selectStatusRapidoM1').value;
    document.getElementById('textareaConsolidadoM2').value = texto;
    document.getElementById('selStatusGeralM2').value = stPlan;

    const resp = await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/resumo`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ resumo_texto: texto, status_planilha: stPlan })
    });
    if (resp.ok) {
        const lbl = document.getElementById('statusResumoSalvo');
        lbl.textContent = '✔ Salvo!';
        setTimeout(() => { lbl.textContent = ''; }, 2500);
        if (mostrarAlerta) alert('Resumo e Status salvos e sincronizados com o Ofício e com a Planilha!');
        await carregarSistema();
    }
}

async function importarResumoPlanilhaBase() {
    if (!ITEM_ATUAL || !DEMANDA_ATUAL_ID) return;
    const histTxt = (ITEM_ATUAL.historico_planilha && ITEM_ATUAL.historico_planilha[DEMANDA_ATUAL_NOME]) || '';
    if (!histTxt || histTxt === 'N/A' || histTxt === 'Preencher') {
        alert('Não há texto histórico pré-cadastrado na planilha inicial para "' + DEMANDA_ATUAL_NOME + '".');
        return;
    }
    document.getElementById('textareaResumoDemanda').value = histTxt;
    document.getElementById('textareaConsolidadoM2').value = histTxt;
    await salvarResumoDemanda(false);
}

async function carregarChecklistDemanda() {
    if (!DEMANDA_ATUAL_ID) return;
    const resp = await fetch(`/api/checklist/${DEMANDA_ATUAL_ID}`, { cache: 'no-store' });
    const dados = await resp.json();
    CHECKLIST_ATUAL = dados.checklist || [];
    ITEM_ATUAL.semaforo = dados.semaforo;

    const dem = dados.demanda || {};
    const resumoSalvo = dem.resumo_texto || '';
    const stSalvo = dem.status_planilha || 'Preencher';
    const dp = dem.dados_planilha || {};

    document.getElementById('textareaResumoDemanda').value = resumoSalvo;
    document.getElementById('selectStatusRapidoM1').value = stSalvo;

    // Popula campos do Módulo 2 (Preenchimento da Planilha PRASS)
    document.getElementById('lblItemCompetenciaM2').textContent = `Item ${ITEM_ATUAL.codigo} • ${DEMANDA_ATUAL_NOME}`;
    document.getElementById('lblColunaExportM2').textContent = DEMANDA_ATUAL_NOME;
    document.getElementById('selStatusGeralM2').value = stSalvo;
    document.getElementById('textareaConsolidadoM2').value = resumoSalvo;

    ['CODER', 'COCAP', 'CONEC'].forEach(s => {
        const info = dp[s] || {};
        document.getElementById('selStatus_' + s).value = info.status || stSalvo;
        document.getElementById('inpSubitem_' + s).value = info.subitem || '';
        document.getElementById('inpRef_' + s).value = info.referencias || '';
        document.getElementById('txtResp_' + s).value = info.texto || '';
    });

    renderizarSemaforoTopo(dados.semaforo);
    renderizarChecklistNaTela();
}

function definirStatusRapidoM2(novoSt) {
    document.getElementById('selStatusGeralM2').value = novoSt;
    document.getElementById('selectStatusRapidoM1').value = novoSt;
    ['CODER', 'COCAP', 'CONEC'].forEach(s => {
        if (ITEM_ATUAL.areas_atribuidas.includes(s)) {
            document.getElementById('selStatus_' + s).value = novoSt;
        }
    });
    salvarModuloPlanilha(false);
}

function consolidarSetoresNaCelula() {
    const partes = [];
    ['CODER', 'COCAP', 'CONEC'].forEach(s => {
        const st = document.getElementById('selStatus_' + s).value;
        const sub = document.getElementById('inpSubitem_' + s).value.trim();
        const ref = document.getElementById('inpRef_' + s).value.trim();
        const txt = document.getElementById('txtResp_' + s).value.trim();
        if (txt || sub || ref) {
            let Bloco = `GERED/${s} [${st}]`;
            if (sub) Bloco += `\nSubitem: ${sub}`;
            if (ref) Bloco += `\nReferências/Chamados: ${ref}`;
            if (txt) Bloco += `\n${txt}`;
            partes.push(Bloco);
        }
    });

    if (CHECKLIST_ATUAL.length > 0) {
        const linhasChk = CHECKLIST_ATUAL.map(c => {
            const st = c.concluido ? 'CONCLUÍDO' : 'EM ANDAMENTO';
            return `• [${c.setor} - ${st}] ${c.titulo}`;
        });
        partes.push(`Resumo das Entregas no Checklist:\n` + linhasChk.join('\n'));
    }

    if (partes.length === 0) {
        alert('Preencha ao menos um campo de CODER, COCAP ou CONEC acima para juntar no texto consolidado.');
        return;
    }
    const unido = partes.join('\n\n');
    document.getElementById('textareaConsolidadoM2').value = unido;
    document.getElementById('textareaResumoDemanda').value = unido;
    salvarModuloPlanilha(false);
}

async function salvarModuloPlanilha(mostrarAlerta) {
    if (!DEMANDA_ATUAL_ID) return;
    const stGeral = document.getElementById('selStatusGeralM2').value;
    const txtCons = document.getElementById('textareaConsolidadoM2').value;
    document.getElementById('selectStatusRapidoM1').value = stGeral;
    document.getElementById('textareaResumoDemanda').value = txtCons;

    const dadosPlanilha = {};
    ['CODER', 'COCAP', 'CONEC'].forEach(s => {
        dadosPlanilha[s] = {
            status: document.getElementById('selStatus_' + s).value,
            subitem: document.getElementById('inpSubitem_' + s).value.trim(),
            referencias: document.getElementById('inpRef_' + s).value.trim(),
            texto: document.getElementById('txtResp_' + s).value.trim()
        };
    });

    const r = await fetch('/api/planilha/salvar-celula', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
            demanda_id: DEMANDA_ATUAL_ID,
            status_planilha: stGeral,
            resumo_texto: txtCons,
            dados_planilha: dadosPlanilha
        })
    });
    if (r.ok) {
        const d = await r.json();
        const msg = document.getElementById('msgSalvoPlanilhaM2');
        msg.textContent = `✔ Planilha salva às ${d.horario}`;
        setTimeout(() => { msg.textContent = ''; }, 3000);
        if (mostrarAlerta) alert('Dados da planilha salvos com sucesso! Você já pode clicar em Exportar Planilha PRASS (.xlsx).');
        await carregarSistema();
    }
}

function renderizarMatrizPanoramica() {
    const box = document.getElementById('tabelaMatrizContainer');
    if (!box || !ESTADO_ITENS.length) return;
    const colsPrincipais = ['Status 08/2026', 'Status 09/2026', 'Status 10/2026', 'Complemento ANS 26/10', 'Status 11/2026', 'Status 12/2026'];

    let html = `<table class="matriz-prass"><thead><tr><th>Item PRASS</th><th>Áreas</th>`;
    colsPrincipais.forEach(c => { html += `<th>${c}</th>`; });
    html += `</tr></thead><tbody>`;

    ESTADO_ITENS.forEach(it => {
        html += `<tr onclick="selecionarItemDaMatriz('${it.codigo}')">`;
        html += `<td><b>${it.titulo}</b></td>`;
        html += `<td>${it.areas_atribuidas.join(', ')}</td>`;
        colsPrincipais.forEach(nomeCol => {
            const dObj = (it.demandas || []).find(d => d.nome === nomeCol);
            const st = dObj ? (dObj.status_planilha || 'Preencher') : 'Preencher';
            const cor = st.startsWith('Finalizado') ? 'VERDE' : (st === 'Preencher' ? 'VERMELHO' : 'AMARELO');
            const resumoCurto = dObj && dObj.resumo_texto ? dObj.resumo_texto.substring(0, 95) + '...' : '(Preencher)';
            html += `<td><span class="area-pill ${cor}">${st.split('/')[0].trim()}</span><div style="margin-top:4px;color:#475569;font-size:11px;">${resumoCurto}</div></td>`;
        });
        html += `</tr>`;
    });

    html += `</tbody></table>`;
    box.innerHTML = html;
}

function selecionarItemDaMatriz(cod) {
    const alvo = ESTADO_ITENS.find(i => i.codigo === cod);
    if (!alvo) return;
    ITEM_ATUAL = alvo;
    renderizarSidebar();
    renderizarItemAtual();
    window.scrollTo({ top: 0, behavior: 'smooth' });
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
                        <a href="/api/anexos/${a.id}/view" target="_blank">👁️ Ver</a>
                        <a href="/api/anexos/${a.id}/download">⬇️ Baixar</a>
                        <button onclick="removerAnexo(${a.id})" style="border:none; background:none; color:#DC3545; font-weight:800; cursor:pointer;">✖</button>
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
    if (!DEMANDA_ATUAL_ID) return alert('Selecione uma demanda primeiro.');
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
        if (!confirm('Este item ainda não possui anexo importado. Deseja marcá-lo como CONCLUÍDO mesmo assim?')) return;
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
    if (!DEMANDA_ATUAL_ID) return alert('Selecione uma demanda primeiro.');
    const resumoAtual = document.getElementById('textareaResumoDemanda').value;
    const resp = await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/gerar-minuta-oficio`, {
        method: 'POST',
        cache: 'no-store',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
            regenerar: false,
            resumo_texto: resumoAtual
        })
    });
    const dados = await resp.json();
    preencherModalOficio(dados);
    document.getElementById('modalOficio').classList.add('open');
}

async function regenerarTextoOficio() {
    if (!confirm('Recriar o Ofício do zero irá montar novamente o texto automático com base no Resumo e nos Anexos atuais, substituindo eventuais edições manuais feitas abaixo. Continuar?')) return;
    const resumoAtual = document.getElementById('textareaResumoDemanda').value;
    const resp = await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/gerar-minuta-oficio`, {
        method: 'POST',
        cache: 'no-store',
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
    const sub = document.getElementById('modalSubtitulo');
    const badge = document.getElementById('badgeStatusOficioModal');

    if (dados.carregou_versao_salva) {
        sub.textContent = '✔ Última versão editada/salva carregada com sucesso do banco de dados.';
        badge.textContent = '✔ Última versão salva ativa';
    } else {
        sub.textContent = 'Minuta gerada automaticamente (Resumo na Seção 1 + Análise das Evidências na Seção 2).';
        badge.textContent = '✨ Nova minuta gerada e salva';
    }

    const box = document.getElementById('resumoArquivosModal');
    const qtdPdf = (dados.arquivos_pdf || []).length;
    const qtdXls = (dados.arquivos_excel || []).length;

    let html = `<b>📦 Composição do Dossiê Unificado:</b> `;
    html += `1º) Ofício de Resposta (Resumo + Análise) | 2º) ${qtdPdf} anexo(s) PDF/Imagem mesclados no mesmo PDF`;
    if (qtdXls > 0) {
        const nomesXls = dados.arquivos_excel.map(x => x.nome_arquivo).join(', ');
        html += ` | 3º) ${qtdXls} planilha(s) citada(s) pelo nome: <code>${nomesXls}</code>.`;
    }
    box.innerHTML = html;
}

async function fecharModalOficio() {
    if (DEMANDA_ATUAL_ID) {
        const texto = document.getElementById('textareaOficio').value;
        await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/salvar-oficio`, {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({ texto_oficio: texto })
        });
    }
    document.getElementById('modalOficio').classList.remove('open');
}

async function salvarRascunhoOficio(silencioso) {
    if (!DEMANDA_ATUAL_ID) return;
    const texto = document.getElementById('textareaOficio').value;
    const resp = await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/salvar-oficio`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ texto_oficio: texto })
    });
    if (resp.ok) {
        const d = await resp.json();
        document.getElementById('badgeStatusOficioModal').textContent = `✔ Última versão salva às ${d.horario || 'agora'}`;
        if (!silencioso) alert('Última versão do Ofício salva com sucesso no banco de dados!');
    }
}

async function baixarOficioEmWord() {
    const texto = document.getElementById('textareaOficio').value;
    await salvarRascunhoOficio(true);
    const resp = await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/exportar-word`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ texto_oficio: texto })
    });
    if (!resp.ok) return alert('Erro ao gerar documento Word.');
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
    await salvarRascunhoOficio(true);
    const resp = await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/exportar-pdf-unificado`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ texto_oficio: texto })
    });
    if (!resp.ok) return alert('Erro ao gerar o PDF unificado.');
    const blob = await resp.blob();
    const nomeSugerido = `Oficio_e_Evidencias_PRASS_${ITEM_ATUAL.codigo.replace(/\./g, '_')}_${DEMANDA_ATUAL_NOME.replace(/[^a-zA-Z0-9]/g, '_')}.pdf`;

    if (window.showSaveFilePicker) {
        try {
            const handle = await window.showSaveFilePicker({
                suggestedName: nomeSugerido,
                types: [{ description: 'Documento PDF Unificado', accept: {'application/pdf': ['.pdf']} }]
            });
            const writable = await handle.createWritable();
            await writable.write(blob);
            await writable.close();
            await fecharModalOficio();
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
    await fecharModalOficio();
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
