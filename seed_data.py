# -*- coding: utf-8 -*-
"""
Dados fixos extraídos integralmente da planilha 'PRASS - Projeto - base inicial.xlsx'
Separados em:
1. apontamento_ans (Coluna B - parte 1)
2. medidas_implementadas (Coluna B - a partir de 'Medidas implementadas e em implementação')
3. observacoes (Coluna C - Tratamento e monitoramento)
"""

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
            "de orientação, prevenção, integridade e responsabilização."
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
            "da avaliação dos resultados alcançados e da efetividade das melhorias implementadas."
        ),
        "historico_planilha": {
            "Status 08/2026": "GERED/CODER: Responsável pelo monitoramento dos processos de Garantia de Atendimento e Reembolso (execução GEREL - UAR/NAR). Recursos Tecnológicos: Em andamento solicitação no APP para informar protocolo de atendimento no Reembolso Integral.",
            "Status 09/2026": "GERED/CODER: Monitoramento de GA e Reembolso mantido. Solicitação tecnológica de vínculo obrigatório de protocolo no APP em homologação.",
            "Status 10/2026": "GERED/CODER: Implementado em 01/10/2026 a obrigatoriedade de o beneficiário informar o número do protocolo de atendimento quando solicitar Reembolso Integral no aplicativo, garantindo rastreabilidade.",
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
            "de evidência."
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
            "providenciadas as autorizações necessárias para as despesas relacionadas ao deslocamento."
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
            "Eventual demora decorrente de insuficiência de rede será analisada sob duas perspectivas:\n"
            "1. causa assistencial: indisponibilidade ou inexistência efetiva do prestador;\n"
            "2. resposta operacional: tempo e adequação das providências adotadas pela Operadora para assegurar alternativa assistencial."
        ),
        "historico_planilha": {
            "Status 08/2026": "GERED: Normativos vigentes da Garantia de Atendimento: MAN 037 (v3 - 11/08/2026), POP 009 (v4 - 13/08/2026), POP 012 (v1 - 13/08/2026), POP 013 (v0 - 13/08/2026), IOT 222 e IOT 224.",
            "Status 09/2026": "GERED: IOT 222 e IOT 224 em revisão. Solicitação de subsídios às UAR/NAR para mapear prestadores ativos e revisão do painel de vulnerabilidade.",
            "Status 10/2026": "GERED/CODER: Reestruturação metodológica do painel de análise de vulnerabilidade com cronograma de entregas.",
            "Complemento ANS 26/10": "GERED/CODER - Subitem Vulnerabilidade de Rede: Reestruturação do modelo de análise de vulnerabilidade para fortalecer a rastreabilidade das bases."
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
            "inclusive contratação direta, Tratamento Fora do Domicílio e, quando cabível, reembolso integral."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• A Postal Saúde vem adotando medidas destinadas ao fortalecimento da rede e à redução das vulnerabilidades "
            "relacionadas à suficiência assistencial.\n"
            "• Entre essas medidas está a implementação do Compartilhamento de Gestão de Riscos e Rede Assistencial, "
            "concebido para ampliar a disponibilidade e a capilaridade da rede, especialmente em localidades nas quais "
            "foram identificadas vulnerabilidades ou necessidade de complementação da estrutura assistencial.\n"
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
            "As ocorrências relacionadas a pagamento particular e posterior reembolso permanecem vinculadas ao "
            "monitoramento específico do processo de reembolso, permitindo a análise da cadeia completa:\n"
            "indisponibilidade de rede → alternativa assistencial → eventual pagamento particular → reembolso → encerramento da demanda."
        ),
        "historico_planilha": {
            "Status 08/2026": "GERED: 1) Garantia de atendimento e reembolso conduzidos conforme item 8.1.2; 2) Gestão da Rede Assistencial (Direta/Indireta) sob gestão da CONEC.",
            "Status 09/2026": "GERED: Ratificação do monitoramento integrado com o item 8.1.2 e revisão dos processos pela CONEC.",
            "Status 10/2026": "GERED/CODER: Implementada em 01/10/2026 a obrigatoriedade do número de protocolo de atendimento na solicitação de Reembolso Integral no APP.",
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
            "A Postal Saúde reconhece que o período analisado foi marcado por relevante restrição financeira decorrente da "
            "irregularidade dos repasses da mantenedora, circunstância que produziu impactos sobre a capacidade de "
            "cumprimento tempestivo de obrigações financeiras perante parte da rede prestadora.\n"
            "A documentação institucional demonstra que a restrição financeira repercutiu sobre o relacionamento com "
            "prestadores, incluindo situações de inadimplência, suspensão de atendimento e descredenciamento.\n"
            "O diagnóstico institucional considera que houve efetiva exposição ao risco de desassistência decorrente da "
            "combinação entre restrição financeira e dependência da rede prestadora, passando a questão central a ser a "
            "adoção de medidas capazes de reduzir a vulnerabilidade assistencial."
        ),
        "medidas_implementadas": (
            "Medidas implementadas e em implementação\n\n"
            "• Como resposta estruturante, a Postal Saúde implementou o Compartilhamento de Gestão de Riscos e Rede "
            "Assistencial, utilizando redes disponibilizadas por operadoras parceiras.\n"
            "• A medida foi concebida para ampliar a disponibilidade de prestadores, aumentar a capilaridade da rede, "
            "reduzir vulnerabilidades assistenciais e proporcionar alternativas diante de eventual indisponibilidade da rede direta.\n"
            "• A documentação relativa ao projeto demonstra que sua implementação foi precedida de análises assistenciais, "
            "regulatórias, operacionais, atuariais, econômico-financeiras e jurídicas.\n"
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
            "reclamações, NIPs e demais sinais de risco relacionados à disponibilidade de prestadores."
        ),
        "historico_planilha": {
            "Status 08/2026": "GERED: Esclarecimentos sobre registros de suspensão, descredenciamento e atuação conjunta com GEREM, GEREL e GECOF.",
            "Status 09/2026": "GERED: Acompanhamento das suspensões e direcionamento assistencial junto às UAR/NAR.",
            "Status 10/2026": "GERED/CODER: Revisão do painel de vulnerabilidade e monitoramento das demandas de Garantia de Atendimento.",
            "Complemento ANS 26/10": "GERED/CODER - Subitem Garantia de Atendimento em virtude das suspensões: Direcionamento do fluxo assistencial para prestadores aptos da rede direta/indireta, Contratação Direta e Reembolso Integral."
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
            "Em dezembro de 2024, a Operadora contava com 12.781 prestadores credenciados na rede direta. Entre janeiro de "
            "2025 e julho de 2026, foram incorporados 305 novos prestadores, alcançando 13.086 prestadores na rede direta, "
            "além de 27.100 prestadores disponibilizados pela rede indireta (CNU), totalizando 40.186 prestadores.\n"
            "No mesmo período, foram registrados 580 descredenciamentos, decorrentes de diferentes situações, incluindo "
            "ausência de atendimento, solicitação do prestador, interesse da Operadora, alteração de CNPJ, encerramento "
            "das atividades e migração de tipo de pessoa."
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
            "• Realização de análises de vulnerabilidade para subsidiar decisões de expansão, substituição e reorganização da rede.\n"
            "• Fortalecimento dos controles sistêmicos e automatizados, com iniciativas voltadas à rastreabilidade e integração.\n"
            "• Revisão contínua dos processos operacionais relacionados à gestão da rede."
        ),
        "observacoes": (
            "Tratamento e monitoramento\n\n"
            "O apontamento está sendo tratado no âmbito do processo permanente de governança, monitoramento e melhoria "
            "contínua da rede assistencial.\n"
            "As movimentações da rede são acompanhadas pela GERED por meio dos sistemas corporativos e do Painel de "
            "Vulnerabilidade, permitindo identificar alterações relevantes na composição, disponibilidade e suficiência da "
            "rede e subsidiar as decisões de credenciamento, substituição, expansão ou reorganização.\n"
            "Paralelamente, a Operadora vem aprimorando os controles sistêmicos e os fluxos operacionais destinados a "
            "assegurar maior consistência entre a situação efetiva da rede, as bases corporativas e as informações "
            "disponibilizadas aos beneficiários."
        ),
        "historico_planilha": {
            "Status 08/2026": "GERED: Ações divididas entre CODER (Painel de Vulnerabilidade), CONEC (Aplicação e revisão de normativos) e COCAP (Qualificação cadastral e portal).",
            "Status 09/2026": "GERED (CODER e CONEC): Revisão contínua dos processos operacionais e normativos da rede.",
            "Status 10/2026": "GERED/CODER: Reestruturação do modelo do painel de análise de vulnerabilidade.",
            "Complemento ANS 26/10": "GERED/CODER - Subitem Vulnerabilidade de Rede: Cronograma e aprimoramento metodológico do painel."
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
            "Status 08/2026": "GERED/CODER: Painel de vulnerabilidade em revisão. COCAP: Análises diárias focadas na qualificação contínua do cadastro e construção de painel de controle.",
            "Status 09/2026": "GERED/CODER e COCAP: Continuidade da qualificação cadastral e desenvolvimento do painel de controle.",
            "Status 10/2026": "GERED/CODER e COCAP: Extração das bases conciliadas, chamado GLPI 31043 (alterações cadastrais), relatório Benner 988 (IOT 068) e revisão das IOTs.",
            "Complemento ANS 26/10": "Entrega das evidências do Bloco B5 (8.5.1): bases CSV, chamado GLPI 31043, rotina GERED, relatório Benner 988 e IOTs 031, 068 e 075."
        }
    },
    {
        "codigo": "8.5.2",
        "titulo": "8.5.2. DIVULGAÇÃO DA REDE ASSISTENCIAL NO PORTAL CORPORATIVO DA OPERADORA, NÃO INTEGRALMENTE NOS TERMOS DA RN ANS Nº 486/2022",
        "areas_atribuidas": ["CODER", "COCAP"],
        "apontamento_ans": (
            "Apontamento da ANS\n"
            "A ANS apontou possível não disponibilização integral das informações da rede assistencial no portal "
            "corporativo da Operadora, nos termos estabelecidos pela regulamentação vigente (RN ANS nº 486/2022, art. 2º).\n\n"
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
            "• Aprimoramento dos controles sistêmicos e automatizados (chamados GLPI nº 30182 e nº 30999).\n"
            "• Fortalecimento da rastreabilidade das alterações cadastrais e contratuais.\n"
            "• Revisão contínua dos fluxos operacionais relacionados à atualização das informações no Guia de Saúde."
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
            "Status 08/2026": "GERED/CODER e GERED/COCAP: Desenvolvimento de painel de controle e acompanhamento dos requisitos do portal.",
            "Status 09/2026": "GERED/CODER e GERED/COCAP: Evolução do módulo de gerenciamento de inconsistências do portal.",
            "Status 10/2026": "GERED/COCAP: 1ª Entrega do Gerenciamento de Reportes do Guia Saúde em homologação (GLPI 30182) e abertura do GLPI 30999 para adequação integral à RN 486/2022.",
            "Complemento ANS 26/10": "Entrega das evidências do Bloco B5 (8.5.2) e Bloco C1 (RN 486/2022, art. 2º): GLPI 30182, GLPI 30999, prints datados com URL e Checklist de Auditoria do art. 2º."
        }
    },
    {
        "codigo": "8.5.3",
        "titulo": "8.5.3. NÃO DISPONIBILIZAÇÃO DE INFORMAÇÃO SOBRE A SUBSTITUIÇÃO DE PRESTADORES DE SERVIÇOS DE ATENÇÃO À SAÚDE NÃO HOSPITALARES, CONFORME DETERMINA A RN ANS Nº 567/2022",
        "areas_atribuidas": ["CODER", "COCAP", "CONEC"],
        "apontamento_ans": (
            "Apontamento da ANS\n"
            "A ANS apontou possível não disponibilização de informações relativas à substituição de prestadores de serviços "
            "de atenção à saúde não hospitalares, nos termos da regulamentação vigente (RN ANS nº 567/2022, art. 11).\n\n"
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
            "• Avaliação da necessidade de substituição ou expansão da rede (CTI DIOPE/GERED 0020/2025).\n"
            "• Fortalecimento dos mecanismos de rastreabilidade das movimentações (GLPI nº 30992 e GLPI nº 31012).\n"
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
            "Status 08/2026": "GERED/CODER e COCAP: Monitoramento das substituições no portal e desenvolvimento de controles sistêmicos.",
            "Status 09/2026": "GERED/CODER e COCAP: Levantamento de descredenciamentos e substituições vinculadas no Benner.",
            "Status 10/2026": "GERED/COCAP: Abertura dos chamados GLPI 30992 (relatório consolidado de substitutos no Benner) e GLPI 31012 (pesquisa isolada RN 567/2022 art. 11 no portal).",
            "Complemento ANS 26/10": "Entrega das evidências dos Blocos B6 e C2: GLPI 30992, IOTs 031 e 036, prints Benner x Portal, CTI 0020/2025, GLPI 31012, prints de 3 substituições e Checklist do art. 11."
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
            "atendimento e efetiva capacidade de agendamento dentro do prazo regulamentar."
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
            "• Emissão e encaminhamento de relatórios analíticos às Regionais, subsidiando a atuação dos gestores locais.\n"
            "• Acompanhamento individualizado das demandas críticas e reincidentes, com priorização das situações com "
            "maior risco de extrapolação de prazo.\n"
            "• Utilização de Contratação Direta, Reembolso Integral e Tratamento Fora do Domicílio, quando configuradas as "
            "hipóteses aplicáveis e necessárias à garantia do atendimento.\n"
            "• Padronização dos fluxos operacionais por meio do Manual da Garantia de Atendimento e dos POPs associados.\n"
            "• Automação e melhoria dos formulários e fluxos eletrônicos, inclusive mediante migração para o E-doc."
        ),
        "observacoes": (
            "Tratamento e Monitoramento\n\n"
            "O tratamento do risco de insuficiência de rede passou a ser conduzido sob perspectiva preventiva e contínua, "
            "mediante integração entre monitoramento da demanda, análise da capacidade assistencial da rede e atuação regional.\n"
            "O acompanhamento diário permite identificar demandas próximas do vencimento dos prazos regulamentares e "
            "direcionar tempestivamente a atuação das áreas responsáveis.\n"
            "A GERED/CODER mantém, ainda, rotina de alinhamento com as Regionais, orientação operacional e "
            "disponibilização de ferramentas de BI para consulta da rede, beneficiários, eventos e demandas de Garantia de Atendimento.\n"
            "O monitoramento também contempla a classificação adequada dos protocolos, especialmente após a alteração "
            "processual implementada em outubro de 2025, de modo a evitar que demandas efetivamente caracterizadas como "
            "Garantia de Atendimento permaneçam indevidamente classificadas como Rede Credenciada."
        ),
        "historico_planilha": {
            "Status 08/2026": "GERED/CODER: Monitoramento diário das UAR/NAR com maior incidência em processos de GA fora do prazo, alinhado ao item 8.1.2.",
            "Status 09/2026": "GERED/CODER: Continuidade do acompanhamento diário via BI e alinhamento com as Regionais.",
            "Status 10/2026": "GERED/CODER: Implementada em 01/10/2026 a obrigatoriedade do protocolo de atendimento no aplicativo para Reembolso Integral oriundo de Garantia de Atendimento.",
            "Complemento ANS 26/10": "N/A"
        }
    }
]

# Checklist inicial pré-carregado para os itens 8.4.2, 8.5.1, 8.5.2 e 8.5.3 na demanda "Complemento ANS 26/10"
CHECKLISTS_INICIAIS_COMPLEMENTO = {
    "8.4.2": {
        "CODER": [
            "Demandas de Garantia de Atendimento geradas por suspensões de atendimento e sua solução",
            "Controle mensal de prestadores suspensos e restabelecidos (indicadores do PRASS)"
        ],
        "CONEC": [
            "Relação dos prestadores que suspenderam ou ameaçaram suspender por inadimplência: prestador, UF, data, valor em aberto e situação",
            "Registros de negociação: e-mails, ofícios, notificações recebidas e atas de reunião com prestadores",
            "Termos de acordo, parcelamento ou confissão de dívida e comprovantes de pagamento",
            "Comunicação de restabelecimento do atendimento, com data"
        ]
    },
    "8.5.1": {
        "COCAP": [
            "Extração da base de rede com data de corte, por tipo, UF e especialidade, conciliada com o portal",
            "Relatório de alterações cadastrais do mês: inclusões, alterações e exclusões, com data e responsável",
            "Registros da rotina da GERED: checklist e amostra de validações",
            "Relatório de higienização da rede (IOT 068) do período",
            "IOT 031, 068, 075, 158 e 224 revisadas ou cronograma formal de revisão"
        ],
        "CODER": [
            "Revisão metodológica e reestruturação do Painel de Vulnerabilidade da Rede Prestadora"
        ]
    },
    "8.5.2": {
        "COCAP": [
            "Relatório das notificações 'Encontrou inconsistência?' do portal e prazo de correção (GLPI 30182)",
            "Cronograma do novo painel de controle e prints da versão atual",
            "Bloco C1 (RN 486/2022, art. 2º) - Cronograma de Implantação",
            "Bloco C1 (RN 486/2022, art. 2º) - Demanda de TI aberta (GLPI 30999) e especificação aprovada",
            "Bloco C1 (RN 486/2022, art. 2º) - Homologação e data de entrada em produção",
            "Bloco C1 (RN 486/2022, art. 2º) - Prints datados com URL, antes e depois, em ao menos três tipos de prestador",
            "Bloco C1 (RN 486/2022, art. 2º) - Checklist de conformidade do art. 2º assinado pela área"
        ]
    },
    "8.5.3": {
        "COCAP": [
            "Bloco B6 - Relação do período: prestador, CNPJ, data, motivo, substituto e data de publicação no portal (GLPI 30992)",
            "Bloco B6 - Amostra de processos completos (IOT 031 e 036): solicitação, parecer, aprovação e aditivo ou distrato",
            "Bloco B6 - Prova da comunicação ao beneficiário: print datado e data de publicação (Benner x Guia de Saúde)",
            "Bloco B6 - Indicador do PRASS: percentual de substituições comunicadas no prazo",
            "Bloco B6 - Análise de equivalência do substituto (mesmo serviço e município)",
            "Bloco B6 - Termos aditivos de habilitação e CTI DIOPE/GERED 0020/2025, prometidos no PRASS",
            "Bloco C2 (RN 567/2022, art. 11) - Cronograma de Implantação",
            "Bloco C2 (RN 567/2022, art. 11) - Demanda de TI, especificação e homologação (GLPI 31012)",
            "Bloco C2 (RN 567/2022, art. 11) - Prints datados com URL, antes e depois, de ao menos três substituições",
            "Bloco C2 (RN 567/2022, art. 11) - Saneamento dos cadastros com erro (datas e telefones) e relatório da correção",
            "Bloco C2 (RN 567/2022, art. 11) - Checklist de conformidade do art. 11 assinado"
        ]
    }
}
