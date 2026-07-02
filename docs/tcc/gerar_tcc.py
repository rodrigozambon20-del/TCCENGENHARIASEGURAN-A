# -*- coding: utf-8 -*-
"""
Gerador do TCC — Gestão digital de riscos psicossociais conforme a NR-1 (Psike).
Produz TCC_Riscos_Psicossociais_NR1_MODELO.docx com formatação ABNT
(A4, Arial 12, espaçamento 1,5, margens 3/3/2/2, recuo de 1,25 cm).

Uso:  python3 gerar_tcc.py
Requer: pip install python-docx

Trechos marcados [PREENCHER: ...] dependem de dados reais de campo e ficam
realçados em amarelo no documento — NÃO devem ser inventados.
"""
import copy
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = "TCC_Riscos_Psicossociais_NR1_MODELO.docx"

# ---------------------------------------------------------------- helpers ----

def set_base_styles(doc):
    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    pf = st.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_after = Pt(6)
    for lvl, size in (("Heading 1", 12), ("Heading 2", 12), ("Heading 3", 12)):
        h = doc.styles[lvl]
        h.font.name = "Arial"
        h.font.size = Pt(size)
        h.font.bold = True
        h.font.color.rgb = None  # preto (herda automático)
        h.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
        h.paragraph_format.space_before = Pt(18)
        h.paragraph_format.space_after = Pt(12)
        h.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

def set_margins(doc):
    for sec in doc.sections:
        sec.page_width, sec.page_height = Cm(21), Cm(29.7)
        sec.top_margin, sec.left_margin = Cm(3), Cm(3)
        sec.bottom_margin, sec.right_margin = Cm(2), Cm(2)

def page_number_header(doc):
    """Número de página no canto superior direito (padrão ABNT)."""
    hdr = doc.sections[0].header
    p = hdr.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run()
    for el, attrs, text in (
        ("w:fldChar", {"w:fldCharType": "begin"}, None),
        ("w:instrText", {"xml:space": "preserve"}, "PAGE"),
        ("w:fldChar", {"w:fldCharType": "end"}, None),
    ):
        node = OxmlElement(el)
        for k, v in attrs.items():
            node.set(qn(k), v)
        if text:
            node.text = text
        run._r.append(node)

def center(doc, text, bold=False, size=12, upper=False, space_after=6):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text.upper() if upper else text)
    r.bold = bold
    r.font.size = Pt(size)
    return p

def add_body(doc, text, indent=True, justify=True):
    """Parágrafo de corpo; trechos [PREENCHER: ...] saem realçados em amarelo."""
    p = doc.add_paragraph()
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent:
        p.paragraph_format.first_line_indent = Cm(1.25)
    rest = text
    while "[PREENCHER" in rest:
        before, _, tail = rest.partition("[PREENCHER")
        marker, _, rest = tail.partition("]")
        if before:
            p.add_run(before)
        r = p.add_run("[PREENCHER" + marker + "]")
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        r.bold = True
    if rest:
        p.add_run(rest)
    return p

def add_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    rest = text
    while "[PREENCHER" in rest:
        before, _, tail = rest.partition("[PREENCHER")
        marker, _, rest = tail.partition("]")
        if before:
            p.add_run(before)
        r = p.add_run("[PREENCHER" + marker + "]")
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        r.bold = True
    if rest:
        p.add_run(rest)
    return p

def add_quote(doc, text):
    """Citação longa ABNT: recuo 4 cm, fonte 10, espaçamento simples."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(4)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    r = p.add_run(text)
    r.font.size = Pt(10)
    return p

def add_ref(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_after = Pt(12)
    p.add_run(text)
    return p

def add_toc(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve")
    instr.text = 'TOC \\o "1-3" \\h \\z \\u'
    sep = OxmlElement("w:fldChar"); sep.set(qn("w:fldCharType"), "separate")
    hint = OxmlElement("w:t")
    hint.text = "Sumário automático: clique com o botão direito e escolha “Atualizar campo” no Word."
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr, sep, hint, end):
        run._r.append(el)

# ---------------------------------------------------------------- documento --

doc = Document()
set_margins(doc)
set_base_styles(doc)
page_number_header(doc)

# ===== Página de instruções do modelo (apagar antes de entregar) =============
center(doc, "COMO USAR ESTE MODELO — APAGUE ESTA PÁGINA ANTES DE ENTREGAR",
       bold=True, size=13, space_after=14)
add_body(doc, "Este documento é o modelo do TCC com todo o texto estruturado. Os "
         "trechos realçados em amarelo, no formato [PREENCHER: ...], dependem de "
         "dados reais (nome da empresa, orientador, resultados de campo, fotos) e "
         "não devem ser inventados. Antes da entrega: (1) preencha todos os campos "
         "amarelos; (2) atualize o Sumário (botão direito sobre ele, “Atualizar "
         "campo”, “Atualizar o índice inteiro”); (3) confira as normas do seu "
         "programa (EPUSP/PECE) quanto a capa, ficha catalográfica e folha de "
         "aprovação; (4) apague esta página.", indent=False)
add_body(doc, "Campos a preencher, na ordem em que aparecem: nome completo do autor; "
         "nome do orientador e titulação; ano/local se diferente; dados da empresa "
         "do estudo de caso; período e amostra da aplicação piloto; resultados "
         "reais por dimensão; fotos/prints da aplicação; plano de ação 5W2H "
         "elaborado com a empresa; e a data das referências normativas consultadas.",
         indent=False)
doc.add_page_break()

# ===== Capa ===================================================================
for _ in range(2):
    center(doc, "")
center(doc, "UNIVERSIDADE DE SÃO PAULO", bold=True)
center(doc, "ESCOLA POLITÉCNICA — PECE", bold=True)
center(doc, "PROGRAMA DE EDUCAÇÃO CONTINUADA EM ENGENHARIA", bold=True)
center(doc, "ESPECIALIZAÇÃO EM ENGENHARIA DE SEGURANÇA DO TRABALHO", bold=True,
       space_after=48)
for _ in range(3):
    center(doc, "")
p = center(doc, "RODRIGO ZAMBON ", bold=True)
r = p.add_run("[PREENCHER: nome completo]")
r.bold = True
r.font.highlight_color = WD_COLOR_INDEX.YELLOW
for _ in range(2):
    center(doc, "")
center(doc, "GESTÃO DIGITAL DE RISCOS PSICOSSOCIAIS CONFORME A NR-1:", bold=True, size=14)
center(doc, "desenvolvimento e aplicação de uma plataforma de gestão de SST "
       "em serviços de distribuição de energia elétrica", bold=True, size=14,
       space_after=48)
for _ in range(6):
    center(doc, "")
center(doc, "São Paulo")
center(doc, "2026")
doc.add_page_break()

# ===== Folha de rosto =========================================================
p = center(doc, "RODRIGO ZAMBON ", bold=True)
r = p.add_run("[PREENCHER: nome completo]")
r.bold = True
r.font.highlight_color = WD_COLOR_INDEX.YELLOW
for _ in range(3):
    center(doc, "")
center(doc, "GESTÃO DIGITAL DE RISCOS PSICOSSOCIAIS CONFORME A NR-1:", bold=True, size=13)
center(doc, "desenvolvimento e aplicação de uma plataforma de gestão de SST "
       "em serviços de distribuição de energia elétrica", bold=True, size=13,
       space_after=36)
nat = doc.add_paragraph()
nat.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
nat.paragraph_format.left_indent = Cm(8)
nat.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
rr = nat.add_run("Monografia apresentada ao Programa de Educação Continuada em "
                 "Engenharia (PECE) da Escola Politécnica da Universidade de São "
                 "Paulo como requisito parcial para a obtenção do título de "
                 "Especialista em Engenharia de Segurança do Trabalho.")
rr.font.size = Pt(11)
nat2 = doc.add_paragraph()
nat2.paragraph_format.left_indent = Cm(8)
nat2.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
rr = nat2.add_run("Orientador(a): Prof. ")
rr.font.size = Pt(11)
rr = nat2.add_run("[PREENCHER: nome e titulação do orientador]")
rr.font.size = Pt(11)
rr.font.highlight_color = WD_COLOR_INDEX.YELLOW
for _ in range(8):
    center(doc, "")
center(doc, "São Paulo")
center(doc, "2026")
doc.add_page_break()

# ===== Resumo =================================================================
center(doc, "RESUMO", bold=True, space_after=18)
add_body(doc, "A atualização da Norma Regulamentadora nº 1 (NR-1) pela Portaria MTE "
         "nº 1.419/2024 tornou explícita a obrigação de identificar perigos e "
         "avaliar riscos ocupacionais considerando os fatores de risco "
         "psicossocial relacionados ao trabalho, no âmbito do Gerenciamento de "
         "Riscos Ocupacionais (GRO) e do Programa de Gerenciamento de Riscos "
         "(PGR). Este trabalho apresenta o desenvolvimento e a aplicação de uma "
         "plataforma digital de baixo custo — denominada Psike — para "
         "operacionalizar essa exigência em uma organização do setor de "
         "distribuição de energia elétrica. A plataforma aplica um instrumento "
         "estruturado de dez itens em seis dimensões psicossociais (carga e ritmo "
         "de trabalho; autonomia e controle; clareza de papel; apoio social e de "
         "liderança; reconhecimento; assédio e violência no trabalho), com coleta "
         "anônima, classificação automática do risco por dimensão em três faixas "
         "e geração automática do texto correspondente para o Inventário de "
         "Riscos Ocupacionais (IRO) do PGR. Módulos complementares digitalizam a "
         "permissão de trabalho para serviços em redes de distribuição (NR-10 e "
         "NR-35), a Avaliação Ergonômica Preliminar (NR-17) e o registro e a "
         "investigação de acidentes e quase acidentes, sobre a mesma base de "
         "dados. A aplicação piloto foi conduzida em [PREENCHER: empresa, período "
         "e número de respondentes]. Os resultados indicam [PREENCHER: síntese "
         "dos resultados reais]. Conclui-se que a digitalização do processo "
         "reduz o custo de conformidade com a NR-1 e melhora a rastreabilidade "
         "exigida para os registros do PGR, respeitando o anonimato dos "
         "respondentes e a Lei Geral de Proteção de Dados.", indent=False)
add_body(doc, "Palavras-chave: riscos psicossociais; NR-1; gerenciamento de riscos "
         "ocupacionais; saúde mental no trabalho; distribuição de energia "
         "elétrica; transformação digital.", indent=False)
doc.add_page_break()

# ===== Abstract ===============================================================
center(doc, "ABSTRACT", bold=True, space_after=18)
add_body(doc, "The update of Brazilian Regulatory Standard No. 1 (NR-1) by MTE "
         "Ordinance No. 1,419/2024 made explicit the obligation to identify "
         "hazards and assess occupational risks considering work-related "
         "psychosocial risk factors within the Occupational Risk Management "
         "(GRO) framework and the Risk Management Program (PGR). This study "
         "presents the development and application of a low-cost digital "
         "platform — named Psike — to operationalize this requirement in an "
         "electric power distribution organization. The platform applies a "
         "structured ten-item instrument across six psychosocial dimensions, "
         "with anonymous data collection, automatic risk classification per "
         "dimension into three bands, and automatic generation of the "
         "corresponding text for the Occupational Risk Inventory (IRO). "
         "Complementary modules digitize work permits for distribution network "
         "services (NR-10 and NR-35), the Preliminary Ergonomic Assessment "
         "(NR-17), and the recording and investigation of accidents and "
         "near misses, on the same data repository. The pilot application was "
         "conducted at [PREENCHER: company, period and sample]. Results indicate "
         "[PREENCHER: summary of actual results]. The study concludes that "
         "digitizing the process reduces the cost of compliance with NR-1 and "
         "improves the traceability required for PGR records, while preserving "
         "respondent anonymity and compliance with the Brazilian General Data "
         "Protection Law (LGPD).", indent=False)
add_body(doc, "Keywords: psychosocial risks; NR-1; occupational risk management; "
         "mental health at work; electric power distribution; digital "
         "transformation.", indent=False)
doc.add_page_break()

# ===== Sumário ================================================================
center(doc, "SUMÁRIO", bold=True, space_after=18)
add_toc(doc)
doc.add_page_break()

# ===== 1 INTRODUÇÃO ===========================================================
doc.add_heading("1 INTRODUÇÃO", level=1)

doc.add_heading("1.1 Contextualização", level=2)
add_body(doc, "Os transtornos mentais e comportamentais relacionados ao trabalho "
         "ocupam posição crescente entre as causas de afastamento no Brasil, e a "
         "Organização Mundial da Saúde reconhece o ambiente de trabalho como "
         "determinante relevante da saúde mental dos trabalhadores (WHO, 2022). "
         "Nesse cenário, a Portaria MTE nº 1.419, de 27 de agosto de 2024, "
         "atualizou a Norma Regulamentadora nº 1 (NR-1) para explicitar que o "
         "levantamento preliminar de perigos e a avaliação de riscos do "
         "Gerenciamento de Riscos Ocupacionais (GRO) devem contemplar os fatores "
         "de risco psicossocial relacionados ao trabalho, ao lado dos "
         "tradicionais agentes físicos, químicos, biológicos, ergonômicos e de "
         "acidentes (BRASIL, 2024). A exigibilidade da nova redação passou a "
         "valer, após prorrogação, a partir de 26 de maio de 2026, alcançando "
         "todo empregador regido pela Consolidação das Leis do Trabalho.")
add_body(doc, "A obrigação não se resume a aplicar um questionário: a organização "
         "deve identificar os fatores, avaliá-los, classificá-los, registrá-los "
         "no Inventário de Riscos Ocupacionais (IRO) do Programa de "
         "Gerenciamento de Riscos (PGR), estabelecer medidas de prevenção com "
         "plano de ação e manter os registros disponíveis e rastreáveis por "
         "longos períodos. Para a maioria das organizações brasileiras — em "
         "especial as de pequeno e médio porte e as equipes de SST enxutas — "
         "esse ciclo completo ainda é executado em papel ou em planilhas "
         "avulsas, com baixa rastreabilidade e alto custo operacional.")
add_body(doc, "No setor de distribuição de energia elétrica, objeto do estudo de "
         "caso deste trabalho, os fatores psicossociais convivem com riscos "
         "ocupacionais graves e bem regulamentados — trabalho em instalações "
         "elétricas energizadas e desenergizadas (NR-10), inclusive no Sistema "
         "Elétrico de Potência (SEP), e trabalho em altura em postes e "
         "estruturas (NR-35). Pressão por restabelecimento rápido do "
         "fornecimento, trabalho em turnos e sobreaviso, exposição a intempéries "
         "e o próprio convívio cotidiano com o risco de acidentes fatais "
         "configuram um conjunto de demandas psicossociais característico da "
         "atividade, o que torna o setor um campo particularmente pertinente "
         "para a gestão integrada proposta.")

doc.add_heading("1.2 Justificativa", level=2)
add_body(doc, "Há uma lacuna prática entre a exigência normativa e os instrumentos "
         "disponíveis. Soluções comerciais de gestão de SST tendem a ter custo "
         "incompatível com organizações menores; questionários validados de "
         "grande porte, como o COPSOQ, exigem competência estatística para "
         "aplicação e interpretação; e a alternativa usual — formulários "
         "impressos tabulados manualmente — fragiliza justamente os atributos "
         "que a NR-1 passou a exigir: anonimato na coleta, consistência na "
         "classificação e rastreabilidade dos registros ao longo do tempo.")
add_body(doc, "Este trabalho se justifica, portanto, por demonstrar a viabilidade "
         "técnica e econômica de uma solução digital de custo praticamente "
         "nulo, construída com tecnologias acessíveis (página web autocontida e "
         "serviços gratuitos de planilha em nuvem), que operacionaliza o ciclo "
         "completo exigido pela NR-1 — da coleta anônima à redação automática "
         "da seção correspondente do IRO — e que se estende, sobre a mesma base "
         "de dados, a outros processos de SST ainda manuais na organização "
         "estudada: permissões de trabalho em campo, avaliação ergonômica "
         "preliminar e registro de acidentes e quase acidentes.")

doc.add_heading("1.3 Objetivos", level=2)
doc.add_heading("1.3.1 Objetivo geral", level=3)
add_body(doc, "Desenvolver e aplicar uma plataforma digital para identificação, "
         "avaliação e documentação de fatores de risco psicossocial conforme a "
         "NR-1, integrada à gestão de riscos ocupacionais de uma organização do "
         "setor de distribuição de energia elétrica.")
doc.add_heading("1.3.2 Objetivos específicos", level=3)
add_bullet(doc, "Sistematizar os requisitos da NR-1 (redação da Portaria MTE nº "
           "1.419/2024) e as referências técnicas aplicáveis à avaliação de "
           "fatores psicossociais, em especial o Guia do MTE e a ISO 45003;")
add_bullet(doc, "Construir um instrumento de avaliação enxuto (dez itens em seis "
           "dimensões), com escala do tipo Likert de cinco pontos, coleta "
           "anônima e critérios objetivos de classificação de risco;")
add_bullet(doc, "Implementar a plataforma digital (frontend web autocontido e "
           "backend gratuito em nuvem) com painel agregado e geração automática "
           "de texto para o Inventário de Riscos Ocupacionais;")
add_bullet(doc, "Digitalizar processos complementares de SST da organização: "
           "permissão de trabalho para serviços em redes de distribuição "
           "(NR-10/NR-35), Avaliação Ergonômica Preliminar (NR-17) e registro "
           "e investigação de acidentes e quase acidentes;")
add_bullet(doc, "Aplicar a plataforma em caráter piloto e discutir os resultados, "
           "as limitações e as condições de generalização da solução.")

doc.add_heading("1.4 Estrutura do trabalho", level=2)
add_body(doc, "Além desta introdução, o trabalho está organizado em quatro seções. "
         "A seção 2 apresenta a fundamentação teórica e normativa. A seção 3 "
         "descreve a metodologia, incluindo a arquitetura da plataforma e o "
         "instrumento de avaliação. A seção 4 apresenta e discute os resultados "
         "da implementação e da aplicação piloto. A seção 5 traz as "
         "considerações finais e sugestões de trabalhos futuros.")
doc.add_page_break()

# ===== 2 FUNDAMENTAÇÃO ========================================================
doc.add_heading("2 FUNDAMENTAÇÃO TEÓRICA E NORMATIVA", level=1)

doc.add_heading("2.1 O GRO e o PGR na NR-1", level=2)
add_body(doc, "Desde a reformulação promovida em 2020, a NR-1 estabelece o "
         "Gerenciamento de Riscos Ocupacionais como processo contínuo composto "
         "por levantamento preliminar de perigos, avaliação de riscos, "
         "classificação, implementação de medidas de prevenção e acompanhamento "
         "do controle. O PGR materializa esse processo em, no mínimo, dois "
         "documentos: o Inventário de Riscos Ocupacionais e o Plano de Ação. A "
         "norma exige que o inventário seja mantido atualizado e que o "
         "histórico das suas atualizações seja retido por período mínimo "
         "estabelecido na própria NR-1 (BRASIL, 2024).")
add_body(doc, "Com a Portaria MTE nº 1.419/2024, os fatores de risco psicossocial "
         "relacionados ao trabalho passaram a integrar expressamente o rol de "
         "perigos a considerar no GRO. Na prática, isso equipara o tratamento "
         "documental dos fatores psicossociais ao dos demais agentes: precisam "
         "ser identificados, avaliados com método consistente, classificados, "
         "inscritos no IRO e vinculados a medidas de prevenção com "
         "responsáveis e prazos no Plano de Ação.")

doc.add_heading("2.2 Fatores de risco psicossocial: conceito e dimensões", level=2)
add_body(doc, "Fatores de risco psicossocial são aspectos da concepção, organização "
         "e gestão do trabalho e de seu contexto social que têm potencial de "
         "causar dano psicológico ou físico ao trabalhador (LEKA; JAIN, 2010). "
         "Modelos clássicos da literatura sustentam a seleção de dimensões de "
         "avaliação: o modelo demanda-controle de Karasek (1979) relaciona o "
         "adoecimento à combinação de altas exigências com baixa autonomia; o "
         "modelo esforço-recompensa de Siegrist (1996) destaca o desequilíbrio "
         "entre o esforço despendido e o reconhecimento recebido; e "
         "instrumentos multidimensionais como o COPSOQ (KRISTENSEN et al., "
         "2005) consolidam dimensões como demandas quantitativas, apoio social, "
         "clareza de papel e comportamentos ofensivos.")
add_body(doc, "Com base nessas referências e nos exemplos do Guia do MTE, este "
         "trabalho adota seis dimensões de avaliação: carga e ritmo de "
         "trabalho; autonomia e controle; clareza de papel; apoio social e de "
         "liderança; reconhecimento; e assédio e violência no trabalho.")

doc.add_heading("2.3 O marco regulatório brasileiro", level=2)
add_body(doc, "A Portaria MTE nº 1.419/2024 alterou a NR-1 para incluir os fatores "
         "de risco psicossocial no processo de identificação de perigos e "
         "avaliação de riscos. Para orientar a aplicação, o Ministério do "
         "Trabalho e Emprego publicou guia informativo sobre fatores de riscos "
         "psicossociais relacionados ao trabalho, com exemplos de fatores, "
         "orientações de método e esclarecimentos sobre o que a fiscalização "
         "espera encontrar documentado (BRASIL, 2025). O guia enfatiza que a "
         "avaliação deve ser coletiva e organizacional — não um diagnóstico "
         "clínico individual — e que o anonimato da coleta é condição para a "
         "fidedignidade das respostas.")
add_body(doc, "Complementarmente, a Lei nº 13.709/2018 (LGPD) condiciona o "
         "tratamento de dados pessoais, o que reforça a opção metodológica por "
         "coleta anônima e análise agregada: não se coletam nome, matrícula ou "
         "qualquer identificador individual, apenas o setor, para permitir a "
         "estratificação mínima exigida pela gestão.")

doc.add_heading("2.4 ISO 45003 e referências internacionais", level=2)
add_body(doc, "A ISO 45003:2021 é a primeira norma internacional dedicada à gestão "
         "de riscos psicossociais, estruturada como diretriz complementar à "
         "ISO 45001:2018. Ela organiza os fatores em três grupos — organização "
         "do trabalho, fatores sociais e ambiente/equipamentos — e recomenda a "
         "integração da gestão psicossocial ao sistema de gestão de SST "
         "existente, em vez de um programa paralelo (ISO, 2021). A Organização "
         "Mundial da Saúde, em suas diretrizes sobre saúde mental no trabalho, "
         "recomenda intervenções organizacionais dirigidas às condições de "
         "trabalho como medida primária, antes de intervenções individuais "
         "(WHO, 2022). Essas referências sustentam duas decisões de projeto da "
         "plataforma: avaliar a organização (e não o indivíduo) e integrar o "
         "resultado diretamente ao PGR.")

doc.add_heading("2.5 Instrumentos de avaliação psicossocial", level=2)
add_body(doc, "Entre os instrumentos consolidados destacam-se o COPSOQ "
         "(Copenhagen Psychosocial Questionnaire), o Job Content Questionnaire "
         "de Karasek e as ferramentas de indicadores de gestão do Health and "
         "Safety Executive britânico. São instrumentos robustos, porém extensos "
         "— versões médias do COPSOQ ultrapassam oitenta itens — e sua "
         "aplicação e interpretação exigem competência técnica pouco disponível "
         "em equipes de SST enxutas. Para viabilizar a adoção, este trabalho "
         "optou por um instrumento próprio e enxuto, de dez itens, com uma a "
         "duas questões por dimensão, assumindo explicitamente o caráter de "
         "triagem (screening): o objetivo é priorizar dimensões para "
         "aprofundamento e ação, e não produzir diagnóstico psicométrico "
         "definitivo. Essa limitação é discutida na seção 4.")

doc.add_heading("2.6 Particularidades do setor de distribuição de energia elétrica",
                level=2)
add_body(doc, "As atividades em redes de distribuição são regidas principalmente "
         "pela NR-10, que estabelece requisitos para serviços em instalações "
         "elétricas desenergizadas (procedimentos de desenergização, bloqueio e "
         "aterramento temporário) e energizadas, inclusive no Sistema Elétrico "
         "de Potência, exigindo trabalhadores autorizados e, para o SEP, "
         "treinamento complementar específico (BRASIL, 2019). O trabalho em "
         "postes e estruturas elevadas sujeita-se ainda à NR-35. A rotina "
         "dessas equipes combina demandas físicas e cognitivas elevadas com "
         "fatores psicossociais típicos: pressão temporal no restabelecimento "
         "de fornecimento, regime de turnos e sobreaviso, trabalho a céu aberto "
         "sob intempéries e a convivência permanente com risco de acidente "
         "grave ou fatal — contexto no qual falhas de atenção têm consequência "
         "severa, o que torna a gestão psicossocial também uma medida de "
         "prevenção de acidentes.")

doc.add_heading("2.7 Transformação digital na gestão de SST", level=2)
add_body(doc, "A digitalização de processos de SST tem como benefícios documentados "
         "a padronização dos registros, a redução do tempo entre coleta e "
         "análise, a eliminação de transcrições manuais e a rastreabilidade "
         "exigida por auditorias e fiscalização. No contexto deste trabalho, a "
         "opção por tecnologias gratuitas e de baixa barreira técnica (página "
         "web estática e planilha em nuvem com script de automação) é uma "
         "decisão deliberada de projeto, voltada à replicabilidade da solução "
         "por profissionais de SST sem apoio de equipe de tecnologia da "
         "informação.")
doc.add_page_break()

# ===== 3 METODOLOGIA ==========================================================
doc.add_heading("3 METODOLOGIA", level=1)

doc.add_heading("3.1 Caracterização da pesquisa", level=2)
add_body(doc, "Trata-se de pesquisa aplicada, de natureza tecnológica, conduzida "
         "na forma de desenvolvimento de artefato (plataforma digital) seguido "
         "de estudo de caso com aplicação piloto. A abordagem é "
         "quali-quantitativa: quantitativa no tratamento das respostas do "
         "instrumento e qualitativa na análise da adequação do processo ao "
         "contexto organizacional. O estudo de caso foi conduzido em "
         "[PREENCHER: caracterização da empresa — porte, região de atuação, "
         "efetivo aproximado e atividades, sem identificar dados sensíveis].")

doc.add_heading("3.2 Arquitetura da plataforma", level=2)
add_body(doc, "A plataforma, denominada Psike, foi construída sobre três decisões "
         "de arquitetura: custo zero de operação, ausência de instalação e "
         "anonimato por construção. O frontend é um único arquivo HTML "
         "autocontido, sem framework nem etapa de compilação, que funciona em "
         "qualquer navegador de computador ou celular. O backend é um script "
         "gratuito do Google Apps Script publicado como aplicativo web, que "
         "grava os registros em abas de uma planilha Google — uma aba por "
         "módulo — servindo como repositório único de dados. O frontend detecta "
         "automaticamente o ambiente: quando a URL do backend está configurada, "
         "os dados são gravados na nuvem e compartilhados entre todos os "
         "dispositivos; sem backend, a aplicação opera em modo local para "
         "demonstração. Não há coleta de nome, e-mail ou identificador "
         "individual em nenhum módulo de avaliação psicossocial.")
add_body(doc, "A plataforma foi organizada em quatro módulos sobre o mesmo "
         "repositório de dados, espelhando o Inventário de Riscos Ocupacionais: "
         "Módulo 1 — avaliação de riscos psicossociais (NR-1/ISO 45003); "
         "Módulo 2 — permissão de trabalho digital para serviços em redes de "
         "distribuição (NR-10 e NR-35); Módulo 3 — Avaliação Ergonômica "
         "Preliminar (NR-17); Módulo 4 — registro e investigação de acidentes "
         "e quase acidentes.")

doc.add_heading("3.3 O instrumento de avaliação psicossocial", level=2)
add_body(doc, "O instrumento contém dez afirmativas redigidas em sentido positivo, "
         "distribuídas em seis dimensões: carga e ritmo de trabalho (2 itens), "
         "autonomia e controle (2), clareza de papel (2), apoio social e de "
         "liderança (2), reconhecimento (1) e assédio e violência no trabalho "
         "(1). O respondente indica concordância em escala Likert de cinco "
         "pontos (1 = discordo totalmente; 5 = concordo totalmente). O único "
         "dado adicional coletado é o setor, opcional, para análise "
         "estratificada. A aplicação é anônima e o respondente é informado "
         "disso antes de iniciar.")
add_body(doc, "Como as afirmativas são positivas, notas altas indicam condição "
         "favorável. A nota média de cada dimensão é classificada em três "
         "faixas de risco: risco baixo (média maior ou igual a 3,8), risco "
         "médio (média entre 2,8 e 3,8) e risco alto (média inferior a 2,8). "
         "Os pontos de corte foram definidos por julgamento técnico, "
         "privilegiando a sensibilidade — na dúvida, a dimensão é classificada "
         "na faixa mais conservadora — e devem ser reavaliados após a "
         "acumulação de séries históricas.")

doc.add_heading("3.4 Geração automática do texto para o IRO", level=2)
add_body(doc, "A cada acesso ao painel, a plataforma consolida as respostas, "
         "calcula as médias por dimensão, aplica a classificação e redige "
         "automaticamente a seção de fatores de risco psicossocial do IRO, "
         "contendo: base normativa, data de geração, método e número de "
         "respondentes, resultado por dimensão com nota média e faixa de "
         "risco, recomendação de plano de ação (5W2H) para as dimensões em "
         "risco médio ou alto e recomendação de reavaliação periódica com "
         "manutenção do histórico. O texto é editável antes da inserção no "
         "documento oficial do PGR, preservando a responsabilidade técnica do "
         "profissional que o subscreve.")

doc.add_heading("3.5 Módulos complementares", level=2)
add_body(doc, "Módulo 2 — Permissão de trabalho digital (NR-10/NR-35): o executante "
         "seleciona o tipo de serviço — rede desenergizada (linha morta), rede "
         "energizada/SEP (linha viva) ou trabalho em altura em postes e "
         "estruturas — e a plataforma apresenta o checklist correspondente. "
         "Para rede desenergizada, o checklist segue a sequência de "
         "desenergização da NR-10 (seccionamento; impedimento de "
         "reenergização; constatação da ausência de tensão; aterramento "
         "temporário com equipotencialização; proteção dos elementos "
         "energizados da zona controlada; sinalização), acrescida da "
         "verificação de trabalhadores autorizados, EPI/EPC isolantes e ordem "
         "de serviço. Para linha viva/SEP, verifica treinamento complementar "
         "SEP, supervisão, luvas isolantes de classe adequada, ferramental "
         "isolado, distâncias de segurança, condições climáticas e comunicação "
         "com o centro de operação. A liberação exige checklist completo, "
         "captura de geolocalização pelo GPS do dispositivo e assinatura "
         "digital em tela — o botão de liberação permanece bloqueado até que "
         "todas as condições sejam atendidas.")
add_body(doc, "Módulo 3 — Avaliação Ergonômica Preliminar (NR-17): formulário "
         "estruturado em cinco blocos de fatores (manuseio de cargas; "
         "mobiliário e postos de trabalho; posturas e exigência física; "
         "organização do trabalho; condições ambientais do posto e trabalho "
         "com telas), avaliados em escala de três pontos. A plataforma "
         "classifica cada bloco e gera parecer automático, indicando a "
         "necessidade de aprofundamento em Análise Ergonômica do Trabalho "
         "(AET) quando algum bloco resulta inadequado, conforme a NR-17.")
add_body(doc, "Módulo 4 — Registro e investigação de acidentes e quase acidentes: "
         "registro rápido em campo de ocorrências, com classificação "
         "(acidente ou quase acidente), descrição, fatores causais em níveis "
         "(causas imediatas, subjacentes e básicas) e ações decorrentes, "
         "alimentando indicadores no mesmo repositório de dados. "
         "[PREENCHER: ajustar esta descrição ao estado final do módulo na "
         "data da entrega].")

doc.add_heading("3.6 Aplicação piloto", level=2)
add_body(doc, "A aplicação piloto foi conduzida em [PREENCHER: empresa/unidade], "
         "no período de [PREENCHER: período], abrangendo [PREENCHER: número] "
         "trabalhadores dos setores de [PREENCHER: setores]. O link da "
         "plataforma foi distribuído por [PREENCHER: canal — ex.: grupos de "
         "trabalho, QR code em DDS], precedido de comunicação sobre o caráter "
         "anônimo e voluntário da participação. [PREENCHER: registrar aqui "
         "eventuais aprovações internas — RH, diretoria — e como foi feita a "
         "sensibilização].")
doc.add_page_break()

# ===== 4 RESULTADOS ===========================================================
doc.add_heading("4 RESULTADOS E DISCUSSÃO", level=1)

doc.add_heading("4.1 A plataforma implementada", level=2)
add_body(doc, "A plataforma foi implementada integralmente e está operacional, "
         "com os módulos 1 a 3 em uso e o módulo 4 em [PREENCHER: estado na "
         "data da entrega]. O fluxo do Módulo 1 — da resposta anônima no "
         "celular ao texto pronto para o IRO — ocorre sem qualquer etapa "
         "manual de tabulação. [PREENCHER: inserir aqui capturas de tela da "
         "aplicação: formulário, painel por dimensão e texto do IRO gerado; "
         "e, se possível, foto do uso em campo do Módulo 2].")
add_body(doc, "O custo de operação é nulo: a hospedagem do arquivo HTML é "
         "gratuita, e o backend utiliza a camada gratuita do Google Apps "
         "Script e do Google Sheets. Como contrapartida, há limites de volume "
         "de requisições e de linhas de planilha, discutidos na seção 4.4.")

doc.add_heading("4.2 Resultados da aplicação piloto", level=2)
add_body(doc, "[PREENCHER: apresentar os resultados reais — número de "
         "respondentes, taxa de adesão, nota média e classificação por "
         "dimensão, comparação entre setores se houver amostra suficiente. "
         "Inserir a tabela de resultados por dimensão e o gráfico do painel. "
         "NÃO inventar dados.]")
add_body(doc, "[PREENCHER: transcrever aqui o texto do IRO gerado pela "
         "plataforma para a aplicação real, como evidência do produto final "
         "do processo.]")

doc.add_heading("4.3 Integração com o PGR e plano de ação", level=2)
add_body(doc, "O texto gerado pela plataforma foi estruturado para inserção "
         "direta na seção de fatores psicossociais do IRO. Para as dimensões "
         "classificadas em risco médio ou alto, recomenda-se a elaboração de "
         "plano de ação no formato 5W2H, com medidas de natureza "
         "organizacional priorizadas sobre medidas individuais, em consonância "
         "com a hierarquia de controles da NR-1 e com as diretrizes da OMS. "
         "[PREENCHER: descrever o plano de ação efetivamente elaborado com a "
         "empresa a partir dos resultados do piloto].")

doc.add_heading("4.4 Limitações", level=2)
add_bullet(doc, "Instrumento de triagem: com dez itens, o instrumento prioriza "
           "viabilidade e adesão; não substitui instrumentos psicométricos "
           "validados quando se exigir diagnóstico aprofundado, e não passou "
           "por validação estatística formal (a validação com amostras "
           "maiores é sugerida como trabalho futuro);")
add_bullet(doc, "Pontos de corte definidos por julgamento técnico, a calibrar "
           "com séries históricas;")
add_bullet(doc, "Ausência de autenticação: qualquer pessoa com o link acessa o "
           "formulário e o painel — aceitável no piloto, mas a segmentação "
           "por empresa/cliente e o controle de acesso são requisitos para "
           "uso comercial;")
add_bullet(doc, "Limites da camada gratuita do Google Sheets/Apps Script quanto "
           "a volume de dados e requisições — a migração para banco de dados "
           "dedicado é o caminho natural em caso de escala;")
add_bullet(doc, "Amostra do piloto restrita a uma organização, o que limita a "
           "generalização direta dos resultados.")

doc.add_heading("4.5 Discussão", level=2)
add_body(doc, "Os resultados [PREENCHER: ajustar conforme os dados reais] "
         "sugerem que a principal barreira à conformidade com a nova NR-1 não "
         "é conceitual, mas operacional: o ciclo identificar–avaliar–"
         "documentar–agir é bem descrito na norma e nas referências, porém "
         "custoso quando executado manualmente. A digitalização de ponta a "
         "ponta reduziu esse custo a praticamente zero na organização "
         "estudada, ao mesmo tempo em que melhorou atributos exigidos pela "
         "norma: anonimato (coleta sem identificadores), consistência "
         "(classificação algorítmica uniforme) e rastreabilidade (registros "
         "com data e hora em repositório único). A extensão da mesma base de "
         "dados aos módulos de permissão de trabalho, ergonomia e acidentes "
         "aponta para o conceito de inventário de riscos vivo, atualizado "
         "pelos próprios processos operacionais de SST, em vez de documento "
         "estático revisado anualmente.")
doc.add_page_break()

# ===== 5 CONSIDERAÇÕES FINAIS =================================================
doc.add_heading("5 CONSIDERAÇÕES FINAIS", level=1)
add_body(doc, "Este trabalho desenvolveu e aplicou uma plataforma digital de "
         "custo zero para a gestão de fatores de risco psicossocial conforme "
         "a NR-1, integrada a módulos complementares de permissão de trabalho "
         "(NR-10/NR-35), avaliação ergonômica preliminar (NR-17) e registro "
         "de acidentes, em uma organização do setor de distribuição de "
         "energia elétrica. Os objetivos propostos foram atingidos: o "
         "instrumento foi construído e aplicado de forma anônima, a "
         "classificação de risco por dimensão foi automatizada e o texto da "
         "seção correspondente do IRO passou a ser gerado automaticamente. "
         "[PREENCHER: uma frase com o principal resultado quantitativo do "
         "piloto].")
add_body(doc, "Como trabalhos futuros, sugerem-se: a validação psicométrica do "
         "instrumento com amostras maiores; a calibração dos pontos de corte "
         "com séries históricas; a implementação de autenticação e segregação "
         "de dados por empresa; a migração do repositório para banco de dados "
         "dedicado; e o acompanhamento longitudinal da eficácia das medidas "
         "de controle implantadas a partir dos resultados, fechando o ciclo "
         "de melhoria contínua previsto no GRO.")
doc.add_page_break()

# ===== REFERÊNCIAS ============================================================
doc.add_heading("REFERÊNCIAS", level=1)
add_ref(doc, "BRASIL. Lei nº 13.709, de 14 de agosto de 2018. Lei Geral de "
        "Proteção de Dados Pessoais (LGPD). Diário Oficial da União, Brasília, "
        "DF, 15 ago. 2018.")
add_ref(doc, "BRASIL. Ministério do Trabalho e Emprego. Norma Regulamentadora "
        "nº 1 (NR-1): disposições gerais e gerenciamento de riscos "
        "ocupacionais. Redação dada pela Portaria MTE nº 1.419, de 27 de "
        "agosto de 2024. Brasília, DF: MTE, 2024.")
add_ref(doc, "BRASIL. Ministério do Trabalho e Emprego. Norma Regulamentadora "
        "nº 10 (NR-10): segurança em instalações e serviços em eletricidade. "
        "Brasília, DF: MTE, 2019.")
add_ref(doc, "BRASIL. Ministério do Trabalho e Emprego. Norma Regulamentadora "
        "nº 17 (NR-17): ergonomia. Redação dada pela Portaria MTP nº 423, de "
        "7 de outubro de 2021. Brasília, DF: MTE, 2021.")
add_ref(doc, "BRASIL. Ministério do Trabalho e Emprego. Norma Regulamentadora "
        "nº 35 (NR-35): trabalho em altura. Brasília, DF: MTE, "
        "[PREENCHER: ano da redação vigente consultada].")
add_ref(doc, "BRASIL. Ministério do Trabalho e Emprego. Guia de informações "
        "sobre os fatores de riscos psicossociais relacionados ao trabalho. "
        "Brasília, DF: MTE, 2025. [PREENCHER: conferir título e ano exatos da "
        "edição consultada].")
add_ref(doc, "INTERNATIONAL ORGANIZATION FOR STANDARDIZATION. ISO 45003:2021 — "
        "Occupational health and safety management — Psychological health and "
        "safety at work: guidelines for managing psychosocial risks. Geneva: "
        "ISO, 2021.")
add_ref(doc, "ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. ABNT NBR ISO 45001:2018 "
        "— Sistemas de gestão de saúde e segurança ocupacional: requisitos com "
        "orientações para uso. Rio de Janeiro: ABNT, 2018.")
add_ref(doc, "KARASEK, R. A. Job demands, job decision latitude, and mental "
        "strain: implications for job redesign. Administrative Science "
        "Quarterly, v. 24, n. 2, p. 285-308, 1979.")
add_ref(doc, "KRISTENSEN, T. S.; HANNERZ, H.; HØGH, A.; BORG, V. The Copenhagen "
        "Psychosocial Questionnaire: a tool for the assessment and improvement "
        "of the psychosocial work environment. Scandinavian Journal of Work, "
        "Environment & Health, v. 31, n. 6, p. 438-449, 2005.")
add_ref(doc, "LEKA, S.; JAIN, A. Health impact of psychosocial hazards at work: "
        "an overview. Geneva: World Health Organization, 2010.")
add_ref(doc, "SIEGRIST, J. Adverse health effects of high-effort/low-reward "
        "conditions. Journal of Occupational Health Psychology, v. 1, n. 1, "
        "p. 27-41, 1996.")
add_ref(doc, "WORLD HEALTH ORGANIZATION. WHO guidelines on mental health at "
        "work. Geneva: WHO, 2022.")

# ===== APÊNDICE ===============================================================
doc.add_page_break()
doc.add_heading("APÊNDICE A — INSTRUMENTO DE AVALIAÇÃO APLICADO", level=1)
add_body(doc, "Itens do questionário (escala: 1 = discordo totalmente a "
         "5 = concordo totalmente; respostas anônimas; único dado adicional: "
         "setor, opcional).", indent=False)
ITENS = [
    ("Carga e ritmo de trabalho", "Tenho tempo suficiente para realizar minhas tarefas com qualidade."),
    ("Carga e ritmo de trabalho", "Raramente preciso trabalhar sob pressão intensa de prazos."),
    ("Autonomia e controle", "Tenho liberdade para decidir como organizar meu trabalho."),
    ("Autonomia e controle", "Minhas opiniões são consideradas nas decisões que afetam meu trabalho."),
    ("Clareza de papel", "Sei exatamente o que se espera de mim no meu cargo."),
    ("Clareza de papel", "Não recebo instruções contraditórias sobre minhas tarefas."),
    ("Apoio social e de liderança", "Posso contar com o apoio da minha liderança quando necessário."),
    ("Apoio social e de liderança", "Existe boa colaboração entre os membros da minha equipe."),
    ("Reconhecimento", "Meu trabalho é reconhecido e valorizado pela organização."),
    ("Assédio e violência no trabalho", "Nos últimos 12 meses, não presenciei nem sofri situações de assédio moral ou sexual no trabalho."),
]
for i, (dim, txt) in enumerate(ITENS, 1):
    add_bullet(doc, f"{i}. ({dim}) {txt}")

doc.save(OUT)
print(f"OK: {OUT} gerado.")
