# -*- coding: utf-8 -*-
"""
Gerador do TCC — Gestão digital de riscos psicossociais conforme a NR-1 (Psike).
Produz TCC_Riscos_Psicossociais_NR1_MODELO.docx com formatação ABNT
(A4, Arial 12, espaçamento 1,5, margens 3/3/2/2, recuo de 1,25 cm).

Versão revisada após a orientação com a Profa. Renata (reunião 03/07/2026):
- Extensão-alvo de 40 a 80 folhas.
- Revisão de literatura ampla, com definições, justificativa das escolhas e
  referências dos últimos cinco anos, além de seção específica sobre burnout.
- Passo a passo da monografia detalhado na metodologia.
- Resultados e discussão integrados.
- Conclusão curta (se o objetivo foi ou não atingido).
- ANONIMATO: não se expõe o nome da empresa nem imagens/dados de pessoas.
  Nos agradecimentos, agradece-se apenas à CERPRO.

Uso:  python3 gerar_tcc.py
Requer: pip install python-docx

Trechos [PREENCHER: ...] dependem de dados reais de campo e ficam realçados em
amarelo — NÃO devem ser inventados, e NÃO devem identificar empresa/pessoas.
"""
import copy
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
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
    for lvl in ("Heading 1", "Heading 2", "Heading 3"):
        h = doc.styles[lvl]
        h.font.name = "Arial"; h.font.size = Pt(12); h.font.bold = True
        h.font.color.rgb = None
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

def center(doc, text, bold=False, size=12, space_after=6):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.bold = bold; r.font.size = Pt(size)
    return p

def _runs_with_marks(p, text):
    """Adiciona texto ao parágrafo, realçando [PREENCHER: ...] em amarelo."""
    rest = text
    while "[PREENCHER" in rest:
        before, _, tail = rest.partition("[PREENCHER")
        marker, _, rest = tail.partition("]")
        if before:
            p.add_run(before)
        r = p.add_run("[PREENCHER" + marker + "]")
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW; r.bold = True
    if rest:
        p.add_run(rest)

def body(doc, text, indent=True):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent:
        p.paragraph_format.first_line_indent = Cm(1.25)
    _runs_with_marks(p, text)
    return p

def bullet(doc, text, numbered=False):
    p = doc.add_paragraph(style="List Number" if numbered else "List Bullet")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _runs_with_marks(p, text)
    return p

def quote(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(4)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    r = p.add_run(text); r.font.size = Pt(10)
    return p

def ref(doc, text):
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
    hint.text = "Sumário automático — no Word, clique com o botão direito e escolha “Atualizar campo”."
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr, sep, hint, end):
        run._r.append(el)

def h1(doc, t): doc.add_heading(t, level=1)
def h2(doc, t): doc.add_heading(t, level=2)
def h3(doc, t): doc.add_heading(t, level=3)

# ---------------------------------------------------------------- documento --

doc = Document()
set_margins(doc)
set_base_styles(doc)
page_number_header(doc)

# ===== Página de instruções (apagar antes de entregar) =======================
center(doc, "COMO USAR ESTE MODELO — APAGUE ESTA PÁGINA ANTES DE ENTREGAR",
       bold=True, size=13, space_after=14)
body(doc, "Este é o modelo do TCC com o texto estruturado conforme as Diretrizes "
     "da USP/ABCD (5ª ed., 2024) e as orientações da supervisão da monografia. "
     "Os trechos realçados em amarelo, no formato [PREENCHER: ...], dependem de "
     "dados reais e não devem ser inventados.", indent=False)
body(doc, "IMPORTANTE — anonimato exigido pela avaliação: NÃO inserir o nome da "
     "empresa, NÃO inserir fotografias de rosto de trabalhadores e NÃO inserir "
     "qualquer dado que identifique a organização ou as pessoas. Refira-se "
     "sempre à “organização estudada”. As figuras devem ser apenas gráficos "
     "agregados e telas do sistema sem dados pessoais. Nos agradecimentos, "
     "agradece-se apenas à CERPRO.", indent=False)
body(doc, "Antes de entregar: (1) preencha os campos amarelos; (2) atualize o "
     "Sumário (botão direito > Atualizar campo > Atualizar o índice inteiro); "
     "(3) confira capa, folha de aprovação e eventual ficha catalográfica "
     "conforme o programa; (4) verifique a extensão (meta de 40 a 80 folhas); "
     "(5) apague esta página.", indent=False)
doc.add_page_break()

# ===== Capa ===================================================================
center(doc, ""); center(doc, "")
center(doc, "UNIVERSIDADE DE SÃO PAULO", bold=True)
center(doc, "ESCOLA POLITÉCNICA", bold=True)
center(doc, "PROGRAMA DE EDUCAÇÃO CONTINUADA EM ENGENHARIA — PECE", bold=True)
center(doc, "ESPECIALIZAÇÃO EM ENGENHARIA DE SEGURANÇA DO TRABALHO",
       bold=True, space_after=48)
for _ in range(3):
    center(doc, "")
p = center(doc, "RODRIGO ZAMBON ", bold=True)
r = p.add_run("[PREENCHER: nome completo]"); r.bold = True
r.font.highlight_color = WD_COLOR_INDEX.YELLOW
center(doc, ""); center(doc, "")
center(doc, "GESTÃO DIGITAL DE RISCOS PSICOSSOCIAIS CONFORME A NR-1:",
       bold=True, size=14)
center(doc, "desenvolvimento e aplicação de uma plataforma de gestão de "
       "segurança e saúde do trabalho em serviços de distribuição de energia "
       "elétrica", bold=True, size=14, space_after=48)
for _ in range(6):
    center(doc, "")
center(doc, "São Paulo")
center(doc, "2026")
doc.add_page_break()

# ===== Folha de rosto =========================================================
p = center(doc, "RODRIGO ZAMBON ", bold=True)
r = p.add_run("[PREENCHER: nome completo]"); r.bold = True
r.font.highlight_color = WD_COLOR_INDEX.YELLOW
for _ in range(3):
    center(doc, "")
center(doc, "GESTÃO DIGITAL DE RISCOS PSICOSSOCIAIS CONFORME A NR-1:",
       bold=True, size=13)
center(doc, "desenvolvimento e aplicação de uma plataforma de gestão de "
       "segurança e saúde do trabalho em serviços de distribuição de energia "
       "elétrica", bold=True, size=13, space_after=36)
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
rr = nat2.add_run("Supervisão da monografia: Profa. ")
rr.font.size = Pt(11)
rr = nat2.add_run("[PREENCHER: nome da supervisora/orientador]")
rr.font.size = Pt(11); rr.font.highlight_color = WD_COLOR_INDEX.YELLOW
for _ in range(7):
    center(doc, "")
center(doc, "São Paulo")
center(doc, "2026")
doc.add_page_break()

# ===== Agradecimentos =========================================================
center(doc, "AGRADECIMENTOS", bold=True, space_after=18)
body(doc, "À CERPRO, pelo apoio e pela viabilização deste trabalho.")
body(doc, "À supervisão da monografia e ao corpo docente do Programa de Educação "
     "Continuada em Engenharia da Escola Politécnica da USP, pelas orientações "
     "ao longo do curso.")
body(doc, "À organização que abriu as portas para a aplicação deste estudo, aqui "
     "preservada em anonimato, e aos trabalhadores que participaram "
     "voluntariamente da avaliação.")
body(doc, "À minha família, pelo apoio durante a especialização.")
doc.add_page_break()

# ===== Resumo =================================================================
center(doc, "RESUMO", bold=True, space_after=18)
body(doc, "A atualização da Norma Regulamentadora nº 1 (NR-1) pela Portaria MTE "
     "nº 1.419/2024 tornou explícita a obrigação de identificar perigos e "
     "avaliar riscos ocupacionais considerando os fatores de risco psicossocial "
     "relacionados ao trabalho, no âmbito do Gerenciamento de Riscos "
     "Ocupacionais (GRO) e do Programa de Gerenciamento de Riscos (PGR), com "
     "exigibilidade a partir de 26 de maio de 2026. Este trabalho teve por "
     "objetivo desenvolver e aplicar uma plataforma digital de baixo custo — "
     "denominada Psike — para operacionalizar essa exigência em uma organização "
     "do setor de distribuição de energia elétrica. Adotou-se pesquisa aplicada, "
     "de natureza tecnológica, na forma de desenvolvimento de artefato seguido "
     "de estudo de caso com aplicação piloto. A plataforma aplica um instrumento "
     "estruturado de dez itens em seis dimensões psicossociais, com coleta "
     "anônima, classificação automática do risco por dimensão em três faixas e "
     "geração automática do texto correspondente para o Inventário de Riscos "
     "Ocupacionais (IRO). Módulos complementares digitalizam a permissão de "
     "trabalho para serviços em redes de distribuição (NR-10 e NR-35), a "
     "Avaliação Ergonômica Preliminar (NR-17) e o registro e a investigação de "
     "acidentes e quase acidentes, sobre a mesma base de dados. A aplicação "
     "piloto foi conduzida em [PREENCHER: período e número de respondentes, sem "
     "identificar a organização]. Os resultados indicam [PREENCHER: síntese dos "
     "resultados reais]. Conclui-se que a digitalização do processo reduz o "
     "custo de conformidade com a NR-1 e melhora a rastreabilidade exigida para "
     "os registros do PGR, preservando o anonimato dos respondentes e a "
     "conformidade com a Lei Geral de Proteção de Dados.", indent=False)
body(doc, "Palavras-chave: riscos psicossociais; NR-1; gerenciamento de riscos "
     "ocupacionais; saúde mental no trabalho; síndrome de burnout; distribuição "
     "de energia elétrica.", indent=False)
doc.add_page_break()

# ===== Abstract ===============================================================
center(doc, "ABSTRACT", bold=True, space_after=18)
body(doc, "The update of Brazilian Regulatory Standard No. 1 (NR-1) by MTE "
     "Ordinance No. 1,419/2024 made explicit the obligation to identify hazards "
     "and assess occupational risks considering work-related psychosocial risk "
     "factors within the Occupational Risk Management (GRO) framework and the "
     "Risk Management Program (PGR), enforceable from May 26, 2026. This study "
     "aimed to develop and apply a low-cost digital platform — named Psike — to "
     "operationalize this requirement in an electric power distribution "
     "organization. Applied, technological research was adopted, as artifact "
     "development followed by a case study with a pilot application. The "
     "platform applies a structured ten-item instrument across six psychosocial "
     "dimensions, with anonymous data collection, automatic risk classification "
     "per dimension into three bands, and automatic generation of the "
     "corresponding text for the Occupational Risk Inventory (IRO). "
     "Complementary modules digitize work permits for distribution network "
     "services (NR-10 and NR-35), the Preliminary Ergonomic Assessment (NR-17) "
     "and the recording and investigation of accidents and near misses, on the "
     "same data repository. The pilot was conducted with [PREENCHER: period and "
     "sample]. Results indicate [PREENCHER: summary of actual results]. The "
     "study concludes that digitizing the process reduces the cost of "
     "compliance with NR-1 and improves the traceability required for PGR "
     "records, while preserving respondent anonymity and compliance with the "
     "Brazilian General Data Protection Law (LGPD).", indent=False)
body(doc, "Keywords: psychosocial risks; NR-1; occupational risk management; "
     "mental health at work; burnout syndrome; electric power distribution.",
     indent=False)
doc.add_page_break()

# ===== Lista de siglas ========================================================
center(doc, "LISTA DE SIGLAS", bold=True, space_after=18)
SIGLAS = [
    ("AEP", "Análise Ergonômica Preliminar"),
    ("AET", "Análise Ergonômica do Trabalho"),
    ("APR", "Análise Preliminar de Risco"),
    ("CAT", "Comunicação de Acidente de Trabalho"),
    ("CID-11", "Classificação Internacional de Doenças, 11ª revisão"),
    ("CLT", "Consolidação das Leis do Trabalho"),
    ("COPSOQ", "Copenhagen Psychosocial Questionnaire"),
    ("GRO", "Gerenciamento de Riscos Ocupacionais"),
    ("INSS", "Instituto Nacional do Seguro Social"),
    ("IRO", "Inventário de Riscos Ocupacionais"),
    ("ISO", "International Organization for Standardization"),
    ("LGPD", "Lei Geral de Proteção de Dados Pessoais"),
    ("MTE", "Ministério do Trabalho e Emprego"),
    ("NR", "Norma Regulamentadora"),
    ("OMS", "Organização Mundial da Saúde"),
    ("PGR", "Programa de Gerenciamento de Riscos"),
    ("SEP", "Sistema Elétrico de Potência"),
    ("SST", "Segurança e Saúde do Trabalho"),
]
for sig, desc in SIGLAS:
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(sig + " — "); r.bold = True
    p.add_run(desc)
doc.add_page_break()

# ===== Sumário ================================================================
center(doc, "SUMÁRIO", bold=True, space_after=18)
add_toc(doc)
doc.add_page_break()

# ============================================================ 1 INTRODUÇÃO ====
h1(doc, "1 INTRODUÇÃO")

h2(doc, "1.1 Contextualização")
body(doc, "A saúde mental relacionada ao trabalho consolidou-se, na última "
     "década, como uma das principais questões de segurança e saúde "
     "ocupacional. Os transtornos mentais e comportamentais figuram entre as "
     "maiores causas de afastamento previdenciário no Brasil, e levantamento da "
     "Associação Nacional de Medicina do Trabalho, com base em dados oficiais do "
     "Instituto Nacional do Seguro Social, apontou crescimento expressivo dos "
     "afastamentos por problemas de saúde mental entre 2023 e 2025 (ANAMT, "
     "2026). No plano internacional, a Organização Mundial da Saúde reconhece o "
     "ambiente de trabalho como determinante relevante da saúde mental e "
     "recomenda intervenções dirigidas às condições de trabalho como medida "
     "primária de prevenção (WHO, 2022).")
body(doc, "Nesse cenário, a Portaria MTE nº 1.419, de 27 de agosto de 2024, "
     "atualizou a NR-1 para explicitar que o levantamento preliminar de perigos "
     "e a avaliação de riscos do Gerenciamento de Riscos Ocupacionais devem "
     "contemplar os fatores de risco psicossocial relacionados ao trabalho, ao "
     "lado dos tradicionais agentes físicos, químicos, biológicos, ergonômicos "
     "e de acidentes (BRASIL, 2024). Após ajustes de cronograma promovidos pela "
     "Portaria MTE nº 765/2025, a exigibilidade da nova redação passou a valer a "
     "partir de 26 de maio de 2026, alcançando todo empregador regido pela "
     "Consolidação das Leis do Trabalho.")
body(doc, "A obrigação não se resume a aplicar um questionário. A organização "
     "deve identificar os fatores, avaliá-los, classificá-los, registrá-los no "
     "Inventário de Riscos Ocupacionais do Programa de Gerenciamento de Riscos, "
     "estabelecer medidas de prevenção com plano de ação e manter os registros "
     "disponíveis e rastreáveis pelos prazos previstos na norma. Para a maioria "
     "das organizações brasileiras — em especial as de pequeno e médio porte e "
     "as equipes de SST enxutas — esse ciclo completo ainda é executado em papel "
     "ou em planilhas avulsas, com baixa rastreabilidade e alto custo "
     "operacional.")
body(doc, "No setor de distribuição de energia elétrica, objeto do estudo de "
     "caso deste trabalho, os fatores psicossociais convivem com riscos "
     "ocupacionais graves e fortemente regulamentados: o trabalho em "
     "instalações elétricas energizadas e desenergizadas, inclusive no Sistema "
     "Elétrico de Potência, regido pela NR-10, e o trabalho em altura em postes "
     "e estruturas, regido pela NR-35. Estudo clássico com trabalhadores desse "
     "setor identificou prevalência de 20,3% de transtornos mentais comuns, "
     "associada à combinação de alta demanda psicológica, baixo controle e "
     "baixo apoio social (SOUZA et al., 2010). Pressão por restabelecimento "
     "rápido do fornecimento, trabalho em turnos e sobreaviso, exposição a "
     "intempéries e a convivência cotidiana com o risco de acidente fatal "
     "configuram um conjunto de demandas psicossociais característico da "
     "atividade, o que torna o setor particularmente pertinente para a gestão "
     "integrada proposta.")

h2(doc, "1.2 Problema de pesquisa")
body(doc, "Diante da nova exigência da NR-1, coloca-se o seguinte problema: como "
     "operacionalizar, de forma tecnicamente consistente e economicamente "
     "viável, a identificação, a avaliação e a documentação dos fatores de "
     "risco psicossocial em uma organização com equipe de SST reduzida, "
     "preservando o anonimato dos trabalhadores e produzindo registros "
     "rastreáveis para o PGR?")

h2(doc, "1.3 Justificativa")
body(doc, "Há uma lacuna prática entre a exigência normativa e os instrumentos "
     "disponíveis. Soluções comerciais de gestão de SST tendem a ter custo "
     "incompatível com organizações menores; questionários validados de grande "
     "porte, como o COPSOQ, exigem competência estatística para aplicação e "
     "interpretação (KRISTENSEN et al., 2005); e a alternativa usual — "
     "formulários impressos tabulados manualmente — fragiliza justamente os "
     "atributos que a NR-1 passou a exigir: anonimato na coleta, consistência na "
     "classificação e rastreabilidade dos registros ao longo do tempo.")
body(doc, "Este trabalho se justifica por demonstrar a viabilidade técnica e "
     "econômica de uma solução digital de custo praticamente nulo, construída "
     "com tecnologias acessíveis, que operacionaliza o ciclo completo exigido "
     "pela NR-1 — da coleta anônima à redação automática da seção "
     "correspondente do IRO — e que se estende, sobre a mesma base de dados, a "
     "outros processos de SST ainda manuais na organização estudada. A "
     "relevância social é reforçada pelo reconhecimento da síndrome de burnout "
     "como doença ocupacional na 11ª revisão da Classificação Internacional de "
     "Doenças, adotada oficialmente no Brasil a partir de 2025 (OMS, 2019; "
     "TREML et al., 2025), o que amplia as consequências jurídicas e "
     "previdenciárias da ausência de gestão dos fatores psicossociais.")

h2(doc, "1.4 Objetivos")
h3(doc, "1.4.1 Objetivo geral")
body(doc, "Desenvolver e aplicar uma plataforma digital para identificação, "
     "avaliação e documentação de fatores de risco psicossocial conforme a "
     "NR-1, integrada à gestão de riscos ocupacionais de uma organização do "
     "setor de distribuição de energia elétrica.")
h3(doc, "1.4.2 Objetivos específicos")
bullet(doc, "Sistematizar os requisitos da NR-1 (redação da Portaria MTE nº "
       "1.419/2024) e as referências técnicas aplicáveis à avaliação de fatores "
       "psicossociais, em especial a ISO 45003 e as diretrizes da OMS;")
bullet(doc, "Revisar a literatura recente sobre fatores de risco psicossocial e "
       "síndrome de burnout, fundamentando as definições e as escolhas do "
       "instrumento;")
bullet(doc, "Construir um instrumento de avaliação enxuto (dez itens em seis "
       "dimensões), com escala do tipo Likert de cinco pontos, coleta anônima e "
       "critérios objetivos de classificação de risco;")
bullet(doc, "Implementar a plataforma digital (frontend web autocontido e "
       "backend gratuito em nuvem), com painel agregado e geração automática de "
       "texto para o Inventário de Riscos Ocupacionais;")
bullet(doc, "Digitalizar processos complementares de SST: permissão de trabalho "
       "para serviços em redes de distribuição (NR-10/NR-35), Avaliação "
       "Ergonômica Preliminar (NR-17) e registro e investigação de acidentes e "
       "quase acidentes;")
bullet(doc, "Aplicar a plataforma em caráter piloto, preservando o anonimato, e "
       "discutir os resultados, as limitações e as condições de generalização.")

h2(doc, "1.5 Estrutura do trabalho")
body(doc, "Além desta introdução, o trabalho está organizado em quatro seções. "
     "A seção 2 apresenta a revisão de literatura e o marco normativo, "
     "incluindo as definições de fatores psicossociais, os modelos teóricos, a "
     "síndrome de burnout e o setor elétrico. A seção 3 descreve a metodologia, "
     "com o passo a passo da pesquisa, os aspectos éticos, a arquitetura da "
     "plataforma e o instrumento. A seção 4 apresenta e discute, de forma "
     "integrada, os resultados da implementação e da aplicação piloto à luz da "
     "literatura. A seção 5 traz a conclusão. Seguem-se as referências e os "
     "apêndices.")
doc.add_page_break()

# ============================================== 2 REVISÃO DE LITERATURA ========
h1(doc, "2 REVISÃO DE LITERATURA E MARCO NORMATIVO")
body(doc, "Esta seção reúne as definições, os modelos e as evidências que "
     "fundamentam o trabalho. Privilegiaram-se referências dos últimos cinco "
     "anos para o estado atual do tema, complementadas por obras seminais dos "
     "modelos teóricos e por documentos normativos vigentes.")

h2(doc, "2.1 Trabalho e saúde mental: a emergência do risco psicossocial")
body(doc, "A relação entre organização do trabalho e adoecimento psíquico é "
     "objeto de estudo consolidado. A transição de um perfil de riscos "
     "predominantemente físicos para um perfil em que os fatores organizacionais "
     "e psicossociais ganham peso acompanha a mudança na natureza do trabalho — "
     "maior carga cognitiva, intensificação do ritmo, pressão por resultados e "
     "novas formas de organização. A OMS, em suas diretrizes sobre saúde mental "
     "no trabalho, estima impacto econômico expressivo da perda de "
     "produtividade associada a depressão e ansiedade e recomenda que as "
     "intervenções atuem primeiro sobre as condições de trabalho, e só "
     "secundariamente sobre o indivíduo (WHO, 2022).")

h2(doc, "2.2 Definições de fatores de risco psicossocial")
body(doc, "Não há uma definição única de fator de risco psicossocial; convém, "
     "portanto, explicitar as principais e justificar a adotada. Para Leka e "
     "Jain (2010), em documento da OMS, fatores de risco psicossocial são "
     "aspectos da concepção, organização e gestão do trabalho, e de seus "
     "contextos social e ambiental, que têm potencial de causar dano "
     "psicológico ou físico. A ISO 45003:2021 define-os como fatores de "
     "natureza social, organizacional e gerencial que podem representar risco à "
     "saúde e à segurança, e os organiza em três grupos: os relativos à "
     "organização do trabalho, os relativos a fatores sociais no trabalho e os "
     "relativos ao ambiente e aos equipamentos (ISO, 2021).")
body(doc, "Instituições europeias de referência convergem nessa direção. A "
     "Agência Europeia para a Segurança e Saúde no Trabalho (EU-OSHA) trata os "
     "riscos psicossociais como decorrentes de deficiências na concepção, "
     "organização e gestão do trabalho, citando entre eles as cargas de "
     "trabalho excessivas, as exigências contraditórias, a falta de clareza "
     "sobre a função, a baixa participação nas decisões e o assédio. O Health "
     "and Safety Executive britânico, por sua vez, operacionaliza a avaliação "
     "por meio dos Management Standards, que organizam os fatores em seis "
     "categorias — demanda, controle, apoio, relacionamentos, função e mudança "
     "—, das quais este trabalho aproveita diretamente a estrutura de "
     "dimensões.")
body(doc, "No plano regulatório brasileiro, a NR-1 não fecha uma definição "
     "taxativa, mas remete a exemplos de fatores a considerar, alinhados aos "
     "documentos internacionais e ao guia publicado pelo Ministério do Trabalho "
     "e Emprego (BRASIL, 2024; BRASIL, 2025). A literatura nacional adverte, "
     "contudo, que a incorporação dos fatores psicossociais à gestão de SST "
     "enfrenta limitações quando reduzida a uma abordagem individualizante, "
     "dissociada da organização do trabalho — razão pela qual se insiste no "
     "caráter coletivo da avaliação. Adota-se, neste trabalho, a definição da "
     "ISO 45003 por ser a mais recente, específica e operacionalizável: ela "
     "nomeia categorias de fatores que se traduzem diretamente em dimensões de "
     "avaliação, o que orienta a construção do instrumento apresentado na seção "
     "3.")

h2(doc, "2.3 Modelos teóricos de estresse ocupacional")
body(doc, "A seleção das dimensões de avaliação apoia-se em modelos clássicos, "
     "cuja vigência é reafirmada pela literatura recente. O modelo "
     "demanda-controle, de Karasek (1979), relaciona o adoecimento à combinação "
     "de altas exigências psicológicas com baixa latitude de decisão "
     "(autonomia): as situações de maior risco são as de alta demanda e baixo "
     "controle. Extensões do modelo incorporam o apoio social como terceiro "
     "eixo protetor. O modelo esforço-recompensa, de Siegrist (1996), destaca o "
     "desequilíbrio entre o esforço despendido e as recompensas recebidas — "
     "salário, estima, reconhecimento e segurança — como fonte de estresse "
     "crônico.")
body(doc, "O modelo demanda-controle foi ampliado por Johnson e Hall (1988), "
     "que acrescentaram o apoio social como terceira dimensão, dando origem ao "
     "modelo demanda-controle-apoio: a situação de maior risco — denominada "
     "iso-strain — combina alta demanda, baixo controle e baixo apoio social. "
     "Essa tríade é particularmente adequada ao trabalho de campo em equipes, "
     "como as de distribuição de energia, em que o apoio dos colegas e da "
     "supervisão tem papel protetor direto.")
body(doc, "Esses modelos têm sustentação empírica no próprio setor elétrico "
     "brasileiro: Souza et al. (2010), ao estudarem 158 trabalhadores da "
     "manutenção de uma empresa de energia, encontraram maior prevalência de "
     "transtornos mentais comuns nos estratos de baixo controle (razão de "
     "prevalência de 1,34), alta demanda psicológica (2,31) e baixo apoio "
     "social (2,82), confirmando a pertinência do modelo "
     "demanda-controle-apoio para a atividade. Instrumentos multidimensionais, "
     "como o COPSOQ (KRISTENSEN et al., 2005), consolidam dimensões como "
     "demandas quantitativas, influência no trabalho, apoio social, clareza de "
     "papel e comportamentos ofensivos, e servem de referência para a seleção "
     "de itens. A convergência entre os modelos de Karasek, de Siegrist e os "
     "instrumentos multidimensionais sustenta a escolha das seis dimensões "
     "adotadas neste trabalho, detalhada na seção 3.")

h2(doc, "2.4 Síndrome de burnout")
body(doc, "Entre os desfechos associados à exposição prolongada a fatores "
     "psicossociais adversos, a síndrome de burnout ocupa posição central. "
     "Maslach e Jackson (1981) a caracterizaram por três dimensões: exaustão "
     "emocional, despersonalização (ou cinismo) e redução da realização "
     "pessoal. Revisões posteriores consolidaram o conceito como resposta "
     "prolongada a estressores crônicos, de natureza interpessoal e "
     "organizacional, no trabalho (MASLACH; SCHAUFELI; LEITER, 2001).")
body(doc, "O reconhecimento institucional do burnout avançou de forma decisiva. "
     "Na 11ª revisão da Classificação Internacional de Doenças (CID-11), a "
     "OMS incluiu o burn-out (código QD85) como fenômeno ocupacional, definido "
     "como síndrome resultante de estresse crônico no trabalho que não foi "
     "administrado com sucesso, caracterizado por exaustão, distanciamento "
     "mental do trabalho e redução da eficácia profissional (OMS, 2019). No "
     "Brasil, a adoção oficial da CID-11 a partir de 2025 reforçou o "
     "enquadramento do burnout como condição relacionada ao trabalho, com "
     "consequências para o nexo técnico e para a emissão de Comunicação de "
     "Acidente de Trabalho.")
body(doc, "A mensuração do burnout apoia-se historicamente no Maslach Burnout "
     "Inventory, que avalia as três dimensões do construto. Do ponto de vista "
     "preventivo, a literatura é convergente com a abordagem organizacional: "
     "como o burnout resulta de estressores crônicos do contexto de trabalho, "
     "as intervenções mais eficazes atuam sobre carga, controle, recompensa, "
     "comunidade, justiça e valores, e não apenas sobre a resiliência "
     "individual (MASLACH; SCHAUFELI; LEITER, 2001). Isso reforça a lógica da "
     "NR-1 de tratar os fatores psicossociais como perigos organizacionais a "
     "serem geridos, e não como fragilidades individuais.")
body(doc, "O quadro epidemiológico brasileiro reforça a urgência do tema. "
     "Estudo de série temporal com dados nacionais indicou tendência de "
     "crescimento das notificações de burnout entre 2014 e 2024, com maior "
     "concentração nas regiões Sudeste e Nordeste e pico em 2024, e "
     "predominância entre mulheres e na faixa de 35 a 49 anos (TREML et al., "
     "2025). Somados aos dados de afastamento do INSS compilados pela ANAMT "
     "(2026), esses achados evidenciam que a gestão dos fatores psicossociais "
     "prevista na NR-1 é também uma medida de prevenção de um desfecho já "
     "reconhecido como doença ocupacional.")
body(doc, "O reconhecimento do burnout como condição relacionada ao trabalho "
     "tem consequências jurídicas e previdenciárias relevantes. Uma vez "
     "estabelecido o nexo entre o adoecimento e as condições de trabalho, "
     "cabe a emissão de Comunicação de Acidente de Trabalho e podem incidir os "
     "efeitos de estabilidade e de responsabilização previstos na legislação. "
     "Para a organização, a ausência de gestão documentada dos fatores "
     "psicossociais — agora exigida pela NR-1 — enfraquece sua posição diante "
     "de eventual questionamento de nexo, o que soma o argumento de "
     "conformidade legal ao argumento de prevenção em saúde.")

h2(doc, "2.5 O Gerenciamento de Riscos Ocupacionais e o PGR na NR-1")
body(doc, "Desde a reformulação de 2020, a NR-1 estabelece o Gerenciamento de "
     "Riscos Ocupacionais como processo contínuo composto por levantamento "
     "preliminar de perigos, avaliação de riscos, classificação, implementação "
     "de medidas de prevenção e acompanhamento do controle. O PGR materializa "
     "esse processo em, no mínimo, dois documentos: o Inventário de Riscos "
     "Ocupacionais e o Plano de Ação. A norma exige que o inventário seja "
     "mantido atualizado e que o histórico de suas atualizações seja retido "
     "(BRASIL, 2024).")
body(doc, "Com a Portaria MTE nº 1.419/2024, os fatores de risco psicossocial "
     "passaram a integrar expressamente o rol de perigos a considerar no GRO. "
     "Na prática, isso equipara o tratamento documental dos fatores "
     "psicossociais ao dos demais agentes: precisam ser identificados, "
     "avaliados com método consistente, classificados, inscritos no IRO e "
     "vinculados a medidas de prevenção com responsáveis e prazos no Plano de "
     "Ação. O guia informativo do MTE sobre fatores de riscos psicossociais "
     "orienta que a avaliação seja coletiva e organizacional — e não um "
     "diagnóstico clínico individual — e que o anonimato da coleta é condição "
     "para a fidedignidade das respostas (BRASIL, 2025).")

h2(doc, "2.6 ISO 45003 e diretrizes internacionais")
body(doc, "A ISO 45003:2021 é a primeira norma internacional dedicada à gestão "
     "de riscos psicossociais, concebida como diretriz complementar à ISO "
     "45001:2018. Ela recomenda a integração da gestão psicossocial ao sistema "
     "de gestão de SST existente, em vez de um programa paralelo, e detalha "
     "exemplos de fatores em cada um dos três grupos que estabelece (ISO, "
     "2021). Alinhada a ela, a OMS propõe uma abordagem escalonada, priorizando "
     "medidas organizacionais (WHO, 2022). Essas referências sustentam duas "
     "decisões de projeto da plataforma: avaliar a organização, e não o "
     "indivíduo, e integrar o resultado diretamente ao PGR.")

h2(doc, "2.7 Instrumentos de avaliação psicossocial")
body(doc, "Entre os instrumentos consolidados destacam-se o COPSOQ, o Job "
     "Content Questionnaire, derivado do modelo de Karasek, e as ferramentas de "
     "indicadores de gestão do Health and Safety Executive britânico. São "
     "instrumentos robustos, porém extensos — versões médias do COPSOQ "
     "ultrapassam oitenta itens — e sua aplicação e interpretação exigem "
     "competência técnica pouco disponível em equipes de SST enxutas. Por isso, "
     "este trabalho optou por um instrumento próprio e enxuto, de dez itens, "
     "assumindo explicitamente o caráter de triagem: o objetivo é priorizar "
     "dimensões para aprofundamento e ação, e não produzir diagnóstico "
     "psicométrico definitivo. Essa opção e suas limitações são discutidas nas "
     "seções 3 e 4.")

h2(doc, "2.8 Riscos psicossociais no setor de distribuição de energia elétrica")
body(doc, "As atividades em redes de distribuição são regidas principalmente "
     "pela NR-10, que estabelece requisitos para serviços em instalações "
     "elétricas desenergizadas — com a sequência de desenergização, "
     "impedimento de reenergização, constatação de ausência de tensão e "
     "aterramento temporário — e energizadas, inclusive no Sistema Elétrico de "
     "Potência, exigindo trabalhadores autorizados e treinamento complementar "
     "específico (BRASIL, 2019). O trabalho em postes e estruturas sujeita-se "
     "ainda à NR-35. Do ponto de vista psicossocial, a literatura específica "
     "do setor evidencia a combinação de altas demandas físicas e cognitivas "
     "com fatores organizacionais adversos, associada a maior prevalência de "
     "transtornos mentais comuns (SOUZA et al., 2010). A pressão temporal no "
     "restabelecimento de fornecimento, o regime de turnos e sobreaviso, o "
     "trabalho a céu aberto e a convivência permanente com risco de acidente "
     "grave compõem um perfil em que a gestão psicossocial é também medida de "
     "prevenção de acidentes, dado que falhas de atenção têm consequência "
     "severa.")

h2(doc, "2.9 Transformação digital na gestão de SST")
body(doc, "A digitalização de processos de SST tem como benefícios documentados "
     "a padronização dos registros, a redução do tempo entre coleta e análise, "
     "a eliminação de transcrições manuais e a rastreabilidade exigida por "
     "auditorias e fiscalização. No contexto deste trabalho, a opção por "
     "tecnologias gratuitas e de baixa barreira técnica é uma decisão "
     "deliberada de projeto, voltada à replicabilidade da solução por "
     "profissionais de SST sem apoio de equipe de tecnologia da informação. "
     "Essa escolha dialoga com a hierarquia de medidas da NR-1: ao reduzir o "
     "custo de identificar e documentar, a ferramenta libera esforço para a "
     "etapa que efetivamente protege — a implementação de medidas de controle.")

h2(doc, "2.10 Estudos correlatos recentes")
body(doc, "A produção recente sobre o tema no Brasil intensificou-se a partir da "
     "publicação da Portaria MTE nº 1.419/2024, concentrando-se em três "
     "vertentes: a análise jurídico-normativa das obrigações decorrentes da "
     "nova NR-1 e sua incorporação ao PGR; a discussão dos limites e das "
     "condições para uma avaliação psicossocial fiel à organização do trabalho; "
     "e os estudos epidemiológicos sobre desfechos de saúde mental, entre os "
     "quais os de burnout (TREML et al., 2025; ANAMT, 2026). Observa-se, "
     "entretanto, escassez de trabalhos que proponham instrumentos e ferramentas "
     "operacionais de baixo custo voltados à realidade de organizações com "
     "equipes de SST reduzidas — lacuna que este trabalho pretende ajudar a "
     "preencher. No recorte setorial, a literatura específica sobre eletricidade "
     "permanece ancorada em estudos como o de Souza et al. (2010), o que "
     "evidencia a oportunidade de novas aplicações no setor sob a ótica da NR-1 "
     "atualizada.")

h2(doc, "2.11 Síntese da revisão")
body(doc, "A literatura recente e o marco normativo convergem em três pontos que "
     "orientam o desenvolvimento: (i) a avaliação psicossocial deve ser "
     "organizacional, anônima e integrada ao sistema de gestão; (ii) as "
     "dimensões a avaliar têm lastro teórico consolidado nos modelos "
     "demanda-controle-apoio (KARASEK, 1979; JOHNSON; HALL, 1988) e "
     "esforço-recompensa (SIEGRIST, 1996); e (iii) a inércia à conformidade é "
     "sobretudo operacional, o que abre espaço para uma solução digital de "
     "baixo custo. Esses três pontos são retomados na metodologia e na "
     "discussão.")
doc.add_page_break()

# ===================================================== 3 METODOLOGIA ==========
h1(doc, "3 METODOLOGIA")

h2(doc, "3.1 Caracterização da pesquisa")
body(doc, "Trata-se de pesquisa aplicada, de natureza tecnológica, conduzida na "
     "forma de desenvolvimento de artefato (a plataforma digital) seguido de "
     "estudo de caso com aplicação piloto. A abordagem é quali-quantitativa: "
     "quantitativa no tratamento das respostas do instrumento e qualitativa na "
     "análise da adequação do processo ao contexto organizacional. Quanto aos "
     "objetivos, é descritiva e propositiva. O estudo de caso foi conduzido em "
     "uma organização do setor de distribuição de energia elétrica, aqui "
     "caracterizada de forma a não permitir sua identificação: [PREENCHER: "
     "porte, região de atuação e natureza das atividades, SEM nome nem dados "
     "identificáveis].")

h2(doc, "3.2 Passo a passo da monografia")
body(doc, "O desenvolvimento seguiu as etapas descritas a seguir, encadeadas de "
     "modo que cada uma fornecesse subsídios à seguinte.")
bullet(doc, "Etapa 1 — Revisão bibliográfica e normativa: levantamento das "
       "definições, modelos e evidências recentes (seção 2) e sistematização "
       "dos requisitos da NR-1 e normas correlatas.", numbered=True)
bullet(doc, "Etapa 2 — Especificação do instrumento: definição das seis "
       "dimensões, redação dos dez itens em sentido positivo, escolha da escala "
       "Likert de cinco pontos e dos critérios de classificação de risco.",
       numbered=True)
bullet(doc, "Etapa 3 — Projeto da arquitetura: definição das premissas de custo "
       "zero, ausência de instalação e anonimato por construção, e escolha das "
       "tecnologias (frontend web autocontido e backend gratuito em nuvem).",
       numbered=True)
bullet(doc, "Etapa 4 — Desenvolvimento da plataforma: implementação do Módulo 1 "
       "(psicossocial) e dos módulos complementares (permissão de trabalho, "
       "ergonomia e acidentes), com painel e geração automática de texto para "
       "o IRO.", numbered=True)
bullet(doc, "Etapa 5 — Testes e validação técnica: verificação de "
       "funcionamento, integridade dos dados e usabilidade em dispositivos "
       "móveis.", numbered=True)
bullet(doc, "Etapa 6 — Aplicação piloto: distribuição do instrumento de forma "
       "anônima aos trabalhadores da organização estudada, com comunicação "
       "prévia sobre o caráter voluntário e o anonimato.", numbered=True)
bullet(doc, "Etapa 7 — Análise e discussão: consolidação dos resultados "
       "agregados, geração do texto do IRO e discussão à luz da literatura.",
       numbered=True)
bullet(doc, "Etapa 8 — Proposição de medidas: recomendação de plano de ação "
       "(5W2H) para as dimensões classificadas como de risco médio ou alto.",
       numbered=True)

h2(doc, "3.3 Aspectos éticos e anonimato")
body(doc, "A coleta foi concebida para não identificar indivíduos nem a "
     "organização. Não se coletam nome, matrícula, e-mail ou qualquer "
     "identificador pessoal; o único dado adicional é o setor, opcional, para "
     "permitir a estratificação mínima da análise. Esse desenho está alinhado à "
     "Lei nº 13.709/2018 (LGPD), que condiciona o tratamento de dados pessoais, "
     "e à orientação do guia do MTE de que a avaliação seja coletiva e anônima "
     "(BRASIL, 2025). Em conformidade com a orientação recebida, este documento "
     "não expõe o nome da organização, não apresenta fotografias de "
     "trabalhadores e não inclui dados que permitam identificação; as "
     "ilustrações limitam-se a gráficos agregados e a telas do sistema sem "
     "dados pessoais.")

h2(doc, "3.4 Arquitetura da plataforma")
body(doc, "A plataforma, denominada Psike, foi construída sobre três decisões de "
     "arquitetura: custo zero de operação, ausência de instalação e anonimato "
     "por construção. O frontend é um único arquivo executado no navegador, sem "
     "necessidade de instalação, o que permite o uso em qualquer computador ou "
     "celular. O backend é um serviço gratuito em nuvem que grava os registros "
     "em uma planilha eletrônica — uma aba por módulo —, servindo como "
     "repositório único de dados. O frontend detecta automaticamente o "
     "ambiente: quando o backend está configurado, os dados são gravados na "
     "nuvem e compartilhados entre dispositivos; na ausência de backend, a "
     "aplicação opera em modo local para demonstração.")
body(doc, "A plataforma foi organizada em quatro módulos sobre o mesmo "
     "repositório de dados, espelhando o Inventário de Riscos Ocupacionais: "
     "Módulo 1 — avaliação de riscos psicossociais (NR-1 e ISO 45003); Módulo 2 "
     "— permissão de trabalho digital para serviços em redes de distribuição "
     "(NR-10 e NR-35); Módulo 3 — Avaliação Ergonômica Preliminar (NR-17); e "
     "Módulo 4 — registro e investigação de acidentes e quase acidentes.")

h2(doc, "3.5 O instrumento de avaliação psicossocial")
body(doc, "O instrumento contém dez afirmativas redigidas em sentido positivo, "
     "distribuídas em seis dimensões: carga e ritmo de trabalho (dois itens), "
     "autonomia e controle (dois), clareza de papel (dois), apoio social e de "
     "liderança (dois), reconhecimento (um) e assédio e violência no trabalho "
     "(um). A escolha das dimensões deriva diretamente dos modelos revisados na "
     "seção 2: carga, autonomia e apoio decorrem do modelo "
     "demanda-controle-apoio; reconhecimento decorre do modelo "
     "esforço-recompensa; clareza de papel e assédio integram instrumentos "
     "multidimensionais como o COPSOQ. O respondente indica concordância em "
     "escala Likert de cinco pontos (1 = discordo totalmente; 5 = concordo "
     "totalmente). O texto integral do instrumento consta do Apêndice A.")

h2(doc, "3.6 Classificação de risco e geração do IRO")
body(doc, "Como as afirmativas são positivas, notas altas indicam condição "
     "favorável. A nota média de cada dimensão é classificada em três faixas: "
     "risco baixo (média maior ou igual a 3,8), risco médio (média entre 2,8 e "
     "3,8) e risco alto (média inferior a 2,8). Os pontos de corte foram "
     "definidos por julgamento técnico, privilegiando a sensibilidade — na "
     "dúvida, a dimensão é classificada na faixa mais conservadora — e devem "
     "ser recalibrados com a acumulação de séries históricas. A cada consulta "
     "ao painel, a plataforma consolida as respostas, calcula as médias, aplica "
     "a classificação e redige automaticamente a seção de fatores psicossociais "
     "do IRO, contendo base normativa, método, número de respondentes, "
     "resultado por dimensão, recomendação de plano de ação para as dimensões "
     "críticas e recomendação de reavaliação periódica. O texto é editável "
     "antes da inserção no PGR, preservando a responsabilidade técnica do "
     "profissional que o subscreve.")

h2(doc, "3.7 Módulos complementares")
body(doc, "Módulo 2 — Permissão de trabalho digital (NR-10/NR-35): o executante "
     "seleciona o tipo de serviço — rede desenergizada (linha morta), rede "
     "energizada ou Sistema Elétrico de Potência (linha viva) e trabalho em "
     "altura em postes e estruturas — e a plataforma apresenta o checklist "
     "correspondente. Para rede desenergizada, o checklist segue a sequência de "
     "desenergização da NR-10. A liberação exige checklist completo, captura de "
     "geolocalização e assinatura digital em tela, com o botão de liberação "
     "bloqueado até que todas as condições sejam atendidas.")
body(doc, "Módulo 3 — Avaliação Ergonômica Preliminar (NR-17): formulário "
     "estruturado em cinco blocos de fatores, avaliados em escala de três "
     "pontos; a plataforma classifica cada bloco e gera parecer automático, "
     "indicando a necessidade de aprofundamento em Análise Ergonômica do "
     "Trabalho quando algum bloco resulta inadequado.")
body(doc, "Módulo 4 — Registro e investigação de acidentes e quase acidentes "
     "(NR-1): registro rápido em campo, com classificação da ocorrência, "
     "descrição, análise de causas em três níveis (imediatas, subjacentes e "
     "básicas) e ações decorrentes, além de indicadores. O módulo reforça que a "
     "emissão da Comunicação de Acidente de Trabalho é obrigação legal do "
     "empregador e não é substituída pelo registro na plataforma.")

h2(doc, "3.8 Aplicação piloto")
body(doc, "A aplicação piloto foi conduzida na organização estudada no período "
     "de [PREENCHER: período], abrangendo [PREENCHER: número] trabalhadores dos "
     "setores de [PREENCHER: setores]. O link da plataforma foi distribuído por "
     "[PREENCHER: canal], precedido de comunicação sobre o caráter anônimo e "
     "voluntário da participação. [PREENCHER: registrar eventuais aprovações "
     "internas e a forma de sensibilização, sem identificar a organização].")

h2(doc, "3.9 Tratamento e análise dos dados")
body(doc, "As respostas foram tratadas de forma agregada. Para cada dimensão, "
     "calcula-se a média aritmética das respostas dos itens que a compõem, "
     "considerando todos os respondentes; a média é então enquadrada em uma das "
     "três faixas de risco definidas na seção 3.6. O painel apresenta, além das "
     "médias e faixas por dimensão, indicadores consolidados — número de "
     "respondentes e contagem de dimensões em cada faixa — que subsidiam a "
     "priorização das ações. Quando a amostra por setor é suficiente, é possível "
     "estratificar os resultados para localizar os focos de risco; quando não "
     "é, os resultados são apresentados apenas de forma global, para não "
     "comprometer o anonimato. Nenhuma análise busca ou permite a identificação "
     "de respondentes individuais.")
doc.add_page_break()

# ============================================ 4 RESULTADOS E DISCUSSÃO =========
h1(doc, "4 RESULTADOS E DISCUSSÃO")
body(doc, "Esta seção integra a apresentação dos resultados e sua discussão à "
     "luz da literatura e do marco normativo revisados na seção 2, conforme "
     "orientação metodológica adotada.")

h2(doc, "4.1 A plataforma desenvolvida")
body(doc, "A plataforma foi implementada integralmente e encontra-se "
     "operacional, com os quatro módulos em uso. O fluxo do Módulo 1 — da "
     "resposta anônima no celular ao texto pronto para o IRO — ocorre sem "
     "qualquer etapa manual de tabulação. O custo de operação é nulo, uma vez "
     "que tanto a hospedagem do frontend quanto o backend utilizam camadas "
     "gratuitas. A Figura 1 apresenta as telas principais do sistema. "
     "[PREENCHER: inserir Figura 1 com telas do sistema, SEM dados pessoais.]")

h2(doc, "4.2 Resultados da aplicação piloto")
body(doc, "[PREENCHER: apresentar os resultados reais — número de respondentes, "
     "taxa de adesão, nota média e classificação por dimensão, e comparação "
     "entre setores se houver amostra suficiente. Inserir a tabela de "
     "resultados por dimensão e o gráfico agregado do painel. NÃO inventar "
     "dados e NÃO identificar a organização.]")
body(doc, "[PREENCHER: transcrever o texto do IRO gerado pela plataforma para a "
     "aplicação real, como evidência do produto final do processo.]")

h2(doc, "4.3 Discussão à luz da literatura")
body(doc, "Os achados [PREENCHER: ajustar conforme os dados reais] são "
     "coerentes com o padrão descrito pela literatura do setor elétrico, em que "
     "as dimensões de demanda, controle e apoio social concentram o maior "
     "potencial de risco (SOUZA et al., 2010). Caso as dimensões de carga e "
     "apoio despontem como críticas, o resultado dialoga diretamente com o "
     "modelo demanda-controle-apoio (KARASEK, 1979) e reforça a pertinência de "
     "medidas organizacionais, como redimensionamento de equipes e revisão de "
     "escalas, em detrimento de intervenções centradas apenas no indivíduo, "
     "conforme recomenda a OMS (2022). A eventual criticidade da dimensão "
     "reconhecimento remete ao modelo esforço-recompensa (SIEGRIST, 1996) e ao "
     "risco de burnout, cujo componente de redução da realização pessoal se "
     "associa à percepção de baixa recompensa (MASLACH; SCHAUFELI; LEITER, "
     "2001).")
body(doc, "A discussão evidencia, ainda, o principal argumento do trabalho: a "
     "barreira à conformidade com a nova NR-1 é sobretudo operacional, e não "
     "conceitual. O ciclo identificar–avaliar–documentar–agir é bem descrito na "
     "norma e nas referências, porém custoso quando executado manualmente. A "
     "digitalização de ponta a ponta reduziu esse custo a praticamente zero na "
     "organização estudada, ao mesmo tempo em que melhorou os atributos "
     "exigidos pela norma: anonimato, consistência da classificação e "
     "rastreabilidade dos registros.")

h2(doc, "4.4 Integração ao PGR e plano de ação")
body(doc, "O texto gerado pela plataforma foi estruturado para inserção direta "
     "na seção de fatores psicossociais do IRO. Para as dimensões classificadas "
     "como de risco médio ou alto, recomenda-se plano de ação no formato 5W2H, "
     "com medidas de natureza organizacional priorizadas, em consonância com a "
     "hierarquia de controles da NR-1 e com as diretrizes da OMS. [PREENCHER: "
     "descrever o plano de ação efetivamente elaborado a partir dos resultados "
     "do piloto, sem identificar a organização.]")

h2(doc, "4.5 Limitações")
bullet(doc, "Instrumento de triagem: com dez itens, prioriza viabilidade e "
       "adesão; não substitui instrumentos psicométricos validados quando se "
       "exigir diagnóstico aprofundado, e não passou por validação estatística "
       "formal;")
bullet(doc, "Pontos de corte definidos por julgamento técnico, a calibrar com "
       "séries históricas;")
bullet(doc, "Ausência de autenticação e de segregação de dados por "
       "organização, aceitável no piloto, mas requisito para uso comercial;")
bullet(doc, "Limites das camadas gratuitas de nuvem quanto a volume de dados e "
       "requisições;")
bullet(doc, "Amostra do piloto restrita a uma organização, o que limita a "
       "generalização direta dos resultados.")
doc.add_page_break()

# ============================================ 5 CONCLUSÃO =====================
h1(doc, "5 CONCLUSÃO")
body(doc, "O objetivo geral — desenvolver e aplicar uma plataforma digital para "
     "a gestão de fatores de risco psicossocial conforme a NR-1, integrada à "
     "gestão de riscos de uma organização do setor de distribuição de energia — "
     "foi [PREENCHER: atingido / parcialmente atingido]. A plataforma foi "
     "desenvolvida e está operacional, com coleta anônima, classificação "
     "automática por dimensão e geração automática do texto para o IRO, e foi "
     "aplicada em caráter piloto. [PREENCHER: uma frase objetiva sobre o "
     "principal resultado do piloto e se ele confirmou ou não a hipótese de "
     "viabilidade.]")
body(doc, "Como trabalhos futuros, sugerem-se a validação psicométrica do "
     "instrumento, a calibração dos pontos de corte com séries históricas, a "
     "implementação de autenticação e segregação por organização, a migração "
     "para banco de dados dedicado e o acompanhamento longitudinal da eficácia "
     "das medidas de controle implantadas.")
doc.add_page_break()

# ============================================ REFERÊNCIAS =====================
h1(doc, "REFERÊNCIAS")
REFS = [
    "ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. ABNT NBR ISO 45001:2018 — "
    "Sistemas de gestão de saúde e segurança ocupacional: requisitos com "
    "orientações para uso. Rio de Janeiro: ABNT, 2018.",
    "ASSOCIAÇÃO NACIONAL DE MEDICINA DO TRABALHO (ANAMT). Levantamento com "
    "dados oficiais do INSS revela crescimento dos afastamentos decorrentes de "
    "problemas de saúde mental entre 2023 e 2025. 2026. Disponível em: "
    "[PREENCHER: URL e data de acesso].",
    "AGÊNCIA EUROPEIA PARA A SEGURANÇA E SAÚDE NO TRABALHO (EU-OSHA). "
    "Psychosocial risks and stress at work. Bilbao: EU-OSHA, [PREENCHER: ano "
    "da publicação consultada].",
    "BRASIL. Lei nº 13.709, de 14 de agosto de 2018. Lei Geral de Proteção de "
    "Dados Pessoais (LGPD). Diário Oficial da União, Brasília, DF, 15 ago. 2018.",
    "HEALTH AND SAFETY EXECUTIVE (HSE). Management Standards for work-related "
    "stress. Sudbury: HSE, [PREENCHER: ano da versão consultada].",
    "JOHNSON, J. V.; HALL, E. M. Job strain, work place social support, and "
    "cardiovascular disease: a cross-sectional study of a random sample of the "
    "Swedish working population. American Journal of Public Health, v. 78, "
    "n. 10, p. 1336-1342, 1988.",
    "BRASIL. Ministério do Trabalho e Emprego. Norma Regulamentadora nº 1 "
    "(NR-1): disposições gerais e gerenciamento de riscos ocupacionais. Redação "
    "dada pela Portaria MTE nº 1.419, de 27 de agosto de 2024. Brasília, DF: "
    "MTE, 2024.",
    "BRASIL. Ministério do Trabalho e Emprego. Portaria MTE nº 765, de 2025 — "
    "define a exigibilidade da gestão de riscos psicossociais na NR-1 a partir "
    "de 26 de maio de 2026. Brasília, DF: MTE, 2025. [PREENCHER: conferir "
    "número/data exatos].",
    "BRASIL. Ministério do Trabalho e Emprego. Norma Regulamentadora nº 10 "
    "(NR-10): segurança em instalações e serviços em eletricidade. Brasília, "
    "DF: MTE, 2019.",
    "BRASIL. Ministério do Trabalho e Emprego. Norma Regulamentadora nº 17 "
    "(NR-17): ergonomia. Redação dada pela Portaria MTP nº 423, de 7 de outubro "
    "de 2021. Brasília, DF: MTE, 2021.",
    "BRASIL. Ministério do Trabalho e Emprego. Norma Regulamentadora nº 35 "
    "(NR-35): trabalho em altura. Brasília, DF: MTE, [PREENCHER: ano da redação "
    "consultada].",
    "BRASIL. Ministério do Trabalho e Emprego. Guia de informações sobre os "
    "fatores de riscos psicossociais relacionados ao trabalho. Brasília, DF: "
    "MTE, 2025. [PREENCHER: conferir título e ano exatos].",
    "INTERNATIONAL ORGANIZATION FOR STANDARDIZATION. ISO 45003:2021 — "
    "Occupational health and safety management — Psychological health and "
    "safety at work: guidelines for managing psychosocial risks. Geneva: ISO, "
    "2021.",
    "KARASEK, R. A. Job demands, job decision latitude, and mental strain: "
    "implications for job redesign. Administrative Science Quarterly, v. 24, "
    "n. 2, p. 285-308, 1979.",
    "KRISTENSEN, T. S.; HANNERZ, H.; HØGH, A.; BORG, V. The Copenhagen "
    "Psychosocial Questionnaire: a tool for the assessment and improvement of "
    "the psychosocial work environment. Scandinavian Journal of Work, "
    "Environment & Health, v. 31, n. 6, p. 438-449, 2005.",
    "LEKA, S.; JAIN, A. Health impact of psychosocial hazards at work: an "
    "overview. Geneva: World Health Organization, 2010.",
    "MASLACH, C.; JACKSON, S. E. The measurement of experienced burnout. "
    "Journal of Occupational Behavior, v. 2, n. 2, p. 99-113, 1981.",
    "MASLACH, C.; SCHAUFELI, W. B.; LEITER, M. P. Job burnout. Annual Review "
    "of Psychology, v. 52, p. 397-422, 2001.",
    "ORGANIZAÇÃO MUNDIAL DA SAÚDE (OMS). Classificação Estatística "
    "Internacional de Doenças e Problemas Relacionados à Saúde (CID-11): "
    "burn-out (QD85). Genebra: OMS, 2019.",
    "SIEGRIST, J. Adverse health effects of high-effort/low-reward conditions. "
    "Journal of Occupational Health Psychology, v. 1, n. 1, p. 27-41, 1996.",
    "SOUZA, S. F.; CARVALHO, F. M.; ARAÚJO, T. M.; PORTO, L. A. Fatores "
    "psicossociais do trabalho e transtornos mentais comuns em eletricitários. "
    "Revista de Saúde Pública, v. 44, n. 4, p. 710-717, 2010.",
    "TREML, M. F. Q. et al. Burnout syndrome in Brazil (2014–2024): regional "
    "variations and temporal trends in an epidemiological study. Revista "
    "Brasileira de Medicina do Trabalho, 2025. [PREENCHER: conferir volume, "
    "número, páginas e DOI].",
    "UNIVERSIDADE DE SÃO PAULO. Agência de Bibliotecas e Coleções Digitais. "
    "Diretrizes para apresentação de dissertações e teses da USP: parte I "
    "(ABNT). 5. ed. São Paulo: ABCD/USP, 2024.",
    "WORLD HEALTH ORGANIZATION (WHO). WHO guidelines on mental health at work. "
    "Geneva: WHO, 2022.",
]
for r in REFS:
    ref(doc, r)

# ============================================ APÊNDICE A =====================
doc.add_page_break()
h1(doc, "APÊNDICE A — INSTRUMENTO DE AVALIAÇÃO APLICADO")
body(doc, "Itens do questionário (escala: 1 = discordo totalmente a 5 = concordo "
     "totalmente; respostas anônimas; único dado adicional: setor, opcional).",
     indent=False)
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
    bullet(doc, f"{i}. ({dim}) {txt}")

doc.add_page_break()
h1(doc, "APÊNDICE B — ARQUITETURA E MÓDULOS DA PLATAFORMA")
body(doc, "A plataforma Psike é composta por um frontend web autocontido e um "
     "backend gratuito em nuvem, com quatro módulos sobre o mesmo repositório "
     "de dados. A Figura B.1 sintetiza a arquitetura. [PREENCHER: inserir "
     "diagrama de arquitetura e telas dos quatro módulos, sem dados pessoais.]")
body(doc, "Quadro-resumo dos módulos: Módulo 1 — riscos psicossociais "
     "(NR-1/ISO 45003); Módulo 2 — permissão de trabalho (NR-10/NR-35); "
     "Módulo 3 — Avaliação Ergonômica Preliminar (NR-17); Módulo 4 — acidentes "
     "e quase acidentes (NR-1).", indent=False)

doc.save(OUT)
print(f"OK: {OUT} gerado.")
