# -*- coding: utf-8 -*-
"""
Gerador do TCC — Sistema digital de registro e investigação de acidentes e
quase acidentes pelo método da árvore de causas (plataforma Psike, Módulo 4).
Produz TCC_Acidentes_ArvoreCausas_MODELO.docx com formatação ABNT
(A4, Arial 12, espaçamento 1,5, margens 3/3/2/2, recuo de 1,25 cm).

Tema repivotado em 24/08/2026 a pedido do autor (foco 100% engenharia de
segurança), mantendo as exigências da supervisão (reunião 03/07/2026):
- 40 a 80 folhas; revisão de literatura com definições e refs recentes;
- burnout citado (como fator humano contribuinte de acidentes);
- passo a passo da monografia; resultados+discussão integrados;
- conclusão curta; anonimato total (agradecimento apenas à CERPRO).

Uso:  python3 gerar_tcc.py     Requer: pip install python-docx
Trechos [PREENCHER: ...] dependem de dados reais — NÃO inventar e NÃO
identificar empresa/pessoas.
"""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = "TCC_Acidentes_ArvoreCausas_MODELO.docx"

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

def ref(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_after = Pt(12)
    _runs_with_marks(p, text)
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
body(doc, "Modelo do TCC estruturado conforme as Diretrizes USP/ABCD (5ª ed., "
     "2024) e as orientações da supervisão. Trechos em amarelo "
     "[PREENCHER: ...] dependem de dados reais e não devem ser inventados.",
     indent=False)
body(doc, "ANONIMATO OBRIGATÓRIO: não inserir nome da empresa, fotografias de "
     "pessoas nem dados identificáveis (número de CAT, nomes, matrículas, "
     "localidades específicas). Referir-se sempre à “organização estudada”. Os "
     "casos de acidente analisados devem ser descritos de forma anonimizada "
     "(função genérica, sem data exata nem local). Agradecimentos: apenas à "
     "CERPRO.", indent=False)
body(doc, "Antes de entregar: (1) preencher campos amarelos; (2) atualizar o "
     "Sumário; (3) conferir a extensão (meta 40–80 folhas); (4) apagar esta "
     "página.", indent=False)
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
center(doc, "SISTEMA DIGITAL DE REGISTRO E INVESTIGAÇÃO DE ACIDENTES E QUASE "
       "ACIDENTES PELO MÉTODO DA ÁRVORE DE CAUSAS:", bold=True, size=14)
center(doc, "desenvolvimento e aplicação na gestão de riscos ocupacionais "
       "(NR-1) em serviços de distribuição de energia elétrica",
       bold=True, size=14, space_after=48)
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
center(doc, "SISTEMA DIGITAL DE REGISTRO E INVESTIGAÇÃO DE ACIDENTES E QUASE "
       "ACIDENTES PELO MÉTODO DA ÁRVORE DE CAUSAS:", bold=True, size=13)
center(doc, "desenvolvimento e aplicação na gestão de riscos ocupacionais "
       "(NR-1) em serviços de distribuição de energia elétrica",
       bold=True, size=13, space_after=36)
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
rr = nat2.add_run("[PREENCHER: nome da supervisora]")
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
body(doc, "À organização que viabilizou a aplicação deste estudo, aqui "
     "preservada em anonimato, e aos profissionais que participaram das "
     "investigações de campo.")
body(doc, "À minha família, pelo apoio durante a especialização.")
doc.add_page_break()

# ===== Resumo =================================================================
center(doc, "RESUMO", bold=True, space_after=18)
body(doc, "A investigação e a análise de acidentes e quase acidentes integram as "
     "exigências do Gerenciamento de Riscos Ocupacionais (GRO) da Norma "
     "Regulamentadora nº 1 (NR-1), que determina ao empregador analisar os "
     "acidentes e as doenças relacionadas ao trabalho e utilizar seus "
     "resultados na revisão do Programa de Gerenciamento de Riscos (PGR). Na "
     "prática, contudo, o registro de ocorrências em muitas organizações ainda "
     "é manual, tardio e centrado na conduta do acidentado — abordagem que a "
     "literatura associa à recorrência dos eventos. Este trabalho teve por "
     "objetivo desenvolver e aplicar um sistema digital de registro e "
     "investigação de acidentes e quase acidentes estruturado no método da "
     "árvore de causas, aplicado a uma organização do setor de distribuição de "
     "energia elétrica. Adotou-se pesquisa aplicada, de natureza tecnológica, "
     "na forma de desenvolvimento de artefato seguido de estudo de caso, com "
     "análise documental retrospectiva de ocorrências anonimizadas e aplicação "
     "piloto em campo. O sistema — módulo central da plataforma Psike — permite "
     "o registro imediato da ocorrência no local, classifica o evento, conduz o "
     "investigador pela análise de causas em três níveis (imediatas, "
     "subjacentes e básicas), vincula ações corretivas e preventivas e calcula "
     "indicadores reativos e proativos, incluindo a razão entre quase acidentes "
     "e acidentes. Módulos complementares digitalizam, sobre a mesma base de "
     "dados, a permissão de trabalho (NR-10 e NR-35), a avaliação de fatores "
     "psicossociais (NR-1) e a Avaliação Ergonômica Preliminar (NR-17). A "
     "aplicação foi conduzida em [PREENCHER: período e escopo, sem identificar "
     "a organização]. Os resultados indicam [PREENCHER: síntese dos resultados "
     "reais]. Conclui-se que a digitalização do ciclo "
     "registro–investigação–ação reduz a subnotificação de quase acidentes, "
     "desloca a análise da culpabilização individual para os fatores "
     "organizacionais e melhora a rastreabilidade exigida pelo PGR.",
     indent=False)
body(doc, "Palavras-chave: acidentes do trabalho; quase acidentes; árvore de "
     "causas; investigação de acidentes; NR-1; distribuição de energia "
     "elétrica.", indent=False)
doc.add_page_break()

# ===== Abstract ===============================================================
center(doc, "ABSTRACT", bold=True, space_after=18)
body(doc, "The investigation and analysis of accidents and near misses are "
     "requirements of the Occupational Risk Management (GRO) framework of "
     "Brazilian Regulatory Standard No. 1 (NR-1), which obliges employers to "
     "analyze work-related accidents and diseases and to use the results in "
     "the review of the Risk Management Program (PGR). In practice, however, "
     "incident recording in many organizations remains manual, late, and "
     "centered on the injured worker's conduct — an approach the literature "
     "associates with event recurrence. This study aimed to develop and apply "
     "a digital system for recording and investigating accidents and near "
     "misses structured on the causal tree method, applied to an electric "
     "power distribution organization. Applied, technological research was "
     "adopted: artifact development followed by a case study, with "
     "retrospective documentary analysis of anonymized occurrences and a field "
     "pilot. The system — the core module of the Psike platform — enables "
     "immediate on-site recording, classifies the event, guides the "
     "investigator through three levels of causal analysis (immediate, "
     "underlying and basic causes), links corrective and preventive actions, "
     "and computes reactive and proactive indicators, including the near "
     "miss-to-accident ratio. Complementary modules digitize, on the same data "
     "repository, work permits (NR-10 and NR-35), psychosocial risk assessment "
     "(NR-1) and the Preliminary Ergonomic Assessment (NR-17). The application "
     "was conducted in [PREENCHER: period and scope]. Results indicate "
     "[PREENCHER: summary]. The study concludes that digitizing the "
     "record–investigate–act cycle reduces near-miss underreporting, shifts "
     "analysis from individual blame to organizational factors, and improves "
     "the traceability required by the PGR.", indent=False)
body(doc, "Keywords: occupational accidents; near misses; causal tree method; "
     "accident investigation; NR-1; electric power distribution.", indent=False)
doc.add_page_break()

# ===== Lista de siglas ========================================================
center(doc, "LISTA DE SIGLAS", bold=True, space_after=18)
SIGLAS = [
    ("ADC", "Árvore de Causas (método de investigação de acidentes)"),
    ("AEP", "Avaliação Ergonômica Preliminar"),
    ("APR", "Análise Preliminar de Risco"),
    ("CAT", "Comunicação de Acidente de Trabalho"),
    ("CID-11", "Classificação Internacional de Doenças, 11ª revisão"),
    ("CLT", "Consolidação das Leis do Trabalho"),
    ("GRO", "Gerenciamento de Riscos Ocupacionais"),
    ("INRS", "Institut National de Recherche et de Sécurité (França)"),
    ("INSS", "Instituto Nacional do Seguro Social"),
    ("IRO", "Inventário de Riscos Ocupacionais"),
    ("LGPD", "Lei Geral de Proteção de Dados Pessoais"),
    ("MTE", "Ministério do Trabalho e Emprego"),
    ("NBR", "Norma Brasileira (ABNT)"),
    ("NR", "Norma Regulamentadora"),
    ("PGR", "Programa de Gerenciamento de Riscos"),
    ("SEP", "Sistema Elétrico de Potência"),
    ("SST", "Segurança e Saúde do Trabalho"),
    ("TF", "Taxa de Frequência de acidentes"),
    ("TG", "Taxa de Gravidade de acidentes"),
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
body(doc, "Os acidentes do trabalho permanecem entre os principais problemas de "
     "saúde pública e de gestão empresarial no Brasil. Os registros oficiais "
     "consolidados no Anuário Estatístico de Acidentes do Trabalho contabilizam "
     "centenas de milhares de ocorrências por ano [PREENCHER: inserir o número "
     "do AEAT mais recente e o ano-base, com a fonte], e a literatura é "
     "unânime em apontar a subnotificação — sobretudo dos eventos sem "
     "afastamento e dos quase acidentes — como limitação estrutural dessas "
     "estatísticas.")
body(doc, "O risco elétrico agrava esse quadro. O anuário estatístico da "
     "Abracopel registrou, no ano-base 2023, 2.089 acidentes de origem "
     "elétrica no país, dos quais 781 resultaram em óbito — letalidade que "
     "supera um terço das ocorrências (ABRACOPEL, 2024). No recorte das redes "
     "de distribuição, o levantamento da Abradee apontou 250 mortes "
     "relacionadas à rede elétrica em 2023 (ABRADEE, 2024). Para as equipes "
     "que constroem e mantêm essas redes — trabalhando em condutores "
     "energizados ou desenergizados e em altura —, o acidente raramente é "
     "leve: a combinação de energia elétrica e queda de altura produz lesões "
     "graves e fatais, o que torna a prevenção baseada em aprendizado com "
     "ocorrências uma prioridade de engenharia.")
body(doc, "A NR-1, na redação vigente, insere a análise de acidentes no núcleo "
     "do Gerenciamento de Riscos Ocupacionais: o empregador deve analisar os "
     "acidentes e as doenças relacionadas ao trabalho, identificar suas causas "
     "e utilizar os resultados na revisão do levantamento de perigos e do "
     "PGR (BRASIL, 2024). A obrigação, portanto, não termina no socorro à "
     "vítima nem na emissão da Comunicação de Acidente de Trabalho: exige "
     "investigação com método, registro rastreável e realimentação do "
     "inventário de riscos e do plano de ação.")
body(doc, "Na prática de muitas organizações, porém, esse ciclo é frágil: o "
     "registro é feito em papel, horas ou dias após o evento; a investigação, "
     "quando ocorre, limita-se a atribuir o acidente a “ato inseguro” do "
     "trabalhador; os quase acidentes — eventos sem lesão que sinalizam as "
     "mesmas falhas sistêmicas — raramente são reportados; e as ações "
     "corretivas não são acompanhadas até a conclusão. A literatura de análise "
     "de acidentes denomina essa prática de abordagem tradicional e "
     "culpabilizadora, e a associa diretamente à recorrência dos eventos "
     "(BINDER; ALMEIDA, 1997; ALMEIDA, 2006).")

h2(doc, "1.2 Problema de pesquisa")
body(doc, "Coloca-se o seguinte problema: como estruturar, em uma organização de "
     "distribuição de energia elétrica com equipe de SST enxuta, um processo "
     "de registro e investigação de acidentes e quase acidentes que seja "
     "imediato, metodologicamente consistente — orientado aos fatores "
     "organizacionais, e não à culpabilização individual —, rastreável para "
     "fins do PGR e de custo compatível com a realidade da organização?")

h2(doc, "1.3 Justificativa")
body(doc, "Três razões justificam o trabalho. Primeiro, a relevância do setor: "
     "a gravidade típica dos acidentes de origem elétrica (ABRACOPEL, 2024) "
     "faz de cada ocorrência — e de cada quase acidente — uma oportunidade de "
     "aprendizado cujo desperdício tem custo potencial em vidas. Segundo, a "
     "exigência normativa: a NR-1 obriga a análise de acidentes como insumo do "
     "GRO, e a fiscalização e as auditorias demandam registros rastreáveis. "
     "Terceiro, a lacuna prática: soluções comerciais de gestão de ocorrências "
     "têm custo incompatível com organizações menores, e o método da árvore de "
     "causas — consolidado na literatura brasileira desde a década de 1990 — "
     "carece de ferramentas digitais acessíveis que o levem ao campo. Este "
     "trabalho demonstra a viabilidade de uma solução de custo praticamente "
     "nulo que integra registro imediato, investigação estruturada e "
     "indicadores, e que se articula com os demais processos de SST da "
     "organização sobre uma mesma base de dados.")

h2(doc, "1.4 Objetivos")
h3(doc, "1.4.1 Objetivo geral")
body(doc, "Desenvolver e aplicar um sistema digital de registro e investigação "
     "de acidentes e quase acidentes, estruturado no método da árvore de "
     "causas e integrado ao Gerenciamento de Riscos Ocupacionais (NR-1), em "
     "uma organização do setor de distribuição de energia elétrica.")
h3(doc, "1.4.2 Objetivos específicos")
bullet(doc, "Sistematizar os requisitos da NR-1 aplicáveis à análise de "
       "acidentes e sua articulação com o PGR, a CAT e a NBR 14280;")
bullet(doc, "Revisar a literatura sobre modelos de causalidade de acidentes e "
       "sobre o método da árvore de causas, incluindo os fatores humanos e "
       "organizacionais contribuintes;")
bullet(doc, "Especificar e implementar o sistema digital: registro imediato em "
       "campo, classificação da ocorrência, análise de causas em três níveis, "
       "gestão de ações e indicadores;")
bullet(doc, "Integrar o sistema aos módulos complementares da plataforma "
       "(permissão de trabalho NR-10/NR-35, fatores psicossociais NR-1 e "
       "AEP NR-17), sobre a mesma base de dados;")
bullet(doc, "Aplicar o sistema em estudo de caso — análise documental "
       "retrospectiva de ocorrências anonimizadas e piloto de campo — e "
       "discutir resultados, limitações e condições de generalização.")

h2(doc, "1.5 Estrutura do trabalho")
body(doc, "Além desta introdução, o trabalho organiza-se em quatro seções: a "
     "seção 2 apresenta a revisão de literatura e o marco normativo; a seção 3 "
     "descreve a metodologia, com o passo a passo da monografia e a "
     "especificação do sistema; a seção 4 apresenta e discute os resultados de "
     "forma integrada; a seção 5 traz a conclusão. Seguem-se as referências e "
     "os apêndices.")
doc.add_page_break()

# ============================================== 2 REVISÃO DE LITERATURA ========
h1(doc, "2 REVISÃO DE LITERATURA E MARCO NORMATIVO")
body(doc, "Esta seção reúne as definições, os modelos de causalidade e as "
     "referências normativas que fundamentam o sistema desenvolvido. "
     "Privilegiaram-se fontes recentes para o estado atual do tema, "
     "complementadas pelas obras seminais dos modelos de análise de acidentes.")

h2(doc, "2.1 Acidente, incidente e quase acidente: definições")
body(doc, "A Lei nº 8.213/1991 define acidente do trabalho, para fins "
     "previdenciários, como o que ocorre pelo exercício do trabalho a serviço "
     "da empresa, provocando lesão corporal ou perturbação funcional que cause "
     "morte, perda ou redução da capacidade para o trabalho (BRASIL, 1991). A "
     "NBR 14280 adota conceito prevencionista mais amplo — ocorrência "
     "imprevista e indesejável, instantânea ou não, relacionada com o "
     "exercício do trabalho, de que resulte ou possa resultar lesão pessoal — "
     "e padroniza o cadastro e as estatísticas de acidentes, incluindo as "
     "taxas de frequência e de gravidade (ABNT, 2001).")
body(doc, "O quase acidente (near miss) é o evento que, por circunstâncias "
     "fortuitas, não produziu lesão nem dano, mas cuja dinâmica é a mesma do "
     "acidente. A ISO 45001 o incorpora ao conceito de “incidente”, exigindo "
     "seu reporte e investigação (ABNT, 2018). Adota-se neste trabalho a "
     "tríade operacional: acidente com afastamento, acidente sem afastamento e "
     "quase acidente — à qual o sistema acrescenta o registro de condições "
     "inseguras observadas, na fronteira entre inspeção e ocorrência.")

h2(doc, "2.2 Modelos de causalidade de acidentes")
body(doc, "A forma de investigar decorre do modelo causal adotado — por isso a "
     "revisão dos modelos precede a escolha do método. A tradição inaugurada "
     "por Heinrich (1931) representou o acidente como sequência linear de "
     "fatores (teoria do dominó), na qual bastaria remover uma peça — "
     "tipicamente o “ato inseguro” — para interromper a cadeia. Do mesmo autor "
     "vem a pirâmide que relaciona, estatisticamente, grandes números de "
     "incidentes menores a cada lesão grave; Bird e Germain (1985) a "
     "atualizaram e estenderam aos danos materiais, fundamentando a prática "
     "moderna de reportar e tratar quase acidentes como matéria-prima da "
     "prevenção.")
body(doc, "Os modelos lineares, contudo, mostraram-se insuficientes para "
     "explicar acidentes em sistemas complexos. Reason (1990; 1997) propôs a "
     "distinção entre falhas ativas — atos na ponta operacional — e condições "
     "latentes — decisões de projeto, gestão e organização que permanecem "
     "adormecidas até se alinharem às falhas ativas, na representação "
     "conhecida como modelo do queijo suíço. A consequência prática é direta: "
     "investigar apenas a conduta do trabalhador deixa intactas as condições "
     "latentes, que voltarão a produzir eventos. Desenvolvimentos recentes, "
     "como a perspectiva Safety-II de Hollnagel (2014), deslocam o foco do "
     "que falha para a variabilidade normal do trabalho real, reforçando o "
     "valor de aprender também com o trabalho cotidiano — o que dá suporte "
     "conceitual ao registro de quase acidentes e condições inseguras.")

h2(doc, "2.3 O método da árvore de causas")
body(doc, "O método da árvore de causas (ADC), desenvolvido no Institut "
     "National de Recherche et de Sécurité francês na década de 1970, parte do "
     "acidente consumado e reconstrói, de trás para frente, a rede de fatos "
     "que o produziram, perguntando sistematicamente o que foi necessário para "
     "que cada fato ocorresse e se ele foi suficiente. O resultado é um "
     "diagrama — a árvore — que explicita as combinações de variações do "
     "trabalho habitual que culminaram no evento, sem juízo de culpa "
     "(BINDER; MONTEAU; ALMEIDA, 1995).")
body(doc, "No Brasil, o método foi difundido por Binder e Almeida, que "
     "demonstraram em estudos de caso sua capacidade de revelar o papel de "
     "fatores gerenciais e de organização do trabalho na gênese dos acidentes "
     "— designação improvisada de trabalhadores a funções, execução deixada à "
     "iniciativa individual, falta de ferramentas adequadas e falhas de "
     "circulação de informação (BINDER; ALMEIDA, 1997). Almeida (2006) "
     "aprofundou a crítica à “atribuição de culpa” como paradigma dominante "
     "nas empresas brasileiras, mostrando que investigações centradas no "
     "comportamento do acidentado produzem recomendações inócuas — treinar, "
     "advertir, punir — e deixam a porta aberta à recorrência.")
body(doc, "O sistema desenvolvido neste trabalho operacionaliza a lógica da "
     "ADC em três níveis de causas, terminologia corrente na prática de SST: "
     "causas imediatas (atos e condições no momento do evento), causas "
     "subjacentes (fatores do posto e da tarefa que possibilitaram as "
     "imediatas) e causas básicas (fatores de gestão e organização que "
     "originaram as subjacentes). A estrutura em níveis força o investigador a "
     "não encerrar a análise na primeira resposta — mecanismo digital "
     "equivalente à pergunta iterativa da árvore.")

h2(doc, "2.4 Fatores humanos e organizacionais: fadiga, estresse e burnout")
body(doc, "A investigação que chega às causas básicas encontra, com "
     "frequência, fatores humanos e organizacionais: jornadas extensas, "
     "pressão de tempo, sobreaviso, fadiga e estados de esgotamento que "
     "degradam a atenção e a tomada de decisão. A NR-1, na redação da Portaria "
     "MTE nº 1.419/2024, passou a exigir a inclusão dos fatores de risco "
     "psicossocial no GRO (BRASIL, 2024) — exigência que conversa diretamente "
     "com a análise de acidentes: os mesmos fatores que adoecem contribuem "
     "para a ocorrência de eventos agudos.")
body(doc, "Entre os desfechos do estresse ocupacional crônico, a síndrome de "
     "burnout — exaustão, distanciamento mental e redução da eficácia — foi "
     "incluída na CID-11 (código QD85) como fenômeno ocupacional (OMS, 2019), "
     "com adoção oficial da classificação no Brasil a partir de 2025. Estudo "
     "epidemiológico nacional recente indica tendência de crescimento das "
     "notificações entre 2014 e 2024 (TREML et al., 2025). No setor elétrico, "
     "Souza et al. (2010) encontraram prevalência de 20,3% de transtornos "
     "mentais comuns em eletricitários, associada a alta demanda, baixo "
     "controle e baixo apoio social. Para este trabalho, a implicação é dupla: "
     "o formulário de investigação inclui fatores humanos e organizacionais "
     "entre as causas selecionáveis, e a plataforma mantém módulo específico "
     "de avaliação psicossocial, permitindo cruzar a sinalização coletiva de "
     "risco com a ocorrência de eventos.")

h2(doc, "2.5 Indicadores de desempenho em SST")
body(doc, "A NBR 14280 padroniza os indicadores reativos clássicos: taxa de "
     "frequência (acidentes por milhão de horas-homem de exposição) e taxa de "
     "gravidade (dias perdidos e debitados por milhão de horas-homem) (ABNT, "
     "2001). A literatura contemporânea recomenda complementá-los com "
     "indicadores proativos — quase acidentes reportados, ações concluídas no "
     "prazo, tempo entre evento e registro —, que medem o funcionamento do "
     "sistema de gestão antes que a lesão ocorra. A razão entre quase "
     "acidentes e acidentes registrados serve como termômetro da cultura de "
     "reporte: valores baixos indicam subnotificação, não segurança. O sistema "
     "calcula ambos os grupos de indicadores automaticamente.")

h2(doc, "2.6 Marco normativo brasileiro")
body(doc, "O arcabouço normativo articula quatro camadas. Na base legal, a Lei "
     "nº 8.213/1991 define o acidente do trabalho e obriga a emissão da CAT — "
     "hoje transmitida eletronicamente pelo evento S-2210 do eSocial — até o "
     "primeiro dia útil seguinte à ocorrência, e imediatamente em caso de "
     "óbito (BRASIL, 1991). Na camada regulamentar, a NR-1 estabelece o GRO e "
     "exige a análise dos acidentes e doenças relacionados ao trabalho com "
     "identificação de causas e realimentação do PGR (BRASIL, 2024). Na "
     "camada técnica, a NBR 14280 padroniza cadastro, classificação e "
     "estatísticas (ABNT, 2001), e a ISO 45001 exige processos de reporte e "
     "investigação de incidentes no sistema de gestão (ABNT, 2018). Na camada "
     "setorial, a NR-10 — com seus requisitos para trabalhos em instalações "
     "desenergizadas, energizadas e no SEP — e a NR-35 regem as atividades "
     "típicas da distribuição de energia (BRASIL, 2019). O sistema digital "
     "materializa as obrigações da NR-1 e da NBR 14280 e registra explicitamente "
     "que não substitui a CAT.")

h2(doc, "2.7 Acidentalidade no setor de distribuição de energia elétrica")
body(doc, "O setor elétrico brasileiro apresenta perfil de acidentalidade de "
     "alta gravidade. O anuário da Abracopel contabilizou 2.089 acidentes de "
     "origem elétrica no ano-base 2023, com 781 óbitos; as áreas de geração e "
     "distribuição concentram parcela expressiva dos choques fatais "
     "(ABRACOPEL, 2024). No recorte das distribuidoras, a Abradee registrou "
     "250 acidentes fatais envolvendo a rede elétrica em 2023 (ABRADEE, 2024). "
     "[PREENCHER: se disponível, acrescentar estatística setorial de "
     "trabalhadores próprios e terceirizados — ex.: relatórios Funcoge — com "
     "fonte e ano.] Além do choque elétrico e do arco voltaico, as equipes de "
     "linha estão expostas à queda de altura, ao trânsito e a intempéries; a "
     "literatura setorial associa a acidentalidade também a fatores "
     "organizacionais como pressão por restabelecimento e terceirização "
     "(SOUZA et al., 2010). Esse perfil — eventos raros e graves — torna o "
     "aprendizado com quase acidentes particularmente valioso: eles fornecem o "
     "volume de sinais que os acidentes, felizmente menos numerosos, não "
     "fornecem.")

h2(doc, "2.8 Digitalização do registro e investigação de ocorrências")
body(doc, "O processo tradicional — formulário de papel preenchido após o "
     "retorno à base, transcrito depois para planilha — introduz perda de "
     "informação (memória, fotos, testemunhas), atraso e inconsistência de "
     "classificação. A digitalização com registro no local, em dispositivo "
     "móvel, endereça as três fragilidades: captura imediata de dados e "
     "evidências, classificação guiada e base única para indicadores e "
     "auditoria. A decisão de projeto deste trabalho — tecnologias gratuitas, "
     "sem instalação, operáveis por equipes de SST sem apoio de TI — visa "
     "remover a barreira de custo que mantém o processo manual nas "
     "organizações menores, liberando esforço para o que efetivamente "
     "previne: a investigação bem-feita e a execução das ações.")

h2(doc, "2.9 Estudos correlatos")
body(doc, "A literatura brasileira sobre análise de acidentes consolidou-se em "
     "torno do método da árvore de causas e da crítica ao paradigma "
     "culpabilizador (BINDER; ALMEIDA, 1997; ALMEIDA, 2006), com aplicações "
     "setoriais diversas. Trabalhos recentes exploram a gestão de ocorrências "
     "em plataformas comerciais de SST e o uso de indicadores proativos; "
     "observa-se, porém, escassez de propostas que levem o método estruturado "
     "de investigação ao campo em ferramenta digital gratuita, acessível a "
     "organizações com equipes enxutas — lacuna que este trabalho pretende "
     "ajudar a preencher. [PREENCHER: se a supervisão solicitar, acrescentar "
     "dois ou três estudos correlatos específicos com citação completa.]")

h2(doc, "2.10 Síntese da revisão")
body(doc, "Três conclusões da revisão orientam o desenvolvimento: (i) o modelo "
     "causal adotado determina a qualidade da investigação — sistemas que "
     "induzem o investigador a parar no “ato inseguro” perpetuam a "
     "recorrência, e a árvore de causas é o antídoto metodológico consolidado; "
     "(ii) o quase acidente é o insumo mais abundante e mais barato do "
     "aprendizado em segurança, mas exige reporte fácil e imediato para "
     "vencer a subnotificação; (iii) a exigência da NR-1 de analisar "
     "ocorrências e realimentar o PGR demanda registro rastreável — atributo "
     "que a digitalização entrega a custo marginal nulo. Esses três pontos "
     "definem os requisitos do sistema especificado na seção 3.")
doc.add_page_break()

# ===================================================== 3 METODOLOGIA ==========
h1(doc, "3 METODOLOGIA")

h2(doc, "3.1 Caracterização da pesquisa")
body(doc, "Trata-se de pesquisa aplicada, de natureza tecnológica, conduzida na "
     "forma de desenvolvimento de artefato (o sistema digital) seguido de "
     "estudo de caso. A abordagem é quali-quantitativa: quantitativa no "
     "tratamento dos registros e indicadores, qualitativa na análise das "
     "investigações e da adequação do processo ao contexto. O estudo de caso "
     "foi conduzido em uma organização do setor de distribuição de energia "
     "elétrica, caracterizada de forma a não permitir identificação: "
     "[PREENCHER: porte, região de atuação e natureza das atividades, SEM nome "
     "nem dados identificáveis].")

h2(doc, "3.2 Passo a passo da monografia")
bullet(doc, "Etapa 1 — Revisão bibliográfica e normativa: modelos de "
       "causalidade, método da árvore de causas, indicadores e requisitos da "
       "NR-1, NBR 14280, ISO 45001, NR-10 e NR-35.", numbered=True)
bullet(doc, "Etapa 2 — Especificação do sistema: taxonomia de ocorrências, "
       "fluxo de registro em campo, roteiro de investigação em três níveis de "
       "causas, gestão de ações e indicadores.", numbered=True)
bullet(doc, "Etapa 3 — Projeto da arquitetura: premissas de custo zero, "
       "ausência de instalação e operação em dispositivo móvel; escolha das "
       "tecnologias (frontend web autocontido e backend gratuito em nuvem).",
       numbered=True)
bullet(doc, "Etapa 4 — Desenvolvimento: implementação do módulo de acidentes "
       "(central) e dos módulos complementares (permissão de trabalho "
       "NR-10/NR-35, fatores psicossociais NR-1, AEP NR-17) sobre a mesma base "
       "de dados.", numbered=True)
bullet(doc, "Etapa 5 — Testes e validação técnica: funcionamento, integridade "
       "dos dados e usabilidade em campo (celular, condições de rua).",
       numbered=True)
bullet(doc, "Etapa 6 — Estudo de caso, parte documental: análise retrospectiva "
       "de ocorrências históricas anonimizadas da organização, reinvestigadas "
       "com o roteiro de três níveis para comparação entre a análise original "
       "e a estruturada.", numbered=True)
bullet(doc, "Etapa 7 — Estudo de caso, parte de campo: uso piloto do sistema "
       "pelas equipes para registro de novas ocorrências e condições "
       "inseguras.", numbered=True)
bullet(doc, "Etapa 8 — Análise e discussão: consolidação dos indicadores, "
       "comparação antes/depois do processo e discussão à luz da literatura; "
       "proposição de plano de ação.", numbered=True)

h2(doc, "3.3 Aspectos éticos e anonimato")
body(doc, "O estudo não expõe a organização nem pessoas. Os casos analisados "
     "são descritos de forma anonimizada — função genérica do envolvido, sem "
     "nome, data exata, local ou número de CAT — e as telas e figuras não "
     "contêm dados pessoais. O tratamento dos registros observa a Lei nº "
     "13.709/2018 (LGPD): os dados de ocorrências são tratados no legítimo "
     "interesse da prevenção, com acesso restrito, e as análises são "
     "apresentadas de forma agregada. Em conformidade com a orientação da "
     "supervisão, os agradecimentos citam apenas a CERPRO.")

h2(doc, "3.4 Arquitetura da plataforma")
body(doc, "O sistema é o módulo central da plataforma Psike, construída sobre "
     "três decisões: custo zero de operação, ausência de instalação e uso "
     "direto no dispositivo móvel do usuário. O frontend é um único arquivo "
     "web executado no navegador; o backend é um serviço gratuito em nuvem que "
     "grava os registros em planilha eletrônica — uma aba por módulo —, "
     "servindo de repositório único. Sem backend configurado, a aplicação "
     "opera em modo local para demonstração. Os quatro módulos — acidentes "
     "(central), permissão de trabalho, fatores psicossociais e AEP — "
     "compartilham a mesma base, o que permite cruzar ocorrências com "
     "permissões emitidas e com a sinalização coletiva de risco psicossocial.")

h2(doc, "3.5 O módulo de registro e investigação de ocorrências")
h3(doc, "3.5.1 Registro em campo")
body(doc, "O registro é projetado para ser feito no local, em menos de cinco "
     "minutos, pelo encarregado ou técnico de segurança: classificação do "
     "evento (acidente com afastamento, acidente sem afastamento, quase "
     "acidente ou condição insegura), data e hora, função genérica do "
     "envolvido, descrição livre do fato e da tarefa em execução, e parte do "
     "corpo atingida quando houver lesão. O sistema alerta, no ato, para a "
     "obrigação legal de emissão da CAT nos casos aplicáveis — deixando "
     "explícito que o registro interno não a substitui.")
h3(doc, "3.5.2 Investigação em três níveis de causas")
body(doc, "A investigação estruturada aplica a lógica da árvore de causas em "
     "três níveis sequenciais e obrigatórios: causas imediatas (o que, no "
     "momento do evento, produziu o dano — atos e condições), causas "
     "subjacentes (o que, na tarefa e no posto, tornou possíveis as causas "
     "imediatas — ferramenta inadequada, procedimento inexistente ou "
     "inaplicável, improvisação, pressa) e causas básicas (o que, na gestão e "
     "na organização, originou as subjacentes — dimensionamento de equipe, "
     "planejamento do serviço, capacitação, manutenção, fatores humanos e "
     "organizacionais, incluindo fadiga e sobrecarga). O formulário não "
     "permite concluir a investigação apenas com causas imediatas — "
     "mecanismo que traduz, no software, a pergunta iterativa do método: o "
     "que foi necessário para que este fato ocorresse?")
h3(doc, "3.5.3 Ações e indicadores")
body(doc, "Cada investigação vincula ações corretivas e preventivas com "
     "responsável e prazo, cujo status é acompanhado no painel. O sistema "
     "calcula automaticamente: contagem de ocorrências por tipo e período; "
     "razão quase acidentes/acidentes (termômetro da cultura de reporte); "
     "tempo mediano entre evento e registro; percentual de ações concluídas no "
     "prazo; e, alimentado com as horas-homem de exposição, as taxas de "
     "frequência e de gravidade da NBR 14280. [PREENCHER: confirmar quais "
     "indicadores foram efetivamente ativados no piloto.]")

h2(doc, "3.6 Módulos complementares")
body(doc, "Permissão de trabalho digital (NR-10/NR-35): checklist específico por "
     "tipo de serviço — rede desenergizada, com a sequência de desenergização "
     "da NR-10; rede energizada/SEP; e trabalho em altura —, com captura de "
     "geolocalização e assinatura em tela; a liberação só se habilita com o "
     "checklist completo. Avaliação de fatores psicossociais (NR-1): "
     "instrumento anônimo de dez itens em seis dimensões, com classificação "
     "automática e geração de texto para o IRO — mantido na plataforma como "
     "fonte de sinais de fatores humanos e organizacionais que a investigação "
     "de acidentes pode corroborar. Avaliação Ergonômica Preliminar (NR-17): "
     "formulário de doze itens em cinco blocos com parecer automático. A "
     "integração das bases permite, por exemplo, verificar se um acidente "
     "ocorreu em serviço coberto por permissão e se o setor envolvido já "
     "sinalizava sobrecarga na avaliação psicossocial.")

h2(doc, "3.7 Estudo de caso: desenho da aplicação")
body(doc, "Parte documental (retrospectiva): [PREENCHER: número] ocorrências "
     "registradas pela organização no período de [PREENCHER: período] foram "
     "anonimizadas e reinvestigadas com o roteiro de três níveis, permitindo "
     "comparar a profundidade causal da análise original com a estruturada. "
     "Parte de campo (piloto): no período de [PREENCHER: período], as equipes "
     "de [PREENCHER: escopo] utilizaram o sistema para registro de novas "
     "ocorrências e condições inseguras. [PREENCHER: registrar aprovações "
     "internas e forma de treinamento das equipes, sem identificar a "
     "organização.]")

h2(doc, "3.8 Tratamento e análise dos dados")
body(doc, "As ocorrências são tratadas de forma agregada e anonimizada. A "
     "análise compara: (i) a distribuição de causas por nível na análise "
     "original versus na estruturada — hipótese de que a estruturada revela "
     "mais causas básicas; (ii) o volume e a proporção de quase acidentes "
     "reportados antes e durante o piloto — hipótese de redução da "
     "subnotificação; e (iii) os tempos entre evento e registro. Nenhuma "
     "análise identifica indivíduos; os casos ilustrativos são descritos com "
     "função genérica e circunstâncias descaracterizadas.")
doc.add_page_break()

# ============================================ 4 RESULTADOS E DISCUSSÃO =========
h1(doc, "4 RESULTADOS E DISCUSSÃO")
body(doc, "Esta seção integra a apresentação dos resultados e sua discussão à "
     "luz da literatura e do marco normativo da seção 2.")

h2(doc, "4.1 O sistema desenvolvido")
body(doc, "O sistema foi implementado integralmente e encontra-se operacional. "
     "O fluxo completo — registro no local, investigação em três níveis, "
     "vinculação de ações e atualização dos indicadores — ocorre sem etapa "
     "manual de transcrição, a custo de operação nulo. A Figura 1 apresenta as "
     "telas principais. [PREENCHER: inserir Figura 1 — telas do módulo de "
     "acidentes, sem dados pessoais.]")

h2(doc, "4.2 Resultados da parte documental (retrospectiva)")
body(doc, "[PREENCHER: apresentar a comparação entre as análises originais e as "
     "reinvestigadas — tabela com a distribuição de causas por nível "
     "(imediatas/subjacentes/básicas) em cada abordagem, e dois ou três casos "
     "ilustrativos anonimizados mostrando as causas básicas que a análise "
     "original não havia alcançado. NÃO identificar organização, pessoas, "
     "datas ou locais.]")

h2(doc, "4.3 Resultados do piloto de campo")
body(doc, "[PREENCHER: apresentar o volume de registros do piloto por tipo "
     "(acidentes, quase acidentes, condições inseguras), a razão quase "
     "acidentes/acidentes, o tempo mediano evento–registro e o status das "
     "ações. Inserir os gráficos do painel. Comparar com o processo anterior "
     "quando houver base.]")

h2(doc, "4.4 Discussão à luz da literatura")
body(doc, "Os resultados [PREENCHER: ajustar conforme os dados reais] permitem "
     "discutir as três hipóteses da revisão. Primeira: se a investigação "
     "estruturada revelou mais causas subjacentes e básicas que a análise "
     "original, confirma-se no caso concreto o diagnóstico de Binder e Almeida "
     "(1997) e Almeida (2006) — o instrumento condiciona a profundidade da "
     "análise, e roteiros que aceitam o “ato inseguro” como resposta final "
     "produzem prevenção inócua. Segunda: se o piloto elevou o reporte de "
     "quase acidentes, corrobora-se a leitura de que a subnotificação é "
     "sobretudo função do atrito do processo — na linha da pirâmide de "
     "Heinrich (1931) e Bird e Germain (1985), o sistema passa a capturar a "
     "base de sinais antes invisível. Terceira: a presença de fatores humanos "
     "e organizacionais — fadiga, pressão de tempo, sobrecarga — entre as "
     "causas básicas articula a análise de acidentes com a gestão de fatores "
     "psicossociais exigida pela NR-1 e com o quadro de adoecimento mental "
     "ocupacional em crescimento no país (TREML et al., 2025), reforçando que "
     "prevenção de acidentes e saúde mental são faces do mesmo sistema de "
     "gestão.")

h2(doc, "4.5 Integração ao PGR e plano de ação")
body(doc, "As causas básicas consolidadas realimentam o inventário de riscos: "
     "cada fator organizacional recorrente identificado nas investigações é "
     "candidato a perigo a reavaliar no IRO, e as ações vinculadas compõem o "
     "plano de ação do PGR com responsáveis e prazos — exatamente o ciclo que "
     "a NR-1 exige (BRASIL, 2024). [PREENCHER: descrever as revisões de "
     "inventário e o plano de ação efetivamente decorrentes do estudo, sem "
     "identificar a organização.]")

h2(doc, "4.6 Limitações")
bullet(doc, "O roteiro em três níveis é uma operacionalização simplificada da "
       "árvore de causas: estrutura a progressão causal, mas não desenha o "
       "diagrama completo de combinações do método original;")
bullet(doc, "A comparação retrospectiva depende da qualidade dos registros "
       "históricos disponíveis;")
bullet(doc, "O piloto restringe-se a uma organização e a um período curto, "
       "limitando a generalização e a leitura de tendência dos indicadores;")
bullet(doc, "Ausência de autenticação e segregação de dados por organização, "
       "aceitável no piloto, requisito para uso amplo;")
bullet(doc, "Limites das camadas gratuitas de nuvem quanto a volume e "
       "requisições.")
doc.add_page_break()

# ============================================ 5 CONCLUSÃO =====================
h1(doc, "5 CONCLUSÃO")
body(doc, "O objetivo geral — desenvolver e aplicar um sistema digital de "
     "registro e investigação de acidentes e quase acidentes pelo método da "
     "árvore de causas, integrado ao GRO da NR-1, em uma organização de "
     "distribuição de energia elétrica — foi [PREENCHER: atingido / "
     "parcialmente atingido]. O sistema foi desenvolvido, está operacional e "
     "foi aplicado em estudo de caso documental e piloto de campo. "
     "[PREENCHER: uma frase objetiva sobre o principal achado — ex.: a "
     "investigação estruturada revelou causas organizacionais ausentes das "
     "análises originais e o reporte de quase acidentes aumentou durante o "
     "piloto.]")
body(doc, "Como trabalhos futuros, sugerem-se: a implementação do diagrama "
     "completo da árvore de causas no software; autenticação e segregação por "
     "organização; acompanhamento longitudinal dos indicadores para leitura de "
     "tendência; e a integração analítica entre os módulos — cruzando "
     "ocorrências, permissões de trabalho e sinalização psicossocial — para "
     "antecipação de cenários de risco.")
doc.add_page_break()

# ============================================ REFERÊNCIAS =====================
h1(doc, "REFERÊNCIAS")
REFS = [
    "ABRACOPEL — ASSOCIAÇÃO BRASILEIRA DE CONSCIENTIZAÇÃO PARA OS PERIGOS DA "
    "ELETRICIDADE. Anuário Estatístico de Acidentes de Origem Elétrica 2024 — "
    "ano-base 2023. Salto: Abracopel, 2024. Disponível em: [PREENCHER: URL e "
    "data de acesso].",
    "ABRADEE — ASSOCIAÇÃO BRASILEIRA DE DISTRIBUIDORES DE ENERGIA ELÉTRICA. "
    "Levantamento de acidentes com a rede elétrica — 2023. 2024. Disponível "
    "em: [PREENCHER: URL e data de acesso].",
    "ALMEIDA, I. M. Trajetória da análise de acidentes: o paradigma "
    "tradicional e os primórdios da ampliação da análise. Interface — "
    "Comunicação, Saúde, Educação, v. 10, n. 19, p. 185-202, 2006. "
    "[PREENCHER: conferir dados exatos do fascículo].",
    "ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. ABNT NBR 14280:2001 — Cadastro "
    "de acidente do trabalho — Procedimento e classificação. Rio de Janeiro: "
    "ABNT, 2001.",
    "ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. ABNT NBR ISO 45001:2018 — "
    "Sistemas de gestão de saúde e segurança ocupacional: requisitos com "
    "orientações para uso. Rio de Janeiro: ABNT, 2018.",
    "BINDER, M. C. P.; ALMEIDA, I. M. Estudo de caso de dois acidentes do "
    "trabalho investigados com o método de árvore de causas. Cadernos de "
    "Saúde Pública, Rio de Janeiro, v. 13, n. 4, p. 749-760, 1997.",
    "BINDER, M. C. P.; MONTEAU, M.; ALMEIDA, I. M. Árvore de causas: método "
    "de investigação de acidentes de trabalho. São Paulo: Publisher Brasil, "
    "1995. [PREENCHER: conferir edição/ano do exemplar consultado].",
    "BIRD, F. E.; GERMAIN, G. L. Practical loss control leadership. Loganville: "
    "International Loss Control Institute, 1985.",
    "BRASIL. Lei nº 8.213, de 24 de julho de 1991. Dispõe sobre os Planos de "
    "Benefícios da Previdência Social. Diário Oficial da União, Brasília, DF, "
    "25 jul. 1991.",
    "BRASIL. Lei nº 13.709, de 14 de agosto de 2018. Lei Geral de Proteção de "
    "Dados Pessoais (LGPD). Diário Oficial da União, Brasília, DF, 15 ago. 2018.",
    "BRASIL. Ministério do Trabalho e Emprego. Norma Regulamentadora nº 1 "
    "(NR-1): disposições gerais e gerenciamento de riscos ocupacionais. "
    "Redação dada pela Portaria MTE nº 1.419, de 27 de agosto de 2024. "
    "Brasília, DF: MTE, 2024.",
    "BRASIL. Ministério do Trabalho e Emprego. Norma Regulamentadora nº 10 "
    "(NR-10): segurança em instalações e serviços em eletricidade. Brasília, "
    "DF: MTE, 2019.",
    "BRASIL. Ministério do Trabalho e Emprego. Norma Regulamentadora nº 35 "
    "(NR-35): trabalho em altura. Brasília, DF: MTE, [PREENCHER: ano da "
    "redação consultada].",
    "BRASIL. Ministério da Previdência Social. Anuário Estatístico de "
    "Acidentes do Trabalho — AEAT. Brasília, DF, [PREENCHER: ano-base e URL "
    "da edição consultada].",
    "HEINRICH, H. W. Industrial accident prevention: a scientific approach. "
    "New York: McGraw-Hill, 1931.",
    "HOLLNAGEL, E. Safety-I and Safety-II: the past and future of safety "
    "management. Farnham: Ashgate, 2014.",
    "ORGANIZAÇÃO MUNDIAL DA SAÚDE (OMS). Classificação Estatística "
    "Internacional de Doenças e Problemas Relacionados à Saúde (CID-11): "
    "burn-out (QD85). Genebra: OMS, 2019.",
    "REASON, J. Human error. Cambridge: Cambridge University Press, 1990.",
    "REASON, J. Managing the risks of organizational accidents. Aldershot: "
    "Ashgate, 1997.",
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
]
for r in REFS:
    ref(doc, r)

# ============================================ APÊNDICES ======================
doc.add_page_break()
h1(doc, "APÊNDICE A — ROTEIRO DE REGISTRO E INVESTIGAÇÃO DE OCORRÊNCIAS")
body(doc, "Campos do registro em campo: tipo de ocorrência (acidente com "
     "afastamento / acidente sem afastamento / quase acidente / condição "
     "insegura); data e hora; função genérica do envolvido; tarefa em "
     "execução; descrição do fato; parte do corpo atingida (se lesão); alerta "
     "de emissão de CAT quando aplicável.", indent=False)
body(doc, "Roteiro de investigação (obrigatório para acidentes; recomendado "
     "para quase acidentes):", indent=False)
bullet(doc, "Nível 1 — Causas imediatas: o que, no momento do evento, produziu "
       "(ou quase produziu) o dano? Atos e condições presentes na cena.")
bullet(doc, "Nível 2 — Causas subjacentes: o que, na tarefa e no posto, tornou "
       "possíveis as causas imediatas? Ferramenta, procedimento, improvisação, "
       "pressa, planejamento do serviço.")
bullet(doc, "Nível 3 — Causas básicas: o que, na gestão e na organização, "
       "originou as causas subjacentes? Dimensionamento, capacitação, "
       "manutenção, fatores humanos e organizacionais (fadiga, sobrecarga, "
       "pressão de tempo).")
bullet(doc, "Ações: corretivas e preventivas, com responsável, prazo e status.")

doc.add_page_break()
h1(doc, "APÊNDICE B — ARQUITETURA E MÓDULOS DA PLATAFORMA")
body(doc, "A plataforma Psike compõe-se de frontend web autocontido e backend "
     "gratuito em nuvem, com quatro módulos sobre o mesmo repositório de "
     "dados: Módulo central — registro e investigação de acidentes e quase "
     "acidentes (NR-1/NBR 14280); Módulo 2 — permissão de trabalho "
     "(NR-10/NR-35); Módulo 3 — avaliação de fatores psicossociais (NR-1); "
     "Módulo 4 — Avaliação Ergonômica Preliminar (NR-17). [PREENCHER: inserir "
     "diagrama de arquitetura e telas, sem dados pessoais.]", indent=False)

doc.save(OUT)
print(f"OK: {OUT} gerado.")
