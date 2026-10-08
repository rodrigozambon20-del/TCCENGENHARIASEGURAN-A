# -*- coding: utf-8 -*-
"""
Gerador do TCC — Sistema digital de registro e investigação de acidentes e
quase acidentes pelo método da árvore de causas (plataforma Psike).
Produz TCC_Acidentes_ArvoreCausas_MODELO.docx.

Segue as regras oficiais da Turma 2025 (PECE/EPUSP):
- Guia Prático da Monografia 2025 (capa, folha de rosto, estilo dos títulos,
  estrutura 1 INTRODUÇÃO / 1.1 OBJETIVO / 1.2 JUSTIFICATIVA, figuras);
- Diretrizes USP 2024, 5. ed. (ABNT NBR 10520:2023 e NBR 6023:2018):
  citação "(Autor, ano)" em caixa baixa, título em negrito nas referências,
  resumo de 150 a 500 palavras precedido da referência do documento,
  palavras-chave separadas por ponto, pré-textuais contados e não numerados;
- Aviso da turma: entrega em PDF pelo Moodle até 22/02/2027 (ano = 2027);
- Orientação da supervisão (03/07/2026): 40-80 folhas, literatura recente,
  burnout citado, passo a passo, resultados e discussão integrados,
  conclusão curta, anonimato, agradecimento apenas à CERPRO.

Uso:  python3 gerar_tcc_acidentes.py     Requer: pip install python-docx
Trechos [PREENCHER: ...] dependem de dados reais e saem realçados em amarelo.
"""
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.text import WD_COLOR_INDEX
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = "TCC_Acidentes_ArvoreCausas_MODELO.docx"
ANO = "2027"
TITULO = ("SISTEMA DIGITAL DE REGISTRO E INVESTIGAÇÃO DE ACIDENTES E QUASE "
          "ACIDENTES PELO MÉTODO DA ÁRVORE DE CAUSAS")
SUBTITULO = ("desenvolvimento e aplicação na gestão de riscos ocupacionais "
             "(NR-1) em serviços de distribuição de energia elétrica")
REF_DOC = ("ZAMBON, R. [PREENCHER: iniciais conforme nome completo]. **Sistema "
           "digital de registro e investigação de acidentes e quase acidentes "
           "pelo método da árvore de causas**: desenvolvimento e aplicação na "
           "gestão de riscos ocupacionais (NR-1) em serviços de distribuição de "
           "energia elétrica. " + ANO + ". [PREENCHER: nº de folhas] f. "
           "Monografia (Especialização em Engenharia de Segurança do Trabalho) "
           "– Escola Politécnica, Universidade de São Paulo, São Paulo, " + ANO + ".")

# ---------------------------------------------------------------- estilos ----

def _plain_font(style, size, bold, italic=False):
    style.font.name = "Arial"
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = italic
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts"); rpr.append(rfonts)
    for att in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(att)) is not None:
            del rfonts.attrib[qn(att)]
    for att in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(att), "Arial")
    style.font.color.rgb = RGBColor(0, 0, 0)

def set_base_styles(doc):
    normal = doc.styles["Normal"]
    _plain_font(normal, 12, False)
    pf = normal.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    # Guia Prático: seção primária CAIXA ALTA negrito; secundária CAIXA ALTA
    # sem negrito; terciária caixa baixa negrito; quaternária sem negrito.
    for name, bold in (("Heading 1", True), ("Heading 2", False),
                       ("Heading 3", True), ("Heading 4", False)):
        h = doc.styles[name]
        _plain_font(h, 12, bold)
        hp = h.paragraph_format
        hp.space_before = Pt(18)
        hp.space_after = Pt(18)
        hp.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        hp.keep_with_next = True
        hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    cap = doc.styles["Caption"]
    _plain_font(cap, 10, False)
    cap.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    cap.paragraph_format.space_before = Pt(12)
    cap.paragraph_format.space_after = Pt(4)
    cap.paragraph_format.keep_with_next = True

def update_fields_on_open(doc):
    """Faz o Word oferecer a atualização de sumário e listas ao abrir."""
    settings = doc.settings.element
    el = OxmlElement("w:updateFields"); el.set(qn("w:val"), "true")
    settings.append(el)

# ---------------------------------------------------------------- seções -----

def set_margins(sec):
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin, sec.left_margin = Cm(3), Cm(3)
    sec.bottom_margin, sec.right_margin = Cm(2), Cm(2)

def new_section(doc, restart_at=None):
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    set_margins(sec)
    sec.header.is_linked_to_previous = False
    inherited = sec._sectPr.find(qn("w:pgNumType"))
    if inherited is not None:
        sec._sectPr.remove(inherited)
    if restart_at is not None:
        pg = OxmlElement("w:pgNumType"); pg.set(qn("w:start"), str(restart_at))
        cols = sec._sectPr.find(qn("w:cols"))
        if cols is not None:
            cols.addprevious(pg)
        else:
            sec._sectPr.append(pg)
    return sec

def page_number_header(sec):
    """Número no canto superior direito, fonte menor (Diretrizes USP 2.2.2)."""
    p = sec.header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(); run.font.size = Pt(10)
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

# ---------------------------------------------------------------- texto ------

def _runs(p, text, size=None):
    """Texto com **negrito** e [PREENCHER: ...] realçado em amarelo."""
    for chunk in re.split(r"(\*\*.+?\*\*|\[PREENCHER[^\]]*\])", text):
        if not chunk:
            continue
        if chunk.startswith("**"):
            r = p.add_run(chunk[2:-2]); r.bold = True
        elif chunk.startswith("[PREENCHER"):
            r = p.add_run(chunk); r.bold = True
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        else:
            r = p.add_run(chunk)
        if size:
            r.font.size = Pt(size)

def center(doc, text, bold=False, size=12, space_after=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(space_after)
    _runs(p, f"**{text}**" if bold and text else text, size)
    return p

def blank(doc, n=1):
    for _ in range(n):
        doc.add_paragraph()

def body(doc, text, indent=True):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent:
        p.paragraph_format.first_line_indent = Cm(1.25)
    _runs(p, text)
    return p

def alineas(doc, items):
    """Alíneas ABNT: a), b)... terminadas em ';' e a última em '.'."""
    for i, item in enumerate(items):
        end = "." if i == len(items) - 1 else ";"
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Cm(1.25)
        p.paragraph_format.first_line_indent = Cm(-0.6)
        _runs(p, f"{chr(97 + i)}) {item}{end}")

def ref(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_after = Pt(12)
    _runs(p, text)

def pretextual_title(doc, text):
    """Títulos sem indicativo numérico: centralizados, fora do sumário."""
    center(doc, text, bold=True, space_after=18)

def h1(doc, t):
    doc.add_heading(t.upper(), level=1)

def h1_center(doc, t):
    """Pós-textuais (referências, apêndices): centralizados, no sumário."""
    h = doc.add_heading(t.upper(), level=1)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

def h2(doc, t):
    doc.add_heading(t.upper(), level=2)

def h3(doc, t):
    doc.add_heading(t, level=3)

def _field(p, instr, cached):
    run = p.add_run()
    b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "begin")
    i = OxmlElement("w:instrText"); i.set(qn("xml:space"), "preserve"); i.text = instr
    s = OxmlElement("w:fldChar"); s.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t"); t.text = cached
    e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end")
    for el in (b, i, s): run._r.append(el)
    r2 = p.add_run(); r2._r.append(t)
    r3 = p.add_run(); r3._r.append(e)

def list_field(doc, instr, hint):
    p = doc.add_paragraph()
    _field(p, instr, hint)

COUNTERS = {"Figura": 0, "Quadro": 0, "Tabela": 0}

def caption(doc, kind, title):
    """Identificação na parte superior: 'Figura 1 – Título' (Guia Prático)."""
    COUNTERS[kind] += 1
    p = doc.add_paragraph(style="Caption")
    p.add_run(f"{kind} ")
    _field(p, f"SEQ {kind} \\* ARABIC", str(COUNTERS[kind]))
    _runs(p, f" – {title}")

def fonte(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_after = Pt(12)
    _runs(p, "Fonte: " + text, size=10)

def figura(doc, title, placeholder, source):
    caption(doc, "Figura", title)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    _runs(p, f"[PREENCHER: {placeholder}]")
    fonte(doc, source)

def equacao(doc, expr, n):
    """Equação centralizada com numeração entre parênteses à direita."""
    p = doc.add_paragraph()
    p.paragraph_format.tab_stops.add_tab_stop(Cm(8), alignment=1)   # centro
    p.paragraph_format.tab_stops.add_tab_stop(Cm(16), alignment=2)  # direita
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    _runs(p, f"\t{expr}\t({n})")

def _borders(el, spec):
    tag = "w:tblBorders" if el.tag == qn("w:tblPr") else "w:tcBorders"
    node = el.find(qn(tag))
    if node is None:
        node = OxmlElement(tag); el.append(node)
    for side, val in spec.items():
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), val)
        if val != "nil":
            b.set(qn("w:sz"), "8"); b.set(qn("w:color"), "000000")
        node.append(b)

def tabela(doc, kind, title, header, rows, source, widths=None):
    """Quadro: bordas fechadas. Tabela: padrão IBGE, laterais abertas."""
    caption(doc, kind, title)
    t = doc.add_table(rows=1 + len(rows), cols=len(header))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = t._tbl.tblPr
    if kind == "Quadro":
        _borders(tblPr, {s: "single" for s in
                         ("top", "left", "bottom", "right", "insideH", "insideV")})
    else:
        _borders(tblPr, {"top": "single", "bottom": "single", "left": "nil",
                         "right": "nil", "insideH": "nil", "insideV": "nil"})
    for r_i, values in enumerate([header] + rows):
        for c_i, val in enumerate(values):
            cell = t.cell(r_i, c_i)
            if widths:
                cell.width = Cm(widths[c_i])
            par = cell.paragraphs[0]
            par.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            par.alignment = (WD_ALIGN_PARAGRAPH.CENTER
                             if kind == "Tabela" and c_i > 0
                             else WD_ALIGN_PARAGRAPH.LEFT)
            _runs(par, f"**{val}**" if r_i == 0 else val, size=10)
            if r_i == 0 and kind == "Tabela":
                _borders(cell._tc.get_or_add_tcPr(), {"bottom": "single"})
    fonte(doc, source)

# ---------------------------------------------------------------- documento --

doc = Document()
set_base_styles(doc)
set_margins(doc.sections[0])
update_fields_on_open(doc)

# ===== Seção 1: instruções + capa (não contadas) =============================
center(doc, "COMO USAR ESTE MODELO — APAGUE ESTA PÁGINA ANTES DE ENTREGAR",
       bold=True, size=13, space_after=14)
body(doc, "Modelo estruturado conforme o Guia Prático da Monografia 2025, a aula "
     "de monografia da Turma 2025 e as Diretrizes USP 2024 (5. ed.), de uso "
     "obrigatório. Trechos [PREENCHER: ...] em amarelo dependem de dados reais "
     "e não devem ser inventados.", indent=False)
body(doc, "PRAZOS (Aviso Monografia Turma 2025): tema por e-mail para "
     "renata@gmirmusp.com.br até 03/11/2026; monografia em PDF pelo Moodle até "
     "22/02/2027 (quem não enviar é reprovado); pôster em PDF até 22/03/2027 "
     "(aprovados); apresentações online de 29/03/2027 a 30/06/2027.",
     indent=False)
body(doc, "ANONIMATO E FOTOS: não citar o nome da empresa nem dados pessoais "
     "(nomes, matrículas, número de CAT, datas exatas, locais). Fotos do estudo "
     "de caso são esperadas — e obrigatórias no pôster —, desde que o rosto dos "
     "trabalhadores e a identificação da empresa em uniformes, veículos e "
     "equipamentos estejam ocultados (desfoque ou tarja). Figuras centralizadas, "
     "de tamanho padronizado, com título acima e fonte abaixo.", indent=False)
body(doc, "ANTES DE ENTREGAR: (1) preencher os campos amarelos; (2) atualizar "
     "sumário e listas (Ctrl+A e F9 no Word); (3) conferir as palavras-chave no "
     "Vocabulário Controlado USP (vocabusp.abcd.usp.br); (4) inserir, em nota de "
     "rodapé no título REFERÊNCIAS, “Elaboradas de acordo com a ABNT NBR "
     "6023:2018”; (5) conferir a extensão (40 a 80 folhas); (6) apagar esta "
     "página e salvar em PDF (Arquivo > Salvar como > PDF) para o Moodle.",
     indent=False)
doc.add_page_break()

# Capa — Guia Prático: autor, título, local, ano
p = center(doc, "RODRIGO ZAMBON ", bold=True)
_runs(p, "[PREENCHER: nome completo]")
blank(doc, 9)
center(doc, TITULO + ":", bold=True)
center(doc, SUBTITULO, bold=True)
blank(doc, 12)
center(doc, "São Paulo")
center(doc, ANO)

# ===== Seção 2: pré-textuais (contados a partir da folha de rosto, sem número)
new_section(doc, restart_at=1)

p = center(doc, "RODRIGO ZAMBON ", bold=True)
_runs(p, "[PREENCHER: nome completo]")
blank(doc, 8)
center(doc, TITULO + ":", bold=True)
center(doc, SUBTITULO, bold=True)
blank(doc, 3)
nat = doc.add_paragraph()
nat.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
nat.paragraph_format.left_indent = Cm(8)
nat.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
_runs(nat, "Monografia apresentada à Escola Politécnica da Universidade de São "
      "Paulo para a obtenção do título de Especialista em Engenharia de "
      "Segurança do Trabalho.")
blank(doc, 9)
center(doc, "São Paulo")
center(doc, ANO)
doc.add_page_break()

pretextual_title(doc, "AGRADECIMENTOS")
body(doc, "À CERPRO, pelo apoio e pela viabilização deste trabalho.")
body(doc, "À supervisão da monografia e ao corpo docente do Programa de Educação "
     "Continuada em Engenharia da Escola Politécnica da Universidade de São "
     "Paulo, pelas orientações ao longo do curso.")
body(doc, "À minha família, pelo apoio durante a especialização.")
doc.add_page_break()

# Resumo — Diretrizes USP 3.1.2.1.8: referência do documento, parágrafo único,
# 150 a 500 palavras, palavras-chave separadas e finalizadas por ponto.
pretextual_title(doc, "RESUMO")
ref(doc, REF_DOC)
body(doc, "A Norma Regulamentadora nº 1 exige que o empregador analise os "
     "acidentes do trabalho e utilize os resultados na revisão do Programa de "
     "Gerenciamento de Riscos; na prática, o registro de ocorrências ainda é "
     "manual e tardio, a investigação tende a encerrar-se na conduta do "
     "acidentado e os quase acidentes raramente são reportados. Este trabalho "
     "teve por objetivo desenvolver e aplicar um sistema digital de registro e "
     "investigação de acidentes e quase acidentes estruturado no método da "
     "árvore de causas, em uma organização do setor de distribuição de energia "
     "elétrica. Adotou-se pesquisa aplicada, com desenvolvimento de artefato "
     "seguido de estudo de caso em duas partes: análise documental "
     "retrospectiva de ocorrências anonimizadas e piloto de campo. O sistema "
     "permite o registro imediato no local pelo celular, conduz o investigador "
     "por três níveis obrigatórios de causas — imediatas, subjacentes e "
     "básicas —, vincula ações com responsável e prazo e calcula indicadores "
     "reativos e proativos. Na análise documental, [PREENCHER: "
     "principal resultado — ex.: proporção de causas básicas identificadas na "
     "análise original e na estruturada]; no piloto, [PREENCHER: principal "
     "resultado — ex.: volume de quase acidentes reportados e tempo entre evento "
     "e registro]. Conclui-se que a digitalização do ciclo de registro, "
     "investigação e ação [PREENCHER: confirmar ou ajustar conforme os "
     "resultados] aprofunda a análise causal em direção aos fatores "
     "organizacionais, reduz a subnotificação de quase acidentes e confere aos "
     "registros a rastreabilidade exigida pela norma, a custo praticamente "
     "nulo.", indent=False)
body(doc, "Palavras-chave: Acidentes do trabalho. Quase acidentes. Árvore de "
     "causas. Investigação de acidentes. Segurança do trabalho. Energia "
     "elétrica.", indent=False)
doc.add_page_break()

pretextual_title(doc, "ABSTRACT")
ref(doc, REF_DOC)
body(doc, "Brazilian Regulatory Standard No. 1 requires employers to analyze "
     "work-related accidents and diseases and to use the results to review the "
     "Risk Management Program; in practice, incident recording remains manual "
     "and late, investigations tend to stop at the injured worker's conduct, and "
     "near misses are rarely reported. This study aimed to develop and apply a "
     "digital system for recording and investigating accidents and near misses "
     "structured on the causal tree method in an electric power distribution "
     "organization. Applied, technological research was adopted, with artifact "
     "development followed by a two-part case study: a retrospective documentary "
     "analysis of anonymized occurrences, reinvestigated with a structured "
     "protocol, and a field pilot. The system enables immediate on-site "
     "recording by smartphone, guides the investigator through three mandatory "
     "levels of causes — immediate, underlying and basic —, links corrective and "
     "preventive actions with owners and deadlines, and computes reactive and "
     "proactive indicators, including the near miss-to-accident ratio and the "
     "frequency and severity rates. Complementary modules digitize, on the same "
     "database, work permits, psychosocial risk assessment and preliminary "
     "ergonomic assessment. In the documentary analysis, [PREENCHER: main "
     "result]; in the pilot, [PREENCHER: main result]. It is concluded that "
     "digitizing the record–investigate–act cycle [PREENCHER: confirm or adjust] "
     "deepens causal analysis toward organizational factors, reduces near-miss "
     "underreporting and gives records the traceability required by "
     "occupational risk management, at virtually zero operating cost.",
     indent=False)
body(doc, "Keywords: Occupational accidents. Near misses. Causal tree method. "
     "Accident investigation. Occupational safety. Electric power.",
     indent=False)
doc.add_page_break()

pretextual_title(doc, "LISTA DE ILUSTRAÇÕES")
list_field(doc, 'TOC \\h \\z \\c "Figura"', "Lista de figuras — atualize no Word (Ctrl+A e F9).")
list_field(doc, 'TOC \\h \\z \\c "Quadro"', "Lista de quadros — atualize no Word (Ctrl+A e F9).")
doc.add_page_break()

pretextual_title(doc, "LISTA DE TABELAS")
list_field(doc, 'TOC \\h \\z \\c "Tabela"', "Lista de tabelas — atualize no Word (Ctrl+A e F9).")
doc.add_page_break()

pretextual_title(doc, "LISTA DE ABREVIATURAS E SIGLAS")
SIGLAS = [
    ("ADC", "Árvore de Causas"),
    ("AEP", "Avaliação Ergonômica Preliminar"),
    ("CAT", "Comunicação de Acidente de Trabalho"),
    ("CID-11", "Classificação Internacional de Doenças, 11ª revisão"),
    ("GRO", "Gerenciamento de Riscos Ocupacionais"),
    ("INRS", "Institut National de Recherche et de Sécurité"),
    ("IRO", "Inventário de Riscos Ocupacionais"),
    ("LGPD", "Lei Geral de Proteção de Dados Pessoais"),
    ("MAPA", "Modelo de Análise e Prevenção de Acidentes de Trabalho"),
    ("MTE", "Ministério do Trabalho e Emprego"),
    ("NBR", "Norma Brasileira"),
    ("NR", "Norma Regulamentadora"),
    ("PGR", "Programa de Gerenciamento de Riscos"),
    ("SEP", "Sistema Elétrico de Potência"),
    ("SST", "Segurança e Saúde no Trabalho"),
    ("TF", "Taxa de Frequência"),
    ("TG", "Taxa de Gravidade"),
]
for sig, desc in SIGLAS:
    p = doc.add_paragraph()
    p.paragraph_format.tab_stops.add_tab_stop(Cm(2.5))
    _runs(p, f"{sig}\t{desc}")

doc.add_page_break()
pretextual_title(doc, "SUMÁRIO")
list_field(doc, 'TOC \\o "1-3" \\h \\z \\u',
           "Sumário automático — atualize no Word (Ctrl+A e F9).")

# ===== Seção 3: textuais e pós-textuais (numeradas) ===========================
sec_txt = new_section(doc)
page_number_header(sec_txt)

# ============================================================ 1 INTRODUÇÃO ====
h1(doc, "1 Introdução")
body(doc, "Os acidentes do trabalho permanecem entre os principais problemas de "
     "saúde pública e de gestão nas organizações brasileiras. Os registros "
     "oficiais do Anuário Estatístico de Acidentes do Trabalho contabilizam "
     "[PREENCHER: número de acidentes registrados no AEAT mais recente] "
     "ocorrências no ano-base [PREENCHER: ano] (Brasil, [PREENCHER: ano da "
     "edição]), número que subestima a realidade, pois os eventos sem "
     "afastamento e os quase acidentes raramente chegam às estatísticas "
     "(Almeida, 2006).")
body(doc, "O risco elétrico agrava esse quadro. O anuário da Associação "
     "Brasileira de Conscientização para os Perigos da Eletricidade registrou, "
     "no ano-base 2023, 2.089 acidentes de origem elétrica no país, dos quais "
     "781 resultaram em óbito (Abracopel, 2024). No recorte das redes de "
     "distribuição, levantamento da associação das distribuidoras apontou 250 "
     "mortes envolvendo a rede elétrica no mesmo ano (Acidentes [...], 2024). "
     "Para as equipes que constroem e mantêm essas redes, trabalhando em "
     "condutores energizados ou desenergizados e em altura, o acidente "
     "raramente é leve.")
body(doc, "A Norma Regulamentadora nº 1 (NR-1), na redação vigente, insere a "
     "análise de acidentes no núcleo do Gerenciamento de Riscos Ocupacionais "
     "(GRO): o empregador deve analisar os acidentes e as doenças relacionadas "
     "ao trabalho, identificar suas causas e utilizar os resultados na revisão "
     "do levantamento de perigos e do Programa de Gerenciamento de Riscos (PGR) "
     "(Brasil, 2024). A obrigação, portanto, não termina na emissão da "
     "Comunicação de Acidente de Trabalho (CAT): exige investigação com método, "
     "registro rastreável e realimentação do inventário de riscos.")
body(doc, "Na prática, porém, esse ciclo é frágil: o registro é feito em papel, "
     "horas ou dias após o evento; a investigação, quando ocorre, limita-se a "
     "atribuir o acidente a um “ato inseguro” do trabalhador; os quase "
     "acidentes não são reportados; e as ações corretivas não são acompanhadas "
     "até a conclusão. A literatura brasileira denomina essa prática de "
     "paradigma tradicional ou culpabilizador e a associa à recorrência dos "
     "eventos (Binder; Almeida, 1997; Almeida, 2006). Coloca-se, assim, o "
     "problema desta pesquisa: como estruturar, em uma organização de "
     "distribuição de energia elétrica com equipe de SST enxuta, um processo de "
     "registro e investigação de ocorrências que seja imediato, orientado aos "
     "fatores organizacionais, rastreável para o PGR e de custo compatível com "
     "a realidade da organização?")

h2(doc, "1.1 Objetivo")
body(doc, "O objetivo deste trabalho é desenvolver e aplicar um sistema digital "
     "de registro e investigação de acidentes e quase acidentes, estruturado no "
     "método da árvore de causas e integrado ao Gerenciamento de Riscos "
     "Ocupacionais da NR-1, em uma organização do setor de distribuição de "
     "energia elétrica. Para tanto, são objetivos específicos:")
alineas(doc, [
    "sistematizar os requisitos da NR-1, da ABNT NBR 14280 e da ABNT NBR ISO "
    "45001 aplicáveis à análise de acidentes e sua articulação com a CAT",
    "revisar os modelos de causalidade de acidentes e o método da árvore de "
    "causas, incluindo os fatores humanos e organizacionais contribuintes",
    "especificar e implementar o sistema: registro em campo, investigação em "
    "três níveis de causas, gestão de ações e indicadores",
    "integrar o sistema aos módulos de permissão de trabalho, fatores "
    "psicossociais e avaliação ergonômica preliminar, sobre a mesma base de "
    "dados",
    "aplicar o sistema em estudo de caso, com análise documental retrospectiva "
    "e piloto de campo, e discutir resultados e limitações",
])

h2(doc, "1.2 Justificativa")
body(doc, "A escolha do tema decorre da vivência profissional do autor em "
     "serviços de distribuição de energia elétrica [PREENCHER: confirmar ou "
     "ajustar o vínculo profissional], setor em que a letalidade dos acidentes "
     "elétricos supera um terço das ocorrências (Abracopel, 2024) e em que cada "
     "evento — e cada quase acidente — representa uma oportunidade de "
     "aprendizado cujo desperdício tem custo potencial em vidas.")
body(doc, "Há também uma lacuna prática. A NR-1 obriga a análise de acidentes "
     "como insumo do GRO (Brasil, 2024), e o método da árvore de causas está "
     "consolidado na literatura brasileira desde a década de 1990 (Binder; "
     "Monteau; Almeida, 1995), mas faltam ferramentas digitais acessíveis que o "
     "levem ao campo: soluções comerciais de gestão de ocorrências têm custo "
     "incompatível com organizações menores, e o registro em papel fragiliza a "
     "rastreabilidade. Este trabalho demonstra a viabilidade de uma solução de "
     "custo praticamente nulo, replicável por profissionais de SST sem apoio de "
     "equipe de tecnologia da informação.")

# ================================================ 2 REVISÃO DA LITERATURA =====
new_page = doc.add_page_break
new_page()
h1(doc, "2 Revisão da literatura")
body(doc, "Esta seção reúne as definições, os modelos de causalidade e o marco "
     "normativo que fundamentam o sistema. A revisão combina obras seminais dos "
     "modelos de análise de acidentes com publicações recentes sobre o tema e "
     "com as estatísticas mais atuais do setor elétrico.")

h2(doc, "2.1 Acidente, incidente e quase acidente")
body(doc, "A Lei nº 8.213/1991 define acidente do trabalho, para fins "
     "previdenciários, como o que ocorre pelo exercício do trabalho a serviço "
     "da empresa, provocando lesão corporal ou perturbação funcional que cause "
     "morte, perda ou redução da capacidade para o trabalho (Brasil, 1991). A "
     "ABNT NBR 14280 adota conceito prevencionista mais amplo — ocorrência "
     "imprevista e indesejável relacionada com o exercício do trabalho, de que "
     "resulte ou possa resultar lesão pessoal — e padroniza o cadastro e as "
     "estatísticas de acidentes, incluindo as taxas de frequência e de "
     "gravidade (Associação Brasileira de Normas Técnicas, 2001).")
body(doc, "O quase acidente é o evento que, por circunstâncias fortuitas, não "
     "produziu lesão nem dano, mas cuja dinâmica é a mesma do acidente. A ABNT "
     "NBR ISO 45001 o abrange no conceito de incidente, exigindo seu reporte e "
     "investigação no sistema de gestão (Associação Brasileira de Normas "
     "Técnicas, 2018). Adota-se, neste trabalho, a classificação operacional em "
     "acidente com afastamento, acidente sem afastamento e quase acidente, à "
     "qual o sistema acrescenta o registro de condições inseguras observadas. A "
     "distinção importa porque a definição previdenciária, centrada na lesão, "
     "deixa de fora justamente os eventos que mais informam a prevenção.")

h2(doc, "2.2 Modelos de causalidade de acidentes")
body(doc, "A forma de investigar decorre do modelo causal adotado — por isso a "
     "revisão dos modelos precede a escolha do método. Heinrich (1931) "
     "representou o acidente como sequência linear de fatores, a teoria do "
     "dominó, na qual bastaria remover uma peça, tipicamente o “ato inseguro”, "
     "para interromper a cadeia. Do mesmo autor vem a pirâmide que relaciona "
     "grandes números de incidentes menores a cada lesão grave; Bird e Germain "
     "(1985) a atualizaram e estenderam aos danos materiais, fundamentando a "
     "prática de reportar e tratar quase acidentes como matéria-prima da "
     "prevenção (Figura 1).")
figura(doc, "Pirâmide de eventos: relação entre lesões graves, lesões leves, "
       "danos materiais e quase acidentes",
       "inserir diagrama da pirâmide redesenhado pelo autor",
       "Adaptado de Bird e Germain (1985).")
body(doc, "Os modelos lineares mostraram-se insuficientes para explicar "
     "acidentes em sistemas complexos. Reason (1990; 1997) distinguiu falhas "
     "ativas — atos na ponta operacional — de condições latentes — decisões de "
     "projeto, gestão e organização que permanecem adormecidas até se alinharem "
     "às falhas ativas, na representação conhecida como modelo do queijo suíço. "
     "A consequência prática é direta: investigar apenas a conduta do "
     "trabalhador deixa intactas as condições latentes, que voltarão a produzir "
     "eventos.")
body(doc, "Os modelos sistêmicos aprofundaram essa leitura. Leveson (2011) "
     "propõe tratar a segurança como problema de controle: acidentes resultam "
     "de restrições de segurança inadequadamente impostas ao comportamento do "
     "sistema, e não de uma cadeia de falhas de componentes. Nessa perspectiva, "
     "perguntar por que o controle falhou — procedimento inexistente, "
     "realimentação ausente, planejamento inadequado — é mais produtivo do que "
     "perguntar quem errou. A perspectiva de Hollnagel (2014), denominada "
     "Safety-II, desloca ainda o foco do que falha para a variabilidade normal "
     "do trabalho real, reforçando o valor de aprender também com o cotidiano — "
     "o que dá suporte conceitual ao registro de quase acidentes e de condições "
     "inseguras.")
body(doc, "A sucessão de modelos não torna os anteriores inúteis, mas desloca "
     "o objeto da investigação: dos atos individuais (Heinrich, 1931) para as "
     "condições latentes (Reason, 1997) e para as estruturas de controle "
     "(Leveson, 2011). Para uma organização com equipe de SST enxuta, o desafio "
     "é operacionalizar essa evolução em um instrumento de uso cotidiano, sem "
     "exigir do investigador domínio teórico aprofundado — o que justifica a "
     "escolha de um método estruturado e já difundido no Brasil, apresentado a "
     "seguir.")

h2(doc, "2.3 O método da árvore de causas")
body(doc, "O método da árvore de causas, desenvolvido no Institut National de "
     "Recherche et de Sécurité (INRS) francês na década de 1970, parte do "
     "acidente consumado e reconstrói, de trás para a frente, a rede de fatos "
     "que o produziram, perguntando sistematicamente o que foi necessário para "
     "que cada fato ocorresse e se ele foi suficiente. O resultado é um "
     "diagrama que explicita as combinações de variações do trabalho habitual "
     "que culminaram no evento, sem juízo de culpa (Binder; Monteau; Almeida, "
     "1995). A Figura 2 ilustra a estrutura do diagrama.")
body(doc, "A aplicação do método percorre três etapas (Binder; Monteau; "
     "Almeida, 1995). Na primeira, a coleta de informações, o investigador "
     "levanta, preferencialmente no local e logo após o evento, os fatos que "
     "diferiram do trabalho habitual — as variações —, registrando fatos "
     "objetivos e não interpretações ou juízos de valor. Na segunda, a "
     "construção da árvore, os fatos são encadeados a partir da lesão, "
     "identificando três tipos de relação lógica: o encadeamento, quando um "
     "único fato antecedente foi necessário e suficiente para produzir o "
     "seguinte; a conjunção, quando dois ou mais antecedentes foram necessários "
     "simultaneamente; e a disjunção, quando um mesmo fato deu origem a dois ou "
     "mais consequentes independentes. Na terceira, a exploração da árvore, o "
     "grupo de análise propõe medidas para neutralizar cada fator e, sobretudo, "
     "identifica os fatores potenciais de acidente — condições que, por "
     "estarem presentes em outras situações de trabalho, podem gerar novos "
     "eventos.")
body(doc, "A etapa de coleta é a mais sensível ao tempo: depoimentos, posições "
     "de ferramentas e condições do local se perdem rapidamente, e a "
     "reconstituição tardia tende a preencher lacunas com suposições — "
     "frequentemente a de que o trabalhador “não seguiu o procedimento” "
     "(Almeida, 2006). Essa constatação fundamenta uma decisão central deste "
     "trabalho: o registro deve acontecer no local, no momento do evento, com o "
     "recurso que o trabalhador já tem em mãos — o celular.")
figura(doc, "Estrutura de um diagrama de árvore de causas",
       "inserir exemplo genérico de árvore de causas redesenhado pelo autor",
       "Adaptado de Binder, Monteau e Almeida (1995).")
body(doc, "No Brasil, o método foi difundido por Binder e Almeida (1997), que "
     "demonstraram em estudos de caso sua capacidade de revelar o papel de "
     "fatores gerenciais e de organização do trabalho na gênese dos acidentes — "
     "designação improvisada de trabalhadores a funções, execução deixada à "
     "iniciativa individual, falta de ferramentas adequadas e falhas na "
     "circulação de informações. Almeida (2006) aprofundou a crítica à "
     "atribuição de culpa como paradigma dominante nas empresas brasileiras, "
     "mostrando que investigações centradas no comportamento do acidentado "
     "produzem recomendações inócuas — treinar, advertir, punir — e deixam "
     "aberta a porta à recorrência.")
body(doc, "Essa linha evoluiu para o Modelo de Análise e Prevenção de Acidentes "
     "de Trabalho (MAPA), que integra a análise da atividade, a análise de "
     "barreiras e os conceitos de Reason em um roteiro voltado à vigilância em "
     "saúde do trabalhador (Almeida; Vilela, 2010; Almeida et al., 2014). "
     "Publicação recente que sistematiza quarenta anos de trajetória brasileira "
     "em análise de acidentes e desastres reafirma a passagem do paradigma "
     "culpabilizador para abordagens sistêmicas e participativas como eixo da "
     "prevenção (Porto, 2024).")
body(doc, "O sistema desenvolvido neste trabalho operacionaliza a lógica da "
     "árvore de causas em três níveis, terminologia corrente na prática de SST: "
     "causas imediatas (atos e condições no momento do evento), causas "
     "subjacentes (fatores da tarefa e do posto que tornaram possíveis as "
     "imediatas) e causas básicas (fatores de gestão e organização que "
     "originaram as subjacentes). A estrutura em níveis impede o investigador "
     "de encerrar a análise na primeira resposta — mecanismo digital "
     "equivalente à pergunta iterativa do método.")

h2(doc, "2.4 Fatores humanos e organizacionais: fadiga, estresse e burnout")
body(doc, "A investigação que chega às causas básicas encontra, com "
     "frequência, fatores humanos e organizacionais: jornadas extensas, pressão "
     "de tempo, sobreaviso, fadiga e estados de esgotamento que degradam a "
     "atenção e a tomada de decisão (Reason, 1997). A NR-1, na redação dada "
     "pela Portaria MTE nº 1.419/2024, passou a exigir a inclusão dos fatores "
     "de risco psicossocial no GRO (Brasil, 2024) — exigência que conversa "
     "diretamente com a análise de acidentes, pois os mesmos fatores que "
     "adoecem contribuem para eventos agudos.")
body(doc, "Entre os desfechos do estresse ocupacional crônico, a síndrome de "
     "burnout — caracterizada por exaustão, distanciamento mental do trabalho e "
     "redução da eficácia profissional — foi incluída na CID-11, sob o código "
     "QD85, como fenômeno ocupacional (Organização Mundial da Saúde, 2019), com "
     "adoção oficial da classificação no Brasil a partir de 2025. Estudo "
     "epidemiológico nacional indica tendência de crescimento das notificações "
     "entre 2014 e 2024 (Treml et al., 2025). No setor elétrico, Souza et al. "
     "(2010) encontraram prevalência de 20,3% de transtornos mentais comuns em "
     "eletricitários, associada a alta demanda, baixo controle e baixo apoio "
     "social. Para este trabalho, a implicação é dupla: o roteiro de "
     "investigação inclui fatores humanos e organizacionais entre as causas "
     "básicas selecionáveis, e a plataforma mantém módulo de avaliação "
     "psicossocial que permite cruzar a sinalização coletiva de risco com a "
     "ocorrência de eventos.")

h2(doc, "2.5 Indicadores de desempenho em segurança")
body(doc, "A ABNT NBR 14280 padroniza os indicadores reativos clássicos: a "
     "taxa de frequência, expressa em acidentes por milhão de horas-homem de "
     "exposição, e a taxa de gravidade, expressa em dias perdidos e debitados "
     "por milhão de horas-homem (Associação Brasileira de Normas Técnicas, "
     "2001), calculadas pelas Equações 1 e 2:")
equacao(doc, "TF = (N × 1.000.000) / HHT", 1)
equacao(doc, "TG = [(DP + DD) × 1.000.000] / HHT", 2)
body(doc, "em que N é o número de acidentes com lesão no período, HHT são as "
     "horas-homem de exposição ao risco, DP são os dias perdidos em razão dos "
     "afastamentos e DD são os dias debitados, atribuídos pela norma aos casos "
     "de morte e de incapacidade permanente.", indent=False)
body(doc, "Por medirem o que já ocorreu, esses indicadores são insuficientes "
     "para gerir eventos raros e graves; a literatura recomenda complementá-los "
     "com indicadores proativos, como quase acidentes reportados, ações "
     "concluídas no prazo e tempo entre evento e registro (Hollnagel, 2014). A "
     "razão entre quase acidentes e acidentes funciona como termômetro da "
     "cultura de reporte: valores baixos indicam subnotificação, e não "
     "segurança (Bird; Germain, 1985).")

h2(doc, "2.6 Marco normativo")
body(doc, "O arcabouço normativo articula a base legal, a camada "
     "regulamentar, a camada técnica e a camada setorial, sintetizadas no "
     "Quadro 1, que relaciona cada instrumento à funcionalidade correspondente "
     "do sistema desenvolvido. A Lei nº 8.213/1991 obriga a emissão da CAT até "
     "o primeiro dia útil seguinte à ocorrência e imediatamente em caso de "
     "óbito (Brasil, 1991); a NR-1 exige a análise de acidentes com "
     "realimentação do PGR (Brasil, 2024); a NR-10 e a NR-35 regem as "
     "atividades típicas da distribuição de energia (Brasil, 2019; Brasil, "
     "[PREENCHER: ano da NR-35 consultada]).")
tabela(doc, "Quadro", "Marco normativo aplicável à análise de acidentes e "
       "funcionalidade correspondente do sistema",
       ["Instrumento", "Exigência principal", "Funcionalidade no sistema"],
       [["Lei nº 8.213/1991", "Define acidente do trabalho; obriga a CAT",
         "Alerta de emissão da CAT no registro"],
        ["NR-1 (Portaria MTE nº 1.419/2024)",
         "Análise de acidentes e revisão do PGR",
         "Investigação em três níveis e vínculo ao plano de ação"],
        ["ABNT NBR 14280:2001", "Cadastro, classificação, TF e TG",
         "Classificação da ocorrência e indicadores"],
        ["ABNT NBR ISO 45001:2018", "Reporte e investigação de incidentes",
         "Registro de quase acidentes e condições inseguras"],
        ["NR-10 e NR-35", "Eletricidade e trabalho em altura",
         "Módulo de permissão de trabalho"]],
       "Elaborado pelo autor com base em Brasil (1991; 2019; 2024) e "
       "Associação Brasileira de Normas Técnicas (2001; 2018).",
       widths=[4.5, 5.5, 6.0])

h2(doc, "2.7 Acidentalidade no setor de distribuição de energia elétrica")
body(doc, "O setor elétrico apresenta perfil de acidentalidade de alta "
     "gravidade: das 2.089 ocorrências de origem elétrica registradas no "
     "ano-base 2023, 781 foram fatais, e as áreas de geração e distribuição "
     "concentram parcela expressiva dos choques fatais (Abracopel, 2024). No "
     "recorte das distribuidoras, foram 250 mortes envolvendo a rede elétrica em "
     "2023, queda de cerca de 8% em relação ao ano anterior (Acidentes [...], "
     "2024). [PREENCHER: se disponível, acrescentar estatística de "
     "trabalhadores próprios e terceirizados do setor, com fonte e ano.] Além "
     "do choque elétrico e do arco voltaico, as equipes de linha estão expostas "
     "à queda de altura, ao trânsito e a intempéries, e a literatura setorial "
     "associa o adoecimento e a acidentalidade a fatores organizacionais como "
     "alta demanda e baixo controle (Souza et al., 2010). Esse perfil — eventos "
     "raros e graves — torna o aprendizado com quase acidentes particularmente "
     "valioso, pois eles fornecem o volume de sinais que os acidentes, menos "
     "numerosos, não fornecem (Bird; Germain, 1985).")
body(doc, "Nos serviços em redes desenergizadas, a NR-10 estabelece que somente "
     "são consideradas desenergizadas as instalações liberadas para trabalho "
     "mediante a seguinte sequência: seccionamento; impedimento de "
     "reenergização; constatação da ausência de tensão; instalação de "
     "aterramento temporário com equipotencialização dos condutores dos "
     "circuitos; proteção dos elementos energizados existentes na zona "
     "controlada; e instalação da sinalização de impedimento de reenergização "
     "(Brasil, 2019). Cada etapa é uma barreira, no sentido de Reason (1997): a "
     "falha em qualquer uma delas — um aterramento omitido, uma constatação de "
     "ausência de tensão feita no ponto errado, uma reenergização por manobra "
     "de terceiros — é causa imediata típica de acidentes graves no setor. "
     "Perguntar por que a barreira falhou conduz às causas subjacentes e "
     "básicas: material de aterramento indisponível na viatura, procedimento "
     "que não contempla a configuração real da rede, comunicação deficiente "
     "com o centro de operação, prazo de restabelecimento incompatível com a "
     "execução completa da sequência.")
body(doc, "Esse encadeamento mostra por que os requisitos da NR-10, embora "
     "detalhados, não bastam por si: o cumprimento das etapas depende de "
     "condições organizacionais — planejamento, recursos, comunicação, "
     "dimensionamento de equipes — que só a investigação estruturada torna "
     "visíveis (Binder; Almeida, 1997). Por isso o sistema articula a "
     "permissão de trabalho, que verifica as barreiras antes do serviço, com a "
     "investigação de ocorrências, que analisa por que elas falharam.")

h2(doc, "2.8 Digitalização do registro e da investigação de ocorrências")
body(doc, "O processo tradicional — formulário de papel preenchido após o "
     "retorno à base e transcrito depois para planilha — introduz perda de "
     "informação, atraso e inconsistência de classificação, fragilidades que a "
     "literatura associa à baixa qualidade das investigações (Almeida, 2006). O "
     "registro no local, em dispositivo móvel, endereça as três fragilidades: "
     "captura imediata de dados e evidências, classificação guiada e base única "
     "para indicadores e auditoria. A opção deste trabalho por tecnologias "
     "gratuitas, sem instalação e operáveis por equipes de SST sem apoio de TI "
     "visa remover a barreira de custo que mantém o processo manual, liberando "
     "esforço para o que efetivamente previne: a investigação bem-feita e a "
     "execução das ações (Porto, 2024).")

h2(doc, "2.9 Estudos correlatos e lacuna")
body(doc, "A produção brasileira sobre análise de acidentes consolidou-se em "
     "torno da árvore de causas e da crítica ao paradigma culpabilizador "
     "(Binder; Almeida, 1997; Almeida, 2006), evoluiu para modelos sistêmicos "
     "como o MAPA (Almeida et al., 2014) e foi recentemente revisitada em "
     "perspectiva histórica (Porto, 2024). [PREENCHER: acrescentar dois ou três "
     "estudos de 2021 em diante sobre investigação de acidentes ou quase "
     "acidentes — busca sugerida no SciELO e na Revista Brasileira de Saúde "
     "Ocupacional com os termos “análise de acidentes”, “quase acidente” e "
     "“near miss”.] Observa-se escassez de propostas que levem o método "
     "estruturado de investigação ao campo em ferramenta digital gratuita, "
     "acessível a organizações com equipes enxutas — lacuna que este trabalho "
     "pretende ajudar a preencher.")
body(doc, "Em síntese, três conclusões da revisão orientam o desenvolvimento: "
     "o modelo causal adotado determina a qualidade da investigação, e a árvore "
     "de causas é o antídoto metodológico consolidado ao paradigma "
     "culpabilizador (Binder; Almeida, 1997); o quase acidente é o insumo mais "
     "abundante do aprendizado, mas exige reporte fácil e imediato (Bird; "
     "Germain, 1985); e a exigência da NR-1 de analisar ocorrências e "
     "realimentar o PGR demanda registro rastreável (Brasil, 2024). Esses "
     "pontos definem os requisitos do sistema apresentado na seção 3.")

# ========================================================= 3 METODOLOGIA =====
new_page()
h1(doc, "3 Materiais e métodos")

h2(doc, "3.1 Caracterização da pesquisa")
body(doc, "Trata-se de pesquisa aplicada, de natureza tecnológica, conduzida na "
     "forma de desenvolvimento de artefato — o sistema digital — seguido de "
     "estudo de caso. A abordagem é quali-quantitativa: quantitativa no "
     "tratamento dos registros e indicadores e qualitativa na análise das "
     "investigações e da adequação do processo ao contexto. O estudo de caso "
     "foi conduzido em uma organização do setor de distribuição de energia "
     "elétrica, caracterizada de forma a não permitir identificação: "
     "[PREENCHER: porte, região de atuação e natureza das atividades, sem nome "
     "nem dados identificáveis].")

h2(doc, "3.2 Etapas da pesquisa")
body(doc, "O trabalho foi desenvolvido nas etapas a seguir, em ordem "
     "cronológica, de modo que cada uma fornecesse subsídios à seguinte:")
alineas(doc, [
    "revisão bibliográfica e normativa: modelos de causalidade, método da "
    "árvore de causas, indicadores e requisitos da NR-1, ABNT NBR 14280, ABNT "
    "NBR ISO 45001, NR-10 e NR-35",
    "especificação do sistema: taxonomia de ocorrências, fluxo de registro em "
    "campo, roteiro de investigação em três níveis de causas, gestão de ações e "
    "indicadores",
    "projeto da arquitetura: premissas de custo zero, ausência de instalação e "
    "operação em dispositivo móvel",
    "desenvolvimento do módulo central de ocorrências e dos módulos "
    "complementares sobre a mesma base de dados",
    "testes de funcionamento, integridade dos dados e usabilidade em campo",
    "estudo de caso, parte documental: reinvestigação de ocorrências "
    "históricas anonimizadas com o roteiro de três níveis",
    "estudo de caso, parte de campo: uso piloto do sistema pelas equipes",
    "análise dos resultados, discussão à luz da literatura e proposição de "
    "plano de ação",
])

h2(doc, "3.3 Aspectos éticos e anonimato")
body(doc, "O estudo não expõe a organização nem pessoas. Os casos são descritos "
     "de forma anonimizada — função genérica do envolvido, sem nome, data "
     "exata, local ou número de CAT —, e nas fotografias o rosto dos "
     "trabalhadores e a identificação da empresa em uniformes, veículos e "
     "equipamentos foram ocultados. O tratamento dos registros observa a Lei nº "
     "13.709/2018 (Brasil, 2018): os dados são tratados com finalidade de "
     "prevenção, com acesso restrito, e as análises são apresentadas de forma "
     "agregada.")

h2(doc, "3.4 Materiais: arquitetura da plataforma")
body(doc, "O sistema é o módulo central da plataforma Psike, construída sobre "
     "três decisões de projeto: custo zero de operação, ausência de instalação "
     "e uso direto no celular do usuário. O frontend é um único arquivo web "
     "executado no navegador; o backend é um serviço gratuito em nuvem (Google "
     "Apps Script) que grava os registros em planilha eletrônica, uma aba por "
     "módulo, servindo de repositório único (Figura 3). Os quatro módulos — "
     "ocorrências, permissão de trabalho, fatores psicossociais e avaliação "
     "ergonômica preliminar — compartilham a mesma base, o que permite cruzar "
     "ocorrências com permissões emitidas e com a sinalização coletiva de risco "
     "psicossocial.")
figura(doc, "Arquitetura da plataforma Psike",
       "inserir diagrama de arquitetura: navegador/celular, backend e planilha",
       "Elaborado pelo autor (2026).")

h2(doc, "3.5 Métodos: o módulo de registro e investigação")
h3(doc, "3.5.1 Registro em campo")
body(doc, "O registro foi projetado para ser feito no local, em menos de cinco "
     "minutos, pelo encarregado ou pelo técnico de segurança: classificação do "
     "evento (acidente com afastamento, acidente sem afastamento, quase "
     "acidente ou condição insegura), data e hora, função genérica do "
     "envolvido, tarefa em execução, descrição do fato e parte do corpo "
     "atingida, quando houver lesão. O sistema alerta para a obrigação legal de "
     "emissão da CAT nos casos aplicáveis, deixando explícito que o registro "
     "interno não a substitui (Brasil, 1991). O Quadro 2 apresenta a "
     "taxonomia adotada e o tratamento previsto para cada tipo de ocorrência.")
tabela(doc, "Quadro", "Taxonomia de ocorrências adotada no sistema",
       ["Tipo", "Definição operacional", "Investigação", "CAT"],
       [["Acidente com afastamento",
         "Lesão que impede o retorno ao trabalho no dia seguinte",
         "Obrigatória, três níveis", "Sim"],
        ["Acidente sem afastamento",
         "Lesão com retorno ao trabalho no dia seguinte",
         "Obrigatória, três níveis", "Sim"],
        ["Quase acidente",
         "Evento sem lesão nem dano, com potencial de causá-los",
         "Recomendada, três níveis", "Não"],
        ["Condição insegura",
         "Situação observada com potencial de gerar ocorrência",
         "Registro e ação corretiva", "Não"]],
       "Elaborado pelo autor com base em Associação Brasileira de Normas "
       "Técnicas (2001; 2018) e Brasil (1991).",
       widths=[3.6, 6.2, 3.8, 1.6])
h3(doc, "3.5.2 Investigação em três níveis de causas")
body(doc, "A investigação aplica a lógica da árvore de causas (Binder; Monteau; "
     "Almeida, 1995) em três níveis sequenciais e obrigatórios (Figura 4): "
     "causas imediatas, que no momento do evento produziram o dano; causas "
     "subjacentes, que na tarefa e no posto tornaram possíveis as imediatas — "
     "ferramenta inadequada, procedimento inexistente ou inaplicável, "
     "improvisação, pressa; e causas básicas, que na gestão e na organização "
     "originaram as subjacentes — dimensionamento de equipe, planejamento do "
     "serviço, capacitação, manutenção e fatores humanos e organizacionais, "
     "incluindo fadiga e sobrecarga (Reason, 1997). O formulário não permite "
     "concluir a investigação apenas com causas imediatas. O Quadro 3 ilustra "
     "a progressão com exemplos típicos de serviços em rede de distribuição.")
tabela(doc, "Quadro", "Exemplos de causas por nível em serviços em rede de "
       "distribuição desenergizada",
       ["Nível", "Pergunta orientadora", "Exemplos"],
       [["Imediatas", "O que produziu o dano na cena?",
         "Aterramento temporário não instalado; constatação de ausência de "
         "tensão em ponto inadequado"],
        ["Subjacentes", "O que na tarefa e no posto tornou isso possível?",
         "Conjunto de aterramento indisponível na viatura; procedimento que não "
         "contempla a configuração real da rede"],
        ["Básicas", "O que na gestão e na organização originou isso?",
         "Planejamento com prazo incompatível com a sequência completa; "
         "inspeção de materiais inexistente; equipe subdimensionada; fadiga "
         "por sobreaviso"]],
       "Elaborado pelo autor com base em Binder, Monteau e Almeida (1995), "
       "Reason (1997) e Brasil (2019).",
       widths=[2.6, 5.0, 7.6])
figura(doc, "Fluxo de registro e investigação em três níveis de causas",
       "inserir fluxograma: registro → causas imediatas → subjacentes → "
       "básicas → ações → indicadores",
       "Elaborado pelo autor (2026).")
h3(doc, "3.5.3 Ações e indicadores")
body(doc, "Cada investigação vincula ações corretivas e preventivas com "
     "responsável e prazo, cujo status é acompanhado no painel. O sistema "
     "calcula automaticamente a contagem de ocorrências por tipo e período, a "
     "razão entre quase acidentes e acidentes, o tempo mediano entre evento e "
     "registro, o percentual de ações concluídas no prazo e, alimentado com as "
     "horas-homem de exposição, as taxas de frequência e de gravidade "
     "(Associação Brasileira de Normas Técnicas, 2001).")

h2(doc, "3.6 Módulos complementares")
body(doc, "A permissão de trabalho digital apresenta checklist específico por "
     "tipo de serviço — rede desenergizada, com a sequência de desenergização "
     "da NR-10; rede energizada; e trabalho em altura —, com captura de "
     "geolocalização e assinatura em tela (Brasil, 2019). A avaliação de "
     "fatores psicossociais aplica instrumento anônimo de dez itens em seis "
     "dimensões, com classificação automática, atendendo à exigência da NR-1 "
     "(Brasil, 2024). A avaliação ergonômica preliminar emprega formulário de "
     "doze itens em cinco blocos com parecer automático. A integração permite "
     "verificar, por exemplo, se um acidente ocorreu em serviço coberto por "
     "permissão e se o setor envolvido já sinalizava sobrecarga.")

h2(doc, "3.7 Estudo de caso: local, população e procedimentos")
body(doc, "Na parte documental, [PREENCHER: número] ocorrências registradas "
     "pela organização no período de [PREENCHER: período] foram anonimizadas e "
     "reinvestigadas com o roteiro de três níveis, permitindo comparar a "
     "profundidade causal da análise original com a da análise estruturada. Na "
     "parte de campo, no período de [PREENCHER: período], as equipes de "
     "[PREENCHER: escopo e número aproximado de trabalhadores] utilizaram o "
     "sistema para registrar novas ocorrências e condições inseguras, após "
     "[PREENCHER: forma de treinamento]. A Figura 5 ilustra a atividade típica "
     "das equipes estudadas.")
figura(doc, "Atividade típica de manutenção em rede de distribuição "
       "(identificações ocultadas)",
       "inserir foto do estudo de caso com rostos e logotipos ocultados",
       "Arquivo do autor (2026).")
body(doc, "Os dados foram tratados de forma agregada. A análise compara a "
     "distribuição de causas por nível na análise original e na estruturada, o "
     "volume e a proporção de quase acidentes reportados antes e durante o "
     "piloto e os tempos entre evento e registro. Quando o número de casos não "
     "permite estatística inferencial, os resultados são apresentados de forma "
     "descritiva.")

# ============================================= 4 RESULTADOS E DISCUSSÃO =====
new_page()
h1(doc, "4 Resultados e discussão")
body(doc, "Os resultados são apresentados em ordem cronológica de obtenção e "
     "discutidos à medida que são apresentados, à luz da literatura revisada "
     "na seção 2.")

h2(doc, "4.1 O sistema desenvolvido")
body(doc, "O sistema foi implementado integralmente e encontra-se "
     "operacional (Figura 6). O ciclo completo — registro no local, "
     "investigação em três níveis, vinculação de ações e atualização dos "
     "indicadores — ocorre sem transcrição manual e a custo de operação nulo. "
     "Esse resultado responde à fragilidade apontada por Almeida (2006) no "
     "processo tradicional: a perda de informação entre o evento e o registro. "
     "Ao exigir a progressão até as causas básicas, o formulário incorpora ao "
     "software a pergunta iterativa da árvore de causas (Binder; Monteau; "
     "Almeida, 1995) e o deslocamento das falhas ativas para as condições "
     "latentes proposto por Reason (1997).")
figura(doc, "Telas do módulo de registro e investigação de ocorrências",
       "inserir capturas de tela do sistema, sem dados pessoais",
       "Elaborado pelo autor (2026).")

h2(doc, "4.2 Análise documental retrospectiva")
body(doc, "A Tabela 1 compara a distribuição das causas identificadas nas "
     "análises originais da organização e nas reinvestigações estruturadas.")
tabela(doc, "Tabela", "Causas identificadas por nível: análise original e "
       "análise estruturada",
       ["Nível de causa", "Análise original (n)", "Análise estruturada (n)"],
       [["Imediatas", "[PREENCHER]", "[PREENCHER]"],
        ["Subjacentes", "[PREENCHER]", "[PREENCHER]"],
        ["Básicas", "[PREENCHER]", "[PREENCHER]"],
        ["Total", "[PREENCHER]", "[PREENCHER]"]],
       "Elaborado pelo autor (2026), com dados anonimizados da organização "
       "estudada.", widths=[5.0, 5.0, 5.0])
body(doc, "[PREENCHER: descrever os resultados da Tabela 1 e dois ou três "
     "casos ilustrativos anonimizados, mostrando as causas básicas que a "
     "análise original não havia alcançado.] [PREENCHER: discutir, conforme os "
     "dados: se a análise estruturada revelou mais causas subjacentes e "
     "básicas, confirma-se no caso concreto o diagnóstico de Binder e Almeida "
     "(1997) e de Almeida (2006) — o instrumento condiciona a profundidade da "
     "análise, e roteiros que aceitam o “ato inseguro” como resposta final "
     "produzem prevenção inócua; se não revelou, discutir as razões, como a "
     "qualidade dos registros históricos.]")

h2(doc, "4.3 Piloto de campo")
body(doc, "A Tabela 2 reúne os indicadores do período piloto, e a Figura 7 "
     "apresenta sua evolução no painel do sistema.")
tabela(doc, "Tabela", "Indicadores do período piloto",
       ["Indicador", "Antes do piloto", "Durante o piloto"],
       [["Acidentes com afastamento", "[PREENCHER]", "[PREENCHER]"],
        ["Acidentes sem afastamento", "[PREENCHER]", "[PREENCHER]"],
        ["Quase acidentes reportados", "[PREENCHER]", "[PREENCHER]"],
        ["Condições inseguras registradas", "[PREENCHER]", "[PREENCHER]"],
        ["Razão quase acidentes / acidentes", "[PREENCHER]", "[PREENCHER]"],
        ["Tempo mediano evento–registro", "[PREENCHER]", "[PREENCHER]"],
        ["Ações concluídas no prazo (%)", "[PREENCHER]", "[PREENCHER]"]],
       "Elaborado pelo autor (2026), com dados anonimizados da organização "
       "estudada.", widths=[6.5, 4.0, 4.0])
figura(doc, "Indicadores do piloto no painel do sistema",
       "inserir gráfico do painel com os dados reais",
       "Elaborado pelo autor (2026).")
body(doc, "[PREENCHER: descrever os resultados da Tabela 2 e da Figura 7.] "
     "[PREENCHER: discutir, conforme os dados: se o reporte de quase acidentes "
     "aumentou, corrobora-se que a subnotificação era sobretudo função do "
     "atrito do processo — na linha da pirâmide de Heinrich (1931) e de Bird e "
     "Germain (1985), o sistema passa a capturar a base de sinais antes "
     "invisível; se a razão permaneceu baixa, discutir barreiras culturais ao "
     "reporte (Hollnagel, 2014).] A presença de fatores humanos e "
     "organizacionais — fadiga, pressão de tempo, sobrecarga — entre as causas "
     "básicas [PREENCHER: confirmar com os dados] articula a análise de "
     "acidentes com a gestão de fatores psicossociais exigida pela NR-1 e com o "
     "quadro de adoecimento mental ocupacional em crescimento no país (Treml et "
     "al., 2025), indicando que prevenção de acidentes e saúde mental são faces "
     "do mesmo sistema de gestão.")

h2(doc, "4.4 Integração ao PGR e plano de ação")
body(doc, "As causas básicas consolidadas realimentam o inventário de riscos: "
     "cada fator organizacional recorrente identificado nas investigações é "
     "candidato a perigo a reavaliar no inventário, e as ações vinculadas "
     "compõem o plano de ação do PGR com responsáveis e prazos — o ciclo que a "
     "NR-1 exige (Brasil, 2024). [PREENCHER: descrever as revisões do "
     "inventário e o plano de ação efetivamente decorrentes do estudo, sem "
     "identificar a organização.]")

h2(doc, "4.5 Limitações")
body(doc, "O roteiro em três níveis é uma operacionalização simplificada da "
     "árvore de causas: estrutura a progressão causal, mas não desenha o "
     "diagrama completo de combinações do método original (Binder; Monteau; "
     "Almeida, 1995). A comparação retrospectiva depende da qualidade dos "
     "registros históricos; o piloto restringe-se a uma organização e a um "
     "período curto, o que limita a generalização e a leitura de tendências; e "
     "a versão atual não possui autenticação nem segregação de dados por "
     "organização, requisitos para uso em escala.")

# ========================================================== 5 CONCLUSÕES =====
new_page()
h1(doc, "5 Conclusões")
body(doc, "O objetivo deste trabalho — desenvolver e aplicar um sistema digital "
     "de registro e investigação de acidentes e quase acidentes pelo método da "
     "árvore de causas, integrado ao Gerenciamento de Riscos Ocupacionais, em "
     "uma organização de distribuição de energia elétrica — foi [PREENCHER: "
     "atingido / parcialmente atingido / não atingido]. O sistema foi "
     "desenvolvido, está operacional e foi aplicado em estudo de caso "
     "documental e piloto de campo. [PREENCHER: uma ou duas frases sobre o "
     "principal achado e a resposta ao problema de pesquisa.]")
body(doc, "Como trabalhos futuros, sugerem-se a implementação do diagrama "
     "completo da árvore de causas no software, a autenticação e segregação "
     "por organização, o acompanhamento longitudinal dos indicadores e o "
     "cruzamento analítico entre ocorrências, permissões de trabalho e "
     "sinalização psicossocial.")

# ============================================================ REFERÊNCIAS ====
new_page()
h1_center(doc, "Referências")
REFS = [
    "ABRACOPEL. **Anuário estatístico de acidentes de origem elétrica 2024**: "
    "ano base 2023. [PREENCHER: local]: Abracopel, 2024. Disponível em: "
    "[PREENCHER: URL]. Acesso em: [PREENCHER: data].",
    "ACIDENTES fatais com a rede elétrica caem 8% em 2023, aponta Abradee. "
    "**Agência Brasil**, Brasília, DF, jul. 2024. Disponível em: "
    "https://agenciabrasil.ebc.com.br/geral/noticia/2024-07/acidentes-fatais-"
    "com-rede-eletrica-caem-8-em-2023-aponta-abradee. Acesso em: [PREENCHER: "
    "data].",
    "ALMEIDA, I. M. Trajetória da análise de acidentes: o paradigma "
    "tradicional e os primórdios da ampliação da análise. **Interface**: "
    "comunicação, saúde, educação, Botucatu, v. 10, n. 19, p. 185-202, 2006. "
    "DOI: 10.1590/S1414-32832006000100013.",
    "ALMEIDA, I. M.; VILELA, R. A. G. **Modelo de análise e prevenção de "
    "acidentes de trabalho**: MAPA. Piracicaba: CEREST, 2010.",
    "ALMEIDA, I. M. et al. Modelo de Análise e Prevenção de Acidentes - MAPA: "
    "ferramenta para a vigilância em saúde do trabalhador. **Ciência & Saúde "
    "Coletiva**, Rio de Janeiro, v. 19, n. 12, p. 4679-4688, 2014.",
    "ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. **NBR 14280**: cadastro de "
    "acidente do trabalho: procedimento e classificação. Rio de Janeiro: ABNT, "
    "2001.",
    "ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. **NBR ISO 45001**: sistemas de "
    "gestão de saúde e segurança ocupacional: requisitos com orientação para "
    "uso. Rio de Janeiro: ABNT, 2018.",
    "BINDER, M. C. P.; ALMEIDA, I. M. Estudo de caso de dois acidentes do "
    "trabalho investigados com o método de árvore de causas. **Cadernos de "
    "Saúde Pública**, Rio de Janeiro, v. 13, n. 4, p. 749-760, 1997.",
    "BINDER, M. C. P.; MONTEAU, M.; ALMEIDA, I. M. **Árvore de causas**: "
    "método de investigação de acidentes de trabalho. São Paulo: Publisher "
    "Brasil, 1995. [PREENCHER: conferir ano da edição consultada].",
    "BIRD, F. E.; GERMAIN, G. L. **Practical loss control leadership**. "
    "Loganville: International Loss Control Institute, 1985.",
    "BRASIL. Lei nº 8.213, de 24 de julho de 1991. Dispõe sobre os Planos de "
    "Benefícios da Previdência Social e dá outras providências. **Diário "
    "Oficial da União**: seção 1, Brasília, DF, 25 jul. 1991.",
    "BRASIL. Lei nº 13.709, de 14 de agosto de 2018. Lei Geral de Proteção de "
    "Dados Pessoais (LGPD). **Diário Oficial da União**: seção 1, Brasília, "
    "DF, 15 ago. 2018.",
    "BRASIL. Ministério da Previdência Social. **Anuário estatístico de "
    "acidentes do trabalho**: AEAT [PREENCHER: ano-base]. Brasília, DF: MPS, "
    "[PREENCHER: ano]. Disponível em: [PREENCHER: URL]. Acesso em: "
    "[PREENCHER: data].",
    "BRASIL. Ministério do Trabalho e Emprego. **NR-1**: disposições gerais e "
    "gerenciamento de riscos ocupacionais. Redação dada pela Portaria MTE nº "
    "1.419, de 27 de agosto de 2024. Brasília, DF: MTE, 2024.",
    "BRASIL. Ministério do Trabalho e Emprego. **NR-10**: segurança em "
    "instalações e serviços em eletricidade. Brasília, DF: MTE, 2019.",
    "BRASIL. Ministério do Trabalho e Emprego. **NR-35**: trabalho em altura. "
    "Brasília, DF: MTE, [PREENCHER: ano da redação consultada].",
    "HEINRICH, H. W. **Industrial accident prevention**: a scientific "
    "approach. New York: McGraw-Hill, 1931.",
    "HOLLNAGEL, E. **Safety-I and Safety-II**: the past and future of safety "
    "management. Farnham: Ashgate, 2014.",
    "LEVESON, N. G. **Engineering a safer world**: systems thinking applied to "
    "safety. Cambridge, MA: MIT Press, 2011.",
    "ORGANIZAÇÃO MUNDIAL DA SAÚDE. **CID-11**: Classificação Internacional de "
    "Doenças, 11ª revisão: QD85 burn-out. Genebra: OMS, 2019.",
    "PORTO, M. F. S. Prevention, social emancipation, and paradigmatic "
    "transition: a 40-year interdisciplinary Brazilian trajectory on accidents "
    "and disasters. **Cadernos de Saúde Pública**, Rio de Janeiro, 2024. "
    "[PREENCHER: conferir volume, número e DOI no PubMed, PMID 38775613].",
    "REASON, J. **Human error**. Cambridge: Cambridge University Press, 1990.",
    "REASON, J. **Managing the risks of organizational accidents**. "
    "Aldershot: Ashgate, 1997.",
    "SOUZA, S. F.; CARVALHO, F. M.; ARAÚJO, T. M.; PORTO, L. A. Fatores "
    "psicossociais do trabalho e transtornos mentais comuns em eletricitários. "
    "**Revista de Saúde Pública**, São Paulo, v. 44, n. 4, p. 710-717, 2010.",
    "TREML, M. F. Q. et al. Burnout syndrome in Brazil (2014–2024): regional "
    "variations and temporal trends in an epidemiological study. **Revista "
    "Brasileira de Medicina do Trabalho**, 2025. [PREENCHER: conferir volume, "
    "número, páginas e DOI].",
    "UNIVERSIDADE DE SÃO PAULO. Agência de Bibliotecas e Coleções Digitais. "
    "**Diretrizes para apresentação de dissertações e teses da USP**: parte I "
    "(ABNT). 5. ed. São Paulo: ABCD/USP, 2024. DOI: 10.11606/9786598386221.",
]
for r in REFS:
    ref(doc, r)

# ============================================================ APÊNDICES ======
new_page()
h1_center(doc, "Apêndice A – Roteiro de registro e investigação de ocorrências")
body(doc, "Campos do registro em campo: tipo de ocorrência (acidente com "
     "afastamento, acidente sem afastamento, quase acidente ou condição "
     "insegura); data e hora; função genérica do envolvido; tarefa em "
     "execução; descrição do fato; parte do corpo atingida, se houver lesão; "
     "alerta de emissão de CAT quando aplicável.", indent=False)
body(doc, "Roteiro de investigação, obrigatório para acidentes e recomendado "
     "para quase acidentes:", indent=False)
alineas(doc, [
    "nível 1, causas imediatas: o que, no momento do evento, produziu ou quase "
    "produziu o dano? Atos e condições presentes na cena",
    "nível 2, causas subjacentes: o que, na tarefa e no posto, tornou possíveis "
    "as causas imediatas? Ferramenta, procedimento, improvisação, pressa, "
    "planejamento do serviço",
    "nível 3, causas básicas: o que, na gestão e na organização, originou as "
    "causas subjacentes? Dimensionamento, capacitação, manutenção, fatores "
    "humanos e organizacionais",
    "ações: corretivas e preventivas, com responsável, prazo e status",
])

new_page()
h1_center(doc, "Apêndice B – Telas dos módulos complementares")
body(doc, "[PREENCHER: inserir, com título acima e fonte abaixo, as telas dos "
     "módulos de permissão de trabalho, fatores psicossociais e avaliação "
     "ergonômica preliminar, sem dados pessoais.]", indent=False)

doc.save(OUT)
print(f"OK: {OUT} gerado.")
