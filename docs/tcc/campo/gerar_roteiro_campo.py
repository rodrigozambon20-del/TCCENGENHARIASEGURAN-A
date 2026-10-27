# -*- coding: utf-8 -*-
"""
Gera Roteiro_Campo_Poda.docx: roteiro de campo, protocolo de dosimetria,
fichas de verificação para imprimir e modelos de autorização do estudo de
caso "poda de árvores junto à rede de distribuição".

Uso:  python3 gerar_roteiro_campo.py      Requer: pip install python-docx

Fontes dos parâmetros: NHO-01 (Fundacentro, 2001); NR-15 Anexo 1; NR-10
(tabela de raios de delimitação); NR-12 Anexo V; NR-35 (Portaria MTP nº
4.218/2022 e alterações); NR-6; PGR 2026 da organização (cargo de podador).
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = "Roteiro_Campo_Poda.docx"
NAVY = RGBColor(0x16, 0x23, 0x3F)

doc = Document()
st = doc.styles["Normal"]
st.font.name = "Arial"; st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
st.paragraph_format.space_after = Pt(4)
for name, size in (("Heading 1", 15), ("Heading 2", 12.5), ("Heading 3", 11)):
    h = doc.styles[name]
    h.font.name = "Arial"; h.font.size = Pt(size); h.font.bold = True
    h.font.color.rgb = NAVY
    h.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    h.paragraph_format.space_before = Pt(10); h.paragraph_format.space_after = Pt(4)
for sec in doc.sections:
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(1.6)
    sec.left_margin = sec.right_margin = Cm(1.8)

# rodapé com número de página
fp = doc.sections[0].footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = fp.add_run("Roteiro de campo — poda junto à rede · pág. ")
r.font.size = Pt(8)
run = fp.add_run(); run.font.size = Pt(8)
for el, attrs, text in (("w:fldChar", {"w:fldCharType": "begin"}, None),
                        ("w:instrText", {"xml:space": "preserve"}, "PAGE"),
                        ("w:fldChar", {"w:fldCharType": "end"}, None)):
    n = OxmlElement(el)
    for k, v in attrs.items():
        n.set(qn(k), v)
    if text:
        n.text = text
    run._r.append(n)


def p(text, bold=False, size=None, italic=False, align=None, after=None):
    par = doc.add_paragraph()
    parts = text.split("**")
    for i, chunk in enumerate(parts):
        if not chunk:
            continue
        rr = par.add_run(chunk)
        rr.bold = bold or (i % 2 == 1)
        rr.italic = italic
        if size:
            rr.font.size = Pt(size)
    if align:
        par.alignment = align
    if after is not None:
        par.paragraph_format.space_after = Pt(after)
    return par


def bullets(items, style="List Bullet"):
    for it in items:
        par = doc.add_paragraph(style=style)
        parts = it.split("**")
        for i, chunk in enumerate(parts):
            if chunk:
                par.add_run(chunk).bold = (i % 2 == 1)


def shade(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def table(header, rows, widths, font=9, blank_rows=0, row_height=None):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]
        c.text = ""
        rr = c.paragraphs[0].add_run(h); rr.bold = True; rr.font.size = Pt(font)
        shade(c, "DCE3EE")
    all_rows = list(rows) + [[""] * len(header) for _ in range(blank_rows)]
    for row in all_rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = ""
            parts = str(val).split("**")
            for k, chunk in enumerate(parts):
                if chunk:
                    rr = cells[i].paragraphs[0].add_run(chunk)
                    rr.font.size = Pt(font); rr.bold = (k % 2 == 1)
    fix_widths(t, widths)
    for row in t.rows:
        if row_height:
            row.height = Cm(row_height)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def fix_widths(t, widths):
    """Larguras fixas (grade + células) para Word e LibreOffice respeitarem."""
    t.autofit = False
    tblPr = t._tbl.tblPr
    layout = OxmlElement("w:tblLayout"); layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    grid = t._tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(int(Cm(w).twips)))
    for i, w in enumerate(widths):
        t.columns[i].width = Cm(w)
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Cm(w)


def field_lines(labels, width=17.4):
    """Campos de preenchimento: rótulo + linha."""
    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    for lab in labels:
        cells = t.add_row().cells
        cells[0].text = ""
        rr = cells[0].paragraphs[0].add_run(lab); rr.font.size = Pt(9); rr.bold = True
        shade(cells[0], "F2F4F7")
        t.rows[-1].height = Cm(0.66)
    fix_widths(t, [6.2, width - 6.2])
    sp = doc.add_paragraph(); sp.paragraph_format.space_after = Pt(0)
    sp.add_run().font.size = Pt(4)


def page_break():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


CNC = "C  /  NC  /  NA"

# =========================================================== CAPA / RESUMO ===
p("ROTEIRO DE CAMPO", bold=True, size=20, align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
p("Estudo de caso: poda de árvores junto à rede de distribuição de energia",
  bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
p("Monografia — Especialização em Engenharia de Segurança do Trabalho (PECE/EPUSP) — "
  "Rodrigo Zambon", size=9.5, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=10)

doc.add_heading("Como usar este roteiro", level=2)
bullets([
    "Leia as seções 1 a 4 antes do campo. **Imprima as fichas F1 a F9** "
    "(seção 5): um jogo de F1, F2 e F3 **para cada jornada medida**; F4 a F7 "
    "uma vez por podador (ou por dia, se as condições mudarem); F8 uma vez por "
    "podador; F9 acompanha todo o estudo.",
    "**Nunca escreva nomes nas fichas.** Use códigos: **P1** e **P2** para os "
    "podadores; equipamentos pelo nº de série/patrimônio. Guarde a relação "
    "código × nome só com você.",
    "Lance os dados na planilha **Planilha_Dados_Poda.xlsx** (mesma pasta) ou "
    "fotografe as fichas e me envie.",
    "Antes do primeiro dia, peça as assinaturas dos modelos dos **Anexos A e B**.",
])

doc.add_heading("O que o estudo de caso vai mostrar", level=2)
p("O PGR 2026 da organização (cargo de podador, págs. 90 a 93 do PDF) e os "
  "laudos dizem uma coisa; o campo pode mostrar outra. Cada hipótese abaixo "
  "nasce de um ponto concreto desses documentos:")
table(["Hip.", "O que os documentos dizem", "O que você vai verificar/medir", "Ficha"],
      [["H1", "Ruído avaliado só de forma qualitativa; fonte citada: “trator e "
              "empilhadeira”; o próprio PGR recomenda avaliação quantitativa. "
              "NR-15 e LTCAT: sem insalubridade / sem exposição.",
        "Dosimetria da jornada (NHO-01 e NR-15 Anexo 1) com motosserra e motopoda.",
        "F1, F2, F3"],
       ["H2", "Choque elétrico: probabilidade rara, severidade leve, risco "
              "trivial, “nenhuma ação necessária”. Laudo NR-16: periculosidade.",
        "Condições reais de trabalho próximo à rede (NR-10).", "F4"],
       ["H3", "Queda acima de 2 m: risco trivial.",
        "Execução do trabalho em altura (NR-35).", "F5"],
       ["H4", "Motosserra/motopoda citadas, sem avaliação dos dispositivos.",
        "Dispositivos obrigatórios e treinamento (NR-12, Anexo V).", "F6"],
       ["H5", "Lista de EPI sem itens específicos para motosserra (calça/perneira "
              "anticorte, protetor facial).",
        "EPI entregue × usado × listado (NR-6).", "F7"],
       ["H6", "Calor não consta do inventário (trabalho a céu aberto).",
        "Opcional: IBUTG (NR-15 Anexo 3), se houver medidor.", "F2 (obs.)"]],
      [1.1, 6.6, 7.2, 2.5])

doc.add_heading("Regras de segurança para você no campo", level=2)
bullets([
    "Você vai **observar e medir, não executar**: não suba, não opere "
    "motosserra/motopoda e **nunca se aproxime da rede**. Distâncias galho–rede "
    "são estimadas do chão (ver F4).",
    "Fique **fora da área isolada** abaixo da poda (queda de galhos) e afastado "
    "da projeção da motosserra.",
    "Use seus EPIs: capacete, óculos, protetor auricular, botina e colete "
    "refletivo (trabalho em via pública).",
    "Se presenciar **risco grave e iminente**, comunique na hora ao encarregado "
    "(NR-1). A pesquisa nunca vem antes da segurança.",
])

page_break()

# ========================================================== 1 PREPARAÇÃO =====
doc.add_heading("1  Antes do campo — preparação", level=1)
table(["#", "Tarefa", "Até", "Feito"],
      [["1", "Enviar o tema à Profa. Renata (renata@gmirmusp.com.br) e aguardar retorno", "03/11/2026", "☐"],
       ["2", "Autorização por escrito da organização para usar PGR/laudos e fazer medições (Anexo A)", "antes da 1ª visita", "☐"],
       ["3", "Concordância dos dois podadores (Anexo B), explicando que é voluntário e anônimo", "antes da 1ª visita", "☐"],
       ["4", "Conseguir o dosímetro (ver 1.1) com certificado de calibração válido", "antes da 1ª medição", "☐"],
       ["5", "Confirmar a tensão nominal da rede onde haverá poda (ex.: 13,8 kV; 34,5 kV) e do secundário (127/220 V)", "antes da 1ª visita", "☐"],
       ["6", "Agendar com o encarregado os dias de poda (meta: 2 jornadas por podador)", "1ª quinzena de nov.", "☐"],
       ["7", "Separar cópia das págs. 90–93 do PGR (podador) para conferir no campo", "antes da 1ª visita", "☐"],
       ["8", "Imprimir fichas F1–F9; prancheta, caneta, relógio/cronômetro, celular carregado", "véspera", "☐"],
       ["9", "Dia piloto: acompanhar 1 jornada sem medir, só para testar as fichas", "1ª visita", "☐"]],
      [0.8, 11.6, 3.4, 1.4])

doc.add_heading("1.1  O dosímetro", level=2)
p("**Dosímetro de ruído** é um medidor que o trabalhador carrega a jornada "
  "inteira; ele soma a exposição e informa a **dose** (%). Onde conseguir, em "
  "ordem de preferência:")
bullets([
    "com a **consultoria que elaborou o PGR** — ela fez dosimetria nos cargos "
    "administrativos, logo possui o equipamento. Pode emprestar ou fazer a "
    "medição com você acompanhando (você registra tudo nas fichas);",
    "com o SESMT/técnico de segurança de outra empresa ou da prefeitura;",
    "aluguel em locadora de instrumentos de SST (alguns dias).",
])
p("Exigências mínimas: **dosímetro com certificado de calibração válido** "
  "(anote nº e validade), **calibrador acústico** também com certificado, e "
  "software ou tela para ler/baixar o resultado. Ideal: **dois dosímetros** "
  "(um por podador no mesmo dia).")

doc.add_heading("1.2  Configuração (peça a quem fornecer o equipamento)", level=2)
table(["Parâmetro", "Canal 1 — NHO-01 (prevenção)", "Canal 2 — NR-15 Anexo 1 (legal)"],
      [["Critério de referência (CR)", "85 dB(A) para 8 h", "85 dB(A) para 8 h"],
       ["Incremento de duplicação de dose (q)", "3 dB", "5 dB"],
       ["Nível limiar de integração (NLI)", "80 dB(A)", "80 dB(A)"],
       ["Ponderação / resposta", "A / lenta (slow)", "A / lenta (slow)"],
       ["Faixa de medição", "no mínimo 80 a 115 dB(A)", "no mínimo 80 a 115 dB(A)"]],
      [6.0, 5.7, 5.7])
p("Se o aparelho só tiver um canal, use a **NHO-01** e anote isso na F2.",
  italic=True, size=9)

doc.add_heading("1.3  Agenda sugerida", level=2)
table(["Período", "O quê"],
      [["até 03/11", "Tema enviado; autorizações; dosímetro garantido"],
       ["novembro", "Dia piloto + 1ª rodada de medições (P1 e P2)"],
       ["até meados de dezembro", "2ª rodada de medições; entrevistas (F8); fotos (F9)"],
       ["até 20/12", "Enviar fichas/planilha, relatórios do dosímetro, certificados e fotos"],
       ["dez. a jan.", "Redação dos capítulos de métodos, resultados e discussão"],
       ["até 22/02/2027", "Entrega da monografia em PDF no Moodle"]],
      [4.5, 12.9])

page_break()

# ===================================================== 2 PROTOCOLO RUÍDO =====
doc.add_heading("2  Protocolo da dosimetria (passo a passo)", level=1)
steps = [
    "**Verificar a calibração antes:** acople o calibrador ao microfone e anote a "
    "leitura na F2 (campo “calibração antes”).",
    "**Explicar ao podador** (P1/P2): o aparelho só mede som, não grava conversa; "
    "não deve tirá-lo, cobri-lo, nem falar/gritar junto ao microfone; manter a "
    "rotina normal.",
    "**Instalar:** dosímetro preso no cinto; **microfone sobre o ombro**, na zona "
    "auditiva (próximo à orelha), do lado mais exposto à motosserra; cabo preso "
    "para não enroscar em galhos; microfone fora do abafador e sem roçar na roupa. "
    "(Conferir a distância exata da zona auditiva no texto da NHO-01.)",
    "**Ligar no início da jornada** e anotar a hora. Medir **a jornada inteira**, "
    "incluindo deslocamentos, sinalização, limpeza e pausas.",
    "**Durante a jornada**, preencher a **F3 (diário de atividades)**: hora de "
    "início/fim de cada atividade e qual equipamento estava ligado. É o diário "
    "que explica a dose no capítulo de resultados.",
    "Anotar na F2 se o **protetor auditivo** estava em uso durante o corte "
    "(sempre / às vezes / não) — o dosímetro mede o ruído **fora** do protetor, "
    "como a norma pede; a eficácia do protetor é discutida à parte.",
    "**Ao final da jornada:** parar a medição, anotar hora de fim, **tempo de "
    "medição (TE)**, **dose do canal 1 e do canal 2 (%)** e o **Leq/Lavg**. "
    "Fotografar a tela de resultado.",
    "**Verificar a calibração depois** e anotar. Se a diferença entre antes e "
    "depois passar de **±1 dB**, a amostra é **inválida**: anote e repita em outro dia.",
    "**Baixar o relatório** do software do dosímetro (PDF/CSV) e guardar com o "
    "código da amostra (ex.: A01_P1_2026-11-10).",
]
for s in steps:
    par = doc.add_paragraph(style="List Number")
    parts = s.split("**")
    for i, chunk in enumerate(parts):
        if chunk:
            par.add_run(chunk).bold = (i % 2 == 1)

doc.add_heading("2.1  Como o resultado será interpretado (a planilha calcula)", level=2)
bullets([
    "**NEN (nível de exposição normalizado, NHO-01):** NEN = 85 + 10·log₁₀(D/100), "
    "com D = dose do canal 1 (%) referida à jornada. Se a medição cobriu só parte "
    "da jornada, a dose é projetada para a jornada (D × jornada/TE), desde que o "
    "período medido seja representativo.",
    "**Critério NHO-01:** NEN < 82 dB(A) = aceitável; 82 a 85 dB(A) = **nível de "
    "ação** (dose ≥ 50%): exige medidas preventivas; ≥ 85 dB(A) = **acima do "
    "limite de exposição**: exige medidas corretivas.",
    "**Critério NR-15 (Anexo 1):** dose do canal 2 > 100% = acima do limite de "
    "tolerância (base legal da insalubridade).",
    "**Quantas amostras:** meta de 2 jornadas por podador (4 amostras). Com menos "
    "de 4, os resultados serão apresentados como indicativos, sem estatística.",
])

page_break()

# ============================================================ 5 FICHAS =======
doc.add_heading("Fichas de campo", level=1)

doc.add_heading("F1 — Identificação da jornada", level=2)
field_lines(["Amostra nº (ex.: A01)", "Data", "Podador (código P1/P2)",
             "Local (genérico: urbano/rural; via pública/propriedade)",
             "Condições do tempo (sol, nublado, chuva, vento) e temperatura aprox.",
             "Tensão da rede próxima (primário / secundário)",
             "Rede: energizada ou desenergizada durante a poda?",
             "Tipo de serviço (poda de galhos, retirada de árvore, limpeza de faixa)",
             "Meio de acesso (chão, escada, subida na árvore, cesto aéreo)",
             "Equipamentos usados (motosserra / motopoda / podão / roçadeira)",
             "Hora de início / fim da jornada",
             "Quantas pessoas na equipe (códigos)"])

doc.add_heading("F2 — Dosimetria", level=2)
field_lines(["Amostra nº / podador", "Dosímetro: marca, modelo, nº de série",
             "Certificado de calibração do dosímetro: nº e validade",
             "Calibrador: nº de série e validade do certificado",
             "Configuração: canal 1 (NHO-01) / canal 2 (NR-15) — conferido?",
             "Calibração ANTES (leitura em dB)", "Calibração DEPOIS (leitura em dB)",
             "Hora de início / fim da medição", "Tempo de medição TE (min)",
             "Jornada do dia (min)", "Dose canal 1 — NHO-01 (%)",
             "Dose canal 2 — NR-15 (%)", "Leq / Lavg (dB(A))",
             "Protetor auditivo usado no corte? (sempre / às vezes / não) e CA",
             "Ocorrências (aparelho tirado, chuva, pausa longa etc.)",
             "Opcional: IBUTG (°C), horário e local da medição"])

page_break()
doc.add_heading("F3 — Diário de atividades da jornada", level=2)
p("Amostra nº: ______   Podador: ______   Data: ____/____/______", size=9.5)
table(["Início", "Fim", "Atividade (corte, subida, sinalização, deslocamento, limpeza, pausa, almoço…)",
       "Equipamento ligado", "Obs."], [], [1.6, 1.6, 7.6, 3.2, 3.4],
      blank_rows=20, row_height=0.72)

page_break()
doc.add_heading("F4 — Risco elétrico: trabalho próximo à rede (NR-10)", level=2)
p("Data: ____/____/______   Amostra/podadores: ______________   Tensão da rede: ________",
  size=9.5)
table(["#", "Item verificado", "Resultado", "Evidência / obs."],
      [["1", "Podadores com treinamento NR-10 básico (40 h) — certificado conferido", CNC, ""],
       ["2", "Podadores com curso complementar SEP (40 h) — o PGR recomenda", CNC, ""],
       ["3", "Podadores formalmente autorizados para serviços em/próximo de instalações elétricas", CNC, ""],
       ["4", "Ordem de serviço / programação do serviço existente", CNC, ""],
       ["5", "Análise de risco (APR) feita ANTES de iniciar, específica do local", CNC, ""],
       ["6", "Se a rede foi desligada: seccionamento, impedimento de reenergização, "
             "constatação de ausência de tensão, aterramento temporário, proteção dos "
             "elementos energizados e sinalização — seguidos nessa ordem", CNC, ""],
       ["7", "Se a rede estava energizada: trabalhadores e ferramentas mantidos fora da "
             "zona de risco (tabela abaixo); zona controlada só com autorizados", CNC, ""],
       ["8", "Ferramentas usadas perto da rede são isolantes (podão/motopoda com cabo "
             "isolante) e com ensaio dielétrico em dia", CNC, ""],
       ["9", "Galhos cortados não caíram/encostaram na rede (houve controle da queda?)", CNC, ""],
       ["10", "Comunicação com o centro de operação/despacho durante o serviço", CNC, ""],
       ["11", "Serviço não realizado por trabalhador sozinho", CNC, ""],
       ["12", "Equipe com pessoa capacitada em primeiros socorros/RCP e plano de emergência", CNC, ""],
       ["13", "Área sinalizada e isolada para pedestres e veículos", CNC, ""]],
      [0.8, 10.4, 2.6, 3.6], font=8.5)
p("**Distâncias da NR-10 (raios de delimitação, em metros)** — conferir a "
  "tabela no texto oficial da NR-10 antes da redação:", size=9)
table(["Tensão nominal", "Rr — limite da zona de risco", "Rc — limite da zona controlada"],
      [["< 1 kV (secundária)", "0,20 m", "0,70 m"],
       ["≥ 10 kV e < 15 kV (ex.: 13,8 kV)", "0,38 m", "1,38 m"],
       ["≥ 15 kV e < 20 kV", "0,40 m", "1,40 m"],
       ["outras tensões", "consultar a tabela da NR-10", "consultar a tabela da NR-10"]],
      [6.0, 5.7, 5.7], font=9)
p("Como estimar a distância galho/trabalhador–condutor **sem se aproximar**: do "
  "chão, compare com uma medida conhecida da estrutura (comprimento da cruzeta, "
  "distância entre condutores) e fotografe. Anote sempre “**estimado**”.",
  italic=True, size=9)

page_break()
doc.add_heading("F5 — Trabalho em altura (NR-35)", level=2)
p("Data: ____/____/______   Amostra/podadores: ______________", size=9.5)
table(["#", "Item verificado", "Resultado", "Evidência / obs."],
      [["1", "Treinamento NR-35 (mín. 8 h) — certificado conferido (verificar se presencial)", CNC, ""],
       ["2", "ASO indica aptidão para trabalho em altura", CNC, ""],
       ["3", "Análise de risco do trabalho em altura feita antes de subir", CNC, ""],
       ["4", "Permissão de trabalho emitida (se atividade não rotineira)", CNC, ""],
       ["5", "Cinturão paraquedista e talabarte inspecionados antes do uso", CNC, ""],
       ["6", "Ancoragem: ponto escolhido e verificado (tronco/galho resistente, escada, cesto)", CNC, ""],
       ["7", "Trabalhador permanece conectado o tempo todo em altura", CNC, ""],
       ["8", "Meio de acesso adequado (escada amarrada/estabilizada, cesto aéreo inspecionado)", CNC, ""],
       ["9", "Condições meteorológicas avaliadas (vento, chuva) — trabalho suspenso se adversas", CNC, ""],
       ["10", "Plano de resgate definido e pessoa apta ao resgate presente", CNC, ""],
       ["11", "Área abaixo isolada contra queda de galhos e ferramentas", CNC, ""],
       ["12", "Motosserra NÃO usada acima da altura dos ombros / com uma só mão", CNC, ""]],
      [0.8, 10.4, 2.6, 3.6], font=8.5)

page_break()
doc.add_heading("F6 — Motosserra e motopoda (NR-12, Anexo V)", level=2)
p("Preencha uma coluna por equipamento.", size=9)
table(["Item", "Equip. 1", "Equip. 2", "Equip. 3"],
      [["Tipo (motosserra/motopoda), marca, modelo", "", "", ""],
       ["Nº de série ou patrimônio", "", "", ""],
       ["Freio de corrente (manual ou automático) presente e funcionando", CNC, CNC, CNC],
       ["Pino pega-corrente", CNC, CNC, CNC],
       ["Protetor da mão direita", CNC, CNC, CNC],
       ["Protetor da mão esquerda", CNC, CNC, CNC],
       ["Trava de segurança do acelerador", CNC, CNC, CNC],
       ["Manual de instruções em português disponível", CNC, CNC, CNC],
       ["Operador com treinamento de motosserra (mín. 8 h) — certificado", CNC, CNC, CNC],
       ["Corrente afiada e tensionada; sabre sem danos", CNC, CNC, CNC],
       ["Abastecimento com motor desligado e longe de fonte de ignição", CNC, CNC, CNC],
       ["Protetor do sabre usado no transporte", CNC, CNC, CNC]],
      [7.4, 3.3, 3.3, 3.4], font=8.5)

page_break()
doc.add_heading("F7 — EPI: listado no PGR × entregue × em uso (NR-6)", level=2)
p("Podador: ______   Data: ____/____/______   Consultar validade do CA em "
  "caepi.mte.gov.br", size=9.5)
S_N = "S / N"
table(["EPI", "No PGR?", "Ficha de entrega?", "Em uso no campo?", "CA válido?", "Estado / obs."],
      [["Capacete classe B (CA 29792)", "Sim", S_N, "S / N / parcial", S_N, ""],
       ["Óculos (CA 10346)", "Sim", S_N, "S / N / parcial", S_N, ""],
       ["Protetor auditivo (CA 35981)", "Sim", S_N, "S / N / parcial", S_N, ""],
       ["Luva contra agentes mecânicos (CA 40029)", "Sim", S_N, "S / N / parcial", S_N, ""],
       ["Perneira (CA 25388) — no PGR ligada a animais peçonhentos", "Sim", S_N, "S / N / parcial", S_N, ""],
       ["Botina tipo B (CA 12160) / calçado com bico", "Sim", S_N, "S / N / parcial", S_N, ""],
       ["Cinturão paraquedista e talabarte (CA 35509)", "Sim", S_N, "S / N / parcial", S_N, ""],
       ["Calça ou perneira de proteção contra corte por motosserra", "Não", S_N, "S / N / parcial", S_N, ""],
       ["Protetor facial (viseira de tela)", "Não", S_N, "S / N / parcial", S_N, ""],
       ["Luva de proteção para motosserra", "Não", S_N, "S / N / parcial", S_N, ""],
       ["Vestimenta para risco elétrico/arco (se houver exposição à rede)", "Não", S_N, "S / N / parcial", S_N, ""],
       ["Colete refletivo (via pública)", "Não", S_N, "S / N / parcial", S_N, ""],
       ["Protetor solar fornecido", "Medida adm.", S_N, "S / N / parcial", "—", ""]],
      [5.6, 1.6, 2.0, 2.4, 1.6, 4.2], font=8.5)
p("Os CAs são os do PGR 2026 da organização; confira se o EPI em uso é o "
  "mesmo. Não fotografe etiquetas com nome do trabalhador.", italic=True, size=8.5)

page_break()
doc.add_heading("F8 — Conversa com o podador (5 a 10 min, anônima e voluntária)", level=2)
p("Código: ______   Data: ____/____/______   Explique antes: é voluntário, "
  "anônimo, pode parar quando quiser e não será usado para avaliação dele.",
  size=9.5)
qs = ["Há quanto tempo trabalha com poda? E nesta função na empresa?",
      "Quais treinamentos recebeu (NR-10, SEP, NR-35, motosserra, primeiros "
      "socorros)? Mais ou menos quando?",
      "Já passou por choque, “quase choque”, galho encostando na rede, queda, "
      "corte ou galho caindo perto? O que aconteceu (sem nomes)?",
      "Usa protetor auricular o tempo todo com a motosserra? Se não, por quê "
      "(incomoda, atrapalha a comunicação, esquece, calor)?",
      "Ao fim do dia sente zumbido ou ouvido “abafado”? (só sim/não/às vezes — "
      "não é avaliação médica)",
      "Os EPIs que recebe são adequados e confortáveis para poda? Falta algum?",
      "Quem decide se a rede será desligada para a poda? Isso costuma acontecer?",
      "O que mais te preocupa no serviço? O que melhoraria a segurança?"]
for i, q in enumerate(qs, 1):
    p(f"{i}. {q}", size=9.5, after=1)
    for _ in range(3):
        p("_" * 98, size=8, after=0)

page_break()
doc.add_heading("F9 — Registro fotográfico", level=2)
p("As fotos do estudo de caso são **obrigatórias no pôster**. Regras: **sem "
  "rostos, sem logotipo da empresa** (uniforme, veículo, equipamento), **sem "
  "placas de veículos** e sem fachadas/nomes de rua identificáveis. Se não "
  "der para evitar, fotografe assim mesmo e me avise: o desfoque é feito antes "
  "de entrar no TCC. Guarde os originais.", size=9.5)
p("Fotos que precisamos ter:", bold=True, size=9.5)
bullets(["visão geral da atividade de poda (sem rosto)",
         "dosímetro instalado (microfone no ombro)",
         "motosserra/motopoda mostrando os dispositivos de segurança",
         "EPIs em uso (incluindo o que falta)",
         "sinalização e isolamento da área",
         "galho próximo à rede, fotografado do chão",
         "meio de acesso (escada, ancoragem, cesto)",
         "tela de resultado do dosímetro"])
table(["Nº foto", "Data", "O que mostra", "Ficha/item", "Precisa desfocar?"],
      [], [1.6, 2.0, 8.0, 2.6, 3.2], blank_rows=14, row_height=0.65)

# =================================================== DEPOIS DO CAMPO ========
page_break()
doc.add_heading("3  Depois do campo — o que me enviar", level=1)
bullets([
    "Planilha preenchida **ou** fotos legíveis de todas as fichas;",
    "relatório do dosímetro de cada amostra (PDF/CSV) e foto da tela de resultado;",
    "certificados de calibração do dosímetro e do calibrador (para a metodologia);",
    "fotos (F9) com a lista de numeração;",
    "respostas da F8;",
    "qualquer documento da organização sobre poda: procedimento, APR padrão, "
    "ordem de serviço, registros de treinamento (anonimizados).",
])
p("Com isso eu escrevo os capítulos 3 (Materiais e métodos) e 4 (Resultados e "
  "discussão), monto as tabelas e gráficos, e atualizo pôster e apresentação.")

# ============================================================ ANEXOS ========
page_break()
doc.add_heading("Anexo A — Modelo de autorização da organização", level=1)
p("AUTORIZAÇÃO PARA REALIZAÇÃO DE ESTUDO DE CASO", bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER)
p("A [nome da organização], inscrita no CNPJ [nº], autoriza [nome completo do "
  "aluno], aluno do curso de Especialização em Engenharia de Segurança do "
  "Trabalho do Programa de Educação Continuada em Engenharia da Escola "
  "Politécnica da Universidade de São Paulo, a realizar estudo de caso para "
  "sua monografia, compreendendo: (a) acompanhamento das atividades de poda de "
  "árvores junto à rede de distribuição; (b) medição de ruído ocupacional por "
  "dosimetria; (c) verificação de condições de segurança, equipamentos e EPIs; "
  "(d) registro fotográfico; e (e) consulta ao Programa de Gerenciamento de "
  "Riscos, ao LTCAT, aos laudos de insalubridade e periculosidade e ao PCMSO "
  "vigentes.")
p("A autorização fica condicionada a que o trabalho: não identifique a "
  "organização, seus trabalhadores ou terceiros, inclusive em fotografias "
  "(rostos, logotipos e placas ocultados); utilize as informações "
  "exclusivamente para fins acadêmicos; e não interfira na execução segura dos "
  "serviços.")
p("[Cidade], ____ de ______________ de 2026.")
p("\n\n______________________________________\n[Nome e cargo do responsável]")

page_break()
doc.add_heading("Anexo B — Modelo de concordância do trabalhador", level=1)
p("CONCORDÂNCIA EM PARTICIPAR DE ESTUDO DE CASO", bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER)
p("Fui informado de que [nome do aluno] realiza estudo de caso para monografia "
  "de especialização da Universidade de São Paulo sobre a segurança na poda de "
  "árvores junto à rede elétrica, e concordo em:")
bullets(["usar um dosímetro de ruído durante a jornada (o aparelho mede apenas "
         "o nível de som e não grava conversas);",
         "ter minhas atividades observadas e responder a uma conversa curta;",
         "ser fotografado durante o trabalho, desde que meu rosto não apareça "
         "ou seja ocultado."])
p("Sei que minha participação é voluntária, que posso desistir a qualquer "
  "momento sem prejuízo, que não serei identificado no trabalho e que as "
  "informações não serão usadas para me avaliar.")
p("Data: ____/____/______")
p("\nCódigo do participante: ______\n\nAssinatura: ______________________________________")

doc.save(OUT)
print(f"OK: {OUT}")
