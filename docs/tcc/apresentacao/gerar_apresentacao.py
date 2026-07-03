# -*- coding: utf-8 -*-
"""
Gerador da apresentação da banca (defesa) do TCC — Psike.
Produz Apresentacao_Banca_Psike.pptx (16:9), com a identidade visual do
projeto (navy #16233F + dourado #B4842A) e telas reais da plataforma.

Uso:  python3 gerar_apresentacao.py
Requer: pip install python-pptx  (e as imagens em ./img)
Notas para a defesa (roteiro de fala) ficam nas notas de cada slide.
Trechos [PREENCHER: ...] dependem de dados reais e devem ser completados.
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

NAVY   = RGBColor(0x16, 0x23, 0x3F)
NAVY2  = RGBColor(0x1F, 0x31, 0x57)
GOLD   = RGBColor(0xB4, 0x84, 0x2A)
PAPER  = RGBColor(0xF1, 0xF3, 0xF6)
SLATE  = RGBColor(0x5B, 0x65, 0x76)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
LINE   = RGBColor(0xDD, 0xE2, 0xE9)

IMG = os.path.join(os.path.dirname(__file__), "img")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

def slide():
    return prs.slides.add_slide(BLANK)

def bg(s, color):
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = color

def box(s, l, t, w, h):
    tb = s.shapes.add_textbox(l, t, w, h)
    tb.text_frame.word_wrap = True
    return tb

def rect(s, l, t, w, h, color):
    from pptx.enum.shapes import MSO_SHAPE
    sp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    sp.fill.solid(); sp.fill.fore_color.rgb = color
    sp.line.fill.background()
    sp.shadow.inherit = False
    return sp

def set_para(p, text, size, color, bold=False, align=PP_ALIGN.LEFT,
             font="Calibri", space_after=8):
    p.text = text
    p.alignment = align
    p.space_after = Pt(space_after)
    for r in p.runs:
        r.font.size = Pt(size); r.font.color.rgb = color
        r.font.bold = bold; r.font.name = font
    return p

def title_bar(s, eyebrow, title):
    """Barra de título padrão dos slides de conteúdo."""
    rect(s, 0, 0, SW, Inches(0.14), GOLD)
    tb = box(s, Inches(0.6), Inches(0.35), Inches(12.1), Inches(1.15))
    tf = tb.text_frame
    set_para(tf.paragraphs[0], eyebrow.upper(), 12, GOLD, bold=True,
             font="Consolas", space_after=2)
    set_para(tf.add_paragraph(), title, 30, NAVY, bold=True, font="Georgia")

def bullets(s, items, top=Inches(1.7), left=Inches(0.7),
            width=Inches(12.0), height=Inches(5.2), size=18, gap=10):
    tb = box(s, left, top, width, height)
    tf = tb.text_frame
    first = True
    for it in items:
        lvl = 0
        txt = it
        if isinstance(it, tuple):
            txt, lvl = it
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        bullet = "•  " if lvl == 0 else "–  "
        set_para(p, bullet + txt, size - (2 if lvl else 0),
                 NAVY if lvl == 0 else SLATE, space_after=gap)
        p.level = lvl
    return tb

def notes(s, text):
    s.notes_slide.notes_text_frame.text = text

def pic_card(s, path, l, t, w):
    """Insere imagem com borda, ajustando altura pela proporção."""
    if not os.path.exists(path):
        ph = rect(s, l, t, w, Inches(3.5), PAPER)
        set_para(ph.text_frame.paragraphs[0], "[imagem]", 14, SLATE,
                 align=PP_ALIGN.CENTER)
        return ph
    from PIL import Image
    iw, ih = Image.open(path).size
    h = int(w * ih / iw)
    pic = s.shapes.add_picture(path, l, t, width=w, height=h)
    pic.line.color.rgb = LINE; pic.line.width = Pt(1)
    return pic

# ============================================================ 1. CAPA =========
s = slide(); bg(s, NAVY)
rect(s, 0, Inches(6.9), SW, Inches(0.14), GOLD)
# marca
mk = rect(s, Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), GOLD)
set_para(mk.text_frame.paragraphs[0], "P", 26, NAVY, bold=True,
         align=PP_ALIGN.CENTER, font="Georgia")
mk.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
tb = box(s, Inches(1.6), Inches(0.72), Inches(6), Inches(0.7))
set_para(tb.text_frame.paragraphs[0], "Psike", 26, WHITE, bold=True, font="Georgia")

tb = box(s, Inches(0.9), Inches(2.4), Inches(11.5), Inches(2.6))
tf = tb.text_frame
set_para(tf.paragraphs[0], "GESTÃO DIGITAL DE RISCOS PSICOSSOCIAIS CONFORME A NR-1",
         34, WHITE, bold=True, font="Georgia", space_after=10)
set_para(tf.add_paragraph(),
         "Desenvolvimento e aplicação de uma plataforma de SST em serviços "
         "de distribuição de energia elétrica", 20, RGBColor(0xC7,0xCE,0xDB),
         font="Calibri")

tb = box(s, Inches(0.9), Inches(5.2), Inches(11.5), Inches(1.6))
tf = tb.text_frame
set_para(tf.paragraphs[0], "Rodrigo Zambon  [PREENCHER: nome completo]",
         18, GOLD, bold=True, space_after=4)
set_para(tf.add_paragraph(),
         "Orientador(a): [PREENCHER]   ·   Especialização em Engenharia de "
         "Segurança do Trabalho — PECE/EPUSP", 14, RGBColor(0xAE,0xB8,0xC9))
set_para(tf.add_paragraph(), "São Paulo · [PREENCHER: data da defesa]",
         14, RGBColor(0xAE,0xB8,0xC9))
notes(s, "Cumprimentar a banca. Apresentar-se, o tema e o orientador. "
         "Frase de abertura: o TCC nasce de uma exigência nova da NR-1 (2024) "
         "e entrega uma solução prática, aplicada no setor de distribuição de "
         "energia. Duração-alvo da defesa: ~15 min + arguição.")

# ============================================ 2. CONTEXTO / PROBLEMA ===========
s = slide(); bg(s, WHITE)
title_bar(s, "Contexto e problema", "Uma exigência nova, um processo ainda manual")
bullets(s, [
    "A Portaria MTE nº 1.419/2024 atualizou a NR-1: fatores de risco "
    "psicossocial passaram a ser obrigatórios no GRO/PGR.",
    "Exigibilidade a partir de 26/05/2026 — alcança todo empregador CLT.",
    "Não basta aplicar um questionário: é preciso identificar, avaliar, "
    "classificar, registrar no Inventário de Riscos (IRO) e agir — com "
    "rastreabilidade de longo prazo.",
    "Na prática, a maioria das empresas faz isso em papel/planilha: baixa "
    "rastreabilidade e alto custo operacional.",
    ("No setor de distribuição de energia, os fatores psicossociais convivem "
     "com riscos graves (NR-10, NR-35) e demandas típicas: pressão por "
     "restabelecimento, turnos, sobreaviso, risco de acidente fatal.", 1),
])
notes(s, "Ancorar o problema: a norma mudou, o prazo é real e próximo, e a "
         "forma como as empresas cumprem hoje é frágil. Situar o setor de "
         "energia como campo pertinente.")

# ============================================ 3. JUSTIFICATIVA =================
s = slide(); bg(s, WHITE)
title_bar(s, "Justificativa", "A lacuna entre a norma e as ferramentas")
bullets(s, [
    "Soluções comerciais de SST: custo alto para empresas menores.",
    "Instrumentos validados extensos (ex.: COPSOQ, 80+ itens): exigem "
    "competência estatística pouco disponível em equipes enxutas.",
    "Formulário impresso tabulado à mão: fragiliza justamente o que a NR-1 "
    "passou a exigir — anonimato, consistência e rastreabilidade.",
    "Oportunidade: demonstrar uma solução digital de custo praticamente nulo, "
    "replicável por profissionais de SST sem apoio de TI.",
])
notes(s, "Deixar claro o 'gap' que justifica o trabalho. Enfatizar custo "
         "quase nulo e replicabilidade como diferencial.")

# ============================================ 4. OBJETIVOS =====================
s = slide(); bg(s, WHITE)
title_bar(s, "Objetivos", "Objetivo geral e específicos")
bullets(s, [
    "GERAL: desenvolver e aplicar uma plataforma digital para identificação, "
    "avaliação e documentação de fatores de risco psicossocial conforme a "
    "NR-1, integrada à gestão de riscos de uma distribuidora de energia.",
    ("Sistematizar os requisitos da NR-1 e referências (Guia MTE, ISO 45003);", 1),
    ("Construir um instrumento enxuto (10 itens, 6 dimensões) com coleta "
     "anônima e classificação objetiva de risco;", 1),
    ("Implementar a plataforma (web + backend gratuito) com painel e geração "
     "automática de texto para o IRO;", 1),
    ("Digitalizar processos complementares: APR (NR-10/35), AEP (NR-17) e "
     "acidentes/near-miss;", 1),
    ("Aplicar em piloto e discutir resultados, limitações e generalização.", 1),
])
notes(s, "Ler o objetivo geral com calma; passar pelos específicos "
         "rapidamente — eles espelham a estrutura da metodologia e dos módulos.")

# ============================================ 5. FUNDAMENTAÇÃO =================
s = slide(); bg(s, WHITE)
title_bar(s, "Fundamentação", "Do conceito às seis dimensões avaliadas")
bullets(s, [
    "Fatores psicossociais: aspectos da organização e gestão do trabalho com "
    "potencial de dano — abordagem organizacional, não diagnóstico individual.",
    "Modelos de base: demanda-controle (Karasek), esforço-recompensa "
    "(Siegrist), multidimensional (COPSOQ).",
    "ISO 45003:2021: primeira norma internacional de riscos psicossociais; "
    "recomenda integrar ao sistema de gestão de SST existente.",
    "Seis dimensões adotadas: carga e ritmo · autonomia e controle · clareza "
    "de papel · apoio social e de liderança · reconhecimento · assédio e "
    "violência.",
])
notes(s, "Mostrar que a seleção das dimensões tem lastro teórico e alinhamento "
         "com a ISO 45003 e o Guia do MTE. Reforçar: avalia a organização.")

# ============================================ 6. METODOLOGIA ===================
s = slide(); bg(s, WHITE)
title_bar(s, "Metodologia", "Arquitetura da plataforma e instrumento")
bullets(s, [
    "Pesquisa aplicada/tecnológica: desenvolvimento de artefato + estudo de "
    "caso com aplicação piloto.",
    "Três decisões de arquitetura: custo zero · sem instalação · anonimato "
    "por construção (LGPD).",
    "Frontend: página web autocontida (abre no navegador do celular). "
    "Backend: Google Apps Script + Sheets como repositório único.",
    "Instrumento: 10 afirmativas positivas, escala Likert 1–5; média por "
    "dimensão classificada em três faixas (baixo ≥ 3,8 · médio 2,8–3,8 · "
    "alto < 2,8).",
], height=Inches(3.4))
notes(s, "Explicar por que Sheets/Apps Script: gratuito e replicável. Explicar "
         "os pontos de corte como julgamento técnico conservador, a calibrar "
         "com histórico.")

# ============================================ 7. A PLATAFORMA (4 MÓDULOS) ======
s = slide(); bg(s, PAPER)
title_bar(s, "A plataforma", "Uma base de dados, quatro módulos de SST")
mods = [
    ("MÓDULO 1", "Riscos psicossociais", "NR-1 · ISO 45003"),
    ("MÓDULO 2", "APR / PT digital", "NR-10 · NR-35 (distribuição)"),
    ("MÓDULO 3", "Ergonomia (AEP)", "NR-17"),
    ("MÓDULO 4", "Acidentes e near-miss", "NR-1 · árvore de causas"),
]
cw, gap = Inches(2.95), Inches(0.25)
x0 = Inches(0.7); y0 = Inches(2.1)
for i, (num, tit, sub) in enumerate(mods):
    x = x0 + i * (cw + gap)
    card = rect(s, x, y0, cw, Inches(3.4), WHITE)
    tf = box(s, x + Inches(0.2), y0 + Inches(0.25), cw - Inches(0.4), Inches(3.0)).text_frame
    set_para(tf.paragraphs[0], "● em uso", 11, RGBColor(0x2E,0x7D,0x63),
             bold=True, font="Consolas", space_after=8)
    set_para(tf.add_paragraph(), num, 11, GOLD, bold=True, font="Consolas",
             space_after=2)
    set_para(tf.add_paragraph(), tit, 18, NAVY, bold=True, font="Georgia",
             space_after=6)
    set_para(tf.add_paragraph(), sub, 13, SLATE)
notes(s, "Este é o coração da contribuição: a mesma base de dados vira um "
         "'inventário de riscos vivo', alimentado por 4 processos operacionais. "
         "Os quatro módulos estão implementados e funcionais.")

# ============================================ 8. MÓDULO 1 — FLUXO + TELAS ======
s = slide(); bg(s, WHITE)
title_bar(s, "Módulo 1 em detalhe", "Da resposta anônima ao texto do IRO — sem tabulação manual")
pic_card(s, os.path.join(IMG, "shot_form.png"), Inches(0.7), Inches(1.9), Inches(6.0))
pic_card(s, os.path.join(IMG, "shot_sobre.png"), Inches(6.9), Inches(1.9), Inches(5.8))
notes(s, "Demonstrar o fluxo: trabalhador responde no celular (anônimo) → "
         "painel classifica por dimensão → app gera o parágrafo do IRO pronto "
         "para colar no PGR. Se possível, fazer uma demo ao vivo aqui.")

# ============================================ 9. MÓDULOS 2–4 — TELAS ===========
s = slide(); bg(s, WHITE)
title_bar(s, "Módulos 2 a 4", "Permissão de trabalho, ergonomia e acidentes")
pic_card(s, os.path.join(IMG, "shot_apr.png"), Inches(0.6), Inches(1.9), Inches(4.0))
pic_card(s, os.path.join(IMG, "shot_aep.png"), Inches(4.75), Inches(1.9), Inches(4.0))
pic_card(s, os.path.join(IMG, "shot_acidentes.png"), Inches(8.9), Inches(1.9), Inches(3.7))
tb = box(s, Inches(0.6), Inches(6.5), Inches(12), Inches(0.7))
set_para(tb.text_frame.paragraphs[0],
         "M2: checklist NR-10/35, geolocalização e assinatura digital  ·  "
         "M3: AEP em 5 blocos com parecer  ·  M4: árvore de causas e CAT",
         13, SLATE, align=PP_ALIGN.CENTER)
notes(s, "Passar rápido: mostrar que a plataforma cobre o campo (APR com "
         "assinatura e GPS), a ergonomia (AEP com parecer automático) e a "
         "investigação de acidentes com árvore de causas em 3 níveis.")

# ============================================ 10. RESULTADOS ===================
s = slide(); bg(s, WHITE)
title_bar(s, "Resultados", "Plataforma operacional + aplicação piloto")
bullets(s, [
    "Plataforma implementada integralmente: 4 módulos operacionais, custo de "
    "operação nulo.",
    "Fluxo do Módulo 1 sem qualquer etapa manual de tabulação.",
    "[PREENCHER: nº de respondentes, taxa de adesão e nota média / faixa de "
    "risco por dimensão — inserir tabela e gráfico do painel real].",
    "[PREENCHER: transcrever o texto do IRO gerado pela plataforma no piloto, "
    "como evidência do produto final].",
], height=Inches(3.6))
notes(s, "Aqui entram os dados reais do piloto. NÃO inventar. Se o piloto "
         "ainda não ocorreu na data da defesa, apresentar como resultado "
         "esperado e mostrar o texto do IRO gerado com dados de demonstração.")

# ============================================ 11. DISCUSSÃO / LIMITAÇÕES =======
s = slide(); bg(s, WHITE)
title_bar(s, "Discussão e limitações", "O que os resultados indicam — e o que não")
bullets(s, [
    "A barreira à conformidade com a NR-1 é operacional, não conceitual: a "
    "digitalização reduz o custo do ciclo a quase zero.",
    "Ganhos alinhados à norma: anonimato, consistência da classificação e "
    "rastreabilidade dos registros.",
    ("Limitações: instrumento de triagem (10 itens), sem validação "
     "psicométrica formal;", 1),
    ("pontos de corte por julgamento técnico, a calibrar com histórico;", 1),
    ("sem autenticação e sem segregação por empresa (uso comercial exige);", 1),
    ("limites da camada gratuita do Sheets; amostra restrita a uma "
     "organização.", 1),
])
notes(s, "Mostrar maturidade reconhecendo limites — a banca valoriza. "
         "Enquadrar o instrumento como triagem (screening), não diagnóstico.")

# ============================================ 12. CONCLUSÃO ====================
s = slide(); bg(s, WHITE)
title_bar(s, "Conclusão", "Contribuição e trabalhos futuros")
bullets(s, [
    "A digitalização viabiliza o cumprimento da NR-1 mesmo com equipes de SST "
    "enxutas, a custo praticamente nulo.",
    "Conceito central: inventário de riscos 'vivo', alimentado pelos próprios "
    "processos operacionais de SST.",
    "Trabalhos futuros: validação psicométrica; calibração dos cortes; "
    "autenticação e segregação por empresa; migração para banco de dados; "
    "acompanhamento longitudinal da eficácia das medidas.",
])
notes(s, "Fechar reafirmando a contribuição prática e a visão de plataforma. "
         "Deixar gancho para os trabalhos futuros — mostra que o projeto tem "
         "continuidade.")

# ============================================ 13. ENCERRAMENTO =================
s = slide(); bg(s, NAVY)
rect(s, 0, Inches(6.9), SW, Inches(0.14), GOLD)
tb = box(s, Inches(0.9), Inches(2.7), Inches(11.5), Inches(2.2))
tf = tb.text_frame
set_para(tf.paragraphs[0], "Obrigado.", 40, WHITE, bold=True, font="Georgia",
         space_after=12)
set_para(tf.add_paragraph(),
         "Rodrigo Zambon  ·  [PREENCHER: e-mail]", 18, GOLD, bold=True,
         space_after=4)
set_para(tf.add_paragraph(),
         "Psike — plataforma de gestão de riscos psicossociais (NR-1)",
         15, RGBColor(0xC7,0xCE,0xDB))
notes(s, "Agradecer à banca e ao orientador. Sinalizar disponibilidade para a "
         "arguição. Ter o app aberto para eventual demonstração ao vivo.")

OUT = os.path.join(os.path.dirname(__file__), "Apresentacao_Banca_Psike.pptx")
prs.save(OUT)
print(f"OK: {os.path.basename(OUT)} gerado — {len(prs.slides.__iter__.__self__._sldIdLst)} slides.")
