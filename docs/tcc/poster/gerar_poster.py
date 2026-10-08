# -*- coding: utf-8 -*-
"""
Gerador do pôster do TCC (Psike) a partir do modelo oficial P_STER_MODELO_2024.pptx.

Preenche as caixas de texto do modelo (Introdução/Objetivo, Metodologia,
Estudo de Caso, Resultados e Discussão, Conclusão, Referências e o cabeçalho
tema/autor/e-mail) com o conteúdo da monografia, PRESERVANDO o layout, as
cores, as fontes e as logomarcas do modelo. Trechos [PREENCHER: ...] dependem
de dados reais de campo e ficam para o autor completar; as figuras de exemplo
do modelo devem ser trocadas por gráficos/telas reais no PowerPoint.

Uso:  python3 gerar_poster.py
Requer: pip install python-pptx
Entrada: _modelo_original.pptx (o modelo oficial, versionado ao lado)
Saída:   Poster_Psike_TCC.pptx  (exportar para PDF no PowerPoint para entrega)
"""
import copy
from pptx import Presentation
from pptx.oxml.ns import qn

SRC = "_modelo_original.pptx"
OUT = "Poster_Psike_TCC.pptx"

# ---- Conteúdo por seção (título mantido do modelo; corpo substituído) --------

INTRO = [
    "A NR-1 exige que o empregador analise os acidentes e as doenças "
    "relacionadas ao trabalho, identifique suas causas e realimente o PGR "
    "(GRO). Na prática, o registro segue manual e tardio, a investigação para "
    "no \u201cato inseguro\u201d e os quase acidentes raramente são reportados.",
    "No setor elétrico a gravidade é extrema: 2.089 acidentes de origem "
    "elétrica no Brasil em 2023, com 781 óbitos (Abracopel, 2024); 250 mortes "
    "envolvendo a rede de distribuição (Acidentes [...], 2024). Cada ocorrência — e "
    "cada quase acidente — é aprendizado que não pode ser desperdiçado.",
    "A literatura brasileira consolidou o método da árvore de causas como "
    "antídoto à investigação culpabilizadora, revelando os fatores gerenciais "
    "e organizacionais na gênese dos acidentes (Binder; Almeida, 1997).",
    "Objetivo: desenvolver e aplicar um sistema digital de registro e "
    "investigação de acidentes e quase acidentes estruturado no método da "
    "árvore de causas, integrado ao GRO/PGR da NR-1, em uma organização de "
    "distribuição de energia elétrica.",
]

METODOLOGIA = [
    "Pesquisa aplicada, de natureza tecnológica: desenvolvimento de artefato "
    "(o sistema digital, módulo central da plataforma Psike) seguido de estudo "
    "de caso — análise documental retrospectiva de ocorrências anonimizadas e "
    "piloto de campo.",
    "Arquitetura de custo zero: frontend em página web autocontida (sem "
    "instalação, uso no celular em campo) e backend gratuito em nuvem como "
    "repositório único. Tratamento de dados anonimizado, conforme a LGPD.",
    "Registro em campo em menos de 5 minutos: classificação (acidente com/sem "
    "afastamento, quase acidente, condição insegura), descrição, tarefa e "
    "alerta automático de CAT. Investigação guiada em 3 níveis de causas: "
    "imediatas → subjacentes → básicas — o formulário não permite concluir só "
    "com causas imediatas.",
    "Ações corretivas/preventivas com responsável e prazo; indicadores "
    "automáticos: razão quase acidentes/acidentes, tempo evento–registro, "
    "ações no prazo e taxas de frequência e gravidade (NBR 14280). Módulos "
    "complementares na mesma base: permissão de trabalho (NR-10/NR-35), "
    "fatores psicossociais (NR-1) e AEP (NR-17).",
]

ESTUDO_CASO = [
    "Organização de distribuição de energia elétrica, preservada em anonimato. "
    "Parte documental: [PREENCHER: nº] ocorrências históricas anonimizadas "
    "reinvestigadas em 3 níveis. Parte de campo: piloto em [PREENCHER: "
    "período] com as equipes de [PREENCHER: escopo], registrando ocorrências "
    "pelo celular.",
]

RESULTADOS = [
    "Sistema operacional, sem transcrição manual e com custo nulo. [PREENCHER: "
    "principal resultado documental e do piloto.] Ao exigir as causas básicas, "
    "a investigação alcança os fatores organizacionais que a análise "
    "tradicional não vê (Binder; Almeida, 1997; Reason, 1997).",
]

CONCLUSAO = [
    "A digitalização do ciclo registro–investigação–ação torna exequível, a "
    "custo praticamente nulo, a exigência da NR-1 de analisar ocorrências e "
    "realimentar o PGR — deslocando a análise da culpabilização individual "
    "para os fatores organizacionais e reduzindo a subnotificação de quase "
    "acidentes.",
    "Trabalhos futuros: diagrama completo da árvore de causas no software, "
    "autenticação e segregação por organização, acompanhamento longitudinal "
    "dos indicadores e cruzamento analítico entre ocorrências, permissões de "
    "trabalho e sinalização psicossocial.",
]

REFERENCIAS = [
    "ABRACOPEL. Anuário estatístico de acidentes de origem elétrica 2024: ano "
    "base 2023. Abracopel, 2024.",
    "BINDER, M. C. P.; ALMEIDA, I. M. Estudo de caso de dois acidentes do "
    "trabalho investigados com o método de árvore de causas. Cadernos de Saúde "
    "Pública, v. 13, n. 4, p. 749-760, 1997.",
    "BRASIL. MTE. NR-1: gerenciamento de riscos ocupacionais. Portaria MTE nº "
    "1.419/2024.",
    "ABNT. NBR 14280: cadastro de acidente do trabalho. Rio de Janeiro, 2001.",
    "REASON, J. Managing the risks of organizational accidents. Ashgate, 1997.",
    "HEINRICH, H. W. Industrial accident prevention. McGraw-Hill, 1931.",
]

HEADER = {
    "tema": "INVESTIGAÇÃO DIGITAL DE ACIDENTES PELA ÁRVORE DE CAUSAS NA "
            "DISTRIBUIÇÃO DE ENERGIA ELÉTRICA",
    "autor": "Rodrigo Zambon [PREENCHER: nome completo]",
    "contato": "[PREENCHER: e-mail]  ·  Apoio: CERPRO",
}
HEADER_SIZES = (36, 28, 24)

# ------------------------------------------------------------------ helpers ---

def _first_run(paragraph):
    runs = paragraph.findall(qn("a:r"))
    return runs[0] if runs else None

def _run_with_text(paragraph):
    """Primeiro run que possui elemento <a:t> (o que carrega texto de fato);
    cai para o primeiro run se nenhum tiver."""
    runs = paragraph.findall(qn("a:r"))
    for r in runs:
        if r.find(qn("a:t")) is not None:
            return r
    return runs[0] if runs else None

def replace_body(text_frame, paragraphs, keep_gap=False):
    """Mantém o 1º parágrafo (título da seção) e substitui o corpo,
    clonando o parágrafo de corpo do modelo para preservar fonte/cor/tamanho."""
    ps = text_frame.paragraphs
    if len(ps) < 2 or not paragraphs:
        return
    txBody = text_frame._txBody
    title_p = ps[0]._p
    # modelo de corpo = primeiro parágrafo (após o título) que tenha texto real
    tmpl_p = None
    for p in ps[1:]:
        if _run_with_text(p._p) is not None:
            tmpl_p = p._p
            break
    if tmpl_p is None:
        tmpl_p = ps[1]._p
    # mantém o título e as linhas em branco que o modelo reserva para imagens
    keep_ps = [title_p]
    if keep_gap:
        for p in ps[1:]:
            if p.text.strip():
                break
            keep_ps.append(p._p)
    for p in list(txBody.findall(qn("a:p"))):
        if not any(p is k for k in keep_ps):
            txBody.remove(p)
    # recria o corpo a partir do template clonado
    for txt in paragraphs:
        newp = copy.deepcopy(tmpl_p)
        keep = _run_with_text(newp)
        for r in newp.findall(qn("a:r")):
            if r is not keep:
                newp.remove(r)
        if keep is not None:
            t = keep.find(qn("a:t"))
            if t is None:
                t = keep.makeelement(qn("a:t"), {})
                keep.append(t)
            t.text = txt
        txBody.append(newp)

def replace_lines(text_frame, lines):
    """Substitui texto linha a linha (cabeçalho), preservando formatação."""
    ps = text_frame.paragraphs
    for i, line in enumerate(lines):
        if i >= len(ps):
            break
        r = _first_run(ps[i]._p)
        if r is not None:
            t = r.find(qn("a:t"))
            if t is not None:
                t.text = line
        # remove runs extras da linha para não sobrar lorem
        for extra in ps[i]._p.findall(qn("a:r"))[1:]:
            ps[i]._p.remove(extra)

# --------------------------------------------------------------------- main ---

prs = Presentation(SRC)
slide = prs.slides[0]

by_name = {sh.name: sh for sh in slide.shapes if sh.has_text_frame}

SECTIONS = {
    "CaixaDeTexto 25": INTRO,          # Introdução/Objetivo
    "CaixaDeTexto 29": METODOLOGIA,    # Metodologia
    "CaixaDeTexto 27": ESTUDO_CASO,    # Estudo de Caso
    "CaixaDeTexto 44": RESULTADOS,     # Resultados e Discussão
    "CaixaDeTexto 46": CONCLUSAO,      # Conclusão
    "CaixaDeTexto 47": REFERENCIAS,    # Referências
}
WITH_IMAGES = {"CaixaDeTexto 27", "CaixaDeTexto 44"}
for name, content in SECTIONS.items():
    if name in by_name:
        replace_body(by_name[name].text_frame, content,
                     keep_gap=name in WITH_IMAGES)

# Cabeçalho (TEMA / autor / contato)
if "CaixaDeTexto 2" in by_name:
    replace_lines(by_name["CaixaDeTexto 2"].text_frame,
                  [HEADER["tema"], HEADER["autor"], HEADER["contato"]])
    from pptx.util import Pt as _HPt
    for para, size in zip(by_name["CaixaDeTexto 2"].text_frame.paragraphs,
                          HEADER_SIZES):
        for r in para.runs:
            r.font.size = _HPt(size)

# Imagens de exemplo do modelo da USP -> espaços reservados identificados.
# A aula da Turma 2025 torna OBRIGATÓRIAS as fotos do estudo de caso no
# pôster (com rostos e logotipos ocultados); os gráficos devem vir do painel.
from pptx.dml.color import RGBColor
from pptx.util import Pt as _Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
PLACEHOLDERS = {
    "Picture 6": "[PREENCHER: FOTO 1 do estudo de caso — rostos e logotipos ocultados]",
    "Picture 8": "[PREENCHER: FOTO 2 do estudo de caso — rostos e logotipos ocultados]",
    "Picture 2": "[PREENCHER: gráfico real do painel — causas por nível]",
    "Picture 4": "[PREENCHER: gráfico real do painel — indicadores do piloto]",
}
all_shapes = {sh.name: sh for sh in slide.shapes}
from pptx.util import Cm as _Cm
fig = slide.shapes.add_shape(1, _Cm(2.3), _Cm(41.0), _Cm(40.0), _Cm(14.0))
fig.fill.solid(); fig.fill.fore_color.rgb = RGBColor(0xFF, 0xF2, 0xA8)
fig.line.color.rgb = RGBColor(0xB4, 0x84, 0x2A); fig.line.width = _Pt(3)
ftf = fig.text_frame; ftf.word_wrap = True; ftf.vertical_anchor = MSO_ANCHOR.MIDDLE
fpara = ftf.paragraphs[0]; fpara.alignment = PP_ALIGN.CENTER
frun = fpara.add_run()
frun.text = ("[PREENCHER: FIGURA do método — exemplo de árvore de causas ou "
             "fluxo registro → 3 níveis de causas → ações]")
frun.font.size = _Pt(28); frun.font.bold = True
frun.font.color.rgb = RGBColor(0x16, 0x23, 0x3F)
for pic_name, label in PLACEHOLDERS.items():
    pic = all_shapes.get(pic_name)
    if pic is None:
        continue
    box = slide.shapes.add_shape(1, pic.left, pic.top, pic.width, pic.height)
    box.fill.solid(); box.fill.fore_color.rgb = RGBColor(0xFF, 0xF2, 0xA8)
    box.line.color.rgb = RGBColor(0xB4, 0x84, 0x2A); box.line.width = _Pt(3)
    tf = box.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]; para.alignment = PP_ALIGN.CENTER
    run = para.add_run(); run.text = label
    run.font.size = _Pt(28); run.font.bold = True
    run.font.color.rgb = RGBColor(0x16, 0x23, 0x3F)
    pic._element.getparent().remove(pic._element)

prs.save(OUT)
print(f"OK: {OUT} gerado a partir de {SRC}.")
