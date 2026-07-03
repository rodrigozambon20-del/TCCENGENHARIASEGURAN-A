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
    "A Portaria MTE nº 1.419/2024 atualizou a NR-1 para exigir, de todo "
    "empregador CLT, a identificação, a avaliação e a documentação dos fatores "
    "de risco psicossocial relacionados ao trabalho dentro do Gerenciamento de "
    "Riscos Ocupacionais (GRO) e do PGR, com exigibilidade a partir de 26 de "
    "maio de 2026.",
    "A urgência é reforçada pelo reconhecimento da síndrome de burnout como "
    "doença ocupacional na CID-11 (código QD85), adotada oficialmente no Brasil "
    "a partir de 2025, e pelo crescimento dos afastamentos por saúde mental "
    "(TREML et al., 2025; ANAMT, 2026).",
    "Na maioria das organizações esse ciclo ainda é feito em papel ou planilhas "
    "avulsas, com baixa rastreabilidade e alto custo. No setor de distribuição "
    "de energia elétrica, os fatores psicossociais (pressão por "
    "restabelecimento, turnos, sobreaviso, risco de acidente grave) convivem "
    "com riscos regulados pela NR-10 e pela NR-35 (SOUZA et al., 2010).",
    "Objetivo: desenvolver e aplicar uma plataforma digital de baixo custo para "
    "operacionalizar a gestão de riscos psicossociais conforme a NR-1, "
    "integrada a outros processos de SST sobre a mesma base de dados.",
]

METODOLOGIA = [
    "Pesquisa aplicada, de natureza tecnológica: desenvolvimento de um artefato "
    "(a plataforma Psike) seguido de estudo de caso com aplicação piloto.",
    "Arquitetura de custo zero: frontend em página web autocontida (sem "
    "instalação) e backend gratuito em Google Apps Script/Sheets como "
    "repositório único. Coleta anônima por construção, em conformidade com a "
    "LGPD.",
    "Instrumento de avaliação: 10 itens em 6 dimensões (carga e ritmo; "
    "autonomia e controle; clareza de papel; apoio social e de liderança; "
    "reconhecimento; assédio e violência), em escala Likert de 5 pontos, com "
    "classificação automática do risco por dimensão em três faixas.",
    "A plataforma reúne quatro módulos sobre a mesma base: (1) riscos "
    "psicossociais (NR-1/ISO 45003); (2) permissão de trabalho para redes de "
    "distribuição (NR-10/NR-35); (3) Avaliação Ergonômica Preliminar (NR-17); "
    "(4) registro e investigação de acidentes e quase acidentes.",
]

ESTUDO_CASO = [
    "A aplicação piloto foi conduzida em uma organização do setor de "
    "distribuição de energia elétrica, preservada em anonimato, no período de "
    "[PREENCHER: período], com [PREENCHER: número] trabalhadores dos setores de "
    "[PREENCHER: setores]. Não se coletam dados que identifiquem a organização "
    "ou as pessoas.",
    "O link da avaliação anônima foi distribuído por [PREENCHER: canal — ex.: "
    "QR code em DDS], precedido de comunicação sobre o caráter voluntário e "
    "anônimo da participação, em conformidade com a LGPD.",
    "As permissões de trabalho (Módulo 2) foram testadas em campo em atividades "
    "de [PREENCHER: ex.: manutenção de rede desenergizada / linha viva / "
    "trabalho em altura em postes], com checklist, geolocalização e assinatura "
    "digital.",
]

RESULTADOS = [
    "A plataforma foi implementada integralmente, com os quatro módulos "
    "operacionais. O fluxo do Módulo 1 — da resposta anônima no celular ao "
    "texto pronto para o Inventário de Riscos Ocupacionais (IRO) — ocorre sem "
    "qualquer tabulação manual, a custo de operação nulo.",
    "[PREENCHER: inserir os resultados reais — nº de respondentes, taxa de "
    "adesão, nota média e faixa de risco por dimensão. Substituir as figuras de "
    "exemplo ao lado pelo gráfico do painel por dimensão e por telas do app.]",
    "Discussão: a principal barreira à conformidade com a nova NR-1 é "
    "operacional, não conceitual. A digitalização de ponta a ponta reduziu o "
    "custo do ciclo e melhorou o anonimato, a consistência da classificação e "
    "a rastreabilidade dos registros exigida pela norma.",
]

CONCLUSAO = [
    "A digitalização viabiliza o cumprimento da NR-1 quanto aos fatores "
    "psicossociais mesmo em organizações com equipes de SST enxutas, a custo "
    "praticamente nulo, e transforma o inventário de riscos em um documento "
    "vivo, alimentado pelos próprios processos operacionais.",
    "Trabalhos futuros: validação psicométrica do instrumento com amostras "
    "maiores, calibração dos pontos de corte com séries históricas, "
    "autenticação e segregação de dados por empresa e acompanhamento "
    "longitudinal da eficácia das medidas de controle.",
]

REFERENCIAS = [
    "BRASIL. NR-1 — Disposições gerais e gerenciamento de riscos ocupacionais. "
    "Redação da Portaria MTE nº 1.419/2024.",
    "BRASIL. NR-10 — Segurança em instalações e serviços em eletricidade.",
    "ISO 45003:2021 — Psychological health and safety at work.",
    "SOUZA, S. F. et al. Fatores psicossociais do trabalho e transtornos "
    "mentais comuns em eletricitários. Rev. Saúde Pública, v. 44, n. 4, 2010.",
    "TREML, M. F. Q. et al. Burnout syndrome in Brazil (2014–2024). Rev. Bras. "
    "Medicina do Trabalho, 2025.",
    "WHO. Guidelines on mental health at work. Geneva, 2022.",
]

HEADER = {
    "tema": "GESTÃO DIGITAL DE RISCOS PSICOSSOCIAIS CONFORME A NR-1: "
            "plataforma de SST aplicada à distribuição de energia elétrica",
    "autor": "Rodrigo Zambon [PREENCHER: nome completo]",
    "contato": "[PREENCHER: e-mail]  ·  Supervisão: [PREENCHER: nome]  ·  "
               "Apoio: CERPRO",
}

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

def replace_body(text_frame, paragraphs):
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
    # remove todos os parágrafos exceto o título
    for p in list(txBody.findall(qn("a:p"))):
        if p is not title_p:
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
for name, content in SECTIONS.items():
    if name in by_name:
        replace_body(by_name[name].text_frame, content)

# Cabeçalho (TEMA / autor / contato)
if "CaixaDeTexto 2" in by_name:
    replace_lines(by_name["CaixaDeTexto 2"].text_frame,
                  [HEADER["tema"], HEADER["autor"], HEADER["contato"]])

prs.save(OUT)
print(f"OK: {OUT} gerado a partir de {SRC}.")
