# Psike — Contexto do projeto (para retomar no Claude Code)

> Cole este arquivo inteiro na primeira mensagem do Claude Code (ou aponte para ele
> com "leia CONTEXTO_PROJETO.md e continue o projeto") para que ele entenda tudo
> o que já foi feito sem precisar reexplicar.

## O que é

**Psike** — plataforma de gestão de SST (Segurança e Saúde do Trabalho), nascida
como TCC de Especialização em Engenharia de Segurança do Trabalho (EPUSP/PECE).
Tema do TCC: gestão digital de riscos psicossociais conforme a NR-1. A plataforma
foi desenhada para crescer em 4 módulos, todos sobre o mesmo repositório de dados
(inspirado no Inventário de Riscos Ocupacionais da NR-1).

## Estado atual (o que já está pronto e funcionando)

- **Módulo 1 — Riscos psicossociais (NR-1 / ISO 45003):** questionário de 10 itens
  em 6 dimensões, classificação automática de risco (baixo/médio/alto), painel
  agregado, geração automática de texto para o IRO/PGR.
- **Módulo 2 — APR / Permissão de Trabalho digital (NR-33/35/12):** checklist
  dinâmico por tipo de atividade, captura de geolocalização via GPS do navegador,
  assinatura digital em canvas, bloqueio do botão de liberação até tudo estar
  preenchido.
- **Módulos 3 e 4 (AEP/AET, investigação de acidentes/near-miss):** ainda não
  implementados — só descritos conceitualmente no TCC e na tela "Plataforma" do app.

## Arquitetura técnica

- **Frontend:** um único arquivo HTML autocontido (`psike-modulo1-riscos-psicossociais.html`),
  sem framework, sem build step — abre direto no navegador. CSS customizado
  (paleta navy `#16233F` + dourado `#B4842A`), fontes Fraunces/Inter/IBM Plex Mono
  via Google Fonts CDN.
- **Backend:** Google Apps Script (`psike-backend-google-apps-script.gs.txt`),
  gratuito, publicado como "App da Web" com acesso "Qualquer pessoa" (para
  respondentes anônimos sem precisar de login Google). Grava numa Google Sheets
  real (duas abas: `respostas` para o Módulo 1, `apr_registros` para o Módulo 2).
  ID da planilha do usuário: `1wM3I40onjQz0vNHc40SIgMNrPc7KZzmiHW_IKbiwJv4`.
  A URL de deployment atual (`/exec`) já está hardcoded na constante
  `SHEETS_WEBAPP_URL` no topo do `<script>` do HTML.
- **Camada de armazenamento no frontend (`Store` / `CloudStore` no JS):**
  detecção automática de ambiente — usa `window.storage` nativo quando roda dentro
  de um Artifact do Claude, cai para `localStorage` quando roda standalone sem
  backend configurado, e usa o Google Apps Script (`CloudStore`) quando a URL
  está configurada. Isso foi necessário porque a primeira versão só usava
  `window.storage` e quebrava fora do ambiente Claude.

## Decisões e problemas já resolvidos (não repetir)

1. `window.storage` só existe dentro do painel de Artifacts do Claude — corrigido
   com fallback para `localStorage`.
2. Colar o código do Apps Script pelo celular causava duplicação/erro de sintaxe
   (parênteses dessincronizados) — resolvido copiando de um `.txt` puro via
   computador em vez de copiar direto do chat.
3. `SpreadsheetApp.getActiveSpreadsheet()` não funciona em projeto avulso do
   Apps Script (só funciona se o script estiver "preso" a uma planilha) — trocado
   para `SpreadsheetApp.openById(SPREADSHEET_ID)` com o ID fixo.
4. Distribuição do HTML: recomendado hospedar via Netlify Drop
   (app.netlify.com/drop, arrastando o arquivo renomeado para `index.html`) para
   gerar um link público, em vez de mandar o arquivo `.html` por WhatsApp/e-mail.

## Limitações conhecidas / dívidas técnicas

- Sem autenticação/senha — qualquer pessoa com o link acessa o formulário e o painel.
- Sem separação de dados por empresa/cliente — hoje é uma planilha única.
- Sem testes automatizados.
- O TCC (`TCC_Riscos_Psicossociais_NR1_MODELO.docx`, gerado à parte, não incluído
  neste pacote) tem seções marcadas `[PREENCHER]` que dependem de dados reais de
  campo (nome da empresa, resultados aplicados, fotos) — não fabricar esses dados.

## Próximos passos sugeridos

1. Implementar Módulos 3 (AEP/AET) e 4 (investigação de acidentes/near-miss).
2. Adicionar autenticação simples (ex.: senha por empresa) antes de uso comercial real.
3. Migrar de Google Sheets para um banco de dados de verdade (Postgres/Supabase)
   se o volume de respostas crescer — Sheets tem limites de linhas/requisições.
4. Configurar deploy contínuo (ex.: repositório Git + Netlify conectado ao repo,
   em vez de upload manual via Netlify Drop).
