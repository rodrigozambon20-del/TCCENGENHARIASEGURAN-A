# Psike — Gestão digital de riscos psicossociais (NR-1)

Plataforma de gestão de SST (Segurança e Saúde do Trabalho) desenvolvida como TCC de
Especialização em Engenharia de Segurança do Trabalho (EPUSP/PECE).
Tema: **gestão digital de riscos psicossociais conforme a NR-1** (Portaria MTE 1.419/2024).

## Estrutura do repositório

| Caminho | Descrição |
|---|---|
| `index.html` | Aplicativo completo (frontend autocontido, sem build) — Módulos 1 a 4 |
| `backend/google-apps-script/Codigo.gs` | Backend em Google Apps Script (grava numa Google Sheets) |
| `docs/CONTEXTO_PROJETO.md` | Histórico do projeto: decisões, problemas resolvidos, dívidas técnicas |
| `docs/tcc/` | Monografia do TCC (`.docx`) e o script `gerar_tcc.py` que a regenera |

## Módulos

| Módulo | Status | Descrição |
|---|---|---|
| 1 — Riscos psicossociais | ✅ em uso | NR-1 / ISO 45003. Questionário de 10 itens em 6 dimensões, classificação automática de risco, painel agregado, texto pronto para o IRO/PGR |
| 2 — APR / Permissão de Trabalho digital | ✅ em uso | NR-10 / NR-35. Serviços em redes de distribuição de energia: rede desenergizada (desenergização NR-10), rede energizada/SEP e trabalho em altura em postes/estruturas. Checklist dinâmico, geolocalização, assinatura digital, bloqueio de liberação |
| 3 — Análise Ergonômica Preliminar (AEP) | ✅ em uso | NR-17. 12 itens em 5 blocos de fatores, classificação por bloco e parecer com indicação de AET quando necessário |
| 4 — Acidentes e quase acidentes | ✅ em uso | NR-1. Registro em campo, análise de causas em três níveis, indicadores e apoio ao registro da CAT |

## Como rodar

O frontend é um único arquivo HTML — basta abrir `index.html` no navegador.
Sem backend configurado, os dados ficam no `localStorage` do navegador.

### Backend (Google Sheets)

1. Crie um projeto em [script.google.com](https://script.google.com) e cole o conteúdo de
   `backend/google-apps-script/Codigo.gs`.
2. Ajuste `SPREADSHEET_ID` para o ID da sua planilha.
3. Implante como **App da Web** com acesso "Qualquer pessoa".
4. Cole a URL `/exec` gerada na constante `SHEETS_WEBAPP_URL` no topo do `<script>`
   do `index.html` (ou na aba **Configurações** do app, para uso só neste navegador).

### Publicação

O arquivo já se chama `index.html`, pronto para deploy contínuo: conecte este
repositório ao Netlify (ou similar) e publique a raiz do projeto. Alternativa manual:
[Netlify Drop](https://app.netlify.com/drop).

## Limitações conhecidas

- Sem autenticação — qualquer pessoa com o link acessa formulário e painel.
- Planilha única, sem separação por empresa/cliente.
- Sem testes automatizados.

Detalhes e próximos passos em `docs/CONTEXTO_PROJETO.md`.

---

*Protótipo acadêmico · dados coletados de forma anônima e agregada, conforme LGPD.*
