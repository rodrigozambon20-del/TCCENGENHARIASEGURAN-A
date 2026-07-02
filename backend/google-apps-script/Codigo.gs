/**
 * PSIKE — Backend em Google Apps Script + Google Sheets
 * ---------------------------------------------------------------
 * Suporta dois módulos da plataforma, cada um em sua própria aba
 * dentro da mesma planilha:
 *   - psico  -> Módulo 1: Riscos psicossociais (NR-1)
 *   - apr    -> Módulo 2: Permissão de Trabalho / APR digital (NR-33/35/12)
 *
 * Este arquivo já está configurado para a planilha:
 * ID: 1wM3I40onjQz0vNHc40SIgMNrPc7KZzmiHW_IKbiwJv4
 *
 * SE VOCÊ JÁ TINHA IMPLANTADO A VERSÃO ANTERIOR (só com o Módulo 1):
 * não precisa criar uma implantação nova. Substitua todo o conteúdo
 * do arquivo Código.gs por este aqui, salve, e vá em
 * Implantar > Gerenciar implantações > ícone de lápis > Nova versão.
 * A URL do app da Web continua a mesma.
 * ---------------------------------------------------------------
 */

const SPREADSHEET_ID = "1wM3I40onjQz0vNHc40SIgMNrPc7KZzmiHW_IKbiwJv4";

const PSICO_SHEET_NAME = "respostas";
const PSICO_HEADERS = ["timestamp", "setor", "q0", "q1", "q2", "q3", "q4", "q5", "q6", "q7", "q8", "q9"];

const APR_SHEET_NAME = "apr_registros";
const APR_HEADERS = [
  "timestamp", "atividade", "local", "responsavel",
  "checklist_json", "geolat", "geolng", "assinante", "liberado"
];

function getSheetGeneric_(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function getPsicoSheet_() { return getSheetGeneric_(PSICO_SHEET_NAME, PSICO_HEADERS); }
function getAprSheet_() { return getSheetGeneric_(APR_SHEET_NAME, APR_HEADERS); }

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * GET ?action=ping
 * GET ?module=psico&action=list
 * GET ?module=apr&action=list
 */
function doGet(e) {
  const action = e.parameter.action || "list";
  const module = e.parameter.module || "psico";

  if (action === "ping") {
    return jsonResponse_({ ok: true, service: "psike-backend", time: new Date().toISOString() });
  }

  if (action === "list" && module === "psico") {
    const sheet = getPsicoSheet_();
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return jsonResponse_({ ok: true, records: [] });
    const [header, ...rows] = values;
    const records = rows.map(row => {
      const rec = { setor: row[1], ts: row[0], answers: {} };
      for (let i = 0; i < 10; i++) {
        const v = row[2 + i];
        if (v !== "" && v !== null && v !== undefined) rec.answers[i] = Number(v);
      }
      return rec;
    });
    return jsonResponse_({ ok: true, records });
  }

  if (action === "list" && module === "apr") {
    const sheet = getAprSheet_();
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return jsonResponse_({ ok: true, records: [] });
    const [header, ...rows] = values;
    const records = rows.map(row => {
      let checklist = [];
      try { checklist = JSON.parse(row[4] || "[]"); } catch (err) { checklist = []; }
      return {
        ts: row[0],
        atividade: row[1],
        local: row[2],
        responsavel: row[3],
        checklist: checklist,
        geolat: row[5],
        geolng: row[6],
        assinante: row[7],
        liberado: row[8],
      };
    });
    return jsonResponse_({ ok: true, records });
  }

  return jsonResponse_({ ok: false, error: "ação/módulo desconhecido" });
}

/**
 * POST body JSON, campo "module": "psico" | "apr"
 *
 * module=psico, action=submit  -> { setor, answers: {0:5,1:3,...} }
 * module=psico, action=reset
 *
 * module=apr,   action=submit  -> { atividade, local, responsavel,
 *                                    checklist:[{item,ok}], geolat, geolng,
 *                                    assinante, liberado }
 * module=apr,   action=reset
 */
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse_({ ok: false, error: "corpo inválido" });
  }

  const module = body.module || "psico";

  if (module === "psico") {
    const sheet = getPsicoSheet_();

    if (body.action === "reset") {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
      return jsonResponse_({ ok: true, cleared: true });
    }

    if (body.action === "submit") {
      const answers = body.answers || {};
      const row = [new Date().toISOString(), body.setor || "Não informado"];
      for (let i = 0; i < 10; i++) {
        row.push(answers[i] !== undefined ? answers[i] : "");
      }
      sheet.appendRow(row);
      return jsonResponse_({ ok: true });
    }
  }

  if (module === "apr") {
    const sheet = getAprSheet_();

    if (body.action === "reset") {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
      return jsonResponse_({ ok: true, cleared: true });
    }

    if (body.action === "submit") {
      const row = [
        new Date().toISOString(),
        body.atividade || "",
        body.local || "",
        body.responsavel || "",
        JSON.stringify(body.checklist || []),
        body.geolat || "",
        body.geolng || "",
        body.assinante || "",
        body.liberado ? "SIM" : "NÃO",
      ];
      sheet.appendRow(row);
      return jsonResponse_({ ok: true });
    }
  }

  return jsonResponse_({ ok: false, error: "ação/módulo desconhecido" });
}
