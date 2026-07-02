/**
 * PSIKE — Backend em Google Apps Script + Google Sheets
 * ---------------------------------------------------------------
 * Suporta quatro módulos da plataforma, cada um em sua própria aba
 * dentro da mesma planilha:
 *   - psico      -> Módulo 1: Riscos psicossociais (NR-1)
 *   - apr        -> Módulo 2: Permissão de Trabalho / APR digital (NR-10/35)
 *   - aep        -> Módulo 3: Análise Ergonômica Preliminar (NR-17)
 *   - acidentes  -> Módulo 4: Acidentes e quase acidentes (NR-1)
 *
 * Este arquivo já está configurado para a planilha:
 * ID: 1wM3I40onjQz0vNHc40SIgMNrPc7KZzmiHW_IKbiwJv4
 *
 * SE VOCÊ JÁ TINHA IMPLANTADO UMA VERSÃO ANTERIOR (com menos módulos):
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

const AEP_SHEET_NAME = "aep_registros";
const AEP_HEADERS = [
  "timestamp", "setor", "funcao", "posto", "avaliador",
  "itens_json", "blocos_json", "observacoes", "conclusao"
];

const ACID_SHEET_NAME = "acidentes_registros";
const ACID_HEADERS = [
  "timestamp", "tipo", "data_ocorrencia", "local", "setor", "atividade",
  "descricao", "lesao", "causas_imediatas", "causas_subjacentes",
  "causas_basicas", "acoes", "cat_emitida", "cat_numero"
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
function getAepSheet_() { return getSheetGeneric_(AEP_SHEET_NAME, AEP_HEADERS); }
function getAcidSheet_() { return getSheetGeneric_(ACID_SHEET_NAME, ACID_HEADERS); }

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * GET ?action=ping
 * GET ?module=psico&action=list
 * GET ?module=apr&action=list
 * GET ?module=aep&action=list
 * GET ?module=acidentes&action=list
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

  if (action === "list" && module === "aep") {
    const sheet = getAepSheet_();
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return jsonResponse_({ ok: true, records: [] });
    const [header, ...rows] = values;
    const records = rows.map(row => {
      let itens = {};
      let blocos = [];
      try { itens = JSON.parse(row[5] || "{}"); } catch (err) { itens = {}; }
      try { blocos = JSON.parse(row[6] || "[]"); } catch (err) { blocos = []; }
      return {
        ts: row[0],
        setor: row[1],
        funcao: row[2],
        posto: row[3],
        avaliador: row[4],
        itens: itens,
        blocos: blocos,
        observacoes: row[7],
        conclusao: row[8],
      };
    });
    return jsonResponse_({ ok: true, records });
  }

  if (action === "list" && module === "acidentes") {
    const sheet = getAcidSheet_();
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return jsonResponse_({ ok: true, records: [] });
    const [header, ...rows] = values;
    const records = rows.map(row => {
      return {
        ts: row[0],
        tipo: row[1],
        data_ocorrencia: row[2],
        local: row[3],
        setor: row[4],
        atividade: row[5],
        descricao: row[6],
        lesao: row[7],
        causas_imediatas: row[8],
        causas_subjacentes: row[9],
        causas_basicas: row[10],
        acoes: row[11],
        cat_emitida: row[12],
        cat_numero: row[13],
      };
    });
    return jsonResponse_({ ok: true, records });
  }

  return jsonResponse_({ ok: false, error: "ação/módulo desconhecido" });
}

/**
 * POST body JSON, campo "module": "psico" | "apr" | "aep" | "acidentes"
 *
 * module=psico, action=submit  -> { setor, answers: {0:5,1:3,...} }
 * module=psico, action=reset
 *
 * module=apr,   action=submit  -> { atividade, local, responsavel,
 *                                    checklist:[{item,ok}], geolat, geolng,
 *                                    assinante, liberado }
 * module=apr,   action=reset
 *
 * module=aep,   action=submit  -> { setor, funcao, posto, avaliador,
 *                                    itens:{0:1,1:2,...}, blocos:[{...}],
 *                                    observacoes, conclusao }
 * module=aep,   action=reset
 *
 * module=acidentes, action=submit -> { tipo, data_ocorrencia, local, setor,
 *                                    atividade, descricao, lesao,
 *                                    causas_imediatas, causas_subjacentes,
 *                                    causas_basicas, acoes, cat_emitida,
 *                                    cat_numero }
 * module=acidentes, action=reset
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

  if (module === "aep") {
    const sheet = getAepSheet_();

    if (body.action === "reset") {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
      return jsonResponse_({ ok: true, cleared: true });
    }

    if (body.action === "submit") {
      const row = [
        new Date().toISOString(),
        body.setor || "Não informado",
        body.funcao || "",
        body.posto || "",
        body.avaliador || "",
        JSON.stringify(body.itens || {}),
        JSON.stringify(body.blocos || []),
        body.observacoes || "",
        body.conclusao || "",
      ];
      sheet.appendRow(row);
      return jsonResponse_({ ok: true });
    }
  }

  if (module === "acidentes") {
    const sheet = getAcidSheet_();

    if (body.action === "reset") {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
      return jsonResponse_({ ok: true, cleared: true });
    }

    if (body.action === "submit") {
      const row = [
        new Date().toISOString(),
        body.tipo || "",
        body.data_ocorrencia || "",
        body.local || "",
        body.setor || "Não informado",
        body.atividade || "",
        body.descricao || "",
        body.lesao || "",
        body.causas_imediatas || "",
        body.causas_subjacentes || "",
        body.causas_basicas || "",
        body.acoes || "",
        body.cat_emitida ? "SIM" : "NÃO",
        body.cat_numero || "",
      ];
      sheet.appendRow(row);
      return jsonResponse_({ ok: true });
    }
  }

  return jsonResponse_({ ok: false, error: "ação/módulo desconhecido" });
}
