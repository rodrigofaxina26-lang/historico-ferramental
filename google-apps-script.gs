/**
 * Google Apps Script - Web App / API central do "Histórico Ferramental".
 *
 * Funções:
 *  - Banco de dados central (planilha): guarda "Ferramentas" e "Ciclos" para
 *    que todas as estações (produção, qualidade, ferramentaria) vejam os
 *    mesmos dados.
 *  - Envio de e-mail automático na saída da ferramenta (com corpo HTML).
 *  - Log opcional de cada saída na aba "Histórico".
 *
 * COMO IMPLANTAR:
 * 1. Acesse https://script.google.com e crie um novo projeto (pode ser com a
 *    mesma conta Gmail/Workspace que deve aparecer como remetente).
 * 2. Apague o conteúdo padrão de "Code.gs" e cole todo o código abaixo.
 * 3. (Opcional, recomendado) Troque o valor de TOKEN por uma senha/segredo
 *    qualquer, ex: "promf-2026-xyz". Esse token evita que outras pessoas usem
 *    sua URL para ler/gravar os dados ou enviar e-mails em seu nome.
 * 4. Crie uma planilha nova no Google Sheets, copie o ID dela (a parte da URL
 *    entre /d/ e /edit) e cole na constante SHEET_ID abaixo. A partir de
 *    agora SHEET_ID é OBRIGATÓRIO — é onde ficam os dados centrais
 *    ("Ferramentas", "Ciclos") e o log opcional ("Histórico"). As abas e
 *    cabeçalhos são criados automaticamente na primeira vez.
 * 5. Clique em "Implantar" > "Nova implantação" (ou, se já existe uma
 *    implantação, "Gerenciar implantações" > editar > "Nova versão").
 *    - Tipo: "App da Web" (Web app)
 *    - Executar como: "Eu" (sua conta)
 *    - Quem pode acessar: "Qualquer pessoa"
 * 6. Autorize as permissões solicitadas (envio de e-mail e acesso à planilha
 *    em seu nome).
 * 7. Copie a URL gerada (termina em /exec) e cole no campo "URL do Web App"
 *    da seção "Central de Dados (Apps Script)" do Histórico Ferramental, em
 *    TODOS os computadores. Cole o mesmo TOKEN no campo "Token de segurança".
 *
 * Limite gratuito de e-mail: contas Gmail comuns ~100 e-mails/dia; contas
 * Google Workspace, ~1500 e-mails/dia.
 */

const TOKEN = ""; // defina um token secreto aqui (mesmo valor configurado no app)
const SHEET_ID = ""; // ID da planilha Google Sheets (obrigatório para Central de Dados)
const HIST_SHEET_NAME = "Histórico";

// E-mails que recebem o resumo semanal de ferramentas precisando de ajuste
// (função enviarResumoSemanal — requer um Acionador de tempo, veja final do arquivo)
const RESUMO_SEMANAL_DESTINATARIOS = [
  // "qualidade@empresa.com",
  // "ferramentaria@empresa.com",
];

const HIST_HEADERS = [
  "Data/Hora Saída", "Nº Ferramenta", "Código PROMETAL", "C.C.", "Cliente",
  "Descrição da Operação", "Nº OP.", "Descrição do Produto", "Cód. Prod. Cliente",
  "Ciclo Nº", "Entrada", "Preparador Entrada", "Máquina",
  "Preparador Saída", "Peças Produzidas", "Peças Acumuladas",
  "Parecer Produção", "Data Parecer Produção", "Responsável Produção",
  "Parecer Qualidade", "Data Parecer Qualidade", "Responsável Qualidade",
  "Assinatura Produção", "Data Assinatura Produção", "Local Assinatura Produção",
  "Assinatura Qualidade", "Data Assinatura Qualidade", "Local Assinatura Qualidade"
];

// ── ENTRY POINTS ──────────────────────────────────────────────────────────────

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    if (TOKEN && params.token !== TOKEN) {
      return jsonResponse({ status: "error", message: "Token inválido." });
    }
    if (params.action === "getData") {
      return jsonResponse({
        status: "ok",
        tools: readRecords("Ferramentas"),
        cycles: readRecords("Ciclos")
      });
    }
    return jsonResponse({ status: "error", message: "Ação desconhecida." });
  } catch (err) {
    return jsonResponse({ status: "error", message: String(err) });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (TOKEN && data.token !== TOKEN) {
      return jsonResponse({ status: "error", message: "Token inválido." });
    }

    switch (data.action) {
      case "upsertTool":
        return upsertRecord("Ferramentas", data.item);
      case "deleteTool":
        return deleteToolCascade(data.id);
      case "upsertCycle":
        return upsertRecord("Ciclos", data.item);
      case "deleteCycle":
        return deleteRecord("Ciclos", data.id);
      case "bulkImport":
        return bulkImport(data.tools, data.cycles);
      case "sendEmail":
        return handleSendEmail(data);
      default:
        return jsonResponse({ status: "error", message: "Ação desconhecida." });
    }
  } catch (err) {
    return jsonResponse({ status: "error", message: String(err) });
  }
}

// ── DADOS CENTRAIS (Ferramentas / Ciclos) ───────────────────────────────────────

function getDataSheet(sheetName) {
  if (!SHEET_ID) throw new Error("SHEET_ID não configurado no script.");
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["id", "json", "updatedAt"]);
  }
  return sheet;
}

function readRecords(sheetName) {
  const sheet = getDataSheet(sheetName);
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const json = rows[i][1];
    if (!json) continue;
    try { out.push(JSON.parse(json)); } catch (err) { /* linha inválida, ignora */ }
  }
  return out;
}

function upsertRecord(sheetName, item) {
  if (!item || !item.id) return jsonResponse({ status: "error", message: "Registro sem id." });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getDataSheet(sheetName);
    const ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 1).getValues();
    const now = new Date().toISOString();
    let rowIndex = -1;
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === item.id) { rowIndex = i + 2; break; }
    }
    const rowValues = [item.id, JSON.stringify(item), now];
    if (rowIndex === -1) {
      sheet.appendRow(rowValues);
    } else {
      sheet.getRange(rowIndex, 1, 1, 3).setValues([rowValues]);
    }
    return jsonResponse({ status: "ok" });
  } finally {
    lock.releaseLock();
  }
}

function deleteRecord(sheetName, id) {
  if (!id) return jsonResponse({ status: "error", message: "id não informado." });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getDataSheet(sheetName);
    const ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (ids[i][0] === id) sheet.deleteRow(i + 2);
    }
    return jsonResponse({ status: "ok" });
  } finally {
    lock.releaseLock();
  }
}

function deleteToolCascade(id) {
  if (!id) return jsonResponse({ status: "error", message: "id não informado." });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Remove a ferramenta
    let sheet = getDataSheet("Ferramentas");
    let ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (ids[i][0] === id) sheet.deleteRow(i + 2);
    }
    // Remove os ciclos cujo toolId === id
    sheet = getDataSheet("Ciclos");
    const rows = sheet.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      const json = rows[i][1];
      if (!json) continue;
      try {
        const cycle = JSON.parse(json);
        if (cycle.toolId === id) sheet.deleteRow(i + 1);
      } catch (err) { /* linha inválida, ignora */ }
    }
    return jsonResponse({ status: "ok" });
  } finally {
    lock.releaseLock();
  }
}

function bulkImport(tools, cycles) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    replaceSheetData("Ferramentas", Array.isArray(tools) ? tools : []);
    replaceSheetData("Ciclos", Array.isArray(cycles) ? cycles : []);
    return jsonResponse({ status: "ok" });
  } finally {
    lock.releaseLock();
  }
}

function replaceSheetData(sheetName, items) {
  const sheet = getDataSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getMaxColumns()).clearContent();
  }
  const now = new Date().toISOString();
  const rows = items.filter(it => it && it.id).map(it => [it.id, JSON.stringify(it), now]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
}

// ── E-MAIL ────────────────────────────────────────────────────────────────────

function handleSendEmail(data) {
  const recipients = Array.isArray(data.recipients) ? data.recipients : [];
  const subject = data.subject || "(sem assunto)";
  const body = data.body || "";
  const htmlBody = data.htmlBody || "";

  if (!recipients.length) {
    return jsonResponse({ status: "error", message: "Nenhum destinatário informado." });
  }

  recipients.forEach(function (to) {
    if (htmlBody) {
      MailApp.sendEmail(to, subject, body, { htmlBody: htmlBody });
    } else {
      MailApp.sendEmail(to, subject, body);
    }
  });

  if (SHEET_ID && Array.isArray(data.record)) {
    appendToHistSheet(data.record);
  }

  return jsonResponse({ status: "ok", count: recipients.length });
}

function appendToHistSheet(record) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(HIST_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HIST_SHEET_NAME);
    sheet.appendRow(HIST_HEADERS);
  }
  sheet.appendRow(record);
}

// ── RESUMO SEMANAL ───────────────────────────────────────────────────────────
//
// Envia um e-mail listando todas as ferramentas que estão com "necessita
// ajuste" marcado e ainda não foram liberadas pela Ferramentaria.
//
// COMO ATIVAR (uma vez só):
// 1. Preencha RESUMO_SEMANAL_DESTINATARIOS no topo deste arquivo com os
//    e-mails que devem receber o resumo.
// 2. No editor do Apps Script, clique no ícone de relógio "Acionadores"
//    (menu lateral esquerdo).
// 3. Clique em "+ Adicionar acionador".
//    - Função a ser executada: enviarResumoSemanal
//    - Origem do evento: Baseado em tempo
//    - Tipo de acionador baseado em tempo: Timer semanal
//    - Selecione o dia da semana e o horário (ex: toda segunda, 08h)
// 4. Salvar. Autorize as permissões se solicitado.

function enviarResumoSemanal() {
  if (!SHEET_ID || !RESUMO_SEMANAL_DESTINATARIOS.length) return;

  const allTools = readRecords("Ferramentas");
  const allCycles = readRecords("Ciclos");

  const pendentes = [];
  allTools.forEach(function (tool) {
    const tc = allCycles.filter(function (c) { return c.toolId === tool.id; });
    if (!tc.length) return;
    const last = tc.slice().sort(function (a, b) {
      return new Date(b.createdAt) - new Date(a.createdAt);
    })[0];
    if (last.necessitaAjuste && !last.liberado) {
      pendentes.push({ tool: tool, cycle: last });
    }
  });

  if (!pendentes.length) return; // nada pendente, não envia e-mail

  const tz = Session.getScriptTimeZone();
  const linhas = pendentes.map(function (p) {
    const saida = p.cycle.saidaData ? Utilities.formatDate(new Date(p.cycle.saidaData), tz, "dd/MM/yyyy") : "—";
    return "<tr>" +
      "<td style='padding:8px;border:1px solid #ddd;'>" + (p.tool.numFerramenta || "") + "</td>" +
      "<td style='padding:8px;border:1px solid #ddd;'>" + (p.tool.descProduto || "") + "</td>" +
      "<td style='padding:8px;border:1px solid #ddd;'>" + (p.tool.cliente || "") + "</td>" +
      "<td style='padding:8px;border:1px solid #ddd;'>" + saida + "</td>" +
      "<td style='padding:8px;border:1px solid #ddd;'>" + (p.cycle.parecerQualidade || "") + "</td>" +
      "</tr>";
  }).join("");

  const subject = "🔧 Resumo Semanal — " + pendentes.length + " ferramenta(s) precisando de ajuste";
  const htmlBody = "<div style='font-family:Arial,sans-serif;'>" +
    "<h2 style='color:#b52a2a;'>🔧 Ferramentas precisando de ajuste</h2>" +
    "<p>" + pendentes.length + " ferramenta(s) aguardando ajuste na Ferramentaria.</p>" +
    "<table style='border-collapse:collapse;width:100%;'>" +
    "<tr style='background:#f3f4f6;'>" +
    "<th style='padding:8px;border:1px solid #ddd;text-align:left;'>Ferramenta</th>" +
    "<th style='padding:8px;border:1px solid #ddd;text-align:left;'>Produto</th>" +
    "<th style='padding:8px;border:1px solid #ddd;text-align:left;'>Cliente</th>" +
    "<th style='padding:8px;border:1px solid #ddd;text-align:left;'>Saída</th>" +
    "<th style='padding:8px;border:1px solid #ddd;text-align:left;'>Parecer Qualidade</th>" +
    "</tr>" + linhas + "</table></div>";

  RESUMO_SEMANAL_DESTINATARIOS.forEach(function (to) {
    MailApp.sendEmail({ to: to, subject: subject, htmlBody: htmlBody });
  });
}

// ── UTIL ──────────────────────────────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
