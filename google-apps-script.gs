/**
 * Google Apps Script - Web App para envio de e-mails e registro em planilha
 * do "Histórico Ferramental".
 *
 * COMO IMPLANTAR:
 * 1. Acesse https://script.google.com e crie um novo projeto (pode ser com a
 *    mesma conta Gmail/Workspace que deve aparecer como remetente).
 * 2. Apague o conteúdo padrão de "Code.gs" e cole todo o código abaixo.
 * 3. (Opcional, recomendado) Troque o valor de TOKEN por uma senha/segredo
 *    qualquer, ex: "promf-2026-xyz". Esse token evita que outras pessoas usem
 *    sua URL para enviar e-mails em seu nome.
 * 4. (Opcional) Para registrar cada saída em uma planilha:
 *    - Crie uma planilha nova no Google Sheets
 *    - Copie o ID dela (parte da URL entre /d/ e /edit)
 *    - Cole o ID na constante SHEET_ID abaixo
 *    - Não precisa criar abas/cabeçalhos: o script cria sozinho na primeira saída
 * 5. Clique em "Implantar" > "Nova implantação".
 *    - Tipo: "App da Web" (Web app)
 *    - Executar como: "Eu" (sua conta)
 *    - Quem pode acessar: "Qualquer pessoa"
 * 6. Autorize as permissões solicitadas (envio de e-mail e acesso ao Drive/Sheets
 *    em seu nome).
 * 7. Copie a URL gerada (termina em /exec) e cole no campo
 *    "URL do Web App (Google Apps Script)" nas configurações de e-mail do
 *    Histórico Ferramental. Cole o mesmo TOKEN no campo "Token de segurança".
 *
 * Limite gratuito: contas Gmail comuns podem enviar ~100 e-mails/dia;
 * contas Google Workspace, ~1500 e-mails/dia.
 */

const TOKEN = ""; // defina um token secreto aqui (mesmo valor configurado no app)
const SHEET_ID = ""; // ID da planilha Google Sheets (deixe vazio para não registrar)
const SHEET_NAME = "Histórico";

const SHEET_HEADERS = [
  "Data/Hora Saída", "Nº Ferramenta", "Código PROMETAL", "C.C.", "Cliente",
  "Descrição da Operação", "Nº OP.", "Descrição do Produto", "Cód. Prod. Cliente",
  "Ciclo Nº", "Entrada", "Preparador Entrada", "Máquina",
  "Preparador Saída", "Peças Produzidas", "Peças Acumuladas",
  "Parecer Produção", "Data Parecer Produção", "Responsável Produção",
  "Parecer Qualidade", "Data Parecer Qualidade", "Responsável Qualidade",
  "Assinatura Produção", "Data Assinatura Produção", "Local Assinatura Produção",
  "Assinatura Qualidade", "Data Assinatura Qualidade", "Local Assinatura Qualidade"
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (TOKEN && data.token !== TOKEN) {
      return jsonResponse({ status: "error", message: "Token inválido." });
    }

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
      appendToSheet(data.record);
    }

    return jsonResponse({ status: "sent", count: recipients.length });
  } catch (err) {
    return jsonResponse({ status: "error", message: String(err) });
  }
}

function appendToSheet(record) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(SHEET_HEADERS);
  }
  sheet.appendRow(record);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
