/**
 * Apps Script vinculado à planilha do Diagnóstico Financeiro (trilha personalizada).
 * Cópia local só de referência — o script publicado vive no editor do Google,
 * vinculado à planilha SPREADSHEET_ID abaixo.
 *
 * DEPLOY (atualizar o que já existe, não criar um novo):
 * 1. Abra a planilha (ID abaixo) → Extensões → Apps Script.
 * 2. Apague o conteúdo de Code.gs e cole este arquivo inteiro.
 * 3. Deploy → Gerenciar implantações → ícone de lápis na implantação ativa
 *    → Versão: "Nova versão" → Implantar.
 *    (Mantém a MESMA URL que o diagnostico-travessia.html já usa —
 *    não precisa mexer no SCRIPT_URL do HTML.)
 *
 * UTM: a página captura utm_source/medium/campaign/content/term da URL
 * (query string) e manda junto no POST. O cabeçalho se auto-corrige a cada
 * envio, então não precisa rodar setupHeaders() de novo numa planilha já
 * existente sem essas colunas.
 */

const SPREADSHEET_ID = '1S8otdQN4Wx_LevYwPyH1SYTY9o1P_Gn2VKkbHQGfBsU';
const SHEET_NAME = 'Respostas';

// Dispara pro n8n toda resposta da trilha, pra cair no Kommo na hora
// (em vez de esperar o motor de polling de 10 em 10 min).
const WEBHOOK_URL = 'https://n8nwebhook.leticiakratka.shop/webhook/travessia-form-trilha';
const WEBHOOK_ENABLED = true;

const HEADERS = [
  'Timestamp', 'Nome', 'Email', 'WhatsApp',
  'É PJ/Autônoma?', 'Tipo de PJ', 'Situação das Dívidas', 'Dinheiro Guardado',
  'Faixa de Renda', 'Profissão',
  'Motivação', 'Dificuldades', 'Expectativa', 'Objetivos 12 meses', 'Compartilhou mais',
  'Interesse em Consultoria',
  'Trilha Recomendada', 'Duplicado?',
  'UTM Source', 'UTM Medium', 'UTM Campaign', 'UTM Content', 'UTM Term'
];

function setupHeaders() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME); }
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  Logger.log('Headers configurados!');
}

function isDuplicate(sheet, email) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const emails = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  return emails.some(function(row) {
    return row[0].toString().toLowerCase() === email.toLowerCase();
  });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    // Reescreve o cabeçalho sempre — idempotente, e "cura" sozinho uma
    // aba antiga que ainda não tinha as colunas de UTM.
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');

    const timestamp   = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    const nome        = data.nome        || '';
    const email       = data.email       || '';
    const whatsapp    = data.whatsapp    || '';
    const pj          = data.pj          || '';
    const pj_tipo     = data.pj_tipo     || '';
    const divida      = data.divida      || '';
    const reserva     = data.reserva     || '';
    const renda       = data.renda       || '';
    const profissao   = data.profissao   || '';
    const motivacao    = data.motivacao    || '';
    const dificuldades = data.dificuldades || '';
    const expectativa  = data.expectativa  || '';
    const objetivos     = data.objetivos     || '';
    const compartilhar  = data.compartilhar  || '';
    const consultoria = data.consultoria || '';
    const trilha      = data.trilha      || '';
    const utm_source   = data.utm_source   || '';
    const utm_medium   = data.utm_medium   || '';
    const utm_campaign = data.utm_campaign || '';
    const utm_content  = data.utm_content  || '';
    const utm_term      = data.utm_term     || '';

    const duplicado = isDuplicate(sheet, email) ? 'SIM' : '';

    const row = [
      timestamp, nome, email, whatsapp, pj, pj_tipo, divida, reserva, renda, profissao,
      motivacao, dificuldades, expectativa, objetivos, compartilhar,
      consultoria, trilha, duplicado,
      utm_source, utm_medium, utm_campaign, utm_content, utm_term
    ];
    sheet.appendRow(row);

    if (duplicado === 'SIM') {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 18).setBackground('#FFD700').setFontWeight('bold');
    }

    if (WEBHOOK_ENABLED) {
      dispararWebhook({
        nome: nome, email: email, whatsapp: whatsapp, trilha: trilha,
        pj: pj, pj_tipo: pj_tipo, divida: divida, reserva: reserva,
        renda: renda, profissao: profissao,
        motivacao: motivacao, dificuldades: dificuldades, expectativa: expectativa,
        objetivos: objetivos, compartilhar: compartilhar,
        consultoria: consultoria,
        utm_source: utm_source, utm_medium: utm_medium, utm_campaign: utm_campaign,
        utm_content: utm_content, utm_term: utm_term
      });
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function dispararWebhook(dados) {
  try {
    const response = UrlFetchApp.fetch(WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(dados),
      muteHttpExceptions: true
    });
    const code = response.getResponseCode();
    Logger.log('Webhook trilha respondeu ' + code + ': ' + response.getContentText());
  } catch (err) {
    Logger.log('❌ Erro ao chamar webhook da trilha: ' + err.toString());
  }
}

function doGet(e) {
  return ContentService.createTextOutput('OK');
}
