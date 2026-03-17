// ─────────────────────────────────────────────────────────────
// DEUNA NPS — Google Apps Script
// Cole este código em script.google.com e faça o deploy como Web App
// ─────────────────────────────────────────────────────────────

const SHEET_NAME = 'Respostas NPS'; // nome da aba no Sheets

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(SHEET_NAME);

    // Cria a aba e os cabeçalhos se não existirem
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Data/Hora',
        'Empresa',
        'NPS (0–10)',
        'Categoria NPS',
        'Resposta Condicional',
        'Suporte Técnico',
        'Comunicação',
        'Resultados da Parceria',
        'Aspectos Valorizados',
        'Sugestões de Melhoria',
      ]);

      // Formata cabeçalhos
      const header = sheet.getRange(1, 1, 1, 10);
      header.setFontWeight('bold');
      header.setBackground('#1C1C1C');
      header.setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    // Categoriza o NPS
    const score = Number(data.nps);
    const categoria = score <= 6 ? 'Detrator' : score <= 8 ? 'Neutro' : 'Promotor';

    // Adiciona a linha de resposta
    sheet.appendRow([
      new Date(data.submitted_at),
      data.empresa       || '',
      score,
      categoria,
      data.cond_resp     || '',
      data.suporte       || '',
      data.comunicacao   || '',
      data.resultados    || '',
      (data.aspectos || []).join(', '),
      data.melhoria      || '',
    ]);

    // Retorna sucesso
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Necessário para permitir requisições cross-origin (CORS)
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'online' }))
    .setMimeType(ContentService.MimeType.JSON);
}
