// ─────────────────────────────────────────────────────────────
// DEUNA NPS — Google Apps Script
// Recebe respostas do formulário (doPost) e
// serve dados para o dashboard (doGet)
// ─────────────────────────────────────────────────────────────

const SHEET_NAME = 'Respostas NPS';

// ── Recebe respostas do formulário ──
function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Data/Hora','Empresa','NPS (0–10)','Categoria NPS',
        'Idioma','Resposta Condicional','Suporte Técnico',
        'Comunicação','Resultados da Parceria',
        'Aspectos Valorizados','Sugestões de Melhoria',
      ]);
      const header = sheet.getRange(1, 1, 1, 11);
      header.setFontWeight('bold');
      header.setBackground('#1C1C1C');
      header.setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    const score     = Number(data.nps);
    const categoria = score <= 6 ? 'Detrator' : score <= 8 ? 'Neutro' : 'Promotor';

    sheet.appendRow([
      new Date(data.submitted_at),
      data.empresa      || '',
      score,
      categoria,
      data.language     || 'pt',
      data.cond_resp    || '',
      data.suporte      || '',
      data.comunicacao  || '',
      data.resultados   || '',
      (data.aspectos || []).join(', '),
      data.melhoria     || '',
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Serve dados para o dashboard ──
function doGet(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet || sheet.getLastRow() <= 1) {
      return jsonResponse({ rows: [], summary: emptySummary() });
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows    = data.slice(1).map(row => ({
      date:        row[0],
      empresa:     row[1],
      nps:         Number(row[2]),
      categoria:   row[3],
      language:    row[4],
      cond_resp:   row[5],
      suporte:     row[6],
      comunicacao: row[7],
      resultados:  row[8],
      aspectos:    row[9],
      melhoria:    row[10],
    }));

    const npsScores   = rows.map(r => r.nps);
    const promotores  = rows.filter(r => r.nps >= 9).length;
    const neutros     = rows.filter(r => r.nps >= 7 && r.nps <= 8).length;
    const detratores  = rows.filter(r => r.nps <= 6).length;
    const total       = rows.length;
    const npsScore    = total > 0
      ? Math.round(((promotores - detratores) / total) * 100)
      : 0;

    // Distribuição 0-10
    const dist = Array(11).fill(0);
    npsScores.forEach(s => { if (s >= 0 && s <= 10) dist[s]++; });

    // Likert counts
    function likertCount(field) {
      const map = { 'Concordo totalmente':0,'Concordo':0,'Discordo':0,'Discordo totalmente':0,
                    'Totalmente de acuerdo':0,'De acuerdo':0,'En desacuerdo':0,'Totalmente en desacuerdo':0,
                    'Strongly agree':0,'Agree':0,'Disagree':0,'Strongly disagree':0 };
      // normalize to PT
      const norm = { 'Totalmente de acuerdo':'Concordo totalmente','De acuerdo':'Concordo',
                     'En desacuerdo':'Discordo','Totalmente en desacuerdo':'Discordo totalmente',
                     'Strongly agree':'Concordo totalmente','Agree':'Concordo',
                     'Disagree':'Discordo','Strongly disagree':'Discordo totalmente' };
      const result = { 'Concordo totalmente':0,'Concordo':0,'Discordo':0,'Discordo totalmente':0 };
      rows.forEach(r => {
        const val = norm[r[field]] || r[field];
        if (result[val] !== undefined) result[val]++;
      });
      return result;
    }

    // Aspectos mais citados
    const aspectMap = {};
    rows.forEach(r => {
      if (!r.aspectos) return;
      r.aspectos.split(', ').forEach(a => {
        if (a.trim()) aspectMap[a.trim()] = (aspectMap[a.trim()] || 0) + 1;
      });
    });
    const aspectos = Object.entries(aspectMap)
      .sort((a, b) => b[1] - a[1])
      .map(([label, count]) => ({ label, count }));

    // Respostas por mês
    const byMonth = {};
    rows.forEach(r => {
      const d = new Date(r.date);
      const key = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
      byMonth[key] = (byMonth[key] || 0) + 1;
    });
    const timeline = Object.entries(byMonth)
      .sort()
      .map(([month, count]) => ({ month, count }));

    const summary = {
      total, npsScore, promotores, neutros, detratores,
      dist,
      suporte:     likertCount('suporte'),
      comunicacao: likertCount('comunicacao'),
      resultados:  likertCount('resultados'),
      aspectos,
      timeline,
    };

    return jsonResponse({ rows, summary });

  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function emptySummary() {
  return {
    total:0, npsScore:0, promotores:0, neutros:0, detratores:0,
    dist: Array(11).fill(0),
    suporte:{}, comunicacao:{}, resultados:{},
    aspectos:[], timeline:[],
  };
}
