// ─────────────────────────────────────────────────────────────
// DEUNA NPS — Google Apps Script
// ─────────────────────────────────────────────────────────────
//
// SETUP (rodar uma vez após publicar):
//   1. Abra o editor do Apps Script → Execute → setupTriggers()
//   2. Autorize as permissões solicitadas
//   Isso cria:
//     • Trigger de edição → reconstrói dashboard ao editar NPS Answers
//     • Trigger diário → rebuild automático todo dia às 8h (backup)
// ─────────────────────────────────────────────────────────────

const SHEET_NAME     = 'NPS Answers';
const DASHBOARD_NAME = 'Dashboard';
const CHARTDATA_NAME = 'ChartData';

// Colunas da planilha de respostas (1-indexed)
const COLS = {
  data:        1,
  empresa:     2,
  nome:        3,
  email:       4,
  nps:         5,
  categoria:   6,
  idioma:      7,
  condResp:    8,
  suporte:     9,
  comtecVel:   10,
  comtecQual:  11,
  comtecProat: 12,
  comunicacao: 13,
  resultados:  14,
  integRapidez:15,
  integQual:   16,
  integFacil:  17,
  aspectos:      18,
  melhoria:      19,
  valorAgregado: 20,
};
const TOTAL_COLS = 20;

const HEADER = [
  'Data/Hora', 'Empresa', 'Nome', 'E-mail',
  'NPS (0–10)', 'Categoria NPS', 'Idioma', 'Resposta Condicional',
  'Suporte Técnico',
  'Com. Técnica — Velocidade', 'Com. Técnica — Qualidade', 'Com. Técnica — Proatividade',
  'Comunicação da Equipe',
  'Resultados (1–10)',
  'Integração — Rapidez', 'Integração — Qualidade', 'Integração — Facilidade',
  'Aspectos Valorizados', 'Sugestões de Melhoria',
  'Clareza Valor Agregado (1–10)',
];

// ── Recebe respostas do formulário ──────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    let sheet  = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADER);
      const h = sheet.getRange(1, 1, 1, TOTAL_COLS);
      h.setFontWeight('bold');
      h.setBackground('#1C1C1C');
      h.setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    const score     = Number(data.nps);
    const categoria = score <= 6 ? 'Detrator' : score <= 8 ? 'Neutro' : 'Promotor';

    sheet.appendRow([
      new Date(data.submitted_at),
      data.empresa       || '',
      data.nome          || '',
      data.email         || '',
      score,
      categoria,
      data.language      || 'pt',
      data.cond_resp     || '',
      data.suporte       || '',
      data.comtec_vel    || '',
      data.comtec_qual   || '',
      data.comtec_proat  || '',
      data.comunicacao   || '',
      Number(data.resultados)   || '',
      Number(data.integ_rapidez)|| '',
      Number(data.integ_qual)   || '',
      Number(data.integ_facil)  || '',
      (data.aspectos || []).join(', '),
      data.melhoria           || '',
      Number(data.valor_agregado) || '',
    ]);

    buildDashboard();

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Serve dados para dashboard externo (opcional) ───────────
function doGet(e) {
  buildDashboard();
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'dashboard updated' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Helpers ─────────────────────────────────────────────────
function avg(arr) {
  const nums = arr.filter(v => v > 0);
  return nums.length ? Math.round((nums.reduce((a, b) => a + b, 0) / nums.length) * 10) / 10 : 0;
}

function faceScore(arr, col) {
  // Converte carinha em número: good=10, neutral=5, bad=1
  const map = { good: 10, neutral: 5, bad: 1 };
  const nums = arr.map(r => map[r[col]] || 0).filter(v => v > 0);
  return nums.length ? Math.round((nums.reduce((a, b) => a + b, 0) / nums.length) * 10) / 10 : 0;
}

const LIKERT_MAP = {
  'Concordo totalmente': 0, 'Concordo': 1, 'Discordo': 2, 'Discordo totalmente': 3,
  'Strongly agree': 0, 'Agree': 1, 'Disagree': 2, 'Strongly disagree': 3,
  'Totalmente de acuerdo': 0, 'De acuerdo': 1, 'En desacuerdo': 2, 'Totalmente en desacuerdo': 3,
};
function likertCounts(data, col) {
  const r = [0, 0, 0, 0];
  data.forEach(row => {
    const idx = LIKERT_MAP[row[col]];
    if (idx !== undefined) r[idx]++;
  });
  return r;
}

// ── Constrói o Dashboard ─────────────────────────────────────
function buildDashboard() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const answers = ss.getSheetByName(SHEET_NAME);

  let cd = ss.getSheetByName(CHARTDATA_NAME);
  if (!cd) { cd = ss.insertSheet(CHARTDATA_NAME); }
  else { cd.clear(); }

  let dash = ss.getSheetByName(DASHBOARD_NAME);
  if (!dash) {
    dash = ss.insertSheet(DASHBOARD_NAME);
    ss.setActiveSheet(dash);
    ss.moveActiveSheet(1);
  } else {
    dash.clear();
    dash.getCharts().forEach(c => dash.removeChart(c));
  }

  // Lê os dados
  const lastRow = answers.getLastRow();
  const firstCell = lastRow >= 1 ? answers.getRange(1, 1).getValue() : '';
  const hasHeader = !(firstCell instanceof Date);
  const dataStartRow = hasHeader ? 2 : 1;
  const dataCount = lastRow - (hasHeader ? 1 : 0);

  if (dataCount < 1) {
    dash.getRange('B2').setValue('Nenhuma resposta ainda.');
    return;
  }

  const data = answers.getRange(dataStartRow, 1, dataCount, TOTAL_COLS).getValues();

  // Índices (0-based)
  const iNps         = COLS.nps - 1;         // 4
  const iSuporte     = COLS.suporte - 1;      // 8
  const iComtecVel   = COLS.comtecVel - 1;    // 9
  const iComtecQual  = COLS.comtecQual - 1;   // 10
  const iComtecProat = COLS.comtecProat - 1;  // 11
  const iComunicacao = COLS.comunicacao - 1;  // 12
  const iResultados  = COLS.resultados - 1;   // 13
  const iIntegRap    = COLS.integRapidez - 1; // 14
  const iIntegQual   = COLS.integQual - 1;    // 15
  const iIntegFacil  = COLS.integFacil - 1;   // 16
  const iAspectos      = COLS.aspectos - 1;       // 17
  const iValorAgregado = COLS.valorAgregado - 1;  // 19

  // Métricas NPS
  const total      = data.length;
  const scores     = data.map(r => Number(r[iNps]));
  const promotores = data.filter(r => Number(r[iNps]) >= 9).length;
  const neutros    = data.filter(r => Number(r[iNps]) >= 7 && Number(r[iNps]) <= 8).length;
  const detratores = data.filter(r => Number(r[iNps]) <= 6).length;
  const npsScore   = Math.round(((promotores - detratores) / total) * 100);

  // Distribuição 0–10
  const dist = Array(11).fill(0);
  scores.forEach(s => { if (s >= 0 && s <= 10) dist[s]++; });

  // Likert (suporte e comunicacao da equipe)
  const suporte     = likertCounts(data, iSuporte);
  const comunicacao = likertCounts(data, iComunicacao);

  // Médias das réguas
  const avgResultados  = avg(data.map(r => Number(r[iResultados])));
  const avgIntegRap    = avg(data.map(r => Number(r[iIntegRap])));
  const avgIntegQual   = avg(data.map(r => Number(r[iIntegQual])));
  const avgIntegFacil    = avg(data.map(r => Number(r[iIntegFacil])));
  const avgValorAgregado = avg(data.map(r => Number(r[iValorAgregado])));

  // Carinhas comunicação técnica (good=10, neutral=5, bad=1)
  const avgComtecVel   = faceScore(data, iComtecVel);
  const avgComtecQual  = faceScore(data, iComtecQual);
  const avgComtecProat = faceScore(data, iComtecProat);

  // Aspectos
  const aspectMap = {};
  data.forEach(r => {
    if (!r[iAspectos]) return;
    String(r[iAspectos]).split(', ').forEach(a => {
      if (a.trim()) aspectMap[a.trim()] = (aspectMap[a.trim()] || 0) + 1;
    });
  });
  const aspectos = Object.entries(aspectMap).sort((a, b) => b[1] - a[1]);

  // ── LAYOUT ──────────────────────────────────────────────────
  dash.setColumnWidth(1, 20);
  dash.setColumnWidth(2, 180);
  dash.setColumnWidth(3, 180);
  dash.setColumnWidth(4, 180);
  dash.setColumnWidth(5, 180);
  dash.setColumnWidth(6, 20);

  // Título
  dash.setRowHeight(1, 10);
  dash.setRowHeight(2, 50);
  const titleCell = dash.getRange('B2:E2');
  titleCell.merge();
  titleCell.setValue('NPS Dashboard — Deuna Partnerships');
  titleCell.setFontSize(18).setFontWeight('bold').setFontColor('#1C1C1C')
           .setFontFamily('Arial').setVerticalAlignment('middle');

  dash.setRowHeight(3, 8);
  dash.setRowHeight(4, 22);
  const updCell = dash.getRange('B4:E4');
  updCell.merge();
  updCell.setValue('Atualizado em: ' + new Date().toLocaleString('pt-BR'));
  updCell.setFontSize(10).setFontColor('#8E8E8E').setFontFamily('Arial');
  dash.setRowHeight(5, 12);

  // ── Cards NPS ──
  dash.setRowHeight(6, 24);
  dash.setRowHeight(7, 48);
  dash.setRowHeight(8, 28);
  dash.setRowHeight(9, 16);

  styleCard(dash, 6, 2, 'Total de Respostas', total, '#FFF3EE', '#FF5500');
  const npsColor = npsScore >= 50 ? '#0B9595' : npsScore >= 0 ? '#B45309' : '#FF614B';
  const npsBg    = npsScore >= 50 ? '#E6F7F7' : npsScore >= 0 ? '#FFF8EC' : '#FEF2F2';
  styleCard(dash, 6, 3, 'NPS Score', (npsScore > 0 ? '+' : '') + npsScore, npsBg, npsColor);
  styleCard(dash, 6, 4, '😊 Promotores (9–10)', promotores, '#E6F7F7', '#0B9595');
  styleCard(dash, 6, 5, '😟 Detratores (0–6)', detratores, '#FEF2F2', '#FF614B');

  // ── Cards de médias (réguas) ──
  dash.setRowHeight(10, 12);
  dash.setRowHeight(11, 20);
  dash.setRowHeight(12, 24);
  dash.setRowHeight(13, 48);
  dash.setRowHeight(14, 28);
  dash.setRowHeight(15, 16);

  const lblRow11 = dash.getRange('B11:E11');
  lblRow11.merge();
  lblRow11.setValue('Médias — Réguas (escala 1–10)');
  lblRow11.setFontSize(10).setFontWeight('bold').setFontColor('#8E8E8E').setFontFamily('Arial');

  // Row 1: Resultados, Valor Agregado, Integ Rapidez, Integ Qualidade
  styleCard(dash, 12, 2, 'Resultados da Parceria',  avgResultados,    '#F0F4FF', '#3B5BDB');
  styleCard(dash, 12, 3, 'Clareza Valor Agregado',  avgValorAgregado, '#F0F4FF', '#3B5BDB');
  styleCard(dash, 12, 4, 'Integração — Rapidez',    avgIntegRap,      '#F0F4FF', '#3B5BDB');
  styleCard(dash, 12, 5, 'Integração — Qualidade',  avgIntegQual,     '#F0F4FF', '#3B5BDB');

  // Row 2: Integ Facilidade (isolada)
  dash.setRowHeight(16, 12);
  dash.setRowHeight(17, 24);
  dash.setRowHeight(18, 48);
  dash.setRowHeight(19, 28);
  styleCard(dash, 17, 2, 'Integração — Facilidade', avgIntegFacil,    '#F0F4FF', '#3B5BDB');

  // ── Cards comunicação técnica (carinhas) ──
  dash.setRowHeight(20, 12);
  dash.setRowHeight(21, 20);
  dash.setRowHeight(22, 24);
  dash.setRowHeight(23, 48);
  dash.setRowHeight(24, 28);
  dash.setRowHeight(25, 16);

  const lblRow21 = dash.getRange('B21:E21');
  lblRow21.merge();
  lblRow21.setValue('Comunicação Técnica — Score médio (good=10 · neutral=5 · bad=1)');
  lblRow21.setFontSize(10).setFontWeight('bold').setFontColor('#8E8E8E').setFontFamily('Arial');

  styleCard(dash, 22, 2, '⚡ Velocidade',   avgComtecVel,   '#FFF8EC', '#B45309');
  styleCard(dash, 22, 3, '✅ Qualidade',    avgComtecQual,  '#FFF8EC', '#B45309');
  styleCard(dash, 22, 4, '📣 Proatividade', avgComtecProat, '#FFF8EC', '#B45309');

  // ── CHARTDATA ────────────────────────────────────────────────

  // A:B — Distribuição NPS
  cd.getRange(1, 1).setValue('Nota');
  cd.getRange(1, 2).setValue('Respostas');
  for (let i = 0; i <= 10; i++) {
    cd.getRange(2 + i, 1).setValue(i);
    cd.getRange(2 + i, 2).setValue(dist[i]);
  }

  // D:E — Segmentação
  cd.getRange(1, 4).setValue('Categoria');
  cd.getRange(1, 5).setValue('Total');
  cd.getRange(2, 4).setValue('Promotores (9–10)'); cd.getRange(2, 5).setValue(promotores);
  cd.getRange(3, 4).setValue('Neutros (7–8)');     cd.getRange(3, 5).setValue(neutros);
  cd.getRange(4, 4).setValue('Detratores (0–6)');  cd.getRange(4, 5).setValue(detratores);

  // G:K — Likert (suporte + comunicação da equipe)
  cd.getRange(1, 7).setValue('Critério');
  cd.getRange(1, 8).setValue('Concordo totalmente');
  cd.getRange(1, 9).setValue('Concordo');
  cd.getRange(1, 10).setValue('Discordo');
  cd.getRange(1, 11).setValue('Discordo totalmente');
  cd.getRange(2, 7).setValue('Suporte Técnico');
  cd.getRange(3, 7).setValue('Comunicação da Equipe');
  [suporte, comunicacao].forEach((arr, i) => {
    arr.forEach((v, j) => cd.getRange(2 + i, 8 + j).setValue(v));
  });

  // M:N — Aspectos valorizados
  cd.getRange(1, 13).setValue('Aspecto');
  cd.getRange(1, 14).setValue('Menções');
  aspectos.slice(0, 8).forEach((a, i) => {
    cd.getRange(2 + i, 13).setValue(a[0]);
    cd.getRange(2 + i, 14).setValue(a[1]);
  });

  // P:Q — Médias das réguas
  cd.getRange(1, 16).setValue('Dimensão');
  cd.getRange(1, 17).setValue('Média');
  const sliderDims = [
    ['Resultados', avgResultados],
    ['Clareza Valor Agregado', avgValorAgregado],
    ['Integração — Rapidez', avgIntegRap],
    ['Integração — Qualidade', avgIntegQual],
    ['Integração — Facilidade', avgIntegFacil],
  ];
  sliderDims.forEach((d, i) => {
    cd.getRange(2 + i, 16).setValue(d[0]);
    cd.getRange(2 + i, 17).setValue(d[1]);
  });

  SpreadsheetApp.flush();

  // ── GRÁFICOS ─────────────────────────────────────────────────

  // 1. Distribuição NPS (barras)
  dash.insertChart(dash.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(cd.getRange(1, 1, 12, 2))
    .setPosition(22, 2, 0, 0)
    .setOption('title', 'Distribuição de Notas NPS')
    .setOption('legend', { position: 'none' })
    .setOption('colors', ['#FF5500'])
    .setOption('hAxis', { title: 'Nota' })
    .setOption('vAxis', { title: 'Respostas', minValue: 0 })
    .setOption('fontName', 'Arial')
    .build());

  // 2. Segmentação (rosca)
  dash.insertChart(dash.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(cd.getRange(1, 4, 4, 2))
    .setPosition(22, 4, 0, 0)
    .setOption('title', 'Promotores · Neutros · Detratores')
    .setOption('pieHole', 0.5)
    .setOption('colors', ['#0B9595', '#FFB84D', '#FF614B'])
    .setOption('fontName', 'Arial')
    .build());

  // 3. Likert — Suporte e Comunicação (barras empilhadas)
  dash.insertChart(dash.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(cd.getRange(1, 7, 3, 5))
    .setPosition(42, 2, 0, 0)
    .setOption('title', 'Suporte e Comunicação da Equipe')
    .setOption('isStacked', true)
    .setOption('colors', ['#0B9595', '#76B4E8', '#FFB84D', '#FF614B'])
    .setOption('fontName', 'Arial')
    .build());

  // 4. Aspectos valorizados (barras horizontais)
  const nAsp = Math.min(aspectos.length, 8) + 1;
  dash.insertChart(dash.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(cd.getRange(1, 13, nAsp, 2))
    .setPosition(42, 4, 0, 0)
    .setOption('title', 'O que os Parceiros Mais Valorizam')
    .setOption('legend', { position: 'none' })
    .setOption('colors', ['#FF5500'])
    .setOption('fontName', 'Arial')
    .build());

  // 5. Médias das réguas (barras)
  dash.insertChart(dash.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(cd.getRange(1, 16, 6, 2))
    .setPosition(58, 2, 0, 0)
    .setOption('title', 'Médias — Avaliações por Régua (1–10)')
    .setOption('legend', { position: 'none' })
    .setOption('colors', ['#3B5BDB'])
    .setOption('vAxis', { minValue: 0, maxValue: 10 })
    .setOption('fontName', 'Arial')
    .build());

  cd.hideSheet();
  SpreadsheetApp.flush();
}

// ── Automação: triggers instaláveis ─────────────────────────

/**
 * Rodar UMA VEZ no editor do Apps Script.
 * Cria trigger de edição e trigger diário de rebuild.
 */
function setupTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remove triggers antigos para evitar duplicatas
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'onNpsAnswersEdit' || fn === 'dailyRebuild') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Trigger de edição na planilha
  ScriptApp.newTrigger('onNpsAnswersEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // Trigger diário às 8h (horário do projeto, America/Sao_Paulo)
  ScriptApp.newTrigger('dailyRebuild')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone('America/Sao_Paulo')
    .create();

  Logger.log('✅ Triggers criados: onNpsAnswersEdit + dailyRebuild (8h diário)');
}

/**
 * Disparado automaticamente ao editar qualquer célula.
 * Reconstrói o dashboard apenas se a aba editada for NPS Answers.
 */
function onNpsAnswersEdit(e) {
  if (!e || !e.source) return;
  const editedSheet = e.range.getSheet().getName();
  if (editedSheet === SHEET_NAME) {
    buildDashboard();
  }
}

/**
 * Rebuild diário às 8h — garante que o dashboard esteja sempre atualizado.
 */
function dailyRebuild() {
  buildDashboard();
}

// ── Helper: estiliza card ────────────────────────────────────
function styleCard(sheet, startRow, col, label, value, bg, color) {
  sheet.getRange(startRow, col)
    .setValue(label)
    .setBackground(bg).setFontColor(color)
    .setFontSize(9).setFontWeight('bold').setFontFamily('Arial')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, false, true, false, false, '#E8E8E8', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(startRow + 1, col)
    .setValue(value)
    .setBackground(bg).setFontColor(color)
    .setFontSize(28).setFontWeight('bold').setFontFamily('Arial')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(false, true, false, true, false, false, '#E8E8E8', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(startRow + 2, col)
    .setBackground(bg)
    .setBorder(false, true, true, true, false, false, '#E8E8E8', SpreadsheetApp.BorderStyle.SOLID);
}
