// ─────────────────────────────────────────────────────────────
// DEUNA NPS — Google Apps Script
// ─────────────────────────────────────────────────────────────

const SHEET_NAME      = 'NPS Answers';
const DASHBOARD_NAME  = 'Dashboard';
const CHARTDATA_NAME  = 'ChartData';

// ── Recebe respostas do formulário ──────────────────────────
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
      const h = sheet.getRange(1,1,1,11);
      h.setFontWeight('bold');
      h.setBackground('#1C1C1C');
      h.setFontColor('#FFFFFF');
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

    // Atualiza o dashboard automaticamente a cada nova resposta
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

// ── Constrói o Dashboard ─────────────────────────────────────
function buildDashboard() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const answers  = ss.getSheetByName(SHEET_NAME);

  // Cria ou limpa a aba ChartData (dados dos gráficos)
  let cd = ss.getSheetByName(CHARTDATA_NAME);
  if (!cd) {
    cd = ss.insertSheet(CHARTDATA_NAME);
    ss.setActiveSheet(cd);
  } else {
    cd.clear();
  }

  // Cria ou limpa a aba Dashboard
  let dash = ss.getSheetByName(DASHBOARD_NAME);
  if (!dash) {
    dash = ss.insertSheet(DASHBOARD_NAME);
    ss.setActiveSheet(dash);
    ss.moveActiveSheet(1); // move para primeira posição
  } else {
    dash.clear();
    // Remove gráficos existentes
    dash.getCharts().forEach(c => dash.removeChart(c));
  }

  // Lê os dados
  const lastRow = answers.getLastRow();

  // Detecta se linha 1 é cabeçalho (texto) ou dado (número na col 3)
  const firstCellCol3 = lastRow >= 1 ? answers.getRange(1, 3).getValue() : '';
  const hasHeader = isNaN(Number(firstCellCol3)) || firstCellCol3 === '';
  const answersDataStartRow = hasHeader ? 2 : 1;
  const dataCount = lastRow - (hasHeader ? 1 : 0);

  if (dataCount < 1) {
    dash.getRange('B2').setValue('Nenhuma resposta ainda.');
    return;
  }

  const data = answers.getRange(answersDataStartRow, 1, dataCount, 11).getValues();

  // Calcula métricas
  const total      = data.length;
  const scores     = data.map(r => Number(r[2]));
  const promotores = data.filter(r => Number(r[2]) >= 9).length;
  const neutros    = data.filter(r => Number(r[2]) >= 7 && Number(r[2]) <= 8).length;
  const detratores = data.filter(r => Number(r[2]) <= 6).length;
  const npsScore   = Math.round(((promotores - detratores) / total) * 100);

  // Distribuição 0–10
  const dist = Array(11).fill(0);
  scores.forEach(s => { if (s >= 0 && s <= 10) dist[s]++; });

  // Likert
  const likertMap = {
    'Concordo totalmente':0,'Concordo':1,'Discordo':2,'Discordo totalmente':3,
    'Strongly agree':0,'Agree':1,'Disagree':2,'Strongly disagree':3,
    'Totalmente de acuerdo':0,'De acuerdo':1,'En desacuerdo':2,'Totalmente en desacuerdo':3,
  };
  function likertCounts(col) {
    const r = [0,0,0,0];
    data.forEach(row => { const idx = likertMap[row[col]]; if(idx !== undefined) r[idx]++; });
    return r;
  }
  const suporte      = likertCounts(6);
  const comunicacao  = likertCounts(7);
  const resultados   = likertCounts(8);

  // Aspectos
  const aspectMap = {};
  data.forEach(r => {
    if (!r[9]) return;
    String(r[9]).split(', ').forEach(a => {
      if (a.trim()) aspectMap[a.trim()] = (aspectMap[a.trim()] || 0) + 1;
    });
  });
  const aspectos = Object.entries(aspectMap).sort((a,b) => b[1]-a[1]);

  // ── LAYOUT ──────────────────────────────────────────────
  dash.setColumnWidth(1, 20);   // margem esquerda
  dash.setColumnWidth(2, 180);
  dash.setColumnWidth(3, 180);
  dash.setColumnWidth(4, 180);
  dash.setColumnWidth(5, 180);
  dash.setColumnWidth(6, 20);   // margem direita

  // Título
  dash.setRowHeight(1, 10);
  dash.setRowHeight(2, 50);
  const titleCell = dash.getRange('B2:E2');
  titleCell.merge();
  titleCell.setValue('NPS Dashboard — deuna Partnerships');
  titleCell.setFontSize(18).setFontWeight('bold').setFontColor('#1C1C1C')
           .setFontFamily('Arial').setVerticalAlignment('middle');

  dash.setRowHeight(3, 8);

  // Última atualização
  dash.setRowHeight(4, 22);
  const updCell = dash.getRange('B4:E4');
  updCell.merge();
  updCell.setValue('Atualizado em: ' + new Date().toLocaleString('pt-BR'));
  updCell.setFontSize(10).setFontColor('#8E8E8E').setFontFamily('Arial');

  dash.setRowHeight(5, 12);

  // ── CARDS de métricas ──
  dash.setRowHeight(6, 24);
  dash.setRowHeight(7, 48);
  dash.setRowHeight(8, 28);
  dash.setRowHeight(9, 16);

  function writeCard(range3rows, label, value, bgColor, fontColor) {
    const cells  = dash.getRange(range3rows);
    const rLabel = range3rows.split(':')[0];
    const [col, startRow] = [rLabel.replace(/\d/g,''), parseInt(rLabel.replace(/\D/g,''))];

    dash.getRange(`${col}${startRow}:${rLabel.replace(/\d/g,'')}${startRow}`).merge()
        .setValue(label)
        .setFontSize(9).setFontWeight('bold').setFontColor(fontColor)
        .setFontFamily('Arial').setBackground(bgColor)
        .setHorizontalAlignment('center').setVerticalAlignment('middle');

    dash.getRange(`${col}${startRow+1}:${col}${startRow+1}`)
        .setValue(value)
        .setFontSize(28).setFontWeight('bold').setFontColor(fontColor)
        .setFontFamily('Arial').setBackground(bgColor)
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  // Card: Total
  styleCard(dash, 6, 2, 3, 'Total de Respostas', total, '#FFF3EE', '#FF5500');
  // Card: NPS
  const npsColor = npsScore >= 50 ? '#0B9595' : npsScore >= 0 ? '#B45309' : '#FF614B';
  const npsBg    = npsScore >= 50 ? '#E6F7F7' : npsScore >= 0 ? '#FFF8EC' : '#FEF2F2';
  styleCard(dash, 6, 3, 3, 'NPS Score', (npsScore > 0 ? '+' : '') + npsScore, npsBg, npsColor);
  // Card: Promotores
  styleCard(dash, 6, 4, 3, '😊 Promotores (9–10)', promotores, '#E6F7F7', '#0B9595');
  // Card: Detratores
  styleCard(dash, 6, 5, 3, '😟 Detratores (0–6)', detratores, '#FEF2F2', '#FF614B');

  dash.setRowHeight(9, 20);

  // ── TABELA DE DADOS PARA OS GRÁFICOS (aba ChartData) ────────
  // Distribuição NPS (col A:B)
  cd.getRange(1, 1).setValue('Nota');
  cd.getRange(1, 2).setValue('Respostas');
  for (let i = 0; i <= 10; i++) {
    cd.getRange(2 + i, 1).setValue(i);
    cd.getRange(2 + i, 2).setValue(dist[i]);
  }

  // Segmentação (col D:E)
  cd.getRange(1, 4).setValue('Categoria');
  cd.getRange(1, 5).setValue('Total');
  cd.getRange(2, 4).setValue('Promotores (9–10)');
  cd.getRange(2, 5).setValue(promotores);
  cd.getRange(3, 4).setValue('Neutros (7–8)');
  cd.getRange(3, 5).setValue(neutros);
  cd.getRange(4, 4).setValue('Detratores (0–6)');
  cd.getRange(4, 5).setValue(detratores);

  // Likert (col G:K)
  cd.getRange(1, 7).setValue('Critério');
  cd.getRange(1, 8).setValue('Concordo totalmente');
  cd.getRange(1, 9).setValue('Concordo');
  cd.getRange(1, 10).setValue('Discordo');
  cd.getRange(1, 11).setValue('Discordo totalmente');
  cd.getRange(2, 7).setValue('Suporte Técnico');
  cd.getRange(3, 7).setValue('Comunicação');
  cd.getRange(4, 7).setValue('Resultados');
  [suporte, comunicacao, resultados].forEach((arr, i) => {
    arr.forEach((v, j) => cd.getRange(2 + i, 8 + j).setValue(v));
  });

  // Aspectos (col M:N)
  cd.getRange(1, 13).setValue('Aspecto');
  cd.getRange(1, 14).setValue('Menções');
  aspectos.slice(0, 8).forEach((a, i) => {
    cd.getRange(2 + i, 13).setValue(a[0]);
    cd.getRange(2 + i, 14).setValue(a[1]);
  });

  // Confirma todos os dados antes de criar os gráficos
  SpreadsheetApp.flush();

  // ── GRÁFICOS ─────────────────────────────────────────────

  // 1. Gráfico distribuição NPS (barras)
  const distData = cd.getRange(1, 1, 12, 2);
  const chartDist = dash.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(distData)
    .setPosition(10, 2, 0, 0)
    .setOption('title', 'Distribuição de Notas NPS')
    .setOption('legend', { position: 'none' })
    .setOption('colors', ['#FF5500'])
    .setOption('hAxis', { title: 'Nota' })
    .setOption('vAxis', { title: 'Respostas', minValue: 0 })
    .setOption('fontName', 'Arial')
    .build();
  dash.insertChart(chartDist);

  // 2. Gráfico segmentação (rosca)
  const segData = cd.getRange(1, 4, 4, 2);
  const chartSeg = dash.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(segData)
    .setPosition(10, 4, 0, 0)
    .setOption('title', 'Promotores · Neutros · Detratores')
    .setOption('pieHole', 0.5)
    .setOption('colors', ['#0B9595','#FFB84D','#FF614B'])
    .setOption('fontName', 'Arial')
    .build();
  dash.insertChart(chartSeg);

  // 3. Likert comparativo (barras empilhadas)
  const likertData = cd.getRange(1, 7, 4, 5);
  const chartLikert = dash.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(likertData)
    .setPosition(26, 2, 0, 0)
    .setOption('title', 'Avaliações por Critério')
    .setOption('isStacked', true)
    .setOption('colors', ['#0B9595','#76B4E8','#FFB84D','#FF614B'])
    .setOption('fontName', 'Arial')
    .build();
  dash.insertChart(chartLikert);

  // 4. Aspectos mais valorizados (barras horizontais)
  const nAsp = Math.min(aspectos.length, 8) + 1;
  const aspData = cd.getRange(1, 13, nAsp, 2);
  const chartAsp = dash.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(aspData)
    .setPosition(26, 4, 0, 0)
    .setOption('title', 'O que os Parceiros Mais Valorizam')
    .setOption('legend', { position: 'none' })
    .setOption('colors', ['#FF5500'])
    .setOption('fontName', 'Arial')
    .build();
  dash.insertChart(chartAsp);

  // Oculta a aba ChartData
  cd.hideSheet();

  SpreadsheetApp.flush();
}

// ── Helper: estiliza card ────────────────────────────────────
function styleCard(sheet, startRow, col, rowSpan, label, value, bg, color) {
  // Label
  const labelRange = sheet.getRange(startRow, col);
  labelRange.setValue(label)
    .setBackground(bg).setFontColor(color)
    .setFontSize(9).setFontWeight('bold').setFontFamily('Arial')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true,true,false,true,false,false,'#E8E8E8', SpreadsheetApp.BorderStyle.SOLID);

  // Value
  const valueRange = sheet.getRange(startRow+1, col);
  valueRange.setValue(value)
    .setBackground(bg).setFontColor(color)
    .setFontSize(28).setFontWeight('bold').setFontFamily('Arial')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(false,true,false,true,false,false,'#E8E8E8', SpreadsheetApp.BorderStyle.SOLID);

  // Bottom border
  const bottomRange = sheet.getRange(startRow+2, col);
  bottomRange.setBackground(bg)
    .setBorder(false,true,true,true,false,false,'#E8E8E8', SpreadsheetApp.BorderStyle.SOLID);
}
