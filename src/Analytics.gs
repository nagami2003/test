// ===== 経営分析 =====
// 公開関数: buildAnalyticsSheet, insertCharts, createMonthlyReportDoc, refreshAnalytics

// ============================================================
// 1. buildAnalyticsSheet — KPI・集計シートの新規作成/再構築
// ============================================================
function buildAnalyticsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kpi = collectKpiData_(ss);

  // シート取得または作成
  let sheet = ss.getSheetByName(CONFIG.SHEETS.ANALYTICS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.ANALYTICS);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
    // 既存チャートを削除
    sheet.getCharts().forEach(c => sheet.removeChart(c));
  }

  let row = 1;

  // ---- タイトル行 ----
  row = writeTitle_(sheet, row, '📊 経営分析レポート　' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日 更新'));

  // ---- セクション1: 支援概況 ----
  row = writeSectionHeader_(sheet, row, '【1】支援概況');
  const section1 = [
    ['支援中外国人数',     kpi.activeWorkers,      '名',  null],
    ['受入企業数',         kpi.companyCount,        '社',  null],
    ['1社あたり平均支援人数', kpi.avgPerCompany,    '名',  '0.0'],
    ['在留期限90日以内',   kpi.visaAlertCount,      '名 ⚠️', null],
  ];
  row = writeKpiTable_(sheet, row, section1);

  // ---- セクション2: 売上状況 ----
  row = writeSectionHeader_(sheet, row, '【2】売上状況');
  const section2 = [
    ['当月売上（請求額）',   kpi.thisMonthSales,     '',  '¥#,##0'],
    ['前月売上（請求額）',   kpi.lastMonthSales,     '',  '¥#,##0'],
    ['前年同月売上',         kpi.lastYearSales,      '',  '¥#,##0'],
    ['直近12ヶ月 売上累計', kpi.sales12Total,       '',  '¥#,##0'],
    ['当月 入金額',          kpi.thisMonthPaid,      '',  '¥#,##0'],
    ['入金率（直近12ヶ月）', kpi.collectionRate,     '',  '0.0%'],
    ['未回収額（未払い合計）', kpi.uncollected,      '',  '¥#,##0'],
  ];
  row = writeKpiTable_(sheet, row, section2);

  // ---- セクション3: 国籍別内訳 ----
  row = writeSectionHeader_(sheet, row, '【3】国籍別 支援人数内訳');
  row = writeBreakdownTable_(sheet, row, kpi.nationalityBreakdown, '国籍', '人数（名）');

  // ---- セクション4: 分野別内訳 ----
  row = writeSectionHeader_(sheet, row, '【4】分野別 支援人数内訳');
  row = writeBreakdownTable_(sheet, row, kpi.sectorBreakdown, '分野', '人数（名）');

  // ---- セクション5: 直近12ヶ月 月次売上推移 ----
  row = writeSectionHeader_(sheet, row, '【5】直近12ヶ月 月次売上推移');
  row = writeMonthlySalesTable_(sheet, row, kpi.monthlySales);

  // 列幅調整
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 80);
  sheet.setFrozenRows(1);

  SpreadsheetApp.flush();
  return sheet;
}

// ============================================================
// 2. insertCharts — 経営分析シートへグラフを挿入
// ============================================================
function insertCharts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.ANALYTICS);
  if (!sheet) {
    sheet = buildAnalyticsSheet();
  }

  // 既存チャートを全削除（重複防止）
  sheet.getCharts().forEach(c => sheet.removeChart(c));

  const kpi = collectKpiData_(ss);

  // ---- チャート用データ範囲を特定 ----
  // 月次売上テーブルのヘッダー行を探す
  const salesTableRange = findTableRange_(sheet, '月');
  // 国籍テーブル
  const nationalityTableRange = findTableRange_(sheet, '国籍');
  // 分野テーブル
  const sectorTableRange = findTableRange_(sheet, '分野');

  let chartAnchorRow = 1;
  let chartAnchorCol = 6; // F列以降に配置

  // ---- グラフ1: 月次売上 棒グラフ ----
  if (salesTableRange && salesTableRange.numRows > 1) {
    const monthCol   = sheet.getRange(salesTableRange.startRow, salesTableRange.startCol,
                                      salesTableRange.numRows, 1);
    const salesCol   = sheet.getRange(salesTableRange.startRow, salesTableRange.startCol + 1,
                                      salesTableRange.numRows, 1);
    const paidCol    = sheet.getRange(salesTableRange.startRow, salesTableRange.startCol + 2,
                                      salesTableRange.numRows, 1);

    const chart1 = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(monthCol)
      .addRange(salesCol)
      .addRange(paidCol)
      .setNumHeaders(1)
      .setOption('title', '直近12ヶ月 月次売上推移')
      .setOption('hAxis', { title: '月' })
      .setOption('vAxis', { title: '金額（円）', format: '¥#,##0' })
      .setOption('colors', [CONFIG.COLORS.HEADER, CONFIG.COLORS.ALERT_GREEN])
      .setOption('legend', { position: 'bottom' })
      .setOption('width', 500)
      .setOption('height', 300)
      .setPosition(chartAnchorRow, chartAnchorCol, 0, 0)
      .build();
    sheet.insertChart(chart1);
    chartAnchorRow += 17; // 次のチャートは下に配置
  }

  // ---- グラフ2: 国籍別 円グラフ ----
  if (nationalityTableRange && nationalityTableRange.numRows > 1 && kpi.activeWorkers > 0) {
    const natRange = sheet.getRange(
      nationalityTableRange.startRow, nationalityTableRange.startCol,
      nationalityTableRange.numRows, 2
    );
    const chart2 = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(natRange)
      .setNumHeaders(1)
      .setOption('title', '国籍別 支援人数')
      .setOption('pieHole', 0.4)
      .setOption('width', 400)
      .setOption('height', 300)
      .setPosition(chartAnchorRow, chartAnchorCol, 0, 0)
      .build();
    sheet.insertChart(chart2);
    chartAnchorRow += 17;
  }

  // ---- グラフ3: 分野別 円グラフ ----
  if (sectorTableRange && sectorTableRange.numRows > 1 && kpi.activeWorkers > 0) {
    const secRange = sheet.getRange(
      sectorTableRange.startRow, sectorTableRange.startCol,
      sectorTableRange.numRows, 2
    );
    const chart3 = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(secRange)
      .setNumHeaders(1)
      .setOption('title', '分野別 支援人数')
      .setOption('pieHole', 0.4)
      .setOption('width', 400)
      .setOption('height', 300)
      .setPosition(chartAnchorRow, chartAnchorCol, 0, 0)
      .build();
    sheet.insertChart(chart3);
  }

  SpreadsheetApp.flush();
}

// ============================================================
// 3. createMonthlyReportDoc — 月次経営レポートを Google Doc で生成
// ============================================================
function createMonthlyReportDoc() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kpi = collectKpiData_(ss);
  const today = new Date();
  const yearMonth = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy年MM月');
  const docTitle = `月次経営レポート_${Utilities.formatDate(today, 'Asia/Tokyo', 'yyyyMM')}_${CONFIG.COMPANY.NAME}`;

  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();
  body.clear();

  // ---- ドキュメントスタイル定数 ----
  const H1 = DocumentApp.ParagraphHeading.HEADING1;
  const H2 = DocumentApp.ParagraphHeading.HEADING2;
  const H3 = DocumentApp.ParagraphHeading.HEADING3;

  // ---- 表紙 ----
  const titlePara = body.appendParagraph(`月次経営レポート\n${yearMonth}`);
  titlePara.setHeading(H1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(`発行: ${CONFIG.COMPANY.NAME}`)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(`作成日: ${Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy年MM月dd日')}`)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('');

  // ---- 1. 今月のハイライト ----
  body.appendParagraph('1. 今月のハイライト').setHeading(H2);
  const highlights = buildHighlights_(kpi, today);
  highlights.forEach(h => body.appendParagraph('• ' + h));
  body.appendParagraph('');

  // ---- 2. 支援概況 ----
  body.appendParagraph('2. 支援概況').setHeading(H2);
  appendDocTable_(body, [
    ['項目', '値'],
    ['支援中外国人数', kpi.activeWorkers + ' 名'],
    ['受入企業数', kpi.companyCount + ' 社'],
    ['1社あたり平均支援人数', formatNum_(kpi.avgPerCompany, 1) + ' 名'],
    ['在留期限90日以内（リスク）', kpi.visaAlertCount + ' 名'],
  ]);
  body.appendParagraph('');

  // ---- 3. 売上状況 ----
  body.appendParagraph('3. 売上状況').setHeading(H2);
  appendDocTable_(body, [
    ['項目', '金額'],
    ['当月売上（請求額）', formatYen_(kpi.thisMonthSales)],
    ['前月売上（請求額）', formatYen_(kpi.lastMonthSales)],
    ['前年同月売上', formatYen_(kpi.lastYearSales)],
    ['直近12ヶ月 売上累計', formatYen_(kpi.sales12Total)],
    ['当月 入金額', formatYen_(kpi.thisMonthPaid)],
    ['入金率（直近12ヶ月）', formatPct_(kpi.collectionRate)],
    ['未回収額（未払い合計）', formatYen_(kpi.uncollected)],
  ]);
  body.appendParagraph('');

  // ---- 4. 国籍別内訳 ----
  body.appendParagraph('4. 国籍別 支援人数内訳').setHeading(H2);
  if (kpi.nationalityBreakdown.length > 0) {
    const natTableData = [['国籍', '人数（名）', '割合']];
    kpi.nationalityBreakdown.forEach(([label, count]) => {
      const pct = kpi.activeWorkers > 0 ? (count / kpi.activeWorkers * 100).toFixed(1) : '0.0';
      natTableData.push([label, String(count), pct + '%']);
    });
    appendDocTable_(body, natTableData);
  } else {
    body.appendParagraph('（データなし）');
  }
  body.appendParagraph('');

  // ---- 5. 分野別内訳 ----
  body.appendParagraph('5. 分野別 支援人数内訳').setHeading(H2);
  if (kpi.sectorBreakdown.length > 0) {
    const secTableData = [['分野', '人数（名）', '割合']];
    kpi.sectorBreakdown.forEach(([label, count]) => {
      const pct = kpi.activeWorkers > 0 ? (count / kpi.activeWorkers * 100).toFixed(1) : '0.0';
      secTableData.push([label, String(count), pct + '%']);
    });
    appendDocTable_(body, secTableData);
  } else {
    body.appendParagraph('（データなし）');
  }
  body.appendParagraph('');

  // ---- 6. 月次売上推移（直近12ヶ月） ----
  body.appendParagraph('6. 月次売上推移（直近12ヶ月）').setHeading(H2);
  if (kpi.monthlySales.length > 0) {
    const salesTableData = [['月', '請求額', '入金額', '未払い額', '支援人数']];
    kpi.monthlySales.forEach(m => {
      salesTableData.push([
        m.ym,
        formatYen_(m.invoiced),
        formatYen_(m.paid),
        formatYen_(m.invoiced - m.paid),
        String(m.workers),
      ]);
    });
    appendDocTable_(body, salesTableData);
  } else {
    body.appendParagraph('（データなし）');
  }
  body.appendParagraph('');

  // ---- 7. 要対応事項 ----
  body.appendParagraph('7. 要対応事項').setHeading(H2);

  body.appendParagraph('7-1. 未払い・入金催促').setHeading(H3);
  const unpaidItems = buildUnpaidItems_(ss, today);
  if (unpaidItems.length > 0) {
    const unpaidTableData = [['企業名', '請求書番号', '請求額', '支払期限', '超過日数']];
    unpaidItems.forEach(u => unpaidTableData.push(u));
    appendDocTable_(body, unpaidTableData);
  } else {
    body.appendParagraph('✅ 未払い請求はありません。');
  }
  body.appendParagraph('');

  body.appendParagraph('7-2. 在留期限90日以内（ビザ更新対応）').setHeading(H3);
  const visaItems = buildVisaAlertItems_(ss, today);
  if (visaItems.length > 0) {
    const visaTableData = [['氏名', '国籍', '在留期限', '残日数', '受入企業']];
    visaItems.forEach(v => visaTableData.push(v));
    appendDocTable_(body, visaTableData);
  } else {
    body.appendParagraph('✅ 在留期限90日以内の外国人はいません。');
  }
  body.appendParagraph('');

  // ---- フッター ----
  body.appendParagraph('---').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(CONFIG.COMPANY.NAME + '　' + CONFIG.COMPANY.REGISTRATION_NUMBER)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(CONFIG.COMPANY.ADDRESS + '　TEL: ' + CONFIG.COMPANY.PHONE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  doc.saveAndClose();

  // REPORTSフォルダがあれば移動
  const reportsFolderId = getFolderIdFromSettings_('ReportsFolderID');
  if (reportsFolderId) {
    try {
      const reportsFolder = DriveApp.getFolderById(reportsFolderId);
      DriveApp.getFileById(doc.getId()).moveTo(reportsFolder);
    } catch (e) {
      console.log('レポートフォルダへの移動に失敗: ' + e.message);
    }
  }

  const docUrl = doc.getUrl();
  SpreadsheetApp.getUi().alert(
    '✅ 月次経営レポートを作成しました。\n\n' +
    'タイトル: ' + docTitle + '\n' +
    'URL: ' + docUrl
  );
  return docUrl;
}

// ============================================================
// 4. refreshAnalytics — buildAnalyticsSheet + insertCharts をまとめて実行
// ============================================================
function refreshAnalytics() {
  buildAnalyticsSheet();
  insertCharts();
  SpreadsheetApp.getUi().alert(
    '✅ 経営分析シートを更新しました。\n\n' +
    'シート「' + CONFIG.SHEETS.ANALYTICS + '」にKPIとグラフを反映しました。'
  );
}

// ============================================================
// 内部: KPIデータ収集
// ============================================================
function collectKpiData_(ss) {
  const today = new Date();
  const thisYm  = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM');
  const lastDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const lastYm  = Utilities.formatDate(lastDate, 'Asia/Tokyo', 'yyyy/MM');
  const lastYearDate = new Date(today.getFullYear() - 1, today.getMonth(), 1);
  const lastYearYm = Utilities.formatDate(lastYearDate, 'Asia/Tokyo', 'yyyy/MM');

  // ---- 外国人データ集計 ----
  let activeWorkers = 0;
  let visaAlertCount = 0;
  const nationalityMap = {};
  const sectorMap = {};

  const workersSheet = ss.getSheetByName(CONFIG.SHEETS.WORKERS);
  if (workersSheet && workersSheet.getLastRow() > 1) {
    const workerData = workersSheet.getDataRange().getValues();
    for (let i = 1; i < workerData.length; i++) {
      const row = workerData[i];
      if (!row[0]) continue;                       // ID空行はスキップ
      if (row[3] !== '在籍中') continue;           // D列: ステータス
      activeWorkers++;

      // 国籍 (E列 = col[4])
      const nat = row[4] || 'その他';
      nationalityMap[nat] = (nationalityMap[nat] || 0) + 1;

      // 分野 (L列 = col[11])
      const sec = row[11] || '不明';
      sectorMap[sec] = (sectorMap[sec] || 0) + 1;

      // 在留期限90日以内 (N列 = col[13])
      const expiry = row[13];
      if (expiry) {
        const daysLeft = Math.floor((new Date(expiry) - today) / 86400000);
        if (daysLeft >= 0 && daysLeft <= 90) visaAlertCount++;
      }
    }
  }

  // ---- 企業数 ----
  let companyCount = 0;
  const companiesSheet = ss.getSheetByName(CONFIG.SHEETS.COMPANIES);
  if (companiesSheet && companiesSheet.getLastRow() > 1) {
    const compData = companiesSheet.getDataRange().getValues();
    for (let i = 1; i < compData.length; i++) {
      if (compData[i][0] && compData[i][13] !== 'キャンセル') companyCount++;
    }
  }

  const avgPerCompany = companyCount > 0 ? activeWorkers / companyCount : 0;

  // ---- 請求書データ集計 ----
  let thisMonthSales  = 0;
  let lastMonthSales  = 0;
  let lastYearSales   = 0;
  let thisMonthPaid   = 0;
  let sales12Total    = 0;
  let paid12Total     = 0;
  let uncollected     = 0;

  // 直近12ヶ月のYM文字列セット
  const last12Yms = new Set();
  for (let i = 0; i < 12; i++) {
    const d = new Date(today.getFullYear(), today.getMonth() - i, 1);
    last12Yms.add(Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM'));
  }

  // 月次データ: { ym -> { invoiced, paid, workers } }
  const monthlyMap = {};
  last12Yms.forEach(ym => { monthlyMap[ym] = { invoiced: 0, paid: 0, workers: 0 }; });

  const invoicesSheet = ss.getSheetByName(CONFIG.SHEETS.INVOICES);
  if (invoicesSheet && invoicesSheet.getLastRow() > 1) {
    const invData = invoicesSheet.getDataRange().getValues();
    for (let i = 1; i < invData.length; i++) {
      const row = invData[i];
      if (!row[0]) continue;
      const ym     = String(row[3]);    // D列: 請求年月 (yyyy/MM)
      const amount = Number(row[5]) || 0;  // F列: 請求金額
      const status = String(row[6]);    // G列: 支払ステータス
      const workers = Number(row[10]) || 0; // K列: 担当外国人数

      if (ym === thisYm)    thisMonthSales += amount;
      if (ym === lastYm)    lastMonthSales += amount;
      if (ym === lastYearYm) lastYearSales += amount;

      if (ym === thisYm && status === '支払済') thisMonthPaid += amount;
      if (status === '未払い') uncollected += amount;

      if (last12Yms.has(ym)) {
        sales12Total += amount;
        if (status === '支払済') paid12Total += amount;
        monthlyMap[ym].invoiced += amount;
        if (status === '支払済') monthlyMap[ym].paid += amount;
        monthlyMap[ym].workers += workers;
      }
    }
  }

  const collectionRate = sales12Total > 0 ? paid12Total / sales12Total : 0;

  // 月次配列: 古い順に並べる
  const monthlySales = Array.from(last12Yms)
    .sort()
    .map(ym => ({ ym, ...monthlyMap[ym] }));

  // 国籍・分野ブレークダウン: 多い順にソート
  const nationalityBreakdown = Object.entries(nationalityMap)
    .sort((a, b) => b[1] - a[1]);
  const sectorBreakdown = Object.entries(sectorMap)
    .sort((a, b) => b[1] - a[1]);

  return {
    activeWorkers,
    companyCount,
    avgPerCompany,
    visaAlertCount,
    thisMonthSales,
    lastMonthSales,
    lastYearSales,
    thisMonthPaid,
    sales12Total,
    collectionRate,
    uncollected,
    monthlySales,
    nationalityBreakdown,
    sectorBreakdown,
  };
}

// ============================================================
// 内部: シート書き込みヘルパー
// ============================================================

/** タイトル行を書き込み、次の行番号を返す */
function writeTitle_(sheet, row, text) {
  const range = sheet.getRange(row, 1, 1, 5).merge();
  range.setValue(text)
    .setFontSize(14).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(row, 40);
  return row + 2;
}

/** セクションヘッダーを書き込み、次の行番号を返す */
function writeSectionHeader_(sheet, row, text) {
  const range = sheet.getRange(row, 1, 1, 4).merge();
  range.setValue(text)
    .setFontWeight('bold').setFontSize(11)
    .setBackground('#e8f0fe')
    .setFontColor(CONFIG.COLORS.HEADER)
    .setBorder(false, false, true, false, false, false,
               CONFIG.COLORS.HEADER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(row, 28);
  return row + 1;
}

/**
 * KPIテーブルを書き込む。
 * rows: [[ラベル, 値, 単位, 書式文字列(null=標準)], ...]
 */
function writeKpiTable_(sheet, row, rows) {
  styleHeaderRow_(sheet, row, 4);
  sheet.getRange(row, 1, 1, 4).setValues([['指標', '値', '単位', '前月比・備考']]);
  row++;
  rows.forEach(([label, value, unit, fmt], idx) => {
    const isAlt = idx % 2 === 1;
    const labelRange = sheet.getRange(row, 1);
    const valueRange = sheet.getRange(row, 2);
    const unitRange  = sheet.getRange(row, 3);

    labelRange.setValue(label).setFontWeight('bold');
    valueRange.setValue(value);
    unitRange.setValue(unit);

    if (fmt) valueRange.setNumberFormat(fmt);
    if (isAlt) {
      sheet.getRange(row, 1, 1, 4).setBackground(CONFIG.COLORS.ROW_ALT);
    }
    row++;
  });
  return row + 1;
}

/**
 * ブレークダウンテーブル（国籍別・分野別）を書き込む。
 * data: [[label, count], ...]
 */
function writeBreakdownTable_(sheet, row, data, col1Header, col2Header) {
  if (data.length === 0) {
    sheet.getRange(row, 1).setValue('（データなし）').setFontColor('#999999');
    return row + 2;
  }
  styleHeaderRow_(sheet, row, 3);
  sheet.getRange(row, 1, 1, 3).setValues([[col1Header, col2Header, '割合']]);
  row++;
  const total = data.reduce((s, [, v]) => s + v, 0);
  data.forEach(([label, count], idx) => {
    const isAlt = idx % 2 === 1;
    const pct = total > 0 ? count / total : 0;
    sheet.getRange(row, 1).setValue(label);
    sheet.getRange(row, 2).setValue(count);
    sheet.getRange(row, 3).setValue(pct).setNumberFormat('0.0%');
    if (isAlt) sheet.getRange(row, 1, 1, 3).setBackground(CONFIG.COLORS.ROW_ALT);
    row++;
  });
  return row + 1;
}

/**
 * 月次売上テーブルを書き込む。
 * monthlySales: [{ ym, invoiced, paid, workers }, ...]
 */
function writeMonthlySalesTable_(sheet, row, monthlySales) {
  if (monthlySales.length === 0) {
    sheet.getRange(row, 1).setValue('（データなし）').setFontColor('#999999');
    return row + 2;
  }
  const headers = ['月', '請求額', '入金額', '未払い額', '支援人数'];
  styleHeaderRow_(sheet, row, headers.length);
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  row++;
  monthlySales.forEach((m, idx) => {
    const isAlt = idx % 2 === 1;
    const unpaid = m.invoiced - m.paid;
    sheet.getRange(row, 1).setValue(m.ym);
    sheet.getRange(row, 2).setValue(m.invoiced).setNumberFormat('¥#,##0');
    sheet.getRange(row, 3).setValue(m.paid).setNumberFormat('¥#,##0');
    sheet.getRange(row, 4).setValue(unpaid).setNumberFormat('¥#,##0');
    sheet.getRange(row, 5).setValue(m.workers);
    if (isAlt) sheet.getRange(row, 1, 1, 5).setBackground(CONFIG.COLORS.ROW_ALT);
    row++;
  });
  return row + 1;
}

// ============================================================
// 内部: チャート用テーブル範囲検索
// ============================================================

/**
 * シート内でヘッダー行を探し { startRow, startCol, numRows } を返す。
 * firstCellValue と完全一致する最初のセルを探す。
 * 見つからない場合は null を返す。
 */
function findTableRange_(sheet, firstCellValue) {
  const data = sheet.getDataRange().getValues();
  for (let r = 0; r < data.length; r++) {
    if (data[r][0] === firstCellValue) {
      // このヘッダー行から空行まで続くデータ行を数える
      let numRows = 1;
      for (let r2 = r + 1; r2 < data.length; r2++) {
        if (data[r2][0] === '' && data[r2][1] === '') break;
        numRows++;
      }
      return { startRow: r + 1, startCol: 1, numRows };
    }
  }
  return null;
}

// ============================================================
// 内部: Google Doc 書き込みヘルパー
// ============================================================

/** Doc本文に表を追加 */
function appendDocTable_(body, rows) {
  if (!rows || rows.length === 0) return;
  const table = body.appendTable(rows);
  // ヘッダー行を太字に
  const headerRow = table.getRow(0);
  for (let c = 0; c < headerRow.getNumCells(); c++) {
    headerRow.getCell(c).editAsText().setBold(true);
  }
}

/** ハイライト文言を組み立てる */
function buildHighlights_(kpi, today) {
  const ymLabel = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy年MM月');
  const highlights = [];

  highlights.push(`支援中外国人数: ${kpi.activeWorkers}名（受入企業: ${kpi.companyCount}社）`);

  if (kpi.thisMonthSales > 0) {
    highlights.push(`${ymLabel}の請求額合計: ${formatYen_(kpi.thisMonthSales)}`);
  }
  if (kpi.lastMonthSales > 0 && kpi.thisMonthSales > 0) {
    const diff = kpi.thisMonthSales - kpi.lastMonthSales;
    const sign = diff >= 0 ? '▲' : '▼';
    highlights.push(`前月比: ${sign}${formatYen_(Math.abs(diff))} (${diff >= 0 ? '+' : ''}${kpi.lastMonthSales > 0 ? Math.round(diff / kpi.lastMonthSales * 100) : 0}%)`);
  }
  if (kpi.lastYearSales > 0 && kpi.thisMonthSales > 0) {
    const diff = kpi.thisMonthSales - kpi.lastYearSales;
    const sign = diff >= 0 ? '▲' : '▼';
    highlights.push(`前年同月比: ${sign}${formatYen_(Math.abs(diff))} (${diff >= 0 ? '+' : ''}${Math.round(diff / kpi.lastYearSales * 100)}%)`);
  }
  if (kpi.uncollected > 0) {
    highlights.push(`未回収額: ${formatYen_(kpi.uncollected)} — 早急な入金確認が必要です。`);
  }
  if (kpi.visaAlertCount > 0) {
    highlights.push(`在留期限90日以内: ${kpi.visaAlertCount}名 — ビザ更新手続きを進めてください。`);
  }
  if (highlights.length === 0) highlights.push('特筆事項はありません。');
  return highlights;
}

/** 未払い請求の一覧行を生成 */
function buildUnpaidItems_(ss, today) {
  const items = [];
  const invoicesSheet = ss.getSheetByName(CONFIG.SHEETS.INVOICES);
  if (!invoicesSheet || invoicesSheet.getLastRow() <= 1) return items;
  const data = invoicesSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || row[6] !== '未払い') continue;
    const dueDate = row[8];
    let overdueStr = '';
    if (dueDate) {
      const overdue = Math.floor((today - new Date(dueDate)) / 86400000);
      overdueStr = overdue > 0 ? overdue + '日超過' : '期限内';
    }
    items.push([
      String(row[2]),
      String(row[0]),
      formatYen_(Number(row[5]) || 0),
      dueDate ? Utilities.formatDate(new Date(dueDate), 'Asia/Tokyo', 'yyyy/MM/dd') : '不明',
      overdueStr,
    ]);
  }
  return items;
}

/** 在留期限90日以内の外国人一覧行を生成 */
function buildVisaAlertItems_(ss, today) {
  const items = [];
  const workersSheet = ss.getSheetByName(CONFIG.SHEETS.WORKERS);
  if (!workersSheet || workersSheet.getLastRow() <= 1) return items;
  const data = workersSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || row[3] !== '在籍中') continue;
    const expiry = row[13];
    if (!expiry) continue;
    const daysLeft = Math.floor((new Date(expiry) - today) / 86400000);
    if (daysLeft >= 0 && daysLeft <= 90) {
      items.push([
        String(row[1]),
        String(row[4]),
        Utilities.formatDate(new Date(expiry), 'Asia/Tokyo', 'yyyy/MM/dd'),
        String(daysLeft) + '日',
        String(row[10]),
      ]);
    }
  }
  // 残日数が少ない順にソート
  items.sort((a, b) => parseInt(a[3]) - parseInt(b[3]));
  return items;
}

// ============================================================
// 内部: 書式ヘルパー
// ============================================================

function formatYen_(value) {
  const n = Number(value) || 0;
  return '¥' + n.toLocaleString('ja-JP');
}

function formatPct_(value) {
  return (Math.round((Number(value) || 0) * 1000) / 10).toFixed(1) + '%';
}

function formatNum_(value, decimals) {
  return (Number(value) || 0).toFixed(decimals || 0);
}

/*
==========================================================================
【Menu.gs 追加項目】
以下を Menu.gs の onOpen() 内の適切な場所に追加してください:

.addItem('📊 経営分析シートを更新', 'refreshAnalytics')
.addItem('📈 グラフを再生成',       'insertCharts')
.addItem('📄 月次経営レポート作成', 'createMonthlyReportDoc')

例（サブメニューとして追加する場合）:
ui.createMenu('📊 経営分析')
  .addItem('経営分析シートを更新', 'refreshAnalytics')
  .addItem('グラフを再生成',       'insertCharts')
  .addItem('月次経営レポート作成', 'createMonthlyReportDoc')
  .addToUi();

==========================================================================
【CONFIG.SHEETS への追加提案】
Config.gs の SHEETS オブジェクトに以下を追加すると
他のファイルから参照しやすくなります:

  ANALYTICS: '経営分析',

追加後は Analytics.gs のローカル定数 CONFIG.SHEETS.ANALYTICS を
CONFIG.SHEETS.ANALYTICS に置き換えてください。

==========================================================================
*/
