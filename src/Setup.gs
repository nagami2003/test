// ===== 初期セットアップ =====

function createFolderStructure() {
  const ui = SpreadsheetApp.getUi();
  const root = getOrCreateFolder_(DriveApp.getRootFolder(), CONFIG.FOLDERS.ROOT);

  const subFolders = [
    CONFIG.FOLDERS.WORKERS,
    CONFIG.FOLDERS.COMPANIES,
    CONFIG.FOLDERS.CONTRACTS,
    CONFIG.FOLDERS.REPORTS,
    CONFIG.FOLDERS.INVOICES,
    CONFIG.FOLDERS.TEMPLATES,
    CONFIG.FOLDERS.CORRESPONDENCE,
  ];

  subFolders.forEach(name => getOrCreateFolder_(root, name));

  // フォルダIDを設定シートに保存
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(CONFIG.SHEETS.SETTINGS);
  }
  settingsSheet.clearContents();
  settingsSheet.getRange('A1:B1').setValues([['設定項目', '値']]);
  settingsSheet.getRange('A2:B2').setValues([['RootFolderID', root.getId()]]);
  settingsSheet.getRange('A3:B3').setValues([['WorkersFolderID', getOrCreateFolder_(root, CONFIG.FOLDERS.WORKERS).getId()]]);
  settingsSheet.getRange('A4:B4').setValues([['CompaniesFolderID', getOrCreateFolder_(root, CONFIG.FOLDERS.COMPANIES).getId()]]);
  settingsSheet.getRange('A5:B5').setValues([['InvoicesFolderID', getOrCreateFolder_(root, CONFIG.FOLDERS.INVOICES).getId()]]);
  settingsSheet.getRange('A6:B6').setValues([['TemplatesFolderID', getOrCreateFolder_(root, CONFIG.FOLDERS.TEMPLATES).getId()]]);
  settingsSheet.getRange('A7:B7').setValues([['ReportsFolderID', getOrCreateFolder_(root, CONFIG.FOLDERS.REPORTS).getId()]]);
  settingsSheet.getRange('A8:B8').setValues([['ContractsFolderID', getOrCreateFolder_(root, CONFIG.FOLDERS.CONTRACTS).getId()]]);

  styleSettingsSheet_(settingsSheet);
  ui.alert('✅ Driveフォルダ構造の作成が完了しました。\n\nフォルダ: ' + CONFIG.FOLDERS.ROOT);
}

function initializeAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  initDashboard_(ss);
  initWorkersSheet_(ss);
  initCompaniesSheet_(ss);
  initSupportPlansSheet_(ss);
  initInterviewsSheet_(ss);
  initInvoicesSheet_(ss);
  initSalesSheet_(ss);
  updateDashboard();
  SpreadsheetApp.getUi().alert('✅ 全シートの初期化が完了しました。');
}

function initDashboard_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.DASHBOARD, 0);
  sheet.clearContents();
  sheet.clearFormats();

  sheet.getRange('A1:F1').merge().setValue('📊 登録支援機関 管理ダッシュボード')
    .setFontSize(18).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER).setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 50);

  const kpiLabels = [
    ['支援中 外国人数', '=COUNTIF(外国人管理!D:D,"在籍中")'],
    ['受入企業数', '=COUNTA(受入企業管理!A:A)-1'],
    ['今月の請求額合計', '=SUMIFS(請求書管理!F:F,請求書管理!C:C,TEXT(TODAY(),"yyyy/mm"))'],
    ['未払い請求件数', '=COUNTIF(請求書管理!G:G,"未払い")'],
    ['期限30日以内のビザ', '=COUNTIFS(外国人管理!K:K,"<="&(TODAY()+30),外国人管理!K:K,">="&TODAY(),外国人管理!D:D,"在籍中")'],
    ['今四半期の未実施面談', '=COUNTIF(定期面談記録!F:F,"未実施")'],
  ];

  kpiLabels.forEach(([label, formula], i) => {
    const col = (i % 3) * 2 + 1;
    const row = Math.floor(i / 3) * 3 + 3;
    sheet.getRange(row, col).setValue(label).setFontWeight('bold').setFontSize(11);
    sheet.getRange(row + 1, col).setFormula(formula).setFontSize(24).setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.getRange(row, col, 2, 1).setBorder(true, true, true, true, false, false, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  });

  sheet.getRange('A9:F9').merge().setValue('⚠️ 直近のアラート')
    .setFontWeight('bold').setBackground('#fce8b2').setFontSize(12);
  sheet.getRange('A10:F10').setValues([['種別', '氏名/企業名', '内容', '期限', 'ステータス', '担当']]);
  styleHeaderRow_(sheet, 10, 6);
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 120);
}

function initWorkersSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.WORKERS);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.WORKERS);
  sheet.clearContents();
  sheet.clearFormats();

  const headers = [
    'ID', '氏名（ローマ字）', '氏名（母国語）', 'ステータス', '国籍', '生年月日',
    '性別', '連絡先（電話）', '連絡先（Email）', '受入企業ID', '受入企業名',
    '分野', '在留資格', '在留期限', '在留カード番号', '入国日', '就労開始日',
    '支援計画書', '個人フォルダURL', '備考',
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, 1, headers.length);

  // データバリデーション
  const statusRange = sheet.getRange('D2:D1000');
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['在籍中', '帰国', '転職', '在留期限切れ', '一時帰国'], true).build();
  statusRange.setDataValidation(statusRule);

  const residenceRange = sheet.getRange('M2:M1000');
  const residenceRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.RESIDENCE_STATUS, true).build();
  residenceRange.setDataValidation(residenceRule);

  const nationalityRange = sheet.getRange('E2:E1000');
  const nationalityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.NATIONALITIES, true).build();
  nationalityRange.setDataValidation(nationalityRule);

  const sectorRange = sheet.getRange('L2:L1000');
  const sectorRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.SECTORS, true).build();
  sectorRange.setDataValidation(sectorRule);

  // 在留期限の条件付き書式（赤→黄→緑）
  const expiryRange = sheet.getRange('N2:N1000');
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(N2<>"",N2<=TODAY()+30)')
      .setBackground(CONFIG.COLORS.ALERT_RED).setFontColor('#ffffff').build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(N2<>"",N2<=TODAY()+90)')
      .setBackground(CONFIG.COLORS.ALERT_YELLOW).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(N2<>"",N2>TODAY()+90)')
      .setBackground(CONFIG.COLORS.ALERT_GREEN).build(),
  ];
  sheet.setConditionalFormatRules(rules);

  autoResizeColumns_(sheet, headers.length);
  sheet.setFrozenRows(1);
}

function initCompaniesSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.COMPANIES);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.COMPANIES);
  sheet.clearContents();
  sheet.clearFormats();

  const headers = [
    '企業ID', '企業名', '法人番号', '代表者名', '担当者名', '担当者連絡先',
    '住所', '業種（分野）', '契約開始日', '契約終了日', '月額支援費',
    '支援中外国人数', '企業フォルダURL', 'ステータス', '備考',
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, 1, headers.length);

  const contractEndRange = sheet.getRange('J2:J1000');
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(J2<>"",J2<=TODAY()+30)')
      .setBackground(CONFIG.COLORS.ALERT_RED).setFontColor('#ffffff').build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(J2<>"",J2<=TODAY()+90)')
      .setBackground(CONFIG.COLORS.ALERT_YELLOW).build(),
  ];
  sheet.setConditionalFormatRules(rules);

  autoResizeColumns_(sheet, headers.length);
  sheet.setFrozenRows(1);
}

function initSupportPlansSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.SUPPORT_PLANS);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.SUPPORT_PLANS);
  sheet.clearContents();
  sheet.clearFormats();

  const baseHeaders = ['計画ID', '外国人ID', '氏名', '受入企業名', '計画作成日', '計画URL'];
  const headers = [...baseHeaders, ...CONFIG.SUPPORT_TYPES, '完了率', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, 1, headers.length);

  // 支援種別のステータス入力規則
  const statusList = ['未実施', '実施済', '該当なし', '実施中'];
  for (let col = baseHeaders.length + 1; col <= baseHeaders.length + CONFIG.SUPPORT_TYPES.length; col++) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(statusList, true).build();
    sheet.getRange(2, col, 1000, 1).setDataValidation(rule);
  }

  autoResizeColumns_(sheet, headers.length);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);
}

function initInterviewsSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.INTERVIEWS);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.INTERVIEWS);
  sheet.clearContents();
  sheet.clearFormats();

  const headers = [
    '面談ID', '外国人ID', '氏名', '受入企業名', '面談予定日', 'ステータス',
    '面談実施日', '面談者', '面談方法', '面談場所',
    '就労状況', '生活状況', '健康状態', '相談内容', '対応内容',
    '次回面談予定', '報告書URL', '備考',
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, 1, headers.length);

  const statusRange = sheet.getRange('F2:F1000');
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未実施', '実施済', 'キャンセル', '延期'], true).build();
  statusRange.setDataValidation(statusRule);

  const methodRange = sheet.getRange('I2:I1000');
  const methodRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['対面', 'オンライン（Google Meet）', '電話', '書面'], true).build();
  methodRange.setDataValidation(methodRule);

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(E2<>"",E2<=TODAY(),F2="未実施")')
      .setBackground(CONFIG.COLORS.ALERT_RED).setFontColor('#ffffff').build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(E2<>"",E2<=TODAY()+14,F2="未実施")')
      .setBackground(CONFIG.COLORS.ALERT_YELLOW).build(),
  ];
  sheet.setConditionalFormatRules(rules);

  autoResizeColumns_(sheet, headers.length);
  sheet.setFrozenRows(1);
}

function initInvoicesSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.INVOICES);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.INVOICES);
  sheet.clearContents();
  sheet.clearFormats();

  const headers = [
    '請求書番号', '受入企業ID', '受入企業名', '請求年月', '発行日',
    '請求金額（税込）', '支払ステータス', '入金日', '支払期限',
    '請求書URL', '担当外国人数', '内訳', '備考',
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, 1, headers.length);

  const statusRange = sheet.getRange('G2:G1000');
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未払い', '支払済', 'キャンセル', '分割払い'], true).build();
  statusRange.setDataValidation(statusRule);

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(I2<>"",I2<TODAY(),G2="未払い")')
      .setBackground(CONFIG.COLORS.ALERT_RED).setFontColor('#ffffff').build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=G2="支払済"')
      .setBackground('#e6f4ea').build(),
  ];
  sheet.setConditionalFormatRules(rules);

  autoResizeColumns_(sheet, headers.length);
  sheet.setFrozenRows(1);
}

function initSalesSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.SALES);
  sheet.clearContents();
  sheet.clearFormats();

  sheet.getRange('A1').setValue('売上サマリー').setFontSize(14).setFontWeight('bold');
  const monthHeaders = ['月', '請求額合計', '入金額合計', '未払い額', '支援人数', '備考'];
  sheet.getRange('A3:F3').setValues([monthHeaders]);
  styleHeaderRow_(sheet, 3, 6);

  const now = new Date();
  for (let i = 11; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const ym = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM');
    const row = 4 + (11 - i);
    sheet.getRange(row, 1).setValue(ym);
    sheet.getRange(row, 2).setFormula(`=SUMIFS(請求書管理!F:F,請求書管理!D:D,"${ym}")`);
    sheet.getRange(row, 3).setFormula(`=SUMIFS(請求書管理!F:F,請求書管理!D:D,"${ym}",請求書管理!G:G,"支払済")`);
    sheet.getRange(row, 4).setFormula(`=B${row}-C${row}`);
    sheet.getRange(row, 5).setFormula(`=SUMIF(請求書管理!D:D,"${ym}",請求書管理!K:K)`);
  }

  sheet.getRange('B4:E15').setNumberFormat('¥#,##0');
  autoResizeColumns_(sheet, 6);
}

function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!sheet) return;

  const today = new Date();
  const alertStartRow = 11;
  sheet.getRange(`A${alertStartRow}:F${alertStartRow + 50}`).clearContent();

  let alertRow = alertStartRow;

  // ビザ期限アラート
  const workersSheet = ss.getSheetByName(CONFIG.SHEETS.WORKERS);
  if (workersSheet) {
    const data = workersSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) break;
      const status = data[i][3];
      const expiry = data[i][13];
      const name = data[i][1];
      if (status !== '在籍中' || !expiry) continue;
      const daysLeft = Math.floor((new Date(expiry) - today) / 86400000);
      if (daysLeft <= 90) {
        const urgency = daysLeft <= 30 ? '🔴 緊急' : daysLeft <= 60 ? '🟠 要注意' : '🟡 注意';
        sheet.getRange(alertRow, 1, 1, 6).setValues([[
          'ビザ期限', name, `在留期限まで${daysLeft}日`,
          Utilities.formatDate(new Date(expiry), 'Asia/Tokyo', 'yyyy/MM/dd'),
          urgency, data[i][9],
        ]]);
        if (daysLeft <= 30) sheet.getRange(alertRow, 1, 1, 6).setBackground('#fce8e6');
        else if (daysLeft <= 60) sheet.getRange(alertRow, 1, 1, 6).setBackground('#fef7e0');
        alertRow++;
      }
    }
  }

  // 未払い請求アラート
  const invoicesSheet = ss.getSheetByName(CONFIG.SHEETS.INVOICES);
  if (invoicesSheet) {
    const data = invoicesSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) break;
      if (data[i][6] !== '未払い') continue;
      const dueDate = data[i][8];
      const daysOverdue = dueDate ? Math.floor((today - new Date(dueDate)) / 86400000) : 0;
      if (daysOverdue >= 0) {
        sheet.getRange(alertRow, 1, 1, 6).setValues([[
          '請求未払い', data[i][2], `${daysOverdue}日超過`,
          Utilities.formatDate(new Date(dueDate), 'Asia/Tokyo', 'yyyy/MM/dd'),
          '🔴 入金確認', data[i][0],
        ]]);
        sheet.getRange(alertRow, 1, 1, 6).setBackground('#fce8e6');
        alertRow++;
      }
    }
  }

  if (alertRow === alertStartRow) {
    sheet.getRange(alertRow, 1, 1, 6).merge().setValue('✅ 現在、緊急のアラートはありません')
      .setHorizontalAlignment('center').setBackground('#e6f4ea');
  }

  SpreadsheetApp.flush();
}

// ---------- 内部ユーティリティ ----------
function getOrCreateFolder_(parent, name) {
  const existing = parent.getFoldersByName(name);
  return existing.hasNext() ? existing.next() : parent.createFolder(name);
}

function styleHeaderRow_(sheet, row, numCols) {
  const range = sheet.getRange(row, 1, 1, numCols);
  range.setBackground(CONFIG.COLORS.HEADER)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, false);
}

function styleSettingsSheet_(sheet) {
  sheet.getRange('A1:B1').setBackground(CONFIG.COLORS.HEADER)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT).setFontWeight('bold');
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 400);
}

function autoResizeColumns_(sheet, numCols) {
  for (let i = 1; i <= numCols; i++) {
    sheet.autoResizeColumn(i);
    const width = sheet.getColumnWidth(i);
    if (width < 80) sheet.setColumnWidth(i, 80);
    if (width > 300) sheet.setColumnWidth(i, 300);
  }
}

function getFolderIdFromSettings_(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!settingsSheet) return null;
  const data = settingsSheet.getDataRange().getValues();
  for (const row of data) {
    if (row[0] === key) return row[1];
  }
  return null;
}
