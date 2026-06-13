// ===== 行政コンプライアンス支援（特定技能 届出管理） =====
//
// 【重要な注意事項】
// 特定技能制度に関する届出の様式・提出期限・提出先は、出入国在留管理庁による
// 制度改正・通達により変更されることがあります。
// 本スクリプトの内容はあくまで業務補助を目的とした参考情報です。
// 実際の届出に際しては、必ず最新の出入国在留管理庁の様式・通達・
// 各地方出入国在留管理局の案内を確認してください。
//
// 参考URL: https://www.moj.go.jp/isa/applications/status/tokutei_index.html

// ---------- ローカル定数 ----------
// Config.gs の SHEETS への追加を提案: COMPLIANCE: 'コンプライアンス届出管理'
var COMPLIANCE_SHEET_NAME = 'コンプライアンス届出管理';

// 届出管理シートの列インデックス（1始まり）
var COMPLIANCE_COL = {
  ID:         1,   // 届出ID
  TYPE:       2,   // 届出種別
  TARGET:     3,   // 対象（受入機関/登録支援機関/氏名等）
  DEADLINE:   4,   // 提出期限
  STATUS:     5,   // ステータス（未提出/提出済/期限超過）
  SUBMIT_DATE:6,   // 提出日
  ASSIGNEE:   7,   // 担当
  NOTES:      8,   // 備考
};

// 届出種別（定期届出）
var NOTIFICATION_TYPES = {
  RSO_QUARTERLY:   '支援実施状況届出（登録支援機関・定期）',
  EMPLOYER_QUARTERLY: '受入れ・活動状況届出（受入機関・定期）',
  AD_HOC:          '随時届出',
};

// ステータス値
var COMPLIANCE_STATUS = {
  PENDING:   '未提出',
  SUBMITTED: '提出済',
  OVERDUE:   '期限超過',
};


// ============================================================
// 1. createQuarterlyReportChecklistDoc()
//    四半期定期届出のチェックリスト＆下書きを Google Doc で生成
// ============================================================
// 【注意】提出期限・様式番号・記載事項は制度改正で変更される場合があります。
// 実際の届出前に必ず出入国在留管理庁の最新情報を確認してください。
function createQuarterlyReportChecklistDoc() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ui   = SpreadsheetApp.getUi();
  var now  = new Date();

  // ---- 対象四半期の算定 ----
  var quarterInfo = getCurrentQuarterInfo_(now);

  // ---- 外国人管理シートからデータ集計 ----
  var workerSummary = collectWorkerSummary_(ss);

  // ---- Google Doc 生成 ----
  var title = '四半期定期届出チェックリスト_'
    + quarterInfo.label
    + '_'
    + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
  var doc  = DocumentApp.create(title);
  var body = doc.getBody();
  body.clear();

  // タイトル
  var titlePara = body.appendParagraph('特定技能 四半期定期届出 チェックリスト（下書き）');
  titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  titlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  // 注意書きボックス相当
  var notice = body.appendParagraph(
    '【重要】本書は届出作業を補助するための内部資料です。'
    + '様式・提出期限・提出先は出入国在留管理庁の最新情報を必ず確認してください。'
    + '（参考: https://www.moj.go.jp/isa/ ）'
  );
  notice.setItalic(true);

  body.appendHorizontalRule();

  // 基本情報
  body.appendParagraph('■ 基本情報').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendTable([
    ['作成日',       Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日')],
    ['作成者',       '（記入）'],
    ['対象四半期',   quarterInfo.label],
    ['四半期期間',   quarterInfo.period],
    ['届出期限目安', quarterInfo.deadline + '（※最新の案内を要確認）'],
    ['支援機関名',   CONFIG.COMPANY.NAME],
    ['登録番号',     CONFIG.COMPANY.REGISTRATION_NUMBER],
  ]).setBorderColor('#dadce0');

  body.appendParagraph('');

  // 提出書類チェックリスト（登録支援機関）
  body.appendParagraph('■ 提出書類チェックリスト（登録支援機関）')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(
    '下記の書類を確認し、完了したものに「✓」を記入してください。'
    + '\n（様式番号は制度改正により変更される場合があります。出入国在留管理庁のサイトで最新版を確認してください。）'
  ).setItalic(true);

  var rsoChecklist = [
    ['□', '支援実施状況に係る届出書（参考様式第3-8号相当）', ''],
    ['□', '支援対象者一覧', ''],
    ['□', '各支援項目の実施状況記録（定期面談記録等）', ''],
    ['□', '定期面談実施確認書類（面談記録の写し等）', ''],
    ['□', '提出先：管轄地方出入国在留管理局（電子届出システム推奨）', ''],
  ];
  var rsoTable = body.appendTable([['確認', '書類名・対応事項', '担当/備考'], ...rsoChecklist]);
  rsoTable.setBorderColor('#dadce0');

  body.appendParagraph('');

  // 提出書類チェックリスト（受入機関向け案内）
  body.appendParagraph('■ 提出書類チェックリスト（受入機関向け案内）')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(
    '受入機関（特定技能所属機関）が提出する書類の確認・案内を行います。'
  ).setItalic(true);

  var empChecklist = [
    ['□', '受入れ状況に係る届出書（参考様式第3-7号相当）', ''],
    ['□', '活動状況に係る届出書', ''],
    ['□', '報酬・待遇に関する資料', ''],
    ['□', '提出先：管轄地方出入国在留管理局（電子届出システム推奨）', ''],
  ];
  var empTable = body.appendTable([['確認', '書類名・対応事項', '担当/備考'], ...empChecklist]);
  empTable.setBorderColor('#dadce0');

  body.appendParagraph('');

  // 外国人支援状況サマリー
  body.appendParagraph('■ 支援対象外国人 状況サマリー')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(
    '外国人管理シート・定期面談記録シートから自動集計（集計日: '
    + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd') + '）'
  );

  var summaryHeader = ['氏名', '受入企業名', 'ステータス', '直近面談日', '次回面談予定', '備考'];
  if (workerSummary.rows.length > 0) {
    body.appendTable([summaryHeader, ...workerSummary.rows]).setBorderColor('#dadce0');
    body.appendParagraph('在籍中: ' + workerSummary.activeCount + '名 / 直近面談実施済: ' + workerSummary.interviewedCount + '名');
  } else {
    body.appendTable([summaryHeader, ['（データなし）', '', '', '', '', '']]).setBorderColor('#dadce0');
    body.appendParagraph('外国人管理シートにデータがありません。');
  }

  body.appendParagraph('');
  body.appendHorizontalRule();

  // 対応手順メモ
  body.appendParagraph('■ 届出対応手順（参考）').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  var steps = [
    '① 外国人管理シート・定期面談記録シートの記載内容を最終確認する',
    '② 出入国在留管理庁の電子届出システム（OICS）または郵送で届出書を提出する',
    '③ 受付番号・提出日を「届出管理」シートに記録し、ステータスを「提出済」に更新する',
    '④ 提出書類の写しをDriveの報告書フォルダに保存する',
    '⑤ 受入機関への報告・共有を行う',
  ];
  steps.forEach(function(s) { body.appendParagraph(s); });

  body.appendParagraph('');
  body.appendParagraph(
    '※ 本書は '
    + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日')
    + ' 時点のシステム情報を基に自動生成されました。'
    + '最終的な内容は担当者が確認・修正してください。'
  ).setItalic(true);

  doc.saveAndClose();
  var docUrl = doc.getUrl();

  // Reports フォルダへ移動
  var reportsFolderId = getFolderIdFromSettings_('ReportsFolderID');
  if (reportsFolderId) {
    try {
      DriveApp.getFileById(doc.getId()).moveTo(DriveApp.getFolderById(reportsFolderId));
    } catch (e) {
      // フォルダ移動失敗は警告のみ（ドキュメントはマイドライブに残る）
    }
  }

  ui.alert(
    '✅ 四半期定期届出チェックリストを作成しました。\n\n'
    + '対象四半期: ' + quarterInfo.label + '\n'
    + '届出期限目安: ' + quarterInfo.deadline + '\n\n'
    + 'URL: ' + docUrl + '\n\n'
    + '※ 提出前に必ず出入国在留管理庁の最新様式・期限を確認してください。'
  );
}


// ============================================================
// 2. initComplianceSheet_() + createNotificationListSheet_()
//    「届出管理」シートの新規作成と初期化
// ============================================================
function initComplianceSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName(COMPLIANCE_SHEET_NAME);
  if (sheet) {
    // 既存シートがある場合は上書き確認せず再初期化（ヘッダのみ）
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // データ行はそのまま保持し、ヘッダのみ再整形
      styleHeaderRow_(sheet, 1, 8);
      autoResizeColumns_(sheet, 8);
      setComplianceConditionalFormats_(sheet);
      return sheet;
    }
    sheet.clearContents();
    sheet.clearFormats();
  } else {
    sheet = ss.insertSheet(COMPLIANCE_SHEET_NAME);
  }

  createNotificationListSheet_(sheet);
  return sheet;
}

// 届出管理シートのヘッダと書式を設定する内部関数
function createNotificationListSheet_(sheet) {
  var headers = [
    '届出ID',       // A: COMPLIANCE_COL.ID
    '届出種別',     // B: COMPLIANCE_COL.TYPE
    '対象',         // C: COMPLIANCE_COL.TARGET
    '提出期限',     // D: COMPLIANCE_COL.DEADLINE
    'ステータス',   // E: COMPLIANCE_COL.STATUS
    '提出日',       // F: COMPLIANCE_COL.SUBMIT_DATE
    '担当',         // G: COMPLIANCE_COL.ASSIGNEE
    '備考',         // H: COMPLIANCE_COL.NOTES
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, 1, headers.length);

  // ステータスのデータ検証
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      COMPLIANCE_STATUS.PENDING,
      COMPLIANCE_STATUS.SUBMITTED,
      COMPLIANCE_STATUS.OVERDUE,
    ], true).build();
  sheet.getRange('E2:E1000').setDataValidation(statusRule);

  // 提出期限の日付フォーマット
  sheet.getRange('D2:D1000').setNumberFormat('yyyy/MM/dd');
  // 提出日の日付フォーマット
  sheet.getRange('F2:F1000').setNumberFormat('yyyy/MM/dd');

  // 条件付き書式（期限超過=赤、14日以内=黄）
  setComplianceConditionalFormats_(sheet);

  autoResizeColumns_(sheet, headers.length);
  // 備考列は幅広めに固定
  sheet.setColumnWidth(8, 250);
  sheet.setFrozenRows(1);
}

// 条件付き書式を設定する内部関数
function setComplianceConditionalFormats_(sheet) {
  var deadlineRange = sheet.getRange('D2:D1000');
  var rules = [
    // 期限超過かつ未提出 → 赤
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(D2<>"",D2<TODAY(),E2="' + COMPLIANCE_STATUS.PENDING + '")')
      .setBackground(CONFIG.COLORS.ALERT_RED)
      .setFontColor('#ffffff')
      .setRanges([deadlineRange])
      .build(),
    // 14日以内かつ未提出 → 黄
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(D2<>"",D2>=TODAY(),D2<=TODAY()+14,E2="' + COMPLIANCE_STATUS.PENDING + '")')
      .setBackground(CONFIG.COLORS.ALERT_YELLOW)
      .setRanges([deadlineRange])
      .build(),
    // 提出済 → 薄緑
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=E2="' + COMPLIANCE_STATUS.SUBMITTED + '"')
      .setBackground('#e6f4ea')
      .setRanges([deadlineRange])
      .build(),
  ];
  sheet.setConditionalFormatRules(rules);
}


// ============================================================
// 3. generateComplianceSchedule()
//    当年度の四半期定期届出4回分を届出管理シートに自動投入
// ============================================================
// 【注意】提出期限はあくまで目安です。
// 最新の出入国在留管理庁の案内・通達を必ず確認してください。
function generateComplianceSchedule() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ui   = SpreadsheetApp.getUi();
  var now  = new Date();
  var year = now.getFullYear();

  // 届出管理シートを取得または初期化
  var sheet = ss.getSheetByName(COMPLIANCE_SHEET_NAME);
  if (!sheet) {
    sheet = initComplianceSheet_();
  }

  // 当年度の四半期定義（期間の終わり月, 届出期限の月）
  // 特定技能の定期届出: 各四半期末の翌々月末が目安（例: 1〜3月期→5月末頃）
  // ただし実際の期限は出入国在留管理庁の通知に従うこと
  var quarters = [
    { label: year + '年 第1四半期（1〜3月）',  deadline: new Date(year, 4, 31),  period: year + '/01/01 〜 ' + year + '/03/31' },
    { label: year + '年 第2四半期（4〜6月）',  deadline: new Date(year, 7, 31),  period: year + '/04/01 〜 ' + year + '/06/30' },
    { label: year + '年 第3四半期（7〜9月）',  deadline: new Date(year, 10, 30), period: year + '/07/01 〜 ' + year + '/09/30' },
    { label: year + '年 第4四半期（10〜12月）',deadline: new Date(year, 1, 28),  period: year + '/10/01 〜 ' + year + '/12/31' },
  ];
  // 第4四半期の届出期限は翌年2月末
  quarters[3].deadline = new Date(year + 1, 1, 28);

  // 既存データを読み込み、重複チェック用に届出IDセットを構築
  var existingIds = getExistingComplianceIds_(sheet);

  var addedCount = 0;

  quarters.forEach(function(q, idx) {
    // 登録支援機関分
    var rsoId = 'RSO-' + year + '-Q' + (idx + 1);
    if (!existingIds[rsoId]) {
      appendComplianceRow_(sheet, {
        id:       rsoId,
        type:     NOTIFICATION_TYPES.RSO_QUARTERLY,
        target:   CONFIG.COMPANY.NAME,
        deadline: q.deadline,
        status:   COMPLIANCE_STATUS.PENDING,
        notes:    q.period + ' / ※最新の様式・期限を要確認',
      });
      addedCount++;
    }

    // 受入機関分（各受入企業）
    var empId = 'EMP-' + year + '-Q' + (idx + 1);
    if (!existingIds[empId]) {
      appendComplianceRow_(sheet, {
        id:       empId,
        type:     NOTIFICATION_TYPES.EMPLOYER_QUARTERLY,
        target:   '全受入機関',
        deadline: q.deadline,
        status:   COMPLIANCE_STATUS.PENDING,
        notes:    q.period + ' / 受入機関各社に提出案内を実施 / ※最新の様式・期限を要確認',
      });
      addedCount++;
    }
  });

  SpreadsheetApp.flush();

  if (addedCount > 0) {
    ui.alert(
      '✅ 届出スケジュールを ' + addedCount + '件 追加しました。\n\n'
      + '対象年度: ' + year + '年\n\n'
      + '【注意】提出期限はあくまで目安です。\n'
      + '出入国在留管理庁の最新の通達を必ずご確認ください。'
    );
  } else {
    ui.alert('ℹ️ ' + year + '年度のスケジュールはすでに登録済みです。');
  }
}


// ============================================================
// 4. checkComplianceDeadlines()
//    届出管理シートを走査し、未提出・期限切れ等をアラート報告
// ============================================================
// 【注意】本関数はシート記録に基づく補助確認です。
// 最終的な届出状況は担当者が直接管理庁へ確認してください。
function checkComplianceDeadlines() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(COMPLIANCE_SHEET_NAME);

  if (!sheet || sheet.getLastRow() <= 1) {
    ui.alert(
      'ℹ️ 届出管理シートにデータがありません。\n\n'
      + '先に「届出スケジュール自動生成」を実行してください。'
    );
    return;
  }

  var today    = new Date();
  today.setHours(0, 0, 0, 0);
  var data     = sheet.getDataRange().getValues();

  var overdueItems  = [];
  var soonItems     = [];   // 14日以内
  var pendingItems  = [];   // 14日超・未提出
  var updatedRows   = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[COMPLIANCE_COL.ID - 1]) continue;

    var id       = row[COMPLIANCE_COL.ID - 1];
    var type     = row[COMPLIANCE_COL.TYPE - 1];
    var target   = row[COMPLIANCE_COL.TARGET - 1];
    var deadline = row[COMPLIANCE_COL.DEADLINE - 1] ? new Date(row[COMPLIANCE_COL.DEADLINE - 1]) : null;
    var status   = row[COMPLIANCE_COL.STATUS - 1];

    if (status === COMPLIANCE_STATUS.SUBMITTED) continue;
    if (!deadline) continue;

    deadline.setHours(0, 0, 0, 0);
    var daysLeft = Math.floor((deadline - today) / 86400000);

    if (daysLeft < 0) {
      // 期限超過 → ステータスを自動更新
      sheet.getRange(i + 1, COMPLIANCE_COL.STATUS).setValue(COMPLIANCE_STATUS.OVERDUE);
      overdueItems.push({ id: id, type: type, target: target, daysLeft: daysLeft, deadline: deadline });
    } else if (daysLeft <= 14) {
      soonItems.push({ id: id, type: type, target: target, daysLeft: daysLeft, deadline: deadline });
    } else {
      pendingItems.push({ id: id, type: type, target: target, daysLeft: daysLeft, deadline: deadline });
    }
  }

  SpreadsheetApp.flush();

  // アラートメッセージ組み立て
  var lines = ['=== 届出コンプライアンス確認レポート ==='];
  lines.push('確認日: ' + Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd'));
  lines.push('');

  if (overdueItems.length > 0) {
    lines.push('🔴 【期限超過 ― 至急対応が必要です】 ' + overdueItems.length + '件');
    overdueItems.forEach(function(item) {
      lines.push(
        '  ・' + item.id + ' ' + item.type
        + ' / 対象: ' + item.target
        + ' / 期限: ' + Utilities.formatDate(item.deadline, 'Asia/Tokyo', 'yyyy/MM/dd')
        + '（' + Math.abs(item.daysLeft) + '日超過）'
      );
    });
    lines.push('');
  }

  if (soonItems.length > 0) {
    lines.push('🟡 【14日以内に期限到来】 ' + soonItems.length + '件');
    soonItems.forEach(function(item) {
      lines.push(
        '  ・' + item.id + ' ' + item.type
        + ' / 対象: ' + item.target
        + ' / 期限: ' + Utilities.formatDate(item.deadline, 'Asia/Tokyo', 'yyyy/MM/dd')
        + '（あと' + item.daysLeft + '日）'
      );
    });
    lines.push('');
  }

  if (pendingItems.length > 0) {
    lines.push('🟢 【未提出（余裕あり）】 ' + pendingItems.length + '件');
    pendingItems.forEach(function(item) {
      lines.push(
        '  ・' + item.id + ' ' + item.type
        + ' / 期限: ' + Utilities.formatDate(item.deadline, 'Asia/Tokyo', 'yyyy/MM/dd')
        + '（あと' + item.daysLeft + '日）'
      );
    });
    lines.push('');
  }

  if (overdueItems.length === 0 && soonItems.length === 0 && pendingItems.length === 0) {
    lines.push('✅ 期限が迫っている届出・超過している届出はありません。');
  }

  lines.push('');
  lines.push('※ 本報告はシート記録に基づく補助確認です。');
  lines.push('  最終的な届出状況は担当者が直接ご確認ください。');

  var reportText = lines.join('\n');

  // UIアラート
  ui.alert(reportText);

  // メール通知（期限超過または14日以内のものがある場合のみ）
  if (overdueItems.length > 0 || soonItems.length > 0) {
    var subject = '【重要】特定技能 届出コンプライアンス確認: ';
    if (overdueItems.length > 0) {
      subject += '期限超過 ' + overdueItems.length + '件';
      if (soonItems.length > 0) subject += '、';
    }
    if (soonItems.length > 0) {
      subject += '期限間近 ' + soonItems.length + '件';
    }
    sendNotificationEmail_(subject, reportText);
  }
}


// ============================================================
// 5. createAdHocNotificationGuideDoc()
//    随時届出が必要なケースのガイドDoc生成
// ============================================================
// 【注意】随時届出の要否・様式・提出期限は制度改正で変更されることがあります。
// 必ず最新の出入国在留管理庁の通達・様式を確認してください。
function createAdHocNotificationGuideDoc() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var now = new Date();

  var title = '特定技能_随時届出ガイド_'
    + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
  var doc  = DocumentApp.create(title);
  var body = doc.getBody();
  body.clear();

  // タイトル
  var titlePara = body.appendParagraph('特定技能 随時届出 判断ガイド');
  titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  titlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph(
    Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日') + ' 作成 ／ '
    + CONFIG.COMPANY.NAME + ' ／ ' + CONFIG.COMPANY.REGISTRATION_NUMBER
  ).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendHorizontalRule();

  // 注意事項
  body.appendParagraph('【重要な注意事項】').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(
    '本書は随時届出の判断を補助するための内部参考資料です。'
    + '届出の要否・様式・提出期限・提出先は出入国在留管理庁による制度改正・通達によって変更されることがあります。'
    + '実際の届出に際しては、必ず最新の出入国在留管理庁のウェブサイト・管轄地方出入国在留管理局に確認してください。'
    + '\n参考URL: https://www.moj.go.jp/isa/'
  ).setItalic(true);

  body.appendParagraph('');

  // 随時届出が必要なケース一覧
  body.appendParagraph('■ 随時届出が必要な主なケース')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(
    '以下のいずれかの事象が発生した場合、速やかに届出の要否を確認し、'
    + '必要に応じて出入国在留管理庁（管轄局）へ届け出てください。'
    + '\n（「〇日以内」等の期限は制度改正により変更される場合があります。最新の通達を必ず確認してください。）'
  ).setItalic(true);

  // 各随時届出ケースの定義
  var adHocCases = [
    {
      title: '① 雇用契約に関する変更（受入機関が届出）',
      items: [
        '・ 雇用契約の変更（賃金・労働時間・業務内容等）',
        '・ 雇用契約の終了（退職・解雇・雇用期間満了等）',
        '・ 新たな雇用契約の締結',
      ],
      timing: '事由発生から14日以内が目安（※最新の案内を要確認）',
      form:   '参考様式第3-1号〜第3-3号相当（※最新様式を要確認）',
    },
    {
      title: '② 支援委託契約に関する変更（登録支援機関が届出）',
      items: [
        '・ 受入機関との支援委託契約の変更',
        '・ 支援委託契約の終了',
        '・ 新たな支援委託契約の締結',
      ],
      timing: '事由発生から14日以内が目安（※最新の案内を要確認）',
      form:   '参考様式第3-9号・第3-10号相当（※最新様式を要確認）',
    },
    {
      title: '③ 1号特定技能外国人支援計画の変更（受入機関／登録支援機関が届出）',
      items: [
        '・ 支援計画の内容の変更（支援内容・担当者等）',
        '・ 登録支援機関への支援の全部委託→一部委託への変更、またはその逆',
        '・ 委託先の登録支援機関の変更',
      ],
      timing: '変更前または事由発生後速やかに（※最新の案内を要確認）',
      form:   '参考様式第3-4号・第3-5号相当（※最新様式を要確認）',
    },
    {
      title: '④ 特定技能外国人の受入れが困難となった場合（受入機関が届出）',
      items: [
        '・ 経営悪化・倒産等により雇用継続が困難になった',
        '・ 外国人が重大な規律違反を行った',
        '・ 天災等の不可抗力による受入継続困難',
      ],
      timing: '事由発生から14日以内が目安（※最新の案内を要確認）',
      form:   '参考様式第3-6号相当（※最新様式を要確認）',
      note:   '→ 外国人への転職支援（支援⑧）が義務付けられる場合あり。転職先の紹介・ハローワークへの同行等を速やかに実施してください。',
    },
    {
      title: '⑤ 特定技能外国人が行方不明になった場合（受入機関・登録支援機関が届出）',
      items: [
        '・ 外国人が無断欠勤・連絡不通となった',
        '・ 住居に戻らなくなった',
      ],
      timing: '判明後速やかに（遅くとも14日以内が目安）（※最新の案内を要確認）',
      form:   '参考様式第3-6号相当（※最新様式を要確認）',
      note:   '→ 警察への相談・捜索届、管轄局への連絡を速やかに行ってください。',
    },
    {
      title: '⑥ 受入機関に関する変更事項（受入機関が届出）',
      items: [
        '・ 商号または名称の変更',
        '・ 代表者の変更',
        '・ 主たる事務所の所在地の変更',
        '・ 事業の廃止・解散',
      ],
      timing: '事由発生から14日以内が目安（※最新の案内を要確認）',
      form:   '参考様式第3-2号相当（※最新様式を要確認）',
    },
    {
      title: '⑦ 登録支援機関に関する変更事項（登録支援機関が届出）',
      items: [
        '・ 商号・名称・住所の変更',
        '・ 代表者の変更',
        '・ 支援担当者の変更',
        '・ 事業の廃止・解散',
      ],
      timing: '事由発生から14日以内が目安（※最新の案内を要確認）',
      form:   '登録支援機関変更届出書（※最新様式を要確認）',
    },
    {
      title: '⑧ 受入機関に係る軽微な変更以外の変更（随時）',
      items: [
        '・ 外国人技能実習機構への届出が必要な変更事項',
        '・ 労働法令違反等が判明した場合',
      ],
      timing: '各事由ごとの規定に従う（※最新の案内を要確認）',
      form:   '各事由に対応した様式（※最新様式を要確認）',
    },
  ];

  adHocCases.forEach(function(c) {
    body.appendParagraph(c.title).setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph('【該当する事象】');
    c.items.forEach(function(item) { body.appendParagraph(item); });
    body.appendParagraph('【届出期限の目安】' + c.timing);
    body.appendParagraph('【主な届出様式】' + c.form);
    if (c.note) {
      body.appendParagraph('【対応ポイント】' + c.note).setItalic(true);
    }
    body.appendParagraph('');
  });

  body.appendHorizontalRule();

  // 対応フロー
  body.appendParagraph('■ 随時届出の対応フロー（参考）')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  var flowSteps = [
    'STEP 1: 事由発生を把握したら、上記一覧で届出の要否を確認する',
    'STEP 2: 出入国在留管理庁のウェブサイトから最新の様式をダウンロードする',
    'STEP 3: 受入機関と連携して届出書類を作成する',
    'STEP 4: 電子届出システム（OICS）または管轄地方出入国在留管理局へ提出する',
    'STEP 5: 受付番号・提出日を「届出管理」シートに記録する（ステータスを「提出済」へ変更）',
    'STEP 6: 提出書類の写しをDriveの報告書フォルダへ保存する',
  ];
  flowSteps.forEach(function(s) { body.appendParagraph(s); });

  body.appendParagraph('');

  // 連絡先
  body.appendParagraph('■ 主な問い合わせ先').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendTable([
    ['出入国在留管理庁（入管庁）', 'https://www.moj.go.jp/isa/'],
    ['電子届出システム（OICS）',   'https://oics.moj.go.jp/'],
    ['自社相談窓口',               CONFIG.COMPANY.PHONE + ' ／ ' + CONFIG.COMPANY.EMAIL],
  ]).setBorderColor('#dadce0');

  body.appendParagraph('');
  body.appendParagraph(
    '本書は ' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日')
    + ' 時点の情報を基に作成されています。'
    + '制度改正により内容が変わる場合がありますので、定期的に最新情報を確認し内容を更新してください。'
  ).setItalic(true);

  doc.saveAndClose();
  var docUrl = doc.getUrl();

  // Reports フォルダへ移動
  var reportsFolderId = getFolderIdFromSettings_('ReportsFolderID');
  if (reportsFolderId) {
    try {
      DriveApp.getFileById(doc.getId()).moveTo(DriveApp.getFolderById(reportsFolderId));
    } catch (e) {
      // フォルダ移動失敗は警告のみ
    }
  }

  ui.alert(
    '✅ 随時届出ガイドを作成しました。\n\n'
    + 'URL: ' + docUrl + '\n\n'
    + '※ 本書は参考資料です。実際の届出前に必ず最新の様式・期限をご確認ください。'
  );
}


// ============================================================
// 内部ユーティリティ関数
// ============================================================

/**
 * 現在日から直近の四半期情報を算出する
 * @param {Date} now 基準日
 * @returns {{label:string, period:string, deadline:string, qNum:number, year:number}}
 */
function getCurrentQuarterInfo_(now) {
  var year  = now.getFullYear();
  var month = now.getMonth() + 1; // 1〜12

  var qNum, startMonth, endMonth, deadlineMonth, deadlineYear;

  if (month <= 3) {
    qNum = 1; startMonth = 1;  endMonth = 3;  deadlineMonth = 5;  deadlineYear = year;
  } else if (month <= 6) {
    qNum = 2; startMonth = 4;  endMonth = 6;  deadlineMonth = 8;  deadlineYear = year;
  } else if (month <= 9) {
    qNum = 3; startMonth = 7;  endMonth = 9;  deadlineMonth = 11; deadlineYear = year;
  } else {
    qNum = 4; startMonth = 10; endMonth = 12; deadlineMonth = 2;  deadlineYear = year + 1;
  }

  // 月末日を求めるユーティリティ（翌月0日 = 当月末）
  var deadlineDateObj = new Date(deadlineYear, deadlineMonth, 0); // 月末日
  var deadlineStr     = Utilities.formatDate(deadlineDateObj, 'Asia/Tokyo', 'yyyy/MM/dd');

  var startStr = year + '/' + pad2_(startMonth) + '/01';
  var endStr   = year + '/' + pad2_(endMonth)   + '/' + getLastDayOfMonth_(year, endMonth);

  return {
    label:    year + '年 第' + qNum + '四半期（' + startMonth + '〜' + endMonth + '月）',
    period:   startStr + ' 〜 ' + endStr,
    deadline: deadlineStr + '頃',
    qNum:     qNum,
    year:     year,
  };
}

/**
 * 外国人管理シート・定期面談記録シートから在籍者の支援状況を集計する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {{rows: Array, activeCount: number, interviewedCount: number}}
 */
function collectWorkerSummary_(ss) {
  var workersSheet    = ss.getSheetByName(CONFIG.SHEETS.WORKERS);
  var interviewsSheet = ss.getSheetByName(CONFIG.SHEETS.INTERVIEWS);

  if (!workersSheet) {
    return { rows: [], activeCount: 0, interviewedCount: 0 };
  }

  // 面談記録を外国人IDでインデックス化（最新実施日を保持）
  var interviewMap = {};
  if (interviewsSheet && interviewsSheet.getLastRow() > 1) {
    var ivData = interviewsSheet.getDataRange().getValues();
    for (var i = 1; i < ivData.length; i++) {
      var ivRow = ivData[i];
      var wId   = ivRow[1]; // 外国人ID（列B）
      var status = ivRow[5]; // ステータス（列F）
      var ivDate = ivRow[6]; // 面談実施日（列G）
      var nextDate = ivRow[15]; // 次回面談予定（列P）
      if (!wId) continue;
      if (status === '実施済' && ivDate) {
        if (!interviewMap[wId] || new Date(ivDate) > new Date(interviewMap[wId].lastDate)) {
          interviewMap[wId] = {
            lastDate: ivDate,
            nextDate: nextDate || '',
          };
        }
      }
    }
  }

  var rows = [];
  var activeCount = 0;
  var interviewedCount = 0;

  if (workersSheet.getLastRow() <= 1) {
    return { rows: rows, activeCount: 0, interviewedCount: 0 };
  }

  var wData = workersSheet.getDataRange().getValues();
  for (var j = 1; j < wData.length; j++) {
    var wRow = wData[j];
    if (!wRow[0]) continue; // IDが空ならスキップ

    var wId2    = wRow[0];  // ID（列A）
    var name    = wRow[1];  // 氏名（列B）
    var company = wRow[10]; // 受入企業名（列K）
    var wStatus = wRow[3];  // ステータス（列D）

    if (wStatus !== '在籍中') continue;
    activeCount++;

    var ivInfo   = interviewMap[wId2] || null;
    var lastDate = ivInfo ? Utilities.formatDate(new Date(ivInfo.lastDate), 'Asia/Tokyo', 'yyyy/MM/dd') : '記録なし';
    var nextDate = (ivInfo && ivInfo.nextDate)
      ? Utilities.formatDate(new Date(ivInfo.nextDate), 'Asia/Tokyo', 'yyyy/MM/dd')
      : '';

    if (ivInfo) interviewedCount++;

    rows.push([
      name    || '',
      company || '',
      wStatus,
      lastDate,
      nextDate,
      '',  // 備考（手動記入欄）
    ]);
  }

  return { rows: rows, activeCount: activeCount, interviewedCount: interviewedCount };
}

/**
 * 届出管理シートの既存届出IDをオブジェクトとして返す
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object} idをキーとするオブジェクト
 */
function getExistingComplianceIds_(sheet) {
  var result = {};
  if (sheet.getLastRow() <= 1) return result;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var id = data[i][COMPLIANCE_COL.ID - 1];
    if (id) result[id] = true;
  }
  return result;
}

/**
 * 届出管理シートに1行追加する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {{id:string, type:string, target:string, deadline:Date, status:string, assignee:string, notes:string}} params
 */
function appendComplianceRow_(sheet, params) {
  var nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, COMPLIANCE_COL.ID).setValue(params.id || '');
  sheet.getRange(nextRow, COMPLIANCE_COL.TYPE).setValue(params.type || '');
  sheet.getRange(nextRow, COMPLIANCE_COL.TARGET).setValue(params.target || '');
  if (params.deadline) {
    sheet.getRange(nextRow, COMPLIANCE_COL.DEADLINE)
      .setValue(params.deadline)
      .setNumberFormat('yyyy/MM/dd');
  }
  sheet.getRange(nextRow, COMPLIANCE_COL.STATUS).setValue(params.status || COMPLIANCE_STATUS.PENDING);
  sheet.getRange(nextRow, COMPLIANCE_COL.SUBMIT_DATE).setValue(params.submitDate || '');
  sheet.getRange(nextRow, COMPLIANCE_COL.ASSIGNEE).setValue(params.assignee || '');
  sheet.getRange(nextRow, COMPLIANCE_COL.NOTES).setValue(params.notes || '');
}

/**
 * 2桁ゼロパディング
 * @param {number} n
 * @returns {string}
 */
function pad2_(n) {
  return n < 10 ? '0' + n : '' + n;
}

/**
 * 指定年月の最終日を返す
 * @param {number} year
 * @param {number} month 1〜12
 * @returns {number}
 */
function getLastDayOfMonth_(year, month) {
  return new Date(year, month, 0).getDate();
}


// ============================================================
// ファイル末尾: Menu.gs 統合時の追加項目メモ
// ============================================================
//
// 【Menu.gs へ追加を提案するメニュー項目】
//
// ─── 行政コンプライアンス ───
//  '📋 届出管理シート初期化'          → initComplianceSheet_
//  '📅 届出スケジュール自動生成'       → generateComplianceSchedule
//  '✅ 届出期限チェック'              → checkComplianceDeadlines
//  '📄 四半期届出チェックリスト作成'   → createQuarterlyReportChecklistDoc
//  '📑 随時届出ガイド作成'            → createAdHocNotificationGuideDoc
//
// 【Config.gs の SHEETS オブジェクトへの追加を提案する定数】
//
//   COMPLIANCE: 'コンプライアンス届出管理',
//
// 上記を Config.gs の SHEETS ブロックへ追加すると、
// 本ファイルの COMPLIANCE_SHEET_NAME 変数を
// CONFIG.SHEETS.COMPLIANCE に置き換えられます。
