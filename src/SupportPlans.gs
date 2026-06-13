// ===== 支援計画管理 =====

function createSupportPlanDoc() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.WORKERS);
  const activeCell = sheet.getActiveCell();
  const row = activeCell.getRow();

  if (row <= 1) {
    SpreadsheetApp.getUi().alert('外国人管理シートで、対象の外国人の行を選択してから実行してください。');
    return;
  }

  const data = sheet.getRange(row, 1, 1, 20).getValues()[0];
  const workerId = data[0];
  const nameRoman = data[1];
  const companyName = data[10];
  const sector = data[11];
  const residenceStatus = data[12];
  const visaExpiry = data[13] ? Utilities.formatDate(new Date(data[13]), 'Asia/Tokyo', 'yyyy年MM月dd日') : '';
  const workStart = data[16] ? Utilities.formatDate(new Date(data[16]), 'Asia/Tokyo', 'yyyy年MM月dd日') : '';

  // テンプレートから Google Doc を作成
  const doc = DocumentApp.create(`支援計画書_${workerId}_${nameRoman}`);
  const body = doc.getBody();

  body.clear();

  const title = body.appendParagraph('特定技能外国人支援計画書');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('').appendText(`作成日: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日')}`);
  body.appendParagraph(`支援機関名: ${CONFIG.COMPANY.NAME}`);
  body.appendParagraph(`登録番号: ${CONFIG.COMPANY.REGISTRATION_NUMBER}`);

  body.appendHorizontalRule();

  body.appendParagraph('1. 外国人の情報').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const workerTable = body.appendTable([
    ['氏名（ローマ字）', nameRoman],
    ['受入企業名', companyName],
    ['分野', sector],
    ['在留資格', residenceStatus],
    ['在留期限', visaExpiry],
    ['就労開始予定日', workStart],
  ]);
  workerTable.setBorderColor('#dadce0');

  body.appendParagraph('2. 支援内容').setHeading(DocumentApp.ParagraphHeading.HEADING2);

  CONFIG.SUPPORT_TYPES.forEach((type, i) => {
    body.appendParagraph(type).setHeading(DocumentApp.ParagraphHeading.HEADING3);
    const details = getSupportTypeDetails_(i);
    body.appendParagraph(details.description);
    body.appendParagraph(`実施時期: ${details.timing}`);
    body.appendParagraph(`実施方法: ${details.method}`);
    body.appendParagraph('');
  });

  body.appendHorizontalRule();
  body.appendParagraph('3. 支援体制').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(`担当支援員: （記入）`);
  body.appendParagraph(`連絡先: ${CONFIG.COMPANY.PHONE}`);
  body.appendParagraph(`相談窓口: ${CONFIG.COMPANY.EMAIL}`);

  body.appendParagraph('').appendParagraph('');
  body.appendParagraph('署名欄');
  const signTable = body.appendTable([
    ['支援機関代表者', '', '外国人本人（署名）', ''],
    ['', '', '', ''],
    ['日付: 　　　　年　　月　　日', '', '日付: 　　　　年　　月　　日', ''],
  ]);

  doc.saveAndClose();
  const docUrl = doc.getUrl();

  // フォルダに移動
  const workersFolderId = getFolderIdFromSettings_('WorkersFolderID');
  if (workersFolderId) {
    const workersFolder = DriveApp.getFolderById(workersFolderId);
    const workerFolders = workersFolder.getFoldersByName(`${workerId}_${nameRoman}`);
    if (workerFolders.hasNext()) {
      const workerFolder = workerFolders.next();
      const planFolders = workerFolder.getFoldersByName('支援計画');
      if (planFolders.hasNext()) {
        DriveApp.getFileById(doc.getId()).moveTo(planFolders.next());
      }
    }
  }

  // シートにURL記録
  sheet.getRange(row, 18).setValue(docUrl);

  const plansSheet = ss.getSheetByName(CONFIG.SHEETS.SUPPORT_PLANS);
  if (plansSheet) {
    const plansData = plansSheet.getDataRange().getValues();
    for (let i = 1; i < plansData.length; i++) {
      if (plansData[i][1] === workerId) {
        plansSheet.getRange(i + 1, 6).setValue(docUrl);
        break;
      }
    }
  }

  SpreadsheetApp.getUi().alert(`✅ 支援計画書を作成しました。\n\nURL: ${docUrl}`);
}

function updateSupportProgress() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const plansSheet = ss.getSheetByName(CONFIG.SHEETS.SUPPORT_PLANS);
  if (!plansSheet) return;

  const data = plansSheet.getDataRange().getValues();
  const baseColCount = 6;

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) break;
    const typeCount = CONFIG.SUPPORT_TYPES.length;
    let implemented = 0;
    for (let j = baseColCount; j < baseColCount + typeCount; j++) {
      if (data[i][j] === '実施済') implemented++;
    }
    const completionRate = Math.round((implemented / typeCount) * 100) + '%';
    plansSheet.getRange(i + 1, baseColCount + typeCount + 1).setValue(completionRate);
  }

  SpreadsheetApp.getUi().alert('✅ 支援進捗を更新しました。');
}

function getSupportTypeDetails_(index) {
  const details = [
    { description: '就労・生活に関する情報をあらかじめ提供する', timing: '入国前・入国直後', method: '対面または書面' },
    { description: '入国時および帰国時の空港等への送迎を行う', timing: '入国時・帰国時', method: '自動車による送迎' },
    { description: '住居の確保および各種生活インフラの契約支援を行う', timing: '入国後速やかに', method: '同行支援' },
    { description: '日本での生活に必要な情報（公的手続き、緊急連絡先等）を提供する', timing: '入国後速やかに', method: '対面（8時間以上）' },
    { description: '日本語学習の機会を定期的に提供する', timing: '就労期間中継続的に', method: '日本語教室、eラーニング等' },
    { description: '相談・苦情を受け付け、適切に対応する', timing: '随時（24時間対応窓口設置）', method: '電話・メール・面談' },
    { description: '日本人との交流を促進する機会を設ける', timing: '年1回以上', method: '地域行事参加、交流会開催等' },
    { description: '受入困難時の転職支援を行う', timing: '必要時', method: '求職活動支援、ハローワーク同行' },
    { description: '定期的に面談を実施し、問題があれば行政機関に通報する', timing: '3ヶ月に1回以上', method: '対面面談（オンライン可）' },
    { description: '生活に必要な住居の確保を支援する', timing: '入国前・入国後', method: '物件探し、契約同行' },
    { description: '医療機関・交通機関・生活インフラの案内を行う', timing: '入国後速やかに', method: '書面・同行' },
    { description: '行政機関への各種届出・手続きを支援する', timing: '必要時', method: '同行または代行（適法な範囲内）' },
  ];
  return details[index] || { description: '', timing: '', method: '' };
}
