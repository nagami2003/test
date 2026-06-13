// ===== 定期面談管理 =====

function showInterviewDialog() {
  const html = HtmlService.createHtmlOutput(getInterviewFormHtml_())
    .setWidth(600).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '定期面談 記録');
}

function recordInterview(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.INTERVIEWS);
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

  // 既存の予定行を更新、または新規追加
  const data = sheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === formData.workerId && data[i][5] === '未実施') {
      // 予定日が最も近い未実施を更新
      if (targetRow === -1) targetRow = i + 1;
    }
  }

  if (targetRow === -1) {
    // 新規行追加
    targetRow = sheet.getLastRow() + 1;
    const interviewId = 'I' + String(targetRow).padStart(5, '0');
    sheet.getRange(targetRow, 1).setValue(interviewId);
    sheet.getRange(targetRow, 2).setValue(formData.workerId);
    sheet.getRange(targetRow, 3).setValue(formData.workerName);
    sheet.getRange(targetRow, 4).setValue(formData.companyName);
  }

  sheet.getRange(targetRow, 5).setValue(formData.scheduledDate);
  sheet.getRange(targetRow, 6).setValue('実施済');
  sheet.getRange(targetRow, 7).setValue(today);
  sheet.getRange(targetRow, 8).setValue(formData.interviewer);
  sheet.getRange(targetRow, 9).setValue(formData.method);
  sheet.getRange(targetRow, 10).setValue(formData.location);
  sheet.getRange(targetRow, 11).setValue(formData.workStatus);
  sheet.getRange(targetRow, 12).setValue(formData.lifeStatus);
  sheet.getRange(targetRow, 13).setValue(formData.healthStatus);
  sheet.getRange(targetRow, 14).setValue(formData.consultationContent);
  sheet.getRange(targetRow, 15).setValue(formData.responseContent);
  sheet.getRange(targetRow, 16).setValue(formData.nextSchedule);
  sheet.getRange(targetRow, 18).setValue(formData.notes);

  // 面談記録 Google Doc を作成
  const docUrl = createInterviewDoc_(formData, today);
  if (docUrl) sheet.getRange(targetRow, 17).setValue(docUrl);

  // 次回面談をカレンダーに追加
  if (formData.nextSchedule) {
    addInterviewToCalendar_(formData.workerId, formData.workerName, formData.companyName, formData.nextSchedule);

    // 次回予定を新規行として追加
    const nextRow = sheet.getLastRow() + 1;
    const nextId = 'I' + String(nextRow).padStart(5, '0');
    sheet.getRange(nextRow, 1, 1, 6).setValues([[
      nextId, formData.workerId, formData.workerName, formData.companyName,
      formData.nextSchedule, '未実施',
    ]]);
  }

  return { success: true };
}

function scheduleInterviews() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workersSheet = ss.getSheetByName(CONFIG.SHEETS.WORKERS);
  const data = workersSheet.getDataRange().getValues();
  let count = 0;

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) break;
    if (data[i][3] !== '在籍中') continue;
    const workerId = data[i][0];
    const name = data[i][1];
    const companyName = data[i][10];
    const workStart = data[i][16];
    if (!workStart) continue;

    scheduleQuarterlyInterviews_(workerId, name, workStart);
    count++;
  }

  SpreadsheetApp.getUi().alert(`✅ ${count}名分の定期面談スケジュールをカレンダーに追加しました。`);
}

function checkPendingInterviews() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.INTERVIEWS);
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  const overdue = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) break;
    if (data[i][5] !== '未実施') continue;
    const scheduled = data[i][4] ? new Date(data[i][4]) : null;
    if (scheduled && scheduled <= today) {
      overdue.push({ name: data[i][2], company: data[i][3], date: data[i][4] });
    }
  }

  if (overdue.length === 0) {
    SpreadsheetApp.getUi().alert('✅ 期限切れの未実施面談はありません。');
    return;
  }

  const message = overdue.map(o =>
    `${o.name}（${o.company}）- 予定日: ${o.date}`
  ).join('\n');
  SpreadsheetApp.getUi().alert(`⚠️ 未実施の面談が${overdue.length}件あります。\n\n${message}`);
}

function addInterviewToCalendar_(workerId, workerName, companyName, dateStr) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const date = new Date(dateStr);
    const endDate = new Date(date.getTime() + 60 * 60 * 1000); // 1時間
    calendar.createEvent(
      `【定期面談】${workerName}（${companyName}）`,
      date,
      endDate,
      {
        description: `特定技能外国人 定期面談\nID: ${workerId}\n受入企業: ${companyName}\n\n※面談後は「定期面談記録」シートに記録してください。`,
        color: CalendarApp.EventColor.BLUE,
      }
    );
  } catch (e) {
    console.log('カレンダー登録エラー: ' + e.message);
  }
}

function createInterviewDoc_(formData, interviewDate) {
  try {
    const doc = DocumentApp.create(
      `面談記録_${formData.workerId}_${formData.workerName}_${interviewDate.replace(/\//g, '')}`
    );
    const body = doc.getBody();
    body.clear();

    const title = body.appendParagraph('定期面談記録');
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    body.appendParagraph(`面談実施日: ${interviewDate}`);
    body.appendParagraph(`面談者: ${formData.interviewer}`);
    body.appendParagraph(`面談方法: ${formData.method}`);
    body.appendHorizontalRule();

    body.appendParagraph('■ 外国人情報').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendTable([
      ['氏名', formData.workerName],
      ['受入企業', formData.companyName],
      ['外国人ID', formData.workerId],
    ]);

    body.appendParagraph('■ 面談内容').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendTable([
      ['就労状況', formData.workStatus],
      ['生活状況', formData.lifeStatus],
      ['健康状態', formData.healthStatus],
    ]);

    if (formData.consultationContent) {
      body.appendParagraph('■ 相談・苦情内容').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(formData.consultationContent);
    }

    if (formData.responseContent) {
      body.appendParagraph('■ 対応内容').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(formData.responseContent);
    }

    body.appendParagraph('■ 次回面談予定日').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(formData.nextSchedule || '未定');

    body.appendHorizontalRule();
    body.appendParagraph('').appendText('担当者署名: ______________________');
    body.appendParagraph(`記録作成日: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日')}`);

    doc.saveAndClose();

    // フォルダに移動
    const workersFolderId = getFolderIdFromSettings_('WorkersFolderID');
    if (workersFolderId) {
      const workersFolder = DriveApp.getFolderById(workersFolderId);
      const workerFolders = workersFolder.getFoldersByName(`${formData.workerId}_${formData.workerName}`);
      if (workerFolders.hasNext()) {
        const workerFolder = workerFolders.next();
        const recordFolders = workerFolder.getFoldersByName('面談記録');
        if (recordFolders.hasNext()) {
          DriveApp.getFileById(doc.getId()).moveTo(recordFolders.next());
        }
      }
    }

    return doc.getUrl();
  } catch (e) {
    console.log('面談記録作成エラー: ' + e.message);
    return '';
  }
}

function getInterviewFormHtml_() {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const threeMonthsLater = Utilities.formatDate(
    new Date(Date.now() + 90 * 24 * 60 * 60 * 1000), 'Asia/Tokyo', 'yyyy-MM-dd'
  );

  return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Noto Sans JP', sans-serif; padding: 15px; font-size: 13px; }
    label { display: block; margin-top: 8px; font-weight: bold; color: #333; }
    input, select, textarea { width: 100%; padding: 5px; margin-top: 2px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
    .row { display: flex; gap: 8px; }
    .row > div { flex: 1; }
    button { margin-top: 12px; padding: 10px 20px; background: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; width: 100%; }
    h3 { color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 4px; margin-top: 0; }
    .section { background: #f8f9fa; padding: 10px; border-radius: 4px; margin-top: 10px; }
  </style>
</head>
<body>
  <h3>定期面談 記録</h3>
  <div class="row">
    <div>
      <label>外国人ID*</label>
      <input type="text" id="workerId" required placeholder="W0001">
    </div>
    <div>
      <label>氏名*</label>
      <input type="text" id="workerName" required>
    </div>
  </div>
  <div class="row">
    <div>
      <label>受入企業名</label>
      <input type="text" id="companyName">
    </div>
    <div>
      <label>面談予定日</label>
      <input type="date" id="scheduledDate" value="${today}">
    </div>
  </div>
  <div class="row">
    <div>
      <label>面談者（担当者名）*</label>
      <input type="text" id="interviewer" required>
    </div>
    <div>
      <label>面談方法*</label>
      <select id="method" required>
        <option value="対面">対面</option>
        <option value="オンライン（Google Meet）">オンライン（Google Meet）</option>
        <option value="電話">電話</option>
        <option value="書面">書面</option>
      </select>
    </div>
  </div>
  <label>面談場所</label>
  <input type="text" id="location" placeholder="会社会議室 / オンライン等">

  <div class="section">
    <b>面談内容</b>
    <label>就労状況</label>
    <select id="workStatus">
      <option value="良好">良好</option>
      <option value="概ね良好">概ね良好</option>
      <option value="要確認">要確認</option>
      <option value="問題あり">問題あり</option>
    </select>
    <label>生活状況</label>
    <select id="lifeStatus">
      <option value="良好">良好</option>
      <option value="概ね良好">概ね良好</option>
      <option value="要確認">要確認</option>
      <option value="問題あり">問題あり</option>
    </select>
    <label>健康状態</label>
    <select id="healthStatus">
      <option value="良好">良好</option>
      <option value="概ね良好">概ね良好</option>
      <option value="要確認">要確認</option>
      <option value="要医療">要医療</option>
    </select>
  </div>

  <label>相談・苦情内容（あれば記載）</label>
  <textarea id="consultationContent" rows="2" placeholder="特になし"></textarea>
  <label>対応内容</label>
  <textarea id="responseContent" rows="2"></textarea>
  <div class="row">
    <div>
      <label>次回面談予定日</label>
      <input type="date" id="nextSchedule" value="${threeMonthsLater}">
    </div>
    <div>
      <label>備考</label>
      <input type="text" id="notes">
    </div>
  </div>
  <button type="button" onclick="submitForm()">記録する</button>
  <script>
    function submitForm() {
      const formData = {
        workerId: document.getElementById('workerId').value,
        workerName: document.getElementById('workerName').value,
        companyName: document.getElementById('companyName').value,
        scheduledDate: document.getElementById('scheduledDate').value,
        interviewer: document.getElementById('interviewer').value,
        method: document.getElementById('method').value,
        location: document.getElementById('location').value,
        workStatus: document.getElementById('workStatus').value,
        lifeStatus: document.getElementById('lifeStatus').value,
        healthStatus: document.getElementById('healthStatus').value,
        consultationContent: document.getElementById('consultationContent').value,
        responseContent: document.getElementById('responseContent').value,
        nextSchedule: document.getElementById('nextSchedule').value,
        notes: document.getElementById('notes').value,
      };
      if (!formData.workerId || !formData.workerName || !formData.interviewer) {
        alert('必須項目（*）を入力してください');
        return;
      }
      google.script.run
        .withSuccessHandler(() => {
          alert('✅ 面談記録を保存しました。');
          google.script.host.close();
        })
        .withFailureHandler(err => alert('エラー: ' + err.message))
        .recordInterview(formData);
    }
  </script>
</body>
</html>`;
}
