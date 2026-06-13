// ===== 外国人管理 =====

function showAddWorkerDialog() {
  const html = HtmlService.createHtmlOutput(getAddWorkerFormHtml_())
    .setWidth(600).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, '外国人 新規登録');
}

function addWorker(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.WORKERS);
  const lastRow = sheet.getLastRow();
  const newId = 'W' + String(lastRow).padStart(4, '0');

  // 個人フォルダを作成
  const workersFolderId = getFolderIdFromSettings_('WorkersFolderID');
  let folderUrl = '';
  if (workersFolderId) {
    const workersFolder = DriveApp.getFolderById(workersFolderId);
    const workerFolder = workersFolder.createFolder(`${newId}_${formData.nameRoman}`);
    folderUrl = workerFolder.getUrl();
    // サブフォルダ作成
    ['在留関係書類', '支援計画', '面談記録', '契約書'].forEach(sub => workerFolder.createFolder(sub));
  }

  const row = [
    newId, formData.nameRoman, formData.nameNative, '在籍中', formData.nationality,
    formData.birthDate, formData.gender, formData.phone, formData.email,
    formData.companyId, formData.companyName, formData.sector,
    formData.residenceStatus, formData.visaExpiry, formData.cardNumber,
    formData.entryDate, formData.workStartDate, '', folderUrl, formData.notes,
  ];

  sheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);

  // カレンダーに定期面談を自動登録
  if (formData.workStartDate) scheduleQuarterlyInterviews_(newId, formData.nameRoman, formData.workStartDate);

  // 支援計画シートに初期行を追加
  addSupportPlanRow_(newId, formData.nameRoman, formData.companyName);

  return { success: true, id: newId };
}

function checkVisaExpiry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.WORKERS);
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  const alerts = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) break;
    if (data[i][3] !== '在籍中') continue;
    const expiry = data[i][13];
    if (!expiry) continue;
    const daysLeft = Math.floor((new Date(expiry) - today) / 86400000);
    if (daysLeft <= 90) {
      alerts.push({ name: data[i][1], daysLeft, expiry });
    }
  }

  if (alerts.length === 0) {
    SpreadsheetApp.getUi().alert('✅ ビザ期限が90日以内に迫っている外国人はいません。');
    return;
  }

  const message = alerts.map(a =>
    `【${a.daysLeft <= 30 ? '緊急' : '要注意'}】${a.name}: あと${a.daysLeft}日 (${a.expiry})`
  ).join('\n');

  SpreadsheetApp.getUi().alert(`⚠️ ビザ期限アラート（${alerts.length}名）\n\n${message}`);
  updateDashboard();
}

function refreshWorkerList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workersSheet = ss.getSheetByName(CONFIG.SHEETS.WORKERS);
  const data = workersSheet.getDataRange().getValues();

  // 企業名を企業IDから自動更新
  const companiesSheet = ss.getSheetByName(CONFIG.SHEETS.COMPANIES);
  const companyData = companiesSheet ? companiesSheet.getDataRange().getValues() : [];
  const companyMap = {};
  for (let i = 1; i < companyData.length; i++) {
    if (companyData[i][0]) companyMap[companyData[i][0]] = companyData[i][1];
  }

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) break;
    const cid = data[i][9];
    if (cid && companyMap[cid]) {
      workersSheet.getRange(i + 1, 11).setValue(companyMap[cid]);
    }
  }
  SpreadsheetApp.getUi().alert('✅ 外国人リストを更新しました。');
}

function addSupportPlanRow_(workerId, name, companyName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SUPPORT_PLANS);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const planId = 'SP' + String(lastRow).padStart(4, '0');
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  const baseData = [planId, workerId, name, companyName, today, ''];
  const supportStatuses = CONFIG.SUPPORT_TYPES.map(() => '未実施');
  const row = [...baseData, ...supportStatuses, '=COUNTIF(G'+lastRow+1+':R'+lastRow+1+',"実施済")/12', ''];
  sheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
}

function scheduleQuarterlyInterviews_(workerId, name, startDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.INTERVIEWS);
  if (!sheet) return;

  const startDate = new Date(startDateStr);
  const today = new Date();
  const companyName = '';
  let lastRow = sheet.getLastRow();

  // 今後1年分の四半期面談を予定
  for (let q = 1; q <= 4; q++) {
    const interviewDate = new Date(startDate);
    interviewDate.setMonth(interviewDate.getMonth() + (q * 3));
    if (interviewDate < today) continue;

    const interviewId = 'I' + String(lastRow + 1).padStart(5, '0');
    const row = [
      interviewId, workerId, name, companyName,
      Utilities.formatDate(interviewDate, 'Asia/Tokyo', 'yyyy/MM/dd'),
      '未実施', '', '', '対面', '', '', '', '', '', '', '', '', '',
    ];
    sheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
    lastRow++;
  }
}

function getAddWorkerFormHtml_() {
  return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Noto Sans JP', sans-serif; padding: 15px; font-size: 14px; }
    label { display: block; margin-top: 10px; font-weight: bold; color: #333; }
    input, select, textarea { width: 100%; padding: 6px; margin-top: 3px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
    .row { display: flex; gap: 10px; }
    .row > div { flex: 1; }
    button { margin-top: 15px; padding: 10px 20px; background: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; width: 100%; }
    button:hover { background: #1557b0; }
    h3 { color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 5px; }
  </style>
</head>
<body>
  <h3>外国人 新規登録</h3>
  <form id="workerForm">
    <div class="row">
      <div>
        <label>氏名（ローマ字）*</label>
        <input type="text" id="nameRoman" required placeholder="NGUYEN VAN A">
      </div>
      <div>
        <label>氏名（母国語）</label>
        <input type="text" id="nameNative" placeholder="グエン・ヴァン・A">
      </div>
    </div>
    <div class="row">
      <div>
        <label>国籍*</label>
        <select id="nationality" required>
          <option value="">選択してください</option>
          ${CONFIG.NATIONALITIES.map(n => `<option value="${n}">${n}</option>`).join('')}
        </select>
      </div>
      <div>
        <label>性別</label>
        <select id="gender">
          <option value="男性">男性</option>
          <option value="女性">女性</option>
          <option value="その他">その他</option>
        </select>
      </div>
    </div>
    <div class="row">
      <div>
        <label>生年月日</label>
        <input type="date" id="birthDate">
      </div>
      <div>
        <label>在留資格*</label>
        <select id="residenceStatus" required>
          ${CONFIG.RESIDENCE_STATUS.map(r => `<option value="${r}">${r}</option>`).join('')}
        </select>
      </div>
    </div>
    <div class="row">
      <div>
        <label>在留期限*</label>
        <input type="date" id="visaExpiry" required>
      </div>
      <div>
        <label>在留カード番号</label>
        <input type="text" id="cardNumber" placeholder="AB12345678CD">
      </div>
    </div>
    <div class="row">
      <div>
        <label>入国日</label>
        <input type="date" id="entryDate">
      </div>
      <div>
        <label>就労開始日</label>
        <input type="date" id="workStartDate">
      </div>
    </div>
    <label>分野*</label>
    <select id="sector" required>
      <option value="">選択してください</option>
      ${CONFIG.SECTORS.map(s => `<option value="${s}">${s}</option>`).join('')}
    </select>
    <label>受入企業ID</label>
    <input type="text" id="companyId" placeholder="C0001">
    <label>受入企業名</label>
    <input type="text" id="companyName">
    <div class="row">
      <div>
        <label>連絡先（電話）</label>
        <input type="text" id="phone" placeholder="080-0000-0000">
      </div>
      <div>
        <label>連絡先（Email）</label>
        <input type="email" id="email">
      </div>
    </div>
    <label>備考</label>
    <textarea id="notes" rows="2"></textarea>
    <button type="button" onclick="submitForm()">登録する</button>
  </form>
  <script>
    function submitForm() {
      const formData = {
        nameRoman: document.getElementById('nameRoman').value,
        nameNative: document.getElementById('nameNative').value,
        nationality: document.getElementById('nationality').value,
        gender: document.getElementById('gender').value,
        birthDate: document.getElementById('birthDate').value,
        residenceStatus: document.getElementById('residenceStatus').value,
        visaExpiry: document.getElementById('visaExpiry').value,
        cardNumber: document.getElementById('cardNumber').value,
        entryDate: document.getElementById('entryDate').value,
        workStartDate: document.getElementById('workStartDate').value,
        sector: document.getElementById('sector').value,
        companyId: document.getElementById('companyId').value,
        companyName: document.getElementById('companyName').value,
        phone: document.getElementById('phone').value,
        email: document.getElementById('email').value,
        notes: document.getElementById('notes').value,
      };
      if (!formData.nameRoman || !formData.nationality || !formData.visaExpiry) {
        alert('必須項目（*）を入力してください');
        return;
      }
      google.script.run
        .withSuccessHandler(result => {
          alert('✅ 登録完了: ID=' + result.id);
          google.script.host.close();
        })
        .withFailureHandler(err => alert('エラー: ' + err.message))
        .addWorker(formData);
    }
  </script>
</body>
</html>`;
}
