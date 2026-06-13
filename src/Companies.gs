// ===== 受入企業管理 =====

function showAddCompanyDialog() {
  const html = HtmlService.createHtmlOutput(getAddCompanyFormHtml_())
    .setWidth(580).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '受入企業 新規登録');
}

function addCompany(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.COMPANIES);
  const lastRow = sheet.getLastRow();
  const newId = 'C' + String(lastRow).padStart(4, '0');

  // 企業フォルダを作成
  const companiesFolderId = getFolderIdFromSettings_('CompaniesFolderID');
  let folderUrl = '';
  if (companiesFolderId) {
    const companiesFolder = DriveApp.getFolderById(companiesFolderId);
    const companyFolder = companiesFolder.createFolder(`${newId}_${formData.name}`);
    folderUrl = companyFolder.getUrl();
    ['契約書', '請求書', '届出書類', '往復文書'].forEach(sub => companyFolder.createFolder(sub));
  }

  const row = [
    newId, formData.name, formData.corporateNumber, formData.representative,
    formData.contactPerson, formData.contactPhone, formData.address,
    formData.sector, formData.contractStart, formData.contractEnd,
    formData.monthlyFee, 0, folderUrl, '契約中', formData.notes,
  ];

  sheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);

  // カレンダーに契約終了アラートを追加
  if (formData.contractEnd) {
    addContractExpiryToCalendar_(newId, formData.name, formData.contractEnd);
  }

  return { success: true, id: newId };
}

function checkContractExpiry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.COMPANIES);
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  const alerts = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][13] !== '契約中') continue;
    const contractEnd = data[i][9];
    if (!contractEnd) continue;
    const daysLeft = Math.floor((new Date(contractEnd) - today) / 86400000);
    if (daysLeft <= 90) {
      alerts.push({ name: data[i][1], daysLeft, contractEnd });
    }
  }

  if (alerts.length === 0) {
    SpreadsheetApp.getUi().alert('✅ 契約期限が90日以内に迫っている企業はありません。');
    return;
  }

  const message = alerts.map(a =>
    `${a.name}: あと${a.daysLeft}日 (${a.contractEnd})`
  ).join('\n');
  SpreadsheetApp.getUi().alert(`⚠️ 契約期限アラート（${alerts.length}社）\n\n${message}`);
}

function addContractExpiryToCalendar_(companyId, companyName, endDateStr) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const endDate = new Date(endDateStr);
    const alertDate = new Date(endDate.getTime() - 60 * 24 * 60 * 60 * 1000); // 60日前
    calendar.createAllDayEvent(
      `【契約更新確認】${companyName}`,
      alertDate,
      {
        description: `受入企業との契約更新確認\n企業ID: ${companyId}\n契約終了日: ${endDateStr}`,
        color: CalendarApp.EventColor.RED,
      }
    );
  } catch (e) {
    console.log('カレンダー登録エラー: ' + e.message);
  }
}

function getAddCompanyFormHtml_() {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const oneYearLater = Utilities.formatDate(
    new Date(Date.now() + 365 * 24 * 60 * 60 * 1000), 'Asia/Tokyo', 'yyyy-MM-dd'
  );

  return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Noto Sans JP', sans-serif; padding: 15px; font-size: 14px; }
    label { display: block; margin-top: 8px; font-weight: bold; color: #333; }
    input, select, textarea { width: 100%; padding: 6px; margin-top: 3px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
    .row { display: flex; gap: 10px; }
    .row > div { flex: 1; }
    button { margin-top: 15px; padding: 10px; background: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; width: 100%; }
    h3 { color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 5px; }
  </style>
</head>
<body>
  <h3>受入企業 新規登録</h3>
  <label>企業名*</label>
  <input type="text" id="name" required>
  <div class="row">
    <div>
      <label>法人番号（13桁）</label>
      <input type="text" id="corporateNumber" placeholder="0000000000000" maxlength="13">
    </div>
    <div>
      <label>代表者名</label>
      <input type="text" id="representative">
    </div>
  </div>
  <div class="row">
    <div>
      <label>担当者名*</label>
      <input type="text" id="contactPerson" required>
    </div>
    <div>
      <label>担当者連絡先*</label>
      <input type="text" id="contactPhone" required placeholder="03-0000-0000">
    </div>
  </div>
  <label>所在地</label>
  <input type="text" id="address" placeholder="東京都〇〇区...">
  <label>分野（業種）*</label>
  <select id="sector" required>
    <option value="">選択してください</option>
    ${CONFIG.SECTORS.map(s => `<option value="${s}">${s}</option>`).join('')}
  </select>
  <div class="row">
    <div>
      <label>契約開始日</label>
      <input type="date" id="contractStart" value="${today}">
    </div>
    <div>
      <label>契約終了日</label>
      <input type="date" id="contractEnd" value="${oneYearLater}">
    </div>
  </div>
  <label>月額支援費（税抜）</label>
  <input type="number" id="monthlyFee" value="${CONFIG.INVOICE.MONTHLY_FEE}" placeholder="20000">
  <label>備考</label>
  <textarea id="notes" rows="2"></textarea>
  <button type="button" onclick="submitForm()">登録する</button>
  <script>
    function submitForm() {
      const formData = {
        name: document.getElementById('name').value,
        corporateNumber: document.getElementById('corporateNumber').value,
        representative: document.getElementById('representative').value,
        contactPerson: document.getElementById('contactPerson').value,
        contactPhone: document.getElementById('contactPhone').value,
        address: document.getElementById('address').value,
        sector: document.getElementById('sector').value,
        contractStart: document.getElementById('contractStart').value,
        contractEnd: document.getElementById('contractEnd').value,
        monthlyFee: parseInt(document.getElementById('monthlyFee').value) || 0,
        notes: document.getElementById('notes').value,
      };
      if (!formData.name || !formData.contactPerson || !formData.contactPhone) {
        alert('必須項目（*）を入力してください');
        return;
      }
      google.script.run
        .withSuccessHandler(result => {
          alert('✅ 登録完了: ID=' + result.id);
          google.script.host.close();
        })
        .withFailureHandler(err => alert('エラー: ' + err.message))
        .addCompany(formData);
    }
  </script>
</body>
</html>`;
}
