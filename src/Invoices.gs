// ===== 請求書管理 =====

function createInvoiceDoc() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const companiesSheet = ss.getSheetByName(CONFIG.SHEETS.COMPANIES);

  if (!companiesSheet) {
    ui.alert('企業管理シートが見つかりません。初期化してください。');
    return;
  }

  // 企業一覧を取得してダイアログで選択
  const companies = [];
  const companyData = companiesSheet.getDataRange().getValues();
  for (let i = 1; i < companyData.length; i++) {
    if (companyData[i][0] && companyData[i][13] !== 'キャンセル') {
      companies.push({ id: companyData[i][0], name: companyData[i][1], fee: companyData[i][10] });
    }
  }

  const html = HtmlService.createHtmlOutput(getInvoiceFormHtml_(companies))
    .setWidth(550).setHeight(500);
  ui.showModalDialog(html, '請求書 作成');
}

function generateInvoice(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.INVOICES);

  const lastRow = sheet.getLastRow();
  const year = new Date().getFullYear();
  const month = String(new Date().getMonth() + 1).padStart(2, '0');
  const invoiceNo = `INV-${year}${month}-${String(lastRow).padStart(3, '0')}`;

  const issueDate = new Date();
  const dueDate = new Date(issueDate.getTime() + CONFIG.INVOICE.PAYMENT_TERMS * 24 * 60 * 60 * 1000);
  const yearMonth = `${year}/${month}`;

  const subtotal = formData.fee * formData.workerCount;
  const tax = Math.floor(subtotal * CONFIG.INVOICE.TAX_RATE);
  const total = subtotal + tax;

  // Google Doc で請求書を作成
  const docUrl = createInvoicePdf_(invoiceNo, formData, issueDate, dueDate, subtotal, tax, total);

  // 請求書シートに記録
  const row = [
    invoiceNo, formData.companyId, formData.companyName, yearMonth,
    Utilities.formatDate(issueDate, 'Asia/Tokyo', 'yyyy/MM/dd'),
    total, '未払い', '',
    Utilities.formatDate(dueDate, 'Asia/Tokyo', 'yyyy/MM/dd'),
    docUrl, formData.workerCount,
    `月次支援費 ${formData.workerCount}名 × ¥${formData.fee.toLocaleString()}`,
    formData.notes,
  ];
  sheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);

  updateSalesSummary();

  return { success: true, invoiceNo, docUrl, total };
}

function createBulkInvoices() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '一括請求書作成',
    '全受入企業の今月分請求書を作成します。\n（既に請求済みの企業はスキップします）\n\n実行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (result !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const companiesSheet = ss.getSheetByName(CONFIG.SHEETS.COMPANIES);
  const workersSheet = ss.getSheetByName(CONFIG.SHEETS.WORKERS);
  const companyData = companiesSheet.getDataRange().getValues();
  const workerData = workersSheet.getDataRange().getValues();

  const yearMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM');
  const invoicesSheet = ss.getSheetByName(CONFIG.SHEETS.INVOICES);
  const existingInvoices = invoicesSheet.getDataRange().getValues();
  const alreadyBilled = new Set();
  for (let i = 1; i < existingInvoices.length; i++) {
    if (existingInvoices[i][3] === yearMonth) alreadyBilled.add(existingInvoices[i][1]);
  }

  let count = 0;
  for (let i = 1; i < companyData.length; i++) {
    const cid = companyData[i][0];
    if (!cid || companyData[i][13] === 'キャンセル') continue;
    if (alreadyBilled.has(cid)) continue;

    // この企業の在籍中外国人数をカウント
    const workerCount = workerData.filter(
      (w, idx) => idx > 0 && w[9] === cid && w[3] === '在籍中'
    ).length;
    if (workerCount === 0) continue;

    generateInvoice({
      companyId: cid,
      companyName: companyData[i][1],
      fee: companyData[i][10] || CONFIG.INVOICE.MONTHLY_FEE,
      workerCount,
      notes: `${yearMonth} 月次支援費`,
    });
    count++;
  }

  SpreadsheetApp.getUi().alert(`✅ ${count}社分の請求書を作成しました。（${yearMonth}）`);
}

function showUnpaidInvoices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.INVOICES);
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  const unpaid = [];
  let total = 0;

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) break;
    if (data[i][6] !== '未払い') continue;
    const due = data[i][8] ? new Date(data[i][8]) : null;
    const overdue = due ? Math.floor((today - due) / 86400000) : 0;
    unpaid.push({
      no: data[i][0], company: data[i][2], amount: data[i][5],
      due: data[i][8], overdue,
    });
    total += Number(data[i][5]) || 0;
  }

  if (unpaid.length === 0) {
    SpreadsheetApp.getUi().alert('✅ 未払い請求はありません。');
    return;
  }

  const message = unpaid.map(u =>
    `${u.no} ${u.company}: ¥${u.amount.toLocaleString()} （期限:${u.due} ${u.overdue > 0 ? u.overdue + '日超過' : ''}）`
  ).join('\n');

  SpreadsheetApp.getUi().alert(
    `📌 未払い請求一覧（${unpaid.length}件）\n合計: ¥${total.toLocaleString()}\n\n${message}`
  );
}

function updateSalesSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
  if (salesSheet) {
    SpreadsheetApp.flush();
  }
}

function createInvoicePdf_(invoiceNo, formData, issueDate, dueDate, subtotal, tax, total) {
  const doc = DocumentApp.create(`請求書_${invoiceNo}_${formData.companyName}`);
  const body = doc.getBody();
  body.clear();

  // 請求書タイトル
  const title = body.appendParagraph('請　求　書');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.setSpacingBefore(20);

  body.appendParagraph(`請求書番号: ${invoiceNo}`)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph(`発行日: ${Utilities.formatDate(issueDate, 'Asia/Tokyo', 'yyyy年MM月dd日')}`)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph(`お支払期限: ${Utilities.formatDate(dueDate, 'Asia/Tokyo', 'yyyy年MM月dd日')}`)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph('');

  // 宛名
  body.appendParagraph(`${formData.companyName}　御中`).setFontSize(14).setFontWeight('bold');
  body.appendParagraph('');

  // 発行者情報
  const issuerPara = body.appendParagraph(
    `${CONFIG.COMPANY.NAME}\n${CONFIG.COMPANY.POSTAL_CODE}\n${CONFIG.COMPANY.ADDRESS}\nTEL: ${CONFIG.COMPANY.PHONE}\nEmail: ${CONFIG.COMPANY.EMAIL}\n${CONFIG.COMPANY.REGISTRATION_NUMBER}`
  );
  issuerPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph('');

  // 請求金額
  const amountPara = body.appendParagraph(`ご請求金額（税込）: ¥${total.toLocaleString()}`);
  amountPara.setFontSize(16).setFontWeight('bold')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('上記の通りご請求申し上げます。').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('');

  // 明細テーブル
  body.appendParagraph('＜請求内訳＞').setFontWeight('bold');
  const detailTable = body.appendTable([
    ['品目', '単価', '数量', '金額'],
    [
      `特定技能外国人 月次支援費（${Utilities.formatDate(issueDate, 'Asia/Tokyo', 'yyyy年MM月')}分）`,
      `¥${Number(formData.fee).toLocaleString()}`,
      `${formData.workerCount}名`,
      `¥${subtotal.toLocaleString()}`,
    ],
  ]);
  if (formData.notes) {
    detailTable.appendTableRow().appendTableCell(formData.notes).colSpan = 4;
  }

  body.appendTable([
    ['小計', `¥${subtotal.toLocaleString()}`],
    [`消費税（${CONFIG.INVOICE.TAX_RATE * 100}%）`, `¥${tax.toLocaleString()}`],
    ['合計（税込）', `¥${total.toLocaleString()}`],
  ]);

  body.appendParagraph('');
  body.appendParagraph('＜お振込先＞').setFontWeight('bold');
  body.appendParagraph(
    `${CONFIG.COMPANY.BANK_NAME}　${CONFIG.COMPANY.ACCOUNT_TYPE}　${CONFIG.COMPANY.ACCOUNT_NUMBER}\n口座名義: ${CONFIG.COMPANY.ACCOUNT_HOLDER}`
  );
  body.appendParagraph('※振込手数料はご負担ください。');

  doc.saveAndClose();

  // 請求書フォルダに移動
  const invoicesFolderId = getFolderIdFromSettings_('InvoicesFolderID');
  if (invoicesFolderId) {
    const invoicesFolder = DriveApp.getFolderById(invoicesFolderId);
    DriveApp.getFileById(doc.getId()).moveTo(invoicesFolder);
  }

  return doc.getUrl();
}

function getInvoiceFormHtml_(companies) {
  const companyOptions = companies.map(c =>
    `<option value="${c.id}" data-fee="${c.fee || CONFIG.INVOICE.MONTHLY_FEE}" data-name="${c.name}">${c.name}</option>`
  ).join('');

  const thisMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM');

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
    .preview { background: #f8f9fa; padding: 10px; border-radius: 4px; margin-top: 10px; }
    h3 { color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 5px; }
  </style>
</head>
<body>
  <h3>請求書 作成</h3>
  <label>受入企業*</label>
  <select id="company" onchange="updateFee()" required>
    <option value="">選択してください</option>
    ${companyOptions}
  </select>
  <div class="row">
    <div>
      <label>請求年月</label>
      <input type="text" id="billingMonth" value="${thisMonth}" placeholder="2024/01">
    </div>
    <div>
      <label>支援中外国人数*</label>
      <input type="number" id="workerCount" min="1" required onchange="calcTotal()">
    </div>
  </div>
  <div class="row">
    <div>
      <label>月額支援費（1名あたり）</label>
      <input type="number" id="fee" value="${CONFIG.INVOICE.MONTHLY_FEE}" onchange="calcTotal()">
    </div>
  </div>
  <label>備考・特記事項</label>
  <textarea id="notes" rows="2"></textarea>
  <div class="preview">
    <b>請求金額プレビュー</b>
    <div id="preview">企業と人数を選択してください</div>
  </div>
  <button type="button" onclick="submitForm()">請求書を作成</button>
  <script>
    const taxRate = ${CONFIG.INVOICE.TAX_RATE};
    function updateFee() {
      const sel = document.getElementById('company');
      const opt = sel.options[sel.selectedIndex];
      if (opt && opt.dataset.fee) {
        document.getElementById('fee').value = opt.dataset.fee;
      }
      calcTotal();
    }
    function calcTotal() {
      const count = parseInt(document.getElementById('workerCount').value) || 0;
      const fee = parseInt(document.getElementById('fee').value) || 0;
      const subtotal = count * fee;
      const tax = Math.floor(subtotal * taxRate);
      const total = subtotal + tax;
      document.getElementById('preview').innerHTML = count > 0
        ? '小計: ¥' + subtotal.toLocaleString() + '　消費税: ¥' + tax.toLocaleString() + '　<b>合計: ¥' + total.toLocaleString() + '</b>'
        : '人数を入力してください';
    }
    function submitForm() {
      const sel = document.getElementById('company');
      const opt = sel.options[sel.selectedIndex];
      if (!sel.value || !document.getElementById('workerCount').value) {
        alert('企業と外国人数を入力してください');
        return;
      }
      const formData = {
        companyId: sel.value,
        companyName: opt.dataset.name,
        workerCount: parseInt(document.getElementById('workerCount').value),
        fee: parseInt(document.getElementById('fee').value),
        billingMonth: document.getElementById('billingMonth').value,
        notes: document.getElementById('notes').value,
      };
      google.script.run
        .withSuccessHandler(result => {
          alert('✅ 請求書を作成しました。\\n請求書番号: ' + result.invoiceNo + '\\n金額: ¥' + result.total.toLocaleString());
          google.script.host.close();
        })
        .withFailureHandler(err => alert('エラー: ' + err.message))
        .generateInvoice(formData);
    }
  </script>
</body>
</html>`;
}
