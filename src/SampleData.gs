// ===== サンプルデータ投入（動作確認用） =====
// メニューには出していません。Apps Scriptエディタから loadSampleData を直接実行してください。
// 本番運用前に clearSampleData で削除できます。

function loadSampleData() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    'サンプルデータ投入',
    '動作確認用のサンプルデータ（企業2社・外国人3名・面談・請求書）を投入します。\n実行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  const fmt = (d) => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
  const addDays = (n) => fmt(new Date(today.getTime() + n * 86400000));

  // --- 受入企業 ---
  const companies = ss.getSheetByName(CONFIG.SHEETS.COMPANIES);
  companies.getRange(2, 1, 2, 15).setValues([
    ['C0001', '株式会社サンプル製造', '1234567890123', '山田太郎', '田中花子', '03-1111-2222',
      '東京都品川区1-1-1', '素形材・産業機械・電気電子情報関連製造業', addDays(-200), addDays(165),
      20000, 2, '', '契約中', 'サンプル'],
    ['C0002', 'サンプル介護サービス株式会社', '9876543210987', '佐藤次郎', '鈴木一郎', '06-3333-4444',
      '大阪府大阪市北区2-2-2', '介護', addDays(-100), addDays(265),
      25000, 1, '', '契約中', 'サンプル'],
  ]);

  // --- 外国人 ---
  const workers = ss.getSheetByName(CONFIG.SHEETS.WORKERS);
  workers.getRange(2, 1, 3, 20).setValues([
    ['W0001', 'NGUYEN VAN A', 'グエン・ヴァン・A', '在籍中', 'ベトナム', '1998/05/12',
      '男性', '080-1111-1111', 'nguyen@example.com', 'C0001', '株式会社サンプル製造',
      '素形材・産業機械・電気電子情報関連製造業', '特定技能1号', addDays(25), 'AB12345678CD',
      addDays(-180), addDays(-175), '', '', 'サンプル（ビザ期限間近）'],
    ['W0002', 'SANTOS MARIA', 'サントス・マリア', '在籍中', 'フィリピン', '1995/11/03',
      '女性', '080-2222-2222', 'santos@example.com', 'C0002', 'サンプル介護サービス株式会社',
      '介護', '特定技能1号', addDays(120), 'EF23456789GH',
      addDays(-90), addDays(-85), '', '', 'サンプル'],
    ['W0003', 'PUTRA BUDI', 'プトラ・ブディ', '在籍中', 'インドネシア', '1999/02/20',
      '男性', '080-3333-3333', 'putra@example.com', 'C0001', '株式会社サンプル製造',
      '素形材・産業機械・電気電子情報関連製造業', '特定技能1号', addDays(300), 'IJ34567890KL',
      addDays(-60), addDays(-55), '', '', 'サンプル'],
  ]);

  // --- 支援計画（各外国人）---
  const plans = ss.getSheetByName(CONFIG.SHEETS.SUPPORT_PLANS);
  [['W0001', 'NGUYEN VAN A', '株式会社サンプル製造'],
   ['W0002', 'SANTOS MARIA', 'サンプル介護サービス株式会社'],
   ['W0003', 'PUTRA BUDI', '株式会社サンプル製造']].forEach((w, i) => {
    const r = 2 + i;
    const statuses = CONFIG.SUPPORT_TYPES.map((_, idx) => idx < 4 ? '実施済' : '未実施');
    const rate = `=COUNTIF(G${r}:R${r},"実施済")/12`;
    const rowData = [`SP000${i + 1}`, w[0], w[1], w[2], fmt(today), '', ...statuses, rate, 'サンプル'];
    plans.getRange(r, 1, 1, rowData.length).setValues([rowData]);
    plans.getRange(r, 6 + CONFIG.SUPPORT_TYPES.length + 1).setNumberFormat('0%');
  });

  // --- 定期面談 ---
  const interviews = ss.getSheetByName(CONFIG.SHEETS.INTERVIEWS);
  interviews.getRange(2, 1, 3, 18).setValues([
    ['I00001', 'W0001', 'NGUYEN VAN A', '株式会社サンプル製造', addDays(-5), '実施済',
      addDays(-5), '田中花子', '対面', '本社会議室', '良好', '良好', '良好',
      '特になし', '継続支援', addDays(85), '', 'サンプル'],
    ['I00002', 'W0002', 'SANTOS MARIA', 'サンプル介護サービス株式会社', addDays(7), '未実施',
      '', '', '対面', '', '', '', '', '', '', '', '', 'サンプル（実施予定）'],
    ['I00003', 'W0003', 'PUTRA BUDI', '株式会社サンプル製造', addDays(-3), '未実施',
      '', '', '対面', '', '', '', '', '', '', '', '', 'サンプル（実施漏れ）'],
  ]);

  // --- 請求書 ---
  const invoices = ss.getSheetByName(CONFIG.SHEETS.INVOICES);
  const ym = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM');
  invoices.getRange(2, 1, 2, 13).setValues([
    ['INV-SAMPLE-001', 'C0001', '株式会社サンプル製造', ym, fmt(today),
      44000, '未払い', '', addDays(30), '', 2, '月次支援費 2名 × ¥20,000', 'サンプル'],
    ['INV-SAMPLE-002', 'C0002', 'サンプル介護サービス株式会社', ym, addDays(-40),
      27500, '未払い', '', addDays(-10), '', 1, '月次支援費 1名 × ¥25,000', 'サンプル（支払期限超過）'],
  ]);

  updateDashboard();
  SpreadsheetApp.flush();
  ui.alert('✅ サンプルデータを投入しました。\nダッシュボードでアラート表示を確認できます。\n\n本番運用前に clearSampleData() で削除してください。');
}

function clearSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [CONFIG.SHEETS.WORKERS, CONFIG.SHEETS.COMPANIES, CONFIG.SHEETS.SUPPORT_PLANS,
   CONFIG.SHEETS.INTERVIEWS, CONFIG.SHEETS.INVOICES].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      const lastCol = data[i].length;
      if (String(data[i][lastCol - 1]).indexOf('サンプル') === 0 ||
          String(data[i][0]).indexOf('SAMPLE') >= 0 ||
          ['W0001', 'W0002', 'W0003', 'C0001', 'C0002'].indexOf(data[i][0]) >= 0) {
        sheet.deleteRow(i + 1);
      }
    }
  });
  updateDashboard();
  SpreadsheetApp.getUi().alert('✅ サンプルデータを削除しました。');
}
