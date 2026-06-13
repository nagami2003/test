// ===== 法定書類 自動生成 =====
// ※本テンプレートは雛形です。実際の運用前に専門家（行政書士・社労士等）の確認を受けてください。

// ---------- 共通ユーティリティ（内部） ----------

/**
 * 生成した Google Doc を ContractsFolderID → なければルートへ moveTo する。
 * @param {GoogleAppsScript.Document.Document} doc
 * @returns {string} doc の URL
 */
function moveDocToContractsFolder_(doc) {
  const contractsFolderId = getFolderIdFromSettings_('ContractsFolderID');
  if (contractsFolderId) {
    try {
      const folder = DriveApp.getFolderById(contractsFolderId);
      DriveApp.getFileById(doc.getId()).moveTo(folder);
    } catch (e) {
      // フォルダが存在しない場合はルートのまま
    }
  }
  return doc.getUrl();
}

/**
 * アクティブ行を安全に取得する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {number|null} 行番号（1=ヘッダー以下なら null）
 */
function getActiveDataRow_(sheet) {
  const row = sheet.getActiveCell().getRow();
  return row <= 1 ? null : row;
}

/** 本日を 'yyyy年MM月dd日' でフォーマットして返す。 */
function todayJp_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日');
}

/**
 * Date または日付文字列を 'yyyy年MM月dd日' にフォーマット。値がなければ空文字。
 * @param {*} val
 * @returns {string}
 */
function dateJp_(val) {
  if (!val) return '';
  try {
    return Utilities.formatDate(new Date(val), 'Asia/Tokyo', 'yyyy年MM月dd日');
  } catch (e) {
    return String(val);
  }
}

/**
 * Document の Body に見出し段落を追加して返す。
 * @param {GoogleAppsScript.Document.Body} body
 * @param {string} text
 * @param {GoogleAppsScript.Document.ParagraphHeading} heading
 * @returns {GoogleAppsScript.Document.Paragraph}
 */
function appendHeading_(body, text, heading) {
  const p = body.appendParagraph(text);
  p.setHeading(heading);
  return p;
}

/**
 * 2列テーブル（項目名｜記入欄）を appendTable で追加して返す。
 * @param {GoogleAppsScript.Document.Body} body
 * @param {string[][]} rows  [[label, value], ...]
 * @returns {GoogleAppsScript.Document.Table}
 */
function appendInfoTable_(body, rows) {
  const table = body.appendTable(rows);
  table.setBorderColor('#dadce0');
  return table;
}

/**
 * 共通フッター（作成者情報・署名欄）を追加する。
 * @param {GoogleAppsScript.Document.Body} body
 * @param {string} signerALabel  左署名者ラベル
 * @param {string} signerBLabel  右署名者ラベル
 */
function appendSignatureBlock_(body, signerALabel, signerBLabel) {
  body.appendParagraph('');
  body.appendHorizontalRule();
  appendHeading_(body, '署名欄', DocumentApp.ParagraphHeading.HEADING2);
  body.appendTable([
    [signerALabel, '（署名・捺印）', signerBLabel, '（署名・捺印）'],
    ['', '', '', ''],
    ['日付：　　　　年　　月　　日', '', '日付：　　　　年　　月　　日', ''],
  ]).setBorderColor('#dadce0');
}

// ============================================================
// 1. 雇用条件書
// ============================================================

/**
 * 雇用条件書（特定技能雇用契約に関する重要事項）を Google Doc として生成する。
 * 外国人管理シートのアクティブ行から外国人情報を読み取り、CONFIG.COMPANY の
 * 自社情報を差し込む。生成した Doc は ContractsFolderID フォルダへ移動する。
 *
 * ※本テンプレートは雛形です。実際の運用前に専門家（行政書士・社労士等）の
 *   確認を受けてください。
 */
function createEmploymentConditionsDoc() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.WORKERS);

  if (!sheet) {
    ui.alert('外国人管理シートが見つかりません。初期化してください。');
    return;
  }

  const row = getActiveDataRow_(sheet);
  if (!row) {
    ui.alert('対象の行を選択してください。');
    return;
  }

  // 外国人管理シート列定義（0始まり）
  // 0:ID 1:氏名(ローマ字) 2:氏名(母国語) 3:ステータス 4:国籍 5:生年月日
  // 6:性別 7:電話 8:Email 9:受入企業ID 10:受入企業名 11:分野
  // 12:在留資格 13:在留期限 14:在留カード番号 15:入国日 16:就労開始日
  const data = sheet.getRange(row, 1, 1, 20).getValues()[0];
  const workerId      = data[0]  || '';
  const nameRoman     = data[1]  || '';
  const nameNative    = data[2]  || '';
  const nationality   = data[4]  || '';
  const birthday      = dateJp_(data[5]);
  const gender        = data[6]  || '';
  const companyName   = data[10] || '';
  const sector        = data[11] || '';
  const residenceStatus = data[12] || '';
  const visaExpiry    = dateJp_(data[13]);
  const workStart     = dateJp_(data[16]);

  const docTitle = `雇用条件書_${workerId}_${nameRoman}`;
  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();
  body.clear();

  // ---- タイトル ----
  const title = body.appendParagraph('雇用条件書');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph(
    '※本書は特定技能雇用契約に関する重要事項を記載した雛形です。' +
    '専門家（行政書士・社労士等）の確認を受けた上でご使用ください。'
  ).setItalic(true);
  body.appendParagraph('');

  // ---- 作成情報 ----
  appendInfoTable_(body, [
    ['作成日',         todayJp_()],
    ['作成機関名',     CONFIG.COMPANY.NAME],
    ['登録番号',       CONFIG.COMPANY.REGISTRATION_NUMBER],
    ['受入企業名',     companyName],
  ]);
  body.appendParagraph('');

  // ---- 第1条 当事者情報 ----
  appendHeading_(body, '第1条　当事者情報', DocumentApp.ParagraphHeading.HEADING2);
  appendInfoTable_(body, [
    ['雇用主（受入企業）名', companyName || '（　　　　　　　　　）'],
    ['外国人氏名（ローマ字）', nameRoman || '（　　　　　　　　　）'],
    ['外国人氏名（母国語）',   nameNative || '（　　　　　　　　　）'],
    ['国籍',                   nationality || '（　　　　）'],
    ['生年月日',               birthday || '（　　　　年　　月　　日）'],
    ['性別',                   gender || '（　　）'],
    ['在留資格',               residenceStatus || '（　　　　　　）'],
    ['在留期限',               visaExpiry || '（　　　　年　　月　　日）'],
  ]);
  body.appendParagraph('');

  // ---- 第2条 業務内容 ----
  appendHeading_(body, '第2条　業務内容', DocumentApp.ParagraphHeading.HEADING2);
  appendInfoTable_(body, [
    ['就労分野',   sector || '（　　　　　　）'],
    ['従事する業務', '（　　　　　　　　　　　　　　　　　　）'],
    ['就業場所',   '（　　　　　　　　　　　　　　　　　　）'],
    ['就労開始予定日', workStart || '（　　　　年　　月　　日）'],
  ]);
  body.appendParagraph('');

  // ---- 第3条 労働時間・休憩・休日 ----
  appendHeading_(body, '第3条　労働時間・休憩・休日', DocumentApp.ParagraphHeading.HEADING2);
  appendInfoTable_(body, [
    ['所定労働時間', '（　）時間（　）分 ／日'],
    ['始業時刻',     '（　　：　　）'],
    ['終業時刻',     '（　　：　　）'],
    ['休憩時間',     '（　　）分'],
    ['所定休日',     '（週　日・その他：　　　　　　　）'],
    ['時間外労働',   '有（月　　時間以内）・無'],
    ['深夜・休日労働', '有・無'],
  ]);
  body.appendParagraph('');

  // ---- 第4条 賃金 ----
  appendHeading_(body, '第4条　賃金', DocumentApp.ParagraphHeading.HEADING2);
  appendInfoTable_(body, [
    ['基本給',           '¥（　　　　　　　）円　／月（時給の場合：¥　　　円）'],
    ['各種手当',         '（　　　　　　　　　　　　　　　）'],
    ['賃金支払日',       '毎月　（　　）日払い（休日の場合：　　日）'],
    ['賃金支払方法',     '銀行振込・現金手渡し（いずれかに〇）'],
    ['控除項目（法定）', '所得税・住民税・社会保険料・雇用保険料'],
    ['控除項目（協定）', '（　　　　　　　　　　　　　　　）'],
    ['最低賃金との比較', '適用地域別最低賃金：¥（　　　）円以上であることを確認'],
  ]);
  body.appendParagraph('');

  // ---- 第5条 雇用契約期間 ----
  appendHeading_(body, '第5条　雇用契約期間', DocumentApp.ParagraphHeading.HEADING2);
  appendInfoTable_(body, [
    ['契約形態',   '期間の定めあり・期間の定めなし（いずれかに〇）'],
    ['契約開始日', workStart || '（　　　　年　　月　　日）'],
    ['契約終了日', visaExpiry || '（在留期限に合わせて設定：　　　　年　　月　　日）'],
    ['更新の有無', '自動更新・協議の上更新・更新なし（いずれかに〇）'],
  ]);
  body.appendParagraph('');

  // ---- 第6条 社会保険・各種保険 ----
  appendHeading_(body, '第6条　社会保険・各種保険', DocumentApp.ParagraphHeading.HEADING2);
  appendInfoTable_(body, [
    ['健康保険',   '加入（　　　　健康保険組合 or 協会けんぽ）・適用除外'],
    ['厚生年金',   '加入・適用除外'],
    ['雇用保険',   '加入・適用除外'],
    ['労災保険',   '加入（全労働者に適用）'],
  ]);
  body.appendParagraph('');

  // ---- 第7条 母国語での確認 ----
  appendHeading_(body, '第7条　母国語での説明・確認', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(
    '本書の内容は、外国人本人が十分理解できる言語（母国語）で説明を行い、' +
    '本人が内容を理解したことを確認しました。'
  );
  body.appendParagraph('');
  appendInfoTable_(body, [
    ['説明使用言語',   '（　　　　　　　語）'],
    ['通訳者氏名等',   '（　　　　　　　　　　　　）または本人が日本語理解可能'],
    ['確認日',         '（　　　　年　　月　　日）'],
  ]);
  body.appendParagraph('');

  // ---- 母国語併記欄（構造のみ） ----
  appendHeading_(body, '【母国語併記欄 / Native Language Column】', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(
    '※ 以下の欄は、外国人本人が確認できるよう母国語（またはやさしい日本語）で' +
    '重要事項を記載・翻訳してください。翻訳は専門家・通訳者が行うことを推奨します。'
  ).setItalic(true);
  appendInfoTable_(body, [
    ['氏名 / Name', nameRoman + ' / ' + nameNative],
    ['業務内容 / Work Description', '（翻訳記入欄）'],
    ['賃金 / Salary',              '（翻訳記入欄）'],
    ['労働時間 / Working Hours',    '（翻訳記入欄）'],
    ['休日 / Holidays',             '（翻訳記入欄）'],
    ['契約期間 / Contract Period',   '（翻訳記入欄）'],
    ['相談窓口 / Contact',          CONFIG.COMPANY.PHONE + ' / ' + CONFIG.COMPANY.EMAIL],
  ]);
  body.appendParagraph('');

  // ---- 署名欄 ----
  appendSignatureBlock_(body, '雇用主（受入企業）代表者', '外国人本人');

  doc.saveAndClose();

  const docUrl = moveDocToContractsFolder_(doc);

  ui.alert('✅ 雇用条件書を作成しました。\n\nURL: ' + docUrl);
}

// ============================================================
// 2. 事前ガイダンス実施確認書
// ============================================================

/**
 * 事前ガイダンス実施確認書を Google Doc として生成する。
 * 外国人管理シートのアクティブ行から外国人情報を読み取る。
 * 法定説明事項のチェックリスト形式。
 *
 * ※本テンプレートは雛形です。実際の運用前に専門家（行政書士・社労士等）の
 *   確認を受けてください。
 */
function createPreGuidanceConfirmationDoc() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.WORKERS);

  if (!sheet) {
    ui.alert('外国人管理シートが見つかりません。初期化してください。');
    return;
  }

  const row = getActiveDataRow_(sheet);
  if (!row) {
    ui.alert('対象の行を選択してください。');
    return;
  }

  const data = sheet.getRange(row, 1, 1, 20).getValues()[0];
  const workerId    = data[0]  || '';
  const nameRoman   = data[1]  || '';
  const nameNative  = data[2]  || '';
  const nationality = data[4]  || '';
  const companyName = data[10] || '';
  const sector      = data[11] || '';

  const docTitle = `事前ガイダンス実施確認書_${workerId}_${nameRoman}`;
  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();
  body.clear();

  // ---- タイトル ----
  const title = body.appendParagraph('事前ガイダンス実施確認書');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph(
    '※本書は特定技能外国人への事前ガイダンス（入管法施行規則第19条の17等）に基づく' +
    '説明実施記録の雛形です。専門家（行政書士・社労士等）の確認を受けてください。'
  ).setItalic(true);
  body.appendParagraph('');

  // ---- 基本情報 ----
  appendHeading_(body, '1. 基本情報', DocumentApp.ParagraphHeading.HEADING2);
  appendInfoTable_(body, [
    ['実施日',           todayJp_() + '　（実施後に記入）'],
    ['実施機関名',       CONFIG.COMPANY.NAME],
    ['登録番号',         CONFIG.COMPANY.REGISTRATION_NUMBER],
    ['担当支援員',       '（　　　　　　　　　）'],
    ['受入企業名',       companyName || '（　　　　　　　　　）'],
    ['外国人氏名（ローマ字）', nameRoman || '（　　　　　　　）'],
    ['外国人氏名（母国語）',   nameNative || '（　　　　　　　）'],
    ['国籍',             nationality || '（　　　　）'],
    ['従事する分野',     sector || '（　　　　　　）'],
    ['実施方法',         '対面・オンライン・書面交付（該当に〇）'],
    ['使用言語',         '（　　　　　　　語）　通訳者：（　　　　　　　）'],
    ['実施場所/URL',     '（　　　　　　　　　　　　　　　　　　）'],
  ]);
  body.appendParagraph('');

  // ---- チェックリスト ----
  appendHeading_(body, '2. 説明事項チェックリスト', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(
    '以下の各事項について説明を行い、外国人本人の理解を確認してください。' +
    '説明済みの項目に「☑」、未説明は「□」を記入してください。'
  );
  body.appendParagraph('');

  const checkItems = [
    ['□', '① 従事する業務内容',
      '担当業務の内容・範囲、使用する機械・器具等の概要'],
    ['□', '② 報酬・賃金の額および支払方法',
      '基本給・各種手当・控除・支払日・最低賃金との比較'],
    ['□', '③ 労働時間・休憩・休日',
      '所定労働時間・時間外労働の有無・休日・有給休暇'],
    ['□', '④ 在留資格（特定技能）の内容・在留期間',
      '特定技能1号/2号の違い・更新手続の概要・在留カードの管理'],
    ['□', '⑤ 入国手続き・入国後の生活',
      '入国の流れ・空港送迎・住居の確保・各種手続の案内'],
    ['□', '⑥ 保証金の徴収・財産管理の禁止',
      '送出機関・仲介業者等が保証金を徴収することは禁止されている旨'],
    ['□', '⑦ 支援費用を本人に負担させないこと',
      '1号特定技能支援計画に基づく支援に関し費用を負担させてはならない旨'],
    ['□', '⑧ 本邦外での活動に関する問題の禁止',
      '不法就労・人身取引等に関与しないこと'],
    ['□', '⑨ 帰国が困難な場合の対応',
      '帰国費用の負担・帰国困難時の支援機関への相談'],
    ['□', '⑩ 相談・苦情窓口',
      CONFIG.COMPANY.NAME + '  TEL: ' + CONFIG.COMPANY.PHONE +
      '  Email: ' + CONFIG.COMPANY.EMAIL],
    ['□', '⑪ 行政機関の相談窓口',
      '出入国在留管理庁・労働基準監督署・都道府県労働局等の連絡先の案内'],
    ['□', '⑫ 宗教・慣習・差別禁止',
      '宗教的慣行への配慮・差別・ハラスメントの禁止'],
  ];

  const tableRows = [['確認', '説明事項', '内容']].concat(checkItems);
  appendInfoTable_(body, tableRows);
  body.appendParagraph('');

  // ---- 確認結果 ----
  appendHeading_(body, '3. 確認結果', DocumentApp.ParagraphHeading.HEADING2);
  appendInfoTable_(body, [
    ['全事項の説明完了', '完了・一部未了（理由：　　　　　　　　　　　　　）'],
    ['外国人本人の理解度', '十分理解・おおむね理解・再説明が必要'],
    ['追加質問・特記事項', '（　　　　　　　　　　　　　　　　　　　　　）'],
  ]);
  body.appendParagraph('');

  // ---- 署名欄 ----
  appendSignatureBlock_(body, '支援機関担当者', '外国人本人');

  doc.saveAndClose();

  const docUrl = moveDocToContractsFolder_(doc);

  ui.alert('✅ 事前ガイダンス実施確認書を作成しました。\n\nURL: ' + docUrl);
}

// ============================================================
// 3. 支援委託契約書
// ============================================================

/**
 * 支援委託契約書（受入機関と登録支援機関の間の業務委託契約書）を
 * Google Doc として生成する。受入企業管理シートのアクティブ行を使用する。
 *
 * ※本テンプレートは雛形です。実際の運用前に専門家（行政書士・社労士等）の
 *   確認を受けてください。
 */
function createSupportEntrustmentContractDoc() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.COMPANIES);

  if (!sheet) {
    ui.alert('受入企業管理シートが見つかりません。初期化してください。');
    return;
  }

  const row = getActiveDataRow_(sheet);
  if (!row) {
    ui.alert('対象の行を選択してください。');
    return;
  }

  // 受入企業管理シート列定義（0始まり）
  // 0:企業ID 1:企業名 2:法人番号 3:代表者名 4:担当者名 5:担当者連絡先
  // 6:住所 7:業種(分野) 8:契約開始日 9:契約終了日 10:月額支援費
  // 11:支援中外国人数 12:企業フォルダURL 13:ステータス 14:備考
  const data = sheet.getRange(row, 1, 1, 15).getValues()[0];
  const companyId        = data[0]  || '';
  const companyName      = data[1]  || '';
  const corporateNumber  = data[2]  || '';
  const representativeName = data[3] || '';
  const contactName      = data[4]  || '';
  const contactTel       = data[5]  || '';
  const address          = data[6]  || '';
  const sector           = data[7]  || '';
  const contractStart    = dateJp_(data[8]);
  const contractEnd      = dateJp_(data[9]);
  const monthlyFee       = data[10] !== '' ? Number(data[10]) : CONFIG.INVOICE.MONTHLY_FEE;

  const docTitle = `支援委託契約書_${companyId}_${companyName}`;
  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();
  body.clear();

  // ---- タイトル ----
  const title = body.appendParagraph('特定技能外国人支援委託契約書');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph(
    '※本書は受入機関と登録支援機関の間の支援委託に関する雛形です。' +
    '専門家（行政書士・社労士等）の確認を受けてください。'
  ).setItalic(true);
  body.appendParagraph('');

  body.appendParagraph(
    '（以下「甲」という）と（以下「乙」という）は、' +
    '特定技能外国人の受入れに関する支援業務の委託につき、以下のとおり契約を締結する。'
  );
  body.appendParagraph('');

  // ---- 当事者情報 ----
  appendHeading_(body, '当事者情報', DocumentApp.ParagraphHeading.HEADING2);
  appendInfoTable_(body, [
    ['甲（委託者・受入機関）企業名', companyName || '（　　　　　　　　　）'],
    ['甲　法人番号',                corporateNumber || '（　　　　　　　　　）'],
    ['甲　代表者名',                representativeName || '（　　　　　　　　　）'],
    ['甲　住所',                    address || '（　　　　　　　　　　　　　）'],
    ['甲　担当者・連絡先',          (contactName || '（　　　）') + '　TEL: ' + (contactTel || '（　　　　　）')],
    ['乙（受託者・登録支援機関）',   CONFIG.COMPANY.NAME],
    ['乙　登録番号',                CONFIG.COMPANY.REGISTRATION_NUMBER],
    ['乙　住所',                    CONFIG.COMPANY.ADDRESS],
    ['乙　連絡先',                  'TEL: ' + CONFIG.COMPANY.PHONE + '　Email: ' + CONFIG.COMPANY.EMAIL],
  ]);
  body.appendParagraph('');

  // ---- 第1条 目的 ----
  appendHeading_(body, '第1条（目的）', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(
    '甲は、出入国管理及び難民認定法（以下「入管法」という）第2条の5第6項の規定に基づき、' +
    '甲が受け入れる特定技能外国人（以下「特定技能外国人」という）に対して実施すべき' +
    '1号特定技能支援計画に定める支援業務を、登録支援機関である乙に委託し、乙はこれを受託する。'
  );
  body.appendParagraph('');

  // ---- 第2条 委託業務範囲 ----
  appendHeading_(body, '第2条（委託する支援業務の範囲）', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('乙が実施する支援業務の範囲は以下のとおりとする。');
  body.appendParagraph('');

  const supportRangeRows = [['支援種別', '実施有無']].concat(
    CONFIG.SUPPORT_TYPES.map(t => [t, '実施・対象外（いずれかに〇）'])
  );
  appendInfoTable_(body, supportRangeRows);
  body.appendParagraph('');

  // ---- 第3条 委託料 ----
  appendHeading_(body, '第3条（委託料）', DocumentApp.ParagraphHeading.HEADING2);
  appendInfoTable_(body, [
    ['委託料（月額・1名あたり）', '¥' + monthlyFee.toLocaleString() + ' 円（税別）'],
    ['消費税', '別途 ' + (CONFIG.INVOICE.TAX_RATE * 100) + '% を加算'],
    ['支払方法', '銀行振込（振込手数料は甲の負担）'],
    ['請求・支払サイクル', '月末締め、翌月' + CONFIG.INVOICE.PAYMENT_TERMS + '日以内払い'],
    ['振込先', CONFIG.COMPANY.BANK_NAME + '　' + CONFIG.COMPANY.ACCOUNT_TYPE + '　' + CONFIG.COMPANY.ACCOUNT_NUMBER],
    ['口座名義', CONFIG.COMPANY.ACCOUNT_HOLDER],
  ]);
  body.appendParagraph('');

  // ---- 第4条 契約期間 ----
  appendHeading_(body, '第4条（契約期間）', DocumentApp.ParagraphHeading.HEADING2);
  appendInfoTable_(body, [
    ['契約開始日', contractStart || '（　　　　年　　月　　日）'],
    ['契約終了日', contractEnd   || '（　　　　年　　月　　日）'],
    ['自動更新',   '期間満了1ヶ月前までに書面による解除通知がない場合、同条件で自動更新'],
    ['中途解約',   '3ヶ月前に書面で相手方に通知することにより解約できる'],
  ]);
  body.appendParagraph('');

  // ---- 第5条 責任分担 ----
  appendHeading_(body, '第5条（責任分担）', DocumentApp.ParagraphHeading.HEADING2);
  appendInfoTable_(body, [
    ['甲の義務',
      '支援に必要な情報・書類の提供、特定技能外国人への支援協力、' +
      '委託料の期日内支払い、入管法その他関係法令の遵守'],
    ['乙の義務',
      '支援計画に基づく支援の誠実な実施、支援記録の作成・保管（5年間）、' +
      '3ヶ月に1回以上の定期報告、関係法令の遵守'],
    ['乙の免責',
      '甲が提供した情報の虚偽・不正確に起因する損害については乙は責任を負わない'],
    ['秘密保持',
      '両者は業務上知り得た相手方の情報を第三者に漏洩してはならない（契約終了後も継続）'],
  ]);
  body.appendParagraph('');

  // ---- 第6条 法令遵守 ----
  appendHeading_(body, '第6条（法令遵守・禁止事項）', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(
    '甲および乙は、本契約の履行にあたり、入管法、労働基準法、最低賃金法その他の' +
    '関係法令を遵守するものとする。また、特定技能外国人から保証金を徴収し、' +
    'または支援費用を負担させる行為を行ってはならない。'
  );
  body.appendParagraph('');

  // ---- 第7条 管轄裁判所 ----
  appendHeading_(body, '第7条（協議・管轄）', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(
    '本契約に定めのない事項または疑義が生じた場合は、甲乙が誠実に協議して解決する。' +
    '訴訟の必要が生じた場合は、乙の所在地を管轄する地方裁判所を専属的合意管轄裁判所とする。'
  );
  body.appendParagraph('');

  // ---- 署名欄 ----
  body.appendParagraph('');
  body.appendParagraph(
    '本契約の成立を証するため、本書2通を作成し、甲乙それぞれ署名捺印の上、各1通を保有する。'
  );
  appendSignatureBlock_(body, '甲（受入機関）代表者', '乙（登録支援機関）代表者');

  doc.saveAndClose();

  const docUrl = moveDocToContractsFolder_(doc);

  ui.alert('✅ 支援委託契約書を作成しました。\n\nURL: ' + docUrl);
}

// ============================================================
// 4. 生活オリエンテーション確認書
// ============================================================

/**
 * 生活オリエンテーション確認書を Google Doc として生成する。
 * 外国人管理シートのアクティブ行から外国人情報を読み取る。
 * チェックリスト形式で生活上の重要事項を網羅する。
 *
 * ※本テンプレートは雛形です。実際の運用前に専門家（行政書士・社労士等）の
 *   確認を受けてください。
 */
function createLifeOrientationChecklistDoc() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.WORKERS);

  if (!sheet) {
    ui.alert('外国人管理シートが見つかりません。初期化してください。');
    return;
  }

  const row = getActiveDataRow_(sheet);
  if (!row) {
    ui.alert('対象の行を選択してください。');
    return;
  }

  const data = sheet.getRange(row, 1, 1, 20).getValues()[0];
  const workerId    = data[0]  || '';
  const nameRoman   = data[1]  || '';
  const nameNative  = data[2]  || '';
  const nationality = data[4]  || '';
  const companyName = data[10] || '';
  const sector      = data[11] || '';

  const docTitle = `生活オリエンテーション確認書_${workerId}_${nameRoman}`;
  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();
  body.clear();

  // ---- タイトル ----
  const title = body.appendParagraph('生活オリエンテーション確認書');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph(
    '※本書は特定技能外国人への生活オリエンテーション（入管法施行規則第19条の17等）に基づく' +
    '説明実施記録の雛形です。専門家（行政書士・社労士等）の確認を受けてください。'
  ).setItalic(true);
  body.appendParagraph('');

  // ---- 基本情報 ----
  appendHeading_(body, '1. 基本情報', DocumentApp.ParagraphHeading.HEADING2);
  appendInfoTable_(body, [
    ['実施日',           todayJp_() + '　（実施後に記入）'],
    ['実施機関名',       CONFIG.COMPANY.NAME],
    ['登録番号',         CONFIG.COMPANY.REGISTRATION_NUMBER],
    ['担当支援員',       '（　　　　　　　　　）'],
    ['受入企業名',       companyName || '（　　　　　　　　　）'],
    ['外国人氏名（ローマ字）', nameRoman || '（　　　　　　　）'],
    ['外国人氏名（母国語）',   nameNative || '（　　　　　　　）'],
    ['国籍',             nationality || '（　　　　）'],
    ['実施方法',         '対面・オンライン・書面（該当に〇）　所要時間：約（　　）時間以上'],
    ['使用言語',         '（　　　　　　　語）　通訳者：（　　　　　　　）'],
  ]);
  body.appendParagraph('');

  // ---- チェックリスト（テーマ別） ----
  appendHeading_(body, '2. 説明事項チェックリスト', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(
    '各項目を説明し、外国人本人が理解したことを確認してください（☑済・□未説明）。' +
    '生活オリエンテーションは合計8時間以上実施することが求められます。'
  );
  body.appendParagraph('');

  // テーマ別チェック項目
  const themes = [
    {
      heading: '【A】住居・住まいのルール',
      items: [
        ['□', '① 住居の場所・設備',    '居室・水道・電気・ガス・インターネット等の使用方法、緊急時の連絡先'],
        ['□', '② 賃料・共益費の支払い', '支払方法・支払日・滞納時の対応'],
        ['□', '③ 住居内のルール',       '騒音・禁煙・ゴミ置き場・共用部分の利用ルール'],
        ['□', '④ 退去時の手続き',       '原状回復義務・事前連絡期間'],
      ],
    },
    {
      heading: '【B】ゴミの分別・排出ルール',
      items: [
        ['□', '① ゴミの分別方法',   '可燃ゴミ・不燃ゴミ・資源ゴミ・粗大ゴミの分類'],
        ['□', '② ゴミ出しの日程',   '地域のゴミ収集日・収集時間・指定袋の有無'],
        ['□', '③ 禁止事項',         '不法投棄・ルール違反時の行政指導・罰則の説明'],
      ],
    },
    {
      heading: '【C】交通ルール・移動手段',
      items: [
        ['□', '① 道路交通法の基本',  '右側通行禁止・信号遵守・飲酒運転禁止・罰則'],
        ['□', '② 自転車のルール',    'ヘルメット着用推奨・二人乗り禁止・歩道走行ルール'],
        ['□', '③ 運転免許',          '外国免許の国際免許証・日本免許取得手続きの案内'],
        ['□', '④ 公共交通機関',      '電車・バスの乗り方・IC カード・通勤経路'],
      ],
    },
    {
      heading: '【D】医療機関・健康管理',
      items: [
        ['□', '① 医療機関の探し方', '近隣の病院・クリニック・夜間救急・救急車（119番）の呼び方'],
        ['□', '② 健康保険の使い方', '保険証の持参・窓口負担・ジェネリック医薬品'],
        ['□', '③ 薬の入手方法',     '処方箋薬局・市販薬・服薬管理'],
        ['□', '④ 心身の健康管理',   '過労防止・メンタルヘルス・相談窓口（外国人総合相談センター等）'],
      ],
    },
    {
      heading: '【E】災害時の対応',
      items: [
        ['□', '① 地震への備え',     '緊急地震速報の意味・身の守り方・家具固定'],
        ['□', '② 避難場所・避難経路', '最寄りの避難場所・避難指示の意味・ハザードマップ'],
        ['□', '③ 非常用持出し品',   '非常袋の準備・在留カード・パスポートの保管'],
        ['□', '④ 緊急連絡',         '110番（警察）・119番（消防・救急）・171番（災害用伝言ダイヤル）'],
        ['□', '⑤ J-ALERT・防災無線', 'Jアラートの意味・防災無線の案内'],
      ],
    },
    {
      heading: '【F】法令遵守・生活規範',
      items: [
        ['□', '① 在留資格の遵守',   '在留資格外の就労禁止・資格外活動許可の概要'],
        ['□', '② 在留カードの管理', '常時携帯義務・紛失時の届出手続き'],
        ['□', '③ 住所変更届',       '転居後14日以内に市区町村への届出義務'],
        ['□', '④ 税金・社会保険',   '所得税・住民税・年金・健康保険料の納付義務'],
        ['□', '⑤ 犯罪・違反の禁止', '万引き・暴力・迷惑行為等の禁止と在留資格への影響'],
        ['□', '⑥ SNS・個人情報',    'SNS 使用上の注意・個人情報の取り扱い・ネット詐欺'],
      ],
    },
    {
      heading: '【G】相談窓口・サポート体制',
      items: [
        ['□', '① 登録支援機関の連絡先',
          CONFIG.COMPANY.NAME + '  TEL: ' + CONFIG.COMPANY.PHONE + '  Email: ' + CONFIG.COMPANY.EMAIL],
        ['□', '② 行政機関の相談窓口',
          '外国人在留総合インフォメーションセンター（0570-013904）・都道府県労働局'],
        ['□', '③ 外国語対応相談機関',
          '法テラス・外国人総合相談センター・よりそいホットライン（0120-279-338）'],
        ['□', '④ ハラスメント・差別',
          '労働局雇用環境・均等部・各種ハラスメント相談窓口の案内'],
      ],
    },
  ];

  themes.forEach(theme => {
    appendHeading_(body, theme.heading, DocumentApp.ParagraphHeading.HEADING3);
    const tableRows = [['確認', '項目', '説明内容・ポイント']].concat(theme.items);
    appendInfoTable_(body, tableRows);
    body.appendParagraph('');
  });

  // ---- 確認結果 ----
  appendHeading_(body, '3. 実施結果・特記事項', DocumentApp.ParagraphHeading.HEADING2);
  appendInfoTable_(body, [
    ['合計実施時間',     '（　　）時間（　　）分'],
    ['全項目説明完了',   '完了・一部未了（継続実施予定日：　　　　年　　月　　日）'],
    ['本人の理解度',     '十分理解・おおむね理解・再説明が必要（再説明予定日：　　　　年　　月　　日）'],
    ['質問・特記事項',   '（　　　　　　　　　　　　　　　　　　　　　　　　　）'],
    ['次回フォロー予定', '（　　　　年　　月　　日　内容：　　　　　　　　　　）'],
  ]);
  body.appendParagraph('');

  // ---- 署名欄 ----
  appendSignatureBlock_(body, '支援機関担当者', '外国人本人');

  doc.saveAndClose();

  const docUrl = moveDocToContractsFolder_(doc);

  ui.alert('✅ 生活オリエンテーション確認書を作成しました。\n\nURL: ' + docUrl);
}

/*
================================================================================
  Menu.gs への追加事項
================================================================================

  以下のメニュー項目を Menu.gs の適切な位置（「法定書類」サブメニュー等）に
  追加してください。

  ─────────────────────────────────────────
  表示名（メニュー項目）                  → 関数名
  ─────────────────────────────────────────
  '📄 雇用条件書を作成'                  → createEmploymentConditionsDoc
  '📋 事前ガイダンス実施確認書を作成'    → createPreGuidanceConfirmationDoc
  '📝 支援委託契約書を作成'              → createSupportEntrustmentContractDoc
  '🏠 生活オリエンテーション確認書を作成' → createLifeOrientationChecklistDoc
  ─────────────────────────────────────────

  Menu.gs 記述例（サブメニューに追加する場合）:

    .addSubMenu(ui.createMenu('📂 法定書類')
      .addItem('📄 雇用条件書を作成',                   'createEmploymentConditionsDoc')
      .addItem('📋 事前ガイダンス実施確認書を作成',     'createPreGuidanceConfirmationDoc')
      .addItem('📝 支援委託契約書を作成',               'createSupportEntrustmentContractDoc')
      .addItem('🏠 生活オリエンテーション確認書を作成', 'createLifeOrientationChecklistDoc')
    )

================================================================================
  CONFIG.gs への追加提案
================================================================================

  以下のキーを CONFIG.FOLDERS に追加することを推奨します。
  ただし Config.gs は直接編集せず、Setup.gs の createFolderStructure() 内で
  サブフォルダ作成・設定シートへの ID 保存処理も追加してください。

    FOLDERS: {
      // 既存 ...
      DOCUMENTS: '📄 法定書類',   // 法定書類専用フォルダ（任意）
    }

  また、getFolderIdFromSettings_ で参照できるよう、
  createFolderStructure() に以下を追加してください:

    settingsSheet.getRange('A8:B8').setValues([['ContractsFolderID',
      getOrCreateFolder_(root, CONFIG.FOLDERS.CONTRACTS).getId()]]);

  ※ ContractsFolderID は既存の CONFIG.FOLDERS.CONTRACTS (📄 契約書) を
     使用しています。createFolderStructure() で既に CONTRACTS フォルダは
     作成されていますが、設定シートへの ID 保存が行われていないため、
     上記を追加することで moveTo 先が正しく解決されます。

================================================================================
*/
