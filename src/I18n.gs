// ===== 多言語対応（外国人本人向け通知・案内） =====
// 対応言語: ja（日本語）/ en（英語）/ vi（ベトナム語）/ id（インドネシア語）

// -----------------------------------------------------------------------
// 1. 多言語辞書
// -----------------------------------------------------------------------

/**
 * LANG_MESSAGES
 * プレースホルダ:
 *   {date}  ... 日付文字列（例: 2025/03/31）
 *   {days}  ... 残日数（例: 30）
 *   {phone} ... 電話番号
 *   {email} ... メールアドレス
 */
var LANG_MESSAGES = {

  // -------- 日本語 --------
  ja: {
    greeting: [
      'ようこそ！{company}によるサポートが始まりました。',
      'わからないことがあれば、いつでも相談してください。',
    ].join('\n'),

    visaExpiryReminder: [
      '【重要】在留期限のお知らせ',
      'あなたの在留期限は {date} です（あと {days} 日）。',
      '期限が切れる前に、必ず更新手続きを行ってください。',
      'ご不明な点は担当者にお問い合わせください。',
    ].join('\n'),

    interviewInvitation: [
      '【お知らせ】定期面談のご案内',
      '{date} に定期面談を予定しています。',
      '日時・場所の詳細は担当者からご連絡します。',
    ].join('\n'),

    consultationContact: [
      '【相談窓口のご案内】',
      '困ったことや不安なことがあれば、いつでも連絡してください。',
      '電話: {phone}',
      'メール: {email}',
      '（平日 9:00〜18:00 対応）',
    ].join('\n'),

    emergencyContact: [
      '【緊急連絡先】',
      '命に関わる緊急事態: 119（救急・消防）',
      '犯罪・事故: 110（警察）',
      '外国語対応（24時間）: #7119（一部地域）',
    ].join('\n'),
  },

  // -------- English --------
  en: {
    greeting: [
      'Welcome! Your support by {company} has started.',
      'Please feel free to contact us anytime if you have questions.',
    ].join('\n'),

    visaExpiryReminder: [
      '[IMPORTANT] Residence Card Expiry Notice',
      'Your residence permit will expire on {date} ({days} days remaining).',
      'Please complete the renewal procedures before the expiry date.',
      'Contact your support officer if you need help.',
    ].join('\n'),

    interviewInvitation: [
      '[Notice] Regular Interview Invitation',
      'A regular interview is scheduled on {date}.',
      'Your support officer will inform you of the details.',
    ].join('\n'),

    consultationContact: [
      '[Support Contact]',
      'If you have any concerns or difficulties, please contact us anytime.',
      'Phone: {phone}',
      'Email: {email}',
      '(Available weekdays, 9:00 AM - 6:00 PM)',
    ].join('\n'),

    emergencyContact: [
      '[Emergency Contacts]',
      'Life-threatening emergency: 119 (Ambulance / Fire)',
      'Crime or accident: 110 (Police)',
      'Multilingual support (24h): #7119 (available in some areas)',
    ].join('\n'),
  },

  // -------- Tieng Viet (Vietnamese) --------
  vi: {
    greeting: [
      'Xin chao! Chuong trinh ho tro cua {company} da bat dau.',
      // Note: 声調記号を含む正式版を下に記載
      'Chào mừng bạn! Chương trình hỗ trợ của {company} đã bắt đầu.',
      'Nếu bạn có bất kỳ thắc mắc nào, hãy liên hệ với chúng tôi bất cứ lúc nào.',
    ].join('\n'),

    visaExpiryReminder: [
      '[QUAN TRỌNG] Thông báo hạn thẻ cư trú',
      'Thẻ cư trú của bạn sẽ hết hạn vào ngày {date} (còn {days} ngày).',
      'Vui lòng làm thủ tục gia hạn trước khi thẻ hết hạn.',
      'Nếu bạn cần hỗ trợ, hãy liên hệ với người phụ trách.',
    ].join('\n'),

    interviewInvitation: [
      '[Thông báo] Lịch phỏng vấn định kỳ',
      'Buổi phỏng vấn định kỳ được lên lịch vào ngày {date}.',
      'Người phụ trách sẽ thông báo chi tiết về thời gian và địa điểm.',
    ].join('\n'),

    consultationContact: [
      '[Thông tin liên hệ hỗ trợ]',
      'Nếu bạn gặp khó khăn hoặc lo lắng về bất cứ điều gì, hãy liên hệ với chúng tôi.',
      'Điện thoại: {phone}',
      'Email: {email}',
      '(Làm việc từ Thứ Hai đến Thứ Sáu, 9:00 - 18:00)',
    ].join('\n'),

    emergencyContact: [
      '[Liên hệ khẩn cấp]',
      'Cấp cứu / Cháy nổ: 119',
      'Tội phạm / Tai nạn: 110 (Cảnh sát)',
      'Hỗ trợ đa ngôn ngữ (24 giờ): #7119 (một số khu vực)',
    ].join('\n'),
  },

  // -------- Bahasa Indonesia --------
  id: {
    greeting: [
      'Selamat datang! Program dukungan dari {company} telah dimulai.',
      'Jangan ragu untuk menghubungi kami kapan saja jika Anda memiliki pertanyaan.',
    ].join('\n'),

    visaExpiryReminder: [
      '[PENTING] Pemberitahuan Masa Berlaku Izin Tinggal',
      'Izin tinggal Anda akan habis masa berlakunya pada {date} ({days} hari lagi).',
      'Harap lakukan prosedur perpanjangan sebelum tanggal kedaluwarsa.',
      'Hubungi petugas pendamping Anda jika membutuhkan bantuan.',
    ].join('\n'),

    interviewInvitation: [
      '[Pemberitahuan] Undangan Wawancara Rutin',
      'Wawancara rutin dijadwalkan pada {date}.',
      'Petugas pendamping akan memberitahu Anda mengenai detail waktu dan tempat.',
    ].join('\n'),

    consultationContact: [
      '[Kontak Layanan Konsultasi]',
      'Jika Anda mengalami kesulitan atau kekhawatiran, silakan hubungi kami kapan saja.',
      'Telepon: {phone}',
      'Email: {email}',
      '(Tersedia hari kerja, pukul 09.00 - 18.00)',
    ].join('\n'),

    emergencyContact: [
      '[Kontak Darurat]',
      'Kedaruratan medis / Kebakaran: 119',
      'Kejahatan / Kecelakaan: 110 (Polisi)',
      'Layanan multibahasa (24 jam): #7119 (tersedia di beberapa wilayah)',
    ].join('\n'),
  },
};


// -----------------------------------------------------------------------
// 2. 国籍 → 言語コード変換
// -----------------------------------------------------------------------

/**
 * nationalityToLang_
 * CONFIG.NATIONALITIES の国名文字列から言語コードを返す。
 * 日本語は常に日本語で対応するため 'ja' は返さない（外国人向け）。
 * 不明・その他は 'en'（共通語として英語）を返す。
 *
 * @param {string} nationality - 国籍名（CONFIG.NATIONALITIES の値）
 * @returns {string} 言語コード: 'vi' | 'id' | 'en'
 */
function nationalityToLang_(nationality) {
  var map = {
    'ベトナム':       'vi',
    'インドネシア':   'id',
    // 以下は英語を共通語として使用
    'フィリピン':     'en',
    'タイ':           'en',
    'ミャンマー':     'en',
    'カンボジア':     'en',
    'モンゴル':       'en',
    '中国':           'en',
    'ネパール':       'en',
    'インド':         'en',
    'スリランカ':     'en',
    'パキスタン':     'en',
    'ウズベキスタン': 'en',
    'その他':         'en',
  };
  return map[nationality] || 'en';
}


// -----------------------------------------------------------------------
// 3. メッセージ取得・プレースホルダ置換
// -----------------------------------------------------------------------

/**
 * getMessage_
 * 辞書からメッセージを取得し、プレースホルダを実値で置換して返す。
 *
 * @param {string} key    - メッセージ種別キー（'greeting' など）
 * @param {string} lang   - 言語コード（'ja' | 'en' | 'vi' | 'id'）
 * @param {Object} params - プレースホルダの置換値 例: { date: '2025/03/31', days: '30' }
 * @returns {string} 置換済みメッセージ。辞書に存在しない場合は空文字。
 */
function getMessage_(key, lang, params) {
  var safeLang = LANG_MESSAGES[lang] ? lang : 'en';
  var template = LANG_MESSAGES[safeLang][key];
  if (!template) return '';

  var result = template;
  if (params) {
    Object.keys(params).forEach(function(k) {
      // {key} をすべて置換（グローバル）
      result = result.split('{' + k + '}').join(String(params[k]));
    });
  }
  return result;
}


// -----------------------------------------------------------------------
// 4. 在留期限リマインダー一括送信
// -----------------------------------------------------------------------

/**
 * sendMultilingualVisaReminder
 * 外国人管理シートを走査し、在留期限が 90 日以内かつ在籍中の外国人に
 * Email で現地語＋日本語併記の在留期限リマインダーを送信する。
 * Email が空欄の行はスキップ。送信件数を alert で報告する。
 *
 * シート列構成（0始まりインデックス）:
 *   0=ID, 1=氏名ローマ字, 2=氏名母国語, 3=在籍状況, 4=国籍,
 *   5=生年月日, 6=性別, 7=電話, 8=Email, 9=会社ID, 10=会社名,
 *   11=分野, 12=在留資格, 13=在留期限
 */
function sendMultilingualVisaReminder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.WORKERS);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('エラー: 外国人管理シートが見つかりません。');
    return;
  }

  var data = sheet.getDataRange().getValues();
  var today = new Date();
  var sentCount = 0;
  var skippedCount = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    // ID が空 → データ終端とみなす
    if (!row[0]) break;

    // 在籍中のみ対象
    if (row[3] !== '在籍中') continue;

    // 在留期限チェック
    var expiry = row[13];
    if (!expiry) continue;
    var expiryDate = new Date(expiry);
    var daysLeft = Math.floor((expiryDate - today) / 86400000);
    if (daysLeft > 90 || daysLeft < 0) continue;

    // Email チェック
    var email = row[8];
    if (!email || String(email).trim() === '') {
      skippedCount++;
      continue;
    }

    var nameRoman   = row[1] || '';
    var nationality = row[4] || 'その他';
    var lang        = nationalityToLang_(nationality);
    var expiryStr   = Utilities.formatDate(expiryDate, 'Asia/Tokyo', 'yyyy/MM/dd');
    var params      = { date: expiryStr, days: String(daysLeft), company: CONFIG.COMPANY.NAME };

    // 現地語メッセージ
    var localMsg = getMessage_('visaExpiryReminder', lang, params);
    // 日本語メッセージ（常に併記）
    var jaMsg    = getMessage_('visaExpiryReminder', 'ja', params);

    // 本文組み立て（現地語 → 日本語の順）
    var body = [
      '--- ' + lang.toUpperCase() + ' ---',
      localMsg,
      '',
      '--- 日本語 ---',
      jaMsg,
      '',
      '---',
      CONFIG.COMPANY.NAME,
      'TEL: ' + CONFIG.COMPANY.PHONE,
      'EMAIL: ' + CONFIG.COMPANY.EMAIL,
    ].join('\n');

    // 件名も現地語＋日本語
    var subjectMap = {
      vi: '[Quan trong] Thong bao han the cu tru / 【重要】在留期限のお知らせ',
      id: '[Penting] Pemberitahuan Izin Tinggal / 【重要】在留期限のお知らせ',
      en: '[Important] Residence Permit Expiry Notice / 【重要】在留期限のお知らせ',
    };
    var subject = subjectMap[lang] || subjectMap['en'];

    GmailApp.sendEmail(
      String(email).trim(),
      subject,
      body,
      { name: CONFIG.COMPANY.NAME + ' サポート窓口' }
    );
    sentCount++;
  }

  var report = [
    '多言語ビザ期限リマインダー送信完了',
    '送信件数: ' + sentCount + ' 件',
  ];
  if (skippedCount > 0) {
    report.push('Email未登録のためスキップ: ' + skippedCount + ' 件');
  }
  SpreadsheetApp.getUi().alert(report.join('\n'));
}


// -----------------------------------------------------------------------
// 5. 選択行の多言語生活案内モーダル表示
// -----------------------------------------------------------------------

/**
 * showMultilingualGuide
 * アクティブシートの選択行から外国人を特定し、
 * HtmlService モーダルで生活案内（挨拶・相談窓口・緊急連絡先）を
 * 現地語＋日本語で表示する。
 * アクティブ行が無効（1行目ヘッダーまたはID空欄）なら案内して return。
 */
function showMultilingualGuide() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getActiveSheet();

  // 外国人管理シート上でのみ動作
  if (sheet.getName() !== CONFIG.SHEETS.WORKERS) {
    SpreadsheetApp.getUi().alert(
      '「' + CONFIG.SHEETS.WORKERS + '」シートを開いた状態で実行してください。'
    );
    return;
  }

  var activeRow = sheet.getActiveCell().getRow();
  if (activeRow <= 1) {
    SpreadsheetApp.getUi().alert('外国人データの行（2行目以降）を選択してから実行してください。');
    return;
  }

  var row = sheet.getRange(activeRow, 1, 1, 14).getValues()[0];
  if (!row[0]) {
    SpreadsheetApp.getUi().alert('選択した行にデータがありません。有効な行を選択してください。');
    return;
  }

  var nameRoman   = row[1] || '（氏名未入力）';
  var nationality = row[4] || 'その他';
  var lang        = nationalityToLang_(nationality);
  var params      = {
    company: CONFIG.COMPANY.NAME,
    phone:   CONFIG.COMPANY.PHONE,
    email:   CONFIG.COMPANY.EMAIL,
  };

  // 各メッセージを現地語・日本語で取得
  var keys  = ['greeting', 'consultationContact', 'emergencyContact'];
  var langLabelMap = { vi: 'Tiếng Việt', id: 'Bahasa Indonesia', en: 'English', ja: '日本語' };
  var langLabel = langLabelMap[lang] || lang.toUpperCase();

  var blocks = keys.map(function(key) {
    var local = getMessage_(key, lang, params);
    var ja    = getMessage_(key, 'ja', params);
    // 同じ言語なら重複しない
    var content = (lang === 'ja')
      ? '<p>' + escapeHtml_(local) + '</p>'
      : '<p>' + escapeHtml_(local) + '</p><hr class="lang-sep"><p class="ja-text">' + escapeHtml_(ja) + '</p>';
    return content;
  });

  var html = buildGuideHtml_(nameRoman, nationality, langLabel, blocks);
  var output = HtmlService.createHtmlOutput(html)
    .setWidth(560)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(output, nameRoman + ' さんへの生活案内');
}

/**
 * escapeHtml_
 * HTML 特殊文字をエスケープする。
 * @param {string} str
 * @returns {string}
 */
function escapeHtml_(str) {
  return String(str)
    .replace(/&/g,  '&amp;')
    .replace(/</g,  '&lt;')
    .replace(/>/g,  '&gt;')
    .replace(/"/g,  '&quot;')
    .replace(/\n/g, '<br>');
}

/**
 * buildGuideHtml_
 * 生活案内モーダル用の HTML 文字列を組み立てる。
 * @param {string}   nameRoman   - 氏名ローマ字
 * @param {string}   nationality - 国籍
 * @param {string}   langLabel   - 言語表示名
 * @param {string[]} blocks      - コンテンツブロック（HTML 文字列配列）
 * @returns {string} HTML
 */
function buildGuideHtml_(nameRoman, nationality, langLabel, blocks) {
  return '<!DOCTYPE html>\n' +
    '<html>\n' +
    '<head>\n' +
    '<meta charset="UTF-8">\n' +
    '<style>\n' +
    '  body { font-family: "Noto Sans", "Noto Sans JP", sans-serif; font-size: 13px;\n' +
    '         padding: 12px 16px; color: #202124; line-height: 1.6; }\n' +
    '  h2 { color: #1a73e8; font-size: 15px; border-bottom: 2px solid #1a73e8;\n' +
    '       padding-bottom: 4px; margin-bottom: 4px; }\n' +
    '  .meta { color: #5f6368; font-size: 12px; margin-bottom: 14px; }\n' +
    '  .section { background: #f8f9fa; border-left: 4px solid #1a73e8;\n' +
    '             padding: 10px 12px; margin-bottom: 12px; border-radius: 0 4px 4px 0; }\n' +
    '  .section p { margin: 0 0 6px; }\n' +
    '  hr.lang-sep { border: none; border-top: 1px dashed #dadce0; margin: 8px 0; }\n' +
    '  .ja-text { color: #5f6368; font-size: 12px; }\n' +
    '</style>\n' +
    '</head>\n' +
    '<body>\n' +
    '<h2>' + escapeHtml_(nameRoman) + ' さん 生活案内</h2>\n' +
    '<p class="meta">国籍: ' + escapeHtml_(nationality) + ' &nbsp;|&nbsp; 言語: ' + escapeHtml_(langLabel) + '</p>\n' +
    blocks.map(function(b) { return '<div class="section">' + b + '</div>'; }).join('\n') +
    '</body>\n</html>';
}


// -----------------------------------------------------------------------
// Menu.gs への追加項目（このファイルは編集不要。担当者が Menu.gs に追加してください）
// -----------------------------------------------------------------------
//
// 追加すべきメニュー項目（表示名 → 関数名）:
//
//   「多言語ビザ期限リマインダー送信」  → sendMultilingualVisaReminder
//   「多言語生活案内を表示（選択行）」  → showMultilingualGuide
//
// 推奨: 既存メニューに「多言語対応」サブメニューを追加し、以下のように設定する:
//
//   .addSubMenu(ui.createMenu('多言語対応')
//     .addItem('ビザ期限リマインダー（本人へ多言語送信）', 'sendMultilingualVisaReminder')
//     .addItem('生活案内プレビュー（選択行）', 'showMultilingualGuide')
//   )
//
// -----------------------------------------------------------------------
