// ===== 登録支援機関 Google Workspace 管理システム =====
// このファイルを自社情報に合わせて編集してください

const CONFIG = {
  // ---------- 自社情報（要変更） ----------
  COMPANY: {
    NAME: '株式会社〇〇登録支援機関',
    POSTAL_CODE: '〒000-0000',
    ADDRESS: '東京都〇〇区〇〇1-1-1',
    PHONE: '03-0000-0000',
    EMAIL: 'support@example.com',
    REGISTRATION_NUMBER: '登録番号 第〇〇〇〇〇〇〇〇〇号',
    BANK_NAME: '〇〇銀行 〇〇支店',
    ACCOUNT_TYPE: '普通',
    ACCOUNT_NUMBER: '0000000',
    ACCOUNT_HOLDER: '株式会社〇〇登録支援機関',
  },

  // ---------- シート名 ----------
  SHEETS: {
    DASHBOARD: 'ダッシュボード',
    WORKERS: '外国人管理',
    COMPANIES: '受入企業管理',
    SUPPORT_PLANS: '支援計画',
    INTERVIEWS: '定期面談記録',
    INVOICES: '請求書管理',
    SALES: '売上管理',
    SETTINGS: '設定',
    COMPLIANCE: 'コンプライアンス届出管理',
    ANALYTICS: '経営分析',
  },

  // ---------- Drive フォルダ名 ----------
  FOLDERS: {
    ROOT: '📁 登録支援機関',
    WORKERS: '👤 外国人別',
    COMPANIES: '🏢 受入企業別',
    CONTRACTS: '📄 契約書',
    REPORTS: '📊 報告書',
    INVOICES: '💴 請求書',
    TEMPLATES: '📋 テンプレート',
    CORRESPONDENCE: '✉️ 往復文書',
  },

  // ---------- 特定技能12種類支援 ----------
  SUPPORT_TYPES: [
    '① 事前ガイダンス',
    '② 出入国時の送迎',
    '③ 住居確保・生活必需品契約支援',
    '④ 生活オリエンテーション',
    '⑤ 日本語学習機会の提供',
    '⑥ 相談・苦情対応',
    '⑦ 日本人との交流促進',
    '⑧ 転職支援',
    '⑨ 定期面談・行政機関への通報',
    '⑩ 入居支援',
    '⑪ 医療機関・交通機関等の案内',
    '⑫ 行政手続き支援',
  ],

  // ---------- 在留資格 ----------
  RESIDENCE_STATUS: ['特定技能1号', '特定技能2号'],

  // ---------- 国籍リスト（特定技能対象国） ----------
  NATIONALITIES: [
    'ベトナム', 'フィリピン', 'インドネシア', 'タイ', 'ミャンマー',
    'カンボジア', 'モンゴル', '中国', 'ネパール', 'インド',
    'スリランカ', 'パキスタン', 'ウズベキスタン', 'その他',
  ],

  // ---------- 分野（特定技能対象分野） ----------
  SECTORS: [
    '介護', 'ビルクリーニング', '素形材・産業機械・電気電子情報関連製造業',
    '建設', '造船・舶用工業', '自動車整備', '航空', '宿泊',
    '農業', '漁業', '飲食料品製造業', '外食業',
  ],

  // ---------- 請求設定 ----------
  INVOICE: {
    PAYMENT_TERMS: 30,        // 支払期限（発行日から何日後）
    TAX_RATE: 0.10,           // 消費税率
    MONTHLY_FEE: 20000,       // 月次支援費用（デフォルト）
    CURRENCY: 'JPY',
  },

  // ---------- 通知設定 ----------
  NOTIFICATIONS: {
    VISA_EXPIRY_ALERT_DAYS: [90, 60, 30, 14], // ビザ期限前アラート（日数）
    INTERVIEW_REMINDER_DAYS: 14,               // 面談リマインダー（何日前）
    REPORT_REMINDER_DAY: 1,                    // 四半期報告リマインダー（月初何日目）
    NOTIFICATION_EMAIL: '',                    // 通知先メールアドレス（空欄=実行者）
  },

  // ---------- カラー設定 ----------
  COLORS: {
    HEADER: '#1a73e8',
    HEADER_TEXT: '#ffffff',
    ALERT_RED: '#ea4335',
    ALERT_YELLOW: '#fbbc04',
    ALERT_GREEN: '#34a853',
    ROW_ALT: '#f8f9fa',
    BORDER: '#dadce0',
  },
};
