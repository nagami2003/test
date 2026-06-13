function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏢 支援機関システム')
    .addSubMenu(ui.createMenu('⚙️ 初期設定')
      .addItem('① Driveフォルダ構造を作成', 'createFolderStructure')
      .addItem('② シートを初期化', 'initializeAllSheets')
      .addItem('③ ダッシュボードを更新', 'updateDashboard'))
    .addSeparator()
    .addSubMenu(ui.createMenu('👤 外国人管理')
      .addItem('外国人を新規登録', 'showAddWorkerDialog')
      .addItem('ビザ期限チェック', 'checkVisaExpiry')
      .addItem('支援対象者一覧を更新', 'refreshWorkerList'))
    .addSubMenu(ui.createMenu('🏢 受入企業管理')
      .addItem('企業を新規登録', 'showAddCompanyDialog')
      .addItem('契約期限チェック', 'checkContractExpiry'))
    .addSubMenu(ui.createMenu('📋 支援計画')
      .addItem('支援計画書を作成（Google Doc）', 'createSupportPlanDoc')
      .addItem('支援進捗を一括更新', 'updateSupportProgress'))
    .addSubMenu(ui.createMenu('📅 定期面談')
      .addItem('面談を記録', 'showInterviewDialog')
      .addItem('面談スケジュールをカレンダーに追加', 'scheduleInterviews')
      .addItem('未実施面談を確認', 'checkPendingInterviews'))
    .addSubMenu(ui.createMenu('💴 請求・売上')
      .addItem('請求書を作成（Google Doc）', 'createInvoiceDoc')
      .addItem('一括請求書作成', 'createBulkInvoices')
      .addItem('未払い一覧を確認', 'showUnpaidInvoices')
      .addItem('売上サマリーを更新', 'updateSalesSummary'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔔 通知・自動化')
      .addItem('定期通知トリガーを設定', 'setupAllTriggers')
      .addItem('トリガーを削除', 'deleteAllTriggers')
      .addItem('通知テスト送信', 'sendTestNotification'))
    .addSubMenu(ui.createMenu('📂 法定書類')
      .addItem('雇用条件書を作成', 'createEmploymentConditionsDoc')
      .addItem('事前ガイダンス実施確認書を作成', 'createPreGuidanceConfirmationDoc')
      .addItem('支援委託契約書を作成', 'createSupportEntrustmentContractDoc')
      .addItem('生活オリエンテーション確認書を作成', 'createLifeOrientationChecklistDoc'))
    .addSubMenu(ui.createMenu('🌐 多言語対応')
      .addItem('ビザ期限リマインダー（本人へ多言語送信）', 'sendMultilingualVisaReminder')
      .addItem('生活案内プレビュー（選択行）', 'showMultilingualGuide'))
    .addSubMenu(ui.createMenu('📋 行政コンプライアンス')
      .addItem('届出管理シートを初期化', 'setupComplianceSheet')
      .addItem('届出スケジュールを自動生成', 'generateComplianceSchedule')
      .addItem('届出期限をチェック', 'checkComplianceDeadlines')
      .addItem('四半期届出チェックリストを作成', 'createQuarterlyReportChecklistDoc')
      .addItem('随時届出ガイドを作成', 'createAdHocNotificationGuideDoc'))
    .addSubMenu(ui.createMenu('📊 経営分析')
      .addItem('経営分析シートを更新', 'refreshAnalytics')
      .addItem('グラフを再生成', 'insertCharts')
      .addItem('月次経営レポートを作成', 'createMonthlyReportDoc'))
    .addSeparator()
    .addItem('❓ ヘルプ・操作説明', 'showHelp')
    .addToUi();
}

function showHelp() {
  const html = HtmlService.createHtmlOutput(`
    <h2>登録支援機関 管理システム - 操作ガイド</h2>
    <h3>初回セットアップ手順</h3>
    <ol>
      <li><b>Config.gs を編集</b> → 自社情報・銀行口座等を入力</li>
      <li><b>初期設定 → Driveフォルダ構造を作成</b> → フォルダが自動生成されます</li>
      <li><b>初期設定 → シートを初期化</b> → 各管理シートが作成されます</li>
      <li><b>通知・自動化 → 定期通知トリガーを設定</b> → 自動アラートが有効になります</li>
    </ol>
    <h3>日常業務フロー</h3>
    <ul>
      <li>外国人受入時：外国人管理 → 新規登録 → 支援計画書を作成</li>
      <li>毎月：請求書を作成 → 送付</li>
      <li>3ヶ月毎：定期面談を記録 → カレンダーに次回を追加</li>
      <li>四半期毎：出入国在留管理庁へ定期報告書を提出</li>
    </ul>
    <p><i>Config.gs の COMPANY セクションを必ず自社情報に変更してください。</i></p>
  `).setWidth(500).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, '操作ガイド');
}
