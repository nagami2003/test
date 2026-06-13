// ===== 通知・自動化トリガー =====

function setupAllTriggers() {
  deleteAllTriggers();

  // 毎朝8時にダッシュボード更新＋ビザ期限チェック
  ScriptApp.newTrigger('dailyMorningCheck')
    .timeBased().atHour(8).everyDays(1).create();

  // 毎月1日に請求書リマインダー
  ScriptApp.newTrigger('monthlyInvoiceReminder')
    .timeBased().onMonthDay(1).atHour(9).create();

  // 毎週月曜に面談リマインダー
  ScriptApp.newTrigger('weeklyInterviewReminder')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();

  SpreadsheetApp.getUi().alert(
    '✅ 自動通知トリガーを設定しました。\n\n' +
    '・毎朝8時: ダッシュボード更新・ビザ期限チェック\n' +
    '・毎月1日9時: 請求書作成リマインダー\n' +
    '・毎週月曜9時: 面談リマインダー'
  );
}

function deleteAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

function dailyMorningCheck() {
  updateDashboard();
  sendVisaExpiryAlerts_();
}

function weeklyInterviewReminder() {
  sendInterviewReminders_();
}

function monthlyInvoiceReminder() {
  sendInvoiceReminder_();
}

function sendVisaExpiryAlerts_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.WORKERS);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const today = new Date();
  const alertDays = CONFIG.NOTIFICATIONS.VISA_EXPIRY_ALERT_DAYS;
  const alerts = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][3] !== '在籍中') continue;
    const expiry = data[i][13];
    if (!expiry) continue;
    const daysLeft = Math.floor((new Date(expiry) - today) / 86400000);
    if (alertDays.includes(daysLeft)) {
      alerts.push({
        name: data[i][1], company: data[i][10], expiry,
        daysLeft, residenceStatus: data[i][12],
      });
    }
  }

  if (alerts.length === 0) return;

  const subject = `【登録支援機関】在留期限アラート: ${alerts.length}名`;
  const body = `
以下の特定技能外国人の在留期限が迫っています。
速やかに更新手続きの確認をお願いします。

${alerts.map(a =>
    `■ ${a.name}（${a.company}）
   在留資格: ${a.residenceStatus}
   期限: ${a.expiry}（あと${a.daysLeft}日）`
  ).join('\n\n')}

---
${CONFIG.COMPANY.NAME}
管理システム 自動通知
  `.trim();

  sendNotificationEmail_(subject, body);
}

function sendInterviewReminders_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.INTERVIEWS);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const today = new Date();
  const reminderDays = CONFIG.NOTIFICATIONS.INTERVIEW_REMINDER_DAYS;
  const upcoming = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][5] !== '未実施') continue;
    const scheduledDate = data[i][4] ? new Date(data[i][4]) : null;
    if (!scheduledDate) continue;
    const daysUntil = Math.floor((scheduledDate - today) / 86400000);
    if (daysUntil >= 0 && daysUntil <= reminderDays) {
      upcoming.push({
        name: data[i][2], company: data[i][3], date: data[i][4], daysUntil,
      });
    }
  }

  if (upcoming.length === 0) return;

  const subject = `【登録支援機関】定期面談リマインダー: ${upcoming.length}件`;
  const body = `
今後${reminderDays}日以内に実施予定の定期面談があります。

${upcoming.map(u =>
    `■ ${u.name}（${u.company}）
   予定日: ${u.date}（あと${u.daysUntil}日）`
  ).join('\n\n')}

面談後はシステムへの記録をお忘れなく。

---
${CONFIG.COMPANY.NAME}
管理システム 自動通知
  `.trim();

  sendNotificationEmail_(subject, body);
}

function sendInvoiceReminder_() {
  const subject = `【登録支援機関】今月の請求書作成リマインダー`;
  const month = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月');
  const body = `
${month}分の請求書作成のお時間です。

スプレッドシートを開き、
「支援機関システム」→「請求・売上」→「一括請求書作成」
を実行してください。

---
${CONFIG.COMPANY.NAME}
管理システム 自動通知
  `.trim();

  sendNotificationEmail_(subject, body);
}

function sendTestNotification() {
  const subject = '【テスト】登録支援機関 管理システム 通知テスト';
  const body = `
このメールは通知設定のテストです。

正常に受信できていれば、メール通知の設定は完了です。

システム情報:
- スプレッドシート: ${SpreadsheetApp.getActiveSpreadsheet().getName()}
- 実行日時: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')}

---
${CONFIG.COMPANY.NAME}
管理システム
  `.trim();

  sendNotificationEmail_(subject, body);
  SpreadsheetApp.getUi().alert('✅ テストメールを送信しました。受信トレイをご確認ください。');
}

function sendNotificationEmail_(subject, body) {
  const recipient = CONFIG.NOTIFICATIONS.NOTIFICATION_EMAIL ||
    Session.getEffectiveUser().getEmail();
  if (!recipient) return;

  GmailApp.sendEmail(recipient, subject, body, {
    name: CONFIG.COMPANY.NAME + ' 管理システム',
  });
}

// 行政報告リマインダー（四半期ごとに出入国在留管理庁への定期報告が必要）
function sendQuarterlyReportReminder() {
  const month = new Date().getMonth() + 1;
  // 1月・4月・7月・10月（四半期末翌月）に送信
  if (![1, 4, 7, 10].includes(month)) return;

  const subject = '【重要】特定技能 四半期定期報告 提出リマインダー';
  const body = `
特定技能外国人に関する四半期定期報告書の提出期限が近づいています。

■ 提出先
出入国在留管理庁（管轄の地方出入国在留管理局）

■ 提出内容
1. 支援実施状況に係る届出（登録支援機関）
2. 受入れ機関による定期報告書（特定技能所属機関）

■ 提出期限
四半期終了後の翌月末日

スプレッドシートから定期面談記録・支援計画の進捗を確認して
報告書を作成してください。

---
${CONFIG.COMPANY.NAME}
管理システム 自動通知
  `.trim();

  sendNotificationEmail_(subject, body);
}
