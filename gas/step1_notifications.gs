/**
 * Step 1: 通知の自動化(メール通知版)
 *
 * setup.gs と同じApps Scriptプロジェクトに新しいファイルとして追加する
 * (シート名などの定数は setup.gs のものを共用)。
 *
 * 方針: 通知は「やることがあるときだけ」届く。
 *   - 毎月1日 9時台: 当月のチェック行を自動生成(これ自体は通知しない)。
 *     更新日が60日以内に迫った契約があればメールで知らせる
 *   - 毎月8日 9時台: 「未確認」が残っていればメールでリマインド
 *
 * 送信先: スクリプト実行者(トリガー登録者)のメールアドレス。
 *   変えたい場合はスクリプト プロパティ NOTIFY_EMAIL に宛先を設定
 *   (カンマ区切りで複数指定も可)。
 *
 * 導入手順:
 *   1. このファイルを追加して保存
 *   2. testEmailNotification を実行して届くことを確認(初回は権限承認あり)
 *   3. setupTriggers を実行してトリガーを登録
 */

const PROP_NOTIFY_EMAIL = 'NOTIFY_EMAIL';
const REMIND_DAY = 8;          // 未確認リマインドを送る日
const RENEWAL_ALERT_DAYS = 60; // 更新日アラートの対象範囲(日)

// ---------------------------------------------------------------------------
// トリガー登録
// ---------------------------------------------------------------------------

/** 月次トリガーを登録する(再実行しても重複しない) */
function setupTriggers() {
  const managed = ['runMonthlyStart', 'runUncheckedReminder'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (managed.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('runMonthlyStart').timeBased().onMonthDay(1).atHour(9).create();
  ScriptApp.newTrigger('runUncheckedReminder').timeBased().onMonthDay(REMIND_DAY).atHour(9).create();
  Logger.log('トリガーを登録しました: 毎月1日9時台(行生成+更新日アラート) / 毎月' + REMIND_DAY + '日9時台(未確認リマインド)');
}

// ---------------------------------------------------------------------------
// 毎月1日: チェック行の自動生成 + 更新日アラート
// ---------------------------------------------------------------------------

function runMonthlyStart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const created = createMonthlyLogRows_(ss);

  if (created === null) {
    // 当月分は既に作成済み(トリガーの重複起動など)。アラートの二重送信も避ける
    Logger.log('当月分のチェック行は作成済みのためスキップしました。');
    return;
  }
  Logger.log('当月のチェック行を ' + created.length + '件 作成しました: ' + created.join(', '));

  const renewals = getUpcomingRenewals_(ss, RENEWAL_ALERT_DAYS);
  if (renewals.length === 0) return;

  const lines = [];
  lines.push(RENEWAL_ALERT_DAYS + '日以内に更新日が来る契約があります。継続/解約を確認してください。');
  lines.push('');
  renewals.forEach(function (r) {
    lines.push('・' + r.tool + (r.plan ? '(' + r.plan + ')' : '') + ' — ' + r.dateText + (r.overdue ? ' ※期限超過' : ''));
  });
  lines.push('');
  lines.push('シート: ' + ss.getUrl());
  sendNotification_('[SaaS管理] 更新日が近い契約があります', lines.join('\n'));
}

/**
 * 当月のチェック行を月次チェックログに追加する。
 * 行構成(Slack(管理者)などの区分行)は前月の行をそのまま引き継ぐ。
 * 前月の行がない場合はツールマスタのチェック対象ツールから生成する。
 *
 * @return {string[]|null} 作成した行のツール名。当月分が既にあれば null
 */
function createMonthlyLogRows_(ss) {
  const sheet = ss.getSheetByName(SHEET_LOG);
  const tz = ss.getSpreadsheetTimeZone();
  const now = new Date();
  const thisMonth = Utilities.formatDate(now, tz, 'yyyy-MM');
  const prevMonth = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth() - 1, 1), tz, 'yyyy-MM');

  // D・E列は数式が先置きされているため、getDataRange はシート下端まで届く。
  // 実データの最終行はA列(年月)で判定する
  const data = sheet.getDataRange().getValues();
  let lastRow = 1;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] !== '') { lastRow = i + 1; break; }
  }

  if (data.some(function (r) { return r[0] === thisMonth; })) return null;

  let names = data
    .filter(function (r) { return r[0] === prevMonth; })
    .map(function (r) { return r[1]; });

  if (names.length === 0) {
    // フォールバック: マスタの月次チェック対象(解約済を除く)からユニークなツール名を取得
    const master = ss.getSheetByName(SHEET_MASTER).getDataRange().getValues();
    const seen = {};
    master.slice(1).forEach(function (r) {
      if (r[12] === true && r[4] !== '解約済' && r[0] !== '') seen[r[0]] = true;
    });
    names = Object.keys(seen);
  }
  if (names.length === 0) return null;

  // D・E列(前月値・増減)の数式を消さないよう、A〜C列とF〜I列を分けて書き込む
  const startRow = lastRow + 1;
  sheet.getRange(startRow, 1, names.length, 3).setValues(
    names.map(function (n) { return [thisMonth, n, '']; })
  );
  sheet.getRange(startRow, 6, names.length, 4).setValues(
    names.map(function () { return ['未確認', '', '', '']; })
  );
  return names;
}

/** 次回更新日が指定日数以内(期限超過含む)の契約を返す */
function getUpcomingRenewals_(ss, days) {
  const master = ss.getSheetByName(SHEET_MASTER).getDataRange().getValues();
  const tz = ss.getSpreadsheetTimeZone();
  const today = new Date();
  const limit = new Date(today.getTime() + days * 24 * 3600 * 1000);

  return master.slice(1)
    .filter(function (r) {
      return r[11] instanceof Date && r[11] <= limit && r[4] !== '解約済';
    })
    .map(function (r) {
      return {
        tool: r[0],
        plan: r[1],
        dateText: Utilities.formatDate(r[11], tz, 'yyyy/MM/dd'),
        overdue: r[11] < today
      };
    });
}

// ---------------------------------------------------------------------------
// 毎月8日: 未確認リマインド
// ---------------------------------------------------------------------------

function runUncheckedReminder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  const tz = ss.getSpreadsheetTimeZone();
  const thisMonth = Utilities.formatDate(new Date(), tz, 'yyyy-MM');

  const unchecked = sheet.getDataRange().getValues()
    .filter(function (r) { return r[0] === thisMonth && r[5] === '未確認'; })
    .map(function (r) { return r[1]; });

  if (unchecked.length === 0) {
    Logger.log('未確認はありません。リマインドは送信しません。');
    return;
  }

  const lines = [];
  lines.push(thisMonth + ' 分で未確認の月次チェックが ' + unchecked.length + '件 残っています:');
  lines.push('');
  lines.push(unchecked.map(function (n) { return '・' + n; }).join('\n'));
  lines.push('');
  lines.push('確認したら「' + SHEET_LOG + '」の結果列をOK/要対応に更新してください。');
  lines.push('シート: ' + ss.getUrl());
  sendNotification_('[SaaS管理] 未確認の月次チェックが残っています(' + thisMonth + ')', lines.join('\n'));
}

// ---------------------------------------------------------------------------
// メール送信
// ---------------------------------------------------------------------------

/** メールで通知する。宛先はNOTIFY_EMAILプロパティ、未設定ならスクリプト実行者 */
function sendNotification_(subject, body) {
  const to = PropertiesService.getScriptProperties().getProperty(PROP_NOTIFY_EMAIL) ||
    Session.getEffectiveUser().getEmail();
  if (!to) {
    Logger.log('送信先メールアドレスが取得できませんでした。通知内容:\n' + subject + '\n' + body);
    return false;
  }
  MailApp.sendEmail(to, subject, body);
  Logger.log('メールを送信しました: ' + to + ' / ' + subject);
  return true;
}

/** 疎通確認用: 自分にテストメールを送る */
function testEmailNotification() {
  sendNotification_('[SaaS管理] 通知テスト',
    'このメールが届いていれば通知設定は完了です。\n' +
    '通知が来るのは「未確認チェックが残っているとき(毎月8日)」と「更新日が60日以内の契約があるとき(毎月1日)」だけです。');
}
