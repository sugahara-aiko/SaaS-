/**
 * Step 1: 通知の自動化
 *
 * setup.gs と同じApps Scriptプロジェクトに新しいファイルとして追加する
 * (シート名などの定数は setup.gs のものを共用)。
 *
 * 導入手順:
 *   1. Slack Incoming WebhookのURLを「プロジェクトの設定 > スクリプト プロパティ」の
 *      キー SLACK_WEBHOOK_URL に設定する(コードには直書きしない)
 *   2. testSlackNotification を実行して疎通確認(初回は権限承認あり)
 *   3. setupTriggers を実行してトリガーを登録する
 *
 * 動作:
 *   - 毎月1日 9時台: 当月のチェック行を自動生成 + 開始通知 + 更新日アラート
 *   - 毎月8日 9時台: 「未確認」が残っていればリマインド(なければ何もしない)
 */

const PROP_SLACK_WEBHOOK = 'SLACK_WEBHOOK_URL';
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
  Logger.log('トリガーを登録しました: 毎月1日9時台(行生成+開始通知+更新日アラート) / 毎月' + REMIND_DAY + '日9時台(未確認リマインド)');
}

// ---------------------------------------------------------------------------
// 毎月1日: チェック行の自動生成 + 開始通知 + 更新日アラート
// ---------------------------------------------------------------------------

function runMonthlyStart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const created = createMonthlyLogRows_(ss);

  if (created === null) {
    // 当月分は既に作成済み(トリガーの重複起動など)。通知も送らない
    Logger.log('当月分のチェック行は作成済みのためスキップしました。');
    return;
  }

  const tz = ss.getSpreadsheetTimeZone();
  const thisMonth = Utilities.formatDate(new Date(), tz, 'yyyy-MM');

  const lines = [];
  lines.push('📋 *[SaaS管理] ' + thisMonth + ' の月次チェックを開始します*');
  lines.push('チェック対象 ' + created.length + '件の行を「' + SHEET_LOG + '」に作成しました:');
  lines.push(created.map(function (n) { return '・' + n; }).join('\n'));

  const renewals = getUpcomingRenewals_(ss, RENEWAL_ALERT_DAYS);
  if (renewals.length > 0) {
    lines.push('');
    lines.push('🔔 *' + RENEWAL_ALERT_DAYS + '日以内に更新日が来る契約:*');
    renewals.forEach(function (r) {
      lines.push('・' + r.tool + (r.plan ? '(' + r.plan + ')' : '') + ' — ' + r.dateText + (r.overdue ? ' ⚠️期限超過' : ''));
    });
  }

  lines.push('');
  lines.push('シート: ' + ss.getUrl());
  sendSlack_(lines.join('\n'));
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
  lines.push('⏰ *[SaaS管理] ' + thisMonth + ' 分で未確認のチェックが ' + unchecked.length + '件 残っています*');
  lines.push(unchecked.map(function (n) { return '・' + n; }).join('\n'));
  lines.push('');
  lines.push('確認したら「' + SHEET_LOG + '」の結果列をOK/要対応に更新してください。');
  lines.push('シート: ' + ss.getUrl());
  sendSlack_(lines.join('\n'));
}

// ---------------------------------------------------------------------------
// Slack送信
// ---------------------------------------------------------------------------

/** Incoming WebhookでSlackに送信する。URL未設定時はログ出力のみ */
function sendSlack_(text) {
  const url = PropertiesService.getScriptProperties().getProperty(PROP_SLACK_WEBHOOK);
  if (!url) {
    Logger.log('スクリプト プロパティ ' + PROP_SLACK_WEBHOOK + ' が未設定です。通知内容:\n' + text);
    return false;
  }
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: text }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    Logger.log('Slack送信に失敗しました: HTTP ' + res.getResponseCode() + ' / ' + res.getContentText());
    return false;
  }
  return true;
}

/** 疎通確認用: Slackにテストメッセージを送る */
function testSlackNotification() {
  const ok = sendSlack_('✅ [SaaS管理] 通知テストです。このメッセージが見えていれば設定完了です。');
  Logger.log(ok ? 'Slackに送信しました。' : '送信できませんでした。ログを確認してください。');
}
