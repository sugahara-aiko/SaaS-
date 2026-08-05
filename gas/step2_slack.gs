/**
 * Step 2-2: Slack ユーザー数の自動取得
 *
 * setup.gs と同じApps Scriptプロジェクトに追加する(このファイル単体で動作する)。
 *
 * 事前準備(READMEにも記載):
 *   1. https://api.slack.com/apps でSlackアプリを作成し、Botトークンを発行
 *      (必要スコープ: users:read, users:read.email)
 *   2. スクリプト プロパティ SLACK_BOT_TOKEN にトークン(xoxb-...)を設定
 *
 * 動作(通知は送らず、結果はすべてシートに書き込む):
 *   - Slackの全ユーザーを取得し、以下の4区分に自動分類して
 *     月次チェックログの当月の各行に人数を記入する:
 *       Slack(管理者) / Slack(メンバー) / Slack(マルチchゲスト) / Slack(シングルchゲスト)
 *   - 管理者・メンバーのうち在籍者マスタで「在籍」になっていない人を
 *     メモ列に記録(退職者残りの検知)。ゲスト行のメモには名前一覧を記録
 *     (不要なゲストがいないかの目視チェック用)
 *   - 結果のOK/要対応の判定は従来どおり人間が行う
 *   - トリガー: 毎月1日 10時台(Step 1の行生成の後に実行される)
 */

const PROP_SLACK_TOKEN = 'SLACK_BOT_TOKEN';

const SLACK_LOG_NAMES = {
  admin: 'Slack(管理者)',
  member: 'Slack(メンバー)',
  mcg: 'Slack(マルチchゲスト)',
  scg: 'Slack(シングルchゲスト)'
};

/** 月次トリガーを登録する(再実行しても重複しない) */
function setupSlackSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runSlackSync') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('runSlackSync').timeBased().onMonthDay(1).atHour(10).create();
  Logger.log('トリガーを登録しました: 毎月1日10時台(Slackユーザー数の自動取得)');
}

function runSlackSync() {
  syncSlackUsers();
}

/** Slackユーザーを取得して月次チェックログに反映する(手動実行可) */
function syncSlackUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const users = fetchSlackMembers_();
  const roster = getActiveRosterEmails_(ss);

  const byCat = { admin: [], member: [], mcg: [], scg: [] };
  users.forEach(function (u) { byCat[u.category].push(u); });

  // 管理者・メンバー: 在籍者マスタで「在籍」になっていない人を要確認としてメモ
  ['admin', 'member'].forEach(function (cat) {
    const unknown = byCat[cat].filter(function (u) {
      return !u.email || !roster[u.email.toLowerCase()];
    });
    const memo = unknown.length > 0
      ? '要確認(在籍者マスタに不在): ' + unknown.map(function (u) { return u.email || u.name; }).join(', ')
      : '';
    writeSlackLogRow_(ss, SLACK_LOG_NAMES[cat], byCat[cat].length, memo);
  });

  // ゲスト: 目視チェック用に名前一覧をメモ
  ['mcg', 'scg'].forEach(function (cat) {
    const memo = byCat[cat].length > 0
      ? '内訳: ' + byCat[cat].map(function (u) { return u.name; }).join(', ')
      : '';
    writeSlackLogRow_(ss, SLACK_LOG_NAMES[cat], byCat[cat].length, memo);
  });

  Logger.log(
    '取得完了: 管理者' + byCat.admin.length + ' / メンバー' + byCat.member.length +
    ' / マルチchゲスト' + byCat.mcg.length + ' / シングルchゲスト' + byCat.scg.length
  );
}

/** Slack users.list APIで全ユーザーを取得し4区分に分類する(Bot・解除済みは除外) */
function fetchSlackMembers_() {
  const token = PropertiesService.getScriptProperties().getProperty(PROP_SLACK_TOKEN);
  if (!token) {
    throw new Error('スクリプト プロパティ ' + PROP_SLACK_TOKEN + ' が未設定です。READMEのStep 2-2の手順でBotトークンを設定してください。');
  }

  const users = [];
  let cursor = '';
  do {
    const url = 'https://slack.com/api/users.list?limit=200' +
      (cursor ? '&cursor=' + encodeURIComponent(cursor) : '');
    const res = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const json = JSON.parse(res.getContentText());
    if (!json.ok) {
      throw new Error('Slack APIエラー: ' + json.error + '(トークンとスコープ users:read, users:read.email を確認してください)');
    }
    (json.members || []).forEach(function (m) {
      if (m.deleted || m.is_bot || m.id === 'USLACKBOT') return;
      let category = 'member';
      if (m.is_ultra_restricted) category = 'scg';
      else if (m.is_restricted) category = 'mcg';
      else if (m.is_admin || m.is_owner || m.is_primary_owner) category = 'admin';
      users.push({
        name: m.real_name || m.name,
        email: (m.profile && m.profile.email) || '',
        category: category
      });
    });
    cursor = (json.response_metadata && json.response_metadata.next_cursor) || '';
  } while (cursor);
  return users;
}

/** 在籍者マスタで「在籍」のメールアドレス一覧を返す(小文字キー) */
function getActiveRosterEmails_(ss) {
  const data = ss.getSheetByName(SHEET_MEMBERS).getDataRange().getValues();
  const roster = {};
  for (let i = 1; i < data.length; i++) {
    const email = String(data[i][1] || '').toLowerCase().trim();
    if (email && data[i][3] === '在籍') roster[email] = true;
  }
  return roster;
}

/**
 * 月次チェックログの当月・指定ツール名の行に人数とメモを書き込む。
 * 行がなければ新規追加する(D・E列の数式は温存)。
 */
function writeSlackLogRow_(ss, toolName, count, extraMemo) {
  const sheet = ss.getSheetByName(SHEET_LOG);
  const tz = ss.getSpreadsheetTimeZone();
  const now = new Date();
  const thisMonth = Utilities.formatDate(now, tz, 'yyyy-MM');
  const memo = '自動取得 ' + Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm') +
    (extraMemo ? ' / ' + extraMemo : '');

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === thisMonth && data[i][1] === toolName) {
      const row = i + 1;
      sheet.getRange(row, 3).setValue(count);        // ユーザー数
      sheet.getRange(row, 7).setValue(memo);         // メモ
      sheet.getRange(row, 8).setValue('自動(GAS)');  // 確認者
      return;
    }
  }

  // 当月行がない場合は追加(A〜CとF〜Iを分けて書き、D・E列の数式を温存)
  let lastRow = 1;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] !== '') { lastRow = i + 1; break; }
  }
  sheet.getRange(lastRow + 1, 1, 1, 3).setValues([[thisMonth, toolName, count]]);
  sheet.getRange(lastRow + 1, 6, 1, 4).setValues([['未確認', memo, '自動(GAS)', '']]);
}
