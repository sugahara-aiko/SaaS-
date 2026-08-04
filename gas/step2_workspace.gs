/**
 * Step 2-1: Google Workspace ユーザー数の自動取得
 *
 * setup.gs / step1_notifications.gs と同じApps Scriptプロジェクトに追加する。
 *
 * 事前準備(READMEにも記載):
 *   1. エディタ左の「サービス +」から「Admin SDK Directory API」を追加
 *      (識別子はデフォルトの AdminDirectory のまま)
 *   2. Google Workspaceの管理者権限があるアカウントで実行すること
 *
 * 動作:
 *   - Workspaceの全ユーザーを取得し、
 *     a. 月次チェックログの当月「GoogleWorkspace」行にアクティブ数を自動記入
 *        (結果のOK/要対応の判定は従来どおり人間が行う)
 *     b. 在籍者マスタを更新(新規ユーザーの追加、停止中→退職への反映)
 *     c. 結果をSlack通知(在籍者マスタで「在籍」なのにWorkspaceにいない人も報告)
 *   - トリガー: 毎月1日 10時台(Step 1の行生成の後に実行される)
 */

const WS_LOG_TOOL_NAME = 'GoogleWorkspace';

/** 月次トリガーを登録する(再実行しても重複しない) */
function setupStep2Triggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runWorkspaceSync') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('runWorkspaceSync').timeBased().onMonthDay(1).atHour(10).create();
  Logger.log('トリガーを登録しました: 毎月1日10時台(Workspaceユーザー数の自動取得)');
}

function runWorkspaceSync() {
  syncWorkspaceUsers();
}

/** Workspaceユーザーを取得してログ・在籍者マスタ・Slackに反映する(手動実行可) */
function syncWorkspaceUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const users = fetchWorkspaceUsers_();
  const active = users.filter(function (u) { return !u.suspended; });
  const suspended = users.filter(function (u) { return u.suspended; });

  const memberResult = updateMembersSheet_(ss, users);
  const sheetOnly = findSheetOnlyMembers_(ss, users);
  const logCell = writeLogCount_(ss, active.length, suspended.length);

  const tz = ss.getSpreadsheetTimeZone();
  const lines = [];
  lines.push('🤖 *[SaaS管理] GoogleWorkspace ユーザー数を自動取得しました*');
  lines.push('アクティブ: ' + active.length + '名 / 停止中: ' + suspended.length + '名');
  lines.push(logCell
    ? '月次チェックログの ' + logCell + ' に反映済み。内容を確認して結果をOKにしてください。'
    : '⚠️ 当月のログ行が見つからなかったため、行を新規追加しました。');
  if (memberResult.added.length > 0) {
    lines.push('');
    lines.push('👤 在籍者マスタに自動追加: ' + memberResult.added.join(', '));
  }
  if (suspended.length > 0) {
    lines.push('');
    lines.push('⏸️ 停止中アカウント(退職処理済みか確認):');
    lines.push(suspended.map(function (u) { return '・' + u.email; }).join('\n'));
  }
  if (sheetOnly.length > 0) {
    lines.push('');
    lines.push('⚠️ 在籍者マスタで「在籍」だがWorkspaceに存在しないメール(業務委託などWorkspace未付与なら問題なし):');
    lines.push(sheetOnly.map(function (e) { return '・' + e; }).join('\n'));
  }
  lines.push('');
  lines.push('取得日時: ' + Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm'));
  sendSlack_(lines.join('\n'));
}

/** Admin SDKで全ユーザーを取得する */
function fetchWorkspaceUsers_() {
  const users = [];
  let pageToken;
  do {
    const res = AdminDirectory.Users.list({
      customer: 'my_customer',
      maxResults: 500,
      orderBy: 'email',
      pageToken: pageToken
    });
    (res.users || []).forEach(function (u) {
      users.push({
        email: u.primaryEmail,
        name: (u.name && u.name.fullName) || '',
        suspended: !!u.suspended || !!u.archived
      });
    });
    pageToken = res.nextPageToken;
  } while (pageToken);
  return users;
}

/**
 * 在籍者マスタを更新する。
 * - 新規メールは行を追加(区分・入社日などは手動入力のため空欄のまま)
 * - 既存行は在籍状況のみ更新(氏名などの手動編集は上書きしない)
 * - Workspaceから消えた人は自動では退職にしない(業務委託がWorkspace未付与の
 *   場合があるため)→ findSheetOnlyMembers_ でSlack報告に回す
 */
function updateMembersSheet_(ss, users) {
  const sheet = ss.getSheetByName(SHEET_MEMBERS);
  const data = sheet.getDataRange().getValues();
  const emailToRow = {}; // メール(小文字) → シート行番号
  for (let i = 1; i < data.length; i++) {
    const email = String(data[i][1] || '').toLowerCase().trim();
    if (email) emailToRow[email] = i + 1;
  }

  const added = [];
  let nextRow = data.length + 1;
  users.forEach(function (u) {
    const key = String(u.email).toLowerCase();
    const status = u.suspended ? '退職' : '在籍';
    if (emailToRow[key]) {
      const row = emailToRow[key];
      if (data[row - 1][3] !== status) {
        sheet.getRange(row, 4).setValue(status);
        if (u.suspended) {
          const note = String(data[row - 1][6] || '');
          if (note.indexOf('停止中') === -1) {
            sheet.getRange(row, 7).setValue((note ? note + ' / ' : '') + 'Workspaceアカウント停止中');
          }
        }
      }
    } else {
      sheet.getRange(nextRow, 1, 1, 7).setValues([[
        u.name, u.email, '', status, '', '',
        '自動追加(Workspace)' + (u.suspended ? ' / アカウント停止中' : '')
      ]]);
      added.push(u.email);
      nextRow++;
    }
  });
  return { added: added };
}

/** 在籍者マスタで「在籍」なのにWorkspaceに存在しないメールを返す */
function findSheetOnlyMembers_(ss, users) {
  const sheet = ss.getSheetByName(SHEET_MEMBERS);
  const data = sheet.getDataRange().getValues();
  const wsEmails = {};
  users.forEach(function (u) { wsEmails[String(u.email).toLowerCase()] = true; });

  const result = [];
  for (let i = 1; i < data.length; i++) {
    const email = String(data[i][1] || '').toLowerCase().trim();
    if (email && data[i][3] === '在籍' && !wsEmails[email]) result.push(data[i][1]);
  }
  return result;
}

/**
 * 月次チェックログの当月「GoogleWorkspace」行にアクティブ数を書き込む。
 * 行がなければ新規追加する(D・E列の数式は温存)。
 *
 * @return {string|null} 書き込んだ位置の説明。新規追加時は null
 */
function writeLogCount_(ss, activeCount, suspendedCount) {
  const sheet = ss.getSheetByName(SHEET_LOG);
  const tz = ss.getSpreadsheetTimeZone();
  const now = new Date();
  const thisMonth = Utilities.formatDate(now, tz, 'yyyy-MM');
  const memo = '自動取得 ' + Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm') +
    '(停止中' + suspendedCount + '名は含まず)';

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === thisMonth && data[i][1] === WS_LOG_TOOL_NAME) {
      const row = i + 1;
      sheet.getRange(row, 3).setValue(activeCount);   // ユーザー数
      sheet.getRange(row, 7).setValue(memo);          // メモ
      sheet.getRange(row, 8).setValue('自動(GAS)');   // 確認者
      return thisMonth + ' の ' + WS_LOG_TOOL_NAME + ' 行';
    }
  }

  // 当月行がない場合は追加(A〜CとF〜Iを分けて書き、D・E列の数式を温存)
  let lastRow = 1;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] !== '') { lastRow = i + 1; break; }
  }
  sheet.getRange(lastRow + 1, 1, 1, 3).setValues([[thisMonth, WS_LOG_TOOL_NAME, activeCount]]);
  sheet.getRange(lastRow + 1, 6, 1, 4).setValues([['未確認', memo, '自動(GAS)', '']]);
  return null;
}
