/**
 * Step 2-3: Claude API(Anthropic)のコスト・APIキー・メンバーの自動取得
 *
 * setup.gs と同じApps Scriptプロジェクトに追加する(このファイル単体で動作する)。
 *
 * 事前準備(READMEにも記載):
 *   1. Claude Consoleで Admin APIキー(sk-ant-admin...)を発行
 *      (組織のadminロールが必要)
 *   2. スクリプト プロパティ ANTHROPIC_ADMIN_KEY にキーを設定
 *
 * 動作(通知は送らず、結果はすべてシートに書き込む):
 *   - 前月のAPIコスト(USD)をワークスペース別に集計し、
 *     月次チェックログの当月「Claude API」行の数量列に合計額を記入
 *     (数値なので前月比のコスト増減が自動計算される)
 *   - アクティブなAPIキー本数・組織メンバー数・在籍者マスタとの不一致をメモ列に記録
 *   - コストが想定内か、失効すべきキーがないかのOK判定は従来どおり人間が行う
 *   - トリガー: 毎月1日 10時台
 */

const PROP_ANTHROPIC_KEY = 'ANTHROPIC_ADMIN_KEY';
const CLAUDE_LOG_TOOL_NAME = 'Claude API';
const ANTHROPIC_API_BASE = 'https://api.anthropic.com';

/** 月次トリガーを登録する(再実行しても重複しない) */
function setupClaudeSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runClaudeSync') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('runClaudeSync').timeBased().onMonthDay(1).atHour(10).create();
  Logger.log('トリガーを登録しました: 毎月1日10時台(Claude APIコスト等の自動取得)');
}

function runClaudeSync() {
  syncClaudeApi();
}

/** Anthropic Admin APIからコスト・キー・メンバーを取得してログに反映する(手動実行可) */
function syncClaudeApi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const users = fetchAnthropicList_('/v1/organizations/users');
  const apiKeys = fetchAnthropicList_('/v1/organizations/api_keys', 'status=active');
  const workspaces = fetchAnthropicList_('/v1/organizations/workspaces');
  const cost = fetchPrevMonthCost_();

  // ワークスペースID → 名前
  const wsNames = {};
  workspaces.forEach(function (w) { wsNames[w.id] = w.name; });

  // コスト内訳の文字列(例: "デフォルト $12.34 / ヤップ $5.67")
  const breakdown = Object.keys(cost.byWorkspace).map(function (wsId) {
    const label = wsId === 'default' ? 'デフォルト' : (wsNames[wsId] || wsId);
    return label + ' $' + cost.byWorkspace[wsId].toFixed(2);
  }).join(' / ');

  // 在籍者マスタとの突合(メンバーのメールが「在籍」にあるか)
  const roster = claudeRosterEmails_(ss);
  const unknown = users.filter(function (u) {
    return !u.email || !roster[String(u.email).toLowerCase()];
  }).map(function (u) { return u.email || u.name; });

  let memo = '前月(' + cost.monthLabel + ')コスト合計 $' + cost.total.toFixed(2) +
    (breakdown ? '(' + breakdown + ')' : '') +
    ' / アクティブAPIキー' + apiKeys.length + '本' +
    ' / メンバー' + users.length + '名';
  if (unknown.length > 0) {
    memo += ' / 要確認(在籍者マスタに不在): ' + unknown.join(', ');
  }

  writeClaudeLogRow_(ss, Math.round(cost.total * 100) / 100, memo);

  Logger.log(
    '取得完了: 前月コスト $' + cost.total.toFixed(2) +
    ' / APIキー' + apiKeys.length + '本 / メンバー' + users.length + '名' +
    (unknown.length > 0 ? ' / マスタ不在: ' + unknown.join(', ') : '')
  );
}

// ---------------------------------------------------------------------------
// Anthropic Admin API
// ---------------------------------------------------------------------------

function anthropicHeaders_() {
  const key = PropertiesService.getScriptProperties().getProperty(PROP_ANTHROPIC_KEY);
  if (!key) {
    throw new Error('スクリプト プロパティ ' + PROP_ANTHROPIC_KEY + ' が未設定です。READMEのStep 2-3の手順でAdmin APIキーを設定してください。');
  }
  return { 'x-api-key': key, 'anthropic-version': '2023-06-01' };
}

function anthropicGet_(url) {
  const res = UrlFetchApp.fetch(url, { headers: anthropicHeaders_(), muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    throw new Error('Anthropic APIエラー: HTTP ' + res.getResponseCode() + ' / ' + res.getContentText() +
      '(Admin APIキー(sk-ant-admin...)か確認してください)');
  }
  return JSON.parse(res.getContentText());
}

/** 一覧系エンドポイントを全ページ取得する(after_idベースのページング) */
function fetchAnthropicList_(path, extraQuery) {
  const items = [];
  let afterId = '';
  do {
    let url = ANTHROPIC_API_BASE + path + '?limit=100' +
      (extraQuery ? '&' + extraQuery : '') +
      (afterId ? '&after_id=' + encodeURIComponent(afterId) : '');
    const json = anthropicGet_(url);
    (json.data || []).forEach(function (item) { items.push(item); });
    afterId = json.has_more ? json.last_id : '';
  } while (afterId);
  return items;
}

/** 前月1ヶ月分のコストをワークスペース別に集計する(単位: USD) */
function fetchPrevMonthCost_() {
  const now = new Date();
  const start = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() - 1, 1));
  const end = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 1));
  const fmt = function (d) { return Utilities.formatDate(d, 'UTC', "yyyy-MM-dd'T'00:00:00'Z'"); };
  const monthLabel = Utilities.formatDate(start, 'UTC', 'yyyy-MM');

  const byWorkspace = {};
  let total = 0;
  let page = '';
  do {
    let url = ANTHROPIC_API_BASE + '/v1/organizations/cost_report' +
      '?starting_at=' + encodeURIComponent(fmt(start)) +
      '&ending_at=' + encodeURIComponent(fmt(end)) +
      '&group_by[]=workspace_id&limit=31' +
      (page ? '&page=' + encodeURIComponent(page) : '');
    const json = anthropicGet_(url);
    (json.data || []).forEach(function (bucket) {
      (bucket.results || []).forEach(function (r) {
        const usd = parseFloat(r.amount || '0') / 100; // amountはセント単位の文字列
        if (isNaN(usd)) return;
        const wsId = r.workspace_id || 'default';
        byWorkspace[wsId] = (byWorkspace[wsId] || 0) + usd;
        total += usd;
      });
    });
    page = json.has_more ? json.next_page : '';
  } while (page);

  return { total: total, byWorkspace: byWorkspace, monthLabel: monthLabel };
}

// ---------------------------------------------------------------------------
// シート書き込み
// ---------------------------------------------------------------------------

/** 在籍者マスタで「在籍」のメールアドレス一覧を返す(小文字キー) */
function claudeRosterEmails_(ss) {
  const data = ss.getSheetByName(SHEET_MEMBERS).getDataRange().getValues();
  const roster = {};
  for (let i = 1; i < data.length; i++) {
    const email = String(data[i][1] || '').toLowerCase().trim();
    if (email && data[i][3] === '在籍') roster[email] = true;
  }
  return roster;
}

/**
 * 月次チェックログの当月「Claude API」行に前月コスト(USD)とメモを書き込む。
 * 行がなければ新規追加する(D・E列の数式は温存)。
 */
function writeClaudeLogRow_(ss, costUsd, extraMemo) {
  const sheet = ss.getSheetByName(SHEET_LOG);
  const tz = ss.getSpreadsheetTimeZone();
  const now = new Date();
  const thisMonth = Utilities.formatDate(now, tz, 'yyyy-MM');
  const memo = '自動取得 ' + Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm') + ' / ' + extraMemo;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === thisMonth && data[i][1] === CLAUDE_LOG_TOOL_NAME) {
      const row = i + 1;
      sheet.getRange(row, 3).setValue(costUsd);       // 数量 = 前月コスト(USD)
      sheet.getRange(row, 7).setValue(memo);          // メモ
      sheet.getRange(row, 8).setValue('自動(GAS)');   // 確認者
      return;
    }
  }

  // 当月行がない場合は追加(A〜CとF〜Iを分けて書き、D・E列の数式を温存)
  let lastRow = 1;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] !== '') { lastRow = i + 1; break; }
  }
  sheet.getRange(lastRow + 1, 1, 1, 3).setValues([[thisMonth, CLAUDE_LOG_TOOL_NAME, costUsd]]);
  sheet.getRange(lastRow + 1, 6, 1, 4).setValues([['未確認', memo, '自動(GAS)', '']]);
}
