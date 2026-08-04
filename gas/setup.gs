/**
 * SaaS管理シート セットアップスクリプト
 *
 * 既存のスプレッドシートに以下の5シートを新規作成し、現行シートのデータを
 * 構造化した形で移行する(既存シートには一切変更を加えない)。
 *
 *   1. ダッシュボード   … 月額費用・更新期限・未チェックの自動集計(数式のみ)
 *   2. ツールマスタ     … 契約情報のマスタ(1プラン1行、単価は月額換算の数値)
 *   3. 月次チェックログ … 毎月のチェック結果を上書きせず行追加で記録
 *   4. 在籍者マスタ     … 退職者突合用の在籍者一覧(Step2で自動取得予定)
 *   5. 設定             … 為替レートなど
 *
 * 使い方: この関数を選択して実行 → 初回のみ権限を承認
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const names = [SHEET_DASHBOARD, SHEET_MASTER, SHEET_LOG, SHEET_MEMBERS, SHEET_CONFIG];
  const existing = names.filter(function (n) { return ss.getSheetByName(n); });
  if (existing.length > 0) {
    throw new Error(
      '同名のシートが既に存在するため中断しました: ' + existing.join(', ') +
      '\n(再実行する場合は該当シートを削除するか名前を変更してください)'
    );
  }

  createConfigSheet_(ss);
  createMasterSheet_(ss);
  createLogSheet_(ss);
  createMembersSheet_(ss);
  createDashboardSheet_(ss);

  // ダッシュボードを先頭へ移動
  ss.setActiveSheet(ss.getSheetByName(SHEET_DASHBOARD));
  ss.moveActiveSheet(1);

  ss.toast('5シートの作成とデータ移行が完了しました。黄色のセル(要確認)を確認してください。', 'セットアップ完了', 10);
  Logger.log('セットアップ完了: ' + names.join(', '));
}

// ---------------------------------------------------------------------------
// シート名・共通定義
// ---------------------------------------------------------------------------

const SHEET_DASHBOARD = 'ダッシュボード';
const SHEET_MASTER = 'ツールマスタ';
const SHEET_LOG = '月次チェックログ';
const SHEET_MEMBERS = '在籍者マスタ';
const SHEET_CONFIG = '設定';

const COLOR_HEADER_BG = '#37474f';
const COLOR_HEADER_TEXT = '#ffffff';
const COLOR_FLAG_BG = '#fff2cc'; // 要確認セル(黄色)
const COLOR_SECTION_BG = '#eceff1';

const LIST_STATUS = ['契約中', '解約予定', '解約済'];
const LIST_BILLING_UNIT = ['人数課金', '固定', '従量', 'プリペイド', '不明'];
const LIST_CURRENCY = ['JPY', 'USD', 'EUR'];
const LIST_CONTRACT_TYPE = ['01_ユーザー課金', '02_開発基盤', '03_事務・固定費'];
const LIST_CHECK_RESULT = ['OK', '要対応', '未確認'];

/** 日付ヘルパー(月は1始まり) */
function d_(y, m, day) {
  return new Date(y, m - 1, day);
}

/** ヘッダー行の書式設定(1行目固定・背景色・太字) */
function styleHeader_(sheet, numCols) {
  sheet.getRange(1, 1, 1, numCols)
    .setBackground(COLOR_HEADER_BG)
    .setFontColor(COLOR_HEADER_TEXT)
    .setFontWeight('bold')
    .setWrap(true);
  sheet.setFrozenRows(1);
}

/** ドロップダウン(リスト選択)のデータ検証を設定 */
function setListValidation_(sheet, a1Range, list) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(a1Range).setDataValidation(rule);
}

// ---------------------------------------------------------------------------
// 1. 設定シート
// ---------------------------------------------------------------------------

function createConfigSheet_(ss) {
  const sheet = ss.insertSheet(SHEET_CONFIG);

  sheet.getRange('A1:D1').setValues([['為替レート', '自動取得', '手動上書き', '採用レート']]);
  sheet.getRange('A2:A3').setValues([['USD/JPY'], ['EUR/JPY']]);
  sheet.getRange('B2').setFormula('=GOOGLEFINANCE("CURRENCY:USDJPY")');
  sheet.getRange('B3').setFormula('=GOOGLEFINANCE("CURRENCY:EURJPY")');
  sheet.getRange('D2').setFormula('=IF($C2<>"",$C2,$B2)');
  sheet.getRange('D3').setFormula('=IF($C3<>"",$C3,$B3)');
  sheet.getRange('C2').setNote('レートを固定したい場合はここに数値を入力(空欄なら自動取得レートを採用)');
  sheet.getRange('B2:D3').setNumberFormat('#,##0.00');

  sheet.getRange('A5').setValue('※ ツールマスタの「月額換算(円)」はこの採用レートで計算されます。');

  styleHeader_(sheet, 4);
  sheet.setColumnWidths(1, 4, 110);
}

// ---------------------------------------------------------------------------
// 2. ツールマスタ
// ---------------------------------------------------------------------------

const MASTER_HEADERS = [
  'ツール名', 'プラン名', '契約区分', 'カテゴリ', 'ステータス',
  '課金単位', '単価(月額換算)', '通貨', '数量(人数)', '月額換算(円)',
  '支払サイクル', '次回更新日', '月次チェック対象', 'チェック項目',
  '担当・メモ', 'ログインURL', '自動化ステータス'
];

// 列: [ツール名, プラン名, 契約区分, カテゴリ, ステータス, 課金単位, 単価, 通貨, 数量,
//      (月額換算=数式のため空), 支払サイクル, 次回更新日, 月次チェック対象, チェック項目, 担当・メモ, URL, 自動化]
// 複数プランのツール(Figma/Claude/ChatGPT)は1プラン1行に分解。
// 単価は「月額換算の単価」(年額契約は12で割った値を入力)。
const MASTER_DATA = [
  ['Slack', 'プロ', '01_ユーザー課金', 'コミュニケーション', '契約中', '人数課金', 1050, 'JPY', '', '',
    '要確認', '', true,
    '1. 組織外の人が混じっていないか\n2. 退職者が残っていないか(非アクティブ)\n3. 不要なシングルチャンネルゲストの削除',
    '', 'https://japanbitcoinindustry.slack.com/account/workspace-settings', '自動化可(users.list API)'],

  ['Notion', 'プラス', '01_ユーザー課金', '情報管理', '契約中', '人数課金', 2000, 'JPY', '', '',
    '年次', d_(2026, 10, 4), true,
    '1. ゲスト枠に不要な人がいないか、いたら削除\n2. 社員・委託以外が「メンバー」になっていないか',
    '', 'https://www.notion.so/login', '一部自動化可(users API/ゲスト取得に制限あり)'],

  ['Figma', 'プロフェッショナル(フル)', '01_ユーザー課金', 'デザイン', '契約中', '人数課金', 3000, 'JPY', 10, '',
    '月次', '毎月末', true,
    '1. 予期せぬ有料ユーザー追加がないか\n2. 追加時はSlack履歴で背景確認(不明時はヒアリング)',
    '閲覧のみ10名(無料)', 'https://www.figma.com/login', '自動化不可(Professionalプランは管理API非対応)'],

  ['Figma', 'プロフェッショナル(Dev)', '01_ユーザー課金', 'デザイン', '契約中', '人数課金', 2250, 'JPY', 1, '',
    '月次', '毎月末', true,
    '(フル席の行と同じチェックをまとめて実施)',
    '', 'https://www.figma.com/login', '自動化不可(Professionalプランは管理API非対応)'],

  ['IVRY', 'スタンダード', '03_事務・固定費', '電話代行', '契約中', '固定', 6980, 'JPY', '', '',
    '月次', '', true,
    '1. 毎月の請求メールを確認\n2. 従量課金が異常値になっていないか',
    '基本料のみ計上、従量課金は別途', 'https://ivry.jp/home/', '一部自動化可(請求メールのGmail解析)'],

  ['Claude API', 'プリペイド', '02_開発基盤', 'AIツール', '契約中', 'プリペイド', '', 'USD', '', '',
    '月次', '', true,
    '1. 月次のトークン消費を確認し、コストが想定範囲内か確認\n2. 利用プロジェクトが現在も有効か確認(不要な連携は停止)\n3. APIキーの発行数を確認し失効すべきキーがないか確認',
    '組織API 上限$100 / ヤップAPI 上限$45',
    'https://platform.claude.com/settings/workspaces/wrkspc_01G44RCmCC1AyEE88VpqRmAW/limits',
    '自動化可(Anthropic Admin APIでコスト・メンバー取得)'],

  ['Claude', 'Pro', '01_ユーザー課金', 'AIツール', '契約中', '人数課金', 20, 'USD', 6, '',
    '月次', '', true,
    '各自のアカウントで管理(解約時には解約画面を受領)\n∟管理台帳と照合',
    '個別付与(詳細は別シート)', '', '自動化不可(個人課金)'],

  ['Claude', 'Max 5x', '01_ユーザー課金', 'AIツール', '契約中', '人数課金', 100, 'USD', 4, '',
    '月次', '', true,
    '各自のアカウントで管理(解約時には解約画面を受領)\n∟管理台帳と照合',
    '個別付与(詳細は別シート)', '', '自動化不可(個人課金)'],

  ['Claude', 'Max 20x', '01_ユーザー課金', 'AIツール', '契約中', '人数課金', 200, 'USD', 8, '',
    '月次', '', true,
    '各自のアカウントで管理(解約時には解約画面を受領)\n∟管理台帳と照合',
    '個別付与(詳細は別シート)', '', '自動化不可(個人課金)'],

  ['ChatGPT', 'Business', '01_ユーザー課金', 'AIツール', '契約中', '人数課金', 30, 'USD', '', '',
    '月次', '', true,
    '各自のアカウントで管理(解約時には解約画面を受領)\n∟管理台帳と照合',
    '個別付与(詳細は別シート)', '', '自動化不可(個人課金)'],

  ['ChatGPT', 'Pro', '01_ユーザー課金', 'AIツール', '契約中', '人数課金', 200, 'USD', 4, '',
    '月次', '', true,
    '各自のアカウントで管理(解約時には解約画面を受領)\n∟管理台帳と照合',
    '個別付与(詳細は別シート)', '', '自動化不可(個人課金)'],

  ['Bitwarden', 'チーム', '03_事務・固定費', 'パスワード管理', '契約中', '人数課金', 4, 'USD', 24, '',
    '年次', d_(2027, 2, 4), true,
    '1. 組織外の人が混じっていないか\n2. 退職者が残っていないか(非アクティブ)\n3. 組織図と権限状態が一致しているか',
    '', 'https://vault.bitwarden.com/#/organizations/fe8ea7f2-6752-4763-9776-b27400d56d9e/members',
    '要検証(TeamsプランのPublic APIでメンバー取得できる可能性)'],

  ['GoogleWorkspace', 'Standard', '03_事務・固定費', '労務管理', '契約中', '人数課金', 1280, 'JPY', 30, '',
    '月次', '毎月末', true,
    '1. 組織外の人が混じっていないか\n2. 退職者が残っていないか(非アクティブ)',
    '', 'https://workspace.google.com/intl/ja/products/admin/', '自動化可(Admin SDK・最優先候補)'],

  ['GoogleWorkspace(カシェイ)', 'Standard', '03_事務・固定費', '労務管理', '契約中', '人数課金', 1280, 'JPY', 4, '',
    '月次', '', false, '',
    '年間契約(月割請求)、岡田さん管理', 'https://workspace.google.com/intl/ja/products/admin/', ''],

  ['Vercel', 'Pro', '02_開発基盤', 'サーバー維持', '契約中', '人数課金', 20, 'USD', '', '',
    '-', '', false, '', '', 'https://vercel.com/login', ''],

  ['Supabase(YAP APP)', 'Team', '02_開発基盤', 'ヤップDB維持', '契約中', '固定', 599, 'USD', '', '',
    '月次', '毎月12日', false, '', 'エンジニアにてウォッチ', 'https://supabase.com/dashboard', ''],

  ['Supabase(YAP AI)', 'Pro', '02_開発基盤', 'ヤップ分析AI接続専用', '契約中', '固定', 25, 'USD', '', '',
    '月次', '毎月7日', false, '', 'ヤップBizチームにてウォッチ', 'https://supabase.com/dashboard', ''],

  ['GitHub', 'GitHub Actions(上限$100)', '02_開発基盤', '開発管理', '契約中', '従量', '', 'USD', '', '',
    '-', '', false, '', '規新さん対応', 'https://github.com/login', '自動化可(Orgメンバー API)'],

  ['codemagic', '年額$3,990', '02_開発基盤', '開発ツール', '契約中', '固定', 332.5, 'USD', '', '',
    '年次', '', false, '', '年額$3,990の月額換算', 'https://codemagic.io/start/', ''],

  ['planetscale', '', '02_開発基盤', 'DB維持', '契約中', '固定', 5, 'USD', '', '',
    '月次', '', false, '', '', 'https://planetscale.com/', ''],

  ['Axiom', '', '02_開発基盤', '開発ツール', '契約中', '固定', 25, 'USD', '', '',
    '月次', '', false, '', '阿部さん、admin@のみ / ユーザー数2', 'https://app.axiom.co/yap-app-gtrm/settings/users', ''],

  ['Cachix', '', '02_開発基盤', '開発ツール', '契約中', '固定', 55, 'EUR', '', '',
    '月次', '', false, '', 'アンドレイさん、seikyu@のみ', 'https://app.cachix.org/organization/jbi/billing', ''],

  ['DesitalPress', '', '02_開発基盤', 'ビッ研ブログhosting', '契約中', '固定', 12.9, 'EUR', '', '',
    '月次', '', false, '', '', '', ''],

  ['AWS', '', '02_開発基盤', 'サーバー', '契約中', '従量', '', 'USD', '', '',
    '-', '', false, '', '規新さん主管理', '', ''],

  ['Twilio', '', '02_開発基盤', 'ヤップSMS認証', '契約中', '従量', '', 'USD', '', '',
    '-', '', false, '', '', '', ''],

  ['Deno', '', '02_開発基盤', 'ヤッププッシュ通知', '契約中', '固定', 20, 'USD', '', '',
    '月次', '毎月1日', false, '', '', '', ''],

  ['nosh', '', '03_事務・固定費', '福利厚生', '契約中', '固定', 29940, 'JPY', '', '',
    '月次', '', false, '',
    '25食/回、残数10個未満の場合発注(※冷凍庫に入れること) / 送料別 / 契約書未格納かも', '', ''],

  ['SmartHR', '', '03_事務・固定費', '労務管理', '契約中', '固定', 0, 'JPY', '', '',
    '-', '', false, '', 'ユーザー数17', 'https://jbi.smarthr.jp/', 'API有(在籍者マスタの供給元候補)'],

  ['バクラク勤怠', 'スタンダード', '03_事務・固定費', '労務管理', '契約中', '人数課金', 1500, 'JPY', 16, '',
    '年次', '', false, '', '', 'https://attendance.layerx.jp/', ''],

  ['マネフォ給与', '基本料金+300円/人(6名〜)', '03_事務・固定費', '給与計算', '契約中', '人数課金', 300, 'JPY', 16, '',
    '-', '', false, '', '基本料金は別途(要確認)', 'https://payroll.moneyforward.com/', ''],

  ['マネフォ経費', '', '03_事務・固定費', '経費精算', '契約中', '不明', '', 'JPY', 19, '',
    '-', '', false, '', 'ユーザー数19(役員含む)',
    'https://accounting.moneyforward.com/home?cti=N-6bX660nXs6IU5SdtUviw', ''],

  ['マネフォPay', '', '03_事務・固定費', 'バーチャルカード', '契約中', '不明', '', 'JPY', '', '',
    '-', '', false, '', '', 'https://biz-pay.moneyforward.com/home', ''],

  ['マネフォクラウド', '', '03_事務・固定費', '財務管理', '契約中', '不明', '', 'JPY', '', '',
    '年次', d_(2026, 12, 25), false, '', '',
    'https://accounting.moneyforward.com/home?cti=N-6bX660nXs6IU5SdtUviw', ''],

  ['クラウドサイン', 'Light', '03_事務・固定費', '契約管理', '契約中', '固定', 10000, 'JPY', '', '',
    '-', '', false, '', '基本料のみ計上、従量¥200/件は別途 / 橋本、菅原のみ', 'https://www.cloudsign.jp/dashboard', ''],

  ['Spotify', 'Premium Standard', '03_事務・固定費', '音楽再生', '契約中', '固定', 1080, 'JPY', '', '',
    '月次', '毎月25日', false, '', '', 'https://www.spotify.com/jp/account/overview/', ''],

  ['Weglot', 'Starter', '03_事務・固定費', 'HP多言語化', '解約済', '固定', 15, 'EUR', '', '',
    '月次', '', false, '', 'コーポレートサイト多言語対応用途、実装できず即日解約済', 'https://www.weglot.com/', ''],

  ['Windsurf', '', '02_開発基盤', '開発ツール', '解約済', '不明', '', 'USD', '', '',
    '-', '', false, '', 'こうへいさん用、橋本アカウントで管理 / 26/2/9解約', 'https://windsurf.com/', ''],

  ['tl;dv', 'Pro', '01_ユーザー課金', '録画ツール', '解約予定', '人数課金', 4980, 'JPY', 5, '',
    '月次', '', true, '不要なメンバーが混じっていないか', 'Pro5 → 解約予定', 'https://tldv.io/', '']
];

// 要確認セル: [ツール名, プラン名, 列番号(1始まり), メモ]
const MASTER_FLAGS = [
  ['Slack', 'プロ', 9, '有料ユーザー数が旧シートに記録なし。要入力'],
  ['Slack', 'プロ', 11, '旧シートで支払サイクルが「?」となっていたため要確認'],
  ['Notion', 'プラス', 9, '有料メンバー数が旧シートに記録なし。要入力'],
  ['ChatGPT', 'Business', 9, '人数不明。旧シートのメモ(Plusx20 1 / Plusx5 3 / Pro 4)がプラン名(Business/Pro)と不整合のため要確認'],
  ['ChatGPT', 'Pro', 9, '旧シートメモの「Pro 4」から推定。要確認'],
  ['Vercel', 'Pro', 9, '有料シート数が旧シートに記録なし。要入力'],
  ['バクラク勤怠', 'スタンダード', 7, '旧シートの「1500円/人」が月額単価か年額単価か要確認(この列は月額換算単価)'],
  ['マネフォ給与', '基本料金+300円/人(6名〜)', 7, '基本料金が不明のため従量単価のみ計上。要確認'],
  ['マネフォ経費', '', 7, '料金不明。要入力'],
  ['マネフォPay', '', 7, '料金不明。要入力'],
  ['マネフォクラウド', '', 7, '料金不明。要入力']
];

function createMasterSheet_(ss) {
  const sheet = ss.insertSheet(SHEET_MASTER);
  const numCols = MASTER_HEADERS.length;
  const numRows = MASTER_DATA.length;

  sheet.getRange(1, 1, 1, numCols).setValues([MASTER_HEADERS]);
  sheet.getRange(2, 1, numRows, numCols).setValues(MASTER_DATA);

  // J列: 月額換算(円) = 単価 × 数量 × 為替レート(解約済・従量・プリペイド・単価空欄は除外)
  const formulas = [];
  for (let r = 2; r <= numRows + 1; r++) {
    formulas.push([
      '=IF(OR($E' + r + '="解約済",$G' + r + '="",$F' + r + '="従量",$F' + r + '="プリペイド"),"",' +
      '$G' + r + '*IF($I' + r + '="",1,$I' + r + ')*' +
      'IF($H' + r + '="JPY",1,IF($H' + r + '="USD",\'' + SHEET_CONFIG + '\'!$D$2,' +
      'IF($H' + r + '="EUR",\'' + SHEET_CONFIG + '\'!$D$3,1))))'
    ]);
  }
  sheet.getRange(2, 10, numRows, 1).setFormulas(formulas);

  // データ検証(ドロップダウン・チェックボックス)
  const lastValidationRow = 200;
  setListValidation_(sheet, 'C2:C' + lastValidationRow, LIST_CONTRACT_TYPE);
  setListValidation_(sheet, 'E2:E' + lastValidationRow, LIST_STATUS);
  setListValidation_(sheet, 'F2:F' + lastValidationRow, LIST_BILLING_UNIT);
  setListValidation_(sheet, 'H2:H' + lastValidationRow, LIST_CURRENCY);
  sheet.getRange(2, 13, numRows, 1).insertCheckboxes();

  // 書式
  sheet.getRange(2, 7, numRows, 1).setNumberFormat('#,##0.0');   // 単価
  sheet.getRange(2, 10, numRows, 1).setNumberFormat('¥#,##0');   // 月額換算(円)
  sheet.getRange(2, 12, numRows, 1).setNumberFormat('yyyy/mm/dd'); // 次回更新日(日付のもののみ)
  sheet.getRange(2, 14, numRows, 2).setWrap(true);

  // 要確認セルのハイライトとメモ
  MASTER_FLAGS.forEach(function (flag) {
    const rowIdx = MASTER_DATA.findIndex(function (row) {
      return row[0] === flag[0] && row[1] === flag[1];
    });
    if (rowIdx === -1) return;
    const cell = sheet.getRange(rowIdx + 2, flag[2]);
    cell.setBackground(COLOR_FLAG_BG).setNote('【要確認】' + flag[3]);
  });

  // 解約済の行はグレーアウト
  MASTER_DATA.forEach(function (row, i) {
    if (row[4] === '解約済') {
      sheet.getRange(i + 2, 1, 1, numCols).setFontColor('#9e9e9e');
    }
  });

  styleHeader_(sheet, numCols);
  sheet.getRange(1, 1, numRows + 1, numCols).createFilter();

  // 列幅
  sheet.setColumnWidth(1, 160);  // ツール名
  sheet.setColumnWidth(2, 170);  // プラン名
  sheet.setColumnWidth(14, 320); // チェック項目
  sheet.setColumnWidth(15, 240); // 担当・メモ
  sheet.setColumnWidth(16, 200); // URL
  sheet.setColumnWidth(17, 220); // 自動化ステータス
}

// ---------------------------------------------------------------------------
// 3. 月次チェックログ
// ---------------------------------------------------------------------------

const LOG_HEADERS = [
  '年月', 'ツール名', 'ユーザー数・数量', '前月値', '増減',
  '結果', 'メモ', '確認者', '確認日'
];

// 現行シートの「最終確認日・メモ」を2026-07分のログとしてシード投入(履歴の起点)
// 列: [年月, ツール名, ユーザー数・数量, (前月値=数式), (増減=数式), 結果, メモ, 確認者, 確認日]
const LOG_SEED = [
  ['2026-07', 'Slack', '', '', '', '未確認', '', '', ''],
  ['2026-07', 'Notion', '', '', '', '未確認', '', '', ''],
  ['2026-07', 'Figma', 'フル10 / Dev1 / 閲覧10', '', '', 'OK', '', '', d_(2026, 7, 11)],
  ['2026-07', 'IVRY', '', '', '', 'OK', '', '', d_(2026, 7, 9)],
  ['2026-07', 'Claude API', '', '', '', '未確認', '前回確認 2026/06/01(組織API上限$100・ヤップAPI上限$45)', '', ''],
  ['2026-07', 'Claude', 'Max20x 8 / Max5x 4 / Pro 6', '', '', 'OK', '', '', d_(2026, 7, 9)],
  ['2026-07', 'ChatGPT', '旧シート記載: Plusx20 1 / Plusx5 3 / Pro 4', '', '', 'OK', 'プラン名との不整合につき人数要確認', '', d_(2026, 7, 9)],
  ['2026-07', 'Bitwarden', 24, '', '', 'OK', '', '', d_(2026, 7, 6)],
  ['2026-07', 'GoogleWorkspace', 30, '', '', 'OK', '', '', d_(2026, 7, 9)],
  ['2026-07', 'tl;dv', 5, '', '', 'OK', 'Pro5 → 解約予定', '', d_(2026, 7, 9)]
];

function createLogSheet_(ss) {
  const sheet = ss.insertSheet(SHEET_LOG);
  const numCols = LOG_HEADERS.length;
  const formulaRows = 500; // 将来の追記行にも数式を先置きしておく

  sheet.getRange(1, 1, 1, numCols).setValues([LOG_HEADERS]);
  sheet.getRange(2, 1, LOG_SEED.length, numCols).setValues(LOG_SEED);

  // D列: 前月値(同ツールの前月行を検索) / E列: 増減
  const formulas = [];
  for (let r = 2; r <= formulaRows + 1; r++) {
    formulas.push([
      '=IF($A' + r + '="","",IFERROR(INDEX(FILTER($C$2:$C$' + (formulaRows + 1) + ',' +
      '$B$2:$B$' + (formulaRows + 1) + '=$B' + r + ',' +
      '$A$2:$A$' + (formulaRows + 1) + '=TEXT(EDATE(DATEVALUE($A' + r + '&"-01"),-1),"yyyy-mm")),1),""))',
      '=IF(OR($A' + r + '="",NOT(ISNUMBER($C' + r + ')),NOT(ISNUMBER($D' + r + '))),"",$C' + r + '-$D' + r + ')'
    ]);
  }
  sheet.getRange(2, 4, formulaRows, 2).setFormulas(formulas);

  setListValidation_(sheet, 'F2:F' + (formulaRows + 1), LIST_CHECK_RESULT);
  sheet.getRange(2, 9, formulaRows, 1).setNumberFormat('yyyy/mm/dd');

  styleHeader_(sheet, numCols);
  sheet.getRange(1, 1, LOG_SEED.length + 1, numCols).createFilter();

  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(7, 300);

  sheet.getRange('A1').setNote(
    '毎月、月次チェック対象ツールの行を追加して記録する(上書きしない)。\n' +
    'Step1の自動化で、月初にその月の行を自動生成する予定。\n年月は「2026-07」形式で入力。'
  );
}

// ---------------------------------------------------------------------------
// 4. 在籍者マスタ
// ---------------------------------------------------------------------------

const MEMBERS_HEADERS = [
  '氏名', 'メールアドレス', '区分', '在籍状況', '入社日', '退職日', '備考'
];

function createMembersSheet_(ss) {
  const sheet = ss.insertSheet(SHEET_MEMBERS);
  sheet.getRange(1, 1, 1, MEMBERS_HEADERS.length).setValues([MEMBERS_HEADERS]);

  setListValidation_(sheet, 'C2:C200', ['社員', '業務委託', 'その他']);
  setListValidation_(sheet, 'D2:D200', ['在籍', '退職']);
  sheet.getRange('E2:F200').setNumberFormat('yyyy/mm/dd');

  styleHeader_(sheet, MEMBERS_HEADERS.length);
  sheet.setColumnWidth(2, 240);

  sheet.getRange('A1').setNote(
    '退職者突合(各SaaSのユーザー一覧との照合)の基準となる在籍者リスト。\n' +
    'Step2でGoogle Workspace(Admin SDK)から自動取得予定。それまでは手動で入力。'
  );
}

// ---------------------------------------------------------------------------
// 5. ダッシュボード
// ---------------------------------------------------------------------------

function createDashboardSheet_(ss) {
  const sheet = ss.insertSheet(SHEET_DASHBOARD);
  const M = "'" + SHEET_MASTER + "'!";
  const L = "'" + SHEET_LOG + "'!";

  sheet.getRange('A1').setValue('SaaS管理ダッシュボード(自動集計・手入力不要)')
    .setFontWeight('bold').setFontSize(14);

  // --- 月額費用サマリー ---
  sheet.getRange('A3').setValue('■ 月額費用(円換算)').setFontWeight('bold').setBackground(COLOR_SECTION_BG);
  sheet.getRange('A4:A8').setValues([['合計'], ['01_ユーザー課金'], ['02_開発基盤'], ['03_事務・固定費'], ['契約中ツール数']]);
  sheet.getRange('B4').setFormula('=SUM(' + M + 'J2:J1000)');
  sheet.getRange('B5').setFormula('=SUMIF(' + M + 'C2:C1000,"01_ユーザー課金",' + M + 'J2:J1000)');
  sheet.getRange('B6').setFormula('=SUMIF(' + M + 'C2:C1000,"02_開発基盤",' + M + 'J2:J1000)');
  sheet.getRange('B7').setFormula('=SUMIF(' + M + 'C2:C1000,"03_事務・固定費",' + M + 'J2:J1000)');
  sheet.getRange('B8').setFormula('=COUNTA(UNIQUE(FILTER(' + M + 'A2:A1000,' + M + 'E2:E1000="契約中")))');
  sheet.getRange('B4:B7').setNumberFormat('¥#,##0');
  sheet.getRange('A9').setValue('※ 従量・プリペイド分(AWS、Claude API等)と未入力の料金は含まない');

  // --- 更新期限アラート ---
  sheet.getRange('D3').setValue('■ 60日以内に更新日が来る契約').setFontWeight('bold').setBackground(COLOR_SECTION_BG);
  sheet.getRange('D4:G4').setValues([['ツール名', 'プラン', '次回更新日', 'サイクル']]).setFontWeight('bold');
  sheet.getRange('D5').setFormula(
    '=IFERROR(FILTER({' + M + 'A2:A1000,' + M + 'B2:B1000,' + M + 'L2:L1000,' + M + 'K2:K1000},' +
    'ISNUMBER(' + M + 'L2:L1000),' + M + 'L2:L1000<=TODAY()+60,' + M + 'E2:E1000<>"解約済"),"該当なし")'
  );
  sheet.getRange('F5:F30').setNumberFormat('yyyy/mm/dd');

  // --- 当月の未チェック ---
  sheet.getRange('I3').setValue('■ 今月の未チェック').setFontWeight('bold').setBackground(COLOR_SECTION_BG);
  sheet.getRange('I4').setValue('ツール名').setFontWeight('bold');
  sheet.getRange('I5').setFormula(
    '=IFERROR(FILTER(' + L + 'B2:B2000,' + L + 'A2:A2000=TEXT(TODAY(),"yyyy-mm"),' + L + 'F2:F2000="未確認"),"なし(今月の行が未作成の場合もなしと表示)")'
  );

  // --- 前月比の変動 ---
  sheet.getRange('K3').setValue('■ 前月比で数量が変動したツール(今月分)').setFontWeight('bold').setBackground(COLOR_SECTION_BG);
  sheet.getRange('K4:N4').setValues([['ツール名', '今月', '前月', '増減']]).setFontWeight('bold');
  sheet.getRange('K5').setFormula(
    '=IFERROR(FILTER({' + L + 'B2:B2000,' + L + 'C2:C2000,' + L + 'D2:D2000,' + L + 'E2:E2000},' +
    L + 'A2:A2000=TEXT(TODAY(),"yyyy-mm"),ISNUMBER(' + L + 'E2:E2000),' + L + 'E2:E2000<>0),"変動なし")'
  );

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(9, 220);
  sheet.setColumnWidth(11, 160);
  sheet.setHiddenGridlines(true);
}
