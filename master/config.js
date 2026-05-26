const SPREADSHEET_ID = '1WKk78wsoCe4MNS8btdwAvvq0QGAxR-F3xZoc4ZjfBbM';
// LINE Developersで取得したチャネルアクセストークンを定義する.
const LINE_ACCESS_TOKEN = "YOUR_LINE_ACCESS_TOKEN";

// シート名の定数.
const SPREADSHEET_NAME = {
  ALL_MEMBER: '全会員名簿',
  ACTIVE_MEMBER: '正准会員名簿',
  PROMOTION_HISTORY: '昇級履歴',
  STATUS_HISTORY: 'ステータス履歴',
  LOG: '実行ログ',
};

// ヘッダー文字列の定数. スプレッドシートの1行目と一致させる.
// 全会員名簿, 正准会員名簿.
const MEMBER_HEADER = {
  LAST_NAME: '姓',
  FIRST_NAME: '名',
  LAST_NAME_FURIGANA: '姓ふりがな',
  FIRST_NAME_FURIGANA: '名ふりがな',
  SCHOOL_YEAR: '学年',
  CURRENT_CLASS: '現在の級',
  NEWEST_CLASS_UP_DATE: '最新昇級日',
  STATUS: '会員区分',
  NOTE: '備考',
  JOIN_DATE: '入会日',
  LEAVE_DATE: '休退会日',
};

// ヘッダー文字列の定数. スプレッドシートの1行目と一致させる.
// 昇級履歴.
const PROMOTION_HEADER = {
  TIMESTAMP: 'タイムスタンプ',
  LAST_NAME: '姓',
  FIRST_NAME: '名',
  OLD_CLASS: '以前の級',
  NEW_CLASS: '新しい級',
  CLASS_UP_DATE: '昇級日',
  NOTE: '備考',
};


// ヘッダー文字列の定数. スプレッドシートの1行目と一致させる.
// ステータス履歴.
const STATUS_HEADER = {
  TIMESTAMP: 'タイムスタンプ',
  LAST_NAME: '姓',
  FIRST_NAME: '名',
  OLD_STATUS: '以前のステータス',
  NEW_STATUS: '新ステータス',
  STATUS_UP_DATE: '変更日',
  NOTE: '備考',
};

// 会員区分.
const MEMBER_STATUS = {
  REGULAR: '正会員',
  ASSOCIATE: '准会員',
  MATE: 'メイト会員',
  LEFT: '休退会',
  NONE: '(新規入会)'
};

// 使い方説明のメッセージテキスト.
const HELP_MESSAGE = "【名簿更新システム 使い方説明】\n" +
  "LINEから特定の形式でメッセージを送信することで, 会員名簿スプレッドシートを更新できます.\n\n" +
  "◆コマンド一覧\n" +
  "各コマンドの「★」から始まる行および各項目を入力して送信してください.\n\n" +
  "1. 昇級処理\n" +
  "★昇級★\n" +
  "フルネーム：[名前]\n" +
  "新級：[新級] (A, B, C, D, E, F, G, 基本, 入門)\n" +
  "日付：[適用日] (例: 12/20. 空欄の場合は今日となります)\n\n" +
  "2. 会員区分変更\n" +
  "★区分変更★\n" +
  "フルネーム：[名前]\n" +
  "新区分：[新区分] (正会員, 准会員, メイト会員)\n" +
  "日付：[適用日] (例: 12/20. 空欄の場合は今日となります)\n\n" +
  "3. 休退会処理\n" +
  "★休退会★\n" +
  "フルネーム：[名前]\n" +
  "日付：[適用日] (例: 12/20. 空欄の場合は今日となります)\n\n" +
  "4. 新規入会/復帰\n" +
  "★入会★ または ★復帰★\n" +
  "姓：[名字]\n" +
  "名：[名前]\n" +
  "姓ふりがな：[姓ひらがな]\n" +
  "名ふりがな：[名ひらがな]\n" +
  "学年：[学年] (例: 小1, 中2, 大1)\n" +
  "級：[現在の級]\n" +
  "区分：[会員区分] (正会員, 准会員, メイト会員)\n" +
  "日付：[適用日] (任意)\n";

/**
 * シートのヘッダー名と列インデックスのマッピングを取得する.
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 * @return {Object} - ヘッダー名と列番号（1始まり）のマッピングオブジェクト
 */
function getColumnMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    map[header] = index + 1;
  });
  return map;
}
