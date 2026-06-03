const PROPERTIES = PropertiesService.getScriptProperties()
const SPREADSHEET_ID = PROPERTIES.getProperty("SPREADSHEET_ID");

// viewer側スプレッドシートのシート名.
const VIEWER_SHEET_NAME = {
  LIST: "名簿一覧",
  COPY_BY_CLASS: "コピペ用名簿",
  SORT_AND_COPY: "ソート名簿コピペ",
};

// ヘッダー文字列の定数. masterスプレッドシートの全会員名簿および正准会員名簿の1行目と一致させる.
const MEMBER_HEADER = {
  LAST_NAME: "姓",
  FIRST_NAME: "名",
  LAST_NAME_FURIGANA: "姓ふりがな",
  FIRST_NAME_FURIGANA: "名ふりがな",
  SCHOOL_YEAR: "学年",
  CURRENT_CLASS: "現在の級",
  NEWEST_CLASS_UP_DATE: "最新昇級日",
  STATUS: "会員区分",
  NOTE: "備考",
  JOIN_DATE: "入会日",
  LEAVE_DATE: "休退会日",
  SCHOOL_YEAR_SORT: "学年ソート用",
};

// 会員区分.
const MEMBER_STATUS = {
  REGULAR: "正会員",
  ASSOCIATE: "准会員",
  MATE: "メイト会員",
  LEFT: "休退会",
};

/**
 * シートのヘッダー名と列インデックス（1始まり）のマッピングを取得する.
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 * @param {number} [headerRow=1] - ヘッダーが存在する行番号（1始まり）
 * @return {Object} - ヘッダー名と列番号（1始まり）のマッピングオブジェクト
 */
function getColumnMap(sheet, headerRow = 1) {
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    map[header] = index + 1;
  });
  return map;
}

/**
 * 学年文字列に対応するソート用の並び順インデックスを取得する.
 * @param {string} schoolYearVal - 学年データ
 * @return {number} - ソート用の数値（小さいほど先）
 */
function getSchoolYearSortKey(schoolYearVal) {
  const order = [
    "小1", "小2", "小3", "小4", "小5", "小6",
    "中1", "中2", "中3",
    "高1", "高2", "高3",
    "大1", "専1", "大2", "専2",
    "大3", "専3", "大4", "専4",
    "大学院",
    "一般"
  ];
  if (schoolYearVal == null) {
    return 99;
  }
  const trimmed = String(schoolYearVal).trim();
  if (trimmed === "") {
    return 99;
  }
  const index = order.indexOf(trimmed);
  return index !== -1 ? index + 1 : 99;
}
