/**
 * スプレッドシートが開かれたときに実行されるトリガー.
 * メニューの追加やシートの初期化, サイドバーの自動表示を行う.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("名簿ビューア")
    .addItem("操作パネル（サイドバー）を開く", "showSidebar")
    .addItem("【緊急用】マスターから手動同期", "syncFromMaster")
    .addSeparator()
    .addItem("【管理者用】自動同期トリガーを設定", "setupAutoSyncTrigger")
    .addItem("【管理者用】自動同期トリガーを解除", "clearAutoSyncTrigger")
    .addToUi();

  // スプレッドシート起動時にサイドバーを自動で表示する.
  // 事務作業時の利便性を高めるためである.
  showSidebar();

  // 必要なシートが存在しない場合は自動的に初期化を行う.
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = spreadsheet.getSheetByName(VIEWER_SHEET_NAME.LIST);
  if (listSheet == null) {
    initializeSheets();
  }
}

/**
 * 操作パネル（サイドバー）を画面右側に表示する.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("sidebar")
    .setTitle("名簿ビューア連携");
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 必要なシートを作成し, 初期レイアウトを設定する.
 */
function initializeSheets() {
  console.log("initializeSheetsを開始します.");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 名簿一覧シートの初期化.
  let listSheet = spreadsheet.getSheetByName(VIEWER_SHEET_NAME.LIST);
  if (listSheet == null) {
    listSheet = spreadsheet.insertSheet(VIEWER_SHEET_NAME.LIST);
  }
  listSheet.clear();

  // 2. コピペ用名簿シートの初期化.
  let copySheet = spreadsheet.getSheetByName(VIEWER_SHEET_NAME.COPY_BY_CLASS);
  if (copySheet == null) {
    copySheet = spreadsheet.insertSheet(VIEWER_SHEET_NAME.COPY_BY_CLASS);
  }
  copySheet.clear();
  // 各列の幅を設定する.
  copySheet.setColumnWidth(1, 280);
  copySheet.setColumnWidth(2, 280);
  copySheet.setColumnWidth(3, 280);

  copySheet.getRange("A1").setValue("LINE更新用名簿プレビュー:").setFontWeight("bold");
  copySheet.getRange("B1").setValue("正准名簿プレビュー:").setFontWeight("bold");
  copySheet.getRange("C1").setValue("メイト名簿プレビュー:").setFontWeight("bold");

  // 3. ソート名簿コピペシートの初期化.
  let sortSheet = spreadsheet.getSheetByName(VIEWER_SHEET_NAME.SORT_AND_COPY);
  if (sortSheet == null) {
    sortSheet = spreadsheet.insertSheet(VIEWER_SHEET_NAME.SORT_AND_COPY);
  }
  sortSheet.clear();

  console.log("シートの初期設定が完了しました.");
  spreadsheet.toast("シートの初期化が完了しました. トリガーを設定するか、手動で同期を行ってください.", "初期化完了", 5);
}


/**
 * クリップボードコピー用の一時モーダルダイアログを表示する.
 * @param {string} text - コピー対象のテキスト
 * @param {string} title - ダイアログのタイトル
 */
function showClipboardDialog(text, title) {
  const template = HtmlService.createTemplateFromFile("clipboard");
  template.text = text;
  
  const htmlOutput = template.evaluate()
    .setWidth(350)
    .setHeight(200);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}

/**
 * LINE更新用の名簿を生成してコピーする.
 */
function copyLineUpdateList() {
  const text = generateLineUpdateListText();
  
  // プレビュー用にスプレッドシートのセルにも書き込む.
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const copySheet = spreadsheet.getSheetByName(VIEWER_SHEET_NAME.COPY_BY_CLASS);
  if (copySheet != null) {
    copySheet.getRange("A2").setValue(text);
  }

  if (text === "") {
    spreadsheet.toast("会員データがありません.", "コピー中止", 5);
    return;
  }

  showClipboardDialog(text, "LINE更新用名簿のコピー");
}

/**
 * 正会員・准会員の級別コピペ名簿を生成してコピーする.
 */
function copyClassListRegular() {
  const text = generateRegularAssociateClassListText();
  
  // プレビュー用にスプレッドシートのセルにも書き込む.
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const copySheet = spreadsheet.getSheetByName(VIEWER_SHEET_NAME.COPY_BY_CLASS);
  if (copySheet != null) {
    copySheet.getRange("B2").setValue(text);
  }

  if (text === "") {
    spreadsheet.toast("正准会員のデータがありません.", "コピー中止", 5);
    return;
  }

  showClipboardDialog(text, "正准会員名簿のコピー");
}

/**
 * メイト会員の級別コピペ名簿を生成してコピーする.
 */
function copyClassListMate() {
  const text = generateMateClassListText();
  
  // プレビュー用にスプレッドシートのセルにも書き込む.
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const copySheet = spreadsheet.getSheetByName(VIEWER_SHEET_NAME.COPY_BY_CLASS);
  if (copySheet != null) {
    copySheet.getRange("C2").setValue(text);
  }

  if (text === "") {
    spreadsheet.toast("メイト会員のデータがありません.", "コピー中止", 5);
    return;
  }

  showClipboardDialog(text, "メイト会員名簿のコピー");
}

/**
 * 現在のソート順の名簿を生成してコピーする.
 */
function copySortedList() {
  const text = generateSortedListText();
  
  if (text === "") {
    SpreadsheetApp.getActiveSpreadsheet().toast("コピー対象のデータがありません.", "コピー中止", 5);
    return;
  }

  showClipboardDialog(text, "ソート名簿のコピー");
}
