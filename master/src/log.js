/**
 * ログバッファ用のグローバル配列.
 */
let logBuffer = [];

/**
 * 情報ログを記録する.
 * @param {string} message - ログメッセージ
 */
function logInfo(message) {
  console.log(message);
  logBuffer.push([new Date(), "INFO", message]);
}

/**
 * 警告ログを記録する.
 * @param {string} message - ログメッセージ
 */
function logWarn(message) {
  console.warn(message);
  logBuffer.push([new Date(), "WARN", message]);
}

/**
 * エラーログを記録する.
 * @param {string} message - ログメッセージ
 */
function logError(message) {
  console.error(message);
  logBuffer.push([new Date(), "ERROR", message]);
}

/**
 * バッファに溜まったログをスプレッドシートへ一括書き込みし, バッファをクリアする.
 */
function flushLogs() {
  if (logBuffer.length === 0) {
    return;
  }

  const lock = LockService.getScriptLock();
  try {
    // 10秒間ロックを試行する.
    lock.waitLock(10000);

    const spreadSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = spreadSheet.getSheetByName(SPREADSHEET_NAME.LOG);
    if (logSheet == null) {
      logSheet = spreadSheet.insertSheet(SPREADSHEET_NAME.LOG);
      logSheet.appendRow(["タイムスタンプ", "ログレベル", "ログ内容"]);
    }

    // ログを一括追記する.
    const startRow = logSheet.getLastRow() + 1;
    const numRows = logBuffer.length;
    const numCols = 3;
    logSheet.getRange(startRow, 1, numRows, numCols).setValues(logBuffer);

    // 最大1000行に制限し, 超過した古いログを削除する.
    const lastRow = logSheet.getLastRow();
    if (lastRow > 1001) {
      const deleteCount = lastRow - 1001;
      logSheet.deleteRows(2, deleteCount);
    }
  } catch (error) {
    console.error("ログのフラッシュ中にエラーが発生しました: " + error.message);
  } finally {
    lock.releaseLock();
    logBuffer = [];
  }
}
