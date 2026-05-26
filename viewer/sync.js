/**
 * master名簿スプレッドシートからアクティブメンバーデータを同期する.
 * @return {boolean} - 同期処理が成功すればtrue, そうでなければfalse
 */
function syncFromMaster() {
  console.log("syncFromMasterを開始します.");
  const lock = LockService.getScriptLock();
  try {
    // 30秒間ロックを試行する. 他の更新処理との競合を防ぐためである.
    lock.waitLock(30000);
    console.log("スクリプトロックを取得しました.");

    // masterのスプレッドシートを開く.
    const masterSpreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    // master側の全会員名簿シート名（ハードコードではなくmasterの仕様に合わせる）.
    const masterSheet = masterSpreadsheet.getSheetByName("全会員名簿");

    if (masterSheet == null) {
      console.error("エラー: master側に「全会員名簿」シートが見つかりません.");
      return false;
    }

    const masterData = masterSheet.getDataRange().getValues();
    const masterColMap = getColumnMap(masterSheet);
    const header = masterData[0];

    // 同期対象（正会員, 准会員, メイト会員）のメンバーを抽出する.
    // 学年ソート用のキーを末尾に追加するため, ヘッダーの末尾にキー名を追加する.
    const extendedHeader = [...header, MEMBER_HEADER.SCHOOL_YEAR_SORT];
    const activeMembers = [extendedHeader];
    
    const statusColIndex = masterColMap[MEMBER_HEADER.STATUS] - 1;
    const schoolYearColIndex = masterColMap[MEMBER_HEADER.SCHOOL_YEAR] - 1;

    for (let i = 1; i < masterData.length; i++) {
      const row = masterData[i];
      const status = row[statusColIndex];
      if (
        status === MEMBER_STATUS.REGULAR ||
        status === MEMBER_STATUS.ASSOCIATE ||
        status === MEMBER_STATUS.MATE
      ) {
        // 元のデータをコピーし, 末尾に学年ソートキーを付加する.
        const schoolYear = row[schoolYearColIndex];
        const sortKey = getSchoolYearSortKey(schoolYear);
        activeMembers.push([...row, sortKey]);
      }
    }

    console.log("同期対象アクティブメンバー数: " + (activeMembers.length - 1));

    // viewer側のスプレッドシートに書き込む.
    const localSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheetNames = [VIEWER_SHEET_NAME.LIST, VIEWER_SHEET_NAME.SORT_AND_COPY];

    targetSheetNames.forEach((sheetName) => {
      let sheet = localSpreadsheet.getSheetByName(sheetName);
      if (sheet == null) {
        // シートが存在しない場合は新規作成する.
        sheet = localSpreadsheet.insertSheet(sheetName);
      }

      // 既存のフィルターを一旦削除する.
      // データのクリアや書き込みによってフィルター範囲が壊れるのを防ぐためである.
      const filter = sheet.getFilter();
      if (filter != null) {
        filter.remove();
      }

      // 既存のコンテンツをクリアする.
      sheet.clear();

      if (activeMembers.length > 0) {
        // 1行目からヘッダーとデータを書き込む.
        const range = sheet.getRange(1, 1, activeMembers.length, activeMembers[0].length);
        range.setValues(activeMembers);

        // 新しいデータ範囲に対してフィルターを作成する.
        // スプレッドシート標準のGUIフィルター機能を利用できるようにするためである.
        range.createFilter();

        // 学年順に自動でソートする.
        const colMap = getColumnMap(sheet, 1);
        const sortColIndex = colMap[MEMBER_HEADER.SCHOOL_YEAR_SORT];
        if (sortColIndex != null) {
          const filter = sheet.getFilter();
          if (filter != null) {
            filter.sort(sortColIndex, true);
          }
        }
      }
    });

    // コピペ用名簿シートにプレビューテキストを自動で書き出す.
    const copySheet = localSpreadsheet.getSheetByName(VIEWER_SHEET_NAME.COPY_BY_CLASS);
    if (copySheet != null) {
      const lineUpdateText = generateLineUpdateListText();
      const regularText = generateRegularAssociateClassListText();
      const mateText = generateMateClassListText();

      // 1行目はヘッダーなので、2行目にそれぞれ値を設定する.
      copySheet.getRange("A2").setValue(lineUpdateText);
      copySheet.getRange("B2").setValue(regularText);
      copySheet.getRange("C2").setValue(mateText);
      console.log("コピペ用名簿のプレビューを更新しました.");
    }

    console.log("同期処理が完了しました.");
    localSpreadsheet.toast("マスター名簿からの同期が完了しました.", "同期完了", 5);
    return true;
  } catch (error) {
    console.error("syncFromMasterエラー: 同期処理中に例外が発生しました. 原因: " + error.message);
    const localSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    localSpreadsheet.toast("同期処理中にエラーが発生しました.", "エラー", 5);
    return false;
  } finally {
    lock.releaseLock();
    console.log("スクリプトロックを解放しました.");
  }
}

/**
 * 現在開いているシートを学年ソート用列に基づいてソートする.
 */
function sortActiveSheetBySchoolYear() {
  console.log("sortActiveSheetBySchoolYearを開始します.");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();

  // 名簿一覧またはソート名簿コピペシート以外では実行しない.
  if (sheetName !== VIEWER_SHEET_NAME.LIST && sheetName !== VIEWER_SHEET_NAME.SORT_AND_COPY) {
    console.warn("ソート対象外のシートです: " + sheetName);
    return;
  }

  const colMap = getColumnMap(sheet, 1);
  const sortColIndex = colMap[MEMBER_HEADER.SCHOOL_YEAR_SORT];
  
  if (sortColIndex == null) {
    console.error("学年ソート用列が見つかりません.");
    return;
  }

  // 既存のフィルターを取得してソートする.
  const filter = sheet.getFilter();
  if (filter != null) {
    // 列インデックス（1始まり）を指定して昇順ソート.
    filter.sort(sortColIndex, true);
    console.log("フィルターを使ったソートを完了しました. 列: " + sortColIndex);
  } else {
    // フィルターがない場合はデータ範囲をソートする.
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow >= 2) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).sort({ column: sortColIndex, ascending: true });
      console.log("範囲指定ソートを完了しました. 列: " + sortColIndex);
    }
  }
}

/**
 * 自動同期トリガーを設定する.
 */
function setupAutoSyncTrigger() {
  // 重複登録を避けるために既存のトリガーを解除する.
  clearAutoSyncTrigger();

  // 毎日午前4時〜5時の間に実行するトリガーを登録する.
  ScriptApp.newTrigger("syncFromMaster")
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .create();

  console.log("自動同期トリガーを設定しました. 毎日午前4時〜5時の間に実行されます.");
  SpreadsheetApp.getActiveSpreadsheet().toast("毎日午前4時〜5時の間の自動同期トリガーを設定しました.", "トリガー設定完了", 5);
}

/**
 * 自動同期トリガーを解除する.
 */
function clearAutoSyncTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let isCleared = false;
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === "syncFromMaster") {
      ScriptApp.deleteTrigger(trigger);
      isCleared = true;
    }
  });

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (isCleared) {
    console.log("自動同期トリガーを解除しました.");
    spreadsheet.toast("自動同期トリガーを解除しました.", "トリガー解除完了", 5);
  } else {
    spreadsheet.toast("設定されている自動同期トリガーはありません.", "トリガー解除", 5);
  }
}
