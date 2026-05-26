/**
 * 昇級処理を実行するメイン関数.
 * @param {string} lastName - 姓
 * @param {string} firstName - 名
 * @param {string} nextLevel - 新しい級
 * @param {Date} [date] - 昇級日
 * @param {string} [note] - 備考
 * @return {boolean} - 手続きに成功すればtrue, そうでなければfalse
 */
function processPromotion(
  lastName, firstName, nextLevel, date = new Date(), note = "GASによる自動更新"
) {
  console.log("processPromotionを開始します. 姓名: " + lastName + " " + firstName + ", 新級: " + nextLevel + ", 日付: " + date + ", 備考: " + note);
  const lock = LockService.getScriptLock();
  try {
    // 30秒間ロックを試行する.
    console.log("スクリプトロックを取得します（最大30秒待機）...");
    lock.waitLock(30000);
    console.log("スクリプトロックを取得しました.");

    // 日付データを文字列化する.
    const formattedDate = Utilities.formatDate(date, "JST", "yyyy/MM/dd");
    console.log("フォーマット後日付: " + formattedDate);

    console.log("スプレッドシートを開きます. ID: " + SPREADSHEET_ID);
    const spreadSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const allSheet = spreadSheet.getSheetByName(SPREADSHEET_NAME.ALL_MEMBER);
    const historySheet = spreadSheet.getSheetByName(SPREADSHEET_NAME.PROMOTION_HISTORY);

    if (allSheet == null) {
      logError("processPromotionエラー: 「" + SPREADSHEET_NAME.ALL_MEMBER + "」シートが見つかりません.");
      return false;
    }
    if (historySheet == null) {
      logError("processPromotionエラー: 「" + SPREADSHEET_NAME.PROMOTION_HISTORY + "」シートが見つかりません.");
      return false;
    }

    const allData = allSheet.getDataRange().getValues();
    const col = getColumnMap(allSheet);

    // 1. 全会員名簿から対象者を検索する.
    console.log("メンバーを検索します...");
    const member = getMemberInfo(allData, col, lastName, firstName);
    if (member == null) {
      logWarn("processPromotion警告: " + lastName + " " + firstName + " さんが見つかりませんでした。");
      return false;
    }
    console.log("メンバーが見つかりました. 行番号: " + member.rowIndex + ", 現在の級: " + member.rowData[col[MEMBER_HEADER.CURRENT_CLASS] - 1]);

    // 既存の最新昇級日を取得して比較する.
    const currentNewestDateVal = member.rowData[col[MEMBER_HEADER.NEWEST_CLASS_UP_DATE] - 1];
    let isNewer = true;
    if (currentNewestDateVal != null && currentNewestDateVal !== "") {
      const currentNewestDate = new Date(currentNewestDateVal);
      if (!isNaN(currentNewestDate.getTime())) {
        // 指定された日付 (date) が既存の最新昇級日以前である場合, 名簿更新をスキップする.
        if (date <= currentNewestDate) {
          isNewer = false;
        }
      }
    }

    if (isNewer) {
      // 2. 全会員名簿を更新する.
      console.log("全会員名簿を更新します. 行: " + member.rowIndex + ", 新級: " + nextLevel + ", 更新日: " + formattedDate);
      allSheet.getRange(member.rowIndex, col[MEMBER_HEADER.CURRENT_CLASS]).setValue(nextLevel);
      allSheet.getRange(member.rowIndex, col[MEMBER_HEADER.NEWEST_CLASS_UP_DATE]).setValue(formattedDate);
    } else {
      console.log("指定された昇級日（" + formattedDate + "）は登録済みの最新昇級日（" + currentNewestDateVal + "）以前であるため, 名簿シートの現在の級および最新昇級日の更新はスキップします.");
    }

    // 3. 昇級履歴シートへ追記する.
    console.log("昇級履歴シートへ追記データを構築します...");
    const histColMap = getColumnMap(historySheet);
    const lastCol = historySheet.getLastColumn();
    const newRowData = new Array(lastCol).fill("");

    newRowData[histColMap[PROMOTION_HEADER.TIMESTAMP] - 1] = new Date();
    newRowData[histColMap[PROMOTION_HEADER.LAST_NAME] - 1] = lastName;
    newRowData[histColMap[PROMOTION_HEADER.FIRST_NAME] - 1] = firstName;
    newRowData[histColMap[PROMOTION_HEADER.OLD_CLASS] - 1] = member.rowData[col[MEMBER_HEADER.CURRENT_CLASS] - 1];
    newRowData[histColMap[PROMOTION_HEADER.NEW_CLASS] - 1] = nextLevel;
    newRowData[histColMap[PROMOTION_HEADER.CLASS_UP_DATE] - 1] = formattedDate;
    newRowData[histColMap[PROMOTION_HEADER.NOTE] - 1] = note;

    console.log("昇級履歴シートにレコードを追加します. データ: " + JSON.stringify(newRowData));
    historySheet.appendRow(newRowData);

    console.log(`${lastName}${firstName}さんの昇級処理を完了しました。`);

    if (isNewer) {
      // 同期処理を実行する.
      console.log("アクティブメンバーシートとの同期処理（syncActiveMembers）を呼び出します...");
      const isSyncSuccess = syncActiveMembers();
      console.log("同期処理の実行結果: " + isSyncSuccess);
    }

    return true;
  } catch (error) {
    logError("processPromotionエラー: 昇級処理中に例外が発生しました. 原因: " + error.message + ", スタック: " + error.stack);
    return false;
  } finally {
    console.log("スクリプトロックを解放します.");
    lock.releaseLock();
    flushLogs();
  }
}

/**
 * 会員区分（休退会, 区分変更）を更新するメイン関数. 入会は別関数とする.
 * @param {string} lastName - 姓
 * @param {string} firstName - 名
 * @param {string} newStatus - 新しい会員区分（正会員, 准会員, 休会, 退会など）
 * @param {Date} [date] - 変更適用日
 * @param {string} [note] - 備考
 * @return {boolean} - 手続きに成功すればtrue, そうでなければfalse
 */
function updateMemberStatus(
  lastName, firstName, newStatus, date = new Date(), note = "ステータス更新(区分変更・休退会)"
) {
  console.log("updateMemberStatusを開始します. 姓名: " + lastName + " " + firstName + ", 新ステータス: " + newStatus + ", 日付: " + date + ", 備考: " + note);
  const lock = LockService.getScriptLock();
  try {
    // 30秒間ロックを試行する.
    console.log("スクリプトロックを取得します（最大30秒待機）...");
    lock.waitLock(30000);
    console.log("スクリプトロックを取得しました.");
    
    // 日付データを文字列化する.
    const formattedDate = Utilities.formatDate(date, "JST", "yyyy/MM/dd");
    console.log("フォーマット後日付: " + formattedDate);

    console.log("スプレッドシートを開きます. ID: " + SPREADSHEET_ID);
    const spreadSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const allSheet = spreadSheet.getSheetByName(SPREADSHEET_NAME.ALL_MEMBER);
    const statusHistorySheet = spreadSheet.getSheetByName(SPREADSHEET_NAME.STATUS_HISTORY);

    if (allSheet == null) {
      logError("updateMemberStatusエラー: 「" + SPREADSHEET_NAME.ALL_MEMBER + "」シートが見つかりません.");
      return false;
    }
    if (statusHistorySheet == null) {
      logError("updateMemberStatusエラー: 「" + SPREADSHEET_NAME.STATUS_HISTORY + "」シートが見つかりません.");
      return false;
    }

    const allData = allSheet.getDataRange().getValues();
    const col = getColumnMap(allSheet);

    // 1. 全会員名簿から対象者を検索する.
    console.log("メンバーを検索します...");
    const member = getMemberInfo(allData, col, lastName, firstName);
    if (member == null) {
      logWarn("updateMemberStatus警告: " + lastName + " " + firstName + " さんが見つかりませんでした。");
      return false;
    }
    console.log("メンバーが見つかりました. 行番号: " + member.rowIndex + ", 現在の区分: " + member.rowData[col[MEMBER_HEADER.STATUS] - 1]);

    // 2. 会員区分を更新する.
    console.log("会員区分を更新します. 行: " + member.rowIndex + ", 新区分: " + newStatus);
    allSheet.getRange(member.rowIndex, col[MEMBER_HEADER.STATUS]).setValue(newStatus);

    // 3. 休退会日を更新する.
    if (newStatus === MEMBER_STATUS.LEFT) {
      console.log("区分が休退会であるため, 休退会日を更新します: " + formattedDate);
      allSheet.getRange(member.rowIndex, col[MEMBER_HEADER.LEAVE_DATE]).setValue(formattedDate);
    }

    // 4. ステータス変更を「ステータス履歴シート」に記録する.
    console.log("ステータス履歴シートへの追記データを構築します...");
    const statusColMap = getColumnMap(statusHistorySheet);
    const lastCol = statusHistorySheet.getLastColumn();
    const newRowData = new Array(lastCol).fill("");

    newRowData[statusColMap[STATUS_HEADER.TIMESTAMP] - 1] = new Date();
    newRowData[statusColMap[STATUS_HEADER.LAST_NAME] - 1] = lastName;
    newRowData[statusColMap[STATUS_HEADER.FIRST_NAME] - 1] = firstName;
    newRowData[statusColMap[STATUS_HEADER.OLD_STATUS] - 1] = member.rowData[col[MEMBER_HEADER.STATUS] - 1];
    newRowData[statusColMap[STATUS_HEADER.NEW_STATUS] - 1] = newStatus;
    newRowData[statusColMap[STATUS_HEADER.STATUS_UP_DATE] - 1] = formattedDate;
    newRowData[statusColMap[STATUS_HEADER.NOTE] - 1] = note;

    console.log("ステータス履歴シートにレコードを追加します. データ: " + JSON.stringify(newRowData));
    statusHistorySheet.appendRow(newRowData);

    console.log(`${lastName}${firstName}さんの区分を「${newStatus}」に更新しました。`);

    // 同期処理を実行する.
    console.log("アクティブメンバーシートとの同期処理（syncActiveMembers）を呼び出します...");
    const isSyncSuccess = syncActiveMembers();
    console.log("同期処理の実行結果: " + isSyncSuccess);

    return true;
  } catch (error) {
    logError("updateMemberStatusエラー: ステータス更新中に例外が発生しました. 原因: " + error.message + ", スタック: " + error.stack);
    return false;
  } finally {
    console.log("スクリプトロックを解放します.");
    lock.releaseLock();
    flushLogs();
  }
}

/**
 * 入会者, 復帰者を更新するメイン関数.
 * @param {string} lastName - 姓
 * @param {string} firstName - 名
 * @param {string} lastNameFurigana - 姓ふりがな
 * @param {string} firstNameFurigana - 名ふりがな
 * @param {string} schoolYear - 学年
 * @param {string} karutaClass - 級
 * @param {string} newStatus - 新しい会員区分（正会員, 准会員）
 * @param {Date} [date] - 入会日/復帰日
 * @param {string} [note] - 備考
 * @return {boolean} - 手続きに成功すればtrue, そうでなければfalse
 */
function processJoin(
  lastName,
  firstName,
  lastNameFurigana,
  firstNameFurigana,
  schoolYear,
  karutaClass,
  newStatus,
  date = new Date(),
  note = "ステータス更新(入会・復帰)"
) {
  console.log("processJoinを開始します. 姓名: " + lastName + " " + firstName + ", ふりがな: " + lastNameFurigana + " " + firstNameFurigana + ", 学年: " + schoolYear + ", 級: " + karutaClass + ", 区分: " + newStatus);
  const lock = LockService.getScriptLock();
  try {
    // 30秒間ロックを試行する.
    console.log("スクリプトロックを取得します（最大30秒待機）...");
    lock.waitLock(30000);
    console.log("スクリプトロックを取得しました.");
    
    // 日付データを文字列化する.
    const formattedDate = Utilities.formatDate(date, "JST", "yyyy/MM/dd");
    console.log("フォーマット後日付: " + formattedDate);

    console.log("スプレッドシートを開きます. ID: " + SPREADSHEET_ID);
    const spreadSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const allSheet = spreadSheet.getSheetByName(SPREADSHEET_NAME.ALL_MEMBER);
    const statusHistorySheet = spreadSheet.getSheetByName(SPREADSHEET_NAME.STATUS_HISTORY);

    if (allSheet == null) {
      logError("processJoinエラー: 「" + SPREADSHEET_NAME.ALL_MEMBER + "」シートが見つかりません.");
      return false;
    }
    if (statusHistorySheet == null) {
      logError("processJoinエラー: 「" + SPREADSHEET_NAME.STATUS_HISTORY + "」シートが見つかりません.");
      return false;
    }

    const allData = allSheet.getDataRange().getValues();
    const col = getColumnMap(allSheet);

    // 1. 全会員名簿から対象者を検索する.
    console.log("既存メンバーを検索します...");
    const member = getMemberInfo(allData, col, lastName, firstName);
    const currentStatus = member != null ? member.rowData[col[MEMBER_HEADER.STATUS] - 1] : MEMBER_STATUS.NONE;
    
    // 見つかった場合は復帰処理を行う.
    if (member != null) {
      console.log("既存メンバーが見つかったため, 復帰処理を行います. 行番号: " + member.rowIndex + ", 旧区分: " + currentStatus + ", 新区分: " + newStatus);
      // 会員区分を更新する.
      allSheet.getRange(member.rowIndex, col[MEMBER_HEADER.STATUS]).setValue(newStatus);
      // 休退会日を空欄にする.
      allSheet.getRange(member.rowIndex, col[MEMBER_HEADER.LEAVE_DATE]).setValue("");
    } else {
      // 見つからない場合は新規入会処理を行う.
      console.log("メンバーが見つからないため, 新規入会登録を行います.");
      const colMap = getColumnMap(allSheet);
      const lastCol = allSheet.getLastColumn();
      const newRowData = new Array(lastCol).fill("");

      newRowData[colMap[MEMBER_HEADER.LAST_NAME] - 1] = lastName;
      newRowData[colMap[MEMBER_HEADER.FIRST_NAME] - 1] = firstName;
      newRowData[colMap[MEMBER_HEADER.LAST_NAME_FURIGANA] - 1] = lastNameFurigana;
      newRowData[colMap[MEMBER_HEADER.FIRST_NAME_FURIGANA] - 1] = firstNameFurigana;
      newRowData[colMap[MEMBER_HEADER.SCHOOL_YEAR] - 1] = schoolYear;
      newRowData[colMap[MEMBER_HEADER.CURRENT_CLASS] - 1] = karutaClass;
      newRowData[colMap[MEMBER_HEADER.STATUS] - 1] = newStatus;
      newRowData[colMap[MEMBER_HEADER.JOIN_DATE] - 1] = formattedDate;

      console.log("全会員名簿にレコードを追加します. データ: " + JSON.stringify(newRowData));
      allSheet.appendRow(newRowData);
    }

    // 4. ステータス変更を「ステータス履歴シート」に記録する.
    console.log("ステータス履歴シートへの追記データを構築します...");
    const statusColMap = getColumnMap(statusHistorySheet);
    const lastCol = statusHistorySheet.getLastColumn();
    const newRowData = new Array(lastCol).fill("");

    newRowData[statusColMap[STATUS_HEADER.TIMESTAMP] - 1] = new Date();
    newRowData[statusColMap[STATUS_HEADER.LAST_NAME] - 1] = lastName;
    newRowData[statusColMap[STATUS_HEADER.FIRST_NAME] - 1] = firstName;
    newRowData[statusColMap[STATUS_HEADER.OLD_STATUS] - 1] = currentStatus;
    newRowData[statusColMap[STATUS_HEADER.NEW_STATUS] - 1] = newStatus;
    newRowData[statusColMap[STATUS_HEADER.STATUS_UP_DATE] - 1] = formattedDate;
    newRowData[statusColMap[STATUS_HEADER.NOTE] - 1] = note;

    console.log("ステータス履歴シートにレコードを追加します. データ: " + JSON.stringify(newRowData));
    statusHistorySheet.appendRow(newRowData);

    console.log(`${lastName}${firstName}さんの区分を「${newStatus}」に更新しました。`);

    // 同期処理を実行する.
    console.log("アクティブメンバーシートとの同期処理（syncActiveMembers）を呼び出します...");
    const isSyncSuccess = syncActiveMembers();
    console.log("同期処理の実行結果: " + isSyncSuccess);

    return true;
  } catch (error) {
    logError("processJoinエラー: 入会/復帰処理中に例外が発生しました. 原因: " + error.message + ", スタック: " + error.stack);
    return false;
  } finally {
    console.log("スクリプトロックを解放します.");
    lock.releaseLock();
    flushLogs();
  }
}

/**
 * 正准会員名簿の同期処理を行う.
 * @return {boolean} - 同期に成功すればtrue, そうでなければfalse
 */
function syncActiveMembers() {
  console.log("syncActiveMembersを開始します. 正准会員名簿の同期を行います.");
  try {
    const spreadSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const allSheet = spreadSheet.getSheetByName(SPREADSHEET_NAME.ALL_MEMBER);
    const activeSheet = spreadSheet.getSheetByName(SPREADSHEET_NAME.ACTIVE_MEMBER);

    if (allSheet == null) {
      logError("syncActiveMembersエラー: 「" + SPREADSHEET_NAME.ALL_MEMBER + "」シートが見つかりません.");
      return false;
    }
    if (activeSheet == null) {
      logError("syncActiveMembersエラー: 「" + SPREADSHEET_NAME.ACTIVE_MEMBER + "」シートが見つかりません.");
      return false;
    }

    const data = allSheet.getDataRange().getValues();
    const col = getColumnMap(allSheet);
    const header = data[0];

    const activeMembers = [header];
    console.log("同期対象メンバーの抽出を開始します. 総行数: " + data.length);

    for (let i = 1; i < data.length; i++) {
      const category = data[i][col[MEMBER_HEADER.STATUS] - 1];
      if (category === MEMBER_STATUS.REGULAR || category === MEMBER_STATUS.ASSOCIATE) {
        activeMembers.push(data[i]);
      }
    }

    console.log("抽出完了. 同期対象の正准会員数: " + (activeMembers.length - 1));

    console.log("正准会員名簿シートをクリアします.");
    activeSheet.clearContents();
    if (activeMembers.length > 0) {
      console.log("正准会員名簿シートにデータを書き込みます. 行数: " + activeMembers.length + ", 列数: " + activeMembers[0].length);
      activeSheet.getRange(1, 1, activeMembers.length, activeMembers[0].length).setValues(activeMembers);
    }
    console.log("syncActiveMembers同期処理が完了しました.");
    return true;
  } catch (error) {
    logError("syncActiveMembersエラー: 同期処理中に例外が発生しました. 原因: " + error.message + ", スタック: " + error.stack);
    return false;
  }
}

/**
 * 共通：メンバーの行データとインデックスを取得する.
 * @param {Array<Array<any>>} allData - スプレッドシートの全行データ
 * @param {Object} colMap - 列名とインデックスのマッピングオブジェクト
 * @param {string} lastName - 姓
 * @param {string} firstName - 名
 * @return {Object|null} - メンバー情報オブジェクト, 見つからない場合は null
 */
function getMemberInfo(allData, colMap, lastName, firstName) {
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][colMap[MEMBER_HEADER.LAST_NAME] - 1] === lastName &&
      allData[i][colMap[MEMBER_HEADER.FIRST_NAME] - 1] === firstName) {
      return { rowIndex: i + 1, rowData: allData[i] };
    }
  }
  return null;
}

/**
 * フルネーム（スペースなし）から会員の姓と名を検索する.
 * @param {string} fullNameWithoutSpace - スペースを除去したフルネーム
 * @return {Object|null} - 姓と名を含むオブジェクト, 見つからない場合は null
 */
function findMemberByName(fullNameWithoutSpace) {
  if (fullNameWithoutSpace == null || fullNameWithoutSpace === "") {
    return null;
  }

  console.log("findMemberByNameを開始します. 検索名: " + fullNameWithoutSpace);
  try {
    const spreadSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const allSheet = spreadSheet.getSheetByName(SPREADSHEET_NAME.ALL_MEMBER);
    if (allSheet == null) {
      logError("findMemberByNameエラー: 「" + SPREADSHEET_NAME.ALL_MEMBER + "」シートが見つかりません.");
      return null;
    }

    const allData = allSheet.getDataRange().getValues();
    const col = getColumnMap(allSheet);
    const lastNameIdx = col[MEMBER_HEADER.LAST_NAME] - 1;
    const firstNameIdx = col[MEMBER_HEADER.FIRST_NAME] - 1;

    for (let i = 1; i < allData.length; i++) {
      const lastNameVal = allData[i][lastNameIdx] || "";
      const firstNameVal = allData[i][firstNameIdx] || "";
      const joinedName = (lastNameVal + firstNameVal).replace(/\s/g, "");
      if (joinedName === fullNameWithoutSpace) {
        console.log("一致するメンバーが見つかりました: " + lastNameVal + " " + firstNameVal);
        return { lastName: lastNameVal, firstName: firstNameVal };
      }
    }
  } catch (error) {
    logError("findMemberByNameエラー: メンバー検索中に例外が発生しました. 原因: " + error.message);
  }
  return null;
}