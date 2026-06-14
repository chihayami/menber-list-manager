/**
 * テスト実行用グローバル変数.
 */
let mockFetchedUrls = [];
let mockSavedProperties = {};
let mockLockAcquired = false;
let mockLockReleased = false;
let mockSetValueCalls = [];
let mockAppendRowCalls = [];

let originalPropertiesService;
let originalUrlFetchApp;
let originalSpreadsheetApp;
let originalLockService;
let originalUtilities;
let originalFlushLogs;

/**
 * テスト用のモックをセットアップする.
 */
function setupMocks() {
  originalPropertiesService = typeof PropertiesService !== 'undefined' ? PropertiesService : null;
  originalUrlFetchApp = typeof UrlFetchApp !== 'undefined' ? UrlFetchApp : null;
  originalSpreadsheetApp = typeof SpreadsheetApp !== 'undefined' ? SpreadsheetApp : null;
  originalLockService = typeof LockService !== 'undefined' ? LockService : null;
  originalUtilities = typeof Utilities !== 'undefined' ? Utilities : null;
  originalFlushLogs = typeof flushLogs !== 'undefined' ? flushLogs : null;

  mockFetchedUrls = [];
  mockSavedProperties = { LINE_TOKEN: "MOCK_LINE_TOKEN" };
  mockLockAcquired = false;
  mockLockReleased = false;
  mockSetValueCalls = [];
  mockAppendRowCalls = [];

  // グローバルモックの定義.
  PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => mockSavedProperties[key],
      getProperties: () => mockSavedProperties
    })
  };

  UrlFetchApp = {
    fetch: (url, params) => {
      mockFetchedUrls.push({ url, params });
      return {
        getContentText: () => "{}",
        getResponseCode: () => 200
      };
    }
  };

  const dummyValues = [
    ['姓', '名', '姓ふりがな', '名ふりがな', '学年', '現在の級', '最新昇級日', '会員区分', '備考', '入会日', '休退会日'],
    ['競技', '太郎', 'きょうぎ', 'たろう', '大1', 'C', '2026/05/20', '正会員', '', '2025/04/01', '']
  ];

  const mockSheet = {
    getDataRange: () => ({
      getValues: () => dummyValues
    }),
    getLastColumn: () => dummyValues[0].length,
    getRange: (row, col) => ({
      setValue: (val) => {
        mockSetValueCalls.push({ row, col, val });
      },
      setValues: () => {},
      getValues: () => [dummyValues[0]]
    }),
    appendRow: (rowData) => {
      mockAppendRowCalls.push(rowData);
    },
    clearContents: () => {}
  };

  SpreadsheetApp = {
    openById: () => ({
      getSheetByName: () => mockSheet
    })
  };

  LockService = {
    getScriptLock: () => ({
      waitLock: () => { mockLockAcquired = true; },
      releaseLock: () => { mockLockReleased = true; }
    })
  };

  Utilities = {
    formatDate: (date, tz, format) => {
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, '0');
      const d = String(date.getDate()).padStart(2, '0');
      return `${y}/${m}/${d}`;
    }
  };

  // テスト時はスプレッドシートへのフラッシュ処理をモック化し, バッファクリアのみを行う.
  flushLogs = () => {
    logBuffer = [];
  };
}

/**
 * モックを解除して元のオブジェクトを復元する.
 */
function teardownMocks() {
  if (originalPropertiesService != null) PropertiesService = originalPropertiesService;
  if (originalUrlFetchApp != null) UrlFetchApp = originalUrlFetchApp;
  if (originalSpreadsheetApp != null) SpreadsheetApp = originalSpreadsheetApp;
  if (originalLockService != null) LockService = originalLockService;
  if (originalUtilities != null) Utilities = originalUtilities;
  if (originalFlushLogs != null) flushLogs = originalFlushLogs;
}

/**
 * 簡易アサーション関数.
 * @param {any} actual - 実際値
 * @param {any} expected - 期待値
 * @param {string} message - エラーメッセージ
 */
function assertEquals(actual, expected, message) {
  if (actual !== expected) {
    throw new Error(`${message} - 期待値: ${expected}, 実際の値: ${actual}`);
  }
}

/**
 * テスト実行用メイン関数.
 */
function runTests() {
  console.log("=== テスト開始 ===");
  setupMocks();

  try {
    // 1. parseTextのテスト.
    console.log("テスト: parseText の検証");
    const parsed1 = parseText("★昇級★\n姓：競技\n名：太郎\n新級：B");
    assertEquals(parsed1.type, "昇級", "タイプが異なります");
    assertEquals(parsed1.params["姓"], "競技", "姓が異なります");
    assertEquals(parsed1.params["名"], "太郎", "名が異なります");
    assertEquals(parsed1.params["新級"], "B", "新級が異なります");

    const parsed2 = parseText("★休退会★  \n名: 太郎 \n  姓：競技  ");
    assertEquals(parsed2.type, "休退会", "タイプが異なります");
    assertEquals(parsed2.params["姓"], "競技", "姓の余白除去ができていません");
    assertEquals(parsed2.params["名"], "太郎", "名の余白除去ができていません");

    // 2. validateKarutaClassのテスト.
    console.log("テスト: validateKarutaClass の検証");
    assertEquals(validateKarutaClass("A"), "A", "アルファベットのパースに失敗しました");
    assertEquals(validateKarutaClass("ａ"), "A", "全角の半角化・大文字化に失敗しました");
    assertEquals(validateKarutaClass("B級"), "B", "「級」の除去に失敗しました");
    assertEquals(validateKarutaClass("入門"), "入門", "漢字の級のパースに失敗しました");
    assertEquals(validateKarutaClass("Z"), null, "無効な級が通過しています");

    // 3. validateMemberStatusのテスト.
    console.log("テスト: validateMemberStatus の検証");
    assertEquals(validateMemberStatus("正会員"), "正会員", "完全一致の検証に失敗しました");
    assertEquals(validateMemberStatus("正"), "正会員", "部分一致 of 補完に失敗しました");
    assertEquals(validateMemberStatus("メイト"), "メイト会員", "部分一致 of 補完に失敗しました");
    assertEquals(validateMemberStatus("一般"), null, "無効な会員区分が通過しています");

    // 4. parseDateParameterのテスト（西暦有無・最寄日自動補完）.
    console.log("テスト: parseDateParameter の検証");
    
    // 西暦あり.
    const date1 = parseDateParameter("2025/12/20", new Date(2026, 4, 25)); // 基準日: 2026/05/25
    assertEquals(date1.getFullYear(), 2025, "西暦ありが正しくパースされていません");
    assertEquals(date1.getMonth(), 11, "月が正しくパースされていません");
    assertEquals(date1.getDate(), 20, "日が正しくパースされていません");

    // 西暦なし: 過去（同月中）.
    const date2 = parseDateParameter("5/20", new Date(2026, 4, 25)); // 基準日: 2026/05/25
    assertEquals(date2.getFullYear(), 2026, "近い過去の日付補完に失敗しました");
    assertEquals(date2.getMonth(), 4, "月が異なります");
    assertEquals(date2.getDate(), 20, "日が異なります");

    // 西暦なし: 未来（同月内・来月）.
    const date3 = parseDateParameter("6/1", new Date(2026, 4, 25)); // 基準日: 2026/05/25
    assertEquals(date3.getFullYear(), 2026, "近い未来の日付補完に失敗しました");
    assertEquals(date3.getMonth(), 5, "月が異なります");
    assertEquals(date3.getDate(), 1, "日が異なります");

    // 西暦なし: 年またぎ（基準日 2025/01/02 で 12/20）.
    const date4 = parseDateParameter("12/20", new Date(2025, 0, 2)); // 基準日: 2025/01/02
    assertEquals(date4.getFullYear(), 2024, "年またぎの過去日付補完に失敗しました");
    assertEquals(date4.getMonth(), 11, "月が異なります");
    assertEquals(date4.getDate(), 20, "日が異なります");

    // 西暦なし: 年またぎ（基準日 2024/12/28 で 1/2）.
    const date5 = parseDateParameter("1/2", new Date(2024, 11, 28)); // 基準日: 2024/12/28
    assertEquals(date5.getFullYear(), 2025, "年またぎの未来日付補完に失敗しました");
    assertEquals(date5.getMonth(), 0, "月が異なります");
    assertEquals(date5.getDate(), 2, "日が異なります");

    // 空・無効値.
    const refDate = new Date(2026, 4, 25);
    const date6 = parseDateParameter("", refDate);
    assertEquals(date6.getTime(), refDate.getTime(), "空指定時のフォールバックが正しくありません");
    const date7 = parseDateParameter("invalid-date", refDate);
    assertEquals(date7.getTime(), refDate.getTime(), "無効指定時のフォールバックが正しくありません");

    // 4.5. normalizeSchoolYearのテスト.
    console.log("テスト: normalizeSchoolYear の検証");
    assertEquals(normalizeSchoolYear("小1"), "小1", "半角のままであるべき");
    assertEquals(normalizeSchoolYear("小１"), "小1", "全角数字が半角数字に変換されるべき");
    assertEquals(normalizeSchoolYear("小一"), "小1", "漢数字が一桁の半角数字に変換されるべき");
    assertEquals(normalizeSchoolYear(" 小 一 "), "小1", "空白が除去されるべき");
    assertEquals(normalizeSchoolYear("中三"), "中3", "漢数字の三が3に変換されるべき");
    assertEquals(normalizeSchoolYear("高10"), "高10", "10は10のままであるべき");
    assertEquals(normalizeSchoolYear("大十"), "大10", "漢数字の十が10に変換されるべき");
    assertEquals(normalizeSchoolYear("一般"), "一般", "数字を含まない文字列はそのまま残るべき");
    assertEquals(normalizeSchoolYear(null), "", "nullの場合は空文字が返るべき");

    // 5. 過去の昇級日のロールバック防止テスト.
    console.log("テスト: 過去日付での昇級処理（名簿更新スキップ）の検証");
    
    // パターンA: 新しい日付での昇級（更新されるべき）.
    mockSetValueCalls = [];
    mockAppendRowCalls = [];
    const resultNew = processPromotion("競技", "太郎", "B", new Date(2026, 4, 25), "最新の昇級"); // 2026/05/25
    assertEquals(resultNew, true, "最新日付での昇級処理に失敗しました");
    assertEquals(mockSetValueCalls.length, 2, "名簿シートが更新されていません");
    assertEquals(mockAppendRowCalls.length, 1, "履歴シートに追記されていません");

    // パターンB: 古い日付での昇級（名簿更新はスキップされ, 履歴追記のみされるべき）.
    mockSetValueCalls = [];
    mockAppendRowCalls = [];
    const resultOld = processPromotion("競技", "太郎", "D", new Date(2026, 4, 10), "過去の昇級"); // 2026/05/10 (登録済みの 2026/05/20 より過去)
    assertEquals(resultOld, true, "過去日付での昇級処理に失敗しました");
    assertEquals(mockSetValueCalls.length, 0, "過去日付なのに名簿シートが更新されてしまいました");
    assertEquals(mockAppendRowCalls.length, 1, "過去日付の履歴が追記されていません");

    // 6. dispatchLineCommandのテスト（LINE返信連携および日付指定）.
    console.log("テスト: dispatchLineCommand の検証");
    
    // 正常系: 昇級（日付指定あり）.
    mockFetchedUrls = [];
    dispatchLineCommand("★昇級★\n姓：競技\n名：太郎\n新級：B\n日付：5/20", "test_reply_token");
    assertEquals(mockFetchedUrls.length, 1, "LINE返信が送信されていません");
    const payload1 = JSON.parse(mockFetchedUrls[0].params.payload);
    assertEquals(payload1.replyToken, "test_reply_token", "replyTokenが正しく渡されていません");
    assertEquals(payload1.messages[0].text.includes("完了しました"), true, "完了メッセージが含まれていません");

    // 正常系: フルネーム指定（スペースあり）.
    mockFetchedUrls = [];
    dispatchLineCommand("★昇級★\nフルネーム：競技 太郎\n新級：B\n日付：5/20", "test_reply_token");
    assertEquals(mockFetchedUrls.length, 1, "LINE返信が送信されていません");
    const payloadFull1 = JSON.parse(mockFetchedUrls[0].params.payload);
    assertEquals(payloadFull1.messages[0].text.includes("競技 太郎 さんの昇級処理（B）を完了しました."), true, "フルネーム（スペースあり）の処理に失敗しました");

    // 正常系: フルネーム指定（スペースなし）.
    mockFetchedUrls = [];
    dispatchLineCommand("★昇級★\nフルネーム：競技太郎\n新級：B\n日付：5/20", "test_reply_token");
    assertEquals(mockFetchedUrls.length, 1, "LINE返信が送信されていません");
    const payloadFull2 = JSON.parse(mockFetchedUrls[0].params.payload);
    assertEquals(payloadFull2.messages[0].text.includes("競技 太郎 さんの昇級処理（B）を完了しました."), true, "フルネーム（スペースなし）の処理に失敗しました");

    // 異常系: フルネーム指定（存在しないメンバー）.
    mockFetchedUrls = [];
    dispatchLineCommand("★昇級★\nフルネーム：存在しないメンバー\n新級：B", "test_reply_token");
    assertEquals(mockFetchedUrls.length, 1, "エラー返信が送信されていません");
    const payloadFullErr = JSON.parse(mockFetchedUrls[0].params.payload);
    assertEquals(payloadFullErr.messages[0].text.includes("エラー：指定されたメンバー「存在しないメンバー」が見つかりません."), true, "存在しないメンバーのエラーメッセージが異なります");

    // 異常系: 入会コマンドでフルネーム指定（個別姓名がないためエラーになるべき）.
    mockFetchedUrls = [];
    dispatchLineCommand("★入会★\nフルネーム：新入生 太郎\n姓ふりがな：しんにゅうせい\n名ふりがな：たろう\n学年：小1\n級：入門\n区分：正会員", "test_reply_token");
    assertEquals(mockFetchedUrls.length, 1, "エラー返信が送信されていません");
    const payloadJoinErr = JSON.parse(mockFetchedUrls[0].params.payload);
    assertEquals(payloadJoinErr.messages[0].text.includes("形式が正しくありません.\n「姓」と「名」の両方を入力してください."), true, "入会コマンドでの姓名個別入力チェックに失敗しました");

    // 正常系: 区分変更（日付指定あり）.
    mockFetchedUrls = [];
    dispatchLineCommand("★区分変更★\n姓：競技\n名：太郎\n新区分：准会員\n日付：2025/12/20", "test_reply_token");
    assertEquals(mockFetchedUrls.length, 1, "LINE返信が送信されていません");
    const payload2 = JSON.parse(mockFetchedUrls[0].params.payload);
    assertEquals(payload2.messages[0].text.includes("准会員"), true, "更新メッセージが含まれていません");

    // 正常系: 説明コマンド.
    mockFetchedUrls = [];
    dispatchLineCommand("★説明★", "test_reply_token");
    assertEquals(mockFetchedUrls.length, 1, "LINE返信が送信されていません");
    const payloadDescription = JSON.parse(mockFetchedUrls[0].params.payload);
    assertEquals(payloadDescription.replyToken, "test_reply_token", "replyTokenが正しく渡されていません");
    assertEquals(payloadDescription.messages[0].text.includes("使い方説明"), true, "説明メッセージが含まれていません");

    // 異常系: 入力不備.
    mockFetchedUrls = [];
    dispatchLineCommand("★昇級★\n姓：競技", "test_reply_token");
    assertEquals(mockFetchedUrls.length, 1, "入力不備エラーが送信されていません");
    const payload4 = JSON.parse(mockFetchedUrls[0].params.payload);
    assertEquals(payload4.messages[0].text.includes("形式が正しくありません"), true, "形式エラーメッセージが異なります");

    console.log("すべてのテストに合格しました.");
  } catch (error) {
    console.error("テスト実行中にエラーが発生しました: " + error.message);
    throw error;
  } finally {
    teardownMocks();
    console.log("=== テスト完了 ===");
  }
}