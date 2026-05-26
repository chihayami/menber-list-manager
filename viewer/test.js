/**
 * テスト実行用グローバル変数およびモックデータ.
 */
let mockSheetData = [];

// 元のオブジェクトを待避する変数.
let originalSpreadsheetApp;
let originalUtilities;

/**
 * テスト用のモック環境をセットアップする.
 */
function setupMocks() {
  originalSpreadsheetApp = typeof SpreadsheetApp !== "undefined" ? SpreadsheetApp : null;
  originalUtilities = typeof Utilities !== "undefined" ? Utilities : null;

  Utilities = {
    formatDate: (date, tz, format) => {
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, '0');
      const d = String(date.getDate()).padStart(2, '0');
      return `${y}/${m}/${d}`;
    }
  };

  // モック用のダミーデータ.
  // 1行目がヘッダー, 2行目以降がデータ.
  mockSheetData = [
    ["姓", "名", "姓ふりがな", "名ふりがな", "学年", "現在の級", "最新昇級日", "会員区分", "備考", "入会日", "学年ソート用"],
    ["競技", "太郎", "きょうぎ", "たろう", "大1", "C", "2026/05/20", "正会員", "", "2025/04/01", 13],
    ["山田", "花子", "やまだ", "はなこ", "小3", "基本", "2026/05/10", "正会員", "", "2026/03/15", 3],
    ["鈴木", "一郎", "すずき", "いちろう", "", "入門", "2026/05/01", "准会員", "", "2026/04/20", 99],
    ["佐藤", "メイト", "さとう", "めいと", "中2", "B", "2026/05/15", "メイト会員", "", "2025/10/01", 8],
    ["田中", "休会", "たなか", "きゅうかい", "一般", "A", "2025/04/01", "休退会", "", "2024/04/01", 17]
  ];

  const mockSheet = {
    getDataRange: () => ({
      getValues: () => mockSheetData
    }),
    getLastColumn: () => mockSheetData[0].length,
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => {
        // 指定された行のヘッダーデータを返す（getColumnMap用）.
        const startRow = row - 1;
        const startCol = col - 1;
        const result = [];
        for (let r = 0; r < numRows; r++) {
          const rowData = [];
          for (let c = 0; c < numCols; c++) {
            rowData.push(mockSheetData[startRow + r][startCol + c]);
          }
          result.push(rowData);
        }
        return result;
      }
    })
  };

  SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: (name) => {
        // 設定されたシート名ごとにモックを返す.
        return mockSheet;
      }
    })
  };
}

/**
 * セットアップしたモック環境を復元する.
 */
function teardownMocks() {
  if (originalSpreadsheetApp != null) {
    SpreadsheetApp = originalSpreadsheetApp;
  }
  if (originalUtilities != null) {
    Utilities = originalUtilities;
  }
}

/**
 * 簡易アサーション関数.
 * @param {any} actual - 実際値
 * @param {any} expected - 期待値
 * @param {string} message - エラーメッセージ
 */
function assertEquals(actual, expected, message) {
  if (actual !== expected) {
    throw new Error(message + " - 期待値: [" + expected + "], 実際の値: [" + actual + "]");
  }
}

/**
 * テスト実行用メイン関数.
 */
function runTests() {
  console.log("=== viewerテスト開始 ===");
  setupMocks();

  try {
    // 1. getFormattedJoinMonth のテスト.
    console.log("テスト: getFormattedJoinMonth の検証");
    
    // Dateオブジェクト.
    assertEquals(getFormattedJoinMonth(new Date(2026, 2, 15)), "2603", "Dateオブジェクトの変換に失敗しました");
    assertEquals(getFormattedJoinMonth(new Date(2025, 11, 1)), "2512", "Dateオブジェクトの12月変換に失敗しました");
    
    // 文字列.
    assertEquals(getFormattedJoinMonth("2026/04/20"), "2604", "文字列(スラッシュ)の変換に失敗しました");
    assertEquals(getFormattedJoinMonth("2025-10-05"), "2510", "文字列(ハイフン)の変換に失敗しました");
    
    // 空文字・不正値.
    assertEquals(getFormattedJoinMonth(""), "", "空文字時の処理に失敗しました");
    assertEquals(getFormattedJoinMonth(null), "", "null時の処理に失敗しました");
    assertEquals(getFormattedJoinMonth("invalid-date-string"), "", "無効な日付文字列の処理に失敗しました");


    // 2. formatSchoolYear のテスト.
    console.log("テスト: formatSchoolYear の検証");
    assertEquals(formatSchoolYear("小5"), "小5", "通常の学年が正しくフォーマットされません");
    assertEquals(formatSchoolYear(""), "一般", "空文字列が一般に変換されません");
    assertEquals(formatSchoolYear("   "), "一般", "空白のみの文字列が一般に変換されません");
    assertEquals(formatSchoolYear(null), "一般", "nullが一般に変換されません");


    // 3. generateClassListText (正会員・准会員) のテスト.
    console.log("テスト: generateClassListText (正会員・准会員) の検証");
    const regularText = generateClassListText(["正会員", "准会員"]);
    
    // 期待されるテキスト:
    // C級
    // 競技 太郎 きょうぎ たろう 大1
    // 
    // 基本級
    // 山田 花子 やまだ はなこ 小3 (2603)
    // 
    // 入門級
    // 鈴木 一郎 すずき いちろう 一般 (2604)
    //
    // ※休退会者(田中)やメイト会員(佐藤)は含まれないこと.
    // ※基本/入門の学年の後ろに入会月(yymm)が付与されていること.
    // ※学年なしの鈴木一郎が「一般」になっていること.
    
        const expectedRegular = 
      "C級\n" +
      "競技太郎\tきょうぎたろう\t大1\n" +
      "\n" +
      "基本級\n" +
      "山田花子\tやまだはなこ\t小3[2603]\n" +
      "\n" +
      "入門級\n" +
      "鈴木一郎\tすずきいちろう\t一般(准)[2604]";
      
    assertEquals(regularText, expectedRegular, "正会員・准会員の級別テキスト生成結果が異なります");


    // 4. generateClassListText (メイト会員) のテスト.
    console.log("テスト: generateClassListText (メイト会員) の検証");
    const mateText = generateClassListText(["メイト会員"]);
    
    // 期待されるテキスト:
    // B級
    // 佐藤 メイト さとう めいと 中2
        const expectedMate = "佐藤メイト\tさとうめいと";
    assertEquals(mateText, expectedMate, "メイト会員の級別テキスト生成結果が異なります");


    // 5. generateSortedListText のテスト.
    console.log("テスト: generateSortedListText の検証");
    
    // ソート名簿コピペシートの現在の並び順そのままで出力されること.
    // 1行目の操作パネル行、2行目のヘッダー行を除き、3行目以降が順に出力される.
    // 期待されるテキスト:
    // 競技 太郎 きょうぎ たろう 大1
    // 山田 花子 やまだ はなこ 小3 (2603)
    // 鈴木 一郎 すずき いちろう 一般 (2604)
    // 佐藤 メイト さとう めいと 中2
    // 田中 休会 たなか きゅうかい 一般
    
    const sortedText = generateSortedListText();
        const expectedSorted = 
      "競技太郎\tきょうぎたろう\t大1\n" +
      "山田花子\tやまだはなこ\t小3[2603]\n" +
      "鈴木一郎\tすずきいちろう\t一般(准)[2604]\n" +
      "佐藤メイト\tさとうめいと";
      
    assertEquals(sortedText, expectedSorted, "ソート名簿のテキスト生成結果が異なります");


    // 6. generateSortedListText (オプション付き) のテスト.
    console.log("テスト: generateSortedListText (オプション付き) の検証");
    // 正准=ON, メイト=OFF, 級表示=ON
    const sortedWithOptionsText = generateSortedListText(true, false, true);
    const expectedSortedWithOptions = 
      "競技太郎\tきょうぎたろう\t大1\tC\n" +
      "山田花子\tやまだはなこ\t小3[2603]\t基本\n" +
      "鈴木一郎\tすずきいちろう\t一般(准)[2604]\t入門";
    assertEquals(sortedWithOptionsText, expectedSortedWithOptions, "オプション付きソート名簿のテキスト生成結果が異なります");


    // 7. getSchoolYearSortKey のテスト.
    console.log("テスト: getSchoolYearSortKey の検証");
    assertEquals(getSchoolYearSortKey("小1"), 1, "小1のソートキーが正しくありません");
    assertEquals(getSchoolYearSortKey("一般"), 22, "一般のソートキーが正しくありません");
    assertEquals(getSchoolYearSortKey(""), 99, "空文字のソートキーが正しくありません");
    assertEquals(getSchoolYearSortKey(null), 99, "nullのソートキーが正しくありません");

    // 8. generateLineUpdateListText のテスト.
    console.log("テスト: generateLineUpdateListText の検証");
    const lineUpdateText = generateLineUpdateListText();
    const todayStr = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
    const expectedLineUpdate = 
      "ちはやふる富士見 会員名簿\n" +
      "更新日: " + todayStr + "\n" +
      "\n" +
      "注意:\n" +
      "関係者外への持ち出しや共有を禁止します.\n" +
      "\n" +
      "C級\n" +
      "競技太郎\tきょうぎたろう\t大1\n" +
      "\n" +
      "基本級\n" +
      "山田花子\tやまだはなこ\t小3[2603]\n" +
      "\n" +
      "入門級\n" +
      "鈴木一郎\tすずきいちろう\t一般(准)[2604]\n" +
      "\n" +
      "メイト会員:\n" +
      "佐藤メイト\tさとうめいと";
    
    assertEquals(lineUpdateText, expectedLineUpdate, "LINE更新用名簿のテキスト生成結果が異なります");

    console.log("すべてのテストに合格しました.");
  } catch (error) {
    console.error("テスト失敗: " + error.message);
    throw error;
  } finally {
    teardownMocks();
    console.log("=== viewerテスト完了 ===");
  }
}
