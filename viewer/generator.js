// 級の出力順序定義.
const CLASS_ORDER = ["A", "B", "C", "D", "E", "F", "G", "基本", "入門"];

/**
 * 入会日データから yymm 形式（例: 2603）の文字列を生成する.
 * @param {any} joinDateVal - スプレッドシートの入会日データ
 * @return {string} - yyMMフォーマットの文字列. 変換できない場合は空文字
 */
function getFormattedJoinMonth(joinDateVal) {
  if (joinDateVal == null || joinDateVal === "") {
    return "";
  }

  let dateObj;
  if (joinDateVal instanceof Date) {
    dateObj = joinDateVal;
  } else {
    // 文字列などの場合は日付へのパースを試みる.
    dateObj = new Date(joinDateVal);
  }

  // 無効な日付の場合は空文字を返す.
  if (isNaN(dateObj.getTime())) {
    return "";
  }

  const fullYear = dateObj.getFullYear();
  const month = dateObj.getMonth() + 1;

  const yearString = String(fullYear).slice(-2);
  const monthString = String(month).padStart(2, "0");

  return yearString + monthString;
}

/**
 * 学年文字列をフォーマットする. 空文字または空白の場合は「一般」とする.
 * @param {string} schoolYearVal - 学年データ
 * @return {string} - フォーマットされた学年
 */
function formatSchoolYear(schoolYearVal) {
  if (schoolYearVal == null) {
    return "一般";
  }
  const trimmed = String(schoolYearVal).trim();
  if (trimmed === "") {
    return "一般";
  }
  return trimmed;
}

/**
 * 級別にグループ化し, ふりがな五十音順に並べたコピペ用名簿テキストを生成する.
 * @param {Array<string>} targetStatuses - 対象とする会員区分のリスト（例: ["正会員", "准会員"]）
 * @return {string} - 生成された名簿テキスト
 */
function generateClassListText(targetStatuses) {
  console.log("generateClassListTextを開始します. 対象区分: " + JSON.stringify(targetStatuses));
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(VIEWER_SHEET_NAME.LIST);

  if (sheet == null) {
    console.warn("名簿一覧シートが存在しないため, 空文字を返します.");
    return "";
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    console.warn("名簿一覧シートにデータがありません.");
    return "";
  }

  const colMap = getColumnMap(sheet, 1);
  const lastNameIdx = colMap[MEMBER_HEADER.LAST_NAME] - 1;
  const firstNameIdx = colMap[MEMBER_HEADER.FIRST_NAME] - 1;
  const lastNameKanaIdx = colMap[MEMBER_HEADER.LAST_NAME_FURIGANA] - 1;
  const firstNameKanaIdx = colMap[MEMBER_HEADER.FIRST_NAME_FURIGANA] - 1;
  const schoolYearIdx = colMap[MEMBER_HEADER.SCHOOL_YEAR] - 1;
  const classIdx = colMap[MEMBER_HEADER.CURRENT_CLASS] - 1;
  const statusIdx = colMap[MEMBER_HEADER.STATUS] - 1;
  const joinDateIdx = colMap[MEMBER_HEADER.JOIN_DATE] - 1;

  // メイト会員のみが対象かどうか判定する.
  const isMateOnly = targetStatuses.indexOf(MEMBER_STATUS.MATE) !== -1 && targetStatuses.indexOf(MEMBER_STATUS.REGULAR) === -1;

  const members = [];

  // 1. 対象のメンバーを抽出する.
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const status = row[statusIdx];

    // 対象の会員区分に含まれているか確認する.
    const isTargetStatus = targetStatuses.indexOf(status) !== -1;
    if (!isTargetStatus) {
      continue;
    }

    const currentClass = String(row[classIdx]).trim();
    const memberObj = {
      lastName: String(row[lastNameIdx]).trim(),
      firstName: String(row[firstNameIdx]).trim(),
      lastNameKana: String(row[lastNameKanaIdx]).trim(),
      firstNameKana: String(row[firstNameKanaIdx]).trim(),
      schoolYear: formatSchoolYear(row[schoolYearIdx]),
      currentClass: currentClass,
      status: status,
      joinDateVal: row[joinDateIdx],
    };
    members.push(memberObj);
  }

  // 2. ふりがなの五十音順（姓ふりがな + 名ふりがな）でソートする.
  members.sort((memberA, memberB) => {
    const fullNameKanaA = memberA.lastNameKana + memberA.firstNameKana;
    const fullNameKanaB = memberB.lastNameKana + memberB.firstNameKana;
    return fullNameKanaA.localeCompare(fullNameKanaB, "ja");
  });

  const lines = [];

  // 3. メイト会員の場合は, 級・学年を表示せず名前とふりがなのみでフラットに出力する.
  if (isMateOnly) {
    members.forEach((member) => {
      lines.push((member.lastName + member.firstName) + "\t" + (member.lastNameKana + member.firstNameKana));
    });
    return lines.join("\n");
  }

  // 4. 正准会員の場合は級ごとにグループ化して出力する.
  const classGroups = {};
  CLASS_ORDER.forEach((cls) => {
    classGroups[cls] = [];
  });
  // 定義外の級を受け止めるためのその他グループ.
  const otherClassGroup = [];

  members.forEach((member) => {
    if (classGroups[member.currentClass] != null) {
      classGroups[member.currentClass].push(member);
    } else {
      otherClassGroup.push(member);
    }
  });

  const addGroupText = (className, groupMembers) => {
    if (groupMembers.length === 0) {
      return;
    }

    // 既にテキストがある場合はグループ間に空行を入れる.
    if (lines.length > 0) {
      lines.push("");
    }

    lines.push(className + "級");
    groupMembers.forEach((member) => {
      let displaySchoolYear = member.schoolYear;
      if (member.status === MEMBER_STATUS.ASSOCIATE) {
        displaySchoolYear = displaySchoolYear + "(准)";
      }
      let suffix = "";
      const isBasicOrIntro = member.currentClass === "基本" || member.currentClass === "入門";
      if (isBasicOrIntro) {
        const yymm = getFormattedJoinMonth(member.joinDateVal);
        if (yymm !== "") {
          suffix = "[" + yymm + "]";
        }
      }
      lines.push(
        (member.lastName + member.firstName) +
        "\t" +
        (member.lastNameKana + member.firstNameKana) +
        "\t" +
        (displaySchoolYear + suffix)
      );
    });
  };

  // 定義順に出力する.
  CLASS_ORDER.forEach((cls) => {
    addGroupText(cls, classGroups[cls]);
  });

  // 定義外の級があれば最後に出力する.
  addGroupText("その他", otherClassGroup);

  return lines.join("\n");
}

/**
 * ソート名簿コピペシートの現在の並び順に基づいてコピペ用名簿テキストを生成する.
 * @param {boolean} [includeRegular=true] - 正会員・准会員を含めるかどうか
 * @param {boolean} [includeMate=true] - メイト会員を含めるかどうか
 * @param {boolean} [showClass=false] - 級を表示するかどうか
 * @return {string} - 生成された名簿テキスト
 */
function generateSortedListText(includeRegular = true, includeMate = true, showClass = false) {
  console.log("generateSortedListTextを開始します. 正准含: " + includeRegular + ", メイト含: " + includeMate + ", 級表示: " + showClass);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(VIEWER_SHEET_NAME.SORT_AND_COPY);

  if (sheet == null) {
    console.warn("ソート名簿コピペシートが存在しないため, 空文字を返します.");
    return "";
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    console.warn("ソート名簿コピペシートにデータがありません.");
    return "";
  }

  const colMap = getColumnMap(sheet, 1);
  const lastNameIdx = colMap[MEMBER_HEADER.LAST_NAME] - 1;
  const firstNameIdx = colMap[MEMBER_HEADER.FIRST_NAME] - 1;
  const lastNameKanaIdx = colMap[MEMBER_HEADER.LAST_NAME_FURIGANA] - 1;
  const firstNameKanaIdx = colMap[MEMBER_HEADER.FIRST_NAME_FURIGANA] - 1;
  const schoolYearIdx = colMap[MEMBER_HEADER.SCHOOL_YEAR] - 1;
  const classIdx = colMap[MEMBER_HEADER.CURRENT_CLASS] - 1;
  const statusIdx = colMap[MEMBER_HEADER.STATUS] - 1;
  const joinDateIdx = colMap[MEMBER_HEADER.JOIN_DATE] - 1;

  const lines = [];

  // スプレッドシート上でソートされた順序そのままで処理する.
  // 2行目（インデックス1）からループを開始する.
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const lastName = String(row[lastNameIdx]).trim();
    const firstName = String(row[firstNameIdx]).trim();
    // 姓名がない空行はスキップする.
    if (lastName === "" && firstName === "") {
      continue;
    }

    const status = row[statusIdx];
    const isRegularOrAssociate = status === MEMBER_STATUS.REGULAR || status === MEMBER_STATUS.ASSOCIATE;
    const isMate = status === MEMBER_STATUS.MATE;

    // フィルタリング処理.
    if (!includeRegular && isRegularOrAssociate) {
      continue;
    }
    if (!includeMate && isMate) {
      continue;
    }
    if (!isRegularOrAssociate && !isMate) {
      continue;
    }

    const lastNameKana = String(row[lastNameKanaIdx]).trim();
    const firstNameKana = String(row[firstNameKanaIdx]).trim();

    // メイト会員の場合は姓名とふりがなのみを出力する.
    if (status === MEMBER_STATUS.MATE) {
      lines.push((lastName + firstName) + "\t" + (lastNameKana + firstNameKana));
      continue;
    }

    const schoolYear = formatSchoolYear(row[schoolYearIdx]);
    const currentClass = String(row[classIdx]).trim();
    const joinDateVal = row[joinDateIdx];

    let suffix = "";
    const isBasicOrIntro = currentClass === "基本" || currentClass === "入門";
    if (isBasicOrIntro) {
      const yymm = getFormattedJoinMonth(joinDateVal);
      if (yymm !== "") {
        suffix = "[" + yymm + "]";
      }
    }

    let displaySchoolYear = schoolYear;
    if (status === MEMBER_STATUS.ASSOCIATE) {
      displaySchoolYear = displaySchoolYear + "(准)";
    }

    let classPart = "";
    if (showClass) {
      classPart = "\t" + currentClass;
    }

    lines.push(
      (lastName + firstName) +
      "\t" +
      (lastNameKana + firstNameKana) +
      "\t" +
      (displaySchoolYear + suffix) +
      classPart
    );
  }

  return lines.join("\n");
}

/**
 * LINE更新用名簿テキストを生成する.
 * @return {string} - 生成された名簿テキスト
 */
function generateLineUpdateListText() {
  console.log("generateLineUpdateListTextを開始します.");
  const dateString = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");

  const regularText = generateClassListText([MEMBER_STATUS.REGULAR, MEMBER_STATUS.ASSOCIATE]);
  const mateText = generateClassListText([MEMBER_STATUS.MATE]);

  const lines = [
    "ちはやふる富士見 会員名簿",
    "更新日: " + dateString,
    "",
    "注意:",
    "関係者外への持ち出しや共有を禁止します.",
    "",
    regularText,
    "",
    "メイト会員:",
    mateText
  ];

  return lines.join("\n");
}

/**
 * 正会員・准会員のコピペ用名簿テキストを生成する.
 * @return {string} - 生成された名簿テキスト
 */
function generateRegularAssociateClassListText() {
  return generateClassListText([MEMBER_STATUS.REGULAR, MEMBER_STATUS.ASSOCIATE]);
}

/**
 * メイト会員のコピペ用名簿テキストを生成する.
 * @return {string} - 生成された名簿テキスト
 */
function generateMateClassListText() {
  return generateClassListText([MEMBER_STATUS.MATE]);
}
