/**
 * LINEからのメッセージに特定文字列が入っていた場合に処理を実行する.
 * 昇級処理・区分変更・休退会に対応する.
 */

/**
 * POSTリクエストを受信したときの処理.
 * @param {Object} e - HTTP POSTイベントオブジェクト
 * @return {HtmlOutput} - LINEサーバーへの応答
 */
function doPost(e) {
  try {
    console.log("doPost関数が呼び出されました. イベントオブジェクト: " + JSON.stringify(e));

    // 1. リクエストペイロードの有無をチェックする.
    if (e == null || e.postData == null || e.postData.contents == null) {
      logWarn("doPost警告: リクエストデータ (e, postData, または contents) が存在しません.");
      return ContentService.createTextOutput("OK");
    }

    // 2. JSONパースを実行する. 失敗した場合は不正リクエストとして無視する.
    let eventData;
    try {
      eventData = JSON.parse(e.postData.contents);
      console.log("JSONパースに成功しました. ペイロード: " + JSON.stringify(eventData));
    } catch (error) {
      logError("doPostエラー: JSONパースに失敗しました. 原因: " + error.message + ", データ: " + e.postData.contents);
      return ContentService.createTextOutput("OK");
    }

    // 3. イベントの検証を行う. テキストメッセージ以外は早期離脱する.
    const event = eventData.events?.[0];
    if (event == null) {
      logWarn("doPost警告: イベントデータ (events) が空です.");
      return ContentService.createTextOutput("OK");
    }

    console.log("受信イベントの概要: type=" + event.type + ", source=" + JSON.stringify(event.source));

    if (event.type !== "message" || event.message?.type !== "text") {
      console.log("doPost情報: テキストメッセージ以外のイベントのため早期離脱します. タイプ: " + event.type + (event.message ? ", メッセージタイプ: " + event.message.type : ""));
      return ContentService.createTextOutput("OK");
    }

    const text = event.message.text.trim();
    console.log("受信テキストメッセージ: " + text);

    // 4. メッセージが名簿更新形式かを判定する.
    if (!/^★(昇級|区分変更|休退会|入会|復帰|説明)★/.test(text)) {
      console.log("doPost情報: 名簿更新用のキーワード（★昇級★など）で始まっていないため処理を終了します.");
      return ContentService.createTextOutput("OK");
    }

    // 5. メインロジックを実行する.
    const replyToken = event.replyToken || "";
    logInfo("名簿更新コマンドを検出しました. コマンド処理を開始します. replyToken: " + replyToken + ", テキスト: " + text);
    try {
      dispatchLineCommand(text, replyToken);
    } catch (error) {
      logError("doPostエラー: LINEコマンド処理中に未キャッチの例外が発生しました: " + error.stack);
    }
  } catch (error) {
    console.error("doPostエラー: 予期しないシステム例外が発生しました: " + error.stack);
  } finally {
    flushLogs();
  }

  // LINEサーバーには200 OKを返す.
  return ContentService.createTextOutput("OK");
}

/**
 * LINEからのコマンドを解析し, 適切な処理へ振り分けて返信する.
 * @param {string} text - 受信したテキストメッセージ
 * @param {string} replyToken - 返信用トークン
 */
function dispatchLineCommand(text, replyToken) {
  console.log("dispatchLineCommandを開始します. テキスト: " + text + ", replyToken: " + replyToken);
  const { type, params } = parseText(text);
  console.log("コマンド解析結果: type=" + type + ", params=" + JSON.stringify(params));

  let lastName = params["姓"];
  let firstName = params["名"];

  // 既存メンバーを対象とするコマンドの場合、フルネーム解決を試みる.
  if (type === "昇級" || type === "区分変更" || type === "休退会") {
    let rawFullName = params["フルネーム"] || params["姓名"] || params["氏名"] || params["名前"];
    if (rawFullName == null && lastName != null && firstName != null) {
      rawFullName = lastName + firstName;
    }

    if (rawFullName != null) {
      const fullNameWithoutSpace = rawFullName.replace(/\s/g, "");
      const resolved = findMemberByName(fullNameWithoutSpace);
      if (resolved != null) {
        lastName = resolved.lastName;
        firstName = resolved.firstName;
      } else {
        logWarn("メンバー検索エラー: フルネーム「" + rawFullName + "」に一致するメンバーが見つかりません.");
        sendLineReply(replyToken, "エラー：指定されたメンバー「" + rawFullName + "」が見つかりません.");
        return;
      }
    }
  }

  const dateStr = params["日付"];
  const eventDate = parseDateParameter(dateStr);

  let replyMessage = "";

  // 共通の入力チェックを行う. 説明コマンド以外では姓名は必須とする.
  if (type !== "説明" && (lastName == null || firstName == null)) {
    logWarn("コマンド検証エラー: 姓名が指定されていません. 姓: " + lastName + ", 名: " + firstName);
    if (type === "昇級" || type === "区分変更" || type === "休退会") {
      sendLineReply(replyToken, "形式が正しくありません.\n「フルネーム」を入力してください.");
    } else {
      sendLineReply(replyToken, "形式が正しくありません.\n「姓」と「名」の両方を入力してください.");
    }
    return;
  }

  console.log("基本検証パス: 姓名=" + lastName + " " + firstName + ", 種別=" + type + ", イベント発生日=" + eventDate);

  // 処理のルーティングを行う.
  if (type === "説明") {
    console.log("説明ルートが選択されました. ヘルプメッセージを生成します.");
    replyMessage = HELP_MESSAGE;
  } else if (type === "昇級") {
    const nextLevel = params["新級"];
    console.log("昇級ルートが選択されました. 新級: " + nextLevel);
    if (nextLevel != null) {
      console.log("昇級処理（processPromotion）を呼び出します...");
      const isSuccess = processPromotion(lastName, firstName, nextLevel, eventDate, "LINEからの更新");
      console.log("昇級処理の実行結果: " + isSuccess);
      replyMessage = isSuccess
        ? `${lastName} ${firstName} さんの昇級処理（${nextLevel}）を完了しました.`
        : `エラー：${lastName} ${firstName} さんが見つからないか, 処理に失敗しました.`;
    } else {
      logWarn("昇級ルート警告: 「新級」が入力されていません.");
      replyMessage = "形式が正しくありません.\n\n★昇級★\nフルネーム：\n新級：";
    }
  } else if (type === "区分変更") {
    const newStatus = params["新区分"];
    console.log("区分変更ルートが選択されました. 新区分: " + newStatus);
    if (newStatus != null) {
      console.log("区分変更処理（updateMemberStatus）を呼び出します...");
      const isSuccess = updateMemberStatus(lastName, firstName, newStatus, eventDate, "LINEからの更新");
      console.log("区分変更処理の実行結果: " + isSuccess);
      replyMessage = isSuccess
        ? `${lastName} ${firstName} さんの区分を「${newStatus}」に更新しました.`
        : `エラー：${lastName} ${firstName} さんが見つからないか, 処理に失敗しました.`;
    } else {
      logWarn("区分変更ルート警告: 「新区分」が入力されていません.");
      replyMessage = "形式が正しくありません.\n\n★区分変更★\nフルネーム：\n新区分：";
    }
  } else if (type === "休退会") {
    console.log("休退会ルートが選択されました. 区分更新処理（updateMemberStatus）を呼び出します...");
    // 休退会は区別せず一律処理する.
    const isSuccess = updateMemberStatus(lastName, firstName, "休退会", eventDate, "LINEからの更新");
    console.log("休退会処理の実行結果: " + isSuccess);
    replyMessage = isSuccess
      ? `${lastName} ${firstName} さんの処理（休退会）を完了しました.`
      : `エラー：${lastName} ${firstName} さんが見つからないか, 処理に失敗しました.`;
  } else if (type === "入会" || type === "復帰") {
    console.log("入会/復帰ルートが選択されました. パラメータの検証を行います.");
    // LINE入力の表記ゆれを吸収する.
    const lastNameFuri = params["姓ふりがな"] || params["ふりがな姓"] || params["せい"] || "";
    const firstNameFuri = params["名ふりがな"] || params["ふりがな名"] || params["めい"] || "";
    const schoolYear = normalizeSchoolYear(params["学年"] || "");

    // 級の検証を行う.
    const rawClass = params["級"] || params["現在の級"];
    const karutaClass = validateKarutaClass(rawClass);

    // 区分の検証を行う.
    const rawStatus = params["区分"];
    const newStatus = validateMemberStatus(rawStatus);

    console.log("入会/復帰パラメータ検証結果: 姓ふりがな=" + lastNameFuri + ", 名ふりがな=" + firstNameFuri + ", 学年=" + schoolYear + ", 検証後級=" + karutaClass + ", 検証後区分=" + newStatus);

    // 必須項目の網羅チェックを行う.
    if (!lastNameFuri || !firstNameFuri || !schoolYear || !rawClass) {
      logWarn("入会/復帰ルート警告: 必須項目が不足しています.");
      replyMessage = `【入力不備】\n必要な項目が不足しています.\n\n★${type}★\n姓：\n名：\n姓ふりがな：\n名ふりがな：\n学年：\n級：\n区分：`;
    } else if (karutaClass == null) {
      logWarn("入会/復帰ルート警告: 許可されていない級が入力されました. 入力値: " + rawClass);
      // 許可されていない級が入力された場合のエラー応答を設定する.
      replyMessage = `【入力エラー】\n「${rawClass}」は登録できません.\n\n級は以下のいずれかを入力してください.\n[ A, B, C, D, E, F, G, 基本, 入門 ]`;
    } else {
      console.log("入会/復帰処理（processJoin）を呼び出します...");
      // 検証を通過したkarutaClassを使用する.
      const isSuccess = processJoin(
        lastName,
        firstName,
        lastNameFuri,
        firstNameFuri,
        schoolYear,
        karutaClass,
        newStatus,
        eventDate,
        `LINEからの${type}申請`
      );
      console.log("入会/復帰処理の実行結果: " + isSuccess);
      replyMessage = isSuccess
        ? `【${type}完了】\n${lastName} ${firstName} さんを「${newStatus}」として登録しました.`
        : `【エラー】\n${lastName} ${firstName} さんの登録処理に失敗しました.`;
    }
  }

  // 結果を返信する.
  if (replyMessage !== "") {
    logInfo("返信メッセージを送信しました. 内容: " + replyMessage);
    sendLineReply(replyToken, replyMessage);
  } else {
    logWarn("警告: 返信メッセージが空のため、送信されませんでした.");
  }
}

/**
 * LINEからのテキストメッセージを解析し, 種別とパラメータを返す.
 * @param {string} text - 解析対象 of テキスト
 * @return {Object} - 種別とパラメータのオブジェクト
 */
function parseText(text) {
  const lines = text.split(/\r\n|\n/);
  const type = lines[0].replace(/\s/g, "").replace(/★/g, "");
  const params = {};

  for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (line === "") continue;

    const parts = line.split(/[:：]/);
    if (parts.length >= 2) {
      const key = parts[0].trim();
      const value = parts.slice(1).join(":").trim();
      params[key] = value;
    }
  }

  return { type, params };
}

/**
 * 級の表記ゆれを補正し, 許可された級のみを返す.
 * @param {string} rawStr - 入力された文字列
 * @return {string|null} - 正しい級文字列, 許可されていない値の場合は null
 */
function validateKarutaClass(rawStr) {
  if (rawStr == null || rawStr === "") return null;

  // 1. 全角英字を半角にし, 大文字に揃え, 前後の空白を削除する.
  let normalized = rawStr
    .replace(/[Ａ-Ｚａ-ｚ]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .toUpperCase()
    .trim();

  // 2. 末尾に「級」と入力されていたら自動で削除する.
  normalized = normalized.replace(/級$/, "");

  // 3. 許可するリストとの照合を行う.
  const VALID_CLASSES = ["A", "B", "C", "D", "E", "F", "G", "基本", "入門"];

  if (VALID_CLASSES.includes(normalized)) {
    return normalized;
  }

  return null;
}

/**
 * 会員区分の表記ゆれを補正し, 許可された区分のみを返す.
 * @param {string} rawStr - 入力された文字列
 * @return {string|null} - 正しい会員区分, 無効な場合は null
 */
function validateMemberStatus(rawStr) {
  if (rawStr == null || rawStr === "") return null;

  // 空白を除去する.
  const normalized = rawStr.trim();

  // 許可リストを定義する.
  const VALID_STATUSES = [MEMBER_STATUS.REGULAR, MEMBER_STATUS.ASSOCIATE, MEMBER_STATUS.MATE];

  // 完全一致チェックを行う.
  if (VALID_STATUSES.includes(normalized)) {
    return normalized;
  }

  // 「会員」が抜けている場合の補完を行う.
  for (const status of VALID_STATUSES) {
    if (status.startsWith(normalized)) {
      return status;
    }
  }

  return null;
}

/**
 * LINEへ返信する.
 * @param {string} replyToken - 返信用トークン
 * @param {string} text - 送信するテキスト
 */
function sendLineReply(replyToken, text) {
  console.log("sendLineReplyが呼び出されました. replyToken: " + replyToken + ", 送信テキスト: " + text);

  if (!replyToken) {
    logError("sendLineReplyエラー: replyTokenが空です. 送信を中断します.");
    return;
  }

  let token = "";
  try {
    token = PropertiesService.getScriptProperties().getProperty("LINE_TOKEN");
    console.log("PropertiesServiceからLINE_TOKENを取得しました. 長さ: " + (token ? token.length : 0));
  } catch (error) {
    logWarn("sendLineReply警告: ScriptPropertiesからLINE_TOKENを取得できませんでした. 原因: " + error.message);
  }

  // PropertiesServiceから取得できない場合はconfig.jsの定数を使用する.
  if (token == null || token === "" || token === "YOUR_LINE_ACCESS_TOKEN") {
    console.log("PropertiesServiceに有効なトークンがないため, config.jsのLINE_ACCESS_TOKENを使用します.");
    token = LINE_ACCESS_TOKEN;
  }

  if (token == null || token === "" || token === "YOUR_LINE_ACCESS_TOKEN") {
    logError("sendLineReplyエラー: 有効なLINEアクセストークンが設定されていません. 送信を中断します.");
    return;
  }

  const url = "https://api.line.me/v2/bot/message/reply";
  const payload = {
    replyToken: replyToken,
    messages: [{ type: "text", text: text }]
  };

  console.log("UrlFetchAppによるLINE APIへのリクエストを送信します. URL: " + url);
  try {
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + token
      },
      payload: JSON.stringify(payload)
    });
    console.log("UrlFetchAppリクエストが正常に終了しました. レスポンスコード: " + response.getResponseCode() + ", 内容: " + response.getContentText());
  } catch (error) {
    logError("sendLineReplyエラー: UrlFetchApp.fetch実行中に例外が発生しました. 原因: " + error.message);
  }
}

/**
 * 日付文字列を解析し, 基準日に最も近い日付オブジェクトを返す.
 * @param {string} rawStr - 日付文字列 (例: "2025/1/2", "12/20", "5/5")
 * @param {Date} [referenceDate] - 基準となる日付（デフォルトは今日）
 * @return {Date} - パースされた日付オブジェクト. 無効または空の場合は基準日
 */
function parseDateParameter(rawStr, referenceDate = new Date()) {
  if (rawStr == null || rawStr === "") {
    return referenceDate;
  }

  const cleaned = rawStr.trim();

  // 1. yyyy/m/d または yyyy-m-d の形式をチェックする.
  const yyyyMmDdRegex = /^(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})$/;
  const yyyyMatch = cleaned.match(yyyyMmDdRegex);
  if (yyyyMatch != null) {
    const year = parseInt(yyyyMatch[1], 10);
    const month = parseInt(yyyyMatch[2], 10) - 1;
    const day = parseInt(yyyyMatch[3], 10);
    const parsedDate = new Date(year, month, day);
    if (!isNaN(parsedDate.getTime())) {
      return parsedDate;
    }
  }

  // 2. m/d または m-d の形式をチェックする（西暦なし）.
  const mmDdRegex = /^(\d{1,2})[/\-](\d{1,2})$/;
  const mmDdMatch = cleaned.match(mmDdRegex);
  if (mmDdMatch != null) {
    const month = parseInt(mmDdMatch[1], 10) - 1;
    const day = parseInt(mmDdMatch[2], 10);

    const referenceYear = referenceDate.getFullYear();

    // 前年, 今年, 来年の同月日候補を生成する.
    const candidatePrev = new Date(referenceYear - 1, month, day);
    const candidateCurr = new Date(referenceYear, month, day);
    const candidateNext = new Date(referenceYear + 1, month, day);

    const diffPrev = Math.abs(candidatePrev.getTime() - referenceDate.getTime());
    const diffCurr = Math.abs(candidateCurr.getTime() - referenceDate.getTime());
    const diffNext = Math.abs(candidateNext.getTime() - referenceDate.getTime());

    let minDiff = diffCurr;
    let bestCandidate = candidateCurr;

    if (!isNaN(diffPrev) && diffPrev < minDiff) {
      minDiff = diffPrev;
      bestCandidate = candidatePrev;
    }
    if (!isNaN(diffNext) && diffNext < minDiff) {
      minDiff = diffNext;
      bestCandidate = candidateNext;
    }

    if (!isNaN(bestCandidate.getTime())) {
      return bestCandidate;
    }
  }

  console.warn("日付のパースに失敗しました. 入力値: " + rawStr + ", 基準日を使用します.");
  return referenceDate;
}

/**
 * 学年の表記ゆれ（半角・全角数字, 漢数字）を半角数字に正規化する.
 * @param {string} rawStr - 入力された学年文字列
 * @return {string} - 正規化された学年文字列
 */
function normalizeSchoolYear(rawStr) {
  if (rawStr == null) {
    return "";
  }

  // 1. 前後の空白および途中の空白を除去する.
  let normalized = rawStr.replace(/\s/g, "");

  // 2. 全角数字を半角数字に変換する.
  normalized = normalized.replace(/[０-９]/g, (ch) => {
    return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
  });

  // 3. 漢数字を半角数字に変換する.
  const KANJI_NUM_MAP = {
    "一": "1",
    "二": "2",
    "三": "3",
    "四": "4",
    "五": "5",
    "六": "6",
    "七": "7",
    "八": "8",
    "九": "9",
    "十": "10"
  };

  // パターン1: 「小」「中」「高」「大」などの学年プレフィックスの後に漢数字が続く場合.
  normalized = normalized.replace(
    /(小|中|高|大|小学|中学|高校|大学)([一二三四五六七八九十])/g,
    (match, p1, p2) => p1 + KANJI_NUM_MAP[p2]
  );

  // パターン2: 文字列全体が漢数字のみの場合.
  if (/^[一二三四五六七八九十]$/.test(normalized)) {
    normalized = KANJI_NUM_MAP[normalized];
  }

  // パターン3: 漢数字の後に「年」「年生」が続く場合.
  normalized = normalized.replace(
    /([一二三四五六七八九十])(年|年生)/g,
    (match, p1, p2) => KANJI_NUM_MAP[p1] + p2
  );

  return normalized;
}