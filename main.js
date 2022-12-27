/**
 * WebHookURLの一覧が記入されたスプレッドシートのURL（例：https://docs.google.com/spreadsheets/d/abc1234567/edit）
 * @type {string}
 */
var spreadSheetURL = "";

/**
 * Googleフォームを提出したときに実行する関数
 * @param {object} formData Googleフォームの情報
 * @returns {void} 戻り値なし
 */
function onFormSubmit(formData) {
  let itemResponses; // ここに全問が集約（問題文なども含まれる）https://developers.google.com/apps-script/reference/forms/form-response
  let responseslist = []; // 回答内容の必要部分だけ格納 0番目が1個目の問の回答 https://developers.google.com/apps-script/reference/forms/item-responseのリスト
  let messtr = ""; // Slackに送信する文字列

  let date = new Date(); // 現在時刻取得
  date = Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss"); // 時刻整形

  try {
    itemResponses = formData.response.getItemResponses();
  }
  catch (e) {
    // 例外エラー処理，稀に回答情報が取得できないことがあるようです．
    console.error("フォームの回答が正常に取得できませんでした．実行ログを確認してください．推定実行時刻: " + date);
    Logger.log(e);
    notifyToSlack("【エラー】フォームの回答が正常に取得できませんでした．実行ログを確認してください．推定実行時刻: " + date, getWebHookURLfromClassName("エラーログ"), true);
    return;
  }

  for (let i = 0; i < itemResponses.length; i++) {
    // Logger.log(itemResponses[i].getResponse());
    responseslist.push(itemResponses[i].getResponse()); // 回答をリストに連結
  }

  console.log(responseslist); // 回答情報デバッグ表示

  // 振替を選択したとき ＊選択肢の文字列と比較するので選択肢と一致させること
  if (responseslist[1] === "振替") {
    messtr = responseslist[0] + "さんが" + responseslist[2] + "クラスから" +
      responseslist[3] + "クラスへの振替を希望しています．以下自由記述: " + responseslist[4];

    // 振替元クラスと振替先クラスが異なるとき（正常）
    if (responseslist[2] !== responseslist[3]) {
      notifyToSlack(messtr, getWebHookURLfromClassName(responseslist[2]), true);
      notifyToSlack(messtr, getWebHookURLfromClassName(responseslist[3]), true);
    }

    // 振替元クラスと振替先クラスが同じとき（エラー）
    if (responseslist[2] === responseslist[3]) {
      notifyToSlack(messtr, getWebHookURLfromClassName("エラーログ"), true);
      console.log("【警告】振替元クラスと振替先クラスが一致していることが検出されました．推定実行時刻: " + date);
      notifyToSlack("【警告】振替元クラスと振替先クラスが一致していることが検出されました．送信者: " + responseslist[0] + "．推定実行時刻: " + date, getWebHookURLfromClassName("エラーログ"), true);
    }
  }
  // 欠席を選択したとき
  else if (responseslist[1] === "欠席") {
    messtr = responseslist[0] + "さんが" + responseslist[2] + "クラスを欠席します．以下自由記述: " + responseslist[3];

    notifyToSlack(messtr, getWebHookURLfromClassName(responseslist[2]), true);
  }
  // ドア開けを選択したとき
  else if (responseslist[1] === "ドア開け") {
    messtr = responseslist[0] + "さんがドアを開けてほしいようです．以下自由記述: " + responseslist[2];

    notifyToSlack(messtr, getWebHookURLfromClassName("全体"), true);
  }
  // それ以外（未実装）の選択肢を選択したとき
  else {
    console.error("【エラー】この選択肢は現在実装されていません．選択肢名: " + responseslist[1]);
    notifyToSlack("【エラー】この選択肢は現在実装されていません．実行ログを確認してください．選択肢名: " + responseslist[1] + "．推定実行時刻: " + date, getWebHookURLfromClassName("エラーログ"), true);
  }
}

/**
 * クラス識別名からWebHookURLを取得する関数
 * @param {string} className クラス識別名
 * @returns {string} 正常時string
 * @returns {null} エラー時null
 */
function getWebHookURLfromClassName(className) {
  const sheet = getWebHookSheet(); // エラー時はnull

  if (sheet === null) {
    return null; // エラー
  }

  const classInfo = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues(); //チャネル識別名とURL
  // Logger.log(classInfo);

  for (let i = 0; i < classInfo.length; i++) {
    if (className === classInfo[i][0]) {
      return classInfo[i][1]; // 正常
    }
  }

  console.error("【エラー】入力されたチャネル識別名が見つかりませんでした．クラス名: " + className);
  notifyToSlack("【エラー】入力されたチャネル識別名が見つかりませんでした．実行ログを確認してください．クラス名: " + className, getWebHookURLfromClassName("エラーログ"), true);
  return null; // エラー
}

/**
 * "WebHookURL"という名前がついたシートを取得する関数（＊スプレッドシートではない）
 * @returns {Sheet} 正常時https://developers.google.com/apps-script/reference/spreadsheet/sheet
 * @returns {null} エラー時null
 */
function getWebHookSheet() {
  let ss; // WebHookURLが記されたシートを含むスプレッドシートhttps://developers.google.com/apps-script/reference/spreadsheet/spreadsheet
  let sheet; // WebHookURLが記されたシートhttps://developers.google.com/apps-script/reference/spreadsheet/sheet

  if (spreadSheetURL === "") {
    console.error("【エラー】スプレッドシートのURLが入力されていません");
    return null; // エラー
  }

  try {
    ss = SpreadsheetApp.openByUrl(spreadSheetURL); // スプレッドシートを開く
  } catch (e) {
    // 例外エラー処理
    console.error("【エラー】スプレッドシートを開く際にエラーが発生しました．エラー内容は以下のとおりです．");
    Logger.log(e);
    return null; // エラー
  }

  try {
    sheet = ss.getSheetByName("WebHookURL"); // シートを開く
  } catch (e) {
    // 例外エラー処理
    console.error("【エラー】シートを開く際にエラーが発生しました．エラー内容は以下のとおりです．");
    Logger.log(e);
    return null; // エラー
  }

  return sheet; // 正常
}

/**
 * フォームの権限取得とWebHookの権限取得
 *     エラーログチャンネルにWebHookのテストメッセージを送信する関数
 * @returns {void} 戻り値なし
 */
function setup() {
  const form = FormApp.getActiveForm(); // フォームを開く（権限取得のためだけに）
  const sheet = getWebHookSheet(); // WebHookURLを含むシート

  if (sheet === null) {
    console.error("【エラー】WebHook処理を行えないため，処理を中止します．");
    return; // エラー
  }

  notifyToSlack("【テスト】送信テストを行います．これは初回動作チェックです．エラーログチャンネルにのみ通知されます．", getWebHookURLfromClassName("エラーログ"), true);
}

/**
 * 全クラスに対してWebHookのテストメッセージを送信する関数
 * @returns {void} 戻り値なし
 */
function classNameCheck() {
  const sheet = getWebHookSheet(); // エラー時はnullが返ってくる

  if (sheet === null) {
    console.error("【エラー】WebHook処理を行えないため，処理を中止します．");
    return; // エラー
  }

  const classInfo = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues(); //チャネル識別名とURL
  // Logger.log(classInfo);

  // 全チャンネルにテストメッセージを送信します．
  for (let i = 0; i < classInfo.length; i++) {
    notifyToSlack("【テスト】送信テストを行います．クラス名が一致しているか確認してください．クラス名: " + classInfo[i][0], classInfo[i][1], true);
  }
}

/**
 * SlackにWebHookを送信する関数
 * @param {string} messtr 送信する文字列
 * @param {string} slackWebHookURL 送信先のWebHookURL
 * @param {boolean} doMention メンション機能の使用
 * @returns {void} 戻り値なし
 */
function notifyToSlack(messtr, slackWebHookURL, doMention) {
  if (slackWebHookURL === "" | slackWebHookURL === null) {
    console.error("【エラー】WebHook URLが入力されていません");
    return; // エラー
  }

  // メンション機能
  if (doMention === undefined || doMention === true) {
    messtr = "<!channel> " + messtr;
  }

  // 投稿ユーザとメッセージ
  const jsonData =
  {
    "username": "振替/欠席/ドア開け連絡",
    "icon_emoji": ":exclamation:",
    "text": messtr
  };
  const payload = JSON.stringify(jsonData);

  const options =
  {
    "method": "post",
    "contentType": "application/json",
    "payload": payload
  };

  try {
    const res = UrlFetchApp.fetch(slackWebHookURL, options); // WebHook送信
    Logger.log(res);
  } catch (e) {
    // 例外エラー処理
    console.error("【エラー】WebHookでエラーが発生しました．エラー内容は以下のとおりです．");
    Logger.log(e);
  }
}
