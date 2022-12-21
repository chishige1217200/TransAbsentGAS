var spreadSheetURL = "https://docs.google.com/spreadsheets/d/1leXb6TdyBbYNxY2Zu3AMg9GU8Y393hrtoRETTg6Cl90/edit"
// WebHookURLの一覧が記入されたスプレッドシートのURL（例：https://docs.google.com/spreadsheets/d/abc1234567/edit）

function onFormSubmit(e) {
  var itemResponses = e.response.getItemResponses(); // ここに全問が集約（問題文なども含まれる）
  var responseslist = []; // 回答内容の必要部分だけ格納 0番目が1つめの問の回答
  var messtr = ""; // Slackに送信する文字列

  for (let i = 0; i < itemResponses.length; i++) {
    // Logger.log(itemResponses[i].getResponse());
    responseslist.push(itemResponses[i].getResponse()); // 回答をリストに連結
  }

  // 振替を選択したとき　＊選択肢の文字列と比較するので選択肢と一致させること
  if (responseslist[1] === "振替") {
    messtr = responseslist[0] + "さんが" + responseslist[2] + "クラスから" +
      responseslist[3] + "クラスへの振替を希望しています．以下自由記述: " + responseslist[4];

    // 振替元クラスと振替先先クラスが異なるとき（正常）
    if (responseslist[2] !== responseslist[3]) {
      notifyToSlack(messtr, getHookURLfromClassName(responseslist[2]));
      notifyToSlack(messtr, getHookURLfromClassName(responseslist[3]));
    }

    // 振替元クラスと振替先クラスが同じとき（エラー）
    if (responseslist[2] === responseslist[3]) {
      notifyToSlack(messtr, getHookURLfromClassName("エラーログ"));
      notifyToSlack("【警告】振替元クラスと振替先クラスが一致していることが検出されました．", getHookURLfromClassName("エラーログ"));
    }
  }
  // 欠席を選択したとき
  else if (responseslist[1] === "欠席") {
    messtr = responseslist[0] + "さんが" + responseslist[2] + "クラスを欠席します．以下自由記述: " + responseslist[3];

    notifyToSlack(messtr, getHookURLfromClassName(responseslist[2]));
  }
  // ドア開けを選択したとき
  else if (responseslist[1] === "ドア開け") {
    messtr = responseslist[0] + "さんがドアを開けてほしいようです．以下自由記述: " + responseslist[2];

    notifyToSlack(messtr, getHookURLfromClassName("全体"));
  }
}

function getHookURLfromClassName(className) {
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

  console.error("入力されたチャネル識別名が見つかりませんでした");
  notifyToSlack("【エラー】入力されたチャネル識別名が見つかりませんでした．実行ログを確認してください．クラス名: " + className, getHookURLfromClassName("エラーログ"));
  return null; // エラー
}

function getWebHookSheet() {
  let ss; // WebHookURLが記されたシートを含むスプレッドシート
  var sheet; // WebHookURLが記されたシート

  if (spreadSheetURL === "") {
    console.error("スプレッドシートのURLが入力されていません");
    return null; // エラー
  }

  try {
    ss = SpreadsheetApp.openByUrl(spreadSheetURL); // スプレッドシートを開く
  } catch (e) {
    // 例外エラー処理
    console.error("スプレッドシートを開く際にエラーが発生しました．エラー内容は以下のとおりです．");
    Logger.log(e);
    return null; // エラー
  }

  try {
    sheet = ss.getSheetByName("WebHookURL"); // シートを開く
  } catch (e) {
    // 例外エラー処理
    console.error("シートを開く際にエラーが発生しました．エラー内容は以下のとおりです．");
    Logger.log(e);
    return null; // エラー
  }

  return sheet; // 正常
}

function setup() { // FormAppの権限取得とWebhookのテスト
  const form = FormApp.getActiveForm(); // フォームを開く（権限取得のためだけに）
  const sheet = getWebHookSheet(); // エラー時はnullが返ってくる

  if (sheet === null) {
    Logger.log("WebHook処理を行えないため，処理を中止します．");
    return; // エラー
  }

  const classInfo = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues(); //チャネル識別名とURL
  // Logger.log(classInfo);

  // 全チャンネルにテストメッセージを送信します．迷惑なので初回セットアップ時にのみテストしてください．
  for (let i = 0; i < classInfo.length; i++) {
    notifyToSlack("【テスト】送信テストを行います．クラス名が一致しているか確認してください．クラス名: " + classInfo[i][0], classInfo[i][1]);
  }
}

function notifyToSlack(messtr, slackWebHookURL) {
  if (slackWebHookURL === "" | slackWebHookURL === null) {
    console.error("WebHook URLが入力されていません");
    return; // エラー
  }

  // 投稿ユーザとメッセージ
  const jsonData =
  {
    "username": "振替/欠席/ドア開け連絡",
    "icon_emoji": ":exclamation:",
    "text": "<!channel>" + messtr
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
    console.error("WebHookでエラーが発生しました．エラー内容は以下のとおりです．");
    Logger.log(e);
  }
}