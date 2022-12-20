function onFormSubmit(e) {
  var itemResponses = e.response.getItemResponses(); // ここに全問が集約
  var responseslist = []; // 回答内容の必要部分だけ格納
  var messtr = ""; // Slackに送信する文字列

  for (let i = 0; i < itemResponses.length; i++) {
    // Logger.log(itemResponses[i].getResponse());
    responseslist.push(itemResponses[i].getResponse());
  }

  if (responseslist[2] === "振替") {
    messtr = responseslist[0] + "さんが" + responseslist[1] + "クラスから" +
      responseslist[3] + "クラスへの振替を希望しています．以下自由記述: " + responseslist[4];
  }
  else if (responseslist[2] === "欠席") {
    messtr = responseslist[0] + "さんが" + responseslist[1] + "クラスを欠席します．以下自由記述: " + responseslist[3];
  }

  notifyToSlack(messtr);
}

function setup() { // FormAppの権限取得とWebhookのテスト
  const form = FormApp.getActiveForm();
  notifyToSlack("これはテストメッセージです．");
}

function notifyToSlack(messtr) {
  // Slack側で作成したボットのウェブフックURL
  const slackWebHookURL = "";

  if (slackWebHookURL === "") {
    console.error("WebHook URLが入力されていません");
    return;
  }

  // 投稿ユーザとメッセージ
  const jsonData =
  {
    "username": "振替/欠席連絡",
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

  UrlFetchApp.fetch(slackWebHookURL, options);
}
