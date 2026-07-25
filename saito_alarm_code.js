var ACCESS_TOKEN = 'Fl9/mn/oY19WSsSnWBQAjzsBvSVdvlZhqglkLm5q7TYWhVLqCAFSPq2SnGUtHYQyGeOo23HgrjNOoBZ1nHpiOI6LbXAe8eUSLJP5CzLHC/1cyf8oB8WaJxTQbv/ibV4sqBjguobK+uTVlmbBf6yqfgdB04t89/1O/w1cDnyilFU=';

var SHEET_KEY = '1cuQuT0LzwbHEzL3VLP_I0WH5NAiS0BNuzbjCdReYhzU';

function authorize() {
  SpreadsheetApp.openById(SHEET_KEY).getSheets();
  UrlFetchApp.fetch('https://api.line.me');
}

function doPost(e) {
  var event = JSON.parse(e.postData.contents).events[0];
  var userMessage = event.message.text;

  var message = "";
  
  var sheetSettings = SpreadsheetApp.openById(SHEET_KEY).getSheetByName('Settings');
  
  if (userMessage === '手続きしたよ') {
    sheetSettings.getRange("B1").setValue(true);
    message = "通知停止";
  } else if (userMessage === 'キャンセル') {
    message = "通知開始";
  } else {
    /*  if ( userMessage === "ID" ) {
      message = tellID(event);
    }
    else {
      message = "メッセージを受け取りました";
    }*/
  }

  // メッセージが空でない場合のみ返信を実行するよう調整
  if (message !== "") {
    replyMessage(event.replyToken, message);
  }
  return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService
    .createTextOutput('ok')
    .setMimeType(ContentService.MimeType.TEXT);
}

function replyMessage(token, message) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + ACCESS_TOKEN,
    },
    'method': 'post',
    'muteHttpExceptions': true,
    'payload': JSON.stringify({
      'replyToken': token,
      'messages': [{
        'type': 'text',
        'text': message,
      }],
    }),
  });
}

function pushMessage(to, message) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + ACCESS_TOKEN,
    },
    'method': 'post',
    'muteHttpExceptions': true,
    'payload': JSON.stringify({
      'to': to,
      'messages': [{
        'type': 'text',
        'text': message,
      }],
    }),
  });
}

function broadcastMessage(message) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + ACCESS_TOKEN,
    },
    'method': 'post',
    'muteHttpExceptions': true,
    'payload': JSON.stringify({
      'messages': [{
        'type': 'text',
        'text': message,
      }],
    }),
  });
}

function tellID(event) {
  // ID
  var userID = event.source.userId;
  var talkID = "";
  if (event.source.type === "group") {
    talkID = event.source.groupId;
  } else if (event.source.type === "room") {
    talkID = event.source.roomId;
  }

  var message = "あなたのID: " + userID;
  if (talkID != "") {
    message += "\nこのチャットのID: " + talkID;
  }

  return message;
}

function notice() {
  var settingsSheet = SpreadsheetApp.openById(SHEET_KEY).getSheetByName('Settings');
  //Trueの時配信を停止
  if (settingsSheet && settingsSheet.getRange("B1").getValue() === true) {
    return;
  }

  var sheet = SpreadsheetApp.openById(SHEET_KEY).getSheetByName('saito_alarm');
  var data  = sheet.getDataRange().getValues();

  var now = new Date();
  for (var i=1; i<data.length; i++) {
    var [day, hour, minute, message, to] = data[i];

    if (message === "") { continue; }

    if ( (day    ==  now.getDate()                   || day === "")
      && (hour   ==  now.getHours()                  || hour       === "")
      && (minute ==  now.getMinutes()                || minute     === "")
      )
      {
        broadcastMessage(message);
        if (to && to != ""){
          pushMessage(to, message);
        }
      }
  }
}

function autoResumeNotification() {
  var sheet = SpreadsheetApp.openById(SHEET_KEY).getSheetByName('Settings');
  if (sheet) {
    sheet.getRange("B1").setValue(false);
  }
}