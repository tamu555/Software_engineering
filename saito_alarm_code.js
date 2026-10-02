// チャネルアクセストークンはスクリプトプロパティに保存
var ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_ACCESS_TOKEN');

var SHEET_KEY = '1cuQuT0LzwbHEzL3VLP_I0WH5NAiS0BNuzbjCdReYhzU';

function authorize() {
  SpreadsheetApp.openById(SHEET_KEY).getSheets();
  UrlFetchApp.fetch('https://api.line.me');
}

function doPost(e) {
  var events = JSON.parse(e.postData.contents).events || [];
  var sheetSettings = SpreadsheetApp.openById(SHEET_KEY).getSheetByName('Settings');

  events.forEach(function (event) {
    // テキストメッセージ以外（スタンプ・画像・友だち追加など）は無視
    if (event.type !== 'message' || event.message.type !== 'text') { return; }

    var userMessage = event.message.text;
    var message = "";

    if (userMessage === '手続きしたよ') {
      if (sheetSettings) { sheetSettings.getRange("B1").setValue(true); }
      message = "通知停止";
    } else if (userMessage === 'キャンセル') {
      if (sheetSettings) { sheetSettings.getRange("B1").setValue(false); }
      message = "通知開始";
    }

    if (message !== "") {
      replyMessage(event.replyToken, message);
    }
  });

  return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService
    .createTextOutput('ok')
    .setMimeType(ContentService.MimeType.TEXT);
}

function callLineApi(endpoint, payload) {
  var response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/' + endpoint, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + ACCESS_TOKEN,
    },
    'method': 'post',
    'muteHttpExceptions': true,
    'payload': JSON.stringify(payload),
  });

  var code = response.getResponseCode();
  if (code !== 200) {
    console.error('LINE API ' + endpoint + ' failed: ' + code + ' ' + response.getContentText());
  }
  return response;
}

function textMessages(message) {
  return [{
    'type': 'text',
    'text': message,
  }];
}

function replyMessage(token, message) {
  return callLineApi('reply', {
    'replyToken': token,
    'messages': textMessages(message),
  });
}

function pushMessage(to, message) {
  return callLineApi('push', {
    'to': to,
    'messages': textMessages(message),
  });
}

function broadcastMessage(message) {
  return callLineApi('broadcast', {
    'messages': textMessages(message),
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

// Settings!B1 が true なら配信停止中
function isStopped(settingsSheet) {
  return settingsSheet !== null && settingsSheet.getRange("B1").getValue() === true;
}

function notice() {
  var spreadsheet = SpreadsheetApp.openById(SHEET_KEY);
  var settingsSheet = spreadsheet.getSheetByName('Settings');

  var sheet = spreadsheet.getSheetByName('saito_alarm');
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
        if (isStopped(settingsSheet)) { return; }
        broadcastMessage(message);

        if (to) {
          if (isStopped(settingsSheet)) { return; }
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
