//メンバーのリストアップ
function getUserIdsFromChannel(token, channelId) {
  // Slack APIのconversations.membersメソッドを使用してチャンネルのメンバーリストを取得
  var apiUrl = "https://slack.com/api/conversations.members";
  //場合によって変える（チャンネルID固定の場合）
  //var channelId = "";
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("SHEETNAME");
  var channelId = sheet.getRange("F7").getValue(); 
  var token = "TOKEN";
  var params = {
    "token": token,
    "channel": channelId
  };
  var response = UrlFetchApp.fetch(apiUrl, {
    "method": "GET",
    "payload": params
  });
  var responseData = JSON.parse(response.getContentText());
  
  // チャンネルのメンバーIDのリストを取得
  var memberIds = responseData.members;
  return memberIds;
}

//チャンネルにいるメンバーをリストアップ
function outputMembersToSpreadsheet(token, channelId, sheetName) {
  var memberIds = getUserIdsFromChannel(token, channelId);
  
  // スプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("SHEETNAME");
  
  // メンバーIDのリストをスプレッドシートに出力
  if (sheet) {
    // ヘッダーを設定
    sheet.getRange("A1").setValue("メンバーID");
    
    // メンバーIDをスプレッドシートに出力
    for (var i = 0; i < memberIds.length; i++) {
      sheet.getRange("A" + (i + 2)).setValue(memberIds[i]);
    }
    
    Logger.log("メンバー情報をスプレッドシートに出力しました。");
    convertMemberIdsToNames();
  } else {
    Logger.log("指定されたシートが見つかりません。");
  }
}
//リストのメンバーIDを名前と連携させて見やすくする
function convertMemberIdsToNames() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName("SHEETNAME");
  var targetSheet = spreadsheet.getSheetByName("LISTSHEETNAME");

  // B列およびD列の範囲を取得
  var targetRangeB = targetSheet.getRange("B:B");
  var targetRangeE = targetSheet.getRange("E:E");
  
  // メンバーの名前が書いてある列の範囲
  var namesRange = targetSheet.getRange("C:C");
  var namesRangeF = targetSheet.getRange("F:F");

  var sourceValues = sourceSheet.getRange("A:A").getValues();

  // B列とD列からメンバーIDを検索して、該当する名前を取得して設定する
  for (var i = 0; i < sourceValues.length; i++) {
    var memberId = sourceValues[i][0];
    if (memberId) { // 空でないセルのみ処理を行う
      var name = "";
      // B列から検索
      var targetValuesB = targetRangeB.getValues();
      for (var j = 0; j < targetValuesB.length; j++) {
        if (targetValuesB[j][0] == memberId) {
          name = namesRange.getValues()[j][0];
          break;
        }
      }
      // D列から検索
      var targetValuesD = targetRangeE.getValues();
      if (!name) {
        for (var k = 0; k < targetValuesD.length; k++) {
          if (targetValuesD[k][0] == memberId) {
            name = namesRangeF.getValues()[k][0];
            break;
          }
        }
      }
      // 名前が取得できた場合は設定する
      if (name) {
        sourceSheet.getRange("B" + (i + 1)).setValue(name);
      }
    }
  }
}


//チャンネルメンバーの退出
function kickChannelSheet(token) {
  var token = "TOKEN";
    //場合によって変える
  //var channelId = "";
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetA = spreadsheet.getSheetByName("SHEETNAME");
  var channelId = sheetA.getRange("F7").getValue(); 
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("SHEETNAME");
  var range = sheet.getRange("A:C"); // A列からC列までの範囲を取得
  var values = range.getValues();
  
  var confirmed = Browser.msgBox(
    "メンバーをキックしますか？",
    "この操作は取り消せません。",
    Browser.Buttons.OK_CANCEL
  );
  
  if (confirmed == "cancel") {
    Logger.log("操作がキャンセルされました。");
    return;
  }
  
  for (var i = 0; i < values.length; i++) {
    var memberId = values[i][0];
    var isChecked = values[i][2]; // C列のチェック状態を取得
    if (!isChecked && memberId) { // チェックが入っていないかつメンバーIDがある場合のみ処理を行う
      kickMemberChannel(token, memberId,channelId);
    }
  }
}

function kickMemberChannel(token, memberId,channelId) {
  // Slack API の conversations.kick メソッドを使用してメンバーをキック
  var apiUrl = "https://slack.com/api/conversations.kick";
  var params = {
    "token": token, // usertoken を使用する
    "channel": channelId,
    "user": memberId
  };
  
  var options = {
    "method": "POST",
    "contentType": "application/json",
    "headers": {
      "Authorization": "Bearer " + token, // Bearer トークンを指定する
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(params)
  };

  var response = UrlFetchApp.fetch(apiUrl, options);
  var responseData = JSON.parse(response.getContentText());
  
  if (responseData.ok) {
    Logger.log("メンバー " + memberId + " をキックしました。");
  } else {
    Logger.log("メンバー " + memberId + " をキックできませんでした。エラー: " + responseData.error);
  }
}



//Botをチャンネルに参加させるやつ
function inviteBotToChannel() {
  var token = "TOKEN";
  var apiUrl = "https://slack.com/api/conversations.join";
  //場合によって付け替える
  //var channelId = "";
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("SHEETNAME");
  var channelId = sheet.getRange("F7").getValue(); // スプレッドシートからチャンネルIDを取得
  
  var params = {
    "token": token,
    "channel": channelId // チャンネルのID
  };
  
  var options = {
    "method": "POST",
    "contentType": "application/json",
    "headers": {
      "Authorization": "Bearer " + token, // Bearer トークンを指定する
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(params)
  };
  
  var response = UrlFetchApp.fetch(apiUrl, options);
  var responseData = JSON.parse(response.getContentText());
  
  if (responseData.ok) {
    Logger.log("ボットがチャンネルに参加しました。");
  } else {
    Logger.log("ボットがチャンネルに参加できませんでした。エラー: " + responseData.error);
  }
}

//A列2行目のリストを全て削除するやつ
function clearSheetData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("SHEETNAME"); // シート名を指定
  
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) { // 2行目以降にデータがある場合
    sheet.getRange("A2:B" + lastRow).clearContent(); // A列の2行目から最終行までのデータをクリア
    sheet.getRange("E7:F").clearContent();//
  } else {
    Logger.log("2行目以降にデータがありません。"); // 2行目以降にデータがない場合
  }
}

//チャンネルに招待する
function inviteMembersToSlackChannel() {
  // Slack Bot Token
  var slackToken = "TOKEN";
  
  // Spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("SHEETNAME");

  // Get Channel ID from Spreadsheet
  var channelId = sheet.getRange("F7").getValue();

  // Get Member IDs from Spreadsheet
  var memberIds = sheet.getRange("A2:A").getValues().flat().filter(String);
  
  // Invite Members to Slack Channel
  memberIds.forEach(function(memberId) {
    var apiUrl = "https://slack.com/api/conversations.invite";
    var payload = {
      channel: channelId,
      users: memberId
    };
    var options = {
      method: "post",
      headers: {
        Authorization: "Bearer " + slackToken
      },
      contentType: "application/json",
      payload: JSON.stringify(payload)
    };
    var response = UrlFetchApp.fetch(apiUrl, options);
    var jsonResponse = JSON.parse(response.getContentText());
    if (!jsonResponse.ok) {
      Logger.log("Error inviting member " + memberId + " to the channel: " + jsonResponse.error);
    } else {
      Logger.log("Member " + memberId + " invited to the channel successfully.");
    }
  });
}

//チェックをつける
function allCheckboxes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SHEETNAME");
  var range = sheet.getRange("C:C"); // C列の範囲を取得
  var values = range.getValues(); // C列の全ての値を取得
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === false) { // チェックボックスが未チェックの場合
      range.getCell(i + 1, 1).setValue(true); // チェックを入れる
    }
  }
}

//チェックを外す
function allResetboxes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SHEETNAME");
  var range = sheet.getRange("C:C"); // C列の範囲を取得
  var values = range.getValues(); // C列の全ての値を取得
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === true) { // チェックボックスがチェックされている場合
      range.getCell(i + 1, 1).setValue(false); // チェックを外す
    }
  }
}

