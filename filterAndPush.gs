function afterDataUpdate(){
  checkhangvasomatve()

}
function filterAndPush() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("LogemailVJ");
  var data = sheet.getDataRange().getValues(); // Lấy toàn bộ dữ liệu

  var listvj = [];
  var rowsToUpdate = []; // Lưu index dòng để tí cập nhật

  // Bỏ header (i = 1)
  for (var i = 1; i < data.length; i++) {
    var colE = data[i][4]; // Cột E
    var colI = data[i][8]; // Cột I

    if (colE && colI.toString().trim() !== "OK") {
      listvj.push(colE);
      rowsToUpdate.push(i + 1); // Dòng trên sheet (bắt đầu từ 1)
    }
  }
  Logger.log(Array.isArray(listvj)); // Phải ra true
  Logger.log(listvj);
  if (listvj.length > 0) {
    var text = listvj.join("\n");
    var result = pushTele("VJ",text);

    if (result === "ok") {
      rowsToUpdate.forEach(function(row) {
        sheet.getRange(row, 9).setValue("OK"); // Cột I = 9
      });
    }
  }
}
function filterAndPushVNA() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("LogemailVNA");
  var data = sheet.getDataRange().getValues(); // Lấy toàn bộ dữ liệu

  var list = [];
  var rowsToUpdate = []; // Lưu index dòng để tí cập nhật

  // Bỏ header (i = 1)
  for (var i = 1; i < data.length; i++) {
    var colE = data[i][4]; // Cột E
    var colI = data[i][8]; // Cột I

    if (colE && colI.toString().trim() !== "OK") {
      list.push(colE);
      rowsToUpdate.push(i + 1); // Dòng trên sheet (bắt đầu từ 1)
    }
  }

  if (list.length > 0) {
    var text = list.join("\n");
    var result = pushTele("VNA",text);

    if (result === "ok") {
      rowsToUpdate.forEach(function(row) {
        sheet.getRange(row, 9).setValue("OK"); // Cột I = 9
      });
    }
  }
}
// Fake hàm pushTele để test
function pushTele(hang, text) {
  var token = "7359295123:AAGz0rHge3L5gM-XJmyzNq6sayULdHO4-qE"; // Bot Token
  var chatId = "-1002520783135"; // ID nhóm hoặc user

  var url = "https://api.telegram.org/bot" + token + "/sendMessage";

  var payload = {
    chat_id: chatId,
    text: "Thông báo có chuyến bay hãng " + hang + " thay đổi : \n" + text
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var result = JSON.parse(response.getContentText());

  if (result.ok) {
    Logger.log("Push thành công dòng: " + hang);
    return "ok";
  } else {
    Logger.log("Push lỗi: " + JSON.stringify(result));
    return "fail";
  }
}

function checkhangvasomatve() {
  var sheetNameTarget = "Danh sách PNR cần gửi mail";
  var sheetNameLog = "Logemailxuấtvéthànhcông";
  var spreadsheetId = "11lIiwBcRyBZMvJxBjdryQDBf9w6wU0nRH8_z5cco78Y";
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var targetSheet = ss.getSheetByName(sheetNameTarget);
  var logSheet = ss.getSheetByName(sheetNameLog);

  // Lấy dữ liệu cột B và G bên sheet target
  var lastRowTarget = targetSheet.getLastRow();
  var dataTarget = targetSheet.getRange(1, 2, lastRowTarget, 6).getValues(); 
  // Cột B (index 0) và Cột G (index 5) trong mảng vì lấy từ cột B → G

  // Lấy dữ liệu cột E và F bên Log
  var lastRowLog = logSheet.getLastRow();
  var dataLog = logSheet.getRange(1, 5, lastRowLog, 2).getValues(); // cột E,F

  // Tạo map tra nhanh
  var logMap = {};
  for (var i = 0; i < dataLog.length; i++) {
    var key = dataLog[i][0];
    var value = dataLog[i][1];
    if (key) {
      logMap[key] = value;
    }
  }
  
  // Duyệt toàn bộ dòng bên target
  for (var row = 0; row < dataTarget.length; row++) {
    var pnr = dataTarget[row][0]; // cột B
    var colG = dataTarget[row][5]; // cột G

    if (pnr && !colG && logMap[pnr] !== undefined) {
      // TH: B có giá trị, G trống → điền
      targetSheet.getRange(row + 1, 7).setValue(logMap[pnr]);
      
      Logger.log(logMap[pnr])
      if (logMap[pnr]=="VJ"){
        targetSheet.getRange(row + 1, 8).setValue(1);
      }else {
        
        sove= checksomatve(pnr)
        if (sove>0){
          targetSheet.getRange(row + 1, 8).setValue(sove);

        }
      }
    } else if (!pnr && colG) {
      // TH: B trống, G có giá trị → xóa
      targetSheet.getRange(row + 1, 7).clearContent();
      targetSheet.getRange(row + 1, 8).clearContent();
    }
  }
}

function checksomatve(pnr) {
  var url = "https://thuhongtour.com/check-so-mat-ve-vna/?pnr=" +pnr + "&ssid=check";
  Logger.log(url)
  var options = {
    method: "post",
    headers: {
      "accept": "application/json"
    },
    payload: "" // tương đương -d ''
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response)
  var result = response.getContentText().trim(); // trả về string, ví dụ "12345"

  // Nếu cần số thì convert sang Number
  var numberResult = Number(result);
  Logger.log(numberResult)
  return isNaN(numberResult) ? result : numberResult;
}
