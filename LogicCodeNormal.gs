function afterDataUpdate(){
  
  checkhangvasomatve()
  SpreadsheetApp.flush();
  callGasBotCheck()
  processPNR_safe_v2()
  
}

function callGasBotCheck() {
  var url = "https://script.google.com/macros/s/AKfycbwtDWmUJuAPDSpvSwY2dWvHCt7rbwRoREvpWg8sgGjgGTndRCNrVjrHKtpjCEFfu18U/exec"
  var endpointGAS = url + "?todo=send";
  try {
    // Gửi xong vứt luôn, không đọc response
    UrlFetchApp.fetch(endpointGAS, { method: "get", muteHttpExceptions: true });
  } catch (e) {
    // Bỏ qua hết
  }
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
  var scriptProps = PropertiesService.getScriptProperties();

  // Lấy trạng thái cũ
  var isRunning = scriptProps.getProperty('checkhangvasomatve_isRunning');
  var skipCount = parseInt(scriptProps.getProperty('checkhangvasomatve_skipCount') || '0', 10);

  // Check chống chạy trùng
  if (isRunning === 'true') {
    skipCount++;
    Logger.log("⚠️ Hàm đang chạy, bỏ qua để tránh trùng. Số lần liên tiếp: " + skipCount);

    if (skipCount >= 3) {
      Logger.log("🔄 Reset cờ isRunning vì bị kẹt quá 3 lần liên tiếp.");
      scriptProps.setProperty('checkhangvasomatve_isRunning', 'false');
      skipCount = 0; // reset lại đếm
    }
    scriptProps.setProperty('checkhangvasomatve_skipCount', skipCount.toString());
    return;
  }

  // Nếu chạy được thì reset skipCount
  scriptProps.setProperty('checkhangvasomatve_skipCount', '0');
  scriptProps.setProperty('checkhangvasomatve_isRunning', 'true');

  try {
    var sheetNameTarget = "Danh sách PNR cần gửi mail";
    var sheetNameLog = "Logemailxuấtvéthànhcông";
    var spreadsheetId = "11lIiwBcRyBZMvJxBjdryQDBf9w6wU0nRH8_z5cco78Y";
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var targetSheet = ss.getSheetByName(sheetNameTarget);
    var logSheet = ss.getSheetByName(sheetNameLog);

    // Lấy dữ liệu cột B → H (7 cột)
    var lastRowTarget = targetSheet.getLastRow();
    if (lastRowTarget < 1) return;
    var dataTarget = targetSheet.getRange(1, 2, lastRowTarget, 7).getValues(); // B → H

    // Lấy dữ liệu cột E,F bên log
    var lastRowLog = logSheet.getLastRow();
    var dataLog = logSheet.getRange(1, 5, lastRowLog, 2).getValues();

    // Map nhanh từ logSheet
    var logMap = {};
    dataLog.forEach(function(row) {
      var key = row[0];
      if (key) logMap[key] = row[1];
    });

    // 🔍 Lọc trước các dòng cần xử lý
    var rowsToProcess = [];
    dataTarget.forEach(function(row, index) {
      var pnr = row[0]; // cột B
      var colF = row[4]; // cột F
      var colG = row[5]; // cột G
      var colH = row[6]; // cột H
      if (pnr && (!colF || !colG || !colH)) {
        rowsToProcess.push({ rowIndex: index + 1, data: row });
      }
    });

    Logger.log(`📌 Tổng dòng cần xử lý: ${rowsToProcess.length}`);

    // Xử lý các dòng lọc được
    rowsToProcess.forEach(function(item) {
      var rowIdx = item.rowIndex;
      var row = item.data;
      var pnr = row[0];
      var colC = row[1]; // cột C
      var colD = row[2]; // cột D
      var colE = row[3]; // cột E
      var colF = row[4];
      var colG = row[5];
      var colH = row[6];

      if (pnr && !colF) {
        var idgopnhom = "R" + rowIdx;
        targetSheet.getRange(rowIdx, 6).setValue(idgopnhom); // cột F
      }
      if (pnr && colG && !colH) {
        if (logMap[pnr] == "VJ") {
          targetSheet.getRange(rowIdx, 8).setValue(1);
        } 
        if (logMap[pnr] == "VNA") {
          targetSheet.getRange(rowIdx, 8).setValue(1);
        } 
      }
      if (pnr && !colG && logMap[pnr] !== undefined) {
        targetSheet.getRange(rowIdx, 7).setValue(logMap[pnr]); // cột G
        Logger.log(logMap[pnr]);
        if (logMap[pnr] == "VJ") {
          targetSheet.getRange(rowIdx, 8).setValue(1);
        }
        if (logMap[pnr] == "VNA") {
          targetSheet.getRange(rowIdx, 8).setValue(1);
        } 
      } else if (!pnr && colG) {
        targetSheet.getRange(rowIdx, 7).clearContent();
        targetSheet.getRange(rowIdx, 8).clearContent();
      }
    });

  } catch (e) {
    Logger.log("❌ Lỗi xảy ra: " + e);
  } finally {
    scriptProps.setProperty('checkhangvasomatve_isRunning', 'false');
  }
}

function checksomatve(pnr) {
  var url = "https://thuhongtour.com/check-so-mat-ve-vna/?pnr=" +pnr + "&ssid=check11";
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
function getLastNonEmptyRowInColumn(sheet, colIndex) { const vals = sheet.getRange(1, colIndex, sheet.getMaxRows()).getValues(); for (let i = vals.length - 1; i >= 0; i--) { const v = vals[i][0]; if (v !== "" && v !== null && v !== undefined && (String(v).trim() !== "")) { return i + 1; } } return 0; }

function doPost(e) {
  const props = PropertiesService.getScriptProperties();
  let debug = [];

  try {
    const data = JSON.parse(e.postData.contents || '{}');
    if (!data.khachHang || !Array.isArray(data.khachHang)) {
      throw new Error("Dữ liệu không hợp lệ, cần mảng 'khachHang'");
    }

    // push queue
    let queue = JSON.parse(props.getProperty("pushemailkhach_queue") || "[]");
    queue.push(data);
    props.setProperty("pushemailkhach_queue", JSON.stringify(queue));
    debug.push("Pushed job, queue len=" + queue.length);

    // thử lấy lock để xử lý luôn
    const lock = LockService.getScriptLock();
    if (lock.tryLock(500)) {
      try {
        debug.push("Got lock, start processing queue");
        processQueue(); // xử lý luôn
      } finally {
        lock.releaseLock();
      }
    } else {
      debug.push("Lock busy, sẽ được xử lý bởi request khác");
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "queued", debug }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString(), debug }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// xử lý queue, chạy lần lượt cho tới khi hết
function processQueue() {
  const props = PropertiesService.getScriptProperties();

  while (true) {
    let queue = JSON.parse(props.getProperty("pushemailkhach_queue") || "[]");
    if (!queue.length) break;

    const job = queue.shift();
    props.setProperty("pushemailkhach_queue", JSON.stringify(queue));

    try {
      processKhachHang(job); // xử lý thật sự
    } catch (err) {
      Logger.log("❌ Lỗi xử lý job: " + err);
      // nếu muốn retry thì push lại queue ở đây
    }
  }
}

function processKhachHang(data) {
  let debugLogs = [];
  const ss = SpreadsheetApp.openById("11lIiwBcRyBZMvJxBjdryQDBf9w6wU0nRH8_z5cco78Y");
  const sheetPNR = ss.getSheetByName("Danh sách PNR cần gửi mail");
  const sheetEmailSDT = ss.getSheetByName("Email-SDT");

  if (!data.khachHang || !Array.isArray(data.khachHang)) {
    throw new Error("Dữ liệu không hợp lệ, cần mảng 'khachHang'");
  }

  let lastPNRRow = getLastNonEmptyRowInColumn(sheetPNR, 2);
  let pnrData = lastPNRRow ? sheetPNR.getRange(1, 1, lastPNRRow, 10).getValues() : [];

  let lastEmailRow = getLastNonEmptyRowInColumn(sheetEmailSDT, 1);
  let emailList = lastEmailRow ? sheetEmailSDT.getRange(1, 1, lastEmailRow, 1).getValues().flat() : [];

  let rowsToInsert = [];

  data.khachHang.forEach(khach => {
    let idGop = null;
    if (khach.guiChung) {
      for (let i = 0; i < pnrData.length; i++) {
        const [ , pnrSheet, emailSheet, , , idSheet, , , , trangThai ] = pnrData[i];
        if (
          pnrSheet &&
          khach.pnrs.includes(pnrSheet.toString().trim()) &&
          emailSheet &&
          emailSheet.toString().trim().toLowerCase() === khach.email.trim().toLowerCase() &&
          (!trangThai || trangThai.toString().trim() === "")
        ) {
          idGop = idSheet;
          break;
        }
      }
      if (!idGop) idGop = "G" + Date.now() + "_" + khach.email;
    }

    // Cập nhật Email-SDT
    if (khach.sdt && khach.sdt.trim() !== "") {
      const emailIndex = emailList.findIndex(email => email && String(email).trim().toLowerCase() === khach.email.trim().toLowerCase());
      if (emailIndex !== -1) {
        sheetEmailSDT.getRange(emailIndex + 1, 2).setNumberFormat("@").setValue(khach.sdt.trim());
      } else {
        let newRow = sheetEmailSDT.getLastRow() + 1;
        sheetEmailSDT.getRange(newRow, 1, 1, 2).setNumberFormat("@");
        sheetEmailSDT.getRange(newRow, 1, 1, 2).setValues([[khach.email, khach.sdt.trim()]]);
        emailList.push(khach.email);
      }
    }

    // Thêm PNR mới
    khach.pnrs.forEach(pnr => {
      const isDuplicate = pnrData.some(row =>
        row[1] && row[1].toString().trim() === pnr.trim() &&
        row[2] && row[2].toString().trim().toLowerCase() === khach.email.trim().toLowerCase() &&
        (!row[9] || row[9].toString().trim() === "")
      );
      if (isDuplicate) return;

      const currentId = khach.guiChung ? idGop : "R" + (lastPNRRow + rowsToInsert.length + 1);

      rowsToInsert.push([
        "", pnr, khach.email, khach.tenKhach || "", khach.xungHo || "", currentId, "", "", khach.banner, ""
      ]);

      pnrData.push([null, pnr, khach.email, khach.tenKhach || "", khach.xungHo || "", currentId, null, null, khach.banner, ""]);
    });
  });

  if (rowsToInsert.length > 0) {
    sheetPNR.getRange(lastPNRRow + 1, 1, rowsToInsert.length, 10).setValues(rowsToInsert);
    SpreadsheetApp.flush();
  }
  afterDataUpdate();
}

function processPNR_safe_v2() {
  var props = PropertiesService.getScriptProperties();

  // 🔒 Check đang chạy hay không
  if (props.getProperty("processPNR_safe_v2_running") === "true") {
    Logger.log("⚠️ Đang chạy, bỏ qua lần này để tránh trùng");
    return;
  }

  // Đánh dấu đang chạy
  props.setProperty("processPNR_safe_v2_running", "true");

  try {
    var sheetNameTarget = "Danh sách PNR cần gửi mail";
    var spreadsheetId = "11lIiwBcRyBZMvJxBjdryQDBf9w6wU0nRH8_z5cco78Y";
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName(sheetNameTarget);

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var lastCol = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var DEBUG = true;

    function safeNum(val) {
      if (typeof val === 'number' && isFinite(val)) return val;
      if (val === null || val === undefined) return 0;
      var s = String(val).trim().replace(/\u00A0/g, '').replace(/,/g, '');
      var m = s.match(/-?\d+(\.\d+)?/);
      if (m) return parseFloat(m[0]);
      return 0;
    }

    function isGEmpty(val) {
      return val === null || val === undefined || (typeof val === 'string' && val.trim() === '');
    }

    // 🔎 lọc trước các dòng có G trống và B có giá trị
    var tasks = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (row[1] && isGEmpty(row[6])) { // B có giá trị, G trống
        tasks.push({ rowIndex: i + 2, row: row });
      }
    }

    if (DEBUG) Logger.log("Có " + tasks.length + " dòng cần xử lý");

    // Nếu không có gì thì thoát
    if (tasks.length === 0) return;

    var outBlock = [];

    tasks.forEach(function(task) {
      var row = task.row;
      var rowIndex = task.rowIndex;

      var colB = row[1];    // B
      var colG = row[6];    // G
      var colH = row[7];    // H
      var colI = row[8];    // I
      var colJ = row[9];    // J
      var colK = row[10];   // K
      var colLraw = row[11]; // L (raw)

      var lNum = safeNum(colLraw);
      if (DEBUG) Logger.log("Row " + rowIndex + " Lraw=" + colLraw + " parsed=" + lNum);

      if (lNum > 10) {
        outBlock.push([rowIndex, ["x", "x", colI, "x", colK, lNum + 1]]);
        return;
      }

      var pnr = String(colB);
      var newK = colK;
      var res;

      if (colK && colK.toString().toUpperCase() === "VJ") {
        res = sendmailVJ(pnr);
      } else if (colK && colK.toString().toUpperCase() === "VNA") {
        res = sendmailVNA(pnr);
      } else {
        res = sendmailVJ(pnr);
        if (!res || res.toLowerCase() === "none") {
          var res2 = sendmailVNA(pnr);
          if (res2 && res2.indexOf("ITINERARY RECEIPT EMAIL SENT") !== -1) newK = "VNA";
        } else if (res.indexOf("VJ") !== -1) {
          newK = "VJ";
        } else {
          var res3 = sendmailVNA(pnr);
          if (res3 && res3.indexOf("ITINERARY RECEIPT EMAIL SENT") !== -1) newK = "VNA";
        }
      }

      outBlock.push([rowIndex, [colG, colH, colI, colJ, newK, lNum + 1]]);
    });

    // 🔄 ghi kết quả cho từng dòng cần xử lý
    outBlock.forEach(function(item) {
      var rowIndex = item[0];
      var values = item[1];
      sheet.getRange(rowIndex, 7, 1, 6).setValues([values]); // G..L
    });

  } catch (err) {
    Logger.log("❌ Lỗi: " + err);
  } finally {
    // ✅ Clear flag khi chạy xong
    PropertiesService.getScriptProperties().deleteProperty("processPNR_safe_v2_running");
  }
}

// dummy funcs tạm
function sendmailVJ(pnr) {
  var url = "https://thuhongtour.com/sendmail_vj?pnr=" + encodeURIComponent(pnr);
  
  try {
    var response = UrlFetchApp.fetch(url, {
      'method': 'get',
      'muteHttpExceptions': true
    });
    Logger.log(response.getContentText())
    // Trả về body response
    return response.getContentText();
  } catch (e) {
    return "Lỗi: " + e.toString();
  }
}
function sendmailVNA(pnr) {
  var url = "https://thuhongtour.com/sendmailvna?ssid=mail&code=" + encodeURIComponent(pnr);
  
  try {
    var response = UrlFetchApp.fetch(url, {
      'method': 'get',
      'muteHttpExceptions': true
    });
    Logger.log(response.getContentText())
    // Trả về body response
    return response.getContentText();
  } catch (e) {
    return "Lỗi: " + e.toString();
  }
}
function copyRowsFromYesterday() {
  const ss = SpreadsheetApp.openById("11lIiwBcRyBZMvJxBjdryQDBf9w6wU0nRH8_z5cco78Y");
  const sheetSrc = ss.getSheetByName("LogGuiEmailPNR-KhachHang");
  const sheetDst = ss.getSheetByName("Gửi mail sau 12h xuất vé");
  const sheetBL = ss.getSheetByName("Danh Sách BL gửi email chăm sóc");

  const data = sheetSrc.getDataRange().getValues();
  if (data.length <= 1) return; // chỉ có header

  // lấy blacklist email từ sheetBL cột A
  const blVals = sheetBL.getRange(1, 1, sheetBL.getLastRow(), 1).getValues();
  const blacklist = new Set(blVals.flat().filter(e => e && e.toString().trim() !== ""));

  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const yesterdayStr = Utilities.formatDate(
    new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1),
    tz,
    "dd/MM/yyyy"
  );

  let rowsToCopy = [];

  for (let i = 1; i < data.length; i++) { // bỏ header
    const value = data[i][17]; // cột R = index 17 (0-based)
    if (!value) continue;

    let dateStr = "";
    if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
      dateStr = Utilities.formatDate(value, tz, "dd/MM/yyyy");
    } else {
      try {
        const parts = value.split(" ");
        const dmy = parts[0].split("/");
        const parsed = new Date(dmy[2], dmy[1] - 1, dmy[0]);
        dateStr = Utilities.formatDate(parsed, tz, "dd/MM/yyyy");
      } catch (e) {
        continue;
      }
    }

    // check trùng ngày và không nằm trong blacklist
    const email = (data[i][2] || "").toString().trim().toLowerCase();
    if (dateStr === yesterdayStr && !blacklist.has(email)) {
      rowsToCopy.push(data[i]);
    }
  }

  // clear sheet đích nhưng giữ lại dòng header (xóa từ dòng 2 trở đi)
  if (sheetDst.getLastRow() > 1) {
    sheetDst.getRange(2, 1, sheetDst.getLastRow() - 1, sheetDst.getLastColumn()).clearContent();
  }

  if (rowsToCopy.length) {
    // giữ dòng cuối cùng theo C+E
    let uniqueMap = new Map();
    rowsToCopy.forEach(row => {
      const key = row[2] + "|" + row[4];
      uniqueMap.set(key, row);
    });

    const uniqueRows = Array.from(uniqueMap.values());

    sheetDst.getRange(2, 1, uniqueRows.length, uniqueRows[0].length).setValues(uniqueRows);
  }
}
function copyRowsTomorrowFlights() {
  const ss = SpreadsheetApp.openById("11lIiwBcRyBZMvJxBjdryQDBf9w6wU0nRH8_z5cco78Y");
  const sheetSrc = ss.getSheetByName("LogGuiEmailPNKhachHang") || ss.getSheetByName("LogGuiEmailPNR-KhachHang");
  const sheetDst = ss.getSheetByName("Gửi mail trước 24h bay");
  const sheetBL = ss.getSheetByName("Danh Sách BL gửi email chăm sóc");

  const data = sheetSrc.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const tomorrow = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, tz, "dd/MM/yyyy");

  // lấy blacklist email
  const blLastRow = Math.max(1, sheetBL.getLastRow());
  const blVals = sheetBL.getRange(1, 1, blLastRow, 1).getValues();
  const blacklist = new Set(blVals.flat().filter(e => e && e.toString().trim() !== "").map(e => e.toString().toLowerCase().trim()));

  let rowsToCopy = [];

  for (let i = 1; i < data.length; i++) { // bỏ header
    const origRow = data[i];
    let row = [...origRow]; // copy để sửa

    const airline = (row[18] || "").toString().trim(); // cột S (index 18)

    // gom các bộ trip,day,time (F–Q) và chuẩn hoá ngay
    let segments = [];
    for (let base = 5; base <= 14; base += 3) { // F(5), I(8), L(11), O(14)
      const tripRaw = row[base];
      const dayRaw = row[base + 1];
      const timeRaw = row[base + 2];
      if (tripRaw && dayRaw) {
        const dayNorm = normalizeDate(dayRaw, tz);           // "dd/MM/yyyy"
        const timeNorm = formatTime(timeRaw, tz);           // "HH:mm"
        const dt = parseDateTime(dayNorm, timeNorm);        // Date object or null
        segments.push({
          trip: tripRaw.toString(),
          dayRaw, timeRaw,
          day: dayNorm,
          time: timeNorm,
          dateTime: dt,
          base
        });
      }
    }

    if (segments.length === 0) continue;

    // tìm xem có ít nhất 1 leg có day = tomorrow (điều kiện ban đầu)
    const hasTomorrow = segments.some(s => s.day === tomorrowStr);
    if (!hasTomorrow) continue;

    // xử lý theo hãng
    if (airline === "VJ") {
      // VJ: chỉ giữ leg có day = tomorrow
      segments = segments.filter(s => s.day === tomorrowStr);

    } else if (airline === "VNA") {
      // VNA: xử lý theo số bộ
      if (segments.length === 1) {
        segments = segments.filter(s => s.day === tomorrowStr);

      } else if (segments.length === 2) {
        const domesticIdx = segments.findIndex(s => isDomesticVN(s.trip));
        if (domesticIdx !== -1) {
          // copy day+time từ bộ có domestic sang bộ còn lại nếu lệch <=24h
          if (domesticIdx === 0) copyDayTimeSafe(segments[0], segments[1]);
          else copyDayTimeSafe(segments[1], segments[0]);
        }
        segments = segments.filter(s => s.day === tomorrowStr);

      } else if (segments.length === 3) {
        const vnIdx = segments.findIndex(s => isDomesticVN(s.trip));
        if (vnIdx === 0) {
          // copy 1 -> 2
          copyDayTimeSafe(segments[0], segments[1]);
        } else if (vnIdx === 2) {
          // copy 2 -> 3
          copyDayTimeSafe(segments[1], segments[2]);
        } else if (vnIdx === 1) {
          // so sánh gần với bộ 1 hay bộ 3
          const nearFirst = closerTo(segments[1], segments[0], segments[2]) === 0;
          if (nearFirst) {
            copyDayTimeSafe(segments[0], segments[1]);
          } else {
            copyDayTimeSafe(segments[1], segments[2]);
          }
        }
        segments = segments.filter(s => s.day === tomorrowStr);

      } else if (segments.length === 4) {
        // copy 1->2, 3->4 (nhưng chỉ copy nếu <=24h)
        copyDayTimeSafe(segments[0], segments[1]);
        copyDayTimeSafe(segments[2], segments[3]);
        segments = segments.filter(s => s.day === tomorrowStr);
      }
    }

    // sau khi xử lý, nếu vẫn có segment thỏa tomorrow thì ghi
    if (segments.length > 0) {
      // clear F→Q
      for (let base = 5; base <= 16; base++) row[base] = "";

      // ghi các segment (bắt đầu từ F)
      segments.forEach((s, idx) => {
        const base = 5 + idx * 3;
        row[base] = s.trip;
        row[base + 1] = s.day;  // dd/MM/yyyy
        row[base + 2] = s.time; // HH:mm
      });

      // kiểm blacklist
      const email = (row[2] || "").toString().trim().toLowerCase();
      if (!blacklist.has(email)) {
        rowsToCopy.push(row);
      }
    }
  }

  // lọc trùng (C|E)
  let uniqueMap = new Map();
  rowsToCopy.forEach(row => {
    const key = (row[2] || "") + "|" + (row[4] || "");
    uniqueMap.set(key, row);
  });
  const uniqueRows = Array.from(uniqueMap.values());

  // clear old (giữ header)
  if (sheetDst.getLastRow() > 1) {
    sheetDst.getRange(2, 1, sheetDst.getLastRow() - 1, sheetDst.getLastColumn()).clearContent();
  }

  if (uniqueRows.length) {
    sheetDst.getRange(2, 1, uniqueRows.length, uniqueRows[0].length).setValues(uniqueRows);
  }
}

/* ================= helper functions ================= */

function normalizeDate(val, tz) {
  if (!val && val !== 0) return "";
  if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val)) {
    return Utilities.formatDate(val, tz, "dd/MM/yyyy");
  }
  const s = val.toString().trim();
  // nếu đã dd/MM/yyyy
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const d = parseInt(m[1], 10);
    const mo = parseInt(m[2], 10) - 1;
    const y = parseInt(m[3], 10);
    return Utilities.formatDate(new Date(y, mo, d), tz, "dd/MM/yyyy");
  }
  // fallback try parse Date
  const dt = new Date(s);
  if (!isNaN(dt)) return Utilities.formatDate(dt, tz, "dd/MM/yyyy");
  return "";
}

function formatTime(val, tz) {
  if (val === null || val === undefined || val === "") return "";
  if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val)) {
    return Utilities.formatDate(val, tz, "HH:mm");
  }
  const s = val.toString().trim();
  const m = s.match(/(\d{1,2}):(\d{2})/);
  if (m) {
    return String(m[1]).padStart(2, "0") + ":" + m[2];
  }
  // try parse
  const dt = new Date(s);
  if (!isNaN(dt)) return Utilities.formatDate(dt, tz, "HH:mm");
  return s;
}

function parseDateTime(dayStr, timeStr) {
  try {
    if (!dayStr) return null;
    const parts = dayStr.split("/");
    if (parts.length !== 3) return null;
    const d = parseInt(parts[0], 10), m = parseInt(parts[1], 10) - 1, y = parseInt(parts[2], 10);
    let hh = 0, mm = 0;
    if (timeStr && typeof timeStr === "string" && timeStr.indexOf(":") >= 0) {
      const t = timeStr.split(":");
      hh = parseInt(t[0], 10) || 0;
      mm = parseInt(t[1], 10) || 0;
    } else if (Object.prototype.toString.call(timeStr) === "[object Date]" && !isNaN(timeStr)) {
      hh = timeStr.getHours();
      mm = timeStr.getMinutes();
    }
    return new Date(y, m, d, hh, mm);
  } catch (e) {
    return null;
  }
}

// trả true nếu cả 2 IATA đều trong danh sách VN
function isDomesticVN(trip) {
  if (!trip) return false;
  const vnIata = ["SGN","HAN","DAD","HPH","VII","CXR","HUI","PQC","VCA","THD","BMV","DIN","VCL","TBB","PXU","UIH","VCS","VKG","DLI"];
  const parts = trip.toString().split("-");
  return parts.length === 2 && vnIata.includes(parts[0]) && vnIata.includes(parts[1]);
}

// copy day+time từ src->dst **chỉ** khi chênh <= 24h, nếu >24h thì không thay đổi dst (giữ nguyên)
function copyDayTimeSafe(src, dst) {
  // src and dst must have .dateTime, .day, .time
  const srcDT = src.dateTime || parseDateTime(src.day, src.time);
  const dstDT = dst.dateTime || parseDateTime(dst.day, dst.time);
  if (srcDT && dstDT) {
    const diff = Math.abs(srcDT - dstDT);
    if (diff <= 24 * 3600 * 1000) {
      dst.day = src.day;
      dst.time = src.time;
      dst.dateTime = srcDT;
    } else {
      // lệch >24h => GIỮ NGUYÊN dst (không thay)
    }
  } else {
    // nếu không parse được, fallback: copy (để an toàn)
    dst.day = src.day;
    dst.time = src.time;
    dst.dateTime = src.dateTime || srcDT || parseDateTime(dst.day, dst.time);
  }
}

// trả 0 nếu gần first hơn, trả 2 nếu gần last hơn
function closerTo(mid, first, last) {
  const midDT = mid.dateTime || parseDateTime(mid.day, mid.time);
  const firstDT = first.dateTime || parseDateTime(first.day, first.time);
  const lastDT = last.dateTime || parseDateTime(last.day, last.time);
  if (!midDT || !firstDT || !lastDT) return 0;
  const d1 = Math.abs(midDT - firstDT);
  const d2 = Math.abs(midDT - lastDT);
  return d1 <= d2 ? 0 : 2;
}



function copyRowsAfter24hFlights() {
  const ss = SpreadsheetApp.openById("11lIiwBcRyBZMvJxBjdryQDBf9w6wU0nRH8_z5cco78Y");
  const sheetSrc = ss.getSheetByName("LogGuiEmailPNR-KhachHang");
  const sheetDst = ss.getSheetByName("Gửi mail sau 24h bay");
  const sheetBL = ss.getSheetByName("Danh Sách BL gửi email chăm sóc");

  const data = sheetSrc.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const yesterdayStr = Utilities.formatDate(
    new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1),
    tz,
    "dd/MM/yyyy"
  );

  // lấy blacklist email
  const blVals = sheetBL.getRange(1, 1, sheetBL.getLastRow(), 1).getValues();
  const blacklist = new Set(
    blVals.flat()
      .filter(e => e && e.toString().trim() !== "")
      .map(e => e.toString().toLowerCase().trim())
  );

  let rowsToCopy = [];

  for (let i = 1; i < data.length; i++) { // bỏ header
    // check theo thứ tự G,J,M,P → lấy cái cuối cùng có giá trị
    const colIdxs = [6, 9, 12, 15];
    let lastVal = null;

    colIdxs.forEach(idx => {
      if (data[i][idx]) {
        lastVal = data[i][idx];
      }
    });

    if (!lastVal) continue;

    let dateStr = "";
    if (Object.prototype.toString.call(lastVal) === "[object Date]" && !isNaN(lastVal)) {
      dateStr = Utilities.formatDate(lastVal, tz, "dd/MM/yyyy");
    } else {
      try {
        const dmy = String(lastVal).split("/");
        const parsed = new Date(dmy[2], dmy[1] - 1, dmy[0]);
        dateStr = Utilities.formatDate(parsed, tz, "dd/MM/yyyy");
      } catch (e) {
        continue;
      }
    }

    const email = (data[i][2] || "").toString().trim().toLowerCase();

    if (dateStr === yesterdayStr && !blacklist.has(email)) {
      rowsToCopy.push(data[i]);
    }
  }

  // clear sheet đích nhưng giữ lại header
  if (sheetDst.getLastRow() > 1) {
    sheetDst.getRange(2, 1, sheetDst.getLastRow() - 1, sheetDst.getLastColumn()).clearContent();
  }

  if (rowsToCopy.length) {
    sheetDst.getRange(2, 1, rowsToCopy.length, rowsToCopy[0].length).setValues(rowsToCopy);
  }
}


