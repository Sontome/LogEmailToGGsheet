function getLastNonEmptyRowInColumn(sheet, colIndex) { const vals = sheet.getRange(1, colIndex, sheet.getMaxRows()).getValues(); for (let i = vals.length - 1; i >= 0; i--) { const v = vals[i][0]; if (v !== "" && v !== null && v !== undefined && (String(v).trim() !== "")) { return i + 1; } } return 0; }

function doPost(e) {
  let debugLogs = [];
  try {
    debugLogs.push("Bắt đầu doPost");

    const ss = SpreadsheetApp.openById("11lIiwBcRyBZMvJxBjdryQDBf9w6wU0nRH8_z5cco78Y");
    debugLogs.push("Mở file: " + ss.getName());

    const sheetPNR = ss.getSheetByName("Danh sách PNR cần gửi mail");
    const sheetEmailSDT = ss.getSheetByName("Email-SDT");

    const data = JSON.parse(e.postData.contents || '{}');
    debugLogs.push("Data nhận: " + JSON.stringify(data));

    if (!data.khachHang || !Array.isArray(data.khachHang)) {
      throw new Error("Dữ liệu không hợp lệ, cần mảng 'khachHang'");
    }

    let lastPNRRow = getLastNonEmptyRowInColumn(sheetPNR, 2);
    let pnrData = lastPNRRow ? sheetPNR.getRange(1, 1, lastPNRRow, 10).getValues() : [];

    let lastEmailRow = getLastNonEmptyRowInColumn(sheetEmailSDT, 1);
    let emailList = lastEmailRow ? sheetEmailSDT.getRange(1, 1, lastEmailRow, 1).getValues().flat() : [];

    let rowsToInsert = [];

    data.khachHang.forEach(khach => {
      debugLogs.push("Xử lý khách: " + khach.email);
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
          "", pnr, khach.email, khach.tenKhach || "", khach.xungHo || "", currentId, "", "", "", ""
        ]);

        pnrData.push([null, pnr, khach.email, khach.tenKhach || "", khach.xungHo || "", currentId, null, null, null, ""]);
      });
    });

    if (rowsToInsert.length > 0) {
      sheetPNR.getRange(lastPNRRow + 1, 1, rowsToInsert.length, 10).setValues(rowsToInsert);
    }
    afterDataUpdate()
    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", message: "Đã ghi dữ liệu vào sheet", debug: debugLogs }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    debugLogs.push("Lỗi: " + err);
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString(), debug: debugLogs }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
