/**
 * GOOGLE APPS SCRIPT FOR ROOMFLOW HOTEL MANAGEMENT SYSTEM
 * 
 * Target Client Sheet Structure (22 Columns):
 * Col A (1) : Sr No
 * Col B (2) : His Name 👦🏻
 * Col C (3) : His Mobile
 * Col D (4) : His Identity Details (Aadhaar / PAN / Passport)
 * Col E (5) : Her Name 👧🏻
 * Col F (6) : Her Mobile
 * Col G (7) : Her Identity Details (Aadhaar / PAN / Passport)
 * Col H (8) : Amount
 * Col I (9) : Discount (₹)
 * Col J (10): Additional Charges (₹)
 * Col K (11): Final Price (₹)
 * Col L (12): Advance Amount (₹)
 * Col M (13): Due / Remaining Amount (₹)
 * Col N (14): Mode of Payment
 * Col O (15): Check-In-Date
 * Col P (16): Check-In-Time
 * Col Q (17): Alloted Room Number
 * Col R (18): Time Duration
 * Col S (19): Hour / Day
 * Col T (20): Check-Out-Date
 * Col U (21): Check-Out-Time
 * Col V (22): Stay ID
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
   
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    var action = data.action || "";
    var targetSheet = data.targetSheet || "";

    if (targetSheet === "Staff Attendance" || action === "staff_add" || action === "staff_delete" || action === "attendance_mark" || action === "batch_attendance_save") {
      return handleStaffAttendance(ss, data);
    } else {
      return handleClientData(ss, data);
    }
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "error", "error": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var action = e ? (e.parameter.action || "") : "";
   
    if (action === "get_staff" || action === "get_attendance" || (e && e.parameter && e.parameter.targetSheet === "Staff Attendance")) {
      return handleStaffAttendance(ss, e ? e.parameter : {});
    }

    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "success", "data": data }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "error", "error": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * HANDLE STAFF ATTENDANCE & STAFF MANAGEMENT (8 Columns)
 */
function handleStaffAttendance(ss, data) {
  var action = data.action || "";
  var sheet = ss.getSheetByName("Staff Attendance");

  var targetHeaders = [
    "Staff ID",
    "Staff Name",
    "Role",
    "Mobile",
    "Date",
    "Time",
    "Status / Details",
    "Action Type"
  ];

  if (!sheet) {
    sheet = ss.insertSheet("Staff Attendance");
    sheet.getRange(1, 1, 1, 8).setValues([targetHeaders]);
  } else {
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 8).setValues([targetHeaders]);
    } else {
      var firstCell = String(sheet.getRange(1, 1).getValue()).trim();
      if (firstCell === "Date" || firstCell === "") {
        sheet.getRange(1, 1, 1, 8).setValues([targetHeaders]);
      }
    }
  }

  var timeZone = ss.getSpreadsheetTimeZone();

  if (action === "batch_attendance_save") {
    var records = data.records || [];
    if (typeof records === "string") {
      try { records = JSON.parse(records); } catch (e) { records = []; }
    }

    var addedCount = 0;
    var updatedCount = 0;

    for (var i = 0; i < records.length; i++) {
      var res = saveOrUpdateAttendanceRow(sheet, records[i], timeZone);
      if (res === "added") addedCount++;
      else if (res === "updated") updatedCount++;
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        result: "success",
        message: "Attendance batch processed",
        added: addedCount,
        updated: updatedCount
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "attendance_mark") {
    var res = saveOrUpdateAttendanceRow(sheet, data, timeZone);
    return ContentService
      .createTextOutput(JSON.stringify({ result: "success", message: "Attendance marked (" + res + ")" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "staff_add" || action === "staff_delete") {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "success", message: "Staff action recorded: " + action }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var dataRange = sheet.getDataRange().getValues();
  return ContentService
    .createTextOutput(JSON.stringify({ result: "success", data: dataRange }))
    .setMimeType(ContentService.MimeType.JSON);
}

function saveOrUpdateAttendanceRow(sheet, rec, timeZone) {
  var staffId = String(rec.staffId || rec.id || rec["Staff ID"] || "").trim();
  var staffName = String(rec.staffName || rec.name || rec["Staff Name"] || "").trim();
  var role = String(rec.role || rec["Role"] || "").trim();
  var mobile = String(rec.mobile || rec.phone || rec.staffMobile || rec.mobileNumber || rec["Mobile"] || "").trim();
  var dateStr = String(rec.date || rec["Date"] || "").trim();
  var timeStr = String(rec.time || rec["Time"] || "").trim();
  var status = String(rec.status || rec.details || rec["Status / Details"] || "").trim();
  var actionType = String(rec.actionType || rec["Action Type"] || "").trim();

  if (!dateStr) {
    dateStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
  }

  if (!timeStr) {
    timeStr = Utilities.formatDate(new Date(), timeZone, "HH:mm");
  }

  var lastRow = sheet.getLastRow();
  var updatedRowIndex = -1;

  if (lastRow > 1) {
    var existingData = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    for (var r = 0; r < existingData.length; r++) {
      var rowStaffId = String(existingData[r][0]).trim();
      var rowDate = existingData[r][4];
      if (rowDate instanceof Date) {
        rowDate = Utilities.formatDate(rowDate, timeZone, "yyyy-MM-dd");
      } else {
        rowDate = String(rowDate).trim();
      }

      if (rowStaffId === staffId && rowDate === dateStr) {
        updatedRowIndex = r + 2;
        break;
      }
    }
  }

  var rowValues = [
    staffId,
    staffName,
    role,
    mobile,
    dateStr,
    timeStr,
    status,
    actionType
  ];

  if (updatedRowIndex > 1) {
    sheet.getRange(updatedRowIndex, 1, 1, 8).setValues([rowValues]);
    return "updated";
  } else {
    sheet.appendRow(rowValues);
    return "added";
  }
}

/**
 * HELPER: Robust Extraction of His Identity / Aadhaar
 */
function getHisIdentity(data) {
  var val = data["His Identity Details (Aadhaar / PAN / Passport)"] ||
            data["His Aadhaar Card Number"] ||
            data["His Identity Details"] ||
            data["His Aadhaar"] ||
            data.hisIdentity ||
            data.hisAadhaar ||
            data.hisAadhar ||
            data.his_identity ||
            data.his_aadhaar ||
            data.his_aadhar ||
            "";
  val = String(val).trim();
  return val !== "" ? val : "-";
}

/**
 * HELPER: Robust Extraction of Her Identity / Aadhaar
 */
function getHerIdentity(data) {
  var val = data["Her Identity Details (Aadhaar / PAN / Passport)"] ||
            data["Her Aadhaar Card Number"] ||
            data["Her Identity Details"] ||
            data["Her Aadhaar"] ||
            data.herIdentity ||
            data.herAadhaar ||
            data.herAadhar ||
            data.her_identity ||
            data.her_aadhaar ||
            data.her_aadhar ||
            "";
  val = String(val).trim();
  return val !== "" ? val : "-";
}

/**
 * HANDLE CLIENT DATA SAVING & UPDATES (EXACT 22 COLUMNS)
 */
function handleClientData(ss, data) {
  var targetSheetName = data.targetSheet || "Sheet1";
  var sheet = ss.getSheetByName(targetSheetName) || ss.getSheets()[0];

  var targetHeaders = [
    "Sr No",
    "His Name 👦🏻",
    "His Mobile",
    "His Identity Details (Aadhaar / PAN / Passport)",
    "Her Name 👧🏻",
    "Her Mobile",
    "Her Identity Details (Aadhaar / PAN / Passport)",
    "Amount",
    "Discount (₹)",
    "Additional Charges (₹)",
    "Final Price (₹)",
    "Advance Amount (₹)",
    "Due / Remaining Amount (₹)",
    "Mode of Payment",
    "Check-In-Date",
    "Check-In-Time",
    "Alloted Room Number",
    "Time Duration",
    "Hour / Day",
    "Check-Out-Date",
    "Check-Out-Time",
    "Stay ID"
  ];

  // Enforce Row 1 Header Alignment to 22 Columns
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 22).setValues([targetHeaders]);
  } else {
    var firstRowVals = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
    if (firstRowVals.length < 22 || String(firstRowVals[8]).indexOf("Discount") === -1) {
      sheet.getRange(1, 1, 1, 22).setValues([targetHeaders]);
    }
  }

  var action = data.action || "";

  // -------------------------------------------------------------
  // 1. UPDATE DUE / REMAINING AMOUNT ONLY (IN-PLACE CELL EDIT)
  // -------------------------------------------------------------
  if (action === "update_due" || action === "due_update") {
    var dueVal = data["Due / Remaining Amount (₹)"] !== undefined ? data["Due / Remaining Amount (₹)"] : (data.remainingAmount !== undefined ? data.remainingAmount : (data.dueAmount !== undefined ? data.dueAmount : ""));
    var targetId = String(data.id || data["Stay ID"] || "").trim();
    var targetRoom = String(data["Alloted Room Number"] || data.allotedRoomNumber || data.room || "").trim();
    var targetName = String(data["His Name 👦🏻"] || data.hisName || "").trim();

    var lastRow = sheet.getLastRow();
    var updatedRowIndex = -1;

    if (lastRow > 1) {
      var allRows = sheet.getRange(2, 1, lastRow - 1, 22).getValues();

      for (var r = allRows.length - 1; r >= 0; r--) {
        var rowId = String(allRows[r][21] || "").trim(); // Col V (22): Stay ID
        var rowRoom = String(allRows[r][16] || "").trim(); // Col Q (17): Alloted Room Number
        var rowHisName = String(allRows[r][1] || "").trim(); // Col B (2): His Name

        if ((targetId && rowId === targetId) ||
            (targetRoom && (rowRoom === targetRoom || rowRoom.indexOf(targetRoom) !== -1 || targetRoom.indexOf(rowRoom) !== -1)) ||
            (targetName && rowHisName === targetName)) {
          updatedRowIndex = r + 2;
          break;
        }
      }
    }

    if (updatedRowIndex > 1) {
      sheet.getRange(updatedRowIndex, 13).setValue(dueVal); // Col M (13): Due / Remaining Amount (₹)
      return ContentService
        .createTextOutput(JSON.stringify({ "result": "success", "message": "Updated due amount on row " + updatedRowIndex }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // MANDATORY RETURN EARLY - NEVER APPEND A NEW ROW ON UPDATE_DUE!
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "success", "message": "No row matched for due update" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // -------------------------------------------------------------
  // 2. CHECKOUT ACTION (IN-PLACE CELL EDIT)
  // -------------------------------------------------------------
  if (action === "checkout") {
    var targetId = String(data.id || data["Stay ID"] || "").trim();
    var targetRoom = String(data["Alloted Room Number"] || data.allotedRoomNumber || data.room || "").trim();
    var targetRoomDigits = targetRoom.replace(/[^0-9a-zA-Z]/g, '').toLowerCase();

    var checkoutDate = data.checkoutDate || data.checkoutDateFormatted || "";
    var checkoutTime = data.checkoutTime || data.checkoutTimeFormatted || "";

    var lastRow = sheet.getLastRow();
    var updatedRowIndex = -1;

    if (lastRow > 1) {
      var allRows = sheet.getRange(2, 1, lastRow - 1, 22).getValues();

      for (var r = allRows.length - 1; r >= 0; r--) {
        var rowId = String(allRows[r][21] || "").trim(); // Col V
        var rowRoom = String(allRows[r][16] || "").trim(); // Col Q
        var rowRoomDigits = rowRoom.replace(/[^0-9a-zA-Z]/g, '').toLowerCase();
        var rowCheckout = String(allRows[r][19] || "").trim(); // Col T

        var isActive = (rowCheckout === "" || rowCheckout === "-" || rowCheckout === "null" || rowCheckout === "undefined");

        if (isActive && (
          (targetId && rowId === targetId) ||
          (targetRoom && (rowRoom === targetRoom || rowRoomDigits === targetRoomDigits || rowRoom.indexOf(targetRoom) !== -1))
        )) {
          updatedRowIndex = r + 2;
          break;
        }
      }
    }

    if (updatedRowIndex > 1) {
      if (checkoutDate) sheet.getRange(updatedRowIndex, 20).setValue(checkoutDate); // Col T (20)
      if (checkoutTime) sheet.getRange(updatedRowIndex, 21).setValue(checkoutTime); // Col U (21)

      return ContentService
        .createTextOutput(JSON.stringify({ "result": "success", "message": "Updated checkout on row " + updatedRowIndex }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ "result": "success", "message": "Checkout logged" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // -------------------------------------------------------------
  // 3. MASTER RESET ACTION (IN-PLACE CELL EDIT)
  // -------------------------------------------------------------
  if (action === "reset") {
    var checkoutDate = data.checkoutDate || "";
    var checkoutTime = data.checkoutTime || "";
    var lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      var allRows = sheet.getRange(2, 1, lastRow - 1, 22).getValues();
      for (var r = 0; r < allRows.length; r++) {
        var rowCheckout = String(allRows[r][19] || "").trim();
        if (rowCheckout === "" || rowCheckout === "-") {
          var rowIndex = r + 2;
          sheet.getRange(rowIndex, 20).setValue(checkoutDate); // Col T
          sheet.getRange(rowIndex, 21).setValue(checkoutTime); // Col U
        }
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ "result": "success", "message": "Dashboard reset updated" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // -------------------------------------------------------------
  // 4. ADD NEW CLIENT RECORD (STRICT 22-COLUMN MAPPING)
  // -------------------------------------------------------------
  var lastRow = sheet.getLastRow();
  var nextSrNo = Math.max(1, lastRow);

  var rawSr = data.srNo;
  if (typeof rawSr === "string" && (rawSr.indexOf("RF_") === 0 || isNaN(Number(rawSr)))) {
    rawSr = null;
  }
  var cleanSrNo = rawSr ? Number(rawSr) : nextSrNo;

  var row22 = [
    cleanSrNo,                                                                          // Col A (1) : Sr No
    data["His Name 👦🏻"] || data.hisName || data.his_name || "-",                       // Col B (2) : His Name
    data["His Mobile"] || data.hisMobile || data.his_mobile || "-",                   // Col C (3) : His Mobile
    getHisIdentity(data),                                                               // Col D (4) : His Aadhaar / Identity
    data["Her Name 👧🏻"] || data.herName || data.her_name || "-",                       // Col E (5) : Her Name
    data["Her Mobile"] || data.herMobile || data.her_mobile || "-",                   // Col F (6) : Her Mobile
    getHerIdentity(data),                                                               // Col G (7) : Her Aadhaar / Identity
    data["Amount"] !== undefined ? data["Amount"] : (data.amount !== undefined ? data.amount : 0), // Col H (8) : Amount
    data["Discount (₹)"] !== undefined ? data["Discount (₹)"] : (data.discount !== undefined ? data.discount : 0), // Col I (9) : Discount (₹)
    data["Additional Charges (₹)"] !== undefined ? data["Additional Charges (₹)"] : (data.additionalCharges !== undefined ? data.additionalCharges : 0), // Col J (10): Additional Charges (₹)
    data["Final Price (₹)"] !== undefined ? data["Final Price (₹)"] : (data.finalPrice !== undefined ? data.finalPrice : 0), // Col K (11): Final Price (₹)
    data["Advance Amount (₹)"] !== undefined ? data["Advance Amount (₹)"] : (data.advanceAmount !== undefined ? data.advanceAmount : ""), // Col L (12): Advance Amount (₹)
    data["Due / Remaining Amount (₹)"] !== undefined ? data["Due / Remaining Amount (₹)"] : (data.remainingAmount !== undefined ? data.remainingAmount : (data.dueAmount !== undefined ? data.dueAmount : "")), // Col M (13): Due / Remaining Amount (₹)
    data["Mode of Payment"] || data.paymentMode || "-",                                // Col N (14): Mode of Payment
    data["Check-In-Date"] || data.checkinDateFormatted || data.checkinDate || "-",     // Col O (15): Check-In-Date
    data["Check-In-Time"] || data.checkinTimeFormatted || data.checkinTime || "-",     // Col P (16): Check-In-Time
    data["Alloted Room Number"] || data.allotedRoomNumber || data.allottedRoom || data.room || "-", // Col Q (17): Alloted Room Number
    data["Time Duration"] || data.duration || "-",                                     // Col R (18): Time Duration
    data["Hour / Day"] || data.durationUnit || data.timeUnit || "Hour",                // Col S (19): Hour / Day
    data["Check-Out-Date"] || data.checkoutDateFormatted || data.checkoutDate || "-", // Col T (20): Check-Out-Date
    data["Check-Out-Time"] || data.checkoutTimeFormatted || data.checkoutTime || "-", // Col U (21): Check-Out-Time
    data["Stay ID"] || data.id || ""                                                    // Col V (22): Stay ID
  ];

  sheet.appendRow(row22);

  return ContentService
    .createTextOutput(JSON.stringify({ "result": "success", "row": row22 }))
    .setMimeType(ContentService.MimeType.JSON);
}