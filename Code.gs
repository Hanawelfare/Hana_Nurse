// HANA Nurse Room System Backend (Google Apps Script API)
// New Spreadsheet URL: https://docs.google.com/spreadsheets/d/1-2iiEsfuHmKwT4FFoFs96BTC-X3eOdAUgRX0M9dqifQ/edit

var SPREADSHEET_ID = "1-2iiEsfuHmKwT4FFoFs96BTC-X3eOdAUgRX0M9dqifQ";

// Serve the API endpoints for external requests (CORS enabled)
function doPost(e) {
  var payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Invalid JSON format" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  var action = payload.action;
  var args = payload.args || [];
  var result;
  
  try {
    if (action === "checkLogin") {
      result = checkLogin(args[0]);
    } else if (action === "getEmployeeData") {
      result = getEmployeeData(args[0]);
    } else if (action === "saveVisitRecord") {
      result = saveVisitRecord(args[0]);
    } else if (action === "updateBedEndTime") {
      result = updateBedEndTime(args[0], args[1]);
    } else if (action === "getServiceHistory") {
      result = getServiceHistory(args[0]);
    } else if (action === "editHistoryRecord") {
      result = editHistoryRecord(args[0]);
    } else if (action === "deleteHistoryRecord") {
      result = deleteHistoryRecord(args[0]);
    } else if (action === "getDashboardData") {
      result = getDashboardData(args[0]);
    } else if (action === "getReportData") {
      result = getReportData(args[0]);
    } else if (action === "getMedicineStock") {
      result = getMedicineStock();
    } else if (action === "addMedicineStock") {
      result = addMedicineStock(args[0], args[1]);
    } else if (action === "getMedicationHistoryData") {
      result = getMedicationHistoryData(args[0], args[1]);
    } else if (action === "initializeSheets") {
      result = initializeSheets();
    } else {
      throw new Error("Action not found");
    }
    
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Serves plain text ping for checkups
function doGet(e) {
  return ContentService.createTextOutput("HANA Nurse Room API is running. Send POST requests to communicate with sheets.")
    .setMimeType(ContentService.MimeType.TEXT);
}

// Helper to open spreadsheet securely
function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (err) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

// Initialize all required sheets with headers and default stock
function initializeSheets() {
  var ss = getSpreadsheet();
  
  // 1. Name Sheet
  var empSheet = ss.getSheetByName("Name");
  if (!empSheet) {
    empSheet = ss.insertSheet("Name");
    empSheet.appendRow(["รหัสพนักงาน", "ชื่อ", "นามสกุล", "แผนก"]);
    empSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#e0e0e0");
  }
  
  // 2. History Sheet
  var histSheet = ss.getSheetByName("History");
  if (!histSheet) {
    histSheet = ss.insertSheet("History");
    histSheet.appendRow([
      "รหัสพนักงาน", "Date", "Time", "Month", "First Name", "Last Name", 
      "Shift", "Department", "ระบบโรค", "อาการ", "DX", "From Work", 
      "Back To Work", "Accident", "Rest", "Refer", "Year", 
      "จ่ายยา (Medication)", "Bed Start (เวลาเริ่ม)", "Bed End (เวลาตื่น)"
    ]);
    histSheet.getRange("A1:T1").setFontWeight("bold").setBackground("#d9ead3");
  }
  
  // 3. Checkup Sheet
  var chkSheet = ss.getSheetByName("Checkup");
  if (!chkSheet) {
    chkSheet = ss.insertSheet("Checkup");
    chkSheet.appendRow([
      "รหัสพนักงาน", "คำนำหน้านาม", "ชื่อ", "นามสกุล", "Dept", "อายุ", 
      "การแพ้ยา", "โรคประจำตัว", "ตรวจร่างกายทั่วไปโดยแพทย์", "สูบบุหรี่", 
      "ดื่มแอลกอฮอล์", "ตาบอดสี", "น้ำหนัก", "ส่วนสูง", "BMI", "ผลแปร BMI", 
      "สายตาคอม", "CBC", "UA", "X-Ray", "Lipid", 
      "ตรวจปัสสาวะระดับน้ำตาลในเลือด", "ตรวจการทำงานของตับ", 
      "ตรวจการทำงานของไต", "ตรวจกรดยูริคในเลือด", "Year"
    ]);
    chkSheet.getRange("A1:Z1").setFontWeight("bold").setBackground("#c9daf8");
  }
  
  // 4. Inventory Sheet (Medicine Stock)
  var invSheet = ss.getSheetByName("Inventory");
  if (!invSheet) {
    invSheet = ss.insertSheet("Inventory");
    invSheet.appendRow(["ชื่อยา", "จำนวนคงเหลือ", "หน่วย", "แจ้งเตือนเมื่อเหลือ ≤"]);
    invSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#fff2cc");
    
    // Add 49 predefined medicines with default stock
    var defaultMeds = [
      "N.S.S. 450ml. 24's K&K.",
      "N.S.S. 100ml. 24's K&K.",
      "LORATADINE LORANOX 10mg. 20x10's S.CHAROEN",
      "MYDOCALM TANDERON 500's PHARMALAND",
      "ALCOHOL 70% 450ml. 24's LEOPARD",
      "ALCOHOL 70% 30ml.12's SIRIBUNCHA",
      "BELCID 240ml. 20's BIOLAB(ALUM MILK)",
      "BALM GOLDEN 2g. 12's",
      "BEMETH-N CREAM 450g. IMEX(BETAMETHASONE-N)",
      "BISOLVON BROMHEXINE 1000's NL.",
      "OP-IZE EYE WASH 110ml. 1x18bott ANB",
      "BRUFEN PROZONE 400mg.50x10's AU",
      "BUSCOPAN CENCOPAN 10mg.50x10's CPL.",
      "CALAMINE LOTION 60ml. 6's LEOPARD",
      "C.P.M.4mg. 1000's AU ฝานําเงิน",
      "COTTON BALL 450 GM",
      "COUNTERPAIN 120g. 24's TAISHO",
      "DRAMAMINE DIMEN 1000's AU.",
      "VOLTAREN DIFENAC 25mg.1000's AU.",
      "GLOVE POWDER Size M 100's PROXAM",
      "MOTILIUM MOTIDOM-M 50x10's TO.",
      "ELASTIC BANDAGE 3\" UTD",
      "ELASTIC BANDAGE 4\" UTD",
      "ORS.0reda3.3gรสส้ม100's",
      "EYE PAD Sterile 25's TG.",
      "AIR-X MINT 50x10's RX",
      "GAUZE Pad Non-Sterile 2\"x2\"-8Ply 10x10's China",
      "GAUZE Pad Non-Sterile 3\"x3\"-8Ply 10x10's China",
      "HISTA OPH Eye Drop 5ml. 12's SANG THAI",
      "KANOLONE CREAM 1g. 1x50's LBS",
      "M.CARMINATIVE 450ml. 12's GPO.",
      "MICROPORE 1\"x10Yds. 12's 3M",
      "PARA CEMOL 500mg.1000's CPL.",
      "PLASTIC BAG \"S\" 100's 6x8cm. LOCAL",
      "PLASTIC BOTTLE 60ml.100's LOCAL(ขายยกแพ็ค)",
      "PLASTIC CREAM BOX 5g.100's LOCAL(ขายยกแพ็ค)",
      "PONSTAN NEOPAIN 500mg. 500's NL.",
      "SILVERDERM CREAM 25g. 6's TO.",
      "TENSOPLAST 100's ผ้า PMC.(Smith ขายยกกลอง)",
      "TRANSPORE 1\"x10Yds.12's 3M",
      "YAHOM BAG 3g. 12's (5 JADEE)",
      "COTTON STICK \"M\" 10x100's SERVA",
      "M.TUSSIS 3800ml. SNAKE BRAND",
      "CONFORM 2\" 12's ZD.",
      "OMEPRAZOLE 20mg. 10x10's องค์การเภสัช",
      "AMMONIA 30ml. 12's วิทยาศรม",
      "BETADINE POLIDINE Solution 30ml. 12's NL.",
      "CA-R-BON 260mg.10x10's GREATER(ULTRACARBON)",
      "สเปรย์เย็น"
    ];
    
    var rows = [];
    for (var i = 0; i < defaultMeds.length; i++) {
      rows.push([defaultMeds[i], 500, "เม็ด", 50]);
    }
    invSheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
  return "Initialized sheets successfully.";
}

// Check Login Password
function checkLogin(password) {
  return password === "hananurse";
}

// Pad Numeric String to 6 Digits
function formatEmployeeId(id) {
  id = String(id).trim().toUpperCase();
  if (/^\d+$/.test(id)) {
    while (id.length < 6) {
      id = "0" + id;
    }
  }
  return id;
}

// Format Date object to d/m/yyyy string format
function formatDateDMY(dateVal) {
  if (dateVal instanceof Date) {
    return dateVal.getDate() + "/" + (dateVal.getMonth() + 1) + "/" + dateVal.getFullYear();
  }
  return String(dateVal).trim();
}

// Get Current Decimal Time format (e.g. 15.46)
function getCurrentTimeDecimal() {
  var now = new Date();
  var h = now.getHours();
  var m = now.getMinutes();
  m = m < 10 ? '0' + m : m;
  return parseFloat(h + "." + m);
}

// Locate History row using composite details
function findHistoryRowIndex(ss, empId, dateStr, timeVal, system) {
  var histSheet = ss.getSheetByName("History");
  var data = histSheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var rowEmpId = formatEmployeeId(data[i][0]);
    var rowDate = data[i][1];
    var rowTime = data[i][2];
    var rowSystem = data[i][8];
    
    var rowDateStr = formatDateDMY(rowDate);
    var rowTimeStr = String(rowTime).trim();
    
    if (rowEmpId === empId && rowDateStr === dateStr && rowTimeStr === String(timeVal).trim() && rowSystem === system) {
      return i + 1; // 1-indexed row number
    }
  }
  return -1;
}

// Update Congenital Disease and Drug Allergies inside Checkup sheet
function updateCheckupAllergiesAndDisease(empId, name, surname, dept, disease, allergy) {
  var ss = getSpreadsheet();
  var chkSheet = ss.getSheetByName("Checkup");
  var chkData = chkSheet.getDataRange().getValues();
  var currentYear = String(new Date().getFullYear());
  
  var rowToUpdateIndex = -1;
  var latestRowIndex = -1;
  var latestYear = 0;
  
  for (var i = 1; i < chkData.length; i++) {
    if (formatEmployeeId(chkData[i][0]) === empId) {
      var rowYear = parseInt(chkData[i][25], 10) || 0;
      if (String(chkData[i][25]) === currentYear) {
        rowToUpdateIndex = i + 1;
      }
      if (rowYear > latestYear) {
        latestYear = rowYear;
        latestRowIndex = i + 1;
      }
    }
  }
  
  if (rowToUpdateIndex !== -1) {
    // Update existing row for current year
    chkSheet.getRange(rowToUpdateIndex, 8).setValue(disease); // Column H (โรคประจำตัว)
    chkSheet.getRange(rowToUpdateIndex, 7).setValue(allergy); // Column G (การแพ้ยา)
  } else if (latestRowIndex !== -1) {
    // Clone latest row, update year and values, and append
    var latestRowValues = chkSheet.getRange(latestRowIndex, 1, 1, 26).getValues()[0];
    latestRowValues[25] = currentYear; // Year (Col Z)
    latestRowValues[7] = disease; // Disease (Col H)
    latestRowValues[6] = allergy; // Allergy (Col G)
    chkSheet.appendRow(latestRowValues);
  } else {
    // Create brand new row
    var newRow = new Array(26).fill("");
    newRow[0] = /^\d+$/.test(empId) ? "'" + empId : empId; // ID
    newRow[2] = name; // First name
    newRow[3] = surname; // Last name
    newRow[4] = dept; // Dept
    newRow[6] = allergy; // Allergy
    newRow[7] = disease; // Disease
    newRow[25] = currentYear; // Year
    chkSheet.appendRow(newRow);
  }
}

// Get Employee Data & Health Stats
function getEmployeeData(empId) {
  initializeSheets();
  empId = formatEmployeeId(empId);
  var ss = getSpreadsheet();
  
  var result = {
    found: false,
    empId: empId,
    name: "",
    surname: "",
    department: "",
    congenitalDisease: "ไม่ทราบข้อมูล",
    drugAllergies: "ไม่ทราบข้อมูล",
    historyStats: { "2024": 0, "2025": 0, "2026": 0 },
    latest3History: [],
    checkups: [],
    warning: false,
    warningMessage: ""
  };
  
  // 1. Check Name Sheet (Master)
  var empSheet = ss.getSheetByName("Name");
  var empData = empSheet.getDataRange().getValues();
  for (var i = 1; i < empData.length; i++) {
    if (formatEmployeeId(empData[i][0]) === empId) {
      result.found = true;
      result.name = empData[i][1] || "";
      result.surname = empData[i][2] || "";
      result.department = empData[i][3] || "";
      break;
    }
  }
  
  // 2. Check History stats and latest visits (Columns A, B, I, J, K, Q, R)
  var histSheet = ss.getSheetByName("History");
  var histData = histSheet.getDataRange().getValues();
  var matchVisits = [];
  
  for (var j = 1; j < histData.length; j++) {
    var rowEmpId = formatEmployeeId(histData[j][0]); // Col A
    if (rowEmpId === empId) {
      var dateVal = histData[j][1]; // Col B
      var yearVal = String(histData[j][16]); // Col Q (index 16)
      
      // Count visits per year
      if (result.historyStats[yearVal] !== undefined) {
        result.historyStats[yearVal]++;
      } else {
        result.historyStats[yearVal] = 1;
      }
      
      matchVisits.push({
        date: dateVal,
        system: histData[j][8], // Col I (index 8)
        symptom: histData[j][9] || histData[j][10] // Col J (อาการ) or Col K (DX)
      });
    }
  }
  
  // Sort visits by date descending to get the latest 3
  matchVisits.sort(function(a, b) {
    return new Date(b.date) - new Date(a.date);
  });
  result.latest3History = matchVisits.slice(0, 3);
  
  // 3. Check Checkups (Max 3 years)
  var chkSheet = ss.getSheetByName("Checkup");
  var chkData = chkSheet.getDataRange().getValues();
  var matchCheckups = [];
  
  var latestDisease = "";
  var latestAllergy = "";
  
  for (var k = 1; k < chkData.length; k++) {
    var chkEmpId = formatEmployeeId(chkData[k][0]);
    if (chkEmpId === empId) {
      if (chkData[k][7] && chkData[k][7] !== "ไม่ทราบข้อมูล") latestDisease = chkData[k][7];
      if (chkData[k][6] && chkData[k][6] !== "ไม่ทราบข้อมูล") latestAllergy = chkData[k][6];
      
      matchCheckups.push({
        allergy: chkData[k][6] || "ไม่ทราบข้อมูล",
        disease: chkData[k][7] || "ไม่ทราบข้อมูล",
        physical: chkData[k][8] || "ไม่ทราบข้อมูล",
        smoke: chkData[k][9] || "ไม่ทราบข้อมูล",
        drink: chkData[k][10] || "ไม่ทราบข้อมูล",
        colorBlind: chkData[k][11] || "ไม่ทราบข้อมูล",
        weight: chkData[k][12] || "-",
        height: chkData[k][13] || "-",
        bmi: chkData[k][14] || "-",
        bmiInterpretation: chkData[k][15] || "-",
        eyeComputer: chkData[k][16] || "ไม่ทราบข้อมูล",
        cbc: chkData[k][17] || "-",
        ua: chkData[k][18] || "-",
        xray: chkData[k][19] || "-",
        lipid: chkData[k][20] || "-",
        sugarUA: chkData[k][21] || "ไม่ทราบข้อมูล",
        liver: chkData[k][22] || "ไม่ทราบข้อมูล",
        kidney: chkData[k][23] || "ไม่ทราบข้อมูล",
        uric: chkData[k][24] || "ไม่ทราบข้อมูล",
        year: chkData[k][25] || "-"
      });
    }
  }
  
  // Sort checkups by year descending and select top 3
  matchCheckups.sort(function(a, b) {
    return b.year - a.year;
  });
  result.checkups = matchCheckups.slice(0, 3);
  
  result.congenitalDisease = latestDisease || "ไม่ทราบข้อมูล";
  result.drugAllergies = latestAllergy || "ไม่ทราบข้อมูล";
  
  // Calculate total history records for warning
  var totalHistoryRecords = matchVisits.length;
  if (totalHistoryRecords >= 3) {
    result.warning = true;
    result.warningMessage = "พนักงานรายนี้มารับบริการห้องพยาบาลแล้ว " + totalHistoryRecords + " ครั้ง (รวมปี 2024 = " + (result.historyStats["2024"] || 0) + " ครั้ง, ปี 2025 = " + (result.historyStats["2025"] || 0) + " ครั้ง, ปี 2026 = " + (result.historyStats["2026"] || 0) + " ครั้ง)";
  }
  
  return result;
}

// Parse medication string format like "PARA CEMOL 500mg.1000's CPL. (4), AMMONIA 30ml. 12's (2)"
function parseMedications(medString) {
  var meds = [];
  if (!medString) return meds;
  var items = medString.split(/,\s+/);
  for (var i = 0; i < items.length; i++) {
    var match = items[i].match(/(.+)\s+\((\d+)\)$/);
    if (match) {
      meds.push({
        name: match[1].trim(),
        qty: parseInt(match[2], 10)
      });
    }
  }
  return meds;
}

// Adjust Medication Stock Quantities in Inventory sheet
function adjustStock(medsList, isDeduct) {
  var ss = getSpreadsheet();
  var invSheet = ss.getSheetByName("Inventory");
  if (!invSheet) return;
  
  var medData = invSheet.getDataRange().getValues();
  
  for (var i = 0; i < medsList.length; i++) {
    var nameToMatch = medsList[i].name;
    var qtyToAdjust = medsList[i].qty;
    
    for (var j = 1; j < medData.length; j++) {
      if (medData[j][0] === nameToMatch) {
        var currentStock = parseFloat(medData[j][1]) || 0;
        var newStock = isDeduct ? (currentStock - qtyToAdjust) : (currentStock + qtyToAdjust);
        invSheet.getRange(j + 1, 2).setValue(newStock); // update Column B
        break;
      }
    }
  }
}

// Save Visit Record
function saveVisitRecord(data) {
  initializeSheets();
  var ss = getSpreadsheet();
  var histSheet = ss.getSheetByName("History");
  
  var timestamp = new Date();
  var dateStr = timestamp.getDate() + "/" + (timestamp.getMonth() + 1) + "/" + timestamp.getFullYear();
  var timeDec = getCurrentTimeDecimal();
  var monthVal = timestamp.getMonth() + 1;
  var yearStr = String(timestamp.getFullYear());
  
  var empId = formatEmployeeId(data.empId);
  var isEmployee = !(/^A[1-6]$/.test(empId));
  
  // If it's an employee and name is entered manually (new employee), add to Name sheet
  if (isEmployee && data.isNewEmployee && data.name && data.surname) {
    var empSheet = ss.getSheetByName("Name");
    var formattedId = /^\d+$/.test(empId) ? "'" + empId : empId;
    empSheet.appendRow([formattedId, data.name, data.surname, data.department]);
  }
  
  // Format medications string
  var medString = "";
  if (data.medications && data.medications.length > 0) {
    var medParts = [];
    for (var i = 0; i < data.medications.length; i++) {
      medParts.push(data.medications[i].name + " (" + data.medications[i].qty + ")");
    }
    medString = medParts.join(", ");
  }
  
  // Map flags
  var isAccident = (data.system === "อุบัติเหตุจากการทำงาน") ? "Yes" : "No";
  var isRest = (data.status === "Rest") ? "Yes" : "No";
  var isRefer = (data.status === "Refer") ? (data.referHospital || "Yes") : "No";
  var isBTW = (data.status === "BTW") ? "Yes" : "No";
  
  // Update congenital disease & allergies in Checkup
  updateCheckupAllergiesAndDisease(empId, data.name, data.surname, data.department || "", data.congenitalDisease, data.drugAllergies);
  
  // Append History Row (Columns A to T)
  var formattedIdForHist = /^\d+$/.test(empId) ? "'" + empId : empId;
  histSheet.appendRow([
    formattedIdForHist,                 // A: รหัสพนักงาน
    dateStr,                            // B: Date
    timeDec,                            // C: Time
    monthVal,                           // D: Month
    data.name,                          // E: First Name
    data.surname,                       // F: Last Name
    data.shift || "d",                  // G: Shift
    data.department || "",              // H: Department
    data.system,                        // I: ระบบโรค
    data.cc || "",                      // J: อาการ
    data.dx || "",                      // K: DX
    data.fromWork || "No",              // L: From Work
    isBTW,                              // M: Back To Work
    isAccident,                         // N: Accident
    isRest,                             // O: Rest
    isRefer,                            // P: Refer (Hospital Name)
    yearStr,                            // Q: Year
    medString,                          // R: จ่ายยา (Medication)
    data.bedStartTime || "",            // S: Bed Start (เวลาเริ่ม)
    data.bedEndTime || ""               // T: Bed End (เวลาตื่น)
  ]);
  
  // Deduct Stock
  if (data.medications && data.medications.length > 0) {
    adjustStock(data.medications, true);
  }
  
  return { success: true, message: "บันทึกข้อมูลการรับบริการสำเร็จ" };
}

// Update Bed End Time in History using Composite ID
function updateBedEndTime(compositeId, endTime) {
  var ss = getSpreadsheet();
  var histSheet = ss.getSheetByName("History");
  
  var idParts = compositeId.split("|");
  var empId = idParts[0];
  var dateStr = idParts[1];
  var timeStr = idParts[2];
  var system = idParts[3];
  
  var rowIndex = findHistoryRowIndex(ss, empId, dateStr, timeStr, system);
  if (rowIndex === -1) {
    return { success: false, message: "ไม่พบประวัติการรับบริการในระบบ" };
  }
  
  histSheet.getRange(rowIndex, 20).setValue(endTime); // Column T: Bed End (เวลาตื่น)
  return { success: true, message: "บันทึกเวลาสิ้นสุดการนอนเรียบร้อย" };
}

// Get Service History for Employee
function getServiceHistory(empId) {
  initializeSheets();
  empId = formatEmployeeId(empId);
  var ss = getSpreadsheet();
  var histSheet = ss.getSheetByName("History");
  var histData = histSheet.getDataRange().getValues();
  
  // Look up latest allergies and disease from Checkup
  var checkupSheet = ss.getSheetByName("Checkup");
  var checkupData = checkupSheet.getDataRange().getValues();
  var latestDisease = "ไม่ทราบข้อมูล";
  var latestAllergy = "ไม่ทราบข้อมูล";
  
  for (var k = 1; k < checkupData.length; k++) {
    if (formatEmployeeId(checkupData[k][0]) === empId) {
      if (checkupData[k][7]) latestDisease = checkupData[k][7];
      if (checkupData[k][6]) latestAllergy = checkupData[k][6];
    }
  }
  
  var results = [];
  for (var i = 1; i < histData.length; i++) {
    if (formatEmployeeId(histData[i][0]) === empId) {
      var dStr = formatDateDMY(histData[i][1]);
      var tStr = String(histData[i][2]).trim();
      var sys = histData[i][8];
      
      var compositeId = empId + "|" + dStr + "|" + tStr + "|" + sys;
      
      results.push({
        id: compositeId,
        date: histData[i][1],
        dateStr: dStr,
        time: histData[i][2],
        month: histData[i][3],
        empId: histData[i][0],
        name: histData[i][4],
        surname: histData[i][5],
        shift: histData[i][6],
        department: histData[i][7],
        congenitalDisease: latestDisease,
        drugAllergies: latestAllergy,
        system: histData[i][8],
        cc: histData[i][9],
        dx: histData[i][10],
        fromWork: histData[i][11],
        btw: histData[i][12],
        acc: histData[i][13],
        rest: histData[i][14],
        refer: histData[i][15],
        year: histData[i][16],
        medication: histData[i][17],
        bedStartTime: histData[i][18],
        bedEndTime: histData[i][19],
        status: histData[i][14] === "Yes" ? "Rest" : (histData[i][15] !== "No" ? "Refer" : "BTW")
      });
    }
  }
  
  // Sort descending
  results.sort(function(a, b) {
    return new Date(b.date) - new Date(a.date);
  });
  
  return results;
}

// Edit Existing History Record
function editHistoryRecord(data) {
  var ss = getSpreadsheet();
  var histSheet = ss.getSheetByName("History");
  
  var idParts = data.id.split("|");
  var empId = idParts[0];
  var dateStr = idParts[1];
  var timeStr = idParts[2];
  var system = idParts[3];
  
  var rowIndex = findHistoryRowIndex(ss, empId, dateStr, timeStr, system);
  if (rowIndex === -1) {
    return { success: false, message: "ไม่พบประวัติการรักษาเพื่อทำการแก้ไข" };
  }
  
  // 1. Restore old medicine stock
  var oldRow = histSheet.getRange(rowIndex, 1, 1, 20).getValues()[0];
  var oldMedString = oldRow[17];
  var oldMeds = parseMedications(oldMedString);
  if (oldMeds.length > 0) {
    adjustStock(oldMeds, false);
  }
  
  // 2. Format new medicine string
  var newMedString = "";
  if (data.medications && data.medications.length > 0) {
    var medParts = [];
    for (var j = 0; j < data.medications.length; j++) {
      medParts.push(data.medications[j].name + " (" + data.medications[j].qty + ")");
    }
    newMedString = medParts.join(", ");
  }
  
  var isAccident = (data.system === "อุบัติเหตุจากการทำงาน") ? "Yes" : "No";
  var isRest = (data.status === "Rest") ? "Yes" : "No";
  var isRefer = (data.status === "Refer") ? (data.referHospital || "Yes") : "No";
  var isBTW = (data.status === "BTW") ? "Yes" : "No";
  
  // Update Congenital disease & allergies in Checkup
  updateCheckupAllergiesAndDisease(empId, data.name, data.surname, data.department || "", data.congenitalDisease, data.drugAllergies);
  
  // 3. Update History cells
  histSheet.getRange(rowIndex, 5).setValue(data.name);
  histSheet.getRange(rowIndex, 6).setValue(data.surname);
  histSheet.getRange(rowIndex, 7).setValue(data.shift || "d");
  histSheet.getRange(rowIndex, 8).setValue(data.department || "");
  histSheet.getRange(rowIndex, 9).setValue(data.system);
  histSheet.getRange(rowIndex, 10).setValue(data.cc || "");
  histSheet.getRange(rowIndex, 11).setValue(data.dx || "");
  histSheet.getRange(rowIndex, 12).setValue(data.fromWork || "No");
  histSheet.getRange(rowIndex, 13).setValue(isBTW);
  histSheet.getRange(rowIndex, 14).setValue(isAccident);
  histSheet.getRange(rowIndex, 15).setValue(isRest);
  histSheet.getRange(rowIndex, 16).setValue(isRefer);
  histSheet.getRange(rowIndex, 18).setValue(newMedString);
  histSheet.getRange(rowIndex, 19).setValue(data.bedStartTime || "");
  histSheet.getRange(rowIndex, 20).setValue(data.bedEndTime || "");
  
  // 4. Deduct new stock
  if (data.medications && data.medications.length > 0) {
    adjustStock(data.medications, true);
  }
  
  return { success: true, message: "แก้ไขข้อมูลประวัติการรักษาสำเร็จ" };
}

// Delete History Record
function deleteHistoryRecord(compositeId) {
  var ss = getSpreadsheet();
  var histSheet = ss.getSheetByName("History");
  
  var idParts = compositeId.split("|");
  var empId = idParts[0];
  var dateStr = idParts[1];
  var timeStr = idParts[2];
  var system = idParts[3];
  
  var rowIndex = findHistoryRowIndex(ss, empId, dateStr, timeStr, system);
  if (rowIndex === -1) {
    return { success: false, message: "ไม่พบข้อมูลประวัติการรับบริการในระบบ" };
  }
  
  // 1. Restore old medicine stock
  var oldRow = histSheet.getRange(rowIndex, 1, 1, 20).getValues()[0];
  var oldMedString = oldRow[17];
  var oldMeds = parseMedications(oldMedString);
  if (oldMeds.length > 0) {
    adjustStock(oldMeds, false);
  }
  
  // 2. Delete the row
  histSheet.deleteRow(rowIndex);
  return { success: true, message: "ลบประวัติการรับบริการและคืนสต๊อกยาเรียบร้อย" };
}

// Helper to check if a date falls in the specified range
function isInDateRange(dateVal, type) {
  var d = new Date(dateVal);
  if (isNaN(d.getTime())) return false;
  
  var now = new Date();
  if (type === "weekly") {
    var diffTime = Math.abs(now - d);
    var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    return diffDays <= 7;
  } else if (type === "monthly") {
    return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
  } else if (type === "yearly") {
    return d.getFullYear() === now.getFullYear();
  }
  return true;
}

// Get Dashboard Aggregated Data
function getDashboardData(filterType) {
  initializeSheets();
  var ss = getSpreadsheet();
  var histSheet = ss.getSheetByName("History");
  var data = histSheet.getDataRange().getValues();
  
  var systemCounts = {};
  var empCounts = {};
  var deptCounts = {};
  var deptSystemMap = {};
  
  for (var i = 1; i < data.length; i++) {
    var dateVal = data[i][1];
    if (!isInDateRange(dateVal, filterType)) continue;
    
    var system = data[i][8]; // Col I (index 8)
    var empId = formatEmployeeId(data[i][0]); // Col A (index 0)
    var dept = data[i][7] || "ไม่ระบุแผนก"; // Col H (index 7)
    
    if (system) {
      systemCounts[system] = (systemCounts[system] || 0) + 1;
    }
    if (empId) {
      empCounts[empId] = (empCounts[empId] || 0) + 1;
    }
    if (dept) {
      deptCounts[dept] = (deptCounts[dept] || 0) + 1;
      if (!deptSystemMap[dept]) {
        deptSystemMap[dept] = {};
      }
      if (system) {
        deptSystemMap[dept][system] = (deptSystemMap[dept][system] || 0) + 1;
      }
    }
  }
  
  var sortedSystems = Object.keys(systemCounts).map(function(key) {
    return { name: key, count: systemCounts[key] };
  }).sort(function(a, b) { return b.count - a.count; }).slice(0, 5);
  
  var sortedEmployees = Object.keys(empCounts).map(function(key) {
    return { empId: key, count: empCounts[key] };
  }).sort(function(a, b) { return b.count - a.count; }).slice(0, 10);
  
  var sortedDepts = Object.keys(deptCounts).map(function(key) {
    var sysMap = deptSystemMap[key] || {};
    var topSys = "";
    var topSysCount = 0;
    Object.keys(sysMap).forEach(function(s) {
      if (sysMap[s] > topSysCount) {
        topSys = s;
        topSysCount = sysMap[s];
      }
    });
    return { 
      name: key, 
      count: deptCounts[key],
      topSystem: topSys || "ไม่มีข้อมูล",
      topSystemCount: topSysCount
    };
  }).sort(function(a, b) { return b.count - a.count; }).slice(0, 5);
  
  return {
    systems: sortedSystems,
    employees: sortedEmployees,
    departments: sortedDepts
  };
}

// Get Filtered Report Data
function getReportData(filters) {
  initializeSheets();
  var ss = getSpreadsheet();
  var histSheet = ss.getSheetByName("History");
  var data = histSheet.getDataRange().getValues();
  
  var results = [];
  
  for (var i = 1; i < data.length; i++) {
    var dateVal = data[i][1];
    var d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    
    // Parse filters
    var yearMatch = true;
    if (filters.year) {
      yearMatch = String(data[i][16]) === String(filters.year); // Col Q (index 16)
    }
    
    var systemMatch = true;
    if (filters.system) {
      systemMatch = data[i][8] === filters.system; // Col I (index 8)
    }
    
    var deptMatch = true;
    if (filters.dept) {
      deptMatch = String(data[i][7]).toLowerCase().indexOf(filters.dept.toLowerCase()) !== -1;
    }
    
    var monthMatch = true;
    if (filters.month && !isNaN(d.getTime())) {
      monthMatch = String(d.getMonth() + 1) === String(filters.month);
    }
    
    var dayMatch = true;
    if (filters.day && !isNaN(d.getTime())) {
      dayMatch = String(d.getDate()) === String(filters.day);
    }
    
    if (yearMatch && systemMatch && deptMatch && monthMatch && dayMatch) {
      results.push({
        date: formatDateDMY(data[i][1]),
        empId: data[i][0],
        name: data[i][4],
        surname: data[i][5],
        shift: data[i][6],
        department: data[i][7],
        symptom: data[i][9] || "",
        system: data[i][8],
        btw: data[i][12] === "Yes" ? "Yes" : "No",
        refer: data[i][15] !== "No" ? data[i][15] : "No", // Show hospital or No
        rest: data[i][14] === "Yes" ? (data[i][18] + " - " + (data[i][19] || "กำลังนอน")) : "No",
        acc: data[i][13] || "No"
      });
    }
  }
  
  return results;
}

// Get Medicine Stock List & Calculate Total Dispensed from Inventory sheet
function getMedicineStock() {
  initializeSheets();
  var ss = getSpreadsheet();
  var invSheet = ss.getSheetByName("Inventory");
  var medData = invSheet.getDataRange().getValues();
  
  var histSheet = ss.getSheetByName("History");
  var histData = histSheet.getDataRange().getValues();
  
  // Aggregate dispensed quantities from all history
  var dispensedCounts = {};
  for (var i = 1; i < histData.length; i++) {
    var medString = histData[i][17]; // Col R (index 17)
    if (medString) {
      var meds = parseMedications(medString);
      for (var k = 0; k < meds.length; k++) {
        dispensedCounts[meds[k].name] = (dispensedCounts[meds[k].name] || 0) + meds[k].qty;
      }
    }
  }
  
  var list = [];
  for (var j = 1; j < medData.length; j++) {
    var name = medData[j][0];
    if (!name) continue;
    var current = parseFloat(medData[j][1]) || 0;
    var unit = medData[j][2] || "เม็ด";
    var threshold = parseFloat(medData[j][3]) || 0;
    var totalDispensed = dispensedCounts[name] || 0;
    
    list.push({
      name: name,
      current: current,
      unit: unit,
      threshold: threshold,
      dispensed: totalDispensed,
      isLow: current <= threshold
    });
  }
  
  return list;
}

// Add Medicine Stock in Inventory
function addMedicineStock(medName, qty) {
  var ss = getSpreadsheet();
  var invSheet = ss.getSheetByName("Inventory");
  var data = invSheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === medName) {
      var currentVal = parseFloat(data[i][1]) || 0;
      invSheet.getRange(i + 1, 2).setValue(currentVal + qty); // update Column B
      return { success: true, message: "เพิ่มยา " + medName + " จำนวน " + qty + " ชิ้น เข้าคลังสำเร็จ" };
    }
  }
  return { success: false, message: "ไม่พบรายการยา" };
}

// Get Monthly/Yearly Dispensed Breakdown for a Medicine
function getMedicationHistoryData(medName, filterType) {
  var ss = getSpreadsheet();
  var histSheet = ss.getSheetByName("History");
  var histData = histSheet.getDataRange().getValues();
  
  var agg = {};
  
  for (var i = 1; i < histData.length; i++) {
    var dateVal = histData[i][1];
    var d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(d.getTime())) continue;
    
    var medString = histData[i][17]; // Col R (index 17)
    if (medString) {
      var meds = parseMedications(medString);
      for (var k = 0; k < meds.length; k++) {
        if (meds[k].name === medName) {
          var key = "";
          if (filterType === "monthly") {
            var month = String(d.getMonth() + 1);
            if (month.length < 2) month = "0" + month;
            key = d.getFullYear() + "-" + month;
          } else {
            key = String(d.getFullYear());
          }
          agg[key] = (agg[key] || 0) + meds[k].qty;
        }
      }
    }
  }
  
  var sortedKeys = Object.keys(agg).sort();
  var results = sortedKeys.map(function(key) {
    return { label: key, qty: agg[key] };
  });
  
  return results;
}
