// ============================================================================
// GSEA ENGLISH LEARNING PORTFOLIO — APPS SCRIPT BACKEND
// Spreadsheet: GSEA_English_Portfolio_DB
// Supports both GET & POST for all operations (Login, Sync, Submission, Diagnostics)
// Wraps ALL handlers in top-level try...catch returning clean JSON (Never HTML)
// ============================================================================

var SPREADSHEET_ID = "1zNagUBMjubKsglB1BGLnfqwUSO09nn4iOo5VMcVqLD4";

/**
 * Diagnostic test you can run from Apps Script Editor (Click Run > testConnection)
 */
function testConnection() {
  var ss = getSpreadsheet();
  var studentsSheet = getStudentsSheet(ss);
  var logsSheet = getLogsSheet(ss);

  Logger.log("=== GSEA DATABASE DIAGNOSTIC ===");
  Logger.log("Spreadsheet: " + (ss ? ss.getName() : "NOT ACCESSIBLE"));
  Logger.log("Students tab: " + (studentsSheet ? "FOUND '" + studentsSheet.getName() + "' (" + studentsSheet.getLastRow() + " rows)" : "NOT FOUND"));
  Logger.log("Logs tab: " + (logsSheet ? "FOUND '" + logsSheet.getName() + "' (" + logsSheet.getLastRow() + " rows)" : "NOT FOUND"));

  if (studentsSheet && studentsSheet.getLastRow() > 1) {
    var sampleStudent = studentsSheet.getRange(2, 1).getValue();
    Logger.log("Sample Student in DB: " + sampleStudent);
  }
}

/**
 * Backwards compatibility alias for older script calls
 */
function loginUser(student, pin) {
  return handleAction({ action: "login", student: student, pin: pin });
}

/**
 * POST handler - Wraps everything in top-level try/catch, returns strict JSON
 */
function doPost(e) {
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    var result = handleAction(data);
    return jsonOut(result);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: err.toString(),
        message: "Server Error: " + err.message,
        stack: err.stack ? String(err.stack) : ""
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GET handler - Wraps everything in top-level try/catch, returns strict JSON
 */
function doGet(e) {
  try {
    var data = (e && e.parameter) ? e.parameter : {};
    var result = handleAction(data);
    return jsonOut(result);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: err.toString(),
        message: "Server Error: " + err.message,
        stack: err.stack ? String(err.stack) : ""
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Central Router: handles login, submission, and diagnostic ping identically for GET and POST
 */
function handleAction(data) {
  data = data || {};
  var action = (data.action || "").toString().toLowerCase().trim();

  // Submission detection: explicit action 'submit' OR payload containing session entry fields
  var isSubmit = action === "submit" || (data.vocab !== undefined || data.contentReview !== undefined || data.workOn !== undefined);
  
  // Login detection: explicit action 'login' OR presence of student and pin (without submission content)
  var isLogin = action === "login" || (data.student && data.pin && !isSubmit);
  // Tutor actions
  var isTutorLogin = action === "tutorlogin";
  var isTutorAssessment = action === "save_assessment" || action === "tutorsubmit" || (data.tutor && data.mark !== undefined);

  if (isTutorAssessment) {
    return handleSaveAssessment(data);
  } else if (isSubmit) {
    return handleSubmit(data);
  } else if (isTutorLogin) {
    return handleTutorLogin(data.student || data.tutor, data.pin);
  } else if (isLogin) {
    return handleLogin(data.student, data.pin);
  } else if (action === "ping" || (!data.student && !data.pin)) {
    return handlePing();
  } else {
    if (data.student && !data.pin) {
      return { success: false, message: "Please provide both student name and 4-digit PIN." };
    }
    return { success: false, message: "Unrecognized request action: '" + action + "'" };
  }
}

/**
 * Handles Tutor Authentication via Students Sheet Tab
 */
function handleTutorLogin(tutorName, tutorPin) {
  var pin = String(tutorPin || "").trim();
  var name = String(tutorName || "").trim();
  if (!pin) {
    return { success: false, message: "Please provide the Tutor password/PIN." };
  }

  var ss = getSpreadsheet();
  var studentsSheet = getStudentsSheet(ss);
  if (!studentsSheet) {
    return { success: false, message: "Database Error: 'Students' tab not found in spreadsheet." };
  }

  var studentsData = studentsSheet.getDataRange().getValues();
  var authorized = false;
  var matchedTutor = "";

  for (var i = 1; i < studentsData.length; i++) {
    var row = studentsData[i];
    var sName = String(row[0] || "").trim();
    var sPin = String(row[1] != null ? row[1] : "").trim().replace(/\.0+$/, "");
    if (/^\d{1,3}$/.test(sPin) && pin.length === 4) sPin = sPin.padStart(4, "0");

    if (name && sName.toLowerCase() === name.toLowerCase()) {
      if (sPin === pin) {
        authorized = true;
        matchedTutor = sName;
        break;
      }
    } else if (!name || sName.toLowerCase().indexOf("tutor") !== -1) {
      if (sPin === pin) {
        authorized = true;
        matchedTutor = sName;
        break;
      }
    }
  }

  if (authorized) {
    var studentList = [];
    for (var j = 1; j < studentsData.length; j++) {
      var nameVal = String(studentsData[j][0] || "").trim();
      if (nameVal && nameVal.toLowerCase().indexOf("tutor") === -1) {
        studentList.push(nameVal);
      }
    }

    var assessmentsMap = {};
    var assessSheet = ss.getSheetByName("Assessments") || ss.getSheetByName("TutorAssessments");
    if (assessSheet && assessSheet.getLastRow() > 1) {
      var aData = assessSheet.getDataRange().getValues();
      for (var a = 1; a < aData.length; a++) {
        var aStudent = String(aData[a][2] || "").trim();
        if (aStudent) {
          assessmentsMap[aStudent] = {
            timestamp: aData[a][0] ? String(aData[a][0]) : "",
            tutor: String(aData[a][1] || ""),
            student: aStudent,
            mark: aData[a][3] !== undefined ? String(aData[a][3]) : "",
            decision: String(aData[a][4] || ""),
            remarks: String(aData[a][5] || "")
          };
        }
      }
    }

    return {
      success: true,
      message: "Tutor authenticated successfully.",
      tutor: matchedTutor || name || "Tutor",
      students: studentList,
      assessments: assessmentsMap
    };
  } else {
    return {
      success: false,
      message: "Invalid Tutor password/PIN. Please check the 'Students' sheet."
    };
  }
}

/**
 * Handles Saving Tutor Assessment (Mark X/20 and GO/NO GO)
 */
function handleSaveAssessment(data) {
  var tutor = String(data.tutor || "Tutor").trim();
  var student = String(data.student || "").trim();
  var mark = data.mark !== undefined ? String(data.mark).trim() : "";
  var decision = String(data.decision || "GO").trim().toUpperCase();
  var remarks = String(data.remarks || "").trim();

  if (!student) {
    return { success: false, message: "Student name is required for assessment." };
  }

  var lock = LockService.getScriptLock();
  var acquired = false;
  try {
    acquired = lock.tryLock(20000);
  } catch (e) {}

  if (!acquired) {
    return { success: false, message: "Server busy recording another entry. Please try again." };
  }

  try {
    var ss = getSpreadsheet();
    
    // 1. Record or update in Assessments tab
    var assessSheet = ss.getSheetByName("Assessments") || ss.getSheetByName("TutorAssessments");
    if (!assessSheet) {
      assessSheet = ss.insertSheet("Assessments");
      assessSheet.appendRow(["Timestamp", "Tutor", "Student", "Mark", "Decision", "Remarks"]);
    }
    
    var aData = assessSheet.getDataRange().getValues();
    var existingRow = -1;
    for (var i = 1; i < aData.length; i++) {
      if (String(aData[i][2] || "").trim().toLowerCase() === student.toLowerCase()) {
        existingRow = i + 1;
        break;
      }
    }
    
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Europe/Paris", "yyyy-MM-dd HH:mm");
    if (existingRow > 0) {
      assessSheet.getRange(existingRow, 1, 1, 6).setValues([[nowStr, tutor, student, mark, decision, remarks]]);
    } else {
      assessSheet.appendRow([nowStr, tutor, student, mark, decision, remarks]);
    }

    // 2. Also log in logsSheet as a Module 4 entry so student logs reflect it
    var logsSheet = getLogsSheet(ss);
    if (logsSheet) {
      logsSheet.appendRow([
        new Date(),
        student,
        "TUTOR",
        "Module 4 (Final Assesment)",
        decision,
        "Decision: " + decision,
        "Final Assessment Mark: " + mark + "/20",
        remarks,
        "Assessed by " + tutor,
        mark,
        "",
        ""
      ]);
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      message: "Tutor assessment recorded successfully for " + student + ".",
      assessment: {
        student: student,
        tutor: tutor,
        mark: mark,
        decision: decision,
        remarks: remarks,
        timestamp: nowStr
      }
    };
  } catch (err) {
    return { success: false, message: "Error saving assessment: " + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * Handles Student PIN Authentication & Log Retrieval
 */
function handleLogin(studentName, studentPin) {
  var student = String(studentName || "").trim();
  var pin = String(studentPin || "").trim();

  if (!student || !pin) {
    return { success: false, message: "Please select a student and provide a 4-digit PIN." };
  }

  var ss = getSpreadsheet();

  // 1. PIN Authentication via Students sheet
  var studentsSheet = getStudentsSheet(ss);
  if (!studentsSheet) {
    var tabNames = ss.getSheets().map(function(s) { return "'" + s.getName() + "'"; }).join(", ");
    return {
      success: false,
      message: "Database Error: 'Students' tab not found in spreadsheet. Available tabs: [" + tabNames + "]"
    };
  }

  var studentsData = studentsSheet.getDataRange().getValues();
  if (studentsData.length <= 1) {
    return { success: false, message: "The 'Students' sheet has no student records registered." };
  }

  var authorized = false;
  var found = false;

  for (var i = 1; i < studentsData.length; i++) {
    var row = studentsData[i];
    var sName = String(row[0] || "").trim();
    var sPin = String(row[1] != null ? row[1] : "").trim().replace(/\.0+$/, "");
    if (/^\d{1,3}$/.test(sPin) && pin.length === 4) sPin = sPin.padStart(4, "0");

    if (sName.toLowerCase() === student.toLowerCase()) {
      found = true;
      if (sPin === pin) {
        authorized = true;
      }
      break;
    }
  }

  if (!found) {
    return { success: false, message: "Student '" + student + "' not found in registered students list." };
  }
  if (!authorized) {
    return { success: false, message: "Incorrect PIN for " + student + ". Please verify your 4-digit code." };
  }

  // 2. Fetch Form Responses / PortfolioLogs
  var logsSheet = getLogsSheet(ss);
  if (!logsSheet) {
    var tabNames = ss.getSheets().map(function(s) { return "'" + s.getName() + "'"; }).join(", ");
    return {
      success: false,
      message: "Database Error: Responses tab ('Form Responses' or 'PortfolioLogs') not found. Available tabs: [" + tabNames + "]"
    };
  }

  var data = logsSheet.getDataRange().getValues();
  var logs = [];

  if (data.length > 1) {
    var rawHeaders = data[0].map(function(item) { return String(item || "").trim().toLowerCase(); });

    // Determine if column 2 is PIN
    var hasPinCol = false;
    for (var hIdx = 0; hIdx < rawHeaders.length; hIdx++) {
      if (rawHeaders[hIdx].indexOf("pin") !== -1) {
        hasPinCol = true;
        break;
      }
    }
    if (!hasPinCol && data.length > 1) {
      var sampleCol2 = String(data[1][2] || "").trim();
      if (/^\d{4}$/.test(sampleCol2)) {
        hasPinCol = true;
      }
    }

    var cTs  = findCol(rawHeaders, ["time", "date", "horodateur"], 0);
    var cSt  = findCol(rawHeaders, ["student", "nom", "name"], 1);
    var cMod = findModuleCol(rawHeaders, hasPinCol ? 3 : 2);
    var cSes = findCol(rawHeaders, ["session"], hasPinCol ? 4 : 3);
    var cVoc = findCol(rawHeaders, ["vocab", "word", "lexique"], hasPinCol ? 5 : 4);
    var cRev = findCol(rawHeaders, ["review", "learnt", "lesson content", "content"], hasPinCol ? 6 : 5);
    var cWrk = findCol(rawHeaders, ["work on", "need to work", "following this session", "work"], hasPinCol ? 7 : 6);
    var cEvH = findEvalHelpCol(rawHeaders, hasPinCol ? 8 : 7);
    var cEvR = findEvalRatingCol(rawHeaders, hasPinCol ? 9 : 8);
    var cToP = findToeicPrepCol(rawHeaders, hasPinCol ? 10 : 9);
    var cToR = findToeicRatingCol(rawHeaders, hasPinCol ? 11 : 10);

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var rowStudent = String(row[cSt] || "").trim();

      if (rowStudent.toLowerCase() === student.toLowerCase()) {
        var ts = row[cTs] ? String(row[cTs]) : "";
        if (row[cTs] instanceof Date) {
          ts = Utilities.formatDate(row[cTs], Session.getScriptTimeZone() || "Europe/Paris", "yyyy-MM-dd HH:mm");
        }

        var modVal = String(row[cMod] || "").trim();
        var sesVal = String(row[cSes] || "").trim();

        // Self-healing: if modVal captured a PIN and sesVal contains "Module", fix column offset
        if (/^\d{4}$/.test(modVal) && /module/i.test(sesVal)) {
          modVal = sesVal;
          sesVal = String(row[cSes + 1] || "").trim();
        }

        logs.push({
          timestamp: ts,
          module: modVal,
          session: sesVal,
          vocab: String(row[cVoc] || ""),
          contentReview: String(row[cRev] || ""),
          workOn: String(row[cWrk] || ""),
          evalHelp: String(row[cEvH] || ""),
          evalRating: String(row[cEvR] || ""),
          toeicPrep: String(row[cToP] || ""),
          toeicRating: String(row[cToR] || "")
        });
      }
    }
  }

  // Look for tutor assessment for this student
  var studentAssessment = null;
  var assessSheet = ss.getSheetByName("Assessments") || ss.getSheetByName("TutorAssessments");
  if (assessSheet && assessSheet.getLastRow() > 1) {
    var aData = assessSheet.getDataRange().getValues();
    for (var a = 1; a < aData.length; a++) {
      if (String(aData[a][2] || "").trim().toLowerCase() === student.toLowerCase()) {
        studentAssessment = {
          timestamp: aData[a][0] ? String(aData[a][0]) : "",
          tutor: String(aData[a][1] || ""),
          student: String(aData[a][2] || ""),
          mark: aData[a][3] !== undefined ? String(aData[a][3]) : "",
          decision: String(aData[a][4] || ""),
          remarks: String(aData[a][5] || "")
        };
        break;
      }
    }
  }

  return {
    success: true,
    student: student,
    logs: logs,
    assessment: studentAssessment,
    serverTime: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Europe/Paris", "yyyy-MM-dd HH:mm:ss")
  };
}

/**
 * Handles Portfolio Entry Submission
 * Concurrency Protection: Uses LockService.getScriptLock() with a 30-second wait
 * to prevent simultaneous student writes from colliding, overwriting rows, or timing out.
 */
function handleSubmit(d) {
  var student = String(d.student || "").trim();
  var pin = String(d.pin || "").trim();

  if (!student || !pin) {
    return {
      success: false,
      message: "Submission rejected: Student name and 4-digit PIN are required."
    };
  }

  var lock = LockService.getScriptLock();
  var acquired = false;
  try {
    acquired = lock.tryLock(25000); // Wait up to 25s in queue for other student writes to complete
  } catch (lockErr) {}

  if (!acquired) {
    return {
      success: false,
      message: "Server is busy recording other submissions. Please submit again in a few seconds."
    };
  }

  try {
    var ss = getSpreadsheet();
    var sheet = getLogsSheet(ss);

    if (!sheet) {
      var tabNames = ss.getSheets().map(function(s) { return "'" + s.getName() + "'"; }).join(", ");
      return {
        success: false,
        message: "Save failed: Responses sheet tab not found. Available tabs: [" + tabNames + "]"
      };
    }

    // Auto-create standard headers if sheet is completely empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp",
        "Student",
        "PIN",
        "Module",
        "Session",
        "Key vocabulary learnt",
        "Lesson content review (What I learnt in this session)",
        "What I need to work on following this session",
        "How will this help me with my formative evaluation?",
        "How well prepared do you think you currently are for your module 1 formative evaluation? (Rate out of 10)",
        "What have you done in terms of TOEIC preparation this week? Please give examples from Global Exam.",
        "How well prepared do you think you currently are for your TOEIC test? (Rate out of 10)"
      ]);
    }

    sheet.appendRow([
      new Date(),
      d.student || "",
      d.pin || "",
      d.module || "",
      d.session || "",
      d.vocab || "",
      d.contentReview || "",
      d.workOn || "",
      d.evalHelp || "",
      d.evalRating || d.formativeRating || "",
      d.toeicPrep || "",
      d.toeicRating || d.toeicPrepRating || ""
    ]);

    SpreadsheetApp.flush();

    return {
      success: true,
      message: "Session entry recorded successfully.",
      timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Europe/Paris", "yyyy-MM-dd HH:mm")
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (relErr) {}
  }
}

/**
 * Diagnostic Ping handler (used by connection checks and test URLs)
 */
function handlePing() {
  var ss = getSpreadsheet();
  var studentsSheet = getStudentsSheet(ss);
  var logsSheet = getLogsSheet(ss);
  var tabNames = ss.getSheets().map(function(s) { return s.getName(); });

  return {
    success: true,
    message: "GSEA Portfolio Web App is active and ready.",
    spreadsheetName: ss.getName(),
    availableTabs: tabNames,
    studentsTab: studentsSheet ? studentsSheet.getName() : "NOT FOUND",
    logsTab: logsSheet ? logsSheet.getName() : "NOT FOUND",
    studentsCount: (studentsSheet && studentsSheet.getLastRow() > 1) ? (studentsSheet.getLastRow() - 1) : 0,
    serverTime: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Europe/Paris", "yyyy-MM-dd HH:mm:ss")
  };
}

/**
 * Obtains Spreadsheet with graceful fallback to openById
 */
function getSpreadsheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch (e) {}

  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    throw new Error("Unable to open Spreadsheet. If this is a standalone Apps Script, ensure SPREADSHEET_ID is valid and shared with your Google account. Details: " + e.message);
  }
}

/**
 * Tolerates any tab name for Form Responses / PortfolioLogs
 */
function getLogsSheet(ss) {
  var candidates = [
    "Form Responses",
    "Form Responses 1",
    "PortfolioLogs",
    "Portfolio Logs",
    "Form_Responses",
    "Portfolio_Logs",
    "Responses",
    "Logs",
    "Réponses au formulaire 1"
  ];

  for (var i = 0; i < candidates.length; i++) {
    var s = ss.getSheetByName(candidates[i]);
    if (s) return s;
  }

  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    var nameLower = all[i].getName().trim().toLowerCase();
    if (nameLower.indexOf("response") !== -1 ||
        nameLower.indexOf("log") !== -1 ||
        nameLower.indexOf("form") !== -1 ||
        nameLower.indexOf("réponse") !== -1) {
      return all[i];
    }
  }

  for (var i = 0; i < all.length; i++) {
    if (all[i].getName().toLowerCase().indexOf("student") === -1) {
      return all[i];
    }
  }

  return all.length > 1 ? all[1] : all[0];
}

/**
 * Tolerates any tab name for Students
 */
function getStudentsSheet(ss) {
  var candidates = ["Students", "Student", "Élèves", "Users"];
  for (var i = 0; i < candidates.length; i++) {
    var s = ss.getSheetByName(candidates[i]);
    if (s) return s;
  }

  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    var nameLower = all[i].getName().trim().toLowerCase();
    if (nameLower.indexOf("student") !== -1 || nameLower.indexOf("élève") !== -1) {
      return all[i];
    }
  }

  return ss.getSheetByName("Students") || (all.length > 0 ? all[0] : null);
}

// ----------------------------------------------------------------------------
// DISAMBIGUATION HELPERS
// ----------------------------------------------------------------------------

function findCol(headers, words, fallback) {
  for (var i = 0; i < headers.length; i++) {
    for (var w = 0; w < words.length; w++) {
      if (headers[i].indexOf(words[w]) !== -1) return i;
    }
  }
  return fallback;
}

function findModuleCol(headers, fallback) {
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (h.indexOf("module") !== -1 &&
        h.indexOf("rate") === -1 &&
        h.indexOf("rating") === -1 &&
        h.indexOf("prepared") === -1 &&
        h.indexOf("10") === -1 &&
        h.indexOf("how well") === -1) {
      return i;
    }
  }
  return fallback;
}

function findEvalHelpCol(headers, fallback) {
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if ((h.indexOf("eval") !== -1 || h.indexOf("formative") !== -1) &&
        (h.indexOf("help") !== -1 || h.indexOf("aide") !== -1)) {
      return i;
    }
  }
  return fallback;
}

function findEvalRatingCol(headers, fallback) {
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if ((h.indexOf("eval") !== -1 || h.indexOf("formative") !== -1 || h.indexOf("prepared") !== -1) &&
        (h.indexOf("rate") !== -1 || h.indexOf("rating") !== -1 || h.indexOf("10") !== -1 || h.indexOf("scale") !== -1)) {
      return i;
    }
  }
  return fallback;
}

function findToeicPrepCol(headers, fallback) {
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (h.indexOf("toeic") !== -1 &&
        (h.indexOf("prep") !== -1 || h.indexOf("global") !== -1 || h.indexOf("week") !== -1 || h.indexOf("done") !== -1) &&
        h.indexOf("rate") === -1 &&
        h.indexOf("rating") === -1 &&
        h.indexOf("10") === -1) {
      return i;
    }
  }
  return fallback;
}

function findToeicRatingCol(headers, fallback) {
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (h.indexOf("toeic") !== -1 &&
        (h.indexOf("rate") !== -1 || h.indexOf("rating") !== -1 || h.indexOf("prepared") !== -1 || h.indexOf("10") !== -1)) {
      return i;
    }
  }
  return fallback;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}
