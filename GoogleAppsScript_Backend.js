// ============================================================================
// GSEA ENGLISH LEARNING PORTFOLIO — APPS SCRIPT BACKEND
// Spreadsheet: GSEA_English_Portfolio_DB
// Compatible with sheet tab names:
// - "Form Responses", "Form Responses 1", "PortfolioLogs", "Portfolio Logs"
// - "Students"
// ============================================================================

function testConnection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentsSheet = getStudentsSheet(ss);
  var logsSheet = getLogsSheet(ss);

  Logger.log("=== GSEA DATABASE DIAGNOSTIC ===");
  Logger.log("Spreadsheet: " + ss.getName());
  Logger.log("Students tab: " + (studentsSheet ? "FOUND '" + studentsSheet.getName() + "' (" + studentsSheet.getLastRow() + " rows)" : "NOT FOUND"));
  Logger.log("Logs tab: " + (logsSheet ? "FOUND '" + logsSheet.getName() + "' (" + logsSheet.getLastRow() + " rows)" : "NOT FOUND"));
  
  if (studentsSheet && studentsSheet.getLastRow() > 1) {
    var sampleStudent = studentsSheet.getRange(2, 1).getValue();
    Logger.log("Sample Student in DB: " + sampleStudent);
  }
}

// Backwards compatibility alias for older script calls
function loginUser(student, pin) {
  return doGet({ parameter: { student: student, pin: pin } });
}

function doGet(e) {
  try {
    var student = (e && e.parameter && e.parameter.student) ? String(e.parameter.student).trim() : "";
    var pin = (e && e.parameter && e.parameter.pin) ? String(e.parameter.pin).trim() : "";

    if (!student || !pin) {
      return jsonOut({ success: true, message: "GSEA Portfolio Web App is active and ready." });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.flush(); // Commit pending writes so latest data is immediately readable

    // 1. PIN Authentication
    var studentsSheet = getStudentsSheet(ss);
    if (!studentsSheet) {
      return jsonOut({ success: false, message: "Database error: 'Students' tab not found in spreadsheet." });
    }

    var students = studentsSheet.getDataRange().getValues();
    var authorized = false;
    var found = false;

    for (var i = 1; i < students.length; i++) {
      var sName = String(students[i][0] || "").trim();
      var sPin = String(students[i][1] != null ? students[i][1] : "").trim();
      if (/^\d{1,3}$/.test(sPin) && pin.length === 4) sPin = sPin.padStart(4, "0");

      if (sName.toLowerCase() === student.toLowerCase()) {
        found = true;
        if (sPin === pin) authorized = true;
        break;
      }
    }

    if (!found) return jsonOut({ success: false, message: "Student '" + student + "' not found in registered students list." });
    if (!authorized) return jsonOut({ success: false, message: "Incorrect PIN for " + student + ". Please verify your 4-digit code." });

    // 2. Fetch Form Responses / PortfolioLogs
    var logsSheet = getLogsSheet(ss);
    if (!logsSheet) {
      return jsonOut({ success: false, message: "Database error: Responses tab ('Form Responses' or 'PortfolioLogs') not found." });
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

      // Column Mappings with Disambiguation & Fallbacks
      var cTs  = findCol(rawHeaders, ["time", "date", "horodateur"], 0);
      var cSt  = findCol(rawHeaders, ["student", "nom", "name"], 1);
      
      // Module column must match "module" but NOT "rate", "rating", "prepared", "how well", "10", "eval"
      var cMod = findModuleCol(rawHeaders, hasPinCol ? 3 : 2);

      // Session column
      var cSes = findCol(rawHeaders, ["session"], hasPinCol ? 4 : 3);

      // Key vocabulary learnt
      var cVoc = findCol(rawHeaders, ["vocab", "word", "lexique"], hasPinCol ? 5 : 4);

      // Lesson content review
      var cRev = findCol(rawHeaders, ["review", "learnt", "lesson content", "content"], hasPinCol ? 6 : 5);

      // What I need to work on
      var cWrk = findCol(rawHeaders, ["work on", "need to work", "following this session", "work"], hasPinCol ? 7 : 6);

      // How will this help me with formative evaluation?
      var cEvH = findEvalHelpCol(rawHeaders, hasPinCol ? 8 : 7);

      // Module 1 formative evaluation preparedness rating (out of 10)
      var cEvR = findEvalRatingCol(rawHeaders, hasPinCol ? 9 : 8);

      // TOEIC prep this week
      var cToP = findToeicPrepCol(rawHeaders, hasPinCol ? 10 : 9);

      // TOEIC test preparedness rating (out of 10)
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

          // Self-healing: if modVal captured a PIN and sesVal contains "Module", automatically fix offset
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

    return jsonOut({
      success: true,
      student: student,
      logs: logs,
      serverTime: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Europe/Paris", "yyyy-MM-dd HH:mm:ss")
    });

  } catch (err) {
    return jsonOut({ success: false, message: "Server Error: " + err.message });
  }
}

function doPost(e) {
  try {
    var contents = e.postData ? e.postData.contents : "";
    var d = {};
    if (contents) {
      try {
        d = JSON.parse(contents);
      } catch (pe) {
        d = e.parameter || {};
      }
    } else if (e.parameter) {
      d = e.parameter;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getLogsSheet(ss);

    if (!sheet) {
      return jsonOut({ success: false, message: "Save failed: Responses sheet tab not found." });
    }

    // Auto-create standard headers if sheet is empty
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

    // Commit write immediately to Google Sheets so subsequent GET reads find the new row
    SpreadsheetApp.flush();

    return jsonOut({
      success: true,
      message: "Session entry recorded successfully.",
      timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Europe/Paris", "yyyy-MM-dd HH:mm")
    });
  } catch (err) {
    return jsonOut({ success: false, message: "Save failed: " + err.message });
  }
}

// ----------------------------------------------------------------------------
// SMART SHEET LOCATORS (Tolerates Any Tab Name: Form Responses, PortfolioLogs, etc.)
// ----------------------------------------------------------------------------

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

  // Case-insensitive & partial match
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

  // Any sheet that is not Students
  for (var i = 0; i < all.length; i++) {
    if (all[i].getName().toLowerCase().indexOf("student") === -1) {
      return all[i];
    }
  }

  return all.length > 1 ? all[1] : all[0];
}

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

  return ss.getSheetByName("Students") || all[0];
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
    // Must contain "module", but NOT be the preparedness rating question
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
