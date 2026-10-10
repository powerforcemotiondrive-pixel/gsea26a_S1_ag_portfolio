// ============================================================================
// GSEA ENGLISH LEARNING PORTFOLIO — APPS SCRIPT BACKEND
// Spreadsheet: GSEA_English_Portfolio_DB
// Tabs: "Students" & "Form Responses"
// ============================================================================

function testConnection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentsSheet = ss.getSheetByName("Students");
  var logsSheet = ss.getSheetByName("Form Responses") || ss.getSheetByName("Form Responses 1") || ss.getSheets()[0];

  Logger.log("Spreadsheet: " + ss.getName());
  Logger.log("Students tab: " + (studentsSheet ? "FOUND (" + studentsSheet.getLastRow() + " rows)" : "NOT FOUND"));
  Logger.log("Responses tab: " + (logsSheet ? "FOUND (" + logsSheet.getName() + ")" : "NOT FOUND"));
}

function doGet(e) {
  try {
    var student = (e && e.parameter && e.parameter.student) ? String(e.parameter.student).trim() : "";
    var pin = (e && e.parameter && e.parameter.pin) ? String(e.parameter.pin).trim() : "";

    if (!student || !pin) {
      return jsonOut({ success: true, message: "GSEA Portfolio Web App is active." });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. PIN Authentication
    var studentsSheet = ss.getSheetByName("Students");
    if (!studentsSheet) {
      return jsonOut({ success: false, message: "Sheet tab 'Students' not found." });
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

    if (!found) return jsonOut({ success: false, message: "Student '" + student + "' not found in Students list." });
    if (!authorized) return jsonOut({ success: false, message: "Incorrect PIN for " + student + "." });

    // 2. Fetch Form Responses
    var logsSheet = ss.getSheetByName("Form Responses") || ss.getSheetByName("Form Responses 1") || ss.getSheets()[0];
    var data = logsSheet.getDataRange().getValues();
    var logs = [];

    if (data.length > 1) {
      var h = data[0].map(function(item) { return String(item || "").toLowerCase(); });
      var cTs  = getCol(h, ["time", "date", "horodateur"], 0);
      var cSt  = getCol(h, ["student", "name", "nom"], 1);
      var cMod = getCol(h, ["module"], 2);
      var cSes = getCol(h, ["session"], 3);
      var cVoc = getCol(h, ["vocab", "word", "lexique"], 4);
      var cRev = getCol(h, ["review", "content", "learnt"], 5);
      var cWrk = getCol(h, ["work on", "need to"], 6);
      var cEvH = getCol(h, ["formative", "eval", "help"], 7);
      var cEvR = getCol(h, ["formative", "module", "rate", "rating", "prepared", "10"], 8);
      var cToP = getCol(h, ["toeic prep", "toeic"], 9);
      var cToR = getCol(h, ["toeic", "rate", "rating", "prepared", "10"], 10);

      for (var r = 1; r < data.length; r++) {
        var row = data[r];
        var rowStudent = String(row[cSt] || "").trim();

        if (rowStudent.toLowerCase() === student.toLowerCase()) {
          var ts = row[cTs] ? String(row[cTs]) : "";
          if (row[cTs] instanceof Date) {
            ts = Utilities.formatDate(row[cTs], Session.getScriptTimeZone() || "Europe/Paris", "yyyy-MM-dd HH:mm");
          }

          logs.push({
            timestamp: ts,
            module: String(row[cMod] || ""),
            session: String(row[cSes] || ""),
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

    return jsonOut({ success: true, student: student, logs: logs });

  } catch (err) {
    return jsonOut({ success: false, message: "Server Error: " + err.message });
  }
}

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Form Responses") || ss.getSheetByName("Form Responses 1") || ss.getSheets()[0];

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

    return jsonOut({ success: true, message: "Saved successfully." });
  } catch (err) {
    return jsonOut({ success: false, message: "Save failed: " + err.message });
  }
}

function getCol(headers, words, fallback) {
  for (var i = 0; i < headers.length; i++) {
    for (var w = 0; w < words.length; w++) {
      if (headers[i].indexOf(words[w]) !== -1) return i;
    }
  }
  return fallback;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

