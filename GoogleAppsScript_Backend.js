/**
 * ============================================================================
 * GSEA ENGLISH LEARNING PORTFOLIO — GOOGLE APPS SCRIPT BACKEND
 * ============================================================================
 * 
 * Instructions:
 * 1. Open your Google Spreadsheet (where your Google Form is linked).
 * 2. Ensure you have two sheets:
 *    - "Students": Column A = Student Name, Column B = 4-Digit Security PIN
 *    - "Form Responses 1": Automatically created when linking your Google Form
 * 3. Go to "Extensions" > "Apps Script".
 * 4. Replace any existing code in Code.gs with this entire file.
 * 5. Click "Deploy" > "New deployment".
 * 6. Select Type: "Web app".
 * 7. Set:
 *    - Execute as: "Me"
 *    - Who has access: "Anyone" (required for student portfolio access)
 * 8. Copy the Web App URL and paste it into index.html as SCRIPT_URL.
 */

// ============================================================================
// 1. GET REQUEST: STUDENT AUTHENTICATION & FETCHING PORTFOLIO LOGS
// ============================================================================
function doGet(e) {
  try {
    const student = (e.parameter.student || "").trim();
    const pin = (e.parameter.pin || "").trim();

    if (!student || !pin) {
      return jsonResponse({
        success: false,
        message: "Please provide both Student Name and 4-digit PIN."
      });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- STEP A: VERIFY PIN IN 'Students' SHEET ---
    const studentsSheet = ss.getSheetByName("Students");
    if (!studentsSheet) {
      return jsonResponse({
        success: false,
        message: "Sheet 'Students' not found. Please ensure Column A has names and Column B has PINs."
      });
    }

    const studentsData = studentsSheet.getDataRange().getValues();
    let isAuthorized = false;

    for (let i = 1; i < studentsData.length; i++) {
      const rowName = String(studentsData[i][0] || "").trim();
      const rowPin = String(studentsData[i][1] || "").trim();

      if (rowName.toLowerCase() === student.toLowerCase()) {
        if (rowPin === pin) {
          isAuthorized = true;
        }
        break;
      }
    }

    if (!isAuthorized) {
      return jsonResponse({
        success: false,
        message: "Authentication failed. The PIN entered does not match our records for " + student + "."
      });
    }

    // --- STEP B: FETCH LOGS FROM GOOGLE FORM RESPONSES SHEET ---
    // Look for the Form Responses sheet (usually "Form Responses 1" or sheet with "Responses")
    let logsSheet = ss.getSheetByName("Form Responses 1");
    if (!logsSheet) {
      const allSheets = ss.getSheets();
      for (let s of allSheets) {
        const name = s.getName().toLowerCase();
        if (name.includes("response") || name.includes("réponse") || name.includes("logs")) {
          logsSheet = s;
          break;
        }
      }
    }
    if (!logsSheet) {
      // Fallback to second sheet or first sheet
      logsSheet = ss.getSheets().length > 1 ? ss.getSheets()[1] : ss.getSheets()[0];
    }

    const logsData = logsSheet.getDataRange().getValues();
    if (logsData.length <= 1) {
      return jsonResponse({
        success: true,
        logs: []
      });
    }

    // Dynamic Header Column Mapping (case-insensitive fuzzy matching)
    const headers = logsData[0].map(h => String(h || "").trim().toLowerCase());
    
    const colTimestamp = headers.findIndex(h => h.includes("timestamp") || h.includes("horodateur"));
    const colStudent   = headers.findIndex(h => h.includes("student") || h.includes("name"));
    const colModule    = headers.findIndex(h => h.includes("module"));
    const colSession   = headers.findIndex(h => h.includes("session"));
    const colVocab     = headers.findIndex(h => h.includes("vocab"));
    const colReview    = headers.findIndex(h => h.includes("content") || h.includes("learnt") || h.includes("review"));
    const colWorkOn    = headers.findIndex(h => h.includes("work on") || h.includes("need to"));
    const colEvalHelp  = headers.findIndex(h => h.includes("formative") && (h.includes("eval") || h.includes("help") || h.includes("tutorial")));
    const colEvalRate  = headers.findIndex(h => (h.includes("rating") || h.includes("rate") || h.includes("prepared")) && (h.includes("module") || h.includes("formative")));
    const colToeicPrep = headers.findIndex(h => h.includes("toeic prep") || (h.includes("toeic") && !h.includes("rating") && !h.includes("rate") && !h.includes("prepared")));
    const colToeicRate = headers.findIndex(h => (h.includes("rating") || h.includes("rate") || h.includes("prepared")) && h.includes("toeic"));

    const logs = [];

    for (let r = 1; r < logsData.length; r++) {
      const row = logsData[r];
      const rowStudent = colStudent !== -1 ? String(row[colStudent] || "").trim() : "";

      if (rowStudent.toLowerCase() === student.toLowerCase()) {
        let tsString = "";
        if (colTimestamp !== -1 && row[colTimestamp]) {
          try {
            tsString = Utilities.formatDate(new Date(row[colTimestamp]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
          } catch (e) {
            tsString = String(row[colTimestamp]);
          }
        }

        logs.push({
          timestamp: tsString,
          module: colModule !== -1 ? String(row[colModule] || "") : "",
          session: colSession !== -1 ? String(row[colSession] || "") : "",
          vocab: colVocab !== -1 ? String(row[colVocab] || "") : "",
          contentReview: colReview !== -1 ? String(row[colReview] || "") : "",
          workOn: colWorkOn !== -1 ? String(row[colWorkOn] || "") : "",
          evalHelp: colEvalHelp !== -1 ? String(row[colEvalHelp] || "") : "",
          evalRating: colEvalRate !== -1 ? String(row[colEvalRate] || "") : "",
          toeicPrep: colToeicPrep !== -1 ? String(row[colToeicPrep] || "") : "",
          toeicRating: colToeicRate !== -1 ? String(row[colToeicRate] || "") : ""
        });
      }
    }

    return jsonResponse({
      success: true,
      logs: logs
    });

  } catch (error) {
    return jsonResponse({
      success: false,
      message: "Server Error: " + error.message
    });
  }
}

// ============================================================================
// 2. POST REQUEST (OPTIONAL FALLBACK WITH LOCKSERVICE CONCURRENCY CONTROL)
// ============================================================================
function doPost(e) {
  // LockService prevents concurrency race conditions and write locks
  const lock = LockService.getScriptLock();
  
  try {
    // Wait up to 30 seconds for concurrent requests to queue cleanly
    lock.waitLock(30000);

    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Verify PIN
    const studentsSheet = ss.getSheetByName("Students");
    if (studentsSheet) {
      const studentsData = studentsSheet.getDataRange().getValues();
      let valid = false;
      for (let i = 1; i < studentsData.length; i++) {
        if (String(studentsData[i][0]).trim().toLowerCase() === String(data.student).trim().toLowerCase()) {
          if (String(studentsData[i][1]).trim() === String(data.pin).trim()) {
            valid = true;
          }
          break;
        }
      }
      if (!valid) {
        return jsonResponse({ success: false, message: "Invalid PIN." });
      }
    }

    let targetSheet = ss.getSheetByName("Form Responses 1") || ss.getSheetByName("Logs") || ss.getSheets()[0];

    targetSheet.appendRow([
      new Date(),
      data.student || "",
      data.pin || "",
      data.module || "",
      data.session || "",
      data.vocab || "",
      data.contentReview || "",
      data.workOn || "",
      data.evalHelp || "",
      data.evalRating || data.formativeRating || "",
      data.toeicPrep || "",
      data.toeicRating || data.toeicPrepRating || ""
    ]);

    return jsonResponse({
      success: true,
      message: "Entry saved successfully."
    });

  } catch (error) {
    return jsonResponse({
      success: false,
      message: "Submission error: " + error.message
    });
  } finally {
    lock.releaseLock();
  }
}

// Helper: JSON response with proper CORS headers
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

