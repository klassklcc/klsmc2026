/**
 * Google Apps Script — Deploy as Web App
 *
 * Hints Used columns per question:
 *   Q2 = Col H (8),  Q3 = Col K (11), Q4 = Col N (14)
 *   Q5 = Col Q (17), Q6 = Col T (20), Q7 = Col W (23), Q8 = Col Z (26)
 *
 * Params: ?teamCode=XXXX&question=Q2
 */

var QUESTION_COLUMNS = {
  Q2: 8,
  Q3: 11,
  Q4: 14,
  Q5: 17,
  Q6: 20,
  Q7: 23,
  Q8: 26,
};

function doGet(e) {
  try {
    var teamCode = e.parameter.teamCode;
    var question = e.parameter.question;
    Logger.log(
      "doGet called with teamCode: " + teamCode + ", question: " + question,
    );

    if (!teamCode) {
      return ContentService.createTextOutput(
        JSON.stringify({ error: "Missing teamCode" }),
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (!question || !QUESTION_COLUMNS[question]) {
      return ContentService.createTextOutput(
        JSON.stringify({
          error:
            "Missing or invalid question. Valid: " +
            Object.keys(QUESTION_COLUMNS).join(", "),
        }),
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var colNumber = QUESTION_COLUMNS[question];
    var colIndex = colNumber - 1;

    // IMPORTANT: Replace with your actual spreadsheet ID
    var ss = SpreadsheetApp.openById("PASTE_YOUR_SPREADSHEET_ID_HERE");
    var sheet = ss.getSheetByName("Sheet1");

    if (!sheet) {
      var sheetNames = ss.getSheets().map(function (s) {
        return s.getName();
      });
      return ContentService.createTextOutput(
        JSON.stringify({
          error: "Sheet not found. Available: " + sheetNames.join(", "),
        }),
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();

    for (var i = 3; i < values.length; i++) {
      var cellVal = String(values[i][1]).trim();
      if (cellVal === teamCode) {
        var currentVal = values[i][colIndex];
        Logger.log(
          "Found team at row " +
            (i + 1) +
            ", col " +
            colNumber +
            " value: " +
            currentVal,
        );
        var newVal =
          typeof currentVal === "number" && !isNaN(currentVal)
            ? currentVal + 1
            : 1;
        sheet.getRange(i + 1, colNumber).setValue(newVal);
        return ContentService.createTextOutput(
          JSON.stringify({
            success: true,
            team: teamCode,
            question: question,
            hintsUsed: newVal,
          }),
        ).setMimeType(ContentService.MimeType.JSON);
      }
    }

    return ContentService.createTextOutput(
      JSON.stringify({ error: "Team not found: " + teamCode }),
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (ex) {
    Logger.log("EXCEPTION: " + ex.toString());
    return ContentService.createTextOutput(
      JSON.stringify({ error: "Exception: " + ex.toString() }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
