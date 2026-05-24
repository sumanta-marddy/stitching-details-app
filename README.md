function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var raw = "";

    if (e && e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
    } else if (e && e.postData && e.postData.contents) {
      raw = e.postData.contents;
      if (raw.indexOf("payload=") === 0) {
        raw = decodeURIComponent(raw.replace("payload=", ""));
      }
    }

    var data = JSON.parse(raw || "{}");
    var headers = data.headers || [];
    var rows = data.rows || [];

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("All Entries") || ss.insertSheet("All Entries");

    sheet.clearContents();

    if (!headers.length) {
      sheet.getRange(1, 1).setValue("NO HEADERS");
      return output_("NO HEADERS");
    }

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    if (rows.length) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    return output_("SUCCESS ROWS: " + rows.length);

  } catch (err) {
    return output_("ERROR: " + err.message);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function doGet(e) {
  return output_("WEB APP WORKING");
}

function output_(text) {
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.TEXT);
}
