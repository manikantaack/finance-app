/**
 * Sahoo Finance Ledger - shared data bridge
 * ------------------------------------------
 * Paste this whole file into a Google Sheet's Apps Script editor
 * (Extensions -> Apps Script), then deploy it as a Web App:
 *   Deploy -> New deployment -> type: Web app
 *   Execute as: Me
 *   Who has access: Anyone
 * Copy the resulting URL (it ends in /exec) into SYNC_URL near the
 * top of src/App.jsx in the website project, then redeploy the site.
 *
 * All the app's data (agents, clients, loans, payments) is stored as
 * one JSON blob in cell A1 of a sheet named "Data" inside this
 * spreadsheet. Everyone using the app - every agent and the owner -
 * reads and writes that same cell, so everyone always sees the same
 * information.
 */

const SHEET_NAME = "Data";
const LOCK_TIMEOUT_MS = 10000;

function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange("A1").setValue("");
  }
  return sheet;
}

// Handles reads: the app's browser calls this with a plain GET request.
function doGet(e) {
  const sheet = getDataSheet_();
  const value = sheet.getRange("A1").getValue() || "";
  return ContentService.createTextOutput(value).setMimeType(ContentService.MimeType.JSON);
}

// Handles saves: the app's browser calls this with a POST request
// whose body is the full JSON data. Content-Type is sent as
// text/plain on purpose (see src/App.jsx) to avoid browser CORS
// preflight requests, which Apps Script Web Apps don't support.
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_TIMEOUT_MS);
  try {
    const body = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
    JSON.parse(body); // throws if it isn't valid JSON - reject bad writes
    const sheet = getDataSheet_();
    sheet.getRange("A1").setValue(body);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
