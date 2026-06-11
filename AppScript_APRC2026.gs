/**
 * APRC 2026 Deposit Coverage Calculator — Survey Data Collection Webhook
 *
 * Purpose: Receives survey submissions from V98-APRC.html and stores in Google Sheet
 *
 * Data Collected (survey tab):
 * - User ID (anonymous, localStorage-based)
 * - Session ID (per page load)
 * - Timestamp
 * - Tool identifier (APRC2026-DCC)
 * - Age group (Demographics Q1)
 * - Jurisdiction selected in calculator
 * - Respondent type (DI = Deposit Insurer / NDI = Non-Deposit Insurer)
 * - Q1 Triggers: withdrawal trigger events (multi-select, comma-separated)
 * - Q2 Destination: where to move money (single select)
 * - Q3 Reassurance: reassurance factors (multi-select, comma-separated)
 *
 * Data Collected (Interactions tab):
 * - Count (sequential, auto-incremented)
 * - Timestamp (yyyymmdd-hhmm) — logged when user reaches Regional Coverage page
 *
 * Created: March 2026
 * Updated: April 2026 — added DCC interaction tracking
 * Replaces: NRS2025 Survey Data Collection webhook
 */

function doPost(e) {
  try {
    // Parse incoming JSON payload
    var data = JSON.parse(e.postData.contents);

    // Route interaction events to the Interactions tab
    if (data.eventType === 'dcc_interaction') {
      return logInteraction(data);
    }

    // Get active sheet (first sheet tab) for survey submissions
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Write header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'User ID',
        'Session ID',
        'Timestamp',
        'Tool',
        'Age Group',
        'Jurisdiction',
        'Respondent Type',
        'Q1 Triggers',
        'Q2 Destination',
        'Q3 Reassurance'
      ]);
    }

    // Append the new response row
    sheet.appendRow([
      data.userId         || '',
      data.sessionId      || '',
      data.timestamp      || new Date().toISOString(),
      data.tool           || 'APRC2026-DCC',
      data.age            || '',
      data.jurisdiction   || '',
      data.respondentType || '',
      data.q1_triggers    || '',
      data.q2_destination || '',
      data.q3_reassurance || ''
    ]);

    // Return success response
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // Log error and return error response
    Logger.log('Error: ' + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Log a single DCC tool interaction to the Interactions tab.
 * Called when a user reaches Page 4 (Regional Coverage Comparison).
 * Columns: Count | Timestamp (yyyymmdd-hhmm)
 */
function logInteraction(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Interactions');

  // Create the Interactions tab if it does not exist yet
  if (!sheet) {
    sheet = ss.insertSheet('Interactions');
  }

  // Write header on first use
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Count', 'Timestamp']);
  }

  // Count = number of existing rows (header = row 1, so first data entry = 1, second = 2, ...)
  var count = sheet.getLastRow();

  sheet.appendRow([count, data.timestamp || '']);

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Test function — run this manually to verify the sheet is accessible
 */
function testSetup() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  Logger.log('Sheet name: ' + sheet.getName());
  Logger.log('Last row: ' + sheet.getLastRow());
  Logger.log('Setup OK');
}
