/**
 * TriagePluse Survey — Google Apps Script Backend
 * ================================================
 * Deploy this as a Web App to collect survey submissions into Google Sheets.
 *
 * SETUP INSTRUCTIONS:
 * 1. Open Google Sheets → Extensions → Apps Script
 * 2. Paste this entire script, replacing any existing code
 * 3. Click "Deploy" → "New deployment" → Type: "Web app"
 * 4. Set "Execute as" = Me, "Who has access" = Anyone
 * 5. Copy the Web App URL and paste it into index.html as WEBHOOK_URL
 *
 * SHEET STRUCTURE:
 * The script automatically creates two sheets:
 *   - "Part A — Triage Nurses"  (for role = nurse)
 *   - "Part B — Physicians/Managers" (for role = receiver)
 */

// ─── Configuration ───────────────────────────────────────────────────────────
var SPREADSHEET_ID = ''; // Leave empty to use the bound spreadsheet, or paste your Sheet ID

var PART_A_SHEET = 'Part A — Triage Nurses';
var PART_B_SHEET = 'Part B — Physicians & Managers';

// Column headers for Part A (Triage Nurse)
var HEADERS_A = [
  'Timestamp',
  'Part',
  'Q1 — Role',
  'Q2 — Years of ED Experience',
  'Q3 — Time to CTAS Assignment',
  'Q4 — Hardest Triage Challenge',
  'Q5 — Times Unsure (Last Month)',
  'Q6 — Action When Unsure',
  'Q7 — Used Aid for CTAS Decision',
  'Q8 — What Aid Used (Optional)',
  'Q9 — Difficult Patient Story ⭐',
  'Q10 — Resources at Triage Station',
  'Q11 — Resources Actually Used (Optional)',
  'Q12 — Would Want Decision Support Tool (1–5)',
  'Q13 — Additional Comments (Optional)'
];

// Column headers for Part B (Physician / Charge Nurse / Manager)
var HEADERS_B = [
  'Timestamp',
  'Part',
  'Q1 — Role',
  'Q2 — Years of ED Experience',
  'QB3 — Frequency of Re-assessment',
  'QB4 — CTAS Mismatch Frequency',
  'QB5 — CTAS Mismatch Story ⭐',
  'QB6 — Action When Level Seems Wrong',
  'QB7 — Most Common Downstream Effect',
  'QB8 — Confidence with Better Decision Support (1–5)',
  'QB9 — Additional Comments (Optional)'
];

// ─── Main Handler ─────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var data = {};

    // Parse form POST data
    if (e && e.parameter) {
      data = e.parameter;
    } else if (e && e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch(err) { data = {}; }
    }

    // Get or create spreadsheet
    var ss = SPREADSHEET_ID
      ? SpreadsheetApp.openById(SPREADSHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();

    var part = (data.part || '').toUpperCase();

    if (part === 'A') {
      appendRow(ss, PART_A_SHEET, HEADERS_A, buildRowA(data));
    } else if (part === 'B') {
      appendRow(ss, PART_B_SHEET, HEADERS_B, buildRowB(data));
    } else {
      // Fallback: write to a raw log sheet
      appendRaw(ss, data);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Also handle GET for testing
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'TriagePluse webhook is live' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Row Builders ─────────────────────────────────────────────────────────────
function buildRowA(d) {
  return [
    d.timestamp || new Date().toISOString(),
    'A — Triage Nurse',
    d.q1 || '',
    d.q2 || '',
    d.q3 || '',
    d.q4 || '',
    d.q5 || '',
    d.q6 || '',
    d.q7 || '',
    d.q8 || '',
    d.q9 || '',
    d.q10 || '',
    d.q11 || '',
    d.q12 || '',
    d.q13 || ''
  ];
}

function buildRowB(d) {
  return [
    d.timestamp || new Date().toISOString(),
    'B — Physician/Manager',
    d.q1 || '',
    d.q2 || '',
    d.qb3 || '',
    d.qb4 || '',
    d.qb5 || '',
    d.qb6 || '',
    d.qb7 || '',
    d.qb8 || '',
    d.qb9 || ''
  ];
}

// ─── Sheet Helpers ────────────────────────────────────────────────────────────

/**
 * Creates a sheet with formatted headers if it doesn't already exist.
 * Returns the sheet object.
 */
function createSheetWithHeaders(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // Add header row with formatting
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#1A2B4A');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(10);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 180);  // Timestamp
    sheet.setColumnWidth(2, 120);  // Part
    // Set wider columns for free-text answers
    [11, 12, 15].forEach(function(col) {
      if (col <= headers.length) sheet.setColumnWidth(col, 320);
    });
    // Standard width for remaining columns
    for (var i = 3; i <= headers.length; i++) {
      if ([11, 12, 15].indexOf(i) === -1) sheet.setColumnWidth(i, 180);
    }
  }
  return sheet;
}

/**
 * Appends a data row to the named sheet, creating it with headers first if needed.
 * rowData must be a non-empty array.
 */
function appendRow(ss, sheetName, headers, rowData) {
  // Guard: skip if rowData is empty (e.g. called from setupSheets)
  if (!rowData || rowData.length === 0) return;

  var sheet = createSheetWithHeaders(ss, sheetName, headers);

  // Append data row
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);

  // Alternate row colors for readability
  if ((lastRow + 1) % 2 === 0) {
    sheet.getRange(lastRow + 1, 1, 1, headers.length).setBackground('#F6F8FB');
  }
}

function appendRaw(ss, data) {
  var sheetName = 'Raw Log';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var hdr = sheet.getRange(1, 1, 1, 2);
    hdr.setValues([['Timestamp', 'Raw JSON']]);
    hdr.setBackground('#CC2229');
    hdr.setFontColor('#FFFFFF');
    hdr.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 600);
  }
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[
    new Date().toISOString(),
    JSON.stringify(data)
  ]]);
}

// ─── Setup Function (run once manually) ──────────────────────────────────────
/**
 * Run this function once from the Apps Script editor to pre-create
 * both sheets with proper headers and formatting.
 */
function setupSheets() {
  var ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  // Rename default "Sheet1" to a summary sheet
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) defaultSheet.setName('Summary');

  // Create Part A and Part B sheets with headers only (no data rows)
  createSheetWithHeaders(ss, PART_A_SHEET, HEADERS_A);
  createSheetWithHeaders(ss, PART_B_SHEET, HEADERS_B);

  // Format Summary sheet
  var summary = ss.getSheetByName('Summary');
  if (summary) {
    summary.getRange('A1').setValue('TriagePluse Survey — Response Tracker');
    summary.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#1A2B4A');
    summary.getRange('A3').setValue('Survey ID:').setFontWeight('bold');
    summary.getRange('B3').setValue('HMG/QID/1397');
    summary.getRange('A4').setValue('Hospital:').setFontWeight('bold');
    summary.getRange('B4').setValue('HMG Takhassusi Hospital — Emergency Department, Riyadh');
    summary.getRange('A5').setValue('Survey URL:').setFontWeight('bold');
    summary.getRange('B5').setValue('https://malneami.github.io/TriagePluse.survey/');
    summary.getRange('A6').setValue('Last Updated:').setFontWeight('bold');
    summary.getRange('B6').setValue(new Date().toLocaleDateString());
    summary.setColumnWidth(1, 160);
    summary.setColumnWidth(2, 400);
  }

  Logger.log('Setup complete! Sheets created successfully.');
}
