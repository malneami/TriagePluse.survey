/**
 * TriagePluse Survey — Google Apps Script Backend + Dashboard
 * ============================================================
 * Deploy this as a Web App to collect survey submissions into Google Sheets
 * and auto-generate a live analytics dashboard.
 *
 * SETUP INSTRUCTIONS:
 * 1. Open Google Sheets → Extensions → Apps Script
 * 2. Paste this entire script, replacing any existing code
 * 3. Run setupSheets() once manually to create all sheets
 * 4. Click "Deploy" → "New deployment" → Type: "Web app"
 * 5. Set "Execute as" = Me, "Who has access" = Anyone
 * 6. Copy the Web App URL and paste it into index.html as WEBHOOK_URL
 * 7. To rebuild the dashboard at any time, run buildDashboard() manually
 *
 * AUTO-REFRESH DASHBOARD:
 * To auto-refresh the dashboard every hour:
 * 1. Click "Triggers" (clock icon) in the Apps Script editor
 * 2. Add trigger → Function: buildDashboard → Time-driven → Hour timer → Every hour
 */

// ─── Configuration ───────────────────────────────────────────────────────────
var SPREADSHEET_ID = ''; // Leave empty to use the bound spreadsheet
var PART_A_SHEET   = 'Part A — Triage Nurses';
var PART_B_SHEET   = 'Part B — Physicians & Managers';
var DASHBOARD_SHEET = 'Dashboard';

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
  'Q9 — Difficult Patient Story',
  'Q10 — Resources at Triage Station',
  'Q11 — Resources Actually Used (Optional)',
  'Q12 — Would Want Decision Support Tool (1-5)',
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
  'QB5 — CTAS Mismatch Story',
  'QB6 — Action When Level Seems Wrong',
  'QB7 — Most Common Downstream Effect',
  'QB8 — Confidence with Better Decision Support (1-5)',
  'QB9 — Additional Comments (Optional)'
];

// ─── Main Handlers ────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var data = {};
    if (e && e.parameter) {
      data = e.parameter;
    } else if (e && e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch(err) { data = {}; }
    }
    var ss = SPREADSHEET_ID
      ? SpreadsheetApp.openById(SPREADSHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();
    var part = (data.part || '').toUpperCase();
    if (part === 'A') {
      appendRow(ss, PART_A_SHEET, HEADERS_A, buildRowA(data));
    } else if (part === 'B') {
      appendRow(ss, PART_B_SHEET, HEADERS_B, buildRowB(data));
    } else {
      appendRaw(ss, data);
    }
    // Rebuild dashboard after each submission
    try { buildDashboard(); } catch(de) {}
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', row: ss.getSheetByName(part === 'A' ? PART_A_SHEET : PART_B_SHEET) ? ss.getSheetByName(part === 'A' ? PART_A_SHEET : PART_B_SHEET).getLastRow() : 0 }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  // Handle ?action=summary — return all rows as JSON for the admin dashboard
  // Supports JSONP via ?callback= parameter to bypass browser CORS restrictions
  var callback = e && e.parameter && e.parameter.callback ? e.parameter.callback : null;
  function respond(obj) {
    var json = JSON.stringify(obj);
    if (callback) {
      // JSONP: wrap in callback function call so browser can read it cross-origin
      return ContentService
        .createTextOutput(callback + '(' + json + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    if (e && e.parameter && e.parameter.action === 'summary') {
      var ss0 = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
      var rows = [];
      var sheetsToRead = [PART_A_SHEET, PART_B_SHEET];
      sheetsToRead.forEach(function(sheetName) {
        var sh = ss0.getSheetByName(sheetName);
        if (!sh || sh.getLastRow() < 2) return;
        var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
        data.forEach(function(row) {
          var obj = {};
          obj.part = (row[1] || '').indexOf('A') !== -1 ? 'A' : 'B';
          obj.timestamp = row[0] ? row[0].toString() : '';
          obj.experience = row[3] || '';
          if (obj.part === 'A') {
            obj.q3 = row[4] || ''; obj.q4 = row[5] || ''; obj.q5 = row[6] || '';
            obj.q6 = row[7] || ''; obj.q7 = row[8] || ''; obj.q8 = row[9] || '';
            obj.q9 = row[10] || ''; obj.q10 = row[11] || ''; obj.q11 = row[12] || '';
            obj.q12 = row[13] || ''; obj.q13 = row[14] || '';
          } else {
            obj.qb3 = row[4] || ''; obj.qb4 = row[5] || ''; obj.qb5 = row[6] || '';
            obj.qb6 = row[7] || ''; obj.qb7 = row[8] || ''; obj.qb8 = row[9] || '';
            obj.qb9 = row[10] || '';
          }
          rows.push(obj);
        });
      });
      return respond({ status: 'ok', rows: rows, total: rows.length });
    }
  } catch(sumErr) {
    return respond({ status: 'error', message: sumErr.toString() });
  }
  // Also handle GET submissions from the browser (no-cors fetch)
  try {
    if (e && e.parameter && e.parameter.part) {
      var data = e.parameter;
      var ss = SPREADSHEET_ID
        ? SpreadsheetApp.openById(SPREADSHEET_ID)
        : SpreadsheetApp.getActiveSpreadsheet();
      var part = (data.part || '').toUpperCase();
      if (part === 'A') {
        appendRow(ss, PART_A_SHEET, HEADERS_A, buildRowA(data));
      } else if (part === 'B') {
        appendRow(ss, PART_B_SHEET, HEADERS_B, buildRowB(data));
      } else {
        appendRaw(ss, data);
      }
      try { buildDashboard(); } catch(de) {}
      var lastRow = ss.getSheetByName(part === 'A' ? PART_A_SHEET : PART_B_SHEET);
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'success', row: lastRow ? lastRow.getLastRow() : 0 }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch(err) {}
  // Default GET response (health check)
  var ss2 = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var countA = 0, countB = 0;
  try { var sa = ss2.getSheetByName(PART_A_SHEET); if(sa) countA = Math.max(0, sa.getLastRow()-1); } catch(e2){}
  try { var sb = ss2.getSheetByName(PART_B_SHEET); if(sb) countB = Math.max(0, sb.getLastRow()-1); } catch(e3){}
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'TriagePulse webhook is LIVE',
      sheet: ss2.getName(),
      responses: countA + countB,
      partA: countA,
      partB: countB,
      time: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Row Builders ─────────────────────────────────────────────────────────────
function buildRowA(d) {
  return [
    d.timestamp || new Date().toISOString(),
    'A - Triage Nurse',
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
    'B - Physician/Manager',
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
function createSheetWithHeaders(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#1A2B4A');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(10);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 140);
    for (var i = 3; i <= headers.length; i++) {
      sheet.setColumnWidth(i, 200);
    }
    // Wider columns for free-text answers
    [9, 11, 13, 15].forEach(function(col) {
      if (col <= headers.length) sheet.setColumnWidth(col, 360);
    });
  }
  return sheet;
}

function appendRow(ss, sheetName, headers, rowData) {
  if (!rowData || rowData.length === 0) return;
  var sheet = createSheetWithHeaders(ss, sheetName, headers);
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
  if ((lastRow) % 2 === 0) {
    sheet.getRange(lastRow + 1, 1, 1, headers.length).setBackground('#EEF2F8');
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

// ─── Dashboard Builder ────────────────────────────────────────────────────────
/**
 * Builds or refreshes the Dashboard sheet with survey analytics.
 * Run manually or set a time-based trigger to auto-refresh.
 */
function buildDashboard() {
  var ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  // Get or create Dashboard sheet
  var dash = ss.getSheetByName(DASHBOARD_SHEET);
  if (!dash) {
    dash = ss.insertSheet(DASHBOARD_SHEET, 0); // Insert as first sheet
  } else {
    dash.clearContents();
    dash.clearFormats();
    // Remove existing charts
    var charts = dash.getCharts();
    charts.forEach(function(c) { dash.removeChart(c); });
  }

  // ── Fetch data ──
  var sheetA = ss.getSheetByName(PART_A_SHEET);
  var sheetB = ss.getSheetByName(PART_B_SHEET);
  var dataA = sheetA && sheetA.getLastRow() > 1
    ? sheetA.getRange(2, 1, sheetA.getLastRow() - 1, HEADERS_A.length).getValues()
    : [];
  var dataB = sheetB && sheetB.getLastRow() > 1
    ? sheetB.getRange(2, 1, sheetB.getLastRow() - 1, HEADERS_B.length).getValues()
    : [];

  var totalA = dataA.length;
  var totalB = dataB.length;
  var total  = totalA + totalB;

  // ── Helper: count values in a column ──
  function countCol(data, colIdx) {
    var counts = {};
    data.forEach(function(row) {
      var val = (row[colIdx] || '').toString().trim();
      if (val) counts[val] = (counts[val] || 0) + 1;
    });
    return counts;
  }

  // ── Helper: average of a column (numeric) ──
  function avgCol(data, colIdx) {
    var sum = 0, n = 0;
    data.forEach(function(row) {
      var v = parseFloat(row[colIdx]);
      if (!isNaN(v)) { sum += v; n++; }
    });
    return n > 0 ? (sum / n).toFixed(2) : 'N/A';
  }

  // ── Style helpers ──
  var HMG_NAVY  = '#1A2B4A';
  var HMG_RED   = '#CC2229';
  var TEAL      = '#0FA89B';
  var AMBER     = '#E8A020';
  var LIGHT_BG  = '#EEF2F8';
  var WHITE     = '#FFFFFF';

  function styleTitle(range, bg, fg) {
    range.setBackground(bg || HMG_NAVY)
         .setFontColor(fg || WHITE)
         .setFontWeight('bold')
         .setFontSize(11)
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
  }

  function styleHeader(range) {
    range.setBackground(HMG_NAVY)
         .setFontColor(WHITE)
         .setFontWeight('bold')
         .setFontSize(10)
         .setHorizontalAlignment('center');
  }

  function styleData(range) {
    range.setFontSize(10)
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
  }

  // ── Set column widths ──
  dash.setColumnWidth(1, 30);   // spacer
  dash.setColumnWidth(2, 220);  // labels
  dash.setColumnWidth(3, 100);  // values
  dash.setColumnWidth(4, 100);  // pct
  dash.setColumnWidth(5, 30);   // spacer
  dash.setColumnWidth(6, 220);  // labels
  dash.setColumnWidth(7, 100);  // values
  dash.setColumnWidth(8, 100);  // pct
  dash.setColumnWidth(9, 30);   // spacer

  var row = 1;

  // ══════════════════════════════════════════════════════════
  // SECTION 1 — MAIN TITLE BANNER
  // ══════════════════════════════════════════════════════════
  dash.setRowHeight(row, 50);
  var titleRange = dash.getRange(row, 1, 1, 9);
  titleRange.merge();
  titleRange.setValue('TriagePluse Survey — Analytics Dashboard | HMG/QID/1397');
  titleRange.setBackground(HMG_NAVY)
            .setFontColor(WHITE)
            .setFontWeight('bold')
            .setFontSize(16)
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle');
  row++;

  dash.setRowHeight(row, 22);
  var subRange = dash.getRange(row, 1, 1, 9);
  subRange.merge();
  subRange.setValue('HMG Takhassusi Hospital — Emergency Department, Riyadh   |   Last updated: ' + new Date().toLocaleString());
  subRange.setBackground('#243B63')
          .setFontColor('rgba(255,255,255,0.8)')
          .setFontSize(10)
          .setHorizontalAlignment('center')
          .setItalic(true);
  row += 2;

  // ══════════════════════════════════════════════════════════
  // SECTION 2 — KPI SUMMARY CARDS (row 4–7)
  // ══════════════════════════════════════════════════════════
  dash.setRowHeight(row, 18);
  var kpiTitle = dash.getRange(row, 1, 1, 9);
  kpiTitle.merge().setValue('RESPONSE OVERVIEW');
  styleTitle(kpiTitle, HMG_RED);
  row++;

  // KPI row
  dash.setRowHeight(row, 52);
  var kpis = [
    ['Total Responses', total],
    ['Part A — Triage Nurses', totalA],
    ['Part B — Physicians/Managers', totalB],
    ['Completion Rate', total > 0 ? '100%' : 'N/A']
  ];
  var kpiCols = [1, 3, 5, 7];
  kpis.forEach(function(kpi, i) {
    var col = kpiCols[i];
    var kpiRange = dash.getRange(row, col, 1, 2);
    kpiRange.merge();
    kpiRange.setValue(kpi[0] + '\n' + kpi[1]);
    kpiRange.setBackground(i === 0 ? HMG_NAVY : i === 1 ? TEAL : i === 2 ? AMBER : '#6B4FBB');
    kpiRange.setFontColor(WHITE)
            .setFontWeight('bold')
            .setFontSize(13)
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle')
            .setWrap(true);
  });
  row += 2;

  // ══════════════════════════════════════════════════════════
  // SECTION 3 — PART A ANALYTICS (left column)
  // ══════════════════════════════════════════════════════════
  dash.setRowHeight(row, 18);
  var aTitle = dash.getRange(row, 1, 1, 4);
  aTitle.merge().setValue('PART A — TRIAGE NURSES (' + totalA + ' responses)');
  styleTitle(aTitle, TEAL);

  var bTitle = dash.getRange(row, 5, 1, 4);
  bTitle.merge().setValue('PART B — PHYSICIANS & MANAGERS (' + totalB + ' responses)');
  styleTitle(bTitle, AMBER);
  row++;

  // ── Q2: Years of Experience (shared) ──
  dash.setRowHeight(row, 14);
  var q2Header = dash.getRange(row, 1, 1, 4);
  q2Header.merge().setValue('Q2 — Years of ED Experience');
  styleHeader(q2Header);
  var q2HeaderB = dash.getRange(row, 5, 1, 4);
  q2HeaderB.merge().setValue('Q2 — Years of ED Experience');
  styleHeader(q2HeaderB);
  row++;

  var expMap = { 'lt1': '< 1 year', '1-3': '1-3 years', '4-7': '4-7 years', '8+': '8+ years' };
  var expOrder = ['lt1', '1-3', '4-7', '8+'];
  var expA = countCol(dataA, 3); // col index 3 = Q2
  var expB = countCol(dataB, 3);

  expOrder.forEach(function(key) {
    var label = expMap[key] || key;
    var cntA = expA[key] || 0;
    var cntB = expB[key] || 0;
    var pctA = totalA > 0 ? Math.round(cntA / totalA * 100) + '%' : '0%';
    var pctB = totalB > 0 ? Math.round(cntB / totalB * 100) + '%' : '0%';
    dash.setRowHeight(row, 18);
    dash.getRange(row, 2).setValue(label).setFontSize(10);
    dash.getRange(row, 3).setValue(cntA).setFontSize(10).setHorizontalAlignment('center');
    dash.getRange(row, 4).setValue(pctA).setFontSize(10).setHorizontalAlignment('center').setFontColor('#555');
    dash.getRange(row, 6).setValue(label).setFontSize(10);
    dash.getRange(row, 7).setValue(cntB).setFontSize(10).setHorizontalAlignment('center');
    dash.getRange(row, 8).setValue(pctB).setFontSize(10).setHorizontalAlignment('center').setFontColor('#555');
    row++;
  });
  row++;

  // ── Q3 (Part A): Time to CTAS Assignment ──
  dash.setRowHeight(row, 14);
  var q3Header = dash.getRange(row, 1, 1, 4);
  q3Header.merge().setValue('Q3 — Time to CTAS Assignment');
  styleHeader(q3Header);
  row++;
  var q3Counts = countCol(dataA, 4);
  var q3Order = ['lt1', '1-2', '2-5', '5-10', 'gt10'];
  var q3Map = { 'lt1': '< 1 min', '1-2': '1-2 min', '2-5': '2-5 min', '5-10': '5-10 min', 'gt10': '> 10 min' };
  q3Order.forEach(function(key) {
    var cnt = q3Counts[key] || 0;
    var pct = totalA > 0 ? Math.round(cnt / totalA * 100) + '%' : '0%';
    dash.setRowHeight(row, 18);
    dash.getRange(row, 2).setValue(q3Map[key] || key).setFontSize(10);
    dash.getRange(row, 3).setValue(cnt).setFontSize(10).setHorizontalAlignment('center');
    dash.getRange(row, 4).setValue(pct).setFontSize(10).setHorizontalAlignment('center').setFontColor('#555');
    row++;
  });
  row++;

  // ── QB3 (Part B): Frequency of Re-assessment ──
  var bRow = row - (q3Order.length + 2); // align with Q3 section start
  // We'll write Part B Q3 in the right column starting from same row as Part A Q3
  // (already written above, so we need to track separately)
  // Instead, write QB3 starting now on the right side
  var rightRow = row - (q3Order.length + 2);
  dash.setRowHeight(rightRow, 14);
  var qb3Header = dash.getRange(rightRow, 5, 1, 4);
  qb3Header.merge().setValue('QB3 — Frequency of Re-assessment');
  styleHeader(qb3Header);
  rightRow++;
  var qb3Counts = countCol(dataB, 4);
  var qb3Order = ['always', 'usually', 'sometimes', 'rarely', 'never'];
  var qb3Map = { 'always': 'Always', 'usually': 'Usually', 'sometimes': 'Sometimes', 'rarely': 'Rarely', 'never': 'Never' };
  qb3Order.forEach(function(key) {
    var cnt = qb3Counts[key] || 0;
    var pct = totalB > 0 ? Math.round(cnt / totalB * 100) + '%' : '0%';
    dash.setRowHeight(rightRow, 18);
    dash.getRange(rightRow, 6).setValue(qb3Map[key] || key).setFontSize(10);
    dash.getRange(rightRow, 7).setValue(cnt).setFontSize(10).setHorizontalAlignment('center');
    dash.getRange(rightRow, 8).setValue(pct).setFontSize(10).setHorizontalAlignment('center').setFontColor('#555');
    rightRow++;
  });

  // ── Q4 (Part A): Hardest Triage Challenge ──
  dash.setRowHeight(row, 14);
  var q4Header = dash.getRange(row, 1, 1, 4);
  q4Header.merge().setValue('Q4 — Hardest Triage Challenge');
  styleHeader(q4Header);
  row++;
  var q4Counts = countCol(dataA, 5);
  var q4Order = ['vitals', 'pain', 'communication', 'overcrowding', 'experience'];
  var q4Map = {
    'vitals': 'Abnormal vitals', 'pain': 'Pain assessment',
    'communication': 'Communication barrier', 'overcrowding': 'Overcrowding',
    'experience': 'Limited experience'
  };
  q4Order.forEach(function(key) {
    var cnt = q4Counts[key] || 0;
    var pct = totalA > 0 ? Math.round(cnt / totalA * 100) + '%' : '0%';
    dash.setRowHeight(row, 18);
    dash.getRange(row, 2).setValue(q4Map[key] || key).setFontSize(10);
    dash.getRange(row, 3).setValue(cnt).setFontSize(10).setHorizontalAlignment('center');
    dash.getRange(row, 4).setValue(pct).setFontSize(10).setHorizontalAlignment('center').setFontColor('#555');
    row++;
  });
  row++;

  // ── QB4 (Part B): CTAS Mismatch Frequency ──
  var qb4StartRow = row - (q4Order.length + 2);
  dash.setRowHeight(qb4StartRow, 14);
  var qb4Header = dash.getRange(qb4StartRow, 5, 1, 4);
  qb4Header.merge().setValue('QB4 — CTAS Mismatch Frequency');
  styleHeader(qb4Header);
  qb4StartRow++;
  var qb4Counts = countCol(dataB, 5);
  var qb4Order = ['daily', 'weekly', 'monthly', 'rarely', 'never'];
  var qb4Map = { 'daily': 'Daily', 'weekly': 'Weekly', 'monthly': 'Monthly', 'rarely': 'Rarely', 'never': 'Never' };
  qb4Order.forEach(function(key) {
    var cnt = qb4Counts[key] || 0;
    var pct = totalB > 0 ? Math.round(cnt / totalB * 100) + '%' : '0%';
    dash.setRowHeight(qb4StartRow, 18);
    dash.getRange(qb4StartRow, 6).setValue(qb4Map[key] || key).setFontSize(10);
    dash.getRange(qb4StartRow, 7).setValue(cnt).setFontSize(10).setHorizontalAlignment('center');
    dash.getRange(qb4StartRow, 8).setValue(pct).setFontSize(10).setHorizontalAlignment('center').setFontColor('#555');
    qb4StartRow++;
  });

  // ── Q5 (Part A): Times Unsure Last Month ──
  dash.setRowHeight(row, 14);
  var q5Header = dash.getRange(row, 1, 1, 4);
  q5Header.merge().setValue('Q5 — Times Unsure (Last Month)');
  styleHeader(q5Header);
  row++;
  var q5Counts = countCol(dataA, 6);
  var q5Order = ['0', '1-2', '3-5', '6-10', 'gt10'];
  var q5Map = { '0': '0 times', '1-2': '1-2 times', '3-5': '3-5 times', '6-10': '6-10 times', 'gt10': '> 10 times' };
  q5Order.forEach(function(key) {
    var cnt = q5Counts[key] || 0;
    var pct = totalA > 0 ? Math.round(cnt / totalA * 100) + '%' : '0%';
    dash.setRowHeight(row, 18);
    dash.getRange(row, 2).setValue(q5Map[key] || key).setFontSize(10);
    dash.getRange(row, 3).setValue(cnt).setFontSize(10).setHorizontalAlignment('center');
    dash.getRange(row, 4).setValue(pct).setFontSize(10).setHorizontalAlignment('center').setFontColor('#555');
    row++;
  });
  row++;

  // ── QB6 (Part B): Action When Level Seems Wrong ──
  var qb6StartRow = row - (q5Order.length + 2);
  dash.setRowHeight(qb6StartRow, 14);
  var qb6Header = dash.getRange(qb6StartRow, 5, 1, 4);
  qb6Header.merge().setValue('QB6 — Action When Level Seems Wrong');
  styleHeader(qb6Header);
  qb6StartRow++;
  var qb6Counts = countCol(dataB, 7);
  var qb6Order = ['reassess', 'discuss', 'override', 'document', 'nothing'];
  var qb6Map = {
    'reassess': 'Reassess patient', 'discuss': 'Discuss with nurse',
    'override': 'Override level', 'document': 'Document concern',
    'nothing': 'Nothing / accept'
  };
  qb6Order.forEach(function(key) {
    var cnt = qb6Counts[key] || 0;
    var pct = totalB > 0 ? Math.round(cnt / totalB * 100) + '%' : '0%';
    dash.setRowHeight(qb6StartRow, 18);
    dash.getRange(qb6StartRow, 6).setValue(qb6Map[key] || key).setFontSize(10);
    dash.getRange(qb6StartRow, 7).setValue(cnt).setFontSize(10).setHorizontalAlignment('center');
    dash.getRange(qb6StartRow, 8).setValue(pct).setFontSize(10).setHorizontalAlignment('center').setFontColor('#555');
    qb6StartRow++;
  });

  // ── Q7 (Part A): Used Aid for CTAS Decision ──
  dash.setRowHeight(row, 14);
  var q7Header = dash.getRange(row, 1, 1, 4);
  q7Header.merge().setValue('Q7 — Used Aid for CTAS Decision');
  styleHeader(q7Header);
  row++;
  var q7Counts = countCol(dataA, 8);
  [['yes-always', 'Yes, always'], ['yes-occ', 'Yes, occasionally'], ['no', 'No'], ['want', 'No, but want one']].forEach(function(pair) {
    var cnt = q7Counts[pair[0]] || 0;
    var pct = totalA > 0 ? Math.round(cnt / totalA * 100) + '%' : '0%';
    dash.setRowHeight(row, 18);
    dash.getRange(row, 2).setValue(pair[1]).setFontSize(10);
    dash.getRange(row, 3).setValue(cnt).setFontSize(10).setHorizontalAlignment('center');
    dash.getRange(row, 4).setValue(pct).setFontSize(10).setHorizontalAlignment('center').setFontColor('#555');
    row++;
  });
  row++;

  // ── QB7 (Part B): Most Common Downstream Effect ──
  var qb7StartRow = row - 6;
  dash.setRowHeight(qb7StartRow, 14);
  var qb7Header = dash.getRange(qb7StartRow, 5, 1, 4);
  qb7Header.merge().setValue('QB7 — Most Common Downstream Effect');
  styleHeader(qb7Header);
  qb7StartRow++;
  var qb7Counts = countCol(dataB, 8);
  var qb7Order = ['delay', 'overcrowd', 'dissatisfy', 'safety', 'none'];
  var qb7Map = {
    'delay': 'Treatment delay', 'overcrowd': 'Overcrowding',
    'dissatisfy': 'Patient dissatisfaction', 'safety': 'Safety incident',
    'none': 'No significant effect'
  };
  qb7Order.forEach(function(key) {
    var cnt = qb7Counts[key] || 0;
    var pct = totalB > 0 ? Math.round(cnt / totalB * 100) + '%' : '0%';
    dash.setRowHeight(qb7StartRow, 18);
    dash.getRange(qb7StartRow, 6).setValue(qb7Map[key] || key).setFontSize(10);
    dash.getRange(qb7StartRow, 7).setValue(cnt).setFontSize(10).setHorizontalAlignment('center');
    dash.getRange(qb7StartRow, 8).setValue(pct).setFontSize(10).setHorizontalAlignment('center').setFontColor('#555');
    qb7StartRow++;
  });

  // ── Q12 (Part A): Decision Support Rating ──
  dash.setRowHeight(row, 14);
  var q12Header = dash.getRange(row, 1, 1, 4);
  q12Header.merge().setValue('Q12 — Want Decision Support Tool (1-5 scale)');
  styleHeader(q12Header);
  row++;
  var q12Avg = avgCol(dataA, 13);
  var q12Counts = countCol(dataA, 13);
  ['1','2','3','4','5'].forEach(function(v) {
    var cnt = q12Counts[v] || 0;
    var pct = totalA > 0 ? Math.round(cnt / totalA * 100) + '%' : '0%';
    dash.setRowHeight(row, 18);
    dash.getRange(row, 2).setValue('Rating ' + v).setFontSize(10);
    dash.getRange(row, 3).setValue(cnt).setFontSize(10).setHorizontalAlignment('center');
    dash.getRange(row, 4).setValue(pct).setFontSize(10).setHorizontalAlignment('center').setFontColor('#555');
    row++;
  });
  dash.setRowHeight(row, 18);
  dash.getRange(row, 2).setValue('Average Rating').setFontSize(10).setFontWeight('bold');
  dash.getRange(row, 3, 1, 2).merge().setValue(q12Avg).setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center').setFontColor(HMG_RED);
  row += 2;

  // ── QB8 (Part B): Confidence with Decision Support ──
  var qb8StartRow = row - 9;
  dash.setRowHeight(qb8StartRow, 14);
  var qb8Header = dash.getRange(qb8StartRow, 5, 1, 4);
  qb8Header.merge().setValue('QB8 — Confidence with Better Decision Support (1-5)');
  styleHeader(qb8Header);
  qb8StartRow++;
  var qb8Avg = avgCol(dataB, 9);
  var qb8Counts = countCol(dataB, 9);
  ['1','2','3','4','5'].forEach(function(v) {
    var cnt = qb8Counts[v] || 0;
    var pct = totalB > 0 ? Math.round(cnt / totalB * 100) + '%' : '0%';
    dash.setRowHeight(qb8StartRow, 18);
    dash.getRange(qb8StartRow, 6).setValue('Rating ' + v).setFontSize(10);
    dash.getRange(qb8StartRow, 7).setValue(cnt).setFontSize(10).setHorizontalAlignment('center');
    dash.getRange(qb8StartRow, 8).setValue(pct).setFontSize(10).setHorizontalAlignment('center').setFontColor('#555');
    qb8StartRow++;
  });
  dash.setRowHeight(qb8StartRow, 18);
  dash.getRange(qb8StartRow, 6).setValue('Average Rating').setFontSize(10).setFontWeight('bold');
  dash.getRange(qb8StartRow, 7, 1, 2).merge().setValue(qb8Avg).setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center').setFontColor(HMG_RED);

  // ── Footer ──
  row++;
  dash.setRowHeight(row, 30);
  var footer = dash.getRange(row, 1, 1, 9);
  footer.merge();
  footer.setValue('Dashboard auto-generated by TriagePluse Apps Script  |  HMG Takhassusi Hospital — Emergency Department  |  HMG/QID/1397');
  footer.setBackground(HMG_NAVY)
        .setFontColor('rgba(255,255,255,0.6)')
        .setFontSize(9)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setItalic(true);

  // ── Add Charts ──
  if (total > 0) {
    addCharts(ss, dash, dataA, dataB, totalA, totalB);
  }

  Logger.log('Dashboard built successfully at ' + new Date().toLocaleString());
}

// ─── Chart Builder ────────────────────────────────────────────────────────────
function addCharts(ss, dash, dataA, dataB, totalA, totalB) {
  // Chart 1: Response distribution (Part A vs Part B) — Pie chart
  var pieData = dash.getRange(4, 2, 1, 2); // dummy range for anchor
  var pieChart = dash.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(buildChartRange(ss, [['Role', 'Responses'], ['Triage Nurses (A)', totalA], ['Physicians/Managers (B)', totalB]]))
    .setPosition(4, 10, 0, 0)
    .setOption('title', 'Respondents by Role')
    .setOption('width', 380)
    .setOption('height', 260)
    .setOption('colors', ['#0FA89B', '#E8A020'])
    .setOption('pieHole', 0.4)
    .setOption('legend', { position: 'bottom' })
    .build();
  dash.insertChart(pieChart);

  // Chart 2: Years of Experience (Part A) — Bar chart
  var expLabels = ['< 1 year', '1-3 years', '4-7 years', '8+ years'];
  var expKeys   = ['lt1', '1-3', '4-7', '8+'];
  function countCol(data, colIdx) {
    var counts = {};
    data.forEach(function(row) {
      var val = (row[colIdx] || '').toString().trim();
      if (val) counts[val] = (counts[val] || 0) + 1;
    });
    return counts;
  }
  var expA = countCol(dataA, 3);
  var expB = countCol(dataB, 3);
  var expChartData = [['Experience', 'Part A (Nurses)', 'Part B (Physicians)']];
  expKeys.forEach(function(k, i) {
    expChartData.push([expLabels[i], expA[k] || 0, expB[k] || 0]);
  });
  var expChart = dash.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(buildChartRange(ss, expChartData))
    .setPosition(18, 10, 0, 0)
    .setOption('title', 'Years of ED Experience')
    .setOption('width', 380)
    .setOption('height', 260)
    .setOption('colors', ['#0FA89B', '#E8A020'])
    .setOption('legend', { position: 'bottom' })
    .build();
  dash.insertChart(expChart);

  // Chart 3: Q12 Decision Support Rating (Part A) — Column chart
  if (totalA > 0) {
    var q12Counts = countCol(dataA, 13);
    var q12Data = [['Rating', 'Count']];
    ['1','2','3','4','5'].forEach(function(v) { q12Data.push(['Rating ' + v, q12Counts[v] || 0]); });
    var q12Chart = dash.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(buildChartRange(ss, q12Data))
      .setPosition(32, 10, 0, 0)
      .setOption('title', 'Q12 — Want Decision Support Tool (Nurses)')
      .setOption('width', 380)
      .setOption('height', 260)
      .setOption('colors', ['#1A2B4A'])
      .setOption('legend', { position: 'none' })
      .build();
    dash.insertChart(q12Chart);
  }
}

// Helper: create a temporary data range for charts from a 2D array
function buildChartRange(ss, data) {
  var tempSheet = ss.getSheetByName('_ChartData');
  if (!tempSheet) {
    tempSheet = ss.insertSheet('_ChartData');
  } else {
    tempSheet.clearContents();
  }
  tempSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  return tempSheet.getRange(1, 1, data.length, data[0].length);
}

// ─── Setup Function (run once manually) ──────────────────────────────────────
function setupSheets() {
  var ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  // Rename default "Sheet1" if it exists
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) defaultSheet.setName('Summary');

  // Create data sheets
  createSheetWithHeaders(ss, PART_A_SHEET, HEADERS_A);
  createSheetWithHeaders(ss, PART_B_SHEET, HEADERS_B);

  // Build the dashboard
  buildDashboard();

  Logger.log('Setup complete! All sheets and dashboard created.');
}
