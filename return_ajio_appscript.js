// ================================================================
// alerts.gs  —  AJIO Return Alerts (Apps Script)
// Paste entire file into Extensions → Apps Script
// ================================================================

const MAIN_SHEET    = "AJIO_RETURN";
const TRACKER_SHEET = "AJIO_TICKETS";

// ── NEW column layout (1-based) ──────────────────────────────
// A  LAST_UPDATE_DATE       1
// B  Return Type             2
// C  SELLER SKU              3
// D  Return QTY              4
// E  Cust Order No           5
// F  Return Created Date     6   ← was G(7)
// G  3PL Delivery Status     7
// H  Return Delivered Date   8   ← was I(9)
// I  Return AWB No           9   ← was F(6)
// J  Actual Delivered Date  10   manual — never written by sync
// K  Quality                11   manual
// L  Notes                  12   manual
// M  Return Carrier Name    13
// N  Cust Return Reason     14
// O  RETURN ORDER NUMBER    15
const COL_SELLER_SKU       =  3;
const COL_CUST_ORDER       =  5;
const COL_RETURN_CREATED   =  6;   // F  (was 7)
const COL_RETURN_DELIVERED =  8;   // H  (was 9)
const COL_RETURN_AWB       =  9;   // I  (was 6)
const COL_ACTUAL_DELIVERED = 10;   // J  manual — Alert 2 checks this
const COL_CARRIER          = 13;
const COL_RETURN_ORDER_NUM = 15;
const TOTAL_COLS           = 15;

const PURPLE = "#9C2BE6";   // Alert 1 — Delivered Not Received
const ORANGE = "#FF9900";   // Alert 2 — 61+ Days No Actual Delivery

const ALERT_DELIVERED = "DELIVERED_NOT_RECEIVED";
const ALERT_61_DAYS   = "61_DAYS_NO_DELIVERY";

// ================================================================
// ENTRY POINTS
// ================================================================

function doPost(e) {
  try {
    runAlerts(true);
    return ContentService.createTextOutput("OK");
  } catch(err) {
    Logger.log("doPost error: " + err.message);
    return ContentService.createTextOutput("ERROR: " + err.message);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 AJIO Alerts")
    .addItem("▶  Run Alerts (with email)",    "runAlertsWithEmail")
    .addItem("▶  Run Alerts (without email)", "runAlertsNoEmail")
    .addSeparator()
    .addItem("🎨  Refresh All Highlights",    "refreshAllHighlights")
    .addToUi();
}


function runAlertsWithEmail()  { runAlerts(true);  }
function runAlertsNoEmail()    { runAlerts(false); }

// ================================================================
// MAIN ALERT LOGIC
// ================================================================

function runAlerts(sendEmail) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainWs    = ss.getSheetByName(MAIN_SHEET);
  const trackerWs = ss.getSheetByName(TRACKER_SHEET);
  if (!mainWs || !trackerWs) {
    Logger.log("ERROR: Required sheet not found");
    return;
  }

  const today = new Date(); today.setHours(0,0,0,0);
  const dataRows = mainWs.getDataRange().getValues().slice(1);
  const tracker  = readTracker(trackerWs);

  resolveClosedCases(mainWs, trackerWs, dataRows);  // ← ADD THIS

  // STEP 1 — Remove highlights where ticket is already raised
  for (const ron in tracker) {
    if (tracker[ron].ticketDate) {
      const i = findRowByRON(dataRows, ron);
      if (i !== -1) mainWs.getRange(i + 2, 1, 1, TOTAL_COLS).setBackground(null);
    }
  }

  // STEP 2 — Alert 1: Delivered Not Received (PURPLE)
  // Trigger window: Return Delivered Date is older than 3 days
  // AND Actual Delivered Date (col J) is still empty
  const wEnd   = new Date(today); wEnd.setDate(today.getDate() - 3);
  const alert1 = [];

  for (let i = 0; i < dataRows.length; i++) {
    const row       = dataRows[i];
    const delivDate = parseDate(row[COL_RETURN_DELIVERED - 1]);
    const actualVal = String(row[COL_ACTUAL_DELIVERED  - 1] || "").trim();
    const ron       = String(row[COL_RETURN_ORDER_NUM  - 1] || "").trim();

    if (!delivDate || !ron)            continue;
    if (delivDate > wEnd)              continue;   // Only skips if it's LESS than 3 days old
    if (actualVal !== "")              continue;   // already physically received
    const t = tracker[ron];
    if (t && t.ticketDate)             continue;   // ticket already raised

    // mainWs.getRange(i + 2, 1, 1, TOTAL_COLS).setBackground(PURPLE);
    if (!t) {
      appendTracker(trackerWs, ron, ALERT_DELIVERED);
      tracker[ron] = { alertType: ALERT_DELIVERED, ticketDate: "" };
    }
    alert1.push({ ron, data: row });
  }

  if (alert1.length > 0 && sendEmail) sendAlertEmail(alert1, ALERT_DELIVERED);

  // STEP 3 — Alert 2: 61+ Days No Actual Delivery (ORANGE)
  // Trigger: Return Created Date is 61+ days ago AND Actual Delivered Date (col J) is empty
  // Day 59 → no. Day 60 → no. Day 61 → YES. Day 62+ → YES.
  const cutoff61 = new Date(today); cutoff61.setDate(today.getDate() - 61);
  const alert2   = [];

  for (let i = 0; i < dataRows.length; i++) {
    const row       = dataRows[i];
    const createdDt = parseDate(row[COL_RETURN_CREATED   - 1]);
    const actualVal = String(row[COL_ACTUAL_DELIVERED    - 1] || "").trim();
    const ron       = String(row[COL_RETURN_ORDER_NUM    - 1] || "").trim();

    if (!createdDt || !ron)    continue;
    if (createdDt >= cutoff61) continue;   // not old enough yet (< 61 days)
    if (actualVal !== "")      continue;   // already received physically
    const t = tracker[ron];
    if (t && t.ticketDate)     continue;   // ticket already raised

    //mainWs.getRange(i + 2, 1, 1, TOTAL_COLS).setBackground(ORANGE);
    if (!t) {
      appendTracker(trackerWs, ron, ALERT_61_DAYS);
      tracker[ron] = { alertType: ALERT_61_DAYS, ticketDate: "" };
    }
    alert2.push({ ron, data: row });
  }

  if (alert2.length > 0 && sendEmail) sendAlertEmail(alert2, ALERT_61_DAYS);

  // ✅ Single source of truth for all highlights
  refreshAllHighlights();
  Logger.log("runAlerts complete ✅");
}

// ================================================================
// onEdit — instant highlight cleanup / re-flag when
//          TICKET_CREATED_DATE in AJIO_TICKETS is changed
// ================================================================

function onEditTickets(e) {
  if (!e || !e.range) return;
  const sh     = e.range.getSheet();
  const shName = sh.getName();

  // ── Case 1: Col J edited in AJIO_RETURN ──
  if (shName === MAIN_SHEET) {
    const editedCols = [];
    for (let j = 0; j < e.range.getNumColumns(); j++) {
      editedCols.push(e.range.getColumn() + j);
    }
    if (!editedCols.includes(COL_ACTUAL_DELIVERED)) return;

    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const trackerWs = ss.getSheetByName(TRACKER_SHEET);
    const mainWs    = sh;

    // For each edited row, check col J and sync AJIO_TICKETS
    for (let i = 0; i < e.range.getNumRows(); i++) {
      const row       = e.range.getRow() + i;
      const ron       = String(mainWs.getRange(row, COL_RETURN_ORDER_NUM).getValue() || "").trim();
      const actualVal = String(mainWs.getRange(row, COL_ACTUAL_DELIVERED).getValue() || "").trim();
      if (!ron) continue;

      const tracker = readTracker(trackerWs);

      if (actualVal === "") {
        // Col J cleared → ensure RON is in AJIO_TICKETS
        if (!tracker[ron]) {
          // Determine alert type from dates
          const createdDt  = parseDate(mainWs.getRange(row, COL_RETURN_CREATED).getValue());
          const delivDate  = parseDate(mainWs.getRange(row, COL_RETURN_DELIVERED).getValue());
          const today      = new Date(); today.setHours(0,0,0,0);
          const cutoff61   = new Date(today); cutoff61.setDate(today.getDate() - 61);
          const wEnd       = new Date(today); wEnd.setDate(today.getDate() - 3);

          let alertType = null;
          if (delivDate && delivDate <= wEnd) alertType = ALERT_DELIVERED;
          else if (createdDt && createdDt < cutoff61)               alertType = ALERT_61_DAYS;

          if (alertType) appendTracker(trackerWs, ron, alertType);
        }
      } else {
        // Col J filled → check if ALL rows for this RON have col J filled
        // If yes, remove from AJIO_TICKETS
        const allData   = mainWs.getDataRange().getValues().slice(1);
        const allFilled = allData
          .filter(r => String(r[COL_RETURN_ORDER_NUM - 1] || "").trim() === ron)
          .every(r  => String(r[COL_ACTUAL_DELIVERED - 1] || "").trim() !== "");

        if (allFilled) {
          deleteTrackerEntry(trackerWs, ron);
        }
      }
    }

    refreshAllHighlights();
    return;
  }

  // ── Case 2: TICKET_CREATED_DATE edited in AJIO_TICKETS ──
  if (shName === TRACKER_SHEET) {
    const headers   = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const ticketCol = headers.indexOf("TICKET_CREATED_DATE") + 1;
    if (!ticketCol) return;

    const editedCols = [];
    for (let j = 0; j < e.range.getNumColumns(); j++) {
      editedCols.push(e.range.getColumn() + j);
    }
    if (!editedCols.includes(ticketCol)) return;

    refreshAllHighlights();
  }
}

function deleteTrackerEntry(trackerWs, ron) {
  const vals   = trackerWs.getDataRange().getValues();
  const hdr    = vals[0];
  const ronCol = hdr.indexOf("RETURN_ORDER_NUMBER");
  if (ronCol < 0) return;

  // Loop bottom-up so row deletion doesn't shift indices
  for (let i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][ronCol] || "").trim() === ron) {
      trackerWs.deleteRow(i + 1);
    }
  }
}

// ================================================================
// TRACKER HELPERS
// ================================================================

function readTracker(ws) {
  const result = {};
  const vals   = ws.getDataRange().getValues();
  if (vals.length <= 1) return result;
  const hdr     = vals[0];
  const iRon    = hdr.indexOf("RETURN_ORDER_NUMBER");
  const iAlert  = hdr.indexOf("ALERT_TYPE");
  const iTicket = hdr.indexOf("TICKET_CREATED_DATE");
  if (iRon < 0) { Logger.log("ERROR: AJIO_TICKETS headers missing"); return result; }
  for (let i = 1; i < vals.length; i++) {
    const r   = vals[i];
    const ron = String(r[iRon] || "").trim();
    if (!ron) continue;
    result[ron] = {
      alertType:  String(iAlert  >= 0 ? r[iAlert]  : "").trim(),
      ticketDate: String(iTicket >= 0 ? r[iTicket] : "").trim(),
    };
  }
  return result;
}

function appendTracker(ws, ron, alertType) {
  const d = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  ws.appendRow([ron, alertType, d, ""]);
}

function findRowByRON(dataRows, ron) {
  for (let i = 0; i < dataRows.length; i++) {
    if (String(dataRows[i][COL_RETURN_ORDER_NUM - 1] || "").trim() === ron) return i;
  }
  return -1;
}

// ================================================================
// DATE PARSER
// ================================================================

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val)) {
    const d = new Date(val); d.setHours(0,0,0,0); return d;
  }
  const s = String(val).trim();
  if (!s) return null;
  // dd-mm-yyyy HH:MM  or  dd-mm-yyyy
  const m1 = s.match(/^(\d{2})-(\d{2})-(\d{4})/);
  if (m1) { const d = new Date(+m1[3], +m1[2]-1, +m1[1]); d.setHours(0,0,0,0); return d; }
  // yyyy-mm-dd
  const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m2) { const d = new Date(+m2[1], +m2[2]-1, +m2[3]); d.setHours(0,0,0,0); return d; }
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

// ================================================================
// EMAIL  (MailApp — no password needed)
// ================================================================

function sendAlertEmail(rows, alertType) {
  const props   = PropertiesService.getScriptProperties();
  const emailTo = props.getProperty("EMAIL_TO") || "";
  if (!emailTo) { Logger.log("EMAIL_TO not set in Script Properties!"); return; }

  const tz = Session.getScriptTimeZone();
  let subject, bodyIntro, colNums, colHeaders;

  if (alertType === ALERT_DELIVERED) {
    subject    = "Ajio Order Marked Delivered but Not Received – Raise Ticket Immediately";
    bodyIntro  = "The following order(s) are marked as delivered in the Ajio system; "
               + "however, they have not been received by us.<br><br>"
               + "Kindly investigate and raise a ticket at the earliest.";
    colNums    = [COL_SELLER_SKU, COL_CUST_ORDER, COL_RETURN_DELIVERED, COL_RETURN_AWB, COL_CARRIER, COL_RETURN_ORDER_NUM];
    colHeaders = ["SELLER SKU", "Cust Order No", "Return Delivered Date", "Return AWB No", "Return Carrier Name", "RETURN ORDER NUMBER"];
  } else {
    subject    = "Ajio Alert – 61+ Days Passed, Return Not Received";
    bodyIntro  = "It has been over 61 days since the Return Creation Date, yet the "
               + "Actual Delivered Date is still not updated.<br><br>"
               + "Kindly raise a ticket immediately.";
    colNums    = [COL_SELLER_SKU, COL_CUST_ORDER, COL_RETURN_CREATED, COL_RETURN_AWB, COL_CARRIER, COL_RETURN_ORDER_NUM];
    colHeaders = ["SELLER SKU", "Cust Order No", "Return Created Date", "Return AWB No", "Return Carrier Name", "RETURN ORDER NUMBER"];
  }

  const th = 'style="border:1px solid #ccc;padding:8px;background:#f0f0f0;font-family:Arial,sans-serif;font-size:13px"';
  const td = 'style="border:1px solid #ccc;padding:8px;font-family:Arial,sans-serif;font-size:13px"';

  let tableHtml = `<table style="border-collapse:collapse"><thead><tr>`;
  colHeaders.forEach(h => tableHtml += `<th ${th}>${h}</th>`);
  tableHtml += `</tr></thead><tbody>`;

  const csvLines = [colHeaders.join(",")];

  rows.forEach(item => {
    tableHtml += "<tr>";
    const csvRow = [];
    colNums.forEach(cn => {
      const raw = item.data[cn - 1];
      const val = raw instanceof Date
        ? Utilities.formatDate(raw, tz, "dd-MM-yyyy HH:mm")
        : String(raw || "").trim();
      tableHtml += `<td ${td}>${val}</td>`;
      csvRow.push(`"${String(val).replace(/"/g, '""')}"`);
    });
    tableHtml += "</tr>";
    csvLines.push(csvRow.join(","));   // ← real \n newline, not \\n
  });

  tableHtml += "</tbody></table>";

  const csvBlob = Utilities.newBlob(
    csvLines.join("\n"),   // ← real newline character
    "text/csv",
    `alert_${alertType.toLowerCase()}.csv`
  );

  MailApp.sendEmail({
    to:          emailTo,
    subject:     subject,
    htmlBody:    `<p>${bodyIntro}</p><br>${tableHtml}`,
    attachments: [csvBlob],
  });

  Logger.log("Email sent: " + subject);
}

function refreshAllHighlights() {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const mainWs = ss.getSheetByName(MAIN_SHEET);
const trackerWs = ss.getSheetByName(TRACKER_SHEET);
if (!mainWs || !trackerWs) { Logger.log("Sheet not found"); return; }

const tracker = readTracker(trackerWs);
const lastRow = mainWs.getLastRow();
if (lastRow < 2) return;

// ✅ Single batch read — all data at once
const allData = mainWs.getRange(2, 1, lastRow - 1, TOTAL_COLS).getValues();

// Build a 2D colour array (one row per data row, one col per column)
// null = clear, colour string = highlight
const colours = [];

for (let i = 0; i < allData.length; i++) {
const row = allData[i];
const ron = String(row[COL_RETURN_ORDER_NUM - 1] || "").trim();
const actualVal = String(row[COL_ACTUAL_DELIVERED - 1] || "").trim();

let colour = null; // default: clear

if (ron) {
const info = tracker[ron];
if (info && !info.ticketDate && actualVal === "") {
colour = (info.alertType === ALERT_DELIVERED) ? PURPLE : ORANGE;
}
}

// Fill all TOTAL_COLS columns with the same colour
colours.push(Array(TOTAL_COLS).fill(colour));
}

// ✅ Single batch write — one API call instead of N
mainWs.getRange(2, 1, allData.length, TOTAL_COLS).setBackgrounds(colours);

Logger.log(`refreshAllHighlights done ✅ — ${allData.length} rows processed`);
}

function resolveClosedCases(mainWs, trackerWs, dataRows) {
  const lastTrackerRow = trackerWs.getLastRow();
  if (lastTrackerRow < 2) return;

  const trackerVals = trackerWs.getDataRange().getValues();
  const hdr         = trackerVals[0];
  const iRon        = hdr.indexOf("RETURN_ORDER_NUMBER");
  if (iRon < 0) return;

  // Loop every row in AJIO_RETURN individually
  for (let i = 0; i < dataRows.length; i++) {
    const ron       = String(dataRows[i][COL_RETURN_ORDER_NUM - 1] || "").trim();
    const actualVal = String(dataRows[i][COL_ACTUAL_DELIVERED - 1] || "").trim();

    if (!ron || actualVal === "") continue;  // no RON or col J empty → skip

    // Col J filled → remove highlight for THIS row only
    mainWs.getRange(i + 2, 1, 1, TOTAL_COLS).setBackground(null);
  }

  // Delete from AJIO_TICKETS only if ALL rows for that RON have col J filled
  // Build map: ron → are all its rows resolved?
  const ronAllResolved = {};
  for (let i = 0; i < dataRows.length; i++) {
    const ron       = String(dataRows[i][COL_RETURN_ORDER_NUM - 1] || "").trim();
    const actualVal = String(dataRows[i][COL_ACTUAL_DELIVERED - 1] || "").trim();
    if (!ron) continue;
    if (!(ron in ronAllResolved)) ronAllResolved[ron] = true;
    if (actualVal === "") ronAllResolved[ron] = false;  // at least one SKU still open
  }

  // Delete ticket rows bottom-up for fully resolved RONs
  const rowsToDelete = [];
  for (let t = 1; t < trackerVals.length; t++) {
    const ron = String(trackerVals[t][iRon] || "").trim();
    if (ron && ronAllResolved[ron] === true) rowsToDelete.push(t + 1);
  }
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    trackerWs.deleteRow(rowsToDelete[i]);
  }

  Logger.log(`resolveClosedCases: highlights cleared per SKU, ${rowsToDelete.length} ticket(s) deleted`);
}
