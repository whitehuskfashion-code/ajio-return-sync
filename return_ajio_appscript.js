// ================================================================
// alerts.gs  —  AJIO Return Alerts (Apps Script)
// Paste entire file into Extensions → Apps Script
// ================================================================

const MAIN_SHEET    = "AJIO_RETURN";

// ── NEW column layout (1-based) ──────────────────────────────
const COL_SELLER_SKU       =  3;
const COL_RETURN_QTY       =  4;
const COL_CUST_ORDER       =  5;
const COL_RETURN_CREATED   =  6;   // F  
const COL_RETURN_DELIVERED =  8;   // H  
const COL_RETURN_AWB       =  9;   // I  
const COL_ACTUAL_DELIVERED = 10;   // J  manual
const COL_QUALITY          = 11;   // K  manual
const COL_CARRIER          = 13;
const COL_RETURN_ORDER_NUM = 15;
const TOTAL_COLS           = 15;

const PURPLE = "#9C2BE6";   // Alert 1 — Delivered Not Received
const ORANGE = "#FF9900";   // Alert 2 — 61+ Days No Actual Delivery

// ================================================================
// ENTRY POINTS
// ================================================================

function doPost(e) {
  try {
    runAlerts();
    return ContentService.createTextOutput("OK");
  } catch(err) {
    Logger.log("doPost error: " + err.message);
    return ContentService.createTextOutput("ERROR: " + err.message);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 AJIO Alerts")
    .addItem("▶  Run Alerts / Check Statuses", "runAlerts")
    .addSeparator()
    .addItem("🎨  Refresh All Highlights",    "refreshAllHighlights")
    .addToUi();
}

// ================================================================
// MAIN ALERT LOGIC
// ================================================================

function runAlerts() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainWs    = ss.getSheetByName(MAIN_SHEET);
  
  if (!mainWs) {
    Logger.log("ERROR: Required sheet not found");
    return;
  }

  const dataRows = mainWs.getDataRange().getValues().slice(1);
  updateCancelledReturns(mainWs, dataRows);
  refreshAllHighlights();
  
  Logger.log("runAlerts complete ✅");
}

function updateCancelledReturns(mainWs, dataRows) {
  const updates = [];
  for (let i = 0; i < dataRows.length; i++) {
    const qty = String(dataRows[i][COL_RETURN_QTY - 1] ?? "").trim();
    const quality = String(dataRows[i][COL_QUALITY - 1] || "").trim();
    
    if (qty === "0" && quality !== "Return Cancelled") {
      updates.push({ row: i + 2, col: COL_QUALITY, val: "Return Cancelled" });
      dataRows[i][COL_QUALITY - 1] = "Return Cancelled"; // update in memory so downstream skips it
    } else if (qty !== "0" && quality === "Return Cancelled") {
      // Revert back to Select if QTY is no longer 0
      updates.push({ row: i + 2, col: COL_QUALITY, val: "Select" });
      dataRows[i][COL_QUALITY - 1] = "Select";
    }
  }
  if (updates.length > 0) {
    updates.forEach(u => mainWs.getRange(u.row, u.col).setValue(u.val));
    Logger.log("Updated " + updates.length + " Quality statuses.");
  }
}

// ================================================================
// onEdit — instant highlight cleanup
// ================================================================

function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== MAIN_SHEET) return;

  const editedCols = [];
  for (let j = 0; j < e.range.getNumColumns(); j++) {
    editedCols.push(e.range.getColumn() + j);
  }
  
  // If either Actual Delivered Date (J) or Quality (K) is edited, refresh highlights
  if (editedCols.includes(COL_ACTUAL_DELIVERED) || editedCols.includes(COL_QUALITY)) {
    refreshAllHighlights();
  }
}

// ================================================================
// HIGHLIGHTS (Single Source of Truth)
// ================================================================

function refreshAllHighlights() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainWs = ss.getSheetByName(MAIN_SHEET);
  if (!mainWs) return;

  const lastRow = mainWs.getLastRow();
  if (lastRow < 2) return;

  // Single batch read
  const allData = mainWs.getRange(2, 1, lastRow - 1, TOTAL_COLS).getValues();
  const colours = [];

  const today = new Date(); 
  today.setHours(0,0,0,0);
  
  const wEnd = new Date(today); 
  wEnd.setDate(today.getDate() - 3);
  
  const cutoff61 = new Date(today); 
  cutoff61.setDate(today.getDate() - 61);

  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const ron = String(row[COL_RETURN_ORDER_NUM - 1] || "").trim();
    const actualVal = String(row[COL_ACTUAL_DELIVERED - 1] || "").trim();
    const qty = String(row[COL_RETURN_QTY - 1] ?? "").trim();
    const quality = String(row[COL_QUALITY - 1] || "").trim();
    
    let colour = null; // default: clear

    // Only consider highlighting if RON exists, QTY != 0, and item is NOT physically received
    if (ron && qty !== "0" && actualVal === "") {
      const delivDate = parseDate(row[COL_RETURN_DELIVERED - 1]);
      const createdDt = parseDate(row[COL_RETURN_CREATED - 1]);

      // Check Purple Condition
      if (delivDate && delivDate <= wEnd && quality !== "Non Delivered 7 days") {
        colour = PURPLE;
      } 
      // Check Orange Condition (Only if not already Purple)
      else if (createdDt && createdDt < cutoff61 && quality !== "Non Delivered 60+ days") {
        colour = ORANGE;
      }
    }

    colours.push(Array(TOTAL_COLS).fill(colour));
  }

  // Single batch write
  mainWs.getRange(2, 1, allData.length, TOTAL_COLS).setBackgrounds(colours);
  Logger.log(`refreshAllHighlights done ✅ — ${allData.length} rows processed`);
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
