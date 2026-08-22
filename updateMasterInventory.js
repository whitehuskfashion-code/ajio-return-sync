/**
 * @OnlyCurrentDoc
 * Employee Inventory Tracker
 * Connects to Print Number and Qty Master Spreadsheet
 */

const MASTER_SPREADSHEET_ID = "1GMfVnYMuXxB-3_m1kuPpBcnuKK6zO02mcg0hmPprAn4";

// IMPORTANT: Replace with actual allowed emails before use.
const ALLOWED_EMAILS = [
  "owner@example.com", 
  "manager@example.com"
];

function checkAuthorization() {
  const userEmail = Session.getActiveUser().getEmail();
  if (userEmail && !ALLOWED_EMAILS.includes(userEmail)) {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Unauthorized Access", `Your email (${userEmail}) does not have permission to run this action.`, ui.ButtonSet.OK);
    return false;
  }
  return true;
}


// =====================================================================
// 1. Weekly Stock Count Logic
// =====================================================================

function replaceWeeklyStock() {
  if (!checkAuthorization()) return;

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const weeklySheet = ss.getSheetByName("WeeklyStockCount");
  const reviewSheet = ss.getSheetByName("ReviewWeeklyStockUpdates");

  if (!weeklySheet || !reviewSheet) {
    ui.alert("Error", "Missing sheets: 'WeeklyStockCount' or 'ReviewWeeklyStockUpdates'", ui.ButtonSet.OK);
    return;
  }

  const rawData = weeklySheet.getRange(2, 1, Math.max(1, weeklySheet.getLastRow() - 1), 7).getValues(); // A to G

  // 1a. Find the Active Slot (Latest Date)
  let latestDateValue = 0;
  let activeSlot = null;

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const color = String(row[1]).trim();
    const dateVal = row[2];

    if (color !== "" && dateVal instanceof Date) {
      if (dateVal.getTime() > latestDateValue) {
        latestDateValue = dateVal.getTime();
        activeSlot = String(row[0]).trim();
      }
    }
  }

  if (!activeSlot) {
    ui.alert("Error", "Could not identify an active slot. Make sure you have entered valid dates in the 'Last Count Date' column.", ui.ButtonSet.OK);
    return;
  }

  // 1b. Validate Active Slot Data
  const activeRows = [];
  const colorSet = new Set();
  let latestDateStr = new Date(latestDateValue).toDateString();

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const slot = String(row[0]).trim();
    const color = String(row[1]).trim();
    const dateVal = row[2];

    if (slot === activeSlot && color !== "") {
      // Date Check - All dates in the active slot must exactly match
      if (!(dateVal instanceof Date) || dateVal.toDateString() !== latestDateStr) {
        ui.alert("Error", `Please ensure all dates are exactly the same for slot '${activeSlot}'.`, ui.ButtonSet.OK);
        return;
      }

      // Duplicate Color Check
      const lowerColor = color.toLowerCase();
      if (colorSet.has(lowerColor)) {
        ui.alert("Error", `Duplicate color '${color}' found in ${activeSlot}. Please consolidate into one row.`, ui.ButtonSet.OK);
        return;
      }
      colorSet.add(lowerColor);

      // Validate Numbers & Blanks
      const sizes = [row[3], row[4], row[5], row[6]];
      const sizeNames = ["S", "M", "L", "XL"];
      for (let s = 0; s < 4; s++) {
        if (sizes[s] === "") {
          ui.alert("Error", `Please fill correct count. Size '${sizeNames[s]}' is missing for color '${color}'.\nType 0 if empty.`, ui.ButtonSet.OK);
          return;
        }
        if (typeof sizes[s] !== 'number') {
          ui.alert("Error", `Invalid number format in Size '${sizeNames[s]}' for color '${color}'. Please enter digits only.`, ui.ButtonSet.OK);
          return;
        }
      }

      activeRows.push({
        rowNum: i + 2,
        color: color,
        S: sizes[0],
        M: sizes[1],
        L: sizes[2],
        XL: sizes[3]
      });
    }
  }

  if (activeRows.length === 0) {
    ui.alert("Notice", "No valid rows found to process in active slot.", ui.ButtonSet.OK);
    return;
  }

  ss.toast("Connecting to Master Inventory...", "Processing");

  // 1c. Fetch Master Inventory
  let remoteSs, masterSheet;
  try {
    remoteSs = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    masterSheet = remoteSs.getSheetByName("master_inventory");
  } catch(e) {
    ui.alert("Error", "Could not connect to remote master spreadsheet.", ui.ButtonSet.OK);
    return;
  }

  const masterData = masterSheet.getRange(2, 1, Math.max(1, masterSheet.getLastRow() - 1), 8).getValues();
  const masterMap = new Map(); // color (lower) -> { rowIdx, S, M, L, XL }

  masterData.forEach((row, idx) => {
    let mColor = String(row[0]).trim();
    if (mColor !== "") {
      masterMap.set(mColor.toLowerCase(), {
        rowIdx: idx, // 0-based relative to data array
        S: Number(row[1]) || 0, // B
        M: Number(row[3]) || 0, // D
        L: Number(row[5]) || 0, // F
        XL: Number(row[7]) || 0 // H
      });
    }
  });

  // Keep a copy of master array to do bulk update later
  const updatedMasterData = masterSheet.getRange(2, 2, Math.max(1, masterSheet.getLastRow() - 1), 7).getValues(); // Fetch B to H

  const reviewOutput = [];
  const reviewHighlights = []; // { idx, color }
  let updatedMasterCount = 0;

  activeRows.forEach(item => {
    let lowerColor = item.color.toLowerCase();
    
    if (!masterMap.has(lowerColor)) {
      reviewOutput.push([item.color, "ALL", "N/A", "N/A", "Color not found in master_inventory"]);
      return; // Skip this color
    }

    const masterStats = masterMap.get(lowerColor);
    const mRowIdx = masterStats.rowIdx;
    const sizesToProcess = [
      { sizeName: "S", counted: item.S, master: masterStats.S, mColIndex: 0 }, // B -> index 0 in updatedMasterData
      { sizeName: "M", counted: item.M, master: masterStats.M, mColIndex: 2 }, // D -> index 2
      { sizeName: "L", counted: item.L, master: masterStats.L, mColIndex: 4 }, // F -> index 4
      { sizeName: "XL", counted: item.XL, master: masterStats.XL, mColIndex: 6 } // H -> index 6
    ];

    sizesToProcess.forEach(s => {
      let diff = s.counted - s.master;
      
      if (Math.abs(diff) <= 1) {
        // Exact match or ±1 Tolerance: Auto-update master
        if (diff !== 0) {
          updatedMasterData[mRowIdx][s.mColIndex] = s.counted;
          updatedMasterCount++;
        }
      } else if (diff > 1) {
        // Overcount (e.g. counted 5, master 2)
        reviewOutput.push([item.color, s.sizeName, s.counted, s.master, `Overcount: Found ${diff} extra`]);
      } else {
        // Undercount (e.g. counted 2, master 5) -> diff is -3
        let missing = Math.abs(diff);
        reviewOutput.push([item.color, s.sizeName, s.counted, s.master, `Undercount: Missing ${missing}`]);
        
        if (missing === 2) {
          reviewHighlights.push({ idx: reviewOutput.length - 1, color: "#ffd966" }); // Light Orange
        } else if (missing > 2) {
          reviewHighlights.push({ idx: reviewOutput.length - 1, color: "#f4cccc" }); // Light Red
        }
      }
    });
  });

  // 1d. Write Data
  ss.toast("Writing updates...", "Processing");

  // Bulk update master inventory
  if (updatedMasterCount > 0) {
    masterSheet.getRange(2, 2, updatedMasterData.length, 7).setValues(updatedMasterData);
  }

  // Write Review Sheet
  if (reviewOutput.length > 0) {
    // Append to bottom
    let startRow = Math.max(2, reviewSheet.getLastRow() + 1);
    reviewSheet.getRange(startRow, 1, reviewOutput.length, 5).setValues(reviewOutput);
    
    // Apply highlights to "Current Counted Inventory" (Column C = 3)
    reviewHighlights.forEach(h => {
      reviewSheet.getRange(startRow + h.idx, 3).setBackground(h.color); 
    });
  }

  // 1e. Clear the OTHER slot
  const clearRanges = [];
  for (let i = 0; i < rawData.length; i++) {
    const slot = String(rawData[i][0]).trim();
    if (slot !== "" && slot !== activeSlot) {
      // Clear columns C to G for this row
      clearRanges.push(`C${i + 2}:G${i + 2}`);
    }
  }

  if (clearRanges.length > 0) {
    weeklySheet.getRangeList(clearRanges).clearContent();
  }

  SpreadsheetApp.flush();
  ui.alert("Success", `Processed Slot: ${activeSlot}\nMaster Inventory Updates: ${updatedMasterCount}\nIssues Logged for Review: ${reviewOutput.length}`, ui.ButtonSet.OK);
}


// =====================================================================
// 2. Review Sheet Updates
// =====================================================================

function updateFromReviewSheet() {
  if (!checkAuthorization()) return;

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("Confirm", "Are you sure you want to forcibly update the Master Inventory with these Review counts?", ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reviewSheet = ss.getSheetByName("ReviewWeeklyStockUpdates");
  
  if (!reviewSheet || reviewSheet.getLastRow() < 2) {
    ui.alert("Notice", "No data to update.", ui.ButtonSet.OK);
    return;
  }

  const reviewData = reviewSheet.getRange(2, 1, reviewSheet.getLastRow() - 1, 3).getValues(); // A to C (Color, Size, Current Count)

  ss.toast("Connecting to Master Inventory...", "Processing");
  let remoteSs, masterSheet;
  try {
    remoteSs = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    masterSheet = remoteSs.getSheetByName("master_inventory");
  } catch(e) {
    ui.alert("Error", "Could not connect to remote master spreadsheet.", ui.ButtonSet.OK);
    return;
  }

  const masterData = masterSheet.getRange(2, 1, Math.max(1, masterSheet.getLastRow() - 1), 8).getValues();
  const masterMap = new Map(); // lowerColor -> rowIdx

  masterData.forEach((row, idx) => {
    let mColor = String(row[0]).trim();
    if (mColor !== "") {
      masterMap.set(mColor.toLowerCase(), idx);
    }
  });

  const updatedMasterData = masterSheet.getRange(2, 2, Math.max(1, masterSheet.getLastRow() - 1), 7).getValues();
  const sizeColMap = { "S": 0, "M": 2, "L": 4, "XL": 6 };

  let updateCount = 0;

  reviewData.forEach(row => {
    const color = String(row[0]).trim().toLowerCase();
    const size = String(row[1]).trim();
    const count = row[2];

    if (color !== "" && size !== "ALL" && masterMap.has(color) && sizeColMap.hasOwnProperty(size)) {
      const mRowIdx = masterMap.get(color);
      const mColIdx = sizeColMap[size];
      
      updatedMasterData[mRowIdx][mColIdx] = count;
      updateCount++;
    }
  });

  // Write Master Data
  if (updateCount > 0) {
    masterSheet.getRange(2, 2, updatedMasterData.length, 7).setValues(updatedMasterData);
  }

  // Clear Review Sheet (Data & Colors)
  reviewSheet.getRange(2, 1, reviewSheet.getLastRow() - 1, 5).clearContent();
  reviewSheet.getRange(2, 3, reviewSheet.getLastRow() - 1, 1).setBackground(null); // Clear background colors
  
  SpreadsheetApp.flush();
  ui.alert("Success", `Forced ${updateCount} updates to Master Inventory.\nReview Sheet has been cleared.`, ui.ButtonSet.OK);
}


// =====================================================================
// 3. Manual Inventory Updates
// =====================================================================

function manualUpdateInventory() {
  if (!checkAuthorization()) return;

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("Confirm", "Are you sure you want to execute these manual inventory updates?", ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const updateSheet = ss.getSheetByName("UpdateInventory");

  if (!updateSheet) {
    ui.alert("Error", "Missing sheet: 'UpdateInventory'", ui.ButtonSet.OK);
    return;
  }

  const rawData = updateSheet.getRange(2, 1, Math.max(1, updateSheet.getLastRow() - 1), 7).getValues(); // A to G
  const validRows = [];
  const colorSet = new Set();

  // Validate inputs
  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const color = String(row[0]).trim();
    const action = String(row[2]).trim();

    if (color !== "") {
      if (action !== "Select" && action !== "") {
        // Check for duplicates
        const lowerColor = color.toLowerCase();
        if (colorSet.has(lowerColor)) {
          ui.alert("Error", `Duplicate color '${color}' found. Please consolidate into one row.`, ui.ButtonSet.OK);
          return;
        }
        colorSet.add(lowerColor);

        // Validate Size columns
        const sizes = [row[3], row[4], row[5], row[6]];
        const sizeNames = ["S", "M", "L", "XL"];
        for (let s = 0; s < 4; s++) {
          if (sizes[s] === "") {
            ui.alert("Error", `Please fill correct count. Size '${sizeNames[s]}' is missing for color '${color}'.\nType 0 if empty.`, ui.ButtonSet.OK);
            return;
          }
          if (typeof sizes[s] !== 'number') {
            ui.alert("Error", `Invalid number format in Size '${sizeNames[s]}' for color '${color}'. Please enter digits only.`, ui.ButtonSet.OK);
            return;
          }
        }

        validRows.push({
          sheetRowIdx: i, // 0-based index for the sheet data array
          color: color,
          action: action,
          S: sizes[0],
          M: sizes[1],
          L: sizes[2],
          XL: sizes[3]
        });
      } else if (action === "Select") {
         // Verify no numbers were entered when action is Select
         const sizes = [row[3], row[4], row[5], row[6]];
         for (let s = 0; s < 4; s++) {
           if (sizes[s] !== "" && sizes[s] !== 0) {
             ui.alert("Error", `You entered a number for color '${color}' but Action is set to 'Select'.\nPlease choose a valid action like 'Add', 'Minus', or 'Replace'.`, ui.ButtonSet.OK);
             return;
           }
         }
      }
    }
  }

  if (validRows.length === 0) {
    ui.alert("Notice", "No valid actions found to process.", ui.ButtonSet.OK);
    return;
  }

  ss.toast("Connecting to Master Inventory...", "Processing");
  let remoteSs, masterSheet;
  try {
    remoteSs = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    masterSheet = remoteSs.getSheetByName("master_inventory");
  } catch(e) {
    ui.alert("Error", "Could not connect to remote master spreadsheet.", ui.ButtonSet.OK);
    return;
  }

  const masterData = masterSheet.getRange(2, 1, Math.max(1, masterSheet.getLastRow() - 1), 8).getValues();
  const masterMap = new Map();

  masterData.forEach((row, idx) => {
    let mColor = String(row[0]).trim();
    if (mColor !== "") {
      masterMap.set(mColor.toLowerCase(), {
        rowIdx: idx,
        S: Number(row[1]) || 0,
        M: Number(row[3]) || 0,
        L: Number(row[5]) || 0,
        XL: Number(row[7]) || 0
      });
    }
  });

  const updatedMasterData = masterSheet.getRange(2, 2, Math.max(1, masterSheet.getLastRow() - 1), 7).getValues();
  const updateDataToReset = updateSheet.getRange(2, 3, rawData.length, 5).getValues(); // C to G
  let processCount = 0;

  validRows.forEach(item => {
    const lowerColor = item.color.toLowerCase();
    
    if (!masterMap.has(lowerColor)) {
      ui.alert("Warning", `Color '${item.color}' not found in master_inventory. Skipping this specific row.`, ui.ButtonSet.OK);
      return;
    }

    const mRowIdx = masterMap.get(lowerColor).rowIdx;
    const mData = masterMap.get(lowerColor);

    const sizes = ["S", "M", "L", "XL"];
    const mCols = [0, 2, 4, 6]; // Corresponds to B, D, F, H in updatedMasterData

    sizes.forEach((s, idx) => {
      let inputVal = item[s];
      let masterVal = mData[s];
      let finalVal = masterVal;

      if (item.action === "Add") {
        finalVal = masterVal + inputVal;
      } else if (item.action === "Minus") {
        finalVal = masterVal - inputVal;
      } else if (item.action === "Replace") {
        finalVal = inputVal;
      }

      updatedMasterData[mRowIdx][mCols[idx]] = finalVal;
    });

    processCount++;
    
    // Reset the row in the local reset array (Columns C to G)
    updateDataToReset[item.sheetRowIdx][0] = "Select"; // Action (C)
    updateDataToReset[item.sheetRowIdx][1] = 0;        // S (D)
    updateDataToReset[item.sheetRowIdx][2] = 0;        // M (E)
    updateDataToReset[item.sheetRowIdx][3] = 0;        // L (F)
    updateDataToReset[item.sheetRowIdx][4] = 0;        // XL (G)
  });

  // Write Master Data
  if (processCount > 0) {
    masterSheet.getRange(2, 2, updatedMasterData.length, 7).setValues(updatedMasterData);
    
    // Write local resets
    updateSheet.getRange(2, 3, updateDataToReset.length, 5).setValues(updateDataToReset);
  }

  SpreadsheetApp.flush();
  ui.alert("Success", `Successfully processed ${processCount} manual inventory updates.`, ui.ButtonSet.OK);
}
