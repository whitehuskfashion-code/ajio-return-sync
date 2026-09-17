// function updateMyntraFromMaster() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var myntraSheet = ss.getSheetByName("Myntra");
//   var masterSheet = ss.getSheetByName("Master");

//   var masterLastRow = masterSheet.getLastRow();
//   if (masterLastRow < 2) return; 

//   // Get Master data: fetch 4 columns starting from Col B (B, C, D, E)
//   // Index 0 = Col B (Style code)
//   // Index 1 = Col C 
//   // Index 2 = Col D 
//   // Index 3 = Col E 
//   var masterData = masterSheet.getRange(2, 2, masterLastRow - 1, 4).getValues(); 

//   // Build a lookup map holding an object for each column we need
//   var lookupMap = {};
//   masterData.forEach(function(row) {
//     var styleCode = row[0].toString().trim();
//     if (styleCode !== "") {
//       lookupMap[styleCode] = {
//         valC: row[1], // Master Col C
//         valD: row[2], // Master Col D
//         valE: row[3]  // Master Col E
//       };
//     }
//   });

//   var myntraLastRow = myntraSheet.getLastRow();
//   if (myntraLastRow < 2) return; 

//   // Get Myntra data: Col C (Style Code)
//   var myntraCData = myntraSheet.getRange(2, 3, myntraLastRow - 1, 1).getValues(); 

//   // Create separate arrays for each target column
//   var outputD = [];
//   var outputE = [];
//   var outputF = [];

//   myntraCData.forEach(function(row) {
//     var cValue = row[0].toString().trim();
//     if (lookupMap.hasOwnProperty(cValue)) {
//       outputD.push([lookupMap[cValue].valC]); // Maps Master C to Myntra D
//       outputE.push([lookupMap[cValue].valD]); // Maps Master D to Myntra E
//       outputF.push([lookupMap[cValue].valE]); // Maps Master E to Myntra F
//     } else {
//       outputD.push([""]); 
//       outputE.push([""]); 
//       outputF.push([""]); 
//     }
//   });

//   // Write results to Myntra strictly in their individual columns
//   myntraSheet.getRange(2, 4, outputD.length, 1).setValues(outputD); // Pastes to D
//   myntraSheet.getRange(2, 5, outputE.length, 1).setValues(outputE); // Pastes to E
//   myntraSheet.getRange(2, 6, outputF.length, 1).setValues(outputF); // Pastes to F
// }



function matchAndCopyMyntraData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTP = ss.getSheetByName("forTP");
  var sheetMyntra = ss.getSheetByName("Myntra");

  // Safety check to ensure both sheets exist
  if (!sheetTP || !sheetMyntra) {
    SpreadsheetApp.getUi().alert("Error: Please make sure you have sheets exactly named 'forTP' and 'Myntra'.");
    return;
  }

  // 1. Get all data from the Myntra sheet
  var myntraData = sheetMyntra.getDataRange().getValues();
  var myntraMap = {};

  // Create a dictionary of Myntra Style IDs (Col A) and their corresponding Col B, C, and E values
  for (var i = 0; i < myntraData.length; i++) {
    var styleId = myntraData[i][0]; // Column A
    var valueB = myntraData[i][1];  // Column B
    var valueC = myntraData[i][2];  // Column C
    var valueE = myntraData[i][4];  // Column E (Index is 4 since arrays are 0-indexed)

    if (styleId !== "" && styleId != null) {
      var cleanStyleId = String(styleId).trim();
      // Store B, C, and E values together
      myntraMap[cleanStyleId] = { colB: valueB, colC: valueC, colE: valueE };
    }
  }

  // 2. Get data from the 'forTP' sheet
  var lastRowTP = sheetTP.getLastRow();
  if (lastRowTP === 0) return; // Stop if the 'forTP' sheet is completely empty

  // Grab Columns A, B, C, and D from 'forTP' 
  // (We grab B, C, and D so we don't overwrite rows that don't find a match)
  var tpRange = sheetTP.getRange(1, 1, lastRowTP, 4);
  var tpValues = tpRange.getValues();

  var outputColsBCD = []; // This will hold the new values for Columns B, C, and D

  // 3. Match and prepare the new data array
  for (var j = 0; j < tpValues.length; j++) {
    var tpStyleId = tpValues[j][0];
    var existingValueB = tpValues[j][1];
    var existingValueC = tpValues[j][2];
    var existingValueD = tpValues[j][3]; // Existing Column D

    if (tpStyleId !== "" && tpStyleId != null) {
      var cleanTpStyleId = String(tpStyleId).trim();

      // If the Style ID exists in our Myntra map, grab all 3 values
      if (myntraMap.hasOwnProperty(cleanTpStyleId)) {
        outputColsBCD.push([
          myntraMap[cleanTpStyleId].colB,
          myntraMap[cleanTpStyleId].colC,
          myntraMap[cleanTpStyleId].colE // Places Myntra's Col E into forTP's Col D
        ]);
      } else {
        // If there is no match, leave existing Columns B, C, and D exactly as they were
        outputColsBCD.push([existingValueB, existingValueC, existingValueD]);
      }
    } else {
      // If Column A is blank, leave existing Columns B, C, and D exactly as they were
      outputColsBCD.push([existingValueB, existingValueC, existingValueD]);
    }
  }

  // 4. Paste the updated values back into Columns B, C, and D of 'forTP'
  // (Start at row 1, column 2, write down 'lastRowTP' rows, across 3 columns)
  sheetTP.getRange(1, 2, lastRowTP, 3).setValues(outputColsBCD);

  SpreadsheetApp.getUi().toast("Columns B, C, and D sync from Myntra sheet is complete!", "Success");
}

// function syncMyntraMRP() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const mrpSheet = ss.getSheetByName("MRP");
//   const myntraSheet = ss.getSheetByName("Myntra");

//   if (!mrpSheet || !myntraSheet) {
//     SpreadsheetApp.getUi().alert("Error: Please ensure both 'MRP' and 'Myntra' sheets exist.");
//     return;
//   }

//   // 1. Get all data from the MRP sheet
//   const mrpData = mrpSheet.getDataRange().getValues();
//   const mrpMap = {};

//   // 2. Build a lookup map from the MRP sheet
//   // Assuming row 1 is headers, so we start at i = 1
//   for (let i = 1; i < mrpData.length; i++) {
//     let styleId = mrpData[i][0];     // Column A: The Style ID to match
//     let valueToCopy = mrpData[i][1]; // Column B: The MRP value to copy

//     if (styleId) {
//       // Convert to string and trim to avoid spacing mismatches
//       mrpMap[styleId.toString().trim()] = valueToCopy;
//     }
//   }

//   // 3. Get all data from the Myntra sheet
//   const myntraData = myntraSheet.getDataRange().getValues();
//   const newColG = [];

//   // 4. Match and prepare the new Column G
//   for (let i = 0; i < myntraData.length; i++) {
//     if (i === 0) {
//       // Keep the existing header for Column G (Index 6)
//       newColG.push([myntraData[i][6] || "Updated Value"]); 
//       continue;
//     }

//     let myntraStyleId = myntraData[i][0]; // Column A (Index 0)
//     let existingColGValue = myntraData[i][6]; // Column G (Index 6)

//     if (myntraStyleId) {
//       let cleanMyntraId = myntraStyleId.toString().trim();

//       if (mrpMap.hasOwnProperty(cleanMyntraId)) {
//         // Match found! Push the copied MRP value from Column B
//         newColG.push([mrpMap[cleanMyntraId]]);
//       } else {
//         // No match found. Keep whatever was already in Column G.
//         newColG.push([existingColGValue]); 
//       }
//     } else {
//        // Blank row, keep existing
//        newColG.push([existingColGValue]);
//     }
//   }

//   // 5. Write the entire array back to Column G (Column number 7)
//   myntraSheet.getRange(1, 7, newColG.length, 1).setValues(newColG);

//   SpreadsheetApp.getUi().alert("Sync complete!");
// }

/**
 * update_tp_prices.js
 * Synchronizes TP Prices from myntraBrandTagMastersheet to Myntra_Ajio_Selling_settlement.
 */

const TAG_MASTER_SS_ID = "12-iec8Uv9eMJdnXJS8N64OD8e_bccjhO4V5NeylVVgM";
const TAG_DATA_SHEET = "myntraBrandTagData";
const MASTER_SHEET = "Master";
const MYNTRA_SHEET = "Myntra";

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Sync Myntra Data')
    .addItem('Sync Now', 'syncMyntraBrandTags')
    .addItem('Match and Copy Myntra Data', 'matchAndCopyMyntraData')
    .addToUi();
}

function syncMyntraBrandTags() {
  const targetSs = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Fetch Source Data
  let tagSs;
  try {
    tagSs = SpreadsheetApp.openById(TAG_MASTER_SS_ID);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Error: Cannot open the Tag Master sheet. Ensure you have access.");
    return;
  }

  const tagWs = tagSs.getSheetByName(TAG_DATA_SHEET);
  if (!tagWs) {
    SpreadsheetApp.getUi().alert(`Error: Sheet '${TAG_DATA_SHEET}' not found in Tag Master.`);
    return;
  }
  const tagData = tagWs.getDataRange().getValues();

  // 2. Fetch Master Data (for financials)
  const masterWs = targetSs.getSheetByName(MASTER_SHEET);
  if (!masterWs) {
    SpreadsheetApp.getUi().alert(`Error: Sheet '${MASTER_SHEET}' not found in current spreadsheet.`);
    return;
  }
  const masterData = masterWs.getDataRange().getValues();

  // 3. Fetch Myntra Sheet Data (Target)
  const myntraWs = targetSs.getSheetByName(MYNTRA_SHEET);
  if (!myntraWs) {
    SpreadsheetApp.getUi().alert(`Error: Sheet '${MYNTRA_SHEET}' not found in current spreadsheet.`);
    return;
  }
  const myntraData = myntraWs.getDataRange().getValues();

  // ==========================================
  // HASH MAP CREATION
  // ==========================================

  // Master Map: price_level -> [Settlement, Selling Value, Event Settlement, Event Selling]
  const masterMap = new Map();
  for (let i = 1; i < masterData.length; i++) {
    const row = masterData[i];
    const priceLevel = String(row[0] || "").trim().toLowerCase();
    if (priceLevel) {
      masterMap.set(priceLevel, {
        c: row[1] || "", // Settlement Amount
        d: row[2] || "", // Selling Value on Myntra
        e: row[3] || "", // Event Settlement
        f: row[4] || ""  // Event Selling on Myntra
      });
    }
  }

  // Target Map: Myntra Style Id -> rowIndex
  const targetMap = new Map();
  for (let i = 1; i < myntraData.length; i++) {
    const styleId = String(myntraData[i][0] || "").trim().toLowerCase();
    if (styleId) {
      targetMap.set(styleId, i);
    }
  }

  // ==========================================
  // AGGREGATE SOURCE DATA
  // ==========================================

  const groupedData = new Map();

  // tagData headers: A:0 (style), B:1 (sku), C:2 (mrp), K:10 (price_level)
  for (let i = 1; i < tagData.length; i++) {
    const row = tagData[i];
    const styleId = String(row[0] || "").trim();
    if (!styleId) continue;

    const styleIdLower = styleId.toLowerCase();
    const skuMstr = String(row[1] || "").trim();
    const mrp = String(row[2] || "").trim();
    const priceLevel = String(row[10] || "").trim();

    // Extract Style Code robustly (non-case-sensitive trim)
    const styleCode = skuMstr.replace(/[-\s](S|M|L|XL|XXL|XXXL|FS)$/i, "").trim();

    if (!groupedData.has(styleIdLower)) {
      groupedData.set(styleIdLower, {
        originalStyleId: styleId,
        styleCode: styleCode,
        mrp: mrp,
        priceLevel: priceLevel,
        hasConflict: false
      });
    } else {
      const existing = groupedData.get(styleIdLower);
      if (existing.mrp !== mrp || existing.priceLevel !== priceLevel) {
        existing.hasConflict = true;
      }
    }
  }

  // ==========================================
  // PREPARE BULK UPDATES
  // ==========================================

  const lastRow = Math.max(1, myntraWs.getLastRow());
  const lastCol = Math.max(9, myntraWs.getLastColumn()); // Guarantee at least up to column I (9)

  // Fetch existing backgrounds safely
  let myntraBackgrounds = [];
  if (myntraWs.getLastRow() > 0) {
    myntraBackgrounds = myntraWs.getRange(1, 1, lastRow, lastCol).getBackgrounds();
  } else {
    myntraBackgrounds = [Array(lastCol).fill(null)];
  }

  const rowsToAppend = [];
  let updateCount = 0;

  for (const [styleIdLower, group] of groupedData.entries()) {

    const searchPriceLevel = group.priceLevel.toLowerCase();
    const masterInfo = masterMap.get(searchPriceLevel);

    // Determine Highlight colors
    let bgMrp = null;
    let bgSettlement = null;

    if (group.hasConflict || !group.mrp) bgMrp = "#ffff00";
    if (group.hasConflict || !group.priceLevel) bgSettlement = "#ffff00";
    if (group.priceLevel && !masterInfo) bgSettlement = "#ff0000"; // Red if not in Master

    const valC = masterInfo ? masterInfo.c : "";
    const valD = masterInfo ? masterInfo.d : "";
    const valE = masterInfo ? masterInfo.e : "";
    const valF = masterInfo ? masterInfo.f : "";

    if (targetMap.has(styleIdLower)) {
      const rIdx = targetMap.get(styleIdLower);

      // Update Values
      myntraData[rIdx][1] = group.styleCode;
      myntraData[rIdx][2] = valC;
      myntraData[rIdx][3] = valD;
      myntraData[rIdx][4] = valE;
      myntraData[rIdx][5] = valF;
      myntraData[rIdx][6] = group.mrp;

      // Safely pad the background row array if it's shorter than lastCol
      while (myntraBackgrounds[rIdx].length < lastCol) {
        myntraBackgrounds[rIdx].push(null);
      }

      // Update Backgrounds
      myntraBackgrounds[rIdx][1] = null;
      myntraBackgrounds[rIdx][2] = bgSettlement;
      myntraBackgrounds[rIdx][3] = bgSettlement;
      myntraBackgrounds[rIdx][4] = bgSettlement;
      myntraBackgrounds[rIdx][5] = bgSettlement;
      myntraBackgrounds[rIdx][6] = bgMrp;

      updateCount++;
    } else {
      const physicalRow = lastRow + rowsToAppend.length + 1;
      const newRow = [
        group.originalStyleId,
        group.styleCode,
        valC,
        valD,
        valE,
        valF,
        group.mrp,
        `=IFERROR((G${physicalRow}-F${physicalRow})/G${physicalRow}*100, "")`,
        `=ROUND(H${physicalRow},0)`
      ];
      // Pad to match original data array length if it's larger than 9
      const targetLen = Math.max(9, myntraData[0].length);
      while (newRow.length < targetLen) newRow.push("");
      rowsToAppend.push(newRow);

      // Background for new row
      const newBgRow = Array(lastCol).fill(null);
      newBgRow[2] = bgSettlement;
      newBgRow[3] = bgSettlement;
      newBgRow[4] = bgSettlement;
      newBgRow[5] = bgSettlement;
      newBgRow[6] = bgMrp;
      myntraBackgrounds.push(newBgRow);
    }
  }

  // ==========================================
  // BULK WRITE TO SPREADSHEET
  // ==========================================

  // 1. Write Updates
  if (myntraData.length > 1) {
    myntraWs.getRange(1, 1, myntraData.length, myntraData[0].length).setValues(myntraData);
  }

  // 2. Append New Rows
  if (rowsToAppend.length > 0) {
    myntraWs.getRange(lastRow + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }

  // 3. Write Backgrounds
  myntraWs.getRange(1, 1, myntraBackgrounds.length, myntraBackgrounds[0].length).setBackgrounds(myntraBackgrounds);

  targetSs.toast(`Sync complete! Updated ${updateCount} rows and appended ${rowsToAppend.length} rows.`, "Success", 8);
}
