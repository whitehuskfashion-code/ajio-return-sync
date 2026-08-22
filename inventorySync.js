/**
 * @OnlyCurrentDoc
 * This script synchronizes Myntra inventory based on physical stock, out-of-stock thresholds, and allocation rules.
 */

// Define Spreadsheet IDs
const AJIO_MASTER_SPREADSHEET_ID = "1GMfVnYMuXxB-3_m1kuPpBcnuKK6zO02mcg0hmPprAn4";

// Adds a custom menu to the spreadsheet UI
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Myntra Tools')
    .addItem('Sync Myntra Inventory', 'syncMyntraInventory')
    .addToUi();
}

/**
 * Main synchronization logic
 */
function syncMyntraInventory() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const myntraBrandTagSheet = ss.getSheetByName("myntraBrandTagData");
  const myntraInventorySheet = ss.getSheetByName("myntraInventory");
  const unknownErrorsSheet = ss.getSheetByName("unkownErrors");
  const ajioInventorySheet = ss.getSheetByName("ajioInventory");
  const myntraButNotInAjioSheet = ss.getSheetByName("myntraButNotInAjio");
  const ajioButNotInMyntraSheet = ss.getSheetByName("ajioButNotInMyntra");

  if (!myntraBrandTagSheet || !myntraInventorySheet || !unknownErrorsSheet || !ajioInventorySheet || !myntraButNotInAjioSheet || !ajioButNotInMyntraSheet) {
    ui.alert("Error", "Missing one of the required sheets: 'myntraBrandTagData', 'myntraInventory', 'unkownErrors', 'ajioInventory', 'myntraButNotInAjio', or 'ajioButNotInMyntra'. Please verify sheet names.", ui.ButtonSet.OK);
    return;
  }

  // Notify user that processing has started
  ss.toast("Connecting to Master Inventory... This may take a moment.", "Processing");

  try {
    // --- 1. Fetch Remote Data from Ajio Master Spreadsheet ---
    const remoteSs = SpreadsheetApp.openById(AJIO_MASTER_SPREADSHEET_ID);
    const mappingSheet = remoteSs.getSheetByName("Mapping Sheet");
    const masterInventorySheet = remoteSs.getSheetByName("master_inventory");
    const inventoryLookupSheet = remoteSs.getSheetByName("inventory_lookup");

    if (!mappingSheet || !masterInventorySheet || !inventoryLookupSheet) {
      throw new Error("Could not find required sheets in the remote 'Print Number and Qty' spreadsheet.");
    }

    const mappingData = mappingSheet.getRange(2, 1, Math.max(1, mappingSheet.getLastRow() - 1), 52).getValues();
    const masterData = masterInventorySheet.getRange(2, 1, Math.max(1, masterInventorySheet.getLastRow() - 1), 8).getValues();
    const lookupData = inventoryLookupSheet.getRange(2, 1, Math.max(1, inventoryLookupSheet.getLastRow() - 1), 18).getValues();

    // --- 2. Build Memory Maps ---
    const mappingMap = new Map(); // sku -> [ {masterProduct, size, error} ]
    mappingData.forEach(row => {
      let masterProduct = String(row[2] || "").trim(); // Column C

      for (let i = 0; i < row.length; i++) {
        let rawSku = String(row[i]);
        if (rawSku.trim() !== "") {
          let sku = rawSku.trim().toLowerCase();
          
          if (!mappingMap.has(sku)) {
            mappingMap.set(sku, []);
          }

          if (i < 3) {
            // Found in Col A, B, or C
            mappingMap.get(sku).push({ error: "SKU found in Mapping Sheet but not in a valid size column" });
          } else {
            // Found in size columns D onwards
            let sizeIndex = (i - 3) % 4;
            let sizeStr = ["S", "M", "L", "XL"][sizeIndex];
            mappingMap.get(sku).push({ masterProduct: masterProduct, size: sizeStr });
          }
        }
      }
    });

    const masterInventoryMap = new Map(); // product -> {S, M, L, XL}
    masterData.forEach(row => {
      let product = String(row[0]).trim();
      if (product) {
        masterInventoryMap.set(product, {
          S: Number(row[1]) || 0,
          M: Number(row[3]) || 0,
          L: Number(row[5]) || 0,
          XL: Number(row[7]) || 0
        });
      }
    });

    const lookupMap = new Map(); // product -> { OOS_TH: {S, M, L, XL}, ALLOC }
    lookupData.forEach(row => {
      let product = String(row[0]).trim();
      if (product) {
        lookupMap.set(product, {
          OOS_TH: {
            S: row[5] !== "" ? Number(row[5]) : null,  // Column F
            M: row[8] !== "" ? Number(row[8]) : null,  // Column I
            L: row[11] !== "" ? Number(row[11]) : null, // Column L
            XL: row[14] !== "" ? Number(row[14]) : null // Column O
          },
          ALLOC: row[17] !== "" ? Number(row[17]) : null // Column R
        });
      }
    });

    // --- 3. Read Myntra & Ajio Input Data ---
    const myntraRawData = myntraBrandTagSheet.getRange(2, 2, Math.max(1, myntraBrandTagSheet.getLastRow() - 1), 7).getValues();
    const processedSanitized = new Set();
    const finalSkusToProcess = [];

    myntraRawData.forEach(row => {
      let rawSku = String(row[0]);
      let colour = String(row[6]);

      // Skip "OutOfStock" labeled SKUs entirely (case-insensitive)
      if (colour.toLowerCase().includes("outofstock")) {
        return;
      }

      if (rawSku.trim() !== "") {
        let sanitized = rawSku.trim().toLowerCase();
        if (!processedSanitized.has(sanitized)) {
          processedSanitized.add(sanitized);
          finalSkusToProcess.push({ original: rawSku, sanitized: sanitized });
        }
      }
    });

    const ajioRawData = ajioInventorySheet.getRange(2, 1, Math.max(1, ajioInventorySheet.getLastRow() - 1), 2).getValues();
    const ajioSkuMap = new Map();
    const ajioQtyArray = [];
    
    ajioRawData.forEach((row, index) => {
      let ajioSku = String(row[0]).trim();
      let ajioQty = row[1];
      ajioQtyArray.push([ajioQty]); // Preserve existing values initially
      if (ajioSku !== "") {
        ajioSkuMap.set(ajioSku.toLowerCase(), index);
      }
    });

    const touchedAjioIndices = new Set();

    // --- 4. Evaluate Logic ---
    const successArray = [];
    const errorArray = [];
    const missingInAjioArray = [];

    finalSkusToProcess.forEach(item => {
      const rawSku = item.original;
      const sanitized = item.sanitized;

      // 4a. Check Mapping Sheet
      if (!mappingMap.has(sanitized)) {
        errorArray.push([rawSku, "SKU not found in Mapping Sheet"]);
        return;
      }

      const mappingEntries = mappingMap.get(sanitized);

      // Check for column boundary error
      const colError = mappingEntries.find(e => e.error);
      if (colError) {
        errorArray.push([rawSku, colError.error]);
        return;
      }

      // Check for conflict errors (multiple products or sizes)
      let firstProduct = mappingEntries[0].masterProduct;
      let firstSize = mappingEntries[0].size;
      let conflict = false;

      for (let i = 1; i < mappingEntries.length; i++) {
        if (mappingEntries[i].masterProduct !== firstProduct) {
          errorArray.push([rawSku, "SKU maps to multiple master products in Mapping Sheet"]);
          conflict = true;
          break;
        }
        if (mappingEntries[i].size !== firstSize) {
          errorArray.push([rawSku, "SKU maps to conflicting sizes in Mapping Sheet"]);
          conflict = true;
          break;
        }
      }
      if (conflict) return;

      if (!firstProduct) {
        errorArray.push([rawSku, "Master product column (C) is empty in Mapping Sheet"]);
        return;
      }

      // 4b. Fetch Physical Stock
      if (!masterInventoryMap.has(firstProduct)) {
        errorArray.push([rawSku, `Master product '${firstProduct}' not found in master_inventory`]);
        return;
      }
      const physicalStock = masterInventoryMap.get(firstProduct)[firstSize];

      // 4c. Fetch Thresholds & Math
      if (!lookupMap.has(firstProduct)) {
        errorArray.push([rawSku, `Master product '${firstProduct}' not found in inventory_lookup`]);
        return;
      }

      const lData = lookupMap.get(firstProduct);
      const oosTh = lData.OOS_TH[firstSize];
      const alloc = lData.ALLOC;

      if (oosTh === null || isNaN(oosTh) || String(oosTh).trim() === "") {
        errorArray.push([rawSku, "Missing OOS_TH in inventory_lookup"]);
        return;
      }

      if (alloc === null || isNaN(alloc) || String(alloc).trim() === "") {
        errorArray.push([rawSku, "Missing ALLOC percentage in inventory_lookup"]);
        return;
      }

      let finalQty = 0;
      if (physicalStock <= oosTh) {
        finalQty = 0; // Out of stock threshold tripped
      } else {
        finalQty = Math.floor(physicalStock * alloc);
      }

      successArray.push([rawSku, finalQty]);

      // 4d. Map back to Ajio Inventory
      if (ajioSkuMap.has(sanitized)) {
        let index = ajioSkuMap.get(sanitized);
        ajioQtyArray[index] = [finalQty]; // Update the quantity
        touchedAjioIndices.add(index);
      } else {
        missingInAjioArray.push(rawSku);
      }
    });

    // Determine untouched Ajio SKUs
    const untouchedAjioSkus = [];
    ajioRawData.forEach((row, index) => {
      let ajioSku = String(row[0]).trim();
      if (ajioSku !== "" && !touchedAjioIndices.has(index)) {
        untouchedAjioSkus.push(ajioSku);
      }
    });

    // --- 5. Write Data to Sheets ---
    ss.toast("Writing bulk results...", "Processing");

    // Clear existing data (keeping row 1 headers)
    if (myntraInventorySheet.getLastRow() > 1) {
      myntraInventorySheet.getRange(2, 1, myntraInventorySheet.getLastRow() - 1, 2).clearContent();
    }
    if (unknownErrorsSheet.getLastRow() > 1) {
      unknownErrorsSheet.getRange(2, 1, unknownErrorsSheet.getLastRow() - 1, 2).clearContent();
    }
    if (myntraButNotInAjioSheet.getLastRow() > 1) {
      myntraButNotInAjioSheet.getRange(2, 1, myntraButNotInAjioSheet.getLastRow() - 1, 2).clearContent();
    }
    if (ajioButNotInMyntraSheet.getLastRow() > 1) {
      ajioButNotInMyntraSheet.getRange(2, 1, ajioButNotInMyntraSheet.getLastRow() - 1, 2).clearContent();
    }

    // Write Myntra data
    if (successArray.length > 0) {
      myntraInventorySheet.getRange(2, 1, successArray.length, 2).setValues(successArray);
    }
    if (errorArray.length > 0) {
      unknownErrorsSheet.getRange(2, 1, errorArray.length, 2).setValues(errorArray);
    }

    // Bulk write updated Ajio column (if there's any data to write)
    if (ajioQtyArray.length > 0) {
      ajioInventorySheet.getRange(2, 2, ajioQtyArray.length, 1).setValues(ajioQtyArray);
    }

    // Write Missing in Ajio data with timestamps
    if (missingInAjioArray.length > 0) {
      const timestamp = new Date().toLocaleString();
      const missingWriteData = missingInAjioArray.map(sku => [sku, timestamp]);
      myntraButNotInAjioSheet.getRange(2, 1, missingWriteData.length, 2).setValues(missingWriteData);
    }

    // Write Untouched Ajio data with reason
    if (untouchedAjioSkus.length > 0) {
      const reason = "SKU present in ajioInventory but missing or skipped in myntraBrandTagData today.";
      const untouchedWriteData = untouchedAjioSkus.map(sku => [sku, reason]);
      ajioButNotInMyntraSheet.getRange(2, 1, untouchedWriteData.length, 2).setValues(untouchedWriteData);
    }

    // Final completion flush and toast
    SpreadsheetApp.flush();

    // Construct final alert message
    let resultMessage = `✅ Sync Complete!\n\n📦 SKUs updated in Myntra & Ajio: ${successArray.length}\n⚠️ General Errors: ${errorArray.length}`;
    
    if (missingInAjioArray.length > 0) {
      resultMessage += `\n❌ SKUs in Myntra but missing from Ajio: ${missingInAjioArray.length} (Check myntraButNotInAjio sheet)`;
    }

    if (untouchedAjioSkus.length > 0) {
      resultMessage += `\n⚠️ Untouched Ajio SKUs (Not updated today): ${untouchedAjioSkus.length} (Check ajioButNotInMyntra sheet)`;
    }

    ss.toast("Done!", "Success");
    ui.alert("Sync Report", resultMessage, ui.ButtonSet.OK);

  } catch (err) {
    ui.alert("Fatal Error", err.toString(), ui.ButtonSet.OK);
  }
}
