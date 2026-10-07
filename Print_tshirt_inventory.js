/**
 * @OnlyCurrentDoc
 * This script automates stock deduction based on a list of SKUs.
 *
 * It reads SKUs from 'test sheet' (Column I), finds the corresponding product
 * in 'Mapping Sheet', locates that product in 'test sheet' (Column B),
 * and decrements the stock count in 'test sheet' (Column D).
 */

// Adds a custom menu to the spreadsheet UI for easy access.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Stock Tools')
    .addItem('Process SKUs & Update Stock', 'processSKUsAndDecrementStock')
    .addItem('Generate Print Order', 'generatePrintOrder')  // 👈 NEW OPTION
    .addSeparator()
    .addItem('Clean Zero RTO Stock', 'cleanZeroRtoStock')
    .addItem('Update Lookups & Thresholds', 'updateInventoryLookupsAndThresholds')
    .addItem('Suggest Fabric Rolls', 'generateFabricRollSuggestions')
    .addSeparator()
    .addItem('Update Weight', 'updateDynamicWeights')
    .addItem('Update Dynamic Thresholds', 'updateDynamicThresholds')
    .addToUi();
}

/**
 * Main function to process SKUs and update stock levels.
 */
function _normalizeSku(sku) {
  if (!sku) return "";
  return String(sku).replace(/[^a-z0-9]/gi, "").toLowerCase();
}

function processSKUsAndDecrementStock() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  const LAST_RUN_KEY = 'LAST_SUCCESSFUL_RUN_TIMESTAMP';
  const FIVE_MINUTES_IN_MS = 5 * 60 * 1000;

  try {
    // --- Safety Check Logic ---
    const lastRunTimestamp = scriptProperties.getProperty(LAST_RUN_KEY);
    if (lastRunTimestamp) {
      const timeSinceLastRun = new Date().getTime() - parseInt(lastRunTimestamp, 10);

      if (timeSinceLastRun < FIVE_MINUTES_IN_MS) {
        const promptMessage = "Heads up! This was run in the last 5 mins ⏱️\n" +
          "Running it again might double-deduct stock.\n\n" +
          "You sure you wanna do this?\n\n" +
          "▶️ Select 'Yes' to proceed (\"Yes, deduct ✅\")\n" +
          "▶️ Select 'No' to cancel (\"No, You saved me 😬\")";

        const response = ui.alert("Are you sure?", promptMessage, ui.ButtonSet.YES_NO);
        if (response !== ui.Button.YES) {
          ui.alert("Action cancelled.", "No changes have been made.", ui.ButtonSet.OK);
          return;
        }
      }
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mappingSheet = ss.getSheetByName("Mapping Sheet");
    const testSheet = ss.getSheetByName("test sheet");
    const masterInventorySheet = ss.getSheetByName("master_inventory");
    const rtoInventorySheet = ss.getSheetByName("rto_inventory");
    const bundleSheet = ss.getSheetByName("Bundle_SKU_Mapping");

    // --- 1. Validate that the required sheets exist ---
    if (!mappingSheet || !testSheet || !masterInventorySheet || !rtoInventorySheet) {
      ui.alert("Error", "Could not find 'Mapping Sheet', 'test sheet', 'master_inventory', or 'rto_inventory'. Please make sure the sheet names are correct.", ui.ButtonSet.OK);
      return;
    }

    // --- 2. Read all data into memory for performance ---
    const mappingData = mappingSheet.getRange("A2:AZ" + mappingSheet.getLastRow()).getValues();
    const testSheetLastRow = testSheet.getLastRow();

    const testDataRange = testSheet.getRange("B2:J" + testSheetLastRow);
    const testData = testDataRange.getValues();

    const skusToProcess = testSheet.getRange("H2:I" + testSheetLastRow).getValues();

    const masterInventoryLastRow = masterInventorySheet.getLastRow();
    const masterInventoryData = masterInventorySheet.getRange("A2:H" + masterInventoryLastRow).getValues();

    const rtoInventoryLastRow = rtoInventorySheet.getLastRow();
    const rtoInventoryData = rtoInventoryLastRow >= 2 ? rtoInventorySheet.getRange("A2:C" + rtoInventoryLastRow).getValues() : [];

    const bundleData = bundleSheet ? bundleSheet.getDataRange().getValues() : [];
    const bundleMap = new Map();
    if (bundleData.length > 0) {
      for (let i = 1; i < bundleData.length; i++) {
        const row = bundleData[i];
        const rawBundleSku = String(row[0] || "").trim().toUpperCase();
        if (rawBundleSku) {
          const bundleNorm = _normalizeSku(rawBundleSku);
          const children = [];
          for (let c = 1; c < row.length; c++) {
            const child = String(row[c] || "").trim().toUpperCase();
            if (child) children.push(child);
          }
          if (children.length > 0) bundleMap.set(bundleNorm, children);
        }
      }
    }

    // --- 3. Build comprehensive SKU lookup maps ---
    const rtoInventoryInfoMap = new Map();
    rtoInventoryData.forEach((row, index) => {
      const rawRtoSku = row[0]; // Column A
      if (rawRtoSku !== null && rawRtoSku !== undefined && rawRtoSku !== "") {
        const rtoSku = String(rawRtoSku).trim().toUpperCase();
        if (rtoSku) {
          let rtoNorm = _normalizeSku(rtoSku);
          if (!rtoInventoryInfoMap.has(rtoNorm)) {
            rtoInventoryInfoMap.set(rtoNorm, {
              rowIndex: index,
              count: Number(row[1]) || 0,
              locked: Number(row[2]) || 0
            });
          }
        }
      }
    });

    const skuToMappingInfo = new Map();
    mappingData.forEach((row, rowIndex) => {
      const rawProductName = row[0];
      const rawMasterProductName = row[2];
      const productName = (rawProductName === null || rawProductName === undefined || rawProductName === "") ? "" : String(rawProductName).trim().toUpperCase();
      const masterProductName = (rawMasterProductName === null || rawMasterProductName === undefined || rawMasterProductName === "") ? "" : String(rawMasterProductName).trim().toUpperCase();

      for (let i = 3; i < row.length; i++) {
        const rawSku = row[i];
        if (rawSku !== null && rawSku !== undefined && rawSku !== "") {
          const sku = String(rawSku).trim().toUpperCase();
          if (sku) {
            const actualColumnIndex = i + 1;
            let size = null;
            if (actualColumnIndex >= 4) {
              const sizeIndex = (actualColumnIndex - 4) % 4;
              switch (sizeIndex) {
                case 0: size = 'S'; break;
                case 1: size = 'M'; break;
                case 2: size = 'L'; break;
                case 3: size = 'XL'; break;
              }
            }
            skuToMappingInfo.set(_normalizeSku(sku), {
              productName: productName || "",
              masterProductName: masterProductName || "",
              size: size,
              columnIndex: actualColumnIndex,
              rowIndex: rowIndex
            });
          }
        }
      }
    });

    const productInfoMap = new Map();
    testData.forEach((row, index) => {
      const rawProductName = row[0];
      const stockCount = row[2];
      if (rawProductName !== null && rawProductName !== undefined && rawProductName !== "") {
        const productName = String(rawProductName).trim().toUpperCase();
        if (productName) {
          if (!productInfoMap.has(productName)) {
            productInfoMap.set(productName, []);
          }
          productInfoMap.get(productName).push({ rowIndex: index, stock: stockCount });
        }
      }
    });

    const masterProductInfoMap = new Map();
    masterInventoryData.forEach((row, index) => {
      const rawMasterProductName = row[0];
      if (rawMasterProductName !== null && rawMasterProductName !== undefined && rawMasterProductName !== "") {
        const masterProductName = String(rawMasterProductName).trim().toUpperCase();
        if (masterProductName) {
          if (!masterProductInfoMap.has(masterProductName)) {
            masterProductInfoMap.set(masterProductName, []);
          }
          masterProductInfoMap.get(masterProductName).push({
            rowIndex: index,
            stocks: { S: row[1], M: row[3], L: row[5], XL: row[7] }
          });
        }
      }
    });

    // --- 4. Process each SKU and prepare the results ---
    const results = [];
    const statusColors = [];
    const rtoColors = [];
    const stockUpdates = testData.map(row => [row[2]]);
    const masterInventoryDecrements = new Map();

    skusToProcess.forEach((row, index) => {
      const rawRtoSku = row[0];
      const rawPrintSku = row[1];
      let status = "";
      let rtoStatus = "";
      let cellSeverity = 0; // 0=Black, 1=Purple, 2=Red
      let rtoSeverity = 0;

      // --- Process PRINT SKU ---
      if (rawPrintSku !== null && rawPrintSku !== undefined && rawPrintSku !== "") {
        const inputSku = String(rawPrintSku).trim().toUpperCase();
        if (inputSku) {
          let printSkusToProcess = [inputSku];
          let inputNorm = _normalizeSku(inputSku);
          if (bundleMap.has(inputNorm)) {
            printSkusToProcess = bundleMap.get(inputNorm);
          }

          let preFlightFailed = false;
          let abortMsg = "";

          // --- PRINT PRE-FLIGHT CHECK ---
          for (let sku of printSkusToProcess) {
            const mappingInfo = skuToMappingInfo.get(_normalizeSku(sku));
            if (!mappingInfo) {
              preFlightFailed = true;
              abortMsg = `${sku} not in Mapping Sheet`;
              break;
            }
            const productName = mappingInfo.productName;
            const masterProductName = mappingInfo.masterProductName;

            if (productName) {
              const productEntries = productInfoMap.get(productName);
              if (!productEntries || productEntries.length === 0) {
                preFlightFailed = true;
                abortMsg = `${productName} not in test sheet`;
                break;
              }
            }
            if (masterProductName) {
              const size = mappingInfo.size;
              if (!size) {
                preFlightFailed = true;
                abortMsg = `Size not determined for column ${mappingInfo.columnIndex}`;
                break;
              }
              const masterEntries = masterProductInfoMap.get(masterProductName);
              if (!masterEntries || masterEntries.length === 0) {
                preFlightFailed = true;
                abortMsg = `${masterProductName} not in master_inventory`;
                break;
              }
            }
          }

          // --- PRINT EXECUTION ---
          if (preFlightFailed) {
            status = `Print: ❌ (ABORTED: ${abortMsg}) | Plain/Ready Merch: ❌ (ABORTED)`;
            cellSeverity = 2;
          } else {
            let printLogs = [];
            let plainLogs = [];

            for (let sku of printSkusToProcess) {
              const mappingInfo = skuToMappingInfo.get(_normalizeSku(sku));
              const productName = mappingInfo.productName;
              const masterProductName = mappingInfo.masterProductName;

              let originalOpStatus = "";
              let masterOpStatus = "";

              if (productName) {
                const productEntries = productInfoMap.get(productName);
                let decrementedCount = 0;
                const warnings = [];
                productEntries.forEach(entry => {
                  const stock = stockUpdates[entry.rowIndex][0];
                  if (typeof stock === 'number') {
                    stockUpdates[entry.rowIndex][0] = stock - 1;
                    decrementedCount++;
                  } else {
                    warnings.push(`typeof stock != 'number'`);
                  }
                });

                if (decrementedCount > 0) {
                  originalOpStatus = `✅ (${sku})`;
                  if (warnings.length) {
                    originalOpStatus += ` Warnings: ${warnings.join('; ')}`;
                  }
                } else {
                  originalOpStatus = `❌ (${sku} warning: ${warnings.join('; ')})`;
                  cellSeverity = 2;
                }
              } else {
                originalOpStatus = `No Print linked❌ (${sku})`;
                cellSeverity = Math.max(cellSeverity, 1);
              }
              printLogs.push(originalOpStatus);

              if (masterProductName) {
                const size = mappingInfo.size;
                if (!masterInventoryDecrements.has(masterProductName)) {
                  masterInventoryDecrements.set(masterProductName, { S: 0, M: 0, L: 0, XL: 0 });
                }
                masterInventoryDecrements.get(masterProductName)[size]++;
                masterOpStatus = `✅ (${sku})`;
              } else {
                if (productName) {
                  masterOpStatus = `⚠️ Please link with master_inventory Column C (${sku})`;
                  cellSeverity = 2; // Matches original script's overriding behavior
                } else {
                  masterOpStatus = `❌ (${sku} Both A and C empty)`;
                  cellSeverity = 2;
                }
              }
              plainLogs.push(masterOpStatus);
            }

            status = `Print: ${printLogs.join(", ")} | Plain/Ready Merch: ${plainLogs.join(", ")}`;
          }
        }
      }

      // --- Process RTO SKU ---
      if (rawRtoSku !== null && rawRtoSku !== undefined && rawRtoSku !== "") {
        const inputRto = String(rawRtoSku).trim().toUpperCase();
        if (inputRto) {
          let rtoSkusToProcess = [inputRto];
          let inputNorm = _normalizeSku(inputRto);
          if (bundleMap.has(inputNorm)) {
            rtoSkusToProcess = bundleMap.get(inputNorm);
          }

          let rtoPreFlightFailed = false;
          let rtoAbortMsg = "";
          let rtoMatches = [];
          const simulatedRtoDeductions = new Map();

          // --- RTO PRE-FLIGHT ---
          for (let sku of rtoSkusToProcess) {
            let normalizedRtoSku = _normalizeSku(sku);
            if (!skuToMappingInfo.has(normalizedRtoSku)) {
              rtoPreFlightFailed = true;
              rtoAbortMsg = `${sku} Not in Mapping Sheet`;
              break;
            }

            const inputMapping = skuToMappingInfo.get(normalizedRtoSku);
            let rtoMatch = null;

            let simulatedDeductExact = simulatedRtoDeductions.get(normalizedRtoSku) || 0;
            if (rtoInventoryInfoMap.has(normalizedRtoSku) && (rtoInventoryInfoMap.get(normalizedRtoSku).count - simulatedDeductExact) > 0) {
              rtoMatch = sku;
              simulatedRtoDeductions.set(normalizedRtoSku, simulatedDeductExact + 1);
            } else {
              for (const [inventorySku, rtoEntry] of rtoInventoryInfoMap.entries()) {
                const inventoryMapping = skuToMappingInfo.get(inventorySku);
                let simulatedDeductAlias = simulatedRtoDeductions.get(inventorySku) || 0;
                if (inventoryMapping &&
                  inventoryMapping.rowIndex === inputMapping.rowIndex &&
                  inventoryMapping.size === inputMapping.size &&
                  (rtoEntry.count - simulatedDeductAlias) > 0) {
                  rtoMatch = inventorySku;
                  simulatedRtoDeductions.set(inventorySku, simulatedDeductAlias + 1);
                  break;
                }
              }
            }

            if (!rtoMatch) {
              rtoPreFlightFailed = true;
              let hasAnyAlias = false;
              if (rtoInventoryInfoMap.has(normalizedRtoSku)) {
                hasAnyAlias = true;
              } else {
                for (const [inventorySku, rtoEntry] of rtoInventoryInfoMap.entries()) {
                  const inventoryMapping = skuToMappingInfo.get(inventorySku);
                  if (inventoryMapping && inventoryMapping.rowIndex === inputMapping.rowIndex && inventoryMapping.size === inputMapping.size) {
                    hasAnyAlias = true;
                    break;
                  }
                }
              }

              if (hasAnyAlias) {
                rtoAbortMsg = `${sku} Stock already 0`;
              } else {
                rtoAbortMsg = `${sku} missing in rto_inventory`;
              }
              break;
            } else {
              rtoMatches.push(_normalizeSku(rtoMatch));
            }
          }

          // --- RTO EXECUTION ---
          if (rtoPreFlightFailed) {
            rtoStatus = `RTO: ❌ (ABORTED: ${rtoAbortMsg})`;
            rtoSeverity = 2;
          } else {
            for (let rtoNorm of rtoMatches) {
              const rtoEntry = rtoInventoryInfoMap.get(rtoNorm);
              rtoEntry.count -= 1;
              rtoEntry.locked = Math.max(0, rtoEntry.locked - 1);
            }
            rtoStatus = `RTO: ✅ (${rtoSkusToProcess.join(", ")} restored)`;
          }
        }
      }

      // --- FORMAT LOGS ---
      let finalStatus = "";
      if (status && rtoStatus) finalStatus = status + " | " + rtoStatus;
      else if (status) finalStatus = status;
      else if (rtoStatus) finalStatus = rtoStatus;

      let highlightColor = "#000000";
      if (cellSeverity === 2) highlightColor = "#FF0000";
      else if (cellSeverity === 1) highlightColor = "#800080";

      let rtoColor = "#000000";
      if (rtoSeverity === 2) rtoColor = "#FF0000";
      else if (rtoSeverity === 1) rtoColor = "#800080";

      if (finalStatus.includes("✅") &&
        !finalStatus.includes("❌") &&
        !finalStatus.includes("⚠️") &&
        !finalStatus.includes("Warning") &&
        !finalStatus.includes("ABORTED")) {
        finalStatus = "✅";
        highlightColor = "#000000";
        rtoColor = "#000000";
      }

      results.push([finalStatus]);
      statusColors.push([highlightColor]);
      rtoColors.push([rtoColor]);
    });

    // --- 5. Apply master inventory decrements ---
    const masterStockUpdates = masterInventoryData.map(row => [row[1], row[3], row[5], row[7]]);
    masterInventoryDecrements.forEach((decrements, masterProductName) => {
      const masterEntries = masterProductInfoMap.get(masterProductName);
      if (masterEntries && masterEntries.length > 0) {
        masterEntries.forEach(masterEntry => {
          ['S', 'M', 'L', 'XL'].forEach((size, sizeIndex) => {
            if (decrements[size] > 0) {
              const currentStock = masterEntry.stocks[size];
              if (typeof currentStock === 'number') {
                masterStockUpdates[masterEntry.rowIndex][sizeIndex] = currentStock - decrements[size];
              }
            }
          });
        });
      }
    });

    // --- 6. Write the updated data back to the sheets ---
    if (results.length > 0) {
      const statusRange = testSheet.getRange(2, 10, results.length, 1);
      statusRange.setValues(results);

      statusColors.forEach((colorRow, index) => {
        statusRange.getCell(index + 1, 1).setFontColor(colorRow[0]);
      });

      testSheet.getRange(2, 8, rtoColors.length, 1).setFontColors(rtoColors);
      testSheet.getRange(2, 4, stockUpdates.length, 1).setValues(stockUpdates);

      if (masterStockUpdates.length > 0) {
        masterInventorySheet.getRange(2, 2, masterStockUpdates.length, 1).setValues(masterStockUpdates.map(row => [row[0]]));
        masterInventorySheet.getRange(2, 4, masterStockUpdates.length, 1).setValues(masterStockUpdates.map(row => [row[1]]));
        masterInventorySheet.getRange(2, 6, masterStockUpdates.length, 1).setValues(masterStockUpdates.map(row => [row[2]]));
        masterInventorySheet.getRange(2, 8, masterStockUpdates.length, 1).setValues(masterStockUpdates.map(row => [row[3]]));
      }

      if (rtoInventoryData.length > 0) {
        const rtoStockUpdates = rtoInventoryData.map(row => {
          const rawSku = row[0];
          if (rawSku !== null && rawSku !== undefined && rawSku !== "") {
            const sku = String(rawSku).trim().toUpperCase();
            let normSku = _normalizeSku(sku);
            if (sku && rtoInventoryInfoMap.has(normSku)) {
              const entry = rtoInventoryInfoMap.get(normSku);
              return [entry.count, entry.locked];
            }
          }
          return [row[1], row[2]];
        });
        rtoInventorySheet.getRange(2, 2, rtoStockUpdates.length, 2).setValues(rtoStockUpdates);
      }
    }

    scriptProperties.setProperty(LAST_RUN_KEY, new Date().getTime().toString());
    ui.alert('✅ Processing Complete!', 'Stock counts and statuses have been updated for both test sheet and master_inventory.', ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(e);
    ui.alert('An unexpected error occurred. Please check the script logs for details.');
  }
}


/**
 * NEW FUNCTION:
 * Generate Print Order when Inventory <= Threshold.
 * Increments inventory by batches (C) until > Threshold
 * and logs the order in column K.
 */
function generatePrintOrder() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Confirmation Required",
    "⚠️ I am going to clean Previous Orders from column G.\nPlease save it into your tracking sheet if not already.\n\nDo you want to proceed?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert("❌ Action cancelled. No changes made.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = ss.getSheetByName("test sheet");

  // ✅ FIX 1: Null check immediately after getSheetByName
  if (!testSheet) {
    ui.alert("Error", "Could not find 'test sheet'. Check the sheet name (case-sensitive).", ui.ButtonSet.OK);
    return;
  }

  const lastRow = testSheet.getLastRow();
  if (lastRow < 2) return;

  // Clear only G and K — same as original
  testSheet.getRange(2, 7, lastRow - 1, 1).clearContent();  // Column G
  testSheet.getRange(2, 11, lastRow - 1, 1).clearContent(); // Column K

  // Get columns B–E (Print, BatchSize, Inventory, Threshold)
  const data = testSheet.getRange(2, 2, lastRow - 1, 4).getValues();

  const updates = []; // Column D (updated inventory)
  const gUpdates = []; // Column G (batches ordered per row)
  const orders = []; // Column K (order log)

  data.forEach((row) => {
    const printName = row[0];  // Column B
    const batchSize = row[1];  // Column C
    let inventory = row[2];  // Column D
    const threshold = row[3];  // Column E

    let batchesOrdered = 0;

    if (
      typeof inventory === "number" &&
      typeof batchSize === "number" &&
      typeof threshold === "number" &&
      batchSize > 0  // ✅ Guard against infinite loop if batchSize is 0
    ) {
      while (inventory <= threshold) {
        inventory += batchSize;
        batchesOrdered++;
      }
    }

    updates.push([inventory]);
    gUpdates.push([batchesOrdered > 0 ? batchesOrdered : ""]); // ✅ Collect G values

    if (batchesOrdered > 0 && printName) {
      orders.push([`${printName} x${batchesOrdered}`]);
    }
  });

  // ✅ FIX 2: All writes outside the loop — batch API calls only
  testSheet.getRange(2, 4, updates.length, 1).setValues(updates);   // Column D
  testSheet.getRange(2, 7, gUpdates.length, 1).setValues(gUpdates);  // Column G

  if (orders.length > 0) {
    testSheet.getRange(2, 11, orders.length, 1).setValues(orders);   // Column K
  }

  ui.alert("✅ Print Order Generated!", "Inventory updated and orders logged in column K.", ui.ButtonSet.OK);
}

/**
 * NEW FUNCTION:
 * Updates PRINT_COUNT in inventory_lookup based on Mapping Sheet,
 * and populates threshold columns based on tier_rules.
 */
function updateInventoryLookupsAndThresholds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inventorySheet = ss.getSheetByName("inventory_lookup");
  const mappingSheet = ss.getSheetByName("Mapping Sheet");
  const tierRulesSheet = ss.getSheetByName("tier_rules");

  // Allow silent running for future time-based triggers
  let activeUi = null;
  try {
    activeUi = SpreadsheetApp.getUi();
  } catch (e) {
    // Running headlessly (e.g., via time-driven trigger)
  }

  if (!inventorySheet || !mappingSheet || !tierRulesSheet) {
    const errorMsg = "Could not find 'inventory_lookup', 'Mapping Sheet', or 'tier_rules'. Please check sheet names.";
    if (activeUi) activeUi.alert("Error", errorMsg, activeUi.ButtonSet.OK);
    else Logger.log(errorMsg);
    return;
  }

  // --- 1. Read all required data ---
  const mappingLastRow = Math.max(2, mappingSheet.getLastRow());
  // Fetch B (Design Name + Weight) and C (Master Product Name)
  const mappingData = mappingSheet.getRange("B2:C" + mappingLastRow).getValues();

  const inventoryLastRow = inventorySheet.getLastRow();
  if (inventoryLastRow < 2) return; // Nothing to process
  const inventoryData = inventorySheet.getRange("A2:C" + inventoryLastRow).getValues();

  const tierRulesLastRow = Math.max(2, tierRulesSheet.getLastRow());
  const tierRulesData = tierRulesSheet.getRange("A2:T" + tierRulesLastRow).getValues();

  // --- 2. Calculate Weighted Print Counts from Mapping Sheet ---
  const mappingCounts = new Map();
  mappingData.forEach(row => {
    let designInfo = row[0]; // Column B
    let name = row[1];       // Column C

    if (name) {
      name = String(name).trim(); // Trim spaces for consistency

      let weight = 0.2; // Default fallback

      if (designInfo) {
        let designStr = String(designInfo).trim();
        // Regex to extract True Velocity inside pipes, e.g. |0.2| or |2.5|
        let match = designStr.match(/\|([\d.]+)\|/);

        if (match && match[1]) {
          let parsedWeight = parseFloat(match[1]);
          // Fallback safely if it's not a number or is negative
          if (!isNaN(parsedWeight) && parsedWeight >= 0) {
            weight = parsedWeight;
          }
        }
      }

      mappingCounts.set(name, (mappingCounts.get(name) || 0) + weight);
    }
  });

  // Fix JS floating point bugs (e.g. 3.9999999 -> 4.00) so tier lookups match exactly
  mappingCounts.forEach((totalWeight, masterProduct) => {
    mappingCounts.set(masterProduct, Math.round(totalWeight * 100) / 100);
  });

  // --- 3. Parse Tier Rules ---
  const tierRulesMap = new Map();
  tierRulesData.forEach(row => {
    let category = row[0];
    if (category) {
      category = String(category).trim();
      let minDesigns = Number(row[1]) || 0;
      let alloc = row[2]; // Column C (ALLOC)
      let thresholds = row.slice(3, 15); // Columns D through O (12 values)
      // Extract the 5 new columns (P, Q, R, S, T) which map to indices 15, 16, 17, 18, 19
      // row.slice(15, 20) safely grabs exactly these 5 items, filling with undefined/"" if missing.
      let unhealthyThresholds = row.slice(15, 20).map(val => val === undefined ? "" : val);

      if (!tierRulesMap.has(category)) {
        tierRulesMap.set(category, []);
      }
      tierRulesMap.get(category).push({ minDesigns: minDesigns, alloc: alloc, thresholds: thresholds, unhealthyThresholds: unhealthyThresholds });
    }
  });

  // Sort tiers by minDesigns descending for each category
  // This allows us to easily find the correct tier by finding the first one where printCount >= minDesigns
  tierRulesMap.forEach((tiers, category) => {
    tiers.sort((a, b) => b.minDesigns - a.minDesigns);
  });

  // --- 4. Process Inventory Lookup ---
  const printCountsToUpdate = [];
  const thresholdsToUpdate = [];
  const allocToUpdate = [];
  const unhealthyToUpdate = [];
  const emptyThresholds = Array(12).fill("");
  const emptyUnhealthy = Array(5).fill("");

  inventoryData.forEach(row => {
    let colorName = String(row[0] || "").trim(); // Column A
    let category = String(row[2] || "").trim();  // Column C

    // Get print count
    let printCount = mappingCounts.get(colorName) || 0;
    printCountsToUpdate.push([printCount]);

    // Get thresholds
    let rowThresholds = emptyThresholds;
    let rowAlloc = "";
    let rowUnhealthy = emptyUnhealthy;

    // Fallback: If it's a readymade category (like readymade_tshirt) but not in tier_rules, 
    // try to fallback to 'readymade_all' if it exists.
    let searchCategory = category;
    if (category && category.startsWith("readymade_") && !tierRulesMap.has(category) && tierRulesMap.has("readymade_all")) {
      searchCategory = "readymade_all";
    }

    if (searchCategory && tierRulesMap.has(searchCategory)) {
      let categoryTiers = tierRulesMap.get(searchCategory);
      for (let i = 0; i < categoryTiers.length; i++) {
        if (printCount >= categoryTiers[i].minDesigns) {
          rowThresholds = categoryTiers[i].thresholds;
          rowAlloc = categoryTiers[i].alloc;
          rowUnhealthy = categoryTiers[i].unhealthyThresholds;
          break;
        }
      }
    }
    thresholdsToUpdate.push(rowThresholds);
    allocToUpdate.push([rowAlloc]);
    unhealthyToUpdate.push(rowUnhealthy);
  });

  // --- 5. Write Data Back ---
  // Write Print Counts to Column B
  inventorySheet.getRange(2, 2, printCountsToUpdate.length, 1).setValues(printCountsToUpdate);
  // Write Thresholds to Columns F through Q (12 columns)
  inventorySheet.getRange(2, 6, thresholdsToUpdate.length, 12).setValues(thresholdsToUpdate);
  // Write ALLOC to Column R (Column 18)
  inventorySheet.getRange(2, 18, allocToUpdate.length, 1).setValues(allocToUpdate);
  // Write Unhealthy Thresholds to Columns S through W (Column 19, 5 columns wide)
  inventorySheet.getRange(2, 19, unhealthyToUpdate.length, 5).setValues(unhealthyToUpdate);

  // Trigger highlighting function so colors are always up-to-date
  highlightMasterInventory();

  Logger.log("✅ Inventory lookups and thresholds updated successfully.");
}

/**
 * NEW FUNCTION:
 * Highlights master_inventory stock cells based on thresholds from inventory_lookup.
 */
function highlightMasterInventory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("master_inventory");
  const lookupSheet = ss.getSheetByName("inventory_lookup");

  if (!masterSheet || !lookupSheet) {
    Logger.log("Error: 'master_inventory' or 'inventory_lookup' sheet not found.");
    return;
  }

  const masterLastRow = Math.max(2, masterSheet.getLastRow());
  const lookupLastRow = Math.max(2, lookupSheet.getLastRow());

  const masterData = masterSheet.getRange("A2:I" + masterLastRow).getValues();
  // Get A to Q to encompass all thresholds
  const lookupData = lookupSheet.getRange("A2:Q" + lookupLastRow).getValues();

  // --- 1. Build Threshold Map ---
  const thresholdMap = new Map();
  lookupData.forEach(row => {
    let product = String(row[0] || "").trim();
    if (product) {
      thresholdMap.set(product, {
        s_oos: row[5], s_new: row[6],
        m_oos: row[8], m_new: row[9],
        l_oos: row[11], l_new: row[12],
        xl_oos: row[14], xl_new: row[15]
      });
    }
  });

  // Helper to determine color based on rules
  function getColor(stock, virtualStock, oos_th, new_order_th) {
    if (stock === "" || stock === null) return null;
    let stockNum = Number(stock);
    if (isNaN(stockNum)) return null;

    let vStockNum = Number(virtualStock);
    if (isNaN(vStockNum)) vStockNum = 0; // Treat empty/invalid virtual stock as 0

    let totalStock = stockNum + vStockNum;

    // Rule 1: RED if (physical + virtual) <= OOS_TH
    let oosStr = String(oos_th || "").trim();
    if (oosStr !== "") {
      let oos = Number(oosStr);
      if (!isNaN(oos) && totalStock <= oos) {
        return "#FF0000"; // Red
      }
    }

    // Rule 2: ORANGE if (physical + virtual) <= NEW_ORDER_TH
    let newOrderStr = String(new_order_th || "").trim();
    if (newOrderStr !== "") {
      let newOrder = Number(newOrderStr);
      if (!isNaN(newOrder) && totalStock <= newOrder) {
        return "#FFA500"; // Orange
      }
    }

    // Rule 3: Clear otherwise
    return null;
  }

  // --- 2. Process Master Inventory Data ---
  const bgS = [];
  const bgM = [];
  const bgL = [];
  const bgXL = [];
  const colLUpdates = []; // To track if >= 2 sizes are Red

  masterData.forEach(row => {
    let product = String(row[0] || "").trim();
    let thresholds = thresholdMap.get(product);

    if (thresholds) {
      let colorS = getColor(row[1], row[2], thresholds.s_oos, thresholds.s_new);
      let colorM = getColor(row[3], row[4], thresholds.m_oos, thresholds.m_new);
      let colorL = getColor(row[5], row[6], thresholds.l_oos, thresholds.l_new);
      let colorXL = getColor(row[7], row[8], thresholds.xl_oos, thresholds.xl_new);

      bgS.push([colorS]);
      bgM.push([colorM]);
      bgL.push([colorL]);
      bgXL.push([colorXL]);

      // Count how many sizes are RED
      let redCount = 0;
      if (colorS === "#FF0000") redCount++;
      if (colorM === "#FF0000") redCount++;
      if (colorL === "#FF0000") redCount++;
      if (colorXL === "#FF0000") redCount++;

      // Update Column L flag
      colLUpdates.push([redCount >= 2 ? "Yes" : "No"]);
    } else {
      bgS.push([null]);
      bgM.push([null]);
      bgL.push([null]);
      bgXL.push([null]);
      colLUpdates.push([product ? "No" : ""]); // If empty row, leave blank
    }
  });

  // --- 3. Apply Backgrounds and Column L ---
  if (bgS.length > 0) {
    masterSheet.getRange(2, 2, bgS.length, 1).setBackgrounds(bgS); // S (Col B)
    masterSheet.getRange(2, 4, bgM.length, 1).setBackgrounds(bgM); // M (Col D)
    masterSheet.getRange(2, 6, bgL.length, 1).setBackgrounds(bgL); // L (Col F)
    masterSheet.getRange(2, 8, bgXL.length, 1).setBackgrounds(bgXL); // XL (Col H)
    masterSheet.getRange(2, 12, colLUpdates.length, 1).setValues(colLUpdates); // Col L Update
  }

  Logger.log("✅ Master inventory highlighting applied.");
}

// ==============================================================================
// FABRIC ROLL SUGGESTION & VIRTUAL INVENTORY SYSTEM
// ==============================================================================

/**
 * Recalculates S_V, M_V, L_V, XL_V based on hidden JSON in active suggestion rows.
 */
function updateVirtualInventorySums() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("master_inventory");
  if (!masterSheet) return;

  const lastRow = Math.max(2, masterSheet.getLastRow());
  const masterData = masterSheet.getRange(2, 1, lastRow - 1, 27).getValues();

  const virtualTotals = new Map();

  for (let i = 0; i < masterData.length; i++) {
    const row = masterData[i];
    const suggProduct = String(row[17] || "").trim(); // Col R
    const hiddenJson = String(row[26] || "").trim(); // Col AB

    if (suggProduct && hiddenJson && hiddenJson.startsWith("{")) {
      try {
        const pieces = JSON.parse(hiddenJson);
        if (!virtualTotals.has(suggProduct)) {
          virtualTotals.set(suggProduct, { S: 0, M: 0, L: 0, XL: 0 });
        }
        let totals = virtualTotals.get(suggProduct);
        totals.S += (pieces.S || 0);
        totals.M += (pieces.M || 0);
        totals.L += (pieces.L || 0);
        totals.XL += (pieces.XL || 0);
      } catch (e) { }
    }
  }

  const vUpdates = [];
  for (let i = 0; i < masterData.length; i++) {
    const product = String(masterData[i][0] || "").trim();
    if (product && virtualTotals.has(product)) {
      const t = virtualTotals.get(product);
      vUpdates.push([t.S, t.M, t.L, t.XL]);
    } else {
      vUpdates.push([0, 0, 0, 0]);
    }
  }

  if (vUpdates.length > 0) {
    masterSheet.getRange(2, 3, vUpdates.length, 1).setValues(vUpdates.map(r => [r[0]])); // C
    masterSheet.getRange(2, 5, vUpdates.length, 1).setValues(vUpdates.map(r => [r[1]])); // E
    masterSheet.getRange(2, 7, vUpdates.length, 1).setValues(vUpdates.map(r => [r[2]])); // G
    masterSheet.getRange(2, 9, vUpdates.length, 1).setValues(vUpdates.map(r => [r[3]])); // I
  }
}

/**
 * Helper to compute the total LOCKED virtual stock for each product
 */
function getLockedVirtualStock(masterData) {
  const lockedTotals = new Map();
  for (let i = 0; i < masterData.length; i++) {
    let row = masterData[i];
    let suggProduct = String(row[17] || "").trim();
    let isLocked = (row[22] === true || row[22] === "TRUE"); // Col X
    let hiddenJson = String(row[26] || "").trim(); // Col AB

    if (suggProduct && isLocked && hiddenJson && hiddenJson.startsWith("{")) {
      try {
        const pieces = JSON.parse(hiddenJson);
        if (!lockedTotals.has(suggProduct)) {
          lockedTotals.set(suggProduct, { S: 0, M: 0, L: 0, XL: 0 });
        }
        let totals = lockedTotals.get(suggProduct);
        totals.S += (pieces.S || 0);
        totals.M += (pieces.M || 0);
        totals.L += (pieces.L || 0);
        totals.XL += (pieces.XL || 0);
      } catch (e) { }
    }
  }
  return lockedTotals;
}

/**
 * Runs daily via Time-Based Trigger.
 */
function generateFabricRollSuggestions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("master_inventory");
  const lookupSheet = ss.getSheetByName("inventory_lookup");

  if (!masterSheet || !lookupSheet) return;

  const masterLastRow = Math.max(2, masterSheet.getLastRow());
  const lookupLastRow = Math.max(2, lookupSheet.getLastRow());

  const masterData = masterSheet.getRange(2, 1, masterLastRow - 1, 27).getValues();
  const lookupData = lookupSheet.getRange(2, 1, lookupLastRow - 1, 17).getValues();

  const lookupMap = new Map();
  lookupData.forEach(row => {
    let product = String(row[0] || "").trim();
    if (product) {
      lookupMap.set(product, {
        fabricOrReady: String(row[3] || "").trim().toUpperCase(),
        pcsPerRoll: Number(row[4]) || 60,
        targetS: Number(row[7]) || 0,
        targetM: Number(row[10]) || 0,
        targetL: Number(row[13]) || 0,
        targetXL: Number(row[16]) || 0,
        oosS: row[5], newOrderS: row[6],
        oosM: row[8], newOrderM: row[9],
        oosL: row[11], newOrderL: row[12],
        oosXL: row[14], newOrderXL: row[15]
      });
    }
  });

  const unpaidRows = new Map();
  const unlockedRows = [];

  for (let i = 0; i < masterData.length; i++) {
    let row = masterData[i];
    let suggProduct = String(row[17] || "").trim(); // Col R
    let isPaid = (row[20] === true || row[20] === "TRUE"); // Col U
    let isLocked = (row[22] === true || row[22] === "TRUE"); // Col W

    if (suggProduct) {
      if (!isPaid) {
        unpaidRows.set(suggProduct, i);
      }
      if (isPaid && !isLocked) {
        unlockedRows.push({ rowIndex: i, product: suggProduct, paidRolls: Number(row[19]) || 0 });
      }
    }
  }

  let firstEmptySuggestionRow = -1;
  for (let i = 0; i < masterData.length; i++) {
    if (String(masterData[i][17] || "").trim() === "") {
      firstEmptySuggestionRow = i;
      break;
    }
  }
  if (firstEmptySuggestionRow === -1) firstEmptySuggestionRow = masterData.length;

  const now = new Date();
  const formattedDate = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss");

  // ==============================================================================
  // STEP 1: Process Unlocked Rows First (Dynamic Ratio Recalculation)
  // ==============================================================================
  const lockedTotals = getLockedVirtualStock(masterData);
  let unlockedUpdates = [];

  unlockedRows.forEach(uRow => {
    let product = uRow.product;
    if (!lookupMap.has(product)) return;
    const lookup = lookupMap.get(product);

    let pS = 0, pM = 0, pL = 0, pXL = 0;
    for (let m = 0; m < masterData.length; m++) {
      if (String(masterData[m][0] || "").trim() === product) {
        pS = isNaN(Number(masterData[m][1])) ? 0 : Number(masterData[m][1]);
        pM = isNaN(Number(masterData[m][3])) ? 0 : Number(masterData[m][3]);
        pL = isNaN(Number(masterData[m][5])) ? 0 : Number(masterData[m][5]);
        pXL = isNaN(Number(masterData[m][7])) ? 0 : Number(masterData[m][7]);
        break;
      }
    }

    let lS = 0, lM = 0, lL = 0, lXL = 0;
    if (lockedTotals.has(product)) {
      const lt = lockedTotals.get(product);
      lS = lt.S; lM = lt.M; lL = lt.L; lXL = lt.XL;
    }

    let intendedS = Math.max(0, lookup.targetS - (pS + lS));
    let intendedM = Math.max(0, lookup.targetM - (pM + lM));
    let intendedL = Math.max(0, lookup.targetL - (pL + lL));
    let intendedXL = Math.max(0, lookup.targetXL - (pXL + lXL));

    let ratio = calculatePerfectRatio(intendedS, intendedM, intendedL, intendedXL);
    let totalPieces = uRow.paidRolls * lookup.pcsPerRoll;
    let distributed = distributePieces(ratio, totalPieces);

    let ratioStr = ratio.join(":");
    let hiddenJson = JSON.stringify({ S: distributed[0], M: distributed[1], L: distributed[2], XL: distributed[3] });

    unlockedUpdates.push({ row: uRow.rowIndex + 2, col: 22, val: ratioStr });
    unlockedUpdates.push({ row: uRow.rowIndex + 2, col: 27, val: hiddenJson });
    unlockedUpdates.push({ row: uRow.rowIndex + 2, col: 26, val: formattedDate });
  });

  if (unlockedUpdates.length > 0) {
    unlockedUpdates.forEach(u => masterSheet.getRange(u.row, u.col).setValue(u.val));
    // Immediately recalculate virtual inventory so the rest of the script sees fresh virtual numbers
    updateVirtualInventorySums();
  }

  // ==============================================================================
  // STEP 2: Suggest Rolls based on ANY breach using FRESH data
  // ==============================================================================
  const freshMasterData = masterSheet.getRange(2, 1, masterLastRow - 1, 27).getValues();
  const updates = [];
  const newCheckboxes = [];

  // Identify which products already have at least one Paid row in the suggestions block
  const productsWithPaidRows = new Set();
  for (let i = 0; i < freshMasterData.length; i++) {
    let suggProduct = String(freshMasterData[i][17] || "").trim(); // Col R
    let isPaid = (freshMasterData[i][20] === true || freshMasterData[i][20] === "TRUE"); // Col U
    if (suggProduct && isPaid) {
      productsWithPaidRows.add(suggProduct);
    }
  }

  for (let i = 0; i < freshMasterData.length; i++) {
    const row = freshMasterData[i];
    const product = String(row[0] || "").trim();
    if (!product || !lookupMap.has(product)) continue;

    const lookup = lookupMap.get(product);
    if (lookup.fabricOrReady !== "FABRIC") continue;

    let needsRolls = false;
    let hasPaidRow = productsWithPaidRows.has(product);

    // Scenario A: Emergency Trigger (Orange Breach)
    const isBreached = (physical, virtual, new_order_th) => {
      let p = Number(physical); if (isNaN(p)) return false;
      let v = Number(virtual); if (isNaN(v)) v = 0;
      let total = p + v;
      let newStr = String(new_order_th || "").trim();
      if (newStr !== "" && !isNaN(Number(newStr)) && total <= Number(newStr)) return true;
      return false;
    };

    // Scenario B: Top-Up Trigger (Target - 5 buffer)
    const isTargetBreached = (physical, virtual, target_stock) => {
      let p = Number(physical); if (isNaN(p)) return false;
      let v = Number(virtual); if (isNaN(v)) v = 0;
      let total = p + v;
      let target = Number(target_stock);
      if (isNaN(target)) return false;
      return total < (target - 5);
    };

    if (hasPaidRow) {
      if (
        isTargetBreached(row[1], row[2], lookup.targetS) ||
        isTargetBreached(row[3], row[4], lookup.targetM) ||
        isTargetBreached(row[5], row[6], lookup.targetL) ||
        isTargetBreached(row[7], row[8], lookup.targetXL)
      ) {
        needsRolls = true;
      }
    } else {
      if (
        isBreached(row[1], row[2], lookup.newOrderS) ||
        isBreached(row[3], row[4], lookup.newOrderM) ||
        isBreached(row[5], row[6], lookup.newOrderL) ||
        isBreached(row[7], row[8], lookup.newOrderXL)
      ) {
        needsRolls = true;
      }
    }

    if (needsRolls) {
      let pS = isNaN(Number(row[1])) ? 0 : Number(row[1]);
      let pM = isNaN(Number(row[3])) ? 0 : Number(row[3]);
      let pL = isNaN(Number(row[5])) ? 0 : Number(row[5]);
      let pXL = isNaN(Number(row[7])) ? 0 : Number(row[7]);

      let vS = isNaN(Number(row[2])) ? 0 : Number(row[2]);
      let vM = isNaN(Number(row[4])) ? 0 : Number(row[4]);
      let vL = isNaN(Number(row[6])) ? 0 : Number(row[6]);
      let vXL = isNaN(Number(row[8])) ? 0 : Number(row[8]);

      let intendedS = Math.max(0, lookup.targetS - (pS + vS));
      let intendedM = Math.max(0, lookup.targetM - (pM + vM));
      let intendedL = Math.max(0, lookup.targetL - (pL + vL));
      let intendedXL = Math.max(0, lookup.targetXL - (pXL + vXL));

      let totalIntended = intendedS + intendedM + intendedL + intendedXL;
      let suggestedRolls = Math.ceil(totalIntended / lookup.pcsPerRoll);

      if (suggestedRolls >= 1) {
        if (unpaidRows.has(product)) {
          let sRowIdx = unpaidRows.get(product);
          updates.push({ row: sRowIdx + 2, col: 19, val: suggestedRolls });
          updates.push({ row: sRowIdx + 2, col: 26, val: formattedDate });
        } else {
          let nRowIdx = firstEmptySuggestionRow++;
          updates.push({ row: nRowIdx + 2, col: 18, val: product });
          updates.push({ row: nRowIdx + 2, col: 19, val: suggestedRolls });
          updates.push({ row: nRowIdx + 2, col: 21, val: false });
          updates.push({ row: nRowIdx + 2, col: 23, val: false });
          updates.push({ row: nRowIdx + 2, col: 24, val: false });
          updates.push({ row: nRowIdx + 2, col: 25, val: formattedDate });
          updates.push({ row: nRowIdx + 2, col: 26, val: formattedDate });
          newCheckboxes.push(nRowIdx + 2);
          masterSheet.getRange(nRowIdx + 2, 18, 1, 1).setBackground("black").setFontColor("white");
        }
      } else {
        if (!hasPaidRow && unpaidRows.has(product)) {
          let sRowIdx = unpaidRows.get(product);
          masterSheet.getRange(sRowIdx + 2, 17, 1, 11).clearContent().clearDataValidations().setBackground(null);
        }
      }
    } else {
      if (!hasPaidRow && unpaidRows.has(product)) {
        let sRowIdx = unpaidRows.get(product);
        masterSheet.getRange(sRowIdx + 2, 17, 1, 11).clearContent().clearDataValidations().setBackground(null);
      }
    }
  }

  if (updates.length > 0) {
    updates.forEach(u => masterSheet.getRange(u.row, u.col).setValue(u.val));
  }

  newCheckboxes.forEach(r => {
    masterSheet.getRange(r, 21).insertCheckboxes();
    masterSheet.getRange(r, 23).insertCheckboxes();
    masterSheet.getRange(r, 24).insertCheckboxes();
  });

  SpreadsheetApp.flush();
  highlightMasterInventory();
}

// ==============================================================================
// FABRIC RATIO ALGORITHMS
// ==============================================================================

function calculatePerfectRatio(intendedS, intendedM, intendedL, intendedXL) {
  let raw = [intendedS, intendedM, intendedL, intendedXL];
  let rawSum = raw.reduce((a, b) => a + b, 0);

  if (rawSum === 0) return [0, 0, 0, 0];

  let baselinePct = raw.map(v => v / rawSum);
  let maxRaw = Math.max(...raw);

  let bestScore = Infinity;
  let bestRatio = [0, 0, 0, 0];

  for (let maxTgt = 1; maxTgt <= 5; maxTgt++) {
    let multiplier = maxTgt / maxRaw;
    let candidate = raw.map(v => {
      return Math.round(v * multiplier);
    });

    let candSum = candidate.reduce((a, b) => a + b, 0);
    let candPct = candidate.map(v => v / candSum);

    let error = 0;
    for (let i = 0; i < 4; i++) {
      error += Math.abs(baselinePct[i] - candPct[i]);
    }

    // Add 1% complexity penalty per maxTgt step
    let score = error + (maxTgt * 0.01);

    if (score < bestScore) {
      bestScore = score;
      bestRatio = candidate;
    }
  }
  return bestRatio;
}

function distributePieces(ratio, totalPieces) {
  let ratioSum = ratio.reduce((a, b) => a + b, 0);
  if (ratioSum === 0) return [0, 0, 0, 0];

  let exact = ratio.map(r => (r / ratioSum) * totalPieces);
  let floored = exact.map(Math.floor);
  let remainders = exact.map((v, i) => ({ val: v - floored[i], idx: i }));

  let currentSum = floored.reduce((a, b) => a + b, 0);
  let needed = totalPieces - currentSum;

  remainders.sort((a, b) => b.val - a.val);

  for (let i = 0; i < needed; i++) {
    floored[remainders[i].idx]++;
  }

  return floored;
}

// ==============================================================================
// AUTOMATIC EVENT HANDLER (Triggers on checkbox click)
// ==============================================================================

function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== "master_inventory") return;

  const col = e.range.getColumn();
  const row = e.range.getRow();

  // U(21), W(23), X(24)
  if (![21, 23, 24].includes(col)) return;

  const val = e.value;
  if (val !== "TRUE") return;

  const ss = e.source;
  const now = new Date();
  const formattedDate = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss");

  const rowData = sh.getRange(row, 17, 1, 11).getValues()[0];
  let product = String(rowData[1] || "").trim();
  if (!product) {
    e.range.setValue(false);
    return;
  }

  // U: Paid
  if (col === 21) {
    let paidRolls = Number(rowData[3]);
    if (isNaN(paidRolls) || paidRolls <= 0) {
      ss.toast("Please enter a valid number of paid rolls in column T before clicking Paid.");
      e.range.setValue(false);
      return;
    }

    // Automatically generate ratio
    generateRatioForSpecificRow(ss, sh, row, product, paidRolls, formattedDate);

    e.range.setBackground("#d9ead3");
    sh.getRange(row, 26).setValue(formattedDate);
    ss.toast("Action successful. Recalculating suggestions...", "Processing");
    SpreadsheetApp.flush();
    generateFabricRollSuggestions();
  }

  // W: Lock Ratio
  else if (col === 23) {
    if (rowData[4] !== true && rowData[4] !== "TRUE") {
      ss.toast("You must mark this as Paid (check Col U) before locking it.");
      e.range.setValue(false);
      return;
    }
    e.range.setBackground("#d9ead3");
    sh.getRange(row, 26).setValue(formattedDate);
    ss.toast("Action successful. Recalculating suggestions...", "Processing");
    SpreadsheetApp.flush();
    generateFabricRollSuggestions();
  }

  // X: Remove Virtual Stock (Archive)
  else if (col === 24) {
    if (rowData[6] !== true && rowData[6] !== "TRUE") {
      ss.toast("You must Lock the ratio (check Col W) before you can remove the virtual stock.");
      e.range.setValue(false);
      return;
    }

    e.range.setBackground("#d9ead3");

    let historySheet = ss.getSheetByName("rolls_history");
    if (!historySheet) {
      historySheet = ss.insertSheet("rolls_history");
      historySheet.appendRow(["Paid Date", "Product", "Suggested Rolls", "Paid Rolls", "Paid", "Ratio", "Lock", "Remove", "Create Date", "Last Update Date", "Hidden Pieces"]);
    }

    historySheet.appendRow(rowData);

    // Clear from master inventory
    sh.getRange(row, 17, 1, 11).clearContent();
    sh.getRange(row, 17, 1, 11).clearDataValidations();
    sh.getRange(row, 17, 1, 11).setBackground(null);

    SpreadsheetApp.flush();
    updateVirtualInventorySums();
    ss.toast("Virtual stock removed. Recalculating suggestions...", "Processing");
    SpreadsheetApp.flush();
    generateFabricRollSuggestions();
  }
}

function generateRatioForSpecificRow(ss, sh, rowNum, product, paidRolls, formattedDate) {
  // Read Master Inventory for physical stock
  let pS = 0, pM = 0, pL = 0, pXL = 0;
  const mData = sh.getRange(2, 1, sh.getLastRow(), 8).getValues();
  for (let i = 0; i < mData.length; i++) {
    if (String(mData[i][0]).trim() === product) {
      pS = Number(mData[i][1]) || 0;
      pM = Number(mData[i][3]) || 0;
      pL = Number(mData[i][5]) || 0;
      pXL = Number(mData[i][7]) || 0;
      break;
    }
  }

  // Determine existing LOCKED virtual stock for this product
  let lS = 0, lM = 0, lL = 0, lXL = 0;
  const fullData = sh.getRange(2, 1, sh.getLastRow(), 27).getValues();
  const lockedTotals = getLockedVirtualStock(fullData);
  if (lockedTotals.has(product)) {
    const lt = lockedTotals.get(product);
    lS = lt.S; lM = lt.M; lL = lt.L; lXL = lt.XL;
  }

  // Read target stock and PCS_PER_ROLL from Lookup
  const lSh = ss.getSheetByName("inventory_lookup");
  const lData = lSh.getRange(2, 1, lSh.getLastRow(), 17).getValues();
  let tS = 0, tM = 0, tL = 0, tXL = 0, pcsPerRoll = 60;
  for (let i = 0; i < lData.length; i++) {
    if (String(lData[i][0]).trim() === product) {
      pcsPerRoll = Number(lData[i][4]) || 60;
      tS = Number(lData[i][7]) || 0;
      tM = Number(lData[i][10]) || 0;
      tL = Number(lData[i][13]) || 0;
      tXL = Number(lData[i][16]) || 0;
      break;
    }
  }

  // Intended calculation includes Locked Virtual Stock
  let intendedS = Math.max(0, tS - (pS + lS));
  let intendedM = Math.max(0, tM - (pM + lM));
  let intendedL = Math.max(0, tL - (pL + lL));
  let intendedXL = Math.max(0, tXL - (pXL + lXL));

  let ratio = calculatePerfectRatio(intendedS, intendedM, intendedL, intendedXL);
  let totalPieces = paidRolls * pcsPerRoll;
  let distributed = distributePieces(ratio, totalPieces);

  let ratioStr = ratio.join(":");
  let hiddenJson = JSON.stringify({ S: distributed[0], M: distributed[1], L: distributed[2], XL: distributed[3] });

  sh.getRange(rowNum, 22).setValue(ratioStr); // V
  sh.getRange(rowNum, 27).setValue(hiddenJson); // AA

  SpreadsheetApp.flush();
  updateVirtualInventorySums();
}

/**
 * Cleans up the rto_inventory sheet by removing rows where the OnHand stock is 0.
 * Uses a bulk read/write method to ensure no empty rows are left behind.
 */
function cleanZeroRtoStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { }
  const rtoSheet = ss.getSheetByName("rto_inventory");

  if (!rtoSheet) {
    //ui.alert("Error", "Could not find 'rto_inventory' sheet.", ui.ButtonSet.OK);
    return;
  }

  const lastRow = rtoSheet.getLastRow();
  if (lastRow <= 1) {
    //ui.alert("Notice", "The rto_inventory sheet is already empty.", ui.ButtonSet.OK);
    return;
  }

  // Read all data (excluding headers)
  const data = rtoSheet.getRange(2, 1, lastRow - 1, 4).getValues();

  // Filter out rows where OnHand is 0 
  // Columns: [SKU, Count, Locked, OnHand] -> OnHand is index 3
  const filteredData = data.filter(row => {
    const onHand = parseFloat(row[3]) || 0;
    return onHand !== 0;
  });

  // NEW: Re-inject the live formula back into Column D
  filteredData.forEach((row, index) => {
    let sheetRow = index + 2; // Offset by 2 because row 1 is headers
    row[3] = `=B${sheetRow}-C${sheetRow}`;
  });

  // Clear the existing data from Row 2 downwards
  rtoSheet.getRange(2, 1, lastRow - 1, 4).clearContent();

  // Write the compacted, filtered data back
  if (filteredData.length > 0) {
    rtoSheet.getRange(2, 1, filteredData.length, 4).setValues(filteredData);
  }
}


// ==============================================================================
// BAYESIAN DYNAMIC WEIGHTING SYSTEM
// ==============================================================================

function updateDynamicWeights() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { }

  // Update the Out-Of-Stock Ledger immediately before running weights math
  recordMidnightOOSSnapshot();

  const mappingSheet = ss.getSheetByName("Mapping Sheet");
  const masterInventorySheet = ss.getSheetByName("master_inventory");
  const ledgerSheet = ss.getSheetByName("oos_ledger");

  if (!mappingSheet || !masterInventorySheet || !ledgerSheet) {
    if (ui) ui.alert("Error", "Could not find required sheets (Mapping Sheet, master_inventory, oos_ledger). Ensure you have completed the manual setup.", ui.ButtonSet.OK);
    return;
  }

  // Source Spreadsheets
  const sourceSpreadsheetId = "10OCuU7CFxuOtG2z6Q6fHWVtHbaVLg4wpZw-XeTzqq1E";
  let sourceSS;
  try {
    sourceSS = SpreadsheetApp.openById(sourceSpreadsheetId);
  } catch (e) {
    if (ui) ui.alert("System Error", "Exact Error: " + e.toString(), ui.ButtonSet.OK);
    return;
  }

  const now = Date.now();
  const ONE_DAY_MS = 24 * 60 * 60 * 1000;
  const HISTORY_WINDOW_ACTIVE_DAYS = 120;
  const LOOKBACK_HORIZON_MS = 240 * ONE_DAY_MS; // Max calendar days to look back

  // Helper for ultra-fast native timezone formatting (100x faster than Utilities.formatDate)
  function getLocalYMD(d) {
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, '0') + "-" + String(d.getDate()).padStart(2, '0');
  }

  // --- PHASE 1: Initialization & The Active-Day Mapper ---

  // 1.1 Read Ledger and build OOS Map
  const ledgerData = ledgerSheet.getDataRange().getValues();
  const oosLog = new Map(); // BlankType -> Set of OOS Date Strings (YYYY-MM-DD)

  for (let i = 1; i < ledgerData.length; i++) {
    const rawDate = ledgerData[i][0];
    const blankType = String(ledgerData[i][1] || "").trim();
    if (!rawDate || !blankType) continue;

    const d = new Date(rawDate);
    if (isNaN(d.getTime())) continue;
    const dateStr = getLocalYMD(d);

    if (!oosLog.has(blankType)) {
      oosLog.set(blankType, new Set());
    }
    oosLog.get(blankType).add(dateStr);
  }

  // Helper to determine if a blank was active on a given day
  function isBlankActiveOnDate(blankType, timestampMs) {
    const d = new Date(timestampMs);
    const dateStr = getLocalYMD(d);
    if (oosLog.has(blankType) && oosLog.get(blankType).has(dateStr)) {
      return false; // It was OOS
    }
    return true; // Assume IN-STOCK if not found
  }

  // 1.2 Build current OOS map (from master_inventory)
  const masterData = masterInventorySheet.getDataRange().getValues();
  const currentOosMap = new Map();
  for (let i = 1; i < masterData.length; i++) {
    const blankType = String(masterData[i][0] || "").trim();
    const isOos = String(masterData[i][11] || "").trim().toLowerCase() === "yes";
    if (blankType) currentOosMap.set(blankType, isOos);
  }

  // 1.3 Build Master Dictionary
  const mappingData = mappingSheet.getDataRange().getValues();
  const bundleSheet = ss.getSheetByName("Bundle_SKU_Mapping");
  const bundleData = bundleSheet ? bundleSheet.getDataRange().getValues() : [];
  const bundleMap = new Map();
  if (bundleData.length > 0) {
    for (let i = 1; i < bundleData.length; i++) {
      const row = bundleData[i];
      const rawBundleSku = String(row[0] || "").trim().toUpperCase();
      if (rawBundleSku) {
        const bundleNorm = _normalizeSku(rawBundleSku);
        const children = [];
        for (let c = 1; c < row.length; c++) {
          const child = String(row[c] || "").trim().toUpperCase();
          if (child) children.push(child);
        }
        if (children.length > 0) bundleMap.set(bundleNorm, children);
      }
    }
  }
  const masterDict = new Map();
  const lookupMap = new Map();
  const duplicateSkus = new Set();

  for (let i = 1; i < mappingData.length; i++) {
    const row = mappingData[i];
    const printCol = String(row[1] || "").trim();
    const blankType = String(row[2] || "").trim();

    let skuCount = 0;
    for (let colIdx = 3; colIdx < row.length; colIdx++) {
      const rawSku = String(row[colIdx] || "").trim().toUpperCase();
      if (rawSku) {
        skuCount++;
        let norm = _normalizeSku(rawSku);
        if (lookupMap.has(norm)) duplicateSkus.add(rawSku);
        else lookupMap.set(norm, i);
      }
    }

    if (skuCount > 0) {
      let launchDateMatch = printCol.match(/<([^>]+)>/);
      let launchDateStr = launchDateMatch ? launchDateMatch[1] : null;
      let launchDateMs = now;
      if (launchDateStr) {
        let parsedDate = Date.parse(launchDateStr);
        if (!isNaN(parsedDate)) launchDateMs = parsedDate;
      }

      let weightMatch = printCol.match(/\[(0(?:\.\d+)?|1(?:\.0+)?)\]/);
      let currentWeight = weightMatch ? parseFloat(weightMatch[1]) : 0.6;

      let velocityMatch = printCol.match(/\|([\d.]+)\|/);
      let currentVelocity = velocityMatch ? parseFloat(velocityMatch[1]) : 0.0;

      masterDict.set(i, {
        salesEMA: 0,
        score15: 0,
        score75: 0,
        launchDateMs: launchDateMs,
        blankType: blankType,
        currentWeight: currentWeight,
        currentVelocity: currentVelocity,
        printStr: printCol
      });
    }
  }

  // 1.4 Precompute Active Day Indexing (O(1) Optimization for Sales Loop)
  const activeDaysMap = new Map();
  const uniqueBlanks = new Set();
  masterDict.forEach(val => uniqueBlanks.add(val.blankType));

  uniqueBlanks.forEach(blank => {
    let activeDayCount = 0;
    let mappingArray = new Array(245); // up to 240 days lookback
    for (let d = 0; d <= 240; d++) {
      let checkTime = now - (d * ONE_DAY_MS);
      mappingArray[d] = activeDayCount; // Assign before incrementing (so today is 0 days old)
      if (isBlankActiveOnDate(blank, checkTime)) {
        activeDayCount++;
      }
    }
    activeDaysMap.set(blank, mappingArray);
  });

  // --- PHASE 2: Data Ingestion (Active-Time Filter) ---
  const missingSkus = new Set();
  const uniqueMissingSkus = new Set();
  let totalProcessedSkus = new Set();

  function processSalesSheet(sheetName, skuColIdx, dateColIdx, isFlipkart = false) {
    const sheet = sourceSS.getSheetByName(sheetName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      let rawDate = data[i][dateColIdx];
      if (!rawDate) continue;

      let saleTime = 0;
      if (rawDate instanceof Date) {
        saleTime = rawDate.getTime();
      } else {
        let numDate = Number(rawDate);
        if (!isNaN(numDate) && numDate > 20000 && numDate < 100000) {
          saleTime = Math.round((numDate - 25569) * 86400 * 1000);
        } else {
          saleTime = Date.parse(rawDate);
        }
      }

      // OPTIMIZATION: Filter out old sales BEFORE doing heavy String coercion & Regex
      if (isNaN(saleTime) || now - saleTime > LOOKBACK_HORIZON_MS) continue;

      let rawSku = String(data[i][skuColIdx] || "").trim().toUpperCase();
      if (!rawSku) continue;

      // Robust SKU Cleaning (removes all stray quotes like """)
      rawSku = rawSku.replace(/"/g, '').trim();
      if (isFlipkart) {
        rawSku = rawSku.replace(/^SKU:\s*/i, '').trim();
      }

      totalProcessedSkus.add(rawSku);

      let skusToProcess = [rawSku];
      let bundleNorm = _normalizeSku(rawSku);
      if (bundleMap.has(bundleNorm)) {
        skusToProcess = bundleMap.get(bundleNorm);
      }

      for (let sku of skusToProcess) {
        let normalizedSku = _normalizeSku(sku);
        if (lookupMap.has(normalizedSku)) {
          let parentRow = lookupMap.get(normalizedSku);
          if (masterDict.has(parentRow)) {
            let design = masterDict.get(parentRow);

            // ACTIVE-TIME FILTER
            if (isBlankActiveOnDate(design.blankType, saleTime)) {
              let calendarDaysAgo = Math.floor(Math.max(0, (now - saleTime) / ONE_DAY_MS));

              if (calendarDaysAgo <= 240) {
                let activeDaysAgo = activeDaysMap.get(design.blankType)[calendarDaysAgo];

                if (activeDaysAgo <= HISTORY_WINDOW_ACTIVE_DAYS) {
                  // The Exponential Decay Distortion Fix
                  let recencyWeight = Math.exp(-activeDaysAgo / 30);
                  design.salesEMA += recencyWeight;

                  let recencyWeight15 = Math.exp(-activeDaysAgo / 15.0);
                  let recencyWeight75 = Math.exp(-activeDaysAgo / 75.0);
                  design.score15 += recencyWeight15;
                  design.score75 += recencyWeight75;
                }
              }
            }
          }
        } else {
          let dateStr = getLocalYMD(new Date(saleTime));
          if (skusToProcess.length > 1) {
            missingSkus.add(`${sheetName} | ${dateStr} | ${rawSku} (Child: ${sku})`);
          } else {
            missingSkus.add(`${sheetName} | ${dateStr} | ${rawSku}`);
          }
          uniqueMissingSkus.add(rawSku);
        }
      }
    }
  }

  processSalesSheet("WEBSITE_FORWARD", 20, 15);
  processSalesSheet("myntra_forward", 2, 0);
  processSalesSheet("AJIO_FORWARD", 4, 1);
  processSalesSheet("FLIPKART_FORWARD", 3, 0, true);

  // --- PHASE 3: Mathematical Transformation (Micro-Economies Model) ---
  const blankDesignCount = new Map(); // blankType -> count of designs
  const blankSalesMap = new Map(); // blankType -> Array of active sales

  masterDict.forEach((val) => {
    // 1. Count designs per blank
    let count = blankDesignCount.get(val.blankType) || 0;
    blankDesignCount.set(val.blankType, count + 1);

    // 2. Group active sales by blank
    if (currentOosMap.get(val.blankType) !== true && val.salesEMA > 0) {
      if (!blankSalesMap.has(val.blankType)) {
        blankSalesMap.set(val.blankType, []);
      }
      blankSalesMap.get(val.blankType).push(val.salesEMA);
    }
  });

  const blankCeilingMap = new Map(); // blankType -> local maxSales (95th percentile)

  blankSalesMap.forEach((salesArray, blankType) => {
    salesArray.sort((a, b) => a - b);
    let p95Index = Math.floor(salesArray.length * 0.95);
    if (p95Index >= salesArray.length) p95Index = salesArray.length - 1;
    let localMax = salesArray[p95Index];
    if (localMax <= 0) localMax = 1;
    blankCeilingMap.set(blankType, localMax);
  });

  const finalUpdates = [];

  masterDict.forEach((val, rowIndex) => {
    let newWeight = val.currentWeight;
    let finalVelocityStr = val.currentVelocity.toFixed(2);
    let isOosToday = currentOosMap.get(val.blankType) === true;
    let numDesignsForBlank = blankDesignCount.get(val.blankType) || 1;

    // Active Age Calculation
    let calendarAgeDays = Math.floor(Math.max(0, (now - val.launchDateMs) / ONE_DAY_MS));
    let activeAgeDays = 0;

    if (calendarAgeDays <= 240) {
      activeAgeDays = activeDaysMap.get(val.blankType)[calendarAgeDays];
    } else {
      // Fallback for extremely old designs
      activeAgeDays = activeDaysMap.get(val.blankType)[240] + (calendarAgeDays - 240);
    }

    if (isOosToday) {
      // OOS Check: Freeze current weight and velocity
      newWeight = val.currentWeight;
      finalVelocityStr = val.currentVelocity.toFixed(2);
    } else {
      // --- 1. LOCAL ALLOCATION WEIGHT [x] ---
      if (numDesignsForBlank === 1) {
        newWeight = 1.0;
      } else {
        let localMaxSales = blankCeilingMap.get(val.blankType) || 1;
        let ratio = Math.min(1.0, val.salesEMA / localMaxSales);

        let dataDrivenWeight = 0.0;
        if (val.salesEMA > 0) {
          dataDrivenWeight = 0.15 + (0.85 * ratio);
        }

        const BASELINE_PRIOR = 0.6;
        let confidence = Math.min(1.0, activeAgeDays / 45.0);
        let blendedWeight = (BASELINE_PRIOR * (1.0 - confidence)) + (dataDrivenWeight * confidence);

        newWeight = Math.round(blendedWeight * 100) / 100;
        if (newWeight < 0.1) newWeight = 0.1;
        if (newWeight > 1.0) newWeight = 1.0;
      }

      // --- 2. TRUE VELOCITY |y| ---
      let denom15 = Math.max(0.1, 15.0 * (1.0 - Math.exp(-activeAgeDays / 15.0)));
      let denom75 = Math.max(0.1, 75.0 * (1.0 - Math.exp(-activeAgeDays / 75.0)));

      let v15 = val.score15 / denom15;
      let v75 = val.score75 / denom75;
      let trueVelocity = (0.6 * v15) + (0.4 * v75);

      // New Launch Assumption (Baseline runway for 0 sales)
      let assumedVelocity = 0.25 * Math.exp(-activeAgeDays / 14.0);
      let finalVelocityNum = Math.max(trueVelocity, assumedVelocity);

      finalVelocityStr = (Math.round(finalVelocityNum * 100) / 100).toFixed(2);
    }

    // Regex Swap for Column B
    let newPrintStr = val.printStr;
    // Strip out all existing brackets and pipes
    newPrintStr = newPrintStr.replace(/\s*\[.*?\]/g, '').replace(/\s*\|.*?\|/g, '').trim();
    // Append fresh values
    newPrintStr = `${newPrintStr} [${newWeight}] |${finalVelocityStr}|`;

    finalUpdates.push({ row: rowIndex + 1, val: newPrintStr });
  });

  // --- PHASE 4: Batch Writing & Reporting ---
  if (finalUpdates.length > 0) {
    let colBData = mappingSheet.getRange(1, 2, mappingData.length, 1).getValues();
    finalUpdates.forEach(update => {
      colBData[update.row - 1][0] = update.val;
    });
    mappingSheet.getRange(1, 2, mappingData.length, 1).setValues(colBData);
  }

  // Email Alert for 2% Rule (using unique SKUs)
  let missingPct = (uniqueMissingSkus.size / (totalProcessedSkus.size || 1)) * 100;
  if (missingSkus.size > 0) {
    let emailBody = "<h3>Unmapped SKUs Alert</h3>";
    emailBody += "<p>The following SKUs had sales but were not found in the Mapping Sheet.</p>";
    emailBody += "<table border='1' cellpadding='5'><tr><th>Platform | Date | Orphaned SKU</th></tr>";
    missingSkus.forEach(skuInfo => {
      emailBody += `<tr><td>${skuInfo}</td></tr>`;
    });
    emailBody += "</table>";

    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: `🚨 Action Required: ${missingSkus.size} Unmapped SKUs found (${missingPct.toFixed(1)}%)`,
      htmlBody: emailBody
    });
  }

  if (duplicateSkus.size > 0) {
    Logger.log("WARNING - Duplicate SKUs found in Mapping Sheet: " + Array.from(duplicateSkus).join(", "));
  }

  try {
    ss.toast("The Mapping Sheet has been successfully updated.", "✅ Dynamic Weights Updated");
  } catch (e) { }
}

// ==============================================================================
// MIDNIGHT SNAPSHOT LOGIC (OOS LEDGER)
// ==============================================================================

function recordMidnightOOSSnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("master_inventory");
  const ledgerSheet = ss.getSheetByName("oos_ledger");

  if (!masterSheet || !ledgerSheet) return;

  const masterData = masterSheet.getDataRange().getValues();
  const ledgerDataToRead = ledgerSheet.getDataRange().getValues();
  const now = new Date();
  const todayStr = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "dd-MMM-yyyy");

  // Build a Set of existing entries to prevent duplicate bloat if run multiple times a day
  const existingLedgerEntries = new Set();
  for (let i = 1; i < ledgerDataToRead.length; i++) {
    let rowDate = new Date(ledgerDataToRead[i][0]);
    if (!isNaN(rowDate.getTime())) {
      let dStr = Utilities.formatDate(rowDate, ss.getSpreadsheetTimeZone(), "dd-MMM-yyyy");
      let bType = String(ledgerDataToRead[i][1] || "").trim();
      existingLedgerEntries.add(dStr + "|" + bType);
    }
  }

  const newRows = [];

  // Skip headers (start at row 1)
  for (let i = 1; i < masterData.length; i++) {
    const blankType = String(masterData[i][0] || "").trim();
    // Col L is index 11
    const isOos = String(masterData[i][11] || "").trim().toLowerCase() === "yes";

    // Data Minimization: ONLY log if it's Out of Stock AND not already logged today!
    if (blankType && isOos) {
      if (!existingLedgerEntries.has(todayStr + "|" + blankType)) {
        newRows.push([todayStr, blankType]);
      }
    }
  }

  if (newRows.length > 0) {
    ledgerSheet.getRange(ledgerSheet.getLastRow() + 1, 1, newRows.length, 2).setValues(newRows);
    SpreadsheetApp.flush(); // Ensure data is physically written immediately
  }

  // Self-Cleaning: Delete rows older than 240 days
  const cutoffDate = new Date(now.getTime() - (240 * 24 * 60 * 60 * 1000));
  const ledgerData = ledgerSheet.getDataRange().getValues();

  let firstValidKeepIndex = -1;
  for (let i = 1; i < ledgerData.length; i++) {
    let rowDate = new Date(ledgerData[i][0]);
    if (!isNaN(rowDate.getTime()) && rowDate >= cutoffDate) {
      firstValidKeepIndex = i;
      break; // Found the first valid date that is recent enough to keep
    }
  }

  // If we found a valid keep date, delete everything before it
  if (firstValidKeepIndex > 1) {
    ledgerSheet.deleteRows(2, firstValidKeepIndex - 1);
  } else if (firstValidKeepIndex === -1 && ledgerData.length > 1) {
    // Entire sheet is old or invalid, delete all data rows
    ledgerSheet.deleteRows(2, ledgerData.length - 1);
  }
}


// ==============================================================================
// DYNAMIC THRESHOLDS (PRINT ON DEMAND)
// ==============================================================================
function updateDynamicThresholds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { }

  const testSheet = ss.getSheetByName("test sheet");
  const mappingSheet = ss.getSheetByName("Mapping Sheet");
  const ledgerSheet = ss.getSheetByName("oos_ledger");

  if (!testSheet || !mappingSheet || !ledgerSheet) {
    if (ui) ui.alert("Error", "Could not find required sheets (test sheet, Mapping Sheet, oos_ledger).", ui.ButtonSet.OK);
    return;
  }

  // Source Spreadsheets
  const sourceSpreadsheetId = "10OCuU7CFxuOtG2z6Q6fHWVtHbaVLg4wpZw-XeTzqq1E";
  let sourceSS;
  try {
    sourceSS = SpreadsheetApp.openById(sourceSpreadsheetId);
  } catch (e) {
    if (ui) ui.alert("System Error", "Exact Error: " + e.toString(), ui.ButtonSet.OK);
    return;
  }

  const now = Date.now();
  const ONE_DAY_MS = 24 * 60 * 60 * 1000;
  const HISTORY_WINDOW_ACTIVE_DAYS = 120;
  const LOOKBACK_HORIZON_MS = 240 * ONE_DAY_MS;

  function getLocalYMD(d) {
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, '0') + "-" + String(d.getDate()).padStart(2, '0');
  }

  // --- PHASE 1: Initialization & The Active-Day Mapper ---
  const ledgerData = ledgerSheet.getDataRange().getValues();
  const oosLog = new Map();

  for (let i = 1; i < ledgerData.length; i++) {
    const rawDate = ledgerData[i][0];
    const blankType = String(ledgerData[i][1] || "").trim();
    if (!rawDate || !blankType) continue;

    const d = new Date(rawDate);
    if (isNaN(d.getTime())) continue;
    const dateStr = getLocalYMD(d);

    if (!oosLog.has(blankType)) {
      oosLog.set(blankType, new Set());
    }
    oosLog.get(blankType).add(dateStr);
  }

  function isBlankActiveOnDate(blankType, timestampMs) {
    const d = new Date(timestampMs);
    const dateStr = getLocalYMD(d);
    if (oosLog.has(blankType) && oosLog.get(blankType).has(dateStr)) {
      return false; // It was OOS
    }
    return true; // Assume IN-STOCK if not found
  }

  // --- PHASE 2: Building the Print Master Dictionary & Festive Calendar ---
  const bundleSheet = ss.getSheetByName("Bundle_SKU_Mapping");
  const bundleData = bundleSheet ? bundleSheet.getDataRange().getValues() : [];
  const bundleMap = new Map();
  if (bundleData.length > 0) {
    for (let i = 1; i < bundleData.length; i++) {
      const row = bundleData[i];
      const rawBundleSku = String(row[0] || "").trim().toUpperCase();
      if (rawBundleSku) {
        const bundleNorm = _normalizeSku(rawBundleSku);
        const children = [];
        for (let c = 1; c < row.length; c++) {
          const child = String(row[c] || "").trim().toUpperCase();
          if (child) children.push(child);
        }
        if (children.length > 0) bundleMap.set(bundleNorm, children);
      }
    }
  }

  const mappingData = mappingSheet.getDataRange().getValues();
  const lookupMap = new Map(); // SKU -> PrintName
  const printDict = new Map(); // PrintName -> { score15: 0, score60: 0, launchDateMs: 0, blankType: '...' }
  const uniqueBlanks = new Set();

  for (let i = 1; i < mappingData.length; i++) {
    const row = mappingData[i];

    // FIX: Using Column A (clean name) which perfectly matches test sheet B
    const printName = String(row[0] || "").trim(); // Column A
    const printCol = String(row[1] || "").trim(); // Column B
    const blankType = String(row[2] || "").trim(); // Column C

    if (!printName) continue;

    uniqueBlanks.add(blankType);
    if (!printDict.has(printName)) {
      let launchDateMs = now;
      let launchDateMatch = printCol.match(/<([^>]+)>/);
      if (launchDateMatch) {
        let launchDateStr = launchDateMatch[1];
        let parsedDate = Date.parse(launchDateStr);
        if (!isNaN(parsedDate)) launchDateMs = parsedDate;
      }

      printDict.set(printName, { score15: 0, score75: 0, launchDateMs: launchDateMs, blankType: blankType });
    }

    for (let colIdx = 3; colIdx < row.length; colIdx++) {
      const rawSku = String(row[colIdx] || "").trim().toUpperCase();
      if (rawSku) {
        lookupMap.set(_normalizeSku(rawSku), printName);
      }
    }
  }

  // Precompute Active Day Indexing (O(1) Optimization)
  const activeDaysMap = new Map();
  uniqueBlanks.forEach(blank => {
    let activeDayCount = 0;
    let mappingArray = new Array(245);
    for (let d = 0; d <= 240; d++) {
      let checkTime = now - (d * ONE_DAY_MS);
      mappingArray[d] = activeDayCount;
      if (isBlankActiveOnDate(blank, checkTime)) {
        activeDayCount++;
      }
    }
    activeDaysMap.set(blank, mappingArray);
  });

  // Festive Calendar Parsing
  const salesDatesSheet = ss.getSheetByName("sales_dates");
  let festiveMultiplier = 1.0;

  if (salesDatesSheet) {
    const calendarData = salesDatesSheet.getDataRange().getValues();

    function getLocalYMD(d) {
      return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, '0') + "-" + String(d.getDate()).padStart(2, '0');
    }
    const todayStr = getLocalYMD(new Date(now));
    const todayMs = Date.parse(todayStr); // Normalized midnight today

    for (let i = 1; i < calendarData.length; i++) {
      let sDateRaw = calendarData[i][0];
      let eDateRaw = calendarData[i][1];
      let boostRaw = parseFloat(calendarData[i][2]);

      if (sDateRaw && eDateRaw && !isNaN(boostRaw)) {
        let sMs = sDateRaw instanceof Date ? sDateRaw.getTime() : Date.parse(sDateRaw);
        let eMs = eDateRaw instanceof Date ? eDateRaw.getTime() : Date.parse(eDateRaw);

        if (!isNaN(sMs) && !isNaN(eMs)) {
          // Offsets: -2 days for start, -1 day for end
          let effectiveStartMs = sMs - (2 * ONE_DAY_MS);
          let effectiveEndMs = eMs - (1 * ONE_DAY_MS);

          if (todayMs >= effectiveStartMs && todayMs <= effectiveEndMs) {
            if (boostRaw > festiveMultiplier) {
              festiveMultiplier = boostRaw;
            }
          }
        }
      }
    }
  }

  // --- PHASE 3: Sales Ingestion (Dual-Velocity EMA) ---
  function processSalesSheet(sheetName, skuColIdx, dateColIdx, isFlipkart = false) {
    const sheet = sourceSS.getSheetByName(sheetName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      let rawDate = data[i][dateColIdx];
      if (!rawDate) continue;

      let saleTime = 0;
      if (rawDate instanceof Date) {
        saleTime = rawDate.getTime();
      } else {
        let numDate = Number(rawDate);
        if (!isNaN(numDate) && numDate > 20000 && numDate < 100000) {
          saleTime = Math.round((numDate - 25569) * 86400 * 1000);
        } else {
          saleTime = Date.parse(rawDate);
        }
      }

      if (isNaN(saleTime) || now - saleTime > LOOKBACK_HORIZON_MS) continue;

      let rawSku = String(data[i][skuColIdx] || "").trim().toUpperCase();
      if (!rawSku) continue;

      rawSku = rawSku.replace(/"/g, '').trim();
      if (isFlipkart) {
        rawSku = rawSku.replace(/^SKU:\s*/i, '').trim();
      }

      let skusToProcess = [rawSku];
      let bundleNorm = _normalizeSku(rawSku);
      if (bundleMap.has(bundleNorm)) {
        skusToProcess = bundleMap.get(bundleNorm);
      }

      for (let sku of skusToProcess) {
        let normalizedSku = _normalizeSku(sku);
        if (lookupMap.has(normalizedSku)) {
          let printName = lookupMap.get(normalizedSku);
          if (printDict.has(printName)) {
            let design = printDict.get(printName);

            if (isBlankActiveOnDate(design.blankType, saleTime)) {
              let calendarDaysAgo = Math.floor(Math.max(0, (now - saleTime) / ONE_DAY_MS));

              if (calendarDaysAgo <= 240) {
                let activeDaysAgo = activeDaysMap.get(design.blankType)[calendarDaysAgo];

                if (activeDaysAgo <= HISTORY_WINDOW_ACTIVE_DAYS) {
                  // Dual-Velocity Decay
                  let recencyWeight15 = Math.exp(-activeDaysAgo / 15.0);
                  let recencyWeight75 = Math.exp(-activeDaysAgo / 75.0);
                  design.score15 += recencyWeight15;
                  design.score75 += recencyWeight75;
                }
              }
            }
          }
        }
      }
    }
  }

  processSalesSheet("WEBSITE_FORWARD", 20, 15);
  processSalesSheet("myntra_forward", 2, 0);
  processSalesSheet("AJIO_FORWARD", 4, 1);
  processSalesSheet("FLIPKART_FORWARD", 3, 0, true);

  // --- PHASE 4 & 5: Poisson Threshold Normalization ---
  const lastRow = testSheet.getLastRow();
  if (lastRow < 2) {
    if (ui) ui.alert("Notice", "test sheet is empty.", ui.ButtonSet.OK);
    return;
  }

  // Get columns B to E (B=PrintName, C=BatchSize, D=Inventory, E=Threshold)
  const targetRange = testSheet.getRange(2, 2, lastRow - 1, 4);
  const targetData = targetRange.getValues();
  const thresholdUpdates = [];

  for (let i = 0; i < targetData.length; i++) {
    const printName = String(targetData[i][0] || "").trim(); // Column B
    const currentThreshold = parseInt(targetData[i][3]); // Column E

    // Fallback if empty
    let finalThreshold = isNaN(currentThreshold) ? 1 : currentThreshold;

    if (printName && finalThreshold !== -1) {
      if (printDict.has(printName)) {
        const design = printDict.get(printName);

        let calendarAgeDays = Math.floor(Math.max(0, (now - design.launchDateMs) / ONE_DAY_MS));
        let activeAgeDays = 0;

        if (calendarAgeDays <= 240) {
          activeAgeDays = activeDaysMap.get(design.blankType)[calendarAgeDays];
        } else {
          activeAgeDays = activeDaysMap.get(design.blankType)[240] + (calendarAgeDays - 240);
        }

        // 1. Calculate Dynamic Denominators (Protects New Launches)
        let denom15 = Math.max(0.1, 15.0 * (1.0 - Math.exp(-activeAgeDays / 15.0)));
        let denom75 = Math.max(0.1, 75.0 * (1.0 - Math.exp(-activeAgeDays / 75.0)));

        // 2. Calculate Velocities
        let v15 = design.score15 / denom15;
        let v75 = design.score75 / denom75;

        // 3. Blended True Velocity
        let trueVelocity = (0.6 * v15) + (0.4 * v75);

        // 4. Volume-Aware Multiplier (1.15x Base for Good Sellers, or Festive Boost)
        let projectedVelocity = trueVelocity;
        if (trueVelocity >= 1.0) {
          let activeMultiplier = Math.max(1.15, festiveMultiplier);
          projectedVelocity = trueVelocity * activeMultiplier;
        }

        // 5. Poisson Reorder Point Formula (Demand + 97% Safety Stock)
        let rawThreshold = projectedVelocity + (2.0 * Math.sqrt(projectedVelocity));

        // 6. Hard Floor & Ceiling Rounding (Extra Safety)
        let calculatedThreshold = Math.max(1, Math.ceil(rawThreshold));

        // Cold Start Protection (7 days)
        if (calendarAgeDays <= 7) {
          calculatedThreshold = Math.max(calculatedThreshold, finalThreshold);
        }

        finalThreshold = calculatedThreshold;
      }
    }

    thresholdUpdates.push([finalThreshold]);
  }

  // Write Thresholds (Column E)
  testSheet.getRange(2, 5, thresholdUpdates.length, 1).setValues(thresholdUpdates);

  try {
    ss.toast(`Thresholds updated! Festive Multiplier: ${festiveMultiplier}x`, "✅ Success");
  } catch (e) { }
}
