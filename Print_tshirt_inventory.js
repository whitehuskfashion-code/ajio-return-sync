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
    .addItem('Update Lookups & Thresholds', 'updateInventoryLookupsAndThresholds')
    .addItem('Suggest Fabric Rolls', 'generateFabricRollSuggestions')
    .addToUi();
}

/**
 * Main function to process SKUs and update stock levels.
 */
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

    // --- 1. Validate that the required sheets exist ---
    if (!mappingSheet || !testSheet || !masterInventorySheet) {
      ui.alert("Error", "Could not find 'Mapping Sheet', 'test sheet', or 'master_inventory'. Please make sure the sheet names are correct.", ui.ButtonSet.OK);
      return;
    }

    // --- 2. Read all data into memory for performance ---
    const mappingData = mappingSheet.getRange("A2:AZ" + mappingSheet.getLastRow()).getValues();
    const testSheetLastRow = testSheet.getLastRow();

    const testDataRange = testSheet.getRange("B2:J" + testSheetLastRow);
    const testData = testDataRange.getValues();

    const skusToProcess = testSheet.getRange("I2:I" + testSheetLastRow).getValues();

    const masterInventoryLastRow = masterInventorySheet.getLastRow();
    const masterInventoryData = masterInventorySheet.getRange("A2:H" + masterInventoryLastRow).getValues();

    // --- 3. Build comprehensive SKU lookup maps ---

    // FIXED: Map ALL SKUs regardless of Column A/C status
    const skuToMappingInfo = new Map(); // SKU -> {productName, masterProductName, size, columnIndex}

    mappingData.forEach(row => {
      const productName = row[0]; // Column A (can be empty)
      const masterProductName = row[2]; // Column C (can be empty)

      // Iterate through SKU columns (D to AZ, which is index 3 to 51)
      for (let i = 3; i < row.length; i++) {
        const sku = row[i];
        if (sku) { // If SKU exists, store mapping info regardless of Column A/C status
          // Determine size based on column position
          const actualColumnIndex = i + 1;
          let size = null;

          if (actualColumnIndex >= 4) { // Start checking from column D
            const sizeIndex = (actualColumnIndex - 4) % 4;
            switch (sizeIndex) {
              case 0: size = 'S'; break;  // D, H, L, P, T, X, AB, AF...
              case 1: size = 'M'; break;  // E, I, M, Q, U, Y, AC, AG...
              case 2: size = 'L'; break;  // F, J, N, R, V, Z, AD, AH...
              case 3: size = 'XL'; break; // G, K, O, S, W, AA, AE, AI...
            }
          }

          skuToMappingInfo.set(sku, {
            productName: productName || "", // Store empty string if null/undefined
            masterProductName: masterProductName || "", // Store empty string if null/undefined
            size: size,
            columnIndex: actualColumnIndex
          });
        }
      }
    });

    // Build product lookup maps for test sheet
    const productInfoMap = new Map();
    testData.forEach((row, index) => {
      const productName = row[0]; // Column B
      const stockCount = row[2];  // Column D
      if (productName) {
        if (!productInfoMap.has(productName)) {
          productInfoMap.set(productName, []);
        }
        productInfoMap.get(productName).push({
          rowIndex: index,
          stock: stockCount
        });
      }
    });

    // Build master product lookup maps
    const masterProductInfoMap = new Map();
    masterInventoryData.forEach((row, index) => {
      const masterProductName = row[0]; // Column A
      if (masterProductName) {
        if (!masterProductInfoMap.has(masterProductName)) {
          masterProductInfoMap.set(masterProductName, []);
        }
        masterProductInfoMap.get(masterProductName).push({
          rowIndex: index,
          stocks: {
            S: row[1],  // Column B
            M: row[3],  // Column D  
            L: row[5],  // Column F
            XL: row[7]  // Column H
          }
        });
      }
    });

    // --- 4. Process each SKU and prepare the results ---
    const results = [];
    const statusColors = [];
    const stockUpdates = testData.map(row => [row[2]]);
    const masterInventoryDecrements = new Map();

    skusToProcess.forEach((row, index) => {
      const sku = row[0];
      let status = "";
      let highlightColor = "#000000"; // Default black

      if (sku) {
        let originalOpStatus = "";
        let masterOpStatus = "";

        // Check if SKU exists in mapping
        const mappingInfo = skuToMappingInfo.get(sku);

        if (mappingInfo) {
          const productName = mappingInfo.productName;
          const masterProductName = mappingInfo.masterProductName;

          // --- ORIGINAL OPERATION Logic ---
          if (productName) {
            // Column A has value - process normally
            const productEntries = productInfoMap.get(productName);
            if (productEntries && productEntries.length > 0) {
              let decrementedCount = 0;
              const warnings = [];

              productEntries.forEach(entry => {
                const stock = stockUpdates[entry.rowIndex][0];
                if (typeof stock === 'number') {
                  stockUpdates[entry.rowIndex][0] = stock - 1;
                  decrementedCount++;
                } else {
                  warnings.push(`${productName} (row ${entry.rowIndex + 2}): typeof stock != 'number'`);
                }
              });

              if (decrementedCount > 0) {
                originalOpStatus = `Print: ✅ (${decrementedCount} decremented)`;
                if (warnings.length) {
                  originalOpStatus += `; Warnings: ${warnings.join('; ')}`;
                }
              } else {
                originalOpStatus = `Print: ❌ (${warnings.join('; ')})`;
                highlightColor = "#FF0000"; // Red
              }
            } else {
              originalOpStatus = `Print: ❌ (${productName}: not in test sheet)`;
              highlightColor = "#FF0000"; // Red
            }
          } else {
            // Column A is empty - no print linked
            originalOpStatus = "Print: No Print linked❌";
          }

          // --- MASTER OPERATION Logic ---
          if (masterProductName) {
            // Column C has value - process master inventory
            const size = mappingInfo.size;
            if (!size) {
              masterOpStatus = `Plain/Ready Merch: ❌ (Size not determined for column ${mappingInfo.columnIndex})`;
              highlightColor = "#FF0000"; // Red
            } else {
              const masterEntries = masterProductInfoMap.get(masterProductName);
              if (masterEntries && masterEntries.length > 0) {
                // Track this decrement
                if (!masterInventoryDecrements.has(masterProductName)) {
                  masterInventoryDecrements.set(masterProductName, { S: 0, M: 0, L: 0, XL: 0 });
                }
                masterInventoryDecrements.get(masterProductName)[size]++;

                masterOpStatus = "Plain/Ready Merch: ✅ (1 decremented)";
              } else {
                masterOpStatus = `Plain/Ready Merch: ❌ (${masterProductName}: not in master_inventory)`;
                highlightColor = "#FF0000"; // Red
              }
            }
          } else {
            // Column C is empty
            if (productName) {
              // Column A has value but Column C is empty
              masterOpStatus = "Plain/Ready Merch: Product: Please link with master_inventory Column C";
              highlightColor = "#007BFF"; // Blue
            } else {
              // Both Column A and C are empty
              masterOpStatus = "Plain/Ready Merch: ❌ (Both columns A and C are empty)";
              highlightColor = "#FF0000"; // Red
            }
          }

          // Determine final highlight color based on scenarios
          if (!productName && masterProductName) {
            // Scenario 1: Column A empty + Column C has value
            highlightColor = "#800080"; // Purple
          } else if (productName && masterProductName) {
            // Scenario 2: Both columns have values - check if both operations succeeded
            if (originalOpStatus.includes("✅") && masterOpStatus.includes("✅")) {
              highlightColor = "#000000"; // Black (no highlighting)
            } else {
              highlightColor = "#FF0000"; // Red for any failures
            }
          } else if (productName && !masterProductName) {
            // Scenario 3: Column A has value + Column C empty
            highlightColor = "#FF0000"; // Red
          }

        } else {
          // SKU not found in mapping sheet
          originalOpStatus = "Print: ❌ (SKU not in Mapping Sheet)";
          masterOpStatus = "Plain/Ready Merch: ❌ (SKU not in Mapping Sheet)";
          highlightColor = "#FF0000"; // Red
        }

        // Combine both operation statuses
        status = originalOpStatus + " | " + masterOpStatus;
      }
      // else {
      //   status = "No SKU provided";
      //   highlightColor = "#FF0000"; // Red
      // }

      results.push([status]);
      statusColors.push([highlightColor]);
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
      // Write all status results to column J
      const statusRange = testSheet.getRange(2, 10, results.length, 1);
      statusRange.setValues(results);

      // Apply color highlighting
      statusColors.forEach((colorRow, index) => {
        statusRange.getCell(index + 1, 1).setFontColor(colorRow[0]);
      });

      // Write all updated stock counts to column D in test sheet
      testSheet.getRange(2, 4, stockUpdates.length, 1).setValues(stockUpdates);

      // Write all updated master inventory stocks
      if (masterStockUpdates.length > 0) {
        masterInventorySheet.getRange(2, 2, masterStockUpdates.length, 1).setValues(masterStockUpdates.map(row => [row[0]])); // Column B
        masterInventorySheet.getRange(2, 4, masterStockUpdates.length, 1).setValues(masterStockUpdates.map(row => [row[1]])); // Column D
        masterInventorySheet.getRange(2, 6, masterStockUpdates.length, 1).setValues(masterStockUpdates.map(row => [row[2]])); // Column F
        masterInventorySheet.getRange(2, 8, masterStockUpdates.length, 1).setValues(masterStockUpdates.map(row => [row[3]])); // Column H
      }
    }

    // --- Store the timestamp of this successful run ---
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
  const mappingData = mappingSheet.getRange("C2:C" + mappingLastRow).getValues();

  const inventoryLastRow = inventorySheet.getLastRow();
  if (inventoryLastRow < 2) return; // Nothing to process
  const inventoryData = inventorySheet.getRange("A2:C" + inventoryLastRow).getValues();

  const tierRulesLastRow = Math.max(2, tierRulesSheet.getLastRow());
  const tierRulesData = tierRulesSheet.getRange("A2:O" + tierRulesLastRow).getValues();

  // --- 2. Calculate Print Counts from Mapping Sheet ---
  const mappingCounts = new Map();
  mappingData.forEach(row => {
    let name = row[0];
    if (name) {
      name = String(name).trim(); // Trim spaces for consistency
      mappingCounts.set(name, (mappingCounts.get(name) || 0) + 1);
    }
  });

  // --- 3. Parse Tier Rules ---
  const tierRulesMap = new Map();
  tierRulesData.forEach(row => {
    let category = row[0];
    if (category) {
      category = String(category).trim();
      let minDesigns = Number(row[1]) || 0;
      let thresholds = row.slice(3, 15); // Columns D through O (12 values)

      if (!tierRulesMap.has(category)) {
        tierRulesMap.set(category, []);
      }
      tierRulesMap.get(category).push({ minDesigns: minDesigns, thresholds: thresholds });
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
  const emptyThresholds = Array(12).fill("");

  inventoryData.forEach(row => {
    let colorName = String(row[0] || "").trim(); // Column A
    let category = String(row[2] || "").trim();  // Column C

    // Get print count
    let printCount = mappingCounts.get(colorName) || 0;
    printCountsToUpdate.push([printCount]);

    // Get thresholds
    let rowThresholds = emptyThresholds;

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
          break;
        }
      }
    }
    thresholdsToUpdate.push(rowThresholds);
  });

  // --- 5. Write Data Back ---
  // Write Print Counts to Column B
  inventorySheet.getRange(2, 2, printCountsToUpdate.length, 1).setValues(printCountsToUpdate);
  // Write Thresholds to Columns F through Q (12 columns)
  inventorySheet.getRange(2, 6, thresholdsToUpdate.length, 12).setValues(thresholdsToUpdate);

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

  masterData.forEach(row => {
    let product = String(row[0] || "").trim();
    let thresholds = thresholdMap.get(product);

    if (thresholds) {
      bgS.push([getColor(row[1], row[2], thresholds.s_oos, thresholds.s_new)]);
      bgM.push([getColor(row[3], row[4], thresholds.m_oos, thresholds.m_new)]);
      bgL.push([getColor(row[5], row[6], thresholds.l_oos, thresholds.l_new)]);
      bgXL.push([getColor(row[7], row[8], thresholds.xl_oos, thresholds.xl_new)]);
    } else {
      bgS.push([null]);
      bgM.push([null]);
      bgL.push([null]);
      bgXL.push([null]);
    }
  });

  // --- 3. Apply Backgrounds ---
  if (bgS.length > 0) {
    masterSheet.getRange(2, 2, bgS.length, 1).setBackgrounds(bgS); // S (Col B)
    masterSheet.getRange(2, 4, bgM.length, 1).setBackgrounds(bgM); // M (Col D)
    masterSheet.getRange(2, 6, bgL.length, 1).setBackgrounds(bgL); // L (Col F)
    masterSheet.getRange(2, 8, bgXL.length, 1).setBackgrounds(bgXL); // XL (Col H)
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
