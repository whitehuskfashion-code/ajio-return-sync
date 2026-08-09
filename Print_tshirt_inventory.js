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

      // Iterate through SKU columns (B to AZ, which is index 1 to 51)
      for (let i = 1; i < row.length; i++) {
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