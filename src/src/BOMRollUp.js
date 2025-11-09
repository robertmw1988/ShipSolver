/**
 * Executes a Bill of Materials (BOM) and rolls up quantities,
 * accounting for existing inventory of intermediate/raw materials (netting),
 * handling all quantities as fractional values (decimals).
 * 
 * Assumes two sheets exist in the spreadsheet:
 * 1. "BOM_Data" with columns: [Parent, Component, Quantity] (Quantity here should be integers/ratios)
 * 2. "Inventory_Stock" with columns: [Component, On Hand] (On Hand here can be fractional)
 * 
 * Can be called as a custom function in Google Sheets:
 * =ROLLUP_BOM_NETTED("1A", 1)
 * 
 * @param {string} topLevelAssembly The ID of the top-level assembly to build (e.g., "1A").
 * @param {number} desiredQuantity The desired quantity of the top-level assembly (can be fractional).
 * @return {Array<Array<string|number>>} A 2D array of rolled-up base components needed (fractional values).
 * @customfunction
 */
function ROLLUP_BOM_NETTED(topLevelAssembly, desiredQuantity) {
  const bomSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BOM_Data");
  const invSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventory_Stock");
  
  if (!bomSheet || !invSheet) {
    throw new Error("BOM_Data or Inventory_Stock sheet not found.");
  }
  
  const bomData = bomSheet.getDataRange().getValues();
  bomData.shift(); // Remove header row

  const invData = invSheet.getDataRange().getValues();
  invData.shift(); // Remove header row

  // Convert inventory list to a quick-lookup object {component: onHandQty}
  const onHandInventory = {};
  invData.forEach(row => {
    // Ensure we parse the second column (index 1) as a float
    onHandInventory[row[0]] = parseFloat(row[1]) || 0;
  });

  // Object to store final rolled-up quantities (key: component name, value: total quantity)
  const rolledUpTotals = {};
  
  // Start the recursive explosion process with netting logic
  explodeBOMNettedRecursive(topLevelAssembly, desiredQuantity, bomData, onHandInventory, rolledUpTotals);

  // Convert results object into a 2D array for Google Sheets output
  const finalOutput = [["Component", "Net Quantity Required (to order/build)"]];
  for (const component in rolledUpTotals) {
    // Only show items that actually have a net requirement > 0
    if (rolledUpTotals[component] > 0) {
        // Output the exact fractional value
        finalOutput.push([component, rolledUpTotals[component]]);
    }
  }

  return finalOutput;
}

/**
 * Recursive helper function to explore the BOM hierarchy with netting logic.
 * Handles all quantities as fractional numbers (floats).
 * 
 * @param {string} parentItem The current item being explored.
 * @param {number} grossRequirement The gross quantity needed of the parent item at this stage.
 * @param {Array<Array<string|number>>} bomData The full BOM data array.
 * @param {Object} onHandInventory The accumulator object for available inventory.
 * @param {Object} totalsAccumulator The accumulator object for total raw materials needed.
 */
function explodeBOMNettedRecursive(parentItem, grossRequirement, bomData, onHandInventory, totalsAccumulator) {
  
  // Calculate the net quantity we need to source for this specific item *at this level*
  const availableStock = onHandInventory[parentItem] || 0;
  let netRequirement = Math.max(0, grossRequirement - availableStock);

  // If we have enough stock or need 0, stop the explosion down this branch.
  if (netRequirement === 0) {
    return; 
  }
  
  // Check if the current item is a sub-assembly (has children in the BOM data)
  const children = bomData.filter(row => row[0] === parentItem);
  
  if (children.length > 0) {
    // It is an assembly; continue exploding its children
    children.forEach(childRow => {
      const component = childRow[1]; // Child item name/ID
      // The quantity in BOM data must be an integer ratio
      const quantityPerParent = parseInt(childRow[2], 10); 
      
      // The quantity needed for the child is based on the *net requirement* of the parent
      const childGrossRequirement = netRequirement * quantityPerParent;
      
      explodeBOMNettedRecursive(component, childGrossRequirement, bomData, onHandInventory, totalsAccumulator);
    });
  } else {
    // It is a raw material (no children found in BOM data)
    // Add the net requirement to the final totals accumulator
    if (totalsAccumulator[parentItem]) {
      totalsAccumulator[parentItem] += netRequirement;
    } else {
      totalsAccumulator[parentItem] = netRequirement;
    }
  }
}