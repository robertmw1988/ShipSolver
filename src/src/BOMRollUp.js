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

/**
 * Configuration used to integrate solver flights with ArtifactsByParams output.
 */
var SOLVER_BOM_CFG = {
  artifactsSheet: 'ArtifactsByParams',
  solverSheet: 'NewSolver',
  flightsColumn: 'Flights',
  keyColumns: {
    shipType: 'Ship type',
    shipDurationType: 'Ship duration type',
    shipLevel: 'Ship level',
    targetArtifact: 'Target artifact'
  },
  cacheKey: 'artifactsByParams:index:v1',
  cacheSeconds: 300
};

/**
 * Build or retrieve a cached lookup of the ArtifactsByParams sheet.
 * @param {boolean=} forceRefresh When true the cache is rebuilt.
 * @return {{artifactHeaders:Array<string>, index:Object, aliasSnapshot:Object}}
 */
function _getArtifactsByParamsIndex(forceRefresh) {
  var cache = null;
  try {
    cache = CacheService.getDocumentCache();
  } catch (e) {}

  if (!forceRefresh && cache) {
    var cached = cache.get(SOLVER_BOM_CFG.cacheKey);
    if (cached) {
      try {
        var parsed = JSON.parse(cached);
        if (parsed && parsed.index) {
          return parsed;
        }
      } catch (e) {}
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SOLVER_BOM_CFG.artifactsSheet);
  if (!sheet) {
    throw new Error('Artifacts sheet not found: ' + SOLVER_BOM_CFG.artifactsSheet);
  }

  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 4) {
    throw new Error('Artifacts sheet is empty or missing header rows.');
  }

  // Allow for "(no matches)" placeholder
  if (values.length === 1 && values[0] && values[0][0] === '(no matches)') {
    return { artifactHeaders: [], index: {}, aliasSnapshot: _buildAliasesFromSheet ? _buildAliasesFromSheet() : {} };
  }

  var headerRow = values[3];
  if (!headerRow || headerRow.length < 5) {
    throw new Error('Artifacts sheet header row is missing expected artifact columns.');
  }

  var artifactHeaders = headerRow.slice(4);
  var dataRows = values.slice(4);
  var aliases = typeof _buildAliasesFromSheet === 'function' ? _buildAliasesFromSheet() : null;
  var index = {};

  for (var i = 0; i < dataRows.length; i++) {
    var row = dataRows[i];
    if (!row || (row[0] === '' && row[1] === '' && row[2] === '' && row[3] === '')) {
      continue;
    }

    var shipRaw = row[0];
    var durationRaw = row[1];
    var levelRaw = row[2];
    var targetRaw = row[3];

    if (shipRaw === '(no matches)') {
      continue;
    }

    var key = _buildFlightKey(shipRaw, durationRaw, levelRaw, targetRaw, aliases);
    if (!key) {
      continue;
    }

    index[key] = _buildDropMapFromRow(headerRow, row);
  }

  var payload = {
    artifactHeaders: artifactHeaders,
    index: index,
    aliasSnapshot: aliases || {}
  };

  if (cache) {
    try {
      cache.put(SOLVER_BOM_CFG.cacheKey, JSON.stringify(payload), SOLVER_BOM_CFG.cacheSeconds);
    } catch (e) {}
  }

  return payload;
}

/**
 * Construct a flight key based on ship, duration, level, and target artifact.
 * Values are normalized via aliases when possible.
 * @param {*} ship
 * @param {*} duration
 * @param {*} level
 * @param {*} target
 * @param {Object} aliases
 * @return {string}
 */
function _buildFlightKey(ship, duration, level, target, aliases) {
  var canonicalShip = _standardizeKeyValue('shipType', ship, aliases);
  var canonicalDuration = _standardizeKeyValue('shipDurationType', duration, aliases);
  var canonicalTarget = _standardizeKeyValue('targetArtifact', target, aliases);
  var levelKey = level == null || level === '' ? '' : String(level);

  if (!canonicalShip || !canonicalDuration || canonicalTarget == null) {
    return '';
  }

  return [canonicalShip, canonicalDuration, levelKey, canonicalTarget].join('|');
}

/**
 * Normalize values using alias helpers when available.
 * @param {string} kind
 * @param {*} value
 * @param {Object} aliases
 * @return {string}
 */
function _standardizeKeyValue(kind, value, aliases) {
  if (value == null) {
    return '';
  }
  if (typeof _normalizeWithAliases === 'function') {
    return _normalizeWithAliases(kind, value, aliases);
  }
  if (typeof _toUpperSnake === 'function') {
    return _toUpperSnake(value);
  }
  return String(value).trim().toUpperCase();
}

/**
 * Build a sparse drop map from a row of the artifacts table.
 * Only non-zero numeric columns are retained to keep cache size small.
 * @param {Array<string>} headers
 * @param {Array<*>} row
 * @return {Object}
 */
function _buildDropMapFromRow(headers, row) {
  var out = {};
  for (var c = 4; c < headers.length; c++) {
    var name = headers[c];
    if (!name) {
      continue;
    }
    var val = row[c];
    if (val == null || val === '') {
      continue;
    }
    var num = typeof val === 'number' ? val : parseFloat(val);
    if (!isNaN(num) && num !== 0) {
      out[name] = num;
    }
  }
  return out;
}

/**
 * Retrieve the cached artifacts index, forcing a rebuild.
 */
function REFRESH_ARTIFACTS_INDEX_CACHE() {
  _getArtifactsByParamsIndex(true);
}

/**
 * Clear both the solver cache and the existing artifacts index cache.
 */
function CLEAR_ARTIFACTS_INDEX_CACHE() {
  try {
    var cache = CacheService.getDocumentCache();
    if (cache) {
      cache.remove(SOLVER_BOM_CFG.cacheKey);
    }
  } catch (e) {}
}

/**
 * Build an aggregate of expected artifact drops based on solver flights.
 * @return {Array<Array<*>>} 2D array suitable for sheet output: [[Artifact, Average Drops], ...]
 */
function GET_SOLVER_ARTIFACT_SUMMARY() {
  var artifactsData = _getArtifactsByParamsIndex(false);
  var artifactHeaders = artifactsData.artifactHeaders || [];
  var aliases = artifactsData.aliasSnapshot || (typeof _buildAliasesFromSheet === 'function' ? _buildAliasesFromSheet() : {});

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var solverSheet = ss.getSheetByName(SOLVER_BOM_CFG.solverSheet);
  if (!solverSheet) {
    throw new Error('Solver sheet not found: ' + SOLVER_BOM_CFG.solverSheet);
  }

  var solverData = solverSheet.getDataRange().getValues();
  if (!solverData || !solverData.length) {
    return [['Artifact', 'Average Drops']];
  }

  var header = solverData[0];
  var headerMap = {};
  for (var i = 0; i < header.length; i++) {
    headerMap[String(header[i]).toLowerCase()] = i;
  }

  var shipIdx = headerMap[SOLVER_BOM_CFG.keyColumns.shipType.toLowerCase()];
  var durationIdx = headerMap[SOLVER_BOM_CFG.keyColumns.shipDurationType.toLowerCase()];
  var levelIdx = headerMap[SOLVER_BOM_CFG.keyColumns.shipLevel.toLowerCase()];
  var targetIdx = headerMap[SOLVER_BOM_CFG.keyColumns.targetArtifact.toLowerCase()];
  var flightsIdx = headerMap[SOLVER_BOM_CFG.flightsColumn.toLowerCase()];

  if (shipIdx == null || durationIdx == null || levelIdx == null || targetIdx == null || flightsIdx == null) {
    throw new Error('NewSolver sheet is missing required columns: ' + JSON.stringify(SOLVER_BOM_CFG));
  }

  var totals = {};
  var missingKeys = [];

  for (var r = 1; r < solverData.length; r++) {
    var row = solverData[r];
    var flights = row[flightsIdx];
    var flightsNum = typeof flights === 'number' ? flights : parseFloat(flights);
    if (isNaN(flightsNum) || flightsNum === 0) {
      continue;
    }

    var key = _buildFlightKey(row[shipIdx], row[durationIdx], row[levelIdx], row[targetIdx], aliases);
    if (!key) {
      continue;
    }

    var dropMap = artifactsData.index[key];
    if (!dropMap) {
      missingKeys.push(key);
      continue;
    }

    for (var artifact in dropMap) {
      if (!dropMap.hasOwnProperty(artifact)) {
        continue;
      }
      var contribution = dropMap[artifact] * flightsNum;
      if (!totals.hasOwnProperty(artifact)) {
        totals[artifact] = 0;
      }
      totals[artifact] += contribution;
    }
  }

  if (missingKeys.length) {
    Logger.log('No artifact row found for solver keys: %s', JSON.stringify(missingKeys));
  }

  var output = [['Artifact', 'Average Drops']];
  for (var j = 0; j < artifactHeaders.length; j++) {
    var artifactName = artifactHeaders[j];
    if (!artifactName) {
      continue;
    }
    var total = totals.hasOwnProperty(artifactName) ? totals[artifactName] : 0;
    output.push([artifactName, total]);
  }

  return output;
}