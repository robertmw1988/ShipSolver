/*******************************************************
 * Egg Inc. Virtue - Mission Retrieval & Solver Prep (ES5)
 * No const/let, no =>, no template literals, no ||= ?? ?.
 *******************************************************/

/* =========================
   CONFIGURATION
   ========================= */

var CFG = {
  // Sheet names
  dataSheet: 'AllArtifactData',          // Source data, used by _buildMissionIndex and _transformALLArtifactData
  sheetRawMissionDataTable: 'MissionDataRaw',
  sheetIndexedMissions: 'MissionDataIndexed',
  paramsSheet: 'Ship_Parameters',         // Parameters like ship levels

  aliasesSheet: 'Aliases',
  resultsSheet: 'ArtifactsByParams',

  dataCols: {
    shipType: 'shipType',
    shipDurationType: 'shipDurationType',
    targetArtifact: 'targetArtifact',
    missionLevel: 'level'
  },

  // Map to the column names found on Ship_Parameters
  paramsCols: {
    shipType: 'Ship type',          // <- header in Ship_Parameters
    shipDurationType: 'Ship duration type',  // <- header in Ship_Parameters (Short/Standard/Extended/Tutorial)
    level: 'Ship level'       // <- user-configured star level per ship & type
  },

  aliasCols: { kind: 'kind', alias: 'alias', canonical: 'canonical' },
  defaultLevelIfMissing: 0,
  includeHeaderRow: true,
  includeUnknowns: true,
  cacheSeconds: 0
};

/* =========================
   CANONICAL KEY ORDER
   ========================= */


const Keys = Object.freeze({
  shipType: Object.freeze([
    'CHICKEN_ONE',
    'CHICKEN_NINE',
    'ATREGGIES',
    'CHICKEN_HEAVY',
    'BCR',
    'MILLENIUM_CHICKEN',
    'CORELLIHEN_CORVETTE',
    'GALEGGTICA',
    'CHICKFIANT',
    'VOYEGGER',
    'HENERPRISE',
  ]),

  shipDurationType: Object.freeze([
    'TUTORIAL',
    'SHORT',
    'LONG',
    'EPIC',
  ]),

  targetArtifact: Object.freeze([
    'UNKNOWN',
    'GOLD_METEORITE',
    'TAU_CETI_GEODE',
    'SOLAR_TITANIUM',
    'LUNAR_TOTEM',
    'DEMETERS_NECKLACE',
    'TUNGSTEN_ANKH',  
    'PUZZLE_CUBE',
    'INTERSTELLAR_COMPASS',
    'QUANTUM_METRONOME',
    'MERCURYS_LENS',
    'ORNATE_GUSSET',
    'THE_CHALICE',
    'BOOK_OF_BASAN',
    'PHOENIX_FEATHER',
    'VIAL_MARTIAN_DUST',
    'AURELIAN_BROOCH',
    'CARVED_RAINSTICK',
    'BEAK_OF_MIDAS',
    'SHIP_IN_A_BOTTLE',
    'TACHYON_DEFLECTOR',
    'DILITHIUM_MONOCLE',
    'TITANIUM_ACTUATOR',
    'NEODYMIUM_MEDALLION',
    'LIGHT_OF_EGGENDIL',
    'LUNAR_STONE_FRAGMENT',
    'LUNAR_STONE',
    'QUANTUM_STONE_FRAGMENT',
    'QUANTUM_STONE',
    'TACHYON_STONE_FRAGMENT',
    'TACHYON_STONE',
    'SOUL_STONE_FRAGMENT',
    'SOUL_STONE',
    'DILITHIUM_STONE_FRAGMENT',
    'DILITHIUM_STONE',
    'SHELL_STONE_FRAGMENT',
    'SHELL_STONE',
    'TERRA_STONE_FRAGMENT',
    'TERRA_STONE',
    'LIFE_STONE_FRAGMENT',
    'LIFE_STONE',
    'PROPHECY_STONE_FRAGMENT',
    'PROPHECY_STONE',
    'CLARITY_STONE_FRAGMENT',
    'CLARITY_STONE',
  ]),
});

/* =========================
    Undroppable Artifacts List
    These are artifacts that cannot be dropped by any ship, even if they appear in the data.
    They are manually added to ensure they are included in the final output.
   ========================= */
const UndroppableArtifacts = Object.freeze([
  "BOOK_OF_BASAN | 3 | COMMON",
  "BOOK_OF_BASAN | 3 | EPIC",
  "BOOK_OF_BASAN | 3 | LEGENDARY",
  "TACHYON_DEFLECTOR | 3 | COMMON",
  "TACHYON_DEFLECTOR | 3 | RARE",
  "TACHYON_DEFLECTOR | 3 | EPIC",
  "TACHYON_DEFLECTOR | 3 | LEGENDARY",
  "CLARITY_STONE | 2 | COMMON",
  "DILITHIUM_STONE | 2 | COMMON",
  "PROPHECY_STONE | 2 | COMMON",
]);

/* =========================
   UTILITIES
   ========================= */
function _toUpperSnake(s) {
  if (s == null) return '';
  return String(s).trim().replace(/[\s\-]+/g, '_').toUpperCase();
}

function _getsheetbyname(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

function _transformALLArtifactData() {
  var sheet = _getsheetbyname(CFG.dataSheet);
  var data = sheet.getDataRange().getValues();
  
  var headers = data[0];

  var rows = data.slice(1);


  // Load ship filters from 'Ship_Parameters' sheet
  var paramSheet = _getsheetbyname(CFG.paramsSheet);
  var paramData = paramSheet.getDataRange().getValues();
  var paramHeaders = paramData[0];
  var shipTypeIndex = paramHeaders.indexOf(CFG.paramsCols.shipType);
  var shipLevelIndex = paramHeaders.indexOf(CFG.paramsCols.level);
  var shipFilters = [];
  for (var i = 1; i < paramData.length; i++) {
    shipFilters.push([paramData[i][shipTypeIndex], paramData[i][shipLevelIndex]]);
  }

  // Filter rows based on shipFilters
  var filteredRows = [];
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    for (var f = 0; f < shipFilters.length; f++) {
      if (row[headers.indexOf("Ship type")] === shipFilters[f][0] &&
          row[headers.indexOf("Ship level")] === shipFilters[f][1]) {
        filteredRows.push(row);
        break;
      }
    }
  }

  var keyCols = ["Ship type", "Ship duration type", "Ship level", "Target artifact"];
  var valueCols = ["Artifact type", "Artifact tier", "Artifact rarity"];
  var totalDropsCol = "Total drops";

  var keyIndexes = [];
  var valueIndexes = [];
  var totalIndex = headers.indexOf(totalDropsCol);

  for (var i = 0; i < keyCols.length; i++) {
    keyIndexes.push(headers.indexOf(keyCols[i]));
  }
  for (var j = 0; j < valueCols.length; j++) {
    valueIndexes.push(headers.indexOf(valueCols[j]));
  }

  var pivot = {};
  var allValueKeys = {};

  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    var key = [];
    for (var k = 0; k < keyIndexes.length; k++) {
      key.push(row[keyIndexes[k]]);
    }
    var keyStr = key.join(" | ");

    var valueKey = [];
    for (var v = 0; v < valueIndexes.length; v++) {
      valueKey.push(row[valueIndexes[v]]);
    }
    var valueStr = valueKey.join(" | ");

    var drops = Number(row[totalIndex]) || 0;

    if (!pivot[keyStr]) {
      pivot[keyStr] = {};
    }
    if (!pivot[keyStr][valueStr]) {
      pivot[keyStr][valueStr] = 0;
    }
    pivot[keyStr][valueStr] += drops;

    allValueKeys[valueStr] = true;
  }

  // Undroppable artifacts manual addition
  for (var m = 0; m < UndroppableArtifacts.length; m++) {
    allValueKeys[UndroppableArtifacts[m]] = true;
  }

  var outputSheet = _getsheetbyname(CFG.sheetRawMissionDataTable);
  var valueKeysList = [];
  for (var vk in allValueKeys) {
    if (allValueKeys.hasOwnProperty(vk)) valueKeysList.push(vk);
  }

  // Order value columns by canonical artifact order from Keys.targetArtifact
  valueKeysList = _orderValueKeysByArtifact(valueKeysList);

  var headerRow = keyCols.concat(valueKeysList);
  var outRows = [headerRow];

  for (var pk in pivot) {
    if (!pivot.hasOwnProperty(pk)) continue;
    var keyParts = pk.split(" | ");
    var rowOut = keyParts.slice();
    for (var i = 0; i < valueKeysList.length; i++) {
      var val = pivot[pk][valueKeysList[i]] || 0;
      rowOut.push(val);
    }
    outRows.push(rowOut);
  }

  // Write all at once
  outputSheet.clearContents();
  if (outRows.length) outputSheet.getRange(1, 1, outRows.length, outRows[0].length).setValues(outRows);
}


// Reads a sheet and returns { header: { colNameLower: index }, rows: [...], all: [...] }
function _readSheet(sheetName) {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet not found: ' + sheetName);
  var lr = sh.getLastRow();
  var lc = sh.getLastColumn();
  if (lr < 1 || lc < 1) return { header: {}, rows: [], all: [] };
  var all = sh.getRange(1, 1, lr, lc).getValues();
  if (!all.length) return { header: {}, rows: [], all: [] };
  var hdrArr = all[0];
  var header = {};
  for (var i = 0; i < hdrArr.length; i++) {
    var h = String(hdrArr[i] == null ? '' : hdrArr[i]).trim();
    header[h.toLowerCase()] = i;
  }
  return { header: header, rows: all.slice(1), all: all };
}

// Cache helpers (document cache)
function _cacheGet(key) {
  try {
    var cache = CacheService.getDocumentCache();
    var v = cache.get(key);
    return v ? JSON.parse(v) : null;
  } catch (e) {
    return null;
  }
}
function _cachePut(key, obj, seconds) {
  try {
    var cache = CacheService.getDocumentCache();
    cache.put(key, JSON.stringify(obj), seconds);
  } catch (e) {}
}

/* =========================
   ALIASES BUILDING
   ========================= */
function _buildAliasesFromSheet() {
  var ck = 'aliases:' + CFG.aliasesSheet;
  var cached = _cacheGet(ck);
  if (cached) return cached;

  var map = { shipType: {}, shipDurationType: {}, targetArtifact: {} };
  var sh = SpreadsheetApp.getActive().getSheetByName(CFG.aliasesSheet);
  if (!sh) { _cachePut(ck, map, CFG.cacheSeconds); return map; }

  var read = _readSheet(CFG.aliasesSheet);
  var header = read.header, rows = read.rows;
  var cK = header[CFG.aliasCols.kind.toLowerCase()];
  var cA = header[CFG.aliasCols.alias.toLowerCase()];
  var cC = header[CFG.aliasCols.canonical.toLowerCase()];

  if (cK == null || cA == null || cC == null) {
    _cachePut(ck, map, CFG.cacheSeconds);
    return map;
  }

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var kind  = String((r[cK] || '')).trim();
    var alias = String((r[cA] || '')).trim();
    var canon = String((r[cC] || '')).trim();
    if (!kind || !alias || !canon) continue;

    if (!map[kind]) map[kind] = {};
    map[kind][alias] = canon;
    map[kind][_toUpperSnake(alias)] = canon; // convenience
  }

  _cachePut(ck, map, CFG.cacheSeconds);
  return map;
}

//TODO: check if this works with current _transformALLArtifactData usage
function _normalizeWithAliases(kind, value, aliases) {
  var raw = String(value == null ? '' : value).trim();
  if (!raw) return '';

  // 1) Exact alias match
  if (aliases && aliases[kind] && aliases[kind][raw]) return aliases[kind][raw];

  // 2) Upper_snake variant alias match
  var up = _toUpperSnake(raw);
  if (aliases && aliases[kind] && aliases[kind][up]) return aliases[kind][up];

  // 3) Heuristic fallback for ships (handle common “display” names)
  if (kind === 'shipType') {
    // Map well-known variants/prefixes to canonical keys
    // Feel free to add lines here if you discover new patterns.
    if (/^ATREGGIES(?:[_\s-]?HENLINER)?$/i.test(raw)) return 'ATREGGIES';
    if (/^CORNISH[\s-]?HEN[\s-]?CORVETTE$/i.test(raw)) return 'CORELLIHEN_CORVETTE';
    if (/^DEFIHENT$/i.test(raw)) return 'CHICKFIANT';      // typo -> canonical
    if (/^VOYEGGER$/i.test(raw)) return 'VOYEGGER';
    if (/^HENERPRISE$/i.test(raw)) return 'HENERPRISE';
    if (/^CHICKEN[\s-]?ONE$/i.test(raw)) return 'CHICKEN_ONE';
    if (/^CHICKEN[\s-]?NINE$/i.test(raw)) return 'CHICKEN_NINE';
    if (/^CHICKEN[\s-]?HEAVY$/i.test(raw)) return 'CHICKEN_HEAVY';
    if (/^BCR$/i.test(raw)) return 'BCR';
    if (/^MILLEN(IA|IU)M[\s-]?CHICKEN$/i.test(raw)) return 'MILLENIUM_CHICKEN';
    if (/^GALEGG(T|TT)ICA$/i.test(raw)) return 'GALEGGTICA';
  }

  // 4) Heuristic fallback for durations (readable -> canonical)
  if (kind === 'shipDurationType') {
    if (/^SHORT$/i.test(raw)) return 'SHORT';
    if (/^STANDARD$/i.test(raw)) return 'STANDARD';
    if (/^EXTENDED$/i.test(raw)) return 'EXTENDED';
    if (/^TUTORIAL$/i.test(raw)) return 'TUTORIAL';
  }

  // 5) Default: upper_snake
  return up;
}


/* =========================
   FLEXIBLE COLUMN RESOLUTION
   ========================= */
var SYNONYMS = {
  shipType:        ['Ship type', 'Shiptype', 'ship','ship_name'],
  shipDurationType:['Ship duration type','durationtype','type','missiontype','duration'],
  targetArtifact:  ['Target artifact','target'],
  missionLevel:    ['Mission level','level','lvl']
};

function _resolveCols(header, desired) {
  var out = {};
  for (var logical in desired) {
    if (!desired.hasOwnProperty(logical)) continue;
    var prefer = String(desired[logical] || '').toLowerCase();
    var idx = header[prefer];
    if (idx == null) {
      var alts = SYNONYMS[logical] || [];
      for (var i = 0; i < alts.length; i++) {
        var alt = alts[i];
        var altKey = String(alt).toLowerCase();
        if (header.hasOwnProperty(altKey)) { idx = header[altKey]; break; }
      }
    }
    if (idx == null) {
      throw new Error("Column '" + desired[logical] + "' for " + logical + " not found and no synonyms matched.");
    }
    out[logical] = idx;
  }
  return out;
}

/* =========================
   PARAM MAP: (stype␞sdur) -> level (object dict)
   ========================= */

//TODO: check if this works with current _transformALLArtifactData usage
function _buildParamLevelMap() {
  var ck = 'paramMap:' + CFG.paramsSheet;
  var cached = _cacheGet(ck);
  if (cached) {
    var obj = {};
    for (var i = 0; i < cached.length; i++) obj[cached[i][0]] = cached[i][1];
    return obj;
  }

  var read = _readSheet(CFG.paramsSheet);
  var header = read.header, rows = read.rows;

  var cSt = header[String(CFG.paramsCols.shipType).toLowerCase()];
  var cDu = header[String(CFG.paramsCols.shipDurationType).toLowerCase()];
  var cLv = header[String(CFG.paramsCols.level).toLowerCase()];
  if (cSt == null || cDu == null || cLv == null) {
    throw new Error(
      "Params '" + CFG.paramsSheet + "' must have headers: " +
      CFG.paramsCols.shipType + ", " + CFG.paramsCols.shipDurationType + ", " + CFG.paramsCols.level
    );
  }

  var aliases = _buildAliasesFromSheet();
  var map = {};
  var serial = [];

  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    // normalize Ship / Type with your existing Aliases
    var st = _normalizeWithAliases('shipType',        row[cSt], aliases);
    var du = _normalizeWithAliases('shipDurationType', row[cDu], aliases);
    var lv = row[cLv]; // <-- USE StarsHelper AS-IS (e.g., 0..8)
    if (!st || !du) continue;

    var key = st + '␞' + du;
    map[key] = lv;
    serial.push([key, lv]);
  }

  _cachePut(ck, serial, CFG.cacheSeconds);
  return map;
}



/* =========================
   BUILD DATA INDEX (AllArtifactData)
   stype -> sdur -> level -> artifact -> [rowIndex]
   ========================= */

   //TODO: check if this works with current _transformALLArtifactData usage
function _buildMissionIndex() {
  var ck = 'index:' + CFG.dataSheet;
  var cached = _cacheGet(ck);
  if (cached) return cached;

  var read = _readSheet(CFG.dataSheet);
  var header = read.header, rows = read.rows, all = read.all;

  var cols = _resolveCols(header, CFG.dataCols);
  var aliases = _buildAliasesFromSheet();

  var index = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var st = _normalizeWithAliases('shipType',        r[cols.shipType], aliases);
    var du = _normalizeWithAliases('shipDurationType', r[cols.shipDurationType], aliases);
    var ta = _normalizeWithAliases('targetArtifact',   r[cols.targetArtifact], aliases);
    var lv = r[cols.missionLevel];

    if (!st || !du || ta == null || lv == null || lv === '') continue;

    var lvKey = String(lv);
    if (!index[st]) index[st] = {};
    if (!index[st][du]) index[st][du] = {};
    if (!index[st][du][lvKey]) index[st][du][lvKey] = {};
    if (!index[st][du][lvKey][ta]) index[st][du][lvKey][ta] = [];
    index[st][du][lvKey][ta].push(i + 2); // store 1-based row index (header offset)
  }

  var packed = { index: index, header: header, all: all, cols: cols };
  _cachePut(ck, packed, CFG.cacheSeconds);
  return packed;
}

/* =========================
   CANONICAL ORDER HELPERS
   ========================= */

function _orderValueKeysByArtifact(valueKeys) {
  // Create map of artifact name to canonical index
  var canonicalIndex = {};
  for (var i = 0; i < Keys.targetArtifact.length; i++) {
    canonicalIndex[Keys.targetArtifact[i]] = i;
  }

  // Sort by extracting artifact name from "ARTIFACT | TIER | RARITY" format
  // and using canonical index or fallback to alphabetical
  return valueKeys.sort(function(a, b) {
    var aArtifact = a.split(" | ")[0];
    var bArtifact = b.split(" | ")[0];
    var aIdx = canonicalIndex[aArtifact];
    var bIdx = canonicalIndex[bArtifact];
    
    // If both have canonical positions, use those
    if (aIdx !== undefined && bIdx !== undefined) {
      return aIdx - bIdx;
    }
    // If only one has position, it goes first
    if (aIdx !== undefined) return -1;
    if (bIdx !== undefined) return 1;
    // Otherwise alphabetical
    return a.localeCompare(b);
  });
}
function _appendUnknowns(canonicalList, actualKeys) {
  var actualSet = {};
  for (var i = 0; i < actualKeys.length; i++) actualSet[actualKeys[i]] = true;

  var seen = {};
  var out = [];
  // canonical first
  for (var c = 0; c < canonicalList.length; c++) {
    var k = canonicalList[c];
    if (actualSet[k]) { out.push(k); seen[k] = true; }
  }
  // then unknowns in actual order
  for (var j = 0; j < actualKeys.length; j++) {
    var a = actualKeys[j];
    if (!seen[a]) out.push(a);
  }
  return out;
}

/* =========================
   SPILL FUNCTION (no level sort)
   ========================= */
/**
 * Returns full rows from AllArtifactData where:
 *   missionLevel == Params level for that (shipType, shipDurationType).
 * Output order: shipType -> shipDurationType -> targetArtifact (canonical).
 * Usage: =GET_MISSION_ROWS_BY_PARAM()
 */

//TODO: check if this works with current _transformALLArtifactData usage
function GET_MISSION_ROWS_BY_PARAM() {
  var built = _buildMissionIndex();
  var index = built.index, all = built.all;
  var paramMap = _buildParamLevelMap();

  var out = [];
  if (CFG.includeHeaderRow) out.push(all[0]);

  var stActual = Object.keys(index);
  var stKeys;
  if (CFG.includeUnknowns) stKeys = _appendUnknowns(Keys.shipType, stActual);
  else {
    stKeys = [];
    for (var si = 0; si < Keys.shipType.length; si++) {
      var sk = Keys.shipType[si];
      if (stActual.indexOf(sk) !== -1) stKeys.push(sk);
    }
  }

  for (var s = 0; s < stKeys.length; s++) {
    var st = stKeys[s];
    var duMap = index[st];
    if (!duMap) continue;

    var duActual = Object.keys(duMap);
    var duKeys;
    if (CFG.includeUnknowns) duKeys = _appendUnknowns(Keys.shipDurationType, duActual);
    else {
      duKeys = [];
      for (var di = 0; di < Keys.shipDurationType.length; di++) {
        var dk = Keys.shipDurationType[di];
        if (duActual.indexOf(dk) !== -1) duKeys.push(dk);
      }
    }

    for (var d = 0; d < duKeys.length; d++) {
      var du = duKeys[d];

      var pLvRaw = paramMap[st + '␞' + du];
      var pLv = (pLvRaw == null || pLvRaw === '') ? CFG.defaultLevelIfMissing : pLvRaw;

      var lvMap = duMap[String(pLv)];
      if (!lvMap) continue;

      var taActual = Object.keys(lvMap);
      var taKeys;
      if (CFG.includeUnknowns) taKeys = _appendUnknowns(Keys.targetArtifact, taActual);
      else {
        taKeys = [];
        for (var ti = 0; ti < Keys.targetArtifact.length; ti++) {
          var tg = Keys.targetArtifact[ti];
          if (taActual.indexOf(tg) !== -1) taKeys.push(tg);
        }
      }

      for (var t = 0; t < taKeys.length; t++) {
        var ta = taKeys[t];
        var rowIdxs = lvMap[ta];
        if (!rowIdxs || !rowIdxs.length) continue;
        for (var r = 0; r < rowIdxs.length; r++) {
          var ri = rowIdxs[r];
          out.push(all[ri - 1]);
        }
      }
    }
  }

  if (out.length === (CFG.includeHeaderRow ? 1 : 0)) return [['(no matches)']];
  return out;
}

/**
 * Optional projection:
 *   =GET_MISSION_ROWS_BY_PARAM_PROJECTED(4)
 *   =GET_MISSION_ROWS_BY_PARAM_PROJECTED("shipType","shipDurationType","missionLevel","targetArtifact")
 */

//TODO: check if this works with current _transformALLArtifactData usage
function GET_MISSION_ROWS_BY_PARAM_PROJECTED() {
  var args = Array.prototype.slice.call(arguments);
  var rows = GET_MISSION_ROWS_BY_PARAM();
  if (!rows || !rows.length) return rows;

  if (args.length === 1 && typeof args[0] === 'number') {
    var n = Math.max(1, args[0]);
    var projA = [];
    for (var i = 0; i < rows.length; i++) projA.push(rows[i].slice(0, n));
    return projA;
  }

  if (args.length > 0 && typeof args[0] === 'string') {
    var header = rows[0];
    var map = {};
    for (var h = 0; h < header.length; h++) map[String(header[h]).toLowerCase()] = h;

    var idxs = [];
    for (var a = 0; a < args.length; a++) {
      var name = String(args[a]).toLowerCase();
      if (map[name] == null) throw new Error('Column not found in output: ' + args[a]);
      idxs.push(map[name]);
    }
    var proj = [];
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r], outRow = [];
      for (var k = 0; k < idxs.length; k++) outRow.push(row[idxs[k]]);
      proj.push(outRow);
    }
    return proj;
  }

  return rows;
}

/* =========================
   WRITERS FOR SOLVER
   ========================= */

function WRITE_RESULTS_SHEET() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CFG.resultsSheet) || ss.insertSheet(CFG.resultsSheet);
  var data = GET_MISSION_ROWS_BY_PARAM();
  sh.clearContents();
  if (!data || !data.length) {
    sh.getRange(1,1,1,1).setValue('(no matches)');
    return;
  }
  sh.getRange(1,1,data.length,data[0].length).setValues(data);
}

function WRITE_RESULTS_KEYS_ONLY() {
  var full = GET_MISSION_ROWS_BY_PARAM();
  if (!full || !full.length) return;

  var header = full[0];
  var map = {};
  for (var i = 0; i < header.length; i++) map[String(header[i]).toLowerCase()] = i;

  // Use consistent column names from CFG
  var colsToExtract = [
    CFG.paramsCols.shipType.toLowerCase(),
    CFG.paramsCols.shipDurationType.toLowerCase(),
    CFG.paramsCols.level.toLowerCase(),
    'target artifact' // From keyCols in _transformALLArtifactData
  ];
  
  // Verify all required columns exist
  for (var n = 0; n < colsToExtract.length; n++) {
    if (map[colsToExtract[n]] == null) {
      throw new Error('Missing column in results: ' + colsToExtract[n]);
    }
  }

  var slim = [['shipType','shipDurationType','level','targetArtifact']];
  for (var r = 1; r < full.length; r++) {
    var row = full[r];
    slim.push([
      row[map[colsToExtract[0]]],
      row[map[colsToExtract[1]]],
      row[map[colsToExtract[2]]],
      row[map[colsToExtract[3]]]
    ]);
  }

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CFG.resultsSheet) || ss.insertSheet(CFG.resultsSheet);
  sh.clearContents();
  sh.getRange(1,1,slim.length,slim[0].length).setValues(slim);
}



function CLEAR_CACHE() {
  var c = CacheService.getDocumentCache();
  c.remove('index:' + CFG.dataSheet);
  c.remove('paramMap:' + CFG.paramsSheet);
  c.remove('aliases:' + CFG.aliasesSheet);
}


/**
 * Show what star levels exist in the data for a given ship + duration pair.
 * This helps diagnose missing data versus configuration mismatches.
 */
function DIAG_LEVELS_FOR_PAIR(ship, duration) {
  var built = _buildMissionIndex();
  var aliases = _buildAliasesFromSheet();
  var st = _normalizeWithAliases('shipType', ship, aliases);
  var du = _normalizeWithAliases('shipDurationType', duration, aliases);
  
  Logger.log('Checking levels in %s for:', CFG.dataSheet);
  Logger.log('  Ship: %s (from "%s")', st, ship);
  Logger.log('  Duration: %s (from "%s")', du, duration);
  
  var duMap = built.index[st];
  if (!duMap) {
    Logger.log('No data found for ship: %s', st);
    Logger.log('Available ships: %s', JSON.stringify(Object.keys(built.index).sort()));
    return;
  }
  
  duMap = duMap[du];
  if (!duMap) {
    Logger.log('No data found for duration: %s', du);
    Logger.log('Available durations for %s: %s', st, JSON.stringify(Object.keys(built.index[st]).sort()));
    return;
  }
  
  var levels = Object.keys(duMap).sort(function(a,b){return Number(a)-Number(b);});
  Logger.log('Found star levels: %s', JSON.stringify(levels));
}

/**
 * Test level lookup with a common ship configuration.
 */
function DIAG_TEST_LEVELS_FOR_PAIR() {
  DIAG_LEVELS_FOR_PAIR('Atreggies Henliner','Short');
  DIAG_PARAM_FOR_PAIR('Atreggies Henliner', 'Short');
}

/**
 * Show what level is configured in parameters for a ship + duration pair.
 */
function DIAG_PARAM_FOR_PAIR(ship, duration) {
  var params = _buildParamLevelMap();
  var aliases = _buildAliasesFromSheet();
  var st = _normalizeWithAliases('shipType', ship, aliases);
  var du = _normalizeWithAliases('shipDurationType', duration, aliases);
  var key = st + '␞' + du;
  
  Logger.log('Checking configuration in %s for:', CFG.paramsSheet);
  Logger.log('  Ship: %s (from "%s")', st, ship);
  Logger.log('  Duration: %s (from "%s")', du, duration);
  Logger.log('  Configured level: %s', params[key]);
}

/**
 * Display headers from the data sheet to help diagnose column name issues.
 */
function DIAG_HEADERS() {
  var read = _readSheet(CFG.dataSheet);
  Logger.log('Data sheet: %s', CFG.dataSheet);
  Logger.log('Headers: %s', JSON.stringify(read.header));
  
  // Also show parameter sheet headers
  var paramRead = _readSheet(CFG.paramsSheet);
  Logger.log('Param sheet: %s', CFG.paramsSheet);
  Logger.log('Param headers: %s', JSON.stringify(paramRead.header));
}

/**
 * Show sample of indexed ships and their parameter mappings.
 */
function DIAG_SAMPLE() {
  var built = _buildMissionIndex();
  var keysSt = Object.keys(built.index);
  Logger.log('Ship keys: %s', JSON.stringify(keysSt.slice(0,10)));
  
  var param = _buildParamLevelMap();
  var some = [];
  var count = 0;
  for (var k in param) {
    if (param.hasOwnProperty(k)) {
      some.push(k + ' -> ' + param[k]);
      if (++count >= 10) break;
    }
  }
  Logger.log('Param pairs (first 10): %s', JSON.stringify(some));
}

/**
 * Compare parameter pairs from Ship_Parameters with the data index from AllArtifactData.
 * Logs:
 *  - Pairs missing entirely in the index (ship or duration mismatch)
 *  - Pairs present but missing the requested StarsHelper level
 * Relies on existing helpers: _buildMissionIndex(), _buildParamLevelMap()
 */
function DIAG_PARAM_INDEX_MISMATCHES() {
  var built = _buildMissionIndex();              // { index, header, all, cols }
  var index = built.index;                       // index[ship][duration][level] -> array of rowIdx
  var params = _buildParamLevelMap();            // key: 'SHIP␞DURATION' -> level (StarsHelper)

  var missingPairs = [];     // pairs not in index (unknown ship or duration)
  var missingLevels = [];    // pair exists, but StarsHelper level bucket is absent
  var ok = 0;

  for (var key in params) {
    if (!params.hasOwnProperty(key)) continue;
    var parts = key.split('␞');
    var st = parts[0], du = parts[1];
    var lv = String(params[key]);

    var shipNode = index[st];
    if (!shipNode) {
      missingPairs.push(st + ' | ' + du + '  (param=' + lv + ')  -- unknown ship');
      continue;
    }
    var durNode = shipNode[du];
    if (!durNode) {
      missingPairs.push(st + ' | ' + du + '  (param=' + lv + ')  -- unknown duration');
      continue;
    }
    if (!durNode.hasOwnProperty(lv)) {
      var levelsPresent = Object.keys(durNode)
        .map(function(s){ return Number(s); })
        .sort(function(a,b){ return a-b; });
      missingLevels.push(st + ' | ' + du + '  (param=' + lv + ', index levels=' + JSON.stringify(levelsPresent) + ')');
      continue;
    }
    ok++;
  }

  Logger.log('PARAM→INDEX coverage: OK=%s  missingPairs=%s  missingLevels=%s', ok, missingPairs.length, missingLevels.length);

  if (missingPairs.length) {
    Logger.log('Pairs missing in index (ship or duration mismatch):\n%s', JSON.stringify(missingPairs, null, 2));
  }
  if (missingLevels.length) {
    Logger.log('Pairs present but missing requested level bucket:\n%s', JSON.stringify(missingLevels, null, 2));
  }
}

//TODO: check if this works with current _transformALLArtifactData usage
function DIAG_WHAT_SHIPS_DO_I_HAVE(prefix) {
  var built = _buildMissionIndex();
  var keys = Object.keys(built.index).sort();
  var p = String(prefix || '').toUpperCase();
  var out = [];
  for (var i = 0; i < keys.length; i++) {
    if (!p || keys[i].indexOf(p) !== -1) out.push(keys[i]);
  }
  Logger.log('Known shipType keys%s: %s', p ? ' (filter=' + p + ')' : '', JSON.stringify(out));
}

/**
 * Check parameters sheet for values that need aliases added.
 * Reports any ship types or duration types that aren't covered by the aliases sheet.
 */
function DIAG_MISSING_ALIASES_IN_PARAMS() {
  var read = _readSheet(CFG.paramsSheet);
  var header = read.header, rows = read.rows;
  
  // Get column indices using consistent CFG names
  var cSt = header[String(CFG.paramsCols.shipType).toLowerCase()];
  var cDu = header[String(CFG.paramsCols.shipDurationType).toLowerCase()];
  if (cSt == null || cDu == null) {
    throw new Error(
      'Params sheet missing columns: ' +
      CFG.paramsCols.shipType + ' or ' + CFG.paramsCols.shipDurationType
    );
  }

  var aliases = _buildAliasesFromSheet();
  var missing = { shipType: {}, shipDurationType: {} };

  for (var r = 0; r < rows.length; r++) {
    var shipRaw = rows[r][cSt], durRaw = rows[r][cDu];
    var st = _normalizeWithAliases('shipType', shipRaw, aliases);
    var du = _normalizeWithAliases('shipDurationType', durRaw, aliases);

    // Check both raw value and upper_snake version
    var hasShip = aliases.shipType && 
      (aliases.shipType[shipRaw] || aliases.shipType[_toUpperSnake(shipRaw)]);
    var hasDur = aliases.shipDurationType && 
      (aliases.shipDurationType[durRaw] || aliases.shipDurationType[_toUpperSnake(durRaw)]);

    if (!st || !hasShip) missing.shipType[String(shipRaw)] = true;
    if (!du || !hasDur) missing.shipDurationType[String(durRaw)] = true;
  }

  var missShips = Object.keys(missing.shipType);
  var missDurs = Object.keys(missing.shipDurationType);
  
  if (!missShips.length && !missDurs.length) {
    Logger.log('All ' + CFG.paramsSheet + ' values are covered by ' + CFG.aliasesSheet + '.');
  } else {
    if (missShips.length) {
      Logger.log('Ship aliases missing for: %s', JSON.stringify(missShips));
    }
    if (missDurs.length) {
      Logger.log('Duration aliases missing for: %s', JSON.stringify(missDurs));
    }
  }
}


/* =========================
   MENU & TRIGGERS
   ========================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Missions')
    .addItem('Write Results (by Param Level)', 'WRITE_RESULTS_SHEET')
    .addItem('Write Results - Keys Only (4 cols)', 'WRITE_RESULTS_KEYS_ONLY')
    .addSeparator()
    .addItem('Bootstrap Aliases', 'BOOTSTRAP_ALIASES')
    .addToUi();
}

function CREATE_ONCHANGE_TRIGGER_FOR_RESULTS() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'WRITE_RESULTS_SHEET') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('WRITE_RESULTS_SHEET')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();
}