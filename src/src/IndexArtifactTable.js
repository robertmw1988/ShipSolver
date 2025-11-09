/*******************************************************
 * Egg Inc. Virtue - Mission Retrieval & Solver Prep (ES5)
 * No const/let, no =>, no template literals, no ||= ?? ?.
 *******************************************************/

/* =========================
   CONFIGURATION
   ========================= */

var CFG = {
  dataSheet: 'AllArtifactData',
  shipParametersSheet: 'Ship_Parameters',
  // Use Ship_Parameters as the parameter source:
  paramsSheet: 'Ship_Parameters',

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
    shipType: 'Ship',          // <- header in Ship_Parameters
    shipDurationType: 'Type',  // <- header in Ship_Parameters (Short/Standard/Extended/Tutorial)
    level: 'StarsHelper'       // <- user-configured star level per ship & type
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
    'SHORT',
    'LONG',
    'EPIC',
    'TUTORIAL',
  ]),

  targetArtifact: Object.freeze([
    'LUNAR_TOTEM',
    'UNKNOWN',
    'TACHYON_STONE',
    'BOOK_OF_BASAN',
    'PHOENIX_FEATHER',
    'TUNGSTEN_ANKH',
    'GOLD_METEORITE',
    'TAU_CETI_GEODE',
    'TACHYON_STONE_FRAGMENT',
    'AURELIAN_BROOCH',
    'CARVED_RAINSTICK',
    'PUZZLE_CUBE',
    'QUANTUM_METRONOME',
    'SHIP_IN_A_BOTTLE',
    'TACHYON_DEFLECTOR',
    'INTERSTELLAR_COMPASS',
    'DILITHIUM_MONOCLE',
    'TITANIUM_ACTUATOR',
    'NEODYMIUM_MEDALLION',
    'MERCURYS_LENS',
    'DILITHIUM_STONE',
    'SHELL_STONE',
    'LUNAR_STONE',
    'SOUL_STONE',
    'QUANTUM_STONE',
    'TERRA_STONE',
    'LIFE_STONE',
    'PROPHECY_STONE',
    'BEAK_OF_MIDAS',
    'CLARITY_STONE',
    'SOLAR_TITANIUM',
    'DILITHIUM_STONE_FRAGMENT',
    'SHELL_STONE_FRAGMENT',
    'LUNAR_STONE_FRAGMENT',
    'SOUL_STONE_FRAGMENT',
    'PROPHECY_STONE_FRAGMENT',
    'QUANTUM_STONE_FRAGMENT',
    'LIGHT_OF_EGGENDIL',
    'TERRA_STONE_FRAGMENT',
    'LIFE_STONE_FRAGMENT',
    'CLARITY_STONE_FRAGMENT',
    'DEMETERS_NECKLACE',
    'VIAL_MARTIAN_DUST',
    'ORNATE_GUSSET',
    'THE_CHALICE',
  ]),
});

/* =========================
   UTILITIES
   ========================= */
function _toUpperSnake(s) {
  if (s == null) return '';
  return String(s).trim().replace(/[\s\-]+/g, '_').toUpperCase();
}

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
   ALIASES (ES5)
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
  shipType:        ['shiptype','ship','ship_name'],
  shipDurationType:['shipdurationtype','durationtype','type','missiontype','duration'],
  targetArtifact:  ['targetartifact','target','artifact'],
  missionLevel:    ['missionlevel','level','lvl']
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
        if (header.hasOwnProperty(alt)) { idx = header[alt]; break; }
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

  var need = ['shiptype','shipdurationtype','targetartifact'];
  for (var n = 0; n < need.length; n++) if (map[need[n]] == null) throw new Error('Missing column in results: ' + need[n]);

  var lvlKey = map.hasOwnProperty('level') ? 'level' : (map.hasOwnProperty('missionlevel') ? 'missionlevel' : null);
  
  if (!lvlKey) throw new Error('Missing column in results: level');
  var slim = [['shipType','shipDurationType','level','targetArtifact']];

  for (var r = 1; r < full.length; r++) {
    var row = full[r];
    slim.push([ row[map['shiptype']], row[map['shipdurationtype']], row[map[lvlKey]], row[map['targetartifact']] ]);
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

function DIAG_LEVELS_FOR_PAIR(ship, duration) {
  var built = _buildMissionIndex();
  var st = _normalizeWithAliases('shipType', ship, _buildAliasesFromSheet());
  var du = _normalizeWithAliases('shipDurationType', duration, _buildAliasesFromSheet());
  var duMap = built.index[st];
  if (!duMap) { Logger.log('No ship in index: %s', st); return; }
  duMap = duMap[du];
  if (!duMap) { Logger.log('No duration in index: %s | %s', st, du); return; }
  var levels = Object.keys(duMap).sort(function(a,b){return Number(a)-Number(b);});
  Logger.log('Index levels present for %s | %s: %s', st, du, JSON.stringify(levels));
}

function DIAG_TEST_LEVELS_FOR_PAIR() {
  DIAG_LEVELS_FOR_PAIR('Atreggies Henliner','Short');
  DIAG_PARAM_FOR_PAIR('Atreggies Henliner', 'Short');
}

function DIAG_PARAM_FOR_PAIR(ship, duration) {
  var params = _buildParamLevelMap();
  var aliases = _buildAliasesFromSheet();
  var st = _normalizeWithAliases('shipType', ship, aliases);
  var du = _normalizeWithAliases('shipDurationType', duration, aliases);
  var key = st + '␞' + du;
  Logger.log('Param level for %s | %s => %s', st, du, params[key]);
}

function DIAG_HEADERS() {
  var read = _readSheet(CFG.dataSheet);
  Logger.log('Index sheet: %s', CFG.dataSheet);
  Logger.log('Headers: %s', JSON.stringify(read.header));
}

function DIAG_SAMPLE() {
  var built = _buildMissionIndex();
  var keysSt = Object.keys(built.index);
  Logger.log('Ship keys: %s', JSON.stringify(keysSt.slice(0,10)));
  var param = _buildParamLevelMap();
  var some = [];
  var count = 0;
  for (var k in param) { if (param.hasOwnProperty(k)) { some.push(k + ' -> ' + param[k]); if (++count>=10) break; } }
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


function DIAG_MISSING_ALIASES_IN_PARAMS() {
  var read    = _readSheet(CFG.paramsSheet);
  var header  = read.header, rows = read.rows;
  var cSt = header[String(CFG.paramsCols.shipType).toLowerCase()];
  var cDu = header[String(CFG.paramsCols.shipDurationType).toLowerCase()];
  if (cSt == null || cDu == null) throw new Error('Params sheet missing Ship and/or Type columns');

  var aliases = _buildAliasesFromSheet();
  var missing = { shipType: {}, shipDurationType: {} };

  for (var r = 0; r < rows.length; r++) {
    var shipRaw = rows[r][cSt], durRaw = rows[r][cDu];
    var st = _normalizeWithAliases('shipType', shipRaw, aliases);
    var du = _normalizeWithAliases('shipDurationType', durRaw, aliases);

    var hasShip =
      (aliases.shipType && (aliases.shipType[shipRaw] || aliases.shipType[_toUpperSnake(shipRaw)]));
    var hasDur =
      (aliases.shipDurationType && (aliases.shipDurationType[durRaw] || aliases.shipDurationType[_toUpperSnake(durRaw)]));

    if (!st || !hasShip) missing.shipType[String(shipRaw)] = true;
    if (!du || !hasDur)  missing.shipDurationType[String(durRaw)] = true;
  }

  var missShips = Object.keys(missing.shipType);
  var missDurs  = Object.keys(missing.shipDurationType);
  if (!missShips.length && !missDurs.length) {
    Logger.log('All Ship_Parameters values are covered by Aliases.');
  } else {
    if (missShips.length) Logger.log('Ship aliases missing for: %s', JSON.stringify(missShips));
    if (missDurs.length)  Logger.log('Duration aliases missing for: %s', JSON.stringify(missDurs));
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