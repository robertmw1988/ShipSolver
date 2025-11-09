/* =========================
   ALIAS BOOTSTRAP (optional)
   ========================= */
function BOOTSTRAP_ALIASES() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CFG.aliasesSheet) || ss.insertSheet(CFG.aliasesSheet);
  sh.clearContents();
  sh.getRange(1,1,1,3).setValues([[ 'kind','alias','canonical' ]]);

  // From Ship_Parameters (ships & durations, if present)
  try {
    var sp = _readSheet(CFG.shipParametersSheet);
    var cShip = sp.header['ship'];
    var cType = sp.header['type'];

    if (cShip != null) {
      var seenShip = {};
      var shipRows = [];
      for (var i = 0; i < sp.rows.length; i++) {
        var s = String(sp.rows[i][cShip] || '').trim();
        if (!s || seenShip[s]) continue;
        seenShip[s] = true;
        shipRows.push(['shipType', s, _toUpperSnake(s)]);
      }
      if (shipRows.length) sh.getRange(sh.getLastRow()+1,1,shipRows.length,3).setValues(shipRows);
    }
    if (cType != null) {
      var seenType = {};
      var typeRows = [];
      for (var j = 0; j < sp.rows.length; j++) {
        var t = String(sp.rows[j][cType] || '').trim();
        if (!t || seenType[t]) continue;
        seenType[t] = true;
        typeRows.push(['shipDurationType', t, _toUpperSnake(t)]);
      }
      if (typeRows.length) sh.getRange(sh.getLastRow()+1,1,typeRows.length,3).setValues(typeRows);
    }
  } catch (e) {
    // optional
  }

  // From AllArtifactData
  try {
    var read = _readSheet(CFG.dataSheet);
    var header = read.header, rows = read.rows;
    var cols = _resolveCols(header, CFG.dataCols);

    var sets = { shipType:{}, shipDurationType:{}, targetArtifact:{} };
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      var st = String(row[cols.shipType] || '').trim();
      var du = String(row[cols.shipDurationType] || '').trim();
      var ta = String(row[cols.targetArtifact] || '').trim();
      if (st) sets.shipType[st] = true;
      if (du) sets.shipDurationType[du] = true;
      if (ta) sets.targetArtifact[ta] = true;
    }
    var aliasRows = [];
    var k, v;
    for (k in sets.shipType) if (sets.shipType.hasOwnProperty(k)) aliasRows.push(['shipType', k, _toUpperSnake(k)]);
    for (k in sets.shipDurationType) if (sets.shipDurationType.hasOwnProperty(k)) aliasRows.push(['shipDurationType', k, _toUpperSnake(k)]);
    for (k in sets.targetArtifact) if (sets.targetArtifact.hasOwnProperty(k)) aliasRows.push(['targetArtifact', k, _toUpperSnake(k)]);

    if (aliasRows.length) sh.getRange(sh.getLastRow()+1,1,aliasRows.length,3).setValues(aliasRows);
  } catch (e2) {
    // optional
  }
}