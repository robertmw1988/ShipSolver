function onEdit(e) {
  const editedSheet = e.range.getSheet();
  if (editedSheet.getName() !== "Comparison") return; // trigger only on Comparison sheet

  updateBestShips();
}

function updateBestShips() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const totalsSheet = ss.getSheetByName("Totals");
  const compSheet = ss.getSheetByName("Comparison");

  // Headers to output
  const outputHeaders = ["Ship", "Target", "Hours", "Ships/Tank", "Artis/Ship", "Artis/Tank", "Artis/Day", "% Score"];

  // Read checkbox table
  const shipNames = compSheet.getRange("L23:L33").getValues().flat().map(normalize);
  const shipChecks = compSheet.getRange("M23:M33").getValues().flat();
  const allowedShips = shipNames.filter((_, i) => shipChecks[i]);

  // Categories
  const infiniteFuelCategories = normalizeList([
    "Chicken One","Chicken Nine","Chicken Heavy","BCR","Quintillion Chicken","Cornish-Hen Corvette","Galeggtica"
  ]);

  const fuelLimitedCategories = normalizeList([
    "Defihent","Voyegger","Henerprise","Atreggies Henliner"
  ]);

  // Pull Totals data including column R (Ships/Tank)
  const rawData = totalsSheet.getRange("A3:R").getValues(); // columns A-R

  // Filter by checkbox & prepare normalized names
  const data = rawData
    .filter(r => r[0] && r[16] !== "" && r[16] !== null) // Ship + % Score exists
    .map(r => ({
      row: r,
      nameNorm: normalize(r[0])
    }))
    // Only include ships allowed by checkbox
    .filter(r => allowedShips.some(prefix => shipMatchesPrefix(r.row[0], prefix)));

  // Split by category
  const infiniteFuelData = data.filter(r => infiniteFuelCategories.some(c => shipMatchesPrefix(r.row[0], c)));
  const fuelLimitedData = data.filter(r => fuelLimitedCategories.some(c => shipMatchesPrefix(r.row[0], c)));

  // Sort descending by % Score (Q = index 16)
  infiniteFuelData.sort((a,b) => b.row[16] - a.row[16]);
  fuelLimitedData.sort((a,b) => b.row[16] - a.row[16]);

  // Top N, including Hours (C = index 2) and Ships/Tank (R = index 17)
  const topInfinite = infiniteFuelData.slice(0,3).map(r => [
    r.row[0],  // Ship
    r.row[1],  // Target
    r.row[2],  // Hours
    r.row[17], // Ships/Tank
    r.row[7],  // Artis/Ship
    r.row[8],  // Artis/Tank
    r.row[9],  // Artis/Day
    r.row[16]  // % Score
  ]);

  const topLimited = fuelLimitedData.slice(0,5).map(r => [
    r.row[0],  // Ship
    r.row[1],  // Target
    r.row[2],  // Hours
    r.row[17], // Ships/Tank
    r.row[7],  // Artis/Ship
    r.row[8],  // Artis/Tank
    r.row[9],  // Artis/Day
    r.row[16]  // % Score
  ]);

  // Clear old output
  compSheet.getRange("P1:W6").clearContent();
  compSheet.getRange("Y1:AF4").clearContent();

  // Write headers + data
  compSheet.getRange("P1:W1").setValues([outputHeaders]);
  if (topLimited.length > 0) {
    compSheet.getRange(2,16,topLimited.length,8).setValues(topLimited);
  }

  compSheet.getRange("Y1:AF1").setValues([outputHeaders]);
  if (topInfinite.length > 0) {
    compSheet.getRange(2,25,topInfinite.length,8).setValues(topInfinite);
  }
}

// Utility functions
function normalize(text) {
  return String(text).trim().toLowerCase();
}

function normalizeList(arr) {
  return arr.map(normalize);
}

// Robust prefix match for multi-word ships
function shipMatchesPrefix(shipName, prefix) {
  const s = normalize(shipName);
  const p = normalize(prefix);
  // Escape regex special characters in prefix
  const safePrefix = p.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
  // Match prefix as whole words at start
  const regex = new RegExp('^' + safePrefix + '(\\s|$)');
  return regex.test(s);
}
