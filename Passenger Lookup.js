function syncAnacData_legacy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("DB_Airports");
  const tempSheet = ss.getSheetByName("TEMP_ANAC_PUBLIC"); 
  
  // 1. Get ALL data from Row 2 downwards
  const tempRange = tempSheet.getRange(2, 1, tempSheet.getLastRow() - 1, tempSheet.getLastColumn());
  const dbRange = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, dbSheet.getLastColumn());
  
  const tempData = tempRange.getValues();
  const dbData = dbRange.getValues();
  
  const tempHeaders = tempData[0];
  const dbHeaders = dbData[0];

  // 2. HELPER: Case-insensitive and Space-trimming header lookup
  const findCol = (headers, name) => {
    const idx = headers.findIndex(h => h.toString().trim().toUpperCase() === name.toUpperCase());
    return idx;
  };

  const tColIcao = findCol(tempHeaders, "CÓDIGO OACI");
  const tColRwy = findCol(tempHeaders, "DESIGNAÇÃO");
  const dColIcao = findCol(dbHeaders, "OACI");
  const dColRwy = findCol(dbHeaders, "Rwy_Ident");

  // 3. SAFETY GATE: Stop if headers don't match
  if (tColIcao === -1 || tColRwy === -1) {
    SpreadsheetApp.getUi().alert("ERROR: Could not find 'CÓDIGO OACI' or 'DESIGNAÇÃO' in Row 2 of TEMP tab. Check for typos!");
    return; // Stop the script here to save your data
  }

  // 4. Map existing DB entries (to preserve Water Zones)
  const dbMap = new Map();
  for (let i = 1; i < dbData.length; i++) {
    const icao = dbData[i][dColIcao];
    const rwy = dbData[i][dColRwy];
    if (icao) {
      const key = icao + "-" + (rwy || "Water");
      dbMap.set(key, dbData[i]);
    }
  }

  const updatedRows = [dbHeaders];

  // 5. Process ANAC Data
  for (let i = 1; i < tempData.length; i++) {
    const row = tempData[i];
    const icao = row[tColIcao];
    const rwyRaw = row[tColRwy];

    if (!icao || !rwyRaw) continue;

    const headings = String(rwyRaw).split('/').map(h => h.trim());
    headings.forEach(h => {
      const key = icao + "-" + h;
      let dbRow = dbMap.get(key) || new Array(dbHeaders.length).fill("");

      dbRow[dColIcao] = icao;
      dbRow[dColRwy] = h;
      dbRow[findCol(dbHeaders, "NOME")] = row[findCol(tempHeaders, "NOME")];
      dbRow[findCol(dbHeaders, "INFRASTRUCTURE_TYPE")] = "Land";
      dbRow[findCol(dbHeaders, "SURFACE_OFFICIAL")] = row[findCol(tempHeaders, "SUPERFÍCIE")];
      dbRow[findCol(dbHeaders, "LENGTH_OFFICIAL")] = row[findCol(tempHeaders, "COMPRIMENTO")];
      
      updatedRows.push(dbRow);
      dbMap.delete(key); // Remove so we don't duplicate
    });
  }

  // 6. ADD BACK manual entries (Water Zones)
  dbMap.forEach(r => updatedRows.push(r));

  // 7. FINAL WRITE
  dbSheet.getRange(2, 1, updatedRows.length, dbHeaders.length).setValues(updatedRows);
  SpreadsheetApp.getUi().alert("Sync Success! Processed " + (updatedRows.length - 1) + " entries.");
}