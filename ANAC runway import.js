function syncAnacPrivateData_legacy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("DB_Airports");
  const tempSheet = ss.getSheetByName("TEMP_ANAC_PRIV"); // <--- DOUBLE CHECK THIS NAME
  
  // 1. Safety Check: Does the temp sheet actually exist?
  if (!tempSheet) {
    SpreadsheetApp.getUi().alert("Error: Sheet 'TEMP_ANAC_PRIV' not found. Please check the tab name!");
    return;
  }
  if (!dbSheet) {
    SpreadsheetApp.getUi().alert("Error: Sheet 'DB_Airports' not found!");
    return;
  }

  // 2. Get Headers
  const dbHeaders = dbSheet.getRange(1, 1, 1, dbSheet.getLastColumn()).getValues()[0];
  const tempHeaders = tempSheet.getRange(2, 1, 1, tempSheet.getLastColumn()).getValues()[0];
  
  const findCol = (headers, nameList) => {
    for (let name of nameList) {
      const idx = headers.findIndex(h => h.toString().trim().toUpperCase() === name.toUpperCase());
      if (idx !== -1) return idx;
    }
    return -1;
  };

  // Map Source Columns
  const tColIcao = findCol(tempHeaders, ["CÓDIGO OACI", "OACI"]);
  const tColRwy  = findCol(tempHeaders, ["DESIGNAÇÃO", "Pista"]);
  const tColAlt  = findCol(tempHeaders, ["ALTITUDE", "ELEV"]);
  const tColLat  = findCol(tempHeaders, ["LATITUDE"]);
  const tColLon  = findCol(tempHeaders, ["LONGITUDE"]);

  // Map Destination Columns
  const dColIcao = findCol(dbHeaders, ["OACI"]);
  const dColRwy  = findCol(dbHeaders, ["Rwy_Ident"]);
  const dColAlt  = findCol(dbHeaders, ["ALTITUDE"]);

  // 3. Helper to handle semicolons and Brazilian Number Formats (Commas)
  const cleanPrivateData = (val) => {
    if (!val) return "";
    // Take the first part before a semicolon, replace comma with dot for math
    let clean = val.toString().split(';')[0].trim().replace(',', '.');
    return clean;
  };

  const toDec = (dms) => {
    if (!dms) return "";
    let cleanDms = cleanPrivateData(dms);
    const p = cleanDms.match(/(\d+)°\s*(\d+)'\s*(\d+)''\s*([NSEW])/);
    if (!p) return cleanDms;
    let d = parseFloat(p[1]) + parseFloat(p[2])/60 + parseFloat(p[3])/3600;
    return (p[4] === 'S' || p[4] === 'W') ? d * -1 : d;
  };

  // 4. Get Data
  const tempLastRow = tempSheet.getLastRow();
  if (tempLastRow < 3) {
    SpreadsheetApp.getUi().alert("The Private Temp sheet appears to have no data below the headers.");
    return;
  }
  const tempData = tempSheet.getRange(3, 1, tempLastRow - 2, tempSheet.getLastColumn()).getValues();
  const dbData = dbSheet.getLastRow() > 1 ? dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, dbSheet.getLastColumn()).getValues() : [];

  const dbMap = new Map();
  dbData.forEach(row => {
    const key = row[dColIcao] + "-" + (row[dColRwy] || "Water");
    dbMap.set(key, row);
  });

  const finalRows = [];

  // 5. Process Rows
  tempData.forEach(row => {
    const icao = cleanPrivateData(row[tColIcao]);
    const rwyRaw = row[tColRwy].toString();

    if (!icao || !rwyRaw) return;

    // Split runway by semicolon OR slash
    const headings = rwyRaw.split(/[;/]+/).map(h => h.trim());

    headings.forEach(h => {
      const key = icao + "-" + h;
      let dbRow = dbMap.get(key) || new Array(dbHeaders.length).fill("");

      dbRow[dColIcao] = icao;
      dbRow[dColRwy]  = h;
      dbRow[findCol(dbHeaders, ["NOME"])] = cleanPrivateData(row[findCol(tempHeaders, ["NOME"])]);
      dbRow[findCol(dbHeaders, ["INFRASTRUCTURE_TYPE"])] = "Private Land";
      
      if (dColAlt !== -1 && tColAlt !== -1) {
        let altVal = cleanPrivateData(row[tColAlt]).replace(/[^\d.-]/g, '');
        dbRow[dColAlt] = altVal ? parseFloat(altVal) : 0;
      }
      
      dbRow[findCol(dbHeaders, ["LATITUDE"])] = toDec(row[tColLat]);
      dbRow[findCol(dbHeaders, ["LONGITUDE"])] = toDec(row[tColLon]);
      
      finalRows.push(dbRow);
      dbMap.delete(key); 
    });
  });

  dbMap.forEach(row => finalRows.push(row));

  // 6. Write back
  dbSheet.getRange(2, 1, finalRows.length, dbHeaders.length).setValues(finalRows);
  SpreadsheetApp.getUi().alert("Private Database Sync Complete! Processed " + finalRows.length + " headings.");
}