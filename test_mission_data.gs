// Quick test to fetch ADS26-001-2 and see what passengers/aircraft/fuel it has
function testMissionData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DB_Dispatch");
  const data = sheet.getDataRange().getValues();
  
  let missionRows = data.filter(r => String(r[1]) === "ADS26-001-2");
  console.log("Found rows for ADS26-001-2:", missionRows.length);
  
  if (missionRows.length > 0) {
    const row = missionRows[0];
    console.log("Date:", row[2]);
    console.log("Aircraft:", row[3]);
    console.log("Pilot:", row[4]);
    console.log("Route:", row[6]);
    console.log("Full JSON (col 9):", row[9]);
    
    try {
      const payload = JSON.parse(row[9] || "{}");
      console.log("Payload legs:", payload.legs ? payload.legs.length : 0);
      if (payload.legs && payload.legs[0]) {
        console.log("First leg pax:", payload.legs[0].pax ? payload.legs[0].pax.length : 0);
      }
    } catch (e) {
      console.log("Parse error:", e.message);
    }
  }
}
