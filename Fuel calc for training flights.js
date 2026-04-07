function calculateRequiredFuel() {

  const ss = SpreadsheetApp.getActive();
  const planSheet = ss.getSheetByName("REF_Syllabus");
  const dbSheet = ss.getSheetByName("DB_Aircraft");

  const planData = planSheet.getDataRange().getValues();
  const dbData = dbSheet.getDataRange().getValues();

  // Build aircraft lookup table
  const aircraftDB = {};
  for (let i = 1; i < dbData.length; i++) {
    const type = dbData[i][0].toString().toUpperCase();
    aircraftDB[type] = {
      burn: Number(dbData[i][1]),
      pattern: Number(dbData[i][2])
    };
  }

  // Get column indexes
  const headers = planData[0];
  const colTarget = headers.indexOf("Target_Role");
  const colHours = headers.indexOf("Required_Hours");
  const colLandings = headers.indexOf("Planned_Landings");
  const colFuel = headers.indexOf("Required_Fuel");

  for (let r = 1; r < planData.length; r++) {

    const roleText = planData[r][colTarget].toString().toUpperCase();
    let aircraft = null;

    if (roleText.includes("C172")) aircraft = "C172";
    else if (roleText.includes("206") && roleText.includes("ANF")) aircraft = "C206 ANF";
    else if (roleText.includes("206")) aircraft = "C206 Land";

    if (!aircraft || !aircraftDB[aircraft]) continue;

    const hours = Number(planData[r][colHours]) || 0;
    const landings = Number(planData[r][colLandings]) || 0;

    const burnLPH = aircraftDB[aircraft].burn;
    const patternBurn = aircraftDB[aircraft].pattern;

    const cruiseFuel = burnLPH * hours;
    const patternFuel = patternBurn * landings;
    const reserveFuel = burnLPH; // 1 hour reserve

    const totalFuel = cruiseFuel + patternFuel + reserveFuel;

    planSheet.getRange(r + 1, colFuel + 1).setValue(totalFuel);
  }
}
