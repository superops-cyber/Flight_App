function convertRangeToFeet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("DB_Airports");
  const range = dbSheet.getRange("I2:I924");
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] && !isNaN(values[i][0])) {
      // Multiply meters by 3.28084 and round to nearest foot
      values[i][0] = Math.round(values[i][0] * 3.28084);
    }
  }
  range.setValues(values);
}