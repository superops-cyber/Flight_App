function logSheetsAndHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  let output = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    const lastCol = sheet.getLastColumn();

    let headers = [];
    if (lastCol > 0) {
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
        .filter(h => h !== '');
    }

    output.push(`=== ${name} ===`);
    output.push(headers.length ? headers.join(', ') : '(no headers)');
    output.push(''); // blank line
  });

  Logger.log(output.join('\n'));
}
