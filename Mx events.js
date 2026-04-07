/**
 * MASTER MAINTENANCE ENGINE
 * Combines ANAC Scraping + Tach Matrix Sync
 */

// 1. THE "BOSS" FUNCTION (Set your Nightly Trigger to this)
function runAllMaintenanceUpdates() {
  console.log("Starting Master Maintenance Update...");
  
  // Step A: Pull fresh dates from ANAC RAB
  updateFleetAndAlert(); 
  
  // Step B: Sync those dates with the Tach Matrix
  nightlyMaintenanceSync();
  
  console.log("Master Update Finished.");
}

// 2. THE ANAC SCRAPER (Date-based legality)
function updateFleetAndAlert() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName("DB_Aircraft");
  if (!dashSheet) return;

  const myEmail = "tecnico.mx@asasdesocorro.org.br";
  const url = "https://sistemas.anac.gov.br/dadosabertos/Aeronaves/RAB/dados_aeronaves.csv";
  
  const dashHeaders = dashSheet.getRange(1, 1, 1, dashSheet.getLastColumn()).getValues()[0];
  const cvaColIdx = dashHeaders.indexOf("Annual Due (CVA)") + 1;
  const tailData = dashSheet.getRange(2, 1, dashSheet.getLastRow()-1, 1).getValues();

  try {
    const response = UrlFetchApp.fetch(url);
    const content = response.getContentText("ISO-8859-1");
    const lines = content.split('\n').filter(line => line.trim().length > 0);
    
    let headerRowIndex = -1;
    for (let i = 0; i < 5; i++) {
      if (lines[i].includes("MARCA")) { headerRowIndex = i; break; }
    }

    let headerLine = lines[headerRowIndex];
    if (headerLine.includes("Atualizado em:")) {
      headerLine = headerLine.substring(headerLine.indexOf("MARCA"));
    }

    const headers = headerLine.split(';').map(h => h.replace(/"/g, '').trim().toUpperCase());
    const marcaIdx = headers.indexOf("MARCA");
    const cvaIdx = headers.indexOf("DT_VALIDADE_CVA");

    let alerts = [];

    tailData.forEach((row, index) => {
      const tailName = row[0];
      if (!tailName) return;

      let searchMark = tailName.replace("-", "").toUpperCase();
      let foundData = null;

      for (let i = headerRowIndex + 1; i < lines.length; i++) {
        let cols = lines[i].split(';');
        if (cols[marcaIdx] && cols[marcaIdx].replace(/"/g, '').trim().toUpperCase() === searchMark) {
          foundData = cols[cvaIdx] ? cols[cvaIdx].replace(/"/g, '').trim() : "";
          break;
        }
      }
      
      if (foundData && foundData !== "null" && foundData !== "") {
        let formattedDate = foundData.replace(/(\d{2})(\d{2})(\d{4})/, "$1/$2/$3");
        dashSheet.getRange(index + 2, cvaColIdx).setValue(formattedDate);
        
        let day = foundData.substring(0,2);
        let month = foundData.substring(2,4);
        let year = foundData.substring(4,8);
        let expiryDate = new Date(year, month - 1, day);
        let today = new Date();
        let diffDays = Math.ceil((expiryDate - today) / (1000 * 60 * 60 * 24));

        if (diffDays <= 30) {
          let status = diffDays < 0 ? 'VENCIDO' : 'vence em ' + diffDays + ' dias';
          alerts.push(`⚠️ ${tailName}: CVA ${status} (${formattedDate})`);
        }
      }
    });

    if (alerts.length > 0) {
      MailApp.sendEmail(myEmail, "ALERTA RAB: Vencimento de CVA", "Atenção:\n\n" + alerts.join("\n"));
    }
  } catch (e) {
    console.log("Erro no Scraper: " + e.message);
  }
}

// 3. THE TACH SYNC (Hour-based legality)
function nightlyMaintenanceSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName("DB_Aircraft");
  const matrixSheet = ss.getSheetByName("DB_Component_Matrix");
  
  if (!dashSheet || !matrixSheet) return;

  const dashData = dashSheet.getDataRange().getValues();
  const matrixData = matrixSheet.getDataRange().getValues();
  const dashHeaders = dashData[0];
  
  const col = {
    tail: dashHeaders.indexOf("Tail #"),
    curTach: dashHeaders.indexOf("Current Tach"),
    nextDue: dashHeaders.indexOf("Next Due (Tach)"),
    hrsLeft: dashHeaders.indexOf("Hours to Insp"),
    cva: dashHeaders.indexOf("Annual Due (CVA)"),
    status: dashHeaders.indexOf("Tech Status")
  };

  for (let i = 1; i < dashData.length; i++) {
    const tail = dashData[i][col.tail];
    const currentTach = parseFloat(dashData[i][col.curTach]) || 0;
    const cvaDateStr = dashData[i][col.cva];
    
    let lowestTachDue = Infinity;

    for (let j = 1; j < matrixData.length; j++) {
      if (matrixData[j][0] === tail) {
        let itemDue = parseFloat(matrixData[j][4]); // "Next Due (Tach)" column
        if (itemDue < lowestTachDue) lowestTachDue = itemDue;
      }
    }

    const hrsRemaining = lowestTachDue === Infinity ? 9999 : (lowestTachDue - currentTach).toFixed(1);
    
    let isDateExpired = false;
    if (cvaDateStr && cvaDateStr !== "Sem data no RAB") {
      const parts = cvaDateStr.split('/');
      const expiry = new Date(parts[2], parts[1] - 1, parts[0]);
      if (expiry < new Date()) isDateExpired = true;
    }

    let techStatus = "READY";
    if (hrsRemaining <= 0) techStatus = "AOG - MX OVERDUE";
    if (isDateExpired) techStatus = "AOG - CVA EXPIRED";
    
    const rowNum = i + 1;
    dashSheet.getRange(rowNum, col.nextDue + 1).setValue(lowestTachDue === Infinity ? "" : lowestTachDue);
    dashSheet.getRange(rowNum, col.hrsLeft + 1).setValue(hrsRemaining);
    dashSheet.getRange(rowNum, col.status + 1).setValue(techStatus);
    
    // Formatting
    const statusCell = dashSheet.getRange(rowNum, col.status + 1);
    if (techStatus.includes("AOG")) {
      statusCell.setBackground("#f4cccc").setFontColor("#990000").setFontWeight("bold");
    } else if (hrsRemaining < 10) {
      statusCell.setBackground("#fff2cc").setFontColor("#bf9000").setFontWeight("bold");
    } else {
      statusCell.setBackground("#d9ead3").setFontColor("#274e13").setFontWeight("normal");
    }
  }
}