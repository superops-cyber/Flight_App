function diagnosticANAC() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DB_Pilots");
  const data = sheet.getRange("A2:E2").getValues()[0]; 
  
  const canac = data[3].toString().trim();
  const cpf = data[1].toString().replace(/\D/g, "");
  const dob = data[2];
  const dobStr = (dob instanceof Date) ? Utilities.formatDate(dob, "GMT", "dd/MM/yyyy") : dob;

  const initialUrl = "https://sistemas.anac.gov.br/habilitacao/ListarCriarAgendamento/ConsultarPublico.do";
  const postUrl = "https://sistemas.anac.gov.br/habilitacao/ListarCriarAgendamento/ResultadoConsultarPublico.do";

  try {
    // STEP 1: Get the Session Cookie
    const initialResponse = UrlFetchApp.fetch(initialUrl);
    const cookie = initialResponse.getAllHeaders()['Set-Cookie'];
    
    // STEP 2: Send the data WITH the cookie
    const payload = {
      "txtCanac": canac,
      "txtCpf": cpf,
      "txtDataNascimento": dobStr,
      "btnConsultar": "Consultar"
    };

    const options = {
      "method": "post",
      "payload": payload,
      "headers": { "Cookie": cookie },
      "followRedirects": true,
      "muteHttpExceptions": true
    };

    const response = UrlFetchApp.fetch(postUrl, options);
    const html = response.getContentText();
    
    // Now check if we actually found the pilot's page
    if (html.includes("Identificação do Piloto") || html.includes(canac)) {
       Logger.log("--- SUCCESS: FOUND PILOT DATA ---");
       // Let's find IFRA now
       const start = html.indexOf("IFRA");
       if (start !== -1) {
         Logger.log(html.substring(start, start + 400));
       } else {
         Logger.log("Found page, but no 'IFRA' text. This is odd.");
       }
    } else {
      Logger.log("STILL REDIRECTED. Title: " + html.match(/<title>(.*?)<\/title>/)[1]);
      Logger.log("Check if CPF/DOB match exactly what is in your sheet.");
    }

  } catch (e) {
    Logger.log("Error: " + e.message);
  }
}