function processLogbookScan(base64Image, sheetId) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey) {
      return { error: "GEMINI_API_KEY not found in script properties." };
    }

    // Call Gemini 1.5 Pro to analyze the image
    var prompt = `
    Analyze the following image of a pilot's flight logbook. 
    Extract the flight entries and format them as a JSON array of objects.
    Each object should have these keys exactly: "date", "aircraftType", "tailNumber", "route", "departureTime", "arrivalTime", "duration", "picName", "landings", "remarks".
    
    If data is missing for a field, use an empty string. 
    Optimize and correct common OCR errors (e.g., 'O' instead of '0', 'I' instead of '1' where appropriate in tail numbers or times).
    Ensure durations are in a consistent format (e.g., HH:MM).

    Return ONLY the JSON array.
    `;

    var payload = {
      "contents": [
        {
          "parts": [
            { "text": prompt },
            {
              "inline_data": {
                "mime_type": "image/jpeg",
                "data": base64Image
              }
            }
          ]
        }
      ]
    };

    var options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    var geminiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro:generateContent?key=" + apiKey;
    var response = UrlFetchApp.fetch(geminiUrl, options);
    var resultText = response.getContentText();
    var resultObj = JSON.parse(resultText);

    if (resultObj.error) {
      return { error: "Gemini API Error: " + resultObj.error.message };
    }

    var extractedText = resultObj.candidates[0].content.parts[0].text;
    
    // Extract JSON from markdown
    var jsonMatch = extractedText.match(/\[[\s\S]*\]/);
    if (!jsonMatch) {
       return { error: "Could not parse JSON from AI response." };
    }
    
    var entries = JSON.parse(jsonMatch[0]);
    if (!entries || entries.length === 0) {
      return { error: "No flight entries found." };
    }

    var targetSheetId = sheetId;
    
    // Create new sheet if none provided
    if (!targetSheetId) {
      var ss = SpreadsheetApp.create("Logbook Scan " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"));
      targetSheetId = ss.getId();
      var sheet = ss.getSheets()[0];
      sheet.setName("Flights");
      sheet.appendRow([
        "Date", "Aircraft Type", "Tail Number", "Route", "Departure", 
        "Arrival", "Duration", "PIC Name", "Landings", "Remarks"
      ]);
      sheet.setFrozenRows(1);
    }

    // Append data
    var ss = SpreadsheetApp.openById(targetSheetId);
    var sheet = ss.getSheetByName("Flights") || ss.getSheets()[0];
    
    var rowsData = entries.map(function(e) {
      return [
        e.date || "", e.aircraftType || "", e.tailNumber || "", e.route || "", 
        e.departureTime || "", e.arrivalTime || "", e.duration || "", 
        e.picName || "", e.landings || "", e.remarks || ""
      ];
    });

    sheet.getRange(sheet.getLastRow() + 1, 1, rowsData.length, rowsData[0].length).setValues(rowsData);

    return { 
      success: true, 
      count: entries.length,
      sheetId: targetSheetId,
      rawData: entries 
    };

  } catch (e) {
    Logger.log(e);
    return { error: e.toString() };
  }
}
