/**
 * Updates Pilot Qualifications based on Training Codes in the Flight Log
 */
function updatePilotQualifications(pilotEmail, trainingCode, isSuccessful) {
  if (!isSuccessful) return "Training not successful; no qualification update.";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pilotSheet = ss.getSheetByName("DB_Pilots");
  const pilotData = pilotSheet.getDataRange().getValues();
  const headers = pilotData[0];
  
  // 1. Identify the Column Index for "Aircraft/Roll"
  const rollColIdx = headers.indexOf("Aircraft/Roll");
  
  // 2. Calculate the new Expiry Date (Today + 365 days)
  const today = new Date();
  const nextYear = new Date(today.getFullYear() + 1, today.getMonth(), today.getDate());
  const expiryStr = Utilities.formatDate(nextYear, "GMT-3", "dd/MM/yyyy");

  // 3. Define the Role based on the Training Code
  // You can expand this mapping as you create more codes
  let newRole = "";
  if (trainingCode.includes("172-CFI")) newRole = "172 Instructor";
  if (trainingCode.includes("206L-OP")) newRole = "206 Land Operational";
  if (trainingCode.includes("206A-INST")) newRole = "206 Amphib Instructor";
  
  if (newRole === "") return "Code recognized, but no role mapping found.";

  // 4. Find the Pilot and update the string
  for (let i = 1; i < pilotData.length; i++) {
    if (pilotData[i][headers.indexOf("Email")] === pilotEmail) {
      let currentRolls = pilotData[i][rollColIdx] || "";
      let newEntry = `${newRole} (Exp: ${expiryStr})`;
      
      let updatedRolls;
      // If the role already exists, replace it (Update). If not, append it (New Qual).
      if (currentRolls.includes(newRole)) {
        // Regex to find the specific role and its old expiry to swap it out
        const regex = new RegExp(newRole + " \\(Exp: .*?\\)");
        updatedRolls = currentRolls.replace(regex, newEntry);
      } else {
        updatedRolls = currentRolls + (currentRolls ? ", " : "") + newEntry;
      }

      pilotSheet.getRange(i + 1, rollColIdx + 1).setValue(updatedRolls);
      return `Success: ${pilotEmail} updated to ${newRole} expiring ${expiryStr}`;
    }
  }
}