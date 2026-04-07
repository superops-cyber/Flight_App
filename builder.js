function calculateFlightPerformance(weight, factors) {
  // 1. BASE DATA LOOKUP
  const perfData = [
    {w: 1200, to: 110, ldg: 138}, {w: 1250, to: 117, ldg: 143},
    {w: 1300, to: 120, ldg: 148}, {w: 1350, to: 127, ldg: 153},
    {w: 1400, to: 133, ldg: 158}, {w: 1450, to: 140, ldg: 163},
    {w: 1500, to: 147, ldg: 168}, {w: 1550, to: 150, ldg: 173},
    {w: 1600, to: 157, ldg: 178}, {w: 1650, to: 163, ldg: 183},
    {w: 1700, to: 167, ldg: 188}, {w: 1723, to: 170, ldg: 193}
  ];

  // 2. INTERPOLATION
  let baseTO, baseLDG;
  for (let i = 0; i < perfData.length - 1; i++) {
    if (weight >= perfData[i].w && weight <= perfData[i+1].w) {
      let ratio = (weight - perfData[i].w) / (perfData[i+1].w - perfData[i].w);
      baseTO = perfData[i].to + ratio * (perfData[i+1].to - perfData[i].to);
      baseLDG = perfData[i].ldg + ratio * (perfData[i+1].ldg - perfData[i].ldg);
      break;
    }
  }

  // 3. APPLY PRECISE MULTIPLIERS FROM CHART
  let calcTO = baseTO;
  let calcLDG = baseLDG;

  // Flaps (Updated: 10 degrees is 1.3x)
  if (factors.flaps === 10) calcTO *= 1.30;
  if (factors.flaps === 20) calcTO *= 1.20;

  // Density Altitude (Lookup based on your new chart)
  calcTO *= getDAMultiplier(factors.da, "TO");
  calcLDG *= getDAMultiplier(factors.da, "LDG");

  // Surface (Lookup Dry/Wet categories)
  calcTO *= getSurfaceMultiplier(factors.surface, factors.isWet, "TO");
  calcLDG *= getSurfaceMultiplier(factors.surface, factors.isWet, "LDG");

  // Wind & Slope
  calcTO *= getWindMultiplier(factors.windComponent, "TO");
  calcLDG *= getWindMultiplier(factors.windComponent, "LDG");
  calcTO *= getSlopeMultiplier(factors.slope, "TO");
  calcLDG *= getSlopeMultiplier(factors.slope, "LDG");

  // 4. SAFETY & ABORT
  let asd = calcTO + calcLDG;
  let decisionPointDist = calcTO / 2;
  let runwayLimit = factors.runwayLength * 0.75;

  return {
    takeoffRoll: Math.round(calcTO),
    landingRoll: Math.round(calcLDG),
    accelStopDist: Math.round(asd),
    decisionPoint: Math.round(decisionPointDist),
    isSafe: (calcTO <= runwayLimit) && (calcLDG <= runwayLimit),
    runwayLimit: Math.round(runwayLimit)
  };
}

// Helper for DA Lookup
function getDAMultiplier(da, phase) {
  if (da <= 1000) return (phase === "TO") ? 1.03 : 1.04;
  if (da <= 2000) return (phase === "TO") ? 1.10 : 1.07;
  if (da <= 3000) return (phase === "TO") ? 1.23 : 1.11;
  if (da <= 4000) return (phase === "TO") ? 1.40 : 1.14;
  return 1.0; // Default fallback
}

function buildNewAviationDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup REF_Lists for Validated Surfaces
  let refSheet = ss.getSheetByName("REF_Lists");
  if (!refSheet) { refSheet = ss.insertSheet("REF_Lists"); }
  refSheet.clear();
  
  const surfaceTypes = [
    ["Surface_Types"],
    ["Paved"],
    ["Short Grass"],
    ["Long Grass (6\"+)"],
    ["Hard Turf"],
    ["Sand / Soft"],
    ["Mud / Marsh"]
  ];
  refSheet.getRange(1, 1, surfaceTypes.length, 1).setValues(surfaceTypes);
  refSheet.getRange(1, 1).setBackground("#38761d").setFontColor("white").setFontWeight("bold");

  // 2. Setup New Directional DB_Airports
  let airportSheet = ss.getSheetByName("DB_Airports");
  if (!airportSheet) { airportSheet = ss.insertSheet("DB_Airports"); }
  airportSheet.clear();
  
  const headers = [
    "ICAO_ID", "Runway_Designator", "Length_m", "Slope_Pct", "Surface_Type", 
    "Elevation_ft", "Runway_Heading", "MTOW_Restriction", "MLW_Restriction"
  ];
  
  airportSheet.getRange(1, 1, 1, headers.length)
              .setValues([headers])
              .setBackground("#073763")
              .setFontColor("white")
              .setFontWeight("bold");

  // 3. Add Data Validation (Dropdowns) for Surface Type
  const surfaceRange = refSheet.getRange("A2:A" + surfaceTypes.length);
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(surfaceRange).build();
  airportSheet.getRange(2, 5, 500, 1).setDataValidation(rule);

  // 4. Add Sample Data for AFBR (Two Directions)
  const sampleData = [
    ["AFBR", "07", 850, 1.5, "Short Grass", 2400, 070, 1723, 1723],
    ["AFBR", "25", 850, -1.5, "Short Grass", 2400, 250, 1723, 1723]
  ];
  airportSheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);

  // Freeze top row for easy scrolling
  airportSheet.setFrozenRows(1);
  
  Browser.msgBox("New Aviation Database Structure Created! Surface validation is active.");
}
function buildMultiPaxDispatch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Headers for 9 individual passengers
  let headers = [
    "Mission_ID", "Flight_ID", "Status", "Date", "ETD", "Aircraft", "Pilot", "Student", "From", "To"
  ];
  
  // Create 9 sets of passenger columns
  for (let i = 1; i <= 9; i++) {
    headers.push(`Pax${i}_Name`, `Pax${i}_Weight`, `Pax${i}_Dest`, `Pax${i}_Fund`, `Pax${i}_Rate`);
  }
  
  headers.push("Cargo_Weight", "Total_Payload", "Leg_Fuel_Req", "Cumulative_Fuel", "Max_Allowable_TOW", "Weight_Margin");

  let sheet = ss.getSheetByName("DB_Dispatch");
  if (!sheet) { sheet = ss.insertSheet("DB_Dispatch"); }
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length)
       .setValues([headers])
       .setBackground("#073763").setFontColor("white").setFontWeight("bold");
  
  // Apply Data Validation for names on all 5 Pax Name columns
  const paxRange = ss.getSheetByName("DB_Passengers").getRange("A2:A1000");
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(paxRange).build();
  
  // Column indices for Pax Names: 11, 16, 21, 26, 31, 36, 41, 46, 51
  [11, 16, 21, 26, 31, 36, 41, 46, 51].forEach(col => {
    sheet.getRange(2, col, 500, 1).setDataValidation(rule);
  });

  Browser.msgBox("Multi-Passenger Dispatch Ready with Split Billing support!");
}
function upgradeMaintenanceSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Update existing DB_Aircraft
  let dashSheet = ss.getSheetByName("DB_Aircraft");
  if (dashSheet) {
    const lastCol = dashSheet.getLastColumn();
    const headersToAdd = ["Current Tach", "Next Due (Tach)", "Hours to Insp", "Annual Due (CVA)", "Open Squawks", "Tech Status"];
    
    // Append headers to the first empty columns
    dashSheet.getRange(1, lastCol + 1, 1, headersToAdd.length)
      .setValues([headersToAdd])
      .setBackground("#27ae60").setFontColor("white").setFontWeight("bold");
  } else {
    SpreadsheetApp.getUi().alert("Error: DB_Aircraft tab not found. Please check the name.");
    return;
  }

  // 2. Create DB_Component_Matrix (The "Future" view)
  let matrixSheet = ss.getSheetByName("DB_Component_Matrix");
  if (!matrixSheet) {
    matrixSheet = ss.insertSheet("DB_Component_Matrix");
  }

  const matrixHeaders = [
    ["Tail #", "Item / AD / Inspection", "Interval (Hrs)", "Last Performed (Tach)", "Next Due (Tach)", "Remaining (Hrs)", "Date Last Done", "Date Due"]
  ];
  matrixSheet.getRange(1, 1, 1, matrixHeaders[0].length).setValues(matrixHeaders)
    .setBackground("#2c3e50").setFontColor("white").setFontWeight("bold");
  matrixSheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert("DB_Aircraft upgraded and Component Matrix created.");
}
function setupChecksDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "DB_Checks";
  const aircraftTypes = ["Cessna 172S", "Cessna U206/520", "Cessna U206/550", "Cessna U206 Amphib"];
  const roles = ["Operational", "Instructor", "Student"];
  
  // 1. Create or Clear the sheet
  let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sheet.clear();
  
  // 2. Define Headers
  const headers = [
    "Pilot_Email", 
    "Aircraft_Type", 
    "Role", 
    "Date_of_Check", 
    "Expiry_Date", 
    "Instructor_Signoff", 
    "Notes"
  ];
  
  // 3. Apply Styling to Headers
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground("#2c3e50")
    .setFontColor("white")
    .setFontWeight("bold");

  // 4. Set Column Formats (Date columns)
  sheet.getRange("D2:E").setNumberFormat("dd/mm/yyyy");

  // 5. Add Dropdowns (Data Validation)
  // Aircraft Type Dropdown
  const acRule = SpreadsheetApp.newDataValidation().requireValueInList(aircraftTypes).build();
  sheet.getRange("B2:B500").setDataValidation(acRule);
  
  // Role Dropdown
  const roleRule = SpreadsheetApp.newDataValidation().requireValueInList(roles).build();
  sheet.getRange("C2:C500").setDataValidation(roleRule);

  // 6. Freeze Header Row
  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert("DB_Checks created. Use this to log annual checks for each aircraft type.");
}
function buildDispatchInfrastructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. BUILD DB_FUNDS
  let fundsSheet = ss.getSheetByName("DB_Funds") || ss.insertSheet("DB_Funds");
  fundsSheet.clear();
  const funds = [
    ["Fund Name", "Category", "Notes"],
    ["Dando Asas a Palavra", "Ministry", "Main Ministry"],
    ["Pontes Aereas de Amor", "Ministry", ""],
    ["Missao Capacitar", "Ministry", ""],
    ["Avançando o Reino", "Ministry", ""],
    ["Training", "Internal", ""],
    ["Maintenance", "Internal", ""],
    ["Ferry", "Internal", ""]
  ];
  fundsSheet.getRange(1, 1, funds.length, 3).setValues(funds).setFontWeight("bold");
  fundsSheet.getRange("A1:C1").setBackground("#4a86e8").setFontColor("white");

  // 2. BUILD ENTRY_MASK
  let maskSheet = ss.getSheetByName("ENTRY_MASK") || ss.insertSheet("ENTRY_MASK");
  maskSheet.clear();
  
  const maskLayout = [
    ["FLIGHT DISPATCH MASK", ""],
    ["Flight_Type", ""],        // B2
    ["Training_Code", ""],      // B3
    ["Date", ""],               // B4
    ["ETD", ""],                // B5
    ["Aircraft", ""],           // B6
    ["Pilot_CFI", ""],          // B7
    ["CoPilot_Student", ""],    // B8
    ["From", ""],               // B9
    ["To", ""],                 // B10
    ["Flight_Time (Hrs)", ""],  // B11
    ["Proposed_Landings", ""],  // B12
    ["Pax1_Name", ""],          // B13
    ["Pax1_Baggage (kg)", ""],  // B14
    ["Pax1_Dest", ""],          // B15
    ["Pax1_Fund", ""],          // B16
    ["Pax1_Rate", ""]           // B17
  ];
  
  maskSheet.getRange(1, 1, maskLayout.length, 2).setValues(maskLayout);
  
  // Formatting the UI
  maskSheet.setColumnWidth(1, 180);
  maskSheet.setColumnWidth(2, 250);
  maskSheet.getRange("A1:B1").merge().setBackground("#0b5394").setFontColor("white").setHorizontalAlignment("center");
  maskSheet.getRange("A2:A17").setBackground("#f3f3f3").setFontWeight("bold");
  maskSheet.getRange("B2:B17").setBackground("#ffffff").setBorder(true, true, true, true, true, true);
  
  // 3. APPLY DATA VALIDATION (Dropdowns)
  // Flight Type
  const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(["MX", "Training", "Check", "Missions", "Air Taxi", "Ferry"], true).build();
  maskSheet.getRange("B2").setDataValidation(typeRule);

  // Funds Dropdown
  const fundRule = SpreadsheetApp.newDataValidation().requireValueInRange(fundsSheet.getRange("A2:A20")).build();
  maskSheet.getRange("B16").setDataValidation(fundRule);

  // Placeholder reminders for the coordinator
  maskSheet.getRange("B14").setValue(15); 

  SpreadsheetApp.getUi().alert("Infrastructure Built! Remember to link dropdowns for Aircraft, Pilots, and Airports manually or via script.");
}
/**
 * Creates the necessary Fuel Management tabs if they don't exist.
 * Run this function once from the script editor.
 */
function setupFuelSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup DB_Fuel_Caches (Current Inventory)
  let cacheSheet = ss.getSheetByName("DB_Fuel_Caches");
  if (!cacheSheet) {
    cacheSheet = ss.insertSheet("DB_Fuel_Caches");
    const headers = [["ICAO", "LOCATION_NAME", "CURRENT_QTY", "FUEL_TYPE", "MIN_THRESHOLD"]];
    cacheSheet.getRange(1, 1, 1, 5).setValues(headers)
      .setBackground("#0b5394").setFontColor("white").setFontWeight("bold");
    
    // Add a sample row
    cacheSheet.appendRow(["M_01", "Jungle Outpost 1", 500, "Avgas", 100]);
    cacheSheet.setFrozenRows(1);
  }

  // 2. Setup DB_Fuel_Logs (Transaction History)
  let logSheet = ss.getSheetByName("DB_Fuel_Logs");
  if (!logSheet) {
    logSheet = ss.insertSheet("DB_Fuel_Logs");
    const headers = [["TIMESTAMP", "ICAO", "AIRCRAFT", "PILOT", "CHANGE_QTY", "TYPE", "VERIFIED"]];
    logSheet.getRange(1, 1, 1, 7).setValues(headers)
      .setBackground("#38761d").setFontColor("white").setFontWeight("bold");
    
    cacheSheet.setFrozenRows(1);
  }
  
  SpreadsheetApp.getUi().alert("Fuel System Tabs Created Successfully!");
}
function createDBPilotAuthorizations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'DB_Pilot_Authorizations';

  // Prevent overwriting if it already exists
  if (ss.getSheetByName(sheetName)) {
    SpreadsheetApp.getUi().alert(sheetName + ' already exists.');
    return;
  }

  // Create the sheet
  const sheet = ss.insertSheet(sheetName);

  // Define headers
  const headers = [
    'Pilot_Name',
    'Pilot_Email',
    'Aircraft_Type',
    'Role',
    'Authorization_Type',
    'Status',
    'Date_Authorized',
    'Expiry_Date',
    'Instructor_Signoff',
    'Source',
    'Notes'
  ];

  // Write headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Formatting
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.autoResizeColumns(1, headers.length);

  // Optional: data validation (recommended)
  addAuthorizationValidations_(sheet);

  SpreadsheetApp.getUi().alert(sheetName + ' created successfully.');
}
function addAuthorizationValidations_(sheet) {
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Student', 'Operational', 'Instructor'], true)
    .setAllowInvalid(false)
    .build();

  const authTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['INITIAL', 'ANNUAL_PROFICIENCY', 'UPGRADE', 'REINSTATEMENT'], true)
    .setAllowInvalid(false)
    .build();

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['ACTIVE', 'SUSPENDED', 'REVOKED'], true)
    .setAllowInvalid(false)
    .build();

  // Apply rules to reasonable column ranges
  sheet.getRange('D2:D').setDataValidation(roleRule);          // Role
  sheet.getRange('E2:E').setDataValidation(authTypeRule);     // Authorization_Type
  sheet.getRange('F2:F').setDataValidation(statusRule);       // Status
}

function _donationEnsureSheet_(ss, sheetName, headers, tabColor, overwriteExisting) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  const shouldReplace = !!overwriteExisting;
  const hasData = sheet.getLastRow() > 0 || sheet.getLastColumn() > 0;

  if (shouldReplace) {
    sheet.clear();
  }

  if (shouldReplace || !hasData) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const existingHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
    let mismatch = false;
    for (let i = 0; i < headers.length; i++) {
      if (String(existingHeaders[i] || '').trim() !== String(headers[i] || '').trim()) {
        mismatch = true;
        break;
      }
    }
    if (mismatch) {
      throw new Error('Sheet ' + sheetName + ' already exists with different headers. Re-run with overwriteExisting=true if you want to replace it.');
    }
  }

  sheet.getRange(1, 1, 1, headers.length)
    .setBackground(tabColor || '#0b5394')
    .setFontColor('white')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

function _ensureDonationImportsDriveFolder_() {
  const folderName = 'MBA_Donation_Imports';
  const props = PropertiesService.getScriptProperties();
  const existingId = String(props.getProperty('DONATION_IMPORTS_FOLDER_ID') || '').trim();

  if (existingId) {
    try {
      return DriveApp.getFolderById(existingId);
    } catch (e) {}
  }

  const folders = DriveApp.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  props.setProperty('DONATION_IMPORTS_FOLDER_ID', folder.getId());
  return folder;
}

function setupDonationFundingSystem(overwriteExisting, showAlert) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const definitions = [
    {
      name: 'DB_Donation_Import_Batches',
      color: '#1f4e78',
      headers: [
        'BATCH_ID',
        'SOURCE_FILENAME',
        'SOURCE_TYPE',
        'SOURCE_FILE_ID',
        'SOURCE_FILE_URL',
        'FILE_HASH',
        'IMPORTED_AT',
        'IMPORTED_BY',
        'STATEMENT_START',
        'STATEMENT_END',
        'ROW_COUNT',
        'STATUS',
        'NOTES'
      ]
    },
    {
      name: 'DB_Donation_Staging',
      color: '#3d85c6',
      headers: [
        'BATCH_ID',
        'ROW_NO',
        'DONOR_RAW',
        'DONOR_NORMALIZED',
        'TX_DATE',
        'AMOUNT_BRL',
        'FUND_ID',
        'DESCRIPTION_RAW',
        'FINGERPRINT_STRICT',
        'FINGERPRINT_FUZZY',
        'MATCH_STATUS',
        'MATCHED_DONATION_ID',
        'MATCH_CONFIDENCE',
        'REVIEW_DECISION',
        'REVIEWED_BY',
        'REVIEWED_AT',
        'NOTES'
      ]
    },
    {
      name: 'DB_Donations_Ledger',
      color: '#38761d',
      headers: [
        'DONATION_ID',
        'TX_DATE',
        'DONOR_NORMALIZED',
        'DONOR_RAW',
        'AMOUNT_BRL',
        'FUND_ID',
        'SOURCE_BATCH_ID',
        'SOURCE_ROW_NO',
        'SOURCE_FILE_ID',
        'DESCRIPTION_RAW',
        'FINGERPRINT_STRICT',
        'CREATED_AT',
        'CREATED_BY',
        'STATUS',
        'ROLLED_BACK_AT',
        'ROLLED_BACK_BY',
        'ROLLBACK_REASON'
      ]
    },
    {
      name: 'DB_Fund_Ledger',
      color: '#6aa84f',
      headers: [
        'LEDGER_ID',
        'FUND_ID',
        'ENTRY_DATE',
        'ENTRY_TYPE',
        'AMOUNT_BRL',
        'RELATED_ID',
        'RELATED_TYPE',
        'SOURCE_BATCH_ID',
        'MISSION_ID',
        'DONATION_ID',
        'NOTE',
        'CREATED_AT',
        'CREATED_BY'
      ]
    },
    {
      name: 'DB_Mission_Funding_Allocations',
      color: '#674ea7',
      headers: [
        'ALLOCATION_ID',
        'MISSION_ID',
        'FLIGHT_ID',
        'FUND_ID',
        'DONATION_ID',
        'ALLOC_AMOUNT_BRL',
        'ALLOCATION_STAGE',
        'RESERVED_AT',
        'FINALIZED_AT',
        'RELEASED_AT',
        'CREATED_BY',
        'NOTES'
      ]
    }
  ];

  definitions.forEach(function(def) {
    _donationEnsureSheet_(ss, def.name, def.headers, def.color, overwriteExisting);
  });

  const folder = _ensureDonationImportsDriveFolder_();

  const result = {
    success: true,
    sheets: definitions.map(function(def) { return def.name; }),
    folderId: folder.getId(),
    folderUrl: folder.getUrl()
  };

  const shouldAlert = (showAlert !== false);
  if (shouldAlert) {
    SpreadsheetApp.getUi().alert(
      'Donation funding schema ready.\n\n'
      + 'Sheets created/validated:\n'
      + definitions.map(function(def) { return '- ' + def.name; }).join('\n')
      + '\n\nDrive folder for uploaded source files:\n'
      + folder.getUrl()
    );
  }

  return result;
}