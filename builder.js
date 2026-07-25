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



/**
 * Creates the necessary Fuel Management tabs if they don't exist.
 * Run this function once from the script editor.
 */
function addAuthorizationValidations_(sheet) {
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Student', 'Operational', 'Instructor'], true)
    .setAllowInvalid(false)
    .build();

  const authTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['INITIAL', 'ANNUAL_PROFICIENCY', 'UPGRADE', 'REINSTATEMENT', 'LINE_CHECK'], true)
    .setAllowInvalid(false)
    .build();

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['ACTIVE', 'PENDING', 'SUSPENDED', 'REVOKED', 'EXPIRED'], true)
    .setAllowInvalid(false)
    .build();

  // Apply rules to reasonable column ranges
  sheet.getRange('F2:F').setDataValidation(roleRule);          // Role
  sheet.getRange('G2:G').setDataValidation(authTypeRule);     // Authorization_Type
  sheet.getRange('H2:H').setDataValidation(statusRule);       // Status
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