/* ==================================================
1. MAIN PORTAL CONTROLLER
================================================== */
function doGet(e) {
 const view = (e && e.parameter && e.parameter.view ? e.parameter.view : "").toLowerCase();
 const pilotParamRaw = (e && e.parameter && e.parameter.pilot != null) ? String(e.parameter.pilot).toLowerCase().trim() : "";
 const pilotParamTrue = pilotParamRaw === "1" || pilotParamRaw === "true" || pilotParamRaw === "yes" || pilotParamRaw === "y";
 const isPilot = view === "pilot" || view === "flightdeck" || view === "pilotapp" || pilotParamTrue;
 const isPortal = !isPilot;


 const fileName = isPilot ? 'PilotApp' : 'Index';
 const title = isPilot ? 'Pilot Flight Deck' : 'Flight Ops Portal';

 const template = HtmlService.createTemplateFromFile(fileName);
 template.webAppUrl = ScriptApp.getService().getUrl();

 return template.evaluate()
   .setTitle(title)
   .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
   .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}






function include(filename) {
return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getInclinometerStandaloneHtml() {
  return HtmlService.createHtmlOutputFromFile('inclinometer_hud_standalone').getContent();
}

function getRequiredSheet_(ss, sheetName, contextLabel) {
const sheet = ss.getSheetByName(sheetName);
if (!sheet) {
  throw new Error((contextLabel || 'Operation') + ': missing required sheet "' + sheetName + '"');
}
return sheet;
}

function mapAirportRowsShared_(rowsObj) {
const rows = rowsObj && Array.isArray(rowsObj.vals) ? rowsObj.vals : [];
const headers = rowsObj && Array.isArray(rowsObj.headers) ? rowsObj.headers : [];
return rows.map(r => {
  const byHeader = (name, fallback) => {
    const idx = headers.indexOf(name);
    return idx >= 0 ? r[idx] : fallback;
  };
  const byAnyHeader = (names, fallback) => {
    const list = Array.isArray(names) ? names : [names];
    for (var ni = 0; ni < list.length; ni++) {
      const value = byHeader(list[ni], null);
      if (value !== null && value !== undefined && String(value).trim() !== '') return value;
    }
    return fallback;
  };
  return {
    icao: byHeader("ICAO", ""),
    nome: byHeader("NOME", ""),
    lat: byHeader("LATITUDE", ""),
    lon: byHeader("LONGITUDE", ""),
    fuelAvailable: byHeader("FUEL_AVAILABLE", ""),
    mtow520: parseFloat(byHeader("MTOW_LIMIT_206_520", 9999)) || 9999,
    mtow550: parseFloat(byHeader("MTOW_LIMIT_206_550", 9999)) || 9999,
    pilotNotes: String(byHeader("PILOT_NOTES", "") || ""),
    airstripPhoto: String(byHeader("AIRSTRIP_PHOTO", "") || ""),
    runwayIdent: byHeader("RWY_IDENT", byHeader("RWY", byHeader("RUNWAY", byHeader("RUNWAY_DESIGNATOR", "")))),
    runwayHeading: byHeader("RUNWAY_HEADING", byHeader("HEADING", "")),
    runwayLength: byHeader("LENGTH_OFFICIAL", byHeader("LENGTH_METERS", byHeader("LENGTH_M", ""))),
    runwayWidth: byHeader("WIDTH_OFFICIAL", byHeader("WIDTH_METERS", byHeader("WIDTH_M", ""))),
    runwaySlopePercent: byHeader("SLOPE_PERCENT", byHeader("SLOPE_PCT", "")),
    runwaySlopeProfile: byHeader("SLOPE_PROFILE", byHeader("RUNWAY_SLOPE_PROFILE", "")),
    elevationFt: byHeader("ELEVATION", byHeader("ALT_FEET", byHeader("ELEVATION_FT", ""))),
    runwaySurfaceActual: byAnyHeader(["SURFACE_ACTUAL", "RUNWAY_SURFACE_ACTUAL", "SURFACE_OFFICIAL", "RUNWAY_SURFACE", "SURFACE_TYPE", "SURFACE"], ""),
    runwaySurfaceCondition: byAnyHeader(["SURFACE_CONDITION", "RUNWAY_SURFACE_CONDITION", "CONDITION", "SURFACE_STATUS"], ""),
    chartUrl: byHeader("CHART_URL", byHeader("PLATE_URL", byHeader("APPROACH_CHART", byHeader("PROCEDURE_PDF", byHeader("PDF_URL", ""))))),
    knownFeatures: byHeader("KNOWN_FEATURES", byHeader("FEATURES", ""))
  };
});
}




/* ==================================================
FIXED: DROPDOWN DATA (PURE JS - NO AUTH CRASH)
================================================== */








function getDropdownData() {
const ss = SpreadsheetApp.getActiveSpreadsheet();
getRequiredSheet_(ss, APP_SHEETS.DISPATCH, "getDropdownData");
















// Helper: Pure JS Date Formatter (No 'Session' calls)
const safeDateStr = (val) => {
if (!val) return "";
// If it's already a date object, convert it
if (val instanceof Date) {
 return val.toISOString().split('T')[0]; // "YYYY-MM-DD"
}
// If it's a string, try to keep it or clean it
return String(val).split('T')[0];
};

const safeDobStr = (val) => {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT", "yyyy-MM-dd");
  }
  const raw = String(val).trim();
  if (!raw) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  if (/^\d{4}-\d{2}-\d{2}T/.test(raw)) return raw.slice(0, 10);

  // Treat slash format as dd/mm/yyyy for Brazilian data consistency.
  const slash = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slash) {
    const dd = String(parseInt(slash[1], 10)).padStart(2, '0');
    const mm = String(parseInt(slash[2], 10)).padStart(2, '0');
    const yyyy = slash[3];
    return `${yyyy}-${mm}-${dd}`;
  }

  const parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, "GMT", "yyyy-MM-dd");
  }
  return "";
};
function getSheetData(name) {
const sheet = ss.getSheetByName(name);
if (!sheet) return { vals: [], headers: [] };
const vals = sheet.getDataRange().getValues();
if (vals.length < 1) return { vals: [], headers: [] };
// Normalize headers
const headers = vals[0].map(h => h.toString().toUpperCase().trim().replace(/\s/g, '_'));
return { vals: vals.slice(1), headers };
}
function mapAirportRows(rowsObj) {
const rows = rowsObj && Array.isArray(rowsObj.vals) ? rowsObj.vals : [];
const headers = rowsObj && Array.isArray(rowsObj.headers) ? rowsObj.headers : [];
return rows.map(r => {
  const byHeader = (name, fallback) => {
    const idx = headers.indexOf(name);
    return idx >= 0 ? r[idx] : fallback;
  };
  const byAnyHeader = (names, fallback) => {
    const list = Array.isArray(names) ? names : [names];
    for (var ni = 0; ni < list.length; ni++) {
      const value = byHeader(list[ni], null);
      if (value !== null && value !== undefined && String(value).trim() !== '') return value;
    }
    return fallback;
  };
  return {
    icao: byHeader("ICAO", ""),
    nome: byHeader("NOME", ""),
    lat: byHeader("LATITUDE", ""),
    lon: byHeader("LONGITUDE", ""),
    fuelAvailable: byHeader("FUEL_AVAILABLE", ""),
    mtow520: parseFloat(byHeader("MTOW_LIMIT_206_520", 9999)) || 9999,
    mtow550: parseFloat(byHeader("MTOW_LIMIT_206_550", 9999)) || 9999,
    pilotNotes: String(byHeader("PILOT_NOTES", "") || ""),
    airstripPhoto: String(byHeader("AIRSTRIP_PHOTO", "") || ""),
    runwayIdent: byHeader("RWY_IDENT", byHeader("RWY", byHeader("RUNWAY", byHeader("RUNWAY_DESIGNATOR", "")))),
    runwayHeading: byHeader("RUNWAY_HEADING", byHeader("HEADING", "")),
    runwayLength: byHeader("LENGTH_OFFICIAL", byHeader("LENGTH_METERS", byHeader("LENGTH_M", ""))),
    runwayWidth: byHeader("WIDTH_OFFICIAL", byHeader("WIDTH_METERS", byHeader("WIDTH_M", ""))),
    runwaySlopePercent: byHeader("SLOPE_PERCENT", byHeader("SLOPE_PCT", "")),
    runwaySlopeProfile: byHeader("SLOPE_PROFILE", byHeader("RUNWAY_SLOPE_PROFILE", "")),
    elevationFt: byHeader("ELEVATION", byHeader("ALT_FEET", byHeader("ELEVATION_FT", ""))),
    runwaySurfaceActual: byAnyHeader(["SURFACE_ACTUAL", "RUNWAY_SURFACE_ACTUAL", "SURFACE_OFFICIAL", "RUNWAY_SURFACE", "SURFACE_TYPE", "SURFACE"], ""),
    runwaySurfaceCondition: byAnyHeader(["SURFACE_CONDITION", "RUNWAY_SURFACE_CONDITION", "CONDITION", "SURFACE_STATUS"], ""),
    chartUrl: byHeader("CHART_URL", byHeader("PLATE_URL", byHeader("APPROACH_CHART", byHeader("PROCEDURE_PDF", byHeader("PDF_URL", ""))))),
    knownFeatures: byHeader("KNOWN_FEATURES", byHeader("FEATURES", ""))
  };
});
}
try {
const air = getSheetData(APP_SHEETS.AIRPORTS);
const acft = getSheetData(APP_SHEETS.AIRCRAFT);
const pax = getSheetData(APP_SHEETS.PASSENGERS);
const syl = getSheetData(APP_SHEETS.SYLLABUS);
const caches = getSheetData(APP_SHEETS.FUEL_CACHES);
const wpData = getSheetData(APP_SHEETS.WAYPOINTS);
const routeData = getSheetData(APP_SHEETS.ROUTES);
const pilots = getSheetData(APP_SHEETS.PILOTS);
const checks = getSheetData(APP_SHEETS.CHECKS);

// Diagnostics: log presence and header info for key sheets to help trace missing data
try {
  appLog_('getDropdownData: DB_Aircraft headers=', acft && acft.headers ? acft.headers : 'N/A', 'rows=', acft && acft.vals ? acft.vals.length : 0);
  appLog_('getDropdownData: DB_Fuel_Caches headers=', caches && caches.headers ? caches.headers : 'N/A', 'rows=', caches && caches.vals ? caches.vals.length : 0);
  appLog_('getDropdownData: DB_Pilots headers=', pilots && pilots.headers ? pilots.headers : 'N/A', 'rows=', pilots && pilots.vals ? pilots.vals.length : 0);
} catch (e) { appLog_('getDropdownData diagnostics error:', e && e.message); }








// PILOTS MAPPING
const pilotNameIdx = pilots.headers.indexOf("NAME") !== -1 ? pilots.headers.indexOf("NAME") : pilots.headers.indexOf("PILOT_NAME");
const pilotWeightIdx = pilots.headers.indexOf("WEIGHT_KGS");
const medicalIdx = pilots.headers.indexOf("MEDICAL_EXPIRY");
const mnteIdx = pilots.headers.indexOf("MNTE_VALIDITY");
const mnafIdx = pilots.headers.indexOf("MNAF_VALIDITY");








const pilotList = pilots.vals.map(r => ({
   name: r[pilotNameIdx],
   weight: pilotWeightIdx !== -1 ? parseFloat(r[pilotWeightIdx]) || 90 : 90,
   Medical_Expiry: medicalIdx !== -1 ? safeDateStr(r[medicalIdx]) : "",
   MNTE_Validity: mnteIdx !== -1 ? safeDateStr(r[mnteIdx]) : "",
   MNAF_Validity: mnafIdx !== -1 ? safeDateStr(r[mnafIdx]) : ""
})).filter(p => p.name);

// PILOT DESTINATION AUTHORIZATION MAPPING (from DB_Checks)
const checksPilotIdx = checks.headers.indexOf("PILOT");
const checksDestIdx = (function() {
  const candidates = [
    "AUTH_DESTINATIONS",
    "AUTHORIZED_DESTINATIONS",
    "DESTINATIONS",
    "DESTINATION",
    "AUTH_DESTINATION"
  ];
  for (let i = 0; i < candidates.length; i++) {
    const idx = checks.headers.indexOf(candidates[i]);
    if (idx !== -1) return idx;
  }
  return -1;
})();

const pilotDestinationChecks = {};
if (checksPilotIdx !== -1 && checksDestIdx !== -1) {
  checks.vals.forEach(r => {
    const pilotName = String(r[checksPilotIdx] || '').trim();
    if (!pilotName) return;
    const key = pilotName.toUpperCase();
    const raw = String(r[checksDestIdx] || '');
    const list = raw
      .split(/[;,]/)
      .map(s => String(s || '').trim().toUpperCase())
      .map(s => s.replace(/[^A-Z0-9]/g, ''))
      .filter(Boolean);
    if (!pilotDestinationChecks[key]) pilotDestinationChecks[key] = [];
    list.forEach(icao => {
      if (pilotDestinationChecks[key].indexOf(icao) === -1) pilotDestinationChecks[key].push(icao);
    });
  });
}








// FUNDS MAPPING
const fundSheet = ss.getSheetByName(APP_SHEETS.FUNDS);
const fundRange = fundSheet ? fundSheet.getDataRange().getValues() : [];
let funds = [];
if(fundRange.length > 1) {
  const fHead = fundRange[0].map(h => h.toString().toUpperCase().trim());
  let nameIdx = fHead.indexOf("NAME");
  if (nameIdx === -1) nameIdx = fHead.indexOf("FUND NAME");
  const balIdx = fHead.indexOf("CURRENT_BALANCE");
  funds = fundRange.slice(1).map(r => {
    const nm = r[nameIdx];
    if(!nm) return null;
    const bal = parseFloat(r[balIdx]) || 0;
    return {
      id: nm,
      displayName: `${nm} (R$ ${bal.toLocaleString('en-US', {minimumFractionDigits: 2})})`,
      balance: bal, limit: 0
    };
  }).filter(f=>f);
}
return {
 nextId: getNextMissionId(),
 pilots: pilotList,
 pilotDestinationChecks: pilotDestinationChecks,
 funds: funds,
 rates: ["1/5", "2/5", "1/2", "3/5", "4/5", "1/1"],
 fuelCaches: caches.vals.map(r => ({
   icao: r[caches.headers.indexOf("ICAO")],
   location: r[caches.headers.indexOf("LOCATION_NAME")],
   qty: parseFloat(r[caches.headers.indexOf("CURRENT_QTY")]) || 0,
   type: r[caches.headers.indexOf("FUEL_TYPE")]
 })),
 syllabus: syl.vals.map(r => ({
     code: r[syl.headers.indexOf("TRAINING_CODE")],
     hours: parseFloat(r[syl.headers.indexOf("REQUIRED_HOURS")]) || 0,
     fuel: parseFloat(r[syl.headers.indexOf("REQUIRED_FUEL")]) || 0,
     route: (function() {
       var routeIdx = syl.headers.indexOf('ROUTE');
       return routeIdx >= 0 ? String(r[routeIdx] || '').trim() : '';
     })()
 })).filter(s => s.code),
 waypoints: wpData.vals.map(r => ({
   wp_id: String(r[wpData.headers.indexOf("WP_ID")]),
   lat: parseFloat(r[wpData.headers.indexOf("LATITUDE")]),
   lon: parseFloat(r[wpData.headers.indexOf("LONGITUDE")]),
   type: String(r[wpData.headers.indexOf("TYPE")] || "")
 })).filter(w => w.wp_id),
































routes: routeData.vals.map((r, idx) => ({
  rowNumber: idx + 2,
  origin: String(r[routeData.headers.indexOf("ORIGIN")] || '').trim().toUpperCase(),
  destination: String(r[routeData.headers.indexOf("DESTINATION")] || '').trim().toUpperCase(),
  waypoint_list: String(r[routeData.headers.indexOf("WAYPOINT_LIST")] || '').trim()
})).filter(rt => rt.origin),
































 aircraft: acft.vals.map(r => {
   const vrCols = {};
   const _idxByAliases = function(aliases) {
     for (var i = 0; i < aliases.length; i++) {
       var idx = acft.headers.indexOf(aliases[i]);
       if (idx >= 0) return idx;
     }
     return -1;
   };
   const _numByAliases = function(aliases, fallback) {
     var idx = _idxByAliases(aliases);
     if (idx < 0) return fallback;
     var val = parseFloat(r[idx]);
     return isNaN(val) ? fallback : val;
   };
   acft.headers.forEach(function(h, idx) {
     const normKey = String(h || '').toUpperCase().trim().replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
     if (!normKey) return;
     if (normKey.indexOf('VR') >= 0 || normKey.indexOf('ROTATE') >= 0) {
       vrCols[normKey] = r[idx];
     }
   });
   return Object.assign({
     reg: r[acft.headers.indexOf("REGISTRATION")],
     aircraftType: r[acft.headers.indexOf("AIRCRAFT_TYPE")] || "",
     typeForPerformance: r[acft.headers.indexOf("TYPE_FOR_PERFORMANCE")] || "",
     speed: parseFloat(r[acft.headers.indexOf("CRUISE_KTS")]) || 130,
     burn: parseFloat(r[acft.headers.indexOf("BURN_LPH")]) || 60,
     rate: parseFloat(r[acft.headers.indexOf("HOURLY_RATE")]) || 0,
     maxFuel: parseFloat(r[acft.headers.indexOf("MAX_FUEL")]) || 400,
     emptyWeight: parseFloat(r[acft.headers.indexOf("EMPTY_WEIGHT")]) || 1000,
     mtow: parseFloat(r[acft.headers.indexOf("MTOW")]) || 1600,
     License_Required: r[acft.headers.indexOf("LICENSE_REQUIRED")] || "MNTE",
    pilotSeat: _numByAliases(["PILOT_SEAT_KGS", "PILOT_SEAT_KG", "PILOT_SEAT_WEIGHT", "PILOT_SEAT"], null),
    midSeat: _numByAliases(["MID_SEAT_KGS", "MID_SEAT_KG", "MID_SEAT_WEIGHT", "MIDDLE_SEAT_KGS", "MIDDLE_SEAT_KG"], null),
    aftSeat: _numByAliases(["AFT_SEAT_KGS", "AFT_SEAT_KG", "AFT_SEAT_WEIGHT", "REAR_SEAT_KGS", "REAR_SEAT_KG"], null),
     NUM_TANKS: r[acft.headers.indexOf("NUM_TANKS")] || 0,
     TANK_NAMES: r[acft.headers.indexOf("TANK_NAMES")] || "",
     MAIN_CAPACITY_L: r[acft.headers.indexOf("MAIN_CAPACITY_L")] || 0,
     TIP_CAPACITY_L: r[acft.headers.indexOf("TIP_CAPACITY_L")] || 0,
     TRANSFER_RATE_LPM: (function(){
       const idx = acft.headers.indexOf("TRANSFER_RATE_LPM");
       return idx >= 0 ? (parseFloat(r[idx]) || 0) : 0;
     })(),
     currentTach: parseFloat(r[acft.headers.indexOf("CURRENT_TACH")]) || 0,
     nextDue: parseFloat(r[acft.headers.indexOf("NEXT_DUE_TACH")]) || 0,
     hoursToInsp: parseFloat(r[acft.headers.indexOf("HOURS_TO_INSPECTION")]) || 0,
      techStatus: (r[acft.headers.indexOf("TECH_STATUS")] || "SERVICEABLE").toUpperCase(),
     openSquawks: r[acft.headers.indexOf("OPEN_SQUAWKS")] || ""
   }, vrCols);
 }),
 passengers: pax.vals.map(r => {
     const h = pax.headers;
     const weightIdx = h.indexOf("WEIGHT_KG") !== -1 ? h.indexOf("WEIGHT_KG") : h.indexOf("WEIGHT_KGS");
     const dobIdx = h.indexOf("DOB");
     const phoneIdx = h.indexOf("PHONE");
     return {
       name: r[h.indexOf("PASSENGER_NAME")] || "Unknown",
       weight: parseFloat(r[weightIdx]) || 80,
       gender: r[h.indexOf("GENDER")] || "U",
       dob: safeDobStr(dobIdx !== -1 ? r[dobIdx] : ""),
       phone: phoneIdx !== -1 ? String(r[phoneIdx] || "") : ""
     };
 }).filter(p => p.name && p.name !== "Unknown"),
 airports: mapAirportRows(air)
};
} catch (e) {
// Return a safe error object instead of crashing
console.log("Dropdown Init Error: " + e.message);
return { error: e.message };
}
}

function getPilotStartupData() {
  var data = getDropdownData();
  if (data && data.error) return data;
  return {
    nextId: data && data.nextId || '',
    pilots: data && Array.isArray(data.pilots) ? data.pilots : [],
    pilotDestinationChecks: data && data.pilotDestinationChecks ? data.pilotDestinationChecks : {},
    funds: data && Array.isArray(data.funds) ? data.funds : [],
    rates: data && Array.isArray(data.rates) ? data.rates : [],
    fuelCaches: data && Array.isArray(data.fuelCaches) ? data.fuelCaches : [],
    syllabus: data && Array.isArray(data.syllabus) ? data.syllabus : [],
    waypoints: data && Array.isArray(data.waypoints) ? data.waypoints : [],
    routes: data && Array.isArray(data.routes) ? data.routes : [],
    airports: data && Array.isArray(data.airports) ? data.airports : [],
    aircraft: data && Array.isArray(data.aircraft) ? data.aircraft : [],
    passengers: data && Array.isArray(data.passengers) ? data.passengers : []
  };
}

function getPilotAirportData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(APP_SHEETS.AIRPORTS);
    if (!sheet) return { airports: [] };
    var vals = sheet.getDataRange().getValues();
    if (!vals || vals.length < 2) return { airports: [] };
    var headers = vals[0].map(function(h) {
      return h.toString().toUpperCase().trim().replace(/\s/g, '_');
    });
    var mapped = mapAirportRowsShared_({ vals: vals.slice(1), headers: headers });
    var asShortText = function(value, maxLen) {
      var text = String(value == null ? '' : value);
      var limit = Number(maxLen || 0) || 120;
      if (text.length <= limit) return text;
      return text.slice(0, limit);
    };
    return {
      airports: mapped.map(function(row) {
        return {
          icao: asShortText(row.icao, 12),
          nome: asShortText(row.nome, 90),
          lat: row.lat,
          lon: row.lon,
          fuelAvailable: asShortText(row.fuelAvailable, 24),
          mtow520: parseFloat(row.mtow520) || 9999,
          mtow550: parseFloat(row.mtow550) || 9999,
          runwayIdent: asShortText(row.runwayIdent, 20),
          runwayHeading: asShortText(row.runwayHeading, 12),
          runwayLength: asShortText(row.runwayLength, 12),
          runwayWidth: asShortText(row.runwayWidth, 12),
          runwaySurfaceActual: asShortText(row.runwaySurfaceActual, 40)
        };
      })
    };
  } catch (e) {
    return { error: e && e.message ? e.message : String(e) };
  }
}

function getPilotAirportDataChunk(offset, limit) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(APP_SHEETS.AIRPORTS);
    if (!sheet) return { airports: [], nextOffset: 0, done: true, total: 0 };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { airports: [], nextOffset: 0, done: true, total: 0 };

    var total = Math.max(0, lastRow - 1);
    var startOffset = Math.max(0, parseInt(offset, 10) || 0);
    var chunkSize = Math.max(100, Math.min(parseInt(limit, 10) || 1200, 2000));
    if (startOffset >= total) {
      return { airports: [], nextOffset: startOffset, done: true, total: total };
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) {
      return h.toString().toUpperCase().trim().replace(/\s/g, '_');
    });

    var rowsToFetch = Math.min(chunkSize, total - startOffset);
    var vals = sheet.getRange(2 + startOffset, 1, rowsToFetch, headers.length).getValues();
    var mapped = mapAirportRowsShared_({ vals: vals, headers: headers });
    var asShortText = function(value, maxLen) {
      var text = String(value == null ? '' : value);
      var max = Number(maxLen || 0) || 120;
      return text.length <= max ? text : text.slice(0, max);
    };

    var airports = mapped.map(function(row) {
      return {
        icao: asShortText(row.icao, 12),
        nome: asShortText(row.nome, 90),
        lat: row.lat,
        lon: row.lon,
        fuelAvailable: asShortText(row.fuelAvailable, 24),
        mtow520: parseFloat(row.mtow520) || 9999,
        mtow550: parseFloat(row.mtow550) || 9999,
        runwayIdent: asShortText(row.runwayIdent, 20),
        runwayHeading: asShortText(row.runwayHeading, 12),
        runwayLength: asShortText(row.runwayLength, 12),
        runwayWidth: asShortText(row.runwayWidth, 12),
        runwaySurfaceActual: asShortText(row.runwaySurfaceActual, 40)
      };
    });

    var nextOffset = startOffset + rowsToFetch;
    return {
      airports: airports,
      nextOffset: nextOffset,
      done: nextOffset >= total,
      total: total
    };
  } catch (e) {
    return { error: e && e.message ? e.message : String(e) };
  }
}
























function getNextMissionId() {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sh = ss.getSheetByName(APP_SHEETS.DISPATCH);
const year = String(new Date().getFullYear()).slice(-2);
if (!sh) return "ADS" + year + "-001";
































const lastRow = sh.getLastRow();
if (lastRow < 2) return "ADS" + year + "-001";
































const ids = sh.getRange(2, 2, lastRow - 1, 1).getValues().flat();
let maxSeq = 0;
ids.forEach(id => {
const match = String(id).match(/^ADS(\d{2})-(\d{3})/);
if (match) {
 const seq = parseInt(match[2], 10);
 if (seq > maxSeq) maxSeq = seq;
}
});
return `ADS${year}-${(maxSeq + 1).toString().padStart(3, "0")}`;
}
































/* ==================================================
3. CALENDAR DATA SOURCE (GROUPED)
================================================== */
































/* ==================================================
FIXED: CALENDAR EVENTS (Uses Column G for Route)
================================================== */
































function getCalendarEvents() {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const dispSheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, "getCalendarEvents");
let events = [];
const data = dispSheet.getDataRange().getValues();
let missions = {};
for (let i = 1; i < data.length; i++) {
 const row = data[i];
 const mId = row[DISPATCH_COL.MISSION_ID];
 if (!mId) continue;
  if (!missions[mId]) {
   missions[mId] = {
     id: mId,
     date: row[DISPATCH_COL.DATE],
     acft: row[DISPATCH_COL.AIRCRAFT],
     pilot: row[DISPATCH_COL.PILOT],
     status: row[DISPATCH_COL.STATUS],
     legs: []
   };
 }
































 // --- THE FIX: Read Route from Column G (Index 6) ---
 const route = splitRoute_(row[DISPATCH_COL.ROUTE]);
 const from = route.from || "?";
 const to = route.to || "?";
  const legTime = safeNumber_(row[DISPATCH_COL.FLIGHT_TIME], 0);


 missions[mId].legs.push({
   from: from,
   to: to,
   time: legTime
 });
}
































Object.values(missions).forEach(m => {
 // Build Route String: "A - B - C"
 let routeDisplay = "Local";
 let totalFlt = 0;
  if (m.legs.length > 0) {
   // Start with the origin of the first leg
   routeDisplay = m.legs[0].from;
    // Append the destination of every leg
   m.legs.forEach(leg => {
     routeDisplay += " - " + leg.to;
     totalFlt += leg.time;
   });
 }
































 // Safe Date Handling
 let dateObj = (m.date instanceof Date) ? m.date : new Date(m.date);
 if (isNaN(dateObj.getTime())) return;
  let status = m.status ? m.status.toString().toUpperCase() : "PENDING";
 let color = "#f57c00"; // Orange
 if (status === "APPROVED") color = "#43a047"; // Green
 if (status === "CANCELLED") color = "#b0bec5"; // Grey
































 events.push({
   start: dateObj.toISOString().split('T')[0],
   color: color,
   extendedProps: {
     type: 'mission',
     id: m.id,
     acft: m.acft,
     pilot: String(m.pilot || '').trim() ? String(m.pilot).trim().split(' ')[0] : 'PILOT TBD',
     fullPilot: String(m.pilot || '').trim() || 'PILOT TBD',
     route: routeDisplay, // Correct String now (e.g. "SDRM - SBBV")
     takeoff: "08:00",
     fltTime: totalFlt.toFixed(1),
     dutyTime: (totalFlt + 1.5).toFixed(1)
   }
 });
});

// --- Flight Time Blocks ---
try {
  const blockSheet = ss.getSheetByName(APP_SHEETS.FLIGHT_BLOCKS);
  if (blockSheet) {
    const blockData = blockSheet.getDataRange().getValues();
    for (let bi = 1; bi < blockData.length; bi++) {
      const br = blockData[bi];
      const bStatus = String(br[FLIGHT_BLOCKS_COL.STATUS] || 'ACTIVE').toUpperCase();
      if (bStatus === 'DELETED') continue;
      const bStartRaw = br[FLIGHT_BLOCKS_COL.DATE_START];
      const bEndRaw   = br[FLIGHT_BLOCKS_COL.DATE_END];
      const bStart = bStartRaw instanceof Date ? bStartRaw : new Date(bStartRaw);
      const bEnd   = bEndRaw   instanceof Date ? bEndRaw   : new Date(bEndRaw);
      if (isNaN(bStart.getTime()) || isNaN(bEnd.getTime())) continue;
      // FullCalendar all-day end is exclusive — add 1 day
      const bEndExcl = new Date(bEnd);
      bEndExcl.setDate(bEndExcl.getDate() + 1);
      events.push({
        start: bStart.toISOString().split('T')[0],
        end:   bEndExcl.toISOString().split('T')[0],
        color: '#6a1b9a',
        extendedProps: {
          type:         'block',
          blockId:      String(br[FLIGHT_BLOCKS_COL.BLOCK_ID] || ''),
          name:         String(br[FLIGHT_BLOCKS_COL.NAME]     || ''),
          acft:         String(br[FLIGHT_BLOCKS_COL.AIRCRAFT]  || ''),
          allocatedHrs: safeNumber_(br[FLIGHT_BLOCKS_COL.ALLOCATED_HRS], 0),
          blockType:    String(br[FLIGHT_BLOCKS_COL.TYPE]      || '')
        }
      });
    }
  }
} catch(e) { /* blocks are non-critical */ }

return events;
}

/* ==================================================
   FLIGHT TIME BLOCKS
================================================== */
function getFlightBlocks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const blockSheet = ss.getSheetByName(APP_SHEETS.FLIGHT_BLOCKS);
  if (!blockSheet) return [];
  const blockData = blockSheet.getDataRange().getValues();
  if (blockData.length <= 1) return [];

  const dispSheet = ss.getSheetByName(APP_SHEETS.DISPATCH);
  const dispData = dispSheet ? dispSheet.getDataRange().getValues() : [];

  const blocks = [];
  for (let i = 1; i < blockData.length; i++) {
    const row = blockData[i];
    const blockId = String(row[FLIGHT_BLOCKS_COL.BLOCK_ID] || '').trim();
    if (!blockId) continue;
    const status = String(row[FLIGHT_BLOCKS_COL.STATUS] || 'ACTIVE').toUpperCase();
    if (status === 'DELETED') continue;

    const aircraft = String(row[FLIGHT_BLOCKS_COL.AIRCRAFT] || '').trim();
    const startRaw = row[FLIGHT_BLOCKS_COL.DATE_START];
    const endRaw   = row[FLIGHT_BLOCKS_COL.DATE_END];
    const dateStart = startRaw instanceof Date ? startRaw : new Date(startRaw);
    const dateEnd   = endRaw   instanceof Date ? endRaw   : new Date(endRaw);
    if (isNaN(dateStart.getTime()) || isNaN(dateEnd.getTime())) continue;

    // Sum dispatch flight hours for this aircraft within the date range
    let usedHours = 0;
    for (let j = 1; j < dispData.length; j++) {
      const dr = dispData[j];
      if (String(dr[DISPATCH_COL.AIRCRAFT] || '').trim() !== aircraft) continue;
      const ds = String(dr[DISPATCH_COL.STATUS] || '').toUpperCase();
      if (ds === 'CANCELLED') continue;
      const dRaw = dr[DISPATCH_COL.DATE];
      const dDate = dRaw instanceof Date ? dRaw : new Date(dRaw);
      if (isNaN(dDate.getTime())) continue;
      if (dDate >= dateStart && dDate <= dateEnd) {
        usedHours += safeNumber_(dr[DISPATCH_COL.FLIGHT_TIME], 0);
      }
    }

    blocks.push({
      blockId:      blockId,
      name:         String(row[FLIGHT_BLOCKS_COL.NAME]          || '').trim(),
      aircraft:     aircraft,
      type:         String(row[FLIGHT_BLOCKS_COL.TYPE]          || '').trim(),
      allocatedHrs: safeNumber_(row[FLIGHT_BLOCKS_COL.ALLOCATED_HRS], 0),
      dateStart:    Utilities.formatDate(dateStart, 'GMT', 'yyyy-MM-dd'),
      dateEnd:      Utilities.formatDate(dateEnd,   'GMT', 'yyyy-MM-dd'),
      notes:        String(row[FLIGHT_BLOCKS_COL.NOTES]         || '').trim(),
      status:       status,
      usedHrs:      Math.round(usedHours * 10) / 10
    });
  }
  return blocks;
}

function saveFlightBlock(data) {
  if (!data || !data.name || !data.aircraft || !data.dateStart || !data.dateEnd) {
    throw new Error('Missing required block fields.');
  }
  const allocHrs = parseFloat(data.allocatedHrs);
  if (isNaN(allocHrs) || allocHrs <= 0) throw new Error('Allocated hours must be a positive number.');
  if (data.dateEnd < data.dateStart) throw new Error('End date must be on or after start date.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let blockSheet = ss.getSheetByName(APP_SHEETS.FLIGHT_BLOCKS);
  if (!blockSheet) {
    blockSheet = ss.insertSheet(APP_SHEETS.FLIGHT_BLOCKS);
    blockSheet.appendRow(['BLOCK_ID','NAME','AIRCRAFT','TYPE','ALLOCATED_HRS','DATE_START','DATE_END','NOTES','STATUS','CREATED_AT']);
  }

  const now = new Date();
  if (data.blockId) {
    const allData = blockSheet.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][FLIGHT_BLOCKS_COL.BLOCK_ID]).trim() === data.blockId) {
        const r = i + 1;
        blockSheet.getRange(r, FLIGHT_BLOCKS_COL.NAME          + 1).setValue(data.name);
        blockSheet.getRange(r, FLIGHT_BLOCKS_COL.AIRCRAFT      + 1).setValue(data.aircraft);
        blockSheet.getRange(r, FLIGHT_BLOCKS_COL.TYPE          + 1).setValue(data.type || '');
        blockSheet.getRange(r, FLIGHT_BLOCKS_COL.ALLOCATED_HRS + 1).setValue(allocHrs);
        blockSheet.getRange(r, FLIGHT_BLOCKS_COL.DATE_START    + 1).setValue(data.dateStart);
        blockSheet.getRange(r, FLIGHT_BLOCKS_COL.DATE_END      + 1).setValue(data.dateEnd);
        blockSheet.getRange(r, FLIGHT_BLOCKS_COL.NOTES         + 1).setValue(data.notes || '');
        blockSheet.getRange(r, FLIGHT_BLOCKS_COL.STATUS        + 1).setValue(data.status || 'ACTIVE');
        return { ok: true, blockId: data.blockId };
      }
    }
    throw new Error('Block not found: ' + data.blockId);
  } else {
    const blockId = 'BLK-' + Utilities.formatDate(now, 'GMT', 'yyyyMMddHHmmss');
    blockSheet.appendRow([
      blockId,
      data.name,
      data.aircraft,
      data.type || '',
      allocHrs,
      data.dateStart,
      data.dateEnd,
      data.notes || '',
      'ACTIVE',
      Utilities.formatDate(now, 'GMT', 'yyyy-MM-dd HH:mm:ss')
    ]);
    return { ok: true, blockId: blockId };
  }
}

function deleteFlightBlock(blockId) {
  if (!blockId) throw new Error('blockId is required.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const blockSheet = ss.getSheetByName(APP_SHEETS.FLIGHT_BLOCKS);
  if (!blockSheet) throw new Error('Flight blocks sheet not found.');
  const allData = blockSheet.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][FLIGHT_BLOCKS_COL.BLOCK_ID]).trim() === blockId) {
      blockSheet.getRange(i + 1, FLIGHT_BLOCKS_COL.STATUS + 1).setValue('DELETED');
      return { ok: true };
    }
  }
  throw new Error('Block not found: ' + blockId);
}

/* ==================================================
4. DISPATCH SAVING & FATIGUE
================================================== */
function saveMission(data) {
 validateMissionPayload_(data);
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const dispatchSheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, "saveMission");
 const transSheet = ss.getSheetByName(APP_SHEETS.TRANSACTIONS);


 const legs = data.legs;
 const header = {
   date: data.date,
   time: data.time || "08:00",
   acft: data.acft,
   pilot: data.pilot,
   copilot: data.copilot,
   type: data.type,
   notes: data.notes
 };


 if (!legs || legs.length === 0) throw new Error("No legs to save.");


 // Determine missionId from first leg
 const firstFlightId = legs[0].flightLegId;
 const missionId = missionIdFromFlightLeg_(firstFlightId);
 if (!missionId) throw new Error("Invalid flight leg id format: " + firstFlightId);


 // Totals for fatigue
 const newMissionFlightTime = legs.reduce((acc, leg) => acc + (parseFloat(leg.time) || 0), 0);
 const totalGround = legs.reduce((acc, leg) => acc + (parseFloat(leg.groundTime) || 0.5), 0);
 const newMissionDutyTime = 1.0 + newMissionFlightTime + totalGround + 0.75;


 // Fatigue warnings
 let fatigueWarnings = [];
 try {
   fatigueWarnings = checkFatigue(header.pilot, header.date, newMissionFlightTime, newMissionDutyTime, missionId);
 } catch(e) { console.log("Fatigue check skipped"); }


 let finalNotes = header.notes || "";
 if (fatigueWarnings.length > 0) {
   finalNotes = "[⚠ " + fatigueWarnings.join(", ") + "] " + finalNotes;
 }


 // Purge old rows for this mission (if editing)
 const dispatchData = dispatchSheet.getDataRange().getValues();
 const keptDispatchRows = [dispatchData[0]].concat(
   dispatchData.slice(1).filter(r => String(r[DISPATCH_COL.MISSION_ID]) !== missionId)
 );
 rewriteSheetData_(dispatchSheet, keptDispatchRows);

 if (transSheet) {
   const transData = transSheet.getDataRange().getValues();
   const keptTransRows = [transData[0]].concat(
     transData.slice(1).filter(r => String(r[0]).indexOf(missionId) !== 0)
   );
   rewriteSheetData_(transSheet, keptTransRows);
 }


 // Save each leg individually
 legs.forEach(leg => {
   const _routeTokensFrom = function(raw) {
     if (Array.isArray(raw)) {
       return raw.map(function(wp) {
         if (typeof wp === 'string') return String(wp || '').trim().toUpperCase();
         if (wp && typeof wp === 'object') {
           return String(wp.fix || wp.wp_id || wp.WP_ID || wp.ident || wp.icao || '').trim().toUpperCase();
         }
         return '';
       }).filter(Boolean);
     }
     const txt = String(raw || '').trim().toUpperCase();
     if (!txt) return [];
     return txt
       .replace(/[→>]/g, ',')
       .split(/[\n\r,;\/|]+/)
       .map(function(part) { return String(part || '').trim().toUpperCase(); })
       .filter(Boolean);
   };

   const fromIcao = String(leg && (leg.from || leg.origin) || '').trim().toUpperCase();
   const toIcao = String(leg && (leg.to || leg.destination) || '').trim().toUpperCase();

   let routeTokens = _routeTokensFrom(leg && leg.route);
   if (routeTokens.length < 2) routeTokens = _routeTokensFrom(leg && leg.waypoints);
   if (fromIcao && (!routeTokens.length || routeTokens[0] !== fromIcao)) routeTokens.unshift(fromIcao);
   if (toIcao && (!routeTokens.length || routeTokens[routeTokens.length - 1] !== toIcao)) routeTokens.push(toIcao);
   routeTokens = routeTokens.filter(function(token, idx, arr) { return idx === 0 || token !== arr[idx - 1]; });

  const routeCol = routeTokens.join(',') || [fromIcao, toIcao].filter(Boolean).join(',');
   const normalizedLeg = {
     ...leg,
     from: fromIcao || (routeTokens[0] || ''),
     to: toIcao || (routeTokens.length ? routeTokens[routeTokens.length - 1] : ''),
     route: routeCol,
     waypoints: routeTokens.length ? routeTokens : (leg && leg.waypoints)
   };

   const singleLegWrapper = JSON.stringify({
     legs: [{ ...normalizedLeg, missionTime: header.time, meta: { time: header.time } }],
     time: header.time
   });


   // For offline flights, add status "DRAFT_OFFLINE" instead of leaving it blank (which defaults to "PENDING")
   const isOfflineFlight = String(header.type || '').toLowerCase().indexOf('offline') >= 0;
   const flightStatus = isOfflineFlight ? 'DRAFT_OFFLINE' : '';
   dispatchSheet.appendRow([
     normalizedLeg.flightLegId,
     missionId,
     header.date,
     header.acft,
     header.pilot,
     header.copilot || "",
     routeCol,
     normalizedLeg.time,
     header.type,
     singleLegWrapper, // Only this leg
     finalNotes,
     flightStatus  // STATUS column (L)
   ]);
   // Log fuel deduction only for cache stops (never supplier fuel).
  const fuelDraw = parseFloat(normalizedLeg.plannedCacheDraw) || 0;
  const isCacheStop = (normalizedLeg && (normalizedLeg.isFuelCacheStop === true || String(normalizedLeg.fuelStopSource || '').toLowerCase() === 'cache'));


   if (fuelDraw > 0 && isCacheStop) {
     // Log the deduction from the specific "FROM" location
     // Using leg.from ensures it deducts from the airport where the fuel was pumped
     logFuelChange(
       normalizedLeg.to,
       -fuelDraw,
       header.acft,
       header.pilot,
       normalizedLeg.flightLegId
     );
   }
   // Save passengers
   if (transSheet && normalizedLeg.pax && Array.isArray(normalizedLeg.pax)) {
     normalizedLeg.pax.forEach(p => {
       const effectiveWeight = (p.name === "FREIGHT") ? p.cargo : p.weight;
       transSheet.appendRow([
         normalizedLeg.flightLegId,
         p.fund || "",
         p.name,
         p.category || "",
         effectiveWeight,
         p.chargeRate,
         p.chargedAmount,
         "PENDING",
         p.phone || "",
         p.description || ""
       ]);
     });
   }
 });

 CacheService.getScriptCache().remove('scheduledMissions:v1');


 return "Success: " + missionId;
}




function checkFatigue(pilotName, missionDateStr, newFlight, newDuty, currentMissionId) {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, "checkFatigue");
const data = sheet.getDataRange().getValues();
appLog_('checkFatigue: pilot=', pilotName, 'currentMissionId=', currentMissionId, 'DB_Dispatch rows=', data.length);
const LIMIT_DAILY_FLIGHT = 9.5;
const LIMIT_DAILY_DUTY = 14.0;
const LIMIT_WEEKLY_DUTY = 55.0;
































const missionDate = new Date(missionDateStr);
missionDate.setHours(0,0,0,0);
const oneDay = 24 * 60 * 60 * 1000;
let existingDailyFlight = 0;
let existingDailyDuty = 0;
let existingWeeklyDuty = 0;
































const processedMissionsDay = new Set();
const processedMissionsWeek = new Set();
































for (let i = 1; i < data.length; i++) {
const rowPilot = data[i][DISPATCH_COL.PILOT];
const rowMissionId = data[i][DISPATCH_COL.MISSION_ID];
const rowFlightTime = parseFloat(data[i][DISPATCH_COL.FLIGHT_TIME]) || 0;
































if (rowPilot !== pilotName) continue;
if (rowMissionId === currentMissionId) continue;
































let rowDate = new Date(data[i][DISPATCH_COL.DATE]);
rowDate.setHours(0,0,0,0);
const diffDays = Math.floor((missionDate - rowDate) / oneDay);
































// SAME DAY
if (diffDays === 0) {
 existingDailyFlight += rowFlightTime;
 if (!processedMissionsDay.has(rowMissionId)) {
   processedMissionsDay.add(rowMissionId);
   existingDailyDuty += 1.75;
 }
 existingDailyDuty += (rowFlightTime + 0.5);
}
































// WEEKLY (Last 6 days)
if (diffDays >= 0 && diffDays <= 6) {
 if (!processedMissionsWeek.has(rowMissionId)) {
   processedMissionsWeek.add(rowMissionId);
   existingWeeklyDuty += 1.75;
 }
 existingWeeklyDuty += (rowFlightTime + 0.5);
}
}








const warnings = [];
if ((existingDailyFlight + newFlight) > LIMIT_DAILY_FLIGHT) warnings.push(`DAILY FLIGHT OVER`);
if ((existingDailyDuty + newDuty) > LIMIT_DAILY_DUTY) warnings.push(`DAILY DUTY OVER`);
if ((existingWeeklyDuty + newDuty) > LIMIT_WEEKLY_DUTY) warnings.push(`WEEKLY DUTY OVER`);








return warnings;
}








/* ==================================================
5. SUPERVISOR DASHBOARD (NO AUTH CHECK)
================================================== */




function getSupervisorDashboard() {
// 1. HARDCODE USER (Stops the "Unsafe Attempt" Auth Crash)
const user = "Admin";




const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName(APP_SHEETS.DISPATCH);
if (!sheet) return { user: user, missions: [] };




const data = sheet.getDataRange().getValues();
const missionMap = {};
// Loop rows (Skip header)
for (let i = 1; i < data.length; i++) {
const mId = data[i][DISPATCH_COL.MISSION_ID];
if (!mId) continue;
const status = data[i][DISPATCH_COL.STATUS] ? data[i][DISPATCH_COL.STATUS].toString().toUpperCase() : "PENDING";
// SAFE DATE HANDLING: Keep raw value, convert to string only if needed
let dateDisp = data[i][DISPATCH_COL.DATE];
if (dateDisp instanceof Date) {
  try { dateDisp = Utilities.formatDate(dateDisp, Session.getScriptTimeZone(), "yyyy-MM-dd"); }
  catch(e) { dateDisp = String(dateDisp); }
} else {
  dateDisp = String(dateDisp);
}
if (!missionMap[mId]) {
 missionMap[mId] = {
   id: mId,
   date: dateDisp,
  acft: String(data[i][DISPATCH_COL.AIRCRAFT]),
  pilot: String(data[i][DISPATCH_COL.PILOT]),
   status: status,
   legs: [],
  warnings: String(data[i][DISPATCH_COL.NOTES] || "")
 };
}
let legSummary = `${data[i][DISPATCH_COL.ROUTE]}`;
// Safely parse float
let timeVal = parseFloat(data[i][DISPATCH_COL.FLIGHT_TIME]);
if(!isNaN(timeVal)) legSummary += ` (${timeVal.toFixed(1)})`;
missionMap[mId].legs.push(legSummary);
}
































// Sort
const missions = Object.values(missionMap).sort((a,b) => {
if (a.status === 'PENDING' && b.status !== 'PENDING') return -1;
if (a.status !== 'PENDING' && b.status === 'PENDING') return 1;
return String(b.date).localeCompare(String(a.date));
});
































return { user: user, missions: missions };
}
































/* ==================================================
FIXED: MISSION DETAILS (PURE JS - NO AUTH CRASH)
================================================== */
































































function getMissionDetailsForSupervisor(missionId) {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName(APP_SHEETS.DISPATCH);
const data = sheet.getDataRange().getValues();
// 1. SMART SEARCH (Check Mission ID, then Flight ID)
let missionRows = data.filter(r => String(r[DISPATCH_COL.MISSION_ID]) === String(missionId));
console.log('getMissionById: rows matching missionId=', missionRows.length);
if (missionRows.length === 0) {
const legRow = data.find(r => String(r[DISPATCH_COL.FLIGHT_ID]) === String(missionId));
if (legRow) {
 const realMissionId = legRow[DISPATCH_COL.MISSION_ID];
 missionRows = data.filter(r => String(r[DISPATCH_COL.MISSION_ID]) === String(realMissionId));
}
}
































if (missionRows.length === 0) throw new Error("Mission not found: " + missionId);
































const mainRow = missionRows[0];
// 2. Safe Date handling
let rawDate = mainRow[DISPATCH_COL.DATE];
let dateStr = (rawDate instanceof Date) ? rawDate.toISOString().split('T')[0] : String(rawDate);
const missionDateObj = (rawDate instanceof Date) ? rawDate : new Date();
































let missionData = {
id: mainRow[DISPATCH_COL.MISSION_ID],
date: dateStr,
acft: String(mainRow[DISPATCH_COL.AIRCRAFT]),
pilot: String(mainRow[DISPATCH_COL.PILOT]),
meta: {
 date: dateStr,
 acft: String(mainRow[DISPATCH_COL.AIRCRAFT]),
 pilot: String(mainRow[DISPATCH_COL.PILOT]),
 copilot: String(mainRow[DISPATCH_COL.COPILOT]),
 notes: String(mainRow[DISPATCH_COL.NOTES] || "")
},
// 3. PARSING LEGS WITH WAYPOINTS
legs: missionRows.map((r) => {
 let json = {};
 let legPayload = {};
































 try {
  json = JSON.parse(r[DISPATCH_COL.RAW_DATA] || "{}");
   if (json.legs && Array.isArray(json.legs) && json.legs.length > 0) {
      legPayload = json.legs[0];
   } else if (Array.isArray(json)) {
      legPayload = json[0] || {};
   } else {
      legPayload = json;
   }
 } catch(e) { legPayload = {}; }
  const safeNum = (val, def) => {
   const n = parseFloat(val);
   return isNaN(n) ? def : n;
 };
































 // Parse route string using comma-delimited policy with legacy-safe fallback.
 const parsedRoute = splitRoute_(r[DISPATCH_COL.ROUTE]);
































 return {
  flightLegId: r[DISPATCH_COL.FLIGHT_ID],
  from: parsedRoute.from || "?",
  to: parsedRoute.to || "?",
    // --- NEW: Pulling waypoints from the JSON ---
   waypoints: legPayload.waypoints || "",
    time: safeNum(r[DISPATCH_COL.FLIGHT_TIME], 0),
   dist: safeNum(legPayload.dist, 0),
   groundTime: safeNum(legPayload.groundTime, 0.5),
   fuel: safeNum(legPayload.fuel, 0),
   takeoffFuel: safeNum(legPayload.takeoffFuel, 0),
   landingFuel: safeNum(legPayload.landingFuel, 0),
   payload: safeNum(legPayload.payload, 0),
   availPayload: safeNum(legPayload.availPayload, 0),
   limit: safeNum(legPayload.limit, 0),
    pax: legPayload.pax || [],
   limitType: legPayload.limitType || "",
   isOver: legPayload.isOver || false,
   missionTime: legPayload.missionTime || "08:00"
 };
})
};
































// 4. Helpers for the Supervisor Sidebar
const pName = String(missionData.meta.pilot);
let timeline = [];
try { timeline = getPilotDutyTimeline(pName, missionDateObj); } catch(e) {}
let authString = "";
try { authString = getAuthorizedDestinations(pName); } catch(e) {}
return {
mission: missionData,
timeline: timeline,
authorizedAirports: authString
};
}
































































// 6. HELPER ACTIONS
































function _getSupervisorApprovalPassword_() {
  const props = PropertiesService.getScriptProperties();
  const keys = [
    'SUPERVISOR_APPROVAL_PASSWORD',
    'SUPERVISOR_PASSWORD',
    'APPROVAL_PASSWORD'
  ];
  for (let i = 0; i < keys.length; i++) {
    const value = props.getProperty(keys[i]);
    if (value && String(value).trim()) return String(value);
  }
  return '';
}

function _verifySupervisorApprovalPassword_(password) {
  const configured = _getSupervisorApprovalPassword_();
  if (!configured) {
    throw new Error('Supervisor approval password not configured. Set script property SUPERVISOR_APPROVAL_PASSWORD.');
  }
  if (String(password || '') !== configured) {
    throw new Error('Invalid supervisor password');
  }
  return true;
}

function approveMission(missionId, approvalPassword) {
_verifySupervisorApprovalPassword_(approvalPassword);
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName(APP_SHEETS.DISPATCH);
if (!sheet) return "Error: DB missing";
const data = sheet.getDataRange().getValues();
const user = "Admin";
for (let i = 1; i < data.length; i++) {
if (String(data[i][DISPATCH_COL.MISSION_ID]) === String(missionId)) {
 const pilotName = String(data[i][DISPATCH_COL.PILOT] || '').trim();
 const pilotKey = pilotName.toUpperCase();
 if (!pilotName || pilotKey === 'PILOT TBD' || pilotKey === 'TBD' || pilotKey === 'UNASSIGNED') {
  throw new Error('Mission cannot be approved without an assigned pilot.');
 }
 sheet.getRange(i + 1, DISPATCH_COL.STATUS + 1).setValue("APPROVED");
}
}
CacheService.getScriptCache().remove('scheduledMissions:v1');
const audit = ss.getSheetByName(APP_SHEETS.AUDIT);
if(audit) audit.appendRow([new Date(), user, missionId, "APPROVE", "PENDING", "APPROVED", "Release"]);
return "Approved";
}
































function waiveDestinationCheck(pilot, icao, missionId, approvalPassword) {
_verifySupervisorApprovalPassword_(approvalPassword);
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName(APP_SHEETS.CHECKS);
if(!sheet) return "Error: DB_Checks missing";
const data = sheet.getDataRange().getValues();
const user = "Admin";
let found = false;
for (let i = 1; i < data.length; i++) {
if (String(data[i][CHECKS_COL.PILOT]) === String(pilot)) {
 const current = data[i][CHECKS_COL.AUTH_DESTINATIONS] || "";
 if (!current.includes(icao)) {
   const newVal = current ? current + ", " + icao : icao;
   sheet.getRange(i + 1, CHECKS_COL.AUTH_DESTINATIONS + 1).setValue(newVal);
   const audit = ss.getSheetByName(APP_SHEETS.AUDIT);
   if(audit) audit.appendRow([new Date(), user, missionId, "WAIVE_CHECK", current, newVal, icao]);
 }
 found = true;
 break;
}
}
return found ? "Check Waived" : "Pilot not found";
}
































function getAuthorizedDestinations(pilotName) {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName(APP_SHEETS.CHECKS);
if (!sheet) return "";
const data = sheet.getDataRange().getValues();
const target = String(pilotName).trim().toUpperCase();
let allDestinations = [];
for (let i = 1; i < data.length; i++) {
const currentPilot = String(data[i][CHECKS_COL.PILOT]).trim().toUpperCase();
if (currentPilot === target) {
 const dests = String(data[i][CHECKS_COL.AUTH_DESTINATIONS] || "");
 if (dests.length > 0) allDestinations.push(dests);
}
}
return allDestinations.join(", ");
}
































/* ==================================================
FIXED: TIMELINE (PURE JS - NO AUTH CRASH)
================================================== */
































function getPilotDutyTimeline(pilotName, centerDate) {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const oneDay = 24*60*60*1000;
if (!(centerDate instanceof Date) || isNaN(centerDate)) centerDate = new Date();
































const startDate = new Date(centerDate.getTime() - (7 * oneDay));
const endDate = new Date(centerDate.getTime() + (7 * oneDay));
let events = [];
const logSheet = ss.getSheetByName(APP_SHEETS.DUTY_LOG);
if (logSheet) {
const logData = logSheet.getDataRange().getValues();
for (let i = 1; i < logData.length; i++) {
 if (logData[i][DUTY_LOG_COL.PILOT] === pilotName) {
   let d = logData[i][DUTY_LOG_COL.DATE];
   if (d instanceof Date && d >= startDate && d <= endDate) {
     events.push({
       date: d.toISOString(), // Safe String
       type: "LOGGED",
       title: String(logData[i][DUTY_LOG_COL.TITLE]),
       desc: String(logData[i][DUTY_LOG_COL.DESC_PRIMARY] || logData[i][DUTY_LOG_COL.DESC_FALLBACK]),
       flightHrs: 0, dutyHrs: 0
     });
   }
 }
}
}
const dispSheet = ss.getSheetByName(APP_SHEETS.DISPATCH);
if (dispSheet) {
const dispData = dispSheet.getDataRange().getValues();
const tracker = {};
for (let i = 1; i < dispData.length; i++) {
 if (dispData[i][DISPATCH_COL.PILOT] === pilotName) {
   let d = dispData[i][DISPATCH_COL.DATE];
   if (d instanceof Date) {
     d.setHours(0,0,0,0);
     if (d >= startDate && d <= endDate) {
       const mId = dispData[i][DISPATCH_COL.MISSION_ID];
       if(!tracker[mId]) {
         tracker[mId] = {
           date: d.toISOString(), // Safe String
           type: "SCHEDULED",
           title: mId,
           desc: dispData[i][DISPATCH_COL.ROUTE],
           flightHrs: 0, dutyHrs: 0
         };
       }
       const flt = parseFloat(dispData[i][DISPATCH_COL.FLIGHT_TIME]) || 0;
       tracker[mId].flightHrs += flt;
       tracker[mId].dutyHrs = tracker[mId].flightHrs + 1.5;
     }
   }
 }
}
Object.values(tracker).forEach(e => events.push(e));
}
return events.sort((a, b) => a.date.localeCompare(b.date));
}
































































/* ==================================================
FIXED: SCHEDULED MISSIONS LIST (PURE JS)
================================================== */
































function getScheduledMissions() {
const cache = CacheService.getScriptCache();
const cacheKey = 'scheduledMissions:v1';
const cached = cache.get(cacheKey);
if (cached) {
  try {
    return JSON.parse(cached);
  } catch (e) {}
}

const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, "getScheduledMissions");

const lastRow = sheet.getLastRow();
if (lastRow < 2) return [];

// Read only needed columns B:L (mission id/date/acft/pilot/route/status)
const data = sheet.getRange(2, 2, lastRow - 1, 11).getValues();
const missions = {};
const DISPATCH_RANGE_COL = {
 MISSION_ID: DISPATCH_COL.MISSION_ID - 1,
 DATE: DISPATCH_COL.DATE - 1,
 AIRCRAFT: DISPATCH_COL.AIRCRAFT - 1,
 PILOT: DISPATCH_COL.PILOT - 1,
 ROUTE: DISPATCH_COL.ROUTE - 1,
 STATUS: DISPATCH_COL.STATUS - 1
};








for (let i = 0; i < data.length; i++) {
 const row = data[i];
 const mId = row[DISPATCH_RANGE_COL.MISSION_ID];
 if (!mId) continue;
  if (!missions[mId]) {
   let d = row[DISPATCH_RANGE_COL.DATE];
   let dStr = "";
   if (d instanceof Date) dStr = d.toISOString().split('T')[0];
   else dStr = String(d || '');








   missions[mId] = {
     id: mId,
     date: dStr,
     acft: row[DISPATCH_RANGE_COL.AIRCRAFT],
     pilot: row[DISPATCH_RANGE_COL.PILOT],
     status: row[DISPATCH_RANGE_COL.STATUS],
     routeStr: ""
   };
 }
  const legRoute = String(row[DISPATCH_RANGE_COL.ROUTE] || '');
  if (missions[mId].routeStr === "") {
    missions[mId].routeStr = legRoute;
  } else {
    const prior = splitRoute_(missions[mId].routeStr);
    const next = splitRoute_(legRoute);
    if (next.to && next.to !== prior.to) {
      missions[mId].routeStr += ',' + next.to;
    }
  }
}








// Convert the object back to an array for the frontend
const result = Object.values(missions).map(m => {
  const routeTokens = routeTokensFromString_(m.routeStr || '');
  const fromIcao = routeTokens[0] || '';
  const toIcao = routeTokens.length ? routeTokens[routeTokens.length - 1] : '';

  return {
    id: m.id,
    date: m.date,
    acft: m.acft,
    pilot: m.pilot,
    status: m.status, // Passes Column L value to the HTML
    from: fromIcao,
    to: toIcao,
    route: m.routeStr
  };
}).reverse().slice(0, 15);

cache.put(cacheKey, JSON.stringify(result), 45);
return result;
}


/* ==================================================
MISSION-WIDE DELETION & FUEL REVERSAL
================================================== */


function cancelMissionFromDatabase(missionId) {
 const ss = SpreadsheetApp.getActiveSpreadsheet();


 // 1. REVERSE ALL FUEL FOR THIS MISSION
 // We do this first so we can read the log data before deleting it
 reverseFuelForMission(missionId);


 // 2. DELETE FROM DB_DISPATCH
 const dbSheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, "cancelMissionFromDatabase");
 const dbData = dbSheet.getDataRange().getValues();
 const keptDispatchRows = [dbData[0]].concat(
   dbData.slice(1).filter(r => String(r[DISPATCH_COL.MISSION_ID]) !== String(missionId))
 );
 rewriteSheetData_(dbSheet, keptDispatchRows);
 CacheService.getScriptCache().remove('scheduledMissions:v1');


 // 3. DELETE FROM DB_TRANSACTIONS
 const transSheet = ss.getSheetByName(APP_SHEETS.TRANSACTIONS);
 if (transSheet) {
   const transData = transSheet.getDataRange().getValues();
   const keptTransRows = [transData[0]].concat(
     transData.slice(1).filter(r => String(r[0]).indexOf(String(missionId)) !== 0)
   );
   rewriteSheetData_(transSheet, keptTransRows);
 }
 return "Success: Mission " + missionId + " fully removed.";
}


function reverseFuelForMission(missionId) {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const logSheet = ss.getSheetByName(APP_SHEETS.FUEL_LOGS);
 const cacheSheet = ss.getSheetByName(APP_SHEETS.FUEL_CACHES);
 if (!logSheet || !cacheSheet) return;


 const logData = logSheet.getDataRange().getValues();
 const cacheData = cacheSheet.getDataRange().getValues();


 for (let i = logData.length - 1; i >= 1; i--) {
   const rowFlightId = String(logData[i][FUEL_LOG_COL.FLIGHT_ID]);
   const verified = String(logData[i][FUEL_LOG_COL.VERIFIED] || "").toUpperCase().trim();


   // MATCH: If the Flight ID starts with the Mission ID (e.g., ADS26-003-1 starts with ADS26-003)
   if (rowFlightId.indexOf(missionId) === 0 && (verified === "NO" || verified === "")) {
     const icao = logData[i][FUEL_LOG_COL.ICAO];
     const amount = safeNumber_(logData[i][FUEL_LOG_COL.CHANGE_QTY], 0);


     // Find cache and refund
     for (let j = 1; j < cacheData.length; j++) {
       if (cacheData[j][FUEL_CACHE_COL.ICAO] === icao) {
         const currentInv = safeNumber_(cacheData[j][FUEL_CACHE_COL.CURRENT_QTY], 0);
         const newInv = currentInv - amount; // Current - (-Draw) = Refund
         cacheSheet.getRange(j + 1, FUEL_CACHE_COL.CURRENT_QTY + 1).setValue(newInv);
        
         // Update local data so multiple legs in one mission don't overwrite each other
         cacheData[j][FUEL_CACHE_COL.CURRENT_QTY] = newInv;
         break;
       }
     }
     logSheet.deleteRow(i + 1);
   }
 }
}
function savePassengerToDB(data) {
try {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const sheet = ss.getSheetByName(APP_SHEETS.PASSENGERS);
  if (!sheet) {
   throw new Error("Sheet 'DB_Passengers' not found!");
 }








 const dobRaw = String(data && data.dob || '').trim();
 const dobCellValue = (function() {
   if (!dobRaw) return '';
   if (/^\d{4}-\d{2}-\d{2}$/.test(dobRaw)) {
     const parts = dobRaw.split('-').map(function(p) { return parseInt(p, 10); });
     return new Date(parts[0], parts[1] - 1, parts[2]);
   }
   return dobRaw;
 })();

 // Mapping the data to your headers:
 // Passenger_Name | ID_Type | ID_Number... | DOB | Weight_kg | Gender | PHONE | Notes | Last_Flown
 const newRow = [
   data.name,        // Passenger_Name
   data.id_type,     // ID_Type
   data.id_num,      // ID_Number_CPF_Passport
   dobCellValue,     // DOB
   data.weight,      // Weight_kg
   data.gender,      // Gender
   data.phone,       // PHONE
   "Added via App",  // Notes
   new Date()        // Last_Flown (Today)
 ];








 sheet.appendRow(newRow);
 return { success: true };
} catch (e) {
 console.error(e.toString());
 return { success: false, error: e.toString() };
}
}

function _toolsNormHeader_(value) {
  return String(value || '').trim().toUpperCase().replace(/\s+/g, '_').replace(/[^A-Z0-9_]/g, '');
}

function _toolsSheetNameFromKind_(kind) {
  var k = String(kind || '').trim().toLowerCase();
  if (k === 'airports') return APP_SHEETS.AIRPORTS;
  if (k === 'pilots') return APP_SHEETS.PILOTS;
  if (k === 'routes') return APP_SHEETS.ROUTES;
  if (k === 'passengers') return APP_SHEETS.PASSENGERS;
  if (k === 'fuelcaches' || k === 'fuel-cache' || k === 'fuel_caches') return APP_SHEETS.FUEL_CACHES;
  if (k === 'contacts' || k === 'fuelcontacts' || k === 'fuel-contacts' || k === 'fuel_contacts') return APP_SHEETS.CONTACTS || 'DB_Contacts';
  if (k === 'syllabus' || k === 'training' || k === 'ref_syllabus') return APP_SHEETS.SYLLABUS;
  throw new Error('Unsupported tools sheet kind: ' + kind);
}

function _toolsSheetHeaderRow_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) throw new Error('Header row not found in ' + sheet.getName());
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

function _toolsRowPayloadFromHeaders_(headerRow, row) {
  var payload = {};
  (headerRow || []).forEach(function(header, idx) {
    var label = String(header || '').trim();
    if (!label) return;
    payload[_toolsNormHeader_(label)] = row && idx < row.length ? row[idx] : '';
  });
  return payload;
}

function getToolsSheetHeaders(kind) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = _toolsSheetNameFromKind_(kind);
    var sh = getRequiredSheet_(ss, sheetName, 'getToolsSheetHeaders');
    var headerRow = _toolsSheetHeaderRow_(sh);
    var headers = headerRow
      .map(function(h) { return String(h || '').trim(); })
      .filter(function(h) { return !!h; })
      .map(function(h) { return { label: h, key: _toolsNormHeader_(h) }; });
    return { success: true, kind: String(kind || ''), sheetName: sheetName, headers: headers };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function addToolsSheetRecord(kind, payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = _toolsSheetNameFromKind_(kind);
    var sh = getRequiredSheet_(ss, sheetName, 'addToolsSheetRecord');
    var headerRow = _toolsSheetHeaderRow_(sh);
    var data = (payload && typeof payload === 'object') ? payload : {};

    var row = headerRow.map(function(header) {
      var label = String(header || '').trim();
      if (!label) return '';
      var key = _toolsNormHeader_(label);
      if (Object.prototype.hasOwnProperty.call(data, key)) return data[key];
      if (Object.prototype.hasOwnProperty.call(data, label)) return data[label];
      return '';
    });

    if (sheetName === APP_SHEETS.ROUTES) {
      var routeHeaders = headerRow.map(function(h) { return _toolsNormHeader_(h); });
      var wpIdx = routeHeaders.indexOf('WAYPOINT_LIST');
      if (wpIdx >= 0) {
        row[wpIdx] = _toolsNormalizeRouteWaypointList_(row[wpIdx]);
      }
    }

    sh.appendRow(row);
    var rowNumber = sh.getLastRow();
    var response = { success: true, sheetName: sheetName, rowNumber: rowNumber };

    if (sheetName === APP_SHEETS.AIRPORTS) {
      var airportSync = _toolsEnsureAirportPhotoFolderForRow_(sh, headerRow, rowNumber);
      if (airportSync && airportSync.success) {
        response.airportPhotoFolder = airportSync;
      }
    }

    return response;
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _toolsFirstHeaderMatch_(headers, candidates) {
  var list = Array.isArray(headers) ? headers : [];
  var norms = list.map(function(h) { return _toolsNormHeader_(h); });
  var keys = Array.isArray(candidates) ? candidates : [candidates];
  for (var i = 0; i < keys.length; i++) {
    var idx = norms.indexOf(_toolsNormHeader_(keys[i]));
    if (idx >= 0) return idx;
  }
  return -1;
}

function _toolsNormalizeKeyValue_(value) {
  return String(value || '').trim().toUpperCase();
}

function _toolsNormalizeRouteWaypointList_(value) {
  var normalized = String(value || '')
    .toUpperCase()
    .replace(/\u2192|->/g, ',')
    .replace(/[;|/\n\r]+/g, ',');
  return normalized
    .split(',')
    .map(function(token) { return String(token || '').trim(); })
    .filter(Boolean)
    .join(', ');
}

function getToolsAircraftBuilderTemplate(sourceRegistration) {
  try {
    var sourceReg = String(sourceRegistration || '').trim();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aircraftSheet = getRequiredSheet_(ss, APP_SHEETS.AIRCRAFT, 'getToolsAircraftBuilderTemplate');
    var airframesSheet = getRequiredSheet_(ss, APP_SHEETS.AIRFRAMES, 'getToolsAircraftBuilderTemplate');
    var envelopesSheet = getRequiredSheet_(ss, APP_SHEETS.ENVELOPES, 'getToolsAircraftBuilderTemplate');
    var rollSheet = getRequiredSheet_(ss, 'Aircraft_Roll_Numbers', 'getToolsAircraftBuilderTemplate');

    var aircraftHeaders = _toolsSheetHeaderRow_(aircraftSheet);
    var airframeHeaders = _toolsSheetHeaderRow_(airframesSheet);
    var envelopeHeaders = _toolsSheetHeaderRow_(envelopesSheet);
    var rollHeaders = _toolsSheetHeaderRow_(rollSheet);

    var aircraftData = aircraftSheet.getDataRange().getValues();
    var airframeData = airframesSheet.getDataRange().getValues();
    var envelopeData = envelopesSheet.getDataRange().getValues();
    var rollData = rollSheet.getDataRange().getValues();

    var regIdx = _toolsFirstHeaderMatch_(aircraftHeaders, ['REGISTRATION', 'REG', 'TAIL', 'TAIL_NUMBER']);
    var typeIdx = _toolsFirstHeaderMatch_(aircraftHeaders, ['AIRCRAFT_TYPE', 'TYPE_FOR_PERFORMANCE']);

    var aircraftOptions = [];
    for (var i = 1; i < aircraftData.length; i++) {
      var row = aircraftData[i];
      var reg = regIdx >= 0 ? String(row[regIdx] || '').trim() : '';
      if (!reg) continue;
      aircraftOptions.push({
        registration: reg,
        aircraftType: typeIdx >= 0 ? String(row[typeIdx] || '').trim() : ''
      });
    }

    var template = {
      aircraftRow: {},
      airframeRows: [],
      envelopeRows: [],
      rollRows: []
    };

    if (sourceReg && regIdx >= 0) {
      var sourceAircraftRow = null;
      for (var ar = 1; ar < aircraftData.length; ar++) {
        var rawReg = String(aircraftData[ar][regIdx] || '').trim();
        if (_toolsNormalizeKeyValue_(rawReg) === _toolsNormalizeKeyValue_(sourceReg)) {
          sourceAircraftRow = aircraftData[ar];
          break;
        }
      }

      if (sourceAircraftRow) {
        template.aircraftRow = _toolsRowPayloadFromHeaders_(aircraftHeaders, sourceAircraftRow);
        var sourceType = '';
        if (typeIdx >= 0) sourceType = String(sourceAircraftRow[typeIdx] || '').trim();

        var afTypeIdx = _toolsFirstHeaderMatch_(airframeHeaders, ['AIRCRAFT_TYPE']);
        var envTypeIdx = _toolsFirstHeaderMatch_(envelopeHeaders, ['AIRCRAFT_TYPE']);
        var rollTypeIdx = _toolsFirstHeaderMatch_(rollHeaders, ['AIRCRAFT_TYPE']);

        template.airframeRows = airframeData.slice(1)
          .filter(function(row) {
            if (!sourceType || afTypeIdx < 0) return false;
            return _toolsNormalizeKeyValue_(row[afTypeIdx]) === _toolsNormalizeKeyValue_(sourceType);
          })
          .map(function(row) { return _toolsRowPayloadFromHeaders_(airframeHeaders, row); });

        template.envelopeRows = envelopeData.slice(1)
          .filter(function(row) {
            if (!sourceType || envTypeIdx < 0) return false;
            return _toolsNormalizeKeyValue_(row[envTypeIdx]) === _toolsNormalizeKeyValue_(sourceType);
          })
          .map(function(row) { return _toolsRowPayloadFromHeaders_(envelopeHeaders, row); });

        template.rollRows = rollData.slice(1)
          .filter(function(row) {
            if (!sourceType || rollTypeIdx < 0) return false;
            return _toolsNormalizeKeyValue_(row[rollTypeIdx]) === _toolsNormalizeKeyValue_(sourceType);
          })
          .map(function(row) { return _toolsRowPayloadFromHeaders_(rollHeaders, row); });
      }
    }

    return {
      success: true,
      aircraftOptions: aircraftOptions,
      sections: {
        aircraft: aircraftHeaders.map(function(h) { return { label: String(h || '').trim(), key: _toolsNormHeader_(h) }; }),
        airframes: airframeHeaders.map(function(h) { return { label: String(h || '').trim(), key: _toolsNormHeader_(h) }; }),
        envelopes: envelopeHeaders.map(function(h) { return { label: String(h || '').trim(), key: _toolsNormHeader_(h) }; }),
        rollnumbers: rollHeaders.map(function(h) { return { label: String(h || '').trim(), key: _toolsNormHeader_(h) }; })
      },
      template: template
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function saveToolsAircraftBundle(payload) {
  try {
    var data = (payload && typeof payload === 'object') ? payload : {};
    var aircraftPayload = (data.aircraftRow && typeof data.aircraftRow === 'object') ? data.aircraftRow : {};
    var airframeRows = Array.isArray(data.airframeRows) ? data.airframeRows : [];
    var envelopeRows = Array.isArray(data.envelopeRows) ? data.envelopeRows : [];
    var rollRows = Array.isArray(data.rollRows) ? data.rollRows : [];

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aircraftSheet = getRequiredSheet_(ss, APP_SHEETS.AIRCRAFT, 'saveToolsAircraftBundle');
    var airframesSheet = getRequiredSheet_(ss, APP_SHEETS.AIRFRAMES, 'saveToolsAircraftBundle');
    var envelopesSheet = getRequiredSheet_(ss, APP_SHEETS.ENVELOPES, 'saveToolsAircraftBundle');
    var rollSheet = getRequiredSheet_(ss, 'Aircraft_Roll_Numbers', 'saveToolsAircraftBundle');

    var aircraftHeaders = _toolsSheetHeaderRow_(aircraftSheet);
    var airframeHeaders = _toolsSheetHeaderRow_(airframesSheet);
    var envelopeHeaders = _toolsSheetHeaderRow_(envelopesSheet);
    var rollHeaders = _toolsSheetHeaderRow_(rollSheet);

    function valueFor(header, rowObj) {
      var label = String(header || '').trim();
      if (!label) return '';
      var key = _toolsNormHeader_(label);
      if (Object.prototype.hasOwnProperty.call(rowObj, key)) return rowObj[key];
      if (Object.prototype.hasOwnProperty.call(rowObj, label)) return rowObj[label];
      return '';
    }

    var regKeyIdx = _toolsFirstHeaderMatch_(aircraftHeaders, ['REGISTRATION', 'REG', 'TAIL', 'TAIL_NUMBER']);
    var typeKeyIdx = _toolsFirstHeaderMatch_(aircraftHeaders, ['AIRCRAFT_TYPE', 'TYPE_FOR_PERFORMANCE']);
    var registration = regKeyIdx >= 0 ? String(valueFor(aircraftHeaders[regKeyIdx], aircraftPayload) || '').trim() : '';
    var aircraftType = typeKeyIdx >= 0 ? String(valueFor(aircraftHeaders[typeKeyIdx], aircraftPayload) || '').trim() : '';
    if (!registration) return { success: false, error: 'Registration is required in DB_Aircraft section.' };
    if (!aircraftType) return { success: false, error: 'Aircraft Type is required in DB_Aircraft section.' };

    var existing = aircraftSheet.getDataRange().getValues();
    if (regKeyIdx >= 0 && existing.length > 1) {
      var regNorm = _toolsNormalizeKeyValue_(registration);
      for (var i = 1; i < existing.length; i++) {
        var existingReg = _toolsNormalizeKeyValue_(existing[i][regKeyIdx]);
        if (existingReg && existingReg === regNorm) {
          return { success: false, error: 'Registration already exists in DB_Aircraft: ' + registration };
        }
      }
    }

    function buildRow(headerRow, rowObj) {
      return headerRow.map(function(header) {
        var key = _toolsNormHeader_(header);
        var val = valueFor(header, rowObj);
        if (key === 'AIRCRAFT_TYPE' && aircraftType) return val || aircraftType;
        return val;
      });
    }

    function isNonEmptyRow(values) {
      return values.some(function(v) { return String(v == null ? '' : v).trim() !== ''; });
    }

    function buildRows(headerRow, rowList) {
      return rowList
        .map(function(rowObj) { return buildRow(headerRow, rowObj || {}); })
        .filter(isNonEmptyRow);
    }

    var aircraftRow = buildRow(aircraftHeaders, aircraftPayload);
    if (!isNonEmptyRow(aircraftRow)) return { success: false, error: 'DB_Aircraft row is empty.' };

    var airframeData = buildRows(airframeHeaders, airframeRows);
    var envelopeData = buildRows(envelopeHeaders, envelopeRows);
    var rollData = buildRows(rollHeaders, rollRows);

    if (!airframeData.length) return { success: false, error: 'Add at least one REF_Airframes row.' };
    if (!envelopeData.length) return { success: false, error: 'Add at least one REF_Envelopes row.' };
    if (!rollData.length) return { success: false, error: 'Add at least one Aircraft_Roll_Numbers row.' };

    aircraftSheet.appendRow(aircraftRow);
    var aircraftRowNumber = aircraftSheet.getLastRow();

    var aircraftFolderUrl = '';
    var aircraftFolderWarning = '';
    try {
      var folderRes = _toolsEnsureAircraftDocumentFolderForRow_(aircraftSheet, aircraftHeaders, aircraftRowNumber);
      if (folderRes && folderRes.success) {
        aircraftFolderUrl = String(folderRes.url || '').trim();
      } else if (folderRes && folderRes.error) {
        aircraftFolderWarning = String(folderRes.error || '').trim();
      }
    } catch (folderErr) {
      aircraftFolderWarning = String(folderErr && folderErr.message ? folderErr.message : folderErr);
    }

    airframesSheet.getRange(airframesSheet.getLastRow() + 1, 1, airframeData.length, airframeData[0].length).setValues(airframeData);
    envelopesSheet.getRange(envelopesSheet.getLastRow() + 1, 1, envelopeData.length, envelopeData[0].length).setValues(envelopeData);
    rollSheet.getRange(rollSheet.getLastRow() + 1, 1, rollData.length, rollData[0].length).setValues(rollData);

    var out = {
      success: true,
      aircraftRowNumber: aircraftRowNumber,
      folderUrl: aircraftFolderUrl,
      counts: {
        aircraft: 1,
        airframes: airframeData.length,
        envelopes: envelopeData.length,
        rollnumbers: rollData.length
      }
    };
    if (aircraftFolderWarning) out.folderWarning = aircraftFolderWarning;
    return out;
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _toolsAircraftDocsRootFolder_() {
  var props = PropertiesService.getScriptProperties();
  var existingId = String(props.getProperty('AIRCRAFT_DOCS_ROOT_FOLDER_ID') || '').trim();
  if (existingId) {
    try {
      return DriveApp.getFolderById(existingId);
    } catch (e) {}
  }

  var folderName = 'Aircraft Docs';
  var folders = DriveApp.getFoldersByName(folderName);
  var root = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  props.setProperty('AIRCRAFT_DOCS_ROOT_FOLDER_ID', root.getId());
  return root;
}

function _toolsAircraftRegIndex_(headerRow) {
  var headers = (headerRow || []).map(function(h) { return _toolsNormHeader_(h); });
  var candidates = ['REGISTRATION', 'AIRCRAFT_REGISTRATION', 'TAIL', 'AIRCRAFT'];
  for (var i = 0; i < candidates.length; i++) {
    var idx = headers.indexOf(candidates[i]);
    if (idx >= 0) return idx;
  }
  return -1;
}

function _toolsAircraftDocsFolderUrlIndex_(sheet, headerRow) {
  var headers = (headerRow || []).map(function(h) { return _toolsNormHeader_(h); });
  var candidates = ['DOCUMENTS_FOLDER_URL', 'DRIVE_FOLDER_URL', 'AIRCRAFT_DOCS_URL', 'DOC_FOLDER_URL'];
  for (var i = 0; i < candidates.length; i++) {
    var idx = headers.indexOf(candidates[i]);
    if (idx >= 0) return idx;
  }

  // Preferred canonical column for this feature.
  var newCol = Math.max(sheet.getLastColumn(), 1) + 1;
  sheet.getRange(1, newCol).setValue('DOCUMENTS_FOLDER_URL');
  if (Array.isArray(headerRow)) headerRow.push('DOCUMENTS_FOLDER_URL');
  return newCol - 1;
}

function _toolsEnsureAircraftDocumentFolderForRow_(sheet, headerRow, rowNumber) {
  var regIdx = _toolsAircraftRegIndex_(headerRow);
  if (regIdx < 0) return { success: false, error: 'REGISTRATION column not found in DB_Aircraft.' };

  var folderIdx = _toolsAircraftDocsFolderUrlIndex_(sheet, headerRow);
  var width = Math.max(sheet.getLastColumn(), 1);
  var row = sheet.getRange(rowNumber, 1, 1, width).getValues()[0];

  var registration = String(row[regIdx] || '').trim().toUpperCase();
  if (!registration) return { success: false, error: 'Aircraft registration is empty.' };

  var root = _toolsAircraftDocsRootFolder_();
  var folderName = registration.replace(/[\\/:*?"<>|]+/g, '-');
  var existing = root.getFoldersByName(folderName);
  var folder = existing.hasNext() ? existing.next() : root.createFolder(folderName);
  var url = folder.getUrl();

  var currentUrl = String(row[folderIdx] || '').trim();
  if (!currentUrl || currentUrl !== url) {
    sheet.getRange(rowNumber, folderIdx + 1).setValue(url);
  }

  return {
    success: true,
    registration: registration,
    folderId: folder.getId(),
    folderName: folder.getName(),
    url: url
  };
}

function ensureAircraftDocumentFolderByRegistration(registration) {
  try {
    var reg = String(registration || '').trim().toUpperCase();
    if (!reg) return { success: false, error: 'Registration is required.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.AIRCRAFT, 'ensureAircraftDocumentFolderByRegistration');
    var headers = _toolsSheetHeaderRow_(sh);
    var regIdx = _toolsAircraftRegIndex_(headers);
    if (regIdx < 0) return { success: false, error: 'REGISTRATION column not found in DB_Aircraft.' };

    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][regIdx] || '').trim().toUpperCase() !== reg) continue;
      return _toolsEnsureAircraftDocumentFolderForRow_(sh, headers, i + 1);
    }
    return { success: false, error: 'Aircraft not found in DB_Aircraft: ' + reg };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function ensureAllAircraftDocumentFolders() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.AIRCRAFT, 'ensureAllAircraftDocumentFolders');
    var headers = _toolsSheetHeaderRow_(sh);
    var regIdx = _toolsAircraftRegIndex_(headers);
    if (regIdx < 0) return { success: false, error: 'REGISTRATION column not found in DB_Aircraft.' };

    var data = sh.getDataRange().getValues();
    var created = 0;
    var updated = 0;
    var skipped = 0;
    var errors = [];

    for (var i = 1; i < data.length; i++) {
      var reg = String(data[i][regIdx] || '').trim().toUpperCase();
      if (!reg) {
        skipped++;
        continue;
      }
      try {
        var before = String(data[i][_toolsAircraftDocsFolderUrlIndex_(sh, headers)] || '').trim();
        var res = _toolsEnsureAircraftDocumentFolderForRow_(sh, headers, i + 1);
        if (!res || !res.success) {
          errors.push(reg + ': ' + String((res && res.error) || 'unknown error'));
          continue;
        }
        if (before) updated++; else created++;
      } catch (rowErr) {
        errors.push(reg + ': ' + String(rowErr && rowErr.message ? rowErr.message : rowErr));
      }
    }

    return {
      success: true,
      created: created,
      updated: updated,
      skipped: skipped,
      errors: errors,
      totalAircraftRows: Math.max(0, data.length - 1)
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _toolsNormalizeAirportCode_(value) {
  return String(value || '').toUpperCase().replace(/[^A-Z0-9]/g, '').trim();
}

function _toolsAirportPhotoRootFolder_() {
  var folderRootName = 'DB_Airports_Airstrip_Photos';
  var existingRoots = DriveApp.getFoldersByName(folderRootName);
  return existingRoots.hasNext() ? existingRoots.next() : DriveApp.createFolder(folderRootName);
}

function _toolsEnsureAirportPhotoFolderForRow_(sheet, headerRow, rowNumber) {
  var headers = Array.isArray(headerRow) ? headerRow.map(function(h) { return _toolsNormHeader_(h); }) : [];
  var icaoIdx = headers.indexOf('ICAO');
  if (icaoIdx < 0) return { success: false, error: 'ICAO column not found in DB_Airports' };

  var row = sheet.getRange(rowNumber, 1, 1, headerRow.length).getValues()[0];
  var icao = _toolsNormalizeAirportCode_(row[icaoIdx]);
  if (!icao) return { success: false, error: 'Airport ICAO is empty' };

  var rootFolder = _toolsAirportPhotoRootFolder_();
  var existing = rootFolder.getFoldersByName(icao);
  var folder = existing.hasNext() ? existing.next() : rootFolder.createFolder(icao);
  var folderUrl = folder.getUrl();

  var photoIdx = headers.indexOf('AIRSTRIP_PHOTO');
  if (photoIdx >= 0) {
    var current = String(row[photoIdx] || '').trim();
    if (current !== folderUrl) {
      sheet.getRange(rowNumber, photoIdx + 1).setValue(folderUrl);
    }
  }

  return {
    success: true,
    icao: icao,
    folderId: folder.getId(),
    folderName: folder.getName(),
    url: folderUrl
  };
}

function ensureAirportPhotoFolderForIcao(icao) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.AIRPORTS, 'ensureAirportPhotoFolderForIcao');
    var headerRow = _toolsSheetHeaderRow_(sh);
    var headers = headerRow.map(function(h) { return _toolsNormHeader_(h); });
    var icaoIdx = headers.indexOf('ICAO');
    if (icaoIdx < 0) throw new Error('ICAO column not found in DB_Airports');

    var target = _toolsNormalizeAirportCode_(icao);
    if (!target) throw new Error('ICAO required');

    var lastRow = sh.getLastRow();
    if (lastRow < 2) throw new Error('DB_Airports is empty');
    var rows = sh.getRange(2, 1, lastRow - 1, headerRow.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (_toolsNormalizeAirportCode_(rows[i][icaoIdx]) === target) {
        return _toolsEnsureAirportPhotoFolderForRow_(sh, headerRow, i + 2);
      }
    }

    throw new Error('Airport ' + target + ' not found in DB_Airports');
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function updateToolsSheetRecord(kind, rowNumber, payload) {
  try {
    var targetRow = Number(rowNumber || 0);
    if (!(targetRow >= 2)) throw new Error('Invalid row number for update');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = _toolsSheetNameFromKind_(kind);
    var sh = getRequiredSheet_(ss, sheetName, 'updateToolsSheetRecord');
    var headerRow = _toolsSheetHeaderRow_(sh);
    var currentRow = sh.getRange(targetRow, 1, 1, headerRow.length).getValues()[0];
    var data = (payload && typeof payload === 'object') ? payload : {};

    var row = headerRow.map(function(header, idx) {
      var label = String(header || '').trim();
      if (!label) return currentRow[idx];
      var key = _toolsNormHeader_(label);
      if (Object.prototype.hasOwnProperty.call(data, key)) return data[key];
      if (Object.prototype.hasOwnProperty.call(data, label)) return data[label];
      return currentRow[idx];
    });

    if (sheetName === APP_SHEETS.ROUTES) {
      var routeHeaders = headerRow.map(function(h) { return _toolsNormHeader_(h); });
      var wpIdx = routeHeaders.indexOf('WAYPOINT_LIST');
      if (wpIdx >= 0) {
        row[wpIdx] = _toolsNormalizeRouteWaypointList_(row[wpIdx]);
      }
    }

    sh.getRange(targetRow, 1, 1, row.length).setValues([row]);
    return { success: true, sheetName: sheetName, rowNumber: targetRow };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function deleteToolsSheetRecord(kind, rowNumber) {
  try {
    var targetRow = Number(rowNumber || 0);
    if (!(targetRow >= 2)) throw new Error('Invalid row number for delete');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = _toolsSheetNameFromKind_(kind);
    var sh = getRequiredSheet_(ss, sheetName, 'deleteToolsSheetRecord');
    var lastRow = sh.getLastRow();
    if (targetRow > lastRow) throw new Error('Row not found for delete');

    sh.deleteRow(targetRow);
    return { success: true, sheetName: sheetName, rowNumber: targetRow };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function searchToolsPassengers(query) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.PASSENGERS, 'searchToolsPassengers');
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: true, records: [] };

    var headers = data[0];
    var normalizedHeaders = headers.map(function(h) { return _toolsNormHeader_(h); });
    var nameIdx = normalizedHeaders.indexOf('PASSENGER_NAME');
    var idIdx = normalizedHeaders.indexOf('ID_NUMBER_CPF_PASSPORT');
    var phoneIdx = normalizedHeaders.indexOf('PHONE');
    var weightIdx = normalizedHeaders.indexOf('WEIGHT_KG') >= 0 ? normalizedHeaders.indexOf('WEIGHT_KG') : normalizedHeaders.indexOf('WEIGHT_KGS');
    var genderIdx = normalizedHeaders.indexOf('GENDER');
    var dobIdx = normalizedHeaders.indexOf('DOB');
    var q = String(query || '').trim().toUpperCase();
    var records = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var hay = [
        nameIdx >= 0 ? row[nameIdx] : '',
        idIdx >= 0 ? row[idIdx] : '',
        phoneIdx >= 0 ? row[phoneIdx] : ''
      ].join(' ').toUpperCase();
      if (q && hay.indexOf(q) === -1) continue;
      records.push({
        rowNumber: i + 1,
        name: nameIdx >= 0 ? String(row[nameIdx] || '').trim() : '',
        id_num: idIdx >= 0 ? String(row[idIdx] || '').trim() : '',
        phone: phoneIdx >= 0 ? String(row[phoneIdx] || '').trim() : '',
        weight: weightIdx >= 0 ? row[weightIdx] : '',
        gender: genderIdx >= 0 ? String(row[genderIdx] || '').trim() : '',
        dob: dobIdx >= 0 ? safeDobStr(row[dobIdx]) : '',
        payload: _toolsRowPayloadFromHeaders_(headers, row)
      });
      if (!q && records.length >= 25) break;
    }

    records.sort(function(a, b) {
      return String(a.name || '').localeCompare(String(b.name || ''));
    });
    return { success: true, records: records };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function addWaypointToDatabase(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getRequiredSheet_(ss, APP_SHEETS.WAYPOINTS, 'addWaypointToDatabase');

  if (!data) throw new Error('Waypoint payload is required');

  const wpId = String(data.wp_id || '').trim().toUpperCase();
  const lat = parseFloat(data.latitude);
  const lon = parseFloat(data.longitude);
  const type = String(data.type || '').trim().toUpperCase();

  if (!wpId) throw new Error('WP_ID is required');
  if (isNaN(lat) || lat < -90 || lat > 90) throw new Error('LATITUDE must be between -90 and 90');
  if (isNaN(lon) || lon < -180 || lon > 180) throw new Error('LONGITUDE must be between -180 and 180');
  if (type !== 'FIX' && type !== 'WATER RUNWAY') throw new Error('TYPE must be FIX or WATER RUNWAY');

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  if (!values || values.length < 1) throw new Error('DB_Waypoints header row not found');

  const headers = values[0].map(function(h) {
    return String(h || '').trim().toUpperCase();
  });
  const idIdx = headers.indexOf('WP_ID');
  const latIdx = headers.indexOf('LATITUDE');
  const lonIdx = headers.indexOf('LONGITUDE');
  const typeIdx = headers.indexOf('TYPE');

  if (idIdx === -1 || latIdx === -1 || lonIdx === -1 || typeIdx === -1) {
    throw new Error('DB_Waypoints must include headers: WP_ID, LATITUDE, LONGITUDE, TYPE');
  }

  for (var i = 1; i < values.length; i++) {
    var existingId = String(values[i][idIdx] || '').trim().toUpperCase();
    if (existingId && existingId === wpId) {
      throw new Error('Waypoint already exists: ' + wpId);
    }
  }

  const newRow = new Array(headers.length).fill('');
  newRow[idIdx] = wpId;
  newRow[latIdx] = lat;
  newRow[lonIdx] = lon;
  newRow[typeIdx] = type;

  sheet.appendRow(newRow);
  return { success: true, wp_id: wpId };
}

function logFuelChange(icao, amount, acft, pilot, flightLegId = "") {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const logSheet = getRequiredSheet_(ss, APP_SHEETS.FUEL_LOGS, 'logFuelChange');
 const cacheSheet = getRequiredSheet_(ss, APP_SHEETS.FUEL_CACHES, 'logFuelChange');


 // Find cache details
 const cacheData = cacheSheet.getDataRange().getValues();
 let airportName = "";
 let fuelType = "";
 for (let i = 1; i < cacheData.length; i++) {
   if (cacheData[i][FUEL_CACHE_COL.ICAO] == icao) {
     airportName = cacheData[i][FUEL_CACHE_COL.LOCATION_NAME];
     fuelType = cacheData[i][FUEL_CACHE_COL.FUEL_TYPE];
     const currentVal = safeNumber_(cacheData[i][FUEL_CACHE_COL.CURRENT_QTY], 0);
     cacheSheet.getRange(i + 1, FUEL_CACHE_COL.CURRENT_QTY + 1).setValue(currentVal + amount);
     break;
   }
 }


 // Log the change with Flight ID
 logSheet.appendRow([
   new Date(),   // TIMESTAMP
   icao,         // ICAO
   airportName,  // AIRPORT_NAME
   acft,         // AIRCRAFT
   pilot,        // PILOT
   amount,       // CHANGE_QTY
   fuelType,     // TYPE
   "No",         // VERIFIED
   flightLegId   // NEW: Flight ID
 ]);
 return true;
}
/**
* Finds a specific mission by its ID to identify which aircraft is assigned.
* Used by the Dashboard to trigger the Tech Status/Squawk lookup.
*/
function reverseFuelForLeg(flightLegId) {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const logSheet = ss.getSheetByName(APP_SHEETS.FUEL_LOGS);
 const cacheSheet = ss.getSheetByName(APP_SHEETS.FUEL_CACHES);
 if (!logSheet || !cacheSheet) return;


 const logData = logSheet.getDataRange().getValues();


 // Loop backwards through logs to find the SPECIFIC leg
 for (let i = logData.length - 1; i >= 1; i--) {
   const rowFlightId = String(logData[i][FUEL_LOG_COL.FLIGHT_ID]);
   const verified = String(logData[i][FUEL_LOG_COL.VERIFIED] || "").toUpperCase().trim();


   // EXACT MATCH ONLY
   if (rowFlightId === flightLegId) {
    
     // Only reverse inventory if NOT verified
     if (verified === "NO" || verified === "") {
       const icao = logData[i][FUEL_LOG_COL.ICAO];
       const amount = safeNumber_(logData[i][FUEL_LOG_COL.CHANGE_QTY], 0);


       // Find the cache and "Refund" the fuel
       const cacheData = cacheSheet.getDataRange().getValues();
       for (let j = 1; j < cacheData.length; j++) {
         if (cacheData[j][FUEL_CACHE_COL.ICAO] === icao) {
           const currentInv = safeNumber_(cacheData[j][FUEL_CACHE_COL.CURRENT_QTY], 0);
           // Reverse: Current - (-96) = Current + 96
           cacheSheet.getRange(j + 1, FUEL_CACHE_COL.CURRENT_QTY + 1).setValue(currentInv - amount);
           break;
         }
       }
      
       // Delete the specific log row
       logSheet.deleteRow(i + 1);
       appLog_("Deleted and reversed fuel for specific leg: " + flightLegId);
     }
   }
 }
}








function getMissionById(missionId) {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, "getMissionById");




const data = sheet.getDataRange().getValues();




// 1. Find all rows for this mission
let missionRows = data.filter(r => String(r[DISPATCH_COL.MISSION_ID]) === String(missionId));
if (missionRows.length === 0) {
  // Maybe they clicked a flightLegId instead
  const legRow = data.find(r => String(r[DISPATCH_COL.FLIGHT_ID]) === String(missionId));
  if (legRow) {
    const realMissionId = legRow[DISPATCH_COL.MISSION_ID];
    missionRows = data.filter(r => String(r[DISPATCH_COL.MISSION_ID]) === String(realMissionId));
  }
}




if (missionRows.length === 0) return null;




const mainRow = missionRows[0];




// 2. Safe Date handling
let rawDate = mainRow[DISPATCH_COL.DATE];
let dateStr = (rawDate instanceof Date) ? rawDate.toISOString().split('T')[0] : String(rawDate);
const missionDateObj = (rawDate instanceof Date) ? rawDate : new Date();




// 3. Build mission object
const missionData = {
  id: mainRow[DISPATCH_COL.MISSION_ID],
  date: dateStr,
  acft: String(mainRow[DISPATCH_COL.AIRCRAFT]),
  pilot: String(mainRow[DISPATCH_COL.PILOT]),
  status: mainRow[DISPATCH_COL.STATUS] ? mainRow[DISPATCH_COL.STATUS].toString().toUpperCase() : "PENDING",
  meta: {
    date: dateStr,
    acft: String(mainRow[DISPATCH_COL.AIRCRAFT]),
    pilot: String(mainRow[DISPATCH_COL.PILOT]),
    copilot: String(mainRow[DISPATCH_COL.COPILOT] || ""),
    type: String(mainRow[DISPATCH_COL.TYPE] || ""),
    notes: String(mainRow[DISPATCH_COL.NOTES] || "")
  },
  // 4. Parse legs
  legs: missionRows.map(r => {
    let legPayload = {};
    try {
      const json = JSON.parse(r[DISPATCH_COL.RAW_DATA] || "{}");
      if (json.legs && Array.isArray(json.legs)) legPayload = json.legs[0];
      else legPayload = json;
    } catch (e) { legPayload = {}; }



    const routeStr = String(r[DISPATCH_COL.ROUTE] || "").trim();
  const parsedRoute = splitRoute_(routeStr);
    const safeNum = (val, def) => isNaN(parseFloat(val)) ? def : parseFloat(val);




    return {
      flightLegId: r[DISPATCH_COL.FLIGHT_ID],
      from: parsedRoute.from || "?",
      to: parsedRoute.to || "?",
      route: routeStr,
      waypoints: legPayload.waypoints || [],
      time: safeNum(r[DISPATCH_COL.FLIGHT_TIME], 0),
      dist: safeNum(legPayload.dist, 0),
      groundTime: safeNum(legPayload.groundTime, 0.5),
      fuel: safeNum(legPayload.fuel, 0),
      takeoffFuel: safeNum(legPayload.takeoffFuel, 0),
      landingFuel: safeNum(legPayload.landingFuel, 0),
      payload: safeNum(legPayload.payload, 0),
      availPayload: safeNum(legPayload.availPayload, 0),
      limit: safeNum(legPayload.limit, 0),
      pax: legPayload.pax || [],
      limitType: legPayload.limitType || "",
      isOver: legPayload.isOver || false,
      missionTime: legPayload.missionTime || "08:00",
      logStatus: 'PENDING',  // enriched below
      bracesRelease: null,
      onBlocks: null
    };
  }),
  actualFuelLogs: getFuelLogsForMission(missionId)
};

// Enrich legs with flight status from LOG_Flights sheet
const logSheetForStatus_ = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
if (logSheetForStatus_) {
  const logData_ = logSheetForStatus_.getDataRange().getValues();
  const logStatusMap_ = {};
  logData_.slice(1).forEach(function(r) {
    const lid = String(r[LOG_FLIGHT_COL.FLIGHT_ID] || '').trim();
    if (!lid) return;
    const released = r[LOG_FLIGHT_COL.BRAKES_RELEASE];
    const onBlks   = r[LOG_FLIGHT_COL.ON_BLOCKS];
    const relStr = (released instanceof Date) ? released.toISOString() : (released ? String(released).trim() : null);
    const obStr  = (onBlks   instanceof Date) ? onBlks.toISOString()   : (onBlks   ? String(onBlks).trim()   : null);
    logStatusMap_[lid] = {
      bracesRelease: relStr || null,
      onBlocks:      obStr  || null,
      logStatus: obStr ? 'COMPLETE' : (relStr ? 'DEPARTED' : 'PENDING')
    };
  });
  missionData.legs.forEach(function(leg) {
    const s = logStatusMap_[leg.flightLegId] || {};
    leg.bracesRelease = s.bracesRelease || null;
    leg.onBlocks      = s.onBlocks      || null;
    leg.logStatus     = s.logStatus     || 'PENDING';
  });
}

return missionData;
}

function getFuelLogsForMission(missionId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(APP_SHEETS.FUEL_LOGS);
  if (!logSheet) return [];
  
  const data = logSheet.getDataRange().getValues();
  if (data.length < 2) return [];

  return data.slice(1).filter(row => {
    // This captures any leg belonging to the mission (e.g., M101 matches M101-A, M101-B)
    return row[FUEL_LOG_COL.FLIGHT_ID] && String(row[FUEL_LOG_COL.FLIGHT_ID]).includes(String(missionId));
  }).map(row => ({
    icao: row[FUEL_LOG_COL.ICAO],
    qty: Math.abs(safeNumber_(row[FUEL_LOG_COL.CHANGE_QTY], 0)), 
    type: row[FUEL_LOG_COL.TYPE],
    verified: row[FUEL_LOG_COL.VERIFIED],
    flightLegId: String(row[FUEL_LOG_COL.FLIGHT_ID])
  }));
}
function submitBriefingToLog(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  if (!logSheet) throw new Error("Sheet 'LOG_Flights' not found.");

  // Construct the row to match your 25-column structure exactly
  const newRow = [
    payload.flightLegId,     // Flight_ID (now using the first leg's ID like ADS26-001-01)
    payload.date,            // Date
    payload.pilot,           // Pilot
    payload.acft,            // Aircraft
    payload.from,            // From
    payload.to,              // To
    payload.totalDist,       // Distance_NM
    payload.startTach,       // Start_Tach
    "",                      // End_Tach (Empty until debrief)
    "",                      // Total_Time (Empty until debrief)
    payload.fuelTotal,       // Fuel_Start
    "",                      // Fuel_End (Empty until debrief)
    "",                      // Fuel_Used (Empty until debrief)
    payload.oil,             // Oil_Added
    payload.volts,           // Battery_Volts
    "",                      // Squawks
    payload.riskMatrix || "",// TO Risk Matrix
    "",                      // Brakes Release (Empty until takeoff)
    payload.actualLoadJSON,  // Actual_Load_JSON
    "",                      // Landing Risk Matrix (Empty until debrief)
    "",                      // Number_Ldgs (Empty until debrief)
    "",                      // Airborne (Empty until takeoff)
    "",                      // Landed (Empty until debrief)
    "",                      // Brakes Applied (Empty until debrief)
    ""                       // Actual TO Roll (Empty until takeoff)
  ];

  logSheet.appendRow(newRow);
  
  // Return the row index or true so the UI knows it worked
  return true;
}

// Backwards-compatible wrapper: client calls saveMissionToLog
function saveMissionToLog(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  if (!logSheet) throw new Error("Sheet 'LOG_Flights' not found.");

  const data = logSheet.getDataRange().getValues();
  if (!data || data.length < 2) {
    return submitBriefingToLog(payload);
  }

  const flightLegId = String(payload && payload.flightLegId || '').trim();
  if (!flightLegId) {
    return submitBriefingToLog(payload);
  }

  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][LOG_FLIGHT_COL.FLIGHT_ID] || '').trim() === flightLegId) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow < 0) {
    return submitBriefingToLog(payload);
  }

  const existing = data[targetRow - 1].slice();
  const setIfProvided = (colIdx, val) => {
    if (val === undefined || val === null || val === '') return;
    existing[colIdx] = val;
  };

  setIfProvided(LOG_FLIGHT_COL.DATE, payload.date);
  setIfProvided(LOG_FLIGHT_COL.PILOT, payload.pilot);
  setIfProvided(LOG_FLIGHT_COL.ACFT, payload.acft);
  setIfProvided(LOG_FLIGHT_COL.FROM, payload.from);
  setIfProvided(LOG_FLIGHT_COL.TO, payload.to);
  setIfProvided(LOG_FLIGHT_COL.DIST, payload.totalDist);
  setIfProvided(LOG_FLIGHT_COL.START_TACH, payload.startTach);
  setIfProvided(LOG_FLIGHT_COL.FUEL_START, payload.fuelTotal);
  setIfProvided(LOG_FLIGHT_COL.OIL, payload.oil);
  setIfProvided(LOG_FLIGHT_COL.VOLTS, payload.volts);
  setIfProvided(LOG_FLIGHT_COL.ACTUAL_LOAD_JSON, payload.actualLoadJSON);

  logSheet.getRange(targetRow, 1, 1, existing.length).setValues([existing]);
  return { success: true, updated: true, row: targetRow, flightLegId: flightLegId };
}

function getFlightLogData(flightLegId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  if (!logSheet) return null;

  const data = logSheet.getDataRange().getValues();
  if (data.length < 2) return null;

  const headerRow = data[0];
  const flightIdCol = headerRow.indexOf('Flight_ID');
  const norm = function(v) { return String(v || '').toUpperCase().trim().replace(/\s+/g, '_'); };
  const headers = headerRow.map(norm);
  const col = function(name, fallback) {
    const idx = headers.indexOf(name);
    return idx >= 0 ? idx : fallback;
  };

  const actualLoadCol = col('ACTUAL_LOAD_JSON', LOG_FLIGHT_COL.ACTUAL_LOAD_JSON);
  const startTachCol = col('START_TACH', LOG_FLIGHT_COL.START_TACH);
  const endTachCol = col('END_TACH', LOG_FLIGHT_COL.END_TACH);
  const fuelStartCol = col('FUEL_START', LOG_FLIGHT_COL.FUEL_START);
  const fuelEndCol = col('FUEL_END', LOG_FLIGHT_COL.FUEL_END);
  const fuelUsedCol = col('FUEL_USED', LOG_FLIGHT_COL.FUEL_USED);
  const brakesReleaseCol = col('BRAKES_RELEASE', LOG_FLIGHT_COL.BRAKES_RELEASE);
  const onBlocksCol = col('ON_BLOCKS', LOG_FLIGHT_COL.ON_BLOCKS);
  const airborneCol = col('AIRBORNE', LOG_FLIGHT_COL.AIRBORNE);
  const landedCol = col('LANDED', LOG_FLIGHT_COL.LANDED);
  const brakesAppliedCol = col('BRAKES_APPLIED', LOG_FLIGHT_COL.BRAKES_APPLIED);
  const numLdgsCol = col('NUMBER_LDGS', LOG_FLIGHT_COL.NUM_LDGS);
  const numTgCol = col('NUMBER_TOUCH_AND_GOS', -1) >= 0 ? col('NUMBER_TOUCH_AND_GOS', -1) : col('NUM_TOUCH_AND_GOS', -1);
  const squawksCol = col('SQUAWKS', LOG_FLIGHT_COL.SQUAWKS);

  if (flightIdCol < 0) return null;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][flightIdCol]).trim() === String(flightLegId).trim()) {
      return {
        flightLegId: data[i][flightIdCol],
        startTach: startTachCol >= 0 ? data[i][startTachCol] : '',
        endTach: endTachCol >= 0 ? data[i][endTachCol] : '',
        fuelStart: fuelStartCol >= 0 ? data[i][fuelStartCol] : '',
        fuelEnd: fuelEndCol >= 0 ? data[i][fuelEndCol] : '',
        fuelUsed: fuelUsedCol >= 0 ? data[i][fuelUsedCol] : '',
        brakesRelease: brakesReleaseCol >= 0 ? data[i][brakesReleaseCol] : '',
        onBlocks: onBlocksCol >= 0 ? data[i][onBlocksCol] : '',
        airborne: airborneCol >= 0 ? data[i][airborneCol] : '',
        landed: landedCol >= 0 ? data[i][landedCol] : '',
        brakesApplied: brakesAppliedCol >= 0 ? data[i][brakesAppliedCol] : '',
        numLdgs: numLdgsCol >= 0 ? data[i][numLdgsCol] : '',
        numTouchAndGos: numTgCol >= 0 ? data[i][numTgCol] : '',
        squawks: squawksCol >= 0 ? data[i][squawksCol] : '',
        actualLoadJSON: actualLoadCol >= 0 ? data[i][actualLoadCol] : ''
      };
    }
  }

  return null;
}

function normalizeWaypointList_(rawWaypoints, origin, destination) {
  var org = String(origin || '').trim().toUpperCase();
  var dst = String(destination || '').trim().toUpperCase();
  var tokens = [];

  if (Array.isArray(rawWaypoints)) {
    tokens = rawWaypoints.map(function(wp) {
      return String(wp || '').trim().toUpperCase();
    });
  } else if (typeof rawWaypoints === 'string') {
    var raw = String(rawWaypoints || '').trim().toUpperCase();
    if (raw) {
      raw = raw.replace(/[→>]/g, ',');
      tokens = raw.split(/[\n\r,;\/|]+/).map(function(part) {
        return String(part || '').trim().toUpperCase();
      });
    }
  }

  var seen = {};
  return tokens.filter(function(token) {
    if (!token || token === org || token === dst || seen[token]) return false;
    seen[token] = true;
    return true;
  });
}

function routeWaypointsFromRouteString_(routeValue) {
  var tokens = routeTokensFromString_(routeValue || '');
  if (tokens.length <= 2) {
    var raw = String(routeValue || '').trim().toUpperCase();
    // Legacy fallback for "AAA - BBB - CCC" only.
    if (/\s+-\s+/.test(raw)) {
      tokens = raw
        .split(/\s+-\s+/)
        .map(function(part) { return String(part || '').trim().toUpperCase(); })
        .filter(Boolean);
    }
  }
  if (tokens.length <= 2) return [];
  return tokens.slice(1, tokens.length - 1);
}

function getFlightRouteData(flightLegId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dispatchSheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, 'getFlightRouteData');

  const data = dispatchSheet.getDataRange().getValues();
  if (data.length < 2) return null;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][DISPATCH_COL.FLIGHT_ID]).trim() === String(flightLegId).trim()) {
      let routeData = null;
      if (data[i][DISPATCH_COL.RAW_DATA]) {
        const raw = safeJsonParse_(data[i][DISPATCH_COL.RAW_DATA], null);
        if (!raw) {
          appLog_('Failed to parse RAW_DATA for route: flightLegId=', flightLegId);
        } else {
          const routeText = String(data[i][DISPATCH_COL.ROUTE] || '');
          const fallbackRoute = splitRoute_(data[i][DISPATCH_COL.ROUTE]);
          const fallbackFrom = fallbackRoute.from || '';
          const fallbackTo = fallbackRoute.to || '';
          const fallbackWps = routeWaypointsFromRouteString_(routeText);

          if (raw.legs && Array.isArray(raw.legs) && raw.legs[0]) {
            const leg = raw.legs[0];
            const from = leg.from || fallbackFrom;
            const to = leg.to || fallbackTo;
            const wpSource = (Array.isArray(fallbackWps) && fallbackWps.length)
              ? fallbackWps
              : (leg.waypoints || fallbackWps);
            routeData = {
              from: from,
              to: to,
              waypoints: normalizeWaypointList_(wpSource, from, to)
            };
          } else if (raw.waypoints) {
            const from = raw.from || fallbackFrom;
            const to = raw.to || fallbackTo;
            routeData = {
              from: from,
              to: to,
              waypoints: normalizeWaypointList_(raw.waypoints || [], from, to)
            };
          }
        }
      }
      
      if (!routeData) {
        const routeText = String(data[i][DISPATCH_COL.ROUTE] || '');
        const fallbackRoute = splitRoute_(data[i][DISPATCH_COL.ROUTE]);
        routeData = {
          from: fallbackRoute.from || '',
          to: fallbackRoute.to || '',
          waypoints: normalizeWaypointList_(routeWaypointsFromRouteString_(routeText), fallbackRoute.from || '', fallbackRoute.to || '')
        };
      }

      // If waypoints are still empty, look up DB_Routes by origin/destination
      if (routeData && routeData.from && (!routeData.waypoints || !routeData.waypoints.length)) {
        try {
          const routeSheet = ss.getSheetByName(APP_SHEETS.ROUTES);
          if (routeSheet) {
            const routeVals = routeSheet.getDataRange().getValues();
            if (routeVals.length > 1) {
              const rHeaders = routeVals[0].map(function(h) { return String(h || '').trim().toUpperCase(); });
              const oriIdx = rHeaders.indexOf('ORIGIN');
              const dstIdx = rHeaders.indexOf('DESTINATION');
              const wpListIdx = rHeaders.indexOf('WAYPOINT_LIST');
              if (oriIdx >= 0 && dstIdx >= 0 && wpListIdx >= 0) {
                for (var ri = 1; ri < routeVals.length; ri++) {
                  const rOri = String(routeVals[ri][oriIdx] || '').trim().toUpperCase();
                  const rDst = String(routeVals[ri][dstIdx] || '').trim().toUpperCase();
                  const isMatch = (rOri === routeData.from && rDst === routeData.to)
                                || (rOri === routeData.to   && rDst === routeData.from);
                  if (isMatch) {
                    const rawWpList = String(routeVals[ri][wpListIdx] || '').trim();
                    const dbTokens = rawWpList.split(/[,;]+/).map(function(s) { return s.trim().toUpperCase(); }).filter(Boolean);
                    if (dbTokens.length) {
                      routeData.waypoints = normalizeWaypointList_(dbTokens, routeData.from, routeData.to);
                      appLog_('getFlightRouteData: DB_Routes waypoints used for', routeData.from, '->', routeData.to, ':', dbTokens.join(','));
                    }
                    break;
                  }
                }
              }
            }
          }
        } catch (e) {
          appLog_('getFlightRouteData DB_Routes lookup error:', e && e.message);
        }
      }

      return routeData;
    }
  }

  return null;
}

function _getChatWebhookUrl(explicitUrl) {
  if (explicitUrl && String(explicitUrl).trim()) return String(explicitUrl).trim();

  const props = PropertiesService.getScriptProperties();
  const keys = [
    'GOOGLE_CHAT_WEBHOOK_URL',
    'CHAT_WEBHOOK_URL',
    'DISPATCH_CHAT_WEBHOOK'
  ];

  for (let i = 0; i < keys.length; i++) {
    const v = props.getProperty(keys[i]);
    if (v && String(v).trim()) return String(v).trim();
  }
  return '';
}

function _sendDispatchPreviewToChat(messageText, explicitWebhookUrl) {
  const webhookUrl = _getChatWebhookUrl(explicitWebhookUrl);
  if (!webhookUrl) {
    return { sent: false, reason: 'Webhook not configured (set GOOGLE_CHAT_WEBHOOK_URL in Script Properties).' };
  }

  const payload = { text: String(messageText || '').trim() || 'Dispatch release triggered.' };
  const resp = UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  if (code >= 200 && code < 300) {
    return { sent: true, code: code };
  }

  return {
    sent: false,
    code: code,
    reason: String(resp.getContentText() || 'Non-2xx response from Google Chat webhook.')
  };
}

function sendDispatchEmail(payload) {
  // Accept payload.emails (array) or legacy payload.email (string)
  var recipients;
  if (payload && Array.isArray(payload.emails) && payload.emails.length) {
    recipients = payload.emails.join(',');
  } else {
    recipients = String(payload && payload.email || '').trim();
  }
  var message = String(payload && payload.message || '').trim();
  var subject = String(payload && payload.subject || 'Dispatch Release').trim();
  if (!recipients) throw new Error('sendDispatchEmail: at least one recipient is required');
  if (!message) throw new Error('sendDispatchEmail: message is required');
  MailApp.sendEmail({
    to: recipients,
    subject: subject,
    body: message
  });
  return { success: true };
}

// ── Flight Following ──────────────────────────────────────────────────────────
var FF_RECIPIENTS_ = ['acompanhamento@asasdesocorro.org.br', 'supervisor.voo@asasdesocorro.org.br'];

function sendFlightFollowEvent(payload) {
  // payload: { event, reg, type, pic, route, pob, depTime, eta, notes, confirmedBy }
  var event       = String(payload.event       || '').toUpperCase(); // AIRBORNE | POSREP | LANDED | OVERDUE
  var reg         = String(payload.reg         || '').trim().toUpperCase();
  var acftType    = String(payload.type        || '').trim();
  var pic         = String(payload.pic         || '—').trim();
  var route       = String(payload.route       || '—').trim();
  var pob         = String(payload.pob         || '—').trim();
  var depTime     = String(payload.depTime     || '').trim();
  var eta         = String(payload.eta         || '').trim();
  var notes       = String(payload.notes       || '').trim();
  var confirmedBy = String(payload.confirmedBy || '—').trim();

  if (!reg)   throw new Error('sendFlightFollowEvent: reg required');
  if (!event) throw new Error('sendFlightFollowEvent: event required');

  var now       = new Date();
  var tz        = 'America/Manaus'; // MAO local — adjust if needed
  var localTime = Utilities.formatDate(now, tz, 'HH:mm');
  var zuluTime  = Utilities.formatDate(now, 'UTC', 'HHmm') + 'Z';

  var icon = { AIRBORNE: '✈', POSREP: '📍', LANDED: '🛬', OVERDUE: '⚠' }[event] || '•';

  var subject = icon + ' ' + event + ' ' + reg + ' ' + zuluTime;
  if (route) subject += ' ' + route;

  var lines = [
    icon + ' FLIGHT FOLLOWING — ' + event,
    'Aircraft : ' + reg + (acftType ? ' (' + acftType + ')' : ''),
    'PIC      : ' + pic,
    'Route    : ' + route,
    'POB      : ' + pob,
    'Time     : ' + localTime + ' MAO / ' + zuluTime,
  ];
  if (depTime) lines.push('Dep      : ' + depTime);
  if (eta)     lines.push('ETA      : ' + eta);
  if (notes)   lines.push('');
  if (notes)   lines.push(notes);
  lines.push('');
  lines.push('Confirmed by: ' + confirmedBy);

  var fullBody = lines.join('\n');

  // Earthmate-compact variant (no emoji, no labels)
  var compact = event + ' ' + reg + ' ' + zuluTime +
    ' PIC:' + pic + ' ROUTE:' + route + ' POB:' + pob +
    (eta ? ' ETA:' + eta : '') +
    (notes ? ' ' + notes.replace(/\n/g, ' ') : '') +
    ' BY:' + confirmedBy;

  var recipients = FF_RECIPIENTS_.join(',');
  MailApp.sendEmail({ to: recipients, subject: subject, body: fullBody });

  return { success: true, compact: compact, subject: subject };
}

function getFlightFollowMissionsForAcft(reg) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(APP_SHEETS.DISPATCH);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  // Build airport fuel lookup map (icao → hasFuel boolean)
  var airportFuelMap = {};
  try {
    var airSheet = ss.getSheetByName(APP_SHEETS.AIRPORTS);
    if (airSheet) {
      var airData = airSheet.getDataRange().getValues();
      if (airData.length > 1) {
        var airH = airData[0].map(function(h) { return String(h || '').toUpperCase().trim().replace(/\s+/g, '_'); });
        var icaoIdx = airH.indexOf('ICAO');
        var fuelIdx = airH.indexOf('FUEL_AVAILABLE');
        if (icaoIdx >= 0 && fuelIdx >= 0) {
          for (var ai = 1; ai < airData.length; ai++) {
            var aIcao = String(airData[ai][icaoIdx] || '').trim().toUpperCase();
            var aFuel = String(airData[ai][fuelIdx] || '').trim().toUpperCase();
            if (aIcao) {
              airportFuelMap[aIcao] = !!(aFuel && aFuel !== 'NONE' && aFuel !== 'N' && aFuel !== 'NO' && aFuel !== '0');
            }
          }
        }
      }
    }
  } catch(e) {}

  // Gather rows for this registration, today ± 1 day (catches night departures)
  var now     = new Date();
  var todayBsb = Utilities.formatDate(now, 'America/Sao_Paulo', 'yyyy-MM-dd');
  var tz      = 'America/Sao_Paulo';

  var missions = {};

  for (var i = 1; i < data.length; i++) {
    var row   = data[i];
    var acft  = String(row[DISPATCH_COL.AIRCRAFT] || '').trim().toUpperCase();
    if (acft !== String(reg || '').trim().toUpperCase()) continue;

    var rawDate = row[DISPATCH_COL.DATE];
    var dateStr = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, tz, 'yyyy-MM-dd')
      : String(rawDate || '').trim().slice(0, 10);
    if (dateStr !== todayBsb) continue;

    var missionId  = String(row[DISPATCH_COL.MISSION_ID] || '').trim();
    var flightLegId = String(row[DISPATCH_COL.FLIGHT_ID] || '').trim();
    if (!missionId || !flightLegId) continue;

    // Parse RAW_DATA blob
    var raw = {};
    try { raw = JSON.parse(String(row[DISPATCH_COL.RAW_DATA] || '{}')); } catch(e) {}

    var legs = Array.isArray(raw.legs) ? raw.legs : [];
    var leg  = legs[0] || {};

    var pax = Array.isArray(leg.pax) ? leg.pax : [];
    var paxFiltered = pax.filter(function(p) {
      return p && String(p.name || '').toUpperCase() !== 'FREIGHT';
    });
    var pobCount = 1 + (String(row[DISPATCH_COL.COPILOT] || '').trim() ? 1 : 0) + paxFiltered.length;

    // Flight plan
    var planId      = String(leg.planId || leg.planDI || raw.planId || '').trim().toUpperCase();
    var takeoffUTC  = String(leg.takeoffUTC || leg.takeoffZulu || raw.takeoffUTC || raw.time || '').trim().replace(/[^0-9]/g,'').slice(0,4);
    var noPlan      = !!(leg.noPlan || raw.noPlan);

    // Waypoints
    var routeStr = String(row[DISPATCH_COL.ROUTE] || '').trim();
    var rawWps   = Array.isArray(leg.waypoints) ? leg.waypoints : [];
    var parsedRoute = splitRoute_(routeStr);
    var from     = String(leg.from || parsedRoute.from || '').trim().toUpperCase();
    var to       = String(leg.to   || parsedRoute.to   || '').trim().toUpperCase();
    var waypoints = [];
    if (rawWps.length > 0) {
      waypoints = rawWps.map(function(w) {
        // Legacy stored waypoints can be plain strings (e.g. ["SBAA","WP1","SBAE"])
        var fix = typeof w === 'string'
          ? w.trim().toUpperCase()
          : String(w && (w.fix || w.name || w.icao) || '').trim().toUpperCase();
        var distNm = (w && typeof w === 'object') ? Number(w.distNm || 0) : 0;
        return { fix: fix, distNm: distNm };
      }).filter(function(w) { return !!w.fix; });
    }
    if (!waypoints.length) {
      // Build waypoints from route string using comma-delimited policy.
      var parts = routeTokensFromString_(routeStr);
      waypoints = parts.map(function(p) { return { fix: p, distNm: 0 }; });
    }
    // Ensure origin and destination are present
    if (!waypoints.length || waypoints[0].fix !== from) waypoints.unshift({ fix: from, distNm: 0 });
    if (waypoints[waypoints.length - 1].fix !== to)    waypoints.push({ fix: to, distNm: 0 });

    // Annotate each waypoint with fuel availability from airports DB
    waypoints = waypoints.map(function(w) {
      var hasFuel = Object.prototype.hasOwnProperty.call(airportFuelMap, w.fix)
        ? airportFuelMap[w.fix] : null; // null = unknown (not an airport in DB)
      return { fix: w.fix, distNm: w.distNm, hasFuel: hasFuel };
    });

    // Planned fuel
    var fuel = Number(leg.fuel || leg.plannedFuel || raw.fuel || 0);

    if (!missions[missionId]) {
      missions[missionId] = {
        missionId:   missionId,
        flightLegId: flightLegId,
        date:        dateStr,
        reg:         acft,
        pilot:       String(row[DISPATCH_COL.PILOT]   || '').trim(),
        copilot:     String(row[DISPATCH_COL.COPILOT] || '').trim(),
        from:        from,
        to:          to,
        routeStr:    routeStr,
        pob:         pobCount,
        pax:         paxFiltered.map(function(p) { return { name: String(p.name || ''), phone: String(p.phone || ''), sex: String(p.sex || ''), age: String(p.age || ''), emergencyContact: String(p.emergencyContact || '') }; }),
        planId:      planId,
        takeoffUTC:  takeoffUTC,
        noPlan:      noPlan,
        fuelL:       fuel,
        waypoints:   waypoints,
        status:      String(row[DISPATCH_COL.STATUS] || '').trim()
      };
    }
  }

  return Object.values(missions);
}

function getFlightFollowMessages(reg) {
  // Read InReach device replies from Gmail for the given registration
  // Returns array of message objects: { timestamp, from, text, verified }
  // In production, this would parse emails from inreach devices
  try {
    var threads = GmailApp.search('label:flight-follow from:inreachmail.com');
    var messages = [];
    
    threads.slice(0, 20).forEach(function(thread) {
      var msgs = thread.getMessages();
      msgs.forEach(function(msg) {
        messages.push({
          timestamp: msg.getDate().toISOString(),
          from: msg.getFrom(),
          subject: msg.getSubject(),
          text: msg.getPlainBody().substring(0, 500),
          verified: false
        });
      });
    });
    
    return messages;
  } catch(e) {
    return [];
  }
}

function getFlightFollowInit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('DB_Aircraft');
  if (!sheet) return { aircraft: [] };
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { aircraft: [] };
  var headers = data[0].map(function(h) { return String(h || '').trim().toUpperCase(); });
  var regIdx  = headers.indexOf('REGISTRATION');
  var typeIdx = headers.indexOf('AIRCRAFT_TYPE');
  if (regIdx < 0) return { aircraft: [] };
  var aircraft = [];
  for (var i = 1; i < data.length; i++) {
    var reg = String(data[i][regIdx] || '').trim();
    if (!reg) continue;
    aircraft.push({
      reg:  reg,
      type: typeIdx >= 0 ? String(data[i][typeIdx] || '').trim() : ''
    });
  }
  return { aircraft: aircraft };
}

function saveMissionFplToDrive(missionId, fplXml) {
  var folderName = 'MissionFlightPlans';
  var folder;
  try {
    var it = DriveApp.getFoldersByName(folderName);
    folder = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);
  } catch(e) {
    folder = DriveApp.getRootFolder();
  }

  var safeName = String(missionId || 'plan').replace(/[^A-Za-z0-9_-]/g, '_');
  var fileName = 'mission_' + safeName + '.fpl';

  // Remove previous version
  try {
    var old = folder.getFilesByName(fileName);
    while (old.hasNext()) { old.next().setTrashed(true); }
  } catch(e2) {}

  var file = folder.createFile(fileName, String(fplXml || ''), 'application/xml');
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    ok: true,
    fileId: file.getId(),
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId()
  };
}

function createDispatchAirportPhotoFolders(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(APP_SHEETS.AIRPORTS) || ss.getSheetByName('DB_Airports');
  if (!sh) throw new Error("Sheet 'DB_Airports' not found.");

  const data = sh.getDataRange().getValues();
  if (data.length < 2) {
    return { success: true, message: 'No airport rows to process', airports: 0, rowsUpdated: 0 };
  }

  const headers = data[0].map(function(h) {
    return String(h || '').toUpperCase().trim().replace(/\s+/g, '_');
  });

  const idx = {
    airstripPhoto: headers.indexOf('AIRSTRIP_PHOTO'),
    icao: headers.indexOf('ICAO')
  };

  if (idx.airstripPhoto < 0) {
    throw new Error('AIRSTRIP_PHOTO column not found in DB_Airports');
  }
  if (idx.icao < 0) {
    throw new Error('ICAO column not found in DB_Airports');
  }

  const normalizeCode = function(v) {
    return String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, '').trim();
  };

  const folderRootName = String(options && options.rootFolderName || 'DB_Airports_Airstrip_Photos').trim();
  const existingRoots = DriveApp.getFoldersByName(folderRootName);
  const rootFolder = existingRoots.hasNext() ? existingRoots.next() : DriveApp.createFolder(folderRootName);

  // Collect unique ICAO codes from the sheet
  const airportSet = {};
  for (let r = 1; r < data.length; r++) {
    const code = normalizeCode(data[r][idx.icao]);
    if (code && code.length >= 3) airportSet[code] = true;
  }

  const airportCodes = Object.keys(airportSet).sort();
  const folderByIcao = {};

  airportCodes.forEach(function(code) {
    const it = rootFolder.getFoldersByName(code);
    folderByIcao[code] = it.hasNext() ? it.next() : rootFolder.createFolder(code);
  });

  let rowsUpdated = 0;
  for (let i = 1; i < data.length; i++) {
    const code = normalizeCode(data[i][idx.icao]);
    if (!code || !folderByIcao[code]) continue;

    const url = folderByIcao[code].getUrl();
    const current = String(data[i][idx.airstripPhoto] || '').trim();
    if (current !== url) {
      sh.getRange(i + 1, idx.airstripPhoto + 1).setValue(url);
      rowsUpdated++;
    }
  }

  return {
    success: true,
    rootFolderName: folderRootName,
    rootFolderId: rootFolder.getId(),
    airports: airportCodes.length,
    rowsUpdated: rowsUpdated,
    sampleAirports: airportCodes.slice(0, 12)
  };
}

function releaseBrakes(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  const dispatchSheet = ss.getSheetByName(APP_SHEETS.DISPATCH);

  if (!logSheet) throw new Error("Sheet 'LOG_Flights' not found.");
  if (!dispatchSheet) throw new Error("Sheet 'DB_Dispatch' not found.");

  const flightLegId = String(payload && payload.flightLegId || '').trim();
  if (!flightLegId) throw new Error('releaseBrakes: flightLegId is required');

  let missionId = String(payload && payload.missionId || '').trim();
  if (!missionId) {
    const parts = flightLegId.split('-');
    missionId = parts.length >= 2 ? (parts[0] + '-' + parts[1]) : flightLegId;
  }

  const riskTotal = parseInt(payload && payload.riskTotal, 10);
  if (isNaN(riskTotal)) throw new Error('releaseBrakes: riskTotal is required');

  const now = new Date();

  // LOG_Flights: Column Q = 17 (TO Risk Matrix), Column R = 18 (Brakes Release)
  const logData = logSheet.getDataRange().getValues();
  let logRow = -1;
  for (let i = 1; i < logData.length; i++) {
    if (String(logData[i][LOG_FLIGHT_COL.FLIGHT_ID] || '').trim() === flightLegId) {
      logRow = i + 1;
      break;
    }
  }
  if (logRow < 0) {
    throw new Error('releaseBrakes: flight not found in LOG_Flights: ' + flightLegId);
  }

  logSheet.getRange(logRow, LOG_FLIGHT_COL.TO_RISK_MATRIX + 1).setValue(riskTotal);
  logSheet.getRange(logRow, LOG_FLIGHT_COL.BRAKES_RELEASE + 1).setValue(now);

  // DB_Dispatch: Column L = 12 status => In-Flight (all rows in mission)
  const dispatchData = dispatchSheet.getDataRange().getValues();
  let dispatchRowsUpdated = 0;
  for (let j = 1; j < dispatchData.length; j++) {
    if (String(dispatchData[j][DISPATCH_COL.MISSION_ID] || '').trim() === missionId) {
      dispatchSheet.getRange(j + 1, DISPATCH_COL.STATUS + 1).setValue('In-Flight');
      dispatchRowsUpdated++;
    }
  }
  CacheService.getScriptCache().remove('scheduledMissions:v1');

  const chatResult = _sendDispatchPreviewToChat(
    payload && payload.chatMessage,
    payload && payload.webhookUrl
  );

  return {
    success: true,
    missionId: missionId,
    flightLegId: flightLegId,
    riskTotal: riskTotal,
    brakesRelease: now.toISOString(),
    dispatchRowsUpdated: dispatchRowsUpdated,
    chat: chatResult
  };
}

function sendEnroutePositionReport(payload) {
  const missionId = String(payload && payload.missionId || '').trim();
  const flightLegId = String(payload && payload.flightLegId || '').trim();
  const text = String(payload && payload.text || '').trim();

  if (!text) throw new Error('sendEnroutePositionReport: text is required');

  const prefix = [
    '📡 ENROUTE POSITION REPORT',
    missionId ? `Mission: ${missionId}` : '',
    flightLegId ? `Leg: ${flightLegId}` : ''
  ].filter(Boolean).join('\n');

  const finalMessage = `${prefix}\n${text}`;
  const chatResult = _sendDispatchPreviewToChat(finalMessage, payload && payload.webhookUrl);

  if (!chatResult.sent) {
    throw new Error('sendEnroutePositionReport: ' + (chatResult.reason || 'Chat send failed'));
  }

  return {
    success: true,
    missionId: missionId,
    flightLegId: flightLegId,
    sentAt: new Date().toISOString(),
    chat: chatResult
  };
}

function recordArrivalOnBlocks(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  if (!logSheet) throw new Error("Sheet 'LOG_Flights' not found.");

  const flightLegId = String(payload && payload.flightLegId || '').trim();
  if (!flightLegId) throw new Error('recordArrivalOnBlocks: flightLegId is required');

  const riskTotal = parseInt(payload && payload.riskTotal, 10);
  if (isNaN(riskTotal)) throw new Error('recordArrivalOnBlocks: riskTotal is required');

  const data = logSheet.getDataRange().getValues();
  if (!data || data.length < 2) throw new Error('recordArrivalOnBlocks: LOG_Flights has no data rows');

  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][LOG_FLIGHT_COL.FLIGHT_ID] || '').trim() === flightLegId) {
      rowIdx = i + 1;
      break;
    }
  }
  if (rowIdx < 0) throw new Error('recordArrivalOnBlocks: flight not found in LOG_Flights: ' + flightLegId);

  const now = new Date();

  // Column T = 20 (Landing Risk Matrix)
  // Column U = 21 (On Blocks)
  logSheet.getRange(rowIdx, LOG_FLIGHT_COL.LANDING_RISK_MATRIX + 1).setValue(riskTotal);
  logSheet.getRange(rowIdx, LOG_FLIGHT_COL.ON_BLOCKS + 1).setValue(now);

  const arrivalJson = payload && payload.arrivalJson ? payload.arrivalJson : null;
  if (arrivalJson) {
    const existingRaw = data[rowIdx - 1][LOG_FLIGHT_COL.ACTUAL_LOAD_JSON] || '';
    let existing = {};
    try {
      existing = existingRaw ? JSON.parse(existingRaw) : {};
    } catch (e) {
      existing = {};
    }

    const merged = {
      ...existing,
      arrival: arrivalJson,
      arrivalSavedAt: now.toISOString()
    };

    // Column S = 19 (Actual_Load_JSON)
    logSheet.getRange(rowIdx, LOG_FLIGHT_COL.ACTUAL_LOAD_JSON + 1).setValue(JSON.stringify(merged));
  }

  return {
    success: true,
    flightLegId: flightLegId,
    riskTotal: riskTotal,
    onBlocks: now.toISOString()
  };
}

function recordDebriefLog(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  if (!logSheet) throw new Error("Sheet 'LOG_Flights' not found.");

  const flightLegId = String(payload && payload.flightLegId || '').trim();
  if (!flightLegId) throw new Error('recordDebriefLog: flightLegId is required');

  const data = logSheet.getDataRange().getValues();
  if (!data || data.length < 2) throw new Error('recordDebriefLog: LOG_Flights has no data rows');

  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][LOG_FLIGHT_COL.FLIGHT_ID] || '').trim() === flightLegId) {
      rowIdx = i + 1;
      break;
    }
  }
  if (rowIdx < 0) throw new Error('recordDebriefLog: flight not found in LOG_Flights: ' + flightLegId);

  const headers = data[0].map(function(h) {
    return String(h || '').toUpperCase().trim().replace(/\s+/g, '_');
  });
  const col = function(name, fallback) {
    const idx = headers.indexOf(name);
    return idx >= 0 ? idx : fallback;
  };

  const colEndTach = col('END_TACH', LOG_FLIGHT_COL.END_TACH);
  const colFuelEnd = col('FUEL_END', LOG_FLIGHT_COL.FUEL_END);
  const colFuelUsed = col('FUEL_USED', LOG_FLIGHT_COL.FUEL_USED);
  const colAirborne = col('AIRBORNE', LOG_FLIGHT_COL.AIRBORNE);
  const colLanded = col('LANDED', LOG_FLIGHT_COL.LANDED);
  const colNumLdgs = col('NUMBER_LDGS', LOG_FLIGHT_COL.NUM_LDGS);
  const colBrakesApplied = col('BRAKES_APPLIED', LOG_FLIGHT_COL.BRAKES_APPLIED);
  const colTotalTime = col('TOTAL_TIME', LOG_FLIGHT_COL.TOTAL_TIME);
  const colNumTg = col('NUMBER_TOUCH_AND_GOS', -1) >= 0 ? col('NUMBER_TOUCH_AND_GOS', -1) : col('NUM_TOUCH_AND_GOS', -1);
  const colSquawks = col('SQUAWKS', LOG_FLIGHT_COL.SQUAWKS);
  const colActualLoadJson = col('ACTUAL_LOAD_JSON', LOG_FLIGHT_COL.ACTUAL_LOAD_JSON);

  const endTach = String(payload && payload.endTach || '').trim();
  const fuelEnd = parseFloat(payload && payload.fuelEnd) || 0;
  const fuelUsed = parseFloat(payload && payload.fuelUsed) || 0;
  const airborne = String(payload && payload.airborne || '').trim();
  const landed = String(payload && payload.landed || '').trim();
  const brakesApplied = String(payload && payload.brakesApplied || '').trim();
  const numLdgs = parseInt(payload && payload.numLdgs, 10) || 1;
  const numTouchAndGos = Math.max(0, parseInt(payload && payload.numTouchAndGos, 10) || 0);
  const squawks = String(payload && payload.squawks || '')
    .split(/[\n,;]+/)
    .map(function(s) { return String(s || '').trim(); })
    .filter(Boolean)
    .join(', ');
  const totalTime = String(payload && payload.totalTime || '').trim();

  if (colEndTach >= 0) logSheet.getRange(rowIdx, colEndTach + 1).setValue(endTach);
  if (colFuelEnd >= 0) logSheet.getRange(rowIdx, colFuelEnd + 1).setValue(fuelEnd);
  if (colFuelUsed >= 0) logSheet.getRange(rowIdx, colFuelUsed + 1).setValue(fuelUsed);
  if (colAirborne >= 0) logSheet.getRange(rowIdx, colAirborne + 1).setValue(airborne);
  if (colLanded >= 0) logSheet.getRange(rowIdx, colLanded + 1).setValue(landed);
  if (colNumLdgs >= 0) logSheet.getRange(rowIdx, colNumLdgs + 1).setValue(numLdgs);
  if (colBrakesApplied >= 0 && brakesApplied) logSheet.getRange(rowIdx, colBrakesApplied + 1).setValue(brakesApplied);
  if (colTotalTime >= 0 && totalTime) logSheet.getRange(rowIdx, colTotalTime + 1).setValue(totalTime);
  if (colSquawks >= 0) logSheet.getRange(rowIdx, colSquawks + 1).setValue(squawks);
  if (colNumTg >= 0) {
    logSheet.getRange(rowIdx, colNumTg + 1).setValue(numTouchAndGos);
  } else if (colActualLoadJson >= 0) {
    const existingRaw = data[rowIdx - 1][colActualLoadJson] || '';
    let existing = {};
    try { existing = existingRaw ? JSON.parse(existingRaw) : {}; } catch (e) { existing = {}; }
    const merged = {
      ...existing,
      debrief: {
        ...(existing.debrief || {}),
        numTouchAndGos: numTouchAndGos,
        totalTime: totalTime,
        debriefAt: new Date().toISOString()
      }
    };
    logSheet.getRange(rowIdx, colActualLoadJson + 1).setValue(JSON.stringify(merged));
  }

  if (squawks) {
    _toolsAddDebriefSquawksToAircraft_({
      aircraftReg: String(data[rowIdx - 1][LOG_FLIGHT_COL.ACFT] || '').trim().toUpperCase(),
      squawks: squawks,
      reportDate: data[rowIdx - 1][LOG_FLIGHT_COL.DATE],
      tachAtReport: endTach,
      reportedBy: String(data[rowIdx - 1][LOG_FLIGHT_COL.PILOT] || '').trim(),
      sourceFlightLegId: flightLegId
    });
  }

  return {
    success: true,
    flightLegId: flightLegId,
    endTach: endTach,
    fuelEnd: fuelEnd,
    fuelUsed: fuelUsed,
    airborne: airborne,
    landed: landed,
    numLdgs: numLdgs,
    brakesApplied: brakesApplied,
    numTouchAndGos: numTouchAndGos,
    squawks: squawks,
    totalTime: totalTime,
    debriefAt: new Date().toISOString()
  };
}

function initializeWB(flightId) {
  if (!flightId) throw new Error('initializeWB: flightId is required');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dispatchSheet = ss.getSheetByName(APP_SHEETS.DISPATCH);
  const logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  const pilotsSheet = ss.getSheetByName(APP_SHEETS.PILOTS);
  const aircraftSheet = ss.getSheetByName(APP_SHEETS.AIRCRAFT);
  const airframesSheet = ss.getSheetByName(APP_SHEETS.AIRFRAMES);
  const envelopesSheet = ss.getSheetByName(APP_SHEETS.ENVELOPES);

  if (!dispatchSheet) throw new Error("Sheet 'DB_Dispatch' not found.");
  if (!logSheet) throw new Error("Sheet 'LOG_Flights' not found.");
  if (!pilotsSheet) throw new Error("Sheet 'DB_Pilots' not found.");
  if (!aircraftSheet) throw new Error("Sheet 'DB_Aircraft' not found.");
  if (!airframesSheet) throw new Error("Sheet 'REF_Airframes' not found.");
  if (!envelopesSheet) throw new Error("Sheet 'REF_Envelopes' not found.");

  const normalize = (v) => String(v || '').toUpperCase().trim().replace(/\s+/g, '_');
  const toNum = (v, d = 0) => {
    const n = parseFloat(v);
    return isNaN(n) ? d : n;
  };

  const getTable = (sheet) => {
    const values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) return { headers: [], rows: [] };
    const headers = values[0].map(normalize);
    return { headers: headers, rows: values.slice(1) };
  };

  const rowToObj = (headers, row) => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  };

  const findBy = (rows, headers, key, val) => {
    const idx = headers.indexOf(normalize(key));
    if (idx < 0) return null;
    const target = String(val || '').trim().toUpperCase();
    return rows.find(r => String(r[idx] || '').trim().toUpperCase() === target) || null;
  };
  const findByAny = (rows, headers, keys, val) => {
    const aliases = Array.isArray(keys) ? keys : [keys];
    for (let i = 0; i < aliases.length; i++) {
      const hit = findBy(rows, headers, aliases[i], val);
      if (hit) return hit;
    }
    return null;
  };

  const dispatchTable = getTable(dispatchSheet);
  const logTable = getTable(logSheet);
  const pilotsTable = getTable(pilotsSheet);
  const aircraftTable = getTable(aircraftSheet);
  const airframesTable = getTable(airframesSheet);
  const envelopesTable = getTable(envelopesSheet);

  const dispatchRow = findBy(dispatchTable.rows, dispatchTable.headers, 'FLIGHT_ID', flightId);
  if (!dispatchRow) throw new Error('initializeWB: flight not found in DB_Dispatch: ' + flightId);
  const dispatch = rowToObj(dispatchTable.headers, dispatchRow);

  const logRow = findBy(logTable.rows, logTable.headers, 'FLIGHT_ID', flightId);
  const logObj = logRow ? rowToObj(logTable.headers, logRow) : {};

  const pilotName = String(dispatch.PILOT || '');
  const aircraftReg = String(dispatch.AIRCRAFT || '');
  const flightDate = String(dispatch.DATE || '');
  
  // Extract takeoff time: try TIME/TAKEOFF/TAKEOFF_TIME first, then parse from RAW_DATA JSON
  let flightTime = String(dispatch.TIME || dispatch.TAKEOFF || dispatch.TAKEOFF_TIME || '');
  if (!flightTime || flightTime === '00:00' || !flightTime.includes(':')) {
    // TIME column is probably flight duration, try to get from RAW_DATA JSON
    try {
      const tempRawData = JSON.parse(dispatch.RAW_DATA || '{}');
      flightTime = String(tempRawData.time || tempRawData.meta?.time || '00:00');
    } catch (e) {
      flightTime = '00:00';
    }
  }

  const pilotRow = findByAny(pilotsTable.rows, pilotsTable.headers, ['PILOT_NAME', 'NAME'], pilotName);
  const pilotObj = pilotRow ? rowToObj(pilotsTable.headers, pilotRow) : {};
  const pilotWeight = toNum(pilotObj.WEIGHT_KGS, 0);

  let aircraftRow = findByAny(aircraftTable.rows, aircraftTable.headers, ['REGISTRATION', 'REG', 'TAIL', 'TAIL_NUMBER'], aircraftReg);
  if (!aircraftRow) {
    aircraftRow = findByAny(aircraftTable.rows, aircraftTable.headers, ['AIRCRAFT_TYPE', 'TYPE_FOR_PERFORMANCE'], aircraftReg);
  }
  if (!aircraftRow) throw new Error('initializeWB: aircraft not found in DB_Aircraft: ' + aircraftReg);
  const aircraftObj = rowToObj(aircraftTable.headers, aircraftRow);
  const aircraftType = String(aircraftObj.AIRCRAFT_TYPE || aircraftObj.TYPE_FOR_PERFORMANCE || aircraftReg);

  const stationRows = airframesTable.rows
    .map(r => rowToObj(airframesTable.headers, r))
    .filter(r => String(r.AIRCRAFT_TYPE || '').trim().toUpperCase() === aircraftType.trim().toUpperCase());

  const envelopeData = envelopesTable.rows
    .map(r => rowToObj(envelopesTable.headers, r))
    .filter(r => {
      return String(r.AIRCRAFT_TYPE || '').trim().toUpperCase() === aircraftType.trim().toUpperCase();
    })
    .map(r => ({
      AIRCRAFT_TYPE: r.AIRCRAFT_TYPE,
      POINT_SEQUENCE: r.POINT_SEQUENCE,
      CG_Arm_X: toNum(r.CG_ARM_X, 0),
      Weight_Y: toNum(r.WEIGHT_Y, 0)
    }))
    .filter(r => {
      const x = toNum(r.CG_Arm_X, NaN);
      const y = toNum(r.Weight_Y, NaN);
      return !isNaN(x) && !isNaN(y);
    })
    .sort((a, b) => toNum(a.POINT_SEQUENCE, 0) - toNum(b.POINT_SEQUENCE, 0));

  if (envelopeData.length < 3) {
    const availableTypes = Array.from(new Set(envelopesTable.rows
      .map(r => rowToObj(envelopesTable.headers, r))
      .map(r => String(r.AIRCRAFT_TYPE || '').trim())
      .filter(Boolean)))
      .slice(0, 12)
      .join(', ');
    throw new Error(
      'initializeWB: envelope points not found in REF_Envelopes for aircraft type "' + aircraftType + '". ' +
      'Expected columns: Aircraft_Type, Point_Sequence, CG_Arm_X, Weight_Y. ' +
      'Available envelope types: ' + (availableTypes || '(none)')
    );
  }

  const getStationArm = (keywords, fallback) => {
    const hit = stationRows.find(s => {
      const name = String(s.STATION_NAME || '').toUpperCase();
      return keywords.some(k => name.indexOf(k) >= 0);
    });
    return hit ? toNum(hit.ARM, fallback) : fallback;
  };

  const pilotArm = getStationArm(['PILOT', 'FRONT', 'FWD'], toNum(aircraftObj.EMPTY_ARM, 0));
  const midArm = getStationArm(['MID', 'MIDDLE', 'ROW2', 'PASS'], pilotArm);
  const aftArm = getStationArm(['AFT', 'REAR', 'ROW3'], midArm);
  const cargoArm = getStationArm(['CARGO', 'BAG', 'FREIGHT'], aftArm);
  const fuelArm = getStationArm(['FUEL', 'TANK'], toNum(aircraftObj.EMPTY_ARM, 0));

  // Extract all cargo areas from REF_Airframes
  const cargoAreas = stationRows
    .filter(s => {
      const name = String(s.STATION_NAME || '').toUpperCase();
      return name.indexOf('CARGO') >= 0 || name.indexOf('POD') >= 0;
    })
    .map(s => {
      const name = String(s.STATION_NAME || '').trim();
      const maxWeightLbs = toNum(s.MAX_WEIGHT_LBS, 0);
      const maxWeightKg = maxWeightLbs > 0 ? maxWeightLbs / 2.20462 : 0; // Convert lbs to kg
      return {
        id: name.replace(/\s+/g, '_').toLowerCase(),
        name: name,
        arm: toNum(s.ARM, cargoArm),
        maxWeightKg: maxWeightKg,
        maxWeightLbs: maxWeightLbs
      };
    })
    .sort((a, b) => {
      // Sort by arm position for logical ordering
      return a.arm - b.arm;
    });

  let rawData = {};
  try {
    rawData = JSON.parse(dispatch.RAW_DATA || '{}');
  } catch (e) {
    rawData = {};
  }

  const legPayload = (rawData && rawData.legs && Array.isArray(rawData.legs) && rawData.legs.length > 0)
    ? rawData.legs[0]
    : (rawData || {});
  const pax = Array.isArray(legPayload.pax) ? legPayload.pax : [];
  
  // Build cargo manifest: list each passenger's cargo and freight separately
  const cargoManifest = [];
  pax.forEach(p => {
    if (String(p.name || '').toUpperCase() === 'FREIGHT') {
      // Freight is listed as a passenger with type='FREIGHT'
      const freightWeight = toNum(p.weight, 0);
      if (freightWeight > 0) {
        cargoManifest.push({
          name: 'Freight',
          plannedWeight: freightWeight,
          actualWeight: freightWeight,
          type: 'freight',
          passengerLinked: false
        });
      }
    } else {
      // Regular passenger cargo
      const cargoWeight = toNum(p.cargo, 0);
      if (cargoWeight > 0) {
        cargoManifest.push({
          name: p.name + ' Cargo',
          plannedWeight: cargoWeight,
          actualWeight: cargoWeight,
          type: 'pax_cargo',
          linkedPassenger: p.name,
          passengerLinked: true
        });
      }
    }
  });
  
  const dispatchCargoWeight = pax.reduce((sum, p) => {
    const cargo = toNum(p.cargo, 0);
    const freightWeight = String(p.name || '').toUpperCase() === 'FREIGHT' ? toNum(p.weight, 0) : 0;
    return sum + cargo + freightWeight;
  }, 0);

  // Extract mission ID to get ALL legs for seat planning
  const missionId = flightId.split('-').slice(0, 2).join('-'); // e.g., ADS26-001
  const allMissionRows = dispatchTable.rows.filter(r => {
    const rowMissionId = String(r[DISPATCH_COL.MISSION_ID] || '').trim();
    return rowMissionId === missionId;
  });

  // Parse all legs to collect ALL unique passengers across the mission
  const allPassengersInMission = [];
  const passsengerSeen = {}; // Track unique passengers by name
  
  allMissionRows.forEach(row => {
    try {
      const legRawData = JSON.parse(String(row[DISPATCH_COL.RAW_DATA] || '{}'));
      const legData = (legRawData && legRawData.legs && Array.isArray(legRawData.legs) && legRawData.legs.length > 0)
        ? legRawData.legs[0]
        : (legRawData || {});
      const legPax = Array.isArray(legData.pax) ? legData.pax : [];
      const nonFreightPax = legPax.filter(p => String(p.name || '').toUpperCase() !== 'FREIGHT');
      nonFreightPax.forEach(pax => {
        const key = String(pax.name || '').trim();
        if (!passsengerSeen[key]) {
          passsengerSeen[key] = true;
          allPassengersInMission.push(pax);
        }
      });
    } catch (e) {}
  });
  
  // Sort by weight (heaviest first)
  allPassengersInMission.sort((a, b) => toNum(b.weight, 0) - toNum(a.weight, 0));
  
  // Determine seats needed: 1 pilot + 1 copilot + N passengers
  const maxPaxInMission = allPassengersInMission.length;

  // Get passengers for THIS leg only (for assigning to seats)
  const thisLegPassengers = pax.filter(p => String(p.name || '').toUpperCase() !== 'FREIGHT')
    .sort((a, b) => toNum(b.weight, 0) - toNum(a.weight, 0));

  const fuelFromLog = toNum(logObj.FUEL_START, 0);
  let fuelFromActualJson = 0;
  let cargoFromActualJson = 0;
  try {
    const existingLoad = JSON.parse(logObj.ACTUAL_LOAD_JSON || '{}');
    // Check for fuel in various locations (Tab 2 saves as fuelTotal in liters)
    if (existingLoad && existingLoad.fuelTotal) fuelFromActualJson = toNum(existingLoad.fuelTotal, 0);
    if (existingLoad && existingLoad.fuel) fuelFromActualJson = toNum(existingLoad.fuel, fuelFromActualJson);
    if (existingLoad && existingLoad.wb && existingLoad.wb.fuel) fuelFromActualJson = toNum(existingLoad.wb.fuel, fuelFromActualJson);

    // Check cargo/load values from briefing payload or saved W&B
    if (existingLoad && existingLoad.cargoTotal) cargoFromActualJson = toNum(existingLoad.cargoTotal, cargoFromActualJson);
    if (existingLoad && existingLoad.cargo) cargoFromActualJson = toNum(existingLoad.cargo, cargoFromActualJson);
    if (existingLoad && existingLoad.wb && existingLoad.wb.cargo) cargoFromActualJson = toNum(existingLoad.wb.cargo, cargoFromActualJson);

    if (existingLoad && existingLoad.wb && Array.isArray(existingLoad.wb.items)) {
      const cargoItem = existingLoad.wb.items.find(item => String(item.type || '').toLowerCase() === 'cargo');
      if (cargoItem) {
        cargoFromActualJson = toNum(cargoItem.actualWeight, toNum(cargoItem.plannedWeight, cargoFromActualJson));
      }
    }
  } catch (e) {}

  // Cargo weight: prioritize saved load JSON, fallback to DB_Dispatch RAW_DATA pax cargo/freight
  let cargoWeight = cargoFromActualJson || dispatchCargoWeight;
  
  // Fuel weight: if from LOG use as-is (already in liters), convert to kg (Avgas ~0.72 kg/L)
  const fuelLiters = fuelFromActualJson || fuelFromLog;
  const fuelWeight = fuelLiters * 0.72; // Convert liters to kg

  // Define seat configuration with loading priority
  const seatDefinitions = [
    { id: 'pilot', label: 'Pilot Seat', arm: pilotArm, weight: toNum(aircraftObj.PILOT_SEAT_KGS, 0), priority: 1, required: true, locked: true },
    { id: 'copilot', label: 'Copilot Seat', arm: pilotArm, weight: toNum(aircraftObj.PILOT_SEAT_KGS, 0), priority: 2, required: false, locked: false },
    { id: 'lh-mid', label: 'LH Mid Seat', arm: midArm, weight: toNum(aircraftObj.MID_SEAT_KGS, 0), priority: 3, required: false, locked: false },
    { id: 'rh-mid', label: 'RH Mid Seat', arm: midArm, weight: toNum(aircraftObj.MID_SEAT_KGS, 0), priority: 4, required: false, locked: false },
    { id: 'lh-aft', label: 'LH Aft Seat', arm: aftArm, weight: toNum(aircraftObj.AFT_SEAT_KGS, 0), priority: 5, required: false, locked: false },
    { id: 'rh-aft', label: 'RH Aft Seat', arm: aftArm, weight: toNum(aircraftObj.AFT_SEAT_KGS, 0), priority: 6, required: false, locked: false }
  ];

  // Determine how many seats to install: 1 pilot + 1 copilot + (passengers for this leg)
  const seatsNeededForThisLeg = Math.min(thisLegPassengers.length + 1, 6); // +1 for pilot
  const totalSeatsOnAircraft = Math.min(maxPaxInMission + 1, 6); // +1 for pilot only

  // Assign passengers on THIS leg to seats (by weight, heaviest front)
  const seatAssignments = {};
  seatDefinitions.forEach((seatDef, idx) => {
    let status = 'base';
    let passenger = null;
    let occupiedWeight = 0;
    let isOccupied = false;

    if (idx === 0) {
      // Pilot seat - always installed and occupied
      status = 'installed';
      passenger = { name: pilotName, weight: pilotWeight };
      occupiedWeight = pilotWeight;
      isOccupied = true;
    } else if (idx < seatsNeededForThisLeg) {
      // Seat needed for this leg - install it
      status = 'installed';
      const paxIdx = idx - 1; // Adjust for pilot
      if (paxIdx < thisLegPassengers.length) {
        passenger = thisLegPassengers[paxIdx];
        occupiedWeight = toNum(passenger.weight, 0);
        isOccupied = true;
      }
    } else if (idx < totalSeatsOnAircraft) {
      // Seat on aircraft but not needed this leg - put in cargo
      status = 'cargo';
    } else {
      // Seat not on aircraft - left at base
      status = 'base';
    }

    seatAssignments[seatDef.id] = {
      label: seatDef.label,
      arm: seatDef.arm,
      seatWeight: seatDef.weight,
      status: status,
      passenger: passenger,
      occupiedWeight: occupiedWeight,
      isOccupied: isOccupied,
      enabled: (status === 'installed'),
      locked: seatDef.locked
    };
  });

  // Build manifest items - ONLY show occupied seats, not separate passenger items
  const items = [
    {
      name: 'Empty Aircraft',
      plannedWeight: toNum(aircraftObj.EMPTY_WEIGHT, 0),
      actualWeight: toNum(aircraftObj.EMPTY_WEIGHT, 0),
      arm: toNum(aircraftObj.EMPTY_ARM, 0),
      type: 'empty'
    }
  ];

  // Add only OCCUPIED seats to manifest (don't duplicate passengers)
  Object.keys(seatAssignments).forEach(seatId => {
    const seat = seatAssignments[seatId];
    if (seat.status === 'installed' && seat.occupiedWeight > 0) {
      items.push({
        name: seat.label + (seat.passenger && seat.passenger.name !== pilotName ? ': ' + seat.passenger.name : ''),
        plannedWeight: seat.occupiedWeight,
        actualWeight: seat.occupiedWeight,
        arm: seat.arm,
        type: 'passenger',
        seatId: seatId
      });
    }
  });

  // Add cargo (includes freight + uninstalled seats)
  items.push({
    name: 'Cargo',
    plannedWeight: cargoWeight,
    actualWeight: cargoWeight,
    arm: cargoArm,
    type: 'cargo'
  });

  // Add fuel
  items.push({
    name: 'Fuel',
    plannedWeight: fuelWeight,
    actualWeight: fuelWeight,
    arm: fuelArm,
    type: 'fuel'
  });

  // Build detailed seat status for UI
  const seats = {};
  Object.keys(seatAssignments).forEach(seatId => {
    const seat = seatAssignments[seatId];
    seats[seat.label] = {
      seatId: seatId,
      weight: seat.seatWeight,
      arm: seat.arm,
      status: seat.status, // 'installed', 'cargo', 'base'
      enabled: seat.enabled,
      passenger: seat.passenger,
      occupiedWeight: seat.occupiedWeight,
      locked: seat.locked
    };
  });

  return {
    flightId: String(flightId),
    aircraft: aircraftReg,
    pilot: pilotName,
    date: flightDate,
    time: flightTime,
    airframeData: {
      Empty_Weight: toNum(aircraftObj.EMPTY_WEIGHT, 0),
      Empty_Arm: toNum(aircraftObj.EMPTY_ARM, 0),
      MTOW: toNum(aircraftObj.MTOW, 0),
      Fuel_Burn_Per_Hour: toNum(aircraftObj.FUEL_BURN_PER_HOUR, 12)
    },
    envelopeData: envelopeData,
    items: items,
    seats: seats,
    seatAssignments: seatAssignments, // Detailed seat info with passengers
    maxPaxInMission: maxPaxInMission,
    thisLegPaxCount: thisLegPassengers.length,
    fuel: fuelWeight,
    fuelArm: fuelArm,
    cargoAreas: cargoAreas,
    cargoManifest: cargoManifest
  };
}

function getWbEnvelopeByAircraft(aircraftReg) {
  const reg = String(aircraftReg || '').trim().toUpperCase();
  if (!reg) throw new Error('getWbEnvelopeByAircraft: aircraftReg is required');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aircraftSheet = ss.getSheetByName(APP_SHEETS.AIRCRAFT);
  const envelopesSheet = ss.getSheetByName(APP_SHEETS.ENVELOPES);

  if (!aircraftSheet) throw new Error("Sheet 'DB_Aircraft' not found.");
  if (!envelopesSheet) throw new Error("Sheet 'REF_Envelopes' not found.");

  const normalize = (v) => String(v || '').toUpperCase().trim().replace(/\s+/g, '_');
  const toNum = (v, d = 0) => {
    const n = parseFloat(v);
    return isNaN(n) ? d : n;
  };
  const getTable = (sheet) => {
    const values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) return { headers: [], rows: [] };
    const headers = values[0].map(normalize);
    return { headers: headers, rows: values.slice(1) };
  };
  const rowToObj = (headers, row) => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  };
  const findByAny = (rows, headers, keys, val) => {
    const target = String(val || '').trim().toUpperCase();
    const aliases = Array.isArray(keys) ? keys : [keys];
    for (let i = 0; i < aliases.length; i++) {
      const idx = headers.indexOf(normalize(aliases[i]));
      if (idx < 0) continue;
      const hit = rows.find(function(r) {
        return String(r[idx] || '').trim().toUpperCase() === target;
      });
      if (hit) return hit;
    }
    return null;
  };

  const aircraftTable = getTable(aircraftSheet);
  const envelopesTable = getTable(envelopesSheet);

  let aircraftRow = findByAny(
    aircraftTable.rows,
    aircraftTable.headers,
    ['REGISTRATION', 'REG', 'TAIL', 'TAIL_NUMBER'],
    reg
  );
  if (!aircraftRow) {
    aircraftRow = findByAny(
      aircraftTable.rows,
      aircraftTable.headers,
      ['AIRCRAFT_TYPE', 'TYPE_FOR_PERFORMANCE'],
      reg
    );
  }
  if (!aircraftRow) {
    throw new Error('getWbEnvelopeByAircraft: aircraft not found in DB_Aircraft: ' + reg);
  }

  const aircraftObj = rowToObj(aircraftTable.headers, aircraftRow);
  const aircraftType = String(aircraftObj.AIRCRAFT_TYPE || aircraftObj.TYPE_FOR_PERFORMANCE || reg).trim();
  const aircraftTypeUpper = aircraftType.toUpperCase();

  const envelopeData = envelopesTable.rows
    .map(function(r) { return rowToObj(envelopesTable.headers, r); })
    .filter(function(r) {
      return String(r.AIRCRAFT_TYPE || '').trim().toUpperCase() === aircraftTypeUpper;
    })
    .map(function(r) {
      return {
        AIRCRAFT_TYPE: r.AIRCRAFT_TYPE,
        POINT_SEQUENCE: r.POINT_SEQUENCE,
        CG_Arm_X: toNum(r.CG_ARM_X, 0),
        Weight_Y: toNum(r.WEIGHT_Y, 0)
      };
    })
    .filter(function(r) {
      const x = toNum(r.CG_Arm_X, NaN);
      const y = toNum(r.Weight_Y, NaN);
      return !isNaN(x) && !isNaN(y);
    })
    .sort(function(a, b) {
      return toNum(a.POINT_SEQUENCE, 0) - toNum(b.POINT_SEQUENCE, 0);
    });

  if (envelopeData.length < 3) {
    const availableTypes = Array.from(new Set(
      envelopesTable.rows
        .map(function(r) { return rowToObj(envelopesTable.headers, r); })
        .map(function(r) { return String(r.AIRCRAFT_TYPE || '').trim(); })
        .filter(Boolean)
    )).slice(0, 12).join(', ');
    throw new Error(
      'getWbEnvelopeByAircraft: envelope points not found in REF_Envelopes for aircraft type "' + aircraftType + '". ' +
      'Available envelope types: ' + (availableTypes || '(none)')
    );
  }

  return {
    aircraft: reg,
    aircraftType: aircraftType,
    cachedAt: new Date().toISOString(),
    envelopeData: envelopeData
  };
}

function saveWBToLog(flightId, wbPayload) {
  if (!flightId) throw new Error('saveWBToLog: flightId is required');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  if (!logSheet) throw new Error("Sheet 'LOG_Flights' not found.");

  const data = logSheet.getDataRange().getValues();
  if (!data || data.length < 2) throw new Error('saveWBToLog: LOG_Flights has no data rows');

  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][LOG_FLIGHT_COL.FLIGHT_ID] || '').trim() === String(flightId).trim()) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow < 0) {
    throw new Error('saveWBToLog: flight not found in LOG_Flights: ' + flightId);
  }

  const existingRaw = data[targetRow - 1][LOG_FLIGHT_COL.ACTUAL_LOAD_JSON] || '';
  let existingJson = {};
  try {
    existingJson = existingRaw ? JSON.parse(existingRaw) : {};
  } catch (e) {
    existingJson = { _parseError: 'Invalid existing Actual_Load_JSON', _raw: String(existingRaw) };
  }

  const merged = {
    ...existingJson,
    wb: wbPayload || {},
    wbSavedAt: new Date().toISOString()
  };

  logSheet.getRange(targetRow, LOG_FLIGHT_COL.ACTUAL_LOAD_JSON + 1).setValue(JSON.stringify(merged));
  return true;
}

function saveTakeoffRollToLog(flightId, rollPayload) {
  if (!flightId) throw new Error('saveTakeoffRollToLog: flightId is required');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  if (!logSheet) throw new Error("Sheet 'LOG_Flights' not found.");

  const data = logSheet.getDataRange().getValues();
  if (!data || data.length < 2) throw new Error('saveTakeoffRollToLog: LOG_Flights has no data rows');

  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][LOG_FLIGHT_COL.FLIGHT_ID] || '').trim() === String(flightId).trim()) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow < 0) {
    throw new Error('saveTakeoffRollToLog: flight not found in LOG_Flights: ' + flightId);
  }

  const existingRaw = data[targetRow - 1][LOG_FLIGHT_COL.ACTUAL_LOAD_JSON] || '';
  let existingJson = {};
  try {
    existingJson = existingRaw ? JSON.parse(existingRaw) : {};
  } catch (e) {
    existingJson = { _parseError: 'Invalid existing Actual_Load_JSON', _raw: String(existingRaw) };
  }

  const merged = {
    ...existingJson,
    takeoffRoll: rollPayload || {},
    takeoffRollSavedAt: new Date().toISOString()
  };

  logSheet.getRange(targetRow, LOG_FLIGHT_COL.ACTUAL_LOAD_JSON + 1).setValue(JSON.stringify(merged));
  return true;
}

function _perfNorm(v) {
  return String(v || '').toUpperCase().trim().replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

function _perfNum(v, d) {
  const n = parseFloat(v);
  return isNaN(n) ? (d || 0) : n;
}

function _perfValue(obj, keys, fallback) {
  for (let i = 0; i < keys.length; i++) {
    const k = _perfNorm(keys[i]);
    if (Object.prototype.hasOwnProperty.call(obj, k) && obj[k] !== '' && obj[k] != null) {
      return obj[k];
    }
  }
  return fallback;
}

function _perfTable(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return { headers: [], rows: [] };
  const vals = sh.getDataRange().getValues();
  if (!vals || vals.length < 2) return { headers: [], rows: [] };

  const headers = vals[0].map(h => _perfNorm(h));
  const rows = vals.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, idx) => obj[h] = r[idx]);
    return obj;
  });

  return { headers: headers, rows: rows };
}

function _interpBaseRoll(baseRows, weightKg) {
  if (!baseRows || !baseRows.length) {
    return { to: 0, ldg: 0 };
  }

  const sorted = baseRows.slice().sort((a, b) => a.w - b.w);
  if (weightKg <= sorted[0].w) return { to: sorted[0].to, ldg: sorted[0].ldg };
  if (weightKg >= sorted[sorted.length - 1].w) {
    const last = sorted[sorted.length - 1];
    return { to: last.to, ldg: last.ldg };
  }

  for (let i = 0; i < sorted.length - 1; i++) {
    const a = sorted[i];
    const b = sorted[i + 1];
    if (weightKg >= a.w && weightKg <= b.w) {
      const ratio = (weightKg - a.w) / ((b.w - a.w) || 1);
      return {
        to: a.to + ratio * (b.to - a.to),
        ldg: a.ldg + ratio * (b.ldg - a.ldg)
      };
    }
  }

  return { to: sorted[0].to, ldg: sorted[0].ldg };
}

function _nearestMultiplier(rows, keyNames, target, valueNames, fallback) {
  const points = [];
  rows.forEach(r => {
    const keyRaw = _perfValue(r, keyNames, '');
    const valRaw = _perfValue(r, valueNames, '');
    const k = parseFloat(keyRaw);
    const v = parseFloat(valRaw);
    if (isNaN(k) || isNaN(v)) return;
    points.push({ k: k, v: v });
  });

  if (!points.length) return fallback || 1;
  points.sort((a, b) => a.k - b.k);

  for (let i = 0; i < points.length; i++) {
    if (points[i].k === target) return points[i].v;
  }

  for (let i = 0; i < points.length - 1; i++) {
    const a = points[i];
    const b = points[i + 1];
    if (target > a.k && target < b.k) {
      const ratio = (target - a.k) / ((b.k - a.k) || 1);
      return a.v + ratio * (b.v - a.v);
    }
  }

  if (target <= points[0].k) return points[0].v;
  if (target >= points[points.length - 1].k) return points[points.length - 1].v;
  return fallback || 1;
}

function _slopeLookupTarget(rows, slopeAbsPercent) {
  const slopeAbs = Math.abs(_perfNum(slopeAbsPercent, 0));
  let maxSlopeKey = 0;
  (rows || []).forEach(function(r) {
    const k = parseFloat(_perfValue(r, ['SLOPE'], ''));
    if (!isNaN(k)) maxSlopeKey = Math.max(maxSlopeKey, Math.abs(k));
  });
  return maxSlopeKey > 1.5 ? slopeAbs : (slopeAbs / 100);
}

function _surfaceMultiplier(rows, surfaceText, valueNames, fallback) {
  const key = _perfNorm(surfaceText || '');
  if (!key) return fallback || 1;

  for (let i = 0; i < rows.length; i++) {
    const rowSurface = _perfNorm(_perfValue(rows[i], ['SURFACE'], ''));
    if (!rowSurface) continue;
    if (rowSurface === key || rowSurface.indexOf(key) >= 0 || key.indexOf(rowSurface) >= 0) {
      const v = parseFloat(_perfValue(rows[i], valueNames, ''));
      if (!isNaN(v)) return v;
    }
  }

  return fallback || 1;
}

function _densityAltitudeFt(elevationFt, qnhHpa, tempC) {
  const elev = _perfNum(elevationFt, 0);
  const qnh = _perfNum(qnhHpa, 1013.25);
  const temp = _perfNum(tempC, 15);

  const pressureAltitude = elev + ((1013.25 - qnh) * 30);
  const isaTemp = 15 - (2 * (elev / 1000));
  return pressureAltitude + (120 * (temp - isaTemp));
}

function _normalizeSlopeProfile(profileRaw, runwayLengthM, fallbackSlope) {
  const runwayLen = Math.max(_perfNum(runwayLengthM, 0), 0);
  const fb = _perfNum(fallbackSlope, 0);

  let profile = Array.isArray(profileRaw)
    ? profileRaw.map(seg => ({
        distance: _perfNum(seg && (seg.distance ?? seg.length ?? seg.segmentM), 0),
        slope: _perfNum(seg && (seg.slope ?? seg.slopePercent ?? seg.grade), fb)
      })).filter(seg => seg.distance > 0 && isFinite(seg.distance) && isFinite(seg.slope))
    : [];

  if (!profile.length) {
    return [{ distance: runwayLen > 0 ? runwayLen : 1, slope: fb }];
  }

  const sumDist = profile.reduce((acc, seg) => acc + seg.distance, 0);
  if (sumDist < runwayLen) {
    profile.push({ distance: runwayLen - sumDist, slope: 0 });
  } else if (sumDist > runwayLen && runwayLen > 0) {
    let remaining = runwayLen;
    const trimmed = [];
    for (let i = 0; i < profile.length && remaining > 0; i++) {
      const d = Math.min(profile[i].distance, remaining);
      if (d > 0) trimmed.push({ distance: d, slope: profile[i].slope });
      remaining -= d;
    }
    profile = trimmed.length ? trimmed : [{ distance: runwayLen, slope: fb }];
  }

  return profile;
}

function _effectiveSlopeOverDistance(profile, distanceM) {
  const target = Math.max(_perfNum(distanceM, 0), 0);
  if (!target) return 0;
  if (!Array.isArray(profile) || !profile.length) return 0;

  let remaining = target;
  let weighted = 0;
  let lastSlope = _perfNum(profile[profile.length - 1].slope, 0);

  for (let i = 0; i < profile.length && remaining > 0; i++) {
    const seg = profile[i];
    const segDist = Math.max(_perfNum(seg.distance, 0), 0);
    const segSlope = _perfNum(seg.slope, 0);
    const used = Math.min(segDist, remaining);
    weighted += used * segSlope;
    remaining -= used;
    lastSlope = segSlope;
  }

  if (remaining > 0) {
    weighted += remaining * lastSlope;
  }

  return weighted / target;
}

function _runwayHeadingFromIdent(rwyIdent) {
  const raw = String(rwyIdent || '').toUpperCase();
  const m = raw.match(/(\d{1,2})/);
  if (!m) return 0;
  const n = parseInt(m[1], 10);
  if (isNaN(n) || n < 1 || n > 36) return 0;
  return (n * 10) % 360;
}

function _runwayDisplayIdent(rwyIdent) {
  const raw = String(rwyIdent == null ? '' : rwyIdent).trim().toUpperCase();
  const m = raw.match(/(\d{1,2})/);
  if (!m) return raw || 'RWY';
  const n = parseInt(m[1], 10);
  if (isNaN(n) || n < 1 || n > 36) return raw || 'RWY';
  return String(n).padStart(2, '0');
}

function _runwayPairKey_(rwyIdent) {
  const raw = String(rwyIdent || '').trim().toUpperCase();
  if (!raw) return '';
  const m = raw.match(/(\d{1,2})([LCR])?/);
  if (!m) return raw;
  const num = parseInt(m[1], 10);
  if (isNaN(num) || num < 1 || num > 36) return raw;
  const suffix = m[2] || '';
  const recipNum = ((num + 18 - 1) % 36) + 1;
  const recipSuffix = suffix === 'L' ? 'R' : (suffix === 'R' ? 'L' : suffix);
  return [
    String(num).padStart(2, '0') + suffix,
    String(recipNum).padStart(2, '0') + recipSuffix
  ].sort().join('/');
}

function _runwayReverseSide_(side) {
  const s = String(side || '').trim().toLowerCase();
  if (s === 'left') return 'right';
  if (s === 'right') return 'left';
  return s || 'right';
}

function _runwayMirrorDistance_(distanceM, lengthM) {
  const dist = Number(distanceM || 0);
  const len = Math.max(Number(lengthM || 0), 0);
  return Math.max(0, Math.round(len - dist));
}

function _runwayAverageSlopeFromSegments_(segments, fallback) {
  const list = Array.isArray(segments) ? segments : [];
  if (!list.length) return Number(fallback || 0) || 0;
  const total = list.reduce(function(sum, seg) {
    return sum + Math.max(Number(seg && (seg.distance || seg.distanceM) || 0), 0);
  }, 0);
  if (!total) return Number(fallback || 0) || 0;
  const weighted = list.reduce(function(sum, seg) {
    const d = Math.max(Number(seg && (seg.distanceM || seg.distance) || 0), 0);
    const s = Number(seg && (seg.slope || seg.slopePercent) || 0) || 0;
    return sum + (d * s);
  }, 0);
  return weighted / total;
}

function _transformSurveyForRunway_(survey, sourceIdent, targetIdent, runwayLengthM) {
  const base = Object.assign({}, survey || {});
  const source = String(sourceIdent || '').trim().toUpperCase();
  const target = String(targetIdent || '').trim().toUpperCase();
  const lengthM = Math.max(Number(base && base.lengthM || runwayLengthM || 0), Number(runwayLengthM || 0), 0);
  if (!source || !target || source === target) {
    return Object.assign({}, base, {
      lengthM: lengthM || Number(base && base.lengthM || 0) || 0
    });
  }

  const mirroredFeatures = (Array.isArray(base.features) ? base.features : []).map(function(item) {
    return Object.assign({}, item, {
      distance: _runwayMirrorDistance_(item && item.distance, lengthM),
      fromThreshold: target,
      side: _runwayReverseSide_(item && item.side)
    });
  }).sort(function(a, b) { return Number(a.distance || 0) - Number(b.distance || 0); });

  const mirroredMarkers = (Array.isArray(base.markers) ? base.markers : []).map(function(item) {
    return Object.assign({}, item, {
      distanceM: _runwayMirrorDistance_(item && item.distanceM, lengthM),
      fromThreshold: target,
      side: _runwayReverseSide_(item && item.side)
    });
  }).sort(function(a, b) { return Number(a.distanceM || 0) - Number(b.distanceM || 0); });

  const mirroredSlopeSegments = (Array.isArray(base.slopeSegments) ? base.slopeSegments : []).slice().reverse().map(function(seg) {
    const segDistance = Math.max(Number(seg && (seg.distanceM != null ? seg.distanceM : seg.distance) || 0), 0);
    const segStartRaw = Number(seg && seg.startDistanceM);
    const segStart = isFinite(segStartRaw) ? segStartRaw : 0;
    const mirroredStart = Math.max(0, Math.round(lengthM - (segStart + segDistance)));
    return Object.assign({}, seg, {
      fromThreshold: target,
      startDistanceM: mirroredStart,
      distanceM: segDistance,
      slope: -1 * (Number(seg && seg.slope || 0) || 0)
    });
  });

  const cutdown = base && base.cutdownAreas ? base.cutdownAreas : {};
  const thresholds = base && base.thresholds ? base.thresholds : {};

  return Object.assign({}, base, {
    lengthM: lengthM || Number(base && base.lengthM || 0) || 0,
    features: mirroredFeatures,
    markers: mirroredMarkers,
    slopeFromThreshold: target,
    slopeSegments: mirroredSlopeSegments,
    cutdownAreas: {
      thrA: cutdown.thrB != null ? cutdown.thrB : (cutdown.thrA != null ? cutdown.thrA : null),
      thrB: cutdown.thrA != null ? cutdown.thrA : (cutdown.thrB != null ? cutdown.thrB : null)
    },
    thresholds: {
      a: thresholds.b || {},
      b: thresholds.a || {}
    },
    obstacleAngles50m: Array.isArray(base.obstacleAngles50m) ? base.obstacleAngles50m.map(function(item) {
      const fromThrRaw = String(item && item.fromThreshold || '').trim().toUpperCase();
      let mirroredThreshold = fromThrRaw;
      if (fromThrRaw === source) mirroredThreshold = target;
      else if (fromThrRaw === target) mirroredThreshold = source;
      return Object.assign({}, item, {
        fromThreshold: mirroredThreshold || target,
        checkpointCorner: String(item && item.checkpointCorner || '').trim().toUpperCase() === 'C' ? 'A' : 'C'
      });
    }) : []
  });
}

function getPerformanceSetup(icao) {
  const airportCode = String(icao || '').trim().toUpperCase();

  const aptTable = _perfTable('DB_Airports');
  const baseTable = _perfTable('Aircraft_Roll_Numbers');
  const perfTable = _perfTable('Performance_Multipliers');

  const toMeters = (v) => {
    if (v == null || v === '') return 0;
    const txt = String(v).trim().toUpperCase();
    const n = parseFloat(txt.replace(',', '.'));
    if (isNaN(n)) return 0;
    if (txt.indexOf('FT') >= 0 || txt.indexOf('FEET') >= 0) {
      return n * 0.3048;
    }
    return n;
  };

  const runways = aptTable.rows
    .filter(r => String(_perfValue(r, ['ICAO'], '')).trim().toUpperCase() === airportCode)
    .map(r => {
      const rawIdent = _perfValue(r, ['RWY_IDENT', 'RWY', 'RUNWAY', 'RUNWAY_DESIGNATOR'], '');
      const rwyIdent = _runwayDisplayIdent(rawIdent);
      const explicitHeading = _perfNum(_perfValue(r, ['RUNWAY_HEADING', 'HEADING'], 0), 0);
      const baseLengthM = toMeters(_perfValue(r, ['LENGTH_OFFICIAL', 'LENGTH_METERS', 'LENGTH_M'], 0));
      const baseWidthM = toMeters(_perfValue(r, ['WIDTH_OFFICIAL', 'WIDTH_METERS', 'WIDTH_M'], 0));
      const slopeRaw = _perfValue(r, ['SLOPE_PERCENT', 'SLOPE_PCT'], 0);
      let defaultSlope = _perfNum(slopeRaw, 0);

      // Parse KNOWN_FEATURES JSON; handle both old format (array) and new format (object with features/metadata)
      let knownFeatures = [];
      let knownObj = {};
      let verifiedOperational = {};
      let officialReference = {};
      const featuresStr = String(_perfValue(r, ['KNOWN_FEATURES', 'FEATURES'], '')).trim();
      if (featuresStr) {
        try {
          const parsed = JSON.parse(featuresStr);
          knownObj = Array.isArray(parsed) ? { features: parsed } : (parsed || {});
          verifiedOperational = (knownObj && knownObj.verifiedOperational && typeof knownObj.verifiedOperational === 'object') ? knownObj.verifiedOperational : {};
          officialReference = (knownObj && knownObj.officialReference && typeof knownObj.officialReference === 'object') ? knownObj.officialReference : {};
          // Support both: array (old) or object with 'features' array (new)
          let featuresList = Array.isArray(parsed)
            ? parsed
            : (Array.isArray(verifiedOperational.features) ? verifiedOperational.features : (parsed.features ? Array.isArray(parsed.features) ? parsed.features : [] : []));
        
          if (featuresList.length) {
            knownFeatures = featuresList.map(f => {
              if (typeof f === 'string') {
                const m = f.match(/^(.*?)(?:\s+(\d+(?:\.\d+)?)\s*m)?(?:\s+(left|right))?$/i);
                return {
                  name: (m && m[1] ? String(m[1]).trim() : f),
                  distance: m && m[2] ? _perfNum(m[2], 0) : 0,
                  side: m && m[3] ? String(m[3]).trim().toLowerCase() : 'right',
                  icon: 'marker'
                };
              }
              return {
                name: String((f && (f.name || f.label || f.feature)) || 'Feature').trim(),
                distance: _perfNum(f && (f.distance ?? f.distanceM ?? f.meters), 0),
                side: String((f && (f.side || f.position)) || 'right').trim().toLowerCase(),
                icon: String((f && (f.icon || f.type)) || 'marker').trim().toLowerCase()
              };
            }).filter(f => f.name && f.distance >= 0 && isFinite(f.distance));
          }
        } catch (e) {
          knownFeatures = [];
        }
      }

      const lengthM = _perfNum(verifiedOperational.lengthM, baseLengthM) || baseLengthM;
      const widthM = _perfNum(verifiedOperational.widthM, baseWidthM) || baseWidthM;

      // Parse segmented slope profile JSON from DB_Airports columns SLOPE_PROFILE / SLOPE_PERCENT
      // Example: [{"distance":100,"slope":4},{"distance":50,"slope":0},{"distance":50,"slope":-2}]
      let slopeProfile = [];
      const slopeProfileRaw = (knownObj && Array.isArray(knownObj.slopeProfile) && knownObj.slopeProfile.length)
        ? knownObj.slopeProfile
        : _perfValue(r, ['SLOPE_PROFILE', 'RUNWAY_SLOPE_PROFILE', 'SLOPE_PERCENT'], '');
      const slopeProfileStr = String(slopeProfileRaw == null ? '' : slopeProfileRaw).trim();
      if (slopeProfileStr) {
        try {
          const parsedProfile = typeof slopeProfileRaw === 'string' ? JSON.parse(slopeProfileStr) : slopeProfileRaw;
          if (Array.isArray(parsedProfile)) {
            slopeProfile = parsedProfile
              .map(seg => ({
                distance: _perfNum(seg && (seg.distance ?? seg.length ?? seg.segmentM), 0),
                slope: _perfNum(seg && (seg.slope ?? seg.slopePercent ?? seg.grade), 0)
              }))
              .filter(seg => seg.distance > 0 && isFinite(seg.distance) && isFinite(seg.slope));

            if (slopeProfile.length) {
              const sumD = slopeProfile.reduce((acc, seg) => acc + seg.distance, 0) || 1;
              const sumWeighted = slopeProfile.reduce((acc, seg) => acc + (seg.distance * seg.slope), 0);
              defaultSlope = sumWeighted / sumD;
            }
          }
        } catch (e) {
          slopeProfile = [];
        }
      }

      if (!slopeProfile.length) {
        slopeProfile = [{ distance: Math.max(lengthM, 0), slope: defaultSlope }];
      }
      if (Array.isArray(verifiedOperational.slopeSegments) && verifiedOperational.slopeSegments.length) {
        defaultSlope = _runwayAverageSlopeFromSegments_(verifiedOperational.slopeSegments, defaultSlope);
      }
      
      return {
        icao: String(_perfValue(r, ['ICAO'], airportCode)).trim().toUpperCase(),
        rwyIdent: rwyIdent,
        headingDeg: explicitHeading > 0 ? explicitHeading : _runwayHeadingFromIdent(rwyIdent),
        length: lengthM,
        width: widthM,
        slope: defaultSlope,
        elevation: _perfNum(_perfValue(r, ['ELEVATION', 'ALT_FEET', 'ELEVATION_FT'], 0), 0),
        surface: String(verifiedOperational.surface || _perfValue(r, ['SURFACE_ACTUAL', 'SURFACE_OFFICIAL', 'SURFACE'], '')).trim(),
        surfaceCondition: String(_perfValue(r, ['SURFACE_CONDITION', 'CONDITION'], '')).trim(),
        pilotNotes: String(_perfValue(r, ['PILOT_NOTES', 'NOTES'], '')).trim(),
        chartUrl: String(_perfValue(r, ['CHART_URL', 'PLATE_URL', 'APPROACH_CHART', 'PROCEDURE_PDF', 'PDF_URL'], '')).trim(),
        airstripPhoto: String(_perfValue(r, ['AIRSTRIP_PHOTO', 'RUNWAY_PHOTO', 'PHOTO_URL'], '')).trim(),
        knownFeatures: knownFeatures,
        slopeProfile: slopeProfile,
        obstacleAngles: Array.isArray(verifiedOperational.obstacleAngles50m) ? verifiedOperational.obstacleAngles50m : (Array.isArray(knownObj.obstacleAngles50m) ? knownObj.obstacleAngles50m : []),
        surveySlopeSegments: Array.isArray(verifiedOperational.slopeSegments) ? verifiedOperational.slopeSegments : [],
        officialReference: {
          lengthM: _perfNum(officialReference.lengthM, baseLengthM) || baseLengthM,
          widthM: _perfNum(officialReference.widthM, baseWidthM) || baseWidthM,
          surface: String(officialReference.surface || _perfValue(r, ['SURFACE_ACTUAL', 'SURFACE_OFFICIAL', 'SURFACE'], '')).trim(),
          headingDeg: _perfNum(officialReference.headingDeg, explicitHeading > 0 ? explicitHeading : _runwayHeadingFromIdent(rwyIdent))
        },
        verifiedOperational: verifiedOperational
      };
    })
    .sort((a, b) => {
      const ha = _perfNum(a.headingDeg, 0);
      const hb = _perfNum(b.headingDeg, 0);
      if (ha !== hb) return ha - hb;
      return String(a.rwyIdent || '').localeCompare(String(b.rwyIdent || ''));
    });

  const flapSet = {};
  perfTable.rows.forEach(r => {
    const flap = parseFloat(_perfValue(r, ['FLAP_SETTING'], ''));
    if (!isNaN(flap)) flapSet[flap] = true;
  });

  const flapOptions = Object.keys(flapSet)
    .map(v => parseFloat(v))
    .filter(v => !isNaN(v))
    .sort((a, b) => a - b);

  const humiditySet = {};
  perfTable.rows.forEach(r => {
    const hum = parseFloat(_perfValue(r, ['HUMIDITY'], ''));
    if (!isNaN(hum)) humiditySet[hum] = true;
  });

  const humidityOptions = Object.keys(humiditySet)
    .map(v => parseFloat(v))
    .filter(v => !isNaN(v))
    .sort((a, b) => a - b);

  const surfaceSet = {};
  perfTable.rows.forEach(r => {
    const s = String(_perfValue(r, ['SURFACE'], '')).trim();
    if (s) surfaceSet[s] = true;
  });

  const surfaceOptions = Object.keys(surfaceSet).sort((a, b) => a.localeCompare(b));



  // Portuguese to English surface translator for ANAC runway surfaces
  const surfaceTranslator = {
    'ASFALTO': 'Paved',
    'CONCRETO': 'Paved',
    'GRAMA': 'Grass to 6"',
    'GRAMA CURTA': 'Short Grass',
    'GRAMA LONGA': 'Long Grass',
    'TERRA': 'Rough',
    'AREIA': 'Sand',
    'CASCALHO': 'Rough',
    'MUD': 'Mud',
    'LAMA': 'Mud',
    'TURF FIRME': 'Firm Turf',
    'TURF': 'Firm Turf'
  };

  return {
    runways: runways,
    flapOptions: flapOptions,
    humidityOptions: humidityOptions,
    surfaceOptions: surfaceOptions,
    surfaceTranslator: surfaceTranslator,
    calcReference: {
      baseRows: baseTable.rows || [],
      perfRows: perfTable.rows || []
    }
  };
}

function getRunwaySurveySurfaceOptions() {
  try {
    const perfTable = _perfTable('Performance_Multipliers');
    const seen = {};
    const options = [];
    (perfTable.rows || []).forEach(function(r) {
      const surface = String(_perfValue(r, ['SURFACE'], '')).trim();
      if (!surface) return;
      const key = surface.toUpperCase();
      if (seen[key]) return;
      seen[key] = true;
      options.push(surface);
    });
    options.sort(function(a, b) { return a.localeCompare(b); });
    return {
      success: true,
      options: options.length ? options : ['Firm Turf', 'Short Grass', 'Grass to 6"', 'Long Grass', 'Rough', 'Mud', 'Sand', 'Asphalt']
    };
  } catch (e) {
    return {
      success: true,
      options: ['Firm Turf', 'Short Grass', 'Grass to 6"', 'Long Grass', 'Rough', 'Mud', 'Sand', 'Asphalt'],
      warning: e && e.message ? e.message : String(e)
    };
  }
}

function calculatePerformanceRolls(payload) {
  try {
    const aircraftType = String(payload && payload.aircraftType || '').trim();
    const weightKg = _perfNum(payload && payload.weightKg, 0);
    const runwayLengthM = _perfNum(payload && payload.runwayLengthM, 0);
    const slopePercent = _perfNum(payload && payload.slopePercent, 0);
    const slopeProfileRaw = payload && payload.slopeProfile;
    const elevationFt = _perfNum(payload && payload.elevationFt, 0);
    const surface = String(payload && payload.surface || '').trim();
    const surfaceCondition = String(payload && payload.surfaceCondition || 'DRY').toUpperCase();
    const qnhHpa = _perfNum(payload && payload.qnhHpa, 1013.25);
    const tempC = _perfNum(payload && payload.tempC, 15);
    const humidityPct = _perfNum(payload && payload.humidityPct, 50);
    const windSpeedKts = _perfNum(payload && payload.windSpeedKts, 0);
    const windType = String(payload && payload.windType || 'HEAD').toUpperCase();
    const windDirectionDeg = _perfNum(payload && payload.windDirectionDeg, 0);
    const runwayHeadingDeg = _perfNum(payload && payload.runwayHeadingDeg, 0);
    const flapSetting = _perfNum(payload && payload.flapSetting, 0);

    if (!aircraftType) return { success: false, error: 'Aircraft type is required.' };
    if (weightKg <= 0) return { success: false, error: 'Weight must be greater than zero.' };
    if (runwayLengthM <= 0) return { success: false, error: 'Runway length must be greater than zero.' };

    const baseTable = _perfTable('Aircraft_Roll_Numbers');
    const perfTable = _perfTable('Performance_Multipliers');

    const baseRows = baseTable.rows
      .filter(r => String(_perfValue(r, ['AIRCRAFT_TYPE'], '')).trim().toUpperCase() === aircraftType.toUpperCase())
      .map(r => ({
        w: _perfNum(_perfValue(r, ['WEIGHT_KG', 'WEIGHT'], 0), 0),
        to: _perfNum(_perfValue(r, ['TO_ROLL_M', 'TO_ROLL'], 0), 0),
        ldg: _perfNum(_perfValue(r, ['LDG_ROLL_M', 'LDG_ROLL'], 0), 0)
      }))
      .filter(r => r.w > 0 && r.to > 0 && r.ldg > 0);

    if (!baseRows.length) {
      return { success: false, error: 'No base roll rows found for aircraft type: ' + aircraftType };
    }

    const base = _interpBaseRoll(baseRows, weightKg);
    const densityAltitudeFt = _densityAltitudeFt(elevationFt, qnhHpa, tempC);

    const daTo = _nearestMultiplier(perfTable.rows, ['DA_FT', 'DA'], densityAltitudeFt, ['TO_MULTIPLIER'], 1);
    const daLdg = _nearestMultiplier(perfTable.rows, ['DA_FT', 'DA'], densityAltitudeFt, ['LDG_MULTIPLIER'], 1);

    const isWet = surfaceCondition === 'WET';
    const surfaceTo = _surfaceMultiplier(perfTable.rows, surface, [isWet ? 'TAKEOFF_WET' : 'TAKEOFF_DRY'], 1);
    const surfaceLdg = _surfaceMultiplier(perfTable.rows, surface, [isWet ? 'LANDING_WET' : 'LANDING_DRY'], 1);

    const slopeProfile = _normalizeSlopeProfile(slopeProfileRaw, runwayLengthM, slopePercent);

    let headTailComponent = 0;
    let crosswindComponent = 0;
    let effectiveWindType = windType;
    let effectiveWindKts = windSpeedKts;

    if (runwayHeadingDeg > 0 && windDirectionDeg >= 0) {
      const rawDelta = ((windDirectionDeg - runwayHeadingDeg) % 360 + 360) % 360;
      const delta = rawDelta > 180 ? rawDelta - 360 : rawDelta;
      const rad = delta * Math.PI / 180;

      headTailComponent = windSpeedKts * Math.cos(rad);
      crosswindComponent = Math.abs(windSpeedKts * Math.sin(rad));
      effectiveWindType = headTailComponent >= 0 ? 'HEAD' : 'TAIL';
      effectiveWindKts = Math.abs(headTailComponent);
    }

    const windColTo = effectiveWindType === 'TAIL' ? 'TAKEOFF_TAIL' : 'TAKEOFF_HEAD';
    const windColLdg = effectiveWindType === 'TAIL' ? 'LANDING_TAIL' : 'LANDING_HEAD';
      const windTo  = effectiveWindKts === 0 ? 1 : _nearestMultiplier(perfTable.rows, ['WIND_KTS', 'WIND'], effectiveWindKts, [windColTo], 1);
      const windLdg = effectiveWindKts === 0 ? 1 : _nearestMultiplier(perfTable.rows, ['WIND_KTS', 'WIND'], effectiveWindKts, [windColLdg], 1);

    const humidity = _nearestMultiplier(perfTable.rows, ['HUMIDITY'], humidityPct, ['HUMIDITY_FACTOR'], 1);
    const flap = _nearestMultiplier(perfTable.rows, ['FLAP_SETTING'], flapSetting, ['FLAP_FACTOR'], 1);

    const toNoSlope = base.to * daTo * surfaceTo * flap * humidity * windTo;
    let slopeTo = 1;
    let effectiveSlopeTakeoff = slopePercent;
    for (let i = 0; i < 3; i++) {
      const estTakeoff = toNoSlope * slopeTo;
      effectiveSlopeTakeoff = _effectiveSlopeOverDistance(slopeProfile, Math.min(estTakeoff, runwayLengthM));
      const slopeAbsTo = Math.abs(effectiveSlopeTakeoff);
      const slopeTargetTo = _slopeLookupTarget(perfTable.rows, slopeAbsTo);
      slopeTo = _nearestMultiplier(
        perfTable.rows,
        ['SLOPE'],
        slopeTargetTo,
        [effectiveSlopeTakeoff >= 0 ? 'TAKEOFF_UP' : 'TAKEOFF_DOWN'],
        1
      );
    }

    const ldgNoSlope = base.ldg * daLdg * surfaceLdg * windLdg;
    let slopeLdg = 1;
    let effectiveSlopeLanding = slopePercent;
    for (let i = 0; i < 3; i++) {
      const estLanding = ldgNoSlope * slopeLdg;
      effectiveSlopeLanding = _effectiveSlopeOverDistance(slopeProfile, Math.min(estLanding, runwayLengthM));
      const slopeAbsLdg = Math.abs(effectiveSlopeLanding);
      const slopeTargetLdg = _slopeLookupTarget(perfTable.rows, slopeAbsLdg);
      slopeLdg = _nearestMultiplier(
        perfTable.rows,
        ['SLOPE'],
        slopeTargetLdg,
        [effectiveSlopeLanding >= 0 ? 'LANDING_UP' : 'LANDING_DOWN'],
        1
      );
    }

    const takeoffRollM = toNoSlope * slopeTo;
    const landingRollM = ldgNoSlope * slopeLdg;

    const combinedRollM = takeoffRollM + landingRollM;
    const halfTakeoffRollM = takeoffRollM * 0.5;
    const takeoff75PctThresholdM = runwayLengthM * 0.75;

    const warnings = [];
    let blocking = false;

    if (takeoffRollM > takeoff75PctThresholdM) {
      warnings.push('Takeoff roll exceeds 75% of runway length.');
    }
    if (landingRollM > runwayLengthM) {
      warnings.push('Landing roll exceeds runway length.');
    }
    if (combinedRollM > runwayLengthM) {
      warnings.push('Takeoff + landing roll is greater than runway length.');
      blocking = true;
    }
    const abortUsesHalfTakeoff = combinedRollM > runwayLengthM;
    const abortPointM = abortUsesHalfTakeoff ? halfTakeoffRollM : takeoffRollM;

    return {
      success: true,
      baseTakeoffM: Math.round(base.to),
      baseLandingM: Math.round(base.ldg),
      densityAltitudeFt: Math.round(densityAltitudeFt),
      takeoffRollM: Math.round(takeoffRollM),
      landingRollM: Math.round(landingRollM),
      combinedRollM: Math.round(combinedRollM),
      halfTakeoffRollM: Math.round(halfTakeoffRollM),
      takeoff75PctThresholdM: Math.round(takeoff75PctThresholdM),
      abortPointM: Math.round(abortPointM),
      abortUsesHalfTakeoff: abortUsesHalfTakeoff,
      headTailComponentKts: Number(headTailComponent.toFixed(1)),
      crosswindComponentKts: Number(crosswindComponent.toFixed(1)),
      effectiveWindKts: Number(effectiveWindKts.toFixed(1)),
      effectiveWindType: effectiveWindType,
      warnings: warnings,
      blocking: blocking,
      factors: {
        daTakeoff: daTo,
        daLanding: daLdg,
        surfaceTakeoff: surfaceTo,
        surfaceLanding: surfaceLdg,
        slopeTakeoff: slopeTo,
        slopeLanding: slopeLdg,
        effectiveSlopeTakeoff: Number(effectiveSlopeTakeoff.toFixed(2)),
        effectiveSlopeLanding: Number(effectiveSlopeLanding.toFixed(2)),
        flap: flap,
        humidity: humidity,
        windTakeoff: windTo,
        windLanding: windLdg
      }
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

  /* ==================================================
  SETUP: Create RUNWAY_WALKTHROUGH_LOG sheet if needed
  =================================================== */

  function setupRunwayWalkthroughLog_() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let logSheet = ss.getSheetByName('RUNWAY_WALKTHROUGH_LOG');
    
      if (!logSheet) {
        logSheet = ss.insertSheet('RUNWAY_WALKTHROUGH_LOG', ss.getSheets().length);
      }
    
      // Clear and setup headers
      logSheet.clear();
      const headers = [
        'STAGING_ID',
        'ICAO',
        'RWY_IDENT',
        'PILOT_NAME',
        'PILOT_EMAIL',
        'WALK_DATE',
        'NOTES',
        'FEATURES_JSON',
        'STATUS',
        'SUPERVISOR_NAME',
        'SUPERVISOR_NOTES',
        'APPROVED_AT',
        'PUBLISHED_AT'
      ];
    
      logSheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight('bold')
        .setBackground('#073763')
        .setFontColor('white');
    
      // Set column widths for readability
      logSheet.setColumnWidth(1, 200);  // STAGING_ID
      logSheet.setColumnWidth(2, 100);  // ICAO
      logSheet.setColumnWidth(3, 100);  // RWY_IDENT
      logSheet.setColumnWidth(4, 150);  // PILOT_NAME
      logSheet.setColumnWidth(8, 300);  // FEATURES_JSON (wider for JSON)
    
      // Freeze header row
      logSheet.setFrozenRows(1);
    
      return { success: true, message: 'RUNWAY_WALKTHROUGH_LOG sheet initialized' };
    } catch (e) {
      return { success: false, error: e && e.message ? e.message : String(e) };
    }
  }

  /* ================================================== 
  RUNWAY WALKTHROUGH & FEATURE EDITS
  ================================================== */

function submitRunwayWalkthrough_(payload) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = ss.getSheetByName('RUNWAY_WALKTHROUGH_LOG');
      if (!logSheet) return { success: false, error: 'RUNWAY_WALKTHROUGH_LOG sheet not found' };

      const pilotName = String(payload && payload.pilotName || '').trim() || 'Unknown Pilot';
      const pilotEmail = String(payload && payload.pilotEmail || '').trim() || '';
      const icao = String(payload && payload.icao || '').trim().toUpperCase();
      const rwyIdent = String(payload && payload.rwyIdent || '').trim();
      const notes = String(payload && payload.notes || '').trim();
      const featuresJson = JSON.stringify(Array.isArray(payload && payload.features) ? payload.features : []);
    
      if (!icao || !rwyIdent) {
        return { success: false, error: 'ICAO and runway identifier required' };
      }

      const stagingId = 'STAG_' + new Date().getTime() + '_' + icao + '_' + rwyIdent.replace(/\s+/g, '');
      const now = new Date().toISOString();
    
      const headers = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
      const newRow = [
        stagingId,
        icao,
        rwyIdent,
        pilotName,
        pilotEmail,
        now,
        notes,
        featuresJson,
        'PENDING',
        '',
        '',
        '',
        ''
      ];
    
      logSheet.appendRow(newRow);
    
      return {
        success: true,
        stagingId: stagingId,
        message: 'Runway walkthrough submitted for review'
      };
    } catch (e) {
      return { success: false, error: e && e.message ? e.message : String(e) };
    }
  }

  function approveRunwayWalkthrough_(stagingId, supervisorName, supervisorNotes) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = ss.getSheetByName('RUNWAY_WALKTHROUGH_LOG');
      const dbSheet = ss.getSheetByName('DB_Airports');
    
      if (!logSheet || !dbSheet) return { success: false, error: 'Required sheets not found' };
    
      // Find staging record
      const logData = logSheet.getDataRange().getValues();
      let stagingRow = -1;
      for (let i = 1; i < logData.length; i++) {
        if (String(logData[i][0]).trim() === stagingId) {
          stagingRow = i;
          break;
        }
      }
    
      if (stagingRow < 0) return { success: false, error: 'Staging record not found' };
    
      const now = new Date().toISOString();
      const icao = String(logData[stagingRow][1]).trim().toUpperCase();
      const rwyIdent = String(logData[stagingRow][2]).trim();
      const featuresJson = String(logData[stagingRow][7]).trim();
      const pilotName = String(logData[stagingRow][3]).trim();
    
      // Update log sheet: mark as APPROVED
      logSheet.getRange(stagingRow + 1, 9).setValue('APPROVED');  // STATUS
      logSheet.getRange(stagingRow + 1, 10).setValue(supervisorName);  // SUPERVISOR_NAME
      logSheet.getRange(stagingRow + 1, 11).setValue(supervisorNotes);  // SUPERVISOR_NOTES
      logSheet.getRange(stagingRow + 1, 12).setValue(now);  // APPROVED_AT
    
      // Update DB_Airports: merge into KNOWN_FEATURES JSON
      const dbData = dbSheet.getDataRange().getValues();
      const headers = dbData[0];
      const knownFeaturesIdx = headers.findIndex(h => String(h).trim().toUpperCase().indexOf('KNOWN_FEATURE') >= 0);
    
      if (knownFeaturesIdx < 0) return { success: false, error: 'KNOWN_FEATURES column not found' };
    
      for (let i = 1; i < dbData.length; i++) {
        if (String(dbData[i][0]).trim().toUpperCase() === icao && String(dbData[i][1]).trim() === rwyIdent) {
          // Found matching runway; merge features into KNOWN_FEATURES
          const currentKnownStr = String(dbData[i][knownFeaturesIdx]).trim();
          let currentKnown = {};
        
          try {
            currentKnown = currentKnownStr ? JSON.parse(currentKnownStr) : {};
          } catch (e) {
            // If it was an old-format array, convert to new format
            try {
              const oldArray = JSON.parse(currentKnownStr);
              if (Array.isArray(oldArray)) {
                currentKnown = { features: oldArray };
              }
            } catch (e2) {
              currentKnown = {};
            }
          }
        
          // Ensure features array exists
          if (!currentKnown.features) currentKnown.features = [];
          if (!Array.isArray(currentKnown.features)) currentKnown.features = [];
        
          // Parse staged features and merge (simple replace for now)
          try {
            const stagedFeatures = JSON.parse(featuresJson);
            if (Array.isArray(stagedFeatures)) {
              currentKnown.features = stagedFeatures;
            }
          } catch (e) {}
        
          // Update metadata
          currentKnown.lastWalked = {
            date: now,
            pilotName: pilotName,
            notes: String(logData[stagingRow][6]).trim(),
            approved: true
          };
        
          // Write back to DB
          dbSheet.getRange(i + 1, knownFeaturesIdx + 1).setValue(JSON.stringify(currentKnown));
        
          // Update log: mark as PUBLISHED
          logSheet.getRange(stagingRow + 1, 9).setValue('PUBLISHED');
          logSheet.getRange(stagingRow + 1, 13).setValue(now);  // PUBLISHED_AT
        
          return { success: true, message: 'Runway features approved and published' };
        }
      }
    
      return { success: false, error: 'Runway not found in DB_Airports' };
    } catch (e) {
      return { success: false, error: e && e.message ? e.message : String(e) };
    }
  }

  function rejectRunwayWalkthrough_(stagingId, supervisorName, supervisorNotes) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = ss.getSheetByName('RUNWAY_WALKTHROUGH_LOG');
    
      if (!logSheet) return { success: false, error: 'RUNWAY_WALKTHROUGH_LOG sheet not found' };
    
      const logData = logSheet.getDataRange().getValues();
      let stagingRow = -1;
      for (let i = 1; i < logData.length; i++) {
        if (String(logData[i][0]).trim() === stagingId) {
          stagingRow = i;
          break;
        }
      }
    
      if (stagingRow < 0) return { success: false, error: 'Staging record not found' };
    
      const now = new Date().toISOString();
      logSheet.getRange(stagingRow + 1, 9).setValue('REJECTED');
      logSheet.getRange(stagingRow + 1, 10).setValue(supervisorName);
      logSheet.getRange(stagingRow + 1, 11).setValue(supervisorNotes);
      logSheet.getRange(stagingRow + 1, 12).setValue(now);
    
      return { success: true, message: 'Runway walkthrough rejected' };
    } catch (e) {
      return { success: false, error: e && e.message ? e.message : String(e) };
    }
  }
/* ==================================================
   RUNWAY DATABASE — save full entry from Pilot App
   Public (no underscore) so google.script.run can reach it.
   ================================================== */
function saveRunwayDatabaseEntry(icao, rwyIdent, featureData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('DB_Airports');
    if (!sheet) return { success: false, error: 'DB_Airports sheet not found' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(function(h) { return String(h || '').trim().toUpperCase(); });

    const icaoCol     = headers.indexOf('ICAO');
    const rwyCol      = headers.indexOf('RWY_IDENT');
    const featuresCol = headers.indexOf('KNOWN_FEATURES');

    if (icaoCol < 0 || rwyCol < 0) {
      return { success: false, error: 'ICAO or RWY_IDENT column not found in DB_Airports' };
    }
    if (featuresCol < 0) {
      return { success: false, error: 'KNOWN_FEATURES column not found in DB_Airports' };
    }

    const targetIcao = String(icao || '').trim().toUpperCase();
    const targetRwy  = String(rwyIdent || '').trim().toUpperCase();

    let foundRow = -1;
    for (let i = 1; i < data.length; i++) {
      const rowIcao = String(data[i][icaoCol] || '').trim().toUpperCase();
      const rowRwy  = String(data[i][rwyCol]  || '').trim().toUpperCase();
      if (rowIcao === targetIcao && rowRwy === targetRwy) {
        foundRow = i;
        break;
      }
    }

    if (foundRow < 0) {
      return { success: false, error: 'Runway not found: ' + targetIcao + ' / ' + targetRwy };
    }

    // Merge with any existing metadata we want to preserve (lastWalked, published, staged)
    let existing = {};
    try {
      const raw = String(data[foundRow][featuresCol] || '').trim();
      if (raw) {
        const parsed = JSON.parse(raw);
        existing = Array.isArray(parsed) ? { features: parsed } : (parsed || {});
      }
    } catch(e) { /* ignore parse errors — overwrite cleanly */ }

    const merged = Object.assign({}, existing, featureData, {
      updatedAt: new Date().toISOString()
    });

    // If featureData has slopeSegments, also write them in slopeProfile-compatible format
    // so Tab4 can read them directly.
    if (featureData && Array.isArray(featureData.slopeSegments) && featureData.slopeSegments.length) {
      merged.slopeProfile = featureData.slopeSegments.map(function(s) {
        return { distance: s.distanceM || 0, slope: s.slope || 0 };
      });
    }

    sheet.getRange(foundRow + 1, featuresCol + 1).setValue(JSON.stringify(merged));
    SpreadsheetApp.flush();

    return { success: true, message: 'Runway database entry saved: ' + targetIcao + ' RWY ' + targetRwy };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _ensureRunwayWalkthroughLogSchema_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('RUNWAY_WALKTHROUGH_LOG');
  if (!logSheet) {
    setupRunwayWalkthroughLog_();
    logSheet = ss.getSheetByName('RUNWAY_WALKTHROUGH_LOG');
  }
  if (!logSheet) throw new Error('RUNWAY_WALKTHROUGH_LOG sheet not found');

  const required = [
    'STAGING_ID', 'ICAO', 'RWY_IDENT', 'PILOT_NAME', 'PILOT_EMAIL', 'WALK_DATE',
    'NOTES', 'FEATURES_JSON', 'STATUS', 'SUPERVISOR_NAME', 'SUPERVISOR_NOTES',
    'APPROVED_AT', 'PUBLISHED_AT', 'ENTRY_KIND', 'SURVEY_JSON', 'OFFICIAL_JSON',
    'CAPTURE_SUMMARY_JSON', 'DEVICE_INFO_JSON'
  ];

  const lastCol = Math.max(logSheet.getLastColumn(), 1);
  const existing = logSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || '').trim();
  });

  required.forEach(function(name) {
    if (existing.indexOf(name) < 0) {
      existing.push(name);
      logSheet.getRange(1, existing.length).setValue(name);
    }
  });

  const index = {};
  existing.forEach(function(h, i) { index[h] = i; });
  return { sheet: logSheet, headers: existing, idx: index };
}

function _runwayDbFindCols_(headers) {
  const norm = headers.map(function(h) { return String(h || '').trim().toUpperCase(); });
  function find() {
    for (var i = 0; i < arguments.length; i++) {
      var k = String(arguments[i] || '').trim().toUpperCase();
      var at = norm.indexOf(k);
      if (at >= 0) return at;
    }
    return -1;
  }
  return {
    icao: find('ICAO'),
    runway: find('RWY_IDENT', 'RWY', 'RUNWAY', 'RUNWAY_DESIGNATOR'),
    knownFeatures: find('KNOWN_FEATURES', 'FEATURES'),
    heading: find('RUNWAY_HEADING', 'HEADING'),
    length: find('LENGTH_OFFICIAL', 'LENGTH_METERS', 'LENGTH_M'),
    width: find('WIDTH_OFFICIAL', 'WIDTH_METERS', 'WIDTH_M'),
    surface: find('SURFACE_ACTUAL', 'SURFACE_OFFICIAL', 'SURFACE')
  };
}

function _parseJsonLoose_(raw, fallback) {
  try {
    if (raw == null || raw === '') return fallback;
    var parsed = typeof raw === 'string' ? JSON.parse(raw) : raw;
    return parsed == null ? fallback : parsed;
  } catch (e) {
    return fallback;
  }
}

function _dbAirportOfficialSnapshot_(row, cols) {
  return {
    lengthM: Number(row[cols.length] || 0) || 0,
    widthM: Number(row[cols.width] || 0) || 0,
    surface: String(row[cols.surface] || '').trim(),
    headingDeg: Number(row[cols.heading] || 0) || 0
  };
}

function _findAirportPhotoFolder_(icao) {
  var code = String(icao || '').trim().toUpperCase();
  if (!code) return null;
  var byExact = DriveApp.getFoldersByName(code);
  if (byExact && byExact.hasNext()) return byExact.next();
  var byPrefixed = DriveApp.getFoldersByName('AIRPORT_' + code);
  if (byPrefixed && byPrefixed.hasNext()) return byPrefixed.next();
  return null;
}

function getAirportPhotoFolderLink(icao) {
  try {
    var folder = _findAirportPhotoFolder_(icao);
    if (!folder) {
      return { success: false, error: 'No Drive folder found for airport ' + String(icao || '').trim().toUpperCase() };
    }
    return {
      success: true,
      icao: String(icao || '').trim().toUpperCase(),
      folderId: folder.getId(),
      folderName: folder.getName(),
      url: folder.getUrl()
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getAirportContacts(icao) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(APP_SHEETS.CONTACTS || 'DB_Contacts');
    if (!sh) return { success: false, error: 'DB_Contacts sheet not found' };
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: false, error: 'DB_Contacts is empty' };
    var headers = data[0].map(function(h) { return String(h || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '_'); });
    var icaoIdx = -1;
    var candidates = ['ICAO', 'C_DIGO', 'CODIGO', 'CDIGO', 'C__DIGO'];
    for (var ci = 0; ci < candidates.length; ci++) {
      var idx = headers.indexOf(candidates[ci]);
      if (idx >= 0) { icaoIdx = idx; break; }
    }
    if (icaoIdx < 0) return { success: false, error: 'ICAO/Código column not found in DB_Contacts. Headers: ' + headers.join(', ') };
    var target = String(icao || '').trim().toUpperCase();
    var row = null;
    var rowNumber = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][icaoIdx] || '').trim().toUpperCase() === target) { row = data[i]; rowNumber = i + 1; break; }
    }
    if (!row) return { success: true, found: false, icao: target };
    var g = function(names) {
      var arr = Array.isArray(names) ? names : [names];
      for (var ni = 0; ni < arr.length; ni++) {
        var hi = headers.indexOf(arr[ni]);
        if (hi >= 0) return String(row[hi] || '').trim();
      }
      return '';
    };
    return {
      success: true,
      found: true,
      rowNumber: rowNumber,
      fields: _toolsRowPayloadFromHeaders_(data[0], row),
      icao: target,
      municipio: g(['MUNIC_PIO___ALDEIA', 'MUNICIPIO___ALDEIA', 'MUNICIPIO', 'MUN']),
      hasFuel: g(['POSSUI_COMBUST_VEL_', 'POSSUI_COMBUSTIVEL_', 'POSSUI_COMBUSTIVEL']),
      permanencia: g(['PERMAN_NCIA_', 'PERMANENCIA_', 'PERMANENCIA']),
      contato: g(['CONTATO']),
      celular: g(['CELULAR']),
      telefone2: g(['TELEFONE_2', 'TELEFONE2']),
      temsCadastro: g(['TEMOS_CADASTRO_', 'TEMOS_CADASTRO']),
      anotacoes: g(['ANOTA__ES', 'ANOTACOES', 'ANOTAÇÕES'])
    };
  } catch(e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function setAirportFuelAvailability(icao, hasFuel) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var code = String(icao || '').trim().toUpperCase();
    if (!code) return { success: false, error: 'ICAO required' };

    var enabled = (hasFuel === true || String(hasFuel || '').trim().toLowerCase() === 'true' || String(hasFuel || '').trim() === '1');
    var airportValue = enabled ? 'Available' : 'None';
    var contactsValue = enabled ? 'YES' : 'NO';

    var airportSheet = getRequiredSheet_(ss, APP_SHEETS.AIRPORTS, 'setAirportFuelAvailability');
    var airportData = airportSheet.getDataRange().getValues();
    if (airportData.length < 2) return { success: false, error: 'DB_Airports is empty' };

    var airportHeaders = airportData[0].map(function(h) { return _toolsNormHeader_(h); });
    var icaoIdx = airportHeaders.indexOf('ICAO');
    var fuelIdx = airportHeaders.indexOf('FUEL_AVAILABLE');
    if (icaoIdx < 0) return { success: false, error: 'ICAO column not found in DB_Airports' };
    if (fuelIdx < 0) return { success: false, error: 'FUEL_AVAILABLE column not found in DB_Airports' };

    var airportRow = 0;
    for (var i = 1; i < airportData.length; i++) {
      if (String(airportData[i][icaoIdx] || '').trim().toUpperCase() === code) { airportRow = i + 1; break; }
    }
    if (!airportRow) return { success: false, error: 'Airport not found in DB_Airports: ' + code };

    airportSheet.getRange(airportRow, fuelIdx + 1).setValue(airportValue);

    var contactsUpdated = false;
    try {
      var contactSheet = ss.getSheetByName(APP_SHEETS.CONTACTS || 'DB_Contacts');
      if (contactSheet) {
        var contactData = contactSheet.getDataRange().getValues();
        if (contactData.length >= 2) {
          var contactHeaders = contactData[0].map(function(h) { return _toolsNormHeader_(h); });

          var contactIcaoIdx = -1;
          var contactIcaoCandidates = ['ICAO', 'C_DIGO', 'CODIGO', 'CDIGO', 'C__DIGO'];
          for (var ci = 0; ci < contactIcaoCandidates.length; ci++) {
            var cidx = contactHeaders.indexOf(contactIcaoCandidates[ci]);
            if (cidx >= 0) { contactIcaoIdx = cidx; break; }
          }

          var contactFuelIdx = -1;
          var contactFuelCandidates = ['POSSUI_COMBUST_VEL_', 'POSSUI_COMBUSTIVEL_', 'POSSUI_COMBUSTIVEL'];
          for (var fi = 0; fi < contactFuelCandidates.length; fi++) {
            var fidx = contactHeaders.indexOf(contactFuelCandidates[fi]);
            if (fidx >= 0) { contactFuelIdx = fidx; break; }
          }

          if (contactIcaoIdx >= 0 && contactFuelIdx >= 0) {
            for (var r = 1; r < contactData.length; r++) {
              if (String(contactData[r][contactIcaoIdx] || '').trim().toUpperCase() === code) {
                contactSheet.getRange(r + 1, contactFuelIdx + 1).setValue(contactsValue);
                contactsUpdated = true;
                break;
              }
            }
          }
        }
      }
    } catch (contactErr) {
      // Keep airport update authoritative even if contacts sync cannot be completed.
    }

    return {
      success: true,
      icao: code,
      fuelAvailable: airportValue,
      contactsUpdated: contactsUpdated
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _toolsHeaderValueByCandidates_(headers, row, candidates, fallback) {
  var norms = (headers || []).map(function(h) { return _toolsNormHeader_(h); });
  var list = Array.isArray(candidates) ? candidates : [candidates];
  for (var i = 0; i < list.length; i++) {
    var idx = norms.indexOf(_toolsNormHeader_(list[i]));
    if (idx >= 0) {
      var value = row && idx < row.length ? row[idx] : '';
      if (String(value || '').trim() !== '') return value;
    }
  }
  return fallback;
}

function getAirportRecordForTools(icao) {
  try {
    var code = String(icao || '').trim().toUpperCase();
    if (!code) return { success: false, error: 'ICAO required' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.AIRPORTS, 'getAirportRecordForTools');
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: false, error: 'DB_Airports is empty' };

    var headers = data[0];
    var norms = headers.map(function(h) { return _toolsNormHeader_(h); });
    var icaoIdx = norms.indexOf('ICAO');
    if (icaoIdx < 0) return { success: false, error: 'ICAO column not found in DB_Airports' };

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][icaoIdx] || '').trim().toUpperCase() === code) {
        var row = data[i];
        return {
          success: true,
          found: true,
          icao: code,
          rowNumber: i + 1,
          fields: _toolsRowPayloadFromHeaders_(headers, row),
          fuelAvailable: _toolsHeaderValueByCandidates_(headers, row, ['FUEL_AVAILABLE'], ''),
          runwaySurface: _toolsHeaderValueByCandidates_(headers, row, ['SURFACE_ACTUAL', 'RUNWAY_SURFACE_ACTUAL', 'SURFACE_OFFICIAL', 'RUNWAY_SURFACE', 'SURFACE_TYPE', 'SURFACE', 'SURFACE_CONDITION', 'RUNWAY_SURFACE_CONDITION', 'CONDITION', 'SURFACE_STATUS'], ''),
          headers: headers.map(function(h) { return String(h || ''); })
        };
      }
    }
    return { success: true, found: false, icao: code };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getAllAirportsForTools() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.AIRPORTS, 'getAllAirportsForTools');
    var data = sh.getDataRange().getValues();
    if (data.length < 2) {
      return { success: true, airports: [], headers: [] };
    }

    var headers = data[0];
    var icaoIdx = _toolsHeaderIndexFromCandidates_(headers, [
      'ICAO',
      'OACI',
      'ICAO_ID',
      'CÓDIGO',
      'CODIGO',
      'C_DIGO',
      'CDIGO',
      'C__DIGO',
      'CÓDIGO OACI',
      'CODIGO OACI',
      'CÓDIGO_OACI',
      'CODIGO_OACI',
      'AERODROMO_OACI',
      'AERODROMO_ICAO'
    ]);
    if (icaoIdx < 0) return { success: false, error: 'ICAO/code column not found in DB_Airports' };

    var airports = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var code = String(row[icaoIdx] || '').trim().toUpperCase();
      if (!code) continue;
      airports.push({
        rowNumber: i + 1,
        icao: code,
        fields: _toolsRowPayloadFromHeaders_(headers, row),
        nome: _toolsHeaderValueByCandidates_(headers, row, ['NOME', 'NAME', 'AIRPORT_NAME'], ''),
        lat: _toolsHeaderValueByCandidates_(headers, row, ['LATITUDE', 'LAT'], ''),
        lon: _toolsHeaderValueByCandidates_(headers, row, ['LONGITUDE', 'LON', 'LONG'], ''),
        fuelAvailable: _toolsHeaderValueByCandidates_(headers, row, ['FUEL_AVAILABLE'], ''),
        runwayIdent: _toolsHeaderValueByCandidates_(headers, row, ['RWY_IDENT', 'RUNWAY_IDENT', 'RUNWAY'], ''),
        runwayLength: _toolsHeaderValueByCandidates_(headers, row, ['LENGTH_OFFICIAL', 'RUNWAY_LENGTH', 'LENGTH_M'], ''),
        runwayWidth: _toolsHeaderValueByCandidates_(headers, row, ['WIDTH_OFFICIAL', 'RUNWAY_WIDTH', 'WIDTH_M'], ''),
        runwaySurfaceActual: _toolsHeaderValueByCandidates_(headers, row, ['SURFACE_ACTUAL', 'RUNWAY_SURFACE_ACTUAL', 'SURFACE_OFFICIAL', 'RUNWAY_SURFACE', 'SURFACE_TYPE', 'SURFACE'], ''),
        runwaySurfaceCondition: _toolsHeaderValueByCandidates_(headers, row, ['SURFACE_CONDITION', 'RUNWAY_SURFACE_CONDITION', 'CONDITION', 'SURFACE_STATUS'], ''),
        runwaySlopePercent: _toolsHeaderValueByCandidates_(headers, row, ['SLOPE_PERCENT', 'RUNWAY_SLOPE', 'SLOPE'], ''),
        elevationFt: _toolsHeaderValueByCandidates_(headers, row, ['ELEVATION', 'ELEVATION_FT'], ''),
        pilotNotes: _toolsHeaderValueByCandidates_(headers, row, ['PILOT_NOTES', 'NOTES'], '')
      });
    }

    return {
      success: true,
      syncedAtMs: Date.now(),
      airports: airports,
      headers: headers.map(function(h) { return String(h || ''); })
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _toolsIsoDateFromAny_(value) {
  var dt = value instanceof Date ? new Date(value.getTime()) : new Date(value);
  if (!(dt instanceof Date) || isNaN(dt.getTime())) return '';
  return dt.getFullYear() + '-' + ('0' + (dt.getMonth() + 1)).slice(-2) + '-' + ('0' + dt.getDate()).slice(-2);
}

function _toolsDiscrepancyStatusLabel_(status) {
  var key = String(status || '').trim().toUpperCase();
  if (key === 'OPEN') return 'Open';
  if (key === 'DEFERRED_50_HOUR') return 'Deferred 50 Hour';
  if (key === 'DEFERRED_100_HOUR') return 'Deferred 100 Hour';
  if (key === 'DEFERRED_TO_DATE') return 'Deferred To Date';
  if (key === 'CLOSED') return 'Closed';
  if (key === 'CANCELED') return 'Canceled';
  return key || 'Open';
}

function _toolsIsOpenDiscrepancyStatus_(status) {
  var key = String(status || '').trim().toUpperCase();
  return key === 'OPEN' || key === 'DEFERRED_50_HOUR' || key === 'DEFERRED_100_HOUR' || key === 'DEFERRED_TO_DATE';
}

function _toolsParseOpenSquawksJson_(raw) {
  var text = String(raw == null ? '' : raw).trim();
  if (!text) return [];
  try {
    var parsed = JSON.parse(text);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function _toolsNextDiscrepancyId_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var year = String((new Date()).getFullYear());
    var key = 'OPEN_SQUAWK_SEQ_' + year;
    var current = parseInt(props.getProperty(key) || '0', 10);
    if (!isFinite(current) || current < 0) current = 0;
    var next = current + 1;
    props.setProperty(key, String(next));
    return 'SQ-' + year + '-' + ('0000' + next).slice(-4);
  } finally {
    lock.releaseLock();
  }
}

function _toolsEnsureOpenSquawksColumn_(sheet, headers, norms) {
  var idx = norms.indexOf('OPEN_SQUAWKS');
  if (idx >= 0) return idx;
  var newCol = headers.length + 1;
  sheet.getRange(1, newCol).setValue('OPEN_SQUAWKS');
  return newCol - 1;
}

function _toolsFindAircraftRowByReg_(rows, regIdx, aircraftReg) {
  var target = String(aircraftReg || '').trim().toUpperCase();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][regIdx] || '').trim().toUpperCase() === target) return i;
  }
  return -1;
}

function _toolsActorEmail_(fallback) {
  var fb = String(fallback || '').trim();
  try {
    var email = String(Session.getActiveUser().getEmail() || '').trim();
    if (email) return email;
  } catch (e) {}
  return fb || 'unknown@local';
}

function _toolsAuditDiscrepancy_(userEmail, aircraftReg, discrepancyId, action, oldValue, newValue, note) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var audit = ss.getSheetByName(APP_SHEETS.AUDIT);
    if (!audit) return;
    audit.appendRow([
      new Date(),
      String(userEmail || ''),
      String(aircraftReg || ''),
      String(action || ''),
      String(oldValue == null ? '' : oldValue),
      String(newValue == null ? '' : newValue),
      String((discrepancyId ? ('Discrepancy ' + discrepancyId + ' | ') : '') + (note || ''))
    ]);
  } catch (e) {}
}

function reportAircraftDiscrepancy(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var aircraftReg = String(body.aircraftReg || '').trim().toUpperCase();
    var tachRaw = String(body.tachAtReport || '').trim();
    var reportedBy = String(body.reportedBy || '').trim();
    var reportDate = String(body.reportDate || '').trim();
    var description = String(body.description || '').trim();
    var status = String(body.status || 'OPEN').trim().toUpperCase();
    var deferredUntilDate = String(body.deferredUntilDate || '').trim();
    var deferredUntilTach = String(body.deferredUntilTach || '').trim();
    var sourceType = String(body.sourceType || 'manual').trim().toLowerCase();
    var sourceFlightLegId = String(body.sourceFlightLegId || '').trim();

    if (!aircraftReg || !tachRaw || !reportedBy || !reportDate || !description) {
      return { success: false, error: 'Aircraft, tach, reporter, date, and description are required.' };
    }

    var validStatuses = {
      OPEN: true,
      DEFERRED_50_HOUR: true,
      DEFERRED_100_HOUR: true,
      DEFERRED_TO_DATE: true,
      CLOSED: true,
      CANCELED: true
    };
    if (!validStatuses[status]) status = 'OPEN';

    var tach = parseFloat(tachRaw);
    if (!isFinite(tach)) return { success: false, error: 'Invalid tach value.' };
    var isoReportDate = _toolsIsoDateFromAny_(reportDate);
    if (!isoReportDate) return { success: false, error: 'Invalid report date.' };

    if (status === 'DEFERRED_50_HOUR') deferredUntilTach = String((tach + 50).toFixed(1));
    if (status === 'DEFERRED_100_HOUR') deferredUntilTach = String((tach + 100).toFixed(1));
    if (status === 'DEFERRED_TO_DATE' && !_toolsIsoDateFromAny_(deferredUntilDate)) {
      return { success: false, error: 'Deferred To Date requires a valid deferred date.' };
    }
    if (status !== 'DEFERRED_TO_DATE') deferredUntilDate = '';
    if (!(status === 'DEFERRED_50_HOUR' || status === 'DEFERRED_100_HOUR' || status === 'DEFERRED_TO_DATE')) deferredUntilTach = '';

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.AIRCRAFT, 'reportAircraftDiscrepancy');
    var data = sh.getDataRange().getValues();
    if (!data || data.length < 2) return { success: false, error: 'DB_Aircraft has no rows.' };
    var headers = data[0];
    var norms = headers.map(function(h) { return _toolsNormHeader_(h); });
    var regIdx = norms.indexOf('REGISTRATION');
    if (regIdx < 0) return { success: false, error: 'REGISTRATION column not found in DB_Aircraft.' };
    var rowIdx = _toolsFindAircraftRowByReg_(data, regIdx, aircraftReg);
    if (rowIdx < 0) return { success: false, error: 'Aircraft not found: ' + aircraftReg };
    var openIdx = _toolsEnsureOpenSquawksColumn_(sh, headers, norms);
    var list = _toolsParseOpenSquawksJson_(sh.getRange(rowIdx + 1, openIdx + 1).getValue());

    var actorEmail = _toolsActorEmail_(reportedBy);
    var nowIso = new Date().toISOString();
    var item = {
      id: _toolsNextDiscrepancyId_(),
      aircraftReg: aircraftReg,
      description: description,
      status: status,
      tachAtReport: tach,
      reportDate: isoReportDate,
      reportedBy: reportedBy,
      createdAt: nowIso,
      sourceType: sourceType,
      sourceFlightLegId: sourceFlightLegId,
      deferredUntilTach: deferredUntilTach ? parseFloat(deferredUntilTach) : '',
      deferredUntilDate: _toolsIsoDateFromAny_(deferredUntilDate) || '',
      mechanicEvaluatedBy: '',
      mechanicEvaluatedAt: '',
      mechanicNotes: ''
    };
    if (status === 'CLOSED') item.closedAt = nowIso;
    if (status === 'CANCELED') item.canceledAt = nowIso;

    list.push(item);
    sh.getRange(rowIdx + 1, openIdx + 1).setValue(JSON.stringify(list));
    var openCount = list.filter(function(sq) { return _toolsIsOpenDiscrepancyStatus_(sq && sq.status); }).length;

    _toolsAuditDiscrepancy_(
      actorEmail,
      aircraftReg,
      item.id,
      'DISCREPANCY_CREATE',
      '',
      status,
      'Source=' + sourceType + '; Reporter=' + reportedBy + '; Tach=' + tach + '; Date=' + isoReportDate
    );

    MailApp.sendEmail({
      to: 'tecnico.mx@asasdesocorro.org.br',
      subject: '[NEW DISCREPANCY] ' + item.id + ' - ' + aircraftReg,
      body: [
        'A new discrepancy was reported.',
        '',
        'ID: ' + item.id,
        'Aircraft: ' + aircraftReg,
        'Status: ' + _toolsDiscrepancyStatusLabel_(status),
        'Date: ' + isoReportDate,
        'Tach: ' + tach,
        'Reporter: ' + reportedBy,
        'Source: ' + sourceType,
        sourceFlightLegId ? ('Flight Leg: ' + sourceFlightLegId) : '',
        item.deferredUntilTach !== '' ? ('Deferred Until Tach: ' + item.deferredUntilTach) : '',
        item.deferredUntilDate ? ('Deferred Until Date: ' + item.deferredUntilDate) : '',
        '',
        'Description:',
        description
      ].filter(Boolean).join('\n')
    });

    return { success: true, item: item, openCount: openCount, actor: actorEmail };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function updateAircraftDiscrepancyStatus(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var aircraftReg = String(body.aircraftReg || '').trim().toUpperCase();
    var discrepancyId = String(body.discrepancyId || '').trim();
    var status = String(body.status || '').trim().toUpperCase();
    var updatedBy = String(body.updatedBy || '').trim() || 'System';
    if (!aircraftReg || !discrepancyId || !status) return { success: false, error: 'aircraftReg, discrepancyId, and status are required.' };

    var validStatuses = { OPEN: true, DEFERRED_50_HOUR: true, DEFERRED_100_HOUR: true, DEFERRED_TO_DATE: true, CLOSED: true, CANCELED: true };
    if (!validStatuses[status]) return { success: false, error: 'Invalid status.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.AIRCRAFT, 'updateAircraftDiscrepancyStatus');
    var data = sh.getDataRange().getValues();
    if (!data || data.length < 2) return { success: false, error: 'DB_Aircraft has no rows.' };
    var headers = data[0];
    var norms = headers.map(function(h) { return _toolsNormHeader_(h); });
    var regIdx = norms.indexOf('REGISTRATION');
    if (regIdx < 0) return { success: false, error: 'REGISTRATION column not found in DB_Aircraft.' };
    var rowIdx = _toolsFindAircraftRowByReg_(data, regIdx, aircraftReg);
    if (rowIdx < 0) return { success: false, error: 'Aircraft not found: ' + aircraftReg };
    var openIdx = _toolsEnsureOpenSquawksColumn_(sh, headers, norms);
    var rows = _toolsParseOpenSquawksJson_(sh.getRange(rowIdx + 1, openIdx + 1).getValue());

    var target = null;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i] && rows[i].id || '').trim() === discrepancyId) { target = rows[i]; break; }
    }
    if (!target) return { success: false, error: 'Discrepancy not found: ' + discrepancyId };

    var actorEmail = _toolsActorEmail_(updatedBy);
    var previousStatus = String(target.status || '');
    target.status = status;
    target.updatedAt = new Date().toISOString();
    target.updatedBy = actorEmail;
    if (status === 'CLOSED') target.closedAt = target.updatedAt;
    if (status === 'CANCELED') target.canceledAt = target.updatedAt;
    if (status === 'OPEN') {
      target.closedAt = '';
      target.canceledAt = '';
    }
    if (status === 'DEFERRED_50_HOUR' || status === 'DEFERRED_100_HOUR') {
      var tach = parseFloat(target.tachAtReport || 0);
      if (isFinite(tach) && tach > 0) target.deferredUntilTach = parseFloat((tach + (status === 'DEFERRED_50_HOUR' ? 50 : 100)).toFixed(1));
      target.deferredUntilDate = '';
    }

    sh.getRange(rowIdx + 1, openIdx + 1).setValue(JSON.stringify(rows));
    var openCount = rows.filter(function(sq) { return _toolsIsOpenDiscrepancyStatus_(sq && sq.status); }).length;
    _toolsAuditDiscrepancy_(
      actorEmail,
      aircraftReg,
      discrepancyId,
      'DISCREPANCY_STATUS_UPDATE',
      previousStatus,
      status,
      'UpdatedBy=' + actorEmail
    );
    return { success: true, discrepancy: target, openCount: openCount, actor: actorEmail };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _toolsAddDebriefSquawksToAircraft_(opts) {
  try {
    var body = (opts && typeof opts === 'object') ? opts : {};
    var aircraftReg = String(body.aircraftReg || '').trim().toUpperCase();
    if (!aircraftReg) return { added: 0, ids: [] };
    var squawks = String(body.squawks || '')
      .split(/[\n,;]+/)
      .map(function(s) { return String(s || '').trim(); })
      .filter(Boolean);
    if (!squawks.length) return { added: 0, ids: [] };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.AIRCRAFT, '_toolsAddDebriefSquawksToAircraft_');
    var data = sh.getDataRange().getValues();
    if (!data || data.length < 2) return { added: 0, ids: [] };
    var headers = data[0];
    var norms = headers.map(function(h) { return _toolsNormHeader_(h); });
    var regIdx = norms.indexOf('REGISTRATION');
    if (regIdx < 0) return { added: 0, ids: [] };
    var rowIdx = _toolsFindAircraftRowByReg_(data, regIdx, aircraftReg);
    if (rowIdx < 0) return { added: 0, ids: [] };
    var openIdx = _toolsEnsureOpenSquawksColumn_(sh, headers, norms);
    var existing = _toolsParseOpenSquawksJson_(sh.getRange(rowIdx + 1, openIdx + 1).getValue());
    var existingKeys = {};
    existing.forEach(function(item) {
      var key = String(item && item.sourceFlightLegId || '') + '|' + String(item && item.description || '').trim().toUpperCase();
      if (key !== '|') existingKeys[key] = true;
    });

    var nowIso = new Date().toISOString();
    var reportDate = _toolsIsoDateFromAny_(body.reportDate) || _toolsIsoDateFromAny_(nowIso);
    var tach = parseFloat(body.tachAtReport || 0);
    var addedItems = [];
    squawks.forEach(function(desc) {
      var key = String(body.sourceFlightLegId || '') + '|' + String(desc || '').trim().toUpperCase();
      if (existingKeys[key]) return;
      var item = {
        id: _toolsNextDiscrepancyId_(),
        aircraftReg: aircraftReg,
        description: desc,
        status: 'OPEN',
        tachAtReport: isFinite(tach) && tach > 0 ? tach : '',
        reportDate: reportDate,
        reportedBy: String(body.reportedBy || 'Pilot').trim() || 'Pilot',
        createdAt: nowIso,
        sourceType: 'debrief',
        sourceFlightLegId: String(body.sourceFlightLegId || '').trim(),
        deferredUntilTach: '',
        deferredUntilDate: '',
        mechanicEvaluatedBy: '',
        mechanicEvaluatedAt: '',
        mechanicNotes: ''
      };
      existing.push(item);
      addedItems.push(item);
      existingKeys[key] = true;
    });
    if (!addedItems.length) return { added: 0, ids: [] };

    sh.getRange(rowIdx + 1, openIdx + 1).setValue(JSON.stringify(existing));

    var debriefActor = _toolsActorEmail_(String(body.reportedBy || 'Pilot').trim() || 'Pilot');
    addedItems.forEach(function(item) {
      _toolsAuditDiscrepancy_(
        debriefActor,
        aircraftReg,
        item.id,
        'DISCREPANCY_CREATE',
        '',
        item.status,
        'Source=debrief; FlightLeg=' + String(body.sourceFlightLegId || '')
      );
    });

    MailApp.sendEmail({
      to: 'tecnico.mx@asasdesocorro.org.br',
      subject: '[NEW DISCREPANCY] ' + aircraftReg + ' (' + addedItems.length + ' from debrief)',
      body: [
        'New discrepancies were reported from debrief.',
        '',
        'Aircraft: ' + aircraftReg,
        'Reporter: ' + (String(body.reportedBy || 'Pilot').trim() || 'Pilot'),
        'Date: ' + reportDate,
        body.sourceFlightLegId ? ('Flight Leg: ' + body.sourceFlightLegId) : '',
        '',
        'Items:',
        addedItems.map(function(item) { return '- ' + item.id + ': ' + item.description; }).join('\n')
      ].filter(Boolean).join('\n')
    });

    return { added: addedItems.length, ids: addedItems.map(function(i) { return i.id; }) };
  } catch (e) {
    return { added: 0, ids: [], error: e && e.message ? e.message : String(e) };
  }
}

function _mxEnsureFrameworkSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function ensure(sheetName, headers) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh) sh = ss.insertSheet(sheetName);
    if (sh.getLastRow() < 1) {
      sh.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setBackground('#1f2937')
        .setFontColor('white')
        .setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  }

  ensure(APP_SHEETS.MAINT_TEMPLATES, [
    'TASK_CODE', 'TASK_NAME', 'AIRCRAFT_TYPE', 'CATEGORY', 'REFERENCE',
    'INTERVAL_HOURS', 'INTERVAL_DAYS', 'ACTIVE', 'SOURCE', 'CAMO_KEY', 'NOTES', 'CREATED_AT', 'UPDATED_AT'
  ]);
  ensure(APP_SHEETS.MAINT_ASSIGNMENTS, [
    'ASSIGNMENT_ID', 'AIRCRAFT_REG', 'TASK_CODE', 'TASK_NAME', 'CATEGORY', 'REFERENCE',
    'INTERVAL_HOURS', 'INTERVAL_DAYS', 'START_TACH', 'START_DATE', 'ACTIVE', 'SOURCE', 'CAMO_KEY', 'NOTES', 'CREATED_AT', 'UPDATED_AT'
  ]);
  ensure(APP_SHEETS.MAINT_LOG, [
    'LOG_ID', 'ASSIGNMENT_ID', 'AIRCRAFT_REG', 'COMPLETED_DATE', 'COMPLETED_TACH',
    'REFERENCE_DOC', 'PERFORMED_BY', 'REMARKS', 'CREATED_AT'
  ]);
}

function setupMaintenanceSchedulingFramework() {
  try {
    _mxEnsureFrameworkSheets_();
    return { success: true, message: 'Maintenance framework ready.' };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _mxNormHeaderMap_(headers) {
  var map = {};
  headers.forEach(function(h, idx) { map[_toolsNormHeader_(h)] = idx; });
  return map;
}

function _mxIsoDate_(value) {
  var dt = value instanceof Date ? new Date(value.getTime()) : new Date(value);
  if (!(dt instanceof Date) || isNaN(dt.getTime())) return '';
  return dt.getFullYear() + '-' + ('0' + (dt.getMonth() + 1)).slice(-2) + '-' + ('0' + dt.getDate()).slice(-2);
}

function _mxAddDaysIso_(isoDate, days) {
  var dt = new Date(isoDate);
  if (isNaN(dt.getTime())) return '';
  dt.setDate(dt.getDate() + days);
  return _mxIsoDate_(dt);
}

function _mxDaysRemaining_(isoDate) {
  if (!isoDate) return '';
  var due = new Date(isoDate + 'T00:00:00');
  if (isNaN(due.getTime())) return '';
  var now = new Date();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  return Math.floor((due - today) / 86400000);
}

function _mxParseNum_(value) {
  var n = parseFloat(value);
  return isFinite(n) ? n : '';
}

function _mxNextId_(sequenceKey, prefix) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var year = String((new Date()).getFullYear());
    var key = sequenceKey + '_' + year;
    var current = parseInt(props.getProperty(key) || '0', 10);
    if (!isFinite(current) || current < 0) current = 0;
    var next = current + 1;
    props.setProperty(key, String(next));
    return prefix + '-' + year + '-' + ('0000' + next).slice(-4);
  } finally {
    lock.releaseLock();
  }
}

function _mxAircraftTachByReg_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getRequiredSheet_(ss, APP_SHEETS.AIRCRAFT, '_mxAircraftTachByReg_');
  var data = sh.getDataRange().getValues();
  if (!data || data.length < 2) return {};
  var headers = data[0];
  var idx = _mxNormHeaderMap_(headers);
  var regIdx = idx.REGISTRATION;
  var tachIdx = idx.CURRENT_TACH;
  var out = {};
  if (!(regIdx >= 0)) return out;
  for (var i = 1; i < data.length; i++) {
    var reg = String(data[i][regIdx] || '').trim().toUpperCase();
    if (!reg) continue;
    out[reg] = _mxParseNum_(data[i][tachIdx]);
  }
  return out;
}

function _mxLatestLogByAssignment_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getRequiredSheet_(ss, APP_SHEETS.MAINT_LOG, '_mxLatestLogByAssignment_');
  var data = sh.getDataRange().getValues();
  if (!data || data.length < 2) return {};
  var headers = data[0];
  var idx = _mxNormHeaderMap_(headers);
  var out = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var assignmentId = String(row[idx.ASSIGNMENT_ID] || '').trim();
    if (!assignmentId) continue;
    var completedDate = _mxIsoDate_(row[idx.COMPLETED_DATE]);
    var createdAt = _mxIsoDate_(row[idx.CREATED_AT]);
    var sortKey = completedDate || createdAt || '';
    if (!sortKey) continue;

    var prev = out[assignmentId];
    if (!prev || String(prev.sortKey) < String(sortKey)) {
      out[assignmentId] = {
        completedDate: completedDate,
        completedTach: _mxParseNum_(row[idx.COMPLETED_TACH]),
        sortKey: sortKey
      };
    }
  }
  return out;
}

function _mxDueState_(hoursRemaining, daysRemaining, hasHours, hasDays, thresholdHours, thresholdDays) {
  var hourOverdue = hasHours && hoursRemaining !== '' && hoursRemaining <= 0;
  var dayOverdue = hasDays && daysRemaining !== '' && daysRemaining <= 0;
  if (hourOverdue || dayOverdue) return 'OVERDUE';

  var hourSoon = hasHours && hoursRemaining !== '' && hoursRemaining <= thresholdHours;
  var daySoon = hasDays && daysRemaining !== '' && daysRemaining <= thresholdDays;
  if (hourSoon || daySoon) return 'DUE_SOON';

  if ((hasHours && hoursRemaining === '') || (hasDays && daysRemaining === '')) return 'UNKNOWN';
  return 'OK';
}

function getMaintenanceScheduleData(payload) {
  try {
    _mxEnsureFrameworkSheets_();
    var body = (payload && typeof payload === 'object') ? payload : {};
    var filterReg = String(body.aircraftReg || '').trim().toUpperCase();
    var thresholdHours = _mxParseNum_(body.thresholdHours);
    var thresholdDays = _mxParseNum_(body.thresholdDays);
    if (thresholdHours === '') thresholdHours = 10;
    if (thresholdDays === '') thresholdDays = 30;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.MAINT_ASSIGNMENTS, 'getMaintenanceScheduleData');
    var data = sh.getDataRange().getValues();
    if (!data || data.length < 2) return { success: true, rows: [], events: [], summary: { total: 0, overdue: 0, dueSoon: 0 } };

    var headers = data[0];
    var idx = _mxNormHeaderMap_(headers);
    var tachByReg = _mxAircraftTachByReg_();
    var latestLog = _mxLatestLogByAssignment_();
    var rows = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var active = String(row[idx.ACTIVE] || 'Y').trim().toUpperCase();
      if (active === 'N') continue;

      var aircraftReg = String(row[idx.AIRCRAFT_REG] || '').trim().toUpperCase();
      if (!aircraftReg) continue;
      if (filterReg && aircraftReg !== filterReg) continue;

      var assignmentId = String(row[idx.ASSIGNMENT_ID] || '').trim();
      var intervalHours = _mxParseNum_(row[idx.INTERVAL_HOURS]);
      var intervalDays = _mxParseNum_(row[idx.INTERVAL_DAYS]);
      var hasHours = intervalHours !== '' && intervalHours > 0;
      var hasDays = intervalDays !== '' && intervalDays > 0;
      if (!hasHours && !hasDays) continue;

      var startTach = _mxParseNum_(row[idx.START_TACH]);
      var startDate = _mxIsoDate_(row[idx.START_DATE]);
      var last = latestLog[assignmentId] || null;
      var baseTach = (last && last.completedTach !== '') ? last.completedTach : startTach;
      var baseDate = (last && last.completedDate) ? last.completedDate : startDate;

      var nextDueTach = '';
      if (hasHours && baseTach !== '') nextDueTach = parseFloat((baseTach + intervalHours).toFixed(1));
      var nextDueDate = '';
      if (hasDays && baseDate) nextDueDate = _mxAddDaysIso_(baseDate, intervalDays);

      var currentTach = tachByReg[aircraftReg];
      var hoursRemaining = '';
      if (hasHours && nextDueTach !== '' && currentTach !== '') {
        hoursRemaining = parseFloat((nextDueTach - currentTach).toFixed(1));
      }
      var daysRemaining = hasDays ? _mxDaysRemaining_(nextDueDate) : '';
      var dueState = _mxDueState_(hoursRemaining, daysRemaining, hasHours, hasDays, thresholdHours, thresholdDays);

      rows.push({
        assignmentId: assignmentId,
        aircraftReg: aircraftReg,
        taskCode: String(row[idx.TASK_CODE] || '').trim(),
        taskName: String(row[idx.TASK_NAME] || '').trim(),
        category: String(row[idx.CATEGORY] || '').trim(),
        reference: String(row[idx.REFERENCE] || '').trim(),
        intervalHours: hasHours ? intervalHours : '',
        intervalDays: hasDays ? intervalDays : '',
        startTach: startTach,
        startDate: startDate,
        currentTach: currentTach,
        lastCompletedDate: last ? last.completedDate : '',
        lastCompletedTach: last ? last.completedTach : '',
        nextDueTach: nextDueTach,
        nextDueDate: nextDueDate,
        hoursRemaining: hoursRemaining,
        daysRemaining: daysRemaining,
        dueState: dueState,
        notes: String(row[idx.NOTES] || '').trim()
      });
    }

    rows.sort(function(a, b) {
      var rank = { OVERDUE: 0, DUE_SOON: 1, UNKNOWN: 2, OK: 3 };
      var ra = rank[a.dueState] != null ? rank[a.dueState] : 9;
      var rb = rank[b.dueState] != null ? rank[b.dueState] : 9;
      if (ra !== rb) return ra - rb;
      return String(a.aircraftReg + '|' + a.taskName).localeCompare(String(b.aircraftReg + '|' + b.taskName));
    });

    var events = rows
      .filter(function(r) { return !!r.nextDueDate; })
      .map(function(r) {
        return {
          id: r.assignmentId,
          title: r.aircraftReg + ' - ' + (r.taskName || r.taskCode || 'Maintenance Item'),
          start: r.nextDueDate,
          allDay: true,
          type: 'maintenance',
          status: r.dueState,
          aircraftReg: r.aircraftReg
        };
      });

    var summary = { total: rows.length, overdue: 0, dueSoon: 0, ok: 0, unknown: 0 };
    rows.forEach(function(r) {
      if (r.dueState === 'OVERDUE') summary.overdue++;
      else if (r.dueState === 'DUE_SOON') summary.dueSoon++;
      else if (r.dueState === 'OK') summary.ok++;
      else summary.unknown++;
    });

    return { success: true, rows: rows, events: events, summary: summary };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function saveMaintenanceTemplate(payload) {
  try {
    _mxEnsureFrameworkSheets_();
    var body = (payload && typeof payload === 'object') ? payload : {};
    var taskCode = String(body.taskCode || '').trim().toUpperCase();
    var taskName = String(body.taskName || '').trim();
    var aircraftType = String(body.aircraftType || '').trim();
    if (!taskCode || !taskName) return { success: false, error: 'taskCode and taskName are required.' };

    var intervalHours = _mxParseNum_(body.intervalHours);
    var intervalDays = _mxParseNum_(body.intervalDays);
    if (intervalHours === '' && intervalDays === '') return { success: false, error: 'Provide intervalHours and/or intervalDays.' };

    var nowIso = _mxIsoDate_(new Date());
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.MAINT_TEMPLATES, 'saveMaintenanceTemplate');
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var idx = _mxNormHeaderMap_(headers);

    for (var i = 1; i < data.length; i++) {
      var existingCode = String(data[i][idx.TASK_CODE] || '').trim().toUpperCase();
      var existingType = String(data[i][idx.AIRCRAFT_TYPE] || '').trim().toUpperCase();
      if (existingCode === taskCode && existingType === String(aircraftType || '').trim().toUpperCase()) {
        return { success: false, error: 'Template already exists for this taskCode and aircraftType.' };
      }
    }

    var row = new Array(headers.length).fill('');
    row[idx.TASK_CODE] = taskCode;
    row[idx.TASK_NAME] = taskName;
    row[idx.AIRCRAFT_TYPE] = aircraftType;
    row[idx.CATEGORY] = String(body.category || '').trim();
    row[idx.REFERENCE] = String(body.reference || '').trim();
    row[idx.INTERVAL_HOURS] = intervalHours === '' ? '' : intervalHours;
    row[idx.INTERVAL_DAYS] = intervalDays === '' ? '' : intervalDays;
    row[idx.ACTIVE] = String(body.active || 'Y').trim().toUpperCase() === 'N' ? 'N' : 'Y';
    row[idx.SOURCE] = String(body.source || 'operator').trim();
    row[idx.CAMO_KEY] = String(body.camoKey || '').trim();
    row[idx.NOTES] = String(body.notes || '').trim();
    row[idx.CREATED_AT] = nowIso;
    row[idx.UPDATED_AT] = nowIso;

    sh.appendRow(row);
    return { success: true, taskCode: taskCode };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function saveMaintenanceAssignment(payload) {
  try {
    _mxEnsureFrameworkSheets_();
    var body = (payload && typeof payload === 'object') ? payload : {};
    var aircraftReg = String(body.aircraftReg || '').trim().toUpperCase();
    var taskCode = String(body.taskCode || '').trim().toUpperCase();
    var taskName = String(body.taskName || '').trim();
    if (!aircraftReg || !taskName) return { success: false, error: 'aircraftReg and taskName are required.' };

    var intervalHours = _mxParseNum_(body.intervalHours);
    var intervalDays = _mxParseNum_(body.intervalDays);
    if (intervalHours === '' && intervalDays === '') return { success: false, error: 'Provide intervalHours and/or intervalDays.' };

    var nowIso = _mxIsoDate_(new Date());
    var assignmentId = _mxNextId_('MAINT_ASSIGNMENT_SEQ', 'MXA');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.MAINT_ASSIGNMENTS, 'saveMaintenanceAssignment');
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var idx = _mxNormHeaderMap_(headers);

    var row = new Array(headers.length).fill('');
    row[idx.ASSIGNMENT_ID] = assignmentId;
    row[idx.AIRCRAFT_REG] = aircraftReg;
    row[idx.TASK_CODE] = taskCode;
    row[idx.TASK_NAME] = taskName;
    row[idx.CATEGORY] = String(body.category || '').trim();
    row[idx.REFERENCE] = String(body.reference || '').trim();
    row[idx.INTERVAL_HOURS] = intervalHours === '' ? '' : intervalHours;
    row[idx.INTERVAL_DAYS] = intervalDays === '' ? '' : intervalDays;
    row[idx.START_TACH] = _mxParseNum_(body.startTach);
    row[idx.START_DATE] = _mxIsoDate_(body.startDate || new Date()) || _mxIsoDate_(new Date());
    row[idx.ACTIVE] = String(body.active || 'Y').trim().toUpperCase() === 'N' ? 'N' : 'Y';
    row[idx.SOURCE] = String(body.source || 'operator').trim();
    row[idx.CAMO_KEY] = String(body.camoKey || '').trim();
    row[idx.NOTES] = String(body.notes || '').trim();
    row[idx.CREATED_AT] = nowIso;
    row[idx.UPDATED_AT] = nowIso;

    sh.appendRow(row);
    return { success: true, assignmentId: assignmentId };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function completeMaintenanceAssignment(payload) {
  try {
    _mxEnsureFrameworkSheets_();
    var body = (payload && typeof payload === 'object') ? payload : {};
    var assignmentId = String(body.assignmentId || '').trim();
    var aircraftReg = String(body.aircraftReg || '').trim().toUpperCase();
    if (!assignmentId || !aircraftReg) return { success: false, error: 'assignmentId and aircraftReg are required.' };

    var completedDate = _mxIsoDate_(body.completedDate || new Date());
    if (!completedDate) return { success: false, error: 'Invalid completedDate.' };
    var completedTach = _mxParseNum_(body.completedTach);
    var logId = _mxNextId_('MAINT_LOG_SEQ', 'MXL');
    var nowIso = _mxIsoDate_(new Date());

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSh = getRequiredSheet_(ss, APP_SHEETS.MAINT_LOG, 'completeMaintenanceAssignment');
    var logData = logSh.getDataRange().getValues();
    var logHeaders = logData[0];
    var logIdx = _mxNormHeaderMap_(logHeaders);

    var logRow = new Array(logHeaders.length).fill('');
    logRow[logIdx.LOG_ID] = logId;
    logRow[logIdx.ASSIGNMENT_ID] = assignmentId;
    logRow[logIdx.AIRCRAFT_REG] = aircraftReg;
    logRow[logIdx.COMPLETED_DATE] = completedDate;
    logRow[logIdx.COMPLETED_TACH] = completedTach === '' ? '' : completedTach;
    logRow[logIdx.REFERENCE_DOC] = String(body.referenceDoc || '').trim();
    logRow[logIdx.PERFORMED_BY] = String(body.performedBy || '').trim();
    logRow[logIdx.REMARKS] = String(body.remarks || '').trim();
    logRow[logIdx.CREATED_AT] = nowIso;
    logSh.appendRow(logRow);

    var asgSh = getRequiredSheet_(ss, APP_SHEETS.MAINT_ASSIGNMENTS, 'completeMaintenanceAssignment');
    var asgData = asgSh.getDataRange().getValues();
    var asgHeaders = asgData[0];
    var asgIdx = _mxNormHeaderMap_(asgHeaders);
    for (var i = 1; i < asgData.length; i++) {
      if (String(asgData[i][asgIdx.ASSIGNMENT_ID] || '').trim() !== assignmentId) continue;
      asgSh.getRange(i + 1, asgIdx.UPDATED_AT + 1).setValue(nowIso);
      break;
    }

    return { success: true, logId: logId };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function applyMaintenanceTemplateToAircraft(payload) {
  try {
    _mxEnsureFrameworkSheets_();
    var body = (payload && typeof payload === 'object') ? payload : {};
    var aircraftReg = String(body.aircraftReg || '').trim().toUpperCase();
    var aircraftType = String(body.aircraftType || '').trim().toUpperCase();
    if (!aircraftReg) return { success: false, error: 'aircraftReg is required.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tplSh = getRequiredSheet_(ss, APP_SHEETS.MAINT_TEMPLATES, 'applyMaintenanceTemplateToAircraft');
    var asgSh = getRequiredSheet_(ss, APP_SHEETS.MAINT_ASSIGNMENTS, 'applyMaintenanceTemplateToAircraft');

    var tplData = tplSh.getDataRange().getValues();
    if (!tplData || tplData.length < 2) return { success: true, inserted: 0, skipped: 0 };
    var tplIdx = _mxNormHeaderMap_(tplData[0]);

    var asgData = asgSh.getDataRange().getValues();
    var asgIdx = _mxNormHeaderMap_(asgData[0]);
    var existing = {};
    for (var i = 1; i < asgData.length; i++) {
      var reg = String(asgData[i][asgIdx.AIRCRAFT_REG] || '').trim().toUpperCase();
      var code = String(asgData[i][asgIdx.TASK_CODE] || '').trim().toUpperCase();
      var active = String(asgData[i][asgIdx.ACTIVE] || 'Y').trim().toUpperCase();
      if (reg && code && active !== 'N') existing[reg + '|' + code] = true;
    }

    var startDate = _mxIsoDate_(body.startDate || new Date()) || _mxIsoDate_(new Date());
    var startTach = _mxParseNum_(body.startTach);
    var nowIso = _mxIsoDate_(new Date());
    var rowsToAdd = [];
    var skipped = 0;

    for (var t = 1; t < tplData.length; t++) {
      var tr = tplData[t];
      var activeTpl = String(tr[tplIdx.ACTIVE] || 'Y').trim().toUpperCase();
      if (activeTpl === 'N') continue;
      var tplType = String(tr[tplIdx.AIRCRAFT_TYPE] || '').trim().toUpperCase();
      if (aircraftType && tplType && tplType !== aircraftType) continue;

      var taskCode = String(tr[tplIdx.TASK_CODE] || '').trim().toUpperCase();
      var taskName = String(tr[tplIdx.TASK_NAME] || '').trim();
      if (!taskName) continue;

      var key = aircraftReg + '|' + taskCode;
      if (taskCode && existing[key]) { skipped++; continue; }

      var newRow = new Array(asgData[0].length).fill('');
      newRow[asgIdx.ASSIGNMENT_ID] = _mxNextId_('MAINT_ASSIGNMENT_SEQ', 'MXA');
      newRow[asgIdx.AIRCRAFT_REG] = aircraftReg;
      newRow[asgIdx.TASK_CODE] = taskCode;
      newRow[asgIdx.TASK_NAME] = taskName;
      newRow[asgIdx.CATEGORY] = String(tr[tplIdx.CATEGORY] || '').trim();
      newRow[asgIdx.REFERENCE] = String(tr[tplIdx.REFERENCE] || '').trim();
      newRow[asgIdx.INTERVAL_HOURS] = _mxParseNum_(tr[tplIdx.INTERVAL_HOURS]);
      newRow[asgIdx.INTERVAL_DAYS] = _mxParseNum_(tr[tplIdx.INTERVAL_DAYS]);
      newRow[asgIdx.START_TACH] = startTach;
      newRow[asgIdx.START_DATE] = startDate;
      newRow[asgIdx.ACTIVE] = 'Y';
      newRow[asgIdx.SOURCE] = String(tr[tplIdx.SOURCE] || 'template').trim();
      newRow[asgIdx.CAMO_KEY] = String(tr[tplIdx.CAMO_KEY] || '').trim();
      newRow[asgIdx.NOTES] = String(tr[tplIdx.NOTES] || '').trim();
      newRow[asgIdx.CREATED_AT] = nowIso;
      newRow[asgIdx.UPDATED_AT] = nowIso;
      rowsToAdd.push(newRow);
    }

    if (rowsToAdd.length) {
      asgSh.getRange(asgSh.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    }
    return { success: true, inserted: rowsToAdd.length, skipped: skipped };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function seedU206MaintenanceBaselineTemplates() {
  try {
    _mxEnsureFrameworkSheets_();
    var base = [
      { taskCode: 'INSP-50H', taskName: '50 Hour Inspection', aircraftType: 'U206', category: 'INSPECTION', reference: 'Operator Baseline', intervalHours: 50 },
      { taskCode: 'INSP-100H', taskName: '100 Hour Inspection', aircraftType: 'U206', category: 'INSPECTION', reference: 'Operator Baseline', intervalHours: 100 },
      { taskCode: 'INSP-200H', taskName: '200 Hour Inspection', aircraftType: 'U206', category: 'INSPECTION', reference: 'Operator Baseline', intervalHours: 200 },
      { taskCode: 'INSP-ANNUAL', taskName: 'Annual Inspection', aircraftType: 'U206', category: 'INSPECTION', reference: 'Operator Baseline', intervalDays: 365 },
      { taskCode: 'GPS-DB', taskName: 'GPS Database Update', aircraftType: 'U206', category: 'NAV_DB', reference: 'Avionics Program', intervalDays: 28 },
      { taskCode: 'BAT-CHECK', taskName: 'Battery Check', aircraftType: 'U206', category: 'BATTERY', reference: 'Battery Program', intervalDays: 180 },
      { taskCode: 'FIRE-EXT', taskName: 'Fire Extinguisher Check', aircraftType: 'U206', category: 'SAFETY', reference: 'Safety Equipment Program', intervalDays: 180 }
    ];

    var inserted = 0;
    var skipped = 0;
    base.forEach(function(item) {
      var res = saveMaintenanceTemplate(item);
      if (res && res.success) inserted++;
      else skipped++;
    });

    return { success: true, inserted: inserted, skipped: skipped };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getToolsReports(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var startDate = (function(value) {
      if (value instanceof Date && !isNaN(value.getTime())) {
        var out = new Date(value.getTime());
        out.setHours(0, 0, 0, 0);
        return out;
      }
      var raw = String(value || '').trim();
      if (!raw) return null;
      if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
        var parts = raw.split('-');
        var dt = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
        dt.setHours(0, 0, 0, 0);
        return isNaN(dt.getTime()) ? null : dt;
      }
      var parsed = new Date(raw);
      if (isNaN(parsed.getTime())) return null;
      parsed.setHours(0, 0, 0, 0);
      return parsed;
    })(body.startDate);
    var endDate = (function(value) {
      if (value instanceof Date && !isNaN(value.getTime())) {
        var out = new Date(value.getTime());
        out.setHours(23, 59, 59, 999);
        return out;
      }
      var raw = String(value || '').trim();
      if (!raw) return null;
      if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
        var parts = raw.split('-');
        var dt = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
        dt.setHours(23, 59, 59, 999);
        return isNaN(dt.getTime()) ? null : dt;
      }
      var parsed = new Date(raw);
      if (isNaN(parsed.getTime())) return null;
      parsed.setHours(23, 59, 59, 999);
      return parsed;
    })(body.endDate);

    var readSheet = function(name) {
      var sh = ss.getSheetByName(name);
      if (!sh) return { headers: [], rows: [] };
      var values = sh.getDataRange().getValues();
      if (!values || !values.length) return { headers: [], rows: [] };
      return {
        headers: values[0].map(function(h) { return String(h || ''); }),
        rows: values.slice(1)
      };
    };
    var indexByAliases = function(headers, aliases, fallback) {
      var norms = (headers || []).map(function(h) { return _toolsNormHeader_(h); });
      var list = Array.isArray(aliases) ? aliases : [aliases];
      for (var i = 0; i < list.length; i++) {
        var idx = norms.indexOf(_toolsNormHeader_(list[i]));
        if (idx >= 0) return idx;
      }
      return typeof fallback === 'number' ? fallback : -1;
    };
    var rowDate = function(value) {
      if (value instanceof Date && !isNaN(value.getTime())) return new Date(value.getTime());
      var raw = String(value || '').trim();
      if (!raw) return null;
      var parsed = new Date(raw);
      return isNaN(parsed.getTime()) ? null : parsed;
    };
    var inRange = function(value) {
      var dt = rowDate(value);
      if (!dt) return (!startDate && !endDate);
      if (startDate && dt.getTime() < startDate.getTime()) return false;
      if (endDate && dt.getTime() > endDate.getTime()) return false;
      return true;
    };
    var num = function(value) {
      var n = parseFloat(value);
      return isNaN(n) ? 0 : n;
    };
    var isoDate = function(value) {
      var dt = rowDate(value);
      if (!dt) return '';
      return dt.getFullYear() + '-' + ('0' + (dt.getMonth() + 1)).slice(-2) + '-' + ('0' + dt.getDate()).slice(-2);
    };
    var isoDateTime = function(value) {
      var dt = rowDate(value);
      return dt ? dt.toISOString() : '';
    };
    var monthKey = function(value) {
      var dt = rowDate(value);
      if (!dt) return '';
      return dt.getFullYear() + '-' + ('0' + (dt.getMonth() + 1)).slice(-2);
    };

    var dispatchData = readSheet(APP_SHEETS.DISPATCH);
    var logData = readSheet(APP_SHEETS.LOG_FLIGHTS);
    var transData = readSheet(APP_SHEETS.TRANSACTIONS);
    var fundData = readSheet(APP_SHEETS.FUNDS);

    var dispatchFlightIdIdx = indexByAliases(dispatchData.headers, ['FLIGHT_ID'], DISPATCH_COL.FLIGHT_ID);
    var dispatchMissionIdIdx = indexByAliases(dispatchData.headers, ['MISSION_ID'], DISPATCH_COL.MISSION_ID);
    var dispatchDateIdx = indexByAliases(dispatchData.headers, ['DATE'], DISPATCH_COL.DATE);
    var dispatchAircraftIdx = indexByAliases(dispatchData.headers, ['AIRCRAFT'], DISPATCH_COL.AIRCRAFT);
    var dispatchPilotIdx = indexByAliases(dispatchData.headers, ['PILOT'], DISPATCH_COL.PILOT);
    var dispatchRouteIdx = indexByAliases(dispatchData.headers, ['ROUTE'], DISPATCH_COL.ROUTE);
    var dispatchTimeIdx = indexByAliases(dispatchData.headers, ['FLIGHT_TIME', 'TOTAL_TIME'], DISPATCH_COL.FLIGHT_TIME);
    var dispatchTypeIdx = indexByAliases(dispatchData.headers, ['TYPE', 'FLIGHT_TYPE'], DISPATCH_COL.TYPE);
    var dispatchStatusIdx = indexByAliases(dispatchData.headers, ['STATUS'], DISPATCH_COL.STATUS);

    var dispatchByFlight = {};
    var dispatchByMission = {};
    dispatchData.rows.forEach(function(row) {
      var flightId = dispatchFlightIdIdx >= 0 ? String(row[dispatchFlightIdIdx] || '').trim() : '';
      var missionId = dispatchMissionIdIdx >= 0 ? String(row[dispatchMissionIdIdx] || '').trim() : '';
      var item = {
        flightId: flightId,
        missionId: missionId,
        date: dispatchDateIdx >= 0 ? row[dispatchDateIdx] : '',
        aircraft: dispatchAircraftIdx >= 0 ? String(row[dispatchAircraftIdx] || '').trim() : '',
        pilot: dispatchPilotIdx >= 0 ? String(row[dispatchPilotIdx] || '').trim() : '',
        route: dispatchRouteIdx >= 0 ? String(row[dispatchRouteIdx] || '').trim() : '',
        totalTime: dispatchTimeIdx >= 0 ? num(row[dispatchTimeIdx]) : 0,
        type: dispatchTypeIdx >= 0 ? String(row[dispatchTypeIdx] || '').trim() : '',
        status: dispatchStatusIdx >= 0 ? String(row[dispatchStatusIdx] || '').trim() : ''
      };
      if (flightId) dispatchByFlight[flightId] = item;
      if (missionId && !dispatchByMission[missionId]) dispatchByMission[missionId] = item;
    });

    var fundNameIdx = indexByAliases(fundData.headers, ['NAME', 'FUND_NAME']);
    var fundBalanceIdx = indexByAliases(fundData.headers, ['CURRENT_BALANCE', 'BALANCE']);
    var fundBalanceMap = {};
    fundData.rows.forEach(function(row) {
      var fundName = fundNameIdx >= 0 ? String(row[fundNameIdx] || '').trim() : '';
      if (!fundName) return;
      fundBalanceMap[fundName] = fundBalanceIdx >= 0 ? num(row[fundBalanceIdx]) : '';
    });

    var transFlightIdx = indexByAliases(transData.headers, ['FLIGHT_ID', 'MISSION_ID'], 0);
    var transFundIdx = indexByAliases(transData.headers, ['FUND', 'FUND_NAME'], 1);
    var transPaxIdx = indexByAliases(transData.headers, ['PASSENGER_NAME', 'NAME'], 2);
    var transWeightIdx = indexByAliases(transData.headers, ['WEIGHT', 'WEIGHT_KG'], 4);
    var transAmountIdx = indexByAliases(transData.headers, ['CHARGED_AMOUNT', 'AMOUNT'], 6);

    var fundAgg = {};
    var totalFundUsage = 0;
    transData.rows.forEach(function(row) {
      var flightId = transFlightIdx >= 0 ? String(row[transFlightIdx] || '').trim() : '';
      if (!flightId) return;
      var missionId = typeof missionIdFromFlightLeg_ === 'function' ? missionIdFromFlightLeg_(flightId) : String(flightId).split('-').slice(0, 2).join('-');
      var dispatchInfo = dispatchByFlight[flightId] || dispatchByMission[missionId] || null;
      var reportDate = dispatchInfo ? dispatchInfo.date : '';
      if (!inRange(reportDate)) return;
      var fundName = transFundIdx >= 0 ? String(row[transFundIdx] || '').trim() : '';
      if (!fundName) fundName = 'Unassigned';
      if (!fundAgg[fundName]) {
        fundAgg[fundName] = {
          fund: fundName,
          amount: 0,
          balance: Object.prototype.hasOwnProperty.call(fundBalanceMap, fundName) ? fundBalanceMap[fundName] : '',
          legs: {},
          passengers: 0,
          weightKg: 0
        };
      }
      fundAgg[fundName].amount += transAmountIdx >= 0 ? num(row[transAmountIdx]) : 0;
      fundAgg[fundName].passengers += (transPaxIdx >= 0 && String(row[transPaxIdx] || '').trim()) ? 1 : 0;
      fundAgg[fundName].weightKg += transWeightIdx >= 0 ? num(row[transWeightIdx]) : 0;
      fundAgg[fundName].legs[flightId] = true;
      totalFundUsage += transAmountIdx >= 0 ? num(row[transAmountIdx]) : 0;
    });
    var fundRows = Object.keys(fundAgg).map(function(name) {
      var row = fundAgg[name];
      return {
        fund: row.fund,
        amount: row.amount,
        balance: row.balance,
        legs: Object.keys(row.legs).length,
        passengers: row.passengers,
        weightKg: row.weightKg
      };
    }).sort(function(a, b) { return b.amount - a.amount; });

    var logFlightIdIdx = indexByAliases(logData.headers, ['FLIGHT_ID'], LOG_FLIGHT_COL.FLIGHT_ID);
    var logDateIdx = indexByAliases(logData.headers, ['DATE'], LOG_FLIGHT_COL.DATE);
    var logPilotIdx = indexByAliases(logData.headers, ['PILOT'], LOG_FLIGHT_COL.PILOT);
    var logAircraftIdx = indexByAliases(logData.headers, ['ACFT', 'AIRCRAFT'], LOG_FLIGHT_COL.ACFT);
    var logFromIdx = indexByAliases(logData.headers, ['FROM', 'ORIGIN'], LOG_FLIGHT_COL.FROM);
    var logToIdx = indexByAliases(logData.headers, ['TO', 'DESTINATION'], LOG_FLIGHT_COL.TO);
    var logTimeIdx = indexByAliases(logData.headers, ['TOTAL_TIME', 'FLIGHT_TIME'], LOG_FLIGHT_COL.TOTAL_TIME);
    var logLdgsIdx = indexByAliases(logData.headers, ['NUMBER_LDGS', 'NUM_LDGS'], LOG_FLIGHT_COL.NUM_LDGS);
    var logAirborneIdx = indexByAliases(logData.headers, ['AIRBORNE'], LOG_FLIGHT_COL.AIRBORNE);
    var logLandedIdx = indexByAliases(logData.headers, ['LANDED'], LOG_FLIGHT_COL.LANDED);

    var flights = [];
    var pilotAgg = {};
    var monthlyAgg = {};
    logData.rows.forEach(function(row) {
      var flightId = logFlightIdIdx >= 0 ? String(row[logFlightIdIdx] || '').trim() : '';
      if (!flightId) return;
      var missionId = typeof missionIdFromFlightLeg_ === 'function' ? missionIdFromFlightLeg_(flightId) : String(flightId).split('-').slice(0, 2).join('-');
      var dispatchInfo = dispatchByFlight[flightId] || dispatchByMission[missionId] || {};
      var logDateValue = logDateIdx >= 0 ? row[logDateIdx] : '';
      var effectiveDate = logDateValue || dispatchInfo.date || '';
      if (!inRange(effectiveDate)) return;
      var totalTime = logTimeIdx >= 0 ? num(row[logTimeIdx]) : 0;
      if (!totalTime) totalTime = num(dispatchInfo.totalTime);
      var landings = logLdgsIdx >= 0 ? num(row[logLdgsIdx]) : 0;
      var pilotName = logPilotIdx >= 0 ? String(row[logPilotIdx] || '').trim() : '';
      if (!pilotName) pilotName = String(dispatchInfo.pilot || '').trim();
      var aircraft = logAircraftIdx >= 0 ? String(row[logAircraftIdx] || '').trim() : '';
      if (!aircraft) aircraft = String(dispatchInfo.aircraft || '').trim();
      var fromVal = logFromIdx >= 0 ? String(row[logFromIdx] || '').trim() : '';
      var toVal = logToIdx >= 0 ? String(row[logToIdx] || '').trim() : '';
      var typeVal = String(dispatchInfo.type || '').trim();
      var routeVal = String(dispatchInfo.route || '').trim() || [fromVal, toVal].filter(Boolean).join('-');
      var flightRow = {
        flightId: flightId,
        missionId: missionId,
        date: isoDate(effectiveDate),
        pilot: pilotName,
        aircraft: aircraft,
        from: fromVal,
        to: toVal,
        route: routeVal,
        type: typeVal,
        totalTime: totalTime,
        landings: landings,
        airborne: logAirborneIdx >= 0 ? isoDateTime(row[logAirborneIdx]) : '',
        landed: logLandedIdx >= 0 ? isoDateTime(row[logLandedIdx]) : ''
      };
      flights.push(flightRow);

      var pilotKey = String(pilotName || 'Unknown').trim() || 'Unknown';
      if (!pilotAgg[pilotKey]) pilotAgg[pilotKey] = { pilot: pilotKey, flights: 0, totalTime: 0, landings: 0 };
      pilotAgg[pilotKey].flights += 1;
      pilotAgg[pilotKey].totalTime += totalTime;
      pilotAgg[pilotKey].landings += landings;

      var month = monthKey(effectiveDate);
      var typeKey = typeVal || 'Unspecified';
      var monthlyKey = month + '|' + typeKey;
      if (month) {
        if (!monthlyAgg[monthlyKey]) monthlyAgg[monthlyKey] = { month: month, type: typeKey, flights: 0, totalTime: 0, landings: 0 };
        monthlyAgg[monthlyKey].flights += 1;
        monthlyAgg[monthlyKey].totalTime += totalTime;
        monthlyAgg[monthlyKey].landings += landings;
      }
    });
    flights.sort(function(a, b) { return String(b.date || '').localeCompare(String(a.date || '')) || String(b.flightId || '').localeCompare(String(a.flightId || '')); });

    var pilotRows = flights.map(function(row) {
      return {
        flightId: row.flightId,
        date: row.date,
        pilot: row.pilot,
        aircraft: row.aircraft,
        from: row.from,
        to: row.to,
        type: row.type,
        totalTime: row.totalTime,
        landings: row.landings
      };
    });
    var pilotSummaries = Object.keys(pilotAgg).map(function(name) { return pilotAgg[name]; })
      .sort(function(a, b) { return b.totalTime - a.totalTime || String(a.pilot).localeCompare(String(b.pilot)); });
    var monthlyRows = Object.keys(monthlyAgg).map(function(key) { return monthlyAgg[key]; })
      .sort(function(a, b) { return String(b.month).localeCompare(String(a.month)) || String(a.type).localeCompare(String(b.type)); });
    var monthlySeen = {};
    monthlyRows.forEach(function(row) { if (row && row.month) monthlySeen[row.month] = true; });

    return {
      success: true,
      filterLabel: startDate || endDate
        ? ((startDate ? isoDate(startDate) : '...') + ' to ' + (endDate ? isoDate(endDate) : '...'))
        : 'all dates',
      summary: {
        totalFundUsage: totalFundUsage,
        totalFlightHours: flights.reduce(function(sum, row) { return sum + num(row.totalTime); }, 0),
        totalFlights: flights.length,
        totalPilots: pilotSummaries.length
      },
      fundUsage: {
        count: fundRows.length,
        totalAmount: totalFundUsage,
        rows: fundRows
      },
      flights: {
        count: flights.length,
        totalHours: flights.reduce(function(sum, row) { return sum + num(row.totalTime); }, 0),
        rows: flights
      },
      pilotLogbook: {
        pilotCount: pilotSummaries.length,
        entryCount: pilotRows.length,
        summaries: pilotSummaries,
        rows: pilotRows
      },
      monthlyTotals: {
        monthCount: Object.keys(monthlySeen).length,
        count: monthlyRows.length,
        rows: monthlyRows
      }
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _schemaNormHeader_(value) {
  return String(value || '').trim().toUpperCase().replace(/\s+/g, '_').replace(/[^A-Z0-9_]/g, '');
}

function _schemaEnsureColumns_(sheet, headersToEnsure) {
  var headerRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  var existingNorms = headerRow.map(function(h) { return _schemaNormHeader_(h); });
  var added = [];
  (headersToEnsure || []).forEach(function(header) {
    var label = String(header || '').trim();
    if (!label) return;
    var norm = _schemaNormHeader_(label);
    if (existingNorms.indexOf(norm) >= 0) return;
    var col = sheet.getLastColumn() + 1;
    sheet.getRange(1, col).setValue(label);
    existingNorms.push(norm);
    added.push(label);
  });
  return added;
}

function _schemaEnsureSheetHeaders_(ss, sheetName, headers) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    if (headers && headers.length) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return { sheet: sh, created: true, added: (headers || []).slice() };
  }
  var added = _schemaEnsureColumns_(sh, headers || []);
  return { sheet: sh, created: false, added: added };
}

function _schemaSeedRowsByKey_(sheet, keyHeader, rows) {
  var out = { inserted: 0 };
  var data = sheet.getDataRange().getValues();
  if (!data || data.length < 1) return out;
  var headers = data[0].map(function(h) { return String(h || '').trim(); });
  var norms = headers.map(function(h) { return _schemaNormHeader_(h); });
  var keyIdx = norms.indexOf(_schemaNormHeader_(keyHeader));
  if (keyIdx < 0) return out;
  var existing = {};
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][keyIdx] || '').trim().toUpperCase();
    if (key) existing[key] = true;
  }
  var append = [];
  (rows || []).forEach(function(rowObj) {
    var key = String((rowObj && rowObj[keyHeader]) || '').trim();
    var normKey = key.toUpperCase();
    if (!key || existing[normKey]) return;
    var row = headers.map(function(h) { return Object.prototype.hasOwnProperty.call(rowObj || {}, h) ? rowObj[h] : ''; });
    append.push(row);
    existing[normKey] = true;
  });
  if (append.length) {
    var start = sheet.getLastRow() + 1;
    sheet.getRange(start, 1, append.length, headers.length).setValues(append);
    out.inserted = append.length;
  }
  return out;
}

function setupStaffSchema() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var staffColumns = [
      'STAFF_ID',
      'EMAIL',
      'ACTIVE',
      'PRIMARY_ROLE',
      'STAFF_ROLES_JSON',
      'CAN_EDIT_DISCREPANCIES',
      'CAN_APPROVE_DEFERMENTS',
      'CAN_EDIT_MAINTENANCE',
      'CAN_FLIGHT_FOLLOW',
      'CAN_COORDINATE_FLIGHTS',
      'CAN_MANAGE_STOCKROOM',
      'CAN_INSTRUCT',
      'CAN_INSPECT',
      'COMPLETED_COURSES_JSON',
      'QUALIFICATIONS_JSON',
      'LAST_TRAINING_SYNC_AT',
      'NOTES'
    ];

    var staffRoleHeaders = [
      'ROLE_CODE',
      'ROLE_NAME',
      'ROLE_GROUP',
      'ACTIVE',
      'DESCRIPTION'
    ];
    var moduleHeaders = [
      'MODULE_ID',
      'MODULE_NAME',
      'ROLE_CODE',
      'COMPONENT',
      'MODULE_TYPE',
      'PASS_THRESHOLD',
      'REQUIRES_PRACTICAL',
      'RECURRENT_DAYS',
      'CLASSROOM_COURSE_ID',
      'CLASSROOM_COURSEWORK_ID',
      'ACTIVE',
      'NOTES'
    ];
    var trainingHeaders = [
      'RECORD_ID',
      'STAFF_ID',
      'STAFF_EMAIL',
      'MODULE_ID',
      'MODULE_NAME',
      'ROLE_CODE',
      'SOURCE',
      'THEORY_SCORE',
      'THEORY_MAX_SCORE',
      'THEORY_PASSED',
      'THEORY_COMPLETED_AT',
      'PRACTICAL_PASSED',
      'PRACTICAL_EVALUATOR',
      'PRACTICAL_COMPLETED_AT',
      'VALID_UNTIL',
      'EXTERNAL_SUBMISSION_ID',
      'EVIDENCE_URL',
      'RECORDED_BY',
      'RECORDED_AT',
      'NOTES'
    ];
    var practicalHeaders = [
      'PRACTICAL_ID',
      'STAFF_ID',
      'STAFF_EMAIL',
      'MODULE_ID',
      'AIRCRAFT',
      'EVALUATOR',
      'RESULT',
      'EVALUATED_AT',
      'TACH_AT_EVAL',
      'LOCATION',
      'OBSERVATIONS',
      'RECORDED_BY',
      'RECORDED_AT'
    ];

    var roles = [
      { ROLE_CODE: 'OP_PILOT_LAND', ROLE_NAME: 'Operational Pilot Land', ROLE_GROUP: 'Operations', ACTIVE: 'Y', DESCRIPTION: 'Operational pilot for land aircraft' },
      { ROLE_CODE: 'OP_INSTR_PILOT_LAND', ROLE_NAME: 'Operational Instructor Pilot Land', ROLE_GROUP: 'Operations', ACTIVE: 'Y', DESCRIPTION: 'Instructor pilot for land operations' },
      { ROLE_CODE: 'OP_PILOT_ANF', ROLE_NAME: 'Operational Pilot ANF', ROLE_GROUP: 'Operations', ACTIVE: 'Y', DESCRIPTION: 'Operational pilot for ANF operations' },
      { ROLE_CODE: 'OP_INSTR_PILOT_ANF', ROLE_NAME: 'Operational Instructor Pilot ANF', ROLE_GROUP: 'Operations', ACTIVE: 'Y', DESCRIPTION: 'Instructor pilot for ANF operations' },
      { ROLE_CODE: 'FLIGHT_INSTRUCTOR', ROLE_NAME: 'Flight Instructor', ROLE_GROUP: 'Training', ACTIVE: 'Y', DESCRIPTION: 'Flight instruction role' },
      { ROLE_CODE: 'MECHANIC_TRAINEE', ROLE_NAME: 'Mechanic In Training', ROLE_GROUP: 'Maintenance', ACTIVE: 'Y', DESCRIPTION: 'Mechanic trainee role' },
      { ROLE_CODE: 'MECHANIC', ROLE_NAME: 'Mechanic', ROLE_GROUP: 'Maintenance', ACTIVE: 'Y', DESCRIPTION: 'Certified mechanic role' },
      { ROLE_CODE: 'INSPECTOR', ROLE_NAME: 'Inspector', ROLE_GROUP: 'Maintenance', ACTIVE: 'Y', DESCRIPTION: 'Inspection authority role' },
      { ROLE_CODE: 'FLIGHT_FOLLOWER', ROLE_NAME: 'Flight Follower', ROLE_GROUP: 'Operations', ACTIVE: 'Y', DESCRIPTION: 'Flight following role' },
      { ROLE_CODE: 'FLIGHT_COORDINATOR', ROLE_NAME: 'Flight Coordinator', ROLE_GROUP: 'Operations', ACTIVE: 'Y', DESCRIPTION: 'Flight coordination role' },
      { ROLE_CODE: 'SRM', ROLE_NAME: 'SRM', ROLE_GROUP: 'Safety', ACTIVE: 'Y', DESCRIPTION: 'Safety/risk management role' },
      { ROLE_CODE: 'STOCKROOM', ROLE_NAME: 'Stockroom', ROLE_GROUP: 'Logistics', ACTIVE: 'Y', DESCRIPTION: 'Stockroom and inventory role' }
    ];

    var pilotsSheet = getRequiredSheet_(ss, APP_SHEETS.PILOTS, 'setupStaffSchema');
    var addedStaffColumns = _schemaEnsureColumns_(pilotsSheet, staffColumns);

    var roleSheetRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.STAFF_ROLES || 'REF_Staff_Roles', staffRoleHeaders);
    var moduleSheetRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.TRAINING_MODULES || 'REF_Training_Modules', moduleHeaders);
    var trainingSheetRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.STAFF_TRAINING || 'DB_Staff_Training_Records', trainingHeaders);
    var practicalSheetRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.STAFF_PRACTICALS || 'DB_Staff_Practical_Evaluations', practicalHeaders);
    var seededRoles = _schemaSeedRowsByKey_(roleSheetRes.sheet, 'ROLE_CODE', roles);

    return {
      success: true,
      pilotsSheet: APP_SHEETS.PILOTS,
      addedStaffColumns: addedStaffColumns,
      roleSheet: {
        name: roleSheetRes.sheet.getName(),
        created: roleSheetRes.created,
        addedColumns: roleSheetRes.added,
        seededRoles: seededRoles.inserted
      },
      moduleSheet: {
        name: moduleSheetRes.sheet.getName(),
        created: moduleSheetRes.created,
        addedColumns: moduleSheetRes.added
      },
      trainingSheet: {
        name: trainingSheetRes.sheet.getName(),
        created: trainingSheetRes.created,
        addedColumns: trainingSheetRes.added
      },
      practicalSheet: {
        name: practicalSheetRes.sheet.getName(),
        created: practicalSheetRes.created,
        addedColumns: practicalSheetRes.added
      }
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _schedulerHeaderIndex_(headers, candidate) {
  var norms = (headers || []).map(function(h) { return _schemaNormHeader_(h); });
  return norms.indexOf(_schemaNormHeader_(candidate));
}

function _schedulerTruthyFlag_(value) {
  var raw = String(value == null ? '' : value).trim().toUpperCase();
  return raw === 'Y' || raw === 'YES' || raw === 'TRUE' || raw === '1' || raw === 'SIM' || raw === 'ATIVO';
}

function _schedulerCurrentUserEmail_() {
  try {
    return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  } catch (e) {
    return '';
  }
}

function _schedulerPermissionRowByEmail_(sheet, email) {
  var target = String(email || '').trim().toLowerCase();
  if (!target) return null;
  var data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return null;
  var headers = data[0];
  var emailIdx = _schedulerHeaderIndex_(headers, 'EMAIL');
  if (emailIdx < 0) return null;
  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][emailIdx] || '').trim().toLowerCase();
    if (rowEmail && rowEmail === target) {
      return { rowNumber: i + 1, headers: headers, row: data[i] };
    }
  }
  return null;
}

function _schedulerAssertPermission_(permissionKey, contextLabel) {
  var email = _schedulerCurrentUserEmail_();
  if (!email) throw new Error((contextLabel || 'scheduler') + ': signed-in email is required');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var permSheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_PERMISSIONS || 'SCHED_Permissions', contextLabel || 'schedulerPermissionCheck');
  var rec = _schedulerPermissionRowByEmail_(permSheet, email);
  if (!rec) throw new Error((contextLabel || 'scheduler') + ': no scheduler permissions found for ' + email);

  var activeIdx = _schedulerHeaderIndex_(rec.headers, 'ACTIVE');
  if (activeIdx >= 0 && !_schedulerTruthyFlag_(rec.row[activeIdx])) {
    throw new Error((contextLabel || 'scheduler') + ': scheduler access is inactive for ' + email);
  }

  if (permissionKey) {
    var idx = _schedulerHeaderIndex_(rec.headers, permissionKey);
    if (idx < 0) throw new Error((contextLabel || 'scheduler') + ': permission column missing: ' + permissionKey);
    if (!_schedulerTruthyFlag_(rec.row[idx])) {
      throw new Error((contextLabel || 'scheduler') + ': permission denied for ' + email + ' (' + permissionKey + ')');
    }
  }

  return {
    email: email,
    rowNumber: rec.rowNumber,
    headers: rec.headers,
    row: rec.row
  };
}

function _schedulerReadConfigMap_(configSheet) {
  var out = {};
  var data = configSheet.getDataRange().getValues();
  if (!data || data.length < 2) return out;
  var headers = data[0];
  var keyIdx = _schedulerHeaderIndex_(headers, 'CONFIG_KEY');
  var valueIdx = _schedulerHeaderIndex_(headers, 'CONFIG_VALUE');
  var activeIdx = _schedulerHeaderIndex_(headers, 'ACTIVE');
  if (keyIdx < 0 || valueIdx < 0) return out;
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][keyIdx] || '').trim();
    if (!key) continue;
    if (activeIdx >= 0 && !_schedulerTruthyFlag_(data[i][activeIdx])) continue;
    out[key] = String(data[i][valueIdx] == null ? '' : data[i][valueIdx]);
  }
  return out;
}

function setupSchedulerSchema() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var now = new Date();

    var configHeaders = [
      'CONFIG_KEY',
      'CONFIG_VALUE',
      'DESCRIPTION',
      'ACTIVE',
      'UPDATED_AT',
      'UPDATED_BY'
    ];
    var permHeaders = [
      'EMAIL',
      'CAN_VIEW',
      'CAN_GENERATE',
      'CAN_EDIT_ASSIGNMENTS',
      'CAN_LOCK_ASSIGNMENTS',
      'CAN_PUBLISH',
      'CAN_EDIT_RULES',
      'CAN_MANAGE_PERMISSIONS',
      'ACTIVE',
      'NOTES',
      'UPDATED_AT',
      'UPDATED_BY'
    ];
    var coverageHeaders = [
      'REQUIREMENT_ID',
      'ROLE_CODE',
      'LOCATION_CODE',
      'DAYS_MASK',
      'SHIFT_START_LOCAL',
      'SHIFT_END_LOCAL',
      'SHIFT_LABEL',
      'REQUIRED_COUNT',
      'PRIORITY',
      'ALLOW_MULTI_ROLE_OVERLAP',
      'ACTIVE',
      'NOTES'
    ];
    var compatHeaders = [
      'RULE_ID',
      'ROLE_CODE_A',
      'ROLE_CODE_B',
      'OVERLAP_ALLOWED',
      'ACTIVE',
      'NOTES'
    ];
    var qualHeaders = [
      'QUAL_ID',
      'STAFF_EMAIL',
      'ROLE_CODE',
      'LOCATION_CODE',
      'VALID_FROM',
      'VALID_UNTIL',
      'ACTIVE',
      'SOURCE',
      'NOTES',
      'UPDATED_AT',
      'UPDATED_BY'
    ];
    var assignHeaders = [
      'ASSIGNMENT_ID',
      'SCHEDULE_MONTH',
      'ASSIGNMENT_DATE',
      'ROLE_CODE',
      'LOCATION_CODE',
      'SHIFT_START_LOCAL',
      'SHIFT_END_LOCAL',
      'SHIFT_START_Z',
      'SHIFT_END_Z',
      'STAFF_EMAIL',
      'SOURCE',
      'LOCKED',
      'LOCK_SCOPE',
      'STATUS',
      'FAIRNESS_SCORE',
      'RULE_REASON',
      'UPDATED_AT',
      'UPDATED_BY'
    ];
    var lockHeaders = [
      'LOCK_ID',
      'ASSIGNMENT_ID',
      'SCHEDULE_MONTH',
      'LOCK_SCOPE',
      'LOCK_REASON',
      'LOCKED_BY',
      'LOCKED_AT',
      'ACTIVE'
    ];
    var alertsHeaders = [
      'ALERT_ID',
      'ALERT_TYPE',
      'SCHEDULE_MONTH',
      'ASSIGNMENT_DATE',
      'ROLE_CODE',
      'LOCATION_CODE',
      'SEVERITY',
      'STATUS',
      'MESSAGE',
      'RECIPIENTS',
      'SENT_AT',
      'CREATED_AT'
    ];
    var publishHeaders = [
      'PUBLISH_ID',
      'SCHEDULE_MONTH',
      'PUBLISHED_BY',
      'PUBLISHED_AT',
      'RESULT',
      'UNFILLED_COUNT',
      'NOTES'
    ];
    var availabilityHeaders = [
      'AVAILABILITY_ID',
      'SCHEDULE_MONTH',
      'STAFF_EMAIL',
      'EVENT_ID',
      'DATE_LOCAL',
      'START_LOCAL',
      'END_LOCAL',
      'START_Z',
      'END_Z',
      'BASE_CODE',
      'TYPE',
      'SOURCE',
      'UPDATED_AT'
    ];

    var configSheetRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.SCHED_CONFIG || 'SCHED_Config', configHeaders);
    var permSheetRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.SCHED_PERMISSIONS || 'SCHED_Permissions', permHeaders);
    var coverageSheetRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.SCHED_COVERAGE_RULES || 'SCHED_Coverage_Requirements', coverageHeaders);
    var compatSheetRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.SCHED_ROLE_COMPAT || 'SCHED_Role_Compatibility', compatHeaders);
    var qualSheetRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.SCHED_STAFF_QUALS || 'SCHED_Staff_Qualifications', qualHeaders);
    var assignSheetRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.SCHED_ASSIGNMENTS || 'SCHED_Assignments', assignHeaders);
    var lockSheetRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.SCHED_LOCKS || 'SCHED_Assignment_Locks', lockHeaders);
    var alertsSheetRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.SCHED_ALERTS || 'SCHED_Alerts', alertsHeaders);
    var publishSheetRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.SCHED_PUBLISH_LOG || 'SCHED_Publish_Log', publishHeaders);
    var availabilitySheetRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.SCHED_AVAILABILITY || 'SCHED_Staff_Availability', availabilityHeaders);

    var roleSheetRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.STAFF_ROLES || 'REF_Staff_Roles', ['ROLE_CODE', 'ROLE_NAME', 'ROLE_GROUP', 'ACTIVE', 'DESCRIPTION']);
    var schedulerRoles = [
      { ROLE_CODE: 'FLIGHT_SUPERVISOR', ROLE_NAME: 'Flight Supervisor', ROLE_GROUP: 'Operations', ACTIVE: 'Y', DESCRIPTION: 'Daily operational flight supervision coverage' },
      { ROLE_CODE: 'MAINTENANCE_SUPERVISOR', ROLE_NAME: 'Maintenance Supervisor', ROLE_GROUP: 'Maintenance', ACTIVE: 'Y', DESCRIPTION: 'Daily maintenance supervision coverage' }
    ];
    var seededSchedulerRoles = _schemaSeedRowsByKey_(roleSheetRes.sheet, 'ROLE_CODE', schedulerRoles);

    var configSeed = [
      { CONFIG_KEY: 'AVAILABILITY_CALENDAR_ID', CONFIG_VALUE: '', DESCRIPTION: 'Shared calendar id with unavailable-only events', ACTIVE: 'Y' },
      { CONFIG_KEY: 'FLIGHTS_CALENDAR_ID', CONFIG_VALUE: '', DESCRIPTION: 'Target Google Calendar id for flight events', ACTIVE: 'Y' },
      { CONFIG_KEY: 'SCHEDULE_CALENDAR_ID', CONFIG_VALUE: '', DESCRIPTION: 'Target Google Calendar id for staffing schedule', ACTIVE: 'Y' },
      { CONFIG_KEY: 'ALERT_RECIPIENTS', CONFIG_VALUE: '', DESCRIPTION: 'Comma-separated email recipients for scheduler alerts', ACTIVE: 'Y' },
      { CONFIG_KEY: 'PUBLISH_DAY_OF_MONTH', CONFIG_VALUE: '15', DESCRIPTION: 'Auto-publish day for next month schedule', ACTIVE: 'Y' },
      { CONFIG_KEY: 'BASE_TZ_BVB', CONFIG_VALUE: 'America/Manaus', DESCRIPTION: 'Local timezone for BVB scheduling', ACTIVE: 'Y' },
      { CONFIG_KEY: 'BASE_TZ_PVH', CONFIG_VALUE: 'America/Porto_Velho', DESCRIPTION: 'Local timezone for PVH scheduling', ACTIVE: 'Y' },
      { CONFIG_KEY: 'BASE_TZ_APS', CONFIG_VALUE: 'America/Sao_Paulo', DESCRIPTION: 'Local timezone for APS scheduling', ACTIVE: 'Y' }
    ];
    var seededConfig = _schemaSeedRowsByKey_(configSheetRes.sheet, 'CONFIG_KEY', configSeed);

    var compatSeed = [
      { RULE_ID: 'COMPAT-FSUP-MSUP', ROLE_CODE_A: 'FLIGHT_SUPERVISOR', ROLE_CODE_B: 'MAINTENANCE_SUPERVISOR', OVERLAP_ALLOWED: 'Y', ACTIVE: 'Y', NOTES: 'Can overlap if qualified' },
      { RULE_ID: 'COMPAT-MSUP-FF', ROLE_CODE_A: 'MAINTENANCE_SUPERVISOR', ROLE_CODE_B: 'FLIGHT_FOLLOWER', OVERLAP_ALLOWED: 'Y', ACTIVE: 'Y', NOTES: 'Can overlap if qualified' },
      { RULE_ID: 'COMPAT-MECH-BLOCK', ROLE_CODE_A: 'MECHANIC', ROLE_CODE_B: '*', OVERLAP_ALLOWED: 'N', ACTIVE: 'Y', NOTES: 'Mechanic assignment blocks overlapping operational roles' },
      { RULE_ID: 'COMPAT-INSP-BLOCK', ROLE_CODE_A: 'INSPECTOR', ROLE_CODE_B: '*', OVERLAP_ALLOWED: 'N', ACTIVE: 'Y', NOTES: 'Inspector assignment blocks overlapping operational roles' }
    ];
    var seededCompat = _schemaSeedRowsByKey_(compatSheetRes.sheet, 'RULE_ID', compatSeed);

    var actorEmail = _schedulerCurrentUserEmail_();
    var seededAdmin = false;
    if (actorEmail) {
      var existingAdmin = _schedulerPermissionRowByEmail_(permSheetRes.sheet, actorEmail);
      if (!existingAdmin) {
        permSheetRes.sheet.appendRow([
          actorEmail,
          'Y',
          'Y',
          'Y',
          'Y',
          'Y',
          'Y',
          'Y',
          'Y',
          'Seeded by setupSchedulerSchema',
          now,
          actorEmail
        ]);
        seededAdmin = true;
      }
    }

    return {
      success: true,
      seededAdmin: seededAdmin,
      currentUser: actorEmail,
      seededRoles: seededSchedulerRoles.inserted,
      seededConfig: seededConfig.inserted,
      seededCompatibilityRules: seededCompat.inserted,
      sheets: {
        config: { name: configSheetRes.sheet.getName(), created: configSheetRes.created, addedColumns: configSheetRes.added },
        permissions: { name: permSheetRes.sheet.getName(), created: permSheetRes.created, addedColumns: permSheetRes.added },
        coverage: { name: coverageSheetRes.sheet.getName(), created: coverageSheetRes.created, addedColumns: coverageSheetRes.added },
        compatibility: { name: compatSheetRes.sheet.getName(), created: compatSheetRes.created, addedColumns: compatSheetRes.added },
        qualifications: { name: qualSheetRes.sheet.getName(), created: qualSheetRes.created, addedColumns: qualSheetRes.added },
        assignments: { name: assignSheetRes.sheet.getName(), created: assignSheetRes.created, addedColumns: assignSheetRes.added },
        locks: { name: lockSheetRes.sheet.getName(), created: lockSheetRes.created, addedColumns: lockSheetRes.added },
        alerts: { name: alertsSheetRes.sheet.getName(), created: alertsSheetRes.created, addedColumns: alertsSheetRes.added },
        publishLog: { name: publishSheetRes.sheet.getName(), created: publishSheetRes.created, addedColumns: publishSheetRes.added },
        availability: { name: availabilitySheetRes.sheet.getName(), created: availabilitySheetRes.created, addedColumns: availabilitySheetRes.added }
      }
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getSchedulerAccessProfile() {
  try {
    var profile = _schedulerAssertPermission_(null, 'getSchedulerAccessProfile');
    var headers = profile.headers || [];
    var row = profile.row || [];
    var out = {
      success: true,
      email: profile.email,
      permissions: {}
    };
    ['CAN_VIEW', 'CAN_GENERATE', 'CAN_EDIT_ASSIGNMENTS', 'CAN_LOCK_ASSIGNMENTS', 'CAN_PUBLISH', 'CAN_EDIT_RULES', 'CAN_MANAGE_PERMISSIONS', 'ACTIVE'].forEach(function(key) {
      var idx = _schedulerHeaderIndex_(headers, key);
      out.permissions[key] = idx >= 0 ? _schedulerTruthyFlag_(row[idx]) : false;
    });
    return out;
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getSchedulerConfig() {
  try {
    _schedulerAssertPermission_('CAN_VIEW', 'getSchedulerConfig');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_CONFIG || 'SCHED_Config', 'getSchedulerConfig');
    return {
      success: true,
      config: _schedulerReadConfigMap_(sheet)
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function saveSchedulerConfigEntries(payload) {
  try {
    var actor = _schedulerAssertPermission_('CAN_EDIT_RULES', 'saveSchedulerConfigEntries');
    var body = (payload && typeof payload === 'object') ? payload : {};
    var entries = [];
    if (Array.isArray(body.entries)) {
      entries = body.entries;
    } else if (body.key) {
      entries = [{ key: body.key, value: body.value, description: body.description, active: body.active }];
    }
    if (!entries.length) return { success: false, error: 'No config entries provided' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_CONFIG || 'SCHED_Config', 'saveSchedulerConfigEntries');
    var data = sheet.getDataRange().getValues();
    var headers = data.length ? data[0] : [];
    var keyIdx = _schedulerHeaderIndex_(headers, 'CONFIG_KEY');
    var valueIdx = _schedulerHeaderIndex_(headers, 'CONFIG_VALUE');
    var descIdx = _schedulerHeaderIndex_(headers, 'DESCRIPTION');
    var activeIdx = _schedulerHeaderIndex_(headers, 'ACTIVE');
    var updAtIdx = _schedulerHeaderIndex_(headers, 'UPDATED_AT');
    var updByIdx = _schedulerHeaderIndex_(headers, 'UPDATED_BY');
    if (keyIdx < 0 || valueIdx < 0) return { success: false, error: 'SCHED_Config is missing required columns' };

    var updated = 0;
    var created = 0;
    var now = new Date();
    entries.forEach(function(entry) {
      var key = String((entry && (entry.key || entry.CONFIG_KEY)) || '').trim();
      if (!key) return;
      var value = entry && Object.prototype.hasOwnProperty.call(entry, 'value') ? entry.value : entry && entry.CONFIG_VALUE;
      var description = entry && Object.prototype.hasOwnProperty.call(entry, 'description') ? entry.description : entry && entry.DESCRIPTION;
      var activeRaw = entry && Object.prototype.hasOwnProperty.call(entry, 'active') ? entry.active : entry && entry.ACTIVE;
      var activeVal = activeRaw == null ? null : (_schedulerTruthyFlag_(activeRaw) ? 'Y' : 'N');

      var foundRow = 0;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][keyIdx] || '').trim() === key) {
          foundRow = i + 1;
          break;
        }
      }

      if (foundRow) {
        sheet.getRange(foundRow, valueIdx + 1).setValue(value == null ? '' : value);
        if (descIdx >= 0 && description != null) sheet.getRange(foundRow, descIdx + 1).setValue(description);
        if (activeIdx >= 0 && activeVal != null) sheet.getRange(foundRow, activeIdx + 1).setValue(activeVal);
        if (updAtIdx >= 0) sheet.getRange(foundRow, updAtIdx + 1).setValue(now);
        if (updByIdx >= 0) sheet.getRange(foundRow, updByIdx + 1).setValue(actor.email);
        updated++;
      } else {
        var row = headers.map(function() { return ''; });
        row[keyIdx] = key;
        row[valueIdx] = value == null ? '' : value;
        if (descIdx >= 0) row[descIdx] = description == null ? '' : description;
        if (activeIdx >= 0) row[activeIdx] = activeVal == null ? 'Y' : activeVal;
        if (updAtIdx >= 0) row[updAtIdx] = now;
        if (updByIdx >= 0) row[updByIdx] = actor.email;
        sheet.appendRow(row);
        created++;
      }
    });

    return { success: true, created: created, updated: updated };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getSchedulerPermissions() {
  try {
    _schedulerAssertPermission_('CAN_MANAGE_PERMISSIONS', 'getSchedulerPermissions');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_PERMISSIONS || 'SCHED_Permissions', 'getSchedulerPermissions');
    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return { success: true, rows: [] };
    var headers = data[0];
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var rowObj = { rowNumber: i + 1 };
      for (var c = 0; c < headers.length; c++) {
        rowObj[String(headers[c] || '').trim()] = data[i][c];
      }
      rows.push(rowObj);
    }
    return { success: true, rows: rows };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function saveSchedulerPermission(payload) {
  try {
    var actor = _schedulerAssertPermission_('CAN_MANAGE_PERMISSIONS', 'saveSchedulerPermission');
    var body = (payload && typeof payload === 'object') ? payload : {};
    var targetEmail = String(body.email || body.EMAIL || '').trim().toLowerCase();
    if (!targetEmail) return { success: false, error: 'Email is required' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_PERMISSIONS || 'SCHED_Permissions', 'saveSchedulerPermission');
    var data = sheet.getDataRange().getValues();
    var headers = data.length ? data[0] : [];
    var now = new Date();

    var rec = _schedulerPermissionRowByEmail_(sheet, targetEmail);
    var rowNumber = rec ? rec.rowNumber : 0;
    if (!rowNumber) {
      var row = headers.map(function() { return ''; });
      row[_schedulerHeaderIndex_(headers, 'EMAIL')] = targetEmail;
      rowNumber = sheet.getLastRow() + 1;
      sheet.appendRow(row);
    }

    var permissionKeys = [
      'CAN_VIEW',
      'CAN_GENERATE',
      'CAN_EDIT_ASSIGNMENTS',
      'CAN_LOCK_ASSIGNMENTS',
      'CAN_PUBLISH',
      'CAN_EDIT_RULES',
      'CAN_MANAGE_PERMISSIONS',
      'ACTIVE'
    ];
    permissionKeys.forEach(function(key) {
      var idx = _schedulerHeaderIndex_(headers, key);
      if (idx < 0) return;
      if (!Object.prototype.hasOwnProperty.call(body, key) && !Object.prototype.hasOwnProperty.call(body, key.toLowerCase())) return;
      var val = Object.prototype.hasOwnProperty.call(body, key) ? body[key] : body[key.toLowerCase()];
      sheet.getRange(rowNumber, idx + 1).setValue(_schedulerTruthyFlag_(val) ? 'Y' : 'N');
    });

    var notesIdx = _schedulerHeaderIndex_(headers, 'NOTES');
    if (notesIdx >= 0 && (Object.prototype.hasOwnProperty.call(body, 'notes') || Object.prototype.hasOwnProperty.call(body, 'NOTES'))) {
      sheet.getRange(rowNumber, notesIdx + 1).setValue(Object.prototype.hasOwnProperty.call(body, 'notes') ? body.notes : body.NOTES);
    }
    var updatedAtIdx = _schedulerHeaderIndex_(headers, 'UPDATED_AT');
    var updatedByIdx = _schedulerHeaderIndex_(headers, 'UPDATED_BY');
    if (updatedAtIdx >= 0) sheet.getRange(rowNumber, updatedAtIdx + 1).setValue(now);
    if (updatedByIdx >= 0) sheet.getRange(rowNumber, updatedByIdx + 1).setValue(actor.email);

    return { success: true, rowNumber: rowNumber, email: targetEmail };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _schedulerMonthKey_(dateObj) {
  if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) return '';
  return Utilities.formatDate(dateObj, 'Etc/UTC', 'yyyy-MM');
}

function _schedulerParseMonthKey_(monthKey) {
  var raw = String(monthKey || '').trim();
  var m = raw.match(/^(\d{4})-(\d{2})$/);
  if (!m) return null;
  var year = Number(m[1]);
  var month = Number(m[2]);
  if (month < 1 || month > 12) return null;
  return {
    monthKey: raw,
    start: new Date(year, month - 1, 1, 0, 0, 0, 0),
    end: new Date(year, month, 1, 0, 0, 0, 0)
  };
}

function _schedulerExtractEmailFromText_(text) {
  var raw = String(text || '');
  var hit = raw.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return hit ? String(hit[0] || '').trim().toLowerCase() : '';
}

function _schedulerExtractBaseFromText_(text) {
  var raw = String(text || '');
  var tagged = raw.match(/(?:base|location)\s*[:=]\s*([A-Z0-9_-]{2,12})/i);
  if (tagged) return String(tagged[1] || '').trim().toUpperCase();
  return '';
}

function _schedulerEventLooksUnavailable_(eventTitle) {
  var title = String(eventTitle || '').trim().toUpperCase();
  if (!title) return true;
  if (title.indexOf('AVAILABLE') === 0) return false;
  if (title.indexOf('DISPONIVEL') === 0) return false;
  return true;
}

function syncSchedulerAvailabilityForMonth(payload) {
  try {
    var actor = _schedulerAssertPermission_('CAN_GENERATE', 'syncSchedulerAvailabilityForMonth');
    var body = (payload && typeof payload === 'object') ? payload : {};
    var requestedMonth = String(body.month || body.scheduleMonth || '').trim();
    var parsedMonth = _schedulerParseMonthKey_(requestedMonth || _schedulerMonthKey_(new Date()));
    if (!parsedMonth) return { success: false, error: 'Invalid month format. Use YYYY-MM.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_CONFIG || 'SCHED_Config', 'syncSchedulerAvailabilityForMonth');
    var cfg = _schedulerReadConfigMap_(configSheet);
    var calendarId = String(cfg.AVAILABILITY_CALENDAR_ID || '').trim();
    if (!calendarId) {
      return { success: false, error: 'Missing SCHED_Config key AVAILABILITY_CALENDAR_ID' };
    }

    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) return { success: false, error: 'Calendar not found for AVAILABILITY_CALENDAR_ID' };

    var availabilitySheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_AVAILABILITY || 'SCHED_Staff_Availability', 'syncSchedulerAvailabilityForMonth');
    var allData = availabilitySheet.getDataRange().getValues();
    var headers = allData.length ? allData[0] : [];
    if (!headers.length) return { success: false, error: 'SCHED_Staff_Availability is missing headers' };

    var idx = {
      month: _schedulerHeaderIndex_(headers, 'SCHEDULE_MONTH'),
      source: _schedulerHeaderIndex_(headers, 'SOURCE'),
      id: _schedulerHeaderIndex_(headers, 'AVAILABILITY_ID'),
      staffEmail: _schedulerHeaderIndex_(headers, 'STAFF_EMAIL'),
      eventId: _schedulerHeaderIndex_(headers, 'EVENT_ID'),
      dateLocal: _schedulerHeaderIndex_(headers, 'DATE_LOCAL'),
      startLocal: _schedulerHeaderIndex_(headers, 'START_LOCAL'),
      endLocal: _schedulerHeaderIndex_(headers, 'END_LOCAL'),
      startZ: _schedulerHeaderIndex_(headers, 'START_Z'),
      endZ: _schedulerHeaderIndex_(headers, 'END_Z'),
      baseCode: _schedulerHeaderIndex_(headers, 'BASE_CODE'),
      type: _schedulerHeaderIndex_(headers, 'TYPE'),
      updatedAt: _schedulerHeaderIndex_(headers, 'UPDATED_AT')
    };

    var keepRows = [headers];
    for (var r = 1; r < allData.length; r++) {
      var row = allData[r];
      var rowMonth = idx.month >= 0 ? String(row[idx.month] || '').trim() : '';
      var rowSource = idx.source >= 0 ? String(row[idx.source] || '').trim().toUpperCase() : '';
      var isTargetMonth = rowMonth === parsedMonth.monthKey;
      var isCalendarSyncRow = rowSource === 'CALENDAR_UNAVAILABLE_SYNC';
      if (isTargetMonth && isCalendarSyncRow) continue;
      keepRows.push(row);
    }

    rewriteSheetData_(availabilitySheet, keepRows);

    var events = calendar.getEvents(parsedMonth.start, parsedMonth.end);
    var appendRows = [];
    var skippedNoEmail = 0;
    var skippedAvailable = 0;
    var imported = 0;

    events.forEach(function(ev) {
      var title = String(ev.getTitle() || '').trim();
      if (!_schedulerEventLooksUnavailable_(title)) {
        skippedAvailable++;
        return;
      }

      var desc = '';
      try { desc = String(ev.getDescription() || ''); } catch (e) { desc = ''; }

      var staffEmail = _schedulerExtractEmailFromText_(title + '\n' + desc);
      if (!staffEmail) {
        try {
          var creators = ev.getCreators();
          if (creators && creators.length) staffEmail = String(creators[0] || '').trim().toLowerCase();
        } catch (e2) {}
      }
      if (!staffEmail) {
        skippedNoEmail++;
        return;
      }

      var start = ev.getStartTime();
      var end = ev.getEndTime();
      var baseCode = _schedulerExtractBaseFromText_(title + '\n' + desc) || String(body.baseCode || 'GLOBAL').trim().toUpperCase();
      var dateLocal = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd');

      var startLocal;
      var endLocal;
      if (ev.isAllDayEvent()) {
        startLocal = '00:00';
        endLocal = '23:59';
      } else {
        startLocal = Utilities.formatDate(start, Session.getScriptTimeZone(), 'HH:mm');
        endLocal = Utilities.formatDate(end, Session.getScriptTimeZone(), 'HH:mm');
      }

      var startZulu = Utilities.formatDate(start, 'Etc/UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
      var endZulu = Utilities.formatDate(end, 'Etc/UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
      var availabilityId = 'AVL-' + parsedMonth.monthKey + '-' + (ev.getId() || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 24) + '-' + staffEmail.replace(/[^a-z0-9]/g, '').slice(0, 16);

      var row = headers.map(function() { return ''; });
      if (idx.id >= 0) row[idx.id] = availabilityId;
      if (idx.month >= 0) row[idx.month] = parsedMonth.monthKey;
      if (idx.staffEmail >= 0) row[idx.staffEmail] = staffEmail;
      if (idx.eventId >= 0) row[idx.eventId] = ev.getId();
      if (idx.dateLocal >= 0) row[idx.dateLocal] = dateLocal;
      if (idx.startLocal >= 0) row[idx.startLocal] = startLocal;
      if (idx.endLocal >= 0) row[idx.endLocal] = endLocal;
      if (idx.startZ >= 0) row[idx.startZ] = startZulu;
      if (idx.endZ >= 0) row[idx.endZ] = endZulu;
      if (idx.baseCode >= 0) row[idx.baseCode] = baseCode;
      if (idx.type >= 0) row[idx.type] = 'UNAVAILABLE';
      if (idx.source >= 0) row[idx.source] = 'CALENDAR_UNAVAILABLE_SYNC';
      if (idx.updatedAt >= 0) row[idx.updatedAt] = new Date();
      appendRows.push(row);
      imported++;
    });

    if (appendRows.length) {
      var startRow = availabilitySheet.getLastRow() + 1;
      availabilitySheet.getRange(startRow, 1, appendRows.length, headers.length).setValues(appendRows);
    }

    return {
      success: true,
      month: parsedMonth.monthKey,
      calendarId: calendarId,
      actor: actor.email,
      stats: {
        eventsRead: events.length,
        imported: imported,
        skippedNoEmail: skippedNoEmail,
        skippedMarkedAvailable: skippedAvailable
      }
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getSchedulerAvailability(payload) {
  try {
    _schedulerAssertPermission_('CAN_VIEW', 'getSchedulerAvailability');
    var body = (payload && typeof payload === 'object') ? payload : {};
    var month = String(body.month || '').trim();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_AVAILABILITY || 'SCHED_Staff_Availability', 'getSchedulerAvailability');
    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return { success: true, rows: [] };

    var headers = data[0];
    var monthIdx = _schedulerHeaderIndex_(headers, 'SCHEDULE_MONTH');
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      if (month && monthIdx >= 0 && String(data[i][monthIdx] || '').trim() !== month) continue;
      var rowObj = { rowNumber: i + 1 };
      for (var c = 0; c < headers.length; c++) {
        rowObj[String(headers[c] || '').trim()] = data[i][c];
      }
      rows.push(rowObj);
    }
    return { success: true, rows: rows };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _toolsHeaderIndexFromCandidates_(headerRow, candidates) {
  var headers = Array.isArray(headerRow) ? headerRow : [];
  var norms = headers.map(function(h) { return _toolsNormHeader_(h); });
  var list = Array.isArray(candidates) ? candidates : [candidates];
  for (var i = 0; i < list.length; i++) {
    var idx = norms.indexOf(_toolsNormHeader_(list[i]));
    if (idx >= 0) return idx;
  }
  return -1;
}

function _toolsTruthyFlag_(value) {
  var raw = String(value == null ? '' : value).trim().toUpperCase();
  return raw === 'Y' || raw === 'YES' || raw === 'TRUE' || raw === '1' || raw === 'SIM' || raw === 'ATIVO';
}

function _toolsCurrentUserEmail_() {
  try {
    return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  } catch (e) {
    return '';
  }
}

function _toolsStaffRecordFromRow_(headers, row, rowNumber) {
  var nameIdx = _toolsHeaderIndexFromCandidates_(headers, ['PILOT', 'NAME', 'STAFF_NAME', 'NOME']);
  var emailIdx = _toolsHeaderIndexFromCandidates_(headers, ['EMAIL', 'E_MAIL']);
  var staffIdIdx = _toolsHeaderIndexFromCandidates_(headers, ['STAFF_ID', 'PILOT_ID', 'ID']);
  var activeIdx = _toolsHeaderIndexFromCandidates_(headers, ['ACTIVE', 'ATIVO']);
  var roleIdx = _toolsHeaderIndexFromCandidates_(headers, ['PRIMARY_ROLE', 'ROLE', 'FUNCAO', 'FUNC\u00c3O']);
  var notesIdx = _toolsHeaderIndexFromCandidates_(headers, ['NOTES']);
  var canEditDiscIdx = _toolsHeaderIndexFromCandidates_(headers, ['CAN_EDIT_DISCREPANCIES']);
  var canApproveDefIdx = _toolsHeaderIndexFromCandidates_(headers, ['CAN_APPROVE_DEFERMENTS']);
  var canEditMxIdx = _toolsHeaderIndexFromCandidates_(headers, ['CAN_EDIT_MAINTENANCE']);
  var canFollowIdx = _toolsHeaderIndexFromCandidates_(headers, ['CAN_FLIGHT_FOLLOW']);
  var canCoordIdx = _toolsHeaderIndexFromCandidates_(headers, ['CAN_COORDINATE_FLIGHTS']);
  var canStockIdx = _toolsHeaderIndexFromCandidates_(headers, ['CAN_MANAGE_STOCKROOM']);
  var canInstructIdx = _toolsHeaderIndexFromCandidates_(headers, ['CAN_INSTRUCT']);
  var canInspectIdx = _toolsHeaderIndexFromCandidates_(headers, ['CAN_INSPECT']);
  return {
    rowNumber: rowNumber,
    staffName: nameIdx >= 0 ? String(row[nameIdx] || '').trim() : '',
    email: emailIdx >= 0 ? String(row[emailIdx] || '').trim().toLowerCase() : '',
    staffId: staffIdIdx >= 0 ? String(row[staffIdIdx] || '').trim() : '',
    active: activeIdx >= 0 ? _toolsTruthyFlag_(row[activeIdx]) : true,
    primaryRole: roleIdx >= 0 ? String(row[roleIdx] || '').trim() : '',
    notes: notesIdx >= 0 ? String(row[notesIdx] || '').trim() : '',
    canEditDiscrepancies: canEditDiscIdx >= 0 ? _toolsTruthyFlag_(row[canEditDiscIdx]) : false,
    canApproveDeferments: canApproveDefIdx >= 0 ? _toolsTruthyFlag_(row[canApproveDefIdx]) : false,
    canEditMaintenance: canEditMxIdx >= 0 ? _toolsTruthyFlag_(row[canEditMxIdx]) : false,
    canFlightFollow: canFollowIdx >= 0 ? _toolsTruthyFlag_(row[canFollowIdx]) : false,
    canCoordinateFlights: canCoordIdx >= 0 ? _toolsTruthyFlag_(row[canCoordIdx]) : false,
    canManageStockroom: canStockIdx >= 0 ? _toolsTruthyFlag_(row[canStockIdx]) : false,
    canInstruct: canInstructIdx >= 0 ? _toolsTruthyFlag_(row[canInstructIdx]) : false,
    canInspect: canInspectIdx >= 0 ? _toolsTruthyFlag_(row[canInspectIdx]) : false
  };
}

function getToolsStaffSetupData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pilotsSheet = getRequiredSheet_(ss, APP_SHEETS.PILOTS, 'getToolsStaffSetupData');
    var roleSheet = getRequiredSheet_(ss, APP_SHEETS.STAFF_ROLES || 'REF_Staff_Roles', 'getToolsStaffSetupData');
    var moduleSheet = getRequiredSheet_(ss, APP_SHEETS.TRAINING_MODULES || 'REF_Training_Modules', 'getToolsStaffSetupData');

    var pData = pilotsSheet.getDataRange().getValues();
    var pHeaders = pData.length ? pData[0] : [];
    var staffRows = [];
    for (var i = 1; i < pData.length; i++) {
      var record = _toolsStaffRecordFromRow_(pHeaders, pData[i], i + 1);
      if (!record.staffName && !record.email) continue;
      staffRows.push(record);
    }
    staffRows.sort(function(a, b) { return String(a.staffName || '').localeCompare(String(b.staffName || '')); });

    var roleData = roleSheet.getDataRange().getValues();
    var roleHeaders = roleData.length ? roleData[0] : [];
    var roleCodeIdx = _toolsHeaderIndexFromCandidates_(roleHeaders, ['ROLE_CODE']);
    var roleNameIdx = _toolsHeaderIndexFromCandidates_(roleHeaders, ['ROLE_NAME']);
    var roleActiveIdx = _toolsHeaderIndexFromCandidates_(roleHeaders, ['ACTIVE']);
    var roles = [];
    for (var r = 1; r < roleData.length; r++) {
      if (roleCodeIdx < 0) continue;
      var code = String(roleData[r][roleCodeIdx] || '').trim();
      if (!code) continue;
      var active = roleActiveIdx < 0 ? true : _toolsTruthyFlag_(roleData[r][roleActiveIdx]);
      if (!active) continue;
      roles.push({
        roleCode: code,
        roleName: roleNameIdx >= 0 ? String(roleData[r][roleNameIdx] || code).trim() : code
      });
    }

    var mData = moduleSheet.getDataRange().getValues();
    var mHeaders = mData.length ? mData[0] : [];
    var moduleIdIdx = _toolsHeaderIndexFromCandidates_(mHeaders, ['MODULE_ID']);
    var moduleNameIdx = _toolsHeaderIndexFromCandidates_(mHeaders, ['MODULE_NAME']);
    var moduleRoleIdx = _toolsHeaderIndexFromCandidates_(mHeaders, ['ROLE_CODE']);
    var moduleComponentIdx = _toolsHeaderIndexFromCandidates_(mHeaders, ['COMPONENT']);
    var moduleTypeIdx = _toolsHeaderIndexFromCandidates_(mHeaders, ['MODULE_TYPE']);
    var modulePassIdx = _toolsHeaderIndexFromCandidates_(mHeaders, ['PASS_THRESHOLD']);
    var modulePracticalIdx = _toolsHeaderIndexFromCandidates_(mHeaders, ['REQUIRES_PRACTICAL']);
    var moduleRecurrentIdx = _toolsHeaderIndexFromCandidates_(mHeaders, ['RECURRENT_DAYS']);
    var moduleNotesIdx = _toolsHeaderIndexFromCandidates_(mHeaders, ['NOTES']);
    var moduleActiveIdx = _toolsHeaderIndexFromCandidates_(mHeaders, ['ACTIVE']);
    var modules = [];
    for (var m = 1; m < mData.length; m++) {
      if (moduleIdIdx < 0) continue;
      var modId = String(mData[m][moduleIdIdx] || '').trim();
      if (!modId) continue;
      var modActive = moduleActiveIdx < 0 ? true : _toolsTruthyFlag_(mData[m][moduleActiveIdx]);
      modules.push({
        rowNumber: m + 1,
        moduleId: modId,
        moduleName: moduleNameIdx >= 0 ? String(mData[m][moduleNameIdx] || modId).trim() : modId,
        roleCode: moduleRoleIdx >= 0 ? String(mData[m][moduleRoleIdx] || '').trim() : '',
        component: moduleComponentIdx >= 0 ? String(mData[m][moduleComponentIdx] || 'THEORY').trim().toUpperCase() : 'THEORY',
        moduleType: moduleTypeIdx >= 0 ? String(mData[m][moduleTypeIdx] || 'INITIAL').trim().toUpperCase() : 'INITIAL',
        passThreshold: modulePassIdx >= 0 ? String(mData[m][modulePassIdx] || '').trim() : '',
        requiresPractical: modulePracticalIdx >= 0 ? _toolsTruthyFlag_(mData[m][modulePracticalIdx]) : false,
        recurrentDays: moduleRecurrentIdx >= 0 ? String(mData[m][moduleRecurrentIdx] || '').trim() : '',
        active: modActive,
        notes: moduleNotesIdx >= 0 ? String(mData[m][moduleNotesIdx] || '').trim() : ''
      });
    }

    var userEmail = _toolsCurrentUserEmail_();
    var me = { email: userEmail, staffName: '', staffId: '' };
    if (userEmail) {
      for (var s = 0; s < staffRows.length; s++) {
        if (String(staffRows[s].email || '').trim().toLowerCase() === userEmail) {
          me.staffName = staffRows[s].staffName;
          me.staffId = staffRows[s].staffId;
          break;
        }
      }
    }

    return {
      success: true,
      currentUser: me,
      staff: staffRows,
      roles: roles,
      modules: modules
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function saveToolsStaffProfile(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.PILOTS, 'saveToolsStaffProfile');
    var headers = _toolsSheetHeaderRow_(sh);
    var data = sh.getDataRange().getValues();

    var rowNumber = Number(body.rowNumber || 0);
    var email = String(body.email || '').trim().toLowerCase();
    var staffId = String(body.staffId || '').trim();
    var staffName = String(body.staffName || '').trim();
    if (!staffName) return { success: false, error: 'Staff name is required' };
    if (!email) return { success: false, error: 'Staff email is required' };

    var emailIdx = _toolsHeaderIndexFromCandidates_(headers, ['EMAIL', 'E_MAIL']);
    var staffIdIdx = _toolsHeaderIndexFromCandidates_(headers, ['STAFF_ID']);
    if (rowNumber < 2) {
      for (var i = 1; i < data.length; i++) {
        var rowEmail = emailIdx >= 0 ? String(data[i][emailIdx] || '').trim().toLowerCase() : '';
        var rowStaffId = staffIdIdx >= 0 ? String(data[i][staffIdIdx] || '').trim() : '';
        if ((email && rowEmail === email) || (staffId && rowStaffId && rowStaffId === staffId)) {
          rowNumber = i + 1;
          break;
        }
      }
    }

    var dataMap = {
      STAFF_ID: staffId,
      EMAIL: email,
      PILOT: staffName,
      PRIMARY_ROLE: String(body.primaryRole || '').trim(),
      ACTIVE: _toolsTruthyFlag_(body.active) ? 'Y' : 'N',
      CAN_EDIT_DISCREPANCIES: _toolsTruthyFlag_(body.canEditDiscrepancies) ? 'Y' : 'N',
      CAN_APPROVE_DEFERMENTS: _toolsTruthyFlag_(body.canApproveDeferments) ? 'Y' : 'N',
      CAN_EDIT_MAINTENANCE: _toolsTruthyFlag_(body.canEditMaintenance) ? 'Y' : 'N',
      CAN_FLIGHT_FOLLOW: _toolsTruthyFlag_(body.canFlightFollow) ? 'Y' : 'N',
      CAN_COORDINATE_FLIGHTS: _toolsTruthyFlag_(body.canCoordinateFlights) ? 'Y' : 'N',
      CAN_MANAGE_STOCKROOM: _toolsTruthyFlag_(body.canManageStockroom) ? 'Y' : 'N',
      CAN_INSTRUCT: _toolsTruthyFlag_(body.canInstruct) ? 'Y' : 'N',
      CAN_INSPECT: _toolsTruthyFlag_(body.canInspect) ? 'Y' : 'N',
      NOTES: String(body.notes || '').trim()
    };

    var record = headers.map(function(header) {
      var key = _toolsNormHeader_(header);
      return Object.prototype.hasOwnProperty.call(dataMap, key) ? dataMap[key] : '';
    });

    if (rowNumber >= 2) {
      var current = sh.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
      var merged = headers.map(function(header, idx) {
        var key = _toolsNormHeader_(header);
        return Object.prototype.hasOwnProperty.call(dataMap, key) ? dataMap[key] : current[idx];
      });
      sh.getRange(rowNumber, 1, 1, merged.length).setValues([merged]);
      return { success: true, action: 'updated', rowNumber: rowNumber };
    }

    sh.appendRow(record);
    return { success: true, action: 'created', rowNumber: sh.getLastRow() };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function saveToolsTrainingModule(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var moduleId = String(body.moduleId || '').trim().toUpperCase();
    var moduleName = String(body.moduleName || '').trim();
    if (!moduleId) return { success: false, error: 'Module ID is required' };
    if (!moduleName) return { success: false, error: 'Module name is required' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.TRAINING_MODULES || 'REF_Training_Modules', 'saveToolsTrainingModule');
    var headers = _toolsSheetHeaderRow_(sh);
    var data = sh.getDataRange().getValues();
    var idIdx = _toolsHeaderIndexFromCandidates_(headers, ['MODULE_ID']);
    var rowNumber = Number(body.rowNumber || 0);

    if (rowNumber < 2 && idIdx >= 0) {
      for (var i = 1; i < data.length; i++) {
        var id = String(data[i][idIdx] || '').trim().toUpperCase();
        if (id && id === moduleId) {
          rowNumber = i + 1;
          break;
        }
      }
    }

    var dataMap = {
      MODULE_ID: moduleId,
      MODULE_NAME: moduleName,
      ROLE_CODE: String(body.roleCode || '').trim(),
      COMPONENT: String(body.component || '').trim().toUpperCase() || 'THEORY',
      MODULE_TYPE: String(body.moduleType || '').trim().toUpperCase() || 'INITIAL',
      PASS_THRESHOLD: String(body.passThreshold == null ? '' : body.passThreshold).trim(),
      REQUIRES_PRACTICAL: _toolsTruthyFlag_(body.requiresPractical) ? 'Y' : 'N',
      RECURRENT_DAYS: String(body.recurrentDays == null ? '' : body.recurrentDays).trim(),
      CLASSROOM_COURSE_ID: String(body.classroomCourseId || '').trim(),
      CLASSROOM_COURSEWORK_ID: String(body.classroomCourseworkId || '').trim(),
      ACTIVE: _toolsTruthyFlag_(body.active) ? 'Y' : 'N',
      NOTES: String(body.notes || '').trim()
    };

    if (rowNumber >= 2) {
      var current = sh.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
      var merged = headers.map(function(header, idx) {
        var key = _toolsNormHeader_(header);
        return Object.prototype.hasOwnProperty.call(dataMap, key) ? dataMap[key] : current[idx];
      });
      sh.getRange(rowNumber, 1, 1, merged.length).setValues([merged]);
      return { success: true, action: 'updated', rowNumber: rowNumber, moduleId: moduleId };
    }

    var row = headers.map(function(header) {
      var key = _toolsNormHeader_(header);
      return Object.prototype.hasOwnProperty.call(dataMap, key) ? dataMap[key] : '';
    });
    sh.appendRow(row);
    return { success: true, action: 'created', rowNumber: sh.getLastRow(), moduleId: moduleId };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _toolsFindStaffByEmailOrId_(staffRows, staffEmail, staffId) {
  var email = String(staffEmail || '').trim().toLowerCase();
  var sid = String(staffId || '').trim();
  for (var i = 0; i < staffRows.length; i++) {
    var row = staffRows[i] || {};
    if (email && String(row.email || '').trim().toLowerCase() === email) return row;
    if (sid && String(row.staffId || '').trim() === sid) return row;
  }
  return null;
}

function saveToolsPracticalEvaluation(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var staffEmail = String(body.staffEmail || '').trim().toLowerCase();
    var staffIdInput = String(body.staffId || '').trim();
    var moduleId = String(body.moduleId || '').trim().toUpperCase();
    var evaluator = String(body.evaluator || '').trim();
    var result = String(body.result || '').trim().toUpperCase();
    var evaluatedAt = String(body.evaluatedAt || '').trim();
    if (!staffEmail && !staffIdInput) return { success: false, error: 'Staff member is required' };
    if (!moduleId) return { success: false, error: 'Module is required' };
    if (!evaluator) return { success: false, error: 'Evaluator is required' };
    if (result !== 'PASS' && result !== 'FAIL') return { success: false, error: 'Result must be PASS or FAIL' };
    if (!evaluatedAt) return { success: false, error: 'Evaluation date is required' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = getRequiredSheet_(ss, APP_SHEETS.PILOTS, 'saveToolsPracticalEvaluation');
    var modulesSheet = getRequiredSheet_(ss, APP_SHEETS.TRAINING_MODULES || 'REF_Training_Modules', 'saveToolsPracticalEvaluation');
    var practicalSheet = getRequiredSheet_(ss, APP_SHEETS.STAFF_PRACTICALS || 'DB_Staff_Practical_Evaluations', 'saveToolsPracticalEvaluation');
    var trainingSheet = getRequiredSheet_(ss, APP_SHEETS.STAFF_TRAINING || 'DB_Staff_Training_Records', 'saveToolsPracticalEvaluation');

    var sData = staffSheet.getDataRange().getValues();
    var sHeaders = sData.length ? sData[0] : [];
    var staffRows = [];
    for (var i = 1; i < sData.length; i++) {
      var rec = _toolsStaffRecordFromRow_(sHeaders, sData[i], i + 1);
      if (rec.staffName || rec.email || rec.staffId) staffRows.push(rec);
    }
    var staff = _toolsFindStaffByEmailOrId_(staffRows, staffEmail, staffIdInput);
    if (!staff) return { success: false, error: 'Staff member not found in DB_Pilots' };

    var mData = modulesSheet.getDataRange().getValues();
    var mHeaders = mData.length ? mData[0] : [];
    var mIdIdx = _toolsHeaderIndexFromCandidates_(mHeaders, ['MODULE_ID']);
    var mNameIdx = _toolsHeaderIndexFromCandidates_(mHeaders, ['MODULE_NAME']);
    var mRoleIdx = _toolsHeaderIndexFromCandidates_(mHeaders, ['ROLE_CODE']);
    var module = null;
    for (var m = 1; m < mData.length; m++) {
      var id = mIdIdx >= 0 ? String(mData[m][mIdIdx] || '').trim().toUpperCase() : '';
      if (id && id === moduleId) {
        module = {
          moduleId: id,
          moduleName: mNameIdx >= 0 ? String(mData[m][mNameIdx] || id).trim() : id,
          roleCode: mRoleIdx >= 0 ? String(mData[m][mRoleIdx] || '').trim() : ''
        };
        break;
      }
    }
    if (!module) return { success: false, error: 'Module not found: ' + moduleId };

    var now = new Date();
    var stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
    var practicalId = 'PRAT-' + stamp;
    var recordedBy = _toolsCurrentUserEmail_() || evaluator;

    var pHeaders = _toolsSheetHeaderRow_(practicalSheet);
    var pMap = {
      PRACTICAL_ID: practicalId,
      STAFF_ID: staff.staffId,
      STAFF_EMAIL: staff.email,
      MODULE_ID: module.moduleId,
      AIRCRAFT: String(body.aircraft || '').trim().toUpperCase(),
      EVALUATOR: evaluator,
      RESULT: result,
      EVALUATED_AT: evaluatedAt,
      TACH_AT_EVAL: String(body.tachAtEval == null ? '' : body.tachAtEval).trim(),
      LOCATION: String(body.location || '').trim(),
      OBSERVATIONS: String(body.observations || '').trim(),
      RECORDED_BY: recordedBy,
      RECORDED_AT: now
    };
    var pRow = pHeaders.map(function(header) {
      var key = _toolsNormHeader_(header);
      return Object.prototype.hasOwnProperty.call(pMap, key) ? pMap[key] : '';
    });
    practicalSheet.appendRow(pRow);

    var tHeaders = _toolsSheetHeaderRow_(trainingSheet);
    var tMap = {
      RECORD_ID: 'TRN-' + stamp,
      STAFF_ID: staff.staffId,
      STAFF_EMAIL: staff.email,
      MODULE_ID: module.moduleId,
      MODULE_NAME: module.moduleName,
      ROLE_CODE: module.roleCode,
      SOURCE: 'PRACTICAL_FORM',
      PRACTICAL_PASSED: result === 'PASS' ? 'Y' : 'N',
      PRACTICAL_EVALUATOR: evaluator,
      PRACTICAL_COMPLETED_AT: evaluatedAt,
      RECORDED_BY: recordedBy,
      RECORDED_AT: now,
      NOTES: String(body.observations || '').trim()
    };
    var tRow = tHeaders.map(function(header) {
      var key = _toolsNormHeader_(header);
      return Object.prototype.hasOwnProperty.call(tMap, key) ? tMap[key] : '';
    });
    trainingSheet.appendRow(tRow);

    return {
      success: true,
      practicalId: practicalId,
      trainingRecordId: tMap.RECORD_ID,
      staffName: staff.staffName,
      moduleName: module.moduleName,
      result: result
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _toolsReadinessDateValue_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  var raw = String(value == null ? '' : value).trim();
  if (!raw) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    var p = raw.split('-');
    var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    return isNaN(d.getTime()) ? null : d;
  }
  var parsed = new Date(raw);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function getToolsStaffReadiness(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var staffEmail = String(body.staffEmail || '').trim().toLowerCase();
    var staffId = String(body.staffId || '').trim();
    var moduleFilter = String(body.moduleId || '').trim().toUpperCase();

    var setup = getToolsStaffSetupData();
    if (!setup || !setup.success) return { success: false, error: (setup && setup.error) || 'Could not load staff data' };
    var staff = _toolsFindStaffByEmailOrId_(setup.staff || [], staffEmail, staffId);
    if (!staff) return { success: false, error: 'Staff member not found' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var moduleSheet = getRequiredSheet_(ss, APP_SHEETS.TRAINING_MODULES || 'REF_Training_Modules', 'getToolsStaffReadiness');
    var trainingSheet = getRequiredSheet_(ss, APP_SHEETS.STAFF_TRAINING || 'DB_Staff_Training_Records', 'getToolsStaffReadiness');
    var practicalSheet = getRequiredSheet_(ss, APP_SHEETS.STAFF_PRACTICALS || 'DB_Staff_Practical_Evaluations', 'getToolsStaffReadiness');

    var moduleData = moduleSheet.getDataRange().getValues();
    var moduleHeaders = moduleData.length ? moduleData[0] : [];
    var modIdIdx = _toolsHeaderIndexFromCandidates_(moduleHeaders, ['MODULE_ID']);
    var modNameIdx = _toolsHeaderIndexFromCandidates_(moduleHeaders, ['MODULE_NAME']);
    var modRoleIdx = _toolsHeaderIndexFromCandidates_(moduleHeaders, ['ROLE_CODE']);
    var modReqPracticalIdx = _toolsHeaderIndexFromCandidates_(moduleHeaders, ['REQUIRES_PRACTICAL']);
    var modRecDaysIdx = _toolsHeaderIndexFromCandidates_(moduleHeaders, ['RECURRENT_DAYS']);
    var modActiveIdx = _toolsHeaderIndexFromCandidates_(moduleHeaders, ['ACTIVE']);

    var modules = [];
    for (var m = 1; m < moduleData.length; m++) {
      if (modIdIdx < 0) continue;
      var id = String(moduleData[m][modIdIdx] || '').trim().toUpperCase();
      if (!id) continue;
      if (moduleFilter && id !== moduleFilter) continue;
      var roleCode = modRoleIdx >= 0 ? String(moduleData[m][modRoleIdx] || '').trim() : '';
      if (!moduleFilter && roleCode && staff.primaryRole && roleCode !== staff.primaryRole) continue;
      var active = modActiveIdx < 0 ? true : _toolsTruthyFlag_(moduleData[m][modActiveIdx]);
      if (!active) continue;
      modules.push({
        moduleId: id,
        moduleName: modNameIdx >= 0 ? String(moduleData[m][modNameIdx] || id).trim() : id,
        roleCode: roleCode,
        requiresPractical: modReqPracticalIdx >= 0 ? _toolsTruthyFlag_(moduleData[m][modReqPracticalIdx]) : false,
        recurrentDays: modRecDaysIdx >= 0 ? Number(moduleData[m][modRecDaysIdx] || 0) : 0
      });
    }

    var trData = trainingSheet.getDataRange().getValues();
    var trHeaders = trData.length ? trData[0] : [];
    var trEmailIdx = _toolsHeaderIndexFromCandidates_(trHeaders, ['STAFF_EMAIL']);
    var trStaffIdIdx = _toolsHeaderIndexFromCandidates_(trHeaders, ['STAFF_ID']);
    var trModuleIdx = _toolsHeaderIndexFromCandidates_(trHeaders, ['MODULE_ID']);
    var trTheoryPassedIdx = _toolsHeaderIndexFromCandidates_(trHeaders, ['THEORY_PASSED']);
    var trTheoryCompletedIdx = _toolsHeaderIndexFromCandidates_(trHeaders, ['THEORY_COMPLETED_AT']);
    var trPracticalPassedIdx = _toolsHeaderIndexFromCandidates_(trHeaders, ['PRACTICAL_PASSED']);
    var trPracticalCompletedIdx = _toolsHeaderIndexFromCandidates_(trHeaders, ['PRACTICAL_COMPLETED_AT']);

    var prData = practicalSheet.getDataRange().getValues();
    var prHeaders = prData.length ? prData[0] : [];
    var prEmailIdx = _toolsHeaderIndexFromCandidates_(prHeaders, ['STAFF_EMAIL']);
    var prStaffIdIdx = _toolsHeaderIndexFromCandidates_(prHeaders, ['STAFF_ID']);
    var prModuleIdx = _toolsHeaderIndexFromCandidates_(prHeaders, ['MODULE_ID']);
    var prResultIdx = _toolsHeaderIndexFromCandidates_(prHeaders, ['RESULT']);
    var prEvalDateIdx = _toolsHeaderIndexFromCandidates_(prHeaders, ['EVALUATED_AT']);

    var now = new Date();
    var rows = [];
    var readyCount = 0;

    modules.forEach(function(module) {
      var theoryPassed = false;
      var practicalPassed = false;
      var lastTheoryDate = null;
      var lastPracticalDate = null;

      for (var i = 1; i < trData.length; i++) {
        var row = trData[i];
        var rowEmail = trEmailIdx >= 0 ? String(row[trEmailIdx] || '').trim().toLowerCase() : '';
        var rowStaffId = trStaffIdIdx >= 0 ? String(row[trStaffIdIdx] || '').trim() : '';
        var rowModule = trModuleIdx >= 0 ? String(row[trModuleIdx] || '').trim().toUpperCase() : '';
        var sameStaff = (staff.email && rowEmail === staff.email) || (staff.staffId && rowStaffId === staff.staffId);
        if (!sameStaff || rowModule !== module.moduleId) continue;

        if (trTheoryPassedIdx >= 0 && _toolsTruthyFlag_(row[trTheoryPassedIdx])) theoryPassed = true;
        if (trPracticalPassedIdx >= 0 && _toolsTruthyFlag_(row[trPracticalPassedIdx])) practicalPassed = true;

        var td = trTheoryCompletedIdx >= 0 ? _toolsReadinessDateValue_(row[trTheoryCompletedIdx]) : null;
        if (td && (!lastTheoryDate || td > lastTheoryDate)) lastTheoryDate = td;
        var pd = trPracticalCompletedIdx >= 0 ? _toolsReadinessDateValue_(row[trPracticalCompletedIdx]) : null;
        if (pd && (!lastPracticalDate || pd > lastPracticalDate)) lastPracticalDate = pd;
      }

      for (var p = 1; p < prData.length; p++) {
        var prow = prData[p];
        var prowEmail = prEmailIdx >= 0 ? String(prow[prEmailIdx] || '').trim().toLowerCase() : '';
        var prowStaffId = prStaffIdIdx >= 0 ? String(prow[prStaffIdIdx] || '').trim() : '';
        var prowModule = prModuleIdx >= 0 ? String(prow[prModuleIdx] || '').trim().toUpperCase() : '';
        var sameStaffEval = (staff.email && prowEmail === staff.email) || (staff.staffId && prowStaffId === staff.staffId);
        if (!sameStaffEval || prowModule !== module.moduleId) continue;
        var passEval = prResultIdx >= 0 ? String(prow[prResultIdx] || '').trim().toUpperCase() === 'PASS' : false;
        if (passEval) practicalPassed = true;
        var pdt = prEvalDateIdx >= 0 ? _toolsReadinessDateValue_(prow[prEvalDateIdx]) : null;
        if (pdt && (!lastPracticalDate || pdt > lastPracticalDate)) lastPracticalDate = pdt;
      }

      var lastCompletion = lastPracticalDate || lastTheoryDate;
      var dueDate = null;
      if (module.recurrentDays > 0 && lastCompletion) {
        dueDate = new Date(lastCompletion.getTime());
        dueDate.setDate(dueDate.getDate() + module.recurrentDays);
      }

      var recurrentValid = !dueDate || dueDate >= now;
      var moduleReady = theoryPassed && (!module.requiresPractical || practicalPassed) && recurrentValid;
      if (moduleReady) readyCount++;

      rows.push({
        moduleId: module.moduleId,
        moduleName: module.moduleName,
        roleCode: module.roleCode,
        theoryPassed: theoryPassed,
        practicalRequired: module.requiresPractical,
        practicalPassed: practicalPassed,
        recurrentDays: module.recurrentDays,
        lastCompletionDate: lastCompletion ? Utilities.formatDate(lastCompletion, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        dueDate: dueDate ? Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        ready: moduleReady
      });
    });

    return {
      success: true,
      staff: {
        staffName: staff.staffName,
        email: staff.email,
        staffId: staff.staffId,
        primaryRole: staff.primaryRole
      },
      summary: {
        moduleCount: rows.length,
        readyCount: readyCount,
        overallReady: rows.length > 0 && readyCount === rows.length
      },
      rows: rows
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function saveAirportFuelProfile(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var code = String(body.icao || '').trim().toUpperCase();
    if (!code) return { success: false, error: 'ICAO required' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var airportSheet = getRequiredSheet_(ss, APP_SHEETS.AIRPORTS, 'saveAirportFuelProfile');
    var airportHeaders = _toolsSheetHeaderRow_(airportSheet);
    var airportData = airportSheet.getDataRange().getValues();
    var airportNorms = airportHeaders.map(function(h) { return _toolsNormHeader_(h); });
    var airportIcaoIdx = airportNorms.indexOf('ICAO');
    if (airportIcaoIdx < 0) return { success: false, error: 'ICAO column not found in DB_Airports' };

    var airportRowNumber = 0;
    var airportRow = null;
    for (var ai = 1; ai < airportData.length; ai++) {
      if (String(airportData[ai][airportIcaoIdx] || '').trim().toUpperCase() === code) {
        airportRowNumber = ai + 1;
        airportRow = airportData[ai];
        break;
      }
    }
    if (!airportRowNumber) return { success: false, error: 'Airport not found in DB_Airports: ' + code };

    var airportPayload = (body.airport && typeof body.airport === 'object') ? body.airport : {};
    var airportNewRow = airportHeaders.map(function(header, idx) {
      var label = String(header || '').trim();
      var key = _toolsNormHeader_(label);
      if (Object.prototype.hasOwnProperty.call(airportPayload, key)) return airportPayload[key];
      if (Object.prototype.hasOwnProperty.call(airportPayload, label)) return airportPayload[label];
      return airportRow[idx];
    });
    airportSheet.getRange(airportRowNumber, 1, 1, airportNewRow.length).setValues([airportNewRow]);

    // Append any new columns from the JSON that don't exist in the sheet yet
    var newColNorms = [];
    var newColVals = [];
    Object.keys(airportPayload).forEach(function(key) {
      var norm = _toolsNormHeader_(key);
      if (norm && airportNorms.indexOf(norm) < 0 && newColNorms.indexOf(norm) < 0) {
        newColNorms.push(norm);
        newColVals.push(airportPayload[key] !== undefined ? airportPayload[key] : '');
      }
    });
    if (newColNorms.length > 0) {
      var appendStart = airportHeaders.length + 1;
      airportSheet.getRange(1, appendStart, 1, newColNorms.length).setValues([newColNorms]);
      airportSheet.getRange(airportRowNumber, appendStart, 1, newColNorms.length).setValues([newColVals]);
    }

    var contactUpdated = false;
    var contactCreated = false;
    var contactPayload = (body.contact && typeof body.contact === 'object') ? body.contact : {};
    var contactHasData = Object.keys(contactPayload).some(function(k) {
      return String(contactPayload[k] == null ? '' : contactPayload[k]).trim() !== '';
    });

    if (contactHasData || body.hasFuel === true) {
      var contactSheet = ss.getSheetByName(APP_SHEETS.CONTACTS || 'DB_Contacts');
      if (contactSheet) {
        var contactData = contactSheet.getDataRange().getValues();
        if (contactData.length >= 1) {
          var contactHeaders = contactData[0];
          var contactNorms = contactHeaders.map(function(h) { return _toolsNormHeader_(h); });

          var contactIcaoIdx = -1;
          ['ICAO', 'C_DIGO', 'CODIGO', 'CDIGO', 'C__DIGO'].some(function(name) {
            var idx = contactNorms.indexOf(_toolsNormHeader_(name));
            if (idx >= 0) { contactIcaoIdx = idx; return true; }
            return false;
          });
          if (contactIcaoIdx < 0) return { success: false, error: 'ICAO/Código column not found in DB_Contacts' };

          var contactFuelIdx = -1;
          ['POSSUI_COMBUST_VEL_', 'POSSUI_COMBUSTIVEL_', 'POSSUI_COMBUSTIVEL'].some(function(name) {
            var idx = contactNorms.indexOf(_toolsNormHeader_(name));
            if (idx >= 0) { contactFuelIdx = idx; return true; }
            return false;
          });

          var contactRowIndex = -1;
          for (var ci = 1; ci < contactData.length; ci++) {
            if (String(contactData[ci][contactIcaoIdx] || '').trim().toUpperCase() === code) {
              contactRowIndex = ci;
              break;
            }
          }

          if (contactRowIndex >= 0) {
            var existingContact = contactData[contactRowIndex];
            var updatedContact = contactHeaders.map(function(header, idx) {
              var label = String(header || '').trim();
              var key = _toolsNormHeader_(label);
              if (Object.prototype.hasOwnProperty.call(contactPayload, key)) return contactPayload[key];
              if (Object.prototype.hasOwnProperty.call(contactPayload, label)) return contactPayload[label];
              return existingContact[idx];
            });
            if (contactFuelIdx >= 0) {
              updatedContact[contactFuelIdx] = body.hasFuel === true ? 'YES' : (body.hasFuel === false ? 'NO' : updatedContact[contactFuelIdx]);
            }
            updatedContact[contactIcaoIdx] = code;
            contactSheet.getRange(contactRowIndex + 1, 1, 1, updatedContact.length).setValues([updatedContact]);
            contactUpdated = true;
          } else {
            var contactNew = contactHeaders.map(function(header) {
              var label = String(header || '').trim();
              var key = _toolsNormHeader_(label);
              if (Object.prototype.hasOwnProperty.call(contactPayload, key)) return contactPayload[key];
              if (Object.prototype.hasOwnProperty.call(contactPayload, label)) return contactPayload[label];
              return '';
            });
            contactNew[contactIcaoIdx] = code;
            if (contactFuelIdx >= 0) contactNew[contactFuelIdx] = body.hasFuel === true ? 'YES' : 'NO';
            contactSheet.appendRow(contactNew);
            contactCreated = true;
          }
        }
      }
    }

    var cacheUpdated = false;
    var cacheCreated = false;
    var cacheRemoved = false;
    if (body.fuelCacheEnabled === true) {
      var cacheSheet = ss.getSheetByName(APP_SHEETS.FUEL_CACHES);
      if (cacheSheet) {
        var cacheData = cacheSheet.getDataRange().getValues();
        if (cacheData.length >= 1) {
          var cacheHeaders = cacheData[0];
          var cacheNorms = cacheHeaders.map(function(h) { return _toolsNormHeader_(h); });
          var cachePayload = (body.fuelCache && typeof body.fuelCache === 'object') ? body.fuelCache : {};

          var cacheIcaoIdx = cacheNorms.indexOf('ICAO');
          if (cacheIcaoIdx < 0) return { success: false, error: 'ICAO column not found in DB_Fuel_Caches' };

          var cacheRowIndex = -1;
          for (var fi = 1; fi < cacheData.length; fi++) {
            if (String(cacheData[fi][cacheIcaoIdx] || '').trim().toUpperCase() === code) {
              cacheRowIndex = fi;
              break;
            }
          }

          if (cacheRowIndex >= 0) {
            var existingCache = cacheData[cacheRowIndex];
            var updatedCache = cacheHeaders.map(function(header, idx) {
              var label = String(header || '').trim();
              var key = _toolsNormHeader_(label);
              if (Object.prototype.hasOwnProperty.call(cachePayload, key)) return cachePayload[key];
              if (Object.prototype.hasOwnProperty.call(cachePayload, label)) return cachePayload[label];
              return existingCache[idx];
            });
            updatedCache[cacheIcaoIdx] = code;
            cacheSheet.getRange(cacheRowIndex + 1, 1, 1, updatedCache.length).setValues([updatedCache]);
            cacheUpdated = true;
          } else {
            var cacheNew = cacheHeaders.map(function(header) {
              var label = String(header || '').trim();
              var key = _toolsNormHeader_(label);
              if (Object.prototype.hasOwnProperty.call(cachePayload, key)) return cachePayload[key];
              if (Object.prototype.hasOwnProperty.call(cachePayload, label)) return cachePayload[label];
              return '';
            });
            cacheNew[cacheIcaoIdx] = code;
            cacheSheet.appendRow(cacheNew);
            cacheCreated = true;
          }
        }
      }
    } else {
      // Explicitly clear stale cache inventory row when mode is not "fuel cache".
      var cleanupSheet = ss.getSheetByName(APP_SHEETS.FUEL_CACHES);
      if (cleanupSheet) {
        var cleanupData = cleanupSheet.getDataRange().getValues();
        if (cleanupData.length >= 2) {
          var cleanupHeaders = cleanupData[0].map(function(h) { return _toolsNormHeader_(h); });
          var cleanupIcaoIdx = cleanupHeaders.indexOf('ICAO');
          if (cleanupIcaoIdx >= 0) {
            for (var cr = cleanupData.length - 1; cr >= 1; cr--) {
              if (String(cleanupData[cr][cleanupIcaoIdx] || '').trim().toUpperCase() === code) {
                cleanupSheet.deleteRow(cr + 1);
                cacheRemoved = true;
              }
            }
          }
        }
      }
    }

    if (body.hasFuel === true || body.hasFuel === false) {
      var fuelRes = setAirportFuelAvailability(code, body.hasFuel);
      if (!fuelRes || !fuelRes.success) {
        return { success: false, error: (fuelRes && fuelRes.error) ? fuelRes.error : 'Failed to set airport fuel availability' };
      }
    }

    return {
      success: true,
      icao: code,
      airportRowNumber: airportRowNumber,
      contactUpdated: contactUpdated,
      contactCreated: contactCreated,
      cacheUpdated: cacheUpdated,
      cacheCreated: cacheCreated,
      cacheRemoved: cacheRemoved
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _donationImportsRootFolder_() {
  try {
    if (typeof _ensureDonationImportsDriveFolder_ === 'function') {
      return _ensureDonationImportsDriveFolder_();
    }
  } catch (e) {}

  var props = PropertiesService.getScriptProperties();
  var existingId = String(props.getProperty('DONATION_IMPORTS_FOLDER_ID') || '').trim();
  if (existingId) {
    try { return DriveApp.getFolderById(existingId); } catch (e) {}
  }
  var folderName = 'MBA_Donation_Imports';
  var folders = DriveApp.getFoldersByName(folderName);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  props.setProperty('DONATION_IMPORTS_FOLDER_ID', folder.getId());
  return folder;
}

function _donationBatchId_() {
  return 'BATCH_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT', 'yyyyMMdd_HHmmss') + '_' + Utilities.getUuid().slice(0, 4).toUpperCase();
}

function _donationNormHeader_(value) {
  return String(value || '').trim().toUpperCase().replace(/\s+/g, '_').replace(/[^A-Z0-9_]/g, '');
}

function _donationNormalizeText_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9 ]+/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function _donationNormalizeDonor_(value) {
  return _donationNormalizeText_(value);
}

function _donationIsoDate_(value) {
  if (!value && value !== 0) return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'GMT', 'yyyy-MM-dd');
  }
  var raw = String(value || '').trim();
  if (!raw) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  var br = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (br) {
    var dd = String(parseInt(br[1], 10)).padStart(2, '0');
    var mm = String(parseInt(br[2], 10)).padStart(2, '0');
    return br[3] + '-' + mm + '-' + dd;
  }
  var parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, 'GMT', 'yyyy-MM-dd');
  return '';
}

function _donationAmountBrl_(value) {
  if (value == null) return NaN;
  var raw = String(value).trim();
  if (!raw) return NaN;
  raw = raw.replace(/R\$/gi, '').replace(/\s+/g, '');
  if (/^-?\d{1,3}(\.\d{3})*,\d{2}$/.test(raw)) {
    raw = raw.replace(/\./g, '').replace(',', '.');
  } else if (/^-?\d+,\d{2}$/.test(raw)) {
    raw = raw.replace(',', '.');
  } else {
    raw = raw.replace(/,/g, '');
  }
  var num = parseFloat(raw);
  return isNaN(num) ? NaN : Math.round(num * 100) / 100;
}

function _donationDigest_(value) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(value || ''), Utilities.Charset.UTF_8);
  return bytes.map(function(b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function _donationSuggestedMap_(headers) {
  var list = Array.isArray(headers) ? headers : [];
  var normalized = list.map(_donationNormHeader_);
  function find(candidates) {
    for (var i = 0; i < candidates.length; i++) {
      var idx = normalized.indexOf(candidates[i]);
      if (idx >= 0) return list[idx];
    }
    return '';
  }
  return {
    donor: find(['DONOR', 'DONOR_NAME', 'NOME', 'NAME', 'CLIENTE', 'PAGADOR', 'BENEFICIARIO']),
    date: find(['DATE', 'DATA', 'TX_DATE', 'TRANSACTION_DATE', 'POSTED_DATE']),
    amount: find(['AMOUNT', 'AMOUNT_BRL', 'VALOR', 'VALUE', 'TOTAL']),
    description: find(['DESCRIPTION', 'HISTORICO', 'HISTORICO_DA_OPERACAO', 'DETAILS', 'MEMO', 'NOTES']),
    campaign: find(['CAMPANHA__PROJETO', 'CAMPANHA_PROJETO', 'CAMPANHA', 'PROJETO', 'CAMPAIGN', 'PROJECT', 'CAUSA', 'PROGRAMA', 'FUND_NAME'])
  };
}

function inspectDonationImportFile(payload) {
  try {
    var fileName = String(payload && payload.fileName || '').trim();
    var mimeType = String(payload && payload.mimeType || '').trim().toLowerCase();
    var base64Data = String(payload && payload.base64Data || '').trim();
    if (!fileName || !base64Data) return { success: false, error: 'File data missing' };

    var ext = fileName.split('.').pop().toLowerCase();
    if (ext === 'pdf' || mimeType.indexOf('pdf') >= 0) {
      return {
        success: true,
        fileType: 'pdf',
        canStage: false,
        error: 'PDF parsing is not enabled in Phase 1. Save the source file, then convert to CSV for staging.'
      };
    }

    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || 'text/csv', fileName);
    var text = blob.getDataAsString('UTF-8').replace(/^\uFEFF/, '');
    var rows = Utilities.parseCsv(text);
    if (!rows || rows.length < 2) return { success: false, error: 'CSV appears empty or has no data rows' };

    var headers = rows[0].map(function(h) { return String(h || '').trim(); });
    var sampleRows = rows.slice(1, 6).map(function(row) {
      var obj = {};
      headers.forEach(function(header, idx) { obj[header] = idx < row.length ? row[idx] : ''; });
      return obj;
    });

    var suggestedMap = _donationSuggestedMap_(headers);
    var campaignValues = [];
    if (suggestedMap.campaign) {
      var campColNorm = _donationNormHeader_(suggestedMap.campaign);
      var campIdx = headers.findIndex(function(h) { return _donationNormHeader_(h) === campColNorm; });
      if (campIdx >= 0) {
        var campCount = {};
        for (var ri = 1; ri < rows.length; ri++) {
          var campRaw = String(rows[ri][campIdx] || '').trim();
          if (campRaw) campCount[campRaw] = (campCount[campRaw] || 0) + 1;
        }
        campaignValues = Object.keys(campCount).sort().map(function(r) { return { raw: r, count: campCount[r] }; });
      }
    }
    return {
      success: true,
      fileType: 'csv',
      canStage: true,
      headers: headers,
      sampleRows: sampleRows,
      suggestedMap: suggestedMap,
      campaignValues: campaignValues
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _donationLedgerRecords_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getRequiredSheet_(ss, APP_SHEETS.DONATIONS_LEDGER, '_donationLedgerRecords_');
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(_donationNormHeader_);
  return data.slice(1).map(function(row) {
    var get = function(name) {
      var idx = headers.indexOf(name);
      return idx >= 0 ? row[idx] : '';
    };
    return {
      donationId: String(get('DONATION_ID') || '').trim(),
      donorNorm: _donationNormalizeDonor_(get('DONOR_NORMALIZED') || get('DONOR_RAW')),
      txDate: _donationIsoDate_(get('TX_DATE')),
      amount: _donationAmountBrl_(get('AMOUNT_BRL')),
      fundId: String(get('FUND_ID') || '').trim(),
      status: String(get('STATUS') || '').trim()
    };
  }).filter(function(r) { return r.donationId; });
}

function stageDonationImport(payload) {
  try {
    var fileName = String(payload && payload.fileName || '').trim();
    var mimeType = String(payload && payload.mimeType || '').trim().toLowerCase();
    var base64Data = String(payload && payload.base64Data || '').trim();
    var fundId = String(payload && payload.fundId || '').trim();
    var campaignColumn = String(payload && payload.campaignColumn || '').trim();
    var campaignMap = (payload && payload.campaignMap && typeof payload.campaignMap === 'object') ? payload.campaignMap : {};
    var importedBy = String(payload && payload.importedBy || Session.getActiveUser().getEmail() || 'Admin').trim();
    var map = payload && payload.columnMap ? payload.columnMap : {};
    if (!fileName || !base64Data) return { success: false, error: 'File data missing' };
    if (!fundId && Object.keys(campaignMap).length === 0) return { success: false, error: 'Fund is required (or provide a campaign-to-fund mapping)' };

    var ext = fileName.split('.').pop().toLowerCase();
    if (ext === 'pdf' || mimeType.indexOf('pdf') >= 0) {
      return { success: false, error: 'PDF parsing is not enabled in Phase 1. Convert the report to CSV before staging.' };
    }

    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || 'text/csv', fileName);
    var text = blob.getDataAsString('UTF-8').replace(/^\uFEFF/, '');
    var rows = Utilities.parseCsv(text);
    if (!rows || rows.length < 2) return { success: false, error: 'CSV appears empty or has no data rows' };

    var headers = rows[0].map(function(h) { return String(h || '').trim(); });
    var headerIndex = {};
    headers.forEach(function(header, idx) { headerIndex[_donationNormHeader_(header)] = idx; });
    function idxFor(headerName) {
      var idx = headerIndex[_donationNormHeader_(headerName)];
      return typeof idx === 'number' ? idx : -1;
    }

    var donorIdx = idxFor(map.donor);
    var dateIdx = idxFor(map.date);
    var amountIdx = idxFor(map.amount);
    var descriptionIdx = idxFor(map.description);
    var campaignIdx = campaignColumn ? idxFor(campaignColumn) : -1;
    if (donorIdx < 0 || dateIdx < 0 || amountIdx < 0) {
      return { success: false, error: 'Column mapping incomplete. Donor, Date, and Amount are required.' };
    }

    var batchId = _donationBatchId_();
    var rootFolder = _donationImportsRootFolder_();
    var subfolder = rootFolder.createFolder(batchId);
    var file = subfolder.createFile(blob);
    var fileHash = _donationDigest_(text);

    var ledgerRows = _donationLedgerRecords_();
    var staging = [];
    var minDate = '';
    var maxDate = '';
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      var donorRaw = String(row[donorIdx] || '').trim();
      var donorNorm = _donationNormalizeDonor_(donorRaw);
      var txDate = _donationIsoDate_(row[dateIdx]);
      var amount = _donationAmountBrl_(row[amountIdx]);
      var descriptionRaw = descriptionIdx >= 0 ? String(row[descriptionIdx] || '').trim() : '';
      var campaignRaw = campaignIdx >= 0 ? String(row[campaignIdx] || '').trim() : '';
      var effectiveFundId = (campaignRaw && campaignMap[campaignRaw]) ? campaignMap[campaignRaw] : fundId;
      if (txDate) {
        if (!minDate || txDate < minDate) minDate = txDate;
        if (!maxDate || txDate > maxDate) maxDate = txDate;
      }

      var strictKey = _donationDigest_([effectiveFundId, donorNorm, txDate, isNaN(amount) ? '' : amount.toFixed(2)].join('|'));
      var fuzzyKey = _donationDigest_([effectiveFundId, donorNorm, isNaN(amount) ? '' : amount.toFixed(2)].join('|'));
      var matchStatus = 'NEW';
      var matchedDonationId = '';
      var matchConfidence = 0;
      var reviewDecision = '';
      var notes = '';

      if (!donorNorm || !txDate || isNaN(amount)) {
        matchStatus = 'INVALID';
        notes = 'Missing donor, date, or amount';
      } else if (!effectiveFundId) {
        matchStatus = 'INVALID';
        notes = 'No fund assigned for campaign: ' + (campaignRaw || '(blank)');
      } else {
        var exact = ledgerRows.find(function(item) {
          return item.fundId === effectiveFundId && item.donorNorm === donorNorm && item.txDate === txDate && Math.abs(item.amount - amount) < 0.001;
        });
        if (exact) {
          matchStatus = 'LIKELY_DUPLICATE';
          matchedDonationId = exact.donationId;
          matchConfidence = 100;
        } else {
          var fuzzy = ledgerRows.find(function(item) {
            if (item.fundId !== effectiveFundId) return false;
            if (item.donorNorm !== donorNorm) return false;
            if (Math.abs(item.amount - amount) > 0.001) return false;
            if (!item.txDate || !txDate) return false;
            var diffMs = Math.abs(new Date(item.txDate).getTime() - new Date(txDate).getTime());
            return diffMs <= (3 * 24 * 60 * 60 * 1000);
          });
          if (fuzzy) {
            matchStatus = 'POSSIBLE_DUPLICATE';
            matchedDonationId = fuzzy.donationId;
            matchConfidence = 70;
          }
        }
      }

      staging.push([
        batchId,
        i,
        donorRaw,
        donorNorm,
        txDate,
        isNaN(amount) ? '' : amount,
        effectiveFundId,
        descriptionRaw,
        strictKey,
        fuzzyKey,
        matchStatus,
        matchedDonationId,
        matchConfidence,
        reviewDecision,
        '',
        '',
        notes
      ]);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var batchSheet = getRequiredSheet_(ss, APP_SHEETS.DONATION_IMPORT_BATCHES, 'stageDonationImport');
    var stageSheet = getRequiredSheet_(ss, APP_SHEETS.DONATION_STAGING, 'stageDonationImport');
    batchSheet.appendRow([
      batchId,
      file.getName(),
      'CSV',
      file.getId(),
      file.getUrl(),
      fileHash,
      new Date().toISOString(),
      importedBy,
      minDate,
      maxDate,
      staging.length,
      'UNDER_REVIEW',
      ''
    ]);
    if (staging.length) {
      stageSheet.getRange(stageSheet.getLastRow() + 1, 1, staging.length, staging[0].length).setValues(staging);
    }

    return { success: true, batchId: batchId, rowCount: staging.length, fileUrl: file.getUrl() };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getDonationBatchReview(batchId) {
  try {
    var target = String(batchId || '').trim();
    if (!target) return { success: false, error: 'batchId required' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var batchSheet = getRequiredSheet_(ss, APP_SHEETS.DONATION_IMPORT_BATCHES, 'getDonationBatchReview');
    var stageSheet = getRequiredSheet_(ss, APP_SHEETS.DONATION_STAGING, 'getDonationBatchReview');

    var batchData = batchSheet.getDataRange().getValues();
    var stageData = stageSheet.getDataRange().getValues();
    var batchHeaders = batchData[0].map(_donationNormHeader_);
    var stageHeaders = stageData[0].map(_donationNormHeader_);

    var batchRow = null;
    for (var i = 1; i < batchData.length; i++) {
      if (String(batchData[i][batchHeaders.indexOf('BATCH_ID')] || '').trim() === target) { batchRow = batchData[i]; break; }
    }
    if (!batchRow) return { success: false, error: 'Batch not found: ' + target };

    function rowToObj(headers, row) {
      var obj = {};
      headers.forEach(function(h, idx) { obj[h] = row[idx]; });
      return obj;
    }
    var batch = rowToObj(batchHeaders, batchRow);
    var rows = stageData.slice(1)
      .filter(function(row) { return String(row[stageHeaders.indexOf('BATCH_ID')] || '').trim() === target; })
      .map(function(row) {
        var obj = rowToObj(stageHeaders, row);
        return {
          rowNo: Number(obj.ROW_NO || 0),
          donorRaw: String(obj.DONOR_RAW || ''),
          donorNormalized: String(obj.DONOR_NORMALIZED || ''),
          txDate: String(obj.TX_DATE || ''),
          amountBrl: obj.AMOUNT_BRL,
          fundId: String(obj.FUND_ID || ''),
          descriptionRaw: String(obj.DESCRIPTION_RAW || ''),
          matchStatus: String(obj.MATCH_STATUS || ''),
          matchedDonationId: String(obj.MATCHED_DONATION_ID || ''),
          matchConfidence: Number(obj.MATCH_CONFIDENCE || 0),
          reviewDecision: String(obj.REVIEW_DECISION || ''),
          notes: String(obj.NOTES || '')
        };
      })
      .sort(function(a, b) { return a.rowNo - b.rowNo; });

    var unresolved = rows.filter(function(r) { return !String(r.reviewDecision || '').trim(); }).length;
    var counts = { NEW: 0, POSSIBLE_DUPLICATE: 0, LIKELY_DUPLICATE: 0, INVALID: 0 };
    rows.forEach(function(r) { counts[r.matchStatus] = (counts[r.matchStatus] || 0) + 1; });

    var allFundIds = rows.map(function(r) { return r.fundId; }).filter(function(f, i, arr) { return f && arr.indexOf(f) === i; });
    var fundSummary = allFundIds.length === 0 ? '' : allFundIds.length === 1 ? allFundIds[0] : allFundIds.join(' | ');

    return {
      success: true,
      batch: {
        batchId: String(batch.BATCH_ID || ''),
        fileName: String(batch.SOURCE_FILENAME || ''),
        fileUrl: String(batch.SOURCE_FILE_URL || ''),
        fundId: fundSummary,
        status: String(batch.STATUS || ''),
        rowCount: Number(batch.ROW_COUNT || rows.length || 0),
        unresolvedCount: unresolved,
        counts: counts
      },
      rows: rows
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function reviewDonationStagingRow(batchId, rowNo, decision) {
  try {
    var target = String(batchId || '').trim();
    var rowNumber = Number(rowNo || 0);
    var nextDecision = String(decision || '').trim().toUpperCase();
    if (!target || !rowNumber) return { success: false, error: 'batchId and rowNo required' };
    if (['COMMIT', 'DUPLICATE', 'IGNORE'].indexOf(nextDecision) === -1) return { success: false, error: 'Invalid decision' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var stageSheet = getRequiredSheet_(ss, APP_SHEETS.DONATION_STAGING, 'reviewDonationStagingRow');
    var data = stageSheet.getDataRange().getValues();
    var headers = data[0].map(_donationNormHeader_);
    var batchIdx = headers.indexOf('BATCH_ID');
    var rowIdx = headers.indexOf('ROW_NO');
    var decisionIdx = headers.indexOf('REVIEW_DECISION');
    var reviewedByIdx = headers.indexOf('REVIEWED_BY');
    var reviewedAtIdx = headers.indexOf('REVIEWED_AT');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][batchIdx] || '').trim() === target && Number(data[i][rowIdx] || 0) === rowNumber) {
        stageSheet.getRange(i + 1, decisionIdx + 1).setValue(nextDecision);
        if (reviewedByIdx >= 0) stageSheet.getRange(i + 1, reviewedByIdx + 1).setValue(Session.getActiveUser().getEmail() || 'Admin');
        if (reviewedAtIdx >= 0) stageSheet.getRange(i + 1, reviewedAtIdx + 1).setValue(new Date().toISOString());
        return { success: true };
      }
    }
    return { success: false, error: 'Staging row not found' };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function commitDonationBatch(batchId) {
  try {
    var target = String(batchId || '').trim();
    if (!target) return { success: false, error: 'batchId required' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var batchSheet = getRequiredSheet_(ss, APP_SHEETS.DONATION_IMPORT_BATCHES, 'commitDonationBatch');
    var stageSheet = getRequiredSheet_(ss, APP_SHEETS.DONATION_STAGING, 'commitDonationBatch');
    var ledgerSheet = getRequiredSheet_(ss, APP_SHEETS.DONATIONS_LEDGER, 'commitDonationBatch');
    var fundLedgerSheet = getRequiredSheet_(ss, APP_SHEETS.FUND_LEDGER, 'commitDonationBatch');

    var stageData = stageSheet.getDataRange().getValues();
    var headers = stageData[0].map(_donationNormHeader_);
    var rows = stageData.slice(1).filter(function(row) { return String(row[headers.indexOf('BATCH_ID')] || '').trim() === target; });
    if (!rows.length) return { success: false, error: 'No staging rows found for batch ' + target };

    var unresolved = rows.filter(function(row) { return !String(row[headers.indexOf('REVIEW_DECISION')] || '').trim(); });
    if (unresolved.length) return { success: false, error: 'Batch is blocked until all rows are reviewed' };

    var committed = 0;
    var ignored = 0;
    var duplicates = 0;
    rows.forEach(function(row) {
      var decision = String(row[headers.indexOf('REVIEW_DECISION')] || '').trim().toUpperCase();
      if (decision === 'COMMIT') {
        var donationId = 'DON_' + Utilities.getUuid().slice(0, 8).toUpperCase();
        var txDate = row[headers.indexOf('TX_DATE')];
        var donorNorm = row[headers.indexOf('DONOR_NORMALIZED')];
        var donorRaw = row[headers.indexOf('DONOR_RAW')];
        var amount = row[headers.indexOf('AMOUNT_BRL')];
        var fundId = row[headers.indexOf('FUND_ID')];
        var sourceRowNo = row[headers.indexOf('ROW_NO')];
        var descriptionRaw = row[headers.indexOf('DESCRIPTION_RAW')];
        var strict = row[headers.indexOf('FINGERPRINT_STRICT')];
        var createdAt = new Date().toISOString();
        var createdBy = Session.getActiveUser().getEmail() || 'Admin';

        ledgerSheet.appendRow([
          donationId,
          txDate,
          donorNorm,
          donorRaw,
          amount,
          fundId,
          target,
          sourceRowNo,
          '',
          descriptionRaw,
          strict,
          createdAt,
          createdBy,
          'AVAILABLE',
          '',
          '',
          ''
        ]);
        fundLedgerSheet.appendRow([
          'LED_' + Utilities.getUuid().slice(0, 8).toUpperCase(),
          fundId,
          txDate,
          'DONATION_IN',
          amount,
          donationId,
          'DONATION',
          target,
          '',
          donationId,
          donorRaw,
          createdAt,
          createdBy
        ]);
        committed++;
      } else if (decision === 'DUPLICATE') {
        duplicates++;
      } else {
        ignored++;
      }
    });

    var batchData = batchSheet.getDataRange().getValues();
    var batchHeaders = batchData[0].map(_donationNormHeader_);
    var batchIdx = batchHeaders.indexOf('BATCH_ID');
    var statusIdx = batchHeaders.indexOf('STATUS');
    var notesIdx = batchHeaders.indexOf('NOTES');
    for (var i = 1; i < batchData.length; i++) {
      if (String(batchData[i][batchIdx] || '').trim() === target) {
        if (statusIdx >= 0) batchSheet.getRange(i + 1, statusIdx + 1).setValue('COMMITTED');
        if (notesIdx >= 0) batchSheet.getRange(i + 1, notesIdx + 1).setValue('Committed: ' + committed + ', Duplicate: ' + duplicates + ', Ignored: ' + ignored);
        break;
      }
    }

    return { success: true, committed: committed, duplicates: duplicates, ignored: ignored };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function uploadRunwaySurveyPhoto(payload) {
  try {
    var icao = String(payload && payload.icao || '').trim().toUpperCase();
    var rwyIdent = String(payload && payload.rwyIdent || '').trim().toUpperCase();
    var base64Data = String(payload && payload.base64Data || '').trim();
    var mimeType = String(payload && payload.mimeType || 'image/jpeg').trim() || 'image/jpeg';
    var source = String(payload && payload.source || '').trim();
    var fileName = String(payload && payload.fileName || '').trim();
    var takenAt = String(payload && payload.takenAt || new Date().toISOString()).trim();

    if (!icao) return { success: false, error: 'ICAO required' };
    if (!base64Data) return { success: false, error: 'Photo data is empty' };

    var folder = _findAirportPhotoFolder_(icao);
    if (!folder) return { success: false, error: 'No Drive folder found for airport ' + icao };

    var safeName = fileName || (icao + '_' + rwyIdent.replace(/\//g, '-') + '_' + new Date().getTime() + '.jpg');
    safeName = safeName.replace(/[^a-zA-Z0-9._-]+/g, '_');
    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, mimeType, safeName);
    var file = folder.createFile(blob);
    var desc = {
      kind: 'runway_survey_photo',
      icao: icao,
      rwyIdent: rwyIdent,
      source: source,
      takenAt: takenAt,
      uploadedAt: new Date().toISOString()
    };
    try {
      file.setDescription(JSON.stringify(desc));
    } catch (e) {}

    return {
      success: true,
      icao: icao,
      rwyIdent: rwyIdent,
      fileId: file.getId(),
      fileName: file.getName(),
      url: file.getUrl(),
      folderId: folder.getId(),
      folderUrl: folder.getUrl()
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function submitRunwaySurvey(payload) {
  try {
    const schema = _ensureRunwayWalkthroughLogSchema_();
    const logSheet = schema.sheet;
    const idx = schema.idx;

    const icao = String(payload && payload.icao || '').trim().toUpperCase();
    const rwyIdent = String(payload && payload.rwyIdent || '').trim().toUpperCase();
    if (!icao || !rwyIdent) return { success: false, error: 'ICAO and runway identifier required' };

    const stagingId = 'SURV_' + new Date().getTime() + '_' + icao + '_' + rwyIdent.replace(/\s+/g, '');
    const nowIso = new Date().toISOString();

    const row = new Array(schema.headers.length).fill('');
    row[idx.STAGING_ID] = stagingId;
    row[idx.ICAO] = icao;
    row[idx.RWY_IDENT] = rwyIdent;
    row[idx.PILOT_NAME] = String(payload && payload.pilotName || '').trim() || 'Unknown Pilot';
    row[idx.PILOT_EMAIL] = String(payload && payload.pilotEmail || '').trim();
    row[idx.WALK_DATE] = nowIso;
    row[idx.NOTES] = String(payload && payload.notes || '').trim();
    row[idx.FEATURES_JSON] = JSON.stringify(Array.isArray(payload && payload.features) ? payload.features : []);
    row[idx.STATUS] = 'PENDING';
    row[idx.ENTRY_KIND] = 'GPS_SURVEY';
    row[idx.SURVEY_JSON] = JSON.stringify(payload && payload.survey ? payload.survey : {});
    row[idx.OFFICIAL_JSON] = JSON.stringify(payload && payload.official ? payload.official : {});
    row[idx.CAPTURE_SUMMARY_JSON] = JSON.stringify(payload && payload.captureSummary ? payload.captureSummary : {});
    row[idx.DEVICE_INFO_JSON] = JSON.stringify(payload && payload.deviceInfo ? payload.deviceInfo : {});

    logSheet.appendRow(row);
    return { success: true, stagingId: stagingId, message: 'Runway survey submitted for supervisor review' };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function submitRunwayApprovalRequest(payload) {
  try {
    const schema = _ensureRunwayWalkthroughLogSchema_();
    const logSheet = schema.sheet;
    const idx = schema.idx;

    const icao = String(payload && payload.icao || '').trim().toUpperCase();
    const rwyIdent = String(payload && payload.rwyIdent || '').trim().toUpperCase() || 'UNKNOWN';
    if (!icao) return { success: false, error: 'ICAO required' };

    const stagingId = 'APPROVAL_' + new Date().getTime() + '_' + icao + '_' + rwyIdent.replace(/\s+/g, '');
    const nowIso = new Date().toISOString();

    const row = new Array(schema.headers.length).fill('');
    row[idx.STAGING_ID] = stagingId;
    row[idx.ICAO] = icao;
    row[idx.RWY_IDENT] = rwyIdent;
    row[idx.PILOT_NAME] = String(payload && payload.pilotName || '').trim() || 'Unknown Pilot';
    row[idx.PILOT_EMAIL] = String(payload && payload.pilotEmail || '').trim();
    row[idx.WALK_DATE] = nowIso;
    row[idx.NOTES] = String(payload && payload.notes || '').trim();
    row[idx.FEATURES_JSON] = JSON.stringify([]);
    row[idx.STATUS] = 'PENDING';
    row[idx.ENTRY_KIND] = 'RUNWAY_APPROVAL';
    row[idx.SURVEY_JSON] = JSON.stringify(payload && payload.survey ? payload.survey : {});
    row[idx.OFFICIAL_JSON] = JSON.stringify(payload && payload.official ? payload.official : {});
    row[idx.CAPTURE_SUMMARY_JSON] = JSON.stringify({ source: 'TAB5_RELEASE' });
    row[idx.DEVICE_INFO_JSON] = JSON.stringify({ source: 'pilot_app_release_tab' });

    logSheet.appendRow(row);
    return { success: true, stagingId: stagingId, message: 'Runway approval request submitted for supervisor review' };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getPendingRunwaySurveyReviews(limit) {
  try {
    const schema = _ensureRunwayWalkthroughLogSchema_();
    const sh = schema.sheet;
    const idx = schema.idx;
    const max = Math.max(1, Number(limit || 100) || 100);

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: true, items: [] };
    const rows = sh.getRange(2, 1, lastRow - 1, schema.headers.length).getValues();

    const items = rows
      .map(function(r, i) {
        return {
          rowNum: i + 2,
          stagingId: String(r[idx.STAGING_ID] || '').trim(),
          icao: String(r[idx.ICAO] || '').trim().toUpperCase(),
          rwyIdent: String(r[idx.RWY_IDENT] || '').trim().toUpperCase(),
          pilotName: String(r[idx.PILOT_NAME] || '').trim(),
          pilotEmail: String(r[idx.PILOT_EMAIL] || '').trim(),
          walkDate: String(r[idx.WALK_DATE] || '').trim(),
          notes: String(r[idx.NOTES] || '').trim(),
          status: String(r[idx.STATUS] || '').trim().toUpperCase(),
          entryKind: String(r[idx.ENTRY_KIND] || '').trim().toUpperCase(),
          survey: _parseJsonLoose_(r[idx.SURVEY_JSON], {}),
          official: _parseJsonLoose_(r[idx.OFFICIAL_JSON], {}),
          captureSummary: _parseJsonLoose_(r[idx.CAPTURE_SUMMARY_JSON], {}),
          deviceInfo: _parseJsonLoose_(r[idx.DEVICE_INFO_JSON], {})
        };
      })
      .filter(function(item) {
        const s = item.status || '';
        return s === 'PENDING' || s === 'QUEUED';
      })
      .sort(function(a, b) {
        return String(b.walkDate || '').localeCompare(String(a.walkDate || ''));
      })
      .slice(0, max);

    return { success: true, items: items };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function approveRunwaySurveyReview(stagingId, supervisorName, supervisorNotes, approvalPassword) {
  try {
    _verifySupervisorApprovalPassword_(approvalPassword);
    const id = String(stagingId || '').trim();
    if (!id) return { success: false, error: 'Missing stagingId' };

    const schema = _ensureRunwayWalkthroughLogSchema_();
    const logSheet = schema.sheet;
    const idx = schema.idx;
    const rows = logSheet.getDataRange().getValues();
    let logRow = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idx.STAGING_ID] || '').trim() === id) {
        logRow = i;
        break;
      }
    }
    if (logRow < 0) return { success: false, error: 'Survey staging record not found' };

    const icao = String(rows[logRow][idx.ICAO] || '').trim().toUpperCase();
    const rwyIdent = String(rows[logRow][idx.RWY_IDENT] || '').trim().toUpperCase();
    const entryKind = String(rows[logRow][idx.ENTRY_KIND] || '').trim().toUpperCase();
    const survey = _parseJsonLoose_(rows[logRow][idx.SURVEY_JSON], {});
    const nowIso = new Date().toISOString();

    if (entryKind === 'RUNWAY_APPROVAL') {
      logSheet.getRange(logRow + 1, idx.STATUS + 1).setValue('PUBLISHED');
      logSheet.getRange(logRow + 1, idx.SUPERVISOR_NAME + 1).setValue(String(supervisorName || '').trim() || 'Supervisor');
      logSheet.getRange(logRow + 1, idx.SUPERVISOR_NOTES + 1).setValue(String(supervisorNotes || '').trim());
      logSheet.getRange(logRow + 1, idx.APPROVED_AT + 1).setValue(nowIso);
      logSheet.getRange(logRow + 1, idx.PUBLISHED_AT + 1).setValue(nowIso);
      SpreadsheetApp.flush();
      return { success: true, message: 'Runway approval request approved', icao: icao, rwyIdent: rwyIdent };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName('DB_Airports');
    if (!dbSheet) return { success: false, error: 'DB_Airports sheet not found' };

    const dbData = dbSheet.getDataRange().getValues();
    if (!dbData.length) return { success: false, error: 'DB_Airports is empty' };
    const cols = _runwayDbFindCols_(dbData[0]);
    if (cols.icao < 0 || cols.runway < 0 || cols.knownFeatures < 0) {
      return { success: false, error: 'DB_Airports missing ICAO/RWY_IDENT/KNOWN_FEATURES columns' };
    }

    let dbRow = -1;
    for (let i = 1; i < dbData.length; i++) {
      const rowIcao = String(dbData[i][cols.icao] || '').trim().toUpperCase();
      const rowRwy = String(dbData[i][cols.runway] || '').trim().toUpperCase();
      if (rowIcao === icao && rowRwy === rwyIdent) {
        dbRow = i;
        break;
      }
    }
    if (dbRow < 0) return { success: false, error: 'Runway not found in DB_Airports' };

    const pairKey = _runwayPairKey_(rwyIdent);
    const publishRows = [];
    for (let j = 1; j < dbData.length; j++) {
      const rowIcao = String(dbData[j][cols.icao] || '').trim().toUpperCase();
      const rowRwy = String(dbData[j][cols.runway] || '').trim().toUpperCase();
      if (rowIcao !== icao) continue;
      if ((pairKey && _runwayPairKey_(rowRwy) === pairKey) || rowRwy === rwyIdent) {
        publishRows.push(j);
      }
    }
    if (!publishRows.length) publishRows.push(dbRow);

    publishRows.forEach(function(rowIndex) {
      const rowRwyIdent = String(dbData[rowIndex][cols.runway] || '').trim().toUpperCase();
      const officialSnapshot = _dbAirportOfficialSnapshot_(dbData[rowIndex], cols);
      const surveyForRow = _transformSurveyForRunway_(survey, rwyIdent, rowRwyIdent, officialSnapshot.lengthM || Number(survey && survey.lengthM || 0));
      const existingRaw = String(dbData[rowIndex][cols.knownFeatures] || '').trim();
      const existingObj = _parseJsonLoose_(existingRaw, {});
      const normalizedExisting = Array.isArray(existingObj) ? { features: existingObj } : (existingObj || {});

      const verifiedOperational = {
        lengthM: Number(surveyForRow && surveyForRow.lengthM || 0) || officialSnapshot.lengthM || 0,
        widthM: Number(surveyForRow && surveyForRow.widthM || 0) || officialSnapshot.widthM || 0,
        surface: String(surveyForRow && surveyForRow.surface || '').trim() || officialSnapshot.surface || '',
        slopeFromThreshold: String(surveyForRow && surveyForRow.slopeFromThreshold || rowRwyIdent).trim(),
        features: Array.isArray(surveyForRow && surveyForRow.features) ? surveyForRow.features : [],
        markers: Array.isArray(surveyForRow && surveyForRow.markers) ? surveyForRow.markers : [],
        obstacles: Array.isArray(surveyForRow && surveyForRow.obstacles) ? surveyForRow.obstacles : [],
        obstacleAngles50m: Array.isArray(surveyForRow && surveyForRow.obstacleAngles50m) ? surveyForRow.obstacleAngles50m : [],
        slopeSegments: Array.isArray(surveyForRow && surveyForRow.slopeSegments) ? surveyForRow.slopeSegments : [],
        perimeterTrace: Array.isArray(surveyForRow && surveyForRow.perimeterTrace) ? surveyForRow.perimeterTrace : [],
        axis: surveyForRow && surveyForRow.axis ? surveyForRow.axis : {},
        thresholds: surveyForRow && surveyForRow.thresholds ? surveyForRow.thresholds : {},
        gpsSummary: surveyForRow && surveyForRow.gpsSummary ? surveyForRow.gpsSummary : {},
        cutdownAreas: surveyForRow && surveyForRow.cutdownAreas ? surveyForRow.cutdownAreas : {},
        cutdownAreaM: Number(surveyForRow && surveyForRow.cutdownAreaM || 0) || null,
        widthObservations: Array.isArray(surveyForRow && surveyForRow.widthObservations) ? surveyForRow.widthObservations : [],
        perimeterSummary: surveyForRow && surveyForRow.perimeterSummary ? surveyForRow.perimeterSummary : {},
        pilotNotes: String(surveyForRow && surveyForRow.notes || '').trim(),
        approvedBy: String(supervisorName || '').trim() || 'Supervisor',
        approvedAt: nowIso
      };

      const merged = Object.assign({}, normalizedExisting, {
        features: verifiedOperational.features,
        verifiedOperational: verifiedOperational,
        verifiedSurvey: {
          version: 2,
          capturedAt: String(rows[logRow][idx.WALK_DATE] || nowIso),
          pilotName: String(rows[logRow][idx.PILOT_NAME] || '').trim(),
          pilotEmail: String(rows[logRow][idx.PILOT_EMAIL] || '').trim(),
          status: 'APPROVED',
          stagingId: id,
          publishedRunway: rowRwyIdent,
          sourceRunway: rwyIdent
        },
        officialReference: {
          lengthM: officialSnapshot.lengthM,
          widthM: officialSnapshot.widthM,
          surface: officialSnapshot.surface,
          headingDeg: officialSnapshot.headingDeg,
          source: 'DB_Airports',
          capturedAt: nowIso
        },
        updatedAt: nowIso
      });

      if (verifiedOperational.slopeSegments.length) {
        merged.slopeProfile = verifiedOperational.slopeSegments.map(function(seg) {
          return { distance: Number(seg.distanceM || 0) || 0, slope: Number(seg.slope || 0) || 0 };
        });
      }

      dbSheet.getRange(rowIndex + 1, cols.knownFeatures + 1).setValue(JSON.stringify(merged));
    });

    logSheet.getRange(logRow + 1, idx.STATUS + 1).setValue('PUBLISHED');
    logSheet.getRange(logRow + 1, idx.SUPERVISOR_NAME + 1).setValue(String(supervisorName || '').trim() || 'Supervisor');
    logSheet.getRange(logRow + 1, idx.SUPERVISOR_NOTES + 1).setValue(String(supervisorNotes || '').trim());
    logSheet.getRange(logRow + 1, idx.APPROVED_AT + 1).setValue(nowIso);
    logSheet.getRange(logRow + 1, idx.PUBLISHED_AT + 1).setValue(nowIso);
    SpreadsheetApp.flush();

    return { success: true, message: 'Runway survey approved and published', icao: icao, rwyIdent: rwyIdent, headingsUpdated: publishRows.length };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function rejectRunwaySurveyReview(stagingId, supervisorName, supervisorNotes, approvalPassword) {
  try {
    _verifySupervisorApprovalPassword_(approvalPassword);
    const id = String(stagingId || '').trim();
    if (!id) return { success: false, error: 'Missing stagingId' };
    const schema = _ensureRunwayWalkthroughLogSchema_();
    const sh = schema.sheet;
    const idx = schema.idx;
    const rows = sh.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idx.STAGING_ID] || '').trim() === id) {
        rowIndex = i;
        break;
      }
    }
    if (rowIndex < 0) return { success: false, error: 'Survey staging record not found' };

    const nowIso = new Date().toISOString();
    sh.getRange(rowIndex + 1, idx.STATUS + 1).setValue('REJECTED');
    sh.getRange(rowIndex + 1, idx.SUPERVISOR_NAME + 1).setValue(String(supervisorName || '').trim() || 'Supervisor');
    sh.getRange(rowIndex + 1, idx.SUPERVISOR_NOTES + 1).setValue(String(supervisorNotes || '').trim());
    sh.getRange(rowIndex + 1, idx.APPROVED_AT + 1).setValue(nowIso);
    return { success: true, message: 'Runway survey rejected' };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _aircraftDocsHeaders_() {
  return [
    'TAIL',
    'DOC_TYPE',
    'DOC_NAME',
    'DRIVE_URL',
    'DRIVE_FILE_ID',
    'REQUIRED',
    'CRITICAL',
    'REVISION',
    'EFFECTIVE_DATE',
    'LAST_VERIFIED_OFFLINE',
    'NOTES',
    'UPDATED_AT',
    'UPDATED_BY'
  ];
}

function _ensureAircraftDocsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = APP_SHEETS.AIRCRAFT_DOCS || 'DB_Aircraft_Docs';
  var sh = ss.getSheetByName(sheetName);
  var required = _aircraftDocsHeaders_();

  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1, 1, 1, required.length).setValues([required]);
  }

  var lastCol = Math.max(sh.getLastColumn(), 1);
  var existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || '').trim();
  });

  required.forEach(function(name) {
    if (existing.indexOf(name) < 0) {
      existing.push(name);
      sh.getRange(1, existing.length).setValue(name);
    }
  });

  var norm = existing.map(function(h) { return _toolsNormHeader_(h); });
  var idx = {};
  existing.forEach(function(h, i) { idx[_toolsNormHeader_(h)] = i; });

  return { sheet: sh, headers: existing, norm: norm, idx: idx };
}

function _aircraftDocsNormalizeTail_(value) {
  return String(value || '').trim().toUpperCase();
}

function _aircraftDocsFlag_(value) {
  var raw = String(value == null ? '' : value).trim().toUpperCase();
  return (raw === 'Y' || raw === 'YES' || raw === 'TRUE' || raw === '1') ? 'Y' : 'N';
}

function _aircraftDocsDriveIdFromUrl_(url) {
  var raw = String(url || '').trim();
  if (!raw) return '';
  var m = raw.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
  if (m && m[1]) return m[1];
  m = raw.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if (m && m[1]) return m[1];
  return '';
}

function _aircraftDocsFolderUrlForTail_(tail) {
  var code = _aircraftDocsNormalizeTail_(tail);
  if (!code) return '';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(APP_SHEETS.AIRCRAFT);
  if (!sh) return '';
  var data = sh.getDataRange().getValues();
  if (!data || data.length < 2) return '';

  var headers = data[0].map(function(h) { return _toolsNormHeader_(h); });
  var regIdx = headers.indexOf('REGISTRATION');
  if (regIdx < 0) regIdx = headers.indexOf('AIRCRAFT_REGISTRATION');
  if (regIdx < 0) regIdx = headers.indexOf('TAIL');
  if (regIdx < 0) regIdx = headers.indexOf('AIRCRAFT');
  if (regIdx < 0) return '';

  var folderCandidates = [
    'DOCUMENTS_FOLDER_URL',
    'DRIVE_FOLDER_URL',
    'DOC_FOLDER_URL',
    'AIRCRAFT_DOCS_URL',
    'AIRCRAFT_DOCS_FOLDER',
    'GOOGLE_DRIVE_FOLDER',
    'DRIVE_URL'
  ].map(function(h) { return _toolsNormHeader_(h); });

  var folderIdx = -1;
  for (var ci = 0; ci < folderCandidates.length; ci++) {
    var at = headers.indexOf(folderCandidates[ci]);
    if (at >= 0) {
      folderIdx = at;
      break;
    }
  }
  if (folderIdx < 0) return '';

  for (var i = 1; i < data.length; i++) {
    if (_aircraftDocsNormalizeTail_(data[i][regIdx]) !== code) continue;
    var url = String(data[i][folderIdx] || '').trim();
    if (url) return url;
  }
  return '';
}

function _aircraftDocsListAircraftRegs_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(APP_SHEETS.AIRCRAFT);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (!data || data.length < 2) return [];

  var headers = data[0].map(function(h) { return _toolsNormHeader_(h); });
  var regIdx = headers.indexOf('REGISTRATION');
  if (regIdx < 0) regIdx = headers.indexOf('AIRCRAFT_REGISTRATION');
  if (regIdx < 0) regIdx = headers.indexOf('TAIL');
  if (regIdx < 0) regIdx = headers.indexOf('AIRCRAFT');
  if (regIdx < 0) return [];

  var out = {};
  for (var i = 1; i < data.length; i++) {
    var reg = _aircraftDocsNormalizeTail_(data[i][regIdx]);
    if (reg) out[reg] = true;
  }
  return Object.keys(out).sort();
}

function getAircraftDocsForTools(tail) {
  try {
    var schema = _ensureAircraftDocsSheet_();
    var sh = schema.sheet;
    var idx = schema.idx;
    var rows = sh.getDataRange().getValues();
    var targetTail = _aircraftDocsNormalizeTail_(tail);
    var docs = [];
    var lastVerified = '';

    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      var rowTail = _aircraftDocsNormalizeTail_(row[idx.TAIL]);
      if (targetTail && rowTail !== targetTail) continue;
      if (!rowTail) continue;

      var verified = String(row[idx.LAST_VERIFIED_OFFLINE] || '').trim();
      if (verified && (!lastVerified || String(verified) > String(lastVerified))) lastVerified = verified;

      docs.push({
        rowNumber: i + 1,
        tail: rowTail,
        docType: String(row[idx.DOC_TYPE] || '').trim(),
        docName: String(row[idx.DOC_NAME] || '').trim(),
        driveUrl: String(row[idx.DRIVE_URL] || '').trim(),
        driveFileId: String(row[idx.DRIVE_FILE_ID] || '').trim(),
        required: _aircraftDocsFlag_(row[idx.REQUIRED]) === 'Y',
        critical: _aircraftDocsFlag_(row[idx.CRITICAL]) === 'Y',
        revision: String(row[idx.REVISION] || '').trim(),
        effectiveDate: String(row[idx.EFFECTIVE_DATE] || '').trim(),
        lastVerifiedOffline: verified,
        notes: String(row[idx.NOTES] || '').trim(),
        updatedAt: String(row[idx.UPDATED_AT] || '').trim(),
        updatedBy: String(row[idx.UPDATED_BY] || '').trim()
      });
    }

    docs.sort(function(a, b) {
      var ak = (a.tail + '|' + a.docType + '|' + a.docName).toUpperCase();
      var bk = (b.tail + '|' + b.docType + '|' + b.docName).toUpperCase();
      return ak.localeCompare(bk);
    });

    return {
      success: true,
      tail: targetTail,
      aircraftRegs: _aircraftDocsListAircraftRegs_(),
      folderUrl: targetTail ? _aircraftDocsFolderUrlForTail_(targetTail) : '',
      lastVerifiedOffline: lastVerified,
      docs: docs
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function saveAircraftDocForTools(body) {
  try {
    var payload = (body && typeof body === 'object') ? body : {};
    var tail = _aircraftDocsNormalizeTail_(payload.tail);
    if (!tail) return { success: false, error: 'Aircraft tail/registration is required' };

    var docType = String(payload.docType || '').trim();
    var docName = String(payload.docName || '').trim();
    var driveUrl = String(payload.driveUrl || '').trim();
    if (!docType) return { success: false, error: 'Document type is required' };
    if (!docName) return { success: false, error: 'Document name is required' };
    if (!driveUrl) return { success: false, error: 'Drive URL is required' };

    var schema = _ensureAircraftDocsSheet_();
    var sh = schema.sheet;
    var idx = schema.idx;
    var rows = sh.getDataRange().getValues();
    var nowIso = new Date().toISOString();
    var by = String(payload.updatedBy || _schedulerCurrentUserEmail_() || 'tools').trim();

    var rowNumber = Number(payload.rowNumber || 0);
    var targetRow = -1;
    if (rowNumber >= 2 && rowNumber <= rows.length) {
      targetRow = rowNumber;
    } else {
      for (var i = 1; i < rows.length; i++) {
        var sameTail = _aircraftDocsNormalizeTail_(rows[i][idx.TAIL]) === tail;
        var sameType = String(rows[i][idx.DOC_TYPE] || '').trim().toUpperCase() === docType.toUpperCase();
        var sameName = String(rows[i][idx.DOC_NAME] || '').trim().toUpperCase() === docName.toUpperCase();
        if (sameTail && sameType && sameName) {
          targetRow = i + 1;
          break;
        }
      }
    }

    var rec = [];
    rec[idx.TAIL] = tail;
    rec[idx.DOC_TYPE] = docType;
    rec[idx.DOC_NAME] = docName;
    rec[idx.DRIVE_URL] = driveUrl;
    rec[idx.DRIVE_FILE_ID] = String(payload.driveFileId || '').trim() || _aircraftDocsDriveIdFromUrl_(driveUrl);
    rec[idx.REQUIRED] = _aircraftDocsFlag_(payload.required);
    rec[idx.CRITICAL] = _aircraftDocsFlag_(payload.critical);
    rec[idx.REVISION] = String(payload.revision || '').trim();
    rec[idx.EFFECTIVE_DATE] = String(payload.effectiveDate || '').trim();
    rec[idx.LAST_VERIFIED_OFFLINE] = String(payload.lastVerifiedOffline || '').trim();
    rec[idx.NOTES] = String(payload.notes || '').trim();
    rec[idx.UPDATED_AT] = nowIso;
    rec[idx.UPDATED_BY] = by;

    for (var c = 0; c < schema.headers.length; c++) {
      if (typeof rec[c] === 'undefined') rec[c] = '';
    }

    if (targetRow >= 2) {
      sh.getRange(targetRow, 1, 1, rec.length).setValues([rec]);
      return { success: true, rowNumber: targetRow, action: 'updated' };
    }

    sh.appendRow(rec);
    return { success: true, rowNumber: sh.getLastRow(), action: 'created' };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function markAircraftDocsOfflineVerifiedForTools(tail, verifiedBy) {
  try {
    var target = _aircraftDocsNormalizeTail_(tail);
    if (!target) return { success: false, error: 'Aircraft tail/registration is required' };

    var schema = _ensureAircraftDocsSheet_();
    var sh = schema.sheet;
    var idx = schema.idx;
    var rows = sh.getDataRange().getValues();
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT', 'yyyy-MM-dd');
    var nowIso = new Date().toISOString();
    var by = String(verifiedBy || _schedulerCurrentUserEmail_() || 'tools').trim();
    var updated = 0;

    for (var i = 1; i < rows.length; i++) {
      if (_aircraftDocsNormalizeTail_(rows[i][idx.TAIL]) !== target) continue;
      sh.getRange(i + 1, idx.LAST_VERIFIED_OFFLINE + 1).setValue(stamp);
      sh.getRange(i + 1, idx.UPDATED_AT + 1).setValue(nowIso);
      sh.getRange(i + 1, idx.UPDATED_BY + 1).setValue(by);
      updated++;
    }

    if (!updated) return { success: false, error: 'No document rows found for ' + target };
    return { success: true, tail: target, lastVerifiedOffline: stamp, updatedRows: updated };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}