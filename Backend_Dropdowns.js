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
return rows.map(function(r) {
  return _airportMapRow_(headers, r);
});
}
try {
const air = getSheetData(APP_SHEETS.AIRPORTS);
const acft = getSheetData(APP_SHEETS.AIRCRAFT);
const pax = getSheetData(APP_SHEETS.PASSENGERS);
const syl = getSheetData(APP_SHEETS.SYLLABUS);
const caches = getSheetData(APP_SHEETS.FUEL_CACHES);
const contacts = getSheetData(APP_SHEETS.CONTACTS);
const wpData = getSheetData(APP_SHEETS.WAYPOINTS);
const routeData = getSheetData(APP_SHEETS.ROUTES);
const pilots = getSheetData(APP_SHEETS.PILOTS);
const checks = getSheetData(APP_SHEETS.CHECKS);
const mxMaintenance = getDispatchMaintenanceStatus({});

// Diagnostics: log presence and header info for key sheets to help trace missing data
try {
  appLog_('getDropdownData: DB_Aircraft headers=', acft && acft.headers ? acft.headers : 'N/A', 'rows=', acft && acft.vals ? acft.vals.length : 0);
  appLog_('getDropdownData: DB_Fuel_Caches headers=', caches && caches.headers ? caches.headers : 'N/A', 'rows=', caches && caches.vals ? caches.vals.length : 0);
  appLog_('getDropdownData: DB_Contacts headers=', contacts && contacts.headers ? contacts.headers : 'N/A', 'rows=', contacts && contacts.vals ? contacts.vals.length : 0);
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
const pilotDestinationChecks = _collectPilotDestinationChecksMap_(checks.headers, checks.vals);

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
 fuelContacts: contacts.vals.map(function(r) {
   return {
     icao: String(_toolsHeaderValueByCandidates_(contacts.headers, r, ['ICAO','C_DIGO','CODIGO','CDIGO','C__DIGO'], '') || '').trim().toUpperCase(),
     providerName: String(_toolsHeaderValueByCandidates_(contacts.headers, r, ['CONTATO','NOME','CONTACT_NAME','FORNECEDOR','PROVIDER'], '') || '').trim(),
     celular: String(_toolsHeaderValueByCandidates_(contacts.headers, r, ['CELULAR','MOBILE','TELEFONE'], '') || '').trim()
   };
 }).filter(function(c) {
   return !!String(c.icao || '').trim();
 }),
 syllabus: syl.vals.map(r => ({
     code: r[syl.headers.indexOf("TRAINING_CODE")],
     description: (function() {
       var idx = syl.headers.indexOf('DESCRIPTION');
       return idx >= 0 ? String(r[idx] || '').trim() : '';
     })(),
     aircraftType: (function() {
       var idx = syl.headers.indexOf('AIRCRAFT_TYPE');
       return idx >= 0 ? String(r[idx] || '').trim() : '';
     })(),
     hours: parseFloat(r[syl.headers.indexOf("REQUIRED_HOURS")]) || 0,
     plannedLandings: (function() {
       var idx = syl.headers.indexOf('PLANNED_LANDINGS');
       return idx >= 0 ? (parseInt(r[idx], 10) || 0) : 0;
     })(),
     fuel: parseFloat(r[syl.headers.indexOf("REQUIRED_FUEL")]) || 0,
     requiredBallast: (function() {
       var idx = syl.headers.indexOf('REQUIRED_BALLAST');
       return idx >= 0 ? (parseFloat(r[idx]) || 0) : 0;
     })(),
     route: (function() {
       var routeIdx = syl.headers.indexOf('ROUTE');
       return routeIdx >= 0 ? String(r[routeIdx] || '').trim() : '';
     })(),
     runwayCheckout: (function() {
       var idx = syl.headers.indexOf('RUNWAY_CHECKOUT');
       return idx >= 0 ? String(r[idx] || '').trim() : '';
     })(),
     lessonPlanJson: (function() {
       var idx = syl.headers.indexOf('LESSON_PLAN_JSON');
       if (idx < 0) idx = syl.headers.indexOf('TRAINING_PLAN_JSON');
       if (idx < 0) idx = syl.headers.indexOf('PLAN_JSON');
       return idx >= 0 ? String(r[idx] || '').trim() : '';
     })(),
     plannedStopsJson: (function() {
       var idx = syl.headers.indexOf('PLANNED_STOPS_JSON');
       if (idx < 0) idx = syl.headers.indexOf('STOP_PLAN_JSON');
       return idx >= 0 ? String(r[idx] || '').trim() : '';
     })(),
     maneuversJson: (function() {
       var idx = syl.headers.indexOf('MANEUVERS_JSON');
       if (idx < 0) idx = syl.headers.indexOf('PLANNED_MANEUVERS_JSON');
       return idx >= 0 ? String(r[idx] || '').trim() : '';
     })()
 })).filter(s => s.code),
 waypoints: wpData.vals.map(r => ({
   wp_id: String(r[wpData.headers.indexOf("WP_ID")]),
   // Keep raw coordinates so clients can parse decimal comma / DMS formats.
   lat: String(r[wpData.headers.indexOf("LATITUDE")] == null ? '' : r[wpData.headers.indexOf("LATITUDE")]).trim(),
   lon: String(r[wpData.headers.indexOf("LONGITUDE")] == null ? '' : r[wpData.headers.indexOf("LONGITUDE")]).trim(),
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
   var reg = String(r[acft.headers.indexOf("REGISTRATION")] || '').trim().toUpperCase();
   var currentHoursToInsp = parseFloat(r[acft.headers.indexOf("HOURS_TO_INSPECTION")]) || 0;
   var maintenance = mxMaintenance && mxMaintenance.success && mxMaintenance.byAircraft ? mxMaintenance.byAircraft[reg] : null;
   var nextInspectionDate = maintenance && maintenance.nextInspectionDate ? String(maintenance.nextInspectionDate || '').trim() : '';
   var forecastHoursToInsp = _mxCurrentOrForecastHoursToInsp_({ reg: reg, hoursToInsp: currentHoursToInsp }, 3, nextInspectionDate);
  var aircraftTypeRaw = String(r[acft.headers.indexOf("AIRCRAFT_TYPE")] || '').trim();
  var typeForPerformanceRaw = String(r[acft.headers.indexOf("TYPE_FOR_PERFORMANCE")] || '').trim();
  var aircraftTypeCanon = String(aircraftTypeRaw || typeForPerformanceRaw || '').trim();
   return Object.assign({
     reg: reg,
     aircraftType: aircraftTypeCanon,
     aircraftTypeRaw: aircraftTypeRaw,
    typeForPerformance: aircraftTypeCanon,
    typeForPerformanceRaw: typeForPerformanceRaw,
    POSICAO_AMBOS: _toolsTruthyFlag_(r[acft.headers.indexOf("POSICAO_AMBOS")]),
     speed: parseFloat(r[acft.headers.indexOf("CRUISE_KTS")]) || 130,
     burn: parseFloat(r[acft.headers.indexOf("BURN_LPH")]) || 60,
     rate: parseFloat(r[acft.headers.indexOf("HOURLY_RATE")]) || 0,
     maxFuel: parseFloat(r[acft.headers.indexOf("MAX_FUEL")]) || 400,
     emptyWeight: parseFloat(r[acft.headers.indexOf("EMPTY_WEIGHT")]) || 1000,
     mtow: parseFloat(r[acft.headers.indexOf("MTOW")]) || 1600,
    mtowLand: _numByAliases(["MTOW_LAND", "MAX_TAKEOFF_WEIGHT_LAND", "TAKEOFF_WEIGHT_LAND", "MTOW_GROUND"], 0),
    mtowWater: _numByAliases(["MTOW_WATER", "MAX_TAKEOFF_WEIGHT_WATER", "TAKEOFF_WEIGHT_WATER", "MTOW_SEA"], 0),
    mlwLand: _numByAliases(["MLW_LAND", "MAX_LANDING_WEIGHT_LAND", "LANDING_WEIGHT_LAND"], 0),
    mlwWater: _numByAliases(["MLW_WATER", "MAX_LANDING_WEIGHT_WATER", "LANDING_WEIGHT_WATER", "MLW_SEA"], 0),
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
     hoursToInsp: currentHoursToInsp,
     hoursToInspForecast: isFinite(Number(forecastHoursToInsp && forecastHoursToInsp.plannedHoursToInsp)) ? Number(forecastHoursToInsp.plannedHoursToInsp) : '',
     plannedFlightHours3Months: isFinite(Number(forecastHoursToInsp && forecastHoursToInsp.plannedFlightHours)) ? Number(forecastHoursToInsp.plannedFlightHours) : 0,
    nextMaintenanceInspectionDate: nextInspectionDate,
     forecastSource: 'scheduled_flights_3_months',
     cvaExpiryDate: (function() {
       var idx = acft.headers.indexOf('ANNUAL_DUE_(CVA)');
       return idx >= 0 ? safeDateStr(r[idx]) : '';
     })(),
      techStatus: (r[acft.headers.indexOf("TECH_STATUS")] || "SERVICEABLE").toUpperCase(),
     openSquawks: r[acft.headers.indexOf("OPEN_SQUAWKS")] || ""
   }, vrCols);
 }),
 passengers: pax.vals.map(r => {
     const h = pax.headers;
     const weightIdx = h.indexOf("WEIGHT_KG") !== -1 ? h.indexOf("WEIGHT_KG") : h.indexOf("WEIGHT_KGS");
     const dobIdx = h.indexOf("DOB");
     const phoneIdx = h.indexOf("PHONE");
     const consentAtIdx = h.indexOf("TERMS_ACCEPTED_AT") !== -1
       ? h.indexOf("TERMS_ACCEPTED_AT")
       : (h.indexOf("AGREED_TO_TERMS_AT") !== -1 ? h.indexOf("AGREED_TO_TERMS_AT") : h.indexOf("TERMS_LAST_ACCEPTED_AT"));
     const consentFlagIdx = h.indexOf("TERMS_ACCEPTED");
     const consentAtRaw = consentAtIdx !== -1 ? r[consentAtIdx] : "";
     const consentAt = _passengerToIsoDate_(consentAtRaw);
     const consentCurrent = _passengerConsentCurrent_(consentAtRaw);
     const name = r[h.indexOf("PASSENGER_NAME")] || "Unknown";
     return {
       name: name,
       nameDisplay: consentCurrent ? (String(name) + ' ✓') : String(name),
       weight: parseFloat(r[weightIdx]) || 80,
       gender: r[h.indexOf("GENDER")] || "U",
       dob: safeDobStr(dobIdx !== -1 ? r[dobIdx] : ""),
       phone: phoneIdx !== -1 ? String(r[phoneIdx] || "") : "",
       termsAcceptedAt: consentAt,
       termsAccepted: consentFlagIdx !== -1 ? _schedulerTruthyFlag_(r[consentFlagIdx]) : !!consentAt,
       termsAcceptedCurrent: consentCurrent
     };
 }).filter(p => p.name && p.name !== "Unknown"),
 airports: mapAirportRows(air),
 passengerConsentConfig: _passengerConsentReadConfigMap_()
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
    passengers: data && Array.isArray(data.passengers) ? data.passengers : [],
    passengerConsentConfig: data && data.passengerConsentConfig ? data.passengerConsentConfig : _passengerConsentReadConfigMap_()
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
          mtowByModel: row.mtowByModel || {},
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
        mtowByModel: row.mtowByModel || {},
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
