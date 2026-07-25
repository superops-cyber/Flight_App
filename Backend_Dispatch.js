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
   notes: data.notes,
   flightDescription: String(data.flightDescription || '').trim(),
   communicationRequests: Array.isArray(data.communicationRequests) ? data.communicationRequests : [],
   training: (data && data.training && typeof data.training === 'object') ? data.training : null,
   cvaPolicy: null,
   mxReview: _buildMissionMxReview_(data.acft, (Array.isArray(data.legs) ? data.legs : []).reduce(function(acc, leg) {
     return acc + (parseFloat(leg && leg.time) || 0);
   }, 0), data && data.mxReview)
 };

 if (!legs || legs.length === 0) throw new Error("No legs to save.");

 // Determine missionId from first leg
 const firstFlightId = legs[0].flightLegId;
 const missionId = missionIdFromFlightLeg_(firstFlightId);
 if (!missionId) throw new Error("Invalid flight leg id format: " + firstFlightId);

 const aircraftType = _supervisorAircraftTypeFromRegistration_(header.acft);
 if (!aircraftType) {
   throw new Error("Could not resolve aircraft type for registration: " + header.acft);
 }

 const pilotAuthorized = isPilotAuthorizedForAircraftRole_(null, header.pilot, aircraftType, 'Operational', header.date)
   || isPilotAuthorizedForAircraftRole_(null, header.pilot, aircraftType, 'Instructor', header.date);
 if (!pilotAuthorized) {
   throw new Error(
     "Pilot " + header.pilot + " is not authorized for aircraft type " + aircraftType + "."
   );
 }

  header.cvaPolicy = _enforceMissionCvaSchedulingPolicy_(ss, header, missionId, data);

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
   _financeReverseMissionTransactions_(ss, missionId, Session.getActiveUser().getEmail() || 'system', 'MISSION_EDIT_REWRITE');
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
     waypoints: routeTokens.length ? routeTokens : (leg && leg.waypoints),
     flightDescription: header.flightDescription,
     communicationRequests: Array.isArray(leg && leg.communicationRequests)
       ? leg.communicationRequests
       : (Array.isArray(header.communicationRequests) ? header.communicationRequests : []),
     training: header.training ? { ...header.training } : null
   };

   const singleLegWrapper = JSON.stringify({
     legs: [{
       ...normalizedLeg,
       missionTime: header.time,
       meta: {
         time: header.time,
         type: header.type,
         training: header.training || null,
         flightDescription: header.flightDescription,
         communicationRequests: Array.isArray(normalizedLeg.communicationRequests) ? normalizedLeg.communicationRequests : [],
         cvaPolicy: header.cvaPolicy,
         mxReview: header.mxReview
       }
     }],
     time: header.time,
     type: header.type,
     flightDescription: header.flightDescription,
     communicationRequests: Array.isArray(header.communicationRequests) ? header.communicationRequests : [],
     missionMeta: {
       flightDescription: header.flightDescription,
       communicationRequests: Array.isArray(header.communicationRequests) ? header.communicationRequests : [],
       cvaPolicy: header.cvaPolicy,
       mxReview: header.mxReview
     },
     cvaPolicy: header.cvaPolicy,
     mxReview: header.mxReview,
     training: header.training || null
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

 invalidateScheduledMissionsCache_();

 try {
   var flightCalendarSync = _mxSyncFlightMissionCalendar_(missionId, header);
   if (flightCalendarSync && flightCalendarSync.success === false) {
     appLog_('Flight calendar sync failed for', missionId, flightCalendarSync.error || 'unknown error');
   }
 } catch (e) {
   appLog_('Flight calendar sync exception for', missionId, e && e.message ? e.message : String(e));
 }

 return "Success: " + missionId;
}

function getMissionDetailsForSupervisor(missionId) {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName(APP_SHEETS.DISPATCH);
const data = sheet.getDataRange().getValues();
let airportLookup = null;
try { airportLookup = _supervisorBuildAirportLookup_(); } catch (e) { airportLookup = null; }
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
const missionFlightDescription = _extractMissionDescriptionFromRaw_(mainRow[DISPATCH_COL.RAW_DATA]);
const missionCommunicationRequests = _extractMissionCommunicationsFromRaw_(mainRow[DISPATCH_COL.RAW_DATA]);
// 2. Safe Date handling
let rawDate = mainRow[DISPATCH_COL.DATE];
let dateStr = _dutyYmd_(rawDate) || ((rawDate instanceof Date) ? rawDate.toISOString().split('T')[0] : String(rawDate));
const missionDateObj = (function(v) {
  const ymd = _dutyYmd_(v);
  if (ymd) {
    const parts = String(ymd).split('-').map(function(n) { return Number(n); });
    const d = new Date(parts[0], parts[1] - 1, parts[2]);
    if (!isNaN(d.getTime())) {
      d.setHours(0, 0, 0, 0);
      return d;
    }
  }
  const fallback = new Date();
  fallback.setHours(0, 0, 0, 0);
  return fallback;
})(rawDate);

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
 notes: String(mainRow[DISPATCH_COL.NOTES] || ""),
 flightDescription: missionFlightDescription,
 communicationRequests: missionCommunicationRequests
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
 const runwayGap = _supervisorBuildRunwayGapForLeg_(airportLookup, parsedRoute, legPayload, r[DISPATCH_COL.ROUTE]);

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
     missionTime: legPayload.missionTime || "08:00",
     flightDescription: String(legPayload.flightDescription || (legPayload.meta && legPayload.meta.flightDescription) || missionFlightDescription || '').trim(),
     communicationRequests: Array.isArray(legPayload.communicationRequests)
       ? legPayload.communicationRequests
       : (Array.isArray(legPayload.meta && legPayload.meta.communicationRequests)
         ? legPayload.meta.communicationRequests
         : missionCommunicationRequests),
     runwayGap: runwayGap
 };
})
};

// 4. Helpers for the Supervisor Sidebar
const pName = String(missionData.meta.pilot);
const cName = String(missionData.meta.copilot || '').trim();
let timeline = [];
try { timeline = getPilotDutyTimeline(pName, missionDateObj); } catch(e) {}
let authString = "";
try { authString = getAuthorizedDestinations(pName); } catch(e) {}
let copilotAuthString = "";
try { copilotAuthString = cName ? getAuthorizedDestinations(cName) : ""; } catch(e) {}
return {
mission: missionData,
timeline: timeline,
authorizedAirports: authString,
authorizedPilotAirports: authString,
authorizedCopilotAirports: copilotAuthString
};
}

function addSupervisorMissionRequest(missionId, payload) {
  const missionKey = String(missionId || '').trim();
  if (!missionKey) throw new Error('missionId is required');
  if (!payload || typeof payload !== 'object') throw new Error('payload is required');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dispatchSheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, 'addSupervisorMissionRequest');
  const data = dispatchSheet.getDataRange().getValues();
  if (!data || data.length < 2) throw new Error('Dispatch sheet is empty');

  const request = {
    requestId: String(payload.requestId || ('REQ-' + missionKey + '-' + Date.now())).trim(),
    type: String(payload.type || 'SURVEY_RUNWAY').trim().toUpperCase(),
    status: String(payload.status || 'OPEN').trim().toUpperCase(),
    createdAt: String(payload.createdAt || new Date().toISOString()),
    createdBy: String(payload.createdBy || Session.getActiveUser().getEmail() || 'supervisor').trim(),
    message: String(payload.message || '').trim(),
    from: String(payload.from || '').trim().toUpperCase(),
    to: String(payload.to || '').trim().toUpperCase(),
    missionId: missionKey,
    flightLegId: String(payload.flightLegId || '').trim(),
    legIndex: Number.isFinite(Number(payload.legIndex)) ? Number(payload.legIndex) : -1
  };

  let updatedRows = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[DISPATCH_COL.MISSION_ID] || '').trim() !== missionKey) continue;

    const parsed = _safeParseJsonObject_(row[DISPATCH_COL.RAW_DATA] || '{}');
    const existing = _extractMissionCommunicationsFromRaw_(parsed);
    const exists = existing.some(function(entry) {
      return String(entry.requestId || '').trim() === request.requestId;
    });
    if (!exists) existing.push(request);

    if (!Array.isArray(parsed.legs) || !parsed.legs.length) parsed.legs = [{}];
    const targetIdx = (request.legIndex >= 0 && request.legIndex < parsed.legs.length) ? request.legIndex : 0;
    const legPayload = parsed.legs[targetIdx] && typeof parsed.legs[targetIdx] === 'object' ? parsed.legs[targetIdx] : {};
    const legExisting = Array.isArray(legPayload.communicationRequests) ? legPayload.communicationRequests.slice() : [];
    if (!legExisting.some(function(entry) { return String(entry.requestId || '').trim() === request.requestId; })) {
      legExisting.push(request);
    }
    legPayload.communicationRequests = legExisting;
    if (!legPayload.meta || typeof legPayload.meta !== 'object') legPayload.meta = {};
    legPayload.meta.communicationRequests = legExisting;
    parsed.legs[targetIdx] = legPayload;

    parsed.communicationRequests = existing;
    if (!parsed.missionMeta || typeof parsed.missionMeta !== 'object') parsed.missionMeta = {};
    parsed.missionMeta.communicationRequests = existing;

    dispatchSheet.getRange(i + 1, DISPATCH_COL.RAW_DATA + 1).setValue(JSON.stringify(parsed));
    updatedRows++;
  }

  if (!updatedRows) throw new Error('Mission not found: ' + missionKey);
  invalidateScheduledMissionsCache_();
  return { success: true, missionId: missionKey, requestId: request.requestId, rowsUpdated: updatedRows };
}

function approveMission(missionId, approvalPassword) {
_verifySupervisorApprovalPassword_(approvalPassword);
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName(APP_SHEETS.DISPATCH);
if (!sheet) return "Error: DB missing";
const data = sheet.getDataRange().getValues();
const user = "Admin";
let missionFound = false;
let missionPilotName = '';
let missionStudentName = '';
let missionType = '';
let missionTotalHours = 0;
for (let i = 1; i < data.length; i++) {
if (String(data[i][DISPATCH_COL.MISSION_ID]) === String(missionId)) {
 missionFound = true;
 const mxReview = _extractMissionMxReviewFromRaw_(data[i][DISPATCH_COL.RAW_DATA]);
 if (mxReview && mxReview.required && String(mxReview.status || '').trim().toUpperCase() !== 'APPROVED') {
  throw new Error('Mission requires chief mechanic approval before supervisor release.');
 }
 const pilotName = String(data[i][DISPATCH_COL.PILOT] || '').trim();
 const pilotKey = pilotName.toUpperCase();
 if (!pilotName || pilotKey === 'PILOT TBD' || pilotKey === 'TBD' || pilotKey === 'UNASSIGNED') {
  throw new Error('Mission cannot be approved without an assigned pilot.');
 }
 if (!missionPilotName) missionPilotName = pilotName;
 if (!missionStudentName) missionStudentName = _resolveTrainingStudentNameFromDispatchRow_(data[i]);
 if (!missionType) missionType = String(data[i][DISPATCH_COL.TYPE] || '').trim();
 missionTotalHours += Number(data[i][DISPATCH_COL.FLIGHT_TIME] || 0);
 sheet.getRange(i + 1, DISPATCH_COL.STATUS + 1).setValue("APPROVED");
}
}
if (missionFound) {
  _financePostMissionTransactions_(ss, missionId, Session.getActiveUser().getEmail() || 'Admin');
  if (missionType.toUpperCase() === 'TRAINING' && missionTotalHours > 0 && missionStudentName) {
    _toolsStudentDebitForMission_(ss, missionId, missionStudentName, missionTotalHours, Session.getActiveUser().getEmail() || 'Admin');
  }
}
invalidateScheduledMissionsCache_();
const audit = ss.getSheetByName(APP_SHEETS.AUDIT);
if(audit) audit.appendRow([new Date(), user, missionId, "APPROVE", "PENDING", "APPROVED", "Release"]);
return "Approved";
}

function scheduleRunwayCheckFromSupervisor(payload, approvalPassword) {
  _verifySupervisorApprovalPassword_(approvalPassword);

  var req = (payload && typeof payload === 'object') ? payload : {};
  var missionId = String(req.missionId || '').trim();
  var instructorName = String(req.instructorName || '').trim();
  var runwayLocation = String(req.runwayLocation || '').trim().toUpperCase();
  var routeRaw = String(req.route || '').trim().toUpperCase();
  var lessonName = String(req.lessonName || '').trim();

  if (!missionId) throw new Error('Mission ID is required.');
  if (!instructorName) throw new Error('Instructor name is required.');
  if (!runwayLocation) throw new Error('Runway check location is required.');
  if (!routeRaw) throw new Error('Route is required.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dispatchSheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, 'scheduleRunwayCheckFromSupervisor');
  var dispatchData = dispatchSheet.getDataRange().getValues();

  var missionRowIndexes = [];
  for (var i = 1; i < dispatchData.length; i++) {
    if (String(dispatchData[i][DISPATCH_COL.MISSION_ID] || '').trim() === missionId) missionRowIndexes.push(i);
  }
  if (!missionRowIndexes.length) throw new Error('Mission not found in dispatch: ' + missionId);

  var firstRow = dispatchData[missionRowIndexes[0]];
  var studentName = String(firstRow[DISPATCH_COL.PILOT] || '').trim();
  var aircraftReg = String(firstRow[DISPATCH_COL.AIRCRAFT] || '').trim().toUpperCase();
  var aircraftType = String(_supervisorAircraftTypeFromRegistration_(aircraftReg) || '').trim() || aircraftReg;
  var flightTime = parseFloat(firstRow[DISPATCH_COL.FLIGHT_TIME]);
  var routeTokens = _supervisorNormalizeRouteTokens_(routeRaw);
  var runwayTargets = _supervisorNormalizeRouteTokens_(runwayLocation);
  runwayTargets = runwayTargets.filter(function(code, idx, arr) { return arr.indexOf(code) === idx; });
  if (routeTokens.length < 2) throw new Error('Route must include at least origin and destination.');
  if (!runwayTargets.length) throw new Error('At least one runway check destination is required.');

  var syllabusSheet = getRequiredSheet_(ss, APP_SHEETS.SYLLABUS, 'scheduleRunwayCheckFromSupervisor');
  var sylHeaders = _toolsSheetHeaderRow_(syllabusSheet);
  var sylRows = syllabusSheet.getDataRange().getValues().slice(1);
  var trainingCode = _supervisorNextRunwayCheckCode_(sylHeaders, sylRows, aircraftType, routeTokens);

  var description = lessonName || ('Cheque de Pista: ' + runwayTargets.join('; ') + ' | ' + routeTokens[0] + ' -> ' + routeTokens[routeTokens.length - 1]);
  var requiredHours = isNaN(flightTime) || flightTime <= 0 ? 1 : Math.max(0.7, flightTime);

  var stops = runwayTargets.map(function(loc, idx) {
    var upperLoc = String(loc || '').trim().toUpperCase();
    return {
      sequence: idx + 1,
      location: upperLoc,
      landingType: 'FULL_STOP',
      breakStop: false,
      notes: '',
      landings: 1,
      touchAndGos: 0
    };
  });

  var planEnvelope = {
    version: 3,
    category: 'RUNWAY_CHECK',
    advisoryOnly: true,
    externalLessonName: lessonName,
    runwayCheckLocation: runwayTargets.join('; '),
    routeCheckPrompt: '',
    stops: stops,
    maneuvers: ['SHORT FIELD LANDING', 'GO-AROUND'],
    customManeuvers: []
  };

  var syllabusPayload = {
    TRAINING_CODE: trainingCode,
    AIRCRAFT_TYPE: aircraftType,
    DESCRIPTION: description,
    REQUIRED_HOURS: requiredHours,
    PLANNED_LANDINGS: Math.max(1, runwayTargets.length),
    REQUIRED_FUEL: '',
    REQUIRED_BALLAST: '',
    ROUTE: routeTokens.join(', '),
    RUNWAY_CHECKOUT: runwayTargets.join(';'),
    LESSON_PLAN_JSON: JSON.stringify(planEnvelope),
    PLANNED_STOPS_JSON: JSON.stringify(stops),
    MANEUVERS_JSON: JSON.stringify({ maneuvers: planEnvelope.maneuvers, customManeuvers: [] })
  };

  var syllabusRes = addToolsSheetRecord('syllabus', syllabusPayload);
  if (!syllabusRes || !syllabusRes.success) {
    throw new Error((syllabusRes && syllabusRes.error) ? syllabusRes.error : 'Failed to create syllabus runway-check lesson.');
  }

  var trainingObj = {
    code: trainingCode,
    description: description,
    requiredHours: requiredHours,
    requiredFuel: 0,
    requiredBallast: 0,
    route: syllabusPayload.ROUTE,
    plannedLandings: Math.max(1, runwayTargets.length),
    runwayCheckout: runwayTargets.join(';'),
    student: studentName,
    instructor: instructorName
  };

  missionRowIndexes.forEach(function(idx) {
    var rowNo = idx + 1;
    var oldNotes = String(dispatchData[idx][DISPATCH_COL.NOTES] || '').trim();
    var noteSuffix = 'RUNWAY CHECK ' + trainingCode + ' | Instrutor: ' + instructorName + ' | Aluno: ' + studentName;
    var newNotes = oldNotes ? (oldNotes + ' | ' + noteSuffix) : noteSuffix;

    dispatchSheet.getRange(rowNo, DISPATCH_COL.PILOT + 1).setValue(instructorName);
    dispatchSheet.getRange(rowNo, DISPATCH_COL.COPILOT + 1).setValue(studentName);
    dispatchSheet.getRange(rowNo, DISPATCH_COL.TYPE + 1).setValue('Training');
    dispatchSheet.getRange(rowNo, DISPATCH_COL.NOTES + 1).setValue(newNotes);
    dispatchSheet.getRange(rowNo, DISPATCH_COL.RAW_DATA + 1).setValue(_supervisorApplyTrainingToDispatchRaw_(dispatchData[idx][DISPATCH_COL.RAW_DATA], trainingObj));
  });

  invalidateScheduledMissionsCache_();

  return {
    success: true,
    missionId: missionId,
    trainingCode: trainingCode,
    student: studentName,
    instructor: instructorName,
    syllabusRowNumber: Number(syllabusRes.rowNumber || 0)
  };
}

function getPassengerConsentConfig() {
  return _passengerConsentReadConfigMap_();
}

function invalidateScheduledMissionsCache_() {
 const cache = CacheService.getScriptCache();
 cache.remove('scheduledMissions:v1');
 cache.remove('scheduledMissions:v2');
}

function getScheduledMissions() {
const cache = CacheService.getScriptCache();
const cacheKey = 'scheduledMissions:v2';
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
 FLIGHT_ID: DISPATCH_COL.FLIGHT_ID - 1,
 DATE: DISPATCH_COL.DATE - 1,
 AIRCRAFT: DISPATCH_COL.AIRCRAFT - 1,
 PILOT: DISPATCH_COL.PILOT - 1,
 COPILOT: DISPATCH_COL.COPILOT - 1,
 ROUTE: DISPATCH_COL.ROUTE - 1,
 RAW_DATA: DISPATCH_COL.RAW_DATA - 1,
 STATUS: DISPATCH_COL.STATUS - 1
};

const flownStatuses_ = { FLOWN: true, COMPLETE: true, COMPLETED: true };
const completedLegMap_ = {};
const departedLegMap_ = {};
const dispatchLegStateMap_ = {};
try {
  const logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  if (logSheet) {
    const logRows = logSheet.getDataRange().getValues();
    for (let li = 1; li < logRows.length; li++) {
      const lr = logRows[li];
      const legId = String(lr[LOG_FLIGHT_COL.FLIGHT_ID] || '').trim();
      if (!legId) continue;
      const released = lr[LOG_FLIGHT_COL.BRAKES_RELEASE];
      const landed = lr[LOG_FLIGHT_COL.LANDED];
      const onBlocks = lr[LOG_FLIGHT_COL.ON_BLOCKS];
      const brakesApplied = lr[LOG_FLIGHT_COL.BRAKES_APPLIED];
      if (released) departedLegMap_[legId] = true;
      if (landed || onBlocks || brakesApplied) completedLegMap_[legId] = true;
    }
  }
} catch (e) {}

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
     copilot: row[DISPATCH_RANGE_COL.COPILOT],
     status: row[DISPATCH_RANGE_COL.STATUS],
     flightDescription: _extractMissionDescriptionFromRaw_(row[DISPATCH_RANGE_COL.RAW_DATA]),
     communicationRequests: _extractMissionCommunicationsFromRaw_(row[DISPATCH_RANGE_COL.RAW_DATA]),
     mxReview: _extractMissionMxReviewFromRaw_(row[DISPATCH_RANGE_COL.RAW_DATA]),
     routeStops: [],
     flightLegIds: []
   };
 }

 const mission = missions[mId];
 const flightLegId = String(row[DISPATCH_RANGE_COL.FLIGHT_ID] || '').trim();
 if (flightLegId) mission.flightLegIds.push(flightLegId);
 const rowStatusRaw = String(row[DISPATCH_RANGE_COL.STATUS] || '').trim().toUpperCase();
 if (flightLegId) {
   const rowState = flownStatuses_[rowStatusRaw]
     ? 2
     : ((rowStatusRaw === 'IN-FLIGHT' || rowStatusRaw === 'IN FLIGHT' || rowStatusRaw === 'DEPARTED') ? 1 : 0);
   dispatchLegStateMap_[flightLegId] = Math.max(Number(dispatchLegStateMap_[flightLegId] || 0), rowState);
 }

 const legRoute = String(row[DISPATCH_RANGE_COL.ROUTE] || '');
 const routeParts = splitRoute_(legRoute);
 const legFrom = String(routeParts.from || '').trim().toUpperCase();
 const legTo = String(routeParts.to || '').trim().toUpperCase();

 if (legFrom) {
   if (!mission.routeStops.length || mission.routeStops[mission.routeStops.length - 1] !== legFrom) {
     mission.routeStops.push(legFrom);
   }
 }
 if (legTo) {
   if (!mission.routeStops.length || mission.routeStops[mission.routeStops.length - 1] !== legTo) {
     mission.routeStops.push(legTo);
   }
 }
}

// Convert the object back to an array for the frontend
const result = Object.values(missions).map(m => {
  const routeStops = Array.isArray(m.routeStops) ? m.routeStops.filter(Boolean) : [];
  const fromIcao = routeStops[0] || '';
  const toIcao = routeStops.length ? routeStops[routeStops.length - 1] : '';
  const routeDisplay = routeStops.length ? routeStops.join(' - ') : '';
  const statusRaw = String(m.status || '').toUpperCase() || 'PENDING';
  const legIds = Array.isArray(m.flightLegIds) ? m.flightLegIds.map(function(v) { return String(v || '').trim(); }).filter(Boolean) : [];
  const legCount = legIds.length;
  const legStates = legIds.map(function(legId) {
    if (completedLegMap_[legId]) return 2;
    if (Number(dispatchLegStateMap_[legId] || 0) >= 2) return 2;
    if (departedLegMap_[legId] || Number(dispatchLegStateMap_[legId] || 0) === 1) return 1;
    return 0;
  });
  let legsFlown = legStates.filter(function(st) { return st >= 2; }).length;
  // If a later leg is already departed, infer prior legs as complete even if logs are sparse.
  let inferredCompleteCount = 0;
  for (let si = 0; si < legStates.length; si++) {
    if (legStates[si] >= 1) inferredCompleteCount = Math.max(inferredCompleteCount, si);
  }
  legsFlown = Math.max(legsFlown, inferredCompleteCount);
  if (legsFlown > legCount) legsFlown = legCount;
  const allLegsComplete = legCount > 0 && legsFlown === legCount;
  const someLegsComplete = legsFlown > 0 && !allLegsComplete;
  let effectiveStatus;
  if (flownStatuses_[statusRaw] || allLegsComplete) {
    effectiveStatus = 'FLOWN';
  } else if (someLegsComplete) {
    effectiveStatus = legsFlown + ' OF ' + legCount + ' LEGS';
  } else {
    effectiveStatus = statusRaw;
  }

  return {
    id: m.id,
    date: m.date,
    acft: m.acft,
    pilot: m.pilot,
    status: effectiveStatus,
    flightDescription: String(m.flightDescription || '').trim(),
    communicationRequests: Array.isArray(m.communicationRequests) ? m.communicationRequests : [],
    mxReview: (m.mxReview && typeof m.mxReview === 'object') ? m.mxReview : {},
    legsFlown: legsFlown,
    legCount: legCount,
    from: fromIcao,
    to: toIcao,
    route: routeDisplay
  };
}).reverse().slice(0, 15);

cache.put(cacheKey, JSON.stringify(result), 45);
return result;
}

function cancelMissionFromDatabase(missionId) {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
  var calendarSync = null;

 // 1. REVERSE ALL FUEL FOR THIS MISSION
 // We do this first so we can read the log data before deleting it
 reverseFuelForMission(missionId);
 _financeReverseMissionTransactions_(ss, missionId, Session.getActiveUser().getEmail() || 'system', 'MISSION_CANCEL');

 // 2. DELETE FROM DB_DISPATCH
 const dbSheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, "cancelMissionFromDatabase");
 const dbData = dbSheet.getDataRange().getValues();
 const keptDispatchRows = [dbData[0]].concat(
   dbData.slice(1).filter(r => String(r[DISPATCH_COL.MISSION_ID]) !== String(missionId))
 );
 rewriteSheetData_(dbSheet, keptDispatchRows);
 invalidateScheduledMissionsCache_();

  try {
    calendarSync = _mxSyncFlightMissionCalendar_(missionId, { action: 'DELETE', deleteOnly: true });
    if (calendarSync && calendarSync.success === false) {
      appLog_('Flight calendar delete sync failed for', missionId, calendarSync.error || 'unknown error');
    }
  } catch (e) {
    appLog_('Flight calendar delete sync exception for', missionId, e && e.message ? e.message : String(e));
  }

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

 _ensurePassengerConsentColumns_(sheet);

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

function backfillAllFormsPassengers(options) {
  var opts = (options && typeof options === 'object') ? options : {};
  var maxRowsPerCall = parseInt(opts.maxRowsPerCall, 10);
  if (!(maxRowsPerCall > 0)) maxRowsPerCall = 200;

  var startForm1 = parseInt(opts.startForm1, 10);
  var startForm2 = parseInt(opts.startForm2, 10);
  var startForm3 = parseInt(opts.startForm3, 10);
  if (!(startForm1 > 1)) startForm1 = 2;
  if (!(startForm2 > 1)) startForm2 = 2;
  if (!(startForm3 > 1)) startForm3 = 2;

  var form1 = ingestForm1ResponsesBatch({
    sheetName: String(APP_SHEETS.FORM1_RESPONSES || 'Form Responses 1'),
    startRow: startForm1,
    maxRows: maxRowsPerCall
  });

  var form2 = ingestForm2ResponsesBatch({
    sheetName: String(APP_SHEETS.FORM2_RESPONSES || 'Form Responses 2'),
    startRow: startForm2,
    maxRows: maxRowsPerCall
  });

  var form3 = ingestForm3ResponsesBatch({
    sheetName: String(APP_SHEETS.FORM3_RESPONSES || 'Form Responses 3'),
    startRow: startForm3,
    maxRows: maxRowsPerCall
  });

  function nextStart(startRow, batchResult) {
    var processed = batchResult && batchResult.processed ? batchResult.processed : 0;
    return startRow + processed;
  }

  return {
    success: true,
    config: {
      maxRowsPerCall: maxRowsPerCall,
      starts: {
        form1: startForm1,
        form2: startForm2,
        form3: startForm3
      }
    },
    forms: {
      form1: form1,
      form2: form2,
      form3: form3
    },
    next: {
      form1: nextStart(startForm1, form1),
      form2: nextStart(startForm2, form2),
      form3: nextStart(startForm3, form3)
    }
  };
}

function backfillAllFormsPassengersUntilDone(options) {
  var opts = (options && typeof options === 'object') ? options : {};
  var maxRowsPerCall = parseInt(opts.maxRowsPerCall, 10);
  var maxRounds = parseInt(opts.maxRounds, 10);
  var maxRuntimeMs = parseInt(opts.maxRuntimeMs, 10);
  if (!(maxRowsPerCall > 0)) maxRowsPerCall = 200;
  if (!(maxRounds > 0)) maxRounds = 10;
  if (!(maxRuntimeMs > 0)) maxRuntimeMs = 240000;

  var starts = {
    form1: parseInt(opts.startForm1, 10),
    form2: parseInt(opts.startForm2, 10),
    form3: parseInt(opts.startForm3, 10)
  };
  if (!(starts.form1 > 1)) starts.form1 = 2;
  if (!(starts.form2 > 1)) starts.form2 = 2;
  if (!(starts.form3 > 1)) starts.form3 = 2;

  var startedAt = Date.now();
  var rounds = [];
  var totals = {
    processed: 0,
    imported: 0,
    skipped: 0,
    failed: 0
  };

  function addTotals(batch) {
    var forms = (batch && batch.forms) ? batch.forms : {};
    ['form1', 'form2', 'form3'].forEach(function(k) {
      var s = forms[k] || {};
      totals.processed += Number(s.processed || 0);
      totals.imported += Number(s.imported || 0);
      totals.skipped += Number(s.skipped || 0);
      totals.failed += Number(s.failed || 0);
    });
  }

  var stopReason = 'max_rounds_reached';
  for (var i = 0; i < maxRounds; i++) {
    if ((Date.now() - startedAt) >= maxRuntimeMs) {
      stopReason = 'max_runtime_reached';
      break;
    }

    var batch = backfillAllFormsPassengers({
      maxRowsPerCall: maxRowsPerCall,
      startForm1: starts.form1,
      startForm2: starts.form2,
      startForm3: starts.form3
    });

    rounds.push({
      round: i + 1,
      starts: { form1: starts.form1, form2: starts.form2, form3: starts.form3 },
      results: batch && batch.forms ? batch.forms : {},
      next: batch && batch.next ? batch.next : {}
    });
    addTotals(batch);

    var next = batch && batch.next ? batch.next : {};
    var nextForm1 = Number(next.form1 || starts.form1);
    var nextForm2 = Number(next.form2 || starts.form2);
    var nextForm3 = Number(next.form3 || starts.form3);

    var progressed = (nextForm1 > starts.form1) || (nextForm2 > starts.form2) || (nextForm3 > starts.form3);
    starts.form1 = nextForm1;
    starts.form2 = nextForm2;
    starts.form3 = nextForm3;

    if (!progressed) {
      stopReason = 'no_more_rows_processed';
      break;
    }
  }

  var elapsedMs = Date.now() - startedAt;
  return {
    success: true,
    complete: stopReason === 'no_more_rows_processed',
    stopReason: stopReason,
    elapsedMs: elapsedMs,
    config: {
      maxRowsPerCall: maxRowsPerCall,
      maxRounds: maxRounds,
      maxRuntimeMs: maxRuntimeMs
    },
    totals: totals,
    next: {
      form1: starts.form1,
      form2: starts.form2,
      form3: starts.form3
    },
    rounds: rounds
  };
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
const staffRows = _toolsLoadStaffRows_();
const resolveCrewName_ = function(raw) {
  const text = String(raw || '').trim();
  if (!text) return '';
  const match = _toolsFindStaffByEmailOrId_(staffRows, text, '');
  return String((match && match.staffName) || text).trim();
};
const resolvedPilot = resolveCrewName_(mainRow[DISPATCH_COL.PILOT]);
const resolvedCopilot = resolveCrewName_(mainRow[DISPATCH_COL.COPILOT] || '');
const aircraftReg = String(mainRow[DISPATCH_COL.AIRCRAFT] || '').trim().toUpperCase();
const resolvedAircraftType = aircraftReg ? String(_supervisorAircraftTypeFromRegistration_(aircraftReg) || '').trim() : '';

const missionFlightDescription = _extractMissionDescriptionFromRaw_(mainRow[DISPATCH_COL.RAW_DATA]);
const missionCommunicationRequests = _extractMissionCommunicationsFromRaw_(mainRow[DISPATCH_COL.RAW_DATA]);

// 2. Safe Date handling
let rawDate = mainRow[DISPATCH_COL.DATE];
let dateStr = (rawDate instanceof Date) ? rawDate.toISOString().split('T')[0] : String(rawDate);
const missionDateObj = (rawDate instanceof Date) ? rawDate : new Date();

// 3. Build mission object
const missionData = {
  id: mainRow[DISPATCH_COL.MISSION_ID],
  date: dateStr,
  acft: aircraftReg,
  aircraftType: resolvedAircraftType,
  pilot: resolvedPilot,
  status: mainRow[DISPATCH_COL.STATUS] ? mainRow[DISPATCH_COL.STATUS].toString().toUpperCase() : "PENDING",
  meta: {
    date: dateStr,
    acft: aircraftReg,
    aircraftType: resolvedAircraftType,
    pilot: resolvedPilot,
    copilot: resolvedCopilot,
    type: String(mainRow[DISPATCH_COL.TYPE] || ""),
    notes: String(mainRow[DISPATCH_COL.NOTES] || ""),
    flightDescription: missionFlightDescription,
    communicationRequests: missionCommunicationRequests,
    training: null
  },
  // 4. Parse legs
  legs: missionRows.map(r => {
    let legPayload = {};
    let rawJson = {};
    try {
      rawJson = JSON.parse(r[DISPATCH_COL.RAW_DATA] || "{}");
      if (rawJson.legs && Array.isArray(rawJson.legs)) legPayload = rawJson.legs[0] || rawJson;
      else legPayload = rawJson;
    } catch (e) { rawJson = {}; legPayload = {}; }

    const legTraining = (legPayload && legPayload.training && typeof legPayload.training === 'object')
      ? legPayload.training
      : ((legPayload && legPayload.meta && legPayload.meta.training && typeof legPayload.meta.training === 'object')
        ? legPayload.meta.training
        : ((rawJson && rawJson.training && typeof rawJson.training === 'object')
          ? rawJson.training
          : null));
    const resolvedTraining = legTraining
      ? Object.assign({}, legTraining, {
          student: resolveCrewName_(legTraining.student || ''),
          instructor: resolveCrewName_(legTraining.instructor || '')
        })
      : null;

    const resolvedPax = Array.isArray(legPayload.pax)
      ? legPayload.pax.map(function(p) {
          if (!p || typeof p !== 'object') return p;
          var out = Object.assign({}, p);
          if (out.name) out.name = resolveCrewName_(out.name);
          return out;
        })
      : (legPayload.pax || []);

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
      pax: resolvedPax,
      limitType: legPayload.limitType || "",
      isOver: legPayload.isOver || false,
      missionTime: legPayload.missionTime || "08:00",
      flightDescription: String(legPayload.flightDescription || (legPayload.meta && legPayload.meta.flightDescription) || missionFlightDescription || "").trim(),
      communicationRequests: Array.isArray(legPayload.communicationRequests)
        ? legPayload.communicationRequests
        : (Array.isArray(legPayload.meta && legPayload.meta.communicationRequests)
          ? legPayload.meta.communicationRequests
          : missionCommunicationRequests),
      training: resolvedTraining,
      logStatus: 'PENDING',  // enriched below
      bracesRelease: null,
      onBlocks: null
    };
  }),
  actualFuelLogs: getFuelLogsForMission(missionId)
};

if (!missionData.meta.training && missionData.legs && missionData.legs.length) {
  missionData.meta.training = missionData.legs[0].training || null;
}

// Enrich legs with flight status from LOG_Flights sheet
const logSheetForStatus_ = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
if (logSheetForStatus_) {
  const logData_ = logSheetForStatus_.getDataRange().getValues();
  const logStatusMap_ = {};
  logData_.slice(1).forEach(function(r) {
    const lid = String(r[LOG_FLIGHT_COL.FLIGHT_ID] || '').trim();
    if (!lid) return;
    const released = r[LOG_FLIGHT_COL.BRAKES_RELEASE];
    const landed = r[LOG_FLIGHT_COL.LANDED];
    const onBlks   = r[LOG_FLIGHT_COL.ON_BLOCKS];
    const brakesApplied = r[LOG_FLIGHT_COL.BRAKES_APPLIED];
    const relStr = (released instanceof Date) ? released.toISOString() : (released ? String(released).trim() : null);
    const ldStr  = (landed   instanceof Date) ? landed.toISOString()   : (landed   ? String(landed).trim()   : null);
    const obStr  = (onBlks   instanceof Date) ? onBlks.toISOString()   : (onBlks   ? String(onBlks).trim()   : null);
    const baStr  = (brakesApplied instanceof Date) ? brakesApplied.toISOString() : (brakesApplied ? String(brakesApplied).trim() : null);
    logStatusMap_[lid] = {
      bracesRelease: relStr || null,
      onBlocks:      obStr  || null,
      brakesApplied: baStr || null,
      logStatus: (ldStr || obStr || baStr) ? 'COMPLETE' : (relStr ? 'DEPARTED' : 'PENDING')
    };
  });
  missionData.legs.forEach(function(leg) {
    const s = logStatusMap_[leg.flightLegId] || {};
    leg.bracesRelease = s.bracesRelease || null;
    leg.onBlocks      = s.onBlocks      || null;
    leg.brakesApplied = s.brakesApplied || null;
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

function getFlightFollowMissionsForAcft(reg) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(APP_SHEETS.DISPATCH);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  // Build airport and waypoint lookup maps for route plotting.
  var airportMetaMap = {};
  var waypointMetaMap = {};
  try {
    var airSheet = ss.getSheetByName(APP_SHEETS.AIRPORTS);
    if (airSheet) {
      var airData = airSheet.getDataRange().getValues();
      if (airData.length > 1) {
        var airH = airData[0].map(function(h) { return String(h || '').toUpperCase().trim().replace(/\s+/g, '_'); });
        var icaoIdx = airH.indexOf('ICAO');
        var fuelIdx = airH.indexOf('FUEL_AVAILABLE');
        var latIdx = airH.indexOf('LATITUDE');
        var lonIdx = airH.indexOf('LONGITUDE');
        var surfIdx = airH.indexOf('SURFACE_ACTUAL');
        if (surfIdx < 0) surfIdx = airH.indexOf('RUNWAY_SURFACE_ACTUAL');
        if (surfIdx < 0) surfIdx = airH.indexOf('SURFACE_OFFICIAL');
        if (surfIdx < 0) surfIdx = airH.indexOf('RUNWAY_SURFACE');
        if (surfIdx < 0) surfIdx = airH.indexOf('SURFACE');
        var nameIdx = airH.indexOf('NOME');
        if (icaoIdx >= 0) {
          for (var ai = 1; ai < airData.length; ai++) {
            var aIcao = String(airData[ai][icaoIdx] || '').trim().toUpperCase();
            var aFuel = fuelIdx >= 0 ? String(airData[ai][fuelIdx] || '').trim().toUpperCase() : '';
            var aLat = latIdx >= 0 ? Number(airData[ai][latIdx]) : NaN;
            var aLon = lonIdx >= 0 ? Number(airData[ai][lonIdx]) : NaN;
            if (aIcao) {
              airportMetaMap[aIcao] = {
                icao: aIcao,
                name: nameIdx >= 0 ? String(airData[ai][nameIdx] || '').trim() : '',
                hasFuel: !!(aFuel && aFuel !== 'NONE' && aFuel !== 'N' && aFuel !== 'NO' && aFuel !== '0'),
                fuelAvailable: aFuel,
                runwaySurfaceActual: surfIdx >= 0 ? String(airData[ai][surfIdx] || '').trim() : '',
                lat: isFinite(aLat) ? aLat : null,
                lon: isFinite(aLon) ? aLon : null,
                kind: 'airport'
              };
            }
          }
        }
      }
    }

    var wpSheet = ss.getSheetByName(APP_SHEETS.WAYPOINTS);
    if (wpSheet) {
      var wpData = wpSheet.getDataRange().getValues();
      if (wpData.length > 1) {
        var wpH = wpData[0].map(function(h) { return String(h || '').toUpperCase().trim().replace(/\s+/g, '_'); });
        var wpIdIdx = wpH.indexOf('WP_ID');
        var wpLatIdx = wpH.indexOf('LATITUDE');
        var wpLonIdx = wpH.indexOf('LONGITUDE');
        var wpTypeIdx = wpH.indexOf('TYPE');
        if (wpIdIdx >= 0 && wpLatIdx >= 0 && wpLonIdx >= 0) {
          for (var wi = 1; wi < wpData.length; wi++) {
            var wpId = String(wpData[wi][wpIdIdx] || '').trim().toUpperCase();
            if (!wpId) continue;
            var wpLat = Number(wpData[wi][wpLatIdx]);
            var wpLon = Number(wpData[wi][wpLonIdx]);
            waypointMetaMap[wpId] = {
              wp_id: wpId,
              lat: isFinite(wpLat) ? wpLat : null,
              lon: isFinite(wpLon) ? wpLon : null,
              type: wpTypeIdx >= 0 ? String(wpData[wi][wpTypeIdx] || '').trim() : '',
              kind: 'waypoint'
            };
          }
        }
      }
    }
  } catch(e) {}

  // Gather all rows for this registration (no day filter)
  var tz      = 'America/Sao_Paulo';

  // Load per-leg fuel and leg plan fields from LOG_Flights (Tab 2 save), keyed by flight leg id.
  var fuelByFlightId = {};
  var planByFlightId = {};
  try {
    var logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
    if (logSheet) {
      var logData = logSheet.getDataRange().getValues();
      if (logData.length > 1) {
        var logHeaders = logData[0].map(function(h) { return String(h || '').trim().toUpperCase().replace(/\s+/g, '_'); });
        var col = function(name, fallbackIdx) {
          var idx = logHeaders.indexOf(String(name || '').toUpperCase());
          return idx >= 0 ? idx : fallbackIdx;
        };
        var fidIdx = col('FLIGHT_ID', LOG_FLIGHT_COL.FLIGHT_ID);
        var fuelStartIdx = col('FUEL_START', LOG_FLIGHT_COL.FUEL_START);
        var actualLoadIdx = col('ACTUAL_LOAD_JSON', LOG_FLIGHT_COL.ACTUAL_LOAD_JSON);
        var toNum = function(v) {
          var n = Number(v);
          return isFinite(n) ? n : 0;
        };

        for (var li = 1; li < logData.length; li++) {
          var fid = String(logData[li][fidIdx] || '').trim();
          if (!fid) continue;

          var fuelFromActualJson = 0;
          try {
            var rawLoad = String(logData[li][actualLoadIdx] || '').trim();
            if (rawLoad) {
              var loadObj = JSON.parse(rawLoad);
              if (loadObj && loadObj.fuelTotal != null) fuelFromActualJson = toNum(loadObj.fuelTotal);
              if (loadObj && Array.isArray(loadObj.legs) && loadObj.legs.length) {
                var legFromLog = loadObj.legs.find(function(x) {
                  return String(x && x.flightLegId || '').trim() === fid;
                }) || loadObj.legs[0] || null;
                if (legFromLog) {
                  var p = String(legFromLog.planId || '').trim().toUpperCase();
                  var t = String(legFromLog.takeoffUTC || '').trim().replace(/[^0-9]/g, '').slice(0, 4);
                  var n = !!legFromLog.noPlan;
                  if (p || t || n) {
                    planByFlightId[fid] = { planId: p, takeoffUTC: t, noPlan: n };
                  }
                }
              }
            }
          } catch (e) {}

          var fuelFromStartCol = toNum(logData[li][fuelStartIdx]);
          var finalFuel = fuelFromActualJson > 0 ? fuelFromActualJson : fuelFromStartCol;
          if (finalFuel > 0) fuelByFlightId[fid] = finalFuel;
        }
      }
    }
  } catch (e) {}

  var missions = {};

  for (var i = 1; i < data.length; i++) {
    var row   = data[i];
    var acft  = String(row[DISPATCH_COL.AIRCRAFT] || '').trim().toUpperCase();
    if (acft !== String(reg || '').trim().toUpperCase()) continue;
    var statusRaw = String(row[DISPATCH_COL.STATUS] || '').trim().toUpperCase();
    if (statusRaw === 'FLOWN' || statusRaw === 'COMPLETE' || statusRaw === 'COMPLETED' || statusRaw === 'CANCELLED') continue;

    var rawDate = row[DISPATCH_COL.DATE];
    var dateStr = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, tz, 'yyyy-MM-dd')
      : String(rawDate || '').trim().slice(0, 10);
    var missionId  = String(row[DISPATCH_COL.MISSION_ID] || '').trim();
    var flightLegId = String(row[DISPATCH_COL.FLIGHT_ID] || '').trim();
    if (!missionId || !flightLegId) continue;

    // Parse RAW_DATA blob
    var raw = {};
    try { raw = JSON.parse(String(row[DISPATCH_COL.RAW_DATA] || '{}')); } catch(e) {}

    var legs = Array.isArray(raw.legs) ? raw.legs : [];
    var leg = legs.find(function(l) {
      return String(l && (l.flightLegId || l.flight_id || l.id) || '').trim() === flightLegId;
    }) || legs[0] || {};

    var pax = Array.isArray(leg.pax) ? leg.pax : [];
    var paxFiltered = pax.filter(function(p) {
      return p && String(p.name || '').toUpperCase() !== 'FREIGHT';
    });
    var pobCount = 1 + (String(row[DISPATCH_COL.COPILOT] || '').trim() ? 1 : 0) + paxFiltered.length;

    // Flight plan
    var planId      = String(leg.planId || leg.planDI || raw.planId || '').trim().toUpperCase();
    var takeoffUTC  = String(leg.takeoffUTC || leg.takeoffZulu || raw.takeoffUTC || raw.time || '').trim().replace(/[^0-9]/g,'').slice(0,4);
    var plannedFlightLevel = String(leg.plannedFlightLevel || leg.flightLevel || raw.plannedFlightLevel || raw.flightLevel || '').trim().toUpperCase();
    var noPlan      = !!(leg.noPlan || raw.noPlan);
    var logPlan = planByFlightId[flightLegId] || null;
    if (logPlan) {
      if (logPlan.planId) planId = String(logPlan.planId || '').trim().toUpperCase();
      if (logPlan.takeoffUTC) takeoffUTC = String(logPlan.takeoffUTC || '').trim().replace(/[^0-9]/g, '').slice(0, 4);
      if (logPlan.noPlan) noPlan = true;
    }

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

    // Annotate each waypoint with airport metadata when available.
    waypoints = waypoints.map(function(w) {
      var airportMeta = Object.prototype.hasOwnProperty.call(airportMetaMap, w.fix)
        ? airportMetaMap[w.fix] : null;
      var wpMeta = !airportMeta && Object.prototype.hasOwnProperty.call(waypointMetaMap, w.fix)
        ? waypointMetaMap[w.fix] : null;
      return {
        fix: w.fix,
        distNm: w.distNm,
        hasFuel: airportMeta ? airportMeta.hasFuel : null,
        fuelAvailable: airportMeta ? airportMeta.fuelAvailable : '',
        surface: airportMeta ? airportMeta.runwaySurfaceActual : '',
        lat: airportMeta ? airportMeta.lat : (wpMeta ? wpMeta.lat : null),
        lon: airportMeta ? airportMeta.lon : (wpMeta ? wpMeta.lon : null),
        type: wpMeta ? wpMeta.type : '',
        kind: airportMeta ? 'airport' : (wpMeta ? 'waypoint' : 'fix')
      };
    });

    var fromMeta = Object.prototype.hasOwnProperty.call(airportMetaMap, from) ? airportMetaMap[from] : null;
    var toMeta   = Object.prototype.hasOwnProperty.call(airportMetaMap, to) ? airportMetaMap[to] : null;

    // Fuel precedence: Tab 2 logged total (LOG_Flights) > takeoff/planned from dispatch raw.
    var fuelFromLog = Number(fuelByFlightId[flightLegId] || 0);
    var fuelFromDispatch = Number(leg.takeoffFuel || leg.fuel || leg.plannedFuel || raw.fuel || 0);
    var fuel = fuelFromLog > 0 ? fuelFromLog : fuelFromDispatch;

    if (!missions[flightLegId]) {
      missions[flightLegId] = {
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
        plannedFlightLevel: plannedFlightLevel,
        takeoffUTC:  takeoffUTC,
        noPlan:      noPlan,
        fuelL:       fuel,
        waypoints:   waypoints,
        resolvedWaypoints: waypoints.filter(function(w) { return isFinite(w.lat) && isFinite(w.lon); }),
        fromCoord:   fromMeta && isFinite(fromMeta.lat) && isFinite(fromMeta.lon) ? { lat: fromMeta.lat, lon: fromMeta.lon } : null,
        toCoord:     toMeta && isFinite(toMeta.lat) && isFinite(toMeta.lon) ? { lat: toMeta.lat, lon: toMeta.lon } : null,
        status:      String(row[DISPATCH_COL.STATUS] || '').trim()
      };
    }
  }

  var out = Object.values(missions);
  out.sort(function(a, b) {
    var ad = String(a.date || '');
    var bd = String(b.date || '');
    if (ad !== bd) return bd.localeCompare(ad); // newest date first
    var at = String(a.takeoffUTC || '');
    var bt = String(b.takeoffUTC || '');
    if (at !== bt) return bt.localeCompare(at); // latest time first
    return String(a.missionId || '').localeCompare(String(b.missionId || ''));
  });
  return out;
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
      { CONFIG_KEY: 'FLIGHTS_CALENDAR_LINK', CONFIG_VALUE: '', DESCRIPTION: 'Google Calendar link for flight events', ACTIVE: 'Y' },
      { CONFIG_KEY: 'MAINTENANCE_CALENDAR_LINK', CONFIG_VALUE: '', DESCRIPTION: 'Google Calendar link for maintenance events', ACTIVE: 'Y' },
      { CONFIG_KEY: 'FLIGHTS_CALENDAR_ID', CONFIG_VALUE: '', DESCRIPTION: 'Target Google Calendar id for flight events', ACTIVE: 'Y' },
      { CONFIG_KEY: 'SCHEDULE_CALENDAR_ID', CONFIG_VALUE: '', DESCRIPTION: 'Target Google Calendar id for staffing schedule', ACTIVE: 'Y' },
      { CONFIG_KEY: 'FLIGHT_FOLLOW_RECIPIENTS', CONFIG_VALUE: 'acompanhamento@asasdesocorro.org.br', DESCRIPTION: 'Flight Following inbox recipient list (comma-separated).', ACTIVE: 'Y' },
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

    var staffPool = _schedulerBuildStaffPool_();
    var roleSet = {};
    var qualSeedRows = [];
    for (var sp = 0; sp < staffPool.length; sp++) {
      var st = staffPool[sp] || {};
      var email = _schedulerEmailKey_(st.email);
      if (!email) continue;
      var roles = _toolsStaffRoleSet_(st).map(function(role) { return _schedulerRoleKey_(role); }).filter(function(role) { return !!role; });
      if (!roles.length) {
        var primary = _schedulerRoleKey_(st.primaryRole || '');
        if (primary) roles = [primary];
      }
      for (var rr = 0; rr < roles.length; rr++) {
        var roleCode = roles[rr];
        roleSet[roleCode] = true;
        var qualId = 'AUTOQUAL-' + email.replace(/[^a-z0-9]/gi, '').toUpperCase() + '-' + roleCode.replace(/[^A-Z0-9]/g, '');
        qualSeedRows.push({
          QUAL_ID: qualId,
          STAFF_EMAIL: email,
          ROLE_CODE: roleCode,
          LOCATION_CODE: '*',
          VALID_FROM: '',
          VALID_UNTIL: '',
          ACTIVE: 'Y',
          SOURCE: 'AUTO_FROM_DB_PILOTS',
          NOTES: 'Auto-seeded from DB_Pilots role data',
          UPDATED_AT: now,
          UPDATED_BY: _schedulerCurrentUserEmail_() || 'system'
        });
      }
    }

    var coverageRows = [];
    Object.keys(roleSet).sort().forEach(function(roleCode) {
      coverageRows.push({
        REQUIREMENT_ID: 'AUTOREQ-' + roleCode,
        ROLE_CODE: roleCode,
        LOCATION_CODE: '*',
        DAYS_MASK: 'MON,TUE,WED,THU,FRI',
        SHIFT_START_LOCAL: '08:00',
        SHIFT_END_LOCAL: '17:00',
        SHIFT_LABEL: 'DAY',
        REQUIRED_COUNT: 1,
        PRIORITY: 100,
        ALLOW_MULTI_ROLE_OVERLAP: 'N',
        ACTIVE: 'Y',
        NOTES: 'Auto-seeded default weekday coverage from DB_Pilots roles'
      });
    });

    var seededQuals = _schemaSeedRowsByKey_(qualSheetRes.sheet, 'QUAL_ID', qualSeedRows);
    var seededCoverage = _schemaSeedRowsByKey_(coverageSheetRes.sheet, 'REQUIREMENT_ID', coverageRows);

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
      seededCoverageRulesFromStaff: seededCoverage.inserted,
      seededQualificationsFromStaff: seededQuals.inserted,
      discoveredStaffRoles: Object.keys(roleSet).sort(),
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

function generateSchedulerAssignments(payload) {
  try {
    var actor = _schedulerAssertPermission_('CAN_GENERATE', 'generateSchedulerAssignments');
    var body = (payload && typeof payload === 'object') ? payload : {};
    var month = String(body.month || body.scheduleMonth || '').trim();
    var parsed = _schedulerParseMonthKey_(month || _schedulerMonthKey_(new Date()));
    if (!parsed) return { success: false, error: 'Invalid month format. Use YYYY-MM.' };
    var dryRun = !Object.prototype.hasOwnProperty.call(body, 'dryRun') || !!body.dryRun;

    var setupRes = setupSchedulerSchema();
    if (!setupRes || !setupRes.success) {
      return {
        success: false,
        error: (setupRes && setupRes.error) ? setupRes.error : 'Scheduler setup failed before generation',
        stage: 'setupSchedulerSchema'
      };
    }

    var diagnostics = { month: parsed.monthKey };
    var staffPool;
    var quals;
    var needs;
    var availabilityMap;
    var compatMap;
    try {
      staffPool = _schedulerBuildStaffPool_();
      diagnostics.staffPoolSize = staffPool.length;
    } catch (eStaff) {
      return { success: false, error: eStaff && eStaff.message ? eStaff.message : String(eStaff), stage: 'buildStaffPool', diagnostics: diagnostics };
    }
    try {
      quals = _schedulerQualificationRows_();
      diagnostics.qualificationRowCount = quals.length;
    } catch (eQual) {
      return { success: false, error: eQual && eQual.message ? eQual.message : String(eQual), stage: 'qualificationRows', diagnostics: diagnostics };
    }
    try {
      needs = _schedulerCoverageNeeds_(parsed.monthKey);
      diagnostics.coverageNeedCount = needs.length;
    } catch (eNeed) {
      return { success: false, error: eNeed && eNeed.message ? eNeed.message : String(eNeed), stage: 'coverageNeeds', diagnostics: diagnostics };
    }
    try {
      availabilityMap = _schedulerAvailabilityMap_(parsed.monthKey);
      diagnostics.availabilityKeys = Object.keys(availabilityMap || {}).length;
    } catch (eAvail) {
      return { success: false, error: eAvail && eAvail.message ? eAvail.message : String(eAvail), stage: 'availabilityMap', diagnostics: diagnostics };
    }
    try {
      compatMap = _schedulerCompatibilityMap_();
      diagnostics.compatibilityRuleCount = Object.keys(compatMap || {}).length;
    } catch (eCompat) {
      return { success: false, error: eCompat && eCompat.message ? eCompat.message : String(eCompat), stage: 'compatibilityMap', diagnostics: diagnostics };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var assignmentsSheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_ASSIGNMENTS || 'SCHED_Assignments', 'generateSchedulerAssignments');
    var locksSheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_LOCKS || 'SCHED_Assignment_Locks', 'generateSchedulerAssignments');

    var assignmentsData = assignmentsSheet.getDataRange().getValues();
    var assignmentHeaders = assignmentsData.length ? assignmentsData[0] : [];
    var lockData = locksSheet.getDataRange().getValues();
    var lockHeaders = lockData.length ? lockData[0] : [];
    var lockIdIdx = _schedulerHeaderIndex_(lockHeaders, 'ASSIGNMENT_ID');
    var lockActiveIdx = _schedulerHeaderIndex_(lockHeaders, 'ACTIVE');
    var activeLockedIds = {};
    for (var l = 1; l < lockData.length; l++) {
      var active = lockActiveIdx < 0 ? true : _schedulerTruthyFlag_(lockData[l][lockActiveIdx]);
      if (!active) continue;
      var aid = lockIdIdx >= 0 ? String(lockData[l][lockIdIdx] || '').trim() : '';
      if (aid) activeLockedIds[aid] = true;
    }

    var existingForMonth = [];
    var keepRows = assignmentHeaders.length ? [assignmentHeaders] : [];
    for (var r = 1; r < assignmentsData.length; r++) {
      var rec = _schedulerAssignmentRowFromHeaders_(assignmentHeaders, assignmentsData[r], r + 1);
      var isMonth = rec.scheduleMonth === parsed.monthKey;
      var isLocked = rec.locked || !!activeLockedIds[rec.assignmentId];
      if (isMonth) existingForMonth.push(rec);

      if (!isMonth) {
        keepRows.push(assignmentsData[r]);
        continue;
      }

      if (isLocked) {
        keepRows.push(assignmentsData[r]);
        continue;
      }

      if (rec.source !== 'AUTO_GENERATED') {
        keepRows.push(assignmentsData[r]);
      }
    }

    var assignedByDateEmail = {};
    var fairness = {};

    existingForMonth.forEach(function(rec) {
      if (!rec.staffEmail || !rec.assignmentDate) return;
      _schedulerAddAssignedIndex_(assignedByDateEmail, rec);
      fairness[rec.staffEmail] = (fairness[rec.staffEmail] || 0) + 1;
    });

    var generated = [];
    var unfilled = [];

    needs.forEach(function(need) {
      var needDateObj = new Date(need.assignmentDate + 'T00:00:00');
      var candidates = staffPool.filter(function(staff) {
        var email = _schedulerEmailKey_(staff.email);
        if (!email) return false;
        if (!_schedulerIsQualified_(staff, need, quals, needDateObj)) return false;
        if (_schedulerIsUnavailable_(availabilityMap, email, need.assignmentDate, need.shiftStartLocal, need.shiftEndLocal)) return false;
        if (_schedulerConflictsWithAssigned_(assignedByDateEmail, compatMap, email, need)) return false;
        return true;
      });

      candidates.sort(function(a, b) {
        var ea = _schedulerEmailKey_(a.email);
        var eb = _schedulerEmailKey_(b.email);
        var fa = fairness[ea] || 0;
        var fb = fairness[eb] || 0;
        if (fa !== fb) return fa - fb;
        return String(a.staffName || ea).localeCompare(String(b.staffName || eb));
      });

      if (!candidates.length) {
        unfilled.push({
          assignmentDate: need.assignmentDate,
          roleCode: need.roleCode,
          locationCode: need.locationCode,
          shiftStartLocal: need.shiftStartLocal,
          shiftEndLocal: need.shiftEndLocal,
          reason: 'No qualified/available staff for slot ' + need.slot
        });
        return;
      }

      var chosen = candidates[0];
      var chosenEmail = _schedulerEmailKey_(chosen.email);
      fairness[chosenEmail] = (fairness[chosenEmail] || 0) + 1;

      var assignment = {
        assignmentId: 'SASG-' + parsed.monthKey.replace('-', '') + '-' + String(generated.length + 1).padStart(4, '0'),
        scheduleMonth: parsed.monthKey,
        assignmentDate: need.assignmentDate,
        roleCode: need.roleCode,
        locationCode: need.locationCode,
        shiftStartLocal: need.shiftStartLocal,
        shiftEndLocal: need.shiftEndLocal,
        shiftStartZ: '',
        shiftEndZ: '',
        staffEmail: chosenEmail,
        source: 'AUTO_GENERATED',
        locked: 'N',
        lockScope: '',
        status: 'PROPOSED',
        fairnessScore: fairness[chosenEmail],
        ruleReason: 'priority=' + need.priority + ';slot=' + need.slot,
        updatedAt: new Date().toISOString(),
        updatedBy: actor.email
      };
      generated.push(assignment);
      _schedulerAddAssignedIndex_(assignedByDateEmail, assignment);
    });

    if (!dryRun) {
      rewriteSheetData_(assignmentsSheet, keepRows);
      if (generated.length) {
        var aHeaders = _toolsSheetHeaderRow_(assignmentsSheet);
        var append = generated.map(function(item) {
          var row = aHeaders.map(function() { return ''; });
          var idx = {
            assignmentId: _schedulerHeaderIndex_(aHeaders, 'ASSIGNMENT_ID'),
            month: _schedulerHeaderIndex_(aHeaders, 'SCHEDULE_MONTH'),
            date: _schedulerHeaderIndex_(aHeaders, 'ASSIGNMENT_DATE'),
            role: _schedulerHeaderIndex_(aHeaders, 'ROLE_CODE'),
            location: _schedulerHeaderIndex_(aHeaders, 'LOCATION_CODE'),
            startLocal: _schedulerHeaderIndex_(aHeaders, 'SHIFT_START_LOCAL'),
            endLocal: _schedulerHeaderIndex_(aHeaders, 'SHIFT_END_LOCAL'),
            startZ: _schedulerHeaderIndex_(aHeaders, 'SHIFT_START_Z'),
            endZ: _schedulerHeaderIndex_(aHeaders, 'SHIFT_END_Z'),
            staffEmail: _schedulerHeaderIndex_(aHeaders, 'STAFF_EMAIL'),
            source: _schedulerHeaderIndex_(aHeaders, 'SOURCE'),
            locked: _schedulerHeaderIndex_(aHeaders, 'LOCKED'),
            lockScope: _schedulerHeaderIndex_(aHeaders, 'LOCK_SCOPE'),
            status: _schedulerHeaderIndex_(aHeaders, 'STATUS'),
            fairness: _schedulerHeaderIndex_(aHeaders, 'FAIRNESS_SCORE'),
            reason: _schedulerHeaderIndex_(aHeaders, 'RULE_REASON'),
            updatedAt: _schedulerHeaderIndex_(aHeaders, 'UPDATED_AT'),
            updatedBy: _schedulerHeaderIndex_(aHeaders, 'UPDATED_BY')
          };
          if (idx.assignmentId >= 0) row[idx.assignmentId] = item.assignmentId;
          if (idx.month >= 0) row[idx.month] = item.scheduleMonth;
          if (idx.date >= 0) row[idx.date] = item.assignmentDate;
          if (idx.role >= 0) row[idx.role] = item.roleCode;
          if (idx.location >= 0) row[idx.location] = item.locationCode;
          if (idx.startLocal >= 0) row[idx.startLocal] = item.shiftStartLocal;
          if (idx.endLocal >= 0) row[idx.endLocal] = item.shiftEndLocal;
          if (idx.startZ >= 0) row[idx.startZ] = item.shiftStartZ;
          if (idx.endZ >= 0) row[idx.endZ] = item.shiftEndZ;
          if (idx.staffEmail >= 0) row[idx.staffEmail] = item.staffEmail;
          if (idx.source >= 0) row[idx.source] = item.source;
          if (idx.locked >= 0) row[idx.locked] = item.locked;
          if (idx.lockScope >= 0) row[idx.lockScope] = item.lockScope;
          if (idx.status >= 0) row[idx.status] = item.status;
          if (idx.fairness >= 0) row[idx.fairness] = item.fairnessScore;
          if (idx.reason >= 0) row[idx.reason] = item.ruleReason;
          if (idx.updatedAt >= 0) row[idx.updatedAt] = item.updatedAt;
          if (idx.updatedBy >= 0) row[idx.updatedBy] = item.updatedBy;
          return row;
        });
        var startRow = assignmentsSheet.getLastRow() + 1;
        assignmentsSheet.getRange(startRow, 1, append.length, aHeaders.length).setValues(append);
      }
    }

    return {
      success: true,
      dryRun: dryRun,
      month: parsed.monthKey,
      staffPoolSize: staffPool.length,
      needsCount: needs.length,
      generatedCount: generated.length,
      unfilledCount: unfilled.length,
      unfilled: unfilled,
      assignments: generated,
      diagnostics: diagnostics
    };
  } catch (e) {
    return {
      success: false,
      error: e && e.message ? e.message : String(e),
      stage: 'generateSchedulerAssignments',
      stack: e && e.stack ? String(e.stack) : ''
    };
  }
}

function previewSchedulerAssignments(payload) {
  var body = (payload && typeof payload === 'object') ? payload : {};
  body.dryRun = true;
  return generateSchedulerAssignments(body);
}

function getSchedulerAssignments(payload) {
  try {
    _schedulerAssertPermission_('CAN_VIEW', 'getSchedulerAssignments');
    var body = (payload && typeof payload === 'object') ? payload : {};
    var month = String(body.month || '').trim();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_ASSIGNMENTS || 'SCHED_Assignments', 'getSchedulerAssignments');
    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return { success: true, rows: [] };
    var headers = data[0] || [];
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

function runSchedulerTestSuite(payload) {
  var body = (payload && typeof payload === 'object') ? payload : {};
  var rawMonth = String(body.month || body.scheduleMonth || '').trim();
  var parsed = _schedulerParseMonthKey_(rawMonth || _schedulerMonthKey_(new Date()));
  if (!parsed) {
    return { success: false, error: 'Invalid month format. Use YYYY-MM.' };
  }

  var tests = [];
  var dryRunSummary = null;
  var unfilledPreview = [];

  function pushTest(id, name, status, message, details) {
    var row = {
      id: String(id || ''),
      name: String(name || id || ''),
      status: String(status || 'FAIL').toUpperCase(),
      message: String(message || '')
    };
    if (details && typeof details === 'object') row.details = details;
    tests.push(row);
  }

  try {
    var profileRes = getSchedulerAccessProfile();
    if (!profileRes || !profileRes.success) {
      pushTest('access_profile', 'Access profile', 'FAIL', (profileRes && profileRes.error) ? profileRes.error : 'Unable to load scheduler permissions.');
      return {
        success: false,
        month: parsed.monthKey,
        tests: tests,
        summary: { passCount: 0, warnCount: 0, failCount: 1, total: 1 }
      };
    }
    var permissions = profileRes.permissions || {};
    pushTest(
      'access_profile',
      'Access profile',
      'PASS',
      'Signed in as ' + String(profileRes.email || 'unknown') + ' (view=' + (permissions.CAN_VIEW ? 'Y' : 'N') + ', generate=' + (permissions.CAN_GENERATE ? 'Y' : 'N') + ')',
      { email: profileRes.email, permissions: permissions }
    );

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var requiredSheets = [
      APP_SHEETS.SCHED_CONFIG || 'SCHED_Config',
      APP_SHEETS.SCHED_PERMISSIONS || 'SCHED_Permissions',
      APP_SHEETS.SCHED_COVERAGE_RULES || 'SCHED_Coverage_Requirements',
      APP_SHEETS.SCHED_ROLE_COMPAT || 'SCHED_Role_Compatibility',
      APP_SHEETS.SCHED_STAFF_QUALS || 'SCHED_Staff_Qualifications',
      APP_SHEETS.SCHED_ASSIGNMENTS || 'SCHED_Assignments',
      APP_SHEETS.SCHED_LOCKS || 'SCHED_Assignment_Locks',
      APP_SHEETS.SCHED_ALERTS || 'SCHED_Alerts',
      APP_SHEETS.SCHED_PUBLISH_LOG || 'SCHED_Publish_Log',
      APP_SHEETS.SCHED_AVAILABILITY || 'SCHED_Staff_Availability'
    ];
    var missingSheets = [];
    for (var s = 0; s < requiredSheets.length; s++) {
      try {
        getRequiredSheet_(ss, requiredSheets[s], 'runSchedulerTestSuite');
      } catch (sheetErr) {
        missingSheets.push(requiredSheets[s]);
      }
    }
    if (missingSheets.length) {
      pushTest('schema_sheets', 'Schema sheets', 'FAIL', 'Missing sheets: ' + missingSheets.join(', '), { missingSheets: missingSheets });
    } else {
      pushTest('schema_sheets', 'Schema sheets', 'PASS', 'All scheduler sheets are present.');
    }

    var cfgRes = getSchedulerConfig();
    if (!cfgRes || !cfgRes.success) {
      pushTest('config_read', 'Config readability', 'FAIL', (cfgRes && cfgRes.error) ? cfgRes.error : 'Could not read scheduler config.');
    } else {
      var cfg = cfgRes.config || {};
      var publishDay = Number(cfg.PUBLISH_DAY_OF_MONTH || 0);
      if (!(publishDay >= 1 && publishDay <= 31)) {
        pushTest('config_publish_day', 'Publish cadence', 'FAIL', 'PUBLISH_DAY_OF_MONTH must be between 1 and 31.', { value: cfg.PUBLISH_DAY_OF_MONTH || '' });
      } else {
        pushTest('config_publish_day', 'Publish cadence', 'PASS', 'PUBLISH_DAY_OF_MONTH=' + publishDay + '.');
      }

      if (String(cfg.AVAILABILITY_CALENDAR_ID || '').trim()) {
        pushTest('config_availability_calendar', 'Availability calendar config', 'PASS', 'AVAILABILITY_CALENDAR_ID is configured.');
      } else {
        pushTest('config_availability_calendar', 'Availability calendar config', 'WARN', 'AVAILABILITY_CALENDAR_ID is blank; calendar sync cannot run.');
      }
    }

    var staffPool = [];
    try {
      staffPool = _schedulerBuildStaffPool_();
      if (staffPool.length) {
        pushTest('staff_pool', 'Staff pool', 'PASS', 'Active staff available: ' + staffPool.length + '.');
      } else {
        pushTest('staff_pool', 'Staff pool', 'FAIL', 'No active staff rows found in DB_Pilots.');
      }
    } catch (staffErr) {
      pushTest('staff_pool', 'Staff pool', 'FAIL', staffErr && staffErr.message ? staffErr.message : String(staffErr));
    }

    try {
      var qualRows = _schedulerQualificationRows_();
      if (qualRows.length) {
        pushTest('qualifications', 'Qualifications', 'PASS', 'Active qualification rows: ' + qualRows.length + '.');
      } else {
        pushTest('qualifications', 'Qualifications', 'FAIL', 'No active rows in SCHED_Staff_Qualifications.');
      }
    } catch (qualErr) {
      pushTest('qualifications', 'Qualifications', 'FAIL', qualErr && qualErr.message ? qualErr.message : String(qualErr));
    }

    try {
      var needs = _schedulerCoverageNeeds_(parsed.monthKey);
      if (needs.length) {
        pushTest('coverage_needs', 'Coverage requirements', 'PASS', 'Coverage slots for ' + parsed.monthKey + ': ' + needs.length + '.');
      } else {
        pushTest('coverage_needs', 'Coverage requirements', 'FAIL', 'No active coverage slots for ' + parsed.monthKey + '.');
      }
    } catch (needsErr) {
      pushTest('coverage_needs', 'Coverage requirements', 'FAIL', needsErr && needsErr.message ? needsErr.message : String(needsErr));
    }

    if (body.syncAvailabilityFirst) {
      var syncRes = syncSchedulerAvailabilityForMonth({ month: parsed.monthKey });
      if (syncRes && syncRes.success) {
        var syncStats = syncRes.stats || {};
        pushTest('availability_sync', 'Availability sync', 'PASS',
          'Synced month ' + parsed.monthKey
          + ' (events=' + Number(syncStats.eventsRead || 0)
          + ', imported=' + Number(syncStats.imported || 0)
          + ', skippedNoEmail=' + Number(syncStats.skippedNoEmail || 0)
          + ').',
          syncStats
        );
      } else {
        pushTest('availability_sync', 'Availability sync', 'FAIL', (syncRes && syncRes.error) ? syncRes.error : 'Availability sync failed.');
      }
    } else {
      var availRes = getSchedulerAvailability({ month: parsed.monthKey });
      if (availRes && availRes.success) {
        var availCount = Array.isArray(availRes.rows) ? availRes.rows.length : 0;
        if (availCount > 0) {
          pushTest('availability_rows', 'Availability rows', 'PASS', 'Rows for ' + parsed.monthKey + ': ' + availCount + '.');
        } else {
          pushTest('availability_rows', 'Availability rows', 'WARN', 'No availability rows for ' + parsed.monthKey + '. Scheduler can still run.');
        }
      } else {
        pushTest('availability_rows', 'Availability rows', 'FAIL', (availRes && availRes.error) ? availRes.error : 'Could not read availability rows.');
      }
    }

    var dryRes = previewSchedulerAssignments({ month: parsed.monthKey });
    if (!dryRes || !dryRes.success) {
      pushTest('dry_run', 'Dry run generation', 'FAIL', (dryRes && dryRes.error) ? dryRes.error : 'Dry run failed.');
    } else {
      dryRunSummary = {
        staffPoolSize: Number(dryRes.staffPoolSize || 0),
        needsCount: Number(dryRes.needsCount || 0),
        generatedCount: Number(dryRes.generatedCount || 0),
        unfilledCount: Number(dryRes.unfilledCount || 0),
        diagnostics: dryRes.diagnostics || {}
      };
      unfilledPreview = Array.isArray(dryRes.unfilled) ? dryRes.unfilled.slice(0, 50) : [];
      if (Number(dryRes.unfilledCount || 0) > 0) {
        pushTest('dry_run', 'Dry run generation', 'WARN',
          'Dry run completed with ' + Number(dryRes.unfilledCount || 0) + ' unfilled slot(s).',
          dryRunSummary
        );
      } else {
        pushTest('dry_run', 'Dry run generation', 'PASS',
          'Dry run completed with full coverage (' + Number(dryRes.generatedCount || 0) + ' assignments).',
          dryRunSummary
        );
      }
    }

    var assignRes = getSchedulerAssignments({ month: parsed.monthKey });
    if (!assignRes || !assignRes.success) {
      pushTest('assignments_read', 'Assignments readability', 'FAIL', (assignRes && assignRes.error) ? assignRes.error : 'Could not load assignment rows.');
    } else {
      var rowCount = Array.isArray(assignRes.rows) ? assignRes.rows.length : 0;
      pushTest('assignments_read', 'Assignments readability', 'PASS', 'Existing assignment rows for ' + parsed.monthKey + ': ' + rowCount + '.');
    }
  } catch (e) {
    pushTest('unexpected_error', 'Unexpected error', 'FAIL', e && e.message ? e.message : String(e));
  }

  var passCount = 0;
  var warnCount = 0;
  var failCount = 0;
  tests.forEach(function(t) {
    var st = String(t.status || '').toUpperCase();
    if (st === 'PASS') passCount++;
    else if (st === 'WARN') warnCount++;
    else failCount++;
  });

  return {
    success: failCount === 0,
    month: parsed.monthKey,
    tests: tests,
    summary: {
      passCount: passCount,
      warnCount: warnCount,
      failCount: failCount,
      total: tests.length
    },
    dryRun: dryRunSummary,
    unfilledPreview: unfilledPreview
  };
}
