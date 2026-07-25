function getToolsBuildVersion() {
  return TOOLS_BUILD_VERSION;
}

function mapAirportRowsShared_(rowsObj) {
const rows = rowsObj && Array.isArray(rowsObj.vals) ? rowsObj.vals : [];
const headers = rowsObj && Array.isArray(rowsObj.headers) ? rowsObj.headers : [];
return rows.map(function(r) {
  return _airportMapRow_(headers, r);
});
}

function getCalendarEvents() {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const dispSheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, "getCalendarEvents");
let events = [];
const data = dispSheet.getDataRange().getValues();
let missions = {};
const maintenanceSnapshot_ = getDispatchMaintenanceStatus({});
const maintenanceByAircraft_ = maintenanceSnapshot_ && maintenanceSnapshot_.success && maintenanceSnapshot_.byAircraft
  ? maintenanceSnapshot_.byAircraft
  : {};

const flownStatuses_ = { FLOWN: true, COMPLETE: true, COMPLETED: true };
const completedLegMap_ = {};
const departedLegMap_ = {};
const dispatchLegStateMap_ = {};

const calendarDateToDateObj_ = function(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return new Date(value.getTime());
  if (typeof value === 'number' && isFinite(value)) {
    var serialBase = new Date(1899, 11, 30);
    var wholeDays = Math.floor(value);
    var millis = Math.round((value - wholeDays) * 24 * 60 * 60 * 1000);
    var serialDate = new Date(serialBase.getTime() + (wholeDays * 24 * 60 * 60 * 1000) + millis);
    if (!isNaN(serialDate.getTime())) return serialDate;
  }
  var raw = String(value == null ? '' : value).trim();
  if (!raw) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    var ymd = new Date(raw + 'T00:00:00');
    return isNaN(ymd.getTime()) ? null : ymd;
  }
  var br = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (br) {
    var dd = Number(br[1]);
    var mm = Number(br[2]);
    var yyyy = Number(br[3]);
    var brDate = new Date(yyyy, mm - 1, dd, 0, 0, 0, 0);
    return isNaN(brDate.getTime()) ? null : brDate;
  }
  var parsed = new Date(raw);
  return isNaN(parsed.getTime()) ? null : parsed;
};

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

for (let i = 1; i < data.length; i++) {
 const row = data[i];
 const mId = row[DISPATCH_COL.MISSION_ID];
 if (!mId) continue;
  if (!missions[mId]) {
     const flightDescription = _extractMissionDescriptionFromRaw_(row[DISPATCH_COL.RAW_DATA]);
     const communicationRequests = _extractMissionCommunicationsFromRaw_(row[DISPATCH_COL.RAW_DATA]);
       const mxReview = _extractMissionMxReviewFromRaw_(row[DISPATCH_COL.RAW_DATA]);
   missions[mId] = {
     id: mId,
     date: row[DISPATCH_COL.DATE],
     acft: row[DISPATCH_COL.AIRCRAFT],
     pilot: row[DISPATCH_COL.PILOT],
     status: row[DISPATCH_COL.STATUS],
       flightDescription: flightDescription,
       communicationRequests: communicationRequests,
         mxReview: mxReview,
     legs: []
   };
 }

 // --- THE FIX: Read Route from Column G (Index 6) ---
 const route = splitRoute_(row[DISPATCH_COL.ROUTE]);
 const from = route.from || "?";
 const to = route.to || "?";
  const legTime = safeNumber_(row[DISPATCH_COL.FLIGHT_TIME], 0);

 missions[mId].legs.push({
   flightLegId: String(row[DISPATCH_COL.FLIGHT_ID] || '').trim(),
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
 let dateObj = calendarDateToDateObj_(m.date);
 if (!dateObj || isNaN(dateObj.getTime())) return;
   let status = m.status ? m.status.toString().toUpperCase() : "PENDING";
   const allLegsComplete = m.legs.length > 0 && m.legs.every(function(leg) {
     return leg.flightLegId && completedLegMap_[String(leg.flightLegId)];
   });
   const isFlown = !!flownStatuses_[status] || allLegsComplete;
   if (isFlown) status = 'FLOWN';
 let color = "#f57c00"; // Orange
   if (status === "APPROVED") color = "#43a047"; // Green
   if (status === "FLOWN") color = "#1565c0"; // Blue
 if (status === "CANCELLED") color = "#b0bec5"; // Grey

 const missionReg = String(m.acft || '').trim().toUpperCase();
 const missionMaint = maintenanceByAircraft_[missionReg] && typeof maintenanceByAircraft_[missionReg] === 'object' ? maintenanceByAircraft_[missionReg] : null;
 const missionNextInspection = missionMaint && missionMaint.nextInspection && typeof missionMaint.nextInspection === 'object'
   ? missionMaint.nextInspection
   : null;
 const missionCurrentHoursToInsp = missionNextInspection ? Number(missionNextInspection.hoursRemaining) : null;
 const missionProjectedHoursToInsp = isFinite(missionCurrentHoursToInsp) ? Number((missionCurrentHoursToInsp - totalFlt).toFixed(1)) : '';

 events.push({
   start: dateObj.toISOString().split('T')[0],
   allDay: true,
   title: String(m.acft || '').trim() + ' | ' + routeDisplay + ' | ' + String(m.pilot || '').trim(),
   display: 'block',
   color: color,
   extendedProps: {
     type: 'mission',
     id: m.id,
     status: status,
     acft: m.acft,
     acftMaintenanceLabel: (function() {
       var reg = String(m.acft || '').trim().toUpperCase();
       var live = maintenanceByAircraft_[reg] && typeof maintenanceByAircraft_[reg] === 'object' ? maintenanceByAircraft_[reg] : null;
       var review = (m.mxReview && typeof m.mxReview === 'object') ? m.mxReview : {};
       var nextInspection = live && live.nextInspection && typeof live.nextInspection === 'object'
         ? live.nextInspection
         : (review.nextInspection && typeof review.nextInspection === 'object' ? review.nextInspection : null);
       var hoursRemaining = nextInspection ? Number(nextInspection.hoursRemaining) : null;
       if (!Number.isFinite(hoursRemaining)) return String(m.acft || '').trim();
       return String(m.acft || '').trim() + ' (' + hoursRemaining.toFixed(1) + ' hrs)';
     })(),
     pilot: String(m.pilot || '').trim() ? String(m.pilot).trim().split(' ')[0] : 'PILOT TBD',
     fullPilot: String(m.pilot || '').trim() || 'PILOT TBD',
     route: routeDisplay, // Correct String now (e.g. "SDRM - SBBV")
     flightDescription: String(m.flightDescription || '').trim(),
     communicationRequests: Array.isArray(m.communicationRequests) ? m.communicationRequests : [],
     mxReview: (m.mxReview && typeof m.mxReview === 'object') ? m.mxReview : {},
     takeoff: "08:00",
     fltTime: totalFlt.toFixed(1),
     flightHoursDisplay: totalFlt.toFixed(1),
     flightHoursLabel: isFlown ? 'FLOWN' : 'EST',
     isPastDate: dateObj.getTime() < (new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate())).getTime(),
    maintenanceCurrentHoursToInsp: Number.isFinite(missionCurrentHoursToInsp) ? Number(missionCurrentHoursToInsp.toFixed(1)) : '',
     maintenanceProjectedHoursToInsp: missionProjectedHoursToInsp,
     maintenanceNextInspectionDate: missionNextInspection && missionNextInspection.nextDueDate ? String(missionNextInspection.nextDueDate || '') : '',
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

try {
  const maintRes = getMaintenanceScheduleData({});
  if (maintRes && maintRes.success && Array.isArray(maintRes.events)) {
    maintRes.events.forEach(function(ev) {
      if (!ev || !ev.start) return;
      events.push({
        id: 'mx-' + String(ev.id || ev.assignmentId || ev.title || '').trim(),
        title: String(ev.title || ev.taskName || ev.taskCode || 'Maintenance'),
        start: ev.start,
        end: ev.end || '',
        allDay: true,
        color: ev.status === 'OVERDUE' ? '#d32f2f' : (ev.status === 'DUE_SOON' ? '#ef6c00' : '#1b5e20'),
        extendedProps: {
          type: 'maintenance',
          assignmentId: String(ev.id || ev.assignmentId || '').trim(),
          status: ev.status || 'OK',
          aircraftReg: String(ev.aircraftReg || ''),
          taskName: String(ev.taskName || ''),
          taskCode: String(ev.taskCode || ''),
          dueDate: String(ev.dueDate || ev.nextDueDate || ev.start || ''),
          dueMode: String(ev.dueMode || ''),
          intervalHours: ev.intervalHours,
          intervalDays: ev.intervalDays,
          intervalMonths: ev.intervalMonths,
          intervalYears: ev.intervalYears,
          lastPerformedDate: String(ev.lastPerformedDate || ''),
          lastPerformedTach: ev.lastPerformedTach,
          notes: String(ev.notes || ''),
          source: 'maintenance'
        }
      });
    });
  }
} catch (e) { /* maintenance events are non-critical */ }

return events;
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

function verifySupervisorApprovalPassword(password) {
  _verifySupervisorApprovalPassword_(password);
  return { success: true };
}

function waiveDestinationCheck(pilot, icao, missionId, approvalPassword) {
_verifySupervisorApprovalPassword_(approvalPassword);
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = _pilotRunwayChecksSheet_();
const headers = _toolsSheetHeaderRow_(sheet);
const data = sheet.getDataRange().getValues();
const user = "Admin";
const pilotName = String(pilot || '').trim();
const normalizedPilot = pilotName.toUpperCase();
const normalizedIcao = _checksNormalizeIcaoToken_(icao);
if (!pilotName || !normalizedIcao) return "Error: pilot and destination are required";

for (let i = 1; i < data.length; i++) {
  const rec = _pilotRunwayCheckRecordFromRow_(headers, data[i], i + 1);
  if (!rec.pilotName || String(rec.pilotName || '').trim().toUpperCase() !== normalizedPilot) continue;
  if (rec.icaos.indexOf(normalizedIcao) === -1) continue;
  if (rec.isActive) return "Check Waived";
}

let staffRecord = null;
const staffSheet = ss.getSheetByName(APP_SHEETS.PILOTS);
if (staffSheet && staffSheet.getLastRow() > 1) {
  const staffData = staffSheet.getDataRange().getValues();
  const staffHeaders = staffData[0] || [];
  for (let s = 1; s < staffData.length; s++) {
    const staff = _toolsStaffRecordFromRow_(staffHeaders, staffData[s], s + 1);
    if (String(staff.staffName || '').trim().toUpperCase() === normalizedPilot) {
      staffRecord = staff;
      break;
    }
  }
}

const now = safeDateStr(new Date());
const dataMap = {
  CHECK_ID: 'CHK_' + new Date().getTime() + '_' + normalizedIcao,
  PILOT_NAME: pilotName,
  PILOT_EMAIL: staffRecord ? String(staffRecord.email || '').trim().toLowerCase() : '',
  STAFF_ID: staffRecord ? String(staffRecord.staffId || '').trim() : '',
  ICAO: normalizedIcao,
  RUNWAY_IDENT: 'ALL',
  AUTH_SCOPE: 'AIRPORT',
  STATUS: 'ACTIVE',
  DATE_CHECKED: now,
  EXPIRY_DATE: '',
  APPROVED_BY: user,
  SOURCE: 'SUPERVISOR_WAIVER',
  NOTES: missionId ? ('Mission ' + missionId) : 'Supervisor waiver',
  CREATED_AT: now,
  UPDATED_AT: now
};
const row = headers.map(function(header) {
  var key = _toolsNormHeader_(header);
  return Object.prototype.hasOwnProperty.call(dataMap, key) ? dataMap[key] : '';
});
sheet.appendRow(row);

const audit = ss.getSheetByName(APP_SHEETS.AUDIT);
if(audit) audit.appendRow([new Date(), user, missionId, "WAIVE_CHECK", '', normalizedIcao, normalizedIcao]);
return "Check Waived";
}

function waiveDestinationChecksBatch(pilot, icaos, missionId, approvalPassword) {
  _verifySupervisorApprovalPassword_(approvalPassword);
  var list = Array.isArray(icaos)
    ? icaos
    : String(icaos || '').split(/[;,]+/);
  var normalized = list
    .map(function(v) { return _checksNormalizeIcaoToken_(v); })
    .filter(Boolean)
    .filter(function(code, idx, arr) { return arr.indexOf(code) === idx; });

  var results = [];
  normalized.forEach(function(code) {
    try {
      results.push({ icao: code, success: true, message: String(waiveDestinationCheck(pilot, code, missionId, approvalPassword) || '') });
    } catch (e) {
      results.push({ icao: code, success: false, error: e && e.message ? e.message : String(e) });
    }
  });
  return { success: true, results: results };
}

function getAuthorizedDestinations(pilotName) {
const sheet = _pilotRunwayChecksSheet_();
if (!sheet || sheet.getLastRow() < 2) return "";
const data = sheet.getDataRange().getValues();
const headers = data[0] || [];
const target = String(pilotName || '').trim().toUpperCase();
const out = [];
const seen = {};
for (let i = 1; i < data.length; i++) {
  const rec = _pilotRunwayCheckRecordFromRow_(headers, data[i], i + 1);
  if (!rec.pilotName || String(rec.pilotName || '').trim().toUpperCase() !== target) continue;
  if (!rec.isActive) continue;
  rec.icaos.forEach(function(code) {
    if (!seen[code]) {
      seen[code] = true;
      out.push(code);
    }
  });
}
return out.join(", ");
}

function ingestForm1ResponseRow(rowNumber, sheetName) {
  var ss = SpreadsheetApp.openById(APP_SHEETS.FORM1_SPREADSHEET_ID);
  var sourceSheet = _formGetSourceSheetByNameOrFirst_(ss, String(sheetName || APP_SHEETS.FORM1_RESPONSES || 'Form Responses 1'));
  if (!sourceSheet) throw new Error('Form 1 response sheet not found');
  var rn = parseInt(rowNumber, 10);
  if (!(rn > 1)) rn = sourceSheet.getLastRow();
  if (!(rn > 1)) throw new Error('No data rows found in Form 1 response sheet');
  var lastCol = sourceSheet.getLastColumn();
  var headers = sourceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var row = sourceSheet.getRange(rn, 1, 1, lastCol).getValues()[0];
  var payload = _formBuildForm1Payload_(headers, row, sourceSheet.getName(), rn);
  return _ingestForm1Payload_(payload);
}

function ingestForm1ResponsesBatch(options) {
  var opts = (options && typeof options === 'object') ? options : {};
  var sheetName = String(opts.sheetName || APP_SHEETS.FORM1_RESPONSES || 'Form Responses 1');
  var startRow = parseInt(opts.startRow, 10);
  var maxRows = parseInt(opts.maxRows, 10);
  if (!(startRow > 1)) startRow = 2;
  if (!(maxRows > 0)) maxRows = 200;

  var ss = SpreadsheetApp.openById(APP_SHEETS.FORM1_SPREADSHEET_ID);
  var sourceSheet = _formGetSourceSheetByNameOrFirst_(ss, sheetName);
  if (!sourceSheet) throw new Error('Form 1 response sheet not found: ' + sheetName);
  var lastRow = sourceSheet.getLastRow();
  var lastCol = sourceSheet.getLastColumn();
  if (lastRow < startRow) return { success: true, processed: 0, imported: 0, failed: 0, skipped: 0 };

  var endRow = Math.min(lastRow, startRow + maxRows - 1);
  var headers = sourceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var rows = sourceSheet.getRange(startRow, 1, endRow - startRow + 1, lastCol).getValues();
  var out = { success: true, processed: 0, imported: 0, failed: 0, skipped: 0, details: [] };

  for (var i = 0; i < rows.length; i++) {
    var rowNumber = startRow + i;
    out.processed += 1;
    try {
      var payload = _formBuildForm1Payload_(headers, rows[i], sourceSheet.getName(), rowNumber);
      var res = _ingestForm1Payload_(payload);
      if (res && res.skipped) out.skipped += 1;
      else if (res && res.success) out.imported += 1;
      else out.failed += 1;
      out.details.push({ row: rowNumber, result: res });
    } catch (err) {
      out.failed += 1;
      out.details.push({ row: rowNumber, error: err && err.message ? err.message : String(err) });
    }

  }

  return out;
}

function ingestForm1OnSubmit(e) {
  if (!e || !e.range) throw new Error('ingestForm1OnSubmit requires an installable on form submit event object');
  var sheet = e.range.getSheet();
  var rowNumber = e.range.getRow();
  return ingestForm1ResponseRow(rowNumber, sheet.getName());
}

function ingestForm2ResponseRow(rowNumber, sheetName) {
  var ss = APP_SHEETS.FORM2_SPREADSHEET_ID ? SpreadsheetApp.openById(APP_SHEETS.FORM2_SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = _formGetSourceSheetByNameOrFirst_(ss, String(sheetName || APP_SHEETS.FORM2_RESPONSES || 'Form Responses 2'));
  if (!sourceSheet) throw new Error('Form 2 response sheet not found');
  var rn = parseInt(rowNumber, 10);
  if (!(rn > 1)) rn = sourceSheet.getLastRow();
  if (!(rn > 1)) throw new Error('No data rows found in Form 2 response sheet');
  var lastCol = sourceSheet.getLastColumn();
  var headers = sourceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var row = sourceSheet.getRange(rn, 1, 1, lastCol).getValues()[0];
  var payload = _formBuildForm2Payload_(headers, row, sourceSheet.getName(), rn);
  return _ingestForm2Payload_(payload);
}

function ingestForm2ResponsesBatch(options) {
  var opts = (options && typeof options === 'object') ? options : {};
  var sheetName = String(opts.sheetName || APP_SHEETS.FORM2_RESPONSES || 'Form Responses 2');
  var startRow = parseInt(opts.startRow, 10);
  var maxRows = parseInt(opts.maxRows, 10);
  if (!(startRow > 1)) startRow = 2;
  if (!(maxRows > 0)) maxRows = 200;

  var ss = APP_SHEETS.FORM2_SPREADSHEET_ID ? SpreadsheetApp.openById(APP_SHEETS.FORM2_SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = _formGetSourceSheetByNameOrFirst_(ss, sheetName);
  if (!sourceSheet) throw new Error('Form 2 response sheet not found: ' + sheetName);
  var lastRow = sourceSheet.getLastRow();
  var lastCol = sourceSheet.getLastColumn();
  if (lastRow < startRow) return { success: true, processed: 0, imported: 0, failed: 0, skipped: 0 };

  var endRow = Math.min(lastRow, startRow + maxRows - 1);
  var headers = sourceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var rows = sourceSheet.getRange(startRow, 1, endRow - startRow + 1, lastCol).getValues();
  var out = { success: true, processed: 0, imported: 0, failed: 0, skipped: 0, details: [] };

  for (var i = 0; i < rows.length; i++) {
    var rowNumber = startRow + i;
    out.processed += 1;
    try {
      var payload = _formBuildForm2Payload_(headers, rows[i], sourceSheet.getName(), rowNumber);
      var res = _ingestForm2Payload_(payload);
      if (res && res.skipped) out.skipped += 1;
      else if (res && res.success) out.imported += 1;
      else out.failed += 1;
      out.details.push({ row: rowNumber, result: res });
    } catch (err) {
      out.failed += 1;
      out.details.push({ row: rowNumber, error: err && err.message ? err.message : String(err) });
    }
  }

  return out;
}

function ingestForm2OnSubmit(e) {
  if (!e || !e.range) throw new Error('ingestForm2OnSubmit requires an installable on form submit event object');
  var sheet = e.range.getSheet();
  var rowNumber = e.range.getRow();
  return ingestForm2ResponseRow(rowNumber, sheet.getName());
}

function ingestForm3ResponseRow(rowNumber, sheetName) {
  var ss = APP_SHEETS.FORM3_SPREADSHEET_ID ? SpreadsheetApp.openById(APP_SHEETS.FORM3_SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = _formGetSourceSheetByNameOrFirst_(ss, String(sheetName || APP_SHEETS.FORM3_RESPONSES || 'Form Responses 3'));
  if (!sourceSheet) throw new Error('Form 3 response sheet not found');
  var rn = parseInt(rowNumber, 10);
  if (!(rn > 1)) rn = sourceSheet.getLastRow();
  if (!(rn > 1)) throw new Error('No data rows found in Form 3 response sheet');
  var lastCol = sourceSheet.getLastColumn();
  var headers = sourceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var row = sourceSheet.getRange(rn, 1, 1, lastCol).getValues()[0];
  var payload = _formBuildForm3Payload_(headers, row, sourceSheet.getName(), rn);
  return _ingestForm3Payload_(payload);
}

function ingestForm3ResponsesBatch(options) {
  var opts = (options && typeof options === 'object') ? options : {};
  var sheetName = String(opts.sheetName || APP_SHEETS.FORM3_RESPONSES || 'Form Responses 3');
  var startRow = parseInt(opts.startRow, 10);
  var maxRows = parseInt(opts.maxRows, 10);
  if (!(startRow > 1)) startRow = 2;
  if (!(maxRows > 0)) maxRows = 200;

  var ss = APP_SHEETS.FORM3_SPREADSHEET_ID ? SpreadsheetApp.openById(APP_SHEETS.FORM3_SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = _formGetSourceSheetByNameOrFirst_(ss, sheetName);
  if (!sourceSheet) throw new Error('Form 3 response sheet not found: ' + sheetName);
  var lastRow = sourceSheet.getLastRow();
  var lastCol = sourceSheet.getLastColumn();
  if (lastRow < startRow) return { success: true, processed: 0, imported: 0, failed: 0, skipped: 0 };

  var endRow = Math.min(lastRow, startRow + maxRows - 1);
  var headers = sourceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var rows = sourceSheet.getRange(startRow, 1, endRow - startRow + 1, lastCol).getValues();
  var out = { success: true, processed: 0, imported: 0, failed: 0, skipped: 0, details: [] };

  for (var i = 0; i < rows.length; i++) {
    var rowNumber = startRow + i;
    out.processed += 1;
    try {
      var payload = _formBuildForm3Payload_(headers, rows[i], sourceSheet.getName(), rowNumber);
      var res = _ingestForm3Payload_(payload);
      if (res && res.skipped) out.skipped += 1;
      else if (res && res.success) out.imported += 1;
      else out.failed += 1;
      out.details.push({ row: rowNumber, result: res });
    } catch (err) {
      out.failed += 1;
      out.details.push({ row: rowNumber, error: err && err.message ? err.message : String(err) });
    }
  }

  return out;
}

function ingestForm3OnSubmit(e) {
  if (!e || !e.range) throw new Error('ingestForm3OnSubmit requires an installable on form submit event object');
  var sheet = e.range.getSheet();
  var rowNumber = e.range.getRow();
  return ingestForm3ResponseRow(rowNumber, sheet.getName());
}

function ingestAnyFormOnSubmit(e) {
  if (!e || !e.range) throw new Error('ingestAnyFormOnSubmit requires an installable on form submit event object');
  var sheet = e.range.getSheet();
  if (!sheet) throw new Error('Could not resolve source sheet from submit event');
  var rowNumber = e.range.getRow();
  var sheetName = String(sheet.getName() || '').trim();

  if (sheetName === String(APP_SHEETS.FORM1_RESPONSES || 'Form Responses 1')) {
    return ingestForm1ResponseRow(rowNumber, sheetName);
  }
  if (sheetName === String(APP_SHEETS.FORM2_RESPONSES || 'Form Responses 2')) {
    return ingestForm2ResponseRow(rowNumber, sheetName);
  }
  if (sheetName === String(APP_SHEETS.FORM3_RESPONSES || 'Form Responses 3')) {
    return ingestForm3ResponseRow(rowNumber, sheetName);
  }

  return {
    success: true,
    skipped: true,
    reason: 'unmapped_sheet',
    sheetName: sheetName,
    rowNumber: rowNumber
  };
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

function getToolsSyllabusRecords() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = APP_SHEETS.SYLLABUS || 'Plano de Aula';
    var sh = getRequiredSheet_(ss, sheetName, 'getToolsSyllabusRecords');
    var headerRow = _toolsSheetHeaderRow_(sh);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) {
      return { success: true, records: [] };
    }
    var width = Math.max(1, sh.getLastColumn());
    var data = sh.getRange(2, 1, lastRow - 1, width).getValues();
    var records = data.map(function(row, idx) {
      var payload = _toolsRowPayloadFromHeaders_(headerRow, row);
      return {
        code: String(payload[_toolsNormHeader_('TRAINING_CODE')] || '').trim(),
        description: String(payload[_toolsNormHeader_('DESCRIPTION')] || '').trim(),
        rowNumber: idx + 2
      };
    }).filter(function(r) { return !!r.code; });
    return { success: true, records: records };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e), records: [] };
  }
}

function getToolsSheetRecordByKey(kind, keyHeader, keyValue) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = _toolsSheetNameFromKind_(kind);
    var sh = getRequiredSheet_(ss, sheetName, 'getToolsSheetRecordByKey');
    var headerRow = _toolsSheetHeaderRow_(sh);
    var keyNorm = _toolsNormHeader_(keyHeader);
    var keyIdx = -1;
    for (var i = 0; i < headerRow.length; i++) {
      if (_toolsNormHeader_(headerRow[i]) === keyNorm) {
        keyIdx = i;
        break;
      }
    }
    if (keyIdx < 0) {
      return {
        success: false,
        found: false,
        error: 'Header not found: ' + String(keyHeader || ''),
        sheetName: sheetName
      };
    }

    var target = _toolsNormalizeKeyValue_(keyValue);
    if (!target) {
      return {
        success: true,
        found: false,
        sheetName: sheetName
      };
    }

    var lastRow = sh.getLastRow();
    if (lastRow < 2) {
      return {
        success: true,
        found: false,
        sheetName: sheetName
      };
    }

    var width = Math.max(1, sh.getLastColumn());
    var data = sh.getRange(2, 1, lastRow - 1, width).getValues();
    for (var r = 0; r < data.length; r++) {
      var candidate = _toolsNormalizeKeyValue_(data[r][keyIdx]);
      if (candidate !== target) continue;
      return {
        success: true,
        found: true,
        sheetName: sheetName,
        rowNumber: r + 2,
        payload: _toolsRowPayloadFromHeaders_(headerRow, data[r])
      };
    }

    return {
      success: true,
      found: false,
      sheetName: sheetName
    };
  } catch (e) {
    return {
      success: false,
      found: false,
      error: e && e.message ? e.message : String(e)
    };
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

function getToolsAircraftBuilderTemplate(sourceRegistration) {
  try {
    var sourceReg = String(sourceRegistration || '').trim();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aircraftSheet = getRequiredSheet_(ss, APP_SHEETS.AIRCRAFT, 'getToolsAircraftBuilderTemplate');
    var airframesSheet = getRequiredSheet_(ss, APP_SHEETS.AIRFRAMES, 'getToolsAircraftBuilderTemplate');
    var envelopesSheet = getRequiredSheet_(ss, APP_SHEETS.ENVELOPES, 'getToolsAircraftBuilderTemplate');
    var rollSheet = getRequiredSheet_(ss, 'Aircraft_Roll_Numbers', 'getToolsAircraftBuilderTemplate');

    var aircraftHeaders = _toolsSheetHeaderRow_(aircraftSheet);
    var aircraftPromptHeaders = ['MTOW_LAND', 'MTOW_WATER', 'MLW_LAND', 'MLW_WATER', 'POSICAO_AMBOS'];
    var aircraftHeadersForTool = aircraftHeaders.slice();
    var aircraftHeaderNormSet = {};
    aircraftHeaders.forEach(function(h) {
      aircraftHeaderNormSet[_toolsNormHeader_(h)] = true;
    });
    aircraftPromptHeaders.forEach(function(h) {
      if (!aircraftHeaderNormSet[_toolsNormHeader_(h)]) aircraftHeadersForTool.push(h);
    });
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
      var aircraftTypeRaw = typeIdx >= 0 ? String(row[typeIdx] || '').trim() : '';
      aircraftOptions.push({
        registration: reg,
        aircraftType: _toolsCanonicalAircraftType_(aircraftTypeRaw),
        aircraftTypeRaw: aircraftTypeRaw
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
        var sourceTypeRaw = '';
        if (typeIdx >= 0) sourceTypeRaw = String(sourceAircraftRow[typeIdx] || '').trim();
        var sourceType = _toolsCanonicalAircraftType_(sourceTypeRaw);
        if (sourceType && typeIdx >= 0) {
          template.aircraftRow[_toolsNormHeader_(aircraftHeaders[typeIdx])] = sourceType;
        }

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
      sourceRegistration: sourceReg,
      sourceAircraftType: sourceType,
      aircraftOptions: aircraftOptions,
      sections: {
        aircraft: aircraftHeadersForTool.map(function(h) { return { label: String(h || '').trim(), key: _toolsNormHeader_(h) }; }),
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
    var originalRegistration = String(data.originalRegistration || '').trim();
    var originalAircraftType = String(data.originalAircraftType || '').trim();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aircraftSheet = getRequiredSheet_(ss, APP_SHEETS.AIRCRAFT, 'saveToolsAircraftBundle');
    var airframesSheet = getRequiredSheet_(ss, APP_SHEETS.AIRFRAMES, 'saveToolsAircraftBundle');
    var envelopesSheet = getRequiredSheet_(ss, APP_SHEETS.ENVELOPES, 'saveToolsAircraftBundle');
    var rollSheet = getRequiredSheet_(ss, 'Aircraft_Roll_Numbers', 'saveToolsAircraftBundle');

    var aircraftHeaders = _toolsSheetHeaderRow_(aircraftSheet);
    var requiredAircraftHeaders = ['MTOW_LAND', 'MTOW_WATER', 'MLW_LAND', 'MLW_WATER', 'POSICAO_AMBOS'];

    function ensureAircraftHeaders_(sheet, headers, requiredList) {
      var changed = false;
      var out = headers.slice();
      var have = {};
      out.forEach(function(h) { have[_toolsNormHeader_(h)] = true; });
      (requiredList || []).forEach(function(h) {
        var norm = _toolsNormHeader_(h);
        if (!have[norm]) {
          out.push(h);
          have[norm] = true;
          changed = true;
        }
      });
      if (changed) {
        sheet.getRange(1, 1, 1, out.length).setValues([out]);
      }
      return out;
    }

    aircraftHeaders = ensureAircraftHeaders_(aircraftSheet, aircraftHeaders, requiredAircraftHeaders);
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
    var aircraftType = typeKeyIdx >= 0 ? _toolsCanonicalAircraftType_(valueFor(aircraftHeaders[typeKeyIdx], aircraftPayload) || '').trim() : '';
    if (!registration) return { success: false, error: 'Registration is required in DB_Aircraft section.' };
    if (!aircraftType) return { success: false, error: 'Aircraft Type is required in DB_Aircraft section.' };

    var existing = aircraftSheet.getDataRange().getValues();
    var editingExisting = !!originalRegistration;
    if (regKeyIdx >= 0 && existing.length > 1) {
      var regNorm = _toolsNormalizeKeyValue_(registration);
      var originalNorm = _toolsNormalizeKeyValue_(originalRegistration || registration);
      var originalRowNumber = -1;
      for (var i = 1; i < existing.length; i++) {
        var existingReg = _toolsNormalizeKeyValue_(existing[i][regKeyIdx]);
        if (editingExisting && existingReg && existingReg === originalNorm) {
          originalRowNumber = i + 1;
          continue;
        }
        if (existingReg && existingReg === regNorm) {
          return { success: false, error: 'Registration already exists in DB_Aircraft: ' + registration };
        }
      }
      if (editingExisting && originalRowNumber < 0) {
        return { success: false, error: 'Original aircraft not found in DB_Aircraft: ' + originalRegistration };
      }
    }

    function buildRow(headerRow, rowObj) {
      return headerRow.map(function(header) {
        var key = _toolsNormHeader_(header);
        var val = valueFor(header, rowObj);
        if ((key === 'AIRCRAFT_TYPE' || key === 'TYPE_FOR_PERFORMANCE') && aircraftType) return val || aircraftType;
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

    var aircraftRowNumber = -1;
    if (editingExisting && regKeyIdx >= 0) {
      for (var j = 1; j < existing.length; j++) {
        if (_toolsNormalizeKeyValue_(existing[j][regKeyIdx]) === _toolsNormalizeKeyValue_(originalRegistration || registration)) {
          aircraftRowNumber = j + 1;
          aircraftSheet.getRange(aircraftRowNumber, 1, 1, aircraftRow.length).setValues([aircraftRow]);
          break;
        }
      }
    }
    if (aircraftRowNumber < 0) {
      aircraftSheet.appendRow(aircraftRow);
      aircraftRowNumber = aircraftSheet.getLastRow();
    }

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

    function deleteMatchingRows_(sheet, headerRow, candidates, matchValue) {
      var target = _toolsNormalizeKeyValue_(matchValue || '');
      if (!target) return 0;
      var matchIdx = _toolsFirstHeaderMatch_(headerRow, candidates);
      if (matchIdx < 0) return 0;
      var sheetData = sheet.getDataRange().getValues();
      var rowsToDelete = [];
      for (var rowIdx = sheetData.length - 1; rowIdx >= 1; rowIdx--) {
        if (_toolsNormalizeKeyValue_(sheetData[rowIdx][matchIdx]) === target) rowsToDelete.push(rowIdx + 1);
      }
      rowsToDelete.forEach(function(rowNumber) { sheet.deleteRow(rowNumber); });
      return rowsToDelete.length;
    }

    if (editingExisting) {
      deleteMatchingRows_(airframesSheet, airframeHeaders, ['REGISTRATION', 'REG', 'TAIL', 'TAIL_NUMBER'], originalRegistration);
      deleteMatchingRows_(airframesSheet, airframeHeaders, ['AIRCRAFT_TYPE', 'TYPE_FOR_PERFORMANCE'], originalAircraftType);
      deleteMatchingRows_(envelopesSheet, envelopeHeaders, ['REGISTRATION', 'REG', 'TAIL', 'TAIL_NUMBER'], originalRegistration);
      deleteMatchingRows_(envelopesSheet, envelopeHeaders, ['AIRCRAFT_TYPE', 'TYPE_FOR_PERFORMANCE'], originalAircraftType);
      deleteMatchingRows_(rollSheet, rollHeaders, ['REGISTRATION', 'REG', 'TAIL', 'TAIL_NUMBER'], originalRegistration);
      deleteMatchingRows_(rollSheet, rollHeaders, ['AIRCRAFT_TYPE', 'TYPE_FOR_PERFORMANCE'], originalAircraftType);
    }

    airframesSheet.getRange(airframesSheet.getLastRow() + 1, 1, airframeData.length, airframeData[0].length).setValues(airframeData);
    envelopesSheet.getRange(envelopesSheet.getLastRow() + 1, 1, envelopeData.length, envelopeData[0].length).setValues(envelopeData);
    rollSheet.getRange(rollSheet.getLastRow() + 1, 1, rollData.length, rollData[0].length).setValues(rollData);

    var out = {
      success: true,
      aircraftRowNumber: aircraftRowNumber,
      edited: editingExisting,
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

function addWaterRunwayToDatabase(data) {
  try {
    var payload = (data && typeof data === 'object') ? data : {};
    var icao = String(payload.icao || '').trim().toUpperCase();
    var pair = String(payload.runwayPair || '').trim().toUpperCase();
    var nameRaw = String(payload.waterName || '').trim();
    var communityRaw = String(payload.communityName || payload.community || '').trim();
    var elev = Number(payload.elevationFt);
    var midpointLat = Number(payload.midpointLat);
    var midpointLon = Number(payload.midpointLon);
    var runwayLengthM = Number(payload.runwayLengthM);
    var runwayWidthM = Number(payload.runwayWidthM);

    function waterAutoName_(rawName) {
      var cleaned = String(rawName || '').trim().replace(/\s+/g, '-').replace(/[^A-Za-z0-9\-]/g, '').toUpperCase();
      return 'W-' + (cleaned || 'UNNAMED');
    }
    function normalizeHeading_(deg) {
      var n = Number(deg);
      if (!isFinite(n)) return NaN;
      n = ((n % 360) + 360) % 360;
      return n === 0 ? 360 : n;
    }
    function headingToIdent_(deg) {
      var n = normalizeHeading_(deg);
      if (!isFinite(n)) return '';
      var ident = Math.round((n % 360) / 10);
      if (ident === 0) ident = 36;
      if (ident < 1) ident = 1;
      if (ident > 36) ident = 36;
      return String(ident).padStart(2, '0');
    }
    function parseHeadings_(raw) {
      var txt = String(raw == null ? '' : raw).trim();
      if (!txt) return [];
      var parts = txt.split(',').map(function(p) { return String(p || '').trim(); }).filter(Boolean);
      var out = [];
      for (var i = 0; i < parts.length; i++) {
        var h = normalizeHeading_(parts[i]);
        if (!isFinite(h)) return [];
        out.push(h);
      }
      if (out.length === 1) {
        out.push(normalizeHeading_(out[0] + 180));
      }
      return out.slice(0, 2);
    }

    if (!icao) return { success: false, error: 'ICAO is required' };
    if (!nameRaw) return { success: false, error: 'Water runway name is required' };
    if (!isFinite(elev)) return { success: false, error: 'Elevation (ft) is required' };
    if (!isFinite(midpointLat) || !isFinite(midpointLon)) return { success: false, error: 'Midpoint coordinates are required' };

    var headings = parseHeadings_(payload.runwayHeadings);
    if (!pair && headings.length >= 2) {
      var identA = headingToIdent_(headings[0]);
      var identB = headingToIdent_(headings[1]);
      if (identA && identB) pair = identA + '/' + identB;
    }
    if (!/^\d{2}\/\d{2}$/.test(pair)) return { success: false, error: 'runwayPair must be like 01/19 (or provide valid runwayHeadings)' };

    var parts = pair.split('/').map(function(p) { return String(p || '').trim().toUpperCase(); }).filter(Boolean);
    if (parts.length !== 2) return { success: false, error: 'Invalid runway pair' };

    var waterNameSource = nameRaw || communityRaw;
    var name = waterAutoName_(waterNameSource);
    var community = String(communityRaw || '').trim().toUpperCase();

    var headingByIdent = {};
    if (headings.length >= 2) {
      headingByIdent[parts[0]] = headings[0];
      headingByIdent[parts[1]] = headings[1];
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getRequiredSheet_(ss, APP_SHEETS.AIRPORTS, 'addWaterRunwayToDatabase');
    var values = sheet.getDataRange().getValues();
    if (!values || !values.length) return { success: false, error: 'DB_Airports header row missing' };

    var headers = values[0].map(function(h) { return String(h || '').trim().toUpperCase(); });
    function idxByAliases(aliases) {
      for (var i = 0; i < aliases.length; i++) {
        var idx = headers.indexOf(String(aliases[i] || '').trim().toUpperCase());
        if (idx >= 0) return idx;
      }
      return -1;
    }

    var idxIcao = idxByAliases(['ICAO']);
    var idxRwy = idxByAliases(['RWY_IDENT', 'RWY', 'RUNWAY', 'RUNWAY_DESIGNATOR']);
    var idxName = idxByAliases(['NOME', 'NAME', 'AIRPORT_NAME', 'RUNWAY_NAME']);
    var idxCommunity = idxByAliases(['MUNICIPIO___ALDEIA', 'MUNIC_PIO___ALDEIA', 'MUNICIPIO', 'MUN', 'COMMUNITY', 'COMUNIDADE']);
    var idxLat = idxByAliases(['LATITUDE', 'LAT']);
    var idxLon = idxByAliases(['LONGITUDE', 'LON', 'LNG']);
    var idxElev = idxByAliases(['ELEVATION', 'ELEVATION_FT', 'ALT_FEET']);
    var idxHeading = idxByAliases(['RUNWAY_HEADING', 'HEADING']);
    var idxLength = idxByAliases(['LENGTH_OFFICIAL', 'RUNWAY_LENGTH', 'LENGTH_M']);
    var idxWidth = idxByAliases(['WIDTH_OFFICIAL', 'RUNWAY_WIDTH', 'WIDTH_M']);
    var idxSurface = idxByAliases(['SURFACE_ACTUAL', 'RUNWAY_SURFACE_ACTUAL', 'SURFACE_OFFICIAL', 'RUNWAY_SURFACE', 'SURFACE_TYPE', 'SURFACE']);
    var idxInfrastructure = idxByAliases(['INFRASTRUCTURE', 'INFRASTRUCTURE_TYPE', 'INFRA_TYPE', 'INFRA', 'INFRAESTRUTURA']);
    var idxKnown = idxByAliases(['KNOWN_FEATURES', 'FEATURES']);

    if (idxIcao < 0 || idxRwy < 0) {
      return { success: false, error: 'DB_Airports must include ICAO and RWY_IDENT columns' };
    }
    if (idxInfrastructure < 0) {
      return { success: false, error: 'DB_Airports must include INFRASTRUCTURE (or INFRASTRUCTURE_TYPE) column for water runways' };
    }

    var existingByIdent = {};
    for (var r = 1; r < values.length; r++) {
      var rowIcao = String(values[r][idxIcao] || '').trim().toUpperCase();
      if (rowIcao !== icao) continue;
      var rowRwy = String(values[r][idxRwy] || '').trim().toUpperCase();
      if (!rowRwy) continue;
      existingByIdent[rowRwy] = r + 1;
    }

    var nowIso = new Date().toISOString();
    var created = [];
    for (var p = 0; p < parts.length; p++) {
      var ident = parts[p];
      if (existingByIdent[ident]) continue;
      var newRow = headers.map(function() { return ''; });
      newRow[idxIcao] = icao;
      newRow[idxRwy] = ident;
      if (idxName >= 0) newRow[idxName] = name;
      if (idxCommunity >= 0 && community) newRow[idxCommunity] = community;
      if (idxLat >= 0) newRow[idxLat] = midpointLat;
      if (idxLon >= 0) newRow[idxLon] = midpointLon;
      if (idxElev >= 0 && isFinite(elev)) newRow[idxElev] = elev;
      if (idxLength >= 0 && isFinite(runwayLengthM) && runwayLengthM > 0) newRow[idxLength] = runwayLengthM;
      if (idxWidth >= 0 && isFinite(runwayWidthM) && runwayWidthM > 0) newRow[idxWidth] = runwayWidthM;
      if (idxHeading >= 0) {
        var heading = headingByIdent[ident];
        newRow[idxHeading] = isFinite(heading) ? heading : (parseInt(ident, 10) * 10);
      }
      if (idxSurface >= 0) newRow[idxSurface] = 'WATER';
      if (idxInfrastructure >= 0) newRow[idxInfrastructure] = 'WATER';
      if (idxKnown >= 0) {
        newRow[idxKnown] = JSON.stringify({
          waterRunway: true,
          runwayPair: pair,
          runwayHeadings: headings,
          waterName: name,
          communityName: community,
          runwayLengthM: isFinite(runwayLengthM) ? runwayLengthM : '',
          runwayWidthM: isFinite(runwayWidthM) ? runwayWidthM : '',
          midpoint: { lat: midpointLat, lon: midpointLon },
          createdAt: nowIso,
          source: 'water_runway_tool'
        });
      }
      sheet.appendRow(newRow);
      created.push(ident);
    }

    return {
      success: true,
      icao: icao,
      runwayPair: pair,
      waterName: name,
      createdIdents: created,
      message: created.length ? ('Created ' + created.join(', ')) : 'Runway entries already existed'
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getLessonEvalHistory(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var pilotFilter = String(body.pilot || '').trim().toLowerCase();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(APP_SHEETS.FLIGHT_LESSON_EVALS || 'DB_Flight_Lesson_Evals');
    if (!sh) return { success: true, records: [] };
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: true, records: [] };
    var headers = data[0].map(function(h) { return String(h || '').trim().toUpperCase(); });
    var map = _lessonEvalHeaderMap_(headers);
    var out = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var pilotVal = map.PILOT >= 0 ? String(row[map.PILOT] || '').trim() : '';
      if (pilotFilter && pilotVal.toLowerCase() !== pilotFilter && pilotVal.toLowerCase().indexOf(pilotFilter) < 0) continue;
      out.push({
        rowNumber:    i + 1,
        evalId:       map.EVAL_ID       >= 0 ? String(row[map.EVAL_ID]       || '') : '',
        missionId:    map.MISSION_ID    >= 0 ? String(row[map.MISSION_ID]    || '') : '',
        trainingCode: map.TRAINING_CODE >= 0 ? String(row[map.TRAINING_CODE] || '') : '',
        pilot:        pilotVal,
        aircraft:     map.AIRCRAFT      >= 0 ? String(row[map.AIRCRAFT]      || '') : '',
        missionDate:  map.MISSION_DATE  >= 0 ? String(row[map.MISSION_DATE]  || '') : '',
        status:       map.STATUS        >= 0 ? String(row[map.STATUS]        || '') : '',
        submittedAt:  map.SUBMITTED_AT  >= 0 ? String(row[map.SUBMITTED_AT]  || '') : '',
        scoresJson:   map.SCORES_JSON   >= 0 ? String(row[map.SCORES_JSON]   || '{}') : '{}',
        notesJson:    map.NOTES_JSON    >= 0 ? String(row[map.NOTES_JSON]    || '{}') : '{}',
        updatedAt:    map.UPDATED_AT    >= 0 ? String(row[map.UPDATED_AT]    || '') : ''
      });
    }
    out.sort(function(a, b) {
      var da = a.missionDate || a.updatedAt || '';
      var db = b.missionDate || b.updatedAt || '';
      return da > db ? -1 : da < db ? 1 : 0;
    });
    return { success: true, records: out };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e), records: [] };
  }
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

  // Persist takeoff performance data from runwaySnapshot + wbSnapshot into ACTUAL_LOAD_JSON
  const runwaySnap = (payload && payload.runwaySnapshot && typeof payload.runwaySnapshot === 'object') ? payload.runwaySnapshot : null;
  const wbSnap = (payload && payload.wbSnapshot && typeof payload.wbSnapshot === 'object') ? payload.wbSnapshot : null;
  if (runwaySnap || wbSnap) {
    const logHeaders = logData[0].map(function(h) { return String(h || '').toUpperCase().trim().replace(/\s+/g, '_'); });
    const actualLoadColIdx = logHeaders.indexOf('ACTUAL_LOAD_JSON') >= 0 ? logHeaders.indexOf('ACTUAL_LOAD_JSON') : LOG_FLIGHT_COL.ACTUAL_LOAD_JSON;
    const existingRaw = String(logData[logRow - 1][actualLoadColIdx] || '');
    let existingLoad = {};
    try { existingLoad = existingRaw ? JSON.parse(existingRaw) : {}; } catch (e2) { existingLoad = {}; }
    const surfCond = String((runwaySnap && runwaySnap.surfaceCondition) || '').trim().toUpperCase();
    const numericOrExisting = function(value, fallback) {
      return (value != null && value !== '' && isFinite(Number(value))) ? Number(value) : fallback;
    };
    const mergedLoad = Object.assign({}, existingLoad, {
      runwayIdent: String((runwaySnap && runwaySnap.runwayIdent) || existingLoad.runwayIdent || '').trim(),
      fromIcao: String((runwaySnap && runwaySnap.fromIcao) || existingLoad.fromIcao || '').trim().toUpperCase(),
      surface: String((runwaySnap && runwaySnap.surface) || existingLoad.surface || '').trim(),
      surfaceCondition: surfCond || existingLoad.surfaceCondition || '',
      wet: surfCond ? (surfCond === 'WET') : existingLoad.wet,
      runwayLengthM: numericOrExisting(runwaySnap && runwaySnap.runwayLengthM, existingLoad.runwayLengthM),
      runwayWidthM: numericOrExisting(runwaySnap && runwaySnap.runwayWidthM, existingLoad.runwayWidthM),
      officialRunwayLengthM: numericOrExisting(runwaySnap && runwaySnap.officialRunwayLengthM, existingLoad.officialRunwayLengthM),
      officialRunwayWidthM: numericOrExisting(runwaySnap && runwaySnap.officialRunwayWidthM, existingLoad.officialRunwayWidthM),
      elevationFt: numericOrExisting(runwaySnap && runwaySnap.elevationFt, existingLoad.elevationFt),
      slopePct: numericOrExisting(runwaySnap && runwaySnap.slopePct, existingLoad.slopePct),
      cutdownAreaM: numericOrExisting(runwaySnap && runwaySnap.cutdownAreaM, existingLoad.cutdownAreaM),
      heading: numericOrExisting(runwaySnap && runwaySnap.heading, existingLoad.heading),
      windDir: numericOrExisting(runwaySnap && runwaySnap.windDir, existingLoad.windDir),
      windSpd: numericOrExisting(runwaySnap && runwaySnap.windSpd, existingLoad.windSpd),
      tempC: numericOrExisting(runwaySnap && runwaySnap.tempC, existingLoad.tempC),
      qnhHpa: numericOrExisting((runwaySnap && (runwaySnap.qnhHpa != null ? runwaySnap.qnhHpa : runwaySnap.qnh)) != null
        ? (runwaySnap.qnhHpa != null ? runwaySnap.qnhHpa : runwaySnap.qnh)
        : (payload && payload.qnhHpa), existingLoad.qnhHpa),
      flaps: numericOrExisting(runwaySnap && runwaySnap.flaps, existingLoad.flaps),
      humidity: numericOrExisting(runwaySnap && runwaySnap.humidity, existingLoad.humidity),
      headTailKts: numericOrExisting(runwaySnap && runwaySnap.headTailKts, existingLoad.headTailKts),
      crosswindKts: numericOrExisting(runwaySnap && runwaySnap.crosswindKts, existingLoad.crosswindKts),
      calcToRoll: (runwaySnap && runwaySnap.takeoffRollM != null) ? Number(runwaySnap.takeoffRollM) : (existingLoad.calcToRoll != null ? existingLoad.calcToRoll : null),
      landingRollM: numericOrExisting(runwaySnap && runwaySnap.landingRollM, existingLoad.landingRollM),
      combinedRollM: numericOrExisting(runwaySnap && runwaySnap.rollStopM, existingLoad.combinedRollM),
      abortPointM: numericOrExisting(runwaySnap && runwaySnap.abortPointM, existingLoad.abortPointM),
      vrDistM: (runwaySnap && runwaySnap.vrDistM != null) ? Number(runwaySnap.vrDistM) : (existingLoad.vrDistM != null ? existingLoad.vrDistM : null),
      mapDistM: (runwaySnap && runwaySnap.mapDistM != null) ? Number(runwaySnap.mapDistM) : (existingLoad.mapDistM != null ? existingLoad.mapDistM : null),
      densityAltitudeFt: numericOrExisting(runwaySnap && runwaySnap.densityAltitudeFt, existingLoad.densityAltitudeFt),
      pressureAltitudeFt: numericOrExisting(runwaySnap && runwaySnap.pressureAltitudeFt, existingLoad.pressureAltitudeFt),
      isaDeviationC: numericOrExisting(runwaySnap && runwaySnap.isaDeviationC, existingLoad.isaDeviationC),
      blocking: typeof (runwaySnap && runwaySnap.blocking) === 'boolean' ? runwaySnap.blocking : existingLoad.blocking,
      warnings: Array.isArray(runwaySnap && runwaySnap.warnings) ? runwaySnap.warnings.slice() : (Array.isArray(existingLoad.warnings) ? existingLoad.warnings : []),
      releasedAt: now.toISOString()
    });
    if (wbSnap && Number(wbSnap.weightKg) > 0) mergedLoad.weightKg = Number(wbSnap.weightKg);
    if (wbSnap && Number(wbSnap.cgIn) > 0) mergedLoad.cgIn = Number(wbSnap.cgIn);
    logSheet.getRange(logRow, actualLoadColIdx + 1).setValue(JSON.stringify(mergedLoad));
  }

  // DB_Dispatch: Column L = 12 status => In-Flight (all rows in mission)
  const dispatchData = dispatchSheet.getDataRange().getValues();
  let dispatchRowsUpdated = 0;
  for (let j = 1; j < dispatchData.length; j++) {
    if (String(dispatchData[j][DISPATCH_COL.MISSION_ID] || '').trim() === missionId) {
      dispatchSheet.getRange(j + 1, DISPATCH_COL.STATUS + 1).setValue('In-Flight');
      dispatchRowsUpdated++;
    }
  }
  invalidateScheduledMissionsCache_();

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

function recordAirborne(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  if (!logSheet) throw new Error("Sheet 'LOG_Flights' not found.");

  const flightLegId = String(payload && payload.flightLegId || '').trim();
  if (!flightLegId) throw new Error('recordAirborne: flightLegId is required');

  const data = logSheet.getDataRange().getValues();
  if (!data || data.length < 2) throw new Error('recordAirborne: LOG_Flights has no data rows');

  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][LOG_FLIGHT_COL.FLIGHT_ID] || '').trim() === flightLegId) {
      rowIdx = i + 1;
      break;
    }
  }
  if (rowIdx < 0) throw new Error('recordAirborne: flight not found in LOG_Flights: ' + flightLegId);

  const ts = payload && payload.airborne ? new Date(payload.airborne) : new Date();
  logSheet.getRange(rowIdx, LOG_FLIGHT_COL.AIRBORNE + 1).setValue(ts);

  return { success: true, flightLegId: flightLegId, airborne: ts.toISOString() };
}

function recordLanded(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  if (!logSheet) throw new Error("Sheet 'LOG_Flights' not found.");

  const flightLegId = String(payload && payload.flightLegId || '').trim();
  if (!flightLegId) throw new Error('recordLanded: flightLegId is required');

  const data = logSheet.getDataRange().getValues();
  if (!data || data.length < 2) throw new Error('recordLanded: LOG_Flights has no data rows');

  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][LOG_FLIGHT_COL.FLIGHT_ID] || '').trim() === flightLegId) {
      rowIdx = i + 1;
      break;
    }
  }
  if (rowIdx < 0) throw new Error('recordLanded: flight not found in LOG_Flights: ' + flightLegId);

  const ts = payload && payload.landed ? new Date(payload.landed) : new Date();
  logSheet.getRange(rowIdx, LOG_FLIGHT_COL.LANDED + 1).setValue(ts);
  const syncRes = _syncDispatchMissionStatusForFlightLeg_(flightLegId);

  return { success: true, flightLegId: flightLegId, landed: ts.toISOString(), missionStatus: syncRes };
}

function clearAutoLoggedLanded(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  if (!logSheet) throw new Error("Sheet 'LOG_Flights' not found.");

  const flightLegId = String(payload && payload.flightLegId || '').trim();
  if (!flightLegId) throw new Error('clearAutoLoggedLanded: flightLegId is required');

  const data = logSheet.getDataRange().getValues();
  if (!data || data.length < 2) throw new Error('clearAutoLoggedLanded: LOG_Flights has no data rows');

  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][LOG_FLIGHT_COL.FLIGHT_ID] || '').trim() === flightLegId) {
      rowIdx = i + 1;
      break;
    }
  }
  if (rowIdx < 0) throw new Error('clearAutoLoggedLanded: flight not found in LOG_Flights: ' + flightLegId);

  logSheet.getRange(rowIdx, LOG_FLIGHT_COL.LANDED + 1).setValue('');
  const syncRes = _syncDispatchMissionStatusForFlightLeg_(flightLegId);
  return { success: true, flightLegId: flightLegId, landed: '', missionStatus: syncRes };
}

function recordTouchAndGo(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  if (!logSheet) throw new Error("Sheet 'LOG_Flights' not found.");

  const flightLegId = String(payload && payload.flightLegId || '').trim();
  if (!flightLegId) throw new Error('recordTouchAndGo: flightLegId is required');

  const touchAndGoCount = Math.max(1, parseInt(payload && payload.touchAndGoCount, 10) || 1);

  const data = logSheet.getDataRange().getValues();
  if (!data || data.length < 2) throw new Error('recordTouchAndGo: LOG_Flights has no data rows');

  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][LOG_FLIGHT_COL.FLIGHT_ID] || '').trim() === flightLegId) {
      rowIdx = i + 1;
      break;
    }
  }
  if (rowIdx < 0) throw new Error('recordTouchAndGo: flight not found in LOG_Flights: ' + flightLegId);

  // NUM_LDGS = T&Gs so far + 1 (the expected final landing still to come)
  logSheet.getRange(rowIdx, LOG_FLIGHT_COL.NUM_LDGS + 1).setValue(touchAndGoCount + 1);

  const tgHeaders = data[0].map(function(v) { return String(v || '').toUpperCase().trim().replace(/\s+/g, '_'); });
  const tgColFn = function(name, fallback) { const i = tgHeaders.indexOf(name); return i >= 0 ? i : fallback; };
  const tgColIdx = tgColFn('NUMBER_TOUCH_AND_GOS', LOG_FLIGHT_COL.NUM_TOUCH_AND_GOS);
  logSheet.getRange(rowIdx, tgColIdx + 1).setValue(touchAndGoCount);

  return {
    success: true,
    flightLegId: flightLegId,
    touchAndGoCount: touchAndGoCount,
    numLdgs: touchAndGoCount + 1
  };
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

  const rollObj = (rollPayload && typeof rollPayload === 'object') ? rollPayload : {};
  const rollDistanceM = (rollObj.distanceAtVrM != null && isFinite(Number(rollObj.distanceAtVrM)))
    ? Math.round(Number(rollObj.distanceAtVrM))
    : ((rollObj.distanceAtVrNm != null && isFinite(Number(rollObj.distanceAtVrNm)))
      ? Math.round(Number(rollObj.distanceAtVrNm) * 1852)
      : null);

  const merged = {
    ...existingJson,
    // Canonical captured metric: meters from Start Roll button press to Vr hit.
    actualToRollM: rollDistanceM,
    vrDistM: rollDistanceM,
    takeoffRoll: rollObj,
    takeoffRollSavedAt: new Date().toISOString()
  };

  logSheet.getRange(targetRow, LOG_FLIGHT_COL.ACTUAL_LOAD_JSON + 1).setValue(JSON.stringify(merged));
  if (rollDistanceM != null) {
    const logHeaders = data[0].map(function(h) { return String(h || '').toUpperCase().trim().replace(/\s+/g, '_'); });
    const actualToRollColIdx = (logHeaders.indexOf('ACTUAL_TO_ROLL') >= 0)
      ? logHeaders.indexOf('ACTUAL_TO_ROLL')
      : (LOG_FLIGHT_COL.ACTUAL_TO_ROLL != null ? LOG_FLIGHT_COL.ACTUAL_TO_ROLL : -1);
    if (actualToRollColIdx >= 0) {
      logSheet.getRange(targetRow, actualToRollColIdx + 1).setValue(rollDistanceM);
    }
  }
  return true;
}

function getPerformanceSetup(icao, aircraftType) {
  const airportCode = String(icao || '').trim().toUpperCase();
  const aircraftTypeKey = _perfAircraftTypeKey_(aircraftType);

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
      let briefingCard = {};
      const featuresStr = String(_perfValue(r, ['KNOWN_FEATURES', 'FEATURES'], '')).trim();
      if (featuresStr) {
        try {
          const parsed = JSON.parse(featuresStr);
          knownObj = Array.isArray(parsed) ? { features: parsed } : (parsed || {});
          verifiedOperational = (knownObj && knownObj.verifiedOperational && typeof knownObj.verifiedOperational === 'object') ? knownObj.verifiedOperational : {};
          officialReference = (knownObj && knownObj.officialReference && typeof knownObj.officialReference === 'object') ? knownObj.officialReference : {};
          briefingCard = (knownObj && knownObj.briefingCard && typeof knownObj.briefingCard === 'object') ? knownObj.briefingCard : {};
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
      const briefingCardCutdownM = _perfNum(
        briefingCard.cutdownAreaM != null ? briefingCard.cutdownAreaM : (briefingCard.clearway != null ? briefingCard.clearway : 0),
        0
      ) || 0;
      
      return {
        icao: String(_perfValue(r, ['ICAO'], airportCode)).trim().toUpperCase(),
        rwyIdent: rwyIdent,
        headingDeg: explicitHeading > 0 ? explicitHeading : _runwayHeadingFromIdent(rwyIdent),
        length: lengthM,
        width: widthM,
        cutdownAreaM: _perfNum(
          verifiedOperational.cutdownAreaM != null
            ? verifiedOperational.cutdownAreaM
            : (briefingCardCutdownM > 0
                ? briefingCardCutdownM
                : (knownObj.cutdownAreaM != null ? knownObj.cutdownAreaM : (knownObj.clearway != null ? knownObj.clearway : 0))),
          0
        ) || 0,
        internalUpdatedAt: String((knownObj.currentSurveyVersion && knownObj.currentSurveyVersion.publishedAt) || (knownObj.verifiedSurvey && knownObj.verifiedSurvey.capturedAt) || knownObj.updatedAt || '').trim(),
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
        briefingCard: briefingCard,
        verifiedOperational: verifiedOperational,
        surveyPilot: String((knownObj.verifiedSurvey && knownObj.verifiedSurvey.pilotName) || '').trim()
      };
    })
    .sort((a, b) => {
      const ha = _perfNum(a.headingDeg, 0);
      const hb = _perfNum(b.headingDeg, 0);
      if (ha !== hb) return ha - hb;
      return String(a.rwyIdent || '').localeCompare(String(b.rwyIdent || ''));
    });

  const flapRows = aircraftTypeKey
    ? perfTable.rows.filter(r => _perfAircraftTypeKey_(_perfValue(r, ['AIRCRAFT_TYPE', 'TYPE_FOR_PERFORMANCE'], '')) === aircraftTypeKey)
    : perfTable.rows.slice();
  const flapSourceRows = flapRows.length ? flapRows : perfTable.rows;

  const flapSet = {};
  flapSourceRows.forEach(r => {
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

    const perfRowsForAircraft = perfTable.rows.filter(r => _perfAircraftTypeKey_(_perfValue(r, ['AIRCRAFT_TYPE', 'TYPE_FOR_PERFORMANCE'], '')) === _perfAircraftTypeKey_(aircraftType));
    const perfRows = perfRowsForAircraft.length ? perfRowsForAircraft : perfTable.rows;

    const daTo = _nearestMultiplier(perfRows, ['DA_FT', 'DA'], densityAltitudeFt, ['TO_MULTIPLIER'], 1);
    const daLdg = _nearestMultiplier(perfRows, ['DA_FT', 'DA'], densityAltitudeFt, ['LDG_MULTIPLIER'], 1);

    const isWet = surfaceCondition === 'WET';
    const surfaceTo = _surfaceMultiplier(perfRows, surface, [isWet ? 'TAKEOFF_WET' : 'TAKEOFF_DRY'], 1);
    const surfaceLdg = _surfaceMultiplier(perfRows, surface, [isWet ? 'LANDING_WET' : 'LANDING_DRY'], 1);

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
      const windTo  = effectiveWindKts === 0 ? 1 : _nearestMultiplier(perfRows, ['WIND_KTS', 'WIND'], effectiveWindKts, [windColTo], 1);
      const windLdg = effectiveWindKts === 0 ? 1 : _nearestMultiplier(perfRows, ['WIND_KTS', 'WIND'], effectiveWindKts, [windColLdg], 1);

    const humidity = _nearestMultiplier(perfRows, ['HUMIDITY'], humidityPct, ['HUMIDITY_FACTOR'], 1);
    const flap = _nearestMultiplier(perfRows, ['FLAP_SETTING'], flapSetting, ['FLAP_FACTOR'], 1);

    const toNoSlope = base.to * daTo * surfaceTo * flap * humidity * windTo;
    let slopeTo = 1;
    let effectiveSlopeTakeoff = slopePercent;
    for (let i = 0; i < 3; i++) {
      const estTakeoff = toNoSlope * slopeTo;
      effectiveSlopeTakeoff = _effectiveSlopeOverDistance(slopeProfile, Math.min(estTakeoff, runwayLengthM));
      const slopeAbsTo = Math.abs(effectiveSlopeTakeoff);
      const slopeTargetTo = _slopeLookupTarget(perfRows, slopeAbsTo);
      slopeTo = _nearestMultiplier(
        perfRows,
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
      const slopeTargetLdg = _slopeLookupTarget(perfRows, slopeAbsLdg);
      slopeLdg = _nearestMultiplier(
        perfRows,
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
    var durationHoursFromAny = function(value) {
      if (value == null || value === '') return null;
      if (value instanceof Date && !isNaN(value.getTime())) {
        if (value.getFullYear() <= 1900) {
          return value.getHours() + (value.getMinutes() / 60) + (value.getSeconds() / 3600) + (value.getMilliseconds() / 3600000);
        }
        return null;
      }
      if (typeof value === 'number' && isFinite(value)) {
        if (value < 0) return null;
        return value;
      }
      var raw = String(value || '').trim();
      if (!raw) return null;
      var hm = raw.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
      if (hm) {
        var hh = Number(hm[1] || 0);
        var mm = Number(hm[2] || 0);
        var ss = Number(hm[3] || 0);
        if (!isNaN(hh) && !isNaN(mm) && !isNaN(ss)) return hh + (mm / 60) + (ss / 3600);
      }
      var n = parseFloat(raw);
      if (!isNaN(n) && isFinite(n) && n >= 0) {
        return n;
      }
      return null;
    };
    var durationHoursBetween = function(startValue, endValue) {
      var startDt = rowDate(startValue);
      var endDt = rowDate(endValue);
      if (!startDt || !endDt) return null;
      var deltaHours = (endDt.getTime() - startDt.getTime()) / 3600000;
      if (!isFinite(deltaHours) || deltaHours <= 0 || deltaHours > 24) return null;
      return deltaHours;
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
        totalTime: dispatchTimeIdx >= 0 ? (durationHoursFromAny(row[dispatchTimeIdx]) || num(row[dispatchTimeIdx])) : 0,
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
    var logBrakesReleaseIdx = indexByAliases(logData.headers, ['BRAKES_RELEASE'], LOG_FLIGHT_COL.BRAKES_RELEASE);
    var logBrakesAppliedIdx = indexByAliases(logData.headers, ['BRAKES_APPLIED', 'ON_BLOCKS'], LOG_FLIGHT_COL.BRAKES_APPLIED);
    var logOnBlocksIdx = indexByAliases(logData.headers, ['ON_BLOCKS', 'BRAKES_APPLIED'], LOG_FLIGHT_COL.ON_BLOCKS);

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
      var totalTime = logTimeIdx >= 0 ? durationHoursFromAny(row[logTimeIdx]) : null;
      if (!(totalTime > 0)) {
        var brakesStart = logBrakesReleaseIdx >= 0 ? row[logBrakesReleaseIdx] : '';
        var brakesEnd = (logBrakesAppliedIdx >= 0 ? row[logBrakesAppliedIdx] : '') || (logOnBlocksIdx >= 0 ? row[logOnBlocksIdx] : '');
        totalTime = durationHoursBetween(brakesStart, brakesEnd);
      }
      if (!(totalTime > 0)) {
        var airStart = logAirborneIdx >= 0 ? row[logAirborneIdx] : '';
        var airEnd = logLandedIdx >= 0 ? row[logLandedIdx] : '';
        totalTime = durationHoursBetween(airStart, airEnd);
      }
      if (!(totalTime > 0)) totalTime = durationHoursFromAny(dispatchInfo.totalTime);
      if (!(totalTime > 0)) totalTime = num(dispatchInfo.totalTime);
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

    var authRes = getPilotAircraftAuthorizations({});

    return {
      success: true,
      currentUser: me,
      staff: staffRows,
      roles: roles,
      modules: modules,
      pilotAuthorizations: authRes && authRes.success && Array.isArray(authRes.rows) ? authRes.rows : [],
      authorizationConfig: _pilotAuthorizationOptionSets_()
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getToolsStudentProfiles() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.PILOTS, 'getToolsStudentProfiles');
    _toolsEnsurePilotStudentHeaders_(sh);
    var data = sh.getDataRange().getValues();
    if (data.length < 1) return { success: true, rows: [], ratingsOptions: _toolsStudentRatingsOptions_() };
    var headers = data[0] || [];
    var ledgerByKey = _toolsStudentLedgerMap_(ss);
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var rec = _toolsStudentRecordFromRow_(headers, data[i], i + 1);
      if (!rec.name && !rec.email && !rec.cpf && !rec.canac) continue;
      if (!rec.studentKey) rec.studentKey = _toolsStudentIdentityKey_(rec);
      rec.ledger = ledgerByKey[rec.studentKey] || [];
      if (rec.ledger.length) {
        var lastEntry = rec.ledger[rec.ledger.length - 1] || {};
        var bal = Number(lastEntry.balanceAfter || rec.balanceBrl || 0);
        rec.balanceBrl = isFinite(bal) ? bal : Number(rec.balanceBrl || 0);
      }
      rows.push(rec);
    }
    rows.sort(function(a, b) { return String(a.name || '').localeCompare(String(b.name || '')); });
    return { success: true, rows: rows, ratingsOptions: _toolsStudentRatingsOptions_() };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e), rows: [], ratingsOptions: _toolsStudentRatingsOptions_() };
  }
}

function saveToolsStudentProfile(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.PILOTS, 'saveToolsStudentProfile');
    var headers = _toolsEnsurePilotStudentHeaders_(sh);
    var data = sh.getDataRange().getValues();

    var rowNumber = Number(body.rowNumber || 0);
    var name = String(body.name || '').trim();
    var email = String(body.email || '').trim().toLowerCase();
    var cpf = String(body.cpf || '').trim();
    var canac = String(body.canac || '').trim();
    if (!name) return { success: false, error: 'Student name is required.' };
    if (!email && !cpf && !canac) return { success: false, error: 'Provide email, CPF or CANAC.' };

    var emailIdx = _toolsHeaderIndexFromCandidates_(headers, ['PILOT_EMAIL', 'EMAIL', 'E_MAIL']);
    var cpfIdx = _toolsHeaderIndexFromCandidates_(headers, ['CPF', 'ID_NUMBER_CPF_PASSPORT']);
    var canacIdx = _toolsHeaderIndexFromCandidates_(headers, ['CANAC', 'ANAC_CODE', 'LICENSE']);

    if (rowNumber < 2) {
      for (var i = 1; i < data.length; i++) {
        var rowEmail = emailIdx >= 0 ? String(data[i][emailIdx] || '').trim().toLowerCase() : '';
        var rowCpf = cpfIdx >= 0 ? String(data[i][cpfIdx] || '').trim() : '';
        var rowCanac = canacIdx >= 0 ? String(data[i][canacIdx] || '').trim() : '';
        if ((email && rowEmail && rowEmail === email) || (cpf && rowCpf && rowCpf === cpf) || (canac && rowCanac && rowCanac === canac)) {
          rowNumber = i + 1;
          break;
        }
      }
    }

    var existingRow = rowNumber >= 2 ? sh.getRange(rowNumber, 1, 1, headers.length).getValues()[0] : headers.map(function() { return ''; });
    var existingRecord = _toolsStudentRecordFromRow_(headers, existingRow, rowNumber >= 2 ? rowNumber : 0);
    var personType = String(body.personType || existingRecord.personType || 'OPS_PILOT').trim().toUpperCase();
    if (personType !== 'STUDENT' && personType !== 'OPS_PILOT' && personType !== 'STAFF') personType = 'OPS_PILOT';
    var studentKey = existingRecord.studentKey || _toolsStudentIdentityKey_({ name: name, email: email, cpf: cpf, canac: canac, dob: String(body.dob || '').trim() });
    var currentBalance = _toolsStudentCurrentBalanceFromLedger_(ss, studentKey, Number(existingRecord.balanceBrl || 0));
    var rate = Number(body.rateBrlHour || existingRecord.rateBrlHour || 0);
    var creditAmount = Number(body.creditAmountBrl || 0);
    var debitHours = Number(body.debitHours || 0);
    if (isNaN(rate) || rate < 0) rate = 0;
    if (isNaN(creditAmount) || creditAmount < 0) creditAmount = 0;
    if (isNaN(debitHours) || debitHours < 0) debitHours = 0;
    if (personType === 'STAFF') {
      rate = 0;
      creditAmount = 0;
    }
    if (debitHours > 0) {
      return { success: false, error: 'Manual debit is disabled. Training mission approval posts debit automatically.' };
    }
    if (creditAmount > 0 && rate <= 0) {
      return { success: false, error: 'Rate (R$/hour) is required for each deposit credit.' };
    }

    var ratingsWithExpiry = Array.isArray(body.ratingsWithExpiry) ? body.ratingsWithExpiry : [];
    var ratings = [];
    var ratingsExpiryMap = {};
    if (ratingsWithExpiry.length > 0) {
      ratingsWithExpiry.forEach(function(it) {
        var code = String(it && it.code || '').trim();
        var expiry = _toolsStudentNormalizeDateIso_(String(it && it.expiry || '').trim());
        if (!code) return;
        ratings.push(code);
        if (expiry) ratingsExpiryMap[code] = expiry;
      });
    } else {
      ratings = Array.isArray(body.ratings) ? body.ratings : String(body.ratings || '').split(/[;,]+/);
      ratings = ratings.map(function(v) { return String(v || '').trim(); }).filter(Boolean);
      ratingsExpiryMap = (existingRecord && existingRecord.ratingsExpiry && typeof existingRecord.ratingsExpiry === 'object')
        ? existingRecord.ratingsExpiry
        : {};
    }

    var now = new Date();
    var nowIso = Utilities.formatDate(now, Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd HH:mm:ss');
    var depositFileUrl = String(body.depositFileUrl || '').trim();
    var depositFileId = String(body.depositFileId || '').trim();
    var dobIso = _toolsStudentNormalizeDateIso_(String(body.dob || '').trim());
    var medExpiryIso = _toolsStudentNormalizeDateIso_(String(body.medicalExpiry || '').trim());

    var ledgerEntries = _toolsStudentLedgerMap_(ss)[studentKey] || [];
    var fifoEntries = ledgerEntries.slice();
    if (creditAmount > 0) {
      fifoEntries.push({
        timestamp: nowIso,
        amount: creditAmount,
        rate: rate,
        note: String(body.creditNote || '').trim(),
        receiptUrl: depositFileUrl,
        by: String(body.updatedBy || _toolsCurrentUserEmail_() || 'tools'),
        balanceAfter: _toolsRoundMoney_(currentBalance + creditAmount)
      });
    }

    var fifoDebit = _toolsStudentComputeFifoDebit_(fifoEntries, 0);
    if (!fifoDebit.success) {
      return { success: false, error: fifoDebit.error };
    }

    var newBalance = _toolsRoundMoney_(currentBalance + creditAmount - Number(fifoDebit.totalAmount || 0));

    var profileJson = {
      name: name,
      email: email,
      cpf: cpf,
      canac: canac,
      dob: dobIso,
      medicalExpiry: medExpiryIso,
      medicalClass: String(body.medicalClass || '').trim(),
      personType: personType,
      ratings: ratings,
      ratingsExpiry: ratingsExpiryMap,
      studentKey: studentKey,
      balanceBrl: newBalance,
      rateBrlHour: rate,
      lastUpdatedAt: nowIso
    };

    var lastDepositAtIdx = _toolsHeaderIndexFromCandidates_(headers, ['LAST_DEPOSIT_AT']);
    var existingLastDepositAt = lastDepositAtIdx >= 0 ? String(existingRow[lastDepositAtIdx] || '').trim() : '';

    var dataMap = {
      PILOT_NAME: name,
      PILOT: name,
      NAME: name,
      PILOT_EMAIL: email,
      EMAIL: email,
      CPF: cpf,
      CANAC: canac,
      DOB: dobIso,
      MEDICAL_EXPIRY: medExpiryIso,
      MEDICAL_CLASS: String(body.medicalClass || '').trim(),
      PERSON_TYPE: personType,
      CURRENT_RATINGS: ratings.join('; '),
      RATINGS_EXPIRY_JSON: JSON.stringify(ratingsExpiryMap),
      MNTE_VALIDITY: String(ratingsExpiryMap.MNTE || ''),
      MNAF_VALIDITY: String(ratingsExpiryMap.MNAF || ''),
      STUDENT_KEY: studentKey,
      STUDENT_PROFILE_JSON: JSON.stringify(profileJson),
      FUNDS_BALANCE_BRL: newBalance,
      FUNDS_RATE_BRL_HOUR: rate,
      LAST_DEPOSIT_FILE_URL: depositFileUrl,
      LAST_DEPOSIT_FILE_ID: depositFileId,
      LAST_DEPOSIT_AT: depositFileUrl ? nowIso : existingLastDepositAt
    };

    var record = headers.map(function(header) {
      var key = _toolsNormHeader_(header);
      return Object.prototype.hasOwnProperty.call(dataMap, key) ? dataMap[key] : '';
    });

    var result;
    if (rowNumber >= 2) {
      var merged = headers.map(function(header, idx) {
        var key = _toolsNormHeader_(header);
        return Object.prototype.hasOwnProperty.call(dataMap, key) ? dataMap[key] : existingRow[idx];
      });
      sh.getRange(rowNumber, 1, 1, merged.length).setValues([merged]);
      result = { success: true, action: 'updated', rowNumber: rowNumber };
    } else {
      sh.appendRow(record);
      result = { success: true, action: 'created', rowNumber: sh.getLastRow() };
      rowNumber = result.rowNumber;
    }

    if (creditAmount > 0) {
      _toolsAppendStudentLedgerEntry_(ss, {
        studentKey: studentKey,
        studentRowNumber: rowNumber,
        studentName: name,
        studentEmail: email,
        txDate: nowIso,
        txType: 'CREDIT',
        amountBrl: creditAmount,
        rateBrlHour: rate,
        balanceAfterBrl: newBalance,
        receiptFileUrl: depositFileUrl,
        receiptFileId: depositFileId,
        notes: String(body.creditNote || '').trim(),
        referenceType: 'MANUAL_CREDIT',
        referenceId: '',
        createdBy: String(body.updatedBy || _toolsCurrentUserEmail_() || 'tools')
      });
    }

    result.balanceBrl = newBalance;
    result.rateBrlHour = rate;
    result.studentKey = studentKey;
    result.debitedHours = 0;
    result.debitedAmountBrl = 0;
    return result;
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function uploadStudentDepositReceipt(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var base64Data = String(body.base64Data || '').trim();
    var fileName = String(body.fileName || '').trim() || 'deposit.pdf';
    var mimeType = String(body.mimeType || 'application/pdf').trim().toLowerCase();
    var studentName = String(body.studentName || '').trim() || 'student';

    if (!base64Data) return { success: false, error: 'No file data provided.' };
    if (mimeType.indexOf('pdf') < 0 && !/\.pdf$/i.test(fileName)) return { success: false, error: 'Only PDF receipts are allowed.' };

    var rootName = 'FlightApp_Student_Deposits';
    var rootFolders = DriveApp.getFoldersByName(rootName);
    var rootFolder = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder(rootName);

    var studentFolderName = String(studentName).replace(/[^A-Z0-9 _-]/gi, '').trim();
    if (!studentFolderName) studentFolderName = 'Student';
    var studentFolders = rootFolder.getFoldersByName(studentFolderName);
    var studentFolder = studentFolders.hasNext() ? studentFolders.next() : rootFolder.createFolder(studentFolderName);

    var bytes = Utilities.base64Decode(base64Data);
    var safeName = String(fileName).replace(/[^A-Z0-9._-]/gi, '_');
    if (!/\.pdf$/i.test(safeName)) safeName += '.pdf';
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'UTC', 'yyyyMMdd_HHmmss');
    var finalName = 'deposit_' + stamp + '_' + safeName;

    var blob = Utilities.newBlob(bytes, 'application/pdf', finalName);
    var file = studentFolder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingErr) {}

    return {
      success: true,
      driveFileId: file.getId(),
      driveUrl: file.getUrl(),
      fileName: finalName,
      uploadedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd HH:mm:ss')
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function voidStudentLedgerEntry(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var txId = String(body.txId || '').trim();
    var studentKey = String(body.studentKey || '').trim();
    var reason = String(body.reason || '').trim();
    var actor = String(body.actorEmail || _toolsCurrentUserEmail_() || 'tools');
    if (!txId) return { success: false, error: 'TX ID required.' };
    if (!studentKey) return { success: false, error: 'Student key required.' };
    if (!reason) return { success: false, error: 'Reason is required to void.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var found = _toolsStudentFindLedgerRowByTxId_(ss, txId);
    if (!found) return { success: false, error: 'Transaction not found: ' + txId };

    var headers = found.headers;
    var row = found.row;
    var rowIndex = found.rowIndex;
    var txTypeIdx = _toolsHeaderIndexFromCandidates_(headers, ['TX_TYPE']);
    var statusIdx = _toolsHeaderIndexFromCandidates_(headers, ['STATUS']);
    var amountIdx = _toolsHeaderIndexFromCandidates_(headers, ['AMOUNT_BRL']);
    var rateIdx = _toolsHeaderIndexFromCandidates_(headers, ['RATE_BRL_HOUR']);
    var studentNameIdx = _toolsHeaderIndexFromCandidates_(headers, ['STUDENT_NAME']);
    var studentEmailIdx = _toolsHeaderIndexFromCandidates_(headers, ['STUDENT_EMAIL']);
    var studentRowIdx = _toolsHeaderIndexFromCandidates_(headers, ['STUDENT_ROW_NUMBER']);

    var txType = txTypeIdx >= 0 ? String(row[txTypeIdx] || '').trim().toUpperCase() : '';
    var currentStatus = statusIdx >= 0 ? String(row[statusIdx] || '').trim().toUpperCase() : '';
    if (txType !== 'CREDIT') return { success: false, error: 'Only CREDIT entries can be voided.' };
    if (currentStatus === 'VOIDED') return { success: true, skipped: true, reason: 'Already voided.' };
    if (currentStatus === 'AMENDED') return { success: false, error: 'Cannot void an amended entry directly.' };

    var originalAmount = Math.abs(amountIdx >= 0 ? Number(row[amountIdx] || 0) : 0);
    var originalRate = rateIdx >= 0 ? Number(row[rateIdx] || 0) : 0;
    var studentName = studentNameIdx >= 0 ? String(row[studentNameIdx] || '').trim() : '';
    var studentEmail = studentEmailIdx >= 0 ? String(row[studentEmailIdx] || '').trim() : '';
    var studentRowNumber = studentRowIdx >= 0 ? Number(row[studentRowIdx] || 0) : 0;

    var nowIso = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd HH:mm:ss');
    var currentBalance = _toolsStudentCurrentBalanceFromLedger_(ss, studentKey, 0);
    var newBalance = _toolsRoundMoney_(currentBalance - originalAmount);

    _toolsAppendStudentLedgerEntry_(ss, {
      studentKey: studentKey,
      studentRowNumber: studentRowNumber,
      studentName: studentName,
      studentEmail: studentEmail,
      txDate: nowIso,
      txType: 'VOID_REVERSAL',
      amountBrl: -originalAmount,
      rateBrlHour: originalRate,
      balanceAfterBrl: newBalance,
      notes: 'Void of ' + txId + '. Reason: ' + reason,
      referenceType: 'VOID_REVERSAL',
      referenceId: txId,
      createdBy: actor
    });
    _toolsStudentMarkLedgerRowStatus_(ss, rowIndex, headers, row, 'VOIDED',
      'Voided ' + nowIso + ' by ' + actor + '. Reason: ' + reason);
    _toolsStudentUpdatePilotBalance_(ss, studentRowNumber, newBalance);

    return { success: true, newBalance: newBalance };
  } catch (ex) {
    return { success: false, error: ex && ex.message ? ex.message : String(ex) };
  }
}

function editStudentLedgerEntry(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var txId = String(body.txId || '').trim();
    var studentKey = String(body.studentKey || '').trim();
    var newAmount = Number(body.newAmount || 0);
    var newRate = Number(body.newRate || 0);
    var newNote = String(body.newNote || '').trim();
    var reason = String(body.reason || '').trim();
    var actor = String(body.actorEmail || _toolsCurrentUserEmail_() || 'tools');
    if (!txId) return { success: false, error: 'TX ID required.' };
    if (!studentKey) return { success: false, error: 'Student key required.' };
    if (!reason) return { success: false, error: 'Reason is required to edit.' };
    if (newAmount <= 0) return { success: false, error: 'New amount must be greater than zero.' };
    if (newRate <= 0) return { success: false, error: 'New rate must be greater than zero.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var found = _toolsStudentFindLedgerRowByTxId_(ss, txId);
    if (!found) return { success: false, error: 'Transaction not found: ' + txId };

    var headers = found.headers;
    var row = found.row;
    var rowIndex = found.rowIndex;
    var txTypeIdx = _toolsHeaderIndexFromCandidates_(headers, ['TX_TYPE']);
    var statusIdx = _toolsHeaderIndexFromCandidates_(headers, ['STATUS']);
    var amountIdx = _toolsHeaderIndexFromCandidates_(headers, ['AMOUNT_BRL']);
    var rateIdx = _toolsHeaderIndexFromCandidates_(headers, ['RATE_BRL_HOUR']);
    var studentNameIdx = _toolsHeaderIndexFromCandidates_(headers, ['STUDENT_NAME']);
    var studentEmailIdx = _toolsHeaderIndexFromCandidates_(headers, ['STUDENT_EMAIL']);
    var studentRowIdx = _toolsHeaderIndexFromCandidates_(headers, ['STUDENT_ROW_NUMBER']);

    var txType = txTypeIdx >= 0 ? String(row[txTypeIdx] || '').trim().toUpperCase() : '';
    var currentStatus = statusIdx >= 0 ? String(row[statusIdx] || '').trim().toUpperCase() : '';
    if (txType !== 'CREDIT') return { success: false, error: 'Only CREDIT entries can be edited.' };
    if (currentStatus === 'VOIDED') return { success: false, error: 'Cannot edit a voided entry.' };
    if (currentStatus === 'AMENDED') return { success: false, error: 'Already amended. Edit the replacement credit instead.' };

    var oldAmount = Math.abs(amountIdx >= 0 ? Number(row[amountIdx] || 0) : 0);
    var oldRate = rateIdx >= 0 ? Number(row[rateIdx] || 0) : 0;
    var studentName = studentNameIdx >= 0 ? String(row[studentNameIdx] || '').trim() : '';
    var studentEmail = studentEmailIdx >= 0 ? String(row[studentEmailIdx] || '').trim() : '';
    var studentRowNumber = studentRowIdx >= 0 ? Number(row[studentRowIdx] || 0) : 0;

    var nowIso = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd HH:mm:ss');
    var currentBalance = _toolsStudentCurrentBalanceFromLedger_(ss, studentKey, 0);
    var balAfterReversal = _toolsRoundMoney_(currentBalance - oldAmount);
    var newBalance = _toolsRoundMoney_(balAfterReversal + newAmount);

    _toolsAppendStudentLedgerEntry_(ss, {
      studentKey: studentKey,
      studentRowNumber: studentRowNumber,
      studentName: studentName,
      studentEmail: studentEmail,
      txDate: nowIso,
      txType: 'EDIT_REVERSAL',
      amountBrl: -oldAmount,
      rateBrlHour: oldRate,
      balanceAfterBrl: balAfterReversal,
      notes: 'Edit reversal of ' + txId + '. Reason: ' + reason,
      referenceType: 'EDIT_REVERSAL',
      referenceId: txId,
      createdBy: actor
    });
    var newTxId = _toolsAppendStudentLedgerEntry_(ss, {
      studentKey: studentKey,
      studentRowNumber: studentRowNumber,
      studentName: studentName,
      studentEmail: studentEmail,
      txDate: nowIso,
      txType: 'CREDIT',
      amountBrl: newAmount,
      rateBrlHour: newRate,
      balanceAfterBrl: newBalance,
      notes: newNote,
      referenceType: 'AMENDED_CREDIT',
      referenceId: txId,
      createdBy: actor
    });
    _toolsStudentMarkLedgerRowStatus_(ss, rowIndex, headers, row, 'AMENDED',
      'Amended ' + nowIso + ' by ' + actor + '. Old: R$' + oldAmount.toFixed(2) + ' @R$' + oldRate.toFixed(2) + '/h. New TX: ' + newTxId + '. Reason: ' + reason);
    _toolsStudentUpdatePilotBalance_(ss, studentRowNumber, newBalance);

    return { success: true, newTxId: newTxId, newBalance: newBalance };
  } catch (ex) {
    return { success: false, error: ex && ex.message ? ex.message : String(ex) };
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

    var emailIdx = _toolsHeaderIndexFromCandidates_(headers, ['PILOT_EMAIL', 'EMAIL', 'E_MAIL']);
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
      PILOT_EMAIL: email,
      EMAIL: email,
      PILOT_NAME: staffName,
      PILOT: staffName,
      PERSON_TYPE: 'STAFF',
      FUNDS_RATE_BRL_HOUR: 0,
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

    var result;
    if (rowNumber >= 2) {
      var current = sh.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
      var merged = headers.map(function(header, idx) {
        var key = _toolsNormHeader_(header);
        return Object.prototype.hasOwnProperty.call(dataMap, key) ? dataMap[key] : current[idx];
      });
      sh.getRange(rowNumber, 1, 1, merged.length).setValues([merged]);
      result = { success: true, action: 'updated', rowNumber: rowNumber };
    } else {
      sh.appendRow(record);
      result = { success: true, action: 'created', rowNumber: sh.getLastRow() };
    }

    _syncStaffIdentityToAuthorizationTables_({
      staffName: staffName,
      email: email,
      staffId: staffId
    });

    return result;
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

function deleteToolsTrainingModule(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var rowNumber = Number(body.rowNumber || 0);
    if (rowNumber < 2) return { success: false, error: 'Invalid row number.' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(APP_SHEETS.TRAINING_MODULES || 'REF_Training_Modules');
    if (!sh) return { success: false, error: 'REF_Training_Modules sheet not found.' };
    if (rowNumber > sh.getLastRow()) return { success: false, error: 'Row ' + rowNumber + ' does not exist.' };
    sh.deleteRow(rowNumber);
    return { success: true, action: 'deleted', rowNumber: rowNumber };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getManeuversForTools() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = APP_SHEETS.MANEUVERS || 'Ref_Maneuvers';
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return { success: true, maneuvers: [] };
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: true, maneuvers: [] };
    var headers = data[0].map(function(h) { return String(h || '').trim().toUpperCase().replace(/\s+/g, '_').replace(/[^A-Z0-9_]/g, ''); });
    var idIdx   = headers.indexOf('MANEUVER_ID');
    var nameIdx = headers.indexOf('NAME');
    var catIdx  = headers.indexOf('CATEGORY');
    var descIdx = headers.indexOf('DESCRIPTION');
    var limIdx  = headers.indexOf('LIMITATIONS');
    var refIdx  = headers.indexOf('REFERENCE');
    var rtIdx   = headers.indexOf('REFERENCE_TYPE');
    var actIdx  = headers.indexOf('ACTIVE');
    var out = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var maneuverIdVal = idIdx >= 0 ? String(row[idIdx] || '').trim() : '';
      if (!maneuverIdVal) continue;
      out.push({
        rowNumber:   i + 1,
        maneuverId:  maneuverIdVal,
        name:        nameIdx >= 0 ? String(row[nameIdx] || '').trim() : '',
        category:    catIdx  >= 0 ? String(row[catIdx]  || '').trim() : '',
        description: descIdx >= 0 ? String(row[descIdx] || '').trim() : '',
        active:      actIdx  >= 0 ? String(row[actIdx]  || '').trim().toUpperCase() !== 'N' : true
      });
    }
    return { success: true, maneuvers: out };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e), maneuvers: [] };
  }
}

function saveToolsManeuver(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var maneuverId = String(body.maneuverId || '').trim().toUpperCase().replace(/\s+/g, '_').replace(/[^A-Z0-9_]/g, '');
    var name = String(body.name || '').trim();
    if (!maneuverId) return { success: false, error: 'Maneuver ID is required' };
    if (!name)       return { success: false, error: 'Name is required' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = APP_SHEETS.MANEUVERS || 'Ref_Maneuvers';
    var sh = ss.getSheetByName(sheetName);
    if (!sh) {
      sh = ss.insertSheet(sheetName);
      sh.appendRow(['MANEUVER_ID', 'NAME', 'CATEGORY', 'DESCRIPTION', 'LIMITATIONS', 'REFERENCE', 'REFERENCE_TYPE', 'ACTIVE']);
    }
    var headers = _toolsSheetHeaderRow_(sh);
    var data = sh.getDataRange().getValues();
    var idIdx = _toolsHeaderIndexFromCandidates_(headers, ['MANEUVER_ID']);
    var rowNumber = Number(body.rowNumber || 0);

    if (rowNumber < 2 && idIdx >= 0) {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][idIdx] || '').trim().toUpperCase() === maneuverId) {
          rowNumber = i + 1;
          break;
        }
      }
    }

    var dataMap = {
      MANEUVER_ID:    maneuverId,
      NAME:           name,
      CATEGORY:       String(body.category    || '').trim().toUpperCase(),
      DESCRIPTION:    String(body.description || '').trim(),
      LIMITATIONS:    String(body.limitations || '').trim(),
      REFERENCE:      String(body.reference   || '').trim(),
      REFERENCE_TYPE: String(body.referenceType || '').trim().toUpperCase(),
      ACTIVE:         _toolsTruthyFlag_(body.active) ? 'Y' : 'N'
    };

    if (rowNumber >= 2) {
      var current = sh.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
      var merged = headers.map(function(header, idx) {
        var key = _toolsNormHeader_(header);
        return Object.prototype.hasOwnProperty.call(dataMap, key) ? dataMap[key] : current[idx];
      });
      sh.getRange(rowNumber, 1, 1, merged.length).setValues([merged]);
      return { success: true, action: 'updated', rowNumber: rowNumber, maneuverId: maneuverId };
    }

    var row = headers.map(function(header) {
      var key = _toolsNormHeader_(header);
      return Object.prototype.hasOwnProperty.call(dataMap, key) ? dataMap[key] : '';
    });
    sh.appendRow(row);
    return { success: true, action: 'created', rowNumber: sh.getLastRow(), maneuverId: maneuverId };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function deleteToolsManeuver(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var rowNumber = Number(body.rowNumber || 0);
    if (rowNumber < 2) return { success: false, error: 'Invalid row number.' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(APP_SHEETS.MANEUVERS || 'Ref_Maneuvers');
    if (!sh) return { success: false, error: 'Ref_Maneuvers sheet not found.' };
    if (rowNumber > sh.getLastRow()) return { success: false, error: 'Row ' + rowNumber + ' does not exist.' };
    sh.deleteRow(rowNumber);
    return { success: true, action: 'deleted', rowNumber: rowNumber };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function ingestLessonPlanFromPdf(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var pdfBase64 = String(body.pdfBase64 || '').trim();
    var mimeType  = String(body.mimeType  || 'application/pdf').trim();
    if (!pdfBase64) return { success: false, error: 'No PDF data provided.' };
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return { success: false, error: 'GEMINI_API_KEY not configured. Add it in GAS Project Settings > Script Properties.' };
    var prompt = [
      'You are an aviation training expert. Extract EVERY evaluation item from this flight training form.',
      'Do NOT summarize. Do NOT merge multiple checklist lines into one.',
      'Include every maneuver, task, question, criterion, and observation line from all pages.',
      'Return ONLY valid JSON (no markdown code blocks) with this exact structure:',
      '{',
      '  "lessonTitle": "full title",',
      '  "category": "VCP or RUNWAY_CHECK or LESSON or ANAC_CHECKRIDE or INTERNAL_CHECKRIDE or ROUTE_CHECK",',
      '  "language": "pt or en",',
      '  "scoringScale": "description of scoring scale",',
      '  "sections": [',
      '    { "sectionId": "I", "sectionTitle": "name",',
      '      "groups": [',
      '        { "groupId": "1", "groupTitle": "name",',
      '          "items": [',
      '            { "name": "exact item text from the form",',
      '              "reference": "regulatory ref e.g. MGO 5.3.4.1",',
      '              "referenceType": "MGO or FAA_ACS or FAA_PTS or ANAC or INTERNAL or OTHER",',
      '              "suggestedLimitations": "operational limits or empty string" }',
      '          ] } ] } ],',
      '  "summaryFields": ["APROVADO", "total hours"],',
      '  "notes": ""',
      '}',
      'Extract EVERY single item without skipping any. If unsure, include the line as an item anyway.'
    ].join('\n');
    var reqBody = {
      contents: [{ parts: [{ text: prompt }, { inlineData: { mimeType: mimeType, data: pdfBase64 } }] }],
      generationConfig: { temperature: 0.1, responseMimeType: 'application/json', maxOutputTokens: 16384 }
    };
    var scriptProps = PropertiesService.getScriptProperties();
    var configuredModel = String(scriptProps.getProperty('GEMINI_MODEL') || '').trim();
    var primaryModel = configuredModel || 'gemini-2.5-flash';

    function _geminiGenerateWithModel_(modelId, bodyToSend) {
      var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(modelId) + ':generateContent?key=' + apiKey;
      return UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(bodyToSend || reqBody),
        muteHttpExceptions: true
      });
    }

    var usedModel = primaryModel;
    var resp = _geminiGenerateWithModel_(usedModel);
    var httpCode = resp.getResponseCode();
    var respText = resp.getContentText();

    // If the configured/default model is unavailable, retry with known supported flash models.
    if (httpCode === 404) {
      var fallbackModels = ['gemini-2.5-flash', 'gemini-1.5-flash', 'gemini-2.0-flash'];
      for (var fm = 0; fm < fallbackModels.length; fm++) {
        var candidate = fallbackModels[fm];
        if (candidate === usedModel) continue;
        var retryResp = _geminiGenerateWithModel_(candidate, reqBody);
        var retryCode = retryResp.getResponseCode();
        var retryText = retryResp.getContentText();
        if (retryCode === 200) {
          usedModel = candidate;
          resp = retryResp;
          httpCode = retryCode;
          respText = retryText;
          break;
        }
      }
    }

    if (httpCode !== 200) {
      var errObj = {}; try { errObj = JSON.parse(respText); } catch (e2) {}
      var msg = (errObj && errObj.error && errObj.error.message) ? errObj.error.message : respText.substring(0, 300);
      return { success: false, error: 'Gemini API error (' + httpCode + ', model ' + usedModel + '): ' + msg };
    }
    var rj = JSON.parse(respText);
    var pts = rj && rj.candidates && rj.candidates[0] && rj.candidates[0].content && rj.candidates[0].content.parts;
    if (!pts || !pts.length) return { success: false, error: 'No response content from Gemini.' };

    function _stripCodeFence_(txt) {
      return String(txt || '').trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '').trim();
    }

    function _extractJsonEnvelope_(txt) {
      var s = String(txt || '');
      var start = s.indexOf('{');
      var end = s.lastIndexOf('}');
      if (start >= 0 && end > start) return s.slice(start, end + 1).trim();
      return s.trim();
    }

    function _parseJsonFromModelText_(txt) {
      var attempts = [];
      var t0 = String(txt || '').trim();
      var t1 = _stripCodeFence_(t0);
      var t2 = _extractJsonEnvelope_(t1);
      if (t0) attempts.push(t0);
      if (t1 && t1 !== t0) attempts.push(t1);
      if (t2 && t2 !== t1) attempts.push(t2);
      var lastErr = null;
      for (var ai = 0; ai < attempts.length; ai++) {
        try {
          return { ok: true, value: JSON.parse(attempts[ai]) };
        } catch (pe) {
          lastErr = pe;
        }
      }
      return { ok: false, error: lastErr, text: t2 || t1 || t0 };
    }

    function _repairJsonWithGemini_(brokenText, modelId) {
      var repairPrompt = [
        'Repair this malformed JSON and return only valid JSON.',
        'Do not remove fields. Do not add markdown. Preserve meaning exactly.',
        '',
        String(brokenText || '')
      ].join('\n');
      var repairBody = {
        contents: [{ parts: [{ text: repairPrompt }] }],
        generationConfig: { temperature: 0, responseMimeType: 'application/json', maxOutputTokens: 16384 }
      };
      var repairResp = _geminiGenerateWithModel_(modelId, repairBody);
      if (repairResp.getResponseCode() !== 200) return null;
      var repairObj = JSON.parse(repairResp.getContentText());
      var repairPts = repairObj && repairObj.candidates && repairObj.candidates[0] && repairObj.candidates[0].content && repairObj.candidates[0].content.parts;
      if (!repairPts || !repairPts.length) return null;
      return String(repairPts.map(function(p) { return String((p && p.text) || ''); }).join('\n') || '').trim();
    }

    var rawJson = String(pts.map(function(p) { return String((p && p.text) || ''); }).join('\n') || '').trim();
    var parsed = _parseJsonFromModelText_(rawJson);
    if (!parsed.ok) {
      var repairedText = null;
      try { repairedText = _repairJsonWithGemini_(parsed.text || rawJson, usedModel); } catch (repairErr) {}
      if (repairedText) parsed = _parseJsonFromModelText_(repairedText);
    }
    if (!parsed.ok) {
      var parseMsg = parsed.error && parsed.error.message ? parsed.error.message : 'Unable to parse model JSON output.';
      return { success: false, error: 'Gemini returned malformed JSON: ' + parseMsg };
    }
    var extracted = parsed.value;

    function _countStructuredItems_(obj) {
      var total = 0;
      var sections = (obj && Array.isArray(obj.sections)) ? obj.sections : [];
      sections.forEach(function(sec) {
        (Array.isArray(sec.groups) ? sec.groups : []).forEach(function(grp) {
          total += (Array.isArray(grp.items) ? grp.items.length : 0);
        });
      });
      return total;
    }

    function _normalizeItemKey_(txt) {
      return String(txt || '').trim().toUpperCase().replace(/\s+/g, ' ');
    }

    function _extractFallbackFlatItems_(modelId) {
      var fallbackPrompt = [
        'Read this full PDF and extract ALL checklist lines exactly as they appear.',
        'Do NOT summarize. Keep one entry per original line/task/question.',
        'Return ONLY valid JSON with this exact shape:',
        '{',
        '  "items": [',
        '    {',
        '      "name": "exact line text",',
        '      "reference": "regulatory reference or empty string",',
        '      "referenceType": "MGO or FAA_ACS or FAA_PTS or ANAC or INTERNAL or OTHER",',
        '      "suggestedLimitations": "limits or empty string"',
        '    }',
        '  ]',
        '}'
      ].join('\n');

      var fbBody = {
        contents: [{ parts: [{ text: fallbackPrompt }, { inlineData: { mimeType: mimeType, data: pdfBase64 } }] }],
        generationConfig: { temperature: 0, responseMimeType: 'application/json', maxOutputTokens: 16384 }
      };
      var fbResp = _geminiGenerateWithModel_(modelId, fbBody);
      if (fbResp.getResponseCode() !== 200) return null;

      var fbObj = JSON.parse(fbResp.getContentText());
      var fbPts = fbObj && fbObj.candidates && fbObj.candidates[0] && fbObj.candidates[0].content && fbObj.candidates[0].content.parts;
      if (!fbPts || !fbPts.length) return null;
      var fbRaw = String(fbPts.map(function(p) { return String((p && p.text) || ''); }).join('\n') || '').trim();
      var fbParsed = _parseJsonFromModelText_(fbRaw);
      return fbParsed.ok ? fbParsed.value : null;
    }

    var extractedItemCount = _countStructuredItems_(extracted);
    var minExpectedRaw = String(scriptProps.getProperty('GEMINI_IMPORT_MIN_ITEMS') || '').trim();
    var minExpectedItems = parseInt(minExpectedRaw, 10);
    if (!isFinite(minExpectedItems) || minExpectedItems < 1) minExpectedItems = 12;

    function _selectBestFallbackFlatItems_(currentModel) {
      var candidates = [currentModel, 'gemini-2.5-flash', 'gemini-1.5-flash', 'gemini-2.0-flash'];
      var uniq = {};
      var best = null;
      var bestCount = 0;
      for (var ci = 0; ci < candidates.length; ci++) {
        var m = String(candidates[ci] || '').trim();
        if (!m || uniq[m]) continue;
        uniq[m] = true;
        try {
          var fb = _extractFallbackFlatItems_(m);
          var fbItems = (fb && Array.isArray(fb.items)) ? fb.items : [];
          if (fbItems.length > bestCount) {
            bestCount = fbItems.length;
            best = fb;
          }
        } catch (e3) {}
      }
      return best;
    }

    if (extractedItemCount < minExpectedItems || extractedItemCount <= 3) {
      var fallback = _selectBestFallbackFlatItems_(usedModel);
      var fallbackItems = (fallback && Array.isArray(fallback.items)) ? fallback.items : [];
      if (fallbackItems.length > 0) {
        if (!Array.isArray(extracted.sections)) extracted.sections = [];
        if (!extracted.sections.length) {
          extracted.sections.push({
            sectionId: 'I',
            sectionTitle: 'Imported Items',
            groups: [{ groupId: '1', groupTitle: 'Checklist', items: [] }]
          });
        }
        if (!Array.isArray(extracted.sections[0].groups)) extracted.sections[0].groups = [];
        if (!extracted.sections[0].groups.length) extracted.sections[0].groups.push({ groupId: '1', groupTitle: 'Checklist', items: [] });
        var targetItems = extracted.sections[0].groups[0].items;
        if (!Array.isArray(targetItems)) {
          extracted.sections[0].groups[0].items = [];
          targetItems = extracted.sections[0].groups[0].items;
        }

        var seenKeys = {};
        extracted.sections.forEach(function(sec) {
          (Array.isArray(sec.groups) ? sec.groups : []).forEach(function(grp) {
            (Array.isArray(grp.items) ? grp.items : []).forEach(function(it) {
              var k = _normalizeItemKey_(it && it.name);
              if (k) seenKeys[k] = true;
            });
          });
        });

        fallbackItems.forEach(function(it) {
          var nm = String((it && it.name) || '').trim();
          var k = _normalizeItemKey_(nm);
          if (!k || seenKeys[k]) return;
          seenKeys[k] = true;
          targetItems.push({
            name: nm,
            reference: String((it && it.reference) || '').trim(),
            referenceType: String((it && it.referenceType) || 'OTHER').trim() || 'OTHER',
            suggestedLimitations: String((it && it.suggestedLimitations) || '').trim()
          });
        });

        extractedItemCount = _countStructuredItems_(extracted);
      }
    }

    return { success: true, extracted: extracted };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function saveLessonPlanFromIngest(payload) {
  try {
    var body         = (payload && typeof payload === 'object') ? payload : {};
    var extracted    = (body.extracted && typeof body.extracted === 'object') ? body.extracted : {};
    var trainingCode = String(body.trainingCode || '').trim().toUpperCase().replace(/[^A-Z0-9_-]/g, '');
    var aircraftType = String(body.aircraftType || '').trim().toUpperCase();
    var description  = String(body.description  || extracted.lessonTitle || '').trim();
    var category     = String(body.category     || extracted.category    || 'VCP').trim();
    var saveToCatalog = body.saveToCatalog !== false;
    if (!trainingCode) return { success: false, error: 'Training code is required.' };

    function _lessonPlanParseJson_(value, fallback) {
      if (!value) return fallback;
      if (typeof value === 'object') return value;
      try { return JSON.parse(String(value || '')); } catch (e) { return fallback; }
    }

    function _lessonPlanNormalizeStop_(stop, seq) {
      var raw = (stop && typeof stop === 'object') ? stop : {};
      var location = String(raw.location || raw.airport || raw.wp || raw.point || '').trim().toUpperCase();
      if (!location) return null;
      var landings = Math.max(0, parseInt(raw.landings, 10) || 0);
      var touchAndGos = Math.max(0, parseInt(raw.touchAndGos, 10) || 0);
      var approaches = Math.max(0, parseInt(raw.approaches, 10) || 0);
      return {
        sequence: Number(raw.sequence || seq || 0) || seq || 1,
        location: location,
        landingType: String(raw.landingType || ((landings > 0 || touchAndGos > 0 || approaches > 0) ? 'FULL_STOP' : 'NONE')).trim().toUpperCase(),
        breakStop: !!raw.breakStop,
        notes: String(raw.notes || '').trim(),
        landings: landings,
        touchAndGos: touchAndGos,
        approaches: approaches
      };
    }

    function _lessonPlanStopsFromRoute_(routeValue) {
      return String(routeValue || '').split(/[,;\n\r]+/).map(function(token) {
        var location = String(token || '').trim().toUpperCase();
        return location ? { location: location, landings: 0, touchAndGos: 0, approaches: 0 } : null;
      }).filter(Boolean);
    }

    var planSource = _lessonPlanParseJson_(body.plan || body.lessonPlan || body.lessonPlanJson, null) || {};
    var explicitStops = Array.isArray(body.stops) ? body.stops
      : (Array.isArray(body.plannedStops) ? body.plannedStops
      : (_lessonPlanParseJson_(body.plannedStopsJson, null) || []));
    var explicitManeuvers = Array.isArray(body.maneuvers) ? body.maneuvers
      : (Array.isArray(body.plannedManeuvers) ? body.plannedManeuvers
      : (_lessonPlanParseJson_(body.plannedManeuversJson, null) || []));
    var explicitCustomManeuvers = Array.isArray(body.customManeuvers) ? body.customManeuvers
      : (Array.isArray(body.plannedCustomManeuvers) ? body.plannedCustomManeuvers
      : (_lessonPlanParseJson_(body.customManeuversJson, null) || []));
    var sections = Array.isArray(extracted.sections) ? extracted.sections : [];
    var allItems = [];
    sections.forEach(function(sec) {
      (Array.isArray(sec.groups) ? sec.groups : []).forEach(function(grp) {
        (Array.isArray(grp.items) ? grp.items : []).forEach(function(item) {
          var nm = String(item.name || '').trim();
          if (!nm) return;
          allItems.push({
            name: nm,
            maneuverId: nm.toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 40),
            category:      String(sec.sectionTitle || '').slice(0, 20),
            groupTitle:    String(grp.groupTitle || ''),
            reference:     String(item.reference || ''),
            referenceType: String(item.referenceType || 'OTHER'),
            limitations:   String(item.suggestedLimitations || item.limitations || '')
          });
        });
      });
    });
    var savedManeuvers = 0;
    if (saveToCatalog && allItems.length > 0) {
      var ss2 = SpreadsheetApp.getActiveSpreadsheet();
      var manSh = ss2.getSheetByName(APP_SHEETS.MANEUVERS || 'Ref_Maneuvers');
      if (!manSh) {
        manSh = ss2.insertSheet(APP_SHEETS.MANEUVERS || 'Ref_Maneuvers');
        manSh.appendRow(['MANEUVER_ID','NAME','CATEGORY','DESCRIPTION','LIMITATIONS','REFERENCE','REFERENCE_TYPE','ACTIVE']);
      }
      var mHdrs = _toolsSheetHeaderRow_(manSh);
      var mData = manSh.getDataRange().getValues();
      var mIdIdx = _toolsHeaderIndexFromCandidates_(mHdrs, ['MANEUVER_ID']);
      var existIds = {};
      for (var ri = 1; ri < mData.length; ri++) {
        var eid = mIdIdx >= 0 ? String(mData[ri][mIdIdx] || '').trim().toUpperCase() : '';
        if (eid) existIds[eid] = true;
      }
      var newRows = [];
      allItems.forEach(function(item) {
        if (!item.maneuverId || existIds[item.maneuverId]) return;
        existIds[item.maneuverId] = true;
        var dm = { MANEUVER_ID: item.maneuverId, NAME: item.name, CATEGORY: item.category,
                   DESCRIPTION: item.groupTitle, LIMITATIONS: item.limitations,
                   REFERENCE: item.reference, REFERENCE_TYPE: item.referenceType, ACTIVE: 'Y' };
        newRows.push(mHdrs.map(function(h) {
          var k = _toolsNormHeader_(h);
          return Object.prototype.hasOwnProperty.call(dm, k) ? dm[k] : '';
        }));
        savedManeuvers++;
      });
      if (newRows.length > 0) {
        manSh.getRange(manSh.getLastRow() + 1, 1, newRows.length, mHdrs.length).setValues(newRows);
      }
    }
    var maneuverIds = []; var seen = {};
    allItems.forEach(function(i) {
      if (i.maneuverId && !seen[i.maneuverId]) { seen[i.maneuverId] = true; maneuverIds.push(i.maneuverId); }
    });
    var planStops = Array.isArray(planSource.stops) && planSource.stops.length ? planSource.stops : explicitStops;
    if ((!Array.isArray(planStops) || !planStops.length) && String(body.route || body.trainingRoute || planSource.route || '').trim()) {
      planStops = _lessonPlanStopsFromRoute_(String(body.route || body.trainingRoute || planSource.route || ''));
    }
    planStops = (Array.isArray(planStops) ? planStops : []).map(function(stop, idx) {
      return _lessonPlanNormalizeStop_(stop, idx + 1);
    }).filter(Boolean);

    var planManeuvers = Array.isArray(planSource.maneuvers) && planSource.maneuvers.length ? planSource.maneuvers.slice() : explicitManeuvers.slice();
    var planCustomManeuvers = Array.isArray(planSource.customManeuvers) && planSource.customManeuvers.length ? planSource.customManeuvers.slice() : explicitCustomManeuvers.slice();
    if (!planManeuvers.length && maneuverIds.length) planManeuvers = maneuverIds.slice();

    var planEnvelope = {
      version: 4, category: category, advisoryOnly: false,
      scoringScale: String(extracted.scoringScale || '0-4'),
      structured: true, sections: sections,
      externalLessonName: String(body.externalLessonName || planSource.externalLessonName || '').trim(),
      routeCheckPrompt: String(body.routeCheckPrompt || planSource.routeCheckPrompt || '').trim(),
      runwayCheckLocation: String(body.runwayCheckLocation || planSource.runwayCheckLocation || '').trim(),
      route: String(body.route || planSource.route || '').trim(),
      stops: planStops,
      maneuvers: planManeuvers.slice(0, 60),
      customManeuvers: planCustomManeuvers
    };
    if (!planEnvelope.route && planEnvelope.stops.length > 0) {
      planEnvelope.route = planEnvelope.stops.map(function(stop) { return stop.location; }).join(', ');
    }
    planEnvelope.plannedLandings = planEnvelope.stops.reduce(function(sum, stop) {
      return sum + Math.max(0, parseInt(stop.landings, 10) || 0) + Math.max(0, parseInt(stop.touchAndGos, 10) || 0);
    }, 0);
    planEnvelope.plannedApproaches = planEnvelope.stops.reduce(function(sum, stop) {
      return sum + Math.max(0, parseInt(stop.approaches, 10) || 0);
    }, 0);
    var syllabusPayload = {
      TRAINING_CODE: trainingCode, AIRCRAFT_TYPE: aircraftType, DESCRIPTION: description,
      LESSON_PLAN_JSON: JSON.stringify(planEnvelope),
      MANEUVERS_JSON: JSON.stringify({ maneuvers: planEnvelope.maneuvers, customManeuvers: planEnvelope.customManeuvers }),
      PLANNED_STOPS_JSON: JSON.stringify(planEnvelope.stops),
      PLANNED_MANEUVERS_JSON: JSON.stringify({ maneuvers: planEnvelope.maneuvers, customManeuvers: planEnvelope.customManeuvers }),
      ROUTE: planEnvelope.route,
      RUNWAY_CHECKOUT: planEnvelope.runwayCheckLocation,
      PLANNED_LANDINGS: String(planEnvelope.plannedLandings || ''),
      PLANNED_APPROACHES: String(planEnvelope.plannedApproaches || '')
    };
    var syllabusRes = addToolsSheetRecord('syllabus', syllabusPayload);
    if (!syllabusRes || !syllabusRes.success) {
      return { success: false, error: (syllabusRes && syllabusRes.error) || 'Failed to save to Ref_Syllabus.' };
    }
    return { success: true, trainingCode: trainingCode, itemCount: allItems.length,
             savedManeuvers: savedManeuvers, syllabusRowNumber: Number(syllabusRes.rowNumber || 0) };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
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

function getToolsStaffReadiness(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var staffEmail = String(body.staffEmail || '').trim().toLowerCase();
    var staffId = String(body.staffId || '').trim();
    var moduleFilter = String(body.moduleId || '').trim().toUpperCase();
    var aircraftType = String(body.aircraftType || '').trim();
    var authorizationRole = String(body.authorizationRole || '').trim();

    var setup = getToolsStaffSetupData();
    if (!setup || !setup.success) return { success: false, error: (setup && setup.error) || 'Could not load staff data' };
    var staff = _toolsFindStaffByEmailOrId_(setup.staff || [], staffEmail, staffId);
    if (!staff) return { success: false, error: 'Staff member not found' };
    var staffRoleSet = _toolsStaffRoleSet_(staff).map(function(role) { return String(role || '').trim().toUpperCase(); });

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
      if (!moduleFilter && roleCode && staffRoleSet.length && staffRoleSet.indexOf(String(roleCode || '').trim().toUpperCase()) < 0) continue;
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
  var authorization = _toolsPilotAuthorizationSummary_(staff, aircraftType, authorizationRole);

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
        primaryRole: staff.primaryRole,
        staffRoles: staff.staffRoles || []
      },
      authorization: authorization,
      summary: {
        moduleCount: rows.length,
        readyCount: readyCount,
        overallReady: rows.length > 0 && readyCount === rows.length && (!authorization.required || authorization.ready)
      },
      rows: rows
    };
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

    const surveyPayload = (payload && payload.survey && typeof payload.survey === 'object') ? payload.survey : {};
    const snapshotPayload = (surveyPayload.runwaySnapshot && typeof surveyPayload.runwaySnapshot === 'object') ? surveyPayload.runwaySnapshot : {};
    const cutdownLabelRaw = String(
      snapshotPayload.cutdownAreaLabel != null ? snapshotPayload.cutdownAreaLabel
        : (surveyPayload.cutdownAreaLabel != null ? surveyPayload.cutdownAreaLabel
          : (payload && payload.official && payload.official.cutdownAreaLabel != null ? payload.official.cutdownAreaLabel : ''))
    ).trim();
    const cutdownSemRestricao = !!(snapshotPayload.cutdownSemRestricao || surveyPayload.cutdownSemRestricao)
      || /^SR\b/i.test(cutdownLabelRaw)
      || /SEM\s*RESTR/i.test(cutdownLabelRaw);
    const cutdownCandidate = Number(
      snapshotPayload.cutdownAreaM != null ? snapshotPayload.cutdownAreaM
        : (surveyPayload.cutdownAreaM != null ? surveyPayload.cutdownAreaM
          : (payload && payload.official && payload.official.cutdownAreaM != null ? payload.official.cutdownAreaM : NaN))
    );
    if (!cutdownSemRestricao && (!isFinite(cutdownCandidate) || cutdownCandidate <= 0)) {
      return { success: false, error: 'Cutdown area is required for runway approval requests.' };
    }

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

function getRunwaySurveyHistory(icao, rwyIdent) {
  try {
    const cleanIcao = String(icao || '').trim().toUpperCase();
    const cleanRwy = String(rwyIdent || '').trim().toUpperCase();
    if (!cleanIcao || !cleanRwy) return { success: false, error: 'ICAO and runway required' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName('DB_Airports');
    if (!dbSheet) return { success: false, error: 'DB_Airports not found' };

    const dbData = dbSheet.getDataRange().getValues();
    if (!dbData.length) return { success: false, error: 'DB_Airports is empty' };
    const cols = _runwayDbFindCols_(dbData[0]);
    if (cols.icao < 0 || cols.runway < 0 || cols.knownFeatures < 0) {
      return { success: false, error: 'DB_Airports missing required columns' };
    }

    // Collect ALL matching rows for this runway pair (e.g. both RWY 09 and RWY 27)
    const pairKey = _runwayPairKey_(cleanRwy);
    const matchedRows = [];
    for (let i = 1; i < dbData.length; i++) {
      const ri = String(dbData[i][cols.icao] || '').trim().toUpperCase();
      if (ri !== cleanIcao) continue;
      const rr = String(dbData[i][cols.runway] || '').trim().toUpperCase();
      const isExact = rr === cleanRwy;
      const isPair  = pairKey && _runwayPairKey_(rr) === pairKey;
      if (isExact || isPair) matchedRows.push({ rowIdx: i, rwy: rr, isPrimary: isExact });
    }

    if (!matchedRows.length) return { success: false, error: 'Runway not found in DB_Airports' };

    // Primary row = the searched direction, fallback to first match
    const primaryMatch = matchedRows.find(function(r) { return r.isPrimary; }) || matchedRows[0];
    const primaryKf = _parseJsonLoose_(String(dbData[primaryMatch.rowIdx][cols.knownFeatures] || '').trim(), {});
    const primaryNorm = Array.isArray(primaryKf) ? { features: primaryKf } : (primaryKf || {});

    // Merge histories from all matched rows, deduplicating by versionId
    const seenVersionIds = {};
    const mergedHistory = [];
    matchedRows.forEach(function(mr) {
      const kf = _parseJsonLoose_(String(dbData[mr.rowIdx][cols.knownFeatures] || '').trim(), {});
      const norm = Array.isArray(kf) ? { features: kf } : (kf || {});
      const hist = Array.isArray(norm.surveyHistory) ? norm.surveyHistory : [];
      hist.forEach(function(entry) {
        if (!entry || typeof entry !== 'object') return;
        const vid = String(entry.versionId || entry.stagingId || '').trim();
        if (vid && seenVersionIds[vid]) return; // skip duplicate
        if (vid) seenVersionIds[vid] = true;
        mergedHistory.push(entry);
      });
    });

    // Sort merged history by publishedAt ascending (oldest first)
    mergedHistory.sort(function(a, b) {
      const ta = new Date(a.publishedAt || 0).getTime();
      const tb = new Date(b.publishedAt || 0).getTime();
      return ta - tb;
    });

    const currentVersion = primaryNorm.currentSurveyVersion || null;
    const activeVerifiedOperational = primaryNorm.verifiedOperational || null;
    const activeVerifiedSurvey = primaryNorm.verifiedSurvey || null;

    // Build the pair label for display (e.g. "09/27")
    const rwyLabels = matchedRows.map(function(r) { return r.rwy; }).sort();
    const rwyPairLabel = rwyLabels.join('/');

    return {
      success: true,
      icao: cleanIcao,
      rwyIdent: cleanRwy,
      rwyPairLabel: rwyPairLabel,
      currentVersion: currentVersion,
      activeVerifiedOperational: activeVerifiedOperational,
      activeVerifiedSurvey: activeVerifiedSurvey,
      history: mergedHistory
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function setActiveRunwaySurveyVersion(icao, rwyIdent, versionId, approvalPassword) {
  try {
    _verifySupervisorApprovalPassword_(approvalPassword);
    const cleanIcao = String(icao || '').trim().toUpperCase();
    const cleanRwy = String(rwyIdent || '').trim().toUpperCase();
    const cleanVersionId = String(versionId || '').trim();
    if (!cleanIcao || !cleanRwy || !cleanVersionId) {
      return { success: false, error: 'ICAO, runway, and versionId required' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName('DB_Airports');
    if (!dbSheet) return { success: false, error: 'DB_Airports not found' };

    const dbData = dbSheet.getDataRange().getValues();
    if (!dbData.length) return { success: false, error: 'DB_Airports is empty' };
    const cols = _runwayDbFindCols_(dbData[0]);
    if (cols.icao < 0 || cols.runway < 0 || cols.knownFeatures < 0) {
      return { success: false, error: 'DB_Airports missing required columns' };
    }

    const pairKey = _runwayPairKey_(cleanRwy);
    const publishRows = [];
    let primaryRow = -1;
    for (let i = 1; i < dbData.length; i++) {
      const ri = String(dbData[i][cols.icao] || '').trim().toUpperCase();
      const rr = String(dbData[i][cols.runway] || '').trim().toUpperCase();
      if (ri !== cleanIcao) continue;
      if (rr === cleanRwy) { primaryRow = i; }
      if ((pairKey && _runwayPairKey_(rr) === pairKey) || rr === cleanRwy) {
        publishRows.push({ rowIndex: i, rowRwy: rr });
      }
    }
    if (primaryRow < 0) return { success: false, error: 'Runway not found in DB_Airports' };
    if (!publishRows.length) publishRows.push({ rowIndex: primaryRow, rowRwy: cleanRwy });

    // Resolve target version from primary row's history
    const primaryKf = _parseJsonLoose_(String(dbData[primaryRow][cols.knownFeatures] || '').trim(), {});
    const primaryNorm = Array.isArray(primaryKf) ? { features: primaryKf } : (primaryKf || {});
    const history = Array.isArray(primaryNorm.surveyHistory) ? primaryNorm.surveyHistory : [];
    const targetEntry = history.find(function(e) { return e && String(e.versionId || '').trim() === cleanVersionId; });
    if (!targetEntry) return { success: false, error: 'Version not found in survey history: ' + cleanVersionId };

    const nowIso = new Date().toISOString();
    publishRows.forEach(function(pr) {
      const rowKf = _parseJsonLoose_(String(dbData[pr.rowIndex][cols.knownFeatures] || '').trim(), {});
      const rowNorm = Array.isArray(rowKf) ? { features: rowKf } : (rowKf || {});

      // Find matching version in this row's history (same stagingId if available, else primary's entry)
      const rowHistory = Array.isArray(rowNorm.surveyHistory) ? rowNorm.surveyHistory : [];
      const rowTarget = rowHistory.find(function(e) { return e && String(e.versionId || '').trim() === cleanVersionId; })
        || rowHistory.find(function(e) { return e && String(e.stagingId || '').trim() === String(targetEntry.stagingId || '').trim(); })
        || targetEntry;

      const nextKf = Object.assign({}, rowNorm, {
        verifiedOperational: rowTarget.verifiedOperational || targetEntry.verifiedOperational || rowNorm.verifiedOperational || {},
        verifiedSurvey: Object.assign({}, rowTarget.verifiedSurvey || targetEntry.verifiedSurvey || {}, {
          reactivatedAt: nowIso,
          reactivatedFromVersionId: cleanVersionId
        }),
        officialReference: rowTarget.officialReference || targetEntry.officialReference || rowNorm.officialReference || {},
        currentSurveyVersion: {
          versionId: cleanVersionId,
          stagingId: String(targetEntry.stagingId || '').trim(),
          publishedAt: String(targetEntry.publishedAt || '').trim(),
          reactivatedAt: nowIso,
          publishedRunway: pr.rowRwy,
          sourceRunway: String(targetEntry.sourceRunway || cleanRwy)
        },
        updatedAt: nowIso
      });

      dbSheet.getRange(pr.rowIndex + 1, cols.knownFeatures + 1).setValue(JSON.stringify(nextKf));
    });

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Active survey version updated to ' + cleanVersionId,
      icao: cleanIcao,
      rwyIdent: cleanRwy,
      versionId: cleanVersionId,
      rowsUpdated: publishRows.length
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function uploadAirstripPhotoToDrive(icao, rwyIdent, base64Data, mimeType, fileName) {
  try {
    var icaoNorm = String(icao || '').trim().toUpperCase();
    var safeFileName = String(fileName || ('airstrip_' + icaoNorm + '.jpg'))
      .replace(/[^a-zA-Z0-9._\-]/g, '_').substring(0, 80);
    var rootFolderName = 'FlightApp_Airstrip_Photos';
    var rootFolders = DriveApp.getFoldersByName(rootFolderName);
    var rootFolder = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder(rootFolderName);
    var subFolders = rootFolder.getFoldersByName(icaoNorm);
    var subFolder = subFolders.hasNext() ? subFolders.next() : rootFolder.createFolder(icaoNorm);
    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', safeFileName);
    var file = subFolder.createFile(blob);

    var sharingWarning = '';
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      sharingWarning = String(shareErr && shareErr.message || shareErr || '');
    }

    var fileId = file.getId();
    var viewUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
    var out = { success: true, url: viewUrl, fileId: fileId };
    if (sharingWarning) out.warning = 'Created file but could not open link-sharing: ' + sharingWarning;
    return out;
  } catch (e) {
    return { success: false, error: String(e && e.message || e) };
  }
}

function getRunwayTakeoffHistory(icao, rwyIdent, limit, targetAircraft) {
  try {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
    if (!logSheet) return { success: true, records: [] };

    var data = logSheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, records: [] };

    var headers = data[0].map(function(h) { return String(h||'').toUpperCase().trim().replace(/\s+/g,'_'); });
    var col = function(name, fallback) { var i = headers.indexOf(name); return i >= 0 ? i : (fallback != null ? fallback : -1); };

    var C = {
      date:         col('DATE',              LOG_FLIGHT_COL.DATE),
      pilot:        col('PILOT',             LOG_FLIGHT_COL.PILOT),
      acft:         col('ACFT',              LOG_FLIGHT_COL.ACFT),
      from:         col('FROM',              LOG_FLIGHT_COL.FROM),
      brakesRel:    col('BRAKES_RELEASE',    LOG_FLIGHT_COL.BRAKES_RELEASE),
      actualLoad:   col('ACTUAL_LOAD_JSON',  LOG_FLIGHT_COL.ACTUAL_LOAD_JSON),
      actualToRoll: col('ACTUAL_TO_ROLL',    LOG_FLIGHT_COL.ACTUAL_TO_ROLL),
      toRiskMatrix: col('TO_RISK_MATRIX',    LOG_FLIGHT_COL.TO_RISK_MATRIX)
    };

    var normalizeRwy = function(v) {
      var t = String(v || '').trim().toUpperCase().replace(/^RWY\s*/, '');
      var m = t.match(/^(\d{1,2})([LCR])?$/);
      if (!m) return t;
      return ('0' + String(parseInt(m[1], 10))).slice(-2) + (m[2] || '');
    };
    var normalizeAcft = function(v) {
      return String(v || '').trim().toUpperCase();
    };

    var icaoNorm  = String(icao || '').trim().toUpperCase();
    var rwyNorm   = normalizeRwy(rwyIdent || '');
    var acftNorm  = normalizeAcft(targetAircraft || '');
    var maxRows   = Math.min(Math.max(Number(limit) || 3, 1), 10);
    var all       = [];

    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      var fromIcao = String(C.from >= 0 ? row[C.from] : '').trim().toUpperCase();
      if (fromIcao !== icaoNorm) continue;

      var weightKg = null;
      var calcToRoll = null;
      var vrDistM = null;
      var actualToRollFromLoad = null;
      var tempC = null;
      var flaps = null;
      var surface = null;
      var wet = null;
      var windDirDeg = null;
      var windSpdKts = null;
      var qnhHpa = null;
      var headTailKts = null;
      var crosswindKts = null;
      var alternates = null;
      var mapDistM = null;
      var rowRwyNorm = '';

      try {
        var loadRaw = C.actualLoad >= 0 ? String(row[C.actualLoad] || '') : '';
        if (loadRaw) {
          var loadObj = JSON.parse(loadRaw);
          var takeoffRoll = (loadObj && typeof loadObj.takeoffRoll === 'object') ? loadObj.takeoffRoll : null;
          var takeoffRollDistM = null;
          if (takeoffRoll) {
            takeoffRollDistM = (takeoffRoll.distanceAtVrM != null && isFinite(Number(takeoffRoll.distanceAtVrM)))
              ? Math.round(Number(takeoffRoll.distanceAtVrM))
              : ((takeoffRoll.distanceAtVrNm != null && isFinite(Number(takeoffRoll.distanceAtVrNm)))
                ? Math.round(Number(takeoffRoll.distanceAtVrNm) * 1852)
                : null);
          }
          rowRwyNorm = normalizeRwy(loadObj.rwyIdent || loadObj.runway || loadObj.departureRunway || loadObj.takeoffRunway || '');
          weightKg   = loadObj.weightKg   != null ? Number(loadObj.weightKg)   : null;
          calcToRoll = loadObj.calcToRoll != null ? Number(loadObj.calcToRoll) : (loadObj.takeoffRollM != null ? Number(loadObj.takeoffRollM) : null);
          actualToRollFromLoad = loadObj.actualToRollM != null ? Number(loadObj.actualToRollM) : takeoffRollDistM;
          vrDistM    = loadObj.vrDistM    != null ? Number(loadObj.vrDistM)    : takeoffRollDistM;
          tempC      = loadObj.tempC      != null ? Number(loadObj.tempC)      : null;
          flaps      = loadObj.flaps      != null ? Number(loadObj.flaps)      : null;
          surface    = loadObj.surface    != null ? String(loadObj.surface)    : null;
          windDirDeg = loadObj.windDir    != null ? Number(loadObj.windDir)    : null;
          windSpdKts = loadObj.windSpd    != null ? Number(loadObj.windSpd)    : null;
          qnhHpa     = loadObj.qnhHpa     != null ? Number(loadObj.qnhHpa)
                      : (loadObj.qnh != null ? Number(loadObj.qnh)
                        : (loadObj.altimeterHpa != null ? Number(loadObj.altimeterHpa) : null));
          headTailKts = loadObj.headTailKts != null ? Number(loadObj.headTailKts) : null;
          crosswindKts = loadObj.crosswindKts != null ? Number(loadObj.crosswindKts) : null;
          alternates = loadObj.alternates != null ? String(loadObj.alternates) : null;
          mapDistM   = loadObj.mapDistM   != null ? Number(loadObj.mapDistM)   : null;
          var hasSurfaceData = loadObj.wet !== undefined || (loadObj.surfaceCondition != null && String(loadObj.surfaceCondition).trim() !== '');
          wet = hasSurfaceData ? (loadObj.wet === true || String(loadObj.surfaceCondition || '').toUpperCase() === 'WET') : null;
        }
      } catch (e) {}

      var actualToRoll = (C.actualToRoll >= 0 && row[C.actualToRoll] !== '' && row[C.actualToRoll] != null)
        ? Number(row[C.actualToRoll])
        : actualToRollFromLoad;

      var dateRaw = C.date >= 0 ? row[C.date] : '';
      var dateStr = '';
      if (dateRaw instanceof Date) dateStr = Utilities.formatDate(dateRaw, 'GMT', 'yyyy-MM-dd');
      else dateStr = String(dateRaw || '').substring(0, 10);

      var acft = String(C.acft >= 0 ? row[C.acft] : '').trim();
      var acftRowNorm = normalizeAcft(acft);
      var runwayMatch = !!rwyNorm && rowRwyNorm === rwyNorm;
      var acftMatch = !!acftNorm && acftRowNorm === acftNorm;

      all.push({
        date:         dateStr,
        pilot:        String(C.pilot >= 0 ? row[C.pilot] : '').trim(),
        acft:         acft,
        weightKg:     weightKg,
        tempC:        tempC,
        qnhHpa:       (qnhHpa != null && isFinite(Number(qnhHpa))) ? Number(qnhHpa) : null,
        flaps:        flaps,
        surface:      surface,
        wet:          wet,
        windDirDeg:   windDirDeg,
        windSpdKts:   windSpdKts,
        headTailKts:  headTailKts,
        crosswindKts: crosswindKts,
        calcToRollM:  calcToRoll,
        actualToRollM:actualToRoll,
        vrDistM:      vrDistM,
        mapDistM:     mapDistM,
        alternates:   alternates,
        _runwayMatch: runwayMatch,
        _acftMatch: acftMatch,
        _weightSort: isFinite(Number(weightKg)) ? Number(weightKg) : -1,
        _recency: i
      });
    }

    if (!all.length) return { success: true, records: [] };

    var selected = [];
    var usedIdx = {};

    // Priority pick: heaviest matching target aircraft at target runway.
    var heavyIdx = -1;
    for (var h = 0; h < all.length; h++) {
      var rec = all[h];
      if (!rec._runwayMatch) continue;
      if (acftNorm && !rec._acftMatch) continue;
      if (heavyIdx < 0) {
        heavyIdx = h;
        continue;
      }
      var cur = all[heavyIdx];
      if (rec._weightSort > cur._weightSort || (rec._weightSort === cur._weightSort && rec._recency > cur._recency)) {
        heavyIdx = h;
      }
    }
    if (heavyIdx >= 0) {
      selected.push(all[heavyIdx]);
      usedIdx[heavyIdx] = true;
    }

    // Fill remaining by strongest relevance then recency.
    var ranked = all.map(function(rec, idx) { return { rec: rec, idx: idx }; });
    ranked.sort(function(a, b) {
      var sa = (a.rec._runwayMatch ? 100 : 0) + (a.rec._acftMatch ? 30 : 0) + Math.max(0, Math.min(20, a.rec._weightSort / 250));
      var sb = (b.rec._runwayMatch ? 100 : 0) + (b.rec._acftMatch ? 30 : 0) + Math.max(0, Math.min(20, b.rec._weightSort / 250));
      if (sb !== sa) return sb - sa;
      return b.rec._recency - a.rec._recency;
    });

    for (var r = 0; r < ranked.length && selected.length < maxRows; r++) {
      var idx = ranked[r].idx;
      if (usedIdx[idx]) continue;
      selected.push(ranked[r].rec);
      usedIdx[idx] = true;
    }

    var clean = selected.map(function(rec) {
      delete rec._runwayMatch;
      delete rec._acftMatch;
      delete rec._weightSort;
      delete rec._recency;
      return rec;
    });

    return { success: true, records: clean };
  } catch (e) {
    return { success: false, error: String(e && e.message || e), records: [] };
  }
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
        required: _aircraftDocsFlag_(row[idx.REQUIRED]) === 'Y',
        critical: _aircraftDocsFlag_(row[idx.CRITICAL]) === 'Y',
        revision: String(row[idx.REVISION] || '').trim(),
        effectiveDate: String(row[idx.EFFECTIVE_DATE] || '').trim(),
        expiryDate: String(row[idx.EXPIRY_DATE] || '').trim(),
        uploadDate: String(row[idx.UPLOAD_DATE] || '').trim(),
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
    rec[idx.EXPIRY_DATE] = String(payload.expiryDate || '').trim();
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

function deleteAircraftDocForTools(request) {
  try {
    var req = (request && typeof request === 'object') ? request : { rowNumber: request };
    var rn = Number(req.rowNumber || 0);
    if (!rn || rn < 2) return { success: false, error: 'Invalid row number' };

    var schema = _ensureAircraftDocsSheet_();
    var sh = schema.sheet;
    var idx = schema.idx;
    var lastRow = sh.getLastRow();
    if (rn > lastRow) return { success: false, error: 'Row ' + rn + ' does not exist' };

    var row = sh.getRange(rn, 1, 1, schema.headers.length).getValues()[0];
    var deleteDriveFile = !!req.deleteDriveFile;
    var driveFileDeleted = false;
    var driveDeleteWarning = '';

    if (deleteDriveFile) {
      var fileId = String(req.driveFileId || row[idx.DRIVE_FILE_ID] || '').trim();
      if (!fileId) {
        var fileUrl = String(req.driveUrl || row[idx.DRIVE_URL] || '').trim();
        fileId = _aircraftDocsDriveIdFromUrl_(fileUrl);
      }
      if (!fileId) {
        driveDeleteWarning = 'Could not find Drive file ID; record was deleted but file was kept.';
      } else {
        try {
          DriveApp.getFileById(fileId).setTrashed(true);
          driveFileDeleted = true;
        } catch (de) {
          driveDeleteWarning = 'Could not move Drive file to trash: ' + (de && de.message ? de.message : String(de));
        }
      }
    }

    sh.deleteRow(rn);
    return {
      success: true,
      deletedRow: rn,
      driveFileDeleted: driveFileDeleted,
      driveDeleteWarning: driveDeleteWarning
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function uploadAircraftDocForTools(payload) {
  try {
    var p = payload || {};
    var tail = _aircraftDocsNormalizeTail_(p.tail);
    if (!tail) return { success: false, error: 'Aircraft tail is required' };

    var base64Data = String(p.base64Data || '').trim();
    var mimeType = String(p.mimeType || 'application/octet-stream').trim();
    var originalFileName = String(p.originalFileName || 'document').trim();
    if (!base64Data) return { success: false, error: 'File data is required' };

    var docType = String(p.docType || '').trim();
    if (!docType) return { success: false, error: 'Doc Type is required' };

    var folderUrl = _aircraftDocsFolderUrlForTail_(tail);
    if (!folderUrl) return { success: false, error: 'No Drive folder found for ' + tail + '. Open the Aircraft panel and save the aircraft to create a folder first.' };

    var folderId = _aircraftDocsDriveIdFromUrl_(folderUrl);
    if (!folderId) return { success: false, error: 'Could not parse folder ID from URL: ' + folderUrl };

    var folder;
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (fe) {
      return { success: false, error: 'Cannot access Drive folder: ' + (fe.message || String(fe)) };
    }

    var normalizedName = _aircraftDocsNormalizeFileName_(
      tail,
      docType,
      String(p.effectiveDate || '').trim(),
      String(p.expiryDate || '').trim(),
      originalFileName
    );

    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, mimeType, normalizedName);
    var file = folder.createFile(blob);
    var driveUrl = file.getUrl();
    var driveFileId = file.getId();

    var uploadDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT', 'yyyy-MM-dd');
    var nowIso = new Date().toISOString();
    var by = String(p.updatedBy || _schedulerCurrentUserEmail_() || 'tools').trim();

    var schema = _ensureAircraftDocsSheet_();
    var sh = schema.sheet;
    var idx = schema.idx;
    var rows = sh.getDataRange().getValues();

    var targetRow = -1;
    for (var i = 1; i < rows.length; i++) {
      var sameTail = _aircraftDocsNormalizeTail_(rows[i][idx.TAIL]) === tail;
      var sameType = String(rows[i][idx.DOC_TYPE] || '').trim().toUpperCase() === docType.toUpperCase();
      var sameName = String(rows[i][idx.DOC_NAME] || '').trim().toUpperCase() === normalizedName.toUpperCase();
      if (sameTail && sameType && sameName) { targetRow = i + 1; break; }
    }

    var rec = [];
    rec[idx.TAIL] = tail;
    rec[idx.DOC_TYPE] = docType;
    rec[idx.DOC_NAME] = normalizedName;
    rec[idx.DRIVE_URL] = driveUrl;
    rec[idx.DRIVE_FILE_ID] = driveFileId;
    rec[idx.REQUIRED] = _aircraftDocsFlag_(p.required);
    rec[idx.CRITICAL] = _aircraftDocsFlag_(p.critical);
    rec[idx.REVISION] = String(p.revision || '').trim();
    rec[idx.EFFECTIVE_DATE] = String(p.effectiveDate || '').trim();
    rec[idx.EXPIRY_DATE] = String(p.expiryDate || '').trim();
    rec[idx.UPLOAD_DATE] = uploadDate;
    rec[idx.LAST_VERIFIED_OFFLINE] = '';
    rec[idx.NOTES] = String(p.notes || '').trim();
    rec[idx.UPDATED_AT] = nowIso;
    rec[idx.UPDATED_BY] = by;
    for (var c = 0; c < schema.headers.length; c++) {
      if (typeof rec[c] === 'undefined') rec[c] = '';
    }

    if (targetRow >= 2) {
      sh.getRange(targetRow, 1, 1, rec.length).setValues([rec]);
    } else {
      sh.appendRow(rec);
    }

    return { success: true, normalizedFileName: normalizedName, driveUrl: driveUrl, driveFileId: driveFileId, uploadDate: uploadDate, action: targetRow >= 2 ? 'updated' : 'created' };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function generateTrainingModuleCourse(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var roleCode      = String(body.roleCode      || '').trim().toUpperCase();
    var moduleType    = String(body.moduleType    || 'INITIAL').trim().toUpperCase();
    var component     = String(body.component     || 'THEORY').trim().toUpperCase();
    var moduleName    = String(body.moduleName    || '').trim();
    var moduleId      = String(body.moduleId      || '').trim();
    var recurrentDays = parseInt(body.recurrentDays || '0', 10) || 0;
    var requiresPractical = _toolsTruthyFlag_(body.requiresPractical);
    var extraContext  = String(body.extraContext   || '').trim();

    var roleNames = { OP_PILOT_LAND:'Piloto Operacional Terrestre', OP_PILOT_ANF:'Piloto Operacional ANF', OP_INSTR_PILOT_LAND:'Piloto Instrutor Terrestre', OP_INSTR_PILOT_ANF:'Piloto Instrutor ANF', FLIGHT_INSTRUCTOR:'Instrutor de Voo', FLIGHT_FOLLOWER:'Acompanhador de Voo', FLIGHT_COORDINATOR:'Coordenador de Voo', FLIGHT_SUPERVISOR:'Supervisor de Voo', MECHANIC:'Mecânico', MECHANIC_TRAINEE:'Mecânico em Treinamento', INSPECTOR:'Inspetor', SRM:'Gerente de Segurança e Risco', STOCKROOM:'Almoxarife' };
    var compNames  = { THEORY:'Teórico', PRACTICAL:'Prático', MIXED:'Teórico-Prático' };
    var typeNames  = { INITIAL:'Inicial', RECURRENT:'Recorrente' };

    var roleName = roleNames[roleCode] || roleCode;
    var compName = compNames[component] || component;
    var typeName = typeNames[moduleType] || moduleType;
    var today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    var docTitle = moduleName || (roleName + ' \u2014 ' + compName + ' ' + typeName);

    // ── Role content definitions ──────────────────────────────────────────────
    var allContent = {};

    allContent['FLIGHT_FOLLOWER'] = {
      description: 'Capacita o acompanhador de voo a exercer controle operacional compartilhado, monitorar voos em tempo real e coordenar com pilotos e base operacional conforme RBAC 135.',
      prerequisites: ['Conhecimento básico de aviação e terminologia aeronáutica', 'Aptidão para comunicação via rádio VHF', 'Familiaridade com o sistema operacional da empresa'],
      duration: '16 horas (teórico) + 4 horas de prática supervisionada',
      objectives: ['Compreender as responsabilidades legais do acompanhador conforme RBAC 135.77', 'Operar sistemas de comunicação e rastreamento de voo', 'Utilizar fraseologia padrão de radiocomunicação ANAC/ICAO', 'Monitorar condições meteorológicas e apoiar decisão de despacho', 'Calcular e controlar combustível e autonomia durante o voo', 'Executar protocolo de emergência e acionamento de SAR', 'Preencher corretamente toda documentação operacional'],
      sections: [
        { title: '1. Papel e Responsabilidades Legais', items: ['Definição de acompanhador de voo conforme RBAC 135.77', 'Controle operacional compartilhado: responsabilidades e limites', 'Autoridade para suspender ou cancelar um voo', 'Cadeia de comunicação: piloto — acompanhador — supervisor', 'Responsabilidade civil e administrativa do acompanhador'] },
        { title: '2. Regulamentação Aplicável', items: ['RBAC 135 — Operações de transporte aéreo não regular', 'RBAC 91 — Regras gerais de voo', 'IS 119-004 — Acompanhamento de voo', 'Manuais operacionais da empresa: MOM e GOM', 'NOTAM e publicações AIP relevantes'] },
        { title: '3. Sistemas de Comunicação e Rastreamento', items: ['Rádio VHF aeronáutico: frequências e canais operacionais', 'Fraseologia padrão ANAC/ICAO para acompanhamento', 'Protocolos de comunicação digital da empresa', 'Rastreamento via ADS-B (FlightAware, FlightRadar24)', 'Checklist de verificação de comunicação pré-voo'] },
        { title: '4. Meteorologia Aplicada', items: ['Fontes de informação: REDEMET, AISWEB, SIGWX', 'Interpretação de METAR, TAF e SIGMET', 'Limites meteorológicos operacionais (mínimos VFR/IFR)', 'Tomada de decisão Go/No-Go com base em meteorologia', 'Fenômenos meteorológicos críticos para a região de operação'] },
        { title: '5. Gestão de Combustível e Autonomia', items: ['Cálculo de consumo horário por tipo de aeronave', 'Verificação de combustível no despacho vs. plano de voo', 'Monitoramento de combustível durante o voo', 'Procedimento de desvio por mínimo de combustível', 'Comunicação proativa do estado de combustível ao piloto'] },
        { title: '6. Planejamento e Acompanhamento Ativo', items: ['Leitura e interpretação do plano de voo operacional', 'Waypoints, rotas e aeródromos alternativos', 'Horários de posição esperada (ETA) e controle de atraso', 'Relatórios de posição periódicos (SELCAL)', 'Identificação de desvio de rota e ação corretiva'] },
        { title: '7. Procedimentos de Emergência', items: ['Protocolo de resposta a emergência declarada pelo piloto', 'Fases INCERFA, ALERFA, DETRESFA — quando e como acionar', 'Acionamento do SAR (Busca e Salvamento)', 'Comunicação com CENIPA e ANAC em ocorrências', 'Coordenação com autoridades locais', 'Estudo de caso: simulação de aeronave em atraso'] },
        { title: '8. Documentação e Registros Operacionais', items: ['Ficha de acompanhamento de voo (preenchimento obrigatório)', 'Diário operacional do acompanhador', 'Registro de combustível e posições por trecho', 'Relatório de ocorrências e irregularidades', 'Uso correto do sistema de controle operacional da empresa'] },
        { title: '9. Segurança e SMS', items: ['Cultura de segurança na função de acompanhamento', 'Identificação e relato de perigos operacionais', 'Just Culture: ambiente de relato sem punição', 'Principais ocorrências históricas na função (lições aprendidas)', 'Papel do acompanhador na prevenção de acidentes'] }
      ],
      normativerefs: ['RBAC 135', 'RBAC 91', 'IS 119-004', 'IS 00-010 (SMS)', 'ICAO Annex 2', 'MOM/GOM da empresa'],
      evalcriteria: 'Avaliação teórica escrita (mínimo 70%). Simulação prática de acompanhamento de voo com cenário de emergência. Avaliação de fraseologia por avaliador designado.'
    };

    allContent['OP_PILOT_LAND'] = {
      description: 'Capacitação para piloto operacional em voos terrestres conforme RBAC 135. Cobre sistemas, procedimentos, regulamentação e CRM aplicados à operação da empresa.',
      prerequisites: ['Licença PPL ou CPL válida', 'Habilitação de tipo na aeronave pertinente', 'Habilitação IFR conforme aplicável à operação'],
      duration: '24 horas (teórico) + dual com instrutor conforme programa de treinamento',
      objectives: ['Dominar os sistemas e limitações da aeronave operada', 'Aplicar procedimentos normais, anormais e de emergência conforme AFM', 'Conhecer e cumprir regulamentação RBAC 135 e 91', 'Executar CRM efetivo em ambiente single e multi-pilot', 'Gerenciar combustível, performance e peso & balanceamento', 'Operar em todas as condições meteorológicas autorizadas'],
      sections: [
        { title: '1. Sistemas da Aeronave', items: ['Estrutura e célula: limites e inspeção pré-voo', 'Sistema de propulsão: operação e limitações', 'Sistemas elétrico, hidráulico e pneumático', 'Aviônica e instrumentos: uso e falhas', 'Sistema de combustível: tanques, válvulas, consumo', 'Limitações operacionais conforme AFM/POH'] },
        { title: '2. Procedimentos Normais e Anormais', items: ['Checklists: pré-voo, partida, voo e pós-voo', 'Procedimentos de decolagem em condições adversas (vento cruzado, densidade altitude)', 'Aproximação e pouso: técnicas e variações', 'Procedimentos anormais conforme AFM', 'MEL — Lista de Equipamentos Mínimos: interpretação e aplicação'] },
        { title: '3. Regulamentação Aplicável', items: ['RBAC 91 — Regras gerais de voo (VFR/IFR)', 'RBAC 135 — Operações de transporte aéreo', 'Duty time e repouso mínimo obrigatório', 'Qualificações requeridas para a operação', 'Documentação obrigatória a bordo'] },
        { title: '4. Navegação e Meteorologia', items: ['Planejamento de rota VFR e IFR', 'Fontes de meteorologia: REDEMET, AISWEB', 'Interpretação de METAR, TAF, SIGMET, AIRMET', 'Mínimos operacionais da empresa', 'Combustível: reservas regulatórias e de política da empresa'] },
        { title: '5. Performance e Peso & Balanceamento', items: ['Cálculo de TOD e LDR em condições variadas', 'Limitações de pista: comprimento, declive, altitude', 'Peso e balanceamento: cálculo e limites', 'Densidade altitude: impacto na performance', 'Casos práticos com aeronaves da frota'] },
        { title: '6. CRM — Crew Resource Management', items: ['Comunicação efetiva piloto-acompanhador-base', 'Tomada de decisão sob pressão (DECIDE model)', 'Gestão de carga de trabalho', 'Situational Awareness: manutenção e recuperação', 'Threat and Error Management (TEM)'] },
        { title: '7. Procedimentos de Emergência', items: ['Falha de motor: procedimentos e pouso de emergência', 'Incêndio a bordo: motor, cabine e elétrico', 'Emergência médica a bordo', 'Comunicação de emergência: MAYDAY e PAN PAN', 'Desorientação espacial: reconhecimento e recuperação'] },
        { title: '8. SMS e Fatores Humanos', items: ['Relato de ocorrências SIPAER', 'Just Culture e ambiente de segurança', 'Fatores humanos: fadiga, estresse, complacência', 'Pressão comercial vs. tomada de decisão segura', 'Lições aprendidas de acidentes relevantes'] }
      ],
      normativerefs: ['RBAC 91', 'RBAC 135', 'IS 135-003', 'AFM/POH da aeronave', 'MOM/GOM da empresa', 'IS 00-010'],
      evalcriteria: 'Prova teórica (mínimo 80%) + checkride com instrutor autorizado. Demonstração de todos os procedimentos de emergência requeridos.'
    };

    allContent['OP_PILOT_ANF'] = allContent['OP_PILOT_LAND'];

    allContent['MECHANIC'] = {
      description: 'Capacitação para mecânico de aviação conforme RBAC 65. Cobre sistemas, documentação, normas de segurança e regulamentação aplicada à manutenção das aeronaves da operação.',
      prerequisites: ['Licença CHT emitida pela ANAC', 'Habilitação na categoria pertinente', 'Experiência mínima conforme programa de treinamento'],
      duration: '20 horas (teórico) + prático supervisionado conforme tarefa',
      objectives: ['Executar manutenção preventiva e corretiva conforme manuais aprovados', 'Preencher documentação de manutenção corretamente (RAD, SIGR)', 'Aplicar normas de segurança e EPI em hangar', 'Utilizar ferramentas e equipamentos de teste corretamente', 'Identificar e registrar discrepâncias conforme procedimentos'],
      sections: [
        { title: '1. Regulamentação de Manutenção', items: ['RBAC 43 — Manutenção, revisões e reparos', 'RBAC 65 — Certificação de técnicos aeronáuticos', 'RBAC 145 — Organizações de manutenção aprovadas', 'Hierarquia documental: MM, IPC, SB, AD/ICA', 'Responsabilidades do mecânico certificado'] },
        { title: '2. Sistemas da Aeronave', items: ['Célula e estrutura: pontos de inspeção', 'Motor e hélice: operação e manutenção', 'Trem de pouso e freios', 'Sistema hidráulico', 'Sistema elétrico e aviônica', 'Sistema de combustível e óleo'] },
        { title: '3. Documentação e Rastreabilidade', items: ['RAD — Registro de Aeronave e Diário de Bordo', 'Ficha de manutenção programada', 'Autorização de Retorno ao Serviço (ARS)', 'SIGR — Registro de discrepâncias', 'Rastreabilidade de peças e materiais aeronáuticos', 'Controle de vida útil de componentes (time-limited parts)'] },
        { title: '4. Ferramentas e Equipamentos', items: ['Ferramentas de torque: uso e calibração', 'Equipamentos de medição elétrica', 'Ferramentas especiais (SE) e controle', 'Programa de calibração de instrumentos', 'Segurança no manuseio de fluidos aeronáuticos'] },
        { title: '5. Segurança no Hangar', items: ['EPI obrigatório por tipo de tarefa', 'Plano de emergência do hangar', 'Prevenção de FOD (Foreign Object Damage)', 'Operação segura de aeronave em solo (pushback, reboque)', 'Manuseio e descarte de materiais inflamáveis e perigosos'] },
        { title: '6. Controle de Qualidade e Inspeção', items: ['Inspeção dupla (two-person rule) em itens críticos', 'Signing off de tarefas de manutenção', 'Revisão independente de trabalho realizado', 'Limites de inspeção: TBO, TBR e intervalos calendário', 'Comunicação de anomalias com o inspetor responsável'] }
      ],
      normativerefs: ['RBAC 43', 'RBAC 65', 'RBAC 145', 'Manuais do fabricante (MM, IPC)', 'ADs e ICAs aplicáveis', 'Manual de Manutenção da empresa'],
      evalcriteria: 'Avaliação escrita (mínimo 75%) + avaliação prática supervisionada em tarefa real ou simulada.'
    };

    allContent['INSPECTOR'] = {
      description: 'Capacitação para inspetor de aeronave. Cobre autoridade de inspeção, critérios de aceitação, documentação e regulamentação RBAC 145.',
      prerequisites: ['Licença CHT com habilitação de inspeção (CHT-I)', 'Experiência mínima conforme RBAC 65', 'Autorização de inspetor emitida pela organização'],
      duration: '16 horas (teórico) + prático supervisionado com inspetor sênior',
      objectives: ['Exercer a autoridade de inspeção com independência técnica', 'Aplicar critérios de aceitação conforme manuais aprovados', 'Documentar inspeções e não-conformidades corretamente', 'Emitir autorização de retorno ao serviço com respaldo regulatório'],
      sections: [
        { title: '1. Autoridade e Responsabilidade', items: ['Definição legal de inspetor conforme RBAC 145', 'Independência funcional do inspetor', 'Responsabilidade civil e criminal na função', 'Limites de autoridade de inspeção', 'Comunicação de não-conformidades à direção'] },
        { title: '2. Critérios de Aceitação', items: ['Tolerâncias conforme manuais do fabricante', 'Critérios de aceitação/rejeição para corrosão', 'Avaliação de danos estruturais e reparos aprovados', 'Substituição obrigatória vs. reparo permitido', 'Componentes de vida limitada: controle e substituição'] },
        { title: '3. Documentação de Inspeção', items: ['Relatório de inspeção: elementos obrigatórios', 'RGA — Registro de Grau de Anomalia', 'ARS — Autorização de Retorno ao Serviço', 'Arquivo e retenção de registros de inspeção', 'Auditoria interna e externa de registros'] },
        { title: '4. Inspeções Regulatórias', items: ['Inspeção de 100 horas: escopo e documentação', 'Inspeção anual: diferenças e exigências adicionais', 'Inspeção de retorno de longa paragem', 'Verificação de Aeronavegabilidade (VA)', 'Coordenação com ANAC para inspeções especiais'] }
      ],
      normativerefs: ['RBAC 43', 'RBAC 65', 'RBAC 145', 'Manuais do fabricante', 'ADs e ICAs aplicáveis'],
      evalcriteria: 'Avaliação escrita (mínimo 80%) + inspeção prática com parecer de inspetor sênior.'
    };

    allContent['FLIGHT_COORDINATOR'] = {
      description: 'Capacitação para coordenador de voo no exercício de controle operacional compartilhado conforme RBAC 135. Cobre planejamento, despacho, documentação e gestão da operação.',
      prerequisites: ['Experiência operacional na empresa', 'Conhecimento de regulamentação RBAC 135', 'Autorização interna para controle operacional'],
      duration: '12 horas (teórico) + acompanhamento prático com coordenador sênior',
      objectives: ['Planejar e despachar voos conforme manuais operacionais', 'Exercer controle operacional compartilhado com o piloto', 'Gerenciar documentação de despacho', 'Tomar decisão Go/No-Go com base em dados operacionais', 'Coordenar resposta a emergências operacionais'],
      sections: [
        { title: '1. Controle Operacional Compartilhado', items: ['Definição e responsabilidades conforme RBAC 135', 'Autoridade do coordenador vs. piloto em comando', 'Tomada de decisão compartilhada', 'Comunicação com a tripulação antes e durante o voo', 'Ação corretiva e cancelamento de voo'] },
        { title: '2. Planejamento de Voo', items: ['Plano de voo ICAO e simplificado', 'Roteamento, waypoints e aeródromos alternativos', 'NOTAM: pesquisa e interpretação', 'Meteorologia aplicada ao despacho', 'Cálculo de combustível, reservas e autonomia'] },
        { title: '3. Documentação de Despacho', items: ['Folha de despacho: elementos obrigatórios', 'Manifesto de passageiros e carga', 'Verificação de documentos da aeronave (CAN, IEM, RA)', 'Controle de documentos e habilitações da tripulação', 'Registros de operação e arquivo'] },
        { title: '4. Gestão de Tripulação', items: ['Controle de duty time e repouso mínimo', 'Verificação de qualificações vigentes', 'Substituição de tripulante em casos de indisponibilidade', 'Briefing operacional com o piloto'] },
        { title: '5. Emergências Operacionais', items: ['Protocolo de aeronave em atraso', 'Acionamento SAR: responsabilidades do coordenador', 'Comunicação com DECEA e ANAC', 'Gestão de crise e comunicação interna', 'Relatório pós-evento'] }
      ],
      normativerefs: ['RBAC 135', 'RBAC 91', 'IS 135-003', 'MOM/GOM da empresa', 'IS 00-010'],
      evalcriteria: 'Avaliação escrita (mínimo 75%) + simulação de despacho de voo com cenário de emergência.'
    };

    allContent['FLIGHT_SUPERVISOR'] = {
      description: 'Capacitação para supervisor de voo. Cobre supervisão da operação, gestão de risco, tomada de decisão e cumprimento de normas operacionais.',
      prerequisites: ['Experiência prévia em funções operacionais (mínimo 1 ano)', 'Treinamento de coordenador ou equivalente', 'Autorização de supervisor emitida pela empresa'],
      duration: '16 horas (teórico) + acompanhamento prático de turno completo',
      objectives: ['Supervisionar a operação diária com foco em segurança', 'Gerenciar riscos operacionais e tomar decisões sob pressão', 'Controlar duty time e prontidão da tripulação', 'Coordenar resposta a emergências', 'Garantir conformidade com manuais e regulamentação'],
      sections: [
        { title: '1. Papel do Supervisor de Voo', items: ['Responsabilidades conforme manuais operacionais', 'Autoridade e limites da função', 'Interface com direção operacional e ANAC', 'Turnover e passagem de turno: elementos críticos'] },
        { title: '2. Gestão de Risco Operacional', items: ['Briefing de risco diário: metodologia e condução', 'Ferramentas de TRM (Threat and Risk Management)', 'Decisão Go/No-Go de alto nível operacional', 'Pressão comercial vs. cultura de segurança', 'Documentação de decisões operacionais críticas'] },
        { title: '3. Controle de Tripulação e Frota', items: ['Status de prontidão da frota em tempo real', 'Controle de qualificações e habilitações vigentes', 'Gestão de duty time: controle e registros', 'Substituição emergencial de tripulante', 'Liberação de aeronave com MEL ativo'] },
        { title: '4. Gerenciamento de Emergências', items: ['Plano de Emergência Operacional (PEO)', 'Acionamento de SAR: responsabilidades do supervisor', 'Comunicação com ANAC, DECEA e CENIPA', 'Gestão de crise: comunicação interna e externa', 'Relatório pós-evento e lições aprendidas'] },
        { title: '5. Conformidade e Auditorias', items: ['Inspeção de linha (observação operacional)', 'Monitoramento de conformidade operacional', 'Relato e tratamento de não-conformidades', 'Ações corretivas e preventivas (CAPA)', 'Preparação para inspeção da ANAC'] }
      ],
      normativerefs: ['RBAC 135', 'RBAC 91', 'IS 135-003', 'IS 00-010', 'MOM/GOM da empresa', 'Plano de Emergência Operacional'],
      evalcriteria: 'Avaliação escrita (mínimo 80%) + simulação de gestão de emergência operacional com avaliador.'
    };

    allContent['SRM'] = {
      description: 'Capacitação para Gerente de Segurança e Risco (SRM) conforme IS 00-010 e requisitos do PGSSO. Cobre SMS, identificação de perigos, avaliação de risco e cultura de segurança.',
      prerequisites: ['Experiência operacional relevante (mínimo 2 anos)', 'Conhecimento de regulamentação ANAC', 'Indicação pela diretoria da empresa'],
      duration: '20 horas (teórico) + workshop prático de análise de risco',
      objectives: ['Implementar e manter o Sistema de Gestão de Segurança (SGSO)', 'Identificar perigos e avaliar riscos sistematicamente', 'Desenvolver cultura de segurança e Just Culture', 'Investigar ocorrências e propor ações corretivas', 'Preparar relatórios de segurança para a ANAC'],
      sections: [
        { title: '1. Sistema de Gestão de Segurança (SMS/SGSO)', items: ['Quatro pilares do SMS (ICAO Doc 9859)', 'Política de segurança: elaboração e divulgação', 'Objetivos e indicadores de segurança (KPIs de segurança)', 'Plano de Segurança Operacional (PSO)', 'Responsabilidades executivas: Accountable Manager'] },
        { title: '2. Identificação de Perigos', items: ['Métodos: LOSA, observação, análise de dados históricos', 'Registro de perigos (hazard log)', 'Fontes de dados: SIPAER, relatórios internos, auditorias', 'Análise de tendências e indicadores preditivos', 'Perigos específicos da operação (ex: operação em área remota, voos noturnos)'] },
        { title: '3. Avaliação e Mitigação de Riscos', items: ['Matriz de risco: probabilidade × severidade', 'Hierarquia de controles de segurança', 'Modelo Swiss Cheese (James Reason)', 'ALARP — tão baixo quanto razoavelmente praticável', 'Documentação, revisão e aprovação de riscos'] },
        { title: '4. Cultura de Segurança e Just Culture', items: ['Definição e importância da cultura de segurança', 'Just Culture: relato sem punição automática', 'Estratégias de comunicação de segurança', 'Treinamento de conscientização para toda a equipe', 'Medição e melhoria da cultura de segurança'] },
        { title: '5. Investigação de Ocorrências', items: ['Tipos: acidente, incidente grave, incidente e ocorrência de solo', 'Metodologia: análise de causa raiz, HFACS, Bow-Tie', 'Coordenação com CENIPA e ANAC', 'Relatório interno: elementos obrigatórios', 'CAPA — Ações Corretivas e Preventivas'] },
        { title: '6. Conformidade Regulatória SMS', items: ['IS 00-010 — Requisitos do PGSSO', 'Relatórios periódicos de segurança à ANAC', 'Auditoria interna do SMS', 'Avaliação de efetividade do SMS (revisão gerencial)', 'Integração SMS com o sistema de qualidade da empresa'] }
      ],
      normativerefs: ['IS 00-010', 'ICAO Doc 9859', 'RBAC 91', 'RBAC 135', 'IS 175-007', 'ICAO Annex 13'],
      evalcriteria: 'Avaliação escrita (mínimo 80%) + apresentação de análise de risco operacional real ou hipotético para a banca avaliadora.'
    };

    allContent['FLIGHT_INSTRUCTOR'] = {
      description: 'Capacitação para instrutor de voo conforme RBAC 61. Cobre técnicas de instrução, planejamento curricular, avaliação de alunos e responsabilidades do instrutor.',
      prerequisites: ['Certificado de instrutor de voo (CFI) válido', 'Habilitação de tipo na aeronave pertinente', 'Experiência mínima conforme RBAC 61'],
      duration: '16 horas (teórico) + checkride com instrutor verificador',
      objectives: ['Planejar e conduzir sessões de instrução efetivas', 'Avaliar competências de alunos de forma objetiva', 'Manter registros de treinamento conforme regulamentação', 'Garantir padrão de segurança durante toda a instrução'],
      sections: [
        { title: '1. Fundamentos da Instrução', items: ['Princípios de aprendizagem adulta (andragogia)', 'Estilos de aprendizado: visual, auditivo, cinestésico', 'Técnicas de briefing e debriefing efetivo', 'Feedback construtivo: como e quando dar', 'Gerenciamento de erros do aluno durante o voo'] },
        { title: '2. Planejamento Curricular', items: ['Objetivos de aprendizado (SMART)', 'Sequência lógica de tópicos e progressão', 'Carga horária por tópico e manobra', 'Material de apoio: apostilas, checklists, referências', 'Adaptação do currículo para diferentes níveis de aluno'] },
        { title: '3. Instrução em Solo', items: ['Técnicas de briefing pré-voo', 'Uso de simuladores, maquetes e quadro branco', 'Avaliação teórica escrita e oral', 'Preparação mental do aluno para o voo', 'Resolução de dúvidas técnicas e conceituais'] },
        { title: '4. Instrução em Voo', items: ['Técnica de transferência de controles', 'Demonstração de manobras com narração', 'Correção de erros em tempo real: quando intervir', 'Gestão da carga de trabalho do aluno', 'Segurança durante instrução: limites de intervenção'] },
        { title: '5. Avaliação e Documentação', items: ['Critérios de aprovação por manobra (ACS/PTS ANAC)', 'Registro de progresso e horas de instrução', 'Recomendação para checkride com examinador', 'Manutenção de arquivos de treinamento', 'Responsabilidade do instrutor na documentação obrigatória'] }
      ],
      normativerefs: ['RBAC 61', 'RBAC 91', 'ACS/PTS aplicável da ANAC', 'MOM da empresa', 'ICAO Doc 9868'],
      evalcriteria: 'Demonstração de briefing e condução de instrução de manobra com avaliador designado. Revisão de plano de aula e registros.'
    };

    allContent['STOCKROOM'] = {
      description: 'Capacitação para almoxarife de aviação. Cobre controle de estoque, rastreabilidade de peças aeronáuticas, documentação e normas de armazenagem.',
      prerequisites: ['Orientação inicial da empresa', 'Conhecimento básico de informática e sistemas de controle'],
      duration: '8 horas (teórico) + prático supervisionado',
      objectives: ['Controlar estoque de peças e materiais aeronáuticos', 'Manter rastreabilidade de componentes certificados', 'Documentar entrada e saída de materiais corretamente', 'Garantir armazenagem correta conforme fabricante e normas'],
      sections: [
        { title: '1. Componentes Aeronáuticos', items: ['Partes certificadas vs. consumíveis', 'Números de parte (P/N) e número de série (S/N)', 'Certificados de conformidade: FAA 8130-3 e EASA Form 1', 'Componentes de vida limitada (life-limited parts)', 'Identificação e segregação de peças suspeitas ou piratas'] },
        { title: '2. Rastreabilidade e Documentação', items: ['Controle de tags e etiquetas de identificação', 'Registro de entrada: fornecedor, certificado, data', 'Registro de saída: aeronave, O/S, mecânico responsável', 'Arquivo físico e digital de certificados', 'Coordenação com mecânico e inspetor'] },
        { title: '3. Armazenagem Correta', items: ['Condições ambientais: temperatura e umidade', 'Proteção anticorrosão e embalagem original', 'Segregação de materiais incompatíveis', 'FIFO — First In First Out: aplicação prática', 'Controle de validade de materiais perecíveis e selantes'] },
        { title: '4. Controle de Estoque e Reposição', items: ['Nível mínimo e máximo de estoque por item', 'Processo de pedido de reposição', 'Inventário periódico: metodologia e registro', 'Controle de ferramentas calibradas em empréstimo', 'Uso do sistema de gestão de almoxarifado da empresa'] }
      ],
      normativerefs: ['RBAC 43', 'RBAC 145', 'Manuais do fabricante', 'Procedimentos internos da empresa'],
      evalcriteria: 'Avaliação escrita (mínimo 70%) + simulação de processo completo de entrada e saída de componente aeronáutico.'
    };

    // Default for unmapped roles
    allContent['_DEFAULT_'] = {
      description: 'Módulo de treinamento para a função ' + roleName + '. Gerado automaticamente — revisar e complementar conforme realidade operacional.',
      prerequisites: ['Experiência na função ou área correlata', 'Orientação inicial da empresa'],
      duration: 'A definir conforme escopo do módulo',
      objectives: ['Conhecer as responsabilidades e escopo da função', 'Cumprir os requisitos regulatórios aplicáveis', 'Operar com segurança dentro do escopo da função'],
      sections: [
        { title: '1. Fundamentos da Função', items: ['Responsabilidades e escopo de atuação', 'Regulamentação aplicável', 'Interface com outras funções da operação'] },
        { title: '2. Procedimentos Operacionais', items: ['Procedimentos normais', 'Procedimentos de emergência', 'Documentação e registros obrigatórios'] },
        { title: '3. Segurança Operacional', items: ['Cultura de segurança e Just Culture', 'Identificação e relato de perigos', 'Lições aprendidas aplicáveis'] }
      ],
      normativerefs: ['Regulamentação ANAC aplicável', 'MOM/GOM da empresa', 'IS 00-010'],
      evalcriteria: 'Avaliação escrita (mínimo 70%) e avaliação prática conforme padrão da empresa.'
    };

    var content = allContent[roleCode] || allContent['_DEFAULT_'];

    // Adapt for recurrent
    if (moduleType === 'RECURRENT') {
      var recDesc = 'MÓDULO DE RECICLAGEM — ' + content.description;
      var recObj  = ['Rever e atualizar conhecimentos adquiridos no treinamento inicial'].concat(content.objectives);
      var recSecs = content.sections.slice();
      recSecs.push({ title: (recSecs.length + 1) + '. Atualização Regulatória e Lições Aprendidas', items: [
        'Alterações regulatórias desde o último treinamento' + (recurrentDays > 0 ? ' (ciclo de ' + recurrentDays + ' dias)' : ''),
        'Ocorrências e acidentes recentes relevantes para a função',
        'Atualizações nos manuais e procedimentos da empresa',
        'Lições aprendidas do período anterior',
        'Revisão dos pontos de maior dificuldade — Q&A'
      ]});
      content = { description: recDesc, prerequisites: content.prerequisites, duration: content.duration, objectives: recObj, sections: recSecs, normativerefs: content.normativerefs, evalcriteria: content.evalcriteria };
    }

    // Add extra context section
    if (extraContext) {
      var extraSecs = content.sections.slice();
      extraSecs.push({ title: (extraSecs.length + 1) + '. Informações Específicas da Operação', items: extraContext.split('\n').filter(function(l) { return l.trim(); }) });
      content = { description: content.description, prerequisites: content.prerequisites, duration: content.duration, objectives: content.objectives, sections: extraSecs, normativerefs: content.normativerefs, evalcriteria: content.evalcriteria };
    }

    // ── Create Google Doc ─────────────────────────────────────────────────────
    var fullTitle = docTitle + ' \u2014 Plano de Aula';
    var doc  = DocumentApp.create(fullTitle);
    var body = doc.getBody();
    body.clear();

    body.appendParagraph(docTitle).setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph('Plano de Aula \u2014 ' + typeName + ' ' + compName).setHeading(DocumentApp.ParagraphHeading.SUBTITLE);

    var metaLines = [
      'Fun\u00e7\u00e3o: ' + roleName,
      'M\u00f3dulo ID: ' + (moduleId || 'N/D') + '   |   Tipo: ' + typeName + '   |   Componente: ' + compName,
      recurrentDays > 0 ? 'Validade: ' + recurrentDays + ' dias' : 'Validade: Sem recorr\u00eancia definida',
      'Avalia\u00e7\u00e3o Pr\u00e1tica: ' + (requiresPractical ? 'Exigida' : 'N\u00e3o exigida'),
      'Data de Gera\u00e7\u00e3o: ' + today,
      'ATEN\u00c7\u00c3O: Documento gerado automaticamente. Revisar antes de usar.'
    ];
    metaLines.forEach(function(line) { body.appendParagraph(line); });

    body.appendParagraph('Descri\u00e7\u00e3o do M\u00f3dulo').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(content.description);

    body.appendParagraph('Carga Hor\u00e1ria').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(content.duration);

    body.appendParagraph('Pr\u00e9-requisitos').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    content.prerequisites.forEach(function(p) { body.appendListItem(p).setGlyphType(DocumentApp.GlyphType.BULLET); });

    body.appendParagraph('Objetivos de Aprendizado').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    content.objectives.forEach(function(o, idx) { body.appendListItem((idx + 1) + '. ' + o).setGlyphType(DocumentApp.GlyphType.BULLET); });

    body.appendParagraph('Conte\u00fado Program\u00e1tico').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    content.sections.forEach(function(sec) {
      body.appendParagraph(sec.title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      sec.items.forEach(function(item) { body.appendListItem(item).setGlyphType(DocumentApp.GlyphType.BULLET); });
    });

    body.appendParagraph('Crit\u00e9rios de Avalia\u00e7\u00e3o').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(content.evalcriteria);

    body.appendParagraph('Refer\u00eancias Normativas').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    content.normativerefs.forEach(function(ref) { body.appendListItem(ref).setGlyphType(DocumentApp.GlyphType.BULLET); });

    body.appendParagraph('Notas para o Instrutor').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    var notes = [
      'Este documento foi gerado automaticamente pelo sistema de treinamento.',
      'O instrutor respons\u00e1vel deve revisar e adaptar o conte\u00fado \u00e0 realidade operacional atual.',
      'Incluir exemplos reais da opera\u00e7\u00e3o e ocorr\u00eancias locais sempre que poss\u00edvel.',
      'Manter registros de presen\u00e7a e avalia\u00e7\u00e3o de todos os participantes.',
      'Versionar este documento conforme SOP de controle documental da empresa.'
    ];
    notes.forEach(function(note) { body.appendListItem(note).setGlyphType(DocumentApp.GlyphType.BULLET); });

    body.appendParagraph('Controle de Aprova\u00e7\u00e3o').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    ['Elaborado por: ________________________  Data: ___/___/_____',
     'Revisado por:  ________________________  Data: ___/___/_____',
     'Aprovado por:  ________________________  Data: ___/___/_____',
     '',
     'Vers\u00e3o: 1.0   |   Data de vig\u00eancia: ___/___/_____'
    ].forEach(function(line) { body.appendParagraph(line); });

    doc.saveAndClose();
    return { success: true, docUrl: doc.getUrl(), docTitle: doc.getName() };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function pushLivePosition(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var reg = String(body.reg || '').trim().toUpperCase();
    if (!reg) return { success: false, error: 'reg required' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = APP_SHEETS.LIVE_TRACK || 'LIVE_Track';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['REG', 'MISSION_ID', 'LAT', 'LNG', 'BEARING', 'GS_KTS', 'UPDATED_AT']);
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h || '').trim().toUpperCase(); });
    var col = function(name) { return headers.indexOf(name); };
    var ri = col('REG'), mi = col('MISSION_ID'), lati = col('LAT'), lngi = col('LNG');
    var beari = col('BEARING'), gsi = col('GS_KTS'), tsi = col('UPDATED_AT');

    var now = new Date();
    var row = headers.map(function() { return ''; });
    if (ri >= 0)   row[ri]   = reg;
    
    // Extract missionId correctly: if body.missionId looks like a full flight leg ID (has 3+ parts),
    // extract just the mission ID portion (first 2 parts)
    var missionId = String(body.missionId || '').trim();
    if (missionId) {
      var parts = missionId.split('-');
      if (parts.length > 2) {
        // This looks like a full flight leg ID (e.g., "OFL240422-12345-01"), extract mission ID
        missionId = parts[0] + '-' + parts[1];
      }
    }
    
    if (mi >= 0)   row[mi]   = missionId;
    if (lati >= 0) row[lati] = Number(body.lat  || 0);
    if (lngi >= 0) row[lngi] = Number(body.lng  || 0);
    if (beari >= 0) row[beari] = Number(body.bearing || 0);
    if (gsi >= 0)  row[gsi]  = Number(body.gsKts || 0);
    if (tsi >= 0)  row[tsi]  = now;

    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][ri] || '').trim().toUpperCase() === reg) { foundRow = i + 1; break; }
    }
    if (foundRow > 0) {
      sheet.getRange(foundRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getLivePositions() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(APP_SHEETS.LIVE_TRACK || 'LIVE_Track');
    if (!sheet) return { success: true, positions: [] };
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, positions: [] };

    var headers = data[0].map(function(h) { return String(h || '').trim().toUpperCase(); });
    var col = function(name) { return headers.indexOf(name); };
    var ri = col('REG'), mi = col('MISSION_ID'), lati = col('LAT'), lngi = col('LNG');
    var beari = col('BEARING'), gsi = col('GS_KTS'), tsi = col('UPDATED_AT');

    var positions = [];
    for (var i = 1; i < data.length; i++) {
      var reg = String(data[i][ri] || '').trim().toUpperCase();
      if (!reg) continue;
      var tsVal = tsi >= 0 && data[i][tsi] ? data[i][tsi] : null;
      var tsMs = tsVal ? (tsVal instanceof Date ? tsVal.getTime() : new Date(tsVal).getTime()) : 0;
      positions.push({
        reg:       reg,
        missionId: mi >= 0 ? String(data[i][mi] || '') : '',
        lat:       lati >= 0 ? Number(data[i][lati] || 0) : 0,
        lng:       lngi >= 0 ? Number(data[i][lngi] || 0) : 0,
        bearing:   beari >= 0 ? Number(data[i][beari] || 0) : 0,
        gsKts:     gsi >= 0 ? Number(data[i][gsi] || 0) : 0,
        updatedAtMs: tsMs
      });
    }
    return { success: true, positions: positions };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function testHistoricalGeminiConnectivity() {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      return { success: false, error: 'GEMINI_API_KEY not configured.' };
    }

    var scriptProps = PropertiesService.getScriptProperties();
    var configuredModel = String(scriptProps.getProperty('GEMINI_MODEL') || '').trim();
    var primaryModel = configuredModel || 'gemini-2.5-flash';

    var body = {
      contents: [{ parts: [{ text: 'Reply with exactly: OK' }] }],
      generationConfig: { temperature: 0, maxOutputTokens: 8 }
    };

    function callModel(modelId) {
      var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(modelId) + ':generateContent?key=' + apiKey;
      return UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(body),
        muteHttpExceptions: true
      });
    }

    var tried = [];
    var candidates = [primaryModel, 'gemini-2.5-flash', 'gemini-1.5-flash', 'gemini-2.0-flash'];
    var seen = {};
    var lastCode = 0;
    var lastText = '';
    for (var i = 0; i < candidates.length; i++) {
      var model = String(candidates[i] || '').trim();
      if (!model || seen[model]) continue;
      seen[model] = true;
      tried.push(model);

      var resp = callModel(model);
      var code = Number(resp.getResponseCode() || 0);
      var text = String(resp.getContentText() || '');
      if (code === 200) {
        return { success: true, model: model, httpCode: 200, triedModels: tried };
      }
      lastCode = code;
      lastText = text;
    }

    var errObj = {};
    try { errObj = JSON.parse(lastText); } catch (parseErr) {}
    var msg = (errObj && errObj.error && errObj.error.message) ? errObj.error.message : String(lastText || '').slice(0, 280);
    return {
      success: false,
      error: 'Gemini API error (' + lastCode + '): ' + msg,
      model: primaryModel,
      httpCode: lastCode,
      triedModels: tried
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function saveToolsRunwayAuthorization(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var sh = _pilotRunwayChecksSheet_();
    var headers = _toolsSheetHeaderRow_(sh);
    var data = sh.getDataRange().getValues();
    var rowNumber = Number(body.rowNumber || 0);
    var pilotName = String(body.pilotName || '').trim();
    var pilotEmail = String(body.pilotEmail || '').trim().toLowerCase();
    var icao = String(body.icao || '').trim().toUpperCase();
    var runwayIdent = String(body.runwayIdent || '').trim().toUpperCase();

    if (!pilotName) return { success: false, error: 'Pilot name is required' };
    if (!icao) return { success: false, error: 'ICAO is required' };

    var status = String(body.status || 'ACTIVE').trim().toUpperCase();
    var source = String(body.source || '').trim();
    var notes = String(body.notes || '').trim();
    var approvedBy = String(body.approvedBy || '').trim();
    var dateChecked = safeDateStr(body.dateChecked || new Date());
    var expiryDate = safeDateStr(body.expiryDate || '');
    var authScope = String(body.authScope || 'RUNWAY').trim().toUpperCase() || 'RUNWAY';

    if (rowNumber < 2) {
      for (var i = 1; i < data.length; i++) {
        var rec = _pilotRunwayCheckRecordFromRow_(headers, data[i], i + 1);
        if (String(rec.pilotName || '').trim().toUpperCase() !== String(pilotName || '').trim().toUpperCase()) continue;
        if (String(rec.icao || '').trim().toUpperCase() !== icao) continue;
        if (String(rec.runwayIdent || '').trim().toUpperCase() !== runwayIdent) continue;
        rowNumber = i + 1;
        break;
      }
    }

    var now = safeDateStr(new Date());
    var dataMap = {
      CHECK_ID: String(body.checkId || ('CHK_' + new Date().getTime() + '_' + icao + '_' + runwayIdent)).trim(),
      PILOT_NAME: pilotName,
      PILOT_EMAIL: pilotEmail,
      STAFF_ID: String(body.staffId || '').trim(),
      ICAO: icao,
      RUNWAY_IDENT: runwayIdent,
      AUTH_SCOPE: authScope,
      STATUS: status,
      DATE_CHECKED: dateChecked,
      EXPIRY_DATE: expiryDate,
      APPROVED_BY: approvedBy,
      SOURCE: source,
      NOTES: notes,
      UPDATED_AT: now
    };

    if (rowNumber >= 2) {
      var current = sh.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
      var merged = headers.map(function(header, idx) {
        var key = _toolsNormHeader_(header);
        if (key === 'CREATED_AT') return current[idx] || now;
        return Object.prototype.hasOwnProperty.call(dataMap, key) ? dataMap[key] : current[idx];
      });
      sh.getRange(rowNumber, 1, 1, merged.length).setValues([merged]);
      return { success: true, action: 'updated', rowNumber: rowNumber };
    }

    dataMap.CREATED_AT = now;
    var row = headers.map(function(header) {
      var key = _toolsNormHeader_(header);
      return Object.prototype.hasOwnProperty.call(dataMap, key) ? dataMap[key] : '';
    });
    sh.appendRow(row);
    return { success: true, action: 'created', rowNumber: sh.getLastRow() };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}
