function getPilotDutyTimeline(pilotName, centerDate) {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const oneDay = 24*60*60*1000;
if (!(centerDate instanceof Date) || isNaN(centerDate)) centerDate = new Date();
centerDate = new Date(centerDate.getTime());
centerDate.setHours(0,0,0,0);

const _toLocalDate_ = function(value) {
  const ymd = _dutyYmd_(value);
  if (!ymd) return null;
  const parts = String(ymd).split('-').map(function(n) { return Number(n); });
  if (parts.length !== 3 || !isFinite(parts[0]) || !isFinite(parts[1]) || !isFinite(parts[2])) return null;
  const out = new Date(parts[0], parts[1] - 1, parts[2]);
  if (isNaN(out.getTime())) return null;
  out.setHours(0,0,0,0);
  return out;
};

const startDate = new Date(centerDate.getTime() - (7 * oneDay));
const endDate = new Date(centerDate.getTime() + (7 * oneDay));
let events = [];
const pilotTarget = String(pilotName || '').trim().toUpperCase();
const logSheet = ss.getSheetByName(APP_SHEETS.DUTY_LOG);
if (logSheet) {
const logData = logSheet.getDataRange().getValues();
const actualByDay = {};
for (let i = 1; i < logData.length; i++) {
 const rowPilot = String(logData[i][DUTY_LOG_COL.PILOT] || '').trim().toUpperCase();
 if (rowPilot !== pilotTarget) continue;
 const dayDate = _toLocalDate_(logData[i][DUTY_LOG_COL.DATE]);
 if (!dayDate) continue;
 if (dayDate < startDate || dayDate > endDate) continue;

 const titleRaw = String(logData[i][DUTY_LOG_COL.TITLE] || '').trim().toUpperCase();
 let payload = {};
 try {
   const jsonRaw = String(logData[i][5] || '').trim();
   payload = jsonRaw ? JSON.parse(jsonRaw) : {};
 } catch (e) {
   payload = {};
 }
 const statusRaw = String((payload && payload.status) || titleRaw || '').trim().toUpperCase();
 const isClosed = statusRaw.indexOf('CLOSED') >= 0 || titleRaw.indexOf('DUTY_CLOSED') >= 0;
 if (!isClosed) continue;

 const ymd = Utilities.formatDate(dayDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
 const startH = _dutyParseHhmmToHours_((payload && payload.startTime) || '');
 const endH = _dutyParseHhmmToHours_((payload && payload.endTime) || '');
 let dutyHrs = 0;
 if (startH > 0 && endH > 0) {
   dutyHrs = endH - startH;
   if (dutyHrs < 0) dutyHrs += 24;
 }
 if (!isFinite(dutyHrs) || dutyHrs < 0) dutyHrs = 0;

 if (!actualByDay[ymd]) {
   actualByDay[ymd] = {
     date: ymd,
     type: 'LOGGED',
     title: 'DUTY_CLOSED',
     desc: 'Actual duty report',
     flightHrs: 0,
     dutyHrs: 0
   };
 }
 actualByDay[ymd].dutyHrs += dutyHrs;
}

 Object.keys(actualByDay).forEach(function(ymd) {
   const daily = _dutyDailyFlightSummary_(ss, pilotName, ymd);
   actualByDay[ymd].flightHrs = Number((daily && daily.totalFlightHours) || 0);
   events.push(actualByDay[ymd]);
 });
}
const dispSheet = ss.getSheetByName(APP_SHEETS.DISPATCH);
if (dispSheet) {
const dispData = dispSheet.getDataRange().getValues();
const tracker = {};
for (let i = 1; i < dispData.length; i++) {
 const rowPilot = String(dispData[i][DISPATCH_COL.PILOT] || '').trim().toUpperCase();
 if (rowPilot === pilotTarget) {
   const d = _toLocalDate_(dispData[i][DISPATCH_COL.DATE]);
   if (!d) continue;
   if (d >= startDate && d <= endDate) {
     const mId = dispData[i][DISPATCH_COL.MISSION_ID];
     if(!tracker[mId]) {
       tracker[mId] = {
         date: Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'), // Local date, not UTC
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
Object.values(tracker).forEach(e => events.push(e));
}
return events.sort((a, b) => a.date.localeCompare(b.date));
}

function getPilotDutyPromptSnapshot(pilotName, dateYmd) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pilot = _dutyNormalizePilotName_(pilotName);
  const ymd = String(dateYmd || _dutyYmd_(new Date())).trim();
  if (!pilot || !ymd) {
    return { success: false, error: 'pilotName and dateYmd are required' };
  }

  const flightSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  const flightData = flightSheet ? flightSheet.getDataRange().getValues() : null;
  const transSheet = ss.getSheetByName(APP_SHEETS.TRANSACTIONS);
  const transData = transSheet ? transSheet.getDataRange().getValues() : null;

  const entry = _dutyReadLatestEntry_(ss, pilot, ymd);
  const autofill = _dutyDailyFlightSummary_(ss, pilot, ymd, { flightData: flightData, transData: transData });

  const closed = !!entry.endTime && String(entry.status || '').indexOf('CLOSE') >= 0;
  const shouldMorningPrompt = !entry.startTime && !closed;
  const shouldEveningPrompt = !!entry.startTime && !entry.endTime;

  return {
    success: true,
    pilotName: pilot,
    dateYmd: ymd,
    entry: entry,
    autofill: autofill,
    dutyConfig: _toolsDutyReadConfigMap_(),
    shouldMorningPrompt: shouldMorningPrompt,
    shouldEveningPrompt: shouldEveningPrompt
  };
}

function savePilotDutyReport(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(APP_SHEETS.DUTY_LOG);
  if (!sh) throw new Error("Sheet 'LOG_Duty' not found.");

  const p = (payload && typeof payload === 'object') ? payload : {};
  const pilot = _dutyNormalizePilotName_(p.pilotName);
  const ymd = String(p.dateYmd || _dutyYmd_(new Date())).trim();
  const status = String(p.status || '').trim().toUpperCase() || 'DUTY_UPDATE';
  const startTime = String(p.startTime || '').trim();
  const endTime = String(p.endTime || '').trim();
  const notes = String(p.notes || '').trim();
  const description = String(p.description || notes || '').trim();
  const department = String(p.department || _dutyDepartmentFromDescription_(description) || '').trim();
  const flightHours = (p.flightHours == null || p.flightHours === '') ? '' : Number(p.flightHours);
  const ignoredMorning = false;
  const ignoredEvening = false;
  const autoSummary = (p.autoSummary && typeof p.autoSummary === 'object') ? p.autoSummary : {};

  if (!pilot) throw new Error('savePilotDutyReport: pilotName is required');
  if (!ymd) throw new Error('savePilotDutyReport: dateYmd is required');
  if (!department) throw new Error('Selecione um departamento antes de salvar o dia.');

  const summaryText = String(
    p.summaryText ||
    notes ||
    autoSummary.summaryLine ||
    description ||
    ''
  ).trim();

  const jsonPayload = {
    status: status,
    startTime: startTime,
    endTime: endTime,
    department: department,
    flightHours: (flightHours === '' || !isFinite(flightHours)) ? '' : Number(flightHours.toFixed(2)),
    notes: notes,
    description: description,
    summaryText: summaryText,
    autoSummary: autoSummary,
    ignoredMorning: ignoredMorning,
    ignoredEvening: ignoredEvening,
    updatedAt: new Date().toISOString()
  };

  const row = [
    ymd,
    pilot,
    status,
    '',
    summaryText,
    JSON.stringify(jsonPayload),
    description || summaryText
  ];
  sh.appendRow(row);

  return {
    success: true,
    rowNumber: sh.getLastRow(),
    pilotName: pilot,
    dateYmd: ymd,
    status: status,
    startTime: startTime,
    endTime: endTime,
    department: department,
    flightHours: jsonPayload.flightHours,
    notes: notes,
    description: description,
    summaryText: summaryText,
    updatedAt: jsonPayload.updatedAt
  };
}

function getToolsDutyAdminBootstrap() {
  try {
    var ctx = _toolsDutyContext_();
    var cfg = _toolsDutyReadConfigMap_();
    return {
      success: true,
      appVersion: DUTY_APP_BUILD_VERSION,
      currentUser: {
        email: ctx.userEmail,
        staffName: ctx.me ? ctx.me.staffName : '',
        staffId: ctx.me ? ctx.me.staffId : ''
      },
      canManageAll: !!ctx.canManageAll,
      canViewAllLogs: _toolsDutyCanViewAllLogs_(ctx),
      canManageConfig: !!ctx.canManageConfig,
      staff: (ctx.staffRows || []).filter(function(s) { return !!s.active; }).map(function(s) {
        return {
          email: String(s.email || '').trim().toLowerCase(),
          staffName: String(s.staffName || '').trim(),
          staffId: String(s.staffId || '').trim()
        };
      }),
      dutyConfig: cfg
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function saveToolsDutyConfiguration(payload) {
  try {
    var ctx = _toolsDutyContext_();
    if (!ctx.canManageConfig) {
      return { success: false, error: 'Permissão insuficiente para alterar a configuração mestre.' };
    }

    var input = (payload && typeof payload === 'object' && payload.config && typeof payload.config === 'object')
      ? payload.config
      : {};
    var defaults = _toolsDutyConfigDefaults_();
    var descMap = _toolsDutyConfigDescriptions_();
    var entries = Object.keys(defaults).map(function(key) {
      return {
        key: key,
        value: Object.prototype.hasOwnProperty.call(input, key) ? input[key] : defaults[key],
        description: descMap[key] || ''
      };
    });

    var writeRes = _toolsDutySaveConfigMap_(entries, ctx.userEmail || 'tools');
    var cfg = _toolsDutyReadConfigMap_();
    return {
      success: true,
      updated: Number(writeRes.updated || 0),
      created: Number(writeRes.created || 0),
      dutyConfig: cfg
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getToolsDutyLogs(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var ctx = _toolsDutyContext_(body.preferredDutyEmail);
    var fromYmd = String(body.fromYmd || '').trim();
    var toYmd = String(body.toYmd || '').trim();
    var pilotFilter = String(body.pilotFilter || '').trim().toUpperCase();
    var canViewAllLogs = _toolsDutyCanViewAllLogs_(ctx);

    var now = new Date();
    if (!toYmd) toYmd = _dutyYmd_(now);
    if (!fromYmd) {
      var from = new Date(now.getTime() - (1000 * 60 * 60 * 24 * 30));
      fromYmd = _dutyYmd_(from);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.DUTY_LOG, 'getToolsDutyLogs');
    var data = sh.getDataRange().getValues();
    if (!data || data.length < 2) {
      return { success: true, canManageAll: !!ctx.canManageAll, rows: [] };
    }

    var flightSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
    var flightData = flightSheet ? flightSheet.getDataRange().getValues() : null;
    var transSheet = ss.getSheetByName(APP_SHEETS.TRANSACTIONS);
    var transData = transSheet ? transSheet.getDataRange().getValues() : null;

    var out = [];
    var dutyDayByPilotDate = {};
    var pilotNamesByUpper = {};
    var dailyFlightCache = {};
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var pilot = String(row[DUTY_LOG_COL.PILOT] || '').trim();
      if (!canViewAllLogs && !_toolsDutyCanAccessPilot_(ctx, pilot)) continue;
      // If pilot was stored as an email, resolve to canonical staff name so flight matching works.
      if (pilot.indexOf('@') >= 0) {
        var resolvedForFlight = _toolsFindStaffByEmailOrId_(ctx.staffRows, pilot, '');
        if (resolvedForFlight && resolvedForFlight.staffName) pilot = resolvedForFlight.staffName;
      }

      var dateYmd = _dutyYmd_(row[DUTY_LOG_COL.DATE]);
      if (!dateYmd) continue;
      if (fromYmd && dateYmd < fromYmd) continue;
      if (toYmd && dateYmd > toYmd) continue;
      if (pilotFilter && !_dutyPilotLikelySame_(pilot, pilotFilter)) continue;

      var pilotUpper = _dutyPilotNorm_(pilot);
      if (!pilotUpper) continue;
      var dayKey = pilotUpper + '|' + dateYmd;
      dutyDayByPilotDate[dayKey] = true;
      if (!pilotNamesByUpper[pilotUpper]) pilotNamesByUpper[pilotUpper] = pilot;

      var status = String(row[DUTY_LOG_COL.TITLE] || '').trim();
      var summaryText = String(row[DUTY_LOG_COL.DESC_FALLBACK] || '').trim();
      var descPrimary = String(row[DUTY_LOG_COL.DESC_PRIMARY] || '').trim();
      var jsonRaw = String(row[5] || '').trim();
      var parsed = {};
      try { parsed = jsonRaw ? JSON.parse(jsonRaw) : {}; } catch (e0) { parsed = {}; }
      var cacheKey = String(pilot || '').trim().toUpperCase() + '|' + dateYmd;
      if (!dailyFlightCache[cacheKey]) {
        dailyFlightCache[cacheKey] = _dutyDailyFlightSummary_(ss, pilot, dateYmd, { flightData: flightData, transData: transData });
      }
      var dailyFlight = dailyFlightCache[cacheKey] || {};
      var parsedFlightHours = String(parsed.flightHours == null ? '' : parsed.flightHours).trim();
      var derivedFlightHours = Number(dailyFlight.totalFlightHours || 0);
      var totalDutyHours = _dutyTotalDutyHours_(parsed.startTime || row[3] || '', parsed.endTime || row[4] || '');
      var resolvedFlightHours = parsedFlightHours;
      if (resolvedFlightHours === '' && isFinite(derivedFlightHours) && derivedFlightHours > 0) {
        resolvedFlightHours = String(Number(derivedFlightHours.toFixed(2)));
      }

      out.push({
        rowNumber: i + 1,
        dateYmd: dateYmd,
        pilotName: pilot,
        status: String(parsed.status || status || '').trim(),
        startTime: String(parsed.startTime || '').trim(),
        endTime: String(parsed.endTime || '').trim(),
        flightHours: resolvedFlightHours,
        storedFlightHours: parsedFlightHours,
        derivedFlightHours: isFinite(derivedFlightHours) ? Number(derivedFlightHours.toFixed(2)) : 0,
        totalDutyHours: isFinite(totalDutyHours) ? Number(totalDutyHours.toFixed(2)) : 0,
        flightHoursSource: parsedFlightHours !== '' ? 'LOG_Duty' : (isFinite(derivedFlightHours) && derivedFlightHours > 0 ? 'LOG_Flights' : ''),
        notes: String(parsed.notes || parsed.description || descPrimary || summaryText || '').trim(),
        description: String(parsed.description || descPrimary || '').trim(),
        summaryText: String(parsed.summaryText || summaryText || '').trim(),
        updatedAt: String(parsed.updatedAt || '').trim(),
        flightSummary: String(dailyFlight.summaryLine || '').trim(),
        flights: Array.isArray(dailyFlight.flights) ? dailyFlight.flights : []
      });
    }

    // Add synthetic rows for days with flights but no duty log entry yet.
    var flightSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
    if (flightSheet) {
      var flightData = flightSheet.getDataRange().getValues();
      if (flightData && flightData.length > 1) {
        for (var j = 1; j < flightData.length; j++) {
          var flightRow = flightData[j];
          var flightPilot = String(flightRow[LOG_FLIGHT_COL.PILOT] || '').trim();
          var flightDateYmd = _dutyYmd_(flightRow[LOG_FLIGHT_COL.DATE]);
          if (!flightDateYmd) continue;
          if (fromYmd && flightDateYmd < fromYmd) continue;
          if (toYmd && flightDateYmd > toYmd) continue;

          var matchingCrewName = '';
          var matchingCrewUpper = '';
          
          // Check if pilot matches
          if (flightPilot && (canViewAllLogs || _toolsDutyCanAccessPilot_(ctx, flightPilot))) {
            matchingCrewName = flightPilot;
            matchingCrewUpper = _dutyPilotNorm_(flightPilot);
          }
          
          // Also check co-pilot/student names in Actual_Load_JSON
          if (!matchingCrewName) {
            try {
              var loadJsonRaw = String(flightRow[LOG_FLIGHT_COL.ACTUAL_LOAD_JSON] || '').trim();
              if (loadJsonRaw) {
                var loadJson = JSON.parse(loadJsonRaw);
                if (loadJson && loadJson.seats && typeof loadJson.seats === 'object') {
                  for (var seatKey in loadJson.seats) {
                    var seat = loadJson.seats[seatKey];
                    if (seat && seat.passenger && seat.passenger.name) {
                      var passengerName = String(seat.passenger.name || '').trim();
                      if (passengerName && (canViewAllLogs || _toolsDutyCanAccessPilot_(ctx, passengerName))) {
                        matchingCrewName = passengerName;
                        matchingCrewUpper = _dutyPilotNorm_(passengerName);
                        break;
                      }
                    }
                  }
                }
              }
            } catch (e) {}
          }
          
          if (!matchingCrewName || !matchingCrewUpper) continue;
          if (pilotFilter && !_dutyPilotLikelySame_(matchingCrewName, pilotFilter)) continue;

          var flightDayKey = matchingCrewUpper + '|' + flightDateYmd;
          if (dutyDayByPilotDate[flightDayKey]) continue;

          if (!pilotNamesByUpper[matchingCrewUpper]) pilotNamesByUpper[matchingCrewUpper] = matchingCrewName;
          if (!dailyFlightCache[flightDayKey]) {
            dailyFlightCache[flightDayKey] = _dutyDailyFlightSummary_(ss, pilotNamesByUpper[matchingCrewUpper], flightDateYmd, { flightData: flightData, transData: transData });
          }
          var dailyFlightSynthetic = dailyFlightCache[flightDayKey] || {};
          var syntheticFlights = Array.isArray(dailyFlightSynthetic.flights) ? dailyFlightSynthetic.flights : [];
          var derivedSyntheticHrs = Number(dailyFlightSynthetic.totalFlightHours || 0);
          if (!syntheticFlights.length) continue;

          dutyDayByPilotDate[flightDayKey] = true;
          out.push({
            rowNumber: 0,
            dateYmd: flightDateYmd,
            pilotName: pilotNamesByUpper[matchingCrewUpper] || matchingCrewName,
            status: 'FLIGHT_ONLY',
            startTime: '',
            endTime: '',
            flightHours: (isFinite(derivedSyntheticHrs) && derivedSyntheticHrs > 0) ? String(Number(derivedSyntheticHrs.toFixed(2))) : '',
            storedFlightHours: '',
            derivedFlightHours: (isFinite(derivedSyntheticHrs) && derivedSyntheticHrs > 0) ? Number(derivedSyntheticHrs.toFixed(2)) : 0,
            totalDutyHours: '',
            flightHoursSource: 'LOG_Flights',
            notes: '',
            description: 'Sem jornada registrada. Voo(s) encontrado(s) no dia.',
            summaryText: String(dailyFlightSynthetic.summaryLine || '').trim(),
            updatedAt: '',
            flightSummary: String(dailyFlightSynthetic.summaryLine || '').trim(),
            flights: syntheticFlights
          });
        }
      }
    }

    out.sort(function(a, b) {
      var d = String(b.dateYmd || '').localeCompare(String(a.dateYmd || ''));
      if (d !== 0) return d;
      return Number(b.rowNumber || 0) - Number(a.rowNumber || 0);
    });

    return {
      success: true,
      canManageAll: !!ctx.canManageAll,
      rows: out,
      applied: { fromYmd: fromYmd, toYmd: toYmd, pilotFilter: pilotFilter }
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function updateToolsDutyLogEntry(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var rowNumber = Number(body.rowNumber || 0);
    var pilotName = String(body.pilotName || '').trim();
    var dateYmd = String(body.dateYmd || '').trim();
    if (!(rowNumber >= 2)) {
      if (!pilotName) return { success: false, error: 'pilotName é obrigatório para criar um log.' };
      if (!dateYmd) return { success: false, error: 'dateYmd é obrigatório para criar um log.' };
      var createDepartment = String(body.department || _dutyDepartmentFromDescription_(body.description) || '').trim();
      if (!createDepartment) return { success: false, error: 'Selecione um departamento antes de salvar o dia.' };
      return savePilotDutyReport({
        pilotName: pilotName,
        dateYmd: _dutyYmd_(body.dateYmd) || dateYmd,
        status: String(body.status || '').trim(),
        startTime: String(body.startTime || '').trim(),
        endTime: String(body.endTime || '').trim(),
        department: createDepartment,
        flightHours: body.flightHours,
        description: String(body.description || '').trim(),
        summaryText: String(body.summaryText || '').trim()
      });
    }

    var ctx = _toolsDutyContext_(body.preferredDutyEmail);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.DUTY_LOG, 'updateToolsDutyLogEntry');
    if (rowNumber > sh.getLastRow()) return { success: false, error: 'Registro não encontrado' };

    var canViewAllLogs = _toolsDutyCanViewAllLogs_(ctx);
    var width = Math.max(7, sh.getLastColumn());
    var current = sh.getRange(rowNumber, 1, 1, width).getValues()[0];
    var pilot = String(current[DUTY_LOG_COL.PILOT] || '').trim();
    if (!canViewAllLogs && !_toolsDutyCanAccessPilot_(ctx, pilot)) {
      return { success: false, error: 'Sem permissão para editar este registro.' };
    }

    var existingJsonRaw = String(current[5] || '').trim();
    var existing = {};
    try { existing = existingJsonRaw ? JSON.parse(existingJsonRaw) : {}; } catch (e0) { existing = {}; }

    var status = String(body.status || existing.status || current[DUTY_LOG_COL.TITLE] || 'DUTY_UPDATE').trim().toUpperCase();
    var startTime = String(body.startTime == null ? (existing.startTime || '') : body.startTime).trim();
    var endTime = String(body.endTime == null ? (existing.endTime || '') : body.endTime).trim();
    var notes = String(body.notes == null ? (existing.notes || existing.description || existing.summaryText || current[DUTY_LOG_COL.DESC_PRIMARY] || current[DUTY_LOG_COL.DESC_FALLBACK] || '') : body.notes).trim();
    var description = String(body.description == null ? (notes || existing.description || current[DUTY_LOG_COL.DESC_PRIMARY] || '') : body.description).trim();
    var department = String(body.department || _dutyDepartmentFromDescription_(description) || _dutyDepartmentFromDescription_(existing.description || current[DUTY_LOG_COL.DESC_PRIMARY] || '') || '').trim();
    if (!department) return { success: false, error: 'Selecione um departamento antes de salvar o dia.' };
    var summaryText = String(body.summaryText == null ? (existing.summaryText || notes || current[DUTY_LOG_COL.DESC_FALLBACK] || description) : body.summaryText).trim();
    var fhRaw = (Object.prototype.hasOwnProperty.call(body, 'flightHours') ? body.flightHours : existing.flightHours);
    var flightHours = (fhRaw === '' || fhRaw == null) ? '' : Number(fhRaw);

    var merged = {
      status: status,
      startTime: startTime,
      endTime: endTime,
      flightHours: (flightHours === '' || !isFinite(flightHours)) ? '' : Number(flightHours.toFixed(2)),
      department: department,
      notes: notes,
      description: description,
      summaryText: summaryText,
      autoSummary: (existing.autoSummary && typeof existing.autoSummary === 'object') ? existing.autoSummary : {},
      ignoredMorning: !!existing.ignoredMorning,
      ignoredEvening: !!existing.ignoredEvening,
      updatedAt: new Date().toISOString()
    };

    current[DUTY_LOG_COL.TITLE] = status;
    if (current.length > DUTY_LOG_COL.DESC_FALLBACK) current[DUTY_LOG_COL.DESC_FALLBACK] = summaryText;
    if (current.length > 5) current[5] = JSON.stringify(merged);
    if (current.length > DUTY_LOG_COL.DESC_PRIMARY) current[DUTY_LOG_COL.DESC_PRIMARY] = description;

    sh.getRange(rowNumber, 1, 1, width).setValues([current]);

    return {
      success: true,
      rowNumber: rowNumber,
      row: {
        rowNumber: rowNumber,
        dateYmd: _dutyYmd_(current[DUTY_LOG_COL.DATE]),
        pilotName: pilot,
        status: status,
        startTime: startTime,
        endTime: endTime,
        department: department,
        flightHours: merged.flightHours,
        description: description,
        summaryText: summaryText,
        updatedAt: merged.updatedAt
      }
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getDutyAppBootstrap(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var preferredEmail = String(body.preferredDutyEmail || '').trim().toLowerCase();
    var sessionEmail = _toolsCurrentUserEmail_('');
    var savedPreferredEmail = _dutyReadSavedPreferredEmail_(body.pilotName);
    var access = _dutyAppResolvePilotAccess_(body.pilotName, preferredEmail || savedPreferredEmail);
    if (!access.success) return access;

    var todayYmd = _dutyAppTodayBsb_();
    var dateYmd = String(body.dateYmd || todayYmd).trim() || todayYmd;
    var toYmd = String(body.toYmd || todayYmd).trim() || todayYmd;
    var fromYmd = String(body.fromYmd || '').trim();
    if (!fromYmd) {
      fromYmd = dateYmd;
    }
    if (!toYmd) {
      toYmd = dateYmd;
    }

    var snapshot = getPilotDutyPromptSnapshot(access.pilotName, dateYmd);
    if (!snapshot || snapshot.success !== true) {
      return {
        success: false,
        error: (snapshot && snapshot.error) ? snapshot.error : 'Falha ao carregar snapshot de duty.'
      };
    }

    var logsRes = getToolsDutyLogs({
      fromYmd: fromYmd,
      toYmd: toYmd,
      pilotFilter: access.pilotName,
      preferredDutyEmail: preferredEmail
    });
    if (!logsRes || logsRes.success !== true) {
      return {
        success: false,
        error: (logsRes && logsRes.error) ? logsRes.error : 'Falha ao carregar logs de duty.'
      };
    }

    return {
      success: true,
      appVersion: DUTY_APP_BUILD_VERSION,
      currentUser: {
        email: String(access.ctx.userEmail || '').trim().toLowerCase(),
        sessionEmail: sessionEmail,
        staffName: String((access.ctx.me && access.ctx.me.staffName) || access.pilotName || '').trim(),
        staffId: String((access.ctx.me && access.ctx.me.staffId) || '').trim(),
        dbPilotName: String(_dutyLookupPilotNameByEmail_(String(access.ctx.userEmail || '').trim().toLowerCase()) || '').trim()
      },
      preferredDutyEmail: savedPreferredEmail,
      canManageAll: !!access.ctx.canManageAll,
      selectedPilotName: access.pilotName,
      staff: (access.ctx.staffRows || []).filter(function(staff) {
        return !!staff.active;
      }).map(function(staff) {
        return {
          email: String(staff.email || '').trim().toLowerCase(),
          staffName: String(staff.staffName || '').trim(),
          staffId: String(staff.staffId || '').trim()
        };
      }),
      dateYmd: dateYmd,
      fromYmd: fromYmd,
      toYmd: toYmd,
      snapshot: snapshot,
      logs: logsRes.rows || [],
      dutyConfig: _toolsDutyReadConfigMap_()
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function saveDutyPreferredEmail(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var pilotName = String(body.pilotName || '').trim();
    if (!pilotName) return { success: false, error: 'pilotName is required.' };
    var access = _dutyAppResolvePilotAccess_(pilotName, _toolsCurrentUserEmail_(''));
    if (!access.success) return access;
    var savedEmail = _dutySavePreferredEmail_(access.pilotName, body.preferredDutyEmail || '');
    return { success: true, pilotName: access.pilotName, preferredDutyEmail: savedEmail };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function saveDutyAppEntry(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var access = _dutyAppResolvePilotAccess_(body.pilotName, body.preferredDutyEmail);
    if (!access.success) return access;

    var safePayload = {
      pilotName: access.pilotName,
      dateYmd: String(body.dateYmd || _dutyAppTodayBsb_()).trim(),
      status: String(body.status || '').trim(),
      startTime: String(body.startTime || '').trim(),
      endTime: String(body.endTime || '').trim(),
      description: String(body.description || '').trim(),
      summaryText: String(body.summaryText || '').trim(),
      ignoredMorning: false,
      ignoredEvening: false,
      autoSummary: (body.autoSummary && typeof body.autoSummary === 'object') ? body.autoSummary : {}
    };

    if (body.flightHours === '' || body.flightHours == null) {
      safePayload.flightHours = '';
    } else {
      var fh = Number(body.flightHours);
      safePayload.flightHours = isFinite(fh) ? fh : '';
    }

    return savePilotDutyReport(safePayload);
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function updateDutyAppLog(payload) {
  try {
    return updateToolsDutyLogEntry(payload);
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function deleteDutyAppEntry(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var rowNumber = Number(body.rowNumber || 0);
    if (!(rowNumber >= 2)) return { success: false, error: 'rowNumber is required' };

    var ctx = _toolsDutyContext_(body.preferredDutyEmail);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.DUTY_LOG, 'deleteDutyAppEntry');
    if (rowNumber > sh.getLastRow()) return { success: false, error: 'Registro não encontrado' };

    var canViewAllLogs = _toolsDutyCanViewAllLogs_(ctx);
    var width = Math.max(7, sh.getLastColumn());
    var current = sh.getRange(rowNumber, 1, 1, width).getValues()[0];
    var pilot = String(current[DUTY_LOG_COL.PILOT] || '').trim();
    if (!canViewAllLogs && !_toolsDutyCanAccessPilot_(ctx, pilot)) {
      return { success: false, error: 'Sem permissão para excluir este registro.' };
    }

    sh.deleteRow(rowNumber);
    return {
      success: true,
      rowNumber: rowNumber,
      pilotName: pilot,
      dateYmd: _dutyYmd_(current[DUTY_LOG_COL.DATE])
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function savePilotLessonEvaluation(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var missionId = String(body.missionId || '').trim();
    var trainingCode = String(body.trainingCode || '').trim();
    if (!missionId || !trainingCode) {
      throw new Error('savePilotLessonEvaluation: missionId and trainingCode are required');
    }

    var sheet = _lessonEvalSheet_();
    var data = sheet.getDataRange().getValues();
    var headers = (data && data.length ? data[0] : _lessonEvalHeaders_()).map(function(h) { return String(h || '').trim(); });
    var map = _lessonEvalHeaderMap_(headers);

    var now = new Date();
    var actor = _lessonEvalActor_();
    var hasScores = Object.prototype.hasOwnProperty.call(body, 'scores');
    var hasNotes = Object.prototype.hasOwnProperty.call(body, 'notes');
    var scores = (body.scores && typeof body.scores === 'object') ? body.scores : {};
    var notes = (body.notes && typeof body.notes === 'object') ? body.notes : {};

    var targetRow = 0;
    var missionIdx = map.MISSION_ID;
    var codeIdx = map.TRAINING_CODE;
    if (missionIdx >= 0 && codeIdx >= 0) {
      for (var r = data.length - 1; r >= 1; r--) {
        var rowMission = String(data[r][missionIdx] || '').trim();
        var rowCode = String(data[r][codeIdx] || '').trim();
        if (rowMission === missionId && rowCode === trainingCode) {
          targetRow = r + 1;
          break;
        }
      }
    }

    var row = new Array(headers.length).fill('');
    if (targetRow > 0) {
      row = sheet.getRange(targetRow, 1, 1, headers.length).getValues()[0];
    }

    if (map.EVAL_ID >= 0 && !String(row[map.EVAL_ID] || '').trim()) row[map.EVAL_ID] = Utilities.getUuid();
    if (map.MISSION_ID >= 0) row[map.MISSION_ID] = missionId;
    if (map.TRAINING_CODE >= 0) row[map.TRAINING_CODE] = trainingCode;
    if (map.PILOT >= 0) row[map.PILOT] = String(body.pilot || '').trim();
    if (map.AIRCRAFT >= 0) row[map.AIRCRAFT] = String(body.aircraft || '').trim();
    if (map.MISSION_DATE >= 0) row[map.MISSION_DATE] = String(body.missionDate || '').trim();
    if (map.SOURCE_TAB >= 0) row[map.SOURCE_TAB] = String(body.sourceTab || 'PILOT_LESSON_MODAL').trim();
    if (hasScores && map.SCORES_JSON >= 0) row[map.SCORES_JSON] = JSON.stringify(scores || {});
    if (hasNotes && map.NOTES_JSON >= 0) row[map.NOTES_JSON] = JSON.stringify(notes || {});
    var newStatus = String(body.status || '').trim().toUpperCase();
    if (newStatus && map.STATUS >= 0) row[map.STATUS] = newStatus;
    var newResult = String(body.result || '').trim().toUpperCase();
    if (newResult && map.RESULT >= 0) row[map.RESULT] = newResult;
    if (newResult && map.RESULT_AT >= 0) row[map.RESULT_AT] = body.resultAt || now;
    if (newStatus === 'SUBMITTED' && map.SUBMITTED_AT >= 0 && !row[map.SUBMITTED_AT]) row[map.SUBMITTED_AT] = now;
    if (map.UPDATED_AT >= 0) row[map.UPDATED_AT] = now;
    if (map.UPDATED_BY >= 0) row[map.UPDATED_BY] = actor;
    if (map.CREATED_AT >= 0 && !row[map.CREATED_AT]) row[map.CREATED_AT] = now;
    if (map.CREATED_BY >= 0 && !String(row[map.CREATED_BY] || '').trim()) row[map.CREATED_BY] = actor;

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, headers.length).setValues([row]);
    } else {
      sheet.appendRow(row);
      targetRow = sheet.getLastRow();
    }

    return {
      success: true,
      missionId: missionId,
      trainingCode: trainingCode,
      rowNumber: targetRow,
      updatedAt: now.toISOString()
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getPilotLessonEvaluation(missionId, trainingCode) {
  try {
    var targetMission = String(missionId || '').trim();
    var targetCode = String(trainingCode || '').trim();
    if (!targetMission || !targetCode) return { success: true, found: false, payload: null };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = APP_SHEETS.FLIGHT_LESSON_EVALS || 'DB_Flight_Lesson_Evals';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: true, found: false, payload: null };

    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return { success: true, found: false, payload: null };
    var headers = data[0].map(function(h) { return String(h || '').trim(); });
    var map = _lessonEvalHeaderMap_(headers);
    if (map.MISSION_ID < 0 || map.TRAINING_CODE < 0) return { success: true, found: false, payload: null };

    var hit = null;
    for (var r = data.length - 1; r >= 1; r--) {
      var row = data[r];
      if (String(row[map.MISSION_ID] || '').trim() !== targetMission) continue;
      if (String(row[map.TRAINING_CODE] || '').trim() !== targetCode) continue;
      hit = row;
      break;
    }
    if (!hit) return { success: true, found: false, payload: null };

    return {
      success: true,
      found: true,
      payload: {
        missionId: targetMission,
        trainingCode: targetCode,
        scores: _lessonEvalJsonSafe_(map.SCORES_JSON >= 0 ? hit[map.SCORES_JSON] : '', {}),
        notes: _lessonEvalJsonSafe_(map.NOTES_JSON >= 0 ? hit[map.NOTES_JSON] : '', {}),
        pilot: map.PILOT >= 0 ? String(hit[map.PILOT] || '') : '',
        aircraft: map.AIRCRAFT >= 0 ? String(hit[map.AIRCRAFT] || '') : '',
        missionDate: map.MISSION_DATE >= 0 ? String(hit[map.MISSION_DATE] || '') : '',
        result: map.RESULT >= 0 ? String(hit[map.RESULT] || '') : '',
        resultAt: map.RESULT_AT >= 0 ? hit[map.RESULT_AT] : '',
        sourceTab: map.SOURCE_TAB >= 0 ? String(hit[map.SOURCE_TAB] || '') : '',
        updatedAt: map.UPDATED_AT >= 0 ? hit[map.UPDATED_AT] : ''
      }
    };
  } catch (e) {
    return { success: false, found: false, error: e && e.message ? e.message : String(e), payload: null };
  }
}

function commitPilotLessonEvaluation(payload) {
  var body = Object.assign({}, (payload && typeof payload === 'object') ? payload : {});
  body.status = 'SUBMITTED';
  return savePilotLessonEvaluation(body);
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
  const aircraftType = String(aircraftObj.AIRCRAFT_TYPE || aircraftObj.TYPE_FOR_PERFORMANCE || aircraftReg).trim();

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
  let pax = Array.isArray(legPayload.pax) ? legPayload.pax.slice() : [];
  
  // For training missions: if there's a copilot or trainee, add them to pax if not already present
  const copilotName = String(dispatch.COPILOT || '').trim();
  const training = (legPayload && legPayload.training && typeof legPayload.training === 'object')
    ? legPayload.training
    : ((rawData && rawData.training && typeof rawData.training === 'object') ? rawData.training : null);
  const studentName = training && String(training.student || '').trim();
  
  // Add trainee/student to pax if they're not the pilot and not already in pax
  const traineeToAdd = studentName || copilotName;
  if (traineeToAdd && String(traineeToAdd).toUpperCase() !== String(pilotName).toUpperCase()) {
    const alreadyInPax = pax.some(p => String(p.name || '').trim().toUpperCase() === String(traineeToAdd).toUpperCase());
    if (!alreadyInPax) {
      // Try to look up trainee weight from pilots sheet
      const traineeRow = findByAny(pilotsTable.rows, pilotsTable.headers, ['PILOT_NAME', 'NAME'], traineeToAdd);
      const traineeWeight = traineeRow ? toNum(rowToObj(pilotsTable.headers, traineeRow).WEIGHT_KGS, 80) : 80;
      pax.push({
        name: traineeToAdd,
        weight: traineeWeight,
        actualWeight: traineeWeight,
        plannedWeight: traineeWeight,
        verified: true
      });
    }
  }
  
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
      // Support amphib limits when available in DB_Aircraft; legacy MTOW remains supported.
      MTOW_LAND: (function() {
        var v = toNum(aircraftObj.MTOW_LAND, 0);
        if (v > 0) return v;
        v = toNum(aircraftObj.MAX_TAKEOFF_WEIGHT_LAND, 0);
        if (v > 0) return v;
        v = toNum(aircraftObj.TAKEOFF_WEIGHT_LAND, 0);
        if (v > 0) return v;
        return 0;
      })(),
      MTOW_WATER: (function() {
        var v = toNum(aircraftObj.MTOW_WATER, 0);
        if (v > 0) return v;
        v = toNum(aircraftObj.MAX_TAKEOFF_WEIGHT_WATER, 0);
        if (v > 0) return v;
        v = toNum(aircraftObj.TAKEOFF_WEIGHT_WATER, 0);
        if (v > 0) return v;
        return 0;
      })(),
      MLW_LAND: (function() {
        var v = toNum(aircraftObj.MLW_LAND, 0);
        if (v > 0) return v;
        v = toNum(aircraftObj.MAX_LANDING_WEIGHT_LAND, 0);
        if (v > 0) return v;
        v = toNum(aircraftObj.LANDING_WEIGHT_LAND, 0);
        if (v > 0) return v;
        return 0;
      })(),
      MLW_WATER: (function() {
        var v = toNum(aircraftObj.MLW_WATER, 0);
        if (v > 0) return v;
        v = toNum(aircraftObj.MAX_LANDING_WEIGHT_WATER, 0);
        if (v > 0) return v;
        v = toNum(aircraftObj.LANDING_WEIGHT_WATER, 0);
        if (v > 0) return v;
        return 0;
      })(),
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

  const pilotName = String(wbPayload && wbPayload.pilot || '').trim();
  if (pilotName) {
    logSheet.getRange(targetRow, LOG_FLIGHT_COL.PILOT + 1).setValue(pilotName);
  }
  logSheet.getRange(targetRow, LOG_FLIGHT_COL.ACTUAL_LOAD_JSON + 1).setValue(JSON.stringify(merged));
  return true;
}

function getPilotAircraftAuthorizations(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var targetEmail = String(body.pilotEmail || body.email || '').trim().toLowerCase();
    var targetName = String(body.pilotName || body.name || '').trim().toUpperCase();
    var targetAircraft = _toolsCanonicalAircraftType_(body.aircraftType || '').trim().toUpperCase();
    var targetRole = String(body.role || '').trim().toUpperCase();

    var sh = _pilotAuthorizationSheet_();
    var data = sh.getDataRange().getValues();
    if (!data || data.length < 2) return { success: true, rows: [] };

    var headers = data[0];
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var rec = _pilotAuthorizationRecordFromRow_(headers, data[i], i + 1);
      if (!rec.pilotName && !rec.pilotEmail) continue;
      if (!_pilotAuthMatchesPerson_(rec, targetEmail, targetName)) continue;
      if (targetAircraft && _toolsCanonicalAircraftType_(rec.aircraftType || '').trim().toUpperCase() !== targetAircraft) continue;
      if (targetRole && String(rec.role || '').trim().toUpperCase() !== targetRole) continue;
      rows.push(rec);
    }

    rows.sort(function(a, b) {
      var ak = [a.pilotName, a.aircraftType, a.role, a.expiryDate].join('|').toUpperCase();
      var bk = [b.pilotName, b.aircraftType, b.role, b.expiryDate].join('|').toUpperCase();
      return ak.localeCompare(bk);
    });

    return { success: true, rows: rows };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function isPilotAuthorizedForAircraftRole_(pilotEmail, pilotName, aircraftType, role, asOfDate) {
  var asOf = safeDateStr(asOfDate || new Date());
  var auth = getPilotAircraftAuthorizations({
    pilotEmail: pilotEmail,
    pilotName: pilotName,
    aircraftType: aircraftType,
    role: role
  });
  if (!auth || !auth.success || !Array.isArray(auth.rows) || !auth.rows.length) return false;

  for (var i = 0; i < auth.rows.length; i++) {
    var rec = auth.rows[i];
    if (!rec.isActive) continue;
    if (rec.expiryDate && rec.expiryDate < asOf) continue;
    return true;
  }
  return false;
}

function saveToolsPilotAuthorization(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var sh = _pilotAuthorizationSheet_();
    var headers = _toolsSheetHeaderRow_(sh);
    var data = sh.getDataRange().getValues();
    var rowNumber = Number(body.rowNumber || 0);
    var authorizationId = String(body.authorizationId || '').trim();
    var pilotName = String(body.pilotName || '').trim();
    var pilotEmail = String(body.pilotEmail || '').trim().toLowerCase();
    var aircraftType = _toolsCanonicalAircraftType_(body.aircraftType || '').trim();
    var role = String(body.role || '').trim();

    if (!pilotName) return { success: false, error: 'Pilot name is required' };
    if (!pilotEmail) return { success: false, error: 'Pilot email is required' };
    if (!aircraftType) return { success: false, error: 'Aircraft type is required' };
    if (!role) return { success: false, error: 'Authorization role is required' };

    if (rowNumber < 2 && authorizationId) {
      for (var i = 1; i < data.length; i++) {
        var rec = _pilotAuthorizationRecordFromRow_(headers, data[i], i + 1);
        if (String(rec.authorizationId || '').trim() === authorizationId) {
          rowNumber = i + 1;
          break;
        }
      }
    }

    var now = safeDateStr(new Date());
    if (!authorizationId) authorizationId = 'PAUTH_' + new Date().getTime() + '_' + String(role || '').trim().toUpperCase();
    var dataMap = {
      AUTHORIZATION_ID: authorizationId,
      PILOT_NAME: pilotName,
      PILOT_EMAIL: pilotEmail,
      STAFF_ID: String(body.staffId || '').trim(),
      AIRCRAFT_TYPE: aircraftType,
      ROLE: role,
      AUTHORIZATION_TYPE: String(body.authorizationType || 'INITIAL').trim().toUpperCase(),
      STATUS: String(body.status || 'ACTIVE').trim().toUpperCase(),
      DATE_AUTHORIZED: safeDateStr(body.dateAuthorized || now),
      EXPIRY_DATE: safeDateStr(body.expiryDate || ''),
      INSTRUCTOR_SIGNOFF: String(body.instructorSignoff || '').trim(),
      SOURCE: String(body.source || '').trim(),
      NOTES: String(body.notes || '').trim(),
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
      return { success: true, action: 'updated', rowNumber: rowNumber, authorizationId: authorizationId };
    }

    dataMap.CREATED_AT = now;
    var row = headers.map(function(header) {
      var key = _toolsNormHeader_(header);
      return Object.prototype.hasOwnProperty.call(dataMap, key) ? dataMap[key] : '';
    });
    sh.appendRow(row);
    return { success: true, action: 'created', rowNumber: sh.getLastRow(), authorizationId: authorizationId };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function deleteToolsPilotAuthorization(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var sh = _pilotAuthorizationSheet_();
    var rowNumber = Number(body.rowNumber || 0);
    if (rowNumber < 2) return { success: false, error: 'Authorization row not found' };
    sh.deleteRow(rowNumber);
    return { success: true, deletedRowNumber: rowNumber };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function saveRunwayBriefingCard(icao, rwyIdent, cardData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dbSheet = ss.getSheetByName('DB_Airports');
    if (!dbSheet) return { success: false, error: 'DB_Airports sheet not found' };

    var dbData = dbSheet.getDataRange().getValues();
    if (!dbData.length) return { success: false, error: 'DB_Airports is empty' };

    var cols = _runwayDbFindCols_(dbData[0]);
    if (cols.icao < 0 || cols.runway < 0 || cols.knownFeatures < 0) {
      return { success: false, error: 'DB_Airports missing ICAO/RWY_IDENT/KNOWN_FEATURES columns' };
    }

    var icaoNorm  = String(icao || '').trim().toUpperCase();
    var pairKey   = _runwayPairKey_(rwyIdent);
    var rwyNorm   = String(rwyIdent || '').trim().toUpperCase();
    var nowIso    = new Date().toISOString();
    var updated   = 0;

    // Sanitise card: remove any injected keys, keep known fields only
    var allowedKeys = ['restrictions','obstructions','weather','dangers','forcedLanding',
               'routes','rwyhist','incidents','othernotes','airstripPhoto',
               'mapDistM','mapDistM2','depTracePoints','revisedAt','revisedBy','verifyFlags',
                             'mapLabel1','mapLabel2',
               'runwayClass','elevation','surface','cutdownAreaM','cutdownAreaLabel','cutdownSemRestricao',
               'internalLengthM','internalWidthM','slopeSegments','noSignificantSlope',
               'knownFeatures','obstacleAngles50m','obstacleAngles','departureArrow'];
    var safeCard = {};
    allowedKeys.forEach(function(k) {
      if (cardData && cardData[k] !== undefined) safeCard[k] = cardData[k];
    });
    safeCard.savedAt = nowIso;

    var restrictions = (safeCard && Array.isArray(safeCard.restrictions)) ? safeCard.restrictions : [];
    var mtowMapFromRestrictions = {};
    restrictions.forEach(function(rec) {
      if (!rec || typeof rec !== 'object') return;
      var mtow = Number(rec.mtow || 0);
      if (!isFinite(mtow) || mtow <= 0) return;
      var key = _airportNormalizeMtowKey_(rec.acft || '');
      if (!key) return;
      mtowMapFromRestrictions[key] = Math.round(mtow);
    });

    for (var i = 1; i < dbData.length; i++) {
      var rowIcao = String(dbData[i][cols.icao] || '').trim().toUpperCase();
      var rowRwy  = String(dbData[i][cols.runway] || '').trim().toUpperCase();
      if (rowIcao !== icaoNorm) continue;
      // Match this direction OR its pair (both directions of same strip)
      var match = rowRwy === rwyNorm || (pairKey && _runwayPairKey_(rowRwy) === pairKey);
      if (!match) continue;

      var existingRaw = String(dbData[i][cols.knownFeatures] || '').trim();
      var existingObj = _parseJsonLoose_(existingRaw, {});
      var normalised  = Array.isArray(existingObj) ? { features: existingObj } : (existingObj || {});
      normalised.briefingCard = safeCard;

      dbSheet.getRange(i + 1, cols.knownFeatures + 1).setValue(JSON.stringify(normalised));

      if (cols.mtowByModel >= 0 && Object.keys(mtowMapFromRestrictions).length) {
        var existingMtowRaw = String(dbData[i][cols.mtowByModel] || '').trim();
        var existingMtowObj = _parseJsonLoose_(existingMtowRaw, {});
        var mtowOut = (existingMtowObj && typeof existingMtowObj === 'object' && !Array.isArray(existingMtowObj)) ? existingMtowObj : {};
        Object.keys(mtowMapFromRestrictions).forEach(function(k) {
          mtowOut[k] = mtowMapFromRestrictions[k];
        });
        dbSheet.getRange(i + 1, cols.mtowByModel + 1).setValue(JSON.stringify(mtowOut));
      }

      updated++;
    }

    if (!updated) return { success: false, error: 'Runway not found in DB_Airports: ' + icaoNorm + ' ' + rwyNorm };
    return { success: true, updatedRows: updated, savedAt: nowIso };
  } catch (e) {
    return { success: false, error: String(e && e.message || e) };
  }
}

function getDutyFlightVisibilityDiagnostics(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var access = _dutyAppResolvePilotAccess_(body.pilotName, body.preferredDutyEmail);
    if (!access.success) return access;

    var ymd = String(body.dateYmd || _dutyAppTodayBsb_()).trim();
    var fromYmd = String(body.fromYmd || '').trim();
    var toYmd = String(body.toYmd || ymd).trim();
    if (!fromYmd) {
      var fromDate = new Date();
      fromDate.setDate(fromDate.getDate() - 30);
      fromYmd = _dutyYmd_(fromDate);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pilot = access.pilotName;
    var pilotKeys = _dutyPilotKeys_(pilot);
    var report = {
      success: true,
      pilotName: pilot,
      pilotKeys: pilotKeys,
      targetDateYmd: ymd,
      range: { fromYmd: fromYmd, toYmd: toYmd },
      flightSummary: _dutyDailyFlightSummary_(ss, pilot, ymd),
      dutyRowsForDay: [],
      flightRowsForDayMatched: [],
      flightRowsRejectedSample: [],
      toolsDutyRowsForDay: []
    };

    var dutySheet = ss.getSheetByName(APP_SHEETS.DUTY_LOG);
    if (dutySheet) {
      var dutyData = dutySheet.getDataRange().getValues();
      for (var i = 1; i < dutyData.length; i++) {
        var drow = dutyData[i];
        var dPilot = String(drow[DUTY_LOG_COL.PILOT] || '').trim();
        var dYmd = _dutyYmd_(drow[DUTY_LOG_COL.DATE]);
        if (dYmd !== ymd) continue;
        if (!_dutyPilotLikelySame_(dPilot, pilot)) continue;
        report.dutyRowsForDay.push({
          rowNumber: i + 1,
          pilot: dPilot,
          dateYmd: dYmd,
          title: String(drow[DUTY_LOG_COL.TITLE] || '').trim()
        });
      }
    }

    var flightSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
    if (flightSheet) {
      var flightData = flightSheet.getDataRange().getValues();
      for (var j = 1; j < flightData.length; j++) {
        var frow = flightData[j];
        var fPilot = String(frow[LOG_FLIGHT_COL.PILOT] || '').trim();
        var fYmd = _dutyYmd_(frow[LOG_FLIGHT_COL.DATE]);
        var samePilot = _dutyPilotLikelySame_(fPilot, pilot);
        var sameDate = (fYmd === ymd);
        var inRange = (!fromYmd || fYmd >= fromYmd) && (!toYmd || fYmd <= toYmd);
        var totalTimeRaw = String(frow[LOG_FLIGHT_COL.TOTAL_TIME] || '').trim();
        var hours = _dutyParseHhmmToHours_(totalTimeRaw);
        var base = {
          rowNumber: j + 1,
          flightId: String(frow[LOG_FLIGHT_COL.FLIGHT_ID] || '').trim(),
          pilot: fPilot,
          dateRaw: String(frow[LOG_FLIGHT_COL.DATE] || ''),
          dateYmd: fYmd,
          totalTime: totalTimeRaw,
          hours: Number((hours || 0).toFixed(2))
        };

        if (samePilot && sameDate) {
          report.flightRowsForDayMatched.push(base);
        } else if (report.flightRowsRejectedSample.length < 25 && (sameDate || inRange)) {
          var reason = [];
          if (!samePilot) reason.push('pilot_mismatch');
          if (!sameDate) reason.push('date_mismatch');
          if (!inRange) reason.push('outside_range');
          report.flightRowsRejectedSample.push({
            rowNumber: base.rowNumber,
            flightId: base.flightId,
            pilot: base.pilot,
            dateRaw: base.dateRaw,
            dateYmd: base.dateYmd,
            reason: reason.join('|')
          });
        }
      }
    }

    var toolsRows = getToolsDutyLogs({
      fromYmd: fromYmd,
      toYmd: toYmd,
      pilotFilter: pilot,
      preferredDutyEmail: body.preferredDutyEmail
    });
    if (toolsRows && toolsRows.success && Array.isArray(toolsRows.rows)) {
      report.toolsDutyRowsForDay = toolsRows.rows
        .filter(function(r) {
          return String(r.dateYmd || '') === ymd && _dutyPilotLikelySame_(String(r.pilotName || ''), pilot);
        })
        .map(function(r) {
          return {
            rowNumber: Number(r.rowNumber || 0),
            status: String(r.status || ''),
            flightHours: String(r.flightHours || ''),
            derivedFlightHours: Number(r.derivedFlightHours || 0),
            hasFlights: Array.isArray(r.flights) && r.flights.length > 0
          };
        });
    }

    report.counts = {
      dutyRowsForDay: report.dutyRowsForDay.length,
      flightRowsMatched: report.flightRowsForDayMatched.length,
      toolsDutyRowsForDay: report.toolsDutyRowsForDay.length
    };

    return report;
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getPilotRunwayAuthorizations(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var targetEmail = String(body.pilotEmail || body.email || '').trim().toLowerCase();
    var targetName = String(body.pilotName || body.name || '').trim().toUpperCase();
    var targetIcao = String(body.icao || body.airportIcao || '').trim().toUpperCase();
    var targetRunway = String(body.runwayIdent || body.runway || '').trim().toUpperCase();

    var sh = _pilotRunwayChecksSheet_();
    var data = sh.getDataRange().getValues();
    if (!data || data.length < 2) return { success: true, rows: [] };

    var headers = data[0];
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var rec = _pilotRunwayCheckRecordFromRow_(headers, data[i], i + 1);
      if (!rec.pilotName && !rec.pilotEmail) continue;
      if (!_pilotAuthMatchesPerson_(rec, targetEmail, targetName)) continue;
      if (targetIcao && String(rec.icao || '').trim().toUpperCase() !== targetIcao) continue;
      if (targetRunway && String(rec.runwayIdent || '').trim().toUpperCase() !== targetRunway) continue;
      rows.push(rec);
    }

    rows.sort(function(a, b) {
      var ak = [a.pilotName, a.icao, a.runwayIdent, a.expiryDate].join('|').toUpperCase();
      var bk = [b.pilotName, b.icao, b.runwayIdent, b.expiryDate].join('|').toUpperCase();
      return ak.localeCompare(bk);
    });

    return { success: true, rows: rows };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}
