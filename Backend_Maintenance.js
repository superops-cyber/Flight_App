function getDispatchCvaDiagnostic_(payload) {
  var body = (payload && typeof payload === 'object') ? payload : {};
  var requestedReg = String(body.reg || body.aircraftReg || body.aircraft || '').trim().toUpperCase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getRequiredSheet_(ss, APP_SHEETS.AIRCRAFT, 'getDispatchCvaDiagnostic_');
  var values = sheet.getDataRange().getValues();
  var headers = values.length ? (values[0] || []) : [];
  var normalizedHeaders = headers.map(function(h) {
    return String(h || '').trim().toUpperCase();
  });
  var regIdx = _cvaHeaderIndexFromCandidates_(headers, ['REGISTRATION', 'REG', 'TAIL', 'TAIL #']);
  var cvaIdx = _cvaHeaderIndexFromCandidates_(headers, ['ANNUAL DUE_(CVA)']);
  var matchedRowIndex = -1;
  var matchedRow = null;

  if (regIdx >= 0 && requestedReg) {
    for (var i = 1; i < values.length; i++) {
      var rowReg = String(values[i][regIdx] || '').trim().toUpperCase();
      if (rowReg && rowReg === requestedReg) {
        matchedRowIndex = i;
        matchedRow = values[i];
        break;
      }
    }
  }

  var dropdownData = null;
  try {
    dropdownData = getDropdownData();
  } catch (e) {
    dropdownData = { error: e && e.message ? e.message : String(e) };
  }

  var appDataAircraft = dropdownData && Array.isArray(dropdownData.aircraft) ? dropdownData.aircraft : [];
  var appDataAircraftEntry = null;
  if (requestedReg && appDataAircraft.length) {
    for (var j = 0; j < appDataAircraft.length; j++) {
      var item = appDataAircraft[j] || {};
      if (String(item.reg || '').trim().toUpperCase() === requestedReg) {
        appDataAircraftEntry = item;
        break;
      }
    }
  }

  return {
    success: true,
    spreadsheetId: ss.getId(),
    sheetName: sheet.getName(),
    requestedReg: requestedReg,
    serverNow: new Date().toISOString(),
    headerCount: headers.length,
    rawHeaders: headers,
    normalizedHeaders: normalizedHeaders,
    registrationColumnIndex: regIdx,
    cvaColumnIndex: cvaIdx,
    cvaHeaderRaw: cvaIdx >= 0 ? String(headers[cvaIdx] || '') : '',
    matchedRowIndex: matchedRowIndex,
    matchedRegistration: matchedRow && regIdx >= 0 ? String(matchedRow[regIdx] || '').trim().toUpperCase() : '',
    cvaRawValue: matchedRow && cvaIdx >= 0 ? matchedRow[cvaIdx] : '',
    cvaIsoValue: matchedRow && cvaIdx >= 0 ? safeDateStr(matchedRow[cvaIdx]) : '',
    cvaValueType: matchedRow && cvaIdx >= 0 ? Object.prototype.toString.call(matchedRow[cvaIdx]) : '',
    firstAircraftRows: values.slice(1, 4),
    appDataAircraftCount: appDataAircraft.length,
    appDataAircraftEntry: appDataAircraftEntry,
    appDataAircraftKeys: appDataAircraftEntry ? Object.keys(appDataAircraftEntry) : [],
    getDropdownDataError: dropdownData && dropdownData.error ? dropdownData.error : ''
  };
}

function getDispatchCvaDiagnosticForPtKps_() {
  return getDispatchCvaDiagnostic_({ reg: 'PT-KPS' });
}

function getMxApprovalQueue() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, 'getMxApprovalQueue');
    var data = sheet.getDataRange().getValues();
    var missions = {};
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var missionId = String(row[DISPATCH_COL.MISSION_ID] || '').trim();
      if (!missionId || missions[missionId]) continue;
      var mxReview = _extractMissionMxReviewFromRaw_(row[DISPATCH_COL.RAW_DATA]);
      if (!mxReview || !mxReview.required) continue;
      var missionFlightHours = 0;
      for (var j = 1; j < data.length; j++) {
        if (String(data[j][DISPATCH_COL.MISSION_ID] || '').trim() !== missionId) continue;
        var rowFlightHours = Number(data[j][DISPATCH_COL.FLIGHT_TIME] || 0);
        if (isFinite(rowFlightHours) && rowFlightHours > 0) missionFlightHours += rowFlightHours;
      }
      var computedReview = _buildMissionMxReview_(row[DISPATCH_COL.AIRCRAFT], missionFlightHours, mxReview);
      missions[missionId] = {
        missionId: missionId,
        date: row[DISPATCH_COL.DATE] instanceof Date ? row[DISPATCH_COL.DATE].toISOString().split('T')[0] : String(row[DISPATCH_COL.DATE] || ''),
        aircraft: String(row[DISPATCH_COL.AIRCRAFT] || '').trim(),
        pilot: String(row[DISPATCH_COL.PILOT] || '').trim(),
        type: String(row[DISPATCH_COL.TYPE] || '').trim(),
        route: String(row[DISPATCH_COL.ROUTE] || '').trim(),
        status: String(row[DISPATCH_COL.STATUS] || 'PENDING').trim().toUpperCase(),
        flightDescription: _extractMissionDescriptionFromRaw_(row[DISPATCH_COL.RAW_DATA]),
        missionFlightHours: missionFlightHours,
        mxReview: computedReview
      };
    }
    var rows = Object.keys(missions).map(function(id) { return missions[id]; }).sort(function(a, b) {
      var aStatus = String(a.mxReview && a.mxReview.status || '').trim().toUpperCase();
      var bStatus = String(b.mxReview && b.mxReview.status || '').trim().toUpperCase();
      if (aStatus === 'PENDING' && bStatus !== 'PENDING') return -1;
      if (aStatus !== 'PENDING' && bStatus === 'PENDING') return 1;
      return String(a.date || '').localeCompare(String(b.date || ''));
    });
    return { success: true, rows: rows };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e), rows: [] };
  }
}

function getMaintenanceScheduleData(payload) {
  try {
    _mxEnsureFrameworkSheets_();
    var body = (payload && typeof payload === 'object') ? payload : {};
    var filterReg = String(body.aircraftReg || '').trim().toUpperCase();
    var filterRegNorm = _toolsNormalizeAircraftReg_(filterReg);
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
      if (filterReg) {
        if (aircraftReg !== filterReg && _toolsNormalizeAircraftReg_(aircraftReg) !== filterRegNorm) continue;
      }

      var assignmentId = String(row[idx.ASSIGNMENT_ID] || '').trim();
      var dueMode = _mxNormalizeDueMode_(row[idx.DUE_MODE], false, false);
      var intervalHours = _mxParseNum_(row[idx.INTERVAL_HOURS]);
      var intervalDays = _mxParseNum_(row[idx.INTERVAL_DAYS]);
      var intervalMonths = _mxParseNum_(row[idx.INTERVAL_MONTHS]);
      var intervalYears = _mxParseNum_(row[idx.INTERVAL_YEARS]);
      var hasHours = intervalHours !== '' && intervalHours > 0;
      var hasCalendar = (intervalDays !== '' && intervalDays > 0) || (intervalMonths !== '' && intervalMonths > 0) || (intervalYears !== '' && intervalYears > 0) || (idx.DUE_DATE != null && String(row[idx.DUE_DATE] || '').trim() !== '');
      dueMode = _mxNormalizeDueMode_(dueMode, hasHours, hasCalendar);
      if (!hasHours && !hasCalendar) continue;

      var startTach = _mxParseNum_(row[idx.START_TACH]);
      var startDate = _mxIsoDate_(row[idx.START_DATE]);
      var lastPerformedTach = idx.LAST_PERFORMED_TACH != null ? _mxParseNum_(row[idx.LAST_PERFORMED_TACH]) : '';
      var lastPerformedDate = idx.LAST_PERFORMED_DATE != null ? _mxIsoDate_(row[idx.LAST_PERFORMED_DATE]) : '';
      var dueDate = (idx.DUE_DATE != null) ? _mxIsoDate_(row[idx.DUE_DATE]) : '';
      var last = latestLog[assignmentId] || null;
      var baseTach = (last && last.completedTach !== '') ? last.completedTach : (lastPerformedTach !== '' ? lastPerformedTach : startTach);
      var baseDate = (last && last.completedDate) ? last.completedDate : (lastPerformedDate || startDate);

      var nextDueTach = '';
      if (hasHours && baseTach !== '') nextDueTach = parseFloat((baseTach + intervalHours).toFixed(1));
      var nextDueDate = '';
      if (dueDate) nextDueDate = dueDate;
      else if (hasCalendar && baseDate) nextDueDate = _mxCalendarDueDate_(baseDate, intervalDays, intervalMonths, intervalYears);

      var currentTach = tachByReg[aircraftReg];
      var hoursRemaining = '';
      if (hasHours && nextDueTach !== '' && currentTach !== '') {
        hoursRemaining = parseFloat((nextDueTach - currentTach).toFixed(1));
      }
      var hasDateTrack = !!nextDueDate;
      var daysRemaining = hasDateTrack ? _mxDaysRemaining_(nextDueDate) : '';
      var dueState = _mxDueState_(hoursRemaining, daysRemaining, dueMode, thresholdHours, thresholdDays);

      rows.push({
        assignmentId: assignmentId,
        aircraftReg: aircraftReg,
        taskCode: String(row[idx.TASK_CODE] || '').trim(),
        taskName: String(row[idx.TASK_NAME] || '').trim(),
        category: String(row[idx.CATEGORY] || '').trim(),
        reference: String(row[idx.REFERENCE] || '').trim(),
        dueMode: dueMode,
        intervalHours: hasHours ? intervalHours : '',
        intervalDays: intervalDays !== '' ? intervalDays : '',
        intervalMonths: intervalMonths !== '' ? intervalMonths : '',
        intervalYears: intervalYears !== '' ? intervalYears : '',
        startTach: startTach,
        startDate: startDate,
        lastPerformedTach: idx.LAST_PERFORMED_TACH != null ? (lastPerformedTach !== '' ? lastPerformedTach : '') : '',
        lastPerformedDate: idx.LAST_PERFORMED_DATE != null ? lastPerformedDate : '',
        dueDate: dueDate,
        dueDateRaw: dueDate,
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
          dueMode: r.dueMode,
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
    var intervalMonths = _mxParseNum_(body.intervalMonths);
    var intervalYears = _mxParseNum_(body.intervalYears);
    var dueMode = _mxNormalizeDueMode_(body.dueMode, intervalHours !== '', (intervalDays !== '' || intervalMonths !== '' || intervalYears !== ''));
    var dueDate = _mxIsoDate_(body.dueDate);
    if (intervalHours === '' && intervalDays === '' && intervalMonths === '' && intervalYears === '' && !dueDate) return { success: false, error: 'Provide at least one interval or a due date.' };

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
    if (idx.DUE_MODE != null) row[idx.DUE_MODE] = dueMode;
    row[idx.INTERVAL_HOURS] = intervalHours === '' ? '' : intervalHours;
    row[idx.INTERVAL_DAYS] = intervalDays === '' ? '' : intervalDays;
    if (idx.INTERVAL_MONTHS != null) row[idx.INTERVAL_MONTHS] = intervalMonths === '' ? '' : intervalMonths;
    if (idx.INTERVAL_YEARS != null) row[idx.INTERVAL_YEARS] = intervalYears === '' ? '' : intervalYears;
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
    var intervalMonths = _mxParseNum_(body.intervalMonths);
    var intervalYears = _mxParseNum_(body.intervalYears);
    var dueMode = _mxNormalizeDueMode_(body.dueMode, intervalHours !== '', (intervalDays !== '' || intervalMonths !== '' || intervalYears !== ''));
    if (intervalHours === '' && intervalDays === '' && intervalMonths === '' && intervalYears === '') return { success: false, error: 'Provide at least one interval.' };

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
    if (idx.DUE_MODE != null) row[idx.DUE_MODE] = dueMode;
    row[idx.INTERVAL_HOURS] = intervalHours === '' ? '' : intervalHours;
    row[idx.INTERVAL_DAYS] = intervalDays === '' ? '' : intervalDays;
    if (idx.INTERVAL_MONTHS != null) row[idx.INTERVAL_MONTHS] = intervalMonths === '' ? '' : intervalMonths;
    if (idx.INTERVAL_YEARS != null) row[idx.INTERVAL_YEARS] = intervalYears === '' ? '' : intervalYears;
    var lastPerformedTach = _mxParseNum_(body.lastPerformedTach !== undefined ? body.lastPerformedTach : body.startTach);
    var lastPerformedDate = _mxIsoDate_(body.lastPerformedDate || body.startDate || new Date()) || _mxIsoDate_(new Date());
    row[idx.START_TACH] = lastPerformedTach;
    row[idx.START_DATE] = lastPerformedDate;
    if (idx.LAST_PERFORMED_TACH != null) row[idx.LAST_PERFORMED_TACH] = lastPerformedTach === '' ? '' : lastPerformedTach;
    if (idx.LAST_PERFORMED_DATE != null) row[idx.LAST_PERFORMED_DATE] = lastPerformedDate || '';
    if (idx.DUE_DATE != null) row[idx.DUE_DATE] = dueDate || '';
    row[idx.ACTIVE] = String(body.active || 'Y').trim().toUpperCase() === 'N' ? 'N' : 'Y';
    row[idx.SOURCE] = String(body.source || 'operator').trim();
    row[idx.CAMO_KEY] = String(body.camoKey || '').trim();
    row[idx.NOTES] = String(body.notes || '').trim();
    row[idx.CREATED_AT] = nowIso;
    row[idx.UPDATED_AT] = nowIso;

    var cvaSync = null;
    var dueDate = idx.DUE_DATE != null ? String(row[idx.DUE_DATE] || '').trim() : '';
    if (dueDate && _mxLooksLikeCvaMaintenance_(taskCode, taskName, body.category, body.reference, body.notes)) {
      cvaSync = _mxSyncAircraftCvaExpiry_(aircraftReg, dueDate);
      if (cvaSync && cvaSync.success === false) {
        throw new Error(cvaSync.error || 'Failed to sync CVA expiry to DB_Aircraft.');
      }
    }

    sh.appendRow(row);

    var calendarSync = _mxSyncMaintenanceAssignmentCalendar_(assignmentId, aircraftReg);
    if (calendarSync && calendarSync.success === false) {
      appLog_('Maintenance calendar sync failed for', assignmentId, calendarSync.error || 'unknown error');
    }

    return { success: true, assignmentId: assignmentId, cvaSynced: !!(cvaSync && cvaSync.success), calendarSynced: !!(calendarSync && calendarSync.success) };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function updateMaintenanceAssignment(payload) {
  try {
    _mxEnsureFrameworkSheets_();
    var body = (payload && typeof payload === 'object') ? payload : {};
    var assignmentId = String(body.assignmentId || '').trim();
    if (!assignmentId) return { success: false, error: 'assignmentId is required.' };

    var supervisorPassword = String(body.supervisorPassword || body.password || '').trim();
    if (!supervisorPassword) return { success: false, error: 'Supervisor password is required.' };
    _verifySupervisorApprovalPassword_(supervisorPassword);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getRequiredSheet_(ss, APP_SHEETS.MAINT_ASSIGNMENTS, 'updateMaintenanceAssignment');
    var data = sh.getDataRange().getValues();
    if (!data || data.length < 2) return { success: false, error: 'Maintenance assignment not found.' };

    var headers = data[0];
    var idx = _mxNormHeaderMap_(headers);
    var rowNumber = -1;
    var currentRow = null;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idx.ASSIGNMENT_ID] || '').trim() !== assignmentId) continue;
      rowNumber = i + 1;
      currentRow = data[i];
      break;
    }
    if (rowNumber < 0 || !currentRow) return { success: false, error: 'Maintenance assignment not found.' };

    var aircraftReg = String(body.aircraftReg != null ? body.aircraftReg : currentRow[idx.AIRCRAFT_REG] || '').trim().toUpperCase();
    var taskCode = String(body.taskCode != null ? body.taskCode : currentRow[idx.TASK_CODE] || '').trim().toUpperCase();
    var taskName = String(body.taskName != null ? body.taskName : currentRow[idx.TASK_NAME] || '').trim();
    if (!aircraftReg || !taskName) return { success: false, error: 'aircraftReg and taskName are required.' };

    var intervalHours = body.intervalHours !== undefined ? _mxParseNum_(body.intervalHours) : _mxParseNum_(currentRow[idx.INTERVAL_HOURS]);
    var intervalDays = body.intervalDays !== undefined ? _mxParseNum_(body.intervalDays) : _mxParseNum_(currentRow[idx.INTERVAL_DAYS]);
    var intervalMonths = body.intervalMonths !== undefined && idx.INTERVAL_MONTHS != null ? _mxParseNum_(body.intervalMonths) : (idx.INTERVAL_MONTHS != null ? _mxParseNum_(currentRow[idx.INTERVAL_MONTHS]) : '');
    var intervalYears = body.intervalYears !== undefined && idx.INTERVAL_YEARS != null ? _mxParseNum_(body.intervalYears) : (idx.INTERVAL_YEARS != null ? _mxParseNum_(currentRow[idx.INTERVAL_YEARS]) : '');
    var dueMode = _mxNormalizeDueMode_(body.dueMode != null ? body.dueMode : currentRow[idx.DUE_MODE], intervalHours !== '', (intervalDays !== '' || intervalMonths !== '' || intervalYears !== ''));
    var dueDate = body.dueDate !== undefined ? _mxIsoDate_(body.dueDate) : (idx.DUE_DATE != null ? _mxIsoDate_(currentRow[idx.DUE_DATE]) : '');
    if (intervalHours === '' && intervalDays === '' && intervalMonths === '' && intervalYears === '' && !dueDate) return { success: false, error: 'Provide at least one interval or a due date.' };

    var lastPerformedTach = body.lastPerformedTach !== undefined ? _mxParseNum_(body.lastPerformedTach) : _mxParseNum_(currentRow[idx.LAST_PERFORMED_TACH] != null ? currentRow[idx.LAST_PERFORMED_TACH] : currentRow[idx.START_TACH]);
    var lastPerformedDate = body.lastPerformedDate !== undefined ? _mxIsoDate_(body.lastPerformedDate) : _mxIsoDate_(currentRow[idx.LAST_PERFORMED_DATE] != null ? currentRow[idx.LAST_PERFORMED_DATE] : currentRow[idx.START_DATE]);
    var activeValue = body.active !== undefined ? String(body.active || 'Y').trim().toUpperCase() : String(currentRow[idx.ACTIVE] || 'Y').trim().toUpperCase();
    var sourceValue = body.source !== undefined ? String(body.source || '').trim() : String(currentRow[idx.SOURCE] || '').trim();
    var camoKeyValue = body.camoKey !== undefined ? String(body.camoKey || '').trim() : String(currentRow[idx.CAMO_KEY] || '').trim();
    var notesValue = body.notes !== undefined ? String(body.notes || '').trim() : String(currentRow[idx.NOTES] || '').trim();
    var categoryValue = body.category !== undefined ? String(body.category || '').trim() : String(currentRow[idx.CATEGORY] || '').trim();
    var referenceValue = body.reference !== undefined ? String(body.reference || '').trim() : String(currentRow[idx.REFERENCE] || '').trim();

    sh.getRange(rowNumber, idx.AIRCRAFT_REG + 1).setValue(aircraftReg);
    sh.getRange(rowNumber, idx.TASK_CODE + 1).setValue(taskCode);
    sh.getRange(rowNumber, idx.TASK_NAME + 1).setValue(taskName);
    sh.getRange(rowNumber, idx.CATEGORY + 1).setValue(categoryValue);
    sh.getRange(rowNumber, idx.REFERENCE + 1).setValue(referenceValue);
    if (idx.DUE_MODE != null) sh.getRange(rowNumber, idx.DUE_MODE + 1).setValue(dueMode);
    sh.getRange(rowNumber, idx.INTERVAL_HOURS + 1).setValue(intervalHours === '' ? '' : intervalHours);
    sh.getRange(rowNumber, idx.INTERVAL_DAYS + 1).setValue(intervalDays === '' ? '' : intervalDays);
    if (idx.INTERVAL_MONTHS != null) sh.getRange(rowNumber, idx.INTERVAL_MONTHS + 1).setValue(intervalMonths === '' ? '' : intervalMonths);
    if (idx.INTERVAL_YEARS != null) sh.getRange(rowNumber, idx.INTERVAL_YEARS + 1).setValue(intervalYears === '' ? '' : intervalYears);
    sh.getRange(rowNumber, idx.START_TACH + 1).setValue(lastPerformedTach === '' ? '' : lastPerformedTach);
    sh.getRange(rowNumber, idx.START_DATE + 1).setValue(lastPerformedDate || '');
    if (idx.LAST_PERFORMED_TACH != null) sh.getRange(rowNumber, idx.LAST_PERFORMED_TACH + 1).setValue(lastPerformedTach === '' ? '' : lastPerformedTach);
    if (idx.LAST_PERFORMED_DATE != null) sh.getRange(rowNumber, idx.LAST_PERFORMED_DATE + 1).setValue(lastPerformedDate || '');
    if (idx.DUE_DATE != null) sh.getRange(rowNumber, idx.DUE_DATE + 1).setValue(dueDate || '');
    sh.getRange(rowNumber, idx.ACTIVE + 1).setValue(activeValue === 'N' ? 'N' : 'Y');
    sh.getRange(rowNumber, idx.SOURCE + 1).setValue(sourceValue);
    sh.getRange(rowNumber, idx.CAMO_KEY + 1).setValue(camoKeyValue);
    sh.getRange(rowNumber, idx.NOTES + 1).setValue(notesValue);
    sh.getRange(rowNumber, idx.UPDATED_AT + 1).setValue(_mxIsoDate_(new Date()));

    var cvaSync = null;
    if (dueDate && _mxLooksLikeCvaMaintenance_(taskCode, taskName, categoryValue, referenceValue, notesValue)) {
      cvaSync = _mxSyncAircraftCvaExpiry_(aircraftReg, dueDate);
      if (cvaSync && cvaSync.success === false) {
        throw new Error(cvaSync.error || 'Failed to sync CVA expiry to DB_Aircraft.');
      }
    }

    var calendarSync = _mxSyncMaintenanceAssignmentCalendar_(assignmentId, aircraftReg);
    if (calendarSync && calendarSync.success === false) {
      appLog_('Maintenance calendar sync failed for', assignmentId, calendarSync.error || 'unknown error');
    }

    return { success: true, assignmentId: assignmentId, cvaSynced: !!(cvaSync && cvaSync.success), calendarSynced: !!(calendarSync && calendarSync.success) };
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
      if (asgIdx.LAST_PERFORMED_DATE != null) asgSh.getRange(i + 1, asgIdx.LAST_PERFORMED_DATE + 1).setValue(completedDate);
      if (asgIdx.LAST_PERFORMED_TACH != null && completedTach !== '') asgSh.getRange(i + 1, asgIdx.LAST_PERFORMED_TACH + 1).setValue(completedTach);
      break;
    }

    var calendarSync = _mxSyncMaintenanceAssignmentCalendar_(assignmentId, aircraftReg);
    if (calendarSync && calendarSync.success === false) {
      appLog_('Maintenance calendar sync failed for completion', assignmentId, calendarSync.error || 'unknown error');
    }

    return { success: true, logId: logId, calendarSynced: !!(calendarSync && calendarSync.success) };
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
      if (asgIdx.DUE_MODE != null) newRow[asgIdx.DUE_MODE] = String(tr[tplIdx.DUE_MODE] || '').trim();
      newRow[asgIdx.INTERVAL_HOURS] = _mxParseNum_(tr[tplIdx.INTERVAL_HOURS]);
      newRow[asgIdx.INTERVAL_DAYS] = _mxParseNum_(tr[tplIdx.INTERVAL_DAYS]);
      if (asgIdx.INTERVAL_MONTHS != null) newRow[asgIdx.INTERVAL_MONTHS] = _mxParseNum_(tr[tplIdx.INTERVAL_MONTHS]);
      if (asgIdx.INTERVAL_YEARS != null) newRow[asgIdx.INTERVAL_YEARS] = _mxParseNum_(tr[tplIdx.INTERVAL_YEARS]);
      newRow[asgIdx.START_TACH] = startTach;
      newRow[asgIdx.START_DATE] = startDate;
      if (asgIdx.LAST_PERFORMED_TACH != null) newRow[asgIdx.LAST_PERFORMED_TACH] = startTach;
      if (asgIdx.LAST_PERFORMED_DATE != null) newRow[asgIdx.LAST_PERFORMED_DATE] = startDate;
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

function getDispatchMaintenanceStatus(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var aircraftReg = String(body.aircraftReg || '').trim().toUpperCase();
    var res = getMaintenanceScheduleData({ aircraftReg: aircraftReg });
    if (!res || !res.success) return { success: false, error: (res && res.error) || 'Could not load maintenance schedule.' };
    var rows = Array.isArray(res.rows) ? res.rows.slice() : [];
    var grouped = {};
    rows.forEach(function(row) {
      var reg = String(row && row.aircraftReg || '').trim().toUpperCase();
      if (!reg) return;
      if (!grouped[reg]) grouped[reg] = [];
      grouped[reg].push(row);
    });
    var byAircraft = {};
    Object.keys(grouped).forEach(function(reg) {
      var list = grouped[reg].slice().sort(function(a, b) {
        var aHours = Number(a && a.hoursRemaining);
        var bHours = Number(b && b.hoursRemaining);
        var aVal = isFinite(aHours) ? aHours : Number.POSITIVE_INFINITY;
        var bVal = isFinite(bHours) ? bHours : Number.POSITIVE_INFINITY;
        return aVal - bVal;
      });
      var nextInspection = list.find(function(item) { return isFinite(Number(item && item.hoursRemaining)); }) || null;
      var nextHours = nextInspection ? Number(nextInspection.hoursRemaining) : null;
      byAircraft[reg] = {
        aircraftReg: reg,
        nextInspectionDate: nextInspection && nextInspection.nextDueDate ? String(nextInspection.nextDueDate || '') : '',
        nextInspection: nextInspection ? {
          assignmentId: String(nextInspection.assignmentId || ''),
          taskCode: String(nextInspection.taskCode || ''),
          taskName: String(nextInspection.taskName || ''),
          dueState: String(nextInspection.dueState || ''),
          nextDueTach: nextInspection.nextDueTach,
          hoursRemaining: nextInspection.hoursRemaining,
          nextDueDate: nextInspection.nextDueDate,
          daysRemaining: nextInspection.daysRemaining
        } : null,
        nextHoursLabel: isFinite(nextHours) ? nextHours.toFixed(1) + ' hrs' : '',
        upcomingInspections: list.slice(0, 5).map(function(item) {
          return {
            assignmentId: String(item.assignmentId || ''),
            taskCode: String(item.taskCode || ''),
            taskName: String(item.taskName || ''),
            dueState: String(item.dueState || ''),
            nextDueTach: item.nextDueTach,
            hoursRemaining: item.hoursRemaining,
            nextDueDate: item.nextDueDate,
            daysRemaining: item.daysRemaining
          };
        })
      };
    });
    return { success: true, byAircraft: byAircraft };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e), byAircraft: {} };
  }
}

function approveMissionMxReview(missionId, supervisorName, supervisorNotes, approvalPassword) {
  _verifyMxApprovalPassword_(approvalPassword);
  var name = String(supervisorName || '').trim() || 'Chief Mechanic';
  var notes = String(supervisorNotes || '').trim();
  var nowIso = new Date().toISOString();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, 'approveMissionMxReview');
  _writeMissionMxReview_(sheet, missionId, {
    status: 'APPROVED',
    approvedBy: name,
    approvedAt: nowIso,
    notes: notes,
    rejectedBy: '',
    rejectedAt: '',
    rejectionNotes: ''
  });
  return { success: true, message: 'Chief mechanic approval recorded', missionId: missionId, approvedBy: name, approvedAt: nowIso };
}

function rejectMissionMxReview(missionId, supervisorName, supervisorNotes, approvalPassword) {
  _verifyMxApprovalPassword_(approvalPassword);
  var name = String(supervisorName || '').trim() || 'Chief Mechanic';
  var notes = String(supervisorNotes || '').trim();
  var nowIso = new Date().toISOString();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, 'rejectMissionMxReview');
  _writeMissionMxReview_(sheet, missionId, {
    status: 'REJECTED',
    rejectedBy: name,
    rejectedAt: nowIso,
    rejectionNotes: notes
  });
  return { success: true, message: 'Chief mechanic review rejected', missionId: missionId, rejectedBy: name, rejectedAt: nowIso };
}
