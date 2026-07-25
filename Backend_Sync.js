function syncAllMissionStatuses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dispatchSheet = ss.getSheetByName(APP_SHEETS.DISPATCH);
  var logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  if (!dispatchSheet || !logSheet) return { success: false, error: 'missing sheet' };

  var dispatchData = dispatchSheet.getDataRange().getValues();
  if (!dispatchData || dispatchData.length < 2) return { success: false, error: 'dispatch empty' };

  var logData = logSheet.getDataRange().getValues();
  var completedLegMap = {};
  for (var li = 1; li < logData.length; li++) {
    var lr = logData[li];
    var lid = String(lr[LOG_FLIGHT_COL.FLIGHT_ID] || '').trim();
    if (!lid) continue;
    if (lr[LOG_FLIGHT_COL.LANDED] || lr[LOG_FLIGHT_COL.ON_BLOCKS] || lr[LOG_FLIGHT_COL.BRAKES_APPLIED]) {
      completedLegMap[lid] = true;
    }
  }

  var missionMap = {};
  for (var i = 1; i < dispatchData.length; i++) {
    var mid = String(dispatchData[i][DISPATCH_COL.MISSION_ID] || '').trim();
    var lid2 = String(dispatchData[i][DISPATCH_COL.FLIGHT_ID] || '').trim();
    if (!mid) continue;
    if (!missionMap[mid]) missionMap[mid] = { rows: [], legs: [] };
    missionMap[mid].rows.push(i + 1);
    if (lid2) missionMap[mid].legs.push(lid2);
  }

  var updated = 0;
  Object.keys(missionMap).forEach(function(mid) {
    var entry = missionMap[mid];
    var legCount = entry.legs.length;
    var legsComplete = entry.legs.filter(function(l) { return !!completedLegMap[l]; }).length;
    var targetStatus = (legCount > 0 && legsComplete === legCount) ? 'FLOWN'
      : (legsComplete > 0 ? 'In-Flight' : null);
    if (!targetStatus) return;
    entry.rows.forEach(function(rowNum) {
      var existing = String(dispatchSheet.getRange(rowNum, DISPATCH_COL.STATUS + 1).getValue() || '').trim().toUpperCase();
      if (existing === 'CANCELLED') return;
      if (existing === targetStatus.toUpperCase()) return;
      dispatchSheet.getRange(rowNum, DISPATCH_COL.STATUS + 1).setValue(targetStatus);
      updated++;
    });
  });

  if (updated > 0) invalidateScheduledMissionsCache_();
  return { success: true, missionsScanned: Object.keys(missionMap).length, rowsUpdated: updated };
}

function _formGetImportLogKeys_(logSheet) {
  var keys = {};
  var data = logSheet.getDataRange().getValues();
  if (!data || data.length < 2) return keys;
  var headers = data[0].map(function(v) { return _schemaNormHeader_(v); });
  var idxResponseKey = headers.indexOf('RESPONSE_KEY');
  var idxStatus = headers.indexOf('STATUS');
  if (idxResponseKey < 0) return keys;
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][idxResponseKey] || '').trim();
    if (!key) continue;
    var status = idxStatus >= 0 ? String(data[i][idxStatus] || '').trim().toUpperCase() : 'OK';
    if (status === 'OK') keys[key] = true;
  }
  return keys;
}

function _formAppendImportLog_(logSheet, payload) {
  logSheet.appendRow([
    new Date(),
    String(payload.formName || 'FORM').trim(),
    payload.sourceSheet || '',
    payload.sourceRow || '',
    payload.responseKey || '',
    payload.status || '',
    payload.detail || ''
  ]);
}

function ensureFlightFollowInboxSyncTrigger() {
  var handler = 'syncFlightFollowInboxToSheet';
  var triggers = ScriptApp.getProjectTriggers();
  var exists = triggers.some(function(t) {
    return t.getHandlerFunction && t.getHandlerFunction() === handler;
  });
  if (!exists) {
    ScriptApp.newTrigger(handler).timeBased().everyMinutes(1).create();
  }
  var triggerCount = ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction && t.getHandlerFunction() === handler;
  }).length;
  return {
    success: true,
    exists: exists,
    handler: handler,
    triggerCount: triggerCount,
    note: 'Run this under the hidden FF mailbox account.'
  };
}

function _flightFollowIsSyncRelevantMessage_(from, to, cc, subj, body, primaryInbox, ffInboxes, outboundAddrs) {
  var fromLower = String(from || '').toLowerCase();
  var toLower = String(to || '').toLowerCase();
  var ccLower = String(cc || '').toLowerCase();
  var subjUpper = String(subj || '').toUpperCase();
  var bodyUpper = String(body || '').toUpperCase();
  var subjUpperFlat = subjUpper.replace(/[^A-Z0-9]/g, ' ');
  var bodyUpperFlat = bodyUpper.replace(/[^A-Z0-9]/g, ' ');

  var hasInreachDomain = fromLower.indexOf('inreachmail.com') >= 0 || fromLower.indexOf('@garmin.com') >= 0 ||
    toLower.indexOf('inreachmail.com') >= 0 || toLower.indexOf('@garmin.com') >= 0 ||
    ccLower.indexOf('inreachmail.com') >= 0 || ccLower.indexOf('@garmin.com') >= 0;

  var hasInreachSystemMail = fromLower.indexOf('no.reply.inreach@garmin.com') >= 0 ||
    fromLower.indexOf('no-reply.inreach@garmin.com') >= 0 ||
    /\binreach\s+message\s+from\s+n[0-9a-z]{3,5}\b/i.test(subjUpper) ||
    /\binreach\s+message\s+from\s+n[0-9a-z]{3,5}\b/i.test(bodyUpper);

  var hasFfTag = subjUpper.indexOf('[FF CHAT]') >= 0 || subjUpper.indexOf('FF CHAT') >= 0 ||
    subjUpperFlat.indexOf('FF CHAT') >= 0 || bodyUpper.indexOf('FLIGHT FOLLOW') >= 0 ||
    bodyUpperFlat.indexOf('FLIGHT FOLLOW') >= 0 ||
    subjUpper.indexOf('INREACH MESSAGE FROM') >= 0 ||
    subjUpperFlat.indexOf('INREACH MESSAGE FROM') >= 0 ||
    bodyUpper.indexOf('INREACH MESSAGE FROM') >= 0 ||
    bodyUpperFlat.indexOf('INREACH MESSAGE FROM') >= 0;

  var hasRegToken = /\b(?:PT|PR|PP|PU)-?[A-Z0-9]{3,4}\b/.test(subjUpper) ||
    /\b(?:PT|PR|PP|PU)-?[A-Z0-9]{3,4}\b/.test(bodyUpper);

  // Hard gate: only enforce if primaryInbox is provided. When executing as the
  // primary inbox account and searching in:inbox, the Gmail query already
  // guarantees delivery to that inbox — BCC/forwarded msgs won't have the
  // address in To:/CC: headers, so skip the gate when primaryInbox is falsy.
  var touchesPrimaryInbox = !!primaryInbox && (
    fromLower.indexOf(primaryInbox) >= 0 ||
    toLower.indexOf(primaryInbox) >= 0 ||
    ccLower.indexOf(primaryInbox) >= 0
  );
  if (primaryInbox && !touchesPrimaryInbox) return false;

  var touchesFfInbox = (ffInboxes || []).some(function(addr) {
    return !!addr && (fromLower.indexOf(addr) >= 0 || toLower.indexOf(addr) >= 0 || ccLower.indexOf(addr) >= 0);
  });

  var touchesOutbound = (outboundAddrs || []).some(function(addr) {
    return !!addr && (fromLower.indexOf(addr) >= 0 || toLower.indexOf(addr) >= 0 || ccLower.indexOf(addr) >= 0 ||
      String(subj || '').toLowerCase().indexOf(addr) >= 0 || String(body || '').toLowerCase().indexOf(addr) >= 0);
  });

  // Single source of truth: require routing through primary inbox and FF-related signal.
  return hasInreachDomain || hasInreachSystemMail || touchesOutbound || touchesFfInbox || hasFfTag || hasRegToken;
}

function _flightFollowSyncIdentityGate_(primaryInbox) {
  var requiredInbox = String(primaryInbox || '').trim().toLowerCase();
  var effectiveUser = '';
  var activeUser = '';
  try { effectiveUser = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase(); } catch (e0) {}
  try { activeUser = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); } catch (e1) {}

  var allowed = !!requiredInbox && !!effectiveUser && effectiveUser === requiredInbox;
  return {
    allowed: allowed,
    requiredInbox: requiredInbox,
    effectiveUser: effectiveUser,
    activeUser: activeUser
  };
}

function syncFlightFollowInboxToSheet(options) {
  // Trigger-ready inbox collector. Run this as the hidden FF mailbox identity.
  var opts = (options && typeof options === 'object') ? options : {};
  var days = Number(opts.days || 1);
  if (!isFinite(days) || days < 1) days = 1;
  if (days > 90) days = 90;
  var maxThreads = Number(opts.maxThreads || 50);
  if (!isFinite(maxThreads) || maxThreads < 5) maxThreads = 50;
  if (maxThreads > 200) maxThreads = 200;

  var syncStartTime = new Date();
  var primaryInbox = String(_flightFollowPrimaryInbox_() || '').trim().toLowerCase();
  var identityGate = _flightFollowSyncIdentityGate_(primaryInbox);
  if (!identityGate.allowed) {
    _flightFollowRecordSyncMetadata_({
      syncTime: new Date().toISOString(),
      scanned: 0,
      inserted: 0,
      threadCount: 0,
      durationMs: 0,
      query: '',
      primaryInbox: primaryInbox,
      ffInboxCount: 0,
      outboundCount: 0,
      executor: identityGate.effectiveUser,
      blocked: true,
      blockReason: 'SYNC_EXECUTOR_MISMATCH'
    });
    return {
      success: false,
      blocked: true,
      error: 'Flight Follow sync is restricted to the primary inbox account: ' + primaryInbox,
      requiredInbox: primaryInbox,
      effectiveUser: identityGate.effectiveUser,
      activeUser: identityGate.activeUser
    };
  }
  var defaultQuery = [
    'in:inbox',
    'newer_than:' + String(Math.floor(days)) + 'd'
  ].join(' ');
  var query = String(opts.query || defaultQuery).trim();
  var sheet = _flightFollowMessagesSheet_();
  var ss = sheet.getParent();
  var sheetRowsBefore = Math.max(0, sheet.getLastRow() - 1);
  var ffInboxes = _flightFollowRecipientsList_().map(function(v) { return String(v || '').toLowerCase(); });
  var outboundAddrs = _flightFollowOutboundRecipientList_();
  var existingIds = {};
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var existingVals = sheet.getRange(2, 11, lastRow - 1, 1).getValues();
    existingVals.forEach(function(r) {
      var id = String((r && r[0]) || '').trim();
      if (id) existingIds[id] = true;
    });
  }

  var threads = GmailApp.search(query, 0, maxThreads);
  var rowsToAppend = [];
  var scanned = 0;

  // Batch-fetch all messages in one API call to minimize quota usage.
  var allThreadMessages = GmailApp.getMessagesForThreads(threads);
  allThreadMessages.forEach(function(msgs, tIdx) {
    var threadId = String(threads[tIdx].getId() || '');
    msgs.forEach(function(msg) {
      // Note: isInInbox() can return false for BCC/filter-delivered messages
      // even when the in:inbox search found them. Rely on the query instead.
      scanned++;
      var messageId = String(msg.getId() || '').trim();
      if (!messageId || existingIds[messageId]) return;

      var from = String(msg.getFrom() || '');
      var to = String(msg.getTo() || '');
      var cc = String(msg.getCc() || '');
      var subj = String(msg.getSubject() || '');
      var body = String(msg.getPlainBody() || '');

      // Pass empty string for primaryInbox so the hard gate is skipped —
      // the Gmail search already constrains to this account's inbox.
      if (!_flightFollowIsSyncRelevantMessage_(from, to, cc, subj, body, '', ffInboxes, outboundAddrs)) return;

      var regNorm = _flightFollowInferRegFromText_([subj, body, from, to, cc].join('\n'));

      rowsToAppend.push({
        timestamp: msg.getDate().toISOString(),
        regNorm: regNorm,
        direction: 'INBOUND',
        from: from,
        to: [to, cc].filter(Boolean).join(', '),
        subject: subj,
        text: body,
        verified: _flightFollowLooksInreach_(from),
        threadId: threadId,
        messageId: messageId,
        source: 'FF_SYNC'
      });
      existingIds[messageId] = true;
    });
  });

  var insertStartRow = sheet.getLastRow() + 1;
  var inserted = _flightFollowAppendMessages_(rowsToAppend);
  var insertEndRow = inserted > 0 ? (insertStartRow + inserted - 1) : 0;
  var sheetRowsAfter = Math.max(0, sheet.getLastRow() - 1);
  var syncEndTime = new Date();
  var durationMs = syncEndTime.getTime() - syncStartTime.getTime();

  Logger.log('FF Sync complete — executor: %s | spreadsheetId: %s | sheet: %s | rows(before->after): %s->%s | appendRows: %s-%s | threads: %s | scanned: %s | passed filter: %s | inserted: %s | query: %s',
    Session.getEffectiveUser().getEmail(), ss.getId(), sheet.getName(), sheetRowsBefore, sheetRowsAfter, insertStartRow, insertEndRow,
    threads.length, scanned, rowsToAppend.length, inserted, query);

  _flightFollowRecordSyncMetadata_({
    syncTime: syncEndTime.toISOString(),
    scanned: scanned,
    inserted: inserted,
    threadCount: threads.length,
    durationMs: durationMs,
    query: query,
    primaryInbox: primaryInbox,
    ffInboxCount: ffInboxes.length,
    outboundCount: outboundAddrs.length,
    executor: Session.getEffectiveUser().getEmail()
  });

  return {
    success: true,
    query: query,
    scannedMessages: scanned,
    insertedMessages: inserted,
    threadCount: threads.length,
    sheetName: sheet.getName(),
    sheetRowsBefore: sheetRowsBefore,
    sheetRowsAfter: sheetRowsAfter,
    appendStartRow: insertStartRow,
    appendEndRow: insertEndRow,
    durationMs: durationMs,
    effectiveUser: identityGate.effectiveUser,
    activeUser: identityGate.activeUser,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName()
  };
}


function _flightFollowRecordSyncMetadata_(metadata) {
  // Store last sync time and stats in config sheet for UI display.
  try {
    var ss = _flightFollowSpreadsheet_();
    var sheet = ss.getSheetByName(APP_SHEETS.SCHED_CONFIG || 'SCHED_Config');
    if (!sheet) return;

    var data = sheet.getDataRange().getValues();
    var keyColIdx = -1;
    var valColIdx = -1;
    if (data.length > 0) {
      var headerRow = data[0];
      for (var i = 0; i < headerRow.length; i++) {
        var h = String(headerRow[i] || '').trim().toUpperCase();
        if (h === 'KEY') keyColIdx = i;
        if (h === 'VALUE') valColIdx = i;
      }
    }
    if (keyColIdx < 0 || valColIdx < 0) return;

    var metaKey = 'FF_SYNC_META';
    var metaValue = JSON.stringify(metadata);
    var found = false;
    for (var ri = 1; ri < data.length; ri++) {
      if (String(data[ri][keyColIdx] || '').trim() === metaKey) {
        sheet.getRange(ri + 1, valColIdx + 1).setValue(metaValue);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([metaKey, metaValue]);
    }
  } catch (e) {}
}

function getFlightFollowSyncMetadata() {
  // Return last sync status for UI display.
  try {
    var ss = _flightFollowSpreadsheet_();
    var sheet = ss.getSheetByName(APP_SHEETS.SCHED_CONFIG || 'SCHED_Config');
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();
    var keyColIdx = -1;
    var valColIdx = -1;
    if (data.length > 0) {
      var headerRow = data[0];
      for (var i = 0; i < headerRow.length; i++) {
        var h = String(headerRow[i] || '').trim().toUpperCase();
        if (h === 'KEY') keyColIdx = i;
        if (h === 'VALUE') valColIdx = i;
      }
    }
    if (keyColIdx < 0 || valColIdx < 0) return null;

    for (var ri = 1; ri < data.length; ri++) {
      if (String(data[ri][keyColIdx] || '').trim() === 'FF_SYNC_META') {
        var raw = String(data[ri][valColIdx] || '');
        try {
          return JSON.parse(raw);
        } catch (e) {
          return null;
        }
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}

function runFlightFollowManualSync(options) {
  // Manual sync entrypoint for UI button.
  var opts = (options && typeof options === 'object') ? options : {};
  if (!Object.prototype.hasOwnProperty.call(opts, 'maxThreads')) opts.maxThreads = 50;
  if (!Object.prototype.hasOwnProperty.call(opts, 'days')) opts.days = 1;
  return syncFlightFollowInboxToSheet(opts);
}

function debugFlightFollowSyncHealth(token) {
  // One-shot health report for trigger/account/query/sheet diagnostics.
  var tkn = String(token || 'KPS').trim();
  var out = {
    token: tkn,
    primaryInbox: _flightFollowPrimaryInbox_(),
    recipients: _flightFollowRecipientsList_(),
    effectiveUser: Session.getEffectiveUser().getEmail(),
    activeUser: Session.getActiveUser().getEmail(),
    triggerCount: 0,
    triggerHandlers: [],
    sheetName: APP_SHEETS.FF_MESSAGES || 'DB_FF_Messages',
    sheetExists: false,
    sheetRows: 0,
    lastSyncMeta: null,
    query: '',
    threadCount: 0,
    inboxMessageCount: 0,
    sampleSubjects: [],
    error: ''
  };

  try {
    var triggers = ScriptApp.getProjectTriggers();
    out.triggerCount = triggers.filter(function(tr) {
      return tr.getHandlerFunction && tr.getHandlerFunction() === 'syncFlightFollowInboxToSheet';
    }).length;
    out.triggerHandlers = triggers.map(function(tr) {
      return tr.getHandlerFunction ? tr.getHandlerFunction() : '';
    }).filter(Boolean);
  } catch (e1) {
    out.error = 'trigger check failed: ' + String(e1 && e1.message || e1);
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(out.sheetName);
    out.sheetExists = !!sh;
    out.sheetRows = sh ? Math.max(0, sh.getLastRow() - 1) : 0;
    out.lastSyncMeta = getFlightFollowSyncMetadata();
  } catch (e2) {
    out.error = (out.error ? out.error + ' | ' : '') + 'sheet check failed: ' + String(e2 && e2.message || e2);
  }

  try {
    var variants = [tkn.toUpperCase(), tkn.toUpperCase().replace(/[^A-Z0-9]/g, '')].filter(Boolean);
    var unique = {};
    variants = variants.filter(function(v) {
      if (unique[v]) return false;
      unique[v] = true;
      return true;
    });
    var clauses = variants.map(function(v) { return '"' + v + '"'; });
    out.query = 'in:inbox newer_than:7d (' + clauses.join(' OR ') + ')';
    var threads = GmailApp.search(out.query, 0, 120);
    out.threadCount = threads.length;
    var msgCount = 0;
    var sample = [];
    threads.forEach(function(th) {
      var msgs = th.getMessages();
      msgs.forEach(function(m) {
        if (!m.isInInbox()) return;
        msgCount++;
        if (sample.length < 10) sample.push(String(m.getSubject() || ''));
      });
    });
    out.inboxMessageCount = msgCount;
    out.sampleSubjects = sample;
  } catch (e3) {
    out.error = (out.error ? out.error + ' | ' : '') + 'gmail check failed: ' + String(e3 && e3.message || e3);
  }

  return out;
}


function _syncDispatchMissionStatusForFlightLeg_(flightLegId) {
  var legKey = String(flightLegId || '').trim();
  if (!legKey) return { success: false, reason: 'missing-flight-leg-id' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dispatchSheet = ss.getSheetByName(APP_SHEETS.DISPATCH);
  var logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  if (!dispatchSheet || !logSheet) return { success: false, reason: 'missing-sheet' };

  var dispatchData = dispatchSheet.getDataRange().getValues();
  if (!dispatchData || dispatchData.length < 2) return { success: false, reason: 'dispatch-empty' };

  var missionId = '';
  var missionRows = [];
  var missionLegIdsMap = {};

  for (var i = 1; i < dispatchData.length; i++) {
    var rowLegId = String(dispatchData[i][DISPATCH_COL.FLIGHT_ID] || '').trim();
    if (rowLegId === legKey) {
      missionId = String(dispatchData[i][DISPATCH_COL.MISSION_ID] || '').trim();
      break;
    }
  }
  if (!missionId) return { success: false, reason: 'mission-not-found' };

  for (var j = 1; j < dispatchData.length; j++) {
    var mid = String(dispatchData[j][DISPATCH_COL.MISSION_ID] || '').trim();
    if (mid !== missionId) continue;
    missionRows.push(j + 1);
    var lid = String(dispatchData[j][DISPATCH_COL.FLIGHT_ID] || '').trim();
    if (lid) missionLegIdsMap[lid] = true;
  }

  var missionLegIds = Object.keys(missionLegIdsMap);
  if (!missionLegIds.length) return { success: false, reason: 'mission-has-no-legs', missionId: missionId };

  var logData = logSheet.getDataRange().getValues();
  var completedLegMap = {};
  for (var li = 1; li < logData.length; li++) {
    var lr = logData[li];
    var lid2 = String(lr[LOG_FLIGHT_COL.FLIGHT_ID] || '').trim();
    if (!lid2) continue;
    var landed = lr[LOG_FLIGHT_COL.LANDED];
    var onBlocks = lr[LOG_FLIGHT_COL.ON_BLOCKS];
    var brakesApplied = lr[LOG_FLIGHT_COL.BRAKES_APPLIED];
    if (landed || onBlocks || brakesApplied) completedLegMap[lid2] = true;
  }

  var legsComplete = missionLegIds.filter(function(lid3) { return !!completedLegMap[lid3]; }).length;
  var allLegsComplete = legsComplete > 0 && legsComplete === missionLegIds.length;
  var anyLegComplete = legsComplete > 0;
  var targetStatus = allLegsComplete ? 'FLOWN' : (anyLegComplete ? 'In-Flight' : '');

  var updatedRows = 0;
  if (targetStatus) {
    missionRows.forEach(function(rowNum) {
      var existing = String(dispatchSheet.getRange(rowNum, DISPATCH_COL.STATUS + 1).getValue() || '').trim();
      var existingUpper = existing.toUpperCase();
      if (existingUpper === 'CANCELLED') return;
      if (existingUpper === String(targetStatus).toUpperCase()) return;
      dispatchSheet.getRange(rowNum, DISPATCH_COL.STATUS + 1).setValue(targetStatus);
      updatedRows++;
    });
    if (updatedRows > 0) invalidateScheduledMissionsCache_();
  }

  return {
    success: true,
    missionId: missionId,
    legsComplete: legsComplete,
    legCount: missionLegIds.length,
    targetStatus: targetStatus || '',
    updatedRows: updatedRows
  };
}

function _mxSyncAircraftCvaExpiry_(aircraftReg, dueDate) {
  var reg = String(aircraftReg || '').trim().toUpperCase();
  var due = _mxIsoDate_(dueDate);
  if (!reg || !due) return { success: false, skipped: true };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getRequiredSheet_(ss, APP_SHEETS.AIRCRAFT, '_mxSyncAircraftCvaExpiry_');
  var data = sh.getDataRange().getValues();
  if (!data || data.length < 2) return { success: false, error: 'DB_Aircraft has no rows.' };

  var headers = data[0];
  var regIdx = _cvaHeaderIndexFromCandidates_(headers, ['REGISTRATION', 'REG', 'TAIL', 'TAIL #']);
  var cvaIdx = _cvaHeaderIndexFromCandidates_(headers, ['ANNUAL DUE_(CVA)']);
  if (regIdx < 0) return { success: false, error: 'CVA sync: DB_Aircraft registration column not found.' };

  if (cvaIdx < 0) {
    cvaIdx = headers.length;
    sh.getRange(1, cvaIdx + 1).setValue('ANNUAL DUE_(CVA)');
  }

  var rowIdx = _toolsFindAircraftRowByReg_(data, regIdx, reg);
  if (rowIdx < 0) return { success: false, error: 'CVA sync: aircraft not found in DB_Aircraft: ' + reg };

  sh.getRange(rowIdx + 1, cvaIdx + 1).setValue(due);
  return { success: true, aircraftReg: reg, dueDate: due };
}

function _mxGetCalendarSyncTarget_(cfg, linkKey, legacyIdKey) {
  var map = cfg && typeof cfg === 'object' ? cfg : {};
  var linkValue = String(map[linkKey] || '').trim();
  var idValue = String(map[legacyIdKey] || '').trim();
  return _mxCalendarIdFromLinkOrId_(linkValue || idValue);
}

function _mxSyncTaggedAllDayCalendarEvent_(calendarTarget, tag, title, dateIso, description) {
  try {
    var calendarId = _mxCalendarIdFromLinkOrId_(calendarTarget);
    if (!calendarId) return { success: false, skipped: true, error: 'Calendar link not configured.' };
    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) return { success: false, error: 'Calendar not found for ' + calendarId };
    var ymd = _mxIsoDate_(dateIso);
    if (!ymd) return { success: false, error: 'Invalid calendar date.' };
    var removed = _mxRemoveTaggedCalendarEvents_(calendar, tag);
    var payloadDescription = String(description || '').trim();
    if (payloadDescription.indexOf(tag) < 0) {
      payloadDescription = payloadDescription ? (payloadDescription + '\n' + tag) : tag;
    }
    calendar.createAllDayEvent(String(title || '').trim() || tag, new Date(ymd + 'T00:00:00'), {
      description: payloadDescription
    });
    return { success: true, calendarId: calendarId, removed: removed };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _mxSyncFlightMissionCalendar_(missionId, data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cfgSheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_CONFIG || 'SCHED_Config', '_mxSyncFlightMissionCalendar_');
    var cfg = _schedulerReadConfigMap_(cfgSheet);
    var calendarTarget = _mxGetCalendarSyncTarget_(cfg, 'FLIGHTS_CALENDAR_LINK', 'FLIGHTS_CALENDAR_ID');
    if (!calendarTarget) return { success: true, skipped: true, reason: 'no_flight_calendar_config' };

    var payload = data && typeof data === 'object' ? data : {};
    var deleteOnly = String(payload.action || '').trim().toUpperCase() === 'DELETE' || payload.deleteOnly === true;
    var tag = '[MBA_FLIGHT:' + missionId + ']';
    if (deleteOnly) {
      var deleteCalendar = CalendarApp.getCalendarById(calendarTarget);
      if (!deleteCalendar) return { success: false, error: 'Calendar not found for ' + calendarTarget };
      var removed = _mxRemoveTaggedCalendarEvents_(deleteCalendar, tag);
      return { success: true, calendarId: calendarTarget, removed: removed, deletedOnly: true };
    }

    var missionDate = String(payload.date || '').trim();
    var missionTime = String(payload.time || '08:00').trim();
    var flightDescription = String(payload.flightDescription || '').trim();
    var acft = String(payload.acft || '').trim().toUpperCase();
    var pilot = String(payload.pilot || '').trim();
    var copilot = String(payload.copilot || '').trim();
    var missionType = String(payload.type || '').trim();
    var legs = Array.isArray(payload.legs) ? payload.legs : [];
    var firstLeg = legs.length ? (legs[0] || {}) : {};
    var route = String(firstLeg.route || '').trim() || [String(firstLeg.from || '').trim(), String(firstLeg.to || '').trim()].filter(Boolean).join(' - ');
    var descriptionLines = [
      'Mission ID: ' + missionId,
      'Aircraft: ' + acft,
      'Pilot: ' + pilot,
      copilot ? ('Copilot: ' + copilot) : '',
      missionType ? ('Type: ' + missionType) : '',
      flightDescription ? ('Description: ' + flightDescription) : '',
      route ? ('Route: ' + route) : ''
    ].filter(Boolean);
    var title = acft + (flightDescription ? (' - ' + flightDescription) : (' - ' + missionId));
    return _mxSyncTaggedAllDayCalendarEvent_(calendarTarget, tag, title, missionDate || _mxIsoDate_(new Date()), descriptionLines.join('\n'));
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function _mxSyncMaintenanceAssignmentCalendar_(assignmentId, aircraftReg) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cfgSheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_CONFIG || 'SCHED_Config', '_mxSyncMaintenanceAssignmentCalendar_');
    var cfg = _schedulerReadConfigMap_(cfgSheet);
    var calendarTarget = _mxGetCalendarSyncTarget_(cfg, 'MAINTENANCE_CALENDAR_LINK', 'SCHEDULE_CALENDAR_ID');
    if (!calendarTarget) return { success: true, skipped: true, reason: 'no_maintenance_calendar_config' };

    var maintRes = getMaintenanceScheduleData({ aircraftReg: aircraftReg });
    if (!maintRes || !maintRes.success || !Array.isArray(maintRes.rows)) {
      return { success: false, error: (maintRes && maintRes.error) ? maintRes.error : 'Unable to read maintenance rows.' };
    }
    var row = maintRes.rows.find(function(item) {
      return String(item && item.assignmentId || '').trim() === String(assignmentId || '').trim();
    });
    if (!row) return { success: false, error: 'Maintenance assignment not found for calendar sync.' };
    if (!row.nextDueDate) return { success: false, error: 'Maintenance assignment has no next due date.' };

    var title = String(row.aircraftReg || aircraftReg || '').trim().toUpperCase() + ' - ' + String(row.taskName || row.taskCode || 'Maintenance').trim();
    var description = [
      'Assignment ID: ' + String(row.assignmentId || assignmentId || ''),
      'Aircraft: ' + String(row.aircraftReg || aircraftReg || '').trim().toUpperCase(),
      'Task Code: ' + String(row.taskCode || ''),
      'Task Name: ' + String(row.taskName || ''),
      'Due Mode: ' + String(row.dueMode || ''),
      row.notes ? ('Notes: ' + String(row.notes || '')) : ''
    ].filter(Boolean).join('\n');
    return _mxSyncTaggedAllDayCalendarEvent_(calendarTarget, '[MBA_MAINT:' + assignmentId + ']', title, row.nextDueDate, description);
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
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

function _syncStaffIdentityToAuthorizationTables_(staff) {
  var staffName = String(staff && staff.staffName || '').trim();
  var email = String(staff && staff.email || '').trim().toLowerCase();
  var staffId = String(staff && staff.staffId || '').trim();
  if (!staffName || !email) return;

  var now = safeDateStr(new Date());
  var idToken = String(staffId || email).replace(/[^A-Z0-9]/gi, '').toUpperCase() || String(new Date().getTime());

  // Seed identity row in DB_Pilot_Authorizations (non-authorizing until checkout/training updates real rows)
  var authSheet = _pilotAuthorizationSheet_();
  var authHeaders = _toolsSheetHeaderRow_(authSheet);
  var authData = authSheet.getDataRange().getValues();
  var authMatchRow = -1;
  for (var i = 1; i < authData.length; i++) {
    var rec = _pilotAuthorizationRecordFromRow_(authHeaders, authData[i], i + 1);
    var sameIdentity = (staffId && rec.staffId === staffId) || (email && rec.pilotEmail === email);
    if (!sameIdentity) continue;
    if (String(rec.source || '').trim().toUpperCase() !== 'STAFF_PROFILE_SYNC') continue;
    authMatchRow = i + 1;
    break;
  }

  var authSeedMap = {
    AUTHORIZATION_ID: 'PAUTH_PROFILE_' + idToken,
    PILOT_NAME: staffName,
    PILOT_EMAIL: email,
    STAFF_ID: staffId,
    AIRCRAFT_TYPE: '',
    ROLE: '',
    AUTHORIZATION_TYPE: '',
    STATUS: 'PENDING',
    DATE_AUTHORIZED: '',
    EXPIRY_DATE: '',
    INSTRUCTOR_SIGNOFF: '',
    SOURCE: 'STAFF_PROFILE_SYNC',
    NOTES: 'Auto-created from Add/Edit Staff',
    CREATED_AT: now,
    UPDATED_AT: now
  };

  if (authMatchRow >= 2) {
    var authCurrent = authSheet.getRange(authMatchRow, 1, 1, authHeaders.length).getValues()[0];
    var authMerged = authHeaders.map(function(header, idx) {
      var key = _toolsNormHeader_(header);
      if (key === 'CREATED_AT') return authCurrent[idx] || now;
      return Object.prototype.hasOwnProperty.call(authSeedMap, key) ? authSeedMap[key] : authCurrent[idx];
    });
    authSheet.getRange(authMatchRow, 1, 1, authMerged.length).setValues([authMerged]);
  } else {
    var authRow = authHeaders.map(function(header) {
      var key = _toolsNormHeader_(header);
      return Object.prototype.hasOwnProperty.call(authSeedMap, key) ? authSeedMap[key] : '';
    });
    authSheet.appendRow(authRow);
  }

  // Seed identity row in DB_Checks (non-authorizing placeholder until real runway/waiver rows are added)
  var checksSheet = _pilotRunwayChecksSheet_();
  var checksHeaders = _toolsSheetHeaderRow_(checksSheet);
  var checksData = checksSheet.getDataRange().getValues();
  var checksMatchRow = -1;
  for (var c = 1; c < checksData.length; c++) {
    var chk = _pilotRunwayCheckRecordFromRow_(checksHeaders, checksData[c], c + 1);
    var sameCheckIdentity = (staffId && chk.staffId === staffId) || (email && chk.pilotEmail === email);
    if (!sameCheckIdentity) continue;
    if (String(chk.source || '').trim().toUpperCase() !== 'STAFF_PROFILE_SYNC') continue;
    checksMatchRow = c + 1;
    break;
  }

  var checksSeedMap = {
    CHECK_ID: 'CHK_PROFILE_' + idToken,
    PILOT_NAME: staffName,
    PILOT_EMAIL: email,
    STAFF_ID: staffId,
    ICAO: '',
    RUNWAY_IDENT: '',
    AUTH_SCOPE: '',
    STATUS: 'PENDING',
    DATE_CHECKED: '',
    EXPIRY_DATE: '',
    APPROVED_BY: '',
    SOURCE: 'STAFF_PROFILE_SYNC',
    NOTES: 'Auto-created from Add/Edit Staff',
    CREATED_AT: now,
    UPDATED_AT: now
  };

  if (checksMatchRow >= 2) {
    var checksCurrent = checksSheet.getRange(checksMatchRow, 1, 1, checksHeaders.length).getValues()[0];
    var checksMerged = checksHeaders.map(function(header, idx) {
      var key = _toolsNormHeader_(header);
      if (key === 'CREATED_AT') return checksCurrent[idx] || now;
      return Object.prototype.hasOwnProperty.call(checksSeedMap, key) ? checksSeedMap[key] : checksCurrent[idx];
    });
    checksSheet.getRange(checksMatchRow, 1, 1, checksMerged.length).setValues([checksMerged]);
  } else {
    var checksRow = checksHeaders.map(function(header) {
      var key = _toolsNormHeader_(header);
      return Object.prototype.hasOwnProperty.call(checksSeedMap, key) ? checksSeedMap[key] : '';
    });
    checksSheet.appendRow(checksRow);
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

function _histFlightImportsRootFolder_() {
  var props = PropertiesService.getScriptProperties();
  var existingId = String(props.getProperty('HIST_FLIGHT_IMPORTS_FOLDER_ID') || '').trim();
  if (existingId) {
    try { return DriveApp.getFolderById(existingId); } catch (e) {}
  }
  var folderName = 'MBA_Historical_Flight_Imports';
  var iter = DriveApp.getFoldersByName(folderName);
  var folder = iter.hasNext() ? iter.next() : DriveApp.createFolder(folderName);
  props.setProperty('HIST_FLIGHT_IMPORTS_FOLDER_ID', folder.getId());
  return folder;
}

function _histFlightEnsureImportSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var batchHeaders = [
    'BATCH_ID',
    'SOURCE_FILENAME',
    'SOURCE_FILE_TYPE',
    'SOURCE_FILE_ID',
    'SOURCE_FILE_URL',
    'SOURCE_FILE_HASH',
    'IMPORTED_AT',
    'IMPORTED_BY',
    'SOURCE_LABEL',
    'DOCUMENT_COUNT',
    'ROW_COUNT',
    'STATUS',
    'NOTES'
  ];
  var stagingHeaders = [
    'BATCH_ID',
    'ROW_NO',
    'SOURCE_FLIGHT_ID',
    'FLIGHT_DATE',
    'PILOT',
    'ACFT',
    'FROM_ICAO',
    'TO_ICAO',
    'TOTAL_TIME_H',
    'NUMBER_LDGS',
    'NUMBER_TOUCH_AND_GOS',
    'SOURCE_NOTE',
    'SOURCE_FILE_ID',
    'SOURCE_FILE_URL',
    'SOURCE_FILE_NAME',
    'SOURCE_MIME_TYPE',
    'EXTRACTION_MODEL',
    'EXTRACTION_CONFIDENCE',
    'EXTRACTION_RAW_JSON',
    'FINGERPRINT_STRICT',
    'FINGERPRINT_FUZZY',
    'MATCH_STATUS',
    'MATCH_REASON',
    'MATCHED_FLIGHT_ID',
    'REVIEW_DECISION',
    'REVIEWED_BY',
    'REVIEWED_AT',
    'COMMIT_FLIGHT_ID',
    'COMMIT_ROW_NO',
    'NOTES'
  ];
  var batchRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.HIST_FLIGHT_IMPORT_BATCHES, batchHeaders);
  var stageRes = _schemaEnsureSheetHeaders_(ss, APP_SHEETS.HIST_FLIGHT_STAGING, stagingHeaders);
  return { success: true, batch: batchRes, staging: stageRes };
}

function setupHistoricalFlightImportSchema() {
  try {
    return _histFlightEnsureImportSheets_();
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getHistoricalFlightImportBatches(limit) {
  try {
    _histFlightEnsureImportSheets_();
    var maxRows = Number(limit || 30);
    if (!isFinite(maxRows) || maxRows < 1) maxRows = 30;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var batchSheet = getRequiredSheet_(ss, APP_SHEETS.HIST_FLIGHT_IMPORT_BATCHES, 'getHistoricalFlightImportBatches');
    var data = batchSheet.getDataRange().getValues();
    if (!data || data.length < 2) return { success: true, rows: [] };

    var headers = data[0].map(_histFlightNormHeader_);
    function get(row, key) {
      var idx = headers.indexOf(key);
      return idx >= 0 ? row[idx] : '';
    }

    var rows = data.slice(1).map(function(row) {
      return {
        batchId: String(get(row, 'BATCH_ID') || '').trim(),
        sourceFileName: String(get(row, 'SOURCE_FILENAME') || '').trim(),
        sourceFileUrl: String(get(row, 'SOURCE_FILE_URL') || '').trim(),
        sourceFileType: String(get(row, 'SOURCE_FILE_TYPE') || '').trim(),
        importedAt: get(row, 'IMPORTED_AT'),
        importedBy: String(get(row, 'IMPORTED_BY') || '').trim(),
        sourceLabel: String(get(row, 'SOURCE_LABEL') || '').trim(),
        documentCount: Number(get(row, 'DOCUMENT_COUNT') || 0),
        rowCount: Number(get(row, 'ROW_COUNT') || 0),
        status: String(get(row, 'STATUS') || '').trim(),
        notes: String(get(row, 'NOTES') || '').trim()
      };
    }).filter(function(r) { return r.batchId; });

    rows.sort(function(a, b) {
      var da = String(a.importedAt || '');
      var db = String(b.importedAt || '');
      return da > db ? -1 : da < db ? 1 : 0;
    });

    return { success: true, rows: rows.slice(0, maxRows) };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e), rows: [] };
  }
}

function inspectHistoricalFlightImportFile(payload) {
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
        error: 'PDF parsing is not enabled in this phase. Convert to CSV before staging.'
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
    return {
      success: true,
      fileType: 'csv',
      canStage: true,
      headers: headers,
      suggestedMap: _histFlightSuggestedMap_(headers),
      sampleRows: sampleRows
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function stageHistoricalFlightImport(payload) {
  try {
    var ensureRes = _histFlightEnsureImportSheets_();
    if (!ensureRes || !ensureRes.success) return { success: false, error: 'Could not ensure import sheets' };

    var fileName = String(payload && payload.fileName || '').trim();
    var mimeType = String(payload && payload.mimeType || '').trim().toLowerCase();
    var base64Data = String(payload && payload.base64Data || '').trim();
    var importedBy = String(payload && payload.importedBy || Session.getActiveUser().getEmail() || 'Admin').trim();
    var sourceLabel = String(payload && payload.sourceLabel || 'HISTORICAL_IMPORT').trim();
    var map = payload && payload.columnMap ? payload.columnMap : {};
    if (!fileName || !base64Data) return { success: false, error: 'File data missing' };

    var ext = fileName.split('.').pop().toLowerCase();
    if (ext === 'pdf' || mimeType.indexOf('pdf') >= 0) {
      return { success: false, error: 'PDF parsing is not enabled in this phase. Convert to CSV before staging.' };
    }

    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || 'text/csv', fileName);
    var text = blob.getDataAsString('UTF-8').replace(/^\uFEFF/, '');
    var rows = Utilities.parseCsv(text);
    if (!rows || rows.length < 2) return { success: false, error: 'CSV appears empty or has no data rows' };

    var headers = rows[0].map(function(h) { return String(h || '').trim(); });
    var headerIndex = {};
    headers.forEach(function(header, idx) { headerIndex[_histFlightNormHeader_(header)] = idx; });
    function idxFor(headerName) {
      var idx = headerIndex[_histFlightNormHeader_(headerName)];
      return typeof idx === 'number' ? idx : -1;
    }

    var dateIdx = idxFor(map.date);
    var pilotIdx = idxFor(map.pilot);
    var acftIdx = idxFor(map.acft);
    var fromIdx = idxFor(map.from);
    var toIdx = idxFor(map.to);
    var routeIdx = idxFor(map.route);
    var totalTimeIdx = idxFor(map.totalTime);
    var numberLdgsIdx = idxFor(map.numberLdgs);
    var touchAndGosIdx = idxFor(map.touchAndGos);
    var flightIdIdx = idxFor(map.flightId);
    var sourceNoteIdx = idxFor(map.sourceNote);

    if (dateIdx < 0 || pilotIdx < 0 || acftIdx < 0 || totalTimeIdx < 0) {
      return { success: false, error: 'Column mapping incomplete. Date, Pilot, Aircraft, and Total Time are required.' };
    }

    var batchId = _histFlightBatchId_();
    var rootFolder = _histFlightImportsRootFolder_();
    var subfolder = rootFolder.createFolder(batchId);
    var file = subfolder.createFile(blob);
    var fileHash = _histFlightDigest_(text);

    var existing = _histFlightReadExistingLogKeys_();
    var staging = [];
    var minDate = '';
    var maxDate = '';

    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      var sourceFlightId = flightIdIdx >= 0 ? String(row[flightIdIdx] || '').trim() : '';
      var flightDate = safeDateYmd_(row[dateIdx]);
      var pilot = String(row[pilotIdx] || '').trim();
      var acft = String(row[acftIdx] || '').trim().toUpperCase();
      var fromTo = _histFlightResolveFromTo_(fromIdx >= 0 ? row[fromIdx] : '', toIdx >= 0 ? row[toIdx] : '', routeIdx >= 0 ? row[routeIdx] : '');
      var totalTimeH = _histFlightParseHours_(row[totalTimeIdx]);
      var numberLdgs = numberLdgsIdx >= 0 ? _histFlightParseInt_(row[numberLdgsIdx]) : '';
      var touchAndGos = touchAndGosIdx >= 0 ? _histFlightParseInt_(row[touchAndGosIdx]) : '';
      var sourceNote = sourceNoteIdx >= 0 ? String(row[sourceNoteIdx] || '').trim() : '';

      if (flightDate) {
        if (!minDate || flightDate < minDate) minDate = flightDate;
        if (!maxDate || flightDate > maxDate) maxDate = flightDate;
      }

      var strictKey = _histFlightMakeStrictKey_({
        date: flightDate,
        pilot: pilot,
        acft: acft,
        from: fromTo.from,
        to: fromTo.to,
        totalTimeH: totalTimeH
      });
      var fuzzyKey = _histFlightMakeFuzzyKey_({ date: flightDate, pilot: pilot, acft: acft });

      var matchStatus = 'NEW';
      var matchReason = '';
      var matchedFlightId = '';

      if (!flightDate || !pilot || !acft || totalTimeH === '') {
        matchStatus = 'INVALID';
        matchReason = 'Missing date, pilot, aircraft, or total time';
      } else if (sourceFlightId && existing.ids[sourceFlightId]) {
        matchStatus = 'LIKELY_DUPLICATE';
        matchReason = 'Flight ID already exists in LOG_Flights';
        matchedFlightId = sourceFlightId;
      } else if (existing.strict[strictKey]) {
        matchStatus = 'LIKELY_DUPLICATE';
        matchReason = 'Exact historical fingerprint already exists';
        matchedFlightId = existing.strict[strictKey];
      } else if (existing.fuzzy[fuzzyKey]) {
        matchStatus = 'POSSIBLE_DUPLICATE';
        matchReason = 'Potential duplicate by date/pilot/aircraft';
        matchedFlightId = existing.fuzzy[fuzzyKey];
      }

      staging.push([
        batchId,
        i,
        sourceFlightId,
        flightDate,
        pilot,
        acft,
        fromTo.from,
        fromTo.to,
        totalTimeH,
        numberLdgs,
        touchAndGos,
        sourceNote,
        file.getId(),
        file.getUrl(),
        file.getName(),
        mimeType || 'text/csv',
        'CSV_IMPORT',
        '',
        '',
        strictKey,
        fuzzyKey,
        matchStatus,
        matchReason,
        matchedFlightId,
        '',
        '',
        '',
        '',
        '',
        ''
      ]);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var batchSheet = getRequiredSheet_(ss, APP_SHEETS.HIST_FLIGHT_IMPORT_BATCHES, 'stageHistoricalFlightImport');
    var stageSheet = getRequiredSheet_(ss, APP_SHEETS.HIST_FLIGHT_STAGING, 'stageHistoricalFlightImport');
    batchSheet.appendRow([
      batchId,
      file.getName(),
      'CSV',
      file.getId(),
      file.getUrl(),
      fileHash,
      new Date().toISOString(),
      importedBy,
      sourceLabel,
      1,
      staging.length,
      'UNDER_REVIEW',
      'Date range: ' + (minDate || '?') + ' to ' + (maxDate || '?')
    ]);
    if (staging.length) {
      stageSheet.getRange(stageSheet.getLastRow() + 1, 1, staging.length, staging[0].length).setValues(staging);
    }
    return { success: true, batchId: batchId, rowCount: staging.length, fileUrl: file.getUrl() };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function syncTrainingQualificationsToScheduler(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var dryRun = body.dryRun === true;
    _schedulerAssertPermission_('CAN_EDIT_RULES', 'syncTrainingQualificationsToScheduler');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var qualSh = ss.getSheetByName(APP_SHEETS.SCHED_STAFF_QUALS || 'SCHED_Staff_Qualifications');
    if (!qualSh) return { success: false, error: 'SCHED_Staff_Qualifications sheet not found. Run SETUP SCHEMA first.' };

    var staffSh = ss.getSheetByName(APP_SHEETS.PILOTS || 'DB_Pilots');
    if (!staffSh) return { success: false, error: 'DB_Pilots sheet not found.' };

    // Load staff list using existing record parser
    var pData = staffSh.getDataRange().getValues();
    var pHeaders = pData.length ? pData[0] : [];
    var staffList = [];
    for (var i = 1; i < pData.length; i++) {
      var rec = _toolsStaffRecordFromRow_(pHeaders, pData[i], i + 1);
      if (!rec.email) continue;
      var roles = [rec.primaryRole].concat(rec.staffRoles || []).map(function(r) { return String(r || '').trim().toUpperCase(); }).filter(Boolean);
      staffList.push({ email: rec.email, roles: roles });
    }

    var roleModuleMap = _trainingRoleModuleMap_();
    var evidenceMap = _trainingStaffModuleEvidence_();
    var today = _trainingQualIsoToday_();

    // Load existing qual rows
    var qRows = qualSh.getDataRange().getValues();
    var qH = qRows[0].map(function(c) { return String(c).trim().toUpperCase(); });
    var qEmail = qH.indexOf('STAFF_EMAIL'); var qRole = qH.indexOf('ROLE_CODE');
    var qActive = qH.indexOf('ACTIVE'); var qValid = qH.indexOf('VALID_UNTIL');
    var qSource = qH.indexOf('SOURCE');

    // Build index of existing rows by "email|role"
    var existingIdx = {};
    for (var r = 1; r < qRows.length; r++) {
      var ek = String(qRows[r][qEmail] || '').toLowerCase().trim() + '|' + String(qRows[r][qRole] || '').toUpperCase().trim();
      existingIdx[ek] = r + 1; // 1-based sheet row
    }

    var actions = [];

    for (var s = 0; s < staffList.length; s++) {
      var member = staffList[s];
      for (var ri = 0; ri < member.roles.length; ri++) {
        var role = member.roles[ri];
        var modules = roleModuleMap[role];
        if (!modules) continue; // role not managed by training system
        var eligibility = _trainingRoleEligibilityForStaff_(member.email, role, modules, evidenceMap, today);
        if (!eligibility.managedByTraining) continue;
        var rowKey = member.email + '|' + role;
        var existingRow = existingIdx[rowKey];
        var newActive = eligibility.eligible ? 'TRUE' : 'FALSE';
        var newValidUntil = eligibility.validUntil || '';
        if (existingRow) {
          var curActive = String(qRows[existingRow - 1][qActive] || '').trim().toUpperCase();
          var curValid = String(qRows[existingRow - 1][qValid] || '').trim();
          var curSource = qSource >= 0 ? String(qRows[existingRow - 1][qSource] || '').trim() : '';
          if (curActive === newActive && curValid === newValidUntil) continue;
          actions.push({ action: eligibility.eligible ? 'ACTIVATE' : 'DEACTIVATE', email: member.email, role: role, active: eligibility.eligible, validUntil: newValidUntil, reason: eligibility.reason, sheetRow: existingRow });
          if (!dryRun) {
            if (qActive >= 0) qualSh.getRange(existingRow, qActive + 1).setValue(newActive);
            if (qValid >= 0) qualSh.getRange(existingRow, qValid + 1).setValue(newValidUntil);
            if (qSource >= 0) qualSh.getRange(existingRow, qSource + 1).setValue('AUTO_FROM_TRAINING_SYNC');
          }
        } else {
          if (!eligibility.eligible) continue; // don't create rows for ineligible staff
          actions.push({ action: 'INSERT', email: member.email, role: role, active: true, validUntil: newValidUntil, reason: eligibility.reason });
          if (!dryRun) {
            var newRow = new Array(qH.length).fill('');
            if (qEmail >= 0) newRow[qEmail] = member.email;
            if (qRole >= 0) newRow[qRole] = role;
            if (qActive >= 0) newRow[qActive] = 'TRUE';
            if (qValid >= 0) newRow[qValid] = newValidUntil;
            if (qSource >= 0) newRow[qSource] = 'AUTO_FROM_TRAINING_SYNC';
            qualSh.appendRow(newRow);
          }
        }
      }
    }

    return { success: true, dryRun: dryRun, actions: actions, count: actions.length };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function previewTrainingQualificationSync(payload) {
  var body = (payload && typeof payload === 'object') ? payload : {};
  body.dryRun = true;
  return syncTrainingQualificationsToScheduler(body);
}

function getHistoricalFlightImportQueueStatus(queueId) {
  try {
    var id = String(queueId || '').trim();
    if (!id) return { success: false, error: 'queueId required' };

    var sheet = _histFlightQueueSheet_();
    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return { success: false, error: 'Queue empty' };

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[0] || '').trim() !== id) continue;
      return {
        success: true,
        queueId: id,
        status: String(row[7] || ''),
        batchId: String(row[1] || ''),
        rowCount: Number(row[12] || 0),
        extractionModel: String(row[13] || ''),
        error: String(row[11] || ''),
        sourceFileId: String(row[2] || '')
      };
    }

    return { success: false, error: 'Queue item not found: ' + id };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function processHistoricalFlightImportQueue() {
  try {
    var sheet = _histFlightQueueSheet_();
    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return { success: true, processed: 0 };

    var rowNo = -1;
    for (var i = 1; i < data.length; i++) {
      var st = String(data[i][7] || '').trim().toUpperCase();
      if (st === 'PENDING') {
        rowNo = i + 1;
        break;
      }
    }
    if (rowNo < 0) return { success: true, processed: 0 };

    sheet.getRange(rowNo, 8).setValue('PROCESSING');
    sheet.getRange(rowNo, 10).setValue(new Date().toISOString());

    var row = sheet.getRange(rowNo, 1, 1, 14).getValues()[0];
    var batchId = String(row[1] || '').trim() || _histFlightBatchId_();
    var fileId = String(row[2] || '').trim();
    var fileName = String(row[3] || '').trim();
    var mimeType = String(row[4] || '').trim() || 'application/octet-stream';
    var sourceLabel = String(row[5] || '').trim() || 'HISTORICAL_PAPER_CAPTURE';
    var importedBy = String(row[6] || '').trim() || (Session.getActiveUser().getEmail() || 'Admin');

    if (!fileId) {
      sheet.getRange(rowNo, 8).setValue('ERROR');
      sheet.getRange(rowNo, 11).setValue(new Date().toISOString());
      sheet.getRange(rowNo, 12).setValue('Missing FILE_ID');
      return { success: false, error: 'Missing FILE_ID' };
    }

    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var bytes = blob.getBytes();
    var base64Data = Utilities.base64Encode(bytes);
    var safeName = String(fileName || file.getName() || 'historical_upload').replace(/[^a-zA-Z0-9._-]+/g, '_');

    var extractRes = _histFlightGeminiExtractFromDocument_(base64Data, mimeType, safeName);
    if (!extractRes || !extractRes.success) {
      sheet.getRange(rowNo, 8).setValue('ERROR');
      sheet.getRange(rowNo, 11).setValue(new Date().toISOString());
      sheet.getRange(rowNo, 12).setValue(extractRes && extractRes.error ? String(extractRes.error) : 'Gemini extraction failed');
      return { success: false, error: extractRes && extractRes.error ? extractRes.error : 'Gemini extraction failed' };
    }

    var flights = Array.isArray(extractRes.flights) ? extractRes.flights : [];
    if (!flights.length) {
      sheet.getRange(rowNo, 8).setValue('ERROR');
      sheet.getRange(rowNo, 11).setValue(new Date().toISOString());
      sheet.getRange(rowNo, 12).setValue('No flight rows detected in document.');
      return { success: false, error: 'No flight rows detected in document.' };
    }

    var stageRes = _histFlightStageRecords_({
      batchId: batchId,
      sourceLabel: sourceLabel,
      importedBy: importedBy,
      file: file,
      fileHash: _histFlightDigest_(base64Data),
      fileType: (mimeType.indexOf('pdf') >= 0 ? 'PDF' : 'IMAGE'),
      mimeType: mimeType,
      extractionModel: extractRes.model,
      extractionRawJson: extractRes.rawJson,
      rows: flights
    });

    if (!stageRes || !stageRes.success) {
      sheet.getRange(rowNo, 8).setValue('ERROR');
      sheet.getRange(rowNo, 11).setValue(new Date().toISOString());
      sheet.getRange(rowNo, 12).setValue(stageRes && stageRes.error ? String(stageRes.error) : 'Stage failed');
      return { success: false, error: stageRes && stageRes.error ? stageRes.error : 'Stage failed' };
    }

    sheet.getRange(rowNo, 8).setValue('DONE');
    sheet.getRange(rowNo, 11).setValue(new Date().toISOString());
    sheet.getRange(rowNo, 12).setValue('');
    sheet.getRange(rowNo, 13).setValue(Number(stageRes.rowCount || 0));
    sheet.getRange(rowNo, 14).setValue(String(extractRes.model || ''));

    var hasMorePending = false;
    var refresh = sheet.getDataRange().getValues();
    for (var r = 1; r < refresh.length; r++) {
      if (String(refresh[r][7] || '').trim().toUpperCase() === 'PENDING') {
        hasMorePending = true;
        break;
      }
    }
    if (hasMorePending) _scheduleHistoricalFlightQueueWorker_();

    return { success: true, processed: 1, queueId: String(row[0] || ''), batchId: batchId, rowCount: Number(stageRes.rowCount || 0) };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}
