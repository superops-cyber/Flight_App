function getInclinometerStandaloneHtml() {
  return HtmlService.createHtmlOutputFromFile('inclinometer_hud_standalone').getContent();
}

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

// Sort (exclude completed/cancelled missions from supervisor sidebar)
const missions = Object.values(missionMap).filter(function(m) {
  var s = String(m.status || '').toUpperCase();
  return s !== 'FLOWN' && s !== 'COMPLETE' && s !== 'COMPLETED' && s !== 'CANCELLED';
}).sort((a,b) => {
if (a.status === 'PENDING' && b.status !== 'PENDING') return -1;
if (a.status !== 'PENDING' && b.status === 'PENDING') return 1;
return String(b.date).localeCompare(String(a.date));
});

let instructors = [];
try {
  const pilotsSheet = ss.getSheetByName(APP_SHEETS.PILOTS);
  if (pilotsSheet && pilotsSheet.getLastRow() > 1) {
    const pData = pilotsSheet.getDataRange().getValues();
    const pHeaders = pData[0] || [];
    instructors = pData.slice(1)
      .map(function(row, idx) { return _toolsStaffRecordFromRow_(pHeaders, row, idx + 2); })
      .filter(function(staff) {
        if (!staff || !staff.staffName || !staff.active) return false;
        const roleText = [staff.primaryRole].concat(staff.staffRoles || []).join(' ').toUpperCase();
        return !!staff.canInstruct || roleText.indexOf('INSTRUCTOR') >= 0 || roleText.indexOf('CHECK PILOT') >= 0 || roleText.indexOf('CHECKPILOT') >= 0;
      })
      .map(function(staff) { return String(staff.staffName || '').trim(); })
      .filter(Boolean)
      .sort(function(a, b) { return String(a).localeCompare(String(b)); });
  }
} catch (e) {}

return { user: user, missions: missions, instructors: instructors };
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

function getPendingRunwaySurveyReviews(limit) {
  try {
    console.log('[getPendingRunwaySurveyReviews] called with limit=', limit);
    const schema = _ensureRunwayWalkthroughLogSchema_();
    const sh = schema.sheet;
    const idx = schema.idx;
    const max = Math.max(1, Number(limit || 100) || 100);
    console.log('[getPendingRunwaySurveyReviews] schema ready, max=', max);

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: true, items: [] };
    const rows = sh.getRange(2, 1, lastRow - 1, schema.headers.length).getValues();

    var dbHeaders = [];
    var dbCols = null;
    var dbByRunway = {};
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const dbSheet = ss.getSheetByName('DB_Airports');
      if (dbSheet) {
        const dbData = dbSheet.getDataRange().getValues();
        if (dbData && dbData.length) {
          dbHeaders = dbData[0].map(function(h) { return String(h || '').trim(); });
          dbCols = _runwayDbFindCols_(dbHeaders);
          for (var di = 1; di < dbData.length; di++) {
            if (!dbCols || dbCols.icao < 0 || dbCols.runway < 0) break;
            var rowIcao = String(dbData[di][dbCols.icao] || '').trim().toUpperCase();
            var rowRwy = String(dbData[di][dbCols.runway] || '').trim().toUpperCase();
            if (!rowIcao || !rowRwy) continue;
            dbByRunway[rowIcao + '|' + rowRwy] = _runwayDbRowToObject_(dbHeaders, dbData[di]);
          }
        }
      }
    } catch (dbErr) {}

    const items = rows
      .map(function(r, i) {
        var itemIcao = String(r[idx.ICAO] || '').trim().toUpperCase();
        var itemRwy = String(r[idx.RWY_IDENT] || '').trim().toUpperCase();
        var dbAirport = dbByRunway[itemIcao + '|' + itemRwy] || null;
        var officialLogId = _runwayOfficialLogIdFromDbRow_(dbAirport, String(r[idx.STAGING_ID] || '').trim());
        var approvedAtRaw = String(r[idx.APPROVED_AT] || '').trim();
        var publishedAtRaw = String(r[idx.PUBLISHED_AT] || '').trim();
        var statusRaw = String(r[idx.STATUS] || '').trim().toUpperCase();
        // Compatibility fallback for legacy rows where STATUS was blank/misaligned.
        if (!statusRaw && !approvedAtRaw && !publishedAtRaw) statusRaw = 'PENDING';
        // Slim survey: strip perimeterTrace (GPS points array - can be thousands of pts) and
        // keep only its length as _perimeterTraceCount to avoid exceeding the 10MB Apps Script limit.
        var surveyParsed = _parseJsonLoose_(r[idx.SURVEY_JSON], {});
        var ptCount = Array.isArray(surveyParsed.perimeterTrace) ? surveyParsed.perimeterTrace.length
          : (typeof surveyParsed._perimeterTraceCount === 'number' ? surveyParsed._perimeterTraceCount : 0);
        var surveySlim = {};
        for (var _sk in surveyParsed) { if (_sk !== 'perimeterTrace') surveySlim[_sk] = surveyParsed[_sk]; }
        surveySlim._perimeterTraceCount = ptCount;
        // Slim dbAirport: only pass display-needed fields (KNOWN_FEATURES/FEATURES) rather than the full row.
        var dbAirportSlim = dbAirport
          ? { KNOWN_FEATURES: dbAirport.KNOWN_FEATURES, FEATURES: dbAirport.FEATURES }
          : null;
        return {
          rowNum: i + 2,
          stagingId: String(r[idx.STAGING_ID] || '').trim(),
          icao: itemIcao,
          rwyIdent: itemRwy,
          pilotName: String(r[idx.PILOT_NAME] || '').trim(),
          pilotEmail: String(r[idx.PILOT_EMAIL] || '').trim(),
          walkDate: String(r[idx.WALK_DATE] || '').trim(),
          notes: String(r[idx.NOTES] || '').trim(),
          status: statusRaw,
          approvedAt: approvedAtRaw,
          publishedAt: publishedAtRaw,
          entryKind: String(r[idx.ENTRY_KIND] || '').trim().toUpperCase(),
          survey: surveySlim,
          official: _parseJsonLoose_(r[idx.OFFICIAL_JSON], {}),
          officialLogId: officialLogId,
          officialSourceLocked: !!dbAirport,
          dbAirport: dbAirportSlim,
          captureSummary: _parseJsonLoose_(r[idx.CAPTURE_SUMMARY_JSON], {}),
          deviceInfo: _parseJsonLoose_(r[idx.DEVICE_INFO_JSON], {})
        };
      })
      .filter(function(item) {
        const s = String(item.status || '').trim().toUpperCase();
        if (s === 'PENDING' || s === 'QUEUED' || s === 'SUBMITTED' || s === 'IN_REVIEW' || s === 'UNDER_REVIEW' || s === 'REVIEW') {
          return true;
        }
        if (!s || s === 'NEW') {
          // Keep unprocessed legacy rows visible if they were never approved/published.
          return !item.approvedAt && !item.publishedAt && !!item.stagingId;
        }
        return false;
      })
      .sort(function(a, b) {
        return String(b.walkDate || '').localeCompare(String(a.walkDate || ''));
      })
      .slice(0, max);

    const response = { success: true, items: items };
    try {
      const responseJson = JSON.stringify(response);
      const responseBytes = responseJson ? responseJson.length : 0;
      console.log('[getPendingRunwaySurveyReviews] returning success with', items.length, 'items; approxBytes=', responseBytes);
    } catch (sizeErr) {
      console.log('[getPendingRunwaySurveyReviews] returning success with', items.length, 'items; sizeLogError=', sizeErr && sizeErr.message ? sizeErr.message : String(sizeErr));
    }
    return response;
  } catch (e) {
    console.error('[getPendingRunwaySurveyReviews] CAUGHT ERROR:', e && e.message ? e.message : String(e), 'FULL:', e);
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function approveRunwaySurveyReview(stagingId, supervisorName, supervisorNotes, approvalPassword, mtowInput) {
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
    const official = _parseJsonLoose_(rows[logRow][idx.OFFICIAL_JSON], {});
    const mtowPayload = (mtowInput && typeof mtowInput === 'object') ? mtowInput : {};
    const nowIso = new Date().toISOString();

    if (entryKind === 'RUNWAY_APPROVAL') {
      const snapshot = (survey && typeof survey === 'object' && survey.runwaySnapshot && typeof survey.runwaySnapshot === 'object')
        ? survey.runwaySnapshot
        : {};
      const cutdownLabelRaw = String(
        mtowPayload.cutdownAreaLabel != null ? mtowPayload.cutdownAreaLabel
          : (snapshot.cutdownAreaLabel != null ? snapshot.cutdownAreaLabel
            : (official && official.cutdownAreaLabel != null ? official.cutdownAreaLabel : ''))
      ).trim();
      const cutdownSemRestricao = !!(mtowPayload.cutdownSemRestricao || snapshot.cutdownSemRestricao)
        || /^SR\b/i.test(cutdownLabelRaw)
        || /SEM\s*RESTR/i.test(cutdownLabelRaw);
      const cutdownCandidateRaw = Number(
        mtowPayload.cutdownAreaM != null ? mtowPayload.cutdownAreaM
          : (snapshot.cutdownAreaM != null ? snapshot.cutdownAreaM
            : (snapshot.cutdownArea != null ? snapshot.cutdownArea
              : (official && official.cutdownAreaM != null ? official.cutdownAreaM : NaN)))
      );
      if (!cutdownSemRestricao && (!isFinite(cutdownCandidateRaw) || cutdownCandidateRaw <= 0)) {
        return { success: false, error: 'Runway approval cannot be published without a valid cutdown area.' };
      }
      const cutdownApprovedM = cutdownSemRestricao ? 0 : Math.round(cutdownCandidateRaw);
      const cutdownApprovedLabel = cutdownSemRestricao ? 'SR - Sem restricao' : (String(cutdownApprovedM) + ' m');
      const mtowModelKeyRaw = String(mtowPayload.modelKey || mtowPayload.model || snapshot.mtowModelKey || '').trim().toUpperCase();
      const mtowModelKey = mtowModelKeyRaw || 'GENERIC';
      const mtowKg = Number(mtowPayload.mtowKg || mtowPayload.value || snapshot.maxTakeoffWeight || 0);
      if (!isFinite(mtowKg) || mtowKg <= 0) {
        return { success: false, error: 'Supervisor MTOW is required to approve this runway release.' };
      }

      const mtowByModel = Object.assign({}, snapshot.mtowByModel || {});
      mtowByModel[mtowModelKey] = mtowKg;

      survey.runwaySnapshot = Object.assign({}, snapshot, {
        maxTakeoffWeight: mtowKg,
        mtowModelKey: mtowModelKey,
        mtowByModel: mtowByModel,
        supervisorMtowKg: mtowKg,
        cutdownAreaM: cutdownApprovedM,
        cutdownAreaLabel: cutdownApprovedLabel,
        cutdownSemRestricao: !!cutdownSemRestricao,
        mtowApprovedBy: String(supervisorName || '').trim() || 'Supervisor',
        mtowApprovedAt: nowIso,
        officialLogId: id
      });

      const officialNext = Object.assign({}, official, {
        maxTakeoffWeight: mtowKg,
        mtowModelKey: mtowModelKey,
        mtowByModel: mtowByModel,
        cutdownAreaM: cutdownApprovedM,
        cutdownAreaLabel: cutdownApprovedLabel,
        cutdownSemRestricao: !!cutdownSemRestricao,
        officialLogId: id,
        source: 'DB_Airports',
        sourceLocked: true,
        capturedAt: nowIso
      });

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const dbSheet = ss.getSheetByName('DB_Airports');
      if (!dbSheet) return { success: false, error: 'DB_Airports sheet not found' };

      const dbData = dbSheet.getDataRange().getValues();
      if (!dbData.length) return { success: false, error: 'DB_Airports is empty' };
      const cols = _runwayDbFindCols_(dbData[0]);
      if (cols.icao < 0 || cols.runway < 0 || cols.knownFeatures < 0) {
        return { success: false, error: 'DB_Airports missing ICAO/RWY_IDENT/KNOWN_FEATURES columns' };
      }

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
      if (!publishRows.length) return { success: false, error: 'Runway not found in DB_Airports' };

      publishRows.forEach(function(rowIndex) {
        const rowRwyIdent = String(dbData[rowIndex][cols.runway] || '').trim().toUpperCase();
        const officialSnapshot = _dbAirportOfficialSnapshot_(dbData[rowIndex], cols);
        const existingRaw = String(dbData[rowIndex][cols.knownFeatures] || '').trim();
        const existingObj = _parseJsonLoose_(existingRaw, {});
        const normalizedExisting = Array.isArray(existingObj) ? { features: existingObj } : (existingObj || {});
        const cutdownLabelRaw = String(snapshot.cutdownAreaLabel || '').trim();
        const cutdownSemRestricaoRow = !!snapshot.cutdownSemRestricao
          || /^SR\b/i.test(cutdownLabelRaw)
          || /SEM\s*RESTR/i.test(cutdownLabelRaw);
        const cutdownLabelParsed = (function(label) {
          var m = String(label || '').match(/-?\d+(?:[\.,]\d+)?/);
          if (!m) return null;
          var n = Number(String(m[0]).replace(',', '.'));
          return isFinite(n) ? n : null;
        })(cutdownLabelRaw);
        const cutdownAreaMRaw = Number(
          snapshot.cutdownAreaM != null ? snapshot.cutdownAreaM
            : (snapshot.cutdownArea != null ? snapshot.cutdownArea
              : (cutdownLabelParsed != null ? cutdownLabelParsed : 0))
        );
        const cutdownAreaM = cutdownSemRestricaoRow
          ? 0
          : ((isFinite(cutdownAreaMRaw) && cutdownAreaMRaw >= 0) ? Math.round(cutdownAreaMRaw) : null);
        const cutdownAreaLabelOut = cutdownSemRestricaoRow
          ? 'SR - Sem restricao'
          : (cutdownAreaM == null ? 'Unknown' : String(cutdownAreaM) + ' m');

        const verifiedOperational = Object.assign({}, (normalizedExisting.verifiedOperational || {}), {
          lengthM: Number(snapshot.runwayLengthM || 0) || officialSnapshot.lengthM || 0,
          widthM: Number(snapshot.runwayWidthM || 0) || officialSnapshot.widthM || 0,
          surface: String(snapshot.surface || '').trim() || officialSnapshot.surface || '',
          slopeFromThreshold: String(snapshot.runwayIdent || rowRwyIdent || rwyIdent).trim() || rowRwyIdent,
          cutdownAreaM: cutdownAreaM,
          cutdownAreaLabel: cutdownAreaLabelOut,
          cutdownSemRestricao: !!cutdownSemRestricaoRow,
          pilotNotes: String(snapshot.pilotNotes || rows[logRow][idx.NOTES] || '').trim(),
          approvedBy: String(supervisorName || '').trim() || 'Supervisor',
          approvedAt: nowIso,
          sourceKind: 'RUNWAY_APPROVAL'
        });
        const _rcFromApproval = String(mtowPayload.runwayClass || '').trim();
        if (_rcFromApproval) verifiedOperational.runwayClass = _rcFromApproval;

        const verifiedSurvey = {
          version: 2,
          capturedAt: String(rows[logRow][idx.WALK_DATE] || nowIso),
          pilotName: String(rows[logRow][idx.PILOT_NAME] || '').trim(),
          pilotEmail: String(rows[logRow][idx.PILOT_EMAIL] || '').trim(),
          status: 'APPROVED',
          stagingId: id,
          publishedRunway: rowRwyIdent,
          sourceRunway: rwyIdent,
          entryKind: 'RUNWAY_APPROVAL'
        };

        const merged = Object.assign({}, normalizedExisting, {
          verifiedOperational: verifiedOperational,
          cutdownAreaM: cutdownAreaM,
          cutdownAreaLabel: cutdownAreaLabelOut,
          cutdownSemRestricao: !!cutdownSemRestricaoRow,
          verifiedSurvey: verifiedSurvey,
          officialReference: {
            lengthM: officialSnapshot.lengthM,
            widthM: officialSnapshot.widthM,
            surface: officialSnapshot.surface,
            headingDeg: officialSnapshot.headingDeg,
            source: 'DB_Airports',
            sourceLocked: true,
            logId: id,
            capturedAt: nowIso
          },
          updatedAt: nowIso
        });

        const versionId = String(id + '__' + rowRwyIdent + '__' + nowIso).replace(/[^A-Z0-9_:\-.T]/ig, '_');
        const historyEntry = {
          versionId: versionId,
          stagingId: id,
          status: 'APPROVED',
          publishedAt: nowIso,
          capturedAt: String(rows[logRow][idx.WALK_DATE] || nowIso),
          pilotName: String(rows[logRow][idx.PILOT_NAME] || '').trim(),
          pilotEmail: String(rows[logRow][idx.PILOT_EMAIL] || '').trim(),
          approvedBy: String(supervisorName || '').trim() || 'Supervisor',
          supervisorNotes: String(supervisorNotes || '').trim(),
          publishedRunway: rowRwyIdent,
          sourceRunway: rwyIdent,
          entryKind: 'RUNWAY_APPROVAL',
          runwaySnapshot: snapshot,
          verifiedSurvey: verifiedSurvey,
          verifiedOperational: verifiedOperational,
          officialReference: merged.officialReference
        };

        const existingHistory = Array.isArray(normalizedExisting.surveyHistory)
          ? normalizedExisting.surveyHistory.filter(function(entry) { return entry && typeof entry === 'object'; })
          : [];
        merged.surveyHistory = existingHistory.concat([historyEntry]).slice(-20);
        merged.currentSurveyVersion = {
          versionId: versionId,
          stagingId: id,
          publishedAt: nowIso,
          publishedRunway: rowRwyIdent,
          sourceRunway: rwyIdent
        };

        dbSheet.getRange(rowIndex + 1, cols.knownFeatures + 1).setValue(JSON.stringify(merged));
      });

      logSheet.getRange(logRow + 1, idx.SURVEY_JSON + 1).setValue(JSON.stringify(survey));
      logSheet.getRange(logRow + 1, idx.OFFICIAL_JSON + 1).setValue(JSON.stringify(officialNext));
      logSheet.getRange(logRow + 1, idx.STATUS + 1).setValue('PUBLISHED');
      logSheet.getRange(logRow + 1, idx.SUPERVISOR_NAME + 1).setValue(String(supervisorName || '').trim() || 'Supervisor');
      logSheet.getRange(logRow + 1, idx.SUPERVISOR_NOTES + 1).setValue(String(supervisorNotes || '').trim());
      logSheet.getRange(logRow + 1, idx.APPROVED_AT + 1).setValue(nowIso);
      logSheet.getRange(logRow + 1, idx.PUBLISHED_AT + 1).setValue(nowIso);
      SpreadsheetApp.flush();
      return {
        success: true,
        message: 'Runway approval request approved',
        icao: icao,
        rwyIdent: rwyIdent,
        mtowKg: mtowKg,
        mtowModelKey: mtowModelKey,
        headingsUpdated: publishRows.length
      };
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
      const _rcGps = String(mtowPayload.runwayClass || '').trim();
      if (_rcGps) verifiedOperational.runwayClass = _rcGps;

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
          sourceLocked: true,
          logId: id,
          capturedAt: nowIso
        },
        updatedAt: nowIso
      });

      const versionId = String(id + '__' + rowRwyIdent + '__' + nowIso).replace(/[^A-Z0-9_:\-.T]/ig, '_');
      const historyEntry = {
        versionId: versionId,
        stagingId: id,
        status: 'APPROVED',
        publishedAt: nowIso,
        capturedAt: String(rows[logRow][idx.WALK_DATE] || nowIso),
        pilotName: String(rows[logRow][idx.PILOT_NAME] || '').trim(),
        pilotEmail: String(rows[logRow][idx.PILOT_EMAIL] || '').trim(),
        approvedBy: String(supervisorName || '').trim() || 'Supervisor',
        supervisorNotes: String(supervisorNotes || '').trim(),
        publishedRunway: rowRwyIdent,
        sourceRunway: rwyIdent,
        verifiedSurvey: merged.verifiedSurvey,
        verifiedOperational: verifiedOperational,
        officialReference: merged.officialReference
      };

      const existingHistory = Array.isArray(normalizedExisting.surveyHistory)
        ? normalizedExisting.surveyHistory.filter(function(entry) { return entry && typeof entry === 'object'; })
        : [];
      const cappedHistory = existingHistory.concat([historyEntry]).slice(-20);
      merged.surveyHistory = cappedHistory;
      merged.currentSurveyVersion = {
        versionId: versionId,
        stagingId: id,
        publishedAt: nowIso,
        publishedRunway: rowRwyIdent,
        sourceRunway: rwyIdent
      };

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
