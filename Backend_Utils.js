function getRequiredSheet_(ss, sheetName, contextLabel) {
const sheet = ss.getSheetByName(sheetName);
if (!sheet) {
  throw new Error((contextLabel || 'Operation') + ': missing required sheet "' + sheetName + '"');
}
return sheet;
}

function safeDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'GMT', 'yyyy-MM-dd');
  }
  var raw = String(val).trim();
  if (!raw) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  var br = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (br) {
    var dd = Number(br[1]);
    var mm = Number(br[2]);
    var yyyy = Number(br[3]);
    var dt = new Date(yyyy, mm - 1, dd);
    if (!isNaN(dt.getTime()) && dt.getFullYear() === yyyy && (dt.getMonth() + 1) === mm && dt.getDate() === dd) {
      return Utilities.formatDate(dt, 'GMT', 'yyyy-MM-dd');
    }
  }
  if (/^\d{4}-\d{2}-\d{2}T/.test(raw)) return raw.slice(0, 10);
  var parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, 'GMT', 'yyyy-MM-dd');
  }
  return '';
}

function _passengerToIsoDate_(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'GMT', 'yyyy-MM-dd');
  }
  var raw = String(value).trim();
  if (!raw) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  if (/^\d{4}-\d{2}-\d{2}T/.test(raw)) return raw.slice(0, 10);
  var slash = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slash) {
    var dd = String(parseInt(slash[1], 10)).padStart(2, '0');
    var mm = String(parseInt(slash[2], 10)).padStart(2, '0');
    var yyyy = slash[3];
    return yyyy + '-' + mm + '-' + dd;
  }
  var parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, 'GMT', 'yyyy-MM-dd');
  }
  return '';
}

function _passengerToDate_(value) {
  var iso = _passengerToIsoDate_(value);
  if (!iso) return null;
  var parts = iso.split('-').map(function(p) { return parseInt(p, 10); });
  if (parts.length !== 3) return null;
  var dt = new Date(parts[0], parts[1] - 1, parts[2]);
  return isNaN(dt.getTime()) ? null : dt;
}

function _passengerConsentCurrent_(value) {
  var acceptedAt = _passengerToDate_(value);
  if (!acceptedAt) return false;
  var now = new Date();
  var cutoff = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 365);
  return acceptedAt.getTime() >= cutoff.getTime();
}

function _airportNormalizeMtowKey_(raw) {
  return String(raw || '')
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function _airportBuildMtowMap_(headers, row, byHeader) {
  var out = {};
  var raw = byHeader('MTOW_BY_MODEL', '');
  if (!raw) return out;
  try {
    var parsed = (typeof raw === 'string') ? JSON.parse(raw) : raw;
    if (!parsed || typeof parsed !== 'object') return out;
    Object.keys(parsed).forEach(function(key) {
      var nKey = _airportNormalizeMtowKey_(key);
      var nVal = Number(parsed[key] || 0);
      if (nKey && isFinite(nVal) && nVal > 0) out[nKey] = nVal;
    });
  } catch (e) {}
  return out;
}

function _airportMapRow_(headers, row) {
  var byHeader = function(name, fallback) {
    var idx = headers.indexOf(name);
    return idx >= 0 ? row[idx] : fallback;
  };
  var byAnyHeader = function(names, fallback) {
    var list = Array.isArray(names) ? names : [names];
    for (var ni = 0; ni < list.length; ni++) {
      var value = byHeader(list[ni], null);
      if (value !== null && value !== undefined && String(value).trim() !== '') return value;
    }
    return fallback;
  };

  var mtowByModel = _airportBuildMtowMap_(headers, row, byHeader);
  var mtowValues = Object.keys(mtowByModel).map(function(k) { return Number(mtowByModel[k] || 0); }).filter(function(v) {
    return isFinite(v) && v > 0;
  });

  return {
    icao: byHeader('ICAO', ''),
    nome: byHeader('NOME', ''),
    lat: byHeader('LATITUDE', ''),
    lon: byHeader('LONGITUDE', ''),
    fuelAvailable: byHeader('FUEL_AVAILABLE', ''),
    mtowByModel: mtowByModel,
    maxMtow: mtowValues.length ? Math.max.apply(null, mtowValues) : 0,
    pilotNotes: String(byHeader('PILOT_NOTES', '') || ''),
    airstripPhoto: String(byHeader('AIRSTRIP_PHOTO', '') || ''),
    runwayIdent: byHeader('RWY_IDENT', byHeader('RWY', byHeader('RUNWAY', byHeader('RUNWAY_DESIGNATOR', '')))),
    runwayHeading: byHeader('RUNWAY_HEADING', byHeader('HEADING', '')),
    runwayLength: byHeader('LENGTH_OFFICIAL', byHeader('LENGTH_METERS', byHeader('LENGTH_M', ''))),
    runwayWidth: byHeader('WIDTH_OFFICIAL', byHeader('WIDTH_METERS', byHeader('WIDTH_M', ''))),
    runwaySlopePercent: byHeader('SLOPE_PERCENT', byHeader('SLOPE_PCT', '')),
    runwaySlopeProfile: byHeader('SLOPE_PROFILE', byHeader('RUNWAY_SLOPE_PROFILE', '')),
    elevationFt: byHeader('ELEVATION', byHeader('ALT_FEET', byHeader('ELEVATION_FT', ''))),
    runwayInfrastructureType: byAnyHeader(['INFRASTRUCTURE_TYPE', 'INFRASTRUCTURE', 'INFRA_TYPE', 'INFRA', 'INFRAESTRUTURA'], ''),
    runwaySurfaceActual: byAnyHeader(['SURFACE_ACTUAL', 'RUNWAY_SURFACE_ACTUAL', 'SURFACE_OFFICIAL', 'RUNWAY_SURFACE', 'SURFACE_TYPE', 'SURFACE'], ''),
    runwaySurfaceCondition: byAnyHeader(['SURFACE_CONDITION', 'RUNWAY_SURFACE_CONDITION', 'CONDITION', 'SURFACE_STATUS'], ''),
    chartUrl: byHeader('CHART_URL', byHeader('PLATE_URL', byHeader('APPROACH_CHART', byHeader('PROCEDURE_PDF', byHeader('PDF_URL', ''))))),
    knownFeatures: byHeader('KNOWN_FEATURES', byHeader('FEATURES', ''))
  };
}

function _safeParseJsonObject_(raw) {
  if (raw === null || raw === undefined || raw === '') return {};
  if (typeof raw === 'object') return raw;
  try {
    var parsed = JSON.parse(String(raw));
    return (parsed && typeof parsed === 'object') ? parsed : {};
  } catch (e) {
    return {};
  }
}

function _extractMissionDescriptionFromRaw_(raw) {
  var parsed = _safeParseJsonObject_(raw);
  var direct = String((parsed && parsed.flightDescription) || '').trim();
  if (direct) return direct;
  var topMeta = parsed && parsed.missionMeta && typeof parsed.missionMeta === 'object' ? parsed.missionMeta : {};
  var metaDesc = String(topMeta.flightDescription || '').trim();
  if (metaDesc) return metaDesc;
  var legs = (parsed && Array.isArray(parsed.legs)) ? parsed.legs : [];
  if (legs.length) {
    var firstLeg = legs[0] || {};
    var legDesc = String(firstLeg.flightDescription || '').trim();
    if (legDesc) return legDesc;
    var legMeta = (firstLeg.meta && typeof firstLeg.meta === 'object') ? firstLeg.meta : {};
    var legMetaDesc = String(legMeta.flightDescription || '').trim();
    if (legMetaDesc) return legMetaDesc;
  }
  return '';
}

function _extractMissionCommunicationsFromRaw_(raw) {
  var parsed = _safeParseJsonObject_(raw);
  var collected = [];
  var pushUnique = function(item) {
    if (!item || typeof item !== 'object') return;
    var id = String(item.requestId || item.id || '').trim();
    if (!id) id = String(item.createdAt || '') + '|' + String(item.type || '') + '|' + String(item.legIndex || '');
    if (!id) return;
    if (collected.some(function(existing) { return String(existing.requestId || existing.id || '') === id; })) return;
    var normalized = Object.assign({}, item);
    normalized.requestId = id;
    normalized.type = String(normalized.type || 'SUPERVISOR_NOTE').trim().toUpperCase();
    normalized.status = String(normalized.status || 'OPEN').trim().toUpperCase();
    normalized.message = String(normalized.message || '').trim();
    normalized.createdAt = String(normalized.createdAt || new Date().toISOString());
    var idx = Number(normalized.legIndex);
    normalized.legIndex = isFinite(idx) ? idx : -1;
    normalized.to = String(normalized.to || '').trim().toUpperCase();
    normalized.from = String(normalized.from || '').trim().toUpperCase();
    collected.push(normalized);
  };

  if (parsed && Array.isArray(parsed.communicationRequests)) parsed.communicationRequests.forEach(pushUnique);
  if (parsed && parsed.missionMeta && Array.isArray(parsed.missionMeta.communicationRequests)) parsed.missionMeta.communicationRequests.forEach(pushUnique);
  if (parsed && Array.isArray(parsed.legs)) {
    parsed.legs.forEach(function(leg) {
      if (!leg || typeof leg !== 'object') return;
      if (Array.isArray(leg.communicationRequests)) leg.communicationRequests.forEach(pushUnique);
      if (leg.meta && Array.isArray(leg.meta.communicationRequests)) leg.meta.communicationRequests.forEach(pushUnique);
    });
  }
  return collected;
}

function _extractMissionMxReviewFromRaw_(raw) {
  var parsed = _safeParseJsonObject_(raw);
  var missionMeta = (parsed && parsed.missionMeta && typeof parsed.missionMeta === 'object') ? parsed.missionMeta : {};
  var mxReview = missionMeta.mxReview && typeof missionMeta.mxReview === 'object'
    ? missionMeta.mxReview
    : (parsed && parsed.mxReview && typeof parsed.mxReview === 'object' ? parsed.mxReview : {});
  return mxReview && typeof mxReview === 'object' ? mxReview : {};
}

function _cvaHeaderIndexFromCandidates_(headers, candidates) {
  var list = Array.isArray(headers) ? headers : [];
  var keys = Array.isArray(candidates) ? candidates : [candidates];
  var normHeaders = list.map(function(h) {
    return String(h || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
  });
  for (var i = 0; i < keys.length; i++) {
    var idx = normHeaders.indexOf(String(keys[i] || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, ''));
    if (idx >= 0) return idx;
  }
  return -1;
}

function _cvaParseIsoDate_(raw) {
  var iso = safeDateStr(raw);
  if (!iso) return null;
  var parts = String(iso || '').split('-').map(function(p) { return Number(p); });
  if (parts.length !== 3 || !isFinite(parts[0]) || !isFinite(parts[1]) || !isFinite(parts[2])) return null;
  var dt = new Date(parts[0], parts[1] - 1, parts[2]);
  if (isNaN(dt.getTime())) return null;
  dt.setHours(0, 0, 0, 0);
  return dt;
}

function _cvaDateAddDays_(dateObj, days) {
  var out = new Date(dateObj.getTime());
  out.setDate(out.getDate() + Number(days || 0));
  out.setHours(0, 0, 0, 0);
  return out;
}

function _cvaIsoFromDate_(dateObj) {
  return Utilities.formatDate(dateObj, 'GMT', 'yyyy-MM-dd');
}

function _cvaDateInRange_(dateObj, startDate, endDate) {
  return dateObj.getTime() >= startDate.getTime() && dateObj.getTime() <= endDate.getTime();
}

function _cvaLongestFreeStreak_(startDate, endDate, occupiedIsoMap) {
  var longest = 0;
  var current = 0;
  var cursor = new Date(startDate.getTime());
  while (cursor.getTime() <= endDate.getTime()) {
    var iso = _cvaIsoFromDate_(cursor);
    if (occupiedIsoMap[iso]) {
      current = 0;
    } else {
      current += 1;
      if (current > longest) longest = current;
    }
    cursor = _cvaDateAddDays_(cursor, 1);
  }
  return longest;
}

function _enforceMissionCvaSchedulingPolicy_(ss, missionHeader, missionId, payload) {
  var body = (payload && typeof payload === 'object') ? payload : {};
  var aircraftReg = String(missionHeader && missionHeader.acft || '').trim().toUpperCase();
  var missionDate = _cvaParseIsoDate_(missionHeader && missionHeader.date);
  if (!aircraftReg) throw new Error('CVA scheduling rule: aircraft registration is required.');
  if (!missionDate) throw new Error('CVA scheduling rule: mission date is invalid.');

  var aircraftSheet = getRequiredSheet_(ss, APP_SHEETS.AIRCRAFT, '_enforceMissionCvaSchedulingPolicy_');
  var aircraftData = aircraftSheet.getDataRange().getValues();
  if (!aircraftData || aircraftData.length < 2) {
    throw new Error('CVA scheduling rule: DB_Aircraft is empty.');
  }

  var headers = aircraftData[0] || [];
  var regIdx = _cvaHeaderIndexFromCandidates_(headers, ['REGISTRATION', 'REG', 'TAIL', 'TAIL #']);
  var cvaIdx = _cvaHeaderIndexFromCandidates_(headers, ['ANNUAL DUE_(CVA)']);
  if (regIdx < 0) throw new Error('CVA scheduling rule: DB_Aircraft registration column not found.');
  if (cvaIdx < 0) throw new Error('CVA scheduling rule: DB_Aircraft CVA column not found (Annual Due_(CVA)).');

  var aircraftRow = null;
  for (var i = 1; i < aircraftData.length; i++) {
    var regVal = String(aircraftData[i][regIdx] || '').trim().toUpperCase();
    if (regVal && regVal === aircraftReg) {
      aircraftRow = aircraftData[i];
      break;
    }
  }
  if (!aircraftRow) throw new Error('CVA scheduling rule: aircraft not found in DB_Aircraft: ' + aircraftReg);

  var cvaRaw = aircraftRow[cvaIdx];
  var cvaDate = _cvaParseIsoDate_(cvaRaw);
  if (!cvaDate) {
    throw new Error('CVA scheduling rule: no current CVA expiry date in DB_Aircraft for ' + aircraftReg + '.');
  }
  if (cvaDate.getTime() < missionDate.getTime()) {
    throw new Error('CVA scheduling rule: CVA expiry for ' + aircraftReg + ' is not current (' + _cvaIsoFromDate_(cvaDate) + ').');
  }

  var scanStart = _cvaDateAddDays_(cvaDate, -30);
  var scanEnd = _cvaDateAddDays_(cvaDate, 10);
  var occupied = {};

  var dispatchSheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, '_enforceMissionCvaSchedulingPolicy_');
  var dispatchData = dispatchSheet.getDataRange().getValues();
  for (var r = 1; r < dispatchData.length; r++) {
    var row = dispatchData[r];
    if (String(row[DISPATCH_COL.AIRCRAFT] || '').trim().toUpperCase() !== aircraftReg) continue;
    if (String(row[DISPATCH_COL.MISSION_ID] || '').trim() === String(missionId || '').trim()) continue;
    var status = String(row[DISPATCH_COL.STATUS] || '').trim().toUpperCase();
    if (status === 'CANCELLED') continue;
    var rowDate = _cvaParseIsoDate_(row[DISPATCH_COL.DATE]);
    if (!rowDate) continue;
    if (!_cvaDateInRange_(rowDate, scanStart, scanEnd)) continue;
    occupied[_cvaIsoFromDate_(rowDate)] = true;
  }

  if (_cvaDateInRange_(missionDate, scanStart, scanEnd)) {
    occupied[_cvaIsoFromDate_(missionDate)] = true;
  }

  var minWindowDays = 10;
  var maxFreeStreakDays = _cvaLongestFreeStreak_(scanStart, scanEnd, occupied);
  var requiresSupervisorApproval = maxFreeStreakDays < minWindowDays;

  var supervisorName = String(body.supervisorApprovalName || '').trim();
  var supervisorApprovalAt = '';
  if (requiresSupervisorApproval) {
    var approvalPassword = String(body.supervisorApprovalPassword || '');
    if (!approvalPassword) {
      throw new Error(
        'CVA_SUPERVISOR_APPROVAL_REQUIRED|Aircraft ' + aircraftReg
        + ' scan window ' + _cvaIsoFromDate_(scanStart) + ' to ' + _cvaIsoFromDate_(scanEnd)
        + ' has max free streak ' + maxFreeStreakDays + ' days (minimum required: ' + minWindowDays + ').'
      );
    }
    _verifySupervisorApprovalPassword_(approvalPassword);
    if (!supervisorName) supervisorName = String(Session.getActiveUser().getEmail() || 'Supervisor').trim();
    supervisorApprovalAt = new Date().toISOString();
  }

  return {
    cvaExpiryDate: _cvaIsoFromDate_(cvaDate),
    missionDate: _cvaIsoFromDate_(missionDate),
    scanWindowStart: _cvaIsoFromDate_(scanStart),
    scanWindowEnd: _cvaIsoFromDate_(scanEnd),
    minWindowDays: minWindowDays,
    maxFreeStreakDays: maxFreeStreakDays,
    requiresSupervisorApproval: requiresSupervisorApproval,
    supervisorApprovalUsed: !!requiresSupervisorApproval,
    supervisorApprovedBy: supervisorName,
    supervisorApprovedAt: supervisorApprovalAt,
    supervisorApprovalReason: String(body.supervisorApprovalReason || '').trim()
  };
}

function _supervisorBuildAirportLookup_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('DB_Airports');
  if (!sh) return null;
  var data = sh.getDataRange().getValues();
  if (!data || data.length < 2) return null;
  var cols = _runwayDbFindCols_(data[0]);
  if (cols.icao < 0 || cols.runway < 0 || cols.knownFeatures < 0) return null;
  var byIcao = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var icao = _supervisorNormalizeIcao_(row[cols.icao]);
    if (!icao) continue;
    if (!byIcao[icao]) byIcao[icao] = [];
    byIcao[icao].push(row);
  }
  return { rowsByIcao: byIcao, cols: cols };
}

function _supervisorNormalizeIcao_(value) {
  var raw = String(value || '').trim().toUpperCase();
  var m = raw.match(/[A-Z]{4}/);
  return m ? m[0] : raw;
}

function _supervisorNormalizeRwyIdent_(value) {
  var raw = String(value || '').trim().toUpperCase();
  var m = raw.match(/(\d{1,2}[LCR]?)/);
  if (!m) return raw;
  var num = String(parseInt(m[1], 10));
  if (!num || num === 'NaN') return raw;
  var suffix = m[1].replace(/\d/g, '');
  return num.padStart(2, '0') + suffix;
}

function _supervisorExtractLegRunway_(legPayload, parsedRoute, rawRoute) {
  var candidates = [];
  if (legPayload && typeof legPayload === 'object') {
    candidates.push(legPayload.rwyIdent, legPayload.runway, legPayload.toRunway, legPayload.destRunway, legPayload.destinationRunway);
  }
  if (parsedRoute && parsedRoute.to) candidates.push(parsedRoute.to);
  candidates.push(rawRoute);
  for (var i = 0; i < candidates.length; i++) {
    var txt = String(candidates[i] || '').trim().toUpperCase();
    if (!txt) continue;
    var m = txt.match(/RWY\s*([0-9]{1,2}[LCR]?)/);
    if (m && m[1]) return _supervisorNormalizeRwyIdent_(m[1]);
    if (/^[0-9]{1,2}[LCR]?$/.test(txt)) return _supervisorNormalizeRwyIdent_(txt);
  }
  return '';
}

function _supervisorBuildRunwayGapForLeg_(lookup, parsedRoute, legPayload, rawRoute) {
  var toIcao = _supervisorNormalizeIcao_(parsedRoute && parsedRoute.to);
  var desiredRwy = _supervisorExtractLegRunway_(legPayload, parsedRoute, rawRoute);
  var baseMissing = ['cutdown area', 'obstacle profile', 'slope profile'];
  if (!toIcao || !lookup || !lookup.rowsByIcao) {
    return {
      icao: toIcao || '',
      rwyIdent: desiredRwy || '',
      missingItems: baseMissing.slice(),
      hasInternalData: false,
      needsAttention: true,
      summary: 'No internal runway dataset found in DB_Airports.',
      runway: null
    };
  }

  var rows = lookup.rowsByIcao[toIcao] || [];
  if (!rows.length) {
    return {
      icao: toIcao,
      rwyIdent: desiredRwy || '',
      missingItems: baseMissing.slice(),
      hasInternalData: false,
      needsAttention: true,
      summary: 'Airport not found in DB_Airports.',
      runway: null
    };
  }

  var cols = lookup.cols;
  var selected = rows[0];
  var selectedRwy = _supervisorNormalizeRwyIdent_(selected[cols.runway]);
  if (desiredRwy) {
    for (var i = 0; i < rows.length; i++) {
      var rowRwy = _supervisorNormalizeRwyIdent_(rows[i][cols.runway]);
      if (rowRwy && rowRwy === desiredRwy) {
        selected = rows[i];
        selectedRwy = rowRwy;
        break;
      }
    }
  }
  // If still no runway ident, scan all rows for any row with a non-empty ident.
  if (!selectedRwy) {
    for (var j = 0; j < rows.length; j++) {
      var anyRwy = _supervisorNormalizeRwyIdent_(rows[j][cols.runway]);
      if (anyRwy) { selected = rows[j]; selectedRwy = anyRwy; break; }
    }
  }

  var knownRaw = String(selected[cols.knownFeatures] || '').trim();
  var knownObj = _parseJsonLoose_(knownRaw, {});
  var knownNorm = Array.isArray(knownObj) ? { features: knownObj } : (knownObj || {});
  var verified = (knownNorm.verifiedOperational && typeof knownNorm.verifiedOperational === 'object') ? knownNorm.verifiedOperational : {};
  var briefingCard = (knownNorm.briefingCard && typeof knownNorm.briefingCard === 'object') ? knownNorm.briefingCard : {};

  var features = Array.isArray(verified.features) ? verified.features : (Array.isArray(knownNorm.features) ? knownNorm.features : []);
  var slopeSegs = Array.isArray(verified.slopeSegments) ? verified.slopeSegments
    : (Array.isArray(knownNorm.slopeSegments) ? knownNorm.slopeSegments
      : (Array.isArray(briefingCard.slopeSegments) ? briefingCard.slopeSegments : []));
  var noSignificantSlope = !!(verified.noSignificantSlope || knownNorm.noSignificantSlope || briefingCard.noSignificantSlope);
  var obstacleList = Array.isArray(verified.obstacleAngles50m) ? verified.obstacleAngles50m
    : (Array.isArray(verified.obstacleAngles) ? verified.obstacleAngles
    : (Array.isArray(knownNorm.obstacleAngles50m) ? knownNorm.obstacleAngles50m
    : (Array.isArray(knownNorm.obstacleAngles) ? knownNorm.obstacleAngles : [])));
  var cutdown = Number(verified.cutdownAreaM != null ? verified.cutdownAreaM : knownNorm.cutdownAreaM || 0) || 0;

  var hasCutdown = cutdown > 0;
  var hasObstacles = obstacleList.length > 0;
  var hasSlope = noSignificantSlope || slopeSegs.length > 0;
  var hasInternalData = hasCutdown || hasObstacles || hasSlope || features.length > 0;
  var missing = [];
  if (!hasCutdown) missing.push('cutdown area');
  if (!hasObstacles) missing.push('obstacle profile');
  if (!hasSlope) missing.push('slope profile');

  var needsAttention = !hasCutdown && !hasObstacles && !hasSlope;
  var summary = needsAttention
    ? 'No internal cutdown / obstacle / slope data for this runway.'
    : (missing.length ? ('Missing: ' + missing.join(', ') + '.') : 'Internal runway dataset is populated.');

  var official = _dbAirportOfficialSnapshot_(selected, cols);
  var runwayObj = {
    icao: toIcao,
    rwyIdent: selectedRwy || desiredRwy || '',
    length: Number(selected[cols.length] || 0) || 0,
    width: Number(selected[cols.width] || 0) || 0,
    slope: Number(verified.slopePercent || knownNorm.slopePercent || 0) || 0,
    surface: String(selected[cols.surface] || '').trim(),
    elevation: Number((cols.elevation >= 0 ? selected[cols.elevation] : null) || verified.elevationFt || knownNorm.elevationFt || 0) || 0,
    knownFeatures: Array.isArray(features) ? features : [],
    surveySlopeSegments: Array.isArray(slopeSegs) ? slopeSegs : [],
    obstacleAngles: Array.isArray(obstacleList) ? obstacleList : [],
    officialReference: official,
    verifiedOperational: verified,
    cutdownAreaM: cutdown,
    airstripPhoto: String(knownNorm.airstripPhoto || '').trim(),
    internalUpdatedAt: String(verified.verifiedAt || knownNorm.updatedAt || '').trim(),
    surveyPilot: String(verified.verifiedBy || knownNorm.updatedBy || '').trim(),
    knownFeaturesRaw: knownRaw || JSON.stringify(knownNorm || {})
  };

  return {
    icao: toIcao,
    rwyIdent: runwayObj.rwyIdent,
    missingItems: missing,
    hasInternalData: hasInternalData,
    needsAttention: needsAttention,
    summary: summary,
    runway: runwayObj
  };
}

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

function _getMxApprovalPassword_() {
  const props = PropertiesService.getScriptProperties();
  const keys = [
    'MX_SUPERVISOR_APPROVAL_PASSWORD',
    'MX_APPROVAL_PASSWORD',
    'CHIEF_MECHANIC_APPROVAL_PASSWORD'
  ];
  for (let i = 0; i < keys.length; i++) {
    const value = props.getProperty(keys[i]);
    if (value && String(value).trim()) return String(value);
  }
  return '';
}

function _financeNormHeader_(value) {
  return String(value || '').trim().toUpperCase().replace(/\s+/g, '_').replace(/[^A-Z0-9_]/g, '');
}

function _financeFindHeaderIndex_(headers, aliases, fallbackIdx) {
  var normalized = (headers || []).map(function(h) { return _financeNormHeader_(h); });
  var list = Array.isArray(aliases) ? aliases : [aliases];
  for (var i = 0; i < list.length; i++) {
    var idx = normalized.indexOf(_financeNormHeader_(list[i]));
    if (idx >= 0) return idx;
  }
  return typeof fallbackIdx === 'number' ? fallbackIdx : -1;
}

function _financeEnsureSheetHeaders_(ss, sheetName, headers) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }
  var existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  var existingNorm = existing.map(_financeNormHeader_);
  (headers || []).forEach(function(header) {
    var label = String(header || '').trim();
    if (!label) return;
    if (existingNorm.indexOf(_financeNormHeader_(label)) >= 0) return;
    sh.getRange(1, sh.getLastColumn() + 1).setValue(label);
    existingNorm.push(_financeNormHeader_(label));
  });
  return sh;
}

function _financeEnsureFundBalanceColumn_(fundSheet) {
  if (!fundSheet) return { nameIdx: -1, balanceIdx: -1 };
  var headers = fundSheet.getRange(1, 1, 1, Math.max(fundSheet.getLastColumn(), 1)).getValues()[0];
  var nameIdx = _financeFindHeaderIndex_(headers, ['NAME', 'FUND_NAME', 'FUND NAME'], 0);
  var balanceIdx = _financeFindHeaderIndex_(headers, ['CURRENT_BALANCE', 'BALANCE'], -1);
  if (balanceIdx < 0) {
    balanceIdx = headers.length;
    fundSheet.getRange(1, balanceIdx + 1).setValue('CURRENT_BALANCE');
  }
  return { nameIdx: nameIdx, balanceIdx: balanceIdx };
}

function _financeEnsureLedgerSheet_(ss) {
  return _financeEnsureSheetHeaders_(ss, APP_SHEETS.FUND_LEDGER, [
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
  ]);
}

function _financeEnsureAllocSheet_(ss) {
  return _financeEnsureSheetHeaders_(ss, APP_SHEETS.MISSION_FUNDING_ALLOCATIONS, [
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
  ]);
}

function _financePostMissionTransactions_(ss, missionId, actorEmail) {
  var transSheet = ss.getSheetByName(APP_SHEETS.TRANSACTIONS);
  if (!transSheet) return { posted: 0, skipped: 0, warnings: ['DB_Transactions not found'] };

  var transData = transSheet.getDataRange().getValues();
  if (!transData || transData.length < 2) return { posted: 0, skipped: 0, warnings: [] };

  var headers = transData[0] || [];
  var flightIdx = _financeFindHeaderIndex_(headers, ['FLIGHT_ID', 'MISSION_ID'], 0);
  var fundIdx = _financeFindHeaderIndex_(headers, ['FUND', 'FUND_NAME'], 1);
  var paxIdx = _financeFindHeaderIndex_(headers, ['PASSENGER_NAME', 'NAME'], 2);
  var amountIdx = _financeFindHeaderIndex_(headers, ['CHARGED_AMOUNT', 'AMOUNT'], 6);
  var statusIdx = _financeFindHeaderIndex_(headers, ['STATUS'], 7);
  if (statusIdx >= headers.length) statusIdx = 7;
  if (headers.length <= statusIdx || _financeNormHeader_(headers[statusIdx]) !== 'STATUS') {
    transSheet.getRange(1, statusIdx + 1).setValue('STATUS');
  }

  var postedAtIdx = _financeFindHeaderIndex_(headers, ['POSTED_AT'], -1);
  if (postedAtIdx < 0) {
    postedAtIdx = Math.max(transSheet.getLastColumn(), statusIdx + 1);
    transSheet.getRange(1, postedAtIdx + 1).setValue('POSTED_AT');
  }

  var fundSheet = ss.getSheetByName(APP_SHEETS.FUNDS);
  if (!fundSheet) return { posted: 0, skipped: 0, warnings: ['DB_Funds not found'] };
  var fundIdxs = _financeEnsureFundBalanceColumn_(fundSheet);
  var fundData = fundSheet.getDataRange().getValues();
  var fundByName = {};
  for (var fr = 1; fr < fundData.length; fr++) {
    var fundName = String(fundData[fr][fundIdxs.nameIdx] || '').trim();
    if (!fundName) continue;
    var bal = parseFloat(fundData[fr][fundIdxs.balanceIdx]);
    if (!isFinite(bal)) bal = 0;
    fundByName[fundName] = { row: fr + 1, balance: bal };
  }

  var ledgerSheet = _financeEnsureLedgerSheet_(ss);
  var allocSheet = _financeEnsureAllocSheet_(ss);
  var now = new Date();
  var nowIso = now.toISOString();
  var actor = String(actorEmail || 'system');

  var posted = 0;
  var skipped = 0;
  var warnings = [];

  for (var i = 1; i < transData.length; i++) {
    var row = transData[i];
    var flightId = String(row[flightIdx] || '').trim();
    if (!flightId || flightId.indexOf(String(missionId)) !== 0) continue;

    var status = String(row[statusIdx] || 'PENDING').trim().toUpperCase();
    if (status === 'POSTED') continue;
    if (status === 'REVERSED') {
      skipped++;
      continue;
    }

    var fundNameRaw = String(row[fundIdx] || '').trim();
    var amount = parseFloat(row[amountIdx]);
    if (!fundNameRaw || !isFinite(amount) || amount <= 0) {
      skipped++;
      warnings.push('Skipped tx row ' + (i + 1) + ' due to missing fund or invalid amount');
      continue;
    }

    var fundRec = fundByName[fundNameRaw];
    if (!fundRec) {
      skipped++;
      warnings.push('Fund not found in DB_Funds: ' + fundNameRaw + ' (tx row ' + (i + 1) + ')');
      continue;
    }

    fundRec.balance = Number((fundRec.balance - amount).toFixed(2));
    fundSheet.getRange(fundRec.row, fundIdxs.balanceIdx + 1).setValue(fundRec.balance);

    transSheet.getRange(i + 1, statusIdx + 1).setValue('POSTED');
    transSheet.getRange(i + 1, postedAtIdx + 1).setValue(nowIso);

    var ledgerId = 'LED_' + Utilities.getUuid().slice(0, 8).toUpperCase();
    var paxName = String(row[paxIdx] || '').trim();
    ledgerSheet.appendRow([
      ledgerId,
      fundNameRaw,
      now,
      'MISSION_OUT',
      -Math.abs(amount),
      flightId,
      'MISSION_TRANSACTION',
      '',
      missionId,
      '',
      (paxName ? ('PAX: ' + paxName + ' | ') : '') + 'Auto-post from transaction',
      nowIso,
      actor
    ]);

    allocSheet.appendRow([
      'ALC_' + Utilities.getUuid().slice(0, 8).toUpperCase(),
      missionId,
      flightId,
      fundNameRaw,
      '',
      Number(amount.toFixed(2)),
      'FINALIZED',
      '',
      nowIso,
      '',
      actor,
      'Auto-posted from DB_Transactions'
    ]);

    posted++;
  }

  return { posted: posted, skipped: skipped, warnings: warnings };
}

function _financeReverseMissionTransactions_(ss, missionId, actorEmail, reason) {
  var transSheet = ss.getSheetByName(APP_SHEETS.TRANSACTIONS);
  if (!transSheet) return { reversed: 0, skipped: 0, warnings: [] };

  var transData = transSheet.getDataRange().getValues();
  if (!transData || transData.length < 2) return { reversed: 0, skipped: 0, warnings: [] };

  var headers = transData[0] || [];
  var flightIdx = _financeFindHeaderIndex_(headers, ['FLIGHT_ID', 'MISSION_ID'], 0);
  var fundIdx = _financeFindHeaderIndex_(headers, ['FUND', 'FUND_NAME'], 1);
  var amountIdx = _financeFindHeaderIndex_(headers, ['CHARGED_AMOUNT', 'AMOUNT'], 6);
  var statusIdx = _financeFindHeaderIndex_(headers, ['STATUS'], 7);
  if (statusIdx >= headers.length) statusIdx = 7;
  if (headers.length <= statusIdx || _financeNormHeader_(headers[statusIdx]) !== 'STATUS') {
    transSheet.getRange(1, statusIdx + 1).setValue('STATUS');
  }

  var reversedAtIdx = _financeFindHeaderIndex_(headers, ['REVERSED_AT'], -1);
  if (reversedAtIdx < 0) {
    reversedAtIdx = Math.max(transSheet.getLastColumn(), statusIdx + 1);
    transSheet.getRange(1, reversedAtIdx + 1).setValue('REVERSED_AT');
  }

  var fundSheet = ss.getSheetByName(APP_SHEETS.FUNDS);
  if (!fundSheet) return { reversed: 0, skipped: 0, warnings: ['DB_Funds not found'] };
  var fundIdxs = _financeEnsureFundBalanceColumn_(fundSheet);
  var fundData = fundSheet.getDataRange().getValues();
  var fundByName = {};
  for (var fr = 1; fr < fundData.length; fr++) {
    var fundName = String(fundData[fr][fundIdxs.nameIdx] || '').trim();
    if (!fundName) continue;
    var bal = parseFloat(fundData[fr][fundIdxs.balanceIdx]);
    if (!isFinite(bal)) bal = 0;
    fundByName[fundName] = { row: fr + 1, balance: bal };
  }

  var ledgerSheet = _financeEnsureLedgerSheet_(ss);
  var now = new Date();
  var nowIso = now.toISOString();
  var actor = String(actorEmail || 'system');
  var noteReason = String(reason || 'MISSION_REVERSE').trim();

  var reversed = 0;
  var skipped = 0;
  var warnings = [];

  for (var i = 1; i < transData.length; i++) {
    var row = transData[i];
    var flightId = String(row[flightIdx] || '').trim();
    if (!flightId || flightId.indexOf(String(missionId)) !== 0) continue;

    var status = String(row[statusIdx] || 'PENDING').trim().toUpperCase();
    if (status !== 'POSTED') continue;

    var fundNameRaw = String(row[fundIdx] || '').trim();
    var amount = parseFloat(row[amountIdx]);
    if (!fundNameRaw || !isFinite(amount) || amount <= 0) {
      skipped++;
      warnings.push('Skipped reversal tx row ' + (i + 1) + ' due to missing fund or invalid amount');
      continue;
    }

    var fundRec = fundByName[fundNameRaw];
    if (!fundRec) {
      skipped++;
      warnings.push('Fund not found in DB_Funds for reversal: ' + fundNameRaw + ' (tx row ' + (i + 1) + ')');
      continue;
    }

    fundRec.balance = Number((fundRec.balance + amount).toFixed(2));
    fundSheet.getRange(fundRec.row, fundIdxs.balanceIdx + 1).setValue(fundRec.balance);

    transSheet.getRange(i + 1, statusIdx + 1).setValue('REVERSED');
    transSheet.getRange(i + 1, reversedAtIdx + 1).setValue(nowIso);

    ledgerSheet.appendRow([
      'LED_' + Utilities.getUuid().slice(0, 8).toUpperCase(),
      fundNameRaw,
      now,
      'MISSION_REVERSAL',
      Number(amount.toFixed(2)),
      flightId,
      'MISSION_TRANSACTION',
      '',
      missionId,
      '',
      noteReason,
      nowIso,
      actor
    ]);

    reversed++;
  }

  return { reversed: reversed, skipped: skipped, warnings: warnings };
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

function _verifyMxApprovalPassword_(password) {
  const configured = _getMxApprovalPassword_();
  if (!configured) {
    throw new Error('MX approval password not configured. Set script property MX_SUPERVISOR_APPROVAL_PASSWORD.');
  }
  if (String(password || '') !== configured) {
    throw new Error('Invalid chief mechanic password');
  }
  return true;
}

function _resolveTrainingStudentNameFromDispatchRow_(dispatchRow) {
  if (!dispatchRow) return '';

  var rawIdx = typeof DISPATCH_COL !== 'undefined' ? DISPATCH_COL.RAW_DATA : -1;
  if (rawIdx >= 0) {
    try {
      var rawText = String(dispatchRow[rawIdx] || '').trim();
      if (rawText) {
        var parsed = JSON.parse(rawText);
        var training = (parsed && typeof parsed.training === 'object') ? parsed.training : null;
        if (training) {
          var trainingCandidates = [
            training.student,
            training.traineeName,
            training.studentName,
            training.aluno
          ];
          for (var i = 0; i < trainingCandidates.length; i++) {
            var candidate = String(trainingCandidates[i] || '').trim();
            if (candidate) return candidate;
          }
        }
      }
    } catch (e) {}
  }

  var copilotIdx = typeof DISPATCH_COL !== 'undefined' ? DISPATCH_COL.COPILOT : -1;
  if (copilotIdx >= 0) {
    var copilotName = String(dispatchRow[copilotIdx] || '').trim();
    if (copilotName && _resolveBillableAlunoNameFromPilotsSheet_(copilotName)) return copilotName;
  }

  return '';
}

function _resolveBillableAlunoNameFromPilotsSheet_(pilotName) {
  var target = String(pilotName || '').trim().toUpperCase();
  if (!target) return '';

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pilotsSh = ss.getSheetByName(APP_SHEETS.PILOTS);
    if (!pilotsSh) return '';

    var headers = _toolsEnsurePilotStudentHeaders_(pilotsSh);
    var data = pilotsSh.getDataRange().getValues();
    var nameIdx = _toolsHeaderIndexFromCandidates_(headers, ['PILOT_NAME', 'PILOT', 'NAME', 'STUDENT_NAME', 'NOME']);
    var personTypeIdx = _toolsHeaderIndexFromCandidates_(headers, ['PERSON_TYPE']);

    for (var i = 1; i < data.length; i++) {
      var rowName = nameIdx >= 0 ? String(data[i][nameIdx] || '').trim().toUpperCase() : '';
      if (!rowName || rowName !== target) continue;

      var personType = personTypeIdx >= 0 ? String(data[i][personTypeIdx] || '').trim().toUpperCase() : '';
      if (personType === 'ALUNO' || personType === 'STUDENT') return String(pilotName || '').trim();
      return '';
    }
  } catch (e) {}

  return '';
}

function _resolveDutyCrewNameFromPilotsSheet_(pilotName) {
  var target = String(pilotName || '').trim().toUpperCase();
  if (!target) return '';

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pilotsSh = ss.getSheetByName(APP_SHEETS.PILOTS);
    if (!pilotsSh) return '';

    var headers = _toolsEnsurePilotStudentHeaders_(pilotsSh);
    var data = pilotsSh.getDataRange().getValues();
    var nameIdx = _toolsHeaderIndexFromCandidates_(headers, ['PILOT_NAME', 'PILOT', 'NAME', 'STUDENT_NAME', 'NOME']);
    var personTypeIdx = _toolsHeaderIndexFromCandidates_(headers, ['PERSON_TYPE']);
    var fallbackName = '';

    for (var i = 1; i < data.length; i++) {
      var rowName = nameIdx >= 0 ? String(data[i][nameIdx] || '').trim().toUpperCase() : '';
      if (!rowName || rowName !== target) continue;

      var resolvedName = nameIdx >= 0 ? String(data[i][nameIdx] || '').trim() : String(pilotName || '').trim();
      if (!resolvedName) continue;

      var personType = personTypeIdx >= 0 ? String(data[i][personTypeIdx] || '').trim().toUpperCase() : '';
      if (personType === 'ALUNO' || personType === 'STUDENT') return resolvedName;
      if (!fallbackName) fallbackName = resolvedName;
    }
  } catch (e) {}

  return fallbackName;
}

function _pilotRunwayCheckHeaders_() {
  return [
    'Check_ID',
    'Pilot_Name',
    'Pilot_Email',
    'Staff_ID',
    'ICAO',
    'Runway_Ident',
    'Auth_Scope',
    'Status',
    'Date_Checked',
    'Expiry_Date',
    'Approved_By',
    'Source',
    'Notes',
    'Created_At',
    'Updated_At'
  ];
}

function _pilotRunwayChecksSheet_() {
  return _ensureToolsSchemaSheet_(APP_SHEETS.CHECKS, _pilotRunwayCheckHeaders_(), '#2c3e50');
}

function _checksNormalizeIcaoToken_(value) {
  return String(value || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function _checksSplitIcaoList_(raw) {
  return String(raw || '')
    .split(/[;,]/)
    .map(function(item) { return _checksNormalizeIcaoToken_(item); })
    .filter(Boolean);
}

function _checksIsActiveStatus_(status, expiryDate) {
  var statusUpper = String(status || 'ACTIVE').trim().toUpperCase();
  if (statusUpper && ['ACTIVE', 'AUTHORIZED', 'VALID'].indexOf(statusUpper) === -1) return false;
  var expiry = safeDateStr(expiryDate || '');
  return !expiry || expiry >= safeDateStr(new Date());
}

function _pilotRunwayCheckRecordFromRow_(headers, row, rowNumber) {
  var pilotIdx = _toolsHeaderIndexFromCandidates_(headers, ['PILOT_NAME', 'PILOT']);
  var emailIdx = _toolsHeaderIndexFromCandidates_(headers, ['PILOT_EMAIL', 'EMAIL']);
  var staffIdIdx = _toolsHeaderIndexFromCandidates_(headers, ['STAFF_ID']);
  var icaoIdx = _toolsHeaderIndexFromCandidates_(headers, ['ICAO', 'AIRPORT_ICAO', 'DESTINATION_ICAO']);
  var runwayIdx = _toolsHeaderIndexFromCandidates_(headers, ['RUNWAY_IDENT', 'RWY_IDENT', 'RUNWAY', 'RWY']);
  var scopeIdx = _toolsHeaderIndexFromCandidates_(headers, ['AUTH_SCOPE', 'SCOPE']);
  var statusIdx = _toolsHeaderIndexFromCandidates_(headers, ['STATUS']);
  var checkedIdx = _toolsHeaderIndexFromCandidates_(headers, ['DATE_CHECKED', 'CHECK_DATE']);
  var expiryIdx = _toolsHeaderIndexFromCandidates_(headers, ['EXPIRY_DATE', 'VALID_UNTIL']);
  var approvedByIdx = _toolsHeaderIndexFromCandidates_(headers, ['APPROVED_BY']);
  var sourceIdx = _toolsHeaderIndexFromCandidates_(headers, ['SOURCE']);
  var notesIdx = _toolsHeaderIndexFromCandidates_(headers, ['NOTES']);
  var authDestIdx = _toolsHeaderIndexFromCandidates_(headers, ['AUTH_DESTINATIONS', 'AUTHORIZED_DESTINATIONS', 'DESTINATIONS', 'DESTINATION', 'AUTH_DESTINATION']);

  var pilotName = pilotIdx >= 0 ? String(row[pilotIdx] || '').trim() : '';
  var icao = icaoIdx >= 0 ? _checksNormalizeIcaoToken_(row[icaoIdx]) : '';
  var legacyIcaos = authDestIdx >= 0 ? _checksSplitIcaoList_(row[authDestIdx]) : [];
  var icaos = legacyIcaos.slice();
  if (icao && icaos.indexOf(icao) === -1) icaos.push(icao);

  return {
    rowNumber: rowNumber,
    pilotName: pilotName,
    pilotEmail: emailIdx >= 0 ? String(row[emailIdx] || '').trim().toLowerCase() : '',
    staffId: staffIdIdx >= 0 ? String(row[staffIdIdx] || '').trim() : '',
    icao: icao,
    runwayIdent: runwayIdx >= 0 ? String(row[runwayIdx] || '').trim().toUpperCase() : '',
    authScope: scopeIdx >= 0 ? String(row[scopeIdx] || '').trim().toUpperCase() : '',
    status: statusIdx >= 0 ? String(row[statusIdx] || '').trim().toUpperCase() : '',
    dateChecked: checkedIdx >= 0 ? safeDateStr(row[checkedIdx]) : '',
    expiryDate: expiryIdx >= 0 ? safeDateStr(row[expiryIdx]) : '',
    approvedBy: approvedByIdx >= 0 ? String(row[approvedByIdx] || '').trim() : '',
    source: sourceIdx >= 0 ? String(row[sourceIdx] || '').trim() : '',
    notes: notesIdx >= 0 ? String(row[notesIdx] || '').trim() : '',
    icaos: icaos,
    isActive: _checksIsActiveStatus_(statusIdx >= 0 ? row[statusIdx] : '', expiryIdx >= 0 ? row[expiryIdx] : '')
  };
}

function _collectPilotDestinationChecksMap_(headers, rows) {
  var map = {};
  (rows || []).forEach(function(row, idx) {
    var rec = _pilotRunwayCheckRecordFromRow_(headers || [], row || [], idx + 2);
    var key = String(rec.pilotName || '').trim().toUpperCase();
    if (!key || !rec.isActive) return;
    if (!map[key]) map[key] = [];
    rec.icaos.forEach(function(code) {
      if (map[key].indexOf(code) === -1) map[key].push(code);
    });
  });
  return map;
}

function _pilotAuthMatchesPerson_(rec, targetEmail, targetName) {
  var recEmail = String(rec && rec.pilotEmail || '').trim().toLowerCase();
  var recName = String(rec && rec.pilotName || '').trim().toUpperCase();
  var email = String(targetEmail || '').trim().toLowerCase();
  var name = String(targetName || '').trim().toUpperCase();
  if (!email && !name) return true;
  if (email && recEmail && recEmail === email) return true;
  if (name && recName && recName === name) return true;
  return false;
}

function _supervisorNormalizeRouteTokens_(raw) {
  return String(raw || '')
    .toUpperCase()
    .replace(/[\u2192>]/g, ',')
    .split(/[\n\r,;]+/)
    .map(function(part) { return String(part || '').trim(); })
    .filter(Boolean);
}

function _supervisorNextRunwayCheckCode_(headers, rows, aircraftType, routeTokens) {
  var codeIdx = _toolsHeaderIndexFromCandidates_(headers, ['TRAINING_CODE']);
  var existing = {};
  if (codeIdx >= 0) {
    (rows || []).forEach(function(row) {
      var code = String((row && row[codeIdx]) || '').trim().toUpperCase();
      if (code) existing[code] = true;
    });
  }

  var acft = String(aircraftType || 'ACFT').toUpperCase().replace(/[^A-Z0-9]+/g, '').slice(0, 12) || 'ACFT';
  var from = String(routeTokens && routeTokens[0] || 'ORIG').toUpperCase().replace(/[^A-Z0-9]+/g, '') || 'ORIG';
  var to = String(routeTokens && routeTokens.length ? routeTokens[routeTokens.length - 1] : 'DEST').toUpperCase().replace(/[^A-Z0-9]+/g, '') || 'DEST';
  var base = ['RWC', acft, from + '_' + to].join('-');
  var candidate = base;
  var seq = 2;
  while (existing[candidate] && seq < 99) {
    candidate = base + '-V' + seq;
    seq++;
  }
  return candidate;
}

function _supervisorApplyTrainingToDispatchRaw_(rawText, training) {
  var parsed = {};
  try { parsed = rawText ? JSON.parse(String(rawText || '{}')) : {}; } catch (e) { parsed = {}; }
  parsed = (parsed && typeof parsed === 'object') ? parsed : {};
  parsed.training = training;
  if (Array.isArray(parsed.legs)) {
    parsed.legs = parsed.legs.map(function(leg) {
      var out = (leg && typeof leg === 'object') ? leg : {};
      out.training = training;
      if (!out.meta || typeof out.meta !== 'object') out.meta = {};
      out.meta.training = training;
      return out;
    });
  }
  return JSON.stringify(parsed);
}

function _supervisorAircraftTypeFromRegistration_(registration) {
  var reg = String(registration || '').trim().toUpperCase();
  if (!reg) return '';
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(APP_SHEETS.AIRCRAFT);
    if (!sh || sh.getLastRow() < 2) return '';
    var values = sh.getDataRange().getValues();
    var headers = values[0] || [];
    var regIdx = _toolsHeaderIndexFromCandidates_(headers, ['REGISTRATION', 'REG', 'TAIL', 'TAIL_NUMBER']);
    var perfTypeIdx = _toolsHeaderIndexFromCandidates_(headers, ['TYPE_FOR_PERFORMANCE']);
    var typeIdx = _toolsHeaderIndexFromCandidates_(headers, ['AIRCRAFT_TYPE', 'TYPE']);
    if (regIdx < 0) return '';
    for (var i = 1; i < values.length; i++) {
      var rowReg = String(values[i][regIdx] || '').trim().toUpperCase();
      if (rowReg !== reg) continue;
      var aircraftType = typeIdx >= 0 ? String(values[i][typeIdx] || '').trim() : '';
      var perfType = perfTypeIdx >= 0 ? String(values[i][perfTypeIdx] || '').trim() : '';
      return String(aircraftType || perfType || '').trim();
    }
  } catch (e) {}
  return '';
}

function _dutyNormalizePilotName_(pilotName) {
  return String(pilotName || '').trim();
}

function _dutyPilotNorm_(value) {
  var text = String(value || '').trim().toUpperCase();
  if (!text) return '';
  try {
    text = text.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  } catch (e) {}
  // Normalize separators so email local parts (e.g. SUPER.OPS) can match names (SUPER OPS).
  return text
    .replace(/[@._-]+/g, ' ')
    .replace(/[^A-Z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function _dutyPilotKeys_(value) {
  var keys = {};
  var raw = String(value || '').trim().toUpperCase();
  if (raw) {
    try {
      raw = raw.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    } catch (e) {}
  }

  var norm = _dutyPilotNorm_(raw);
  if (norm) {
    keys[norm] = true;
    keys[norm.replace(/\s+/g, '')] = true;
  }

  var emailMatch = raw.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/);
  var email = emailMatch ? String(emailMatch[0] || '').trim() : '';
  if (email) {
    keys[email] = true;
    var local = String(email.split('@')[0] || '').trim();
    if (local) {
      keys[local] = true;
      var localNorm = _dutyPilotNorm_(local);
      if (localNorm) {
        keys[localNorm] = true;
        keys[localNorm.replace(/\s+/g, '')] = true;
      }
    }
  }
  return Object.keys(keys);
}

function _dutyPilotLikelySame_(a, b) {
  var aa = _dutyPilotKeys_(a);
  var bb = _dutyPilotKeys_(b);
  if (!aa.length || !bb.length) return false;
  for (var i = 0; i < aa.length; i++) {
    if (bb.indexOf(aa[i]) >= 0) return true;
  }
  var aNorm = _dutyPilotNorm_(a);
  var bNorm = _dutyPilotNorm_(b);
  if (aNorm && bNorm) {
    if (aNorm.length >= 6 && bNorm.indexOf(aNorm) >= 0) return true;
    if (bNorm.length >= 6 && aNorm.indexOf(bNorm) >= 0) return true;
  }
  return false;
}

function _dutyYmd_(value) {
  if (value == null || value === '') return '';
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  if (typeof value === 'number' && isFinite(value)) {
    // Google Sheets serial date values use 1899-12-30 as day 0.
    var serialDate = new Date(Math.round((value - 25569) * 86400000));
    if (!isNaN(serialDate.getTime())) {
      return Utilities.formatDate(serialDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  }

  var text = String(value || '').trim();
  if (!text) return '';
  var iso = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return text;
  var br = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
  if (br) {
    var day = Number(br[1] || 0);
    var month = Number(br[2] || 0) - 1;
    var year = Number(br[3] || 0);
    var parsed = new Date(year, month, day);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  }

  var d = new Date(text);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _dutyParseHhmmToHours_(raw) {
  if (raw instanceof Date) {
    if (isNaN(raw.getTime())) return 0;
    return Math.max(0, raw.getHours() + (raw.getMinutes() / 60) + (raw.getSeconds() / 3600));
  }
  if (typeof raw === 'number' && isFinite(raw)) {
    return raw > 0 && raw < 1 ? raw * 24 : raw;
  }
  const text = String(raw || '').trim();
  if (!text) return 0;
  const mm = text.match(/^(\d{1,2}):(\d{2})$/);
  if (mm) {
    const hh = Math.max(0, Math.min(23, Number(mm[1] || 0)));
    const mn = Math.max(0, Math.min(59, Number(mm[2] || 0)));
    return hh + (mn / 60);
  }
  const dec = parseFloat(text);
  return isFinite(dec) && dec > 0 ? dec : 0;
}

function _dutyTotalDutyHours_(startTime, endTime) {
  const startHours = _dutyParseHhmmToHours_(startTime);
  const endHours = _dutyParseHhmmToHours_(endTime);
  if (!isFinite(startHours) || !isFinite(endHours) || startHours < 0 || endHours < 0) return 0;
  var diff = endHours - startHours;
  if (diff < 0) diff += 24;
  return diff > 0 ? diff : 0;
}

function _dutyFmtHoursPt_(hours) {
  const n = Number(hours || 0);
  if (!isFinite(n) || n <= 0) return '0,0';
  return n.toFixed(1).replace('.', ',');
}

function _dutyFmtDurationAny_(value) {
  if (value == null || value === '') return '';
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    const year = value.getFullYear();
    if (year <= 1900) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
    }
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  }

  if (typeof value === 'number' && isFinite(value)) {
    var serialDate = new Date(Math.round((value - 25569) * 86400000));
    if (!isNaN(serialDate.getTime())) {
      if (serialDate.getFullYear() <= 1900) {
        return Utilities.formatDate(serialDate, Session.getScriptTimeZone(), 'HH:mm');
      }
      return Utilities.formatDate(serialDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    }
    return String(value);
  }

  var text = String(value || '').trim();
  if (!text) return '';
  if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(text)) return text.slice(0, 5);
  var parsed = new Date(text);
  if (isNaN(parsed.getTime())) return text;
  if (parsed.getFullYear() <= 1900) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'HH:mm');
  }
  return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function _dutyGetHeaderIndexByAliases_(headers, aliases, fallbackIdx) {
  const norm = function(v) {
    return String(v || '').toUpperCase().trim().replace(/\s+/g, '_');
  };
  const normalizedHeaders = (Array.isArray(headers) ? headers : []).map(norm);
  for (let i = 0; i < aliases.length; i++) {
    const idx = normalizedHeaders.indexOf(norm(aliases[i]));
    if (idx >= 0) return idx;
  }
  return Number(fallbackIdx || -1);
}

function _dutyFlightRowCompleted_(row) {
  if (!row || !Array.isArray(row)) return false;
  const landed = row[LOG_FLIGHT_COL.LANDED];
  const onBlocks = row[LOG_FLIGHT_COL.ON_BLOCKS];
  const brakesApplied = row[LOG_FLIGHT_COL.BRAKES_APPLIED];
  const asText = function(v) {
    if (v instanceof Date) return isNaN(v.getTime()) ? '' : v.toISOString();
    return String(v || '').trim();
  };
  return !!(asText(landed) || asText(onBlocks) || asText(brakesApplied));
}

function _dutyDailyFlightSummary_(ss, pilotName, ymd, options) {
  const summary = {
    totalFlightHours: 0,
    flights: [],
    funds: [],
    summaryLine: ''
  };
  const opts = (options && typeof options === 'object') ? options : {};
  const pilotTarget = _dutyNormalizePilotName_(pilotName);
  if (!pilotTarget || !ymd) return summary;

  const resolvedLogData = Array.isArray(opts.flightData) ? opts.flightData : (function() {
    const logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
    return logSheet ? logSheet.getDataRange().getValues() : null;
  })();
  if (resolvedLogData && resolvedLogData.length > 1) {
    for (let i = 1; i < resolvedLogData.length; i++) {
      const row = resolvedLogData[i];
      const rowPilot = String(row[LOG_FLIGHT_COL.PILOT] || '').trim();
      const rowDateYmd = _dutyYmd_(row[LOG_FLIGHT_COL.DATE]);
      if (rowDateYmd !== ymd) continue;
      if (!_dutyFlightRowCompleted_(row)) continue;

      let isMatchingFlight = false;

      if (rowPilot && _dutyPilotLikelySame_(rowPilot, pilotTarget)) {
        isMatchingFlight = true;
      }

      if (!isMatchingFlight) {
        try {
          const loadJsonRaw = String(row[LOG_FLIGHT_COL.ACTUAL_LOAD_JSON] || '').trim();
          if (loadJsonRaw) {
            const loadJson = JSON.parse(loadJsonRaw);
            if (loadJson && loadJson.seats && typeof loadJson.seats === 'object') {
              for (const seatKey in loadJson.seats) {
                const seat = loadJson.seats[seatKey];
                if (seat && seat.passenger && seat.passenger.name) {
                  const passengerName = String(seat.passenger.name || '').trim();
                  if (passengerName && _dutyPilotLikelySame_(passengerName, pilotTarget)) {
                    isMatchingFlight = true;
                    break;
                  }
                }
              }
            }
          }
        } catch (e) {}
      }

      if (!isMatchingFlight) continue;

      const flightId = String(row[LOG_FLIGHT_COL.FLIGHT_ID] || '').trim();
      const acft = String(row[LOG_FLIGHT_COL.ACFT] || '').trim().toUpperCase();
      const from = String(row[LOG_FLIGHT_COL.FROM] || '').trim().toUpperCase();
      const to = String(row[LOG_FLIGHT_COL.TO] || '').trim().toUpperCase();
      const totalTimeRaw = row[LOG_FLIGHT_COL.TOTAL_TIME];
      const flightHours = _dutyParseHhmmToHours_(totalTimeRaw);
      const totalTimeLabel = _dutyFmtDurationAny_(totalTimeRaw);

      summary.totalFlightHours += flightHours;
      summary.flights.push({
        flightId: flightId,
        acft: acft,
        from: from,
        to: to,
        totalTime: totalTimeLabel,
        flightHours: Number(flightHours.toFixed(2))
      });
    }
  }

  try {
    const resolvedTransData = Array.isArray(opts.transData) ? opts.transData : (function() {
      const transSheet = ss.getSheetByName(APP_SHEETS.TRANSACTIONS);
      return transSheet ? transSheet.getDataRange().getValues() : null;
    })();
    if (resolvedTransData && resolvedTransData.length > 1) {
      const headers = resolvedTransData[0] || [];
      const flightIdIdx = _dutyGetHeaderIndexByAliases_(headers, ['FLIGHT_ID', 'MISSION_ID'], 0);
      const fundIdx = _dutyGetHeaderIndexByAliases_(headers, ['FUND', 'FUND_NAME'], 1);
      const flightIds = {};
      summary.flights.forEach(function(f) { if (f.flightId) flightIds[f.flightId] = true; });

      const funds = {};
      for (let i = 1; i < resolvedTransData.length; i++) {
        const row = resolvedTransData[i];
        const fid = String((row[flightIdIdx] || '')).trim();
        if (!fid || !flightIds[fid]) continue;
        const fund = String((row[fundIdx] || '')).trim();
        if (fund) funds[fund] = true;
      }
      summary.funds = Object.keys(funds).sort();
    }
  } catch (e) {}

  summary.totalFlightHours = Number(summary.totalFlightHours.toFixed(2));
  const lines = summary.flights.map(function(f) {
    const route = [f.from, f.to].filter(Boolean).join(' - ');
    const timeTxt = f.totalTime ? (' (' + f.totalTime + ')') : '';
    return [f.flightId, f.acft, route].filter(Boolean).join(', ') + timeTxt;
  });
  const fundsTxt = summary.funds.length ? (' | Fundos: ' + summary.funds.join(', ')) : '';
  summary.summaryLine = lines.length
    ? (lines.join(' ; ') + fundsTxt)
    : ('Sem voos registrados no dia' + fundsTxt);

  return summary;
}

function _dutyReadLatestEntry_(ss, pilotName, ymd) {
  const result = {
    status: '',
    startTime: '',
    endTime: '',
    flightHours: '',
    description: '',
    ignoredMorning: false,
    ignoredEvening: false,
    updatedAt: '',
    sourceRow: 0
  };

  const sh = ss.getSheetByName(APP_SHEETS.DUTY_LOG);
  if (!sh) return result;
  const rows = sh.getDataRange().getValues();
  if (!rows || rows.length < 2) return result;

  const pilotTarget = _dutyNormalizePilotName_(pilotName).toUpperCase();
  for (let i = rows.length - 1; i >= 1; i--) {
    const row = rows[i];
    const rowDate = _dutyYmd_(row[DUTY_LOG_COL.DATE]);
    const rowPilot = String(row[DUTY_LOG_COL.PILOT] || '').trim().toUpperCase();
    if (rowDate !== ymd || rowPilot !== pilotTarget) continue;

    const title = String(row[DUTY_LOG_COL.TITLE] || '').trim().toUpperCase();
    if (title.indexOf('DUTY') < 0 && title.indexOf('JORNADA') < 0) continue;

    const jsonRaw = String(row[5] || '').trim();
    let parsed = {};
    try { parsed = jsonRaw ? JSON.parse(jsonRaw) : {}; } catch (e) { parsed = {}; }

    return {
      status: String(parsed.status || title || '').trim().toUpperCase(),
      startTime: String(parsed.startTime || '').trim(),
      endTime: String(parsed.endTime || '').trim(),
      flightHours: String(parsed.flightHours == null ? '' : parsed.flightHours).trim(),
      description: String(parsed.description || row[DUTY_LOG_COL.DESC_PRIMARY] || row[DUTY_LOG_COL.DESC_FALLBACK] || '').trim(),
      ignoredMorning: !!parsed.ignoredMorning,
      ignoredEvening: !!parsed.ignoredEvening,
      updatedAt: String(parsed.updatedAt || '').trim(),
      sourceRow: i + 1
    };
  }

  return result;
}

function _dutyDepartmentFromDescription_(text) {
  var raw = String(text || '').trim();
  if (!raw) return '';
  var match = raw.match(/^Departamento:\s*(.+)$/im);
  return match ? String(match[1] || '').trim() : '';
}

function _toolsLoadStaffRows_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pilotsSheet = getRequiredSheet_(ss, APP_SHEETS.PILOTS, '_toolsLoadStaffRows_');
  var data = pilotsSheet.getDataRange().getValues();
  if (!data || data.length < 2) return [];
  var headers = data[0] || [];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var staff = _toolsStaffRecordFromRow_(headers, data[i], i + 1);
    if (!staff || (!staff.staffName && !staff.email)) continue;
    rows.push(staff);
  }
  return rows;
}

function _toolsSupervisorLike_(staff) {
  if (!staff) return false;
  var roleText = [staff.primaryRole].concat(staff.staffRoles || []).join(' ').toUpperCase();
  if (roleText.indexOf('SUPERVISOR') >= 0) return true;
  if (roleText.indexOf('MANAGER') >= 0) return true;
  if (roleText.indexOf('GERENTE') >= 0) return true;
  if (roleText.indexOf('COORDENADOR') >= 0) return true;
  return !!staff.canCoordinateFlights || !!staff.canApproveDeferments;
}

function _toolsDutyContext_(overrideEmailRaw) {
  var userEmail = _toolsCurrentUserEmail_(overrideEmailRaw);
  var staffRows = _toolsLoadStaffRows_();
  var me = _toolsFindStaffByEmailOrId_(staffRows, userEmail, '') || null;
  var schedulerCanManageConfig = false;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var permSheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_PERMISSIONS || 'SCHED_Permissions', '_toolsDutyContext_');
    var rec = _schedulerPermissionRowByEmail_(permSheet, userEmail);
    if (rec) {
      var activeIdx = _schedulerHeaderIndex_(rec.headers, 'ACTIVE');
      var canEditRulesIdx = _schedulerHeaderIndex_(rec.headers, 'CAN_EDIT_RULES');
      var canManagePermsIdx = _schedulerHeaderIndex_(rec.headers, 'CAN_MANAGE_PERMISSIONS');
      var isActive = activeIdx < 0 ? true : _schedulerTruthyFlag_(rec.row[activeIdx]);
      schedulerCanManageConfig = isActive && (
        (canEditRulesIdx >= 0 && _schedulerTruthyFlag_(rec.row[canEditRulesIdx])) ||
        (canManagePermsIdx >= 0 && _schedulerTruthyFlag_(rec.row[canManagePermsIdx]))
      );
    }
  } catch (e) {
    schedulerCanManageConfig = false;
  }

  var canManageAll = _toolsSupervisorLike_(me);
  return {
    userEmail: userEmail,
    me: me,
    staffRows: staffRows,
    canManageAll: canManageAll,
    canManageConfig: canManageAll || schedulerCanManageConfig
  };
}

function _dutyPreferredEmailKey_(pilotNameRaw) {
  var pilotKey = String(pilotNameRaw || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
  return pilotKey ? 'duty_preferred_email_' + pilotKey : '';
}

function _dutyReadSavedPreferredEmail_(pilotNameRaw) {
  var key = _dutyPreferredEmailKey_(pilotNameRaw);
  if (!key) return '';
  try {
    return String(PropertiesService.getScriptProperties().getProperty(key) || '').trim().toLowerCase();
  } catch (e) {
    return '';
  }
}

function _dutySavePreferredEmail_(pilotNameRaw, emailRaw) {
  var key = _dutyPreferredEmailKey_(pilotNameRaw);
  var email = String(emailRaw || '').trim().toLowerCase();
  if (!key) return '';
  try {
    var props = PropertiesService.getScriptProperties();
    if (email) props.setProperty(key, email);
    else props.deleteProperty(key);
  } catch (e) {}
  return email;
}

function _toolsDutyCanAccessPilot_(ctx, pilotNameRaw) {
  if (ctx && ctx.canManageAll) return true;
  var pilotName = String(pilotNameRaw || '').trim();
  if (!pilotName) return false;
  var me = (ctx && ctx.me) || null;
  var myName = String((me && me.staffName) || '').trim();
  var myEmail = String((me && me.email) || (ctx && ctx.userEmail) || '').trim();
  if (_dutyPilotLikelySame_(pilotName, myName)) return true;
  if (_dutyPilotLikelySame_(pilotName, myEmail)) return true;
  return false;
}

function _toolsDutyNormalizeEmailList_(rawValue) {
  return String(rawValue || '')
    .split(/[\n,;]+/)
    .map(function(value) { return String(value || '').trim().toLowerCase(); })
    .filter(function(value) { return /.+@.+\..+/.test(value); });
}

function _toolsDutyCanViewAllLogs_(ctx) {
  if (ctx && ctx.canManageAll) return true;
  var currentEmail = String((ctx && ctx.userEmail) || '').trim().toLowerCase();
  if (!currentEmail) return false;
  try {
    var cfg = _toolsDutyReadConfigMap_();
    return _toolsDutyNormalizeEmailList_(cfg.DUTY_LOG_VIEWERS || '').indexOf(currentEmail) >= 0;
  } catch (e) {
    return false;
  }
}

function _toolsDutyConfigDefaults_() {
  return {
    DUTY_GEOFENCE_COORDS: '',
    DUTY_MORNING_ALERT_TIME: '08:00',
    DUTY_EVENING_ALERT_TIME: '17:00',
    DUTY_MORNING_WINDOW_MIN: '120',
    DUTY_EVENING_WINDOW_MIN: '180',
    DUTY_GEOFENCE_RADIUS_KM: '8000',
    DUTY_LANDING_THRESHOLD_KTS: '30',
    DUTY_TOUCH_AND_GO_CONFIRM_KTS: '40',
    DUTY_ALERT_RECIPIENTS: '',
    DUTY_LOG_VIEWERS: '',
    PASSENGER_TERMS_STATEMENT_1: 'Doacoes: Entendo que Asas de Socorro e uma organizacao sem fins lucrativos e depende de doacoes para realizar voos.',
    PASSENGER_TERMS_STATEMENT_2: 'Seguranca: A Asas de Socorro adota padroes de seguranca, mas todo voo tem riscos inerentes.',
    PASSENGER_TERMS_STATEMENT_3: 'Operacao: Os voos sao diurnos e podem sofrer atraso, desvio ou cancelamento por clima, seguranca, motivo tecnico ou logistica.',
    PASSENGER_TERMS_STATEMENT_4: 'Responsabilidade por atraso: Nao ha responsabilizacao por perdas decorrentes de atrasos.',
    PASSENGER_TERMS_QR_PROMPT: 'Vou escanear este QR Code para apoiar ou compartilhar este ministerio.',
    PASSENGER_TERMS_QR_URL: '',
    POSTFLIGHT_EMAIL_INCLUDE_OPS_PILOT: 'Y',
    POSTFLIGHT_EMAIL_INCLUDE_INSTRUCTOR: 'N',
    POSTFLIGHT_EMAIL_INCLUDE_STUDENT: 'Y',
    POSTFLIGHT_EMAIL_EXTRA_RECIPIENTS: ''
  };
}

function _toolsDutyConfigDescriptions_() {
  return {
    DUTY_GEOFENCE_COORDS: 'Lista de coordenadas base para gatilhos de jornada (uma por linha, formato flexível).',
    DUTY_MORNING_ALERT_TIME: 'Horário alvo para prompt de início de jornada (HH:MM).',
    DUTY_EVENING_ALERT_TIME: 'Horário alvo para prompt de encerramento de jornada (HH:MM).',
    DUTY_MORNING_WINDOW_MIN: 'Janela em minutos para prompt de manhã.',
    DUTY_EVENING_WINDOW_MIN: 'Janela em minutos para prompt de tarde/noite.',
    DUTY_GEOFENCE_RADIUS_KM: 'Raio de proximidade (metros) para gatilho por localização.',
    DUTY_LANDING_THRESHOLD_KTS: 'Velocidade máxima para considerar pouso concluído.',
    DUTY_TOUCH_AND_GO_CONFIRM_KTS: 'Velocidade mínima para confirmar recolagem em touch-and-go.',
    DUTY_ALERT_RECIPIENTS: 'Emails para alertas de duty time (separados por vírgula).',
    DUTY_LOG_VIEWERS: 'Emails autorizados a ver todos os logs de duty time (separados por vírgula).',
    PASSENGER_TERMS_STATEMENT_1: 'Texto curto do termo de passageiro #1.',
    PASSENGER_TERMS_STATEMENT_2: 'Texto curto do termo de passageiro #2.',
    PASSENGER_TERMS_STATEMENT_3: 'Texto curto do termo de passageiro #3.',
    PASSENGER_TERMS_STATEMENT_4: 'Texto curto do termo de passageiro #4.',
    PASSENGER_TERMS_QR_PROMPT: 'Texto exibido junto ao QR de apoio/compartilhamento.',
    PASSENGER_TERMS_QR_URL: 'URL de destino do QR de apoio/compartilhamento.',
    POSTFLIGHT_EMAIL_INCLUDE_OPS_PILOT: 'Enviar e-mail pós-voo para o piloto operacional (Y/N).',
    POSTFLIGHT_EMAIL_INCLUDE_INSTRUCTOR: 'Enviar e-mail pós-voo para o instrutor (Y/N).',
    POSTFLIGHT_EMAIL_INCLUDE_STUDENT: 'Enviar e-mail pós-voo para o aluno/trainee (Y/N).',
    POSTFLIGHT_EMAIL_EXTRA_RECIPIENTS: 'Emails adicionais para receber pós-voo (separados por vírgula).'
  };
}

function _passengerConsentReadConfigMap_() {
  var out = {
    PASSENGER_TERMS_STATEMENT_1: '',
    PASSENGER_TERMS_STATEMENT_2: '',
    PASSENGER_TERMS_STATEMENT_3: '',
    PASSENGER_TERMS_STATEMENT_4: '',
    PASSENGER_TERMS_QR_PROMPT: '',
    PASSENGER_TERMS_QR_URL: ''
  };
  var defaults = _toolsDutyConfigDefaults_();
  Object.keys(out).forEach(function(key) {
    out[key] = String(defaults[key] || '');
  });

  try {
    var cfg = _toolsDutyReadConfigMap_();
    Object.keys(out).forEach(function(key) {
      if (Object.prototype.hasOwnProperty.call(cfg, key)) {
        out[key] = String(cfg[key] == null ? '' : cfg[key]).trim() || String(defaults[key] || '');
      }
    });
  } catch (e) {}

  return out;
}

function _flightFollowConfigDefaults_() {
  return {
    FLIGHT_FOLLOW_RECIPIENTS: 'acompanhamento@asasdesocorro.org.br',
    FLIGHT_FOLLOW_OUTBOUND_MAP: '',
    FLIGHT_FOLLOW_MAPSHARE_MAP: ''
  };
}

function _flightFollowConfigDescriptions_() {
  return {
    FLIGHT_FOLLOW_RECIPIENTS: 'Email(s) de destino do Flight Following (separados por vírgula). Primeiro email também é usado para leitura do inbox.',
    FLIGHT_FOLLOW_OUTBOUND_MAP: 'Mapeamento por aeronave para outbound (uma linha por registro: PTABC=email@dominio.com).',
    FLIGHT_FOLLOW_MAPSHARE_MAP: 'Links MapShare por aeronave (uma linha por registro: PTABC=https://share.garmin.com/DEVICE).'
  };
}

function _flightFollowMapshareByReg_() {
  var out = {};
  try {
    var cfg = _flightFollowReadConfigMap_();
    var raw = String(cfg.FLIGHT_FOLLOW_MAPSHARE_MAP || '').trim();
    if (!raw) return out;
    raw.split(/\r?\n/).forEach(function(line) {
      var cleaned = String(line || '').trim();
      if (!cleaned || cleaned.indexOf('#') === 0) return;
      var parts = cleaned.split(/[=:]/);
      if (parts.length < 2) return;
      var reg = String(parts[0] || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
      var url = String(parts.slice(1).join(':') || '').trim();
      if (!reg || !url) return;
      if (!/^https?:\/\//i.test(url)) url = 'https://' + url;
      out[reg] = url;
    });
  } catch (e) {}
  return out;
}

function _flightFollowReadConfigMap_() {
  var out = _flightFollowConfigDefaults_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfgSheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_CONFIG || 'SCHED_Config', '_flightFollowReadConfigMap_');
  var cfg = _schedulerReadConfigMap_(cfgSheet);
  Object.keys(out).forEach(function(key) {
    if (Object.prototype.hasOwnProperty.call(cfg, key)) {
      out[key] = String(cfg[key] == null ? '' : cfg[key]).trim();
    }
  });
  return out;
}

function _toolsDutyReadConfigMap_() {
  var out = _toolsDutyConfigDefaults_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfgSheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_CONFIG || 'SCHED_Config', '_toolsDutyReadConfigMap_');
  var cfg = _schedulerReadConfigMap_(cfgSheet);
  Object.keys(out).forEach(function(key) {
    if (Object.prototype.hasOwnProperty.call(cfg, key)) {
      out[key] = String(cfg[key] == null ? '' : cfg[key]).trim();
    }
  });
  return out;
}

function _toolsDutySaveConfigMap_(entries, actorEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_CONFIG || 'SCHED_Config', '_toolsDutySaveConfigMap_');
  var data = sheet.getDataRange().getValues();
  var headers = data.length ? data[0] : [];
  var keyIdx = _schedulerHeaderIndex_(headers, 'CONFIG_KEY');
  var valueIdx = _schedulerHeaderIndex_(headers, 'CONFIG_VALUE');
  var descIdx = _schedulerHeaderIndex_(headers, 'DESCRIPTION');
  var activeIdx = _schedulerHeaderIndex_(headers, 'ACTIVE');
  var updAtIdx = _schedulerHeaderIndex_(headers, 'UPDATED_AT');
  var updByIdx = _schedulerHeaderIndex_(headers, 'UPDATED_BY');
  if (keyIdx < 0 || valueIdx < 0) throw new Error('SCHED_Config is missing required columns');

  var now = new Date();
  var changed = 0;
  var created = 0;
  for (var e = 0; e < entries.length; e++) {
    var entry = entries[e] || {};
    var key = String(entry.key || '').trim();
    if (!key) continue;
    var value = String(entry.value == null ? '' : entry.value).trim();
    var desc = String(entry.description == null ? '' : entry.description).trim();

    var foundRow = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][keyIdx] || '').trim() === key) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow) {
      sheet.getRange(foundRow, valueIdx + 1).setValue(value);
      if (descIdx >= 0) sheet.getRange(foundRow, descIdx + 1).setValue(desc);
      if (activeIdx >= 0) sheet.getRange(foundRow, activeIdx + 1).setValue('Y');
      if (updAtIdx >= 0) sheet.getRange(foundRow, updAtIdx + 1).setValue(now);
      if (updByIdx >= 0) sheet.getRange(foundRow, updByIdx + 1).setValue(actorEmail || 'tools');
      changed++;
    } else {
      var row = headers.map(function() { return ''; });
      row[keyIdx] = key;
      row[valueIdx] = value;
      if (descIdx >= 0) row[descIdx] = desc;
      if (activeIdx >= 0) row[activeIdx] = 'Y';
      if (updAtIdx >= 0) row[updAtIdx] = now;
      if (updByIdx >= 0) row[updByIdx] = actorEmail || 'tools';
      sheet.appendRow(row);
      created++;
    }
  }

  return { success: true, updated: changed, created: created };
}

function _dutyAppTodayBsb_() {
  return Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');
}

function _dutyLooksLikeEmail_(value) {
  return /@/.test(String(value || '').trim());
}

function _dutyLookupPilotNameByEmail_(emailRaw) {
  var email = String(emailRaw || '').trim().toLowerCase();
  if (!email) return '';
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(APP_SHEETS.PILOTS);
    if (!sh) return '';
    var data = sh.getDataRange().getValues();
    if (!data || data.length < 2) return '';
    var headers = data[0] || [];
    var nameIdx = _toolsHeaderIndexFromCandidates_(headers, ['PILOT_NAME', 'PILOT', 'NAME', 'STAFF_NAME', 'NOME']);
    var emailIdx = _toolsHeaderIndexFromCandidates_(headers, ['PILOT_EMAIL', 'EMAIL', 'E_MAIL']);
    if (nameIdx < 0 || emailIdx < 0) return '';
    for (var i = 1; i < data.length; i++) {
      var rowEmail = String(data[i][emailIdx] || '').trim().toLowerCase();
      if (!rowEmail || rowEmail !== email) continue;
      var rowName = String(data[i][nameIdx] || '').trim();
      if (rowName && !_dutyLooksLikeEmail_(rowName)) return rowName;
    }
  } catch (e) {}
  return '';
}

function _dutyResolvePilotDisplayName_(ctx, requestedRaw) {
  var requested = _dutyNormalizePilotName_(requestedRaw);
  if (requested && !_dutyLooksLikeEmail_(requested)) return requested;

  var me = (ctx && ctx.me) || null;
  var meName = _dutyNormalizePilotName_(me && me.staffName);
  if (meName && !_dutyLooksLikeEmail_(meName)) return meName;

  var requestedEmail = _dutyLooksLikeEmail_(requested) ? requested : '';
  var userEmail = _dutyNormalizePilotName_((ctx && ctx.userEmail) || '');
  var emailToLookup = requestedEmail || userEmail;

  var fromPilots = _dutyLookupPilotNameByEmail_(emailToLookup);
  if (fromPilots) return fromPilots;

  var fromStaffByUser = _toolsFindStaffByEmailOrId_((ctx && ctx.staffRows) || [], userEmail, '');
  var fromStaffName = _dutyNormalizePilotName_(fromStaffByUser && fromStaffByUser.staffName);
  if (fromStaffName && !_dutyLooksLikeEmail_(fromStaffName)) return fromStaffName;

  var fromStaffByReq = _toolsFindStaffByEmailOrId_((ctx && ctx.staffRows) || [], requestedEmail, '');
  var fromReqName = _dutyNormalizePilotName_(fromStaffByReq && fromStaffByReq.staffName);
  if (fromReqName && !_dutyLooksLikeEmail_(fromReqName)) return fromReqName;

  return requested || userEmail;
}

function _dutyAppResolvePilotAccess_(pilotNameRaw, overrideEmailRaw) {
  var ctx = _toolsDutyContext_(overrideEmailRaw);
  var resolvedPilot = _dutyResolvePilotDisplayName_(ctx, pilotNameRaw);

  if (!resolvedPilot) {
    return {
      success: false,
      error: 'Nao foi possivel identificar o piloto logado. Confira o email no DB_Pilots.'
    };
  }

  if (!_toolsDutyCanAccessPilot_(ctx, resolvedPilot)) {
    return {
      success: false,
      error: 'Sem permissao para acessar este piloto.'
    };
  }

  return {
    success: true,
    ctx: ctx,
    pilotName: resolvedPilot
  };
}

function _ensurePassengerConsentColumns_(sheet) {
  if (!sheet) return;
  var wanted = ['TERMS_ACCEPTED', 'TERMS_ACCEPTED_AT', 'TERMS_SOURCE'];
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(v) { return String(v || '').trim(); });
  var norms = headerRow.map(function(v) {
    return String(v || '').trim().toUpperCase().replace(/\s+/g, '_').replace(/[^A-Z0-9_]/g, '');
  });
  wanted.forEach(function(col) {
    var norm = String(col).toUpperCase().replace(/\s+/g, '_').replace(/[^A-Z0-9_]/g, '');
    if (norms.indexOf(norm) >= 0) return;
    var nextCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, nextCol).setValue(col);
    norms.push(norm);
  });
}

function _formNormText_(value) {
  var raw = String(value || '').trim().toUpperCase();
  try {
    raw = raw.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  } catch (e) {}
  return raw.replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

function _formToNumber_(value, fallbackValue) {
  var raw = String(value == null ? '' : value).trim().replace(/\./g, '').replace(',', '.');
  var n = parseFloat(raw);
  return isNaN(n) ? (fallbackValue || 0) : n;
}

function _formToDateIso_(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, 'GMT', 'yyyy-MM-dd');
  var raw = String(value).trim();
  if (!raw) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  var br = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (br) {
    var dd = String(parseInt(br[1], 10)).padStart(2, '0');
    var mm = String(parseInt(br[2], 10)).padStart(2, '0');
    var yyyy = br[3];
    return yyyy + '-' + mm + '-' + dd;
  }
  var parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, 'GMT', 'yyyy-MM-dd');
  return '';
}

function _formFindHeaderIdxByAliases_(headers, aliases) {
  var norms = (headers || []).map(function(h) { return _formNormText_(h); });
  for (var a = 0; a < aliases.length; a++) {
    var alias = _formNormText_(aliases[a]);
    var idx = norms.indexOf(alias);
    if (idx >= 0) return idx;
  }
  // Fallback for long/multiline form questions where exact text can drift.
  for (var b = 0; b < aliases.length; b++) {
    var aliasPart = _formNormText_(aliases[b]);
    if (!aliasPart || aliasPart.length < 12) continue;
    for (var i = 0; i < norms.length; i++) {
      if (norms[i].indexOf(aliasPart) >= 0) return i;
    }
  }
  return -1;
}

function _formFindHeaderIdxsByAliases_(headers, aliases) {
  var norms = (headers || []).map(function(h) { return _formNormText_(h); });
  var aliasNorms = (aliases || []).map(function(v) { return _formNormText_(v); });
  var out = [];
  for (var i = 0; i < norms.length; i++) {
    if (aliasNorms.indexOf(norms[i]) >= 0) out.push(i);
  }
  return out;
}

function _formGetSourceSheetByNameOrFirst_(ss, sheetName) {
  var preferred = String(sheetName || '').trim();
  var sh = preferred ? ss.getSheetByName(preferred) : null;
  if (sh) return sh;
  var sheets = ss.getSheets();
  return (sheets && sheets.length) ? sheets[0] : null;
}

function _formInferIdType_(idValue) {
  var raw = String(idValue || '').trim();
  if (!raw) return '';
  var digits = raw.replace(/\D+/g, '');
  if (digits.length === 11) return 'CPF';
  if (digits.length === 9) return 'RG';
  if (/[A-Z]/i.test(raw) && /\d/.test(raw)) return 'PASSPORT';
  return 'ID';
}

function _formNormalizeGender_(value) {
  var v = _formNormText_(value || '');
  if (!v) return '';
  if (v === 'M' || v.indexOf('MASC') === 0 || v === 'MALE') return 'M';
  if (v === 'F' || v.indexOf('FEM') === 0 || v === 'FEMALE') return 'F';
  return '';
}

function _formExtractPassengerBlocksForm1_(headers, row) {
  var out = [];
  var h = headers || [];
  for (var i = 0; i < h.length; i++) {
    var norm = _formNormText_(h[i]);
    var looksLikeName = (norm.indexOf('NOME_COMPLETO') >= 0 || norm.indexOf('FULL_NAME') >= 0 || norm.indexOf('COMPLETE_NAME') >= 0);
    var isRequesterName = (norm.indexOf('DIGITE_SEU_NOME_COMPLETO') >= 0 || norm.indexOf('TYPE_YOUR_FULL_NAME') >= 0);
    if (!looksLikeName || isRequesterName) continue;

    var name = String((row[i] || '')).trim();
    if (!name) continue;

    out.push({
      idx: out.length + 1,
      name: name,
      idNum: String((row[i + 1] || '')).trim(),
      phone: String((row[i + 2] || '')).trim(),
      weightKg: _formToNumber_(row[i + 3], 0),
      dob: _formToDateIso_(row[i + 4])
    });
  }
  return out;
}

function _formPassengerKey_(pax, requesterEmail) {
  var idNum = String(pax && pax.idNum || '').replace(/\s+/g, '').toUpperCase();
  if (idNum) return 'ID:' + idNum;
  var email = String(requesterEmail || '').trim().toLowerCase();
  var n = _formNormText_(pax && pax.name || '');
  var dob = String(pax && pax.dob || '').trim();
  if (email && n) return 'EMAILNAME:' + email + '|' + n;
  if (n && dob) return 'NAMEDOB:' + n + '|' + dob;
  return 'NAME:' + n;
}

function _formGetOrCreateSheetByHeaders_(ss, sheetName, headers) {
  return _schemaEnsureSheetHeaders_(ss, sheetName, headers || []).sheet;
}

function _formUpsertPassengerFromIntake_(passenger, ctx) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getRequiredSheet_(ss, APP_SHEETS.PASSENGERS, '_formUpsertPassengerFromIntake_');
  _ensurePassengerConsentColumns_(sh);
  var data = sh.getDataRange().getValues();
  if (!data || data.length < 1) throw new Error('DB_Passengers header row missing');
  var headers = data[0].map(function(v) { return String(v || '').trim(); });
  var norms = headers.map(function(v) { return _schemaNormHeader_(v); });

  function idxByAliases(aliases) {
    for (var i = 0; i < aliases.length; i++) {
      var idx = norms.indexOf(_schemaNormHeader_(aliases[i]));
      if (idx >= 0) return idx;
    }
    return -1;
  }

  var idxName = idxByAliases(['PASSENGER_NAME', 'NAME', 'NOME', 'NOME_COMPLETO']);
  var idxIdType = idxByAliases(['ID_TYPE', 'ID_TIPO']);
  var idxIdNum = idxByAliases(['ID_NUMBER_CPF_PASSPORT', 'ID_NUMBER', 'CPF', 'DOCUMENTO', 'ID_NUM']);
  var idxDob = idxByAliases(['DOB', 'DATE_OF_BIRTH', 'DATA_DE_NASCIMENTO']);
  var idxWeight = idxByAliases(['WEIGHT_KG', 'WEIGHT_KGS', 'PESO_KG']);
  var idxGender = idxByAliases(['GENDER', 'SEXO']);
  var idxPhone = idxByAliases(['PHONE', 'TELEFONE']);
  var idxNotes = idxByAliases(['NOTES', 'OBSERVACOES', 'OBS']);
  var idxLastFlown = idxByAliases(['LAST_FLOWN', 'LAST_FLOWN_DATE', 'ULTIMO_VOO']);
  var idxTermsAccepted = idxByAliases(['TERMS_ACCEPTED', 'AGREED_TO_TERMS']);
  var idxTermsAcceptedAt = idxByAliases(['TERMS_ACCEPTED_AT', 'AGREED_TO_TERMS_AT', 'TERMS_LAST_ACCEPTED_AT']);
  var idxTermsSource = idxByAliases(['TERMS_SOURCE', 'CONSENT_SOURCE']);

  var matchRow = -1;
  var paxId = String(passenger.idNum || '').replace(/\s+/g, '').toUpperCase();
  var paxNameNorm = _formNormText_(passenger.name || '');
  var paxDob = String(passenger.dob || '').trim();

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var rowId = idxIdNum >= 0 ? String(row[idxIdNum] || '').replace(/\s+/g, '').toUpperCase() : '';
    var rowNameNorm = idxName >= 0 ? _formNormText_(row[idxName] || '') : '';
    var rowDob = idxDob >= 0 ? _formToDateIso_(row[idxDob]) : '';
    if (paxId && rowId && paxId === rowId) {
      matchRow = r + 1;
      break;
    }
    if (!paxId && paxNameNorm && rowNameNorm === paxNameNorm && paxDob && rowDob && paxDob === rowDob) {
      matchRow = r + 1;
      break;
    }
  }

  var values = headers.map(function() { return ''; });
  if (matchRow > 0) {
    values = sh.getRange(matchRow, 1, 1, headers.length).getValues()[0];
  }

  if (idxName >= 0) values[idxName] = passenger.name;
  if (idxIdType >= 0) {
    var inferredIdType = _formInferIdType_(passenger.idNum);
    if (inferredIdType) values[idxIdType] = inferredIdType;
    else values[idxIdType] = String(values[idxIdType] || '');
  }
  if (idxIdNum >= 0 && paxId) values[idxIdNum] = passenger.idNum;
  if (idxDob >= 0 && passenger.dob) values[idxDob] = passenger.dob;
  if (idxWeight >= 0 && passenger.weightKg > 0) values[idxWeight] = passenger.weightKg;
  if (idxGender >= 0) {
    var parsedGender = _formNormalizeGender_(passenger.gender);
    if (parsedGender) values[idxGender] = parsedGender;
  }
  if (idxPhone >= 0 && passenger.phone) values[idxPhone] = passenger.phone;
  if (idxNotes >= 0) {
    var sourceLabel = String(ctx && ctx.formName || 'FORM').trim();
    var note = 'Imported from ' + sourceLabel + ' on ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    values[idxNotes] = String(values[idxNotes] || '').trim();
    if (!values[idxNotes]) values[idxNotes] = note;
    else if (values[idxNotes].indexOf(note) < 0) values[idxNotes] += ' | ' + note;
  }
  if (idxLastFlown >= 0 && !values[idxLastFlown]) values[idxLastFlown] = new Date();

  var intakeDate = _passengerToDate_(ctx && ctx.submittedAt ? ctx.submittedAt : '');
  if (idxTermsAccepted >= 0) values[idxTermsAccepted] = 'Y';
  if (idxTermsAcceptedAt >= 0 && intakeDate) values[idxTermsAcceptedAt] = intakeDate;
  if (idxTermsSource >= 0) values[idxTermsSource] = String(ctx && ctx.formName || 'FORM_INTAKE');

  if (matchRow > 0) sh.getRange(matchRow, 1, 1, headers.length).setValues([values]);
  else sh.appendRow(values);

  return { action: matchRow > 0 ? 'updated' : 'created', rowNumber: matchRow > 0 ? matchRow : sh.getLastRow() };
}

function _formBuildForm1Payload_(headers, row, sourceSheetName, rowNumber) {
  var idxTimestamp = _formFindHeaderIdxByAliases_(headers, ['Carimbo de data/hora', 'Timestamp']);
  var idxConfirmEmail = _formFindHeaderIdxByAliases_(headers, ['Email para confirmação', 'Email for confirmation']);
  var idxRoutePt = _formFindHeaderIdxByAliases_(headers, ['Origem e Destino do voo']);
  var idxRouteEn = _formFindHeaderIdxByAliases_(headers, ['Origin and Destination']);
  var idxChecklistPt = _formFindHeaderIdxByAliases_(headers, ['Leia atentamente cada item abaixo e marque todas as caixas. Ao marcar todos os itens, você declara estar ciente e de acordo com as condições descritas.']);
  var idxConfirmPt = _formFindHeaderIdxByAliases_(headers, ['Você confirma que leu e compreendeu todos os pontos acima?']);
  var idxPaxCount = _formFindHeaderIdxByAliases_(headers, ['Quantos passageiros?', 'How many passengers']);
  var idxSupport = _formFindHeaderIdxByAliases_(headers, ['Apoie este trabalho - Escolha pelo menos duas formas de contribuição:', 'Help us - Choose at least two ways to contribute:']);
  var idxDeclaration = _formFindHeaderIdxByAliases_(headers, ['Escolha uma declaração / Choose one:']);
  var idxValidity = _formFindHeaderIdxByAliases_(headers, ['Validade desta autorização', 'Validity of this authorization']);
  var idxOther = _formFindHeaderIdxByAliases_(headers, ['Você tem mais alguma coisa que gostaria de compartilhar?', "Do you have anything else you'd like to let us know?"]);
  var idxRequesterEmail = _formFindHeaderIdxByAliases_(headers, ['Endereço de e-mail']);
  var idxRequesterName = _formFindHeaderIdxByAliases_(headers, ['Digite seu nome completo / Type your full name']);
  var idxLanguage = _formFindHeaderIdxByAliases_(headers, ['Choose your preferred language/Escolha a seu idioma:']);
  var idxImageAuth = _formFindHeaderIdxByAliases_(headers, ['Uso de Imagem', 'Use of Image Authorization', 'Autorizo também o uso da minha imagem']);
  var idxProhibitedPt = _formFindHeaderIdxByAliases_(headers, ['Declaração de Itens Proibidos', 'Ao solicitar este voo, declaro que não estou transportando nenhum dos seguintes itens:']);
  var idxBaggagePt = _formFindHeaderIdxByAliases_(headers, ['Entendo que Asas de Socorro se reserva o direito de inspecionar minhas bagagens antes do embarque para garantir a segurança de todos os passageiros.']);
  var idxProhibitedEn = _formFindHeaderIdxByAliases_(headers, ['Prohibited Items Declaration', 'By requesting this flight, I declare that I am not carrying any of the following items:']);
  var idxBaggageEn = _formFindHeaderIdxByAliases_(headers, ['I understand that Asas de Socorro reserves the right to inspect my baggage before boarding to ensure the safety of all passengers.']);

  var passengers = _formExtractPassengerBlocksForm1_(headers, row);
  var submittedAtRaw = idxTimestamp >= 0 ? row[idxTimestamp] : '';
  var submittedAtIso = (submittedAtRaw instanceof Date && !isNaN(submittedAtRaw.getTime()))
    ? Utilities.formatDate(submittedAtRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    : String(submittedAtRaw || '').trim();
  var requesterEmail = String((idxRequesterEmail >= 0 ? row[idxRequesterEmail] : '') || '').trim() || String((idxConfirmEmail >= 0 ? row[idxConfirmEmail] : '') || '').trim();
  var route = String((idxRoutePt >= 0 ? row[idxRoutePt] : '') || '').trim() || String((idxRouteEn >= 0 ? row[idxRouteEn] : '') || '').trim();
  var responseKey = [sourceSheetName, rowNumber, submittedAtIso, requesterEmail, route].join('|');

  return {
    formName: 'FORM_1',
    sourceSheet: sourceSheetName,
    sourceRow: rowNumber,
    responseKey: responseKey,
    submittedAt: submittedAtIso,
    confirmationEmail: String((idxConfirmEmail >= 0 ? row[idxConfirmEmail] : '') || '').trim(),
    requesterEmail: requesterEmail,
    requesterName: String((idxRequesterName >= 0 ? row[idxRequesterName] : '') || '').trim(),
    route: route,
    language: String((idxLanguage >= 0 ? row[idxLanguage] : '') || '').trim(),
    passengerCountDeclared: parseInt(String((idxPaxCount >= 0 ? row[idxPaxCount] : '') || '').trim(), 10) || 0,
    checklistAcknowledgement: String((idxChecklistPt >= 0 ? row[idxChecklistPt] : '') || '').trim(),
    declarationConfirmed: String((idxConfirmPt >= 0 ? row[idxConfirmPt] : '') || '').trim(),
    supportActions: String((idxSupport >= 0 ? row[idxSupport] : '') || '').trim(),
    declarationChoice: String((idxDeclaration >= 0 ? row[idxDeclaration] : '') || '').trim(),
    releaseValidity: String((idxValidity >= 0 ? row[idxValidity] : '') || '').trim(),
    otherNotes: String((idxOther >= 0 ? row[idxOther] : '') || '').trim(),
    imageUseAuthorization: String((idxImageAuth >= 0 ? row[idxImageAuth] : '') || '').trim(),
    prohibitedItemsPt: String((idxProhibitedPt >= 0 ? row[idxProhibitedPt] : '') || '').trim(),
    baggageInspectionPt: String((idxBaggagePt >= 0 ? row[idxBaggagePt] : '') || '').trim(),
    prohibitedItemsEn: String((idxProhibitedEn >= 0 ? row[idxProhibitedEn] : '') || '').trim(),
    baggageInspectionEn: String((idxBaggageEn >= 0 ? row[idxBaggageEn] : '') || '').trim(),
    passengers: passengers,
    rawJson: JSON.stringify({ headers: headers, row: row })
  };
}

function _formBuildForm2Payload_(headers, row, sourceSheetName, rowNumber) {
  var idxTimestamp = _formFindHeaderIdxByAliases_(headers, ['Carimbo de data/hora', 'Timestamp']);
  var idxConfirmEmail = _formFindHeaderIdxByAliases_(headers, ['Email para confirmação', 'Email for confirmation']);
  var idxMedicalSituation = _formFindHeaderIdxByAliases_(headers, ['Descreva brevemente a situação médica que motiva esta solicitação']);
  var idxRoute = _formFindHeaderIdxByAliases_(headers, ['Origem e Destino do voo', 'Origin and Destination']);
  var idxUrgency = _formFindHeaderIdxByAliases_(headers, ['Escolha a gravidade da situação, de acordo com os exemplos abaixo']);
  var idxDeclaration = _formFindHeaderIdxByAliases_(headers, ['Declaração', 'Declaração e Assinatura']);
  var whoDeclares = _formFindHeaderIdxsByAliases_(headers, ['Quem está fazendo essa declaração?']);
  var idxWhoDeclares1 = whoDeclares.length ? whoDeclares[0] : -1;
  var idxWhoDeclares2 = whoDeclares.length > 1 ? whoDeclares[1] : -1;
  var idxTreatmentProof = _formFindHeaderIdxByAliases_(headers, ['Comprovante de tratamento/agendamento']);
  var idxMedicalDeclarationRequired = _formFindHeaderIdxByAliases_(headers, ['Declaração médica obrigatória']);
  var idxDeclarantName = _formFindHeaderIdxByAliases_(headers, ['Nome de quem declara']);
  var idxFormalDeclaration = _formFindHeaderIdxByAliases_(headers, ['Encaminhe uma declaração formal de um profissional de saúde explicando por que este voo é necessário, enviada para o WhatsApp (62) 98197-6055.']);
  var idxPublicHealthGap = _formFindHeaderIdxByAliases_(headers, ['Por que os serviços públicos de saúde não conseguiram responder a esta situação? Descreva tentativas de utilização do SUS/SESAI ou outros serviços e o motivo da indisponibilidade.']);
  var idxTransportLimits = _formFindHeaderIdxByAliases_(headers, ['Por que outros meios de transporte não são viáveis?', 'Explique limitações logísticas, climáticas ou de acesso (barco, carro, ônibus etc.)']);
  var idxChecklist = _formFindHeaderIdxByAliases_(headers, ['Leia atentamente cada item abaixo e marque todas as caixas. Ao marcar todos os itens, você declara estar ciente e de acordo com as condições descritas.']);
  var idxConfirmChecklist = _formFindHeaderIdxByAliases_(headers, ['Você confirma que leu e compreendeu todos os pontos acima?']);

  var idxPatientName = _formFindHeaderIdxByAliases_(headers, ['Nome completo', 'Nome completo / Full Name']);
  var idxPatientCpf = _formFindHeaderIdxByAliases_(headers, ['CPF']);
  var idxPatientPhone = _formFindHeaderIdxByAliases_(headers, ['Telefone']);
  var idxPatientWeight = _formFindHeaderIdxByAliases_(headers, ['Peso (kg)', 'Peso']);
  var idxPatientDob = _formFindHeaderIdxByAliases_(headers, ['Data de nascimento']);

  var idxCompanionName = _formFindHeaderIdxByAliases_(headers, ['Quem acompanhará o paciente?                          Nome:']);
  var idxCompanionCpf = idxCompanionName >= 0 ? idxCompanionName + 1 : -1;
  var idxCompanionPhone = idxCompanionName >= 0 ? idxCompanionName + 2 : -1;
  var idxCompanionWeight = idxCompanionName >= 0 ? idxCompanionName + 3 : -1;
  var idxCompanionDob = idxCompanionName >= 0 ? idxCompanionName + 4 : -1;

  var idxSupport = _formFindHeaderIdxByAliases_(headers, ['Escolha pelo menos duas formas de contribuição:']);
  var idxRequesterEmail = _formFindHeaderIdxByAliases_(headers, ['Endereço de e-mail']);

  var submittedAtRaw = idxTimestamp >= 0 ? row[idxTimestamp] : '';
  var submittedAtIso = (submittedAtRaw instanceof Date && !isNaN(submittedAtRaw.getTime()))
    ? Utilities.formatDate(submittedAtRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    : String(submittedAtRaw || '').trim();

  var requesterEmail = String((idxRequesterEmail >= 0 ? row[idxRequesterEmail] : '') || '').trim() || String((idxConfirmEmail >= 0 ? row[idxConfirmEmail] : '') || '').trim();
  var route = String((idxRoute >= 0 ? row[idxRoute] : '') || '').trim();
  var medicalSituation = String((idxMedicalSituation >= 0 ? row[idxMedicalSituation] : '') || '').trim();
  var responseKey = [sourceSheetName, rowNumber, submittedAtIso, requesterEmail, route, _formNormText_(medicalSituation).slice(0, 60)].join('|');

  var passengers = [];
  var patientName = String((idxPatientName >= 0 ? row[idxPatientName] : '') || '').trim();
  if (patientName) {
    passengers.push({
      idx: 1,
      role: 'PATIENT',
      name: patientName,
      idNum: String((idxPatientCpf >= 0 ? row[idxPatientCpf] : '') || '').trim(),
      phone: String((idxPatientPhone >= 0 ? row[idxPatientPhone] : '') || '').trim(),
      weightKg: _formToNumber_(idxPatientWeight >= 0 ? row[idxPatientWeight] : '', 0),
      dob: _formToDateIso_(idxPatientDob >= 0 ? row[idxPatientDob] : '')
    });
  }

  var companionName = String((idxCompanionName >= 0 ? row[idxCompanionName] : '') || '').trim();
  if (companionName) {
    passengers.push({
      idx: passengers.length + 1,
      role: 'COMPANION',
      name: companionName,
      idNum: String((idxCompanionCpf >= 0 ? row[idxCompanionCpf] : '') || '').trim(),
      phone: String((idxCompanionPhone >= 0 ? row[idxCompanionPhone] : '') || '').trim(),
      weightKg: _formToNumber_(idxCompanionWeight >= 0 ? row[idxCompanionWeight] : '', 0),
      dob: _formToDateIso_(idxCompanionDob >= 0 ? row[idxCompanionDob] : '')
    });
  }

  return {
    formName: 'FORM_2',
    sourceSheet: sourceSheetName,
    sourceRow: rowNumber,
    responseKey: responseKey,
    submittedAt: submittedAtIso,
    confirmationEmail: String((idxConfirmEmail >= 0 ? row[idxConfirmEmail] : '') || '').trim(),
    requesterEmail: requesterEmail,
    requesterName: String((idxDeclarantName >= 0 ? row[idxDeclarantName] : '') || '').trim(),
    route: route,
    language: '',
    passengerCountDeclared: passengers.length,
    checklistAcknowledgement: String((idxChecklist >= 0 ? row[idxChecklist] : '') || '').trim(),
    declarationConfirmed: String((idxConfirmChecklist >= 0 ? row[idxConfirmChecklist] : '') || '').trim(),
    supportActions: String((idxSupport >= 0 ? row[idxSupport] : '') || '').trim(),
    declarationChoice: String((idxDeclaration >= 0 ? row[idxDeclaration] : '') || '').trim(),
    releaseValidity: '',
    otherNotes: '',
    imageUseAuthorization: '',
    prohibitedItemsPt: '',
    baggageInspectionPt: '',
    prohibitedItemsEn: '',
    baggageInspectionEn: '',
    medicalSituation: medicalSituation,
    urgencyLevel: String((idxUrgency >= 0 ? row[idxUrgency] : '') || '').trim(),
    declarationBy1: String((idxWhoDeclares1 >= 0 ? row[idxWhoDeclares1] : '') || '').trim(),
    declarationBy2: String((idxWhoDeclares2 >= 0 ? row[idxWhoDeclares2] : '') || '').trim(),
    treatmentProof: String((idxTreatmentProof >= 0 ? row[idxTreatmentProof] : '') || '').trim(),
    medicalDeclarationRequired: String((idxMedicalDeclarationRequired >= 0 ? row[idxMedicalDeclarationRequired] : '') || '').trim(),
    declarantName: String((idxDeclarantName >= 0 ? row[idxDeclarantName] : '') || '').trim(),
    formalDeclarationInstruction: String((idxFormalDeclaration >= 0 ? row[idxFormalDeclaration] : '') || '').trim(),
    publicHealthReason: String((idxPublicHealthGap >= 0 ? row[idxPublicHealthGap] : '') || '').trim(),
    transportLimitations: String((idxTransportLimits >= 0 ? row[idxTransportLimits] : '') || '').trim(),
    passengers: passengers,
    rawJson: JSON.stringify({ headers: headers, row: row })
  };
}

function _formBuildForm3Payload_(headers, row, sourceSheetName, rowNumber) {
  var idxTimestamp = _formFindHeaderIdxByAliases_(headers, ['Carimbo de data/hora', 'Timestamp']);
  var idxConfirmEmail = _formFindHeaderIdxByAliases_(headers, ['Email para confirmação', 'Email for confirmation']);
  var idxChecklist = _formFindHeaderIdxByAliases_(headers, ['Leia atentamente cada item abaixo e marque todas as caixas. Ao marcar todos os itens, você declara estar ciente e de acordo com as condições descritas.']);
  var idxVillageUnderstanding = _formFindHeaderIdxByAliases_(headers, ['Entendimentos com a Asas de Socorro – Antes e Durante Sua Estadia na Aldeia', 'Marque abaixo para indicar que você compreende os pontos apresentados.']);
  var idxFamilyCount = _formFindHeaderIdxByAliases_(headers, ['Quantas pessoas fazem parte da sua família ou equipe pelas quais você é responsável?', '(Inclua você mesmo na contagem.)']);
  var idxContributionContext = _formFindHeaderIdxByAliases_(headers, ['Abaixo, por favor, indique de que forma você poderia contribuir.', 'Opção "Outra"', 'Por favor, marque abaixo tudo o que se aplica à sua situação (pelo menos 2):']);
  var idxDeclaration = _formFindHeaderIdxByAliases_(headers, ['Declaração e Assinatura']);
  var idxValidity = _formFindHeaderIdxByAliases_(headers, ['Validade desta autorização']);
  var idxRequesterEmail = _formFindHeaderIdxByAliases_(headers, ['Endereço de e-mail']);

  var passengers = _formExtractPassengerBlocksForm1_(headers, row).map(function(p, idx) {
    return Object.assign({}, p, { role: idx === 0 ? 'RESPONSIBLE' : 'FAMILY_OR_TEAM' });
  });

  var submittedAtRaw = idxTimestamp >= 0 ? row[idxTimestamp] : '';
  var submittedAtIso = (submittedAtRaw instanceof Date && !isNaN(submittedAtRaw.getTime()))
    ? Utilities.formatDate(submittedAtRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    : String(submittedAtRaw || '').trim();

  var requesterEmail = String((idxRequesterEmail >= 0 ? row[idxRequesterEmail] : '') || '').trim() || String((idxConfirmEmail >= 0 ? row[idxConfirmEmail] : '') || '').trim();
  var declaration = String((idxDeclaration >= 0 ? row[idxDeclaration] : '') || '').trim();
  var responseKey = [sourceSheetName, rowNumber, submittedAtIso, requesterEmail, _formNormText_(declaration).slice(0, 60)].join('|');

  return {
    formName: 'FORM_3',
    sourceSheet: sourceSheetName,
    sourceRow: rowNumber,
    responseKey: responseKey,
    submittedAt: submittedAtIso,
    confirmationEmail: String((idxConfirmEmail >= 0 ? row[idxConfirmEmail] : '') || '').trim(),
    requesterEmail: requesterEmail,
    requesterName: passengers.length ? String(passengers[0].name || '').trim() : '',
    route: '',
    language: '',
    passengerCountDeclared: parseInt(String((idxFamilyCount >= 0 ? row[idxFamilyCount] : '') || '').trim(), 10) || passengers.length,
    checklistAcknowledgement: String((idxChecklist >= 0 ? row[idxChecklist] : '') || '').trim(),
    declarationConfirmed: declaration,
    supportActions: String((idxContributionContext >= 0 ? row[idxContributionContext] : '') || '').trim(),
    declarationChoice: declaration,
    releaseValidity: String((idxValidity >= 0 ? row[idxValidity] : '') || '').trim(),
    otherNotes: '',
    imageUseAuthorization: '',
    prohibitedItemsPt: '',
    baggageInspectionPt: '',
    prohibitedItemsEn: '',
    baggageInspectionEn: '',
    familyGroupCount: String((idxFamilyCount >= 0 ? row[idxFamilyCount] : '') || '').trim(),
    villageStayAcknowledgement: String((idxVillageUnderstanding >= 0 ? row[idxVillageUnderstanding] : '') || '').trim(),
    understandingsAcknowledgement: String((idxChecklist >= 0 ? row[idxChecklist] : '') || '').trim(),
    contributionContext: String((idxContributionContext >= 0 ? row[idxContributionContext] : '') || '').trim(),
    passengers: passengers,
    rawJson: JSON.stringify({ headers: headers, row: row })
  };
}

function _ingestPassengerLiabilityPayload_(submission) {
  var formName = String(submission && submission.formName || 'FORM').trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var liabilityHeaders = [
    'INGESTED_AT',
    'FORM_NAME',
    'SOURCE_SHEET',
    'SOURCE_ROW',
    'RESPONSE_KEY',
    'SUBMITTED_AT',
    'CONFIRMATION_EMAIL',
    'REQUESTER_EMAIL',
    'REQUESTER_NAME',
    'ROUTE',
    'LANGUAGE',
    'PASSENGER_COUNT_DECLARED',
    'DECLARATION_CONFIRMED',
    'CHECKLIST_ACKNOWLEDGEMENT',
    'SUPPORT_ACTIONS',
    'DECLARATION_CHOICE',
    'RELEASE_VALIDITY',
    'IMAGE_USE_AUTHORIZATION',
    'PROHIBITED_ITEMS_PT',
    'BAGGAGE_INSPECTION_PT',
    'PROHIBITED_ITEMS_EN',
    'BAGGAGE_INSPECTION_EN',
    'OTHER_NOTES',
    'PASSENGER_INDEX',
    'PASSENGER_NAME',
    'PASSENGER_ID',
    'PASSENGER_PHONE',
    'PASSENGER_WEIGHT_KG',
    'PASSENGER_DOB',
    'PASSENGER_ROLE',
    'MEDICAL_SITUATION',
    'URGENCY_LEVEL',
    'DECLARATION_BY_1',
    'DECLARATION_BY_2',
    'TREATMENT_PROOF',
    'MEDICAL_DECLARATION_REQUIRED',
    'DECLARANT_NAME',
    'FORMAL_DECLARATION_INSTRUCTION',
    'PUBLIC_HEALTH_REASON',
    'TRANSPORT_LIMITATIONS',
    'FAMILY_GROUP_COUNT',
    'VILLAGE_STAY_ACKNOWLEDGEMENT',
    'UNDERSTANDINGS_ACKNOWLEDGEMENT',
    'CONTRIBUTION_CONTEXT',
    'PASSENGER_KEY',
    'RAW_JSON'
  ];

  var logHeaders = ['INGESTED_AT', 'FORM_NAME', 'SOURCE_SHEET', 'SOURCE_ROW', 'RESPONSE_KEY', 'STATUS', 'DETAIL'];
  var liabilitySheet = _formGetOrCreateSheetByHeaders_(ss, APP_SHEETS.LIABILITY_RELEASES || 'DB_Liability_Releases', liabilityHeaders);
  var logSheet = _formGetOrCreateSheetByHeaders_(ss, APP_SHEETS.FORM_IMPORT_LOG || 'DB_Form_Import_Log', logHeaders);
  var processedKeys = _formGetImportLogKeys_(logSheet);

  if (processedKeys[submission.responseKey]) {
    return { success: true, skipped: true, reason: 'already_processed', responseKey: submission.responseKey };
  }

  var pax = Array.isArray(submission.passengers) ? submission.passengers : [];
  if (!pax.length) {
    _formAppendImportLog_(logSheet, {
      formName: formName,
      sourceSheet: submission.sourceSheet,
      sourceRow: submission.sourceRow,
      responseKey: submission.responseKey,
      status: 'ERROR',
      detail: 'No passenger blocks found'
    });
    return { success: false, error: 'No passenger blocks found', responseKey: submission.responseKey };
  }

  var passengerResults = [];
  pax.forEach(function(p) {
    passengerResults.push(_formUpsertPassengerFromIntake_(p, submission));
    var passengerKey = _formPassengerKey_(p, submission.requesterEmail);
    liabilitySheet.appendRow([
      new Date(),
      formName,
      submission.sourceSheet,
      submission.sourceRow,
      submission.responseKey,
      submission.submittedAt,
      submission.confirmationEmail,
      submission.requesterEmail,
      submission.requesterName,
      submission.route,
      submission.language,
      submission.passengerCountDeclared,
      submission.declarationConfirmed,
      submission.checklistAcknowledgement,
      submission.supportActions,
      submission.declarationChoice,
      submission.releaseValidity,
      submission.imageUseAuthorization,
      submission.prohibitedItemsPt,
      submission.baggageInspectionPt,
      submission.prohibitedItemsEn,
      submission.baggageInspectionEn,
      submission.otherNotes,
      p.idx,
      p.name,
      p.idNum,
      p.phone,
      p.weightKg,
      p.dob,
      String(p.role || ''),
      String(submission.medicalSituation || ''),
      String(submission.urgencyLevel || ''),
      String(submission.declarationBy1 || ''),
      String(submission.declarationBy2 || ''),
      String(submission.treatmentProof || ''),
      String(submission.medicalDeclarationRequired || ''),
      String(submission.declarantName || ''),
      String(submission.formalDeclarationInstruction || ''),
      String(submission.publicHealthReason || ''),
      String(submission.transportLimitations || ''),
      String(submission.familyGroupCount || ''),
      String(submission.villageStayAcknowledgement || ''),
      String(submission.understandingsAcknowledgement || ''),
      String(submission.contributionContext || ''),
      passengerKey,
      submission.rawJson
    ]);
  });

  _formAppendImportLog_(logSheet, {
    formName: formName,
    sourceSheet: submission.sourceSheet,
    sourceRow: submission.sourceRow,
    responseKey: submission.responseKey,
    status: 'OK',
    detail: 'Passengers imported: ' + passengerResults.length
  });

  return { success: true, responseKey: submission.responseKey, passengersImported: passengerResults.length, passengerResults: passengerResults };
}

function _ingestForm1Payload_(submission) {
  var payload = Object.assign({}, submission || {}, { formName: 'FORM_1' });
  return _ingestPassengerLiabilityPayload_(payload);
}

function _ingestForm2Payload_(submission) {
  var payload = Object.assign({}, submission || {}, { formName: 'FORM_2' });
  return _ingestPassengerLiabilityPayload_(payload);
}

function _ingestForm3Payload_(submission) {
  var payload = Object.assign({}, submission || {}, { formName: 'FORM_3' });
  return _ingestPassengerLiabilityPayload_(payload);
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
  if (k === 'maneuvers' || k === 'ref_maneuvers') return APP_SHEETS.MANEUVERS || 'Ref_Maneuvers';
  if (k === 'schedcoverage' || k === 'sched_coverage' || k === 'scheduler_coverage') return APP_SHEETS.SCHED_COVERAGE_RULES || 'SCHED_Coverage_Requirements';
  if (k === 'schedcompat' || k === 'sched_compat' || k === 'scheduler_compat') return APP_SHEETS.SCHED_ROLE_COMPAT || 'SCHED_Role_Compatibility';
  if (k === 'schedconfig' || k === 'sched_config' || k === 'scheduler_config') return APP_SHEETS.SCHED_CONFIG || 'SCHED_Config';
  if (k === 'schedquals' || k === 'sched_quals' || k === 'scheduler_quals') return APP_SHEETS.SCHED_STAFF_QUALS || 'SCHED_Staff_Qualifications';
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

function _toolsTruthyFlag_(value) {
  var raw = String(value == null ? '' : value).trim().toUpperCase();
  return raw === 'Y' || raw === 'YES' || raw === 'TRUE' || raw === '1' || raw === 'SIM' || raw === 'ATIVO';
}

function _toolsCanonicalAircraftType_(value) {
  var text = String(value || '').trim();
  return text ? text.replace(/\s+/g, ' ').trim() : '';
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

function _lessonEvalHeaders_() {
  return [
    'EVAL_ID',
    'MISSION_ID',
    'TRAINING_CODE',
    'PILOT',
    'AIRCRAFT',
    'MISSION_DATE',
    'STATUS',
    'RESULT',
    'RESULT_AT',
    'SUBMITTED_AT',
    'SOURCE_TAB',
    'SCORES_JSON',
    'NOTES_JSON',
    'UPDATED_AT',
    'UPDATED_BY',
    'CREATED_AT',
    'CREATED_BY'
  ];
}

function _lessonEvalSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = APP_SHEETS.FLIGHT_LESSON_EVALS || 'DB_Flight_Lesson_Evals';
  return _schemaEnsureSheetHeaders_(ss, sheetName, _lessonEvalHeaders_()).sheet;
}

function _lessonEvalHeaderMap_(headers) {
  var out = {};
  (headers || []).forEach(function(h, i) {
    out[String(h || '').trim().toUpperCase()] = i;
  });
  return out;
}

function _lessonEvalJsonSafe_(raw, fallback) {
  if (raw == null || raw === '') return fallback;
  try {
    var parsed = (typeof raw === 'string') ? JSON.parse(raw) : raw;
    return parsed == null ? fallback : parsed;
  } catch (e) {
    return fallback;
  }
}

function _lessonEvalActor_() {
  try {
    var email = String(Session.getActiveUser().getEmail() || '').trim();
    if (email) return email;
  } catch (e0) {}
  try {
    var fallback = String(Session.getEffectiveUser().getEmail() || '').trim();
    if (fallback) return fallback;
  } catch (e1) {}
  return 'system';
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

function _flightFollowArchiveFolder_() {
  var folderName = 'FF_Archived_Snapshots';
  var it = DriveApp.getFoldersByName(folderName);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(folderName);
}

function _flightFollowArchiveSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Active spreadsheet not available for FF archive.');
  var name = APP_SHEETS.FF_ARCHIVE || 'DB_FF_Archive';
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  var lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    sheet.getRange(1, 1, 1, FF_ARCHIVE_HEADERS_.length).setValues([FF_ARCHIVE_HEADERS_]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  var currentHeaders = sheet.getRange(1, 1, 1, FF_ARCHIVE_HEADERS_.length).getValues()[0].map(function(h) {
    return String(h || '').trim().toUpperCase();
  });
  var expectedHeaders = FF_ARCHIVE_HEADERS_.map(function(h) { return String(h || '').trim().toUpperCase(); });
  var mismatch = false;
  for (var i = 0; i < expectedHeaders.length; i++) {
    if (currentHeaders[i] !== expectedHeaders[i]) { mismatch = true; break; }
  }
  if (mismatch) {
    sheet.getRange(1, 1, 1, FF_ARCHIVE_HEADERS_.length).setValues([FF_ARCHIVE_HEADERS_]);
  }
  return sheet;
}

function _flightFollowSpreadsheet_() {
  var configuredId = '';
  try {
    configuredId = String(PropertiesService.getScriptProperties().getProperty(FF_MESSAGES_SPREADSHEET_ID_PROP_) || '').trim();
  } catch (e0) {}

  var targetId = configuredId || String(FF_MESSAGES_SPREADSHEET_ID_DEFAULT_ || '').trim();

  if (targetId) {
    try {
      return SpreadsheetApp.openById(targetId);
    } catch (e1) {
      throw new Error('Flight Follow spreadsheet ID is configured/defaulted but inaccessible: ' + targetId + ' | ' + String(e1 && e1.message || e1));
    }
  }

  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e2) {}

  throw new Error('Flight Follow spreadsheet unavailable. Set script property ' + FF_MESSAGES_SPREADSHEET_ID_PROP_ + ' to a shared spreadsheet ID.');
}

function _flightFollowNormalizeReg_(value) {
  return String(value || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function _flightFollowTokenVariants_(regNorm) {
  var v = [];
  var norm = _flightFollowNormalizeReg_(regNorm);
  if (!norm) return v;
  v.push(norm);
  if (norm.length >= 5) v.push(norm.slice(0, 2) + '-' + norm.slice(2));
  return v;
}

function _flightFollowContainsAnyToken_(text, tokens) {
  var hay = String(text || '').toUpperCase();
  if (!hay || !tokens || !tokens.length) return false;
  for (var i = 0; i < tokens.length; i++) {
    if (hay.indexOf(String(tokens[i] || '').toUpperCase()) >= 0) return true;
  }
  return false;
}

function _flightFollowLooksInreach_(sender) {
  var s = String(sender || '').toLowerCase();
  return s.indexOf('inreachmail.com') >= 0 || s.indexOf('@garmin.com') >= 0;
}

function _flightFollowInferRegFromText_(text) {
  var upper = String(text || '').toUpperCase();
  if (!upper) return '';
  var m = upper.match(/\b(?:PT|PR|PP|PU)-?[A-Z0-9]{3,4}\b/);
  return m && m[0] ? String(m[0]).replace(/[^A-Z0-9]/g, '') : '';
}

function _flightFollowMessagesSheet_() {
  var ss = _flightFollowSpreadsheet_();
  var name = APP_SHEETS.FF_MESSAGES || 'DB_FF_Messages';
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    sheet.getRange(1, 1, 1, FF_MESSAGES_HEADERS_.length).setValues([FF_MESSAGES_HEADERS_]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  var currentHeaders = sheet.getRange(1, 1, 1, FF_MESSAGES_HEADERS_.length).getValues()[0].map(function(h) {
    return String(h || '').trim().toUpperCase();
  });
  var expectedHeaders = FF_MESSAGES_HEADERS_.map(function(h) { return String(h || '').trim().toUpperCase(); });
  var mismatch = false;
  for (var i = 0; i < expectedHeaders.length; i++) {
    if (currentHeaders[i] !== expectedHeaders[i]) {
      mismatch = true;
      break;
    }
  }
  if (mismatch) {
    sheet.getRange(1, 1, 1, FF_MESSAGES_HEADERS_.length).setValues([FF_MESSAGES_HEADERS_]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _flightFollowThreadRegMap_(sheet, maxRows) {
  var targetSheet = sheet || _flightFollowMessagesSheet_();
  var lastRow = targetSheet.getLastRow();
  if (lastRow < 2) return {};

  var limit = Number(maxRows || 2000);
  if (!isFinite(limit) || limit < 1) limit = 2000;
  var startRow = Math.max(2, lastRow - limit + 1);
  var values = targetSheet.getRange(startRow, 1, lastRow - startRow + 1, FF_MESSAGES_HEADERS_.length).getValues();
  var out = {};

  values.forEach(function(r) {
    var regNorm = _flightFollowNormalizeReg_(r[2]);
    var threadId = String(r[9] || '').trim();
    if (!regNorm || !threadId || out[threadId]) return;
    out[threadId] = regNorm;
  });

  return out;
}

function _flightFollowAppendMessages_(messages) {
  var rows = Array.isArray(messages) ? messages : [];
  if (!rows.length) return 0;
  var sheet = _flightFollowMessagesSheet_();
  var nowIso = new Date().toISOString();
  var threadRegMap = _flightFollowThreadRegMap_(sheet, 2500);
  var values = rows.map(function(m) {
    var ts = m.timestamp ? new Date(m.timestamp) : new Date();
    var tsIso = isNaN(ts.getTime()) ? nowIso : ts.toISOString();
    var threadId = String(m.threadId || '').trim();
    var regNorm = _flightFollowNormalizeReg_(m.regNorm || m.reg || threadRegMap[threadId] || _flightFollowInferRegFromText_([
      m.subject,
      m.text,
      m.from,
      m.to
    ].join('\n')));
    return [
      tsIso,
      regNorm ? (regNorm.length >= 5 ? regNorm.slice(0, 2) + '-' + regNorm.slice(2) : regNorm) : '',
      regNorm,
      String(m.direction || 'INBOUND').trim().toUpperCase(),
      String(m.from || '').trim(),
      String(m.to || '').trim(),
      String(m.subject || '').trim(),
      String(m.text || '').trim().slice(0, 8000),
      m.verified ? 'TRUE' : 'FALSE',
      threadId,
      String(m.messageId || '').trim(),
      String(m.source || '').trim() || 'FF_SYNC',
      nowIso
    ];
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, FF_MESSAGES_HEADERS_.length).setValues(values);
  return values.length;
}

function _flightFollowOutboundRecipientList_() {
  // Parse outbound map and return unique email targets.
  var out = [];
  try {
    var ss = _flightFollowSpreadsheet_();
    var sheet = ss.getSheetByName(APP_SHEETS.SCHED_CONFIG || 'SCHED_Config');
    if (!sheet) return out;
    var cfg = _schedulerReadConfigMap_(sheet);
    var raw = String(cfg.FLIGHT_FOLLOW_OUTBOUND_MAP || '').trim();
    if (!raw) return out;

    var seen = {};
    raw.split(/\r?\n/).forEach(function(line) {
      var cleaned = String(line || '').trim();
      if (!cleaned || cleaned.indexOf('#') === 0) return;
      var parts = cleaned.split(/[=:]/);
      if (parts.length < 2) return;
      var addr = String(parts.slice(1).join(':') || '').trim().toLowerCase();
      if (!addr || seen[addr]) return;
      seen[addr] = true;
      out.push(addr);
    });
  } catch (e) {}
  return out;
}

function _flightFollowRecipientsList_() {
  var fallback = (FF_RECIPIENTS_ || []).map(function(v) {
    return String(v || '').trim().toLowerCase();
  }).filter(Boolean);

  try {
    var ss = _flightFollowSpreadsheet_();
    var sheet = ss.getSheetByName(APP_SHEETS.SCHED_CONFIG || 'SCHED_Config');
    if (!sheet) return fallback;
    var cfg = _schedulerReadConfigMap_(sheet);
    var raw = String(cfg.FLIGHT_FOLLOW_RECIPIENTS || '').trim();
    if (!raw) return fallback;

    var list = raw.split(/[;,\n]+/).map(function(v) {
      return String(v || '').trim().toLowerCase();
    }).filter(Boolean);
    if (!list.length) return fallback;

    var seen = {};
    return list.filter(function(email) {
      if (seen[email]) return false;
      seen[email] = true;
      return true;
    });
  } catch (e) {
    return fallback;
  }
}

function _flightFollowPrimaryInbox_() {
  var list = _flightFollowRecipientsList_();
  return list.length ? list[0] : 'acompanhamento@asasdesocorro.org.br';
}

function _flightFollowOutboundRecipientForReg_(reg) {
  var regNorm = String(reg || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (!regNorm) return '';
  try {
    var ss = _flightFollowSpreadsheet_();
    var sheet = ss.getSheetByName(APP_SHEETS.SCHED_CONFIG || 'SCHED_Config');
    if (!sheet) return '';
    var cfg = _schedulerReadConfigMap_(sheet);
    var raw = String(cfg.FLIGHT_FOLLOW_OUTBOUND_MAP || '').trim();
    if (!raw) return '';

    var lines = raw.split(/\r?\n/);
    for (var i = 0; i < lines.length; i++) {
      var line = String(lines[i] || '').trim();
      if (!line || line.indexOf('#') === 0) continue;
      var parts = line.split(/[=:]/);
      if (parts.length < 2) continue;
      var key = String(parts[0] || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
      var value = String(parts.slice(1).join(':') || '').trim().toLowerCase();
      if (!key || !value) continue;
      if (key === regNorm) return value;
    }
    return '';
  } catch (e) {
    return '';
  }
}

function _flightFollowScanMessages_(reg, debugMode) {
  // Read inbox replies arriving to Flight Following and keep only messages tied to this aircraft.
  // Returns messages list or debug object when debugMode=true.
  var regNorm = String(reg || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (!regNorm) {
    return debugMode ? {
      reg: '',
      regNorm: '',
      inboxAddress: _flightFollowPrimaryInbox_(),
      search: '',
      effectiveUser: Session.getEffectiveUser().getEmail(),
      activeUser: Session.getActiveUser().getEmail(),
      threadCount: 0,
      messageCount: 0,
      matchedCount: 0,
      samples: [],
      note: 'Missing reg parameter'
    } : [];
  }

  var inboxAddress = _flightFollowPrimaryInbox_();
  var outboundAddress = String(_flightFollowOutboundRecipientForReg_(reg) || '').trim().toLowerCase();
  var dashedReg = regNorm.length >= 5 ? (regNorm.slice(0, 2) + '-' + regNorm.slice(2)) : regNorm;
  var search = [
    'in:inbox',
    'newer_than:30d'
  ].join(' ');
  var fallbackSearch = [
    'in:anywhere',
    'newer_than:30d',
    '("' + regNorm + '" OR "' + dashedReg + '")'
  ].join(' ');
  var outboundFallbackSearch = outboundAddress ? [
    'in:anywhere',
    'newer_than:30d',
    '"' + outboundAddress + '"'
  ].join(' ') : '';

  var tokenVariants = [regNorm];
  // Add dashed variant (e.g., PTABC) -> PT-ABC for easier text matching.
  if (regNorm.length >= 5) {
    tokenVariants.push(regNorm.slice(0, 2) + '-' + regNorm.slice(2));
  }

  var containsReg_ = function(text) {
    var t = String(text || '').toUpperCase();
    if (!t) return false;
    return tokenVariants.some(function(tok) {
      return t.indexOf(tok) >= 0;
    });
  };

  var looksInreach_ = function(sender) {
    var s = String(sender || '').toLowerCase();
    return s.indexOf('inreachmail.com') >= 0 || s.indexOf('@garmin.com') >= 0;
  };

  try {
    var threads = GmailApp.search(search, 0, 120);
    var threadSeen = {};
    threads.forEach(function(t) {
      threadSeen[t.getId()] = true;
    });

    // Fallback: include reg-linked threads from anywhere (including Sent) so replies
    // can still be discovered even when inbox first-message subjects are noisy.
    var fallbackThreads = GmailApp.search(fallbackSearch, 0, 120);
    fallbackThreads.forEach(function(t) {
      if (threadSeen[t.getId()]) return;
      threads.push(t);
      threadSeen[t.getId()] = true;
    });

    if (outboundFallbackSearch) {
      var outboundThreads = GmailApp.search(outboundFallbackSearch, 0, 120);
      outboundThreads.forEach(function(t) {
        if (threadSeen[t.getId()]) return;
        threads.push(t);
        threadSeen[t.getId()] = true;
      });
    }

    var out = [];
    var sample = [];
    var msgCount = 0;
    var sourceThreadCount = threads.length;

    threads.forEach(function(thread) {
      var subject = String(thread.getFirstMessageSubject() || '');
      var threadHasReg = containsReg_(subject);
      var msgs = thread.getMessages();
      var threadHasOutboundAddress = false;
      if (outboundAddress) {
        for (var mi = 0; mi < msgs.length; mi++) {
          var tm = msgs[mi];
          var tFrom = String(tm.getFrom() || '').toLowerCase();
          var tTo = String(tm.getTo() || '').toLowerCase();
          if (tFrom.indexOf(outboundAddress) >= 0 || tTo.indexOf(outboundAddress) >= 0) {
            threadHasOutboundAddress = true;
            break;
          }
        }
      }
      msgs.forEach(function(msg) {
        if (!msg.isInInbox()) return;
        msgCount++;

        var from = String(msg.getFrom() || '');
        var subj = String(msg.getSubject() || '');
        var body = String(msg.getPlainBody() || '');
        var to = String(msg.getTo() || '');
        var toNormLower = String(to || '').toLowerCase();
        var fromNormLower = String(from || '').toLowerCase();
        var addressMatched = !!outboundAddress && (
          toNormLower.indexOf(outboundAddress) >= 0 ||
          fromNormLower.indexOf(outboundAddress) >= 0 ||
          String(subj || '').toLowerCase().indexOf(outboundAddress) >= 0 ||
          String(body || '').toLowerCase().indexOf(outboundAddress) >= 0
        );

        // Match by explicit reg in subject/body or in a reg-matched thread.
        var matched = threadHasReg || containsReg_(subj) || containsReg_(body) || containsReg_(to) || containsReg_(from) || threadHasOutboundAddress || addressMatched;
        if (!matched) return;

        // Prefer FF mailbox routing but still allow reg-matched replies that lost explicit To headers.
        var isToPrimaryInbox = !!inboxAddress && toNormLower.indexOf(String(inboxAddress || '').toLowerCase()) >= 0;
        var isLikelyReplyInThread = threadHasReg || threadHasOutboundAddress || addressMatched;
        if (!isToPrimaryInbox && !isLikelyReplyInThread) return;

        if (debugMode && sample.length < 30) {
          sample.push({
            date: msg.getDate().toISOString(),
            from: from,
            to: to,
            subject: subj,
            threadHasReg: threadHasReg,
            threadHasOutboundAddress: threadHasOutboundAddress,
            outboundAddressMatched: addressMatched,
            isToPrimaryInbox: isToPrimaryInbox,
            bodyPreview: body.substring(0, 180)
          });
        }

        out.push({
          timestamp: msg.getDate().toISOString(),
          from: from,
          subject: subj,
          text: body.substring(0, 1200),
          verified: looksInreach_(from)
        });
      });
    });

    out.sort(function(a, b) {
      return String(a.timestamp || '').localeCompare(String(b.timestamp || ''));
    });

    if (debugMode) {
      return {
        reg: reg,
        regNorm: regNorm,
        tokenVariants: tokenVariants,
        inboxAddress: inboxAddress,
        outboundAddress: outboundAddress,
        search: search,
        fallbackSearch: fallbackSearch,
        outboundFallbackSearch: outboundFallbackSearch,
        effectiveUser: Session.getEffectiveUser().getEmail(),
        activeUser: Session.getActiveUser().getEmail(),
        threadCount: sourceThreadCount,
        messageCount: msgCount,
        matchedCount: out.length,
        samples: sample
      };
    }

    return out.slice(-200);
  } catch (e) {
    return debugMode ? {
      reg: reg,
      regNorm: regNorm,
      inboxAddress: inboxAddress,
      outboundAddress: outboundAddress,
      search: search,
    fallbackSearch: fallbackSearch,
      outboundFallbackSearch: outboundFallbackSearch,
      effectiveUser: Session.getEffectiveUser().getEmail(),
      activeUser: Session.getActiveUser().getEmail(),
      threadCount: 0,
      messageCount: 0,
      matchedCount: 0,
      samples: [],
      error: String(e && e.message || e)
    } : [];
  }
}

function _flightFollowParseKmlDocument_(root, context) {
  context = context || {};
  var sourceUrl = String(context.sourceUrl || '').trim();
  var depth = Number(context.depth || 0);
  var seen = context.seen || {};
  var ns = XmlService.getNamespace('http://www.opengis.net/kml/2.2');
  var gxNs = XmlService.getNamespace('http://www.google.com/kml/ext/2.2');
  var out = {
    title: '',
    description: '',
    placemarks: [],
    summary: {
      placemarkCount: 0,
      featureCount: 0,
      pointCount: 0,
      lineCount: 0,
      trackCount: 0,
      polygonCount: 0,
      coordinateCount: 0
    },
    bounds: null
  };

  if (!root) return out;

  out.title = _flightFollowKmlChildText_(root, 'name', ns);
  out.description = _flightFollowKmlChildText_(root, 'description', ns);

  function visit(container, containerUrl, containerDepth) {
    if (!container) return;
    var placemarks = container.getChildren('Placemark', ns) || [];
    for (var i = 0; i < placemarks.length; i++) {
      var parsedPlacemark = _flightFollowParseKmlPlacemark_(placemarks[i], ns);
      if (parsedPlacemark) {
        out.placemarks.push(parsedPlacemark);
        out.summary.placemarkCount += 1;
        out.summary.featureCount += parsedPlacemark.features.length;
        for (var j = 0; j < parsedPlacemark.features.length; j++) {
          var feature = parsedPlacemark.features[j];
          if (feature.type === 'Point') out.summary.pointCount += 1;
          if (feature.type === 'LineString') out.summary.lineCount += 1;
          if (feature.type === 'Track') out.summary.trackCount += 1;
          if (feature.type === 'Polygon') out.summary.polygonCount += 1;
          out.summary.coordinateCount += feature.coordinates.length;
          for (var k = 0; k < feature.coordinates.length; k++) {
            out.bounds = _flightFollowExtendBounds_(out.bounds, feature.coordinates[k]);
          }
        }
      }
    }

    if (containerDepth < 2) {
      var networkLinks = container.getChildren('NetworkLink', ns) || [];
      for (var n = 0; n < networkLinks.length; n++) {
        var networkLink = networkLinks[n];
        var href = '';
        try {
          var linkNode = networkLink.getChild('Link', ns) || networkLink.getChild('Url', ns);
          if (linkNode) href = _flightFollowKmlChildText_(linkNode, 'href', ns) || _flightFollowKmlChildText_(linkNode, 'url', ns);
        } catch (e) {}
        href = _flightFollowResolveKmlUrl_(href, containerUrl || sourceUrl);
        if (!href || seen[href]) continue;
        seen[href] = true;
        try {
          var linkedResp = UrlFetchApp.fetch(href, {
            followRedirects: true,
            muteHttpExceptions: true,
            headers: {
              Accept: 'application/vnd.google-earth.kml+xml, application/xml;q=0.9, text/xml;q=0.8, */*;q=0.1'
            }
          });
          if (linkedResp.getResponseCode() >= 200 && linkedResp.getResponseCode() < 300) {
            var linkedText = String(linkedResp.getContentText() || '').trim();
            if (linkedText) {
              var linkedDoc = XmlService.parse(linkedText);
              var linkedParsed = _flightFollowParseKmlDocument_(linkedDoc.getRootElement(), {
                sourceUrl: href,
                depth: containerDepth + 1,
                seen: seen
              });
              if (linkedParsed && linkedParsed.placemarks && linkedParsed.placemarks.length) {
                for (var lp = 0; lp < linkedParsed.placemarks.length; lp++) out.placemarks.push(linkedParsed.placemarks[lp]);
                out.summary.placemarkCount += linkedParsed.summary.placemarkCount;
                out.summary.featureCount += linkedParsed.summary.featureCount;
                out.summary.pointCount += linkedParsed.summary.pointCount;
                out.summary.lineCount += linkedParsed.summary.lineCount;
                out.summary.trackCount += linkedParsed.summary.trackCount || 0;
                out.summary.polygonCount += linkedParsed.summary.polygonCount;
                out.summary.coordinateCount += linkedParsed.summary.coordinateCount;
                out.bounds = _flightFollowMergeBounds_(out.bounds, linkedParsed.bounds);
              }
            }
          }
        } catch (e) {}
      }
    }

    var folders = container.getChildren('Folder', ns) || [];
    for (var f = 0; f < folders.length; f++) visit(folders[f], containerUrl, containerDepth);
    var documents = container.getChildren('Document', ns) || [];
    for (var d = 0; d < documents.length; d++) visit(documents[d], containerUrl, containerDepth);
  }

  visit(root, sourceUrl, depth);
  return out;
}

function _flightFollowParseKmlPlacemark_(placemark, ns) {
  if (!placemark) return null;
  var placemarkPointMeta = _flightFollowParseKmlPointMeta_(placemark, ns);
  var out = {
    name: _flightFollowKmlChildText_(placemark, 'name', ns),
    description: _flightFollowKmlChildText_(placemark, 'description', ns),
    features: []
  };

  var geometries = placemark.getChildren() || [];
  for (var i = 0; i < geometries.length; i++) {
    _flightFollowParseKmlGeometry_(geometries[i], ns, out.features, placemarkPointMeta);
  }

  if (!out.features.length) return null;
  return out;
}

function _flightFollowParseKmlGeometry_(node, ns, features, inheritedPointMeta) {
  if (!node || !features) return;
  var name = String(node.getName() || '').trim();
  if (name === 'Point') {
    var pointCoords = _flightFollowParseKmlCoordinates_(_flightFollowKmlChildText_(node, 'coordinates', ns));
    if (pointCoords.length) {
      var pointMeta = _flightFollowParseKmlPointMeta_(node, ns);
      if (inheritedPointMeta) {
        if (!pointMeta.when && inheritedPointMeta.when) pointMeta.when = inheritedPointMeta.when;
        if (inheritedPointMeta.raw) {
          for (var inheritedKey in inheritedPointMeta.raw) {
            if (!Object.prototype.hasOwnProperty.call(inheritedPointMeta.raw, inheritedKey)) continue;
            if (!Object.prototype.hasOwnProperty.call(pointMeta.raw, inheritedKey) || !pointMeta.raw[inheritedKey]) {
              pointMeta.raw[inheritedKey] = inheritedPointMeta.raw[inheritedKey];
            }
          }
        }
        if (isNaN(pointMeta.altitudeM) && isFinite(Number(inheritedPointMeta.altitudeM))) pointMeta.altitudeM = Number(inheritedPointMeta.altitudeM);
        if (isNaN(pointMeta.speedKts) && isFinite(Number(inheritedPointMeta.speedKts))) pointMeta.speedKts = Number(inheritedPointMeta.speedKts);
        if (isNaN(pointMeta.headingDeg) && isFinite(Number(inheritedPointMeta.headingDeg))) pointMeta.headingDeg = Number(inheritedPointMeta.headingDeg);
      }
      features.push({
        type: 'Point',
        coordinates: [pointCoords[0]],
        rawCoordinates: pointCoords,
        point: pointMeta
      });
    }
    return;
  }
  if (name === 'LineString') {
    var lineCoords = _flightFollowParseKmlCoordinates_(_flightFollowKmlChildText_(node, 'coordinates', ns));
    if (lineCoords.length) {
      features.push({ type: 'LineString', coordinates: lineCoords });
    }
    return;
  }
  if (name === 'Track') {
    var trackPoints = [];
    var pendingWhen = '';
    var children = node.getChildren() || [];
    for (var t = 0; t < children.length; t++) {
      var child = children[t];
      var childName = String(child && child.getName() || '').trim();
      if (childName === 'when') {
        pendingWhen = String(child.getText() || '').trim();
        continue;
      }
      if (childName !== 'coord') continue;
      var coordText = String(child.getText() || '').trim();
      if (!coordText) continue;
      var coordParts = coordText.split(/\s+/);
      if (coordParts.length < 2) continue;
      var lon = Number(coordParts[0]);
      var lat = Number(coordParts[1]);
      var alt = coordParts.length > 2 ? Number(coordParts[2]) : NaN;
      if (!isFinite(lat) || !isFinite(lon)) continue;
      trackPoints.push({
        lat: lat,
        lon: lon,
        altM: isFinite(alt) ? alt : NaN,
        when: pendingWhen
      });
      pendingWhen = '';
    }
    if (trackPoints.length) {
      features.push({
        type: 'Track',
        coordinates: trackPoints.map(function(point) {
          return isFinite(point.altM) ? [point.lat, point.lon, point.altM] : [point.lat, point.lon];
        }),
        points: trackPoints
      });
    }
    return;
  }
  if (name === 'Polygon') {
    var outer = node.getChild('outerBoundaryIs', ns);
    var ring = outer && outer.getChild('LinearRing', ns);
    var polyCoords = ring ? _flightFollowParseKmlCoordinates_(_flightFollowKmlChildText_(ring, 'coordinates', ns)) : [];
    if (polyCoords.length) {
      features.push({ type: 'Polygon', coordinates: polyCoords });
    }
    return;
  }
  if (name === 'MultiGeometry' || name === 'MultiTrack') {
    var children = node.getChildren() || [];
    for (var c = 0; c < children.length; c++) {
      _flightFollowParseKmlGeometry_(children[c], ns, features, inheritedPointMeta);
    }
  }
}

function _flightFollowParseKmlPointMeta_(node, ns) {
  var meta = {
    when: '',
    altitudeM: NaN,
    speedKts: NaN,
    headingDeg: NaN,
    raw: {}
  };
  if (!node) return meta;
  try {
    var whenNode = node.getChild('TimeStamp', ns);
    if (whenNode) meta.when = _flightFollowKmlChildText_(whenNode, 'when', ns);
    if (!meta.when) {
      whenNode = node.getChild('when', ns);
      if (whenNode) meta.when = String(whenNode.getText() || '').trim();
    }
  } catch (e) {}

  var ext = null;
  try {
    ext = node.getChild('ExtendedData', ns);
  } catch (e) {}
  if (ext) {
    var dataNodes = ext.getChildren('Data', ns) || [];
    for (var i = 0; i < dataNodes.length; i++) {
      var dataNode = dataNodes[i];
      var key = String(dataNode && dataNode.getAttribute('name') ? dataNode.getAttribute('name').getValue() : '').trim();
      var value = String(_flightFollowKmlChildText_(dataNode, 'value', ns) || dataNode.getText() || '').trim();
      if (!key) continue;
      var normKey = key.toLowerCase().replace(/[^a-z0-9]+/g, '');
      meta.raw[normKey] = value;
    }
    try {
      var schemaDataNodes = ext.getChildren('SchemaData', ns) || [];
      for (var s = 0; s < schemaDataNodes.length; s++) {
        var schemaData = schemaDataNodes[s];
        var simpleNodes = schemaData.getChildren('SimpleData', ns) || [];
        for (var j = 0; j < simpleNodes.length; j++) {
          var simpleNode = simpleNodes[j];
          var simpleKey = String(simpleNode && simpleNode.getAttribute('name') ? simpleNode.getAttribute('name').getValue() : '').trim();
          var simpleValue = String(simpleNode.getText() || '').trim();
          if (!simpleKey) continue;
          meta.raw[simpleKey.toLowerCase().replace(/[^a-z0-9]+/g, '')] = simpleValue;
        }
      }
    } catch (e) {}
  }

  if (!meta.when) {
    meta.when = String(
      meta.raw.timeutc ||
      meta.raw.time ||
      meta.raw.timestamp ||
      meta.raw.whenu ||
      ''
    ).trim();
  }

  var altitudeRaw = meta.raw.altitudeft || meta.raw.altitudem || meta.raw.altitude || meta.raw.alt || meta.raw.elevation || '';
  var speedRaw = meta.raw.speedkts || meta.raw.groundspeedkts || meta.raw.speed || meta.raw.groundspeed || '';
  var headingRaw = meta.raw.headingdeg || meta.raw.heading || meta.raw.course || '';
  var altitudeM = Number(altitudeRaw);
  if (isFinite(altitudeM) && Math.abs(altitudeM) > 1000 && String(altitudeRaw).indexOf('ft') >= 0) {
    altitudeM = altitudeM / 3.28084;
  }
  if (isFinite(altitudeM)) meta.altitudeM = altitudeM;
  var speedKts = Number(speedRaw);
  if (isFinite(speedKts)) meta.speedKts = speedKts;
  var headingDeg = Number(headingRaw);
  if (isFinite(headingDeg)) meta.headingDeg = headingDeg;
  return meta;
}

function _flightFollowResolveKmlUrl_(href, sourceUrl) {
  href = String(href || '').trim();
  if (!href) return '';
  if (/^https?:\/\//i.test(href)) return href;
  sourceUrl = String(sourceUrl || '').trim();
  if (!sourceUrl) return href;
  var originMatch = sourceUrl.match(/^(https?:\/\/[^\/]+)(\/.*)?$/i);
  if (!originMatch) return href;
  var origin = originMatch[1];
  if (href.charAt(0) === '/') return origin + href;
  var basePath = (originMatch[2] || '/').replace(/\/[^\/]*$/, '/');
  return origin + basePath + href.replace(/^\.\//, '');
}

function _flightFollowMergeBounds_(left, right) {
  if (!right) return left || null;
  if (!left) {
    return {
      minLat: right.minLat,
      maxLat: right.maxLat,
      minLon: right.minLon,
      maxLon: right.maxLon
    };
  }
  if (right.minLat < left.minLat) left.minLat = right.minLat;
  if (right.maxLat > left.maxLat) left.maxLat = right.maxLat;
  if (right.minLon < left.minLon) left.minLon = right.minLon;
  if (right.maxLon > left.maxLon) left.maxLon = right.maxLon;
  return left;
}

function _flightFollowParseKmlCoordinates_(raw) {
  var coords = [];
  String(raw || '').trim().split(/\s+/).forEach(function(chunk) {
    var cleaned = String(chunk || '').trim();
    if (!cleaned) return;
    var parts = cleaned.split(',');
    if (parts.length < 2) return;
    var lon = Number(parts[0]);
    var lat = Number(parts[1]);
    var alt = parts.length > 2 ? Number(parts[2]) : NaN;
    if (!isFinite(lat) || !isFinite(lon)) return;
    if (isFinite(alt)) coords.push([lat, lon, alt]);
    else coords.push([lat, lon]);
  });
  return coords;
}

function _flightFollowKmlChildText_(node, childName, ns) {
  if (!node) return '';
  try {
    var child = node.getChild(childName, ns);
    if (child) return String(child.getText() || '').trim();
  } catch (e) {}
  return '';
}

function _flightFollowExtendBounds_(bounds, coord) {
  if (!coord || !coord.length) return bounds || null;
  var lat = Number(coord[0]);
  var lon = Number(coord[1]);
  if (!isFinite(lat) || !isFinite(lon)) return bounds || null;
  if (!bounds) {
    return { minLat: lat, maxLat: lat, minLon: lon, maxLon: lon };
  }
  if (lat < bounds.minLat) bounds.minLat = lat;
  if (lat > bounds.maxLat) bounds.maxLat = lat;
  if (lon < bounds.minLon) bounds.minLon = lon;
  if (lon > bounds.maxLon) bounds.maxLon = lon;
  return bounds;
}

function _debr8TrainingStaffRecord_(ss, pilotName) {
  const target = String(pilotName || '').trim().toUpperCase();
  if (!target) return null;
  const staffSheet = ss.getSheetByName(APP_SHEETS.PILOTS);
  if (!staffSheet || staffSheet.getLastRow() < 2) return null;
  const staffData = staffSheet.getDataRange().getValues();
  const staffHeaders = staffData[0] || [];
  for (let i = 1; i < staffData.length; i++) {
    const staff = _toolsStaffRecordFromRow_(staffHeaders, staffData[i], i + 1);
    const nameKey = String(staff.staffName || '').trim().toUpperCase();
    if (nameKey === target) return staff;
  }
  return null;
}

function _debr8AppendTrainingCheckouts_(ss, flightLegId, pilotName, trainingDebrief) {
  const checksSheet = _pilotRunwayChecksSheet_();
  const headers = _toolsSheetHeaderRow_(checksSheet);
  const rows = checksSheet.getDataRange().getValues();
  const traineeName = String(trainingDebrief && trainingDebrief.traineeName || pilotName || '').trim();
  const instructorName = String(trainingDebrief && trainingDebrief.instructorName || '').trim();
  const normalizedPilot = traineeName.toUpperCase();
  const trainingCode = String(trainingDebrief && trainingDebrief.trainingCode || '').trim().toUpperCase();
  const sourceTag = 'TRAINING_CHECKOUT_DEBRIEF';
  const today = safeDateStr(new Date());
  const staff = _debr8TrainingStaffRecord_(ss, traineeName);

  const candidateChecks = Array.isArray(trainingDebrief && trainingDebrief.runwayChecks)
    ? trainingDebrief.runwayChecks
    : [];

  let appended = 0;
  candidateChecks.forEach(function(item) {
    const icao = _checksNormalizeIcaoToken_(item && item.icao);
    const successful = !!(item && item.successful);
    const landings = Math.max(0, parseInt(item && item.landings, 10) || 0);
    if (!icao || !successful || landings <= 0) return;

    const already = rows.slice(1).some(function(r, idx) {
      const rec = _pilotRunwayCheckRecordFromRow_(headers, r, idx + 2);
      if (String(rec.pilotName || '').trim().toUpperCase() !== normalizedPilot) return false;
      if (rec.icaos.indexOf(icao) === -1) return false;
      if (String(rec.source || '').trim().toUpperCase() !== sourceTag) return false;
      return String(rec.notes || '').indexOf(flightLegId) >= 0;
    });
    if (already) return;

    const note = [
      'Training checkout',
      trainingCode ? ('Code ' + trainingCode) : '',
      'Flight ' + flightLegId,
      'Landings ' + landings
    ].filter(Boolean).join(' | ');

    const dataMap = {
      CHECK_ID: 'CHK_' + new Date().getTime() + '_' + icao,
      PILOT_NAME: traineeName,
      PILOT_EMAIL: staff ? String(staff.email || '').trim().toLowerCase() : '',
      STAFF_ID: staff ? String(staff.staffId || '').trim() : '',
      ICAO: icao,
      RUNWAY_IDENT: 'ALL',
      AUTH_SCOPE: 'AIRPORT',
      STATUS: 'ACTIVE',
      DATE_CHECKED: today,
      EXPIRY_DATE: '',
      APPROVED_BY: instructorName || traineeName,
      SOURCE: sourceTag,
      NOTES: note,
      CREATED_AT: today,
      UPDATED_AT: today
    };

    const rowOut = headers.map(function(header) {
      const key = _toolsNormHeader_(header);
      return Object.prototype.hasOwnProperty.call(dataMap, key) ? dataMap[key] : '';
    });
    checksSheet.appendRow(rowOut);
    appended += 1;
  });

  return appended;
}

  function _dutyHasEntryForPilotDate_(ss, pilotName, ymd) {
    var sh = ss.getSheetByName(APP_SHEETS.DUTY_LOG);
    if (!sh) return false;
    var pilotTarget = _dutyNormalizePilotName_(pilotName).toUpperCase();
    if (!pilotTarget || !ymd) return false;

    var rows = sh.getDataRange().getValues();
    if (!rows || rows.length < 2) return false;

    for (var i = 1; i < rows.length; i++) {
      var rowDate = _dutyYmd_(rows[i][DUTY_LOG_COL.DATE]);
      var rowPilot = String(rows[i][DUTY_LOG_COL.PILOT] || '').trim().toUpperCase();
      if (rowDate === ymd && rowPilot === pilotTarget) return true;
    }
    return false;
  }

  function _dutyDepartmentForPilotDate_(ss, pilotName, ymd) {
    var sh = ss.getSheetByName(APP_SHEETS.DUTY_LOG);
    if (!sh) return '';
    var pilotTarget = _dutyNormalizePilotName_(pilotName).toUpperCase();
    if (!pilotTarget || !ymd) return '';

    var rows = sh.getDataRange().getValues();
    if (!rows || rows.length < 2) return '';

    for (var i = rows.length - 1; i >= 1; i--) {
      var rowDate = _dutyYmd_(rows[i][DUTY_LOG_COL.DATE]);
      var rowPilot = String(rows[i][DUTY_LOG_COL.PILOT] || '').trim().toUpperCase();
      if (rowDate !== ymd || rowPilot !== pilotTarget) continue;

      var jsonRaw = String(rows[i][5] || '').trim();
      if (!jsonRaw) break;
      try {
        var parsed = JSON.parse(jsonRaw);
        var department = String(parsed && parsed.department || '').trim();
        if (department) return department;
      } catch (e) {}
      break;
    }

    return '';
  }

  function _dutyCollectOnboardAlumnoNamesFromLoadJson_(loadJsonRaw, primaryPilotName) {
    var names = [];
    var seen = {};
    var payload = {};
    try {
      payload = loadJsonRaw ? JSON.parse(String(loadJsonRaw)) : {};
    } catch (e) {
      payload = {};
    }

    var seats = payload && payload.seats && typeof payload.seats === 'object' ? payload.seats : null;
    if (!seats) return names;

    var primaryNorm = _dutyPilotNorm_(primaryPilotName);
    var copilotName = String(payload && (payload.copilot || payload.copilotName || (payload.crew && payload.crew.copilot)) || '').trim();
    if (copilotName) {
      var resolvedCopilot = _resolveDutyCrewNameFromPilotsSheet_(copilotName);
      var copilotNorm = _dutyPilotNorm_(resolvedCopilot || copilotName);
      if (copilotNorm && copilotNorm !== primaryNorm && !seen[copilotNorm]) {
        seen[copilotNorm] = true;
        names.push(resolvedCopilot || copilotName);
      }
    }
    for (var seatKey in seats) {
      var seat = seats[seatKey];
      var passengerName = String(seat && seat.passenger && seat.passenger.name || '').trim();
      if (!passengerName) continue;
      if (primaryNorm && _dutyPilotNorm_(passengerName) === primaryNorm) continue;

      var resolved = _resolveDutyCrewNameFromPilotsSheet_(passengerName) || _resolveBillableAlunoNameFromPilotsSheet_(passengerName);
      if (!resolved) continue;

      var resolvedNorm = _dutyPilotNorm_(resolved);
      if (!resolvedNorm || seen[resolvedNorm]) continue;
      seen[resolvedNorm] = true;
      names.push(resolved);
    }

    return names;
  }

  function _dutyPersistOnboardAlunoRowsFromFlight_(ss, flightRow, flightLegId) {
    var flightDateYmd = _dutyYmd_(flightRow && flightRow[LOG_FLIGHT_COL.DATE]);
    if (!flightDateYmd) return [];

    var mainPilotName = String(flightRow && flightRow[LOG_FLIGHT_COL.PILOT] || '').trim();
    var loadJsonRaw = String(flightRow && flightRow[LOG_FLIGHT_COL.ACTUAL_LOAD_JSON] || '').trim();
    var onboardAlumnoNames = _dutyCollectOnboardAlumnoNamesFromLoadJson_(loadJsonRaw, mainPilotName);
    if (!onboardAlumnoNames.length) return [];

    var persistedRows = [];
    var sharedDepartment = _dutyDepartmentForPilotDate_(ss, mainPilotName, flightDateYmd);
    if (!sharedDepartment) sharedDepartment = 'Operações';

    var airborne = String(flightRow && flightRow[LOG_FLIGHT_COL.AIRBORNE] || '').trim();
    var landed = String(flightRow && flightRow[LOG_FLIGHT_COL.LANDED] || '').trim();
    var fallbackDescription = 'Voo automático registrado a partir do Flight_ID ' + String(flightLegId || '').trim();

    onboardAlumnoNames.forEach(function(studentName) {
      if (_dutyHasEntryForPilotDate_(ss, studentName, flightDateYmd)) return;

      var dailyFlight = _dutyDailyFlightSummary_(ss, studentName, flightDateYmd) || {};
      var flightHours = Number(dailyFlight.totalFlightHours || 0);
      if (!(flightHours > 0)) return;

      var summaryLine = String(dailyFlight.summaryLine || '').trim() || fallbackDescription;
      var result = savePilotDutyReport({
        pilotName: studentName,
        dateYmd: flightDateYmd,
        status: 'DUTY_CLOSED',
        startTime: airborne,
        endTime: landed,
        department: sharedDepartment,
        flightHours: Number(flightHours.toFixed(2)),
        description: fallbackDescription + (mainPilotName ? ' | Piloto: ' + mainPilotName : ''),
        summaryText: summaryLine
      });

      persistedRows.push(result);
    });

    return persistedRows;
  }

function _postFlightNormName_(value) {
  return String(value || '').trim().toUpperCase().replace(/\s+/g, ' ');
}

function _postFlightResolveStudentEmail_(ss, traineeNameCandidates) {
  var pilotsSh = ss.getSheetByName(APP_SHEETS.PILOTS);
  if (!pilotsSh) return '';

  var wantedMap = {};
  (Array.isArray(traineeNameCandidates) ? traineeNameCandidates : []).forEach(function(name) {
    var key = _postFlightNormName_(name);
    if (key) wantedMap[key] = true;
  });
  if (!Object.keys(wantedMap).length) return '';

  var headers = _toolsEnsurePilotStudentHeaders_(pilotsSh);
  var data = pilotsSh.getDataRange().getValues();
  if (!data || data.length < 2) return '';

  var firstEmailMatch = '';
  for (var i = 1; i < data.length; i++) {
    var rec = _toolsStudentRecordFromRow_(headers, data[i], i + 1);
    var recNameKey = _postFlightNormName_(rec && rec.name);
    var recEmail = String(rec && rec.email || '').trim().toLowerCase();
    if (!recNameKey || !recEmail || !wantedMap[recNameKey]) continue;

    var recType = String(rec && rec.personType || '').trim().toUpperCase();
    if (recType === 'STUDENT') return recEmail;
    if (!firstEmailMatch) firstEmailMatch = recEmail;
  }

  return firstEmailMatch;
}

function _postFlightResolveEmailByNames_(ss, nameCandidates, preferredType) {
  var pilotsSh = ss.getSheetByName(APP_SHEETS.PILOTS);
  if (!pilotsSh) return '';

  var wantedMap = {};
  (Array.isArray(nameCandidates) ? nameCandidates : []).forEach(function(name) {
    var key = _postFlightNormName_(name);
    if (key) wantedMap[key] = true;
  });
  if (!Object.keys(wantedMap).length) return '';

  var pref = String(preferredType || '').trim().toUpperCase();
  var headers = _toolsEnsurePilotStudentHeaders_(pilotsSh);
  var data = pilotsSh.getDataRange().getValues();
  if (!data || data.length < 2) return '';

  var firstEmailMatch = '';
  for (var i = 1; i < data.length; i++) {
    var rec = _toolsStudentRecordFromRow_(headers, data[i], i + 1);
    var recNameKey = _postFlightNormName_(rec && rec.name);
    var recEmail = String(rec && rec.email || '').trim().toLowerCase();
    if (!recNameKey || !recEmail || !wantedMap[recNameKey]) continue;

    var recType = String(rec && rec.personType || '').trim().toUpperCase();
    if (pref && recType === pref) return recEmail;
    if (!firstEmailMatch) firstEmailMatch = recEmail;
  }

  return firstEmailMatch;
}

function _postFlightSplitEmails_(raw) {
  return String(raw || '')
    .split(/[\n,;]+/)
    .map(function(v) { return String(v || '').trim().toLowerCase(); })
    .filter(function(v) { return /.+@.+\..+/.test(v); });
}

function _postFlightResolveOnboardPilotEmails_(ss, report, payload, trainingDebrief) {
  var summary = (report && report.summary && typeof report.summary === 'object') ? report.summary : {};
  var lesson = (report && report.lessonPlanSummary && typeof report.lessonPlanSummary === 'object') ? report.lessonPlanSummary : {};
  var funds = (report && report.studentFunds && typeof report.studentFunds === 'object') ? report.studentFunds : {};
  var names = [];

  var addName = function(value) {
    var name = String(value || '').trim();
    if (name) names.push(name);
  };

  addName(summary.pilot);
  addName(summary.copilot);
  addName(payload && payload.pilotName);
  addName(payload && payload.copilotName);
  addName(payload && payload.copilot);
  addName(lesson.instructorName);
  addName(lesson.traineeName);
  var seen = {};
  var recipients = [];
  var resolveEmailForName = function(name) {
    var email = _postFlightResolveEmailByNames_(ss, [name], '');
    return String(email || '').trim().toLowerCase();
  };

  names.forEach(function(name) {
    var email = resolveEmailForName(name);
    if (!email || seen[email]) return;
    seen[email] = true;
    recipients.push(email);
  });

  return recipients;
}

function _postFlightBuildStudentEmailBody_(report, externalLink) {
  function formatPtDate_(value) {
    var text = String(value || '').trim();
    if (!text) return '---';
    var months = { '01': 'jan', '02': 'fev', '03': 'mar', '04': 'abr', '05': 'mai', '06': 'jun', '07': 'jul', '08': 'ago', '09': 'set', '10': 'out', '11': 'nov', '12': 'dez' };
    var ymd = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (ymd) return ymd[3] + '-' + (months[ymd[2]] || ymd[2]) + '-' + ymd[1];
    var d = new Date(text);
    if (!isNaN(d.getTime())) {
      var dd = String(d.getDate()).padStart(2, '0');
      var mm = String(d.getMonth() + 1).padStart(2, '0');
      var yyyy = String(d.getFullYear());
      return dd + '-' + (months[mm] || mm) + '-' + yyyy;
    }
    return text;
  }

  function formatDurationHhMm_(hoursValue) {
    var hours = parseFloat(hoursValue);
    if (!isFinite(hours) || hours < 0) return '---';
    var totalMinutes = Math.round(hours * 60);
    var hh = String(Math.floor(totalMinutes / 60)).padStart(2, '0');
    var mm = String(totalMinutes % 60).padStart(2, '0');
    return hh + ':' + mm;
  }

  var summary = (report && report.summary && typeof report.summary === 'object') ? report.summary : {};
  var log = (report && report.log && typeof report.log === 'object') ? report.log : {};
  var wb = (report && report.wb && typeof report.wb === 'object') ? report.wb : {};
  var lesson = (report && report.lessonPlanSummary && typeof report.lessonPlanSummary === 'object') ? report.lessonPlanSummary : {};
  var funds = (report && report.studentFunds && typeof report.studentFunds === 'object') ? report.studentFunds : {};

  var lines = [
    'Ola,',
    '',
    'Seu relatorio pos-voo ja esta disponivel.',
    '',
    'Flight ID: ' + String(report && report.flightId || '---'),
    'Data: ' + formatPtDate_(summary.date || log.date || '---'),
    'Aeronave: ' + String(summary.aircraft || log.aircraft || '---'),
    'Rota: ' + String(report.routeDisplay || summary.route || log.route || [summary.from || log.from || '---', summary.to || log.to || '---'].join(' -> ')),
    'Tempo total: ' + String(report.totalTimeDisplay || formatDurationHhMm_(log.totalTimeHours) || '---'),
    'Pousos: ' + String(log.numLandings == null ? '---' : log.numLandings),
    'Touch and Go: ' + String(log.numTouchAndGos == null ? '---' : log.numTouchAndGos),
    'W&B Peso/CG: ' + String(wb.grossWeight == null ? '---' : wb.grossWeight) + ' kg / ' + String(wb.cgPosition == null ? '---' : wb.cgPosition),
    'Plano de aula: ' + String(lesson.lessonName || report.lessonName || '---'),
    'Resultado: ' + String(lesson.flightResult || '---'),
    'Saldo do fundo: ' + (funds.balanceBrl == null ? '---' : ('R$ ' + Number(funds.balanceBrl || 0).toFixed(2))),
    ''
  ];

  var tab8 = (report && report.debrief && report.debrief.tab8 && typeof report.debrief.tab8 === 'object') ? report.debrief.tab8 : {};
  lines.push('TAB 8 - DEBRIEF');
  lines.push('Tempo total Tab8: ' + String(report.totalTimeDisplay || formatDurationHhMm_(tab8.totalTimeHours) || tab8.totalTime || '---'));
  lines.push('Tempo de instrucao: ' + String(report.instructionTimeDisplay || tab8.instructionTimeDisplay || '---'));
  lines.push('Dual recebido: ' + String(report.dualRecvdDisplay || tab8.dualRecvdDisplay || '---'));
  lines.push('Dual dado: ' + String(report.dualGivenDisplay || tab8.dualGivenDisplay || '---'));
  lines.push('Pousos noturnos: ' + String(tab8.nightLdgs == null ? '---' : tab8.nightLdgs));
  lines.push('Horas noturnas: ' + String(tab8.nightHrs == null ? '---' : tab8.nightHrs));
  lines.push('Horas IFR: ' + String(tab8.ifrHrs == null ? '---' : tab8.ifrHrs));
  lines.push('Debrief em: ' + String(tab8.debriefAt || '---'));
  lines.push('Landing log Tab8: ' + JSON.stringify(Array.isArray(tab8.landingLog) ? tab8.landingLog : []));
  lines.push('Training debrief Tab8: ' + String(tab8.trainingDebriefText || (tab8.trainingDebrief ? JSON.stringify(tab8.trainingDebrief) : '---')));
  lines.push('');

  if (externalLink) {
    lines.push('Abrir relatorio completo:');
    lines.push(String(externalLink));
    lines.push('');
  }

  lines.push('Atenciosamente,');
  lines.push('Equipe Flight Ops');

  return lines.join('\n');
}

function _sendPostFlightReportToStudent_(ss, flightLegId, payload, trainingDebrief) {
  try {
    var appUrl = String(ScriptApp.getService().getUrl() || '').replace(/\/$/, '');
    var externalLink = appUrl ? (appUrl + '?view=flightreport&flightId=' + encodeURIComponent(String(flightLegId || ''))) : '';

    var report = getToolsFlightDetailReport({ flightId: flightLegId });
    if (!report || !report.success) {
      return {
        sent: false,
        skipped: true,
        reason: 'post-flight report unavailable'
      };
    }

    var cfg = {};
    try {
      cfg = _toolsDutyReadConfigMap_();
    } catch (cfgErr) {
      cfg = _toolsDutyConfigDefaults_();
    }

    var extraRecipients = _postFlightSplitEmails_(cfg.POSTFLIGHT_EMAIL_EXTRA_RECIPIENTS || '');

    var summary = (report.summary && typeof report.summary === 'object') ? report.summary : {};
    var lesson = (report.lessonPlanSummary && typeof report.lessonPlanSummary === 'object') ? report.lessonPlanSummary : {};
    var funds = (report.studentFunds && typeof report.studentFunds === 'object') ? report.studentFunds : {};

    var recipientsMap = {};
    var addRecipient = function(email) {
      var e = String(email || '').trim().toLowerCase();
      if (!/.+@.+\..+/.test(e)) return;
      recipientsMap[e] = true;
    };

    _postFlightResolveOnboardPilotEmails_(ss, report, payload, trainingDebrief).forEach(addRecipient);

    extraRecipients.forEach(addRecipient);
    var recipients = Object.keys(recipientsMap);
    if (!recipients.length) {
      return {
        sent: false,
        skipped: true,
        reason: 'no configured recipients resolved'
      };
    }

    var subject = '[Pos-Voo] Relatorio do voo ' + String(flightLegId || '');
    var body = _postFlightBuildStudentEmailBody_(report, externalLink);
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      body: body
    });

    return {
      sent: true,
      to: recipients,
      subject: subject,
      link: externalLink
    };
  } catch (e) {
    return {
      sent: false,
      skipped: false,
      error: e && e.message ? e.message : String(e)
    };
  }
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

function _perfAircraftTypeKey_(value) {
  return String(value || '')
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '');
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
    mtowByModel: find('MTOW_BY_MODEL', 'MTOW_MODEL_LIMITS', 'MTOW_LIMITS_BY_MODEL'),
    heading: find('RUNWAY_HEADING', 'HEADING'),
    length: find('LENGTH_OFFICIAL', 'LENGTH_METERS', 'LENGTH_M'),
    width: find('WIDTH_OFFICIAL', 'WIDTH_METERS', 'WIDTH_M'),
    surface: find('SURFACE_ACTUAL', 'SURFACE_OFFICIAL', 'SURFACE'),
    elevation: find('ELEVATION', 'ALT_FEET', 'ELEVATION_FT')
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

function _toolsNormalizeAircraftReg_(value) {
  return String(value || '')
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
}

function _toolsFindAircraftRowByReg_(rows, regIdx, aircraftReg) {
  var target = String(aircraftReg || '').trim().toUpperCase();
  var targetNorm = _toolsNormalizeAircraftReg_(aircraftReg);
  for (var i = 1; i < rows.length; i++) {
    var rowReg = String(rows[i][regIdx] || '').trim().toUpperCase();
    if (rowReg === target) return i;
    if (targetNorm && _toolsNormalizeAircraftReg_(rowReg) === targetNorm) return i;
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

  function ensureMissingHeaders_(sheet, desiredHeaders) {
    if (!sheet) return;
    var lastCol = sheet.getLastColumn();
    var existing = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    var present = {};
    existing.forEach(function(h) { present[_toolsNormHeader_(h)] = true; });
    var additions = [];
    desiredHeaders.forEach(function(h) {
      if (!present[_toolsNormHeader_(h)]) additions.push(h);
    });
    if (!additions.length) return;
    sheet.getRange(1, lastCol + 1, 1, additions.length)
      .setValues([additions])
      .setBackground('#1f2937')
      .setFontColor('white')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  ensure(APP_SHEETS.MAINT_TEMPLATES, [
    'TASK_CODE', 'TASK_NAME', 'AIRCRAFT_TYPE', 'CATEGORY', 'REFERENCE',
    'DUE_MODE', 'INTERVAL_HOURS', 'INTERVAL_DAYS', 'INTERVAL_MONTHS', 'INTERVAL_YEARS', 'ACTIVE', 'SOURCE', 'CAMO_KEY', 'NOTES', 'CREATED_AT', 'UPDATED_AT'
  ]);
  ensure(APP_SHEETS.MAINT_ASSIGNMENTS, [
    'ASSIGNMENT_ID', 'AIRCRAFT_REG', 'TASK_CODE', 'TASK_NAME', 'CATEGORY', 'REFERENCE',
    'DUE_MODE', 'INTERVAL_HOURS', 'INTERVAL_DAYS', 'INTERVAL_MONTHS', 'INTERVAL_YEARS', 'START_TACH', 'START_DATE', 'LAST_PERFORMED_TACH', 'LAST_PERFORMED_DATE', 'DUE_DATE', 'ACTIVE', 'SOURCE', 'CAMO_KEY', 'NOTES', 'CREATED_AT', 'UPDATED_AT'
  ]);
  ensure(APP_SHEETS.MAINT_LOG, [
    'LOG_ID', 'ASSIGNMENT_ID', 'AIRCRAFT_REG', 'COMPLETED_DATE', 'COMPLETED_TACH',
    'REFERENCE_DOC', 'PERFORMED_BY', 'REMARKS', 'CREATED_AT'
  ]);

  ensureMissingHeaders_(ss.getSheetByName(APP_SHEETS.MAINT_ASSIGNMENTS), [
    'ASSIGNMENT_ID', 'AIRCRAFT_REG', 'TASK_CODE', 'TASK_NAME', 'CATEGORY', 'REFERENCE',
    'DUE_MODE', 'INTERVAL_HOURS', 'INTERVAL_DAYS', 'INTERVAL_MONTHS', 'INTERVAL_YEARS', 'START_TACH', 'START_DATE', 'LAST_PERFORMED_TACH', 'LAST_PERFORMED_DATE', 'DUE_DATE', 'ACTIVE', 'SOURCE', 'CAMO_KEY', 'NOTES', 'CREATED_AT', 'UPDATED_AT'
  ]);
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

function _mxParseWholeNum_(value) {
  var n = parseFloat(value);
  return isFinite(n) ? n : '';
}

function _mxCalendarDueDate_(baseDate, intervalDays, intervalMonths, intervalYears) {
  var dt = baseDate instanceof Date ? new Date(baseDate.getTime()) : new Date(baseDate);
  if (!(dt instanceof Date) || isNaN(dt.getTime())) return '';
  var years = _mxParseWholeNum_(intervalYears);
  var months = _mxParseWholeNum_(intervalMonths);
  var days = _mxParseWholeNum_(intervalDays);
  if (years !== '') dt.setFullYear(dt.getFullYear() + Math.round(years));
  if (months !== '') dt.setMonth(dt.getMonth() + Math.round(months));
  if (days !== '') dt.setDate(dt.getDate() + Math.round(days));
  return _mxIsoDate_(dt);
}

function _mxNormalizeDueMode_(value, hasHours, hasCalendar) {
  var raw = String(value || '').trim().toUpperCase();
  if (!raw) {
    if (hasHours && hasCalendar) return 'AMBOS';
    if (hasHours) return 'HORAS';
    if (hasCalendar) return 'DATA';
    return '';
  }
  if (raw === 'HORA' || raw === 'HORAS' || raw === 'HOURS' || raw === 'HOUR' || raw === 'TACH') return 'HORAS';
  if (raw === 'TEMPO' || raw === 'DATA' || raw === 'TIME' || raw === 'DATE') return 'DATA';
  if (raw === 'AMBOS' || raw === 'AMBAS' || raw === 'BOTH' || raw === 'WHICHEVER' || raw === 'WHICHEVER_FIRST') return 'AMBOS';
  if (hasHours && hasCalendar) return 'AMBOS';
  if (hasHours) return 'HORAS';
  if (hasCalendar) return 'DATA';
  return raw;
}

function _mxDueModeLabel_(mode) {
  var key = String(mode || '').trim().toUpperCase();
  if (key === 'HORAS') return 'Horas';
  if (key === 'DATA') return 'Tempo';
  if (key === 'AMBOS') return 'Horas ou tempo, o que vier primeiro';
  return key || '-';
}

function _mxSingleDueState_(remaining, threshold) {
  if (remaining === '' || remaining == null || !isFinite(Number(remaining))) return 'UNKNOWN';
  var value = Number(remaining);
  if (value <= 0) return 'OVERDUE';
  if (threshold !== '' && value <= threshold) return 'DUE_SOON';
  return 'OK';
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

function _mxAddMonthsIso_(isoDate, months) {
  var dt = new Date(isoDate);
  if (isNaN(dt.getTime())) return '';
  dt.setMonth(dt.getMonth() + Math.round(Number(months) || 0));
  return _mxIsoDate_(dt);
}

function _mxAddYearsIso_(isoDate, years) {
  var dt = new Date(isoDate);
  if (isNaN(dt.getTime())) return '';
  dt.setFullYear(dt.getFullYear() + Math.round(Number(years) || 0));
  return _mxIsoDate_(dt);
}

function _mxPlannedFlightHoursNextMonths_(aircraftReg, monthsAhead) {
  var reg = String(aircraftReg || '').trim().toUpperCase();
  var horizonMonths = Number(monthsAhead);
  if (!reg || !isFinite(horizonMonths) || horizonMonths <= 0) return 0;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, '_mxPlannedFlightHoursNextMonths_');
  var data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return 0;

  var headers = data[0];
  var idx = _mxNormHeaderMap_(headers);
  var dateIdx = idx.DATE;
  var aircraftIdx = idx.AIRCRAFT;
  var statusIdx = idx.STATUS;
  var flightTimeIdx = idx.FLIGHT_TIME;
  if (!(dateIdx >= 0) || !(aircraftIdx >= 0) || !(flightTimeIdx >= 0)) return 0;

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var horizon = new Date(today.getTime());
  horizon.setMonth(horizon.getMonth() + Math.round(horizonMonths));
  horizon.setHours(23, 59, 59, 999);

  var total = 0;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[aircraftIdx] || '').trim().toUpperCase() !== reg) continue;
    if (statusIdx >= 0) {
      var status = String(row[statusIdx] || '').trim().toUpperCase();
      if (status && status !== 'PENDING' && status !== 'APPROVED' && status !== 'SCHEDULED' && status !== 'DRAFT') continue;
    }
    var rawDate = row[dateIdx];
    var rowDate = rawDate instanceof Date ? new Date(rawDate.getTime()) : new Date(rawDate);
    if (isNaN(rowDate.getTime())) continue;
    if (rowDate < today || rowDate > horizon) continue;
    var hrs = Number(row[flightTimeIdx] || 0);
    if (isFinite(hrs) && hrs > 0) total += hrs;
  }
  return Number(total.toFixed(1));
}

function _mxPlannedFlightHoursBetweenDates_(aircraftReg, startIsoDate, endIsoDate) {
  var reg = String(aircraftReg || '').trim().toUpperCase();
  var start = String(startIsoDate || '').trim();
  var end = String(endIsoDate || '').trim();
  if (!reg || !start || !end) return 0;

  var startDate = new Date(start + 'T00:00:00');
  var endDate = new Date(end + 'T23:59:59');
  if (isNaN(startDate.getTime()) || isNaN(endDate.getTime()) || endDate < startDate) return 0;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getRequiredSheet_(ss, APP_SHEETS.DISPATCH, '_mxPlannedFlightHoursBetweenDates_');
  var data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return 0;

  var headers = data[0];
  var idx = _mxNormHeaderMap_(headers);
  var dateIdx = idx.DATE;
  var aircraftIdx = idx.AIRCRAFT;
  var statusIdx = idx.STATUS;
  var flightTimeIdx = idx.FLIGHT_TIME;
  if (!(dateIdx >= 0) || !(aircraftIdx >= 0) || !(flightTimeIdx >= 0)) return 0;

  var total = 0;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[aircraftIdx] || '').trim().toUpperCase() !== reg) continue;
    if (statusIdx >= 0) {
      var status = String(row[statusIdx] || '').trim().toUpperCase();
      if (status && status !== 'PENDING' && status !== 'APPROVED' && status !== 'SCHEDULED' && status !== 'DRAFT') continue;
    }
    var rawDate = row[dateIdx];
    var rowDate = rawDate instanceof Date ? new Date(rawDate.getTime()) : new Date(rawDate);
    if (isNaN(rowDate.getTime())) continue;
    if (rowDate < startDate || rowDate > endDate) continue;
    var hrs = Number(row[flightTimeIdx] || 0);
    if (isFinite(hrs) && hrs > 0) total += hrs;
  }
  return Number(total.toFixed(1));
}

function _mxCurrentOrForecastHoursToInsp_(aircraftRow, monthsAhead, nextInspectionDate) {
  var current = aircraftRow && isFinite(Number(aircraftRow.hoursToInsp)) ? Number(aircraftRow.hoursToInsp) : null;
  var reg = aircraftRow && aircraftRow.reg ? String(aircraftRow.reg).trim().toUpperCase() : '';
  var planned = 0;
  var plannedRemaining = null;
  var inspectionDate = String(nextInspectionDate || '').trim();
  var todayIso = _mxIsoDate_(new Date());
  if (inspectionDate) {
    planned = _mxPlannedFlightHoursBetweenDates_(reg, todayIso, inspectionDate);
  } else {
    planned = _mxPlannedFlightHoursNextMonths_(reg, monthsAhead);
  }
  if (current !== null) plannedRemaining = Number((current - planned).toFixed(1));
  return {
    currentHoursToInsp: current,
    plannedFlightHours: planned,
    plannedHoursToInsp: plannedRemaining,
    calendarHoursToInsp: plannedRemaining,
    inspectionDate: inspectionDate
  };
}

function _mxLooksLikeCvaMaintenance_(taskCode, taskName, category, reference, notes) {
  var text = [taskCode, taskName, category, reference, notes].map(function(value) {
    return String(value || '').trim().toUpperCase();
  }).join(' ');
  return /\bCVA\b/.test(text) || /ANNUAL\s+DUE/.test(text);
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

function _mxDueState_(hoursRemaining, daysRemaining, dueMode, thresholdHours, thresholdDays) {
  var mode = String(dueMode || '').trim().toUpperCase();
  var hourState = _mxSingleDueState_(hoursRemaining, thresholdHours);
  var dayState = _mxSingleDueState_(daysRemaining, thresholdDays);
  if (mode === 'HORAS') return hourState;
  if (mode === 'DATA') return dayState;
  if (hourState === 'OVERDUE' || dayState === 'OVERDUE') return 'OVERDUE';
  if (hourState === 'DUE_SOON' || dayState === 'DUE_SOON') return 'DUE_SOON';
  if (hourState === 'UNKNOWN' && dayState === 'UNKNOWN') return 'UNKNOWN';
  return 'OK';
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
  if (!email) {
    try {
      email = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
    } catch (e0) {
      email = '';
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var permSheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_PERMISSIONS || 'SCHED_Permissions', contextLabel || 'schedulerPermissionCheck');
  if (!email) {
    return {
      email: 'system',
      rowNumber: 0,
      headers: [],
      row: []
    };
  }

  var rec = _schedulerPermissionRowByEmail_(permSheet, email);
  if (!rec) {
    permSheet.appendRow([
      email,
      'Y',
      'Y',
      'Y',
      'Y',
      'Y',
      'Y',
      'Y',
      'Y',
      'Auto-seeded by _schedulerAssertPermission_',
      new Date(),
      email
    ]);
    rec = _schedulerPermissionRowByEmail_(permSheet, email);
  }
  if (!rec) throw new Error((contextLabel || 'scheduler') + ': no scheduler permissions found for ' + email);

  var activeIdx = _schedulerHeaderIndex_(rec.headers, 'ACTIVE');
  if (activeIdx >= 0 && !_schedulerTruthyFlag_(rec.row[activeIdx])) {
    throw new Error((contextLabel || 'scheduler') + ': scheduler access is inactive for ' + email);
  }

  if (permissionKey) {
    var idx = _schedulerHeaderIndex_(rec.headers, permissionKey);
    if (idx < 0) return {
      email: email,
      rowNumber: rec.rowNumber,
      headers: rec.headers,
      row: rec.row
    };
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

function _mxCalendarIdFromLinkOrId_(value) {
  var raw = String(value || '').trim();
  if (!raw) return '';
  try {
    var url = new URL(raw);
    var cid = url.searchParams.get('cid');
    if (cid) return decodeURIComponent(cid);
    var src = url.searchParams.get('src');
    if (src) return decodeURIComponent(src);
  } catch (e) {}
  return raw;
}

function _mxRemoveTaggedCalendarEvents_(calendar, tag) {
  var removed = 0;
  if (!calendar || !tag) return removed;
  var start = new Date(2000, 0, 1);
  var end = new Date(2100, 0, 1);
  var events = calendar.getEvents(start, end) || [];
  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    var text = String(ev.getTitle() || '') + '\n' + String(ev.getDescription() || '');
    if (text.indexOf(tag) < 0) continue;
    ev.deleteEvent();
    removed++;
  }
  return removed;
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

function _schedulerRoleKey_(value) {
  return String(value || '').trim().toUpperCase();
}

function _schedulerNormalizeLocalTime_(value) {
  if (value == null || value === '') return '';

  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
  }

  if (typeof value === 'number' && isFinite(value)) {
    var n = value;
    var frac = n % 1;
    if (frac < 0) frac += 1;
    var minsNum = Math.round(frac * 24 * 60) % (24 * 60);
    var hhNum = Math.floor(minsNum / 60);
    var mmNum = minsNum % 60;
    return String(hhNum).padStart(2, '0') + ':' + String(mmNum).padStart(2, '0');
  }

  var raw = String(value).trim();
  if (!raw) return '';

  var hhmm = raw.match(/^(\d{1,2}):(\d{2})/);
  if (hhmm) {
    var hh = Math.max(0, Math.min(23, Number(hhmm[1]) || 0));
    var mm = Math.max(0, Math.min(59, Number(hhmm[2]) || 0));
    return String(hh).padStart(2, '0') + ':' + String(mm).padStart(2, '0');
  }

  var parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'HH:mm');
  }

  return raw;
}

function _schedulerLocalTimeToMinutes_(value) {
  var raw = _schedulerNormalizeLocalTime_(value);
  if (!raw) return null;
  var m = raw.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  var hh = Number(m[1]);
  var mm = Number(m[2]);
  if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
  return hh * 60 + mm;
}

function _schedulerIntervalsOverlap_(aStart, aEnd, bStart, bEnd) {
  if (aStart == null || aEnd == null || bStart == null || bEnd == null) return true;
  return aStart < bEnd && bStart < aEnd;
}

function _schedulerDayMaskSet_(maskRaw) {
  var raw = String(maskRaw || '').trim().toUpperCase();
  if (!raw) return null;
  var out = {};
  var cleanedDigits = raw.replace(/[^0-9]/g, '');
  if (cleanedDigits.length) {
    cleanedDigits.split('').forEach(function(ch) {
      var d = Number(ch);
      if (d >= 1 && d <= 7) out[d % 7] = true;
      if (d === 0) out[0] = true;
    });
    if (Object.keys(out).length) return out;
  }

  raw.split(/[\s,;|/]+/).forEach(function(tok) {
    var t = String(tok || '').trim().toUpperCase();
    if (!t) return;
    if (t === 'MON' || t === 'MONDAY' || t === 'SEG' || t === 'SEGUNDA') out[1] = true;
    if (t === 'TUE' || t === 'TUESDAY' || t === 'TER' || t === 'TERCA' || t === 'TERÇA') out[2] = true;
    if (t === 'WED' || t === 'WEDNESDAY' || t === 'QUA' || t === 'QUARTA') out[3] = true;
    if (t === 'THU' || t === 'THURSDAY' || t === 'QUI' || t === 'QUINTA') out[4] = true;
    if (t === 'FRI' || t === 'FRIDAY' || t === 'SEX' || t === 'SEXTA') out[5] = true;
    if (t === 'SAT' || t === 'SATURDAY' || t === 'SAB' || t === 'SABADO' || t === 'SÁBADO') out[6] = true;
    if (t === 'SUN' || t === 'SUNDAY' || t === 'DOM' || t === 'DOMINGO') out[0] = true;
  });

  return Object.keys(out).length ? out : null;
}

function _schedulerDateMatchesMask_(dateObj, dayMaskSet) {
  if (!dayMaskSet) return true;
  return !!dayMaskSet[dateObj.getDay()];
}

function _schedulerEmailKey_(value) {
  return String(value || '').trim().toLowerCase();
}

function _schedulerAssignmentRowFromHeaders_(headers, row, rowNumber) {
  var idx = {
    assignmentId: _schedulerHeaderIndex_(headers, 'ASSIGNMENT_ID'),
    month: _schedulerHeaderIndex_(headers, 'SCHEDULE_MONTH'),
    date: _schedulerHeaderIndex_(headers, 'ASSIGNMENT_DATE'),
    role: _schedulerHeaderIndex_(headers, 'ROLE_CODE'),
    location: _schedulerHeaderIndex_(headers, 'LOCATION_CODE'),
    startLocal: _schedulerHeaderIndex_(headers, 'SHIFT_START_LOCAL'),
    endLocal: _schedulerHeaderIndex_(headers, 'SHIFT_END_LOCAL'),
    staffEmail: _schedulerHeaderIndex_(headers, 'STAFF_EMAIL'),
    source: _schedulerHeaderIndex_(headers, 'SOURCE'),
    locked: _schedulerHeaderIndex_(headers, 'LOCKED'),
    status: _schedulerHeaderIndex_(headers, 'STATUS')
  };
  return {
    rowNumber: rowNumber,
    assignmentId: idx.assignmentId >= 0 ? String(row[idx.assignmentId] || '').trim() : '',
    scheduleMonth: idx.month >= 0 ? String(row[idx.month] || '').trim() : '',
    assignmentDate: idx.date >= 0 ? safeDateStr(row[idx.date]) : '',
    roleCode: idx.role >= 0 ? String(row[idx.role] || '').trim() : '',
    locationCode: idx.location >= 0 ? String(row[idx.location] || '').trim() : '',
    shiftStartLocal: idx.startLocal >= 0 ? String(row[idx.startLocal] || '').trim() : '',
    shiftEndLocal: idx.endLocal >= 0 ? String(row[idx.endLocal] || '').trim() : '',
    staffEmail: idx.staffEmail >= 0 ? _schedulerEmailKey_(row[idx.staffEmail]) : '',
    source: idx.source >= 0 ? String(row[idx.source] || '').trim().toUpperCase() : '',
    locked: idx.locked >= 0 ? _schedulerTruthyFlag_(row[idx.locked]) : false,
    status: idx.status >= 0 ? String(row[idx.status] || '').trim().toUpperCase() : ''
  };
}

function _schedulerCompatibilityMap_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_ROLE_COMPAT || 'SCHED_Role_Compatibility', 'schedulerCompatibilityMap');
  var data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return {};
  var headers = data[0] || [];
  var roleAIdx = _schedulerHeaderIndex_(headers, 'ROLE_CODE_A');
  var roleBIdx = _schedulerHeaderIndex_(headers, 'ROLE_CODE_B');
  var overlapIdx = _schedulerHeaderIndex_(headers, 'OVERLAP_ALLOWED');
  var activeIdx = _schedulerHeaderIndex_(headers, 'ACTIVE');
  var map = {};

  function setPair(a, b, allowed) {
    var key = a + '|' + b;
    map[key] = allowed;
  }

  for (var i = 1; i < data.length; i++) {
    var active = activeIdx < 0 ? true : _schedulerTruthyFlag_(data[i][activeIdx]);
    if (!active) continue;
    var a = _schedulerRoleKey_(roleAIdx >= 0 ? data[i][roleAIdx] : '');
    var b = _schedulerRoleKey_(roleBIdx >= 0 ? data[i][roleBIdx] : '');
    if (!a || !b) continue;
    var allowed = _schedulerTruthyFlag_(overlapIdx >= 0 ? data[i][overlapIdx] : 'N');
    setPair(a, b, allowed);
    if (a !== '*' && b !== '*') setPair(b, a, allowed);
  }
  return map;
}

function _schedulerOverlapAllowed_(compatMap, roleA, roleB) {
  var a = _schedulerRoleKey_(roleA);
  var b = _schedulerRoleKey_(roleB);
  if (!a || !b) return false;
  if (a === b) return true;
  if (Object.prototype.hasOwnProperty.call(compatMap, a + '|' + b)) return !!compatMap[a + '|' + b];
  if (Object.prototype.hasOwnProperty.call(compatMap, a + '|*')) return !!compatMap[a + '|*'];
  if (Object.prototype.hasOwnProperty.call(compatMap, '*|' + b)) return !!compatMap['*|' + b];
  return false;
}

function _schedulerBuildStaffPool_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pilotsSheet = getRequiredSheet_(ss, APP_SHEETS.PILOTS, 'schedulerBuildStaffPool');
  var pData = pilotsSheet.getDataRange().getValues();
  if (!pData || pData.length < 2) return [];
  var headers = pData[0] || [];
  var out = [];
  for (var i = 1; i < pData.length; i++) {
    var rec = _toolsStaffRecordFromRow_(headers, pData[i], i + 1);
    if (!rec || !rec.active) continue;
    if (!rec.email) continue;
    out.push(rec);
  }
  return out;
}

function _schedulerQualificationRows_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_STAFF_QUALS || 'SCHED_Staff_Qualifications', 'schedulerQualificationRows');
  var data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return [];
  var headers = data[0] || [];
  var idx = {
    staffEmail: _schedulerHeaderIndex_(headers, 'STAFF_EMAIL'),
    role: _schedulerHeaderIndex_(headers, 'ROLE_CODE'),
    location: _schedulerHeaderIndex_(headers, 'LOCATION_CODE'),
    validFrom: _schedulerHeaderIndex_(headers, 'VALID_FROM'),
    validUntil: _schedulerHeaderIndex_(headers, 'VALID_UNTIL'),
    active: _schedulerHeaderIndex_(headers, 'ACTIVE')
  };
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var active = idx.active < 0 ? true : _schedulerTruthyFlag_(data[i][idx.active]);
    if (!active) continue;
    var email = idx.staffEmail >= 0 ? _schedulerEmailKey_(data[i][idx.staffEmail]) : '';
    var role = idx.role >= 0 ? _schedulerRoleKey_(data[i][idx.role]) : '';
    if (!email || !role) continue;
    out.push({
      staffEmail: email,
      roleCode: role,
      locationCode: idx.location >= 0 ? _schedulerRoleKey_(data[i][idx.location]) : '',
      validFrom: idx.validFrom >= 0 ? safeDateStr(data[i][idx.validFrom]) : '',
      validUntil: idx.validUntil >= 0 ? safeDateStr(data[i][idx.validUntil]) : ''
    });
  }
  return out;
}

function _schedulerCoverageNeeds_(monthKey) {
  var parsed = _schedulerParseMonthKey_(monthKey);
  if (!parsed) return [];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_COVERAGE_RULES || 'SCHED_Coverage_Requirements', 'schedulerCoverageNeeds');
  var data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return [];
  var headers = data[0] || [];
  var idx = {
    role: _schedulerHeaderIndex_(headers, 'ROLE_CODE'),
    location: _schedulerHeaderIndex_(headers, 'LOCATION_CODE'),
    daysMask: _schedulerHeaderIndex_(headers, 'DAYS_MASK'),
    startLocal: _schedulerHeaderIndex_(headers, 'SHIFT_START_LOCAL'),
    endLocal: _schedulerHeaderIndex_(headers, 'SHIFT_END_LOCAL'),
    shiftLabel: _schedulerHeaderIndex_(headers, 'SHIFT_LABEL'),
    requiredCount: _schedulerHeaderIndex_(headers, 'REQUIRED_COUNT'),
    priority: _schedulerHeaderIndex_(headers, 'PRIORITY'),
    active: _schedulerHeaderIndex_(headers, 'ACTIVE')
  };

  var needs = [];
  for (var r = 1; r < data.length; r++) {
    var active = idx.active < 0 ? true : _schedulerTruthyFlag_(data[r][idx.active]);
    if (!active) continue;
    var roleCode = idx.role >= 0 ? _schedulerRoleKey_(data[r][idx.role]) : '';
    if (!roleCode) continue;
    var requiredCount = Math.max(0, Number(data[r][idx.requiredCount] || 0));
    if (!requiredCount) continue;
    var dayMaskSet = _schedulerDayMaskSet_(idx.daysMask >= 0 ? data[r][idx.daysMask] : '');
    var startLocal = idx.startLocal >= 0 ? _schedulerNormalizeLocalTime_(data[r][idx.startLocal]) : '';
    var endLocal = idx.endLocal >= 0 ? _schedulerNormalizeLocalTime_(data[r][idx.endLocal]) : '';
    var priority = Number(data[r][idx.priority] || 100) || 100;
    var locationCode = idx.location >= 0 ? _schedulerRoleKey_(data[r][idx.location]) : '';
    var shiftLabel = idx.shiftLabel >= 0 ? String(data[r][idx.shiftLabel] || '').trim() : '';

    for (var d = new Date(parsed.start); d < parsed.end; d.setDate(d.getDate() + 1)) {
      if (!_schedulerDateMatchesMask_(d, dayMaskSet)) continue;
      var dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      for (var slot = 1; slot <= requiredCount; slot++) {
        needs.push({
          scheduleMonth: parsed.monthKey,
          assignmentDate: dateStr,
          roleCode: roleCode,
          locationCode: locationCode,
          shiftStartLocal: startLocal,
          shiftEndLocal: endLocal,
          shiftLabel: shiftLabel,
          priority: priority,
          slot: slot
        });
      }
    }
  }

  needs.sort(function(a, b) {
    if (a.assignmentDate !== b.assignmentDate) return a.assignmentDate.localeCompare(b.assignmentDate);
    if (a.priority !== b.priority) return a.priority - b.priority;
    if (a.roleCode !== b.roleCode) return a.roleCode.localeCompare(b.roleCode);
    return a.slot - b.slot;
  });
  return needs;
}

function _schedulerAvailabilityMap_(monthKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getRequiredSheet_(ss, APP_SHEETS.SCHED_AVAILABILITY || 'SCHED_Staff_Availability', 'schedulerAvailabilityMap');
  var data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return {};
  var headers = data[0] || [];
  var idx = {
    month: _schedulerHeaderIndex_(headers, 'SCHEDULE_MONTH'),
    email: _schedulerHeaderIndex_(headers, 'STAFF_EMAIL'),
    dateLocal: _schedulerHeaderIndex_(headers, 'DATE_LOCAL'),
    startLocal: _schedulerHeaderIndex_(headers, 'START_LOCAL'),
    endLocal: _schedulerHeaderIndex_(headers, 'END_LOCAL'),
    type: _schedulerHeaderIndex_(headers, 'TYPE')
  };

  var out = {};
  for (var i = 1; i < data.length; i++) {
    var month = idx.month >= 0 ? String(data[i][idx.month] || '').trim() : '';
    if (monthKey && month !== monthKey) continue;
    var email = idx.email >= 0 ? _schedulerEmailKey_(data[i][idx.email]) : '';
    var dateLocal = idx.dateLocal >= 0 ? String(data[i][idx.dateLocal] || '').trim() : '';
    if (!email || !dateLocal) continue;
    var type = idx.type >= 0 ? String(data[i][idx.type] || '').trim().toUpperCase() : 'UNAVAILABLE';
    if (type && type.indexOf('AVAILABLE') === 0) continue;
    var startM = _schedulerLocalTimeToMinutes_(idx.startLocal >= 0 ? data[i][idx.startLocal] : '00:00');
    var endM = _schedulerLocalTimeToMinutes_(idx.endLocal >= 0 ? data[i][idx.endLocal] : '23:59');
    if (startM == null) startM = 0;
    if (endM == null) endM = 24 * 60;
    if (endM <= startM) endM = 24 * 60;

    var key = email + '|' + dateLocal;
    if (!out[key]) out[key] = [];
    out[key].push({ startM: startM, endM: endM });
  }
  return out;
}

function _schedulerIsUnavailable_(availabilityMap, email, dateLocal, startLocal, endLocal) {
  var key = _schedulerEmailKey_(email) + '|' + String(dateLocal || '').trim();
  var rows = availabilityMap[key] || [];
  if (!rows.length) return false;
  var needStart = _schedulerLocalTimeToMinutes_(startLocal);
  var needEnd = _schedulerLocalTimeToMinutes_(endLocal);
  if (needStart == null || needEnd == null || needEnd <= needStart) {
    needStart = 0;
    needEnd = 24 * 60;
  }
  for (var i = 0; i < rows.length; i++) {
    if (_schedulerIntervalsOverlap_(needStart, needEnd, rows[i].startM, rows[i].endM)) return true;
  }
  return false;
}

function _schedulerIsQualified_(staff, need, quals, needDate) {
  var email = _schedulerEmailKey_(staff && staff.email);
  var role = _schedulerRoleKey_(need && need.roleCode);
  var location = _schedulerRoleKey_(need && need.locationCode);
  var dateStr = safeDateStr(needDate);

  var hasAnyQualForRole = false;
  for (var i = 0; i < quals.length; i++) {
    var q = quals[i];
    if (q.roleCode !== role) continue;
    hasAnyQualForRole = true;
    if (q.staffEmail !== email) continue;
    var qLoc = _schedulerRoleKey_(q.locationCode);
    if (qLoc && qLoc !== '*' && location && qLoc !== location) continue;
    if (q.validFrom && q.validFrom > dateStr) continue;
    if (q.validUntil && q.validUntil < dateStr) continue;
    return true;
  }

  if (hasAnyQualForRole) return false;

  var roles = _toolsStaffRoleSet_(staff).map(function(v) { return _schedulerRoleKey_(v); });
  return roles.indexOf(role) >= 0 || _schedulerRoleKey_(staff && staff.primaryRole) === role;
}

function _schedulerConflictsWithAssigned_(assignedByDateEmail, compatMap, email, need) {
  var dateKey = String(need.assignmentDate || '').trim();
  var emailKey = _schedulerEmailKey_(email);
  var key = dateKey + '|' + emailKey;
  var assigned = assignedByDateEmail[key] || [];
  if (!assigned.length) return false;

  var needStart = _schedulerLocalTimeToMinutes_(need.shiftStartLocal);
  var needEnd = _schedulerLocalTimeToMinutes_(need.shiftEndLocal);
  if (needStart == null || needEnd == null || needEnd <= needStart) {
    needStart = 0;
    needEnd = 24 * 60;
  }

  for (var i = 0; i < assigned.length; i++) {
    var cur = assigned[i];
    if (!_schedulerIntervalsOverlap_(needStart, needEnd, cur.startM, cur.endM)) continue;
    if (!_schedulerOverlapAllowed_(compatMap, need.roleCode, cur.roleCode)) return true;
  }
  return false;
}

function _schedulerAddAssignedIndex_(assignedByDateEmail, assignment) {
  var dateKey = String(assignment.assignmentDate || '').trim();
  var emailKey = _schedulerEmailKey_(assignment.staffEmail);
  if (!dateKey || !emailKey) return;
  var key = dateKey + '|' + emailKey;
  if (!assignedByDateEmail[key]) assignedByDateEmail[key] = [];
  var startM = _schedulerLocalTimeToMinutes_(assignment.shiftStartLocal);
  var endM = _schedulerLocalTimeToMinutes_(assignment.shiftEndLocal);
  if (startM == null || endM == null || endM <= startM) {
    startM = 0;
    endM = 24 * 60;
  }
  assignedByDateEmail[key].push({ roleCode: _schedulerRoleKey_(assignment.roleCode), startM: startM, endM: endM });
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

function _toolsCurrentUserEmail_(overrideEmailRaw) {
  var overrideEmail = String(overrideEmailRaw || '').trim().toLowerCase();
  if (overrideEmail && _dutyLooksLikeEmail_(overrideEmail)) {
    return overrideEmail;
  }
  try {
    return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  } catch (e) {
    return '';
  }
}

function _toolsStaffRecordFromRow_(headers, row, rowNumber) {
  var nameIdx = _toolsHeaderIndexFromCandidates_(headers, ['PILOT_NAME', 'PILOT', 'NAME', 'STAFF_NAME', 'NOME']);
  var emailIdx = _toolsHeaderIndexFromCandidates_(headers, ['PILOT_EMAIL', 'EMAIL', 'E_MAIL']);
  var staffIdIdx = _toolsHeaderIndexFromCandidates_(headers, ['STAFF_ID', 'PILOT_ID', 'ID']);
  var activeIdx = _toolsHeaderIndexFromCandidates_(headers, ['ACTIVE', 'ATIVO']);
  var roleIdx = _toolsHeaderIndexFromCandidates_(headers, ['PRIMARY_ROLE', 'ROLE', 'FUNCAO', 'FUNC\u00c3O']);
  var rolesJsonIdx = _toolsHeaderIndexFromCandidates_(headers, ['STAFF_ROLES_JSON']);
  var qualificationsJsonIdx = _toolsHeaderIndexFromCandidates_(headers, ['QUALIFICATIONS_JSON']);
  var notesIdx = _toolsHeaderIndexFromCandidates_(headers, ['NOTES']);
  var canEditDiscIdx = _toolsHeaderIndexFromCandidates_(headers, ['CAN_EDIT_DISCREPANCIES']);
  var canApproveDefIdx = _toolsHeaderIndexFromCandidates_(headers, ['CAN_APPROVE_DEFERMENTS']);
  var canEditMxIdx = _toolsHeaderIndexFromCandidates_(headers, ['CAN_EDIT_MAINTENANCE']);
  var canFollowIdx = _toolsHeaderIndexFromCandidates_(headers, ['CAN_FLIGHT_FOLLOW']);
  var canCoordIdx = _toolsHeaderIndexFromCandidates_(headers, ['CAN_COORDINATE_FLIGHTS']);
  var canStockIdx = _toolsHeaderIndexFromCandidates_(headers, ['CAN_MANAGE_STOCKROOM']);
  var canInstructIdx = _toolsHeaderIndexFromCandidates_(headers, ['CAN_INSTRUCT']);
  var canInspectIdx = _toolsHeaderIndexFromCandidates_(headers, ['CAN_INSPECT']);
  var parsedRoles = _parseJsonLoose_(rolesJsonIdx >= 0 ? row[rolesJsonIdx] : '', []);
  var parsedQualifications = _parseJsonLoose_(qualificationsJsonIdx >= 0 ? row[qualificationsJsonIdx] : '', []);
  return {
    rowNumber: rowNumber,
    staffName: nameIdx >= 0 ? String(row[nameIdx] || '').trim() : '',
    email: emailIdx >= 0 ? String(row[emailIdx] || '').trim().toLowerCase() : '',
    staffId: staffIdIdx >= 0 ? String(row[staffIdIdx] || '').trim() : '',
    active: activeIdx >= 0 ? _toolsTruthyFlag_(row[activeIdx]) : true,
    primaryRole: roleIdx >= 0 ? String(row[roleIdx] || '').trim() : '',
    staffRoles: Array.isArray(parsedRoles) ? parsedRoles.map(function(v) { return String(v || '').trim(); }).filter(Boolean) : [],
    qualifications: Array.isArray(parsedQualifications) ? parsedQualifications.map(function(v) { return String(v || '').trim(); }).filter(Boolean) : [],
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

function _toolsStudentRatingsOptions_() {
  return [
    { code: 'PP-A', label: 'Piloto Privado Avião (PP-A)' },
    { code: 'PC-A', label: 'Piloto Comercial Avião (PC-A)' },
    { code: 'IFR', label: 'Voo por Instrumentos (IFR)' },
    { code: 'INVA', label: 'Instrutor de Voo Avião (INVA)' },
    { code: 'MLTE', label: 'Classe Multimotor Terrestre (MLTE)' },
    { code: 'MNTE', label: 'Classe Monomotor Terrestre (MNTE)' },
    { code: 'MNAF', label: 'Monomotor Anfíbio (MNAF)' },
    { code: 'MLTE-IFR', label: 'Multimotor + IFR' },
    { code: 'JET', label: 'Tipo Jato' },
    { code: 'ANAC-AGR', label: 'Piloto Agrícola' },
    { code: 'ANAC-CFI', label: 'Instrutor Credenciado' }
  ];
}

function _toolsEnsurePilotStudentHeaders_(sheet) {
  var headers = _toolsSheetHeaderRow_(sheet);
  var existingNorms = headers.map(function(h) { return _toolsNormHeader_(h); });
  var required = [
    'CPF',
    'CANAC',
    'DOB',
    'MEDICAL_EXPIRY',
    'MEDICAL_CLASS',
    'PERSON_TYPE',
    'CURRENT_RATINGS',
    'STUDENT_KEY',
    'STUDENT_PROFILE_JSON',
    'FUNDS_BALANCE_BRL',
    'FUNDS_RATE_BRL_HOUR',
    'RATINGS_EXPIRY_JSON',
    'LAST_DEPOSIT_FILE_URL',
    'LAST_DEPOSIT_FILE_ID',
    'LAST_DEPOSIT_AT'
  ];

  var missing = [];
  required.forEach(function(label) {
    if (existingNorms.indexOf(_toolsNormHeader_(label)) === -1) missing.push(label);
  });

  if (missing.length) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
  }

  return headers;
}

function _toolsStudentFundsLedgerHeaders_() {
  return [
    'TX_ID',
    'STUDENT_KEY',
    'STUDENT_ROW_NUMBER',
    'STUDENT_NAME',
    'STUDENT_EMAIL',
    'TX_DATE',
    'TX_TYPE',
    'AMOUNT_BRL',
    'RATE_BRL_HOUR',
    'BALANCE_AFTER_BRL',
    'RECEIPT_FILE_URL',
    'RECEIPT_FILE_ID',
    'NOTES',
    'REFERENCE_TYPE',
    'REFERENCE_ID',
    'CREATED_AT',
    'CREATED_BY',
    'STATUS'
  ];
}

function _toolsStudentFundsLedgerSheet_(ss) {
  var target = APP_SHEETS.STUDENT_FUNDS_LEDGER || 'DB_Student_Funds_Ledger';
  var ensured = _schemaEnsureSheetHeaders_(ss, target, _toolsStudentFundsLedgerHeaders_());
  return ensured.sheet;
}

function _toolsStudentIdentityKey_(body) {
  var email = String(body && body.email || '').trim().toLowerCase();
  var cpf = String(body && body.cpf || '').trim().replace(/\D+/g, '');
  var canac = String(body && body.canac || '').trim().toUpperCase();
  var name = String(body && body.name || '').trim().toUpperCase().replace(/\s+/g, ' ');
  var dob = String(body && body.dob || '').trim();
  if (cpf) return 'CPF:' + cpf;
  if (canac) return 'CANAC:' + canac;
  if (email) return 'EMAIL:' + email;
  return 'NAME:' + name + '|DOB:' + dob;
}

function _toolsStudentTxDateTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value || '').trim();
}

function _toolsStudentLedgerMap_(ss) {
  var map = {};
  var sh = _toolsStudentFundsLedgerSheet_(ss);
  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return map;

  var headers = values[0] || [];
  var keyIdx = _toolsHeaderIndexFromCandidates_(headers, ['STUDENT_KEY']);
  var amountIdx = _toolsHeaderIndexFromCandidates_(headers, ['AMOUNT_BRL']);
  var rateIdx = _toolsHeaderIndexFromCandidates_(headers, ['RATE_BRL_HOUR']);
  var txDateIdx = _toolsHeaderIndexFromCandidates_(headers, ['TX_DATE']);
  var noteIdx = _toolsHeaderIndexFromCandidates_(headers, ['NOTES']);
  var receiptUrlIdx = _toolsHeaderIndexFromCandidates_(headers, ['RECEIPT_FILE_URL']);
  var byIdx = _toolsHeaderIndexFromCandidates_(headers, ['CREATED_BY']);
  var balAfterIdx = _toolsHeaderIndexFromCandidates_(headers, ['BALANCE_AFTER_BRL']);
  var txIdIdx = _toolsHeaderIndexFromCandidates_(headers, ['TX_ID']);
  var txTypeIdx = _toolsHeaderIndexFromCandidates_(headers, ['TX_TYPE']);
  var refTypeIdx = _toolsHeaderIndexFromCandidates_(headers, ['REFERENCE_TYPE']);
  var refIdIdx = _toolsHeaderIndexFromCandidates_(headers, ['REFERENCE_ID']);
  var statusIdx = _toolsHeaderIndexFromCandidates_(headers, ['STATUS']);
  var studentNameLedIdx = _toolsHeaderIndexFromCandidates_(headers, ['STUDENT_NAME']);

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var studentKey = keyIdx >= 0 ? String(row[keyIdx] || '').trim() : '';
    if (!studentKey) continue;
    if (!map[studentKey]) map[studentKey] = [];
    map[studentKey].push({
      txId: txIdIdx >= 0 ? String(row[txIdIdx] || '').trim() : '',
      txType: txTypeIdx >= 0 ? String(row[txTypeIdx] || '').trim().toUpperCase() : '',
      referenceType: refTypeIdx >= 0 ? String(row[refTypeIdx] || '').trim().toUpperCase() : '',
      referenceId: refIdIdx >= 0 ? String(row[refIdIdx] || '').trim() : '',
      status: statusIdx >= 0 ? String(row[statusIdx] || '').trim().toUpperCase() : '',
      studentName: studentNameLedIdx >= 0 ? String(row[studentNameLedIdx] || '').trim() : '',
      timestamp: txDateIdx >= 0 ? _toolsStudentTxDateTime_(row[txDateIdx]) : '',
      amount: amountIdx >= 0 ? Number(row[amountIdx] || 0) : 0,
      rate: rateIdx >= 0 ? Number(row[rateIdx] || 0) : 0,
      note: noteIdx >= 0 ? String(row[noteIdx] || '').trim() : '',
      receiptUrl: receiptUrlIdx >= 0 ? String(row[receiptUrlIdx] || '').trim() : '',
      by: byIdx >= 0 ? String(row[byIdx] || '').trim() : '',
      balanceAfter: balAfterIdx >= 0 ? Number(row[balAfterIdx] || 0) : 0
    });
  }

  Object.keys(map).forEach(function(k) {
    map[k].sort(function(a, b) {
      return String(a.timestamp || '').localeCompare(String(b.timestamp || ''));
    });
  });

  return map;
}

function _toolsAppendStudentLedgerEntry_(ss, entry) {
  var sh = _toolsStudentFundsLedgerSheet_(ss);
  var headers = _toolsSheetHeaderRow_(sh);
  var txId = 'STX_' + Utilities.getUuid().slice(0, 8).toUpperCase();
  var createdAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd HH:mm:ss');
  var txDate = String(entry && entry.txDate || '').trim() || createdAt;
  var dataMap = {
    TX_ID: txId,
    STUDENT_KEY: String(entry && entry.studentKey || '').trim(),
    STUDENT_ROW_NUMBER: Number(entry && entry.studentRowNumber || 0),
    STUDENT_NAME: String(entry && entry.studentName || '').trim(),
    STUDENT_EMAIL: String(entry && entry.studentEmail || '').trim().toLowerCase(),
    TX_DATE: txDate,
    TX_TYPE: String(entry && entry.txType || 'CREDIT').trim().toUpperCase(),
    AMOUNT_BRL: Number(entry && entry.amountBrl || 0),
    RATE_BRL_HOUR: Number(entry && entry.rateBrlHour || 0),
    BALANCE_AFTER_BRL: Number(entry && entry.balanceAfterBrl || 0),
    RECEIPT_FILE_URL: String(entry && entry.receiptFileUrl || '').trim(),
    RECEIPT_FILE_ID: String(entry && entry.receiptFileId || '').trim(),
    NOTES: String(entry && entry.notes || '').trim(),
    REFERENCE_TYPE: String(entry && entry.referenceType || '').trim(),
    REFERENCE_ID: String(entry && entry.referenceId || '').trim(),
    CREATED_AT: createdAt,
    CREATED_BY: String(entry && entry.createdBy || _toolsCurrentUserEmail_() || 'tools')
  };

  var row = headers.map(function(header) {
    var key = _toolsNormHeader_(header);
    return Object.prototype.hasOwnProperty.call(dataMap, key) ? dataMap[key] : '';
  });
  sh.appendRow(row);
  return txId;
}

function _toolsRoundMoney_(value) {
  return Number((Number(value || 0)).toFixed(2));
}

function _toolsStudentNormalizeDateIso_(value) {
  return safeDateStr(value);
}

function _toolsStudentCurrentBalanceFromLedger_(ss, studentKey, fallback) {
  var key = String(studentKey || '').trim();
  if (!key) return Number(fallback || 0);
  var ledger = _toolsStudentLedgerMap_(ss)[key] || [];
  if (!ledger.length) return Number(fallback || 0);
  var last = ledger[ledger.length - 1] || {};
  var bal = Number(last.balanceAfter || 0);
  return isFinite(bal) ? bal : Number(fallback || 0);
}

function _toolsStudentBuildFifoLots_(entries) {
  var rows = Array.isArray(entries) ? entries.slice() : [];
  rows.sort(function(a, b) {
    return String(a && a.timestamp || '').localeCompare(String(b && b.timestamp || ''));
  });

  var lots = [];
  for (var i = 0; i < rows.length; i++) {
    var item = rows[i] || {};
    var amount = Number(item.amount || 0);
    var rate = Number(item.rate || 0);
    if (amount > 0 && rate > 0) {
      lots.push({
        rate: rate,
        remainingAmount: _toolsRoundMoney_(amount),
        sourceTimestamp: String(item.timestamp || ''),
        sourceNote: String(item.note || '')
      });
      continue;
    }
    if (amount >= 0) continue;

    var toConsume = _toolsRoundMoney_(Math.abs(amount));
    while (toConsume > 0.00001 && lots.length) {
      var lot = lots[0];
      var take = Math.min(lot.remainingAmount, toConsume);
      lot.remainingAmount = _toolsRoundMoney_(lot.remainingAmount - take);
      toConsume = _toolsRoundMoney_(toConsume - take);
      if (lot.remainingAmount <= 0.00001) lots.shift();
    }
  }
  return lots;
}

function _toolsStudentComputeFifoDebit_(entries, debitHours) {
  var hoursWanted = Number(debitHours || 0);
  if (!isFinite(hoursWanted) || hoursWanted <= 0) {
    return { success: true, lines: [], totalAmount: 0, availableHours: 0, shortageHours: 0 };
  }

  var lots = _toolsStudentBuildFifoLots_(entries);
  var availableHours = lots.reduce(function(acc, lot) {
    var r = Number(lot.rate || 0);
    if (!r || r <= 0) return acc;
    return acc + (Number(lot.remainingAmount || 0) / r);
  }, 0);

  var remainingHours = Number(hoursWanted.toFixed(4));
  var lines = [];
  var totalAmount = 0;

  for (var i = 0; i < lots.length && remainingHours > 0.00001; i++) {
    var lot = lots[i];
    var rate = Number(lot.rate || 0);
    var remainingAmount = Number(lot.remainingAmount || 0);
    if (!rate || rate <= 0 || remainingAmount <= 0) continue;

    var lotHours = remainingAmount / rate;
    var useHours = Math.min(remainingHours, lotHours);
    var useAmount = _toolsRoundMoney_(useHours * rate);
    if (useAmount > remainingAmount) useAmount = _toolsRoundMoney_(remainingAmount);
    if (useAmount <= 0) continue;

    var exactHours = useAmount / rate;
    lines.push({
      rate: rate,
      hours: Number(exactHours.toFixed(4)),
      amount: useAmount,
      sourceTimestamp: lot.sourceTimestamp,
      sourceNote: lot.sourceNote
    });

    remainingHours = Number((remainingHours - exactHours).toFixed(4));
    totalAmount = _toolsRoundMoney_(totalAmount + useAmount);
  }

  if (remainingHours > 0.00001) {
    return {
      success: false,
      error: 'Insufficient FIFO credit to cover ' + hoursWanted.toFixed(2) + ' h. Available: ' + availableHours.toFixed(2) + ' h.',
      lines: lines,
      totalAmount: totalAmount,
      availableHours: availableHours,
      shortageHours: Number(remainingHours.toFixed(4))
    };
  }

  return {
    success: true,
    lines: lines,
    totalAmount: totalAmount,
    availableHours: availableHours,
    shortageHours: 0
  };
}

function _toolsStudentDebitForMission_(ss, missionId, pilotName, flightHours, actorEmail) {
  try {
    var hours = Number(flightHours || 0);
    if (!hours || hours <= 0) return { success: true, skipped: true, reason: 'No flight hours' };
    if (!pilotName || !missionId) return { success: false, error: 'Missing pilotName or missionId' };

    var sh = ss.getSheetByName(APP_SHEETS.PILOTS);
    if (!sh) return { success: false, error: 'DB_Pilots not found' };

    var headers = _toolsEnsurePilotStudentHeaders_(sh);
    var data = sh.getDataRange().getValues();
    var nameIdx = _toolsHeaderIndexFromCandidates_(headers, ['PILOT_NAME', 'PILOT', 'NAME', 'STUDENT_NAME', 'NOME']);
    var balanceIdx = _toolsHeaderIndexFromCandidates_(headers, ['FUNDS_BALANCE_BRL', 'BALANCE_BRL']);

    var pilotNameNorm = String(pilotName).trim().toUpperCase();
    var pilotRowNumber = -1;
    var pilotRow = null;
    for (var i = 1; i < data.length; i++) {
      var rowName = nameIdx >= 0 ? String(data[i][nameIdx] || '').trim().toUpperCase() : '';
      if (rowName && rowName === pilotNameNorm) {
        pilotRowNumber = i + 1;
        pilotRow = data[i];
        break;
      }
    }
    if (!pilotRow) return { success: false, error: 'Pilot not found in DB_Pilots: ' + pilotName };

    var rec = _toolsStudentRecordFromRow_(headers, pilotRow, pilotRowNumber);
    var personType = String(rec.personType || '').trim().toUpperCase();
    if (personType !== 'STUDENT') {
      return { success: true, skipped: true, reason: 'Person type ' + (personType || 'UNKNOWN') + ' is non-billable for training debit' };
    }
    if (!rec.studentKey) rec.studentKey = _toolsStudentIdentityKey_(rec);
    var studentKey = rec.studentKey;

    // Idempotency: skip if a FLIGHT_DEBIT_FIFO for this missionId already exists
    var existingEntries = _toolsStudentLedgerMap_(ss)[studentKey] || [];
    for (var e = 0; e < existingEntries.length; e++) {
      if (String(existingEntries[e].referenceId || '') === String(missionId) &&
          String(existingEntries[e].referenceType || '').toUpperCase() === 'FLIGHT_DEBIT_FIFO') {
        return { success: true, skipped: true, reason: 'Debit already posted for mission ' + missionId };
      }
    }

    var currentBalance = _toolsStudentCurrentBalanceFromLedger_(ss, studentKey, Number(rec.balanceBrl || 0));
    var fifoDebit = _toolsStudentComputeFifoDebit_(existingEntries, hours);
    if (!fifoDebit.success) return { success: false, error: fifoDebit.error };
    if (!fifoDebit.lines || !fifoDebit.lines.length) return { success: true, skipped: true, reason: 'No debit lines produced' };

    var nowIso = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd HH:mm:ss');
    var actor = String(actorEmail || _toolsCurrentUserEmail_() || 'system');
    var runningBalance = Number(currentBalance);

    for (var fi = 0; fi < fifoDebit.lines.length; fi++) {
      var d = fifoDebit.lines[fi] || {};
      var debitAmount = _toolsRoundMoney_(Math.abs(Number(d.amount || 0)));
      if (debitAmount <= 0) continue;
      runningBalance = _toolsRoundMoney_(runningBalance - debitAmount);
      _toolsAppendStudentLedgerEntry_(ss, {
        studentKey: studentKey,
        studentRowNumber: pilotRowNumber,
        studentName: rec.name,
        studentEmail: rec.email,
        txDate: nowIso,
        txType: 'DEBIT',
        amountBrl: -debitAmount,
        rateBrlHour: Number(d.rate || 0),
        balanceAfterBrl: runningBalance,
        notes: 'Flight debit ' + Number(d.hours || 0).toFixed(2) + 'h @ R$' + Number(d.rate || 0).toFixed(2) + '/h | Mission: ' + missionId,
        referenceType: 'FLIGHT_DEBIT_FIFO',
        referenceId: missionId,
        createdBy: actor
      });
    }

    // Update FUNDS_BALANCE_BRL snapshot in DB_Pilots row
    if (balanceIdx >= 0) {
      sh.getRange(pilotRowNumber, balanceIdx + 1).setValue(runningBalance);
    }

    return {
      success: true,
      studentKey: studentKey,
      pilotName: rec.name,
      flightHours: hours,
      totalDebited: fifoDebit.totalAmount,
      newBalance: runningBalance
    };
  } catch (ex) {
    return { success: false, error: ex && ex.message ? ex.message : String(ex) };
  }
}

function _toolsStudentRecordFromRow_(headers, row, rowNumber) {
  var nameIdx = _toolsHeaderIndexFromCandidates_(headers, ['PILOT_NAME', 'PILOT', 'NAME', 'STUDENT_NAME', 'NOME']);
  var emailIdx = _toolsHeaderIndexFromCandidates_(headers, ['PILOT_EMAIL', 'EMAIL', 'E_MAIL']);
  var cpfIdx = _toolsHeaderIndexFromCandidates_(headers, ['CPF', 'ID_NUMBER_CPF_PASSPORT']);
  var canacIdx = _toolsHeaderIndexFromCandidates_(headers, ['CANAC', 'ANAC_CODE', 'LICENSE']);
  var dobIdx = _toolsHeaderIndexFromCandidates_(headers, ['DOB', 'DATE_OF_BIRTH', 'BIRTH_DATE']);
  var medExpiryIdx = _toolsHeaderIndexFromCandidates_(headers, ['MEDICAL_EXPIRY', 'MEDICAL_EXP', 'CMA_EXPIRY']);
  var medClassIdx = _toolsHeaderIndexFromCandidates_(headers, ['MEDICAL_CLASS', 'CMA_CLASS']);
  var personTypeIdx = _toolsHeaderIndexFromCandidates_(headers, ['PERSON_TYPE']);
  var staffIdIdx = _toolsHeaderIndexFromCandidates_(headers, ['STAFF_ID']);
  var ratingsIdx = _toolsHeaderIndexFromCandidates_(headers, ['CURRENT_RATINGS', 'RATINGS']);
  var studentKeyIdx = _toolsHeaderIndexFromCandidates_(headers, ['STUDENT_KEY']);
  var balanceIdx = _toolsHeaderIndexFromCandidates_(headers, ['FUNDS_BALANCE_BRL', 'BALANCE_BRL', 'CREDIT_BALANCE']);
  var rateIdx = _toolsHeaderIndexFromCandidates_(headers, ['FUNDS_RATE_BRL_HOUR', 'RATE_BRL_HOUR', 'HOURLY_RATE']);
  var profileIdx = _toolsHeaderIndexFromCandidates_(headers, ['STUDENT_PROFILE_JSON']);
  var ratingsExpiryIdx = _toolsHeaderIndexFromCandidates_(headers, ['RATINGS_EXPIRY_JSON']);
  var mnteValidityIdx = _toolsHeaderIndexFromCandidates_(headers, ['MNTE_VALIDITY']);
  var mnafValidityIdx = _toolsHeaderIndexFromCandidates_(headers, ['MNAF_VALIDITY', 'MANF_VALIDITY']);
  var lastUrlIdx = _toolsHeaderIndexFromCandidates_(headers, ['LAST_DEPOSIT_FILE_URL', 'DEPOSIT_FILE_URL']);
  var lastIdIdx = _toolsHeaderIndexFromCandidates_(headers, ['LAST_DEPOSIT_FILE_ID', 'DEPOSIT_FILE_ID']);

  var ratingsRaw = ratingsIdx >= 0 ? row[ratingsIdx] : '';
  var ratings = [];
  try {
    var parsedRatings = typeof ratingsRaw === 'string' ? JSON.parse(ratingsRaw) : ratingsRaw;
    if (Array.isArray(parsedRatings)) ratings = parsedRatings.map(function(v) { return String(v || '').trim(); }).filter(Boolean);
  } catch (e) {
    ratings = String(ratingsRaw || '').split(/[;,]+/).map(function(v) { return String(v || '').trim(); }).filter(Boolean);
  }

  var ratingsExpiry = {};
  try {
    var parsedMap = ratingsExpiryIdx >= 0 ? _parseJsonLoose_(row[ratingsExpiryIdx], {}) : {};
    if (parsedMap && typeof parsedMap === 'object') ratingsExpiry = parsedMap;
  } catch (eRe) {}
  if (!ratingsExpiry.MNTE && mnteValidityIdx >= 0) ratingsExpiry.MNTE = safeDateStr(row[mnteValidityIdx]);
  if (!ratingsExpiry.MNAF && mnafValidityIdx >= 0) ratingsExpiry.MNAF = safeDateStr(row[mnafValidityIdx]);

  var profileJson = {};
  try {
    profileJson = profileIdx >= 0 ? _parseJsonLoose_(row[profileIdx], {}) : {};
  } catch (e3) {}

  var rawPersonType = personTypeIdx >= 0 ? String(row[personTypeIdx] || '').trim().toUpperCase() : '';
  var inferredPersonType = rawPersonType;
  if (!inferredPersonType) {
    var staffIdVal = staffIdIdx >= 0 ? String(row[staffIdIdx] || '').trim() : '';
    inferredPersonType = staffIdVal ? 'STAFF' : 'STUDENT';
  }

  return {
    rowNumber: rowNumber,
    name: nameIdx >= 0 ? String(row[nameIdx] || '').trim() : '',
    email: emailIdx >= 0 ? String(row[emailIdx] || '').trim().toLowerCase() : '',
    cpf: cpfIdx >= 0 ? String(row[cpfIdx] || '').trim() : '',
    canac: canacIdx >= 0 ? String(row[canacIdx] || '').trim() : '',
    dob: dobIdx >= 0 ? safeDateStr(row[dobIdx]) : '',
    medicalExpiry: medExpiryIdx >= 0 ? safeDateStr(row[medExpiryIdx]) : '',
    medicalClass: medClassIdx >= 0 ? String(row[medClassIdx] || '').trim() : '',
    personType: inferredPersonType,
    ratings: ratings,
    ratingsExpiry: ratingsExpiry,
    studentKey: studentKeyIdx >= 0 ? String(row[studentKeyIdx] || '').trim() : '',
    balanceBrl: balanceIdx >= 0 ? Number(row[balanceIdx] || 0) : Number(profileJson.balanceBrl || 0),
    rateBrlHour: rateIdx >= 0 ? Number(row[rateIdx] || 0) : Number(profileJson.rateBrlHour || 0),
    ledger: [],
    lastDepositFileUrl: lastUrlIdx >= 0 ? String(row[lastUrlIdx] || '').trim() : '',
    lastDepositFileId: lastIdIdx >= 0 ? String(row[lastIdIdx] || '').trim() : ''
  };
}

function _toolsStudentFindLedgerRowByTxId_(ss, txId) {
  var sh = _toolsStudentFundsLedgerSheet_(ss);
  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return null;
  var headers = values[0] || [];
  var txIdIdx = _toolsHeaderIndexFromCandidates_(headers, ['TX_ID']);
  if (txIdIdx < 0) return null;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][txIdIdx] || '').trim() === txId) {
      return { rowIndex: i + 1, headers: headers, row: values[i] };
    }
  }
  return null;
}

function _toolsStudentMarkLedgerRowStatus_(ss, rowIndex, headers, row, newStatus, auditNote) {
  var sh = _toolsStudentFundsLedgerSheet_(ss);
  var statusIdx = _toolsHeaderIndexFromCandidates_(headers, ['STATUS']);
  var noteIdx = _toolsHeaderIndexFromCandidates_(headers, ['NOTES']);
  if (statusIdx >= 0) sh.getRange(rowIndex, statusIdx + 1).setValue(newStatus);
  if (noteIdx >= 0 && auditNote) {
    var existingNote = String(row[noteIdx] || '').trim();
    sh.getRange(rowIndex, noteIdx + 1).setValue(existingNote ? existingNote + ' | ' + auditNote : auditNote);
  }
}

function _toolsStudentUpdatePilotBalance_(ss, studentRowNumber, newBalance) {
  if (!studentRowNumber || studentRowNumber < 2) return;
  var sh = ss.getSheetByName(APP_SHEETS.PILOTS);
  if (!sh) return;
  var headers = _toolsEnsurePilotStudentHeaders_(sh);
  var balanceIdx = _toolsHeaderIndexFromCandidates_(headers, ['FUNDS_BALANCE_BRL', 'BALANCE_BRL']);
  if (balanceIdx >= 0) sh.getRange(studentRowNumber, balanceIdx + 1).setValue(newBalance);
}

function _ensureToolsSchemaSheet_(sheetName, headers, tabColor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);

  var lastCol = sh.getLastColumn();
  if (lastCol < 1) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    var current = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var currentNorms = current.map(function(header) { return _toolsNormHeader_(header); });
    var missing = [];
    headers.forEach(function(header) {
      if (currentNorms.indexOf(_toolsNormHeader_(header)) === -1) missing.push(header);
    });
    if (missing.length) sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  }

  var finalCol = sh.getLastColumn();
  if (finalCol > 0) {
    sh.getRange(1, 1, 1, finalCol)
      .setFontWeight('bold')
      .setBackground(tabColor || '#1f2937')
      .setFontColor('white');
  }
  sh.setFrozenRows(1);
  return sh;
}

function _pilotAuthorizationHeaders_() {
  return [
    'Authorization_ID',
    'Pilot_Name',
    'Pilot_Email',
    'Staff_ID',
    'Aircraft_Type',
    'Role',
    'Authorization_Type',
    'Status',
    'Date_Authorized',
    'Expiry_Date',
    'Instructor_Signoff',
    'Source',
    'Notes',
    'Created_At',
    'Updated_At'
  ];
}

function _pilotAuthorizationOptionSets_() {
  return {
    roles: ['Student', 'Operational', 'Instructor'],
    authorizationTypes: ['INITIAL', 'ANNUAL_PROFICIENCY', 'UPGRADE', 'REINSTATEMENT', 'LINE_CHECK'],
    statuses: ['ACTIVE', 'PENDING', 'SUSPENDED', 'REVOKED', 'EXPIRED']
  };
}

function _ensurePilotAuthorizationSheetSchema_() {
  return _ensureToolsSchemaSheet_(APP_SHEETS.PILOT_AUTHORIZATIONS || 'DB_Pilot_Authorizations', _pilotAuthorizationHeaders_(), '#2c3e50');
}

function _pilotAuthorizationSheet_() {
  return _ensurePilotAuthorizationSheetSchema_();
}

function _pilotAuthorizationRecordFromRow_(headers, row, rowNumber) {
  var authIdIdx = _toolsHeaderIndexFromCandidates_(headers, ['AUTHORIZATION_ID']);
  var nameIdx = _toolsHeaderIndexFromCandidates_(headers, ['PILOT_NAME', 'PILOT']);
  var emailIdx = _toolsHeaderIndexFromCandidates_(headers, ['PILOT_EMAIL', 'EMAIL']);
  var staffIdIdx = _toolsHeaderIndexFromCandidates_(headers, ['STAFF_ID']);
  var aircraftIdx = _toolsHeaderIndexFromCandidates_(headers, ['AIRCRAFT_TYPE', 'AIRCRAFT']);
  var roleIdx = _toolsHeaderIndexFromCandidates_(headers, ['ROLE']);
  var authTypeIdx = _toolsHeaderIndexFromCandidates_(headers, ['AUTHORIZATION_TYPE', 'AUTH_TYPE']);
  var statusIdx = _toolsHeaderIndexFromCandidates_(headers, ['STATUS']);
  var authorizedIdx = _toolsHeaderIndexFromCandidates_(headers, ['DATE_AUTHORIZED', 'AUTHORIZED_AT']);
  var expiryIdx = _toolsHeaderIndexFromCandidates_(headers, ['EXPIRY_DATE', 'VALID_UNTIL']);
  var signoffIdx = _toolsHeaderIndexFromCandidates_(headers, ['INSTRUCTOR_SIGNOFF', 'SIGNOFF']);
  var sourceIdx = _toolsHeaderIndexFromCandidates_(headers, ['SOURCE']);
  var notesIdx = _toolsHeaderIndexFromCandidates_(headers, ['NOTES']);
  var createdIdx = _toolsHeaderIndexFromCandidates_(headers, ['CREATED_AT']);
  var updatedIdx = _toolsHeaderIndexFromCandidates_(headers, ['UPDATED_AT']);

  var pilotName = nameIdx >= 0 ? String(row[nameIdx] || '').trim() : '';
  var pilotEmail = emailIdx >= 0 ? String(row[emailIdx] || '').trim().toLowerCase() : '';
  var aircraftType = aircraftIdx >= 0 ? String(row[aircraftIdx] || '').trim() : '';
  var role = roleIdx >= 0 ? String(row[roleIdx] || '').trim() : '';
  var authorizationType = authTypeIdx >= 0 ? String(row[authTypeIdx] || '').trim() : '';
  var status = statusIdx >= 0 ? String(row[statusIdx] || '').trim() : '';
  var dateAuthorized = authorizedIdx >= 0 ? safeDateStr(row[authorizedIdx]) : '';
  var expiryDate = expiryIdx >= 0 ? safeDateStr(row[expiryIdx]) : '';
  var statusUpper = status.toUpperCase();
  var isActive = !statusUpper || statusUpper === 'ACTIVE' || statusUpper === 'AUTHORIZED' || statusUpper === 'VALID';
  var isExpired = !!expiryDate && expiryDate < safeDateStr(new Date());

  return {
    rowNumber: rowNumber,
    authorizationId: authIdIdx >= 0 ? String(row[authIdIdx] || '').trim() : '',
    pilotName: pilotName,
    pilotEmail: pilotEmail,
    staffId: staffIdIdx >= 0 ? String(row[staffIdIdx] || '').trim() : '',
    aircraftType: aircraftType,
    role: role,
    authorizationType: authorizationType,
    status: status,
    dateAuthorized: dateAuthorized,
    expiryDate: expiryDate,
    instructorSignoff: signoffIdx >= 0 ? String(row[signoffIdx] || '').trim() : '',
    source: sourceIdx >= 0 ? String(row[sourceIdx] || '').trim() : '',
    notes: notesIdx >= 0 ? String(row[notesIdx] || '').trim() : '',
    createdAt: createdIdx >= 0 ? safeDateStr(row[createdIdx]) : '',
    updatedAt: updatedIdx >= 0 ? safeDateStr(row[updatedIdx]) : '',
    isActive: isActive,
    isExpired: isExpired
  };
}

function _toolsFindStaffByEmailOrId_(staffRows, staffEmail, staffId) {
  var email = String(staffEmail || '').trim().toLowerCase();
  var sid = String(staffId || '').trim();
  for (var i = 0; i < staffRows.length; i++) {
    var row = staffRows[i] || {};
    var rowEmail = String(row.email || '').trim().toLowerCase();
    if (email && rowEmail && (rowEmail === email || _dutyPilotLikelySame_(rowEmail, email))) return row;
    if (sid && String(row.staffId || '').trim() === sid) return row;
  }
  return null;
}

function _toolsStaffRoleSet_(staff) {
  var out = {};
  var primary = String((staff && staff.primaryRole) || '').trim();
  if (primary) out[primary.toUpperCase()] = primary;
  var roles = (staff && Array.isArray(staff.staffRoles)) ? staff.staffRoles : [];
  roles.forEach(function(role) {
    var raw = String(role || '').trim();
    if (raw) out[raw.toUpperCase()] = raw;
  });
  return Object.keys(out).map(function(key) { return out[key]; });
}

function _toolsPilotAuthorizationSummary_(staff, aircraftType, authorizationRole) {
  var authRes = getPilotAircraftAuthorizations({
    pilotEmail: staff && staff.email,
    pilotName: staff && staff.staffName,
    aircraftType: aircraftType,
    role: authorizationRole
  });

  if (!authRes || !authRes.success) {
    return {
      required: !!(aircraftType || authorizationRole),
      ready: false,
      error: authRes && authRes.error ? authRes.error : 'Could not read pilot authorizations',
      aircraftType: aircraftType || '',
      role: authorizationRole || '',
      rows: []
    };
  }

  var today = safeDateStr(new Date());
  var rows = Array.isArray(authRes.rows) ? authRes.rows : [];
  var activeRows = rows.filter(function(rec) { return !!rec.isActive; });
  var validRows = activeRows.filter(function(rec) { return !rec.expiryDate || rec.expiryDate >= today; });

  return {
    required: !!(aircraftType || authorizationRole),
    ready: validRows.length > 0,
    aircraftType: aircraftType || '',
    role: authorizationRole || '',
    totalCount: rows.length,
    activeCount: activeRows.length,
    validCount: validRows.length,
    rows: rows
  };
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

function _histFlightNormHeader_(value) {
  return String(value || '').trim().toUpperCase().replace(/\s+/g, '_').replace(/[^A-Z0-9_]/g, '');
}

function _histFlightDigest_(value) {
  var raw = String(value == null ? '' : value);
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return bytes.map(function(b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function _histFlightBatchId_() {
  var d = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  return 'HIST_' + d + '_' + Utilities.getUuid().slice(0, 8).toUpperCase();
}

function _histFlightSuggestedMap_(headers) {
  var list = (headers || []).map(function(h) { return String(h || '').trim(); });
  var normalized = list.map(_histFlightNormHeader_);
  function find(candidates) {
    for (var i = 0; i < candidates.length; i++) {
      var idx = normalized.indexOf(_histFlightNormHeader_(candidates[i]));
      if (idx >= 0) return list[idx];
    }
    return '';
  }
  return {
    flightId: find(['FLIGHT_ID', 'LEG_ID', 'ID_VOO', 'ID', 'VOO_ID']),
    date: find(['DATE', 'DATA', 'FLIGHT_DATE', 'DIA']),
    pilot: find(['PILOT', 'ALUNO', 'PILOTO', 'NOME_ALUNO', 'STUDENT']),
    acft: find(['ACFT', 'AIRCRAFT', 'AERONAVE', 'REGISTRATION', 'MATRICULA']),
    from: find(['FROM', 'ORIGIN', 'ORIGEM', 'DEP', 'DEPARTURE']),
    to: find(['TO', 'DESTINATION', 'DESTINO', 'ARR', 'ARRIVAL']),
    route: find(['ROUTE', 'ROTA', 'PERCURSO']),
    totalTime: find(['TOTAL_TIME', 'FLIGHT_TIME', 'TEMPO_TOTAL', 'HORAS', 'TIME_H']),
    numberLdgs: find(['NUMBER_LDGS', 'NUM_LDGS', 'LANDINGS', 'POUSOS']),
    touchAndGos: find(['NUMBER_TOUCH_AND_GOS', 'TOUCH_AND_GOS', 'TGL', 'TNG']),
    sourceNote: find(['NOTES', 'OBS', 'OBSERVACOES', 'DESCRIPTION', 'DESCRICAO'])
  };
}

function _histFlightParseHours_(value) {
  if (value == null || value === '') return '';
  if (typeof value === 'number' && isFinite(value)) return value;
  var raw = String(value || '').trim();
  if (!raw) return '';
  var n = parseFloat(raw.replace(',', '.'));
  if (isFinite(n) && /^[0-9]+([\.,][0-9]+)?$/.test(raw)) return n;
  var hm = raw.match(/^(\d{1,3})\s*:\s*(\d{1,2})$/);
  if (hm) {
    var hh = Number(hm[1]);
    var mm = Number(hm[2]);
    if (isFinite(hh) && isFinite(mm)) return hh + (mm / 60);
  }
  return '';
}

function _histFlightParseInt_(value) {
  if (value == null || value === '') return '';
  var n = parseInt(String(value).replace(/[^0-9-]/g, ''), 10);
  return isFinite(n) ? n : '';
}

function _histFlightResolveFromTo_(fromRaw, toRaw, routeRaw) {
  var from = String(fromRaw || '').trim().toUpperCase();
  var to = String(toRaw || '').trim().toUpperCase();
  if (from && to) return { from: from, to: to };
  var split = splitRoute_(routeRaw || '');
  if (!from) from = String(split.from || '').trim().toUpperCase();
  if (!to) to = String(split.to || '').trim().toUpperCase();
  return { from: from, to: to };
}

function _histFlightMakeStrictKey_(obj) {
  var hours = obj.totalTimeH === '' ? '' : Number(obj.totalTimeH || 0).toFixed(2);
  return _histFlightDigest_([
    String(obj.date || ''),
    String(obj.pilot || '').toUpperCase(),
    String(obj.acft || '').toUpperCase(),
    String(obj.from || '').toUpperCase(),
    String(obj.to || '').toUpperCase(),
    hours
  ].join('|'));
}

function _histFlightMakeFuzzyKey_(obj) {
  return _histFlightDigest_([
    String(obj.date || ''),
    String(obj.pilot || '').toUpperCase(),
    String(obj.acft || '').toUpperCase()
  ].join('|'));
}

function _histFlightParseTimeToken_(value) {
  var raw = String(value || '').trim();
  if (!raw) return '';
  var m = raw.match(/^(\d{1,2})\s*[:hH]\s*(\d{2})$/);
  if (!m) return '';
  var hh = Number(m[1]);
  var mm = Number(m[2]);
  if (!isFinite(hh) || !isFinite(mm)) return '';
  if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return '';
  return String(hh).padStart(2, '0') + ':' + String(mm).padStart(2, '0');
}

function _histFlightFormatNumber_(value) {
  if (value == null || value === '') return '';
  var n = Number(String(value).replace(',', '.'));
  if (!isFinite(n)) return '';
  return String(n);
}

function _histFlightBuildStructuredTags_(src) {
  var out = [];
  function add(key, value, formatter) {
    var v = formatter ? formatter(value) : String(value || '').trim();
    if (!v) return;
    out.push(String(key || '').trim().toUpperCase() + '=' + String(v).trim());
  }
  add('INSTR_INIT', src && src.instructorInitials);
  add('MISSION_CODE', src && src.missionTrainingCode);
  add('AIRCRAFT_TYPE_RAW', src && src.aircraftTypeRaw);
  add('FUEL_GAL_START', src && src.fuelGallonsStart, _histFlightFormatNumber_);
  add('FUEL_GAL_END', src && src.fuelGallonsEnd, _histFlightFormatNumber_);
  add('TACH_START', src && src.tachStart, _histFlightFormatNumber_);
  add('TACH_END', src && src.tachEnd, _histFlightFormatNumber_);
  add('BRAKES_RELEASE', src && src.brakesRelease, _histFlightParseTimeToken_);
  add('BRAKES_APPLIED', src && src.brakesApplied, _histFlightParseTimeToken_);
  add('AIRBORNE', src && src.airborne, _histFlightParseTimeToken_);
  add('LANDED', src && src.landed, _histFlightParseTimeToken_);
  return out;
}

function _histFlightMergeSourceNote_(baseNote, src) {
  var note = String(baseNote || '').trim();
  var tags = _histFlightBuildStructuredTags_(src);
  if (!tags.length) return note;
  var block = '[HIST:' + tags.join(';') + ']';
  if (!note) return block;
  if (note.indexOf('[HIST:') >= 0) return note;
  return note + ' ' + block;
}

function _histFlightReadStructuredTags_(sourceNote) {
  var out = {};
  var raw = String(sourceNote || '');
  var m = raw.match(/\[HIST:([^\]]+)\]/i);
  if (!m || !m[1]) return out;
  var pairs = String(m[1] || '').split(';');
  for (var i = 0; i < pairs.length; i++) {
    var part = String(pairs[i] || '').trim();
    if (!part) continue;
    var eq = part.indexOf('=');
    if (eq < 0) continue;
    var key = String(part.slice(0, eq) || '').trim().toUpperCase();
    var val = String(part.slice(eq + 1) || '').trim();
    if (!key) continue;
    out[key] = val;
  }
  return out;
}

function _histFlightInstructorNameFromInitials_(initials) {
  var key = String(initials || '').trim().toUpperCase();
  var map = {
    AD: 'Augusto Diego',
    LS: 'Leandro Siqueira',
    VS: 'Victor Scaparro'
  };
  return map[key] || '';
}

function _histFlightAircraftTypeLookup_() {
  var out = { reg: {}, alias: {} };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(APP_SHEETS.AIRCRAFT);
  if (!sh) return out;
  var data = sh.getDataRange().getValues();
  if (!data || data.length < 2) return out;

  var headers = data[0].map(_histFlightNormHeader_);
  var regIdx = _histFlightHeaderIndex_(headers, ['REGISTRATION', 'REG', 'TAIL', 'TAIL_NUMBER', 'AIRCRAFT_REGISTRATION', 'MATRICULA']);
  var typeIdx = _histFlightHeaderIndex_(headers, ['TYPE_FOR_PERFORMANCE', 'AIRCRAFT_TYPE', 'TYPE', 'MODEL']);
  var altTypeIdx = _histFlightHeaderIndex_(headers, ['AIRCRAFT_TYPE', 'TYPE', 'MODEL']);

  function norm(v) { return String(v || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, ''); }
  function setAlias(raw, typeVal) {
    var k = norm(raw);
    var t = String(typeVal || '').trim();
    if (!k || !t) return;
    if (!out.alias[k]) out.alias[k] = t;
  }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var reg = regIdx >= 0 ? String(row[regIdx] || '').trim().toUpperCase() : '';
    var typePrimary = typeIdx >= 0 ? String(row[typeIdx] || '').trim() : '';
    var typeAlt = altTypeIdx >= 0 ? String(row[altTypeIdx] || '').trim() : '';
    var finalType = typePrimary || typeAlt;
    if (!finalType) continue;

    if (reg) {
      out.reg[norm(reg)] = finalType;
      setAlias(reg, finalType);
    }
    setAlias(typePrimary, finalType);
    setAlias(typeAlt, finalType);
  }
  return out;
}

function _histFlightResolveAircraftType_(lookup, acftValue, rawType) {
  function norm(v) { return String(v || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, ''); }
  var lk = lookup || { reg: {}, alias: {} };
  var byReg = lk.reg[norm(acftValue)] || '';
  if (byReg) return byReg;
  var byRaw = lk.alias[norm(rawType)] || '';
  if (byRaw) return byRaw;
  var byAcft = lk.alias[norm(acftValue)] || '';
  if (byAcft) return byAcft;
  return String(rawType || acftValue || '').trim();
}

function _histFlightReadExistingLogKeys_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getRequiredSheet_(ss, APP_SHEETS.LOG_FLIGHTS, '_histFlightReadExistingLogKeys_');
  var data = sh.getDataRange().getValues();
  var out = { ids: {}, strict: {}, fuzzy: {}, rowCount: Math.max(data.length - 1, 0) };
  if (!data || data.length < 2) return out;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var flightId = String(row[LOG_FLIGHT_COL.FLIGHT_ID] || '').trim();
    if (flightId) out.ids[flightId] = true;
    var obj = {
      date: safeDateYmd_(row[LOG_FLIGHT_COL.DATE]),
      pilot: String(row[LOG_FLIGHT_COL.PILOT] || '').trim(),
      acft: String(row[LOG_FLIGHT_COL.ACFT] || '').trim(),
      from: String(row[LOG_FLIGHT_COL.FROM] || '').trim(),
      to: String(row[LOG_FLIGHT_COL.TO] || '').trim(),
      totalTimeH: _histFlightParseHours_(row[LOG_FLIGHT_COL.TOTAL_TIME])
    };
    out.strict[_histFlightMakeStrictKey_(obj)] = flightId || ('ROW_' + (i + 1));
    out.fuzzy[_histFlightMakeFuzzyKey_(obj)] = flightId || ('ROW_' + (i + 1));
  }
  return out;
}

function _histFlightNormalizeText_(value) {
  var raw = String(value || '').trim();
  if (!raw) return '';
  try {
    raw = raw.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  } catch (e) {}
  return raw.toUpperCase().replace(/[^A-Z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function _histFlightNormalizeReg_(value) {
  return String(value || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function _histFlightHeaderIndex_(headers, candidates) {
  var list = Array.isArray(headers) ? headers : [];
  var cand = Array.isArray(candidates) ? candidates : [];
  for (var i = 0; i < cand.length; i++) {
    var idx = list.indexOf(_histFlightNormHeader_(cand[i]));
    if (idx >= 0) return idx;
  }
  return -1;
}

function _histFlightBuildReferenceLookup_() {
  var refs = {
    peopleExact: {},
    peopleEntries: [],
    aircraftExact: {},
    airportExact: {}
  };

  function addPerson(name, role, source) {
    var display = String(name || '').trim();
    if (!display) return;
    var key = _histFlightNormalizeText_(display);
    if (!key) return;
    var entry = { key: key, name: display, role: String(role || ''), source: String(source || '') };
    refs.peopleEntries.push(entry);
    if (!refs.peopleExact[key]) refs.peopleExact[key] = [];
    refs.peopleExact[key].push(entry);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var pilotsSh = ss.getSheetByName(APP_SHEETS.PILOTS);
  if (pilotsSh) {
    var pilotsData = pilotsSh.getDataRange().getValues();
    if (pilotsData && pilotsData.length > 1) {
      var pHeaders = pilotsData[0].map(_histFlightNormHeader_);
      var pNameIdx = _histFlightHeaderIndex_(pHeaders, ['PILOT_NAME', 'NAME', 'STAFF_NAME', 'NOME']);
      for (var p = 1; p < pilotsData.length; p++) {
        if (pNameIdx < 0) break;
        addPerson(pilotsData[p][pNameIdx], 'INSTRUCTOR', 'DB_Pilots');
      }
    }
  }

  var acftSh = ss.getSheetByName(APP_SHEETS.AIRCRAFT);
  if (acftSh) {
    var acftData = acftSh.getDataRange().getValues();
    if (acftData && acftData.length > 1) {
      var aHeaders = acftData[0].map(_histFlightNormHeader_);
      var regIdx = _histFlightHeaderIndex_(aHeaders, ['REGISTRATION', 'REG', 'TAIL', 'TAIL_NUMBER', 'AIRCRAFT_REGISTRATION', 'MATRICULA']);
      for (var a = 1; a < acftData.length; a++) {
        if (regIdx < 0) break;
        var reg = String(acftData[a][regIdx] || '').trim().toUpperCase();
        var regKey = _histFlightNormalizeReg_(reg);
        if (!regKey) continue;
        refs.aircraftExact[regKey] = reg;
      }
    }
  }

  var airSh = ss.getSheetByName(APP_SHEETS.AIRPORTS);
  if (airSh) {
    var airData = airSh.getDataRange().getValues();
    if (airData && airData.length > 1) {
      var iHeaders = airData[0].map(_histFlightNormHeader_);
      var icaoIdx = _histFlightHeaderIndex_(iHeaders, ['ICAO', 'AIRPORT_ICAO', 'CODE']);
      var nameIdx = _histFlightHeaderIndex_(iHeaders, ['NOME', 'NAME', 'AIRPORT_NAME']);
      for (var i = 1; i < airData.length; i++) {
        if (icaoIdx < 0) break;
        var icao = String(airData[i][icaoIdx] || '').trim().toUpperCase();
        if (!icao) continue;
        refs.airportExact[icao] = {
          icao: icao,
          name: nameIdx >= 0 ? String(airData[i][nameIdx] || '').trim() : ''
        };
      }
    }
  }

  return refs;
}

function _histFlightResolvePersonMatch_(candidate, refs) {
  var input = String(candidate || '').trim();
  if (!input) {
    return { input: '', status: 'MISSING', matchedName: '', matchedSource: '', matchedRole: '' };
  }

  var byInitials = _histFlightInstructorNameFromInitials_(input);
  if (byInitials) {
    return {
      input: input,
      status: 'CONFIRMED',
      matchedName: byInitials,
      matchedSource: 'INSTRUCTOR_INITIALS_MAP',
      matchedRole: 'INSTRUCTOR'
    };
  }

  var key = _histFlightNormalizeText_(input);
  var exact = refs && refs.peopleExact ? refs.peopleExact[key] : null;
  if (exact && exact.length === 1) {
    return {
      input: input,
      status: 'CONFIRMED',
      matchedName: String(exact[0].name || ''),
      matchedSource: String(exact[0].source || ''),
      matchedRole: String(exact[0].role || '')
    };
  }
  if (exact && exact.length > 1) {
    return {
      input: input,
      status: 'POSSIBLE_AMBIGUOUS',
      matchedName: String(exact[0].name || ''),
      matchedSource: String(exact[0].source || ''),
      matchedRole: String(exact[0].role || '')
    };
  }

  var entries = refs && Array.isArray(refs.peopleEntries) ? refs.peopleEntries : [];
  var fuzzy = [];
  for (var i = 0; i < entries.length; i++) {
    var ek = String(entries[i].key || '');
    if (!ek) continue;
    if (ek.indexOf(key) >= 0 || key.indexOf(ek) >= 0) fuzzy.push(entries[i]);
    if (fuzzy.length >= 4) break;
  }

  if (fuzzy.length === 1) {
    return {
      input: input,
      status: 'POSSIBLE',
      matchedName: String(fuzzy[0].name || ''),
      matchedSource: String(fuzzy[0].source || ''),
      matchedRole: String(fuzzy[0].role || '')
    };
  }
  if (fuzzy.length > 1) {
    return {
      input: input,
      status: 'POSSIBLE_AMBIGUOUS',
      matchedName: String(fuzzy[0].name || ''),
      matchedSource: String(fuzzy[0].source || ''),
      matchedRole: String(fuzzy[0].role || '')
    };
  }

  return { input: input, status: 'NOT_FOUND', matchedName: '', matchedSource: '', matchedRole: '' };
}

function _histFlightResolveAircraftMatch_(acft, refs) {
  var input = String(acft || '').trim().toUpperCase();
  if (!input) return { input: '', status: 'MISSING', matchedRegistration: '' };
  var key = _histFlightNormalizeReg_(input);
  var reg = refs && refs.aircraftExact ? refs.aircraftExact[key] : '';
  if (reg) return { input: input, status: 'CONFIRMED', matchedRegistration: String(reg || '') };
  return { input: input, status: 'NOT_FOUND', matchedRegistration: '' };
}

function _histFlightResolveAirportMatch_(icao, refs) {
  var input = String(icao || '').trim().toUpperCase();
  if (!input) return { input: '', status: 'MISSING', matchedIcao: '', matchedName: '' };
  var hit = refs && refs.airportExact ? refs.airportExact[input] : null;
  if (hit) {
    return {
      input: input,
      status: 'CONFIRMED',
      matchedIcao: String(hit.icao || input),
      matchedName: String(hit.name || '')
    };
  }
  return { input: input, status: 'NOT_FOUND', matchedIcao: '', matchedName: '' };
}

function _histFlightExtractCrewCandidates_(pilot, sourceNote) {
  var note = String(sourceNote || '');
  function pick(labels) {
    for (var i = 0; i < labels.length; i++) {
      var rx = new RegExp(labels[i] + '\\s*[:=-]\\s*([^|;,\\n\\r]+)', 'i');
      var m = note.match(rx);
      if (m && m[1]) return String(m[1] || '').trim();
    }
    return '';
  }

  var student = pick(['aluno', 'student', 'estudante']);
  var instructor = pick(['instrutor', 'instructor', 'cfi']);
  var pilotName = String(pilot || '').trim();
  if (!student && pilotName) student = pilotName;

  return {
    student: student,
    instructor: instructor,
    pilot: pilotName
  };
}

function _histFlightBuildRowValidation_(row, refs) {
  var crew = _histFlightExtractCrewCandidates_(row && row.pilot, row && row.sourceNote);
  var studentMatch = _histFlightResolvePersonMatch_(crew.student, refs);
  var instructorMatch = _histFlightResolvePersonMatch_(crew.instructor, refs);
  var pilotMatch = _histFlightResolvePersonMatch_(crew.pilot, refs);
  var aircraftMatch = _histFlightResolveAircraftMatch_(row && row.acft, refs);
  var fromAirport = _histFlightResolveAirportMatch_(row && row.from, refs);
  var toAirport = _histFlightResolveAirportMatch_(row && row.to, refs);

  var issues = [];
  function pushIssue(label, matchObj, allowMissing) {
    var st = String(matchObj && matchObj.status || '');
    if (st === 'CONFIRMED') return;
    if (allowMissing && st === 'MISSING') return;
    issues.push(label + ': ' + st);
  }

  pushIssue('Student', studentMatch, false);
  pushIssue('Instructor', instructorMatch, true);
  pushIssue('Aircraft', aircraftMatch, false);
  pushIssue('From airport', fromAirport, false);
  pushIssue('To airport', toAirport, false);

  return {
    overallStatus: issues.length ? 'REVIEW_REQUIRED' : 'OK',
    issues: issues,
    student: studentMatch,
    instructor: instructorMatch,
    pilot: pilotMatch,
    aircraft: aircraftMatch,
    fromAirport: fromAirport,
    toAirport: toAirport
  };
}

function _histFlightBuildLogPreviewFromStageObject_(stageObj, batchId, existingIds, logHeaders) {
  var sourceFlightId = String((stageObj && stageObj.SOURCE_FLIGHT_ID) || '').trim();
  var rowNo = Number((stageObj && stageObj.ROW_NO) || 0);
  var sourceNote = String((stageObj && stageObj.SOURCE_NOTE) || '').trim();
  var tags = _histFlightReadStructuredTags_(sourceNote);
  var aircraftTypeLookup = _histFlightAircraftTypeLookup_();
  var instructorInitials = String(tags.INSTR_INIT || '').trim().toUpperCase();
  var instructorName = _histFlightInstructorNameFromInitials_(instructorInitials);
  var resolvedAircraftType = _histFlightResolveAircraftType_(aircraftTypeLookup, stageObj && stageObj.ACFT, tags.AIRCRAFT_TYPE_RAW);
  var fuelGalStart = Number(String(tags.FUEL_GAL_START || '').replace(',', '.'));
  var fuelGalEnd = Number(String(tags.FUEL_GAL_END || '').replace(',', '.'));
  var fuelStartL = isFinite(fuelGalStart) ? Math.round(fuelGalStart * 3.78541 * 10) / 10 : '';
  var fuelEndL = isFinite(fuelGalEnd) ? Math.round(fuelGalEnd * 3.78541 * 10) / 10 : '';
  var computedFlightId = sourceFlightId;
  if (!computedFlightId || (existingIds && existingIds[computedFlightId])) {
    computedFlightId = String(batchId || '') + '-' + ('000' + rowNo).slice(-4);
  }

  var out = new Array(LOG_FLIGHT_COL.NUM_TOUCH_AND_GOS + 1).fill('');
  out[LOG_FLIGHT_COL.FLIGHT_ID] = computedFlightId;
  out[LOG_FLIGHT_COL.DATE] = String((stageObj && stageObj.FLIGHT_DATE) || '').trim();
  out[LOG_FLIGHT_COL.PILOT] = String((stageObj && stageObj.PILOT) || '').trim();
  out[LOG_FLIGHT_COL.ACFT] = String((stageObj && stageObj.ACFT) || '').trim().toUpperCase();
  out[LOG_FLIGHT_COL.FROM] = String((stageObj && stageObj.FROM_ICAO) || '').trim().toUpperCase();
  out[LOG_FLIGHT_COL.TO] = String((stageObj && stageObj.TO_ICAO) || '').trim().toUpperCase();
  out[LOG_FLIGHT_COL.TOTAL_TIME] = stageObj && stageObj.TOTAL_TIME_H;
  out[LOG_FLIGHT_COL.START_TACH] = tags.TACH_START || '';
  out[LOG_FLIGHT_COL.END_TACH] = tags.TACH_END || '';
  out[LOG_FLIGHT_COL.FUEL_START] = fuelStartL;
  out[LOG_FLIGHT_COL.FUEL_END] = fuelEndL;
  out[LOG_FLIGHT_COL.BLOCK_OUT] = tags.BRAKES_RELEASE || '';
  out[LOG_FLIGHT_COL.OFF] = tags.AIRBORNE || '';
  out[LOG_FLIGHT_COL.ON] = tags.LANDED || '';
  out[LOG_FLIGHT_COL.BLOCK_IN] = tags.BRAKES_APPLIED || '';
  out[LOG_FLIGHT_COL.NUM_LDGS] = stageObj && stageObj.NUMBER_LDGS;
  out[LOG_FLIGHT_COL.NUM_TOUCH_AND_GOS] = stageObj && stageObj.NUMBER_TOUCH_AND_GOS;
  out[LOG_FLIGHT_COL.SQUAWKS] = sourceNote;
  out[LOG_FLIGHT_COL.ACTUAL_LOAD_JSON] = JSON.stringify({
    entryKind: 'HIST_IMPORT_PREVIEW',
    batchId: String(batchId || ''),
    sourceFlightId: sourceFlightId,
    sourceFileId: String((stageObj && stageObj.SOURCE_FILE_ID) || ''),
    sourceFileUrl: String((stageObj && stageObj.SOURCE_FILE_URL) || ''),
    sourceFileName: String((stageObj && stageObj.SOURCE_FILE_NAME) || ''),
    sourceMimeType: String((stageObj && stageObj.SOURCE_MIME_TYPE) || ''),
    extractionModel: String((stageObj && stageObj.EXTRACTION_MODEL) || ''),
    extractionConfidence: stageObj && stageObj.EXTRACTION_CONFIDENCE,
    trainingCode: String(tags.MISSION_CODE || '').trim(),
    instructorInitials: instructorInitials,
    instructorName: instructorName,
    aircraftType: resolvedAircraftType
  });

  var cols = [];
  for (var i = 0; i < out.length; i++) {
    var colName = (logHeaders && logHeaders[i]) ? String(logHeaders[i]) : ('COL_' + (i + 1));
    cols.push({
      index: i,
      column: colName,
      value: out[i]
    });
  }

  return {
    flightId: computedFlightId,
    columns: cols
  };
}

function _runwayDbRowToObject_(headers, row) {
  var out = {};
  for (var i = 0; i < headers.length; i++) {
    var key = String(headers[i] || '').trim();
    if (!key) continue;
    out[key] = row[i];
  }
  return out;
}

function _runwayOfficialLogIdFromDbRow_(dbRowObj, fallbackId) {
  var row = dbRowObj || {};
  var candidates = [
    'OFFICIAL_LOGID',
    'OFFICIAL_LOG_ID',
    'RUNWAY_LOGID',
    'RUNWAY_LOG_ID',
    'LOGID',
    'LOG_ID',
    'SOURCE_LOGID',
    'SOURCE_LOG_ID'
  ];
  for (var i = 0; i < candidates.length; i++) {
    var key = candidates[i];
    var val = String(row[key] == null ? '' : row[key]).trim();
    if (val) return val;
  }
  return String(fallbackId || '').trim();
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
    'EXPIRY_DATE',
    'UPLOAD_DATE',
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
  var m = raw.match(/\/folders\/([a-zA-Z0-9_-]{10,})/);
  if (m && m[1]) return m[1];
  m = raw.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
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

function _aircraftDocsNormalizeFileName_(tail, docType, effectiveDate, expiryDate, originalName) {
  var t = String(tail || '').trim().toUpperCase();
  var dt = String(docType || '').trim().toUpperCase().replace(/\s+/g, '_').replace(/[^A-Z0-9_-]/g, '') || 'DOC';
  var orig = String(originalName || '');
  var ext = '';
  var dotIdx = orig.lastIndexOf('.');
  if (dotIdx >= 0) ext = orig.slice(dotIdx).toLowerCase();
  var eff = String(effectiveDate || '').trim();
  var exp = String(expiryDate || '').trim();
  var effYear = eff ? eff.slice(0, 4) : '';
  var expYear = exp ? exp.slice(0, 4) : '';
  var dateLabel = '';
  if (effYear && expYear && effYear !== expYear) dateLabel = effYear + '-' + expYear;
  else if (effYear) dateLabel = effYear;
  else if (expYear) dateLabel = expYear;
  var parts = [t, dt];
  if (dateLabel) parts.push(dateLabel);
  return parts.join('_') + ext;
}

function _trainingQualIsoToday_() {
  var d = new Date();
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _trainingQualDateMax_(a, b) {
  if (!a) return b;
  if (!b) return a;
  return a >= b ? a : b;
}

function _trainingQualDateToIso_(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(val).trim().substring(0, 10);
}

function _trainingRoleModuleMap_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(APP_SHEETS.TRAINING_MODULES || 'REF_Training_Modules');
  if (!sh) return {};
  var rows = sh.getDataRange().getValues();
  if (rows.length < 2) return {};
  var h = rows[0].map(function(c) { return String(c).trim().toUpperCase(); });
  var iMod = h.indexOf('MODULE_ID');
  var iRole = h.indexOf('ROLE_CODE');
  var iPrac = h.indexOf('REQUIRES_PRACTICAL');
  var iRec = h.indexOf('RECURRENT_DAYS');
  if (iMod < 0 || iRole < 0) return {};
  var map = {};
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var moduleId = String(r[iMod] || '').trim().toUpperCase();
    var roleCode = String(r[iRole] || '').trim().toUpperCase();
    if (!moduleId || !roleCode) continue;
    if (!map[roleCode]) map[roleCode] = [];
    map[roleCode].push({
      moduleId: moduleId,
      requiresPractical: String(r[iPrac] || '').trim().toUpperCase() === 'TRUE' || r[iPrac] === true || r[iPrac] === 1,
      recurrentDays: parseInt(r[iRec] || '0', 10) || 0
    });
  }
  return map;
}

function _trainingStaffModuleEvidence_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tSh = ss.getSheetByName(APP_SHEETS.STAFF_TRAINING || 'DB_Staff_Training_Records');
  var pSh = ss.getSheetByName(APP_SHEETS.STAFF_PRACTICALS || 'DB_Staff_Practical_Evaluations');
  var evidence = {}; // key: "email|moduleId" → {theoryDate, practicalDate, practicalPass}

  function key_(email, mod) { return String(email).toLowerCase().trim() + '|' + String(mod).toUpperCase().trim(); }

  if (tSh && tSh.getLastRow() > 1) {
    var tRows = tSh.getDataRange().getValues();
    var th = tRows[0].map(function(c) { return String(c).trim().toUpperCase(); });
    var tiEmail = th.indexOf('STAFF_EMAIL'); var tiMod = th.indexOf('MODULE_ID');
    var tiDate = th.indexOf('COMPLETION_DATE'); var tiPass = th.indexOf('RESULT');
    if (tiEmail >= 0 && tiMod >= 0) {
      for (var i = 1; i < tRows.length; i++) {
        var r = tRows[i];
        var passed = String(r[tiPass] || '').trim().toUpperCase();
        if (passed && passed !== 'PASS' && passed !== 'TRUE' && passed !== '1') continue;
        var k = key_(r[tiEmail], r[tiMod]);
        var iso = _trainingQualDateToIso_(tiDate >= 0 ? r[tiDate] : '');
        if (!evidence[k]) evidence[k] = {};
        evidence[k].theoryDate = _trainingQualDateMax_(evidence[k].theoryDate, iso);
      }
    }
  }

  if (pSh && pSh.getLastRow() > 1) {
    var pRows = pSh.getDataRange().getValues();
    var ph = pRows[0].map(function(c) { return String(c).trim().toUpperCase(); });
    var piEmail = ph.indexOf('STAFF_EMAIL'); var piMod = ph.indexOf('MODULE_ID');
    var piDate = ph.indexOf('EVAL_DATE'); var piPass = ph.indexOf('RESULT');
    if (piEmail >= 0 && piMod >= 0) {
      for (var j = 1; j < pRows.length; j++) {
        var pr = pRows[j];
        var pPassed = String(pr[piPass] || '').trim().toUpperCase();
        var k2 = key_(pr[piEmail], pr[piMod]);
        var iso2 = _trainingQualDateToIso_(piDate >= 0 ? pr[piDate] : '');
        if (!evidence[k2]) evidence[k2] = {};
        if (pPassed === 'PASS' || pPassed === 'TRUE' || pPassed === '1') {
          evidence[k2].practicalDate = _trainingQualDateMax_(evidence[k2].practicalDate, iso2);
          evidence[k2].practicalPass = true;
        }
      }
    }
  }
  return evidence;
}

function _trainingRoleEligibilityForStaff_(staffEmail, roleCode, modulesForRole, evidenceMap, asOfDate) {
  if (!modulesForRole || modulesForRole.length === 0) {
    return { eligible: true, managedByTraining: false, validUntil: '', reason: 'No training modules required for this role' };
  }
  var today = asOfDate || _trainingQualIsoToday_();
  var minValidUntil = '';
  for (var i = 0; i < modulesForRole.length; i++) {
    var m = modulesForRole[i];
    var k = String(staffEmail).toLowerCase().trim() + '|' + m.moduleId;
    var ev = evidenceMap[k] || {};
    if (!ev.theoryDate) {
      return { eligible: false, managedByTraining: true, validUntil: '', reason: 'Missing theory: ' + m.moduleId };
    }
    if (m.requiresPractical && !ev.practicalPass) {
      return { eligible: false, managedByTraining: true, validUntil: '', reason: 'Missing practical: ' + m.moduleId };
    }
    if (m.recurrentDays > 0) {
      var baseDate = m.requiresPractical && ev.practicalDate ? ev.practicalDate : ev.theoryDate;
      var expiry = new Date(baseDate);
      expiry.setDate(expiry.getDate() + m.recurrentDays);
      var expiryIso = Utilities.formatDate(expiry, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (expiryIso < today) {
        return { eligible: false, managedByTraining: true, validUntil: expiryIso, reason: 'Recurrency expired for module: ' + m.moduleId };
      }
      minValidUntil = minValidUntil ? (expiryIso < minValidUntil ? expiryIso : minValidUntil) : expiryIso;
    }
  }
  return { eligible: true, managedByTraining: true, validUntil: minValidUntil, reason: 'All modules current' };
}

function _histFlightFindBatchRow_(sheet, batchId) {
  var target = String(batchId || '').trim();
  if (!target) return null;
  var data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return null;
  var headers = data[0].map(_histFlightNormHeader_);
  var batchIdx = headers.indexOf('BATCH_ID');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][batchIdx] || '').trim() === target) {
      return { rowNumber: i + 1, headers: headers, row: data[i] };
    }
  }
  return null;
}

function _histFlightJsonFromModelText_(txt) {
  function stripCodeFence(s) {
    return String(s || '').trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '').trim();
  }
  function extractJsonEnvelope(s) {
    var raw = String(s || '');
    var start = raw.indexOf('{');
    var end = raw.lastIndexOf('}');
    if (start >= 0 && end > start) return raw.slice(start, end + 1).trim();
    return raw.trim();
  }
  var attempts = [];
  var t0 = String(txt || '').trim();
  var t1 = stripCodeFence(t0);
  var t2 = extractJsonEnvelope(t1);
  if (t0) attempts.push(t0);
  if (t1 && t1 !== t0) attempts.push(t1);
  if (t2 && t2 !== t1) attempts.push(t2);
  var lastErr = null;
  for (var i = 0; i < attempts.length; i++) {
    try {
      return { ok: true, value: JSON.parse(attempts[i]), rawText: attempts[i] };
    } catch (e) {
      lastErr = e;
    }
  }
  return { ok: false, error: lastErr, rawText: t2 || t1 || t0 };
}

function _histFlightGeminiExtractFromDocument_(base64Data, mimeType, fileName) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return { success: false, error: 'GEMINI_API_KEY not configured. Add it in GAS Project Settings > Script Properties.' };
  }

  var prompt = [
    'You are extracting historical flight records from a scanned flight sheet.',
    'Read the full image/PDF and extract every flight row you can identify.',
    'Return ONLY valid JSON with this exact shape:',
    '{',
    '  "flights": [',
    '    {',
    '      "sourceFlightId": "",',
    '      "date": "YYYY-MM-DD or empty",',
    '      "pilot": "",',
    '      "acft": "",',
    '      "from": "",',
    '      "to": "",',
    '      "route": "",',
    '      "missionTrainingCode": "",',
    '      "instructorInitials": "",',
    '      "aircraftTypeRaw": "",',
    '      "tachStart": "number or empty",',
    '      "tachEnd": "number or empty",',
    '      "fuelGallonsStart": "number or empty",',
    '      "fuelGallonsEnd": "number or empty",',
    '      "brakesRelease": "HH:MM or empty",',
    '      "airborne": "HH:MM or empty",',
    '      "landed": "HH:MM or empty",',
    '      "brakesApplied": "HH:MM or empty",',
    '      "totalTimeH": "number or HH:MM or empty",',
    '      "numberLdgs": "integer or empty",',
    '      "numberTouchAndGos": "integer or empty",',
    '      "sourceNote": "include uncertainties or remarks",',
    '      "confidence": "0-100"',
    '    }',
    '  ],',
    '  "documentSummary": "short summary of page content"',
    '}',
    'Rules:',
    '- Keep one object per detected flight entry.',
    '- If a field is unreadable, return empty string and mention uncertainty in sourceNote.',
    '- Do not invent values.',
    '- Preserve Portuguese names exactly when possible.'
  ].join('\n');

  var reqBody = {
    contents: [{
      parts: [
        { text: prompt },
        { inlineData: { mimeType: String(mimeType || 'application/pdf'), data: String(base64Data || '') } }
      ]
    }],
    generationConfig: { temperature: 0.1, responseMimeType: 'application/json', maxOutputTokens: 8192 }
  };

  var scriptProps = PropertiesService.getScriptProperties();
  var configuredModel = String(scriptProps.getProperty('GEMINI_MODEL') || '').trim();
  var usedModel = configuredModel || 'gemini-2.5-flash';

  function geminiGenerate(modelId, bodyToSend) {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(modelId) + ':generateContent?key=' + apiKey;
    return UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(bodyToSend || reqBody),
      muteHttpExceptions: true
    });
  }

  var resp = geminiGenerate(usedModel, reqBody);
  var code = resp.getResponseCode();
  var text = resp.getContentText();
  if (code === 404) {
    var fallbackModels = ['gemini-2.5-flash', 'gemini-1.5-flash', 'gemini-2.0-flash'];
    for (var i = 0; i < fallbackModels.length; i++) {
      var candidate = fallbackModels[i];
      if (candidate === usedModel) continue;
      var retry = geminiGenerate(candidate, reqBody);
      if (retry.getResponseCode() === 200) {
        usedModel = candidate;
        resp = retry;
        code = 200;
        text = retry.getContentText();
        break;
      }
    }
  }
  if (code !== 200) {
    var errObj = {};
    try { errObj = JSON.parse(text); } catch (e) {}
    var msg = (errObj && errObj.error && errObj.error.message) ? errObj.error.message : String(text || '').slice(0, 350);
    return { success: false, error: 'Gemini API error (' + code + ', model ' + usedModel + '): ' + msg };
  }

  var parsedResponse = JSON.parse(text);
  var parts = parsedResponse && parsedResponse.candidates && parsedResponse.candidates[0] && parsedResponse.candidates[0].content && parsedResponse.candidates[0].content.parts;
  if (!parts || !parts.length) {
    return { success: false, error: 'No response content from Gemini.' };
  }

  var rawJson = String(parts.map(function(p) { return String((p && p.text) || ''); }).join('\n') || '').trim();
  var parsed = _histFlightJsonFromModelText_(rawJson);
  if (!parsed.ok) {
    var parseMsg = parsed.error && parsed.error.message ? parsed.error.message : 'Unable to parse model JSON output.';
    return { success: false, error: 'Gemini returned malformed JSON: ' + parseMsg, model: usedModel, raw: rawJson };
  }

  var out = parsed.value || {};
  var flights = Array.isArray(out.flights) ? out.flights : [];
  return {
    success: true,
    model: usedModel,
    rawJson: rawJson,
    flights: flights,
    documentSummary: String(out.documentSummary || '')
  };
}

function _histFlightStageRecords_(args) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var batchSheet = getRequiredSheet_(ss, APP_SHEETS.HIST_FLIGHT_IMPORT_BATCHES, '_histFlightStageRecords_');
  var stageSheet = getRequiredSheet_(ss, APP_SHEETS.HIST_FLIGHT_STAGING, '_histFlightStageRecords_');
  var existing = _histFlightReadExistingLogKeys_();

  var batchId = String(args.batchId || '').trim();
  var sourceLabel = String(args.sourceLabel || 'HISTORICAL_IMPORT').trim();
  var importedBy = String(args.importedBy || Session.getActiveUser().getEmail() || 'Admin').trim();
  var file = args.file;
  var fileHash = String(args.fileHash || '').trim();
  var rowsIn = Array.isArray(args.rows) ? args.rows : [];

  var stageData = stageSheet.getDataRange().getValues();
  var nextRowNo = 1;
  if (stageData && stageData.length > 1) {
    var stageHeaders = stageData[0].map(_histFlightNormHeader_);
    var bIdx = stageHeaders.indexOf('BATCH_ID');
    var rIdx = stageHeaders.indexOf('ROW_NO');
    for (var i = 1; i < stageData.length; i++) {
      if (String(stageData[i][bIdx] || '').trim() !== batchId) continue;
      var rv = Number(stageData[i][rIdx] || 0);
      if (rv >= nextRowNo) nextRowNo = rv + 1;
    }
  }

  var staging = [];
  var minDate = '';
  var maxDate = '';
  for (var j = 0; j < rowsIn.length; j++) {
    var src = rowsIn[j] || {};
    var sourceFlightId = String(src.sourceFlightId || '').trim();
    var flightDate = safeDateYmd_(src.date);
    var pilot = String(src.pilot || '').trim();
    var acft = String(src.acft || '').trim().toUpperCase();
    var fromTo = _histFlightResolveFromTo_(src.from, src.to, src.route);
    var totalTimeH = _histFlightParseHours_(src.totalTimeH);
    var numberLdgs = _histFlightParseInt_(src.numberLdgs);
    var touchAndGos = _histFlightParseInt_(src.numberTouchAndGos);
    var sourceNote = String(src.sourceNote || '').trim();
    sourceNote = _histFlightMergeSourceNote_(sourceNote, src);
    var confidence = src.confidence == null ? '' : Number(src.confidence);
    if (!isFinite(confidence)) confidence = '';

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
      nextRowNo++,
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
      file ? file.getId() : '',
      file ? file.getUrl() : '',
      file ? file.getName() : '',
      String(args.mimeType || ''),
      String(args.extractionModel || ''),
      confidence,
      String(args.extractionRawJson || ''),
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

  if (staging.length) {
    stageSheet.getRange(stageSheet.getLastRow() + 1, 1, staging.length, staging[0].length).setValues(staging);
  }

  var batchRow = _histFlightFindBatchRow_(batchSheet, batchId);
  if (!batchRow) {
    batchSheet.appendRow([
      batchId,
      file ? file.getName() : 'DOCUMENT_IMPORT',
      String(args.fileType || 'DOCUMENT'),
      file ? file.getId() : '',
      file ? file.getUrl() : '',
      fileHash,
      new Date().toISOString(),
      importedBy,
      sourceLabel,
      1,
      staging.length,
      'UNDER_REVIEW',
      'Date range: ' + (minDate || '?') + ' to ' + (maxDate || '?')
    ]);
  } else {
    var headers = batchRow.headers;
    var rowNo = batchRow.rowNumber;
    var docCountIdx = headers.indexOf('DOCUMENT_COUNT');
    var rowCountIdx = headers.indexOf('ROW_COUNT');
    var notesIdx = headers.indexOf('NOTES');
    var sourceFileNameIdx = headers.indexOf('SOURCE_FILENAME');
    if (docCountIdx >= 0) {
      var prevDocCount = Number(batchRow.row[docCountIdx] || 0);
      batchSheet.getRange(rowNo, docCountIdx + 1).setValue(prevDocCount + 1);
    }
    if (rowCountIdx >= 0) {
      var prevRows = Number(batchRow.row[rowCountIdx] || 0);
      batchSheet.getRange(rowNo, rowCountIdx + 1).setValue(prevRows + staging.length);
    }
    if (sourceFileNameIdx >= 0 && file && !String(batchRow.row[sourceFileNameIdx] || '').trim()) {
      batchSheet.getRange(rowNo, sourceFileNameIdx + 1).setValue(file.getName());
    }
    if (notesIdx >= 0) {
      var prevNotes = String(batchRow.row[notesIdx] || '').trim();
      var nextNotes = 'Date range: ' + (minDate || '?') + ' to ' + (maxDate || '?');
      batchSheet.getRange(rowNo, notesIdx + 1).setValue(prevNotes ? (prevNotes + ' | ' + nextNotes) : nextNotes);
    }
  }

  return { success: true, batchId: batchId, rowCount: staging.length, minDate: minDate, maxDate: maxDate, sourceFileUrl: file ? file.getUrl() : '' };
}

function _histChunkCache_() {
  return CacheService.getUserCache();
}

function _histChunkKey_(uploadId, suffix) {
  return 'HIST_CHUNK_' + String(uploadId || '') + '_' + String(suffix || '');
}

function _histFlightQueueSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = 'HIST_FLIGHT_IMPORT_QUEUE';
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow([
      'QUEUE_ID',
      'BATCH_ID',
      'FILE_ID',
      'FILE_NAME',
      'MIME_TYPE',
      'SOURCE_LABEL',
      'IMPORTED_BY',
      'STATUS',
      'CREATED_AT',
      'STARTED_AT',
      'FINISHED_AT',
      'ERROR',
      'ROW_COUNT',
      'EXTRACTION_MODEL'
    ]);
  }
  return sheet;
}

function _scheduleHistoricalFlightQueueWorker_() {
  try {
    var fn = 'processHistoricalFlightImportQueue';
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (String(triggers[i].getHandlerFunction() || '') === fn) return;
    }
    ScriptApp.newTrigger(fn).timeBased().after(5000).create();
  } catch (e) {}
}

function _buildMissionMxReview_(aircraftReg, projectedFlightHours, existingReview) {
  var reg = String(aircraftReg || '').trim().toUpperCase();
  var plannedHours = Number(projectedFlightHours);
  if (!isFinite(plannedHours) || plannedHours < 0) plannedHours = 0;
  var prior = (existingReview && typeof existingReview === 'object') ? existingReview : {};
  var scheduleRes = getMaintenanceScheduleData({ aircraftReg: reg });
  var rows = (scheduleRes && scheduleRes.success && Array.isArray(scheduleRes.rows)) ? scheduleRes.rows.slice() : [];
  rows = rows.filter(function(row) {
    return String(row && row.aircraftReg || '').trim().toUpperCase() === reg;
  }).sort(function(a, b) {
    var aHours = Number(a && a.hoursRemaining);
    var bHours = Number(b && b.hoursRemaining);
    var aVal = isFinite(aHours) ? aHours : Number.POSITIVE_INFINITY;
    var bVal = isFinite(bHours) ? bHours : Number.POSITIVE_INFINITY;
    return aVal - bVal;
  });

  var nextInspection = null;
  for (var i = 0; i < rows.length; i++) {
    var hrs = Number(rows[i] && rows[i].hoursRemaining);
    if (isFinite(hrs)) {
      nextInspection = rows[i];
      break;
    }
  }

  var currentHoursRemaining = nextInspection ? Number(nextInspection.hoursRemaining) : null;
  if (!isFinite(currentHoursRemaining)) currentHoursRemaining = null;
  var projectedHoursRemaining = currentHoursRemaining === null ? null : parseFloat((currentHoursRemaining - plannedHours).toFixed(1));
  var thresholdHours = 10;
  var criticalHours = 5;
  var requiresApproval = projectedHoursRemaining !== null && projectedHoursRemaining < thresholdHours;
  var projectedStatus = 'OK';
  var reviewReason = '';
  if (projectedHoursRemaining !== null) {
    if (projectedHoursRemaining < 0) {
      projectedStatus = 'OVERDUE';
      reviewReason = 'Less than 0 hrs remaining after mission.';
    } else if (projectedHoursRemaining < criticalHours) {
      projectedStatus = 'CRITICAL';
      reviewReason = 'Less than 5 hrs remaining after mission.';
    } else if (projectedHoursRemaining < thresholdHours) {
      projectedStatus = 'REVIEW_REQUIRED';
      reviewReason = 'Less than 10 hrs remaining after mission.';
    }
  }

  var reviewStatus = String(prior.status || '').trim().toUpperCase();
  if (!requiresApproval) reviewStatus = 'NOT_REQUIRED';
  else if (reviewStatus !== 'APPROVED' && reviewStatus !== 'REJECTED') reviewStatus = 'PENDING';

  return {
    required: requiresApproval,
    status: reviewStatus || (requiresApproval ? 'PENDING' : 'NOT_REQUIRED'),
    thresholdHours: thresholdHours,
    criticalHours: criticalHours,
    missionFlightHours: parseFloat(plannedHours.toFixed(1)),
    currentHoursRemaining: currentHoursRemaining,
    projectedHoursRemaining: projectedHoursRemaining,
    projectedStatus: projectedStatus,
    reviewReason: reviewReason,
    aircraftReg: reg,
    nextInspection: nextInspection ? {
      assignmentId: String(nextInspection.assignmentId || ''),
      taskCode: String(nextInspection.taskCode || ''),
      taskName: String(nextInspection.taskName || ''),
      dueState: String(nextInspection.dueState || ''),
      currentTach: nextInspection.currentTach,
      nextDueTach: nextInspection.nextDueTach,
      hoursRemaining: nextInspection.hoursRemaining,
      nextDueDate: nextInspection.nextDueDate,
      daysRemaining: nextInspection.daysRemaining
    } : null,
    upcomingInspections: rows.slice(0, 5).map(function(row) {
      return {
        assignmentId: String(row.assignmentId || ''),
        taskCode: String(row.taskCode || ''),
        taskName: String(row.taskName || ''),
        dueState: String(row.dueState || ''),
        currentTach: row.currentTach,
        nextDueTach: row.nextDueTach,
        hoursRemaining: row.hoursRemaining,
        nextDueDate: row.nextDueDate,
        daysRemaining: row.daysRemaining
      };
    }),
    approvedBy: String(prior.approvedBy || '').trim(),
    approvedAt: String(prior.approvedAt || '').trim(),
    notes: String(prior.notes || '').trim(),
    rejectedBy: String(prior.rejectedBy || '').trim(),
    rejectedAt: String(prior.rejectedAt || '').trim(),
    rejectionNotes: String(prior.rejectionNotes || '').trim()
  };
}

function _writeMissionMxReview_(dispatchSheet, missionId, reviewPatch) {
  var sh = dispatchSheet;
  if (!sh) throw new Error('Dispatch sheet not available');
  var missionKey = String(missionId || '').trim();
  if (!missionKey) throw new Error('missionId is required');
  var data = sh.getDataRange().getValues();
  var touched = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][DISPATCH_COL.MISSION_ID] || '').trim() !== missionKey) continue;
    var parsed = _safeParseJsonObject_(data[i][DISPATCH_COL.RAW_DATA]);
    if (!parsed || typeof parsed !== 'object') parsed = {};
    if (!parsed.missionMeta || typeof parsed.missionMeta !== 'object') parsed.missionMeta = {};
    var current = parsed.missionMeta.mxReview && typeof parsed.missionMeta.mxReview === 'object' ? parsed.missionMeta.mxReview : {};
    parsed.missionMeta.mxReview = Object.assign({}, current, reviewPatch || {});
    parsed.mxReview = parsed.missionMeta.mxReview;
    sh.getRange(i + 1, DISPATCH_COL.RAW_DATA + 1).setValue(JSON.stringify(parsed));
    touched++;
  }
  if (!touched) throw new Error('Mission not found: ' + missionKey);
  invalidateScheduledMissionsCache_();
  return touched;
}
