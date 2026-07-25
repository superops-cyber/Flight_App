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

function getToolsFlightFollowConfig() {
  try {
    var ctx = _toolsDutyContext_();
    var cfg = _flightFollowReadConfigMap_();
    return {
      success: true,
      canManageConfig: !!ctx.canManageConfig,
      ffConfig: cfg
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function saveToolsFlightFollowConfig(payload) {
  try {
    var ctx = _toolsDutyContext_();
    if (!ctx.canManageConfig) {
      return { success: false, error: 'Permissão insuficiente para alterar a configuração Flight Following.' };
    }

    var input = (payload && typeof payload === 'object' && payload.config && typeof payload.config === 'object')
      ? payload.config
      : {};
    var defaults = _flightFollowConfigDefaults_();
    var descMap = _flightFollowConfigDescriptions_();
    var entries = Object.keys(defaults).map(function(key) {
      return {
        key: key,
        value: Object.prototype.hasOwnProperty.call(input, key) ? input[key] : defaults[key],
        description: descMap[key] || ''
      };
    });

    var writeRes = _toolsDutySaveConfigMap_(entries, ctx.userEmail || 'tools');
    var cfg = _flightFollowReadConfigMap_();
    return {
      success: true,
      updated: Number(writeRes.updated || 0),
      created: Number(writeRes.created || 0),
      ffConfig: cfg
    };
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
  if (type !== 'FIX') throw new Error('TYPE must be FIX');

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
  const numTgCol = col('NUMBER_TOUCH_AND_GOS', LOG_FLIGHT_COL.NUM_TOUCH_AND_GOS);
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
        landingLog: (function() {
          var raw = actualLoadCol >= 0 ? data[i][actualLoadCol] : '';
          try {
            var parsed = raw ? JSON.parse(String(raw)) : {};
            var list = parsed && parsed.debrief && Array.isArray(parsed.debrief.landingLog) ? parsed.debrief.landingLog : [];
            return list.map(function(entry) {
              return {
                timeZ: String(entry && entry.timeZ || '').trim(),
                icao: String(entry && entry.icao || '').trim().toUpperCase(),
                name: String(entry && entry.name || '').trim(),
                locDesc: String(entry && entry.locDesc || '').trim(),
                type: String(entry && entry.type || 'FULL').trim().toUpperCase()
              };
            }).filter(function(entry) {
              return !!(entry.timeZ || entry.icao || entry.locDesc);
            });
          } catch (e) {
            return [];
          }
        })(),
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

function archiveFlightFollowingSnapshot(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var flightId = String(body.flightId || body.flightLegId || '').trim();
    var missionId = String(body.missionId || '').trim();
    var flightLegId = String(body.flightLegId || '').trim();
    var reg = String(body.reg || '').trim().toUpperCase();
    var reason = String(body.reason || 'MANUAL_ARCHIVE').trim();
    var snapshot = (body.snapshot && typeof body.snapshot === 'object') ? body.snapshot : body;
    if (!flightId) return { success: false, error: 'flightId is required.' };

    var archiveSheet = _flightFollowArchiveSheet_();
    var archiveId = Utilities.getUuid();
    var archivedAt = new Date();
    var archivedBy = String((Session.getActiveUser() && Session.getActiveUser().getEmail && Session.getActiveUser().getEmail()) || '').trim() || 'unknown';

    var snapshotText = '';
    try {
      snapshotText = JSON.stringify(snapshot || {});
    } catch (serr) {
      snapshotText = JSON.stringify({ _serializeError: String(serr && serr.message || serr), _rawType: typeof snapshot });
    }

    var snapshotFileId = '';
    var snapshotFileUrl = '';
    if (snapshotText.length > 45000) {
      var folder = _flightFollowArchiveFolder_();
      var fileName = 'ff_snapshot_' + String(flightId).replace(/[^A-Za-z0-9_-]+/g, '_') + '_' + archiveId + '.json';
      var file = folder.createFile(fileName, snapshotText, MimeType.PLAIN_TEXT);
      snapshotFileId = file.getId();
      snapshotFileUrl = file.getUrl();
      inlineJson = JSON.stringify({ external: true, fileId: snapshotFileId, bytes: snapshotText.length });
    }

    archiveSheet.appendRow([
      archivedAt,
      archiveId,
      flightId,
      missionId,
      flightLegId,
      reg,
      archivedBy,
      reason,
      inlineJson,
      snapshotFileId,
      snapshotFileUrl,
      snapshotText.length
    ]);

    return {
      success: true,
      archiveId: archiveId,
      flightId: flightId,
      archivedAt: archivedAt.toISOString(),
      archivedBy: archivedBy,
      snapshotBytes: snapshotText.length,
      externalized: !!snapshotFileId,
      snapshotFileId: snapshotFileId,
      snapshotFileUrl: snapshotFileUrl
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getToolsFlightFollowingArchive(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var flightId = String(body.flightId || body.flightLegId || '').trim();
    if (!flightId) return { success: false, error: 'flightId is required.' };

    var sheet = _flightFollowArchiveSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'No archived FF snapshots found.' };
    var rows = sheet.getRange(2, 1, lastRow - 1, FF_ARCHIVE_HEADERS_.length).getValues();

    var idx = {
      archivedAt: 0,
      archiveId: 1,
      flightId: 2,
      missionId: 3,
      flightLegId: 4,
      reg: 5,
      archivedBy: 6,
      reason: 7,
      snapshotJson: 8,
      snapshotFileId: 9,
      snapshotFileUrl: 10,
      snapshotSize: 11
    };

    var matches = rows.filter(function(r) {
      return String(r[idx.flightId] || '').trim() === flightId;
    });
    if (!matches.length) return { success: false, error: 'No archived FF snapshot for flight ' + flightId + '.' };

    matches.sort(function(a, b) {
      var at = (a[idx.archivedAt] instanceof Date) ? a[idx.archivedAt].getTime() : new Date(a[idx.archivedAt] || 0).getTime();
      var bt = (b[idx.archivedAt] instanceof Date) ? b[idx.archivedAt].getTime() : new Date(b[idx.archivedAt] || 0).getTime();
      return bt - at;
    });

    var row = matches[0];
    var snapshotText = String(row[idx.snapshotJson] || '');
    var fileId = String(row[idx.snapshotFileId] || '').trim();
    var fileUrl = String(row[idx.snapshotFileUrl] || '').trim();

    if (fileId) {
      try {
        snapshotText = DriveApp.getFileById(fileId).getBlob().getDataAsString();
      } catch (ferr) {
        return { success: false, error: 'Snapshot file unreadable: ' + String(ferr && ferr.message || ferr) };
      }
    }

    var snapshot = {};
    try {
      snapshot = snapshotText ? JSON.parse(snapshotText) : {};
    } catch (jerr) {
      snapshot = { _parseError: String(jerr && jerr.message || jerr), raw: String(snapshotText || '').slice(0, 1000) };
    }

    return {
      success: true,
      archive: {
        archiveId: String(row[idx.archiveId] || ''),
        flightId: String(row[idx.flightId] || ''),
        missionId: String(row[idx.missionId] || ''),
        flightLegId: String(row[idx.flightLegId] || ''),
        reg: String(row[idx.reg] || ''),
        archivedAt: (row[idx.archivedAt] instanceof Date) ? row[idx.archivedAt].toISOString() : String(row[idx.archivedAt] || ''),
        archivedBy: String(row[idx.archivedBy] || ''),
        reason: String(row[idx.reason] || ''),
        snapshotSize: Number(row[idx.snapshotSize] || 0),
        snapshotFileId: fileId,
        snapshotFileUrl: fileUrl
      },
      snapshot: snapshot
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function setFlightFollowMessagesSpreadsheetId(spreadsheetId) {
  var id = String(spreadsheetId || '').trim();
  var ss;
  if (!id) {
    id = String(FF_MESSAGES_SPREADSHEET_ID_DEFAULT_ || '').trim();
    if (id) {
      ss = SpreadsheetApp.openById(id);
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) throw new Error('spreadsheetId is required (no default or active spreadsheet found)');
      id = ss.getId();
    }
  } else {
    ss = SpreadsheetApp.openById(id);
  }
  PropertiesService.getScriptProperties().setProperty(FF_MESSAGES_SPREADSHEET_ID_PROP_, id);
  return {
    success: true,
    property: FF_MESSAGES_SPREADSHEET_ID_PROP_,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName()
  };
}

function getFlightFollowMessagesSpreadsheetTarget() {
  var configuredId = String(PropertiesService.getScriptProperties().getProperty(FF_MESSAGES_SPREADSHEET_ID_PROP_) || '').trim();
  var ss = _flightFollowSpreadsheet_();
  return {
    success: true,
    property: FF_MESSAGES_SPREADSHEET_ID_PROP_,
    configuredId: configuredId || '',
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName()
  };
}

function ensureFlightFollowMessageStore() {
  var sheet = _flightFollowMessagesSheet_();
  return {
    success: true,
    sheetName: sheet.getName(),
    headers: FF_MESSAGES_HEADERS_.slice(),
    rowCount: Math.max(0, sheet.getLastRow() - 1)
  };
}

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

  var recipients = _flightFollowRecipientsList_().join(',');
  var replyInbox = _flightFollowPrimaryInbox_();
  MailApp.sendEmail({
    to: recipients,
    subject: subject,
    body: fullBody,
    replyTo: replyInbox
  });

  return { success: true, compact: compact, subject: subject };
}

function sendFlightFollowMessage(payload) {
  // payload: { reg, missionId, flightLegId, route, sender, text }
  var body = (payload && typeof payload === 'object') ? payload : {};
  var reg = String(body.reg || '').trim().toUpperCase();
  var text = String(body.text || '').trim();
  if (!reg) throw new Error('sendFlightFollowMessage: reg required');
  if (!text) throw new Error('sendFlightFollowMessage: text required');

  var missionId = String(body.missionId || '').trim();
  var flightLegId = String(body.flightLegId || '').trim();
  var route = String(body.route || '').trim();
  var sender = String(body.sender || 'Dispatch UI').trim();
  var isAutoLogInbound = sender.toLowerCase() === 'autolog';
  var directionOverride = String(body.direction || '').trim().toUpperCase();
  if (directionOverride !== 'INBOUND' && directionOverride !== 'OUTBOUND') directionOverride = 'OUTBOUND';
  var recipients = _flightFollowRecipientsList_();
  if (!recipients.length) throw new Error('sendFlightFollowMessage: no configured recipients');
  var outbound = isAutoLogInbound ? '' : _flightFollowOutboundRecipientForReg_(reg);
  var toRecipients = outbound ? [outbound] : recipients;
  var ccRecipients = outbound ? recipients.filter(function(addr) {
    return String(addr || '').trim().toLowerCase() !== String(outbound || '').trim().toLowerCase();
  }) : [];

  var now = new Date();
  var zuluTime = Utilities.formatDate(now, 'UTC', 'HHmm') + 'Z';
  var localTime = Utilities.formatDate(now, 'America/Manaus', 'HH:mm');
  var subject = '[FF CHAT] ' + reg + ' ' + zuluTime + (route ? ' ' + route : '');

  var lines = [
    'FLIGHT FOLLOW MESSAGE',
    'Aircraft : ' + reg,
    'Time     : ' + localTime + ' MAO / ' + zuluTime
  ];
  if (missionId) lines.push('Mission  : ' + missionId);
  if (flightLegId) lines.push('Leg      : ' + flightLegId);
  if (route) lines.push('Route    : ' + route);
  lines.push('From     : ' + sender);
  lines.push('');
  lines.push(text);

  MailApp.sendEmail({
    to: toRecipients.join(','),
    cc: ccRecipients.join(','),
    subject: subject,
    body: lines.join('\n'),
    replyTo: _flightFollowPrimaryInbox_()
  });

  // For AutoLog events we rely on Gmail sync (FF_SYNC) as the single message source
  // to avoid duplicate rows (FF_UI + FF_SYNC) for the same auto-generated event.
  if (!isAutoLogInbound) {
    try {
      _flightFollowAppendMessages_([{
        timestamp: now.toISOString(),
        reg: reg,
        regNorm: reg,
        direction: directionOverride,
        from: Session.getEffectiveUser().getEmail() || sender,
        to: toRecipients.concat(ccRecipients).join(','),
        subject: subject,
        text: text,
        verified: false,
        source: 'FF_UI'
      }]);
    } catch (storeErr) {}
  }

  return {
    success: true,
    to: toRecipients,
    cc: ccRecipients,
    subject: subject,
    timestamp: now.toISOString()
  };
}

function getFlightFollowMessages(reg) {
  var regNorm = _flightFollowNormalizeReg_(reg);
  if (!regNorm) return [];

  var outboundAddress = String(_flightFollowOutboundRecipientForReg_(reg) || '').trim().toLowerCase();
  var tokens = _flightFollowTokenVariants_(regNorm);

  try {
    var sheet = _flightFollowMessagesSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var maxRows = 1500;
    var startRow = Math.max(2, lastRow - maxRows + 1);
    var values = sheet.getRange(startRow, 1, lastRow - startRow + 1, FF_MESSAGES_HEADERS_.length).getValues();
    var threadRegMap = _flightFollowThreadRegMap_(sheet, maxRows);
    var out = [];

    var previewTextForBody = function(body, subject) {
      var raw = String(body || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
      if (!raw) return '';

      var cutMarkers = [
        /\n\s*>\s*On .+ wrote:/i,
        /\s+>\s*On .+ wrote:/i,
        /\nOn .+ wrote:/i,
        /\n-+\s*Original Message\s*-+/i,
        /\nFrom:\s/i,
        /\nSent:\s/i
      ];
      cutMarkers.forEach(function(rx) {
        raw = raw.replace(rx, '\n__FF_QUOTE_BREAK__');
      });
      raw = String(raw.split('__FF_QUOTE_BREAK__')[0] || '').trim();

      var lines = raw.split('\n').map(function(line) {
        return String(line || '').trim();
      }).filter(function(line) {
        return !!line;
      });

      var subjectText = String(subject || '').trim().toLowerCase();
      lines = lines.filter(function(line) {
        var norm = String(line || '').trim().toLowerCase();
        if (!norm) return false;
        if (subjectText && norm === subjectText) return false;
        if (/^re:\s*\[ff chat\]/i.test(line)) return false;
        if (/^\[ff chat\]/i.test(line)) return false;
        return true;
      });

      return lines.join(' ').replace(/\s+/g, ' ').trim();
    };

    values.forEach(function(r) {
      var ts = String(r[0] || '');
      var direction = String(r[3] || '').trim().toUpperCase();
      var threadId = String(r[9] || '').trim();
      var rowRegNorm = _flightFollowNormalizeReg_(r[2]) || _flightFollowNormalizeReg_(threadRegMap[threadId]);
      var from = String(r[4] || '');
      var to = String(r[5] || '');
      var subject = String(r[6] || '');
      var text = String(r[7] || '');
      var verified = String(r[8] || '').toUpperCase() === 'TRUE';
      var fullText = [from, to, subject, text].join('\n');

      var regMatched = rowRegNorm === regNorm || _flightFollowContainsAnyToken_(fullText, tokens);
      var addressMatched = !!outboundAddress && String(fullText || '').toLowerCase().indexOf(outboundAddress) >= 0;
      if (!regMatched && !addressMatched) return;

      out.push({
        timestamp: ts,
        direction: direction || (_flightFollowLooksInreach_(from) ? 'INBOUND' : 'OUTBOUND'),
        from: from || 'unknown',
        subject: subject,
        text: text,
        previewText: previewTextForBody(text, subject),
        verified: verified
      });
    });

    out.sort(function(a, b) {
      return String(a.timestamp || '').localeCompare(String(b.timestamp || ''));
    });
    return out.slice(-200);
  } catch (e) {
    return [];
  }
}

function debugFlightFollowSheetMessages(reg) {
  var regNorm = _flightFollowNormalizeReg_(reg);
  if (!regNorm) {
    return { success: false, reason: 'Missing reg', rows: [] };
  }

  var outboundAddress = String(_flightFollowOutboundRecipientForReg_(reg) || '').trim().toLowerCase();
  var tokens = _flightFollowTokenVariants_(regNorm);

  try {
    var sheet = _flightFollowMessagesSheet_();
    var ss = sheet.getParent();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return {
        success: true,
        reg: regNorm,
        spreadsheetId: ss.getId(),
        spreadsheetName: ss.getName(),
        sheetName: sheet.getName(),
        lastRow: lastRow,
        rowCount: 0,
        matchedCount: 0,
        rows: [],
        tailRows: []
      };
    }

    var maxRows = 1500;
    var startRow = Math.max(2, lastRow - maxRows + 1);
    var values = sheet.getRange(startRow, 1, lastRow - startRow + 1, FF_MESSAGES_HEADERS_.length).getValues();
    var threadRegMap = _flightFollowThreadRegMap_(sheet, maxRows);
    var matched = [];
    var tailStart = Math.max(2, lastRow - 14);
    var tailValues = sheet.getRange(tailStart, 1, lastRow - tailStart + 1, FF_MESSAGES_HEADERS_.length).getValues();
    var tailRows = tailValues.map(function(r, idx) {
      return {
        sheetRow: tailStart + idx,
        timestamp: String(r[0] || ''),
        regNorm: _flightFollowNormalizeReg_(r[2]),
        direction: String(r[3] || ''),
        from: String(r[4] || ''),
        to: String(r[5] || ''),
        subject: String(r[6] || ''),
        threadId: String(r[9] || ''),
        messageId: String(r[10] || ''),
        source: String(r[11] || '')
      };
    });

    values.forEach(function(r, idx) {
      var ts = String(r[0] || '');
      var rowRegNorm = _flightFollowNormalizeReg_(r[2]);
      var direction = String(r[3] || '');
      var from = String(r[4] || '');
      var to = String(r[5] || '');
      var subject = String(r[6] || '');
      var text = String(r[7] || '');
      var threadId = String(r[9] || '').trim();
      var fullText = [from, to, subject, text].join('\n');
      var effectiveRegNorm = rowRegNorm || _flightFollowNormalizeReg_(threadRegMap[threadId]);
      var regMatched = effectiveRegNorm === regNorm || _flightFollowContainsAnyToken_(fullText, tokens);
      var addressMatched = !!outboundAddress && String(fullText || '').toLowerCase().indexOf(outboundAddress) >= 0;
      if (!regMatched && !addressMatched) return;

      matched.push({
        sheetRow: startRow + idx,
        timestamp: ts,
        direction: direction,
        rowRegNorm: rowRegNorm,
        effectiveRegNorm: effectiveRegNorm,
        threadId: threadId,
        from: from,
        to: to,
        subject: subject,
        textPreview: text.slice(0, 160),
        regMatched: regMatched,
        addressMatched: addressMatched
      });
    });

    matched.sort(function(a, b) {
      return String(a.timestamp || '').localeCompare(String(b.timestamp || ''));
    });

    return {
      success: true,
      reg: regNorm,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      sheetName: sheet.getName(),
      lastRow: lastRow,
      outboundAddress: outboundAddress,
      rowCount: values.length,
      matchedCount: matched.length,
      rows: matched.slice(-25),
      tailRows: tailRows
    };
  } catch (e) {
    return { success: false, reg: regNorm, error: String(e && e.message || e) };
  }
}

function debugFlightFollowInbox(reg) {
  // Admin diagnostic helper to confirm mailbox context and matching logic.
  return _flightFollowScanMessages_(reg, true);
}

function debugFlightFollowInboxByToken(token) {
  // Diagnostic helper to inspect inbox visibility by arbitrary token (e.g., KPS).
  var rawToken = String(token || '').trim();
  if (!rawToken) rawToken = 'KPS';

  var variants = [rawToken.toUpperCase()];
  var compact = rawToken.toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (compact && variants.indexOf(compact) < 0) variants.push(compact);
  if (compact.length >= 5) {
    var dashed = compact.slice(0, 2) + '-' + compact.slice(2);
    if (variants.indexOf(dashed) < 0) variants.push(dashed);
  }

  var clauses = variants.map(function(v) { return '"' + v + '"'; });
  var query = 'in:inbox newer_than:30d (' + clauses.join(' OR ') + ')';

  try {
    var threads = GmailApp.search(query, 0, 120);
    var rows = [];

    threads.forEach(function(thread) {
      var msgs = thread.getMessages();
      msgs.forEach(function(msg) {
        if (!msg.isInInbox()) return;
        var from = String(msg.getFrom() || '');
        var to = String(msg.getTo() || '');
        var subj = String(msg.getSubject() || '');
        var body = String(msg.getPlainBody() || '');
        rows.push({
          timestamp: msg.getDate().toISOString(),
          from: from,
          to: to,
          subject: subj,
          bodyPreview: body.substring(0, 180)
        });
      });
    });

    rows.sort(function(a, b) {
      return String(b.timestamp || '').localeCompare(String(a.timestamp || ''));
    });

    return {
      token: rawToken,
      variants: variants,
      query: query,
      effectiveUser: Session.getEffectiveUser().getEmail(),
      activeUser: Session.getActiveUser().getEmail(),
      threadCount: threads.length,
      inboxMessageCount: rows.length,
      messages: rows.slice(0, 200)
    };
  } catch (e) {
    return {
      token: rawToken,
      variants: variants,
      query: query,
      effectiveUser: Session.getEffectiveUser().getEmail(),
      activeUser: Session.getActiveUser().getEmail(),
      threadCount: 0,
      inboxMessageCount: 0,
      messages: [],
      error: String(e && e.message || e)
    };
  }
}

function getFlightFollowInit() {
  var currentUserEmail = '';
  try {
    currentUserEmail = String(Session.getActiveUser().getEmail() || '').trim();
  } catch (e0) {
    currentUserEmail = '';
  }
  if (!currentUserEmail) {
    try {
      currentUserEmail = String(Session.getEffectiveUser().getEmail() || '').trim();
    } catch (e1) {
      currentUserEmail = '';
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('DB_Aircraft');
  if (!sheet) return { aircraft: [], airports: [], waypoints: [], fuelCaches: [], currentUserEmail: currentUserEmail };
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { aircraft: [], airports: [], waypoints: [], fuelCaches: [], currentUserEmail: currentUserEmail };
  var headers = data[0].map(function(h) { return String(h || '').trim().toUpperCase(); });
  var regIdx  = headers.indexOf('REGISTRATION');
  var perfTypeIdx = headers.indexOf('TYPE_FOR_PERFORMANCE');
  var typeIdx = headers.indexOf('AIRCRAFT_TYPE');
  var burnIdx = headers.indexOf('BURN_LPH');
  if (regIdx < 0) return { aircraft: [], currentUserEmail: currentUserEmail };
  var aircraft = [];
  for (var i = 1; i < data.length; i++) {
    var reg = String(data[i][regIdx] || '').trim();
    if (!reg) continue;
    var burnVal = burnIdx >= 0 ? Number(data[i][burnIdx]) : NaN;
    var aircraftTypeRaw = typeIdx >= 0 ? String(data[i][typeIdx] || '').trim() : '';
    var typeForPerformanceRaw = perfTypeIdx >= 0 ? String(data[i][perfTypeIdx] || '').trim() : '';
    var aircraftTypeCanon = String(aircraftTypeRaw || '').trim();
    aircraft.push({
      reg: reg,
      aircraftType: aircraftTypeCanon,
      aircraftTypeRaw: aircraftTypeRaw,
      typeForPerformance: aircraftTypeCanon,
      typeForPerformanceRaw: typeForPerformanceRaw || aircraftTypeRaw,
      speed: 130,
      burn: isFinite(burnVal) ? burnVal : 0
    });
  }

  var airports = [];
  try {
    var airSheet = ss.getSheetByName(APP_SHEETS.AIRPORTS);
    if (airSheet) {
      var airVals = airSheet.getDataRange().getValues();
      if (airVals.length > 1) {
        var airHeaders = airVals[0].map(function(h) { return String(h || '').toUpperCase().trim().replace(/\s+/g, '_'); });
        airports = airVals.slice(1).map(function(r) {
          return _airportMapRow_(airHeaders, r);
        }).filter(function(a) {
          return String(a && a.icao || '').trim();
        });
      }
    }
  } catch (e) {}

  var waypoints = [];
  try {
    var wpSheet = ss.getSheetByName(APP_SHEETS.WAYPOINTS);
    if (wpSheet) {
      var wpVals = wpSheet.getDataRange().getValues();
      if (wpVals.length > 1) {
        var wpHeaders = wpVals[0].map(function(h) { return String(h || '').toUpperCase().trim().replace(/\s+/g, '_'); });
        var wpIdIdx = wpHeaders.indexOf('WP_ID');
        var wpLatIdx = wpHeaders.indexOf('LATITUDE');
        var wpLonIdx = wpHeaders.indexOf('LONGITUDE');
        var wpTypeIdx = wpHeaders.indexOf('TYPE');
        if (wpIdIdx >= 0 && wpLatIdx >= 0 && wpLonIdx >= 0) {
          waypoints = wpVals.slice(1).map(function(r) {
            return {
              wp_id: String(r[wpIdIdx] || '').trim().toUpperCase(),
              lat: Number(r[wpLatIdx]),
              lon: Number(r[wpLonIdx]),
              type: wpTypeIdx >= 0 ? String(r[wpTypeIdx] || '').trim() : ''
            };
          }).filter(function(w) {
            return !!w.wp_id && isFinite(w.lat) && isFinite(w.lon);
          });
        }
      }
    }
  } catch (e2) {}

  var fuelCaches = [];
  try {
    var cacheSheet = ss.getSheetByName(APP_SHEETS.FUEL_CACHES);
    if (cacheSheet) {
      var cacheVals = cacheSheet.getDataRange().getValues();
      if (cacheVals.length > 1) {
        var cacheHeaders = cacheVals[0].map(function(h) { return String(h || '').toUpperCase().trim().replace(/\s+/g, '_'); });
        var icaoIdx = cacheHeaders.indexOf('ICAO');
        var qtyIdx = cacheHeaders.indexOf('CURRENT_QTY');
        if (qtyIdx < 0) qtyIdx = cacheHeaders.indexOf('QTY');
        if (qtyIdx < 0) qtyIdx = cacheHeaders.indexOf('CURRENT_INVENTORY');
        if (qtyIdx < 0) qtyIdx = cacheHeaders.indexOf('INVENTORY');
        if (icaoIdx >= 0 && qtyIdx >= 0) {
          fuelCaches = cacheVals.slice(1).map(function(r) {
            return {
              icao: String(r[icaoIdx] || '').trim().toUpperCase(),
              qty: Number(r[qtyIdx] || 0)
            };
          }).filter(function(c) { return !!c.icao; });
        }
      }
    }
  } catch (e3) {}

  var inreachByReg = _flightFollowMapshareByReg_();

  return {
    aircraft: aircraft,
    airports: airports,
    waypoints: waypoints,
    fuelCaches: fuelCaches,
    inreachByReg: inreachByReg,
    currentUserEmail: currentUserEmail
  };
}

function fetchFlightFollowKml(url) {
  var sourceUrl = String(url || '').trim();
  if (!sourceUrl) return { ok: false, error: 'missing_url' };
  if (!/^https?:\/\//i.test(sourceUrl)) sourceUrl = 'https://' + sourceUrl;

  try {
    var resp = UrlFetchApp.fetch(sourceUrl, {
      followRedirects: true,
      muteHttpExceptions: true,
      headers: {
        Accept: 'application/vnd.google-earth.kml+xml, application/xml;q=0.9, text/xml;q=0.8, */*;q=0.1'
      }
    });
    var status = resp.getResponseCode();
    var text = String(resp.getContentText() || '');
    if (status < 200 || status >= 300) {
      return {
        ok: false,
        sourceUrl: sourceUrl,
        status: status,
        error: 'http_' + status,
        snippet: text.slice(0, 300)
      };
    }
    if (!String(text || '').trim()) {
      return { ok: false, sourceUrl: sourceUrl, status: status, error: 'empty_response' };
    }

    var doc;
    try {
      doc = XmlService.parse(text);
    } catch (parseErr) {
      return {
        ok: false,
        sourceUrl: sourceUrl,
        status: status,
        error: 'parse_error: ' + String(parseErr && parseErr.message || parseErr),
        snippet: text.slice(0, 500)
      };
    }

    var parsed = _flightFollowParseKmlDocument_(doc.getRootElement(), {
      sourceUrl: sourceUrl,
      depth: 0,
      seen: {}
    });
    parsed.ok = true;
    parsed.sourceUrl = sourceUrl;
    parsed.status = status;
    return parsed;
  } catch (e) {
    return {
      ok: false,
      sourceUrl: sourceUrl,
      error: String(e && e.message || e)
    };
  }
}

function diagnoseFlightFollowKml(url) {
  var sourceUrl = String(url || '').trim();
  if (!sourceUrl) return { ok: false, error: 'missing_url' };
  if (!/^https?:\/\//i.test(sourceUrl)) sourceUrl = 'https://' + sourceUrl;

  try {
    var resp = UrlFetchApp.fetch(sourceUrl, {
      followRedirects: true,
      muteHttpExceptions: true,
      headers: {
        Accept: 'application/vnd.google-earth.kml+xml, application/xml;q=0.9, text/xml;q=0.8, */*;q=0.1'
      }
    });
    var status = resp.getResponseCode();
    var text = String(resp.getContentText() || '');
    var out = {
      ok: status >= 200 && status < 300,
      sourceUrl: sourceUrl,
      status: status,
      rawLength: text.length,
      rawPreview: text.slice(0, 6000)
    };
    if (status < 200 || status >= 300) {
      out.error = 'http_' + status;
      return out;
    }
    if (!String(text || '').trim()) {
      out.error = 'empty_response';
      return out;
    }

    try {
      var doc = XmlService.parse(text);
      out.parsed = _flightFollowParseKmlDocument_(doc.getRootElement(), {
        sourceUrl: sourceUrl,
        depth: 0,
        seen: {}
      });
    } catch (parseErr) {
      out.error = 'parse_error: ' + String(parseErr && parseErr.message || parseErr);
    }
    return out;
  } catch (e) {
    return {
      ok: false,
      sourceUrl: sourceUrl,
      error: String(e && e.message || e)
    };
  }
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

  invalidateScheduledMissionsCache_();
  const syncRes = _syncDispatchMissionStatusForFlightLeg_(flightLegId);

  return {
    success: true,
    flightLegId: flightLegId,
    riskTotal: riskTotal,
    onBlocks: now.toISOString(),
    missionStatus: syncRes
  };
}

function recordDebriefLog(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(APP_SHEETS.LOG_FLIGHTS);
  if (!logSheet) throw new Error("Sheet 'LOG_Flights' not found.");

  const normalizeHeader_ = function(v) {
    return String(v || '').toUpperCase().trim().replace(/\s+/g, '_');
  };

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
  const colBrakesRelease = col('BRAKES_RELEASE', LOG_FLIGHT_COL.BRAKES_RELEASE);
  const colAirborne = col('AIRBORNE', LOG_FLIGHT_COL.AIRBORNE);
  const colLanded = col('LANDED', LOG_FLIGHT_COL.LANDED);
  const colNumLdgs = col('NUMBER_LDGS', LOG_FLIGHT_COL.NUM_LDGS);
  const colBrakesApplied = col('BRAKES_APPLIED', LOG_FLIGHT_COL.BRAKES_APPLIED);
  const colTotalTime = col('TOTAL_TIME', LOG_FLIGHT_COL.TOTAL_TIME);
  const colNumTg = col('NUMBER_TOUCH_AND_GOS', LOG_FLIGHT_COL.NUM_TOUCH_AND_GOS);
  const colSquawks = col('SQUAWKS', LOG_FLIGHT_COL.SQUAWKS);
  const colActualLoadJson = col('ACTUAL_LOAD_JSON', LOG_FLIGHT_COL.ACTUAL_LOAD_JSON);

  const endTach = String(payload && payload.endTach || '').trim();
  const fuelEnd = parseFloat(payload && payload.fuelEnd) || 0;
  const fuelUsed = parseFloat(payload && payload.fuelUsed) || 0;
  const brakesRelease = String(payload && payload.brakesRelease || '').trim();
  const airborne = String(payload && payload.airborne || '').trim();
  const landed = String(payload && payload.landed || '').trim();
  const brakesApplied = String(payload && payload.brakesApplied || '').trim();
  const numLdgs = parseInt(payload && payload.numLdgs, 10) || 1;
  const nightLdgs = Math.max(0, parseInt(payload && payload.nightLdgs, 10) || 0);
  const nightHrs = Math.max(0, parseFloat(payload && payload.nightHrs) || 0);
  const ifrHrs = Math.max(0, parseFloat(payload && payload.ifrHrs) || 0);
  const numTouchAndGos = Math.max(0, parseInt(payload && payload.numTouchAndGos, 10) || 0);
  const squawks = String(payload && payload.squawks || '')
    .split(/[\n,;]+/)
    .map(function(s) { return String(s || '').trim(); })
    .filter(Boolean)
    .join(', ');
  const totalTime = String(payload && payload.totalTime || '').trim();
    const computeDurationHhMm_ = function(startHhMm, endHhMm) {
      var parseMinutes = function(v) {
        var m = String(v || '').trim().match(/^(\d{2}):(\d{2})$/);
        if (!m) return null;
        var hh = Number(m[1]);
        var mm = Number(m[2]);
        if (!isFinite(hh) || !isFinite(mm)) return null;
        if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
        return (hh * 60) + mm;
      };

      var startMin = parseMinutes(startHhMm);
      var endMin = parseMinutes(endHhMm);
      if (startMin == null || endMin == null) return '';
      var delta = endMin - startMin;
      if (delta < 0) delta += (24 * 60);
      var hhOut = String(Math.floor(delta / 60)).padStart(2, '0');
      var mmOut = String(delta % 60).padStart(2, '0');
      return hhOut + ':' + mmOut;
    };

    var persistedTotalTime = computeDurationHhMm_(brakesRelease, (colBrakesApplied >= 0 ? brakesApplied : '') || String(data[rowIdx - 1][LOG_FLIGHT_COL.ON_BLOCKS] || '').trim());
    if (!persistedTotalTime) persistedTotalTime = computeDurationHhMm_(airborne, landed);
    if (!persistedTotalTime) persistedTotalTime = totalTime;

  const trainingDebrief = (payload && payload.trainingDebrief && typeof payload.trainingDebrief === 'object')
    ? payload.trainingDebrief
    : null;
  const landingLog = Array.isArray(payload && payload.landingLog)
    ? payload.landingLog.map(function(entry) {
        return {
          timeZ: String(entry && entry.timeZ || '').trim(),
          icao: String(entry && entry.icao || '').trim().toUpperCase(),
          name: String(entry && entry.name || '').trim(),
          locDesc: String(entry && entry.locDesc || '').trim(),
          type: String(entry && entry.type || 'FULL').trim().toUpperCase()
        };
      }).filter(function(entry) {
        return !!(entry.timeZ || entry.icao || entry.locDesc);
      })
    : [];

  if (colEndTach >= 0) logSheet.getRange(rowIdx, colEndTach + 1).setValue(endTach);
  if (colFuelEnd >= 0) logSheet.getRange(rowIdx, colFuelEnd + 1).setValue(fuelEnd);
  if (colFuelUsed >= 0) logSheet.getRange(rowIdx, colFuelUsed + 1).setValue(fuelUsed);
  if (colBrakesRelease >= 0 && brakesRelease) logSheet.getRange(rowIdx, colBrakesRelease + 1).setValue(brakesRelease);
  if (colAirborne >= 0) logSheet.getRange(rowIdx, colAirborne + 1).setValue(airborne);
  if (colLanded >= 0) logSheet.getRange(rowIdx, colLanded + 1).setValue(landed);
  if (colNumLdgs >= 0) logSheet.getRange(rowIdx, colNumLdgs + 1).setValue(numLdgs);
  if (colBrakesApplied >= 0 && brakesApplied) logSheet.getRange(rowIdx, colBrakesApplied + 1).setValue(brakesApplied);
  if (colTotalTime >= 0 && persistedTotalTime) logSheet.getRange(rowIdx, colTotalTime + 1).setValue(persistedTotalTime);
  if (colSquawks >= 0) logSheet.getRange(rowIdx, colSquawks + 1).setValue(squawks);
  logSheet.getRange(rowIdx, colNumTg + 1).setValue(numTouchAndGos);

  if (colActualLoadJson >= 0) {
    const existingRaw = data[rowIdx - 1][colActualLoadJson] || '';
    let existing = {};
    try { existing = existingRaw ? JSON.parse(existingRaw) : {}; } catch (e) { existing = {}; }
    const merged = {
      ...existing,
      debrief: {
        ...(existing.debrief || {}),
        numLdgs: numLdgs,
        numTouchAndGos: numTouchAndGos,
        nightLdgs: nightLdgs,
        nightHrs: nightHrs,
        ifrHrs: ifrHrs,
        dualRecvd: Math.max(0, parseFloat(payload && payload.dualRecvd) || 0),
        dualGiven: Math.max(0, parseFloat(payload && payload.dualGiven) || 0),
        landingLog: landingLog,
        trainingDebrief: trainingDebrief,
        totalTime: totalTime,
        debriefAt: new Date().toISOString()
      }
    };
    logSheet.getRange(rowIdx, colActualLoadJson + 1).setValue(JSON.stringify(merged));
  }

  const endTachNum = parseFloat(endTach);
  if (isFinite(endTachNum)) {
    const aircraftReg = String(data[rowIdx - 1][LOG_FLIGHT_COL.ACFT] || '').trim().toUpperCase();
    const aircraftSheet = ss.getSheetByName(APP_SHEETS.AIRCRAFT);
    if (aircraftSheet && aircraftReg) {
      const acftValues = aircraftSheet.getDataRange().getValues();
      if (acftValues && acftValues.length >= 2) {
        const acftHeaders = (acftValues[0] || []).map(normalizeHeader_);
        const regCol = acftHeaders.indexOf('REGISTRATION');
        const currentTachCol = acftHeaders.indexOf('CURRENT_TACH');
        const nextDueCol = acftHeaders.indexOf('NEXT_DUE_TACH');
        const hoursToInspCol = acftHeaders.indexOf('HOURS_TO_INSPECTION');

        if (regCol >= 0 && currentTachCol >= 0) {
          for (let i = 1; i < acftValues.length; i++) {
            const reg = String(acftValues[i][regCol] || '').trim().toUpperCase();
            if (reg !== aircraftReg) continue;

            const rowNo = i + 1;
            aircraftSheet.getRange(rowNo, currentTachCol + 1).setValue(endTachNum);

            if (nextDueCol >= 0 && hoursToInspCol >= 0) {
              const nextDueNum = parseFloat(acftValues[i][nextDueCol]);
              if (isFinite(nextDueNum)) {
                const hrsRemaining = parseFloat((nextDueNum - endTachNum).toFixed(1));
                aircraftSheet.getRange(rowNo, hoursToInspCol + 1).setValue(hrsRemaining);
              }
            }
            break;
          }
        }
      }
    }
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

  try {
    _dutyPersistOnboardAlunoRowsFromFlight_(ss, data[rowIdx - 1], flightLegId);
  } catch (eOnboard) {
    console.warn('recordDebriefLog: unable to persist onboard aluno duty rows', eOnboard);
  }

  let trainingChecksAdded = 0;
  if (trainingDebrief && String(trainingDebrief.flightResult || '').trim()) {
    const logRow = data[rowIdx - 1] || [];
    const traineeName = String(trainingDebrief.traineeName || payload && payload.pilotName || logRow[LOG_FLIGHT_COL.PILOT] || '').trim();
    trainingChecksAdded = _debr8AppendTrainingCheckouts_(ss, flightLegId, traineeName, trainingDebrief);
    savePilotLessonEvaluation({
      missionId: String(payload && payload.missionId || '').trim(),
      trainingCode: String(trainingDebrief.trainingCode || '').trim(),
      pilot: traineeName,
      aircraft: String(logRow[LOG_FLIGHT_COL.ACFT] || '').trim(),
      missionDate: String(logRow[LOG_FLIGHT_COL.DATE] || '').trim(),
      sourceTab: 'TAB8_DEBRIEF',
      result: String(trainingDebrief.flightResult || '').trim().toUpperCase(),
      resultAt: new Date().toISOString()
    });
  }

  invalidateScheduledMissionsCache_();
  const syncRes = _syncDispatchMissionStatusForFlightLeg_(flightLegId);
  const postFlightEmail = _sendPostFlightReportToStudent_(ss, flightLegId, payload, trainingDebrief);

  return {
    success: true,
    flightLegId: flightLegId,
    endTach: endTach,
    fuelEnd: fuelEnd,
    fuelUsed: fuelUsed,
    brakesRelease: brakesRelease,
    airborne: airborne,
    landed: landed,
    numLdgs: numLdgs,
    nightLdgs: nightLdgs,
    nightHrs: nightHrs,
    ifrHrs: ifrHrs,
    brakesApplied: brakesApplied,
    numTouchAndGos: numTouchAndGos,
    landingLog: landingLog,
    squawks: squawks,
    trainingDebrief: trainingDebrief,
    trainingChecksAdded: trainingChecksAdded,
    totalTime: persistedTotalTime,
    debriefAt: new Date().toISOString(),
    postFlightEmail: postFlightEmail
  };
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

function getToolsFlightDetailReport(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var flightId = String(body.flightId || '').trim();
    if (!flightId) return { success: false, error: 'Flight ID is required.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();

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
    var normText = function(value) {
      return String(value == null ? '' : value).trim();
    };
    var asNumber = function(value, fallback) {
      var n = parseFloat(value);
      if (isNaN(n)) return typeof fallback === 'number' ? fallback : 0;
      return n;
    };
    var asNumberOrNull = function(value) {
      if (value == null || value === '') return null;
      var n = parseFloat(value);
      return isNaN(n) ? null : n;
    };
    var asRateText = function(value) {
      if (value instanceof Date && !isNaN(value.getTime())) {
        return String(value.getMonth() + 1) + '/' + String(value.getDate());
      }
      var raw = String(value == null ? '' : value).trim();
      if (!raw) return '';
      var parsed = new Date(raw);
      if (!isNaN(parsed.getTime()) && /^\d{4}-\d{2}-\d{2}/.test(raw)) {
        return String(parsed.getMonth() + 1) + '/' + String(parsed.getDate());
      }
      return raw;
    };
    var spreadsheetSerialToDate_ = function(value) {
      if (typeof value !== 'number' || !isFinite(value)) return null;
      var base = new Date(1899, 11, 30);
      var wholeDays = Math.floor(value);
      var millis = Math.round((value - wholeDays) * 24 * 60 * 60 * 1000);
      return new Date(base.getTime() + (wholeDays * 24 * 60 * 60 * 1000) + millis);
    };
    var normalizeDateObj_ = function(value) {
      if (value instanceof Date && !isNaN(value.getTime())) return new Date(value.getTime());
      if (typeof value === 'number' && isFinite(value)) {
        var serialDate = spreadsheetSerialToDate_(value);
        if (serialDate && !isNaN(serialDate.getTime())) return serialDate;
      }
      var raw = String(value == null ? '' : value).trim();
      if (!raw) return null;
      if (/^\d+(?:\.\d+)?$/.test(raw)) {
        var serial = spreadsheetSerialToDate_(parseFloat(raw));
        if (serial && !isNaN(serial.getTime())) return serial;
      }
      var parsed = new Date(raw);
      return isNaN(parsed.getTime()) ? null : parsed;
    };
    var dateToIso = function(value) {
      if (!value) return '';
      var dt = normalizeDateObj_(value);
      return dt ? dt.toISOString() : String(value || '').trim();
    };
    var dateToYmd = function(value) {
      var dt = normalizeDateObj_(value);
      if (dt) return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var iso = dateToIso(value);
      return iso && iso.indexOf('T') > 0 ? iso.slice(0, 10) : iso;
    };
    var rowDateToDateObj_ = function(value) {
      return normalizeDateObj_(value);
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
      var startDt = rowDateToDateObj_(startValue);
      var endDt = rowDateToDateObj_(endValue);
      if (!startDt || !endDt) return null;
      var deltaHours = (endDt.getTime() - startDt.getTime()) / 3600000;
      if (!isFinite(deltaHours) || deltaHours <= 0 || deltaHours > 24) return null;
      return deltaHours;
    };
    var formatDurationHhMm_ = function(hoursValue) {
      var hours = parseFloat(hoursValue);
      if (!isFinite(hours) || hours < 0) return '';
      var totalMinutes = Math.round(hours * 60);
      var hh = String(Math.floor(totalMinutes / 60)).padStart(2, '0');
      var mm = String(totalMinutes % 60).padStart(2, '0');
      return hh + ':' + mm;
    };
    var zuluHHMM = function(value) {
      var dt = rowDateToDateObj_(value);
      if (!dt) {
        var raw = String(value || '').trim();
        var hm = raw.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
        if (hm) return ('0' + Number(hm[1] || 0)).slice(-2) + String(hm[2] || '00');
        var compact = raw.match(/^(\d{2})(\d{2})$/);
        if (compact) return compact[1] + compact[2];
        return '';
      }
      return Utilities.formatDate(dt, 'UTC', 'HHmm');
    };
    var truncateText = function(value, maxLen) {
      var text = String(value == null ? '' : value);
      var limit = typeof maxLen === 'number' && maxLen > 0 ? maxLen : 2000;
      return text.length > limit ? text.slice(0, limit) + ' ...[truncated]' : text;
    };
    var toCompactJsonText = function(value, maxLen) {
      if (value == null) return '';
      try {
        return truncateText(JSON.stringify(sanitizeForClient(value, 3)), maxLen || 800);
      } catch (e) {
        return truncateText(String(value), maxLen || 800);
      }
    };
    var sanitizeForClient = function(value, depth) {
      if (depth <= 0) return '[max-depth]';
      if (value == null) return value;
      if (value instanceof Date) return dateToIso(value);
      var valueType = typeof value;
      if (valueType === 'string') return truncateText(value, 4000);
      if (valueType === 'number') return isFinite(value) ? value : null;
      if (valueType === 'boolean') return value;
      if (Array.isArray(value)) {
        return value.slice(0, 50).map(function(item) {
          return sanitizeForClient(item, depth - 1);
        });
      }
      if (valueType === 'object') {
        var out = {};
        Object.keys(value).slice(0, 50).forEach(function(key) {
          out[key] = sanitizeForClient(value[key], depth - 1);
        });
        return out;
      }
      return truncateText(value, 200);
    };

    var dispatchData = readSheet(APP_SHEETS.DISPATCH);
    var logData = readSheet(APP_SHEETS.LOG_FLIGHTS);
    var transData = readSheet(APP_SHEETS.TRANSACTIONS);

    var logFlightIdIdx = indexByAliases(logData.headers, ['FLIGHT_ID'], LOG_FLIGHT_COL.FLIGHT_ID);
    if (logFlightIdIdx < 0) return { success: false, error: 'LOG_Flights is missing FLIGHT_ID column.' };

    var logRow = null;
    for (var li = 0; li < logData.rows.length; li++) {
      if (normText(logData.rows[li][logFlightIdIdx]) === flightId) {
        logRow = logData.rows[li];
        break;
      }
    }
    if (!logRow) return { success: false, error: 'Flight not found in LOG_Flights: ' + flightId };

    var logDateIdx = indexByAliases(logData.headers, ['DATE'], LOG_FLIGHT_COL.DATE);
    var logPilotIdx = indexByAliases(logData.headers, ['PILOT'], LOG_FLIGHT_COL.PILOT);
    var logAircraftIdx = indexByAliases(logData.headers, ['ACFT', 'AIRCRAFT'], LOG_FLIGHT_COL.ACFT);
    var logFromIdx = indexByAliases(logData.headers, ['FROM', 'ORIGIN'], LOG_FLIGHT_COL.FROM);
    var logToIdx = indexByAliases(logData.headers, ['TO', 'DESTINATION'], LOG_FLIGHT_COL.TO);
    var logDistIdx = indexByAliases(logData.headers, ['DISTANCE_NM', 'DIST'], LOG_FLIGHT_COL.DIST);
    var logStartTachIdx = indexByAliases(logData.headers, ['START_TACH'], LOG_FLIGHT_COL.START_TACH);
    var logEndTachIdx = indexByAliases(logData.headers, ['END_TACH'], LOG_FLIGHT_COL.END_TACH);
    var logTotalTimeIdx = indexByAliases(logData.headers, ['TOTAL_TIME', 'FLIGHT_TIME'], LOG_FLIGHT_COL.TOTAL_TIME);
    var logFuelStartIdx = indexByAliases(logData.headers, ['FUEL_START'], LOG_FLIGHT_COL.FUEL_START);
    var logFuelEndIdx = indexByAliases(logData.headers, ['FUEL_END'], LOG_FLIGHT_COL.FUEL_END);
    var logFuelUsedIdx = indexByAliases(logData.headers, ['FUEL_USED'], LOG_FLIGHT_COL.FUEL_USED);
    var logOilIdx = indexByAliases(logData.headers, ['OIL_ADDED', 'OIL'], LOG_FLIGHT_COL.OIL);
    var logVoltsIdx = indexByAliases(logData.headers, ['BATTERY_VOLTS', 'VOLTS'], LOG_FLIGHT_COL.VOLTS);
    var logSquawksIdx = indexByAliases(logData.headers, ['SQUAWKS'], LOG_FLIGHT_COL.SQUAWKS);
    var logToRiskIdx = indexByAliases(logData.headers, ['TO_RISK_MATRIX', 'TAKEOFF_RISK_MATRIX'], LOG_FLIGHT_COL.TO_RISK_MATRIX);
    var logLandingRiskIdx = indexByAliases(logData.headers, ['LANDING_RISK_MATRIX'], LOG_FLIGHT_COL.LANDING_RISK_MATRIX);
    var logBrakesReleaseIdx = indexByAliases(logData.headers, ['BRAKES_RELEASE'], LOG_FLIGHT_COL.BRAKES_RELEASE);
    var logAirborneIdx = indexByAliases(logData.headers, ['AIRBORNE'], LOG_FLIGHT_COL.AIRBORNE);
    var logLandedIdx = indexByAliases(logData.headers, ['LANDED'], LOG_FLIGHT_COL.LANDED);
    var logBrakesAppliedIdx = indexByAliases(logData.headers, ['BRAKES_APPLIED', 'ON_BLOCKS'], LOG_FLIGHT_COL.BRAKES_APPLIED);
    var logOnBlocksIdx = indexByAliases(logData.headers, ['ON_BLOCKS', 'BRAKES_APPLIED'], LOG_FLIGHT_COL.ON_BLOCKS);
    var logNumLdgsIdx = indexByAliases(logData.headers, ['NUMBER_LDGS', 'NUM_LDGS'], LOG_FLIGHT_COL.NUM_LDGS);
    var logNumTgIdx = indexByAliases(logData.headers, ['NUMBER_TOUCH_AND_GOS', 'NUM_TOUCH_AND_GOS'], LOG_FLIGHT_COL.NUM_TOUCH_AND_GOS);
    var logActualLoadIdx = indexByAliases(logData.headers, ['ACTUAL_LOAD_JSON'], LOG_FLIGHT_COL.ACTUAL_LOAD_JSON);

    var actualLoadRaw = (logActualLoadIdx >= 0 ? logRow[logActualLoadIdx] : '') || '';
    var actualLoad = {};
    try {
      actualLoad = actualLoadRaw ? JSON.parse(String(actualLoadRaw)) : {};
    } catch (e) {
      actualLoad = { _parseError: 'Invalid ACTUAL_LOAD_JSON', _raw: String(actualLoadRaw) };
    }

    var dispatchFlightIdIdx = indexByAliases(dispatchData.headers, ['FLIGHT_ID'], DISPATCH_COL.FLIGHT_ID);
    var dispatchMissionIdIdx = indexByAliases(dispatchData.headers, ['MISSION_ID'], DISPATCH_COL.MISSION_ID);
    var dispatchDateIdx = indexByAliases(dispatchData.headers, ['DATE'], DISPATCH_COL.DATE);
    var dispatchAircraftIdx = indexByAliases(dispatchData.headers, ['AIRCRAFT'], DISPATCH_COL.AIRCRAFT);
    var dispatchPilotIdx = indexByAliases(dispatchData.headers, ['PILOT'], DISPATCH_COL.PILOT);
    var dispatchCopilotIdx = indexByAliases(dispatchData.headers, ['COPILOT'], DISPATCH_COL.COPILOT);
    var dispatchRouteIdx = indexByAliases(dispatchData.headers, ['ROUTE'], DISPATCH_COL.ROUTE);
    var dispatchTypeIdx = indexByAliases(dispatchData.headers, ['TYPE', 'FLIGHT_TYPE'], DISPATCH_COL.TYPE);
    var dispatchStatusIdx = indexByAliases(dispatchData.headers, ['STATUS'], DISPATCH_COL.STATUS);
    var dispatchRawIdx = indexByAliases(dispatchData.headers, ['RAW_DATA'], DISPATCH_COL.RAW_DATA);

    var dispatchRow = null;
    for (var di = 0; di < dispatchData.rows.length; di++) {
      if (normText(dispatchData.rows[di][dispatchFlightIdIdx]) === flightId) {
        dispatchRow = dispatchData.rows[di];
        break;
      }
    }

    var transFlightIdx = indexByAliases(transData.headers, ['FLIGHT_ID', 'MISSION_ID'], 0);
    var transFundIdx = indexByAliases(transData.headers, ['FUND', 'FUND_NAME'], 1);
    var transPaxIdx = indexByAliases(transData.headers, ['PASSENGER_NAME', 'NAME'], 2);
    var transCategoryIdx = indexByAliases(transData.headers, ['CATEGORY', 'PAX_CATEGORY'], 3);
    var transWeightIdx = indexByAliases(transData.headers, ['WEIGHT', 'WEIGHT_KG'], 4);
    var transRateIdx = indexByAliases(transData.headers, ['CHARGE_RATE', 'RATE'], 5);
    var transAmountIdx = indexByAliases(transData.headers, ['CHARGED_AMOUNT', 'AMOUNT'], 6);

    var passengersFromTransactions = transData.rows
      .filter(function(row) { return normText(row[transFlightIdx]) === flightId; })
      .map(function(row) {
        return {
          name: normText(row[transPaxIdx]),
          category: normText(row[transCategoryIdx]),
          fund: normText(row[transFundIdx]),
          weightKg: asNumber(row[transWeightIdx], 0),
          chargeRate: asRateText(row[transRateIdx]),
          amount: asNumber(row[transAmountIdx], 0)
        };
      });

    var wb = (actualLoad && typeof actualLoad === 'object' && actualLoad.wb && typeof actualLoad.wb === 'object') ? actualLoad.wb : {};
    var wbSeats = wb && wb.seats && typeof wb.seats === 'object' ? wb.seats : {};
    var passengersFromWb = [];
    Object.keys(wbSeats).forEach(function(seatKey) {
      var seat = wbSeats[seatKey] || {};
      var p = seat.passenger || null;
      if (!p || !normText(p.name)) return;
      passengersFromWb.push({
        source: 'seat',
        seat: normText(seatKey),
        name: normText(p.name),
        weightKg: asNumber((p.actualWeight != null ? p.actualWeight : p.weight), 0),
        plannedWeightKg: asNumber((p.plannedWeight != null ? p.plannedWeight : p.weight), 0),
        cargoKg: asNumber(p.cargo, 0),
        fund: normText(p.fund),
        category: normText(p.category),
        chargeRate: normText(p.chargeRate),
        chargedAmount: asNumber(p.chargedAmount, 0)
      });
    });
    (Array.isArray(wb.waitingPassengers) ? wb.waitingPassengers : []).forEach(function(p) {
      if (!p || !normText(p.name)) return;
      passengersFromWb.push({
        source: 'waiting',
        seat: '',
        name: normText(p.name),
        weightKg: asNumber((p.actualWeight != null ? p.actualWeight : p.weight), 0),
        plannedWeightKg: asNumber((p.plannedWeight != null ? p.plannedWeight : p.weight), 0),
        cargoKg: asNumber(p.cargo, 0),
        fund: normText(p.fund),
        category: normText(p.category),
        chargeRate: normText(p.chargeRate),
        chargedAmount: asNumber(p.chargedAmount, 0)
      });
    });

    var rawFields = [];
    for (var hi = 0; hi < logData.headers.length; hi++) {
      var fieldName = String(logData.headers[hi] || '').trim();
      if (!fieldName) fieldName = 'COL_' + (hi + 1);
      var rawVal = (hi < logRow.length) ? logRow[hi] : '';
      var fieldNorm = _toolsNormHeader_(fieldName);
      var fieldText = (rawVal instanceof Date) ? rawVal.toISOString() : String(rawVal == null ? '' : rawVal);
      if (fieldNorm === 'ACTUAL_LOAD_JSON') fieldText = '[omitted from compact payload; bytes=' + String(actualLoadRaw || '').length + ']';
      rawFields.push({
        field: fieldName,
        value: truncateText(fieldText, 240)
      });
    }

    var wbGraph = {
      envelopeData: Array.isArray(wb.envelopeData) ? wb.envelopeData.slice(0, 24).map(function(point) {
        return {
          x: asNumberOrNull(point && (point.CG_Arm_X != null ? point.CG_Arm_X : point.x)),
          y: asNumberOrNull(point && (point.Weight_Y != null ? point.Weight_Y : point.y))
        };
      }).filter(function(point) {
        return point.x != null && point.y != null && isFinite(point.x) && isFinite(point.y);
      }) : [],
      point: {
        x: asNumberOrNull(wb.cgPosition != null ? wb.cgPosition : wb.cg),
        y: asNumberOrNull(wb.grossWeight != null ? wb.grossWeight : wb.totalWeight)
      }
    };

    // If this flight did not save envelope points in ACTUAL_LOAD_JSON, try live reference envelope by aircraft.
    if (!wbGraph.envelopeData.length) {
      try {
        var envAircraft = normText((dispatchRow && dispatchAircraftIdx >= 0) ? dispatchRow[dispatchAircraftIdx] : '')
          || normText(logAircraftIdx >= 0 ? logRow[logAircraftIdx] : '');
        if (envAircraft) {
          var envRef = getWbEnvelopeByAircraft(envAircraft);
          wbGraph.envelopeData = (envRef && Array.isArray(envRef.envelopeData) ? envRef.envelopeData : []).slice(0, 24).map(function(point) {
            return {
              x: asNumberOrNull(point && (point.CG_Arm_X != null ? point.CG_Arm_X : point.x)),
              y: asNumberOrNull(point && (point.Weight_Y != null ? point.Weight_Y : point.y))
            };
          }).filter(function(point) {
            return point.x != null && point.y != null && isFinite(point.x) && isFinite(point.y);
          });
        }
      } catch (envErr) {
        // Keep report resilient; envelope fallback is best-effort only.
      }
    }

    var debriefTab8 = (actualLoad && actualLoad.debrief && typeof actualLoad.debrief === 'object') ? actualLoad.debrief : {};
    var arrivalData = (actualLoad && actualLoad.arrival && typeof actualLoad.arrival === 'object') ? actualLoad.arrival : {};
    var takeoffRollData = (actualLoad && actualLoad.takeoffRoll && typeof actualLoad.takeoffRoll === 'object') ? actualLoad.takeoffRoll : {};
    var dispatchRaw = {};
    if (dispatchRow && dispatchRawIdx >= 0) {
      try {
        dispatchRaw = JSON.parse(String(dispatchRow[dispatchRawIdx] || '{}'));
      } catch (drErr) {
        dispatchRaw = {};
      }
    }
    var dispatchTraining = (dispatchRaw && dispatchRaw.training && typeof dispatchRaw.training === 'object')
      ? dispatchRaw.training
      : ((dispatchRaw && dispatchRaw.meta && dispatchRaw.meta.training && typeof dispatchRaw.meta.training === 'object')
        ? dispatchRaw.meta.training
        : {});
    var trainingDebriefRaw = debriefTab8.trainingDebrief;
    var trainingDebriefObj = {};
    if (trainingDebriefRaw && typeof trainingDebriefRaw === 'object') {
      trainingDebriefObj = trainingDebriefRaw;
    } else if (trainingDebriefRaw) {
      try {
        var parsedTrainingDebrief = JSON.parse(String(trainingDebriefRaw));
        if (parsedTrainingDebrief && typeof parsedTrainingDebrief === 'object') trainingDebriefObj = parsedTrainingDebrief;
      } catch (tdErr) {
        trainingDebriefObj = {};
      }
    }
    var missionIdValue = normText((dispatchRow && dispatchMissionIdIdx >= 0) ? dispatchRow[dispatchMissionIdIdx] : '');
    if (!missionIdValue && typeof missionIdFromFlightLeg_ === 'function') {
      missionIdValue = String(missionIdFromFlightLeg_(flightId) || '').trim();
    }
    var trainingCodeValue = normText(trainingDebriefObj.trainingCode || dispatchTraining.trainingCode || dispatchTraining.code || '');
    var lessonNameValue = normText(trainingDebriefObj.lessonName || dispatchTraining.lessonName || dispatchTraining.externalLessonName || '');
    var traineeNameValue = normText(trainingDebriefObj.traineeName || dispatchTraining.student || dispatchTraining.traineeName || '');
    var instructorNameValue = normText(trainingDebriefObj.instructorName || dispatchTraining.instructor || dispatchTraining.instructorName || ((dispatchRow && dispatchPilotIdx >= 0) ? dispatchRow[dispatchPilotIdx] : ''));
    var copilotNameValue = normText((actualLoad && (actualLoad.copilot || actualLoad.copilotName || (actualLoad.crew && actualLoad.crew.copilot))) || ((dispatchRow && dispatchCopilotIdx >= 0) ? dispatchRow[dispatchCopilotIdx] : ''));
    var runwayChecksList = Array.isArray(trainingDebriefObj.runwayChecks) ? trainingDebriefObj.runwayChecks : [];
    var runwayChecksPassed = runwayChecksList.filter(function(item) {
      return !!(item && (item.ok === true || String(item.ok || '').toUpperCase() === 'Y'));
    }).length;

    var lessonEvalSummary = null;
    if (missionIdValue && trainingCodeValue && typeof getPilotLessonEvaluation === 'function') {
      try {
        var evalRes = getPilotLessonEvaluation(missionIdValue, trainingCodeValue);
        if (evalRes && evalRes.success && evalRes.found) {
          var evalScores = (evalRes.scores && typeof evalRes.scores === 'object') ? evalRes.scores : {};
          var evalValues = Object.keys(evalScores).map(function(k) { return Number(evalScores[k] || 0); }).filter(function(v) {
            return isFinite(v) && v > 0;
          });
          var evalAvg = evalValues.length ? Number((evalValues.reduce(function(a, b) { return a + b; }, 0) / evalValues.length).toFixed(2)) : null;
          lessonEvalSummary = {
            found: true,
            status: normText(evalRes.status || ''),
            result: normText(evalRes.result || ''),
            scoreCount: evalValues.length,
            averageScore: evalAvg,
            submittedAt: normText(evalRes.submittedAt || ''),
            reviewer: normText(evalRes.reviewer || ''),
            notesCount: (evalRes.notes && typeof evalRes.notes === 'object') ? Object.keys(evalRes.notes).length : 0
          };
        }
      } catch (evalErr) {
        lessonEvalSummary = null;
      }
    }

    if (!traineeNameValue) {
      for (var txi = 0; txi < passengersFromTransactions.length; txi++) {
        var txItem = passengersFromTransactions[txi] || {};
        var txCat = String(txItem.category || '').trim().toUpperCase();
        if (txCat.indexOf('STUDENT') >= 0 && String(txItem.name || '').trim()) {
          traineeNameValue = String(txItem.name || '').trim();
          break;
        }
      }
    }

    var routeRawValue = normText((dispatchRow && dispatchRouteIdx >= 0) ? dispatchRow[dispatchRouteIdx] : '');
    var routeSplitValue = splitRoute_(routeRawValue);
    var routeTokens = [];
    if (routeSplitValue && routeSplitValue.from) routeTokens.push(String(routeSplitValue.from || '').trim().toUpperCase());
    routeTokens = routeTokens.concat(routeWaypointsFromRouteString_(routeRawValue).map(function(token) {
      return String(token || '').trim().toUpperCase();
    }));
    if (routeSplitValue && routeSplitValue.to) routeTokens.push(String(routeSplitValue.to || '').trim().toUpperCase());
    routeTokens = routeTokens.filter(Boolean);
    var routeDisplay = routeTokens.length ? routeTokens.join(' - ') : routeRawValue;

    var studentFundsSummary = {
      studentName: traineeNameValue,
      studentKey: '',
      balanceBrl: null,
      lastTransactions: []
    };
    if (traineeNameValue) {
      try {
        var pilotsSh = ss.getSheetByName(APP_SHEETS.PILOTS);
        if (pilotsSh) {
          var pilotsHeaders = _toolsEnsurePilotStudentHeaders_(pilotsSh);
          var pilotsVals = pilotsSh.getDataRange().getValues();
          var nameIdxPilot = _toolsHeaderIndexFromCandidates_(pilotsHeaders, ['PILOT_NAME', 'PILOT', 'NAME', 'STUDENT_NAME', 'NOME']);
          var matchedRec = null;
          var traineeNorm = String(traineeNameValue || '').trim().toUpperCase();
          for (var pi = 1; pi < pilotsVals.length; pi++) {
            var rn = nameIdxPilot >= 0 ? String(pilotsVals[pi][nameIdxPilot] || '').trim().toUpperCase() : '';
            if (rn && rn === traineeNorm) {
              matchedRec = _toolsStudentRecordFromRow_(pilotsHeaders, pilotsVals[pi], pi + 1);
              break;
            }
          }
          if (matchedRec) {
            var studentKeyValue = String(matchedRec.studentKey || _toolsStudentIdentityKey_(matchedRec) || '').trim();
            var allLedgers = _toolsStudentLedgerMap_(ss);
            var ledgerRows = allLedgers[studentKeyValue] || [];
            var liveRows = ledgerRows.filter(function(tx) {
              var statusKey = String(tx && tx.status || '').trim().toUpperCase();
              return statusKey !== 'VOIDED';
            });
            var balanceValue = _toolsStudentCurrentBalanceFromLedger_(ss, studentKeyValue, Number(matchedRec.balanceBrl || 0));
            studentFundsSummary.studentName = matchedRec.name || traineeNameValue;
            studentFundsSummary.studentKey = studentKeyValue;
            studentFundsSummary.balanceBrl = Number(isFinite(balanceValue) ? balanceValue : 0);
            studentFundsSummary.lastTransactions = liveRows.slice(-3).reverse().map(function(tx) {
              return {
                txId: String(tx && tx.txId || '').trim(),
                timestamp: String(tx && tx.timestamp || '').trim(),
                txType: String(tx && tx.txType || '').trim().toUpperCase(),
                amountBrl: Number(tx && tx.amount || 0),
                rateBrlHour: Number(tx && tx.rate || 0),
                note: String(tx && tx.note || '').trim(),
                referenceType: String(tx && tx.referenceType || '').trim().toUpperCase(),
                referenceId: String(tx && tx.referenceId || '').trim(),
                balanceAfterBrl: Number(tx && tx.balanceAfter || 0),
                receiptUrl: String(tx && tx.receiptUrl || '').trim()
              };
            });
          }
        }
      } catch (sfErr) {
        // Optional section for reporting; keep the rest of report intact.
      }
    }

    var totalTimeHours = durationHoursFromAny(logTotalTimeIdx >= 0 ? logRow[logTotalTimeIdx] : null);
    if (!(totalTimeHours > 0)) {
      totalTimeHours = durationHoursBetween(
        logBrakesReleaseIdx >= 0 ? logRow[logBrakesReleaseIdx] : null,
        (logBrakesAppliedIdx >= 0 ? logRow[logBrakesAppliedIdx] : null) || (logOnBlocksIdx >= 0 ? logRow[logOnBlocksIdx] : null)
      );
    }
    if (!(totalTimeHours > 0)) {
      totalTimeHours = durationHoursBetween(
        logAirborneIdx >= 0 ? logRow[logAirborneIdx] : null,
        logLandedIdx >= 0 ? logRow[logLandedIdx] : null
      );
    }
    var totalTimeDisplay = (totalTimeHours > 0) ? formatDurationHhMm_(totalTimeHours) : '';
    var dualRecvdHours = asNumber(debriefTab8.dualRecvd, 0);
    var dualGivenHours = asNumber(debriefTab8.dualGiven, 0);
    var dualRecvdDisplay = dualRecvdHours > 0 ? formatDurationHhMm_(dualRecvdHours) : '';
    var dualGivenDisplay = dualGivenHours > 0 ? formatDurationHhMm_(dualGivenHours) : '';
    var instructionTimeDisplay = dualGivenDisplay;
    var lessonPlanSummary = {
      missionId: missionIdValue,
      trainingCode: trainingCodeValue,
      lessonName: lessonNameValue,
      traineeName: traineeNameValue,
      instructorName: instructorNameValue,
      flightResult: normText(trainingDebriefObj.flightResult || ''),
      runwayChecksTotal: runwayChecksList.length,
      runwayChecksPassed: runwayChecksPassed,
      evaluation: lessonEvalSummary
    };
    var landingTimeline = (function() {
      var lines = [];
      var list = Array.isArray(debriefTab8.landingLog) ? debriefTab8.landingLog : [];
      if (list.length) {
        list.forEach(function(entry) {
          var typeCode = String(entry && entry.type || '').trim().toUpperCase();
          var label = typeCode === 'TNG' ? 'Touch-and-go' : 'Landing';
          var loc = String(entry && (entry.icao || entry.locDesc || entry.name) || '').trim().toUpperCase();
          if (!loc) loc = normText(logToIdx >= 0 ? logRow[logToIdx] : '').toUpperCase() || '---';
          var time = String(entry && entry.timeZ || '').trim();
          if (!time) time = zuluHHMM(entry && entry.time);
          var text = label + ' at ' + loc;
          if (time) text += ' ' + time + 'Z';
          lines.push(text);
        });
        return lines;
      }
      var tngCount = Math.max(0, Math.round(asNumber(logNumTgIdx >= 0 ? logRow[logNumTgIdx] : 0, 0)));
      var landingCount = Math.max(0, Math.round(asNumber(logNumLdgsIdx >= 0 ? logRow[logNumLdgsIdx] : 0, 0)));
      var loc = normText(logToIdx >= 0 ? logRow[logToIdx] : '').toUpperCase() || '---';
      var time = zuluHHMM(logLandedIdx >= 0 ? logRow[logLandedIdx] : '');
      for (var i = 0; i < tngCount; i++) {
        lines.push('Touch-and-go at ' + loc + (time ? (' ' + time + 'Z') : ''));
      }
      if (landingCount > 0 || (logLandedIdx >= 0 && logRow[logLandedIdx])) {
        lines.push('Landing at ' + loc + (time ? (' ' + time + 'Z') : ''));
      }
      return lines;
    })();

    var response = {
      success: true,
      flightId: flightId,
      generatedAt: new Date().toISOString(),
      appUrl: ScriptApp.getService().getUrl() || '',
      summary: {
        date: dateToYmd(logDateIdx >= 0 ? logRow[logDateIdx] : ''),
        pilot: normText(logPilotIdx >= 0 ? logRow[logPilotIdx] : ''),
        copilot: copilotNameValue,
        aircraft: normText(logAircraftIdx >= 0 ? logRow[logAircraftIdx] : ''),
        from: normText(logFromIdx >= 0 ? logRow[logFromIdx] : ''),
        to: normText(logToIdx >= 0 ? logRow[logToIdx] : ''),
        route: routeDisplay,
        flightType: normText((dispatchRow && dispatchTypeIdx >= 0) ? dispatchRow[dispatchTypeIdx] : ''),
        status: normText((dispatchRow && dispatchStatusIdx >= 0) ? dispatchRow[dispatchStatusIdx] : ''),
        totalTime: normText(logTotalTimeIdx >= 0 ? logRow[logTotalTimeIdx] : ''),
        totalTimeDisplay: totalTimeDisplay,
        fuelUsed: asNumber(logFuelUsedIdx >= 0 ? logRow[logFuelUsedIdx] : 0, 0),
        landings: asNumber(logNumLdgsIdx >= 0 ? logRow[logNumLdgsIdx] : 0, 0)
      },
      log: {
        date: dateToYmd(logDateIdx >= 0 ? logRow[logDateIdx] : ''),
        pilot: normText(logPilotIdx >= 0 ? logRow[logPilotIdx] : ''),
        copilot: copilotNameValue,
        aircraft: normText(logAircraftIdx >= 0 ? logRow[logAircraftIdx] : ''),
        from: normText(logFromIdx >= 0 ? logRow[logFromIdx] : ''),
        to: normText(logToIdx >= 0 ? logRow[logToIdx] : ''),
        distanceNm: asNumber(logDistIdx >= 0 ? logRow[logDistIdx] : 0, 0),
        startTach: normText(logStartTachIdx >= 0 ? logRow[logStartTachIdx] : ''),
        endTach: normText(logEndTachIdx >= 0 ? logRow[logEndTachIdx] : ''),
        totalTime: normText(logTotalTimeIdx >= 0 ? logRow[logTotalTimeIdx] : ''),
        totalTimeHours: (totalTimeHours > 0) ? Number(totalTimeHours.toFixed(3)) : null,
        totalTimeDisplay: totalTimeDisplay,
        fuelStart: asNumber(logFuelStartIdx >= 0 ? logRow[logFuelStartIdx] : 0, 0),
        fuelEnd: asNumber(logFuelEndIdx >= 0 ? logRow[logFuelEndIdx] : 0, 0),
        fuelUsed: asNumber(logFuelUsedIdx >= 0 ? logRow[logFuelUsedIdx] : 0, 0),
        oilAdded: asNumber(logOilIdx >= 0 ? logRow[logOilIdx] : 0, 0),
        batteryVolts: asNumber(logVoltsIdx >= 0 ? logRow[logVoltsIdx] : 0, 0),
        squawks: normText(logSquawksIdx >= 0 ? logRow[logSquawksIdx] : ''),
        toRiskMatrix: normText(logToRiskIdx >= 0 ? logRow[logToRiskIdx] : ''),
        landingRiskMatrix: normText(logLandingRiskIdx >= 0 ? logRow[logLandingRiskIdx] : ''),
        brakesRelease: dateToIso(logBrakesReleaseIdx >= 0 ? logRow[logBrakesReleaseIdx] : ''),
        airborne: dateToIso(logAirborneIdx >= 0 ? logRow[logAirborneIdx] : ''),
        landed: dateToIso(logLandedIdx >= 0 ? logRow[logLandedIdx] : ''),
        brakesApplied: dateToIso(logBrakesAppliedIdx >= 0 ? logRow[logBrakesAppliedIdx] : ''),
        onBlocks: dateToIso(logOnBlocksIdx >= 0 ? logRow[logOnBlocksIdx] : ''),
        numLandings: asNumber(logNumLdgsIdx >= 0 ? logRow[logNumLdgsIdx] : 0, 0),
        numTouchAndGos: asNumber(logNumTgIdx >= 0 ? logRow[logNumTgIdx] : 0, 0)
      },
      passengers: {
        transactions: passengersFromTransactions.slice(0, 30),
        wbManifest: passengersFromWb.slice(0, 30)
      },
      wb: {
        grossWeight: asNumberOrNull(wb.grossWeight != null ? wb.grossWeight : wb.totalWeight),
        cgPosition: asNumberOrNull(wb.cgPosition != null ? wb.cgPosition : wb.cg),
        isSafe: (typeof wb.isSafe === 'boolean') ? wb.isSafe : null,
        fuelKg: asNumberOrNull(wb.fuel),
        savedAt: normText(wb.savedAt),
        graph: wbGraph
      },
      debrief: {
        tab8: {
          totalTime: normText(debriefTab8.totalTime),
          totalTimeHours: (totalTimeHours > 0) ? Number(totalTimeHours.toFixed(3)) : null,
          totalTimeDisplay: totalTimeDisplay,
          nightLdgs: asNumber(debriefTab8.nightLdgs, 0),
          nightHrs: asNumber(debriefTab8.nightHrs, 0),
          ifrHrs: asNumber(debriefTab8.ifrHrs, 0),
          dualRecvd: dualRecvdHours,
          dualGiven: dualGivenHours,
          dualRecvdDisplay: dualRecvdDisplay,
          dualGivenDisplay: dualGivenDisplay,
          instructionTimeDisplay: instructionTimeDisplay,
          landingLog: sanitizeForClient(Array.isArray(debriefTab8.landingLog) ? debriefTab8.landingLog : [], 3),
          debriefAt: normText(debriefTab8.debriefAt),
          trainingDebrief: sanitizeForClient(trainingDebriefObj, 3),
          trainingDebriefText: toCompactJsonText(trainingDebriefObj || {}, 500)
        },
        arrival: toCompactJsonText(arrivalData, 700),
        takeoffRoll: toCompactJsonText(takeoffRollData, 700)
      },
      lessonPlanSummary: lessonPlanSummary,
      lessonName: lessonNameValue,
      routeDisplay: routeDisplay,
      totalTimeDisplay: totalTimeDisplay,
      dualRecvdDisplay: dualRecvdDisplay,
      dualGivenDisplay: dualGivenDisplay,
      instructionTimeDisplay: instructionTimeDisplay,
      studentFunds: studentFundsSummary,
      landingTimeline: landingTimeline,
      rawFields: rawFields.slice(0, 40),
      debug: {
        rawFieldCount: rawFields.length,
        rawFieldsReturned: Math.min(rawFields.length, 40),
        transactionsCount: passengersFromTransactions.length,
        wbManifestCount: passengersFromWb.length,
        actualLoadBytes: String(actualLoadRaw || '').length,
        graphPointsReturned: wbGraph.envelopeData.length
      }
    };

    try {
      var ffArchiveRes = getToolsFlightFollowingArchive({ flightId: flightId });
      if (ffArchiveRes && ffArchiveRes.success) {
        response.ffArchive = ffArchiveRes.archive || null;
      } else {
        response.ffArchive = null;
      }
    } catch (ffArcErr) {
      response.ffArchive = null;
    }

    response.debug.responseBytes = JSON.stringify(response).length;
    return response;
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function getToolsFlightDetailProbe(payload) {
  try {
    var body = (payload && typeof payload === 'object') ? payload : {};
    var flightId = String(body.flightId || '').trim();
    if (!flightId) return { success: false, error: 'Flight ID is required.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var readSheet = function(name) {
      var sh = ss.getSheetByName(name);
      if (!sh) return { name: name, found: false, headers: [], rows: [] };
      var values = sh.getDataRange().getValues();
      if (!values || !values.length) return { name: name, found: true, headers: [], rows: [] };
      return {
        name: name,
        found: true,
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
    var normText = function(value) {
      return String(value == null ? '' : value).trim();
    };
    var previewRow = function(headers, row) {
      if (!row) return null;
      var out = {};
      for (var i = 0; i < headers.length && i < row.length && i < 16; i++) {
        var key = String(headers[i] || '').trim() || ('COL_' + (i + 1));
        var cell = row[i];
        out[key] = cell instanceof Date ? cell.toISOString() : String(cell == null ? '' : cell);
      }
      return out;
    };

    var logData = readSheet(APP_SHEETS.LOG_FLIGHTS);
    var dispatchData = readSheet(APP_SHEETS.DISPATCH);
    var transData = readSheet(APP_SHEETS.TRANSACTIONS);

    var logFlightIdIdx = indexByAliases(logData.headers, ['FLIGHT_ID'], LOG_FLIGHT_COL.FLIGHT_ID);
    var dispatchFlightIdIdx = indexByAliases(dispatchData.headers, ['FLIGHT_ID'], DISPATCH_COL.FLIGHT_ID);
    var transFlightIdIdx = indexByAliases(transData.headers, ['FLIGHT_ID', 'MISSION_ID'], 0);

    var logMatchIndex = -1;
    for (var li = 0; li < logData.rows.length; li++) {
      if (normText(logData.rows[li][logFlightIdIdx]) === flightId) {
        logMatchIndex = li;
        break;
      }
    }

    var dispatchMatchIndex = -1;
    for (var di = 0; di < dispatchData.rows.length; di++) {
      if (normText(dispatchData.rows[di][dispatchFlightIdIdx]) === flightId) {
        dispatchMatchIndex = di;
        break;
      }
    }

    var transMatchCount = 0;
    for (var ti = 0; ti < transData.rows.length; ti++) {
      if (normText(transData.rows[ti][transFlightIdIdx]) === flightId) transMatchCount++;
    }

    return {
      success: true,
      flightId: flightId,
      sheets: {
        log: {
          name: APP_SHEETS.LOG_FLIGHTS,
          found: !!logData.found,
          headerCount: logData.headers.length,
          rowCount: logData.rows.length,
          flightIdIdx: logFlightIdIdx,
          headersPreview: logData.headers.slice(0, 20)
        },
        dispatch: {
          name: APP_SHEETS.DISPATCH,
          found: !!dispatchData.found,
          headerCount: dispatchData.headers.length,
          rowCount: dispatchData.rows.length,
          flightIdIdx: dispatchFlightIdIdx,
          headersPreview: dispatchData.headers.slice(0, 20)
        },
        transactions: {
          name: APP_SHEETS.TRANSACTIONS,
          found: !!transData.found,
          headerCount: transData.headers.length,
          rowCount: transData.rows.length,
          flightIdIdx: transFlightIdIdx,
          headersPreview: transData.headers.slice(0, 20)
        }
      },
      matches: {
        logFound: logMatchIndex >= 0,
        dispatchFound: dispatchMatchIndex >= 0,
        transactionCount: transMatchCount,
        logRowPreview: previewRow(logData.headers, logMatchIndex >= 0 ? logData.rows[logMatchIndex] : null),
        dispatchRowPreview: previewRow(dispatchData.headers, dispatchMatchIndex >= 0 ? dispatchData.rows[dispatchMatchIndex] : null)
      }
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
    var contactNormPayload = {};
    Object.keys(contactPayload).forEach(function(k) {
      contactNormPayload[_toolsNormHeader_(k)] = contactPayload[k];
    });
    var contactPick = function(candidates) {
      var arr = Array.isArray(candidates) ? candidates : [candidates];
      for (var i = 0; i < arr.length; i++) {
        var key = _toolsNormHeader_(arr[i]);
        if (Object.prototype.hasOwnProperty.call(contactNormPayload, key)) {
          var value = contactNormPayload[key];
          if (String(value == null ? '' : value).trim() !== '') return value;
        }
      }
      return '';
    };
    var contactCanonical = {
      CONTATO: contactPick(['CONTATO', 'NOME', 'CONTACT_NAME']),
      CELULAR: contactPick(['CELULAR', 'MOBILE', 'TELEFONE']),
      TELEFONE_2: contactPick(['TELEFONE_2', 'TELEFONE2', 'PHONE2']),
      PERMANENCIA: contactPick(['PERMANENCIA', 'PERMAN_NCIA_', 'PERMAN_NCIA', 'PERMANCIA']),
      ANOTACOES: contactPick(['ANOTACOES', 'ANOTA__ES', 'ANOTAES', 'NOTES'])
    };
    var contactHasData = Object.keys(contactCanonical).some(function(k) {
      return String(contactCanonical[k] == null ? '' : contactCanonical[k]).trim() !== '';
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
              var key = _toolsNormHeader_(String(header || '').trim());
              if (key === 'CONTATO') return contactCanonical.CONTATO || existingContact[idx];
              if (key === 'CELULAR') return contactCanonical.CELULAR || existingContact[idx];
              if (key === 'TELEFONE_2' || key === 'TELEFONE2' || key === 'PHONE2') return contactCanonical.TELEFONE_2 || existingContact[idx];
              if (key === 'PERMANENCIA' || key === 'PERMAN_NCIA_' || key === 'PERMAN_NCIA' || key === 'PERMANCIA') return contactCanonical.PERMANENCIA || existingContact[idx];
              if (key === 'ANOTA__ES' || key === 'ANOTACOES' || key === 'ANOTAES' || key === 'NOTES') return contactCanonical.ANOTACOES || existingContact[idx];
              if (Object.prototype.hasOwnProperty.call(contactNormPayload, key)) {
                var directVal = contactNormPayload[key];
                if (String(directVal == null ? '' : directVal).trim() !== '') return directVal;
              }
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
              var key = _toolsNormHeader_(String(header || '').trim());
              if (key === 'CONTATO') return contactCanonical.CONTATO;
              if (key === 'CELULAR') return contactCanonical.CELULAR;
              if (key === 'TELEFONE_2' || key === 'TELEFONE2' || key === 'PHONE2') return contactCanonical.TELEFONE_2;
              if (key === 'PERMANENCIA' || key === 'PERMAN_NCIA_' || key === 'PERMAN_NCIA' || key === 'PERMANCIA') return contactCanonical.PERMANENCIA;
              if (key === 'ANOTA__ES' || key === 'ANOTACOES' || key === 'ANOTAES' || key === 'NOTES') return contactCanonical.ANOTACOES;
              if (Object.prototype.hasOwnProperty.call(contactNormPayload, key)) return contactNormPayload[key];
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

function getHistoricalFlightBatchReview(batchId) {
  try {
    var target = String(batchId || '').trim();
    if (!target) return { success: false, error: 'batchId required' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var batchSheet = getRequiredSheet_(ss, APP_SHEETS.HIST_FLIGHT_IMPORT_BATCHES, 'getHistoricalFlightBatchReview');
    var stageSheet = getRequiredSheet_(ss, APP_SHEETS.HIST_FLIGHT_STAGING, 'getHistoricalFlightBatchReview');

    var batchData = batchSheet.getDataRange().getValues();
    var stageData = stageSheet.getDataRange().getValues();
    if (!batchData || !batchData.length) return { success: false, error: 'Import batch sheet has no headers' };
    if (!stageData || !stageData.length) return { success: false, error: 'Staging sheet has no headers' };

    var batchHeaders = batchData[0].map(_histFlightNormHeader_);
    var stageHeaders = stageData[0].map(_histFlightNormHeader_);
    var batchIdx = batchHeaders.indexOf('BATCH_ID');
    var rowBatch = null;
    for (var i = 1; i < batchData.length; i++) {
      if (String(batchData[i][batchIdx] || '').trim() === target) { rowBatch = batchData[i]; break; }
    }
    if (!rowBatch) return { success: false, error: 'Batch not found: ' + target };

    var refs = _histFlightBuildReferenceLookup_();
    var existing = _histFlightReadExistingLogKeys_();
    var logSheet = getRequiredSheet_(ss, APP_SHEETS.LOG_FLIGHTS, 'getHistoricalFlightBatchReview/logPreview');
    var logHeaderRow = logSheet.getLastColumn() > 0 ? logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0] : [];

    function rowToObj(headers, row) {
      var obj = {};
      headers.forEach(function(h, idx) { obj[h] = row[idx]; });
      return obj;
    }

    var batch = rowToObj(batchHeaders, rowBatch);
    var rows = stageData.slice(1)
      .filter(function(row) { return String(row[stageHeaders.indexOf('BATCH_ID')] || '').trim() === target; })
      .map(function(row) {
        var obj = rowToObj(stageHeaders, row);
        var item = {
          rowNo: Number(obj.ROW_NO || 0),
          sourceFlightId: String(obj.SOURCE_FLIGHT_ID || ''),
          date: String(obj.FLIGHT_DATE || ''),
          pilot: String(obj.PILOT || ''),
          acft: String(obj.ACFT || ''),
          from: String(obj.FROM_ICAO || ''),
          to: String(obj.TO_ICAO || ''),
          totalTimeH: obj.TOTAL_TIME_H,
          numberLdgs: obj.NUMBER_LDGS,
          touchAndGos: obj.NUMBER_TOUCH_AND_GOS,
          sourceNote: String(obj.SOURCE_NOTE || ''),
          sourceFileId: String(obj.SOURCE_FILE_ID || ''),
          sourceFileUrl: String(obj.SOURCE_FILE_URL || ''),
          sourceFileName: String(obj.SOURCE_FILE_NAME || ''),
          sourceMimeType: String(obj.SOURCE_MIME_TYPE || ''),
          extractionModel: String(obj.EXTRACTION_MODEL || ''),
          extractionConfidence: obj.EXTRACTION_CONFIDENCE,
          matchStatus: String(obj.MATCH_STATUS || ''),
          matchReason: String(obj.MATCH_REASON || ''),
          matchedFlightId: String(obj.MATCHED_FLIGHT_ID || ''),
          reviewDecision: String(obj.REVIEW_DECISION || ''),
          commitFlightId: String(obj.COMMIT_FLIGHT_ID || ''),
          commitRowNo: obj.COMMIT_ROW_NO,
          notes: String(obj.NOTES || '')
        };
        item.validation = _histFlightBuildRowValidation_(item, refs);
        item.logPreview = _histFlightBuildLogPreviewFromStageObject_(obj, target, existing.ids, logHeaderRow);
        return item;
      })
      .sort(function(a, b) { return a.rowNo - b.rowNo; });

    var unresolved = rows.filter(function(r) { return !String(r.reviewDecision || '').trim(); }).length;
    var counts = { NEW: 0, POSSIBLE_DUPLICATE: 0, LIKELY_DUPLICATE: 0, INVALID: 0 };
    rows.forEach(function(r) { counts[r.matchStatus] = (counts[r.matchStatus] || 0) + 1; });

    return {
      success: true,
      batch: {
        batchId: String(batch.BATCH_ID || ''),
        fileName: String(batch.SOURCE_FILENAME || ''),
        fileUrl: String(batch.SOURCE_FILE_URL || ''),
        sourceLabel: String(batch.SOURCE_LABEL || ''),
        documentCount: Number(batch.DOCUMENT_COUNT || 0),
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

function reviewHistoricalFlightStagingRow(batchId, rowNo, decision, notes) {
  try {
    var target = String(batchId || '').trim();
    var rowNumber = Number(rowNo || 0);
    var nextDecision = String(decision || '').trim().toUpperCase();
    var noteText = String(notes || '').trim();
    if (!target || !rowNumber) return { success: false, error: 'batchId and rowNo required' };
    if (['COMMIT', 'DUPLICATE', 'IGNORE'].indexOf(nextDecision) === -1) return { success: false, error: 'Invalid decision' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var stageSheet = getRequiredSheet_(ss, APP_SHEETS.HIST_FLIGHT_STAGING, 'reviewHistoricalFlightStagingRow');
    var data = stageSheet.getDataRange().getValues();
    if (!data || !data.length) return { success: false, error: 'Staging sheet has no headers' };
    var headers = data[0].map(_histFlightNormHeader_);
    var batchIdx = headers.indexOf('BATCH_ID');
    var rowIdx = headers.indexOf('ROW_NO');
    var decisionIdx = headers.indexOf('REVIEW_DECISION');
    var reviewedByIdx = headers.indexOf('REVIEWED_BY');
    var reviewedAtIdx = headers.indexOf('REVIEWED_AT');
    var notesIdx = headers.indexOf('NOTES');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][batchIdx] || '').trim() === target && Number(data[i][rowIdx] || 0) === rowNumber) {
        stageSheet.getRange(i + 1, decisionIdx + 1).setValue(nextDecision);
        if (reviewedByIdx >= 0) stageSheet.getRange(i + 1, reviewedByIdx + 1).setValue(Session.getActiveUser().getEmail() || 'Admin');
        if (reviewedAtIdx >= 0) stageSheet.getRange(i + 1, reviewedAtIdx + 1).setValue(new Date().toISOString());
        if (notesIdx >= 0 && noteText) stageSheet.getRange(i + 1, notesIdx + 1).setValue(noteText);
        return { success: true };
      }
    }
    return { success: false, error: 'Staging row not found' };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function commitHistoricalFlightBatch(batchId) {
  try {
    var target = String(batchId || '').trim();
    if (!target) return { success: false, error: 'batchId required' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var batchSheet = getRequiredSheet_(ss, APP_SHEETS.HIST_FLIGHT_IMPORT_BATCHES, 'commitHistoricalFlightBatch');
    var stageSheet = getRequiredSheet_(ss, APP_SHEETS.HIST_FLIGHT_STAGING, 'commitHistoricalFlightBatch');
    var logSheet = getRequiredSheet_(ss, APP_SHEETS.LOG_FLIGHTS, 'commitHistoricalFlightBatch');

    var stageData = stageSheet.getDataRange().getValues();
    if (!stageData || stageData.length < 2) return { success: false, error: 'No staging rows found' };
    var headers = stageData[0].map(_histFlightNormHeader_);
    var rows = stageData.slice(1).filter(function(row) { return String(row[headers.indexOf('BATCH_ID')] || '').trim() === target; });
    if (!rows.length) return { success: false, error: 'No staging rows found for batch ' + target };

    var unresolved = rows.filter(function(row) { return !String(row[headers.indexOf('REVIEW_DECISION')] || '').trim(); });
    if (unresolved.length) return { success: false, error: 'Batch is blocked until all rows are reviewed' };

    var existing = _histFlightReadExistingLogKeys_();
    var aircraftTypeLookup = _histFlightAircraftTypeLookup_();
    var committed = 0;
    var ignored = 0;
    var duplicates = 0;
    var invalid = 0;

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var decision = String(row[headers.indexOf('REVIEW_DECISION')] || '').trim().toUpperCase();
      if (decision !== 'COMMIT') {
        if (decision === 'DUPLICATE') duplicates++;
        else ignored++;
        continue;
      }

      var flightDate = String(row[headers.indexOf('FLIGHT_DATE')] || '').trim();
      var pilot = String(row[headers.indexOf('PILOT')] || '').trim();
      var acft = String(row[headers.indexOf('ACFT')] || '').trim().toUpperCase();
      var from = String(row[headers.indexOf('FROM_ICAO')] || '').trim().toUpperCase();
      var to = String(row[headers.indexOf('TO_ICAO')] || '').trim().toUpperCase();
      var totalTimeH = row[headers.indexOf('TOTAL_TIME_H')];
      var numberLdgs = row[headers.indexOf('NUMBER_LDGS')];
      var touchAndGos = row[headers.indexOf('NUMBER_TOUCH_AND_GOS')];
      var sourceFlightId = String(row[headers.indexOf('SOURCE_FLIGHT_ID')] || '').trim();
      var sourceNote = String(row[headers.indexOf('SOURCE_NOTE')] || '').trim();
      var sourceFileId = String(row[headers.indexOf('SOURCE_FILE_ID')] || '').trim();
      var sourceFileUrl = String(row[headers.indexOf('SOURCE_FILE_URL')] || '').trim();
      var sourceFileName = String(row[headers.indexOf('SOURCE_FILE_NAME')] || '').trim();
      var sourceMimeType = String(row[headers.indexOf('SOURCE_MIME_TYPE')] || '').trim();
      var extractionModel = String(row[headers.indexOf('EXTRACTION_MODEL')] || '').trim();
      var extractionConfidence = row[headers.indexOf('EXTRACTION_CONFIDENCE')];
      var strictKey = String(row[headers.indexOf('FINGERPRINT_STRICT')] || '').trim();
      var tags = _histFlightReadStructuredTags_(sourceNote);
      var instructorInitials = String(tags.INSTR_INIT || '').trim().toUpperCase();
      var instructorName = _histFlightInstructorNameFromInitials_(instructorInitials);
      var resolvedAircraftType = _histFlightResolveAircraftType_(aircraftTypeLookup, acft, tags.AIRCRAFT_TYPE_RAW);
      var fuelGalStart = Number(String(tags.FUEL_GAL_START || '').replace(',', '.'));
      var fuelGalEnd = Number(String(tags.FUEL_GAL_END || '').replace(',', '.'));
      var fuelStartL = isFinite(fuelGalStart) ? Math.round(fuelGalStart * 3.78541 * 10) / 10 : '';
      var fuelEndL = isFinite(fuelGalEnd) ? Math.round(fuelGalEnd * 3.78541 * 10) / 10 : '';

      if (!flightDate || !pilot || !acft || totalTimeH === '' || !strictKey) {
        invalid++;
        continue;
      }
      if (existing.strict[strictKey] || (sourceFlightId && existing.ids[sourceFlightId])) {
        duplicates++;
        continue;
      }

      var importFlightId = sourceFlightId;
      if (!importFlightId || existing.ids[importFlightId]) {
        var rowNo = Number(row[headers.indexOf('ROW_NO')] || 0);
        importFlightId = target + '-' + ('000' + rowNo).slice(-4);
      }

      var out = new Array(LOG_FLIGHT_COL.NUM_TOUCH_AND_GOS + 1).fill('');
      out[LOG_FLIGHT_COL.FLIGHT_ID] = importFlightId;
      out[LOG_FLIGHT_COL.DATE] = flightDate;
      out[LOG_FLIGHT_COL.PILOT] = pilot;
      out[LOG_FLIGHT_COL.ACFT] = acft;
      out[LOG_FLIGHT_COL.FROM] = from;
      out[LOG_FLIGHT_COL.TO] = to;
      out[LOG_FLIGHT_COL.TOTAL_TIME] = totalTimeH;
      out[LOG_FLIGHT_COL.START_TACH] = tags.TACH_START || '';
      out[LOG_FLIGHT_COL.END_TACH] = tags.TACH_END || '';
      out[LOG_FLIGHT_COL.FUEL_START] = fuelStartL;
      out[LOG_FLIGHT_COL.FUEL_END] = fuelEndL;
      out[LOG_FLIGHT_COL.BLOCK_OUT] = tags.BRAKES_RELEASE || '';
      out[LOG_FLIGHT_COL.OFF] = tags.AIRBORNE || '';
      out[LOG_FLIGHT_COL.ON] = tags.LANDED || '';
      out[LOG_FLIGHT_COL.BLOCK_IN] = tags.BRAKES_APPLIED || '';
      out[LOG_FLIGHT_COL.NUM_LDGS] = numberLdgs;
      out[LOG_FLIGHT_COL.NUM_TOUCH_AND_GOS] = touchAndGos;
      out[LOG_FLIGHT_COL.SQUAWKS] = sourceNote;
      out[LOG_FLIGHT_COL.ACTUAL_LOAD_JSON] = JSON.stringify({
        entryKind: 'HIST_IMPORT',
        batchId: target,
        sourceFlightId: sourceFlightId,
        importedAt: new Date().toISOString(),
        sourceNote: sourceNote,
        sourceFileId: sourceFileId,
        sourceFileUrl: sourceFileUrl,
        sourceFileName: sourceFileName,
        sourceMimeType: sourceMimeType,
        extractionModel: extractionModel,
        extractionConfidence: extractionConfidence,
        trainingCode: String(tags.MISSION_CODE || '').trim(),
        instructorInitials: instructorInitials,
        instructorName: instructorName,
        aircraftType: resolvedAircraftType
      });

      logSheet.appendRow(out);
      var logRowNo = logSheet.getLastRow();
      committed++;

      var absRow = 2;
      for (var s = 1; s < stageData.length; s++) {
        if (String(stageData[s][headers.indexOf('BATCH_ID')] || '').trim() === target && Number(stageData[s][headers.indexOf('ROW_NO')] || 0) === Number(row[headers.indexOf('ROW_NO')] || 0)) {
          absRow = s + 1;
          break;
        }
      }
      var commitFlightIdIdx = headers.indexOf('COMMIT_FLIGHT_ID');
      var commitRowNoIdx = headers.indexOf('COMMIT_ROW_NO');
      if (commitFlightIdIdx >= 0) stageSheet.getRange(absRow, commitFlightIdIdx + 1).setValue(importFlightId);
      if (commitRowNoIdx >= 0) stageSheet.getRange(absRow, commitRowNoIdx + 1).setValue(logRowNo);

      existing.ids[importFlightId] = true;
      existing.strict[strictKey] = importFlightId;
      existing.fuzzy[String(row[headers.indexOf('FINGERPRINT_FUZZY')] || '').trim()] = importFlightId;
    }

    var batchData = batchSheet.getDataRange().getValues();
    if (batchData && batchData.length) {
      var batchHeaders = batchData[0].map(_histFlightNormHeader_);
      var batchIdx = batchHeaders.indexOf('BATCH_ID');
      var statusIdx = batchHeaders.indexOf('STATUS');
      var notesIdx = batchHeaders.indexOf('NOTES');
      for (var bi = 1; bi < batchData.length; bi++) {
        if (String(batchData[bi][batchIdx] || '').trim() === target) {
          if (statusIdx >= 0) batchSheet.getRange(bi + 1, statusIdx + 1).setValue('COMMITTED');
          if (notesIdx >= 0) {
            batchSheet.getRange(bi + 1, notesIdx + 1).setValue(
              'Committed: ' + committed + ', Duplicate: ' + duplicates + ', Ignored: ' + ignored + ', Invalid: ' + invalid
            );
          }
          break;
        }
      }
    }

    var actor = Session.getActiveUser().getEmail() || 'Admin';
    var audit = ss.getSheetByName(APP_SHEETS.AUDIT);
    if (audit) {
      audit.appendRow([
        new Date(),
        actor,
        target,
        'HIST_FLIGHT_IMPORT_COMMIT',
        'UNDER_REVIEW',
        'COMMITTED',
        'Committed: ' + committed + ', Duplicate: ' + duplicates + ', Ignored: ' + ignored + ', Invalid: ' + invalid
      ]);
    }

    return { success: true, committed: committed, duplicates: duplicates, ignored: ignored, invalid: invalid };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function ingestHistoricalFlightDocument(payload) {
  try {
    var ensureRes = _histFlightEnsureImportSheets_();
    if (!ensureRes || !ensureRes.success) return { success: false, error: 'Could not ensure import sheets' };

    var body = (payload && typeof payload === 'object') ? payload : {};
    var base64Data = String(body.base64Data || '').trim();
    var fileName = String(body.fileName || '').trim() || ('historical_' + new Date().getTime() + '.pdf');
    var mimeType = String(body.mimeType || 'application/pdf').trim();
    var importedBy = String(body.importedBy || Session.getActiveUser().getEmail() || 'Admin').trim();
    var sourceLabel = String(body.sourceLabel || 'HISTORICAL_PAPER_CAPTURE').trim();
    var targetBatchId = String(body.batchId || '').trim() || _histFlightBatchId_();

    if (!base64Data) return { success: false, error: 'No file data provided.' };

    var bytes = Utilities.base64Decode(base64Data);
    var safeName = fileName.replace(/[^a-zA-Z0-9._-]+/g, '_');
    var blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', safeName);
    var rootFolder = _histFlightImportsRootFolder_();
    var subfolders = rootFolder.getFoldersByName(targetBatchId);
    var batchFolder = subfolders.hasNext() ? subfolders.next() : rootFolder.createFolder(targetBatchId);
    var file = batchFolder.createFile(blob);

    var extractRes = _histFlightGeminiExtractFromDocument_(base64Data, mimeType, safeName);
    if (!extractRes || !extractRes.success) {
      return { success: false, error: extractRes && extractRes.error ? extractRes.error : 'Gemini extraction failed', batchId: targetBatchId, sourceFileUrl: file.getUrl() };
    }
    var flights = Array.isArray(extractRes.flights) ? extractRes.flights : [];
    if (!flights.length) {
      return { success: false, error: 'No flight rows detected in document.', batchId: targetBatchId, sourceFileUrl: file.getUrl() };
    }

    var fileHash = _histFlightDigest_(base64Data);
    var stageRes = _histFlightStageRecords_({
      batchId: targetBatchId,
      sourceLabel: sourceLabel,
      importedBy: importedBy,
      file: file,
      fileHash: fileHash,
      fileType: (mimeType.indexOf('pdf') >= 0 ? 'PDF' : 'IMAGE'),
      mimeType: mimeType,
      extractionModel: extractRes.model,
      extractionRawJson: extractRes.rawJson,
      rows: flights
    });
    if (!stageRes || !stageRes.success) return stageRes;

    return {
      success: true,
      batchId: stageRes.batchId,
      rowCount: stageRes.rowCount,
      sourceFileUrl: file.getUrl(),
      extractionModel: extractRes.model,
      extractedRowsPreview: flights.slice(0, 5),
      documentSummary: extractRes.documentSummary || ''
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function initHistoricalFlightChunkUpload(meta) {
  try {
    var body = (meta && typeof meta === 'object') ? meta : {};
    var totalChunks = Number(body.totalChunks || 0);
    if (!isFinite(totalChunks) || totalChunks < 1) {
      return { success: false, error: 'Invalid totalChunks' };
    }

    var uploadId = Utilities.getUuid().replace(/-/g, '').slice(0, 24);
    var cache = _histChunkCache_();
    var record = {
      uploadId: uploadId,
      createdAt: new Date().toISOString(),
      totalChunks: totalChunks,
      fileName: String(body.fileName || ''),
      mimeType: String(body.mimeType || 'application/octet-stream'),
      batchId: String(body.batchId || ''),
      sourceLabel: String(body.sourceLabel || 'HISTORICAL_PAPER_CAPTURE'),
      importedBy: String(body.importedBy || Session.getActiveUser().getEmail() || 'Admin')
    };

    cache.put(_histChunkKey_(uploadId, 'meta'), JSON.stringify(record), 1800);
    return { success: true, uploadId: uploadId, totalChunks: totalChunks };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function appendHistoricalFlightChunk(uploadId, chunkIndex, chunkData) {
  try {
    var id = String(uploadId || '').trim();
    if (!id) return { success: false, error: 'uploadId required' };

    var idx = Number(chunkIndex || 0);
    if (!isFinite(idx) || idx < 0) return { success: false, error: 'Invalid chunkIndex' };

    var data = String(chunkData || '');
    if (!data) return { success: false, error: 'Empty chunkData' };
    if (data.length > 90000) return { success: false, error: 'Chunk too large' };

    var cache = _histChunkCache_();
    var metaRaw = cache.get(_histChunkKey_(id, 'meta'));
    if (!metaRaw) return { success: false, error: 'Upload session expired. Start again.' };

    cache.put(_histChunkKey_(id, 'c_' + idx), data, 1800);
    return { success: true, uploadId: id, chunkIndex: idx };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function finalizeHistoricalFlightChunkUpload(uploadId) {
  try {
    var id = String(uploadId || '').trim();
    if (!id) return { success: false, error: 'uploadId required' };

    var cache = _histChunkCache_();
    var metaRaw = cache.get(_histChunkKey_(id, 'meta'));
    if (!metaRaw) return { success: false, error: 'Upload session expired. Start again.' };

    var meta = JSON.parse(metaRaw);
    var total = Number(meta.totalChunks || 0);
    if (!isFinite(total) || total < 1) return { success: false, error: 'Invalid upload metadata' };

    var missing = [];
    var parts = [];
    for (var i = 0; i < total; i++) {
      var piece = cache.get(_histChunkKey_(id, 'c_' + i));
      if (!piece) {
        missing.push(i);
      } else {
        parts.push(piece);
      }
    }
    if (missing.length) {
      return { success: false, error: 'Missing chunks: ' + missing.slice(0, 8).join(',') + (missing.length > 8 ? '...' : '') };
    }

    var base64Data = parts.join('');
    var mimeType = String(meta.mimeType || 'application/octet-stream');
    var safeName = String(meta.fileName || ('historical_' + new Date().getTime() + '.jpg')).replace(/[^a-zA-Z0-9._-]+/g, '_');
    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, mimeType, safeName);
    var targetBatchId = String(meta.batchId || '').trim() || _histFlightBatchId_();
    var rootFolder = _histFlightImportsRootFolder_();
    var subfolders = rootFolder.getFoldersByName(targetBatchId);
    var batchFolder = subfolders.hasNext() ? subfolders.next() : rootFolder.createFolder(targetBatchId);
    var file = batchFolder.createFile(blob);

    var queueSheet = _histFlightQueueSheet_();
    var queueId = Utilities.getUuid().replace(/-/g, '').slice(0, 24);
    queueSheet.appendRow([
      queueId,
      targetBatchId,
      file.getId(),
      file.getName(),
      mimeType,
      String(meta.sourceLabel || 'HISTORICAL_PAPER_CAPTURE'),
      String(meta.importedBy || Session.getActiveUser().getEmail() || 'Admin'),
      'PENDING',
      new Date().toISOString(),
      '',
      '',
      '',
      0,
      ''
    ]);

    _scheduleHistoricalFlightQueueWorker_();

    var keys = [_histChunkKey_(id, 'meta')];
    for (var k = 0; k < total; k++) keys.push(_histChunkKey_(id, 'c_' + k));
    cache.removeAll(keys);

    return {
      success: true,
      queued: true,
      queueId: queueId,
      batchId: targetBatchId,
      sourceFileUrl: file.getUrl()
    };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}

function updateHistoricalFlightStagingRow(batchId, rowNo, patch) {
  try {
    var target = String(batchId || '').trim();
    var targetRowNo = Number(rowNo || 0);
    var changes = (patch && typeof patch === 'object') ? patch : {};
    if (!target || !(targetRowNo > 0)) return { success: false, error: 'batchId and rowNo are required' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var stageSheet = getRequiredSheet_(ss, APP_SHEETS.HIST_FLIGHT_STAGING, 'updateHistoricalFlightStagingRow');
    var data = stageSheet.getDataRange().getValues();
    if (!data || data.length < 2) return { success: false, error: 'No staging data' };

    var headers = data[0].map(_histFlightNormHeader_);
    function hIdx(name) { return headers.indexOf(name); }

    var batchIdx = hIdx('BATCH_ID');
    var rowIdx = hIdx('ROW_NO');
    var absRow = -1;
    var current = null;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][batchIdx] || '').trim() === target && Number(data[i][rowIdx] || 0) === targetRowNo) {
        absRow = i + 1;
        current = data[i];
        break;
      }
    }
    if (absRow < 0 || !current) return { success: false, error: 'Staging row not found' };

    function asText(v) { return String(v == null ? '' : v).trim(); }
    function applyText(key, headerName, transform) {
      if (!Object.prototype.hasOwnProperty.call(changes, key)) return;
      var idx = hIdx(headerName);
      if (idx < 0) return;
      var raw = asText(changes[key]);
      var val = transform ? transform(raw) : raw;
      stageSheet.getRange(absRow, idx + 1).setValue(val);
      current[idx] = val;
    }

    applyText('date', 'FLIGHT_DATE', function(v) { return safeDateYmd_(v); });
    applyText('pilot', 'PILOT', function(v) { return v; });
    applyText('acft', 'ACFT', function(v) { return v.toUpperCase(); });
    applyText('from', 'FROM_ICAO', function(v) { return v.toUpperCase(); });
    applyText('to', 'TO_ICAO', function(v) { return v.toUpperCase(); });
    applyText('sourceFlightId', 'SOURCE_FLIGHT_ID', function(v) { return v; });
    applyText('sourceNote', 'SOURCE_NOTE', function(v) { return v; });

    if (Object.prototype.hasOwnProperty.call(changes, 'totalTimeH')) {
      var tIdx = hIdx('TOTAL_TIME_H');
      if (tIdx >= 0) {
        var pHours = _histFlightParseHours_(changes.totalTimeH);
        stageSheet.getRange(absRow, tIdx + 1).setValue(pHours === '' ? '' : pHours);
        current[tIdx] = pHours === '' ? '' : pHours;
      }
    }
    if (Object.prototype.hasOwnProperty.call(changes, 'numberLdgs')) {
      var lIdx = hIdx('NUMBER_LDGS');
      if (lIdx >= 0) {
        var pLdg = _histFlightParseInt_(changes.numberLdgs);
        stageSheet.getRange(absRow, lIdx + 1).setValue(pLdg === '' ? '' : pLdg);
        current[lIdx] = pLdg === '' ? '' : pLdg;
      }
    }
    if (Object.prototype.hasOwnProperty.call(changes, 'touchAndGos')) {
      var gIdx = hIdx('NUMBER_TOUCH_AND_GOS');
      if (gIdx >= 0) {
        var pTg = _histFlightParseInt_(changes.touchAndGos);
        stageSheet.getRange(absRow, gIdx + 1).setValue(pTg === '' ? '' : pTg);
        current[gIdx] = pTg === '' ? '' : pTg;
      }
    }

    var obj = {
      date: String(current[hIdx('FLIGHT_DATE')] || ''),
      pilot: String(current[hIdx('PILOT')] || '').trim(),
      acft: String(current[hIdx('ACFT')] || '').trim().toUpperCase(),
      from: String(current[hIdx('FROM_ICAO')] || '').trim().toUpperCase(),
      to: String(current[hIdx('TO_ICAO')] || '').trim().toUpperCase(),
      totalTimeH: current[hIdx('TOTAL_TIME_H')]
    };

    var strictKey = _histFlightMakeStrictKey_(obj);
    var fuzzyKey = _histFlightMakeFuzzyKey_(obj);
    var sourceFlightId = String(current[hIdx('SOURCE_FLIGHT_ID')] || '').trim();
    var existing = _histFlightReadExistingLogKeys_();

    var matchStatus = 'NEW';
    var matchReason = '';
    var matchedFlightId = '';
    if (!obj.date || !obj.pilot || !obj.acft || obj.totalTimeH === '') {
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

    var strictIdx = hIdx('FINGERPRINT_STRICT');
    var fuzzyIdx = hIdx('FINGERPRINT_FUZZY');
    var msIdx = hIdx('MATCH_STATUS');
    var mrIdx = hIdx('MATCH_REASON');
    var mfIdx = hIdx('MATCHED_FLIGHT_ID');
    if (strictIdx >= 0) stageSheet.getRange(absRow, strictIdx + 1).setValue(strictKey);
    if (fuzzyIdx >= 0) stageSheet.getRange(absRow, fuzzyIdx + 1).setValue(fuzzyKey);
    if (msIdx >= 0) stageSheet.getRange(absRow, msIdx + 1).setValue(matchStatus);
    if (mrIdx >= 0) stageSheet.getRange(absRow, mrIdx + 1).setValue(matchReason);
    if (mfIdx >= 0) stageSheet.getRange(absRow, mfIdx + 1).setValue(matchedFlightId);

    return { success: true, matchStatus: matchStatus, matchReason: matchReason, matchedFlightId: matchedFlightId };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  }
}
