function syncAnacData() {
  const summary = runAirportSyncPhase1_({
    sourceTag: 'ANAC_PUBLIC',
    tempSheetName: 'TEMP_ANAC_PUBLIC',
    tempHeaderRow: 2,
    tempDataStartRow: 3,
    infrastructureType: 'Land'
  });

  SpreadsheetApp.getUi().alert(buildAirportSyncSummaryMessage_(summary));
}

function buildAirportSyncMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Airport Sync')
    .addItem('Set Public CSV URL', 'setAnacPublicCsvUrl')
    .addItem('Set Private CSV URL', 'setAnacPrivateCsvUrl')
    .addSeparator()
    .addItem('Pull Public -> TEMP', 'pullAnacPublicToTemp')
    .addItem('Pull Private -> TEMP', 'pullAnacPrivateToTemp')
    .addItem('Pull Both -> TEMP', 'pullAnacBothToTemp')
    .addItem('Pull Both + Sync Both', 'pullAndSyncAnacBothData')
    .addSeparator()
    .addItem('Run Public Sync', 'syncAnacData')
    .addItem('Run Private Sync', 'syncAnacPrivateData')
    .addItem('Run Both + Report', 'syncAnacBothData')
    .addSeparator()
    .addItem('Debug ICAO Mapping', 'debugAirportSyncForIcao')
    .addItem('Open Sync Report', 'openAirportSyncReport')
    .addToUi();
}

function debugAirportSyncForIcao() {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt('Debug ICAO Mapping', 'Enter ICAO code to inspect (e.g., SBGR):', ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return;

  const icao = normalizeIcao_(prompt.getResponseText());
  if (!icao) {
    ui.alert('Please provide a valid ICAO code.');
    return;
  }

  const lines = [];
  lines.push('ICAO Debug: ' + icao);
  lines.push('Timestamp: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
  lines.push('');
  lines.push(debugTempSheetForIcao_('TEMP_ANAC_PUBLIC', icao));
  lines.push('');
  lines.push(debugTempSheetForIcao_('TEMP_ANAC_PRIV', icao));
  lines.push('');
  lines.push(debugDbForIcao_(icao));

  const report = lines.join('\n');
  writeAirportDebugReport_(report);
  ui.alert('ICAO debug captured in LOG_AirportDebug.\n\n' + report.slice(0, 3500));
}

function setAnacPublicCsvUrl() {
  setAnacCsvUrlByPrompt_('ANAC_PUBLIC_CSV_URL', 'Set ANAC Public CSV URL', 'Paste the direct CSV URL for Aeródromos de Uso Público:');
}

function setAnacPrivateCsvUrl() {
  setAnacCsvUrlByPrompt_('ANAC_PRIVATE_CSV_URL', 'Set ANAC Private CSV URL', 'Paste the direct CSV URL for Aeródromos de Uso Privativo:');
}

function pullAnacPublicToTemp() {
  const result = pullAnacCsvToTemp_({
    sourceTag: 'ANAC_PUBLIC',
    propKey: 'ANAC_PUBLIC_CSV_URL',
    tempSheetName: 'TEMP_ANAC_PUBLIC',
    tempHeaderRow: 2,
    tempDataStartRow: 3
  });
  SpreadsheetApp.getUi().alert(buildAnacPullSummaryMessage_(result));
}

function pullAnacPrivateToTemp() {
  const result = pullAnacCsvToTemp_({
    sourceTag: 'ANAC_PRIVATE',
    propKey: 'ANAC_PRIVATE_CSV_URL',
    tempSheetName: 'TEMP_ANAC_PRIV',
    tempHeaderRow: 2,
    tempDataStartRow: 3,
    delimiter: ';',
    charset: 'windows-1252',
    expandMultiRunway: true
  });
  SpreadsheetApp.getUi().alert(buildAnacPullSummaryMessage_(result));
}

function pullAnacBothToTemp() {
  const pub = pullAnacCsvToTemp_({
    sourceTag: 'ANAC_PUBLIC',
    propKey: 'ANAC_PUBLIC_CSV_URL',
    tempSheetName: 'TEMP_ANAC_PUBLIC',
    tempHeaderRow: 2,
    tempDataStartRow: 3
  });
  const priv = pullAnacCsvToTemp_({
    sourceTag: 'ANAC_PRIVATE',
    propKey: 'ANAC_PRIVATE_CSV_URL',
    tempSheetName: 'TEMP_ANAC_PRIV',
    tempHeaderRow: 2,
    tempDataStartRow: 3,
    delimiter: ';',
    charset: 'windows-1252'
  });

  SpreadsheetApp.getUi().alert([
    'ANAC pull complete (Both)',
    '',
    '[PUBLIC] rows=' + pub.rowsWritten + ', headers=' + pub.headerCount,
    '[PRIVATE] rows=' + priv.rowsWritten + ', headers=' + priv.headerCount
  ].join('\n'));
}

function pullAndSyncAnacBothData() {
  pullAnacBothToTemp();
  syncAnacBothData();
}

function syncAnacBothData() {
  const publicSummary = runAirportSyncPhase1_({
    sourceTag: 'ANAC_PUBLIC',
    tempSheetName: 'TEMP_ANAC_PUBLIC',
    tempHeaderRow: 2,
    tempDataStartRow: 3,
    infrastructureType: 'Land'
  });

  const privateSummary = runAirportSyncPhase1_({
    sourceTag: 'ANAC_PRIVATE',
    tempSheetName: 'TEMP_ANAC_PRIV',
    tempHeaderRow: 2,
    tempDataStartRow: 3,
    infrastructureType: 'Private Land'
  });

  const msg = [
    'Airport sync complete (Both)',
    '',
    '[PUBLIC] inserted=' + publicSummary.inserted + ', updated=' + publicSummary.updated + ', locked=' + publicSummary.lockedSkipped + ', parseErrors=' + publicSummary.parseErrors,
    '[PRIVATE] inserted=' + privateSummary.inserted + ', updated=' + privateSummary.updated + ', locked=' + privateSummary.lockedSkipped + ', parseErrors=' + privateSummary.parseErrors,
    '',
    'Report tab: LOG_AirportSync'
  ].join('\n');

  SpreadsheetApp.getUi().alert(msg);
}

function openAirportSyncReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('LOG_AirportSync');
  if (!sh) {
    SpreadsheetApp.getUi().alert("Sheet 'LOG_AirportSync' does not exist yet. Run a sync first.");
    return;
  }
  ss.setActiveSheet(sh);
}

function debugTempSheetForIcao_(sheetName, icao) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return '[' + sheetName + '] Missing sheet';

  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  if (lastCol < 1 || lastRow < 2) return '[' + sheetName + '] Empty sheet';

  const detected = detectTempHeaderRowForSync_(sh, lastCol, 30);
  const headerRow = detected ? detected.headerRow : 2;
  const dataStart = headerRow + 1;
  const rawHeaders = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const headers = buildEffectiveTempHeadersForSync_(sh, headerRow, lastCol, rawHeaders);
  const idx = buildAirportColumnIndex_([], headers).temp;

  const runwayCols = collectRunwayDesignColumnIndices_(headers);
  const runwayColInfo = runwayCols.map(function(c) { return (c + 1) + ':' + String(headers[c]); }).join(' | ');
  const primaryRunwayHeader = idx.runway >= 0 ? String(headers[idx.runway]) : '(not found)';
  const normHeaders = headers.map(function(h) { return normalizeHeader_(h); });
  const normRunwayHeaders = normHeaders.filter(function(h, i) { return i === idx.runway || (i >= 10 && i <= 25); });

  if (lastRow < dataStart) {
    return [
      '[' + sheetName + ']',
      'HeaderRow=' + headerRow + ', DataStart=' + dataStart,
      'ICAO col=' + (idx.icao + 1) + ', primary runway col=' + (idx.runway + 1) + ' (' + primaryRunwayHeader + ')',
      'Runway design columns=' + runwayCols.length,
      'No data rows'
    ].join('\n');
  }

  const rows = sh.getRange(dataStart, 1, lastRow - dataStart + 1, lastCol).getValues();
  const matched = [];
  const uniqueRunways = {};
  const samples = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowIcao = normalizeIcao_(safeCell_(row, idx.icao));
    if (rowIcao !== icao) continue;

    const fromAllCols = tokenizeRunwaysFromColumns_(row, runwayCols);
    const fromPrimary = splitRunwayDesignators_(safeCell_(row, idx.runway));

    fromAllCols.forEach(function(r) { uniqueRunways[r] = true; });
    matched.push(i + dataStart);

    if (samples.length < 8) {
      samples.push(
        'R' + (i + dataStart) +
        ' primaryRaw=[' + String(safeCell_(row, idx.runway)) + ']' +
        ' primaryParsed=[' + fromPrimary.join(',') + ']' +
        ' allParsed=[' + fromAllCols.join(',') + ']'
      );
    }
  }

  const uniqueList = Object.keys(uniqueRunways).sort();
  return [
    '[' + sheetName + ']',
    'HeaderRow=' + headerRow + ', DataStart=' + dataStart,
    'ICAO col=' + (idx.icao + 1) + ', primary runway col=' + (idx.runway + 1) + ' (' + primaryRunwayHeader + ')',
    'Runway design columns=' + runwayCols.length + (runwayColInfo ? ' -> ' + runwayColInfo : ''),
    'Sample normalized headers (cols 11-26): ' + normRunwayHeaders.join(' | '),
    'Matched rows for ' + icao + '=' + matched.length,
    'Unique parsed runways=' + (uniqueList.length ? uniqueList.join(',') : '(none)'),
    'Samples:',
    (samples.length ? samples.join('\n') : '(none)')
  ].join('\n');
}

function debugDbForIcao_(icao) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('DB_Airports');
  if (!sh) return '[DB_Airports] Missing sheet';

  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  if (lastCol < 1 || lastRow < 2) return '[DB_Airports] Empty sheet';

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = buildAirportColumnIndex_(headers, []).db;
  const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const runwaySet = {};
  const samples = [];
  let matched = 0;

  rows.forEach(function(row, i) {
    if (normalizeIcao_(safeCell_(row, idx.icao)) !== icao) return;
    matched += 1;
    const rwy = normalizeRunwayToken_(safeCell_(row, idx.runway));
    if (rwy) runwaySet[rwy] = true;
    if (samples.length < 10) {
      samples.push(
        'R' + (i + 2) +
        ' rwy=[' + String(safeCell_(row, idx.runway)) + ']' +
        ' length=[' + String(safeCell_(row, idx.length)) + ']' +
        ' width=[' + String(safeCell_(row, idx.width)) + ']'
      );
    }
  });

  return [
    '[DB_Airports]',
    'Matched rows for ' + icao + '=' + matched,
    'Unique runway values=' + Object.keys(runwaySet).sort().join(','),
    'Samples:',
    (samples.length ? samples.join('\n') : '(none)')
  ].join('\n');
}

function collectRunwayDesignColumnIndices_(headers) {
  const out = [];
  const norm = (headers || []).map(function(h) { return normalizeHeader_(h); });
  for (let i = 0; i < norm.length; i++) {
    const h = norm[i];
    if (isRunwayDesignHeaderToken_(h)) {
      out.push(i);
    }
  }
  return out;
}

function isRunwayDesignHeaderToken_(token) {
  const h = String(token == null ? '' : token).toUpperCase();
  if (!h) return false;
  if (h === 'CABECEIRA' || h === 'RWY' || h === 'RUNWAY') return true;
  if (h.indexOf('DESIGNA') >= 0) return true;
  if (h.indexOf('PISTA') >= 0 && h.indexOf('COMPRIMENTO') < 0 && h.indexOf('LARGURA') < 0 && h.indexOf('SUPERFICIE') < 0 && h.indexOf('RESISTENCIA') < 0) return true;
  return false;
}

function tokenizeRunwaysFromColumns_(row, indices) {
  const unique = {};
  (indices || []).forEach(function(i) {
    splitRunwayDesignators_(safeCell_(row, i)).forEach(function(token) {
      unique[token] = true;
    });
  });
  return Object.keys(unique).sort();
}

function writeAirportDebugReport_(reportText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = 'LOG_AirportDebug';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, 2).setValues([['RUN_AT', 'REPORT']]);
  }

  sh.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    reportText
  ]);
}

function syncAnacPrivateData() {
  const summary = runAirportSyncPhase1_({
    sourceTag: 'ANAC_PRIVATE',
    tempSheetName: 'TEMP_ANAC_PRIV',
    tempHeaderRow: 2,
    tempDataStartRow: 3,
    infrastructureType: 'Private Land'
  });

  SpreadsheetApp.getUi().alert(buildAirportSyncSummaryMessage_(summary));
}

function runAirportSyncPhase1_(cfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('DB_Airports');
  const tempSheet = ss.getSheetByName(cfg.tempSheetName);

  if (!dbSheet) throw new Error("Sheet 'DB_Airports' not found");
  if (!tempSheet) throw new Error("Sheet '" + cfg.tempSheetName + "' not found");

  ensureDbAirportSyncColumns_(dbSheet, ['SOURCE', 'MANUAL_LOCK', 'LAST_SYNC_AT', 'WIDTH_OFFICIAL']);

  const dbLastCol = dbSheet.getLastColumn();
  if (dbLastCol < 1) {
    throw new Error("Sheet 'DB_Airports' has no columns/header row");
  }

  const tempLastCol = tempSheet.getLastColumn();
  if (tempLastCol < 1) {
    throw new Error("Sheet '" + cfg.tempSheetName + "' has no columns/header row");
  }

  const detectedTempHeader = detectTempHeaderRowForSync_(tempSheet, tempLastCol, cfg.tempHeaderScanRows || 20);
  const resolvedTempHeaderRow = detectedTempHeader ? detectedTempHeader.headerRow : cfg.tempHeaderRow;
  const resolvedTempDataStartRow = detectedTempHeader ? (resolvedTempHeaderRow + 1) : cfg.tempDataStartRow;

  if (tempSheet.getLastRow() < resolvedTempHeaderRow) {
    throw new Error("Sheet '" + cfg.tempSheetName + "' is missing the configured/detected header row " + resolvedTempHeaderRow);
  }

  const dbHeaders = dbSheet.getRange(1, 1, 1, dbLastCol).getValues()[0];
  const dbRows = dbSheet.getLastRow() > 1
    ? dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, dbLastCol).getValues()
    : [];

  const rawTempHeaders = detectedTempHeader
    ? detectedTempHeader.headers
    : tempSheet.getRange(resolvedTempHeaderRow, 1, 1, tempLastCol).getValues()[0];
  const tempHeaders = buildEffectiveTempHeadersForSync_(tempSheet, resolvedTempHeaderRow, tempLastCol, rawTempHeaders);
  const tempLastRow = tempSheet.getLastRow();
  const tempRows = tempLastRow >= resolvedTempDataStartRow
    ? tempSheet.getRange(resolvedTempDataStartRow, 1, tempLastRow - resolvedTempDataStartRow + 1, tempLastCol).getValues()
    : [];

  let idx = buildAirportColumnIndex_(dbHeaders, tempHeaders);
  const protectedCols = buildProtectedColumnIndexSet_(dbHeaders, [
    'SURFACE_ACTUAL',
    'MTOW_LIMIT_206_520',
    'MTOW_LIMIT_206_550',
    'SLOPE_PERCENT',
    'ONE_WAY',
    'PILOT_NOTES',
    'AIRSTRIP_PHOTO',
    'FUEL_AVAILABLE',
    'KNOWN_FEATURES'
  ]);

  if (idx.temp.icao < 0 || idx.temp.runway < 0) {
    const foundHeaders = tempHeaders
      .map(function(h) { return String(h == null ? '' : h).trim(); })
      .filter(function(h) { return h !== ''; })
      .slice(0, 25)
      .join(' | ');
    throw new Error(
      'Missing required temp columns: ICAO/OACI and runway designator. ' +
      'Expected aliases include [CÓDIGO OACI, CODIGO_OACI, ICAO] and [DESIGNAÇÃO, DESIGNACAO, DESIGNACAO_DA_PISTA, PISTA, RUNWAY]. ' +
      'Found headers: ' + foundHeaders
    );
  }
  if (idx.db.icao < 0 || idx.db.runway < 0) {
    throw new Error('Missing required DB columns: ICAO/OACI and runway identifier');
  }

  const expandedResult = expandMultiRunwayRows_(tempHeaders, tempRows);
  let syncHeaders = tempHeaders;
  let syncRows = tempRows;
  if (expandedResult && expandedResult.rows && expandedResult.rows.length > 0) {
    syncHeaders = expandedResult.headers;
    syncRows = expandedResult.rows;
    idx = buildAirportColumnIndex_(dbHeaders, syncHeaders);
  }

  const stats = {
    source: cfg.sourceTag,
    tempSheet: cfg.tempSheetName,
    tempHeaderRowUsed: resolvedTempHeaderRow,
    tempDataStartRowUsed: resolvedTempDataStartRow,
    rowsRead: tempRows.length,
    runwayRowsParsed: 0,
    inserted: 0,
    updated: 0,
    unchanged: 0,
    lockedSkipped: 0,
    parseErrors: 0,
    warnings: []
  };

  const resultRows = dbRows.map(r => r.slice());
  const dbKeyToRowIndex = new Map();

  if (detectedTempHeader && resolvedTempHeaderRow !== cfg.tempHeaderRow) {
    stats.warnings.push('Auto-detected temp header row ' + resolvedTempHeaderRow + ' (configured ' + cfg.tempHeaderRow + ')');
  }

  for (let i = 0; i < resultRows.length; i++) {
    const key = rowKey_(resultRows[i], idx.db);
    if (!key) continue;
    if (!dbKeyToRowIndex.has(key)) {
      dbKeyToRowIndex.set(key, i);
    }
  }

  const nowIso = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");

  for (let i = 0; i < syncRows.length; i++) {
    const srcRow = syncRows[i];
    const icao = normalizeIcao_(safeCell_(srcRow, idx.temp.icao));
    const rawRunway = safeCell_(srcRow, idx.temp.runway);

    if (!icao || !rawRunway) continue;

    const runwayList = splitRunwayDesignators_(rawRunway);
    if (!runwayList.length) {
      stats.parseErrors += 1;
      stats.warnings.push('No runway token parsed for ICAO ' + icao + ' from value: ' + String(rawRunway));
      continue;
    }

    const mapped = mapTempAirportData_(srcRow, idx.temp);

    runwayList.forEach(function(rwy) {
      stats.runwayRowsParsed += 1;
      const key = icao + '-' + rwy;
      const existingIndex = dbKeyToRowIndex.has(key) ? dbKeyToRowIndex.get(key) : -1;

      if (existingIndex >= 0) {
        const row = resultRows[existingIndex];
        const existingSource = String(safeCell_(row, idx.db.source)).trim().toUpperCase();
        const locked = isManualLocked_(safeCell_(row, idx.db.manualLock));
        const manualProtected = locked || existingSource.indexOf('MANUAL') === 0;

        if (manualProtected) {
          stats.lockedSkipped += 1;
          return;
        }

        const changed = applyAirportRowUpdate_(row, idx.db, {
          icao: icao,
          runway: rwy,
          source: cfg.sourceTag,
          infrastructureType: cfg.infrastructureType,
          timestamp: nowIso,
          mapped: mapped,
          protectedCols: protectedCols
        });

        if (changed) stats.updated += 1;
        else stats.unchanged += 1;
      } else {
        const row = new Array(dbHeaders.length).fill('');
        applyAirportRowUpdate_(row, idx.db, {
          icao: icao,
          runway: rwy,
          source: cfg.sourceTag,
          infrastructureType: cfg.infrastructureType,
          timestamp: nowIso,
          mapped: mapped,
          protectedCols: protectedCols
        });
        resultRows.push(row);
        dbKeyToRowIndex.set(key, resultRows.length - 1);
        stats.inserted += 1;
      }
    });
  }

  if (resultRows.length) {
    if (idx.db.runway >= 0) {
      dbSheet.getRange(2, idx.db.runway + 1, resultRows.length, 1).setNumberFormat('@');
    }
    dbSheet.getRange(2, 1, resultRows.length, dbHeaders.length).setValues(resultRows);
    if (dbSheet.getLastRow() > resultRows.length + 1) {
      dbSheet.getRange(resultRows.length + 2, 1, dbSheet.getLastRow() - (resultRows.length + 1), dbHeaders.length).clearContent();
    }
  }

  writeAirportSyncReport_(ss, stats);
  return stats;
}

function detectTempHeaderRowForSync_(sheet, lastCol, maxScanRows) {
  const lastRow = sheet.getLastRow();
  const scanRows = Math.min(lastRow, maxScanRows || 20);
  if (scanRows < 1 || lastCol < 1) return null;

  const matrix = sheet.getRange(1, 1, scanRows, lastCol).getValues();
  for (let i = 0; i < matrix.length; i++) {
    const candidateHeaders = matrix[i];
    const idx = buildAirportColumnIndex_([], candidateHeaders);
    if (idx.temp.icao >= 0 && idx.temp.runway >= 0) {
      return { headerRow: i + 1, headers: candidateHeaders };
    }
  }
  return null;
}

function buildEffectiveTempHeadersForSync_(sheet, headerRow, lastCol, headers) {
  const baseHeaders = stringifyRow_(headers || []);
  if (!sheet || headerRow <= 1 || lastCol < 1 || !baseHeaders.length) return baseHeaders;

  const top = sheet.getRange(headerRow - 1, 1, 1, lastCol).getValues()[0];
  const topNorm = top.map(function(c) { return normalizeHeader_(c); });

  const runwayGroupByCol = new Array(lastCol).fill('');
  let activeRunwayGroup = '';

  for (var i = 0; i < topNorm.length; i++) {
    var t = topNorm[i];
    if (t) {
      var m = t.match(/PISTA[_ ]?(\d+)/);
      if (m) {
        activeRunwayGroup = m[1];
      } else {
        activeRunwayGroup = '';
      }
    }
    runwayGroupByCol[i] = activeRunwayGroup;
  }

  const out = baseHeaders.slice();
  for (var j = 0; j < out.length; j++) {
    var groupNum = runwayGroupByCol[j];
    if (!groupNum) continue;

    var hNorm = normalizeHeader_(out[j]);
    var isRunwayField = hNorm.indexOf('DESIGNACAO') >= 0 ||
      hNorm.indexOf('COMPRIMENTO') >= 0 ||
      hNorm.indexOf('LARGURA') >= 0 ||
      hNorm.indexOf('RESISTENCIA') >= 0 ||
      hNorm.indexOf('SUPERFICIE') >= 0;
    if (!isRunwayField) continue;

    out[j] = String(out[j] == null ? '' : out[j])
      .replace(/[\s_]?\d+$/, '')
      .trim() + ' ' + groupNum;
  }

  return out;
}

function pullAnacCsvToTemp_(cfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = getRequiredScriptProperty_(cfg.propKey);
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('ANAC pull failed (' + cfg.sourceTag + ') HTTP ' + code + '. URL: ' + url);
  }

  const charset = cfg.charset || 'UTF-8';
  const text = String(response.getContentText(charset) || '').replace(/^\uFEFF/, '');
  const delimiter = cfg.delimiter || detectCsvDelimiter_(text);
  let rows = Utilities.parseCsv(text, delimiter);
  rows = normalizeCsvRows_(rows);

  if (!rows.length || !rows[0].length) {
    throw new Error('CSV parsed empty for ' + cfg.sourceTag + '. Check URL/content.');
  }

  // Skip leading metadata lines (e.g. "Atualizado em: 2026-03-25") to find real header row
  const headerRowIdx = findCsvHeaderRowIndex_(rows);
  let headers = rows[headerRowIdx];
  let dataRows = rows.slice(headerRowIdx + 1).filter(function(r) {
    return r.some(function(c) { return String(c || '').trim() !== ''; });
  });

  // Expansion now happens during sync, not pull, for consistency
  headers = stringifyRow_(headers);
  dataRows = dataRows.map(function(r) { return stringifyRow_(r); });

  let tempSheet = ss.getSheetByName(cfg.tempSheetName);
  if (!tempSheet) tempSheet = ss.insertSheet(cfg.tempSheetName);

  tempSheet.clear();
  if (tempSheet.getMaxRows() < cfg.tempDataStartRow) {
    tempSheet.insertRowsAfter(tempSheet.getMaxRows(), cfg.tempDataStartRow - tempSheet.getMaxRows());
  }
  if (tempSheet.getMaxColumns() < headers.length) {
    tempSheet.insertColumnsAfter(tempSheet.getMaxColumns(), headers.length - tempSheet.getMaxColumns());
  }

  const pulledAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  tempSheet.getRange(1, 1).setNumberFormat('@').setValue('Source=' + cfg.sourceTag + ' | PulledAt=' + pulledAt + ' | URL=' + url);
  tempSheet.getRange(cfg.tempHeaderRow, 1, 1, headers.length).setNumberFormat('@').setValues([headers]);

  if (dataRows.length) {
    tempSheet.getRange(cfg.tempDataStartRow, 1, dataRows.length, headers.length).setNumberFormat('@').setValues(dataRows);
  }

  return {
    source: cfg.sourceTag,
    tempSheet: cfg.tempSheetName,
    url: url,
    headerCount: headers.length,
    rowsWritten: dataRows.length
  };
}

function findCsvHeaderRowIndex_(rows) {
  // Prefer a row that explicitly looks like ANAC aerodrome headers
  for (var i = 0; i < Math.min(rows.length, 20); i++) {
    var normCells = rows[i].map(function(c) { return normalizeHeader_(c); });
    var hasIcao = normCells.indexOf('CODIGO_OACI') >= 0 || normCells.indexOf('OACI') >= 0 || normCells.indexOf('ICAO') >= 0;
    var hasRunway = normCells.some(function(c) {
      return c.indexOf('DESIGNACAO') >= 0 || c.indexOf('PISTA') >= 0 || c === 'RUNWAY' || c === 'RWY';
    });
    if (hasIcao && hasRunway) return i;
  }

  // Fallback: skip rows that look like single-cell metadata (e.g. "Atualizado em: 2026-03-25")
  for (var j = 0; j < Math.min(rows.length, 20); j++) {
    var nonEmpty = rows[j].filter(function(c) { return String(c || '').trim() !== ''; });
    if (nonEmpty.length >= 3) return j;
  }
  return 0;
}

function expandMultiRunwayRows_(headers, dataRows) {
  var norm = headers.map(function(h) { return normalizeHeader_(h); });
  var repeatedFallback = function() {
    return expandRepeatedRunwayHeaderRows_(headers, dataRows, norm);
  };

  // Find column groups with numeric suffix: e.g. DESIGNACAO_1, COMPRIMENTO_1 -> slot 1
  var slotColIndices = {};
  norm.forEach(function(h, i) {
    var m = h.match(/^(.+?)(?:[_ ]?)(\d+)$/);
    if (!m) return;
    var num = parseInt(m[2]);
    if (!slotColIndices[num]) slotColIndices[num] = [];
    slotColIndices[num].push(i);
  });

  var slotNums = Object.keys(slotColIndices).map(Number).sort(function(a,b){return a-b;});
  if (slotNums.length < 2) {
    return repeatedFallback();
  }

  // Collect all slotted indices
  var slottedSet = {};
  slotNums.forEach(function(n) {
    slotColIndices[n].forEach(function(i) { slottedSet[i] = true; });
  });

  // Base (non-slotted) column indices
  var baseIndices = [];
  norm.forEach(function(_, i) {
    if (!slottedSet[i]) baseIndices.push(i);
  });

  // Canonical (unsuffixed) headers from slot 1
  var slot1Indices = slotColIndices[slotNums[0]];
  var canonicalSlotHeaders = slot1Indices.map(function(i) {
    return headers[i].replace(/[\s_]?\d+$/, '');
  });
  var canonicalBaseKeys = slot1Indices.map(function(i) {
    return norm[i].replace(/[_ ]?\d+$/, '');
  });

  var slotFieldIndex = {};
  slotNums.forEach(function(n) {
    slotFieldIndex[n] = {};
    slotColIndices[n].forEach(function(i) {
      var base = norm[i].replace(/[_ ]?\d+$/, '');
      if (!slotFieldIndex[n][base]) slotFieldIndex[n][base] = i;
    });
  });

  var outHeaders = baseIndices.map(function(i) { return headers[i]; }).concat(canonicalSlotHeaders);

  var outRows = [];
  var slotsWithDesignColumn = 0;
  dataRows.forEach(function(row) {
    slotNums.forEach(function(n) {
      var slotIndices = slotColIndices[n];

      // Find designator column in this slot
      var designIdx = -1;
      Object.keys(slotFieldIndex[n]).forEach(function(base) {
        if (designIdx >= 0) return;
        if (base.indexOf('DESIGNA') >= 0 || base.indexOf('PISTA') >= 0 || base.indexOf('RUNWAY') >= 0 || base === 'RWY') {
          designIdx = slotFieldIndex[n][base];
        }
      });
      if (designIdx < 0) {
        slotIndices.forEach(function(i) {
          if (norm[i].indexOf('DESIGNA') >= 0 || norm[i].indexOf('PISTA') >= 0 || norm[i].indexOf('RUNWAY') >= 0 || norm[i] === 'RWY') designIdx = i;
        });
      }
      if (designIdx >= 0) slotsWithDesignColumn += 1;

      // Skip empty runway slots
      var desigVal = String(designIdx >= 0 ? (row[designIdx] || '') : '').trim();
      if (!desigVal) return;

      var newRow = baseIndices.map(function(i) { return row[i] != null ? row[i] : ''; });
      canonicalBaseKeys.forEach(function(base) {
        var srcIdx = slotFieldIndex[n][base];
        newRow.push(srcIdx != null && row[srcIdx] != null ? row[srcIdx] : '');
      });
      outRows.push(newRow);
    });
  });

  if (slotsWithDesignColumn === 0 || !outRows.length) {
    return repeatedFallback();
  }

  return { headers: outHeaders, rows: outRows };
}

function expandRepeatedRunwayHeaderRows_(headers, dataRows, normHeaders) {
  var norm = Array.isArray(normHeaders) ? normHeaders : headers.map(function(h) { return normalizeHeader_(h); });

  var runwayDesignIdxs = [];
  norm.forEach(function(h, i) {
    if (isRunwayDesignHeaderToken_(h)) {
      runwayDesignIdxs.push(i);
    }
  });

  if (runwayDesignIdxs.length < 2) return null;

  var runwayGroups = runwayDesignIdxs.map(function(designIdx, pos) {
    var nextDesignIdx = pos + 1 < runwayDesignIdxs.length ? runwayDesignIdxs[pos + 1] : Math.min(norm.length, designIdx + 10);
    var group = { design: designIdx, length: -1, width: -1, resistance: -1, surface: -1 };

    for (var i = designIdx + 1; i < nextDesignIdx; i++) {
      if (group.length < 0 && norm[i].indexOf('COMPRIMENTO') >= 0) group.length = i;
      if (group.width < 0 && norm[i].indexOf('LARGURA') >= 0) group.width = i;
      if (group.resistance < 0 && norm[i].indexOf('RESISTENCIA') >= 0) group.resistance = i;
      if (group.surface < 0 && norm[i].indexOf('SUPERFICIE') >= 0) group.surface = i;
    }
    return group;
  });

  var groupedSet = {};
  runwayGroups.forEach(function(g) {
    [g.design, g.length, g.width, g.resistance, g.surface].forEach(function(idx) {
      if (idx >= 0) groupedSet[idx] = true;
    });
  });

  var baseIndices = [];
  norm.forEach(function(_, i) {
    if (!groupedSet[i]) baseIndices.push(i);
  });

  var outHeaders = baseIndices.map(function(i) { return headers[i]; }).concat([
    'DESIGNAÇÃO',
    'COMPRIMENTO',
    'LARGURA',
    'RESISTÊNCIA',
    'SUPERFÍCIE'
  ]);

  var outRows = [];
  dataRows.forEach(function(row) {
    runwayGroups.forEach(function(g) {
      var desigVal = String(g.design >= 0 ? (row[g.design] || '') : '').trim();
      if (!desigVal || desigVal === '-') return;

      var newRow = baseIndices.map(function(i) { return row[i] != null ? row[i] : ''; });
      newRow.push(g.design >= 0 && row[g.design] != null ? row[g.design] : '');
      newRow.push(g.length >= 0 && row[g.length] != null ? row[g.length] : '');
      newRow.push(g.width >= 0 && row[g.width] != null ? row[g.width] : '');
      newRow.push(g.resistance >= 0 && row[g.resistance] != null ? row[g.resistance] : '');
      newRow.push(g.surface >= 0 && row[g.surface] != null ? row[g.surface] : '');
      outRows.push(newRow);
    });
  });

  if (!outRows.length) return null;
  return { headers: outHeaders, rows: outRows };
}

function normalizeCsvRows_(rows) {
  if (!rows || !rows.length) return [];
  let maxCols = 0;
  rows.forEach(function(r) {
    if (Array.isArray(r) && r.length > maxCols) maxCols = r.length;
  });
  return rows.map(function(r) {
    const out = Array.isArray(r) ? r.slice() : [];
    while (out.length < maxCols) out.push('');
    return out;
  });
}

function detectCsvDelimiter_(text) {
  const lines = String(text || '').split(/\r?\n/).slice(0, 12);
  var bestSemi = 0;
  var bestComma = 0;

  lines.forEach(function(line) {
    if (!line) return;
    var semi = (line.match(/;/g) || []).length;
    var comma = (line.match(/,/g) || []).length;
    if (semi > bestSemi) bestSemi = semi;
    if (comma > bestComma) bestComma = comma;
  });

  if (bestSemi === 0 && bestComma === 0) return ',';
  return bestSemi >= bestComma ? ';' : ',';
}

function setAnacCsvUrlByPrompt_(propKey, title, promptText) {
  const ui = SpreadsheetApp.getUi();
  const current = PropertiesService.getScriptProperties().getProperty(propKey) || '';
  const prompt = ui.prompt(title, promptText + '\n\nCurrent: ' + (current || '(not set)'), ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return;

  const val = String(prompt.getResponseText() || '').trim();
  if (!/^https?:\/\//i.test(val)) {
    ui.alert('Invalid URL. Please provide a full http(s) CSV URL.');
    return;
  }

  PropertiesService.getScriptProperties().setProperty(propKey, val);
  ui.alert('Saved ' + propKey + '.');
}

function getRequiredScriptProperty_(key) {
  const value = String(PropertiesService.getScriptProperties().getProperty(key) || '').trim();
  if (!value) {
    throw new Error('Missing script property: ' + key + '. Use Airport Sync menu to set it first.');
  }
  return value;
}

function buildAnacPullSummaryMessage_(summary) {
  return [
    'ANAC pull complete: ' + summary.source,
    'Temp sheet: ' + summary.tempSheet,
    'Headers: ' + summary.headerCount,
    'Rows written: ' + summary.rowsWritten
  ].join('\n');
}

function ensureDbAirportSyncColumns_(sheet, requiredNames) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    sheet.insertColumnsBefore(1, requiredNames.length);
    sheet.getRange(1, 1, 1, requiredNames.length).setValues([requiredNames]);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const normalized = headers.map(h => normalizeHeader_(h));
  const toAdd = requiredNames.filter(function(name) {
    return normalized.indexOf(normalizeHeader_(name)) === -1;
  });

  if (!toAdd.length) return;

  const startCol = sheet.getLastColumn() + 1;
  sheet.insertColumnsAfter(sheet.getLastColumn(), toAdd.length);
  sheet.getRange(1, startCol, 1, toAdd.length).setValues([toAdd]);
}

function buildAirportColumnIndex_(dbHeaders, tempHeaders) {
  const dbFind = function() {
    return findHeaderIndex_(dbHeaders, Array.prototype.slice.call(arguments));
  };
  const tempFind = function() {
    return findHeaderIndexSmart_(tempHeaders, Array.prototype.slice.call(arguments));
  };
  const tempFindAll = function() {
    const aliases = Array.prototype.slice.call(arguments);
    const seen = {};
    const out = [];
    for (let i = 0; i < aliases.length; i++) {
      const idx = findHeaderIndexSmart_(tempHeaders, [aliases[i]]);
      if (idx >= 0 && !seen[idx]) {
        seen[idx] = true;
        out.push(idx);
      }
    }
    return out;
  };
  const tempFindContainsAll = function(tokens) {
    return findHeaderIndicesByContains_(tempHeaders, tokens || []);
  };

  return {
    db: {
      icao: dbFind('ICAO', 'OACI', 'ICAO_ID'),
      runway: dbFind('RWY_IDENT', 'RWY', 'RUNWAY', 'RUNWAY_DESIGNATOR'),
      name: dbFind('NOME', 'NAME', 'AIRPORT_NAME'),
      infraType: dbFind('INFRASTRUCTURE_TYPE'),
      surface: dbFind('SURFACE_OFFICIAL', 'SURFACE'),
      length: dbFind('LENGTH_OFFICIAL', 'LENGTH_METERS', 'LENGTH_M', 'COMPRIMENTO'),
      width: dbFind('WIDTH_OFFICIAL', 'WIDTH_METERS', 'WIDTH_M', 'LARGURA'),
      elevation: dbFind('ELEVATION', 'ALTITUDE', 'ELEVATION_FT', 'ALT_FEET'),
      lat: dbFind('LATITUDE', 'LAT'),
      lon: dbFind('LONGITUDE', 'LON', 'LNG'),
      slope: dbFind('SLOPE_PERCENT', 'SLOPE_PCT'),
      source: dbFind('SOURCE'),
      manualLock: dbFind('MANUAL_LOCK'),
      lastSyncAt: dbFind('LAST_SYNC_AT')
    },
    temp: {
      icao: tempFind('CÓDIGO OACI', 'CODIGO OACI', 'CODIGO_OACI', 'OACI', 'ICAO', 'ICAO_ID', 'AERODROMO_OACI', 'AERODROMO_ICAO'),
      runway: tempFind('DESIGNAÇÃO', 'DESIGNACAO', 'DESIGNAÇÃO_1', 'DESIGNACAO_1', 'DESIGNAÇÃO 1', 'DESIGNACAO 1', 'DESIGNAÇÃO DA PISTA', 'DESIGNACAO DA PISTA', 'DESIGNACAO_DA_PISTA', 'IDENTIFICADOR_DA_PISTA', 'PISTA', 'PISTA_1', 'RUNWAY', 'RWY', 'CABECEIRA'),
      name: tempFind('NOME', 'AERÓDROMO', 'AERODROMO', 'AIRPORT_NAME', 'NOME DO AERODROMO'),
      surface: tempFind('SUPERFÍCIE', 'SUPERFICIE', 'SUPERFÍCIE_1', 'SUPERFICIE_1', 'SURFACE', 'TIPO SUPERFICIE', 'TIPO_SUPERFICIE', 'REVESTIMENTO'),
      length: tempFind('COMPRIMENTO', 'COMPRIMENTO_1', 'LENGTH', 'LENGTH_M', 'LENGTH_METERS', 'COMPRIMENTO_PISTA', 'COMPRIMENTO_TOTAL'),
      width: tempFind('LARGURA', 'LARGURA_1', 'WIDTH', 'WIDTH_M', 'WIDTH_METERS', 'LARGURA_PISTA', 'LARGURA_TOTAL'),
      elevation: tempFind('ALTITUDE', 'ELEV', 'ELEVATION', 'ELEVATION_FT', 'ELEVACAO'),
      lat: tempFind('LATITUDE', 'LAT', 'LATGEOPOINT', 'LAT_GEOPOINT', 'LATITUDE_GEOPOINT', 'LAT_GEO', 'COORD_LAT'),
      lon: tempFind('LONGITUDE', 'LON', 'LNG', 'LONGEOPOINT', 'LON_GEOPOINT', 'LONGITUDE_GEOPOINT', 'LON_GEO', 'COORD_LON'),
      slope: tempFind('DECLIVIDADE', 'SLOPE', 'SLOPE_PCT', 'SLOPE_PERCENT'),
      surfaceCandidates: mergeUniqueIndices_(
        tempFindAll('SUPERFÍCIE', 'SUPERFICIE', 'SUPERFÍCIE_1', 'SUPERFICIE_1', 'SURFACE', 'TIPO SUPERFICIE', 'TIPO_SUPERFICIE', 'REVESTIMENTO'),
        tempFindContainsAll(['SUPERFICIE', 'REVESTIMENTO'])
      ),
      lengthCandidates: mergeUniqueIndices_(
        tempFindAll('COMPRIMENTO', 'COMPRIMENTO_1', 'LENGTH', 'LENGTH_M', 'LENGTH_METERS', 'COMPRIMENTO_PISTA', 'COMPRIMENTO_TOTAL', 'COMPRIMENTO_M'),
        tempFindContainsAll(['COMPRIMENTO'])
      ),
      widthCandidates: mergeUniqueIndices_(
        tempFindAll('LARGURA', 'LARGURA_1', 'WIDTH', 'WIDTH_M', 'WIDTH_METERS', 'LARGURA_PISTA', 'LARGURA_TOTAL', 'LARGURA_M'),
        tempFindContainsAll(['LARGURA'])
      ),
      elevationCandidates: mergeUniqueIndices_(
        tempFindAll('ALTITUDE', 'ELEV', 'ELEVATION', 'ELEVATION_FT', 'ELEVACAO'),
        tempFindContainsAll(['ALTITUDE', 'ELEV'])
      ),
      latCandidates: mergeUniqueIndices_(
        tempFindAll('LATITUDE', 'LAT', 'LATGEOPOINT', 'LAT_GEOPOINT', 'LATITUDE_GEOPOINT', 'LAT_GEO', 'COORD_LAT'),
        tempFindContainsAll(['LATITUDE', 'LAT'])
      ),
      lonCandidates: mergeUniqueIndices_(
        tempFindAll('LONGITUDE', 'LON', 'LNG', 'LONGEOPOINT', 'LON_GEOPOINT', 'LONGITUDE_GEOPOINT', 'LON_GEO', 'COORD_LON'),
        tempFindContainsAll(['LONGITUDE', 'LON', 'LNG'])
      ),
      slopeCandidates: mergeUniqueIndices_(
        tempFindAll('DECLIVIDADE', 'SLOPE', 'SLOPE_PCT', 'SLOPE_PERCENT'),
        tempFindContainsAll(['DECLIVIDADE', 'SLOPE'])
      )
    }
  };
}

function mapTempAirportData_(row, tempIdx) {
  const pick = function(primaryIdx, candidateIndices) {
    return firstNonBlankCell_(row, candidateIndices, primaryIdx);
  };

  return {
    name: safeCell_(row, tempIdx.name),
    surface: pick(tempIdx.surface, tempIdx.surfaceCandidates),
    length: parseRunwayMetricLoose_(pick(tempIdx.length, tempIdx.lengthCandidates), 'length'),
    width: parseRunwayMetricLoose_(pick(tempIdx.width, tempIdx.widthCandidates), 'width'),
    elevation: parseNumberLoose_(pick(tempIdx.elevation, tempIdx.elevationCandidates)),
    lat: parseCoordinateLoose_(pick(tempIdx.lat, tempIdx.latCandidates), 'lat'),
    lon: parseCoordinateLoose_(pick(tempIdx.lon, tempIdx.lonCandidates), 'lon'),
    slope: parseNumberLoose_(pick(tempIdx.slope, tempIdx.slopeCandidates))
  };
}

function firstNonBlankCell_(row, candidateIndices, fallbackIdx) {
  const candidates = Array.isArray(candidateIndices) ? candidateIndices : [];
  for (let i = 0; i < candidates.length; i++) {
    const val = safeCell_(row, candidates[i]);
    if (String(val == null ? '' : val).trim() !== '') return val;
  }
  return safeCell_(row, fallbackIdx);
}

function applyAirportRowUpdate_(row, dbIdx, payload) {
  let changed = false;
  const isProtected = function(colIdx) {
    return payload && payload.protectedCols && payload.protectedCols[colIdx] === true;
  };

  changed = setCellIfChanged_(row, dbIdx.icao, payload.icao, false, isProtected(dbIdx.icao)) || changed;
  changed = setCellIfChanged_(row, dbIdx.runway, payload.runway, false, isProtected(dbIdx.runway)) || changed;
  changed = setCellIfChanged_(row, dbIdx.source, payload.source, false, isProtected(dbIdx.source)) || changed;
  changed = setCellIfChanged_(row, dbIdx.lastSyncAt, payload.timestamp, false, isProtected(dbIdx.lastSyncAt)) || changed;
  changed = setCellIfChanged_(row, dbIdx.infraType, payload.infrastructureType, false, isProtected(dbIdx.infraType)) || changed;

  changed = setCellIfChanged_(row, dbIdx.name, payload.mapped.name, true, isProtected(dbIdx.name)) || changed;
  changed = setCellIfChanged_(row, dbIdx.surface, payload.mapped.surface, true, isProtected(dbIdx.surface)) || changed;
  changed = setCellIfChanged_(row, dbIdx.length, payload.mapped.length, true, isProtected(dbIdx.length)) || changed;
  changed = setCellIfChanged_(row, dbIdx.width, payload.mapped.width, true, isProtected(dbIdx.width)) || changed;
  changed = setCellIfChanged_(row, dbIdx.elevation, payload.mapped.elevation, true, isProtected(dbIdx.elevation)) || changed;
  changed = setCellIfChanged_(row, dbIdx.lat, payload.mapped.lat, true, isProtected(dbIdx.lat)) || changed;
  changed = setCellIfChanged_(row, dbIdx.lon, payload.mapped.lon, true, isProtected(dbIdx.lon)) || changed;
  changed = setCellIfChanged_(row, dbIdx.slope, payload.mapped.slope, true, isProtected(dbIdx.slope)) || changed;

  return changed;
}

function setCellIfChanged_(row, idx, value, skipIfBlank, skipIfProtected) {
  if (idx < 0) return false;
  if (skipIfProtected) return false;
  if (skipIfBlank && (value === '' || value == null)) return false;

  const oldVal = row[idx];
  if (String(oldVal) === String(value)) return false;
  row[idx] = value;
  return true;
}

function rowKey_(row, dbIdx) {
  const icao = normalizeIcao_(safeCell_(row, dbIdx.icao));
  const rwy = normalizeRunwayToken_(safeCell_(row, dbIdx.runway));
  if (!icao || !rwy) return '';
  return icao + '-' + rwy;
}

function isManualLocked_(value) {
  const txt = String(value == null ? '' : value).trim().toUpperCase();
  return txt === 'Y' || txt === 'YES' || txt === 'TRUE' || txt === '1' || txt === 'LOCK';
}

function splitRunwayDesignators_(rawValue) {
  const raw = runwayValueToString_(rawValue)
    .toUpperCase()
    .replace(/\s+E\s+/g, '/')
    .replace(/\s+/g, ' ')
    .trim();
  if (!raw) return [];

  const parts = raw
    .split(/[\/;|,]+/)
    .map(function(p) { return normalizeRunwayToken_(p); })
    .filter(function(p) { return !!p; });

  const unique = [];
  const seen = {};
  parts.forEach(function(p) {
    if (seen[p]) return;
    seen[p] = true;
    unique.push(p);
  });
  return unique;
}

function normalizeRunwayToken_(token) {
  const raw = String(token == null ? '' : token).trim().toUpperCase();
  if (!raw) return '';

  if (/^WATER$/.test(raw)) return 'WATER';
  if (/^H\d{1,2}$/.test(raw)) return raw;
  if (/\s/.test(raw)) return '';

  const m = raw.match(/^(\d{1,2})([LRC]?)$/);
  if (!m) return '';

  const n = parseInt(m[1], 10);
  if (isNaN(n) || n < 1 || n > 36) return raw;
  return String(n).padStart(2, '0') + (m[2] || '');
}

function runwayValueToString_(rawValue) {
  if (rawValue == null || rawValue === '') return '';
  if (Object.prototype.toString.call(rawValue) === '[object Date]' && !isNaN(rawValue.getTime())) {
    const mm = String(rawValue.getMonth() + 1).padStart(2, '0');
    const dd = String(rawValue.getDate()).padStart(2, '0');
    return mm + '/' + dd;
  }
  return String(rawValue);
}

function normalizeIcao_(value) {
  return String(value == null ? '' : value).trim().toUpperCase();
}

function parseNumberLoose_(value) {
  if (value == null || value === '') return '';
  let cleaned = String(value)
    .split(';')[0]
    .replace(/\u00A0/g, ' ')
    .replace(/\t/g, ' ')
    .trim();

  cleaned = cleaned.replace(/\s+/g, '');
  cleaned = cleaned.replace(/[^0-9,\.\-]/g, '');
  if (!cleaned) return '';

  const commaCount = (cleaned.match(/,/g) || []).length;
  const dotCount = (cleaned.match(/\./g) || []).length;
  const lastComma = cleaned.lastIndexOf(',');
  const lastDot = cleaned.lastIndexOf('.');

  if (lastComma >= 0 && lastDot >= 0) {
    if (lastComma > lastDot) {
      cleaned = cleaned.replace(/\./g, '').replace(/,/g, '.');
    } else {
      cleaned = cleaned.replace(/,/g, '');
    }
  } else if (lastComma >= 0 && dotCount === 0) {
    const parts = cleaned.split(',');
    const allTriplets = parts.length > 1 && parts.slice(1).every(function(p) { return p.length === 3; });

    if (parts.length === 2 && parts[1].length <= 2) {
      cleaned = parts[0] + '.' + parts[1];
    } else if (allTriplets) {
      cleaned = parts.join('');
    } else {
      cleaned = parts.join('.');
    }
  } else if (lastDot >= 0 && commaCount === 0) {
    const dotParts = cleaned.split('.');
    const allTripletsDot = dotParts.length > 1 && dotParts.slice(1).every(function(p) { return p.length === 3; });
    if (allTripletsDot) cleaned = dotParts.join('');
  }

  cleaned = cleaned.replace(/[^0-9.\-]/g, '');
  if (!cleaned) return '';
  const num = parseFloat(cleaned);
  if (isNaN(num)) return '';

  // Heuristic for malformed runway metrics like 18,000,000 that should be 1800.
  if (num > 1000000) {
    let scaled = num;
    while (scaled > 10000 && scaled % 10 === 0) {
      scaled = scaled / 10;
    }
    return scaled;
  }

  return num;
}

function parseRunwayMetricLoose_(value, metricType) {
  if (value == null || value === '') return '';

  const raw = String(value).trim();
  let num = parseNumberLoose_(raw);
  if (num === '' || isNaN(num)) return '';

  const limit = metricType === 'width' ? 500 : 10000;
  const minReasonable = metricType === 'width' ? 3 : 50;

  while (num > limit && Number.isInteger(num) && num % 10 === 0) {
    num = num / 10;
  }

  if (num > limit) {
    const digitsOnly = raw.replace(/[^0-9\-]/g, '');
    if (digitsOnly) {
      let scaled = parseInt(digitsOnly, 10);
      while (scaled > limit && scaled % 10 === 0) {
        scaled = scaled / 10;
      }
      if (scaled >= minReasonable && scaled <= limit) {
        num = scaled;
      }
    }
  }

  if (num < minReasonable || num > limit) return '';
  return num;
}

function findHeaderIndexSmart_(headers, aliases) {
  const idx = findHeaderIndex_(headers, aliases);
  if (idx >= 0) return idx;

  const normalizedHeaders = headers.map(function(h) { return normalizeHeader_(h); });
  for (let i = 0; i < aliases.length; i++) {
    const token = normalizeHeader_(aliases[i]);
    if (!token) continue;
    const partialIdx = normalizedHeaders.findIndex(function(h) {
      if (h === token) return true;
      const parts = h.split('_').filter(function(p) { return !!p; });
      if (parts.indexOf(token) >= 0) return true;
      if (token.length >= 5 && h.indexOf(token) >= 0) return true;
      return false;
    });
    if (partialIdx >= 0) return partialIdx;
  }
  return -1;
}

function stringifyRow_(row) {
  return (row || []).map(function(cell) {
    return cell == null ? '' : String(cell);
  });
}

function findHeaderIndicesByContains_(headers, tokens) {
  const normalizedHeaders = headers.map(function(h) { return normalizeHeader_(h); });
  const normalizedTokens = (tokens || [])
    .map(function(t) { return normalizeHeader_(t); })
    .filter(function(t) { return !!t; });

  const out = [];
  normalizedHeaders.forEach(function(h, idx) {
    if (!h) return;
    const hasToken = normalizedTokens.some(function(t) {
      return h.indexOf(t) >= 0;
    });
    if (hasToken) out.push(idx);
  });
  return out;
}

function mergeUniqueIndices_() {
  const out = [];
  const seen = {};
  for (let i = 0; i < arguments.length; i++) {
    const list = Array.isArray(arguments[i]) ? arguments[i] : [];
    for (let j = 0; j < list.length; j++) {
      const idx = list[j];
      if (typeof idx !== 'number' || idx < 0 || seen[idx]) continue;
      seen[idx] = true;
      out.push(idx);
    }
  }
  return out;
}

function parseCoordinateLoose_(value, axis) {
  if (value == null || value === '') return '';
  const txt = String(value).trim();

  const spacedDecimal = txt.match(/^\s*([+\-]?\d{1,3})\s+(\d{3,})\s*$/);
  if (spacedDecimal) {
    const signNum = parseInt(spacedDecimal[1], 10);
    const sign = signNum < 0 ? '-' : '';
    const combined = sign + String(Math.abs(signNum)) + '.' + spacedDecimal[2];
    const n = parseFloat(combined);
    if (!isNaN(n) && isCoordinateInRange_(n, axis)) return n;
  }

  const dmsMatch = txt.match(/(\d{1,3})[^0-9]+(\d{1,2})(?:[^0-9]+(\d{1,2}(?:[\.,]\d+)?))?[^0-9A-Z]*([NSEW])/i);
  if (dmsMatch) {
    const deg = parseFloat(dmsMatch[1]);
    const min = parseFloat(dmsMatch[2]);
    const sec = dmsMatch[3] ? parseFloat(String(dmsMatch[3]).replace(',', '.')) : 0;
    const hemi = String(dmsMatch[4]).toUpperCase();
    if (!isNaN(deg) && !isNaN(min) && !isNaN(sec)) {
      let out = deg + (min / 60) + (sec / 3600);
      if (hemi === 'S' || hemi === 'W') out *= -1;
      if (isCoordinateInRange_(out, axis)) return out;
    }
  }

  const fallback = parseNumberLoose_(txt);
  if (fallback === '' || isNaN(fallback)) return '';
  return isCoordinateInRange_(fallback, axis) ? fallback : '';
}

function isCoordinateInRange_(value, axis) {
  if (value == null || value === '' || isNaN(value)) return false;
  if (axis === 'lat') return value >= -90 && value <= 90;
  if (axis === 'lon') return value >= -180 && value <= 180;
  return value >= -180 && value <= 180;
}

function findHeaderIndex_(headers, aliases) {
  const norm = headers.map(function(h) { return normalizeHeader_(h); });
  for (let i = 0; i < aliases.length; i++) {
    const idx = norm.indexOf(normalizeHeader_(aliases[i]));
    if (idx >= 0) return idx;
  }
  return -1;
}

function buildProtectedColumnIndexSet_(dbHeaders, protectedHeaderAliases) {
  const set = {};
  (protectedHeaderAliases || []).forEach(function(name) {
    const idx = findHeaderIndex_(dbHeaders, [name]);
    if (idx >= 0) set[idx] = true;
  });
  return set;
}

function normalizeHeader_(text) {
  return String(text == null ? '' : text)
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9\s_]/g, '')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_');
}

function safeCell_(row, idx) {
  if (!Array.isArray(row) || idx < 0 || idx >= row.length) return '';
  return row[idx];
}

function writeAirportSyncReport_(ss, stats) {
  const name = 'LOG_AirportSync';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  const headers = [
    'RUN_AT',
    'SOURCE',
    'TEMP_SHEET',
    'ROWS_READ',
    'RUNWAY_ROWS_PARSED',
    'INSERTED',
    'UPDATED',
    'UNCHANGED',
    'LOCKED_SKIPPED',
    'PARSE_ERRORS',
    'WARNINGS'
  ];

  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const runAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const warningText = (stats.warnings || []).slice(0, 10).join(' | ');

  sh.appendRow([
    runAt,
    stats.source,
    stats.tempSheet,
    stats.rowsRead,
    stats.runwayRowsParsed,
    stats.inserted,
    stats.updated,
    stats.unchanged,
    stats.lockedSkipped,
    stats.parseErrors,
    warningText
  ]);
}

function buildAirportSyncSummaryMessage_(summary) {
  return [
    'Airport sync complete: ' + summary.source,
    'Temp rows read: ' + summary.rowsRead,
    'Runway rows parsed: ' + summary.runwayRowsParsed,
    'Inserted: ' + summary.inserted,
    'Updated: ' + summary.updated,
    'Unchanged: ' + summary.unchanged,
    'Locked skipped: ' + summary.lockedSkipped,
    'Parse errors: ' + summary.parseErrors,
    "Report tab: LOG_AirportSync"
  ].join('\n');
}
