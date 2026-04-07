/**
 * Document Pack sync (Google Docs fichas -> Master Google Sheet)
 *
 * What it updates per row (matched by drive_doc_id):
 * - codigo_interno
 * - revisao
 * - vigencia
 * - status
 * - local_vigente
 * - retencao_minima
 * - ultima_atualizacao
 * - observacoes (on error)
 */

const DOCPACK_PROP_SHEET_ID = 'DOCPACK_MASTER_SHEET_ID';
const DOCPACK_PROP_SHEET_TAB = 'DOCPACK_MASTER_SHEET_TAB';
const DOCPACK_DEFAULT_SHEET_URL = 'https://docs.google.com/spreadsheets/d/195Yps812fb5l2ftWD4lJ3zoAkExEZQyj3J7rZ185G4s/edit';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Document Pack Sync')
    .addItem('Run sync now', 'runDocumentPackSyncFromMenu')
    .addItem('Setup default sheet', 'setupDocumentPackSyncFromMenu')
    .addSeparator()
    .addItem('Create daily trigger', 'createDocumentPackSyncTriggerFromMenu')
    .addItem('Create hourly trigger', 'createDocumentPackSyncTriggerHourlyFromMenu')
    .addItem('Clear sync triggers', 'clearDocumentPackSyncTriggersFromMenu')
    .addToUi();

  if (typeof buildAirportSyncMenu === 'function') {
    buildAirportSyncMenu();
  }
}

function onInstall() {
  onOpen();
}

function setDocumentPackSyncConfig(sheetId, sheetTabName) {
  if (!sheetId) throw new Error('sheetId is required');
  const props = PropertiesService.getScriptProperties();
  props.setProperty(DOCPACK_PROP_SHEET_ID, sheetId);
  if (sheetTabName) props.setProperty(DOCPACK_PROP_SHEET_TAB, sheetTabName);
  Logger.log('Document Pack sync config saved.');
}

function setDocumentPackSyncConfigFromUrl(sheetUrl, sheetTabName) {
  const candidateUrl = String(sheetUrl || '').trim() || DOCPACK_DEFAULT_SHEET_URL;
  const sheetId = extractGoogleId_(candidateUrl);
  if (!sheetId) {
    throw new Error('Could not extract sheet id from URL. Provide a full Google Sheets URL or set DOCPACK_DEFAULT_SHEET_URL.');
  }
  setDocumentPackSyncConfig(sheetId, sheetTabName || '');
}

function setupDocumentPackSync() {
  setDocumentPackSyncConfigFromUrl(DOCPACK_DEFAULT_SHEET_URL, '');
  Logger.log('Document Pack sync configured with default sheet URL.');
}

function setupDocumentPackSyncFromMenu() {
  setupDocumentPackSync();
  SpreadsheetApp.getUi().alert('Document Pack sync configured for the default master sheet.');
}

function createDocumentPackSyncTrigger() {
  clearDocumentPackSyncTriggers();
  ScriptApp.newTrigger('syncDocumentPackFichasToMaster')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
  Logger.log('Trigger created: syncDocumentPackFichasToMaster daily (around 02:00).');
}

function createDocumentPackSyncTriggerFromMenu() {
  createDocumentPackSyncTrigger();
  SpreadsheetApp.getUi().alert('Daily sync trigger created (around 02:00).');
}

function createDocumentPackSyncTriggerHourly() {
  clearDocumentPackSyncTriggers();
  ScriptApp.newTrigger('syncDocumentPackFichasToMaster')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('Trigger created: syncDocumentPackFichasToMaster hourly.');
}

function createDocumentPackSyncTriggerHourlyFromMenu() {
  createDocumentPackSyncTriggerHourly();
  SpreadsheetApp.getUi().alert('Hourly sync trigger created.');
}

function clearDocumentPackSyncTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncDocumentPackFichasToMaster')
    .forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('Existing sync triggers removed.');
}

function clearDocumentPackSyncTriggersFromMenu() {
  clearDocumentPackSyncTriggers();
  SpreadsheetApp.getUi().alert('Document Pack sync triggers removed.');
}

function runDocumentPackSyncNow() {
  return syncDocumentPackFichasToMaster();
}

function runDocumentPackSyncFromMenu() {
  const summary = syncDocumentPackFichasToMaster();
  SpreadsheetApp.getUi().alert(
    'Document Pack sync complete.\n\n' +
    'Inspected: ' + summary.inspected + '\n' +
    'Updated: ' + summary.updated + '\n' +
    'Errors: ' + summary.errors
  );
  return summary;
}

function syncDocumentPackFichasToMaster() {
  const config = getSyncConfig_();
  const spreadsheet = SpreadsheetApp.openById(config.sheetId);
  const sheet = config.sheetTab
    ? spreadsheet.getSheetByName(config.sheetTab)
    : spreadsheet.getSheets()[0];

  if (!sheet) throw new Error('Master sheet tab not found.');

  ensureSyncColumns_(sheet);

  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) {
    return { inspected: 0, updated: 0, errors: 0, message: 'No data rows found.' };
  }

  const headers = values[0].map(h => String(h || '').trim());
  const col = buildHeaderIndex_(headers);

  const required = ['drive_doc_id', 'status', 'revisao', 'ultima_atualizacao'];
  const missing = required.filter(k => col[k] == null);
  if (missing.length) {
    throw new Error('Missing required columns in sheet: ' + missing.join(', '));
  }

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  let inspected = 0;
  let updated = 0;
  let errors = 0;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const docId = String(row[col.drive_doc_id] || '').trim();
    if (!docId) continue;

    inspected++;

    try {
      const text = DocumentApp.openById(docId).getBody().getText();
      const parsed = parseFichaControlFields_(text);
      let changed = false;

      if (col.codigo_interno != null && parsed.codigoInterno) {
        if (String(row[col.codigo_interno] || '').trim() !== parsed.codigoInterno) {
          row[col.codigo_interno] = parsed.codigoInterno;
          changed = true;
        }
      }

      if (parsed.revisao) {
        if (String(row[col.revisao] || '').trim() !== parsed.revisao) {
          row[col.revisao] = parsed.revisao;
          changed = true;
        }
      }

      if (col.vigencia != null && parsed.vigencia) {
        if (String(row[col.vigencia] || '').trim() !== parsed.vigencia) {
          row[col.vigencia] = parsed.vigencia;
          changed = true;
        }
      }

      if (parsed.status) {
        if (String(row[col.status] || '').trim() !== parsed.status) {
          row[col.status] = parsed.status;
          changed = true;
        }
      }

      if (col.local_vigente != null && parsed.localVigente) {
        if (String(row[col.local_vigente] || '').trim() !== parsed.localVigente) {
          row[col.local_vigente] = parsed.localVigente;
          changed = true;
        }
      }

      if (col.retencao_minima != null && parsed.retencaoMinima) {
        if (String(row[col.retencao_minima] || '').trim() !== parsed.retencaoMinima) {
          row[col.retencao_minima] = parsed.retencaoMinima;
          changed = true;
        }
      }

      if (changed) {
        row[col.ultima_atualizacao] = now;
        if (col.observacoes != null) {
          row[col.observacoes] = 'Atualizado automaticamente a partir da ficha';
        }
        updated++;
      }
    } catch (err) {
      errors++;
      if (col.observacoes != null) {
        const message = 'Erro sync ficha: ' + String(err && err.message ? err.message : err);
        row[col.observacoes] = message.slice(0, 250);
      }
    }
  }

  range.setValues(values);

  const summary = { inspected, updated, errors, sheetId: config.sheetId, sheetTab: sheet.getName() };
  Logger.log(JSON.stringify(summary));
  return summary;
}

function ensureSyncColumns_(sheet) {
  const requiredOptionalHeaders = [
    'vigencia',
    'local_vigente',
    'retencao_minima',
  ];

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headerRange = sheet.getRange(1, 1, 1, lastColumn);
  const headers = headerRange.getValues()[0].map(h => String(h || '').trim());
  const normalized = headers.map(normalizeHeader_);

  const missing = requiredOptionalHeaders.filter(name => !normalized.includes(name));
  if (!missing.length) return;

  let insertAt = lastColumn;
  const revisaoIndex = normalized.indexOf('revisao');
  if (revisaoIndex >= 0) {
    insertAt = revisaoIndex + 2;
  }

  missing.forEach((name, idx) => {
    sheet.insertColumnAfter(insertAt + idx - 1 < 1 ? 1 : insertAt + idx - 1);
    sheet.getRange(1, insertAt + idx).setValue(name);
  });
}

function getSyncConfig_() {
  const props = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty(DOCPACK_PROP_SHEET_ID);
  const sheetTab = props.getProperty(DOCPACK_PROP_SHEET_TAB) || '';
  if (!sheetId) {
    throw new Error('Missing config. Run setDocumentPackSyncConfig(...) first.');
  }
  return { sheetId, sheetTab };
}

function buildHeaderIndex_(headers) {
  const idx = {};
  headers.forEach((h, i) => {
    const key = normalizeHeader_(h);
    if (key) idx[key] = i;
  });
  return idx;
}

function normalizeHeader_(header) {
  return String(header || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

function extractGoogleId_(urlOrId) {
  const value = String(urlOrId || '').trim();
  if (!value) return '';
  if (!value.includes('/')) return value;

  const docMatch = value.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (docMatch) return docMatch[1];

  const folderMatch = value.match(/\/folders\/([a-zA-Z0-9-_]+)/);
  if (folderMatch) return folderMatch[1];

  const idMatch = value.match(/[?&]id=([a-zA-Z0-9-_]+)/);
  if (idMatch) return idMatch[1];

  return '';
}

function parseFichaControlFields_(rawText) {
  const text = String(rawText || '').replace(/\r/g, '\n');
  const control = extractControlSection_(text);

  const codigo = sanitizeFieldValue_(extractLabeledValue_(control, '(?:Código interno|Codigo interno)'));
  const revisao = sanitizeFieldValue_(extractLabeledValue_(control, '(?:Revisão|Revisao)'));
  const vigencia = sanitizeFieldValue_(extractLabeledValue_(control, '(?:Vigência|Vigencia)'));
  const status = extractStatus_(control);
  const localVigente = sanitizeFieldValue_(extractLabeledValue_(control, 'Local oficial de armazenamento'));
  const retencaoMinima = sanitizeFieldValue_(extractLabeledValue_(control, '(?:Prazo de retenção|Prazo de retencao)'));

  return {
    codigoInterno: codigo,
    revisao: revisao,
    vigencia: vigencia,
    status: status,
    localVigente: localVigente,
    retencaoMinima: retencaoMinima,
  };
}

function extractControlSection_(text) {
  const match = text.match(/Ficha de controle([\s\S]*?)(?:Finalidade e regra de uso|Formulário controlado|Formulario controlado)/i);
  return match ? match[1] : text;
}

function extractLabeledValue_(text, labelRegex) {
  const re = new RegExp(labelRegex + '\\s*[:\\-]?\\s*([^\\n]+)', 'i');
  const match = String(text || '').match(re);
  return match ? match[1].trim() : '';
}

function sanitizeFieldValue_(value) {
  const v = String(value || '').trim();
  if (!v) return '';
  if (/^[_\-.\s]+$/.test(v)) return '';
  if (/^(\[?\s*\]?|n\/a|na)$/i.test(v)) return '';
  return v;
}

function extractStatus_(text) {
  const t = String(text || '');

  const options = [
    { label: 'Em elaboração', re: /(☑|✅|✔|✓|\[x\]|\(x\)|\bx\b)\s*Em elaboração/i },
    { label: 'Em revisão', re: /(☑|✅|✔|✓|\[x\]|\(x\)|\bx\b)\s*Em revisão/i },
    { label: 'Vigente', re: /(☑|✅|✔|✓|\[x\]|\(x\)|\bx\b)\s*Vigente/i },
    { label: 'Obsoleto', re: /(☑|✅|✔|✓|\[x\]|\(x\)|\bx\b)\s*Obsoleto/i },
  ];

  for (const opt of options) {
    if (opt.re.test(t)) return opt.label;
  }

  const raw = extractLabeledValue_(t, 'Status');
  if (/elabor/i.test(raw)) return 'Em elaboração';
  if (/revis/i.test(raw)) return 'Em revisão';
  if (/vigen/i.test(raw)) return 'Vigente';
  if (/obsol/i.test(raw)) return 'Obsoleto';

  return '';
}
