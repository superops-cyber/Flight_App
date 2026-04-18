const APP_DEBUG = false;

const APP_SHEETS = {
  AIRPORTS: 'DB_Airports',
  AIRCRAFT: 'DB_Aircraft',
  PASSENGERS: 'DB_Passengers',
  DISPATCH: 'DB_Dispatch',
  PILOTS: 'DB_Pilots',
  PILOT_AUTHORIZATIONS: 'DB_Pilot_Authorizations',
  FUNDS: 'DB_Funds',
  CHECKS: 'DB_Checks',
  AUDIT: 'LOG_Audit',
  DUTY_LOG: 'LOG_Duty',
  TRANSACTIONS: 'DB_Transactions',
  SYLLABUS: 'Ref_Syllabus',
  WAYPOINTS: 'DB_Waypoints',
  ROUTES: 'DB_Routes',
  LOG_FLIGHTS: 'LOG_Flights',
  AIRFRAMES: 'REF_Airframes',
  ENVELOPES: 'REF_Envelopes',
  AIRCRAFT_DOCS: 'DB_Aircraft_Docs',
  FUEL_LOGS: 'DB_Fuel_Logs',
  FUEL_CACHES: 'DB_Fuel_Caches',
  FLIGHT_BLOCKS: 'DB_Flight_Blocks',
  CONTACTS: 'DB_Contacts',
  DONATION_IMPORT_BATCHES: 'DB_Donation_Import_Batches',
  DONATION_STAGING: 'DB_Donation_Staging',
  DONATIONS_LEDGER: 'DB_Donations_Ledger',
  FUND_LEDGER: 'DB_Fund_Ledger',
  MISSION_FUNDING_ALLOCATIONS: 'DB_Mission_Funding_Allocations',
  STAFF_ROLES: 'REF_Staff_Roles',
  TRAINING_MODULES: 'REF_Training_Modules',
  STAFF_TRAINING: 'DB_Staff_Training_Records',
  STAFF_PRACTICALS: 'DB_Staff_Practical_Evaluations',
  MAINT_TEMPLATES: 'DB_Maint_Templates',
  MAINT_ASSIGNMENTS: 'DB_Maint_Assignments',
  MAINT_LOG: 'DB_Maint_Log',
  SCHED_CONFIG: 'SCHED_Config',
  SCHED_PERMISSIONS: 'SCHED_Permissions',
  SCHED_COVERAGE_RULES: 'SCHED_Coverage_Requirements',
  SCHED_ROLE_COMPAT: 'SCHED_Role_Compatibility',
  SCHED_STAFF_QUALS: 'SCHED_Staff_Qualifications',
  SCHED_ASSIGNMENTS: 'SCHED_Assignments',
  SCHED_LOCKS: 'SCHED_Assignment_Locks',
  SCHED_ALERTS: 'SCHED_Alerts',
  SCHED_PUBLISH_LOG: 'SCHED_Publish_Log',
  SCHED_AVAILABILITY: 'SCHED_Staff_Availability',
  FORM1_RESPONSES: 'Respostas ao formulário 1',
  FORM2_RESPONSES: 'Form Responses 2',
  FORM3_RESPONSES: 'Form Responses 3',
  FORM1_SPREADSHEET_ID: '1cGgfm9A8GYUvVEv4jv7ZIn-ilRqpFkZhpd2aEsh0x2Y',
  FORM2_SPREADSHEET_ID: '1RVK-sbIvuQex7tgpCwVW1yYT0cKD1w4LQRxd4bOW640',
  FORM3_SPREADSHEET_ID: '1kc_SJAitPgTgfZUo09tuhLJKG-qo_b1TTCNZSAmgnf0',
  LIABILITY_RELEASES: 'DB_Liability_Releases',
  FORM_IMPORT_LOG: 'DB_Form_Import_Log'
};

const DISPATCH_COL = {
  FLIGHT_ID: 0,
  MISSION_ID: 1,
  DATE: 2,
  AIRCRAFT: 3,
  PILOT: 4,
  COPILOT: 5,
  ROUTE: 6,
  FLIGHT_TIME: 7,
  TYPE: 8,
  RAW_DATA: 9,
  NOTES: 10,
  STATUS: 11
};

const FUEL_LOG_COL = {
  TIMESTAMP: 0,
  ICAO: 1,
  AIRPORT_NAME: 2,
  AIRCRAFT: 3,
  PILOT: 4,
  CHANGE_QTY: 5,
  TYPE: 6,
  VERIFIED: 7,
  FLIGHT_ID: 8
};

const FUEL_CACHE_COL = {
  ICAO: 0,
  LOCATION_NAME: 1,
  CURRENT_QTY: 2,
  FUEL_TYPE: 3,
  MIN_THRESHOLD: 4
};

const LOG_FLIGHT_COL = {
  FLIGHT_ID: 0,
  DATE: 1,
  PILOT: 2,
  ACFT: 3,
  FROM: 4,
  TO: 5,
  DIST: 6,
  START_TACH: 7,
  TOTAL_TIME: 9,
  FUEL_START: 10,
  OIL: 13,
  VOLTS: 14,
  SQUAWKS: 15,
  TO_RISK_MATRIX: 16,
  BRAKES_RELEASE: 17,
  ACTUAL_LOAD_JSON: 18,
  LANDING_RISK_MATRIX: 19,
  END_TACH: 8,
  FUEL_END: 11,
  FUEL_USED: 12,
  NUM_LDGS: 20,
  AIRBORNE: 21,
  LANDED: 22,
  BRAKES_APPLIED: 23,
  ACTUAL_TO_ROLL: 24,
  NUM_TOUCH_AND_GOS: 25,
  ON_BLOCKS: 23
};

const CHECKS_COL = {
  PILOT: 0,
  AUTH_DESTINATIONS: 7
};

const DUTY_LOG_COL = {
  DATE: 0,
  PILOT: 1,
  TITLE: 2,
  DESC_FALLBACK: 4,
  DESC_PRIMARY: 6
};

const FLIGHT_BLOCKS_COL = {
  BLOCK_ID:      0,
  NAME:          1,
  AIRCRAFT:      2,
  TYPE:          3,
  ALLOCATED_HRS: 4,
  DATE_START:    5,
  DATE_END:      6,
  NOTES:         7,
  STATUS:        8,
  CREATED_AT:    9
};

function appLog_() {
  if (!APP_DEBUG || typeof console === 'undefined' || !console.log) return;
  console.log.apply(console, arguments);
}

function safeNumber_(val, fallbackValue) {
  const parsed = parseFloat(val);
  return isNaN(parsed) ? (fallbackValue || 0) : parsed;
}

function safeDateYmd_(val) {
  if (!val) return '';
  if (val instanceof Date) return val.toISOString().split('T')[0];
  const parsed = new Date(val);
  if (!isNaN(parsed.getTime())) return parsed.toISOString().split('T')[0];
  return String(val);
}

function routeTokensFromString_(routeValue) {
  const raw = String(routeValue || '').trim().toUpperCase();
  if (!raw) return [];
  return raw
    .replace(/[→>]/g, ',')
    .split(/[\n\r,;\/|]+/)
    .map(function(part) { return String(part || '').trim().toUpperCase(); })
    .filter(Boolean);
}

function splitRoute_(routeValue) {
  const route = String(routeValue || '');
  let parts = routeTokensFromString_(route);
  // Legacy fallback: old route strings used "AAA - BBB - CCC".
  if (parts.length < 2 && /\s+-\s+/.test(route)) {
    parts = route
      .split(/\s+-\s+/)
      .map(function(part) { return String(part || '').trim().toUpperCase(); })
      .filter(Boolean);
  }
  // Last-resort legacy fallback for endpoint extraction only.
  if (parts.length < 2 && route.indexOf('-') >= 0) {
    parts = route
      .split('-')
      .map(function(part) { return String(part || '').trim().toUpperCase(); })
      .filter(Boolean);
  }
  return {
    from: parts[0] || '',
    to: parts[parts.length - 1] || ''
  };
}

function safeJsonParse_(rawValue, fallbackValue) {
  if (!rawValue) return fallbackValue;
  try {
    return JSON.parse(rawValue);
  } catch (e) {
    return fallbackValue;
  }
}

function missionIdFromFlightLeg_(flightLegId) {
  const parts = String(flightLegId || '').split('-');
  if (parts.length < 2) return '';
  return parts[0] + '-' + parts[1];
}

function validateMissionPayload_(data) {
  const errors = [];

  if (!data || typeof data !== 'object') {
    throw new Error('Invalid mission payload: payload is empty or not an object');
  }

  if (!String(data.date || '').trim()) errors.push('date is required');
  if (!String(data.acft || '').trim()) errors.push('acft is required');
  if (!String(data.pilot || '').trim()) errors.push('pilot is required');

  const legs = Array.isArray(data.legs) ? data.legs : [];
  if (!legs.length) errors.push('at least one leg is required');

  legs.forEach((leg, index) => {
    const label = 'leg #' + (index + 1);
    if (!leg || typeof leg !== 'object') {
      errors.push(label + ' is missing');
      return;
    }
    if (!String(leg.flightLegId || '').trim()) errors.push(label + ' flightLegId is required');
    if (!String(leg.from || '').trim()) errors.push(label + ' from is required');
    if (!String(leg.to || '').trim()) errors.push(label + ' to is required');
    if (safeNumber_(leg.time, NaN) !== safeNumber_(leg.time, NaN) || safeNumber_(leg.time, 0) < 0) {
      errors.push(label + ' time must be a valid number >= 0');
    }
  });

  if (errors.length) {
    throw new Error('Invalid mission payload: ' + errors.join('; '));
  }
}

function rewriteSheetData_(sheet, rows) {
  if (!sheet) throw new Error('rewriteSheetData_: sheet is required');
  const outRows = Array.isArray(rows) ? rows : [];

  sheet.clearContents();
  if (!outRows.length) return;

  const columnCount = outRows[0].length;
  sheet.getRange(1, 1, outRows.length, columnCount).setValues(outRows);
}
