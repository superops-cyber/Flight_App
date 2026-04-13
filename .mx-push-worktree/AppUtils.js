const APP_DEBUG = false;

const APP_SHEETS = {
  AIRPORTS: 'DB_Airports',
  AIRCRAFT: 'DB_Aircraft',
  PASSENGERS: 'DB_Passengers',
  DISPATCH: 'DB_Dispatch',
  PILOTS: 'DB_Pilots',
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
  FUEL_LOGS: 'DB_Fuel_Logs',
  FUEL_CACHES: 'DB_Fuel_Caches',
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
  MAINT_LOG: 'DB_Maint_Log'
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

function splitRoute_(routeValue) {
  const route = String(routeValue || '');
  const parts = route.split('-');
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
