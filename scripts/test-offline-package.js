const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const backendFiles = ['AppUtils.js', 'Backend_Utils.js', 'Backend_Maintenance.js', 'Backend_Dropdowns.js',
  'Backend_Pilot.js', 'Backend_Dispatch.js', 'Backend_Flight.js', 'Backend_Other.js', 'Backend_OfflinePackage.js'];

function fixture(count = 2) {
  const tables = {};
  function table(name, headers, records = []) {
    tables[name] = [headers, ...records.map(record => headers.map(header => record[header] ?? ''))];
  }
  const dispatchHeaders = ['FLIGHT_ID', 'MISSION_ID', 'DATE', 'AIRCRAFT', 'PILOT', 'COPILOT', 'ROUTE',
    'FLIGHT_TIME', 'TYPE', 'RAW_DATA', 'NOTES', 'STATUS'];
  const missions = [];
  for (let i = 1; i <= count; i++) {
    const id = 'ADS26-' + String(i).padStart(3, '0');
    for (let leg = 1; leg <= 2; leg++) {
      missions.push({
        FLIGHT_ID: id + '-0' + leg, MISSION_ID: id, DATE: '2026-11-06',
        AIRCRAFT: i % 2 ? 'PT-ONE' : 'PT-TWO', PILOT: 'Test Pilot', COPILOT: '', ROUTE: 'SBGR - FIX - SBRJ',
        FLIGHT_TIME: 1, TYPE: 'MISSION', STATUS: 'APPROVED',
        RAW_DATA: JSON.stringify({ pax: [{ name: 'Passenger', weight: 70, cargo: 5 }] })
      });
    }
  }
  table('DB_Dispatch', dispatchHeaders, missions);
  table('DB_Pilots', ['NAME', 'EMAIL', 'STAFF_ID', 'ACTIVE', 'PRIMARY_ROLE', 'WEIGHT_KGS',
    'MEDICAL_EXPIRY', 'MNTE_VALIDITY', 'MNAF_VALIDITY'], [{
    NAME: 'Test Pilot', EMAIL: 'pilot@example.org', STAFF_ID: 'P1', ACTIVE: 'Y', PRIMARY_ROLE: 'PILOT',
    WEIGHT_KGS: 80, MEDICAL_EXPIRY: '2027-01-01', MNTE_VALIDITY: '2027-01-01', MNAF_VALIDITY: '2027-01-01'
  }]);
  const aircraftHeaders = ['REGISTRATION', 'AIRCRAFT_TYPE', 'EMPTY_WEIGHT', 'EMPTY_ARM', 'MTOW', 'CURRENT_TACH',
    'MAX_FUEL', 'NUM_TANKS', 'TANK_NAMES', 'PILOT_SEAT_KGS', 'MID_SEAT_KGS', 'AFT_SEAT_KGS', 'TECH_STATUS',
    'CRUISE_KTS', 'BURN_LPH', 'MAIN_CAPACITY_L', 'TIP_CAPACITY_L', 'TRANSFER_RATE_LPM',
    'NEXT_DUE_TACH', 'HOURS_TO_INSPECTION', 'LICENSE_REQUIRED'];
  table('DB_Aircraft', aircraftHeaders, ['PT-ONE', 'PT-TWO'].map((reg, index) => ({
    REGISTRATION: reg, AIRCRAFT_TYPE: index ? 'C206' : 'C172', EMPTY_WEIGHT: 900, EMPTY_ARM: 40,
    MTOW: 1600, CURRENT_TACH: 100, MAX_FUEL: 200, NUM_TANKS: 2, TANK_NAMES: 'LM,RM',
    PILOT_SEAT_KGS: 10, MID_SEAT_KGS: 10, AFT_SEAT_KGS: 10, TECH_STATUS: 'SERVICEABLE',
    CRUISE_KTS: 120, BURN_LPH: 40, MAIN_CAPACITY_L: 100, TIP_CAPACITY_L: 0,
    TRANSFER_RATE_LPM: 0, NEXT_DUE_TACH: 200, HOURS_TO_INSPECTION: 100, LICENSE_REQUIRED: 'MNTE'
  })));
  table('REF_Airframes', ['AIRCRAFT_TYPE', 'STATION_NAME', 'ARM', 'MAX_WEIGHT_LBS'],
    ['C172', 'C206'].flatMap(type => ['PILOT', 'MID', 'AFT', 'CARGO', 'FUEL'].map(name => ({
      AIRCRAFT_TYPE: type, STATION_NAME: name, ARM: 42, MAX_WEIGHT_LBS: 300
    }))));
  table('REF_Envelopes', ['AIRCRAFT_TYPE', 'POINT_SEQUENCE', 'CG_ARM_X', 'WEIGHT_Y'],
    ['C172', 'C206'].flatMap(type => [[35, 1000], [45, 1000], [45, 1600], [35, 1600]].map((point, index) => ({
      AIRCRAFT_TYPE: type, POINT_SEQUENCE: index, CG_ARM_X: point[0], WEIGHT_Y: point[1]
    }))));
  table('DB_Airports', ['ICAO', 'NOME', 'LATITUDE', 'LONGITUDE', 'RWY_IDENT', 'LENGTH_OFFICIAL',
    'WIDTH_OFFICIAL', 'SURFACE_ACTUAL', 'ELEVATION', 'RUNWAY_HEADING'],
  ['SBGR', 'SBRJ'].map(icao => ({
    ICAO: icao, NOME: icao, LATITUDE: -23, LONGITUDE: -46, RWY_IDENT: '09',
    LENGTH_OFFICIAL: 2000, WIDTH_OFFICIAL: 30, SURFACE_ACTUAL: 'Paved', ELEVATION: 100, RUNWAY_HEADING: 90
  })));
  table('DB_Routes', ['ORIGIN', 'DESTINATION', 'WAYPOINT_LIST'], [{ ORIGIN: 'SBGR', DESTINATION: 'SBRJ', WAYPOINT_LIST: 'FIX' }]);
  table('DB_Waypoints', ['WP_ID', 'LATITUDE', 'LONGITUDE', 'TYPE'], [{ WP_ID: 'FIX', LATITUDE: -23.5, LONGITUDE: -45.5, TYPE: 'FIX' }]);
  table('DB_Passengers', ['PASSENGER_NAME', 'WEIGHT_KG'], [{ PASSENGER_NAME: 'Passenger', WEIGHT_KG: 70 }]);
  table('DB_Checks', ['PILOT_NAME', 'AUTH_DESTINATIONS'], [{ PILOT_NAME: 'Test Pilot', AUTH_DESTINATIONS: 'SBGR,SBRJ' }]);
  table('DB_Funds', ['NAME', 'CURRENT_BALANCE'], [{ NAME: 'General', CURRENT_BALANCE: 0 }]);
  table('Ref_Syllabus', ['TRAINING_CODE', 'REQUIRED_HOURS', 'REQUIRED_FUEL']);
  table('DB_Fuel_Caches', ['ICAO', 'LOCATION_NAME', 'CURRENT_QTY', 'FUEL_TYPE'], [{
    ICAO: 'SBGR', LOCATION_NAME: 'Test', CURRENT_QTY: 1000, FUEL_TYPE: 'AVGAS'
  }]);
  table('DB_Fuel_Logs', ['TIMESTAMP', 'ICAO', 'AIRPORT_NAME', 'AIRCRAFT', 'PILOT', 'CHANGE_QTY', 'TYPE', 'VERIFIED', 'FLIGHT_ID']);
  table('LOG_Flights', ['Flight_ID', 'DATE', 'PILOT', 'ACFT', 'FROM', 'TO', 'DIST', 'START_TACH', 'END_TACH',
    'TOTAL_TIME', 'FUEL_START', 'FUEL_END', 'FUEL_USED', 'OIL', 'VOLTS', 'SQUAWKS', 'TO_RISK_MATRIX',
    'BRAKES_RELEASE', 'ACTUAL_LOAD_JSON', 'LANDING_RISK_MATRIX', 'NUMBER_LDGS', 'AIRBORNE', 'LANDED', 'BRAKES_APPLIED', 'ACTUAL_TO_ROLL']);
  table('SCHED_Config', ['CONFIG_KEY', 'CONFIG_VALUE', 'ACTIVE'], [{
    CONFIG_KEY: 'DUTY_LANDING_THRESHOLD_KTS', CONFIG_VALUE: '30', ACTIVE: 'Y'
  }]);
  table('DB_Maint_Assignments', ['ASSIGNMENT_ID', 'AIRCRAFT_REG', 'TASK_CODE', 'TASK_NAME', 'CATEGORY', 'REFERENCE',
    'DUE_MODE', 'INTERVAL_HOURS', 'INTERVAL_DAYS', 'INTERVAL_MONTHS', 'INTERVAL_YEARS',
    'START_TACH', 'START_DATE', 'ACTIVE', 'NOTES']);
  table('DB_Maint_Log', ['ASSIGNMENT_ID', 'COMPLETED_DATE', 'COMPLETED_TACH', 'CREATED_AT']);
  table('Aircraft_Roll_Numbers', ['AIRCRAFT_TYPE', 'WEIGHT_KG', 'TO_ROLL_M', 'LDG_ROLL_M'],
    ['C172', 'C206'].map(type => ({ AIRCRAFT_TYPE: type, WEIGHT_KG: 1400, TO_ROLL_M: 400, LDG_ROLL_M: 350 })));
  const multiplierHeaders = ['AIRCRAFT_TYPE', 'FLAP_SETTING', 'TO_MULTIPLIER', 'LDG_MULTIPLIER',
    'DA_FT', 'WIND_KTS', 'HUMIDITY', 'HUMIDITY_FACTOR', 'FLAP_FACTOR', 'SLOPE', 'SURFACE',
    'TAKEOFF_HEAD', 'TAKEOFF_TAIL', 'LANDING_HEAD', 'LANDING_TAIL',
    'TAKEOFF_UP', 'TAKEOFF_DOWN', 'LANDING_UP', 'LANDING_DOWN',
    'TAKEOFF_DRY', 'TAKEOFF_WET', 'LANDING_DRY', 'LANDING_WET'];
  table('Performance_Multipliers', multiplierHeaders, ['C172', 'C206'].map(type => ({
    ...Object.fromEntries(multiplierHeaders.map(header => [header, 1])),
    AIRCRAFT_TYPE: type, FLAP_SETTING: 0, DA_FT: 0, WIND_KTS: 0, HUMIDITY: 50, SLOPE: 0, SURFACE: 'Paved'
  })));
  return tables;
}

function harness(tables = fixture(), email = 'pilot@example.org') {
  const writes = [];
  const reads = {};
  const deny = label => (...args) => { writes.push(label); throw new Error('Forbidden side effect: ' + label); };
  const sheet = name => {
    if (!tables[name]) return null;
    return new Proxy({
      getName: () => name,
      getLastRow: () => tables[name].length,
      getLastColumn: () => tables[name][0].length,
      getDataRange: () => ({ getValues: () => {
        reads[name] = (reads[name] || 0) + 1;
        return structuredClone(tables[name]);
      } }),
      getRange: (r, c, nr = 1, nc = 1) => new Proxy({
        getValues: () => tables[name].slice(r - 1, r - 1 + nr).map(row => row.slice(c - 1, c - 1 + nc))
      }, { get: (target, key) => key in target ? target[key] : deny(name + '.' + String(key)) })
    }, { get: (target, key) => key in target ? target[key] : deny(name + '.' + String(key)) });
  };
  class Clock extends Date {
    constructor(...args) { super(...(args.length ? args : ['2026-11-06T15:00:00Z'])); }
    static now() { return new Clock().getTime(); }
  }
  const ss = new Proxy({
    getSheetByName: sheet, getSpreadsheetTimeZone: () => 'America/Sao_Paulo'
  }, { get: (target, key) => key in target ? target[key] : deny('Spreadsheet.' + String(key)) });
  const context = {
    Date: Clock, console: { log() {}, warn() {} },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss },
    Session: { getActiveUser: () => ({ getEmail: () => email }), getScriptTimeZone: () => 'America/Sao_Paulo' },
    Utilities: {
      getUuid: () => crypto.randomUUID(),
      formatDate: (date, zone) => new Intl.DateTimeFormat('en-CA', {
        timeZone: zone === 'GMT' ? 'UTC' : zone, year: 'numeric', month: '2-digit', day: '2-digit'
      }).format(date),
      computeDigest: (algorithm, text) => [...crypto.createHash('sha256').update(text).digest()],
      base64EncodeWebSafe: bytes => Buffer.from(bytes).toString('base64url'),
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' }
    }
  };
  for (const service of ['MailApp', 'UrlFetchApp', 'DriveApp', 'CacheService', 'PropertiesService', 'CalendarApp', 'ScriptApp']) {
    context[service] = new Proxy({}, { get: (_, key) => deny(service + '.' + String(key)) });
  }
  vm.createContext(context);
  for (const file of backendFiles) vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), context, { filename: file });
  return { context, writes, reads, export: days => context.exportOfflineMissionPackage(days) };
}

test('real readers export all missions and legs beyond the dashboard limit, with shared references', () => {
  const tables = fixture(18);
  const before = JSON.stringify(tables);
  const h = harness(tables);
  const pack = h.export(14);
  assert.equal(pack.completeness.ready, true, JSON.stringify(pack.completeness));
  assert.equal(pack.missions.length, 18);
  assert.ok(pack.missions.every(mission => Object.keys(mission.legs).length === 2));
  assert.equal(Object.keys(pack.references.aircraftRecords).length, 2);
  assert.equal(Object.keys(pack.references.envelopes).length, 2);
  assert.equal(Object.keys(pack.references.performanceSetups).length, 4);
  assert.equal(Object.keys(pack.references.performanceReferences).length, 1);
  for (const mission of pack.missions) for (const leg of Object.values(mission.legs)) {
    assert.ok(pack.references.wbConfigurations[leg.wbConfigurationRef]);
    assert.ok(pack.references.envelopes[leg.envelopeRef]);
    assert.equal('envelopeData' in leg.wb, false);
    assert.equal(leg.flightLog, null); // No prior log row is needed before Briefing.
    assert.deepEqual(Array.from(leg.route.waypoints), ['FIX']);
  }
  assert.equal(JSON.stringify(tables), before);
  assert.deepEqual(h.writes, []);
});

test('7/14-day windows include today and exclude the exact upper bound', () => {
  const tables = fixture(4);
  for (const row of tables.DB_Dispatch.slice(1)) {
    row[2] = ['2026-11-06', '2026-11-12', '2026-11-13', '2026-11-20'][Number(row[1].slice(-3)) - 1];
  }
  const h = harness(tables);
  assert.equal(h.export(7).missions.length, 2);
  assert.equal(h.export(14).missions.length, 3);
  assert.throws(() => h.export(10), /7 or 14/);
  assert.deepEqual(h.writes, []);
});

test('a remaining leg inside the window includes the entire mission and preserves approval statuses', () => {
  const tables = fixture(1);
  tables.DB_Dispatch[1][2] = '2026-11-05';
  tables.DB_Dispatch[1][11] = 'FLOWN';
  const h = harness(tables);
  const pack = h.export(7);
  assert.equal(pack.missions.length, 1);
  assert.equal(Object.keys(pack.missions[0].legs).length, 2);
  assert.equal(pack.missions[0].summary.status, '1 OF 2 LEGS');
  assert.equal(pack.missions[0].detail.status, 'FLOWN');
  assert.equal(pack.missions[0].legs['ADS26-001-02'].dispatchStatus, 'APPROVED');
  assert.equal(pack.completeness.ready, false); // Completed leg is missing its log baseline.
  assert.match(pack.completeness.missionResults[0].missing.join(';'), /flight-log baseline/);
  assert.deepEqual(h.writes, []);
});

test('approval, all-leg assignment and active authenticated identity are required', () => {
  const tables = fixture(5);
  for (const row of tables.DB_Dispatch.slice(1)) {
    if (row[1].endsWith('002')) row[11] = 'PENDING';
    if (row[1].endsWith('003')) row[4] = 'Other Pilot';
    if (row[1].endsWith('004') && row[0].endsWith('02')) row[11] = 'CANCELLED';
    if (row[1].endsWith('005')) row[11] = 'FLOWN';
  }
  assert.equal(harness(tables).export(7).missions.length, 1);
  assert.throws(() => harness(tables, '').export(7), /authenticated/);
  assert.throws(() => harness(tables, 'unknown@example.org').export(7), /active staff/);
  tables.DB_Pilots[1][3] = 'N';
  assert.throws(() => harness(tables).export(7), /active staff/);
});

test('existing supervisor role permits other assignments without changing approval', () => {
  const tables = fixture(2);
  tables.DB_Pilots[1][4] = 'SUPERVISOR';
  tables.DB_Dispatch[3][4] = 'Other Pilot';
  tables.DB_Dispatch[4][4] = 'Other Pilot';
  const pack = harness(tables).export(7);
  assert.equal(pack.missions.length, 2);
  assert.equal(pack.missions[1].detail.pilot, 'Other Pilot');
});

test('one missing aircraft reference blocks only its mission and whole-package readiness', () => {
  const tables = fixture();
  tables.REF_Envelopes = tables.REF_Envelopes.filter((row, i) => !i || row[0] !== 'C206');
  const h = harness(tables);
  const pack = h.export(7);
  assert.equal(pack.completeness.ready, false);
  assert.equal(pack.completeness.missionResults[0].ready, true);
  assert.equal(pack.completeness.missionResults[1].ready, false);
  assert.match(pack.completeness.missionResults[1].missing.join(';'), /envelope/);
  assert.deepEqual(h.writes, []);
});

test('missing maintenance sheets/columns are reported without repair or hidden writes', () => {
  for (const damage of [
    tables => { delete tables.DB_Maint_Log; },
    tables => { tables.DB_Maint_Assignments[0][1] = 'WRONG_COLUMN'; }
  ]) {
    const tables = fixture();
    damage(tables);
    const h = harness(tables);
    const pack = h.export(7);
    assert.equal(pack.completeness.ready, false);
    assert.ok(pack.completeness.missionResults.every(result => !result.ready));
    assert.match(JSON.stringify(pack.completeness), /DB_Maint/);
    assert.deepEqual(h.writes, []);
  }
});

test('normal startup and maintenance callers still initialize schemas', () => {
  const h = harness();
  vm.runInContext('var schemaCalls = 0; _mxEnsureFrameworkSheets_ = function() { schemaCalls++; };', h.context);
  assert.equal(h.context.getMaintenanceScheduleData({}).success, true);
  assert.equal(h.context.getDispatchMaintenanceStatus({}).success, true);
  assert.ok(h.context.getDropdownData().aircraft.length);
  assert.ok(h.context.getPilotStartupData().aircraft.length);
  assert.equal(h.context.schemaCalls, 4);
  assert.ok(h.context._getPilotStartupDataReadOnly_().aircraft.length);
  assert.equal(h.context.schemaCalls, 4);
  assert.deepEqual(h.writes, []);
});

test('maintenance assignments, history and forecasts are read without schema initialization', () => {
  const tables = fixture(1);
  const record = {
    ASSIGNMENT_ID: 'inspection-1', AIRCRAFT_REG: 'PT-ONE', TASK_CODE: '100H', TASK_NAME: 'Inspection',
    CATEGORY: 'Inspection', REFERENCE: 'Manual', DUE_MODE: 'HORAS', INTERVAL_HOURS: 100,
    START_TACH: 100, START_DATE: '2026-10-01', ACTIVE: 'Y'
  };
  tables.DB_Maint_Assignments.push(tables.DB_Maint_Assignments[0].map(header => record[header] ?? ''));
  tables.DB_Maint_Log.push(['inspection-1', '2026-10-20', 120, '2026-10-20']);
  const h = harness(tables);
  const pack = h.export(7);
  assert.equal(pack.completeness.ready, true, JSON.stringify(pack.completeness));
  assert.ok(h.reads.DB_Maint_Log >= 2);
  assert.deepEqual(h.writes, []);
});

test('missing runway, performance type, waypoint, malformed Dispatch and duplicate legs fail completeness', () => {
  const damages = [
    tables => { tables.DB_Airports = tables.DB_Airports.filter((row, i) => !i || row[0] !== 'SBRJ'); },
    tables => { tables.Aircraft_Roll_Numbers = tables.Aircraft_Roll_Numbers.filter((row, i) => !i || row[0] !== 'C172'); },
    tables => { tables.DB_Waypoints = [tables.DB_Waypoints[0]]; },
    tables => { tables.REF_Envelopes[1][2] = 'invalid'; },
    tables => { tables.REF_Airframes = tables.REF_Airframes.filter((row, i) => !i || row[1] !== 'MID'); },
    tables => { tables.DB_Dispatch[1][9] = '{broken'; },
    tables => { tables.DB_Dispatch[2][0] = tables.DB_Dispatch[1][0]; }
  ];
  for (const damage of damages) {
    const tables = fixture(1);
    damage(tables);
    const h = harness(tables);
    const pack = h.export(7);
    assert.equal(pack.completeness.ready, false);
    assert.equal(pack.completeness.missionResults[0].ready, false);
    assert.ok(pack.completeness.missionResults[0].missing.length);
    assert.deepEqual(h.writes, []);
  }
});

test('ANAC runway surfaces use the existing performance translator without masking missing factors', () => {
  for (const surface of ['ASFALTO', 'CONCRETO', 'GRAMA', 'TERRA']) {
    const tables = fixture(1);
    for (const row of tables.DB_Airports.slice(1)) row[7] = surface;
    const translated = { ASFALTO: 'Paved', CONCRETO: 'Paved', GRAMA: 'Grass to 6"', TERRA: 'Rough' }[surface];
    for (const row of tables.Performance_Multipliers.slice(1)) row[10] = translated;
    const h = harness(tables);
    const pack = h.export(7);
    assert.equal(pack.completeness.ready, true, JSON.stringify(pack.completeness));
    assert.equal(pack.references.performanceSetups['SBGR|C172'].runways[0].surface, surface);
    assert.deepEqual(h.writes, []);
    for (const row of tables.Performance_Multipliers.slice(1)) row[19] = '';
    assert.equal(h.export(7).completeness.ready, false);
  }
});

test('reordered positional flight-log and fuel-log columns are incomplete rather than misread', () => {
  for (const [sheet, header] of [
    ['DB_Fuel_Logs', 'FLIGHT_ID'], ['DB_Fuel_Logs', 'CHANGE_QTY'],
    ['LOG_Flights', 'Flight_ID'], ['LOG_Flights', 'LANDED'], ['LOG_Flights', 'BRAKES_APPLIED']
  ]) {
    const tables = fixture(1);
    const headers = tables[sheet][0];
    const index = headers.indexOf(header);
    [headers[index], headers[1]] = [headers[1], headers[index]];
    const h = harness(tables);
    const pack = h.export(7);
    assert.equal(pack.completeness.ready, false);
    assert.match(pack.completeness.errors.join(';'), /unexpected column position/);
    assert.ok(pack.completeness.missionResults.every(result => !result.ready));
    assert.deepEqual(h.writes, []);
  }
});

test('flight-log baselines and Date values are exported as JSON-safe data', () => {
  const tables = fixture(1);
  const header = tables.LOG_Flights[0];
  const log = header.map(() => '');
  log[0] = 'ADS26-001-01';
  log[17] = new Date('2026-11-06T14:00:00Z');
  log[18] = '{"tanks":{"LM":50,"RM":50}}';
  tables.LOG_Flights.push(log);
  const h = harness(tables);
  const pack = h.export(7);
  assert.equal(pack.completeness.ready, true, JSON.stringify(pack.completeness));
  assert.equal(pack.missions[0].legs['ADS26-001-01'].flightLog.brakesRelease, '2026-11-06T14:00:00.000Z');
  assert.equal(JSON.parse(JSON.stringify(pack)).missions.length, 1);
  assert.deepEqual(h.writes, []);
});

test('empty authorized collection is not reported ready', () => {
  const h = harness(fixture(0));
  const pack = h.export(7);
  assert.equal(pack.completeness.ready, false);
  assert.match(pack.completeness.errors.join(';'), /No eligible missions/);
  assert.deepEqual(h.writes, []);
});

test('completed flight logs exclude already-flown missions even when Dispatch status is stale', () => {
  const tables = fixture(1);
  for (const row of tables.DB_Dispatch.slice(1)) {
    const log = tables.LOG_Flights[0].map(() => '');
    log[0] = row[0];
    log[23] = '2026-11-06T14:00:00Z';
    tables.LOG_Flights.push(log);
  }
  const h = harness(tables);
  assert.equal(h.export(7).missions.length, 0);
  assert.deepEqual(h.writes, []);
});

test('source changes during export invalidate every mission', () => {
  const tables = fixture();
  const h = harness(tables);
  h.context.mutateSource = () => { tables.DB_Aircraft[1][4] += 10; };
  vm.runInContext(`
    var originalStartupReader = _getPilotStartupDataReadOnly_;
    _getPilotStartupDataReadOnly_ = function() {
      var result = originalStartupReader();
      mutateSource();
      return result;
    };
  `, h.context);
  const pack = h.export(7);
  assert.equal(pack.completeness.ready, false);
  assert.match(pack.completeness.errors.join(';'), /Source changed during export: DB_Aircraft/);
  assert.ok(pack.completeness.missionResults.every(result => !result.ready));
  assert.deepEqual(h.writes, []);
});

test('invalid coordinates, pilot source weight and empty multiplier values cannot claim readiness', () => {
  for (const damage of [
    tables => { tables.DB_Airports[1][2] = 'not coordinates'; },
    tables => { tables.DB_Pilots[1][5] = ''; },
    tables => { tables.Performance_Multipliers[1][2] = ''; },
    tables => { tables.Performance_Multipliers[0][7] = 'MISSING_HUMIDITY_FACTOR'; },
    tables => { for (const row of tables.Performance_Multipliers.slice(1)) row[19] = ''; }
  ]) {
    const tables = fixture(1);
    damage(tables);
    const h = harness(tables);
    assert.equal(h.export(7).completeness.ready, false);
    assert.deepEqual(h.writes, []);
  }
});

test('download interface exposes JSON only after response and distinguishes incomplete packages', async () => {
  const tools = fs.readFileSync(path.join(root, 'View_Tools.html'), 'utf8');
  const script = tools.slice(tools.indexOf("  var _toolsOfflineDownloadUrl = ''"), tools.indexOf('  var _toolsWpPreviewMap ='));
  for (const ready of [true, false]) {
    const elements = Object.fromEntries(['prepare', 'status', 'report', 'download', 'days'].map(id => [
      'tools-offline-' + id, {
        disabled: false, value: '14', style: {}, textContent: '', children: [],
        removeAttribute(name) { delete this[name]; }, appendChild(child) { this.children.push(child); }
      }
    ]));
    let response;
    let failure;
    let receivedDays;
    let blob;
    const runner = {
      withSuccessHandler(fn) { response = fn; return this; },
      withFailureHandler(fn) { failure = fn; return this; },
      exportOfflineMissionPackage(days) { receivedDays = days; }
    };
    const context = {
      document: { getElementById: id => elements[id], createElement: () => ({ textContent: '' }) },
      google: { script: { run: runner } }, Blob,
      URL: { createObjectURL(value) { blob = value; return 'blob:package'; }, revokeObjectURL() {} }
    };
    context.window = context;
    vm.createContext(context);
    vm.runInContext(script, context);
    context.toolsPrepareOfflineDownload();
    assert.equal(receivedDays, 14);
    assert.equal(elements['tools-offline-download'].style.display, 'none');
    response({
      kind: 'flightapp.offline-package', packageId: 'test',
      window: { fromDate: '2026-11-06', days: 14 },
      completeness: { ready, eligibleCount: 2, errors: [], missionResults: [
        { missionId: '<mission>', ready, missing: ['Missing envelope'] }
      ] }
    });
    assert.equal(elements['tools-offline-prepare'].disabled, false);
    assert.equal(elements['tools-offline-download'].href, 'blob:package');
    assert.equal(JSON.parse(await blob.text()).completeness.ready, ready);
    assert.match(elements['tools-offline-status'].textContent, ready ? /data complete/ : /NOT READY/);
    assert.match(elements['tools-offline-report'].children[0].textContent, /<mission>/);
    context.toolsPrepareOfflineDownload();
    failure(new Error('Authentication expired'));
    assert.match(elements['tools-offline-status'].textContent, /Authentication expired/);
    assert.equal(elements['tools-offline-download'].style.display, 'none');
    assert.equal(elements['tools-offline-prepare'].disabled, false);
  }
});

test('changed backend files and Tools script parse; production dashboard retains its limit', () => {
  for (const file of backendFiles) new vm.Script(fs.readFileSync(path.join(root, file), 'utf8'), { filename: file });
  const tools = fs.readFileSync(path.join(root, 'View_Tools.html'), 'utf8');
  for (const block of tools.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)) new vm.Script(block[1], { filename: 'View_Tools.html' });
  assert.match(fs.readFileSync(path.join(root, 'Backend_Dispatch.js'), 'utf8'), /\.reverse\(\)\.slice\(0, 15\)/);
});
