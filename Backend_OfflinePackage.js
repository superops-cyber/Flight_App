function exportOfflineMissionPackage(windowDays) {
  var days = Number(windowDays);
  if (days !== 7 && days !== 14) throw new Error('Offline download window must be 7 or 14 days.');
  var email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) throw new Error('Offline download requires an authenticated organizational account.');
  var staffRows = _toolsLoadStaffRows_();
  var staff = staffRows.filter(function(person) { return person.active && person.email === email; });
  if (staff.length !== 1) throw new Error('Offline download requires one active staff record for your account.');
  var person = staff[0];
  var canCoordinate = _toolsSupervisorLike_(person);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeZone = ss.getSpreadsheetTimeZone();
  var now = new Date();
  var fromDate = Utilities.formatDate(now, timeZone, 'yyyy-MM-dd');
  var end = new Date(fromDate + 'T12:00:00Z');
  end.setUTCDate(end.getUTCDate() + days);
  var toDate = Utilities.formatDate(end, 'UTC', 'yyyy-MM-dd');
  var tables = {};
  var issues = [];

  function read(name, columns) {
    try {
      var table = _offlinePackageTable_(ss, name, columns);
      tables[name] = table;
      return table;
    } catch (error) {
      issues.push(error.message);
      return null;
    }
  }

  var dispatch = read(APP_SHEETS.DISPATCH, ['FLIGHT_ID', 'MISSION_ID', 'DATE', 'AIRCRAFT', 'PILOT', 'COPILOT', 'ROUTE', 'RAW_DATA', 'STATUS']);
  if (!dispatch) throw new Error(issues.join('; '));
  // Existing mission/W&B readers use positional Dispatch columns.
  ['FLIGHT_ID', 'MISSION_ID', 'DATE', 'AIRCRAFT', 'PILOT', 'COPILOT', 'ROUTE', 'RAW_DATA', 'STATUS'].forEach(function(key) {
    if (dispatch.headers[DISPATCH_COL[key]] !== key) throw new Error('DB_Dispatch: unexpected column position for ' + key);
  });
  var groups = Object.create(null);
  dispatch.rows.forEach(function(row) {
    var id = String(row[DISPATCH_COL.MISSION_ID] || '').trim();
    if (id) {
      if (!groups[id]) groups[id] = [];
      groups[id].push(row);
    }
  });
  var candidates = Object.keys(groups).filter(function(id) {
    var rows = groups[id];
    var done = ['FLOWN', 'COMPLETE', 'COMPLETED'];
    var loadable = ['APPROVED', 'IN-FLIGHT', 'IN FLIGHT', 'DEPARTED'].concat(done);
    if (!rows.every(function(row) { return loadable.indexOf(String(row[DISPATCH_COL.STATUS] || '').trim().toUpperCase()) >= 0; })) return false;
    if (!rows.some(function(row) {
      var date = _offlinePackageDate_(row[DISPATCH_COL.DATE], timeZone);
      return date && date >= fromDate && date < toDate
        && done.indexOf(String(row[DISPATCH_COL.STATUS] || '').trim().toUpperCase()) < 0;
    })) return false;
    return canCoordinate || rows.every(function(row) {
      return [row[DISPATCH_COL.PILOT], row[DISPATCH_COL.COPILOT]].some(function(crew) {
        var value = String(crew || '').trim().toUpperCase();
        return value && [person.email, person.staffId, person.staffName].some(function(identity) {
          return identity && String(identity).trim().toUpperCase() === value;
        });
      });
    });
  }).sort();
  var flightIdCounts = Object.create(null);
  dispatch.rows.forEach(function(row) {
    var fid = String(row[DISPATCH_COL.FLIGHT_ID] || '').trim();
    if (fid) flightIdCounts[fid] = (flightIdCounts[fid] || 0) + 1;
  });

  var specs = [
    [APP_SHEETS.AIRCRAFT, ['REGISTRATION', 'AIRCRAFT_TYPE', 'EMPTY_WEIGHT', 'EMPTY_ARM', 'MTOW', 'CURRENT_TACH',
      'MAX_FUEL', 'NUM_TANKS', 'TANK_NAMES', 'PILOT_SEAT_KGS', 'MID_SEAT_KGS', 'AFT_SEAT_KGS', 'TECH_STATUS', 'LICENSE_REQUIRED']],
    [APP_SHEETS.PILOTS, ['NAME|PILOT_NAME', 'EMAIL|PILOT_EMAIL', 'WEIGHT_KGS', 'MEDICAL_EXPIRY']],
    [APP_SHEETS.AIRFRAMES, ['AIRCRAFT_TYPE', 'STATION_NAME', 'ARM']],
    [APP_SHEETS.ENVELOPES, ['AIRCRAFT_TYPE', 'POINT_SEQUENCE', 'CG_ARM_X', 'WEIGHT_Y']],
    [APP_SHEETS.AIRPORTS, ['ICAO', 'LATITUDE', 'LONGITUDE', 'RWY_IDENT|RWY|RUNWAY|RUNWAY_DESIGNATOR',
      'LENGTH_OFFICIAL|LENGTH_METERS|LENGTH_M', 'SURFACE_ACTUAL|SURFACE_OFFICIAL|SURFACE']],
    [APP_SHEETS.ROUTES, ['ORIGIN', 'DESTINATION', 'WAYPOINT_LIST']],
    [APP_SHEETS.WAYPOINTS, ['WP_ID', 'LATITUDE', 'LONGITUDE']],
    [APP_SHEETS.PASSENGERS, ['PASSENGER_NAME', 'WEIGHT_KG|WEIGHT_KGS']],
    [APP_SHEETS.CHECKS, ['PILOT_NAME|PILOT', 'AUTH_DESTINATIONS|AUTHORIZED_DESTINATIONS|DESTINATIONS|DESTINATION|AUTH_DESTINATION|ICAO|AIRPORT_ICAO|DESTINATION_ICAO']],
    [APP_SHEETS.FUNDS, ['NAME|FUND_NAME', 'CURRENT_BALANCE']],
    [APP_SHEETS.SYLLABUS, ['TRAINING_CODE', 'REQUIRED_HOURS', 'REQUIRED_FUEL']],
    [APP_SHEETS.FUEL_CACHES, ['ICAO', 'CURRENT_QTY', 'FUEL_TYPE']],
    [APP_SHEETS.FUEL_LOGS, ['FLIGHT_ID', 'ICAO', 'CHANGE_QTY', 'TYPE', 'VERIFIED']],
    [APP_SHEETS.LOG_FLIGHTS, ['FLIGHT_ID', 'ACTUAL_LOAD_JSON', 'BRAKES_RELEASE', 'AIRBORNE', 'LANDED']],
    [APP_SHEETS.SCHED_CONFIG, ['CONFIG_KEY', 'CONFIG_VALUE']],
    [APP_SHEETS.MAINT_ASSIGNMENTS, ['ASSIGNMENT_ID']],
    [APP_SHEETS.MAINT_LOG, ['ASSIGNMENT_ID']],
    ['Aircraft_Roll_Numbers', ['AIRCRAFT_TYPE', 'WEIGHT_KG|WEIGHT', 'TO_ROLL_M|TO_ROLL', 'LDG_ROLL_M|LDG_ROLL']],
    ['Performance_Multipliers', ['AIRCRAFT_TYPE', 'FLAP_SETTING', 'TO_MULTIPLIER', 'LDG_MULTIPLIER',
      'DA_FT|DA', 'WIND_KTS|WIND', 'HUMIDITY', 'HUMIDITY_FACTOR', 'FLAP_FACTOR', 'SLOPE', 'SURFACE',
      'TAKEOFF_HEAD', 'TAKEOFF_TAIL', 'LANDING_HEAD', 'LANDING_TAIL',
      'TAKEOFF_UP', 'TAKEOFF_DOWN', 'LANDING_UP', 'LANDING_DOWN',
      'TAKEOFF_DRY', 'TAKEOFF_WET', 'LANDING_DRY', 'LANDING_WET']]
  ];
  specs.forEach(function(spec) { read(spec[0], spec[1]); });
  // These existing readers still use fixed column indexes.
  [
    [APP_SHEETS.FUEL_LOGS, FUEL_LOG_COL, ['ICAO', 'CHANGE_QTY', 'TYPE', 'VERIFIED', 'FLIGHT_ID']],
    [APP_SHEETS.LOG_FLIGHTS, LOG_FLIGHT_COL, ['FLIGHT_ID', 'BRAKES_RELEASE', 'LANDED', 'BRAKES_APPLIED']]
  ].forEach(function(spec) {
    var table = tables[spec[0]];
    if (!table) return;
    spec[2].forEach(function(key) {
      var actual = table.headers[spec[1][key]];
      if (actual !== key && !(key === 'BRAKES_APPLIED' && actual === 'ON_BLOCKS')) {
        issues.push(spec[0] + ': unexpected column position for ' + key);
      }
    });
  });
  if (tables[APP_SHEETS.LOG_FLIGHTS]) {
    var completedLegs = Object.create(null);
    tables[APP_SHEETS.LOG_FLIGHTS].objects.forEach(function(row) {
      if (row.LANDED || row.ON_BLOCKS || row.BRAKES_APPLIED) completedLegs[String(row.FLIGHT_ID || '').trim()] = true;
    });
    candidates = candidates.filter(function(id) {
      return groups[id].some(function(row) {
        return !completedLegs[String(row[DISPATCH_COL.FLIGHT_ID] || '').trim()]
          && !/^(FLOWN|COMPLETE|COMPLETED)$/.test(String(row[DISPATCH_COL.STATUS] || '').trim().toUpperCase());
      });
    });
  }
  var startup = null;
  var settings = null;
  if (!issues.length) {
    try {
      startup = _getPilotStartupDataReadOnly_();
      if (!startup || startup.error) throw new Error(startup && startup.error || 'Startup data unavailable');
      var config = _toolsDutyReadConfigMap_();
      settings = {};
      Object.keys(config).filter(function(key) {
        return /^PASSENGER_TERMS_/.test(key) || key === 'DUTY_LANDING_THRESHOLD_KTS' || key === 'DUTY_TOUCH_AND_GO_CONFIRM_KTS';
      }).forEach(function(key) { settings[key] = config[key]; });
    } catch (error) {
      startup = null;
      settings = null;
      issues.push(error.message);
    }
  }

  var pack = {
    kind: 'flightapp.offline-package',
    schemaVersion: 1,
    packageId: Utilities.getUuid(),
    exportedAt: now.toISOString(),
    exportedFor: email,
    window: { days: days, fromDate: fromDate, toDateExclusive: toDate, timeZone: timeZone },
    missions: [],
    references: {
      startup: startup, operationalSettings: settings,
      aircraftRecords: Object.create(null), airframeStations: Object.create(null),
      envelopes: Object.create(null), wbConfigurations: Object.create(null),
      performanceSetups: Object.create(null), performanceReferences: Object.create(null)
    },
    completeness: { ready: false, eligibleCount: candidates.length, missionResults: [], errors: issues },
    limitations: ['Application assets, live tracking, remote charts, photos and map tiles are not embedded in this JSON file.']
  };
  var refs = pack.references;
  var wbConfigurations = Object.create(null);
  var performanceErrors = Object.create(null);
  var envelopeErrors = Object.create(null);
  var performance = null;
  if (startup) {
    var crewNames = Object.create(null);
    var aircraftRegs = Object.create(null);
    var passengerNames = Object.create(null);
    candidates.forEach(function(id) {
      groups[id].forEach(function(row) {
        aircraftRegs[String(row[DISPATCH_COL.AIRCRAFT] || '').trim().toUpperCase()] = true;
        [row[DISPATCH_COL.PILOT], row[DISPATCH_COL.COPILOT]].forEach(function(crew) {
          var record = _toolsFindStaffByEmailOrId_(staffRows, String(crew || ''), '');
          crewNames[String(record && record.staffName || crew || '').trim().toUpperCase()] = true;
        });
        try {
          var raw = JSON.parse(String(row[DISPATCH_COL.RAW_DATA] || '{}'));
          var leg = Array.isArray(raw.legs) ? raw.legs[0] : raw;
          (leg && leg.pax || []).forEach(function(pax) { passengerNames[String(pax.name || '').trim().toUpperCase()] = true; });
        } catch (error) {
          // The affected mission receives the malformed-data error below.
        }
      });
    });
    startup.aircraft = startup.aircraft.filter(function(aircraft) { return !!aircraftRegs[String(aircraft.reg).toUpperCase()]; });
    startup.pilots = startup.pilots.filter(function(pilot) { return !!crewNames[String(pilot.name).toUpperCase()]; });
    startup.passengers = startup.passengers.filter(function(pax) { return !!passengerNames[String(pax.name).toUpperCase()]; });
    Object.keys(startup.pilotDestinationChecks).forEach(function(name) {
      if (!crewNames[String(name).toUpperCase()]) delete startup.pilotDestinationChecks[name];
    });
    tables[APP_SHEETS.AIRCRAFT].objects.forEach(function(row) {
      var reg = String(row.REGISTRATION || '').trim().toUpperCase();
      if (aircraftRegs[reg]) refs.aircraftRecords[reg] = row;
    });
    performance = {
      baseRows: tables.Aircraft_Roll_Numbers.objects,
      perfRows: tables.Performance_Multipliers.objects
    };
    refs.performanceReferences.global = performance;
  }

  candidates.forEach(function(id) {
    var missing = issues.slice();
    var rows = groups[id];
    var entry = { id: id, sourceRevision: _offlinePackageDigest_(rows), summary: null, detail: null, legs: {} };
    function check(ok, message) { if (!ok) missing.push(message); }
    function attempt(label, action) {
      try { return action(); } catch (error) { missing.push(label + ': ' + error.message); return null; }
    }
    if (!issues.length) {
      rows.forEach(function(row) {
        check(!!_offlinePackageDate_(row[DISPATCH_COL.DATE], timeZone), 'Invalid Dispatch date: ' + row[DISPATCH_COL.FLIGHT_ID]);
        attempt('Dispatch payload ' + row[DISPATCH_COL.FLIGHT_ID], function() {
          var raw = JSON.parse(String(row[DISPATCH_COL.RAW_DATA] || ''));
          if (!raw || typeof raw !== 'object' || Array.isArray(raw)) throw new Error('Expected Dispatch JSON object');
        });
      });
      var mission = attempt('Mission ' + id, function() { return getMissionById(id); });
      entry.detail = mission;
      check(mission && Array.isArray(mission.legs) && mission.legs.length === rows.length, 'Missing mission details or flight legs');
      if (mission && Array.isArray(mission.legs)) {
        var ownLegIds = mission.legs.map(function(leg) { return String(leg.flightLegId || '').trim(); });
        mission.actualFuelLogs = (mission.actualFuelLogs || []).filter(function(log) {
          return ownLegIds.indexOf(String(log.flightLegId).trim()) >= 0;
        });
        var reg = String(mission.acft || '').trim().toUpperCase();
        var aircraft = refs.aircraftRecords[reg];
        var type = String(aircraft && aircraft.AIRCRAFT_TYPE || '').trim();
        check(!!aircraft && !!type, 'Missing aircraft configuration: ' + reg);
        check(aircraft && _offlinePackagePositive_(aircraft.EMPTY_WEIGHT) && _offlinePackageFinite_(aircraft.EMPTY_ARM)
          && _offlinePackagePositive_(aircraft.MTOW), 'Invalid aircraft empty weight, arm or MTOW: ' + reg);
        check(aircraft && _offlinePackagePositive_(aircraft.MAX_FUEL) && _offlinePackagePositive_(aircraft.NUM_TANKS)
          && String(aircraft.TANK_NAMES || '').trim(), 'Missing aircraft fuel tank configuration: ' + reg);
        check(aircraft && ['PILOT_SEAT_KGS', 'MID_SEAT_KGS', 'AFT_SEAT_KGS'].every(function(key) {
          return _offlinePackageFinite_(aircraft[key]) && Number(aircraft[key]) >= 0;
        }) && String(aircraft.LICENSE_REQUIRED || '').trim(), 'Missing seat weights or required license: ' + reg);
        check(startup.pilots.some(function(pilot) {
          return String(pilot.name).trim().toUpperCase() === String(mission.pilot).trim().toUpperCase()
            && _offlinePackagePositive_(pilot.weight) && pilot.Medical_Expiry;
        }), 'Missing pilot weight/medical reference: ' + mission.pilot);
        check(tables[APP_SHEETS.PILOTS].objects.some(function(pilot) {
          return String(pilot.NAME || pilot.PILOT_NAME || '').trim().toUpperCase() === String(mission.pilot).trim().toUpperCase()
            && _offlinePackagePositive_(pilot.WEIGHT_KGS);
        }), 'Missing source pilot weight: ' + mission.pilot);
        if (type && !refs.airframeStations[type]) {
          refs.airframeStations[type] = tables[APP_SHEETS.AIRFRAMES].objects.filter(function(row) {
            return String(row.AIRCRAFT_TYPE || '').trim().toUpperCase() === type.toUpperCase();
          });
        }
        var stations = refs.airframeStations[type] || [];
        check(stations.length && stations.every(function(row) { return _offlinePackageFinite_(row.ARM); }), 'Missing/invalid airframe stations: ' + type);
        ['PILOT|FRONT|FWD', 'MID|MIDDLE|ROW2|PASS', 'AFT|REAR|ROW3', 'CARGO|BAG|FREIGHT', 'FUEL|TANK'].forEach(function(pattern) {
          check(stations.some(function(row) { return new RegExp(pattern).test(String(row.STATION_NAME).toUpperCase()); }), 'Missing airframe station ' + pattern + ': ' + type);
        });
        var sourceEnvelope = tables[APP_SHEETS.ENVELOPES].objects.filter(function(row) {
          return String(row.AIRCRAFT_TYPE || '').trim().toUpperCase() === type.toUpperCase();
        }).sort(function(a, b) {
          return Number(a.POINT_SEQUENCE) - Number(b.POINT_SEQUENCE);
        });
        check(sourceEnvelope.every(function(point) { return _offlinePackageFinite_(point.POINT_SEQUENCE); })
          && _offlinePackageEnvelopeValid_(sourceEnvelope.map(function(point) {
            return { CG_Arm_X: point.CG_ARM_X, Weight_Y: point.WEIGHT_Y };
          })), 'Missing/invalid source envelope coordinates: ' + type);
        if (type && !refs.envelopes[type] && !envelopeErrors[type]) {
          try {
            var envelope = getWbEnvelopeByAircraft(reg);
            if (!_offlinePackageEnvelopeValid_(envelope.envelopeData)) throw new Error('Invalid CG polygon');
            refs.envelopes[type] = envelope.envelopeData;
          } catch (error) { envelopeErrors[type] = error.message; }
        }
        check(!!refs.envelopes[type], 'Missing envelope ' + type + ': ' + (envelopeErrors[type] || 'unavailable'));
        var typeKey = _perfAircraftTypeKey_(type);
        var baseRows = performance.baseRows.filter(function(row) { return _perfAircraftTypeKey_(row.AIRCRAFT_TYPE) === typeKey; });
        check(baseRows.length && baseRows.every(function(row) {
          return _offlinePackagePositive_(_perfValue(row, ['WEIGHT_KG', 'WEIGHT'], ''))
            && _offlinePackagePositive_(_perfValue(row, ['TO_ROLL_M', 'TO_ROLL'], ''))
            && _offlinePackagePositive_(_perfValue(row, ['LDG_ROLL_M', 'LDG_ROLL'], ''));
        }), 'Missing/invalid performance base rows: ' + type);
        check(performance.perfRows.some(function(row) {
          return _perfAircraftTypeKey_(row.AIRCRAFT_TYPE) === typeKey
            && _offlinePackageFinite_(row.FLAP_SETTING)
            && _offlinePackagePositive_(row.TO_MULTIPLIER) && _offlinePackagePositive_(row.LDG_MULTIPLIER);
        }),
          'Missing performance multipliers for aircraft type: ' + type);
        [
          [['DA_FT', 'DA'], ['TO_MULTIPLIER', 'LDG_MULTIPLIER']],
          [['WIND_KTS', 'WIND'], ['TAKEOFF_HEAD', 'TAKEOFF_TAIL', 'LANDING_HEAD', 'LANDING_TAIL']],
          [['HUMIDITY'], ['HUMIDITY_FACTOR']],
          [['FLAP_SETTING'], ['FLAP_FACTOR']],
          [['SLOPE'], ['TAKEOFF_UP', 'TAKEOFF_DOWN', 'LANDING_UP', 'LANDING_DOWN']]
        ].forEach(function(group) {
          group[1].forEach(function(valueColumn) {
            check(performance.perfRows.some(function(row) {
              return _offlinePackageFinite_(_perfValue(row, group[0], '')) && _offlinePackagePositive_(row[valueColumn]);
            }), 'Missing calculation reference: ' + group[0].join('/') + ' -> ' + valueColumn);
          });
        });
        var legIds = Object.create(null);
        mission.legs.forEach(function(leg) {
          var fid = String(leg.flightLegId || '').trim();
          check(fid && !legIds[fid] && flightIdCounts[fid] === 1
            && rows.some(function(row) { return String(row[DISPATCH_COL.FLIGHT_ID]).trim() === fid; }), 'Missing, duplicate or mismatched flight leg: ' + fid);
          legIds[fid] = true;
          var dispatchLeg = rows.find(function(row) { return String(row[DISPATCH_COL.FLIGHT_ID]).trim() === fid; });
          var dispatchStatus = String(dispatchLeg && dispatchLeg[DISPATCH_COL.STATUS] || '').trim().toUpperCase();
          var legData = { dispatchStatus: dispatchStatus, wb: null, wbConfigurationRef: '', envelopeRef: type, route: null, flightLog: null, performanceSetupRefs: [] };
          entry.legs[fid] = legData;
          var wb = attempt('W&B ' + fid, function() { return initializeWB(fid); });
          if (wb) {
            check(wb.flightId === fid && String(wb.aircraft).toUpperCase() === reg, 'W&B aircraft/leg mismatch: ' + fid);
            check(Array.isArray(wb.items) && wb.items.length && wb.seats && Object.keys(wb.seats).length
              && _offlinePackagePositive_(wb.airframeData && wb.airframeData.Empty_Weight)
              && _offlinePackageFinite_(wb.fuelArm) && _offlinePackageEnvelopeValid_(wb.envelopeData), 'Incomplete W&B payload: ' + fid);
            check(wb.seatAssignments && wb.seatAssignments.pilot
              && _offlinePackagePositive_(wb.seatAssignments.pilot.occupiedWeight), 'Missing W&B pilot weight: ' + fid);
            var configuration = { airframeData: wb.airframeData, fuelArm: wb.fuelArm, cargoAreas: wb.cargoAreas };
            var configurationKey = JSON.stringify(configuration);
            if (!wbConfigurations[configurationKey]) {
              wbConfigurations[configurationKey] = 'wb-' + (Object.keys(refs.wbConfigurations).length + 1);
              refs.wbConfigurations[wbConfigurations[configurationKey]] = configuration;
            }
            legData.wbConfigurationRef = wbConfigurations[configurationKey];
            delete wb.airframeData;
            delete wb.fuelArm;
            delete wb.cargoAreas;
            delete wb.envelopeData;
            legData.wb = wb;
          }
          var route = attempt('Route ' + fid, function() { return getFlightRouteData(fid); });
          legData.route = route;
          check(route && route.from === leg.from && route.to === leg.to, 'Missing/mismatched route: ' + fid);
          (route && route.waypoints || []).forEach(function(point) {
            var key = String(point && typeof point === 'object' ? point.wp_id || point.name || point.id : point).trim().toUpperCase();
            check(startup.waypoints.some(function(wp) {
              return String(wp.wp_id).trim().toUpperCase() === key && _offlinePackageCoordinatesValid_(wp.lat, wp.lon);
            }) || startup.airports.some(function(airport) {
              return String(airport.icao).trim().toUpperCase() === key && _offlinePackageCoordinatesValid_(airport.lat, airport.lon);
            }), 'Missing/invalid route waypoint: ' + key);
          });
          legData.flightLog = attempt('Flight log ' + fid, function() { return getFlightLogData(fid); });
          if (legData.flightLog) {
            check(String(legData.flightLog.flightLegId).trim() === fid, 'Flight-log leg mismatch: ' + fid);
            attempt('Flight-log JSON ' + fid, function() {
              var load = JSON.parse(legData.flightLog.actualLoadJSON || '{}');
              if (!load || typeof load !== 'object' || Array.isArray(load)) throw new Error('Expected load JSON object');
            });
          } else {
            check(!tables[APP_SHEETS.LOG_FLIGHTS].objects.some(function(row) { return String(row.FLIGHT_ID).trim() === fid; })
              && dispatchStatus === 'APPROVED'
              && !/^(IN-FLIGHT|IN FLIGHT|DEPARTED|COMPLETE)$/.test(String(leg.logStatus).toUpperCase()), 'Missing existing flight-log baseline: ' + fid);
          }
          [leg.from, leg.to].forEach(function(icao) {
            var code = String(icao || '').trim().toUpperCase();
            var key = code + '|' + type;
            check(startup.airports.some(function(airport) {
              return String(airport.icao).trim().toUpperCase() === code && _offlinePackageCoordinatesValid_(airport.lat, airport.lon);
            }), 'Missing airport/coordinates: ' + code);
            if (!refs.performanceSetups[key] && !performanceErrors[key]) {
              try {
                var setup = getPerformanceSetup(code, type);
                if (!setup || !setup.runways || !setup.runways.length || !setup.runways.every(function(runway) {
                  return runway.rwyIdent && _offlinePackagePositive_(runway.length) && runway.surface;
                })) throw new Error('Missing/invalid runway setup');
                setup.runways.forEach(function(runway) {
                  var surface = _perfNorm(runway.surface);
                  var translations = setup.surfaceTranslator || {};
                  var translationKey = Object.keys(translations).find(function(key) {
                    return _perfNorm(key) === surface;
                  });
                  if (translationKey) surface = _perfNorm(translations[translationKey]);
                  ['TAKEOFF_DRY', 'TAKEOFF_WET', 'LANDING_DRY', 'LANDING_WET'].forEach(function(column) {
                    if (!performance.perfRows.some(function(row) {
                      var key = _perfNorm(row.SURFACE);
                      return key && (key === surface || key.indexOf(surface) >= 0 || surface.indexOf(key) >= 0)
                        && _offlinePackagePositive_(row[column]);
                    })) throw new Error('Missing surface reference ' + runway.surface + ' -> ' + column);
                  });
                });
                delete setup.calcReference;
                setup.performanceReference = 'global';
                refs.performanceSetups[key] = setup;
              } catch (error) { performanceErrors[key] = error.message; }
            }
            check(!!refs.performanceSetups[key], 'Missing performance setup ' + key + ': ' + (performanceErrors[key] || 'unavailable'));
            if (legData.performanceSetupRefs.indexOf(key) < 0) legData.performanceSetupRefs.push(key);
          });
        });
        var flown = mission.legs.filter(function(leg) {
          return leg.logStatus === 'COMPLETE' || /^(FLOWN|COMPLETE|COMPLETED)$/.test(entry.legs[leg.flightLegId].dispatchStatus);
        }).length;
        var status = flown > 0 ? flown + ' OF ' + mission.legs.length + ' LEGS' : mission.status;
        entry.summary = {
          id: id, date: mission.date, acft: reg, pilot: mission.pilot, status: status,
          from: mission.legs[0] && mission.legs[0].from,
          to: mission.legs.length && mission.legs[mission.legs.length - 1].to,
          route: mission.legs.map(function(leg) { return leg.route; }).join(' / '),
          legCount: mission.legs.length,
          legsFlown: flown
        };
      }
    }
    pack.missions.push(entry);
    pack.completeness.missionResults.push({ missionId: id, ready: missing.length === 0, missing: Array.from(new Set(missing)) });
  });
  // A package must not combine records from different revisions of Sheets.
  Object.keys(tables).forEach(function(name) {
    var current = ss.getSheetByName(name);
    if (!current || JSON.stringify(current.getDataRange().getValues()) !== tables[name].snapshot) {
      pack.completeness.errors.push('Source changed during export: ' + name + '. Download again.');
    }
  });
  if (pack.completeness.errors.length) {
    pack.completeness.missionResults.forEach(function(result) {
      result.ready = false;
      result.missing = Array.from(new Set(result.missing.concat(pack.completeness.errors)));
    });
  }
  pack.completeness.ready = candidates.length > 0 && !pack.completeness.errors.length
    && pack.completeness.missionResults.every(function(result) { return result.ready; });
  if (!candidates.length) pack.completeness.errors.push('No eligible missions in this window for your account.');
  pack.sourceRevision = _offlinePackageDigest_(Object.keys(tables).map(function(name) { return [name, tables[name].snapshot]; }));
  return _offlinePackageJsonSafe_(pack);
}

function _offlinePackageTable_(ss, name, requiredColumns) {
  var values = getRequiredSheet_(ss, name, 'Offline download').getDataRange().getValues();
  var headers = (values[0] || []).map(function(value) { return String(value || '').trim().toUpperCase().replace(/\s+/g, '_'); });
  var missing = requiredColumns.filter(function(group) {
    return !group.split('|').some(function(key) { return headers.indexOf(key) >= 0; });
  });
  if (missing.length) throw new Error(name + ': missing columns ' + missing.join(', '));
  return {
    headers: headers, rows: values.slice(1), snapshot: JSON.stringify(values),
    objects: values.slice(1).map(function(row) {
      var result = {};
      headers.forEach(function(key, index) { result[key] = row[index]; });
      return result;
    })
  };
}

function _offlinePackageDate_(value, timeZone) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
  var text = String(value || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return '';
  var date = new Date(text + 'T12:00:00Z');
  return !isNaN(date.getTime()) && date.toISOString().slice(0, 10) === text ? text : '';
}

function _offlinePackageFinite_(value) {
  return value !== null && value !== undefined && String(value).trim() !== '' && isFinite(Number(value));
}

function _offlinePackagePositive_(value) {
  return _offlinePackageFinite_(value) && Number(value) > 0;
}

function _offlinePackageCoordinatesValid_(lat, lon) {
  function parse(value) {
    var text = String(value == null ? '' : value).trim().replace(/(\d),(\d)/g, '$1.$2');
    if (/^-?\d+(\.\d+)?$/.test(text)) return Number(text);
    var parts = text.replace(/[\u00b0'"]/g, ' ').trim().match(/^([NSEW])?\s*(\d+)\s+(\d+(?:\.\d+)?)(?:\s+(\d+(?:\.\d+)?))?\s*([NSEW])?$/i);
    if (!parts || Number(parts[3]) >= 60 || Number(parts[4] || 0) >= 60) return NaN;
    var cardinal = String(parts[1] || parts[5] || '').toUpperCase();
    return (Number(parts[2]) + Number(parts[3]) / 60 + Number(parts[4] || 0) / 3600) * (/[SW]/.test(cardinal) ? -1 : 1);
  }
  var x = parse(lat);
  var y = parse(lon);
  return isFinite(x) && isFinite(y) && Math.abs(x) <= 90 && Math.abs(y) <= 180;
}

function _offlinePackageEnvelopeValid_(points) {
  if (!Array.isArray(points) || points.length < 3 || !points.every(function(point) {
    return _offlinePackageFinite_(point.CG_Arm_X) && _offlinePackagePositive_(point.Weight_Y);
  })) return false;
  var area = points.reduce(function(sum, point, index) {
    var next = points[(index + 1) % points.length];
    return sum + Number(point.CG_Arm_X) * Number(next.Weight_Y) - Number(next.CG_Arm_X) * Number(point.Weight_Y);
  }, 0);
  return Math.abs(area) > 0;
}

function _offlinePackageDigest_(value) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, JSON.stringify(value), Utilities.Charset.UTF_8
  ));
}

function _offlinePackageJsonSafe_(value) {
  return JSON.parse(JSON.stringify(value, function(key, item) {
    if (typeof item === 'number' && !isFinite(item)) throw new Error('Offline package contains a non-finite number: ' + key);
    if (item === undefined) return null;
    return item;
  }));
}
