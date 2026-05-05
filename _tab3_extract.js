

(function() {
  var modal = document.getElementById('wb-keypad-modal');
  if (modal && modal.parentElement !== document.body) document.body.appendChild(modal);
})();

// ============================================
// WEIGHT & BALANCE - DATA MODEL & CALCULATIONS
// ============================================

window.wbData = {
  flightId: '',
  aircraft: '',
  route: '',
  pilot: '',
  pilotWeightEstimated: false,
  pilotWeightEstimateReason: '',
  date: '',
  airframeData: null,
  envelopeData: [],
  items: [],  // { name, plannedWeight, actualWeight, arm, type }
  seats: {},  // { seatName: { weight, arm, enabled } }
  fuel: 0,    // Actual fuel loaded
  fuelArm: 0,
  waitingPassengers: [], // Passengers not currently assigned to a seat
};

function wbNum_(value, fallback) {
  const n = parseFloat(value);
  return isNaN(n) ? (fallback || 0) : n;
}

function wbNormName_(value) {
  return String(value || '').trim();
}

function wbFirstName_(value) {
  const full = wbNormName_(value);
  if (!full) return '';
  return full.split(/\s+/)[0] || full;
}

function wbMissionIdFromFlightId_(flightId) {
  const fid = String(flightId || '').trim();
  if (!fid) return '';
  const parts = fid.split('-');
  if (parts.length >= 3) return parts.slice(0, 2).join('-');
  return fid;
}

function wbReadCachedMissionByFlightId_(flightId) {
  try {
    const missionId = wbMissionIdFromFlightId_(flightId);
    if (!missionId) return null;
    const raw = localStorage.getItem('mba_cache_mission_' + missionId);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && parsed.value ? parsed.value : null;
  } catch (e) {
    return null;
  }
}

function wbFindMissionLegByFlightId_(mission, flightId) {
  const legs = mission && Array.isArray(mission.legs) ? mission.legs : [];
  if (!legs.length) return null;
  const fid = String(flightId || '').trim();
  if (!fid) return legs[0];
  return legs.find(function(leg) {
    return String(leg && leg.flightLegId || '').trim() === fid;
  }) || legs[0];
}

function wbFindAircraft_(reg) {
  const list = (window.appData && Array.isArray(window.appData.aircraft)) ? window.appData.aircraft : [];
  const key = String(reg || '').trim().toUpperCase();
  return list.find(function(item) {
    return String(item && item.reg || '').trim().toUpperCase() === key;
  }) || null;
}

function wbIsPilotTbd_(pilotName) {
  const key = wbNormName_(pilotName).toUpperCase();
  return !key || key === 'PILOT TBD' || key === 'TBD' || key === 'UNASSIGNED';
}

function wbAveragePilotWeight_() {
  const list = (window.appData && Array.isArray(window.appData.pilots)) ? window.appData.pilots : [];
  const weights = list
    .map(function(item) { return wbNum_(item && item.weight, 0); })
    .filter(function(w) { return w > 0; });
  if (!weights.length) return 90;
  const sum = weights.reduce(function(acc, val) { return acc + val; }, 0);
  return Math.round((sum / weights.length) * 10) / 10;
}

function wbResolvePilotWeight_(pilotName) {
  const avgWeight = wbAveragePilotWeight_();
  if (wbIsPilotTbd_(pilotName)) {
    return {
      weight: avgWeight,
      estimated: true,
      reason: 'Pilot not assigned. Using average pilot weight.'
    };
  }

  const list = (window.appData && Array.isArray(window.appData.pilots)) ? window.appData.pilots : [];
  const key = wbNormName_(pilotName).toUpperCase();
  const hit = list.find(function(item) {
    return wbNormName_(item && item.name).toUpperCase() === key;
  });
  if (hit) {
    return {
      weight: wbNum_(hit && hit.weight, avgWeight),
      estimated: false,
      reason: ''
    };
  }

  return {
    weight: avgWeight,
    estimated: true,
    reason: 'Pilot weight unavailable. Using average pilot weight.'
  };
}

function wbFindPilotWeight_(pilotName) {
  return wbResolvePilotWeight_(pilotName).weight;
}

function wbShowPilotWeightEstimateInfo_() {
  const msg = String(window.wbData && window.wbData.pilotWeightEstimateReason || 'Estimated pilot weight used for this calculation.');
  if (window.M && typeof M.toast === 'function') {
    M.toast({ html: msg, classes: 'orange darken-3', displayLength: 4200 });
    return;
  }
  if (window.alert) window.alert(msg);
}

function wbReadCachedDropdownData_() {
  try {
    const raw = localStorage.getItem('mba_cache_dropdown_data_v1');
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return parsed && parsed.value && typeof parsed.value === 'object' ? parsed.value : {};
  } catch (e) {
    return {};
  }
}

function wbDropdownSource_() {
  const live = (window.appData && typeof window.appData === 'object') ? window.appData : {};
  const cached = wbReadCachedDropdownData_();
  return {
    passengers: Array.isArray(live.passengers) && live.passengers.length ? live.passengers : (Array.isArray(cached.passengers) ? cached.passengers : []),
    funds: Array.isArray(live.funds) && live.funds.length ? live.funds : (Array.isArray(cached.funds) ? cached.funds : []),
    rates: Array.isArray(live.rates) && live.rates.length ? live.rates : (Array.isArray(cached.rates) ? cached.rates : [])
  };
}

function wbMissionPassengers_() {
  const mission = (window.currentBriefingMission && Array.isArray(window.currentBriefingMission.legs)) ? window.currentBriefingMission : wbReadCachedMissionByFlightId_(window.wbData && window.wbData.flightId);
  const seen = {};
  const list = [];
  const legs = mission && Array.isArray(mission.legs) ? mission.legs : [];
  legs.forEach(function(leg) {
    const pax = Array.isArray(leg && leg.pax) ? leg.pax : [];
    pax.forEach(function(p) {
      const name = wbNormName_(p && p.name);
      if (!name || name.toUpperCase() === 'FREIGHT' || seen[name.toUpperCase()]) return;
      seen[name.toUpperCase()] = true;
      list.push({
        name: name,
        weight: wbNum_(p && (p.actualWeight != null ? p.actualWeight : p.weight), 80),
        fund: String(p && p.fund || ''),
        chargeRate: String(p && p.chargeRate || '')
      });
    });
  });
  return list;
}

function wbEnsureInputId_(inputEl) {
  if (!inputEl) return '';
  if (inputEl.id) return inputEl.id;
  var autoId = 'wb-num-' + Date.now() + '-' + Math.floor(Math.random() * 10000);
  inputEl.id = autoId;
  return autoId;
}

window.wbKeypadOpen_ = function(inputEl, label, opts) {
  if (!inputEl) return;

  var safeOpts = opts || {};
  window._wbKeypadEl = inputEl;
  window._wbKeypadOpts = safeOpts;

  var titleEl = document.getElementById('wb-keypad-title');
  var displayEl = document.getElementById('wb-keypad-display');
  var modalEl = document.getElementById('wb-keypad-modal');
  var dotRow = document.getElementById('wb-keypad-dot-row');
  if (!titleEl || !displayEl || !modalEl) return;

  titleEl.textContent = label || 'Entry';
  displayEl.textContent = '';
  if (dotRow) dotRow.style.display = safeOpts.decimal ? '' : 'none';
  modalEl.style.display = 'flex';
};

window.wbKeypadPress_ = function(val) {
  var displayEl = document.getElementById('wb-keypad-display');
  var opts = window._wbKeypadOpts || {};
  if (!displayEl) return;

  var cur = displayEl.textContent || '';
  if (val === 'BKSP') {
    displayEl.textContent = cur.slice(0, -1);
    return;
  }

  if (val === 'C') {
    displayEl.textContent = '';
    return;
  }

  if (val === '.') {
    if (opts.decimal && cur.indexOf('.') === -1) {
      displayEl.textContent = cur ? (cur + '.') : '0.';
    }
    return;
  }

  var maxLen = opts.maxLen || 10;
  if (cur.length < maxLen) {
    displayEl.textContent = cur + String(val || '');
  }
};

window.wbKeypadOK_ = function() {
  var inputEl = window._wbKeypadEl;
  var opts = window._wbKeypadOpts || {};
  var displayEl = document.getElementById('wb-keypad-display');
  var nextVal = displayEl ? (displayEl.textContent || '') : '';

  if (inputEl) {
    inputEl.value = nextVal;
    inputEl.dispatchEvent(new Event('input', { bubbles: true }));
    inputEl.dispatchEvent(new Event('change', { bubbles: true }));
    if (typeof opts.onApply === 'function') {
      try {
        opts.onApply(nextVal);
      } catch (e) {
      }
    }
  }

  window.wbKeypadCancel_();
};

window.wbKeypadCancel_ = function() {
  var modalEl = document.getElementById('wb-keypad-modal');
  if (modalEl) modalEl.style.display = 'none';
  window._wbKeypadEl = null;
  window._wbKeypadOpts = null;
};

function wbKeypadLabelForInput_(inputEl) {
  if (!inputEl) return 'Numeric Input';
  if (inputEl.dataset && inputEl.dataset.keypadLabel) return String(inputEl.dataset.keypadLabel);
  var field = inputEl.closest('.wb-field');
  var fieldLabel = field && field.querySelector('.wb-field-label');
  if (fieldLabel && fieldLabel.textContent) return String(fieldLabel.textContent).trim();
  return 'Numeric Input';
}

function wbOpenNumericKeypadForInput_(inputEl) {
  if (!inputEl || inputEl.disabled) return false;
  var label = wbKeypadLabelForInput_(inputEl);

  if (typeof window.wbKeypadOpen_ === 'function') {
    window.wbKeypadOpen_(inputEl, label, { decimal: true });
    return true;
  }

  if (typeof window.flightDeckOpenNumericKeypad === 'function') {
    var inputId = wbEnsureInputId_(inputEl);
    if (!inputId) return false;

    var min = parseFloat(inputEl.min);
    window.flightDeckOpenNumericKeypad(inputId, {
      title: label,
      allowSign: isFinite(min) && min < 0,
      clearOnOpen: false
    });
    return true;
  }

  if (typeof window.briefOpenKeypad_ === 'function') {
    window.briefOpenKeypad_(inputEl, label, { decimal: true });
    return true;
  }

  return false;
}

function wbBindNumericKeypad_() {
  if (window._wbNumericKeypadBound) return;
  window._wbNumericKeypadBound = true;

  document.addEventListener('pointerdown', function(e) {
    var input = e.target && e.target.closest ? e.target.closest('#wb-container input[type="number"]') : null;
    if (!input || input.disabled || input.dataset.keypad === 'off') return;
    if (wbOpenNumericKeypadForInput_(input)) {
      e.preventDefault();
      input.blur();
    }
  }, true);

  document.addEventListener('focusin', function(e) {
    var input = e.target && e.target.matches && e.target.matches('#wb-container input[type="number"]') ? e.target : null;
    if (!input || input.disabled || input.dataset.keypad === 'off') return;
    if (wbOpenNumericKeypadForInput_(input)) {
      input.blur();
    }
  });
}

wbBindNumericKeypad_();

function wbDefaultEnvelope_(emptyArm, mtow) {
  const minWeight = Math.max(900, Math.round((mtow || 1600) * 0.58));
  const maxWeight = Math.max(minWeight + 200, Math.round(mtow || 1600));
  const fwdMin = Math.max(34, wbNum_(emptyArm, 47) - 5.8);
  const fwdMax = fwdMin + 1.7;
  const aftMax = fwdMax + 6.2;
  const aftMin = aftMax - 1.4;
  return [
    { POINT_SEQUENCE: 1, CG_Arm_X: fwdMin, Weight_Y: minWeight },
    { POINT_SEQUENCE: 2, CG_Arm_X: fwdMax, Weight_Y: maxWeight },
    { POINT_SEQUENCE: 3, CG_Arm_X: aftMax, Weight_Y: maxWeight },
    { POINT_SEQUENCE: 4, CG_Arm_X: aftMin, Weight_Y: minWeight }
  ];
}

function wbNormalizeEnvelope_(rawEnvelope) {
  if (!Array.isArray(rawEnvelope)) return [];
  return rawEnvelope
    .map(function(point, idx) {
      return {
        POINT_SEQUENCE: wbNum_(point && point.POINT_SEQUENCE, idx + 1),
        CG_Arm_X: wbNum_(point && point.CG_Arm_X, NaN),
        Weight_Y: wbNum_(point && point.Weight_Y, NaN)
      };
    })
    .filter(function(point) {
      return isFinite(point.CG_Arm_X) && isFinite(point.Weight_Y);
    })
    .sort(function(a, b) {
      return wbNum_(a.POINT_SEQUENCE, 0) - wbNum_(b.POINT_SEQUENCE, 0);
    });
}

function wbFindCachedEnvelopeForAircraft_(aircraftReg) {
  const byAircraftCache = wbReadRealEnvelopeForAircraft_(aircraftReg);
  if (byAircraftCache.length >= 3) return byAircraftCache;

  const target = String(aircraftReg || '').trim().toUpperCase();
  if (!target) return [];

  try {
    const keys = Object.keys(localStorage || {});
    for (let i = 0; i < keys.length; i++) {
      const key = keys[i];
      if (!String(key || '').startsWith('mba_cache_wb_')) continue;
      const raw = localStorage.getItem(key);
      if (!raw) continue;
      const parsed = JSON.parse(raw);
      const payload = parsed && parsed.value ? parsed.value : null;
      if (!payload || String(payload.aircraft || '').trim().toUpperCase() !== target) continue;
      const normalized = wbNormalizeEnvelope_(payload.envelopeData);
      if (normalized.length >= 3) return normalized;
    }
  } catch (e) {}

  return [];
}

function wbEnvelopeCacheKey_(aircraftReg) {
  const key = String(aircraftReg || '').trim().toUpperCase();
  return key ? ('mba_cache_envelope_' + key) : '';
}

function wbCacheRealEnvelopeForAircraft_(aircraftReg, envelopeData) {
  try {
    const key = wbEnvelopeCacheKey_(aircraftReg);
    const normalized = wbNormalizeEnvelope_(envelopeData);
    if (!key || normalized.length < 3) return;
    localStorage.setItem(key, JSON.stringify({
      aircraft: String(aircraftReg || '').trim().toUpperCase(),
      cachedAt: new Date().toISOString(),
      envelopeData: normalized
    }));
  } catch (e) {}
}

function wbReadRealEnvelopeForAircraft_(aircraftReg) {
  try {
    const key = wbEnvelopeCacheKey_(aircraftReg);
    if (!key) return [];
    const raw = localStorage.getItem(key);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return wbNormalizeEnvelope_(parsed && parsed.envelopeData);
  } catch (e) {
    return [];
  }
}

function wbPassengerExists_(name) {
  const needle = wbNormName_(name).toUpperCase();
  if (!needle) return false;

  const inSeats = Object.keys(window.wbData.seats || {}).some(function(seatId) {
    const s = window.wbData.seats[seatId];
    return wbNormName_(s && s.passenger && s.passenger.name).toUpperCase() === needle;
  });
  if (inSeats) return true;

  return (window.wbData.waitingPassengers || []).some(function(p) {
    return wbNormName_(p && p.name).toUpperCase() === needle;
  });
}

function wbSetMissionCache_(missionId, missionObj) {
  try {
    if (!missionId || !missionObj) return;
    localStorage.setItem('mba_cache_mission_' + missionId, JSON.stringify({
      savedAt: Date.now(),
      value: missionObj
    }));
  } catch (e) {
    console.warn('Failed writing mission cache from W&B', e);
  }
}

function wbPatchOutboxMission_(missionId, flightId, paxList) {
  try {
    const raw = localStorage.getItem('mba_outbox_v1');
    if (!raw) return;
    const parsed = JSON.parse(raw);
    const items = parsed && Array.isArray(parsed.value) ? parsed.value : [];
    let changed = false;

    items.forEach(function(item) {
      if (String(item && item.method || '') !== 'saveMission') return;
      const args = Array.isArray(item.args) ? item.args : [];
      const payload = args[0];
      if (!payload || typeof payload !== 'object') return;
      const legs = Array.isArray(payload.legs) ? payload.legs : [];
      legs.forEach(function(leg) {
        const legFlightId = String(leg && leg.flightLegId || '').trim();
        const legMissionId = wbMissionIdFromFlightId_(legFlightId);
        if (legFlightId === String(flightId || '').trim() || (missionId && legMissionId === missionId)) {
          leg.pax = Array.isArray(paxList) ? paxList.map(function(p) { return { ...p }; }) : [];
          changed = true;
        }
      });
    });

    if (changed) {
      localStorage.setItem('mba_outbox_v1', JSON.stringify({
        savedAt: Date.now(),
        value: items
      }));
    }
  } catch (e) {
    console.warn('Failed patching outbox from W&B', e);
  }
}

function wbBuildPaxForSync_() {
  const byName = {};

  const cargoByPax = {};
  (window.wbData.cargoManifest || []).forEach(function(c) {
    if (String(c && c.type || '').toLowerCase() !== 'pax_cargo') return;
    const linked = wbNormName_(c && c.linkedPassenger);
    if (!linked) return;
    const key = linked.toUpperCase();
    cargoByPax[key] = wbNum_(c.actualWeight, wbNum_(c.plannedWeight, 0));
  });

  Object.keys(window.wbData.seats || {}).forEach(function(seatId) {
    const seat = window.wbData.seats[seatId] || {};
    const passenger = seat.passenger || null;
    if (!passenger || !passenger.name) return;
    const isPilot = wbNormName_(window.wbData.pilot).toUpperCase() === wbNormName_(passenger.name).toUpperCase();
    if (isPilot) return;

    const key = wbNormName_(passenger.name).toUpperCase();
    byName[key] = {
      name: wbNormName_(passenger.name),
      weight: wbNum_(passenger.actualWeight, wbNum_(passenger.weight, 0)),
      cargo: wbNum_(cargoByPax[key], wbNum_(passenger.cargo, 0)),
      fund: String(passenger.fund || ''),
      category: String(passenger.category || ''),
      chargeRate: String(passenger.chargeRate || ''),
      chargedAmount: wbNum_(passenger.chargedAmount, 0),
      phone: String(passenger.phone || '')
    };
  });

  (window.wbData.waitingPassengers || []).forEach(function(passenger) {
    if (!passenger || !passenger.name) return;
    const key = wbNormName_(passenger.name).toUpperCase();
    byName[key] = {
      name: wbNormName_(passenger.name),
      weight: wbNum_(passenger.actualWeight, wbNum_(passenger.weight, 0)),
      cargo: wbNum_(cargoByPax[key], wbNum_(passenger.cargo, 0)),
      fund: String(passenger.fund || ''),
      category: String(passenger.category || ''),
      chargeRate: String(passenger.chargeRate || ''),
      chargedAmount: wbNum_(passenger.chargedAmount, 0),
      phone: String(passenger.phone || '')
    };
  });

  const paxList = Object.keys(byName).map(function(key) {
    return byName[key];
  });

  (window.wbData.cargoManifest || []).forEach(function(cargo) {
    if (String(cargo && cargo.type || '').toLowerCase() !== 'freight') return;
    const weight = wbNum_(cargo.actualWeight, wbNum_(cargo.plannedWeight, 0));
    if (!(weight > 0)) return;
    paxList.push({
      name: 'FREIGHT',
      weight: weight,
      cargo: weight,
      fund: String(cargo.fund || ''),
      category: 'FREIGHT',
      chargeRate: String(cargo.chargeRate || ''),
      chargedAmount: wbNum_(cargo.chargedAmount, 0),
      description: String(cargo.name || 'Freight')
    });
  });

  return paxList;
}

function wbPersistMissionPax_() {
  try {
    const flightId = String(window.wbData && window.wbData.flightId || '').trim();
    if (!flightId) return;

    const missionId = wbMissionIdFromFlightId_(flightId);
    const mission = (window.currentBriefingMission && Array.isArray(window.currentBriefingMission.legs) && window.currentBriefingMission.legs.length)
      ? window.currentBriefingMission
      : wbReadCachedMissionByFlightId_(flightId);
    if (!mission || !Array.isArray(mission.legs) || !mission.legs.length) return;

    const targetLeg = wbFindMissionLegByFlightId_(mission, flightId);
    if (!targetLeg) return;

    const paxList = wbBuildPaxForSync_();
    targetLeg.pax = paxList;

    if (window.currentBriefingMission && String(window.currentBriefingMission.id || '') === String(mission.id || missionId || '')) {
      window.currentBriefingMission = mission;
    }

    wbSetMissionCache_(String(mission.id || missionId || ''), mission);
    wbPatchOutboxMission_(String(mission.id || missionId || ''), flightId, paxList);
  } catch (e) {
    console.warn('wbPersistMissionPax_ failed', e);
  }
}

function wbBuildOfflinePayload_(flightId) {
  const mission = (window.currentBriefingMission && Array.isArray(window.currentBriefingMission.legs) && window.currentBriefingMission.legs.length)
    ? window.currentBriefingMission
    : wbReadCachedMissionByFlightId_(flightId);
  if (!mission || !Array.isArray(mission.legs) || !mission.legs.length) return null;

  const leg = wbFindMissionLegByFlightId_(mission, flightId);
  if (!leg) return null;

  const aircraftReg = String(mission.acft || '').trim();
  const pilotName = String(mission.pilot || '').trim();
  const aircraftObj = wbFindAircraft_(aircraftReg) || {};
  const wbTemplate = (function() {
    const target = String(aircraftReg || '').trim().toUpperCase();
    if (!target) return null;
    try {
      const keys = Object.keys(localStorage || {});
      for (let i = 0; i < keys.length; i++) {
        const key = String(keys[i] || '');
        if (!key.startsWith('mba_cache_wb_')) continue;
        const raw = localStorage.getItem(key);
        if (!raw) continue;
        const parsed = JSON.parse(raw);
        const payload = parsed && parsed.payload ? parsed.payload : null;
        if (!payload) continue;
        if (String(payload.aircraft || '').trim().toUpperCase() !== target) continue;
        if (!payload.seats || !payload.items) continue;
        return payload;
      }
    } catch (e) {}
    return null;
  })();

  const emptyArm = wbNum_(wbTemplate && wbTemplate.airframeData && wbTemplate.airframeData.Empty_Arm, wbNum_(aircraftObj.emptyArm, 47));
  const mtow = wbNum_(wbTemplate && wbTemplate.airframeData && wbTemplate.airframeData.MTOW, wbNum_(aircraftObj.mtow, 1600));
  const emptyWeight = wbNum_(wbTemplate && wbTemplate.airframeData && wbTemplate.airframeData.Empty_Weight, wbNum_(aircraftObj.emptyWeight, 1000));
  const fuelBurn = wbNum_(wbTemplate && wbTemplate.airframeData && wbTemplate.airframeData.Fuel_Burn_Per_Hour, wbNum_(aircraftObj.burn, 60));
  const pilotArm = Math.max(30, emptyArm - 9);
  const fuelArm = wbNum_(wbTemplate && wbTemplate.fuelArm, emptyArm);
  const midArm = pilotArm + 28;
  const aftArm = midArm + 22;
  const cargoArm = aftArm + 16;

  const seatDefs = (function() {
    if (wbTemplate && wbTemplate.seats && typeof wbTemplate.seats === 'object') {
      const knownOrder = ['pilot', 'copilot', 'rh-mid', 'lh-mid', 'rh-aft', 'lh-aft'];
      const defs = Object.keys(wbTemplate.seats).map(function(seatName) {
        const seat = wbTemplate.seats[seatName] || {};
        const key = String(seat.seatId || seatName || '').trim();
        const seatNameNorm = String(seatName || '').trim().toUpperCase();
        const keyNorm = String(key || '').trim().toLowerCase();
        const isPilot = (keyNorm === 'pilot') || /^PILOT\b/.test(seatNameNorm);
        return {
          key: key || String(seatName || '').trim().toLowerCase().replace(/\s+/g, '-'),
          label: String(seatName || seat.label || 'Seat'),
          arm: wbNum_(seat.arm, pilotArm),
          seatWeight: wbNum_(seat.weight, wbNum_(aircraftObj.pilotSeat, 12)),
          locked: !!(seat.locked || isPilot)
        };
      }).filter(function(def) { return !!def.key; });

      defs.sort(function(a, b) {
        const ai = knownOrder.indexOf(String(a.key || '').toLowerCase());
        const bi = knownOrder.indexOf(String(b.key || '').toLowerCase());
        const sa = ai >= 0 ? ai : 999;
        const sb = bi >= 0 ? bi : 999;
        if (sa !== sb) return sa - sb;
        return String(a.label || '').localeCompare(String(b.label || ''));
      });
      if (defs.length >= 2) return defs;
    }

    return [
      { key: 'pilot', label: 'Pilot Seat', arm: pilotArm, seatWeight: wbNum_(aircraftObj.pilotSeat, 13), locked: true },
      { key: 'copilot', label: 'Copilot Seat', arm: pilotArm, seatWeight: wbNum_(aircraftObj.pilotSeat, 13), locked: false },
      { key: 'lh-mid', label: 'LH Mid Seat', arm: midArm, seatWeight: wbNum_(aircraftObj.midSeat, 11), locked: false },
      { key: 'rh-mid', label: 'RH Mid Seat', arm: midArm, seatWeight: wbNum_(aircraftObj.midSeat, 11), locked: false },
      { key: 'lh-aft', label: 'LH Aft Seat', arm: aftArm, seatWeight: wbNum_(aircraftObj.aftSeat, 10), locked: false },
      { key: 'rh-aft', label: 'RH Aft Seat', arm: aftArm, seatWeight: wbNum_(aircraftObj.aftSeat, 10), locked: false }
    ];
  })();

  const rawPax = Array.isArray(leg.pax) ? leg.pax : [];
  const paxList = rawPax
    .filter(function(p) { return String(p && p.name || '').toUpperCase() !== 'FREIGHT'; })
    .map(function(p) {
      const w = wbNum_(p.actualWeight, wbNum_(p.weight, 0));
      return {
        ...p,
        name: wbNormName_(p.name),
        weight: w,
        plannedWeight: wbNum_(p.plannedWeight, w),
        actualWeight: w,
        cargo: wbNum_(p.cargo, 0)
      };
    })
    .filter(function(p) { return !!p.name; })
    .sort(function(a, b) { return wbNum_(b.weight, 0) - wbNum_(a.weight, 0); });

  const freightList = rawPax
    .filter(function(p) { return String(p && p.name || '').toUpperCase() === 'FREIGHT'; })
    .map(function(p) {
      const w = wbNum_(p.cargo, wbNum_(p.weight, 0));
      return {
        name: 'Freight',
        plannedWeight: w,
        actualWeight: w,
        type: 'freight',
        passengerLinked: false,
        fund: String(p.fund || ''),
        chargeRate: String(p.chargeRate || ''),
        chargedAmount: wbNum_(p.chargedAmount, 0)
      };
    })
    .filter(function(c) { return wbNum_(c.actualWeight, 0) > 0; });

  const seatAssignments = {};
  const seats = {};
  const maxPaxSeats = seatDefs.length - 1;
  const assignedPax = Math.min(paxList.length, maxPaxSeats);
  const installCount = Math.min(1 + paxList.length, seatDefs.length);
  const pilotWeightInfo = wbResolvePilotWeight_(pilotName);
  const pilotWeight = pilotWeightInfo.weight;

  seatDefs.forEach(function(def, idx) {
    let status = 'base';
    let passenger = null;
    let occupiedWeight = 0;
    let isOccupied = false;

    if (idx === 0) {
      status = 'installed';
      passenger = {
        name: pilotName,
        weight: pilotWeight,
        plannedWeight: pilotWeight,
        actualWeight: pilotWeight,
        verified: true
      };
      occupiedWeight = pilotWeight;
      isOccupied = true;
    } else if (idx < installCount) {
      status = 'installed';
      const pax = paxList[idx - 1] || null;
      if (pax) {
        passenger = { ...pax };
        occupiedWeight = wbNum_(pax.actualWeight, wbNum_(pax.weight, 0));
        isOccupied = occupiedWeight > 0;
      }
    }

    seatAssignments[def.key] = {
      label: def.label,
      arm: def.arm,
      seatWeight: def.seatWeight,
      status: status,
      passenger: passenger,
      occupiedWeight: occupiedWeight,
      isOccupied: isOccupied,
      enabled: status === 'installed',
      locked: !!def.locked
    };

    seats[def.label] = {
      seatId: def.key,
      weight: def.seatWeight,
      arm: def.arm,
      status: status,
      enabled: status === 'installed',
      passenger: passenger,
      occupiedWeight: occupiedWeight,
      locked: !!def.locked
    };
  });

  const waitingPassengers = paxList.slice(assignedPax).map(function(p) { return { ...p }; });

  const cargoManifest = [];
  paxList.forEach(function(p) {
    const cargo = wbNum_(p.cargo, 0);
    if (cargo > 0) {
      cargoManifest.push({
        name: p.name + ' Cargo',
        plannedWeight: cargo,
        actualWeight: cargo,
        type: 'pax_cargo',
        linkedPassenger: p.name,
        passengerLinked: true,
        fund: String(p.fund || ''),
        chargeRate: String(p.chargeRate || ''),
        chargedAmount: wbNum_(p.chargedAmount, 0)
      });
    }
  });
  freightList.forEach(function(item) { cargoManifest.push(item); });

  const cargoAreas = (function() {
    if (wbTemplate && Array.isArray(wbTemplate.cargoAreas) && wbTemplate.cargoAreas.length) {
      return wbTemplate.cargoAreas.map(function(area) {
        return {
          id: String(area && area.id || '').trim() || ('cargo_' + Math.random().toString(36).slice(2, 6)),
          name: String(area && area.name || 'Cargo').trim(),
          arm: wbNum_(area && area.arm, cargoArm),
          maxWeightKg: wbNum_(area && area.maxWeightKg, 0),
          maxWeightLbs: wbNum_(area && area.maxWeightLbs, 0)
        };
      });
    }
    return [
      { id: 'cargo_a', name: 'Cargo A', arm: cargoArm - 10, maxWeightKg: 90, maxWeightLbs: 198 },
      { id: 'cargo_b', name: 'Cargo B', arm: cargoArm, maxWeightKg: 110, maxWeightLbs: 243 },
      { id: 'cargo_c', name: 'Cargo C', arm: cargoArm + 10, maxWeightKg: 90, maxWeightLbs: 198 }
    ];
  })();

  const cargoAreaWeights = {};
  cargoAreas.forEach(function(area) {
    cargoAreaWeights[area.id] = { planned: 0, actual: 0 };
  });
  const plannedCargo = cargoManifest.reduce(function(sum, c) {
    return sum + wbNum_(c.actualWeight, wbNum_(c.plannedWeight, 0));
  }, 0);
  if (plannedCargo > 0 && cargoAreas.length) {
    cargoAreaWeights[cargoAreas[0].id] = { planned: plannedCargo, actual: plannedCargo };
  }

  const fuelLiters = wbNum_(leg.takeoffFuel, wbNum_(leg.fuel, 0));
  const fuelWeight = fuelLiters > 0 ? (fuelLiters * 0.72) : 0;

  const items = [
    {
      name: 'Empty Aircraft',
      plannedWeight: emptyWeight,
      actualWeight: emptyWeight,
      arm: emptyArm,
      type: 'empty'
    }
  ];

  Object.keys(seats).forEach(function(seatName) {
    const seat = seats[seatName];
    if (seat.status === 'installed' && seat.occupiedWeight > 0 && seat.passenger) {
      const isPilot = wbNormName_(seat.passenger.name).toUpperCase() === wbNormName_(pilotName).toUpperCase();
      items.push({
        name: isPilot ? seatName : (seatName + ': ' + seat.passenger.name),
        plannedWeight: wbNum_(seat.passenger.plannedWeight, wbNum_(seat.passenger.weight, 0)),
        actualWeight: wbNum_(seat.passenger.actualWeight, wbNum_(seat.passenger.weight, 0)),
        arm: wbNum_(seat.arm, 0),
        type: 'passenger',
        seatId: seatName
      });
    }
  });

  items.push({
    name: 'Cargo',
    plannedWeight: plannedCargo,
    actualWeight: plannedCargo,
    arm: cargoArm,
    type: 'cargo'
  });

  items.push({
    name: 'Fuel',
    plannedWeight: fuelWeight,
    actualWeight: fuelWeight,
    arm: fuelArm,
    type: 'fuel'
  });

  const envelopeData = wbFindCachedEnvelopeForAircraft_(aircraftReg);
  if (envelopeData.length < 3) return null;

  return {
    flightId: String(leg.flightLegId || flightId || ''),
    aircraft: aircraftReg,
    pilot: pilotName,
    pilotWeightEstimated: !!pilotWeightInfo.estimated,
    pilotWeightEstimateReason: String(pilotWeightInfo.reason || ''),
    date: String(mission.date || ''),
    time: String(leg.time || mission.time || ''),
    route: `${String(leg.from || '').trim().toUpperCase()} -> ${String(leg.to || '').trim().toUpperCase()}`,
    airframeData: {
      Empty_Weight: emptyWeight,
      Empty_Arm: emptyArm,
      MTOW: mtow,
      Fuel_Burn_Per_Hour: fuelBurn
    },
    envelopeData: envelopeData,
    items: items,
    seats: seats,
    seatAssignments: seatAssignments,
    waitingPassengers: waitingPassengers,
    maxPaxInMission: Math.max(paxList.length, assignedPax),
    thisLegPaxCount: paxList.length,
    fuel: fuelWeight,
    fuelArm: fuelArm,
    cargoAreas: cargoAreas,
    cargoAreaWeights: cargoAreaWeights,
    cargoManifest: cargoManifest
  };
}

window.onWbPaxPick = function() {
  const paxSelect = document.getElementById('wb-add-pax-select');
  if (!paxSelect) return;
  const name = wbNormName_(paxSelect.value);
  if (!name) return;
};

window.refreshOfflineLoadTools = function() {
  const paxSelect = document.getElementById('wb-add-pax-select');
  const fundSelect = document.getElementById('wb-add-pax-fund');
  const rateSelect = document.getElementById('wb-add-pax-rate');
  const freightFund = document.getElementById('wb-add-freight-fund');
  const freightRate = document.getElementById('wb-add-freight-rate');
  if (!paxSelect || !fundSelect || !rateSelect || !freightFund || !freightRate) return;

  const source = wbDropdownSource_();
  const passengers = (Array.isArray(source.passengers) ? source.passengers.slice() : []).concat(wbMissionPassengers_());
  const seenPax = {};
  const dedupedPassengers = passengers.filter(function(p) {
    const name = wbNormName_(p && p.name).toUpperCase();
    if (!name || seenPax[name]) return false;
    seenPax[name] = true;
    return true;
  });
  dedupedPassengers.sort(function(a, b) {
    return String(a && a.name || '').localeCompare(String(b && b.name || ''));
  });
  paxSelect.innerHTML = '<option value="">Select Passenger</option>' + dedupedPassengers.map(function(p) {
    const name = wbNormName_(p && p.name);
    const weight = wbNum_(p && p.weight, 80);
    return `<option value="${name}" data-wb-weight="${weight}">${name} (${Math.round(weight)}kg)</option>`;
  }).join('');

  const funds = Array.isArray(source.funds) ? source.funds : [];
  const fundHtml = '<option value="">No Fund</option>' + funds.map(function(f) {
    const id = String(f && (f.id || f.displayName) || '').trim();
    const label = String(f && (f.displayName || f.id) || id);
    return `<option value="${id}">${label}</option>`;
  }).join('');
  fundSelect.innerHTML = fundHtml;
  freightFund.innerHTML = fundHtml;

  const rates = Array.isArray(source.rates) ? source.rates : [];
  const rateHtml = '<option value="">Rate</option>' + rates.map(function(r) {
    const val = String(r || '').trim();
    return `<option value="${val}">${val}</option>`;
  }).join('');
  rateSelect.innerHTML = rateHtml;
  freightRate.innerHTML = rateHtml;

  const freightTools = document.getElementById('wb-freight-tools');
  if (freightTools) {
    freightTools.style.display = 'flex';
    freightTools.dataset.initialized = '1';
  }

  window.onWbPaxPick();
};

window.toggleOfflineFreightTools = function() {
  const freightTools = document.getElementById('wb-freight-tools');
  if (!freightTools) return;
  freightTools.style.display = (freightTools.style.display === 'none' || !freightTools.style.display) ? 'flex' : 'none';
};

window.addOfflinePaxToWb = function() {
  const paxSelectEl = document.getElementById('wb-add-pax-select');
  const name = wbNormName_((paxSelectEl || {}).value);
  const selectedOpt = paxSelectEl && paxSelectEl.selectedOptions ? paxSelectEl.selectedOptions[0] : null;
  const weight = wbNum_(selectedOpt && selectedOpt.getAttribute('data-wb-weight'), 0);
  const cargo = Math.max(0, wbNum_((document.getElementById('wb-add-pax-cargo') || {}).value, 0));
  const fund = String((document.getElementById('wb-add-pax-fund') || {}).value || '').trim();
  const rate = String((document.getElementById('wb-add-pax-rate') || {}).value || '').trim();

  if (!name) {
    if (window.M) M.toast({ html: 'Select a passenger to add', classes: 'orange' });
    return;
  }
  if (!(weight > 0)) {
    if (window.M) M.toast({ html: 'Enter a valid passenger weight', classes: 'orange' });
    return;
  }
  if (wbPassengerExists_(name)) {
    if (window.M) M.toast({ html: 'Passenger already in manifest/waiting list', classes: 'orange' });
    return;
  }

  const passenger = {
    name: name,
    weight: weight,
    plannedWeight: weight,
    actualWeight: weight,
    cargo: cargo,
    fund: fund,
    chargeRate: rate,
    chargedAmount: 0,
    verified: false
  };

  if (!Array.isArray(window.wbData.waitingPassengers)) window.wbData.waitingPassengers = [];
  window.wbData.waitingPassengers.push(passenger);

  if (!window.wbData.seatAssignments || typeof window.wbData.seatAssignments !== 'object') {
    window.wbData.seatAssignments = {};
  }
  window.wbData.seatAssignments['offline_' + Date.now()] = {
    label: 'Offline Added',
    passenger: { ...passenger },
    isOccupied: false,
    status: 'waiting'
  };

  if (cargo > 0) {
    if (!Array.isArray(window.wbData.cargoManifest)) window.wbData.cargoManifest = [];
    const existingIdx = window.wbData.cargoManifest.findIndex(function(c) {
      return String(c && c.type || '').toLowerCase() === 'pax_cargo'
        && wbNormName_(c && c.linkedPassenger).toUpperCase() === name.toUpperCase();
    });
    const nextItem = {
      name: name + ' Cargo',
      plannedWeight: cargo,
      actualWeight: cargo,
      type: 'pax_cargo',
      linkedPassenger: name,
      passengerLinked: true,
      fund: fund,
      chargeRate: rate,
      chargedAmount: 0
    };
    if (existingIdx >= 0) {
      window.wbData.cargoManifest[existingIdx] = nextItem;
    } else {
      window.wbData.cargoManifest.push(nextItem);
    }
  }

  wbPersistMissionPax_();
  window.renderManifest();
  window.renderPaxCargoSummary();
  window.updateWBUI();

  const cargoInput = document.getElementById('wb-add-pax-cargo');
  if (cargoInput) cargoInput.value = '0';

  if (window.M) M.toast({ html: name + ' added to waiting passengers', classes: 'green' });
};

window.addOfflineFreightToWb = function() {
  const weight = Math.max(0, wbNum_((document.getElementById('wb-add-freight-kg') || {}).value, 0));
  const fund = String((document.getElementById('wb-add-freight-fund') || {}).value || '').trim();
  const rate = String((document.getElementById('wb-add-freight-rate') || {}).value || '').trim();

  if (!(weight > 0)) {
    if (window.M) M.toast({ html: 'Enter freight weight in kg', classes: 'orange' });
    return;
  }

  if (!Array.isArray(window.wbData.cargoManifest)) window.wbData.cargoManifest = [];
  window.wbData.cargoManifest.push({
    name: 'Freight',
    plannedWeight: weight,
    actualWeight: weight,
    type: 'freight',
    passengerLinked: false,
    fund: fund,
    chargeRate: rate,
    chargedAmount: 0
  });

  wbPersistMissionPax_();
  window.renderManifest();
  window.renderPaxCargoSummary();
  window.updateWBUI();

  const input = document.getElementById('wb-add-freight-kg');
  if (input) input.value = '0';

  if (window.M) M.toast({ html: 'Freight added', classes: 'green' });
};

/**
 * Calculate moment for a single item
 * Arms are in inches, so moment = weight (kg) ├ù arm (inches)
 */
window.calculateMoment = function(weight, arm) {
  return weight * arm; // Moment = Weight ├ù Arm (kg-in)
};

function wbSeatStowArm_() {
  const areas = Array.isArray(window.wbData.cargoAreas) ? window.wbData.cargoAreas : [];
  const maxCargoArm = areas.reduce(function(max, area) {
    const a = wbNum_(area && area.arm, NaN);
    return isFinite(a) ? Math.max(max, a) : max;
  }, -Infinity);
  if (isFinite(maxCargoArm)) return maxCargoArm + 6;
  return wbNum_(window.wbData && window.wbData.fuelArm, 48) + 28;
}

/**
 * Calculate total W&B
 */
window.calculateWB = function() {
  let totalWeight = 0;
  let totalMoment = 0;

  const seatStowArm = wbSeatStowArm_();

  // Calculate adjusted empty weight (subtract seats left at base)
  let emptyWeight = window.wbData.airframeData.Empty_Weight;
  Object.values(window.wbData.seats).forEach(seat => {
    if (seat.status === 'base') {
      emptyWeight -= seat.weight;
    }
  });

  // Add all manifest items (including fuel which is now in the items array)
  window.wbData.items.forEach(item => {
    if (item.type === 'cargo') return;
    const weight = (item.type === 'empty') ? emptyWeight : (parseFloat(item.actualWeight) || 0);
    const arm = parseFloat(item.arm) || 0;
    totalWeight += weight;
    totalMoment += window.calculateMoment(weight, arm);
  });

  // Empty weight already includes installed seat structure weight.
  // Only apply deltas for seats that moved from their installed location.
  Object.keys(window.wbData.seats).forEach(seatName => {
    const seat = window.wbData.seats[seatName];
    const seatWeight = wbNum_(seat && seat.weight, 0);
    const seatArm = wbNum_(seat && seat.arm, 0);
    if (!(seatWeight > 0)) return;

    if (seat.status === 'cargo') {
      totalMoment += window.calculateMoment(seatWeight, (seatStowArm - seatArm));
    }
  });

  // Add cargo area weights
  if (window.wbData.cargoAreas && window.wbData.cargoAreaWeights) {
    window.wbData.cargoAreas.forEach(area => {
      const areaWeights = window.wbData.cargoAreaWeights[area.id];
      if (areaWeights) {
        const weight = parseFloat(areaWeights.actual) || 0;
        const arm = parseFloat(area.arm) || 0;
        if (weight > 0) {
          totalWeight += weight;
          totalMoment += window.calculateMoment(weight, arm);
        }
      }
    });
  }

  // Add individual cargo manifest weights
  if (window.wbData.cargoManifest) {
    window.wbData.cargoManifest.forEach(cargo => {
      // Cargo manifest items don't have a specific arm, they use the cargo arm already in items
      // But we might want to track them separately for accountability only, not in W&B
      // So we skip them here to avoid double-counting
    });
  }

  const cg = totalWeight > 0 ? (totalMoment / totalWeight) : 0;

  return {
    totalWeight,
    totalMoment,
    cg,
    empty: {
      weight: parseFloat(window.wbData.airframeData?.Empty_Weight) || 0,
      arm: parseFloat(window.wbData.airframeData?.Empty_Arm) || 0
    },
    mtow: parseFloat(window.wbData.airframeData?.MTOW) || 0,
  };
};

/**
 * Check if CG is within envelope
 */
window.checkEnvelope = function(cgArm, weight) {
  const envelope = wbNormalizeEnvelope_(window.wbData.envelopeData);
  if (!envelope || envelope.length < 3) return false;

  // Ray casting algorithm for point-in-polygon
  let inside = false;
  for (let i = 0, j = envelope.length - 1; i < envelope.length; j = i++) {
    const xi = parseFloat(envelope[i].CG_Arm_X);
    const yi = parseFloat(envelope[i].Weight_Y);
    const xj = parseFloat(envelope[j].CG_Arm_X);
    const yj = parseFloat(envelope[j].Weight_Y);

    const intersect = ((yi > weight) !== (yj > weight))
      && (cgArm < (xj - xi) * (weight - yi) / (yj - yi) + xi);
    if (intersect) inside = !inside;
  }

  return inside;
};

/**
 * Draw CG envelope and current position
 */
window.drawCGGraph = function() {
  const canvas = document.getElementById('wb-cg-canvas');
  if (!canvas) return;

  const ctx = canvas.getContext('2d');
  const rect = canvas.getBoundingClientRect();
  const dpr = window.devicePixelRatio || 1;
  canvas.width = Math.max(1, Math.round(rect.width * dpr));
  canvas.height = Math.max(1, Math.round(rect.height * dpr));
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

  const padding = 48;
  const w = rect.width - 2 * padding;
  const h = rect.height - 2 * padding;

  // Draw grid background
  ctx.fillStyle = '#f9f9f9';
  ctx.fillRect(padding, padding, w, h);

  // Get envelope data
  const envelope = wbNormalizeEnvelope_(window.wbData.envelopeData);
  if (!envelope || envelope.length === 0) {
    ctx.fillStyle = '#d32f2f';
    ctx.font = 'bold 14px Arial';
    ctx.fillText('NO REAL CG ENVELOPE AVAILABLE', padding + 20, rect.height / 2 - 8);
    ctx.fillStyle = '#666';
    ctx.font = '12px Arial';
    ctx.fillText('Cannot validate aircraft loading offline.', padding + 20, rect.height / 2 + 14);
    return;
  }

  // Determine scale
  const arms = envelope.map(p => parseFloat(p.CG_Arm_X));
  const weights = envelope.map(p => parseFloat(p.Weight_Y));
  const minArm = Math.min(...arms);
  const maxArm = Math.max(...arms);
  const minWeight = Math.min(...weights);
  const maxWeight = Math.max(...weights);

  const armScale = w / (maxArm - minArm || 1);
  const weightScale = h / (maxWeight - minWeight || 1);

  // Convert data to canvas coords
  const toCanvasX = (arm) => padding + (arm - minArm) * armScale;
  const toCanvasY = (weight) => padding + h - (weight - minWeight) * weightScale;

  // Draw envelope polygon
  ctx.strokeStyle = '#0b5394';
  ctx.fillStyle = 'rgba(11, 83, 148, 0.1)';
  ctx.lineWidth = 2;
  ctx.beginPath();
  envelope.forEach((point, idx) => {
    const x = toCanvasX(parseFloat(point.CG_Arm_X));
    const y = toCanvasY(parseFloat(point.Weight_Y));
    if (idx === 0) ctx.moveTo(x, y);
    else ctx.lineTo(x, y);
  });
  ctx.closePath();
  ctx.fill();
  ctx.stroke();

  // Draw envelope points
  ctx.fillStyle = '#0b5394';
  envelope.forEach(point => {
    const x = toCanvasX(parseFloat(point.CG_Arm_X));
    const y = toCanvasY(parseFloat(point.Weight_Y));
    ctx.beginPath();
    ctx.arc(x, y, 4, 0, Math.PI * 2);
    ctx.fill();
  });

  // Draw current CG point
  const wb = window.calculateWB();
  if (wb.totalWeight > 0) {
    const cgX = toCanvasX(wb.cg);
    const cgY = toCanvasY(wb.totalWeight);

    const isSafe = window.checkEnvelope(wb.cg, wb.totalWeight);
    ctx.fillStyle = isSafe ? '#2e7d32' : '#d32f2f';
    ctx.beginPath();
    ctx.arc(cgX, cgY, 6, 0, Math.PI * 2);
    ctx.fill();

    // Draw crosshairs
    ctx.strokeStyle = isSafe ? '#2e7d32' : '#d32f2f';
    ctx.lineWidth = 1;
    ctx.setLineDash([4, 4]);
    ctx.beginPath();
    ctx.moveTo(cgX, padding);
    ctx.lineTo(cgX, padding + h);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(padding, cgY);
    ctx.lineTo(padding + w, cgY);
    ctx.stroke();
    ctx.setLineDash([]);

    // Add labels for current CG and weight
    ctx.fillStyle = isSafe ? '#2e7d32' : '#d32f2f';
    ctx.font = 'bold 11px Arial';
    ctx.fillText(`CG: ${wb.cg.toFixed(1)}"`, cgX + 10, cgY - 20);
    ctx.fillText(`${Math.round(wb.totalWeight)} kg`, cgX + 10, cgY - 8);
    ctx.font = '10px Arial';
    ctx.fillText(`(TO)`, cgX + 10, cgY + 4);

    // Draw projected landing CG point
    if (window.wbData.time && window.wbData.airframeData && window.wbData.airframeData.Fuel_Burn_Per_Hour) {
      const timeStr = String(window.wbData.time || '00:00');
      const [hours, minutes] = timeStr.split(':').map(x => parseInt(x, 10));
      const flightHours = hours + (minutes / 60);
      const fuelBurnRate = window.wbData.airframeData.Fuel_Burn_Per_Hour || 12; // liters/hour
      const fuelBurned = flightHours * fuelBurnRate * 0.72;
      const landingWeight = Math.max(wb.totalWeight - fuelBurned, 0);
      
      if (landingWeight > 0 && fuelBurned > 0) {
        // Calculate landing CG - fuel burns from fuel tanks, which affects CG
        const fuelItem = window.wbData.items.find(item => item.type === 'fuel');
        const fuelArm = fuelItem ? parseFloat(fuelItem.arm) : 48; // Default fuel arm
        
        // Recalculate moment and CG for landing
        const landingMoment = wb.totalMoment - (fuelBurned * fuelArm);
        const landingCG = landingMoment / landingWeight;
        
        const landingCgX = toCanvasX(landingCG);
        const landingCgY = toCanvasY(landingWeight);
        
        const landingSafe = window.checkEnvelope(landingCG, landingWeight);
        ctx.fillStyle = landingSafe ? '#4caf50' : '#ff9800';
        ctx.beginPath();
        ctx.arc(landingCgX, landingCgY, 5, 0, Math.PI * 2);
        ctx.fill();
        
        // Draw dashed line between takeoff and landing
        ctx.strokeStyle = '#666';
        ctx.lineWidth = 1;
        ctx.setLineDash([3, 3]);
        ctx.beginPath();
        ctx.moveTo(cgX, cgY);
        ctx.lineTo(landingCgX, landingCgY);
        ctx.stroke();
        ctx.setLineDash([]);
        
        // Label landing point - position to avoid overlap
        ctx.fillStyle = landingSafe ? '#4caf50' : '#ff9800';
        ctx.font = '10px Arial';
        // Position label based on relative position to takeoff point
        const labelOffsetX = landingCgX < cgX ? -55 : 10;
        const labelOffsetY = landingCgY < cgY ? -8 : 12;
        ctx.fillText(`${Math.round(landingWeight)} kg (LND)`, landingCgX + labelOffsetX, landingCgY + labelOffsetY);
      }
    }
  }

  // Draw axes
  ctx.strokeStyle = '#333';
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(padding, padding + h);
  ctx.lineTo(padding + w, padding + h);
  ctx.stroke();
  ctx.beginPath();
  ctx.moveTo(padding, padding);
  ctx.lineTo(padding, padding + h);
  ctx.stroke();

  // Draw axis labels with better spacing
  ctx.fillStyle = '#333';
  ctx.font = 'bold 13px Arial';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillText('CG Arm (inches)', padding + w / 2, padding + h + 25);
  
  ctx.save();
  ctx.translate(padding - 74, padding + h / 2);
  ctx.rotate(-Math.PI / 2);
  ctx.fillText('Weight (kg)', 0, 0);
  ctx.restore();
  
  // Add scale markers
  ctx.font = '12px Arial';
  ctx.fillStyle = '#666';
  
  // X-axis values
  const xSteps = 4;
  for (let i = 0; i <= xSteps; i++) {
    const arm = minArm + (maxArm - minArm) * (i / xSteps);
    const x = toCanvasX(arm);
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    ctx.fillText(arm.toFixed(0), x, padding + h + 5);
  }
  
  // Y-axis values
  const ySteps = 4;
  for (let i = 0; i <= ySteps; i++) {
    const weight = minWeight + (maxWeight - minWeight) * (i / ySteps);
    const y = toCanvasY(weight);
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    ctx.fillText(weight.toFixed(0), padding - 12, y);
  }
};

/**
 * Update the entire W&B UI
 */
window.updateWBUI = function() {
  const wb = window.calculateWB();
  const envelope = wbNormalizeEnvelope_(window.wbData.envelopeData);
  const hasRealEnvelope = envelope.length >= 3;
  const isSafe = window.checkEnvelope(wb.cg, wb.totalWeight);

  // Update header - only update elements that exist
  if (document.getElementById('wb-flight-id')) {
    document.getElementById('wb-flight-id').textContent = window.wbData.flightId || '---';
  }
  if (document.getElementById('wb-route')) {
    let routeText = window.wbData.route || '';
    if (!routeText && window.currentBriefingMission && Array.isArray(window.currentBriefingMission.legs) && window.currentBriefingMission.legs.length) {
      const firstLeg = window.currentBriefingMission.legs[0] || {};
      routeText = `${String(firstLeg.from || '').trim().toUpperCase()} -> ${String(firstLeg.to || '').trim().toUpperCase()}`;
    }
    document.getElementById('wb-route').textContent = routeText || '---';
  }
  if (document.getElementById('wb-aircraft')) {
    document.getElementById('wb-aircraft').textContent = window.wbData.aircraft || '---';
  }

  // Update totals
  document.getElementById('wb-gross-weight').textContent = Math.round(wb.totalWeight) + ' kg';
  document.getElementById('wb-total-moment').textContent = Math.round(wb.totalMoment);
  document.getElementById('wb-cg-position').textContent = (wb.cg || 0).toFixed(1) + ' in';
  document.getElementById('wb-empty-weight').textContent = Math.round(wb.empty.weight) + ' kg';
  document.getElementById('wb-mtow').textContent = Math.round(wb.mtow) + ' kg';
  document.getElementById('wb-weight-remaining').textContent = Math.round(wb.mtow - wb.totalWeight) + ' kg';

  // Update envelope status
  const envelopeEl = document.getElementById('wb-envelope-status');
  if (!hasRealEnvelope) {
    envelopeEl.className = 'envelope-status unsafe';
    envelopeEl.textContent = '[WARN] NO REAL ENVELOPE DATA - FLIGHT MUST NOT PROCEED';
  } else if (isSafe) {
    envelopeEl.className = 'envelope-status safe';
    envelopeEl.textContent = '[OK] WITHIN ENVELOPE';
  } else {
    envelopeEl.className = 'envelope-status unsafe';
    envelopeEl.textContent = '[WARN] OUTSIDE ENVELOPE - CG or Weight Limit Exceeded';
  }

  // Render passenger & cargo summary table
  window.renderPaxCargoSummary();

  // Redraw graph
  window.drawCGGraph();
};

/**
 * Render passenger and cargo summary table
 */
window.renderPaxCargoSummary = function() {
  const tableBody = document.getElementById('wb-pax-cargo-table');
  if (!tableBody) return;

  const rows = [];
  const seats = window.wbData.seats || {};
  
  // Add each passenger (installed seats only), excluding pilot
  Object.keys(seats).forEach(seatId => {
    const seat = seats[seatId];
    // Skip pilot
    if (seat.status === 'installed' && seat.passenger && seat.passenger.name) {
      const isPilot = window.wbData.pilot && seat.passenger.name === window.wbData.pilot;
      if (isPilot) return; // Skip pilot
      
      const planned = seat.passenger.plannedWeight || seat.passenger.weight || 0;
      const actual = seat.passenger.actualWeight || seat.passenger.weight || 0;
      const diff = actual - planned;
      const diffClass = Math.abs(diff) < 0.5 ? 'match' : 'mismatch';
      rows.push(`
        <tr>
          <td>${seat.passenger.name} <span style="color:#999; font-size:1.275rem;">(${seatId})</span></td>
          <td style="text-align:right;">${Math.round(planned)}</td>
          <td style="text-align:right;">${Math.round(actual)}</td>
          <td style="text-align:right; font-weight:bold; color:${Math.abs(diff) < 0.5 ? '#2e7d32' : '#d32f2f'};">${Math.round(diff)}</td>
        </tr>
      `);
    }
  });

  // Add total cargo: planned from dispatch manifest, actual from cargo-area entries
  let cargoPlanned = 0;
  let cargoActual = 0;
  if (window.wbData.cargoManifest) {
    window.wbData.cargoManifest.forEach(cargo => {
      cargoPlanned += parseFloat(cargo.plannedWeight) || 0;
    });
  }
  if (window.wbData.cargoAreas && window.wbData.cargoAreaWeights) {
    window.wbData.cargoAreas.forEach(area => {
      const areaWeights = window.wbData.cargoAreaWeights[area.id];
      if (areaWeights) {
        cargoActual += parseFloat(areaWeights.actual) || 0;
      }
    });
  }
  
  if (cargoPlanned > 0 || cargoActual > 0) {
    const cargoDiff = cargoActual - cargoPlanned;
    rows.push(`
      <tr>
        <td><b>Total Cargo</b></td>
        <td style="text-align:right;">${Math.round(cargoPlanned)}</td>
        <td style="text-align:right;">${Math.round(cargoActual)}</td>
        <td style="text-align:right; font-weight:bold; color:${Math.abs(cargoDiff) < 0.5 ? '#2e7d32' : '#d32f2f'};">${Math.round(cargoDiff)}</td>
      </tr>
    `);
  }

  // Calculate totals
  let totalPlanned = 0, totalActual = 0;
  
  // Passengers total
  Object.keys(seats).forEach(seatId => {
    const seat = seats[seatId];
    if (seat.status === 'installed' && seat.passenger) {
      const isPilot = window.wbData.pilot && seat.passenger.name === window.wbData.pilot;
      if (isPilot) return;
      totalPlanned += (seat.passenger.plannedWeight || seat.passenger.weight) || 0;
      totalActual += (seat.passenger.actualWeight || seat.passenger.weight) || 0;
    }
  });
  
  // Cargo totals
  totalPlanned += cargoPlanned;
  totalActual += cargoActual;

  const totalDiff = totalActual - totalPlanned;
  rows.push(`
    <tr class="summary-total">
      <td>Passengers + Cargo</td>
      <td style="text-align:right;">${Math.round(totalPlanned)}</td>
      <td style="text-align:right;">${Math.round(totalActual)}</td>
      <td style="text-align:right; color:${Math.abs(totalDiff) < 0.5 ? '#2e7d32' : '#d32f2f'};">${Math.round(totalDiff)}</td>
    </tr>
  `);

  tableBody.innerHTML = rows.join('');
};

/**
 * Determine which cargo areas should be visible based on seat configuration
 */
window.getAvailableCargoAreas = function() {
  const cargoAreas = window.wbData.cargoAreas || [];
  const seats = window.wbData.seats || {};
  
  const available = [];
  
  cargoAreas.forEach(area => {
    const areaName = (area.name || '').toUpperCase();
    let shouldShow = false;
    
    if (areaName.includes('POD') || areaName.includes('CARGO D')) {
      // Pods and Cargo D always show
      shouldShow = true;
    } else if (areaName.includes('CARGO A')) {
      // Show only if copilot seat is not installed
      const copilotSeats = Object.keys(seats).filter(k => {
        const s = seats[k];
        return k.toLowerCase().includes('copilot') && s.status === 'installed';
      });
      shouldShow = copilotSeats.length === 0;
    } else if (areaName.includes('CARGO B')) {
      // Show if at least one mid seat is NOT installed
      const totalMidSeats = Object.keys(seats).filter(k => 
        k.toLowerCase().includes('mid') || k.toLowerCase().includes('row2')
      ).length;
      const installedMidSeats = Object.keys(seats).filter(k => {
        const s = seats[k];
        return (k.toLowerCase().includes('mid') || k.toLowerCase().includes('row2')) && s.status === 'installed';
      }).length;
      shouldShow = totalMidSeats > 0 && installedMidSeats < totalMidSeats;
    } else if (areaName.includes('CARGO C')) {
      // Show if at least one aft seat is NOT installed
      const totalAftSeats = Object.keys(seats).filter(k => 
        k.toLowerCase().includes('aft') || k.toLowerCase().includes('row3')
      ).length;
      const installedAftSeats = Object.keys(seats).filter(k => {
        const s = seats[k];
        return (k.toLowerCase().includes('aft') || k.toLowerCase().includes('row3')) && s.status === 'installed';
      }).length;
      shouldShow = totalAftSeats > 0 && installedAftSeats < totalAftSeats;
    }
    
    if (shouldShow) {
      available.push(area);
    }
  });
  
  return available;
};

function wbNextSeatStatus_(status) {
  const cur = String(status || 'installed');
  if (cur === 'installed') return 'cargo';
  if (cur === 'cargo') return 'base';
  return 'installed';
}

function wbStatusButtonMeta_(status) {
  const cur = String(status || 'installed');
  if (cur === 'installed') return { text: 'INSTALLED', css: 'install' };
  if (cur === 'cargo') return { text: 'CARGO', css: 'cargo' };
  return { text: 'BASE', css: 'base' };
}

function wbStatusCycleButtonHtml_(seatId, status) {
  const meta = wbStatusButtonMeta_(status);
  const next = wbNextSeatStatus_(status);
  const nextLabel = wbStatusButtonMeta_(next).text;
  return `
    <button onclick="window.cycleSeatStatus('${seatId}')"
            class="seat-status-btn ${meta.css}"
            title="Now: ${meta.text}. Tap to move to ${nextLabel}.">
      ${meta.text}
    </button>
  `;
}

window.cycleSeatStatus = function(seatName) {
  const seat = window.wbData.seats && window.wbData.seats[seatName];
  if (!seat || seat.locked) return;
  window.changeSeatStatus(seatName, wbNextSeatStatus_(seat.status));
};

/**
 * Render manifest rows (includes seats)
 */
window.renderManifest = function() {
  const container = document.getElementById('wb-manifest-rows');
  
  // Guard against missing DOM element in case of layout changes
  if (!container) {
    console.warn('Manifest container not found - layout may have changed');
    return;
  }
  
  // Get ALL passengers from mission (not just unassigned)
  const allMissionPassengers = [];
  const seenPassengers = {};
  
  // First, get passengers from all seat assignments
  Object.values(window.wbData.seatAssignments || {}).forEach(s => {
    if (s.passenger && s.passenger.name && !s.locked && s.passenger.name !== window.wbData.pilot) {
      const key = s.passenger.name;
      if (!seenPassengers[key]) {
        seenPassengers[key] = true;
        allMissionPassengers.push(s.passenger);
      }
    }
  });
  
  // Add waiting passengers not already in the list
  window.wbData.waitingPassengers.forEach(p => {
    const key = p.name;
    if (!seenPassengers[key]) {
      seenPassengers[key] = true;
      allMissionPassengers.push(p);
    }
  });
  
  const rows = [];
  
  // 1. Empty Aircraft
  const emptyItem = window.wbData.items.find(item => item.type === 'empty');
  if (emptyItem) {
    const weight = parseFloat(emptyItem.actualWeight) || 0;
    rows.push(`
      <div class="manifest-row">
        <div><b>${emptyItem.name}</b></div>
        <div>${weight.toFixed(1)}</div>
        <div>${(parseFloat(emptyItem.arm) || 0).toFixed(1)}</div>
        <div></div>
      </div>
    `);
  }
  
  // Separate seats by status
  const installedSeats = [];
  const cargoSeats = [];
  const baseSeats = [];
  
  Object.keys(window.wbData.seats).forEach(seatId => {
    const seat = window.wbData.seats[seatId];
    const status = seat.status || 'installed';
    if (status === 'installed') {
      installedSeats.push({ seatId, seat });
    } else if (status === 'cargo') {
      cargoSeats.push({ seatId, seat });
    } else if (status === 'base') {
      baseSeats.push({ seatId, seat });
    }
  });

  const SEAT_ORDER = ['Pilot Seat', 'Copilot Seat', 'RH Mid Seat', 'LH Mid Seat', 'RH Aft Seat', 'LH Aft Seat'];
  installedSeats.sort(function(a, b) {
    const ai = SEAT_ORDER.indexOf(a.seatId);
    const bi = SEAT_ORDER.indexOf(b.seatId);
    if (ai !== -1 && bi !== -1) return ai - bi;
    if (ai !== -1) return -1;
    if (bi !== -1) return 1;
    return String(a.seatId || '').localeCompare(String(b.seatId || ''));
  });
  
  // 2. All Installed Seats
  installedSeats.forEach(({ seatId, seat }) => {
    const seatWeight = seat.weight || 0;
    const occupantWeight = seat.occupiedWeight || 0;

    // Single-line seat row: seat + seat weight + first-name occupant.
    let nameHTML = `<b>${seatId} <span style="color:#666; font-weight:700;">(${seatWeight}kg)</span></b>`;
    if (seat.passenger && seat.passenger.name) {
      nameHTML += ` <span style="color:#455a64;">${wbFirstName_(seat.passenger.name)}</span>`;
    } else if (!seat.locked) {
      nameHTML += ` <span style="color:#999;">Empty</span>`;
    }
    
    // Planned weight column - original dispatch weight
    let plannedWeightHTML = '';
    let actualWeightHTML = '';
    let deltaHTML = '';
    
    if (seat.passenger && seat.passenger.name) {
      // Store original planned weight if not already stored
      if (!seat.passenger.plannedWeight) {
        seat.passenger.plannedWeight = seat.passenger.weight;
      }
      
      const plannedWeight = seat.passenger.plannedWeight || 0;
      const actualWeight = seat.passenger.actualWeight || seat.passenger.weight;
      const isVerified = seat.passenger.verified || false;
      const isPilotSeat = window.wbData.pilot && seat.passenger.name === window.wbData.pilot;
      const diff = actualWeight - plannedWeight;
      const diffClass = Math.abs(diff) < 0.5 ? 'match' : 'mismatch';
      
      plannedWeightHTML = `<span>${plannedWeight.toFixed(1)}</span>`;
      
      // Actual weight with verify/edit/lock controls (pilot is fixed, non-editable)
      if (isPilotSeat) {
        actualWeightHTML = `
          <div style="display:flex; align-items:center; gap:4px;">
            <span id="weight-display-${seatId}" style="font-weight:600; color:#2e7d32;">${actualWeight.toFixed(1)}</span>
          </div>
        `;
      } else {
        actualWeightHTML = `
        <div style="display:flex; align-items:center; gap:4px;">
          <span id="weight-display-${seatId}" style="font-weight:${isVerified ? '600' : 'normal'}; color:${isVerified ? '#2e7d32' : '#333'}; ${isVerified ? '' : 'display:none;'}">
            ${actualWeight.toFixed(1)}
          </span>
          <input id="weight-input-${seatId}" type="number" inputmode="decimal" step="0.1" value="${actualWeight}" 
                 style="border-color:#2196f3; ${isVerified ? 'display:none;' : 'display:inline-block; width:62px;'}"
                 onkeydown="if(event.key==='Enter') window.savePassengerWeight('${seatId}')">
          <button onclick="window.verifyPassengerWeight('${seatId}')" 
                  id="verify-btn-${seatId}"
                  title="Verify weight"
                  class="wb-manifest-btn verify" ${isVerified ? 'style="display:none;"' : ''}>
            VERIFY
          </button>
          <button onclick="window.editPassengerWeight('${seatId}')"
                  id="unlock-btn-${seatId}"
                  title="Unlock to edit"
                  class="wb-lock-btn" ${isVerified ? '' : 'style="display:none;"'}>
            🔒
          </button>
          <span id="verified-chip-${seatId}" class="wb-verified-chip" ${isVerified ? '' : 'style="display:none;"'}>
            <span style="font-size:0.9rem; line-height:1;">✓</span>
            <span>OK</span>
          </span>
          <button onclick="window.savePassengerWeight('${seatId}')" 
                  id="save-btn-${seatId}"
                  title="Save weight"
                  class="wb-manifest-btn save" style="display:none;">
            Save
          </button>
          <button onclick="window.cancelEditWeight('${seatId}')" 
                  id="cancel-btn-${seatId}"
                  title="Cancel"
                  class="wb-manifest-btn cancel" style="display:none;">
            Cancel
          </button>
        </div>
      `;
      }
      
      deltaHTML = `<span class="weight-diff ${diffClass}">${diff >= 0 ? '+' : ''}${diff.toFixed(1)}</span>`;
    } else {
      plannedWeightHTML = `<span>-</span>`;
      actualWeightHTML = `<span>-</span>`;
      deltaHTML = `<span>-</span>`;
    }
    
    // Controls column
    let controlsHTML = '';
    if (seat.locked) {
      controlsHTML = '';
    } else {
      controlsHTML = `<div style="display:flex; justify-content:flex-end;">${wbStatusCycleButtonHtml_(seatId, seat.status)}</div>`;
    }
    
    rows.push(`
      <div class="manifest-row">
        <div>${nameHTML}</div>
        <div>${actualWeightHTML}</div>
        <div>${(parseFloat(seat.arm) || 0).toFixed(1)}</div>
        <div>${controlsHTML}</div>
      </div>
    `);
  });
  
  // 3. Fuel
  const fuelItem = window.wbData.items.find(item => item.type === 'fuel');
  if (fuelItem) {
    const actualWeight = parseFloat(fuelItem.actualWeight) || 0;
    const roundedFuelWeight = Math.round(actualWeight);
    const fuelLiters = actualWeight > 0 ? Math.round(actualWeight / 0.72) : 0;
    rows.push(`
      <div class="manifest-row">
        <div><b>${fuelItem.name} (${fuelLiters} L)</b></div>
        <div><span style="font-weight:bold; color:#0b5394;">${roundedFuelWeight}</span></div>
        <div>${(parseFloat(fuelItem.arm) || 0).toFixed(1)}</div>
        <div></div>
      </div>
    `);
  }
  
  // 4. Conditional cargo areas (Pod A, Pod B, Cargo A, B, C)
  const cargoSortOrder = ['POD A', 'POD B', 'CARGO A', 'CARGO B', 'CARGO C', 'CARGO D'];
  const availableCargoAreas = window.getAvailableCargoAreas().slice().sort(function(a, b) {
    const aName = String(a && a.name || '').toUpperCase();
    const bName = String(b && b.name || '').toUpperCase();
    const aIdx = cargoSortOrder.findIndex(function(k) { return aName.indexOf(k) >= 0; });
    const bIdx = cargoSortOrder.findIndex(function(k) { return bName.indexOf(k) >= 0; });
    const safeA = aIdx >= 0 ? aIdx : 999;
    const safeB = bIdx >= 0 ? bIdx : 999;
    if (safeA !== safeB) return safeA - safeB;
    return aName.localeCompare(bName);
  });
  availableCargoAreas.forEach(area => {
    // Initialize cargo areas in wbData if not present
    if (!window.wbData.cargoAreaWeights) {
      window.wbData.cargoAreaWeights = {};
    }
    if (window.wbData.cargoAreaWeights[area.id] === undefined) {
      window.wbData.cargoAreaWeights[area.id] = { planned: 0, actual: 0 };
    }
    const areaWeights = window.wbData.cargoAreaWeights[area.id];
    const plannedWeight = areaWeights.planned || 0;
    const actualWeight = areaWeights.actual || 0;
    const diff = actualWeight - plannedWeight;
    const diffClass = Math.abs(diff) < 0.5 ? 'match' : 'mismatch';
    const maxWeightKg = parseFloat(area.maxWeightKg) || 0;
    const maxExceeded = maxWeightKg > 0 && actualWeight > maxWeightKg;
    const maxDisplay = maxWeightKg > 0 ? ` (Max: ${maxWeightKg.toFixed(1)}kg)` : '';
    
    let areaLabel = `<b>${area.name}</b>${maxDisplay}`;
    if (maxExceeded) {
      areaLabel += ` <span style="color:#d32f2f; font-weight:bold; font-size:1.26rem;">[X] EXCEEDS MAX</span>`;
    }
    
    rows.push(`
      <div class="manifest-row">
        <div>${areaLabel}</div>
        <div>
          <input type="number"
                 inputmode="decimal"
                 value="${actualWeight}"
                 step="0.1"
                 title="${maxDisplay}"
                 class="manifest-cargo-input"
                 onchange="
                   const val = parseFloat(this.value) || 0;
                   const max = ${maxWeightKg} || 0;
                   if (max > 0 && val > max) {
                     M.toast({html: '[X] Exceeds max weight of ' + Math.round(max) + 'kg', displayLength: 3000});
                     this.value = '${actualWeight}';
                     return;
                   }
                   if (!window.wbData.cargoAreaWeights) window.wbData.cargoAreaWeights = {};
                   window.wbData.cargoAreaWeights['${area.id}'] = { planned: ${plannedWeight}, actual: val };
                   window.updateWBUI();
                   window.renderManifest();
                   if (typeof window.persistWBState_ === 'function') window.persistWBState_();
                 "
                 style="border-color:${maxExceeded ? '#d32f2f' : '#2196f3'}; background:${maxExceeded ? '#ffebee' : 'white'};">
        </div>
        <div>${(parseFloat(area.arm) || 0).toFixed(1)}</div>
        <div></div>
      </div>
    `);
  });
  
  // 4b. Add "Cargo D" for stowed seats
  const storedSeats = Object.keys(window.wbData.seats)
    .filter(seatId => window.wbData.seats[seatId].status === 'cargo')
    .map(seatId => window.wbData.seats[seatId]);
  
  if (storedSeats.length > 0) {
    const storedWeight = storedSeats.reduce((sum, seat) => sum + (seat.weight || 0), 0);
    const storedArm = storedSeats.reduce((s, seat) => s + (seat.arm || 0), 0) / storedSeats.length;
    
    rows.push(`
      <div class="manifest-row" style="background:#fff8e1;">
        <div><b>Cargo D (Stowed)</b></div>
        <div><span style="background:#2196f3; color:white; padding:2px 4px; border-radius:2px; font-size:1.17rem;">${storedWeight.toFixed(1)}</span></div>
        <div>${storedArm.toFixed(1)}</div>
        <div><span style="color:#999; font-size:1.2rem;">Auto</span></div>
      </div>
    `);
  }
  
  // 5. Cargo Seats (at bottom)
  cargoSeats.forEach(({ seatId, seat }) => {
    const seatWeight = seat.weight || 0;
    const occupantWeight = seat.occupiedWeight || 0;

    let nameHTML = `<b>${seatId} <span style="color:#666; font-weight:700;">(${seatWeight}kg)</span></b>`;
    if (seat.passenger && seat.passenger.name) {
      nameHTML += ` <span style="color:#455a64;">${wbFirstName_(seat.passenger.name)}</span>`;
    }
    
    const controlsHTML = seat.locked
      ? ''
      : `<div style="display:flex; justify-content:flex-end;">${wbStatusCycleButtonHtml_(seatId, seat.status)}</div>`;
    
    rows.push(`
      <div class="manifest-row">
        <div>${nameHTML}</div>
        <div>-</div>
        <div>${(parseFloat(seat.arm) || 0).toFixed(1)}</div>
        <div>${controlsHTML}</div>
      </div>
    `);
  });
  
  // 6. Base Seats (at very bottom)
  if (baseSeats.length > 0) {
    rows.push(`
      <div class="manifest-row" style="background:#f5f5f5; margin-top:8px;">
        <div style="grid-column:1/-1"><b style="color:#666;">LEFT AT BASE</b></div>
      </div>
    `);
  }
  
  baseSeats.forEach(({ seatId, seat }) => {
    const seatWeight = seat.weight || 0;

    let nameHTML = `<b>${seatId} <span style="color:#666; font-weight:700;">(${seatWeight}kg)</span></b>`;
    
    const controlsHTML = seat.locked
      ? ''
      : `<div style="display:flex; justify-content:flex-end;">${wbStatusCycleButtonHtml_(seatId, seat.status)}</div>`;
    
    rows.push(`
      <div class="manifest-row">
        <div>${nameHTML}</div>
        <div>-</div>
        <div>${(parseFloat(seat.arm) || 0).toFixed(1)}</div>
        <div>${controlsHTML}</div>
      </div>
    `);
  });
  
  container.innerHTML = rows.join('');

  // Render waiting passengers section
  const waitingContainer = document.getElementById('wb-waiting-passengers');
  const waitingList = document.getElementById('wb-waiting-list');
  
  if (waitingContainer && waitingList && window.wbData.waitingPassengers && window.wbData.waitingPassengers.length > 0) {
    waitingContainer.style.display = 'block';
    const waitingHTML = window.wbData.waitingPassengers.map(p => `
      <div style="display:flex; flex-wrap:wrap; justify-content:space-between; align-items:center; gap:6px 8px; padding:6px; background:white; border-radius:3px;">
        <span style="font-weight:600; font-size:1.425rem;">${p.name} <span style="color:#999; font-size:1.35rem;">(${p.weight}kg)</span></span>
        <div style="display:flex; flex-wrap:wrap; gap:4px; justify-content:flex-end;">
          ${Object.keys(window.wbData.seats).filter(seatId => {
            const s = window.wbData.seats[seatId];
            return !s.locked && s.status === 'installed' && !s.passenger;
          }).map(seatId => `
            <button onclick="window.assignPassengerToSeat('${seatId}', '${p.name}')" 
                    class="wb-manifest-btn edit">
              Assign ${seatId}
            </button>
          `).join('')}
        </div>
      </div>
    `).join('');
    
    waitingList.innerHTML = waitingHTML;
  } else if (waitingContainer && waitingList) {
    waitingContainer.style.display = 'none';
    waitingList.innerHTML = '';
  }

  // Update totals using the same seat-aware WB model used by summary/graph.
  const wbTotals = (typeof window.calculateWB === 'function') ? window.calculateWB() : null;
  const actualTotal = wbTotals ? Number(wbTotals.totalWeight || 0) : 0;
  document.getElementById('wb-total-actual').textContent = String(Math.round(actualTotal));
  
  // Hide the old seat toggle section since seats are now in manifest
  const togglesContainer = document.getElementById('wb-seat-toggles');
  if (togglesContainer) {
    togglesContainer.style.display = 'none';
  }
};

/**
 * Verify passenger weight
 */
window.verifyPassengerWeight = function(seatId) {
  const seat = window.wbData.seats[seatId];
  if (!seat || !seat.passenger) return;

  const inputEl = document.getElementById(`weight-input-${seatId}`);
  const typedWeight = wbNum_(inputEl && inputEl.value, 0);
  if (!(typedWeight > 0)) {
    if (window.M) M.toast({ html: 'Enter a valid weight before verify', classes: 'orange', displayLength: 2000 });
    return;
  }
  seat.passenger.actualWeight = typedWeight;
  seat.passenger.verified = true;
  
  // Update occupiedWeight to use actual weight
  seat.occupiedWeight = seat.passenger.actualWeight;
  
  // Update UI
  window.renderManifest();
  window.updateWBUI();
  if (typeof window.persistWBState_ === 'function') window.persistWBState_();
  
  M.toast({html: `Weight verified for ${seat.passenger.name}`, displayLength: 2000});
};

/**
 * Edit passenger weight
 */
window.editPassengerWeight = function(seatId) {
  const seat = window.wbData.seats[seatId];
  if (!seat || !seat.passenger) return;
  
  // Show input, hide display
  const displayEl = document.getElementById(`weight-display-${seatId}`);
  const inputEl = document.getElementById(`weight-input-${seatId}`);
  const verifyBtn = document.getElementById(`verify-btn-${seatId}`);
  const unlockBtn = document.getElementById(`unlock-btn-${seatId}`);
  const verifiedChip = document.getElementById(`verified-chip-${seatId}`);
  const saveBtn = document.getElementById(`save-btn-${seatId}`);
  const cancelBtn = document.getElementById(`cancel-btn-${seatId}`);
  
  if (displayEl) displayEl.style.display = 'none';
  if (inputEl) {
    inputEl.style.display = 'inline-block';
    inputEl.style.width = '62px';
    inputEl.value = seat.passenger.actualWeight || seat.passenger.weight;
    inputEl.focus();
    inputEl.select();
  }
  if (verifyBtn) verifyBtn.style.display = 'none';
  if (unlockBtn) unlockBtn.style.display = 'none';
  if (verifiedChip) verifiedChip.style.display = 'none';
  if (saveBtn) saveBtn.style.display = 'inline-block';
  if (cancelBtn) cancelBtn.style.display = 'inline-block';
  if (seat.passenger) seat.passenger.verified = false;
};

/**
 * Save edited passenger weight
 */
window.savePassengerWeight = function(seatId) {
  const seat = window.wbData.seats[seatId];
  if (!seat || !seat.passenger) return;
  
  const inputEl = document.getElementById(`weight-input-${seatId}`);
  const newWeight = parseFloat(inputEl.value);
  
  if (isNaN(newWeight) || newWeight <= 0) {
    M.toast({html: 'Invalid weight value', displayLength: 2000});
    return;
  }
  
  // Update passenger actual weight
  seat.passenger.actualWeight = newWeight;
  seat.occupiedWeight = newWeight;
  seat.passenger.verified = true;
  
  // Rebuild manifest items with new actual weight
  const newItems = [];
  window.wbData.items.forEach(item => {
    if (item.type !== 'passenger') {
      newItems.push(item);
    }
  });
  
  Object.keys(window.wbData.seats).forEach(seatName => {
    const s = window.wbData.seats[seatName];
    if (s.status === 'installed' && s.occupiedWeight > 0 && s.passenger) {
      const paxActualWeight = s.passenger.actualWeight || s.passenger.weight;
      newItems.push({
        name: seatName + ': ' + s.passenger.name,
        plannedWeight: s.passenger.plannedWeight || s.passenger.weight,
        actualWeight: paxActualWeight,
        arm: s.arm,
        type: 'passenger',
        seatId: seatName
      });
    }
  });
  
  window.wbData.items = newItems;
  
  // Refresh UI
  window.renderManifest();
  window.updateWBUI();
  if (typeof window.persistWBState_ === 'function') window.persistWBState_();
  
  M.toast({html: `Weight updated for ${seat.passenger.name}: ${newWeight}kg`, displayLength: 2000});
};

/**
 * Cancel weight edit
 */
window.cancelEditWeight = function(seatId) {
  const seat = window.wbData.seats[seatId];
  if (!seat || !seat.passenger) return;
  
  // Reset input to current actual value
  const inputEl = document.getElementById(`weight-input-${seatId}`);
  if (inputEl) inputEl.value = seat.passenger.actualWeight || seat.passenger.weight;
  
  // Show display, hide input
  const displayEl = document.getElementById(`weight-display-${seatId}`);
  const verifyBtn = document.getElementById(`verify-btn-${seatId}`);
  const unlockBtn = document.getElementById(`unlock-btn-${seatId}`);
  const verifiedChip = document.getElementById(`verified-chip-${seatId}`);
  const saveBtn = document.getElementById(`save-btn-${seatId}`);
  const cancelBtn = document.getElementById(`cancel-btn-${seatId}`);
  
  if (displayEl) displayEl.style.display = 'inline';
  if (inputEl) inputEl.style.display = 'none';
  if (saveBtn) saveBtn.style.display = 'none';
  if (cancelBtn) cancelBtn.style.display = 'none';
  
  if (seat.passenger.verified) {
    if (verifyBtn) verifyBtn.style.display = 'none';
    if (unlockBtn) unlockBtn.style.display = 'inline-flex';
    if (verifiedChip) verifiedChip.style.display = 'inline-flex';
    if (displayEl) displayEl.style.display = 'inline';
    if (inputEl) inputEl.style.display = 'none';
  } else {
    if (displayEl) displayEl.style.display = 'none';
    if (inputEl) {
      inputEl.style.display = 'inline-block';
      inputEl.style.width = '62px';
    }
    if (verifyBtn) verifyBtn.style.display = 'inline-block';
    if (unlockBtn) unlockBtn.style.display = 'none';
    if (verifiedChip) verifiedChip.style.display = 'none';
  }
};

/**
 * Render seat toggles with status and passenger info
 */
window.renderSeatToggles = function() {
  const container = document.getElementById('wb-seat-toggles');
  
  // Element no longer exists in new layout - seats are in manifest instead
  if (!container) {
    return;
  }
  
  if (!window.wbData.seats || Object.keys(window.wbData.seats).length === 0) {
    container.innerHTML = '<p>No seats configured</p>';
    return;
  }

  // Filter out locked seats (pilot seat)
  const html = Object.keys(window.wbData.seats)
    .filter(seatName => !window.wbData.seats[seatName].locked)
    .map(seatName => {
    const seat = window.wbData.seats[seatName];
    const status = seat.status || 'installed';
    const statusClass = 'status-' + status;
    
    let passengerInfo = '';
    if (seat.passenger && seat.passenger.name) {
      // Get all mission passengers (from seatAssignments)
      const allPassengers = [];
      Object.values(window.wbData.seatAssignments || {}).forEach(s => {
        if (s.passenger && s.passenger.name && s.label !== 'Pilot Seat') {
          const key = s.passenger.name;
          if (!allPassengers.find(p => p.name === key)) {
            allPassengers.push(s.passenger);
          }
        }
      });
      
      const paxDropdown = `
        <select class="pax-reassign-select" onchange="window.reassignPassenger('${seatName}', this.value)">
          <option value="">-- Reassign passenger --</option>
          ${allPassengers.map(p => `<option value="${p.name}" ${p.name === seat.passenger.name ? 'selected' : ''}>${p.name} (${p.weight}kg)</option>`).join('')}
        </select>
      `;
      
      passengerInfo = `
        <div class="seat-passenger-info">
          Pax: ${seat.passenger.name} (${seat.occupiedWeight || 0}kg)
          ${paxDropdown}
        </div>
      `;
    } else if (status === 'installed') {
      // Empty installed seat - show dropdown to assign passenger
      const allPassengers = [];
      Object.values(window.wbData.seatAssignments || {}).forEach(s => {
        if (s.passenger && s.passenger.name && s.label !== 'Pilot Seat') {
          const key = s.passenger.name;
          if (!allPassengers.find(p => p.name === key)) {
            allPassengers.push(s.passenger);
          }
        }
      });
      
      const emptyDropdown = `
        <select class="pax-reassign-select" onchange="window.assignPassengerToSeat('${seatName}', this.value)">
          <option value="">-- Assign passenger --</option>
          ${allPassengers.map(p => `<option value="${p.name}">${p.name} (${p.weight}kg)</option>`).join('')}
        </select>
      `;
      
      passengerInfo = `
        <div class="seat-passenger-info" style="color:#999;">
          Empty seat
          ${emptyDropdown}
        </div>
      `;
    }

    // Hide controls for locked seats (e.g., pilot)
    const controlsHTML = seat.locked
      ? `<div class="seat-controls" style="color:#999; font-size:1.275rem; font-style:italic;">Locked - cannot modify</div>`
      : `<div class="seat-controls">
          <button onclick="window.changeSeatStatus('${seatName}', 'installed')" 
                  ${status === 'installed' ? 'disabled' : ''}>
            Install
          </button>
          <button onclick="window.changeSeatStatus('${seatName}', 'cargo')" 
                  class="secondary"
                  ${status === 'cargo' ? 'disabled' : ''}>
            To Cargo
          </button>
          <button onclick="window.changeSeatStatus('${seatName}', 'base')" 
                  class="secondary"
                  ${status === 'base' ? 'disabled' : ''}>
            Leave at Base
          </button>
        </div>`;

    return `
      <div class="seat-toggle ${statusClass}">
        <div class="seat-toggle-header">
          <span>${seatName} (${seat.weight}kg)</span>
          <span class="seat-badge ${status}">${status}</span>
        </div>
        ${passengerInfo}
        ${controlsHTML}
      </div>
    `;
  }).join('');

  container.innerHTML = html;
  
  // Add summary info
  const installedCount = Object.values(window.wbData.seats).filter(s => s.status === 'installed').length;
  const cargoCount = Object.values(window.wbData.seats).filter(s => s.status === 'cargo').length;
  const baseCount = Object.values(window.wbData.seats).filter(s => s.status === 'base').length;
  
  const summary = `
    <div style="margin-top:8px; padding:6px; background:#f5f5f5; border-radius:4px; font-size:1.2rem;">
      <b>Seat Summary:</b> 
      ${installedCount} installed, 
      ${cargoCount} in cargo, 
      ${baseCount} left at base
      ${window.wbData.maxPaxInMission ? `<br><i>Mission max passengers: ${window.wbData.maxPaxInMission}</i>` : ''}
    </div>
  `;
  
  container.innerHTML += summary;
};

/**
 * Change seat status and recalculate W&B
 */
window.changeSeatStatus = function(seatName, newStatus) {
  const seat = window.wbData.seats[seatName];
  if (!seat) return;
  
  const oldStatus = seat.status;
  if (oldStatus === newStatus) return;
  
  // If moving away from installed and there's a passenger, move them to waiting
  if (oldStatus === 'installed' && newStatus !== 'installed' && seat.passenger) {
    // Check if this passenger is already in waiting list
    const alreadyWaiting = window.wbData.waitingPassengers.find(p => p.name === seat.passenger.name);
    if (!alreadyWaiting) {
      window.wbData.waitingPassengers.push({ ...seat.passenger });
    }
  }
  
  // Clear passenger from seat if moving to non-installed status
  if (newStatus !== 'installed') {
    seat.passenger = null;
    seat.occupiedWeight = 0;
  }
  
  // Update seat status
  seat.status = newStatus;
  seat.enabled = (newStatus === 'installed');
  
  // Rebuild items array to reflect seat changes
  // Remove all passenger items and rebuild from current seat state
  const newItems = [];
  window.wbData.items.forEach(item => {
    // Keep non-passenger items
    if (item.type !== 'passenger') {
      newItems.push(item);
    }
  });
  
  // Add passenger items only from installed seats
  Object.keys(window.wbData.seats).forEach(seatId => {
    const s = window.wbData.seats[seatId];
    if (s.status === 'installed' && s.passenger && s.occupiedWeight > 0) {
      const passengerName = s.passenger.name;
      const plannedWeight = s.passenger.plannedWeight || s.passenger.weight;
      const actualWeight = s.passenger.actualWeight || s.passenger.weight;
      const passengerMoment = actualWeight * s.arm;
      
      newItems.push({
        name: `${seatId}: ${passengerName}`,
        plannedWeight: plannedWeight,
        actualWeight: actualWeight,
        arm: s.arm,
        moment: passengerMoment,
        type: 'passenger',
        seatId: seatId
      });
    }
  });
  
  // Update empty aircraft weight to account for seats left at base
  let adjustedEmptyWeight = window.wbData.airframeData.Empty_Weight;
  Object.keys(window.wbData.seats).forEach(seatId => {
    const s = window.wbData.seats[seatId];
    if (s.status === 'base') {
      adjustedEmptyWeight -= s.weight;
    }
  });
  
  const emptyItem = newItems.find(item => item.type === 'empty');
  if (emptyItem) {
    emptyItem.plannedWeight = adjustedEmptyWeight;
    emptyItem.actualWeight = adjustedEmptyWeight;
  }
  
  window.wbData.items = newItems;
  
  // Refresh UI
  window.renderSeatToggles();
  window.renderAircraftDiagram();
  window.renderCargoPodDiagram();
  window.renderManifest();
  window.updateWBUI();
  if (typeof window.persistWBState_ === 'function') window.persistWBState_();
  
  M.toast({html: `${seatName}: moved to ${newStatus}`, displayLength: 2000});
};

/**
 * Reassign passenger to different seat
 */
window.reassignPassenger = function(targetSeatName, passengerName) {
  if (!passengerName) return;
  
  // Find the seat that currently has this passenger
  let sourceSeatName = null;
  
  Object.keys(window.wbData.seats).forEach(sName => {
    const s = window.wbData.seats[sName];
    if (s.passenger && s.passenger.name === passengerName) {
      sourceSeatName = sName;
    }
  });
  
  if (!sourceSeatName || sourceSeatName === targetSeatName) return;
  const targetSeat = window.wbData.seats[targetSeatName];
  const sourceSeat = window.wbData.seats[sourceSeatName];
  
  // Swap passengers and weights
  const tempPassenger = targetSeat.passenger;
  const tempWeight = targetSeat.occupiedWeight;
  
  targetSeat.passenger = sourceSeat.passenger;
  targetSeat.occupiedWeight = sourceSeat.occupiedWeight;
  
  sourceSeat.passenger = tempPassenger;
  sourceSeat.occupiedWeight = tempWeight;
  
  // Rebuild manifest items from current seat state
  const newItems = [];
  window.wbData.items.forEach(item => {
    if (item.type !== 'passenger') {
      newItems.push(item);
    }
  });
  
  // Add passenger items from installed seats
  Object.keys(window.wbData.seats).forEach(seatName => {
    const seat = window.wbData.seats[seatName];
    if (seat.status === 'installed' && seat.occupiedWeight > 0 && seat.passenger) {
      const paxPlanned = seat.passenger.plannedWeight || seat.passenger.weight;
      const paxActual = seat.passenger.actualWeight || seat.passenger.weight;
      newItems.push({
        name: seatName + ': ' + seat.passenger.name,
        plannedWeight: paxPlanned,
        actualWeight: paxActual,
        arm: seat.arm,
        type: 'passenger',
        seatId: seatName
      });
    }
  });
  
  window.wbData.items = newItems;
  
  // Refresh UI
  window.renderSeatToggles();
  window.renderAircraftDiagram();
  window.renderCargoPodDiagram();
  window.renderManifest();
  window.updateWBUI();
  if (typeof window.persistWBState_ === 'function') window.persistWBState_();
  
  M.toast({html: `${passengerName} reassigned to ${targetSeatName}`, displayLength: 2000});
};

/**
 * Assign a passenger to an empty seat
 */
window.assignPassengerToSeat = function(seatName, passengerName) {
  if (!passengerName) return;
  
  const targetSeat = window.wbData.seats[seatName];
  if (!targetSeat) return;
  
  // Find this passenger in any seat or in waiting list
  let sourceSeatName = null;
  let passengerData = null;
  let isFromWaiting = false;
  
  // Check seats first
  Object.keys(window.wbData.seats).forEach(seatId => {
    const s = window.wbData.seats[seatId];
    if (s.passenger && s.passenger.name === passengerName) {
      sourceSeatName = seatId;
      passengerData = s.passenger;
    }
  });
  
  // If passenger not found in any seat, check waiting list
  if (!passengerData) {
    const waitingIndex = window.wbData.waitingPassengers.findIndex(p => p.name === passengerName);
    if (waitingIndex >= 0) {
      passengerData = window.wbData.waitingPassengers[waitingIndex];
      isFromWaiting = true;
    }
  }
  
  // If still not found, look in seatAssignments
  if (!passengerData) {
    Object.values(window.wbData.seatAssignments || {}).forEach(s => {
      if (s.passenger && s.passenger.name === passengerName) {
        passengerData = s.passenger;
      }
    });
  }
  
  if (!passengerData) return;
  
  // Initialize plannedWeight and actualWeight if not set
  if (!passengerData.plannedWeight) {
    passengerData.plannedWeight = passengerData.weight;
  }
  if (!passengerData.actualWeight) {
    passengerData.actualWeight = passengerData.weight;
  }
  
  // Assign to target seat
  targetSeat.passenger = passengerData;
  targetSeat.occupiedWeight = passengerData.actualWeight || passengerData.weight;
  
  // Remove from source seat if there was one
  if (sourceSeatName) {
    const sourceSeat = window.wbData.seats[sourceSeatName];
    sourceSeat.passenger = null;
    sourceSeat.occupiedWeight = 0;
  }
  
  // Remove from waiting list if it was from there
  if (isFromWaiting) {
    window.wbData.waitingPassengers = window.wbData.waitingPassengers.filter(p => p.name !== passengerName);
  }
  
  // Rebuild manifest
  const newItems = [];
  window.wbData.items.forEach(item => {
    if (item.type !== 'passenger') {
      newItems.push(item);
    }
  });
  
  Object.keys(window.wbData.seats).forEach(seatId => {
    const seat = window.wbData.seats[seatId];
    if (seat.status === 'installed' && seat.occupiedWeight > 0 && seat.passenger) {
      const paxPlanned = seat.passenger.plannedWeight || seat.passenger.weight;
      const paxActual = seat.passenger.actualWeight || seat.passenger.weight;
      newItems.push({
        name: seatId + ': ' + seat.passenger.name,
        plannedWeight: paxPlanned,
        actualWeight: paxActual,
        arm: seat.arm,
        type: 'passenger',
        seatId: seatId
      });
    }
  });
  
  window.wbData.items = newItems;
  
  // Refresh UI
  window.renderSeatToggles();
  window.renderAircraftDiagram();
  window.renderCargoPodDiagram();
  window.renderManifest();
  window.updateWBUI();
  if (typeof window.persistWBState_ === 'function') window.persistWBState_();
  
  M.toast({html: `${passengerName} assigned to ${seatName}`, displayLength: 2000});
};

window.moveSeatPassengerToWaiting = function(seatName) {
  const seat = window.wbData.seats && window.wbData.seats[seatName];
  if (!seat || seat.locked || !seat.passenger) return;

  const passenger = { ...seat.passenger };
  seat.passenger = null;
  seat.occupiedWeight = 0;

  if (!Array.isArray(window.wbData.waitingPassengers)) {
    window.wbData.waitingPassengers = [];
  }
  const alreadyWaiting = window.wbData.waitingPassengers.some(function(p) {
    return wbNormName_(p && p.name).toUpperCase() === wbNormName_(passenger && passenger.name).toUpperCase();
  });
  if (!alreadyWaiting) {
    window.wbData.waitingPassengers.push(passenger);
  }

  const newItems = [];
  window.wbData.items.forEach(function(item) {
    if (item.type !== 'passenger') newItems.push(item);
  });
  Object.keys(window.wbData.seats).forEach(function(seatId) {
    const s = window.wbData.seats[seatId];
    if (s.status === 'installed' && s.occupiedWeight > 0 && s.passenger) {
      const paxPlanned = s.passenger.plannedWeight || s.passenger.weight;
      const paxActual = s.passenger.actualWeight || s.passenger.weight;
      newItems.push({
        name: seatId + ': ' + s.passenger.name,
        plannedWeight: paxPlanned,
        actualWeight: paxActual,
        arm: s.arm,
        type: 'passenger',
        seatId: seatId
      });
    }
  });
  window.wbData.items = newItems;

  window.renderAircraftDiagram();
  window.renderManifest();
  window.updateWBUI();
  if (typeof window.persistWBState_ === 'function') window.persistWBState_();
  if (window.M) M.toast({ html: passenger.name + ' moved to needing to be seated', classes: 'blue' });
};

/**
 * Render aircraft diagram with seat positions
 */
window.renderAircraftDiagram = function() {
  const container = document.getElementById('wb-aircraft-diagram');
  if (!container) return;
  
  // Get available cargo areas
  const availableCargoAreas = window.getAvailableCargoAreas();
  const availableCargoIds = availableCargoAreas.map(a => a.id);

  // Helper to get seat display text
  const getSeatDisplay = (seatName, seat) => {
    if (seat.passenger && seat.passenger.name) {
      return seat.passenger.name.split(' ')[0].substring(0, 12);
    }
    return '';
  };

  // Helper to get seat CSS class
  const getSeatClass = (seat) => {
    if (seat.status === 'installed') return 'installed';
    if (seat.status === 'cargo') return 'cargo';
    if (seat.status === 'base') return 'empty';
    return 'empty';
  };

  // Helper to render a seat or cargo zone
  const seatTitleFor = (seatName) => {
    const map = {
      'Pilot Seat': 'PILOT',
      'Copilot Seat': 'COPILOT',
      'LH Mid Seat': 'LH MID',
      'RH Mid Seat': 'RH MID',
      'LH Aft Seat': 'LH AFT',
      'RH Aft Seat': 'RH AFT'
    };
    return map[seatName] || String(seatName || '').replace(' Seat', '').toUpperCase();
  };

  const renderSeatOrZone = (seatName, abbreviation, zoneLabel, forceZone) => {
    const seat = window.wbData.seats[seatName];
    if (!seat) return '';
    const title = seatTitleFor(seatName);

    if (forceZone || seat.status === 'cargo') {
      return `
        <div class="aircraft-seat-slot">
          <div class="aircraft-seat-title">${title}</div>
          <div class="aircraft-zone" style="font-size:13.5px; padding:3px 6px;">${zoneLabel}</div>
        </div>
      `;
    }
    
    const display = getSeatDisplay(seatName, seat);
    const seatClass = getSeatClass(seat);
    const canUnseat = seat.status === 'installed' && seat.passenger && seat.passenger.name && !seat.locked;
    const unseatHint = canUnseat ? ' title="Double-click to move passenger to needing to be seated" ondblclick="window.moveSeatPassengerToWaiting(\'' + seatName + '\')"' : '';
    return `
      <div class="aircraft-seat-slot">
        <div class="aircraft-seat-title">${title}</div>
        <div class="aircraft-seat ${seatClass}"${unseatHint}>${display || abbreviation}</div>
      </div>
    `;
  };

  let fuselageHTML = '<div class="aircraft-fuselage">';

  // Row 1: Pilot & Copilot
  const pilotSeat = window.wbData.seats['Pilot Seat'];
  const copilotSeat = window.wbData.seats['Copilot Seat'];
  
  fuselageHTML += `<div class="aircraft-row">`;
  fuselageHTML += renderSeatOrZone('Pilot Seat', 'P', 'Zone A', false);
  
  // Copilot OR Zone A if available
  if (availableCargoIds.includes('cargo_a') && copilotSeat.status !== 'installed') {
    fuselageHTML += renderSeatOrZone('Copilot Seat', 'CP', 'Zone A', true);
  } else {
    fuselageHTML += renderSeatOrZone('Copilot Seat', 'CP', 'Zone A', false);
  }
  fuselageHTML += `</div>`;

  // Row 2: LH Mid & RH Mid
  fuselageHTML += `<div class="aircraft-row">`;
  
  const lhMidSeat = window.wbData.seats['LH Mid Seat'];
  const rhMidSeat = window.wbData.seats['RH Mid Seat'];
  
  if (availableCargoIds.includes('cargo_b') && lhMidSeat.status !== 'installed') {
    fuselageHTML += renderSeatOrZone('LH Mid Seat', 'LM', 'Zone B', true);
  } else {
    fuselageHTML += renderSeatOrZone('LH Mid Seat', 'LM', 'Zone B', false);
  }
  
  if (availableCargoIds.includes('cargo_b') && rhMidSeat.status !== 'installed') {
    fuselageHTML += renderSeatOrZone('RH Mid Seat', 'RM', 'Zone B', true);
  } else {
    fuselageHTML += renderSeatOrZone('RH Mid Seat', 'RM', 'Zone B', false);
  }
  fuselageHTML += `</div>`;

  // Row 3: LH Aft & RH Aft
  fuselageHTML += `<div class="aircraft-row">`;
  
  const lhAftSeat = window.wbData.seats['LH Aft Seat'];
  const rhAftSeat = window.wbData.seats['RH Aft Seat'];
  
  if (availableCargoIds.includes('cargo_c') && lhAftSeat.status !== 'installed') {
    fuselageHTML += renderSeatOrZone('LH Aft Seat', 'LA', 'Zone C', true);
  } else {
    fuselageHTML += renderSeatOrZone('LH Aft Seat', 'LA', 'Zone C', false);
  }
  
  if (availableCargoIds.includes('cargo_c') && rhAftSeat.status !== 'installed') {
    fuselageHTML += renderSeatOrZone('RH Aft Seat', 'RA', 'Zone C', true);
  } else {
    fuselageHTML += renderSeatOrZone('RH Aft Seat', 'RA', 'Zone C', false);
  }
  fuselageHTML += `</div>`;

  // Zone D - always visible, show stored seats
  const storedSeats = [];
  const baseSeats = [];
  Object.keys(window.wbData.seats).forEach(seatName => {
    const seat = window.wbData.seats[seatName];
    // Extract abbreviation from seat name
    let abbr = seatName.replace(' Seat', '').replace('Pilot', 'P')
                       .replace('Copilot', 'CP').replace('LH Mid', 'LM')
                       .replace('RH Mid', 'RM').replace('LH Aft', 'LA')
                       .replace('RH Aft', 'RA');
    
    if (seat.status === 'cargo') {
      storedSeats.push(abbr);
    } else if (seat.status === 'base') {
      baseSeats.push(abbr);
    }
  });

  fuselageHTML += `<div class="aircraft-zone-d">`;
  fuselageHTML += `<span style="margin-right:4px;">Zone D:</span>`;
  if (storedSeats.length > 0) {
    storedSeats.forEach(abbr => {
      fuselageHTML += `<div class="stored-seat-label">Seat ${abbr}</div>`;
    });
  } else {
    fuselageHTML += `<span style="font-size:13.5px; color:#999;">Available</span>`;
  }
  fuselageHTML += `</div>`;
  fuselageHTML += `</div>`;

  // BASE box - show seats left at base (OUTSIDE the fuselage)
  let baseBoxHTML = '';
  if (baseSeats.length > 0) {
    baseBoxHTML = `
      <div style="margin-top:10px; padding:6px 10px; border:2px dashed #999; border-radius:8px; background:#e8e8e8; display:flex; flex-wrap:wrap; gap:4px; align-items:center; justify-content:center;">
        <span style="margin-right:4px; color:#666; font-weight:bold; font-size:16.5px;">BASE:</span>
        ${baseSeats.map(abbr => `<div class="stored-seat-label" style="background:#9e9e9e;">Seat ${abbr}</div>`).join('')}
      </div>
    `;
  }

  container.innerHTML = `
    <div style="display:flex; flex-direction:column; align-items:center;">
      <div style="font-weight:bold; margin-bottom:8px; font-size:1.425rem; text-align:center;">Aircraft Layout</div>
      ${fuselageHTML}
      ${baseBoxHTML}
    </div>
  `;
};

/**
 * Render cargo pod side view with dimension markers
 */
window.renderCargoPodDiagram = function() {
  const container = document.getElementById('wb-cargo-pod-diagram');
  if (!container) return;
  
  // Larger compact version showing key reference points
  const svgWidth = 360;
  const svgHeight = 165;
  const marginLeft = 18;
  const axisY = 74;
  
  // Reference points: 10, 30, 50, 67, 84 (inches from datum)
  const points = [
    { pos: 10, label: '10"' },
    { pos: 30, label: '30"' },
    { pos: 50, label: '50"' },
    { pos: 67, label: '67"' },
    { pos: 84, label: '84"' }
  ];
  
  // Scale: pod spans from 10" to 84" (74" total)
  const podStart = 10;
  const podEnd = 84;
  const scale = (svgWidth - marginLeft * 2) / (podEnd - podStart);
  
  // Helper to convert position to x coordinate
  const posToX = (pos) => marginLeft + ((pos - podStart) * scale);
  
  let svg = `
    <svg width="${svgWidth}" height="${svgHeight}" viewBox="0 0 ${svgWidth} ${svgHeight}">
      <!-- Reference axis line (starts at 10") -->
      <line x1="${marginLeft}" y1="${axisY}" x2="${svgWidth - marginLeft}" y2="${axisY}" 
            stroke="#999" stroke-width="2"/>
      
      <!-- Pod A region (10" to 50") -->
      <rect x="${posToX(10)}" y="${axisY - 24}" width="${(50 - 10) * scale}" height="48" 
            fill="#e3f2fd" opacity="0.6" stroke="#2196f3" stroke-width="1.5"/>
      <text x="${posToX(30)}" y="${axisY - 7}" text-anchor="middle" 
        fill="#1565c0" font-size="20" font-weight="bold">Pod A</text>
      
      <!-- Pod B region (50" to 84") -->
      <rect x="${posToX(50)}" y="${axisY - 24}" width="${(84 - 50) * scale}" height="48" 
            fill="#fff3e0" opacity="0.6" stroke="#f57c00" stroke-width="1.5"/>
      <text x="${posToX(67)}" y="${axisY - 7}" text-anchor="middle" 
        fill="#e65100" font-size="20" font-weight="bold">Pod B</text>
  `;
  
  // Add tick marks and labels
  points.forEach(point => {
    const x = posToX(point.pos);
    svg += `
      <line x1="${x}" y1="${axisY - 4}" x2="${x}" y2="${axisY + 4}" 
            stroke="#666" stroke-width="2"/>
      <text x="${x}" y="${axisY + 26}" text-anchor="middle" 
        fill="#666" font-size="17" font-weight="bold">${point.label}</text>
    `;
  });
  
  svg += `</svg>`;
  
  container.innerHTML = svg;
};

/**
 * Save W&B to LOG_Flights
 */
window.saveWBLog = function(options) {
  const opts = (options && typeof options === 'object') ? options : {};
  const silent = !!opts.silent;
  const wb = window.calculateWB();
  const envelope = wbNormalizeEnvelope_(window.wbData.envelopeData);
  if (envelope.length < 3) {
    if (typeof window.flightAppAlert === 'function') {
      window.flightAppAlert('No real CG envelope is available for this aircraft. Flight must not proceed until online sync loads the envelope.', { title: 'Weight & Balance' });
    } else if (window.M) {
      M.toast({ html: 'No real envelope data. Flight must not proceed.', classes: 'red' });
    }
    return;
  }
  
  // Extract fuel weight from items array
  const fuelItem = window.wbData.items.find(item => item.type === 'fuel');
  const fuelWeight = fuelItem ? parseFloat(fuelItem.actualWeight) || 0 : 0;
  
  const wbPayload = {
    flightId: window.wbData.flightId,
    grossWeight: wb.totalWeight,
    cgPosition: wb.cg,
    isSafe: window.checkEnvelope(wb.cg, wb.totalWeight),
    items: window.wbData.items,
    seats: window.wbData.seats,
    fuel: fuelWeight,
    savedAt: new Date().toISOString()
  };

  if (typeof window.runOrQueueServerAction === 'function') {
    window.runOrQueueServerAction({
      method: 'saveWBToLog',
      args: [window.wbData.flightId, wbPayload],
      label: 'W&B save'
    }, {
      onSuccess: function(resp) {
        console.log('W&B saved successfully', resp);
        if (!silent && window.M) M.toast({ html: 'Weight & Balance saved to LOG_Flights', classes: 'green' });
      },
      onQueued: function() {
        if (!silent && window.M) M.toast({ html: 'Offline: W&B save queued for sync', classes: 'orange' });
      },
      onFailure: function(err) {
        console.error('Save failed', err);
        if (!silent && window.M) M.toast({ html: 'Failed to save W&B - check console', classes: 'red' });
      }
    });
    return;
  }

  if (!window.google || !google.script || !google.script.run) {
    console.log('W&B Payload (local):', wbPayload);
    if (!silent && window.M) M.toast({ html: 'Local: W&B payload logged to console', classes: 'orange' });
    return;
  }

  google.script.run.saveWBToLog(window.wbData.flightId, wbPayload);
};

/**
 * Reset W&B
 */
window.resetWB = async function() {
  const ok = await window.flightAppConfirm('Clear all W&B data?', { title: 'Flight App asks you to verify', okText: 'Clear' });
  if (!ok) return;
  window.wbData.items.forEach(item => item.actualWeight = item.plannedWeight);
  Object.keys(window.wbData.seats).forEach(seatName => {
    window.wbData.seats[seatName].enabled = true;
  });
  window.updateWBUI();
  window.renderManifest();
  window.renderSeatToggles();
  if (typeof window.persistWBState_ === 'function') window.persistWBState_();
};

function wbCacheKey_(flightId) {
  return 'mba_cache_wb_' + String(flightId || '').trim();
}

function cacheWBPayload_(flightId, data) {
  try {
    const key = wbCacheKey_(flightId);
    if (!key || !flightId || !data) return;
    localStorage.setItem(key, JSON.stringify({
      flightId: String(flightId || '').trim(),
      cachedAt: new Date().toISOString(),
      payload: data
    }));
  } catch (e) {
    console.warn('W&B cache write failed', e);
  }
}

function readCachedWBPayload_(flightId) {
  try {
    const key = wbCacheKey_(flightId);
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && parsed.payload ? parsed.payload : null;
  } catch (e) {
    console.warn('W&B cache read failed', e);
    return null;
  }
}

function wbPersistCurrentState_(flightId) {
  const fid = String(flightId || (window.wbData && window.wbData.flightId) || '').trim();
  if (!fid || !window.wbData) return;
  cacheWBPayload_(fid, window.wbData);
  wbPersistMissionPax_();
}

window.cacheWBPayload_ = cacheWBPayload_;
window.readCachedWBPayload_ = readCachedWBPayload_;
window.persistWBState_ = wbPersistCurrentState_;

function applyWBPayload_(data) {
  if (!data || typeof data !== 'object') {
    throw new Error('Invalid W&B payload');
  }

  const missionRoute = (window.currentBriefingMission && Array.isArray(window.currentBriefingMission.legs) && window.currentBriefingMission.legs.length)
    ? `${String(window.currentBriefingMission.legs[0].from || '').trim().toUpperCase()} -> ${String(window.currentBriefingMission.legs[0].to || '').trim().toUpperCase()}`
    : '';

  const aircraftReg = String(data.aircraft || '').trim();
  const incomingEnvelope = wbNormalizeEnvelope_(data.envelopeData);
  if (incomingEnvelope.length >= 3) {
    wbCacheRealEnvelopeForAircraft_(aircraftReg, incomingEnvelope);
  }
  const envelopeData = wbFindCachedEnvelopeForAircraft_(aircraftReg);

  window.wbData = {
    ...window.wbData,
    ...data,
    route: data.route || missionRoute,
    seats: data.seats || {},
    envelopeData: envelopeData
  };

  window.renderManifest();
  window.renderSeatToggles();
  window.renderAircraftDiagram();
  window.renderCargoPodDiagram();
  window.updateWBUI();
  window.refreshOfflineLoadTools();
}

/**
 * Initialize W&B from flight data
 */
window.setupWB = async function(flightId) {
  try {
    const cachedPayload = readCachedWBPayload_(flightId);
    const offlinePayload = wbBuildOfflinePayload_(flightId);

    if (!window.google || !google.script || !google.script.run) {
      if (cachedPayload && wbFindCachedEnvelopeForAircraft_(cachedPayload.aircraft).length >= 3) {
        applyWBPayload_(cachedPayload);
        return;
      }
      if (offlinePayload) {
        cacheWBPayload_(flightId, offlinePayload);
        applyWBPayload_(offlinePayload);
        if (window.M) M.toast({ html: 'Offline: initialized W&B from cached mission', classes: 'orange' });
        return;
      }
      if (typeof window.flightAppAlert === 'function') {
        window.flightAppAlert('No real CG envelope is cached for this aircraft. Connect online and open W&B once before attempting this offline flight.', { title: 'Weight & Balance' });
      } else if (window.M) {
        M.toast({ html: 'No real envelope cached for aircraft', classes: 'red' });
      }
      console.log('No Apps Script context and no valid cached/offline W&B payload with real envelope');
      return;
    }

    google.script.run
      .withSuccessHandler(function(data) {
        cacheWBPayload_(flightId, data);
        applyWBPayload_(data);
      })
      .withFailureHandler(function(err) {
        console.error('Failed to load W&B data', err);
        if (cachedPayload) {
          applyWBPayload_(cachedPayload);
          if (window.M) M.toast({ html: 'Offline: loaded cached W&B data', classes: 'orange' });
          return;
        }
        if (offlinePayload) {
          cacheWBPayload_(flightId, offlinePayload);
          applyWBPayload_(offlinePayload);
          if (window.M) M.toast({ html: 'Offline flight: built W&B from local mission', classes: 'orange' });
          return;
        }
        const msg = err && err.message ? err.message : String(err || 'unknown error');
        if (typeof window.flightAppAlert === 'function') {
          window.flightAppAlert('Failed to load W&B data:\n' + msg, { title: 'Weight & Balance' });
        } else if (window.M) {
          M.toast({ html: 'Failed to load W&B data', classes: 'red' });
        }
      })
      .initializeWB(flightId);
  } catch (e) {
    console.error('Exception initializing W&B', e);
  }
};

window.onBriefingFuelUpdated_ = function(flightId, fuelLiters) {
  const fid = String(flightId || '').trim();
  if (!fid) return;

  const liters = Number(fuelLiters || 0);
  const fuelKg = liters > 0 ? (liters * 0.72) : 0;

  const patchFuel = function(payload) {
    if (!payload || typeof payload !== 'object') return payload;
    const next = { ...payload };
    const items = Array.isArray(next.items) ? next.items.slice() : [];
    const fuelItem = items.find(item => String(item && item.type || '').toLowerCase() === 'fuel');
    if (fuelItem) {
      fuelItem.actualWeight = fuelKg;
      fuelItem.plannedWeight = fuelKg;
      fuelItem.name = 'Fuel';
    }
    next.items = items;
    next.fuel = fuelKg;
    return next;
  };

  const cached = readCachedWBPayload_(fid);
  if (cached) {
    cacheWBPayload_(fid, patchFuel(cached));
  }

  if (window.wbData && String(window.wbData.flightId || '').trim() === fid) {
    window.wbData = patchFuel(window.wbData);
    window.renderManifest();
    window.updateWBUI();
  }
};

window.tab3ValidateBeforeProceed = function() {
  const seats = window.wbData && window.wbData.seats ? window.wbData.seats : {};
  const unverified = Object.keys(seats).filter(function(seatId) {
    const seat = seats[seatId];
    if (!seat || seat.status !== 'installed' || !seat.passenger || !seat.passenger.name) return false;
    const isPilot = window.wbData.pilot && seat.passenger.name === window.wbData.pilot;
    if (isPilot) return false;
    return !seat.passenger.verified;
  });

  if (unverified.length) {
    if (window.M) M.toast({ html: 'Verify passenger weights before proceeding', classes: 'orange', displayLength: 2800 });
    return false;
  }
  return true;
};

