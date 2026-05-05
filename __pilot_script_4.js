
  window.runwayWalkthroughData = {
    icao: '',
    rwyIdent: '',
    features: [],
    notes: ''
  };

  window.openRunwayWalkthrough = async function(fromIcao, toIcao, legIndex) {
    if (!fromIcao || !window.appData) return;
    
    const departurePt = String(fromIcao || '').trim().toUpperCase();
    window.runwayWalkthroughData.icao = departurePt;
    
    const runwayRows = (window.appData.airports || []).filter(a => String(a && a.icao || '').trim().toUpperCase() === departurePt);
    const runways = runwayRows.map(function(r) {
      const knownRaw = r && (r.knownFeatures || r.KNOWN_FEATURES || '');
      let features = [];
      let walkedMetadata = null;
      try {
        const parsed = typeof knownRaw === 'string' && knownRaw.trim() ? JSON.parse(knownRaw) : knownRaw;
        if (Array.isArray(parsed)) {
          features = parsed;
        } else if (parsed && typeof parsed === 'object') {
          features = Array.isArray(parsed.features) ? parsed.features : [];
          walkedMetadata = {
            lastWalked: parsed.lastWalked || null,
            published: parsed.published || null,
            staged: parsed.staged || null
          };
        }
      } catch (e) {
        features = [];
      }

      const lengthVal = Number(r && (r.runwayLength || r.length || r.lengthM || r.LENGTH_OFFICIAL || 0));
      return {
        rwyIdent: String(r && (r.runwayIdent || r.rwyIdent || r.runway || r.RWY_IDENT || '')).trim() || 'RWY',
        length: isNaN(lengthVal) ? 0 : lengthVal,
        knownFeatures: Array.isArray(features) ? features : [],
        walkedMetadata: walkedMetadata
      };
    });
    
    if (!runways.length) {
      if (window.M) M.toast({ html: 'No runway data found for ' + departurePt, classes: 'orange' });
      return;
    }
    
    // Default to first runway (or let pilot choose)
    const selectedRwy = runways[0];
    window.runwayWalkthroughData.rwyIdent = selectedRwy.rwyIdent || 'Unknown';
    window.runwayWalkthroughData.features = (selectedRwy.knownFeatures || []).slice();
    
    const title = `${departurePt} RWY ${selectedRwy.rwyIdent} • ${Math.round(selectedRwy.length)}m`;
    document.getElementById('walkthrough-title').textContent = title;
    
    let lastWalkedHtml = '';
    if (selectedRwy.walkedMetadata && selectedRwy.walkedMetadata.lastWalked) {
      const lw = selectedRwy.walkedMetadata.lastWalked;
      const lastDate = new Date(lw.date).toLocaleDateString();
      lastWalkedHtml = `<div style="background:#e8f5e9; border:1px solid #2e7d32; padding:10px; border-radius:6px; margin-bottom:15px;">
        <b>Last Walked:</b> ${lastDate} by ${lw.pilotName}<br>
        <small>${lw.notes || ''}</small>
      </div>`;
    }
    
    const featuresHtml = (window.runwayWalkthroughData.features || []).map((f, idx) => `
      <div style="background:#f5f5f5; padding:8px; margin:5px 0; border-radius:4px; display:flex; justify-content:space-between; align-items:center;">
        <span><b>${f.name}</b> @ ${Math.round(f.distance || 0)}m (${f.side || 'right'})</span>
        <button onclick="window.removeFeature(${idx})" style="background:#d32f2f; color:white; border:none; padding:4px 8px; border-radius:3px; cursor:pointer;">Remove</button>
      </div>
    `).join('');
    
    const contentHtml = `
      ${lastWalkedHtml}
      <div style="margin-bottom:15px;">
        <label style="font-weight:bold; display:block; margin-bottom:5px;">Pilot Notes:</label>
        <textarea id="walkthrough-notes" style="width:100%; height:80px; padding:8px; border:1px solid #ccc; border-radius:4px; font-family:monospace;" placeholder="Observations from runway walk..."></textarea>
      </div>
      <div style="margin-bottom:15px;">
        <h4 style="margin:10px 0 10px 0; font-size:0.95rem;">Known Features:</h4>
        ${featuresHtml || '<p style="color:#999;"><i>No features recorded</i></p>'}
      </div>
    `;
    
    document.getElementById('walkthrough-content').innerHTML = contentHtml;
    document.getElementById('runway-walkthrough-modal').style.display = 'block';
  };

  window.closeRunwayWalkthrough = function() {
    document.getElementById('runway-walkthrough-modal').style.display = 'none';
  };

  window.removeFeature = function(idx) {
    if (window.runwayWalkthroughData.features && idx >= 0 && idx < window.runwayWalkthroughData.features.length) {
      window.runwayWalkthroughData.features.splice(idx, 1);
      // Refresh modal
      window.openRunwayWalkthrough(window.runwayWalkthroughData.icao);
    }
  };

  window.submitRunwayWalkthrough = async function() {
    const btn = document.getElementById('submit-walkthrough-btn');
    btn.disabled = true;
    btn.textContent = 'Submitting...';
    
    try {
      const notes = document.getElementById('walkthrough-notes')?.value || '';
      const payload = {
        icao: window.runwayWalkthroughData.icao,
        rwyIdent: window.runwayWalkthroughData.rwyIdent,
        pilotName: window.currentBriefingMission?.pilot || 'Unknown',
        pilotEmail: window.currentBriefingMission?.meta?.pilotEmail || '',
        features: window.runwayWalkthroughData.features,
        notes: notes
      };
      
      if (typeof window.runOrQueueServerAction === 'function') {
        window.runOrQueueServerAction({
          method: 'submitRunwayWalkthrough_',
          args: [payload],
          label: 'Runway walkthrough'
        }, {
          onSuccess: function(result) {
            btn.disabled = false;
            btn.textContent = '✓ I Walked This Runway';
            if (result && result.success) {
              if (window.M) M.toast({ html: 'Runway walkthrough submitted for review', classes: 'green', displayLength: 3000 });
              window.closeRunwayWalkthrough();
            } else if (window.M) {
              M.toast({ html: (result && result.error) ? result.error : 'Submission failed', classes: 'red' });
            }
          },
          onQueued: function() {
            btn.disabled = false;
            btn.textContent = '✓ I Walked This Runway';
            if (window.M) M.toast({ html: 'Offline: runway walkthrough queued', classes: 'orange' });
            window.closeRunwayWalkthrough();
          },
          onFailure: function(err) {
            btn.disabled = false;
            btn.textContent = '✓ I Walked This Runway';
            if (window.M) M.toast({ html: 'Error: ' + (err && err.message ? err.message : 'Unknown'), classes: 'red' });
          }
        });
        return;
      }

      btn.disabled = false;
      btn.textContent = '✓ I Walked This Runway';
      if (window.M) M.toast({ html: 'Runway submit unavailable', classes: 'red' });
      return;
    } catch (e) {
      btn.disabled = false;
      btn.textContent = '✓ I Walked This Runway';
      if (window.M) M.toast({ html: 'Error: ' + e.message, classes: 'red' });
    }
  };

function calculateBriefFuelTally() {
  const tallyEl = document.getElementById('brief-fuel-tally');
  const warningEl = document.getElementById('brief-fuel-warning');
  const totalBox = document.getElementById('brief-fuel-total-box');
  const launch = parseFloat(tallyEl.dataset.launch) || 0;
  const activeMain = String(document.getElementById('brief_startup_tank')?.value || '').trim().toUpperCase();

  let total = 0;
  const tanks = { LM: 0, RM: 0, LT: 0, RT: 0 };

  document.querySelectorAll('.brief-tank-input').forEach(input => {
    const key = String(input.dataset.tankKey || '').trim().toUpperCase();
    const max = parseFloat(input.dataset.max) || 0;
    let val = parseFloat(input.value) || 0;
    if (val < 0) val = 0;
    if (max > 0 && val > max) {
      val = max;
      input.value = String(max);
      if (window.M) M.toast({ html: `${key} cannot exceed ${Math.round(max)}L`, classes: 'orange', displayLength: 2200 });
    }
    total += val;
    if (Object.prototype.hasOwnProperty.call(tanks, key)) {
      tanks[key] = val;
    }
  });

  tallyEl.innerText = `${Math.round(total)}L`;
  window.briefFuelSnapshot = {
    ...tanks,
    activeMain: activeMain,
    total: total,
    launch: launch,
    updatedAt: new Date().toISOString()
  };

  // Persist whenever values change so they survive tab switches
  briefWriteFuelCache_(window.briefFuelSnapshot);

  // Show mismatch warning with explicit liters above/below planned launch fuel.
  if (launch > 0) {
    const diff = Math.round(total - launch);
    if (diff !== 0) {
      const absDiff = Math.abs(diff);
      const direction = diff > 0 ? 'above' : 'below';
      warningEl.textContent = `⚠ Fuel mismatch: ${absDiff}L ${direction} planned`;
      warningEl.style.display = 'block';
      tallyEl.style.color = '#ffb3b3';
      if (totalBox) totalBox.classList.add('warn');
      return;
    }
  }

  warningEl.style.display = 'none';
  tallyEl.style.color = '#fff';
  if (totalBox) totalBox.classList.remove('warn');
}

window.briefRefreshStartupTankUi_ = function() {
  const activeMain = String(document.getElementById('brief_startup_tank')?.value || '').trim().toUpperCase();
  document.querySelectorAll('#briefing-content .brief-tank-btn[data-tank-key="LM"], #briefing-content .brief-tank-btn[data-tank-key="RM"]').forEach(function(button) {
    button.classList.toggle('startup-selected', String(button.dataset.tankKey || '').trim().toUpperCase() === activeMain);
  });
};

window.briefSetStartupTank_ = function(tank, btnEl) {
  const normalized = String(tank || '').trim().toUpperCase();
  if (normalized !== 'LM' && normalized !== 'RM') return;
  const input = document.getElementById('brief_startup_tank');
  if (input) input.value = normalized;
  window.briefRefreshStartupTankUi_();
  if (btnEl && window.M) M.toast({ html: `Startup tank set to ${normalized}`, classes: 'green', displayLength: 1800 });
  calculateBriefFuelTally();
};

window.briefOpenKeypad_ = function(inputEl, label, opts) {
  if (!inputEl) {
    console.warn('briefOpenKeypad_ called without input element');
    return;
  }

  window._briefKeypadEl = inputEl;
  window._briefKeypadOpts = opts || {};
  var titleEl = document.getElementById('brief-keypad-title');
  var displayEl = document.getElementById('brief-keypad-display');
  var modalEl = document.getElementById('brief-keypad-modal');
  if (!titleEl || !displayEl || !modalEl) {
    console.error('Tab2 keypad modal elements are missing from DOM');
    return;
  }
  titleEl.textContent = label || 'Entry';
  // Opening keypad implies user intends to replace the value.
  displayEl.textContent = '';
  var dotRow = document.getElementById('brief-keypad-dot-row');
  if (dotRow) dotRow.style.display = (opts && opts.decimal) ? '' : 'none';
  modalEl.style.display = 'flex';
};

window.briefKeypadPress_ = function(val) {
  var disp = document.getElementById('brief-keypad-display');
  var opts = window._briefKeypadOpts || {};
  var cur = disp.textContent || '';
  if (val === '⌫') {
    disp.textContent = cur.slice(0, -1);
  } else if (val === 'C') {
    disp.textContent = '';
  } else if (val === '.') {
    if (opts.decimal && !cur.includes('.')) disp.textContent = cur + '.';
  } else {
    var maxLen = opts.maxLen || 10;
    if (cur.length < maxLen) disp.textContent = cur + val;
  }
};

window.briefKeypadOK_ = function() {
  var el = window._briefKeypadEl;
  var opts = window._briefKeypadOpts || {};
  var nextVal = document.getElementById('brief-keypad-display').textContent || '';
  if (el) {
    el.value = nextVal;
    if (el.classList && el.classList.contains('volts-example') && nextVal !== '') {
      el.classList.remove('volts-example');
    }
    if (el.classList && el.classList.contains('brief-tank-input')) {
      calculateBriefFuelTally();
    }
  }
  if (opts && typeof opts.onApply === 'function') {
    try { opts.onApply(nextVal); } catch (e) {}
  }
  var modalEl = document.getElementById('brief-keypad-modal');
  if (modalEl) modalEl.style.display = 'none';
  window._briefKeypadEl = null;
};

window.briefKeypadCancel_ = function() {
  var modalEl = document.getElementById('brief-keypad-modal');
  if (modalEl) modalEl.style.display = 'none';
  window._briefKeypadEl = null;
};

// ── Fuel cache helpers ────────────────────────────────────────────────────────
function briefFuelCacheKey_() {
  const mission = window.currentBriefingMission || {};
  const missionId = String(mission.id || '').trim();
  if (missionId) return 'mba_cache_tab2_fuel_' + missionId;
  const firstLeg = Array.isArray(mission.legs) ? mission.legs[0] : null;
  const fallbackId = String(firstLeg && firstLeg.flightLegId || '').trim();
  return fallbackId ? ('mba_cache_tab2_fuel_' + fallbackId) : '';
}

function briefReadFuelCache_() {
  try {
    const key = briefFuelCacheKey_();
    if (!key) return null;
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) { return null; }
}

function briefWriteFuelCache_(snap) {
  try {
    const key = briefFuelCacheKey_();
    if (!key) return;
    localStorage.setItem(key, JSON.stringify(snap));
  } catch (e) {}
}

window.briefHydrateFuelInputs_ = function() {
  const cached = briefReadFuelCache_();
  if (!cached) return;
  // Restore each tank: update the hidden input AND the visible button
  document.querySelectorAll('#briefing-content .brief-tank-input').forEach(function(input) {
    const key = String(input.dataset.tankKey || '').trim().toUpperCase();
    if (!key || cached[key] == null) return;
    input.value = String(cached[key]);
    var wrap = input.closest('.brief-tank');
    if (wrap) {
      var btn = wrap.querySelector('.brief-tank-btn');
      if (btn) btn.textContent = String(Math.round(Number(cached[key])) || '');
    }
  });
  // Restore startup tank selection
  if (cached.activeMain) {
    const stInput = document.getElementById('brief_startup_tank');
    if (stInput) stInput.value = String(cached.activeMain);
    window.briefRefreshStartupTankUi_();
  }
  // Recompute totals (will also re-persist with current values)
  calculateBriefFuelTally();
};
// ──────────────────────────────────────────────────────────────────────────────

function briefLegInputsCacheKey_() {
  const mission = window.currentBriefingMission || {};
  const missionId = String(mission.id || '').trim();
  if (missionId) return 'mba_cache_tab2_leg_inputs_' + missionId;
  const firstLeg = mission && Array.isArray(mission.legs) ? mission.legs[0] : null;
  const fallbackId = String(firstLeg && firstLeg.flightLegId || '').trim();
  return fallbackId ? ('mba_cache_tab2_leg_inputs_' + fallbackId) : '';
}

function briefZeroOilPromptCacheKey_() {
  const mission = window.currentBriefingMission || {};
  const missionId = String(mission.id || '').trim();
  if (missionId) return 'mba_cache_tab2_zero_oil_prompt_' + missionId;
  const firstLeg = mission && Array.isArray(mission.legs) ? mission.legs[0] : null;
  const fallbackId = String(firstLeg && firstLeg.flightLegId || '').trim();
  return fallbackId ? ('mba_cache_tab2_zero_oil_prompt_' + fallbackId) : '';
}

function briefHasSeenZeroOilPrompt_() {
  try {
    const key = briefZeroOilPromptCacheKey_();
    if (!key) return false;
    const raw = localStorage.getItem(key);
    if (!raw) return false;
    const parsed = JSON.parse(raw);
    return !!(parsed && parsed.promptShown);
  } catch (e) {
    return false;
  }
}

function briefMarkZeroOilPromptSeen_() {
  try {
    const key = briefZeroOilPromptCacheKey_();
    if (!key) return;
    localStorage.setItem(key, JSON.stringify({
      promptShown: true,
      updatedAt: new Date().toISOString()
    }));
  } catch (e) {}
}

function briefReadLegInputsCache_() {
  try {
    const key = briefLegInputsCacheKey_();
    if (!key) return {};
    const raw = localStorage.getItem(key);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return parsed && parsed.entries && typeof parsed.entries === 'object' ? parsed.entries : {};
  } catch (e) {
    return {};
  }
}

function briefWriteLegInputsCache_(entries) {
  try {
    const key = briefLegInputsCacheKey_();
    if (!key) return;
    localStorage.setItem(key, JSON.stringify({
      updatedAt: new Date().toISOString(),
      entries: entries || {}
    }));
  } catch (e) {}
}

function briefCollectLegInputs_() {
  const entries = {};
  document.querySelectorAll('#briefing-content .leg-card').forEach(function(card, idx) {
    const flightLegId = String(card && card.dataset && card.dataset.flightLegId || '').trim();
    const key = flightLegId || String(idx + 1);
    const planInput = card.querySelector('.plan-id-input');
    const zuluInput = card.querySelector('.zulu-input');
    const noPlanInput = card.querySelector('.no-plan-checkbox');
    const noPlan = Boolean(noPlanInput && noPlanInput.checked);
    entries[key] = {
      flightLegId: flightLegId,
      planId: noPlan ? '' : String(planInput && planInput.value || '').trim().toUpperCase(),
      takeoffUTC: noPlan ? '' : String(zuluInput && zuluInput.value || '').trim(),
      noPlan: noPlan
    };
  });
  return entries;
}

window.briefCaptureLegInputs_ = function() {
  const entries = briefCollectLegInputs_();
  briefWriteLegInputsCache_(entries);

  const mission = window.currentBriefingMission;
  if (mission && Array.isArray(mission.legs)) {
    mission.legs.forEach(function(leg) {
      const fid = String(leg && leg.flightLegId || '').trim();
      const row = entries[fid];
      if (!row) return;
      leg.planId = row.planId || '';
      leg.takeoffUTC = row.takeoffUTC || '';
      leg.noFlightPlan = !!row.noPlan;
    });
  }

  return entries;
};

window.briefGetLegInput_ = function(flightLegId) {
  const key = String(flightLegId || '').trim();
  if (!key) return null;
  const live = briefCollectLegInputs_();
  if (live[key]) return live[key];
  const cached = briefReadLegInputsCache_();
  return cached[key] || null;
};

window.briefHydrateLegInputs_ = function() {
  const cached = briefReadLegInputsCache_();
  document.querySelectorAll('#briefing-content .leg-card').forEach(function(card, idx) {
    const flightLegId = String(card && card.dataset && card.dataset.flightLegId || '').trim();
    const key = flightLegId || String(idx + 1);
    const row = cached[key];
    if (!row) return;

    const planInput = card.querySelector('.plan-id-input');
    const zuluInput = card.querySelector('.zulu-input');
    const noPlanInput = card.querySelector('.no-plan-checkbox');

    if (planInput) planInput.value = String(row.planId || '').trim().toUpperCase();
    if (zuluInput) zuluInput.value = String(row.takeoffUTC || '').trim();
    if (noPlanInput) {
      noPlanInput.checked = !!row.noPlan;
      window.briefToggleNoPlan_(noPlanInput, true);
    }
  });
};

window.briefToggleNoPlan_ = function(cb, skipPersist) {
  var card = cb.closest('.leg-card');
  if (!card) return;
  var planInput = card.querySelector('input[placeholder="X42T70QM"]');
  var takeoffInput = card.querySelector('.zulu-input');
  if (planInput) {
    if (cb.checked) {
      planInput.dataset.savedVal = planInput.value;
      planInput.value = '';
      planInput.disabled = true;
    } else {
      planInput.disabled = false;
      planInput.value = planInput.dataset.savedVal || '';
    }
  }
  if (takeoffInput) {
    if (cb.checked) {
      takeoffInput.dataset.savedVal = takeoffInput.value;
      takeoffInput.value = '';
      takeoffInput.disabled = true;
    } else {
      takeoffInput.disabled = false;
      takeoffInput.value = takeoffInput.dataset.savedVal || '';
    }
  }
  if (!skipPersist && typeof window.briefCaptureLegInputs_ === 'function') {
    window.briefCaptureLegInputs_();
  }
};

window.briefSetOil_ = function(val, btnEl) {
  var oilInput = document.getElementById('brief_oil');
  if (oilInput) oilInput.value = String(val);
  var row = btnEl ? btnEl.closest('.oil-choice-row') : null;
  if (!row) return;
  row.querySelectorAll('.oil-choice').forEach(function(b) {
    b.classList.toggle('active', b === btnEl);
  });
};

window.briefOpenTankKeypad_ = function(btnEl, label) {
  if (!btnEl) return;
  var wrap = btnEl.closest('.brief-tank');
  if (!wrap) return;
  var hiddenInput = wrap.querySelector('.brief-tank-input');
  if (!hiddenInput) {
    console.warn('No tank input found for', label);
    return;
  }
  window.briefOpenKeypad_(hiddenInput, label, {
    decimal: false,
    onApply: function(v) {
      hiddenInput.value = v;
      btnEl.textContent = String(v || '');
      calculateBriefFuelTally();
    }
  });
};

window.briefHandleTankTap_ = function(btnEl) {
  if (!btnEl) return;
  var tankKey = String(btnEl.dataset.tankKey || '').trim().toUpperCase();
  var isMain = tankKey === 'LM' || tankKey === 'RM';
  if (!isMain) {
    window.briefOpenTankKeypad_(btnEl, btnEl.dataset.label || 'Tank');
    return;
  }

  var timer = window._briefTankTapTimer;
  if (timer && timer.button === btnEl) {
    clearTimeout(timer.id);
    window._briefTankTapTimer = null;
    window.briefSetStartupTank_(tankKey, btnEl);
    return;
  }

  if (timer && timer.id) {
    clearTimeout(timer.id);
    window._briefTankTapTimer = null;
  }

  window._briefTankTapTimer = {
    button: btnEl,
    id: setTimeout(function() {
      window._briefTankTapTimer = null;
      window.briefOpenTankKeypad_(btnEl, btnEl.dataset.label || 'Tank');
    }, 260)
  };
};

window.bindBriefingInteractions_ = function() {
  if (window._briefDelegatedHandlersBound) return;
  window._briefDelegatedHandlersBound = true;
  document.addEventListener('click', function(e) {
    var oilBtn = e.target.closest('#briefing-content .oil-choice');
    if (oilBtn) {
      e.preventDefault();
      window.briefSetOil_(Number(oilBtn.dataset.oil || 0), oilBtn);
      return;
    }

    var tankBtn = e.target.closest('#briefing-content .js-tank-keypad');
    if (tankBtn) {
      e.preventDefault();
      window.briefHandleTankTap_(tankBtn);
      return;
    }

    var fieldBox = e.target.closest('#briefing-content .js-keypad-field');
    if (fieldBox) {
      e.preventDefault();
      var targetSel = fieldBox.dataset.target || '';
      var input = targetSel ? fieldBox.querySelector(targetSel) : fieldBox.querySelector('input');
      if (!input) {
        console.warn('No input found for keypad field', fieldBox.dataset.label || 'Entry');
        return;
      }
      window.briefOpenKeypad_(input, fieldBox.dataset.label || 'Entry', { decimal: fieldBox.dataset.decimal === '1' });
      return;
    }

    var legTo = e.target.closest('#briefing-content .js-leg-to-keypad');
    if (legTo) {
      e.preventDefault();
      var toInput = legTo.querySelector('input[readonly]');
      if (!toInput) return;
      window.briefOpenKeypad_(toInput, 'T/O UTC', {
        digits: true,
        maxLen: 4,
        onApply: function() {
          if (typeof window.briefCaptureLegInputs_ === 'function') {
            window.briefCaptureLegInputs_();
          }
        }
      });
    }
  }, true);

  document.addEventListener('input', function(e) {
    var target = e.target;
    if (!target || !target.closest) return;
    if (!target.closest('#briefing-content')) return;
    if (target.classList && (target.classList.contains('plan-id-input') || target.classList.contains('zulu-input'))) {
      if (typeof window.briefCaptureLegInputs_ === 'function') {
        window.briefCaptureLegInputs_();
      }
    }
  }, true);
};

function renderLegs(mission) {
  const legs = mission.legs || [];
  const acftConfig = (appData && Array.isArray(appData.aircraft))
    ? appData.aircraft.find(a => String(a?.reg || '').trim().toUpperCase() === String(mission?.acft || '').trim().toUpperCase())
    : null;

  const legViaText_ = function(leg) {
    const tokensFrom = function(raw) {
      if (Array.isArray(raw)) {
        return raw.map(function(wp) {
          if (typeof wp === 'string') return String(wp || '').trim().toUpperCase();
          if (wp && typeof wp === 'object') {
            return String(wp.fix || wp.wp_id || wp.WP_ID || wp.ident || wp.icao || wp.name || wp.label || '').trim().toUpperCase();
          }
          return '';
        }).filter(Boolean);
      }
      const txt = String(raw || '').trim().toUpperCase();
      if (!txt) return [];
      return txt
        .replace(/[→>]/g, ',')
        .split(/[\n\r,;\/|]+/)
        .map(function(part) { return String(part || '').trim().toUpperCase(); })
        .filter(Boolean);
    };

    const routeTokens = tokensFrom(leg && (leg.route || leg.routeStr || leg.via));
    const wpTokens = tokensFrom(leg && leg.waypoints);
    const from = String(leg && (leg.from || leg.origin) || '').trim().toUpperCase();
    const to = String(leg && (leg.to || leg.destination) || '').trim().toUpperCase();

    let picked = wpTokens.length > routeTokens.length ? wpTokens : routeTokens;
    if (!picked.length) picked = [from, to].filter(Boolean);
    if (from && (!picked.length || picked[0] !== from)) picked.unshift(from);
    if (to && (!picked.length || picked[picked.length - 1] !== to)) picked.push(to);
    picked = picked.filter(function(token, idx, arr) { return idx === 0 || token !== arr[idx - 1]; });

    return picked.join(', ');
  };
  // Determine active leg: first non-COMPLETE, or last leg if all complete
  let autoActiveIdx = legs.findIndex(l => (l.logStatus || 'PENDING') !== 'COMPLETE');
  if (autoActiveIdx < 0) autoActiveIdx = legs.length - 1;

  // Honour any pilot-selected leg from this session
  let activeIdx = autoActiveIdx;
  const savedFlightId = window.activeLegFlightId || null;
  if (savedFlightId) {
    const savedIdx = legs.findIndex(l => l.flightLegId === savedFlightId);
    if (savedIdx >= 0) activeIdx = savedIdx;
  }

  // Persist selection globally so downstream tabs can read it
  window.activeLegIndex = activeIdx;
  window.activeLegFlightId = legs[activeIdx] ? legs[activeIdx].flightLegId : (legs[0] ? legs[0].flightLegId : '');

  const legsHtml = legs.map((leg, i) => {
    const status = leg.logStatus || 'PENDING';
    const isComplete = status === 'COMPLETE';
    const isDeparted = status === 'DEPARTED';
    const isActive   = (i === activeIdx) && !isComplete;

    const takeoffFuel = Math.round(leg.takeoffFuel || 0);
    const burnFuel = Math.round(leg.fuel || 0);
    const landingFuel = Math.round(leg.landingFuel || 0);
    const groundTime = parseFloat(leg.groundTime) || 0.5;
    const cacheDraw = leg.isFuelCacheStop ? Math.round(leg.plannedCacheDraw || 0) : 0;
    const routeText = legViaText_(leg);
    const runwayLimitWeight = Math.round(Number(leg.limit || leg.runwayLimit || leg.rwyLimit || 0) || 0);
    const aircraftMaxTakeoffWeight = Math.round(Number(
      leg.maxTakeoffWeight ||
      leg.maxTOWeight ||
      leg.mtow ||
      leg.maxGrossWeight ||
      mission.maxTakeoffWeight ||
      mission.maxTOWeight ||
      mission.mtow ||
      mission.maxGrossWeight ||
      acftConfig?.mtow ||
      acftConfig?.MTOW ||
      0
    ) || 0);
    const effectiveMaxTakeoffWeight = runwayLimitWeight > 0 && aircraftMaxTakeoffWeight > 0
      ? Math.min(runwayLimitWeight, aircraftMaxTakeoffWeight)
      : (runwayLimitWeight || aircraftMaxTakeoffWeight || Math.round(leg.finalPayload || leg.payload || 0));
    const runwayPenalty = (aircraftMaxTakeoffWeight > 0 && effectiveMaxTakeoffWeight > 0 && effectiveMaxTakeoffWeight < aircraftMaxTakeoffWeight)
      ? (aircraftMaxTakeoffWeight - effectiveMaxTakeoffWeight)
      : 0;
    const runwayLimitNote = runwayLimitWeight > 0
      ? `${String(leg.limitType || 'Runway limit').trim() || 'Runway limit'} ${runwayLimitWeight}kg${runwayPenalty > 0 ? ` (${runwayPenalty}kg less than max ${aircraftMaxTakeoffWeight}kg)` : ''}`
      : '';
    const payloadRemaining = Math.round((leg.finalAvailPayload || leg.availPayload || 0) - (leg.payload || 0));

    const actualLog = (mission.actualFuelLogs || [])
      .find(log => log.flightLegId === leg.flightLegId);

    let fuelStatusHtml = '';
    if (actualLog) {
      const isVerified = actualLog.verified === 'YES';
      const statusClass = isVerified ? 'verified' : 'pending';
      const statusText = isVerified ? 'VERIFIED' : 'PENDING REVIEW';
      fuelStatusHtml = `
        <div class="brief-chip-panel cache-usage-alert">
          <div class="cache-usage-title">!!! FUEL CACHE USAGE ALERT !!!</div>
          <div class="cache-usage-row"><span class="qty">${Math.round(actualLog.qty)}L</span> used at ${actualLog.icao}</div>
          <span class="cache-status-pill ${statusClass}">CACHE USE ${statusText}</span>
        </div>`;
    }

    const paxRows = (leg.pax || []).map(p => {
      const isFreight = String(p && p.name || '').toUpperCase() === 'FREIGHT';
      const sex = isFreight ? '-' : (p.gender || '-');
      const cat = isFreight ? '-' : (p.category || '-');
      const bodyKg = Math.round(Number(p && p.weight || 0) || 0);
      const cargoKg = Math.round(Number(p && p.cargo || 0) || 0);
      return `
      <div class="pax-row">
        <span class="pax-col name"><span class="good">${p.name || 'PAX'}</span></span>
        <span class="pax-col sex">${sex}</span>
        <span class="pax-col cat">${cat}</span>
        <span class="pax-col weight">Body: ${bodyKg}kg</span>
        <span class="pax-col cargo">Cargo: ${cargoKg}kg</span>
      </div>`;
    }).join('');

    const legStateText = isComplete ? 'Complete' : (isDeparted ? 'In Flight' : (isActive ? 'Active Leg' : 'Upcoming'));
    const selectBtn = isComplete
      ? `<div class="leg-select-btn complete">✓ Done</div>`
      : `<div class="leg-select-btn" onclick="briefSetActiveLeg_(${i})">${isActive ? 'Active Leg' : 'Select'}</div>`;

    return `
      <div class="leg-card ${isComplete ? 'is-complete' : ''}" data-flight-leg-id="${leg.flightLegId}">
        <div class="leg-head">
          <div>
            <div class="leg-route-line">LEG ${i+1}: ${leg.from} → ${leg.to}</div>
            <div class="leg-via-line">Via ${routeText} | ${legStateText}</div>
          </div>
          ${selectBtn}
        </div>
        <div class="leg-body">
          <div class="leg-summary">
            <div class="brief-chip-row">
              <span class="brief-badge red">Burn: -${burnFuel}L</span>
              <span class="brief-badge green">Max T/O Wt: ${effectiveMaxTakeoffWeight}kg</span>
              <span class="brief-badge gray">Ldg: ${landingFuel}L</span>
              ${cacheDraw > 0 ? `<span class="brief-badge cache-draw-alert">FUEL CACHE DRAW +${cacheDraw}L</span>` : ''}
            </div>
            <div class="brief-data-line">Rem: <span class="good">${payloadRemaining}kg</span> | ${(leg.time||0).toFixed(1)}h flt | ${groundTime.toFixed(1)}h gnd | T/O: ${takeoffFuel}L${runwayLimitNote ? ` | <span class="weight-limit-note">${runwayLimitNote}</span>` : ''}</div>
            ${fuelStatusHtml}
          </div>
          <div class="leg-inputs">
            <div class="fieldbox">
              <label>Plan ID</label>
              <input type="text" class="plan-id-input" placeholder="X42T70QM" maxlength="8">
            </div>
            <div class="fieldbox js-leg-to-keypad">
              <label>T/O UTC</label>
              <input type="text" class="zulu-input" readonly placeholder="1345">
            </div>
            <div class="fieldbox no-plan-box">
              <label>No Flight Plan</label>
              <div class="no-plan-wrap">
                <input type="checkbox" class="no-plan-checkbox" onchange="briefToggleNoPlan_(this)">
                <span>No flight plan filed</span>
              </div>
            </div>
          </div>
          <div class="pax-block">
            ${paxRows || '<div class="pax-load">No PAX</div>'}
          </div>
        </div>
      </div>`;
  }).join('');

  document.getElementById('briefing-legs').innerHTML = legsHtml;
  window.bindBriefingInteractions_();
  if (typeof window.briefHydrateLegInputs_ === 'function') {
    window.briefHydrateLegInputs_();
  }
}

// Called when pilot taps "Select This Leg" on a non-complete leg card
window.briefSetActiveLeg_ = function(idx) {
  const mission = window.currentBriefingMission;
  if (!mission || !mission.legs) return;
  const leg = mission.legs[idx];
  if (!leg) return;
  if ((leg.logStatus || 'PENDING') === 'COMPLETE') {
    if (window.M) M.toast({ html: 'That leg is already complete', classes: 'orange' });
    return;
  }
  window.activeLegIndex   = idx;
  window.activeLegFlightId = leg.flightLegId;
  renderLegs(mission);
  if (window.M) M.toast({ html: `Active leg: ${leg.from} ➔ ${leg.to} (${leg.flightLegId})`, classes: 'green', displayLength: 2200 });
};

window.setupBriefing = function(mission) {
  const container = document.getElementById('briefing-container') || document.getElementById('tab2');

  if (!mission || !mission.legs || mission.legs.length === 0) {
    container.innerHTML = "<p style='padding:20px; color:red;'>No mission data found.</p>";
    return;
  }

  const totalFlight = mission.legs.reduce((s,l)=>s+(l.time||0),0);
  const totalGround = mission.legs.reduce((s,l)=>s+(parseFloat(l.groundTime) || 0.5),0);
  const totalDuty = mission.time || (1.0 + totalFlight + totalGround + 0.75);
  const launchFuel = Math.round(mission.legs[0]?.takeoffFuel || 0);

  const pilot = mission.pilot || '';
  const copilot = mission.meta?.copilot || '';
  const missionDate = mission.date || '';

  const acftConfig = appData.aircraft?.find(a => String(a?.reg || '').trim().toUpperCase() === String(mission.acft || '').trim().toUpperCase());
  const mCap = parseFloat(acftConfig?.MAIN_CAPACITY_L || 0);
  const tCap = parseFloat(acftConfig?.TIP_CAPACITY_L || 0);
  const safeMainCap = Math.max(0, mCap);
  const safeTipCap  = Math.max(0, tCap);

  container.innerHTML = `
  <div id="briefing-content">
    <div class="brief-main">
      <div class="brief-hero">
        <h2 class="brief-title">Mission Brief</h2>
        <div style="font-size:0.62rem; color:#d6e5ff; margin-top:2px;">Tab2 Build: v180</div>
      </div>

      <div class="brief-info">Date: ${missionDate} | Pilot: ${pilot || '—'} | Copilot: ${copilot || '—'}</div>

      <div class="stats-row">
        <div class="stat-card"><span class="stat-val">${totalDuty.toFixed(1)} hr</span><span class="stat-label">Total Duty</span></div>
        <div class="stat-card"><span class="stat-val">${totalFlight.toFixed(1)} hr</span><span class="stat-label">Flight Time</span></div>
        <div class="stat-card"><span class="stat-val">${totalGround.toFixed(1)} hr</span><span class="stat-label">Ground Time</span></div>
        <div class="stat-card"><span class="stat-val green">${launchFuel} L</span><span class="stat-label">Launch Fuel</span></div>
        <div class="stat-card"><span class="stat-val blue">${mission.id || 'N/A'}</span><span class="stat-label">Mission ID</span></div>
      </div>

      <div class="brief-panel">
        <h3 class="brief-section-title">Active Mission Legs</h3>
        <div id="briefing-legs"></div>
      </div>

      <div class="brief-panel">
        <h3 class="brief-section-title green">Pre-Flight Entries</h3>
        <div class="brief-entry-grid">
          <div class="brief-entry-box js-keypad-field" data-label="Tach" data-decimal="1" data-target="#brief_startTach">
            <label>Tach</label>
            <input type="text" readonly id="brief_startTach" value="${acftConfig?.currentTach || '55'}">
          </div>
          <div class="brief-entry-box js-keypad-field" data-label="Volts" data-decimal="1" data-target="#brief_volts">
            <label>Volts</label>
            <input type="text" readonly id="brief_volts" class="volts-example" value="24.2">
          </div>
          <div class="brief-entry-box">
            <label>Oil (L)</label>
            <div class="oil-choice-row">
              <button type="button" class="oil-choice active" data-oil="0">0</button>
              <button type="button" class="oil-choice" data-oil="1">1</button>
              <button type="button" class="oil-choice" data-oil="2">2</button>
            </div>
            <input type="hidden" id="brief_oil" value="0">
          </div>
          <div class="brief-tank group-start">
            <label>Left Tip</label>
            <button type="button" class="brief-tank-btn js-tank-keypad" data-label="Left Tip (L)" data-tank-key="LT"></button>
            <input type="number" style="display:none" class="brief-tank-input" data-tank-key="LT" data-max="${safeTipCap}" value="">
          </div>
          <div class="brief-tank">
            <label>Left Main</label>
            <button type="button" class="brief-tank-btn js-tank-keypad" data-label="Left Main (L)" data-tank-key="LM"></button>
            <input type="number" style="display:none" class="brief-tank-input" data-tank-key="LM" data-max="${safeMainCap}" value="">
          </div>
          <div class="brief-tank">
            <label>Right Main</label>
            <button type="button" class="brief-tank-btn js-tank-keypad" data-label="Right Main (L)" data-tank-key="RM"></button>
            <input type="number" style="display:none" class="brief-tank-input" data-tank-key="RM" data-max="${safeMainCap}" value="">
          </div>
          <div class="brief-tank">
            <label>Right Tip</label>
            <button type="button" class="brief-tank-btn js-tank-keypad" data-label="Right Tip (L)" data-tank-key="RT"></button>
            <input type="number" style="display:none" class="brief-tank-input" data-tank-key="RT" data-max="${safeTipCap}" value="">
          </div>
          <input type="hidden" id="brief_startup_tank" value="">
          <div id="brief-fuel-total-box" class="brief-total-box">
            <label>Total</label>
            <b id="brief-fuel-tally" data-launch="${launchFuel}">0L</b>
          </div>
        </div>
        <div id="brief-fuel-warning" class="brief-warning-inline">⚠ Fuel below planned</div>
        <div class="brief-note">Tap each box to enter values. Double-tap LM or RM to choose startup tank. Use LOAD AIRCRAFT to continue to W&amp;B.</div>
      </div>
    </div>
  </div>
  `;

  renderLegs(mission);
  window.currentBriefingMission = mission;
  window.briefFuelSnapshot = {
    LM: 0,
    RM: 0,
    LT: 0,
    RT: 0,
    activeMain: '',
    total: 0,
    launch: launchFuel,
    updatedAt: new Date().toISOString()
  };
  window.briefRefreshStartupTankUi_();
  calculateBriefFuelTally();
  // Restore any previously entered fuel values (survives tab switches)
  window.briefHydrateFuelInputs_();
  window.bindBriefingInteractions_();
};

window.submitBriefingLog = async function(missionId, opts) {
  try {
    const options = opts || {};
    const tach = document.getElementById('brief_startTach')?.value || '';
    const volts = document.getElementById('brief_volts')?.value || '';
    const oilValue = Number(document.getElementById('brief_oil')?.value || 0);

    if (oilValue === 0 && !briefHasSeenZeroOilPrompt_()) {
      // Only ask once per mission/offline cache lifecycle.
      briefMarkZeroOilPromptSeen_();
      const okNoOil = await window.flightAppConfirm('Oil shows 0 L. Confirm that no oil was added before loading aircraft.', { title: 'Flight App asks you to verify' });
      if (!okNoOil) return false;
    }

    const tanks = {};
    document.querySelectorAll('.brief-tank-input').forEach(input => {
      const key = String(input.dataset.tankKey || '').trim().toUpperCase();
      const max = parseFloat(input.dataset.max) || 0;
      let val = parseFloat(input.value) || 0;
      if (val < 0) val = 0;
      if (max > 0 && val > max) {
        if (window.M) M.toast({ html: `${key} cannot exceed ${Math.round(max)}L`, classes: 'orange' });
        val = max;
        input.value = String(max);
      }
      if (key) tanks[key] = val;
    });
    const activeMain = String(document.getElementById('brief_startup_tank')?.value || '').trim().toUpperCase();
    if (activeMain !== 'LM' && activeMain !== 'RM') {
      if (window.M) M.toast({ html: 'Select the startup tank: LM or RM', classes: 'orange', displayLength: 2800 });
      return false;
    }
    if (Number(tanks[activeMain] || 0) <= 0) {
      if (window.M) M.toast({ html: `${activeMain} must have fuel for startup`, classes: 'orange', displayLength: 2800 });
      return false;
    }

    const legsData = Array.from(document.querySelectorAll('.leg-card')).map((card, idx) => {
      const plan = card.querySelector('input[placeholder="X42T70QM"]')?.value || '';
      const takeoff = card.querySelector('input[placeholder="1345"]')?.value || '';
      const noPlan = !!(card.querySelector('.no-plan-checkbox')?.checked);
      return {
        index: idx + 1,
        flightLegId: String(card.dataset.flightLegId || '').trim(),
        planId: noPlan ? '' : plan,
        takeoffUTC: noPlan ? '' : takeoff,
        noPlan: noPlan
      };
    });

    if (typeof window.briefCaptureLegInputs_ === 'function') {
      window.briefCaptureLegInputs_();
    }

    // Calculate total distance from mission legs
    const totalDist = (window.currentBriefingMission && window.currentBriefingMission.legs)
      ? window.currentBriefingMission.legs.reduce((s, leg) => s + (leg.distance || leg.dist || 0), 0)
      : 0;

    // Use the pilot-selected active leg (set via "Select This Leg" in Tab 2); fall back to leg[0]
    const activeLeg_ = (function() {
      const legs_ = window.currentBriefingMission && window.currentBriefingMission.legs;
      if (!legs_ || !legs_.length) return null;
      if (window.activeLegFlightId) {
        const found = legs_.find(l => l.flightLegId === window.activeLegFlightId);
        if (found) return found;
      }
      return legs_[0];
    })();

    const firstLegFlightId = activeLeg_ ? activeLeg_.flightLegId : (missionId || '');

    if (activeLeg_) {
      activeLeg_.startTach = tach;
    }
    if (window.currentBriefingMission) {
      window.currentBriefingMission.startTach = tach;
      if (Array.isArray(window.currentBriefingMission.legs)) {
        const match = window.currentBriefingMission.legs.find(function(l) { return String(l && l.flightLegId || '') === String(firstLegFlightId || ''); });
        if (match) match.startTach = tach;
      }
    }
    try {
      const mid = String((window.currentBriefingMission && window.currentBriefingMission.id) || '').trim();
      if (mid) {
        const key = 'mission_' + mid;
        const cached = JSON.parse(localStorage.getItem(key) || 'null');
        if (cached && Array.isArray(cached.legs)) {
          const cachedLeg = cached.legs.find(function(l) { return String(l && l.flightLegId || '') === String(firstLegFlightId || ''); });
          if (cachedLeg) cachedLeg.startTach = tach;
          cached.startTach = tach;
          localStorage.setItem(key, JSON.stringify(cached));
        }
      }
    } catch (e) {}

    const fuelTotal = Object.values(tanks).reduce((s, v) => s + (parseFloat(v) || 0), 0) || 0;

    const payload = {
      flightLegId: firstLegFlightId,  // active leg's flight ID (e.g., ADS26-001-02)
      date: (window.currentBriefingMission && window.currentBriefingMission.date) || new Date().toISOString().split('T')[0],
      pilot: (window.currentBriefingMission && window.currentBriefingMission.pilot) || '',
      acft: (window.currentBriefingMission && window.currentBriefingMission.acft) || '',
      from: activeLeg_ ? (activeLeg_.from || '') : '',
      to: activeLeg_ ? (activeLeg_.to || '') : '',
      totalDist: activeLeg_ ? (activeLeg_.distance || activeLeg_.dist || 0) : totalDist,
      startTach: tach,
      fuelTotal: fuelTotal,
      activeMain: activeMain,
      oil: oilValue,
      volts: volts,
      actualLoadJSON: JSON.stringify({ tanks: tanks, activeMain: activeMain, fuelTotal: fuelTotal, oil: oilValue, legs: legsData, savedAt: new Date().toISOString() }),
      savedAt: new Date().toISOString(),
      userAgent: navigator.userAgent
    };

    if (typeof window.runOrQueueServerAction === 'function') {
      window.runOrQueueServerAction({
        method: 'saveMissionToLog',
        args: [payload],
        label: 'Briefing save'
      }, {
        onSuccess: function(resp) {
          console.log('Save success', resp);
          if (typeof window.onBriefingFuelUpdated_ === 'function') {
            window.onBriefingFuelUpdated_(firstLegFlightId, fuelTotal, payload.actualLoadJSON);
          }
          if (window.M) M.toast({ html: 'Aircraft fuel loaded', classes: 'green' });
          if (options.advanceToTab3 && typeof window.switchTab === 'function') {
            window.__skipTab2GateOnce = true;
            window.switchTab(3);
          }
        },
        onQueued: function() {
          if (typeof window.onBriefingFuelUpdated_ === 'function') {
            window.onBriefingFuelUpdated_(firstLegFlightId, fuelTotal, payload.actualLoadJSON);
          }
          if (window.M) M.toast({ html: 'Offline: aircraft load queued', classes: 'orange' });
          if (options.advanceToTab3 && typeof window.switchTab === 'function') {
            window.__skipTab2GateOnce = true;
            window.switchTab(3);
          }
        },
        onFailure: function(err) {
          console.error('Save failed', err);
          if (window.M) M.toast({ html: 'Load failed — check connection', classes: 'red' });
        }
      });
      return true;
    } else if (window.google && google.script && google.script.run) {
      google.script.run.saveMissionToLog(payload);
      return true;
    } else {
      console.log('submitBriefingLog payload (no Apps Script):', payload);
      if (window.M) M.toast({ html: 'Local mode: payload logged', classes: 'orange' });
      return true;
    }
  } catch (err) {
    console.error('submitBriefingLog error', err);
    if (window.M) M.toast({ html: 'Unexpected error loading aircraft', classes: 'red' });
    return false;
  }
};

window.validateFuelAtLeastOneTank = function() {
  let hasAnyFuel = false;
  document.querySelectorAll('.brief-tank-input').forEach(input => {
    const val = parseFloat(input.value) || 0;
    if (val > 0) hasAnyFuel = true;
  });
  return hasAnyFuel;
};

window.tab2LoadAircraftThenAdvance = function() {
  const missionId = (window.currentBriefingMission && window.currentBriefingMission.id) || '';
  if (!missionId) {
    if (window.M) M.toast({ html: 'No mission loaded in briefing', classes: 'orange' });
    return;
  }
  // Safeguard: require at least some fuel in tanks before advancing
  if (!window.validateFuelAtLeastOneTank()) {
    if (window.M) M.toast({ html: 'Please enter fuel in at least one tank before loading aircraft', classes: 'orange', displayLength: 3000 });
    return;
  }
  window.submitBriefingLog(missionId, { advanceToTab3: true });
};


