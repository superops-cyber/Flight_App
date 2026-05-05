
(function() {
  var state = {
    aircraft: [],
    missionsByReg: {},
    activeFlights: new Set(),
    selectedFlightKey: '',
    pendingOpenKey: '',
    standaloneMode: false,
    messagesByReg: {},
    checksByReg: {},
    livePositions: {},      // reg -> { lat, lng, bearing, gsKts, updatedAtMs }
    mapModeByKey: {},       // key -> 'inhouse' | 'inreach' | 'both'
    inreachUrlByKey: {},    // key -> url string
    leafletMapsByKey: {},   // key -> L.map instance
    leafletMarkersByKey: {}, // key -> L.marker instance
    satMapsByKey: {}        // key -> L.map instance (satellite-only pane)
  };
  var _ffPollTimer = null;
  var _ffLeafletReady = false;
  var _ffMissingGoogleWarned = false;

  function _ffHasGoogleScriptRun_() {
    return !!(window.google && google.script && google.script.run);
  }

  function _ffWarnGoogleMissing_() {
    if (_ffMissingGoogleWarned) return;
    _ffMissingGoogleWarned = true;
    renderEmpty('Flight Following indisponivel neste contexto: google.script.run nao encontrado.');
  }

  function _ffGetUrlParams_() {
    try {
      return new URLSearchParams(window.location.search || '');
    } catch (e) {
      return { get: function() { return ''; } };
    }
  }

  function _ffApplyStandaloneMode_() {
    var params = _ffGetUrlParams_();
    state.standaloneMode = params.get('ffStandalone') === '1';
    state.pendingOpenKey = String(params.get('ffKey') || '').trim();
    var host = document.getElementById('ff-container');
    if (host && state.standaloneMode) host.classList.add('ff-standalone');
  }

  _ffApplyStandaloneMode_();
  loadAircraftAndMissions();
  _ffStartPolling();

  function _ffStartPolling() {
    if (_ffPollTimer) clearInterval(_ffPollTimer);
    _ffPollTimer = setInterval(_ffPollLivePositions, 15000);
  }

  function _ffPollLivePositions() {
    if (!state.activeFlights.size) return;
    if (!_ffHasGoogleScriptRun_()) return;
    google.script.run
      .withSuccessHandler(function(res) {
        if (!res || !res.success) return;
        (res.positions || []).forEach(function(p) {
          state.livePositions[p.reg] = p;
        });
        _ffApplyAllPositions();
      })
      .withFailureHandler(function() {})
      .getLivePositions();
  }

  function _ffApplyAllPositions() {
    state.activeFlights.forEach(function(key) {
      var reg = key.split('|')[0];
      var pos = state.livePositions[reg] || null;
      var sreg = safeId(reg);
      var ageMs = pos ? (Date.now() - pos.updatedAtMs) : Infinity;
      var isLive = pos && ageMs < 90000;  // < 90 s = live
      var isStale = pos && !isLive;

      // Update badge
      var badge = document.getElementById('ff-map-badge-' + sreg);
      if (badge) {
        if (isLive)  { badge.textContent = 'Live (' + Math.round(ageMs/1000) + 's ago)'; badge.className = 'ff-map-source-badge live'; }
        else if (isStale) { badge.textContent = 'Stale (' + Math.round(ageMs/60000) + 'min ago)'; badge.className = 'ff-map-source-badge stale'; }
        else { badge.textContent = 'No position'; badge.className = 'ff-map-source-badge none'; }
      }

      // Auto-switch: if no live data and inReach URL set, prefer inReach
      var mode = state.mapModeByKey[key] || 'inhouse';
      if (!isLive && state.inreachUrlByKey[key] && mode === 'inhouse') {
        // Auto-revert only if not manually set â€” track auto vs manual with a flag
        if (!state._mapModeManual || !state._mapModeManual[key]) {
          mode = 'inreach';
          state.mapModeByKey[key] = 'inreach';
          _ffApplyMapModeUI(key, reg, mode);
        }
      } else if (isLive && mode === 'inreach') {
        // Auto-switch back to inhouse when live data returns
        if (!state._mapModeManual || !state._mapModeManual[key]) {
          mode = 'inhouse';
          state.mapModeByKey[key] = 'inhouse';
          _ffApplyMapModeUI(key, reg, mode);
        }
      }

      // Update Leaflet marker
      if (isLive && mode !== 'inreach') {
        _ffUpdateMapMarker(key, reg, pos);
      }
    });
  }

  function _ffEnsureLeaflet(cb) {
    if (window.L && typeof window.L.map === 'function') { cb(); return; }
    if (!document.getElementById('ff-leaflet-css')) {
      var link = document.createElement('link');
      link.id = 'ff-leaflet-css'; link.rel = 'stylesheet';
      link.href = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css';
      document.head.appendChild(link);
    }
    var script = document.createElement('script');
    script.src = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js';
    script.onload = function() { cb(); };
    document.head.appendChild(script);
  }

  function _ffPlotMissionTrace(map, mission, opts) {
    if (!map || !mission) return false;
    var options = opts || {};
    var color = options.color || '#1565c0';
    var weight = options.weight || 2;
    var opacity = options.opacity || 0.7;
    var dashArray = options.dashArray || '5 5';
    var addEndpoints = !!options.addEndpoints;

    var waypoints = Array.isArray(mission.resolvedWaypoints) ? mission.resolvedWaypoints : [];
    var lls = waypoints.filter(function(w){ return isFinite(w.lat) && isFinite(w.lon); }).map(function(w){ return [w.lat, w.lon]; });
    if (lls.length >= 2) {
      L.polyline(lls, { color: color, weight: weight, opacity: opacity, dashArray: dashArray }).addTo(map);
      if (addEndpoints) {
        L.circleMarker(lls[0], { radius: 4, color: color, fillColor: color, fillOpacity: 0.95, weight: 2 }).addTo(map);
        L.circleMarker(lls[lls.length - 1], { radius: 4, color: color, fillColor: color, fillOpacity: 0.95, weight: 2 }).addTo(map);
      }
      map.fitBounds(L.latLngBounds(lls), { padding: [20, 20] });
      return true;
    }

    var fromCoord = mission.fromCoord && isFinite(mission.fromCoord.lat) && isFinite(mission.fromCoord.lon)
      ? [mission.fromCoord.lat, mission.fromCoord.lon] : null;
    var toCoord = mission.toCoord && isFinite(mission.toCoord.lat) && isFinite(mission.toCoord.lon)
      ? [mission.toCoord.lat, mission.toCoord.lon] : null;
    if (fromCoord && toCoord) {
      L.polyline([fromCoord, toCoord], { color: color, weight: weight, opacity: opacity, dashArray: dashArray }).addTo(map);
      if (addEndpoints) {
        L.circleMarker(fromCoord, { radius: 4, color: color, fillColor: color, fillOpacity: 0.95, weight: 2 }).addTo(map);
        L.circleMarker(toCoord, { radius: 4, color: color, fillColor: color, fillOpacity: 0.95, weight: 2 }).addTo(map);
      }
      map.fitBounds(L.latLngBounds([fromCoord, toCoord]), { padding: [20, 20] });
      return true;
    }
    return false;
  }

  function _ffInitMap(key, reg, mission) {
    if (state.leafletMapsByKey[key]) return; // already initialized
    var sreg = safeId(reg);
    var mapEl = document.getElementById('ff-live-map-' + sreg);
    if (!mapEl) return;
    _ffEnsureLeaflet(function() {
      var map = L.map('ff-live-map-' + sreg, { zoomControl: true, attributionControl: false });
      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 16, minZoom: 2 }).addTo(map);
      state.leafletMapsByKey[key] = map;
      if (_ffPlotMissionTrace(map, mission, { color: '#1565c0', weight: 2, opacity: 0.7, dashArray: '8 6', addEndpoints: true })) return;
      map.setView([-3.5, -60.0], 5); // Default: central Brazil
    });
  }

  function _ffInitSatMap(key, reg, mission) {
    if (state.satMapsByKey[key]) return state.satMapsByKey[key];
    var sreg = safeId(reg);
    var mapEl = document.getElementById('ff-sat-map-' + sreg);
    if (!mapEl) return null;
    var satMap = L.map('ff-sat-map-' + sreg, { zoomControl: true, attributionControl: false });
    state.satMapsByKey[key] = satMap;
    if (!_ffPlotMissionTrace(satMap, mission, { color: '#8ad4ff', weight: 2, opacity: 0.95, dashArray: '7 6', addEndpoints: true })) {
      satMap.setView([-3.5, -60.0], 5);
    }
    return satMap;
  }

  function _ffUpdateMapMarker(key, reg, pos) {
    var sreg = safeId(reg);
    _ffEnsureLeaflet(function() {
      var map = state.leafletMapsByKey[key];
      if (!map) { _ffInitMap(key, reg, null); return; }

      var latlng = [pos.lat, pos.lng];
      // Rotated aircraft icon using divIcon
      var iconHtml = '<div style="width:24px;height:24px;transform:rotate(' + Math.round(pos.bearing) + 'deg);font-size:22px;line-height:1;">âœˆ</div>';
      var icon = L.divIcon({ html: iconHtml, iconSize: [24, 24], iconAnchor: [12, 12], className: '' });

      if (state.leafletMarkersByKey[key]) {
        state.leafletMarkersByKey[key].setLatLng(latlng).setIcon(icon);
        state.leafletMarkersByKey[key].setPopupContent(
          reg + ' â€” ' + Math.round(pos.gsKts) + ' kts / hdg ' + Math.round(pos.bearing) + 'Â°'
        );
      } else {
        var marker = L.marker(latlng, { icon: icon })
          .bindPopup(reg + ' â€” ' + Math.round(pos.gsKts) + ' kts / hdg ' + Math.round(pos.bearing) + 'Â°')
          .addTo(map);
        state.leafletMarkersByKey[key] = marker;
        map.setView(latlng, Math.max(map.getZoom(), 9));
      }
    });
  }

  window.ffSetMapMode = function(key, mode) {
    if (!state._mapModeManual) state._mapModeManual = {};
    state._mapModeManual[key] = true; // user manually set â€” no auto-override
    state.mapModeByKey[key] = mode;
    var reg = key.split('|')[0];
    _ffApplyMapModeUI(key, reg, mode);
  };

  function _ffApplyMapModeUI(key, reg, mode) {
    var sreg = safeId(reg);
    var liveWrap    = document.getElementById('ff-live-wrap-' + sreg);
    var inreachWrap = document.getElementById('ff-inreach-wrap-' + sreg);
    var btnIn  = document.getElementById('ff-maptab-inhouse-' + sreg);
    var btnIr  = document.getElementById('ff-maptab-inreach-' + sreg);
    var btnBoth= document.getElementById('ff-maptab-both-' + sreg);
    var showLive    = mode === 'inhouse' || mode === 'both';
    var showInreach = mode === 'inreach' || mode === 'both';
    if (liveWrap)    liveWrap.style.display    = showLive    ? '' : 'none';
    if (inreachWrap) inreachWrap.style.display = showInreach ? '' : 'none';
    if (btnIn)   btnIn.className   = 'ff-map-toggle-btn' + (mode === 'inhouse' ? ' active' : '');
    if (btnIr)   btnIr.className   = 'ff-map-toggle-btn' + (mode === 'inreach' ? ' active' : '');
    if (btnBoth) btnBoth.className = 'ff-map-toggle-btn' + (mode === 'both'    ? ' active' : '');
    // Invalidate Leaflet size if shown
    if (showLive) {
      var map = state.leafletMapsByKey[key];
      if (map) setTimeout(function(){ map.invalidateSize(); }, 50);
    }
    var satMap = state.satMapsByKey[key];
    if (satMap) setTimeout(function(){ satMap.invalidateSize(); }, 50);
    // Load InReach iframe if url set
    if (showInreach) _ffRenderInreachFrame(key, reg);
  }

  window.ffSetInreachUrl = function(key, url) {
    url = String(url || '').trim();
    if (url && !/^https?:\/\//i.test(url)) url = 'https://' + url;
    state.inreachUrlByKey[key] = url;
    if (!state._mapModeManual) state._mapModeManual = {};
    state._mapModeManual[key] = true;
    var reg = key.split('|')[0];
    var mode = state.mapModeByKey[key] || 'inhouse';
    if (mode === 'inhouse') { // auto-show InReach alongside if user just entered URL
      state.mapModeByKey[key] = 'both';
      _ffApplyMapModeUI(key, reg, 'both');
    } else {
      _ffRenderInreachFrame(key, reg);
    }
  };

  function _ffRenderInreachFrame(key, reg) {
    var sreg = safeId(reg);
    var wrap = document.getElementById('ff-inreach-wrap-' + sreg);
    if (!wrap) return;
    var url = state.inreachUrlByKey[key] || '';
    if (!url) { /* leave fallback text */ return; }
    var existing = wrap.querySelector('iframe');
    if (existing && existing.src === url) return; // already loaded
    wrap.innerHTML = '<iframe class="ff-inreach-iframe" src="' + url.replace(/"/g,'&quot;') + '" allowfullscreen sandbox="allow-scripts allow-forms"></iframe>';
  }

  function loadAircraftAndMissions() {
    if (!_ffHasGoogleScriptRun_()) {
      _ffWarnGoogleMissing_();
      return;
    }
    google.script.run
      .withSuccessHandler(function(data) {
        state.aircraft = Array.isArray(data && data.aircraft) ? data.aircraft : [];
        renderFlightList();
        renderSelectedFlight();

        state.aircraft.forEach(function(acft) {
          google.script.run
            .withSuccessHandler(function(missions) {
              state.missionsByReg[acft.reg] = Array.isArray(missions) ? missions : [];
              renderFlightList();
            })
            .withFailureHandler(function() {
              state.missionsByReg[acft.reg] = [];
              renderFlightList();
            })
            .getFlightFollowMissionsForAcft(acft.reg);
        });
      })
      .withFailureHandler(function(err) {
        renderEmpty('Falha ao carregar voos: ' + esc(String((err && err.message) || err || 'erro desconhecido')));
      })
      .getFlightFollowInit();
  }

  function renderFlightList() {
    var host = document.getElementById('ff-flight-list');
    if (!host) return;

    var rows = [];
    state.aircraft.forEach(function(acft) {
      var missions = state.missionsByReg[acft.reg] || [];
      missions.forEach(function(mission) {
        rows.push({
          key: flightKey(acft.reg, mission.flightLegId || mission.missionId),
          reg: acft.reg,
          missionId: mission.missionId,
          flightLegId: mission.flightLegId || '',
          route: (mission.from || '--') + ' - ' + (mission.to || '--'),
          pilot: mission.pilot || '--',
          date: mission.date || '--',
          takeoffUTC: mission.takeoffUTC || ''
        });
      });
    });

    rows.sort(function(a, b) {
      var ad = String(a.date || '');
      var bd = String(b.date || '');
      if (ad !== bd) return bd.localeCompare(ad);
      var at = String(a.takeoffUTC || '');
      var bt = String(b.takeoffUTC || '');
      if (at !== bt) return bt.localeCompare(at);
      return String(a.missionId || '').localeCompare(String(b.missionId || ''));
    });

    if (!rows.length) {
      host.innerHTML = '<div class="ff-empty-state" style="padding:26px 10px;">Nenhum voo encontrado.</div>';
      return;
    }

    host.innerHTML = '<table class="ff-list-table"><thead><tr><th>Aeronave</th><th>Missao</th><th>Flight ID</th><th>Data</th><th>Rota</th><th>PIC</th><th></th></tr></thead><tbody>' +
      rows.map(function(r) {
        var isActive = r.key === state.selectedFlightKey;
        return '<tr class="ff-list-row' + (isActive ? ' active' : '') + '" onclick="ffOpenFlight(\'' + esc(r.key) + '\')">' +
          '<td>' + esc(r.reg) + '</td>' +
          '<td>' + esc(r.missionId) + '</td>' +
          '<td>' + esc(r.flightLegId || '--') + '</td>' +
          '<td>' + esc(r.date || '--') + '</td>' +
          '<td>' + esc(r.route) + '</td>' +
          '<td>' + esc(r.pilot) + '</td>' +
          '<td><button class="ff-btn-primary" style="padding:4px 8px; font-size:0.68rem;" onclick="event.stopPropagation(); ffOpenFlightInNewTab(\'' + esc(r.key) + '\')">FULL</button></td>' +
        '</tr>';
      }).join('') +
      '</tbody></table>';

    if (state.pendingOpenKey) {
      var found = rows.some(function(r) { return r.key === state.pendingOpenKey; });
      if (found) {
        var keyToOpen = state.pendingOpenKey;
        state.pendingOpenKey = '';
        window.ffOpenFlight(keyToOpen);
      }
    }
  }

  window.ffOpenFlight = function(key) {
    state.selectedFlightKey = String(key || '');
    state.activeFlights = new Set();
    if (state.selectedFlightKey) state.activeFlights.add(state.selectedFlightKey);
    renderFlightList();
    renderSelectedFlight();
    var reg = state.selectedFlightKey.split('|')[0] || '';
    if (reg) loadMessages(reg);
    _ffPollLivePositions();
  };

  window.ffOpenFlightInNewTab = function(key) {
    var k = String(key || '').trim();
    if (!k) return;
    try {
      var u = new URL(window.location.href);
      u.searchParams.set('ffStandalone', '1');
      u.searchParams.set('ffKey', k);
      u.hash = '#view-flightfollow';
      window.open(u.toString(), '_blank', 'noopener');
    } catch (e) {
      var q = '?ffStandalone=1&ffKey=' + encodeURIComponent(k) + '#view-flightfollow';
      window.open(String(window.location.pathname || '') + q, '_blank', 'noopener');
    }
  };

  window.ffLoadFlights = function() {
    loadAircraftAndMissions();
  };

  window.ffToggleAllFlights = function() {
    state.activeFlights = new Set();
    state.selectedFlightKey = '';
    renderFlightList();
    renderSelectedFlight();
  };

  function renderSelectedFlight() {
    var host = document.getElementById('ff-detail-host');
    if (!host) return;
    if (!state.selectedFlightKey) {
      host.innerHTML = '<div class="ff-empty-state">Selecione um voo na lista para abrir o painel de acompanhamento.</div>';
      return;
    }

    var parts = state.selectedFlightKey.split('|');
    var reg = parts[0] || '';
    var flightLegId = parts[1] || '';
    var acft = state.aircraft.find(function(a) { return a.reg === reg; }) || { reg: reg, type: '' };
    var mission = (state.missionsByReg[reg] || []).find(function(m) {
      return String(m && (m.flightLegId || m.missionId) || '') === String(flightLegId || '');
    }) || null;
    if (!mission) {
      host.innerHTML = '<div class="ff-empty-state">Dados do voo ainda carregando. Tente novamente em alguns segundos.</div>';
      return;
    }
    host.innerHTML = buildFocusCard(acft, mission);
    setTimeout(function() {
      _ffInitMap(state.selectedFlightKey, reg, mission);
      renderMessages(reg);
    }, 80);
  }

  function ffFormatDuration_(hours) {
    var h = Number(hours);
    if (!isFinite(h) || h <= 0) return '--:--';
    var totalMin = Math.round(h * 60);
    var hh = Math.floor(totalMin / 60);
    var mm = totalMin % 60;
    return String(hh).padStart(2, '0') + ':' + String(mm).padStart(2, '0');
  }

  function buildFocusCard(acft, mission) {
    var reg = String(acft && acft.reg || '').trim();
    var sreg = safeId(reg);
    var key = flightKey(reg, mission && (mission.flightLegId || mission.missionId) || '');
    var burnLph = Number(acft && (acft.burnLph || acft.burn) || 0);
    var fuelL = Number(mission && mission.fuelL || 0);
    var range = (burnLph > 0 && fuelL > 0) ? ffFormatDuration_(fuelL / burnLph) : '--:--';
    var plannedTime = String(mission && mission.takeoffUTC || '').trim();
    var dateText = String(mission && mission.date || '--').trim();
    var planCode = mission && mission.noPlan ? 'NO PLAN' : String(mission && mission.planId || '--').trim();
    var roleText = String(mission && mission.copilot || '').trim() || 'Sem co-piloto/aluno';
    var pob = 1 + (Array.isArray(mission && mission.pax) ? mission.pax.length : 0);
    return '<div class="ff-focus-card">' +
      '<h3 class="ff-focus-title">' + esc(reg) + ' | Missao ' + esc(String(mission.missionId || '--')) + ' | ' + esc((mission.from || '--') + ' - ' + (mission.to || '--')) + '</h3>' +
      '<div class="ff-focus-grid-top">' +
        ffFieldReadonly_('Piloto', mission.pilot || '--') +
        ffFieldInput_('Flight Follower', 'Nome do acompanhante') +
        ffFieldReadonly_('Data', dateText) +
        ffFieldReadonly_('PLVO Codigo', planCode) +
        ffFieldReadonly_('PLVO Horario', plannedTime ? plannedTime + 'Z' : '--') +
        ffFieldInput_('Horario de Decolagem', '00:00') +
      '</div>' +
      '<div class="ff-focus-grid-mid">' +
        ffFieldReadonly_('Callsign', reg || '--') +
        ffFieldReadonly_('Co-piloto / Aluno', roleText) +
        ffFieldInput_('Follower Backup', 'Nome (opcional)') +
        ffFieldInput_('ETE', '00:00') +
        ffFieldReadonly_('Combustivel a bordo (L)', isFinite(fuelL) && fuelL > 0 ? String(Math.round(fuelL)) : '--') +
        ffFieldReadonly_('Range (combustivel/burn)', range) +
      '</div>' +
      '<div class="ff-focus-grid-mid" style="grid-template-columns: repeat(3, minmax(120px, 1fr)); margin-bottom:0;">' +
        ffFieldReadonly_('POB (piloto + pax)', String(pob)) +
        ffFieldReadonly_('Burn rate (DB_Aircraft)', isFinite(burnLph) && burnLph > 0 ? burnLph.toFixed(1) + ' L/h' : '--') +
        ffFieldReadonly_('Aeronave', acft.type || '--') +
      '</div>' +

      '<div class="ff-map-panel" id="ff-map-panel-' + sreg + '" style="margin-top:10px;">' +
        '<div class="ff-map-toolbar">' +
          '<span class="ff-map-source-badge none" id="ff-map-badge-' + sreg + '">No data</span>' +
          '<button class="ff-map-toggle-btn active" id="ff-maptab-inhouse-' + sreg + '" onclick="ffSetMapMode(\''+key+'\',\'inhouse\')">In-House</button>' +
          '<button class="ff-map-toggle-btn" id="ff-maptab-inreach-' + sreg + '" onclick="ffSetMapMode(\''+key+'\',\'inreach\')">InReach</button>' +
          '<button class="ff-map-toggle-btn" id="ff-maptab-both-' + sreg + '" onclick="ffSetMapMode(\''+key+'\',\'both\')">Both</button>' +
          '<input class="ff-map-inreach-url" id="ff-inreach-url-' + sreg + '" type="url" placeholder="share.garmin.com/DeviceID" onblur="ffSetInreachUrl(\''+key+'\',this.value)" onkeydown="if(event.key===\'Enter\')ffSetInreachUrl(\''+key+'\',this.value)">' +
        '</div>' +
        _ffBuildWxBar(key, sreg) +
        '<div class="ff-map-content-grid">' +
          '<div class="ff-map-primary-col">' +
            '<div class="ff-live-map-wrap" id="ff-live-wrap-' + sreg + '">' +
              '<div class="ff-leaflet-map" id="ff-live-map-' + sreg + '"></div>' +
            '</div>' +
            '<div class="ff-inreach-wrap" id="ff-inreach-wrap-' + sreg + '" style="display:none">' +
              '<div class="ff-inreach-fallback" id="ff-inreach-fallback-' + sreg + '">' +
                'Paste the Garmin MapShare URL above<br>' +
                '<span style="font-size:0.72rem;font-weight:400;color:#607d8b;">e.g. share.garmin.com/YourDevice</span>' +
              '</div>' +
            '</div>' +
          '</div>' +
          '<div class="ff-sat-wrap" id="ff-sat-wrap-' + sreg + '">' +
            '<div class="ff-sat-label" id="ff-sat-label-' + sreg + '" onclick="ffWxCycleLayer(\'' + key + '\')" title="Click to switch Visible / Infrared">Visible</div>' +
            '<div class="ff-sat-map" id="ff-sat-map-' + sreg + '"></div>' +
            '<div class="ff-sat-empty" id="ff-sat-empty-' + sreg + '">Satellite weather frames load here.<br>Use Visible / Infrared and Play when needed.</div>' +
          '</div>' +
        '</div>' +
      '</div>' +

      '<div class="ff-chat-container" style="margin-top:10px;">' +
        '<div class="ff-chat-header">InReach Messages</div>' +
        '<div class="ff-chat-messages" id="ff-chat-' + sreg + '"></div>' +
        '<div class="ff-chat-input-group">' +
          '<input class="ff-chat-input" id="ff-chat-input-' + sreg + '" placeholder="Message to aircraft via inReach email" onkeypress="if(event.key===\'Enter\'){ffSendChat(\'' + esc(reg) + '\');}">' +
          '<button class="ff-chat-send" onclick="ffSendChat(\'' + esc(reg) + '\')">Send</button>' +
        '</div>' +
      '</div>' +
    '</div>';
  }

  function ffFieldReadonly_(label, value) {
    return '<div class="ff-field"><div class="ff-field-label">' + esc(label) + '</div><div class="ff-field-value">' + esc(value) + '</div></div>';
  }

  function ffFieldInput_(label, placeholder) {
    return '<div class="ff-field"><div class="ff-field-label">' + esc(label) + '</div><input class="ff-field-input" type="text" placeholder="' + esc(placeholder || '') + '"></div>';
  }

  function renderGrid() {
    var grid = document.getElementById('ff-cards-grid');
    if (!grid) return;

    if (!state.activeFlights.size) {
      grid.innerHTML = '<div class="ff-empty-state">No flights selected.</div>';
      return;
    }

    var cards = [];
    state.activeFlights.forEach(function(key) {
      var parts = key.split('|');
      var reg = parts[0];
      var missionId = parts[1];
      var acft = state.aircraft.find(function(a) { return a.reg === reg; }) || { reg: reg, type: '' };
      var mission = (state.missionsByReg[reg] || []).find(function(m) { return m.missionId === missionId; }) || null;
      if (mission) cards.push(buildCard(acft, mission));
    });

    grid.innerHTML = cards.length ? cards.join('') : '<div class="ff-empty-state">Waiting for mission data...</div>';
    state.activeFlights.forEach(function(key) {
      loadMessages(key.split('|')[0]);
      // Init Leaflet map for each card after DOM is ready
      var parts = key.split('|');
      var reg = parts[0];
      var missionId = parts[1];
      var mission = (state.missionsByReg[reg] || []).find(function(m) { return m.missionId === missionId; }) || null;
      setTimeout(function() { _ffInitMap(key, reg, mission); }, 80);
    });
    // Kick off a position poll immediately when flights change
    _ffPollLivePositions();
  }

  function buildCard(acft, mission) {
    var reg = acft.reg;
    var sreg = safeId(reg);
    var missionId = mission.missionId || '';
    var key = flightKey(reg, missionId);
    var checks = state.checksByReg[key] || { fuel: false, data: false };
    var pax = Array.isArray(mission.pax) ? mission.pax : [];
    var waypoints = Array.isArray(mission.waypoints) ? mission.waypoints : [];
    var flMatch = String(mission.planId || '').match(/F(\d{3})/i);
    var fl = flMatch ? flMatch[1] : '--';

    var paxRows = pax.length ? pax.map(function(p) {
      var ageSex = [p.sex || '', p.age || ''].filter(Boolean).join(' / ');
      return '<tr>' +
        '<td style="border:1px solid #c0c0c0;padding:6px;">' + esc(p.name || '') + '</td>' +
        '<td style="border:1px solid #c0c0c0;padding:6px;">' + esc(ageSex || '--') + '</td>' +
        '<td style="border:1px solid #c0c0c0;padding:6px;">' + esc(p.phone || '') + '</td>' +
        '<td style="border:1px solid #c0c0c0;padding:6px;">' + esc(p.emergencyContact || '') + '</td>' +
      '</tr>';
    }).join('') : '<tr><td colspan="4" style="border:1px solid #c0c0c0;padding:8px;color:#78909c;">No passengers listed</td></tr>';

    var posRows = waypoints.length ? waypoints.map(function(wp, idx) {
      var isDest = idx === waypoints.length - 1;
      var fuelCell = '';
      if (wp.hasFuel === true) {
        fuelCell = '<span style="color:#2e7d32;font-weight:700;font-size:0.8rem;">â›½</span>';
      } else if (wp.hasFuel === false && wp.fix) {
        fuelCell = '<span style="color:#999;font-size:0.75rem;">â€“</span>';
      }
      return '<tr>' +
        '<td>' + esc(wp.fix || '') + '</td>' +
        '<td style="text-align:center;width:28px;">' + fuelCell + '</td>' +
        '<td><input class="ff-input-cell" placeholder="--:--"></td>' +
        '<td><input class="ff-input-cell" placeholder="--:--"></td>' +
        '<td>' + esc(fl) + '</td>' +
        '<td style="text-align:center;">' + (isDest ? '<span style="color:#2e7d32;font-weight:700;">DEST</span>' : '<button style="background:#1565c0;color:#fff;border:none;border-radius:3px;padding:4px 8px;font-size:0.72rem;font-weight:700;cursor:pointer;">PASS</button>') + '</td>' +
      '</tr>';
    }).join('') : '<tr><td colspan="6" style="padding:8px;color:#78909c;">No route points</td></tr>';

    return '<div class="ff-card">' +
      '<div class="ff-card-header">' +
        '<div><div class="ff-card-reg">' + esc(reg) + '</div><div class="ff-card-type">' + esc(acft.type || '') + '</div></div>' +
        '<div style="font-size:0.8rem;font-weight:700;color:#455a64;">' + esc(missionId) + '</div>' +
      '</div>' +
      '<div class="ff-card-content">' +
        '<table class="ff-table">' +
          '<thead><tr><th>Piloto/INVA</th><th>Acomp</th><th>Data</th><th>PLVO CÃ³digo</th><th>PLVO HorÃ¡rio</th><th>HorÃ¡rio DEC</th></tr></thead>' +
          '<tbody><tr>' +
            '<td><input class="ff-input-cell" value="' + esc(mission.pilot || '') + '"></td>' +
            '<td><input class="ff-input-cell" value="' + esc(mission.copilot || '') + '"></td>' +
            '<td>' + esc(mission.date || '--') + '</td>' +
            '<td>' + esc(mission.noPlan ? 'NO PLAN' : (mission.planId || '--')) + '</td>' +
            '<td>' + esc(mission.takeoffUTC ? mission.takeoffUTC + 'Z' : '--') + '</td>' +
            '<td><input class="ff-input-cell" value="" placeholder="00:00"></td>' +
          '</tr></tbody>' +
        '</table>' +

        '<table class="ff-table">' +
          '<thead><tr><th>Co-Piloto/Aluno</th><th>2Âº Acompanhador</th><th>ETE</th><th>Auto. (LTS)</th><th>Autonomia</th><th>POB</th></tr></thead>' +
          '<tbody><tr>' +
            '<td><input class="ff-input-cell" value="' + esc(mission.copilot || '') + '"></td>' +
            '<td><input class="ff-input-cell" value=""></td>' +
            '<td><input class="ff-input-cell" value=""></td>' +
            '<td>' + esc(mission.fuelL || '--') + '</td>' +
            '<td><input class="ff-input-cell" value="0:00"></td>' +
            '<td>' + esc(mission.pob || '--') + '</td>' +
          '</tr></tbody>' +
        '</table>' +

        '<div class="ff-checklist">' +
          '<div class="ff-check-row"><span>CombustÃ­vel: Coerente com o FS e suficiente para o voo?</span><button class="ff-check-box ' + (checks.fuel ? 'checked' : '') + '" onclick="ffToggleCheck(\'' + esc(key) + '\',\'fuel\')">' + (checks.fuel ? 'OK' : '') + '</button></div>' +
          '<div class="ff-check-row"><span>Outros dados: Coerente com o plano de voo para o dia?</span><button class="ff-check-box ' + (checks.data ? 'checked' : '') + '" onclick="ffToggleCheck(\'' + esc(key) + '\',\'data\')">' + (checks.data ? 'OK' : '') + '</button></div>' +
        '</div>' +

        '<table class="ff-table">' +
          '<thead><tr><th style="width:18%;">Origem/Destino</th><th style="width:28px;">â›½</th><th>PosiÃ§Ã£o</th><th>ETA</th><th>ATA</th><th style="background:#d7f0df;color:#1b5e20;">FL</th><th></th></tr></thead>' +
          '<tbody>' + posRows + '</tbody>' +
        '</table>' +

        '<div style="display:grid;grid-template-columns:1.4fr 1fr;gap:12px;">' +
          '<div>' +
            '<div style="font-size:0.82rem;font-weight:700;margin-bottom:6px;color:#455a64;">Passengers onboard</div>' +
            '<table class="ff-pax-table">' +
              '<thead><tr style="background:#404040;color:#fff;"><th style="border:1px solid #c0c0c0;padding:6px;">Name</th><th style="border:1px solid #c0c0c0;padding:6px;">Sex/Age</th><th style="border:1px solid #c0c0c0;padding:6px;">Phone</th><th style="border:1px solid #c0c0c0;padding:6px;">Emergency Contact</th></tr></thead>' +
              '<tbody>' + paxRows + '</tbody>' +
            '</table>' +
          '</div>' +
          '<div>' +
            '<div style="font-size:0.82rem;font-weight:700;margin-bottom:6px;color:#455a64;">AnotaÃ§Ãµes</div>' +
            '<textarea class="ff-notes-textarea" placeholder="Live notes, map checks, timeline, concerns..."></textarea>' +
          '</div>' +
        '</div>' +

        '<div class="ff-map-panel" id="ff-map-panel-' + sreg + '">' +
          '<div class="ff-map-toolbar">' +
            '<span class="ff-map-source-badge none" id="ff-map-badge-' + sreg + '">No data</span>' +
            '<button class="ff-map-toggle-btn active" id="ff-maptab-inhouse-' + sreg + '" onclick="ffSetMapMode(\''+key+'\',\'inhouse\')">In-House</button>' +
            '<button class="ff-map-toggle-btn" id="ff-maptab-inreach-' + sreg + '" onclick="ffSetMapMode(\''+key+'\',\'inreach\')">InReach</button>' +
            '<button class="ff-map-toggle-btn" id="ff-maptab-both-' + sreg + '" onclick="ffSetMapMode(\''+key+'\',\'both\')">Both</button>' +
            '<input class="ff-map-inreach-url" id="ff-inreach-url-' + sreg + '" type="url" placeholder="share.garmin.com/DeviceID" onblur="ffSetInreachUrl(\''+key+'\',this.value)" onkeydown="if(event.key===\'Enter\')ffSetInreachUrl(\''+key+'\',this.value)">' +
          '</div>' +
          _ffBuildWxBar(key, sreg) +
          '<div class="ff-map-content-grid">' +
            '<div class="ff-map-primary-col">' +
              '<div class="ff-live-map-wrap" id="ff-live-wrap-' + sreg + '">' +
                '<div class="ff-leaflet-map" id="ff-live-map-' + sreg + '"></div>' +
              '</div>' +
              '<div class="ff-inreach-wrap" id="ff-inreach-wrap-' + sreg + '" style="display:none">' +
                '<div class="ff-inreach-fallback" id="ff-inreach-fallback-' + sreg + '">' +
                  'Paste the Garmin MapShare URL above<br>' +
                  '<span style="font-size:0.72rem;font-weight:400;color:#607d8b;">e.g. share.garmin.com/YourDevice</span>' +
                '</div>' +
              '</div>' +
            '</div>' +
            '<div class="ff-sat-wrap" id="ff-sat-wrap-' + sreg + '">' +
              '<div class="ff-sat-label" id="ff-sat-label-' + sreg + '" onclick="ffWxCycleLayer(\'' + key + '\')" title="Click to switch Visible / Infrared">Visible</div>' +
              '<div class="ff-sat-map" id="ff-sat-map-' + sreg + '"></div>' +
              '<div class="ff-sat-empty" id="ff-sat-empty-' + sreg + '">Satellite weather frames load here.<br>Use Visible / Infrared and Play when needed.</div>' +
            '</div>' +
          '</div>' +
        '</div>' +

        '<div class="ff-chat-container">' +
          '<div class="ff-chat-header">InReach Messages</div>' +
          '<div class="ff-chat-messages" id="ff-chat-' + sreg + '"></div>' +
          '<div class="ff-chat-input-group">' +
            '<input class="ff-chat-input" id="ff-chat-input-' + sreg + '" placeholder="Message to aircraft via inReach email" onkeypress="if(event.key===\'Enter\'){ffSendChat(\'' + esc(reg) + '\');}">' +
            '<button class="ff-chat-send" onclick="ffSendChat(\'' + esc(reg) + '\')">Send</button>' +
          '</div>' +
        '</div>' +
      '</div>' +
    '</div>';
  }

  function loadMessages(reg) {
    if (!_ffHasGoogleScriptRun_()) return;
    google.script.run
      .withSuccessHandler(function(messages) {
        state.messagesByReg[reg] = Array.isArray(messages) ? messages : [];
        renderMessages(reg);
      })
      .withFailureHandler(function() {})
      .getFlightFollowMessages(reg);
  }

  function renderMessages(reg) {
    var container = document.getElementById('ff-chat-' + safeId(reg));
    if (!container) return;
    var messages = state.messagesByReg[reg] || [];
    if (!messages.length) {
      container.innerHTML = '<div style="color:#90a4ae;">No inbound messages yet.</div>';
      return;
    }
    container.innerHTML = messages.map(function(msg) {
      var from = String(msg.from || 'unknown');
      var isDevice = from.toLowerCase().indexOf('inreachmail.com') >= 0;
      var when = msg.timestamp ? new Date(msg.timestamp).toLocaleString() : '';
      return '<div class="ff-chat-msg ' + (isDevice ? 'device' : '') + '">' +
        '<div style="font-weight:700;">' + esc(from) + '</div>' +
        '<div style="font-size:0.68rem;color:#607d8b;">' + esc(when) + '</div>' +
        '<div style="margin-top:4px;">' + esc(msg.text || '') + '</div>' +
      '</div>';
    }).join('');
  }

  window.ffToggleCheck = function(key, field) {
    if (!state.checksByReg[key]) state.checksByReg[key] = { fuel: false, data: false };
    state.checksByReg[key][field] = !state.checksByReg[key][field];
    renderGrid();
  };

  window.ffSendChat = function(reg) {
    var input = document.getElementById('ff-chat-input-' + safeId(reg));
    if (!input) return;
    var text = String(input.value || '').trim();
    if (!text) return;

    if (!state.messagesByReg[reg]) state.messagesByReg[reg] = [];
    state.messagesByReg[reg].push({
      timestamp: new Date().toISOString(),
      from: 'dispatch@local',
      text: text
    });
    renderMessages(reg);
    input.value = '';

    if (window.M) {
      M.toast({ html: 'Chat queued for ' + esc(reg) + '. Email send wiring comes next.', classes: 'blue darken-2', displayLength: 2500 });
    }
  };

  function renderEmpty(msg) {
    var host = document.getElementById('ff-detail-host');
    if (host) host.innerHTML = '<div class="ff-empty-state">' + esc(msg) + '</div>';
  }

  function flightKey(reg, missionId) {
    return String(reg || '') + '|' + String(missionId || '');
  }

  function safeId(value) {
    return String(value || '').replace(/[^A-Za-z0-9]/g, '_');
  }

  function esc(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // â”€â”€ NOAA GOES-16 Satellite WX Overlay (NASA GIBS WMTS) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  var _GIBS_LAYERS = {
    geocolor: {
      label: 'GeoColor',
      url: 'https://gibs.earthdata.nasa.gov/wmts/epsg3857/best/GOES-East_ABI_GeoColor/default/{time}/GoogleMapsCompatible_Level7/{z}/{y}/{x}.png',
      maxNativeZoom: 7
    },
    ir: {
      label: 'IR Band 13',
      url: 'https://gibs.earthdata.nasa.gov/wmts/epsg3857/best/GOES-East_ABI_Band13_Clean_Infrared/default/{time}/GoogleMapsCompatible_Level6/{z}/{y}/{x}.png',
      maxNativeZoom: 6
    }
  };

  var _wxByKey = {};  // key â†’ { enabled, frames, frameIdx, playing, timer, layer, pendingLayer, layerKey, opacity, mission }

  function _ffWxBuildUrl(layerKey, time) {
    var cfg = _GIBS_LAYERS[layerKey] || _GIBS_LAYERS.geocolor;
    return cfg.url.replace('{time}', time);
  }

  function _ffWxBuildTimes(numFrames, intervalMin) {
    var times = [];
    var interval = intervalMin * 60000;
    // Keep frames fresh while leaving enough ingest time for GIBS publishing.
    var latest = Math.floor((Date.now() - 30 * 60000) / interval) * interval;
    for (var i = numFrames - 1; i >= 0; i--) {
      var d = new Date(latest - i * interval);
      times.push(d.toISOString().replace(/\.\d{3}Z$/, 'Z'));
    }
    return times;
  }

  function _ffWxUpdateDisplay(key) {
    var ws = _wxByKey[key];
    if (!ws) return;
    var sreg = safeId(key.split('|')[0]);
    var timeEl = document.getElementById('ff-wx-time-' + sreg);
    if (timeEl && ws.frames.length) {
      var d   = new Date(ws.frames[ws.frameIdx]);
      var mon = String(d.getUTCMonth() + 1).padStart(2, '0');
      var day = String(d.getUTCDate()).padStart(2, '0');
      var hh  = String(d.getUTCHours()).padStart(2, '0');
      var mn  = String(d.getUTCMinutes()).padStart(2, '0');
      timeEl.textContent = mon + '/' + day + ' ' + hh + ':' + mn + 'Z';
    }
    var playBtn = document.getElementById('ff-wx-play-' + sreg);
    if (playBtn) playBtn.textContent = ws.playing ? 'â¸' : 'â–¶';
    var satLabel = document.getElementById('ff-sat-label-' + sreg);
    if (satLabel) satLabel.textContent = ws.layerKey === 'ir' ? 'Infrared' : 'Visible';
    var modeSel = document.querySelector('#ff-wx-controls-' + sreg + ' .ff-wx-select');
    if (modeSel) modeSel.value = ws.layerKey;
    var dotsEl = document.getElementById('ff-wx-dots-' + sreg);
    if (dotsEl && ws.frames.length) {
      var html = '';
      for (var i = 0; i < ws.frames.length; i++) {
        html += '<span style="display:inline-block;width:5px;height:5px;border-radius:50%;' +
          'margin:0 1px;background:' + (i === ws.frameIdx ? '#42a5f5' : '#37474f') + '"></span>';
      }
      dotsEl.innerHTML = html;
    }
  }

  function _ffWxRenderFrame(key) {
    var ws = _wxByKey[key];
    if (!ws || !ws.enabled || !ws.frames.length) return;
    var reg = key.split('|')[0];
    var map = state.satMapsByKey[key] || _ffInitSatMap(key, reg, ws.mission || null);
    if (!map) return;
    var cfg = _GIBS_LAYERS[ws.layerKey] || _GIBS_LAYERS.geocolor;
    var url = _ffWxBuildUrl(ws.layerKey, ws.frames[ws.frameIdx]);
    var nextLayer = L.tileLayer(url, {
      tileSize: 256,
      maxNativeZoom: cfg.maxNativeZoom || 6,
      maxZoom: 13,
      opacity: 0,
      attribution: 'NOAA GOES-16 / NASA GIBS',
      updateWhenIdle: true,
      keepBuffer: 2
    }).addTo(map);

    if (ws.pendingLayer) {
      try { map.removeLayer(ws.pendingLayer); } catch(e2) {}
      ws.pendingLayer = null;
    }
    ws.pendingLayer = nextLayer;

    nextLayer.on('load', function() {
      var current = _wxByKey[key];
      if (!current || current.pendingLayer !== nextLayer) return;
      nextLayer.setOpacity(current.opacity == null ? 0.85 : current.opacity);
      if (current.layer) {
        try { map.removeLayer(current.layer); } catch(e3) {}
      }
      current.layer = nextLayer;
      current.pendingLayer = null;
      _ffWxUpdateDisplay(key);
    });

    nextLayer.on('tileerror', function() {
      var current = _wxByKey[key];
      if (!current || current.pendingLayer !== nextLayer) return;
      try { map.removeLayer(nextLayer); } catch(e4) {}
      current.pendingLayer = null;
      if (!current.playing && current.frameIdx > 0) {
        current.frameIdx -= 1;
        _ffWxRenderFrame(key);
        return;
      }
      _ffWxUpdateDisplay(key);
    });

    if (!ws.layer) {
      nextLayer.setOpacity(ws.opacity == null ? 0.85 : ws.opacity);
      ws.layer = nextLayer;
      ws.pendingLayer = null;
      _ffWxUpdateDisplay(key);
    } else {
      _ffWxUpdateDisplay(key);
    }
  }

  function _ffBuildWxBar(key, sreg) {
    var k = esc(key);
    return '<div class="ff-wx-bar">' +
      '<button class="ff-wx-toggle" id="ff-wx-toggle-' + sreg + '" onclick="ffToggleWx(\'' + k + '\')">SAT &#9729;</button>' +
      '<span class="ff-wx-controls" id="ff-wx-controls-' + sreg + '" style="display:none;">' +
        '<button class="ff-wx-btn" id="ff-wx-play-' + sreg + '" onclick="ffWxPlayPause(\'' + k + '\')">&#9654;</button>' +
        '<button class="ff-wx-btn" onclick="ffWxStep(\'' + k + '\',-1)">&#9664;</button>' +
        '<button class="ff-wx-btn" onclick="ffWxStep(\'' + k + '\',1)">&#9654;&#9654;</button>' +
        '<span class="ff-wx-time" id="ff-wx-time-' + sreg + '">--/-- --:--Z</span>' +
        '<span id="ff-wx-dots-' + sreg + '" style="margin:0 3px;"></span>' +
        '<select class="ff-wx-select" onchange="ffWxSetLayer(\'' + k + '\',this.value)">' +
          '<option value="geocolor">Visible</option>' +
          '<option value="ir">Infrared</option>' +
        '</select>' +
        '<input type="range" min="0.25" max="1" step="0.05" value="0.85" class="ff-wx-opacity" title="Opacity" oninput="ffWxSetOpacity(\'' + k + '\',this.value)">' +
      '</span>' +
    '</div>';
  }

  window.ffToggleWx = function(key) {
    if (!_wxByKey[key]) {
      _wxByKey[key] = { enabled: false, frames: [], frameIdx: 0, playing: false, timer: null, layer: null, pendingLayer: null, layerKey: 'geocolor', opacity: 0.85, mission: null };
    }
    var ws = _wxByKey[key];
    ws.enabled = !ws.enabled;
    var sreg     = safeId(key.split('|')[0]);
    var btn      = document.getElementById('ff-wx-toggle-'   + sreg);
    var controls = document.getElementById('ff-wx-controls-' + sreg);
    var satWrap  = document.getElementById('ff-sat-wrap-' + sreg);
    var satEmpty = document.getElementById('ff-sat-empty-' + sreg);
    if (ws.enabled) {
      ws.frames   = _ffWxBuildTimes(8, 10);
      ws.frameIdx = ws.frames.length - 1;
      ws.mission  = (state.missionsByReg[key.split('|')[0]] || []).find(function(m) { return m.missionId === key.split('|')[1]; }) || null;
      if (btn)      { btn.classList.add('active'); btn.innerHTML = 'SAT &#9729; ON'; }
      if (controls) controls.style.display = '';
      if (satWrap) satWrap.classList.add('active');
      if (satEmpty) satEmpty.style.display = 'none';
      _ffEnsureLeaflet(function() {
        _ffInitSatMap(key, key.split('|')[0], ws.mission || null);
        var satMap = state.satMapsByKey[key];
        if (satMap) setTimeout(function(){ satMap.invalidateSize(); }, 60);
        _ffWxRenderFrame(key);
      });
    } else {
      if (ws.timer) { clearInterval(ws.timer); ws.timer = null; }
      ws.playing = false;
      var map = state.satMapsByKey[key];
      if (map && ws.layer) { try { map.removeLayer(ws.layer); } catch(e2) {} }
      if (map && ws.pendingLayer) { try { map.removeLayer(ws.pendingLayer); } catch(e3) {} }
      ws.layer = null;
      ws.pendingLayer = null;
      if (btn)      { btn.classList.remove('active'); btn.innerHTML = 'SAT &#9729;'; }
      if (controls) controls.style.display = 'none';
      if (satWrap) satWrap.classList.remove('active');
      if (satEmpty) satEmpty.style.display = '';
    }
  };

  window.ffWxStep = function(key, dir) {
    var ws = _wxByKey[key];
    if (!ws || !ws.enabled || !ws.frames.length) return;
    if (ws.playing) { window.ffWxPlayPause(key); }
    ws.frameIdx = (ws.frameIdx + dir + ws.frames.length) % ws.frames.length;
    _ffWxRenderFrame(key);
  };

  window.ffWxPlayPause = function(key) {
    var ws = _wxByKey[key];
    if (!ws || !ws.enabled) return;
    ws.playing = !ws.playing;
    if (ws.playing) {
      ws.timer = setInterval(function() {
        ws.frameIdx = (ws.frameIdx + 1) % ws.frames.length;
        _ffWxRenderFrame(key);
      }, 1800);
    } else {
      if (ws.timer) { clearInterval(ws.timer); ws.timer = null; }
    }
    _ffWxUpdateDisplay(key);
  };

  window.ffWxSetLayer = function(key, layerKey) {
    var ws = _wxByKey[key];
    if (!ws) return;
    ws.layerKey = layerKey;
    if (ws.enabled) {
      _ffWxRenderFrame(key);
    } else {
      _ffWxUpdateDisplay(key);
    }
  };

  window.ffWxCycleLayer = function(key) {
    var ws = _wxByKey[key];
    if (!ws) return;
    window.ffWxSetLayer(key, ws.layerKey === 'ir' ? 'geocolor' : 'ir');
  };

  window.ffWxSetOpacity = function(key, val) {
    var ws = _wxByKey[key];
    if (!ws) return;
    ws.opacity = Number(val);
    if (ws.layer) ws.layer.setOpacity(ws.opacity);
    if (ws.pendingLayer) ws.pendingLayer.setOpacity(0);
  };

})();

