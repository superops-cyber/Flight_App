
  function openRunwayDatabase() {
    openRunwaySurveyTool();
  }

  function openTab1Inclinometer() {
    rwyInclinometerOpenStandaloneTab_('tab1-inclinometer-angle', 'Salvar ângulo');
  }

  function openAircraftDocsModal() {
    const overlay = document.getElementById('modal-acft-docs');
    if (!overlay) return;

    overlay.style.display = 'flex';
    document.body.style.overflow = 'hidden';

    // Resolve active mission's aircraft tail (may be empty if no mission selected)
    const missions = cacheGet(OFFLINE_CACHE_KEYS.SCHEDULED_MISSIONS) || [];
    const mission = missions.find(function(m) { return String(m && m.id || '') === String(activeMission || ''); });
    const activeTail = String((mission && mission.acft) || '').trim().toUpperCase();

    const cached = cacheGet(OFFLINE_CACHE_KEYS.AIRCRAFT_DOCS);
    _renderAircraftDocsModal_(cached, activeTail);

    // Background refresh when online — re-render silently when done
    if (window.google && google.script && google.script.run) {
      google.script.run
        .withSuccessHandler(function(result) {
          if (result && result.success) {
            cacheSet(OFFLINE_CACHE_KEYS.AIRCRAFT_DOCS, result);
            // Only re-render if modal is still open
            if (document.getElementById('modal-acft-docs').style.display !== 'none') {
              _renderAircraftDocsModal_(result, activeTail);
            }
          }
        })
        .withFailureHandler(function() { /* silent — already showing cached */ })
        .getAircraftDocsForTools('');
    }
  }

  function _renderAircraftDocsModal_(result, activeTail) {
    const title = document.getElementById('acft-docs-title');
    const sub = document.getElementById('acft-docs-subtitle');
    const body = document.getElementById('acft-docs-body');
    const folderRow = document.getElementById('acft-docs-folder-row');
    const folderLink = document.getElementById('acft-docs-folder-link');

    if (folderRow) folderRow.style.display = 'none';

    if (!result || !result.success || !result.docs || result.docs.length === 0) {
      if (title) title.textContent = 'Aircraft Manuals & Documents';
      if (sub) sub.textContent = '';
      if (body) body.innerHTML = '<p class="orange-text" style="margin-top:18px; font-weight:700; text-align:center;">No documents cached yet.<br><span style="font-size:0.82rem; color:#607d8b; font-weight:500;">Press REFRESH on Tab 1 to sync documents.</span></p>';
      return;
    }

    const allDocs = result.docs;

    // Collect all unique tails, active tail first
    const tailSet = {};
    allDocs.forEach(function(d) { tailSet[d.tail] = true; });
    const allTails = Object.keys(tailSet).sort();
    if (activeTail && tailSet[activeTail]) {
      allTails.splice(allTails.indexOf(activeTail), 1);
      allTails.unshift(activeTail);
    }
    const selectedTail = (activeTail && tailSet[activeTail]) ? activeTail : (allTails[0] || '');

    if (title) title.textContent = 'Manuals & Documents';

    // Build tail filter pills
    let filterHtml = '<div id="acft-docs-filter" style="display:flex; flex-wrap:wrap; gap:7px; margin-bottom:14px;">';
    allTails.forEach(function(t) {
      const isActive = t === selectedTail;
      filterHtml += '<button type="button" onclick="_acftDocsFilterTail_(\'' + _escAttr_(t) + '\')" '
        + 'id="acft-docs-pill-' + _escAttr_(t) + '" '
        + 'style="height:30px; padding:0 13px; border-radius:20px; border:2px solid ' + (isActive ? '#0b5394' : '#cfd8dc') + '; '
        + 'background:' + (isActive ? '#0b5394' : '#fff') + '; color:' + (isActive ? '#fff' : '#455a64') + '; '
        + 'font-size:0.78rem; font-weight:900; cursor:pointer;">' + _escHtml_(t) + '</button>';
    });
    filterHtml += '</div>';

    if (body) body.innerHTML = filterHtml + '<div id="acft-docs-list"></div>';

    if (sub) {
      const verified = result.lastVerifiedOffline ? 'Cached · last verified ' + result.lastVerifiedOffline : 'Cached offline';
      sub.textContent = allTails.length + ' aircraft  ·  ' + allDocs.length + ' total docs  ·  ' + verified;
    }

    _acftDocsRenderList_(allDocs, selectedTail, folderLink, folderRow);
  }

  // Store docs in closure so filter pills can access them
  let _acftDocsAllDocs_ = [];

  function _acftDocsFilterTail_(tail) {
    // Update pill styles
    document.querySelectorAll('#acft-docs-filter button').forEach(function(btn) {
      const isActive = btn.id === 'acft-docs-pill-' + tail;
      btn.style.background = isActive ? '#0b5394' : '#fff';
      btn.style.color = isActive ? '#fff' : '#455a64';
      btn.style.borderColor = isActive ? '#0b5394' : '#cfd8dc';
    });
    const folderRow = document.getElementById('acft-docs-folder-row');
    const folderLink = document.getElementById('acft-docs-folder-link');
    _acftDocsRenderList_(_acftDocsAllDocs_, tail, folderLink, folderRow);
  }

  function _acftDocsRenderList_(allDocs, tail, folderLink, folderRow) {
    _acftDocsAllDocs_ = allDocs;
    const list = document.getElementById('acft-docs-list');
    if (!list) return;

    const docs = allDocs.filter(function(d) { return d.tail === tail; });

    // Update folder link
    const cached = cacheGet(OFFLINE_CACHE_KEYS.AIRCRAFT_DOCS);
    const folderUrl = cached && cached.folderUrl ? cached.folderUrl : '';
    if (folderRow && folderLink) {
      if (folderUrl) {
        folderLink.href = folderUrl;
        folderRow.style.display = '';
      } else {
        folderRow.style.display = 'none';
      }
    }

    if (docs.length === 0) {
      list.innerHTML = '<p class="grey-text" style="margin-top:12px; text-align:center;">No documents on file for ' + _escHtml_(tail) + '.</p>';
      return;
    }

    // Group by docType
    const groups = {};
    docs.forEach(function(d) {
      const grp = d.docType || 'Other';
      if (!groups[grp]) groups[grp] = [];
      groups[grp].push(d);
    });

    let html = '';
    Object.keys(groups).sort().forEach(function(grp) {
      html += '<div style="margin-top:14px;">';
      html += '<div style="font-size:0.72rem; font-weight:900; text-transform:uppercase; color:#7d8b97; letter-spacing:0.06em; margin-bottom:6px;">' + _escHtml_(grp) + '</div>';
      groups[grp].forEach(function(d) {
        const hasLink = !!d.driveUrl;
        const isExpired = d.expiryDate && (new Date(d.expiryDate) < new Date());
        const iconColor = isExpired ? '#c62828' : (d.critical ? '#e65100' : '#1565c0');
        html += '<div style="display:flex; align-items:center; gap:10px; padding:9px 0; border-bottom:1px solid #f0f3f5;">';
        html += '<i class="material-icons" style="font-size:1.3rem; color:' + iconColor + ';">' + (hasLink ? 'description' : 'insert_drive_file') + '</i>';
        html += '<div style="flex:1; min-width:0;">';
        if (hasLink) {
          html += '<a href="' + _escAttr_(d.driveUrl) + '" target="_blank" rel="noopener noreferrer" style="font-weight:800; color:#0b5394; text-decoration:none; font-size:0.9rem; display:block; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">' + _escHtml_(d.docName || d.docType) + '</a>';
        } else {
          html += '<span style="font-weight:800; color:#37474f; font-size:0.9rem;">' + _escHtml_(d.docName || d.docType) + '</span>';
        }
        const meta = [];
        if (d.revision) meta.push('Rev ' + _escHtml_(d.revision));
        if (d.effectiveDate) meta.push('Eff ' + _escHtml_(d.effectiveDate));
        if (d.expiryDate) meta.push((isExpired ? '⚠ Expired ' : 'Exp ') + _escHtml_(d.expiryDate));
        if (meta.length) html += '<div style="font-size:0.72rem; color:#78909c; margin-top:2px;">' + meta.join('  ·  ') + '</div>';
        html += '</div></div>';
      });
      html += '</div>';
    });
    list.innerHTML = html;
  }

  function closeAircraftDocsModal() {
    const overlay = document.getElementById('modal-acft-docs');
    if (overlay) overlay.style.display = 'none';
    document.body.style.overflow = '';
  }

  function _escHtml_(s) {
    return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  function _escAttr_(s) {
    return String(s || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
  }

  const INCLINOMETER_HTML_CACHE_KEY = 'mba_cache_inclinometer_standalone_v1';

  function rwyUpdateInclinometerCacheIndicator_() {
    var el = document.getElementById('incl-cache-status');
    if (!el) return;
    var ready = !!rwyGetCachedStandaloneHtml_();
    el.textContent = ready
      ? 'Inclinometer offline cache: READY'
      : 'Inclinometer offline cache: NOT READY (open once online)';
    el.style.color = ready ? '#2e7d32' : '#ef6c00';
  }

  function rwyGetCachedStandaloneHtml_() {
    try {
      const raw = localStorage.getItem(INCLINOMETER_HTML_CACHE_KEY);
      if (!raw) return '';
      const parsed = JSON.parse(raw);
      const html = parsed && typeof parsed.html === 'string' ? parsed.html : '';
      return html && html.trim() ? html : '';
    } catch (e) {
      return '';
    }
  }

  function rwySetCachedStandaloneHtml_(html) {
    try {
      const text = String(html || '');
      if (!text.trim()) return;
      localStorage.setItem(INCLINOMETER_HTML_CACHE_KEY, JSON.stringify({
        savedAt: Date.now(),
        html: text
      }));
      rwyUpdateInclinometerCacheIndicator_();
    } catch (e) {}
  }

  function rwyLoadStandaloneHtmlIntoPopup_(popup, html, targetInputId, saveLabel) {
    popup.document.open();
    popup.document.write(String(html || ''));
    popup.document.close();
    try {
      popup._incTargetInputId = String(targetInputId || 'angleInput');
      popup._incSaveLabel = String(saveLabel || 'Save angle');
      popup._incReturnUrl = String(window.location.href || WEB_APP_URL || '');
      popup._incStorageKey = 'missionBriefing.inclinometerResult';
    } catch(e2) {}
    try { popup.focus(); } catch (e) {}
  }

  window.setInclinometerValue = function(value, targetInputId) {
    const inputId = String(targetInputId || 'tab1-inclinometer-angle');
    const input = document.getElementById(inputId);
    if (!input) return false;
    input.value = String(value || '');
    try { input.dispatchEvent(new Event('input', { bubbles: true })); } catch (e) {}
    try { input.dispatchEvent(new Event('change', { bubbles: true })); } catch (e) {}
    return true;
  };

  function rwyConsumeStandaloneInclinometerResult_() {
    let payload = null;
    try {
      payload = JSON.parse(localStorage.getItem('missionBriefing.inclinometerResult') || 'null');
    } catch (e) {
      payload = null;
    }
    if (!payload || payload.value == null) return false;
    const applied = window.setInclinometerValue(payload.value, payload.targetInputId);
    if (applied) {
      try { localStorage.removeItem('missionBriefing.inclinometerResult'); } catch (e) {}
    }
    return applied;
  }

  window.addEventListener('focus', rwyConsumeStandaloneInclinometerResult_);
  document.addEventListener('visibilitychange', function() {
    if (!document.hidden) rwyConsumeStandaloneInclinometerResult_();
  });
  document.addEventListener('DOMContentLoaded', rwyUpdateInclinometerCacheIndicator_);
  window.addEventListener('online', rwyUpdateInclinometerCacheIndicator_);

  function rwyInclinometerOpenStandaloneTab_(targetInputId, saveLabel) {
    rwyConsumeStandaloneInclinometerResult_();
    let popup = null;
    try {
      popup = window.open('', '_blank');
    } catch (e) {
      popup = null;
    }

    if (!popup) {
      if (window.M) M.toast({ html: 'Popup bloqueado. Permita popups para abrir modo câmera.', classes: 'orange' });
      return false;
    }

    try {
      popup.document.open();
      popup.document.write('<!doctype html><html><head><meta name="viewport" content="width=device-width, initial-scale=1"><title>Loading...</title></head><body style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;padding:16px;background:#0b1118;color:#d9e6ef;">Abrindo inclinômetro...</body></html>');
      popup.document.close();
    } catch (e) {}

    const cachedHtml = rwyGetCachedStandaloneHtml_();

    if (!google || !google.script || !google.script.run) {
      if (cachedHtml) {
        try {
          rwyLoadStandaloneHtmlIntoPopup_(popup, cachedHtml, targetInputId, saveLabel);
          if (window.M) M.toast({ html: 'Modo offline: inclinômetro carregado do cache.', classes: 'blue darken-2' });
          return true;
        } catch (e) {
          try { popup.close(); } catch (e2) {}
          if (window.M) M.toast({ html: 'Falha ao abrir inclinômetro offline.', classes: 'orange' });
          return false;
        }
      }
      try { popup.close(); } catch (e) {}
      if (window.M) M.toast({ html: 'Inclinômetro offline indisponível. Abra uma vez online para salvar cache.', classes: 'orange' });
      return false;
    }

    google.script.run
      .withSuccessHandler(function(html) {
        try {
          rwySetCachedStandaloneHtml_(html);
          rwyLoadStandaloneHtmlIntoPopup_(popup, html, targetInputId, saveLabel);
        } catch (e) {
          if (window.M) M.toast({ html: 'Falha ao carregar HTML standalone no popup.', classes: 'orange' });
        }
      })
      .withFailureHandler(function(err) {
        if (cachedHtml) {
          try {
            rwyLoadStandaloneHtmlIntoPopup_(popup, cachedHtml, targetInputId, saveLabel);
            if (window.M) M.toast({ html: 'Backend offline. Inclinômetro aberto do cache.', classes: 'blue darken-2' });
            return;
          } catch (e2) {}
        }
        try { popup.close(); } catch (e) {}
        if (window.M) M.toast({ html: 'Falha no backend ao abrir standalone: ' + (err && err.message ? err.message : 'erro'), classes: 'orange' });
      })
      .getInclinometerStandaloneHtml();

    return true;
  }

  function rwyInclinometerOpenTopLevel_() {
    try {
      // Never navigate to userCodeAppPanel/googleusercontent iframe URLs directly;
      // they can render blank when opened top-level on iOS.
      const hardcodedExec = 'https://script.google.com/macros/s/AKfycbwOSmd0vN35IZe1LyRzyH3EN18gHSiMpSsOLf0M0XhVwFTLI_9Hh5etXAeev_iUl0fwig/exec';
      const base = String(WEB_APP_URL || hardcodedExec || '');
      if (!base) throw new Error('Missing web app URL');
      const sep = base.indexOf('?') >= 0 ? '&' : '?';
      const url = base + sep + 'view=pilot&camtop=1&t=' + Date.now();

      // Prefer same-tab navigation first so popup blockers cannot interfere.
      try {
        if (window.top && window.top !== window) {
          window.top.location.href = url;
        } else {
          window.location.href = url;
        }
        return true;
      } catch (e) {
        // Fall back to new-tab path below.
      }

      // iOS Safari can create about:blank first; explicitly assign the target URL.
      const opened = window.open('', '_blank');
      if (opened) {
        opened.location.href = url;
        return true;
      }

      // Fallback path if popup object is blocked but click navigation is allowed.
      const a = document.createElement('a');
      a.href = url;
      a.target = '_blank';
      a.rel = 'noopener';
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      return true;
    } catch (err) {
      if (window.M) M.toast({ html: 'Falha ao abrir aba nova para modo câmera.', classes: 'orange' });
      return false;
    }
  }
  const runwayInclinometerState = {
    stream: null,
    orientationBound: false,
    zeroOffset: 0,
    filteredPitch: 0,
    rawBeta: null,
    target: null,
    currentDisplayAngle: 0,
    pixelsPerDegree: 6,
    saveInputId: 'rwysurvey-obs-a-popup',
    saveLabel: 'SALVAR ÂNGULO'
  };

  function rwyInclinometerClamp_(v, lo, hi) {
    return Math.max(lo, Math.min(hi, Number(v || 0)));
  }

  function rwyInclinometerBuildLadder_() {
    const ladder = document.getElementById('rwysurvey-inclinometer-ladder');
    if (!ladder) return;
    ladder.innerHTML = '';
    const ppd = runwayInclinometerState.pixelsPerDegree;
    for (let deg = -30; deg <= 30; deg += 5) {
      const line = document.createElement('div');
      line.style.position = 'absolute';
      line.style.left = '10%';
      line.style.width = '80%';
      line.style.height = (deg % 10 === 0) ? '2px' : '1px';
      line.style.background = 'rgba(0,255,136,' + (deg % 10 === 0 ? '0.95' : '0.55') + ')';
      line.style.top = 'calc(50% + ' + (-deg * ppd) + 'px)';
      ladder.appendChild(line);

      if (deg % 10 === 0 && deg !== 0) {
        const txt = document.createElement('div');
        txt.textContent = String(Math.abs(deg));
        txt.style.position = 'absolute';
        txt.style.right = '6%';
        txt.style.top = 'calc(50% + ' + (-deg * ppd - 9) + 'px)';
        txt.style.color = 'rgba(0,255,136,0.85)';
        txt.style.fontSize = '12px';
        txt.style.fontWeight = '700';
        ladder.appendChild(txt);
      }
    }
  }

  function rwyInclinometerUpdateReadout_(angleDeg) {
    const degEl = document.getElementById('rwysurvey-inclinometer-deg');
    const pctEl = document.getElementById('rwysurvey-inclinometer-pct');
    const a = rwyInclinometerClamp_(angleDeg, -30, 30);
    runwayInclinometerState.currentDisplayAngle = a;
    const pct = Math.tan(a * Math.PI / 180) * 100;
    if (degEl) degEl.textContent = a.toFixed(1) + '°';
    if (pctEl) pctEl.textContent = pct.toFixed(1) + '%';
  }

  function rwyInclinometerRenderHud_() {
    const horizon = document.getElementById('rwysurvey-inclinometer-horizon');
    const ladder = document.getElementById('rwysurvey-inclinometer-ladder');
    const target = runwayInclinometerState.target;
    const ppd = runwayInclinometerState.pixelsPerDegree;
    const pitch = rwyInclinometerClamp_(runwayInclinometerState.filteredPitch, -30, 30);
    const offsetY = pitch * ppd;
    if (horizon) horizon.style.transform = 'translateY(' + offsetY + 'px)';
    if (ladder) ladder.style.transform = 'translateY(' + offsetY + 'px)';

    let displayAngle = pitch;
    if (target) {
      const hud = document.getElementById('rwysurvey-inclinometer-hud');
      if (hud) {
        const rect = hud.getBoundingClientRect();
        const centerY = rect.top + (rect.height / 2);
        const pixelDelta = centerY - target.y;
        const targetOffsetDeg = pixelDelta / ppd;
        displayAngle = pitch + targetOffsetDeg;
      }
    }
    rwyInclinometerUpdateReadout_(displayAngle);
  }

  function rwyInclinometerHandleOrientation_(event) {
    if (!event || typeof event.beta !== 'number') return;
    const st = runwayInclinometerState;
    st.rawBeta = Number(event.beta);
    const corrected = st.rawBeta - Number(st.zeroOffset || 0);
    const alpha = 0.14;
    st.filteredPitch = (st.filteredPitch * (1 - alpha)) + (corrected * alpha);
    rwyInclinometerRenderHud_();
  }

  async function rwyInclinometerRequestOrientation_() {
    if (!window.DeviceOrientationEvent) return true;
    try {
      if (typeof DeviceOrientationEvent.requestPermission === 'function') {
        const status = await DeviceOrientationEvent.requestPermission();
        return status === 'granted';
      }
      return true;
    } catch (err) {
      return false;
    }
  }

  async function rwyInclinometerStartCamera_() {
    const video = document.getElementById('rwysurvey-inclinometer-video');
    if (!window.isSecureContext) {
      throw new Error('Camera requires HTTPS secure context');
    }
    if (!video || !navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      throw new Error('Camera API unavailable in this browser context');
    }

    video.setAttribute('playsinline', 'true');
    video.setAttribute('autoplay', 'true');
    video.muted = true;

    const tries = [
      { audio: false, video: { facingMode: { ideal: 'environment' }, width: { ideal: 1920 }, height: { ideal: 1080 } } },
      { audio: false, video: { facingMode: { exact: 'environment' } } },
      { audio: false, video: { facingMode: 'environment' } },
      { audio: false, video: true }
    ];

    let stream = null;
    let lastErr = null;
    for (let i = 0; i < tries.length; i++) {
      try {
        stream = await navigator.mediaDevices.getUserMedia(tries[i]);
        break;
      } catch (err) {
        lastErr = err;
      }
    }

    if (!stream) {
      const errName = String(lastErr && lastErr.name || 'Error');
      const errMsg = String(lastErr && lastErr.message || 'Unknown camera error');
      throw new Error(errName + ': ' + errMsg);
    }

    runwayInclinometerState.stream = stream;
    video.srcObject = stream;
    try {
      await video.play();
    } catch (e) {
      // iOS can still show first frame even if autoplay promise rejects.
    }
  }

  function rwyInclinometerCameraErrorText_(err) {
    const msg = String(err && err.message || err || '').toLowerCase();
    const host = String((window.location && window.location.hostname) || 'este site');
    if (!window.isSecureContext) return 'Abra pelo link HTTPS do web app (câmera bloqueada fora de contexto seguro).';
    if (msg.indexOf('notallowederror') >= 0 || msg.indexOf('permission') >= 0 || msg.indexOf('denied') >= 0) {
      if (window.top !== window.self) {
        return 'Permissão de câmera negada em visualização embutida. Abra em aba nova no Safari e permita câmera para ' + host + '. (Localização não libera câmera)';
      }
      return 'Permissão de câmera negada para ' + host + '. No Safari (aA > Ajustes do Site), defina Câmera = Permitir. Localização não libera câmera.';
    }
    if (msg.indexOf('notreadableerror') >= 0 || msg.indexOf('trackstarterror') >= 0 || msg.indexOf('in use') >= 0) {
      return 'Câmera ocupada por outro app/aba. Feche outros usos da câmera e tente novamente.';
    }
    if (msg.indexOf('notfounderror') >= 0 || msg.indexOf('overconstrained') >= 0) {
      return 'Câmera traseira não encontrada. Tentando câmera padrão.';
    }
    return 'Não foi possível abrir a câmera neste dispositivo/browser.';
  }

  function rwyInclinometerStopCamera_() {
    const st = runwayInclinometerState;
    if (st.stream) {
      try {
        st.stream.getTracks().forEach(function(track) { track.stop(); });
      } catch (e) {}
      st.stream = null;
    }
    const video = document.getElementById('rwysurvey-inclinometer-video');
    if (video) video.srcObject = null;
  }

  async function openRunwaySurveyInclinometer(targetInputId, saveLabelOverride) {
    const st = runwayInclinometerState;
    st.saveInputId = String(targetInputId || 'rwysurvey-obs-a-popup');
    const active = (runwaySurveyToolState.ui && runwaySurveyToolState.ui.activeObstaclePrompt) || { corner: 'A', distanceM: 50 };
    st.saveLabel = String(saveLabelOverride || ('Salvar ângulo ' + Number(active.distanceM || 50) + 'm'));
    const saveBtn = document.getElementById('rwysurvey-inclinometer-save-btn');
    if (saveBtn) saveBtn.textContent = st.saveLabel.toUpperCase();

    const modal = document.getElementById('rwysurvey-inclinometer-modal');
    if (modal) modal.style.display = 'block';
    const meta = document.getElementById('rwysurvey-inclinometer-meta');
    if (meta) meta.textContent = 'Solicitando câmera e sensores...';

    const hud = document.getElementById('rwysurvey-inclinometer-hud');
    if (hud) {
      const rect = hud.getBoundingClientRect();
      st.pixelsPerDegree = Math.max(4, Math.min(11, rect.height / 58));
    }
    rwyInclinometerBuildLadder_();
    rwyInclinometerClearTarget();

    const granted = await rwyInclinometerRequestOrientation_();
    if (!granted) {
      if (meta) meta.textContent = 'Movimento negado; câmera ainda pode funcionar. Ative Motion & Orientation para horizonte dinâmico.';
      if (window.M) M.toast({ html: 'Motion denied: camera mode only', classes: 'orange' });
    } else if (!st.orientationBound) {
      window.addEventListener('deviceorientation', rwyInclinometerHandleOrientation_, true);
      st.orientationBound = true;
    }

    try {
      await rwyInclinometerStartCamera_();
      if (meta) meta.textContent = granted
        ? 'Toque no alvo para medir o ângulo'
        : 'Câmera ativa sem sensores. Ative Motion & Orientation para medir inclinação.';
    } catch (e) {
      const human = rwyInclinometerCameraErrorText_(e);
      const errMsg = String(e && e.message || '').toLowerCase();
      const isPermissionError = errMsg.indexOf('notallowederror') >= 0 || errMsg.indexOf('permission') >= 0 || errMsg.indexOf('denied') >= 0;
      const isEmbedded = (window.top !== window.self);
      if (meta) meta.textContent = human;
      if (window.M) M.toast({ html: human, classes: 'orange' });

      if (isPermissionError && isEmbedded) {
        if (meta) meta.textContent = 'Modo embutido bloqueia câmera. Use o botão abaixo para abrir em aba nova.';
        if (window.M) {
          M.toast({
            html: '<button class="btn blue darken-2" onclick="rwyInclinometerOpenStandaloneTab_();">ABRIR MODO CÂMERA STANDALONE</button>',
            displayLength: 12000,
            classes: 'blue-grey darken-3'
          });
        }
      }
    }
  }

  function closeRunwaySurveyInclinometer() {
    rwyInclinometerStopCamera_();
    const modal = document.getElementById('rwysurvey-inclinometer-modal');
    if (modal) modal.style.display = 'none';
  }

  function rwyInclinometerHandleTap(event) {
    const hud = document.getElementById('rwysurvey-inclinometer-hud');
    const marker = document.getElementById('rwysurvey-inclinometer-target');
    if (!hud || !marker) return;
    const rect = hud.getBoundingClientRect();
    const x = Number(event.clientX || (event.touches && event.touches[0] && event.touches[0].clientX) || 0);
    const y = Number(event.clientY || (event.touches && event.touches[0] && event.touches[0].clientY) || 0);
    const clampedX = Math.max(rect.left + 8, Math.min(rect.right - 8, x));
    const clampedY = Math.max(rect.top + 8, Math.min(rect.bottom - 8, y));
    runwayInclinometerState.target = { x: clampedX, y: clampedY };
    marker.style.left = (clampedX - rect.left) + 'px';
    marker.style.top = (clampedY - rect.top) + 'px';
    marker.style.display = 'block';
    const meta = document.getElementById('rwysurvey-inclinometer-meta');
    if (meta) meta.textContent = 'Alvo travado · ajuste a mira e salve';
    rwyInclinometerRenderHud_();
  }

  function rwyInclinometerClearTarget() {
    runwayInclinometerState.target = null;
    const marker = document.getElementById('rwysurvey-inclinometer-target');
    if (marker) marker.style.display = 'none';
    const meta = document.getElementById('rwysurvey-inclinometer-meta');
    if (meta) meta.textContent = 'Toque no alvo para medir o ângulo';
    rwyInclinometerRenderHud_();
  }

  function rwyInclinometerSetZero() {
    const st = runwayInclinometerState;
    if (typeof st.rawBeta !== 'number') {
      if (window.M) M.toast({ html: 'No sensor data yet', classes: 'orange' });
      return;
    }
    st.zeroOffset = Number(st.rawBeta || 0);
    st.filteredPitch = 0;
    rwyInclinometerRenderHud_();
    if (window.M) M.toast({ html: 'Zero calibrated', classes: 'green' });
  }

  function rwyInclinometerSave() {
    const angle = rwyInclinometerClamp_(runwayInclinometerState.currentDisplayAngle, -30, 30);
    const input = document.getElementById(runwayInclinometerState.saveInputId || 'rwysurvey-obs-a-popup');
    if (input) input.value = angle.toFixed(1);
    closeRunwaySurveyInclinometer();
    if (window.M) M.toast({ html: 'Ângulo salvo: ' + angle.toFixed(1) + '°', classes: 'green' });
  }

  window.waterRunwaySurveyState = {
    icao: '',
    selectedPair: '',
    selectedDirection: '',
    waterName: '',
    elevationFt: null,
    midpoint: null,
    created: false,
    gps: { watchId: null, tracking: false, current: null, points: [], startFix: null, lastProjectedDistanceM: 0 },
    headingDeg: 0,
    measuredLengthM: 0,
    captures: [],
    features: [],
    pendingAnglePhoto: null,
    notes: '',
    sourceRunwayRow: null
  };

  function _waterRwyResetState_() {
    window.waterRunwaySurveyState = {
      icao: '',
      selectedPair: '',
      selectedDirection: '',
      waterName: '',
      elevationFt: null,
      midpoint: null,
      created: false,
      gps: { watchId: null, tracking: false, current: null, points: [], startFix: null, lastProjectedDistanceM: 0 },
      headingDeg: 0,
      measuredLengthM: 0,
      captures: [],
      features: [],
      pendingAnglePhoto: null,
      notes: '',
      sourceRunwayRow: null
    };
  }

  function _waterRwyRunwayPairs_() {
    const out = [];
    for (let n = 1; n <= 18; n++) {
      const a = String(n).padStart(2, '0');
      const b = String(n + 18).padStart(2, '0');
      out.push(a + '/' + b);
    }
    return out;
  }

  function _waterRwyReciprocal_(ident) {
    const num = parseInt(String(ident || '').replace(/\D+/g, ''), 10);
    if (!isFinite(num) || num < 1 || num > 36) return '';
    return String((((num + 18 - 1) % 36) + 1)).padStart(2, '0');
  }

  function _waterRwyPairFromIdent_(ident) {
    const a = String(ident || '').trim().toUpperCase().replace(/^RWY\s*/i, '').replace(/[^0-9]/g, '');
    if (!a) return '';
    const main = String(parseInt(a, 10) || '').padStart(2, '0');
    if (!main || main === '00') return '';
    const b = _waterRwyReciprocal_(main);
    if (!b) return '';
    return [main, b].sort(function(x, y) { return x.localeCompare(y); }).join('/');
  }

  function _waterRwyDirectionOptionsFromPair_(pair) {
    const p = String(pair || '').trim().toUpperCase();
    if (!p) return [];
    return p.split('/').map(function(v) { return String(v || '').trim(); }).filter(Boolean);
  }

  function _waterRwyHeadingFromDirection_(ident) {
    const num = parseInt(String(ident || '').replace(/\D+/g, ''), 10);
    if (!isFinite(num) || num < 1 || num > 36) return 0;
    return num * 10;
  }

  function _waterRwySurfaceText_(row) {
    return String(row && (row.runwaySurfaceActual || row.surface || row.runwaySurfaceCondition || '') || '').trim().toUpperCase();
  }

  function _waterRwyIsWaterRow_(row) {
    const s = _waterRwySurfaceText_(row);
    return /WATER|SEAPLANE|LAGO|LAKE|RIVER|RIO|AGUA/.test(s);
  }

  function _waterRwyRowsForIcao_(icao) {
    const rows = (window.appData && Array.isArray(window.appData.airports)) ? window.appData.airports : [];
    const target = String(icao || '').trim().toUpperCase();
    return rows.filter(function(r) {
      const rowIcao = String(r && (r.airportICAO || r.icao || r.ICAO) || '').trim().toUpperCase();
      return rowIcao === target;
    });
  }

  function _waterRwyUniqueIcaos_() {
    const rows = (window.appData && Array.isArray(window.appData.airports)) ? window.appData.airports : [];
    const seen = {};
    rows.forEach(function(r) {
      const icao = String(r && (r.airportICAO || r.icao || r.ICAO) || '').trim().toUpperCase();
      if (icao) seen[icao] = true;
    });
    return Object.keys(seen).sort();
  }

  function _waterRwyAutoName_(rawName) {
    const cleaned = String(rawName || '').trim().replace(/\s+/g, '-').replace(/[^A-Za-z0-9\-]/g, '').toUpperCase();
    return 'W-' + (cleaned || 'UNNAMED');
  }

  function openWaterRunwaySurveyTool() {
    const modal = document.getElementById('water-rwysurvey-modal');
    if (!modal) return;
    _waterRwyResetState_();
    modal.style.display = 'block';

    const icaos = _waterRwyUniqueIcaos_();
    const icaoInput = document.getElementById('water-rwy-icao');
    if (icaoInput) icaoInput.value = icaos.length ? icaos[0] : '';
    waterRwyRenderAutoName();
    waterRwyInitPairOptions();
    waterRwyRefreshRunwayOptions();
    waterRwySetAnglePhotoStatus_('No angle photo selected.');
    waterRwyRenderLists();
  }

  function closeWaterRunwaySurveyTool() {
    waterRwyStopGps(true);
    const modal = document.getElementById('water-rwysurvey-modal');
    if (modal) modal.style.display = 'none';
  }

  function waterRwyInitPairOptions() {
    const sel = document.getElementById('water-rwy-pair');
    if (!sel) return;
    const pairs = _waterRwyRunwayPairs_();
    sel.innerHTML = '';
    pairs.forEach(function(p) {
      const opt = document.createElement('option');
      opt.value = p;
      opt.textContent = p;
      sel.appendChild(opt);
    });
    sel.value = pairs[0] || '';
  }

  function waterRwyRenderAutoName() {
    const raw = String((document.getElementById('water-rwy-name') || {}).value || '').trim();
    const auto = _waterRwyAutoName_(raw);
    const el = document.getElementById('water-rwy-auto-name');
    if (el) el.textContent = auto;
    waterRunwaySurveyState.waterName = auto;
  }

  function waterRwyToggleCreatePanel() {
    const panel = document.getElementById('water-rwy-create-panel');
    if (!panel) return;
    panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
  }

  function waterRwyRefreshRunwayOptions() {
    const icao = String((document.getElementById('water-rwy-icao') || {}).value || '').trim().toUpperCase();
    waterRunwaySurveyState.icao = icao;
    const rows = _waterRwyRowsForIcao_(icao).filter(_waterRwyIsWaterRow_);
    const pairMap = {};
    rows.forEach(function(r) {
      const ident = String(r && (r.runwayIdent || r.rwyIdent || '') || '').trim().toUpperCase();
      const pair = _waterRwyPairFromIdent_(ident) || ident;
      if (!pair) return;
      if (!pairMap[pair]) pairMap[pair] = r;
    });

    const sel = document.getElementById('water-rwy-existing');
    if (!sel) return;
    sel.innerHTML = '';
    const ph = document.createElement('option');
    ph.value = '';
    ph.textContent = rows.length ? '-- Select existing water runway --' : '-- No water runways found --';
    sel.appendChild(ph);
    Object.keys(pairMap).sort().forEach(function(pair) {
      const opt = document.createElement('option');
      opt.value = pair;
      opt.textContent = pair;
      sel.appendChild(opt);
    });
    sel.value = '';
    waterRwyApplyDirectionOptions('');
  }

  function waterRwyApplyExistingSelection() {
    const pair = String((document.getElementById('water-rwy-existing') || {}).value || '').trim().toUpperCase();
    if (!pair) return;
    waterRunwaySurveyState.selectedPair = pair;
    waterRunwaySurveyState.created = true;
    waterRwyApplyDirectionOptions(pair);
    const status = document.getElementById('water-rwy-create-status');
    if (status) status.textContent = 'Using existing water runway pair ' + pair + '.';
  }

  function waterRwyApplyDirectionOptions(pair) {
    const dirSel = document.getElementById('water-rwy-direction');
    if (!dirSel) return;
    const dirs = _waterRwyDirectionOptionsFromPair_(pair);
    dirSel.innerHTML = '';
    dirs.forEach(function(d) {
      const opt = document.createElement('option');
      opt.value = d;
      opt.textContent = d;
      dirSel.appendChild(opt);
    });
    const first = dirs[0] || '';
    dirSel.value = first;
    waterRunwaySurveyState.selectedDirection = first;
    waterRunwaySurveyState.headingDeg = _waterRwyHeadingFromDirection_(first);
  }

  function waterRwyMarkMidpointHere() {
    if (!navigator.geolocation) {
      if (window.M) M.toast({ html: 'Geolocation not available', classes: 'red' });
      return;
    }
    navigator.geolocation.getCurrentPosition(function(pos) {
      const c = pos && pos.coords ? pos.coords : null;
      if (!c) return;
      waterRunwaySurveyState.midpoint = {
        lat: Number(c.latitude || 0),
        lon: Number(c.longitude || 0),
        acc: Number(c.accuracy || 9999),
        markedAt: new Date().toISOString()
      };
      const el = document.getElementById('water-rwy-midpoint');
      if (el) {
        el.textContent = 'Midpoint: ' + waterRunwaySurveyState.midpoint.lat.toFixed(6) + ', ' + waterRunwaySurveyState.midpoint.lon.toFixed(6)
          + ' (±' + Math.round(waterRunwaySurveyState.midpoint.acc) + 'm)';
      }
      if (window.M) M.toast({ html: 'Midpoint marked', classes: 'green' });
    }, function(err) {
      if (window.M) M.toast({ html: 'GPS error: ' + (err && err.message ? err.message : 'unknown'), classes: 'red' });
    }, { enableHighAccuracy: true, maximumAge: 0, timeout: 12000 });
  }

  function waterRwySaveInitialRunway() {
    const st = waterRunwaySurveyState;
    const icao = String((document.getElementById('water-rwy-icao') || {}).value || '').trim().toUpperCase();
    const pair = String((document.getElementById('water-rwy-pair') || {}).value || '').trim().toUpperCase();
    const elevRaw = (document.getElementById('water-rwy-elevation') || {}).value;
    const elevationFt = elevRaw === '' ? null : Number(elevRaw);
    waterRwyRenderAutoName();
    if (!icao) {
      if (window.M) M.toast({ html: 'ICAO is required', classes: 'orange' });
      return;
    }
    if (!pair) {
      if (window.M) M.toast({ html: 'Runway pair is required', classes: 'orange' });
      return;
    }
    if (!st.midpoint || !isFinite(st.midpoint.lat) || !isFinite(st.midpoint.lon)) {
      if (window.M) M.toast({ html: 'Mark midpoint first', classes: 'orange' });
      return;
    }

    window.runOrQueueServerAction({
      method: 'addWaterRunwayToDatabase',
      args: [{
        icao: icao,
        runwayPair: pair,
        waterName: st.waterName,
        elevationFt: isFinite(elevationFt) ? elevationFt : null,
        midpointLat: st.midpoint.lat,
        midpointLon: st.midpoint.lon
      }],
      label: 'Create water runway ' + icao + ' ' + pair
    }, {
      onSuccess: function(resp) {
        if (!resp || !resp.success) {
          if (window.M) M.toast({ html: (resp && resp.error) ? resp.error : 'Failed to save runway', classes: 'red' });
          return;
        }
        st.icao = icao;
        st.selectedPair = pair;
        st.created = true;
        st.elevationFt = isFinite(elevationFt) ? elevationFt : null;
        waterRwyApplyDirectionOptions(pair);
        waterRwyPushLocalAirportRows_(icao, pair, st.waterName, st.elevationFt, st.midpoint);
        waterRwyRefreshRunwayOptions();
        const status = document.getElementById('water-rwy-create-status');
        if (status) status.textContent = 'Saved ' + pair + ' for ' + icao + '. Start measuring.';
        if (window.M) M.toast({ html: 'Initial water runway saved', classes: 'green' });
      },
      onQueued: function() {
        st.icao = icao;
        st.selectedPair = pair;
        st.created = true;
        st.elevationFt = isFinite(elevationFt) ? elevationFt : null;
        waterRwyApplyDirectionOptions(pair);
        if (window.M) M.toast({ html: 'Offline: runway creation queued', classes: 'orange' });
      },
      onFailure: function(err) {
        if (window.M) M.toast({ html: 'Create failed: ' + (err && err.message ? err.message : String(err)), classes: 'red' });
      }
    });
  }

  function waterRwyPushLocalAirportRows_(icao, pair, waterName, elevationFt, midpoint) {
    if (!window.appData) window.appData = {};
    if (!Array.isArray(window.appData.airports)) window.appData.airports = [];
    const dirs = _waterRwyDirectionOptionsFromPair_(pair);
    dirs.forEach(function(ident) {
      const exists = window.appData.airports.some(function(r) {
        const rowIcao = String(r && (r.airportICAO || r.icao || r.ICAO) || '').trim().toUpperCase();
        const rowRwy = String(r && (r.runwayIdent || r.rwyIdent || '') || '').trim().toUpperCase();
        return rowIcao === icao && rowRwy === ident;
      });
      if (exists) return;
      window.appData.airports.push({
        airportICAO: icao,
        icao: icao,
        runwayIdent: ident,
        runwayHeading: _waterRwyHeadingFromDirection_(ident),
        runwaySurfaceActual: 'WATER',
        elevationFt: elevationFt,
        lat: midpoint && midpoint.lat,
        lon: midpoint && midpoint.lon,
        nome: waterName
      });
    });
  }

  function _waterRwyDistanceMeters_(a, b) {
    return _rwySurveyDistanceMetersBetween_(a, b);
  }

  function _waterRwyProjectedDistance_(startFix, currentFix, headingDeg) {
    if (!startFix || !currentFix) return 0;
    const latRad = _rwySurveyDeg2Rad_(Number(startFix.lat || 0));
    const north = (Number(currentFix.lat || 0) - Number(startFix.lat || 0)) * 111320;
    const east = (Number(currentFix.lon || 0) - Number(startFix.lon || 0)) * 111320 * Math.max(Math.cos(latRad), 0.2);
    const br = _rwySurveyDeg2Rad_(Number(headingDeg || 0));
    const ux = Math.sin(br);
    const uy = Math.cos(br);
    return Math.max(0, (east * ux) + (north * uy));
  }

  function waterRwyStartGps() {
    const st = waterRunwaySurveyState;
    if (!st.created && !st.selectedPair) {
      if (window.M) M.toast({ html: 'Pick existing or create runway first', classes: 'orange' });
      return;
    }
    st.selectedDirection = String((document.getElementById('water-rwy-direction') || {}).value || st.selectedDirection || '').trim().toUpperCase();
    st.headingDeg = _waterRwyHeadingFromDirection_(st.selectedDirection);
    if (!st.headingDeg) {
      if (window.M) M.toast({ html: 'Select measurement runway direction', classes: 'orange' });
      return;
    }
    if (!navigator.geolocation) {
      if (window.M) M.toast({ html: 'Geolocation not available', classes: 'red' });
      return;
    }
    if (st.gps.watchId != null && navigator.geolocation) navigator.geolocation.clearWatch(st.gps.watchId);
    st.gps.watchId = navigator.geolocation.watchPosition(function(pos) {
      const c = pos && pos.coords ? pos.coords : null;
      if (!c) return;
      const fix = {
        lat: Number(c.latitude || 0),
        lon: Number(c.longitude || 0),
        acc: Number(c.accuracy || 9999),
        ts: Date.now()
      };
      st.gps.current = fix;
      st.gps.points.push(fix);
      if (!st.gps.startFix) st.gps.startFix = fix;
      st.gps.lastProjectedDistanceM = Math.round(_waterRwyProjectedDistance_(st.gps.startFix, fix, st.headingDeg));
      const live = document.getElementById('water-rwy-live');
      if (live) {
        live.textContent = 'Tracking ' + st.selectedDirection + ' · distance ' + st.gps.lastProjectedDistanceM + 'm · acc ±' + Math.round(fix.acc) + 'm';
      }
    }, function(err) {
      if (window.M) M.toast({ html: 'GPS error: ' + (err && err.message ? err.message : 'unknown'), classes: 'red' });
    }, { enableHighAccuracy: true, maximumAge: 0, timeout: 12000 });
    st.gps.tracking = true;
    const btn = document.getElementById('water-rwy-gps-toggle');
    if (btn) btn.textContent = 'MEASURING...';
    const live = document.getElementById('water-rwy-live');
    if (live) live.textContent = 'GPS started. Taxi straight and tap angle capture.';
  }

  function waterRwyStopGps(hard) {
    const st = waterRunwaySurveyState;
    if (st.gps.watchId != null && navigator.geolocation) {
      navigator.geolocation.clearWatch(st.gps.watchId);
      st.gps.watchId = null;
    }
    st.gps.tracking = false;
    if (hard) {
      st.gps.current = null;
      st.gps.points = [];
      st.gps.startFix = null;
      st.gps.lastProjectedDistanceM = 0;
    }
    const btn = document.getElementById('water-rwy-gps-toggle');
    if (btn) btn.textContent = 'START MEASURING';
  }

  function waterRwyToggleMeasuring() {
    if (waterRunwaySurveyState.gps.tracking) {
      waterRwyStopGps(false);
      const live = document.getElementById('water-rwy-live');
      if (live) live.textContent = 'GPS paused.';
      return;
    }
    waterRwyStartGps();
  }

  function waterRwyStopAndStoreLength() {
    const st = waterRunwaySurveyState;
    st.measuredLengthM = Math.max(st.measuredLengthM, Number(st.gps.lastProjectedDistanceM || 0));
    waterRwyStopGps(false);
    const live = document.getElementById('water-rwy-live');
    if (live) live.textContent = 'Run ended at ' + Math.round(st.measuredLengthM || 0) + 'm.';
  }

  function waterRwyTapMeasureAngle() {
    const st = waterRunwaySurveyState;
    if (!st.gps.tracking || !st.gps.current || !st.gps.startFix) {
      if (window.M) M.toast({ html: 'Start measuring first', classes: 'orange' });
      return;
    }
    const dist = Math.round(Number(st.gps.lastProjectedDistanceM || 0));
    if (!st.captures.length && dist < 200) {
      if (window.M) M.toast({ html: 'First angle capture needs at least 200m from start', classes: 'orange' });
      return;
    }
    let angle = Number((document.getElementById('water-rwy-angle-input') || {}).value);
    if (!isFinite(angle)) {
      const promptRaw = window.prompt('Enter obstacle angle in degrees');
      angle = Number(promptRaw);
    }
    if (!isFinite(angle)) {
      if (window.M) M.toast({ html: 'Angle is required', classes: 'orange' });
      return;
    }
    const note = String((document.getElementById('water-rwy-angle-note') || {}).value || '').trim();
    st.captures.push({
      distanceM: dist,
      angleDeg: Number(angle.toFixed(1)),
      note: note,
      photo: st.pendingAnglePhoto,
      lat: Number(st.gps.current.lat || 0),
      lon: Number(st.gps.current.lon || 0),
      ts: new Date().toISOString()
    });
    st.pendingAnglePhoto = null;
    waterRwySetAnglePhotoStatus_('No angle photo selected.');
    const angleInput = document.getElementById('water-rwy-angle-input');
    const angleNote = document.getElementById('water-rwy-angle-note');
    if (angleInput) angleInput.value = '';
    if (angleNote) angleNote.value = '';
    waterRwyRenderLists();
    if (window.M) M.toast({ html: 'Angle captured at ' + dist + 'm', classes: 'green' });
  }

  function waterRwyAddFeatureAtCurrent() {
    const st = waterRunwaySurveyState;
    if (!st.gps.current) {
      if (window.M) M.toast({ html: 'No GPS fix yet', classes: 'orange' });
      return;
    }
    const type = String((document.getElementById('water-rwy-feature-type') || {}).value || 'other').trim();
    const note = String((document.getElementById('water-rwy-feature-note') || {}).value || '').trim();
    st.features.push({
      type: type,
      note: note,
      distanceM: Math.round(Number(st.gps.lastProjectedDistanceM || 0)),
      fromThreshold: String(st.selectedDirection || '').trim(),
      gps: { lat: Number(st.gps.current.lat || 0), lon: Number(st.gps.current.lon || 0), acc: Number(st.gps.current.acc || 9999) },
      capturedAt: new Date().toISOString()
    });
    const noteEl = document.getElementById('water-rwy-feature-note');
    if (noteEl) noteEl.value = '';
    waterRwyRenderLists();
    if (window.M) M.toast({ html: 'Feature added', classes: 'green' });
  }

  function waterRwyRenderLists() {
    const st = waterRunwaySurveyState;
    const angleEl = document.getElementById('water-rwy-angle-list');
    const featEl = document.getElementById('water-rwy-feature-list');
    const sumEl = document.getElementById('water-rwy-summary');
    if (sumEl) sumEl.textContent = 'Angles: ' + st.captures.length + ' · Features: ' + st.features.length;

    if (angleEl) {
      if (!st.captures.length) {
        angleEl.textContent = 'No angles captured.';
      } else {
        angleEl.innerHTML = st.captures.map(function(c, i) {
          return '<div style="padding:4px 0; border-bottom:1px solid #eceff1;">#' + (i + 1) + ' · ' + c.distanceM + 'm · ' + c.angleDeg + '°' + (c.note ? ' · ' + c.note : '') + '</div>';
        }).join('');
      }
    }

    if (featEl) {
      if (!st.features.length) {
        featEl.textContent = 'No features captured.';
      } else {
        featEl.innerHTML = st.features.map(function(f, i) {
          return '<div style="padding:4px 0; border-bottom:1px solid #eceff1;">#' + (i + 1) + ' · ' + String(f.type || '') + ' · ' + String(f.distanceM || 0) + 'm' + (f.note ? ' · ' + f.note : '') + '</div>';
        }).join('');
      }
    }
  }

  function waterRwySetAnglePhotoStatus_(text) {
    const el = document.getElementById('water-rwy-angle-photo-status');
    if (el) el.textContent = String(text || 'No angle photo selected.');
  }

  function waterRwyPickAnglePhoto() {
    const st = waterRunwaySurveyState;
    if (!st.icao) {
      if (window.M) M.toast({ html: 'Set ICAO first', classes: 'orange' });
      return;
    }
    const input = document.getElementById('water-rwy-angle-photo-input');
    if (input) {
      input.value = '';
      input.click();
    }
  }

  function onWaterRwyAnglePhotoInputChange(input) {
    try {
      const file = input && input.files && input.files[0] ? input.files[0] : null;
      if (!file) return;
      waterRwyUploadAnglePhoto_(file);
    } finally {
      if (input) input.value = '';
    }
  }

  async function waterRwyUploadAnglePhoto_(file) {
    const st = waterRunwaySurveyState;
    const icao = String(st.icao || '').trim().toUpperCase();
    const rwyIdent = String(st.selectedDirection || '').trim().toUpperCase();
    if (!icao) return;
    waterRwySetAnglePhotoStatus_('Uploading photo...');
    try {
      const rawDataUrl = await _rwySurveyReadFileAsDataUrl_(file);
      const compressedDataUrl = await _rwySurveyShrinkImageDataUrl_(rawDataUrl);
      const parts = compressedDataUrl.split(',');
      if (parts.length < 2) throw new Error('Invalid image data');
      const header = parts[0] || '';
      const mimeMatch = header.match(/data:([^;]+);base64/i);
      const mimeType = (mimeMatch && mimeMatch[1]) ? mimeMatch[1] : 'image/jpeg';
      const base64Data = parts[1];
      const ts = new Date();
      const stamp = ts.toISOString().replace(/[.:]/g, '-');
      const fileName = _rwySurveyPhotoSafeName_(icao + '_' + String(rwyIdent || 'WTRWY') + '_ANGLE_' + stamp + '_' + (file.name || 'photo.jpg'));

      window.runOrQueueServerAction({
        method: 'uploadRunwaySurveyPhoto',
        args: [{
          icao: icao,
          rwyIdent: rwyIdent,
          fileName: fileName,
          mimeType: mimeType,
          base64Data: base64Data,
          source: 'water_angle',
          takenAt: ts.toISOString()
        }],
        label: 'Water runway angle photo ' + icao + ' ' + rwyIdent
      }, {
        onSuccess: function(resp) {
          st.pendingAnglePhoto = {
            name: fileName,
            status: 'uploaded',
            fileId: resp && resp.fileId ? String(resp.fileId) : '',
            url: resp && resp.url ? String(resp.url) : '',
            source: 'water_angle',
            takenAt: ts.toISOString()
          };
          waterRwySetAnglePhotoStatus_('Photo ready: ' + fileName);
          if (window.M) M.toast({ html: 'Angle photo uploaded', classes: 'green' });
        },
        onQueued: function() {
          st.pendingAnglePhoto = {
            name: fileName,
            status: 'queued',
            fileId: '',
            url: '',
            source: 'water_angle',
            takenAt: ts.toISOString()
          };
          waterRwySetAnglePhotoStatus_('Photo queued offline: ' + fileName);
          if (window.M) M.toast({ html: 'Offline: angle photo queued', classes: 'orange' });
        },
        onFailure: function(err) {
          st.pendingAnglePhoto = null;
          waterRwySetAnglePhotoStatus_('Photo upload failed');
          if (window.M) M.toast({ html: 'Angle photo upload failed: ' + (err && err.message ? err.message : String(err)), classes: 'red' });
        }
      });
    } catch (e) {
      st.pendingAnglePhoto = null;
      waterRwySetAnglePhotoStatus_('Photo upload failed');
      if (window.M) M.toast({ html: 'Angle photo error: ' + (e && e.message ? e.message : String(e)), classes: 'red' });
    }
  }

  function openWaterRunwayPreview() {
    const st = waterRunwaySurveyState;
    const existing = document.getElementById('water-rwy-preview-modal');
    if (existing) existing.remove();
    const modal = document.createElement('div');
    modal.id = 'water-rwy-preview-modal';
    modal.style.cssText = 'position:fixed; inset:0; z-index:10270; background:rgba(0,0,0,0.66); display:flex; align-items:center; justify-content:center; padding:12px;';
    const card = document.createElement('div');
    card.style.cssText = 'width:min(720px,100%); max-height:90vh; overflow:auto; background:#fff; border-radius:10px; box-shadow:0 10px 30px rgba(0,0,0,0.35);';
    card.innerHTML = ''
      + '<div style="display:flex; justify-content:space-between; align-items:center; background:#00695c; color:#fff; padding:10px 12px;">'
      + '<div><div style="font-weight:900;">Water Runway Preview</div><div style="font-size:0.78rem; opacity:0.88;">' + String(st.icao || '--') + ' · ' + String(st.selectedPair || '--') + '</div></div>'
      + '<button onclick="this.closest(\'#water-rwy-preview-modal\').remove()" style="border:none; background:rgba(255,255,255,0.2); color:#fff; border-radius:6px; padding:6px 10px;">✕</button>'
      + '</div>'
      + '<div style="padding:12px; display:grid; gap:10px;">'
      + '<div style="font-size:0.85rem;"><b>Direction:</b> ' + String(st.selectedDirection || '--') + ' · <b>Measured length:</b> ' + Math.round(Number(st.measuredLengthM || st.gps.lastProjectedDistanceM || 0)) + 'm</div>'
      + '<div style="font-size:0.85rem;"><b>Angles:</b> ' + st.captures.length + ' · <b>Features:</b> ' + st.features.length + '</div>'
      + '<pre style="margin:0; max-height:360px; overflow:auto; background:#f5f7fa; border:1px solid #e0e6ed; border-radius:8px; padding:10px; font-size:0.73rem;">'
      + JSON.stringify({
        icao: st.icao,
        runwayPair: st.selectedPair,
        direction: st.selectedDirection,
        midpoint: st.midpoint,
        measuredLengthM: Math.round(Number(st.measuredLengthM || st.gps.lastProjectedDistanceM || 0)),
        captures: st.captures,
        features: st.features,
        notes: String((document.getElementById('water-rwy-notes') || {}).value || '')
      }, null, 2)
      + '</pre>'
      + '</div>';
    modal.appendChild(card);
    modal.onclick = function(ev) { if (ev.target === modal) modal.remove(); };
    document.body.appendChild(modal);
  }

  function submitWaterRunwaySurvey() {
    const st = waterRunwaySurveyState;
    const icao = String((document.getElementById('water-rwy-icao') || {}).value || st.icao || '').trim().toUpperCase();
    const direction = String((document.getElementById('water-rwy-direction') || {}).value || st.selectedDirection || '').trim().toUpperCase();
    const pair = String(st.selectedPair || (document.getElementById('water-rwy-existing') || {}).value || '').trim().toUpperCase();
    const notes = String((document.getElementById('water-rwy-notes') || {}).value || '').trim();
    const measuredLengthM = Math.round(Number(st.measuredLengthM || st.gps.lastProjectedDistanceM || 0));

    if (!icao || !direction) {
      if (window.M) M.toast({ html: 'ICAO and direction are required', classes: 'orange' });
      return;
    }
    if (measuredLengthM <= 0) {
      if (window.M) M.toast({ html: 'Measure runway first', classes: 'orange' });
      return;
    }

    const startFix = st.gps.startFix || st.midpoint || null;
    const endFix = st.gps.current || null;
    const btn = document.getElementById('water-rwy-submit');
    if (btn) { btn.disabled = true; btn.textContent = 'SUBMITTING...'; }

    const payload = {
      icao: icao,
      rwyIdent: direction,
      pilotName: String((window.currentBriefingMission && window.currentBriefingMission.pilot) || 'Unknown Pilot'),
      pilotEmail: String((window.currentBriefingMission && window.currentBriefingMission.meta && window.currentBriefingMission.meta.pilotEmail) || ''),
      notes: notes,
      features: st.features,
      survey: {
        mode: 'WATER_STRAIGHT',
        runwayPair: pair,
        runwayWaterName: st.waterName,
        lengthM: measuredLengthM,
        widthM: 0,
        surface: 'WATER',
        elevationFt: st.elevationFt,
        features: st.features,
        markers: st.features.map(function(f) {
          return { label: f.type, distanceM: f.distanceM, fromThreshold: f.fromThreshold, gps: f.gps, notes: f.note || '' };
        }),
        obstacleAngles50m: st.captures,
        obstacles: st.captures,
        slopeSegments: [],
        axis: { headingDeg: _waterRwyHeadingFromDirection_(direction), lengthM: measuredLengthM },
        thresholds: {
          a: startFix ? { ident: direction, lat: Number(startFix.lat || 0), lon: Number(startFix.lon || 0) } : {},
          b: endFix ? { ident: _waterRwyReciprocal_(direction), lat: Number(endFix.lat || 0), lon: Number(endFix.lon || 0) } : {}
        },
        midpoint: st.midpoint,
        perimeterTrace: Array.isArray(st.gps.points) ? st.gps.points : [],
        notes: notes,
        gpsSummary: {
          points: Array.isArray(st.gps.points) ? st.gps.points.length : 0,
          lastProjectedDistanceM: Number(st.gps.lastProjectedDistanceM || 0),
          measuredLengthM: measuredLengthM
        }
      },
      official: {
        lengthM: measuredLengthM,
        widthM: 0,
        surface: 'WATER',
        headingDeg: _waterRwyHeadingFromDirection_(direction)
      },
      captureSummary: {
        pointCount: Array.isArray(st.gps.points) ? st.gps.points.length : 0,
        anglesCaptured: st.captures.length,
        featuresCaptured: st.features.length,
        measuredLengthM: measuredLengthM
      },
      deviceInfo: {
        userAgent: String(navigator.userAgent || ''),
        platform: String(navigator.platform || ''),
        submittedAt: new Date().toISOString(),
        source: 'water_runway_tool'
      }
    };

    window.runOrQueueServerAction({
      method: 'submitRunwaySurvey',
      args: [payload],
      label: 'Water runway survey ' + icao + ' ' + direction
    }, {
      onSuccess: function(resp) {
        if (btn) { btn.disabled = false; btn.textContent = 'SUBMIT FOR REVIEW'; }
        if (resp && resp.success) {
          if (window.M) M.toast({ html: 'Water runway survey submitted', classes: 'green' });
          closeWaterRunwaySurveyTool();
        } else if (window.M) {
          M.toast({ html: (resp && resp.error) ? resp.error : 'Submit failed', classes: 'red' });
        }
      },
      onQueued: function() {
        if (btn) { btn.disabled = false; btn.textContent = 'SUBMIT FOR REVIEW'; }
        if (window.M) M.toast({ html: 'Offline: water survey queued', classes: 'orange' });
        closeWaterRunwaySurveyTool();
      },
      onFailure: function(err) {
        if (btn) { btn.disabled = false; btn.textContent = 'SUBMIT FOR REVIEW'; }
        if (window.M) M.toast({ html: 'Submit failed: ' + (err && err.message ? err.message : String(err)), classes: 'red' });
      }
    });
  }

  window.runwaySurveyToolState = {
    icao: '',
    rwyIdent: '',
    runway: null,
    surfaceOptions: [],
    official: null,
    internalUpdatedAt: '',
    headingDeg: 0,
    lengthM: 0,
    startThresholdIdent: '',
    reciprocalThresholdIdent: '',
    thresholdA: null,
    thresholdB: null,
    gps: {
      watchId: null,
      tracking: false,
      paused: false,
      points: [],
      current: null,
      bestAcc: Infinity,
      avgAcc: 0,
      samples: 0
    },
    features: [],
    obstacleAngles50m: [],
    slopeSegments: [],
    photos: [],
    surfaceObserved: '',
    cutdownAreas: { thrA: null, thrB: null },
    perimeter: {
      lastMarkFix: null,
      startTs: null,
      cornerMarks: [],
      segments: [],
      lengthWalkedM: 0,
      widthWalkedM: 0,
      liveSinceMark: 0,
      closed: false
    },
    capture: {
      pending: false,
      kind: '',
      label: '',
      startedTs: null,
      settleMs: 2500,
      timerId: null
    },
    ui: {
      pausedByPopup: false,
      activeObstaclePrompt: { corner: 'A', distanceM: 50, fromThreshold: '', operation: 'landing', manual: false },
      obstaclePhoto: null,
      pendingObstacleSaveAfterPhoto: false
    },
    prompts: {
      a50Shown: false,
      a50Completed: false,
      a50Lat: null,
      a50Lon: null,
      a300Shown: false,
      a300Completed: false,
      a300Lat: null,
      a300Lon: null,
      c50Shown: false,
      c50Completed: false,
      c50Lat: null,
      c50Lon: null,
      c300Shown: false,
      c300Completed: false,
      c300Lat: null,
      c300Lon: null
    },
    slopeCapture: {
      active: false,
      startIdx: -1,
      startTs: null,
      pendingStartDistanceM: 0,
      pendingDistanceM: 0,
      pendingFromThreshold: ''
    },
    measureTool: {
      active: false,
      startIdx: -1,
      distanceM: 0
    },
    widthObservations: [],
    debugEvents: [],
    thresholdAlerts: {
      thrA: { m300: false, m50: false },
      thrB: { m300: false, m50: false }
    }
  };

  function _rwySurveyPerimeterDefaults_() {
    return { lastMarkFix: null, startTs: null, cornerMarks: [], segments: [], lengthWalkedM: 0, widthWalkedM: 0, liveSinceMark: 0, closed: false };
  }

  function _rwySurveyCornerSequence_() {
    return ['A', 'B', 'C', 'D', 'E'];
  }

  function _rwySurveyCaptureDefaults_() {
    return { pending: false, kind: '', label: '', startedTs: null, settleMs: 2500, timerId: null };
  }

  function _rwySurveyMeasureDefaults_() {
    return { active: false, startIdx: -1, distanceM: 0 };
  }

  function _rwySurveyUiDefaults_() {
    return {
      pausedByPopup: false,
      activeObstaclePrompt: { corner: 'A', distanceM: 50, fromThreshold: '', operation: 'landing', manual: false },
      obstaclePhoto: null,
      pendingObstacleSaveAfterPhoto: false
    };
  }

  function _rwySurveyNormalizeThresholdIdent_(raw) {
    const txt = String(raw || '').trim().toUpperCase().replace(/^RWY\s*/i, '');
    const m = txt.match(/^(\d{1,2})([LCR])?$/);
    if (!m) return txt;
    const num = parseInt(m[1], 10);
    if (!(num >= 1 && num <= 36)) return txt;
    return String(num).padStart(2, '0') + (m[2] || '');
  }

  function _rwySurveyGetActiveStartThreshold_() {
    const state = runwaySurveyToolState;
    return String(state.startThresholdIdent || state.rwyIdent || '').trim().toUpperCase();
  }

  function _rwySurveyGetOppositeStartThreshold_() {
    const state = runwaySurveyToolState;
    return String(state.reciprocalThresholdIdent || _rwySurveyReciprocalIdent_(_rwySurveyGetActiveStartThreshold_()) || '').trim().toUpperCase();
  }

  function _rwySurveyThresholdForCorner_(cornerLabel) {
    const corner = String(cornerLabel || 'A').trim().toUpperCase();
    return corner === 'C' ? _rwySurveyGetOppositeStartThreshold_() : _rwySurveyGetActiveStartThreshold_();
  }

  function _rwySurveyThresholdChoices_() {
    const raw = String(runwaySurveyToolState.rwyIdent || '').trim().toUpperCase();
    let list = [];
    if (raw.indexOf('/') >= 0) {
      list = raw.split('/').map(function(p) { return _rwySurveyNormalizeThresholdIdent_(p); }).filter(Boolean);
    } else {
      const a = _rwySurveyNormalizeThresholdIdent_(raw);
      const b = _rwySurveyNormalizeThresholdIdent_(_rwySurveyReciprocalIdent_(a));
      list = [a, b].filter(Boolean);
    }
    const uniq = [];
    list.forEach(function(v) {
      if (uniq.indexOf(v) < 0) uniq.push(v);
    });
    return uniq;
  }

  async function _rwySurveyEnsureStartThresholdChosen_() {
    const state = runwaySurveyToolState;
    if (String(state.startThresholdIdent || '').trim()) return true;

    const choices = _rwySurveyThresholdChoices_();
    const defaultValue = String(choices[0] || _rwySurveyNormalizeThresholdIdent_(state.rwyIdent || '') || '');
    const help = choices.length ? choices.join(' or ') : defaultValue;
    const message = 'Starting threshold for MARK A? Use ' + help + '.';

    // Native prompt is most reliable in this modal/PWA flow.
    const chosenRaw = window.prompt(message, defaultValue);
    let chosen = String(chosenRaw || '').trim();

    chosen = _rwySurveyNormalizeThresholdIdent_(chosen || '');
    if (!chosen) {
      if (window.M) M.toast({ html: 'Start threshold selection cancelled', classes: 'orange' });
      return false;
    }

    const valid = choices.filter(Boolean);
    if (valid.indexOf(chosen) < 0) {
      if (window.M) M.toast({ html: 'Invalid threshold. Use ' + help + '.', classes: 'orange' });
      return false;
    }

    state.startThresholdIdent = chosen;
    state.reciprocalThresholdIdent = valid.find(function(v) { return v !== chosen; }) || _rwySurveyNormalizeThresholdIdent_(_rwySurveyReciprocalIdent_(chosen));
    _rwySurveyLogEvent_('mark', 'Start threshold set', { threshold: state.startThresholdIdent, reciprocal: state.reciprocalThresholdIdent });
    renderRunwaySurveyA50Status();
    if (window.M) M.toast({ html: 'Start threshold: ' + state.startThresholdIdent, classes: 'green' });
    return true;
  }

  function _rwySurveyClearCaptureTimer_() {
    const capture = runwaySurveyToolState.capture || {};
    if (capture.timerId) {
      clearTimeout(capture.timerId);
      capture.timerId = null;
    }
  }

  function _rwySurveyCancelPendingCapture_(reason, silent) {
    const capture = runwaySurveyToolState.capture || {};
    const hadPending = !!capture.pending;
    const label = capture.label || 'capture';
    _rwySurveyClearCaptureTimer_();
    runwaySurveyToolState.capture = _rwySurveyCaptureDefaults_();
    if (hadPending && !silent) {
      _rwySurveyLogEvent_('capture', 'Capture cancelled: ' + label, reason ? { reason: reason } : null);
      if (window.M) M.toast({ html: reason || 'Capture cancelled', classes: 'blue-grey' });
    }
    renderRunwaySurveyStatus();
    renderRunwaySurveyPerimeterTally();
    renderRunwaySurveyActionButtons();
  }

  function _rwySurveyLogEvent_(kind, message, extra) {
    const state = runwaySurveyToolState;
    if (!Array.isArray(state.debugEvents)) state.debugEvents = [];
    state.debugEvents.push({
      ts: new Date().toISOString(),
      kind: String(kind || 'info'),
      message: String(message || ''),
      extra: extra || null
    });
    if (state.debugEvents.length > 80) state.debugEvents = state.debugEvents.slice(state.debugEvents.length - 80);
    renderRunwaySurveyMarkHistory();
  }

  function renderRunwaySurveyMarkHistory() {
    const el = document.getElementById('rwysurvey-mark-history');
    if (!el) return;
    const events = Array.isArray(runwaySurveyToolState.debugEvents) ? runwaySurveyToolState.debugEvents : [];
    if (!events.length) {
      el.textContent = 'Marks: none yet.';
      return;
    }
    const tail = events.slice(-4).map(function(ev) {
      const t = String(ev.ts || '').slice(11, 19);
      return t + ' ' + ev.kind.toUpperCase() + ': ' + ev.message;
    });
    el.innerHTML = '<b>Recent:</b> ' + tail.join(' &nbsp;|&nbsp; ');
  }

  function renderRunwaySurveyA50Status() {
    const el = document.getElementById('rwysurvey-a50-status');
    if (!el) return;
    const prompts = runwaySurveyToolState.prompts || {};
    const doneA50 = !!prompts.a50Completed;
    const doneA300 = !!prompts.a300Completed;
    const doneC50 = !!prompts.c50Completed;
    const doneC300 = !!prompts.c300Completed;
    const allDone = doneA50 && doneA300 && doneC50 && doneC300;
    const inProgress = !!(prompts.a50Shown || prompts.a300Shown || prompts.c50Shown || prompts.c300Shown) && !allDone;
    const startThr = _rwySurveyGetActiveStartThreshold_() || 'START';
    const oppThr = _rwySurveyGetOppositeStartThreshold_() || 'OPPOSITE';
    if (allDone) {
      el.textContent = 'Obstacle-angle checks complete: ' + startThr + '+50m ✓ | ' + startThr + '+300m ✓ | ' + oppThr + '+50m ✓ | ' + oppThr + '+300m ✓';
      el.style.background = '#e8f5e9';
      el.style.borderColor = '#a5d6a7';
      el.style.color = '#1b5e20';
      return;
    }
    if (inProgress) {
      el.textContent = 'Obstacle-angle checks: ' + startThr + '+50m ' + (doneA50 ? '✓' : '…') + ' | ' + startThr + '+300m ' + (doneA300 ? '✓' : '…') + ' | ' + oppThr + '+50m ' + (doneC50 ? '✓' : '…') + ' | ' + oppThr + '+300m ' + (doneC300 ? '✓' : '…');
      el.style.background = '#fff8e1';
      el.style.borderColor = '#ffcc80';
      el.style.color = '#6d4c41';
      return;
    }
    el.textContent = 'Obstacle-angle checks pending: ' + startThr + '+50m, ' + startThr + '+300m, ' + oppThr + '+50m, ' + oppThr + '+300m';
    el.style.background = '#fff3e0';
    el.style.borderColor = '#ffcc80';
    el.style.color = '#6d4c41';
  }

  function openRunwaySurveyFeaturePopup() {
    const gps = runwaySurveyToolState.gps || {};
    if (!gps.current) {
      if (window.M) M.toast({ html: 'No current GPS fix', classes: 'orange' });
      return;
    }
    if (Number(gps.current.acc || 9999) > 10) {
      if (window.M) M.toast({ html: 'Current GPS accuracy >10m. Wait for better fix.', classes: 'orange' });
      return;
    }
    if (gps.tracking && !gps.paused) {
      runwaySurveyToolState.ui.pausedByPopup = true;
      pauseRunwaySurveyGps();
    } else {
      runwaySurveyToolState.ui.pausedByPopup = false;
    }
    const fixEl = document.getElementById('rwysurvey-feature-fix');
    if (fixEl) {
      fixEl.textContent = 'Current fix: ±' + Math.round(Number(gps.current.acc || 0)) + 'm';
    }
    const popup = document.getElementById('rwysurvey-feature-popup');
    if (popup) popup.style.display = 'block';
  }

  function closeRunwaySurveyFeaturePopup(cancelled) {
    const popup = document.getElementById('rwysurvey-feature-popup');
    if (popup) popup.style.display = 'none';
    const gps = runwaySurveyToolState.gps || {};
    if (runwaySurveyToolState.ui && runwaySurveyToolState.ui.pausedByPopup && gps.tracking && gps.paused) {
      runwaySurveyToolState.ui.pausedByPopup = false;
      resumeRunwaySurveyGps();
    }
    if (cancelled) renderRunwaySurveyActionButtons();
  }

  function rwySurveySetFeatureType_(type) {
    const el = document.getElementById('rwysurvey-feature-type-popup');
    if (!el) return;
    el.value = String(type || '').trim().toLowerCase();
  }

  function rwySurveyGetFeatureType_() {
    const raw = String((document.getElementById('rwysurvey-feature-type-popup') || {}).value || '').trim().toLowerCase();
    return raw.replace(/\s+/g, '_');
  }

  function rwySurveySetObstacleType_(type) {
    const el = document.getElementById('rwysurvey-obs-type-popup');
    if (!el) return;
    el.value = String(type || '').trim().toLowerCase();
  }

  function rwySurveyGetObstacleType_() {
    const raw = String((document.getElementById('rwysurvey-obs-type-popup') || {}).value || '').trim().toLowerCase();
    return raw.replace(/\s+/g, '_');
  }

  function _rwySurveyObstaclePhotoStatusText_() {
    const statusEl = document.getElementById('rwysurvey-obs-photo-status');
    if (!statusEl) return;
    const photo = runwaySurveyToolState.ui && runwaySurveyToolState.ui.obstaclePhoto ? runwaySurveyToolState.ui.obstaclePhoto : null;
    if (!photo) {
      statusEl.textContent = 'No angle photo selected.';
      statusEl.style.color = '#546e7a';
      return;
    }
    const status = String(photo.status || 'uploading');
    const shortName = String(photo.name || '').slice(0, 44);
    statusEl.textContent = 'Photo: ' + shortName + ' (' + status + ')';
    statusEl.style.color = status === 'uploaded' ? '#2e7d32' : (status === 'failed' ? '#c62828' : '#546e7a');
  }

  function openRunwaySurveyObstaclePhotoCapture_() {
    const input = document.getElementById('rwysurvey-obstacle-photo-input');
    if (!input) return;
    input.value = '';
    input.click();
  }

  function _rwySurveyObstacleContextFromUi_() {
    const state = runwaySurveyToolState;
    const active = state.ui && state.ui.activeObstaclePrompt ? state.ui.activeObstaclePrompt : {};
    const startThr = _rwySurveyGetActiveStartThreshold_() || state.rwyIdent || 'A';
    const oppThr = _rwySurveyGetOppositeStartThreshold_() || _rwySurveyReciprocalIdent_(startThr);
    const thresholdRaw = String((document.getElementById('rwysurvey-obs-threshold-popup') || {}).value || active.fromThreshold || startThr).trim().toUpperCase();
    const distanceRaw = Number((document.getElementById('rwysurvey-obs-distance-popup') || {}).value);
    const distanceM = isFinite(distanceRaw) ? Math.max(1, Math.round(distanceRaw)) : Math.max(1, Math.round(Number(active.distanceM || 50)));
    const operationRaw = String((document.getElementById('rwysurvey-obs-operation-popup') || {}).value || active.operation || 'auto').trim().toLowerCase();
    let operation = operationRaw;
    if (operation === 'auto') {
      if (distanceM <= 60) operation = 'landing';
      else if (distanceM >= 250) operation = 'takeoff';
      else operation = '';
    }
    const checkpointCorner = thresholdRaw === oppThr ? 'C' : 'A';
    const thresholdRef = thresholdRaw === oppThr ? oppThr : startThr;
    return {
      fromThreshold: thresholdRef,
      checkpointCorner: checkpointCorner,
      checkpointDistanceM: distanceM,
      operation: operation,
      startThreshold: startThr,
      oppositeThreshold: oppThr
    };
  }

  function _rwySurveyBuildObstaclePhotoName_(ctx) {
    const state = runwaySurveyToolState;
    const icao = String(state.icao || 'XXXX').trim().toUpperCase() || 'XXXX';
    const rwy = String(state.rwyIdent || 'RWY').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
    const thr = String(ctx && ctx.fromThreshold || state.rwyIdent || 'THR').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
    const now = new Date();
    const y = String(now.getUTCFullYear());
    const m = String(now.getUTCMonth() + 1).padStart(2, '0');
    const d = String(now.getUTCDate()).padStart(2, '0');
    const hh = String(now.getUTCHours()).padStart(2, '0');
    const mm = String(now.getUTCMinutes()).padStart(2, '0');
    const ss = String(now.getUTCSeconds()).padStart(2, '0');
    const dist = Math.max(1, Math.round(Number(ctx && ctx.checkpointDistanceM || 0)));
    return _rwySurveyPhotoSafeName_(icao + '_' + y + m + d + '_' + hh + mm + ss + '_RWY' + rwy + '_THR' + thr + '_' + dist + 'M.jpg');
  }

  function onRunwaySurveyObstaclePhotoInputChange(input) {
    try {
      const files = input && input.files ? Array.from(input.files) : [];
      if (!files.length) return;
      const ctx = _rwySurveyObstacleContextFromUi_();
      const forcedName = _rwySurveyBuildObstaclePhotoName_(ctx);
      _rwySurveyUploadSelectedPhotos_(files.slice(0, 1), 'obstacle-angle', { forcedFileName: forcedName }).then(function(localRefs) {
        const photo = Array.isArray(localRefs) && localRefs.length ? localRefs[0] : null;
        runwaySurveyToolState.ui.obstaclePhoto = photo;
        _rwySurveyObstaclePhotoStatusText_();
        if (runwaySurveyToolState.ui.pendingObstacleSaveAfterPhoto) {
          runwaySurveyToolState.ui.pendingObstacleSaveAfterPhoto = false;
          addRunwaySurveyObstacleAngle();
        }
      });
    } finally {
      if (input) input.value = '';
    }
  }

  function openRunwaySurveyObstaclePopup(autoTriggered, cornerLabel, distanceM) {
    const gps = runwaySurveyToolState.gps || {};
    if (gps.tracking && !gps.paused) {
      runwaySurveyToolState.ui.pausedByPopup = true;
      pauseRunwaySurveyGps();
    } else {
      runwaySurveyToolState.ui.pausedByPopup = false;
    }
    if (!runwaySurveyToolState.prompts) runwaySurveyToolState.prompts = {
      a50Shown: false, a50Completed: false, a50Lat: null, a50Lon: null,
      a300Shown: false, a300Completed: false, a300Lat: null, a300Lon: null,
      c50Shown: false, c50Completed: false, c50Lat: null, c50Lon: null,
      c300Shown: false, c300Completed: false, c300Lat: null, c300Lon: null
    };
    const state = runwaySurveyToolState;
    const derived = _rwySurveyDerivedDimensions_();
    const runwayLengthM = Math.max(1, Math.round(Number(derived.lengthM || state.lengthM || 0) || 1));
    const startThr = _rwySurveyGetActiveStartThreshold_() || state.rwyIdent || 'A';
    const oppThr = _rwySurveyGetOppositeStartThreshold_() || _rwySurveyReciprocalIdent_(startThr);
    const manual = !(Number(distanceM || 0) > 0);
    let corner = String(cornerLabel || 'A').trim().toUpperCase();
    let dist = Math.max(1, Number(distanceM || 50));
    let thresholdRef = _rwySurveyThresholdForCorner_(corner) || corner;
    if (manual) {
      const fix = gps.current;
      const per = state.perimeter || {};
      const marks = Array.isArray(per.cornerMarks) ? per.cornerMarks : [];
      const origin = marks.length ? marks[0] : state.thresholdA;
      let heading = Number(state.headingDeg || 0);
      if (marks.length >= 2) {
        const a = marks[0], b = marks[1];
        const dy = Number(b.lat || 0) - Number(a.lat || 0);
        const dx = Number(b.lon || 0) - Number(a.lon || 0);
        heading = (Math.atan2(dx, dy) * 180 / Math.PI + 360) % 360;
      }
      let along = 0;
      if (fix && origin) {
        along = Math.max(0, Math.min(runwayLengthM, _rwySurveyProjectAlongAxisM_(fix.lat, fix.lon, origin.lat, origin.lon, heading)));
      }
      if (along > (runwayLengthM / 2)) {
        corner = 'C';
        thresholdRef = oppThr;
        dist = Math.round(runwayLengthM - along);
      } else {
        corner = 'A';
        thresholdRef = startThr;
        dist = Math.round(along);
      }
      dist = Math.max(1, dist || 1);
    }
    thresholdRef = String(thresholdRef || startThr).trim().toUpperCase();
    const mode = dist >= 300 ? 'takeoff' : 'landing';
    runwaySurveyToolState.ui.activeObstaclePrompt = {
      corner: corner,
      distanceM: dist,
      fromThreshold: thresholdRef,
      operation: mode,
      manual: manual
    };
    runwaySurveyToolState.ui.obstaclePhoto = null;
    runwaySurveyToolState.ui.pendingObstacleSaveAfterPhoto = false;
    const titleEl = document.getElementById('rwysurvey-obstacle-title');
    if (titleEl) titleEl.textContent = 'Obstacle Angles @ ' + dist + 'm from THR ' + thresholdRef;
    const context = document.getElementById('rwysurvey-obstacle-context');
    if (context) {
      const approachHint = mode === 'landing'
        ? 'Measure behind you for landing clearance on THR ' + thresholdRef + '.'
        : 'Measure ahead for takeoff climb from THR ' + thresholdRef + '.';
      if (autoTriggered) {
        context.textContent = 'Paused at ~' + dist + 'm from THR ' + thresholdRef + '. ' + approachHint;
      } else {
        context.textContent = 'Capture obstacle angle(s) at THR ' + thresholdRef + ' +' + dist + 'm. ' + approachHint;
      }
    }
    const inclBtn = document.getElementById('rwysurvey-open-inclinometer-btn');
    if (inclBtn) inclBtn.textContent = 'INCLINÔMETRO ' + dist + 'M';
    const thresholdSel = document.getElementById('rwysurvey-obs-threshold-popup');
    if (thresholdSel) {
      const opts = '<option value="' + startThr + '">' + startThr + '</option><option value="' + oppThr + '">' + oppThr + '</option>';
      thresholdSel.innerHTML = opts;
      thresholdSel.value = thresholdRef === oppThr ? oppThr : startThr;
      thresholdSel.disabled = !!autoTriggered;
    }
    const distInput = document.getElementById('rwysurvey-obs-distance-popup');
    if (distInput) {
      distInput.value = String(Math.max(1, Math.round(dist)));
      distInput.readOnly = !!autoTriggered;
    }
    const opSel = document.getElementById('rwysurvey-obs-operation-popup');
    if (opSel) {
      opSel.value = autoTriggered ? mode : (manual ? 'auto' : mode);
    }
    if (document.getElementById('rwysurvey-obs-type-popup')) {
      const current = rwySurveyGetObstacleType_();
      if (!current) rwySurveySetObstacleType_('rock');
    }
    _rwySurveyObstaclePhotoStatusText_();
    const popup = document.getElementById('rwysurvey-obstacle-popup');
    if (popup) popup.style.display = 'block';
    renderRunwaySurveyA50Status();
  }

  function closeRunwaySurveyObstaclePopup(silent) {
    closeRunwaySurveyInclinometer();
    const popup = document.getElementById('rwysurvey-obstacle-popup');
    if (popup) popup.style.display = 'none';
    const prompts = runwaySurveyToolState.prompts || {};
    const active = runwaySurveyToolState.ui && runwaySurveyToolState.ui.activeObstaclePrompt ? runwaySurveyToolState.ui.activeObstaclePrompt : { corner: 'A', distanceM: 50 };
    if (active.corner === 'A' && active.distanceM >= 300 && prompts.a300Shown && !prompts.a300Completed) prompts.a300Completed = true;
    if (active.corner === 'A' && active.distanceM < 300 && prompts.a50Shown && !prompts.a50Completed) prompts.a50Completed = true;
    if (active.corner === 'C' && active.distanceM >= 300 && prompts.c300Shown && !prompts.c300Completed) prompts.c300Completed = true;
    if (active.corner === 'C' && active.distanceM < 300 && prompts.c50Shown && !prompts.c50Completed) prompts.c50Completed = true;
    const gps = runwaySurveyToolState.gps || {};
    if (runwaySurveyToolState.ui && runwaySurveyToolState.ui.pausedByPopup && gps.tracking && gps.paused) {
      runwaySurveyToolState.ui.pausedByPopup = false;
      resumeRunwaySurveyGps();
    }
    if (!silent && window.M) M.toast({ html: 'Continue to next corner', classes: 'blue darken-2' });
    renderRunwaySurveyA50Status();
    renderRunwaySurveyActionButtons();
  }

  function renderRunwaySurveyActionButtons() {
    const state = runwaySurveyToolState;
    const gps = state.gps || {};
    const per = state.perimeter || {};
    const capture = state.capture || _rwySurveyCaptureDefaults_();
    const hasRunway = !!String(state.rwyIdent || '').trim();
    const cornerMarks = Array.isArray(per.cornerMarks) ? per.cornerMarks : [];
    const corners = _rwySurveyCornerSequence_();
    const nextCornerIdx = Math.min(cornerMarks.length, corners.length - 1);
    const nextCorner = corners[nextCornerIdx];
    const pending = !!capture.pending;
    const canMark = !!(hasRunway && gps.tracking && !gps.paused && !pending);

    const gpsToggle = document.getElementById('rwysurvey-gps-toggle');
    if (gpsToggle) {
      let label = 'START GPS';
      if (pending) label = 'HOLD STILL...';
      else if (gps.tracking && gps.paused) label = 'RESUME GPS';
      else if (gps.tracking) label = 'STOP GPS';
      gpsToggle.textContent = label;
      gpsToggle.disabled = pending;
    }

    const reset = document.getElementById('rwysurvey-reset');
    if (reset) reset.disabled = pending;

    const cornerA = document.getElementById('rwysurvey-corner-a');
    const cornerB = document.getElementById('rwysurvey-corner-b');
    const cornerC = document.getElementById('rwysurvey-corner-c');
    const cornerD = document.getElementById('rwysurvey-corner-d');
    const cornerE = document.getElementById('rwysurvey-corner-e');
    if (cornerA) cornerA.disabled = !canMark || per.closed || nextCorner !== 'A';
    if (cornerB) cornerB.disabled = !canMark || per.closed || nextCorner !== 'B';
    if (cornerC) cornerC.disabled = !canMark || per.closed || nextCorner !== 'C';
    if (cornerD) cornerD.disabled = !canMark || per.closed || nextCorner !== 'D';
    if (cornerE) cornerE.disabled = !canMark || per.closed || nextCorner !== 'E';

    if (cornerA) cornerA.textContent = cornerMarks.length > 0 ? 'A ✓' : 'MARK A';
    if (cornerB) cornerB.textContent = cornerMarks.length > 1 ? 'B ✓' : 'MARK B';
    if (cornerC) cornerC.textContent = cornerMarks.length > 2 ? 'C ✓' : 'MARK C';
    if (cornerD) cornerD.textContent = cornerMarks.length > 3 ? 'D ✓' : 'MARK D';
    if (cornerE) cornerE.textContent = cornerMarks.length > 4 ? 'E ✓' : 'MARK E';

    const narrow = document.getElementById('rwysurvey-width-narrow');
    if (narrow) narrow.disabled = !canMark;
    const wide = document.getElementById('rwysurvey-width-wide');
    if (wide) wide.disabled = !canMark;
    const slopeBtn = document.getElementById('rwysurvey-slope-toggle');
    if (slopeBtn) slopeBtn.disabled = pending;
    const photoBtn = document.getElementById('rwysurvey-photo-add-btn');
    if (photoBtn) photoBtn.disabled = pending || !runwaySurveyToolState.icao;
    const measureBtn = document.getElementById('rwysurvey-measure-toggle');
    if (measureBtn) measureBtn.disabled = pending || !gps.tracking || gps.paused;
    renderRunwaySurveyMeasureTool();
  }

  function _rwySurveyPhotoSafeName_(name) {
    return String(name || 'runway_photo.jpg').replace(/[^a-zA-Z0-9._-]+/g, '_');
  }

  function openRunwaySurveyPhotoOptions() {
    const state = runwaySurveyToolState;
    if (!state.icao) {
      if (window.M) M.toast({ html: 'Select airport first', classes: 'orange' });
      return;
    }
    document.getElementById('rwysurvey-photo-sheet').style.display = 'block';
  }

  function closeRunwaySurveyPhotoSheet() {
    document.getElementById('rwysurvey-photo-sheet').style.display = 'none';
  }

  function _rwsShtPick_(choice) {
    closeRunwaySurveyPhotoSheet();
    if (choice === 'camera') {
      const camInput = document.getElementById('rwysurvey-photo-capture-input');
      if (camInput) { camInput.value = ''; camInput.click(); }
    } else if (choice === 'library') {
      const libInput = document.getElementById('rwysurvey-photo-library-input');
      if (libInput) { libInput.value = ''; libInput.click(); }
    } else if (choice === 'drive') {
      openRunwaySurveyPhotoDriveFolder();
    }
  }

  function openRunwaySurveyPhotoDriveFolder() {
    const state = runwaySurveyToolState;
    if (!state.icao) {
      if (window.M) M.toast({ html: 'Select airport first', classes: 'orange' });
      return;
    }
    window.runOrQueueServerAction({
      method: 'getAirportPhotoFolderLink',
      args: [String(state.icao || '').trim().toUpperCase()],
      label: 'Airport photo folder ' + String(state.icao || '')
    }, {
      onSuccess: function(resp) {
        if (!resp || !resp.success || !resp.url) {
          if (window.M) M.toast({ html: (resp && resp.error) ? resp.error : 'Airport folder not found', classes: 'orange' });
          return;
        }
        try {
          window.open(String(resp.url), '_blank', 'noopener');
        } catch (e) {}
      },
      onQueued: function() {
        if (window.M) M.toast({ html: 'Offline: cannot open Drive folder right now', classes: 'orange' });
      },
      onFailure: function(err) {
        if (window.M) M.toast({ html: 'Folder open failed: ' + (err && err.message ? err.message : String(err)), classes: 'red' });
      }
    });
  }

  function onRunwaySurveyPhotoInputChange(input, source) {
    try {
      const files = input && input.files ? Array.from(input.files) : [];
      if (!files.length) return;
      _rwySurveyUploadSelectedPhotos_(files, String(source || 'library'));
    } finally {
      if (input) input.value = '';
    }
  }

  function _rwySurveyReadFileAsDataUrl_(file) {
    return new Promise(function(resolve, reject) {
      const reader = new FileReader();
      reader.onload = function() { resolve(String(reader.result || '')); };
      reader.onerror = function() { reject(new Error('Unable to read image file')); };
      reader.readAsDataURL(file);
    });
  }

  function _rwySurveyShrinkImageDataUrl_(dataUrl) {
    return new Promise(function(resolve) {
      const img = new Image();
      img.onload = function() {
        try {
          const maxW = 1600;
          const maxH = 1600;
          const ratio = Math.min(maxW / Math.max(1, img.width), maxH / Math.max(1, img.height), 1);
          const w = Math.max(1, Math.round(img.width * ratio));
          const h = Math.max(1, Math.round(img.height * ratio));
          const canvas = document.createElement('canvas');
          canvas.width = w;
          canvas.height = h;
          const ctx = canvas.getContext('2d');
          if (!ctx) return resolve(dataUrl);
          ctx.drawImage(img, 0, 0, w, h);
          resolve(canvas.toDataURL('image/jpeg', 0.78));
        } catch (e) {
          resolve(dataUrl);
        }
      };
      img.onerror = function() { resolve(dataUrl); };
      img.src = dataUrl;
    });
  }

  async function _rwySurveyUploadSelectedPhotos_(files, source, options) {
    const state = runwaySurveyToolState;
    const icao = String(state.icao || '').trim().toUpperCase();
    const rwy = String(state.rwyIdent || '').trim().toUpperCase();
    const opts = options && typeof options === 'object' ? options : {};
    const createdRefs = [];
    if (!icao) {
      if (window.M) M.toast({ html: 'Select airport first', classes: 'orange' });
      return createdRefs;
    }

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      if (!file) continue;
      try {
        const rawDataUrl = await _rwySurveyReadFileAsDataUrl_(file);
        const compressedDataUrl = await _rwySurveyShrinkImageDataUrl_(rawDataUrl);
        const parts = compressedDataUrl.split(',');
        if (parts.length < 2) throw new Error('Invalid image data');
        const header = parts[0] || '';
        const mimeMatch = header.match(/data:([^;]+);base64/i);
        const mimeType = (mimeMatch && mimeMatch[1]) ? mimeMatch[1] : 'image/jpeg';
        const base64Data = parts[1];
        const ts = new Date();
        const stamp = ts.toISOString().replace(/[.:]/g, '-');
        const originalName = _rwySurveyPhotoSafeName_(file.name || ('photo_' + stamp + '.jpg'));
        const rwyTag = rwy ? rwy.replace(/\//g, '-') + '_' : '';
        const fileName = String((i === 0 && opts.forcedFileName) ? opts.forcedFileName : (icao + '_' + rwyTag + stamp + '_' + originalName));

        const localRef = {
          name: fileName,
          source: source,
          status: 'uploading',
          queuedAt: ts.toISOString(),
          sizeKb: Math.round((base64Data.length * 3 / 4) / 1024)
        };
        state.photos.push(localRef);
        createdRefs.push(localRef);
        renderRunwaySurveyPhotoList();

        window.runOrQueueServerAction({
          method: 'uploadRunwaySurveyPhoto',
          args: [{
            icao: icao,
            rwyIdent: rwy,
            fileName: fileName,
            mimeType: mimeType,
            base64Data: base64Data,
            source: source,
            takenAt: ts.toISOString()
          }],
          label: 'Runway photo ' + icao + ' ' + rwy
        }, {
          onSuccess: function(resp) {
            localRef.status = 'uploaded';
            localRef.fileId = resp && resp.fileId ? String(resp.fileId) : '';
            localRef.url = resp && resp.url ? String(resp.url) : '';
            localRef.folderUrl = resp && resp.folderUrl ? String(resp.folderUrl) : '';
            renderRunwaySurveyPhotoList();
            if (window.M) M.toast({ html: 'Photo uploaded', classes: 'green' });
          },
          onQueued: function() {
            localRef.status = 'queued';
            renderRunwaySurveyPhotoList();
            if (window.M) M.toast({ html: 'Offline: photo queued for sync', classes: 'orange' });
          },
          onFailure: function(err) {
            localRef.status = 'failed';
            localRef.error = err && err.message ? String(err.message) : String(err || 'Upload failed');
            renderRunwaySurveyPhotoList();
            if (window.M) M.toast({ html: 'Photo upload failed', classes: 'red' });
          }
        });
      } catch (e) {
        if (window.M) M.toast({ html: 'Photo skipped: ' + (e && e.message ? e.message : 'read error'), classes: 'orange' });
      }
    }
    return createdRefs;
  }

  function renderRunwaySurveyPhotoList() {
    const el = document.getElementById('rwysurvey-photo-list');
    const summary = document.getElementById('rwysurvey-photo-summary');
    const items = Array.isArray(runwaySurveyToolState.photos) ? runwaySurveyToolState.photos : [];
    if (summary) summary.textContent = 'Photos: ' + items.length;
    if (!el) return;
    if (!items.length) {
      el.innerHTML = '<div style="font-size:0.8rem; color:#999; margin-top:4px;">No photos attached yet.</div>';
      return;
    }
    const badgeColor = function(status) {
      if (status === 'uploaded') return '#2e7d32';
      if (status === 'queued') return '#ef6c00';
      if (status === 'failed') return '#c62828';
      return '#455a64';
    };
    el.innerHTML = items.map(function(p, idx) {
      const status = String(p && p.status || 'pending').toLowerCase();
      const url = String(p && p.url || '');
      const openLink = url ? '<a href="' + url + '" target="_blank" rel="noopener" style="margin-left:8px; color:#1565c0; text-decoration:underline;">open</a>' : '';
      return '<div style="display:flex; justify-content:space-between; align-items:center; gap:8px; margin-top:4px; padding:5px 8px; border:1px solid #dbe7f3; border-radius:6px; background:#f7fbff; font-size:0.8rem;">'
        + '<span><b>Photo ' + (idx + 1) + '</b> · ' + Math.max(0, Number(p && p.sizeKb || 0)) + 'KB · <span style="color:' + badgeColor(status) + '; font-weight:800;">' + status.toUpperCase() + '</span>' + openLink + '</span>'
        + '<button onclick="removeRunwaySurveyPhoto(' + idx + ')" style="border:none; background:none; color:#d32f2f; cursor:pointer;">✕</button></div>';
    }).join('');
  }

  function removeRunwaySurveyPhoto(i) {
    const arr = runwaySurveyToolState.photos;
    if (!Array.isArray(arr)) return;
    if (i < 0 || i >= arr.length) return;
    arr.splice(i, 1);
    renderRunwaySurveyPhotoList();
  }

  function copyRunwaySurveyDebugLog() {
    const state = runwaySurveyToolState;
    const payload = {
      icao: state.icao,
      rwyIdent: state.rwyIdent,
      gpsSummary: {
        points: (state.gps && state.gps.points || []).length,
        currentAccM: Number(state.gps && state.gps.current && state.gps.current.acc || 0),
        bestAccM: isFinite(state.gps && state.gps.bestAcc) ? Number(state.gps.bestAcc) : null,
        avgAccM: state.gps && state.gps.samples ? Number(state.gps.avgAcc) : null
      },
      perimeter: state.perimeter,
      segments: (state.perimeter && state.perimeter.segments) || [],
      events: state.debugEvents || []
    };
    const text = JSON.stringify(payload, null, 2);
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function() {
        if (window.M) M.toast({ html: 'Debug log copied', classes: 'green' });
      }).catch(function() {
        window.prompt('Copy debug log:', text);
      });
    } else {
      window.prompt('Copy debug log:', text);
    }
  }

  function _rwySurveyListAirports_() {
    const rows = (window.appData && Array.isArray(window.appData.airports)) ? window.appData.airports : [];
    const seen = {};
    const list = [];
    rows.forEach(function(r) {
      const icao = String(r && (r.airportICAO || r.icao || r.ICAO) || '').trim().toUpperCase();
      if (!icao || seen[icao]) return;
      seen[icao] = true;
      list.push({ icao: icao, name: String((r && (r.airportName || r.name || r.nome)) || icao) });
    });
    list.sort(function(a, b) { return a.icao.localeCompare(b.icao); });
    return list;
  }

  function _rwySurveyRenderAirportOptions_(filterText) {
    const airports = _rwySurveyListAirports_();
    const q = String(filterText || '').trim().toUpperCase();
    const filtered = !q ? airports : airports.filter(function(a) {
      const txt = (String(a.icao || '') + ' ' + String(a.name || '')).toUpperCase();
      return txt.indexOf(q) >= 0;
    });
    return filtered;
  }

  function _rwySurveyDefaultSurfaceOptions_() {
    return ['Firm Turf', 'Short Grass', 'Grass to 6"', 'Long Grass', 'Rough', 'Mud', 'Sand', 'Asphalt'];
  }

  function _rwySurveyRenderSurfaceOptions_() {
    const sel = document.getElementById('rwysurvey-surface');
    if (!sel) return;
    const options = Array.isArray(runwaySurveyToolState.surfaceOptions) && runwaySurveyToolState.surfaceOptions.length
      ? runwaySurveyToolState.surfaceOptions
      : _rwySurveyDefaultSurfaceOptions_();
    const current = String(sel.value || '').trim();
    sel.innerHTML = '<option value="">-- Select --</option>' + options.map(function(opt) {
      const text = String(opt || '').trim();
      return '<option value="' + text.replace(/"/g, '&quot;') + '">' + text + '</option>';
    }).join('');
    if (current && options.indexOf(current) >= 0) sel.value = current;
  }

  function _rwySurveyLoadSurfaceOptions_() {
    _rwySurveyRenderSurfaceOptions_();
    if (!window.google || !google.script || !google.script.run) return;
    google.script.run
      .withSuccessHandler(function(resp) {
        const options = resp && Array.isArray(resp.options) && resp.options.length ? resp.options : _rwySurveyDefaultSurfaceOptions_();
        runwaySurveyToolState.surfaceOptions = options;
        _rwySurveyRenderSurfaceOptions_();
      })
      .withFailureHandler(function() {
        runwaySurveyToolState.surfaceOptions = _rwySurveyDefaultSurfaceOptions_();
        _rwySurveyRenderSurfaceOptions_();
      })
      .getRunwaySurveySurfaceOptions();
  }

  function _rwySurveyMeasuredDistanceFromIndex_(startIdx) {
    const pts = Array.isArray(runwaySurveyToolState.gps && runwaySurveyToolState.gps.points) ? runwaySurveyToolState.gps.points : [];
    const from = Math.max(0, Number(startIdx || 0));
    if (pts.length - from < 2) return 0;
    let total = 0;
    for (let i = from + 1; i < pts.length; i++) {
      const prev = pts[i - 1];
      const next = pts[i];
      const dt = Math.max(0.2, (Number(next && next.ts || 0) - Number(prev && prev.ts || 0)) / 1000);
      const dist = _rwySurveyDistanceMetersBetween_(prev, next);
      const speed = dist / dt;
      if (dist > 18 && speed > 8) continue;
      total += dist;
    }
    return Math.round(total);
  }

  function renderRunwaySurveyMeasureTool() {
    const tool = runwaySurveyToolState.measureTool || _rwySurveyMeasureDefaults_();
    const btn = document.getElementById('rwysurvey-measure-toggle');
    const el = document.getElementById('rwysurvey-measure-readout');
    if (btn) {
      btn.textContent = tool.active ? 'END MEASURING' : 'START MEASURING';
      btn.className = tool.active ? 'btn deep-orange darken-2' : 'btn white blue-text text-darken-2';
      btn.style.border = tool.active ? 'none' : '2px solid #1976d2';
      btn.style.boxShadow = 'none';
    }
    if (!el) return;
    if (tool.active) {
      const liveDist = _rwySurveyMeasuredDistanceFromIndex_(tool.startIdx);
      el.textContent = 'Standalone measuring in progress: ' + liveDist + ' m (not submitted)';
      return;
    }
    if (Number(tool.distanceM || 0) > 0) {
      el.textContent = 'Standalone measured distance: ' + Math.round(Number(tool.distanceM || 0)) + ' m (not submitted)';
      return;
    }
    el.textContent = 'Standalone measurer idle (not submitted).';
  }

  function toggleRunwaySurveyMeasureTool() {
    const gps = runwaySurveyToolState.gps || {};
    if (!gps.tracking || gps.paused) {
      if (window.M) M.toast({ html: 'Start GPS tracking before measuring', classes: 'orange' });
      return;
    }
    if (!runwaySurveyToolState.measureTool) runwaySurveyToolState.measureTool = _rwySurveyMeasureDefaults_();
    const tool = runwaySurveyToolState.measureTool;
    if (!tool.active) {
      tool.active = true;
      tool.startIdx = Math.max(0, (gps.points || []).length - 1);
      tool.distanceM = 0;
      renderRunwaySurveyMeasureTool();
      if (window.M) M.toast({ html: 'Measuring started', classes: 'blue darken-2' });
      return;
    }
    tool.active = false;
    tool.distanceM = _rwySurveyMeasuredDistanceFromIndex_(tool.startIdx);
    renderRunwaySurveyMeasureTool();
    if (window.M) M.toast({ html: 'Measured ' + Math.round(tool.distanceM || 0) + ' m', classes: 'green' });
  }

  function _rwySurveyShowAirportOptions_() {
    _rwySurveyFilterAirports_();
  }

  function _rwySurveySelectAirportFromSearch_(icao) {
    const input = document.getElementById('rwysurvey-icao-search');
    if (input) input.value = String(icao || '');
    const dropdown = document.getElementById('rwysurvey-airport-dropdown');
    if (dropdown) dropdown.style.display = 'none';
    runwaySurveyToolState.icao = String(icao || '').trim().toUpperCase();
    _rwySurveyOnAirportChange_();
  }

  function _rwySurveyGetAirportName_(icao) {
    const rows = (window.appData && Array.isArray(window.appData.airports)) ? window.appData.airports : [];
    const target = String(icao || '').trim().toUpperCase();
    let bestName = '';
    for (var i = 0; i < rows.length; i++) {
      const r = rows[i];
      const thisIcao = String(r && (r.airportICAO || r.icao || r.ICAO) || '').trim().toUpperCase();
      if (thisIcao !== target) continue;
      const candidate = String((r && (r.airportName || r.name || r.nome || r.NOME || r.airport || r.AIRPORT_NAME)) || '').trim();
      if (candidate) {
        bestName = candidate;
        break;
      }
    }
    return bestName;
  }

  function _rwySurveyFilterAirports_() {
    const input = document.getElementById('rwysurvey-icao-search');
    const searchText = (input && input.value) ? String(input.value).trim().toUpperCase() : '';
    const dropdown = document.getElementById('rwysurvey-airport-dropdown');
    if (!dropdown) return;
    const matches = _rwySurveyRenderAirportOptions_(searchText);
    if (!searchText) {
      dropdown.style.display = 'none';
      return;
    }
    dropdown.innerHTML = '';
    matches.forEach(function(item) {
      const icao = String(item && item.icao || '').trim().toUpperCase();
      const name = String(item && item.name || _rwySurveyGetAirportName_(icao) || '').trim();
      if (!icao) return;
      const div = document.createElement('div');
      div.style.cssText = 'padding:8px 12px; cursor:pointer; border-bottom:1px solid #e0e0e0; font-size:0.9rem;';
      div.textContent = icao + (name ? (' · ' + name) : '');
      div.onmouseover = function() { div.style.background = '#f5f5f5'; };
      div.onmouseout = function() { div.style.background = 'transparent'; };
      div.onclick = function() { _rwySurveySelectAirportFromSearch_(icao); };
      dropdown.appendChild(div);
    });
    if (matches.length) {
      dropdown.style.display = 'block';
    } else {
      dropdown.innerHTML = '<div style="padding:8px 12px; color:#999; font-size:0.85rem;">No airports found</div>';
      dropdown.style.display = 'block';
    }
  }

  function _rwySurveyApplyAirportSelection_() {
    const input = document.getElementById('rwysurvey-icao-search');
    const raw = String((input && input.value) || '').trim().toUpperCase();
    const icao = raw.split(/\s+/)[0] || '';
    if (!icao) {
      if (window.M) M.toast({ html: 'No airport selected', classes: 'orange' });
      return;
    }
    runwaySurveyToolState.icao = icao;
    _rwySurveyOnAirportChange_();
  }

  function _rwySurveyRunwaysForIcao_(icao) {
    const rows = (window.appData && Array.isArray(window.appData.airports)) ? window.appData.airports : [];
    const target = String(icao || '').trim().toUpperCase();
    return rows.filter(function(r) {
      return String(r && (r.airportICAO || r.icao || r.ICAO) || '').trim().toUpperCase() === target;
    });
  }

  function _rwySurveyToNumber_(v) {
    const n = Number(v);
    return isFinite(n) ? n : 0;
  }

  function _rwySurveyHeadingFromIdent_(ident) {
    const num = parseInt(String(ident || '').replace(/\D+/g, ''), 10);
    if (isNaN(num) || num < 1 || num > 36) return 0;
    return num * 10;
  }

  function _rwySurveyReciprocalIdent_(ident) {
    const num = parseInt(String(ident || '').replace(/\D+/g, ''), 10);
    if (isNaN(num) || num < 1 || num > 36) return 'REV';
    return String((((num + 18 - 1) % 36) + 1)).padStart(2, '0');
  }

  function _rwySurveyRunwayPairKey_(ident) {
    const raw = String(ident || '').trim().toUpperCase();
    if (!raw) return '';

    if (raw.indexOf('/') >= 0) {
      const parts = raw.split('/').map(function(p) { return String(p || '').trim().toUpperCase(); }).filter(Boolean);
      if (!parts.length) return raw;
      parts.sort(function(a, b) { return a.localeCompare(b); });
      return parts.join('/');
    }

    const m = raw.match(/(\d{1,2})([LCR])?/);
    if (!m) return raw;
    const num = parseInt(m[1], 10);
    if (!(num >= 1 && num <= 36)) return raw;
    const suffix = m[2] || '';
    const recipNum = ((num + 18 - 1) % 36) + 1;
    const recipSuffix = suffix === 'L' ? 'R' : (suffix === 'R' ? 'L' : suffix);
    const a = String(num).padStart(2, '0') + suffix;
    const b = String(recipNum).padStart(2, '0') + recipSuffix;
    return [a, b].sort(function(x, y) { return x.localeCompare(y); }).join('/');
  }

  function _rwySurveyBuildRunwayChoices_(runways) {
    const rows = Array.isArray(runways) ? runways : [];
    const byPair = {};
    const ordered = [];

    rows.forEach(function(r, i) {
      const identRaw = String(r && (r.runwayIdent || r.rwyIdent || '') || '').trim().toUpperCase();
      const ident = identRaw || ('RWY-' + (i + 1));
      const pairKey = _rwySurveyRunwayPairKey_(ident) || ident;
      if (!byPair[pairKey]) {
        byPair[pairKey] = {
          pairKey: pairKey,
          value: ident,
          label: pairKey,
          rows: [r],
          primaryRow: r
        };
        ordered.push(byPair[pairKey]);
      } else {
        byPair[pairKey].rows.push(r);
        if (ident.localeCompare(byPair[pairKey].value) < 0) {
          byPair[pairKey].value = ident;
          byPair[pairKey].primaryRow = r;
        }
      }
    });

    ordered.sort(function(a, b) { return String(a.label || '').localeCompare(String(b.label || '')); });
    return ordered;
  }

  function openRunwaySurveyTool() {
    const modal = document.getElementById('rwysurvey-modal');
    if (!modal) return;
    modal.style.display = 'block';

    const airports = _rwySurveyRenderAirportOptions_('');
    const search = document.getElementById('rwysurvey-icao-search');

    const remembered = String(runwaySurveyToolState.icao || '').trim().toUpperCase();
    const firstIcao = airports.length ? airports[0].icao : '';
    const selectedIcao = airports.some(function(a) { return a.icao === remembered; }) ? remembered : firstIcao;
    runwaySurveyToolState.icao = selectedIcao;
    if (search) search.value = selectedIcao;
    _rwySurveyLoadSurfaceOptions_();
    _rwySurveyOnAirportChange_(selectedIcao);
    renderRunwaySurveyA50Status();
    renderRunwaySurveyFeatureList();
    renderRunwaySurveyPhotoList();
    renderRunwaySurveyObstacleList();
    renderRunwaySurveySlopeCaptureUi();
    renderRunwaySurveyMeasureTool();
    renderRunwaySurveyActionButtons();
  }

  function closeRunwaySurveyTool() {
    closeRunwayDiagramPreview();
    closeRunwaySurveyFeaturePopup(true);
    closeRunwaySurveyObstaclePopup(true);
    closeRunwaySurveyInclinometer();
    _rwySurveyCancelPendingCapture_('Capture cancelled', true);
    stopRunwaySurveyGps(true);
    closeRunwaySurveyInfo();
    const modal = document.getElementById('rwysurvey-modal');
    if (modal) modal.style.display = 'none';
  }

  function openRunwaySurveyInfo() {
    const modal = document.getElementById('rwysurvey-info-modal');
    if (modal) modal.style.display = 'block';
  }

  function closeRunwaySurveyInfo() {
    const modal = document.getElementById('rwysurvey-info-modal');
    if (modal) modal.style.display = 'none';
  }

  function _rwySurveyOnAirportChange_(icaoOverride) {
    const input = document.getElementById('rwysurvey-icao-search');
    const dropdown = document.getElementById('rwysurvey-airport-dropdown');
    const fromInput = String((input && input.value) || '').trim().toUpperCase().split(/\s+/)[0] || '';
    const chosenIcao = String(icaoOverride || runwaySurveyToolState.icao || fromInput || '').trim().toUpperCase();
    runwaySurveyToolState.icao = chosenIcao;
    if (dropdown) dropdown.style.display = 'none';
    if (input && chosenIcao) {
      input.value = chosenIcao;
    }
    const runways = _rwySurveyRunwaysForIcao_(runwaySurveyToolState.icao);
    const runwayChoices = _rwySurveyBuildRunwayChoices_(runways);
    runwaySurveyToolState.runwayChoices = runwayChoices;
    const rwySel = document.getElementById('rwysurvey-rwy');
    if (rwySel) {
      rwySel.innerHTML = '';
      const ph = document.createElement('option');
      ph.value = '';
      ph.textContent = '-- Select --';
      rwySel.appendChild(ph);
      runwayChoices.forEach(function(choice) {
        const opt = document.createElement('option');
        opt.value = String(choice.value || '').trim().toUpperCase();
        opt.textContent = String(choice.label || opt.value);
        rwySel.appendChild(opt);
      });
      rwySel.disabled = runwayChoices.length === 0;
      rwySel.value = '';
    }

    if (!runwayChoices.length && window.M) {
      M.toast({ html: 'No runway found for this airport', classes: 'orange' });
    }
    _rwySurveyOnRunwayChange_();
  }

  function _rwySurveyOnRunwayChange_() {
    const rwySel = document.getElementById('rwysurvey-rwy');
    const ident = String((rwySel && rwySel.value) || '').trim().toUpperCase();
    if (!ident) {
      runwaySurveyToolState.rwyIdent = '';
      runwaySurveyToolState.startThresholdIdent = '';
      runwaySurveyToolState.reciprocalThresholdIdent = '';
      runwaySurveyToolState.runway = null;
      runwaySurveyToolState.headingDeg = 0;
      runwaySurveyToolState.lengthM = 0;
      runwaySurveyToolState.official = { lengthM: 0, widthM: 0, surface: '', headingDeg: 0 };
      runwaySurveyToolState.internalUpdatedAt = '';
      runwaySurveyToolState.photos = [];
      runwaySurveyToolState.perimeter = _rwySurveyPerimeterDefaults_();
      runwaySurveyToolState.capture = _rwySurveyCaptureDefaults_();
      runwaySurveyToolState.thresholdAlerts = { thrA: { m300: false, m50: false }, thrB: { m300: false, m50: false } };

      const officialElEmpty = document.getElementById('rwysurvey-official');
      if (officialElEmpty) officialElEmpty.innerHTML = '<b>Official runway reference:</b> select a runway';

      renderRunwaySurveyPerimeterTally();
      renderRunwaySurveyMarkHistory();
      renderRunwaySurveyStatus();
      renderRunwaySurveyA50Status();
      renderRunwaySurveyPhotoList();
      renderRunwaySurveySlopeCaptureUi();
      renderRunwaySurveyActionButtons();
      return;
    }

    const choices = Array.isArray(runwaySurveyToolState.runwayChoices) ? runwaySurveyToolState.runwayChoices : [];
    const chosen = choices.find(function(c) {
      return String(c && c.value || '').trim().toUpperCase() === ident;
    }) || choices[0] || null;

    runwaySurveyToolState.rwyIdent = String((chosen && chosen.value) || ident || '').trim().toUpperCase();
    runwaySurveyToolState.startThresholdIdent = '';
    runwaySurveyToolState.reciprocalThresholdIdent = '';

    const runways = _rwySurveyRunwaysForIcao_(runwaySurveyToolState.icao);
    const row = (chosen && chosen.primaryRow) || runways.find(function(r) {
      return String(r.runwayIdent || r.rwyIdent || '').trim().toUpperCase() === runwaySurveyToolState.rwyIdent;
    }) || runways[0] || null;

    runwaySurveyToolState.runway = row;
    runwaySurveyToolState.headingDeg = _rwySurveyToNumber_(row && (row.runwayHeading || row.headingDeg)) || _rwySurveyHeadingFromIdent_(ident);
    runwaySurveyToolState.lengthM = _rwySurveyToNumber_(row && (row.runwayLength || row.length));

    let known = {};
    try {
      const raw = row && row.knownFeatures;
      known = raw ? (typeof raw === 'string' ? JSON.parse(raw) : raw) : {};
    } catch (e) {
      known = {};
    }
    if (Array.isArray(known)) known = { features: known };
    const officialRef = known && known.officialReference ? known.officialReference : {};
    const verifiedOperational = known && known.verifiedOperational && typeof known.verifiedOperational === 'object' ? known.verifiedOperational : {};
    const verifiedSurvey = known && known.verifiedSurvey && typeof known.verifiedSurvey === 'object' ? known.verifiedSurvey : {};
    const currentVersion = known && known.currentSurveyVersion && typeof known.currentSurveyVersion === 'object' ? known.currentSurveyVersion : {};
    runwaySurveyToolState.internalUpdatedAt = String(currentVersion.publishedAt || verifiedSurvey.capturedAt || known.updatedAt || '').trim();
    runwaySurveyToolState.official = {
      lengthM: _rwySurveyToNumber_(officialRef.lengthM || runwaySurveyToolState.lengthM),
      widthM: _rwySurveyToNumber_(officialRef.widthM || row && (row.runwayWidth || row.width)),
      surface: String(officialRef.surface || row && (row.runwaySurfaceActual || row.surface) || '').trim(),
      headingDeg: _rwySurveyToNumber_(officialRef.headingDeg || runwaySurveyToolState.headingDeg)
    };

    runwaySurveyToolState.features = [];
    runwaySurveyToolState.obstacleAngles50m = [];
    runwaySurveyToolState.slopeSegments = [];
    runwaySurveyToolState.surfaceObserved = '';
    runwaySurveyToolState.cutdownAreas = { thrA: null, thrB: null };
    runwaySurveyToolState.perimeter = _rwySurveyPerimeterDefaults_();
    runwaySurveyToolState.capture = _rwySurveyCaptureDefaults_();
    runwaySurveyToolState.ui = _rwySurveyUiDefaults_();
    runwaySurveyToolState.prompts = {
      a50Shown: false, a50Completed: false, a50Lat: null, a50Lon: null,
      a300Shown: false, a300Completed: false, a300Lat: null, a300Lon: null,
      c50Shown: false, c50Completed: false, c50Lat: null, c50Lon: null,
      c300Shown: false, c300Completed: false, c300Lat: null, c300Lon: null
    };
    runwaySurveyToolState.slopeCapture = { active: false, startIdx: -1, startTs: null, pendingStartDistanceM: 0, pendingDistanceM: 0, pendingFromThreshold: '' };
    runwaySurveyToolState.photos = [];
    runwaySurveyToolState.measureTool = _rwySurveyMeasureDefaults_();
    runwaySurveyToolState.widthObservations = [];
    runwaySurveyToolState.debugEvents = [];
    runwaySurveyToolState.thresholdAlerts = { thrA: { m300: false, m50: false }, thrB: { m300: false, m50: false } };
    const srfSel = document.getElementById('rwysurvey-surface');
    if (srfSel) srfSel.value = '';
    const cutdownEl = document.getElementById('rwysurvey-cutdown');
    if (cutdownEl) {
      const savedCutdown = Number(verifiedOperational.cutdownAreaM || known.cutdownAreaM || 0);
      cutdownEl.value = (isFinite(savedCutdown) && savedCutdown >= 0) ? String(Math.round(savedCutdown)) : '';
    }
    const widthObsEl = document.getElementById('rwysurvey-width-obs-list');
    if (widthObsEl) widthObsEl.textContent = '';
    clearRunwaySurveyTrace();

    const thrA = runwaySurveyToolState.rwyIdent || 'RWY';
    const thrB = _rwySurveyReciprocalIdent_(runwaySurveyToolState.rwyIdent);
    const opts = '<option value="' + thrA + '">' + thrA + '</option><option value="' + thrB + '">' + thrB + '</option>';
    const thrSel1 = document.getElementById('rwysurvey-feature-thr');
    const thrSel2 = document.getElementById('rwysurvey-slope-thr');
    if (thrSel1) thrSel1.innerHTML = opts;
    if (thrSel2) thrSel2.innerHTML = opts;

    const officialEl = document.getElementById('rwysurvey-official');
    if (officialEl) {
      const internalLen = Math.round(Number(verifiedOperational.lengthM || runwaySurveyToolState.lengthM || 0)) || 0;
      const internalWid = Math.round(Number(verifiedOperational.widthM || row && (row.runwayWidth || row.width) || 0)) || 0;
      const internalSurface = String(verifiedOperational.surface || row && (row.runwaySurfaceActual || row.surface) || '').trim() || '-';
      const internalStamp = String(runwaySurveyToolState.internalUpdatedAt || '').trim();
      const internalDateLabel = internalStamp ? internalStamp.slice(0, 16).replace('T', ' ') + 'Z' : 'N/A';
      officialEl.innerHTML = '<b>Official runway reference:</b> '
        + (Math.round(runwaySurveyToolState.official.lengthM) || 0) + 'm length, '
        + (Math.round(runwaySurveyToolState.official.widthM) || 0) + 'm width, '
        + (runwaySurveyToolState.official.surface || '-')
        + '<br><b>Internal reference:</b> '
        + internalLen + 'm length, '
        + internalWid + 'm width, '
        + internalSurface
        + ' · date ' + internalDateLabel;
    }

    renderRunwaySurveyFeatureList();
    renderRunwaySurveyPhotoList();
    renderRunwaySurveyObstacleList();
    renderRunwaySurveySlopeList();
    renderRunwaySurveyPerimeterTally();
    renderRunwaySurveyMarkHistory();
    renderRunwaySurveyStatus();
    renderRunwaySurveyA50Status();
    renderRunwaySurveySlopeCaptureUi();
    renderRunwaySurveyMeasureTool();
    renderRunwaySurveyActionButtons();
  }

  function _rwySurveyDeg2Rad_(deg) { return (Number(deg || 0) * Math.PI) / 180; }

  function _rwySurveyMetersToLatLon_(baseLat, baseLon, eastM, northM) {
    const latRad = _rwySurveyDeg2Rad_(baseLat);
    const dLat = northM / 111320;
    const dLon = eastM / (111320 * Math.max(Math.cos(latRad), 0.2));
    return { lat: Number(baseLat || 0) + dLat, lon: Number(baseLon || 0) + dLon };
  }

  function _rwySurveyProjectAlongAxisM_(lat, lon, thrLat, thrLon, headingDeg) {
    const latRad = _rwySurveyDeg2Rad_(thrLat);
    const north = (Number(lat || 0) - Number(thrLat || 0)) * 111320;
    const east = (Number(lon || 0) - Number(thrLon || 0)) * 111320 * Math.max(Math.cos(latRad), 0.2);
    const br = _rwySurveyDeg2Rad_(headingDeg);
    return (east * Math.sin(br)) + (north * Math.cos(br));
  }

  function _rwySurveyDistanceMetersBetween_(a, b) {
    if (!a || !b) return 0;
    const toRad = function(d) { return (Number(d || 0) * Math.PI) / 180; };
    const R = 6371000;
    const lat1 = toRad(a.lat), lon1 = toRad(a.lon), lat2 = toRad(b.lat), lon2 = toRad(b.lon);
    const dLat = lat2 - lat1;
    const dLon = lon2 - lon1;
    const h = Math.sin(dLat / 2) * Math.sin(dLat / 2)
      + Math.cos(lat1) * Math.cos(lat2) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
    return 2 * R * Math.asin(Math.min(1, Math.sqrt(h)));
  }

  function _rwySurveyAveragedFix_() {
    // Average the last N high-accuracy GPS points to get a stable corner position.
    // This is far more accurate than any single GPS reading for marking corners.
    const gps = (runwaySurveyToolState && runwaySurveyToolState.gps) || {};
    const allPoints = Array.isArray(gps.points) ? gps.points : [];
    const maxAcc = 6; // only use points with <=6m accuracy
    const windowMs = 8000; // look back up to 8 seconds
    const now = Date.now();

    const candidates = allPoints.filter(function(p) {
      const age = now - Number(p && p.ts || 0);
      return age >= 0 && age <= windowMs && Number(p && p.acc || 9999) <= maxAcc;
    });

    if (!candidates.length) {
      // fallback to last known current fix
      const cur = gps.current;
      return cur ? { lat: Number(cur.lat), lon: Number(cur.lon), ts: cur.ts, acc: Number(cur.acc || 99), averaged: false, n: 1 } : null;
    }

    // Weight each point by 1/acc² so better fixes dominate
    let wLat = 0, wLon = 0, wSum = 0;
    candidates.forEach(function(p) {
      const w = 1 / Math.max(Number(p.acc || 1), 0.5) / Math.max(Number(p.acc || 1), 0.5);
      wLat += Number(p.lat) * w;
      wLon += Number(p.lon) * w;
      wSum += w;
    });

    const avgAcc = candidates.reduce(function(s, p) { return s + Number(p.acc || 4); }, 0) / candidates.length;
    return {
      lat: wLat / wSum,
      lon: wLon / wSum,
      ts: Number(candidates[candidates.length - 1].ts || now),
      acc: avgAcc,
      averaged: true,
      n: candidates.length
    };
  }

  function _rwySurveyAveragedFixInWindow_(startTs, endTs, maxAcc, fallbackFix) {
    const gps = (runwaySurveyToolState && runwaySurveyToolState.gps) || {};
    const allPoints = Array.isArray(gps.points) ? gps.points : [];
    const start = Number(startTs || 0);
    const end = Number(endTs || Date.now());
    const accLimit = Number(maxAcc || 6);

    const candidates = allPoints.filter(function(p) {
      const ts = Number(p && p.ts || 0);
      return ts >= start && ts <= end && Number(p && p.acc || 9999) <= accLimit;
    });

    if (!candidates.length) {
      if (!fallbackFix) return null;
      return {
        lat: Number(fallbackFix.lat),
        lon: Number(fallbackFix.lon),
        ts: Number(fallbackFix.ts || end),
        acc: Number(fallbackFix.acc || 99),
        averaged: false,
        n: 1
      };
    }

    let wLat = 0, wLon = 0, wSum = 0;
    candidates.forEach(function(p) {
      const w = 1 / Math.max(Number(p.acc || 1), 0.5) / Math.max(Number(p.acc || 1), 0.5);
      wLat += Number(p.lat) * w;
      wLon += Number(p.lon) * w;
      wSum += w;
    });

    const avgAcc = candidates.reduce(function(s, p) { return s + Number(p.acc || 4); }, 0) / candidates.length;
    return {
      lat: wLat / wSum,
      lon: wLon / wSum,
      ts: Number(candidates[candidates.length - 1].ts || end),
      acc: avgAcc,
      averaged: true,
      n: candidates.length
    };
  }

  function _rwySurveyQueueSettledCapture_(options) {
    const state = runwaySurveyToolState;
    const gps = state.gps || {};
    const rawFix = gps.current;
    const opts = options || {};
    const settleMs = Math.max(1200, Number(opts.settleMs || 2500));
    const maxCurrentAcc = Number(opts.maxCurrentAcc || 10);

    if (!gps.tracking) {
      if (window.M) M.toast({ html: 'Start GPS tracking first', classes: 'orange' });
      return;
    }
    if (gps.paused) {
      if (window.M) M.toast({ html: 'Resume GPS first', classes: 'orange' });
      return;
    }
    if ((state.capture && state.capture.pending) || false) {
      if (window.M) M.toast({ html: 'Capture already in progress. Hold still.', classes: 'blue darken-2' });
      return;
    }
    if (!rawFix) {
      _rwySurveyLogEvent_('reject', 'Capture ignored: no GPS fix');
      if (window.M) M.toast({ html: 'No current GPS fix', classes: 'orange' });
      return;
    }
    if (Number(rawFix.acc || 9999) > maxCurrentAcc) {
      _rwySurveyLogEvent_('reject', 'Capture ignored: accuracy >' + maxCurrentAcc + 'm', { acc: Number(rawFix.acc || 9999) });
      if (window.M) M.toast({ html: 'GPS accuracy >' + maxCurrentAcc + 'm. Hold still for a better fix.', classes: 'orange' });
      return;
    }

    const startedTs = Date.now();
    const label = String(opts.label || 'point').trim();
    const fallbackFix = { lat: rawFix.lat, lon: rawFix.lon, ts: rawFix.ts || startedTs, acc: rawFix.acc };
    state.capture = {
      pending: true,
      kind: String(opts.kind || 'capture'),
      label: label,
      startedTs: startedTs,
      settleMs: settleMs,
      timerId: null
    };
    _rwySurveyLogEvent_('capture', 'Capture started: ' + label, { settleMs: settleMs, acc: Number(rawFix.acc || 0) });
    if (window.M) M.toast({ html: 'Hold still · saving ' + label + ' in ' + (settleMs / 1000).toFixed(1) + 's', classes: 'blue darken-2', displayLength: settleMs + 900 });
    renderRunwaySurveyStatus();
    renderRunwaySurveyPerimeterTally();
    renderRunwaySurveyActionButtons();

    state.capture.timerId = setTimeout(function() {
      const activeCapture = runwaySurveyToolState.capture || {};
      if (!activeCapture.pending || Number(activeCapture.startedTs || 0) !== startedTs) return;

      _rwySurveyClearCaptureTimer_();
      const fix = _rwySurveyAveragedFixInWindow_(startedTs, Date.now(), Number(opts.maxAverageAcc || 6), runwaySurveyToolState.gps.current || fallbackFix);
      runwaySurveyToolState.capture = _rwySurveyCaptureDefaults_();
      renderRunwaySurveyStatus();
      renderRunwaySurveyPerimeterTally();
      renderRunwaySurveyActionButtons();

      if (!fix) {
        _rwySurveyLogEvent_('reject', 'Capture failed: no usable fix', { label: label });
        if (window.M) M.toast({ html: 'Could not save ' + label + '. Try again.', classes: 'orange' });
        return;
      }
      if (Number(fix.acc || 9999) > maxCurrentAcc) {
        _rwySurveyLogEvent_('reject', 'Capture failed: final accuracy >' + maxCurrentAcc + 'm', { acc: Number(fix.acc || 9999), label: label });
        if (window.M) M.toast({ html: 'Saved fix still too loose. Try again after standing still longer.', classes: 'orange' });
        return;
      }
      if (typeof opts.onComplete === 'function') opts.onComplete(fix);
    }, settleMs);
  }

  function _rwySurveySegmentTypeForCornerStep_(fromCorner, toCorner) {
    const pair = String(fromCorner || '') + String(toCorner || '');
    if (pair === 'AB' || pair === 'CD') return 'length';
    if (pair === 'BC' || pair === 'DE') return 'width';
    return 'length';
  }

  function _rwySurveyCommitCornerMark_(cornerLabel, fix) {
    const state = runwaySurveyToolState;
    if (!state.perimeter) state.perimeter = _rwySurveyPerimeterDefaults_();
    if (!Array.isArray(state.perimeter.cornerMarks)) state.perimeter.cornerMarks = [];

    const per = state.perimeter;
    const corners = _rwySurveyCornerSequence_();
    const expectedCorner = corners[Math.min(per.cornerMarks.length, corners.length - 1)];
    const requested = String(cornerLabel || '').trim().toUpperCase();
    if (requested !== expectedCorner) {
      _rwySurveyLogEvent_('reject', 'Corner ignored: expected ' + expectedCorner + ', got ' + requested);
      if (window.M) M.toast({ html: 'Tap corner ' + expectedCorner + ' next', classes: 'orange' });
      renderRunwaySurveyActionButtons();
      return;
    }

    per.cornerMarks.push({
      label: requested,
      lat: fix.lat,
      lon: fix.lon,
      ts: fix.ts || Date.now(),
      acc: fix.acc || null,
      averaged: !!fix.averaged,
      n: Number(fix.n || 1)
    });

    if (!per.startTs) per.startTs = Number(fix.ts || Date.now());
    per.lastMarkFix = fix;
    per.liveSinceMark = 0;

    const count = per.cornerMarks.length;
    if (count >= 2) {
      const prevCorner = per.cornerMarks[count - 2];
      const currCorner = per.cornerMarks[count - 1];
      const segType = _rwySurveySegmentTypeForCornerStep_(prevCorner.label, currCorner.label);
      const distM = Math.round(_rwySurveyDistanceMetersBetween_(prevCorner, currCorner));
      if (distM > 0) {
        per.segments.push({
          type: segType,
          distanceM: distM,
          fromCorner: prevCorner.label,
          toCorner: currCorner.label,
          from: { lat: prevCorner.lat, lon: prevCorner.lon, ts: prevCorner.ts || null, acc: prevCorner.acc || null },
          to: { lat: currCorner.lat, lon: currCorner.lon, ts: currCorner.ts || null, acc: currCorner.acc || null },
          markedAt: new Date().toISOString()
        });
        if (segType === 'length') per.lengthWalkedM += distM;
        else per.widthWalkedM += distM;
      }
    }

    if (count === 5) {
      const a = per.cornerMarks[0];
      const e = per.cornerMarks[4];
      per.closureErrorM = Math.round(_rwySurveyDistanceMetersBetween_(e, a));
      per.closed = true;
      _rwySurveyLogEvent_('mark', 'Corner E saved · perimeter complete', { acc: Number(fix.acc || 0), pointsAveraged: Number(fix.n || 1), corners: 5, closureErrorM: Number(per.closureErrorM || 0) });
      if (window.M) M.toast({ html: 'Corner E saved · perimeter complete', classes: 'green darken-1', displayLength: 3200 });
    } else {
      const nextCorner = corners[Math.min(count, corners.length - 1)];
      _rwySurveyLogEvent_('mark', 'Corner ' + requested + ' saved', { acc: Number(fix.acc || 0), pointsAveraged: Number(fix.n || 1), corners: count });
      if (window.M) M.toast({ html: 'Corner ' + requested + ' saved. Next: ' + nextCorner, classes: 'green' });
    }

    renderRunwaySurveyPerimeterTally();
    renderRunwaySurveyActionButtons();
  }

  function _rwySurveyCommitWidthObservation_(label, fix) {
    const state = runwaySurveyToolState;
    if (!Array.isArray(state.widthObservations)) state.widthObservations = [];
    state.widthObservations.push({
      label: label,
      lat: fix.lat,
      lon: fix.lon,
      acc: fix.acc,
      n: fix.n || 1,
      ts: new Date().toISOString()
    });
    const el = document.getElementById('rwysurvey-width-obs-list');
    if (el) {
      el.textContent = state.widthObservations.map(function(o) {
        return o.label.toUpperCase();
      }).join(' · ');
    }
    _rwySurveyLogEvent_('mark', 'Width observation saved: ' + label, { acc: Number(fix.acc || 0), pointsAveraged: Number(fix.n || 1) });
    renderRunwaySurveyActionButtons();
    if (window.M) M.toast({ html: label.toUpperCase() + ' spot marked at ±' + Math.round(fix.acc) + 'm accuracy', classes: label === 'narrow' ? 'light-blue darken-3' : 'deep-orange darken-2' });
  }

  async function markRunwayCorner(cornerLabel) {
    const requested = String(cornerLabel || '').trim().toUpperCase();
    if (!String(runwaySurveyToolState.rwyIdent || '').trim()) {
      if (window.M) M.toast({ html: 'Select runway before marking corners', classes: 'orange' });
      renderRunwaySurveyActionButtons();
      return;
    }
    const corners = _rwySurveyCornerSequence_();
    if (corners.indexOf(requested) < 0) {
      if (window.M) M.toast({ html: 'Invalid corner label', classes: 'orange' });
      return;
    }
    const per = runwaySurveyToolState.perimeter || _rwySurveyPerimeterDefaults_();
    const expected = corners[Math.min((Array.isArray(per.cornerMarks) ? per.cornerMarks.length : 0), corners.length - 1)];
    if (requested !== expected) {
      _rwySurveyLogEvent_('reject', 'Corner ignored: expected ' + expected + ', got ' + requested);
      if (window.M) M.toast({ html: 'Please mark corner ' + expected + ' next', classes: 'orange' });
      renderRunwaySurveyActionButtons();
      return;
    }

    if (requested === 'A' && (!per.cornerMarks || !per.cornerMarks.length)) {
      const okStartThreshold = await _rwySurveyEnsureStartThresholdChosen_();
      if (!okStartThreshold) return;
    }

    _rwySurveyQueueSettledCapture_({
      kind: 'corner',
      label: 'corner ' + requested,
      settleMs: 2500,
      maxCurrentAcc: 10,
      maxAverageAcc: 6,
      onComplete: function(fix) {
        _rwySurveyCommitCornerMark_(requested, fix);
      }
    });
  }

  function markRunwayThresholdTurn() {
    markRunwayCorner('A');
  }

  function markRunwayNextSideline() {
    markRunwayCorner('B');
  }

  function closeRunwayPerimeter() {
    markRunwayCorner('D');
  }

  function markRunwayWidthObservation(label) {
    _rwySurveyQueueSettledCapture_({
      kind: 'width-observation',
      label: label + ' width spot',
      settleMs: 2200,
      maxCurrentAcc: 15,
      maxAverageAcc: 8,
      onComplete: function(fix) {
        _rwySurveyCommitWidthObservation_(label, fix);
      }
    });
  }

  function _rwySurveyMedian_(arr) {
    const values = (Array.isArray(arr) ? arr : []).map(function(v) { return Number(v); }).filter(function(v) { return isFinite(v); }).sort(function(a, b) { return a - b; });
    if (!values.length) return 0;
    const mid = Math.floor(values.length / 2);
    return values.length % 2 ? values[mid] : (values[mid - 1] + values[mid]) / 2;
  }

  function _rwySurveyTrimmedMean_(arr, trimRatio) {
    const values = (Array.isArray(arr) ? arr : []).map(function(v) { return Number(v); }).filter(function(v) { return isFinite(v); }).sort(function(a, b) { return a - b; });
    if (!values.length) return 0;
    const trim = Math.max(0, Math.min(Math.floor(values.length * Number(trimRatio || 0)), Math.floor(values.length / 2) - 1));
    const kept = trim > 0 ? values.slice(trim, values.length - trim) : values;
    return kept.reduce(function(s, v) { return s + v; }, 0) / kept.length;
  }

  function _rwySurveyLateralWidthEstimate_() {
    const per = runwaySurveyToolState.perimeter || {};
    const marks = Array.isArray(per.cornerMarks) ? per.cornerMarks : [];
    const gps = runwaySurveyToolState.gps || {};
    const points = Array.isArray(gps.points) ? gps.points : [];

    if (!marks.length || !points.length) return { widthM: 0, samples: 0 };
    const byLabel = {};
    marks.forEach(function(m) { byLabel[String(m.label || '').toUpperCase()] = m; });
    const a = byLabel.A, b = byLabel.B, c = byLabel.C, d = byLabel.D;
    if (!a || !b || !c || !d) return { widthM: 0, samples: 0 };

    const tsAB0 = Math.min(Number(a.ts || 0), Number(b.ts || 0));
    const tsAB1 = Math.max(Number(a.ts || 0), Number(b.ts || 0));
    const tsCD0 = Math.min(Number(c.ts || 0), Number(d.ts || 0));
    const tsCD1 = Math.max(Number(c.ts || 0), Number(d.ts || 0));

    const sideAB = points.filter(function(p) {
      const ts = Number(p && p.ts || 0);
      return ts >= tsAB0 && ts <= tsAB1 && Number(p && p.acc || 9999) <= 12;
    });
    const sideCD = points.filter(function(p) {
      const ts = Number(p && p.ts || 0);
      return ts >= tsCD0 && ts <= tsCD1 && Number(p && p.acc || 9999) <= 12;
    });

    if (sideAB.length < 4 || sideCD.length < 4) return { widthM: 0, samples: 0 };

    const latRad = _rwySurveyDeg2Rad_(a.lat);
    const toEN = function(p) {
      const north = (Number(p.lat || 0) - Number(a.lat || 0)) * 111320;
      const east = (Number(p.lon || 0) - Number(a.lon || 0)) * 111320 * Math.max(Math.cos(latRad), 0.2);
      return { east: east, north: north };
    };

    const bEN = toEN(b);
    const mag = Math.sqrt((bEN.east * bEN.east) + (bEN.north * bEN.north));
    if (!(mag > 10)) return { widthM: 0, samples: 0 };
    const ux = bEN.east / mag;
    const uy = bEN.north / mag;
    const nx = -uy;
    const ny = ux;

    const project = function(p) {
      const en = toEN(p);
      return {
        along: (en.east * ux) + (en.north * uy),
        cross: (en.east * nx) + (en.north * ny)
      };
    };

    const abProj = sideAB.map(project).sort(function(x, y) { return x.along - y.along; });
    const cdProj = sideCD.map(project);
    const widthSamples = [];

    cdProj.forEach(function(cp) {
      let best = null;
      let bestDelta = Infinity;
      for (let i = 0; i < abProj.length; i++) {
        const delta = Math.abs(abProj[i].along - cp.along);
        if (delta < bestDelta) {
          bestDelta = delta;
          best = abProj[i];
        }
      }
      if (best && bestDelta <= 45) {
        widthSamples.push(Math.abs(cp.cross - best.cross));
      }
    });

    if (widthSamples.length >= 6) {
      return { widthM: _rwySurveyTrimmedMean_(widthSamples, 0.15), samples: widthSamples.length };
    }

    const abCross = _rwySurveyMedian_(abProj.map(function(p) { return p.cross; }));
    const cdCross = _rwySurveyMedian_(cdProj.map(function(p) { return p.cross; }));
    const fallback = Math.abs(cdCross - abCross);
    return { widthM: fallback > 0 ? fallback : 0, samples: widthSamples.length };
  }

  function _rwySurveyDerivedDimensions_() {
    const per = runwaySurveyToolState.perimeter || {};
    const segs = Array.isArray(per.segments) ? per.segments : [];

    // If perimeter was closed (pilot started mid-side), merge the first and last
    // segments of the same type into one combined measurement before averaging.
    let mergedSegs = segs.slice();
    if (per.closed && segs.length >= 3) {
      const first = segs[0];
      const last = segs[segs.length - 1];
      if (first && last && first.type === last.type) {
        const combined = Object.assign({}, first, {
          distanceM: Number(first.distanceM || 0) + Number(last.distanceM || 0),
          merged: true
        });
        mergedSegs = [combined].concat(segs.slice(1, segs.length - 1));
      }
    }

    const lenSegs = mergedSegs.filter(function(s) { return s && s.type === 'length' && Number(s.distanceM || 0) > 0; });
    const widSegs = mergedSegs.filter(function(s) { return s && s.type === 'width' && Number(s.distanceM || 0) > 0; });
    const lenAvg = lenSegs.length ? (lenSegs.reduce(function(acc, s) { return acc + Number(s.distanceM || 0); }, 0) / lenSegs.length) : 0;
    const widAvg = widSegs.length ? (widSegs.reduce(function(acc, s) { return acc + Number(s.distanceM || 0); }, 0) / widSegs.length) : 0;
    const lateral = _rwySurveyLateralWidthEstimate_();
    const widthLateral = Number(lateral.widthM || 0);
    const widthFinal = widthLateral > 0 ? widthLateral : widAvg;
    return {
      lengthM: lenAvg,
      widthM: widthFinal,
      widthCornerM: widAvg,
      widthLateralM: widthLateral,
      widthLateralSamples: Number(lateral.samples || 0),
      lengthSegments: lenSegs.length,
      widthSegments: widSegs.length,
      lengthWalkedM: Number(per.lengthWalkedM || 0),
      widthWalkedM: Number(per.widthWalkedM || 0),
      closureErrorM: Number(per.closureErrorM || 0)
    };
  }

  function renderRunwaySurveyPerimeterTally() {
    const el = document.getElementById('rwysurvey-perimeter-tally');
    if (!el) return;
    const per = runwaySurveyToolState.perimeter || {};
    const cornerMarks = Array.isArray(per.cornerMarks) ? per.cornerMarks : [];
    const corners = _rwySurveyCornerSequence_();
    const segs = Array.isArray(per.segments) ? per.segments : [];
    const d = _rwySurveyDerivedDimensions_();
    const capture = runwaySurveyToolState.capture || {};
    const lenDer = d.lengthM ? Math.round(d.lengthM) + 'm' : '--';
    const widDer = d.widthM ? Math.round(d.widthM) + 'm' : '--';

    if (!String(runwaySurveyToolState.rwyIdent || '').trim()) {
      el.textContent = 'Selecione uma pista para habilitar MARK A/B/C/D/E.';
      return;
    }

    if (!cornerMarks.length) {
      if (capture.pending) {
        el.innerHTML = '<b>Hold still:</b> saving ' + String(capture.label || 'point') + '...';
        return;
      }
      el.textContent = 'Stand still at corner A, tap MARK A, and keep still until it confirms.';
      return;
    }

    const live = Number(per.liveSinceMark || 0);
    const lastCorner = cornerMarks[Math.max(cornerMarks.length - 1, 0)] || {};
    let html = '<b>Corners:</b> ' + cornerMarks.map(function(c) { return c.label; }).join(' → ');

    if (!per.closed) {
      const nextCorner = corners[Math.min(cornerMarks.length, corners.length - 1)];
      html += '<span style="color:#888;"> &nbsp;|&nbsp; </span>Next: <b>' + nextCorner + '</b>';
    }

    html += '<span style="color:#888;"> &nbsp;|&nbsp; </span>';
    html += 'since ' + String(lastCorner.label || 'last') + ': <b>' + live + 'm</b>';

    if (capture.pending) {
      html += '<span style="color:#888;"> &nbsp;|&nbsp; </span>';
      html += '<span style="color:#1565c0; font-weight:700;">saving ' + String(capture.label || 'point') + '...</span>';
    }

    if (segs.length) {
      html += '<span style="color:#888;"> &nbsp;|&nbsp; </span>';
      html += segs.map(function(s, i) {
        const label = s.type === 'length' ? 'L' : 'W';
        const color = s.type === 'length' ? '#1565c0' : '#6a1b9a';
        return '<span style="color:' + color + '; font-weight:700;">' + label + (i + 1) + ':' + Math.round(s.distanceM) + 'm</span>';
      }).join('<span style="color:#aaa;"> · </span>');
    }

    html += '<span style="color:#888;"> &nbsp;|&nbsp; </span>';
    html += 'derived: <b>' + lenDer + ' × ' + widDer + '</b>';

    if (Number(d.widthLateralSamples || 0) > 0) {
      html += '<span style="color:#888;"> &nbsp;|&nbsp; </span>';
      html += 'lat avg: <b>' + Math.round(Number(d.widthLateralM || 0)) + 'm</b> (' + Number(d.widthLateralSamples || 0) + ' pts)';
    }

    if (Number(d.closureErrorM || 0) > 0) {
      html += '<span style="color:#888;"> &nbsp;|&nbsp; </span>';
      html += 'closure err: <b>' + Math.round(Number(d.closureErrorM || 0)) + 'm</b>';
    }

    el.innerHTML = html;
  }

  function _rwySurveyMaybeNotifyThresholdProximity_(fix) {
    const per = runwaySurveyToolState.perimeter || {};
    const marks = Array.isArray(per.cornerMarks) ? per.cornerMarks : [];
    const prompts = runwaySurveyToolState.prompts || {
      a50Shown: false, a50Completed: false, a50Lat: null, a50Lon: null,
      a300Shown: false, a300Completed: false, a300Lat: null, a300Lon: null,
      c50Shown: false, c50Completed: false, c50Lat: null, c50Lon: null,
      c300Shown: false, c300Completed: false, c300Lat: null, c300Lon: null
    };
    runwaySurveyToolState.prompts = prompts;
    if (!fix) return;

    const distSinceA = Number(per.liveSinceMark || 0);
    const distSinceLastMark = Number(per.liveSinceMark || 0);

    // A+50m checkpoint: only when exactly 1 corner marked and not closed
    if (!prompts.a50Shown && !prompts.a50Completed && marks.length === 1 && !per.closed && distSinceA >= 50) {
      prompts.a50Shown = true;
      prompts.a50Lat = Number(fix.lat || 0);
      prompts.a50Lon = Number(fix.lon || 0);
      _rwySurveyLogEvent_('prompt', 'A+50m checkpoint reached');
      openRunwaySurveyObstaclePopup(true, 'A', 50);
      return;
    }

    if (!prompts.a300Shown && !prompts.a300Completed && marks.length === 1 && !per.closed && distSinceA >= 300) {
      prompts.a300Shown = true;
      prompts.a300Lat = Number(fix.lat || 0);
      prompts.a300Lon = Number(fix.lon || 0);
      _rwySurveyLogEvent_('prompt', 'A+300m checkpoint reached');
      openRunwaySurveyObstaclePopup(true, 'A', 300);
      return;
    }

    // C+50m checkpoint: only when exactly 3 corners marked and not closed
    if (!prompts.c50Shown && !prompts.c50Completed && marks.length === 3 && !per.closed && distSinceLastMark >= 50) {
      prompts.c50Shown = true;
      prompts.c50Lat = Number(fix.lat || 0);
      prompts.c50Lon = Number(fix.lon || 0);
      _rwySurveyLogEvent_('prompt', 'C+50m checkpoint reached');
      openRunwaySurveyObstaclePopup(true, 'C', 50);
      return;
    }

    if (!prompts.c300Shown && !prompts.c300Completed && marks.length === 3 && !per.closed && distSinceLastMark >= 300) {
      prompts.c300Shown = true;
      prompts.c300Lat = Number(fix.lat || 0);
      prompts.c300Lon = Number(fix.lon || 0);
      _rwySurveyLogEvent_('prompt', 'C+300m checkpoint reached');
      openRunwaySurveyObstaclePopup(true, 'C', 300);
      return;
    }
  }

  function startRunwaySurveyGps() {
    if (!navigator.geolocation) {
      if (window.M) M.toast({ html: 'Geolocation not supported on this device', classes: 'red' });
      return;
    }
    const gps = runwaySurveyToolState.gps;
    if (gps.watchId != null) navigator.geolocation.clearWatch(gps.watchId);
    gps.tracking = true;
    gps.paused = false;

    gps.watchId = navigator.geolocation.watchPosition(function(pos) {
      const c = pos && pos.coords ? pos.coords : null;
      if (!c || gps.paused) return;
      const fix = {
        ts: Date.now(),
        lat: Number(c.latitude || 0),
        lon: Number(c.longitude || 0),
        acc: Number(c.accuracy || 9999)
      };
      gps.current = fix;
      gps.points.push(fix);
      gps.samples += 1;
      gps.bestAcc = Math.min(gps.bestAcc, fix.acc);
      gps.avgAcc = ((gps.avgAcc * (gps.samples - 1)) + fix.acc) / gps.samples;

      if (!runwaySurveyToolState.thresholdA) {
        runwaySurveyToolState.thresholdA = { ident: runwaySurveyToolState.rwyIdent, lat: fix.lat, lon: fix.lon };
        const heading = runwaySurveyToolState.headingDeg || 0;
        const lengthM = runwaySurveyToolState.lengthM || runwaySurveyToolState.official.lengthM || 0;
        const br = _rwySurveyDeg2Rad_(heading);
        const east = Math.sin(br) * lengthM;
        const north = Math.cos(br) * lengthM;
        const thrBpt = _rwySurveyMetersToLatLon_(fix.lat, fix.lon, east, north);
        runwaySurveyToolState.thresholdB = { ident: _rwySurveyReciprocalIdent_(runwaySurveyToolState.rwyIdent), lat: thrBpt.lat, lon: thrBpt.lon };
        // NOTE: perimeter anchor is NOT auto-set here. Pilot must stand still
        // at their start position and press a corner button to set the anchor.
      }

      // Update live distance from last mark so pilot can see how far they've walked
      if (runwaySurveyToolState.perimeter && runwaySurveyToolState.perimeter.lastMarkFix) {
        runwaySurveyToolState.perimeter.liveSinceMark =
          Math.round(_rwySurveyDistanceMetersBetween_(runwaySurveyToolState.perimeter.lastMarkFix, fix));
      }

      _rwySurveyMaybeNotifyThresholdProximity_(fix);

      renderRunwaySurveyStatus();
      renderRunwaySurveyPerimeterTally();
      renderRunwaySurveyMeasureTool();
      renderRunwaySurveyTrace();
    }, function(err) {
      if (window.M) M.toast({ html: 'GPS error: ' + (err && err.message ? err.message : 'unknown'), classes: 'red' });
    }, { enableHighAccuracy: true, maximumAge: 0, timeout: 10000 });

    renderRunwaySurveyStatus();
    renderRunwaySurveyActionButtons();
  }

  function toggleRunwaySurveyGps() {
    const gps = runwaySurveyToolState.gps || {};
    if (!gps.tracking) {
      startRunwaySurveyGps();
      return;
    }
    if (gps.paused) {
      resumeRunwaySurveyGps();
      return;
    }
    stopRunwaySurveyGps(false);
  }

  function pauseRunwaySurveyGps() {
    runwaySurveyToolState.gps.paused = true;
    renderRunwaySurveyStatus();
    renderRunwaySurveyActionButtons();
  }

  function resumeRunwaySurveyGps() {
    if (!runwaySurveyToolState.gps.tracking) {
      startRunwaySurveyGps();
      return;
    }
    runwaySurveyToolState.gps.paused = false;
    renderRunwaySurveyStatus();
    renderRunwaySurveyActionButtons();
  }

  function stopRunwaySurveyGps(hard) {
    _rwySurveyCancelPendingCapture_('Capture cancelled', true);
    const gps = runwaySurveyToolState.gps;
    if (gps.watchId != null && navigator.geolocation) {
      navigator.geolocation.clearWatch(gps.watchId);
      gps.watchId = null;
    }
    gps.tracking = false;
    gps.paused = false;
    if (hard) gps.current = null;
    renderRunwaySurveyStatus();
    renderRunwaySurveyActionButtons();
  }

  function clearRunwaySurveyTrace() {
    _rwySurveyCancelPendingCapture_('Survey reset', true);
    closeRunwaySurveyFeaturePopup(true);
    closeRunwaySurveyObstaclePopup(true);
    const gps = runwaySurveyToolState.gps;
    gps.points = [];
    gps.current = null;
    gps.bestAcc = Infinity;
    gps.avgAcc = 0;
    gps.samples = 0;
    runwaySurveyToolState.thresholdA = null;
    runwaySurveyToolState.thresholdB = null;
    runwaySurveyToolState.startThresholdIdent = '';
    runwaySurveyToolState.reciprocalThresholdIdent = '';
    runwaySurveyToolState.perimeter = _rwySurveyPerimeterDefaults_();
    runwaySurveyToolState.capture = _rwySurveyCaptureDefaults_();
    runwaySurveyToolState.ui = { pausedByPopup: false, activeObstaclePrompt: { corner: 'A', distanceM: 50 } };
    runwaySurveyToolState.prompts = {
      a50Shown: false, a50Completed: false, a50Lat: null, a50Lon: null,
      a300Shown: false, a300Completed: false, a300Lat: null, a300Lon: null,
      c50Shown: false, c50Completed: false, c50Lat: null, c50Lon: null,
      c300Shown: false, c300Completed: false, c300Lat: null, c300Lon: null
    };
    runwaySurveyToolState.slopeCapture = { active: false, startIdx: -1, startTs: null, pendingStartDistanceM: 0, pendingDistanceM: 0, pendingFromThreshold: '' };
    runwaySurveyToolState.measureTool = _rwySurveyMeasureDefaults_();
    runwaySurveyToolState.widthObservations = [];
    runwaySurveyToolState.debugEvents = [];
    runwaySurveyToolState.thresholdAlerts = { thrA: { m300: false, m50: false }, thrB: { m300: false, m50: false } };
    const widthObsEl = document.getElementById('rwysurvey-width-obs-list');
    if (widthObsEl) widthObsEl.textContent = '';
    renderRunwaySurveyStatus();
    renderRunwaySurveyPerimeterTally();
    renderRunwaySurveyMarkHistory();
    renderRunwaySurveyFeatureList();
    renderRunwaySurveyPhotoList();
    renderRunwaySurveyObstacleList();
    renderRunwaySurveyA50Status();
    renderRunwaySurveySlopeCaptureUi();
    renderRunwaySurveyMeasureTool();
    renderRunwaySurveyTrace();
    renderRunwaySurveyActionButtons();
  }

  function renderRunwaySurveyStatus() {
    const el = document.getElementById('rwysurvey-status');
    if (!el) return;
    const gps = runwaySurveyToolState.gps;
    const capture = runwaySurveyToolState.capture || {};
    const mode = gps.tracking ? (gps.paused ? 'Paused' : 'Tracking') : 'Idle';
    const curAcc = gps.current ? Math.round(Number(gps.current.acc || 0)) : 0;
    const best = isFinite(gps.bestAcc) ? Math.round(gps.bestAcc) : 0;
    const avg = gps.samples ? Math.round(gps.avgAcc) : 0;
    let text = 'GPS ' + mode + ' · points: ' + gps.points.length + ' · current acc: ' + curAcc + 'm · best: ' + best + 'm · avg: ' + avg + 'm';
    if (capture.pending) {
      const remainMs = Math.max(0, Number(capture.startedTs || 0) + Number(capture.settleMs || 0) - Date.now());
      text += ' · saving ' + String(capture.label || 'point') + ' (' + (remainMs / 1000).toFixed(1) + 's)';
    }
    el.textContent = text;
  }

  function renderRunwaySurveyTrace() {
    const svg = document.getElementById('rwysurvey-trace');
    if (!svg) return;
    const allPts = runwaySurveyToolState.gps.points;
    const startTs = Number(runwaySurveyToolState.perimeter && runwaySurveyToolState.perimeter.startTs || 0);
    const rawPts = startTs > 0 ? allPts.filter(function(p) { return Number(p && p.ts || 0) >= startTs; }) : allPts;
    if (!rawPts.length) { svg.innerHTML = ''; return; }

    const filtered = [];
    rawPts.forEach(function(p) {
      if (!p) return;
      if (!filtered.length) {
        filtered.push(p);
        return;
      }
      const prev = filtered[filtered.length - 1];
      const dt = Math.max(0.2, (Number(p.ts || 0) - Number(prev.ts || 0)) / 1000);
      const d = _rwySurveyDistanceMetersBetween_(prev, p);
      const speed = d / dt;
      if (d > 18 && speed > 8) return;
      filtered.push(p);
    });

    const pts = filtered.map(function(_, idx) {
      const a = filtered[Math.max(0, idx - 2)];
      const b = filtered[Math.max(0, idx - 1)];
      const c = filtered[idx];
      const d = filtered[Math.min(filtered.length - 1, idx + 1)];
      const e = filtered[Math.min(filtered.length - 1, idx + 2)];
      return {
        lat: (a.lat + b.lat + c.lat + d.lat + e.lat) / 5,
        lon: (a.lon + b.lon + c.lon + d.lon + e.lon) / 5,
        ts: c.ts,
        acc: c.acc
      };
    });
    if (!pts.length) { svg.innerHTML = ''; return; }

    let minLat = pts[0].lat, maxLat = pts[0].lat, minLon = pts[0].lon, maxLon = pts[0].lon;
    pts.forEach(function(p) {
      minLat = Math.min(minLat, p.lat); maxLat = Math.max(maxLat, p.lat);
      minLon = Math.min(minLon, p.lon); maxLon = Math.max(maxLon, p.lon);
    });

    const cornerMarks = (runwaySurveyToolState.perimeter && runwaySurveyToolState.perimeter.cornerMarks) || [];
    cornerMarks.forEach(function(c) {
      minLat = Math.min(minLat, c.lat); maxLat = Math.max(maxLat, c.lat);
      minLon = Math.min(minLon, c.lon); maxLon = Math.max(maxLon, c.lon);
    });
    const features = runwaySurveyToolState.features || [];
    features.forEach(function(f) {
      const g = f && f.gps;
      if (!g) return;
      minLat = Math.min(minLat, Number(g.lat || 0)); maxLat = Math.max(maxLat, Number(g.lat || 0));
      minLon = Math.min(minLon, Number(g.lon || 0)); maxLon = Math.max(maxLon, Number(g.lon || 0));
    });

    const dLat = Math.max(maxLat - minLat, 0.000001);
    const dLon = Math.max(maxLon - minLon, 0.000001);
    const cosLat = Math.cos(((minLat + maxLat) / 2) * Math.PI / 180);
    const latSpanM = dLat * 111320;
    const lonSpanM = dLon * 111320 * cosLat;

    // Auto-orient: put the longer GPS extent along the vertical axis for portrait display.
    // latDominant means the runway runs N/S — already portrait-natural.
    // lonDominant means E/W runway — we transpose axes so it still reads top-to-bottom.
    const W = 360, H = 520, pad = 16;
    const latDominant = latSpanM >= lonSpanM;

    var toX, toY;
    if (latDominant) {
      // Runway N/S: lon→X, lat→Y (north = top)
      var sx = (W - pad * 2) / dLon;
      var sy = (H - pad * 2) / dLat;
      toX = function(p) { return pad + (p.lon - minLon) * sx; };
      toY = function(p) { return H - pad - (p.lat - minLat) * sy; };
    } else {
      // Runway E/W: transpose so runway goes top-to-bottom
      var sx = (W - pad * 2) / dLat;
      var sy = (H - pad * 2) / dLon;
      toX = function(p) { return pad + (p.lat - minLat) * sx; };
      toY = function(p) { return pad + (maxLon - p.lon) * sy; };
    }

    svg.setAttribute('viewBox', '0 0 ' + W + ' ' + H);

    const poly = pts.map(function(p) {
      return toX(p).toFixed(1) + ',' + toY(p).toFixed(1);
    }).join(' ');

    let html = '<polyline points="' + poly + '" fill="none" stroke="#0b5394" stroke-width="1.5" />';

    // Corner dots and labels A/B/C/D
    cornerMarks.forEach(function(c) {
      if (!c) return;
      const cx = toX(c).toFixed(1);
      const cy = toY(c).toFixed(1);
      html += '<circle cx="' + cx + '" cy="' + cy + '" r="6" fill="#6a1b9a" stroke="#fff" stroke-width="1.5" />';
      html += '<text x="' + (Number(cx) + 8) + '" y="' + (Number(cy) + 4) + '" font-size="12" fill="#4a148c" font-weight="900">' + String(c.label || '') + '</text>';
    });

    // Orange triangles for width observations
    const wobs = runwaySurveyToolState.widthObservations || [];
    wobs.forEach(function(o) {
      const cx = toX(o).toFixed(1);
      const cy = toY(o).toFixed(1);
      const fill = o.label === 'narrow' ? '#0277bd' : '#e64a19';
      html += '<polygon points="' + cx + ',' + (Number(cy)-6) + ' ' + (Number(cx)-5) + ',' + (Number(cy)+5) + ' ' + (Number(cx)+5) + ',' + (Number(cy)+5) + '" fill="' + fill + '" stroke="#fff" stroke-width="1" />';
    });

    // Feature markers (auto-populate as captured)
    features.forEach(function(f, i) {
      const g = f && f.gps;
      if (!g) return;
      const cx = toX({ lat: Number(g.lat || 0), lon: Number(g.lon || 0) }).toFixed(1);
      const cy = toY({ lat: Number(g.lat || 0), lon: Number(g.lon || 0) }).toFixed(1);
      html += '<rect x="' + (Number(cx)-4) + '" y="' + (Number(cy)-4) + '" width="8" height="8" rx="1" fill="#1565c0" stroke="#fff" stroke-width="1.2" />';
      html += '<text x="' + (Number(cx) + 6) + '" y="' + (Number(cy) + 4) + '" font-size="10" fill="#0d47a1" font-weight="700">F' + (i + 1) + '</text>';
    });

    svg.innerHTML = html;
  }

  function addRunwaySurveyFeatureFromCurrentGps() {
    openRunwaySurveyFeaturePopup();
  }

  function saveRunwaySurveyFeaturePopup() {
    const gps = runwaySurveyToolState.gps || {};
    const fix = gps.current;
    if (!fix) {
      if (window.M) M.toast({ html: 'No current GPS fix', classes: 'orange' });
      return;
    }
    if (Number(fix.acc || 9999) > 12) {
      if (window.M) M.toast({ html: 'Current GPS accuracy >12m. Wait for better fix.', classes: 'orange' });
      return;
    }

    const type = rwySurveyGetFeatureType_();
    const side = String((document.getElementById('rwysurvey-feature-side-popup') || {}).value || 'right').trim().toLowerCase();
    if (!type) {
      if (window.M) M.toast({ html: 'Select or type a feature type', classes: 'orange' });
      return;
    }
    const per = runwaySurveyToolState.perimeter || {};
    const marks = Array.isArray(per.cornerMarks) ? per.cornerMarks : [];
    const origin = marks.length ? marks[0] : runwaySurveyToolState.thresholdA;
    if (!origin) {
      if (window.M) M.toast({ html: 'Mark corner A first', classes: 'orange' });
      return;
    }

    let heading = runwaySurveyToolState.headingDeg || 0;
    if (marks.length >= 2) {
      const a = marks[0], b = marks[1];
      const dy = Number(b.lat || 0) - Number(a.lat || 0);
      const dx = Number(b.lon || 0) - Number(a.lon || 0);
      heading = (Math.atan2(dx, dy) * 180 / Math.PI + 360) % 360;
    }

    let along = _rwySurveyProjectAlongAxisM_(fix.lat, fix.lon, origin.lat, origin.lon, heading);
    along = Math.max(0, Math.round(along));

    runwaySurveyToolState.features.push({
      name: type.replace(/_/g, ' '),
      type: type,
      distance: along,
      fromThreshold: _rwySurveyGetActiveStartThreshold_() || 'A',
      side: side,
      gps: { lat: fix.lat, lon: fix.lon, accuracyM: fix.acc, at: new Date(fix.ts).toISOString() }
    });
    _rwySurveyLogEvent_('mark', 'Feature added: ' + type + ' @ ' + along + 'm', { side: side, acc: Number(fix.acc || 0) });
    closeRunwaySurveyFeaturePopup(false);
    renderRunwaySurveyFeatureList();
    renderRunwaySurveyTrace();
    if (window.M) M.toast({ html: 'Feature added at ~' + along + 'm from ' + (_rwySurveyGetActiveStartThreshold_() || 'start threshold'), classes: 'green' });
  }

  function renderRunwaySurveyFeatureList() {
    const el = document.getElementById('rwysurvey-feature-list');
    if (!el) return;
    const items = runwaySurveyToolState.features;
    const summary = document.getElementById('rwysurvey-feature-summary');
    if (summary) summary.textContent = 'Features: ' + items.length;
    if (!items.length) {
      el.innerHTML = '<div style="font-size:0.8rem; color:#999; margin-top:4px;">No captured features yet.</div>';
      return;
    }
    el.innerHTML = items.map(function(f, i) {
      const fromThr = String(f && f.fromThreshold || _rwySurveyGetActiveStartThreshold_() || 'A');
      return '<div style="display:flex; justify-content:space-between; align-items:center; gap:8px; margin-top:4px; padding:5px 8px; border:1px solid #e0e0e0; border-radius:6px; font-size:0.82rem;">'
        + '<span><b>' + f.name + '</b> · ' + f.distance + 'm from ' + fromThr + ' · ' + f.side + ' · acc ' + Math.round(Number(f.gps && f.gps.accuracyM || 0)) + 'm</span>'
        + '<button onclick="removeRunwaySurveyFeature(' + i + ')" style="border:none; background:none; color:#d32f2f; cursor:pointer;">✕</button></div>';
    }).join('');
  }

  function removeRunwaySurveyFeature(i) {
    runwaySurveyToolState.features.splice(i, 1);
    renderRunwaySurveyFeatureList();
    renderRunwaySurveyTrace();
  }

  function addRunwaySurveyObstacleAngle() {
    const type = rwySurveyGetObstacleType_();
    const a = Number((document.getElementById('rwysurvey-obs-a-popup') || {}).value);
    const notes = String((document.getElementById('rwysurvey-obs-notes-popup') || {}).value || '').trim();
    if (!type) return;
    if (!isFinite(a)) {
      if (window.M) M.toast({ html: 'Enter valid obstacle angle', classes: 'orange' });
      return;
    }
    const obsCtx = _rwySurveyObstacleContextFromUi_();
    if (!obsCtx.operation) {
      if (window.M) M.toast({ html: 'Choose landing or takeoff for non-50m/300m obstacle entries', classes: 'orange' });
      return;
    }
    const photoRef = runwaySurveyToolState.ui && runwaySurveyToolState.ui.obstaclePhoto ? runwaySurveyToolState.ui.obstaclePhoto : null;
    if (!photoRef) {
      runwaySurveyToolState.ui.pendingObstacleSaveAfterPhoto = true;
      if (window.M) M.toast({ html: 'Capture obstacle photo before saving angle', classes: 'blue' });
      openRunwaySurveyObstaclePhotoCapture_();
      return;
    }
    runwaySurveyToolState.obstacleAngles50m.push({
      type: type,
      angleDeg: isFinite(a) ? a : null,
      fromThrA50mDeg: isFinite(a) ? a : null,
      fromThrB50mDeg: null,
      fromThreshold: obsCtx.fromThreshold,
      operation: obsCtx.operation,
      checkpointCorner: obsCtx.checkpointCorner,
      checkpointDistanceM: Number(obsCtx.checkpointDistanceM || 50),
      notes: notes,
      photo: {
        name: String(photoRef.name || ''),
        status: String(photoRef.status || ''),
        url: String(photoRef.url || ''),
        fileId: String(photoRef.fileId || ''),
        source: String(photoRef.source || ''),
        queuedAt: String(photoRef.queuedAt || '')
      }
    });
    if (document.getElementById('rwysurvey-obs-a-popup')) document.getElementById('rwysurvey-obs-a-popup').value = '';
    if (document.getElementById('rwysurvey-obs-notes-popup')) document.getElementById('rwysurvey-obs-notes-popup').value = '';
    const checkpointDistance = Number(obsCtx.checkpointDistanceM || 0);
    if (obsCtx.checkpointCorner === 'A' && checkpointDistance >= 250) runwaySurveyToolState.prompts.a300Completed = true;
    else if (obsCtx.checkpointCorner === 'A' && checkpointDistance <= 100) runwaySurveyToolState.prompts.a50Completed = true;
    else if (obsCtx.checkpointCorner === 'C' && checkpointDistance >= 250) runwaySurveyToolState.prompts.c300Completed = true;
    else if (obsCtx.checkpointCorner === 'C' && checkpointDistance <= 100) runwaySurveyToolState.prompts.c50Completed = true;
    _rwySurveyLogEvent_('mark', obsCtx.checkpointCorner + checkpointDistance + ' obstacle angle added: ' + type, { angle: a, operation: obsCtx.operation });
    runwaySurveyToolState.ui.obstaclePhoto = null;
    runwaySurveyToolState.ui.pendingObstacleSaveAfterPhoto = false;
    _rwySurveyObstaclePhotoStatusText_();
    renderRunwaySurveyObstacleList();
    renderRunwaySurveyA50Status();
    if (window.M) M.toast({ html: 'Obstacle angle saved', classes: 'green' });
  }

  function renderRunwaySurveyObstacleList() {
    const summary = document.getElementById('rwysurvey-obstacle-summary');
    const items = runwaySurveyToolState.obstacleAngles50m;
    if (summary) summary.textContent = 'Obstacles: ' + items.length;
  }

  function removeRunwaySurveyObstacle(i) {
    runwaySurveyToolState.obstacleAngles50m.splice(i, 1);
    if (!runwaySurveyToolState.obstacleAngles50m.length && runwaySurveyToolState.prompts) {
      runwaySurveyToolState.prompts.a50Completed = false;
      runwaySurveyToolState.prompts.a300Completed = false;
      runwaySurveyToolState.prompts.c50Completed = false;
      runwaySurveyToolState.prompts.c300Completed = false;
    }
    renderRunwaySurveyObstacleList();
    renderRunwaySurveyA50Status();
  }

  function addRunwaySurveySlopeSegment() {
    const fromThr = String((document.getElementById('rwysurvey-slope-thr') || {}).value || '').trim();
    const d = Number((document.getElementById('rwysurvey-slope-dist') || {}).value);
    const s = Number((document.getElementById('rwysurvey-slope-pct') || {}).value);
    if (!(d > 0) || !isFinite(s)) {
      if (window.M) M.toast({ html: 'Enter valid slope segment distance and slope', classes: 'orange' });
      return;
    }
    runwaySurveyToolState.slopeSegments.push({ fromThreshold: fromThr, distanceM: d, slope: s });
    if (document.getElementById('rwysurvey-slope-dist')) document.getElementById('rwysurvey-slope-dist').value = '';
    if (document.getElementById('rwysurvey-slope-pct')) document.getElementById('rwysurvey-slope-pct').value = '';
    renderRunwaySurveySlopeList();
  }

  function toggleRunwaySurveySlopeCapture() {
    const state = runwaySurveyToolState;
    const gps = state.gps || {};
    if (!state.slopeCapture) {
      state.slopeCapture = { active: false, startIdx: -1, startTs: null, pendingStartDistanceM: 0, pendingDistanceM: 0, pendingFromThreshold: '' };
    }

    const getSlopeStartDistance_ = function() {
      const fix = gps.current;
      const per = state.perimeter || {};
      const marks = Array.isArray(per.cornerMarks) ? per.cornerMarks : [];
      const origin = marks.length ? marks[0] : state.thresholdA;
      if (!fix || !origin) return 0;
      let heading = state.headingDeg || 0;
      if (marks.length >= 2) {
        const a = marks[0], b = marks[1];
        const dy = Number(b.lat || 0) - Number(a.lat || 0);
        const dx = Number(b.lon || 0) - Number(a.lon || 0);
        heading = (Math.atan2(dx, dy) * 180 / Math.PI + 360) % 360;
      }
      return Math.max(0, Math.round(_rwySurveyProjectAlongAxisM_(fix.lat, fix.lon, origin.lat, origin.lon, heading)));
    };

    if (!state.slopeCapture.active) {
      if (!gps.tracking || gps.paused) {
        if (window.M) M.toast({ html: 'Start GPS tracking before slope capture', classes: 'orange' });
        return;
      }
      state.slopeCapture.active = true;
      state.slopeCapture.startIdx = Math.max(0, (gps.points || []).length - 1);
      state.slopeCapture.startTs = Date.now();
      state.slopeCapture.pendingStartDistanceM = getSlopeStartDistance_();
      state.slopeCapture.pendingDistanceM = 0;
      state.slopeCapture.pendingFromThreshold = String(_rwySurveyGetActiveStartThreshold_() || state.rwyIdent || 'RWY');
      renderRunwaySurveySlopeCaptureUi();
      if (window.M) M.toast({ html: 'Slope capture started at ~' + Math.round(Number(state.slopeCapture.pendingStartDistanceM || 0)) + 'm', classes: 'green' });
      return;
    }

    const points = Array.isArray(gps.points) ? gps.points : [];
    const startIdx = Math.max(0, Number(state.slopeCapture.startIdx || 0));
    let distanceM = 0;
    for (let i = startIdx + 1; i < points.length; i++) {
      distanceM += _rwySurveyDistanceMetersBetween_(points[i - 1], points[i]);
    }
    state.slopeCapture.active = false;
    state.slopeCapture.pendingDistanceM = Math.max(0, Math.round(distanceM));
    state.slopeCapture.pendingFromThreshold = String(_rwySurveyGetActiveStartThreshold_() || state.rwyIdent || 'RWY');
    openRunwaySurveySlopePopup();
    renderRunwaySurveySlopeCaptureUi();
  }

  function renderRunwaySurveySlopeCaptureUi() {
    const state = runwaySurveyToolState;
    const cap = state.slopeCapture || { active: false };
    const btn = document.getElementById('rwysurvey-slope-toggle');
    const live = document.getElementById('rwysurvey-slope-live');
    if (btn) {
      btn.textContent = cap.active ? 'END SLOPE' : 'ADD SLOPE';
      btn.className = cap.active ? 'btn deep-orange darken-2' : 'btn green darken-2';
    }
    if (live) {
      live.textContent = cap.active ? 'Slope capture in progress… tap END SLOPE when done' : 'Slope capture idle';
    }
  }

  function openRunwaySurveySlopePopup() {
    const cap = runwaySurveyToolState.slopeCapture || {};
    const popup = document.getElementById('rwysurvey-slope-popup');
    const context = document.getElementById('rwysurvey-slope-popup-context');
    if (context) context.textContent = 'Captured segment: start ~' + Math.round(Number(cap.pendingStartDistanceM || 0)) + 'm, length ' + Math.round(Number(cap.pendingDistanceM || 0)) + 'm';
    const pct = document.getElementById('rwysurvey-slope-popup-pct');
    if (pct) pct.value = '';
    if (popup) popup.style.display = 'block';
  }

  function closeRunwaySurveySlopePopup(cancelled) {
    const popup = document.getElementById('rwysurvey-slope-popup');
    if (popup) popup.style.display = 'none';
    if (cancelled && window.M) M.toast({ html: 'Slope capture cancelled', classes: 'blue-grey' });
  }

  function saveRunwaySurveySlopePopup() {
    const cap = runwaySurveyToolState.slopeCapture || {};
    const s = Number((document.getElementById('rwysurvey-slope-popup-pct') || {}).value);
    if (!isFinite(s)) {
      if (window.M) M.toast({ html: 'Enter valid slope %', classes: 'orange' });
      return;
    }
    const distanceM = Math.max(1, Number(cap.pendingDistanceM || 0));
    const startDistanceM = Math.max(0, Number(cap.pendingStartDistanceM || 0));
    const notes = String((document.getElementById('rwysurvey-slope-popup-notes') || {}).value || '').trim();
    runwaySurveyToolState.slopeSegments.push({
      fromThreshold: String(cap.pendingFromThreshold || runwaySurveyToolState.rwyIdent || 'RWY'),
      startDistanceM: startDistanceM,
      distanceM: distanceM,
      slope: s,
      notes: notes || undefined
    });
    closeRunwaySurveySlopePopup(false);
    renderRunwaySurveySlopeList();
    renderRunwaySurveySlopeCaptureUi();
    if (window.M) M.toast({ html: 'Slope segment saved', classes: 'green' });
  }

  function saveRunwaySurveySlopePopupDeferred() {
    const cap = runwaySurveyToolState.slopeCapture || {};
    const distanceM = Math.max(1, Number(cap.pendingDistanceM || 0));
    const startDistanceM = Math.max(0, Number(cap.pendingStartDistanceM || 0));
    const notes = String((document.getElementById('rwysurvey-slope-popup-notes') || {}).value || '').trim();
    runwaySurveyToolState.slopeSegments.push({
      fromThreshold: String(cap.pendingFromThreshold || runwaySurveyToolState.rwyIdent || 'RWY'),
      startDistanceM: startDistanceM,
      distanceM: distanceM,
      slope: null,
      deferred: true,
      notes: notes || 'Slope to be measured'
    });
    closeRunwaySurveySlopePopup(false);
    renderRunwaySurveySlopeList();
    renderRunwaySurveySlopeCaptureUi();
    if (window.M) M.toast({ html: 'Slope noted — measure later', classes: 'blue-grey' });
  }

  function renderRunwaySurveySlopeList() {
    const el = document.getElementById('rwysurvey-slope-list');
    if (!el) return;
    const items = runwaySurveyToolState.slopeSegments;
    if (!items.length) {
      el.innerHTML = '<div style="font-size:0.8rem; color:#999;">No slope segments yet.</div>';
      return;
    }
    el.innerHTML = items.map(function(seg, i) {
      const isDeferred = seg.deferred || seg.slope === null;
      const slopeStr = isDeferred
        ? '<span style="color:#e65100; font-weight:800;">⏱ TBD</span>' + (seg.notes ? ' — ' + seg.notes : '')
        : (seg.slope >= 0 ? '+' : '') + seg.slope + '%';
      const borderColor = isDeferred ? '#ffe0b2' : '#e1bee7';
      const bgColor = isDeferred ? '#fff8f0' : '#fff';
      return '<div style="display:flex; justify-content:space-between; align-items:center; gap:8px; margin-top:4px; padding:5px 8px; border:1px solid ' + borderColor + '; border-radius:6px; background:' + bgColor + '; font-size:0.82rem;">'
        + '<span><b>' + seg.fromThreshold + '</b> · start ' + Math.round(Number(seg.startDistanceM || 0)) + 'm · len ' + Math.round(seg.distanceM) + 'm · ' + slopeStr + '</span>'
        + '<button onclick="removeRunwaySurveySlopeSegment(' + i + ')" style="border:none; background:none; color:#d32f2f; cursor:pointer;">✕</button></div>';
    }).join('');
  }

  function removeRunwaySurveySlopeSegment(i) {
    runwaySurveyToolState.slopeSegments.splice(i, 1);
    renderRunwaySurveySlopeList();
  }

  function submitRunwaySurveyTool() {
    const state = runwaySurveyToolState;
    if (!state.icao || !state.rwyIdent) {
      if (window.M) M.toast({ html: 'Select airport and runway', classes: 'orange' });
      return;
    }

    const surfaceObserved = String((document.getElementById('rwysurvey-surface') || {}).value || '').trim();
    if (!surfaceObserved) {
      if (window.M) M.toast({ html: 'Surface observed is required', classes: 'orange' });
      return;
    }

    const btn = document.getElementById('rwysurvey-submit-btn');
    if (btn) { btn.disabled = true; btn.textContent = 'SUBMITTING…'; }

    const notes = String((document.getElementById('rwysurvey-notes') || {}).value || '').trim();
    const cutdownRaw = (document.getElementById('rwysurvey-cutdown') || {}).value;
    const cutdownM = cutdownRaw !== '' && !isNaN(Number(cutdownRaw)) ? Number(cutdownRaw) : null;
    const derived = _rwySurveyDerivedDimensions_();
    const marks = (state.perimeter && Array.isArray(state.perimeter.cornerMarks)) ? state.perimeter.cornerMarks : [];
    const submitThresholdA = Object.assign({}, state.thresholdA || {});
    const submitThresholdB = Object.assign({}, state.thresholdB || {});
    const aLatValid = isFinite(Number(submitThresholdA.lat));
    const aLonValid = isFinite(Number(submitThresholdA.lon));
    if ((!aLatValid || !aLonValid) && marks.length) {
      submitThresholdA.ident = state.rwyIdent;
      submitThresholdA.lat = Number(marks[0].lat || 0);
      submitThresholdA.lon = Number(marks[0].lon || 0);
    }
    const effectiveLengthM = Math.max(0, Math.round(derived.lengthM || state.lengthM || state.official.lengthM || 0));
    const hasA = isFinite(Number(submitThresholdA.lat)) && isFinite(Number(submitThresholdA.lon));
    const hasB = isFinite(Number(submitThresholdB.lat)) && isFinite(Number(submitThresholdB.lon));
    const aToBM = (hasA && hasB) ? _rwySurveyDistanceMetersBetween_(submitThresholdA, submitThresholdB) : 0;
    const shouldRebuildB = hasA && effectiveLengthM > 0 && (!hasB || aToBM < Math.max(10, effectiveLengthM * 0.2));
    if (shouldRebuildB) {
      let heading = Number(state.headingDeg || 0);
      if (marks.length >= 2) {
        const a = marks[0], b = marks[1];
        const dy = Number(b.lat || 0) - Number(a.lat || 0);
        const dx = Number(b.lon || 0) - Number(a.lon || 0);
        heading = (Math.atan2(dx, dy) * 180 / Math.PI + 360) % 360;
      }
      const br = _rwySurveyDeg2Rad_(heading);
      const east = Math.sin(br) * effectiveLengthM;
      const north = Math.cos(br) * effectiveLengthM;
      const thrBpt = _rwySurveyMetersToLatLon_(Number(submitThresholdA.lat || 0), Number(submitThresholdA.lon || 0), east, north);
      submitThresholdB.ident = _rwySurveyReciprocalIdent_(state.rwyIdent);
      submitThresholdB.lat = Number(thrBpt.lat || 0);
      submitThresholdB.lon = Number(thrBpt.lon || 0);
    }
    const gps = state.gps || {};
    const payload = {
      icao: state.icao,
      rwyIdent: state.rwyIdent,
      startThresholdIdent: state.startThresholdIdent || state.rwyIdent,
      reciprocalThresholdIdent: state.reciprocalThresholdIdent || _rwySurveyReciprocalIdent_(state.rwyIdent),
      pilotName: String((window.currentBriefingMission && window.currentBriefingMission.pilot) || 'Unknown Pilot'),
      pilotEmail: String((window.currentBriefingMission && window.currentBriefingMission.meta && window.currentBriefingMission.meta.pilotEmail) || ''),
      notes: notes,
      features: state.features,
      photoUploads: Array.isArray(state.photos) ? state.photos.map(function(p) {
        return {
          name: String(p && p.name || ''),
          status: String(p && p.status || ''),
          url: String(p && p.url || ''),
          fileId: String(p && p.fileId || ''),
          source: String(p && p.source || ''),
          queuedAt: String(p && p.queuedAt || '')
        };
      }) : [],
      survey: {
        lengthM: Math.round(derived.lengthM || state.lengthM || state.official.lengthM || 0),
        widthM: Math.round(derived.widthM || state.official.widthM || 0),
        surface: surfaceObserved,
        surfaceObserved: surfaceObserved,
        cutdownAreaM: cutdownM,
        cutdownAreas: { thrA: cutdownM, thrB: cutdownM },
        slopeFromThreshold: String((document.getElementById('rwysurvey-slope-thr') || {}).value || state.rwyIdent),
        features: state.features,
        markers: state.features.map(function(f) {
          return { label: f.name, type: f.type, distanceM: f.distance, fromThreshold: f.fromThreshold, side: f.side, gps: f.gps };
        }),
        obstacles: [],
        obstacleAngles50m: state.obstacleAngles50m,
        slopeSegments: state.slopeSegments,
        perimeterSegments: (state.perimeter && Array.isArray(state.perimeter.segments)) ? state.perimeter.segments : [],
        perimeterSummary: {
          lengthWalkedM: Math.round(derived.lengthWalkedM || 0),
          widthWalkedM: Math.round(derived.widthWalkedM || 0),
          derivedLengthM: Math.round(derived.lengthM || 0),
          derivedWidthM: Math.round(derived.widthM || 0),
          cornerWidthM: Math.round(derived.widthCornerM || 0),
          lateralWidthM: Math.round(derived.widthLateralM || 0),
          lateralWidthSamples: Number(derived.widthLateralSamples || 0),
          closureErrorM: Math.round(derived.closureErrorM || 0),
          lengthSegments: Number(derived.lengthSegments || 0),
          widthSegments: Number(derived.widthSegments || 0)
        },
        debugEvents: (state.debugEvents || []).slice(-40),
        perimeterTrace: gps.points || [],
        widthObservations: state.widthObservations || [],
        photoUploads: Array.isArray(state.photos) ? state.photos.map(function(p) {
          return {
            name: String(p && p.name || ''),
            status: String(p && p.status || ''),
            url: String(p && p.url || ''),
            fileId: String(p && p.fileId || ''),
            source: String(p && p.source || ''),
            queuedAt: String(p && p.queuedAt || '')
          };
        }) : [],
        axis: { headingDeg: state.headingDeg || 0, lengthM: Math.round(derived.lengthM || state.lengthM || 0) },
        thresholds: { a: submitThresholdA, b: submitThresholdB },
        thresholdReference: { start: state.startThresholdIdent || state.rwyIdent, opposite: state.reciprocalThresholdIdent || _rwySurveyReciprocalIdent_(state.rwyIdent) },
        notes: notes,
        gpsSummary: {
          points: Array.isArray(gps.points) ? gps.points.length : 0,
          bestAccuracyM: isFinite(gps.bestAcc) ? gps.bestAcc : null,
          avgAccuracyM: gps.samples ? gps.avgAcc : null,
          maxRateMode: true
        }
      },
      official: {
        lengthM: state.official.lengthM || 0,
        widthM: state.official.widthM || 0,
        surface: String(state.official.surface || ''),
        headingDeg: state.official.headingDeg || 0
      },
      captureSummary: {
        pointCount: Array.isArray(gps.points) ? gps.points.length : 0,
        bestAccuracyM: isFinite(gps.bestAcc) ? gps.bestAcc : null,
        avgAccuracyM: gps.samples ? gps.avgAcc : null,
        lengthWalkedM: Math.round(derived.lengthWalkedM || 0),
        widthWalkedM: Math.round(derived.widthWalkedM || 0)
      },
      deviceInfo: {
        userAgent: String(navigator.userAgent || ''),
        platform: String(navigator.platform || ''),
        submittedAt: new Date().toISOString()
      }
    };

    window.runOrQueueServerAction({
      method: 'submitRunwaySurvey',
      args: [payload],
      label: 'Runway survey ' + state.icao + ' RWY ' + state.rwyIdent
    }, {
      onSuccess: function(resp) {
        if (btn) { btn.disabled = false; btn.textContent = 'SUBMIT FOR SUPERVISOR REVIEW'; }
        if (resp && resp.success) {
          if (window.M) M.toast({ html: 'Runway survey submitted for supervisor review', classes: 'green' });
          closeRunwaySurveyTool();
        } else if (window.M) {
          M.toast({ html: (resp && resp.error) ? resp.error : 'Submit failed', classes: 'red' });
        }
      },
      onQueued: function() {
        if (btn) { btn.disabled = false; btn.textContent = 'SUBMIT FOR SUPERVISOR REVIEW'; }
        if (window.M) M.toast({ html: 'Offline: runway survey queued', classes: 'orange' });
        closeRunwaySurveyTool();
      },
      onFailure: function(err) {
        if (btn) { btn.disabled = false; btn.textContent = 'SUBMIT FOR SUPERVISOR REVIEW'; }
        if (window.M) M.toast({ html: 'Submit failed: ' + (err && err.message ? err.message : String(err)), classes: 'red' });
      }
    });
  }

  function openRunwayDiagramPreview() {
    const state = runwaySurveyToolState;
    if (!state.icao || !state.rwyIdent) {
      if (window.M) M.toast({ html: 'Select airport and runway first', classes: 'orange' });
      return;
    }
    const existing = document.getElementById('rwy-diagram-preview-modal');
    if (existing) existing.remove();

    const derived = _rwySurveyDerivedDimensions_();
    const runwayLengthM = Math.max(Math.round(derived.lengthM || state.lengthM || state.official && state.official.lengthM || 0), 1);
    const runwayWidthM = Math.max(Math.round(derived.widthM || state.official && state.official.widthM || 0), 0);
    const surfaceValue = String((document.getElementById('rwysurvey-surface') || {}).value || state.surfaceObserved || state.official && state.official.surface || '').trim();
    const cutdownValue = String((document.getElementById('rwysurvey-cutdown') || {}).value || '').trim();
    const features = Array.isArray(state.features) ? state.features.slice() : [];
    const obstacles = Array.isArray(state.obstacleAngles50m) ? state.obstacleAngles50m.slice() : [];
    const startThr = _rwySurveyGetActiveStartThreshold_() || state.rwyIdent || 'A';
    const oppThr = _rwySurveyGetOppositeStartThreshold_() || _rwySurveyReciprocalIdent_(startThr);
    const slopeSegments = Array.isArray(state.slopeSegments) && state.slopeSegments.length
      ? state.slopeSegments.slice()
      : [{ startDistanceM: 0, distanceM: runwayLengthM, slope: 0, fromThreshold: startThr }];

    const modal = document.createElement('div');
    modal.id = 'rwy-diagram-preview-modal';
    modal.style.cssText = 'position:fixed; inset:0; z-index:10280; background:rgba(0,0,0,0.68); display:flex; align-items:center; justify-content:center; padding:16px;';

    const content = document.createElement('div');
    content.style.cssText = 'background:#fff; border-radius:12px; width:min(760px, 100%); max-height:92vh; overflow:auto; box-shadow:0 12px 36px rgba(0,0,0,0.35);';

    const header = document.createElement('div');
    header.style.cssText = 'display:flex; justify-content:space-between; align-items:center; padding:12px 14px; background:#0b5394; color:#fff;';
    header.innerHTML = '<div><div style="font-size:1rem; font-weight:900;">RUNWAY PREVIEW</div><div style="font-size:0.78rem; opacity:0.9;">' + state.icao + ' · RWY ' + state.rwyIdent + '</div></div><button onclick="closeRunwayDiagramPreview()" style="border:none; background:rgba(255,255,255,0.18); color:#fff; border-radius:6px; padding:6px 10px; font-size:1rem; cursor:pointer;">✕</button>';

    const body = document.createElement('div');
    body.style.cssText = 'padding:14px; display:grid; gap:10px;';

    const summary = document.createElement('div');
    summary.style.cssText = 'display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:8px;';
    const internalStamp = String(state.internalUpdatedAt || '').trim();
    const internalDateLabel = internalStamp ? internalStamp.slice(0, 16).replace('T', ' ') + 'Z' : 'N/A';
    summary.innerHTML = ''
      + '<div style="border:1px solid #dbe7f3; border-radius:8px; padding:8px; background:#f7fbff; font-size:0.84rem;"><b style="color:#0b5394;">Internal/Surveyed</b><br>' + runwayLengthM + 'm × ' + runwayWidthM + 'm' + (surfaceValue ? ' • ' + surfaceValue : '') + (cutdownValue ? ' • cutdown ' + cutdownValue + 'm' : '') + '<br><span style="font-size:0.76rem; color:#607d8b;">Internal date: ' + internalDateLabel + '</span></div>'
      + '<div style="border:1px solid #e0e0e0; border-radius:8px; padding:8px; background:#fafafa; font-size:0.84rem;"><b style="color:#555;">Official</b><br>' + Math.round(Number(state.official && state.official.lengthM || 0)) + 'm × ' + Math.round(Number(state.official && state.official.widthM || 0)) + 'm' + (state.official && state.official.surface ? ' • ' + state.official.surface : '') + '</div>';

    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('width', '100%');
    svg.setAttribute('height', '360');
    svg.setAttribute('viewBox', '0 0 640 360');
    svg.style.cssText = 'border:1px solid #d7dde3; border-radius:10px; background:#f9fbfd;';

    const makeEl = function(tag, attrs, text) {
      const el = document.createElementNS('http://www.w3.org/2000/svg', tag);
      Object.keys(attrs || {}).forEach(function(k) { el.setAttribute(k, String(attrs[k])); });
      if (text != null) el.textContent = text;
      return el;
    };
    const norm = String(surfaceValue || '').toUpperCase();
    const isPaved = /ASPHALT|PAVED|CONCRETE|ASFALTO|CONCRETO/.test(norm);
    const isGrass = /GRASS|TURF|GRAMA/.test(norm);

    const defs = makeEl('defs', {});
    const grassPattern = makeEl('pattern', { id: 'surveyPreviewGrass', patternUnits: 'userSpaceOnUse', width: 12, height: 12 });
    grassPattern.appendChild(makeEl('rect', { x: 0, y: 0, width: 12, height: 12, fill: '#6d9557' }));
    grassPattern.appendChild(makeEl('path', { d: 'M0,12 L12,0 M-3,9 L3,3 M9,15 L15,9', stroke: '#87b36e', 'stroke-width': 1 }));
    defs.appendChild(grassPattern);
    svg.appendChild(defs);

    const topX = 48;
    const topY = 46;
    const topH = 72;
    const usableW = 540;
    const scale = usableW / runwayLengthM;
    const runwayPx = runwayLengthM * scale;
    const fill = isPaved ? '#5f5f5f' : (isGrass ? 'url(#surveyPreviewGrass)' : '#9c8762');

    svg.appendChild(makeEl('text', { x: topX, y: 22, 'font-size': 12, fill: '#234', 'font-weight': '700' }, 'TOP VIEW'));
    svg.appendChild(makeEl('rect', { x: topX, y: topY, width: runwayPx, height: topH, fill: fill, stroke: '#263238', 'stroke-width': 2, rx: 4, ry: 4 }));
    if (isPaved) {
      svg.appendChild(makeEl('line', { x1: topX, y1: topY + (topH / 2), x2: topX + runwayPx, y2: topY + (topH / 2), stroke: '#f3f3f3', 'stroke-width': 3, 'stroke-dasharray': '18,10' }));
    }
    svg.appendChild(makeEl('text', { x: topX + 10, y: topY + topH - 12, 'font-size': 18, fill: '#fff', 'font-weight': '800' }, String(state.rwyIdent || 'RWY')));
    svg.appendChild(makeEl('text', { x: topX + runwayPx - 10, y: topY + topH - 12, 'font-size': 18, fill: '#fff', 'font-weight': '800', 'text-anchor': 'end' }, String(_rwySurveyReciprocalIdent_(state.rwyIdent) || '')));
    svg.appendChild(makeEl('text', { x: topX + (runwayPx / 2), y: topY + topH + 22, 'font-size': 12, fill: '#0b5394', 'font-weight': '700', 'text-anchor': 'middle' }, runwayLengthM + 'm × ' + runwayWidthM + 'm' + (surfaceValue ? ' • ' + surfaceValue : '')));

    const iconForFeature = function(name) {
      const s = String(name || '').toUpperCase();
      if (s.indexOf('TREE') >= 0) return '🌳';
      if (s.indexOf('SCHOOL') >= 0) return '🏫';
      if (s.indexOf('HOUSE') >= 0 || s.indexOf('CHURCH') >= 0) return '🏠';
      if (s.indexOf('WINDSOCK') >= 0) return '🎏';
      if (s.indexOf('TOWER') >= 0) return '🗼';
      if (s.indexOf('POWER') >= 0) return '⚡';
      return '📍';
    };

    const iconForObstacle = function(type) {
      const t = String(type || '').toLowerCase();
      if (t.indexOf('tree') >= 0) return '🌳';
      if (t.indexOf('building') >= 0 || t.indexOf('house') >= 0) return '🏢';
      if (t.indexOf('hill') >= 0) return '⛰';
      if (t.indexOf('rock') >= 0) return '🪨';
      if (t.indexOf('power') >= 0) return '⚡';
      return '📍';
    };

    features.sort(function(a, b) { return Number(a.distance || 0) - Number(b.distance || 0); }).forEach(function(feat, idx) {
      const distM = Math.max(0, Math.min(runwayLengthM, Number(feat && feat.distance || 0)));
      const x = topX + (distM * scale);
      const side = String(feat && feat.side || 'right').toLowerCase() === 'left' ? 'left' : 'right';
      const iconY = side === 'left' ? topY - 14 : topY + topH + 26;
      const connectorY = side === 'left' ? topY : topY + topH;
      svg.appendChild(makeEl('line', { x1: x, y1: connectorY, x2: x, y2: side === 'left' ? iconY + 6 : iconY - 10, stroke: '#fb8c00', 'stroke-width': 1.5 }));
      svg.appendChild(makeEl('text', { x: x, y: iconY, 'font-size': 15, 'text-anchor': 'middle' }, iconForFeature(feat && (feat.name || feat.type))));
      svg.appendChild(makeEl('text', { x: x, y: side === 'left' ? iconY - 8 : iconY + 12, 'font-size': 9, fill: '#455a64', 'text-anchor': 'middle' }, Math.round(distM) + 'm'));
    });

    // Obstacles are shown only in side view so the top view stays uncluttered.

    const sideX = 48;
    const sideY = 212;
    const sideH = 80;
    const profile = slopeSegments.map(function(seg) {
      const segDist = Math.max(0, Number(seg && (seg.distanceM != null ? seg.distanceM : seg.distance) || 0));
      const segFromThreshold = String(seg && seg.fromThreshold || startThr).trim().toUpperCase();
      const rawStart = Number(seg && seg.startDistanceM);
      let start = isFinite(rawStart) ? rawStart : 0;
      if (!isFinite(rawStart)) {
        start = segFromThreshold === oppThr ? Math.max(0, runwayLengthM - segDist) : 0;
      }
      start = Math.max(0, Math.min(runwayLengthM, start));
      const distance = segDist;
      const end = Math.max(start, Math.min(runwayLengthM, start + distance));
      return {
        startDistanceM: start,
        endDistanceM: end,
        distance: Math.max(0, end - start),
        slope: Number(seg && seg.slope || 0) || 0,
        fromThreshold: String(seg && seg.fromThreshold || startThr)
      };
    }).filter(function(seg) { return seg.distance > 0; }).sort(function(a, b) { return a.startDistanceM - b.startDistanceM; });

    const normalizedProfile = profile.length ? profile : [{ startDistanceM: 0, endDistanceM: runwayLengthM, distance: runwayLengthM, slope: 0, fromThreshold: startThr }];
    let currentElevation = 0;
    let cursorDist = 0;
    const profilePts = [{ dist: 0, elev: 0 }];

    normalizedProfile.forEach(function(seg) {
      if (seg.startDistanceM > cursorDist) {
        profilePts.push({ dist: seg.startDistanceM, elev: currentElevation });
      }
      currentElevation += (seg.distance * seg.slope / 100);
      profilePts.push({ dist: seg.endDistanceM, elev: currentElevation });
      cursorDist = seg.endDistanceM;
    });
    if (cursorDist < runwayLengthM) {
      profilePts.push({ dist: runwayLengthM, elev: currentElevation });
    }

    const elevations = profilePts.map(function(p) { return p.elev; });
    const minEl = Math.min.apply(null, elevations);
    const maxEl = Math.max.apply(null, elevations);
    const spanEl = Math.max(0.5, maxEl - minEl);
    const elScale = Math.max(0.8, Math.min(sideH / spanEl, 4));
    const usedHeight = spanEl * elScale;
    const yOffset = (sideH - usedHeight) / 2;
    svg.appendChild(makeEl('text', { x: sideX, y: sideY - 8, 'font-size': 12, fill: '#234', 'font-weight': '700' }, 'SIDE VIEW'));
    svg.appendChild(makeEl('rect', { x: sideX, y: sideY, width: runwayPx, height: sideH, fill: '#f2f5f8', stroke: '#d7dde3' }));
    const points = profilePts.map(function(p) {
      const px = sideX + (p.dist * scale);
      const py = sideY + yOffset + usedHeight - ((p.elev - minEl) * elScale);
      return px + ',' + py;
    });
    svg.appendChild(makeEl('polyline', { points: points.join(' '), fill: 'none', stroke: '#2e7d32', 'stroke-width': 3 }));

    const yForDist = function(distA) {
      const d = Math.max(0, Math.min(runwayLengthM, Number(distA || 0)));
      for (let i = 1; i < profilePts.length; i++) {
        if (d <= profilePts[i].dist) {
          const a = profilePts[i - 1];
          const b = profilePts[i];
          const span = Math.max(0.0001, Number(b.dist - a.dist));
          const t = (d - a.dist) / span;
          const elev = a.elev + (b.elev - a.elev) * t;
          return sideY + yOffset + usedHeight - ((elev - minEl) * elScale);
        }
      }
      const last = profilePts[profilePts.length - 1];
      return sideY + yOffset + usedHeight - ((last.elev - minEl) * elScale);
    };

    const endAnchorX = { left: sideX - 34, right: sideX + runwayPx + 34 };
    const endStacks = { left: 0, right: 0 };

    const drawObstacleShape_ = function(type, anchorX, anchorY, sideDir, container) {
      const t = String(type || '').toLowerCase();
      const c = container || svg;
      if (t.indexOf('tree') >= 0) {
        c.appendChild(makeEl('text', { x: anchorX - (sideDir * 10), y: anchorY + 4, 'font-size': 12, 'text-anchor': 'middle' }, '🌳'));
        c.appendChild(makeEl('text', { x: anchorX, y: anchorY + 4, 'font-size': 12, 'text-anchor': 'middle' }, '🌳'));
        c.appendChild(makeEl('text', { x: anchorX + (sideDir * 10), y: anchorY + 4, 'font-size': 12, 'text-anchor': 'middle' }, '🌳'));
        return;
      }
      if (t.indexOf('hill') >= 0) {
        c.appendChild(makeEl('path', {
          d: 'M' + (anchorX - 18) + ',' + (anchorY + 10)
            + ' L' + (anchorX - 6) + ',' + (anchorY - 4)
            + ' L' + anchorX + ',' + (anchorY + 2)
            + ' L' + (anchorX + 10) + ',' + (anchorY - 8)
            + ' L' + (anchorX + 22) + ',' + (anchorY + 10),
          fill: 'none',
          stroke: '#6d4c41',
          'stroke-width': 2.4
        }));
        return;
      }
      c.appendChild(makeEl('text', { x: anchorX, y: anchorY + 4, 'font-size': 13, 'text-anchor': 'middle' }, iconForObstacle(t)));
    };

    obstacles.forEach(function(obs, obsIdx) {
      const obsGroup = makeEl('g', { id: 'pilotObsGrp_' + obsIdx });
      const cornerRaw = String(obs && obs.checkpointCorner || 'A').toUpperCase();
      const fromThreshold = String(obs && obs.fromThreshold || '').trim().toUpperCase();
      const corner = fromThreshold === oppThr ? 'C' : (fromThreshold === startThr ? 'A' : cornerRaw);
      const distM = Math.max(0, Math.min(runwayLengthM, Number(obs && (obs.checkpointDistanceM != null ? obs.checkpointDistanceM : obs.distanceM) || 0)));
      const fromA = corner === 'C' ? (runwayLengthM - distM) : distM;
      const baseX = sideX + (fromA * scale);
      const baseY = yForDist(fromA);
      const operation = String(obs && obs.operation || '').toLowerCase() || (Number(distM) >= 300 ? 'takeoff' : 'landing');
      const dir = corner === 'C'
        ? (operation === 'landing' ? 1 : -1)
        : (operation === 'landing' ? -1 : 1);
      const sideKey = dir > 0 ? 'right' : 'left';
      const rawAngle = Math.max(0, Number(obs && (obs.angleDeg != null ? obs.angleDeg : (obs.fromThrA50mDeg != null ? obs.fromThrA50mDeg : obs.fromThrB50mDeg)) || 0));
      const angle = Math.min(45, rawAngle);
      const rad = angle * Math.PI / 180;
      const stackIndex = endStacks[sideKey]++;
      const anchorX = endAnchorX[sideKey] + (dir * Math.min(18, stackIndex * 4));
      const anchorY = sideY + 18 + (stackIndex * 16);
      const tipY = Math.min(baseY - 8, anchorY + (Math.abs(anchorX - baseX) * Math.tan(rad)));
      const thr = String(obs && obs.fromThreshold || _rwySurveyThresholdForCorner_(corner) || (corner === 'C' ? oppThr : startThr));

      obsGroup.appendChild(makeEl('circle', { cx: baseX, cy: baseY, r: 2.2, fill: '#37474f' }));
      obsGroup.appendChild(makeEl('line', { x1: baseX, y1: baseY, x2: anchorX, y2: tipY, stroke: '#ef6c00', 'stroke-width': 2.2 }));
      obsGroup.appendChild(makeEl('circle', { cx: anchorX, cy: tipY, r: 1.7, fill: '#ef6c00' }));
      drawObstacleShape_(obs && obs.type, anchorX, anchorY, dir, obsGroup);
      obsGroup.appendChild(makeEl('text', {
        x: anchorX,
        y: anchorY + 18,
        'font-size': 8,
        fill: '#37474f',
        'text-anchor': 'middle',
        'font-weight': '700'
      }, Math.round(distM) + 'm • ' + rawAngle.toFixed(1) + '°'));
      obsGroup.appendChild(makeEl('text', {
        x: anchorX,
        y: anchorY + 28,
        'font-size': 7.5,
        fill: '#546e7a',
        'text-anchor': 'middle'
      }, thr + ' ' + operation));
      svg.appendChild(obsGroup);
    });

    normalizedProfile.forEach(function(seg) {
      const midX = sideX + (((seg.startDistanceM + seg.endDistanceM) / 2) * scale);
      svg.appendChild(makeEl('text', { x: midX, y: sideY + sideH + 14, 'font-size': 9, fill: '#4b4b4b', 'text-anchor': 'middle' }, (seg.slope >= 0 ? '+' : '') + seg.slope.toFixed(1) + '% / ' + Math.round(seg.distance) + 'm'));
      svg.appendChild(makeEl('text', { x: midX, y: sideY + sideH + 25, 'font-size': 8, fill: '#6b6b6b', 'text-anchor': 'middle' }, 'start ' + Math.round(seg.startDistanceM) + 'm'));
    });

    body.appendChild(summary);

    // Build flex row: SVG + obstacle checklist panel
    const normThrPilot = function(raw) {
      const txt = String(raw || '').trim().toUpperCase().replace(/^RWY\s*/, '');
      const m = txt.match(/^(\d{1,2})([LCR])?$/);
      if (!m) return txt;
      const n = parseInt(m[1], 10);
      if (!isFinite(n) || n < 1 || n > 36) return txt;
      return String(n).padStart(2, '0') + (m[2] || '');
    };
    const diagramRow = document.createElement('div');
    diagramRow.style.cssText = 'display:flex; gap:10px; align-items:flex-start;';
    diagramRow.appendChild(svg);
    if (obstacles.length) {
      const obsPanel = document.createElement('div');
      obsPanel.style.cssText = 'border:1px solid #e0e6eb; border-radius:8px; padding:8px 10px; background:#fafefe; min-width:150px; max-width:185px; flex-shrink:0; font-size:0.8rem;';
      obsPanel.innerHTML = '<div style="font-weight:800; color:#0b5394; margin-bottom:6px; font-size:0.78rem;">Obstacle Angles</div>'
        + obstacles.map(function(obs, idx) {
          const dist = Math.round(Number(obs && obs.checkpointDistanceM || 0));
          const thr = normThrPilot(obs && obs.fromThreshold || '') || startThr;
          const operation = String(obs && obs.operation || '').toLowerCase() || (dist >= 300 ? 'takeoff' : 'landing');
          const label = 'RWY ' + thr + ' \u00b7 ' + dist + 'm \u00b7 ' + operation;
          return '<label style="display:flex;align-items:center;gap:4px;margin-bottom:4px;cursor:pointer;white-space:nowrap;">'
            + '<input type="checkbox" checked onchange="_pilotObsToggle_(this,' + idx + ')">'
            + '<span>' + label + '</span></label>';
        }).join('');
      diagramRow.appendChild(obsPanel);
    }
    body.appendChild(diagramRow);

    const footer = document.createElement('div');
    footer.style.cssText = 'display:grid; gap:6px; font-size:0.82rem; color:#37474f;';
    footer.innerHTML = ''
      + '<div style="background:#eef6ff; border-left:4px solid #1976d2; padding:8px 10px; border-radius:6px;">Features: ' + (features.length || 0) + ' • Obstacles: ' + (obstacles.length || 0) + ' • Slope segments: ' + (slopeSegments.length || 0) + '</div>';
    body.appendChild(footer);

    content.appendChild(header);
    content.appendChild(body);
    modal.appendChild(content);
    document.body.appendChild(modal);
    modal.onclick = function(e) { if (e.target === modal) closeRunwayDiagramPreview(); };
  }

  function closeRunwayDiagramPreview() {
    const modal = document.getElementById('rwy-diagram-preview-modal');
    if (modal) modal.remove();
  }

  window._pilotObsToggle_ = function(cb, idx) {
    var g = document.getElementById('pilotObsGrp_' + idx);
    if (g) g.style.display = cb.checked ? '' : 'none';
  };

  
