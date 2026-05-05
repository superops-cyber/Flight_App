
  try { if (window.__pilotDiag) window.__pilotDiag.mainScriptReady = true; } catch (_e) {}

    let currentTab = 1;
    let activeMission = null;
    let appData = {};
    let _activeLessonContext = null;
    const PILOT_TAB_VERSIONS = Object.freeze({
      1: Object.freeze({ shortLabel: 'Gnd', fullLabel: 'Ground', version: 105 }),
      2: Object.freeze({ shortLabel: 'Brief', fullLabel: 'Briefing', version: 180 }),
      3: Object.freeze({ shortLabel: 'W&B', fullLabel: 'Weight & Balance', version: 103 }),
      4: Object.freeze({ shortLabel: 'Perf', fullLabel: 'Performance', version: 104 }),
      5: Object.freeze({ shortLabel: 'Release', fullLabel: 'Release', version: 105 }),
      6: Object.freeze({ shortLabel: 'Enroute', fullLabel: 'Enroute', version: 130 }),
      7: Object.freeze({ shortLabel: 'Arrv', fullLabel: 'Arrival', version: 107 }),
      8: Object.freeze({ shortLabel: 'Log', fullLabel: 'Debrief Log', version: 123 })
    });
    const OFFLINE_CACHE_KEYS = {
      DROPDOWN_DATA: 'mba_cache_dropdown_data_v1',
      SCHEDULED_MISSIONS: 'mba_cache_scheduled_missions_v1',
      PREFETCH_META: 'mba_cache_prefetch_meta_v1',
      OUTBOX: 'mba_outbox_v1',
      DUTY_PROMPT_STATE: 'mba_duty_prompt_state_v1',
      AIRCRAFT_DOCS: 'mba_cache_aircraft_docs_v1'
    };
    const NEW_FLIGHT_URL = '';
    const WEB_APP_URL = <?!= JSON.stringify(typeof webAppUrl === 'string' ? webAppUrl : '') ?>;
    let _pilotClockTimer = null;

    function _pilotClockText_(timeZone) {
      return new Intl.DateTimeFormat('en-GB', {
        timeZone: timeZone,
        hour12: false,
        hour: '2-digit',
        minute: '2-digit'
      }).format(new Date());
    }

    function _pilotMonthAbbrev_(monthIndex) {
      const months = ['Jan.', 'Feb.', 'Mar.', 'Apr.', 'May.', 'Jun.', 'Jul.', 'Aug.', 'Sep.', 'Oct.', 'Nov.', 'Dec.'];
      return months[Math.max(0, Math.min(11, Number(monthIndex) || 0))];
    }

    function _pilotFormatDateMonDayYear_(value, fallback) {
      const s = value instanceof Date ? value : String(value || '').trim();
      // Parse date-only strings as local time (not UTC) to avoid timezone off-by-one
      const date = (typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s)) ? new Date(s + 'T00:00:00') : new Date(s);
      if (isNaN(date.getTime())) return fallback || '';
      const month = _pilotMonthAbbrev_(date.getMonth());
      const day = String(date.getDate());
      const year = String(date.getFullYear());
      return `${month} ${day}, ${year}`;
    }

    function _pilotFormatBsbDateTime_(value, fallback) {
      const date = new Date(value || '');
      if (isNaN(date.getTime())) return fallback || '';
      const datePart = _pilotFormatDateMonDayYear_(date, '');
      const timePart = new Intl.DateTimeFormat('en-GB', {
        timeZone: 'America/Sao_Paulo',
        hour12: false,
        hour: '2-digit',
        minute: '2-digit'
      }).format(date);
      return `${datePart} ${timePart} BSB`;
    }

    function _renderPilotDispatchClocks_() {
      const z = document.getElementById('pilot-time-z');
      const mao = document.getElementById('pilot-time-mao');
      const bsb = document.getElementById('pilot-time-bsb');
      if (!z || !mao || !bsb) return;
      z.textContent = _pilotClockText_('UTC');
      mao.textContent = _pilotClockText_('America/Manaus');
      bsb.textContent = _pilotClockText_('America/Sao_Paulo');
    }

    function initPilotDispatchClocks_() {
      _renderPilotDispatchClocks_();
      if (_pilotClockTimer) clearInterval(_pilotClockTimer);
      _pilotClockTimer = setInterval(_renderPilotDispatchClocks_, 15000);
    }

    function _pilotNormalizeScriptUrl_(raw) {
      try {
        const u = new URL(String(raw || ''));
        if (!/script\.google\.com$/i.test(u.hostname)) return '';
        if (!/\/s\//i.test(u.pathname)) return '';
        if (/\/dev\/?$/i.test(u.pathname)) {
          u.pathname = u.pathname.replace(/\/dev\/?$/i, '/exec');
        }
        if (!/\/(exec|dev)\/?$/i.test(u.pathname)) return '';
        u.search = '';
        u.hash = '';
        return u.toString();
      } catch (e) {
        return '';
      }
    }

    function _lessonSafeParseJson_(raw, fallback) {
      try { return raw ? JSON.parse(raw) : fallback; } catch (e) { return fallback; }
    }

    function _buildLessonContextFromMission_(mission) {
      var training = mission && mission.meta ? mission.meta.training : null;
      var code = String(training && training.code || '').trim();
      if (!code) return null;
      var syllabus = Array.isArray(appData && appData.syllabus) ? appData.syllabus : [];
      var row = syllabus.find(function(s) { return String(s && s.code || '') === code; }) || null;
      var plan = _lessonSafeParseJson_(row && row.lessonPlanJson, null);
      var stops = [];
      if (plan && Array.isArray(plan.stops)) {
        stops = plan.stops;
      } else if (row && row.plannedStopsJson) {
        stops = _lessonSafeParseJson_(row.plannedStopsJson, []);
      } else {
        stops = String(training && training.route || '').split(',').map(function(s) {
          var loc = String(s || '').trim().toUpperCase();
          return loc ? { location: loc, landings: 0, touchAndGos: 0 } : null;
        }).filter(Boolean);
      }
      var maneuvers = [];
      if (plan && Array.isArray(plan.maneuvers)) maneuvers = plan.maneuvers.slice();
      else if (row && row.maneuversJson) {
        var mv = _lessonSafeParseJson_(row.maneuversJson, {});
        maneuvers = Array.isArray(mv && mv.maneuvers) ? mv.maneuvers.slice() : [];
      }
      return {
        code: code,
        title: String((row && row.description) || (training && training.description) || code),
        category: String((plan && plan.category) || ''),
        routeCheckPrompt: String((plan && plan.routeCheckPrompt) || ''),
        runwayCheckLocation: String((plan && plan.runwayCheckLocation) || ''),
        externalLessonName: String((plan && plan.externalLessonName) || ''),
        stops: stops,
        maneuvers: maneuvers,
        mission: mission || {}
      };
    }

    function renderLessonLinkBar_() {
      var bar = document.getElementById('lesson-link-bar');
      var codeEl = document.getElementById('lesson-link-code');
      if (!bar || !codeEl) return;
      if (!_activeLessonContext || !_activeLessonContext.code) {
        bar.style.display = 'none';
        codeEl.textContent = '---';
        return;
      }
      codeEl.textContent = _activeLessonContext.code;
      bar.style.display = 'flex';
    }

    function loadActiveLessonContext_() {
      if (!activeMission) {
        _activeLessonContext = null;
        renderLessonLinkBar_();
        return;
      }
      fetchMissionDetails(activeMission, function(mission) {
        _activeLessonContext = _buildLessonContextFromMission_(mission);
        renderLessonLinkBar_();
      }, function() {
        _activeLessonContext = null;
        renderLessonLinkBar_();
      });
    }

    function _lessonRatingSelectHtml_(name) {
      return '<select data-lesson-score="' + name + '" class="browser-default" style="height:30px; font-size:0.78rem;">'
        + '<option value="">-</option>'
        + '<option value="1">1</option><option value="2">2</option><option value="3">3</option><option value="4">4</option>'
        + '</select>';
    }

    function lessonStorageKey_() {
      var ctx = _activeLessonContext || {};
      var missionId = String(ctx.mission && ctx.mission.id || activeMission || '').trim();
      var code = String(ctx.code || '').trim();
      if (!missionId || !code) return '';
      return 'mba_lesson_eval_' + missionId + '_' + code;
    }

    function lessonCollectEvaluation_() {
      var key = lessonStorageKey_();
      if (!key) return;
      var payload = { scores: {}, notes: {} };
      document.querySelectorAll('[data-lesson-score]').forEach(function(sel) {
        var k = String(sel.getAttribute('data-lesson-score') || '');
        payload.scores[k] = String(sel.value || '');
      });
      document.querySelectorAll('.lesson-note-list').forEach(function(wrap, idx) {
        var items = [];
        wrap.querySelectorAll('[data-lesson-note-item]').forEach(function(div) {
          items.push(String(div.textContent || '').replace(/^•\s*/, '').trim());
        });
        payload.notes[String(idx)] = items;
      });
      try { localStorage.setItem(key, JSON.stringify(payload)); } catch (e) {}
    }

    function lessonRestoreEvaluation_() {
      var key = lessonStorageKey_();
      if (!key) return;
      var raw = '';
      try { raw = localStorage.getItem(key) || ''; } catch (e) { raw = ''; }
      if (!raw) return;
      var parsed = _lessonSafeParseJson_(raw, null);
      if (!parsed || typeof parsed !== 'object') return;
      var scores = parsed.scores || {};
      var notes = parsed.notes || {};
      document.querySelectorAll('[data-lesson-score]').forEach(function(sel) {
        var k = String(sel.getAttribute('data-lesson-score') || '');
        if (Object.prototype.hasOwnProperty.call(scores, k)) sel.value = String(scores[k] || '');
      });
      document.querySelectorAll('.lesson-note-list').forEach(function(wrap, idx) {
        var keyIdx = String(idx);
        var list = Array.isArray(notes[keyIdx]) ? notes[keyIdx] : [];
        wrap.innerHTML = '';
        list.forEach(function(text) {
          var div = document.createElement('div');
          div.style.fontSize = '0.74rem';
          div.style.marginTop = '4px';
          div.style.color = '#455a64';
          div.textContent = '• ' + String(text || '');
          div.setAttribute('data-lesson-note-item', '1');
          wrap.appendChild(div);
        });
      });
    }

    function _lessonRenderGradeRows_(items, kind) {
      if (!items || !items.length) return '<div style="font-size:0.82rem; color:#78909c; padding:6px 0;">Sem itens.</div>';
      return '<table style="width:100%; border-collapse:collapse; font-size:0.79rem;">'
        + '<thead><tr><th style="border:1px solid #e5edf5; background:#eff5fb; padding:6px; text-align:left;">#</th><th style="border:1px solid #e5edf5; background:#eff5fb; padding:6px; text-align:left;">' + (kind === 'landing' ? 'Ponto/Pousos' : 'Manobra') + '</th><th style="border:1px solid #e5edf5; background:#eff5fb; padding:6px; text-align:left;">Nota (1-4)</th><th style="border:1px solid #e5edf5; background:#eff5fb; padding:6px; text-align:left;">Notas</th></tr></thead>'
        + '<tbody>' + items.map(function(item, idx) {
          var label = kind === 'landing'
            ? (String(item.location || '') + ' | Pousos: ' + Number(item.landings || 0) + ' | T&G: ' + Number(item.touchAndGos || 0))
            : String(item || '');
          return '<tr>'
            + '<td style="border:1px solid #e5edf5; padding:6px;">' + (idx + 1) + '</td>'
            + '<td style="border:1px solid #e5edf5; padding:6px;">' + label.replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</td>'
            + '<td style="border:1px solid #e5edf5; padding:6px; min-width:92px;">' + _lessonRatingSelectHtml_(kind + '-' + idx) + '</td>'
            + '<td style="border:1px solid #e5edf5; padding:6px;"><button type="button" class="btn-flat" style="font-size:0.72rem; font-weight:800; padding:0 6px;" onclick="lessonAddInlineNote_(this)">+ NOTA</button><div class="lesson-note-list"></div></td>'
            + '</tr>';
        }).join('') + '</tbody></table>';
    }

    function openLessonPlanModal_() {
      if (!_activeLessonContext) {
        if (window.M) M.toast({ html: 'Nenhum plano de aula ativo nesta missão.', classes: 'orange' });
        return;
      }
      var ctx = _activeLessonContext;
      var modal = document.getElementById('lesson-plan-modal');
      var sub = document.getElementById('lesson-plan-subtitle');
      var content = document.getElementById('lesson-plan-content');
      if (!modal || !sub || !content) return;

      var mission = ctx.mission || {};
      var meta = mission.meta || {};
      var firstLeg = (mission.legs && mission.legs[0]) || {};
      sub.textContent = String(ctx.code || '') + ' • ' + String(ctx.title || 'Plano de Aula');
      content.innerHTML = ''
        + '<div style="font-size:0.8rem; line-height:1.6; color:#37474f; margin-bottom:10px;">'
        + '<b>Data:</b> ' + String(meta.date || mission.date || '') + ' &nbsp; <b>Matrícula:</b> ' + String(meta.acft || mission.acft || '') + ' &nbsp; <b>Piloto:</b> ' + String(meta.pilot || mission.pilot || '') + ' &nbsp; <b>Instrutor:</b> ' + String(meta.copilot || '')
        + '<br><b>Rota:</b> ' + String(firstLeg.route || (ctx.stops || []).map(function(s){ return s.location; }).join(' -> '))
        + (ctx.externalLessonName ? ('<br><b>Ref. Externa:</b> ' + String(ctx.externalLessonName)) : '')
        + (ctx.routeCheckPrompt ? ('<br><b>Rota Check:</b> ' + String(ctx.routeCheckPrompt)) : '')
        + (ctx.runwayCheckLocation ? ('<br><b>Pista Check:</b> ' + String(ctx.runwayCheckLocation)) : '')
        + '</div>'
        + '<h6 style="margin:8px 0 6px 0; color:#0b5394; font-weight:900;">Avaliação de Pousos</h6>'
        + _lessonRenderGradeRows_(ctx.stops || [], 'landing')
        + '<h6 style="margin:12px 0 6px 0; color:#0b5394; font-weight:900;">Avaliação de Manobras</h6>'
        + _lessonRenderGradeRows_(ctx.maneuvers || [], 'maneuver');

      content.querySelectorAll('[data-lesson-score]').forEach(function(sel) {
        sel.addEventListener('change', lessonCollectEvaluation_);
      });
      lessonRestoreEvaluation_();

      modal.style.display = 'flex';
    }

    function closeLessonPlanModal_() {
      var modal = document.getElementById('lesson-plan-modal');
      if (modal) modal.style.display = 'none';
    }

    function lessonAddInlineNote_(btn) {
      var wrap = btn && btn.parentElement ? btn.parentElement.querySelector('.lesson-note-list') : null;
      if (!wrap) return;
      var text = prompt('Digite a observação:');
      if (text == null) return;
      var safe = String(text || '').trim();
      if (!safe) return;
      var div = document.createElement('div');
      div.style.fontSize = '0.74rem';
      div.style.marginTop = '4px';
      div.style.color = '#455a64';
      div.textContent = '• ' + safe;
      div.setAttribute('data-lesson-note-item', '1');
      wrap.appendChild(div);
      lessonCollectEvaluation_();
    }

    function _pilotBestBaseUrl_() {
      return _pilotNormalizeScriptUrl_(window.location.href)
        || _pilotNormalizeScriptUrl_(document.referrer)
        || _pilotNormalizeScriptUrl_(WEB_APP_URL)
        || '';
    }

    function _getPilotTabVersionInfo_(tabNum) {
      return PILOT_TAB_VERSIONS[Number(tabNum)] || null;
    }

    function renderPilotTabVersions_() {
      Object.keys(PILOT_TAB_VERSIONS).forEach(function(tabKey) {
        const tabNum = Number(tabKey);
        const info = _getPilotTabVersionInfo_(tabNum);
        const tabEl = document.getElementById('tab' + tabNum);
        if (info && tabEl) {
          let badge = tabEl.querySelector('.app-tab-version-badge');
          if (!badge) {
            badge = document.createElement('div');
            badge.className = 'app-tab-version-badge';
            tabEl.appendChild(badge);
          }
          badge.textContent = 'Tab ' + tabNum + ' v' + String(info.version);
          badge.setAttribute('aria-label', 'Tab ' + tabNum + ' ' + info.fullLabel + ' version ' + info.version);
          badge.title = info.fullLabel + ' version ' + info.version;
        }

        const stepEl = document.getElementById('step' + tabNum);
        if (info && stepEl) {
          stepEl.setAttribute('title', 'Tab ' + tabNum + ' version ' + info.version);
          stepEl.setAttribute('aria-label', 'Tab ' + tabNum + ' ' + info.shortLabel + ' version ' + info.version);
        }
      });

      const listEl = document.getElementById('tab1-version-list');
      if (!listEl) return;
      listEl.innerHTML = Object.keys(PILOT_TAB_VERSIONS).map(function(tabKey) {
        const tabNum = Number(tabKey);
        const info = _getPilotTabVersionInfo_(tabNum);
        return ''
          + '<div class="tab1-version-row">'
          + '<span class="tab1-version-label">Tab ' + tabNum + ' ' + info.shortLabel + '</span>'
          + '<span class="tab1-version-value">v' + String(info.version) + '</span>'
          + '</div>';
      }).join('');
    }

    window.getPilotTabVersion = function(tabNum) {
      const info = _getPilotTabVersionInfo_(tabNum);
      return info ? info.version : null;
    };

    window.getPilotTabVersions = function() {
      return JSON.parse(JSON.stringify(PILOT_TAB_VERSIONS));
    };

    function _getDispatchPortalUrl_() {
      try {
        const base = _pilotBestBaseUrl_() || window.location.href;
        const url = new URL(base);
        url.searchParams.delete('createOAuthDialog');
        url.searchParams.set('view', 'portal');
        url.searchParams.set('src', 'pilot-app');
        url.searchParams.set('t', String(Date.now()));
        url.hash = 'view-dispatch';
        return url.toString();
      } catch (e) {
        const base = String(_pilotBestBaseUrl_() || '').trim();
        return base ? (base + '?view=portal&src=pilot-app&t=' + Date.now() + '#view-dispatch') : '?view=portal&src=pilot-app&t=' + Date.now() + '#view-dispatch';
      }
    }
    let hasPrefetchedMissionsThisSession = false;
    let missionPrefetchInProgress = false;
    let missionPrefetchProgress = { total: 0, progress: 0 };
    let _lastMissionListServerFetchTs = 0;
    let outboxSyncInProgress = false;
    const OUTBOX_MAX_ATTEMPTS = 5;
    const OFFLINE_CACHE_STALE_MS = 24 * 60 * 60 * 1000;
    const OFFLINE_ENVELOPE_STALE_MS = 30 * 24 * 60 * 60 * 1000;
    const OFFLINE_CACHE_MAX_MISSIONS = 60;
    let _dutyPromptTimer = null;
    let _dutyPromptBusy = false;
    let _dutyGeoVerifyMap = null;
    let _dutyGeoVerifyCurrentMarker = null;
    let _dutyGeoVerifyTargetMarker = null;
    let _dutyGeoVerifyCircle = null;
    const FLIGHTAPP_DIALOG_STATE = {
      resolver: null,
      mode: 'alert'
    };

    function _flightAppDialogElements() {
      return {
        overlay: document.getElementById('flightapp-dialog-overlay'),
        title: document.getElementById('flightapp-dialog-title'),
        message: document.getElementById('flightapp-dialog-message'),
        input: document.getElementById('flightapp-dialog-input'),
        ok: document.getElementById('flightapp-dialog-ok'),
        cancel: document.getElementById('flightapp-dialog-cancel')
      };
    }

    function _flightAppDialogClose(result) {
      const els = _flightAppDialogElements();
      if (els.overlay) els.overlay.style.display = 'none';
      const resolver = FLIGHTAPP_DIALOG_STATE.resolver;
      FLIGHTAPP_DIALOG_STATE.resolver = null;
      FLIGHTAPP_DIALOG_STATE.mode = 'alert';
      if (resolver) resolver(result);
    }

    function _flightAppOpenDialog(mode, message, options) {
      return new Promise(function(resolve) {
        const opts = options || {};
        const els = _flightAppDialogElements();
        if (!els.overlay || !els.title || !els.message || !els.input || !els.ok || !els.cancel) {
          resolve(mode === 'prompt' ? null : true);
          return;
        }

        FLIGHTAPP_DIALOG_STATE.resolver = resolve;
        FLIGHTAPP_DIALOG_STATE.mode = mode;

        els.title.textContent = String(opts.title || 'Flight App');
        els.message.textContent = String(message || '');
        els.ok.textContent = String(opts.okText || (mode === 'alert' ? 'OK' : 'Confirm'));
        els.cancel.textContent = String(opts.cancelText || 'Cancel');
        els.cancel.style.display = mode === 'alert' ? 'none' : 'inline-flex';
        els.input.style.display = mode === 'prompt' ? 'block' : 'none';
        els.input.value = mode === 'prompt' ? String(opts.defaultValue || '') : '';
        els.input.type = String(opts.inputType || 'text');
        if (opts.inputMode) els.input.setAttribute('inputmode', String(opts.inputMode));
        else els.input.removeAttribute('inputmode');
        if (opts.step != null) els.input.setAttribute('step', String(opts.step));
        else els.input.removeAttribute('step');
        if (opts.min != null) els.input.setAttribute('min', String(opts.min));
        else els.input.removeAttribute('min');
        if (opts.max != null) els.input.setAttribute('max', String(opts.max));
        else els.input.removeAttribute('max');
        els.overlay.style.display = 'flex';

        els.ok.onclick = function() {
          if (mode === 'prompt') {
            _flightAppDialogClose(els.input.value);
            return;
          }
          _flightAppDialogClose(true);
        };
        els.cancel.onclick = function() {
          _flightAppDialogClose(mode === 'prompt' ? null : false);
        };
        els.overlay.onclick = function(evt) {
          if (evt.target === els.overlay && mode !== 'alert') {
            _flightAppDialogClose(mode === 'prompt' ? null : false);
          }
        };
        els.input.onkeydown = function(evt) {
          if (evt.key === 'Enter') {
            evt.preventDefault();
            els.ok.click();
          } else if (evt.key === 'Escape' && mode !== 'alert') {
            evt.preventDefault();
            els.cancel.click();
          }
        };
        if (mode === 'prompt') {
          setTimeout(function() { els.input.focus(); els.input.select(); }, 20);
        } else {
          setTimeout(function() { els.ok.focus(); }, 20);
        }
      });
    }

    window.flightAppAlert = function(message, options) {
      return _flightAppOpenDialog('alert', message, options);
    };

    window.flightAppConfirm = function(message, options) {
      return _flightAppOpenDialog('confirm', message, options);
    };

    window.flightAppPrompt = function(message, defaultValue, options) {
      return _flightAppOpenDialog('prompt', message, { ...(options || {}), defaultValue: defaultValue });
    };

    function _dutyGetCachedPromptState_() {
      const v = cacheGet(OFFLINE_CACHE_KEYS.DUTY_PROMPT_STATE);
      return (v && typeof v === 'object') ? v : {};
    }

    function _dutySetCachedPromptState_(state) {
      const next = (state && typeof state === 'object') ? state : {};
      cacheSet(OFFLINE_CACHE_KEYS.DUTY_PROMPT_STATE, next);
    }

    function _dutyNowBsb_() {
      const now = new Date();
      const ymd = new Intl.DateTimeFormat('sv-SE', { timeZone: 'America/Sao_Paulo', year: 'numeric', month: '2-digit', day: '2-digit' }).format(now);
      const hh = Number(new Intl.DateTimeFormat('en-GB', { timeZone: 'America/Sao_Paulo', hour12: false, hour: '2-digit' }).format(now));
      const mm = Number(new Intl.DateTimeFormat('en-GB', { timeZone: 'America/Sao_Paulo', hour12: false, minute: '2-digit' }).format(now));
      const hhmm = String(hh).padStart(2, '0') + ':' + String(mm).padStart(2, '0');
      return { ymd: ymd, hour: hh, minute: mm, hhmm: hhmm };
    }

    function _dutyResolvePilotName_() {
      const current = window.currentBriefingMission;
      if (current && String(current.pilot || '').trim()) return String(current.pilot || '').trim();

      const missions = cacheGet(OFFLINE_CACHE_KEYS.SCHEDULED_MISSIONS) || [];
      if (activeMission) {
        const selected = missions.find(function(m) { return String(m && m.id || '') === String(activeMission || ''); });
        if (selected && String(selected.pilot || '').trim()) return String(selected.pilot || '').trim();
      }

      const nonTbd = missions.find(function(m) {
        const p = String((m && m.pilot) || '').trim();
        const up = p.toUpperCase();
        return p && up !== 'PILOT TBD' && up !== 'TBD' && up !== 'UNASSIGNED';
      });
      if (nonTbd) return String(nonTbd.pilot || '').trim();
      return '';
    }

    function _dutyGetConfig_(snapshot) {
      return (snapshot && snapshot.dutyConfig && typeof snapshot.dutyConfig === 'object') ? snapshot.dutyConfig : {};
    }

    function _dutyGetGeofenceRadiusKm_(snapshot) {
      const cfg = _dutyGetConfig_(snapshot);
      const raw = String(cfg.DUTY_GEOFENCE_RADIUS_KM || '').trim().replace(',', '.');
      const parsed = Number(raw);
      return (isFinite(parsed) && parsed > 0) ? parsed : 8;
    }

    function _dutyParseGeofenceLine_(lineRaw) {
      const line = String(lineRaw || '').trim();
      if (!line) return null;

      const hemiPair = line.match(/^(.+?[NSns])\s*[,;|\/]\s*(.+?[EWew])$/);
      if (hemiPair) {
        const latH = _parseCoordinate(hemiPair[1]);
        const lonH = _parseCoordinate(hemiPair[2]);
        if (isFinite(latH) && isFinite(lonH)) return { lat: latH, lon: lonH };
      }

      const parts = line.split(/\s*[;|]\s*/).filter(Boolean);
      if (parts.length === 2) {
        const latS = _parseCoordinate(parts[0]);
        const lonS = _parseCoordinate(parts[1]);
        if (isFinite(latS) && isFinite(lonS)) return { lat: latS, lon: lonS };
      }

      const commaPair = line.match(/^\s*([+\-]?\d+(?:[\.,]\d+)?)\s*,\s*([+\-]?\d+(?:[\.,]\d+)?)\s*$/);
      if (commaPair) {
        const latC = _parseCoordinate(commaPair[1]);
        const lonC = _parseCoordinate(commaPair[2]);
        if (isFinite(latC) && isFinite(lonC)) return { lat: latC, lon: lonC };
      }

      const nums = line.match(/[+\-]?\d+(?:[\.,]\d+)?/g) || [];
      if (nums.length === 2) {
        const latN = _parseCoordinate(nums[0]);
        const lonN = _parseCoordinate(nums[1]);
        if (isFinite(latN) && isFinite(lonN)) return { lat: latN, lon: lonN };
      }

      return null;
    }

    function _dutyParseGeofencePoints_(snapshot) {
      const cfg = _dutyGetConfig_(snapshot);
      const raw = String(cfg.DUTY_GEOFENCE_COORDS || cfg.DUTY_HOME_AIRPORTS || '').trim();
      if (!raw) return [];
      const lines = raw.split(/\r?\n+/).map(function(v) { return String(v || '').trim(); }).filter(Boolean);
      const points = [];
      lines.forEach(function(line, idx) {
        const parsed = _dutyParseGeofenceLine_(line);
        if (!parsed) return;
        if (!isFinite(parsed.lat) || !isFinite(parsed.lon)) return;
        if (parsed.lat < -90 || parsed.lat > 90 || parsed.lon < -180 || parsed.lon > 180) return;
        points.push({
          lat: parsed.lat,
          lon: parsed.lon,
          label: 'Ponto ' + (idx + 1)
        });
      });
      return points;
    }

    function _dutyDistanceKm_(lat1, lon1, lat2, lon2) {
      const nm = _offfltHaversineNm(lat1, lon1, lat2, lon2);
      return isFinite(nm) ? (nm * 1.852) : NaN;
    }

    function _dutyGetPositionContext_(snapshot) {
      return new Promise(function(resolve) {
        const points = _dutyParseGeofencePoints_(snapshot);
        const radiusKm = _dutyGetGeofenceRadiusKm_(snapshot);
        const base = {
          hasFence: points.length > 0,
          points: points,
          radiusKm: radiusKm,
          available: false,
          inFence: points.length === 0,
          hint: ''
        };

        if (!points.length) {
          resolve(base);
          return;
        }

        if (!(navigator && navigator.geolocation)) {
          resolve(Object.assign({}, base, { hint: 'Geolocalização indisponível para validar geofence.' }));
          return;
        }
        navigator.geolocation.getCurrentPosition(function(pos) {
          const c = pos && pos.coords ? pos.coords : {};
          const lat = Number(c.latitude);
          const lon = Number(c.longitude);
          if (!isFinite(lat) || !isFinite(lon)) {
            resolve(Object.assign({}, base, { hint: 'Posição inválida para validar geofence.' }));
            return;
          }

          let nearest = null;
          points.forEach(function(point) {
            const distKm = _dutyDistanceKm_(lat, lon, point.lat, point.lon);
            if (!isFinite(distKm)) return;
            if (!nearest || distKm < nearest.distKm) {
              nearest = { point: point, distKm: distKm };
            }
          });

          const inFence = !!(nearest && nearest.distKm <= radiusKm);
          const hint = nearest
            ? ('Distância até geofence: ' + nearest.distKm.toFixed(2) + ' km (raio ' + radiusKm.toFixed(1) + ' km).')
            : 'Não foi possível calcular distância até a geofence.';
          resolve(Object.assign({}, base, {
            available: true,
            lat: lat,
            lon: lon,
            nearest: nearest,
            inFence: inFence,
            hint: hint
          }));
        }, function() {
          resolve(Object.assign({}, base, { hint: 'Permissão de localização negada ou indisponível.' }));
        }, {
          enableHighAccuracy: false,
          timeout: 7000,
          maximumAge: 10 * 60 * 1000
        });
      });
    }

    function _dutyGeoVerifyElements_() {
      return {
        overlay: document.getElementById('duty-geo-verify-overlay'),
        map: document.getElementById('duty-geo-verify-map'),
        coords: document.getElementById('duty-geo-verify-coords'),
        ok: document.getElementById('duty-geo-verify-ok'),
        cancel: document.getElementById('duty-geo-verify-cancel')
      };
    }

    function _dutyGeoRenderVerifyMap_(geo) {
      const els = _dutyGeoVerifyElements_();
      if (!els.map || !geo || !geo.nearest || !isFinite(geo.lat) || !isFinite(geo.lon)) return;

      _wpEnsureLeaflet(function() {
        if (!_dutyGeoVerifyMap) {
          _dutyGeoVerifyMap = L.map('duty-geo-verify-map', { zoomControl: true, attributionControl: false });
          L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 16, minZoom: 2 }).addTo(_dutyGeoVerifyMap);
        }

        if (_dutyGeoVerifyCurrentMarker) {
          try { _dutyGeoVerifyMap.removeLayer(_dutyGeoVerifyCurrentMarker); } catch (e) {}
          _dutyGeoVerifyCurrentMarker = null;
        }
        if (_dutyGeoVerifyTargetMarker) {
          try { _dutyGeoVerifyMap.removeLayer(_dutyGeoVerifyTargetMarker); } catch (e) {}
          _dutyGeoVerifyTargetMarker = null;
        }
        if (_dutyGeoVerifyCircle) {
          try { _dutyGeoVerifyMap.removeLayer(_dutyGeoVerifyCircle); } catch (e) {}
          _dutyGeoVerifyCircle = null;
        }

        _dutyGeoVerifyCurrentMarker = L.marker([geo.lat, geo.lon]).bindTooltip('Sua posição').addTo(_dutyGeoVerifyMap);
        _dutyGeoVerifyTargetMarker = L.marker([geo.nearest.point.lat, geo.nearest.point.lon]).bindTooltip('Centro geofence').addTo(_dutyGeoVerifyMap);
        _dutyGeoVerifyCircle = L.circle([geo.nearest.point.lat, geo.nearest.point.lon], {
          radius: Number(geo.radiusKm || 0) * 1000,
          color: '#1976d2',
          weight: 2,
          fillColor: '#90caf9',
          fillOpacity: 0.2
        }).addTo(_dutyGeoVerifyMap);

        const bounds = L.latLngBounds([
          [geo.lat, geo.lon],
          [geo.nearest.point.lat, geo.nearest.point.lon]
        ]);
        _dutyGeoVerifyMap.fitBounds(bounds.pad(0.6));
        setTimeout(function() { if (_dutyGeoVerifyMap) _dutyGeoVerifyMap.invalidateSize(); }, 40);
      });
    }

    function _dutyOpenGeoVerifyModal_(geo) {
      return new Promise(function(resolve) {
        const els = _dutyGeoVerifyElements_();
        if (!els.overlay || !els.ok || !els.cancel || !els.coords) {
          resolve(true);
          return;
        }

        const nearest = geo && geo.nearest ? geo.nearest : null;
        const distText = nearest ? nearest.distKm.toFixed(2) : '--';
        const radiusText = isFinite(geo && geo.radiusKm) ? Number(geo.radiusKm).toFixed(1) : '--';
        els.coords.textContent = 'Atual: ' + Number(geo.lat).toFixed(6) + ', ' + Number(geo.lon).toFixed(6)
          + ' | Centro: ' + Number(nearest && nearest.point.lat || 0).toFixed(6) + ', ' + Number(nearest && nearest.point.lon || 0).toFixed(6)
          + ' | Distância: ' + distText + ' km de ' + radiusText + ' km';

        els.overlay.style.display = 'flex';
        _dutyGeoRenderVerifyMap_(geo);

        const close = function(ok) {
          els.overlay.style.display = 'none';
          els.ok.onclick = null;
          els.cancel.onclick = null;
          els.overlay.onclick = null;
          resolve(!!ok);
        };

        els.ok.onclick = function() { close(true); };
        els.cancel.onclick = function() { close(false); };
        els.overlay.onclick = function(ev) {
          if (ev.target === els.overlay) close(false);
        };
      });
    }

    function _dutySavePayload_(payload, doneCb) {
      runOrQueueServerAction({
        method: 'savePilotDutyReport',
        args: [payload],
        label: 'Jornada'
      }, {
        onSuccess: function() {
          if (typeof doneCb === 'function') doneCb(true);
        },
        onQueued: function() {
          if (window.M) M.toast({ html: 'Offline: jornada enfileirada para sincronizar', classes: 'orange' });
          if (typeof doneCb === 'function') doneCb(true);
        },
        onFailure: function(err) {
          if (window.M) M.toast({ html: 'Falha ao salvar jornada', classes: 'red' });
          console.error('savePilotDutyReport failed', err);
          if (typeof doneCb === 'function') doneCb(false);
        }
      });
    }

    async function _dutyPromptMorningStart_(ctx, snapshot, geo) {
      const now = _dutyNowBsb_();
      const pilot = ctx.pilotName;

      if (geo && geo.hasFence) {
        if (!geo.available) {
          return { prompted: false, blocked: 'location-unavailable', message: geo.hint || 'Sem localização para validar geofence.' };
        }
        if (!geo.inFence) {
          return { prompted: false, blocked: 'outside-geofence', message: geo.hint || 'Fora da geofence configurada.' };
        }
        const verified = await _dutyOpenGeoVerifyModal_(geo);
        if (!verified) {
          return { prompted: false, blocked: 'map-cancel', message: 'Verificação de posição cancelada.' };
        }
      }

      const msg = [
        'Iniciar jornada de trabalho agora?',
        '',
        'Piloto: ' + pilot,
        'Horário sugerido: ' + now.hhmm + ' BSB',
        (geo && geo.hasFence) ? (geo.hint || 'Geofence validada por coordenadas.') : 'Dica: você também pode confirmar pelo horário padrão das 08:00.'
      ].join('\n');

      const ok = await window.flightAppConfirm(msg, {
        title: 'Jornada Diária',
        okText: 'Iniciar',
        cancelText: 'Ignorar'
      });

      if (!ok) {
        _dutySavePayload_({
          pilotName: pilot,
          dateYmd: ctx.dateYmd,
          status: 'DUTY_MORNING_IGNORED',
          ignoredMorning: true,
          summaryText: 'Prompt matinal ignorado'
        }, function() {
          const cache = _dutyGetCachedPromptState_();
          cache[ctx.dateYmd] = Object.assign({}, cache[ctx.dateYmd] || {}, { morningPrompted: true });
          _dutySetCachedPromptState_(cache);
        });
        return { prompted: true };
      }

      const entered = await window.flightAppPrompt('Informe o horário de início da jornada (HH:MM):', now.hhmm, {
        title: 'Jornada Diária',
        okText: 'Salvar',
        cancelText: 'Cancelar',
        inputType: 'text',
        inputMode: 'numeric'
      });

      const start = String(entered || '').trim();
      const valid = /^(\d{1,2}):(\d{2})$/.test(start);
      if (!valid) {
        if (window.M) M.toast({ html: 'Horário inválido. Use HH:MM', classes: 'orange' });
        return { prompted: true };
      }

      _dutySavePayload_({
        pilotName: pilot,
        dateYmd: ctx.dateYmd,
        status: 'DUTY_OPEN',
        startTime: start,
        summaryText: 'Jornada iniciada às ' + start
      }, function(okSave) {
        if (okSave && window.M) M.toast({ html: 'Jornada iniciada às ' + start, classes: 'green' });
        const cache = _dutyGetCachedPromptState_();
        cache[ctx.dateYmd] = Object.assign({}, cache[ctx.dateYmd] || {}, { morningPrompted: true });
        _dutySetCachedPromptState_(cache);
      });
      return { prompted: true };
    }

    function _dutyCloseoutElements_() {
      return {
        overlay: document.getElementById('duty-closeout-overlay'),
        start: document.getElementById('duty-closeout-start'),
        flightHours: document.getElementById('duty-closeout-flight-hours'),
        summary: document.getElementById('duty-closeout-summary'),
        desc: document.getElementById('duty-closeout-desc'),
        submit: document.getElementById('duty-closeout-submit'),
        ignore: document.getElementById('duty-closeout-ignore')
      };
    }

    function _dutyOpenCloseoutModal_(ctx, snapshot) {
      return new Promise(function(resolve) {
        const els = _dutyCloseoutElements_();
        if (!els.overlay || !els.start || !els.flightHours || !els.summary || !els.desc || !els.submit || !els.ignore) {
          resolve({ action: 'ignore' });
          return;
        }

        const auto = (snapshot && snapshot.autofill && typeof snapshot.autofill === 'object') ? snapshot.autofill : {};
        const entry = (snapshot && snapshot.entry && typeof snapshot.entry === 'object') ? snapshot.entry : {};
        const start = String(entry.startTime || '').trim() || '--:--';
        const defaultHours = (entry.flightHours !== '' && entry.flightHours != null)
          ? String(entry.flightHours)
          : String(Number(auto.totalFlightHours || 0).toFixed(1));

        els.start.textContent = start;
        els.flightHours.value = defaultHours;
        els.summary.value = String(auto.summaryLine || 'Sem voos registrados no dia');
        els.desc.value = String(entry.description || auto.summaryLine || '');
        els.overlay.style.display = 'flex';

        const cleanup = function(result) {
          els.overlay.style.display = 'none';
          els.submit.onclick = null;
          els.ignore.onclick = null;
          els.overlay.onclick = null;
          resolve(result);
        };

        els.submit.onclick = function() {
          cleanup({
            action: 'submit',
            flightHours: String(els.flightHours.value || '').trim(),
            summaryText: String(els.summary.value || '').trim(),
            description: String(els.desc.value || '').trim()
          });
        };
        els.ignore.onclick = function() {
          cleanup({ action: 'ignore' });
        };
        els.overlay.onclick = function(ev) {
          if (ev.target === els.overlay) cleanup({ action: 'ignore' });
        };
      });
    }

    async function _dutyPromptEveningClose_(ctx, snapshot, geo) {
      if (geo && geo.hasFence) {
        if (!geo.available) {
          return { prompted: false, blocked: 'location-unavailable', message: geo.hint || 'Sem localização para validar geofence.' };
        }
        if (!geo.inFence) {
          return { prompted: false, blocked: 'outside-geofence', message: geo.hint || 'Fora da geofence configurada.' };
        }
      }

      const res = await _dutyOpenCloseoutModal_(ctx, snapshot);
      const cache = _dutyGetCachedPromptState_();
      cache[ctx.dateYmd] = Object.assign({}, cache[ctx.dateYmd] || {}, { eveningPrompted: true });
      _dutySetCachedPromptState_(cache);

      if (!res || res.action !== 'submit') {
        _dutySavePayload_({
          pilotName: ctx.pilotName,
          dateYmd: ctx.dateYmd,
          status: 'DUTY_EVENING_IGNORED',
          startTime: String(snapshot && snapshot.entry && snapshot.entry.startTime || '').trim(),
          ignoredEvening: true,
          summaryText: 'Prompt de encerramento ignorado'
        });
        return { prompted: true };
      }

      const now = _dutyNowBsb_();
      const flightHours = Number(String(res.flightHours || '').replace(',', '.'));
      _dutySavePayload_({
        pilotName: ctx.pilotName,
        dateYmd: ctx.dateYmd,
        status: 'DUTY_CLOSED',
        startTime: String(snapshot && snapshot.entry && snapshot.entry.startTime || '').trim(),
        endTime: now.hhmm,
        flightHours: isFinite(flightHours) ? flightHours : '',
        summaryText: res.summaryText,
        description: res.description,
        autoSummary: (snapshot && snapshot.autofill) ? snapshot.autofill : {}
      }, function(okSave) {
        if (okSave && window.M) {
          M.toast({ html: 'Jornada encerrada às ' + now.hhmm, classes: 'green' });
        }
      });
      return { prompted: true };
    }

    function _dutyWarnGeofenceOnce_(ymd, key, message) {
      const cache = _dutyGetCachedPromptState_();
      const todayState = (cache[ymd] && typeof cache[ymd] === 'object') ? cache[ymd] : {};
      if (todayState[key]) return;
      todayState[key] = true;
      cache[ymd] = todayState;
      _dutySetCachedPromptState_(cache);
      if (window.M && message) M.toast({ html: message, classes: 'orange' });
    }

    function _dutyCanRunByTime_(mode, now) {
      if (!now) return false;
      if (mode === 'morning') return now.hour >= 8 && now.hour <= 11;
      if (mode === 'evening') return now.hour >= 17 && now.hour <= 23;
      return false;
    }

    function _dutyCheckPrompts_() {
      if (_dutyPromptBusy) return;
      _dutyPromptBusy = true;

      const now = _dutyNowBsb_();
      const pilotName = _dutyResolvePilotName_();
      if (!pilotName) {
        _dutyPromptBusy = false;
        return;
      }

      const cache = _dutyGetCachedPromptState_();
      const todayState = (cache[now.ymd] && typeof cache[now.ymd] === 'object') ? cache[now.ymd] : {};
      const context = { pilotName: pilotName, dateYmd: now.ymd };

      const finish = function() {
        _dutyPromptBusy = false;
      };

      if (!(window.google && google.script && google.script.run)) {
        finish();
        return;
      }

      google.script.run
        .withSuccessHandler(async function(snapshot) {
          try {
            if (!snapshot || snapshot.success !== true) { finish(); return; }
            const geo = await _dutyGetPositionContext_(snapshot);

            if (_dutyCanRunByTime_('morning', now) && snapshot.shouldMorningPrompt && !todayState.morningPrompted) {
              const morningRes = await _dutyPromptMorningStart_(context, snapshot, geo);
              if (morningRes && morningRes.blocked) {
                _dutyWarnGeofenceOnce_(context.dateYmd, 'morningGeofenceWarned', morningRes.message || 'Geofence não validada para prompt matinal.');
              }
              finish();
              return;
            }

            if (_dutyCanRunByTime_('evening', now) && snapshot.shouldEveningPrompt && !todayState.eveningPrompted) {
              const eveningRes = await _dutyPromptEveningClose_(context, snapshot, geo);
              if (eveningRes && eveningRes.blocked) {
                _dutyWarnGeofenceOnce_(context.dateYmd, 'eveningGeofenceWarned', eveningRes.message || 'Geofence não validada para prompt de encerramento.');
              }
              finish();
              return;
            }
          } catch (e) {
            console.error('Duty prompt flow failed', e);
          }
          finish();
        })
        .withFailureHandler(function(err) {
          console.warn('getPilotDutyPromptSnapshot failed', err);
          finish();
        })
        .getPilotDutyPromptSnapshot(pilotName, now.ymd);
    }

    function initDutyPromptScheduler_() {
      if (_dutyPromptTimer) clearInterval(_dutyPromptTimer);
      _dutyPromptTimer = setInterval(_dutyCheckPrompts_, 60000);
      setTimeout(_dutyCheckPrompts_, 3500);
    }

    function cacheSet(key, value) {
      const envelope = {
        savedAt: Date.now(),
        value: value
      };
      const serialized = JSON.stringify(envelope);

      try {
        localStorage.setItem(key, serialized);
        if (isMissionCacheKey(key)) {
          pruneMissionCache(OFFLINE_CACHE_MAX_MISSIONS);
        }
        return true;
      } catch (e) {
        console.warn('cacheSet failed, trying prune', key, e);
        try {
          pruneMissionCache(Math.max(10, OFFLINE_CACHE_MAX_MISSIONS - 10));
          localStorage.setItem(key, serialized);
          if (isMissionCacheKey(key)) {
            pruneMissionCache(OFFLINE_CACHE_MAX_MISSIONS);
          }
          return true;
        } catch (retryErr) {
          console.warn('cacheSet retry failed', key, retryErr);
          return false;
        }
      }
    }

    function getCacheEnvelope(key) {
      try {
        const raw = localStorage.getItem(key);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        return parsed && typeof parsed === 'object' ? parsed : null;
      } catch (e) {
        return null;
      }
    }

    function isMissionCacheKey(key) {
      return String(key || '').startsWith('mba_cache_mission_');
    }

    function listMissionCacheEntries() {
      const entries = [];
      try {
        for (let i = 0; i < localStorage.length; i++) {
          const key = String(localStorage.key(i) || '');
          if (!isMissionCacheKey(key)) continue;
          const envelope = getCacheEnvelope(key) || {};
          entries.push({
            key: key,
            savedAt: Number(envelope.savedAt || 0)
          });
        }
      } catch (e) {
        console.warn('listMissionCacheEntries failed', e);
      }
      return entries.sort((a, b) => a.savedAt - b.savedAt);
    }

    function pruneMissionCache(maxToKeep) {
      const limit = Math.max(0, Number(maxToKeep || 0));
      const entries = listMissionCacheEntries();
      if (entries.length <= limit) return 0;

      let removed = 0;
      const removeCount = entries.length - limit;
      for (let i = 0; i < removeCount; i++) {
        try {
          localStorage.removeItem(entries[i].key);
          removed += 1;
        } catch (e) {
          console.warn('Failed removing cached mission', entries[i].key, e);
        }
      }
      return removed;
    }

    function cacheGet(key) {
      try {
        const raw = localStorage.getItem(key);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        return parsed && Object.prototype.hasOwnProperty.call(parsed, 'value') ? parsed.value : null;
      } catch (e) {
        console.warn('cacheGet failed', key, e);
        return null;
      }
    }

    function getOutboxItems() {
      const items = cacheGet(OFFLINE_CACHE_KEYS.OUTBOX);
      return Array.isArray(items) ? items : [];
    }

    function setOutboxItems(items) {
      return cacheSet(OFFLINE_CACHE_KEYS.OUTBOX, Array.isArray(items) ? items : []);
    }

    function getPendingOutboxCount() {
      return getOutboxItems().filter(i => String(i.status || 'pending') === 'pending').length;
    }

    function getFailedOutboxCount() {
      return getOutboxItems().filter(i => String(i.status || 'pending') === 'failed').length;
    }

    function updateOutboxStatusPanel() {
      const el = document.getElementById('outbox-status');
      if (!el) return;
      const pending = getPendingOutboxCount();
      const failed = getFailedOutboxCount();
      if (failed > 0) {
        el.style.background = '#fff8e1';
        el.style.borderColor = '#ffecb3';
        el.style.color = '#8a6d1b';
      } else if (pending > 0) {
        el.style.background = '#e3f2fd';
        el.style.borderColor = '#bbdefb';
        el.style.color = '#0b5394';
      } else {
        el.style.background = '#f9f9f9';
        el.style.borderColor = '#ececec';
        el.style.color = '#555';
      }
      el.textContent = `Outbox: ${pending} pending, ${failed} failed`;
    }

    function updateOutboxButtonState() {
      const btn = document.getElementById('btn-sync-outbox');
      const retryBtn = document.getElementById('btn-retry-failed-outbox');
      if (!btn) return;

      const pending = getPendingOutboxCount();
      const failed = getFailedOutboxCount();
      btn.textContent = `SYNC QUEUED (${pending})`;
      btn.disabled = outboxSyncInProgress || pending === 0 || !isServerAvailable();
      btn.style.opacity = btn.disabled ? '0.6' : '1';

      if (retryBtn) {
        retryBtn.textContent = `RETRY FAILED (${failed})`;
        retryBtn.disabled = outboxSyncInProgress || failed === 0 || !isServerAvailable();
        retryBtn.style.opacity = retryBtn.disabled ? '0.6' : '1';
      }

      updateOutboxStatusPanel();
    }

    function retryFailedOutboxItems() {
      const items = getOutboxItems();
      const updated = items.map(item => {
        if (String(item.status || 'pending') !== 'failed') return item;
        return {
          ...item,
          status: 'pending',
          attempts: 0,
          lastError: ''
        };
      });
      setOutboxItems(updated);
      updateOutboxButtonState();
      processOutboxQueue();
    }

    function executeServerMethod(methodName, args, onSuccess, onFailure) {
      if (!isServerAvailable()) {
        if (typeof onFailure === 'function') onFailure(new Error('Server unavailable'));
        return;
      }
      const method = String(methodName || '').trim();
      if (!method || !google.script.run[method]) {
        if (typeof onFailure === 'function') onFailure(new Error(`Unknown server method: ${method}`));
        return;
      }

      google.script.run
        .withSuccessHandler(resp => {
          if (typeof onSuccess === 'function') onSuccess(resp);
        })
        .withFailureHandler(err => {
          if (typeof onFailure === 'function') onFailure(err);
        })[method].apply(google.script.run, Array.isArray(args) ? args : []);
    }

    function enqueueOutboxAction(action) {
      const items = getOutboxItems();
      items.push({
        id: `outbox_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
        method: String(action.method || '').trim(),
        args: Array.isArray(action.args) ? action.args : [],
        label: String(action.label || action.method || 'Action'),
        queuedAt: new Date().toISOString(),
        status: 'pending',
        attempts: 0,
        lastError: ''
      });
      setOutboxItems(items);
      updateOutboxButtonState();
    }

    function runOrQueueServerAction(action, handlers) {
      const h = handlers || {};
      if (!action || !action.method) {
        if (typeof h.onFailure === 'function') h.onFailure(new Error('Invalid action'));
        return;
      }

      if (!isServerAvailable()) {
        enqueueOutboxAction(action);
        if (typeof h.onQueued === 'function') h.onQueued();
        if (window.M) M.toast({ html: `${action.label || action.method} queued for sync`, classes: 'orange' });
        return;
      }

      executeServerMethod(action.method, action.args,
        resp => {
          if (typeof h.onSuccess === 'function') h.onSuccess(resp);
        },
        err => {
          if (!isServerAvailable()) {
            enqueueOutboxAction(action);
            if (typeof h.onQueued === 'function') h.onQueued();
            if (window.M) M.toast({ html: `${action.label || action.method} queued after disconnect`, classes: 'orange' });
            return;
          }
          if (typeof h.onFailure === 'function') h.onFailure(err);
        }
      );
    }

    function processOutboxQueue() {
      if (!isServerAvailable()) {
        updateOutboxButtonState();
        return;
      }
      if (outboxSyncInProgress) return;

      const items = getOutboxItems();
      const pending = items.filter(i => String(i.status || 'pending') === 'pending');
      if (!pending.length) {
        updateOutboxButtonState();
        return;
      }

      outboxSyncInProgress = true;
      updateOutboxButtonState();

      let index = 0;
      let synced = 0;
      let failed = 0;

      const step = function() {
        const currentList = getOutboxItems();
        const queue = currentList.filter(i => String(i.status || 'pending') === 'pending');

        if (index >= queue.length) {
          outboxSyncInProgress = false;
          updateOutboxButtonState();
          if (window.M) {
            const failedTotal = getFailedOutboxCount();
            const cls = failedTotal ? 'orange' : 'green';
            M.toast({ html: `Queued sync complete: ${synced} synced${failedTotal ? `, ${failedTotal} failed` : ''}`, classes: cls, displayLength: 4500 });
          }
          return;
        }

        const item = queue[index++];
        executeServerMethod(item.method, item.args,
          () => {
            const after = getOutboxItems().filter(it => it.id !== item.id);
            setOutboxItems(after);
            synced += 1;
            updateOutboxButtonState();
            setTimeout(step, 40);
          },
          err => {
            const after = getOutboxItems().map(it => {
              if (it.id !== item.id) return it;
              const attempts = Number(it.attempts || 0) + 1;
              const status = attempts >= OUTBOX_MAX_ATTEMPTS ? 'failed' : 'pending';
              return {
                ...it,
                attempts: attempts,
                status: status,
                lastError: err && err.message ? err.message : String(err || 'sync failure')
              };
            });
            setOutboxItems(after);
            failed += 1;
            updateOutboxButtonState();
            setTimeout(step, 120);
          }
        );
      };

      step();
    }

    function syncOutboxNow() {
      processOutboxQueue();
    }

    window.runOrQueueServerAction = runOrQueueServerAction;

    function getCachedMissionCount() {
      try {
        let count = 0;
        for (let i = 0; i < localStorage.length; i++) {
          const key = String(localStorage.key(i) || '');
          if (key.startsWith('mba_cache_mission_')) count += 1;
        }
        return count;
      } catch (e) {
        console.warn('getCachedMissionCount failed', e);
        return 0;
      }
    }

    function formatSyncTime(isoString) {
      if (!isoString) return 'never';
      return _pilotFormatBsbDateTime_(isoString, 'unknown');
    }

    function formatStatusDate(isoString) {
      return _pilotFormatDateMonDayYear_(isoString || '', '');
    }

    function _cacheHasValidEnvelopePayload_(payload) {
      const points = Array.isArray(payload && payload.envelopeData) ? payload.envelopeData : [];
      const valid = points.filter(function(p) {
        const x = parseFloat(p && (p.CG_Arm_X ?? p.cgArmX ?? p.cgArm ?? p.arm ?? p.x));
        const y = parseFloat(p && (p.Weight_Y ?? p.weightY ?? p.weight ?? p.y));
        return !isNaN(x) && !isNaN(y);
      });
      return valid.length >= 3;
    }

    function getCachedEnvelopeCount() {
      try {
        let count = 0;
        const keys = Object.keys(localStorage || {});
        for (let i = 0; i < keys.length; i++) {
          const key = String(keys[i] || '');
          if (!key.startsWith('mba_cache_envelope_')) continue;
          const raw = localStorage.getItem(key);
          if (!raw) continue;
          let parsed;
          try { parsed = JSON.parse(raw); } catch (e) { continue; }
          const payload = (parsed && parsed.value && typeof parsed.value === 'object') ? parsed.value : parsed;
          if (_cacheHasValidEnvelopePayload_(payload)) count += 1;
        }
        return count;
      } catch (e) {
        return 0;
      }
    }

    function _envelopeCacheKey_(aircraftReg) {
      const reg = String(aircraftReg || '').trim().toUpperCase();
      return reg ? ('mba_cache_envelope_' + reg) : '';
    }

    function _readEnvelopeCachePayload_(aircraftReg) {
      try {
        const key = _envelopeCacheKey_(aircraftReg);
        if (!key) return null;
        const raw = localStorage.getItem(key);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        return (parsed && parsed.value && typeof parsed.value === 'object') ? parsed.value : parsed;
      } catch (e) {
        return null;
      }
    }

    function _isEnvelopePayloadFresh_(payload) {
      try {
        if (!_cacheHasValidEnvelopePayload_(payload)) return false;
        const cachedAt = String((payload && payload.cachedAt) || '').trim();
        if (!cachedAt) return true;
        const t = new Date(cachedAt).getTime();
        if (!isFinite(t) || t <= 0) return true;
        return (Date.now() - t) <= OFFLINE_ENVELOPE_STALE_MS;
      } catch (e) {
        return false;
      }
    }

    function _knownAircraftRegsFromMissions_(missions) {
      return Array.from(new Set(
        (Array.isArray(missions) ? missions : [])
          .map(function(m) { return String(m && m.acft || '').trim().toUpperCase(); })
          .filter(Boolean)
      ));
    }

    function _knownAircraftRegsForEnvelopeSync_() {
      const regs = [];
      const add = function(v) {
        const reg = String(v || '').trim().toUpperCase();
        if (reg) regs.push(reg);
      };

      try {
        const fromAppData = (window.appData && Array.isArray(window.appData.aircraft)) ? window.appData.aircraft : [];
        fromAppData.forEach(function(a) { add(a && a.reg); });
      } catch (e) {}

      try {
        const cachedDropdown = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA) || {};
        const fromCached = Array.isArray(cachedDropdown.aircraft) ? cachedDropdown.aircraft : [];
        fromCached.forEach(function(a) { add(a && a.reg); });
      } catch (e) {}

      try {
        const missions = cacheGet(OFFLINE_CACHE_KEYS.SCHEDULED_MISSIONS) || [];
        _knownAircraftRegsFromMissions_(missions).forEach(add);
      } catch (e) {}

      return Array.from(new Set(regs));
    }

    function _storeEnvelopePayloadForAircraft_(aircraftReg, payload) {
      try {
        const reg = String(aircraftReg || '').trim().toUpperCase();
        const env = Array.isArray(payload && payload.envelopeData) ? payload.envelopeData : [];
        if (!reg || env.length < 3) return false;
        localStorage.setItem('mba_cache_envelope_' + reg, JSON.stringify({
          aircraft: reg,
          cachedAt: new Date().toISOString(),
          envelopeData: env
        }));
        return true;
      } catch (e) {
        return false;
      }
    }

    function syncEnvelopeCacheForAircraftRegs_(aircraftRegs, options) {
      const opts = options || {};
      const regs = Array.from(new Set(
        (Array.isArray(aircraftRegs) ? aircraftRegs : [])
          .map(function(r) { return String(r || '').trim().toUpperCase(); })
          .filter(Boolean)
      ));

      if (!regs.length) {
        updateOfflineCacheStatus();
        updateOfflineFlightButtonState();
        return;
      }
      if (!isServerAvailable()) return;
      if (!(window.google && google.script && google.script.run)) return;

      regs.forEach(function(reg) {
        try {
          const cached = _readEnvelopeCachePayload_(reg);
          if (!opts.forceRefresh && _isEnvelopePayloadFresh_(cached)) return;
        } catch (e) {}

        google.script.run
          .withSuccessHandler(function(payload) {
            if (_storeEnvelopePayloadForAircraft_(reg, payload)) {
              updateOfflineCacheStatus();
              updateOfflineFlightButtonState();
            }
          })
          .withFailureHandler(function(err) {
            console.warn('syncEnvelopeCacheForAircraftRegs failed', reg, err);
          })
          .getWbEnvelopeByAircraft(reg);
      });
    }

    function updateOfflineFlightButtonState() {
      const btn = document.getElementById('btn-offline-flight');
      if (!btn) return;
      const hasEnvelope = getCachedEnvelopeCount() > 0;
      btn.style.border = hasEnvelope ? '2px solid #2e7d32' : '2px solid #c62828';
      btn.style.boxShadow = hasEnvelope ? '0 0 0 2px rgba(46,125,50,0.15)' : '0 0 0 2px rgba(198,40,40,0.12)';
      btn.title = hasEnvelope
        ? 'Offline flight ready: CG envelope cache present'
        : 'Offline flight blocked: no CG envelope cache available yet';
    }

    // Track per-row readiness so overall icon can be computed any time
    window._readinessState = { missions: 'none', cg: 'none', unknown: 'none', mapdb: 'none' };

    function _isMissionReadyForLock_() {
      const s = window._readinessState || {};
      return s.missions === 'ok' && s.cg === 'ok' && s.unknown === 'ok';
    }

    function _updateMissionReadyLockUI_() {
      const ready = _isMissionReadyForLock_();
      const actionButtons = document.querySelectorAll('#tab1 .action-grid .action-btn');
      actionButtons.forEach(function(btn) {
        if (!btn || btn.id === 'btn-acft-docs') return; // manuals always accessible
        btn.classList.toggle('readiness-locked', !ready);
        btn.title = ready ? 'Mission ready' : 'Offline pack not ready. Press REFRESH and wait for Mission Ready.';
      });

      for (let i = 2; i <= 8; i++) {
        const step = document.getElementById('step' + i);
        if (!step) continue;
        step.classList.toggle('readiness-step-locked', !ready);
      }

      const btnNext = document.getElementById('btnNext');
      if (btnNext && Number(currentTab || 1) === 1) {
        btnNext.disabled = !ready;
        btnNext.style.opacity = ready ? '1' : '0.6';
        btnNext.title = ready ? '' : 'Offline pack not ready. Press REFRESH and wait for Mission Ready.';
      }
    }

    function _computeUnknownPackReadiness_() {
      const cachedDropdown = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA) || {};
      const source = (window.appData && typeof window.appData === 'object') ? window.appData : cachedDropdown;
      const airports = Array.isArray(source.airports) ? source.airports.length : 0;
      const waypoints = Array.isArray(source.waypoints) ? source.waypoints.length : 0;
      const routes = Array.isArray(source.routes) ? source.routes.length : 0;
      const aircraft = Array.isArray(source.aircraft) ? source.aircraft.length : 0;
      const pilots = Array.isArray(source.pilots) ? source.pilots.length : 0;
      const envelopes = getCachedEnvelopeCount();

      const ready = airports > 0 && waypoints > 0 && routes > 0 && aircraft > 0 && pilots > 0 && envelopes > 0;
      const missing = [];
      if (airports <= 0) missing.push('airports');
      if (waypoints <= 0) missing.push('waypoints');
      if (routes <= 0) missing.push('routes');
      if (aircraft <= 0) missing.push('aircraft');
      if (pilots <= 0) missing.push('pilots');
      if (envelopes <= 0) missing.push('envelopes');

      return {
        ready: ready,
        label: ready
          ? 'Unknown-flight pack loaded'
          : (missing.length ? ('Missing: ' + missing.slice(0, 3).join(', ') + (missing.length > 3 ? '...' : '')) : 'Not ready')
      };
    }

    function tab1GuardedAction_(actionName) {
      if (!_isMissionReadyForLock_()) {
        if (window.M) M.toast({ html: 'Offline pack not ready. Press REFRESH and wait for MISSION READY.', classes: 'orange', displayLength: 3500 });
        return;
      }
      const fn = window[String(actionName || '').trim()];
      if (typeof fn === 'function') fn();
    }

    function _setReadinessDot_(id, state) {
      // state: 'ok' | 'bad' | 'warn' | 'sync' | 'none'
      const el = document.getElementById(id);
      if (!el) return;
      const colors = { ok: '#2e7d32', bad: '#c62828', warn: '#f9a825', sync: '#1565c0', none: '#bdbdbd' };
      el.style.background = colors[state] || colors.none;
    }

    function _recomputeOverallReadiness_() {
      const s = window._readinessState;
      const icon = document.getElementById('readiness-overall-icon');
      const label = document.getElementById('readiness-overall-label');
      const allOk = s.missions === 'ok' && s.cg === 'ok' && s.unknown === 'ok';
      const anySyncing = s.missions === 'sync' || s.cg === 'sync' || s.unknown === 'sync';
      if (anySyncing) {
        if (icon) {
          icon.textContent = 'Syncing';
          icon.style.color = '#ffffff';
          icon.style.background = '#1565c0';
          icon.style.borderColor = '#0d47a1';
        }
        if (label) { label.textContent = 'Building Bush Pack'; label.style.color = '#1565c0'; }
      } else if (allOk) {
        if (icon) {
          icon.textContent = 'MISSION READY';
          icon.style.color = '#ffffff';
          icon.style.background = '#2e7d32';
          icon.style.borderColor = '#1b5e20';
        }
        if (label) { label.textContent = 'Bush Pilot Ready'; label.style.color = '#2e7d32'; }
      } else {
        if (icon) {
          icon.textContent = 'Not Ready';
          icon.style.color = '#ffffff';
          icon.style.background = '#8d6e63';
          icon.style.borderColor = '#5d4037';
        }
        if (label) { label.textContent = 'Offline Pack Needed'; label.style.color = '#c62828'; }
      }
      _updateMissionReadyLockUI_();
    }
    window._recomputeOverallReadiness_ = _recomputeOverallReadiness_;

    function updateOfflineCacheStatus(options) {
      const opts = options || {};
      try {
        if (typeof window.refreshOfflineMapPackStatus === 'function') {
          window.refreshOfflineMapPackStatus();
        }
      } catch (e) {
        try { if (window._setOfflineMapDbIndicator) window._setOfflineMapDbIndicator('warn', 'Status unavailable'); } catch (_) {}
      }

      const hasExplicitState = Object.prototype.hasOwnProperty.call(opts, 'state');
      const effectiveState = hasExplicitState
        ? String(opts.state || '').toLowerCase()
        : (missionPrefetchInProgress ? 'syncing' : (isServerAvailable() ? 'ready' : 'offline'));

      const meta = cacheGet(OFFLINE_CACHE_KEYS.PREFETCH_META) || {};
      const cachedCount = Number(typeof opts.cachedCount === 'number' ? opts.cachedCount : getCachedMissionCount());
      const total = Number(typeof opts.total === 'number' ? opts.total : (missionPrefetchProgress.total || meta.total || 0));
      const progress = Number(typeof opts.progress === 'number' ? opts.progress : (missionPrefetchProgress.progress || 0));
      const syncedAt = opts.syncedAt || meta.syncedAt || null;
      const syncedAtMs = syncedAt ? new Date(syncedAt).getTime() : 0;
      const isStale = syncedAtMs > 0 && (Date.now() - syncedAtMs) > OFFLINE_CACHE_STALE_MS;
      const envelopeCount = Number(typeof opts.envelopeCount === 'number' ? opts.envelopeCount : getCachedEnvelopeCount());
      const hasEnvelope = envelopeCount > 0;
      const hasAnyFlightCache = cachedCount > 0 || total > 0;
      const state = effectiveState;

      const missionLabel = document.getElementById('rlabel-missions');
      const cgLabel = document.getElementById('rlabel-cg');
      const unknownLabel = document.getElementById('rlabel-unknown');
      const progDiv = document.getElementById('readiness-progress');
      const progBar = document.getElementById('readiness-progress-bar');
      const progTxt = document.getElementById('readiness-progress-text');
      const progPct = document.getElementById('readiness-progress-percent');

      if (state === 'syncing') {
        window._readinessState.missions = 'sync';
        _setReadinessDot_('rdot-missions', 'sync');
        window._readinessState.unknown = 'sync';
        _setReadinessDot_('rdot-unknown', 'sync');
        const pct = total > 0 ? Math.max(0, Math.min(100, Math.round((progress / total) * 100))) : 0;
        const stageText = String(opts.message || '').trim();
        if (missionLabel) missionLabel.textContent = stageText || `Downloading ${progress}/${total || '…'}`;
        if (unknownLabel) unknownLabel.textContent = 'Building unknown-flight pack...';
        if (progDiv) progDiv.style.display = 'block';
        if (progBar) progBar.style.width = pct + '%';
        if (progTxt) progTxt.textContent = stageText || `Missions ${progress}/${total || 0}`;
        if (progPct) progPct.textContent = pct + '%';
        _recomputeOverallReadiness_();
        updateOfflineFlightButtonState();
        return;
      }

      if (progDiv) progDiv.style.display = 'none';
      if (progBar) progBar.style.width = '0%';
      if (progTxt) progTxt.textContent = '';
      if (progPct) progPct.textContent = '0%';

      const flightsReady = !hasAnyFlightCache || (cachedCount > 0 && !isStale);
      if (state === 'empty') {
        window._readinessState.missions = 'none';
        _setReadinessDot_('rdot-missions', 'none');
        if (missionLabel) missionLabel.textContent = 'Not cached';
      } else if (isStale) {
        window._readinessState.missions = 'bad';
        _setReadinessDot_('rdot-missions', 'bad');
        if (missionLabel) missionLabel.textContent = `Stale · ${formatSyncTime(syncedAt)}`;
      } else if (flightsReady && !hasAnyFlightCache) {
        window._readinessState.missions = 'ok';
        _setReadinessDot_('rdot-missions', 'ok');
        if (missionLabel) missionLabel.textContent = 'No flights scheduled';
      } else if (flightsReady) {
        window._readinessState.missions = 'ok';
        _setReadinessDot_('rdot-missions', 'ok');
        if (missionLabel) missionLabel.textContent = `${cachedCount} cached`;
      } else {
        window._readinessState.missions = 'bad';
        _setReadinessDot_('rdot-missions', 'bad');
        if (missionLabel) missionLabel.textContent = `${cachedCount}/${total || cachedCount} cached`;
      }

      window._readinessState.cg = hasEnvelope ? 'ok' : 'bad';
      _setReadinessDot_('rdot-cg', hasEnvelope ? 'ok' : 'bad');
      if (cgLabel) cgLabel.textContent = hasEnvelope ? `${envelopeCount} envelope${envelopeCount !== 1 ? 's' : ''}` : 'Not cached';

      const unknownPack = _computeUnknownPackReadiness_();
      window._readinessState.unknown = unknownPack.ready ? 'ok' : 'bad';
      _setReadinessDot_('rdot-unknown', unknownPack.ready ? 'ok' : 'bad');
      if (unknownLabel) unknownLabel.textContent = unknownPack.label;

      _recomputeOverallReadiness_();
      updateOfflineFlightButtonState();
    }

    function clearOfflineCacheNow() {
      const ok = window.confirm('Clear mission offline data (missions, envelopes, W&B/performance cache, and queue)? Maps will be kept.');
      if (!ok) return;
      try {
        const keys = Object.keys(localStorage || {});
        keys.forEach(function(key) {
          const k = String(key || '');
          if (k.startsWith('mba_cache_') || k.startsWith('mission_') || k.startsWith('mba_autolog_ldglog_') || k === OFFLINE_CACHE_KEYS.OUTBOX || k.startsWith('mba_outbox_')) {
            localStorage.removeItem(k);
          }
        });
      } catch (e) {}
      updateOfflineCacheStatus({ state: 'empty', cachedCount: 0, total: 0, envelopeCount: 0 });
      updateOutboxStatusPanel();
      updateOutboxButtonState();
      updateOfflineFlightButtonState();
      if (window.M) M.toast({ html: 'Mission offline data cleared (maps kept)', classes: 'orange' });
    }

    async function clearMapsCacheNow() {
      let ok = false;
      if (typeof window.flightAppConfirm === 'function') {
        ok = await window.flightAppConfirm('Are you sure you want to clear all cached map tiles?', { title: 'Clear Maps', okText: 'Clear Maps' });
      } else {
        ok = window.confirm('Are you sure you want to clear all cached map tiles?');
      }
      if (!ok) return;

      if (typeof window.clearMapPackDirect === 'function') {
        await window.clearMapPackDirect();
      } else {
        try { localStorage.removeItem('enrouteMapPackMeta'); } catch (e) {}
      }

      if (typeof window.refreshOfflineMapPackStatus === 'function') {
        window.refreshOfflineMapPackStatus();
      }
      if (window.M) M.toast({ html: 'Map cache cleared', classes: 'orange' });
    }

    function ensureEnvelopeCacheOnStartup() {
      if (!isServerAvailable()) {
        updateOfflineCacheStatus();
        updateOfflineFlightButtonState();
        return;
      }
      if (!(window.google && google.script && google.script.run)) return;

      // Respect 24h freshness on cold start so reopen does not force a full sync.
      if (isOfflineCacheFresh_()) {
        updateOfflineCacheStatus();
        updateOfflineFlightButtonState();
        return;
      }

      syncEnvelopeCacheForAircraftRegs_(_knownAircraftRegsForEnvelopeSync_(), { forceRefresh: false });

      google.script.run
        .withSuccessHandler(function(missions) {
          const list = Array.isArray(missions) ? missions : [];
          const regs = _knownAircraftRegsFromMissions_(list);
          if (regs.length) {
            syncEnvelopeCacheForAircraftRegs_(regs, { forceRefresh: false });
          }
          if (!list.length) {
            updateOfflineCacheStatus();
            return;
          }
          prefetchScheduledMissionDetails(list, false);
          setTimeout(function() {
            updateOfflineCacheStatus();
            updateOfflineFlightButtonState();
          }, 1800);
        })
        .withFailureHandler(function() {
          updateOfflineCacheStatus();
          updateOfflineFlightButtonState();
        })
        .getScheduledMissions();
    }

    const MAP_PACK_STALE_DAYS = 15;

    // ============================================================
    // OFFLINE MAP PACK ENGINE — available from startup, Tab 1+
    // ============================================================
    (function() {
      const _IDB_NAME  = 'enrouteMapPackDB';
      const _IDB_STORE = 'tiles';
      let   _idbInst   = null;
      let   _packDownloadActive = false;

      function _idbOpen() {
        return new Promise(function(resolve, reject) {
          if (_idbInst) { resolve(_idbInst); return; }
          const req = indexedDB.open(_IDB_NAME, 1);
          req.onupgradeneeded = function(e) {
            const db = e.target.result;
            if (!db.objectStoreNames.contains(_IDB_STORE)) db.createObjectStore(_IDB_STORE);
          };
          req.onsuccess = function(e) { _idbInst = e.target.result; resolve(_idbInst); };
          req.onerror   = function(e) { reject(e.target.error); };
        });
      }
      function _idbPut(key, val) {
        return _idbOpen().then(function(db) {
          return new Promise(function(resolve, reject) {
            const tx = db.transaction(_IDB_STORE, 'readwrite');
            tx.objectStore(_IDB_STORE).put(val, key);
            tx.oncomplete = resolve;
            tx.onerror    = function(e) { reject(e.target.error); };
          });
        });
      }
      function _idbGet(key) {
        return _idbOpen().then(function(db) {
          return new Promise(function(resolve, reject) {
            const tx  = db.transaction(_IDB_STORE, 'readonly');
            const req = tx.objectStore(_IDB_STORE).get(key);
            req.onsuccess = function(e) { resolve(e.target.result || null); };
            req.onerror   = function(e) { reject(e.target.error); };
          });
        });
      }
      function _idbClearAll() {
        return _idbOpen().then(function(db) {
          return new Promise(function(resolve) {
            try {
              const tx = db.transaction(_IDB_STORE, 'readwrite');
              tx.objectStore(_IDB_STORE).clear();
              tx.oncomplete = resolve; tx.onerror = resolve;
            } catch(e) { resolve(); }
          });
        }).catch(function() {});
      }

      // Expose IDB helpers globally so Tab 6 tile rendering can use them
      window._mapIdb = {
        get:       function(k)          { return _idbGet(k); },
        put:       function(k, v)       { return _idbPut(k, v); },
        getTile:   function(z, x, y)    { return _idbGet('t_'+z+'_'+x+'_'+y).catch(function(){return null;}); },
        putTile:   function(z, x, y, b) { return _idbPut('t_'+z+'_'+x+'_'+y, b); },
        getMeta:   function()           { return _idbGet('meta_pack').catch(function(){return null;}); },
        putMeta:   function(m)          { return _idbPut('meta_pack', m).catch(function(){}); },
        clearAll:  function()           { return _idbClearAll(); }
      };

      // Tile coordinate math
      function _latLonToTile(lat, lon, zoom) {
        const n = Math.pow(2, zoom);
        const x = Math.floor((lon + 180) / 360 * n);
        const lRad = lat * Math.PI / 180;
        const y = Math.floor((1 - Math.log(Math.tan(lRad) + 1/Math.cos(lRad)) / Math.PI) / 2 * n);
        return { x: Math.max(0, Math.min(n-1, x)), y: Math.max(0, Math.min(n-1, y)) };
      }
      function _enumTiles(minLat, minLon, maxLat, maxLon, minZ, maxZ) {
        const tiles = [];
        for (let z = minZ; z <= maxZ; z++) {
          const tl = _latLonToTile(maxLat, minLon, z);
          const br = _latLonToTile(minLat, maxLon, z);
          for (let x = tl.x; x <= br.x; x++)
            for (let y = tl.y; y <= br.y; y++)
              tiles.push({ z:z, x:x, y:y });
        }
        return tiles;
      }
      function _enumTilesForBboxes(bboxes, minZ, maxZ) {
        const seen = new Set(), out = [];
        (Array.isArray(bboxes) ? bboxes : []).forEach(function(b) {
          if (!b) return;
          const minLat=Number(b.minLat), minLon=Number(b.minLon), maxLat=Number(b.maxLat), maxLon=Number(b.maxLon);
          if (![minLat,minLon,maxLat,maxLon].every(isFinite)) return;
          _enumTiles(minLat,minLon,maxLat,maxLon,minZ,maxZ).forEach(function(t) {
            const k=t.z+'_'+t.x+'_'+t.y;
            if (!seen.has(k)) { seen.add(k); out.push(t); }
          });
        });
        return out;
      }
      function _tileUrlFromTemplate(template, z, x, y) {
        const s = ['a','b','c'][Math.floor(Math.random()*3)];
        return template.replace('{z}',z).replace('{x}',x).replace('{y}',y).replace('{s}',s);
      }
      function _getTileTemplate() {
        // Tab 6 may store its chosen template; fall back to OSM
        return (window.enroute6 && window.enroute6.tileTemplate) ||
               String(localStorage.getItem('enrouteTileTemplate') || '').trim() ||
               'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png';
      }

      // Coverage bboxes
      function _offlineWacBbox() {
        return { minLat: -34.0, maxLat: 6.0, minLon: -74.0, maxLon: -28.0 };
      }
      function _offlineAirportCoverageBboxes(bufferDeg) {
        const buf = Math.max(0.08, Number(bufferDeg || 0.2));
        const airports = (window.appData && Array.isArray(window.appData.airports)) ? window.appData.airports : [];
        const out = airports.map(function(apt) {
          const lat=Number(apt&&apt.lat), lon=Number(apt&&apt.lon);
          if (!isFinite(lat)||!isFinite(lon)) return null;
          return { minLat:Math.max(-85,lat-buf), maxLat:Math.min(85,lat+buf), minLon:Math.max(-180,lon-buf), maxLon:Math.min(180,lon+buf) };
        }).filter(Boolean);
        return out.length ? out : [_offlineWacBbox()];
      }

      // Status display
      function _packDateLabel(isoDate) {
        const d = new Date(isoDate||'');
        if (isNaN(d.getTime())) return '';
        const mo = ['Jan.','Feb.','Mar.','Apr.','May.','Jun.','Jul.','Aug.','Sep.','Oct.','Nov.','Dec.'][d.getMonth()]||'---';
        return mo+' '+d.getDate()+', '+d.getFullYear();
      }
      function _setOfflineMapDbIndicator(state, label) {
        const dot  = document.getElementById('rdot-mapdb');
        const text = document.getElementById('rlabel-mapdb');
        if (text) text.textContent = label || 'Not downloaded';
        if (dot) {
          const colors = { none:'#bdbdbd', ok:'#2e7d32', warn:'#f9a825', expired:'#d32f2f', syncing:'#1565c0' };
          dot.style.background = colors[state] || colors.none;
        }
        const stateMap = { ok:'ok', syncing:'sync', expired:'bad', warn:'warn', none:'none' };
        if (window._readinessState) window._readinessState.mapdb = stateMap[state] || 'none';
        if (typeof window._recomputeOverallReadiness_ === 'function') window._recomputeOverallReadiness_();
      }
      window._setOfflineMapDbIndicator = _setOfflineMapDbIndicator;

      function _renderPackStatusFromMeta(meta) {
        const els = [
          document.getElementById('enroute-mappack-status'),
          document.getElementById('offline-map-status')
        ].filter(Boolean);
        if (!meta || !meta.savedAt) {
          els.forEach(function(el){ el.textContent='Maps not downloaded.'; el.className='pack-status pack-none'; });
          _setOfflineMapDbIndicator('none', 'Not downloaded');
          return;
        }
        const ageDays = Math.floor((Date.now()-new Date(meta.savedAt).getTime())/86400000);
        const dateLabel = _packDateLabel(meta.savedAt);
        if (ageDays > MAP_PACK_STALE_DAYS) {
          els.forEach(function(el){ el.textContent='✖ MAPS EXPIRED — last download '+(dateLabel||'unknown date'); el.className='pack-status pack-expired'; });
          _setOfflineMapDbIndicator('expired', 'Expired');
        } else {
          els.forEach(function(el){ el.textContent='✔ MAPS DOWNLOADED — '+(dateLabel||'date unavailable'); el.className='pack-status pack-ok'; });
          _setOfflineMapDbIndicator('ok', 'Ready — '+dateLabel);
        }
      }

      function _updatePackStatus() {
        try {
          const metaRaw = localStorage.getItem('enrouteMapPackMeta');
          if (metaRaw) {
            let meta; try { meta = JSON.parse(metaRaw); } catch(e) { meta = null; }
            _renderPackStatusFromMeta(meta);
            return;
          }
          const els = [document.getElementById('enroute-mappack-status'), document.getElementById('offline-map-status')].filter(Boolean);
          els.forEach(function(el){ el.textContent='Maps not downloaded.'; el.className='pack-status pack-none'; });
          _setOfflineMapDbIndicator('none', 'Not downloaded');
        } catch (e) {
          _setOfflineMapDbIndicator('warn', 'Status unavailable');
        }
      }
      window.refreshOfflineMapPackStatus = _updatePackStatus;

      // Download progress UI (Tab 1 panel elements + Tab 6 elements, whichever exist)
      function _progEl(id) { return document.getElementById(id); }
      function _updateDownloadProgress(done, failed, total) {
        const pct = Math.round((done+failed)/total*100);
        ['enroute-pack-bar',  'tab1-map-pack-bar' ].forEach(function(id){ const el=_progEl(id); if(el) el.style.width=pct+'%'; });
        ['enroute-pack-progress-text','tab1-map-pack-text'].forEach(function(id){ const el=_progEl(id); if(el) el.textContent=(done+failed)+'/'+total+' tiles ('+failed+' failed)'; });
      }

      // Main download function — works from any tab
      window.downloadMapPack = async function(scope) {
        if (_packDownloadActive) {
          if (window.M) M.toast({ html: 'Maps downloading — see progress below', classes: 'blue darken-2' });
          return;
        }
        const isRoute = String(scope||'base').toLowerCase() === 'route';
        const defaults = isRoute ? { minZoom:7, maxZoom:11, label:'Route pack' } : { minZoom:6, maxZoom:10, label:'Airport coverage pack' };
        // For route scope we need Tab 6 waypoints — fall back to base if not available
        let bboxes;
        if (isRoute && window.enroute6 && Array.isArray(window.enroute6.waypoints) && window.enroute6.waypoints.length) {
          const wps = window.enroute6.waypoints;
          const coords = wps.map(function(wp){
            if (typeof window._lookupFix === 'function') return window._lookupFix(wp.fix);
            const airports = window.appData && window.appData.airports || [];
            const waypoints = window.appData && window.appData.waypoints || [];
            return [...airports,...waypoints].find(function(a){ return (a.icao||a.fix||a.id||'').toUpperCase() === String(wp.fix||'').toUpperCase(); }) || null;
          }).filter(Boolean);
          if (coords.length) {
            const lats=coords.map(function(c){return c.lat;}), lons=coords.map(function(c){return c.lon;});
            bboxes = [{ minLat:Math.min.apply(null,lats)-0.5, maxLat:Math.max.apply(null,lats)+0.5, minLon:Math.min.apply(null,lons)-0.5, maxLon:Math.max.apply(null,lons)+0.5 }];
          }
        }
        if (!bboxes || !bboxes.length) bboxes = _offlineAirportCoverageBboxes(0.2);
        if (!bboxes.length) { if (window.M) M.toast({ html: 'No airport data loaded yet — sync missions first', classes: 'orange' }); return; }

        const template = _getTileTemplate();
        const minZ = defaults.minZoom, maxZ = defaults.maxZoom;
        const tiles = _enumTilesForBboxes(bboxes, minZ, maxZ);
        if (tiles.length > 2000) {
          const ok = await window.flightAppConfirm(defaults.label+' download: ~'+tiles.length+' tiles (~'+Math.round(tiles.length*22/1024)+' MB). Proceed?', { title:'Offline Map Pack', okText:'Download' });
          if (!ok) return;
        }

        _packDownloadActive = true;
        _setOfflineMapDbIndicator('syncing', 'Downloading…');
        // Show progress in whichever panel is visible
        ['enroute-pack-progress','tab1-map-pack-progress'].forEach(function(id){ const el=_progEl(id); if(el) el.style.display='block'; });
        ['enroute-pack-btn','tab1-map-pack-btn'].forEach(function(id){ const el=_progEl(id); if(el){ el.textContent='DOWNLOADING…'; el.disabled=true; } });

        let done=0, failed=0;
        const BATCH=6;
        for (let i=0; i<tiles.length; i+=BATCH) {
          if (!_packDownloadActive) break;
          await Promise.all(tiles.slice(i,i+BATCH).map(async function(t) {
            try {
              const url = _tileUrlFromTemplate(template, t.z, t.x, t.y);
              const resp = await fetch(url, { cache:'no-store' });
              if (!resp.ok) throw new Error('HTTP '+resp.status);
              const blob = await resp.blob();
              await window._mapIdb.putTile(t.z, t.x, t.y, blob);
              done++;
            } catch(e) { failed++; }
            _updateDownloadProgress(done, failed, tiles.length);
          }));
        }

        _packDownloadActive = false;
        ['enroute-pack-btn','tab1-map-pack-btn'].forEach(function(id){ const el=_progEl(id); if(el){ el.textContent='SYNC MAPS'; el.disabled=false; } });
        const meta = {
          scope: String(scope||'base').toLowerCase(),
          savedAt: new Date().toISOString(),
          tileCount: done, failedCount: failed,
          bboxes: bboxes, minZoom: minZ, maxZoom: maxZ,
          tileTemplate: template
        };
        localStorage.setItem('enrouteMapPackMeta', JSON.stringify(meta));
        window._mapIdb.putMeta(meta);
        _updatePackStatus();
        if (window.M) M.toast({
          html: '\uD83D\uDDFA '+defaults.label+' ready: '+done+' tiles'+(failed?' ('+failed+' failed)':''),
          classes: done>0?'green':'orange', displayLength:6000
        });
        // If Tab 6 map is already rendered, refresh its tile layer
        if (done>0 && typeof window._enrouteApplyOfflineTileLayer === 'function') window._enrouteApplyOfflineTileLayer();
      };

      window.clearMapPack = async function() {
        const ok = await window.flightAppConfirm('Clear all cached map tiles?', { title:'Offline Map Pack', okText:'Clear' });
        if (!ok) return;
        await _idbClearAll();
        localStorage.removeItem('enrouteMapPackMeta');
        _packDownloadActive = false;
        _updatePackStatus();
        if (typeof window._enrouteApplyOfflineTileLayer === 'function') window._enrouteApplyOfflineTileLayer();
        if (window.M) M.toast({ html:'Map pack cleared', classes:'orange' });
      };
      window.clearMapPackDirect = async function() {
        await _idbClearAll();
        localStorage.removeItem('enrouteMapPackMeta');
        _packDownloadActive = false;
        _updatePackStatus();
      };
    })();
    // ============================================================
    // END OFFLINE MAP PACK ENGINE
    // ============================================================

    function getOfflinePrefetchMeta_() {
      return cacheGet(OFFLINE_CACHE_KEYS.PREFETCH_META) || {};
    }

    function isOfflineCacheFresh_() {
      const meta = getOfflinePrefetchMeta_();
      const syncedAt = meta && meta.syncedAt ? new Date(meta.syncedAt).getTime() : 0;
      if (!syncedAt || Number.isNaN(syncedAt)) return false;
      // Use a 30-minute gate for auto-sync on open so missions are always fresh,
      // but rapid reopens within 30 min don't hammer the server.
      return (Date.now() - syncedAt) <= (30 * 60 * 1000);
    }

    function _isMapPackStale_() {
      try {
        const raw = localStorage.getItem('enrouteMapPackMeta');
        if (!raw) return true;
        const meta = JSON.parse(raw);
        if (!meta || !meta.savedAt) return true;
        const ageDays = (Date.now() - new Date(meta.savedAt).getTime()) / 86400000;
        return ageDays > MAP_PACK_STALE_DAYS;
      } catch (e) { return true; }
    }

    function _maybeRefreshMapPackNow_() {
      // Maps are only downloaded manually via the ↺ MAPS button.
      // Just refresh the displayed status here — never auto-download.
      if (typeof window.refreshOfflineMapPackStatus === 'function') {
        window.refreshOfflineMapPackStatus();
      }
    }

    function downloadOfflineCacheNow() {
      refreshOfflineCacheNow();
    }

    // Explicitly expose handlers for inline onclick reliability across webview/browser variants.
    window.downloadOfflineCacheNow = downloadOfflineCacheNow;
    window.clearOfflineCacheNow = clearOfflineCacheNow;
    window.clearMapsCacheNow = clearMapsCacheNow;

    function refreshOfflineCacheNow() {
      if (!isServerAvailable()) {
        M.toast({ html: 'Offline: unable to refresh cache right now', classes: 'orange' });
        updateOfflineCacheStatus();
        return;
      }
      const btn = document.getElementById('btn-refresh-offline-cache');
      if (btn) {
        btn.disabled = true;
        btn.textContent = 'REFRESHING...';
      }

      updateOfflineCacheStatus({ state: 'syncing', progress: 0, total: 100, cachedCount: getCachedMissionCount(), message: 'Starting download...' });

      const finalizeDownloadSuccess_ = function() {
        setTimeout(function() {
          updateOfflineCacheStatus({ state: 'ready' });
          if (btn) {
            btn.disabled = false;
            btn.textContent = 'REFRESH';
          }
        }, 1200);
      };

      const finalizeDownloadFailure_ = function(message) {
        const msg = message || 'unknown error';
        M.toast({ html: `Download failed: ${msg}`, classes: 'red' });
        updateOfflineCacheStatus();
        if (btn) {
          btn.disabled = false;
          btn.textContent = 'REFRESH';
        }
      };

      const downloadScheduledMissions_ = function() {
        updateOfflineCacheStatus({ state: 'syncing', progress: 25, total: 100, cachedCount: getCachedMissionCount(), message: 'Loading mission list...' });
        google.script.run
          .withSuccessHandler(function(missions) {
            const list = Array.isArray(missions) ? missions : [];
            cacheSet(OFFLINE_CACHE_KEYS.SCHEDULED_MISSIONS, list);
            updateOfflineCacheStatus({ state: 'syncing', progress: 35, total: 100, cachedCount: getCachedMissionCount(), message: `Caching ${list.length} mission details...` });
            prefetchScheduledMissionDetails(list, true);
            syncEnvelopeCacheForAircraftRegs_(_knownAircraftRegsForEnvelopeSync_().concat(_knownAircraftRegsFromMissions_(list)), { forceRefresh: true });
            // Cache all aircraft docs for all tails — works offline, survives last-minute aircraft switch
            google.script.run
              .withSuccessHandler(function(docsResult) {
                if (docsResult && docsResult.success) cacheSet(OFFLINE_CACHE_KEYS.AIRCRAFT_DOCS, docsResult);
              })
              .withFailureHandler(function() { /* silent — docs cache best-effort */ })
              .getAircraftDocsForTools('');
            finalizeDownloadSuccess_();
          })
          .withFailureHandler(function(err) {
            const msg = err && err.message ? err.message : String(err || 'unknown error');
            finalizeDownloadFailure_(msg);
          })
          .getScheduledMissions();
      };

      // Pull full startup dropdown data so offline flow includes DB_Routes and other core tables.
      google.script.run
        .withSuccessHandler(function(startupData) {
          updateOfflineCacheStatus({ state: 'syncing', progress: 15, total: 100, cachedCount: getCachedMissionCount(), message: 'Loading core tables...' });
          const payload = (startupData && typeof startupData === 'object') ? startupData : {};
          if (payload.error) {
            M.toast({ html: 'Startup data refresh warning. Continuing with mission download...', classes: 'orange' });
            downloadScheduledMissions_();
            return;
          }

          const cachedDropdown = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA) || {};
          const mergedDropdown = Object.assign({}, cachedDropdown, payload);
          cacheSet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA, mergedDropdown);
          updateOfflineCacheStatus({ state: 'syncing', progress: 22, total: 100, cachedCount: getCachedMissionCount(), message: 'Core tables cached' });

          if (window.appData && typeof window.appData === 'object') {
            window.appData = Object.assign({}, window.appData, mergedDropdown);
          } else {
            window.appData = mergedDropdown;
          }
          // Keep local module state in sync with window.appData.
          appData = window.appData;

          downloadScheduledMissions_();
        })
        .withFailureHandler(function() {
          M.toast({ html: 'Core cache refresh skipped (network). Downloading missions only...', classes: 'orange' });
          downloadScheduledMissions_();
        })
        .getPilotStartupData();
    }

    function missionCacheKey(missionId) {
      return `mba_cache_mission_${String(missionId || '').trim()}`;
    }

    function isServerAvailable() {
      return Boolean(window.google && google.script && google.script.run && navigator.onLine);
    }

    function updateConnectivityBanner() {
      const banner = document.getElementById('connectivity-banner');
      if (!banner) return;
      if (isServerAvailable()) {
        banner.style.display = 'none';
        return;
      }
      banner.style.display = 'block';
      banner.style.background = '#fff3cd';
      banner.style.color = '#7a5b00';
      banner.textContent = 'OFFLINE MODE: showing cached data';
    }

    function fetchMissionDetails(missionId, onSuccess, onFailure) {
      const id = String(missionId || '').trim();
      if (!id) {
        if (typeof onFailure === 'function') onFailure(new Error('Missing missionId'));
        return;
      }

      if (isServerAvailable()) {
        google.script.run
          .withSuccessHandler(mission => {
            if (mission) {
              cacheSet(missionCacheKey(id), mission);
              window.currentBriefingMission = mission;
              prefetchWbForMission(mission, false);
              prefetchPerformanceForMission(mission, false);
              _activeLessonContext = _buildLessonContextFromMission_(mission);
              renderLessonLinkBar_();
            }
            if (typeof onSuccess === 'function') onSuccess(mission);
          })
          .withFailureHandler(err => {
            const cachedMission = cacheGet(missionCacheKey(id));
            if (cachedMission) {
              M.toast({ html: 'Server unavailable. Loaded cached mission.', classes: 'orange' });
              window.currentBriefingMission = cachedMission;
              _activeLessonContext = _buildLessonContextFromMission_(cachedMission);
              renderLessonLinkBar_();
              if (typeof onSuccess === 'function') onSuccess(cachedMission);
              return;
            }
            if (typeof onFailure === 'function') onFailure(err);
          })
          .getMissionById(id);
        return;
      }

      const cachedMission = cacheGet(missionCacheKey(id));
      if (cachedMission) {
        window.currentBriefingMission = cachedMission;
        _activeLessonContext = _buildLessonContextFromMission_(cachedMission);
        renderLessonLinkBar_();
        if (typeof onSuccess === 'function') onSuccess(cachedMission);
        return;
      }

      if (typeof onFailure === 'function') onFailure(new Error('Offline and no cached mission data available.'));
    }

    function _tab1MissionForAccept_(missionId, onSuccess, onFailure) {
      const id = String(missionId || '').trim();
      const current = window.currentBriefingMission;
      if (current && String(current.id || '').trim() === id) {
        if (typeof onSuccess === 'function') onSuccess(current);
        return;
      }

      const cachedMission = cacheGet(missionCacheKey(id));
      if (cachedMission) {
        window.currentBriefingMission = cachedMission;
        if (typeof onSuccess === 'function') onSuccess(cachedMission);
        return;
      }

      fetchMissionDetails(id, onSuccess, onFailure);
    }

    function _tab1BuildOzRunwaysPayload_(mission) {
      const firstLeg = mission && Array.isArray(mission.legs) ? (mission.legs[0] || {}) : {};
      const origin = String(firstLeg.from || '').trim().toUpperCase();
      const destination = String(firstLeg.to || '').trim().toUpperCase();
      const raw = String(firstLeg.waypoints || '').trim();
      let waypoints = raw
        ? raw.split(',').map(function(part) { return String(part || '').trim().toUpperCase(); }).filter(Boolean)
        : [];

      if (origin && (!waypoints.length || waypoints[0] !== origin)) waypoints.unshift(origin);
      if (destination && (!waypoints.length || waypoints[waypoints.length - 1] !== destination)) waypoints.push(destination);
      waypoints = waypoints.filter(function(token, idx) {
        return idx === 0 || token !== waypoints[idx - 1];
      });

      return {
        missionId: String(mission && mission.id || '').trim(),
        flightId: String(firstLeg.flightLegId || mission && mission.id || '').trim(),
        aircraft: String(mission && mission.acft || '').trim().toUpperCase(),
        pilot: String(mission && mission.pilot || '').trim(),
        origin: origin,
        destination: destination,
        waypoints: waypoints,
        route: waypoints.join(' ')
      };
    }

    // ── Garmin .fpl XML builder ───────────────────────────────────────────
    function _tab1BuildFplXml_(payload) {
      const airports = (window.appData && window.appData.airports) || {};
      const wpts = Array.isArray(payload.waypoints) ? payload.waypoints : [];
      const created = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');

      function ex(s) {
        return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
      }
      function isApt(id) { return /^[A-Z]{4}$/.test(id); }

      const wtTable = wpts.map(function(icao) {
        const a = airports[icao] || {};
        const lat = a.lat != null ? parseFloat(a.lat).toFixed(6) : '0.000000';
        const lon = a.lon != null ? parseFloat(a.lon).toFixed(6) : '0.000000';
        return [
          '    <waypoint>',
          '      <identifier>' + ex(icao) + '</identifier>',
          '      <type>' + (isApt(icao) ? 'AIRPORT' : 'USER WAYPOINT') + '</type>',
          '      <country-code>BR</country-code>',
          '      <lat>' + lat + '</lat>',
          '      <lon>' + lon + '</lon>',
          '      <comment></comment>',
          '    </waypoint>'
        ].join('\n');
      }).join('\n');

      const rtPts = wpts.map(function(icao) {
        return [
          '    <route-point>',
          '      <waypoint-identifier>' + ex(icao) + '</waypoint-identifier>',
          '      <waypoint-type>' + (isApt(icao) ? 'AIRPORT' : 'USER WAYPOINT') + '</waypoint-type>',
          '      <waypoint-country-code>BR</waypoint-country-code>',
          '    </route-point>'
        ].join('\n');
      }).join('\n');

      return [
        '<' + '?xml version="1.0" encoding="UTF-8"?>',
        '<flight-plan xmlns="http://www8.garmin.com/xmlschemas/FlightPlan/v1">',
        '  <created>' + created + '</created>',
        '  <waypoint-table>',
        wtTable,
        '  </waypoint-table>',
        '  <route>',
        '    <route-name>' + ex(payload.missionId || 'MISSION') + '</route-name>',
        '    <flight-plan-index>1</flight-plan-index>',
        rtPts,
        '  </route>',
        '</flight-plan>'
      ].join('\n');
    }

    // ── Save .fpl to Drive, return promise({ok, downloadUrl}) ────────────────
    function _tab1SaveFplToDrive_(missionId, fplXml) {
      return new Promise(function(resolve) {
        if (!isServerAvailable()) { resolve(null); return; }
        google.script.run
          .withSuccessHandler(function(r) { resolve(r && r.ok ? r : null); })
          .withFailureHandler(function(e) {
            console.warn('saveMissionFplToDrive failed', e);
            resolve(null);
          })
          .saveMissionFplToDrive(missionId, fplXml);
      });
    }

    let __tab1AcceptBusy = false;

    function tab1AcceptMissionThenAdvance() {
      if (__tab1AcceptBusy) return;
      if (!activeMission) {
        M.toast({ html: 'Please select a mission first!', classes: 'orange' });
        return;
      }

      __tab1AcceptBusy = true;
      const btn = document.getElementById('btnNext');
      const prevLabel = btn ? btn.innerText : 'ACCEPT MISSION';
      if (btn) {
        btn.disabled = true;
        btn.innerText = 'OPENING FLIGHT PLAN...';
      }

      const finish = function() {
        __tab1AcceptBusy = false;
        if (btn) {
          btn.disabled = false;
          btn.innerText = prevLabel;
        }
        switchTab(2);
      };

      _tab1MissionForAccept_(activeMission, function(mission) {
        M.toast({ html: 'Mission accepted', classes: 'green', displayLength: 2000 });
        finish();
      }, function(err) {
        console.warn('Mission load for accept failed', err);
        M.toast({ html: 'Mission accepted', classes: 'orange', displayLength: 2500 });
        finish();
      });
    }

    function prefetchWbForMission(mission, forceRefresh) {
      if (!isServerAvailable()) return;
      if (!(window.google && google.script && google.script.run)) return;

      const firstLegId = mission && mission.legs && mission.legs[0]
        ? String(mission.legs[0].flightLegId || '').trim()
        : '';
      if (!firstLegId) return;

      try {
        if (!forceRefresh && typeof window.readCachedWBPayload_ === 'function' && window.readCachedWBPayload_(firstLegId)) {
          return;
        }
      } catch (e) {}

      google.script.run
        .withSuccessHandler(function(wb) {
          try {
            if (wb && typeof window.cacheWBPayload_ === 'function') {
              window.cacheWBPayload_(firstLegId, wb);
            }
            if (wb) {
              const aircraftReg = String((wb && wb.aircraft) || (mission && mission.acft) || '').trim().toUpperCase();
              const env = Array.isArray(wb && wb.envelopeData) ? wb.envelopeData : [];
              if (aircraftReg && env.length >= 3) {
                localStorage.setItem('mba_cache_envelope_' + aircraftReg, JSON.stringify({
                  aircraft: aircraftReg,
                  cachedAt: new Date().toISOString(),
                  envelopeData: env
                }));
                updateOfflineCacheStatus();
                updateOfflineFlightButtonState();
              }
            }
          } catch (e) {
            console.warn('prefetchWbForMission cache save failed', e);
          }
        })
        .withFailureHandler(function(err) {
          console.warn('prefetchWbForMission failed', err);
        })
        .initializeWB(firstLegId);
    }

    function prefetchPerformanceForMission(mission, forceRefresh) {
      if (!isServerAvailable()) return;
      if (!(window.google && google.script && google.script.run)) return;

      const icaos = Array.from(new Set(
        (Array.isArray(mission && mission.legs) ? mission.legs : [])
          .flatMap(function(leg) {
            return [
              String(leg && leg.from || '').trim().toUpperCase(),
              String(leg && leg.to || '').trim().toUpperCase()
            ];
          })
          .filter(Boolean)
      ));

      if (!icaos.length) return;

      icaos.forEach(function(icao) {
        try {
          if (!forceRefresh && typeof window.readCachedPerformanceSetup_ === 'function' && window.readCachedPerformanceSetup_(icao)) {
            return;
          }
        } catch (e) {}

        google.script.run
          .withSuccessHandler(function(setup) {
            try {
              if (setup && typeof window.cachePerformanceSetup_ === 'function') {
                window.cachePerformanceSetup_(icao, setup);
              }
            } catch (e) {
              console.warn('prefetchPerformanceForMission cache save failed', e);
            }
          })
          .withFailureHandler(function(err) {
            console.warn('prefetchPerformanceForMission failed', err);
          })
          .getPerformanceSetup(icao);
      });
    }

    function prefetchScheduledMissionDetails(missions, forceRefresh) {
      if (!isServerAvailable()) return;
      if (hasPrefetchedMissionsThisSession && !forceRefresh) return;

      const missionIds = Array.from(new Set(
        (Array.isArray(missions) ? missions : [])
          .map(m => String((m && m.id) || '').trim())
          .filter(Boolean)
      ));

      if (!missionIds.length) return;

      hasPrefetchedMissionsThisSession = true;
      missionPrefetchInProgress = true;
      missionPrefetchProgress = { total: missionIds.length, progress: 0 };

      let index = 0;
      let cachedCount = 0;
      let failedCount = 0;

      M.toast({ html: `Syncing ${missionIds.length} flights for offline use...`, classes: 'blue darken-2' });
      updateOfflineCacheStatus({ state: 'syncing', total: missionIds.length, progress: 0, cachedCount: getCachedMissionCount(), message: `Syncing flight details (0/${missionIds.length})...` });

      const step = function() {
        if (index >= missionIds.length) {
          missionPrefetchInProgress = false;
          missionPrefetchProgress = { total: missionIds.length, progress: missionIds.length };
          const syncedAtIso = new Date().toISOString();
          cacheSet(OFFLINE_CACHE_KEYS.PREFETCH_META, {
            syncedAt: syncedAtIso,
            total: missionIds.length,
            cached: cachedCount,
            failed: failedCount
          });

          updateOfflineCacheStatus({
            state: failedCount ? 'partial' : 'ready',
            total: missionIds.length,
            cachedCount: getCachedMissionCount(),
            syncedAt: syncedAtIso
          });

          // Queue a base-pack map download if maps have never been cached.
          // The download itself runs the first time Tab 6 opens (code lives there).
          if (!localStorage.getItem('enrouteMapPackMeta')) {
            window._pendingMapPackDownload = 'base';
            const mapLbl = document.getElementById('rlabel-mapdb');
            if (mapLbl) mapLbl.textContent = 'Will download when Tab 6 opens';
          }

          const summaryClass = failedCount ? 'orange' : 'green';
          M.toast({
            html: `Offline sync complete: ${cachedCount}/${missionIds.length} flights cached${failedCount ? ` (${failedCount} failed)` : ''}`,
            classes: summaryClass,
            displayLength: 5000
          });
          return;
        }

        const id = missionIds[index++];
        google.script.run
          .withSuccessHandler(mission => {
            if (mission && cacheSet(missionCacheKey(id), mission)) {
              cachedCount += 1;
            } else {
              failedCount += 1;
            }
            missionPrefetchProgress = { total: missionIds.length, progress: index };
            if (mission) {
              prefetchWbForMission(mission, !!forceRefresh);
              prefetchPerformanceForMission(mission, !!forceRefresh);
            }
            updateOfflineCacheStatus({
              state: 'syncing',
              total: missionIds.length,
              progress: index,
              cachedCount: getCachedMissionCount(),
              message: `Syncing flight details (${index}/${missionIds.length})...`
            });
            setTimeout(step, 35);
          })
          .withFailureHandler(() => {
            failedCount += 1;
            missionPrefetchProgress = { total: missionIds.length, progress: index };
            updateOfflineCacheStatus({
              state: 'syncing',
              total: missionIds.length,
              progress: index,
              cachedCount: getCachedMissionCount(),
              message: `Syncing flight details (${index}/${missionIds.length})...`
            });
            setTimeout(step, 80);
          })
          .getMissionById(id);
      };

      step();
    }

    function updateNewFlightButtonState() {
      const btn = document.getElementById('btn-new-flight');
      if (!btn) return;

      const online = isServerAvailable();
      btn.style.opacity = online ? '1' : '0.55';
      btn.style.filter = online ? 'none' : 'grayscale(0.2)';
      btn.title = online ? 'Open dispatch / flight creation portal' : 'Offline: opens local offline flight form';
    }

    function openNewFlightApp() {
      if (isServerAvailable()) {
        const targetUrl = _getDispatchPortalUrl_();
        try {
          if (window.top && window.top !== window) {
            window.top.location.assign(targetUrl);
          } else {
            window.location.assign(targetUrl);
          }
        } catch (e) {
          const opened = window.open(targetUrl, '_blank', 'noopener');
          if (opened) return;
          window.location.href = targetUrl;
        }
        return;
      }
      openOfflineFlightModal();
    }

    function _offfltNormIcao(v) {
      return String(v || '').trim().toUpperCase();
    }

    function _offfltFindAirport(icao) {
      const key = _offfltNormIcao(icao);
      const airportsPrimary = (appData && Array.isArray(appData.airports)) ? appData.airports : [];
      const cachedDropdown = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA) || {};
      const airportsFallback = Array.isArray(cachedDropdown.airports) ? cachedDropdown.airports : [];
      const findMatch = function(list) {
        return (Array.isArray(list) ? list : []).find(function(a) {
          const aptIcao = _offfltNormIcao((a && (a.icao || a.ICAO || a.code || a.CODE)) || '');
          return aptIcao === key;
        }) || null;
      };
      return findMatch(airportsPrimary) || findMatch(airportsFallback);
    }

    function _offfltFindAircraft(reg) {
      const key = String(reg || '').trim().toUpperCase();
      const aircraftPrimary = (appData && Array.isArray(appData.aircraft)) ? appData.aircraft : [];
      const cachedDropdown = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA) || {};
      const aircraftFallback = Array.isArray(cachedDropdown.aircraft) ? cachedDropdown.aircraft : [];
      const findMatch = function(list) {
        return (Array.isArray(list) ? list : []).find(function(a) {
          return String((a && a.reg) || '').trim().toUpperCase() === key;
        }) || null;
      };
      return findMatch(aircraftPrimary) || findMatch(aircraftFallback);
    }

    function _offfltAirportLatLon(airport) {
      if (!airport || typeof airport !== 'object') return null;
      const latRaw = airport.lat != null ? airport.lat : (airport.latitude != null ? airport.latitude : airport.LATITUDE);
      const lonRaw = airport.lon != null ? airport.lon : (airport.longitude != null ? airport.longitude : airport.LONGITUDE);
      const lat = typeof _parseCoordinate === 'function' ? _parseCoordinate(latRaw) : parseFloat(latRaw);
      const lon = typeof _parseCoordinate === 'function' ? _parseCoordinate(lonRaw) : parseFloat(lonRaw);
      if (!isFinite(lat) || !isFinite(lon)) return null;
      return { lat: lat, lon: lon };
    }

    function _offfltFindPilot(name) {
      const key = String(name || '').trim().toUpperCase();
      const pilotsPrimary = (appData && Array.isArray(appData.pilots)) ? appData.pilots : [];
      const cachedDropdown = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA) || {};
      const pilotsFallback = Array.isArray(cachedDropdown.pilots) ? cachedDropdown.pilots : [];
      const findMatch = function(list) {
        return (Array.isArray(list) ? list : []).find(function(p) {
          return String((p && p.name) || '').trim().toUpperCase() === key;
        }) || null;
      };
      return findMatch(pilotsPrimary) || findMatch(pilotsFallback);
    }

    function _offfltDistanceNm(fromAirport, toAirport) {
      if (!fromAirport || !toAirport) return 0;
      const a = _offfltAirportLatLon(fromAirport);
      const b = _offfltAirportLatLon(toAirport);
      if (!a || !b) return 0;
      return _offfltHaversineNm(a.lat, a.lon, b.lat, b.lon);
    }

    function _offfltHaversineNm(lat1, lon1, lat2, lon2) {
      if ([lat1, lon1, lat2, lon2].some(v => !isFinite(v))) return 0;

      const toRad = d => d * Math.PI / 180;
      const Rnm = 3440.065;
      const dLat = toRad(lat2 - lat1);
      const dLon = toRad(lon2 - lon1);
      const a = Math.sin(dLat / 2) * Math.sin(dLat / 2)
        + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
      const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
      return Math.max(0, Rnm * c);
    }

    function _offfltParseRouteTokens(routeStr) {
      const raw = String(routeStr || '').trim().toUpperCase();
      if (!raw) return [];
      const normalized = raw.replace(/[→>]/g, ',').replace(/\n|\r/g, ',');
      const tokens = normalized.split(',').map(s => s.trim()).filter(Boolean);
      return tokens.filter((tok, idx) => idx === 0 || tok !== tokens[idx - 1]);
    }

    function _offfltResolveLocation(token) {
      const key = String(token || '').trim().toUpperCase();
      if (!key) return null;

      const apt = _offfltFindAirport(key);
      if (apt) {
        const coords = _offfltAirportLatLon(apt);
        if (coords) {
          return { ident: key, lat: coords.lat, lon: coords.lon };
        }
      }

      const cachedDropdown = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA) || {};
      const waypointsPrimary = (appData && Array.isArray(appData.waypoints)) ? appData.waypoints : [];
      const waypointsFallback = Array.isArray(cachedDropdown.waypoints) ? cachedDropdown.waypoints : [];
      const waypoints = waypointsPrimary.length ? waypointsPrimary : waypointsFallback;
      const wp = waypoints.find(function(w) {
        return String((w && (w.wp_id || w.WP_ID || w.ident || w.ID || w.code || w.CODE || w.name || w.NAME)) || '').trim().toUpperCase() === key;
      });
      if (wp) {
        const latRaw = wp.lat != null ? wp.lat :
          (wp.latitude != null ? wp.latitude :
          (wp.LATITUDE != null ? wp.LATITUDE : wp.LAT));
        const lonRaw = wp.lon != null ? wp.lon :
          (wp.longitude != null ? wp.longitude :
          (wp.LONGITUDE != null ? wp.LONGITUDE : wp.LON));
        const lat = typeof _parseCoordinate === 'function' ? _parseCoordinate(latRaw) : parseFloat(latRaw);
        const lon = typeof _parseCoordinate === 'function' ? _parseCoordinate(lonRaw) : parseFloat(lonRaw);
        if (isFinite(lat) && isFinite(lon)) {
          return { ident: key, lat: lat, lon: lon };
        }
      }

      return null;
    }

    function _offfltBuildRouteWithEndpoints(waypointList, fromIcao, toIcao) {
      const from = _offfltNormIcao(fromIcao);
      const to = _offfltNormIcao(toIcao);
      const tokens = _offfltParseRouteTokens(waypointList);
      if (from && (!tokens.length || tokens[0] !== from)) tokens.unshift(from);
      if (to && (!tokens.length || tokens[tokens.length - 1] !== to)) tokens.push(to);
      return tokens.join(', ');
    }

    function _offfltRouteDistanceNm(routeText) {
      const ids = _offfltParseRouteTokens(routeText);
      if (ids.length < 2) return 0;
      let total = 0;
      for (let i = 0; i < ids.length - 1; i++) {
        const start = _offfltResolveLocation(ids[i]);
        const end = _offfltResolveLocation(ids[i + 1]);
        if (!start || !end) continue;
        total += _offfltHaversineNm(start.lat, start.lon, end.lat, end.lon);
      }
      return total;
    }

    function _offfltBuildRoutePlan(fromIcao, toIcao) {
      const from = _offfltNormIcao(fromIcao);
      const to = _offfltNormIcao(toIcao);
      const fromAirport = _offfltFindAirport(from);
      const toAirport = _offfltFindAirport(to);

      let routeText = [from, to].filter(Boolean).join(', ');
      let distance = _offfltDistanceNm(fromAirport, toAirport);

      const cachedDropdown = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA) || {};
      const routesPrimary = (appData && Array.isArray(appData.routes)) ? appData.routes : [];
      const routesFallback = Array.isArray(cachedDropdown.routes) ? cachedDropdown.routes : [];
      const routes = routesPrimary.length ? routesPrimary : routesFallback;
      const routeOrigin_ = function(r) {
        return _offfltNormIcao((r && (r.origin || r.ORIGIN || r.from || r.FROM)) || '');
      };
      const routeDestination_ = function(r) {
        return _offfltNormIcao((r && (r.destination || r.DESTINATION || r.to || r.TO)) || '');
      };
      const routeWaypointList_ = function(r) {
        return String((r && (r.waypoint_list || r.WAYPOINT_LIST || r.route || r.ROUTE || r.route_text || r.ROUTE_TEXT)) || '').trim();
      };
      const directRoute = routes.find(function(r) {
        return routeOrigin_(r) === from && routeDestination_(r) === to;
      });
      const reverseRoute = !directRoute ? routes.find(function(r) {
        return routeOrigin_(r) === to && routeDestination_(r) === from;
      }) : null;

      const directWaypoints = directRoute ? routeWaypointList_(directRoute) : '';
      const reverseWaypoints = reverseRoute ? routeWaypointList_(reverseRoute) : '';

      if (directWaypoints) {
        routeText = _offfltBuildRouteWithEndpoints(directWaypoints, from, to);
        distance = _offfltRouteDistanceNm(routeText) || distance;
      } else if (reverseWaypoints) {
        const reversed = _offfltParseRouteTokens(reverseWaypoints).reverse().join(', ');
        routeText = _offfltBuildRouteWithEndpoints(reversed, from, to);
        distance = _offfltRouteDistanceNm(routeText) || distance;
      }

      return {
        routeText: routeText,
        distanceNm: Math.max(0, distance)
      };
    }

    function _offfltHasRealEnvelope(acftReg) {
      try {
        const reg = String(acftReg || '').trim().toUpperCase();
        if (!reg) return false;

        const hasValidPoints = function(points) {
          const arr = Array.isArray(points) ? points : [];
          const valid = arr.filter(function(p) {
            const x = parseFloat(
              p && (p.CG_Arm_X ?? p.cgArmX ?? p.cgArm ?? p.arm ?? p.x)
            );
            const y = parseFloat(
              p && (p.Weight_Y ?? p.weightY ?? p.weight ?? p.y)
            );
            return !isNaN(x) && !isNaN(y);
          });
          return valid.length >= 3;
        };

        // Primary cache: per-aircraft envelope cache
        const key = 'mba_cache_envelope_' + reg;
        const raw = localStorage.getItem(key);
        if (raw) {
          const parsed = JSON.parse(raw);
          const payload = (parsed && parsed.value && typeof parsed.value === 'object') ? parsed.value : parsed;
          if (hasValidPoints(payload && payload.envelopeData)) return true;
        }

        // Fallback cache: any cached W&B payload for same aircraft
        const keys = Object.keys(localStorage || {});
        for (let i = 0; i < keys.length; i++) {
          const k = String(keys[i] || '');
          if (!k.startsWith('mba_cache_wb_')) continue;
          const itemRaw = localStorage.getItem(k);
          if (!itemRaw) continue;
          let parsed;
          try { parsed = JSON.parse(itemRaw); } catch (e) { continue; }
          const wbPayload = parsed && parsed.payload ? parsed.payload : (parsed && parsed.value && parsed.value.payload ? parsed.value.payload : null);
          if (!wbPayload) continue;
          const payloadReg = String((wbPayload.aircraft || wbPayload.acft || '')).trim().toUpperCase();
          if (payloadReg !== reg) continue;
          if (hasValidPoints(wbPayload.envelopeData)) return true;
        }

        return false;
      } catch (e) {
        return false;
      }
    }

    function _offfltUpdateAuthStatus(pilotName, destinationIcao) {
      const el = document.getElementById('offflt-auth-status');
      const authWrap = document.getElementById('offflt_ack_auth_wrap');
      const authAck = document.getElementById('offflt_ack_auth');
      if (!el) return { allowed: true, reason: 'no-ui' };

      const pilotKey = String(pilotName || '').trim();
      const dest = _offfltNormIcao(destinationIcao);
      if (!dest) {
        el.style.background = '#f7fbff';
        el.style.borderColor = '#cfe2f3';
        el.style.color = '#0b5394';
        el.textContent = 'Select destination, then confirm runway authorization.';
        if (authWrap) authWrap.style.display = 'block';
        if (authAck) authAck.checked = false;
        return { allowed: true, reason: 'incomplete' };
      }

      if (!pilotKey || pilotKey.toUpperCase() === 'PILOT TBD' || pilotKey.toUpperCase() === 'TBD') {
        el.style.background = '#fff8e1';
        el.style.borderColor = '#ffd54f';
        el.style.color = '#8d6e00';
        el.textContent = `PILOT TBD selected for ${dest}. Mission can be queued, but cannot be approved until a pilot is assigned.`;
        if (authWrap) authWrap.style.display = 'block';
        return { allowed: true, reason: 'pilot-tbd' };
      }

      el.style.background = '#f7fbff';
      el.style.borderColor = '#cfe2f3';
      el.style.color = '#0b5394';
      el.textContent = `Pilot acknowledgement required for ${pilotKey} to queue offline flight to ${dest}.`;
      if (authWrap) authWrap.style.display = 'block';
      return { allowed: true, reason: 'pilot-ack-required' };
    }

    function _offfltSetCalcValue(id, text, color) {
      const el = document.getElementById(id);
      if (!el) return;
      el.textContent = text;
      if (color) el.style.color = color;
    }

    function _offfltBuildComputeIssue(calc) {
      const details = calc && typeof calc === 'object' ? calc : {};
      if (!details.fromAirport && !details.toAirport) return 'Compute blocked: origin and destination airports were not resolved from startup data.';
      if (!details.fromAirport) return 'Compute blocked: origin airport was not resolved from startup data.';
      if (!details.toAirport) return 'Compute blocked: destination airport was not resolved from startup data.';
      if (!details.acftObj) return 'Compute blocked: aircraft profile was not resolved from startup data.';
      if (!(Number(details.distNm || 0) > 0)) return 'Compute blocked: route distance resolved to 0 nm.';
      if (!(Number(details.speed || 0) > 0)) return 'Compute blocked: aircraft cruise speed resolved to 0 kts.';
      return '';
    }

    function _offfltRecompute() {
      const from = _offfltNormIcao((document.getElementById('offflt_from') || {}).value);
      const to = _offfltNormIcao((document.getElementById('offflt_to') || {}).value);
      const acft = String((document.getElementById('offflt_acft') || {}).value || '').trim();
      const pilot = String((document.getElementById('offflt_pilot') || {}).value || '').trim();

      const fromAirport = _offfltFindAirport(from);
      const toAirport = _offfltFindAirport(to);
      const acftObj = _offfltFindAircraft(acft);

      const routePlan = _offfltBuildRoutePlan(from, to);
      const distNm = Number(routePlan.distanceNm || 0);
      const speed = acftObj ? (parseFloat(acftObj.speed) || 0) : 0;
      const burn = acftObj ? (parseFloat(acftObj.burn) || 0) : 0;
      const suggestedTime = (distNm > 0 && speed > 0) ? (distNm / speed) : 0;
      const useTime = suggestedTime;
      const taxiTakeoffFuel = useTime > 0 ? 5 : 0;
      const estFuel = (useTime > 0 && burn > 0) ? ((useTime * burn) + taxiTakeoffFuel) : 0;
      const reserveFuel = burn > 0 ? burn : 0;
      const requiredFuel = estFuel + reserveFuel;
      const estGroundTime = 0.5;
      const estDuty = useTime > 0 ? (1.0 + useTime + estGroundTime + 0.75) : 0;

      const resolveAirportMtowLimit = function(airport, aircraft) {
        const map = (airport && airport.mtowByModel && typeof airport.mtowByModel === 'object') ? airport.mtowByModel : {};
        const keys = Object.keys(map);
        if (!keys.length) return 9999;
        const acftRef = String(
          (aircraft && (aircraft.typeForPerformance || aircraft.aircraftType || aircraft.reg)) || ''
        ).toUpperCase().replace(/[^A-Z0-9]+/g, '_');
        if (acftRef) {
          const sorted = keys.slice().sort(function(a, b) { return String(b || '').length - String(a || '').length; });
          const matched = sorted.find(function(k) { return acftRef.indexOf(String(k || '').toUpperCase()) >= 0; });
          if (matched) return Number(map[matched] || 0) || 9999;
        }
        if (map.GENERIC) return Number(map.GENERIC || 0) || 9999;
        const vals = keys.map(function(k) { return Number(map[k] || 0); }).filter(function(v) { return isFinite(v) && v > 0; });
        return vals.length ? Math.max.apply(null, vals) : 9999;
      };
      const aptLimit = resolveAirportMtowLimit(fromAirport, acftObj);
      const acftMtow = acftObj ? (Number(acftObj.mtow || acftObj.MTOW || 9999) || 9999) : 9999;
      const takeoffLimitKg = Math.min(acftMtow, aptLimit);
      const limitType = aptLimit < acftMtow ? 'Airport Limit' : 'Aircraft MTOW';

      _offfltSetCalcValue('offflt_calc_dist', distNm > 0 ? `${distNm.toFixed(1)} nm` : '— nm', '#0b5394');
      _offfltSetCalcValue('offflt_calc_time', suggestedTime > 0 ? `${suggestedTime.toFixed(1)} hr` : '— hr', '#0b5394');
      _offfltSetCalcValue('offflt_calc_fuel', requiredFuel > 0 ? `${Math.round(requiredFuel)} L` : '— L', '#0b5394');
      _offfltSetCalcValue('offflt_calc_duty', estDuty > 0 ? `${estDuty.toFixed(1)} hr` : '— hr', estDuty > 14 ? '#c62828' : '#0b5394');
      _offfltSetCalcValue('offflt_calc_route', routePlan.routeText || '—', '#455a64');
      _offfltSetCalcValue('offflt_calc_diag', '', '#8d6e63');

      const auth = _offfltUpdateAuthStatus(pilot, to);
      const authEl = document.getElementById('offflt-auth-status');
      const hasEnvelope = _offfltHasRealEnvelope(acft);
      if (authEl && acft && !hasEnvelope) {
        authEl.style.background = '#ffebee';
        authEl.style.borderColor = '#ef9a9a';
        authEl.style.color = '#b71c1c';
        authEl.textContent = `No real CG envelope is cached for ${acft}. Connect online and open W&B for this aircraft before creating offline flight.`;
      }

      const computeIssue = _offfltBuildComputeIssue({
        fromAirport: fromAirport,
        toAirport: toAirport,
        acftObj: acftObj,
        distNm: distNm,
        speed: speed
      });
      if (computeIssue) {
        _offfltSetCalcValue('offflt_calc_diag', computeIssue, '#b26a00');
      }

      return {
        fromAirport: fromAirport,
        toAirport: toAirport,
        distNm: distNm,
        speed: speed,
        suggestedTime: suggestedTime,
        usedTime: useTime,
        estFuel: estFuel,
        reserveFuel: reserveFuel,
        requiredFuel: requiredFuel,
        estDuty: estDuty,
        routeText: routePlan.routeText,
        takeoffLimitKg: takeoffLimitKg,
        limitType: limitType,
        estGroundTime: estGroundTime,
        auth: auth,
        hasEnvelope: hasEnvelope,
        acftObj: acftObj
      };
    }

    function _offfltBindEventsOnce() {
      const elem = document.getElementById('modalOfflineFlight');
      if (!elem || elem.dataset.boundOfflineFlight === '1') return;
      elem.dataset.boundOfflineFlight = '1';
      ['offflt_from','offflt_to','offflt_acft','offflt_pilot'].forEach(function(id) {
        const el = document.getElementById(id);
        if (!el) return;
        el.addEventListener('input', _offfltRecompute);
        el.addEventListener('change', _offfltRecompute);
      });
    }

    function openOfflineFlightModal() {
      const dateEl = document.getElementById('offflt_date');
      const timeEl = document.getElementById('offflt_time');
      const acftEl = document.getElementById('offflt_acft');
      const pilotEl = document.getElementById('offflt_pilot');
      const fromEl = document.getElementById('offflt_from');
      const toEl = document.getElementById('offflt_to');
      const ackLic = document.getElementById('offflt_ack_license');
      const ackDuty = document.getElementById('offflt_ack_duty');
      const ackAuth = document.getElementById('offflt_ack_auth');
      const airportsDl = document.getElementById('offflt-airports-list');

      if (dateEl && !dateEl.value) {
        const now = new Date();
        dateEl.value = now.toISOString().slice(0, 10);
      }
      if (timeEl && !timeEl.value) timeEl.value = '08:00';

      if (acftEl) {
        const cachedDropdown = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA) || {};
        const acftPrimary = (appData && Array.isArray(appData.aircraft)) ? appData.aircraft : [];
        const acftFallback = Array.isArray(cachedDropdown.aircraft) ? cachedDropdown.aircraft : [];
        const acftList = acftPrimary.length ? acftPrimary : acftFallback;
        acftEl.innerHTML = '<option value="" disabled selected>Select Aircraft</option>' +
          acftList.map(a => `<option value="${String(a.reg || '')}">${String(a.reg || '')}</option>`).join('');
      }

      if (pilotEl) {
        const cachedDropdown = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA) || {};
        const pilotPrimary = (appData && Array.isArray(appData.pilots)) ? appData.pilots : [];
        const pilotFallback = Array.isArray(cachedDropdown.pilots) ? cachedDropdown.pilots : [];
        const pilotList = pilotPrimary.length ? pilotPrimary : pilotFallback;
        pilotEl.innerHTML = '<option value="" disabled selected>Select Pilot</option>' +
          '<option value="PILOT TBD">PILOT TBD (Pending Only)</option>' +
          pilotList.map(p => `<option value="${String(p.name || '')}">${String(p.name || '')}</option>`).join('');
      }

      if (airportsDl) {
        const cachedDropdown = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA) || {};
        const airportsPrimary = (appData && Array.isArray(appData.airports)) ? appData.airports : [];
        const airportsFallback = Array.isArray(cachedDropdown.airports) ? cachedDropdown.airports : [];
        const airports = airportsPrimary.length ? airportsPrimary : airportsFallback;
        const seen = {};
        airportsDl.innerHTML = airports.map(a => {
          const icao = _offfltNormIcao(a.icao);
          if (!icao || seen[icao]) return '';
          seen[icao] = true;
          const name = String(a.nome || '').trim();
          return `<option value="${icao}">${icao}${name ? ' — ' + name : ''}</option>`;
        }).filter(Boolean).join('');
      }

      if (fromEl) fromEl.value = _offfltNormIcao(fromEl.value);
      if (toEl) toEl.value = _offfltNormIcao(toEl.value);
      if (ackLic) ackLic.checked = false;
      if (ackDuty) ackDuty.checked = false;
      if (ackAuth) ackAuth.checked = false;

      _offfltBindEventsOnce();
      _offfltRecompute();

      const elem = document.getElementById('modalOfflineFlight');
      const modal = M.Modal.getInstance(elem) || M.Modal.init(elem);
      modal.open();
    }

    function buildOfflineFlightIds(dateStr) {
      const compact = String(dateStr || '').replace(/-/g, '').slice(2);
      const stamp = String(Date.now()).slice(-5);
      const missionId = `OFL${compact}-${stamp}`;
      const flightLegId = `${missionId}-01`;
      return { missionId, flightLegId };
    }

    function addDraftToLocalMissionCaches(payload, missionId) {
      const list = cacheGet(OFFLINE_CACHE_KEYS.SCHEDULED_MISSIONS) || [];
      const firstLeg = payload && payload.legs && payload.legs[0] ? payload.legs[0] : {};
      const newItem = {
        id: missionId,
        date: payload.date,
        acft: payload.acft,
        pilot: payload.pilot,
        from: firstLeg.from || '',
        to: firstLeg.to || '',
        status: 'DRAFT_OFFLINE'
      };

      const withoutSame = list.filter(m => String(m.id) !== String(missionId));
      withoutSame.unshift(newItem);
      cacheSet(OFFLINE_CACHE_KEYS.SCHEDULED_MISSIONS, withoutSame);

      const missionDetails = {
        id: missionId,
        date: payload.date,
        acft: payload.acft,
        pilot: payload.pilot,
        meta: {
          acft: payload.acft,
          pilot: payload.pilot,
          copilot: payload.copilot || '',
          notes: payload.notes || ''
        },
        legs: payload.legs.map(l => ({
          waypoints: (function() {
            if (Array.isArray(l.waypoints) && l.waypoints.length) {
              return l.waypoints.map(function(wp) { return String(wp || '').trim().toUpperCase(); }).filter(Boolean);
            }
            const raw = String(l.route || `${l.from || ''}-${l.to || ''}` || '').trim().toUpperCase();
            return raw
              .replace(/[→>]/g, ',')
              .split(/[\n\r,;\/|]+/)
              .map(function(part) { return String(part || '').trim().toUpperCase(); })
              .filter(Boolean);
          })(),
          flightLegId: l.flightLegId,
          from: l.from,
          to: l.to,
          route: String(l.route || `${l.from}-${l.to}`),
          time: Number(l.time || 0),
          groundTime: Number(l.groundTime || 0.5),
          dist: Number(l.dist || 0),
          distance: Number(l.dist || 0),
          fuel: Number(l.fuel || 0),
          takeoffFuel: Number(l.takeoffFuel || 0),
          landingFuel: Number(l.landingFuel || 0),
          limit: Number(l.limit || 0),
          limitType: String(l.limitType || ''),
          pax: Array.isArray(l.pax) ? l.pax : []
        }))
      };
      cacheSet(missionCacheKey(missionId), missionDetails);
    }

    function submitOfflineFlightDraft() {
      const date = String(document.getElementById('offflt_date')?.value || '').trim();
      const missionTime = String(document.getElementById('offflt_time')?.value || '08:00').trim() || '08:00';
      const acft = String(document.getElementById('offflt_acft')?.value || '').trim();
      const pilot = String(document.getElementById('offflt_pilot')?.value || '').trim() || 'PILOT TBD';
      const from = String(document.getElementById('offflt_from')?.value || '').trim().toUpperCase();
      const to = String(document.getElementById('offflt_to')?.value || '').trim().toUpperCase();
      const ackLicense = !!(document.getElementById('offflt_ack_license') || {}).checked;
      const ackDuty = !!(document.getElementById('offflt_ack_duty') || {}).checked;
      const ackAuth = !!(document.getElementById('offflt_ack_auth') || {}).checked;

      const calc = _offfltRecompute();
      const legTime = (calc.usedTime > 0 ? calc.usedTime : 0);

      if (!date || !acft || !from || !to) {
        M.toast({ html: 'Fill Date, Aircraft, From, To', classes: 'orange' });
        return;
      }

      if (!(legTime > 0)) {
        M.toast({ html: _offfltBuildComputeIssue(calc) || 'Could not compute flight time from route + aircraft speed', classes: 'orange' });
        return;
      }

      if (!calc.fromAirport || !calc.toAirport) {
        M.toast({ html: 'Use valid ICAO airports from cached DB airports', classes: 'orange' });
        return;
      }

      if (!calc.hasEnvelope) {
        M.toast({ html: 'No real CG envelope cached for this aircraft. Go online and open W&B first.', classes: 'red' });
        return;
      }

      if (!ackAuth) {
        M.toast({ html: 'Acknowledge runway authorization before continuing', classes: 'orange' });
        return;
      }

      if (!ackLicense || !ackDuty) {
        M.toast({ html: 'Flight App asks you to verify license + duty acknowledgements', classes: 'orange' });
        return;
      }

      const acftObj = calc.acftObj || _offfltFindAircraft(acft);
      const burn = acftObj ? (parseFloat(acftObj.burn) || 0) : 0;
      const tripFuel = burn > 0 ? (legTime * burn) : Number(calc.estFuel || 0);
      const takeoffExtraFuel = legTime > 0 ? 5 : 0;
      const adjustedTripFuel = tripFuel + takeoffExtraFuel;
      const reserveFuel = burn > 0 ? burn : Number(calc.reserveFuel || 0);
      const estFuel = Math.round(adjustedTripFuel + reserveFuel);
      const estGroundTime = Number(calc.estGroundTime || 0.5);
      const estDuty = 1.0 + legTime + estGroundTime + 0.75;
      const limitKg = Number(calc.takeoffLimitKg || 0);
      const limitType = String(calc.limitType || '').trim();

      let licenseWarning = '';
      const pilotObj = _offfltFindPilot(pilot);
      if (pilotObj && acftObj) {
        const required = String(acftObj.License_Required || 'MNTE').toUpperCase();
        const expiryStr = required.indexOf('MNAF') >= 0 ? pilotObj.MNAF_Validity : pilotObj.MNTE_Validity;
        if (expiryStr) {
          const expiry = new Date(expiryStr + 'T00:00:00');
          const missionDate = new Date(date + 'T00:00:00');
          if (!isNaN(expiry.getTime()) && !isNaN(missionDate.getTime()) && expiry < missionDate) {
            licenseWarning = ` [WARN LICENSE ${required} EXPIRED ${expiryStr}]`;
          }
        }
      }

      const dutyWarning = estDuty > 14 ? ` [WARN DUTY EST ${estDuty.toFixed(1)}h > 14h]` : '';

      let routeTokens = _offfltParseRouteTokens(calc.routeText || `${from}, ${to}`);
      if (from && (!routeTokens.length || routeTokens[0] !== from)) routeTokens.unshift(from);
      if (to && (!routeTokens.length || routeTokens[routeTokens.length - 1] !== to)) routeTokens.push(to);
      routeTokens = routeTokens.filter(function(token, idx, arr) { return idx === 0 || token !== arr[idx - 1]; });
      const routeHyphen = routeTokens.join('-') || `${from}-${to}`;

      const ids = buildOfflineFlightIds(date);
      const payload = {
        date: date,
        time: missionTime,
        acft: acft,
        pilot: pilot,
        copilot: '',
        type: 'Offline Flight',
        notes: (`[OFFLINE FLIGHT][ROUTE ${String(routeHyphen || (from + '-' + to)).replace(/\s+/g, ' ').trim()}][RESERVE 1.0H ${Math.round(reserveFuel)}L][TKOF EXTRA 5L]${licenseWarning}${dutyWarning}${ackAuth ? ' [ACK RUNWAY AUTHORIZATION]' : ''}`).trim(),
        legs: [{
          flightLegId: ids.flightLegId,
          from: from,
          to: to,
          route: routeHyphen,
          waypoints: routeTokens,
          time: legTime,
          groundTime: estGroundTime,
          dist: calc.distNm > 0 ? Number(calc.distNm.toFixed(1)) : 0,
          fuel: Math.round(adjustedTripFuel) > 0 ? Math.round(adjustedTripFuel) : 0,
          takeoffFuel: estFuel > 0 ? estFuel : 0,
          landingFuel: Math.round(reserveFuel),
          limit: limitKg > 0 ? Math.round(limitKg) : undefined,
          limitType: limitType || undefined,
          plannedCacheDraw: 0,
          pax: []
        }]
      };

      const btn = document.getElementById('btn-offflt-save');
      if (btn) {
        btn.disabled = true;
        btn.textContent = 'QUEUEING...';
      }

      window.runOrQueueServerAction({
        method: 'saveMission',
        args: [payload],
        label: 'Offline flight'
      }, {
        onSuccess: function() {
          addDraftToLocalMissionCaches(payload, ids.missionId);
          renderMissionList();
          M.toast({ html: 'Offline flight created', classes: 'green' });
          const modal = M.Modal.getInstance(document.getElementById('modalOfflineFlight'));
          if (modal) modal.close();
          if (btn) {
            btn.disabled = false;
            btn.textContent = 'QUEUE OFFLINE FLIGHT';
          }
        },
        onQueued: function() {
          addDraftToLocalMissionCaches(payload, ids.missionId);
          renderMissionList();
          M.toast({ html: 'Offline flight queued for sync', classes: 'orange' });
          const modal = M.Modal.getInstance(document.getElementById('modalOfflineFlight'));
          if (modal) modal.close();
          if (btn) {
            btn.disabled = false;
            btn.textContent = 'QUEUE OFFLINE FLIGHT';
          }
        },
        onFailure: function(err) {
          M.toast({ html: 'Create failed: ' + (err && err.message ? err.message : String(err || 'unknown')), classes: 'red' });
          if (btn) {
            btn.disabled = false;
            btn.textContent = 'QUEUE OFFLINE FLIGHT';
          }
        }
      });
    }

    // Initial Load
    document.addEventListener('DOMContentLoaded', function() {
      pruneMissionCache(OFFLINE_CACHE_MAX_MISSIONS);
      renderPilotTabVersions_();
      renderLessonLinkBar_();
      initPilotDispatchClocks_();
      updateConnectivityBanner();
      updateOfflineCacheStatus();
      updateOutboxButtonState();
      updateNewFlightButtonState();
      updateOfflineFlightButtonState();
      ensureEnvelopeCacheOnStartup();
      initDutyPromptScheduler_();
      window.addEventListener('online', updateConnectivityBanner);
      window.addEventListener('offline', updateConnectivityBanner);
      window.addEventListener('online', function() {
        updateOfflineCacheStatus();
        updateOutboxButtonState();
        updateNewFlightButtonState();
        updateOfflineFlightButtonState();
        ensureEnvelopeCacheOnStartup();
        processOutboxQueue();
        // Auto-refresh mission list so status changes (e.g. PENDING->APPROVED) are visible immediately
        if (typeof renderMissionList === 'function') renderMissionList({ forceServerList: true, skipPrefetch: true, silent: true });
      });
      window.addEventListener('offline', function() {
        updateOfflineCacheStatus();
        updateOutboxButtonState();
        updateNewFlightButtonState();
        updateOfflineFlightButtonState();
      });
      setTimeout(loadInitialData, 100);

      var lpModal = document.getElementById('lesson-plan-modal');
      if (lpModal) {
        lpModal.addEventListener('click', function(ev) {
          if (ev.target === lpModal) closeLessonPlanModal_();
        });
      }
    });

    window.addEventListener('beforeunload', function() {
      if (_pilotClockTimer) clearInterval(_pilotClockTimer);
      _pilotClockTimer = null;
      if (_dutyPromptTimer) clearInterval(_dutyPromptTimer);
      _dutyPromptTimer = null;
    });

    function loadInitialData() {
      if (typeof M === 'undefined') {
        setTimeout(loadInitialData, 500);
        return;
      }

      const consumeCamTopLaunchRequest_ = function() {
        try {
          const url = new URL(window.location.href);
          const camTop = String(url.searchParams.get('camtop') || '');
          if (camTop !== '1') return false;
          url.searchParams.delete('camtop');
          url.searchParams.delete('t');
          history.replaceState(null, '', url.toString());
          return true;
        } catch (e) {
          return (window.location.search || '').indexOf('camtop=1') >= 0;
        }
      };

      const applyDataAndStart = function(data, fromCache, options) {
        const incoming = (data && typeof data === 'object') ? data : {};
        const opts = (options && typeof options === 'object') ? options : {};
        if (opts.merge && appData && typeof appData === 'object') {
          appData = Object.assign({}, appData, incoming);
        } else {
          appData = incoming;
        }
        window.appData = appData;
        M.AutoInit();
        renderMissionList();
        updateOfflineCacheStatus();
        updateOutboxButtonState();
        updateNewFlightButtonState();
        updateOfflineFlightButtonState();
        if (!fromCache) {
          syncEnvelopeCacheForAircraftRegs_(_knownAircraftRegsForEnvelopeSync_(), { forceRefresh: false });
        }
        if (fromCache && !opts.silentToast) {
          M.toast({ html: 'Offline startup: loaded cached data', classes: 'orange' });
        }

        if (consumeCamTopLaunchRequest_()) {
          setTimeout(function() {
            try { switchTab(1); } catch (e) {}
            try {
              openTab1Inclinometer();
            } catch (e2) {
              if (window.M) M.toast({ html: 'Falha ao abrir inclinômetro automaticamente.', classes: 'orange' });
            }
          }, 450);
        }
      };

      const isTransientStartupError_ = function(err) {
        const msg = err && err.message ? String(err.message) : String(err || '');
        return /HTTP 0|NetworkError|Connection failure|Failed to fetch|ScriptError/i.test(msg);
      };

      const startupDiag = Array.isArray(window.__startupDiag) ? window.__startupDiag : [];
      window.__startupDiag = startupDiag;
      const pushStartupDiag_ = function(stage, details) {
        try {
          startupDiag.push({
            ts: new Date().toISOString(),
            stage: String(stage || ''),
            online: (typeof navigator !== 'undefined' && typeof navigator.onLine === 'boolean') ? navigator.onLine : null,
            details: details || {}
          });
          if (startupDiag.length > 40) startupDiag.splice(0, startupDiag.length - 40);
        } catch (e) {}
      };
      const escHtml_ = function(v) {
        return String(v == null ? '' : v)
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;')
          .replace(/'/g, '&#39;');
      };

      const renderStartupState_ = function(kind, msg) {
        const missionList = document.getElementById('mission-list-container');
        if (!missionList) return;
        if (kind === 'retrying') {
          missionList.innerHTML = "<div class='card-panel orange-text' style='margin:10px;'>Connection unstable during startup. Retrying automatically... <button type='button' id='startup-retry-now' class='btn-flat' style='margin-left:8px; color:#1565c0; font-weight:700;'>Retry now</button></div>";
          const btn = document.getElementById('startup-retry-now');
          if (btn) {
            btn.onclick = function() {
              try { if (window.__startupRetryTimer) clearTimeout(window.__startupRetryTimer); } catch (e) {}
              window.__startupRetryTimer = null;
              requestStartupData_(0);
            };
          }
          return;
        }
        missionList.innerHTML = `<div class='card-panel red-text' style='margin:10px;'>System startup failed: ${escHtml_(msg || 'Unknown startup error')} <button type='button' id='startup-retry-now' class='btn-flat' style='margin-left:8px; color:#1565c0; font-weight:700;'>Retry now</button></div>`;
        const btn = document.getElementById('startup-retry-now');
        if (btn) btn.onclick = function() { requestStartupData_(0); };
      };

      const requestStartupData_ = function(attempt) {
        try {
          if (window.__startupRetryTimer) {
            clearTimeout(window.__startupRetryTimer);
            window.__startupRetryTimer = null;
          }
        } catch (e) {}

        pushStartupDiag_('request', {
          attempt: attempt,
          href: (window && window.location && window.location.href) ? String(window.location.href).slice(0, 200) : ''
        });

        google.script.run.withSuccessHandler(data => {
          window.__startupRetryBackoffMs = 2500;
          if (data && data.error) {
            pushStartupDiag_('server-error-payload', { attempt: attempt, message: String(data.error || '') });
            renderStartupState_('failed', String(data.error || 'Startup payload error'));
            return;
          } else {
            pushStartupDiag_('success', { attempt: attempt });
          }
          const cached = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA) || {};
          const merged = Object.assign({}, cached, data || {});
          cacheSet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA, merged);
          applyDataAndStart(merged, false);
          const loadAirportChunks_ = function(offset, collected) {
            const startOffset = Number(offset || 0);
            const acc = Array.isArray(collected) ? collected : [];
            google.script.run.withSuccessHandler(function(airportChunk) {
              if (airportChunk && airportChunk.error) {
                pushStartupDiag_('airport-error-payload', { attempt: attempt, message: String(airportChunk.error || '') });
                return;
              }
              const part = Array.isArray(airportChunk && airportChunk.airports) ? airportChunk.airports : [];
              const mergedPart = acc.concat(part);
              const done = !!(airportChunk && airportChunk.done);
              const nextOffset = Number(airportChunk && airportChunk.nextOffset || 0);

              if (!done) {
                pushStartupDiag_('airport-chunk', {
                  attempt: attempt,
                  received: mergedPart.length,
                  nextOffset: nextOffset,
                  total: Number(airportChunk && airportChunk.total || 0)
                });
                loadAirportChunks_(nextOffset, mergedPart);
                return;
              }

              const airportData = { airports: mergedPart };
              const cachedAfterCore = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA) || {};
              const mergedAirports = Object.assign({}, cachedAfterCore, airportData);
              cacheSet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA, mergedAirports);
              pushStartupDiag_('airport-success', {
                attempt: attempt,
                airports: mergedPart.length
              });
              applyDataAndStart(airportData, false, { merge: true, silentToast: true });
            }).withFailureHandler(function(err) {
              const msgText = err && err.message ? String(err.message) : String(err || '');
              pushStartupDiag_('airport-failure', {
                attempt: attempt,
                transient: isTransientStartupError_(err),
                message: msgText.slice(0, 500)
              });
            }).getPilotAirportDataChunk(startOffset, 1200);
          };

          loadAirportChunks_(0, []);
        }).withFailureHandler(err => {
          const msgText = err && err.message ? String(err.message) : String(err || '');
          pushStartupDiag_('failure', {
            attempt: attempt,
            transient: isTransientStartupError_(err),
            message: msgText.slice(0, 500)
          });
          const cached = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA);
          if (cached) {
            window.__startupRetryBackoffMs = 2500;
            pushStartupDiag_('fallback-cache', { attempt: attempt });
            applyDataAndStart(cached, true);
            return;
          }

          if (isTransientStartupError_(err) && attempt < 4) {
            const retryDelayMs = attempt === 0 ? 700 : (attempt === 1 ? 1200 : 1800);
            if (attempt === 0) M.toast({ html: 'Startup network hiccup. Retrying…', classes: 'orange', displayLength: 1600 });
            setTimeout(function() {
              requestStartupData_(attempt + 1);
            }, retryDelayMs);
            return;
          }

          if (isTransientStartupError_(err)) {
            renderStartupState_('retrying');
            const nextDelay = Math.max(2500, Math.min(Number(window.__startupRetryBackoffMs || 2500), 12000));
            window.__startupRetryBackoffMs = Math.min(nextDelay + 1500, 12000);
            window.__startupRetryTimer = setTimeout(function() {
              requestStartupData_(0);
            }, nextDelay);
            return;
          }

          const msg = err && err.message ? err.message : String(err || 'Unknown startup error');
          renderStartupState_('failed', msg);
          M.toast({ html: `Startup failed: ${msg}`, classes: 'red', displayLength: 6000 });
        }).getPilotStartupData();
      };

      const cachedStartup = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA);
      const freshCache = isOfflineCacheFresh_();

      if (cachedStartup && freshCache) {
        applyDataAndStart(cachedStartup, true, { silentToast: true });
        return;
      }

      if (isServerAvailable()) {
        requestStartupData_(0);
        return;
      }

      if (cachedStartup) {
        applyDataAndStart(cachedStartup, true);
        return;
      }

      const cachedMissionsFallback = cacheGet(OFFLINE_CACHE_KEYS.SCHEDULED_MISSIONS) || [];
      const cachedMissionDetailsCount = getCachedMissionCount();
      if (cachedMissionsFallback.length || cachedMissionDetailsCount > 0) {
        applyDataAndStart({ missions: cachedMissionsFallback }, true, { silentToast: true });
        M.toast({ html: 'Offline startup: using cached missions only', classes: 'orange', displayLength: 3500 });
        return;
      }

      const missionList = document.getElementById('mission-list-container');
      if (missionList) {
        missionList.innerHTML = "<div class='card-panel orange-text' style='margin:10px;'>Offline and no cached startup data available yet. Connect once to seed cache.</div>";
      }
    }

    // Tab Logic
    const TAB_FIELD_STATE_PREFIX = 'mba_tab_field_state_v1';

    function _tabFieldStateKey_(tabNum) {
      const missionKey = String(activeMission || '').trim();
      if (!missionKey) return '';
      return `${TAB_FIELD_STATE_PREFIX}:${missionKey}:tab${Number(tabNum || 0)}`;
    }

    function _tabCollectFieldState_(tabNum) {
      const tabEl = document.getElementById('tab' + tabNum);
      if (!tabEl) return {};
      const fields = tabEl.querySelectorAll('input[id], select[id], textarea[id]');
      const state = {};
      fields.forEach(function(el) {
        if (!el || !el.id) return;
        const tag = String(el.tagName || '').toLowerCase();
        const type = String(el.type || '').toLowerCase();
        if (el.disabled || el.readOnly) return;
        if (tag === 'input' && ['hidden', 'password', 'file', 'button', 'submit', 'reset'].includes(type)) return;
        if (type === 'checkbox' || type === 'radio') {
          state[el.id] = { t: 'checked', v: !!el.checked };
        } else {
          state[el.id] = { t: 'value', v: String(el.value == null ? '' : el.value) };
        }
      });
      return state;
    }

    function _tabCollectExtraState_(tabNum) {
      const n = Number(tabNum || 0);

      if (n === 2) {
        const tanks = {};
        document.querySelectorAll('#tab2 .brief-tank-input').forEach(function(input) {
          const key = String(input && input.dataset && input.dataset.tankKey || '').trim().toUpperCase();
          if (!key) return;
          tanks[key] = String(input.value == null ? '' : input.value);
        });

        return {
          tanks: tanks,
          oil: String(document.getElementById('brief_oil')?.value || ''),
          startupTank: String(document.getElementById('brief_startup_tank')?.value || ''),
          startTach: String(document.getElementById('brief_startTach')?.value || ''),
          volts: String(document.getElementById('brief_volts')?.value || '')
        };
      }

      if (n === 4) {
        return {
          windDirection: String(document.getElementById('perf-wind-direction')?.value || ''),
          windSpeed: String(document.getElementById('perf-wind-speed')?.value || ''),
          windConfirmed: !!(window.performanceTab4 && window.performanceTab4.windConfirmed)
        };
      }

      if (n === 5) {
        return {
          selected: (window.releaseTab5State && window.releaseTab5State.selected)
            ? Object.assign({}, window.releaseTab5State.selected)
            : {}
        };
      }

      if (n === 7) {
        return {
          qnh: String(document.getElementById('arr7-qnh')?.value || ''),
          temp: String(document.getElementById('arr7-temp')?.value || ''),
          windDir: String(document.getElementById('arr7-wind-dir-input')?.value || ''),
          windSpd: String(document.getElementById('arr7-wind-spd-input')?.value || ''),
          windConfirmed: !!(window.arrival7 && window.arrival7.windConfirmed)
        };
      }

      return null;
    }

    function _tabRestoreExtraState_(tabNum, extra) {
      if (!extra || typeof extra !== 'object') return;
      const n = Number(tabNum || 0);

      if (n === 2) {
        const tanks = (extra.tanks && typeof extra.tanks === 'object') ? extra.tanks : {};
        Object.keys(tanks).forEach(function(key) {
          const normalized = String(key || '').trim().toUpperCase();
          if (!normalized) return;
          const input = document.querySelector('#tab2 .brief-tank-input[data-tank-key="' + normalized + '"]');
          if (input) {
            input.value = String(tanks[key] == null ? '' : tanks[key]);
          }
          const btn = document.querySelector('#tab2 .brief-tank-btn[data-tank-key="' + normalized + '"]');
          if (btn) {
            btn.textContent = String(tanks[key] == null ? '' : tanks[key]);
          }
        });

        const oilInput = document.getElementById('brief_oil');
        if (oilInput) oilInput.value = String(extra.oil == null ? '' : extra.oil);
        document.querySelectorAll('#tab2 .oil-choice-row .oil-choice').forEach(function(btn) {
          const v = String(btn && btn.dataset && btn.dataset.oil || '');
          btn.classList.toggle('active', v === String(extra.oil == null ? '' : extra.oil));
        });

        const startup = document.getElementById('brief_startup_tank');
        if (startup) startup.value = String(extra.startupTank == null ? '' : extra.startupTank);

        const tach = document.getElementById('brief_startTach');
        if (tach) tach.value = String(extra.startTach == null ? '' : extra.startTach);

        const volts = document.getElementById('brief_volts');
        if (volts) volts.value = String(extra.volts == null ? '' : extra.volts);

        if (typeof window.briefRefreshStartupTankUi_ === 'function') {
          window.briefRefreshStartupTankUi_();
        }
        if (typeof window.calculateBriefFuelTally === 'function') {
          window.calculateBriefFuelTally();
        }
        return;
      }

      if (n === 4) {
        const dir = document.getElementById('perf-wind-direction');
        const spd = document.getElementById('perf-wind-speed');
        if (dir) dir.value = String(extra.windDirection == null ? '' : extra.windDirection);
        if (spd) spd.value = String(extra.windSpeed == null ? '' : extra.windSpeed);

        if (!window.performanceTab4) window.performanceTab4 = {};
        window.performanceTab4.windConfirmed = !!extra.windConfirmed;

        if (typeof window._perfUpdateWindBtn_ === 'function') {
          window._perfUpdateWindBtn_();
        }
        if (typeof window.schedulePerformanceRecalc === 'function') {
          window.schedulePerformanceRecalc();
        }
        return;
      }

      if (n === 5) {
        const selected = (extra.selected && typeof extra.selected === 'object') ? extra.selected : null;
        if (!selected || typeof window.selectRiskLevel !== 'function') return;

        ['Pilot', 'Aircraft', 'Runway', 'Weather', 'Mission'].forEach(function(category) {
          const level = Number(selected[category] || 0);
          if (level >= 1 && level <= 3) {
            window.selectRiskLevel(category, level);
          }
        });
        return;
      }

      if (n === 7) {
        const extra7 = extra;
        const qnh = document.getElementById('arr7-qnh');
        const temp = document.getElementById('arr7-temp');
        const windDir = document.getElementById('arr7-wind-dir-input');
        const windSpd = document.getElementById('arr7-wind-spd-input');

        if (qnh) qnh.value = String(extra7.qnh == null ? '' : extra7.qnh);
        if (temp) temp.value = String(extra7.temp == null ? '' : extra7.temp);
        if (windDir) windDir.value = String(extra7.windDir == null ? '' : extra7.windDir);
        if (windSpd) windSpd.value = String(extra7.windSpd == null ? '' : extra7.windSpd);

        if (window.arrival7) window.arrival7.windConfirmed = !!extra7.windConfirmed;

        if (typeof window._arr7CalcIfReadyGlobal_ === 'function') {
          window._arr7CalcIfReadyGlobal_();
        }
      }
    }

    function _tabPersistFieldState_(tabNum) {
      try {
        const key = _tabFieldStateKey_(tabNum);
        if (!key) return;
        const state = _tabCollectFieldState_(tabNum);
        const extra = _tabCollectExtraState_(tabNum);
        localStorage.setItem(key, JSON.stringify({ savedAt: new Date().toISOString(), state: state, extra: extra }));
      } catch (e) {}
    }

    function _tabRestoreFieldState_(tabNum) {
      try {
        const key = _tabFieldStateKey_(tabNum);
        if (!key) return;
        const raw = localStorage.getItem(key);
        if (!raw) return;
        const parsed = JSON.parse(raw);
        const state = (parsed && parsed.state && typeof parsed.state === 'object') ? parsed.state : {};
        const extra = (parsed && parsed.extra && typeof parsed.extra === 'object') ? parsed.extra : null;
        Object.keys(state).forEach(function(id) {
          const el = document.getElementById(id);
          const entry = state[id];
          if (!el || !entry) return;
          if (entry.t === 'checked') {
            el.checked = !!entry.v;
          } else {
            el.value = String(entry.v == null ? '' : entry.v);
          }
          el.dispatchEvent(new Event('input', { bubbles: true }));
          el.dispatchEvent(new Event('change', { bubbles: true }));
        });
        _tabRestoreExtraState_(tabNum, extra);
      } catch (e) {}
    }

    if (!window.__tabFieldStateHooksBound) {
      const persistNow = function() {
        if (!currentTab || !activeMission) return;
        _tabPersistFieldState_(currentTab);
      };
      document.addEventListener('change', persistNow);
      document.addEventListener('input', function(ev) {
        const t = ev && ev.target;
        if (!t || !t.id) return;
        if (!currentTab || !activeMission) return;
        if (t.matches && t.matches('input, select, textarea')) {
          _tabPersistFieldState_(currentTab);
        }
      });
      window.__tabFieldStateHooksBound = true;
    }

    function switchTab(tabNum) {
      if (tabNum < 1 || tabNum > 8) return;

      if (tabNum > 1 && !_isMissionReadyForLock_()) {
        if (window.M) M.toast({ html: 'Offline pack not ready. Press REFRESH and wait for MISSION READY.', classes: 'orange', displayLength: 3200 });
        return;
      }

      const skipTab2GateOnce = window.__skipTab2GateOnce === true;
      if (skipTab2GateOnce) {
        window.__skipTab2GateOnce = false;
      }

      if (!skipTab2GateOnce && currentTab === 2 && tabNum > 2 && typeof window.tab2LoadAircraftThenAdvance === 'function') {
        window.tab2LoadAircraftThenAdvance();
        return;
      }

      if (currentTab === 3 && tabNum > 3 && typeof window.tab3ValidateBeforeProceed === 'function') {
        if (!window.tab3ValidateBeforeProceed()) return;
        // Persist Tab 3 W&B automatically so Flight Detail always has saved wb payload.
        if (typeof window.saveWBLog === 'function') {
          try {
            window.saveWBLog({ silent: true });
          } catch (wbSaveErr) {
            console.warn('Auto-save W&B on tab transition failed', wbSaveErr);
          }
        }
      }

      if (tabNum === 5 && window.releaseTab5Hidden) {
        tabNum = 6;
      }

      // 1. Validation: Don't let them go to Tab 2 without a mission selected
      if (tabNum > 1 && !activeMission) {
        M.toast({html: 'Please select a mission first!', classes: 'orange'});
        return;
      }

      if (currentTab && activeMission) {
        _tabPersistFieldState_(currentTab);
      }

      document.querySelectorAll('.app-tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.step').forEach(s => s.classList.remove('active'));
      document.getElementById('tab' + tabNum).classList.add('active');
      document.getElementById('step' + tabNum).classList.add('active');
      
      currentTab = tabNum;
      document.getElementById('btnBack').style.visibility = (tabNum === 1) ? 'hidden' : 'visible';
      updateFooterButton(tabNum);

      if (tabNum === 1 && isServerAvailable() && typeof renderMissionList === 'function') {
        const now = Date.now();
        if ((now - Number(_lastMissionListServerFetchTs || 0)) > 15000) {
          _lastMissionListServerFetchTs = now;
          renderMissionList({ forceServerList: true, skipPrefetch: true, silent: true });
        }
      }

      // 2. TRIGGER BRIEFING RENDER: If switching to Tab 2, load the data
      if (tabNum === 2) {
        initiateBriefing(activeMission);
      }

      // 3. TRIGGER W&B RENDER: If switching to Tab 3, initialize using first leg flight ID
      if (tabNum === 3) {
        initiateWB(activeMission);
      }

      // 4. TRIGGER PERFORMANCE RENDER: preload Tab 4 with known values from Tab 2/3
      if (tabNum === 4) {
        initiatePerformance(activeMission);
      }

      // 5. TRIGGER RELEASE RENDER: pull from mission + tab 3/4 computed data
      if (tabNum === 5) {
        initiateRelease(activeMission);
      }

      // 6. TRIGGER ENROUTE RENDER
      if (tabNum === 6) {
        initiateEnroute(activeMission);
      }

      // 7. TRIGGER ARRIVAL RENDER
      if (tabNum === 7) {
        initiateArrival(activeMission);
      }

      // 8. TRIGGER DEBRIEF RENDER
      if (tabNum === 8) {
        initiateDebrief(activeMission);
      }

      if (activeMission) {
        setTimeout(function() { _tabRestoreFieldState_(tabNum); }, 220);
      }
    }

    function nextTab() {
      if (currentTab === 1) {
        if (!_isMissionReadyForLock_()) {
          if (window.M) M.toast({ html: 'Offline pack not ready. Press REFRESH and wait for MISSION READY.', classes: 'orange', displayLength: 3200 });
          return;
        }
        tab1AcceptMissionThenAdvance();
        return;
      }

      if (currentTab === 2 && typeof window.tab2LoadAircraftThenAdvance === 'function') {
        window.tab2LoadAircraftThenAdvance();
        return;
      }

      if (currentTab === 3 && typeof window.tab3ValidateBeforeProceed === 'function') {
        if (!window.tab3ValidateBeforeProceed()) return;
      }

      if (currentTab === 5 && typeof window.tab5BrakesReleaseThenAdvance === 'function') {
        window.tab5BrakesReleaseThenAdvance();
        return;
      }
      if (currentTab === 7 && typeof window.tab7OnBlocksThenAdvance === 'function') {
        window.tab7OnBlocksThenAdvance();
        return;
      }
      if (currentTab === 8 && typeof window.submitDebriefLog === 'function') {
        window.submitDebriefLog();
        return;
      }
      switchTab(currentTab + 1);
    }
    function prevTab() {
      let target = currentTab - 1;
      if (target === 5 && window.releaseTab5Hidden) target = 4;
      switchTab(target);
    }

    function updateFooterButton(tab) {
      const btn = document.getElementById('btnNext');
      btn.disabled = false; // always re-enable when changing tabs
      const labels = {
        1: "ACCEPT MISSION", 2: "LOAD AIRCRAFT", 3: "CALCULATE PERFORMANCE",
        4: "REVIEW RISK", 5: "BRAKES RELEASE", 6: "APPROACH BRIEFING",
        7: "LANDED / ON BLOCKS", 8: "SUBMIT DAILY LOG"
      };
      btn.innerText = labels[tab] || "NEXT";
    }
    // Mission Logic
    function renderMissionList(options) {
      const opts = (options && typeof options === 'object') ? options : {};
      const forceServerList = !!opts.forceServerList;
      const skipPrefetch = !!opts.skipPrefetch;
      const silent = !!opts.silent;
      const render = function(missions) {
        const container = document.getElementById('mission-list-container');
        if (!missions || missions.length === 0) {
          container.innerHTML = "<p class='center'>No missions scheduled.</p>";
          updateOfflineCacheStatus();
          return;
        }
        container.innerHTML = missions.map(m => {
          const dateRaw = String(m.date || '');
          const dateObj = new Date(/^\d{4}-\d{2}-\d{2}$/.test(dateRaw) ? dateRaw + 'T00:00:00' : dateRaw);
          const dateStr = !isNaN(dateObj) ? formatStatusDate(dateObj) : '---';
          const origin = String(m.from || m.origin || '').trim().toUpperCase();
          const destination = String(m.to || m.destination || '').trim().toUpperCase();
          const routeLabel = origin && destination
            ? `${origin} → ${destination}`
            : (destination || origin || 'N/A');

          // --- UPDATED LOGIC HERE ---
          // Check if status exists and is "APPROVED" (case insensitive)
          const statusRaw = String(m.status || '').toUpperCase();
          const isFlown = statusRaw === 'FLOWN' || statusRaw === 'COMPLETE' || statusRaw === 'COMPLETED';
          const isPartial = /\/.*LEGS/.test(statusRaw);
          const isApproved = statusRaw === 'APPROVED';
          const isOfflineDraft = statusRaw === 'DRAFT_OFFLINE';
          const isLoadable = isApproved || isOfflineDraft || isFlown || isPartial;
          
          const statusText = isFlown ? 'FLOWN' : (isPartial ? m.status : (isApproved ? 'APPROVED' : (isOfflineDraft ? 'OFFLINE FLIGHT' : 'PENDING')));
          const statusClass = isFlown ? 'text-flown' : (isPartial ? 'text-partial' : (isApproved ? 'text-approved' : 'text-pending'));

          return `
            <div class="mission-row" id="row-${m.id}" data-loadable="${isLoadable ? '1' : '0'}" data-status="${statusRaw}" onclick="selectMission('${m.id}')" style="${isLoadable ? '' : 'opacity:0.7;'}">
              <span class="m-date">${dateStr}</span>
              <span class="m-id">${m.id}</span>
              <span class="m-acft">${m.acft}</span>
              <span class="m-route">${routeLabel}</span>
              <span class="m-status ${statusClass}">${statusText}</span>
              <span class="m-pilot">${m.pilot ? m.pilot.split(' ')[0] : 'N/A'}</span>
            </div>
          `;
        }).join('');
        updateOfflineCacheStatus();
      };

      const cachedMissions = cacheGet(OFFLINE_CACHE_KEYS.SCHEDULED_MISSIONS) || [];

      if (cachedMissions.length) {
        render(cachedMissions);
      }

      if (isServerAvailable()) {
        if (!forceServerList && cachedMissions.length && isOfflineCacheFresh_()) {
          if (!skipPrefetch) prefetchScheduledMissionDetails(cachedMissions, false);
          return;
        }

        google.script.run.withSuccessHandler(missions => {
          cacheSet(OFFLINE_CACHE_KEYS.SCHEDULED_MISSIONS, missions || []);
          render(missions || []);
          if (!skipPrefetch) prefetchScheduledMissionDetails(missions || []);
        }).withFailureHandler(err => {
          const cached = cacheGet(OFFLINE_CACHE_KEYS.SCHEDULED_MISSIONS) || [];
          if (cached.length) {
            if (!silent && window.M) M.toast({ html: 'Loaded cached missions', classes: 'orange' });
            render(cached);
            return;
          }
          const container = document.getElementById('mission-list-container');
          const msg = err && err.message ? err.message : String(err || 'Unknown mission list error');
          container.innerHTML = `<div class='card-panel red-text' style='margin:10px;'>Mission list failed: ${msg}</div>`;
        }).getScheduledMissions();
        return;
      }

      if (cachedMissions.length) {
        render(cachedMissions);
      } else {
        const container = document.getElementById('mission-list-container');
        container.innerHTML = "<p class='center orange-text'>Offline and no cached mission list available.</p>";
        updateOfflineCacheStatus();
      }
    }

    function selectMission(id) {
      const selectedRow = document.getElementById('row-' + id);
      if (selectedRow && String(selectedRow.dataset.loadable || '') !== '1') {
        M.toast({ html: 'Mission is still pending approval', classes: 'orange', displayLength: 2500 });
        return;
      }

      document.querySelectorAll('.mission-row').forEach(row => row.classList.remove('selected'));
      if (selectedRow) selectedRow.classList.add('selected');
      activeMission = id;
      loadActiveLessonContext_();
      checkMissionLegality(id);
      M.toast({ html: 'Selected: ' + id, displayLength: 1500 });
    }

    function openRunwayWalkthroughFromTab1() {
      if (!activeMission) {
        if (window.M) M.toast({ html: 'Select a mission first', classes: 'orange' });
        return;
      }

      fetchMissionDetails(activeMission, function(mission) {
        const firstLeg = mission && Array.isArray(mission.legs) ? (mission.legs[0] || {}) : null;
        const fromIcao = String(firstLeg && firstLeg.from || '').trim().toUpperCase();
        const toIcao = String(firstLeg && firstLeg.to || '').trim().toUpperCase();
        if (!fromIcao) {
          if (window.M) M.toast({ html: 'Mission has no departure ICAO', classes: 'orange' });
          return;
        }
        if (typeof window.openRunwayWalkthrough !== 'function') {
          if (window.M) M.toast({ html: 'Runway walkthrough UI not available yet', classes: 'red' });
          return;
        }
        window.openRunwayWalkthrough(fromIcao, toIcao, 0);
      }, function(err) {
        if (window.M) M.toast({ html: 'Could not load mission for runway walkthrough', classes: 'red' });
      });
    }

    function checkMissionLegality(missionId) {
      const acftArea = document.getElementById('aircraft-status-area');
      const btnNext = document.getElementById('btnNext');

      // Show a small loading state while we fetch mission details
      acftArea.innerHTML = `<div class="center" style="padding:18px;"><div class="preloader-wrapper small active"><div class="spinner-layer spinner-blue-only"><div class="circle-clipper left"><div class="circle"></div></div></div></div><p class="grey-text" style="margin:8px 0 0 0;">Checking aircraft status...</p></div>`;

      fetchMissionDetails(missionId, mission => {
        try {
          console.log('checkMissionLegality mission:', mission);
          const liveAppData = (window.appData && typeof window.appData === 'object') ? window.appData : appData;
          const cachedDropdown = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA) || {};
          console.log('appData.aircraft count:', liveAppData && liveAppData.aircraft ? liveAppData.aircraft.length : 0);
          const acftList = (liveAppData && Array.isArray(liveAppData.aircraft) && liveAppData.aircraft.length)
            ? liveAppData.aircraft
            : (Array.isArray(cachedDropdown.aircraft) ? cachedDropdown.aircraft : []);
          const missionAcft = String((mission && mission.acft) || '').trim().toUpperCase();
          if (!mission || !acftList.length || !missionAcft) {
            acftArea.innerHTML = `<div class="card-panel orange-text">Offline: mission or aircraft cache is incomplete. Re-sync when online.</div>`;
            return;
          }

          // Find the aircraft in our local database
          const acft = acftList.find(function(a) {
            return String((a && a.reg) || '').trim().toUpperCase() === missionAcft;
          });

          if (!acft) {
            acftArea.innerHTML = `<div class="card-panel red-text">Aircraft ${missionAcft} not found in database.</div>`;
            return;
          }

        // GROUNDED LOGIC
        const isGrounded = acft.techStatus === "GROUNDED";
        btnNext.disabled = isGrounded; // Gray out the "Confirm & Proceed" button
        const statusColor = isGrounded ? '#d32f2f' : '#2e7d32'; 
        const statusLabel = isGrounded ? 'GROUNDED - NO FLY' : 'SERVICEABLE';

        // SQUAWK LOGIC (Parsing the JSON)
        let squawkHtml = '<p class="green-text" style="font-size:0.9rem;">✅ No open squawks</p>';
        const openSquawksRaw = String((acft && acft.openSquawks) || '').trim();
        if (openSquawksRaw !== "") {
          try {
            const squawks = JSON.parse(openSquawksRaw);
            const openStatuses = ['OPEN', 'DEFERRED_50_HOUR', 'DEFERRED_100_HOUR', 'DEFERRED_TO_DATE'];
            const visible = Array.isArray(squawks) ? squawks.filter(s => openStatuses.includes(String((s && s.status) || '').trim().toUpperCase())) : [];
            if (visible.length) {
              squawkHtml = visible.map(s => {
                const status = String((s && s.status) || 'OPEN').trim().toUpperCase();
                let statusLabel = status.replace(/_/g, ' ');
                if (status === 'DEFERRED_50_HOUR') statusLabel = 'Deferred 50 Hour';
                if (status === 'DEFERRED_100_HOUR') statusLabel = 'Deferred 100 Hour';
                if (status === 'DEFERRED_TO_DATE') statusLabel = 'Deferred To Date';
                const deferBits = [];
                if (s && s.deferredUntilTach !== '' && s && s.deferredUntilTach != null) deferBits.push('Until Tach ' + s.deferredUntilTach);
                if (s && s.deferredUntilDate) deferBits.push('Until ' + s.deferredUntilDate);
                const dateLabel = (s && (s.reportDate || s.date)) ? (s.reportDate || s.date) : 'NEW';
                return `
                  <div style="border-left: 3px solid #fbc02d; padding: 5px 10px; margin-bottom: 5px; background: #fffde7; font-size: 0.85rem;">
                    <b>${dateLabel}:</b> ${s && s.description ? s.description : ''} <i>(${statusLabel}${deferBits.length ? ' • ' + deferBits.join(' • ') : ''})</i>
                  </div>
                `;
              }).join('');
            }
          } catch (e) {
            squawkHtml = `<div class="orange-text">⚠️ ${openSquawksRaw}</div>`;
          }
        }

        const currentTach = parseFloat(acft.currentTach);
        const nextDue = parseFloat(acft.nextDue);
        const hoursToInsp = parseFloat(acft.hoursToInsp);
        const tachText = isFinite(currentTach) ? currentTach.toFixed(1) : '—';
        const nextDueText = isFinite(nextDue) ? nextDue.toFixed(1) : '—';
        const hrsText = isFinite(hoursToInsp) ? hoursToInsp.toFixed(1) : '—';

        // RENDER THE UI
        acftArea.innerHTML = `
          <div class="card-panel" style="border: 2px solid ${statusColor}; padding: 15px; border-radius: 8px;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
              <span style="font-weight:900; font-size:1.4rem;">${acft.reg}</span>
              <span style="background:${statusColor}; color:white; padding:4px 12px; border-radius:4px; font-weight:bold; font-size:0.8rem;">
                ${statusLabel}
              </span>
            </div>

            <div class="row" style="margin-bottom:0; background: #f5f5f5; border-radius: 4px; padding: 10px 0;">
              <div class="col s4 center">
                <small style="display:block; color:#666;">TACH</small>
                <b>${tachText}</b>
              </div>
              <div class="col s4 center">
                <small style="display:block; color:#666;">NEXT DUE</small>
                <b>${nextDueText}</b>
              </div>
              <div class="col s4 center">
                <small style="display:block; color:#666;">REMAINING</small>
                <b class="${isFinite(hoursToInsp) && hoursToInsp < 10 ? 'red-text' : ''}">${hrsText}</b>
              </div>
            </div>

            <div style="margin-top:15px;">
              <h6 style="font-size:0.7rem; font-weight:bold; color:#999; text-transform:uppercase;">Maintenance Squawks</h6>
              ${squawkHtml}
            </div>
          </div>
        `;
        } catch (e) {
          console.error('checkMissionLegality render failed', e);
          acftArea.innerHTML = `<div class="card-panel orange-text">Offline data parse issue. Continue with cached mission and re-sync when online.</div>`;
        }
        }, err => {
          console.error('getMissionById failed', err);
          acftArea.innerHTML = `<div class="card-panel orange-text">Offline mission detail unavailable. Continue with cached data if needed.</div>`;
          M.toast({ html: 'Mission details unavailable offline', classes: 'orange' });
        });
    }


    let selectedCache = null; // To track which row we clicked

    function renderFuelList() {
      const container = document.getElementById('fuelList');
      if (!appData || !appData.fuelCaches) return;

      // Render rows with ICAO, Location, Qty, and Type
      container.innerHTML = appData.fuelCaches.map(f => `
        <div class="cache-row-clickable" onclick="openAdjustModal('${f.icao}')" 
            style="padding: 12px; border-bottom: 1px solid #eee; cursor: pointer;">
          <div style="display: flex; justify-content: space-between;">
            <span style="font-weight: bold; color: #0b5394;">${f.icao} - ${f.location || 'Unknown'}</span>
            <span class="badge blue white-text" style="border-radius:4px;">${f.qty}L</span>
          </div>
          <small class="grey-text">${f.type || 'AVGAS'}</small>
        </div>
      `).join('');
    }

    function filterFuelList() {
      const input = document.getElementById('fuelSearch').value.toUpperCase();
      const rows = document.getElementsByClassName('cache-row-clickable');
      
      for (let i = 0; i < rows.length; i++) {
        const txt = rows[i].textContent || rows[i].innerText;
        rows[i].style.display = txt.toUpperCase().indexOf(input) > -1 ? "" : "none";
      }
    }
    function openFuelModal() {
      const modalElem = document.getElementById('modalFuel');
      const instance = M.Modal.getInstance(modalElem) || M.Modal.init(modalElem);
      
      // This ensures the list is drawn before the window opens
      renderFuelList(); 
      
      instance.open();
    }

    function closeFuelModals() {
      ['modalFuelAdjust', 'modalFuel'].forEach(function(id) {
        const elem = document.getElementById(id);
        if (!elem || !window.M || !M.Modal) return;
        const instance = M.Modal.getInstance(elem);
        if (instance) instance.close();
      });

      document.querySelectorAll('.modal-overlay').forEach(function(o) { o.remove(); });
      document.body.style.overflow = '';

      const amountEl = document.getElementById('adjustAmount');
      if (amountEl) amountEl.value = '';
      const searchEl = document.getElementById('fuelSearch');
      if (searchEl) searchEl.value = '';
      selectedCache = null;
      setFuelAdjustMode('add');

      try { switchTab(1); } catch (e) { /* stay on current tab if unavailable */ }
    }

    function setFuelAdjustMode(mode) {
      const normalized = mode === 'subtract' ? 'subtract' : 'add';
      const modeEl = document.getElementById('adjustMode');
      const addBtn = document.getElementById('fuelModeAddBtn');
      const subtractBtn = document.getElementById('fuelModeSubtractBtn');
      const labelEl = document.getElementById('adjustAmountLabel');

      if (modeEl) modeEl.value = normalized;
      if (labelEl) labelEl.textContent = normalized === 'subtract' ? 'Liters to remove' : 'Liters to add';
      if (addBtn) {
        addBtn.style.background = normalized === 'add' ? '#2e7d32' : '#c8e6c9';
        addBtn.style.color = normalized === 'add' ? '#fff' : '#2e7d32';
      }
      if (subtractBtn) {
        subtractBtn.style.background = normalized === 'subtract' ? '#c62828' : '#ffcdd2';
        subtractBtn.style.color = normalized === 'subtract' ? '#fff' : '#8b0000';
      }
    }

    function setFuelAmount(amount) {
      const amountEl = document.getElementById('adjustAmount');
      if (!amountEl) return;
      amountEl.value = String(Math.max(1, Math.round(Number(amount) || 0)));
      amountEl.focus();
    }

    function openAdjustModal(icao) {
      selectedCache = appData.fuelCaches.find(f => f.icao === icao);
      if (!selectedCache) return;

      // Set Labels (Rounding current qty just in case)
      document.getElementById('adjustTitle').innerText = `Fuel: ${selectedCache.icao}`;
      document.getElementById('adjustSub').innerText = `${selectedCache.location} (${Math.round(selectedCache.qty)}L Available)`;

      // Populate Pilots (Notice the p.name)
      const pilotSelect = document.getElementById('adjustPilot');
      if (appData.pilots) {
        pilotSelect.innerHTML = '<option value="" disabled selected>Select Pilot</option>' + 
          appData.pilots.map(p => `<option value="${p.name}">${p.name}</option>`).join('');
      }

      // Populate Aircraft
      const acftSelect = document.getElementById('adjustAcft');
      if (appData.aircraft) {
        acftSelect.innerHTML = '<option value="" disabled selected>Select Aircraft</option>' + 
          appData.aircraft.map(a => `<option value="${a.reg}">${a.reg}</option>`).join('');
      }

      const amountEl = document.getElementById('adjustAmount');
      if (amountEl) amountEl.value = '';
      setFuelAdjustMode('add');

      const elem = document.getElementById('modalFuelAdjust');
      const instance = M.Modal.getInstance(elem) || M.Modal.init(elem);
      instance.open();
    }
    function submitFuelAdjustment() {
      const amountInput = document.getElementById('adjustAmount').value;
      const enteredAmount = Math.round(parseFloat(amountInput));
      const mode = ((document.getElementById('adjustMode') || {}).value || 'add') === 'subtract' ? 'subtract' : 'add';
      const amount = mode === 'subtract' ? -Math.abs(enteredAmount) : Math.abs(enteredAmount);
      const pilot = document.getElementById('adjustPilot').value;
      const acft = document.getElementById('adjustAcft').value;

      // Validation
      if (!pilot || !acft || isNaN(amount) || amount === 0) {
        M.toast({ html: 'Fill all fields: Pilot, Aircraft, and Amount' });
        return;
      }

      // Disable button to prevent double submit
      const btn = document.querySelector('#modalFuelAdjust .btn');
      if (btn) btn.disabled = true;

      const onFuelDone = () => {
          M.toast({ html: `Logged ${amount}L for ${selectedCache.icao}` });

          // Update local cache qty and refresh list
          selectedCache.qty += amount;
          renderFuelList();

          // Persist the updated qty back to localStorage so it survives an offline reload
          try {
            const cached = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA);
            if (cached) {
              cached.fuelCaches = appData.fuelCaches;
              cacheSet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA, cached);
            }
          } catch(e) { /* non-fatal */ }

          // --- PROPER MODAL TEARDOWN ---
          const adjustElem = document.getElementById('modalFuelAdjust');
          const searchElem = document.getElementById('modalFuel');

          const adjustModal = M.Modal.getInstance(adjustElem);
          if (adjustModal) {
            adjustModal.close();
            adjustModal.destroy();
          }

          const searchModal = M.Modal.getInstance(searchElem);
          if (searchModal) {
            searchModal.close();
            searchModal.destroy();
          }

          // HARD CLEANUP: remove any orphaned overlays
          document.querySelectorAll('.modal-overlay').forEach(o => o.remove());

          // Reset input + button
          document.getElementById('adjustAmount').value = '';
          if (btn) btn.disabled = false;

          // Return to Mission Dashboard
            closeFuelModals();

      };

      window.runOrQueueServerAction({
        method: 'logFuelChange',
        args: [selectedCache.icao, amount, acft, pilot],
        label: 'Fuel change'
      }, {
        onSuccess: onFuelDone,
        onQueued: onFuelDone,
        onFailure: function(err) {
          M.toast({ html: 'Error: ' + (err && err.message ? err.message : String(err || 'unknown')) });
          if (btn) btn.disabled = false;
        }
      });
    }






    // Passenger Logic
    function openPaxModal() {
      const elem = document.getElementById('modalPax');
      const instance = M.Modal.getInstance(elem) || M.Modal.init(elem);
      instance.open();
    }

    function _parseCoordinate(str) {
      if (str === null || str === undefined || str === '') return NaN;
      // Normalize Brazilian comma decimal separator (-2,123456 → -2.123456)
      const s = String(str).trim().replace(/(\d),(\d)/g, '$1.$2');
      // Plain decimal (e.g. -2.123456 or 6.0)
      if (/^-?\d+(\.\d+)?$/.test(s)) return parseFloat(s);
      // Normalize: strip degree/min/sec symbols, collapse spaces
      const norm = s.replace(/[°'"]/g, ' ').replace(/\s+/g, ' ').trim();
      // DMS: [NSEW] DD MM SS.ss or DD MM SS.ss [NSEW]
      const dmsRe = /^([NSEWnsew])?\s*(\d+)\s+(\d+)\s+(\d+(?:\.\d+)?)\s*([NSEWnsew])?$/;
      let m = norm.match(dmsRe);
      if (m) {
        const card = (m[1] || m[5] || '').toUpperCase();
        const dec = parseFloat(m[2]) + parseFloat(m[3]) / 60 + parseFloat(m[4]) / 3600;
        return (card === 'S' || card === 'W') ? -dec : dec;
      }
      // Decimal-minutes: [NSEW] DD MM.mmm or DD MM.mmm [NSEW]
      const dmRe = /^([NSEWnsew])?\s*(\d+)\s+(\d+(?:\.\d+)?)\s*([NSEWnsew])?$/;
      m = norm.match(dmRe);
      if (m) {
        const card = (m[1] || m[4] || '').toUpperCase();
        const dec = parseFloat(m[2]) + parseFloat(m[3]) / 60;
        return (card === 'S' || card === 'W') ? -dec : dec;
      }
      // Last resort
      const fallback = parseFloat(s);
      return isFinite(fallback) ? fallback : NaN;
    }

    function openWaypointModal() {
      const elem = document.getElementById('modalWaypoint');
      const instance = M.Modal.getInstance(elem) || M.Modal.init(elem);
      instance.open();
      _wpUpdatePreview();
      setTimeout(function() { if (_wpPreviewMap) _wpPreviewMap.invalidateSize(); }, 350);
    }

    let _wpPreviewMap = null;
    let _wpPreviewMarker = null;

    function _wpEnsureLeaflet(cb) {
      if (window.L && typeof window.L.map === 'function') { cb(); return; }
      if (!document.getElementById('wp-leaflet-css')) {
        const link = document.createElement('link');
        link.id = 'wp-leaflet-css'; link.rel = 'stylesheet';
        link.href = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css';
        document.head.appendChild(link);
      }
      const existing = document.getElementById('wp-leaflet-js');
      if (existing) {
        const t = setInterval(function() {
          if (window.L && typeof window.L.map === 'function') { clearInterval(t); cb(); }
        }, 100);
        return;
      }
      const script = document.createElement('script');
      script.id = 'wp-leaflet-js';
      script.src = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js';
      script.onload = cb;
      document.head.appendChild(script);
    }

    function _wpUpdatePreview() {
      const lat = _parseCoordinate(((document.getElementById('wp_lat') || {}).value) || '');
      const lon = _parseCoordinate(((document.getElementById('wp_lon') || {}).value) || '');
      const mapDiv = document.getElementById('wp-preview-map');
      const coordsDiv = document.getElementById('wp-preview-coords');
      if (!mapDiv) return;
      if (isNaN(lat) || isNaN(lon) || lat < -90 || lat > 90 || lon < -180 || lon > 180) {
        mapDiv.style.display = 'none';
        if (coordsDiv) coordsDiv.style.display = 'none';
        return;
      }
      mapDiv.style.display = 'block';
      if (coordsDiv) {
        coordsDiv.style.display = 'block';
        coordsDiv.textContent = '\u2192 ' + lat.toFixed(6) + ', ' + lon.toFixed(6);
      }
      _wpEnsureLeaflet(function() {
        if (!_wpPreviewMap) {
          _wpPreviewMap = L.map('wp-preview-map', { zoomControl: true, attributionControl: false });
          const tileUrl = 'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png';
          if (typeof _idbGetTile === 'function') {
            const OfflineLayer = L.TileLayer.extend({
              createTile: function(coords, done) {
                const img = document.createElement('img'); img.alt = '';
                const netUrl = tileUrl
                  .replace('{s}', ['a','b','c'][Math.floor(Math.random()*3)])
                  .replace('{z}', coords.z).replace('{x}', coords.x).replace('{y}', coords.y);
                _idbGetTile(coords.z, coords.x, coords.y).then(function(blob) {
                  if (blob) {
                    const u = URL.createObjectURL(blob);
                    img.onload = function() { URL.revokeObjectURL(u); done(null, img); };
                    img.onerror = function() { URL.revokeObjectURL(u); img.src = netUrl; };
                    img.src = u;
                  } else {
                    img.onload = function() { done(null, img); };
                    img.onerror = function(e) { done(e, img); };
                    img.src = netUrl;
                  }
                }).catch(function() {
                  img.onload = function() { done(null, img); };
                  img.onerror = function(e) { done(e, img); };
                  img.src = netUrl;
                });
                return img;
              }
            });
            new OfflineLayer(tileUrl, { maxZoom: 16, minZoom: 2 }).addTo(_wpPreviewMap);
          } else {
            L.tileLayer(tileUrl, { maxZoom: 16, minZoom: 2 }).addTo(_wpPreviewMap);
          }
        }
        _wpPreviewMap.setView([lat, lon], 11);
        if (_wpPreviewMarker) {
          _wpPreviewMarker.setLatLng([lat, lon]);
        } else {
          _wpPreviewMarker = L.marker([lat, lon]).addTo(_wpPreviewMap);
        }
        setTimeout(function() { if (_wpPreviewMap) _wpPreviewMap.invalidateSize(); }, 50);
      });
    }

    function saveNewWaypoint() {
      const data = {
        wp_id: String(document.getElementById('wp_id').value || '').trim().toUpperCase(),
        latitude: _parseCoordinate(document.getElementById('wp_lat').value),
        longitude: _parseCoordinate(document.getElementById('wp_lon').value),
        type: String(document.getElementById('wp_type').value || 'FIX').trim().toUpperCase()
      };

      if (!data.wp_id) {
        M.toast({ html: 'Waypoint ID is required' });
        return;
      }
      if (isNaN(data.latitude) || data.latitude < -90 || data.latitude > 90) {
        M.toast({ html: 'Latitude: unrecognized format or out of range (±90)' });
        return;
      }
      if (isNaN(data.longitude) || data.longitude < -180 || data.longitude > 180) {
        M.toast({ html: 'Longitude: unrecognized format or out of range (±180)' });
        return;
      }
      if (data.type !== 'FIX' && data.type !== 'WATER RUNWAY') {
        M.toast({ html: 'Type must be FIX or WATER RUNWAY' });
        return;
      }

      const btn = document.querySelector('#modalWaypoint button');
      if (btn) btn.disabled = true;

      const onWaypointDone = function() {
        M.toast({ html: 'Waypoint saved: ' + data.wp_id, classes: 'green' });

        if (!appData.waypoints) appData.waypoints = [];
        appData.waypoints.push({
          wp_id: data.wp_id,
          lat: Number(data.latitude),
          lon: Number(data.longitude),
          type: data.type
        });

        try {
          const cached = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA);
          if (cached) {
            cached.waypoints = appData.waypoints;
            cacheSet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA, cached);
          }
        } catch (e) { /* non-fatal */ }

        const modal = M.Modal.getInstance(document.getElementById('modalWaypoint'));
        if (modal) modal.close();

        const idEl = document.getElementById('wp_id');
        const latEl = document.getElementById('wp_lat');
        const lonEl = document.getElementById('wp_lon');
        const typeEl = document.getElementById('wp_type');
        if (idEl) idEl.value = '';
        if (latEl) latEl.value = '';
        if (lonEl) lonEl.value = '';
        if (typeEl) typeEl.value = 'FIX';
        if (btn) btn.disabled = false;

        switchTab(1);
      };

      window.runOrQueueServerAction({
        method: 'addWaypointToDatabase',
        args: [data],
        label: 'Waypoint save'
      }, {
        onSuccess: function(res) {
          if (res && res.success) {
            onWaypointDone();
          } else {
            M.toast({ html: 'Error saving waypoint' });
            if (btn) btn.disabled = false;
          }
        },
        onQueued: onWaypointDone,
        onFailure: function(err) {
          M.toast({ html: 'Error: ' + (err && err.message ? err.message : String(err || 'unknown')), classes: 'red' });
          if (btn) btn.disabled = false;
        }
      });
    }

    function saveNewPax() {
      const data = {
        name: document.getElementById('pax_name').value,
        id_type: document.getElementById('pax_id_type').value,
        id_num: document.getElementById('pax_id_num').value,
        dob: document.getElementById('pax_dob').value,
        weight: document.getElementById('pax_weight').value,
        gender: document.getElementById('pax_gender').value,
        phone: document.getElementById('pax_phone').value
      };

      if (!data.name || !data.weight) {
        M.toast({html: 'Name and Weight are required'});
        return;
      }

      // 1. Disable the button to prevent multiple clicks
      const btn = document.querySelector('#modalPax button');
      if (btn) btn.disabled = true;

      const onPaxDone = () => {
          M.toast({html: 'Passenger Registered!'});

          // Add to in-memory appData.passengers immediately so they're usable in pax dropdowns
          if (!appData.passengers) appData.passengers = [];
          appData.passengers.push({
            name: data.name,
            weight: parseFloat(data.weight) || 80,
            gender: data.gender || 'U',
            dob: data.dob || '',
            phone: data.phone || ''
          });
          appData.passengers.sort((a, b) => String(a.name).localeCompare(String(b.name)));

          // Persist the new passenger into the DROPDOWN_DATA cache so they survive an offline reload
          try {
            const cached = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA);
            if (cached) {
              cached.passengers = appData.passengers;
              cacheSet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA, cached);
            }
          } catch(e) { /* non-fatal */ }

          // Close the modal
          const paxModal = M.Modal.getInstance(document.getElementById('modalPax'));
          if (paxModal) paxModal.close();

          // Clear all inputs and re-enable button
          document.querySelectorAll('#modalPax input').forEach(i => i.value = '');
          if (btn) btn.disabled = false;

          // Snap back to the Mission Dashboard
          switchTab(1);
      };

      window.runOrQueueServerAction({
        method: 'savePassengerToDB',
        args: [data],
        label: 'Passenger save'
      }, {
        onSuccess: (res) => {
          if(res.success) {
            onPaxDone();
          } else {
            M.toast({html: 'Error: ' + res.error});
            if (btn) btn.disabled = false; // Re-enable so they can try to fix the error
          }
        },
        onQueued: onPaxDone,
        onFailure: function(err) {
          M.toast({html: 'Error: ' + (err && err.message ? err.message : String(err || 'unknown'))});
          if (btn) btn.disabled = false;
        }
      });
    }
    function initiateBriefing(missionId) {
      const container = document.getElementById('tab2'); 
      
      if(container) {
        container.innerHTML = `
          <div class="center" style="margin-top:50px;">
            <div class="preloader-wrapper small active">
              <div class="spinner-layer spinner-blue-only">
                <div class="circle-clipper left"><div class="circle"></div></div>
              </div>
            </div>
            <p class="grey-text">Loading Briefing Data...</p>
          </div>`;
      }

      fetchMissionDetails(missionId, mission => {
          // --- DEBUGGING LINE START ---
          console.log("Mission Data Received from Server:", mission);
          // --- DEBUGGING LINE END ---

          if (typeof setupBriefing === "function") {
            setupBriefing(mission);
          } else {
            console.error("Critical: setupBriefing function not found.");
            container.innerHTML = "<p class='red-text center'>Error: Briefing Script not loaded.</p>";
          }
        }, err => {
          console.error('initiateBriefing fetchMissionDetails failed', err);
          M.toast({html: 'Failed loading briefing mission', classes: 'red'});
        });
    }

    // Returns the pilot-selected active leg, or the first non-COMPLETE leg, or legs[0]
    function getActiveLeg_(mission) {
      if (!mission || !Array.isArray(mission.legs) || !mission.legs.length) return null;
      const legs = mission.legs;
      // 1. Honour pilot's explicit Tab-2 selection stored in global
      if (window.activeLegFlightId) {
        const found = legs.find(l => l.flightLegId === window.activeLegFlightId);
        if (found) return found;
      }
      // 2. First non-complete leg
      const firstPending = legs.find(l => (l.logStatus || 'PENDING') !== 'COMPLETE');
      if (firstPending) return firstPending;
      // 3. Fall back to last leg
      return legs[legs.length - 1];
    }

    function initiateWB(missionId) {
      const tab3El = document.getElementById('tab3');
      if (tab3El && !document.getElementById('wb-container')) {
        tab3El.innerHTML = "<p class='red-text center' style='padding:20px;'>W&B view not loaded.</p>";
        return;
      }

      const runSetup = function(flightId) {
        if (!flightId) {
          M.toast({html: 'No flight leg found for W&B', classes: 'orange'});
          return;
        }
        if (typeof window.setupWB === 'function') {
          window.setupWB(flightId);
        } else {
          console.error('setupWB function not found.');
          M.toast({html: 'W&B script not loaded', classes: 'red'});
        }
      };

      // Preferred source: already-loaded briefing mission — use active leg, not always legs[0]
      if (window.currentBriefingMission && Array.isArray(window.currentBriefingMission.legs) && window.currentBriefingMission.legs.length > 0) {
        const activeLeg = getActiveLeg_(window.currentBriefingMission);
        runSetup(activeLeg ? (activeLeg.flightLegId || '') : '');
        return;
      }

      // Fallback: fetch mission data and use active leg
      fetchMissionDetails(missionId, function(mission) {
          if (mission) window.currentBriefingMission = mission;
          const activeLeg = getActiveLeg_(mission);
          runSetup(activeLeg ? (activeLeg.flightLegId || '') : '');
        }, function(err) {
          console.error('initiateWB getMissionById failed', err);
          M.toast({html: 'Failed loading mission for W&B', classes: 'red'});
        });
    }

    function initiatePerformance(missionId) {
      const runSetup = function(mission) {
        if (!mission || !mission.legs || !mission.legs.length) {
          M.toast({html: 'Mission legs missing for performance', classes: 'orange'});
          return;
        }

        const firstLeg = getActiveLeg_(mission) || mission.legs[0] || {};
        const fromIcao = String(firstLeg.from || '').trim();
        const flightId = String(firstLeg.flightLegId || '').trim();

        let wbWeight = 0;
        try {
          if (typeof window.calculateWB === 'function' && window.wbData && window.wbData.flightId) {
            const wb = window.calculateWB();
            wbWeight = Number(wb && wb.totalWeight ? wb.totalWeight : 0);
          }
        } catch (e) {
          wbWeight = 0;
        }

        const fallbackWeight = Number(firstLeg.limit || firstLeg.payload || 0);
        const grossWeight = wbWeight > 0 ? wbWeight : fallbackWeight;

        const acftReg = String(mission.acft || '').trim();
        const liveAircraft = (appData && Array.isArray(appData.aircraft)) ? appData.aircraft : [];
        let cachedAircraft = [];
        try {
          const cachedDropdown = cacheGet(OFFLINE_CACHE_KEYS.DROPDOWN_DATA) || {};
          cachedAircraft = Array.isArray(cachedDropdown.aircraft) ? cachedDropdown.aircraft : [];
        } catch (e) { cachedAircraft = []; }
        const aircraftList = liveAircraft.length ? liveAircraft : cachedAircraft;
        const acftObj = aircraftList.find(a => String(a.reg || '').trim().toUpperCase() === acftReg.toUpperCase()) || null;
        // Use TYPE_FOR_PERFORMANCE if present, otherwise fall back to AIRCRAFT_TYPE
        const typeForPerf = acftObj ? String(acftObj.typeForPerformance || '').trim() : '';
        const aircraftType = typeForPerf || (acftObj ? String(acftObj.aircraftType || '') : '');

        if (typeof window.setupPerformanceTab === 'function') {
          window.setupPerformanceTab({
            missionId: mission.id,
            flightId: flightId,
            fromIcao: fromIcao,
            weightKg: grossWeight,
            aircraftReg: acftReg,
            aircraftType: aircraftType
          });
        } else {
          M.toast({html: 'Performance tab script not loaded', classes: 'red'});
        }
      };

      if (window.currentBriefingMission && window.currentBriefingMission.id === missionId) {
        runSetup(window.currentBriefingMission);
        return;
      }

      fetchMissionDetails(missionId, function(mission) {
          if (mission) window.currentBriefingMission = mission;
          runSetup(mission);
        }, function(err) {
          console.error('initiatePerformance getMissionById failed', err);
          M.toast({html: 'Failed loading mission for performance', classes: 'red'});
        });
    }

    function initiateRelease(missionId) {
      const runSetup = function(mission) {
        if (!mission || !mission.legs || !mission.legs.length) {
          M.toast({html: 'Mission legs missing for release', classes: 'orange'});
          return;
        }
        if (typeof window.setupReleaseTab === 'function') {
          window.setupReleaseTab(mission);
        } else {
          M.toast({html: 'Release tab script not loaded', classes: 'red'});
        }
      };

      if (window.currentBriefingMission && window.currentBriefingMission.id === missionId) {
        runSetup(window.currentBriefingMission);
        return;
      }

      fetchMissionDetails(missionId, function(mission) {
          if (mission) window.currentBriefingMission = mission;
          runSetup(mission);
        }, function(err) {
          console.error('initiateRelease getMissionById failed', err);
          M.toast({html: 'Failed loading mission for release', classes: 'red'});
        });
    }

    function initiateEnroute(missionId) {
      const runSetup = function(mission, attempt) {
        if (!mission || !mission.legs || !mission.legs.length) {
          M.toast({html: 'Mission legs missing for enroute tab', classes: 'orange'});
          return;
        }
        if (typeof window.setupEnrouteTab === 'function') {
          window.setupEnrouteTab(mission);
        } else {
          const n = Number(attempt || 0);
          if (n < 1) {
            setTimeout(function() { runSetup(mission, n + 1); }, 350);
            return;
          }
          const readyFlag = window.__enrouteScriptReady ? 'ready-flag=true' : 'ready-flag=false';
          M.toast({html: 'Enroute tab script not loaded. Try refresh (' + readyFlag + ')', classes: 'red', displayLength: 5000});
        }
      };

      if (window.currentBriefingMission && window.currentBriefingMission.id === missionId) {
        runSetup(window.currentBriefingMission, 0);
        return;
      }

      fetchMissionDetails(missionId, function(mission) {
          if (mission) window.currentBriefingMission = mission;
          runSetup(mission, 0);
        }, function(err) {
          console.error('initiateEnroute getMissionById failed', err);
          M.toast({html: 'Failed loading mission for enroute', classes: 'red'});
        });
    }

    function initiateArrival(missionId) {
      const runSetup = function(mission) {
        if (!mission || !mission.legs || !mission.legs.length) {
          M.toast({html: 'Mission legs missing for arrival tab', classes: 'orange'});
          return;
        }
        if (typeof window.setupArrivalTab === 'function') {
          window.setupArrivalTab(mission);
        } else {
          M.toast({html: 'Arrival tab script not loaded', classes: 'red'});
        }
      };

      if (window.currentBriefingMission && window.currentBriefingMission.id === missionId) {
        runSetup(window.currentBriefingMission);
        return;
      }

      fetchMissionDetails(missionId, function(mission) {
          if (mission) window.currentBriefingMission = mission;
          runSetup(mission);
        }, function(err) {
          console.error('initiateArrival getMissionById failed', err);
          M.toast({html: 'Failed loading mission for arrival', classes: 'red'});
        });
    }

    function initiateDebrief(missionId) {
      const runSetup = function(mission) {
        if (!mission || !mission.legs || !mission.legs.length) {
          M.toast({html: 'Mission legs missing for debrief tab', classes: 'orange'});
          return;
        }
        if (typeof window.setupDebriefTab === 'function') {
          window.setupDebriefTab(mission);
        } else {
          M.toast({html: 'Debrief tab script not loaded', classes: 'red'});
        }
      };

      if (window.currentBriefingMission && String(window.currentBriefingMission.id || '').trim() === String(missionId || '').trim()) {
        runSetup(window.currentBriefingMission);
        return;
      }

      fetchMissionDetails(missionId, function(mission) {
          if (mission) window.currentBriefingMission = mission;
          runSetup(mission);
        }, function(err) {
          console.error('initiateDebrief getMissionById failed', err);
          M.toast({html: 'Failed loading mission for debrief', classes: 'red'});
        });
    }
    
    function OLD_initiateBriefing(missionId) {
      // Show a loader in the briefing area
      const container = document.getElementById('briefing-content');
      if(container) container.innerHTML = '<div class="center" style="margin-top:50px;"><div class="preloader-wrapper small active"><div class="spinner-layer spinner-blue-only"><div class="circle-clipper left"><div class="circle"></div></div></div></div></div>';

      fetchMissionDetails(missionId, mission => {
        if (typeof setupBriefing === "function") {
          setupBriefing(mission);
        } else {
          console.error("setupBriefing function not found. Is Tab2_Briefing.html included correctly?");
        }
      }, err => {
        console.error('OLD_initiateBriefing fetchMissionDetails failed', err);
      });
    }

  
