
    // --- STATE ---
    let supMissionId = null;
    let supPilot = null;
    let lastLoadedData = null; // Store data to allow re-rendering without server call
    let supInstructors = [];
    let supervisorRunwayCheckCtx = null;
    let supervisorRefreshTimer = null;
    let supervisorTabActive = false;
    let supervisorPasswordPromptState = { resolve: null };

    function supFormatDateMonDayYear_(value, fallback) {
      const s = String(value || '').trim();
      // Parse date-only strings as local time (not UTC) to avoid timezone off-by-one
      const d = /^\d{4}-\d{2}-\d{2}$/.test(s) ? new Date(s + 'T00:00:00') : new Date(s);
      if (isNaN(d.getTime())) return fallback || '';
      const months = ['Jan.', 'Feb.', 'Mar.', 'Apr.', 'May.', 'Jun.', 'Jul.', 'Aug.', 'Sep.', 'Oct.', 'Nov.', 'Dec.'];
      return `${months[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
    }

    function openSupervisorPasswordPrompt_(subtitle) {
      return new Promise(function(resolve) {
        supervisorPasswordPromptState.resolve = resolve;
        const overlay = document.getElementById('supervisor-password-overlay');
        const input = document.getElementById('supervisor-password-input');
        const sub = document.getElementById('supervisor-password-sub');
        const err = document.getElementById('supervisor-password-error');
        if (sub) sub.textContent = subtitle || 'Enter supervisor password to continue.';
        if (input) input.value = '';
        if (err) { err.style.display = 'none'; err.textContent = ''; }
        if (overlay) overlay.style.display = 'flex';
        setTimeout(function() { if (input) input.focus(); }, 20);
      });
    }

    function closeSupervisorPasswordPrompt_(confirmed) {
      const overlay = document.getElementById('supervisor-password-overlay');
      const input = document.getElementById('supervisor-password-input');
      if (overlay) overlay.style.display = 'none';
      const resolver = supervisorPasswordPromptState.resolve;
      supervisorPasswordPromptState.resolve = null;
      if (resolver) resolver(confirmed ? String(input && input.value || '') : '');
    }

    function submitSupervisorPasswordPrompt_() {
      const input = document.getElementById('supervisor-password-input');
      const err = document.getElementById('supervisor-password-error');
      const value = String(input && input.value || '');
      if (!value) {
        if (err) { err.textContent = 'Password required.'; err.style.display = 'block'; }
        if (input) input.focus();
        return;
      }
      closeSupervisorPasswordPrompt_(true);
    }

    function supervisorPasswordKeypadAppend_(digit) {
      const input = document.getElementById('supervisor-password-input');
      if (!input) return;
      input.value = String(input.value || '') + String(digit || '');
      try { input.dispatchEvent(new Event('input', { bubbles: true })); } catch (e) {}
      try { input.focus(); } catch (e2) {}
    }

    function supervisorPasswordKeypadBackspace_() {
      const input = document.getElementById('supervisor-password-input');
      if (!input) return;
      input.value = String(input.value || '').slice(0, -1);
      try { input.dispatchEvent(new Event('input', { bubbles: true })); } catch (e) {}
      try { input.focus(); } catch (e2) {}
    }

    function supervisorPasswordKeypadClear_() {
      const input = document.getElementById('supervisor-password-input');
      if (!input) return;
      input.value = '';
      try { input.dispatchEvent(new Event('input', { bubbles: true })); } catch (e) {}
      try { input.focus(); } catch (e2) {}
    }

    async function supervisorRequirePassword_(subtitle) {
      const password = await openSupervisorPasswordPrompt_(subtitle);
      return String(password || '');
    }

    document.addEventListener('keydown', function(evt) {
      const overlay = document.getElementById('supervisor-password-overlay');
      if (!overlay || overlay.style.display !== 'flex') return;
      if (evt.key === 'Escape') {
        evt.preventDefault();
        closeSupervisorPasswordPrompt_(false);
        return;
      }
      if (evt.key === 'Enter') {
        evt.preventDefault();
        submitSupervisorPasswordPrompt_();
      }
    });

    function onSupervisorTabShown() {
      supervisorTabActive = true;
      if (typeof window.openRunwayBriefingCard !== 'function' && typeof window.__rbcOpenRef === 'function') {
        window.openRunwayBriefingCard = window.__rbcOpenRef;
      }
      if (typeof window.runwayBriefingDeepProbe === 'function') {
        try { window.runwayBriefingDeepProbe(); } catch (e) {}
      }
      refreshSupervisorDashboard(false);
      if (supervisorRefreshTimer) clearInterval(supervisorRefreshTimer);
      supervisorRefreshTimer = setInterval(function() {
        refreshSupervisorDashboard(true);
      }, 12000);
    }

    function onSupervisorTabHidden() {
      supervisorTabActive = false;
      if (supervisorRefreshTimer) {
        clearInterval(supervisorRefreshTimer);
        supervisorRefreshTimer = null;
      }
    }

    function refreshSupervisorDashboard(isSilent) {
      if (!(google.script.run && google.script.run.getSupervisorDashboard)) return;
      google.script.run
        .withSuccessHandler(function(data) {
          initSidebar(data);
          if (supMissionId && supervisorTabActive) {
            google.script.run
              .withSuccessHandler(function(details) {
                if (details && details.mission && details.mission.id === supMissionId) {
                  lastLoadedData = details;
                  renderDetails(details);
                }
              })
              .getMissionDetailsForSupervisor(supMissionId);
          }
        })
        .withFailureHandler(function(err) {
          if (!isSilent) showSupervisorError(err);
        })
        .getSupervisorDashboard();
    }

    document.addEventListener('DOMContentLoaded', () => {
      // Ensure we have a server function for this, or use dummy data for now
      if(google.script.run && google.script.run.getSupervisorDashboard) {
          refreshSupervisorDashboard(false);
      }

      window.addEventListener('mission-saved', function() {
        refreshSupervisorDashboard(true);
      });

      window.addEventListener('storage', function(evt) {
        if (!evt || evt.key !== 'mba_mission_saved_at') return;
        refreshSupervisorDashboard(true);
      });
    });

    function showSupervisorError(err) { M.toast({html: 'Error: ' + err.message}); }
    
    // --- SIDEBAR ---
    function initSidebar(data) {
      if(!data) return;
      const userDisplayEl = document.getElementById('userDisplay');
      if (userDisplayEl) userDisplayEl.innerText = data.user || 'Admin';
      supInstructors = Array.isArray(data.instructors) ? data.instructors.slice() : [];
      const sb = document.getElementById('sbList');
      sb.innerHTML = `<div class="sb-header"><h6 style="margin:0">DISPATCH SUPERVISOR</h6><div style="display:flex;justify-content:space-between;align-items:center;margin-top:6px;gap:8px;"><small>${data.user || 'Admin'}</small><button type="button" onclick="refreshSupervisorDashboard(false)" style="border:none;border-radius:6px;background:#0b5394;color:#fff;padding:4px 8px;font-size:0.72rem;font-weight:800;cursor:pointer;">REFRESH</button></div></div>`;
      
      if(!data.missions) return;

      data.missions.forEach(m => {
        const routeSummary = m.legs.map(leg => leg.split('(')[0].trim()).join(' âž ');
        const div = document.createElement('div');
        div.className = 'mission-item';
        div.setAttribute('data-mission-id', String(m.id || ''));
        div.innerHTML = `
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:5px;">
            <div style="font-weight:bold; font-size:1.1rem;">${m.id}</div>
            <span class="m-status ${
              m.status === 'APPROVED' ? 'st-approved' :
              m.status === 'FLOWN' ? 'st-flown' :
              /\/.*LEGS/.test(m.status) ? 'st-partial' :
              'st-pending'
            }">${m.status}</span>
          </div>
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:5px; font-size:0.9rem;">
             <div style="font-weight:bold; color:#4fc3f7;">${m.acft}</div> 
             <div style="opacity:0.7;">${supFormatDateMonDayYear_(m.date, '--')}</div>
          </div>
          <div style="font-size:0.85rem; line-height:1.4em; opacity:0.9;">
            <div style="margin-bottom:2px;">${m.pilot}</div>
            <div style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis; color:#bbb;">
               <span style="font-size:0.8em">âœˆ</span> ${routeSummary}
            </div>
          </div>
          ${m.warnings ? `<div style="margin-top:6px; color:#ff5252; font-size:0.75rem; font-weight:bold; background:rgba(255,0,0,0.1); padding:4px; border-radius:4px;">âš  ${m.warnings}</div>` : ''}
        `;
        if (supMissionId && m.id === supMissionId) div.classList.add('active');
        div.onclick = () => loadSupervisorMission(m.id, div);
        sb.appendChild(div);
      });

      if (supMissionId && !data.missions.some(function(m) { return m.id === supMissionId; })) {
        supMissionId = null;
        document.getElementById('mTitle').innerText = 'Select a Mission';
        document.getElementById('detailPanel').innerHTML = '<p style="text-align:center; margin-top:50px; color:#999;">Select a mission from the left.</p>';
        document.getElementById('timelineContainer').innerHTML = '';
      }
    }

    function loadSupervisorMission(id, el) {
      document.querySelectorAll('.mission-item').forEach(d => d.classList.remove('active'));
      if (el && el.classList) el.classList.add('active');
      supMissionId = id;
      document.getElementById('mTitle').innerText = "Loading " + id + "...";
      google.script.run.withSuccessHandler(data => {
        lastLoadedData = data; 
        renderDetails(data);
      }).withFailureHandler(showSupervisorError).getMissionDetailsForSupervisor(id);
    }
    // --- DETAILS RENDERER ---
    function renderDetails(data) {
      const m = data.mission;
      const t = data.timeline;
      supPilot = m.meta.pilot;

      const normalizedLegs = (Array.isArray(m.legs) ? m.legs : []).map(function(leg) {
        const safeLeg = leg || {};
        const num = function(v, d) {
          const n = Number(v);
          return isFinite(n) ? n : d;
        };
        const waypointsText = Array.isArray(safeLeg.waypoints)
          ? safeLeg.waypoints.filter(Boolean).join(',')
          : String(safeLeg.waypoints || '');
        return {
          flightLegId: String(safeLeg.flightLegId || ''),
          from: String(safeLeg.from || '?'),
          to: String(safeLeg.to || '?'),
          waypointsText: waypointsText,
          time: num(safeLeg.time, 0),
          dist: num(safeLeg.dist, 0),
          groundTime: num(safeLeg.groundTime, 0.5),
          fuel: num(safeLeg.fuel, 0),
          takeoffFuel: num(safeLeg.takeoffFuel, 0),
          landingFuel: num(safeLeg.landingFuel, 0),
          payload: num(safeLeg.payload, 0),
          availPayload: num(safeLeg.availPayload, 0),
          limit: num(safeLeg.limit, 0),
          pax: Array.isArray(safeLeg.pax) ? safeLeg.pax : [],
          limitType: String(safeLeg.limitType || ''),
          isOver: !!safeLeg.isOver,
          missionTime: String(safeLeg.missionTime || '08:00'),
          runwayGap: (safeLeg.runwayGap && typeof safeLeg.runwayGap === 'object') ? safeLeg.runwayGap : null
        };
      });
      
      document.getElementById('mTitle').innerText = `${m.id}: ${m.meta.pilot} (${m.meta.acft})`;
      document.getElementById('btnApprove').disabled = false;
      
      // 1. CALCULATE TOTALS
      let totalFlight = 0, totalFuel = 0, totalGround = 0;
      normalizedLegs.forEach(leg => {
        totalFlight += parseFloat(leg.time) || 0;
        totalFuel += parseFloat(leg.fuel) || 0;
        totalGround += parseFloat(leg.groundTime) || 0.5;
      });
      const PREFLIGHT_DUTY = 1.0;
      const POSTFLIGHT_DUTY = 0.75;
      const totalDuty = totalFlight + totalGround + PREFLIGHT_DUTY + POSTFLIGHT_DUTY;

      // Smart Checks 
      let warnings = [];
      if (totalFlight > 8.0) warnings.push("Flight Time > 8h");
      if (totalDuty > 14.0) warnings.push("Duty Time > 14h");
      
      const authRaw = data.authorizedAirports || "";
      const authList = authRaw.split(',').map(s => s.trim()); 
      let unauthorizedLegs = false;
      normalizedLegs.forEach(leg => {
        const isAuth = authList.some(a => a.trim().toUpperCase() === leg.to.trim().toUpperCase());
        if (!isAuth) unauthorizedLegs = true;
      });
      const runwayGapLegs = normalizedLegs.filter(function(leg) {
        return !!(leg && leg.runwayGap && leg.runwayGap.needsAttention);
      });
      
      if (unauthorizedLegs && !m.meta.copilot) warnings.push("Dest Check Needed");
      if (runwayGapLegs.length) warnings.push("Runway Internal Data Alert");
      if (m.meta.notes) warnings.push(m.meta.notes);

      const isOk = warnings.length === 0;
      const statusColor = isOk ? '#4caf50' : '#d32f2f';
      const statusTitle = isOk ? "OK" : "WARN";
      const statusDesc = isOk ? "Ready for Release" : warnings.join("<br>");

      // 2. RENDER TOP SUMMARY
      const container = document.getElementById('detailPanel');
      container.innerHTML = `
        <div style="display:flex; gap:10px; margin-bottom:20px; flex-wrap:wrap;">
          <div class="card center" style="flex:1; padding:15px; margin:0; min-width:100px;">
            <h5 style="margin:0; color:#0b5394;">${totalDuty.toFixed(1)}</h5>
            <small style="color:#777; font-weight:bold;">Est. Duty</small>
          </div>
          <div class="card center" style="flex:1; padding:15px; margin:0; min-width:100px;">
            <h5 style="margin:0; color:#0b5394;">${totalFlight.toFixed(1)}</h5>
            <small style="color:#777; font-weight:bold;">Flight Time</small>
          </div>
          <div class="card center" style="flex:1; padding:15px; margin:0; min-width:100px;">
            <h5 style="margin:0; color:#0b5394;">${Math.round(totalFuel)} L</h5>
            <small style="color:#777; font-weight:bold;">Total Fuel</small>
          </div>
          <div class="card center" style="flex:1; padding:15px; margin:0; min-width:120px; border-bottom: 3px solid ${statusColor}">
            <h5 style="margin:0; color:${statusColor};">${statusTitle}</h5>
            <small style="color:${isOk ? '#777' : '#d32f2f'}; font-weight:bold; font-size:0.75rem; line-height:1.1em; display:block; margin-top:4px;">${statusDesc}</small>
          </div>
        </div>
        <h6 style="margin-bottom:15px; border-bottom:1px solid #ccc; padding-bottom:5px;">Mission Profile</h6>
      `;

      // 3. LEG CARDS
      normalizedLegs.forEach((leg, idx) => {
        const isAuthorized = authList.some(auth => auth.trim().toUpperCase() === leg.to.trim().toUpperCase());
        const scheduleBtn = `<button class="btn-small blue lighten-1 z-depth-0" style="font-size:0.7rem; font-weight:bold; margin-left:6px;" onclick="openSupervisorRunwayCheckModal_(${idx})">SCHEDULE RUNWAY CHECK</button>`;
        const authHtml = isAuthorized
          ? `<span class="new badge green" data-badge-caption="Checked Out âœ…">Pilot</span>${scheduleBtn}`
          : `<button class="btn-small red lighten-2 z-depth-0" style="font-size:0.7rem; font-weight:bold;" onclick="waiveCheck('${leg.to}')">WAIVE CHECK</button>${scheduleBtn}`;
        const runwayGap = leg.runwayGap || null;
        const missingItems = runwayGap && Array.isArray(runwayGap.missingItems) ? runwayGap.missingItems : [];
        const runwayAlertHtml = runwayGap && runwayGap.needsAttention ? `
          <div style="margin:8px 0 10px; padding:8px 10px; border:1px solid #ef9a9a; border-left:4px solid #c62828; background:#ffebee; border-radius:4px;">
            <div style="display:flex; justify-content:space-between; align-items:center; gap:8px; flex-wrap:wrap;">
              <div style="font-size:0.75rem; color:#b71c1c; font-weight:800; text-transform:uppercase; letter-spacing:0.03em;">Runway Internal Data Alert</div>
              <button class="btn-small red darken-2 z-depth-0" style="font-size:0.68rem; font-weight:800;" onclick="openSupervisorRunwayBriefing_(${idx})">OPEN BRIEFING CARD</button>
            </div>
            <div style="font-size:0.78rem; color:#5d4037; margin-top:4px;">${runwayGap.summary || 'No internal cutdown / obstacle / slope data found for this destination runway.'}</div>
            ${missingItems.length ? `<div style="margin-top:5px; font-size:0.74rem; color:#6d4c41;"><b>Missing:</b> ${missingItems.join(', ')}</div>` : ''}
          </div>
        ` : '';
        
        const routingHtml = leg.waypointsText ? `
        <div style="margin: 8px 0; padding: 6px 10px; background: #e3f2fd; border-left: 4px solid #2196f3; border-radius: 0 4px 4px 0;">
          <div style="font-size: 0.65rem; font-weight: bold; color: #1565c0; text-transform: uppercase; letter-spacing: 0.5px;">Planned Routing (VIA)</div>
          <div style="font-family: 'Consolas', 'Monaco', monospace; font-size: 0.85rem; color: #0d47a1; font-weight: bold;">
            ${leg.waypointsText.replace(/,/g, ' âž” ')}
          </div>
        </div>
      ` : '';

        // Manifest
        const manifestHtml = leg.pax.length > 0 ? leg.pax.map(p => {
            const isMale = p.gender === 'M';
            const genderTag = p.gender ? `<span class="pax-tag ${isMale ? 'tag-m' : 'tag-f'}">${p.gender}</span>` : '';
            const ageTag = p.category ? `<span class="pax-tag tag-age">${p.category}</span>` : '';
            const moneyStr = (p.chargedAmount || 0).toLocaleString('en-US', {minimumFractionDigits: 2});
            return `
              <div class="pax-line">
                <span>
                  <b>${p.name}</b> <small style="color:#777">(${p.fund || '-'})</small> ${genderTag} ${ageTag}
                  <div style="font-size:0.75rem; color:#0b5394; margin-top:2px;">Rate: ${p.chargeRate || '?'} (R$ ${moneyStr})</div>
                  ${p.description ? `<div style="font-size:0.75rem; color:#666; font-style:italic;">â†³ ${p.description}</div>` : ''}
                </span>
                <span style="font-size:0.8rem; text-align:right;">
                  <div>Cargo: <b>${p.cargo}kg</b></div>
                  <div style="color:#888;">+ ${p.weight}kg</div>
                </span>
              </div>
            `;
        }).join('') : '<div style="padding:10px; color:#ccc; font-style:italic;">No passengers loaded.</div>';

        const limitDisp = leg.limit ? Math.round(leg.limit) : '---';
        const remPayload = leg.availPayload ? Math.round(leg.availPayload - leg.payload) : '---';
        const limitTypeClass = leg.limitType && leg.limitType.includes('Airport') ? 'bg-perf' : 'bg-struct';

        container.innerHTML += `
          <div class="leg-card" style="${leg.isOver ? 'border-left-color: #d32f2f; background: #ffebee;' : ''}">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:5px;">
              <div>
                <span style="font-weight:bold; font-size:1.1rem; color:#333;">LEG ${idx+1}: ${leg.from} âžœ ${leg.to}</span> 
                <br><small style="color:#777;">ID: ${m.id}-${idx+1}</small>
              </div>
              <div>${authHtml}</div>
            </div>
            
            ${routingHtml}

            <div style="font-size: 0.85rem; color: #333; margin-bottom: 8px; border-bottom: 1px solid #ddd; padding-bottom: 6px; line-height: 1.5;">
              <b>Dist:</b> ${Math.round(leg.dist || 0)} nm | 
              <b>Time:</b> ${leg.time.toFixed(1)} hr | 
              <b>Ground:</b> ${leg.groundTime.toFixed(1)} hr
              <br>
              <span style="color: #0b5394;">
                <b>Takeoff:</b> ${Math.round(leg.takeoffFuel || 0)}L | 
                <b>Enroute:</b> âˆ’${Math.round(leg.fuel || 0)}L | 
                <b>Landing:</b> ${Math.round(leg.landingFuel || 0)}L
              </span>
            </div>

            ${runwayAlertHtml}
            
            <div style="font-size:0.75rem; margin-bottom:10px;">
              <span class="limit-tag ${limitTypeClass}">${leg.limitType || 'Limit'}: ${limitDisp}kg</span>
              <span style="margin-left:10px">Payload Rem: <b>${remPayload}kg</b></span>
            </div>
            
            <div style="background:#f9f9f9; border-radius:4px; border:1px solid #eee;">
              ${manifestHtml}
            </div>

            <div class="dispatch-link" onclick="jumpToDispatch('${m.id}')">
              EDIT MISSION IN DISPATCH TERMINAL
              <div style="font-size:0.7rem; color:#90a4ae;">(Alter Fuel, Ground Time, or Assign Instructor)</div>
            </div>

          </div>
        `;
      });

      // 4. DUTY CONTEXT (simplified): logged before flight + planned before/after
      const tl = document.getElementById('timelineContainer');
      const _parseLocalDate_ = function(v) {
        const s = String(v || '').trim();
        return /^\d{4}-\d{2}-\d{2}$/.test(s) ? new Date(s + 'T00:00:00') : new Date(s);
      };
      const missionDate = _parseLocalDate_(m && m.meta ? m.meta.date : '');
      const validTimeline = Array.isArray(t) ? t.filter(function(ev) {
        return ev && ev.date && !isNaN(_parseLocalDate_(ev.date).getTime());
      }) : [];

      const _summarizeDutyByDay_ = function(events) {
        const acc = {};
        events.forEach(function(ev) {
          const key = String(ev.date || '').slice(0, 10);
          if (!key) return;
          const duty = Number(ev.dutyHrs || 0);
          const flight = Number(ev.flightHrs || 0);
          if (!acc[key]) acc[key] = { date: key, dutyHrs: 0, flightHrs: 0 };
          if (isFinite(duty)) acc[key].dutyHrs += duty;
          if (isFinite(flight)) acc[key].flightHrs += flight;
        });
        return Object.keys(acc).map(function(k) { return acc[k]; }).filter(function(row) {
          return row.dutyHrs > 0;
        });
      };

      const _sortByDateAsc_ = function(a, b) {
        return _parseLocalDate_(a.date).getTime() - _parseLocalDate_(b.date).getTime();
      };

      const actualBefore = _summarizeDutyByDay_(validTimeline.filter(function(ev) {
        const evDate = _parseLocalDate_(ev.date);
        return String(ev.type || '').toUpperCase() === 'LOGGED' && evDate < missionDate;
      })).sort(_sortByDateAsc_);

      const plannedFuture = _summarizeDutyByDay_(validTimeline.filter(function(ev) {
        const evDate = _parseLocalDate_(ev.date);
        return String(ev.type || '').toUpperCase() === 'SCHEDULED' && evDate >= missionDate;
      })).sort(_sortByDateAsc_);

      const _renderDayRow_ = function(row, tone, modeLabel) {
        return ''
          + '<div style="border:1px solid #e3e7ec; border-left:4px solid ' + tone + '; border-radius:6px; padding:8px; background:#fff; margin-bottom:7px; display:flex; justify-content:space-between; align-items:center; gap:8px;">'
          + '<div>'
          + '<div style="font-size:0.74rem; color:#78909c; font-weight:700;">' + supFormatDateMonDayYear_(row.date, '--') + '</div>'
          + '<div style="font-size:0.72rem; color:#546e7a; text-transform:uppercase; letter-spacing:0.03em; font-weight:800;">' + _supervisorEscHtml_(modeLabel) + '</div>'
          + '<div style="font-size:0.76rem; color:#37474f; margin-top:2px;">Total duty time: ' + row.dutyHrs.toFixed(1) + ' hours</div>'
          + '<div style="font-size:0.76rem; color:#37474f;">Flight hours: ' + row.flightHrs.toFixed(1) + '</div>'
          + '</div>'
          + '<div style="font-size:1.02rem; font-weight:900; color:#263238; white-space:nowrap;">' + row.dutyHrs.toFixed(1) + 'h</div>'
          + '</div>';
      };

      const _renderSection_ = function(label, list, tone, modeLabel, emptyText) {
        const itemsHtml = list.length
          ? list.map(function(row) { return _renderDayRow_(row, tone, modeLabel); }).join('')
          : '<div style="font-size:0.78rem; color:#90a4ae; font-style:italic; padding:8px 6px;">' + _supervisorEscHtml_(emptyText) + '</div>';
        return ''
          + '<div style="margin-bottom:12px;">'
          + '<div style="font-size:0.72rem; font-weight:900; color:#607d8b; text-transform:uppercase; letter-spacing:0.05em; margin:0 0 6px 0;">' + _supervisorEscHtml_(label) + '</div>'
          + itemsHtml
          + '</div>';
      };

      tl.innerHTML = ''
        + _renderSection_('Past Days (Actual)', actualBefore, '#2e7d32', 'Actual', 'No actual duty-hour records before this flight.')
        + _renderSection_('Upcoming (Planned)', plannedFuture, '#1565c0', 'Planned', 'No planned duty hours from this flight onward.');
    } 

    function openSupervisorRunwayCheckModal_(legIdx) {
      if (!lastLoadedData || !lastLoadedData.mission || !Array.isArray(lastLoadedData.mission.legs)) return;
      const mission = lastLoadedData.mission;
      const leg = mission.legs[Number(legIdx)] || null;
      if (!leg) return;

      const unauthorizedTargets = supervisorGetUnauthorizedRunwayTargets_();
      const selectedRunway = String(leg.to || '').trim().toUpperCase();
      const addAllTargets = unauthorizedTargets.length > 1 && unauthorizedTargets.indexOf(selectedRunway) >= 0
        ? window.confirm(
            'Existem ' + unauthorizedTargets.length + ' destinos sem cheque (' + unauthorizedTargets.join(', ') + ').\n\n' +
            'OK = adicionar todos ao cheque agora\nCancelar = manter apenas ' + selectedRunway + ' e waivar os demais.'
          )
        : false;
      const chosenTargets = addAllTargets
        ? unauthorizedTargets
        : [selectedRunway].filter(Boolean);
      const autoWaiveTargets = addAllTargets
        ? []
        : unauthorizedTargets.filter(function(code) { return code !== selectedRunway; });

      const routeTokens = String(leg.waypoints || '').trim()
        ? String(leg.waypoints || '').split(',').map(function(s){ return String(s || '').trim().toUpperCase(); }).filter(Boolean)
        : [String(leg.from || '').trim().toUpperCase(), String(leg.to || '').trim().toUpperCase()].filter(Boolean);
      const routeText = addAllTargets ? supervisorBuildMissionRoute_() : routeTokens.join(', ');
      const student = String(mission.meta && mission.meta.pilot || mission.pilot || '').trim();
      const runwayLocation = chosenTargets.join('; ');

      supervisorRunwayCheckCtx = {
        missionId: String(mission.id || ''),
        legIdx: Number(legIdx),
        flightLegId: String(leg.flightLegId || ''),
        student: student,
        route: routeText,
        runwayLocation: runwayLocation,
        autoWaiveTargets: autoWaiveTargets,
        acft: String(mission.meta && mission.meta.acft || mission.acft || '').trim().toUpperCase()
      };

      const sel = document.getElementById('supervisor-runway-check-instructor');
      const studentEl = document.getElementById('supervisor-runway-check-student');
      const runwayEl = document.getElementById('supervisor-runway-check-location');
      const routeEl = document.getElementById('supervisor-runway-check-route');
      const nameEl = document.getElementById('supervisor-runway-check-lesson-name');
      const subEl = document.getElementById('supervisor-runway-check-sub');
      const modal = document.getElementById('supervisor-runway-check-modal');

      if (sel) {
        sel.innerHTML = '<option value="">Selecione...</option>' + (supInstructors || []).map(function(name) {
          return '<option value="' + String(name).replace(/"/g, '&quot;') + '">' + String(name) + '</option>';
        }).join('');
      }
      if (studentEl) studentEl.value = student;
      if (runwayEl) runwayEl.value = runwayLocation;
      if (routeEl) routeEl.value = routeText;
      if (nameEl) nameEl.value = 'Cheque de Pista ' + (chosenTargets.join('-') || selectedRunway) + ' - ' + String(mission.id || '');
      if (subEl) {
        const waiveText = autoWaiveTargets.length ? (' â€¢ Waive automÃ¡tico: ' + autoWaiveTargets.join(', ')) : '';
        subEl.textContent = 'MissÃ£o ' + String(mission.id || '') + ' â€¢ Leg ' + (Number(legIdx) + 1) + waiveText;
      }

      if (modal) modal.style.display = 'flex';
    }

    function supervisorGetUnauthorizedRunwayTargets_() {
      if (!lastLoadedData || !lastLoadedData.mission || !Array.isArray(lastLoadedData.mission.legs)) return [];
      const authList = String(lastLoadedData.authorizedAirports || '')
        .split(',')
        .map(function(s){ return String(s || '').trim().toUpperCase(); })
        .filter(Boolean);
      const uniq = [];
      lastLoadedData.mission.legs.forEach(function(leg) {
        const to = String((leg && leg.to) || '').trim().toUpperCase();
        if (!to) return;
        if (authList.indexOf(to) >= 0) return;
        if (uniq.indexOf(to) >= 0) return;
        uniq.push(to);
      });
      return uniq;
    }

    function openSupervisorRunwayBriefing_(legIdx) {
      if (!lastLoadedData || !lastLoadedData.mission || !Array.isArray(lastLoadedData.mission.legs)) return;
      const mission = lastLoadedData.mission;
      const leg = mission.legs[Number(legIdx)] || null;
      if (!leg) return;
      const gap = (leg.runwayGap && typeof leg.runwayGap === 'object') ? leg.runwayGap : {};
      const icao = String(gap.icao || leg.to || '').trim().toUpperCase();
      const rwyIdent = String(gap.rwyIdent || '').trim().toUpperCase();
      const forcedRunway = (gap.runway && typeof gap.runway === 'object') ? gap.runway : { icao: icao, rwyIdent: rwyIdent };
      if (typeof openRunwayBriefingCard !== 'function' && typeof window.__rbcOpenRef === 'function') {
        window.openRunwayBriefingCard = window.__rbcOpenRef;
      }
      if (typeof openRunwayBriefingCard !== 'function') {
        const diag = (typeof window.runwayBriefingDeepProbe === 'function')
          ? window.runwayBriefingDeepProbe()
          : ((typeof window.runwayBriefingDiag === 'function') ? window.runwayBriefingDiag() : {
              hasOpenFn: false,
              hasOverlay: !!document.getElementById('rbc-overlay'),
              hasState: !!window._rbcState,
              activeView: (document.querySelector('.view-section.view-active') || {}).id || '',
              href: window.location.href
            });
        try { console.error('[SUPERVISOR][RBC] component missing', diag); } catch (e) {}
        if (window.M) M.toast({ html: 'Runway briefing component is not loaded.', classes: 'red' });
        return;
      }
      openRunwayBriefingCard(icao, rwyIdent, 'supervisor', forcedRunway);
    }

    window.runwayBriefingDiag = function() {
      const overlay = document.getElementById('rbc-overlay');
      const active = document.querySelector('.view-section.view-active');
      const out = {
        hasOpenFn: typeof window.openRunwayBriefingCard === 'function',
        hasCloseFn: typeof window.rbcClose === 'function',
        hasOverlay: !!overlay,
        overlayHidden: overlay ? !!overlay.hidden : null,
        overlayDisplay: overlay ? String(overlay.style.display || '') : '',
        hasState: !!window._rbcState,
        activeView: active ? String(active.id || '') : '',
        hash: String(window.location.hash || ''),
        href: String(window.location.href || '')
      };
      try { console.table(out); } catch (e) { console.log(out); }
      return out;
    };

    window.runwayBriefingDeepProbe = function() {
      const overlay = document.getElementById('rbc-overlay');
      const active = document.querySelector('.view-section.view-active');
      const scripts = Array.prototype.slice.call(document.querySelectorAll('script'));
      const hasInlineRbcSignature = scripts.some(function(s) {
        const txt = String(s && s.textContent || '');
        return txt.indexOf('RUNWAY BRIEFING CARD') >= 0 || txt.indexOf('__rbcScriptVersion') >= 0;
      });
      const out = {
        activeView: active ? String(active.id || '') : '',
        hasOpenFn: typeof window.openRunwayBriefingCard === 'function',
        hasBackupOpenFn: typeof window.__rbcOpenRef === 'function',
        hasCloseFn: typeof window.rbcClose === 'function',
        hasRbcState: !!window._rbcState,
        hasOverlay: !!overlay,
        overlayHidden: overlay ? !!overlay.hidden : null,
        overlayDisplay: overlay ? String(overlay.style.display || '') : '',
        hasInlineRbcSignature: hasInlineRbcSignature,
        rbcLoadedAt: String(window.__rbcLoadedAt || ''),
        rbcScriptVersion: String(window.__rbcScriptVersion || ''),
        href: String(window.location.href || '')
      };
      try { console.group('[SUPERVISOR][RBC][DEEP-PROBE]'); console.table(out); console.groupEnd(); } catch (e) { console.log(out); }
      return out;
    };

    function supervisorBuildMissionRoute_() {
      if (!lastLoadedData || !lastLoadedData.mission || !Array.isArray(lastLoadedData.mission.legs)) return '';
      const legs = lastLoadedData.mission.legs;
      const route = [];
      legs.forEach(function(leg, idx) {
        const from = String((leg && leg.from) || '').trim().toUpperCase();
        const to = String((leg && leg.to) || '').trim().toUpperCase();
        if (idx === 0 && from) route.push(from);
        if (to) route.push(to);
      });
      return route.filter(Boolean).join(', ');
    }

    function closeSupervisorRunwayCheckModal_() {
      const modal = document.getElementById('supervisor-runway-check-modal');
      if (modal) modal.style.display = 'none';
      supervisorRunwayCheckCtx = null;
    }

    async function submitSupervisorRunwayCheckSchedule_() {
      if (!supervisorRunwayCheckCtx) return;

      const instructorName = String((document.getElementById('supervisor-runway-check-instructor') || {}).value || '').trim();
      const runwayLocation = String((document.getElementById('supervisor-runway-check-location') || {}).value || '').trim().toUpperCase();
      const route = String((document.getElementById('supervisor-runway-check-route') || {}).value || '').trim().toUpperCase();
      const lessonName = String((document.getElementById('supervisor-runway-check-lesson-name') || {}).value || '').trim();

      if (!instructorName) {
        if (window.M) M.toast({ html: 'Selecione o instrutor.', classes: 'orange' });
        return;
      }
      if (!runwayLocation || !route) {
        if (window.M) M.toast({ html: 'Informe pista/local e rota.', classes: 'orange' });
        return;
      }

      const approvalPassword = await supervisorRequirePassword_('Informe a senha para agendar o cheque de pista.');
      if (!approvalPassword) return;

      const request = {
        missionId: supervisorRunwayCheckCtx.missionId,
        flightLegId: supervisorRunwayCheckCtx.flightLegId,
        instructorName: instructorName,
        runwayLocation: runwayLocation,
        route: route,
        lessonName: lessonName
      };

      const finishSchedule = function() {
        google.script.run
          .withSuccessHandler(function(resp) {
            if (!resp || !resp.success) {
              if (window.M) M.toast({ html: (resp && resp.error) ? resp.error : 'Falha ao agendar cheque.', classes: 'red' });
              return;
            }
            if (window.M) M.toast({ html: 'Cheque agendado: ' + String(resp.trainingCode || '') + ' â€¢ Agora aprove a missÃ£o.', classes: 'green' });
            closeSupervisorRunwayCheckModal_();
            refreshSupervisorDashboard(true);
            if (supMissionId) {
              google.script.run.withSuccessHandler(function(data) {
                lastLoadedData = data;
                renderDetails(data);
              }).withFailureHandler(showSupervisorError).getMissionDetailsForSupervisor(supMissionId);
            }
          })
          .withFailureHandler(function(err) {
            if (window.M) M.toast({ html: 'Erro: ' + (err && err.message ? err.message : String(err)), classes: 'red' });
          })
          .scheduleRunwayCheckFromSupervisor(request, approvalPassword);
      };

      const autoWaive = Array.isArray(supervisorRunwayCheckCtx.autoWaiveTargets)
        ? supervisorRunwayCheckCtx.autoWaiveTargets.filter(Boolean)
        : [];
      if (!autoWaive.length) {
        finishSchedule();
        return;
      }

      google.script.run
        .withSuccessHandler(function() {
          finishSchedule();
        })
        .withFailureHandler(function(err) {
          if (window.M) M.toast({ html: 'Falha ao waivar destinos extras: ' + (err && err.message ? err.message : String(err)), classes: 'red' });
        })
        .waiveDestinationChecksBatch(supPilot, autoWaive, supervisorRunwayCheckCtx.missionId, approvalPassword);
    }


    async function waiveCheck(icao) {
      if(confirm(`Authorize ${icao} for ${supPilot}?`)) {
         const approvalPassword = await supervisorRequirePassword_('Enter supervisor password to authorize ' + icao + '.');
         if (!approvalPassword) return;
         const currentList = lastLoadedData.authorizedAirports ? lastLoadedData.authorizedAirports.split(',').map(s=>s.trim()) : [];
         currentList.push(icao);
         lastLoadedData.authorizedAirports = currentList.join(', ');
         renderDetails(lastLoadedData);
         google.script.run
           .withSuccessHandler(msg => M.toast({html: 'Saved to Database'}))
           .withFailureHandler(err => M.toast({html: 'Authorization failed: ' + (err && err.message ? err.message : String(err)), classes:'red'}))
           .waiveDestinationCheck(supPilot, icao, supMissionId, approvalPassword);
      }
    }
    async function approveCurrent() {
      if(!confirm("Approve and Release Mission " + supMissionId + "?")) return;
      const approvingMissionId = String(supMissionId || '');
      const approvalPassword = await supervisorRequirePassword_('Enter supervisor password to approve mission ' + supMissionId + '.');
      if (!approvalPassword) return;
      
      // Disable the button immediately to prevent double-clicks
      document.getElementById('btnApprove').disabled = true;

      google.script.run.withSuccessHandler(() => { 
        // Optimistic local status update so sidebar reflects approval immediately.
        const approvedItem = document.querySelector('.mission-item[data-mission-id="' + approvingMissionId + '"]');
        if (approvedItem) {
          const badge = approvedItem.querySelector('.m-status');
          if (badge) {
            badge.classList.remove('st-pending');
            badge.classList.add('st-approved');
            badge.textContent = 'APPROVED';
          }
        }

        M.toast({html: 'Mission ' + supMissionId + ' Approved!'});

        // Notify portal views (same tab + other tabs) that mission status changed.
        try {
          if (typeof window.notifyMissionSaved === 'function') window.notifyMissionSaved();
          window.dispatchEvent(new CustomEvent('mission-saved', { detail: { when: Date.now(), source: 'supervisor-approve' } }));
          localStorage.setItem('mba_mission_saved_at', String(Date.now()));
        } catch (e) {}
        
        // 1. Reset the UI details to "Select a Mission" state
        document.getElementById('mTitle').innerText = "Select a Mission";
        document.getElementById('detailPanel').innerHTML = '<p style="text-align:center; margin-top:50px; color:#999;">Select a mission from the left.</p>';
        document.getElementById('timelineContainer').innerHTML = "";
        
        // 2. Refresh ONLY the sidebar list (No white screen reload!)
        google.script.run.withSuccessHandler(initSidebar).getSupervisorDashboard();
        
        // 3. Clear our local tracker
        supMissionId = null;

      }).withFailureHandler(err => {
        M.toast({html: 'Mission approval failed: ' + (err && err.message ? err.message : String(err)), classes:'red'});
        document.getElementById('btnApprove').disabled = false;
      }).approveMission(supMissionId, approvalPassword);
    }

    function openRunwaySurveyReview() {
      const el = document.getElementById('runway-survey-review-modal');
      if (!el || !window.M || !M.Modal) {
        if (window.M) M.toast({ html: 'Modal not available', classes: 'red' });
        return;
      }
      const modal = M.Modal.getInstance(el) || M.Modal.init(el);
      modal.open();
      refreshRunwaySurveyReviewList();
    }

    function refreshRunwaySurveyReviewList() {
      const list = document.getElementById('runway-survey-review-list');
      if (!list) return;
      list.innerHTML = '<div style="padding:12px; color:#888;">Loading pending runway surveysâ€¦</div>';
      _setRunwaySurveyReviewDiag_('request:start', {
        note: 'Calling getPendingRunwaySurveyReviews',
        limit: 25,
        pageHref: String(window.location.href || '')
      });

      google.script.run
        .withSuccessHandler(function(resp) {
          _setRunwaySurveyReviewDiag_('response:success-handler', resp);
          if (!resp || resp.success !== true) {
            list.innerHTML = '<div style="padding:12px; color:#d32f2f;">Failed loading runway surveys: ' + ((resp && resp.error) ? resp.error : 'Unknown response') + '</div>';
            return;
          }

          const items = Array.isArray(resp.items) ? resp.items : [];
          if (!items.length) {
            list.innerHTML = '<div style="padding:12px; color:#2e7d32; font-weight:700;">No pending runway surveys.</div>';
            return;
          }

          list.innerHTML = items.map(function(item) {
            const survey = item.survey || {};
            const official = item.official || {};
            const dbAirport = item.dbAirport || {};
            const summary = item.captureSummary || {};
            const isApproval = String(item.entryKind || '') === 'RUNWAY_APPROVAL';
            const dbKnownRaw = String(dbAirport.KNOWN_FEATURES || dbAirport.FEATURES || '').trim();
            let dbKnown = {};
            try { dbKnown = dbKnownRaw ? JSON.parse(dbKnownRaw) : {}; } catch (e) { dbKnown = {}; }
            if (Array.isArray(dbKnown)) dbKnown = { features: dbKnown };
            const dbVerified = (dbKnown && dbKnown.verifiedOperational && typeof dbKnown.verifiedOperational === 'object') ? dbKnown.verifiedOperational : {};
            const dbCurrentVersion = (dbKnown && dbKnown.currentSurveyVersion && typeof dbKnown.currentSurveyVersion === 'object') ? dbKnown.currentSurveyVersion : {};
            const dbVerifiedSurvey = (dbKnown && dbKnown.verifiedSurvey && typeof dbKnown.verifiedSurvey === 'object') ? dbKnown.verifiedSurvey : {};
            const internalStamp = String(dbCurrentVersion.publishedAt || dbVerifiedSurvey.capturedAt || dbKnown.updatedAt || '').trim();
            const internalDateLabel = internalStamp ? internalStamp.slice(0, 16).replace('T', ' ') + 'Z' : 'N/A';
            const previewPayload = encodeURIComponent(JSON.stringify({
              stagingId: item.stagingId,
              entryKind: item.entryKind,
              icao: item.icao,
              rwyIdent: item.rwyIdent,
              survey: survey,
              official: official,
              dbAirport: dbAirport,
              officialLogId: item.officialLogId,
              officialSourceLocked: !!item.officialSourceLocked
            }));
            const traceCount = survey._perimeterTraceCount != null ? survey._perimeterTraceCount : (Array.isArray(survey.perimeterTrace) ? survey.perimeterTrace.length : 0);
            const obsAngles = Array.isArray(survey.obstacleAngles50m) ? survey.obstacleAngles50m.length : 0;
            const featCount = Array.isArray(survey.features) ? survey.features.length : 0;
            const markerCount = Array.isArray(survey.markers) ? survey.markers.length : 0;
            const slopeCount = Array.isArray(survey.slopeSegments) ? survey.slopeSegments.length : 0;
            const runwaySnapshot = survey.runwaySnapshot || {};
            const riskDetail = survey.riskDetail || {};
            const mtowDefault = Number(runwaySnapshot.supervisorMtowKg || runwaySnapshot.maxTakeoffWeight || official.maxTakeoffWeight || 0);
            const mtowModelDefault = String(runwaySnapshot.mtowModelKey || official.mtowModelKey || '').trim();
            return `
              <div data-survey-staging="${item.stagingId}" data-entry-kind="${item.entryKind || ''}" data-preview-payload="${previewPayload}" style="background:#fff; border:1px solid #d9e2ef; border-radius:8px; padding:10px; margin-bottom:10px;">
                <div style="display:flex; justify-content:space-between; gap:10px; flex-wrap:wrap;">
                  <div>
                    <div style="font-weight:800; color:#0b5394;">${item.icao} RWY ${item.rwyIdent} ${isApproval ? 'Â· RUNWAY APPROVAL' : 'Â· GPS SURVEY'}</div>
                    <div style="font-size:0.8rem; color:#666;">Pilot: ${item.pilotName || 'Unknown'} Â· ${item.walkDate || ''}</div>
                    ${isApproval ? `
                      <div style="font-size:0.8rem; color:#37474f; margin-top:4px;"><b>Request Type</b> Pilot cutdown edit</div>
                      <div style="font-size:0.8rem; color:#37474f;"><b>Cutdown change</b> ${runwaySnapshot.cutdownAreaLabel || (runwaySnapshot.cutdownAreaM != null ? (String(runwaySnapshot.cutdownAreaM) + ' m') : 'Unknown')} (new) vs ${(runwaySnapshot.cutdownBaselineM != null ? (String(runwaySnapshot.cutdownBaselineM) + ' m') : (dbVerified.cutdownAreaLabel || (dbVerified.cutdownAreaM != null ? (String(dbVerified.cutdownAreaM) + ' m') : 'Unknown')))} (current)</div>
                      <div style="font-size:0.78rem; color:#777; margin-top:2px;">Internal baseline date: ${internalDateLabel} Â· Source: ${runwaySnapshot.editSource || 'TAB5_RELEASE'}</div>
                    ` : `
                      <div style="font-size:0.8rem; color:#666; margin-top:4px;">Trace pts: ${traceCount} Â· Features: ${featCount} Â· Markers: ${markerCount} Â· 50m obstacle angles: ${obsAngles} Â· Slope segs: ${slopeCount}</div>
                      <div style="font-size:0.8rem; color:#37474f; margin-top:4px;"><b>Observed</b> ${Number(survey.lengthM || 0) || 0}m x ${Number(survey.widthM || 0) || 0}m (${survey.surface || '-'})</div>
                      <div style="font-size:0.8rem; color:#37474f;"><b>Official</b> ${Number(official.lengthM || 0) || 0}m x ${Number(official.widthM || 0) || 0}m (${official.surface || '-'})</div>
                      <div style="font-size:0.78rem; color:#777; margin-top:2px;">Avg GPS acc: ${Math.round(Number(summary.avgAccuracyM || 0)) || 0}m Â· Best: ${Math.round(Number(summary.bestAccuracyM || 0)) || 0}m</div>
                    `}
                  </div>
                  <div style="display:flex; flex-direction:column; gap:6px; min-width:200px;">
                    <textarea id="survey-note-${item.stagingId}" placeholder="Supervisor notes" style="width:100%; min-height:58px; border:1px solid #ccc; border-radius:6px; padding:6px; font-size:0.82rem;"></textarea>
                    ${isApproval ? `
                      <div style="display:flex; gap:6px; align-items:center; flex-wrap:wrap;">
                        <input id="survey-mtow-model-${item.stagingId}" type="text" value="${mtowModelDefault}" placeholder="MTOW Model Key" style="flex:1 1 140px; height:32px; border:1px solid #ccc; border-radius:6px; padding:0 8px; font-size:0.78rem;">
                        <input id="survey-mtow-${item.stagingId}" type="number" inputmode="numeric" min="1" step="1" value="${mtowDefault > 0 ? mtowDefault : ''}" placeholder="MTOW kg" style="width:110px; height:32px; border:1px solid #ccc; border-radius:6px; padding:0 8px; font-size:0.78rem;">
                        <input id="survey-cutdown-${item.stagingId}" type="number" inputmode="numeric" min="1" step="1" value="" placeholder="Cutdown m" style="width:110px; height:32px; border:1px solid #ccc; border-radius:6px; padding:0 8px; font-size:0.78rem;">
                      </div>
                    ` : ''}
                    <div style="display:flex; gap:6px; align-items:center; flex-wrap:wrap;">
                      <label style="font-size:0.78rem; font-weight:700; color:#546e7a; white-space:nowrap;">Runway Class</label>
                      <select id="survey-rwyclass-${item.stagingId}" style="height:32px; border:1px solid #ccc; border-radius:6px; padding:0 8px; font-size:0.78rem; background:#fff;">
                        <option value="">-- assign --</option>
                        <option value="1">Class 1 (â‰¥900 m)</option>
                        <option value="2">Class 2 (600â€“899 m)</option>
                        <option value="3">Class 3 (&lt;600 m)</option>
                      </select>
                    </div>
                    <div style="display:flex; gap:6px; justify-content:flex-end;">
                      <button class="btn-small teal darken-2" onclick="openSupervisorRunwayPreview('${item.stagingId}')">Preview</button>
                      ${!isApproval ? `<button class="btn-small blue darken-1" onclick="openPendingVsActiveComparison('${item.stagingId}')">Compare Active</button>` : ''}
                      <button class="btn-small red lighten-1" onclick="rejectRunwaySurveyItem('${item.stagingId}')">Reject</button>
                      <button class="btn-small green" onclick="approveRunwaySurveyItem('${item.stagingId}')">Approve</button>
                    </div>
                  </div>
                </div>
                ${item.notes ? `<div style="margin-top:8px; font-size:0.82rem; color:#444; background:#f7f9fc; border-left:3px solid #90caf9; padding:6px 8px;">${item.notes}</div>` : ''}
              </div>
            `;
          }).join('');
        })
        .withFailureHandler(function(err) {
          _setRunwaySurveyReviewDiag_('response:failure-handler', {
            message: err && err.message ? err.message : String(err),
            raw: err
          });
          list.innerHTML = `<div style="padding:12px; color:#d32f2f;">Failed loading runway surveys: ${err && err.message ? err.message : String(err)}</div>`;
        })
        .getPendingRunwaySurveyReviews(25);
    }

    function _setRunwaySurveyReviewDiag_(stage, payload) {
      const el = document.getElementById('runway-survey-review-diag');
      if (!el) return;
      let safePayload = payload;
      try {
        safePayload = JSON.parse(JSON.stringify(payload));
      } catch (e) {
        safePayload = { nonSerializablePayload: String(payload) };
      }
      const envelope = {
        at: new Date().toISOString(),
        stage: String(stage || ''),
        payloadType: Array.isArray(payload) ? 'array' : typeof payload,
        payloadKeys: (payload && typeof payload === 'object' && !Array.isArray(payload)) ? Object.keys(payload) : [],
        payload: safePayload
      };
      try {
        el.textContent = JSON.stringify(envelope, null, 2);
      } catch (jsonErr) {
        el.textContent = 'Diagnostic JSON stringify failed: ' + (jsonErr && jsonErr.message ? jsonErr.message : String(jsonErr));
      }
    }

    function _getPendingSurveyCardData_() {
      const list = document.getElementById('runway-survey-review-list');
      if (!list) return [];
      return Array.from(list.querySelectorAll('[data-survey-staging]'));
    }

    window.openSupervisorRunwayPreview = function(stagingId) {
      const list = document.getElementById('runway-survey-review-list');
      if (!list || !stagingId) return;
      const card = list.querySelector('[data-survey-staging="' + String(stagingId) + '"]');
      const payload = card ? decodeURIComponent(card.dataset.previewPayload || '') : '';
      if (!payload) {
        if (window.M) M.toast({ html: 'Preview data unavailable', classes: 'orange' });
        return;
      }
      let item = null;
      try { item = JSON.parse(payload); } catch (e) { item = null; }
      if (!item) {
        if (window.M) M.toast({ html: 'Failed to parse preview data', classes: 'red' });
        return;
      }
      renderSupervisorRunwayPreview(item);
    };

    function closePendingVsActiveComparison() {
      const modal = document.getElementById('runway-pending-active-compare-modal');
      if (modal) modal.style.display = 'none';
    }

    function _pendingCompareFormat_(value) {
      if (value == null || value === '') return 'â€”';
      if (typeof value === 'number' && isFinite(value)) return String(Math.round(value * 100) / 100);
      return String(value);
    }

    function _pendingCompareArrayCount_(value) {
      return Array.isArray(value) ? value.length : 0;
    }

    function _pendingCompareRowHtml_(label, pendingValue, activeValue) {
      const pendingDisplay = _pendingCompareFormat_(pendingValue);
      const activeDisplay = _pendingCompareFormat_(activeValue);
      const changed = pendingDisplay !== activeDisplay;
      const pendingStyle = changed ? 'background:#fff8e1; color:#5d4037; font-weight:700;' : 'color:#263238;';
      const activeStyle = changed ? 'background:#e8f5e9; color:#1b5e20; font-weight:700;' : 'color:#263238;';
      return ''
        + '<tr style="border-bottom:1px solid #eef3f7;">'
        + '<td style="padding:6px 8px; color:#546e7a; font-weight:700; white-space:nowrap;">' + _supervisorEscHtml_(label) + '</td>'
        + '<td style="padding:6px 8px; ' + pendingStyle + '">' + _supervisorEscHtml_(pendingDisplay) + '</td>'
        + '<td style="padding:6px 8px; ' + activeStyle + '">' + _supervisorEscHtml_(activeDisplay) + '</td>'
        + '</tr>';
    }

    window.openPendingVsActiveComparison = function(stagingId) {
      const list = document.getElementById('runway-survey-review-list');
      const modal = document.getElementById('runway-pending-active-compare-modal');
      const title = document.getElementById('runway-pending-active-compare-title');
      const sub = document.getElementById('runway-pending-active-compare-sub');
      const body = document.getElementById('runway-pending-active-compare-body');
      if (!list || !modal || !title || !sub || !body || !stagingId) return;

      const card = list.querySelector('[data-survey-staging="' + String(stagingId) + '"]');
      const payload = card ? decodeURIComponent(card.dataset.previewPayload || '') : '';
      if (!payload) {
        if (window.M) M.toast({ html: 'Comparison data unavailable', classes: 'orange' });
        return;
      }

      let item = null;
      try { item = JSON.parse(payload); } catch (e) { item = null; }
      if (!item) {
        if (window.M) M.toast({ html: 'Failed to parse comparison data', classes: 'red' });
        return;
      }
      if (String(item.entryKind || '').toUpperCase() !== 'GPS_SURVEY') {
        if (window.M) M.toast({ html: 'Compare Active is only available for GPS survey items', classes: 'orange' });
        return;
      }

      title.textContent = 'PENDING VS ACTIVE';
      sub.textContent = String(item.icao || '--') + ' Â· RWY ' + String(item.rwyIdent || '--') + ' Â· Staging ' + String(item.stagingId || '');
      body.innerHTML = '<div style="padding:12px; color:#888;">Loading active approved versionâ€¦</div>';
      modal.style.display = 'flex';

      google.script.run
        .withSuccessHandler(function(resp) {
          if (!resp || !resp.success) {
            body.innerHTML = '<div style="padding:12px; color:#d32f2f;">Failed to load active survey: ' + _supervisorEscHtml_((resp && resp.error) || 'Unknown error') + '</div>';
            return;
          }

          const pendingSurvey = item.survey || {};
          const activeOperational = resp.activeVerifiedOperational || {};
          const activeSurvey = resp.activeVerifiedSurvey || {};
          const currentVersion = resp.currentVersion || {};
          const hasActive = Object.keys(activeOperational).length > 0 || Object.keys(activeSurvey).length > 0;

          if (!hasActive) {
            body.innerHTML = ''
              + '<div style="padding:12px; border-left:4px solid #ffb300; background:#fff8e1; color:#5d4037; border-radius:6px;">'
              + 'No active approved version exists yet for this runway. Approve this pending survey to create the first baseline.'
              + '</div>'
              + _supervisorRenderDataTable_('Pending Survey (new submission)', pendingSurvey, '#1565c0');
            return;
          }

          const pendingObstacleAngles = _pendingCompareArrayCount_(pendingSurvey.obstacleAngles50m);
          const activeObstacleAngles = _pendingCompareArrayCount_(activeOperational.obstacleAngles50m);
          const pendingCaptured = String(item.walkDate || '').trim();
          const activeCaptured = String(activeSurvey.capturedAt || '').trim();

          const rows = [];
          rows.push(_pendingCompareRowHtml_('Captured At', pendingCaptured || 'â€”', activeCaptured || 'â€”'));
          rows.push(_pendingCompareRowHtml_('Pilot', item.pilotName || 'â€”', activeSurvey.pilotName || 'â€”'));
          rows.push(_pendingCompareRowHtml_('Length (m)', Number(pendingSurvey.lengthM || 0) || 0, Number(activeOperational.lengthM || 0) || 0));
          rows.push(_pendingCompareRowHtml_('Width (m)', Number(pendingSurvey.widthM || 0) || 0, Number(activeOperational.widthM || 0) || 0));
          rows.push(_pendingCompareRowHtml_('Surface', pendingSurvey.surface || 'â€”', activeOperational.surface || 'â€”'));
          rows.push(_pendingCompareRowHtml_('Features', _pendingCompareArrayCount_(pendingSurvey.features), _pendingCompareArrayCount_(activeOperational.features)));
          rows.push(_pendingCompareRowHtml_('Markers', _pendingCompareArrayCount_(pendingSurvey.markers), _pendingCompareArrayCount_(activeOperational.markers)));
          rows.push(_pendingCompareRowHtml_('Obstacle Angles (50m)', pendingObstacleAngles, activeObstacleAngles));
          rows.push(_pendingCompareRowHtml_('Slope Segments', _pendingCompareArrayCount_(pendingSurvey.slopeSegments), _pendingCompareArrayCount_(activeOperational.slopeSegments)));
          rows.push(_pendingCompareRowHtml_('Cutdown Area (m2)', Number(pendingSurvey.cutdownAreaM || 0) || 0, Number(activeOperational.cutdownAreaM || 0) || 0));
          rows.push(_pendingCompareRowHtml_('Notes', String(pendingSurvey.notes || '').trim() || 'â€”', String(activeOperational.pilotNotes || '').trim() || 'â€”'));

          body.innerHTML = ''
            + '<div style="display:flex; justify-content:space-between; flex-wrap:wrap; gap:8px; align-items:center; padding:8px 10px; border:1px solid #d9e2ef; border-radius:8px; background:#f7fbff;">'
            + '<div style="font-size:0.8rem; color:#455a64;">Active version ID: <b>' + _supervisorEscHtml_(String(currentVersion.versionId || '').slice(-28) || 'N/A') + '</b></div>'
            + '<div style="font-size:0.78rem; color:#2e7d32; font-weight:700;">Yellow/green rows indicate differences</div>'
            + '</div>'
            + '<div style="overflow:auto; border:1px solid #d9e2ef; border-radius:8px;">'
            + '<table style="width:100%; border-collapse:collapse; min-width:680px;">'
            + '<thead><tr>'
            + '<th style="text-align:left; padding:8px; background:#eceff1; color:#455a64; font-size:0.76rem;">Parameter</th>'
            + '<th style="text-align:left; padding:8px; background:#fff8e1; color:#6d4c41; font-size:0.76rem;">Pending Survey</th>'
            + '<th style="text-align:left; padding:8px; background:#e8f5e9; color:#1b5e20; font-size:0.76rem;">Active Approved</th>'
            + '</tr></thead>'
            + '<tbody>' + rows.join('') + '</tbody>'
            + '</table>'
            + '</div>'
            + '<div style="display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:10px;">'
            + _supervisorRenderDataTable_('Pending Survey Payload', pendingSurvey, '#1565c0')
            + _supervisorRenderDataTable_('Active Approved Payload', activeOperational, '#2e7d32')
            + '</div>';
        })
        .withFailureHandler(function(err) {
          body.innerHTML = '<div style="padding:12px; color:#d32f2f;">Failed to load comparison: ' + _supervisorEscHtml_(err && err.message ? err.message : String(err)) + '</div>';
        })
        .getRunwaySurveyHistory(item.icao, item.rwyIdent);
    };

    function _supervisorEscHtml_(value) {
      return String(value == null ? '' : value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    }

    function _supervisorFlattenObject_(value, prefix, out) {
      const keyPrefix = String(prefix || '');
      if (Array.isArray(value)) {
        if (!value.length) {
          out.push({ key: keyPrefix || '(array)', value: '[]' });
          return;
        }
        value.forEach(function(v, i) {
          const nextKey = keyPrefix ? (keyPrefix + '[' + i + ']') : ('[' + i + ']');
          _supervisorFlattenObject_(v, nextKey, out);
        });
        return;
      }
      if (value && typeof value === 'object') {
        const keys = Object.keys(value);
        if (!keys.length) {
          out.push({ key: keyPrefix || '(object)', value: '{}' });
          return;
        }
        keys.forEach(function(k) {
          const nextKey = keyPrefix ? (keyPrefix + '.' + k) : k;
          _supervisorFlattenObject_(value[k], nextKey, out);
        });
        return;
      }
      out.push({ key: keyPrefix || '(value)', value: value == null ? '' : value });
    }

    function _supervisorRenderDataTable_(titleText, sourceObj, accentColor) {
      const rows = [];
      _supervisorFlattenObject_(sourceObj || {}, '', rows);
      const tableRows = rows.map(function(r) {
        const displayVal = (typeof r.value === 'string' || typeof r.value === 'number' || typeof r.value === 'boolean')
          ? String(r.value)
          : JSON.stringify(r.value);
        return '<tr>'
          + '<td style="vertical-align:top; padding:4px 8px; border-bottom:1px solid #eef3f7; color:#455a64; font-family:monospace; font-size:0.76rem;">' + _supervisorEscHtml_(r.key) + '</td>'
          + '<td style="vertical-align:top; padding:4px 8px; border-bottom:1px solid #eef3f7; color:#263238; font-family:monospace; font-size:0.76rem;">' + _supervisorEscHtml_(displayVal) + '</td>'
          + '</tr>';
      }).join('');
      return ''
        + '<div style="margin-top:10px; border:1px solid #d9e2ef; border-radius:8px; overflow:hidden;">'
        + '<div style="padding:6px 8px; font-size:0.8rem; font-weight:800; color:#fff; background:' + (accentColor || '#0b5394') + ';">' + _supervisorEscHtml_(titleText) + '</div>'
        + '<div style="max-height:240px; overflow:auto; background:#fff;">'
        + '<table style="width:100%; border-collapse:collapse; table-layout:fixed;">'
        + '<thead><tr>'
        + '<th style="width:42%; text-align:left; padding:6px 8px; border-bottom:1px solid #d9e2ef; background:#f7fbff; font-size:0.74rem; color:#546e7a;">Parameter</th>'
        + '<th style="text-align:left; padding:6px 8px; border-bottom:1px solid #d9e2ef; background:#f7fbff; font-size:0.74rem; color:#546e7a;">Value</th>'
        + '</tr></thead>'
        + '<tbody>' + (tableRows || '<tr><td colspan="2" style="padding:8px; color:#777;">No parameters found.</td></tr>') + '</tbody>'
        + '</table>'
        + '</div>'
        + '</div>';
    }

    function renderSupervisorRunwayPreview(item) {
      const modal = document.getElementById('runway-survey-preview-modal');
      const title = document.getElementById('runway-survey-preview-title');
      const sub = document.getElementById('runway-survey-preview-sub');
      const body = document.getElementById('runway-survey-preview-body');
      if (!modal || !title || !sub || !body) return;

      const survey = item && item.survey ? item.survey : {};
      const official = item && item.official ? item.official : {};
      const dbAirport = item && item.dbAirport ? item.dbAirport : {};
      const dbKnownRaw = String(dbAirport.KNOWN_FEATURES || dbAirport.FEATURES || '').trim();
      let dbKnown = {};
      try { dbKnown = dbKnownRaw ? JSON.parse(dbKnownRaw) : {}; } catch (e) { dbKnown = {}; }
      if (Array.isArray(dbKnown)) dbKnown = { features: dbKnown };
      const dbVerified = (dbKnown && dbKnown.verifiedOperational && typeof dbKnown.verifiedOperational === 'object') ? dbKnown.verifiedOperational : {};
      const dbCurrentVersion = (dbKnown && dbKnown.currentSurveyVersion && typeof dbKnown.currentSurveyVersion === 'object') ? dbKnown.currentSurveyVersion : {};
      const dbVerifiedSurvey = (dbKnown && dbKnown.verifiedSurvey && typeof dbKnown.verifiedSurvey === 'object') ? dbKnown.verifiedSurvey : {};
      const internalStamp = String(dbCurrentVersion.publishedAt || dbVerifiedSurvey.capturedAt || dbKnown.updatedAt || '').trim();
      const internalDateLabel = internalStamp ? internalStamp.slice(0, 16).replace('T', ' ') + 'Z' : 'N/A';
      const officialLogId = String(item && item.officialLogId || official && official.officialLogId || item && item.stagingId || '').trim();
      const sourceLocked = !!(item && item.officialSourceLocked);
      
      // Check if this is a RUNWAY_APPROVAL entry kind
      const isRunwayApproval = String(item && item.entryKind || '').trim().toUpperCase() === 'RUNWAY_APPROVAL';
      
      const lengthM = Math.max(1, Math.round(Number(survey.lengthM || official.lengthM || 0)));
      const widthM = Math.max(0, Math.round(Number(survey.widthM || official.widthM || 0)));
      const surface = String(survey.surface || official.surface || '-');
      const features = Array.isArray(survey.features) ? survey.features : [];
      const obstacles = Array.isArray(survey.obstacleAngles50m) ? survey.obstacleAngles50m : [];
      const slopes = Array.isArray(survey.slopeSegments) ? survey.slopeSegments : [];

      title.textContent = 'RUNWAY PREVIEW';
      sub.textContent = String(item.icao || '--') + ' Â· RWY ' + String(item.rwyIdent || '--');

      // â”€â”€â”€ SPECIAL RENDERING FOR RUNWAY_APPROVAL â”€â”€â”€
      if (isRunwayApproval) {
        const briefingCard = (dbKnown && dbKnown.briefingCard && typeof dbKnown.briefingCard === 'object') 
          ? dbKnown.briefingCard 
          : {};
        
        // Extract survey and baseline cutdown values
        const surveyRunway = (survey && survey.runwaySnapshot) || survey;
        const surveyCutdown = Number(surveyRunway && surveyRunway.cutdownAreaM || 0);
        const cutdownBaseline = Number(surveyRunway && surveyRunway.cutdownBaselineM || 0);
        const currentDbCutdown = Number(dbVerified && dbVerified.cutdownAreaM || 0);
        
        // Briefing card parameters
        const briefCardLen = Math.round(Number(briefingCard.internalLengthM || lengthM || 0));
        const briefCardWid = Math.round(Number(briefingCard.internalWidthM || widthM || 0));
        const briefCardSurface = String(briefingCard.surface || surface || '-');
        const briefCardCutdown = Math.round(Number(briefingCard.cutdownAreaM || surveyCutdown || 0));
        const briefCardElev = Math.round(Number(briefingCard.elevation || 0));
        
        // Build briefing card display (simplified for supervisor)
        const briefingCardHtml = ''
          + '<div style="background:#f5f9ff; border:1px solid #0b5394; border-radius:10px; padding:16px; margin-bottom:16px;">'
          + '<div style="font-size:0.75rem; font-weight:900; color:#0b5394; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:10px;">Briefing Card (Current)</div>'
          + '<div style="display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:8px;">'
          + '<div style="background:#fff; border:1px solid #c5d9f1; border-radius:6px; padding:8px;">'
          +   '<div style="font-size:0.65rem; font-weight:900; color:#0b5394; text-transform:uppercase; margin-bottom:3px;">Length</div>'
          +   '<div style="font-size:1rem; font-weight:900; color:#1a1a1a;">' + briefCardLen + ' <span style="font-size:0.75rem; color:#666;">m</span></div>'
          + '</div>'
          + '<div style="background:#fff; border:1px solid #c5d9f1; border-radius:6px; padding:8px;">'
          +   '<div style="font-size:0.65rem; font-weight:900; color:#0b5394; text-transform:uppercase; margin-bottom:3px;">Width</div>'
          +   '<div style="font-size:1rem; font-weight:900; color:#1a1a1a;">' + briefCardWid + ' <span style="font-size:0.75rem; color:#666;">m</span></div>'
          + '</div>'
          + '<div style="background:#fff; border:1px solid #c5d9f1; border-radius:6px; padding:8px;">'
          +   '<div style="font-size:0.65rem; font-weight:900; color:#0b5394; text-transform:uppercase; margin-bottom:3px;">Surface</div>'
          +   '<div style="font-size:0.85rem; font-weight:700; color:#1a1a1a;">' + briefCardSurface + '</div>'
          + '</div>'
          + '<div style="background:#fff; border:1px solid #c5d9f1; border-radius:6px; padding:8px;">'
          +   '<div style="font-size:0.65rem; font-weight:900; color:#0b5394; text-transform:uppercase; margin-bottom:3px;">Cutdown</div>'
          +   '<div style="font-size:1rem; font-weight:900; color:#1a1a1a;">' + briefCardCutdown + ' <span style="font-size:0.75rem; color:#666;">m</span></div>'
          + '</div>'
          + '<div style="background:#fff; border:1px solid #c5d9f1; border-radius:6px; padding:8px;">'
          +   '<div style="font-size:0.65rem; font-weight:900; color:#0b5394; text-transform:uppercase; margin-bottom:3px;">Elevation</div>'
          +   '<div style="font-size:0.9rem; font-weight:700; color:#1a1a1a;">' + (briefCardElev || 'â€“') + ' <span style="font-size:0.75rem; color:#666;">ft</span></div>'
          + '</div>'
          + '<div style="background:#fff; border:1px solid #c5d9f1; border-radius:6px; padding:8px;">'
          +   '<div style="font-size:0.65rem; font-weight:900; color:#0b5394; text-transform:uppercase; margin-bottom:3px;">Class</div>'
          +   '<div style="font-size:0.9rem; font-weight:700; color:#1a1a1a;">' + (briefingCard.runwayClass ? 'Class ' + briefingCard.runwayClass : 'â€“') + '</div>'
          + '</div>'
          + '</div>'
          + '</div>';
        
        // Cutdown change comparison sidebar
        const cutdownChange = surveyCutdown - currentDbCutdown;
        const cutdownChangeColor = cutdownChange > 0 ? '#ef6c00' : (cutdownChange < 0 ? '#1976d2' : '#999');
        const cutdownChangeLabel = cutdownChange > 0 ? '+' + Math.round(cutdownChange) : Math.round(cutdownChange);
        
        const changeComparisonHtml = ''
          + '<div style="background:#fff9f0; border:1px solid #ffb74d; border-radius:10px; padding:12px; flex:0 0 200px;">'
          + '<div style="font-size:0.72rem; font-weight:900; color:#e65100; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:10px;">Pilot Changes</div>'
          + '<div style="background:#fff; border:1px solid #ffe0b2; border-radius:6px; padding:10px; margin-bottom:8px;">'
          +   '<div style="font-size:0.72rem; font-weight:800; color:#bf360c; text-transform:uppercase; margin-bottom:6px;">Cutdown Area</div>'
          +   '<div style="display:flex; align-items:baseline; justify-content:space-between; margin-bottom:6px;">'
          +     '<span style="font-size:0.8rem; color:#666;">Before:</span>'
          +     '<span style="font-size:0.95rem; font-weight:700; color:#1a1a1a;">' + Math.round(currentDbCutdown) + ' m</span>'
          +   '</div>'
          +   '<div style="display:flex; align-items:baseline; justify-content:space-between; margin-bottom:6px;">'
          +     '<span style="font-size:0.8rem; color:#666;">After:</span>'
          +     '<span style="font-size:0.95rem; font-weight:700; color:#1a1a1a;">' + Math.round(surveyCutdown) + ' m</span>'
          +   '</div>'
          +   '<div style="height:4px; background:#f0f0f0; border-radius:2px; margin:6px 0;"></div>'
          +   '<div style="display:flex; align-items:center; justify-content:space-between;">'
          +     '<span style="font-size:0.75rem; font-weight:900; color:' + cutdownChangeColor + ';">CHANGE</span>'
          +     '<span style="font-size:1.1rem; font-weight:900; color:' + cutdownChangeColor + ';">' + cutdownChangeLabel + ' m</span>'
          +   '</div>'
          + '</div>'
          + ((Number(dbVerified && dbVerified.internalUpdatedAt) > 0) || internalDateLabel !== 'N/A'
            ? '<div style="font-size:0.75rem; color:#999; margin-top:6px;">Last updated: ' + internalDateLabel + '</div>'
            : '<div style="background:#fff3cd; border:1px solid #ffeaa7; border-radius:4px; padding:6px 8px; font-size:0.75rem; color:#856404; margin-top:6px;">'
            +   '<strong>âš  No internal data</strong> â€” Pilot will survey'
            +   '<button onclick="(function(el){el.parentElement.style.display=\'none\'; if(window.M)M.toast({html:\'Pilot survey noted.\',classes:\'blue\'});})(this)" '
            +   'style="display:block; margin-top:4px; background:#856404; color:#fff; border:none; border-radius:3px; padding:3px 8px; font-size:0.7rem; font-weight:700; cursor:pointer; width:100%;">'
            +   'Pilot Will Survey'
            + '</button></div>')
          + '</div>';
        
        // Render RUNWAY_APPROVAL-specific view
        body.innerHTML = ''
          + briefingCardHtml
          + '<div style="display:flex; gap:12px; align-items:flex-start;">'
          + changeComparisonHtml
          + '<div style="flex:1; min-width:0;">'
          + '<div style="background:#f0f5ff; border:1px solid #d7dde3; border-radius:10px; padding:12px; margin-bottom:12px;">'
          + '<div style="font-size:0.72rem; font-weight:900; color:#0b5394; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:8px;">Runway Diagram</div>'
          + '<svg width="100%" height="180" viewBox="0 0 480 200" style="border:1px solid #ddd; border-radius:6px; background:#fff;">'
          + '<rect x="20" y="50" width="380" height="50" fill="#616161" stroke="#263238" stroke-width="2" rx="3"></rect>'
          + '<line x1="20" y1="75" x2="400" y2="75" stroke="#f3f3f3" stroke-width="2" stroke-dasharray="18,10"></line>'
          + '<text x="30" y="85" font-size="14" fill="#fff" font-weight="bold">' + String(item.rwyIdent || 'RWY') + '</text>'
          + '<text x="390" y="85" font-size="14" fill="#fff" font-weight="bold" text-anchor="end">' + (survey && survey.reciprocalIdent ? String(survey.reciprocalIdent) : '') + '</text>'
          + '<text x="20" y="30" font-size="11" fill="#666" font-weight="700">' + briefCardLen + ' m Ã— ' + briefCardWid + ' m</text>'
          + '</svg>'
          + '</div>'
          + _supervisorRenderDataTable_('Runway Snapshot (pilot survey)', surveyRunway, '#0b5394')
          + '</div>'
          + '</div>';

        modal.style.display = 'flex';
        return;
      }
      
      // â”€â”€â”€ STANDARD RENDERING FOR GPS_SURVEY â”€â”€â”€
      const officialTitle = sourceLocked
        ? ('&#128274; Official (DB_Airports Â· LOGID ' + _supervisorEscHtml_(officialLogId || 'N/A') + ')')
        : ('Official' + (officialLogId ? (' (LOGID ' + _supervisorEscHtml_(officialLogId) + ')') : ''));

      const topSummary = ''
        + '<div style="display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:8px;">'
        + '<div style="border:1px solid #dbe7f3; border-radius:8px; padding:8px; background:#f7fbff; font-size:0.84rem;"><b style="color:#0b5394;">Surveyed/Internal</b><br>' + lengthM + 'm Ã— ' + widthM + 'm â€¢ ' + surface + '<br><span style="font-size:0.76rem; color:#607d8b;">Current internal: ' + (Number(dbVerified.lengthM || 0) || 0) + 'm Ã— ' + (Number(dbVerified.widthM || 0) || 0) + 'm â€¢ ' + (dbVerified.surface || '-') + ' Â· ' + internalDateLabel + '</span></div>'
        + '<div style="border:1px solid #e0e0e0; border-radius:8px; padding:8px; background:#fafafa; font-size:0.84rem;"><b style="color:#555;">' + officialTitle + '</b><br>' + Math.round(Number(official.lengthM || 0)) + 'm Ã— ' + Math.round(Number(official.widthM || 0)) + 'm â€¢ ' + String(official.surface || '-') + '</div>'
        + '</div>';

      const runwayPx = 480;
      const scale = runwayPx / lengthM;
      const topX = 48;
      const topY = 52;
      const topH = 72;
      const sideX = topX;
      const sideY = topY + topH + 30;
      const sideH = 60;
      const viewW = 600;
      const viewH = sideY + sideH + 80;
      const endX = topX + runwayPx;

      const topFeatureDots = features.slice(0, 40).map(function(f) {
        const d = Math.max(0, Math.min(lengthM, Number(f && f.distance || 0)));
        const x = topX + (d * scale);
        const side = String(f && f.side || 'right').toLowerCase() === 'left' ? 'left' : 'right';
        const y = side === 'left' ? topY - 10 : topY + topH + 14;
        return '<circle cx="' + x.toFixed(1) + '" cy="' + y.toFixed(1) + '" r="3" fill="#fb8c00"></circle>';
      }).join('');

      // Slope profile from survey slopeSegments
      const sortedSlopes = slopes.map(function(s) {
        const sd = Math.max(0, Number(s && s.startDistanceM || 0));
        const dist = Math.max(0, Number(s && (s.distanceM != null ? s.distanceM : s.distance) || 0));
        return { startDistanceM: sd, distance: dist, slope: Number(s && s.slope || 0) };
      }).filter(function(s) { return s.distance > 0; }).sort(function(a, b) { return a.startDistanceM - b.startDistanceM; });
      const profPts = [];
      let cursor = 0, elev = 0;
      const elevArr = [0];
      const profileSegs = [];
      sortedSlopes.forEach(function(seg) {
        if (seg.startDistanceM > cursor) profileSegs.push({ startM: cursor, endM: seg.startDistanceM, slope: 0 });
        const endM = Math.min(lengthM, seg.startDistanceM + seg.distance);
        profileSegs.push({ startM: seg.startDistanceM, endM: endM, slope: seg.slope });
        cursor = endM;
      });
      if (cursor < lengthM) profileSegs.push({ startM: cursor, endM: lengthM, slope: 0 });
      const allElevs = [0];
      let ePos = 0;
      profileSegs.forEach(function(s) { ePos += ((s.endM - s.startM) * s.slope / 100); allElevs.push(ePos); });
      if (!profileSegs.length) { profileSegs.push({ startM: 0, endM: lengthM, slope: 0 }); allElevs.push(0); }
      const minEl = Math.min.apply(null, allElevs);
      const maxEl = Math.max.apply(null, allElevs);
      const spanEl = Math.max(0.5, maxEl - minEl);
      const elScale = Math.max(0.8, Math.min(sideH / spanEl, 4));
      const usedH = spanEl * elScale;
      const yOff = (sideH - usedH) / 2;
      let eIdx = 0;
      const sidePolyPts = [sideX + ',' + (sideY + yOff + usedH - ((allElevs[0] - minEl) * elScale))];
      profileSegs.forEach(function(s, i) {
        sidePolyPts.push((sideX + (s.endM * scale)).toFixed(1) + ',' + (sideY + yOff + usedH - ((allElevs[i + 1] - minEl) * elScale)).toFixed(1));
      });

      const yForDist = function(d) {
        const dm = Math.max(0, Math.min(lengthM, Number(d || 0)));
        let cur = 0, elStart = 0;
        for (let i = 0; i < profileSegs.length; i++) {
          const s = profileSegs[i];
          if (dm <= s.endM || i === profileSegs.length - 1) {
            const span = Math.max(0.0001, s.endM - s.startM);
            const t = Math.max(0, Math.min(1, (dm - s.startM) / span));
            const e = elStart + ((s.endM - s.startM) * s.slope / 100) * t;
            return sideY + yOff + usedH - ((e - minEl) * elScale);
          }
          elStart += ((s.endM - s.startM) * s.slope / 100);
        }
        return sideY + yOff + usedH;
      };

      const normThr = function(raw) {
        const txt = String(raw || '').trim().toUpperCase().replace(/^RWY\s*/, '');
        const m = txt.match(/^(\d{1,2})([LCR])?$/);
        if (!m) return txt;
        const n = parseInt(m[1], 10);
        if (!isFinite(n) || n < 1 || n > 36) return txt;
        return String(n).padStart(2, '0') + (m[2] || '');
      };
      const thisThrSup = normThr(item.rwyIdent || '');

      const iconForObs = function(type) {
        const t = String(type || '').toLowerCase();
        if (t.indexOf('tree') >= 0) return '\ud83c\udf33';
        if (t.indexOf('building') >= 0 || t.indexOf('house') >= 0) return '\ud83c\udfe2';
        if (t.indexOf('hill') >= 0) return '\u26f0';
        if (t.indexOf('rock') >= 0) return '\ud83e\udea8';
        if (t.indexOf('power') >= 0) return '\u26a1';
        return '\ud83d\udccd';
      };

      // Build obstacle SVG groups (all visible by default, user can toggle)
      const obsDropY = sideY + sideH + 2;
      const obsIconY = obsDropY + 20;
      const obsDistY = obsDropY + 33;
      const obsGroupsSvg = obstacles.map(function(obs, idx) {
        const dist = Math.max(0, Math.min(lengthM, Number(obs && obs.checkpointDistanceM || 0)));
        const thr = normThr(obs && obs.fromThreshold || '');
        const corner = String(obs && obs.checkpointCorner || '').trim().toUpperCase();
        const fromA = (thr && thr !== thisThrSup) || (!thr && corner === 'C') ? (lengthM - dist) : dist;
        const operation = String(obs && obs.operation || '').trim().toLowerCase() || (dist >= 300 ? 'takeoff' : 'landing');
        const bx = (sideX + (fromA * scale)).toFixed(1);
        const by = yForDist(fromA).toFixed(1);
        const icon = iconForObs(obs && obs.type);
        const angleDeg = Number(obs && (obs.angleDeg != null ? obs.angleDeg : (obs.fromThrA50mDeg != null ? obs.fromThrA50mDeg : obs.fromThrB50mDeg)) || 0).toFixed(1);
        const lineColor = operation === 'landing' ? '#1976d2' : '#ef6c00';
        return '<g id="supObsGrp_' + idx + '">'
          + '<line x1="' + bx + '" y1="' + by + '" x2="' + bx + '" y2="' + obsDropY + '" stroke="' + lineColor + '" stroke-width="3"></line>'
          + '<circle cx="' + bx + '" cy="' + by + '" r="6" fill="#37474f"></circle>'
          + '<text x="' + bx + '" y="' + obsIconY + '" font-size="16" fill="#2e7d32" text-anchor="middle">' + icon + ' ' + angleDeg + '\u00b0</text>'
          + '<text x="' + bx + '" y="' + obsDistY + '" font-size="10" fill="' + lineColor + '" text-anchor="middle">' + Math.round(dist) + 'm</text>'
          + '</g>';
      }).join('');

      const slopeLabels = profileSegs.map(function(s, i) {
        const midX = (sideX + ((s.startM + (s.endM - s.startM) / 2) * scale)).toFixed(1);
        const segPx = (s.endM - s.startM) * scale;
        if (segPx < 28) return '';
        return '<text x="' + midX + '" y="' + (sideY + sideH - 6) + '" font-size="9" fill="#607d8b" text-anchor="middle">'
          + (s.slope >= 0 ? '+' : '') + s.slope.toFixed(1) + '%</text>';
      }).join('');

      // Checkbox list label for each obstacle
      const obsCheckboxes = obstacles.map(function(obs, idx) {
        const dist = Math.round(Number(obs && obs.checkpointDistanceM || 0));
        const thr = normThr(obs && obs.fromThreshold || '') || normThr(item.rwyIdent || '');
        const operation = String(obs && obs.operation || '').trim().toLowerCase() || (dist >= 300 ? 'takeoff' : 'landing');
        const label = 'RWY ' + thr + ' Â· ' + dist + 'm Â· ' + operation;
        return '<label style="display:flex; align-items:center; gap:4px; font-size:0.77rem; color:#37474f; cursor:pointer; white-space:nowrap;">'
          + '<input type="checkbox" checked onchange="(function(el,id){var g=document.getElementById(id);if(g)g.style.display=el.checked?\'\':\'none\';})(this,\'supObsGrp_' + idx + '\')">'
          + label + '</label>';
      }).join('');

      body.innerHTML = ''
        + topSummary
        + '<div style="display:flex; gap:10px; align-items:flex-start;">'
        + '<svg width="100%" height="' + viewH + '" viewBox="0 0 ' + viewW + ' ' + viewH + '" style="border:1px solid #d7dde3; border-radius:10px; background:#f9fbfd; flex:1 1 auto;">'
        + '<text x="48" y="24" font-size="12" fill="#234" font-weight="700">TOP VIEW</text>'
        + '<rect x="' + topX + '" y="' + topY + '" width="' + runwayPx + '" height="' + topH + '" fill="#616161" stroke="#263238" stroke-width="2" rx="4" ry="4"></rect>'
        + '<line x1="' + topX + '" y1="' + (topY + topH / 2) + '" x2="' + endX + '" y2="' + (topY + topH / 2) + '" stroke="#f3f3f3" stroke-width="3" stroke-dasharray="18,10"></line>'
        + '<text x="' + (topX + 10) + '" y="' + (topY + topH - 10) + '" font-size="17" fill="#fff" font-weight="800">' + String(item.rwyIdent || 'RWY') + '</text>'
        + '<text x="' + (endX - 10) + '" y="' + (topY + topH - 10) + '" font-size="17" fill="#fff" font-weight="800" text-anchor="end">' + String(survey.reciprocalIdent || '') + '</text>'
        + topFeatureDots
        + '<text x="48" y="' + (sideY - 8) + '" font-size="12" fill="#234" font-weight="700">SIDE VIEW</text>'
        + '<rect x="' + sideX + '" y="' + sideY + '" width="' + runwayPx + '" height="' + sideH + '" fill="#f2f5f8" stroke="#d7dde3"></rect>'
        + '<polyline points="' + sidePolyPts.join(' ') + '" fill="none" stroke="#2e7d32" stroke-width="2.5"></polyline>'
        + slopeLabels
        + obsGroupsSvg
        + '</svg>'
        + (obstacles.length ? '<div style="border:1px solid #e0e6eb; border-radius:8px; padding:8px; background:#fafefe; min-width:160px; max-width:200px; flex-shrink:0;">'
          + '<div style="font-size:0.78rem; font-weight:800; color:#0b5394; margin-bottom:6px;">Obstacle Angles</div>'
          + '<div style="display:flex; flex-direction:column; gap:5px;">' + obsCheckboxes + '</div>'
          + '</div>' : '')
        + '</div>'
        + '<div style="background:#eef6ff; border-left:4px solid #1976d2; padding:8px 10px; border-radius:6px; font-size:0.82rem; color:#37474f;">Features: ' + features.length + ' \u2022 Obstacles: ' + obstacles.length + ' \u2022 Slope segments: ' + slopes.length + '</div>'
        + _supervisorRenderDataTable_('Measured Parameters (survey payload)', survey, '#0b5394')
        + _supervisorRenderDataTable_('DB_Airports Parameters (official source)', dbAirport, '#546e7a');

      modal.style.display = 'flex';
    }

    function closeSupervisorRunwayPreview() {
      const modal = document.getElementById('runway-survey-preview-modal');
      if (modal) modal.style.display = 'none';
    }

    function _supervisorNameForSurvey_() {
      const raw = document.getElementById('userDisplay') ? document.getElementById('userDisplay').innerText : 'Supervisor';
      return String(raw || '').trim() || 'Supervisor';
    }

    async function approveRunwaySurveyItem(stagingId) {
      if (!stagingId) return;
      if (!confirm('Approve runway survey ' + stagingId + '?')) return;
      const approvalPassword = await supervisorRequirePassword_('Enter supervisor password to approve runway survey ' + stagingId + '.');
      if (!approvalPassword) return;
      const notes = (document.getElementById('survey-note-' + stagingId) || {}).value || '';
      const list = document.getElementById('runway-survey-review-list');
      const card = list ? list.querySelector('[data-survey-staging="' + String(stagingId) + '"]') : null;
      const entryKind = String(card && card.dataset && card.dataset.entryKind || '').trim().toUpperCase();
      const runwayClass = String((document.getElementById('survey-rwyclass-' + stagingId) || {}).value || '').trim();
      let mtowInput = null;
      if (entryKind === 'RUNWAY_APPROVAL') {
        const mtowValRaw = (document.getElementById('survey-mtow-' + stagingId) || {}).value;
        const mtowKg = Number(mtowValRaw || 0);
        if (!isFinite(mtowKg) || mtowKg <= 0) {
          if (window.M) M.toast({ html: 'Supervisor MTOW (kg) is required for runway approval.', classes: 'orange' });
          return;
        }
        const cutdownValRaw = (document.getElementById('survey-cutdown-' + stagingId) || {}).value;
        const cutdownAreaM = Number(cutdownValRaw || 0);
        if (!isFinite(cutdownAreaM) || cutdownAreaM <= 0) {
          if (window.M) M.toast({ html: 'Supervisor cutdown value is required for runway approval.', classes: 'orange' });
          return;
        }
        const mtowModel = String((document.getElementById('survey-mtow-model-' + stagingId) || {}).value || '').trim().toUpperCase();
        mtowInput = {
          modelKey: mtowModel || 'GENERIC',
          mtowKg: Math.round(mtowKg),
          cutdownAreaM: Math.round(cutdownAreaM),
          runwayClass: runwayClass
        };
      } else if (runwayClass) {
        mtowInput = { runwayClass: runwayClass };
      }
      google.script.run
        .withSuccessHandler(function(resp) {
          if (resp && resp.success) {
            if (window.M) M.toast({ html: 'Runway survey approved', classes: 'green' });
            refreshRunwaySurveyReviewList();
          } else if (window.M) {
            M.toast({ html: (resp && resp.error) ? resp.error : 'Approve failed', classes: 'red' });
          }
        })
        .withFailureHandler(function(err) {
          if (window.M) M.toast({ html: 'Approve failed: ' + (err && err.message ? err.message : String(err)), classes: 'red' });
        })
        .approveRunwaySurveyReview(stagingId, _supervisorNameForSurvey_(), notes, approvalPassword, mtowInput);
    }

    async function rejectRunwaySurveyItem(stagingId) {
      if (!stagingId) return;
      if (!confirm('Reject runway survey ' + stagingId + '?')) return;
      const approvalPassword = await supervisorRequirePassword_('Enter supervisor password to reject runway survey ' + stagingId + '.');
      if (!approvalPassword) return;
      const notes = (document.getElementById('survey-note-' + stagingId) || {}).value || '';
      google.script.run
        .withSuccessHandler(function(resp) {
          if (resp && resp.success) {
            if (window.M) M.toast({ html: 'Runway survey rejected', classes: 'orange' });
            refreshRunwaySurveyReviewList();
          } else if (window.M) {
            M.toast({ html: (resp && resp.error) ? resp.error : 'Reject failed', classes: 'red' });
          }
        })
        .withFailureHandler(function(err) {
          if (window.M) M.toast({ html: 'Reject failed: ' + (err && err.message ? err.message : String(err)), classes: 'red' });
        })
        .rejectRunwaySurveyReview(stagingId, _supervisorNameForSurvey_(), notes, approvalPassword);
    }

    // â”€â”€â”€ Runway Survey History â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    var _rshCompareSlots_ = [null, null]; // up to 2 entries for comparison

    function openRunwaySurveyHistoryModal() {
      const el = document.getElementById('runway-survey-history-modal');
      if (!el) return;
      // Close the Materialize parent modal first so its overlay doesn't block inputs
      const parentModal = document.getElementById('runway-survey-review-modal');
      if (parentModal && window.M && M.Modal) {
        const instance = M.Modal.getInstance(parentModal);
        if (instance) instance.close();
      }
      el.style.display = 'flex';
      // Focus ICAO input after a short delay for iOS keyboard
      setTimeout(function() {
        const icaoEl = document.getElementById('rsh-icao');
        if (icaoEl) icaoEl.focus();
      }, 200);
    }

    function closeRunwaySurveyHistoryModal() {
      const el = document.getElementById('runway-survey-history-modal');
      if (el) el.style.display = 'none';
      // Reopen the parent Materialize modal
      const parentModal = document.getElementById('runway-survey-review-modal');
      if (parentModal && window.M && M.Modal) {
        const instance = M.Modal.getInstance(parentModal);
        if (instance) instance.open();
      }
    }

    function loadRunwaySurveyHistory() {
      const icao = String((document.getElementById('rsh-icao') || {}).value || '').trim().toUpperCase();
      const rwy = String((document.getElementById('rsh-rwy') || {}).value || '').trim().toUpperCase();
      if (!icao || !rwy) {
        if (window.M) M.toast({ html: 'Enter ICAO and runway', classes: 'orange' });
        return;
      }
      const body = document.getElementById('rsh-body');
      const sub = document.getElementById('rsh-sub');
      if (body) body.innerHTML = '<div style="color:#888; padding:16px; text-align:center;">Loadingâ€¦</div>';
      if (sub) sub.textContent = icao + ' Â· RWY ' + rwy;
      _rshCompareSlots_ = [null, null];
      _rshUpdateComparePanel_();

      google.script.run
        .withSuccessHandler(function(resp) {
          if (!resp || !resp.success) {
            if (body) body.innerHTML = '<div style="color:#d32f2f; padding:12px;">' + ((resp && resp.error) || 'Load failed') + '</div>';
            return;
          }
          // Update sub-header with pair label (e.g. "09/27") if available
          if (sub) sub.textContent = resp.icao + ' Â· RWY ' + (resp.rwyPairLabel || resp.rwyIdent);
          _renderRunwaySurveyHistory_(resp);
        })
        .withFailureHandler(function(err) {
          if (body) body.innerHTML = '<div style="color:#d32f2f; padding:12px;">Error: ' + (err && err.message ? err.message : String(err)) + '</div>';
        })
        .getRunwaySurveyHistory(icao, rwy);
    }

    function _rshFmtDate_(iso) {
      if (!iso) return 'â€”';
      try {
        const d = new Date(iso);
        if (isNaN(d.getTime())) return String(iso).slice(0, 19).replace('T', ' ');
        return d.toLocaleDateString() + ' ' + d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      } catch (e) { return String(iso).slice(0, 19).replace('T', ' '); }
    }

    function _renderRunwaySurveyHistory_(resp) {
      const body = document.getElementById('rsh-body');
      if (!body) return;
      const history = Array.isArray(resp.history) ? resp.history : [];
      const currentVersionId = resp.currentVersion && resp.currentVersion.versionId ? resp.currentVersion.versionId : null;
      if (!history.length) {
        body.innerHTML = '<div style="padding:16px; color:#555; text-align:center;">No approved survey history found for ' + (resp.icao || '--') + ' RWY ' + (resp.rwyIdent || '--') + '.</div>';
        return;
      }

      const reversedHistory = history.slice().reverse(); // newest first

      body.innerHTML = '<div style="font-size:0.78rem; color:#666; margin-bottom:8px;">' + history.length + ' approved version(s). Check up to 2 to compare side-by-side.</div>'
        + reversedHistory.map(function(entry, i) {
          const vid = String(entry.versionId || '');
          const isActive = vid === currentVersionId;
          const vo = entry.verifiedOperational || {};
          const vs = entry.verifiedSurvey || {};
          const originalIndex = history.length - 1 - i; // for data attribute
          return '<div data-rsh-vid="' + _rshEsc_(vid) + '" data-rsh-idx="' + originalIndex + '" style="border:2px solid ' + (isActive ? '#2e7d32' : '#d9e2ef') + '; border-radius:8px; padding:10px; margin-bottom:8px; background:' + (isActive ? '#f1f8e9' : '#fff') + ';">'
            + '<div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:8px;">'
            + '<div style="flex:1 1 auto;">'
            + (isActive ? '<span style="display:inline-block; background:#2e7d32; color:#fff; font-size:0.7rem; font-weight:800; border-radius:4px; padding:1px 7px; margin-bottom:4px;">ACTIVE</span> ' : '')
            + '<span style="font-weight:800; font-size:0.88rem; color:#1565c0;">v' + (history.length - i) + ' Â· ' + _rshFmtDate_(entry.publishedAt) + '</span>'
            + '<div style="font-size:0.78rem; color:#455a64; margin-top:2px;">Pilot: ' + _rshEsc_(entry.pilotName || 'â€”') + ' Â· Approved by: ' + _rshEsc_(entry.approvedBy || 'â€”') + '</div>'
            + '<div style="font-size:0.78rem; color:#455a64;">Captured: ' + _rshFmtDate_(entry.capturedAt) + ' Â· RWY: ' + _rshEsc_(entry.publishedRunway || 'â€”') + '</div>'
            + '<div style="font-size:0.78rem; color:#37474f; margin-top:2px;"><b>Observed</b> ' + (Number(vo.lengthM || 0) || 0) + 'm Ã— ' + (Number(vo.widthM || 0) || 0) + 'm Â· ' + _rshEsc_(vo.surface || 'â€”') + '</div>'
            + (entry.supervisorNotes ? '<div style="font-size:0.76rem; color:#777; font-style:italic; margin-top:2px;">' + _rshEsc_(entry.supervisorNotes) + '</div>' : '')
            + '</div>'
            + '<div style="display:flex; flex-direction:column; gap:5px; min-width:130px; align-items:flex-end;">'
            + '<label style="display:flex; align-items:center; gap:4px; font-size:0.78rem; color:#1565c0; cursor:pointer; white-space:nowrap;">'
            + '<input type="checkbox" data-rsh-compare-vid="' + _rshEsc_(vid) + '" data-rsh-compare-idx="' + i + '" onchange="rshToggleCompare(this, ' + i + ')" style="cursor:pointer;">'
            + 'Compare</label>'
            + (!isActive ? '<button onclick="rshSetActive(\'' + _rshEsc_(vid) + '\', \'' + _rshEsc_(resp.icao) + '\', \'' + _rshEsc_(resp.rwyIdent) + '\')" style="border:none; background:#1565c0; color:#fff; border-radius:6px; padding:4px 12px; font-size:0.78rem; font-weight:800; cursor:pointer;">Set Active</button>' : '')
            + '</div>'
            + '</div>'
            + '</div>';
        }).join('');

      // Store resp on window for compare use
      window._rshCurrentResp_ = resp;
    }

    function _rshEsc_(v) {
      return String(v == null ? '' : v)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    }

    function rshToggleCompare(checkbox, listIndex) {
      const vid = String(checkbox.dataset.rshCompareVid || '');
      if (!window._rshCurrentResp_) return;
      const reversedHistory = (window._rshCurrentResp_.history || []).slice().reverse();
      const entry = reversedHistory[listIndex] || null;
      if (checkbox.checked) {
        const emptySlot = _rshCompareSlots_.indexOf(null);
        if (emptySlot < 0) {
          checkbox.checked = false;
          if (window.M) M.toast({ html: 'Only 2 versions can be compared at once. Uncheck one first.', classes: 'orange' });
          return;
        }
        _rshCompareSlots_[emptySlot] = { vid: vid, entry: entry };
      } else {
        const slot = _rshCompareSlots_.findIndex(function(s) { return s && s.vid === vid; });
        if (slot >= 0) _rshCompareSlots_[slot] = null;
      }
      _rshUpdateComparePanel_();
    }

    function clearRshCompare() {
      _rshCompareSlots_ = [null, null];
      _rshUpdateComparePanel_();
      // Uncheck all compare checkboxes
      const body = document.getElementById('rsh-body');
      if (body) body.querySelectorAll('[data-rsh-compare-vid]').forEach(function(cb) { cb.checked = false; });
    }

    function _rshUpdateComparePanel_() {
      const panel = document.getElementById('rsh-compare-panel');
      const compareBody = document.getElementById('rsh-compare-body');
      if (!panel || !compareBody) return;
      const filled = _rshCompareSlots_.filter(Boolean);
      if (filled.length < 2) { panel.style.display = 'none'; return; }
      panel.style.display = 'block';

      compareBody.innerHTML = filled.map(function(slot, si) {
        const entry = slot.entry || {};
        const vo = entry.verifiedOperational || {};
        const vs = entry.verifiedSurvey || {};
        const isActive = window._rshCurrentResp_ && window._rshCurrentResp_.currentVersion
          && window._rshCurrentResp_.currentVersion.versionId === slot.vid;

        const rows = [
          ['Version ID', String(slot.vid).slice(-24) + 'â€¦'],
          ['Published', _rshFmtDate_(entry.publishedAt)],
          ['Captured', _rshFmtDate_(entry.capturedAt)],
          ['Pilot', entry.pilotName || 'â€”'],
          ['Approved by', entry.approvedBy || 'â€”'],
          ['Notes', entry.supervisorNotes || 'â€”'],
          ['Length (m)', Number(vo.lengthM || 0) || 0],
          ['Width (m)', Number(vo.widthM || 0) || 0],
          ['Surface', vo.surface || 'â€”'],
          ['Features', Array.isArray(vo.features) ? vo.features.length + ' items' : 'â€”'],
          ['Obstacles', Array.isArray(vo.obstacles) ? vo.obstacles.length + ' items' : 'â€”'],
          ['Slope segments', Array.isArray(vo.slopeSegments) ? vo.slopeSegments.length + ' segments' : 'â€”'],
          ['Cutdown area mÂ²', vo.cutdownAreaM != null ? vo.cutdownAreaM : 'â€”'],
          ['Status', vs.status || 'â€”'],
          ['Source RWY', entry.sourceRunway || 'â€”'],
          ['Published RWY', entry.publishedRunway || 'â€”']
        ];

        return '<div style="border:2px solid ' + (isActive ? '#2e7d32' : '#90caf9') + '; border-radius:8px; overflow:hidden;">'
          + '<div style="padding:6px 8px; background:' + (isActive ? '#2e7d32' : '#1565c0') + '; color:#fff; font-size:0.78rem; font-weight:800;">'
          + 'Version ' + (si + 1) + (isActive ? ' Â· ACTIVE' : '') + '</div>'
          + '<table style="width:100%; border-collapse:collapse; font-size:0.77rem;">'
          + rows.map(function(r) {
            return '<tr style="border-bottom:1px solid #eef3f7;">'
              + '<td style="padding:4px 7px; color:#546e7a; white-space:nowrap; font-weight:700;">' + _rshEsc_(r[0]) + '</td>'
              + '<td style="padding:4px 7px; color:#263238;">' + _rshEsc_(String(r[1])) + '</td>'
              + '</tr>';
          }).join('')
          + '</table>'
          + '</div>';
      }).join('');
    }

    async function rshSetActive(versionId, icao, rwyIdent) {
      if (!versionId || !icao || !rwyIdent) return;
      if (!confirm('Set version ' + String(versionId).slice(-20) + 'â€¦ as the active survey for ' + icao + ' RWY ' + rwyIdent + '?')) return;
      const approvalPassword = await supervisorRequirePassword_('Enter supervisor password to change active survey version.');
      if (!approvalPassword) return;

      google.script.run
        .withSuccessHandler(function(resp) {
          if (resp && resp.success) {
            if (window.M) M.toast({ html: 'Active survey version updated', classes: 'green' });
            loadRunwaySurveyHistory(); // refresh list
          } else {
            if (window.M) M.toast({ html: (resp && resp.error) || 'Failed to update active version', classes: 'red' });
          }
        })
        .withFailureHandler(function(err) {
          if (window.M) M.toast({ html: 'Error: ' + (err && err.message ? err.message : String(err)), classes: 'red' });
        })
        .setActiveRunwaySurveyVersion(icao, rwyIdent, versionId, approvalPassword);
    }

