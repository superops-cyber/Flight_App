#!/usr/bin/env python3
import re

file_path = '/Users/jacobanderson/mission-briefing-app/PilotApp.html'

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Update 1: Modify _rwySurveyFilterAirports_ function to show dropdown
old_filter = '''  function _rwySurveyFilterAirports_() {
    const input = document.getElementById('rwysurvey-icao-search');
    const list = _rwySurveyRenderAirportOptions_(input && input.value);
    if (list.length) {
      const sel = document.getElementById('rwysurvey-icao');
      if (sel) sel.value = list[0].icao;
    }
  }'''

new_filter = '''  function _rwySurveyShowAirportOptions_() {
    const dropdown = document.getElementById('rwysurvey-airport-dropdown');
    if (dropdown) dropdown.style.display = 'block';
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
    for (var i = 0; i < rows.length; i++) {
      const r = rows[i];
      const thisIcao = String(r && (r.airportICAO || r.icao || r.ICAO) || '').trim().toUpperCase();
      if (thisIcao === target) return String(r.airportName || r.name || '').trim();
    }
    return '';
  }

  function _rwySurveyFilterAirports_() {
    const input = document.getElementById('rwysurvey-icao-search');
    const searchText = (input && input.value) ? String(input.value).trim().toUpperCase() : '';
    const dropdown = document.getElementById('rwysurvey-airport-dropdown');
    if (!dropdown) return;
    const rows = (window.appData && Array.isArray(window.appData.airports)) ? window.appData.airports : [];
    const icaoSet = {};
    rows.forEach(function(r) {
      const icao = String(r && (r.airportICAO || r.icao || r.ICAO) || '').trim().toUpperCase();
      if (icao && !icaoSet[icao]) {
        const matches = icao.indexOf(searchText) >= 0 || (r && r.airportName && String(r.airportName || '').toUpperCase().indexOf(searchText) >= 0);
        icaoSet[icao] = matches;
      }
    });
    const matches = Object.keys(icaoSet).filter(function(i) { return icaoSet[i]; }).sort();
    if (!searchText) {
      dropdown.style.display = 'none';
      return;
    }
    dropdown.innerHTML = '';
    matches.forEach(function(icao) {
      const name = _rwySurveyGetAirportName_(icao);
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
  }'''

content = content.replace(old_filter, new_filter)

# Update 2: Add checkpoint detection for A+50m and C+50m
old_threshold = '''  function _rwySurveyMaybeNotifyThresholdProximity_(fix) {
    const per = runwaySurveyToolState.perimeter || {};
    const marks = Array.isArray(per.cornerMarks) ? per.cornerMarks : [];
    const prompts = runwaySurveyToolState.prompts || { a50Shown: false, a50Completed: false, a50Lat: null, a50Lon: null };
    runwaySurveyToolState.prompts = prompts;
    if (!fix || prompts.a50Shown || prompts.a50Completed) return;
    if (marks.length !== 1 || per.closed) return;

    const distSinceA = Number(per.liveSinceMark || 0);
    if (distSinceA < 50) return;

    prompts.a50Shown = true;
    prompts.a50Lat = Number(fix.lat || 0);
    prompts.a50Lon = Number(fix.lon || 0);
    _rwySurveyLogEvent_('prompt', 'A+50m checkpoint reached');
    openRunwaySurveyObstaclePopup(true);
  }'''

new_threshold = '''  function _rwySurveyMaybeNotifyThresholdProximity_(fix) {
    const per = runwaySurveyToolState.perimeter || {};
    const marks = Array.isArray(per.cornerMarks) ? per.cornerMarks : [];
    const prompts = runwaySurveyToolState.prompts || { a50Shown: false, a50Completed: false, a50Lat: null, a50Lon: null, c50Shown: false, c50Completed: false, c50Lat: null, c50Lon: null };
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
      openRunwaySurveyObstaclePopup(true, 'A');
      return;
    }

    // C+50m checkpoint: only when exactly 3 corners marked and not closed
    if (!prompts.c50Shown && !prompts.c50Completed && marks.length === 3 && !per.closed && distSinceLastMark >= 50) {
      prompts.c50Shown = true;
      prompts.c50Lat = Number(fix.lat || 0);
      prompts.c50Lon = Number(fix.lon || 0);
      _rwySurveyLogEvent_('prompt', 'C+50m checkpoint reached');
      openRunwaySurveyObstaclePopup(true, 'C');
      return;
    }
  }'''

content = content.replace(old_threshold, new_threshold)

# Update 3: Add diagram preview function before the closing script tag
diagram_func = '''

  function openRunwayDiagramPreview() {
    const state = runwaySurveyToolState;
    if (!state.perimeter || !state.perimeter.cornerMarks || state.perimeter.cornerMarks.length < 2) {
      if (window.M) M.toast({ html: 'Mark at least 2 corners first (A and B)', classes: 'orange' });
      return;
    }

    const modal = document.createElement('div');
    modal.id = 'rwy-diagram-preview-modal';
    modal.style.cssText = 'position:fixed; inset:0; z-index:9999; background:rgba(0,0,0,0.6); display:flex; align-items:center; justify-content:center; padding:20px;';
    
    const content = document.createElement('div');
    content.style.cssText = 'background:#fff; border-radius:8px; padding:20px; max-width:600px; width:100%; box-shadow:0 8px 32px rgba(0,0,0,0.3);';
    
    const header = document.createElement('div');
    header.style.cssText = 'display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; border-bottom:2px solid #1565c0; padding-bottom:8px;';
    header.innerHTML = '<h3 style="margin:0; color:#1565c0;">Runway Diagram Preview</h3><button onclick="document.getElementById(\\'rwy-diagram-preview-modal\\').remove()" style="border:none; background:none; font-size:20px; cursor:pointer; color:#999;">×</button>';
    
    const diagram = document.createElement('svg');
    diagram.setAttribute('width', '500');
    diagram.setAttribute('height', '300');
    diagram.setAttribute('viewBox', '0 0 500 300');
    diagram.style.cssText = 'border:1px solid #ccc; border-radius:4px; background:#f9f9f9; width:100%; height:auto;';
    
    // Draw runway perimeter based on corners
    const marks = state.perimeter.cornerMarks || [];
    if (marks.length >= 2) {
      // Simple representation: draw runway as box
      const corners = { A: marks[0], B: marks[1], C: marks[2], D: marks[3] };
      const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      rect.setAttribute('x', '50');
      rect.setAttribute('y', '50');
      rect.setAttribute('width', '400');
      rect.setAttribute('height', '200');
      rect.setAttribute('fill', 'none');
      rect.setAttribute('stroke', '#333');
      rect.setAttribute('stroke-width', '2');
      diagram.appendChild(rect);
      
      // Add corner labels
      const labels = ['A', 'B', 'C', 'D'];
      const positions = [[50, 50], [450, 50], [450, 250], [50, 250]];
      labels.forEach((label, idx) => {
        const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        text.setAttribute('x', positions[idx][0]);
        text.setAttribute('y', positions[idx][1] - 10);
        text.setAttribute('font-size', '14');
        text.setAttribute('font-weight', 'bold');
        text.setAttribute('fill', '#6a1b9a');
        text.textContent = label;
        diagram.appendChild(text);
      });
      
      // Draw features
      (state.features || []).forEach((feat, idx) => {
        const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
        circle.setAttribute('cx', String(80 + idx * 30));
        circle.setAttribute('cy', '120');
        circle.setAttribute('r', '4');
        circle.setAttribute('fill', '#1976d2');
        diagram.appendChild(circle);
        
        const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        label.setAttribute('x', String(80 + idx * 30));
        label.setAttribute('y', '140');
        label.setAttribute('font-size', '10');
        label.setAttribute('text-anchor', 'middle');
        label.setAttribute('fill', '#333');
        label.textContent = (feat.name || 'F').substring(0, 1) + (idx + 1);
        diagram.appendChild(label);
      });
      
      // Note about obstacles
      const note = document.createElement('div');
      note.style.cssText = 'margin-top:12px; padding:8px; background:#e3f2fd; border-left:4px solid #1976d2; font-size:0.85rem; color:#1565c0;';
      note.textContent = 'Diagram shows basic runway representation with ' + (state.features || []).length + ' feature(s) and ' + (state.obstacleAngles50m || []).length + ' obstacle angle(s) captured.';
      
      content.appendChild(header);
      content.appendChild(diagram);
      content.appendChild(note);
    }
    
    modal.appendChild(content);
    document.body.appendChild(modal);
    modal.onclick = function(e) { if (e.target === modal) modal.remove(); };
  }
'''

# Insert diagram function before the final closing script - find </script> tag near end
script_close_pattern = r'(\n\s*</script>)'
content = re.sub(script_close_pattern, diagram_func + r'\1', content)

# Write updated content back
with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("✓ Updated _rwySurveyFilterAirports_ with autocomplete dropdown")
print("✓ Updated _rwySurveyMaybeNotifyThresholdProximity_ for A+50m and C+50m")
print("✓ Added openRunwayDiagramPreview() function")
print("\nDone!")
