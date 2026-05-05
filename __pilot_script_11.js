
// ══════════════════════════════════════════════════════════════════
//  RUNWAY BRIEFING CARD  –  controller
// ══════════════════════════════════════════════════════════════════

window._rbcState = {
  icao: '',
  rwyIdent: '',
  runway: null,          // full runway object from performanceTab4 / appData
  card: {},              // loaded/saved card data blob
  editMode: false,
  dirty: false,
  mapDistM: null,        // go-around distance in metres from threshold
  diagramScale: 1,       // pixels per metre (set when diagram is drawn)
  diagramTopX: 0,
  diagramTopY: 0,
  diagramTopH: 0,
  diagramRunwayPx: 0,
  callerContext: 'tab5'  // 'tab5' | 'tab7'
};

// ── open / close ─────────────────────────────────────────────────

window.openRunwayBriefingCard = function(icao, rwyIdent, callerContext) {
  const s = window._rbcState;
  s.icao = String(icao || '').trim().toUpperCase();
  s.rwyIdent = String(rwyIdent || '').trim().toUpperCase();
  s.callerContext = callerContext || 'tab5';
  s.editMode = false;
  s.dirty = false;
  s.mapDistM = null;

  // Resolve runway object
  s.runway = _rbcFindRunway(s.icao, s.rwyIdent);

  const overlay = document.getElementById('rbc-overlay');
  if (overlay) overlay.style.display = 'block';
  document.body.style.overflow = 'hidden';

  rbcSwitchPage(1);
  _rbcRenderPage1();
  _rbcLoadAndRenderPage2();
  _rbcSetEditUI(false);
};

window.rbcClose = function() {
  const overlay = document.getElementById('rbc-overlay');
  if (overlay) overlay.style.display = 'none';
  document.body.style.overflow = '';
};

// ── page switching ────────────────────────────────────────────────

window.rbcSwitchPage = function(n) {
  [1,2].forEach(function(i) {
    const page = document.getElementById('rbc-page-' + i);
    const tab = document.getElementById('rbc-tab-btn-' + i);
    if (page) page.classList.toggle('active', i === n);
    if (tab)  tab.classList.toggle('active', i === n);
  });
};

// ── helpers ────────────────────────────────────────────────────────

function _rbcEsc(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function _rbcFindRunway(icao, rwyIdent) {
  const perf = window.performanceTab4 || {};
  const runways = Array.isArray(perf.runways) ? perf.runways : [];
  // Try exact match first
  let rwy = runways.find(r => String(r.rwyIdent||'').toUpperCase() === rwyIdent);
  if (!rwy) {
    // Try ICAO match via appData
    const airports = (window.appData && Array.isArray(window.appData.airports)) ? window.appData.airports : [];
    const apt = airports.find(a => String(a.icao||'').toUpperCase() === icao);
    if (apt) {
      // Synthesise a minimal runway object from apt
      rwy = {
        icao: icao,
        rwyIdent: rwyIdent,
        length: parseFloat(apt.runwayLength) || 0,
        width: parseFloat(apt.runwayWidth) || 0,
        slope: parseFloat(apt.runwaySlopePercent) || 0,
        surface: apt.runwaySurfaceActual || '',
        elevation: parseFloat(apt.elevationFt) || 0,
        knownFeatures: [],
        slopeProfile: [],
        surveySlopeSegments: [],
        obstacleAngles: [],
        officialReference: { lengthM: parseFloat(apt.runwayLength)||0, widthM: parseFloat(apt.runwayWidth)||0 },
        verifiedOperational: {},
        cutdownAreaM: 0,
        airstripPhoto: apt.airstripPhoto || '',
        pilotNotes: apt.pilotNotes || '',
        knownFeaturesRaw: apt.knownFeatures || ''
      };
    }
  }
  return rwy || null;
}

function _rbcLoadCard() {
  // Card data lives in KNOWN_FEATURES JSON under key "briefingCard"
  const s = window._rbcState;
  const rwy = s.runway;
  if (!rwy) return {};

  // Try from verifiedOperational.briefingCard or top-level knownObj
  let blob = {};
  try {
    const raw = String((rwy && rwy.knownFeaturesRaw) || '').trim();
    if (raw) {
      const parsed = JSON.parse(raw);
      const obj = Array.isArray(parsed) ? {} : (parsed || {});
      blob = (obj.briefingCard && typeof obj.briefingCard === 'object') ? obj.briefingCard : {};
    }
  } catch (e) {}

  // Also check verifiedOperational directly (set by newer approval flow)
  if (!Object.keys(blob).length) {
    const vo = (rwy && rwy.verifiedOperational) || {};
    blob = (vo.briefingCard && typeof vo.briefingCard === 'object') ? vo.briefingCard : {};
  }

  return blob;
}

function _rbcRunwayClass(rwy) {
  if (!rwy) return '?';
  const l = Number(rwy.length || 0);
  if (l >= 900) return '1';
  if (l >= 600) return '2';
  return '3';
}

// ── PAGE 1 render ─────────────────────────────────────────────────

function _rbcRenderPage1() {
  const s = window._rbcState;
  const rwy = s.runway;
  s.card = _rbcLoadCard();
  const card = s.card;

  // Header title
  const title = document.getElementById('rbc-header-title');
  if (title) title.textContent = (s.icao || '?') + ' · ' + (s.rwyIdent || '?') + ' Briefing Card';

  // Ident grid
  _rbcSet('rbc-class', 'Class ' + _rbcRunwayClass(rwy));
  _rbcSet('rbc-ident', s.rwyIdent || '–');
  _rbcSet('rbc-icao', s.icao || '–');

  // Coordinates from verifiedOperational thresholds or knownObj
  let coordsText = '–';
  try {
    const vo = (rwy && rwy.verifiedOperational) || {};
    const thr = Array.isArray(vo.thresholds) ? vo.thresholds : [];
    if (thr.length) {
      coordsText = thr.map(function(t) {
        const lat = Number(t.lat || t.latitude || 0).toFixed(5);
        const lon = Number(t.lon || t.longitude || 0).toFixed(5);
        const id  = String(t.ident || t.rwy || '').toUpperCase();
        return (id ? id + ': ' : '') + lat + ', ' + lon;
      }).join('\n');
    }
  } catch (e) {}
  _rbcSet('rbc-coords', coordsText);

  // Dimensions
  const intLen  = rwy ? Math.round(Number(rwy.length || 0)) : 0;
  const intWid  = rwy ? Math.round(Number(rwy.width  || 0)) : 0;
  const offLen  = rwy ? Math.round(Number((rwy.officialReference && rwy.officialReference.lengthM) || intLen || 0)) : 0;
  const offWid  = rwy ? Math.round(Number((rwy.officialReference && rwy.officialReference.widthM)  || intWid || 0)) : 0;
  const cutdown = rwy ? (Number(rwy.cutdownAreaM || 0) > 0 ? Math.round(rwy.cutdownAreaM) + ' m' : '–') : '–';
  const elev    = rwy ? Math.round(Number(rwy.elevation || 0)) + ' ft' : '–';
  const surface = rwy ? (rwy.surface || '–') : '–';
  const slopePct = rwy ? Number(rwy.slope || 0).toFixed(1) + '%' : '–';
  const survDate = rwy ? String(rwy.internalUpdatedAt || '').substring(0, 10) || '–' : '–';

  _rbcSet('rbc-int-length', intLen ? intLen + ' m' : '–');
  _rbcSet('rbc-off-length', offLen ? offLen + ' m' : '–');
  _rbcSet('rbc-int-width',  intWid ? intWid + ' m' : '–');
  _rbcSet('rbc-off-width',  offWid ? offWid + ' m' : '–');
  _rbcSet('rbc-cutdown',    cutdown);
  _rbcSet('rbc-elevation',  elev);
  _rbcSet('rbc-surface',    surface);
  _rbcSet('rbc-slope',      slopePct);
  _rbcSet('rbc-survey-date', survDate);

  // Slope detail labels
  const slopeDetailEl = document.getElementById('rbc-slope-detail-row');
  if (slopeDetailEl && rwy) {
    const segs = Array.isArray(rwy.surveySlopeSegments) && rwy.surveySlopeSegments.length
      ? rwy.surveySlopeSegments
      : (Array.isArray(rwy.slopeProfile) ? rwy.slopeProfile : []);
    if (segs.length > 1) {
      const labels = segs.map(function(seg) {
        const dist = Math.round(Number(seg.distanceM || seg.distance || 0));
        const slp  = Number(seg.slope || 0).toFixed(1);
        return dist + 'm ' + (seg.slope >= 0 ? '+' : '') + slp + '%';
      }).join('  →  ');
      slopeDetailEl.innerHTML = '<div style="font-size:0.72rem; color:#546e7a; font-weight:700; padding-bottom:4px;">⛰ Slope profile: ' + _rbcEsc(labels) + '</div>';
    } else {
      slopeDetailEl.innerHTML = '';
    }
  }

  // Diagram
  _rbcDrawDiagram();

  // Photo
  const photoBox = document.getElementById('rbc-photo-box');
  if (photoBox) {
    const photoUrl = String((rwy && rwy.airstripPhoto) || (card && card.airstripPhoto) || '').trim();
    if (photoUrl) {
      photoBox.innerHTML = '<img src="' + _rbcEsc(photoUrl) + '" alt="Airstrip photo" onerror="this.parentNode.innerHTML=\'<span>Photo unavailable</span>\'">';
    } else {
      photoBox.innerHTML = '<span>No photo — tap EDIT to add URL</span>';
    }
  }

  // Restrictions
  _rbcRenderRestrictions();

  // Text blocks
  _rbcSetBlock('obstructions', card.obstructions  || (rwy && _rbcBuildDefaultObstructions(rwy)) || '–');
  _rbcSetBlock('weather',      card.weather       || '–');
  _rbcSetBlock('dangers',      card.dangers       || '–');
  _rbcSetBlock('forced',       card.forcedLanding || '–');

  // MAP / go-around
  if (card.mapDistM != null) {
    s.mapDistM = Number(card.mapDistM);
    _rbcShowMapInfo();
  }

  // Revision footer
  const el1 = document.getElementById('rbc-revision-date');
  const el2 = document.getElementById('rbc-revised-by');
  if (el1) el1.textContent = String(card.revisedAt || '').substring(0,10) || '–';
  if (el2) el2.textContent = String(card.revisedBy || '') || '–';
}

function _rbcBuildDefaultObstructions(rwy) {
  const feats = Array.isArray(rwy.knownFeatures) ? rwy.knownFeatures : [];
  if (!feats.length) return '';
  return feats.map(function(f) {
    return '• ' + String(f.name || '') + (f.distance ? ' @ ' + Math.round(f.distance) + 'm' : '') + (f.side ? ' (' + f.side + ')' : '');
  }).join('\n');
}

function _rbcSet(id, text) {
  const el = document.getElementById(id);
  if (el) el.textContent = text;
}

function _rbcSetBlock(key, text) {
  const view = document.getElementById('rbc-' + key + '-view');
  const edit  = document.getElementById('rbc-' + key + '-edit');
  if (view) view.textContent = text || '–';
  if (edit) edit.value = (text && text !== '–') ? text : '';
}

function _rbcRenderRestrictions() {
  const s = window._rbcState;
  const card = s.card;
  const rwy = s.runway;
  const tbody = document.getElementById('rbc-restrictions-tbody');
  if (!tbody) return;

  let rows = [];

  // From card
  if (Array.isArray(card.restrictions) && card.restrictions.length) {
    rows = card.restrictions.slice();
  }

  // Auto-populate from mtowByModel if no card rows
  if (!rows.length) {
    const airports = (window.appData && Array.isArray(window.appData.airports)) ? window.appData.airports : [];
    const apt = airports.find(function(a) { return String(a.icao||'').toUpperCase() === s.icao; });
    const mtow = (apt && apt.mtowByModel && typeof apt.mtowByModel === 'object') ? apt.mtowByModel : {};
    Object.keys(mtow).forEach(function(modelKey) {
      const kg = Number(mtow[modelKey] || 0);
      if (kg > 0) rows.push({ acft: modelKey, dir: 'All', mtow: kg, notes: 'From DB_Airports MTOW' });
    });
  }

  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="4" style="color:#90a4ae; text-align:center;">No restrictions recorded</td></tr>';
    return;
  }

  tbody.innerHTML = rows.map(function(r) {
    const mtowVal = Number(r.mtow || 0);
    const mtowClass = mtowVal > 0 ? 'ok' : '';
    return '<tr><td>' + _rbcEsc(r.acft||'–') + '</td>'
      + '<td>' + _rbcEsc(r.dir||'–') + '</td>'
      + '<td class="' + mtowClass + '">' + (mtowVal > 0 ? mtowVal + ' kg' : '–') + '</td>'
      + '<td>' + _rbcEsc(r.notes||'–') + '</td></tr>';
  }).join('');
}

// ── DIAGRAM (reuses window.drawRunwayDiagram logic, self-contained) ──

function _rbcDrawDiagram() {
  const s = window._rbcState;
  const rwy = s.runway;
  if (!rwy) return;

  // Delegate to Tab4's drawRunwayDiagram if available, drawing into our SVG
  // We swap the element id temporarily
  const svg = document.getElementById('rbc-diagram-svg');
  if (!svg) return;

  // Use the shared drawing function by temporarily redirecting it
  if (typeof window.drawRunwayDiagram === 'function') {
    // We need to draw into OUR svg. The Tab4 function uses 'perf-runway-diagram'.
    // We'll call rbcDrawRunwayDiagramInto() — our own standalone version.
  }
  _rbcDrawRunwayDiagramInto(svg, rwy, null, s.mapDistM);
}

function _rbcDrawRunwayDiagramInto(svg, runway, result, mapDistM) {
  if (!svg || !runway) return;

  const makeEl = function(tag, attrs, text) {
    const el = document.createElementNS('http://www.w3.org/2000/svg', tag);
    Object.keys(attrs || {}).forEach(function(k) { el.setAttribute(k, String(attrs[k])); });
    if (text != null) el.textContent = text;
    return el;
  };

  const surfaceNorm = String(runway.surface || '').toUpperCase().replace(/[^A-Z]/g, '');
  const isAsphalt = /PAVED|ASPHALT|CONCRETE|ASFALTO|CONCRETO/.test(surfaceNorm);
  const isGrass   = /GRASS|TURF|GRAMA/.test(surfaceNorm);
  const lengthM   = Math.max(Number(runway.length || 0), 1);
  const widthM    = Math.max(Number(runway.width  || 0), 0);

  const canvasW   = 700;
  const canvasH   = 470;
  const topX      = 40;
  const topY      = 42;
  const topHeight = 62;
  const usableWidth = 620;
  const scale     = usableWidth / lengthM;
  const runwayPx  = lengthM * scale;

  // Store for MAP placement
  const st = window._rbcState;
  st.diagramScale     = scale;
  st.diagramTopX      = topX;
  st.diagramTopY      = topY;
  st.diagramTopH      = topHeight;
  st.diagramRunwayPx  = runwayPx;

  svg.innerHTML = '';

  const defs = makeEl('defs', {});
  const gp = makeEl('pattern', { id: 'rbcGrass', patternUnits: 'userSpaceOnUse', width: 12, height: 12 });
  gp.appendChild(makeEl('rect', { x:0, y:0, width:12, height:12, fill:'#6d9557' }));
  gp.appendChild(makeEl('path', { d:'M0,12 L12,0 M-3,9 L3,3 M9,15 L15,9', stroke:'#7fab66', 'stroke-width':1 }));
  defs.appendChild(gp);
  const rp = makeEl('pattern', { id: 'rbcRough', patternUnits: 'userSpaceOnUse', width: 8, height: 8 });
  rp.appendChild(makeEl('rect', { x:0, y:0, width:8, height:8, fill:'#8b6f47' }));
  rp.appendChild(makeEl('circle', { cx:2, cy:2, r:1, fill:'#6f5837' }));
  rp.appendChild(makeEl('circle', { cx:6, cy:5, r:1, fill:'#6f5837' }));
  defs.appendChild(rp);
  svg.appendChild(defs);

  const fill = isAsphalt ? '#5f5f5f' : (isGrass ? 'url(#rbcGrass)' : (/ROUGH|MUD|SAND/.test(surfaceNorm) ? 'url(#rbcRough)' : '#9a9a9a'));

  const topGroup = makeEl('g', {});
  topGroup.appendChild(makeEl('text', { x:topX, y:24, 'font-size':14, fill:'#333', 'font-weight':'700' }, 'TOP VIEW'));
  topGroup.appendChild(makeEl('rect', { x:topX, y:topY, width:runwayPx, height:topHeight, fill:fill, stroke:'#2a2a2a', 'stroke-width':1.5, rx:2 }));

  if (isAsphalt) {
    topGroup.appendChild(makeEl('line', { x1:topX, y1:topY+topHeight/2, x2:topX+runwayPx, y2:topY+topHeight/2, stroke:'#f3f3f3', 'stroke-width':2, 'stroke-dasharray':'12,8' }));
  }

  const identNum = parseInt(String(runway.rwyIdent || '').replace(/\D/g,''), 10);
  const thisRwy  = (!isNaN(identNum) && identNum >= 1 && identNum <= 36) ? String(identNum).padStart(2,'0') : String(runway.rwyIdent || '??');
  const recipN   = (!isNaN(identNum) && identNum >= 1 && identNum <= 36) ? String((((identNum+18-1)%36)+1)).padStart(2,'0') : '';

  const addBadge = function(cx, angle, label) {
    const g = makeEl('g', { transform:'rotate('+angle+' '+cx+' '+(topY+topHeight/2)+')' });
    g.appendChild(makeEl('rect', { x:cx-18, y:(topY+topHeight/2)-13, width:36, height:26, fill:'rgba(0,0,0,0.45)', rx:3 }));
    g.appendChild(makeEl('text', { x:cx, y:(topY+topHeight/2)+6, 'font-size':16, fill:'#fff', 'font-weight':'900', 'text-anchor':'middle' }, label));
    topGroup.appendChild(g);
  };
  addBadge(topX+22, 90, thisRwy);
  if (recipN) addBadge(topX+runwayPx-22, -90, recipN);

  // Width label
  topGroup.appendChild(makeEl('text', { x:topX+(runwayPx/2), y:topY+topHeight+22, 'font-size':13, fill:'#333', 'text-anchor':'middle' },
    Math.round(lengthM)+'m × '+ Math.round(widthM||0)+'m · '+(runway.surface||'')));

  // Official ref
  const offRef = runway.officialReference || {};
  if (offRef.lengthM) {
    topGroup.appendChild(makeEl('text', { x:topX+(runwayPx/2), y:topY+topHeight+36, 'font-size':12, fill:'#0b5394', 'font-weight':'700', 'text-anchor':'middle' },
      'OFFICIAL: '+Math.round(offRef.lengthM||0)+'m × '+Math.round(offRef.widthM||0)+'m'));
  }

  // Features
  const features = Array.isArray(runway.knownFeatures) ? runway.knownFeatures : [];
  features.forEach(function(feat) {
    const dist = Math.max(0, Number(feat.distance || 0));
    if (dist > lengthM) return;
    const side = String(feat.side||'right').toLowerCase();
    const fx = topX + (dist * scale);
    const fy = side === 'left' ? topY - 18 : topY + topHeight + 18;
    topGroup.appendChild(makeEl('line', { x1:fx, y1:side==='left'?topY:topY+topHeight, x2:fx, y2:fy, stroke:'#ff9800', 'stroke-width':1.5 }));
    topGroup.appendChild(makeEl('text', { x:fx, y:fy, 'font-size':14, 'text-anchor':'middle' }, '●'));
    topGroup.appendChild(makeEl('text', { x:fx, y:fy+(side==='left'?-6:12), 'font-size':11, fill:'#555', 'text-anchor':'middle' }, Math.round(dist)+'m'));
  });

  // MAP / go-around Maltese cross
  if (mapDistM != null && mapDistM >= 0) {
    const mx = topX + (mapDistM * scale);
    const my = topY + topHeight/2;
    const arm = 10;
    topGroup.appendChild(makeEl('line', { x1:mx-arm, y1:my, x2:mx+arm, y2:my, stroke:'#6a1b9a', 'stroke-width':3 }));
    topGroup.appendChild(makeEl('line', { x1:mx, y1:my-arm, x2:mx, y2:my+arm, stroke:'#6a1b9a', 'stroke-width':3 }));
    topGroup.appendChild(makeEl('line', { x1:mx-arm*0.6, y1:my-arm*0.6, x2:mx+arm*0.6, y2:my+arm*0.6, stroke:'#6a1b9a', 'stroke-width':2 }));
    topGroup.appendChild(makeEl('line', { x1:mx+arm*0.6, y1:my-arm*0.6, x2:mx-arm*0.6, y2:my+arm*0.6, stroke:'#6a1b9a', 'stroke-width':2 }));
    topGroup.appendChild(makeEl('text', { x:mx, y:topY-6, 'font-size':11, fill:'#6a1b9a', 'font-weight':'800', 'text-anchor':'middle' }, 'MAP'));
  }

  svg.appendChild(topGroup);

  // ── SIDE VIEW ─────────────────────────────────────────────────

  let profile = [];
  const segs = Array.isArray(runway.surveySlopeSegments) && runway.surveySlopeSegments.length
    ? runway.surveySlopeSegments
    : (Array.isArray(runway.slopeProfile) ? runway.slopeProfile : []);

  if (segs.length) {
    const ident = String(runway.rwyIdent||'').toUpperCase();
    let cursor = 0;
    const sorted = segs.map(function(seg) {
      const segDist = Math.max(Number(seg.distanceM != null ? seg.distanceM : (seg.distance||0)), 0);
      const rawStart = Number(seg.startDistanceM);
      const startM = isFinite(rawStart) ? rawStart : cursor;
      return { startDistanceM: Math.max(0, Math.min(lengthM, startM)), distance: segDist, slope: Number(seg.slope||0) };
    }).filter(function(s) { return s.distance > 0; }).sort(function(a,b) { return a.startDistanceM - b.startDistanceM; });
    cursor = 0;
    sorted.forEach(function(seg) {
      if (seg.startDistanceM > cursor) profile.push({ distance: seg.startDistanceM - cursor, slope: 0 });
      const clipped = Math.max(0, Math.min(seg.distance, lengthM - seg.startDistanceM));
      if (clipped > 0) { profile.push({ distance: clipped, slope: seg.slope }); cursor = seg.startDistanceM + clipped; }
    });
    if (cursor < lengthM) profile.push({ distance: lengthM - cursor, slope: 0 });
  } else {
    profile = [{ distance: lengthM, slope: Number(runway.slope||0) }];
  }

  const sideX = topX, sideY = 308, sideH = 84;
  const elevations = [0];
  let curEl = 0;
  profile.forEach(function(seg) { curEl += seg.distance*seg.slope/100; elevations.push(curEl); });
  const minEl = Math.min.apply(null, elevations);
  const maxEl = Math.max.apply(null, elevations);
  const spanEl = Math.max(0.5, maxEl - minEl);
  const elScale = Math.max(0.8, Math.min(sideH / spanEl, 4));
  const usedH = spanEl * elScale;
  const yOff = (sideH - usedH) / 2;

  let xCur = 0;
  const pts = [sideX + ',' + (sideY + yOff + usedH - ((elevations[0]-minEl)*elScale))];
  profile.forEach(function(seg, i) {
    xCur += seg.distance;
    const px = sideX + (xCur * scale);
    const py = sideY + yOff + usedH - ((elevations[i+1]-minEl)*elScale);
    pts.push(px + ',' + py);
  });

  const sideGroup = makeEl('g', {});
  sideGroup.appendChild(makeEl('text', { x:sideX, y:sideY-6, 'font-size':13, fill:'#333', 'font-weight':'700' }, 'SIDE VIEW'));
  sideGroup.appendChild(makeEl('rect', { x:sideX, y:sideY, width:runwayPx, height:sideH, fill:'#f3f3f3', stroke:'#ddd' }));
  sideGroup.appendChild(makeEl('polyline', { points:pts.join(' '), fill:'none', stroke:'#00695c', 'stroke-width':3 }));

  // Slope labels
  let segPos = 0;
  profile.forEach(function(seg) {
    const midX = sideX + ((segPos + seg.distance/2) * scale);
    const lbl = Math.round(seg.distance)+'m '+(seg.slope>=0?'+':'')+Number(seg.slope).toFixed(1)+'%';
    sideGroup.appendChild(makeEl('text', { x:midX, y:sideY+sideH+22, 'font-size':11, fill:'#4b4b4b', 'text-anchor':'middle' }, lbl));
    segPos += seg.distance;
  });

  sideGroup.appendChild(makeEl('text', { x:sideX+2, y:sideY+sideH+14, 'font-size':11, fill:'#546e7a' }, thisRwy||'??'));
  if (recipN) sideGroup.appendChild(makeEl('text', { x:sideX+runwayPx-2, y:sideY+sideH+14, 'font-size':11, fill:'#546e7a', 'text-anchor':'end' }, recipN));

  svg.appendChild(sideGroup);
  svg.setAttribute('viewBox', '0 0 ' + canvasW + ' ' + canvasH);
}

function _rbcShowMapInfo() {
  const s = window._rbcState;
  const el = document.getElementById('rbc-map-info');
  const ds = document.getElementById('rbc-map-dist');
  if (!el) return;
  if (s.mapDistM != null) {
    el.style.display = 'block';
    if (ds) ds.textContent = Math.round(s.mapDistM) + ' m';
  } else {
    el.style.display = 'none';
  }
}

// ── MAP placement (tap on diagram) ───────────────────────────────

function _rbcEnableDiagramTap(enabled) {
  const container = document.getElementById('rbc-diagram-container');
  const hint = document.getElementById('rbc-diagram-drop-hint');
  const editHint = document.getElementById('rbc-diagram-edit-hint');
  if (!container) return;

  if (enabled) {
    container.style.cursor = 'crosshair';
    if (hint) hint.textContent = 'Tap the diagram to place the Go-Around Point (MAP) ✛';
    if (editHint) editHint.style.display = 'inline';
    container.onclick = function(e) {
      const svgEl = document.getElementById('rbc-diagram-svg');
      if (!svgEl) return;
      const rect = svgEl.getBoundingClientRect();
      // Map click to SVG coordinates
      const svgW = 700;
      const svgH = 470;
      const clickX = (e.clientX - rect.left) * (svgW / rect.width);
      const st = window._rbcState;
      const localX = clickX - st.diagramTopX;
      const distM = Math.max(0, Math.min(Math.round(localX / st.diagramScale), Math.round(Number((st.runway && st.runway.length)||0))));
      st.mapDistM = distM;
      st.dirty = true;
      _rbcDrawDiagram();
      _rbcShowMapInfo();
    };
  } else {
    container.style.cursor = '';
    if (hint) hint.textContent = '';
    if (editHint) editHint.style.display = 'none';
    container.onclick = null;
  }
}

// ── inline text block editing ────────────────────────────────────

window.rbcEditBlock = function(key) {
  if (!window._rbcState.editMode) return;
  const view = document.getElementById('rbc-' + key + '-view');
  const edit  = document.getElementById('rbc-' + key + '-edit');
  if (!view || !edit) return;
  view.style.display = 'none';
  edit.style.display = 'block';
  edit.focus();
  edit.onblur = function() {
    const newVal = edit.value.trim() || '–';
    view.textContent = newVal;
    view.style.display = 'block';
    edit.style.display = 'none';
    window._rbcState.dirty = true;
    // Persist to card blob
    const keyMap = { obstructions:'obstructions', weather:'weather', dangers:'dangers', forced:'forcedLanding', routes:'routes', rwyhist:'rwyhist', incidents:'incidents', othernotes:'othernotes' };
    if (keyMap[key]) window._rbcState.card[keyMap[key]] = newVal !== '–' ? newVal : '';
  };
};

// ── restrictions ──────────────────────────────────────────────────

window.rbcAddRestrictionRow = function() {
  const acft  = String(document.getElementById('rbc-r-acft')?.value || '').trim();
  const dir   = String(document.getElementById('rbc-r-dir')?.value  || '').trim();
  const mtow  = Number(document.getElementById('rbc-r-mtow')?.value  || 0);
  const notes = String(document.getElementById('rbc-r-notes')?.value || '').trim();
  if (!acft) { if (window.M) M.toast({html:'Enter aircraft / model name', classes:'orange'}); return; }

  const s = window._rbcState;
  if (!Array.isArray(s.card.restrictions)) s.card.restrictions = [];
  s.card.restrictions.push({ acft, dir: dir||'All', mtow: mtow||0, notes });
  s.dirty = true;
  _rbcRenderRestrictions();
  ['rbc-r-acft','rbc-r-dir','rbc-r-mtow','rbc-r-notes'].forEach(function(id) {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
};

// ── edit mode ────────────────────────────────────────────────────

function _rbcSetEditUI(editing) {
  const s = window._rbcState;
  s.editMode = editing;

  const editBtn   = document.getElementById('rbc-edit-btn');
  const saveBtn   = document.getElementById('rbc-save-btn');
  const cancelBtn = document.getElementById('rbc-cancel-btn');
  const reEditor  = document.getElementById('rbc-restrictions-editor');

  if (editBtn)   editBtn.style.display   = editing ? 'none'  : 'inline-block';
  if (saveBtn)   saveBtn.style.display   = editing ? 'inline-block' : 'none';
  if (cancelBtn) cancelBtn.style.display = editing ? 'inline-block' : 'none';
  if (reEditor)  reEditor.style.display  = editing ? 'block' : 'none';

  // Show/hide editable hints
  const blocks = ['obstructions','weather','dangers','forced','routes','rwyhist','incidents','othernotes'];
  blocks.forEach(function(key) {
    const view = document.getElementById('rbc-' + key + '-view');
    if (view) view.classList.toggle('editable-block', editing);
  });

  _rbcEnableDiagramTap(editing);
}

window.rbcToggleEdit = function() {
  _rbcSetEditUI(true);
  _rbcSetStatus('Editing — tap text blocks to edit. Tap diagram to place MAP ✛.');
};

window.rbcCancelEdit = function() {
  _rbcSetEditUI(false);
  window._rbcState.dirty = false;
  _rbcRenderPage1(); // reload original
  _rbcSetStatus('');
};

// ── save ─────────────────────────────────────────────────────────

window.rbcSave = function() {
  const s = window._rbcState;
  if (!s.icao || !s.rwyIdent) {
    _rbcSetStatus('Error: no runway selected.');
    return;
  }

  // Collect any open textareas
  ['obstructions','weather','dangers','forced','routes','rwyhist','incidents','othernotes'].forEach(function(key) {
    const edit = document.getElementById('rbc-' + key + '-edit');
    if (edit && edit.style.display !== 'none') {
      const keyMap = { obstructions:'obstructions', weather:'weather', dangers:'dangers', forced:'forcedLanding', routes:'routes', rwyhist:'rwyhist', incidents:'incidents', othernotes:'othernotes' };
      s.card[keyMap[key]] = edit.value.trim();
    }
  });

  s.card.mapDistM   = s.mapDistM != null ? s.mapDistM : null;
  s.card.revisedAt  = new Date().toISOString().substring(0,10);
  s.card.revisedBy  = _rbcCurrentUser();

  const saveBtn = document.getElementById('rbc-save-btn');
  if (saveBtn) saveBtn.disabled = true;
  _rbcSetStatus('Saving...');

  if (!(window.google && google.script && google.script.run)) {
    // Offline — persist to AIRCRAFT_DOCS cache as best-effort
    _rbcSetStatus('Offline: changes saved locally only. Sync when online.');
    _rbcSetEditUI(false);
    if (saveBtn) saveBtn.disabled = false;
    return;
  }

  google.script.run
    .withSuccessHandler(function(result) {
      if (saveBtn) saveBtn.disabled = false;
      if (result && result.success) {
        _rbcSetStatus('Saved ✓ — ' + (s.card.revisedAt||''));
        _rbcSetEditUI(false);
        s.dirty = false;
        // Refresh revision footer
        const el1 = document.getElementById('rbc-revision-date');
        const el2 = document.getElementById('rbc-revised-by');
        if (el1) el1.textContent = s.card.revisedAt || '–';
        if (el2) el2.textContent = s.card.revisedBy || '–';
      } else {
        _rbcSetStatus('Save failed: ' + _rbcEsc(String((result && result.error) || 'unknown')));
      }
    })
    .withFailureHandler(function(err) {
      if (saveBtn) saveBtn.disabled = false;
      _rbcSetStatus('Save error: ' + _rbcEsc(String(err || 'unknown')));
    })
    .saveRunwayBriefingCard(s.icao, s.rwyIdent, s.card);
};

function _rbcCurrentUser() {
  // Best-effort: get pilot name from mission context
  try {
    const m = window.currentBriefingMission;
    return String((m && m.pilot) || '').trim() || 'Supervisor';
  } catch (e) { return 'Supervisor'; }
}

function _rbcSetStatus(msg) {
  const el = document.getElementById('rbc-status-msg');
  if (el) el.textContent = msg || '';
}

// ── PAGE 2 — performance history ─────────────────────────────────

function _rbcLoadAndRenderPage2() {
  const s = window._rbcState;

  // Text blocks
  const card = s.card;
  _rbcSetBlock('routes',    card.routes    || '');
  _rbcSetBlock('rwyhist',   card.rwyhist   || '');
  _rbcSetBlock('incidents', card.incidents || '');
  _rbcSetBlock('othernotes',card.othernotes|| '');

  // History
  const loadingEl = document.getElementById('rbc-history-loading');
  const tableEl   = document.getElementById('rbc-history-table');
  if (loadingEl) loadingEl.style.display = 'block';
  if (tableEl)   tableEl.style.display   = 'none';

  if (!(window.google && google.script && google.script.run)) {
    if (loadingEl) loadingEl.textContent = 'Offline — flight history not available.';
    return;
  }

  google.script.run
    .withSuccessHandler(function(result) {
      if (loadingEl) loadingEl.style.display = 'none';
      _rbcRenderHistory(result);
    })
    .withFailureHandler(function() {
      if (loadingEl) loadingEl.textContent = 'Could not load flight history.';
    })
    .getRunwayTakeoffHistory(s.icao, s.rwyIdent, 2);
}

function _rbcRenderHistory(result) {
  const tbody = document.getElementById('rbc-history-tbody');
  const table = document.getElementById('rbc-history-table');
  const loading = document.getElementById('rbc-history-loading');
  if (!tbody) return;

  const records = (result && Array.isArray(result.records)) ? result.records : [];
  if (!records.length) {
    if (loading) { loading.style.display = 'block'; loading.textContent = 'No departure records found for this runway.'; }
    return;
  }

  tbody.innerHTML = records.map(function(r) {
    const actual = r.actualToRollM != null ? Math.round(r.actualToRollM) + ' m' : '–';
    const calc   = r.calcToRollM   != null ? Math.round(r.calcToRollM)   + ' m' : '–';
    const wt     = r.weightKg      != null ? Math.round(r.weightKg) + ' kg' : '–';
    const vrDist = r.vrDistM       != null ? Math.round(r.vrDistM)  + ' m' : '–';
    const mapD   = r.mapDistM      != null ? Math.round(r.mapDistM) + ' m' : '–';
    return '<tr>'
      + '<td>' + _rbcEsc(r.date||'–') + '</td>'
      + '<td>' + _rbcEsc(r.pilot||'–') + '</td>'
      + '<td>' + _rbcEsc(r.acft||'–') + '</td>'
      + '<td>' + _rbcEsc(wt) + '</td>'
      + '<td>' + _rbcEsc(String(r.tempC!=null?r.tempC+'°C':'–')) + '</td>'
      + '<td>' + _rbcEsc(String(r.flaps!=null?r.flaps+'°':'–')) + '</td>'
      + '<td>' + _rbcEsc(r.surface||'–') + '</td>'
      + '<td>' + _rbcEsc(r.wet ? 'WET' : 'DRY') + '</td>'
      + '<td>' + _rbcEsc(calc) + '</td>'
      + '<td><b>' + _rbcEsc(actual) + '</b></td>'
      + '<td>' + _rbcEsc(vrDist) + '</td>'
      + '<td>' + _rbcEsc(mapD) + '</td>'
      + '<td style="font-size:0.7rem;">' + _rbcEsc(r.alternates||'–') + '</td>'
      + '</tr>';
  }).join('');

  if (table)   table.style.display   = '';
  if (loading) loading.style.display = 'none';
}

