import os
import re
import json

def compile_gas_includes(src_filename):
    with open(src_filename, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
        
    # Pattern to find <?!= include('FileName'); ?>
    include_regex = re.compile(r'<\?!=?\s*include\([\'"]([a-zA-Z0-9_$]+)[\'"]\);\s*\?>')
    
    def replace_include(match):
        filename = match.group(1) + '.html'
        if os.path.exists(filename):
            print(f"Embedding {filename}...")
            with open(filename, 'r', encoding='utf-8', errors='ignore') as inc_file:
                # Recursively resolve any includes inside the included file
                inc_content = inc_file.read()
                return compile_gas_includes_str(inc_content)
        else:
            print(f"Warning: Include file {filename} not found.")
            return f"<!-- Include {filename} not found -->"
            
    return include_regex.sub(replace_include, content)

def compile_gas_includes_str(content):
    include_regex = re.compile(r'<\?!=?\s*include\([\'"]([a-zA-Z0-9_$]+)[\'"]\);\s*\?>')
    
    def replace_include(match):
        filename = match.group(1) + '.html'
        if os.path.exists(filename):
            with open(filename, 'r', encoding='utf-8', errors='ignore') as inc_file:
                return compile_gas_includes_str(inc_file.read())
        else:
            return f"<!-- Include {filename} not found -->"
            
    return include_regex.sub(replace_include, content)

def build():
    build_dir = 'offline_build'
    os.makedirs(build_dir, exist_ok=True)
    
    print("Compiling PilotApp.html...")
    compiled_html = compile_gas_includes('PilotApp.html')
    
    # Inject Mock google.script.run Proxy right after <script> in the compiled HTML
    proxy_script = """
    // PWA OFFLINE MOCK FOR GOOGLE.SCRIPT.RUN
    function buildDemoModeProxy() {
      const demoMission = {
        id: 'DEMO-2026-11-06-001',
        date: '2026-11-06',
        acft: 'PP-ACF',
        pilot: 'Ana Silva',
        copilot: 'R. Oliveira',
        type: 'DEMO MISSION',
        notes: 'Demo mission for isolated offline pilot workflow.',
        meta: {
          acft: 'PP-ACF',
          pilot: 'Ana Silva',
          copilot: 'R. Oliveira',
          notes: 'Demo mission for isolated offline pilot workflow.',
          training: { code: 'VFR-AB', route: 'SBGR,SBRJ' }
        },
        legs: [{
          flightLegId: 'DEMO-2026-11-06-001-01',
          from: 'SBGR',
          to: 'SBRJ',
          route: 'SBGR-SBRJ',
          waypoints: ['SBGR', 'SBSP', 'SBRJ'],
          time: 2.9,
          groundTime: 0.6,
          dist: 221.5,
          fuel: 680,
          takeoffFuel: 1180,
          landingFuel: 420,
          limit: 3200,
          limitType: 'MTOW',
          pax: []
        }]
      };

      const demoEnvelope = {
        aircraft: 'PP-ACF',
        aircraftType: 'PC12',
        cachedAt: new Date().toISOString(),
        envelopeData: [
          { AIRCRAFT_TYPE: 'PC12', POINT_SEQUENCE: 1, CG_Arm_X: 80.5, Weight_Y: 2100 },
          { AIRCRAFT_TYPE: 'PC12', POINT_SEQUENCE: 2, CG_Arm_X: 92.0, Weight_Y: 2850 },
          { AIRCRAFT_TYPE: 'PC12', POINT_SEQUENCE: 3, CG_Arm_X: 101.5, Weight_Y: 3300 },
          { AIRCRAFT_TYPE: 'PC12', POINT_SEQUENCE: 4, CG_Arm_X: 109.0, Weight_Y: 2950 }
        ]
      };

      const demoStartup = {
        nextId: 'DEMO-2026-11-06-001',
        pilots: [{ name: 'Ana Silva', weight: 72, Medical_Expiry: '2026-12-12', MNTE_Validity: '2026-12-12', MNAF_Validity: '2026-12-12' }],
        pilotDestinationChecks: {},
        funds: [],
        rates: ['1/5', '2/5', '1/2', '3/5', '4/5', '1/1'],
        fuelCaches: [{ icao: 'SBGR', location: 'Guarulhos', qty: 2400, type: 'JET A1' }],
        syllabus: [{ code: 'VFR-AB', description: 'Demo VFR route validation', aircraftType: 'PC12', hours: 2.5, plannedLandings: 1, fuel: 120, requiredBallast: 0, route: 'SBGR-SBSP-SBRJ', runwayCheckout: 'RUNWAY 09L', lessonPlanJson: JSON.stringify({ stops: ['SBGR', 'SBSP', 'SBRJ'] }), plannedStopsJson: JSON.stringify(['SBGR', 'SBSP', 'SBRJ']), maneuversJson: JSON.stringify({}) }],
        waypoints: [{ wp_id: 'SBGR', lat: '-23.4356', lon: '-46.4731', type: 'APT' }, { wp_id: 'SBSP', lat: '-23.5000', lon: '-46.6167', type: 'APT' }, { wp_id: 'SBRJ', lat: '-22.8100', lon: '-43.2506', type: 'APT' }],
        routes: [{ rowNumber: 2, origin: 'SBGR', destination: 'SBRJ', waypoint_list: 'SBGR,SBSP,SBRJ' }],
        airports: [{ icao: 'SBGR', nome: 'Guarulhos Intl', lat: -23.4356, lon: -46.4731, fuelAvailable: 'JET A1', runwayIdent: '09L/27R', runwayLength: '3600', runwayWidth: '45', runwaySurfaceActual: 'asphalt' }, { icao: 'SBSP', nome: 'Congonhas', lat: -23.5000, lon: -46.6167, fuelAvailable: 'JET A1', runwayIdent: '15/33', runwayLength: '1800', runwayWidth: '35', runwaySurfaceActual: 'asphalt' }, { icao: 'SBRJ', nome: 'Rio Santos Dumont', lat: -22.8100, lon: -43.2506, fuelAvailable: 'JET A1', runwayIdent: '02R/20L', runwayLength: '2200', runwayWidth: '45', runwaySurfaceActual: 'asphalt' }],
        aircraft: [{ reg: 'PP-ACF', aircraftType: 'PC12', model: 'Pilatus PC-12 NG', mtow: 4750, maxFuel: 1160, burn: 215, maxPax: 6, License_Required: 'MNTE', currentTach: 1832.4, hoursToInspection: 94.1, currentHours: 1832.4, status: 'DEMO' }],
        passengers: [],
        passengerConsentConfig: { requireConsent: false },
        missions: [demoMission]
      };

      const persist = (method, payload) => {
        try {
          const key = 'mba_demo_persisted_' + method;
          localStorage.setItem(key, JSON.stringify({ savedAt: Date.now(), value: payload }));
        } catch (err) {}
      };

      return {
        call: function(method, args) {
          const a = Array.isArray(args) ? args : [];
          if (method === 'getPilotStartupData' || method === 'getDropdownData') return demoStartup;
          if (method === 'getScheduledMissions') return [demoMission];
          if (method === 'getMissionById') return a[0] === demoMission.id ? demoMission : null;
          if (method === 'getWbEnvelopeByAircraft') return demoEnvelope;
          if (method === 'saveMission' || method === 'saveMissionToLog' || method === 'saveWBToLog') {
            persist(method, a[0] || {});
            return { success: true, demo: true, storedLocally: true, missionId: demoMission.id };
          }
          return { success: true, demo: true, action: method };
        }
      };
    }
    
    if (typeof google === 'undefined') {
      console.log('PWA Mode: Initializing google.script.run Proxy interceptor.');
      window.google = { script: { run: (function() {
        const demoProxy = buildDemoModeProxy();
        class RunProxy {
          constructor() {
            this._success = null;
            this._failure = null;
            return new Proxy(this, {
              get: (target, prop) => {
                if (prop === 'withSuccessHandler') return (fn) => { target._success = fn; return target; };
                if (prop === 'withFailureHandler') return (fn) => { target._failure = fn; return target; };
                if (prop === 'then' || prop === 'catch' || prop === 'finally') return undefined;
                return (...args) => {
                  try {
                    const result = demoProxy.call(prop, args);
                    if (target._success) target._success(result);
                  } catch (err) {
                    if (target._failure) target._failure(err);
                  }
                };
              }
            });
          }
        }
        return new RunProxy();
      })() } };
    }
    
    if (window.location && /(?:[?&])demo=1(?:[&]|$)/.test(window.location.search || '')) {
      window.__DEMO_MODE__ = true;
      const demoIndicator = document.createElement('div');
      demoIndicator.textContent = 'DEMO MODE';
      demoIndicator.style.position = 'fixed';
      demoIndicator.style.top = '12px';
      demoIndicator.style.right = '12px';
      demoIndicator.style.zIndex = '2147483647';
      demoIndicator.style.background = '#ffc107';
      demoIndicator.style.color = '#111';
      demoIndicator.style.fontSize = '11px';
      demoIndicator.style.fontWeight = '900';
      demoIndicator.style.padding = '6px 10px';
      demoIndicator.style.borderRadius = '999px';
      demoIndicator.style.letterSpacing = '0.12em';
      demoIndicator.style.boxShadow = '0 4px 12px rgba(0,0,0,0.18)';
      document.body && document.body.appendChild(demoIndicator);
    }
    """
    
    # Locate <script> tag and inject
    if '<script>' in compiled_html:
        compiled_html = compiled_html.replace('<script>', f'<script>\n{proxy_script}\n', 1)
    
    with open(f"{build_dir}/index.html", 'w', encoding='utf-8') as f:
        f.write(compiled_html)
    print(f"Created {build_dir}/index.html")
    
    # 2. Write manifest.json
    manifest = {
        "short_name": "FlightDeck",
        "name": "Pilot Flight Deck Offline",
        "icons": [
            {
                "src": "https://fonts.gstatic.com/s/i/materialicons/flight_takeoff/v6/24px.svg",
                "type": "image/svg+xml",
                "sizes": "192x192 512x512"
            }
        ],
        "start_url": "index.html",
        "background_color": "#0b5394",
        "theme_color": "#0b5394",
        "display": "standalone",
        "orientation": "portrait"
    }
    with open(f"{build_dir}/manifest.json", 'w', encoding='utf-8') as f:
        json.dump(manifest, f, indent=2)
    print(f"Created {build_dir}/manifest.json")
    
    # 3. Write service-worker.js
    sw_code = """
const CACHE_NAME = 'flightdeck-cache-v1';
const ASSETS = [
  'index.html',
  'manifest.json',
  'https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/css/materialize.min.css',
  'https://fonts.googleapis.com/icon?family=Material+Icons',
  'https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/js/materialize.min.js'
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', (e) => {
  // Let the browser handle standard POST fetch requests directly (GAS backend)
  if (e.request.method === 'POST') {
    return;
  }
  
  e.respondWith(
    caches.match(e.request).then((cachedResponse) => {
      return cachedDistanceResponse || fetch(e.request);
    }).catch(() => fetch(e.request))
  );
});
"""
    # Quick fix in the fetch handler for typo
    sw_code = sw_code.replace("cachedDistanceResponse", "cachedResponse")
    
    with open(f"{build_dir}/service-worker.js", 'w', encoding='utf-8') as f:
        f.write(sw_code.strip() + "\n")
    print(f"Created {build_dir}/service-worker.js")
    
    print("\nOffline build complete! The 'offline_build' folder is ready to host on Vercel, Netlify, or GitHub Pages.")

if __name__ == "__main__":
    build()
