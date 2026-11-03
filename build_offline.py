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
    if (typeof google === 'undefined') {
      console.log("PWA Mode: Initializing google.script.run Proxy interceptor.");
      window.google = {
        script: {
          run: (function() {
            class RunProxy {
              constructor() {
                this._success = null;
                this._failure = null;
                
                return new Proxy(this, {
                  get: (target, prop) => {
                    if (prop === 'withSuccessHandler') {
                      return (fn) => { target._success = fn; return target; };
                    }
                    if (prop === 'withFailureHandler') {
                      return (fn) => { target._failure = fn; return target; };
                    }
                    
                    // Route any other function call directly to fetch apiCall()
                    return async (...args) => {
                      try {
                        const payload = args.length > 0 ? args[0] : null;
                        const result = await apiCall(prop, payload);
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
          })()
        }
      };
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
