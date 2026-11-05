import os
import re
import json
import html

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

    # A static PWA cannot evaluate Apps Script template expressions. Require
    # the active deployment URL explicitly at build time instead of silently
    # generating a bundle with unresolved template syntax.
    web_app_url = os.environ.get('FLIGHT_APP_WEB_APP_URL', '').strip()
    if not web_app_url.startswith('https://script.google.com/') or not web_app_url.endswith('/exec'):
        raise RuntimeError(
            'Set FLIGHT_APP_WEB_APP_URL to the active Apps Script /exec URL before building.'
        )

    escaped_web_app_url = html.escape(web_app_url, quote=True)
    compiled_html, config_replacements = re.subn(
        r'(data-web-app-url=")[^"]*(")',
        r'\1' + escaped_web_app_url + r'\2',
        compiled_html,
        count=1
    )
    if config_replacements != 1:
        raise RuntimeError(
            'Could not locate the app-config data-web-app-url attribute in compiled PilotApp.html.'
        )

    if '<?!=' in compiled_html or '<?= ' in compiled_html:
        raise RuntimeError(
            'Unresolved Apps Script template syntax remains in the static PWA build.'
        )
    
    # Static PWA compatibility bridge. It must call HTTP directly and must
    # never route through apiCall(), which would recurse back into this bridge.
    proxy_script = """
    // PWA HTTP bridge. Expose the google.script.run compatibility facade
    // only after a real Apps Script health check succeeds, so offline
    // branches continue to detect an unavailable backend.
    (function() {
      var pwaRunner = null;
      function pwaHttpCall(action, args) {
        var url = String(
          (document.getElementById('app-config') &&
           document.getElementById('app-config').getAttribute('data-web-app-url')) || ''
        ).trim();
        if (!url) return Promise.reject(new Error('PWA is missing its Apps Script /exec URL.'));
        var callArgs = Array.isArray(args) ? args : [];
        var requestBody = { action: String(action || ''), args: callArgs, payload: callArgs.length ? callArgs[0] : null };
        return fetch(url, {
          method: 'POST',
          credentials: 'include',
          redirect: 'follow',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(requestBody)
        }).then(function(response) {
          return response.text().then(function(raw) {
            var result;
            try { result = JSON.parse(raw); }
            catch (parseError) { throw new Error('Apps Script returned non-JSON (HTTP ' + response.status + '). Check Google sign-in and cross-origin access.'); }
            if (!response.ok) throw new Error('Apps Script HTTP ' + response.status);
            if (!result || result.success !== true) throw new Error(result && result.error ? result.error : 'Apps Script request failed.');
            return result.data;
          });
        }).catch(function(error) {
          window.__PWA_SERVER_READY__ = false;
          removeBridge();
          throw error;
        });
      }
      function makeRunner(successHandler, failureHandler) {
        var runner = {
          withSuccessHandler: function(fn) { return makeRunner(typeof fn === 'function' ? fn : null, failureHandler); },
          withFailureHandler: function(fn) { return makeRunner(successHandler, typeof fn === 'function' ? fn : null); }
        };
        return new Proxy(runner, {
          get: function(target, property) {
            if (property in target) return target[property];
            if (typeof property !== 'string') return target[property];
            return function() {
              var args = Array.prototype.slice.call(arguments);
              pwaHttpCall(property, args).then(function(value) { if (successHandler) successHandler(value); })
                .catch(function(error) { if (failureHandler) failureHandler(error); else console.error('PWA backend call failed: ' + property, error); });
            };
          }
        });
      }
      function installBridge() {
        if (window.google && window.google.script && window.google.script.run && !window.__PWA_HTTP_RUNNER__) return;
        window.google = window.google || {};
        window.google.script = window.google.script || {};
        pwaRunner = makeRunner(null, null);
        window.google.script.run = pwaRunner;
        window.__PWA_HTTP_RUNNER__ = pwaRunner;
      }
      function removeBridge() {
        if (window.__PWA_HTTP_RUNNER__ && window.google && window.google.script && window.google.script.run === window.__PWA_HTTP_RUNNER__) {
          delete window.google.script.run;
        }
        window.__PWA_HTTP_RUNNER__ = null;
        pwaRunner = null;
      }
      function probeBackend() {
        window.__PWA_SERVER_READY__ = false;
        removeBridge();
        pwaHttpCall('pwaHealthCheck', []).then(function() {
          window.__PWA_SERVER_READY__ = true;
          installBridge();
          window.dispatchEvent(new Event('pwa-backend-ready'));
        }).catch(function(error) {
          console.warn('PWA backend health check failed:', error);
        });
      }
      window.__PWA_SERVER_READY__ = false;
      window.addEventListener('online', probeBackend);
      window.addEventListener('offline', function() { window.__PWA_SERVER_READY__ = false; removeBridge(); });
      window.__PWA_HTTP_TRANSPORT__ = pwaHttpCall;
      probeBackend();
    })();
    """

    # The PWA bridge above is the transport; navigator.onLine only indicates
    # that the browser has a network connection, not that Apps Script replied.
    # Existing synchronous call sites use this as "attempt server calls".
    server_check = re.compile(
        r'function isServerAvailable\(\)\s*\{\s*return Boolean\([\s\S]*?\);\s*\}',
        re.MULTILINE
    )
    compiled_html, server_check_replacements = server_check.subn(
        'function isServerAvailable() { return Boolean(window.__PWA_SERVER_READY__); }',
        compiled_html,
        count=1
    )
    if server_check_replacements != 1:
        raise RuntimeError('Could not replace isServerAvailable() for the static PWA build.')

    release_check = re.compile(
        r'function _releaseDispatchOnline_\(\)\s*\{\s*return Boolean\([\s\S]*?\);\s*\}',
        re.MULTILINE
    )
    compiled_html, release_check_replacements = release_check.subn(
        'function _releaseDispatchOnline_() { return Boolean(window.__PWA_SERVER_READY__); }',
        compiled_html,
        count=1
    )
    if release_check_replacements != 1:
        raise RuntimeError('Could not replace _releaseDispatchOnline_() for the static PWA build.')

    # Inject the PWA bridge into the first inline script block.
    if '<script>' in compiled_html:
        compiled_html = compiled_html.replace('<script>', '<script>\n' + proxy_script + '\n', 1)
    else:
        raise RuntimeError('Could not find an inline script block for the PWA transport bridge.')
    
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
