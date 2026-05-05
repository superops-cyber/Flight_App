
    window.__pilotDiag = { startedAt: Date.now(), mainScriptReady: false };

    (function() {
      function showDiag(msg, level) {
        try {
          var el = document.getElementById('pilot-startup-diag');
          if (!el) return;
          el.style.display = 'block';
          el.textContent = String(msg || 'Startup diagnostic');
          if (level === 'error') {
            el.style.background = '#ffebee';
            el.style.borderColor = '#ef9a9a';
            el.style.color = '#b71c1c';
          } else if (level === 'warn') {
            el.style.background = '#fff8e1';
            el.style.borderColor = '#ffe082';
            el.style.color = '#6d4c41';
          } else {
            el.style.background = '#e8f5e9';
            el.style.borderColor = '#a5d6a7';
            el.style.color = '#1b5e20';
          }
        } catch (_e) {}
      }

      window.addEventListener('error', function(ev) {
        try {
          var src = (ev && ev.filename) ? String(ev.filename).split('/').pop() : 'unknown';
          var ln = (ev && typeof ev.lineno === 'number') ? ev.lineno : 0;
          var col = (ev && typeof ev.colno === 'number') ? ev.colno : 0;
          var msg = (ev && ev.message) ? String(ev.message) : 'unknown error';
          showDiag('Startup error: ' + msg + ' @ ' + src + ':' + ln + ':' + col, 'error');
        } catch (_e2) {}
      });

      window.addEventListener('unhandledrejection', function(ev) {
        try {
          var reason = ev && ev.reason ? ev.reason : null;
          var msg = reason && reason.message ? String(reason.message) : String(reason || 'promise rejected');
          showDiag('Unhandled rejection: ' + msg, 'warn');
        } catch (_e3) {}
      });

      document.addEventListener('DOMContentLoaded', function() {
        try {
          if (!window.__pilotDiag.mainScriptReady) {
            showDiag('Startup warning: main script not ready', 'warn');
          }
        } catch (_e4) {}
      });
    })();
  
