function doGet(e) {
 function getQueryParam_(name) {
   var key = String(name || '').toLowerCase();
   if (!key) return '';
   var map = (e && e.parameter) ? e.parameter : {};
   var maps = (e && e.parameters) ? e.parameters : {};
   if (map && map[name] != null) return String(map[name] || '');
   if (map && map[key] != null) return String(map[key] || '');
   if (maps && maps[name] != null) {
     return Array.isArray(maps[name]) ? String((maps[name][0] || '')) : String(maps[name] || '');
   }
   if (maps && maps[key] != null) {
     return Array.isArray(maps[key]) ? String((maps[key][0] || '')) : String(maps[key] || '');
   }

   var qs = String((e && e.queryString) || '');
   if (qs) {
     var parts = qs.split('&');
     for (var i = 0; i < parts.length; i += 1) {
       var pair = parts[i].split('=');
       var k = decodeURIComponent(String(pair[0] || '')).toLowerCase();
       if (k === key) {
         return decodeURIComponent(String(pair.slice(1).join('=') || '').replace(/\+/g, ' '));
       }
     }
   }
   return '';
 }

 const pathInfo = String((e && e.pathInfo) || '').toLowerCase();
 const queryRaw = String((e && e.queryString) || '').toLowerCase();
 const looseHistHint = /view=histcapture|view=historicalcapture|view=historical|histcapture|historicalcapture/.test(queryRaw);
 const viewRaw = getQueryParam_('view');
 const view = String(viewRaw || '').toLowerCase();
 const pathHasHistCapture = /histcapture|historicalcapture|historical/.test(pathInfo);
 const isFlightReport = view === "flightreport" || view === "postflight";
 const isHistCapture = view === "histcapture" || view === "historicalcapture" || view === "historical" || pathHasHistCapture || looseHistHint;
 const isDuty = view === "duty" || view === "dutyapp";
 const pilotParamTrue = String(getQueryParam_('pilot') || '').toLowerCase() === 'true';
 const isPilot = !isDuty && !isFlightReport && !isHistCapture && (view === "dev_pilot" || view === "pilot" || view === "flightdeck" || view === "pilotapp" || pilotParamTrue);
 const diag = String(getQueryParam_('diag') || '').toLowerCase();
 const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'UTC', "yyyy-MM-dd HH:mm:ss") + ' | ' + Utilities.getUuid().slice(0, 8);
 const isDiagProbe = diag === 'probe' || diag === 'includes' || diag === '1';

 try {
   console.log('[doGet] view=%s duty=%s pilot=%s hist=%s path=%s hash=%s', view, isDuty, isPilot, isHistCapture, pathInfo, getQueryParam_('hash'));
 } catch (logErr) {}

 if (diag === 'cva') {
   var diagPayload = getDispatchCvaDiagnostic_({
     reg: String(getQueryParam_('reg') || getQueryParam_('aircraft') || '').trim()
   });
   return ContentService
     .createTextOutput(JSON.stringify(diagPayload, null, 2))
     .setMimeType(ContentService.MimeType.JSON);
 }

 const fileName = isDiagProbe ? 'PortalDiag' : (isFlightReport ? 'View_FlightReport' : (isHistCapture ? 'View_HistoricalCapture' : (isDuty ? 'DutyApp' : (isPilot ? 'PilotApp' : 'Index'))));
 const title = isDiagProbe ? 'Portal Diagnostics' : (isFlightReport ? 'Relatorio Pos-Voo' : (isHistCapture ? 'Captura Historico de Voos' : (isDuty ? 'Duty Time' : (isPilot ? 'Pilot Flight Deck' : 'Flight Ops Portal'))));

 const template = HtmlService.createTemplateFromFile(fileName);
 template.webAppUrl = ScriptApp.getService().getUrl();
 template.buildStamp = stamp;
 template.routeView = view;
 template.routeFlightId = String(getQueryParam_('flightId') || '');
 template.routePathInfo = pathInfo;
 template.routeQueryString = queryRaw;
 template.routeDiagMode = diag;
 template.routeDiagTarget = String(getQueryParam_('part') || '');
 template.routeFileName = fileName;
 template.routeTitle = title;

 return template.evaluate()
   .setTitle(title)
   .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
   .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getPortalProbeHtml(part) {
  var key = String(part || '').trim().toLowerCase();
  var map = {
    home: 'Index',
    dispatch: 'View_Dispatch',
    supervisor: 'View_Supervisor',
    flightfollow: 'View_FlightFollowing',
    tools: 'View_Tools',
    maintenance: 'View_Maintenance'
  };
  var fileName = map[key];
  if (!fileName) {
    return {
      ok: false,
      error: 'Unknown probe target: ' + key,
      fileName: ''
    };
  }

  var template = HtmlService.createTemplateFromFile(fileName);
  template.webAppUrl = ScriptApp.getService().getUrl();
  template.buildStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'UTC', "yyyy-MM-dd HH:mm:ss") + ' | ' + Utilities.getUuid().slice(0, 8);
  template.routeView = key;
  template.routeFlightId = '';
  template.routePathInfo = '';
  template.routeQueryString = 'diag=probe&part=' + key;
  template.routeDiagMode = 'probe';
  template.routeDiagTarget = key;
  template.routeFileName = fileName;
  template.routeTitle = 'Portal Probe: ' + fileName;

  var rendered = template.evaluate().getContent();
  var scriptCount = (rendered.match(/<script\b[^>]*>[\s\S]*?<\/script>/gi) || []).length;
  var previewHtml = String(rendered || '')
    .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<iframe\b[^>]*>[\s\S]*?<\/iframe>/gi, '');

  return {
    ok: true,
    fileName: fileName,
    scriptCount: scriptCount,
    html: previewHtml
  };
}

function doPost(e) {
  var output = { success: false, error: 'Unsupported request' };
  try {
    var raw = String((e && e.postData && e.postData.contents) || '').trim();
    if (!raw) {
      output = { success: false, error: 'Empty POST body' };
    } else {
      var body = JSON.parse(raw);
      var action = String((body && body.action) || '').trim();
      if (action === 'ingestHistoricalFlightDocument') {
        output = ingestHistoricalFlightDocument((body && body.payload) || {});
      } else if (action === 'testHistoricalGeminiConnectivity') {
        output = testHistoricalGeminiConnectivity();
      } else {
        output = { success: false, error: 'Unknown action: ' + action };
      }
    }
  } catch (err) {
    output = { success: false, error: err && err.message ? err.message : String(err) };
  }

  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}
