const fs = require('fs');
const read = (f) => fs.readFileSync(f, 'utf8').split(/\r?\n/);
const root = read('PilotApp.html');
const includes = {
  Styles: read('Styles.html'),
  Tab2_Briefing: read('Tab2_Briefing.html'),
  Tab3_WB: read('Tab3_WB.html'),
  Tab4_Performance: read('Tab4_Performance.html'),
  Tab5_Release: read('Tab5_Release.html'),
  Tab6_Enroute: read('Tab6_Enroute.html'),
  Tab7_Arrival: read('Tab7_Arrival.html'),
  Tab8_Debrief: read('Tab8_Debrief.html'),
  RunwayBriefingCard: read('RunwayBriefingCard.html')
};
let rendered = [];
for (let i = 0; i < root.length; i++) {
  const line = root[i];
  const m = line.match(/include\('([^']+)'\)/);
  if (m && includes[m[1]]) rendered.push(...includes[m[1]]);
  else rendered.push(line);
}
for (let n = 3818; n <= 3828; n++) {
  const line = rendered[n - 1] || '';
  console.log('L' + n + ' len=' + line.length + ' ' + JSON.stringify(line));
  for (let i = 0; i < line.length; i++) {
    const cp = line.charCodeAt(i);
    if (cp < 32 || cp > 126) console.log('  c' + (i + 1) + ' U+' + cp.toString(16).toUpperCase().padStart(4, '0'));
  }
}
