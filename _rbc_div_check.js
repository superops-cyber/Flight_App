const fs = require('fs');
const s = fs.readFileSync('RunwayBriefingCard.html', 'utf8');
const lines = s.split(/\r?\n/);
let depth = 0;
let inOverlay = false;
let overlayDepthStart = null;
for (let i = 0; i < lines.length; i++) {
  const line = lines[i];
  const open = (line.match(/<div\b[^>]*>/g) || []).length;
  const close = (line.match(/<\/div>/g) || []).length;
  if (line.includes('id="rbc-overlay"')) {
    inOverlay = true;
    overlayDepthStart = depth;
  }
  depth += open - close;
  if (inOverlay && depth <= overlayDepthStart) {
    console.log('overlay closes at line', i + 1, 'depth', depth, 'line', line.trim());
    inOverlay = false;
  }
  if (depth < 0) {
    console.log('negative depth at', i + 1);
    break;
  }
}
console.log('final depth', depth);
