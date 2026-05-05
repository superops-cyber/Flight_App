const fs = require('fs');
const text = fs.readFileSync('Tab3_WB.html', 'utf8');
const target = new Set([0x2018,0x2019,0x201C,0x201D,0x00A0,0x2028,0x2029]);
let line = 1, col = 1;
for (let i = 0; i < text.length; i++) {
  const cp = text.charCodeAt(i);
  if (target.has(cp)) {
    console.log(`L${line}:C${col} U+${cp.toString(16).toUpperCase().padStart(4,'0')}`);
  }
  const ch = text[i];
  if (ch === '\n') { line++; col = 1; }
  else { col++; }
}
