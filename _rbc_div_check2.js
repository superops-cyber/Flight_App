const fs = require('fs');
const s = fs.readFileSync('RunwayBriefingCard.html', 'utf8');
const lines = s.split(/\r?\n/);
const stack = [];
for (let i = 0; i < lines.length; i++) {
  const line = lines[i];
  const tags = line.match(/<\/?div\b[^>]*>/g) || [];
  for (const t of tags) {
    if (t.startsWith('</div')) {
      if (!stack.length) {
        console.log('EXTRA close at line', i + 1, line.trim());
      } else {
        stack.pop();
      }
    } else {
      stack.push({ line: i + 1, tag: t });
    }
  }
}
console.log('Unclosed opens:', stack.length);
if (stack.length) {
  console.log('Last 5 opens:');
  stack.slice(-5).forEach(x => console.log(x.line, x.tag));
}
