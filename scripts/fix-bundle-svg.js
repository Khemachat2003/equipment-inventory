const fs = require('fs');
const path = require('path');

const filepath = path.join(__dirname, '..', 'public', 'js', 'bundle.js');
let content = fs.readFileSync(filepath, 'utf8');

// Fix the remove button SVG
content = content.replace(
  /<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2\.5" stroke-linecap="round"><line x1="18" y1="6" x2="6" y2="18"\/><line x1="6" y1="6" x2="18" y2="18"\/><\/svg>/g,
  '<span class="material-symbols-rounded icon-sm icon-regular icon-filled icon-danger">delete</span>'
);

fs.writeFileSync(filepath, content, 'utf8');
console.log('✅ Fixed remove button SVG');