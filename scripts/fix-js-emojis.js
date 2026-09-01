const fs = require('fs');
const path = require('path');

const JS_DIR = path.join(__dirname, '..', 'public', 'js');

const EMOJI_REPLACEMENTS = {
  '📜': '<span class="material-symbols-rounded icon-sm icon-regular icon-outline icon-primary">description</span>',
  '📷': '<span class="material-symbols-rounded icon-sm icon-regular icon-outline icon-primary">camera_alt</span>',
  '🚚': '<span class="material-symbols-rounded icon-sm icon-regular icon-outline icon-primary">local_shipping</span>',
  '✏️': '<span class="material-symbols-rounded icon-sm icon-regular icon-outline icon-secondary">edit</span>',
  '✔': '<span class="material-symbols-rounded icon-sm icon-regular icon-filled icon-success">check</span>',
};

const JS_FILES = ['asset.js', 'bundle.js', 'farm.js', 'stock.js', 'history.js', 'dashboard.js', 'app.js', 'scan.js'];

function processFile(filename) {
  const filepath = path.join(JS_DIR, filename);
  if (!fs.existsSync(filepath)) {
    console.log('Skip (not found):', filename);
    return;
  }
  let content = fs.readFileSync(filepath, 'utf8');
  let changed = false;
  for (const [emoji, replacement] of Object.entries(EMOJI_REPLACEMENTS)) {
    if (content.includes(emoji)) {
      const escaped = emoji.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      content = content.replace(new RegExp(escaped, 'g'), replacement);
      changed = true;
    }
  }
  if (changed) {
    fs.writeFileSync(filepath, content, 'utf8');
    console.log('Fixed:', filename);
  } else {
    console.log('No emojis:', filename);
  }
}

JS_FILES.forEach(processFile);
console.log('\nDone!');