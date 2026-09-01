const fs = require('fs');
const path = require('path');

const JS_DIR = path.join(__dirname, '..', 'public', 'js');

// Use SINGLE quotes for HTML strings to avoid conflict with double quotes in class attributes
const EMOJI_REPLACEMENTS = {
  '📦': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>inventory_2</span>',
  '🔧': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>build</span>',
  '⚙️': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-secondary\'>settings</span>',
  '🔩': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-secondary\'>precision_manufacturing</span>',
  '🗃️': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-secondary\'>folder</span>',
  '💻': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>computer</span>',
  '🖥️': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>desktop_mac</span>',
  '⌨️': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-secondary\'>keyboard</span>',
  '🖨️': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-secondary\'>print</span>',
  '📡': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>wifi</span>',
  '📋': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>clipboard_list</span>',
  '📭': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-muted\'>inbox</span>',
  '📍': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>place</span>',
  '📅': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>event</span>',
  '🔄': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>refresh</span>',
  '🏢': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>business</span>',
  '🗑️': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-filled icon-danger\'>delete</span>',
  '🔍': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>search</span>',
  '📤': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>send</span>',
  '🐔': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>egg</span>',
  '🐄': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>pets</span>',
  '🐷': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>pets</span>',
  '🏭': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>factory</span>',
  '🌐': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>public</span>',
  '🌱': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>eco</span>',
  '📊': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>analytics</span>',
  '🔴': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-filled icon-danger\'>circle</span>',
  '📜': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>description</span>',
  '📷': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>camera_alt</span>',
  '🚚': '<span class=\'material-symbols-rounded icon-sm icon-regular icon-outline icon-primary\'>local_shipping</span>',
};

const JS_FILES = ['asset.js', 'bundle.js', 'dashboard.js', 'farm.js', 'history.js', 'stock.js'];

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