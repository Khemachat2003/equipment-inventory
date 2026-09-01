const fs = require('fs');
const path = require('path');

const JS_DIR = path.join(__dirname, '..', 'public', 'js');

const ICON_MAP = {
  '📦': { name: 'inventory_2', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '🔧': { name: 'build', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '⚙️': { name: 'settings', classes: 'icon-sm icon-regular icon-outline icon-secondary' },
  '🔩': { name: 'precision_manufacturing', classes: 'icon-sm icon-regular icon-outline icon-secondary' },
  '🗃️': { name: 'folder', classes: 'icon-sm icon-regular icon-outline icon-secondary' },
  '💻': { name: 'computer', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '🖥️': { name: 'desktop_mac', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '⌨️': { name: 'keyboard', classes: 'icon-sm icon-regular icon-outline icon-secondary' },
  '🖨️': { name: 'print', classes: 'icon-sm icon-regular icon-outline icon-secondary' },
  '📡': { name: 'wifi', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '📋': { name: 'clipboard_list', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '📭': { name: 'inbox', classes: 'icon-sm icon-regular icon-outline icon-muted' },
  '📍': { name: 'place', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '📅': { name: 'event', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '🔄': { name: 'refresh', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '🏢': { name: 'business', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '🗑️': { name: 'delete', classes: 'icon-sm icon-regular icon-filled icon-danger' },
  '🔍': { name: 'search', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '📤': { name: 'send', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '🐔': { name: 'egg', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '🐄': { name: 'pets', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '🐷': { name: 'pets', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '🏭': { name: 'factory', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '🌐': { name: 'public', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '🌱': { name: 'eco', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '📊': { name: 'analytics', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '🔴': { name: 'circle', classes: 'icon-sm icon-regular icon-filled icon-danger' },
  '📜': { name: 'description', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '📷': { name: 'camera_alt', classes: 'icon-sm icon-regular icon-outline icon-primary' },
  '🚚': { name: 'local_shipping', classes: 'icon-sm icon-regular icon-outline icon-primary' },
};

const ICON_HELPER = `// Icon helper for template literals - returns Material Symbols HTML
function ICON(name, classes = 'icon-sm icon-regular icon-outline icon-primary') {
  return '<span class="material-symbols-rounded ' + classes + '">' + name + '</span>';
}
`;

const JS_FILES = ['asset.js', 'bundle.js', 'dashboard.js', 'farm.js', 'history.js', 'stock.js', 'app.js'];

function processFile(filename) {
  const filepath = path.join(JS_DIR, filename);
  if (!fs.existsSync(filepath)) {
    console.log('Skip (not found):', filename);
    return;
  }
  let content = fs.readFileSync(filepath, 'utf8');
  
  // Add ICON helper at the top (after any initial comments)
  if (!content.includes('function ICON(')) {
    const lines = content.split('\n');
    let insertAt = 0;
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (line && !line.startsWith('//') && !line.startsWith('/*') && !line.startsWith('*')) {
        insertAt = i;
        break;
      }
    }
    lines.splice(insertAt, 0, ICON_HELPER);
    content = lines.join('\n');
  }

  // Replace emojis in template literals (backtick strings) with ${ICON(...)} calls
  // This is tricky - we need to find emojis inside backtick strings
  // Simpler approach: replace emoji characters globally, but ONLY in template literal contexts
  // For safety, we'll do a targeted replacement for known patterns
  
  let changed = false;
  for (const [emoji, icon] of Object.entries(ICON_MAP)) {
    if (content.includes(emoji)) {
      const escaped = emoji.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      // Replace emoji with ${ICON('name', 'classes')} - works in template literals
      const replacement = `\${ICON('${icon.name}', '${icon.classes}')}`;
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