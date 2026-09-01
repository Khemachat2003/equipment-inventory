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

// This function replaces emojis ONLY inside template literals (backtick strings)
function replaceEmojisInTemplateLiterals(content) {
  let result = '';
  let i = 0;
  let inTemplate = false;
  let templateStart = -1;
  let templateContent = '';
  let inString = false;
  let stringChar = '';
  let escapeNext = false;

  while (i < content.length) {
    const ch = content[i];
    const next = content[i + 1];

    if (!inTemplate && !inString) {
      // Check for template literal start
      if (ch === '`') {
        inTemplate = true;
        templateStart = result.length;
        templateContent = '';
      } else if (ch === '\'' || ch === '"') {
        inString = true;
        stringChar = ch;
      }
      result += ch;
    } else if (inTemplate) {
      // Inside template literal
      templateContent += ch;
      result += ch;
      if (ch === '`') {
        // Template literal ends
        // Now process the template content for emoji replacement
        let processedContent = templateContent;
        for (const [emoji, icon] of Object.entries(ICON_MAP)) {
          if (processedContent.includes(emoji)) {
            const escaped = emoji.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const replacement = `\${ICON('${icon.name}', '${icon.classes}')}`;
            processedContent = processedContent.replace(new RegExp(escaped, 'g'), replacement);
          }
        }
        // Replace the template content in result
        result = result.slice(0, templateStart) + processedContent;
        inTemplate = false;
        templateContent = '';
      } else if (ch === '\\' && !escapeNext) {
        escapeNext = true;
      } else {
        escapeNext = false;
      }
    } else if (inString) {
      // Inside single/double quoted string - don't replace
      result += ch;
      if (ch === stringChar && !escapeNext) {
        inString = false;
        stringChar = '';
      }
      escapeNext = ch === '\\' && !escapeNext;
    }
    i++;
  }
  return result;
}

function processFile(filename) {
  const filepath = path.join(JS_DIR, filename);
  if (!fs.existsSync(filepath)) {
    console.log('Skip (not found):', filename);
    return;
  }
  let content = fs.readFileSync(filepath, 'utf8');
  
  // Add ICON helper at the top
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

  // Replace emojis only in template literals
  const newContent = replaceEmojisInTemplateLiterals(content);
  
  if (newContent !== content) {
    fs.writeFileSync(filepath, newContent, 'utf8');
    console.log('Fixed:', filename);
  } else {
    console.log('No emojis in template literals:', filename);
  }
}

JS_FILES.forEach(processFile);
console.log('\nDone!');