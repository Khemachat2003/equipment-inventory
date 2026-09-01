const fs = require('fs');
const path = require('path');

const PUBLIC_DIR = path.join(__dirname, '..', 'public');

// Emoji to Lucide icon mapping
const EMOJI_MAP = {
  // Navigation & Actions
  '🔄': 'refresh-cw',
  '🔁': 'repeat',
  '🔍': 'search',
  '📦': 'package',
  '🌐': 'globe',
  '📄': 'file-text',
  '👤': 'user',
  '📊': 'bar-chart-2',
  '📈': 'trending-up',
  '📷': 'camera',
  '📋': 'clipboard-list',
  '💾': 'save',
  '💡': 'lightbulb',
  '🌱': 'sprout',
  '🏠': 'home',
  '🏢': 'building',
  '🏭': 'factory',
  '🚚': 'truck',
  '🔧': 'wrench',
  '👁': 'eye',
  '👁️': 'eye',

  // Status circles
  '🟢': 'check-circle-2',
  '🟡': 'alert-circle',
  '🔴': 'x-circle',

  // Farm animals (use generic icons since Lucide doesn't have animals)
  '🐔': 'egg',         // closest for chicken/poultry
  '🐄': 'cow',         // not in lucide, use generic
  '🐷': 'pig',         // not in lucide, use generic

  // UI Elements
  '📥': 'download',
  '📤': 'upload',
  '🗃': 'archive',
  '🎯': 'target',
  '📂': 'folder-open',
  '🗑': 'trash-2',
  '🗑️': 'trash-2',
  '📅': 'calendar',
  '📝': 'edit-2',
  '📭': 'inbox',
  '📭': 'mail',
  '🔒': 'lock',
  '🔐': 'lock',
  '🔑': 'key',
  '🔓': 'unlock',
  '🕐': 'clock',
  '📜': 'scroll',
  '📜': 'file-text',
  '📍': 'map-pin',
  '🖨': 'printer',
  '💬': 'message-square',
  '🔔': 'bell',
  '📱': 'smartphone',
  '📡': 'wifi',
  '📴': 'wifi-off',
  '🚀': 'rocket',
  '🎨': 'palette',
  '🏷': 'tag',
  '🔖': 'bookmark',
  '🌱': 'seedling',
  '💡': 'lightbulb',

  // Additional from trace.html, scan.html, qr.html
  '📥': 'download',
  '📤': 'upload',
  '📂': 'folder-open',
  '🗑': 'trash-2',
  '🗑️': 'trash-2',
};

// Files to process
const HTML_FILES = [
  'index.html',
  'asset.html',
  'audit.html',
  'backup-view.html',
  'qr.html',
  'scan.html',
  'trace.html'
];

function replaceEmojisInContent(content) {
  let replaced = content;

  // Replace each emoji with Lucide icon
  for (const [emoji, iconName] of Object.entries(EMOJI_MAP)) {
    // Escape regex special chars in emoji
    const escapedEmoji = emoji.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    // Replace emoji when it appears as standalone or with text around it
    // We'll replace the emoji character itself with an icon element
    const iconHtml = `<i data-lucide="${iconName}" class="lucide-icon" style="width:1em;height:1em;vertical-align:middle;display:inline-block;"></i>`;
    replaced = replaced.replace(new RegExp(escapedEmoji, 'g'), iconHtml);
  }

  return replaced;
}

function addLucideToHead(content) {
  // Check if lucide already added
  if (content.includes('lucide.min.js')) {
    return content;
  }

  // Add lucide script after chart.js or in head
  const chartScript = '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js" crossorigin="anonymous"></script>';
  const lucideScript = '<script src="https://unpkg.com/lucide@latest/dist/umd/lucide.min.js" crossorigin="anonymous"></script>';

  if (content.includes(chartScript)) {
    return content.replace(chartScript, chartScript + '\n' + lucideScript);
  }

  // Fallback: add before </head>
  return content.replace('</head>', lucideScript + '\n</head>');
}

function addLucideInit(content) {
  // Check if already initialized
  if (content.includes('lucide.createIcons()')) {
    return content;
  }

  // Add initialization before closing </script> at end of body or in existing DOMContentLoaded
  // Find the last </script> before </body>
  const lastScriptEnd = content.lastIndexOf('</script>');
  const bodyEnd = content.indexOf('</body>');

  if (lastScriptEnd > 0 && bodyEnd > lastScriptEnd) {
    const initCode = `
<script>
  // Initialize Lucide icons
  if (typeof lucide !== 'undefined') {
    lucide.createIcons();
  }
</script>
`;
    return content.slice(0, lastScriptEnd + 9) + initCode + content.slice(lastScriptEnd + 9);
  }

  // Fallback: add before </body>
  return content.replace('</body>', `
<script>
  if (typeof lucide !== 'undefined') {
    lucide.createIcons();
  }
</script>
</body>`);
}

function processFile(filename) {
  const filepath = path.join(PUBLIC_DIR, filename);
  let content = fs.readFileSync(filepath, 'utf8');

  // Add Lucide to head
  content = addLucideToHead(content);

  // Replace emojis
  content = replaceEmojisInContent(content);

  // Add Lucide init
  content = addLucideInit(content);

  fs.writeFileSync(filepath, content, 'utf8');
  console.log(`✅ Processed ${filename}`);
}

// Run
HTML_FILES.forEach(processFile);
console.log('\n🎉 All files processed!');