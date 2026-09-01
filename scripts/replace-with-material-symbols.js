const fs = require('fs');
const path = require('path');

const PUBLIC_DIR = path.join(__dirname, '..', 'public');

// Lucide icon name -> Material Symbols name mapping
const LUCIDE_TO_MATERIAL = {
  'refresh-cw': 'refresh',
  'repeat': 'replay',
  'search': 'search',
  'package': 'inventory_2',
  'globe': 'public',
  'file-text': 'description',
  'user': 'person',
  'bar-chart-2': 'analytics',
  'trending-up': 'trending_up',
  'camera': 'camera_alt',
  'clipboard-list': 'assignment',
  'save': 'save',
  'lightbulb': 'lightbulb',
  'sprout': 'eco',
  'home': 'home',
  'factory': 'factory',
  'truck': 'local_shipping',
  'wrench': 'build',
  'eye': 'visibility',
  'check-circle-2': 'check_circle',
  'alert-circle': 'warning',
  'x-circle': 'cancel',
  'feather': 'pets',           // was egg -> chicken
  'heart': 'pets',             // was cow -> generic animal
  'bone': 'pets',              // was pig -> generic animal
  'download': 'download',
  'upload': 'upload',
  'archive': 'archive',
  'target': 'track_changes',
  'folder-open': 'folder_open',
  'trash-2': 'delete',
  'calendar': 'event',
  'edit-2': 'edit',
  'inbox': 'inbox',
  'mail': 'mail',
  'lock': 'lock',
  'key': 'vpn_key',
  'unlock': 'lock_open',
  'clock': 'schedule',
  'scroll': 'article',
  'map-pin': 'place',
  'printer': 'print',
  'message-square': 'chat',
  'bell': 'notifications',
  'smartphone': 'phone_android',
  'wifi': 'wifi',
  'wifi-off': 'wifi_off',
  'rocket': 'rocket_launch',
  'palette': 'palette',
  'tag': 'label',
  'bookmark': 'bookmark',
  'seedling': 'eco',
};

// HTML files to process
const HTML_FILES = [
  'index.html',
  'asset.html',
  'audit.html',
  'backup-view.html',
  'qr.html',
  'scan.html',
  'trace.html'
];

function replaceLucideWithMaterial(content) {
  let replaced = content;

  // 1. Replace Lucide CDN with Material Symbols font
  const lucideScript = '<script src="https://unpkg.com/lucide@latest/dist/umd/lucide.min.js" crossorigin="anonymous"></script>';
  const materialLink = '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-25..200" crossorigin="anonymous" />';

  if (replaced.includes(lucideScript)) {
    replaced = replaced.replace(lucideScript, materialLink);
  }

  // 2. Replace <i data-lucide="name" ...> with <span class="material-symbols-rounded" style="...">name</span>
  // Pattern: <i data-lucide="icon-name" class="lucide-icon" style="width:1em;height:1em;vertical-align:middle;display:inline-block;"></i>
  const iconRegex = /<i\s+data-lucide="([^"]+)"\s+class="lucide-icon"[^>]*><\/i>/g;

  replaced = replaced.replace(iconRegex, (match, iconName) => {
    const materialName = LUCIDE_TO_MATERIAL[iconName] || iconName;
    // Default style: weight 400, fill 0 (outline), grade 0, optical size 24
    const style = "font-variation-settings:'wght' 400,'FILL' 0,'GRAD' 0,'opsz' 24;font-size:1.2em;vertical-align:middle;display:inline-block;";
    return `<span class="material-symbols-rounded" style="${style}">${materialName}</span>`;
  });

  return replaced;
}

function replaceLucideInit(content) {
  // Remove lucide.createIcons() calls since Material Symbols auto-renders
  return content.replace(/if\s*\(\s*typeof\s+lucide\s*!==\s*['"]undefined['"]\s*\)\s*\{\s*lucide\.createIcons\(\);\s*\}/g, '');
}

function processFile(filename) {
  const filepath = path.join(PUBLIC_DIR, filename);
  let content = fs.readFileSync(filepath, 'utf8');

  // Replace Lucide with Material Symbols
  content = replaceLucideWithMaterial(content);

  // Remove lucide.init calls
  content = replaceLucideInit(content);

  fs.writeFileSync(filepath, content, 'utf8');
  console.log(`✅ Converted ${filename} to Material Symbols`);
}

// Run
HTML_FILES.forEach(processFile);
console.log('\n🎉 All files converted to Material Symbols!');