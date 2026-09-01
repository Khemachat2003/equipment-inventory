const fs = require('fs');
const path = require('path');

const PUBLIC_DIR = path.join(__dirname, '..', 'public');

// Map old inline styles to new utility classes
function upgradeInlineIcons(content) {
  let replaced = content;

  // Pattern: <span class="material-symbols-rounded" style="font-variation-settings:'wght' 400,'FILL' 0,'GRAD' 0,'opsz' 24;font-size:1.2em;vertical-align:middle;display:inline-block;">icon-name</span>
  const inlineStyleRegex = /<span class="material-symbols-rounded" style="font-variation-settings:'wght' 400,'FILL' 0,'GRAD' 0,'opsz' 24;font-size:1\.2em;vertical-align:middle;display:inline-block;">([^<]+)<\/span>/g;

  replaced = replaced.replace(inlineStyleRegex, (match, iconName) => {
    // Determine appropriate class based on context/icon
    let className = 'material-symbols-rounded icon-md icon-regular icon-outline';
    
    // Add color based on common icon names
    if (['refresh', 'search', 'inventory_2', 'package', 'description', 'person', 'analytics', 'trending_up', 'camera_alt', 'assignment', 'save', 'lightbulb', 'eco', 'home', 'factory', 'local_shipping', 'build', 'visibility', 'check_circle', 'warning', 'cancel', 'pets', 'download', 'upload', 'archive', 'track_changes', 'folder_open', 'delete', 'event', 'edit', 'inbox', 'mail', 'lock', 'vpn_key', 'lock_open', 'schedule', 'article', 'place', 'print', 'chat', 'notifications', 'phone_android', 'wifi', 'wifi_off', 'rocket_launch', 'palette', 'label', 'bookmark'].includes(iconName)) {
      // Default to primary color for action icons
      className += ' icon-primary';
    }
    
    // Special cases
    if (['check_circle'].includes(iconName)) className = className.replace('icon-primary', 'icon-success').replace('icon-outline', 'icon-filled');
    if (['warning'].includes(iconName)) className = className.replace('icon-primary', 'icon-warning').replace('icon-outline', 'icon-filled');
    if (['cancel'].includes(iconName)) className = className.replace('icon-primary', 'icon-danger').replace('icon-outline', 'icon-filled');
    if (['visibility', 'search'].includes(iconName)) className += ' icon-secondary';
    
    return `<span class="${className}">${iconName}</span>`;
  });

  return replaced;
}

// Also need to add the font link and icon system CSS to other HTML files
const ICON_SYSTEM_CSS = `
/* ═══════════════════════════════════════════════════════════════
   ICON DESIGN SYSTEM — Premium Material Symbols
   Size × Weight × Fill × Grade × Color × State
═══════════════════════════════════════════════════════════════ */

/* Base icon element */
.material-symbols-rounded {
  font-family: 'Material Symbols Rounded';
  font-style: normal;
  font-variation-settings:
    'wght' var(--icon-w-regular, 400),
    'FILL' var(--icon-fill-none, 0),
    'GRAD' var(--icon-grade-normal, 0),
    'opsz' var(--icon-opsz-md, 24);
  display: inline-flex;
  align-items: center;
  justify-content: center;
  flex-shrink: 0;
  line-height: 1;
  vertical-align: middle;
  transition:
    font-variation-settings var(--icon-transition),
    color var(--icon-transition),
    transform var(--icon-transition);
}

/* ─── Size Modifiers ─── */
.icon-xs { font-size: var(--icon-xs); --icon-opsz-md: var(--icon-opsz-sm); }
.icon-sm { font-size: var(--icon-sm); --icon-opsz-md: var(--icon-opsz-sm); }
.icon-md { font-size: var(--icon-md); --icon-opsz-md: var(--icon-opsz-md); }
.icon-lg { font-size: var(--icon-lg); --icon-opsz-md: var(--icon-opsz-lg); }
.icon-xl { font-size: var(--icon-xl); --icon-opsz-md: var(--icon-opsz-lg); }
.icon-2xl { font-size: var(--icon-2xl); --icon-opsz-md: var(--icon-opsz-xl); }

/* ─── Weight Modifiers ─── */
.icon-light     { --icon-w-regular: var(--icon-w-light); }
.icon-regular   { --icon-w-regular: var(--icon-w-regular); }
.icon-medium    { --icon-w-regular: var(--icon-w-medium); }
.icon-semibold  { --icon-w-regular: var(--icon-w-semibold); }
.icon-bold      { --icon-w-regular: var(--icon-w-bold); }

/* ─── Fill Modifiers (semantic) ─── */
.icon-outline   { --icon-fill-none: var(--icon-fill-none); }
.icon-hover     { --icon-fill-none: var(--icon-fill-subtle); }
.icon-active    { --icon-fill-none: var(--icon-fill-half); }
.icon-filled    { --icon-fill-none: var(--icon-fill-full); }

/* ─── Grade Modifiers ─── */
.icon-grade-normal { --icon-grade-normal: var(--icon-grade-normal); }
.icon-grade-dark   { --icon-grade-normal: var(--icon-grade-dark); }
.icon-grade-light  { --icon-grade-normal: var(--icon-grade-light); }

/* ─── Semantic Color Variants ─── */
.icon-primary   { color: var(--blue); }
.icon-secondary { color: var(--tsub); }
.icon-muted     { color: var(--tmuted); }
.icon-success   { color: var(--emerald); }
.icon-warning   { color: var(--amber); }
.icon-danger    { color: var(--red); }
.icon-purple    { color: var(--purple); }
.icon-white     { color: #fff; }
.icon-inherit   { color: inherit; }

/* ─── Composite Utility Classes ─── */
.nav-icon {
  --icon-w-regular: var(--icon-w-medium);
  --icon-fill-none: var(--icon-fill-none);
  --icon-grade-normal: var(--icon-grade-normal);
  --icon-opsz-md: var(--icon-opsz-md);
  font-size: var(--icon-md);
  color: var(--tsub);
  transition: color var(--icon-transition), font-variation-settings var(--icon-transition);
}
.nav-icon:hover,
.nav-icon.active {
  color: var(--blue);
  --icon-fill-none: var(--icon-fill-full);
  --icon-w-regular: var(--icon-w-semibold);
}

.btn-icon {
  --icon-w-regular: var(--icon-w-semibold);
  --icon-fill-none: var(--icon-fill-none);
  font-size: var(--icon-sm);
}
.btn-primary .btn-icon    { color: var(--blue); }
.btn-primary:hover .btn-icon    { --icon-fill-none: var(--icon-fill-subtle); }
.btn-primary:active .btn-icon   { --icon-fill-none: var(--icon-fill-half); }
.btn-secondary .btn-icon  { color: var(--tsub); }
.btn-outline .btn-icon    { color: var(--blue); }
.btn-ghost .btn-icon      { color: var(--tsub); }
.btn-ghost:hover .btn-icon    { color: var(--blue); --icon-fill-none: var(--icon-fill-subtle); }

.status-icon {
  --icon-w-regular: var(--icon-w-semibold);
  --icon-fill-none: var(--icon-fill-full);
  font-size: var(--icon-xs);
}
.status-success .status-icon { color: var(--emerald); }
.status-warning .status-icon { color: var(--amber); }
.status-danger .status-icon  { color: var(--red); }
.status-info .status-icon    { color: var(--blue); }

.card-icon {
  --icon-w-regular: var(--icon-w-regular);
  --icon-fill-none: var(--icon-fill-none);
  font-size: var(--icon-xl);
  color: var(--blue);
}
.card:hover .card-icon {
  --icon-fill-none: var(--icon-fill-subtle);
  --icon-w-regular: var(--icon-w-medium);
}

.fab-icon {
  --icon-w-regular: var(--icon-w-bold);
  --icon-fill-none: var(--icon-fill-full);
  font-size: var(--icon-lg);
  color: #fff;
}

.table-action-icon {
  --icon-w-regular: var(--icon-w-medium);
  --icon-fill-none: var(--icon-fill-none);
  font-size: var(--icon-sm);
  color: var(--tmuted);
  transition: color var(--icon-transition), font-variation-settings var(--icon-transition), transform var(--icon-transition);
}
.table-action-icon:hover {
  color: var(--blue);
  --icon-fill-none: var(--icon-fill-subtle);
  transform: scale(1.1);
}

.empty-icon {
  --icon-w-regular: var(--icon-w-light);
  --icon-fill-none: var(--icon-fill-none);
  font-size: var(--icon-2xl);
  color: var(--g300);
}

.header-icon {
  --icon-w-regular: var(--icon-w-semibold);
  --icon-fill-none: var(--icon-fill-full);
  font-size: var(--icon-lg);
  color: var(--ink);
}

.material-symbols-rounded:not(.no-transition) {
  transition: font-variation-settings 150ms ease, color 150ms ease, transform 150ms ease;
}

.material-symbols-rounded:focus-visible {
  outline: 2px solid var(--blue);
  outline-offset: 2px;
  border-radius: 4px;
}

@media (prefers-reduced-motion: reduce) {
  .material-symbols-rounded {
    transition: none;
  }
}
`;

const MATERIAL_FONT_LINK = '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-25..200" crossorigin="anonymous" />';

const HTML_FILES = [
  'asset.html',
  'audit.html',
  'backup-view.html',
  'qr.html',
  'scan.html',
  'trace.html'
];

function processFile(filename) {
  const filepath = path.join(PUBLIC_DIR, filename);
  let content = fs.readFileSync(filepath, 'utf8');

  // 1. Add Material Symbols font link if not present
  if (!content.includes('Material+Symbols+Rounded')) {
    // Add after chart.js script or in head
    const chartScript = '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js" crossorigin="anonymous"></script>';
    if (content.includes(chartScript)) {
      content = content.replace(chartScript, chartScript + '\n' + MATERIAL_FONT_LINK);
    } else {
      content = content.replace('</head>', MATERIAL_FONT_LINK + '\n</head>');
    }
  }

  // 2. Add Icon Design System CSS before closing </style> in head
  if (!content.includes('ICON DESIGN SYSTEM')) {
    content = content.replace('</style>', ICON_SYSTEM_CSS + '\n</style>');
  }

  // 3. Upgrade inline icon styles
  content = upgradeInlineIcons(content);

  fs.writeFileSync(filepath, content, 'utf8');
  console.log(`✅ Upgraded ${filename}`);
}

// Run on all files except index.html (already done)
HTML_FILES.forEach(processFile);
console.log('\n🎉 All files upgraded with Icon Design System!');