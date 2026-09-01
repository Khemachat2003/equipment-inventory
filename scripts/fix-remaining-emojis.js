const fs = require('fs');
const path = require('path');

const PUBLIC_DIR = path.join(__dirname, '..', 'public');

const EMOJI_REPLACEMENTS = {
  '✕': '<span class="material-symbols-rounded icon-sm icon-regular icon-outline icon-muted">close</span>',
  '❌': '<span class="material-symbols-rounded icon-sm icon-regular icon-filled icon-danger">cancel</span>',
  '✅': '<span class="material-symbols-rounded icon-sm icon-regular icon-filled icon-success">check_circle</span>',
  '⚠️': '<span class="material-symbols-rounded icon-sm icon-regular icon-filled icon-warning">warning</span>',
  '➕': '<span class="material-symbols-rounded icon-sm icon-regular icon-outline icon-primary">add_circle</span>',
  '⚙️': '<span class="material-symbols-rounded icon-sm icon-regular icon-outline icon-secondary">settings</span>',
  '✏️': '<span class="material-symbols-rounded icon-sm icon-regular icon-outline icon-secondary">edit</span>',
};

const files = ['index.html', 'asset.html', 'audit.html', 'backup-view.html', 'qr.html', 'scan.html', 'trace.html'];

files.forEach(f => {
  const filepath = path.join(PUBLIC_DIR, f);
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
    console.log('Fixed:', f);
  }
});
console.log('Done');