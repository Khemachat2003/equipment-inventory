// emoji-walker.js - Runtime emoji to Material Symbols replacement

(function() {
  const EMOJI_ICON_MAP = {
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
    '⚙️': { name: 'settings', classes: 'icon-sm icon-regular icon-outline icon-secondary' },
    '🗑️': { name: 'delete', classes: 'icon-sm icon-regular icon-filled icon-danger' },
    '📭': { name: 'inbox', classes: 'icon-sm icon-regular icon-outline icon-muted' },
    '📍': { name: 'place', classes: 'icon-sm icon-regular icon-outline icon-primary' },
    '📅': { name: 'event', classes: 'icon-sm icon-regular icon-outline icon-primary' },
    '⚙️': { name: 'settings', classes: 'icon-sm icon-regular icon-outline icon-secondary' },
    '✏️': { name: 'edit', classes: 'icon-sm icon-regular icon-outline icon-secondary' },
    '✔': { name: 'check', classes: 'icon-sm icon-regular icon-filled icon-success' },
    '🔴': { name: 'circle', classes: 'icon-sm icon-regular icon-filled icon-danger' },
  };

  function createIconSpan(icon) {
    const span = document.createElement('span');
    span.className = 'material-symbols-rounded ' + icon.classes;
    span.textContent = icon.name;
    span.style.cssText = "font-variation-settings:'wght' 400,'FILL' 0,'GRAD' 0,'opsz' 24;font-size:1.2em;vertical-align:middle;display:inline-block;";
    return span;
  }

  function isValidNode(node) {
    return node && node.nodeType && (node.nodeType === Node.ELEMENT_NODE || node.nodeType === Node.TEXT_NODE);
  }

  function replaceEmojiInTextNode(textNode) {
    if (!isValidNode(textNode) || textNode.nodeType !== Node.TEXT_NODE) return false;
    
    let text = textNode.textContent;
    let replaced = false;
    
    for (const [emoji, icon] of Object.entries(EMOJI_ICON_MAP)) {
      if (text.includes(emoji)) {
        const parent = textNode.parentNode;
        if (!parent) return false;
        
        if (text === emoji) {
          parent.replaceChild(createIconSpan(icon), textNode);
        } else {
          const parts = text.split(emoji);
          const frag = document.createDocumentFragment();
          parts.forEach((part, i) => {
            if (part) frag.appendChild(document.createTextNode(part));
            if (i < parts.length - 1) frag.appendChild(createIconSpan(icon).cloneNode(true));
          });
          parent.replaceChild(frag, textNode);
        }
        replaced = true;
        break;
      }
    }
    return replaced;
  }

  function replaceEmojiInTextNodes(root) {
    if (!isValidNode(root)) return;
    
    try {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null, false);
      const textNodes = [];
      let node;
      while (node = walker.nextNode()) {
        if (isValidNode(node)) textNodes.push(node);
      }
      textNodes.forEach(replaceEmojiInTextNode);
    } catch (e) {
      // Silently ignore invalid roots
    }
  }

  function replaceEmojiInElement(element) {
    if (!isValidNode(element) || element.nodeType !== Node.ELEMENT_NODE) return;
    
    // Replace in text nodes within element
    replaceEmojiInTextNodes(element);
    
    // Also check element's own textContent for emojis
    for (const [emoji] of Object.entries(EMOJI_ICON_MAP)) {
      if (element.textContent && element.textContent.includes(emoji)) {
        replaceEmojiInTextNodes(element);
        break;
      }
    }
  }

  function replaceEmojiInNode(node) {
    if (!isValidNode(node)) return;
    
    if (node.nodeType === Node.ELEMENT_NODE) {
      replaceEmojiInElement(node);
    } else if (node.nodeType === Node.TEXT_NODE) {
      replaceEmojiInTextNode(node);
    }
  }

  function init() {
    if (document.body) {
      replaceEmojiInTextNodes(document.body);
    }
  }

  // Expose global API
  window.EMS = window.EMS || {};
  window.EMS.replaceEmojis = function(root) {
    if (!root) {
      if (document.body) replaceEmojiInTextNodes(document.body);
      return;
    }
    replaceEmojiInNode(root);
  };

  // Safe init
  function safeInit() {
    try {
      if (document.body) {
        init();
      } else {
        document.addEventListener('DOMContentLoaded', safeInit, { once: true });
        return;
      }
    } catch (e) {
      // Ignore init errors
    }

    // Observe dynamic content
    try {
      const observer = new MutationObserver(function(mutations) {
        mutations.forEach(function(m) {
          m.addedNodes.forEach(function(node) {
            if (node && node.nodeType) {
              replaceEmojiInNode(node);
            }
          });
        });
      });
      observer.observe(document.body, { childList: true, subtree: true });
    } catch (e) {
      // Ignore observer errors
    }

    // Periodic checks for async content (first 10 seconds)
    let checkCount = 0;
    const interval = setInterval(function() {
      try {
        if (document.body) replaceEmojiInTextNodes(document.body);
      } catch (e) {}
      checkCount++;
      if (checkCount >= 20) clearInterval(interval);
    }, 500);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', safeInit, { once: true });
  } else {
    safeInit();
  }
})();