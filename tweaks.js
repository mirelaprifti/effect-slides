/* Tweaks wiring shared by both decks.
 * Listens for parent host's edit-mode messages, manages a small floating panel,
 * and persists tweaks via __edit_mode_set_keys so they survive reload via the
 * EDITMODE-BEGIN/END block in the host HTML.
 *
 * Bound DOM targets:
 *   [data-tt="talk-title"]       <- talk title (block)
 *   [data-tt="talk-title-inline"]<- talk title (one-line, lowercase)
 *   [data-tt="speaker-name"]
 *   [data-tt="speaker-handle"]
 *   [data-tt="speaker-handle-bare"] <- handle w/o leading @
 *   [data-tt="talk-loc"]
 *   [data-tt="talk-date"]
 */
(function () {
  const root = document.documentElement;

  function applyTheme(theme) {
    root.setAttribute('data-theme', theme);
  }

  function applyText(t) {
    document.querySelectorAll('[data-tt="talk-title"]').forEach(el => el.textContent = t.talkTitle);
    document.querySelectorAll('[data-tt="talk-title-inline"]').forEach(el => el.textContent = t.talkTitle);
    document.querySelectorAll('[data-tt="speaker-name"]').forEach(el => el.textContent = t.speakerName);
    document.querySelectorAll('[data-tt="speaker-handle"]').forEach(el => el.textContent = t.speakerHandle.startsWith('@') ? t.speakerHandle : '@' + t.speakerHandle);
    document.querySelectorAll('[data-tt="speaker-handle-bare"]').forEach(el => el.textContent = t.speakerHandle.replace(/^@/, ''));
    document.querySelectorAll('[data-tt="talk-loc"]').forEach(el => el.textContent = t.talkLoc);
    document.querySelectorAll('[data-tt="talk-date"]').forEach(el => el.textContent = t.talkDate);
  }

  function apply(t) {
    applyTheme(t.theme);
    applyText(t);
  }

  // Initial paint
  document.addEventListener('DOMContentLoaded', () => apply(window.__TWEAKS));

  // Build the panel
  let panel = null;
  function buildPanel() {
    if (panel) return panel;
    panel = document.createElement('div');
    panel.id = 'tweaks-panel';
    panel.innerHTML = `
      <div class="tp-head">
        <span class="tp-title">Tweaks</span>
        <button class="tp-close" aria-label="Close">×</button>
      </div>
      <div class="tp-body">
        <div class="tp-row">
          <label class="tp-label">Theme</label>
          <div class="tp-seg" data-tp-seg="theme">
            <button data-v="dark">Dark</button>
            <button data-v="light">Light</button>
          </div>
        </div>
        <div class="tp-row">
          <label class="tp-label">Talk title</label>
          <textarea data-tp-text="talkTitle" rows="2"></textarea>
        </div>
        <div class="tp-row tp-row-2">
          <div>
            <label class="tp-label">Speaker</label>
            <input type="text" data-tp-text="speakerName" />
          </div>
          <div>
            <label class="tp-label">Handle</label>
            <input type="text" data-tp-text="speakerHandle" />
          </div>
        </div>
        <div class="tp-row tp-row-2">
          <div>
            <label class="tp-label">Date</label>
            <input type="text" data-tp-text="talkDate" />
          </div>
          <div>
            <label class="tp-label">Location</label>
            <input type="text" data-tp-text="talkLoc" />
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(panel);

    // Wire close
    panel.querySelector('.tp-close').addEventListener('click', () => {
      hidePanel();
      window.parent.postMessage({ type: '__edit_mode_dismissed' }, '*');
    });

    // Wire theme segmented control
    const seg = panel.querySelector('[data-tp-seg="theme"]');
    seg.addEventListener('click', (e) => {
      const btn = e.target.closest('button[data-v]');
      if (!btn) return;
      const v = btn.dataset.v;
      window.__TWEAKS.theme = v;
      apply(window.__TWEAKS);
      syncPanel();
      window.parent.postMessage({ type: '__edit_mode_set_keys', edits: { theme: v } }, '*');
    });

    // Wire text inputs
    panel.querySelectorAll('[data-tp-text]').forEach(input => {
      input.addEventListener('input', () => {
        const key = input.dataset.tpText;
        window.__TWEAKS[key] = input.value;
        applyText(window.__TWEAKS);
        window.parent.postMessage({ type: '__edit_mode_set_keys', edits: { [key]: input.value } }, '*');
      });
    });

    syncPanel();
    return panel;
  }

  function syncPanel() {
    if (!panel) return;
    const t = window.__TWEAKS;
    panel.querySelectorAll('[data-tp-seg="theme"] button').forEach(b => {
      b.classList.toggle('on', b.dataset.v === t.theme);
    });
    panel.querySelectorAll('[data-tp-text]').forEach(input => {
      const key = input.dataset.tpText;
      if (document.activeElement !== input) input.value = t[key] ?? '';
    });
  }

  function showPanel() {
    buildPanel();
    panel.classList.add('on');
    syncPanel();
  }

  function hidePanel() {
    if (panel) panel.classList.remove('on');
  }

  // Host protocol
  window.addEventListener('message', (e) => {
    const d = e.data;
    if (!d || typeof d !== 'object') return;
    if (d.type === '__activate_edit_mode') showPanel();
    if (d.type === '__deactivate_edit_mode') hidePanel();
  });

  // Announce after listener is live
  window.parent.postMessage({ type: '__edit_mode_available' }, '*');

  // Always-visible floating controls — theme toggle + PPTX export.
  function buildThemeToggle() {
    if (document.getElementById('deck-controls')) return;
    const wrap = document.createElement('div');
    wrap.id = 'deck-controls';
    wrap.className = 'export-hidden';
    wrap.innerHTML = `
      <button id="theme-toggle" type="button" aria-label="Toggle light/dark">
        <span class="tt-icon" aria-hidden="true"></span><span class="tt-label"></span>
      </button>
    `;
    const style = document.createElement('style');
    style.textContent = `
      #deck-controls {
        position: fixed; top: 16px; right: 16px;
        z-index: 99998;
        display: inline-flex; align-items: center; gap: 8px;
      }
      #deck-controls button {
        display: inline-flex; align-items: center; gap: 8px;
        padding: 8px 12px;
        background: var(--panel, #15151a);
        color: var(--fg, #ededf0);
        border: 1px solid var(--hairline-2, #2a2a31);
        border-radius: 999px;
        font-family: 'JetBrains Mono', ui-monospace, monospace;
        font-size: 11px; letter-spacing: 0.14em; text-transform: uppercase;
        cursor: pointer;
        box-shadow: 0 6px 18px rgba(0,0,0,0.18);
      }
      #deck-controls button:hover { border-color: var(--muted-2, #4a4a55); }
      #deck-controls button:disabled { opacity: 0.6; cursor: progress; }
      #theme-toggle .tt-icon {
        width: 14px; height: 14px; border-radius: 50%;
        background: var(--fg, #ededf0);
        box-shadow: inset -4px -4px 0 0 var(--panel, #15151a);
      }
      html[data-theme="light"] #theme-toggle .tt-icon {
        background: var(--fg); box-shadow: none;
      }
    `;
    document.head.appendChild(style);
    document.body.appendChild(wrap);

    wrap.querySelector('#theme-toggle').addEventListener('click', () => {
      const next = window.__TWEAKS.theme === 'dark' ? 'light' : 'dark';
      window.__TWEAKS.theme = next;
      apply(window.__TWEAKS);
      syncPanel();
      syncToggle();
      window.parent.postMessage({ type: '__edit_mode_set_keys', edits: { theme: next } }, '*');
    });

    syncToggle();
  }
  function syncToggle() {
    const btn = document.getElementById('theme-toggle');
    if (!btn) return;
    const t = window.__TWEAKS.theme;
    btn.querySelector('.tt-label').textContent = t === 'dark' ? 'Dark' : 'Light';
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', buildThemeToggle);
  } else {
    buildThemeToggle();
  }
})();
