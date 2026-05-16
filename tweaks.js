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
})();
