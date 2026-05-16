/* Native PPTX exporter — generates an editable .pptx from the slide DOM.
 *
 * Uses PptxGenJS loaded on demand from CDN. One PPTX slide per <section> in
 * <deck-stage>. Layout is the predefined widescreen 13.333 × 7.5 in (16:9),
 * matching 1920×1080 at 144 DPI ⇒ 1 design px = 1/144 in.
 *
 * Conservative output: plain text only (no rich runs), no custom shapes,
 * named hex colors, guards on every addText so empty strings are skipped.
 * Keeps Keynote happy.
 */
(function () {
  const PPTXGEN_URL = 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';

  let pptxgenPromise = null;
  function loadPptxgen() {
    if (pptxgenPromise) return pptxgenPromise;
    pptxgenPromise = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = PPTXGEN_URL;
      s.onload = () => resolve(window.PptxGenJS);
      s.onerror = () => reject(new Error('Failed to load PptxGenJS'));
      document.head.appendChild(s);
    });
    return pptxgenPromise;
  }

  // 1920×1080 design @ 144 DPI ⇒ 13.333 × 7.5 in.
  const px = (n) => +(Math.max(0, n) / 144).toFixed(3);

  function readHex(name, fallback) {
    let v = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
    if (v.startsWith('#')) v = v.slice(1);
    return /^[0-9a-fA-F]{6}$/.test(v) ? v.toUpperCase() : fallback;
  }

  function colors() {
    const isDark = (document.documentElement.getAttribute('data-theme') || 'dark') === 'dark';
    return {
      bg: readHex('--bg', isDark ? '0A0A0B' : 'F7F7F5'),
      fg: readHex('--fg', isDark ? 'EDEDF0' : '14141A'),
      muted: readHex('--muted', '7A7A82'),
      accent: readHex('--accent', '00D492'),
    };
  }

  function pickText(el, sel) {
    if (!el) return '';
    const n = el.querySelector(sel);
    return n ? n.textContent.trim().replace(/\s+/g, ' ') : '';
  }
  function pickHTMLText(el) {
    if (!el) return '';
    // Flatten an element's text, replacing <br> with newlines.
    return Array.from(el.childNodes).map(n => {
      if (n.nodeType === 3) return n.textContent;
      if (n.tagName === 'BR') return '\n';
      return n.textContent || '';
    }).join('').trim().replace(/[ \t]+/g, ' ');
  }
  function pickAll(el, sel) {
    return el ? Array.from(el.querySelectorAll(sel)) : [];
  }

  function add(slide, str, opts) {
    if (str == null) return;
    const s = String(str).trim();
    if (!s) return;
    slide.addText(s, Object.assign({
      fontFace: 'Geist',
      valign: 'top',
      margin: 0,
      isTextBox: true,
    }, opts));
  }

  function bg(slide, c) {
    slide.background = { color: c.bg };
  }

  function eyebrow(slide, section, c) {
    const eb = pickText(section, '.eyebrow');
    if (!eb) return;
    add(slide, eb, {
      x: px(112), y: px(110), w: px(900), h: px(36),
      fontSize: 14, fontFace: 'Courier New', bold: true, color: c.fg,
    });
  }

  function chrome(slide, section, c, n, total) {
    add(slide, 'EFFECT', {
      x: px(48), y: px(28), w: px(220), h: px(24),
      fontSize: 11, fontFace: 'Courier New', bold: true, color: c.fg,
    });
    const label = (section.dataset.label || '').toUpperCase();
    if (label) add(slide, label, {
      x: px(760), y: px(28), w: px(400), h: px(24),
      fontSize: 11, fontFace: 'Courier New', color: c.muted, align: 'center',
    });
    add(slide, `${String(n).padStart(2, '0')} / ${String(total).padStart(2, '0')}`, {
      x: px(1648), y: px(28), w: px(220), h: px(24),
      fontSize: 11, fontFace: 'Courier New', color: c.muted, align: 'right',
    });
    add(slide, 'Effect — 2026', {
      x: px(1500), y: px(1032), w: px(372), h: px(20),
      fontSize: 11, fontFace: 'Courier New', color: c.muted, align: 'right',
    });
  }

  function buildTitle(pptx, section, c) {
    const slide = pptx.addSlide();
    bg(slide, c);
    const title = pickText(section, 'h1.title') || pickText(section, 'h1');
    const speakerName = pickText(section, '.t-by .name');
    const speakerHandle = pickText(section, '.t-by .handle');
    const loc = pickText(section, '.t-meta .where b');
    const dateEl = section.querySelector('.t-meta [data-tt="talk-date"]');
    const date = dateEl ? dateEl.textContent.trim() : '';

    add(slide, '~ presentation/title.md', {
      x: px(112), y: px(112), w: px(900), h: px(28),
      fontSize: 13, fontFace: 'Courier New', color: c.muted,
    });
    const where = [loc, date].filter(Boolean).join('   ');
    if (where) add(slide, where, {
      x: px(1000), y: px(112), w: px(808), h: px(28),
      fontSize: 13, fontFace: 'Courier New', color: c.muted, align: 'right', bold: true,
    });
    add(slide, title || 'Talk title', {
      x: px(112), y: px(360), w: px(1696), h: px(420),
      fontSize: 72, fontFace: 'Geist', bold: true, color: c.fg,
    });
    add(slide, 'SPEAKER', {
      x: px(112), y: px(840), w: px(300), h: px(24),
      fontSize: 12, fontFace: 'Courier New', color: c.muted,
    });
    add(slide, speakerName, {
      x: px(112), y: px(872), w: px(900), h: px(44),
      fontSize: 24, fontFace: 'Geist', bold: true, color: c.fg,
    });
    add(slide, speakerHandle, {
      x: px(112), y: px(918), w: px(900), h: px(28),
      fontSize: 14, fontFace: 'Courier New', color: c.muted,
    });
    return slide;
  }

  function buildAgenda(pptx, section, c) {
    const slide = pptx.addSlide();
    bg(slide, c);
    eyebrow(slide, section, c);
    add(slide, pickText(section, 'h2'), {
      x: px(112), y: px(180), w: px(820), h: px(180),
      fontSize: 60, fontFace: 'Geist', bold: true, color: c.fg,
    });
    add(slide, pickText(section, 'p'), {
      x: px(112), y: px(380), w: px(820), h: px(200),
      fontSize: 18, fontFace: 'Geist', color: c.muted,
    });
    const items = pickAll(section, 'ol > li').map((li, i) => {
      const label = pickText(li, 'span:not(.dur)') || li.textContent.trim();
      const dur = pickText(li, '.dur');
      return `${String(i + 1).padStart(2, '0')}   ${label}${dur ? '   ' + dur : ''}`;
    });
    if (items.length) add(slide, items.join('\n'), {
      x: px(980), y: px(180), w: px(828), h: px(800),
      fontSize: 20, fontFace: 'Geist', color: c.fg, paraSpaceAfter: 8,
    });
    return slide;
  }

  function buildSpeaker(pptx, section, c) {
    const slide = pptx.addSlide();
    bg(slide, c);
    eyebrow(slide, section, c);
    add(slide, pickText(section, 'h2'), {
      x: px(112), y: px(180), w: px(1696), h: px(140),
      fontSize: 56, fontFace: 'Geist', bold: true, color: c.fg,
    });
    add(slide, pickText(section, '.who .role'), {
      x: px(112), y: px(340), w: px(1696), h: px(40),
      fontSize: 20, fontFace: 'Geist', color: c.muted,
    });
    add(slide, pickText(section, '.who p'), {
      x: px(112), y: px(400), w: px(1100), h: px(200),
      fontSize: 18, fontFace: 'Geist', color: c.muted,
    });
    const items = pickAll(section, '.who ul li').map(li => `— ${li.textContent.trim()}`);
    if (items.length) add(slide, items.join('\n'), {
      x: px(112), y: px(640), w: px(1100), h: px(320),
      fontSize: 18, fontFace: 'Geist', color: c.fg, paraSpaceAfter: 6,
    });
    return slide;
  }

  function buildStatement(pptx, section, c) {
    const slide = pptx.addSlide();
    bg(slide, c);
    eyebrow(slide, section, c);
    const txt = pickHTMLText(section.querySelector('h2'));
    add(slide, txt, {
      x: px(112), y: px(280), w: px(1696), h: px(560),
      fontSize: 72, fontFace: 'Geist', bold: true, color: c.fg, valign: 'middle',
    });
    return slide;
  }

  function buildQuote(pptx, section, c) {
    const slide = pptx.addSlide();
    bg(slide, c);
    add(slide, '"', {
      x: px(112), y: px(220), w: px(200), h: px(140),
      fontSize: 96, fontFace: 'Courier New', bold: true, color: c.accent,
    });
    const q = pickHTMLText(section.querySelector('blockquote') || section.querySelector('.q'));
    add(slide, q, {
      x: px(112), y: px(360), w: px(1696), h: px(440),
      fontSize: 48, fontFace: 'Geist', color: c.fg,
    });
    const cite = pickText(section, 'cite') || pickText(section, '.cite');
    if (cite) add(slide, `— ${cite}`, {
      x: px(112), y: px(840), w: px(1696), h: px(40),
      fontSize: 18, fontFace: 'Courier New', color: c.muted,
    });
    return slide;
  }

  function buildStats(pptx, section, c) {
    const slide = pptx.addSlide();
    bg(slide, c);
    eyebrow(slide, section, c);
    add(slide, pickText(section, 'h2'), {
      x: px(112), y: px(180), w: px(1696), h: px(80),
      fontSize: 44, fontFace: 'Geist', bold: true, color: c.fg,
    });
    const stats = pickAll(section, '.stat');
    const cols = Math.max(1, stats.length);
    const colW = (1696) / cols;
    stats.forEach((s, i) => {
      const x = 112 + i * colW;
      const vEl = s.querySelector('.v')?.cloneNode(true);
      let value = '';
      let pre = '';
      if (vEl) {
        pre = pickText(vEl, '.pre');
        vEl.querySelectorAll('.pre, .sub').forEach(n => n.remove());
        value = vEl.textContent.trim().replace(/\s+/g, ' ');
      }
      const lbl = pickText(s, '.lbl');
      const sub = pickText(s, '.sub');
      if (pre) add(slide, pre, {
        x: px(x), y: px(380), w: px(colW - 32), h: px(28),
        fontSize: 14, fontFace: 'Courier New', color: c.accent,
      });
      add(slide, value, {
        x: px(x), y: px(412), w: px(colW - 32), h: px(140),
        fontSize: 72, fontFace: 'Geist', bold: true, color: c.fg,
      });
      if (lbl) add(slide, lbl, {
        x: px(x), y: px(560), w: px(colW - 32), h: px(28),
        fontSize: 13, fontFace: 'Courier New', color: c.muted,
      });
      if (sub) add(slide, sub, {
        x: px(x), y: px(596), w: px(colW - 32), h: px(120),
        fontSize: 14, fontFace: 'Geist', color: c.muted,
      });
    });
    return slide;
  }

  function buildCode(pptx, section, c) {
    const slide = pptx.addSlide();
    bg(slide, c);
    eyebrow(slide, section, c);
    add(slide, pickText(section, 'h2'), {
      x: px(112), y: px(180), w: px(1696), h: px(80),
      fontSize: 36, fontFace: 'Geist', bold: true, color: c.fg,
    });
    const pre = section.querySelector('pre');
    if (pre) {
      const code = pre.textContent.replace(/\n+$/, '');
      add(slide, code, {
        x: px(112), y: px(290), w: px(1696), h: px(720),
        fontSize: 16, fontFace: 'Courier New', color: c.fg,
      });
    }
    return slide;
  }

  function buildResources(pptx, section, c) {
    const slide = pptx.addSlide();
    bg(slide, c);
    eyebrow(slide, section, c);
    add(slide, pickText(section, 'h2'), {
      x: px(112), y: px(180), w: px(1696), h: px(80),
      fontSize: 52, fontFace: 'Geist', bold: true, color: c.fg,
    });
    pickAll(section, '.item').forEach((it, i) => {
      const y = 320 + i * 110;
      const title = pickText(it, '.t') || pickText(it, 'h3') || pickText(it, '.title');
      const url = pickText(it, '.url');
      const desc = pickText(it, '.desc') || pickText(it, 'p');
      add(slide, title, { x: px(112), y: px(y), w: px(700), h: px(36), fontSize: 22, bold: true, color: c.fg });
      if (desc) add(slide, desc, { x: px(112), y: px(y + 38), w: px(700), h: px(60), fontSize: 14, color: c.muted });
      if (url) add(slide, url, { x: px(840), y: px(y + 6), w: px(900), h: px(36), fontSize: 16, fontFace: 'Courier New', color: c.accent });
    });
    return slide;
  }

  function buildThanks(pptx, section, c) {
    const slide = pptx.addSlide();
    bg(slide, c);
    const heading = pickHTMLText(section.querySelector('h2'));
    add(slide, heading, {
      x: px(112), y: px(240), w: px(1696), h: px(360),
      fontSize: 104, fontFace: 'Geist', bold: true, color: c.fg, valign: 'middle',
    });
    pickAll(section, '.links .row').forEach((r, i) => {
      const x = 112 + i * 440;
      add(slide, pickText(r, '.lbl'), {
        x: px(x), y: px(720), w: px(420), h: px(28),
        fontSize: 12, fontFace: 'Courier New', color: c.muted,
      });
      add(slide, pickText(r, '.val'), {
        x: px(x), y: px(752), w: px(420), h: px(40),
        fontSize: 22, fontFace: 'Geist', bold: true, color: c.fg,
      });
    });
    return slide;
  }

  function buildGeneric(pptx, section, c) {
    const slide = pptx.addSlide();
    bg(slide, c);
    eyebrow(slide, section, c);
    const h = pickHTMLText(section.querySelector('h2')) || pickText(section, 'h1');
    if (h) add(slide, h, {
      x: px(112), y: px(180), w: px(1696), h: px(220),
      fontSize: 52, fontFace: 'Geist', bold: true, color: c.fg,
    });
    const p = pickText(section, 'p');
    if (p) add(slide, p, {
      x: px(112), y: px(420), w: px(1696), h: px(140),
      fontSize: 20, fontFace: 'Geist', color: c.muted,
    });
    const lis = pickAll(section, 'ul li, ol li').map(li => `— ${li.textContent.trim()}`);
    if (lis.length) add(slide, lis.join('\n'), {
      x: px(112), y: px(580), w: px(1696), h: px(440),
      fontSize: 22, fontFace: 'Geist', color: c.fg, paraSpaceAfter: 6,
    });
    return slide;
  }

  const BUILDERS = {
    's-title': buildTitle,
    's-agenda': buildAgenda,
    's-speaker': buildSpeaker,
    's-statement': buildStatement,
    's-section': buildStatement,
    's-quote': buildQuote,
    's-stats': buildStats,
    's-code': buildCode,
    's-codecom': buildCode,
    's-resources': buildResources,
    's-thanks': buildThanks,
  };

  function pickBuilder(section) {
    for (const cls of section.classList) {
      if (BUILDERS[cls]) return BUILDERS[cls];
    }
    return buildGeneric;
  }

  async function exportPPTX() {
    const PptxGenJS = await loadPptxgen();
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // predefined 13.333 × 7.5 in (16:9)
    const c = colors();
    const sections = document.querySelectorAll('deck-stage > section');
    const total = sections.length;
    sections.forEach((section, i) => {
      try {
        const build = pickBuilder(section);
        const slide = build(pptx, section, c);
        if (slide && !section.classList.contains('s-title')) {
          chrome(slide, section, c, i + 1, total);
        }
      } catch (err) {
        console.error('slide', i + 1, 'failed:', err);
      }
    });

    const title = (window.__TWEAKS && window.__TWEAKS.talkTitle) || 'effect-slides';
    const safe = title.replace(/[^a-z0-9-]+/gi, '-').replace(/^-+|-+$/g, '').toLowerCase() || 'slides';
    await pptx.writeFile({ fileName: `${safe}.pptx` });
  }

  window.exportPPTX = exportPPTX;
})();
