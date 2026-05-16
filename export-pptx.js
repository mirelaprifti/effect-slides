/* Native PPTX exporter — generates an editable .pptx from the slide DOM.
 *
 * Uses PptxGenJS loaded on demand from CDN. Each <section> in <deck-stage>
 * becomes one PPTX slide. Layout is 13.333 x 7.5 in (matching 1920x1080 at
 * 144 DPI, so 1 design px = 1/144 in).
 *
 * Slide builders are dispatched by the section's class (s-title, s-agenda,
 * s-speaker, ...). Anything without a dedicated builder falls back to a
 * generic eyebrow + heading + body layout.
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

  // 1920x1080 design @ 144 DPI ⇒ 13.333 x 7.5 in.
  const px = (n) => +(n / 144).toFixed(3);

  function cssVar(name) {
    const v = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
    return v.startsWith('#') ? v.slice(1) : v;
  }

  function text(el, sel) {
    if (!el) return '';
    const n = el.querySelector(sel);
    return n ? n.textContent.trim().replace(/\s+/g, ' ') : '';
  }

  function all(el, sel) {
    return el ? Array.from(el.querySelectorAll(sel)) : [];
  }

  function colors() {
    return {
      bg: cssVar('--bg') || '0a0a0b',
      fg: cssVar('--fg') || 'ededf0',
      muted: cssVar('--muted') || '7a7a82',
      accent: cssVar('--accent') || '00D492',
    };
  }

  function bgFor(slide, c) {
    slide.background = { color: c.bg };
  }

  function addText(slide, str, opts) {
    if (!str) return;
    slide.addText(str, Object.assign({
      fontFace: 'Geist',
      color: opts.color || colors().fg,
      valign: 'top',
      margin: 0,
    }, opts));
  }

  function addEyebrow(slide, section, c) {
    const eb = text(section, '.eyebrow');
    if (!eb) return;
    addText(slide, eb, {
      x: px(112), y: px(110), w: px(800), h: px(36),
      fontSize: 14, fontFace: 'JetBrains Mono', bold: true,
      color: c.fg, charSpacing: 5,
    });
  }

  function addChrome(slide, section, c, pageNo, total) {
    const label = section.dataset.label || '';
    // Top-left: Effect mark (text-only since we can't inline SVG mark here)
    addText(slide, 'Effect', {
      x: px(48), y: px(28), w: px(200), h: px(24),
      fontSize: 11, fontFace: 'Geist', bold: true, color: c.fg, charSpacing: 4,
    });
    // Top-center: slide label
    if (label) {
      addText(slide, label.toUpperCase(), {
        x: px(760), y: px(28), w: px(400), h: px(24),
        fontSize: 11, fontFace: 'JetBrains Mono', color: c.muted,
        align: 'center', charSpacing: 4,
      });
    }
    // Top-right: page number
    addText(slide, `${String(pageNo).padStart(2, '0')} / ${String(total).padStart(2, '0')}`, {
      x: px(1672), y: px(28), w: px(200), h: px(24),
      fontSize: 11, fontFace: 'JetBrains Mono', color: c.muted, align: 'right',
    });
    // Bottom-right: Effect 2026
    addText(slide, 'Effect 2026', {
      x: px(1572), y: px(1032), w: px(300), h: px(20),
      fontSize: 11, fontFace: 'JetBrains Mono', color: c.muted, align: 'right',
    });
  }

  function buildTitle(pptx, section, c) {
    const slide = pptx.addSlide();
    bgFor(slide, c);

    const title = text(section, 'h1.title') || text(section, 'h1');
    const speakerName = text(section, '.t-by .name');
    const speakerHandle = text(section, '.t-by .handle');
    const loc = text(section, '.t-meta .where b');
    const dateEl = section.querySelector('.t-meta .where [data-tt="talk-date"]');
    const date = dateEl ? dateEl.textContent.trim() : '';
    const meta = text(section, '.t-meta > span:first-child');

    addText(slide, meta || '~ presentation/title.md', {
      x: px(112), y: px(112), w: px(1000), h: px(28),
      fontSize: 13, fontFace: 'JetBrains Mono', color: c.muted,
    });
    addText(slide, [loc, date].filter(Boolean).join('   '), {
      x: px(1100), y: px(112), w: px(708), h: px(28),
      fontSize: 13, fontFace: 'JetBrains Mono', color: c.muted, align: 'right', bold: true,
    });

    addText(slide, title, {
      x: px(112), y: px(360), w: px(1696), h: px(360),
      fontSize: 76, fontFace: 'Geist', bold: true, color: c.fg,
    });

    addText(slide, 'SPEAKER', {
      x: px(112), y: px(840), w: px(200), h: px(24),
      fontSize: 12, fontFace: 'JetBrains Mono', color: c.muted, charSpacing: 5,
    });
    addText(slide, speakerName, {
      x: px(112), y: px(870), w: px(800), h: px(40),
      fontSize: 24, fontFace: 'Geist', bold: true, color: c.fg,
    });
    addText(slide, speakerHandle, {
      x: px(112), y: px(914), w: px(800), h: px(28),
      fontSize: 14, fontFace: 'JetBrains Mono', color: c.muted,
    });

    return slide;
  }

  function buildAgenda(pptx, section, c) {
    const slide = pptx.addSlide();
    bgFor(slide, c);
    addEyebrow(slide, section, c);
    const heading = text(section, 'h2');
    const lede = text(section, 'p');
    addText(slide, heading, {
      x: px(112), y: px(180), w: px(820), h: px(180),
      fontSize: 64, fontFace: 'Geist', bold: true, color: c.fg,
    });
    if (lede) addText(slide, lede, {
      x: px(112), y: px(380), w: px(820), h: px(200),
      fontSize: 18, fontFace: 'Geist', color: c.muted,
    });
    const items = all(section, 'ol > li').map((li, i) => ({
      n: String(i + 1).padStart(2, '0'),
      label: text(li, 'span:not(.dur)'),
      dur: text(li, '.dur'),
    }));
    const rows = items.map(it => [
      { text: `${it.n}`, options: { fontFace: 'JetBrains Mono', color: c.muted, bold: true } },
      { text: `   ${it.label}`, options: { fontFace: 'Geist', color: c.fg, bold: true } },
      { text: it.dur ? `   ${it.dur}` : '', options: { fontFace: 'JetBrains Mono', color: c.muted } },
    ]);
    rows.forEach((row, i) => {
      slide.addText(row, {
        x: px(980), y: px(180 + i * 60), w: px(828), h: px(48),
        fontSize: 18, valign: 'top',
      });
    });
    return slide;
  }

  function buildSpeaker(pptx, section, c) {
    const slide = pptx.addSlide();
    bgFor(slide, c);
    addEyebrow(slide, section, c);
    const name = text(section, 'h2');
    const role = text(section, '.who .role');
    const bio = text(section, '.who p');
    const items = all(section, '.who ul li').map(li => li.textContent.trim());

    addText(slide, name, {
      x: px(112), y: px(180), w: px(1100), h: px(160),
      fontSize: 60, fontFace: 'Geist', bold: true, color: c.fg,
    });
    if (role) addText(slide, role, {
      x: px(112), y: px(340), w: px(1100), h: px(40),
      fontSize: 20, fontFace: 'Geist', color: c.muted,
    });
    if (bio) addText(slide, bio, {
      x: px(112), y: px(420), w: px(1100), h: px(160),
      fontSize: 18, fontFace: 'Geist', color: c.muted,
    });
    if (items.length) {
      slide.addText(items.map(t => ({ text: `— ${t}`, options: { breakLine: true } })), {
        x: px(112), y: px(620), w: px(1100), h: px(280),
        fontSize: 18, fontFace: 'Geist', color: c.fg, valign: 'top',
      });
    }
    return slide;
  }

  function buildStatement(pptx, section, c) {
    const slide = pptx.addSlide();
    bgFor(slide, c);
    addEyebrow(slide, section, c);
    const heading = section.querySelector('h2');
    if (heading) {
      // Preserve <em> as accent color via rich text runs.
      const runs = [];
      heading.childNodes.forEach(n => {
        if (n.nodeType === 3) {
          runs.push({ text: n.textContent, options: { color: c.fg } });
        } else if (n.tagName === 'BR') {
          runs.push({ text: '\n', options: { breakLine: true } });
        } else if (n.tagName && n.tagName.toLowerCase() === 'em') {
          runs.push({ text: n.textContent, options: { color: c.accent } });
        } else {
          runs.push({ text: n.textContent, options: { color: c.fg } });
        }
      });
      slide.addText(runs, {
        x: px(112), y: px(280), w: px(1696), h: px(560),
        fontSize: 80, fontFace: 'Geist', bold: true,
        valign: 'middle',
      });
    }
    return slide;
  }

  function buildQuote(pptx, section, c) {
    const slide = pptx.addSlide();
    bgFor(slide, c);
    const open = text(section, '.open') || '"';
    const quote = text(section, 'blockquote') || text(section, '.q');
    const cite = text(section, 'cite') || text(section, '.cite');
    addText(slide, open, {
      x: px(112), y: px(220), w: px(200), h: px(120),
      fontSize: 96, fontFace: 'JetBrains Mono', bold: true, color: c.accent,
    });
    addText(slide, quote, {
      x: px(112), y: px(360), w: px(1696), h: px(440),
      fontSize: 56, fontFace: 'Geist', color: c.fg,
    });
    if (cite) addText(slide, `— ${cite}`, {
      x: px(112), y: px(840), w: px(1696), h: px(40),
      fontSize: 18, fontFace: 'JetBrains Mono', color: c.muted,
    });
    return slide;
  }

  function buildStats(pptx, section, c) {
    const slide = pptx.addSlide();
    bgFor(slide, c);
    addEyebrow(slide, section, c);
    addText(slide, text(section, 'h2'), {
      x: px(112), y: px(180), w: px(1696), h: px(80),
      fontSize: 48, fontFace: 'Geist', bold: true, color: c.fg,
    });
    const stats = all(section, '.stat');
    const cols = Math.max(1, stats.length);
    const colW = (1920 - 224) / cols;
    stats.forEach((s, i) => {
      const pre = text(s, '.v .pre');
      const v = (s.querySelector('.v')?.cloneNode(true));
      let value = '';
      if (v) {
        v.querySelectorAll('.pre, .sub').forEach(n => n.remove());
        value = v.textContent.trim().replace(/\s+/g, ' ');
      }
      const lbl = text(s, '.lbl');
      const sub = text(s, '.sub');
      const isDark = s.classList.contains('dark');
      const x = 112 + i * colW;
      slide.addShape('rect', {
        x: px(x), y: px(360), w: px(colW - 24), h: px(360),
        fill: { color: isDark ? c.fg : c.bg },
        line: { color: isDark ? c.fg : c.muted, width: 1 },
      });
      if (pre) addText(slide, pre, {
        x: px(x + 32), y: px(388), w: px(colW - 88), h: px(28),
        fontSize: 14, fontFace: 'JetBrains Mono', color: c.accent,
      });
      addText(slide, value, {
        x: px(x + 32), y: px(420), w: px(colW - 88), h: px(140),
        fontSize: 84, fontFace: 'Geist', bold: true,
        color: isDark ? c.bg : c.fg,
      });
      if (lbl) addText(slide, lbl, {
        x: px(x + 32), y: px(580), w: px(colW - 88), h: px(28),
        fontSize: 13, fontFace: 'JetBrains Mono',
        color: isDark ? c.bg : c.muted, charSpacing: 4,
      });
      if (sub) addText(slide, sub, {
        x: px(x + 32), y: px(620), w: px(colW - 88), h: px(80),
        fontSize: 14, fontFace: 'Geist',
        color: isDark ? c.bg : c.muted,
      });
    });
    return slide;
  }

  function buildCode(pptx, section, c) {
    const slide = pptx.addSlide();
    bgFor(slide, c);
    addEyebrow(slide, section, c);
    addText(slide, text(section, 'h2'), {
      x: px(112), y: px(180), w: px(1696), h: px(80),
      fontSize: 40, fontFace: 'Geist', bold: true, color: c.fg,
    });
    const pre = section.querySelector('pre');
    if (pre) {
      const code = pre.textContent.replace(/\n$/, '');
      slide.addShape('rect', {
        x: px(112), y: px(290), w: px(1696), h: px(720),
        fill: { color: c.bg }, line: { color: c.muted, width: 1 },
      });
      slide.addText(code, {
        x: px(140), y: px(310), w: px(1640), h: px(680),
        fontSize: 16, fontFace: 'JetBrains Mono', color: c.fg, valign: 'top',
      });
    }
    return slide;
  }

  function buildResources(pptx, section, c) {
    const slide = pptx.addSlide();
    bgFor(slide, c);
    addEyebrow(slide, section, c);
    addText(slide, text(section, 'h2'), {
      x: px(112), y: px(180), w: px(1696), h: px(80),
      fontSize: 56, fontFace: 'Geist', bold: true, color: c.fg,
    });
    const items = all(section, '.item').map(it => ({
      title: text(it, '.t') || text(it, 'h3') || text(it, '.title'),
      url: text(it, '.url'),
      desc: text(it, '.desc') || text(it, 'p'),
    }));
    items.forEach((it, i) => {
      const y = 320 + i * 110;
      addText(slide, it.title, {
        x: px(112), y: px(y), w: px(700), h: px(36),
        fontSize: 22, fontFace: 'Geist', bold: true, color: c.fg,
      });
      if (it.desc) addText(slide, it.desc, {
        x: px(112), y: px(y + 38), w: px(700), h: px(60),
        fontSize: 14, fontFace: 'Geist', color: c.muted,
      });
      if (it.url) addText(slide, it.url, {
        x: px(840), y: px(y + 6), w: px(900), h: px(36),
        fontSize: 16, fontFace: 'JetBrains Mono', color: c.accent,
      });
    });
    return slide;
  }

  function buildThanks(pptx, section, c) {
    const slide = pptx.addSlide();
    bgFor(slide, c);
    const heading = section.querySelector('h2');
    if (heading) {
      const runs = [];
      heading.childNodes.forEach(n => {
        if (n.nodeType === 3) runs.push({ text: n.textContent, options: { color: c.fg } });
        else if (n.tagName === 'BR') runs.push({ text: '\n', options: { breakLine: true } });
        else if (n.tagName.toLowerCase() === 'em') runs.push({ text: n.textContent, options: { color: c.accent } });
        else runs.push({ text: n.textContent, options: { color: c.fg } });
      });
      slide.addText(runs, {
        x: px(112), y: px(240), w: px(1696), h: px(360),
        fontSize: 120, fontFace: 'Geist', bold: true, valign: 'middle',
      });
    }
    const rows = all(section, '.links .row').map(r => ({
      lbl: text(r, '.lbl'),
      val: text(r, '.val'),
    }));
    rows.forEach((row, i) => {
      const x = 112 + i * 440;
      addText(slide, row.lbl, {
        x: px(x), y: px(720), w: px(420), h: px(28),
        fontSize: 12, fontFace: 'JetBrains Mono', color: c.muted, charSpacing: 4,
      });
      addText(slide, row.val, {
        x: px(x), y: px(752), w: px(420), h: px(40),
        fontSize: 22, fontFace: 'Geist', bold: true, color: c.fg,
      });
    });
    return slide;
  }

  // Generic fallback — eyebrow, h2, p, optional list.
  function buildGeneric(pptx, section, c) {
    const slide = pptx.addSlide();
    bgFor(slide, c);
    addEyebrow(slide, section, c);
    const h = text(section, 'h2') || text(section, 'h1');
    if (h) addText(slide, h, {
      x: px(112), y: px(180), w: px(1696), h: px(180),
      fontSize: 56, fontFace: 'Geist', bold: true, color: c.fg,
    });
    const p = text(section, 'p');
    if (p) addText(slide, p, {
      x: px(112), y: px(400), w: px(1696), h: px(120),
      fontSize: 20, fontFace: 'Geist', color: c.muted,
    });
    const lis = all(section, 'ul li, ol li').map(li => li.textContent.trim());
    if (lis.length) {
      slide.addText(lis.map(t => ({ text: `— ${t}`, options: { breakLine: true } })), {
        x: px(112), y: px(540), w: px(1696), h: px(440),
        fontSize: 22, fontFace: 'Geist', color: c.fg, valign: 'top',
      });
    }
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
    pptx.defineLayout({ name: 'EFFECT_16_9', width: 13.333, height: 7.5 });
    pptx.layout = 'EFFECT_16_9';

    const c = colors();
    const sections = document.querySelectorAll('deck-stage > section');
    const total = sections.length;
    sections.forEach((section, i) => {
      const build = pickBuilder(section);
      const slide = build(pptx, section, c);
      if (slide && !section.classList.contains('s-title')) {
        addChrome(slide, section, c, i + 1, total);
      }
    });

    const title = (window.__TWEAKS && window.__TWEAKS.talkTitle) || 'Effect Slides';
    const safe = title.replace(/[^a-z0-9-]+/gi, '-').replace(/^-+|-+$/g, '').toLowerCase() || 'slides';
    await pptx.writeFile({ fileName: `${safe}.pptx` });
  }

  window.exportPPTX = exportPPTX;
})();
