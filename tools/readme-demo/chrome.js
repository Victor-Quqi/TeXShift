// Builds the shared OneNote-like window and renders math/Mermaid placeholders.
(function () {
  'use strict';

  const GRIP = '<svg width="14" height="2"><g fill="#757778"><rect width="2" height="2"/>' +
    '<rect x="4" width="2" height="2"/><rect x="8" width="2" height="2"/><rect x="12" width="2" height="2"/></g></svg>' +
    '<svg class="resize" width="12" height="6"><g fill="#757778"><path d="M0 3L3.3 0V6Z"/><path d="M12 3L8.7 0V6Z"/></g></svg>';

  const TABS = ['File', 'Home', 'Insert', 'Draw', 'History', 'Review', 'View', 'Help', 'TeXShift'];

  // opts: pageTitle, notebook, sections [[name, color, textColor], ...] (first is open),
  // pages [...] (first is open), height (stage height in px).
  function buildWindow(opts) {
    const I = window.Demo.icons;
    const root = document.createElement('div');
    root.className = 'window';
    root.style.setProperty('--section', opts.sections[0][1]);
    if (opts.height) root.style.setProperty('--stage-h', opts.height + 'px');
    root.innerHTML = `
      <div class="titlebar">
        ${I.app()}
        <div class="qat">${I.qatUndo()}${I.qatPrint()}${I.qatRedo()}${I.qatMore()}</div>
        <div class="doc-title">${opts.pageTitle}  -  OneNote</div>
        <div class="search">${I.search()}<span>Search</span></div>
        <div class="win-controls"><span>${I.win.min}</span><span>${I.win.max}</span><span>${I.win.close}</span></div>
      </div>
      <div class="tabs">${TABS.map(t => `<div class="tab${t === 'TeXShift' ? ' active' : ''}">${t}</div>`).join('')}</div>
      <div class="ribbon">
        <div class="rbtn" data-btn="convert"><i class="rbtn-hl"></i>${I.convert()}<span>Convert</span></div>
        <div class="rbtn" data-btn="reverse"><i class="rbtn-hl"></i>${I.reverse()}<span>Reverse Convert</span></div>
        <div class="rsep"></div>
        <div class="rbtn" data-btn="settings"><i class="rbtn-hl"></i>${I.settings()}<span>Settings</span></div>
        <div class="rsep"></div>
        <div class="ribbon-switch">${I.chevronDown(14)}</div>
      </div>
      <div class="nbbar">
        <div class="nb-name">${I.notebook()}<span>${opts.notebook}</span>${I.chevronDown(14)}</div>
        ${opts.sections.map(([name, color, ink], i) => `
          <div class="stab${i ? '' : ' active'}" style="background:${color};color:${ink}">${name}</div>`).join('')}
        <div class="stab add">+</div>
      </div>
      <div class="section-line"></div>
      <div class="body">
        <nav class="pages">
          <div class="pages-top"><div class="add-page">${I.addPage()}<span>Add Page</span></div>${I.sort()}</div>
          ${opts.pages.map((p, i) => `<div class="page-item${i === 0 ? ' active' : ''}">${p}</div>`).join('')}
        </nav>
        <main class="canvas">
          <div class="page-head">
            <div class="page-title">${opts.pageTitle}</div>
            <div class="page-date"><span>Tuesday, October 6, 2026</span><span>10:24 PM</span></div>
          </div>
          <div class="outline">
            <div class="outline-bar">${GRIP}</div>
            <div class="outline-body"></div>
          </div>
        </main>
      </div>
      <div class="overlay"></div>`;
    document.body.appendChild(root);

    const q = s => root.querySelector(s);
    return {
      root,
      overlay: q('.overlay'),
      canvas: q('.canvas'),
      outline: q('.outline'),
      body: q('.outline-body'),
      buttons: {
        convert: q('[data-btn="convert"]'),
        reverse: q('[data-btn="reverse"]'),
        settings: q('[data-btn="settings"]'),
      },
      // Center of an element in window coordinates, optionally offset.
      center(el, dx = 0, dy = 0) {
        const r = el.getBoundingClientRect();
        const w = root.getBoundingClientRect();
        return [r.left - w.left + r.width / 2 + dx, r.top - w.top + r.height / 2 + dy];
      },
      // Idle pointer spot on the page, right of the note container.
      restPoint(gap) {
        const o = this.rect(this.outline);
        return [Math.min(o.left + o.width + gap, root.offsetWidth - 60), Math.min(540, root.offsetHeight - 90)];
      },
      rect(el) {
        const r = el.getBoundingClientRect();
        const w = root.getBoundingClientRect();
        return { left: r.left - w.left, top: r.top - w.top, width: r.width, height: r.height };
      },
    };
  }

  // Replaces [data-tex] / [data-tex-block] nodes with native MathML (rendered in Cambria Math,
  // the same font OneNote uses for its equations).
  async function renderMath(scope) {
    const nodes = scope.querySelectorAll('[data-tex], [data-tex-block]');
    if (!nodes.length) return;
    await window.MathJax.startup.promise;
    for (const node of nodes) {
      const display = node.hasAttribute('data-tex-block');
      const tex = node.getAttribute(display ? 'data-tex-block' : 'data-tex');
      node.innerHTML = window.MathJax.tex2mml(tex, { display });
    }
  }

  async function renderMermaid(scope) {
    const nodes = scope.querySelectorAll('[data-mermaid]');
    if (!nodes.length) return;
    window.mermaid.initialize({ startOnLoad: false, theme: 'default' });
    let i = 0;
    for (const node of nodes) {
      const { svg } = await window.mermaid.render('mmd' + (i++), node.getAttribute('data-mermaid'));
      node.innerHTML = svg;
      const el = node.querySelector('svg');
      const width = node.getAttribute('data-width');
      if (width) {
        el.style.maxWidth = 'none';
        el.setAttribute('width', width);
        el.removeAttribute('height');
      }
    }
  }

  const outerHeight = el => {
    const cs = getComputedStyle(el);
    return el.offsetHeight + parseFloat(cs.marginTop) + parseFloat(cs.marginBottom);
  };

  // Converts .blk blocks in place (source -> output, or back when reverse is set) while a
  // soft highlight band sweeps over them. Block heights are tweened so the text below
  // reflows smoothly. Call with blocks in their pre-morph layout.
  function morph(tl, body, blocks, start, { reverse = false, sweep = 0.9 } = {}) {
    const base = body.getBoundingClientRect().top;
    const top = blocks[0].getBoundingClientRect().top - base;
    const bottom = blocks[blocks.length - 1].getBoundingClientRect().bottom - base;
    const span = Math.max(bottom - top, 1);

    const band = document.createElement('div');
    band.className = 'shimmer';
    band.style.top = top - 70 + 'px';
    body.appendChild(band);
    tl.to(band, start, sweep, { y: [0, span * 1.35 + 140] }, 'inOutSine');
    tl.to(band, start, 0.12, { opacity: [0, 1] }, 'linear');
    tl.to(band, start + sweep - 0.12, 0.12, { opacity: [1, 0] }, 'linear');

    for (const blk of blocks) {
      const src = blk.querySelector(':scope > .src');
      const out = blk.querySelector(':scope > .out');
      const [from, to] = reverse ? [out, src] : [src, out];
      const at = start + sweep * 0.8 * ((blk.getBoundingClientRect().top - base - top) / span);
      tl.to(from, at, 0.2, { opacity: [1, 0], blur: [0, 1.5] }, 'outCubic');
      tl.to(to, 0, 0, { y: [0, 0], blur: [0, 0] }, 'linear');
      tl.to(to, at + 0.1, 0.5, { opacity: [0, 1], y: [6, 0], blur: [2, 0] }, 'outCubic');
      tl.to(blk, at, 0.55, { height: [outerHeight(from), outerHeight(to)] }, 'inOutCubic');
    }
    return start + sweep + 0.5;
  }

  // Blinking text caret while time falls inside one of the [from, to) windows.
  function caret(tl, el, windows) {
    tl.hook(t => {
      const w = windows.find(([a, b]) => t >= a && t < b);
      el.style.opacity = w && Math.floor((t - w[0]) / 0.53) % 2 === 0 ? '1' : '0';
    });
  }

  // Clones <template id=...> into the note container.
  function mountNote(ui, templateId) {
    const note = document.getElementById(templateId).content.firstElementChild.cloneNode(true);
    ui.body.appendChild(note);
    return note;
  }

  // Glides the pointer onto a ribbon button icon and clicks it at time `at`.
  function pressButton(tl, pointer, ui, btn, at) {
    const [x, y] = ui.center(btn.querySelector('svg'), 4, 5);
    const hover = btn.querySelector('.rbtn-hl');
    pointer.move(at - 1.1, 0.9, x, y);
    tl.to(hover, at - 0.4, 0.14, { opacity: [0, 1] }, 'outCubic');
    pointer.click(at);
    tl.set(btn, at, 'pressed').set(btn, at + 0.15, 'pressed', false);
    tl.to(btn, at, 0.08, { scale: [1, 0.96] }, 'outCubic');
    tl.to(btn, at + 0.08, 0.3, { scale: [0.96, 1] }, 'outBack');
    tl.to(hover, at + 0.6, 0.22, { opacity: [1, 0] }, 'inOutSine');
  }

  // Convert, hold, Reverse Convert, hold: the loop shared by the round-trip scenes.
  function roundTrip(ui, note, { hold = 2.7 } = {}) {
    const blocks = [...note.querySelectorAll('.blk')];
    const tl = new window.Demo.Timeline();
    const rest = ui.restPoint(70);
    const pointer = new window.Demo.Pointer(tl, ui.overlay, rest[0], rest[1]).ibeamWithin(ui.rect(ui.canvas));

    const tc = 1.8;
    pressButton(tl, pointer, ui, ui.buttons.convert, tc);
    const converted = morph(tl, ui.body, blocks, tc + 0.2);
    pointer.move(tc + 0.55, 0.9, rest[0] - 40, rest[1] - 60, -0.1);

    const tr = converted + hold;
    pressButton(tl, pointer, ui, ui.buttons.reverse, tr);
    const reversed = morph(tl, ui.body, blocks, tr + 0.2, { reverse: true });
    pointer.move(tr + 0.55, 1.0, rest[0], rest[1], 0.1);

    // End on a full caret blink cycle so the loop restarts seamlessly.
    const caretBack = reversed - 0.3;
    tl.duration = caretBack + 1.06 * 2;
    caret(tl, note.querySelector('.caret'), [[0, tc + 0.1], [caretBack, tl.duration + 1]]);
    return tl;
  }

  // Shared bootstrap: ?t=1.5 freezes one moment, ?capture leaves seeking to render.py,
  // and a plain open autoplays in a loop for quick previews.
  async function start(build) {
    await Promise.all([document.fonts.ready, window.Demo.icons.ready]);
    const tl = await build();
    window.__demo = { duration: tl.duration, seek: t => tl.seek(t) };
    const params = new URLSearchParams(location.search);
    if (params.has('t')) {
      tl.seek(parseFloat(params.get('t')));
    } else if (params.has('capture')) {
      tl.seek(0);
    } else {
      const t0 = performance.now();
      const loop = now => {
        tl.seek(((now - t0) / 1000) % tl.duration);
        requestAnimationFrame(loop);
      };
      requestAnimationFrame(loop);
    }
    window.__ready = true;
  }

  Object.assign(window.Demo, { buildWindow, renderMath, renderMermaid, outerHeight, morph, caret, mountNote, pressButton, roundTrip, start });
})();
