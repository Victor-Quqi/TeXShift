// Deterministic timeline engine for the README demo scenes.
// Every animated value is a pure function of time, so render.py can seek
// frame by frame and the captured GIF is identical on every run.
(function () {
  'use strict';

  const ease = {
    linear: p => p,
    inCubic: p => p * p * p,
    outCubic: p => 1 - Math.pow(1 - p, 3),
    inOutCubic: p => (p < 0.5 ? 4 * p * p * p : 1 - Math.pow(-2 * p + 2, 3) / 2),
    outQuint: p => 1 - Math.pow(1 - p, 5),
    inOutQuint: p => (p < 0.5 ? 16 * p * p * p * p * p : 1 - Math.pow(-2 * p + 2, 5) / 2),
    inOutSine: p => -(Math.cos(Math.PI * p) - 1) / 2,
    outBack: p => {
      const c1 = 1.4, c3 = c1 + 1;
      return 1 + c3 * Math.pow(p - 1, 3) + c1 * Math.pow(p - 1, 2);
    },
  };

  const lerp = (a, b, p) => a + (b - a) * p;
  const clamp01 = v => Math.min(1, Math.max(0, v));

  let nextId = 0;
  const idOf = el => el.__tlId || (el.__tlId = 'e' + (++nextId));

  // Composite channels are merged into one transform/filter per element.
  const TRANSFORM = new Set(['x', 'y', 'scale', 'scaleX', 'scaleY', 'rotate']);
  const FX_DEFAULTS = { x: 0, y: 0, scale: 1, scaleX: 1, scaleY: 1, rotate: 0, blur: 0 };

  class Timeline {
    constructor() {
      this.groups = new Map();
      this.hooks = [];
      this.touched = new Set();
      this.duration = 0;
    }

    // Registers one keyed tween; at seek time the latest started tween per key wins.
    add(key, start, dur, apply, easing) {
      const fn = typeof easing === 'function' ? easing : ease[easing || 'inOutCubic'];
      if (!this.groups.has(key)) this.groups.set(key, []);
      this.groups.get(key).push({ start, dur, apply, fn });
      this.duration = Math.max(this.duration, start + dur);
      return this;
    }

    // Tweens element channels: { opacity: [0, 1], y: [8, 0], height: [10, 40] }.
    to(el, start, dur, props, easing) {
      for (const [prop, [from, to]] of Object.entries(props)) {
        this.add(idOf(el) + ':' + prop, start, dur, p => this.write(el, prop, lerp(from, to, p)), easing);
      }
      return this;
    }

    // Toggles a class at an instant; before the first toggle the opposite state holds.
    set(el, time, cls, on = true) {
      return this.add(idOf(el) + '.' + cls, time, 0, p => el.classList.toggle(cls, p >= 1 ? on : !on), 'linear');
    }

    // Runs fn(t) on every seek, for effects that are not simple tweens (caret blink, typing).
    hook(fn) {
      this.hooks.push(fn);
      return this;
    }

    write(el, prop, value) {
      if (prop in FX_DEFAULTS) {
        el.__fx = el.__fx || Object.assign({}, FX_DEFAULTS);
        el.__fx[prop] = value;
        this.touched.add(el);
      } else if (prop === 'opacity') {
        el.style.opacity = value.toFixed(4);
      } else {
        el.style[prop] = value.toFixed(2) + 'px';
      }
    }

    seek(t) {
      for (const tweens of this.groups.values()) {
        let active = null;
        for (const tw of tweens) {
          if (tw.start <= t && (!active || tw.start >= active.start)) active = tw;
        }
        if (active) {
          const p = active.dur === 0 ? 1 : clamp01((t - active.start) / active.dur);
          active.apply(active.fn(p));
        } else {
          const first = tweens.reduce((a, b) => (b.start < a.start ? b : a));
          first.apply(first.fn(0));
        }
      }
      for (const fn of this.hooks) fn(t);
      for (const el of this.touched) {
        const f = el.__fx;
        const parts = [];
        if (f.x || f.y) parts.push(`translate(${f.x.toFixed(2)}px, ${f.y.toFixed(2)}px)`);
        if (f.rotate) parts.push(`rotate(${f.rotate.toFixed(2)}deg)`);
        if (f.scale !== 1) parts.push(`scale(${f.scale.toFixed(4)})`);
        if (f.scaleX !== 1 || f.scaleY !== 1) parts.push(`scale(${f.scaleX.toFixed(4)}, ${f.scaleY.toFixed(4)})`);
        el.style.transform = parts.join(' ');
        el.style.filter = f.blur > 0.01 ? `blur(${f.blur.toFixed(2)}px)` : '';
      }
    }
  }

  // Mouse pointer that glides along gently curved paths and shows click ripples.
  class Pointer {
    constructor(tl, layer, x, y) {
      this.tl = tl;
      this.layer = layer;
      this.x = x;
      this.y = y;
      this.el = document.createElement('div');
      this.el.className = 'pointer';
      this.el.innerHTML = POINTER_SVG.arrow + POINTER_SVG.ibeam;
      layer.appendChild(this.el);
      this.moves = 0;
      tl.add('pointer:pos', 0, 0, () => this.place(x, y), 'linear');
    }

    place(x, y) {
      this.tl.write(this.el, 'x', x);
      this.tl.write(this.el, 'y', y);
    }

    // Moves to (x, y); arc bends the path sideways for a natural hand motion.
    move(start, dur, x, y, arc = 0.12, easing = 'inOutCubic') {
      const x0 = this.x, y0 = this.y;
      const dx = x - x0, dy = y - y0;
      const bend = Math.hypot(dx, dy) * arc;
      const nx = -dy / (Math.hypot(dx, dy) || 1), ny = dx / (Math.hypot(dx, dy) || 1);
      this.tl.add('pointer:pos', start, dur, p => {
        const b = Math.sin(Math.PI * p) * bend;
        this.place(lerp(x0, x, p) + nx * b, lerp(y0, y, p) + ny * b);
      }, easing);
      this.x = x;
      this.y = y;
      return this;
    }

    // Shows the text I-beam whenever the pointer is inside rect, as over a OneNote page.
    ibeamWithin(rect) {
      this.tl.hook(() => {
        const { x, y } = this.el.__fx;
        const inside = x >= rect.left && x <= rect.left + rect.width && y >= rect.top && y <= rect.top + rect.height;
        this.el.classList.toggle('ibeam', inside);
      });
      return this;
    }

    // Press feedback: pointer dips slightly and a soft ring expands from the tip.
    click(time) {
      this.tl.to(this.el, time, 0.09, { scale: [1, 0.86] }, 'outCubic');
      this.tl.to(this.el, time + 0.09, 0.18, { scale: [0.86, 1] }, 'outCubic');
      const ring = document.createElement('div');
      ring.className = 'click-ring';
      ring.style.left = this.x + 'px';
      ring.style.top = this.y + 'px';
      this.layer.appendChild(ring);
      this.tl.to(ring, time, 0.06, { opacity: [0, 0.9] }, 'linear');
      this.tl.to(ring, time + 0.06, 0.5, { opacity: [0.9, 0] }, 'outCubic');
      this.tl.to(ring, time, 0.56, { scale: [0.25, 1] }, 'outQuint');
      return this;
    }
  }

  const POINTER_SVG = {
    arrow: '<svg class="pointer-arrow" width="22" height="26" viewBox="0 0 22 26">' +
      '<path d="M2 1.5 L2 20.5 L6.7 16.3 L9.9 23.6 L13.4 22.1 L10.3 15 L16.6 15 Z" ' +
      'fill="#fff" stroke="#111" stroke-width="1.3" stroke-linejoin="round"/></svg>',
    ibeam: '<svg class="pointer-ibeam" width="12" height="22" viewBox="0 0 12 22">' +
      '<path d="M2 1.5 h3 q1 0 1 1 v17 q0 1 -1 1 h-3 M10 1.5 h-3 q-1 0 -1 1 v17 q0 1 1 1 h3" ' +
      'fill="none" stroke="#fff" stroke-width="3.2" stroke-linecap="round"/>' +
      '<path d="M2 1.5 h3 q1 0 1 1 v17 q0 1 -1 1 h-3 M10 1.5 h-3 q-1 0 -1 1 v17 q0 1 1 1 h3" ' +
      'fill="none" stroke="#111" stroke-width="1.2" stroke-linecap="round"/></svg>',
  };

  window.Demo = { Timeline, Pointer, ease, lerp, clamp01 };
})();
