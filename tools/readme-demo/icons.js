// Icons traced from OneNote for Microsoft 365 screenshots. Coordinates are in the 150%-DPI
// pixel space they were measured in; each viewBox is shown at 1/1.5 of its size in CSS px.
(function () {
  'use strict';

  const INK = '#3b3a39';
  const LINE = `fill="none" stroke="${INK}" stroke-width="1.5" stroke-linecap="butt" stroke-linejoin="miter"`;

  const svg = (viewBox, size, body) => {
    const [, , w, h] = viewBox.split(' ').map(Number);
    const width = (size || w / 1.5).toFixed(2);
    const height = ((size || w / 1.5) * h / w).toFixed(2);
    return `<svg width="${width}" height="${height}" viewBox="${viewBox}">${body}</svg>`;
  };

  const GEAR =
    'M27.58 26.42L30.10 25.78A8.2 8.2 0 0 0 30.10 23.22L27.58 22.58A5.9 5.9 0 0 0 26.45 20.63L27.16 18.13' +
    'A8.2 8.2 0 0 0 24.94 16.84L23.13 18.71A5.9 5.9 0 0 0 20.87 18.71L19.06 16.84A8.2 8.2 0 0 0 16.84 18.13' +
    'L17.55 20.63A5.9 5.9 0 0 0 16.42 22.58L13.90 23.22A8.2 8.2 0 0 0 13.90 25.78L16.42 26.42' +
    'A5.9 5.9 0 0 0 17.55 28.37L16.84 30.87A8.2 8.2 0 0 0 19.06 32.16L20.87 30.29A5.9 5.9 0 0 0 23.13 30.29' +
    'L24.94 32.16A8.2 8.2 0 0 0 27.16 30.87L26.45 28.37A5.9 5.9 0 0 0 27.58 26.42Z';

  // Ribbon button glyphs (imageMso Repeat, Undo, ComAddInsDialog) at the simplified-ribbon size.
  const convert = (size = 20) => svg('6 3 30 30', size,
    `<path ${LINE} d="M25 5.5A13 13 0 1 1 15 6.6"/><path ${LINE} d="M8 5H17V15"/>`);

  const reverse = (size = 20, color = INK) => svg('5 2.5 30 30', size,
    `<g fill="none" stroke="${color}" stroke-width="1.5"><path d="M9 4.5V16H20"/>` +
    '<path d="M9.6 15.4L16.9 7.67A8.25 8.25 0 0 1 28.6 19.33L17.6 30.6"/></g>');

  const settings = (size = 20) => svg('3 3.5 30 30', size,
    `<path ${LINE} d="M14.5 31H6V5H26V14"/>` +
    '<rect x="10.9" y="10.9" width="3.2" height="3.2" fill="none" stroke="#797774" stroke-width="1.8"/>' +
    '<path d="M16 13H22" stroke="#797774" stroke-width="1.5"/>' +
    `<path d="${GEAR}" fill="#fff" stroke="#1e8bcd" stroke-width="1.5" stroke-linejoin="round"/>` +
    '<circle cx="22" cy="24.5" r="2.1" fill="none" stroke="#1e8bcd" stroke-width="1.5"/>');

  // Title bar: render.py extracts the official icon from the local ONENOTE.EXE into the ignored
  // .cache folder (it is never committed); the traced version is used when it is missing.
  const APP_ICON = '.cache/onenote.png';
  const probe = new Image();
  probe.src = APP_ICON;
  let hasAppIcon = false;
  const ready = probe.decode().then(() => { hasAppIcon = true; }, () => {});

  const app = () => (hasAppIcon ? `<img class="app-logo" src="${APP_ICON}" alt="">` : tracedApp());

  const tracedApp = () => svg('0 0 30 30', 20,
    '<rect x="8" y="2.5" width="20" height="23" rx="2.2" fill="#d66eff"/>' +
    '<path d="M18.5 10.5H28V18.5H18.5Z" fill="#a73ad8"/><path d="M18.5 10.5H20V18.5H18.5Z" fill="#c55bf1"/>' +
    '<path d="M8 18.5H28V23.3A2.2 2.2 0 0 1 25.8 25.5H10.2A2.2 2.2 0 0 1 8 23.3Z" fill="#7a16a5"/>' +
    '<rect x="4" y="10" width="14.5" height="14.5" rx="2" fill="#630395"/>' +
    '<path d="M8.1 21.2V13.4H9.8L13 18.3V13.4H14.6V21.2H12.9L9.7 16.3V21.2Z" fill="#fff"/>');

  const qatUndo = () => reverse(16, '#a19f9d');
  const qatRedo = () => reverse(16, '#a19f9d').replace('<g ', '<g transform="matrix(-1 0 0 1 40 0)" ');

  const qatPrint = () => svg('0 0 30 30', 16,
    `<g stroke="${INK}" stroke-width="1.3">` +
    '<rect x="9.5" y="3.5" width="11" height="7.5" fill="#f0f0f0"/>' +
    '<rect x="3.5" y="11.5" width="23" height="9" fill="#fafafa"/>' +
    '<rect x="9.5" y="18.5" width="11" height="8" fill="#f0f0f0"/></g>' +
    `<rect x="5.6" y="13.6" width="1.4" height="1.4" fill="${INK}"/>`);

  const qatMore = () => svg('0 0 30 30', 16,
    `<rect x="10" y="11" width="10" height="2" fill="${INK}"/>` +
    `<path d="M12 15.3L15 18.6L18 15.3" fill="none" stroke="${INK}" stroke-width="1.5"/>`);

  // Notebook bar and page list.
  const notebook = () => svg('0 0 36 36', 24,
    '<path fill="#029dd4" fill-rule="evenodd" d="M8.5 2H26A2.5 2.5 0 0 1 28.5 4.5V29.5A2.5 2.5 0 0 1 26 32H8.5' +
    'A2.5 2.5 0 0 1 6 29.5V4.5A2.5 2.5 0 0 1 8.5 2ZM11.5 7.5H22.5V11.5H11.5Z"/>' +
    '<g fill="#029dd4"><rect x="30" y="9.5" width="2.6" height="4"/><rect x="30" y="15.5" width="2.6" height="4"/>' +
    '<rect x="30" y="21.5" width="2.6" height="4"/></g>');

  const addPage = () => svg('2 2 26 26', 17,
    '<g fill="none" stroke="#7719aa" stroke-width="1.5">' +
    '<path d="M17.5 6.75H8.5A1.75 1.75 0 0 0 6.75 8.5V20.75A1.75 1.75 0 0 0 8.5 22.5H20.75A1.75 1.75 0 0 0 22.5 20.75V12"/>' +
    '<path d="M12.5 17.5L24.5 5"/></g>');

  const sort = () => svg('0 0 30 30', 18,
    `<g fill="none" stroke="${INK}" stroke-width="1.5"><path d="M10 4V23"/><path d="M5 18.5L10 23.5L15 18.5"/>` +
    '<path d="M14 6H25M14 10H22M14 14.75H18"/></g>');

  const chevronDown = (size = 12) => svg('0 0 16 16', size,
    `<path d="M3.5 6L8 10.5L12.5 6" fill="none" stroke="${INK}" stroke-width="1.6"/>`);

  const search = () => svg('0 0 16 16', 14,
    `<g fill="none" stroke="${INK}" stroke-width="1.2"><circle cx="6.5" cy="6.5" r="4.5"/><path d="M10 10L14 14"/></g>`);

  const win = {
    min: '<svg width="10" height="10"><path d="M0 5.5H10" stroke="#3b3a39"/></svg>',
    max: '<svg width="10" height="10"><rect x=".5" y=".5" width="9" height="9" fill="none" stroke="#3b3a39"/></svg>',
    close: '<svg width="10" height="10"><path d="M0 0L10 10M10 0L0 10" stroke="#3b3a39"/></svg>',
  };

  window.Demo = window.Demo || {};
  window.Demo.icons = {
    ready,
    convert, reverse, settings, app, qatUndo, qatRedo, qatPrint, qatMore,
    notebook, addPage, sort, chevronDown, search, win,
  };
})();
