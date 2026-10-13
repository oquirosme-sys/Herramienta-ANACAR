/* Ilustraciones vectoriales de equipos eléctricos para el glosario y el catálogo (módulo Ayuda).
   Cada función devuelve un SVG (texto) de 240 × 180. Se usan en la web y se convierten a PNG para el Excel. */
(function (g) {
  'use strict';
  const C = { caja: '#cfd4d9', borde: '#4f565d', frente: '#e9ecef', bk: '#30363b', mango: '#f4f6f8', verde: '#2e7a3a', cobre: '#c8793a', texto: '#33393f' };
  const svg = (cuerpo, titulo) => '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 240 180" role="img" aria-label="' + (titulo || '') + '">' +
    '<rect width="240" height="180" fill="#fff"/>' + cuerpo + '</svg>';
  const t = (x, y, s, sz, anc, col) => '<text x="' + x + '" y="' + y + '" font-family="Century Gothic,Arial,sans-serif" font-size="' + (sz || 9) + '" fill="' + (col || C.texto) + '" text-anchor="' + (anc || 'middle') + '">' + s + '</text>';
  const r = (x, y, w, h, f, s, rx, sw) => '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" fill="' + f + '" stroke="' + (s || 'none') + '" rx="' + (rx || 0) + '" stroke-width="' + (sw || 1) + '"/>';
  const l = (x1, y1, x2, y2, s, w, dash) => '<line x1="' + x1 + '" y1="' + y1 + '" x2="' + x2 + '" y2="' + y2 + '" stroke="' + (s || C.borde) + '" stroke-width="' + (w || 1) + '"' + (dash ? ' stroke-dasharray="' + dash + '"' : '') + '/>';
  const c = (x, y, rr, f, s, sw) => '<circle cx="' + x + '" cy="' + y + '" r="' + rr + '" fill="' + f + '" stroke="' + (s || 'none') + '" stroke-width="' + (sw || 1) + '"/>';

  /** Breaker enchufable visto de frente (ancho w por polo). */
  function bkFrente(x, y, polos, w, h, txt) {
    let s = '';
    for (let i = 0; i < polos; i++) s += r(x + i * w, y, w - 1, h, C.bk, null, 1) + r(x + i * w + w * 0.2, y + h * 0.3, w * 0.6 - 1, h * 0.4, C.mango, null, 1);
    if (polos > 1) s += l(x + w * 0.3, y + h * 0.5, x + (polos - 0.3) * w, y + h * 0.5, C.mango, 2);
    if (txt) s += t(x + polos * w / 2, y + h + 9, txt, 8);
    return s;
  }

  /** Tablero (panelboard): gabinete, frente muerto, principal arriba y dos columnas de ramales. */
  function tablero(o) {
    o = o || {};
    const filas = o.filas || 9, zap = o.zapatas, x = 70, y = 8, w = 100, h = o.etq ? 152 : 164;
    let s = r(x, y, w, h, C.caja, C.borde, 4, 1.5) + r(x + 8, y + 8, w - 16, h - 16, C.frente, C.borde, 2);
    if (zap) s += [0, 1, 2].map(i => c(x + 32 + i * 18, y + 22, 5, C.cobre, C.borde)).join('') + t(x + w / 2, y + 38, 'ZAPATAS', 7);
    else s += bkFrente(x + 32, y + 14, 3, 12, 22) + t(x + w / 2, y + 47, 'PRINCIPAL', 7);
    for (let i = 0; i < filas; i++) { const yy = y + 52 + i * 10; s += r(x + 14, yy, 30, 8, C.bk, null, 1) + r(x + 22, yy + 2, 14, 4, C.mango) + r(x + w - 44, yy, 30, 8, C.bk, null, 1) + r(x + w - 36, yy + 2, 14, 4, C.mango); }
    s += l(x + w / 2, y + 50, x + w / 2, y + 52 + filas * 10, C.cobre, 3);
    s += t(x + 8, y + 58, '1', 6, 'start') + t(x + w - 8, y + 58, '2', 6, 'end');
    if (o.etq) s += t(120, 178, o.etq, 9, 'middle', C.verde);
    return svg(s, o.etq || 'Tablero');
  }
  /** Centro de carga (residencial): gabinete pequeño con puerta. */
  function centroCarga() {
    let s = r(80, 20, 80, 130, C.caja, C.borde, 4, 1.5) + r(86, 26, 68, 118, C.frente, C.borde, 2) + bkFrente(108, 32, 2, 12, 18);
    for (let i = 0; i < 6; i++) s += r(92, 60 + i * 12, 24, 9, C.bk, null, 1) + r(124, 60 + i * 12, 24, 9, C.bk, null, 1);
    s += r(158, 70, 4, 26, C.borde, null, 1) + t(120, 168, 'Centro de carga', 9, 'middle', C.verde);
    return svg(s, 'Centro de carga');
  }
  /** Subestación / switchboard: varias secciones de piso. */
  function subestacion() {
    let s = '';
    for (let i = 0; i < 3; i++) {
      const x = 20 + i * 68;
      s += r(x, 10, 66, 150, C.caja, C.borde, 2, 1.5) + r(x + 5, 16, 56, 40, C.frente, C.borde, 1) + r(x + 5, 60, 56, 94, C.frente, C.borde, 1);
      if (i === 0) s += r(x + 14, 22, 38, 28, C.bk, null, 2) + r(x + 26, 30, 14, 12, C.mango) + t(x + 33, 68, 'PRINCIPAL', 7);
      else for (let j = 0; j < 3; j++) s += r(x + 12, 68 + j * 28, 42, 22, C.bk, null, 2) + r(x + 27, 74 + j * 28, 12, 10, C.mango);
    }
    s += r(18, 160, 206, 6, C.borde) + t(120, 177, 'Subestación / switchboard', 9, 'middle', C.verde);
    return svg(s, 'Subestación');
  }
  /** Zapatas frente a interruptor principal. */
  function principal() {
    let s = r(15, 20, 95, 120, C.caja, C.borde, 3, 1.5) + r(130, 20, 95, 120, C.caja, C.borde, 3, 1.5);
    s += [0, 1, 2].map(i => r(34 + i * 20, 40, 14, 22, C.cobre, C.borde, 2) + c(41 + i * 20, 51, 3, C.borde)).join('') + l(41, 62, 41, 130, C.cobre, 3) + l(61, 62, 61, 130, C.cobre, 3) + l(81, 62, 81, 130, C.cobre, 3);
    s += bkFrente(155, 36, 3, 15, 30) + l(162, 66, 162, 130, C.cobre, 3) + l(177, 66, 177, 130, C.cobre, 3) + l(192, 66, 192, 130, C.cobre, 3);
    s += t(62, 158, 'Zapatas (main lugs)', 9) + t(177, 158, 'Interruptor principal', 9);
    return svg(s, 'Zapatas e interruptor principal');
  }
  /** Breakers enchufables 1P, 2P y 3P. */
  function enchufable(o) {
    o = o || {};
    let s = bkFrente(25, 50, 1, 26, 70, '1 polo') + bkFrente(80, 50, 2, 26, 70, '2 polos') + bkFrente(160, 50, 3, 22, 70, '3 polos');
    s += t(120, 30, o.etq || 'Enchufables / atornillables (BAB, QBHW, GHB, QOB)', 9, 'middle', C.verde);
    return svg(s, 'Breakers enchufables');
  }
  /** Interruptor de caja moldeada. */
  function cajaMoldeada(o) {
    o = o || {};
    let s = r(70, 20, 100, 140, C.bk, C.borde, 6, 1.5);
    for (let i = 0; i < 3; i++) s += r(82 + i * 28, 26, 20, 14, '#8a9096', null, 2) + r(82 + i * 28, 140, 20, 14, '#8a9096', null, 2);
    s += r(100, 70, 40, 40, '#4a5157', null, 4) + r(110, 76, 20, 28, C.mango, null, 3);
    if (o.electronico) s += r(84, 116, 72, 18, '#1d4d2a', null, 2) + t(120, 128, 'L  S  I  G', 9, 'middle', '#c7f0d0');
    else s += t(120, 128, o.marco || 'MARCO 250 A', 8, 'middle', '#e9ecef');
    s += t(120, 176, o.etq || 'Caja moldeada (MCCB)', 9, 'middle', C.verde);
    return svg(s, 'Interruptor de caja moldeada');
  }
  /** GFCI / AFCI con botón de prueba y cable neutro. */
  function gfci() {
    let s = r(95, 30, 50, 110, C.bk, C.borde, 3) + r(105, 60, 30, 34, C.mango, null, 2) + c(120, 112, 7, '#d9e86b', C.borde) + t(120, 115, 'T', 8);
    s += '<path d="M120 140 C 120 160, 160 150, 175 165" stroke="#f2f2f2" stroke-width="5" fill="none"/><path d="M120 140 C 120 160, 160 150, 175 165" stroke="#9aa0a6" stroke-width="1" fill="none"/>';
    s += t(178, 160, 'neutro', 8, 'start') + t(160, 112, 'prueba (TEST)', 8, 'start') + t(120, 22, 'GFCI / AFCI', 10, 'middle', C.verde);
    return svg(s, 'GFCI / AFCI');
  }
  /** Supresor de picos. */
  function spd() {
    let s = r(75, 30, 90, 110, '#dfe3e6', C.borde, 6, 1.5);
    for (let i = 0; i < 3; i++) s += c(100 + i * 20, 60, 6, '#4cd964', C.borde) + t(100 + i * 20, 80, 'ABC'[i], 9);
    s += t(120, 105, 'SPD', 16, 'middle', C.verde) + t(120, 122, '250 kA', 9);
    for (let i = 0; i < 4; i++) s += l(95 + i * 16, 140, 95 + i * 16, 165, ['#111', '#c00', '#1565c0', '#2e7d32'][i], 3);
    return svg(s, 'Supresor de picos');
  }
  /** Transformador seco. */
  function transformador() {
    let s = r(55, 25, 130, 125, C.caja, C.borde, 4, 1.5);
    for (let i = 0; i < 7; i++) s += r(70 + i * 15, 45, 8, 50, '#a9b0b6', null, 2);
    s += r(80, 105, 80, 30, C.frente, C.borde, 2) + t(120, 118, 'kVA · Z %', 9) + t(120, 130, '480 / 208Y-120', 8);
    s += r(50, 150, 140, 8, C.borde) + t(120, 174, 'Transformador seco', 9, 'middle', C.verde);
    return svg(s, 'Transformador');
  }
  /** Transferencia automática / manual. */
  function transferencia() {
    let s = r(70, 15, 100, 150, C.caja, C.borde, 3, 1.5) + r(85, 28, 70, 26, '#1d2b36', null, 2) + t(120, 45, 'NORMAL  ▸', 9, 'middle', '#7fe08f');
    s += r(95, 70, 50, 50, C.bk, null, 4) + r(112, 78, 16, 34, C.mango, null, 3);
    s += l(10, 40, 70, 40, '#1565c0', 3) + t(12, 34, 'Normal', 8, 'start') + l(10, 100, 70, 100, '#c62828', 3, '6 4') + t(12, 94, 'Emergencia', 8, 'start');
    s += l(170, 70, 230, 70, C.borde, 3) + t(228, 64, 'Carga', 8, 'end') + t(120, 178, 'ATS / MTS', 9, 'middle', C.verde);
    return svg(s, 'Transferencia');
  }
  /** UPS. */
  function ups() {
    let s = r(85, 15, 70, 150, '#2f3438', C.borde, 5, 1.5) + r(95, 28, 50, 24, '#173b22', null, 2) + t(120, 44, '100 %', 10, 'middle', '#7fe08f');
    for (let i = 0; i < 6; i++) s += r(95, 64 + i * 12, 50, 6, '#454c52', null, 1);
    s += t(120, 178, 'UPS', 9, 'middle', C.verde);
    return svg(s, 'UPS');
  }
  /** Generador. */
  function generador() {
    let s = r(30, 50, 180, 90, '#e2b23a', C.borde, 6, 1.5) + r(40, 60, 50, 70, '#c9971f', null, 3);
    for (let i = 0; i < 6; i++) s += l(46, 66 + i * 11, 84, 66 + i * 11, '#7a5b10', 2);
    s += r(110, 70, 80, 40, '#f4d27a', C.borde, 3) + t(150, 95, "G  X''d", 11);
    s += r(25, 140, 190, 10, C.borde) + t(120, 172, 'Generador', 9, 'middle', C.verde);
    return svg(s, 'Generador');
  }
  /** Conductores: corte transversal y tipos de aislamiento. */
  function conductores() {
    const col = ['#111', '#c62828', '#1565c0', '#f5f5f5', '#2e7d32'], nom = ['A', 'B', 'C', 'N', 'T'];
    let s = '';
    for (let i = 0; i < 5; i++) s += c(40 + i * 40, 70, 16, col[i], C.borde, 1.5) + c(40 + i * 40, 70, 9, C.cobre, '#8a4b1c') + t(40 + i * 40, 104, nom[i], 10);
    s += t(120, 30, 'THHN · XHHW-2 · RHW (CU / AL)', 10, 'middle', C.verde) + t(120, 130, 'Calibre AWG / kcmil · paralelos: 2x, 3x…', 9);
    s += t(120, 150, '3#2/0 + 1#2/0 N + 1#4 T', 10);
    return svg(s, 'Conductores');
  }
  /** Tubería EMT / PVC y ductobarra. */
  function tuberia() {
    let s = r(20, 40, 200, 26, '#b9c0c7', C.borde, 13) + r(110, 36, 22, 34, '#9aa2aa', C.borde, 4) + t(60, 57, 'EMT', 10);
    s += r(20, 90, 200, 26, '#9e9e9e', C.borde, 13) + t(60, 107, 'PVC SCH40', 10, 'middle', '#fff');
    s += r(20, 132, 200, 20, '#8d6e63', C.borde, 2) + t(120, 146, 'Ductobarra', 9, 'middle', '#fff') + t(120, 172, 'Canalizaciones', 9, 'middle', C.verde);
    return svg(s, 'Tubería');
  }
  /** Barras del tablero. */
  function barras() {
    let s = '';
    ['A', 'B', 'C'].forEach((x, i) => { s += r(40, 30 + i * 26, 160, 12, C.cobre, '#8a4b1c', 2) + t(30, 40 + i * 26, x, 10); });
    s += r(40, 110, 160, 10, '#d7d7d7', C.borde, 2) + t(30, 119, 'N', 10) + r(40, 132, 160, 10, '#7cb342', C.borde, 2) + t(30, 141, 'T', 10);
    s += t(120, 172, 'Barras de fase, neutro y tierra (A)', 9, 'middle', C.verde);
    return svg(s, 'Barras');
  }
  /** Factores de demanda y diversidad (barras). */
  function factores() {
    const b = [['Conectada', 120, '#90a4ae'], ['× F. demanda', 90, '#4f9a5b'], ['× F. diversidad', 72, '#2e7a3a']];
    let s = '';
    b.forEach((x, i) => { s += r(30 + i * 65, 150 - x[1], 45, x[1], x[2], null, 2) + t(52 + i * 65, 165, x[0], 8); });
    s += t(120, 18, 'kVA demandados = kVA × FD × Fdiv (0 – 1)', 9, 'middle', C.verde);
    return svg(s, 'Factores');
  }
  /** Caída de voltaje y cortocircuito en cascada. */
  function cascada() {
    let s = l(20, 40, 220, 40, C.borde, 2) + t(20, 32, 'Acometida', 8, 'start');
    [60, 130, 200].forEach((x, i) => { s += l(x, 40, x, 80, C.borde, 2) + r(x - 22, 80, 44, 10, C.cobre) + t(x, 104, 'T' + (i + 1), 9); });
    s += '<polyline points="20,130 60,136 130,146 200,158" fill="none" stroke="#c62828" stroke-width="2"/>' + t(222, 172, 'ΔV', 9, 'end', '#c62828');
    s += '<polyline points="20,170 60,160 130,150 200,142" fill="none" stroke="#1565c0" stroke-width="2"/>' + t(222, 134, 'Icc', 9, 'end', '#1565c0');
    return svg(s, 'Cascada');
  }

  const LISTA = { tablero, centroCarga, subestacion, principal, enchufable, cajaMoldeada, gfci, spd, transformador, transferencia, ups, generador, conductores, tuberia, barras, factores, cascada };
  g.Ilus = Object.assign({ get: (k, o) => (LISTA[k] ? LISTA[k](o) : '') }, LISTA);
})(window);
