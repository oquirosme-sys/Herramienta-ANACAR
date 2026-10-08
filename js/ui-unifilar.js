/* Diagrama unifilar (estilo ETAP / SKM): red o acometida → transformador → UPS → interruptor → alimentador → barra de cada tablero,
   con segunda acometida (generador o bypass) por ATS / MTS / interruptor. El mismo dibujo se muestra en SVG y se exporta a DXF
   (AutoCAD lo abre directamente y puede guardarlo como DWG). Opcionalmente, una tabla de datos junto a cada tablero (como en los DU). */
(function () {
  'use strict';
  const h = U.h, f2 = v => U.fmt(v, 2), f1 = v => U.fmt(v, 1), f0 = v => U.fmt(v, 0);
  const NS = 'http://www.w3.org/2000/svg';
  const W = 150, GAP = 30;
  let NIVEL = 330;
  let conTabla = false;
  const anchoDatos = () => (conTabla ? 330 : 250);
  const colorV = V => (V >= 440 ? '#c0392b' : V >= 230 ? '#1b7f4b' : '#0b5cab');

  /* ---------- transformador ---------- */
  function editarTrafo(t, R) {
    const tr = t.trafo, r = R.res[t.id], padreV = r && r.padre ? r.padre.V : '';
    const d = { activo: true, kva: tr.kva, z: tr.z, xr: tr.xr, primario: tr.primario || padreV || '', fases: tr.fases || (Number(t.fases) === 3 ? 3 : 1), nombre: tr.nombre || '' };
    const pres = UI.select(UI.opts(Store.catalog.transformadores.map(x => ({ value: x.id, label: x.nombre + ' · Z ' + x.z + ' %', group: x.tipo })), '— Elegir de la lista (opcional) —'), '', v => {
      const x = Store.catalog.transformadores.find(o => String(o.id) === v); if (!x) return;
      const m = /(\d+(?:[.,]\d+)?)\s*KVA/i.exec(x.nombre); if (m) { d.kva = Number(m[1].replace(',', '.')); kva.value = String(d.kva).replace('.', ','); }
      d.z = x.z; z.value = String(x.z).replace('.', ','); d.fases = /MONO/i.test(x.tipo) ? 1 : 3; fs.value = d.fases; d.nombre = d.nombre || x.nombre; nom.value = d.nombre;
    });
    const kva = UI.input(d, 'kva', { type: 'num', save: false, after: () => {} }), z = UI.input(d, 'z', { type: 'num', save: false, after: () => {} });
    const xr = UI.input(d, 'xr', { type: 'num', save: false, after: () => {}, placeholder: '4' }), vp = UI.input(d, 'primario', { type: 'num', save: false, after: () => {} });
    const fs = UI.select([{ value: 3, label: 'Trifásico' }, { value: 1, label: 'Monofásico' }], d.fases, v => { d.fases = Number(v); });
    const nom = UI.input(d, 'nombre', { save: false, after: () => {}, placeholder: 'T-1' });
    UI.modal('Transformador que alimenta a ' + Calc.nombreTablero(t), h('div', null,
      h('p', { class: 'muted' }, 'Se ubica entre ' + (r && r.padre ? r.padre.nombre : 'la red') + ' y el alimentador de ' + Calc.nombreTablero(t) + '. El secundario es el sistema del tablero (' + t.sistema + ' V).'),
      h('div', { class: 'grid2' }, UI.field('De la lista del Excel', pres, null, 'span2'), UI.field('Nombre', nom), UI.field('Fases', fs), UI.field('Capacidad (kVA)', kva), UI.field('Impedancia Z (%)', z),
        UI.field('Relación X/R', xr, 'Para la caída de tensión en el transformador; 4 si no se conoce'), UI.field('Voltaje primario (V)', vp, padreV ? 'Voltaje del tablero que lo alimenta: ' + padreV + ' V' : 'Voltaje de la red'))), [
      tr.activo ? { label: 'Quitar transformador', cls: 'danger', onclick: () => { tr.activo = false; Store.save(); App.refresh(); } } : null,
      { label: 'Cancelar' },
      { label: 'Guardar', cls: 'primary', onclick: () => {
        if (!Number(d.kva) || !Number(d.z)) { UI.alert('Indique la capacidad (kVA) y la impedancia (%).'); return false; }
        Object.assign(tr, d, { activo: true });
        const padre = Store.tablero(t.padreId), circ = padre && padre.circuitos.find(c => c.tableroHijoId === t.id);
        if (circ && Number(padre.voltaje) !== Number(t.voltaje)) circ.detalleId = Store.detalleTablero(t, padre);
        Store.save(); App.refresh();
      } },
    ].filter(Boolean));
  }

  /* ---------- dibujo abstracto (lo usan SVG y DXF) ---------- */
  /** Recorre el proyecto y emite primitivas: linea, rect, circulo, texto (con capa, clase y acción al hacer clic). */
  function dibujar(R, D) {
    const P = Store.project, LBL = anchoDatos();
    NIVEL = conTabla ? 480 : 330;
    const raiz = R.orden.filter(r => r.nivel === 0);
    const span = {}, ancho = {}, hijos = r => (R.hijos[r.tab.id] || []).map(t => R.res[t.id]).filter(Boolean);
    const medir = r => { const hs = hijos(r); span[r.tab.id] = Math.max(W, hs.reduce((a, x) => a + medir(x), 0) + GAP * Math.max(0, hs.length - 1)); ancho[r.tab.id] = span[r.tab.id] + LBL; return ancho[r.tab.id]; };
    const totalW = raiz.reduce((a, r) => a + medir(r) + GAP, GAP) + 40;
    const prof = r => 1 + Math.max(0, ...hijos(r).map(prof));
    D.tamano(totalW, 120 + NIVEL * Math.max(1, ...raiz.map(prof)) + 120);

    const interruptor = (x, y, txt, mal, ir) => {
      D.rect(x - 9, y - 9, 18, 18, 'u-bk' + (mal ? ' mal' : ''), 'PROTECCION', ir);
      D.texto(x + 16, y + 4, txt, 'u-t', 'TEXTO');
    };
    /** Bajada desde (x, y0) hasta la barra en y1. */
    function bajada(x, y0, y1, r, circ) {
      const A = r.alim, tr = r.trafo; let y = y0;
      D.linea(x, y0, x, y1, 'u-lin', 'ALIMENTADOR');
      const bkA = circ ? circ.AF : A.AF139, b = circ ? circ.breaker : r.bkMain;
      y += circ ? (conTabla ? 262 : 96) : 26;
      interruptor(x, y, (bkA || '–') + ' A ' + (b ? (b.polos || '') + 'P ' + (b.unidad === 'STD' ? 'TM' : b.unidad || '') + ' · ' + (b.ref || b.modelo || '') : ''), b && b.sccrBajo, () => App.go('memoria', circ ? r.padre.tab.id : r.tab.id));
      if (tr) {
        y += 58;
        D.circulo(x, y - 11, 15, 'u-trc', 'TRAFO', () => editarTrafo(r.tab, R)); D.circulo(x, y + 11, 15, 'u-trc', 'TRAFO', () => editarTrafo(r.tab, R));
        D.texto(x + 22, y - 12, (tr.nombre ? tr.nombre + ' · ' : '') + f1(tr.kva) + ' kVA · Z ' + f2(tr.z) + ' %', 'u-t b', 'TEXTO');
        D.texto(x + 22, y + 3, (tr.Vp ? f0(tr.Vp) : '?') + ' / ' + r.tab.sistema + ' V · carga ' + f1(tr.carga * 100) + ' %', 'u-t', 'TEXTO');
        D.texto(x + 22, y + 18, 'Icc sec. ' + f2(tr.iccSec / 1000) + ' kA · ΔV ' + f2(tr.reg) + ' %', 'u-t', 'TEXTO');
      } else {
        y += 50;
        D.boton(x, y, '+ trafo', () => editarTrafo(r.tab, R));
      }
      if (r.ups) {
        y += 44;
        D.rect(x - 22, y - 13, 44, 26, 'u-ups', 'EQUIPOS', () => App.go('memoria', r.tab.id)); D.texto(x, y + 4, 'UPS', 'u-t b', 'TEXTO', 'c');
        D.texto(x + 30, y + 4, r.ups.nombre + ' ' + f1(r.ups.kva) + ' kVA · Icc ' + f2(r.ups.icc / 1000) + ' kA', 'u-t', 'TEXTO');
      }
      y = Math.max(y + 46, y1 - (r.alterna ? 120 : 74));
      // con segunda acometida los datos del cable van a la izquierda de la línea
      const cx = r.alterna ? x - 10 : x + 10, al = r.alterna ? 'e' : null;
      D.texto(cx, y, A.fasesTxt + (A.neutroTxt ? '+' + A.neutroTxt + 'N' : '') + (A.tierraTxt ? '+' + A.tierraTxt + 'T' : '') + ' ' + A.mat + ' ' + A.ais, 'u-t b', 'TEXTO', al);
      D.texto(cx, y + 15, f1(A.M139) + ' m · ' + (A.preN || '') + A.AR139 + ' mm ' + A.tuberia + ' · ΔV ' + f2(A.AV139 / r.V * 100) + ' %' + (A.ajustado ? ' · ajustado' : ''), 'u-t', 'TEXTO', al);
      if (A.val && !A.val.ok) D.texto(cx, y + 30, '⚠ no cumple ampacidad corregida', 'u-t mal', 'TEXTO', al);
      // segunda acometida: generador o bypass → ATS / MTS / IP sobre la barra
      if (r.alterna) {
        const ya = y1 - 46, al = r.alterna;
        D.rect(x - 20, ya - 12, 40, 24, 'u-ats', 'EQUIPOS', () => App.go('memoria', r.tab.id)); D.texto(x, ya + 4, al.equipo, 'u-t b', 'TEXTO', 'c');
        const gx = x + 140;
        D.linea(x + 20, ya, gx, ya, 'u-lin2', 'ALIMENTADOR'); D.linea(gx, ya, gx, ya - 70, 'u-lin2', 'ALIMENTADOR');
        if (al.tipo === 'generador') { D.circulo(gx, ya - 86, 16, 'u-gen', 'EQUIPOS', () => App.go('memoria', r.tab.id)); D.texto(gx, ya - 81, 'G', 'u-t b', 'TEXTO', 'c'); }
        else D.texto(gx, ya - 76, '↑', 'u-t b', 'TEXTO', 'c');
        D.texto(gx + 22, ya - 92, al.tipo === 'generador' ? (al.origen || 'Generador') + (al.xd ? " · X''d " + f1(al.xd) + ' %' : '') : 'Bypass desde ' + (al.origen || '?'), 'u-t b', 'TEXTO');
        D.texto(gx + 8, ya - 34, al.fasesTxt + ' ' + al.mat + ' · ' + f1(al.longitud) + ' m · ΔV ' + f2(al.cv) + ' %', 'u-t', 'TEXTO');
        D.texto(gx + 8, ya - 19, 'Icc ' + (al.iccA ? f2(al.iccA / 1000) + ' kA' : '—') + (al.carga !== null ? ' · carga ' + f1(al.carga * 100) + ' %' : ''), 'u-t', 'TEXTO');
      }
    }

    /** Tabla de datos del tablero (como las tablas de los DU). */
    function tablaDatos(r, x, y) {
      const A = r.alim, w1 = 170, w2 = 140, hh = 15;
      const filas = [['EQUIPO', r.nombre], ['CORRIENTE CORTOCIRCUITO', r.iccKA ? f2(r.iccKA) + ' kA' : '—'], ['kVA TOTALES', f2(r.W130)], ['FACTOR DE DEMANDA', f2(A.R139)], ['kVA DEMANDADOS', f2(A.L139)],
        ['FACTOR DE POTENCIA', f2(A.T139)], ['ACOMETIDA', ''], ['FASES', A.fasesTxt], ['NEUTRO', A.neutroTxt || '—'], ['TIERRA', A.tierraTxt || '—'], ['AISLAMIENTO / COND.', A.ais + ' / ' + A.mat],
        ['CANALIZACIÓN', (A.preN || '') + A.AR139 + ' mm ' + A.tuberia], ['LONGITUD', f1(A.M139) + ' m'], ['VOLTAJE NOMINAL', r.tab.sistema + ' V'], ['VOLTAJE CALCULADO', f2(A.AW139) + ' V'], ['% CAÍDA DE VOLTAJE', f2(A.AY139) + ' %']];
      D.rect(x, y, w1 + w2, hh * (filas.length + 1), 'u-tab', 'TABLA');
      D.texto(x + (w1 + w2) / 2, y + 11, 'TABLA RESUMEN DEL TABLERO', 'u-t b', 'TABLA', 'c');
      filas.forEach((f, i) => {
        const yy = y + hh * (i + 1);
        D.linea(x, yy, x + w1 + w2, yy, 'u-tabl', 'TABLA');
        D.texto(x + 4, yy + 11, f[0], 'u-t s', 'TABLA'); D.texto(x + w1 + 4, yy + 11, f[1], 'u-t s b', 'TABLA');
      });
      D.linea(x + w1, y + hh, x + w1, y + hh * (filas.length + 1), 'u-tabl', 'TABLA');
    }

    function nodo(r, x0, y0) {
      const sw = span[r.tab.id], hs = hijos(r), yb = y0, cx = x0 + sw / 2;
      const ch = hs.reduce((a, x) => a + ancho[x.tab.id], 0) + GAP * Math.max(0, hs.length - 1) - (hs.length ? LBL : 0);
      const xs = [], xi = []; let x = cx - ch / 2;
      hs.forEach(hj => { xi.push(x); xs.push(x + span[hj.tab.id] / 2); x += ancho[hj.tab.id] + GAP; });
      const xa = Math.min(cx - W / 2, ...xs), xb = Math.max(cx + W / 2, ...xs), tx = xb + 14, mal = r.avisos.length;
      D.barra(xa, xb, yb, colorV(r.V), () => App.go('memoria', r.tab.id), r.nombre);
      D.texto(tx, yb - 22, r.nombre, 'u-n', 'TEXTO', null, () => App.go('memoria', r.tab.id));
      if (conTabla) tablaDatos(r, tx, yb - 8);
      else {
        D.texto(tx, yb - 6, r.tab.sistema + ' V · ' + (r.zapatas ? 'zapatas' : (r.alim.AF139 || '–') + ' A') + ' · ' + (r.cat ? r.cat.modelo : ''), 'u-t', 'TEXTO');
        D.texto(tx, yb + 9, f2(r.W130) + ' kVA · ' + f2(r.alim.L139) + ' kVA dem.', 'u-t', 'TEXTO');
        D.texto(tx, yb + 24, 'Bornes ' + f2(r.alim.AW139) + ' V · ΔV ' + f2(r.alim.AY139) + ' %', 'u-t' + (r.alim.AY139 > P.cvMaxAlim ? ' mal' : ''), 'TEXTO');
        D.texto(tx, yb + 39, 'Icc ' + (r.iccKA ? f2(r.iccKA) + ' kA' : '—') + (r.iccFuente ? ' (' + r.iccFuente + ')' : ''), 'u-t b' + (r.bkMain && r.bkMain.sccrBajo ? ' mal' : ''), 'TEXTO');
        if (mal) D.texto(tx, yb + 54, '⚠ ' + mal + ' aviso(s)', 'u-t warn', 'TEXTO', null, null, r.avisos.join('\n'));
        D.boton(tx + 31, yb + (mal ? 70 : 55), '+ tablero', () => nuevoDerivado(r.tab));
      }
      hs.forEach((hj, i) => {
        const yh = yb + NIVEL;
        bajada(xs[i], yb, yh, hj, hj.circuitoPadre);
        nodo(hj, xi[i], yh);
      });
    }

    let x = GAP;
    raiz.forEach(r => {
      const cx = x + span[r.tab.id] / 2, y0 = 30;
      D.rect(cx - 42, y0, 84, 34, 'u-red', 'EQUIPOS');
      D.texto(cx, y0 + 22, r.trafo ? 'RED' : (r.tab.conectadoA || 'ACOMETIDA').slice(0, 14), 'u-t b', 'TEXTO', 'c');
      D.texto(cx + 50, y0 + 14, r.trafo ? (P.iccRed ? 'Icc ' + f2(P.iccRed) + ' kA' : 'Icc infinita') : '', 'u-t', 'TEXTO');
      if (!r.trafo && r.iccFuente === 'manual') D.texto(cx + 50, y0 + 29, 'Icc en bornes ' + f2(r.iccKA) + ' kA (dato)', 'u-t', 'TEXTO');
      bajada(cx, y0 + 34, y0 + 34 + NIVEL - 60, r, null);
      nodo(r, x, y0 + 34 + NIVEL - 60);
      x += ancho[r.tab.id] + GAP;
    });
  }

  /* ---------- salida SVG ---------- */
  function svgDibujo(R) {
    const S = (tag, attrs, ...kids) => {
      const el = document.createElementNS(NS, tag);
      Object.entries(attrs || {}).forEach(([k, v]) => { if (v === null || v === undefined || v === false) return; if (k === 'onclick') el.addEventListener('click', v); else el.setAttribute(k, v); });
      kids.flat().forEach(k => { if (k === null || k === undefined) return; el.appendChild(typeof k === 'string' || typeof k === 'number' ? document.createTextNode(String(k)) : k); });
      return el;
    };
    let svg = null, g0 = null;
    const tit = t => (t ? S('title', null, t) : null);
    const D = {
      tamano(w, hh) {
        svg = S('svg', { class: 'unifilar', viewBox: '0 0 ' + w + ' ' + hh, width: w, height: hh, xmlns: NS, role: 'img', 'aria-label': 'Diagrama unifilar' });
        svg.appendChild(S('defs', null, S('pattern', { id: 'rayas', width: 6, height: 6, patternUnits: 'userSpaceOnUse', patternTransform: 'rotate(45)' }, S('line', { x1: 0, y1: 0, x2: 0, y2: 6, class: 'u-ray' }))));
        g0 = S('g'); svg.appendChild(g0);
      },
      linea: (x1, y1, x2, y2, cls) => g0.appendChild(S('line', { x1, y1, x2, y2, class: cls })),
      rect: (x, y, w, hh, cls, capa, ir) => g0.appendChild(S('rect', { x, y, width: w, height: hh, class: cls + (ir ? ' clic' : ''), onclick: ir })),
      circulo: (cx, cy, r, cls, capa, ir) => g0.appendChild(S('circle', { cx, cy, r, class: cls + (ir ? ' clic' : ''), onclick: ir })),
      texto: (x, y, t, cls, capa, al, ir, title) => g0.appendChild(S('text', { x, y, class: cls + (ir ? ' clic' : ''), 'text-anchor': al === 'c' ? 'middle' : al === 'e' ? 'end' : null, onclick: ir }, t, tit(title))),
      barra: (x1, x2, y, color, ir, nombre) => g0.appendChild(S('line', { x1, y1: y, x2, y2: y, class: 'u-bus', stroke: color, onclick: ir }, tit('Abrir memoria de cálculo de ' + nombre))),
      boton: (cx, cy, t, ir) => g0.appendChild(S('g', { class: 'u-add no-print', onclick: ir }, S('rect', { x: cx - 32, y: cy - 10, width: 64, height: 18, rx: 9 }), S('text', { x: cx, y: cy + 3, 'text-anchor': 'middle' }, t))),
    };
    dibujar(R, D);
    return svg;
  }

  /* ---------- salida DXF (AutoCAD R12, abre en AutoCAD / AutoCAD LT y se guarda como DWG) ---------- */
  const CAPAS = { TABLERO: 5, ALIMENTADOR: 7, PROTECCION: 1, TRAFO: 3, EQUIPOS: 6, TEXTO: 2, TABLA: 8 };
  function dxf(R) {
    let H = 0; const ent = [];
    const yy = y => +(H - y).toFixed(2), xx = x => +x.toFixed(2);
    const esc = t => String(t).replace(/[^\x20-\x7e]/g, ch => (ch === 'Δ' ? 'D' : ch === '⚠' ? '!' : ch === '↑' ? '^' : '\\U+' + ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')));
    const linea = (x1, y1, x2, y2, capa) => ent.push(['0', 'LINE', '8', capa, '10', xx(x1), '20', yy(y1), '30', 0, '11', xx(x2), '21', yy(y2), '31', 0]);
    const ALTO = { 'u-n': 11, 'u-t s': 7, 'u-t s b': 7 };
    const D = {
      tamano: (w, hh) => { H = hh; },
      linea: (x1, y1, x2, y2, cls, capa) => linea(x1, y1, x2, y2, capa || 'ALIMENTADOR'),
      rect: (x, y, w, hh, cls, capa) => { const c = capa || 'EQUIPOS'; linea(x, y, x + w, y, c); linea(x + w, y, x + w, y + hh, c); linea(x + w, y + hh, x, y + hh, c); linea(x, y + hh, x, y, c); if (cls === 'u-red') for (let i = 8; i < w + hh; i += 8) linea(x + Math.max(0, i - hh), y + Math.min(hh, i), x + Math.min(w, i), y + Math.max(0, i - w), c); },
      circulo: (cx, cy, r, cls, capa) => ent.push(['0', 'CIRCLE', '8', capa || 'EQUIPOS', '10', xx(cx), '20', yy(cy), '30', 0, '40', r]),
      texto: (x, y, t, cls, capa, al) => {
        const hgt = ALTO[cls] || 8.5, e = ['0', 'TEXT', '8', capa || 'TEXTO', '7', 'SINERGIA', '10', xx(x), '20', yy(y), '30', 0, '40', hgt, '1', esc(t)];
        if (al === 'c' || al === 'e') e.push('72', al === 'c' ? 1 : 2, '11', xx(x), '21', yy(y), '31', 0);
        ent.push(e);
      },
      barra: (x1, x2, y) => { [-2, 0, 2].forEach(d => linea(x1, y + d, x2, y + d, 'TABLERO')); },
      boton: () => {},
    };
    dibujar(R, D);
    const out = ['0', 'SECTION', '2', 'HEADER', '9', '$ACADVER', '1', 'AC1009', '9', '$DWGCODEPAGE', '3', 'ANSI_1252', '9', '$INSUNITS', '70', 4, '0', 'ENDSEC',
      '0', 'SECTION', '2', 'TABLES',
      '0', 'TABLE', '2', 'LTYPE', '70', 1, '0', 'LTYPE', '2', 'CONTINUOUS', '70', 0, '3', 'Solid line', '72', 65, '73', 0, '40', 0, '0', 'ENDTAB',
      '0', 'TABLE', '2', 'LAYER', '70', Object.keys(CAPAS).length].concat(...Object.entries(CAPAS).map(([k, c]) => ['0', 'LAYER', '2', 'E-' + k, '70', 0, '62', c, '6', 'CONTINUOUS']), ['0', 'ENDTAB',
      '0', 'TABLE', '2', 'STYLE', '70', 1, '0', 'STYLE', '2', 'SINERGIA', '70', 0, '40', 0, '41', 1, '50', 0, '71', 0, '42', 8, '3', 'GOTHIC.TTF', '4', '', '0', 'ENDTAB',
      '0', 'ENDSEC', '0', 'SECTION', '2', 'ENTITIES']);
    ent.forEach(e => { const k = e.slice(); k[3] = 'E-' + k[3]; out.push(...k); });
    out.push('0', 'ENDSEC', '0', 'EOF');
    const txt = []; for (let i = 0; i < out.length; i += 2) txt.push(String(out[i]).padStart(3, ' '), String(out[i + 1]));
    return txt.join('\r\n') + '\r\n';
  }

  function nuevoDerivado(padre) {
    UI.prompt('Nuevo tablero derivado de ' + Calc.nombreTablero(padre), 'Nombre del tablero', '', v => (v ? '' : 'Escriba un nombre')).then(nombre => {
      if (!nombre) return;
      const t = Store.nuevoTablero({ nombre, tipo: padre.tipo, sistema: padre.sistema });
      Store.cambiarPadre(t, padre.id); App.refresh(); UI.toast('Tablero ' + nombre + ' conectado a ' + Calc.nombreTablero(padre), 'ok');
    });
  }

  App.views.unifilar = function (view, R) {
    const P = Store.project;
    view.appendChild(h('div', { class: 'page-h row no-print' }, h('div', null, h('h2', null, 'Diagrama unifilar'), h('p', { class: 'muted' }, 'Clic en una barra para abrir su memoria, en un transformador para editarlo; "+ trafo" agrega un transformador aguas arriba del tablero.')),
      h('div', { class: 'toolbar' },
        UI.field('Icc de la red en el primario (kA)', UI.input(P, 'iccRed', { type: 'num', fk: 'u:icc', placeholder: 'infinita', class: 'w-num' }), null, 'inline'),
        h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: conTabla, onchange: e => { conTabla = e.target.checked; App.refresh(); } }), ' Tabla de datos junto a cada tablero'),
        UI.btn('+ Tablero principal', async () => { const n = await UI.prompt('Nuevo tablero principal', 'Nombre', '', v => (v ? '' : 'Escriba un nombre')); if (n) { Store.nuevoTablero({ nombre: n }); App.refresh(); } }, 'small'),
        UI.btn('Descargar DXF (AutoCAD)', () => { U.download(App.fileBase() + ' - unifilar.dxf', dxf(R), 'application/dxf'); UI.toast('DXF generado: ábralo en AutoCAD y use Guardar como… DWG', 'ok'); }, 'primary small', 'AutoCAD abre el DXF directamente; con "Guardar como" queda en DWG'),
        UI.btn('Descargar SVG', () => { const s = document.querySelector('svg.unifilar'); if (s) U.download(App.fileBase() + ' - unifilar.svg', '<?xml version="1.0" encoding="UTF-8"?>' + s.outerHTML.replace('<svg', '<svg style="font-family:Arial;background:#fff"'), 'image/svg+xml'); }, 'small'),
        UI.btn('Imprimir', () => window.print(), 'small'))));
    if (!R.orden.length) { view.appendChild(h('div', { class: 'empty' }, 'No hay tableros. Créelos en Proyecto o impórtelos de Revit.')); return; }
    view.appendChild(h('div', { class: 'u-wrap', 'data-scroll': 'unifilar' }, svgDibujo(R)));
    view.appendChild(h('div', { class: 'u-ley' }, h('span', null, h('i', { style: { background: '#c0392b' } }), '480 V'), h('span', null, h('i', { style: { background: '#1b7f4b' } }), '240 V'), h('span', null, h('i', { style: { background: '#0b5cab' } }), '208 V y menos'),
      h('span', null, '□ interruptor · ◎ transformador · G generador · ATS/MTS transferencia · Icc punto a punto (Bussmann) · ΔV acumulada desde la acometida · DXF con capas E-TABLERO, E-ALIMENTADOR, E-PROTECCION, E-TRAFO, E-EQUIPOS, E-TEXTO, E-TABLA')));
  };
  window.Unifilar = { editarTrafo, dxf };
})();
