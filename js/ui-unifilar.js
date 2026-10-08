/* Diagrama unifilar (estilo ETAP / SKM): red o acometida → transformador → interruptor → alimentador → barra de cada tablero.
   Muestra en cada barra el voltaje en bornes, la caída acumulada y la corriente de cortocircuito; permite agregar transformadores
   y tableros derivados directamente en el diagrama. */
(function () {
  'use strict';
  const h = U.h, f2 = v => U.fmt(v, 2), f1 = v => U.fmt(v, 1);
  const NS = 'http://www.w3.org/2000/svg';
  const S = (tag, attrs, ...kids) => {
    const el = document.createElementNS(NS, tag);
    Object.entries(attrs || {}).forEach(([k, v]) => { if (v === null || v === undefined || v === false) return; if (k.startsWith('on')) el.addEventListener(k.slice(2), v); else el.setAttribute(k, v); });
    kids.flat().forEach(k => { if (k === null || k === undefined || k === false) return; el.appendChild(typeof k === 'string' || typeof k === 'number' ? document.createTextNode(String(k)) : k); });
    return el;
  };
  const W = 150, LBL = 250, GAP = 30, NIVEL = 300;
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

  /* ---------- dibujo ---------- */
  function diagrama(R) {
    const raiz = R.orden.filter(r => r.nivel === 0);
    const ancho = {}, hijos = r => (R.hijos[r.tab.id] || []).map(t => R.res[t.id]).filter(Boolean);
    // ancho de cada subárbol: barra (o la suma de sus derivados) + espacio para los datos a la derecha de la barra
    const span = {}, medir = r => { const hs = hijos(r); span[r.tab.id] = Math.max(W, hs.reduce((a, x) => a + medir(x), 0) + GAP * Math.max(0, hs.length - 1)); ancho[r.tab.id] = span[r.tab.id] + LBL; return ancho[r.tab.id]; };
    const totalW = raiz.reduce((a, r) => a + medir(r) + GAP, GAP);
    const prof = r => 1 + Math.max(0, ...hijos(r).map(prof));
    const totalH = 120 + NIVEL * Math.max(1, ...raiz.map(prof)) + 40;
    const svg = S('svg', { class: 'unifilar', viewBox: '0 0 ' + totalW + ' ' + totalH, width: totalW, height: totalH, xmlns: NS, role: 'img', 'aria-label': 'Diagrama unifilar' });
    const P = Store.project, g0 = S('g');
    svg.appendChild(S('defs', null, S('pattern', { id: 'rayas', width: 6, height: 6, patternUnits: 'userSpaceOnUse', patternTransform: 'rotate(45)' }, S('line', { x1: 0, y1: 0, x2: 0, y2: 6, class: 'u-ray' }))));
    svg.appendChild(g0);

    /** Bajada desde (x, y0) hasta la barra en y1: interruptor, transformador opcional y cable. */
    function bajada(x, y0, y1, r, circ) {
      const A = r.alim, tr = r.trafo, grp = S('g', { class: 'u-baj' });
      let y = y0;
      grp.appendChild(S('line', { x1: x, y1: y, x2: x, y2: y1, class: 'u-lin' }));
      // interruptor del tablero padre (o principal de la acometida)
      const bkA = circ ? circ.AF : A.AF139, bkTxt = circ ? (circ.breaker ? (circ.breaker.ref || circ.breaker.modelo || '') : '') : (r.bkMain ? (r.bkMain.ref || r.bkMain.modelo || '') : '');
      y += circ ? 96 : 26; // debajo de los datos de la barra del tablero padre
      grp.appendChild(S('rect', { x: x - 9, y: y - 9, width: 18, height: 18, class: 'u-bk' + ((circ ? circ.breaker && circ.breaker.sccrBajo : r.bkMain && r.bkMain.sccrBajo) ? ' mal' : ''), onclick: () => App.go('memoria', circ ? r.padre.tab.id : r.tab.id) }, S('title', null, 'Interruptor ' + (bkA || '') + ' A ' + bkTxt)));
      grp.appendChild(S('text', { x: x + 16, y: y + 4, class: 'u-t' }, (bkA || '–') + ' A' + (bkTxt ? ' · ' + bkTxt : '')));
      // transformador
      if (tr) {
        y += 58;
        const t2 = S('g', { class: 'u-tr', onclick: () => editarTrafo(r.tab, R) },
          S('circle', { cx: x, cy: y - 11, r: 15 }), S('circle', { cx: x, cy: y + 11, r: 15 }), S('title', null, 'Editar transformador'));
        grp.appendChild(t2);
        grp.appendChild(S('text', { x: x + 22, y: y - 12, class: 'u-t b' }, (tr.nombre ? tr.nombre + ' · ' : '') + f1(tr.kva) + ' kVA · Z ' + f2(tr.z) + ' %'));
        grp.appendChild(S('text', { x: x + 22, y: y + 3, class: 'u-t' }, (tr.Vp ? f1(tr.Vp) : '?') + ' / ' + r.tab.sistema + ' V · carga ' + f1(tr.carga * 100) + ' %'));
        grp.appendChild(S('text', { x: x + 22, y: y + 18, class: 'u-t' }, 'Icc sec. ' + f2(tr.iccSec / 1000) + ' kA · ΔV ' + f2(tr.reg) + ' %'));
      } else {
        y += 50;
        grp.appendChild(S('g', { class: 'u-add no-print', onclick: () => editarTrafo(r.tab, R) }, S('rect', { x: x - 34, y: y - 10, width: 68, height: 18, rx: 9 }), S('text', { x, y: y + 3, 'text-anchor': 'middle' }, '+ trafo'), S('title', null, 'Agregar transformador')));
      }
      // cable del alimentador
      y = Math.max(y + 46, y1 - 74);
      grp.appendChild(S('text', { x: x + 10, y, class: 'u-t b' }, A.fasesTxt + (A.neutroTxt ? '+' + A.neutroTxt + 'N' : '') + (A.tierraTxt ? '+' + A.tierraTxt + 'T' : '') + ' ' + A.mat + ' ' + A.ais));
      grp.appendChild(S('text', { x: x + 10, y: y + 15, class: 'u-t' }, f1(A.M139) + ' m · ' + (A.preN || '') + A.AR139 + ' mm ' + A.tuberia + ' · ΔV ' + f2(A.AV139 / r.V * 100) + ' %' + (A.ajustado ? ' · ajustado' : '')));
      if (A.val && !A.val.ok) grp.appendChild(S('text', { x: x + 10, y: y + 30, class: 'u-t mal' }, '⚠ no cumple ampacidad corregida'));
      return grp;
    }

    /** Barra del tablero y sus derivados; los datos de la barra van a la derecha de su extremo (como en ETAP). */
    function nodo(r, x0, y0) {
      const sw = span[r.tab.id], hs = hijos(r), yb = y0, cx = x0 + sw / 2;
      const ch = hs.reduce((a, x) => a + ancho[x.tab.id], 0) + GAP * Math.max(0, hs.length - 1) - (hs.length ? LBL : 0);
      const xs = [], xi = []; let x = cx - ch / 2;
      hs.forEach(hj => { xi.push(x); xs.push(x + span[hj.tab.id] / 2); x += ancho[hj.tab.id] + GAP; });
      const xa = Math.min(cx - W / 2, ...xs), xb = Math.max(cx + W / 2, ...xs), tx = xb + 14;
      const col = colorV(r.V), mal = r.avisos.length;
      g0.appendChild(S('line', { x1: xa, y1: yb, x2: xb, y2: yb, class: 'u-bus', stroke: col, onclick: () => App.go('memoria', r.tab.id) }, S('title', null, 'Abrir memoria de cálculo de ' + r.nombre)));
      g0.appendChild(S('text', { x: tx, y: yb - 22, class: 'u-n', onclick: () => App.go('memoria', r.tab.id) }, r.nombre));
      g0.appendChild(S('text', { x: tx, y: yb - 6, class: 'u-t' }, r.tab.sistema + ' V · ' + (r.alim.AF139 || '–') + ' A · ' + (r.cat ? r.cat.modelo : '')));
      g0.appendChild(S('text', { x: tx, y: yb + 9, class: 'u-t' }, f2(r.W130) + ' kVA · ' + f2(r.alim.L139) + ' kVA dem.'));
      g0.appendChild(S('text', { x: tx, y: yb + 24, class: 'u-t' + (r.alim.AY139 > P.cvMaxAlim ? ' mal' : '') }, 'Bornes ' + f2(r.alim.AW139) + ' V · ΔV ' + f2(r.alim.AY139) + ' %'));
      g0.appendChild(S('text', { x: tx, y: yb + 39, class: 'u-t b' + (r.bkMain && r.bkMain.sccrBajo ? ' mal' : '') }, 'Icc ' + (r.iccKA ? f2(r.iccKA) + ' kA' : '—') + (r.iccFuente ? ' (' + r.iccFuente + ')' : '')));
      if (mal) g0.appendChild(S('text', { x: tx, y: yb + 54, class: 'u-t warn' }, '⚠ ' + mal + ' aviso(s)', S('title', null, r.avisos.join('\n'))));
      g0.appendChild(S('g', { class: 'u-add no-print', onclick: () => nuevoDerivado(r.tab) }, S('rect', { x: tx, y: yb + (mal ? 62 : 47), width: 62, height: 18, rx: 9 }), S('text', { x: tx + 31, y: yb + (mal ? 75 : 60), 'text-anchor': 'middle' }, '+ tablero'), S('title', null, 'Agregar tablero derivado')));
      hs.forEach((hj, i) => {
        const yh = yb + NIVEL;
        g0.appendChild(bajada(xs[i], yb, yh, hj, hj.circuitoPadre));
        nodo(hj, xi[i], yh);
      });
    }

    let x = GAP;
    raiz.forEach(r => {
      const w = ancho[r.tab.id], cx = x + span[r.tab.id] / 2, y0 = 30;
      // red / acometida
      g0.appendChild(S('rect', { x: cx - 42, y: y0, width: 84, height: 34, class: 'u-red' }));
      g0.appendChild(S('text', { x: cx, y: y0 + 22, 'text-anchor': 'middle', class: 'u-t b' }, r.trafo ? 'RED' : (r.tab.conectadoA || 'ACOMETIDA').slice(0, 14)));
      g0.appendChild(S('text', { x: cx + 50, y: y0 + 14, class: 'u-t' }, r.trafo ? (P.iccRed ? 'Icc ' + f2(P.iccRed) + ' kA' : 'Icc infinita') : (r.tab.conectadoA ? '' : 'Acometida')));
      if (!r.trafo && r.iccFuente === 'manual') g0.appendChild(S('text', { x: cx + 50, y: y0 + 29, class: 'u-t' }, 'Icc en bornes ' + f2(r.iccKA) + ' kA (dato)'));
      g0.appendChild(bajada(cx, y0 + 34, y0 + 34 + NIVEL - 60, r, null));
      nodo(r, x, y0 + 34 + NIVEL - 60);
      x += w + GAP;
    });
    return svg;
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
        UI.btn('+ Tablero principal', async () => { const n = await UI.prompt('Nuevo tablero principal', 'Nombre', '', v => (v ? '' : 'Escriba un nombre')); if (n) { Store.nuevoTablero({ nombre: n }); App.refresh(); } }, 'small'),
        UI.btn('Descargar SVG', () => { const s = document.querySelector('svg.unifilar'); if (s) U.download(App.fileBase() + ' - unifilar.svg', '<?xml version="1.0" encoding="UTF-8"?>' + s.outerHTML.replace('<svg', '<svg style="font-family:Arial;background:#fff"'), 'image/svg+xml'); }, 'small'),
        UI.btn('Imprimir', () => window.print(), 'small'))));
    if (!R.orden.length) { view.appendChild(h('div', { class: 'empty' }, 'No hay tableros. Créelos en Proyecto o impórtelos de Revit.')); return; }
    view.appendChild(h('div', { class: 'u-wrap', 'data-scroll': 'unifilar' }, diagrama(R)));
    view.appendChild(h('div', { class: 'u-ley' }, h('span', null, h('i', { style: { background: '#c0392b' } }), '480 V'), h('span', null, h('i', { style: { background: '#1b7f4b' } }), '240 V'), h('span', null, h('i', { style: { background: '#0b5cab' } }), '208 V y menos'),
      h('span', null, '□ interruptor · ◎ transformador · Icc por el método punto a punto (Bussmann) · ΔV acumulada desde la acometida')));
  };
  window.Unifilar = { editarTrafo };
})();
