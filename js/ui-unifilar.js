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
  const fasesPolos = r => (r.fases === 3 ? 3 : 2);
  /** Textos movibles: cada texto lleva una clave estable (contexto: tablero + número de orden) y su desplazamiento se guarda en el proyecto (unifilarPos). */
  function llaves() {
    let ctx = ''; const cont = {};
    return { ctx(k) { const p = ctx; if (k !== undefined) ctx = k; return p; }, key() { cont[ctx] = (cont[ctx] || 0) + 1; return ctx + '#' + cont[ctx]; } };
  }
  const posiciones = () => (Store.project && Store.project.unifilarPos) || {};
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
    d.protPrim = tr.protPrim !== false; d.protSec = tr.protSec === true; d.protPrimA = tr.protPrimA || ''; d.protSecA = tr.protSecA || '';
    const chk = (k, txt) => h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: !!d[k], onchange: e => { d[k] = e.target.checked; } }), ' ' + txt);
    UI.modal('Transformador que alimenta a ' + Calc.nombreTablero(t), h('div', null,
      h('p', { class: 'muted' }, 'Se ubica entre ' + (r && r.padre ? r.padre.nombre : 'la red') + ' y el alimentador de ' + Calc.nombreTablero(t) + '. El secundario es el sistema del tablero (' + t.sistema + ' V).'),
      h('div', { class: 'grid2' }, UI.field('De la lista del Excel', pres, null, 'span2'), UI.field('Nombre', nom), UI.field('Fases', fs), UI.field('Capacidad (kVA)', kva), UI.field('Impedancia Z (%)', z),
        UI.field('Relación X/R', xr, 'Para la caída de tensión en el transformador; 4 si no se conoce'), UI.field('Voltaje primario (V)', vp, padreV ? 'Voltaje del tablero que lo alimenta: ' + padreV + ' V' : 'Voltaje de la red'),
        UI.field('Protección en el primario', chk('protPrim', 'Dibujar interruptor en el primario'), 'Desmarque si no aplica: no aparece en el unifilar', 'span2'),
        UI.field('Protección primaria (A)', UI.input(d, 'protPrimA', { type: 'num', save: false, after: () => {}, placeholder: 'Auto (125 % de la corriente nominal)' }), 'Vacío = automática, NEC 450.3(B)'),
        UI.field('Protección en el secundario', chk('protSec', 'Dibujar interruptor en el secundario'), 'Marque si el transformador lleva protección en el secundario', 'span2'),
        UI.field('Protección secundaria (A)', UI.input(d, 'protSecA', { type: 'num', save: false, after: () => {}, placeholder: 'Auto (125 % de la corriente nominal)' }), 'Vacío = automática, NEC 450.3(B)'))), [
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
    const span = {}, ancho = {}, todos = r => (R.hijos[r.tab.id] || []).map(t => R.res[t.id]).filter(Boolean);
    const UXW = 520;
    /** La UPS cuelga de la barra del tablero que la alimenta (aguas arriba) y el tablero con la UPS queda bajo su MTS; un principal con UPS y sin tablero aguas arriba la dibuja en su propia barra. */
    const propio = r => !!(r.ups && (r.tab.ups.propio === undefined ? !r.padre : r.tab.ups.propio));   // la UPS cuelga de su propia barra
    const hijoUps = r => todos(r).find(x => x.ups && !propio(x)) || null;
    const ramaUps = r => !!(hijoUps(r) || propio(r));
    const hijoU = r => hijoUps(r) || (propio(r) ? todos(r).find(x => x.tab.deUps) : null) || null;   // tablero conectado aguas abajo del MTS de la UPS
    const hijos = r => todos(r).filter(x => x !== hijoU(r));
    const medir = r => { const hs = hijos(r), hu = hijoU(r); span[r.tab.id] = Math.max(W, hs.reduce((a, x) => a + medir(x), 0) + GAP * Math.max(0, hs.length - 1)); const eu = hu ? Math.max(UXW, medir(hu) + 170) : UXW; ancho[r.tab.id] = span[r.tab.id] + LBL + (ramaUps(r) ? eu : 0); return ancho[r.tab.id]; };
    const totalW = raiz.reduce((a, r) => a + medir(r) + GAP, GAP) + 40;
    const prof = r => 1 + Math.max(0, ...todos(r).map(prof));
    D.tamano(totalW, 120 + NIVEL * Math.max(1, ...raiz.map(prof)) + 120 + (R.orden.some(hijoU) ? NIVEL : 0));

    const HB = SIMB.INTERRUPTOR_3P.alto * SIMB.INTERRUPTOR_3P.k;     // alto del símbolo de interruptor
    /** Interruptor: bloque de 1, 2 o 3 polos centrado en (x, y) con atributos (polos, amperios, unidad, modelo, SCCR) y texto a la derecha. Devuelve [y superior, y inferior]. */
    const interruptor = (x, y, txt, mal, ir, d) => {
      d = d || {};
      const np = Math.max(1, Math.min(3, Number(d.polos) || 3));
      D.simbolo('INTERRUPTOR_' + np + 'P', x, y - HB / 2, { POLOS: np + 'p', AMPERIOS: d.amp ? d.amp + 'A' : '', UNIDAD: d.unidad || 'TM', MODELO: d.modelo || '', SCCR: d.sccr || '' }, { mal, ir, title: txt });
      D.texto(x + 19, y + 4, txt, 'u-t', 'TEXTO');
      return [y - HB / 2, y + HB / 2];
    };
    /** Línea vertical con huecos donde hay símbolos (los símbolos traen sus propios terminales). */
    const lineaV = (x, y0, y1, huecos, capa) => {
      let a = y0;
      huecos.slice().sort((m, n) => m[0] - n[0]).forEach(([h0, h1]) => { if (h0 > a) D.linea(x, a, x, h0, 'u-lin', capa); a = Math.max(a, h1); });
      if (y1 > a) D.linea(x, a, x, y1, 'u-lin', capa);
    };
    /** Bajada desde (x, y0) hasta la barra en y1. */
    function bajada_(x, y0, y1, r, circ) {
      const A = r.alim, tr = r.trafo; let y = y0; const huecos = [];
      const bkA = circ ? circ.AF : A.AF139, b = circ ? circ.breaker : r.bkMain;
      y += circ ? (conTabla ? 262 : 96) + (r.padre && r.padre.alterna ? 40 : 0) : 26;
      if (!tr || tr.protPrim) huecos.push(interruptor(x, y, (bkA || '–') + ' A ' + (b ? (b.polos || '') + 'P ' + (b.unidad === 'STD' ? 'TM' : b.unidad || '') + ' · ' + (b.ref || b.modelo || '') : ''), b && b.sccrBajo, () => App.go('memoria', circ ? r.padre.tab.id : r.tab.id),
        { polos: b ? b.polos : fasesPolos(r), amp: bkA, unidad: b ? (b.unidad === 'STD' || !b.unidad ? 'TM' : b.unidad) : 'TM', modelo: b ? (b.ref || b.modelo || '') : '', sccr: b ? b.sccr : '' }));
      if (tr) {
        y += 58;
        const kt = SIMB.TRAFO_SECO.k, ht = SIMB.TRAFO_SECO.alto * kt;
        D.simbolo('TRAFO_SECO', x, y - ht / 2, { KVA: f1(tr.kva) + ' kVA', IMPEDANCIA: 'Z ' + f2(tr.z) + ' %', VOLTAJES: (tr.Vp ? f0(tr.Vp) : '?') + ' / ' + r.tab.sistema + ' V' }, { ir: () => (tr.compartido ? App.go('proyecto') : editarTrafo(r.tab, R)), title: tr.compartido ? 'Transformador compartido: se edita en Proyecto → Fuentes' : 'Editar transformador' });
        huecos.push([y - ht / 2 - 3 * kt, y + ht / 2 + 3 * kt]);
        D.texto(x + 22, y - 12, (tr.nombre ? tr.nombre + ' · ' : '') + f1(tr.kva) + ' kVA · Z ' + f2(tr.z) + ' %' + (tr.compartido ? ' · compartido (' + tr.salidas + ' salidas)' : ''), 'u-t b', 'TEXTO');
        D.texto(x + 22, y + 3, (tr.Vp ? f0(tr.Vp) : '?') + ' / ' + r.tab.sistema + ' V · carga ' + f1(tr.carga * 100) + ' %', 'u-t', 'TEXTO');
        D.texto(x + 22, y + 18, 'Icc sec. ' + f2(tr.iccSec / 1000) + ' kA · ΔV ' + f2(tr.reg) + ' %', 'u-t', 'TEXTO');
        if (tr.protSec && tr.sec) {   // protección en el secundario del transformador (NEC 450.3(B)); solo se dibuja si el transformador la tiene marcada
          const b = tr.sec.bk, ys = y + ht / 2 + 3 * kt + 34;
          huecos.push(interruptor(x, ys, tr.sec.A + ' A ' + (b ? (b.polos || '') + 'P ' + (b.unidad === 'STD' ? 'TM' : b.unidad || '') + ' · ' + (b.ref || b.modelo || '') : (tr.fases === 3 ? '3' : '2') + 'P · secundario'), b && b.sccrBajo, () => (tr.compartido ? App.go('proyecto') : editarTrafo(r.tab, R)),
            { polos: b ? b.polos : (tr.fases === 3 ? 3 : 2), amp: tr.sec.A, unidad: b ? (b.unidad === 'STD' || !b.unidad ? 'TM' : b.unidad) : 'TM', modelo: b ? (b.ref || b.modelo || '') : '', sccr: b ? b.sccr : '' }));
          y = ys + 6;
        }
      } else {
        y += 50;
        D.boton(x, y, '+ trafo', () => editarTrafo(r.tab, R));
      }
      y = Math.max(y + 46, y1 - (r.alterna ? 176 : 74));
      // con segunda acometida los datos del cable van a la izquierda de la línea
      const cx = r.alterna ? x - 10 : x + 10, al = r.alterna ? 'e' : null;
      D.texto(cx, y, A.fasesTxt + (A.neutroTxt ? '+' + A.neutroTxt + 'N' : '') + (A.tierraTxt ? '+' + A.tierraTxt + 'T' : '') + ' ' + A.mat + ' ' + A.ais, 'u-t b', 'TEXTO', al);
      const dN = r.alterna && /ATS|MTS/.test(r.alterna.equipo) ? r.alterna.distNormal : A.M139;
      D.texto(cx, y + 15, (r.alterna ? 'Normal → ' + r.alterna.equipo + ' ' : '') + f1(dN) + ' m · ' + (A.preN || '') + A.AR139 + ' mm ' + A.tuberia + ' · ΔV ' + f2(A.AV139 / r.V * 100) + ' %' + (A.ajustado ? ' · ajustado' : ''), 'u-t', 'TEXTO', al);
      if (A.val && !A.val.ok) D.texto(cx, y + 30, '⚠ no cumple ampacidad corregida', 'u-t mal', 'TEXTO', al);
      const fin = () => lineaV(x, y0, y1, huecos, 'ALIMENTADOR');
      // segunda acometida: generador o bypass → ATS / MTS / IP sobre la barra
      if (r.alterna) {
        // equipo de transferencia pegado a la barra del tablero; la rama alterna con su propia protección
        const al = r.alterna, sym = al.equipo === 'MTS' ? 'MTS' : al.equipo === 'ATS' ? 'ATS' : null;
        let ya;
        const gx = x + 140;
        if (sym) {
          // transferencia sobre la barra: normal por arriba, alterna por la derecha, carga hacia la barra
          const S0 = SIMB[sym], ha = S0.alto * S0.k, oy = y1 - 58 - ha, pa = S0.puertos.alterna;
          ya = oy - pa[1] * S0.k;
          D.simbolo(sym, x, oy, { MODELO: al.ats ? al.ats.modelo : '', AMPERIOS: al.ats ? String(al.ats.amperios) : String(A.AF139 || ''), WCR: al.wcr ? String(al.wcr) : '', TRANSICION: al.transicion || '' }, { ir: () => App.go('memoria', r.tab.id), title: al.equipo });
          huecos.push([oy, oy + ha]);
          D.texto(x - 6 - (S0.k * 10), oy + 12, al.equipo, 'u-t b', 'TEXTO', 'e');
          if (al.ats || al.distAts) D.texto(x - S0.k * 10 - 6, oy + 27, (al.ats ? al.ats.modelo + ' · WCR ' + al.wcr + ' kA' : '') + (al.distAts ? ' · ' + f1(al.distAts) + ' m al tablero' : ''), 'u-t', 'TEXTO', 'e');
          D.texto(x + 6, oy - 6, 'NORMAL' + (r.trafo ? ' (transformador)' : r.padre ? ' (' + r.padre.nombre + ')' : r.tab.conectadoA ? ' (' + r.tab.conectadoA + ')' : ''), 'u-t s b', 'TEXTO');
          // tramo común ATS → tablero: espacio para el calibre, canalización, distancia y caída
          D.texto(x - 10, y1 - 38, A.fasesTxt + (A.neutroTxt ? '+' + A.neutroTxt + 'N' : '') + (A.tierraTxt ? '+' + A.tierraTxt + 'T' : '') + ' ' + A.mat + ' ' + A.ais, 'u-t b', 'TEXTO', 'e');
          D.texto(x - 10, y1 - 23, al.equipo + ' → tablero ' + f1(al.distAts) + ' m · ' + (A.preN || '') + A.AR139 + ' mm ' + A.tuberia + ' · ΔV ' + f2(al.caidaComun / r.V * 100) + ' %', 'u-t', 'TEXTO', 'e');
          D.texto(x + pa[0] * S0.k + 6, ya + 12, 'EMERGENCIA', 'u-t s b', 'TEXTO');
          D.linea(x + pa[0] * S0.k, ya, gx, ya, 'u-lin', 'ALIMENTADOR');
        } else {
          ya = y1 - 15;
          D.rect(x - 22, ya - 12, 44, 24, 'u-ats', 'EQUIPOS', () => App.go('memoria', r.tab.id)); D.texto(x, ya + 4, al.equipo, 'u-t b', 'TEXTO', 'c'); huecos.push([ya - 12, ya + 12]);
          D.linea(x + 22, ya, gx, ya, 'u-lin', 'ALIMENTADOR');
        }
        const yb2 = ya - 58;
        // interruptor de la rama alterna: el circuito del bypass en el tablero de origen o el principal del generador
        const origen = al.tipo !== 'generador' ? Object.values(R.res).find(o => o.rows.some(z => z.c.respaldoDe === r.tab.id)) : null;
        const cb = origen ? origen.rows.find(z => z.c.respaldoDe === r.tab.id) : null;
        const bkTxt = cb ? (cb.AF || '–') + ' A ' + (cb.breaker ? cb.polosBreaker + 'P ' + (cb.breaker.ref || cb.breaker.modelo || '') : '') : (A.AF139 || '–') + ' A ' + (fasesPolos(r)) + 'P';
        const hg = interruptor(gx, yb2, bkTxt, cb && cb.breaker && cb.breaker.sccrBajo, () => App.go('memoria', origen ? origen.tab.id : r.tab.id),
          { polos: cb && cb.breaker ? cb.polosBreaker : fasesPolos(r), amp: cb ? cb.AF : A.AF139, unidad: 'TM', modelo: cb && cb.breaker ? (cb.breaker.ref || cb.breaker.modelo || '') : '', sccr: cb && cb.breaker ? cb.breaker.sccr : '' });
        const alBus = al.tipo !== 'generador' && r.padre && al.origenId === r.padre.tab.id;   // el bypass sube hasta la barra del tablero que alimenta
        D.linea(gx, ya, gx, hg[1], 'u-lin', 'ALIMENTADOR'); D.linea(gx, hg[0], gx, alBus ? y0 : yb2 - 52, 'u-lin', 'ALIMENTADOR');
        if (alBus) D.punto && D.punto(gx, y0);
        if (al.tipo === 'generador') {
          D.simbolo('GENERADOR', gx, yb2 - 52, { NOMBRE: al.origen || 'Generador', KVA: al.kva ? f1(al.kva) + ' kVA' : '', XD: al.xd ? f1(al.xd) + ' %' : '' }, { ir: () => App.go('memoria', r.tab.id), title: 'Generador' });
          D.texto(gx + 22, yb2 - 74, (al.origen || 'Generador') + (al.xd ? " · X''d " + f1(al.xd) + ' %' : ''), 'u-t b', 'TEXTO');
        } else {
          D.texto(gx + 10, yb2 - 60, 'BYPASS desde ' + (al.origen || '?'), 'u-t b', 'TEXTO');
          if (cb) D.texto(gx + 10, yb2 - 45, 'circuito [' + cb.polos.join(',') + '] en ' + origen.nombre, 'u-t', 'TEXTO');
        }
        D.texto(gx + 8, ya - 26, (al.tipo === 'generador' ? 'Gen → ' : 'Bypass → ') + al.equipo + ' ' + al.fasesTxt + ' ' + al.mat + ' · ' + f1(al.longitud) + ' m · ΔV ' + f2(al.cv) + ' %', 'u-t', 'TEXTO');
        D.texto(gx + 8, ya - 11, 'Icc ' + (al.iccA ? f2(al.iccA / 1000) + ' kA' : '—') + (al.carga !== null ? ' · carga ' + f1(al.carga * 100) + ' %' : '') + (al.iccAtsKA ? ' · en ' + al.equipo + ' ' + f2(al.iccAtsKA) + ' kA' : ''), 'u-t', 'TEXTO');
      }
      fin();
    }

    /** Tabla de datos del tablero (como las tablas de los DU). */
    function tablaDatos(r, x, y) {
      const A = r.alim, w1 = 170, w2 = 140, hh = 15;
      const filas = [['EQUIPO', r.nombre], ['CORRIENTE CORTOCIRCUITO', r.iccKA ? f2(r.iccKA) + ' kA' : '—'], ['kVA TOTALES', f2(r.W130)], ['FACTOR DE DEMANDA', f2(A.R139)], ['kVA DEMANDADOS', f2(A.L139)],
        ['FACTOR DE POTENCIA', f2(A.T139)], ['ACOMETIDA', ''], ['FASES', A.fasesTxt], ['NEUTRO', A.neutroTxt || '—'], ['TIERRA', A.tierraTxt || '—'], ['AISLAMIENTO / COND.', A.ais + ' / ' + A.mat],
        ['CANALIZACIÓN', (A.preN || '') + A.AR139 + ' mm ' + A.tuberia], ['LONGITUD', f1(A.M139) + ' m'], ['VOLTAJE NOMINAL', r.tab.sistema + ' V'], ['VOLTAJE CALCULADO', f2(A.AW139) + ' V'], ['% CAÍDA DE VOLTAJE', f2(A.AY139) + ' %']];
      if (r.alterna) {
        const a = r.alterna;
        filas.push([a.equipo + ' / ' + (a.tipo === 'generador' ? 'GENERADOR' : 'BYPASS'), ''], ['MODELO ' + a.equipo, a.ats ? a.ats.modelo + ' · ' + a.ats.amperios + ' A' : (A.AF139 || '–') + ' A'], ['TRANSICIÓN / POLOS', (a.equipo === 'MTS' ? 'manual' : a.transicion || '—') + ' · ' + (a.polosAts || '—') + 'P'],
          ['WCR / Icc EN EL ' + a.equipo, (a.wcr || '—') + ' / ' + (a.iccAtsKA ? f2(a.iccAtsKA) : '—') + ' kA'], ['NORMAL → ' + a.equipo, f1(a.distNormal) + ' m'], [(a.tipo === 'generador' ? 'GENERADOR' : 'BYPASS') + ' → ' + a.equipo, f1(a.longitud) + ' m'],
          [a.equipo + ' → TABLERO', f1(a.distAts) + ' m'], ['GENERADOR', a.tipo === 'generador' ? (a.kva ? f1(a.kva) + ' kVA' : '—') : (a.origen || '—')], ['CALIBRE 2ª ACOMETIDA', a.fasesTxt + ' ' + a.mat], ['CAÍDA 2ª ACOMETIDA', f2(a.cv) + ' %']);
      }
      if (r.upsMts) filas.push(['MTS (BYPASS UPS)', ''], ['MODELO MTS', r.upsMts.ats ? r.upsMts.ats.modelo + ' · ' + r.upsMts.ats.amperios + ' A' : (A.AF139 || '–') + ' A'], ['WCR / Icc EN EL MTS', (r.upsMts.wcr || '—') + ' / ' + (r.iccKA ? f2(r.iccKA) : '—') + ' kA']);
      if (r.ups) filas.push(['UPS', r.ups.nombre || ''], ['CAPACIDAD UPS', f1(r.ups.kva) + ' kVA'], ['Icc DEL INVERSOR', f2(r.ups.icc / 1000) + ' kA']);
      D.rect(x, y, w1 + w2, hh * (filas.length + 1), 'u-tab', 'TABLA');
      D.texto(x + (w1 + w2) / 2, y + 11, 'TABLA RESUMEN DEL TABLERO', 'u-t b', 'TABLA', 'c');
      filas.forEach((f, i) => {
        const yy = y + hh * (i + 1);
        D.linea(x, yy, x + w1 + w2, yy, 'u-tabl', 'TABLA');
        D.texto(x + 4, yy + 11, f[0], 'u-t s', 'TABLA'); D.texto(x + w1 + 4, yy + 11, f[1], 'u-t s b', 'TABLA');
      });
      D.linea(x + w1, y + hh, x + w1, y + hh * (filas.length + 1), 'u-tabl', 'TABLA');
    }

    /** UPS del tablero principal: dos interruptores en la barra (UPS y bypass) → UPS → MTS → interruptor de salida, aguas abajo del principal. */
    function ramaUpsDibujo_(r, xu, yb, hu) {
      const up = hu && hu.ups ? hu : r, u = up.ups, al = up.tab.alterna, mts = up === r ? !!r.tab.ups.mts : !!(al && al.activo && al.equipo === 'MTS' && al.origenId === r.tab.id), x1 = mts ? xu - 45 : xu, ir = () => App.go('memoria', up.tab.id);
      const fases = fasesPolos(r), hb1 = interruptor(x1, yb + 30, 'UPS ' + f1(u.kva) + ' kVA', false, ir, { polos: fases, unidad: 'TM' });
      const yU = yb + 120;
      D.linea(x1, yb, x1, hb1[0], 'u-lin', 'ALIMENTADOR'); D.linea(x1, hb1[1], x1, yU, 'u-lin', 'ALIMENTADOR');
      D.rect(x1 - 22, yU, 44, 30, 'u-ups', 'EQUIPOS', ir); D.texto(x1, yU + 19, 'UPS', 'u-t b', 'TEXTO', 'c');
      D.texto(x1 - 28, yU + 8, (u.nombre || 'UPS') + ' ' + f1(u.kva) + ' kVA', 'u-t', 'TEXTO', 'e');
      D.texto(x1 - 28, yU + 22, 'Icc ' + f2(u.icc / 1000) + ' kA', 'u-t', 'TEXTO', 'e');
      let yS = yU + 30;
      if (mts) {
        const S0 = SIMB.MTS, ha = S0.alto * S0.k, pa = S0.puertos.alterna, oy = yU + 30 + 40, xb2 = xu + 45;
        const hb2 = interruptor(xb2, yb + 30, 'Bypass', false, ir, { polos: fases, unidad: 'TM' });
        D.simbolo('MTS', x1, oy, { MODELO: '', AMPERIOS: '', WCR: '', TRANSICION: 'manual' }, { ir, title: 'MTS' });
        D.linea(x1, yU + 30, x1, oy, 'u-lin', 'ALIMENTADOR');
        const ya = oy - pa[1] * S0.k, xa2 = x1 + pa[0] * S0.k;
        D.linea(xb2, yb, xb2, hb2[0], 'u-lin', 'ALIMENTADOR'); D.linea(xb2, hb2[1], xb2, ya, 'u-lin', 'ALIMENTADOR'); D.linea(xa2, ya, xb2, ya, 'u-lin', 'ALIMENTADOR');
        D.texto(x1 - 30, oy + 12, 'MTS (bypass UPS)', 'u-t b', 'TEXTO', 'e');
        yS = oy + ha;
      }
      if (hu) return [x1, yS];   // el tablero conectado aguas abajo se dibuja con su formato normal
      const yO = yS + 28, ho = interruptor(x1, yO, 'Salida UPS', false, ir, { polos: fases, unidad: 'TM' });
      D.linea(x1, yS, x1, ho[0], 'u-lin', 'ALIMENTADOR');
      D.linea(x1, ho[1], x1, ho[1] + 22, 'u-lin', 'ALIMENTADOR'); D.rect(x1 - 36, ho[1] + 22, 72, 44, 'u-tab', 'EQUIPOS');
      D.texto(x1, ho[1] + 48, 'Cargas críticas', 'u-t s', 'TEXTO', 'c');
      return [x1, yS];
    }

    function nodo_(r, x0, y0) {
      const sw = span[r.tab.id], hs = hijos(r), yb = y0, cx = x0 + sw / 2;
      const ch = hs.reduce((a, x) => a + ancho[x.tab.id], 0) + GAP * Math.max(0, hs.length - 1) - (hs.length ? LBL : 0);
      const xs = [], xi = []; let x = cx - ch / 2;
      hs.forEach(hj => { xi.push(x); xs.push(x + span[hj.tab.id] / 2); x += ancho[hj.tab.id] + GAP; });
      const xu = Math.max(cx + W / 2, ...xs) + (xs.length ? 380 : 90), xa = Math.min(cx - W / 2, ...xs), xb = ramaUps(r) ? xu + 150 : Math.max(cx + W / 2, ...xs), tx = xb + 14, mal = r.avisos.length;
      const o = r.alterna ? 40 : 0;   // con segunda acometida los datos van debajo de la barra (arriba está la rama alterna)
      D.barra(xa, xb, yb, colorV(r.V), () => App.go('memoria', r.tab.id), r.nombre);
      D.texto(tx, yb - 22 + o, r.nombre, 'u-n', 'TEXTO', null, () => App.go('memoria', r.tab.id));
      if (!r.padre) {   // barra de tierra del tablero de acometida (conductor del electrodo, 250.66)
        D.linea(xa + 8, yb, xa + 8, yb + 14, 'u-lin', 'TABLERO');
        D.simbolo('TIERRA', xa + 8, yb + 14, {}, { title: 'Barra de tierra' });
        D.texto(xa + 22, yb + 33, 'Tierra' + (r.alim.AP140 ? ' ' + r.alim.AP140 + ' AWG' : ''), 'u-t s', 'TEXTO');
      }
      if (conTabla) tablaDatos(r, tx, yb - 8 + o);
      else {
        D.texto(tx, yb - 6 + o, r.tab.sistema + ' V · ' + (r.zapatas ? 'zapatas' : (r.alim.AF139 || '–') + ' A') + ' · ' + (r.cat ? r.cat.modelo : ''), 'u-t', 'TEXTO');
        D.texto(tx, yb + 9 + o, f2(r.W130) + ' kVA · ' + f2(r.alim.L139) + ' kVA dem.', 'u-t', 'TEXTO');
        D.texto(tx, yb + 24 + o, 'Bornes ' + f2(r.alim.AW139) + ' V · ΔV ' + f2(r.alim.AY139) + ' %', 'u-t' + (r.alim.AY139 > P.cvMaxAlim ? ' mal' : ''), 'TEXTO');
        D.texto(tx, yb + 39 + o, 'Icc ' + (r.iccKA ? f2(r.iccKA) + ' kA' : '—') + (r.iccFuente ? ' (' + r.iccFuente + ')' : ''), 'u-t b' + (r.bkMain && r.bkMain.sccrBajo ? ' mal' : ''), 'TEXTO');
        if (mal) D.texto(tx, yb + 54 + o, '⚠ ' + mal + ' aviso(s)', 'u-t warn', 'TEXTO', null, null, r.avisos.join('\n'));
        D.boton(tx + 31, yb + o + (mal ? 70 : 55), '+ tablero', () => nuevoDerivado(r.tab));
      }
      if (ramaUps(r)) {
        const hu = hijoU(r), [x1, yS] = ramaUpsDibujo(r, xu, yb, hu);
        if (hu) { const yh = yS + NIVEL; bajada(x1, yS, yh, hu, hu.circuitoPadre); nodo(hu, x1 - span[hu.tab.id] / 2, yh); }
      }
      hs.forEach((hj, i) => {
        const yh = yb + NIVEL;
        bajada(xs[i], yb, yh, hj, hj.circuitoPadre);
        nodo(hj, xi[i], yh);
      });
    }

    let x = GAP;
    const conCtx = (k, f) => { const prev = D.ctx ? D.ctx(k) : ''; try { return f(); } finally { if (D.ctx) D.ctx(prev); } };
    const bajada = (x, y0, y1, r, circ) => conCtx(r.tab.id + (circ ? ':c' : ':b'), () => bajada_(x, y0, y1, r, circ));
    const ramaUpsDibujo = (r, xu, yb, hu) => conCtx(r.tab.id + ':u', () => ramaUpsDibujo_(r, xu, yb, hu));
    const nodo = (r, x0, y0) => conCtx(r.tab.id + ':n', () => nodo_(r, x0, y0));
    raiz.forEach(r => conCtx('r:' + r.tab.id, () => {
      const cx = x + span[r.tab.id] / 2, y0 = 30;
      D.rect(cx - 42, y0, 84, 34, 'u-red', 'EQUIPOS');
      D.texto(cx, y0 + 22, r.trafo ? 'RED' : (Store.fuente(r.tab.fuenteId) ? ((Store.fuente(r.tab.fuenteId).nombre || 'GENERADOR')) : (r.tab.conectadoA || 'ACOMETIDA')).slice(0, 14), 'u-t b', 'TEXTO', 'c');
      D.texto(cx + 50, y0 + 14, r.trafo ? (P.iccRed ? 'Icc ' + f2(P.iccRed) + ' kA' : 'Icc infinita') : '', 'u-t', 'TEXTO');
      if (!r.trafo && r.iccFuente === 'manual') D.texto(cx + 50, y0 + 29, 'Icc en bornes ' + f2(r.iccKA) + ' kA (dato)', 'u-t', 'TEXTO');
      bajada(cx, y0 + 34, y0 + 34 + NIVEL - 60, r, null);
      nodo(r, x, y0 + 34 + NIVEL - 60);
      x += ancho[r.tab.id] + GAP;
    }));
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
    const L = llaves(), pos = posiciones();
    const D = {
      ctx: L.ctx,
      tamano(w, hh) {
        svg = S('svg', { class: 'unifilar', viewBox: '0 0 ' + w + ' ' + hh, width: w, height: hh, xmlns: NS, role: 'img', 'aria-label': 'Diagrama unifilar' });
        svg.appendChild(S('defs', null, S('pattern', { id: 'rayas', width: 6, height: 6, patternUnits: 'userSpaceOnUse', patternTransform: 'rotate(45)' }, S('line', { x1: 0, y1: 0, x2: 0, y2: 6, class: 'u-ray' }))));
        g0 = S('g'); svg.appendChild(g0);
      },
      linea: (x1, y1, x2, y2, cls) => g0.appendChild(S('line', { x1, y1, x2, y2, class: cls })),
      rect: (x, y, w, hh, cls, capa, ir) => g0.appendChild(S('rect', { x, y, width: w, height: hh, class: cls + (ir ? ' clic' : ''), onclick: ir })),
      circulo: (cx, cy, r, cls, capa, ir) => g0.appendChild(S('circle', { cx, cy, r, class: cls + (ir ? ' clic' : ''), onclick: ir })),
      texto: (x, y, t, cls, capa, al, ir, title) => {
        const k = capa === 'TEXTO' ? L.key() : null, o = k && pos[k];
        return g0.appendChild(S('text', { x, y, class: cls + (ir ? ' clic' : '') + (k ? ' u-mov' : ''), 'text-anchor': al === 'c' ? 'middle' : al === 'e' ? 'end' : null, 'data-k': k, transform: o ? 'translate(' + o.dx + ' ' + o.dy + ')' : null, onclick: ir }, t, tit(title || (k ? 'Arrastre para mover este texto' : null))));
      },
      barra: (x1, x2, y, color, ir, nombre) => g0.appendChild(S('line', { x1, y1: y, x2, y2: y, class: 'u-bus', stroke: color, onclick: ir }, tit('Abrir memoria de cálculo de ' + nombre))),
      simbolo: (nombre, x, y, vals, o) => {
        o = o || {}; const Y = SIMB[nombre], k = Y.k, X = px => +(x + px * k).toFixed(2), V = py => +(y - py * k).toFixed(2);
        const el = S('g', { class: 'u-sim' + (o.mal ? ' mal' : '') + (o.ir ? ' clic' : ''), onclick: o.ir }, tit(o.title));
        Y.prims.forEach(m => {
          if (m[0] === 'L') el.appendChild(S('line', { x1: X(m[1]), y1: V(m[2]), x2: X(m[3]), y2: V(m[4]) }));
          else if (m[0] === 'C') el.appendChild(S('circle', { cx: X(m[1]), cy: V(m[2]), r: +(m[3] * k).toFixed(2) }));
          else if (m[0] === 'F') el.appendChild(S('circle', { cx: X(m[1]), cy: V(m[2]), r: +(m[3] * k).toFixed(2), class: 'rell' }));
          else if (m[0] === 'A') {
            const [cx, cy, r, a1, a2] = [m[1], m[2], m[3], m[4] * Math.PI / 180, m[5] * Math.PI / 180], g = a => [X(cx + r * Math.cos(a)), V(cy + r * Math.sin(a))];
            const [sx, sy] = g(a1), [ex, ey] = g(a2), grande = ((m[5] - m[4] + 720) % 360) > 180 ? 1 : 0;
            el.appendChild(S('path', { d: 'M' + sx + ' ' + sy + 'A' + +(r * k).toFixed(2) + ' ' + +(r * k).toFixed(2) + ' 0 ' + grande + ' 0 ' + ex + ' ' + ey }));
          } else if (m[0] === 'T') el.appendChild(S('text', { x: X(m[1]), y: V(m[2]), class: 'u-st', 'font-size': +(m[3] * k).toFixed(1), 'text-anchor': 'middle', 'dominant-baseline': 'central' }, m[4]));
        });
        g0.appendChild(el);
      },
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
    const L = llaves(), pos = posiciones();
    const D = {
      ctx: L.ctx,
      tamano: (w, hh) => { H = hh; },
      linea: (x1, y1, x2, y2, cls, capa) => linea(x1, y1, x2, y2, capa || 'ALIMENTADOR'),
      rect: (x, y, w, hh, cls, capa) => { const c = capa || 'EQUIPOS'; linea(x, y, x + w, y, c); linea(x + w, y, x + w, y + hh, c); linea(x + w, y + hh, x, y + hh, c); linea(x, y + hh, x, y, c); if (cls === 'u-red') for (let i = 8; i < w + hh; i += 8) linea(x + Math.max(0, i - hh), y + Math.min(hh, i), x + Math.min(w, i), y + Math.max(0, i - w), c); },
      circulo: (cx, cy, r, cls, capa) => ent.push(['0', 'CIRCLE', '8', capa || 'EQUIPOS', '10', xx(cx), '20', yy(cy), '30', 0, '40', r]),
      texto: (x, y, t, cls, capa, al) => {
        if (capa === 'TEXTO') { const o = pos[L.key()]; if (o) { x += o.dx; y += o.dy; } }
        const hgt = ALTO[cls] || 8.5, e = ['0', 'TEXT', '8', capa || 'TEXTO', '7', 'SINERGIA', '10', xx(x), '20', yy(y), '30', 0, '40', hgt, '1', esc(t)];
        if (al === 'c' || al === 'e') e.push('72', al === 'c' ? 1 : 2, '11', xx(x), '21', yy(y), '31', 0);
        ent.push(e);
      },
      barra: (x1, x2, y) => { [-2, 0, 2].forEach(d => linea(x1, y + d, x2, y + d, 'TABLERO')); },
      boton: () => {},
      simbolo: (nombre, x, y, vals) => {
        const Y = SIMB[nombre], conAtr = (Y.attrs || []).length > 0, e = ['0', 'INSERT', '8', Y.capa].concat(conAtr ? ['66', 1] : [], ['2', nombre, '10', xx(x), '20', yy(y), '30', 0, '41', Y.k, '42', Y.k, '43', Y.k, '50', 0]);
        ent.push(e);
        if (conAtr) {
          (Y.attrs || []).forEach(([tag, , ax, ay, ah, vis]) => ent.push(['0', 'ATTRIB', '8', Y.capa, '10', xx(x + ax * Y.k), '20', yy(y) + +(ay * Y.k).toFixed(2), '30', 0, '40', +(ah * Y.k).toFixed(2), '1', esc((vals || {})[tag] || ' '), '7', 'SINERGIA', '2', tag, '70', vis ? 0 : 1]));
          ent.push(['0', 'SEQEND', '8', Y.capa]);
        }
      },
    };
    dibujar(R, D);
    // definiciones de bloque (capa 0: toman la capa del INSERT)
    const bloques = [];
    Object.entries(SIMB).forEach(([nombre, Y]) => {
      const a = ['0', 'BLOCK', '8', '0', '2', nombre, '70', (Y.attrs || []).length ? 2 : 0, '10', 0, '20', 0, '30', 0, '3', nombre];
      Y.prims.forEach(m => {
        if (m[0] === 'L') a.push('0', 'LINE', '8', '0', '10', m[1], '20', m[2], '30', 0, '11', m[3], '21', m[4], '31', 0);
        else if (m[0] === 'C') a.push('0', 'CIRCLE', '8', '0', '10', m[1], '20', m[2], '30', 0, '40', m[3]);
        else if (m[0] === 'F') { const w = m[3], r2 = +(w / 2).toFixed(3); a.push('0', 'POLYLINE', '8', '0', '66', 1, '10', 0, '20', 0, '30', 0, '40', w, '41', w, '70', 1,
          '0', 'VERTEX', '8', '0', '10', +(m[1] - r2).toFixed(3), '20', m[2], '30', 0, '42', 1, '0', 'VERTEX', '8', '0', '10', +(m[1] + r2).toFixed(3), '20', m[2], '30', 0, '42', 1, '0', 'SEQEND', '8', '0'); }
        else if (m[0] === 'A') a.push('0', 'ARC', '8', '0', '10', m[1], '20', m[2], '30', 0, '40', m[3], '50', m[4], '51', m[5]);
        else if (m[0] === 'T') a.push('0', 'TEXT', '8', '0', '7', 'SINERGIA', '10', m[1], '20', m[2], '30', 0, '40', m[3], '1', esc(m[4]), '72', 1, '73', 2, '11', m[1], '21', m[2], '31', 0);
      });
      (Y.attrs || []).forEach(([tag, prompt, ax, ay, ah, vis]) => a.push('0', 'ATTDEF', '8', '0', '10', ax, '20', ay, '30', 0, '40', ah, '1', ' ', '3', esc(prompt), '2', tag, '70', vis ? 0 : 1, '7', 'SINERGIA'));
      a.push('0', 'ENDBLK', '8', '0');
      bloques.push(...a);
    });
    const out = ['0', 'SECTION', '2', 'HEADER', '9', '$ACADVER', '1', 'AC1009', '9', '$DWGCODEPAGE', '3', 'ANSI_1252', '9', '$INSUNITS', '70', 4, '0', 'ENDSEC',
      '0', 'SECTION', '2', 'TABLES',
      '0', 'TABLE', '2', 'LTYPE', '70', 1, '0', 'LTYPE', '2', 'CONTINUOUS', '70', 0, '3', 'Solid line', '72', 65, '73', 0, '40', 0, '0', 'ENDTAB',
      '0', 'TABLE', '2', 'LAYER', '70', Object.keys(CAPAS).length].concat(...Object.entries(CAPAS).map(([k, c]) => ['0', 'LAYER', '2', 'E-' + k, '70', 0, '62', c, '6', 'CONTINUOUS']), ['0', 'ENDTAB',
      '0', 'TABLE', '2', 'STYLE', '70', 1, '0', 'STYLE', '2', 'SINERGIA', '70', 0, '40', 0, '41', 1, '50', 0, '71', 0, '42', 8, '3', 'GOTHIC.TTF', '4', '', '0', 'ENDTAB',
      '0', 'ENDSEC', '0', 'SECTION', '2', 'BLOCKS'].concat(bloques, ['0', 'ENDSEC', '0', 'SECTION', '2', 'ENTITIES']));
    ent.forEach(e => { const k = e.slice(); if (k[2] === '8') k[3] = 'E-' + k[3]; out.push(...k); });
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

  let zoomU = 1, manoU = true, quitarTeclas = null, quitarArrastre = null;   // ampliación del diagrama y herramienta mano (arrastrar para moverse); se conservan al refrescar
  const ZMIN = 0.1, ZMAX = 4;
  App.views.unifilar = function (view, R) {
    const P = Store.project;
    let etqZoom = null, wrap = null, btnMano = null;
    /** Cambia la ampliación (el centro visible se mantiene) y ajusta el tamaño del SVG sin volver a dibujar. */
    const poner = (z, cx, cy) => {
      const s = wrap && wrap.querySelector('svg.unifilar'); if (!s) return;
      const w0 = Number(s.dataset.w), h0 = Number(s.dataset.h), antes = zoomU;
      zoomU = Math.min(ZMAX, Math.max(ZMIN, z));
      const fx = cx === undefined ? wrap.clientWidth / 2 : cx, fy = cy === undefined ? wrap.clientHeight / 2 : cy;
      const px = (wrap.scrollLeft + fx) / antes, py = (wrap.scrollTop + fy) / antes;
      s.setAttribute('width', Math.round(w0 * zoomU)); s.setAttribute('height', Math.round(h0 * zoomU));
      wrap.scrollLeft = px * zoomU - fx; wrap.scrollTop = py * zoomU - fy;
      if (etqZoom) etqZoom.textContent = Math.round(zoomU * 100) + ' %';
    };
    const ajustar = () => { const s = wrap && wrap.querySelector('svg.unifilar'); if (s) poner(Math.min(1, (wrap.clientWidth - 4) / Number(s.dataset.w))); };
    view.appendChild(h('div', { class: 'page-h row no-print' }, h('div', null, h('h2', null, 'Diagrama unifilar'), h('p', { class: 'muted' }, 'Clic en una barra para abrir su memoria, en un transformador para editarlo; "+ trafo" agrega un transformador aguas arriba del tablero.')),
      h('div', { class: 'toolbar' },
        UI.field('Icc de la red en el primario (kA)', UI.input(P, 'iccRed', { type: 'num', fk: 'u:icc', placeholder: 'infinita', class: 'w-num' }), null, 'inline'),
        h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: conTabla, onchange: e => { conTabla = e.target.checked; App.refresh(); } }), ' Tabla de datos junto a cada tablero'),
        UI.btn('+ Tablero principal', async () => { const n = await UI.prompt('Nuevo tablero principal', 'Nombre', '', v => (v ? '' : 'Escriba un nombre')); if (n) { Store.nuevoTablero({ nombre: n }); App.refresh(); } }, 'small'),
        UI.btn('Restablecer textos', async () => { if (!Object.keys(posiciones()).length) { UI.toast('No hay textos movidos', 'ok'); return; } if (await UI.confirm('¿Devolver todos los textos del unifilar a su posición automática?', 'Restablecer')) { Store.project.unifilarPos = {}; Store.save(); App.refresh(); } }, 'small', 'Arrastre cualquier texto del diagrama para acomodarlo; esto los devuelve a su lugar'),
        UI.btn('Descargar DXF (AutoCAD)', () => { U.download(App.fileBase() + ' - unifilar.dxf', dxf(R), 'application/dxf'); UI.toast('DXF generado: ábralo en AutoCAD y use Guardar como… DWG', 'ok'); }, 'primary small', 'AutoCAD abre el DXF directamente; con "Guardar como" queda en DWG'),
        UI.btn('Descargar SVG', () => { const s = document.querySelector('svg.unifilar'); if (s) U.download(App.fileBase() + ' - unifilar.svg', '<?xml version="1.0" encoding="UTF-8"?>' + s.outerHTML.replace(/ width="\d+"/, ' width="' + s.dataset.w + '"').replace(/ height="\d+"/, ' height="' + s.dataset.h + '"').replace('<svg', '<svg style="font-family:Arial;background:#fff"'), 'image/svg+xml'); }, 'small'),
        h('span', { class: 'zoom-u', title: 'Ctrl + rueda del ratón también acerca y aleja' },
          UI.btn('−', () => poner(zoomU / 1.25), 'small', 'Alejar'), etqZoom = h('span', { class: 'zoom-v' }, Math.round(zoomU * 100) + ' %'), UI.btn('+', () => poner(zoomU * 1.25), 'small', 'Acercar'),
          UI.btn('100 %', () => poner(1), 'small', 'Tamaño real'), UI.btn('Ajustar', ajustar, 'small', 'Cabe todo el ancho en pantalla'),
          btnMano = UI.btn('✋ Mover', () => { manoU = !manoU; if (wrap) wrap.classList.toggle('mano', manoU); btnMano.classList.toggle('primary', manoU); }, 'small' + (manoU ? ' primary' : ''), 'Arrastre el diagrama para moverse (también con el botón central del ratón o con la barra espaciadora pulsada)')),
        UI.btn('Imprimir', () => window.print(), 'small'))));
    if (!R.orden.length) { view.appendChild(h('div', { class: 'empty' }, 'No hay tableros. Créelos en Proyecto o impórtelos de Revit.')); return; }
    wrap = h('div', { class: 'u-wrap', 'data-scroll': 'unifilar' }, svgDibujo(R));
    view.appendChild(wrap);
    wrap.classList.toggle('mano', manoU);
    {   // textos movibles: arrastrar un texto lo desplaza; la posición se guarda en el proyecto
      let dr = null, movio = false;
      wrap.addEventListener('pointerdown', e => {
        const tx = e.target.closest && e.target.closest('text.u-mov'); if (!tx || e.button !== 0) return;
        const sv = wrap.querySelector('svg.unifilar'), m = sv && sv.getScreenCTM(); if (!m) return;
        const k = tx.getAttribute('data-k'), o = posiciones()[k] || { dx: 0, dy: 0 };
        dr = { tx, k, x: e.clientX, y: e.clientY, dx0: o.dx, dy0: o.dy, s: m.a || 1, moved: false };
        e.stopPropagation(); e.preventDefault();
      }, true);
      const mover = e => {
        if (!dr) return;
        const dx = (e.clientX - dr.x) / dr.s, dy = (e.clientY - dr.y) / dr.s;
        if (!dr.moved && Math.hypot(e.clientX - dr.x, e.clientY - dr.y) < 3) return;
        dr.moved = true; dr.nx = dr.dx0 + dx; dr.ny = dr.dy0 + dy;
        dr.tx.setAttribute('transform', 'translate(' + dr.nx + ' ' + dr.ny + ')'); wrap.classList.add('arrastrando');
      };
      const soltar = () => {
        if (!dr) return;
        wrap.classList.remove('arrastrando');
        if (dr.moved) {
          const P = Store.project; P.unifilarPos = P.unifilarPos || {};
          P.unifilarPos[dr.k] = { dx: Math.round(dr.nx * 10) / 10, dy: Math.round(dr.ny * 10) / 10 }; Store.save();
          movio = true; setTimeout(() => { movio = false; }, 0);
        }
        dr = null;
      };
      wrap.addEventListener('click', e => { if (movio) { e.stopPropagation(); e.preventDefault(); } }, true);
      if (quitarArrastre) quitarArrastre();
      window.addEventListener('pointermove', mover); window.addEventListener('pointerup', soltar);
      quitarArrastre = () => { window.removeEventListener('pointermove', mover); window.removeEventListener('pointerup', soltar); };
    }
    {   // mano: arrastrar para desplazar el diagrama; un arrastre no cuenta como clic sobre una barra o un tablero
      let ini = null, movio = false, espacio = false;
      const activo = e => manoU || e.button === 1 || espacio;
      wrap.addEventListener('pointerdown', e => { if (!activo(e) || (e.button !== 0 && e.button !== 1) || e.target.closest('.u-add')) return; ini = { x: e.clientX, y: e.clientY, sl: wrap.scrollLeft, st: wrap.scrollTop }; movio = false; if (e.button === 1) e.preventDefault(); });
      wrap.addEventListener('pointermove', e => { if (!ini) return; const dx = e.clientX - ini.x, dy = e.clientY - ini.y; if (!movio && Math.hypot(dx, dy) < 4) return; movio = true; try { wrap.setPointerCapture(e.pointerId); } catch (x) { /* sin captura */ } wrap.classList.add('arrastrando'); wrap.scrollLeft = ini.sl - dx; wrap.scrollTop = ini.st - dy; });
      const fin = e => { if (!ini) return; ini = null; wrap.classList.remove('arrastrando'); try { wrap.releasePointerCapture(e.pointerId); } catch (x) { /* ya liberado */ } };
      wrap.addEventListener('pointerup', fin); wrap.addEventListener('pointercancel', fin);
      wrap.addEventListener('click', e => { if (movio) { e.stopPropagation(); e.preventDefault(); movio = false; } }, true);
      const tecla = (v, e) => { if (e.code === 'Space' && !/INPUT|TEXTAREA|SELECT/.test((e.target.tagName || ''))) { espacio = v; wrap.classList.toggle('mano', manoU || espacio); if (v && document.querySelector('.u-wrap') === wrap) e.preventDefault(); } };
      if (quitarTeclas) quitarTeclas(); const kd = e => tecla(true, e), ku = e => tecla(false, e); document.addEventListener('keydown', kd); document.addEventListener('keyup', ku); quitarTeclas = () => { document.removeEventListener('keydown', kd); document.removeEventListener('keyup', ku); };
    }
    { const sv = wrap.querySelector('svg.unifilar'); if (sv) { sv.dataset.w = sv.getAttribute('width'); sv.dataset.h = sv.getAttribute('height'); if (zoomU !== 1) { sv.setAttribute('width', Math.round(sv.dataset.w * zoomU)); sv.setAttribute('height', Math.round(sv.dataset.h * zoomU)); } } }
    wrap.addEventListener('wheel', e => { if (!e.ctrlKey) return; e.preventDefault(); const b = wrap.getBoundingClientRect(); poner(zoomU * (e.deltaY < 0 ? 1.1 : 1 / 1.1), e.clientX - b.left, e.clientY - b.top); }, { passive: false });
    view.appendChild(h('div', { class: 'u-ley' }, h('span', null, h('i', { style: { background: '#c0392b' } }), '480 V'), h('span', null, h('i', { style: { background: '#1b7f4b' } }), '240 V'), h('span', null, h('i', { style: { background: '#0b5cab' } }), '208 V y menos'),
      h('span', null, '□ interruptor · ◎ transformador · G generador · ATS/MTS transferencia · Icc punto a punto (Bussmann) · ΔV acumulada desde la acometida · DXF con capas E-TABLERO, E-ALIMENTADOR, E-PROTECCION, E-TRAFO, E-EQUIPOS, E-TEXTO, E-TABLA')));
  };
  window.Unifilar = { editarTrafo, dxf };
})();
