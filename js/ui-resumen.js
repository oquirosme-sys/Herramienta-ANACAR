/* Tablas resumen para los planos de Revit (hojas TABLA RESUMEN, (VERTICAL), DU(VERTICAL) y ORIGINAL) y orden/exclusión de tableros
   (hoja ORDEN TABLEROS). Resumen.* construye los datos; la vista y la exportación a Excel los usan igual. */
(function () {
  'use strict';
  const h = U.h;
  const t2 = v => (v === null || v === undefined ? '' : String(v).trim());
  const r2 = v => (v === null || v === undefined || v === '' || isNaN(v) ? '' : Math.round(Number(v) * 100) / 100);

  const Resumen = {
    /** Tableros en el orden de la tabla resumen: número de orden y luego el orden de alimentación; sin los excluidos. */
    ordenados(R) {
      const l = R.orden.map((r, i) => ({ r, i })).filter(x => !x.r.tab.excluirResumen)
        .sort((a, b) => (num(a.r.tab.orden) - num(b.r.tab.orden)) || (a.i - b.i)).map(x => x.r);
      Resumen.conTransf = l.some(r => r.alterna || r.ups || r.upsMts);   // si algún tablero lleva ATS/MTS o UPS la tabla horizontal incluye sus datos
      Resumen.conFuentes = l.some(r => r.alterna || r.ups || r.upsMts);   // las tablas verticales agregan el bloque de fuentes solo si algún tablero lo usa
      return l;
    },
    conFuentes: false,
    conTransf: false,
    /** Datos de la segunda acometida (ATS / MTS, generador o bypass) y de la UPS de un tablero. */
    transf(r) {
      const a = r.alterna, u = r.ups, A = r.alim, m = !a && r.upsMts;
      if (m) return ['MTS', m.ats ? m.ats.modelo : '', m.ats ? m.ats.amperios : (A.AF139 || ''), 'manual', 3, m.wcr || '', r.iccKA ? r2(r.iccKA) : '', 'Bypass desde la barra', '', '', '', '', '', u ? r2(u.kva) : '', u ? r2(u.icc / 1000) : ''];
      return [a ? a.equipo : '', a && a.ats ? a.ats.modelo : '', a ? (a.ats ? a.ats.amperios : (A.AF139 || '')) : '', a ? (a.equipo === 'MTS' ? 'manual' : a.transicion) : '', a ? a.polosAts : '', a ? (a.wcr || '') : '', a && a.iccAtsKA ? r2(a.iccAtsKA) : '',
        a ? (a.tipo === 'generador' ? 'Generador' + (a.kva ? ' ' + a.kva + ' kVA' : '') : 'Bypass ' + (a.origen || '')) : '', a ? r2(a.distNormal) : '', a ? r2(a.longitud) : '', a ? r2(a.distAts) : '',
        a ? a.fasesTxt + ' ' + a.mat : '', a ? r2(a.cv) : '', u ? r2(u.kva) : '', u ? r2(u.icc / 1000) : ''];
    },
    subTransf: ['Equipo', 'Modelo de referencia', 'Amperios (A)', 'Transición', '# polos', 'WCR (kA)', 'Icc en el ATS (kA)', 'Fuente alterna', 'Normal → ATS (m)', 'Alterna → ATS (m)', 'ATS → tablero (m)', 'Conductores 2ª acometida', 'Caída 2ª acometida (%)', 'UPS (kVA)', 'Icc del inversor (kA)'],
    /** Bloque de UPS y segunda acometida (ATS / MTS) de las tablas verticales; mismas filas para todos los tableros. */
    fuentes(r) {
      if (!Resumen.conFuentes) return [];
      const a = r.alterna, tr = a && /ATS|MTS/.test(a.equipo), u = r.ups, d = a ? a.distAts || 0 : 0, M = r.alim.M139;
      // '–' y no vacío: la tabla toma el estilo de cada fila de la primera columna (una fila vacía se vería como encabezado)
      return [['UPS / Segunda acometida', ''],
        ['UPS', u ? u.nombre + ' ' + r2(u.kva) + ' kVA' : '–'],
        ['Segunda acometida', a ? (a.tipo === 'generador' ? 'Generador' : 'Bypass') + (a.origen ? ' · ' + a.origen : '') : '–'],
        ['Equipo de transferencia', a ? a.equipo + (a.ats ? ' ' + a.ats.modelo + ' · ' + a.ats.amperios + ' A' : '') : '–'],
        ['WCR / Icc en el equipo (kA)', tr ? (a.wcr || '—') + ' / ' + (a.iccAtsKA ? r2(a.iccAtsKA) : '—') : '–'],
        ['Conductores fuente alterna', a ? a.fasesTxt + ' ' + a.mat : '–'],
        ['Generador / bypass → ATS-MTS (m)', a ? r2(a.longitud) : '–'],
        ['Transformador / medidores → ATS-MTS (m)', tr ? r2(Math.max(0, M - d)) : '–'],
        ['ATS-MTS → tablero (m)', tr ? r2(d) : '–']];
    },
    grupos: [['Tablero / Equipo', 1], ['Alimentado desde', 1], ['kVA Totales', 1], ['kVA Demandados', 1], ['Factor demanda', 1], ['Factor diversidad', 1], ['Factor potencia', 1],
      ['Calibres conductores alimentador', 8], ['Tubería', 3], ['Longitud (m)', 1], ['Voltaje bornes (V)', 1], ['Caída de voltaje total (V)', 1], ['Caída de voltaje total (%)', 1],
      ['Corriente cortocircuito disponible (kA)', 1], ['Capacidad de breaker (A)', 1], ['DATOS DEL TABLERO', 8], ['DATOS DEL SUPRESOR', 7], ['DATOS INTERRUPTOR PRINCIPAL', 7], ['DATOS TRANSFERENCIA (ATS / MTS) Y UPS', 15]],
    sub: ['', '', '', '', '', '', '', 'Fases (AWG)', '', 'Neutro (AWG)', '', 'Tierra (AWG)', '', 'Material', 'Aislamiento', 'Cantidad', 'Ø (mm)', 'Tipo', '', '', '', '', '', '',
      'Tablero / Equipo', 'Modelo de referencia', 'Fabricante', 'Barras Fase', 'Barras Neutro', 'Barras Tierra', 'Espacios', 'Montaje',
      'Fabricante', 'Modelo de referencia', 'Montaje', 'Supresión L-L (kA)', 'Supresión L-N / L-T / N-T (kA)', 'Voltaje (V)', 'Fases',
      'Fabricante', 'Modelo de referencia', 'Marco (A)', 'Amperios (A)', 'Tipo de unidad', '# polos', 'SCCR (kA)'].concat(['Equipo', 'Modelo de referencia', 'Amperios (A)', 'Transición', '# polos', 'WCR (kA)', 'Icc en el ATS (kA)', 'Fuente alterna', 'Normal → ATS (m)', 'Alterna → ATS (m)', 'ATS → tablero (m)', 'Conductores 2ª acometida', 'Caída 2ª acometida (%)', 'UPS (kVA)', 'Icc del inversor (kA)']),
    /** Fila de la TABLA RESUMEN horizontal (columnas I:BB de la fila 144 del Machote). */
    horizontal(r) {
      const A = r.alim, cat = r.cat || {}, spd = r.spd || {}, bk = r.bkMain || {};
      return [r.nombre, r.alimentadoDesde, r2(r.W130), r2(A.L139), r2(A.R139), r2(A.S139), r2(A.T139),
        A.preF, t2(A.AL139), t2(A.preN), t2(A.AN139), t2(A.preN), t2(A.AP139), A.mat, A.ais, t2(A.preN), A.AR139, A.tuberia,
        r2(A.M139), r2(A.AW139), r2(A.AX139), r2(A.AY139), r.iccKA ? r2(r.iccKA) : '', A.AF139 || '',
        r.nombre, cat.modelo || '', r.marca || '', cat.barraFase || '', cat.barraNeutro || '', cat.barraTierra || '', cat.espacios || '', r.tab.montaje || '',
        spd.modelo ? r.marca : '', spd.modelo || '', spd.montaje || '', spd.kaLL || '', spd.kaLN || '', spd.voltaje || '', spd.fases || '',
        bk.modelo !== undefined ? r.marca : '', bk.modelo || '', Number(bk.marco) ? bk.marco : '', bk.amperios || '', bk.unidad || '', bk.polos || '', bk.sccr || ''].concat(Resumen.transf(r));
    },
    /** TABLA RESUMEN (VERTICAL): [etiqueta, valor, valor2] (BU18:BW39). */
    vertical(r) {
      const A = r.alim;
      return [['Tablero / Equipo', r.nombre], ['Alimentado desde', r.alimentadoDesde], ['kVA Conectados', r2(r.W130)], ['kVA Demandados', r2(A.L139)], ['Factor demanda', r2(A.R139)],
        ['Factor potencia', r2(A.T139)], ['Factor de diversidad', r2(A.S139)], ['Calibres conductores alimentador', ''], ['Fases (AWG)', A.preF, t2(A.AL139)],
        ['Neutro (AWG)', t2(A.preN), t2(A.AN139)], ['Tierra (AWG)', t2(A.preN), t2(A.AP139)], ['Material', A.mat], ['Aislamiento', A.ais], ['Tubería', ''], ['Ø (mm)', t2(A.preN), A.AR139], ['Tipo', A.tuberia],
        ['Longitud (m)', r2(A.M139)], ['Voltaje bornes (V)', r2(A.AW139)], ['Caída de voltaje total (V)', r2(A.AX139)], ['Caída de voltaje total (%)', r2(A.AY139)],
        ['Corriente cortocircuito disponible (kA)', r.iccKA ? r2(r.iccKA) : ''], ['Capacidad de breaker (A)', A.AF139 || '']].concat(Resumen.fuentes(r));
    },
    /** TABLA RESUMEN DU (VERTICAL) (BR18:BT30). */
    du(r) {
      const A = r.alim;
      return [['Tablero / Equipo', r.nombre], ['Conectado a', r.alimentadoDesde], ['kVA Conectados', r2(r.W130)], ['kVA Demandados', r2(A.L139)], ['Factor de demanda', r2(A.R139)],
        ['Capacidad de breaker (A)', A.AF139 || ''], ['Fases (AWG)', A.preF, t2(A.AL139)], ['Neutro (AWG)', t2(A.preN), t2(A.AN139)], ['Tierra (AWG)', t2(A.preN), t2(A.AP139)],
        ['Material', A.mat], ['Aislamiento', A.ais], ['Ø (mm)', t2(A.preN), A.AR139], ['Tipo', A.tuberia]].concat(Resumen.fuentes(r));
    },
    /** TABLA RESUMEN ORIGINAL (BZ18:CB56), incluye los datos del transformador si se indicaron. */
    original(r) {
      const A = r.alim, tr = r.tab.trafo || {};
      return [['Tablero', r.nombre], ['Capacidad del transformador (kVA)', tr.kva || ''], ['Impedancia (%Z)', tr.z || ''], ['Fases', tr.fases || ''], ['Voltaje devanado primario (kV)', tr.primario ? Number(tr.primario) / 1000 : ''],
        ['Voltaje devanado secundario (V)', tr.kva ? r.tab.sistema : ''], ['Corriente de cortocircuito calculada (kA)', r.trafo && r.trafo.iccSec ? r2(r.trafo.iccSec / 1000) : (r.iccKA ? r2(r.iccKA) : '')], ['Potencia total (kVA)', r2(r.W130)], ['Potencia neta (kVA)', r2(A.L139)],
        ['Factor de demanda', r2(A.R139)], ['Factor de potencia', r2(A.T139)], ['Factor de diversidad', r2(A.S139)], ['Alimentadores', ''],
        ['Fases', A.preF, t2(A.AL139)], ['Neutro', t2(A.preN), t2(A.AN139)], ['Puesta a tierra', t2(A.preN), t2(A.AP139)], ['Distancia (m)', r2(A.M139)],
        ['Voltaje nominal (V)', r.V], ['Voltaje en bornes (V)', r2(A.AW139)], ['Caída de voltaje (V)', r2(A.AX139)], ['Caída de voltaje (%)', r2(A.AY139)]].concat(Resumen.fuentes(r));
    },
  };
  const num = v => (v === '' || v === null || v === undefined || isNaN(Number(v)) ? 1e9 : Number(v));

  App.views.resumen = function (view, R, sub) {
    sub = sub || 'horizontal';
    const lista = Resumen.ordenados(R);
    const Z = Hoja.zoomUI('resumen');
    view.appendChild(h('div', { class: 'page-h row no-print' }, h('div', null, h('h2', null, 'Tablas resumen'), h('p', { class: 'muted' }, 'Tablas para los planos de Revit. Se actualizan solas; ya no hace falta correr macros.')),
      h('div', { class: 'toolbar' }, h('label', { class: 'chk', title: 'Sin datos del tablero, del supresor ni del interruptor principal' }, h('input', { type: 'checkbox', checked: !!Store.project.resumenSimple, onchange: e => { Store.project.resumenSimple = e.target.checked; Store.save(); App.refresh(); } }), ' Solo tabla resumen'), Z.barra(), UI.btn('Exportar a Excel', () => ExportExcel.download(R), 'primary small'),
        UI.btn('CSV (tabla horizontal)', () => U.download(App.fileBase() + ' - tabla resumen.csv', U.csv([Resumen.grupos.flatMap(([g, n]) => [g].concat(Array(n - 1).fill(''))), Resumen.sub].concat(lista.map(Resumen.horizontal))), 'text/csv'), 'small'),
        UI.btn('Imprimir', () => window.print(), 'small'))));

    // orden / exclusión (hoja ORDEN TABLEROS)
    view.appendChild(h('details', { class: 'card fold no-print' }, h('summary', { class: 'card-h' }, h('h3', null, 'Orden y tableros incluidos (' + lista.length + ' de ' + R.orden.length + ')')),
      h('div', { class: 'card-b' }, h('p', { class: 'muted' }, 'Número para ordenar; desmarque para excluir de las tablas. Sin número se ubican al final en el orden de alimentación.'),
        h('table', { class: 'tbl narrow' }, h('thead', null, h('tr', null, h('th', null, 'Tablero'), h('th', null, 'Orden'), h('th', null, 'Incluir'))),
          h('tbody', null, R.orden.map(r => h('tr', null, h('td', null, r.nombre), h('td', null, UI.input(r.tab, 'orden', { type: 'num', class: 'w-num', fk: 'ord:' + r.tab.id })),
            h('td', null, h('input', { type: 'checkbox', checked: !r.tab.excluirResumen, onchange: e => { r.tab.excluirResumen = !e.target.checked; Store.save(); App.refresh(); } })))))))));

    const tabs = [['horizontal', 'Horizontal'], ['vertical', 'Vertical'], ['du', 'DU (vertical)'], ['original', 'Original']];
    view.appendChild(h('nav', { class: 'chips no-print' }, tabs.map(([k, l]) => h('a', { class: 'chip' + (sub === k ? ' active' : ''), href: '#resumen/' + k }, l))));

    const hoja = sub === 'horizontal' ? Hoja.resumen(lista, Store.project.resumenSimple)
      : sub === 'vertical' ? Hoja.vertical(lista, Resumen.vertical, 'TABLA RESUMEN - TABLEROS ELÉCTRICOS')
        : sub === 'du' ? Hoja.vertical(lista, Resumen.du, 'TABLA RESUMEN DU - TABLEROS ELÉCTRICOS')
          : Hoja.vertical(lista, Resumen.original, 'TABLA RESUMEN - TABLEROS ELÉCTRICOS');
    const wrap = lista.length ? Hoja.html(hoja, Z.opts()) : null;
    view.appendChild(h('article', { class: 'hoja' }, wrap || h('div', { class: 'empty' }, 'Sin tableros.')));
    if (wrap) Z.montar(wrap);
    if (sub === 'original') view.appendChild(h('p', { class: 'hint no-print' }, 'Los datos del transformador se toman del transformador de cada tablero (memoria de cálculo o diagrama unifilar).'));
  };
  window.Resumen = Resumen;
})();
