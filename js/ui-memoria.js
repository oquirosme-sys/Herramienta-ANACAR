/* Memoria de cálculo por tablero (equivale a un bloque de la hoja ANACAR): datos del tablero, circuitos ramales,
   factores de demanda, alimentador con caída de voltaje en cascada, cortocircuito y balanceo. */
(function () {
  'use strict';
  const h = U.h, f2 = v => U.fmt(v, 2), f1 = v => U.fmt(v, 1), f0 = v => U.fmt(v, 0);
  const pos = c => (c.polos || []).join(',');

  /** Selector de tableros en orden de alimentación. */
  function selector(R, id) {
    return h('nav', { class: 'chips', 'aria-label': 'Tableros' }, R.orden.map(r => h('a', {
      class: 'chip' + (r.tab.id === id ? ' active' : '') + (r.avisos.length ? ' warn' : ''), href: '#memoria/' + r.tab.id,
      style: { marginLeft: r.nivel ? '0' : null }, title: r.avisos.join('\n') || null,
    }, (r.nivel ? '↳ '.repeat(1) : '') + r.nombre)));
  }

  const detOpts = () => {
    const tipos = {}; Store.catalog.tiposCarga.forEach(t => { tipos[t.id] = t.nombre; });
    return Store.catalog.detallesCarga.map(d => ({ value: d.id, label: d.id + ' · ' + d.descripcion + ' (' + d.v + ' V ' + d.fases + 'F)', group: tipos[d.tipo] || 'Tipo ' + d.tipo }));
  };
  const bkOpts = (marca, auto) => [{ value: '', label: auto ? 'Auto: ' + auto : 'Automático' }].concat(Store.catalog.breakers.filter(b => b.marca === marca).map(b => ({
    value: b.id, label: b.id + ' · ' + (b.modelo || '—') + ' ' + (b.amperios || '') + ' A ' + (b.unidad || '') + ' ' + (b.polos || '') + 'P ' + (b.sccr || '') + ' kA @' + (b.vSccr || '') + ' V', group: (b.polos || '?') + ' polos',
  })));

  function datosTablero(t, r, R) {
    const L = Store.catalog.listas, k = 'd:' + t.id + ':', C = Store.catalog;
    const marcaOpts = UI.opts(Store.catalog.marcas, 'Del proyecto (' + Store.project.marcaDefecto + ')');
    const catOpts = [{ value: '', label: r.cat ? 'Auto: ' + r.cat.modelo + ' (' + r.cat.fabricante + ')' : 'Automático' }].concat(C.tablerosCat.map(x => ({ value: x.id, label: x.id + ' · ' + x.modelo + ' · ' + x.barraFase + ' A · ' + x.espacios + ' esp.', group: x.fabricante })));
    const spdOpts = [{ value: '', label: r.spd ? 'Auto: ' + r.spd.modelo : 'Automático' }].concat(C.supresores.filter(x => x.marca === r.marca).map(x => ({ value: x.id, label: x.id + ' · ' + x.modelo + ' · ' + x.voltaje + ' V ' + x.fases + 'F · ' + x.kaLL + ' kA', group: x.montaje })));
    const padres = Store.posiblesPadres(t);
    return UI.card('Datos del tablero', h('div', { class: 'grid4' },
      UI.field('Prefijo', UI.input(t, 'prefijo', { fk: k + 'pre' })), UI.field('Nombre / ID', UI.input(t, 'nombre', { fk: k + 'nom' })),
      UI.field('Tipo', UI.select(['3F', '1F'], t.tipo, v => { t.tipo = v; Object.assign(t, Store.voltajeDe(t.sistema, v)); Store.save(); App.refresh(); }, { fk: k + 'tipo' })),
      UI.field('Sistema (V)', UI.select(L.sistemas, t.sistema, v => { t.sistema = v; Object.assign(t, Store.voltajeDe(v, t.tipo)); Store.save(); App.refresh(); }, { fk: k + 'sis', disabled: !!t.fuenteId }), (t.fuenteId ? 'Heredado del secundario de la fuente · ' : '') + 'Voltaje nominal ' + t.voltaje + ' V · ' + t.fases + ' fases · ' + t.hilos + ' hilos'),
      UI.field('Fases', UI.bind(t, 'fases', [3, 2, 1], { num: true, fk: k + 'fas' })), UI.field('Hilos', UI.bind(t, 'hilos', [4, 3, 2], { num: true, fk: k + 'hil' })),
      UI.field('% Reserva', UI.input(t, 'reserva', { type: 'num', fk: k + 'res' }), 'Fracción: 0,1 = 10 %'),
      UI.field('Factor diversidad del tablero', UI.input(t, 'fdivTablero', { type: 'num', fk: k + 'fdv', placeholder: '1', rango: [0, 1] }), 'Entre 0 y 1 (multiplica la demanda)'),
      UI.field('Alimentado desde', UI.select(UI.opts(Store.opcionesAlimentacion(t), '— Acometida / externo —'), Store.valorAlimentacion(t), v => { Store.alimentarDesde(t, v); App.refresh(); }, { fk: k + 'pad' }), (Store.project.fuentes || []).length ? null : 'Los transformadores y generadores compartidos se crean en Proyecto → Fuentes'),
      !t.padreId ? UI.field('Conectado a', UI.input(t, 'conectadoA', { fk: k + 'con', placeholder: 'MÓDULO MEDIDORES' })) : UI.field('Circuito en ' + (r.padre ? r.padre.nombre : ''), h('div', { class: 'ro' }, r.circuitoPadre ? 'Posición ' + pos(r.circuitoPadre.c) : 'Sin circuito')),
      UI.field(t.alterna.activo && /ATS|MTS/.test(t.alterna.equipo || 'ATS') ? 'Longitud ' + (t.alterna.equipo || 'ATS') + ' → tablero (m)' : 'Longitud alimentador (m)', UI.input(t, 'longitud', { type: 'num', fk: k + 'lon' })),
      UI.field('Icc disponible (kA)', UI.input(t, 'iccManual', { type: 'num', fk: k + 'icc', placeholder: r.iccKA ? f2(r.iccKA) + ' (' + r.iccFuente + ')' : 'Ingrese o elija transformador' }), t.padreId ? 'Vacío = cascada desde ' + (r.padre ? r.padre.nombre : 'el padre') : 'En bornes del tablero'),
      UI.field('Transformador aguas arriba', h('div', { class: 'row' }, h('span', { class: 'ro grow' }, r.trafo ? (r.trafo.nombre ? r.trafo.nombre + ' · ' : '') + U.fmt(r.trafo.kva, 1) + ' kVA · Z ' + f2(r.trafo.z) + ' %' : 'Ninguno'), r.trafo && r.trafo.compartido ? null : UI.btn(r.trafo ? 'Editar' : '+ Agregar', () => Unifilar.editarTrafo(t, App.R), 'small')),
        r.trafo ? (r.trafo.compartido ? 'Fuente compartida (' + r.trafo.salidas + ' salida(s), ' + f1(r.trafo.kvaServidos) + ' kVA servidos) · ' : '') + 'Carga ' + f1(r.trafo.carga * 100) + ' % · Icc secundario ' + f2(r.trafo.iccSec / 1000) + ' kA · ΔV ' + f2(r.trafo.reg) + ' %' : 'Cambio de voltaje o acometida en media tensión'),
      UI.field('Tipo de equipo', UI.bind(t, 'clase', [{ value: 'tablero', label: 'Tablero (panelboard)' }, { value: 'subestacion', label: 'Subestación / switchboard (QED-2, I-Line, PRL4, Pow-R-Line Xpert)' }], { fk: k + 'cla' })),
      UI.field('Principal', UI.bind(t, 'principal', [{ value: 'interruptor', label: 'Interruptor principal' }, { value: 'zapatas', label: 'Zapatas (main lugs)' }], { fk: k + 'pri' }), t.principal === 'zapatas' ? 'Protegido por el interruptor del tablero que lo alimenta' : null),
      UI.field('Espacios (polos)', UI.bind(t, 'espacios', [{ value: '', label: 'Auto (' + r.espaciosUsados + ' usados + ' + (Store.project.reservaEspacios || 0) + ' %)' }].concat(Array.from(new Set(C.tablerosCat.filter(x => x.fabricante === r.marca && n(x.espacios) < 999).map(x => n(x.espacios)))).sort((a, b) => a - b).map(v => ({ value: v, label: v + ' espacios' }))), { num: true, fk: k + 'esp' })),
      UI.field('Marca', UI.bind(t, 'marca', marcaOpts, { fk: k + 'mar' })),
      UI.field('Tablero (catálogo)', UI.bind(t, 'catalogoId', catOpts, { fk: k + 'cat' }), r.cat ? (r.famT ? 'Familia ' + r.famT.familia + ' · ' : '') + 'Barras ' + r.cat.barraFase + '/' + r.cat.barraNeutro + '/' + r.cat.barraTierra + ' A · ' + r.cat.espacios + ' espacios (usa ' + r.espaciosUsados + ', reserva ' + (Store.project.reservaEspacios || 0) + ' %)' : null, 'span2'),
      UI.field('Montaje', UI.bind(t, 'montaje', L.montajes, { fk: k + 'mon' })),
      UI.field('Supresor (SPD)', UI.bind(t, 'supresorId', spdOpts, { fk: k + 'spd' }), r.spd ? r.spd.montaje + ' · ' + r.spd.kaLL + '/' + r.spd.kaLN + ' kA' : null),
    ));
  }

  const n = v => Number(v) || 0;
  /** UPS en la alimentación normal y segunda acometida (generador o bypass) con ATS / MTS / interruptor. */
  function alimentacion(t, r) {
    const u = t.ups, a = t.alterna, k = 'f:' + t.id + ':', L = Store.catalog.listas, A = r.alterna;
    const chk = (obj, key, txt) => h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: !!obj[key], onchange: e => { obj[key] = e.target.checked; Store.save(); App.refresh(); } }), ' ' + txt);
    const fila = (l, v, cls) => h('tr', null, h('th', null, l), h('td', { class: cls || '' }, v));
    const otros = Store.project.tableros.filter(o => o.id !== t.id);
    const gens = (Store.project.fuentes || []).filter(x => x.tipo === 'generador').map(x => ({ value: x.id, label: (x.nombre || '') + (Number(x.kva) ? ' · ' + x.kva + ' kVA' : '') + (x.sistema ? ' · ' + x.sistema + ' V' : '') }));
    const propio = !!(t.ups.activo && (t.ups.propio === undefined ? !t.padreId : t.ups.propio));
    const antecesores = new Set(); for (let o = Store.tablero(t.padreId); o && !antecesores.has(o.id); o = Store.tablero(o.padreId)) antecesores.add(o.id);
    return UI.card('Fuentes de alimentación', h('div', { class: 'cols' },
      h('div', null,
        h('h4', null, 'Alimentación normal'),
        h('p', { class: 'muted' }, r.trafo ? 'Por transformador ' + (r.trafo.nombre || '') + ' (' + U.fmt(r.trafo.kva, 1) + ' kVA).' : r.padre ? 'Desde ' + r.padre.nombre + '.' : 'Acometida: ' + (t.conectadoA || 'red') + '.'),
        chk(u, 'activo', 'A través de una UPS'),
        u.activo ? h('div', { class: 'grid3 mt' }, UI.field('Nombre', UI.input(u, 'nombre', { fk: k + 'un' })), UI.field('Capacidad (kVA)', UI.input(u, 'kva', { type: 'num', fk: k + 'uk' })),
          UI.field('Icc del inversor (× In)', UI.input(u, 'factorIcc', { type: 'num', fk: k + 'uf', placeholder: '2' }), 'Corriente que entrega en falla')) : null,
        u.activo ? UI.field('Dónde se conecta la UPS', UI.select(UI.opts([{ value: '__propio', label: 'A la barra de este tablero (la UPS cuelga de él)' }].concat(Store.posiblesPadres(t).map(o => ({ value: o.id, label: 'A la barra de ' + Calc.nombreTablero(o) + ' (este tablero queda aguas abajo)' }))), '— elegir —'), propio ? '__propio' : (t.padreId || ''), v => {
          if (v === '__propio') u.propio = true; else if (v) { u.propio = false; Store.cambiarPadre(t, v); if (a.activo && a.equipo === 'MTS' && a.tipo === 'tablero') a.origenId = v; }
          Store.save(); App.refresh(); }), 'La UPS cuelga de la barra del tablero elegido, con sus interruptores; el circuito se crea en ese tablero') : null,
        u.activo && !propio ? h('label', { class: 'chk mt', title: 'El tablero que alimenta esta UPS lleva dos circuitos con su interruptor: uno a la UPS y otro al bypass (MTS)' },
          h('input', { type: 'checkbox', checked: !!(a.activo && a.equipo === 'MTS' && a.tipo === 'tablero' && a.origenId === t.padreId), onchange: e => {
            if (e.target.checked) Object.assign(a, { activo: true, equipo: 'MTS', tipo: 'tablero', origenId: t.padreId, polosAts: a.polosAts || 3 }); else if (a.equipo === 'MTS' && a.tipo === 'tablero') a.activo = false;
            Store.save(); App.refresh(); } }), ' La UPS lleva MTS (bypass de mantenimiento desde ' + (r.padre ? r.padre.nombre : 'el tablero que la alimenta') + ')') : null,
        u.activo && propio ? chk(u, 'mts', 'La UPS lleva MTS (bypass de mantenimiento): dos acometidas en la barra de este tablero') : null,
        u.activo && propio ? UI.field('Tablero conectado aguas abajo de la UPS / MTS', UI.select(UI.opts(Store.project.tableros.filter(o => o.id !== t.id && !antecesores.has(o.id)).map(o => ({ value: o.id, label: Calc.nombreTablero(o) })), '— ninguno —'), (Store.project.tableros.find(o => o.padreId === t.id && o.deUps) || {}).id || '', v => { Store.project.tableros.forEach(o => { o.deUps = o.id === v; }); if (v) { const d = Store.tablero(v); if (d.padreId !== t.id) Store.cambiarPadre(d, t.id); } Store.save(); App.refresh(); }), 'Se dibuja bajo el MTS de la UPS con su formato normal; se conecta a este tablero (su circuito y su carga se suman aquí)') : null,
        r.ups ? h('p', { class: 'hint' }, 'Carga ' + U.fmt(r.ups.carga * 100, 1) + ' % · Icc limitado a ' + f2(r.ups.icc / 1000) + ' kA · voltaje regulado a la salida') : null),
      h('div', null,
        h('h4', null, 'Segunda acometida'),
        chk(a, 'activo', 'Lleva segunda acometida (generador o bypass)'),
        a.activo ? h('div', { class: 'grid3 mt' },
          UI.field('Equipo', UI.bind(a, 'equipo', [{ value: 'ATS', label: 'Transferencia automática (ATS)' }, { value: 'MTS', label: 'Transferencia manual (MTS)' }, { value: 'IP', label: 'Interruptor principal (enclavado)' }], { fk: k + 'eq' })),
          UI.field('Fuente', UI.bind(a, 'tipo', [{ value: 'generador', label: 'Generador' }, { value: 'tablero', label: 'Bypass directo desde otro tablero' }], { fk: k + 'ti' })),
          a.tipo === 'generador' ? UI.field('Generador del proyecto', UI.bind(a, 'fuenteId', UI.opts(gens, '— Manual (solo este tablero) —'), { fk: k + 'gf' }), gens.length ? 'Un mismo generador puede alimentar varios ATS y tableros' : 'Cree generadores compartidos en Proyecto → Fuentes') : UI.field('Tablero de origen', UI.bind(a, 'origenId', UI.opts(otros.map(o => ({ value: o.id, label: Calc.nombreTablero(o) })), '— elegir —'), { fk: k + 'or' })),
          a.tipo === 'generador' && !a.fuenteId ? UI.field('Nombre', UI.input(a, 'nombre', { fk: k + 'gn' })) : null,
          a.tipo === 'generador' && !a.fuenteId ? UI.field('Generador (kVA)', UI.input(a, 'kva', { type: 'num', fk: k + 'gk' })) : null,
          a.tipo === 'generador' && !a.fuenteId ? UI.field("Reactancia X''d (%)", UI.input(a, 'xd', { type: 'num', fk: k + 'gx', placeholder: '12' })) : null,
          a.equipo === 'ATS' || !a.equipo ? UI.field('Tipo de transferencia', UI.bind(a, 'transicion', [{ value: 'abierta', label: 'Transición abierta (estándar)' }, { value: 'retardada', label: 'Transición retardada (motores)' }, { value: 'cerrada', label: 'Transición cerrada' }, { value: 'bypass', label: 'Con bypass-aislamiento' }], { fk: k + 'tr' })) : null,
          /ATS|MTS/.test(a.equipo || 'ATS') ? UI.field('Polos', UI.bind(a, 'polosAts', [{ value: 3, label: '3 polos (neutro sólido)' }, { value: 4, label: '4 polos (neutro conmutado)' }], { num: true, fk: k + 'po' }), 'Con 4 polos el generador es fuente derivada separadamente (250.30)') : null,
          /ATS|MTS/.test(a.equipo || 'ATS') ? UI.field('Modelo (catálogo)', UI.bind(a, 'atsId', [{ value: '', label: A && A.ats ? 'Auto: ' + A.ats.modelo + ' · ' + A.ats.amperios + ' A · ' + A.ats.wcr + ' kA' : 'Automático' }].concat((Store.catalog.ats || []).filter(x => x.equipo === (a.equipo === 'MTS' ? 'MTS' : 'ATS') && x.transicion === (a.equipo === 'MTS' ? 'manual' : (a.transicion || 'abierta')) && Number(x.polos) === (Number(a.polosAts) || 3)).map(x => ({ value: x.id, label: x.marca + ' · ' + x.modelo + ' · ' + x.amperios + ' A · WCR ' + x.wcr + ' kA', group: x.transicion + ' · ' + x.polos + ' polos' }))), { num: true, fk: k + 'mo' })) : null,
          /ATS|MTS/.test(a.equipo || 'ATS') ? UI.field('WCR / SCCR del ' + (a.equipo || 'ATS') + ' (kA)', UI.input(a, 'sccr', { type: 'num', fk: k + 'sc', placeholder: A && A.ats ? 'Auto ' + A.ats.wcr : '' }), 'Vacío = el del modelo') : null,
          UI.field(/ATS|MTS/.test(a.equipo || 'ATS') ? (a.tipo === 'generador' ? 'Generador' : 'Tablero de origen') + ' → ' + (a.equipo || 'ATS') + ' (m)' : 'Longitud (m)', UI.input(a, 'longitud', { type: 'num', fk: k + 'al' })),
          /ATS|MTS/.test(a.equipo || 'ATS') ? UI.field((r.trafo ? 'Transformador' : r.padre ? r.padre.nombre : t.conectadoA || 'Fuente normal') + ' → ' + (a.equipo || 'ATS') + ' (m)', UI.input(a, 'distNormal', { type: 'num', fk: k + 'dn', placeholder: '0' }), 'Acometida normal hasta el ' + (a.equipo || 'ATS')) : null,
          /ATS|MTS/.test(a.equipo || 'ATS') ? UI.field((a.equipo || 'ATS') + ' → tablero (m)', UI.input(t, 'longitud', { type: 'num', fk: k + 'lon2' }), 'Se toma de la longitud de este tablero, el que se alimenta desde el ' + (a.equipo || 'ATS')) : null,
          UI.field('Material', UI.bind(a, 'material', L.materiales, { fk: k + 'am' })), UI.field('Aislamiento', UI.bind(a, 'aislamiento', Store.catalog.aislamientos, { fk: k + 'aa' })),
          UI.field('Tubería', UI.bind(a, 'tuberia', L.tuberias, { fk: k + 'at' })), UI.field('# en paralelo', UI.input(a, 'paralelos', { type: 'num', fk: k + 'ap', placeholder: A ? 'Auto ' + A.par : 'Auto' }))) : null,
        A ? h('table', { class: 'kv mt' }, h('tbody', null,
          fila('Origen', A.origen || '—'),
          A.compartido ? fila('Generador compartido', A.salidas + ' salida(s) · el % de carga suma todos los tableros que cuelgan de él') : null,
          fila('Conductores', h('b', null, A.fasesTxt + (A.neutroTxt ? ' + ' + A.neutroTxt + ' N' : '') + (A.tierraTxt ? ' + ' + A.tierraTxt + ' T' : '') + ' ' + A.mat + ' ' + A.ais)),
          fila('Tubería', A.tubo + ' mm ' + A.tuberia),
          fila('Ampacidad corregida', A.val ? f1(A.val.cap) + ' A ' + (A.val.ok ? '✓' : '✗') : '—', A.val && !A.val.ok ? 'bad' : ''),
          A.ats ? fila(A.equipo + ' (catálogo)', h('b', null, A.ats.modelo + ' · ' + A.ats.amperios + ' A · ' + A.ats.polos + ' polos · transición ' + A.ats.transicion)) : null,
          /ATS|MTS/.test(A.equipo) ? fila('Acometida normal → ' + A.equipo, f1(A.distNormal) + ' m · ΔV ' + f2(A.caidaNormal) + ' V') : null,
          /ATS|MTS/.test(A.equipo) ? fila((A.tipo === 'generador' ? 'Generador' : 'Bypass') + ' → ' + A.equipo, f1(A.longitud) + ' m · ΔV ' + f2(A.caidaFuente) + ' V') : null,
          /ATS|MTS/.test(A.equipo) ? fila(A.equipo + ' → tablero (tramo común)', f1(A.distAts) + ' m · ΔV ' + f2(A.caidaComun) + ' V') : null,
          fila('Caída / voltaje en bornes', f2(A.caida) + ' V · ' + f2(A.vBornes) + ' V · ΔV ' + f2(A.cv) + ' %', A.cv > Store.project.cvMaxAlim ? 'bad' : ''),
          /ATS|MTS/.test(A.equipo) ? fila('Icc en el ' + A.equipo + ' / WCR', (A.iccAtsKA ? f2(A.iccAtsKA) + ' kA' : '—') + ' · WCR ' + (A.wcr || '—') + ' kA ' + (A.wcrOk ? '✓' : '✗'), A.wcrOk ? '' : 'bad') : null,
          fila('Icc por esta fuente (en bornes)', A.iccA ? f2(A.iccA / 1000) + ' kA' : '—'),
          A.carga !== null ? fila('Carga del generador', f1(A.carga * 100) + ' %', A.carga > 1 ? 'bad' : '') : null)) : null)));
  }

  /** Revisión NEC 2020 del tablero. */
  function revisionNec(r) {
    if (!r.nec || !r.nec.length) return null;
    const ico = { error: '✗', aviso: '⚠', info: 'ℹ' };
    return UI.card('Revisión NEC 2020 (' + r.nec.filter(x => x.nivel !== 'info').length + ' por atender)', h('ul', { class: 'nec' }, r.nec.map(x => h('li', { class: x.nivel }, h('b', null, ico[x.nivel] + ' ' + x.art + ' '), x.txt))),
      h('small', { class: 'muted' }, 'Según los resúmenes de cambios NEC 2020 de Eaton/CITEC/CIEMI'));
  }

  function alimentador(t, r) {
    const a = t.alim, A = r.alim, L = Store.catalog.listas, k = 'a:' + t.id + ':';
    const fila = (l, v, cls) => h('tr', { class: cls || '' }, h('th', null, l), h('td', null, v));
    return UI.card('Alimentador / acometida', h('div', null, h('div', { class: 'cols' },
      h('div', { class: 'grid3' },
        UI.field('Material', UI.bind(a, 'material', L.materiales, { fk: k + 'mat' })),
        UI.field('Aislamiento', UI.bind(a, 'aislamiento', Store.catalog.aislamientos, { fk: k + 'ais' })),
        UI.field('Tubería', UI.bind(a, 'tuberia', L.tuberias, { fk: k + 'tub' })),
        UI.field('# en paralelo', UI.input(a, 'paralelos', { type: 'num', fk: k + 'par', placeholder: 'Auto ' + A.AI139 })),
        UI.field('Aumento de calibre (%)', UI.pct(a, 'aumento', { modo: 'aumento', fk: k + 'aum', placeholder: '0' }), 'Sobre la ampacidad de la protección'),
        UI.field('Factor de potencia', UI.input(a, 'fp', { type: 'num', fk: k + 'fp' })),
        UI.field('Factor de protección (%)', a.rated100 ? h('div', { class: 'ro' }, '100 % (interruptor 100 % rated)') : UI.pct(a, 'mult', { fk: k + 'mul', placeholder: '125' }), '125 % carga continua'),
        UI.field('Protección (A)', UI.input(a, 'prot', { type: 'num', fk: k + 'pro', placeholder: 'Auto ' + (A.AF139 || '') })),
        t.principal === 'zapatas' ? null : UI.field('Unidad de disparo', UI.bind(a, 'unidad', [{ value: '', label: 'Termomagnético (TM)' }, { value: 'LI', label: 'Electrónico LI' }, { value: 'LSI', label: 'Electrónico LSI' }, { value: 'LSIG', label: 'Electrónico LSIG' }], { fk: k + 'uni' })),
        t.principal === 'zapatas' ? null : UI.field('Capacidad', h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: !!a.rated100, onchange: e => { a.rated100 = e.target.checked; Store.save(); App.refresh(); } }), ' 100 % rated'), 'Lleva el 100 % de la carga continua (215.3 / 230.42 exc.)'),
        t.principal === 'zapatas' ? UI.field('Interruptor principal', h('div', { class: 'ro' }, 'Zapatas (sin interruptor principal)'), null, 'span3') : UI.field('Interruptor principal', UI.bind(a, 'breakerId', bkOpts(r.marca, r.bkMain && !a.breakerId ? (r.bkMain.ref || r.bkMain.modelo) + ' · ' + r.bkMain.sccr + ' kA' : ''), { fk: k + 'bk' }), r.famT ? 'Familias permitidas en ' + r.famT.familia + ': ' + r.famT.principales : null, 'span3'),
        UI.field('Temp. ambiente (°C)', UI.bind(a, 'tempAmb', L.temperaturas, { fk: k + 'ta' })),
        UI.field('Temp. bornes (°C)', UI.bind(a, 'tempBorne', [60, 75, 90], { num: true, fk: k + 'tb' }), 'Columna 310.16 que limita'),
        UI.field('Cond. portadores (agrup.)', UI.bind(a, 'agrupamiento', [{ value: '', label: 'Auto (' + A.agrup + ')' }].concat(L.agrupamientos.map(x => ({ value: x, label: x }))), { fk: k + 'ag' })),
        UI.field('Calibre de fases (manual)', UI.bind(a, 'calFases', [{ value: '', label: 'Auto (' + (A.auto ? A.auto.fases : '') + ')' }].concat(Calc.CALIBRES.map(x => ({ value: x, label: x }))), { fk: k + 'cf' }), 'Si elige uno, se respeta y solo se avisa si no cumple'),
        UI.field('Calibre de neutro (manual)', UI.bind(a, 'calNeutro', [{ value: '', label: 'Auto (' + (A.auto ? A.auto.neutro || '—' : '') + ')' }].concat(Calc.CALIBRES.map(x => ({ value: x, label: x }))), { fk: k + 'cn' })),
        UI.field('Calibre de tierra (manual)', UI.bind(a, 'calTierra', [{ value: '', label: 'Auto (' + (A.auto ? A.auto.tierra : '') + ')' }].concat(Calc.CALIBRES.map(x => ({ value: x, label: x }))), { fk: k + 'ct' }), 'Mínimo 10 AWG en alimentadores'),
        UI.field('Calibre', h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: !!a.calibreFijo, onchange: e => { a.calibreFijo = e.target.checked; Store.save(); App.refresh(); } }), ' No aumentar automáticamente'), null, 'span3'),
      ),
      h('table', { class: 'kv' }, h('tbody', null,
        fila('Corriente demandada A / B / C (A)', [0, 1, 2].slice(0, r.fases === 3 ? 3 : 2).map(i => f1(A.AA139[i])).join(' / ')),
        fila('Ampacidad requerida (A)', f2(A.AE139)),
        fila('Protección (A)', A.AF139 || '—'),
        fila('Conductores', h('b', null, A.fasesTxt + (A.neutroTxt ? ' + ' + A.neutroTxt + ' N' : '') + (A.tierraTxt ? ' + ' + A.tierraTxt + ' T' : '') + ' AWG ' + A.mat + ' ' + A.ais)),
        fila('Tubería', (A.preN || '') + A.AR139 + ' mm ' + A.tuberia),
        fila('Conductor del electrodo (250.66)', A.AP140 ? A.AP140 + ' AWG' : '—'),
        fila('FAC tabla / ajustado', f2(A.AS139) + ' / ' + f2(A.AT139)),
        fila('Voltaje de partida (V)', h('span', null, f2(A.vInicio), r.padre && A.vInicio !== r.V ? h('small', { class: 'muted' }, ' (bornes de ' + r.padre.nombre + ')') : null)),
        fila('Caída en el alimentador (V)', f2(A.AV139)),
        fila('Voltaje en bornes L-L / L-N (V)', f2(A.AW139) + ' / ' + f2(A.AW140)),
        fila('Caída acumulada (V / %)', h('b', { class: A.AY139 > Store.project.cvMaxAlim ? 'bad' : '' }, f2(A.AX139) + ' V · ' + f2(A.AY139) + ' %')),
        fila('Interruptor principal', r.bkMain ? h('span', { class: r.bkMain.sccrBajo ? 'bad' : '' }, (r.bkMain.ref || r.bkMain.modelo || '—') + ' · ' + (A.AF139 || '') + ' A · ' + r.bkMain.unidad + ' · ' + r.bkMain.polos + 'P · ' + r.bkMain.sccr + ' kA') : '—'),
        r.bkMain && r.bkMain.sel ? fila('Interruptor: catálogo', r.bkMain.sel.desc + ' · terminal ' + r.bkMain.sel.terminal) : null,
        r.bkMain && r.bkMain.bornes ? fila('Bornes vs cableado', r.bkMain.bornes.ok ? '✓ ' + r.bkMain.bornes.actual : r.bkMain.bornes.ok === false ? '⚠ ' + r.bkMain.bornes.msg : r.bkMain.bornes.msg, r.bkMain.bornes.ok === false ? 'bad' : '') : null,
        fila('Cortocircuito en bornes', r.iccKA ? f2(r.iccKA) + ' kA (' + r.iccFuente + ')' : '—'),
      ))), validacion(t, r)));
  }

  /** Tabla de validación del alimentador por temperatura y agrupamiento (NEC 310.15 y 310.16). */
  function validacion(t, r) {
    const A = r.alim, v = A.val;
    if (!v) return null;
    const fila = (l, x, cls) => h('tr', null, h('th', null, l), h('td', { class: cls || '' }, x));
    const est = v.ok ? h('span', { class: 'badge ok' }, 'Cumple') : h('span', { class: 'badge bad' }, 'No cumple');
    return h('div', { class: 'valid' }, h('h4', null, 'Validación por temperatura y agrupamiento (310.15 / 310.16) ', est),
      h('table', { class: 'kv' }, h('tbody', null,
        fila('Calibre', (A.AI139 > 1 ? A.AI139 + ' × ' : '') + v.cal + ' ' + A.mat + ' ' + A.ais + (A.ajustado ? ' (antes ' + (A.parBase > 1 ? A.parBase + ' × ' : '') + A.calBase + ')' : '')),
        fila('Ampacidad a ' + v.tA + ' °C (aislamiento)', f1(v.amp90) + ' A'),
        fila('Factor de temperatura ' + (t.alim.tempAmb || '26-30') + ' °C · Ft', U.fmt(v.ft, 2), v.ft ? '' : 'bad'),
        fila('Factor de agrupamiento ' + A.agrup + ' conductores · Fg', U.fmt(v.fg, 2)),
        fila('Ampacidad corregida (× ' + A.AI139 + ' en paralelo)', f1(v.corr) + ' A'),
        fila('Límite por bornes a ' + (t.alim.tempBorne || 75) + ' °C', f1(v.term * A.AI139) + ' A'),
        fila('Ampacidad disponible', h('b', null, f1(v.cap) + ' A')),
        fila('Ampacidad requerida (I × ' + U.fmt(A.AD139, 2) + ')', f1(A.AE139) + ' A', v.okI ? 'okc' : 'bad'),
        fila('Protección (240.4(B) permite el siguiente tamaño estándar)', (A.AF139 || '—') + ' A', v.okP ? 'okc' : 'bad'))));
  }

  /** Barra de desplazamiento horizontal también arriba de la tabla, sincronizada con la de abajo. */
  function conBarraSuperior(wrap) {
    const inner = h('div', { style: { height: '1px' } }), top = h('div', { class: 'scroll-top', 'aria-hidden': 'true' }, inner);
    const sync = () => { inner.style.width = wrap.scrollWidth + 'px'; top.style.display = wrap.scrollWidth > wrap.clientWidth + 1 ? 'block' : 'none'; if (top.scrollLeft !== wrap.scrollLeft) top.scrollLeft = wrap.scrollLeft; };
    top.addEventListener('scroll', () => { if (wrap.scrollLeft !== top.scrollLeft) wrap.scrollLeft = top.scrollLeft; });
    wrap.addEventListener('scroll', () => { if (top.scrollLeft !== wrap.scrollLeft) top.scrollLeft = wrap.scrollLeft; });
    requestAnimationFrame(sync); setTimeout(sync, 150); setTimeout(sync, 600);
    if (window.ResizeObserver) { const ro = new ResizeObserver(sync); ro.observe(wrap); if (wrap.firstElementChild) ro.observe(wrap.firstElementChild); }
    return h('div', null, top, wrap);
  }

  function circuitos(t, r, R) {
    const k = c => 'c:' + c.id + ':', L = Store.catalog.listas, dOpts = detOpts();
    const cr = Store.project.pais === 'PA', DEV = Simbolos.disp.map((d, i) => [d.campo, d.nombre, 'dv' + i, d.clave]);
    const head = h('thead', null,
      h('tr', { class: 'grp' }, h('th', { colspan: 12 }, 'Datos de circuitos ramales / alimentadores'), cr ? h('th', { colspan: DEV.length, class: 'dev' }, 'Cantidad de dispositivos') : null, h('th', { colspan: r.fases === 3 ? 3 : 2 }, 'Balance (kVA)'),
        h('th', { colspan: 8 }, 'Cálculo de conductores'), h('th', { colspan: 3 }, 'Caída de voltaje'), h('th', { colspan: 5 }, 'Interruptor ramal'), h('th', null, '')),
      h('tr', null, ['Fijo', 'Posición', 'Detalle de carga', 'Descripción', 'kVA', 'Long. (m)', 'V', 'F', 'Prot. %', 'F. uso', 'F. div.', 'I (A)'].concat(cr ? DEV.map(d => ({ sim: d })) : [], r.fases === 3 ? ['A', 'B', 'C'] : ['A', 'B'],
        ['Amp. req.', 'Prot. (A)', 'Material', 'Aislam.', '# par.', 'Aum. %', 'Calibres F / N / T', 'Tubo (mm)', 'V real', 'ΔV (V)', 'ΔV total %', 'Unidad', 'Tipo (catálogo)', 'Modelo de referencia', 'Polos', 'SCCR', '']).map(x => x && x.sim ? h('th', { class: 'dev sim', title: x.sim[1] }, Object.assign(h('span', { class: 'simb' }), { innerHTML: Simbolos.svg(x.sim[3], 22, 33) })) : h('th', null, x))));
    const body = h('tbody');
    r.rows.forEach(x => {
      const c = x.c, kk = k(c), resp = c.respaldoDe && Store.tablero(c.respaldoDe), hijo = (c.tableroHijoId && Store.tablero(c.tableroHijoId)) || resp;
      const polos = UI.input({ v: pos(c) }, 'v', { fk: kk + 'pol', class: 'w-pos', label: 'Posición', save: false, after: v => {
        const arr = String(v).split(/[^0-9]+/).map(Number).filter(z => z > 0);
        Store.registrarCambios(t, [{ id: c.id, de: c.polos, a: arr, descripcion: x.descripcion }], 'manual'); App.refresh();
      } });
      body.appendChild(h('tr', { class: (x.err.length ? 'row-warn' : '') + (hijo ? ' row-link' : ''), title: x.err.join('\n') || null },
        h('td', { class: 'c' }, h('input', { type: 'checkbox', checked: !!c.fijo, title: 'Fijar posición (el autobalanceo no lo mueve)', onchange: e => { c.fijo = e.target.checked; Store.save(); } })),
        h('td', { class: 'sticky' }, polos),
        h('td', null, hijo ? h('a', { href: '#memoria/' + hijo.id, class: 'link' }, (resp ? '↺ bypass de ' : '→ ') + Calc.nombreTablero(hijo)) : null, UI.bind(c, 'detalleId', dOpts, { num: true, fk: kk + 'det', class: 'w-det', label: 'Detalle de carga' })),
        h('td', null, hijo ? h('span', { class: 'ro' }, x.descripcion) : UI.input(c, 'descripcion', { fk: kk + 'des', class: 'w-desc', placeholder: x.det.descripcion || '' })),
        h('td', null, hijo ? h('span', { class: 'ro r' }, f2(x.J)) : UI.input(c, 'kva', { type: 'num', fk: kk + 'kva', class: 'w-num' })),
        h('td', null, hijo ? h('span', { class: 'ro r' }, f1(c.longitud)) : UI.input(c, 'longitud', { type: 'num', fk: kk + 'lon', class: 'w-num' })),
        h('td', { class: 'r' }, x.O || ''), h('td', { class: 'r' }, x.P || ''),
        h('td', null, UI.pct(c, 'mult', { fk: kk + 'mul', class: 'w-xs', placeholder: U.fmt(x.AD * 100, 0) })),
        h('td', null, UI.input(c, 'fu', { type: 'num', fk: kk + 'fu', class: 'w-xxs', placeholder: f2(x.R), title: 'Factor de uso (demanda) del circuito' })),
        h('td', null, UI.input(c, 'fdivC', { type: 'num', fk: kk + 'fdv', class: 'w-xxs', placeholder: f2(x.S), title: 'Factor de diversidad del circuito (0 a 1)', rango: [0, 1] })),
        h('td', { class: 'r' }, x.J ? f1(Math.max(...x.I)) : ''),
        cr ? DEV.map(d => h('td', { class: 'dev' }, hijo ? null : UI.input(c, d[0], { type: 'num', fk: kk + d[2], class: 'w-xxs', label: d[1] }))) : null,
        x.fase.slice(0, r.fases === 3 ? 3 : 2).map((v, i) => h('td', { class: 'r ph ph' + i }, v ? f2(v) : '')),
        h('td', { class: 'r' }, x.AE ? f1(x.AE) : ''),
        h('td', null, UI.input(c, 'prot', { type: 'num', fk: kk + 'pro', class: 'w-xs', placeholder: x.AF && !c.prot ? String(x.AF) : '' })),
        h('td', null, UI.bind(c, 'material', L.materiales, { fk: kk + 'mat', class: 'w-xs' })),
        h('td', null, UI.bind(c, 'aislamiento', Store.catalog.aislamientos, { fk: kk + 'ais', class: 'w-ais' })),
        h('td', null, UI.input(c, 'paralelos', { type: 'num', fk: kk + 'par', class: 'w-xxs', placeholder: String(x.AI) })),
        h('td', null, UI.pct(c, 'aumento', { modo: 'aumento', fk: kk + 'aum', class: 'w-xxs', placeholder: '0' })),
        h('td', { class: 'nowrap' }, x.AL ? x.fasesTxt + (x.neutroTxt ? ' · ' + x.neutroTxt : '') + (x.tierraTxt ? ' · ' + x.tierraTxt : '') : ''),
        h('td', { class: 'r' }, x.tuboTxt), h('td', { class: 'r' }, x.AU ? f1(x.AU) : ''), h('td', { class: 'r' }, x.AV !== null ? f2(x.AV) : ''),
        h('td', { class: 'r ' + (x.AY > Store.project.cvMaxTotal ? 'bad' : '') }, x.AY !== null ? f2(x.AY) : ''),
        h('td', { title: x.nec ? 'NEC ' + x.nec.motivo : null }, UI.bind(c, 'unidad', [{ value: '', label: x.nec ? 'Auto: ' + x.nec.unidad : 'Auto (STD)' }, { value: 'STD', label: 'STD' }].concat(['GFCI', 'AFCI', 'AFCI/GFCI', 'HACR', 'SHT'].map(u => ({ value: u, label: u }))), { fk: kk + 'uni', class: 'w-xs' })),
        h('td', null, UI.bind(c, 'breakerId', bkOpts(r.marca, x.breaker && !c.breakerId ? (x.breaker.fam || x.breaker.modelo || x.breaker.unidad) : ''), { fk: kk + 'bk', class: 'w-bk' })),
        h('td', { class: x.breaker && x.breaker.sccrBajo ? 'bad' : '' }, x.breaker ? (x.breaker.ref || x.breaker.modelo || '—') + (x.breaker.unidad && x.breaker.unidad !== 'STD' ? ' ' + x.breaker.unidad : '') : ''), h('td', { class: 'r' }, x.J ? x.polosBreaker : ''), h('td', { class: 'r' }, x.breaker ? x.breaker.sccr : ''),
        h('td', { class: 'acc' },
          UI.iconBtn('⧉', 'Duplicar circuito', () => { const n = U.clone(c); n.id = U.uid(); n.tableroHijoId = ''; n.auto = false; n.polos = Store.posicionLibre(t, Math.max(1, c.polos.length)); t.circuitos.push(n); Store.save(); App.refresh(); }),
          UI.iconBtn('🗑', 'Eliminar circuito', async () => {
            if (hijo && !(await UI.confirm('Este circuito alimenta a ' + Calc.nombreTablero(hijo) + '. ¿Eliminarlo? El tablero quedará como acometida/externo.', 'Eliminar'))) return;
            t.circuitos = t.circuitos.filter(z => z !== c); if (hijo) hijo.padreId = ''; Store.save(); App.refresh();
          }, 'danger'))));
    });
    const tot = h('tfoot', null, h('tr', null, h('th', { colspan: 4, class: 'r' }, 'Totales'), h('th', { class: 'r' }, f2(r.J116)), h('th', { colspan: 7 }),
      cr ? DEV.map(d => h('th', { class: 'r dev' }, r.rows.reduce((a, z) => a + (Number(z.c[d[0]]) || 0), 0) || '')) : null,
      r.U116.slice(0, r.fases === 3 ? 3 : 2).map(v => h('th', { class: 'r' }, f2(v))), h('th', { colspan: 17, class: 'l' }, 'Desbalance máximo: ' + f2(r.desbalance) + ' %')));
    const hijosSinCircuito = Store.project.tableros.filter(o => o.padreId === t.id && !t.circuitos.some(c => c.tableroHijoId === o.id));
    return UI.card('Circuitos ramales (' + r.rows.length + ')', h('div', null,
      r.rows.length ? conBarraSuperior(h('div', { class: 'tbl-wrap tall circ', 'data-scroll': 'circ-' + t.id }, h('table', { class: 'tbl calc' }, head, body, tot))) : h('div', { class: 'empty' }, 'Sin circuitos. Agréguelos o impórtelos de Revit.'),
      h('div', { class: 'toolbar' },
        UI.btn('+ Circuito', () => { const c = Store.nuevoCircuito(t); App.refresh(); setTimeout(() => { const e = document.querySelector('[data-fk="c:' + c.id + ':kva"]'); if (e) e.focus(); }, 0); }, 'primary small'),
        UI.btn('+ Tablero derivado…', () => nuevoDerivado(t), 'small', 'Crea un tablero alimentado desde este y su circuito'),
        hijosSinCircuito.length ? UI.btn('Asignar circuito a derivados (' + hijosSinCircuito.length + ')', () => { hijosSinCircuito.forEach(o => Store.cambiarPadre(o, t.id)); App.refresh(); }, 'small') : null,
        UI.btn('⚖ Autobalancear', () => balanceo(t, r), 'primary small', 'Reubica circuitos para reducir el desbalance y anota los cambios para Revit'),
        t.cambiosRevit.some(x => !x.aplicado) ? UI.btn('Cambios para Revit (' + t.cambiosRevit.filter(x => !x.aplicado).length + ')', () => App.go('revit', 'cambios'), 'small') : null,
        h('span', { class: 'grow' }),
        h('small', { class: 'muted' }, 'Posición: polos separados por coma (1,3,5). Campos vacíos usan el valor automático que se muestra en gris.'))));
  }

  function nuevoDerivado(padre) {
    UI.prompt('Nuevo tablero derivado de ' + Calc.nombreTablero(padre), 'Nombre del tablero', '', v => (v ? '' : 'Escriba un nombre')).then(nombre => {
      if (!nombre) return;
      const t = Store.nuevoTablero({ nombre, tipo: padre.tipo, sistema: padre.sistema });
      Store.cambiarPadre(t, padre.id); App.refresh(); UI.toast('Tablero ' + nombre + ' creado y conectado', 'ok');
    });
  }

  function balanceo(t, r) {
    const obj = Store.project.desbalanceMax > 0 ? Math.min(5, Store.project.desbalanceMax) : 0;
    const b = Calc.balanceo(t, Store.catalog, { espacios: r.cat ? r.cat.espacios : 0, maxMov: 40, objetivo: obj });
    const fx = arr => arr.slice(0, r.fases === 3 ? 3 : 2).map(v => f2(v)).join(' / ');
    if (!b.cambios.length) { UI.alert('No se encontraron cambios que mejoren el balance (desbalance ' + U.fmt(b.antes, 2) + ' %). Los circuitos fijados no se mueven.', 'Autobalanceo'); return; }
    UI.modal('Autobalanceo — ' + Calc.nombreTablero(t), h('div', null,
      h('p', null, 'Desbalance ', h('b', null, U.fmt(b.antes, 2) + ' %'), ' (kVA ' + fx(b.fasesAntes) + ') → ', h('b', { class: 'okc' }, U.fmt(b.despues, 2) + ' %'), ' (kVA ' + fx(b.fasesDespues) + ')'),
      h('h4', null, 'Cambios de circuito para Revit (' + b.cambios.length + ')'),
      h('table', { class: 'tbl' }, h('thead', null, h('tr', null, ['Carga', 'Circuito actual', 'Circuito nuevo'].map(x => h('th', null, x)))),
        h('tbody', null, b.cambios.map(c => h('tr', null, h('td', null, c.descripcion), h('td', null, c.de.join(',')), h('td', null, h('b', null, c.a.join(','))))))),
      h('p', { class: 'muted' }, 'Intercambia circuitos con el mismo número de polos y mueve circuitos a espacios libres; no toca los fijados. Al aplicar, los cambios quedan en Revit ▸ Cambios para Revit para pasarlos al modelo.')), [
      { label: 'Cancelar' },
      { label: 'Aplicar y anotar para Revit', cls: 'primary', onclick: () => { Store.registrarCambios(t, b.cambios, 'balanceo'); App.refresh(); UI.toast(b.cambios.length + ' cambios aplicados y anotados para Revit', 'ok'); } },
    ], { wide: true });
  }

  function demanda(t, r) {
    const body = h('tbody');
    r.tipos.filter(x => x.conectados || (t.fd || {})[x.tipo.id] !== undefined).forEach(x => {
      const id = x.tipo.id;
      body.appendChild(h('tr', null, h('td', null, x.tipo.nombre), h('td', { class: 'r' }, x.n), h('td', { class: 'r' }, f2(x.conectados)), h('td', { class: 'r' }, f2(x.reserva)), h('td', { class: 'r' }, f2(x.total)),
        h('td', null, UI.input(t.fd, String(id), { type: 'num', fk: 'fd:' + t.id + ':' + id, class: 'w-num', placeholder: U.fmt(x.fd, 3) + (x.tipo.metodo === 'tomas' ? ' (10 kVA + 50 %)' : '') })),
        h('td', { class: 'r' }, f2(x.fdiv)), h('td', { class: 'r' }, f2(x.demandados))));
    });
    return UI.card('Factores de demanda y diversidad', h('div', { class: 'tbl-wrap' }, h('table', { class: 'tbl' },
      h('thead', null, h('tr', null, ['Tipo de carga', 'Circuitos', 'kVA conectados', 'Reserva', 'kVA totales', 'Factor demanda', 'F. diversidad', 'kVA demandados'].map(x => h('th', null, x)))),
      body, h('tfoot', null, h('tr', null, h('th', null, 'Totales'), h('th'), h('th', { class: 'r' }, f2(r.W128)), h('th', { class: 'r' }, f2(r.W129)), h('th', { class: 'r' }, f2(r.W130)), h('th', { class: 'r' }, f2(r.alim.R139)), h('th'), h('th', { class: 'r' }, f2(r.J134)))))),
      h('div', { class: 'hint' }, h('b', null, 'Cómo se aplican los factores: '),
        '1) En cada circuito, el factor de uso (demanda) y el de diversidad del detalle de carga —o el valor propio del circuito, columnas F. uso y F. div.— dan los kVA demandados y la corriente con la que se dimensionan protección y conductor (I = kVA × F. uso × F. div.; los factores de diversidad van de 0 a 1). ',
        '2) En el tablero, la carga conectada (más la reserva) se agrupa por tipo de carga y se aplica el factor de demanda del tipo (tomas: 100 % de los primeros 10 kVA + 50 % del resto, 220.44) y su diversidad. ',
        '3) El alimentador usa los kVA demandados × el factor de diversidad del tablero (0 a 1). ',
        '4) Los tableros derivados suben al tablero que los alimenta con sus kVA ' + (Store.project.cargaDerivados === 'demandada' ? 'demandados (y ya no se les vuelve a aplicar demanda)' : 'conectados (como el Excel) y ahí se les aplica la demanda del tipo de carga del circuito') + '; se cambia en Proyecto ▸ Criterios. ',
        'El circuito de bypass de otro tablero ocupa espacios pero no suma carga.'));
  }

  App.views.memoria = function (view, R, id) {
    if (!R.orden.length) { view.appendChild(h('div', { class: 'empty' }, 'No hay tableros. Créelos en la pestaña Proyecto o impórtelos de Revit.')); return; }
    const r = R.res[id] || R.orden[0], t = r.tab;
    view.appendChild(selector(R, t.id));
    view.appendChild(h('div', { class: 'page-h row' }, h('div', null, h('h2', null, r.nombre), h('p', { class: 'muted' }, t.tipo + ' · ' + t.sistema + ' V · alimentado desde ' + (r.alimentadoDesde || '—'))),
      h('div', { class: 'toolbar' }, UI.btn('Ver tablero ' + t.tipo, () => App.go(t.tipo === '1F' ? 'tab1f' : 'tab3f', t.id), 'small'),
        UI.btn('Memoria: PDF / Word / Excel', () => App.go('reporte', t.id), 'primary small'))));
    if (r.avisos.length) view.appendChild(UI.avisos(r.avisos));
    const A = r.alim;
    view.appendChild(h('div', { class: 'kpis' },
      UI.kpi('kVA conectados', f2(r.W130), 'con ' + f0(n100(t.reserva)) + ' % reserva'), UI.kpi('kVA demandados', f2(A.L139), 'FD ' + f2(A.R139)),
      UI.kpi('Corriente por fase', [0, 1, 2].slice(0, r.fases === 3 ? 3 : 2).map(i => f0(A.AA139[i])).join(' / ') + ' A'),
      UI.kpi('Desbalance', f2(r.desbalance) + ' %', null, r.desbalance > Store.project.desbalanceMax ? 'warn' : ''),
      UI.kpi('Interruptor principal', (A.AF139 || '—') + ' A', A.fasesTxt + ' AWG'),
      UI.kpi('Voltaje en bornes', f2(A.AW139) + ' V', 'ΔV acumulada ' + f2(A.AY139) + ' %', A.AY139 > Store.project.cvMaxAlim ? 'warn' : ''),
      UI.kpi('Cortocircuito', r.iccKA ? f2(r.iccKA) + ' kA' : '—', r.iccFuente)));
    view.appendChild(h('div', { class: 'cols' }, datosTablero(t, r, R)));
    view.appendChild(circuitos(t, r, R));
    view.appendChild(h('div', { class: 'cols' }, demanda(t, r)));
    view.appendChild(alimentador(t, r));
    view.appendChild(alimentacion(t, r));
    const rn = revisionNec(r); if (rn) view.appendChild(rn);
  };
  const n100 = v => (Number(v) || 0) * 100;
})();
