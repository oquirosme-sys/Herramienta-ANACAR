/* Importar la tabla de circuitos de Revit (Electrical Circuit Schedule): .txt/.csv exportado, el Excel "Circuitos revit-excel" o filas pegadas.
   Agrupa por tablero (Panel), propone el detalle de carga de cada circuito y crea o actualiza los tableros. */
(function () {
  'use strict';
  const h = U.h;
  let sesion = null; // { paneles: [{ panel, destino, tipo, sistema, modo, incluir, circuitos: [...] }], unidad }

  /* ---------- lectura ---------- */
  const COLS = [
    ['voltaje', /volt/i], ['nombre', /load name|nombre|descrip/i], ['carga', /true load|apparent load|carga|load|kva|potencia/i], ['panel', /panel|tablero/i],
    ['circuito', /circuit|circuito/i], ['polos', /pole|polo/i], ['longitud', /length|longitud|largo/i],
  ];
  function mapear(filas) {
    for (let i = 0; i < Math.min(filas.length, 30); i++) {
      const f = filas[i].map(x => String(x === undefined || x === null ? '' : x));
      const m = {};
      f.forEach((txt, j) => { const c = COLS.find(([k, re]) => m[k] === undefined && re.test(txt) && !(k === 'carga' && /name|nombre/i.test(txt))); if (c) m[c[0]] = j; });
      if (m.panel !== undefined && m.circuito !== undefined) return { inicio: i + 1, m };
    }
    return { inicio: 0, m: { voltaje: 0, carga: 1, panel: 2, circuito: 3, polos: 4, nombre: 5, longitud: 6 } }; // orden de la macro ImportarCircuitosRevit
  }
  function numero(v) {
    if (typeof v === 'number') return v;
    let s = String(v || '').trim(); if (!s) return null;
    const pies = /(\d+(?:[.,]\d+)?)\s*'\s*(?:-?\s*(\d+(?:[.,]\d+)?)\s*")?/.exec(s);
    if (pies) return { pies: Number(pies[1].replace(',', '.')) + (pies[2] ? Number(pies[2].replace(',', '.')) / 12 : 0) };
    const m = /-?\d[\d.,]*/.exec(s); if (!m) return null;
    let t = m[0];
    if (t.includes(',') && t.includes('.')) t = t.lastIndexOf(',') > t.lastIndexOf('.') ? t.replace(/\./g, '').replace(',', '.') : t.replace(/,/g, '');
    else if (t.includes(',')) t = t.replace(',', '.');
    let n = Number(t);
    if (/\bva\b/i.test(s) && !/kva/i.test(s)) n = n / 1000; // VA → kVA
    return isFinite(n) ? n : null;
  }
  /** Repara textos UTF-8 leídos como ANSI por la macro de Excel ("HabitaciÃ³n" → "Habitación"). */
  const CP1252 = (() => { const m = {}, d = new TextDecoder('windows-1252'); for (let b = 0; b < 256; b++) m[d.decode(new Uint8Array([b]))] = b; return m; })();
  function reparar(s) {
    s = String(s === undefined || s === null ? '' : s);
    if (!/[ÃÂ][\u0080-ÿŒ-™]/.test(s)) return s;
    try { return new TextDecoder('utf-8', { fatal: true }).decode(new Uint8Array(Array.from(s, ch => (CP1252[ch] === undefined ? 63 : CP1252[ch])))); } catch (e) { return s; }
  }
  function polosDe(v, n) {
    let arr;
    if (typeof v === 'number') arr = [Math.trunc(v)]; // "20,22" llega como 20.22 (o 18,20 como 18.2): se toma el primero y se completan los demás
    else arr = String(v || '').split(/[^0-9]+/).map(Number);
    arr = arr.filter(x => x > 0);
    n = Number(n) || arr.length || 1;
    if (arr.length && arr.length < n) for (let i = arr.length; i < n; i++) arr.push(arr[0] + 2 * i);
    return arr.slice(0, Math.max(n, 1));
  }
  function procesar(filas, origen) {
    filas = filas.filter(f => f && f.some(x => x !== '' && x !== null && x !== undefined));
    const { inicio, m } = mapear(filas), grupos = {};
    filas.slice(inicio).forEach(f => {
      const panel = reparar(f[m.panel]).trim();
      if (!panel || /electrical|schedule|^panel$/i.test(panel)) return;
      const nPolos = Number(numero(f[m.polos])) || 1, v = Number(numero(f[m.voltaje])) || 0;
      const lon = numero(f[m.longitud]);
      const c = { nombre: reparar(f[m.nombre]).trim(), voltaje: v, polos: polosDe(f[m.circuito], nPolos), nPolos, kva: Number(numero(f[m.carga])) || 0,
        longitud: lon && lon.pies !== undefined ? lon.pies : lon, enPies: !!(lon && lon.pies !== undefined), circuitoTxt: String(f[m.circuito]) };
      const d = Calc.detalleRevit(c.nombre, v, nPolos === 2 && v >= 208 ? 2 : nPolos, Store.catalog);
      c.detalleId = d.detalle ? d.detalle.id : ''; c.seguro = d.seguro;
      (grupos[panel] = grupos[panel] || []).push(c);
    });
    const paneles = Object.keys(grupos).sort().map(panel => {
      const cs = grupos[panel], existente = Store.project.tableros.find(t => norm(t.nombre) === norm(panel) || norm(Calc.nombreTablero(t)) === norm(panel));
      const tres = cs.some(c => c.nPolos === 3), vmax = Math.max(...cs.map(c => c.voltaje));
      const sistema = vmax >= 440 ? '277/480' : vmax === 240 ? '120/240' : '120/208';
      return { panel, destino: existente ? existente.id : 'nuevo', tipo: tres || sistema === '120/208' ? '3F' : '1F', sistema, modo: 'reemplazar', incluir: true, circuitos: cs };
    });
    if (!paneles.length) throw new Error('No se encontraron circuitos. Verifique que la tabla tenga las columnas Panel y Circuit Number.');
    sesion = { paneles, origen, unidad: paneles.some(p => p.circuitos.some(c => c.enPies)) ? 'pies-detectado' : 'm' };
    App.refresh();
  }
  const norm = s => String(s || '').toUpperCase().replace(/^TABLERO\s+/, '').replace(/\s+/g, '');

  function textoAFilas(txt) {
    const lineas = txt.split(/\r?\n/).filter(l => l.trim());
    const sep = lineas.some(l => l.includes('\t')) ? '\t' : (lineas.filter(l => l.includes(';')).length > lineas.length / 2 ? ';' : ',');
    return lineas.map(l => {
      if (sep !== ',') return l.split(sep).map(x => x.replace(/^"|"$/g, '').trim());
      const out = []; let cur = '', q = false;
      for (const ch of l) { if (ch === '"') q = !q; else if (ch === ',' && !q) { out.push(cur.trim()); cur = ''; } else cur += ch; }
      out.push(cur.trim()); return out;
    });
  }
  async function abrir() {
    const f = await UI.pickFile('.txt,.csv,.tsv,.xlsx,.xlsm'); if (!f) return;
    try {
      if (/\.xls[xm]$/i.test(f.name)) {
        const hojas = await XlsxRead.read(await UI.readBuffer(f));
        const hoja = hojas.find(s => s.state === 'visible' && s.rows.some(r => r.some(x => /panel/i.test(String(x))))) || hojas[0];
        procesar(hoja.rows, f.name + ' · ' + hoja.name);
      } else procesar(textoAFilas(await UI.readText(f)), f.name);
    } catch (e) { UI.alert(e.message, 'No se pudo importar'); }
  }

  /* ---------- aplicar ---------- */
  function importar() {
    let nT = 0, nC = 0;
    const factor = sesion.unidad === 'pies' || sesion.unidad === 'pies-detectado' ? 0.3048 : 1;
    sesion.paneles.filter(p => p.incluir).forEach(p => {
      let t = Store.tablero(p.destino);
      if (!t) { t = Store.nuevoTablero({ nombre: p.panel, tipo: p.tipo, sistema: p.sistema }); nT++; }
      if (p.modo === 'reemplazar') t.circuitos = t.circuitos.filter(c => c.tableroHijoId);
      p.circuitos.forEach(c => {
        const libres = c.polos.filter(x => !t.circuitos.some(o => (o.polos || []).includes(x)));
        t.circuitos.push({ id: U.uid(), polos: libres.length === c.polos.length ? c.polos : Store.posicionLibre(t, c.polos.length), detalleId: c.detalleId, descripcion: c.nombre, kva: c.kva,
          longitud: c.longitud === null ? '' : Math.round(c.longitud * factor * 100) / 100, material: 'CU', aislamiento: 'THHN', mult: '', paralelos: '', aumento: 1, breakerId: '', prot: '', tableroHijoId: '', origen: 'revit', circuitoRevit: c.polos.join(',') });
        nC++;
      });
    });
    Store.save(); sesion = null;
    UI.toast(nC + ' circuitos importados' + (nT ? ', ' + nT + ' tablero(s) nuevos' : ''), 'ok');
    App.go('proyecto');
  }

  /* ---------- cambios de circuito pendientes para pasar a Revit ---------- */
  let verAplicados = false;
  function filasCambios() {
    const out = [];
    Store.project.tableros.forEach(t => (t.cambiosRevit || []).forEach(e => { if (verAplicados || !e.aplicado) out.push({ t, e }); }));
    return out;
  }
  function autobalancearTodos() {
    const R = App.R, res = [];
    R.orden.forEach(r => {
      if (r.fases === 1) return;
      const b = Calc.balanceo(r.tab, Store.catalog, { espacios: r.cat ? r.cat.espacios : 0, maxMov: 40, objetivo: Math.min(5, Store.project.desbalanceMax || 5) });
      if (b.cambios.length) res.push({ r, b });
    });
    if (!res.length) { UI.alert('No hay cambios que mejoren el balance de los tableros.', 'Autobalanceo'); return; }
    UI.modal('Autobalancear todos los tableros', h('div', null,
      h('table', { class: 'tbl' }, h('thead', null, h('tr', null, ['Tablero', 'Desbalance actual', 'Con cambios', 'Circuitos a mover'].map(x => h('th', null, x)))),
        h('tbody', null, res.map(({ r, b }) => h('tr', null, h('td', null, r.nombre), h('td', { class: 'r' }, U.fmt(b.antes, 2) + ' %'), h('td', { class: 'r okc' }, U.fmt(b.despues, 2) + ' %'), h('td', { class: 'r' }, b.cambios.length))))),
      h('p', { class: 'muted' }, 'Los circuitos fijados no se mueven. Los cambios quedan anotados en esta lista para pasarlos a Revit.')), [
      { label: 'Cancelar' },
      { label: 'Aplicar en todos', cls: 'primary', onclick: () => { res.forEach(({ r, b }) => Store.registrarCambios(r.tab, b.cambios, 'balanceo')); App.refresh(); UI.toast('Cambios aplicados en ' + res.length + ' tablero(s)', 'ok'); } },
    ], { wide: true });
  }
  function vistaCambios(view) {
    const filas = filasCambios(), pend = filas.filter(x => !x.e.aplicado);
    const tabla = () => [['Panel (Revit)', 'Carga', 'Circuito en Revit', 'Circuito nuevo', 'Origen', 'Fecha']].concat(filas.map(({ t, e }) => [t.nombre, e.descripcion, e.revit || e.de, e.a, e.origen, e.fecha]));
    view.appendChild(h('div', { class: 'toolbar' },
      UI.btn('⚖ Autobalancear todos los tableros', autobalancearTodos, 'primary small'),
      UI.btn('Copiar tabla', () => { navigator.clipboard.writeText(tabla().map(r => r.join('\t')).join('\n')).then(() => UI.toast('Tabla copiada: péguela en Excel', 'ok'), () => UI.alert('No se pudo copiar.')); }, 'small'),
      UI.btn('Exportar CSV', () => U.download(App.fileBase() + ' - cambios para Revit.csv', U.csv(tabla()), 'text/csv'), 'small'),
      pend.length ? UI.btn('Marcar todos como aplicados', async () => { if (await UI.confirm('¿Marcar los ' + pend.length + ' cambios como ya aplicados en Revit?', 'Marcar', false)) { pend.forEach(x => { x.e.aplicado = true; }); Store.save(); App.refresh(); } }, 'small') : null,
      h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: verAplicados, onchange: e => { verAplicados = e.target.checked; App.refresh(); } }), ' Ver también los aplicados')));
    if (!filas.length) { view.appendChild(h('div', { class: 'empty' }, 'No hay cambios pendientes. Al autobalancear o cambiar la posición de un circuito, el cambio aparece aquí para pasarlo al modelo de Revit.')); return; }
    view.appendChild(UI.card('Cambios de circuito para Revit (' + pend.length + ' pendientes)', h('div', null, h('div', { class: 'tbl-wrap' }, h('table', { class: 'tbl' },
      h('thead', null, h('tr', null, ['Aplicado', 'Panel (Revit)', 'Carga', 'Circuito en Revit', '→', 'Circuito nuevo', 'Origen', 'Fecha', ''].map(x => h('th', null, x)))),
      h('tbody', null, filas.map(({ t, e }) => h('tr', { class: e.aplicado ? 'dim' : '' },
        h('td', { class: 'c' }, h('input', { type: 'checkbox', checked: !!e.aplicado, onchange: ev => { e.aplicado = ev.target.checked; Store.save(); App.refresh(); } })),
        h('td', null, h('a', { href: '#memoria/' + t.id }, t.nombre)), h('td', null, e.descripcion), h('td', null, e.revit || e.de), h('td', { class: 'muted' }, '→'), h('td', null, h('b', null, e.a)),
        h('td', null, e.origen === 'balanceo' ? 'Autobalanceo' : 'Manual'), h('td', null, e.fecha),
        h('td', { class: 'acc' }, UI.iconBtn('↶', 'Deshacer este cambio (vuelve a la posición anterior)', () => {
          const c = t.circuitos.find(x => x.id === e.circuitoId);
          if (c) c.polos = String(e.de).split(',').map(Number).filter(Boolean);
          t.cambiosRevit.splice(t.cambiosRevit.indexOf(e), 1); Store.save(); App.refresh();
        }))))))),
      h('p', { class: 'hint' }, 'En Revit: abra el tablero (Panel), seleccione el circuito y use "Mover a" o edite el número de circuito. Al terminar márquelo como aplicado.'))));
  }

  /* ---------- vista ---------- */
  App.views.revit = function (view, R, sub) {
    view.appendChild(h('nav', { class: 'chips no-print' }, h('a', { class: 'chip' + (sub !== 'cambios' ? ' active' : ''), href: '#revit' }, 'Importar circuitos'),
      h('a', { class: 'chip' + (sub === 'cambios' ? ' active' : ''), href: '#revit/cambios' }, 'Cambios para Revit (' + Store.project.tableros.reduce((a, t) => a + (t.cambiosRevit || []).filter(e => !e.aplicado).length, 0) + ')')));
    if (sub === 'cambios') { view.appendChild(h('div', { class: 'page-h' }, h('h2', null, 'Cambios para Revit'), h('p', { class: 'muted' }, 'Lista de circuitos que cambiaron de posición (autobalanceo o edición) para actualizar el modelo de Revit.'))); vistaCambios(view); return; }
    view.appendChild(h('div', { class: 'page-h' }, h('h2', null, 'Importar circuitos de Revit'), h('p', { class: 'muted' },
      'Los tableros se alimentan con la tabla de circuitos de Revit (Electrical Circuit Schedule) con las columnas Voltage, True Load, Panel, Circuit Number, Number of Poles, Load Name y Length.')));
    if (!sesion) {
      const area = h('textarea', { rows: 8, placeholder: 'Pegue aquí las filas copiadas de Revit o de Excel (incluya la fila de encabezados)…' });
      view.appendChild(h('div', { class: 'cols' },
        UI.card('1. Abrir archivo', h('div', null,
          h('p', null, 'Acepta el .txt/.csv que exporta Revit (Exportar ▸ Informes ▸ Tabla de planificación) o el Excel ', h('i', null, 'Circuitos revit-excel.xlsm'), '.'),
          UI.btn('Elegir archivo…', abrir, 'primary'))),
        UI.card('… o pegar filas', h('div', null, area, h('div', { class: 'toolbar' }, UI.btn('Leer filas pegadas', () => { try { procesar(textoAFilas(area.value), 'Pegado'); } catch (e) { UI.alert(e.message); } }))))));
      view.appendChild(UI.card('Cómo se interpreta', h('ul', { class: 'help' },
        h('li', null, 'Cada valor de ', h('b', null, 'Panel'), ' es un tablero. Si ya existe uno con ese nombre se actualiza; si no, se crea.'),
        h('li', null, h('b', null, 'Circuit Number'), ' da la posición: "1,3,5" o "20,22". Si solo viene el primero se completan los polos de 2 en 2.'),
        h('li', null, 'El ', h('b', null, 'detalle de carga'), ' (tipo, voltaje, fases, factores) se propone por el nombre, el voltaje y los polos; revise los marcados con ⚠.'),
        h('li', null, h('b', null, 'True Load'), ' se toma en kVA (si viene en VA se convierte). ', h('b', null, 'Length'), ' en metros; las longitudes en pies (12\' - 6") se convierten.'))));
      return;
    }
    const dOpts = Store.catalog.detallesCarga.map(d => ({ value: d.id, label: d.id + ' · ' + d.descripcion + ' (' + d.v + ' V ' + d.fases + 'F)' }));
    const destinos = [{ value: 'nuevo', label: '+ Crear tablero nuevo' }].concat(Store.project.tableros.map(t => ({ value: t.id, label: Calc.nombreTablero(t) })));
    const total = sesion.paneles.reduce((a, p) => a + p.circuitos.length, 0), dudosos = sesion.paneles.reduce((a, p) => a + p.circuitos.filter(c => !c.seguro).length, 0);
    view.appendChild(h('div', { class: 'toolbar sticky-bar' },
      h('span', null, h('b', null, total + ' circuitos'), ' en ', h('b', null, sesion.paneles.length + ' tableros'), ' · ' + sesion.origen + (dudosos ? ' · ' : ''), dudosos ? h('span', { class: 'warnc' }, dudosos + ' por revisar ⚠') : null),
      h('span', { class: 'grow' }),
      UI.field('Longitudes en', UI.select([{ value: 'm', label: 'metros' }, { value: 'pies', label: 'pies' }], sesion.unidad === 'pies-detectado' ? 'pies' : sesion.unidad, v => { sesion.unidad = v; }), null, 'inline'),
      UI.btn('Cancelar', () => { sesion = null; App.refresh(); }), UI.btn('Importar', importar, 'primary')));
    sesion.paneles.forEach((p, pi) => {
      const head = h('div', { class: 'grid4' },
        UI.field('Incluir', h('input', { type: 'checkbox', checked: p.incluir, onchange: e => { p.incluir = e.target.checked; } })),
        UI.field('Destino', UI.select(destinos, p.destino, v => { p.destino = v; App.refresh(); })),
        p.destino === 'nuevo' ? UI.field('Tipo / sistema', h('div', { class: 'row' }, UI.select(['3F', '1F'], p.tipo, v => { p.tipo = v; }), UI.select(Store.catalog.listas.sistemas, p.sistema, v => { p.sistema = v; }))) :
          UI.field('Circuitos existentes', UI.select([{ value: 'reemplazar', label: 'Reemplazar (conserva los de tableros derivados)' }, { value: 'agregar', label: 'Agregar' }], p.modo, v => { p.modo = v; })));
      const tbl = h('table', { class: 'tbl' }, h('thead', null, h('tr', null, ['Circuito', 'Posición', 'Nombre de carga (Revit)', 'V', 'Polos', 'kVA', 'Long.', 'Detalle de carga'].map(x => h('th', null, x)))),
        h('tbody', null, p.circuitos.map((c, ci) => h('tr', { class: c.seguro ? '' : 'row-warn' }, h('td', null, c.circuitoTxt), h('td', null, c.polos.join(',')), h('td', null, c.nombre),
          h('td', { class: 'r' }, c.voltaje), h('td', { class: 'r' }, c.nPolos), h('td', { class: 'r' }, U.fmt(c.kva, 2)), h('td', { class: 'r' }, c.longitud === null ? '' : U.fmt(c.longitud, 2)),
          h('td', null, c.seguro ? null : h('span', { title: 'Revisar: detalle propuesto por coincidencia aproximada' }, '⚠ '), UI.select(dOpts, c.detalleId, v => { c.detalleId = Number(v); c.seguro = true; }, { class: 'w-det', fk: 'rv:' + pi + ':' + ci }))))));
      view.appendChild(UI.card('Panel ' + p.panel + ' (' + p.circuitos.length + ' circuitos · ' + U.fmt(p.circuitos.reduce((a, c) => a + c.kva, 0), 2) + ' kVA)', h('div', null, head, h('div', { class: 'tbl-wrap', 'data-scroll': 'rv' + pi }, tbl))));
    });
  };
})();
