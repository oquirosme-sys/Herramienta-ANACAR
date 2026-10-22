/* Pestaña Proyecto: datos generales, criterios, alta de tableros (cuántos y de cuál se alimenta cada uno) y diagrama de alimentación. */
(function () {
  'use strict';
  const h = U.h, f2 = v => U.fmt(v, 2);

  function crearVarios() {
    const cant = h('input', { type: 'number', min: 1, max: 60, value: 3 }), pref = h('input', { type: 'text', value: 'T' });
    const tipo = UI.select(['3F', '1F'], '3F', () => {}), sis = UI.select(Store.catalog.listas.sistemas, '120/208', () => {});
    const padre = UI.select(UI.opts(Store.project.tableros.map(t => ({ value: t.id, label: Calc.nombreTablero(t) })), '— Ninguno (acometida / externo) —'), '', () => {});
    const nombres = h('textarea', { rows: 3, placeholder: 'Opcional: un nombre por línea (TA, TB, TC… o 2G1, 2G2…). Si lo deja vacío se usa el prefijo + letra.' });
    UI.modal('Crear tableros', h('div', { class: 'grid2' },
      UI.field('Cantidad', cant), UI.field('Prefijo del nombre', pref, 'TA, TB, TC…'), UI.field('Tipo', tipo), UI.field('Sistema (V)', sis),
      UI.field('Alimentados desde', padre, 'Puede cambiarlo luego para cada tablero.', 'span2'), UI.field('Nombres', nombres, null, 'span2')), [
      { label: 'Cancelar' },
      { label: 'Crear', cls: 'primary', onclick: () => {
        const lista = nombres.value.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
        const n = lista.length || Math.max(1, Math.min(60, Number(cant.value) || 1));
        for (let i = 0; i < n; i++) {
          const t = Store.nuevoTablero({ nombre: lista[i] || (pref.value.trim() + String.fromCharCode(65 + (i % 26)) + (i >= 26 ? Math.floor(i / 26) : '')), tipo: tipo.value, sistema: sis.value });
          if (padre.value) Store.cambiarPadre(t, padre.value);
        }
        App.refresh(); UI.toast(n + ' tablero(s) creados', 'ok');
      } },
    ]);
  }

  /** Árbol de alimentación (unifilar simplificado). */
  /** Transformadores y generadores del proyecto: un mismo equipo puede alimentar varios tableros y ATS; el voltaje de esos tableros es el secundario nominal. */
  function tarjetaFuentes(R) {
    const p = Store.project, L = Store.catalog.listas;
    const info = id => (R.fuentes || []).find(x => x.fuente.id === id) || { salidas: 0, carga: 0, cargaPct: null, tableros: [] };
    const filas = p.fuentes.map(fu => {
      const k = 'fu:' + fu.id + ':', gen = fu.tipo === 'generador', i = info(fu.id);
      const sis = UI.select(L.sistemas, fu.sistema, v => { fu.sistema = v; Store.sincronizarFuentes(); Store.save(); App.refresh(); }, { fk: k + 'sis', label: 'Secundario' });
      const lista = gen ? null : UI.select(UI.opts(Store.catalog.transformadores.map(x => ({ value: x.id, label: x.nombre + ' · Z ' + x.z + ' %', group: x.tipo })), '— elegir —'), '', v => {
        const x = Store.catalog.transformadores.find(o => String(o.id) === v); if (!x) return;
        const m = /(\d+(?:[.,]\d+)?)\s*KVA/i.exec(x.nombre); if (m) fu.kva = Number(m[1].replace(',', '.'));
        fu.z = x.z; fu.fases = /MONO/i.test(x.tipo) ? 1 : 3; Store.save(); App.refresh();
      }, { fk: k + 'lis', label: 'De la lista del Excel' });
      const nombres = i.tableros.map(id => Store.tablero(id)).filter(Boolean).map(Calc.nombreTablero).join(', ');
      return h('tr', null,
        h('td', null, UI.input(fu, 'nombre', { fk: k + 'nom', class: 'w-nm', label: 'Nombre' })), h('td', null, gen ? 'Generador' : 'Transformador'),
        h('td', null, UI.input(fu, 'kva', { type: 'num', fk: k + 'kva', class: 'w-num', label: 'kVA' })), h('td', null, sis),
        h('td', null, gen ? h('span', { class: 'muted' }, '—') : UI.input(fu, 'z', { type: 'num', fk: k + 'z', class: 'w-num', label: 'Z %' })),
        h('td', null, gen ? UI.input(fu, 'xd', { type: 'num', fk: k + 'xd', class: 'w-num', label: "X''d %", placeholder: '12' }) : UI.input(fu, 'xr', { type: 'num', fk: k + 'xr', class: 'w-num', label: 'X/R', placeholder: '4' })),
        h('td', null, gen ? h('span', { class: 'muted' }, '—') : UI.input(fu, 'primario', { type: 'num', fk: k + 'pri', class: 'w-num', label: 'Primario (V)' })),
        h('td', null, gen ? h('span', { class: 'muted' }, '—') : UI.bind(fu, 'fases', [{ value: 3, label: '3F' }, { value: 1, label: '1F' }], { num: true, fk: k + 'fas', label: 'Fases' })),
        h('td', null, gen ? h('span', { class: 'muted' }, '—') : h('div', { class: 'col' },
          h('label', { class: 'chk', title: 'Interruptor en el primario del transformador' }, h('input', { type: 'checkbox', checked: fu.protPrim !== false, onchange: e => { fu.protPrim = e.target.checked; Store.save(); App.refresh(); } }), ' Primario'),
          h('label', { class: 'chk', title: 'Interruptor en el secundario del transformador' }, h('input', { type: 'checkbox', checked: fu.protSec === true, onchange: e => { fu.protSec = e.target.checked; Store.save(); App.refresh(); } }), ' Secundario'))),
        h('td', null, lista || h('span', { class: 'muted' }, '—')),
        h('td', null, i.salidas ? h('span', { title: nombres }, i.salidas + ' · ' + f2(i.carga) + ' kVA' + (i.cargaPct !== null ? ' (' + f2(i.cargaPct * 100) + ' %)' : '')) : h('span', { class: 'muted' }, 'Sin tableros')),
        h('td', { class: 'acc' }, UI.iconBtn('🗑', 'Eliminar', async () => { if (await UI.confirm('¿Eliminar ' + (fu.nombre || 'la fuente') + '? Los tableros que cuelgan de ella quedan sin fuente.', 'Eliminar')) { Store.eliminarFuente(fu.id); App.refresh(); } }, 'danger')));
    });
    return UI.card('Fuentes: transformadores y generadores (' + p.fuentes.length + ')',
      p.fuentes.length ? h('div', { class: 'tbl-wrap' }, h('table', { class: 'tbl' }, h('thead', null, h('tr', null, ['Nombre', 'Tipo', 'kVA', 'Secundario (V)', 'Z (%)', "X/R · X''d (%)", 'Primario (V)', 'Fases', 'Protección', 'De la lista del Excel', 'Tableros · carga', ''].map(x => h('th', null, x)))), h('tbody', null, filas)))
        : h('p', { class: 'muted' }, 'Opcional. Cree aquí el transformador de pedestal o el generador y luego, en "Alimentado desde" de cada tablero (o en la segunda acometida de un ATS), elija esa fuente: varios tableros comparten el mismo equipo y toman su voltaje secundario.'),
      h('div', { class: 'toolbar' }, UI.btn('+ Transformador', () => { Store.nuevaFuente('transformador'); App.refresh(); }, 'small'), UI.btn('+ Generador', () => { Store.nuevaFuente('generador'); App.refresh(); }, 'small')));
  }

  function arbol(R) {
    const raiz = R.orden.filter(r => r.nivel === 0);
    const nodo = r => h('li', null,
      h('a', { class: 'node' + (r.avisos.length ? ' warn' : ''), href: '#memoria/' + r.tab.id, title: r.avisos.join('\n') || 'Abrir memoria de cálculo' },
        h('b', null, r.nombre), h('span', { class: 'tag' }, r.tab.tipo + ' · ' + r.tab.sistema + ' V'),
        h('small', null, f2(r.W130) + ' kVA · ' + (r.alim.AF139 || '–') + ' A'),
        h('small', null, 'Bornes ' + f2(r.alim.AW139) + ' V · ΔV ' + f2(r.alim.AY139) + ' %'),
        h('small', null, 'Icc ' + (r.iccKA ? f2(r.iccKA) + ' kA' : '—'))),
      (R.hijos[r.tab.id] || []).length ? h('ul', null, R.hijos[r.tab.id].map(t => R.res[t.id]).filter(Boolean).map(nodo)) : null);
    return h('div', { class: 'tree', 'data-scroll': 'arbol' }, raiz.length ? h('ul', null, raiz.map(r => h('li', { class: 'root' }, h('div', { class: 'source' }, r.tab.conectadoA || 'Acometida'), h('ul', null, nodo(r))))) : h('p', { class: 'muted' }, 'Sin tableros.'));
  }

  App.views.proyecto = function (view, R) {
    const p = Store.project, L = Store.catalog.listas;
    // al abrir la herramienta el proyecto está en blanco: se ofrecen los guardados
    if (!p.nombre && !p.numero && !p.tableros.length && Store.listaProyectos().length) {
      view.appendChild(UI.card('Proyectos guardados en este navegador', App.proyectosGuardados(true), h('small', { class: 'muted' }, 'Abra uno para continuar o empiece este proyecto en blanco')));
    }
    view.appendChild(h('div', { class: 'page-h' }, h('h2', null, 'Proyecto'), h('p', { class: 'muted' }, 'Indique los datos del proyecto, cuántos tableros tiene y de cuál se alimenta cada uno. Las caídas de voltaje y el cortocircuito se calculan en cascada desde la acometida.')));

    view.appendChild(h('div', { class: 'cols' },
      UI.card('Datos del proyecto', h('div', { class: 'grid3' },
        UI.field('Nombre', UI.input(p, 'nombre', { fk: 'p:nombre' }), null, 'span2'), UI.field('Proyecto N.º', UI.input(p, 'numero', { fk: 'p:numero' })),
        UI.field('Ubicación', UI.input(p, 'ubicacion', { fk: 'p:ubic' }), null, 'span2'), UI.field('Fecha', UI.input(p, 'fecha', { type: 'date', fk: 'p:fecha' })),
        UI.field('Elaboró', UI.input(p, 'elaboro', { fk: 'p:elab' }), null, 'span2'), UI.field('Revisión', UI.input(p, 'revision', { fk: 'p:rev' })),
        UI.field('Formato del cuadro de cargas', UI.bind(p, 'pais', [{ value: 'CR', label: 'Formato original (Costa Rica)' }, { value: 'PA', label: 'Panamá (formato TL4)' }], { fk: 'p:pais' }), 'Panamá dibuja el cuadro 3F/1F con dispositivos por circuito (apagadores, dimmers, tomas, extractores); el formato original no cambia', 'span3'))),
      UI.card('Criterios de diseño', h('div', { class: 'grid3' },
        UI.field('ΔV máx. alimentador (%)', UI.input(p, 'cvMaxAlim', { type: 'num', fk: 'p:cva' }), 'Acumulada hasta los bornes del tablero'),
        UI.field('ΔV máx. total (%)', UI.input(p, 'cvMaxTotal', { type: 'num', fk: 'p:cvt' }), 'Alimentadores + circuito ramal'),
        UI.field('Desbalance máx. (%)', UI.input(p, 'desbalanceMax', { type: 'num', fk: 'p:des' })),
        UI.field('Marca por defecto', UI.bind(p, 'marcaDefecto', Store.catalog.marcas, { fk: 'p:marca' }), 'Para elegir tablero, breakers y supresor automáticamente'),
        UI.field('Long. máx. para Icc (m)', UI.input(p, 'iccLongMax', { type: 'num', fk: 'p:icl' }), 'Igual que el Excel (20 m). 0 = sin límite'),
        UI.field('Ocupación (NEC)', UI.bind(p, 'ocupacion', [{ value: 'comercial', label: 'Comercial / oficinas' }, { value: 'vivienda', label: 'Vivienda' }, { value: 'hotel', label: 'Hotel' }, { value: 'hospital', label: 'Hospital' }, { value: 'industrial', label: 'Industrial' }], { fk: 'p:ocu' }), 'Para GFCI/AFCI (210.8, 210.12) y SPD (230.67)'),
        UI.field('Carga de tableros derivados', UI.bind(p, 'cargaDerivados', [{ value: 'conectada', label: 'kVA conectados (como el Excel)' }, { value: 'demandada', label: 'kVA demandados (no repite demanda)' }], { fk: 'p:cder' }), 'Con qué carga sube cada tablero al que lo alimenta'),
        UI.field('Reserva de espacios (%)', UI.input(p, 'reservaEspacios', { type: 'num', fk: 'p:res' }), 'Para elegir el tablero de catálogo'),
        UI.field('Icc de la red (kA)', UI.input(p, 'iccRed', { type: 'num', fk: 'p:iccr', placeholder: 'infinita' }), 'En el primario de los transformadores principales'),
        UI.field('Alimentadores', h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: p.autoAmpacidad !== false, onchange: e => { p.autoAmpacidad = e.target.checked; Store.save(); App.refresh(); } }), ' Aumentar calibre si no cumple 310.15'), 'Temperatura y agrupamiento'),
      ))));

    view.appendChild(tarjetaFuentes(R));

    // ---- tableros
    const cuerpo = h('tbody');
    R.orden.concat(p.tableros.filter(t => !R.res[t.id]).map(t => ({ tab: t, nivel: 0, avisos: ['No calculado'], alim: {} }))).forEach(r => {
      const t = r.tab, k = 't:' + t.id + ':';
      const padres = Store.posiblesPadres(t);
      const selPadre = UI.select(UI.opts(Store.opcionesAlimentacion(t), '— Acometida / externo —'), Store.valorAlimentacion(t), v => { Store.alimentarDesde(t, v); App.refresh(); }, { fk: k + 'padre', label: 'Alimentado desde', class: 'w-padre' });
      cuerpo.appendChild(h('tr', null,
        h('td', { class: 'nm', style: { paddingLeft: (0.4 + r.nivel * 1.1) + 'rem' } }, r.nivel ? h('span', { class: 'muted' }, '└ ') : null, UI.input(t, 'nombre', { fk: k + 'nombre', class: 'w-nm', label: 'Nombre del tablero' })),
        h('td', null, UI.select(['3F', '1F'], t.tipo, v => { t.tipo = v; Object.assign(t, Store.voltajeDe(t.sistema, v)); Store.save(); App.refresh(); }, { fk: k + 'tipo', label: 'Tipo' })),
        h('td', null, UI.select(L.sistemas, t.sistema, v => { t.sistema = v; Object.assign(t, Store.voltajeDe(v, t.tipo)); Store.save(); App.refresh(); }, { fk: k + 'sis', label: 'Sistema', disabled: !!t.fuenteId, title: t.fuenteId ? 'Heredado del secundario de la fuente' : null })),
        h('td', null, selPadre, !t.padreId ? UI.input(t, 'conectadoA', { fk: k + 'con', placeholder: 'Conectado a (p. ej. MÓDULO MEDIDORES)', class: 'mt' }) : null),
        h('td', null, UI.input(t, 'longitud', { type: 'num', fk: k + 'long', class: 'w-num', label: 'Longitud del alimentador (m)' })),
        h('td', { class: 'r' }, f2(r.W130)), h('td', { class: 'r' }, r.alim.AF139 || ''),
        h('td', { class: 'r' }, f2(r.alim.AW139)), h('td', { class: 'r ' + (r.alim.AY139 > p.cvMaxAlim ? 'bad' : '') }, f2(r.alim.AY139)),
        h('td', { class: 'r' }, r.iccKA ? f2(r.iccKA) : h('span', { class: 'muted' }, '—'), r.iccFuente ? h('small', { class: 'muted d' }, r.iccFuente) : null),
        h('td', null, r.avisos.length ? h('span', { class: 'badge warn', title: r.avisos.join('\n') }, r.avisos.length + ' aviso(s)') : h('span', { class: 'badge ok' }, 'OK')),
        h('td', { class: 'acc' },
          UI.iconBtn('✎', 'Memoria de cálculo', () => App.go('memoria', t.id)),
          UI.iconBtn('⧉', 'Duplicar', () => { Store.duplicarTablero(t.id); App.refresh(); }),
          UI.iconBtn('🗑', 'Eliminar', async () => { if (await UI.confirm('¿Eliminar ' + Calc.nombreTablero(t) + ' y sus ' + t.circuitos.length + ' circuitos?', 'Eliminar')) { Store.eliminarTablero(t.id); App.refresh(); } }, 'danger'))));
    });
    view.appendChild(UI.card('Tableros (' + p.tableros.length + ')',
      p.tableros.length ? h('div', { class: 'tbl-wrap', 'data-scroll': 'tabs' }, h('table', { class: 'tbl' },
        h('thead', null, h('tr', null, ['Tablero', 'Tipo', 'Sistema (V)', 'Alimentado desde', 'Long. alim. (m)', 'kVA totales', 'Breaker (A)', 'V bornes', 'ΔV acum. (%)', 'Icc (kA)', 'Estado', ''].map(x => h('th', null, x)))), cuerpo))
        : h('div', { class: 'empty' }, 'Aún no hay tableros. Créelos aquí o impórtelos desde la tabla de circuitos de Revit.'),
      h('div', { class: 'toolbar' },
        UI.btn('+ Tablero', () => { const t = Store.nuevoTablero({ nombre: 'T' + String.fromCharCode(65 + (p.tableros.length % 26)) }); App.refresh(); setTimeout(() => { const e = document.querySelector('[data-fk="t:' + t.id + ':nombre"]'); if (e) e.select(); }, 0); }, 'primary small'),
        UI.btn('Crear varios…', crearVarios, 'small'),
        UI.btn('Importar de Revit…', () => App.go('revit'), 'small'))));

    view.appendChild(UI.card('Diagrama de alimentación', arbol(R), h('a', { class: 'btn small', href: '#unifilar' }, 'Ver diagrama unifilar con transformadores y cortocircuito →')));
  };
})();
