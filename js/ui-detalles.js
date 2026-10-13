/* Detalles de carga (equivale a la hoja DCARGAS del ANACAR original): cada detalle define voltaje, fases, hilos,
   tipo de carga (demanda), factores, factor de potencia, si es continua y el multiplicador de protección.
   Las "palabras clave (Revit)" permiten que la importación de Revit reconozca el detalle por el nombre de la carga. */
(function () {
  'use strict';
  const h = U.h;
  let sec = 'detalles';

  const usos = id => {
    const l = [];
    Store.project.tableros.forEach(t => (t.circuitos || []).forEach(c => { if (String(c.detalleId) === String(id)) l.push(Calc.nombreTablero(t) + ' [' + (c.polos || []).join(',') + ']'); }));
    return l;
  };

  App.views.detalles = function (view) {
    const base = App.defCatalogo('detallesCarga'), tipos = App.defCatalogo('tiposCarga');
    const defD = Object.assign({}, base, {
      ayuda: 'Cada circuito usa un detalle de carga: voltaje y fases (reparto en el tablero), hilos (neutro), tipo de carga (factor de demanda del grupo), factor de demanda y de diversidad (0 a 1), FP, si es continua y el multiplicador de la protección (1,25 = 125 %). ' +
        'Las palabras clave son nombres o prefijos con que vienen las cargas en Revit, separados por coma (p. ej. "uc-, condensadora"): al importar, un circuito cuyo nombre contiene una palabra clave toma ese detalle sin ⚠. Un detalle monofásico a 208/240 V conectado en 2 polos reparte su carga entre los 2 polos.',
      idNuevo: l => Math.max(0, ...l.map(x => Number(x.id) || 0)) + 1,
      enUso: x => { const u = usos(x.id); return u.length ? 'El detalle ' + x.id + ' se usa en ' + u.length + ' circuito(s) del proyecto: ' + u.slice(0, 8).join(', ') + (u.length > 8 ? '…' : '') + '. Cámbielos antes de eliminarlo.' : ''; },
    });
    const defT = Object.assign({}, tipos, {
      ayuda: 'Los tipos de carga agrupan los circuitos para aplicar el factor de demanda (método fijo, tomas 220.44, cocina 220.56 o 220.53) y la diversidad del grupo.',
      enUso: x => (Store.catalog.detallesCarga.some(d => String(d.tipo) === String(x.id)) ? 'Hay detalles de carga de este tipo. Cámbielos antes de eliminarlo.' : ''),
    });
    view.appendChild(h('div', { class: 'page-h' }, h('h2', null, 'Detalles de carga'), h('p', { class: 'muted' }, 'Catálogo de cargas (como la hoja DCARGAS del Excel). Los cambios se aplican de inmediato a los cálculos y a la importación de Revit.')));
    view.appendChild(h('nav', { class: 'chips' },
      h('button', { type: 'button', class: 'chip' + (sec === 'detalles' ? ' active' : ''), onclick: () => { sec = 'detalles'; App.refresh(); } }, 'Detalles de carga (' + Store.catalog.detallesCarga.length + ')'),
      h('button', { type: 'button', class: 'chip' + (sec === 'tipos' ? ' active' : ''), onclick: () => { sec = 'tipos'; App.refresh(); } }, 'Tipos de carga (demanda)')));
    view.appendChild(UI.card(sec === 'detalles' ? 'Detalles de carga' : 'Tipos de carga', sec === 'detalles' ? App.tablaCatalogo('detallesCarga', defD) : App.tablaCatalogo('tiposCarga', defT)));
  };
})();
