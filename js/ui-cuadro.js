/* Pestañas "Tableros 3F" y "Tableros 1F": cuadro de cargas de cada tablero con el mismo diseño de las hojas TABLEROS 3F / 1F del Excel. */
(function () {
  'use strict';
  const h = U.h;

  function vista(tipo) {
    const ruta = tipo === '3F' ? 'tab3f' : 'tab1f';
    return function (view, R, id) {
      const lista = R.orden.filter(r => r.tab.tipo === tipo);
      view.appendChild(h('div', { class: 'page-h row no-print' }, h('div', null, h('h2', null, 'Tableros ' + tipo), h('p', { class: 'muted' }, 'Mismo diseño de la hoja TABLEROS ' + tipo + ' del Excel. Las posiciones impares van primero y luego las pares.')),
        h('div', { class: 'toolbar' },
          h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: Hoja.real, onchange: e => { Hoja.real = e.target.checked; App.refresh(); } }), ' Tamaño real'),
          UI.btn('Imprimir / PDF', () => window.print(), 'small'), UI.btn('Exportar a Excel', () => ExportExcel.download(R, { soloTipo: tipo }), 'primary small'))));
      if (!lista.length) { view.appendChild(h('div', { class: 'empty' }, 'No hay tableros ' + tipo + '.')); return; }
      const todos = id === 'todos', sel = todos ? null : (lista.find(r => r.tab.id === id) || lista[0]);
      view.appendChild(h('nav', { class: 'chips no-print' }, lista.map(r => h('a', { class: 'chip' + (sel && r.tab.id === sel.tab.id ? ' active' : '') + (r.avisos.length ? ' warn' : ''), href: '#' + ruta + '/' + r.tab.id, title: r.avisos.join('\n') || null }, r.nombre)),
        h('a', { class: 'chip' + (todos ? ' active' : ''), href: '#' + ruta + '/todos' }, 'Todos (para imprimir)')));
      (todos ? lista : [sel]).forEach(r => {
        view.appendChild(h('article', { class: 'hoja' },
          h('div', { class: 'hoja-acc no-print' }, UI.btn('Memoria de cálculo', () => App.go('memoria', r.tab.id), 'small ghost'), r.avisos.length ? h('span', { class: 'badge warn', title: r.avisos.join('\n') }, r.avisos.length + ' aviso(s)') : null),
          Hoja.html(Hoja.tablero(r), { real: Hoja.real })));
      });
    };
  }
  App.views.tab3f = vista('3F');
  App.views.tab1f = vista('1F');
})();
