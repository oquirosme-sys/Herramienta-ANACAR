/* Arranque, pestañas y menú Archivo. Cada vista se registra en App.views y recibe el cálculo completo del proyecto. */
(function (g) {
  'use strict';
  const h = U.h;

  const TABS = [
    { id: 'proyecto', label: 'Proyecto' },
    { id: 'unifilar', label: 'Diagrama unifilar' },
    { id: 'memoria', label: 'Memoria de cálculo' },
    { id: 'tab3f', label: 'Tableros 3F' },
    { id: 'tab1f', label: 'Tableros 1F' },
    { id: 'resumen', label: 'Tablas resumen' },
    { id: 'detalles', label: 'Detalles de carga' },
    { id: 'prevista', label: 'Prevista' },
    { id: 'revit', label: 'Revit' },
    { id: 'ayuda', label: 'Ayuda' },
    { id: 'admin', label: 'Administración', admin: true },
    { id: 'reporte', label: 'Reporte', oculto: true },
  ];

  const App = {
    views: {},
    R: null,
    route: { view: 'proyecto', id: '' },

    calc() { Store.sincronizarBypass(); App.R = Calc.proyecto(Store.project, Store.catalog); return App.R; },
    go(view, id) { location.hash = '#' + view + (id ? '/' + id : ''); },
    readHash() {
      const [v, id] = decodeURIComponent(location.hash.slice(1)).split('/');
      App.route = { view: TABS.some(t => t.id === v) ? v : 'proyecto', id: id || '' };
      if (App.route.view === 'admin' && !Auth.isAdmin()) App.route.view = 'proyecto';
    },

    /** Vuelve a calcular y dibujar la vista actual conservando el foco y el desplazamiento. */
    refresh() {
      const a = document.activeElement, fk = a && a.getAttribute && a.getAttribute('data-fk');
      const sx = window.scrollX, sy = window.scrollY;
      const wraps = Array.from(document.querySelectorAll('[data-scroll]')).map(e => [e.getAttribute('data-scroll'), e.scrollLeft, e.scrollTop]);
      App.render();
      window.scrollTo(sx, sy);
      wraps.forEach(([k, l, t]) => { const e = document.querySelector('[data-scroll="' + k + '"]'); if (e) { e.scrollLeft = l; e.scrollTop = t; } });
      if (fk) { const e = document.querySelector('[data-fk="' + CSS.escape(fk) + '"]'); if (e) e.focus({ preventScroll: true }); }
    },
    render() {
      App.calc();
      const nav = U.clear(document.getElementById('tabs'));
      TABS.filter(t => !t.oculto && (!t.admin || Auth.isAdmin())).forEach(t => nav.appendChild(h('a', { class: 'tab' + (t.admin ? ' admin' : '') + (App.route.view === t.id ? ' active' : ''), href: '#' + t.id, role: 'tab', 'aria-selected': App.route.view === t.id ? 'true' : 'false' }, t.label)));
      const view = U.clear(document.getElementById('view'));
      try { App.views[App.route.view](view, App.R, App.route.id); }
      catch (e) { console.error(e); view.appendChild(h('div', { class: 'empty err' }, 'Error al mostrar la vista: ' + e.message)); }
      App.actions();
    },

    actions() {
      const box = U.clear(document.getElementById('actions'));
      const st = h('span', { class: 'save-status' + (Store.status === 'error' ? ' error' : '') }, Store.status === 'saving' ? 'Guardando…' : Store.status === 'error' ? 'No se pudo guardar en el navegador' : 'Guardado en este navegador');
      const menu = h('details', { class: 'menu' }, h('summary', { class: 'btn' }, 'Archivo ▾'), h('div', { class: 'menu-b' },
        UI.btn('Nuevo proyecto (en blanco)', () => { Store.newProject(); App.go('proyecto'); App.render(); UI.toast('Proyecto nuevo. El anterior quedó en Proyectos guardados.', 'ok'); }, 'ghost'),
        UI.btn('Proyectos guardados…', () => App.proyectosGuardados(), 'ghost'),
        UI.btn('Abrir proyecto (.json)…', async () => { const f = await UI.pickFile('.json,application/json'); if (!f) return; try { Store.importProject(await UI.readText(f)); App.render(); UI.toast('Proyecto abierto', 'ok'); } catch (e) { UI.alert(e.message); } }, 'ghost'),
        UI.btn('Importar proyecto de ANACAR en Excel (.xlsm)…', () => App.importarExcel(), 'ghost'),
        UI.btn('Revisar el proyecto contra las pestañas 3F / 1F de un Excel…', () => App.revisarExcel(), 'ghost'),
        UI.btn('Guardar proyecto (.json)', () => U.download(App.fileBase() + '.json', Store.exportProject(), 'application/json'), 'ghost'),
        h('hr'),
        UI.btn('Exportar a Excel (tablas resumen y tableros)', () => ExportExcel.download(App.R), 'ghost'),
        UI.btn('Memoria de cálculo de todos (PDF / Word / Excel)', () => App.go('reporte', 'todos'), 'ghost'),
        UI.btn('Imprimir / PDF de la vista', () => window.print(), 'ghost')));
      menu.addEventListener('click', e => { if (e.target.closest('.menu-b .btn')) menu.open = false; });
      const adm = Auth.isAdmin()
        ? UI.btn('Administrador ✓', async () => { if (await UI.confirm('¿Salir del modo administrador?', 'Salir', false)) { Auth.logout(); } }, 'admin-on', 'Salir del modo administrador')
        : UI.btn('Administrador', () => App.login(), 'ghost', 'Entrar como administrador para editar catálogos');
      box.append(st, menu, adm);
    },
    /** Importa un proyecto hecho en ANACAR (Excel): crea el proyecto en la herramienta y muestra el reporte de lo copiado y lo recalculado. */
    async importarExcel() {
      const f = await UI.pickFile('.xlsm,.xlsx'); if (!f) return;
      try {
        UI.toast('Importando ' + f.name + '… los proyectos grandes tardan unos segundos', 'ok'); await new Promise(r => setTimeout(r, 60));
        const rep = await ImportAnacar.importar(await UI.readBuffer(f), f.name);
        App.render();
        const cv = rep.dif.filter(d => /caída de voltaje/i.test(d.que)), otras = rep.dif.filter(d => !/caída de voltaje/i.test(d.que));
        const fila = d => h('tr', { title: d.nota || null }, h('td', null, d.tablero), h('td', null, d.que + (d.nota ? ' ⓘ' : '')), h('td', { class: 'r' }, String(d.excel)), h('td', { class: 'r' }, String(d.web)));
        const tabla = rep.dif.length ? h('div', null, otras.length ? h('div', { class: 'tbl-wrap' }, h('table', { class: 'tbl' }, h('thead', null, h('tr', null, ['Tablero', 'Dato', 'Excel', 'Herramienta'].map(x => h('th', null, x)))), h('tbody', null, otras.slice(0, 80).map(fila)))) : null, cv.length ? h('p', { class: 'muted' }, cv.length + ' valor(es) de caída de voltaje difieren en más de 0,15 %: el Excel y la herramienta acumulan la caída de forma distinta entre tableros (la herramienta la suma desde la acometida).') : null, otras.some(d => d.nota) ? h('ul', null, Array.from(new Set(otras.filter(d => d.nota).map(d => d.nota))).map(n => h('li', { class: 'muted' }, 'ⓘ ' + n))) : null) : h('p', { class: 'muted' }, 'Sin diferencias: la herramienta recalculó los mismos valores que dejó Excel.');
        UI.modal('Proyecto importado: ' + Store.project.nombre, h('div', null,
          h('p', null, rep.tableros + ' tablero(s) y ' + rep.circuitos + ' circuito(s) copiados. Cargas, longitudes, posiciones, protecciones, calibres, interruptores y tableros se tomaron del Excel; caídas de voltaje, balanceo, avisos y cortocircuito los recalculó la herramienta.'),
          rep.aviso.length ? h('div', null, h('h4', null, 'Para revisar (' + rep.aviso.length + ')'), h('ul', null, rep.aviso.map(a => h('li', null, a)))) : null,
          h('h4', null, 'Comparación con el Excel' + (rep.dif.length ? ' (' + otras.length + ' diferencia(s) de cálculo' + (otras.length > 80 ? ', se muestran 80' : '') + (cv.length ? ' y ' + cv.length + ' de caída de voltaje' : '') + ')' : '')), tabla), [rep.rev && rep.rev.hayFichas ? { label: 'Revisión contra pestañas 3F / 1F (' + rep.rev.resumen.revisar + ' para revisar)', onclick: () => { RevisorAnacar.mostrar(rep.rev); return false; } } : null, { label: 'Cerrar', cls: 'primary' }].filter(Boolean));
        App.go('proyecto');
      } catch (e) { UI.alert('No se pudo importar: ' + e.message); }
    },
    /** Revisa los tableros del proyecto abierto contra las fichas de las pestañas 3F / 1F de un Excel de ANACAR. */
    async revisarExcel() {
      const f = await UI.pickFile('.xlsm,.xlsx'); if (!f) return;
      try { UI.toast('Leyendo ' + f.name + '…', 'ok'); await new Promise(r => setTimeout(r, 60)); RevisorAnacar.mostrar(await RevisorAnacar.revisarArchivo(await UI.readBuffer(f), f.name)); } catch (e) { UI.alert('No se pudo revisar: ' + e.message); }
    },
    /** Lista de proyectos guardados en este navegador (abrir o eliminar). */
    proyectosGuardados(contenedor) {
      const lista = Store.listaProyectos(); let cerrar = null;
      const tabla = lista.length ? h('table', { class: 'tbl' }, h('thead', null, h('tr', null, ['Proyecto', 'N.º', 'Tableros', 'Modificado', ''].map(x => h('th', null, x)))),
        h('tbody', null, lista.map(p => h('tr', { class: p.id === Store.project.id ? 'row-link' : '' }, h('td', null, h('b', null, p.nombre || '(sin nombre)')), h('td', null, p.numero || ''), h('td', { class: 'r' }, p.tableros),
          h('td', null, String(p.modificado).replace('T', ' ').slice(0, 16)),
          h('td', { class: 'acc' }, UI.btn(p.id === Store.project.id ? 'Abierto' : 'Abrir', () => { Store.abrirProyecto(p.id); App.go('proyecto'); App.render(); if (cerrar) cerrar(); }, 'small'),
            UI.iconBtn('🗑', 'Eliminar', async () => { if (await UI.confirm('¿Eliminar el proyecto "' + (p.nombre || 'sin nombre') + '" de este navegador? Si no lo exportó a .json se pierde.', 'Eliminar')) { Store.eliminarProyecto(p.id); if (cerrar) cerrar(); App.render(); if (!contenedor) App.proyectosGuardados(); } }, 'danger')))))) 
        : h('p', { class: 'muted' }, 'No hay proyectos guardados en este navegador.');
      if (contenedor) return tabla;
      cerrar = UI.modal('Proyectos guardados en este navegador', tabla, [{ label: 'Cerrar' }], { wide: true });
      return null;
    },
    fileBase() { const p = Store.project; return ('ANACAR ' + (p.numero || '') + ' ' + (p.nombre || 'proyecto')).trim().replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, ' '); },
    login() {
      const inp = h('input', { type: 'password', autocomplete: 'current-password' }), err = h('small', { class: 'err' });
      UI.modal('Modo administrador', h('div', null, h('p', { class: 'muted' }, 'El administrador puede agregar o modificar marcas, tableros, breakers, supresores, tipos de carga y tablas de conductores.'), UI.field('Contraseña', inp), err), [
        { label: 'Cancelar' },
        { label: 'Entrar', cls: 'primary', onclick: () => { if (!Auth.login(inp.value)) { err.textContent = 'Contraseña incorrecta.'; return false; } UI.toast('Modo administrador activo', 'ok'); } },
      ]);
      inp.addEventListener('keydown', e => { if (e.key === 'Enter') document.querySelector('.modal-f .btn.primary').click(); });
    },
  };

  g.App = App;
  document.addEventListener('DOMContentLoaded', () => {
    Store.load(); Auth.init();
    Store.onChange(() => { const s = document.querySelector('.save-status'); if (s) { s.textContent = Store.status === 'saving' ? 'Guardando…' : Store.status === 'error' ? 'No se pudo guardar en el navegador' : 'Guardado en este navegador'; s.classList.toggle('error', Store.status === 'error'); } });
    Auth.onChange(() => { App.readHash(); App.render(); });
    window.addEventListener('hashchange', () => { App.readHash(); App.render(); window.scrollTo(0, 0); });
    App.readHash(); App.render();
  });
})(window);
