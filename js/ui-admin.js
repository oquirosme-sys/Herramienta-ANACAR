/* Administración (solo administrador): catálogos que en el Excel estaban en las hojas DCARGAS, CONDISEÑO y DATOS (ocultas).
   Marcas, tableros, breakers, supresores, tipos y detalles de carga, tipos de cableado, tablas NEC, transformadores, seguridad y respaldo. */
(function () {
  'use strict';
  const h = U.h;
  let filtro = '', filtroMarca = '';

  const C = () => Store.catalog;
  const marcas = () => C().marcas.map(m => ({ value: m, label: m }));
  const tipos = () => C().tiposCarga.map(t => ({ value: t.id, label: t.id + ' · ' + t.nombre }));
  const SECC = {
    marcas: { title: 'Marcas', especial: true },
    tiposCarga: { title: 'Tipos de carga (demanda)', cols: [['id', 'Código', 'num'], ['nombre', 'Nombre'], ['metodo', 'Método', 'sel', () => [{ value: 'fijo', label: 'Factor fijo' }, { value: 'tomas', label: 'Tomas: 10 kVA + 50 %' }, { value: 'cocina', label: 'Cocina: tabla 220.56' }, { value: '220.53', label: '≥ 4 equipos: 75 % (220.53)' }]], ['fd', 'F. demanda', 'num'], ['fdiv', 'F. diversidad', 'num']] },
    detallesCarga: { title: 'Detalles de carga (DCARGAS)', cols: [['id', 'Detalle', 'num'], ['tipo', 'Tipo de carga', 'sel', tipos, true], ['descripcion', 'Descripción'], ['v', 'V', 'num'], ['fases', 'Fases', 'num'], ['hilos', 'Hilos', 'num'], ['fd', 'F. dem.', 'num'], ['fdiv', 'F. div.', 'num'], ['fp', 'FP', 'num'], ['continua', 'Continua', 'bool'], ['mult', 'Mult.', 'num'], ['fundamento', 'Fundamento NEC']] },
    tablerosCat: { title: 'Tableros', marca: 'fabricante', cols: [['id', 'Tipo', 'num'], ['modelo', 'Modelo de referencia'], ['fabricante', 'Fabricante', 'sel', marcas], ['fases', 'Fases', 'sel', () => ['3F', '1F', '1F/3F']], ['barraFase', 'Barra fase (A)', 'num'], ['barraNeutro', 'Barra neutro (A)', 'num'], ['barraTierra', 'Barra tierra (A)', 'num'], ['espacios', 'Espacios', 'num'], ['nota', 'Nota']] },
    breakers: { title: 'Breakers', marca: 'marca', cols: [['marca', 'Marca', 'sel', marcas], ['id', 'Tipo', 'num'], ['modelo', 'Modelo'], ['marco', 'Marco (A)', 'num'], ['amperios', 'Amperios (rango)'], ['unidad', 'Tipo de unidad'], ['polos', 'Polos', 'num'], ['sccr', 'SCCR (kA)', 'num'], ['vSccr', 'SCCR a (V)', 'num'], ['nota', 'Nota']] },
    reglasTableros: { title: 'Reglas: familias de tableros', get: () => C().reglas.tableros, ayuda: 'El tablero automático se elige entre las familias de la marca cuyo voltaje máximo y fases sirven, con barras ≥ protección principal y espacios ≥ usados + reserva; primero la de menor "Preferencia". Los breakers ramales y principales se buscan solo en las familias indicadas (separadas por coma) y se elige el de menor SCCR que supere el Icc del tablero.', cols: [['marca', 'Marca', 'sel', marcas], ['familia', 'Familia'], ['patron', 'Patrón del modelo (expresión regular)'], ['vMax', 'V máx.', 'num'], ['fases', 'Fases (1F/3F)'], ['ramales', 'Breakers ramales'], ['principales', 'Breakers principales'], ['orden', 'Preferencia', 'num'], ['nota', 'Nota']] },
    reglasBreakers: { title: 'Reglas: familias de breakers', get: () => C().reglas.breakers, ayuda: 'La familia de cada breaker se reconoce por su modelo. La plantilla arma el número de catálogo: {p} polos, {a} amperios, {a3} amperios con 3 dígitos, {m3} primeras 3 letras del modelo (p. ej. QOB{p}{a} → QOB120; EDB{p}4{a3} → EDB34020). Vacía = se usa el modelo del catálogo.', cols: [['marca', 'Marca', 'sel', marcas], ['familia', 'Familia'], ['patron', 'Patrón del modelo (expresión regular)'], ['vMax', 'V máx.', 'num'], ['plantilla', 'Plantilla del número de catálogo']] },
    supresores: { title: 'Supresores (SPD)', marca: 'marca', cols: [['marca', 'Marca', 'sel', marcas], ['id', 'Tipo', 'num'], ['modelo', 'Modelo'], ['montaje', 'Montaje', 'sel', () => ['Interno', 'Externo']], ['kaLL', 'kA L-L', 'num'], ['kaLN', 'kA L-N', 'num'], ['voltaje', 'Voltaje'], ['fases', 'Fases', 'num'], ['conexion', 'Conexión']] },
    cableado: { title: 'Tipos de cableado', especial: true },
    ampCU: { title: 'Calibre por ampacidad — cobre (310.16)', cols: [['amp', 'Desde (A)', 'num'], ['cal', 'Calibre']] },
    ampAL: { title: 'Calibre por ampacidad — aluminio (310.16)', cols: [['amp', 'Desde (A)', 'num'], ['cal', 'Calibre']] },
    protecciones: { title: 'Protección por ampacidad requerida', cols: [['amp', 'Desde (A)', 'num'], ['prot', 'Protección (A)', 'num']] },
    tierras: { title: 'Tierra de equipo (250.122)', cols: [['amp', 'Hasta (A)', 'num'], ['CU', 'CU'], ['AL', 'AL'], ['CU-MC', 'CU-MC'], ['AL-MC', 'AL-MC']] },
    conduit: { title: 'Tubería por calibre (C.10, mm)', cols: [['cal', 'Calibre'], ['THHN', 'THHN', 'num'], ['RHW', 'RHW', 'num'], ['XHHW', 'XHHW', 'num']] },
    electrodo: { title: 'Conductor del electrodo (250.66)', cols: [['calCU', 'Calibre CU'], ['CU', 'Electrodo CU'], ['CU-MC', 'CU-MC'], ['calAL', 'Calibre AL'], ['AL', 'Electrodo AL'], ['AL-MC', 'AL-MC']] },
    transformadores: { title: 'Transformadores (Icc)', cols: [['id', '#', 'num'], ['nombre', 'Nombre'], ['tipo', 'Tipo', 'sel', () => ['MONOFASICO', 'TRIFASICO']], ['kacc480', 'KACC 480 V', 'num'], ['kacc240', 'KACC 240 V', 'num'], ['kacc208', 'KACC 208 V', 'num'], ['z', 'Z %', 'num']] },
    avanzado: { title: 'Tablas avanzadas (JSON)', especial: true },
    seguridad: { title: 'Seguridad y respaldo', especial: true },
  };

  function tabla(key, def) {
    let lista = def.get ? def.get() : C()[key];
    const q = filtro.toLowerCase();
    const filas = lista.map((x, i) => ({ x, i })).filter(({ x }) => (!def.marca || !filtroMarca || x[def.marca] === filtroMarca) && (!q || def.cols.some(c => String(x[c[0]] === undefined || x[c[0]] === null ? '' : x[c[0]]).toLowerCase().includes(q))));
    const cel = (x, c, i) => {
      const [k, , tp, op] = c, fk = 'adm:' + key + ':' + i + ':' + k;
      if (tp === 'bool') return h('input', { type: 'checkbox', checked: !!x[k], 'data-fk': fk, onchange: e => { x[k] = e.target.checked; Store.save(); } });
      if (tp === 'sel') return UI.select(UI.opts(op(), ''), x[k], v => { x[k] = c[4] || (typeof x[k] === 'number' && v !== '') ? Number(v) : v; Store.save(); }, { fk });
      return UI.input(x, k, { type: tp === 'num' ? 'num' : 'text', fk, class: tp === 'num' ? 'w-num' : '', after: () => {} });
    };
    const body = h('tbody', null, filas.slice(0, 400).map(({ x, i }) => h('tr', null, def.cols.map(c => h('td', null, cel(x, c, i))),
      h('td', { class: 'acc' }, UI.iconBtn('⧉', 'Duplicar', () => { lista.splice(i + 1, 0, U.clone(x)); Store.save(); App.refresh(); }), UI.iconBtn('🗑', 'Eliminar', () => { lista.splice(i, 1); Store.save(); App.refresh(); }, 'danger')))));
    const nuevo = () => {
      const o = {}; def.cols.forEach(([k, , tp]) => { o[k] = tp === 'num' ? '' : tp === 'bool' ? false : ''; });
      if (def.cols[0][0] === 'id' || def.cols.some(c => c[0] === 'id')) o.id = Math.max(0, ...lista.filter(x => !def.marca || !filtroMarca || x[def.marca] === filtroMarca).map(x => Number(x.id) || 0)) + 1;
      if (def.marca && filtroMarca) o[def.marca] = filtroMarca;
      lista.push(o); Store.save(); filtro = ''; App.refresh();
      setTimeout(() => { const w = document.querySelector('[data-scroll="adm-' + key + '"]'); if (w) w.scrollTop = w.scrollHeight; }, 0);
    };
    return h('div', null, def.ayuda ? h('p', { class: 'hint' }, def.ayuda) : null,
      h('div', { class: 'toolbar' },
        h('input', { type: 'search', placeholder: 'Buscar…', value: filtro, oninput: e => { filtro = e.target.value; clearTimeout(tabla.t); tabla.t = setTimeout(App.refresh, 250); }, 'data-fk': 'adm-buscar' }),
        def.marca ? UI.select(UI.opts(C().marcas, 'Todas las marcas'), filtroMarca, v => { filtroMarca = v; App.refresh(); }) : null,
        h('span', { class: 'muted' }, filas.length + ' de ' + lista.length + (filas.length > 400 ? ' (se muestran 400; use el buscador)' : '')), h('span', { class: 'grow' }),
        UI.btn('+ Agregar', nuevo, 'primary small')),
      h('div', { class: 'tbl-wrap tall', 'data-scroll': 'adm-' + key }, h('table', { class: 'tbl admin' }, h('thead', null, h('tr', null, def.cols.map(c => h('th', null, c[1])), h('th'))), body)),
      def.marca ? h('p', { class: 'hint' }, 'El "Tipo" es el número que se elige en la memoria de cálculo; es propio de cada marca (como en el Excel).') : null);
  }

  function seccMarcas() {
    const usos = m => C().tablerosCat.filter(t => t.fabricante === m).length + ' tableros · ' + C().breakers.filter(b => b.marca === m).length + ' breakers · ' + C().supresores.filter(b => b.marca === m).length + ' supresores';
    return h('div', null,
      h('table', { class: 'tbl narrow' }, h('thead', null, h('tr', null, h('th', null, 'Marca'), h('th', null, 'Uso en el catálogo'), h('th'))),
        h('tbody', null, C().marcas.map((m, i) => h('tr', null, h('td', null, h('b', null, m)), h('td', { class: 'muted' }, usos(m)), h('td', { class: 'acc' },
          UI.btn('Renombrar', async () => {
            const n = await UI.prompt('Renombrar marca', 'Nombre', m, v => (!v ? 'Escriba un nombre' : (v !== m && C().marcas.includes(v) ? 'Ya existe' : ''))); if (!n || n === m) return;
            C().marcas[i] = n; C().tablerosCat.forEach(t => { if (t.fabricante === m) t.fabricante = n; }); C().breakers.forEach(b => { if (b.marca === m) b.marca = n; }); C().supresores.forEach(b => { if (b.marca === m) b.marca = n; });
            Store.project.tableros.forEach(t => { if (t.marca === m) t.marca = n; }); if (Store.project.marcaDefecto === m) Store.project.marcaDefecto = n; Store.save(); App.refresh();
          }, 'small'),
          UI.btn('Copiar catálogo a…', async () => {
            const n = await UI.prompt('Nueva marca a partir de ' + m, 'Nombre de la nueva marca', '', v => (!v ? 'Escriba un nombre' : (C().marcas.includes(v) ? 'Ya existe' : ''))); if (!n) return;
            C().marcas.push(n);
            ['breakers', 'supresores'].forEach(k => C()[k].filter(b => b.marca === m).forEach(b => C()[k].push(Object.assign(U.clone(b), { marca: n }))));
            Store.save(); App.refresh(); UI.toast('Marca ' + n + ' creada con los breakers y supresores de ' + m + ' para editarlos', 'ok');
          }, 'small', 'Crea la marca con una copia de los breakers y supresores para ajustar modelos'),
          UI.iconBtn('🗑', 'Eliminar marca', async () => {
            if (!(await UI.confirm('¿Eliminar la marca ' + m + ' y todos sus tableros, breakers y supresores del catálogo?', 'Eliminar'))) return;
            C().marcas.splice(i, 1); C().tablerosCat = C().tablerosCat.filter(t => t.fabricante !== m); C().breakers = C().breakers.filter(b => b.marca !== m); C().supresores = C().supresores.filter(b => b.marca !== m);
            Store.save(); App.refresh();
          }, 'danger')))))),
      h('div', { class: 'toolbar' }, UI.btn('+ Marca', async () => { const n = await UI.prompt('Nueva marca', 'Nombre', '', v => (!v ? 'Escriba un nombre' : (C().marcas.includes(v) ? 'Ya existe' : ''))); if (n) { C().marcas.push(n); Store.save(); App.refresh(); } }, 'primary small')));
  }

  function lista(titulo, arr, ayuda) {
    return UI.card(titulo, h('div', null, ayuda ? h('p', { class: 'hint' }, ayuda) : null,
      h('div', { class: 'tags' }, arr.map((x, i) => h('span', { class: 'tag-ed' }, x, UI.iconBtn('✕', 'Quitar ' + x, () => { arr.splice(i, 1); Store.save(); App.refresh(); })))),
      h('div', { class: 'toolbar' }, UI.btn('+ Agregar', async () => { const v = await UI.prompt(titulo, 'Valor', '', s => (!s ? 'Escriba un valor' : arr.includes(s) ? 'Ya existe' : '')); if (v) { arr.push(v); Store.save(); App.refresh(); } }, 'small'))));
  }
  function seccCableado() {
    const L = C().listas;
    return h('div', { class: 'cols' },
      lista('Aislamientos', C().aislamientos, 'Para la tubería se usan las columnas THHN, RHW y XHHW de la tabla C.10; otros aislamientos usan THHN. Los terminados en "/ MC" o materiales -MC se consideran cable MC.'),
      lista('Materiales', L.materiales, 'CU y AL usan sus tablas de ampacidad; CU-MC y AL-MC son cable armado (sin tubería).'),
      lista('Tipos de tubería', L.tuberias, 'Se usan con las constantes C del cálculo de cortocircuito.'),
      lista('Sistemas de voltaje', L.sistemas, 'Formato L-N/L-L, p. ej. 120/208.'));
  }
  function seccAvanzado() {
    const claves = ['fac', 'constC', 'ampacidad31016', 'tempFactor', 'agrupamiento', 'demanda22056', 'listas', 'leyendaUnidades'];
    let sel = claves[0];
    const ta = h('textarea', { rows: 22, class: 'mono', spellcheck: 'false' }); ta.value = JSON.stringify(C()[sel], null, 1);
    const s = UI.select(claves, sel, v => { sel = v; ta.value = JSON.stringify(C()[v], null, 1); });
    return h('div', null, h('p', { class: 'hint' }, 'FAC de caída de tensión, constantes C de cortocircuito, ampacidades 310.16 y factores de corrección. Edite con cuidado: el JSON debe ser válido.'),
      h('div', { class: 'toolbar' }, s, UI.btn('Guardar tabla', () => { try { C()[sel] = JSON.parse(ta.value); Store.save(); UI.toast('Tabla guardada', 'ok'); App.refresh(); } catch (e) { UI.alert('JSON no válido: ' + e.message); } }, 'primary small')), ta);
  }
  function seccSeguridad() {
    const p1 = h('input', { type: 'password', autocomplete: 'new-password' }), p2 = h('input', { type: 'password', autocomplete: 'new-password' });
    return h('div', { class: 'cols' },
      UI.card('Contraseña de administrador', h('div', null, UI.field('Nueva contraseña', p1), UI.field('Repetir', p2),
        h('div', { class: 'toolbar' }, UI.btn('Cambiar', () => { if (p1.value.length < 6) return UI.alert('Use al menos 6 caracteres.'); if (p1.value !== p2.value) return UI.alert('Las contraseñas no coinciden.'); Store.setAdminHash(U.sha256(p1.value)); p1.value = p2.value = ''; UI.toast('Contraseña actualizada', 'ok'); }, 'primary small')),
        h('p', { class: 'hint' }, 'En esta versión el control es local (solo oculta la edición). En el paso 2, Supabase Auth con roles y Row Level Security restringirá la escritura del catálogo a los administradores.'))),
      UI.card('Respaldo del catálogo', h('div', null,
        h('p', null, 'Comparta el catálogo con el equipo exportándolo y que cada usuario lo importe (en el paso 2 vendrá de la base de datos).'),
        h('div', { class: 'toolbar' },
          UI.btn('Exportar catálogo (.json)', () => U.download('Catalogo ANACAR ' + U.today() + '.json', Store.exportCatalog(), 'application/json'), 'small'),
          UI.btn('Importar catálogo…', async () => { const f = await UI.pickFile('.json'); if (!f) return; try { Store.importCatalog(await UI.readText(f)); App.refresh(); UI.toast('Catálogo importado', 'ok'); } catch (e) { UI.alert(e.message); } }, 'small'),
          UI.btn('Restablecer al Excel original', async () => { if (await UI.confirm('¿Reemplazar todo el catálogo por el del Excel ANACAR 2026.1? Se pierden los cambios del administrador.', 'Restablecer')) { Store.resetCatalog(); App.refresh(); } }, 'danger small')))));
  }

  App.views.admin = function (view, R, sec) {
    if (!Auth.isAdmin()) { view.appendChild(h('div', { class: 'empty' }, 'Solo el administrador puede ver esta pestaña.')); return; }
    sec = SECC[sec] ? sec : 'marcas';
    if (App._admSec !== sec) { filtro = ''; App._admSec = sec; }
    view.appendChild(h('div', { class: 'page-h' }, h('h2', null, 'Administración'), h('p', { class: 'muted' }, 'Catálogos compartidos por todos los proyectos. Los cambios se reflejan de inmediato en los cálculos.')));
    view.appendChild(h('div', { class: 'admin-layout' },
      h('nav', { class: 'side' }, Object.keys(SECC).map(k => h('a', { class: k === sec ? 'active' : '', href: '#admin/' + k }, SECC[k].title))),
      h('div', { class: 'admin-main' }, UI.card(SECC[sec].title,
        sec === 'marcas' ? seccMarcas() : sec === 'cableado' ? seccCableado() : sec === 'avanzado' ? seccAvanzado() : sec === 'seguridad' ? seccSeguridad() : tabla(sec, SECC[sec])))));
  };
})();
