/* Estado de la aplicación y persistencia.
   Toda la lectura/escritura pasa por Store.* para que el paso 2 (Supabase) solo reemplace este archivo
   (o agregue un adaptador) sin tocar la interfaz. Hoy: localStorage + importar/exportar JSON.
   Modelo:
     Catálogo (solo administrador): tiposCarga, detallesCarga, tablerosCat, breakers, supresores, transformadores, tablas NEC…
     Proyecto: datos generales + tableros[] → cada tablero con su alimentador (alim), circuitos[] y padreId (de quién se alimenta). */
(function (g) {
  'use strict';
  const K_PROJECT = 'anacar.proyecto.v1';      // versión anterior (un solo proyecto): se pasa a la lista
  const K_PROYECTOS = 'anacar.proyectos.v1';    // { id: proyecto } guardados en este navegador
  const K_ACTUAL = 'anacar.actual.v1';          // proyecto abierto en esta pestaña (sessionStorage)
  const K_CATALOG = 'anacar.catalogo.v1';
  const K_ADMINPW = 'anacar.adminhash.v1';

  const listeners = [];
  let saveTimer = null;

  const Store = {
    project: null,
    catalog: null,
    status: 'ok',

    onChange(fn) { listeners.push(fn); },
    _emit() { listeners.forEach(f => f(Store.status)); },

    /** Al abrir la herramienta se empieza con un proyecto limpio; al recargar la pestaña se sigue con el que estaba abierto.
        Los proyectos con contenido quedan en la lista de proyectos guardados de este navegador. */
    load() {
      let c = null;
      try { c = JSON.parse(localStorage.getItem(K_CATALOG)); } catch (e) { /* ignora */ }
      Store.catalog = c && c.detallesCarga && c.breakers ? c : Store.seedCatalog();
      Store.actualizarCatalogo();
      const lista = Store.proyectos();
      try { // proyecto único de la versión anterior → lista
        const viejo = JSON.parse(localStorage.getItem(K_PROJECT));
        if (viejo && viejo.tableros) { if (!lista[viejo.id]) { lista[viejo.id] = viejo; localStorage.setItem(K_PROYECTOS, JSON.stringify(lista)); } localStorage.removeItem(K_PROJECT); }
      } catch (e) { /* ignora */ }
      let actual = null; try { actual = sessionStorage.getItem(K_ACTUAL); } catch (e) { /* ignora */ }
      Store.project = actual && lista[actual] ? lista[actual] : Store.emptyProject();
      Store.migrate();
    },
    /** Un catálogo guardado en este navegador recibe, una sola vez por revisión, los tableros nuevos del catálogo base (no se borra ni se cambia nada de lo que el administrador editó). */
    actualizarCatalogo() {
      const seed = window.SEED || {}, c = Store.catalog;
      if (!c || (c.catRev || 1) >= (seed.catRev || 1)) { if (c && !c.catRev) c.catRev = seed.catRev || 1; return; }
      const llave = x => [x.fabricante, x.modelo, x.barraFase, x.espacios, x.fases].join('|');
      const hay = new Set((c.tablerosCat || []).map(llave));
      let id = Math.max(0, ...(c.tablerosCat || []).map(x => Number(x.id) || 0));
      (seed.tablerosCat || []).forEach(x => { if (!hay.has(llave(x))) c.tablerosCat.push(Object.assign({}, x, { id: ++id })); });
      c.catRev = seed.catRev; try { localStorage.setItem(K_CATALOG, JSON.stringify(c)); } catch (e) { /* ignora */ }
    },
    proyectos() { try { return JSON.parse(localStorage.getItem(K_PROYECTOS)) || {}; } catch (e) { return {}; } },
    /** Lista para mostrar: los guardados, del más reciente al más antiguo. */
    listaProyectos() {
      return Object.values(Store.proyectos()).map(p => ({ id: p.id, nombre: p.nombre, numero: p.numero, tableros: (p.tableros || []).length, modificado: p.modificado || p.fecha || '' }))
        .sort((a, b) => String(b.modificado).localeCompare(String(a.modificado)));
    },
    abrirProyecto(id) {
      const p = Store.proyectos()[id]; if (!p) return false;
      Store.flush(); Store.project = p; Store.migrate();
      try { sessionStorage.setItem(K_ACTUAL, p.id); } catch (e) { /* ignora */ }
      return true;
    },
    eliminarProyecto(id) {
      const l = Store.proyectos(); delete l[id]; localStorage.setItem(K_PROYECTOS, JSON.stringify(l));
      if (Store.project.id === id) Store.newProject();
    },
    seedCatalog() { const s = U.clone(window.SEED); s.marcas = Store.marcasDe(s); return s; },
    marcasDe(c) { return Array.from(new Set([].concat(c.tablerosCat.map(t => t.fabricante), c.breakers.map(b => b.marca), c.supresores.map(b => b.marca)).filter(Boolean))).sort(); },
    emptyProject() {
      return {
        id: U.uid(), nombre: '', numero: '', ubicacion: '', fecha: U.today(), elaboro: '', revision: '',
        cvMaxRamal: 3, cvMaxAlim: 3, cvMaxTotal: 5, desbalanceMax: 10, iccLongMax: 20, marcaDefecto: 'Eaton', reservaEspacios: 20, autoAmpacidad: true, iccRed: '', resumenSimple: false, ocupacion: 'comercial', cargaDerivados: 'conectada', pais: 'CR', previstas: [], fuentes: [], tableros: [],
      };
    },
    migrate() {
      const c = Store.catalog, s = window.SEED, norm0 = x => String(x || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
      // catálogos guardados antes de las reglas de fabricante: tablas NEC completas, SCCR a 240/480 V y fases de cada tablero
      if (!c.reglas) {
        ['reglas', 'ampacidad31016', 'tempFactor'].forEach(k => { c[k] = U.clone(s[k]); });
        c.listas.temperaturas = s.listas.temperaturas.slice();
      }
      // catálogos anteriores a las subestaciones (QED-2, Pow-R-Line Xpert) y al tipo de principal (zapatas / interruptor)
      if (!c.reglas.tableros.some(f => f.familia === 'QED-2')) {
        c.reglas.tableros.forEach(f => { const o = s.reglas.tableros.find(x => x.marca === f.marca && x.familia === f.familia); if (o && f.clase === undefined) f.clase = o.clase; });
        s.reglas.tableros.filter(x => x.clase === 'subestacion').forEach(x => c.reglas.tableros.push(U.clone(x)));
        c.reglas.breakers.forEach(f => { if (/^PDG/.test(f.familia) && !f.plantilla) f.plantilla = '{pdg}'; });
        s.tablerosCat.filter(x => /^(QED|PRLX)/.test(x.modelo)).forEach(x => { if (!c.tablerosCat.some(y => y.id === x.id)) c.tablerosCat.push(U.clone(x)); });
      }
      c.reglas.breakers.forEach(f => { if (f.familia === 'GFCI/AFCI' && !f.plantilla) { const o = s.reglas.breakers.find(x => x.marca === f.marca && x.familia === f.familia); if (o) f.plantilla = o.plantilla; } });
      c.tablerosCat.forEach(t => { if (t.principal === undefined) { const o = s.tablerosCat.find(x => x.id === t.id && x.modelo === t.modelo); t.principal = o ? o.principal : 'Ambos'; } });
      c.breakers.forEach(b => { if (b.vSccr === undefined) { const o = s.breakers.find(x => x.marca === b.marca && x.id === b.id); if (o) b.vSccr = o.vSccr; } });
      c.tablerosCat.forEach(t => { if (t.fases === undefined) { const o = s.tablerosCat.find(x => x.id === t.id); if (o) t.fases = o.fases; } });
      Object.keys(s).forEach(k => { if (c[k] === undefined) c[k] = U.clone(s[k]); });
      if (!c.marcas) c.marcas = Store.marcasDe(c);
      // factores de diversidad: antes dividían (≥ 1, como el Excel); ahora van de 0 a 1 y multiplican (1/x da el mismo resultado)
      const inv = v => (Number(v) > 1 ? Math.round(1e6 / Number(v)) / 1e6 : v);
      // palabras clave de Revit en los detalles de carga y detalles nuevos del catálogo base (p. ej. deshumidificador)
      if (!c.claves01) {
        (c.detallesCarga || []).forEach(d => { if (d.claves === undefined) { const o = s.detallesCarga.find(x => x.id === d.id && norm0(x.descripcion) === norm0(d.descripcion)); d.claves = o ? o.claves : ''; } });
        s.detallesCarga.filter(x => !(c.detallesCarga || []).some(d => d.id === x.id)).forEach(x => c.detallesCarga.push(U.clone(x)));
        c.claves01 = true;
      }
      if (!c.fdiv01) { (c.detallesCarga || []).forEach(d => { d.fdiv = inv(d.fdiv); }); (c.tiposCarga || []).forEach(t => { t.fdiv = inv(t.fdiv); }); c.fdiv01 = true; }
      const p = Store.project;
      ['cvMaxRamal', 'cvMaxAlim', 'cvMaxTotal', 'desbalanceMax', 'iccLongMax', 'marcaDefecto', 'reservaEspacios', 'autoAmpacidad', 'iccRed', 'resumenSimple', 'ocupacion', 'cargaDerivados', 'pais', 'previstas', 'fuentes'].forEach(k => { if (p[k] === undefined) p[k] = Store.emptyProject()[k]; });
      p.tableros.forEach(t => Store.completarTablero(t));
      Store.sincronizarFuentes();
      if (!p.fdiv01) {
        p.tableros.forEach(t => { t.fdivTablero = inv(t.fdivTablero); (t.circuitos || []).forEach(x => { x.fdivC = inv(x.fdivC); }); Object.keys(t.fdiv || {}).forEach(k => { t.fdiv[k] = inv(t.fdiv[k]); }); });
        p.fdiv01 = true;
      }
    },

    /* ---------- tableros y circuitos ---------- */
    /** Voltaje línea-línea y fases a partir del sistema (120/208, 277/480…) y del tipo 3F/1F. */
    voltajeDe(sistema, tipo) {
      const nums = String(sistema || '').split('/').map(Number).filter(x => x);
      const vmax = nums.length ? Math.max(...nums) : 208;
      return { voltaje: vmax, fases: tipo === '1F' ? (nums.length > 1 ? 2 : 1) : 3, hilos: tipo === '1F' ? (nums.length > 1 ? 3 : 2) : 4 };
    },
    completarTablero(t) {
      if (!t.id) t.id = U.uid();
      if (!t.tipo) t.tipo = Number(t.fases) === 3 || t.fases === undefined ? '3F' : '1F';
      if (!t.sistema) t.sistema = t.tipo === '1F' ? '120/240' : '120/208';
      if (t.voltaje === undefined || t.fases === undefined) Object.assign(t, Store.voltajeDe(t.sistema, t.tipo));
      if (t.prefijo === undefined) t.prefijo = 'TABLERO';
      if (t.fuenteId === undefined) t.fuenteId = '';   // transformador o generador del proyecto del que cuelga directamente
      if (t.reserva === undefined) t.reserva = 0.1;
      if (t.longitud === undefined) t.longitud = 10;
      t.alim = Object.assign({ material: 'CU', aislamiento: 'XHHW-2', fp: 0.9, mult: 1.25, tuberia: 'EMT', paralelos: '', aumento: 1, tempAmb: '26-30', tempBorne: 75, agrupamiento: '', breakerId: '', prot: '' }, t.alim || {});
      t.circuitos = t.circuitos || []; t.fd = t.fd || {}; t.cambiosRevit = t.cambiosRevit || [];
      // transformador aguas arriba (antes solo datos para la tabla original; primario en kV)
      t.trafo = Object.assign({ activo: false, kva: '', z: '', xr: '', primario: '', fases: '', nombre: '' }, t.trafo || {});
      if (Number(t.trafo.primario) && Number(t.trafo.primario) < 100) t.trafo.primario = Number(t.trafo.primario) * 1000;
      // versión anterior: transformador de la lista del Excel solo para el Icc → transformador aguas arriba
      if (t.transformadorId && !t.trafo.activo) {
        const x = Store.catalog.transformadores.find(o => String(o.id) === String(t.transformadorId)), m = x && /(\d+(?:[.,]\d+)?)\s*KVA/i.exec(x.nombre);
        if (x && m) Object.assign(t.trafo, { activo: true, kva: Number(m[1].replace(',', '.')), z: x.z, fases: /MONO/i.test(x.tipo) ? 1 : 3, nombre: x.nombre });
        delete t.transformadorId;
      }
      t.circuitos.forEach(c => { if (!c.id) c.id = U.uid(); if (!Array.isArray(c.polos)) c.polos = []; });
      if (t.montaje === undefined) t.montaje = 'Superficial';
      if (!t.principal) t.principal = 'interruptor';
      if (!t.clase) t.clase = 'tablero';
      if (t.espacios === undefined) t.espacios = '';
      t.ups = Object.assign({ activo: false, kva: '', factorIcc: 2, nombre: 'UPS' }, t.ups || {});
      t.alterna = Object.assign({ activo: false, equipo: 'ATS', tipo: 'generador', nombre: 'GEN-1', kva: '', xd: 12, origenId: '', fuenteId: '', longitud: '', material: 'CU', aislamiento: 'XHHW-2', tuberia: 'EMT', paralelos: '' }, t.alterna || {});
      return t;
    },
    nuevoTablero(datos) {
      const t = Store.completarTablero(Object.assign({ id: U.uid(), nombre: '', tipo: '3F', sistema: '120/208', padreId: '', conectadoA: '', circuitos: [] }, datos || {}));
      Object.assign(t, Store.voltajeDe(t.sistema, t.tipo));
      Store.project.tableros.push(t); Store.save();
      return t;
    },
    tablero(id) { return Store.project.tableros.find(t => t.id === id) || null; },
    eliminarTablero(id) {
      const p = Store.project;
      p.tableros.forEach(t => {
        if (t.padreId === id) t.padreId = '';
        t.circuitos.forEach(c => { if (c.tableroHijoId === id) c.tableroHijoId = ''; });
      });
      p.tableros = p.tableros.filter(t => t.id !== id); Store.save();
    },
    duplicarTablero(id) {
      const o = Store.tablero(id); if (!o) return null;
      const t = U.clone(o); t.id = U.uid(); t.nombre = (o.nombre || '') + ' (copia)';
      t.circuitos.forEach(c => { c.id = U.uid(); c.tableroHijoId = ''; });
      Store.project.tableros.splice(Store.project.tableros.indexOf(o) + 1, 0, t); Store.save();
      return t;
    },
    /** Siguiente posición libre en el tablero para un circuito de n polos (los polos de un mismo lado van de 2 en 2). */
    posicionLibre(t, nPolos) {
      const usados = new Set(); t.circuitos.forEach(c => (c.polos || []).forEach(p => usados.add(Number(p))));
      for (let p = 1; p < 400; p++) {
        const set = Array.from({ length: nPolos }, (_, i) => p + 2 * i);
        if (set.every(x => !usados.has(x))) return set;
      }
      return [];
    },
    nuevoCircuito(t, datos) {
      const det = Store.catalog.detallesCarga.find(d => String(d.id) === String((datos || {}).detalleId || 8)) || {};
      const nPolos = Number(det.fases) === 3 ? 3 : (Number(det.v) >= 208 && Number(det.v) <= 240 && Number(det.fases) === 2 ? 2 : 1);
      const c = Object.assign({ id: U.uid(), polos: Store.posicionLibre(t, nPolos), detalleId: det.id || 8, descripcion: '', kva: '', longitud: '', material: 'CU', aislamiento: 'THHN', mult: '', paralelos: '', aumento: 1, breakerId: '', prot: '', tableroHijoId: '' }, datos || {});
      t.circuitos.push(c); Store.save();
      return c;
    },

    /** El tablero de origen de un bypass lleva el circuito (con su protección) que alimenta la segunda acometida. */
    sincronizarBypass() {
      const P = Store.project; let cambio = false;
      P.tableros.forEach(o => {
        const antes = o.circuitos.length;
        o.circuitos = o.circuitos.filter(c => { if (!c.respaldoDe) return true; const t = Store.tablero(c.respaldoDe), a = t && t.alterna; return a && a.activo && a.tipo === 'tablero' && a.origenId === o.id; });
        if (o.circuitos.length !== antes) cambio = true;
      });
      P.tableros.forEach(t => {
        const a = t.alterna, o = a && a.activo && a.tipo === 'tablero' && Store.tablero(a.origenId);
        if (o && o.id !== t.id && !o.circuitos.some(c => c.respaldoDe === t.id)) {
          const det = Store.detalleTablero(t, o), dd = Store.catalog.detallesCarga.find(x => x.id === det) || {};
          o.circuitos.push({ id: U.uid(), polos: Store.posicionLibre(o, Number(dd.fases) === 3 ? 3 : 2), detalleId: det, descripcion: '', kva: 0, longitud: a.longitud, material: 'CU', aislamiento: 'THHN', mult: '', paralelos: '', aumento: 1, breakerId: '', prot: '', tableroHijoId: '', respaldoDe: t.id, auto: true });
          cambio = true;
        }
      });
      if (cambio) Store.save();
    },
    /** Aplica cambios de posición y los anota en la lista de cambios pendientes para Revit (un registro por circuito: posición original → nueva). */
    registrarCambios(t, cambios, origen) {
      const fecha = U.today();
      cambios.forEach(ch => {
        const c = t.circuitos.find(x => x.id === ch.id); if (!c) return;
        const de = (ch.de || c.polos || []).join(','), a = ch.a.join(',');
        c.polos = ch.a.slice();
        const e = t.cambiosRevit.find(x => x.circuitoId === ch.id && !x.aplicado);
        if (e) { e.a = a; e.fecha = fecha; e.origen = origen || e.origen; if (e.a === e.de) t.cambiosRevit.splice(t.cambiosRevit.indexOf(e), 1); }
        else if (de !== a) t.cambiosRevit.push({ id: U.uid(), circuitoId: ch.id, descripcion: ch.descripcion || c.descripcion || '', de, a, revit: c.circuitoRevit || de, fecha, origen: origen || 'manual', aplicado: false });
      });
      Store.save();
    },
    /** Detalle de carga (DCARGAS) adecuado para el circuito que alimenta a un tablero derivado. */
    detalleTablero(t, padre) {
      const d = Store.catalog.detallesCarga, f = Number(t.fases) === 3 ? 3 : 2, V = Number(padre ? padre.voltaje : t.voltaje);
      if (padre && Number(padre.voltaje) !== Number(t.voltaje)) { const tr = d.find(x => /transformador/i.test(x.descripcion) && Number(x.v) === V); if (tr) return tr.id; }
      return (d.find(x => /tablero/i.test(x.descripcion) && Number(x.fases) === f && Number(x.v) === V) ||
        d.find(x => /tablero/i.test(x.descripcion) && Number(x.v) === V) ||
        d.find(x => /tablero/i.test(x.descripcion) && Number(x.fases) === f) || d[0]).id;
    },
    /** Cambia de quién se alimenta un tablero y crea/mueve el circuito que lo alimenta en el tablero padre. */
    cambiarPadre(t, padreId) {
      const p = Store.project;
      p.tableros.forEach(o => { o.circuitos = o.circuitos.filter(c => !(c.tableroHijoId === t.id && c.auto)); o.circuitos.forEach(c => { if (c.tableroHijoId === t.id) c.tableroHijoId = ''; }); });
      t.padreId = padreId || ''; if (padreId) t.fuenteId = '';
      const padre = Store.tablero(padreId);
      if (padre) {
        const det = Store.detalleTablero(t, padre), dd = Store.catalog.detallesCarga.find(x => x.id === det) || {};
        const np = Number(dd.fases) === 3 ? 3 : 2;
        padre.circuitos.push({ id: U.uid(), polos: Store.posicionLibre(padre, np), detalleId: det, descripcion: '', kva: 0, longitud: t.longitud, material: 'CU', aislamiento: 'THHN', mult: '', paralelos: '', aumento: 1, breakerId: '', prot: '', tableroHijoId: t.id, auto: true });
      }
      Store.save();
    },
    /* ---------- fuentes del proyecto: transformadores y generadores compartidos por varios tableros / ATS ---------- */
    fuente(id) { return id ? (Store.project.fuentes || []).find(x => x.id === id) || null : null; },
    nuevaFuente(tipo) {
      const gen = tipo === 'generador', p = Store.project, n = p.fuentes.filter(x => x.tipo === (gen ? 'generador' : 'transformador')).length + 1;
      const x = gen ? { id: U.uid(), tipo, nombre: 'GEN-' + n, kva: '', sistema: '120/208', xd: 12 }
        : { id: U.uid(), tipo: 'transformador', nombre: 'T-' + n, kva: '', z: '', xr: '', primario: '', fases: 3, sistema: '120/208' };
      p.fuentes.push(x); Store.save(); return x;
    },
    eliminarFuente(id) {
      const p = Store.project; p.fuentes = p.fuentes.filter(x => x.id !== id);
      p.tableros.forEach(t => { if (t.fuenteId === id) t.fuenteId = ''; if (t.alterna && t.alterna.fuenteId === id) t.alterna.fuenteId = ''; });
      Store.save();
    },
    /** Los tableros que cuelgan de una fuente toman el sistema (voltaje del secundario) de ella. */
    sincronizarFuentes() {
      const p = Store.project;
      p.tableros.forEach(t => {
        if (t.alterna && t.alterna.fuenteId && !Store.fuente(t.alterna.fuenteId)) t.alterna.fuenteId = '';
        if (!t.fuenteId) return;
        const f = Store.fuente(t.fuenteId);
        if (!f) { t.fuenteId = ''; return; }
        if (f.sistema && t.sistema !== f.sistema) { t.sistema = f.sistema; Object.assign(t, Store.voltajeDe(t.sistema, t.tipo)); }
      });
    },
    /** Cuelga el tablero directamente de una fuente (transformador o generador) o lo suelta. */
    asignarFuente(t, id) {
      if (id) { Store.cambiarPadre(t, ''); t.fuenteId = id; } else t.fuenteId = '';
      Store.sincronizarFuentes(); Store.save();
    },
    /** Opciones de "Alimentado desde": tableros y fuentes del proyecto (valor 'F:' + id para las fuentes). */
    opcionesAlimentacion(t) {
      const tabs = Store.posiblesPadres(t).map(o => ({ value: o.id, label: Calc.nombreTablero(o), group: 'Tableros' }));
      const fts = (Store.project.fuentes || []).map(f => ({ value: 'F:' + f.id, label: (f.tipo === 'generador' ? 'Generador ' : 'Transformador ') + (f.nombre || '') + (Number(f.kva) ? ' · ' + f.kva + ' kVA' : '') + (f.sistema ? ' · ' + f.sistema + ' V' : ''), group: 'Fuentes del proyecto' }));
      return fts.concat(tabs);
    },
    valorAlimentacion(t) { return t.fuenteId ? 'F:' + t.fuenteId : (t.padreId || ''); },
    alimentarDesde(t, v) {
      if (String(v).startsWith('F:')) Store.asignarFuente(t, String(v).slice(2));
      else { t.fuenteId = ''; Store.cambiarPadre(t, v); }
    },

    /** Tableros que pueden alimentar a t (no él mismo ni sus derivados). */
    posiblesPadres(t) {
      const p = Store.project, desc = new Set([t.id]);
      let cambio = true;
      while (cambio) { cambio = false; p.tableros.forEach(o => { if (o.padreId && desc.has(o.padreId) && !desc.has(o.id)) { desc.add(o.id); cambio = true; } }); }
      return p.tableros.filter(o => !desc.has(o.id));
    },

    /* ---------- persistencia ---------- */
    save() {
      clearTimeout(saveTimer);
      Store.status = 'saving'; Store._emit();
      saveTimer = setTimeout(Store.flush, 250);
    },
    flush() {
      clearTimeout(saveTimer);
      try {
        const p = Store.project;
        // un proyecto sin datos no se guarda en la lista (así la herramienta siempre arranca limpia)
        if (p.nombre || p.numero || (p.tableros || []).length || (p.previstas || []).length) {
          p.modificado = new Date().toISOString();
          const l = Store.proyectos(); l[p.id] = p; localStorage.setItem(K_PROYECTOS, JSON.stringify(l));
          try { sessionStorage.setItem(K_ACTUAL, p.id); } catch (e) { /* ignora */ }
        }
        localStorage.setItem(K_CATALOG, JSON.stringify(Store.catalog));
        Store.status = 'ok';
      } catch (e) { Store.status = 'error'; }
      Store._emit();
    },
    exportProject() { return JSON.stringify({ tipo: 'anacar-proyecto', version: 1, proyecto: Store.project }, null, 1); },
    importProject(text) {
      const o = JSON.parse(text), p = o && (o.proyecto || o);
      if (!p || !Array.isArray(p.tableros)) throw new Error('El archivo no es un proyecto de esta herramienta.');
      Store.project = p; Store.migrate(); Store.save();
    },
    newProject() { Store.flush(); Store.project = Store.emptyProject(); try { sessionStorage.removeItem(K_ACTUAL); } catch (e) { /* ignora */ } Store._emit(); },
    exportCatalog() { return JSON.stringify({ tipo: 'anacar-catalogo', version: 1, catalogo: Store.catalog }, null, 1); },
    importCatalog(text) {
      const o = JSON.parse(text), c = o && (o.catalogo || o);
      if (!c || !Array.isArray(c.detallesCarga) || !Array.isArray(c.breakers)) throw new Error('El archivo no es un catálogo de esta herramienta.');
      Store.catalog = c; Store.migrate(); Store.save();
    },
    resetCatalog() { Store.catalog = Store.seedCatalog(); Store.save(); },
    adminHash() { try { return localStorage.getItem(K_ADMINPW) || window.CONFIG.ADMIN_PASSWORD_SHA256; } catch (e) { return window.CONFIG.ADMIN_PASSWORD_SHA256; } },
    setAdminHash(h) { localStorage.setItem(K_ADMINPW, h); },
  };
  window.addEventListener('beforeunload', () => { if (Store.status === 'saving') Store.flush(); });
  g.Store = Store;
})(window);
