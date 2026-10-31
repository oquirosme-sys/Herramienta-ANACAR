/* Importar un proyecto de ANACAR en Excel (hoja ANACAR con un bloque «Machote» por tablero) a la herramienta en línea.
   Se copian las entradas (cargas, longitudes, posiciones, protecciones, calibres, breakers, tablero de catálogo, jerarquía, ATS/generador)
   y se recalcula en la herramienta; después se compara contra los valores que Excel dejó calculados y se informan las diferencias.
   Cada bloque mide 143 filas: la fila «ID» del primero es la 9 (rel. 9); los circuitos van de la rel. 16 a la 115 y el alimentador está en la rel. 139.
   Una celda con fórmula en Excel es un dato calculado; una celda sin fórmula es una entrada del ingeniero (se respeta). */
(function (g) {
  'use strict';
  const PASO = 143, R0 = 9;
  const ci = c => { let n = 0; for (const ch of c) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  const num = v => (v === '' || v === null || v === undefined || isNaN(Number(v)) ? 0 : Number(v));
  const txt = v => (v === null || v === undefined ? '' : String(v).trim());
  const redondea = (v, d) => { const k = Math.pow(10, d || 2); return Math.round(num(v) * k) / k; };
  /** En el Excel el factor de diversidad DIVIDE la demanda; en la herramienta MULTIPLICA y va de 0 a 1. Equivale 1/S cuando S ≥ 1; con S < 1 (aumenta la demanda) la herramienta no lo admite. */
  const diversidad = s => { const x = num(s); return !x || x === 1 ? { web: '', ok: true } : x > 1 ? { web: Math.round(100000 / x) / 100000, ok: true } : { web: '', ok: false }; };
  const nombreTab = t => ((t.prefijo === undefined ? 'TABLERO' : t.prefijo) + ' ' + (t.nombre || '')).trim();

  /** Lee los bloques de la hoja ANACAR: cada uno con helpers de lectura relativos a la fila 9 del Machote. */
  function bloques(hoja) {
    const out = [];
    for (let r = R0 - 1; r < hoja.rows.length; r++) {
      const f = hoja.rows[r] || [];
      if (txt(f[1]) === 'ID' && /Voltaje Nominal/i.test(txt(f[9]))) {
        const off = r + 1 - R0, k = out.length;
        const fila = rel => hoja.rows[rel + off - 1] || [];
        const b = {
          k, off,
          v: (rel, c) => fila(rel)[ci(c)],
          f: (rel, c) => !!(hoja.formulas[rel + off - 1] || [])[ci(c)],
        };
        b.nombre = txt(b.v(10, 'I')); b.prefijo = txt(b.v(10, 'B'));
        out.push(b);
      }
    }
    return out;
  }

  /** Marca de un breaker de la hoja (modelo de referencia, unidad, polos, kA) en el formato de la herramienta. */
  function breakerExcel(b, rel) {
    return { modelo: txt(b.v(rel, 'BA')), marco: num(b.v(rel, 'BB')), amperios: txt(b.v(rel, 'BC')), unidad: txt(b.v(rel, 'BD')) || 'STD', polos: num(b.v(rel, 'BE')), sccr: num(b.v(rel, 'BF')) };
  }

  /** Construye el borrador del proyecto con las entradas del Excel. */
  function construir(sheets, nombreArchivo) {
    const hoja = sheets.find(s => /^ANACAR$/i.test(s.name));
    if (!hoja) throw new Error('El archivo no tiene la hoja «ANACAR». ¿Es un proyecto de ANACAR en Excel?');
    const todos = bloques(hoja).filter(b => b.nombre && !/^machote$/i.test(b.nombre));
    if (!todos.length) throw new Error('No se encontraron tableros en la hoja ANACAR.');
    const esTab = b => /^TABLERO$/i.test(b.prefijo);
    const porNombre = {}; todos.forEach(b => { porNombre[b.nombre.toUpperCase()] = b; });
    // número de circuitos de un bloque (fila con detalle de carga y posición)
    const nC = b => { let n = 0; for (let rel = 16; rel <= 115; rel++) if (num(b.v(rel, 'K')) > 0 && ['D', 'F', 'H'].some(c => num(b.v(rel, c)) > 0)) n++; return n; };
    // transferencia con generador: GENERADOR → ATS/MTS ← IP(s) ← fuente; son bloques de paso (sin circuitos propios o con uno solo hacia el siguiente)
    const cadenas = [], equipoSet = new Set();
    todos.filter(g => /^GEN/i.test(g.nombre) && nC(g) <= 1).forEach(g => {
      const a = porNombre[txt(g.v(10, 'P')).toUpperCase()];
      if (!a || !/^(ATS|MTS)/i.test(a.nombre) || nC(a) > 1) return;
      const ips = []; let e = porNombre[txt(a.v(10, 'P')).toUpperCase()];
      while (e && /^IP/i.test(e.nombre) && nC(e) <= 1 && ips.length < 4) { ips.push(e); e = porNombre[txt(e.v(10, 'P')).toUpperCase()]; }
      cadenas.push({ g, a, ips, arriba: e && !equipoSet.has(e) ? e : null, arribaTxt: txt((ips.length ? ips[ips.length - 1] : a).v(10, 'P')) });
      [g, a].concat(ips).forEach(x => equipoSet.add(x));
    });
    const equipos = todos.filter(b => equipoSet.has(b) || !esTab(b));
    const tabs = todos.filter(b => !equipos.includes(b));
    const aviso = [];
    const mapa = {};   // nombre del bloque → tablero de la herramienta

    tabs.forEach(b => {
      const sist = txt(b.v(10, 'V')) || '120/208', fases = num(b.v(10, 'L')) || 3;
      const t = {
        id: U.uid(), prefijo: 'TABLERO', nombre: b.nombre, tipo: fases === 3 ? '3F' : '1F', sistema: sist, padreId: '', conectadoA: txt(b.v(10, 'P')), circuitos: [],
        voltaje: num(b.v(10, 'J')) || 208, fases, hilos: num(b.v(10, 'M')) || (fases === 3 ? 4 : 3), reserva: num(b.v(10, 'O')), longitud: num(b.v(10, 'R')),
        montaje: txt(b.v(11, 'BP')) || 'Superficial', marca: txt(b.v(11, 'BK')) || undefined,
      };
      if (txt(b.v(10, 'T')) !== '') t.iccManual = num(b.v(10, 'T'));
      t.alim = {
        material: txt(b.v(139, 'AG')) || 'CU', aislamiento: txt(b.v(139, 'AH')) || 'XHHW-2', paralelos: b.f(139, 'AI') ? '' : num(b.v(139, 'AI')) || '', aumento: num(b.v(139, 'AJ')) || 1,
        mult: num(b.v(139, 'AD')) || 1.25, fp: num(b.v(139, 'T')) || 0.9, tuberia: txt(b.v(140, 'AQ')) || 'EMT', tempAmb: txt(b.v(10, 'Y')) || '26-30', tempBorne: num(b.v(10, 'Z')) || 75, agrupamiento: txt(b.v(10, 'AB')),
      };
      if (!b.f(139, 'R') && num(b.v(139, 'R')) > 0) t.fdTablero = num(b.v(139, 'R'));   // factor de demanda del tablero escrito a mano en el Excel
      if (!b.f(139, 'S') && num(b.v(139, 'S')) && num(b.v(139, 'S')) !== 1) {   // diversidad del tablero: según la versión del Excel multiplica o divide; se deduce de los valores calculados
        const S = num(b.v(139, 'S')), J = num(b.v(139, 'J')), Rr = num(b.v(139, 'R')), L = num(b.v(139, 'L')), rho = J * Rr ? L / (J * Rr) : null;
        if (rho !== null && Math.abs(rho - S) <= 0.01 * Math.max(1, S)) { if (S <= 1) t.fdivTablero = S; else aviso.push('«' + nombreTab(t) + '»: factor de diversidad del tablero ' + S + ' (multiplica y aumenta la demanda); la herramienta solo admite de 0 a 1, se dejó en 1.'); }
        else if (rho !== null && Math.abs(rho - 1 / S) <= 0.01) { const dv = diversidad(S); if (dv.ok) t.fdivTablero = dv.web; else aviso.push('«' + nombreTab(t) + '»: factor de diversidad del tablero ' + S + ' (divide y aumenta la demanda); la herramienta solo admite de 0 a 1, se dejó en 1.'); }
        else aviso.push('«' + nombreTab(t) + '»: no se pudo deducir cómo aplica el Excel el factor de diversidad del tablero (' + S + '); se dejó en 1.');
      }
      t._bloque = b; t._cat = { fab: txt(b.v(11, 'BK')), modelo: txt(b.v(11, 'BJ')), barraFase: num(b.v(11, 'BL')), barraNeutro: num(b.v(11, 'BM')), barraTierra: num(b.v(11, 'BN')), espacios: num(b.v(11, 'BO')), idExcel: num(b.v(11, 'BI')) };
      t._spdId = num(b.v(16, 'BI'));
      // circuitos (rel. 16 a 115): el detalle de carga (K) marca una fila usada
      for (let rel = 16; rel <= 115; rel++) {
        const det = num(b.v(rel, 'K')); if (det <= 0) continue;
        const polos = ['D', 'F', 'H'].map(c => num(b.v(rel, c))).filter(x => x > 0);
        if (!polos.length) continue;
        const c = { id: U.uid(), polos, detalleId: det, descripcion: '', kva: 0, longitud: num(b.v(rel, 'M')), material: txt(b.v(rel, 'AG')) || 'CU', aislamiento: txt(b.v(rel, 'AH')) || 'THHN',
          mult: num(b.v(rel, 'AD')) || '', paralelos: '', aumento: num(b.v(rel, 'AJ')) || 1, breakerId: '', prot: '', tableroHijoId: '', _rel: rel };
        const desc = txt(b.v(rel, 'I'));
        c._descExcel = desc;
        if (!b.f(rel, 'I')) c.descripcion = desc;
        c.kva = num(b.v(rel, 'J')); c._kvaF = b.f(rel, 'J');
        if (!b.f(rel, 'AF') && num(b.v(rel, 'AF'))) c.prot = num(b.v(rel, 'AF'));
        if (!b.f(rel, 'AI') && num(b.v(rel, 'AI')) > 0) c.paralelos = num(b.v(rel, 'AI'));
        if (!b.f(rel, 'R') && txt(b.v(rel, 'R')) !== '') c.fu = num(b.v(rel, 'R'));
        if (!b.f(rel, 'S') && txt(b.v(rel, 'S')) !== '' && num(b.v(rel, 'S')) !== 1) { const dv = diversidad(b.v(rel, 'S')); if (dv.ok) c.fdivC = dv.web; }
        t.circuitos.push(c);
      }
      t.cambiosRevit = [];
      // factores de demanda y diversidad por tipo de carga (filas 131 y 132): solo los escritos a mano que difieren del catálogo
      const nm2 = s => String(s || '').toUpperCase().replace(/[^A-Z0-9Á-Ú]/g, '');
      'JKLMNOPQRSTUV'.split('').forEach(col => {
        const tp = Store.catalog.tiposCarga.find(x => nm2(x.nombre) === nm2(b.v(121, col))); if (!tp) return;
        if (!b.f(131, col) && txt(b.v(131, col)) !== '' && Math.abs(num(b.v(131, col)) - num(tp.fd)) > 1e-9) (t.fd = t.fd || {})[tp.id] = num(b.v(131, col));
        if (!b.f(132, col) && txt(b.v(132, col)) !== '' && num(b.v(132, col)) !== 1) {
          const dv = diversidad(b.v(132, col));
          if (dv.ok) (t.fdiv = t.fdiv || {})[tp.id] = dv.web;
          else aviso.push('«' + nombreTab(t) + '»: factor de diversidad ' + num(b.v(132, col)) + ' en ' + tp.nombre + '. En el Excel divide la demanda (la aumenta ' + Math.round((1 / num(b.v(132, col)) - 1) * 100) + ' %); la herramienta solo admite diversidad de 0 a 1 que la reduce, así que se dejó en 1 y la demanda de ese tipo saldrá menor que en el Excel.');
        }
      });
      mapa[b.nombre.toUpperCase()] = t;
    });

    // jerarquía: un circuito cuya descripción es «TABLERO <nombre>» alimenta a ese tablero; si no, vale «Conectado a»
    tabs.forEach(b => {
      const t = mapa[b.nombre.toUpperCase()];
      t.circuitos.forEach(c => {
        const m = /^TABLERO\s+(.+)$/i.exec(c._descExcel), hijo = m && mapa[m[1].trim().toUpperCase()];
        if (hijo && hijo !== t) { c.tableroHijoId = hijo.id; c.descripcion = ''; c.kva = 0; hijo.padreId = t.id; }
      });
    });
    tabs.forEach(b => {
      const t = mapa[b.nombre.toUpperCase()];
      const p = mapa[t.conectadoA.toUpperCase()];
      if (p && p !== t && !t.padreId) { t.padreId = p.id; aviso.push('«' + nombreTab(t) + '» cuelga de «' + nombreTab(p) + '» por «Conectado a», pero ese tablero no tiene un circuito para él: la herramienta le crea uno.'); }
    });

    // transferencia con generador: el tablero que cuelga del ATS lleva la segunda acometida y pasa a colgar de lo que alimentaba la cadena
    cadenas.forEach(ch => {
      const bajo = tabs.filter(b => txt(b.v(10, 'P')).toUpperCase() === ch.a.nombre.toUpperCase());
      const distNormal = [ch.a].concat(ch.ips).reduce((x, e) => x + num(e.v(10, 'R')), 0);
      const primero = (ch.ips.length ? ch.ips[ch.ips.length - 1] : ch.a).nombre;
      bajo.forEach(b => {
        const t = mapa[b.nombre.toUpperCase()], up = ch.arriba && esTab(ch.arriba) ? mapa[ch.arriba.nombre.toUpperCase()] : null;
        t.alterna = { activo: true, equipo: /^MTS/i.test(ch.a.nombre) ? 'MTS' : 'ATS', tipo: 'generador', nombre: ch.g.nombre, kva: '', xd: 12, origenId: '', fuenteId: '', longitud: num(ch.g.v(10, 'R')), distNormal, material: 'CU', aislamiento: 'XHHW-2', tuberia: 'EMT', paralelos: '' };
        if (up) {
          const clave = x => String(x || '').toUpperCase().split(' ').filter(Boolean).join(' '), c = up.circuitos.find(x => clave(x._descExcel) === clave('TABLERO ' + primero));
          if (c) { c.tableroHijoId = t.id; c.descripcion = ''; c.kva = 0; t.padreId = up.id; t.conectadoA = ''; }
          else { t.padreId = up.id; t.conectadoA = ''; aviso.push('«' + nombreTab(up) + '» no tiene el circuito hacia «' + primero + '»: la herramienta crea uno para «' + nombreTab(t) + '».'); }
        } else t.conectadoA = ch.arribaTxt || t.conectadoA;
        aviso.push('«' + nombreTab(t) + '»: ' + [ch.a].concat(ch.ips).map(x => x.nombre).join(' ← ') + ' y generador «' + ch.g.nombre + '» pasaron a la segunda acometida (' + t.alterna.equipo + '). Revise la capacidad del generador (kVA) y la distancia normal → ' + t.alterna.equipo + ' (' + distNormal + ' m, suma de ' + [ch.a].concat(ch.ips).map(x => x.nombre).join(' + ') + ').');
      });
    });
    // bloques que no son tableros y no forman parte de una transferencia con generador: queda el texto en «Conectado a»
    equipos.filter(e => !equipoSet.has(e)).forEach(e => aviso.push('«' + e.nombre + '» no se importó como tablero (' + (esTab(e) ? 'bloque de equipo' : 'sin prefijo TABLERO') + '); revise si es un equipo que la herramienta no modela.'));

    // orden y exclusión de las tablas resumen (hoja ORDEN TABLEROS)
    const ord = sheets.find(s => /^ORDEN TABLEROS$/i.test(s.name));
    if (ord) ord.rows.forEach(r => {
      const nm = txt(r[1]).replace(/^TABLERO\s+/i, '').toUpperCase(), t = mapa[nm]; if (!t) return;
      const o = r[2]; if (/^x$/i.test(txt(o))) t.excluirResumen = true; else if (num(o)) t.orden = num(o);
    });

    const proy = Store.emptyProject();
    proy.nombre = (nombreArchivo || '').replace(/\.(xlsm|xlsx)$/i, '').replace(/^ANACAR[\s\d.]*-?\s*/i, '').trim() || 'Proyecto importado de Excel';
    proy.tableros = tabs.map(b => mapa[b.nombre.toUpperCase()]);
    return { proy, tabs: tabs.map(b => mapa[b.nombre.toUpperCase()]), mapa, aviso, equipos };
  }

  /** Detalles de carga: cada Excel trae su propia hoja DCARGAS y los números pueden no coincidir con los de la herramienta.
      Se busca por descripción, voltaje y fases; si el tipo de carga o los factores difieren se actualiza el catálogo con el valor del Excel (queda en el reporte); si no existe se agrega. */
  function conciliarDetalles(sheets, res) {
    const hoja = sheets.find(s => /^DCARGAS$/i.test(s.name)), nota = [];
    if (!hoja) return nota;
    const ex = {}; hoja.rows.forEach(r => { if (typeof r[1] === 'number' && r[1] > 0 && txt(r[3])) ex[r[1]] = { tipo: num(r[2]), desc: txt(r[3]), v: num(r[4]), fases: num(r[5]), hilos: num(r[6]), fd: num(r[7]), fdiv: num(r[8]), fp: num(r[9]) }; });
    const C = Store.catalog, norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();
    const memo = {}, tocados = [], nuevos = [];
    res.tabs.forEach(t => t.circuitos.forEach(c => {
      const id = c.detalleId, e = ex[id]; if (!e) return;
      if (memo[id] === undefined) {
        // mismo número y descripción → es el mismo detalle (se actualiza si el Excel cambió sus datos); si no, mismo texto, voltaje, fases y tipo; si no, se agrega
        let d = C.detallesCarga.find(x => x.id === id && norm(x.descripcion) === norm(e.desc)) || C.detallesCarga.find(x => norm(x.descripcion) === norm(e.desc) && Number(x.v) === e.v && Number(x.fases) === e.fases && x.tipo === e.tipo);
        if (!d) {
          d = { id: Math.max(0, ...C.detallesCarga.map(x => x.id)) + 1, tipo: e.tipo, descripcion: e.desc, v: e.v, fases: e.fases, hilos: e.hilos, fd: e.fd, fdiv: diversidad(e.fdiv).web || 1, fp: e.fp, continua: false, mult: 1.25, fundamento: 'Importado de un Excel de ANACAR', claves: '' };
          C.detallesCarga.push(d); nuevos.push(e.desc);
        } else if (d.tipo !== e.tipo || Number(d.hilos) !== e.hilos || Math.abs(num(d.fd) - e.fd) > 1e-9 || Math.abs((num(d.fdiv) || 1) - (diversidad(e.fdiv).web || 1)) > 1e-9 || Math.abs(num(d.fp) - e.fp) > 1e-9) {
          tocados.push(d.descripcion + ': tipo ' + d.tipo + ' → ' + e.tipo);
          Object.assign(d, { tipo: e.tipo, hilos: e.hilos, fd: e.fd, fdiv: diversidad(e.fdiv).web || 1, fp: e.fp });
        }
        memo[id] = d.id;
      }
      c.detalleId = memo[id];
    }));
    if (tocados.length) nota.push('Detalles de carga del catálogo actualizados con los del Excel (' + tocados.length + '): ' + tocados.join('; ') + '.');
    if (nuevos.length) nota.push('Detalles de carga que no existían y se agregaron al catálogo: ' + nuevos.join(', ') + '.');
    return nota;
  }

  /** Tablero de catálogo y supresor: se busca el mismo modelo en el catálogo de la herramienta; si no existe se escribe a mano con los datos del Excel. */
  function conciliarCatalogos(res) {
    const C = Store.catalog, nota = [];
    res.tabs.forEach(t => {
      const c = t._cat, nm = nombreTab(t);
      const hall = C.tablerosCat.find(x => x.fabricante === c.fab && x.modelo === c.modelo && Number(x.barraFase) === c.barraFase && Number(x.espacios) === c.espacios) || C.tablerosCat.find(x => x.fabricante === c.fab && x.modelo === c.modelo && Number(x.barraFase) === c.barraFase);
      if (!c.modelo) return;
      if (hall) { t.catalogoId = hall.id; t.marca = hall.fabricante; if (!t.espacios && Number(hall.espacios) !== c.espacios) t.espacios = c.espacios; }
      else { t.catalogoId = 'manual'; t.tableroManual = { modelo: c.modelo, barraFase: c.barraFase, barraNeutro: c.barraNeutro, barraTierra: c.barraTierra, espacios: c.espacios }; nota.push('«' + nm + '»: el tablero ' + c.modelo + ' (' + c.fab + ' ' + c.barraFase + ' A, ' + c.espacios + ' espacios) no está en el catálogo; quedó como tablero escrito a mano.'); }
      const sp = c._spdId && C.supresores.find(x => String(x.id) === String(c._spdId));
      if (sp) {
        const modeloExcel = txt(t._bloque.v(11, 'BJ'));
        void modeloExcel;
      }
    });
    return nota;
  }

  /** Compara lo calculado por la herramienta con lo que Excel dejó calculado. Devuelve filas de diferencias. */
  function comparar(res, R) {
    const dif = [], tol = { kva: 0.02, cv: 0.05, amp: 0.5 };
    res.tabs.forEach(t => {
      const b = t._bloque, r = R.res[t.id], nm = nombreTab(t); if (!r) return;
      const add = (que, excel, web, ok, tipo, nota) => { if (!ok) dif.push({ tablero: nm, que, excel, web, tipo: tipo || 'calculo', nota: nota || '' }); };
      const eW130 = num(b.v(130, 'W')), eL139 = num(b.v(139, 'L')), eAF = num(b.v(139, 'AF')), eAL = txt(b.v(139, 'AL')), eAY = num(b.v(139, 'AY')), eDes = num(b.v(118, 'U'));
      add('kVA conectados totales', redondea(eW130, 2), redondea(r.W130, 2), Math.abs(eW130 - r.W130) <= Math.max(tol.kva, eW130 * 0.002));
      add('kVA demandados', redondea(eL139, 2), redondea(r.alim.L139, 2), Math.abs(eL139 - r.alim.L139) <= Math.max(tol.kva, eL139 * 0.002));
      // el Excel de algunos proyectos calcula la ampacidad requerida con la corriente de la fase A y no con la más cargada
      const eAA = ['AA', 'AB', 'AC'].map(c => num(b.v(139, c))), eAE = num(b.v(139, 'AE')), eAD = num(b.v(139, 'AD')) || 1;
      const soloA = eAE > 0 && Math.abs(eAE - eAA[0] * eAD) < 0.5 && eAE < Math.max.apply(null, eAA) * eAD - 0.5;
      const notaA = soloA ? 'El Excel usa la corriente de la fase A (' + redondea(eAA[0], 0) + ' A) para la ampacidad requerida; la herramienta usa la fase más cargada (' + redondea(Math.max.apply(null, eAA), 0) + ' A).' : '';
      add('Protección del alimentador (A)', eAF || '—', r.alim.AF139 || '—', !eAF || eAF === r.alim.AF139, 'proteccion', notaA);
      add('Calibre del alimentador', eAL || '—', r.alim.AL139 || '—', !eAL || String(eAL) === String(r.alim.AL139), 'cable', notaA);
      add('Caída de voltaje acumulada en el alimentador (%)', redondea(eAY, 2), redondea(r.alim.AY139, 2), !eAY || Math.abs(eAY - r.alim.AY139) <= 0.15);
      add('Desbalance máximo (%)', redondea(eDes, 1), redondea(r.desbalance, 1), Math.abs(eDes - r.desbalance) <= 0.2);
      // circuitos: se emparejan por la primera posición
      t.circuitos.forEach(c => {
        const x = r.rows.find(z => z.c === c); if (!x) return;
        const rel = c._rel, pos = '[' + c.polos.join(',') + '] ' + (x.descripcion || c._descExcel || '');
        const eP = num(b.v(rel, 'AF')), eC = txt(b.v(rel, 'AL')), eY = num(b.v(rel, 'AY'));
        if (c.tableroHijoId) return;
        if (eP && x.AF && eP !== x.AF) dif.push({ tablero: nm, que: 'Circuito ' + pos + ': protección (A)', excel: eP, web: x.AF, tipo: 'proteccion' });
        if (eC && x.AL && String(eC) !== String(x.AL)) dif.push({ tablero: nm, que: 'Circuito ' + pos + ': calibre', excel: eC, web: x.AL, tipo: 'cable' });
        const acumC = r.padre ? r.padre.alim.AY139 : 0;
        if (eY && x.AY !== null && Math.abs(eY - (x.AY - acumC)) > 0.15 && Math.abs(eY - x.AY) > 0.15) dif.push({ tablero: nm, que: 'Circuito ' + pos + ': caída de voltaje (%, sin la acumulada de arriba)', excel: redondea(eY, 2), web: redondea(x.AY - acumC, 2), tipo: 'calculo' });
      });
    });
    return dif;
  }

  /** Importa el archivo: crea el proyecto en la herramienta y devuelve el reporte. */
  async function importar(buf, nombreArchivo) {
    const sheets = await XlsxRead.read(buf);
    const res = construir(sheets, nombreArchivo);
    Store.newProject();
    const P = Store.project; P.nombre = res.proy.nombre;
    res.tabs.forEach(t => { Store.completarTablero(t); P.tableros.push(t); });
    const nota = conciliarDetalles(sheets, res).concat(conciliarCatalogos(res));
    // los breakers escritos en el Excel se concilian con el catálogo después del primer cálculo
    Store.sincronizarBypass && Store.sincronizarBypass();
    let R = Calc.proyecto(P, Store.catalog);
    conciliarAlimentadores(res, R, nota, 'prot');
    R = Calc.proyecto(P, Store.catalog);
    conciliarAlimentadores(res, R, nota, 'cable');
    R = Calc.proyecto(P, Store.catalog);
    conciliarBreakers(res, R, nota);
    R = Calc.proyecto(P, Store.catalog);
    const dif = comparar(res, R);
    const noCumple = R.orden.filter(r => r.alim.val && !r.alim.val.ok).map(r => r.tab.nombre);
    if (noCumple.length) nota.push('Según la herramienta, ' + noCumple.length + ' alimentador(es) del Excel no cumplen la ampacidad corregida por temperatura y agrupamiento (310.15): ' + noCumple.join(', ') + '. Se conservó el cableado del Excel; revise el aviso en la memoria de cada tablero.');
    res.tabs.forEach(t => { delete t._bloque; delete t._cat; delete t._spdId; t.circuitos.forEach(c => { delete c._rel; delete c._descExcel; delete c._kvaF; }); });
    Store.save();
    return { tableros: res.tabs.length, circuitos: res.tabs.reduce((a, t) => a + t.circuitos.length, 0), equipos: res.equipos.map(e => e.nombre), aviso: res.aviso.concat(nota), dif };
  }

  /** Alimentador de cada tablero: se respeta lo que dejó el Excel (protección y calibres). Si la herramienta calcula otra cosa se fija el valor del Excel;
      la herramienta lo conserva y avisa si no cumple la ampacidad corregida. */
  function conciliarAlimentadores(res, R, nota, que) {
    const cambios = [];
    res.tabs.forEach(t => {
      const b = t._bloque, r = R.res[t.id]; if (!r) return;
      if (que === 'prot') {
        const eAF = num(b.v(139, 'AF'));
        if (eAF > 0 && r.alim.AF139 !== eAF) { t.alim.prot = eAF; cambios.push(nombreTab(t) + ' ' + eAF + ' A'); }
      } else {
        const eF = txt(b.v(139, 'AL')), eN = txt(b.v(139, 'AN')), eT = txt(b.v(139, 'AP'));
        let c = 0;
        if (eF && String(r.alim.AL139) !== eF) { t.alim.calFases = eF; c++; }
        if (eN && eN !== ' ' && String(r.alim.AN139) !== eN) { t.alim.calNeutro = eN; c++; }
        if (eT && eT !== ' ' && String(r.alim.AP139) !== eT) { t.alim.calTierra = eT; c++; }
        if (c) cambios.push(nombreTab(t) + ' calibre ' + eF);
      }
    });
    if (cambios.length) nota.push((que === 'prot' ? 'Protección del alimentador fijada con el valor del Excel en ' : 'Calibres del alimentador fijados con los del Excel en ') + cambios.length + ' tablero(s): ' + cambios.join('; ') + '. La herramienta los respeta y avisa en la memoria si no cumplen la ampacidad corregida (310.15).');
  }

  /** Interruptores: se compara el que eligió Excel con el de la herramienta (polos, unidad, kA); si no coinciden se fija el del Excel (del catálogo o escrito a mano). */
  function conciliarBreakers(res, R, nota) {
    const C = Store.catalog;
    const igual = (bk, e) => bk && Number(bk.polos) === e.polos && (bk.unidad || 'STD') === (e.unidad || 'STD') && Number(bk.sccr) === e.sccr;
    const buscar = (marca, e) => C.breakers.find(b => b.marca === marca && String(b.modelo || '') === e.modelo && Number(b.polos) === e.polos && (b.unidad || 'STD') === e.unidad && String(b.amperios) === e.amperios && Number(b.sccr) === e.sccr)
      || C.breakers.find(b => b.marca === marca && Number(b.polos) === e.polos && (b.unidad || 'STD') === e.unidad && Number(b.sccr) === e.sccr && String(b.modelo || '') === e.modelo);
    res.tabs.forEach(t => {
      const b = t._bloque, r = R.res[t.id]; if (!r) return;
      const marca = r.marca, nm = nombreTab(t);
      // principal
      const eP = breakerExcel(b, 139);
      if (r.bkMain && eP.modelo && !igual(r.bkMain, eP)) {
        const hall = buscar(marca, eP);
        if (hall) t.alim.breakerId = hall.id;
        else { t.alim.breakerId = 'manual'; t.alim.breakerManual = { modelo: eP.modelo, amperios: num(String(eP.amperios).split('-').pop()) || '', sccr: eP.sccr, unidad: eP.unidad === 'STD' ? '' : eP.unidad }; nota.push('«' + nm + '»: interruptor principal ' + eP.modelo + ' (' + eP.sccr + ' kA) escrito a mano porque no coincide con el catálogo.'); }
      }
      // ramales
      t.circuitos.forEach(c => {
        const x = r.rows.find(z => z.c === c); if (!x || !x.breaker || c.tableroHijoId) return;
        const e = breakerExcel(b, c._rel); if (!e.modelo && !e.sccr) return;
        if (igual(x.breaker, e)) return;
        const hall = buscar(marca, e);
        if (hall) c.breakerId = hall.id;
        else { c.breakerId = 'manual'; c.breakerManual = { modelo: e.modelo, sccr: e.sccr, unidad: e.unidad === 'STD' ? '' : e.unidad }; }
      });
    });
  }

  g.ImportAnacar = { importar, construir, bloques, comparar };
})(window);
