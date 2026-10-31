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

  /** Lee los bloques de la hoja ANACAR: cada uno con helpers de lectura relativos a la fila 9 del Machote.
      Hay dos distribuciones: la actual (bloques de 143 filas, un circuito por polo) y la antigua (bloques de largo variable, un circuito por fila). */
  function bloques(hoja) {
    const out = [];
    for (let r = R0 - 1; r < hoja.rows.length; r++) {
      const f = hoja.rows[r] || [];
      if (txt(f[1]) === 'ID' && /Voltaje Nominal/i.test(txt(f[9]))) {
        const off = r + 1 - R0, k = out.length;
        const fila = rel => hoja.rows[rel + off - 1] || [];
        const b = {
          k, off, version: '2026',
          v: (rel, c) => fila(rel)[ci(c)],
          f: (rel, c) => !!(hoja.formulas[rel + off - 1] || [])[ci(c)],
        };
        b.nombre = txt(b.v(10, 'I')).split(' ').filter(Boolean).join(' '); b.prefijo = txt(b.v(10, 'B'));
        out.push(b);
      }
    }
    if (out.length) return out;
    // distribución antigua: la fila «ID» trae «Voltaje Nominal» en la columna E
    const ini = [];
    for (let r = R0 - 1; r < hoja.rows.length; r++) { const f = hoja.rows[r] || []; if (txt(f[1]) === 'ID' && /Voltaje Nominal/i.test(txt(f[4]))) ini.push(r); }
    ini.forEach((r, k) => out.push(bloqueAntiguo(hoja, r, ini[k + 1] === undefined ? hoja.rows.length : ini[k + 1], k)));
    return out;
  }

  /* Distribución antigua → se traduce a la actual: v(rel, columna) devuelve la celda que corresponde. */
  const MAP_BASE = { I: 'D', J: 'E', K: 'F', L: 'G', M: 'H', N: 'I', O: 'J', P: 'K', R: 'L', S: 'M', T: 'N', AD: 'X', AE: 'Y', AF: 'Z', AG: 'AA', AH: 'AB', AI: 'AC', AJ: 'AD', AK: 'AE', AL: 'AF', AN: 'AH', AP: 'AJ', AQ: 'AK', AU: 'AO', AV: 'AP', AX: 'AR', AY: 'AS', AZ: 'AT', BA: 'AU', BB: 'AV', BC: 'AW', BD: 'AX', BE: 'AY', BF: 'AZ' };
  const MAP_ALIM = Object.assign({}, MAP_BASE, { U: 'O', V: 'P', W: 'Q', AA: 'U', AB: 'V', AC: 'W' });
  const MAP_ENC = { B: 'B', I: 'D', J: 'E', L: 'G', M: 'H', O: 'I', P: 'J', R: 'L', T: 'N', V: 'P' };
  const MAP_CAT = { BI: 'BC', BJ: 'BD', BK: 'BE', BL: 'BF', BM: 'BG', BN: 'BH', BO: 'BI', BP: 'BJ' };
  const TIPOS_COL = { J: 'E', K: 'F', L: 'G', M: 'H', N: 'I', O: 'J', P: 'K', Q: 'L', R: 'M', S: 'N', T: 'O', W: 'P' };
  function bloqueAntiguo(hoja, r0, fin, k) {
    const R = i => hoja.rows[i] || [], F = i => hoja.formulas[i] || [];
    const nombre = txt(R(r0 + 1)[ci('D')]).split(' ').filter(Boolean).join(' '), prefijo = txt(R(r0 + 1)[ci('B')]);
    // filas de circuito: columna B = nombre del tablero y columna C numérica
    const circ = []; let ultima = r0 + 6;
    for (let i = r0 + 6; i < fin; i++) {
      const f = R(i);
      if (typeof f[2] === 'number' && txt(f[1]).split(' ').filter(Boolean).join(' ') === nombre) { circ.push(i); ultima = i; } else if (circ.length) break;
    }
    const buscar = (etq, desde) => { for (let i = desde; i < fin; i++) if (txt(R(i)[ci('D')]).indexOf(etq) === 0) return i; return -1; };
    const rAcom = buscar('Datos de acometida / alimentador', ultima), rFeed = rAcom < 0 ? -1 : rAcom + 3;
    const rRes = buscar('Tablero / Equipo', rAcom < 0 ? ultima : rAcom), rTub = rRes < 0 ? -1 : rRes + 2;
    const rTot = buscar('kVA Conectados totales', ultima), rFd = buscar('Factor demanda', ultima), rDv = buscar('Factor diversidad', ultima), rDm = buscar('kVA Demandados', ultima), rTd = buscar('Total kVA Demandados', ultima);
    let rCab = -1; for (let i = ultima; i < fin; i++) if (txt(R(i)[ci('D')]) === 'ID' && /ILUMINACI/i.test(txt(R(i)[ci('E')]))) { rCab = i; break; }
    const nFases = num(R(r0 + 1)[ci('G')]) || 3;
    // polos de cada circuito: la fila no trae polos, solo un número correlativo; se asignan según las fases con carga (A, B, C)
    const usados = new Set(), polos = {};
    const fasePolo = p => { const kk = Math.floor((p - 1) / 2); return nFases === 3 ? kk % 3 : nFases === 2 ? kk % 2 : 0; };
    circ.forEach(i => {
      const f = R(i); if (!(num(f[ci('F')]) > 0)) return;
      const np = Math.max(1, Math.min(3, num(f[ci('J')]) || 1)), ph = [num(f[ci('O')]), num(f[ci('P')]), num(f[ci('Q')])].map(x => x > 0), con = [0, 1, 2].filter(x => ph[x]);
      const base = con.length ? con[0] : (num(f[2]) - 1) % 3;
      let sel = null;
      if (nFases !== 3) { for (let p = 1; !sel && p < 200; p++) { const l = []; for (let q = 0; q < np; q++) l.push(p + 2 * q); if (l.every(x => !usados.has(x))) sel = l; } }
      else if (np === 1) { for (let p = 1; p < 200; p++) if (!usados.has(p) && fasePolo(p) === base) { sel = [p]; break; } }
      else if (np === 3) { for (let p = 1; p < 200; p++) { if (fasePolo(p) !== 0) continue; const l = [p, p + 2, p + 4]; if (l.every(x => !usados.has(x))) { sel = l; break; } } }
      else {
        const a = con.length === 2 ? con[0] : base, c2 = con.length === 2 ? con[1] : (base + 1) % 3;
        for (let p = 1; p < 200; p++) { if (fasePolo(p) !== a || fasePolo(p + 2) !== c2) continue; const l = [p, p + 2]; if (l.every(x => !usados.has(x))) { sel = l; break; } }
      }
      if (!sel) sel = [Math.max(0, ...Array.from(usados)) + 1];
      sel.forEach(x => usados.add(x)); polos[i] = sel;
    });
    const resuelve = (rel, c) => {   // → { i: fila (0-based), col: letra } o { valor }
      if (rel === 10) return MAP_ENC[c] ? { i: r0 + 1, col: MAP_ENC[c] } : null;
      if (rel === 11) return MAP_CAT[c] ? { i: r0 + 2, col: MAP_CAT[c] } : null;
      if (rel >= 16 && rel <= 115) {
        const i = circ[rel - 16]; if (i === undefined) return null;
        if (c === 'D' || c === 'F' || c === 'H') { const p = polos[i]; return { valor: p ? (p[c === 'D' ? 0 : c === 'F' ? 1 : 2] || '') : '' }; }
        if (rel === 16 && c === 'BI') return { i, col: 'BC' };
        return MAP_BASE[c] ? { i, col: MAP_BASE[c] } : null;
      }
      if (rel === 139) return rFeed >= 0 && MAP_ALIM[c] ? { i: rFeed, col: MAP_ALIM[c] } : null;
      if (rel === 140) return rTub >= 0 && c === 'AQ' ? { i: rTub, col: 'U' } : null;
      const tipos = { 121: rCab, 130: rTot, 131: rFd, 132: rDv, 133: rDm }[rel];
      if (tipos !== undefined) return tipos >= 0 && TIPOS_COL[c] ? { i: tipos, col: TIPOS_COL[c] } : null;
      if (rel === 134 && c === 'J' && rTd >= 0) return { i: rTd, col: 'E' };
      return null;
    };
    return {
      k, off: r0, version: '2024', nombre, prefijo,
      v: (rel, c) => { const t = resuelve(rel, c); return !t ? undefined : t.valor !== undefined ? t.valor : R(t.i)[ci(t.col)]; },
      f: (rel, c) => { const t = resuelve(rel, c); return !t || t.valor !== undefined ? false : !!F(t.i)[ci(t.col)]; },
    };
  }

  /** Marca de un breaker de la hoja (modelo de referencia, unidad, polos, kA) en el formato de la herramienta. */
  function breakerExcel(b, rel) {
    return { modelo: txt(b.v(rel, 'BA')), marco: num(b.v(rel, 'BB')), amperios: txt(b.v(rel, 'BC')), unidad: txt(b.v(rel, 'BD')) || 'STD', polos: num(b.v(rel, 'BE')), sccr: num(b.v(rel, 'BF')) };
  }

  /** Construye el borrador del proyecto con las entradas del Excel. */
  function construir(sheets, nombreArchivo) {
    const hoja = sheets.find(s => /^ANACAR$/i.test(s.name));
    if (!hoja) throw new Error('El archivo no tiene la hoja «ANACAR». ¿Es un proyecto de ANACAR en Excel?');
    const todos = bloques(hoja).filter(b => b.nombre && !/^machote$/i.test(b.nombre) && !/^x+$/i.test(b.nombre));   // «Machote» y «xxxxx» son los bloques plantilla
    if (!todos.length) throw new Error('No se encontraron tableros en la hoja ANACAR.');
    const esTab = b => /^TABLERO$/i.test(b.prefijo);
    const cl = x => String(x || '').toUpperCase().split(' ').filter(Boolean).join(' ');   // nombre sin diferencias de mayúsculas ni espacios
    const porNombre = {}; todos.forEach(b => { porNombre[cl(b.nombre)] = b; });
    // circuitos de un bloque (fila con detalle de carga y posición) y a qué bloque alimenta cada uno (descripción «TABLERO <nombre>»)
    const filas = b => { const out = []; for (let rel = 16; rel <= 115; rel++) if (num(b.v(rel, 'K')) > 0 && ['D', 'F', 'H'].some(c => num(b.v(rel, c)) > 0)) out.push(rel); return out; };
    const nC = b => filas(b).length;
    const hijosDe = b => filas(b).map(rel => { const d = cl(b.v(rel, 'I')); const h = d.indexOf('TABLERO ') === 0 ? porNombre[d.slice(8)] : null; return h && h !== b ? h : null; }).filter(Boolean);
    const padresDe = {}; todos.forEach(b => hijosDe(b).forEach(h => { const k = cl(h.nombre); (padresDe[k] = padresDe[k] || []); if (!padresDe[k].includes(b)) padresDe[k].push(b); }));
    // transferencia con generador: GENERADOR → ATS/MTS; el ATS y los IP que lo preceden son bloques de paso (con un circuito o ninguno)
    const gens = todos.filter(g => /^GEN/i.test(g.nombre) && nC(g) <= 3);
    const cadenas = [], equipoSet = new Set();
    todos.filter(a => /^(ATS|MTS)/i.test(a.nombre) && nC(a) <= 1).forEach(a => {
      const G = gens.filter(g => cl(txt(g.v(10, 'P'))) === cl(a.nombre) || (padresDe[cl(a.nombre)] || []).includes(g));
      if (!G.length) return;
      const ips = []; let U = null;
      const normal = (padresDe[cl(a.nombre)] || []).filter(p => !gens.includes(p));
      if (normal.length) {   // por circuitos: quién tiene un circuito hacia este ATS
        let P = normal[0];
        while (P && /^IP/i.test(P.nombre) && nC(P) <= 1 && ips.length < 4) { ips.push(P); P = (padresDe[cl(P.nombre)] || []).filter(p => !gens.includes(p))[0] || null; }
        U = P;
      } else {   // bloques de paso sin circuitos: se sigue «Conectado a»
        let e = porNombre[cl(txt(a.v(10, 'P')))];
        while (e && /^IP/i.test(e.nombre) && nC(e) <= 1 && ips.length < 4) { ips.push(e); e = porNombre[cl(txt(e.v(10, 'P')))]; }
        U = e;
      }
      cadenas.push({ g: G[0], a, ips, arriba: U && !equipoSet.has(U) ? U : null, arribaTxt: txt((ips.length ? ips[ips.length - 1] : a).v(10, 'P')) });
      [a].concat(ips, G).forEach(x => equipoSet.add(x));
    });
    const vacio = b => nC(b) === 0 && !(padresDe[cl(b.nombre)] || []).length && !equipoSet.has(b);   // bloque sin circuitos al que nadie alimenta (p. ej. un bloque en blanco)
    const equipos = todos.filter(b => equipoSet.has(b) || !esTab(b) || vacio(b));
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
      if (b.version === '2024' && num(b.v(139, 'S')) > 0 && num(b.v(139, 'S')) !== 1) {   // distribución antigua: la demanda del tablero se divide entre el factor de diversidad
        const dv = diversidad(b.v(139, 'S')); if (dv.ok) t.fdivTablero = dv.web; else aviso.push('«' + nombreTab(t) + '»: factor de diversidad del tablero ' + num(b.v(139, 'S')) + ' (divide y aumenta la demanda); la herramienta solo admite de 0 a 1, se dejó en 1.');
      } else if (!b.f(139, 'S') && num(b.v(139, 'S')) && num(b.v(139, 'S')) !== 1) {   // diversidad del tablero: según la versión del Excel multiplica o divide; se deduce de los valores calculados
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
        c._jConst = !b.f(rel, 'J'); c._jVal = num(b.v(rel, 'J'));
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
      mapa[cl(b.nombre)] = t;
    });

    // --- jerarquía ---
    // el circuito que alimenta a otro tablero tiene por descripción «TABLERO <nombre>». Un circuito no asignado conserva su carga como carga común.
    const sueltos = (c, t) => { c.tableroHijoId = ''; c.descripcion = c._descExcel || c.descripcion; };
    const circHacia = (p, hijoNombre) => p.circuitos.filter(c => cl(c._descExcel) === 'TABLERO ' + cl(hijoNombre));
    const enlazar = (p, t, via) => {
      const cs = circHacia(p, via || t.nombre); if (!cs.length) return false;
      // kVA escritos a mano en el circuito que alimenta al tablero (distintos de los del tablero de abajo): se conservan
      if (cs[0]._jConst && !via && Math.abs(cs[0]._jVal - num(t._bloque.v(139, 'J'))) > 0.05) { cs[0].kvaManual = cs[0]._jVal; aviso.push('«' + nombreTab(p) + '»: el circuito hacia «' + nombreTab(t) + '» tiene ' + cs[0]._jVal + ' kVA escritos a mano en el Excel (el tablero de abajo suma ' + redondea(num(t._bloque.v(139, 'J')), 2) + '); se conservó el valor del Excel.'); }
      cs[0].tableroHijoId = t.id; cs[0].descripcion = ''; cs[0].kva = 0; t.padreId = p.id; return true; };
    // transferencia con generador: el tablero que cuelga del ATS lleva la segunda acometida y pasa a colgar de lo que alimentaba la cadena
    cadenas.forEach(ch => {
      const nombresCadena = [ch.a].concat(ch.ips);
      const bajoCirc = hijosDe(ch.a).filter(h => esTab(h) && !equipos.includes(h));
      const bajoTxt = tabs.filter(b => cl(txt(b.v(10, 'P'))) === cl(ch.a.nombre));
      const bajo = Array.from(new Set(bajoCirc.concat(bajoTxt)));
      const distNormal = nombresCadena.reduce((x, e) => x + num(e.v(10, 'R')), 0);
      const primero = (ch.ips.length ? ch.ips[ch.ips.length - 1] : ch.a).nombre;
      bajo.forEach(b => {
        const t = mapa[cl(b.nombre)], up = ch.arriba && esTab(ch.arriba) && !equipos.includes(ch.arriba) ? mapa[cl(ch.arriba.nombre)] : null;
        if (!t) return;
        t.alterna = { activo: true, equipo: /^MTS/i.test(ch.a.nombre) ? 'MTS' : 'ATS', tipo: 'generador', nombre: ch.g.nombre, kva: '', xd: 12, origenId: '', fuenteId: '', longitud: num(ch.g.v(10, 'R')), distNormal, material: 'CU', aislamiento: 'XHHW-2', tuberia: 'EMT', paralelos: '' };
        t._padreFijo = true;
        if (up) {
          if (!enlazar(up, t, primero)) { t.padreId = up.id; aviso.push('«' + nombreTab(up) + '» no tiene el circuito hacia «' + primero + '»: la herramienta crea uno para «' + nombreTab(t) + '».'); }
          t.conectadoA = '';
        } else t.conectadoA = ch.arribaTxt || t.conectadoA;
        aviso.push('«' + nombreTab(t) + '»: ' + nombresCadena.map(x => x.nombre).join(' ← ') + ' y generador «' + ch.g.nombre + '» pasaron a la segunda acometida (' + t.alterna.equipo + '). Revise la capacidad del generador (kVA) y la distancia normal → ' + t.alterna.equipo + ' (' + distNormal + ' m, suma de ' + nombresCadena.map(x => x.nombre).join(' + ') + ').');
      });
    });
    // padre de los demás tableros: quien tenga un circuito hacia él; con varios candidatos, el que coincide con «Conectado a» y si no el primero
    const multi = [];
    tabs.forEach(b => {
      const t = mapa[cl(b.nombre)]; if (t._padreFijo) return;
      const cand = (padresDe[cl(b.nombre)] || []).filter(p => !equipos.includes(p)).map(p => mapa[cl(p.nombre)]).filter(p => p && p !== t);
      if (!cand.length) return;
      const p = cand.find(x => cl(x.nombre) === cl(t.conectadoA)) || cand[0];
      enlazar(p, t);
      if (cand.length > 1) multi.push('«' + nombreTab(t) + '» (' + cand.map(x => x.nombre).join(' y ') + ' → se tomó ' + p.nombre + ')');
    });
    // sin circuito que lo alimente: no se cuelga de nadie (la herramienta tendría que crear un circuito y sumar su carga al tablero de arriba); «Conectado a» queda como texto
    const huerfanos = tabs.map(b => mapa[cl(b.nombre)]).filter(t => !t.padreId && mapa[cl(t.conectadoA)] && mapa[cl(t.conectadoA)] !== t);
    if (huerfanos.length) aviso.push(huerfanos.length + ' tablero(s) dicen estar conectados a otro tablero pero este no tiene un circuito hacia ellos (por ejemplo tableros típicos de habitación): quedaron como tableros independientes y «Conectado a» quedó como texto, para no sumar su carga al otro: ' + huerfanos.slice(0, 12).map(t => nombreTab(t) + ' (' + t.conectadoA + ')').join('; ') + (huerfanos.length > 12 ? '; …' : '') + '.');
    // ciclos: se corta el enlace del tablero que cierra el ciclo
    const byId = {}; tabs.forEach(b => { const t = mapa[cl(b.nombre)]; byId[t.id] = t; });
    tabs.forEach(b => {
      const t = mapa[cl(b.nombre)], vistos = new Set([t.id]); let x = byId[t.padreId];
      while (x) { if (vistos.has(x.id)) { const p = byId[t.padreId]; if (p) p.circuitos.filter(c => c.tableroHijoId === t.id).forEach(c => sueltos(c, t)); aviso.push('Se cortó el enlace «' + nombreTab(p) + '» → «' + nombreTab(t) + '» porque cerraba un ciclo; el circuito quedó como carga común.'); t.padreId = ''; break; } vistos.add(x.id); x = byId[x.padreId]; }
    });
    // circuitos que apuntaban a un tablero pero no quedaron enlazados: carga común con los kVA del Excel
    tabs.forEach(b => { const t = mapa[cl(b.nombre)]; t.circuitos.forEach(c => { if (!c.tableroHijoId && /^TABLERO /i.test(cl(c._descExcel)) && porNombre[cl(c._descExcel).slice(8)]) { c.descripcion = c._descExcel; } }); });
    if (multi.length) aviso.push('Tableros alimentados por más de un circuito (la herramienta admite un solo tablero de arriba; los otros circuitos quedaron como carga común con los kVA del Excel): ' + multi.join('; ') + '.');
    // bloques que no son tableros y no forman parte de una transferencia con generador
    equipos.filter(e => !equipoSet.has(e)).forEach(e => aviso.push('«' + e.nombre + '» no se importó como tablero (' + (vacio(e) ? 'bloque sin circuitos al que nadie alimenta' : esTab(e) ? 'bloque de equipo' : 'sin prefijo TABLERO') + ').'));

    // orden y exclusión de las tablas resumen (hoja ORDEN TABLEROS)
    const ord = sheets.find(s => /^ORDEN TABLEROS$/i.test(s.name));
    if (ord) ord.rows.forEach(r => {
      const nm = txt(r[1]).replace(/^TABLERO\s+/i, '').toUpperCase(), t = mapa[nm]; if (!t) return;
      const o = r[2]; if (/^x$/i.test(txt(o))) t.excluirResumen = true; else if (num(o)) t.orden = num(o);
    });

    const proy = Store.emptyProject();
    proy.nombre = (nombreArchivo || '').replace(/\.(xlsm|xlsx)$/i, '').replace(/^ANACAR[\s\d.]*-?\s*/i, '').trim() || 'Proyecto importado de Excel';
    proy.tableros = tabs.map(b => mapa[cl(b.nombre)]);
    return { proy, tabs: tabs.map(b => mapa[cl(b.nombre)]), mapa, aviso, equipos };
  }

  /** Detalles de carga: cada Excel trae su propia hoja DCARGAS y los números pueden no coincidir con los de la herramienta.
      Se busca por descripción, voltaje y fases; si el tipo de carga o los factores difieren se actualiza el catálogo con el valor del Excel (queda en el reporte); si no existe se agrega. */
  function conciliarDetalles(sheets, res) {
    const hoja = sheets.find(s => /^DCARGAS$/i.test(s.name)), nota = [];
    if (!hoja) return nota;
    const ex = {}; hoja.rows.forEach(r => { if (typeof r[1] === 'number' && r[1] > 0 && txt(r[3])) ex[r[1]] = { tipo: num(r[2]), desc: txt(r[3]), v: num(r[4]), fases: num(r[5]), hilos: num(r[6]), fd: num(r[7]), fdiv: num(r[8]), fp: num(r[9]) }; });
    // distribución antigua: los tipos de carga se identifican por el nombre de la columna (el número puede no coincidir con el de la herramienta)
    const bl0 = res.tabs.length ? res.tabs[0]._bloque : null, tipoWeb = {}, tNuevos = [];
    if (bl0 && bl0.version === '2024') {
      const n2 = x => String(x || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
      'JKLMNOPQRST'.split('').forEach((col, i) => {
        const nombre = txt(bl0.v(121, col)); if (!nombre) return;
        let tp = Store.catalog.tiposCarga.find(x => n2(x.nombre) === n2(nombre));
        if (!tp) { tp = { id: Math.max(0, ...Store.catalog.tiposCarga.map(x => x.id)) + 1, nombre, metodo: 'fijo', fd: num(bl0.v(131, col)) || 0.8, fdiv: 1 }; Store.catalog.tiposCarga.push(tp); tNuevos.push(nombre); }
        tipoWeb[i + 1] = tp.id;
      });
    }
    const C = Store.catalog, norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();
    const memo = {}, tocados = [], nuevos = [];
    res.tabs.forEach(t => t.circuitos.forEach(c => {
      const id = c.detalleId, e = ex[id]; if (!e) return;
      if (!e._ok && tipoWeb[e.tipo]) { e.tipo = tipoWeb[e.tipo]; e._ok = true; }
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
    if (tNuevos.length) nota.push('Tipos de carga del Excel que no existían y se agregaron al catálogo: ' + tNuevos.join(', ') + '.');
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
      add('Caída de voltaje acumulada en el alimentador (%)', redondea(eAY, 2), redondea(r.alim.AY139, 2), !eAY || Math.abs(eAY - r.alim.AY139) <= Math.max(0.15, 0.3 * eAY));
      if (b.v(118, 'U') !== undefined) add('Desbalance máximo (%)', redondea(eDes, 1), redondea(r.desbalance, 1), Math.abs(eDes - r.desbalance) <= 0.2);
      // circuitos: se emparejan por la primera posición
      t.circuitos.forEach(c => {
        const x = r.rows.find(z => z.c === c); if (!x) return;
        const rel = c._rel, pos = '[' + c.polos.join(',') + '] ' + (x.descripcion || c._descExcel || '');
        const eP = num(b.v(rel, 'AF')), eC = txt(b.v(rel, 'AL')), eY = num(b.v(rel, 'AY'));
        if (c.tableroHijoId) return;
        if (eP && x.AF && eP !== x.AF) dif.push({ tablero: nm, que: 'Circuito ' + pos + ': protección (A)', excel: eP, web: x.AF, tipo: 'proteccion' });
        if (eC && x.AL && String(eC) !== String(x.AL)) dif.push({ tablero: nm, que: 'Circuito ' + pos + ': calibre', excel: eC, web: x.AL, tipo: 'cable' });
        const acumC = r.padre ? r.padre.alim.AY139 : 0;
        if (eY && x.AY !== null && Math.abs(eY - (x.AY - acumC)) > Math.max(0.15, 0.3 * eY) && Math.abs(eY - x.AY) > Math.max(0.15, 0.3 * eY)) dif.push({ tablero: nm, que: 'Circuito ' + pos + ': caída de voltaje (%, sin la acumulada de arriba)', excel: redondea(eY, 2), web: redondea(x.AY - acumC, 2), tipo: 'calculo' });
      });
    });
    return dif;
  }

  /** Importa el archivo: crea el proyecto en la herramienta y devuelve el reporte. */
  async function importar(buf, nombreArchivo) {
    const sheets = await XlsxRead.read(buf, { hojas: ['ANACAR', 'DCARGAS', 'ORDEN TABLEROS'] });
    const res = construir(sheets, nombreArchivo);
    Store.newProject();
    const P = Store.project; P.nombre = res.proy.nombre;
    P.cargaDerivados = 'sinreserva';   // el Excel suma al tablero de arriba los kVA conectados del de abajo sin su reserva
    res.tabs.forEach(t => { Store.completarTablero(t); P.tableros.push(t); });
    const nota = conciliarDetalles(sheets, res).concat(conciliarCatalogos(res));
    // los breakers escritos en el Excel se concilian con el catálogo después del primer cálculo
    Store.sincronizarBypass && Store.sincronizarBypass();
    let R = Calc.proyecto(P, Store.catalog);
    conciliarCircuitos(res, R, nota);
    conciliarAlimentadores(res, R, nota, 'prot');
    R = Calc.proyecto(P, Store.catalog);
    conciliarAlimentadores(res, R, nota, 'cable');
    R = Calc.proyecto(P, Store.catalog);
    conciliarBreakers(res, R, nota);
    R = Calc.proyecto(P, Store.catalog);
    const dif = comparar(res, R);
    const noCumple = R.orden.filter(r => r.alim.val && !r.alim.val.ok).map(r => r.tab.nombre);
    if (noCumple.length) nota.push('Según la herramienta, ' + noCumple.length + ' alimentador(es) del Excel no cumplen la ampacidad corregida por temperatura y agrupamiento (310.15): ' + noCumple.join(', ') + '. Se conservó el cableado del Excel; revise el aviso en la memoria de cada tablero.');
    res.tabs.forEach(t => { delete t._bloque; delete t._cat; delete t._spdId; delete t._padreFijo; t.circuitos.forEach(c => { delete c._rel; delete c._descExcel; delete c._kvaF; delete c._jConst; delete c._jVal; }); });
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

  /** Protección de cada circuito ramal: si la herramienta calcula otro amperaje se fija el del Excel (cada Excel trae su propia lista de tamaños y a veces calcula distinto). */
  function conciliarCircuitos(res, R, nota) {
    let n = 0; const ej = [];
    res.tabs.forEach(t => {
      const b = t._bloque, r = R.res[t.id]; if (!r) return;
      t.circuitos.forEach(c => {
        if (c.tableroHijoId || c.prot) return;
        const x = r.rows.find(z => z.c === c); if (!x) return;
        const eP = num(b.v(c._rel, 'AF'));
        if (eP > 0 && x.AF && x.AF !== eP) { c.prot = eP; n++; if (ej.length < 5) ej.push(nombreTab(t) + ' [' + c.polos.join(',') + '] ' + eP + ' A'); }
      });
    });
    if (n) nota.push('Protección de ' + n + ' circuito(s) ramal(es) fijada con el valor del Excel (' + ej.join('; ') + (n > ej.length ? '; …' : '') + ').');
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
