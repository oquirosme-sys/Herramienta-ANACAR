/* Hojas con el diseño del Excel: TABLEROS 3F / 1F, TABLA RESUMEN, (VERTICAL), DU(VERTICAL) y ORIGINAL.
   Cada hoja se arma como una cuadrícula con las mismas columnas, celdas combinadas y estilos de la hoja original;
   la misma cuadrícula se dibuja en pantalla (Hoja.html) y se exporta a Excel (export-excel.js). */
(function (g) {
  'use strict';
  const h = U.h;

  /** Índice de columna de Excel: 'A' → 0, 'B' → 1, 'AJ' → 35. */
  const col = L => { let n = 0; for (const ch of L) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  const LET = n => { let s = ''; for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };

  /** Estilos (iguales a las hojas del Excel: Century Gothic, títulos verdes, barras grises). */
  const ESTILOS = {
    tit: { fn: 'Century Gothic', sz: 48, b: 1, color: '#006600', al: 'center', border: 'medium' },
    tit2: { fn: 'Century Gothic', sz: 40, b: 1, color: '#006600', al: 'center', border: 'medium', wrap: true },
    alimL: { fn: 'Century Gothic', sz: 18, b: 1, bg: '#D9D9D9', al: 'center', border: 'thin' },
    alimV: { fn: 'Century Gothic', sz: 18, b: 1, al: 'center', border: 'medium', wrap: true },
    barra: { fn: 'Century Gothic', sz: 18, b: 1, bg: '#D9D9D9', al: 'center', border: 'thin' },
    th: { fn: 'Century Gothic', sz: 12, b: 1, al: 'center', wrap: true, border: 'thin' },
    thg: { fn: 'Century Gothic', sz: 12, b: 1, bg: '#D9D9D9', al: 'center', wrap: true, border: 'thin' },
    td: { fn: 'Century Gothic', sz: 12, b: 1, al: 'center', border: 'thin' },
    tdn: { fn: 'Century Gothic', sz: 12, al: 'center', border: 'thin' },
    tdl: { fn: 'Century Gothic', sz: 12, b: 1, al: 'left', border: 'thin', wrap: true },
    tdr: { fn: 'Century Gothic', sz: 12, b: 1, al: 'right', border: 'thin' },
    gris: { fn: 'Century Gothic', sz: 12, b: 1, bg: '#D9D9D9', al: 'center', border: 'thin' },
    grisr: { fn: 'Century Gothic', sz: 12, b: 1, bg: '#D9D9D9', al: 'right', border: 'thin' },
    tot: { fn: 'Century Gothic', sz: 12, b: 1, bg: '#D9D9D9', al: 'center', border: 'thin', wrap: true },
    ley: { fn: 'Century Gothic', sz: 9, al: 'left', wrap: true, border: 'none' },
    logo: { border: 'thin' },
  };

  /** Cuadrícula: celdas {y, x, y2, x2, v, s, f} (f = decimales si v es número). */
  function Grid(anchos) {
    const cells = [], altos = [];
    return {
      anchos, cells, altos,
      put(y, x1, y2, x2, v, s, f) { cells.push({ y, x: typeof x1 === 'string' ? col(x1) : x1, y2, x2: typeof x2 === 'string' ? col(x2) : x2, v: v === undefined || v === null ? '' : v, s, f }); },
      alto(y, pt) { altos[y] = pt; },
      filas() { return cells.reduce((m, c) => Math.max(m, c.y2), 0) + 1; },
    };
  }
  const num = (v, d) => (v === null || v === undefined || v === '' || isNaN(v) ? '' : Number(v));
  const txt = v => (v === null || v === undefined ? '' : String(v).trim());

  /* ---------- TABLEROS 3F / 1F ---------- */
  const ANCHOS_TAB = (() => {
    const w = { A: 1.5, B: 22.8, C: 2.8, D: 6.2, E: 5, F: 4.8, G: 21.1, H: 32.1, I: 18.1, J: 15.8, K: 13.8, L: 13.9, M: 13.5, N: 16.1, O: 22.2, P: 13.9, Q: 12, R: 13, S: 18.1, T: 16.2, U: 13.2, V: 15.8, W: 15.2,
      X: 18.1, Y: 2.8, Z: 6.1, AA: 5.1, AB: 5.5, AC: 7, AD: 7.5, AE: 6, AF: 5.8, AG: 7.5, AH: 5.8, AI: 12.5, AJ: 4.5 };
    return Object.keys(w).map(k => w[k]);
  })();

  function tablero(r) {
    const G = Grid(ANCHOS_TAB), t = r.tab, A = r.alim, cat = r.cat || {}, bk = r.bkMain || {}, spd = r.spd || {}, tres = r.fases === 3;
    const P = (y, a, b, v, s, f, y2) => G.put(y, a, y2 === undefined ? y : y2, b, v, s, f);
    // encabezado (filas 6 y 7 del Excel)
    P(0, 'B', 'L', r.nombre, 'tit', null, 1); P(0, 'M', 'O', 'Alimentado desde:', 'alimL'); P(1, 'M', 'O', r.alimentadoDesde || '', 'alimV'); P(0, 'P', 'AJ', '', 'logo', null, 1);
    G.alto(0, 61); G.alto(1, 56);
    // datos de cálculos eléctricos / especificación del tablero (filas 8 a 11)
    P(2, 'B', 'O', 'Datos de cálculos eléctricos', 'barra'); P(2, 'P', 'AJ', 'Especificación del tablero', 'barra'); G.alto(2, 48);
    [['B', 'F', 'KVA Conectados'], ['G', 'G', 'KVA Demandados'], ['H', 'H', 'kVA Reserva'], ['I', 'I', 'Factor demanda'], ['J', 'J', 'Factor diversidad'], ['K', 'K', 'Factor potencia'], ['O', 'O', 'Icc disponible (kA)'],
      ['P', 'Q', 'Tensión Nominal (Voltios)'], ['R', 'R', 'Fases'], ['S', 'S', 'Hilos'], ['T', 'U', 'Modelo de referencia'], ['V', 'V', 'Fabricante'], ['W', 'W', 'Montaje'], ['AF', 'AG', 'Espacios'], ['AH', 'AJ', '% DESBALANCEO MAXIMO']]
      .forEach(([a, b, v]) => P(3, a, b, v, 'th', null, 4));
    P(3, 'L', 'N', 'Amperios totales por fase', 'th'); P(4, 'L', 'L', 'Fase A', 'th'); P(4, 'M', 'M', 'Fase B', 'th'); P(4, 'N', 'N', 'Fase C', 'th');
    P(3, 'X', 'AE', 'Amperaje de Barras', 'th'); P(4, 'X', 'AA', 'Fases', 'th'); P(4, 'AB', 'AC', 'Neutro', 'th'); P(4, 'AD', 'AE', 'Tierra', 'th');
    G.alto(3, 23.1); G.alto(4, 23.1);
    const v5 = [['B', 'F', r.W130, 2], ['G', 'G', A.L139, 2], ['H', 'H', Math.round((Number(t.reserva) || 0) * 100) + '%'], ['I', 'I', A.R139, 2], ['J', 'J', A.S139, 2], ['K', 'K', A.T139, 2],
      ['L', 'L', A.AA139[0], 1], ['M', 'M', A.AA139[1], 1], ['N', 'N', tres ? A.AA139[2] : '', 1], ['O', 'O', r.iccKA ? num(r.iccKA) : '', 2], ['P', 'Q', t.sistema], ['R', 'R', r.fases], ['S', 'S', r.hilos],
      ['T', 'U', cat.modelo], ['V', 'V', r.marca], ['W', 'W', t.montaje], ['X', 'AA', cat.barraFase], ['AB', 'AC', cat.barraNeutro], ['AD', 'AE', cat.barraTierra], ['AF', 'AG', cat.espacios], ['AH', 'AJ', num(r.desbalance) / 100, 'pct']];
    v5.forEach(([a, b, v, f]) => P(5, a, b, v, 'td', f)); G.alto(5, 38.45);
    // alimentador / interruptor principal / supresor (filas 12 a 14)
    P(6, 'B', 'O', 'Datos del alimentador/acometida', 'barra'); P(6, 'P', 'W', 'Datos de Interruptor Principal', 'barra'); P(6, 'X', 'AJ', 'Datos Supresor (SPD)', 'barra'); G.alto(6, 38.45);
    [['B', 'F', 'Fases (AWG)'], ['G', 'G', 'Neutro (AWG)'], ['H', 'H', 'Tierra (AWG)'], ['I', 'I', 'Material'], ['J', 'J', 'Aislamiento'], ['K', 'K', 'Tubería (Ø mm)'], ['L', 'L', 'Longitud (m)'], ['M', 'M', 'Voltaje (V)'],
      ['N', 'N', 'Caída voltaje (V)'], ['O', 'O', 'Caída voltaje total (%)'], ['P', 'R', 'Modelo de referencia'], ['S', 'S', 'Amperaje (A)'], ['T', 'T', 'Marco (A)'], ['U', 'U', 'Tipo'], ['V', 'V', 'Polos'], ['W', 'W', 'SCCR (kA)'],
      ['X', 'AC', 'Modelo de referencia'], ['AD', 'AF', 'Capacidad de supresión (kA)'], ['AI', 'AJ', 'Montaje']].forEach(([a, b, v]) => P(7, a, b, v, 'th'));
    P(7, 'AG', 'AH', '', 'th', null, 8);
    G.alto(7, 38.45);
    [['B', 'F', A.fasesTxt], ['G', 'G', A.neutroTxt], ['H', 'H', A.tierraTxt], ['I', 'I', A.mat], ['J', 'J', A.ais], ['K', 'K', (A.preN || '') + txt(A.AR139)], ['L', 'L', A.M139, 1], ['M', 'M', A.AW139, 2],
      ['N', 'N', A.AX139, 2], ['O', 'O', A.AY139, 2], ['P', 'R', bk.ref || bk.modelo], ['S', 'S', A.AF139], ['T', 'T', Number(bk.marco) ? bk.marco : ''], ['U', 'U', bk.unidad], ['V', 'V', bk.polos], ['W', 'W', bk.sccr],
      ['X', 'AC', spd.modelo], ['AD', 'AF', spd.kaLL], ['AI', 'AJ', spd.montaje]].forEach(([a, b, v, f]) => P(8, a, b, v, 'td', f));
    G.alto(8, 38.45);
    // circuitos ramales / posición en el tablero (filas 15 a 17)
    P(9, 'B', 'W', 'Datos de circuitos ramales / alimentadores', 'barra'); P(9, 'X', 'AJ', 'Posición en el tablero', 'barra'); G.alto(9, 38.45);
    [['B', 'F', 'ID de circuito'], ['G', 'H', 'Descripción de carga'], ['I', 'I', 'kVA Conectados'], ['J', 'J', 'Caída voltaje total (%)'], ['X', 'AB', 'ID de circuito']].forEach(([a, b, v]) => P(10, a, b, v, 'th', null, 11));
    P(10, 'K', 'Q', 'Datos de Interruptores Ramales', 'th'); P(10, 'R', 'V', 'Datos conductores ramales / alimentadores', 'th'); P(10, 'W', 'W', 'Tubo EMT', 'th');
    [['K', 'L', 'Modelo de referencia'], ['M', 'M', 'Marco (A)'], ['N', 'N', 'Amperios (A)'], ['O', 'O', 'Tipo de unidad'], ['P', 'P', 'Polos'], ['Q', 'Q', 'SCCR (kA)'], ['R', 'R', 'Fases'], ['S', 'S', 'Neutro'],
      ['T', 'T', 'Tierra'], ['U', 'U', 'Tipo'], ['V', 'V', 'Aislamiento'], ['W', 'W', '(Ø mm)']].forEach(([a, b, v]) => P(11, a, b, v, 'th'));
    const FC = tres ? [['AC', 'AE'], ['AF', 'AH'], ['AI', 'AJ']] : [['AC', 'AF'], ['AG', 'AJ']];
    FC.forEach(([a, b], i) => { P(10, a, b, 'KVA', 'th'); P(11, a, b, 'FASE ' + 'ABC'[i], 'th'); });
    G.alto(10, 22.35); G.alto(11, 45.6);

    // una fila por posición: primero las impares y luego las pares (como la hoja del Excel)
    const esp = Math.max(Number(cat.espacios) || 0, r.espaciosUsados, 2);
    const total = esp % 2 ? esp + 1 : esp;
    const pos = [];
    for (let p = 1; p <= total; p += 2) pos.push(p);
    for (let p = 2; p <= total; p += 2) pos.push(p);
    const inicio = {}, cont = {};
    r.rows.forEach(x => { if (x.polos.length) { inicio[x.polos[0]] = x; x.polos.slice(1).forEach(p => { cont[p] = x; }); } });
    let y = 12;
    pos.forEach(p => {
      const x = inicio[p], c = cont[p], nm = t.nombre || '';
      const ph = Calc.fasePolo(p, r.fases);
      if (x) {
        const bkr = x.breaker || {};
        P(y, 'B', 'B', nm, 'tdr'); P(y, 'C', 'C', '-', 'td'); P(y, 'D', 'D', x.polos[0], 'tdr'); P(y, 'E', 'E', x.polos[1] || '', 'tdr'); P(y, 'F', 'F', x.polos[2] || '', 'tdr');
        P(y, 'G', 'H', x.descripcion, x.J ? 'td' : 'gris'); P(y, 'I', 'I', x.J || '', 'td', 2); P(y, 'J', 'J', x.AY === null ? '' : x.AY, 'td', 2);
        P(y, 'K', 'L', x.J ? (bkr.ref || bkr.modelo || '') : '', 'td'); P(y, 'M', 'M', Number(bkr.marco) ? bkr.marco : '', 'td'); P(y, 'N', 'N', x.AF || '', 'td'); P(y, 'O', 'O', x.J ? bkr.unidad || '' : '', 'td');
        P(y, 'P', 'P', x.J ? x.polosBreaker : '', 'td'); P(y, 'Q', 'Q', x.J ? bkr.sccr || '' : '', 'td'); P(y, 'R', 'R', x.AL ? x.fasesTxt.trim() : '', 'td'); P(y, 'S', 'S', x.neutroTxt, 'td');
        P(y, 'T', 'T', x.tierraTxt, 'td'); P(y, 'U', 'U', x.J ? (x.mat === 'AL-MC' ? 'AL' : x.mat) : '', 'td'); P(y, 'V', 'V', x.J ? x.ais : '', 'td'); P(y, 'W', 'W', x.tuboTxt, 'td');
        P(y, 'X', 'X', nm, 'tdr'); P(y, 'Y', 'Y', '-', 'td'); P(y, 'Z', 'Z', x.polos[0], 'tdr'); P(y, 'AA', 'AA', x.polos[1] || '', 'tdr'); P(y, 'AB', 'AB', x.polos[2] || '', 'tdr');
        FC.forEach(([a, b], i) => P(y, a, b, x.fase[i] || '', x.fase[i] || i === ph ? 'td' : 'gris', 2));
      } else {
        P(y, 'B', 'B', nm, 'tdr'); P(y, 'C', 'C', '-', 'td'); P(y, 'D', 'D', p, 'tdr'); P(y, 'E', 'E', '', 'tdr'); P(y, 'F', 'F', '', 'tdr');
        P(y, 'G', 'H', c ? '' : '', 'gris');
        ['I', 'J', 'K', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W'].forEach(L => P(y, L, L === 'K' ? 'L' : L, '', 'td'));
        P(y, 'X', 'X', nm, 'tdr'); P(y, 'Y', 'Y', '-', 'td'); P(y, 'Z', 'Z', p, 'tdr'); P(y, 'AA', 'AA', '', 'tdr'); P(y, 'AB', 'AB', '', 'tdr');
        FC.forEach(([a, b], i) => P(y, a, b, '', !c && i === ph ? 'td' : 'gris'));
      }
      G.alto(y, 23.45); y++;
    });
    P(y, 'B', 'W', Store.catalog.leyendaUnidades || '', 'ley', null, y + 1); P(y, 'X', 'AB', 'kVA Conectados sin reserva', 'tot', null, y + 1);
    FC.forEach(([a, b], i) => P(y, a, b, r.U116[i], 'tot', 2, y + 1));
    G.alto(y, 23.45); G.alto(y + 1, 23.45);
    return G;
  }

  /* ---------- TABLA RESUMEN (horizontal, B:AU) ---------- */
  const ANCHOS_RES = (() => {
    const w = { A: 4, B: 35.8, C: 30.2, D: 18.8, E: 18.9, F: 15.5, G: 15.8, H: 15.1, I: 15, J: 9, K: 15.5, L: 13.9, M: 13.9, N: 9, O: 13.8, P: 18.2, Q: 16.8, R: 10, S: 17.8, T: 12, U: 14, V: 14, W: 14,
      X: 19.1, Y: 18.2, Z: 32.1, AA: 23.5, AB: 16.5, AC: 15, AD: 10, AE: 10, AF: 16, AG: 20.2, AH: 21.8, AI: 34.9, AJ: 21.8, AK: 12, AL: 14, AM: 12, AN: 21.5, AO: 30.9, AP: 43, AQ: 16.2, AR: 21.8, AS: 14, AT: 10, AU: 12 };
    return Object.keys(w).map(k => w[k]);
  })();
  function resumen(lista) {
    const G = Grid(ANCHOS_RES), P = (y, a, b, v, s, f, y2) => G.put(y, a, y2 === undefined ? y : y2, b, v, s, f);
    P(0, 'B', 'N', 'TABLA RESUMEN - TABLEROS ELÉCTRICOS', 'tit', null, 7); P(0, 'O', 'Y', '', 'logo', null, 7);
    P(0, 'Z', 'AG', 'DATOS DEL TABLERO', 'tit2', null, 7); P(0, 'AH', 'AN', 'DATOS DEL SUPRESOR', 'tit2', null, 7); P(0, 'AO', 'AU', 'DATOS INTERRUPTOR PRINCIPAL', 'tit2', null, 7);
    for (let i = 0; i < 8; i++) G.alto(i, 22.35);
    const dos = [['B', 'Tablero / Equipo'], ['C', 'Alimentado desde'], ['D', 'kVA Totales'], ['E', 'kVA Demandados'], ['F', 'Factor demanda'], ['G', 'Factor diversidad'], ['H', 'Factor potencia'],
      ['T', 'Longitud (m)'], ['U', 'Voltaje bornes (V)'], ['V', 'Caída de voltaje total (V)'], ['W', 'Caída de voltaje total (%)'], ['X', 'Corriente cortocircuito disponible (kA)'], ['Y', 'Capacidad de breaker (A)'],
      ['Z', 'Tablero / Equipo'], ['AA', 'Modelo de referencia'], ['AB', 'Fabricante'], ['AF', 'Espacios'], ['AG', 'Montaje'], ['AH', 'Fabricante'], ['AI', 'Modelo de referencia'], ['AJ', 'Montaje'],
      ['AM', 'Voltaje (V)'], ['AN', 'Fases'], ['AO', 'Fabricante'], ['AP', 'Modelo de referencia'], ['AQ', 'Marco (A)'], ['AR', 'Amperios (A)'], ['AS', 'Tipo de unidad'], ['AT', '# polos'], ['AU', 'SCCR (kA)']];
    dos.forEach(([L, v]) => P(8, L, L, v, 'thg', null, 9));
    P(8, 'I', 'P', 'Calibres conductores alimentador', 'thg'); [['I', 'J', 'Fases (AWG)'], ['K', 'L', 'Neutro (AWG)'], ['M', 'N', 'Tierra (AWG)'], ['O', 'O', 'Material'], ['P', 'P', 'Aislamiento']].forEach(([a, b, v]) => P(9, a, b, v, 'thg'));
    P(8, 'Q', 'S', 'Tubería', 'thg'); [['Q', 'Cantidad'], ['R', '(Ø mm)'], ['S', 'Tipo']].forEach(([L, v]) => P(9, L, L, v, 'thg'));
    P(8, 'AC', 'AE', 'Barras', 'thg'); [['AC', 'Fase'], ['AD', 'Neutro'], ['AE', 'Tierra']].forEach(([L, v]) => P(9, L, L, v, 'thg'));
    P(8, 'AK', 'AL', 'Capacidad de supresión (kA)', 'thg'); P(9, 'AK', 'AK', 'L-L', 'thg'); P(9, 'AL', 'AL', 'L-N / L-T / N-T', 'thg');
    G.alto(8, 23.45); G.alto(9, 44.45);
    lista.forEach((r, i) => {
      const v = Resumen.horizontal(r), y = 10 + i;
      v.forEach((x, k) => P(y, 1 + k, 1 + k, x, k === 0 || k === 24 ? 'td' : 'tdn', typeof x === 'number' && !Number.isInteger(x) ? 2 : null));
      G.alto(y, 32.45);
    });
    return G;
  }

  /* ---------- tablas verticales: (VERTICAL), DU(VERTICAL), ORIGINAL ---------- */
  function vertical(lista, fn, titulo) {
    const nT = Math.max(lista.length, 1), anchos = [3, 42].concat(Array(nT * 2).fill(13));
    const G = Grid(anchos), P = (y, a, b, v, s, f, y2) => G.put(y, a, y2 === undefined ? y : y2, b, v, s, f);
    const fin = 1 + nT * 2;
    P(0, 1, Math.max(2, fin - 3), titulo, 'tit2', null, 5); P(0, Math.max(3, fin - 2), fin, '', 'logo', null, 5);
    for (let i = 0; i < 6; i++) G.alto(i, 20);
    P(6, 1, 1, 'Tablero / Alimentadores', 'thg');
    lista.forEach((r, j) => P(6, 2 + 2 * j, 3 + 2 * j, r.nombre, 'thg'));
    G.alto(6, 32);
    if (!lista.length) return G;
    const filas = lista.map(fn);
    filas[0].forEach((f, i) => {
      const y = 7 + i, sub = f.length === 2 && f[1] === '' && i > 0;
      P(y, 1, 1, f[0], sub ? 'thg' : 'tdl');
      filas.forEach((fl, j) => {
        const x = fl[i];
        if (x.length > 2) { P(y, 2 + 2 * j, 2 + 2 * j, x[1], 'tdr'); P(y, 3 + 2 * j, 3 + 2 * j, x[2], 'td', typeof x[2] === 'number' && !Number.isInteger(x[2]) ? 2 : null); }
        else P(y, 2 + 2 * j, 3 + 2 * j, x[1], sub ? 'thg' : 'td', typeof x[1] === 'number' && !Number.isInteger(x[1]) ? 2 : null);
      });
      G.alto(y, 26);
    });
    return G;
  }

  /* ---------- dibujo en pantalla ---------- */
  function fmtCelda(c) {
    if (typeof c.v === 'number') {
      if (c.f === 'pct') return U.fmt(c.v * 100, 2) + '%';
      if (c.f !== null && c.f !== undefined) return U.fmt(c.v, c.f);
      return U.fmt(c.v, Number.isInteger(c.v) ? 0 : 2);
    }
    return c.v;
  }
  function html(G, opts) {
    opts = opts || {};
    const filas = G.filas(), ocup = {}, inicio = {};
    G.cells.forEach(c => { inicio[c.y + ':' + c.x] = c; for (let y = c.y; y <= c.y2; y++) for (let x = c.x; x <= c.x2; x++) ocup[y + ':' + x] = c; });
    const x0 = Math.min(...G.cells.map(c => c.x)), x1 = Math.max(...G.cells.map(c => c.x2));
    const tb = h('tbody');
    for (let y = 0; y < filas; y++) {
      const tr = h('tr', { style: { height: Math.round((G.altos[y] || 20) * 1.333) + 'px' } });
      for (let x = x0; x <= x1; x++) {
        const c = inicio[y + ':' + x];
        if (c) {
          const td = h('td', { class: 's-' + c.s, rowspan: c.y2 > c.y ? c.y2 - c.y + 1 : null, colspan: c.x2 > c.x ? c.x2 - c.x + 1 : null });
          if (c.s === 'logo') td.appendChild(h('img', { src: 'img/logo-sinergia.png', alt: 'Sinergia Ingeniería' }));
          else td.textContent = fmtCelda(c);
          tr.appendChild(td);
        } else if (!ocup[y + ':' + x]) tr.appendChild(h('td', { class: 's-vacio' }));
      }
      tb.appendChild(tr);
    }
    const table = h('table', { class: 'hx' }, h('colgroup', null, G.anchos.slice(x0, x1 + 1).map(w => h('col', { style: { width: Math.round(w * 7 + 5) + 'px' } }))), tb);
    const wrap = h('div', { class: 'hx-wrap' + (opts.real ? ' real' : '') }, table);
    // se ajusta al ancho disponible (como "Ajustar a la página" de Excel)
    const ajustar = () => {
      if (!wrap.isConnected || opts.real || !wrap.clientWidth) return;
      table.style.zoom = 1;
      const k = Math.min(1, (wrap.clientWidth - 2) / table.scrollWidth); table.style.zoom = k > 0 ? k : 1;
    };
    requestAnimationFrame(ajustar); setTimeout(ajustar, 60);
    if (window.ResizeObserver) new ResizeObserver(() => ajustar()).observe(wrap);
    return wrap;
  }

  g.Hoja = { tablero, resumen, vertical, html, ESTILOS, col, LET, real: false };
})(window);
