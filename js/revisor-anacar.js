/* Revisor de la importación de ANACAR Excel: compara los tableros que quedaron en la herramienta con los de las pestañas «TABLEROS 3F» y «TABLEROS 1F» del Excel
   (la ficha de cada tablero: datos de cálculo, especificación, alimentador, interruptor principal, supresor y la lista de circuitos con su interruptor y su cable)
   y arma un reporte con lo que coincide y lo que no. No cambia nada del proyecto. Soporta las dos distribuciones de esas pestañas (las columnas se buscan por su rótulo). */
(function (g) {
  'use strict';
  const ci = c => { let n = 0; for (const ch of c) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  const txt = v => { const s = v === null || v === undefined || v === false ? '' : String(v).trim(); return /^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!)/i.test(s) ? '' : s; };   // los errores de Excel (#N/A…) cuentan como celda vacía
  const norm = s => txt(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ');
  const cl = s => txt(s).toUpperCase().replace(/\s+/g, ' ');
  const num = v => { if (typeof v === 'number') return v; const s = txt(v).replace(/\s/g, ''); if (!s) return null; const pct = /%$/.test(s); const x = parseFloat(s.replace(',', '.').replace(/[^0-9.\-]/g, '')); return isNaN(x) ? null : (pct ? x / 100 : x); };
  const r2 = x => Math.round(x * 100) / 100;
  const fmt = v => (v === null || v === undefined || v === '' ? '—' : typeof v === 'number' ? String(r2(v)).replace('.', ',') : String(v));

  /* ---------- prioridad de cada diferencia ---------- */
  const ORD = { alta: 0, media: 1, baja: 2 }, ETQ_PRIO = { alta: 'Alta', media: 'Media', baja: 'Baja' };
  const idxCal = c => { const i = (window.Calc && Calc.CALIBRES ? Calc.CALIBRES : []).indexOf(String(c)); return i; };
  /** Alta (críticos): cambió el cableado (cables, material, aislamiento), el tamaño de la protección, las distancias, la descripción, el tipo de unidad del interruptor (GFCI / AFCI),
      los factores de demanda o de diversidad o el número de circuitos; también los circuitos que no cruzan.
      Media: cambia el modelo de las protecciones, el SCCR del interruptor o el supresor (y los demás datos de equipo o de cálculo sin clasificar). Baja: caídas de voltaje, montaje, hilos y las notas. */
  function prioridad(campo, ex, we, nivel) {
    if (nivel === 'info') return 'baja';
    const c = norm(campo);
    if (/caida|voltaje en bornes/.test(c)) return 'baja';                                         // caídas de voltaje
    if (/^cable:|alimentador: (fases|neutro|tierra|material|aislamiento)/.test(c)) return 'alta';   // cableado
    if (/amper|longitud|descripcion|tipo de unidad|factor de demanda|factor de diversidad|numero de circuitos/.test(c)) return 'alta';   // protección (tamaño), distancias, descripción, GFCI/AFCI, factores, circuitos
    if (/montaje|hilos/.test(c)) return 'baja';
    return 'media';                                                                                  // modelo de protecciones, SCCR, supresor y demás
  }

  /** Rótulos de una fila: [{ col, t }] con el texto normalizado. */
  const etiq = row => { const o = []; (row || []).forEach((v, col) => { if (txt(v)) o.push({ col, t: norm(v) }); }); return o; };
  const col = (L, pref, k) => { const f = L.filter(x => x.t.startsWith(pref)); const e = f[k || 0]; return e ? e.col : -1; };
  const colEq = (L, t) => { const e = L.find(x => x.t === t); return e ? e.col : -1; };

  /** Calibre de un cable escrito como «2x2#3/0», «3#4», « 10», «2x63,5»: { par, n, cal }. */
  const cable = s => { const t = txt(s); if (!t) return null; const m = /^(?:(\d+)\s*x\s*)?(?:(\d+)\s*#\s*)?(.+)$/i.exec(t); return m ? { par: Number(m[1] || 1), n: m[2] ? Number(m[2]) : null, cal: m[3].trim().replace(',', '.') } : { par: 1, n: null, cal: t }; };

  /* ---------- lectura de las fichas ---------- */
  function leerBloque(rows, r0) {
    const R = i => rows[i] || [];
    const b = { fila: r0 + 1, nombre: txt(R(r0)[1]).replace(/^TABLERO\s+/i, '').trim(), circuitos: [] };
    const ad = (R(r0) || []).findIndex(v => /alimentado desde/i.test(txt(v)));
    b.alimentadoDesde = ad >= 0 ? txt(R(r0 + 1)[ad]).replace(/^TABLERO\s+/i, '') : '';
    const L3 = etiq(R(r0 + 3)), L4 = etiq(R(r0 + 4)), V5 = R(r0 + 5);
    const v3 = (pref, k) => { const c = col(L3, pref, k); return c < 0 ? undefined : V5[c]; };
    const v4 = t => { const c = colEq(L4, t); return c < 0 ? undefined : V5[c]; };
    b.cab = {
      kvaConectados: v3('kva conectados'), kvaDemandados: v3('kva demandados'), reserva: v3('kva reserva') !== undefined ? v3('kva reserva') : v3('% reserva'),
      fd: v3('factor demanda'), fdiv: v3('factor diversidad'), fp: v3('factor potencia'), icc: v3('icc disponible'), sistema: v3('tension nominal'),
      fases: v3('fases'), hilos: v3('hilos'), modelo: v3('modelo de referencia'), fabricante: v3('fabricante'), montaje: v3('montaje'), espacios: v3('espacios'),
      barraFase: v4('fases'), barraNeutro: v4('neutro'), barraTierra: v4('tierra'), desbalance: v3('% desbalanceo'),
      amp: ['fase a', 'fase b', 'fase c'].map(t => v4(t)),
    };
    const L7 = etiq(R(r0 + 7)), V8 = R(r0 + 8);
    const v7 = (pref, k) => { const c = col(L7, pref, k); return c < 0 ? undefined : V8[c]; };
    b.alim = { fases: v7('fases'), neutro: v7('neutro'), tierra: v7('tierra'), material: v7('material'), aislamiento: v7('aislamiento'), tuberia: v7('tuberia'), longitud: v7('longitud'), voltaje: v7('voltaje'), caidaV: v7('caida voltaje (v)'), caidaPct: v7('caida voltaje total') };
    b.bk = { modelo: v7('modelo de referencia', 0), amperaje: v7('amperaje'), marco: v7('marco'), tipo: v7('tipo'), polos: v7('polos'), sccr: v7('sccr') };
    b.spd = { modelo: v7('modelo de referencia', 1), capacidad: v7('capacidad de supresion'), montaje: v7('montaje') };
    // circuitos
    const H = etiq(R(r0 + 10)), S = etiq(R(r0 + 11));
    const cDesc = col(H, 'descripcion'), cKva = col(H, 'kva conectados'), cCaida = col(H, 'caida voltaje total');
    const cMod = col(S, 'modelo de referencia'), cMarco = col(S, 'marco'), cAmp = col(S, 'amperios'), cUni = col(S, 'tipo de unidad'), cPol = col(S, 'polos'), cScc = col(S, 'sccr');
    const cFas = colEq(S, 'fases'), cNeu = colEq(S, 'neutro'), cTie = colEq(S, 'tierra'), cTip = colEq(S, 'tipo'), cAis = colEq(S, 'aislamiento'), cTub = (S.find(x => /mm\)$/.test(x.t)) || { col: -1 }).col;
    for (let i = r0 + 12; i < rows.length; i++) {
      const f = R(i), id = txt(f[1]); if (!id || /^TABLERO\s/i.test(id)) break;
      const d = cDesc >= 0 ? txt(f[cDesc]) : '', k = cKva >= 0 ? num(f[cKva]) : null;
      if ((!d || d === '0') && !(k > 0)) continue;   // posición libre (el Excel deja «0» en la descripción de las filas vacías)
      const polos = txt(f[2]) === '-' ? [3, 4, 5].map(c => num(f[c])).filter(x => x > 0) : [];
      b.circuitos.push({ id, fila: i + 1, desc: d, kva: k, polos, caida: cCaida >= 0 ? num(f[cCaida]) : null,
        bk: { modelo: cMod >= 0 ? txt(f[cMod]) : '', marco: cMarco >= 0 ? num(f[cMarco]) : null, amperios: cAmp >= 0 ? num(f[cAmp]) : null, tipo: cUni >= 0 ? txt(f[cUni]) : '', polos: cPol >= 0 ? num(f[cPol]) : null, sccr: cScc >= 0 ? num(f[cScc]) : null },
        cab: { fases: cFas >= 0 ? f[cFas] : '', neutro: cNeu >= 0 ? f[cNeu] : '', tierra: cTie >= 0 ? f[cTie] : '', tipo: cTip >= 0 ? txt(f[cTip]) : '', aislamiento: cAis >= 0 ? txt(f[cAis]) : '', tubo: cTub >= 0 ? num(f[cTub]) : null } });
    }
    return b;
  }
  function leerFichas(sheets) {
    const out = [];
    ['TABLEROS 3F', 'TABLEROS 1F'].forEach(nombre => {
      const hoja = sheets.find(s => s.name.toUpperCase() === nombre); if (!hoja) return;
      for (let r = 0; r < hoja.rows.length - 3; r++) {
        const f = hoja.rows[r] || [];
        if (/^TABLERO\s/i.test(txt(f[1])) && /^datos de c/i.test(norm((hoja.rows[r + 2] || [])[1]))) { const b = leerBloque(hoja.rows, r); b.hoja = nombre.slice(-2); out.push(b); }
      }
    });
    return out;
  }

  /* ---------- comparación ---------- */
  function revisar(sheets, P, R, archivo) {
    const fichas = leerFichas(sheets).filter(b => !/^(machote|x+)$/i.test(b.nombre)), usadas = new Set(), tableros = [];   // «Machote» y «xxxxx» son las plantillas
    const limpiaNombre = s => cl(s).replace(/^TABLERO\s+/, '');
    const hallar = (t, k) => {   // k-ésimo tablero de la herramienta con ese nombre (los nombres repetidos se importan como «NOMBRE (2)»)
      const m = /^(.*) \((\d+)\)$/.exec(txt(t.nombre)), base = limpiaNombre(m ? m[1] : t.nombre), n = m ? Number(m[2]) - 1 : 0;
      const pref = t.tipo === '3F' ? '3F' : '1F', cand = fichas.filter(b => limpiaNombre(b.nombre) === base);
      const porHoja = h => cand.filter(b => b.hoja === h);
      const lista = porHoja(pref).length > n ? porHoja(pref) : porHoja(pref === '3F' ? '1F' : '3F').length > n ? porHoja(pref === '3F' ? '1F' : '3F') : cand;
      return lista[n] || null;
    };
    P.tableros.forEach(t => {
      const r = R.res[t.id]; if (!r) return;
      const b = hallar(t); const inf = { nombre: Store && Calc ? Calc.nombreTablero(t) : t.nombre, id: t.id, tipo: t.tipo, hoja: b ? b.hoja : null, fila: b ? b.fila : null, cab: [], circ: [], aviso: [] };
      tableros.push(inf); if (!b) { inf.estado = 'sin ficha'; inf.nAlta = 1; inf.nMedia = 0; inf.nBaja = 0; inf.nDato = 1; inf.nInfo = 0; return; }
      usadas.add(b);
      const dif = (lista, campo, ex, we, o) => { o = o || {}; const nivel = o.nivel || 'dato'; lista.push({ campo, excel: ex, web: we, nivel, nota: o.nota || '', prio: prioridad(campo, ex, we, nivel) }); };
      const cmpN = (lista, campo, ex, we, tol, o) => { const e = num(ex); if (e === null || we === null || we === undefined) return; if (Math.abs(e - Number(we)) > (tol || 0.01) * (o && o.rel ? Math.max(1, Math.abs(e)) : 1)) dif(lista, campo, e, Number(we), o); };
      const cmpT = (lista, campo, ex, we, o) => { const e = cl(ex); if (!e) return; if (e !== cl(we)) dif(lista, campo, txt(ex), txt(we), o); };
      const cmpModelo = (lista, campo, ex, we) => { const e = cl(ex), w = cl(we); if (!e || e === '0' || e === w) return; if (w && (w.startsWith(e) || e.startsWith(w))) dif(lista, campo, txt(ex), txt(we), { nivel: 'info', nota: 'Misma familia: el Excel anota la familia y la herramienta el modelo completo' }); else dif(lista, campo, txt(ex), txt(we)); };
      const cmpCable = (lista, campo, ex, we, o) => {
        const e = cable(ex), w = cable(we); if (!e) return; if (!w) { dif(lista, campo, txt(ex), txt(we), o); return; }
        const dis = e.cal !== w.cal || e.par !== w.par;
        if (dis) dif(lista, campo, txt(ex), txt(we), o);
        else if (e.n && w.n && e.n !== w.n) dif(lista, campo, txt(ex), txt(we), Object.assign({}, o, { nivel: 'info', nota: 'Misma sección; cambia el número de conductores por juego (en los tableros de 120/240 V el Excel usa 2 con «Fases = 1»)' }));
      };
      const C = b.cab, A = r.alim, cat = r.cat || {};
      const h = inf.cab;
      cmpN(h, 'kVA conectados (con reserva)', C.kvaConectados, r.W130, 0.05);
      cmpN(h, 'kVA demandados', C.kvaDemandados, A.L139, 0.05);
      { const e = num(C.reserva); if (e !== null) { const ee = e > 1 ? e / 100 : e; if (Math.abs(ee - Number(t.reserva || 0)) > 0.001) dif(h, 'Reserva', ee, Number(t.reserva || 0)); } }
      cmpN(h, 'Factor de demanda', C.fd, A.R139, 0.005); cmpN(h, 'Factor de diversidad', C.fdiv, A.S139, 0.005); cmpN(h, 'Factor de potencia', C.fp, A.T139, 0.005);
      (C.amp || []).forEach((a, i) => { if (num(a) !== null && num(a) > 0) cmpN(h, 'Corriente por fase ' + 'ABC'[i], a, (A.AA139 || [])[i], 0.03, { rel: true, nivel: 'info', nota: 'La herramienta calcula la corriente del alimentador con la fórmula trifásica; el Excel viejo de 120/240 V divide entre 120 V' }); });
      cmpN(h, 'Icc disponible (kA)', C.icc, r.iccKA, 0.05);
      { const e = cl(C.sistema); if (e && e !== cl(t.sistema)) { const vmax = Math.max(0, ...e.split('/').map(Number).filter(x => x)); dif(h, 'Sistema (tensión nominal)', txt(C.sistema), t.sistema, vmax && vmax !== r.V ? { nivel: 'info', nota: 'El texto de la ficha no coincide con el voltaje del tablero (' + r.V + ' V); la herramienta lo deduce del voltaje' } : {}); } }
      { const e = num(C.fases); if (e !== null) { const ok = e === r.fases || (e === 1 && (r.fases === 2 || r.fases === 3)); if (!ok) dif(h, 'Fases', e, r.fases); else if (e !== r.fases) dif(h, 'Fases', e, r.fases, { nivel: 'info', nota: '«Fases = 1» en un tablero de 120/240 V: la herramienta usa ' + r.fases + ' según las fases con carga' }); } }
      cmpN(h, 'Hilos', C.hilos, r.hilos, 0.1, { nivel: 'info', nota: 'Los monofásicos de 2 fases se importan con 3 hilos' });
      cmpT(h, 'Modelo del tablero', C.modelo, cat.modelo); cmpT(h, 'Fabricante', C.fabricante, cat.fabricante); cmpT(h, 'Montaje', C.montaje, t.montaje);
      cmpN(h, 'Barras de fases (A)', C.barraFase, cat.barraFase, 0.1); cmpN(h, 'Barras de neutro (A)', C.barraNeutro, cat.barraNeutro, 0.1); cmpN(h, 'Barras de tierra (A)', C.barraTierra, cat.barraTierra, 0.1);
      cmpN(h, 'Espacios', C.espacios, cat.espacios, 0.1);
      // alimentador
      const AL = b.alim, f = inf.cab;
      cmpCable(f, 'Alimentador: fases', AL.fases, A.fasesTxt); cmpCable(f, 'Alimentador: neutro', AL.neutro, A.neutroTxt); cmpCable(f, 'Alimentador: tierra', AL.tierra, A.tierraTxt);
      cmpT(f, 'Alimentador: material', AL.material, A.mat); cmpT(f, 'Alimentador: aislamiento', AL.aislamiento, A.ais);
      { const e = cable(AL.tuberia); if (e && num(e.cal) !== null) { const w = Number(A.AR139); if (Math.abs(num(e.cal) - w) > 0.2 || e.par !== (Number(A.AI139) || 1)) dif(f, 'Alimentador: tubería (mm)', txt(AL.tuberia), (Number(A.AI139) > 1 ? A.AI139 + 'x' : '') + fmt(w)); } }
      cmpN(f, 'Alimentador: longitud (m)', AL.longitud, A.M139, 0.05); cmpN(f, 'Alimentador: voltaje en bornes (V)', AL.voltaje, A.AW139, 0.5, { nivel: 'info', nota: 'La herramienta suma la caída de los tableros de arriba y del transformador' });
      cmpN(f, 'Alimentador: caída de voltaje (V)', AL.caidaV, A.AV139, 0.3, { nivel: 'info' }); cmpN(f, 'Alimentador: caída de voltaje total (%)', AL.caidaPct, A.AY139, 0.15, { nivel: 'info', nota: 'La herramienta acumula la caída de los tableros de arriba; el Excel viejo solo la del propio alimentador' });
      // interruptor principal y supresor
      const bk = r.bkMain || {};
      cmpModelo(f, 'Interruptor principal: modelo', b.bk.modelo, bk.ref || bk.modelo); cmpN(f, 'Interruptor principal: amperaje (A)', b.bk.amperaje, A.AF139, 0.5);
      cmpT(f, 'Interruptor principal: tipo', b.bk.tipo, bk.unidad === 'STD' || !bk.unidad ? 'STD' : bk.unidad); cmpN(f, 'Interruptor principal: polos', b.bk.polos, bk.polos, 0.1); cmpN(f, 'Interruptor principal: SCCR (kA)', b.bk.sccr, bk.sccr, 0.5);
      const sp = r.spd || {};
      cmpT(f, 'Supresor (SPD): modelo', b.spd.modelo, sp.modelo || sp.ref); cmpN(f, 'Supresor (SPD): capacidad (kA)', b.spd.capacidad, sp.kaLL, 0.5); cmpT(f, 'Supresor (SPD): montaje', b.spd.montaje, sp.montaje);
      // circuitos: por polos si la ficha los trae, si no por descripción y orden
      const circs = t.circuitos.filter(c => !c.respaldoDe), filaWeb = c => r.rows.find(z => z.c === c), descWeb = c => cl(c.descripcion || (filaWeb(c) || {}).descripcion);
      const mismosPolos = (a, c) => a.length && a.length === (c.polos || []).length && a.every(x => (c.polos || []).map(Number).includes(x));
      // emparejado de los circuitos de la ficha con los de la herramienta: por polos (si la ficha los trae), por descripción y kVA, por descripción y, si sobran los mismos de cada lado, por orden
      const par = new Array(b.circuitos.length).fill(null), libres = circs.slice();
      const tomar = (k, pred) => { if (par[k]) return; const c = libres.find(pred); if (c) { par[k] = c; libres.splice(libres.indexOf(c), 1); } };
      b.circuitos.forEach((sc, k) => { if (sc.polos.length) tomar(k, x => mismosPolos(sc.polos, x)); });
      b.circuitos.forEach((sc, k) => { if (cl(sc.desc)) tomar(k, x => descWeb(x) === cl(sc.desc) && Math.abs(Number((filaWeb(x) || {}).J || 0) - Number(sc.kva || 0)) < 0.011); });
      b.circuitos.forEach((sc, k) => { if (cl(sc.desc)) tomar(k, x => descWeb(x) === cl(sc.desc)); });
      { const falta = b.circuitos.map((_, k) => k).filter(k => !par[k]); if (falta.length === libres.length) falta.forEach((k, m) => { par[k] = libres[m]; }); else if (falta.length) falta.forEach(k => tomar(k, x => Math.abs(Number((filaWeb(x) || {}).J || 0) - Number(b.circuitos[k].kva || 0)) < 0.011)); }
      b.circuitos.forEach((sc, idx) => {
        const c = par[idx];
        const ci1 = { id: sc.id, desc: sc.desc, fila: sc.fila, difs: [] };
        inf.circ.push(ci1);
        if (!c) { ci1.sinPar = true; return; }
        const x = filaWeb(c); if (!x) { ci1.sinPar = true; return; }
        const d = ci1.difs;
        ci1.polos = (c.polos || []).join(',');
        cmpN(d, 'kVA conectados', sc.kva, x.J, 0.01);
        cmpT(d, 'Descripción', sc.desc, x.descripcion || c.descripcion || '');
        const bb = x.breaker || {};
        cmpModelo(d, 'Interruptor: modelo', sc.bk.modelo === '0' ? '' : sc.bk.modelo, bb.ref || bb.modelo); cmpN(d, 'Interruptor: amperios (A)', sc.bk.amperios, x.AF, 0.5);
        cmpT(d, 'Interruptor: tipo de unidad', sc.bk.tipo, bb.unidad === 'STD' || !bb.unidad ? 'STD' : bb.unidad); cmpN(d, 'Interruptor: polos', sc.bk.polos, x.polosBreaker || bb.polos, 0.1); cmpN(d, 'Interruptor: SCCR (kA)', sc.bk.sccr, bb.sccr, 0.5);
        cmpCable(d, 'Cable: fases', sc.cab.fases, x.fasesTxt || x.AL); cmpCable(d, 'Cable: neutro', sc.cab.neutro, x.neutroTxt || x.AN); cmpCable(d, 'Cable: tierra', sc.cab.tierra, x.tierraTxt || x.AP);
        cmpT(d, 'Cable: tipo', sc.cab.tipo, x.mat); cmpT(d, 'Cable: aislamiento', sc.cab.aislamiento, x.ais); cmpN(d, 'Tubo EMT (mm)', sc.cab.tubo, x.AR, 0.2);
        cmpN(d, 'Caída de voltaje total (%)', sc.caida, x.AY, 0.15, { nivel: 'info', nota: 'La herramienta incluye la caída acumulada de arriba' });
      });
      if (b.circuitos.length !== circs.length) dif(inf.cab, 'Número de circuitos', b.circuitos.length, circs.length);
      libres.forEach(c => { const x = filaWeb(c); if (x && x.J) inf.circ.push({ id: '—', desc: x.descripcion || c.descripcion, polos: (c.polos || []).join(','), difs: [], soloWeb: true }); });
      const n = lista => lista.filter(d => d.nivel === 'dato').length;
      inf.nDato = n(inf.cab) + inf.circ.reduce((a, c) => a + n(c.difs) + (c.sinPar || c.soloWeb ? 1 : 0), 0);
      const todos = inf.cab.concat(...inf.circ.map(c => c.difs)).filter(d => d.nivel === 'dato'), sinP = inf.circ.filter(c => c.sinPar || c.soloWeb).length;
      inf.nAlta = todos.filter(d => d.prio === 'alta').length + sinP; inf.nMedia = todos.filter(d => d.prio === 'media').length; inf.nBaja = inf.cab.concat(...inf.circ.map(c => c.difs)).filter(d => d.prio === 'baja').length;   // las notas (caídas de voltaje y demás) cuentan como baja
      inf.nInfo = inf.cab.filter(d => d.nivel === 'info').length + inf.circ.reduce((a, c) => a + c.difs.filter(d => d.nivel === 'info').length, 0);
      inf.nCircFicha = b.circuitos.length; inf.nCircWeb = circs.length;
      inf.estado = inf.nDato ? 'revisar' : inf.nInfo ? 'coincide (con notas)' : 'coincide';
    });
    const nombresWeb = new Set(P.tableros.map(t => limpiaNombre(txt(t.nombre).replace(/ \(\d+\)$/, ''))));
    const soloExcel = fichas.filter(b => !usadas.has(b) && !nombresWeb.has(limpiaNombre(b.nombre))).map(b => ({ nombre: b.nombre, hoja: b.hoja, fila: b.fila }));
    tableros.sort((a, b) => (b.nAlta - a.nAlta) || (b.nMedia - a.nMedia) || (b.nBaja - a.nBaja));   // los más críticos primero
    const criticos = [];
    tableros.forEach(t => {
      if (t.estado === 'sin ficha') { criticos.push({ tablero: t.nombre, seccion: 'Tablero', dato: 'Sin ficha en las pestañas 3F / 1F', excel: '', web: '' }); return; }
      t.cab.filter(d => d.prio === 'alta' && d.nivel === 'dato').forEach(d => criticos.push({ tablero: t.nombre, seccion: 'Tablero / alimentador', dato: d.campo, excel: d.excel, web: d.web }));
      t.circ.forEach(c => { if (c.sinPar) criticos.push({ tablero: t.nombre, seccion: 'Circuito ' + c.id + ' ' + (c.desc || ''), dato: 'No se encontró en la herramienta', excel: '', web: '' }); else if (c.soloWeb) criticos.push({ tablero: t.nombre, seccion: 'Circuito ' + (c.desc || ''), dato: 'No está en la ficha del Excel', excel: '', web: '' });
        c.difs.filter(d => d.prio === 'alta' && d.nivel === 'dato').forEach(d => criticos.push({ tablero: t.nombre, seccion: 'Circuito ' + c.id + ' ' + (c.desc || ''), dato: d.campo, excel: d.excel, web: d.web })); });
    });
    const resumen = { tableros: tableros.length, coinciden: tableros.filter(x => /^coincide/.test(x.estado)).length, revisar: tableros.filter(x => x.estado === 'revisar').length, sinFicha: tableros.filter(x => x.estado === 'sin ficha').length, soloExcel: soloExcel.length, fichas: fichas.length };
    resumen.alta = tableros.reduce((a, t) => a + t.nAlta, 0); resumen.media = tableros.reduce((a, t) => a + t.nMedia, 0); resumen.baja = tableros.reduce((a, t) => a + t.nBaja, 0);
    return { archivo, fecha: new Date().toISOString().slice(0, 16).replace('T', ' '), tableros, soloExcel, criticos, resumen, hayFichas: fichas.length > 0 };
  }

  /* ---------- reporte ---------- */
  const porPrio = ds => ds.slice().sort((a, b) => (ORD[a.prio] - ORD[b.prio]) || (a.nivel === b.nivel ? 0 : a.nivel === 'dato' ? -1 : 1));
  const prioCirc = c => (c.sinPar || c.soloWeb ? 0 : Math.min(3, ...c.difs.filter(d => d.nivel === 'dato').map(d => ORD[d.prio])));
  const COL = { alta: ['#b42318', '#fdecea'], media: ['#8a5a00', '#fff4d6'], baja: ['#64707d', '#eef1f5'] };
  function filasCsv(inf) {
    const o = [['Prioridad', 'Tablero', 'Pestaña', 'Sección', 'Circuito', 'Dato', 'Excel', 'Herramienta', 'Nota']];
    inf.tableros.forEach(t => {
      if (t.estado === 'sin ficha') { o.push(['Alta', t.nombre, '', 'Tablero', '', 'Sin ficha en las pestañas 3F / 1F', '', '', '']); return; }
      porPrio(t.cab).forEach(d => o.push([ETQ_PRIO[d.prio], t.nombre, t.hoja, 'Tablero y alimentador', '', d.campo, fmt(d.excel), fmt(d.web), d.nota]));
      t.circ.slice().sort((a, b) => prioCirc(a) - prioCirc(b)).forEach(c => { if (c.sinPar) o.push(['Alta', t.nombre, t.hoja, 'Circuito', c.id + ' ' + c.desc, 'No se encontró en la herramienta', '', '', '']); else if (c.soloWeb) o.push(['Alta', t.nombre, t.hoja, 'Circuito', c.desc, 'No está en la ficha del Excel', '', '', '']); porPrio(c.difs).forEach(d => o.push([ETQ_PRIO[d.prio], t.nombre, t.hoja, 'Circuito', c.id + ' ' + c.desc, d.campo, fmt(d.excel), fmt(d.web), d.nota])); });
    });
    inf.soloExcel.forEach(x => o.push(['Alta', x.nombre, x.hoja, 'Ficha', '', 'Hay ficha en el Excel y no hay tablero en la herramienta', '', '', '']));
    return o;
  }
  const esc = x => String(x === undefined || x === null ? '' : x).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const badge = pr => '<span class="p p-' + pr + '">' + ETQ_PRIO[pr] + '</span>';
  function html(inf, abrir) {
    const fila = d => '<tr class="' + d.nivel + '"><td>' + esc(d.campo) + '</td><td class="r">' + badge(d.prio) + '</td><td class="r">' + esc(fmt(d.excel)) + '</td><td class="r">' + esc(fmt(d.web)) + '</td><td>' + esc(d.nota || '') + '</td></tr>';
    const cab = '<colgroup><col class="a"><col class="p1"><col class="b"><col class="c"><col class="d"></colgroup><thead><tr><th>Dato</th><th class="r">Prioridad</th><th class="r">Excel</th><th class="r">Herramienta</th><th>Nota</th></tr></thead>';
    const tabla = ds => ds.length ? '<table>' + cab + '<tbody>' + porPrio(ds).map(fila).join('') + '</tbody></table>' : '';
    const T = inf.tableros.map(t => '<details' + (abrir ? ' open' : '') + '><summary><b>' + esc(t.nombre) + '</b> · ' + esc(t.hoja ? 'pestaña ' + t.hoja + ', fila ' + t.fila : '') + ' · <span class="e-' + esc(t.estado.split(' ')[0]) + '">' + esc(t.estado) + '</span> · ' + badge('alta') + ' ' + t.nAlta + ' ' + badge('media') + ' ' + t.nMedia + ' ' + badge('baja') + ' ' + t.nBaja
      + (t.nCircFicha !== undefined ? ' · circuitos: ' + t.nCircFicha + ' en la ficha, ' + t.nCircWeb + ' en la herramienta' : '') + '</summary>'
      + tabla(t.cab) + t.circ.filter(c => c.sinPar || c.soloWeb || c.difs.some(d => d.nivel === 'dato')).sort((a, b) => prioCirc(a) - prioCirc(b)).map(c => '<h4>Circuito ' + esc(c.id + ' ' + (c.desc || '')) + (c.polos ? ' [' + esc(c.polos) + ']' : '') + (c.sinPar ? ' — no se encontró en la herramienta' : c.soloWeb ? ' — no está en la ficha del Excel' : '') + '</h4>' + tabla(c.difs)).join('') + '</details>').join('');
    const K = inf.criticos.length ? '<table><colgroup><col style="width:16%"><col style="width:30%"><col style="width:28%"><col style="width:13%"><col style="width:13%"></colgroup><thead><tr><th>Tablero</th><th>Sección</th><th>Dato</th><th class="r">Excel</th><th class="r">Herramienta</th></tr></thead><tbody>' + inf.criticos.map(k => '<tr><td>' + esc(k.tablero) + '</td><td>' + esc(k.seccion) + '</td><td>' + esc(k.dato) + '</td><td class="r">' + esc(fmt(k.excel)) + '</td><td class="r">' + esc(fmt(k.web)) + '</td></tr>').join('') + '</tbody></table>' : '<p>Sin diferencias críticas.</p>';
    return '<!doctype html><html lang="es"><meta charset="utf-8"><title>Revisión de ' + esc(inf.archivo) + '</title><style>@page{margin:14mm}@media print{details{break-inside:avoid-page}h4{break-after:avoid}}body{font:13px system-ui,Segoe UI,Arial;margin:20px;color:#1d2733}table{border-collapse:collapse;margin:6px 0 12px;table-layout:fixed;width:100%}col.a{width:30%}col.p1{width:10%}col.b,col.c{width:15%}col.d{width:30%}td,th{border:1px solid #ccc;padding:2px 8px;text-align:left}td.r,th.r{text-align:center}tr.info td{color:#64707d}summary{cursor:pointer;margin:8px 0}h4{margin:10px 0 2px}.e-revisar,.e-sin{color:#b42318;font-weight:600}.e-coincide{color:#1b7f4b;font-weight:600}.p{display:inline-block;padding:0 7px;border-radius:9px;font-size:11px;font-weight:600}.p-alta{background:' + COL.alta[1] + ';color:' + COL.alta[0] + '}.p-media{background:' + COL.media[1] + ';color:' + COL.media[0] + '}.p-baja{background:' + COL.baja[1] + ';color:' + COL.baja[0] + '}</style><body><h2>Revisión del proyecto importado contra las pestañas TABLEROS 3F y 1F</h2><p>Archivo: ' + esc(inf.archivo) + ' · ' + esc(inf.fecha) + '<br>Tableros revisados: ' + inf.resumen.tableros + ' · coinciden: ' + inf.resumen.coinciden + ' · para revisar: ' + inf.resumen.revisar + ' · sin ficha: ' + inf.resumen.sinFicha + ' · fichas del Excel sin tablero: ' + inf.resumen.soloExcel + '<br>Diferencias: ' + badge('alta') + ' ' + inf.resumen.alta + ' ' + badge('media') + ' ' + inf.resumen.media + ' ' + badge('baja') + ' ' + inf.resumen.baja + '</p>'
      + '<h3>Críticos (prioridad alta): ' + inf.criticos.length + '</h3>' + K + '<h3>Tableros, del más crítico al menos crítico</h3>' + T + (inf.soloExcel.length ? '<h3>Fichas del Excel sin tablero en la herramienta</h3><ul>' + inf.soloExcel.map(x => '<li>' + esc(x.nombre) + ' (pestaña ' + esc(x.hoja) + ', fila ' + x.fila + ')</li>').join('') + '</ul>' : '')
      + '<p><b>Prioridad.</b> Alta (críticos): cambió el cableado, el tamaño de la protección, las distancias, la descripción, el tipo de unidad (GFCI / AFCI), los factores de demanda o de diversidad o el número de circuitos. Media: cambió el modelo de las protecciones, el SCCR o el supresor (y otros datos de equipo o de cálculo). Baja: caídas de voltaje, montaje, hilos y notas. Las filas en gris son notas: diferencias esperadas por la forma de calcular de la herramienta (caída de voltaje acumulada, 120/240 V).</p></body></html>';
  }
  /** Abre el cuadro de impresión con el reporte completo (todos los tableros desplegados): ahí se elige «Guardar como PDF». */
  function pdf(inf) {
    const f = document.createElement('iframe'); f.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0'; document.body.append(f);
    const d = f.contentWindow.document; d.open(); d.write(html(inf, true)); d.close();
    setTimeout(() => { f.contentWindow.focus(); f.contentWindow.print(); setTimeout(() => f.remove(), 60000); }, 300);
  }

  function mostrar(inf) {
    const h = U.h, R = inf.resumen;
    if (!inf.hayFichas) { UI.alert('El Excel no trae fichas en las pestañas «TABLEROS 3F» / «TABLEROS 1F» (o no se pudieron leer), así que no hay contra qué comparar.', 'Revisión'); return; }
    const chipP = (pr, n) => h('span', { style: 'display:inline-block;padding:1px 8px;border-radius:10px;font-size:12px;font-weight:600;background:' + COL[pr][1] + ';color:' + COL[pr][0] }, ETQ_PRIO[pr] + (n !== undefined ? ' ' + n : ''));
    const chip = e => h('span', { style: 'padding:1px 8px;border-radius:10px;font-size:12px;background:' + (/^coincide/.test(e) ? 'var(--ok-bg)' : 'var(--bad-bg)') + ';color:' + (/^coincide/.test(e) ? 'var(--ok)' : 'var(--bad)') }, e);
    const W = ['30%', '9%', '15%', '15%', '31%'], cen = 'text-align:center;white-space:nowrap';   // las mismas columnas en todas las tablas, para que Excel y Herramienta queden alineadas
    const tabla = ds => (ds.length ? h('div', { class: 'tbl-wrap' }, h('table', { class: 'tbl', style: 'table-layout:fixed;width:100%' }, h('colgroup', null, W.map(w => h('col', { style: 'width:' + w }))),
      h('thead', null, h('tr', null, ['Dato', 'Prioridad', 'Excel', 'Herramienta', 'Nota'].map((x, i) => h('th', { style: i >= 1 && i <= 3 ? cen : null }, x)))),
      h('tbody', null, porPrio(ds).map(d => h('tr', { style: d.nivel === 'info' ? 'opacity:.65' : null, title: d.nota || null }, h('td', null, d.campo), h('td', { style: cen }, chipP(d.prio)), h('td', { style: cen }, fmt(d.excel)), h('td', { style: cen }, fmt(d.web)), h('td', { class: 'muted', style: 'font-size:12px' }, d.nota || '')))))) : null);
    const crit = inf.criticos.length ? h('div', { class: 'tbl-wrap', style: 'max-height:220px;overflow:auto' }, h('table', { class: 'tbl', style: 'table-layout:fixed;width:100%' }, h('colgroup', null, ['16%', '29%', '27%', '14%', '14%'].map(w => h('col', { style: 'width:' + w }))),
      h('thead', null, h('tr', null, ['Tablero', 'Sección', 'Dato', 'Excel', 'Herramienta'].map((x, i) => h('th', { style: i >= 3 ? cen : null }, x)))),
      h('tbody', null, inf.criticos.map(k => h('tr', null, h('td', null, k.tablero), h('td', { style: 'font-size:12px' }, k.seccion), h('td', null, k.dato), h('td', { style: cen }, fmt(k.excel)), h('td', { style: cen }, fmt(k.web))))))) : h('p', { class: 'muted' }, 'Sin diferencias críticas.');
    const cuerpo = inf.tableros.map(t => h('details', null, h('summary', null, h('b', null, t.nombre), ' ', chip(t.estado), ' ', chipP('alta', t.nAlta), ' ', chipP('media', t.nMedia), ' ', chipP('baja', t.nBaja), ' ',
      h('span', { class: 'muted' }, t.hoja ? 'pestaña ' + t.hoja + ', fila ' + t.fila + (t.nCircFicha !== undefined ? ' · circuitos ' + t.nCircFicha + ' (ficha) / ' + t.nCircWeb + ' (herramienta)' : '') : 'no hay ficha con ese nombre')),
      tabla(t.cab), t.circ.filter(c => c.sinPar || c.soloWeb || c.difs.some(d => d.nivel === 'dato')).sort((a, b) => prioCirc(a) - prioCirc(b)).map(c => h('div', null, h('h4', null, 'Circuito ' + c.id + ' ' + (c.desc || '') + (c.polos ? ' [' + c.polos + ']' : '') + (c.sinPar ? ' — no se encontró en la herramienta' : c.soloWeb ? ' — no está en la ficha del Excel' : '')), tabla(c.difs)))));
    UI.modal('Revisión contra las pestañas 3F / 1F', h('div', null,
      h('p', null, R.tableros + ' tablero(s) revisados: ', h('b', null, R.coinciden + ' coinciden'), ', ', h('b', null, R.revisar + ' para revisar'), R.sinFicha ? ', ' + R.sinFicha + ' sin ficha' : '', R.soloExcel ? ', ' + R.soloExcel + ' ficha(s) del Excel sin tablero' : '', '. Diferencias: ', chipP('alta', R.alta), ' ', chipP('media', R.media), ' ', chipP('baja', R.baja)),
      h('p', { class: 'muted' }, 'Alta (críticos): cableado, tamaño de la protección, distancias, descripción, tipo de unidad (GFCI / AFCI), factores de demanda y diversidad, número de circuitos. Media: modelo de las protecciones, SCCR, supresor. Baja: caídas de voltaje, montaje, hilos y notas (las filas atenuadas son diferencias esperadas por cómo calcula la herramienta).'),
      h('div', { style: 'max-height:62vh;overflow:auto' },
        h('h4', null, 'Críticos (prioridad alta): ' + inf.criticos.length), crit,
        h('h4', null, 'Tableros, del más crítico al menos crítico'), cuerpo, inf.soloExcel.length ? h('div', null, h('h4', null, 'Fichas del Excel sin tablero en la herramienta'), h('ul', null, inf.soloExcel.map(x => h('li', null, x.nombre + ' (pestaña ' + x.hoja + ', fila ' + x.fila + ')')))) : null)),
    [
      { label: 'Descargar reporte (.html)', onclick: () => { U.download('Revisión ' + inf.archivo.replace(/\.[^.]+$/, '') + '.html', html(inf), 'text/html'); return false; } },
      { label: 'Exportar PDF', cls: 'primary', onclick: () => { pdf(inf); return false; } },
      { label: 'Descargar (.csv)', onclick: () => { U.download('Revisión ' + inf.archivo.replace(/\.[^.]+$/, '') + '.csv', '\ufeff' + U.csv(filasCsv(inf)), 'text/csv'); return false; } },
      { label: 'Cerrar' },
    ], { wide: true });
  }

  /** Revisa el proyecto abierto contra las fichas de un Excel (sin importarlo de nuevo). */
  async function revisarArchivo(buf, nombre) {
    const sheets = await XlsxRead.read(buf, { hojas: ['TABLEROS 3F', 'TABLEROS 1F'] });
    return revisar(sheets, Store.project, Calc.proyecto(Store.project, Store.catalog), nombre);
  }
  g.RevisorAnacar = { revisar, revisarArchivo, mostrar, leerFichas, html, filasCsv, pdf };
})(window);
