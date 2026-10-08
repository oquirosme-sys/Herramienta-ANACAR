/* Motor de cálculo (sin DOM). Replica las fórmulas de la hoja ANACAR (bloque "Machote") y agrega la cascada entre tableros:
   - la carga de un tablero derivado sube al circuito del tablero que lo alimenta,
   - el voltaje en bornes de un tablero es el punto de partida del alimentador de sus derivados (caída de voltaje acumulada),
   - la corriente de cortocircuito baja por cada alimentador (método punto a punto).
   Referencias a celdas del Excel entre corchetes: [AE16] = ampacidad requerida del circuito, [AW139] = voltaje en bornes, etc. */
(function (g) {
  'use strict';
  let K = null; // catálogo activo

  const n = v => (v === null || v === undefined || v === '' || isNaN(Number(v)) ? 0 : Number(v));
  const r2 = v => Math.round(v * 100) / 100;

  /* ---------- búsquedas en tablas (equivalentes a VLOOKUP/INDEX-MATCH del Excel) ---------- */
  /** VLOOKUP(x, tabla, , TRUE): fila con el mayor umbral <= x. */
  function approx(tabla, campo, x) {
    let res = null;
    for (const f of tabla) { if (Number(f[campo]) <= x) res = f; else break; }
    return res;
  }
  /** MATCH(TRUE, x <= umbrales, 0): primera fila con umbral >= x. */
  function firstGE(tabla, campo, x) { return tabla.find(f => Number(f[campo]) >= x) || null; }
  const calStr = c => (c === null || c === undefined ? '' : String(c));

  const esAL = m => m === 'AL' || m === 'AL-MC';
  const esMC = m => m === 'CU-MC' || m === 'AL-MC';

  /** Protección (breaker) estándar para una ampacidad requerida [CONDISEÑO L124:M154]. */
  const proteccion = amp => { const f = approx(K.protecciones, 'amp', amp); return f ? f.prot : null; };
  /** Calibre por ampacidad [CONDISEÑO C6:D26 CU / C44:D64 AL]. */
  const calibre = (amp, mat) => { const f = approx(esAL(mat) ? K.ampAL : K.ampCU, 'amp', amp); return f ? f.cal : ''; };
  /** Tierra de equipo, tabla 250.122 [CONDISEÑO L76:P98]. */
  const tierra = (prot, mat) => { const f = firstGE(K.tierras, 'amp', prot); return f ? (f[mat] || f.CU) : ''; };
  /** Conductor del electrodo, tabla 250.66 [CONDISEÑO U76:Z95]. */
  function electrodo(cal, mat) {
    const al = esAL(mat), f = K.electrodo.find(x => (al ? x.calAL : x.calCU) === calStr(cal));
    return f ? (f[mat] || (al ? f.AL : f.CU)) : '';
  }
  /** Diámetro de tubería, tabla C.10 [CONDISEÑO L101:O118]. Igual que el Excel: "XHHW-2" usa la columna THHN. */
  function conduit(cal, mat, ais) {
    if (esMC(mat)) return 'MC';
    const f = K.conduit.find(x => x.cal === calStr(cal)); if (!f) return '';
    return ais === 'RHW' ? f.RHW : ais === 'XHHW' ? f.XHHW : f.THHN;
  }
  /** FAC de caída de tensión por calibre y factor de potencia [CONDISEÑO H5:AC9 CU / H21:AC25 AL, conduit magnético]. */
  function fac(cal, mat, fp) {
    const t = K.fac[esAL(mat) ? 'AL_mag' : 'CU_mag'];
    const fila = fp === 1 ? '1' : fp > 0.949 ? '0.95' : fp > 0.899 ? '0.9' : fp > 0.799 ? '0.8' : null;
    if (!fila || !t[fila]) return null;
    const v = t[fila][calStr(cal)];
    return v === undefined ? null : v;
  }
  /** FAC ajustado [AT16]: L-N y L-L monofásicos. */
  function facAjustado(f, O, P, extra240x2) {
    if (f === null) return null;
    if (P === 1 && (O === 120 || O === 230 || O === 277)) return f * 1.1547;
    if (P === 1 && (O === 208 || O === 240)) return f * 1.15;
    if (extra240x2 && P === 2 && O === 240) return f * 1.15;
    return f;
  }
  /** Voltaje línea-neutro del sistema [AU126:AV134, XLOOKUP exacto o siguiente mayor]. */
  function vLN(V) {
    const t = K.listas.lnVoltaje.slice().sort((a, b) => a[0] - b[0]);
    const f = t.find(x => x[0] >= V); return f ? f[1] : V / 1.732;
  }
  /** Corriente por fase [AA16]: mismas condiciones que el Excel. U = kVA conectados en la fase, X = kVA demandados. */
  function corriente(U, X, O, P) {
    if (!O) return 0;
    if (O === 120 && P === 1) return X * 1000 / O;
    if (O === 208 && P === 1) return X * 1000 / O;
    if (O === 208 && P === 2) return U * 2 * 1000 / O;
    if (O === 230 && P === 1) return X * 1000 / O;
    if (O === 240 && P === 1) return X * 1000 / (O / 2);
    if (O === 277 && P === 1) return X * 1000 / O;
    return X * 1000 / (O / 1.732);
  }
  /** Corriente del alimentador [AA139]. */
  function corrienteAlim(X, O, P) {
    if (O === 120 && P === 1) return X * 1000 / O;
    if (O === 208 && P === 1) return X * 1000 / O;
    if (O === 240 && P === 2) return X * 1000 / (O / 2);
    if (O === 230 && P === 1) return X * 1000 / O;
    if (O === 240 && P === 1) return X * 1000 / (O / 2);
    if (O === 277 && P === 1) return X * 1000 / O;
    return X * 1000 / (O / 1.732);
  }
  /** Conductores en paralelo automáticos [AI17]. */
  const paralelosAuto = amp => (amp > 1600 ? 5 : amp > 1000 ? 4 : amp > 600 ? 3 : amp > 300 ? 2 : 1);

  /** Fase de un polo (funciones VBA POLA/POLB/POLC y POLD/POLE): 3F → A,B,C cada 2 polos; 1F → A,B. */
  function fasePolo(polo, fasesTablero) {
    const k = Math.floor((polo - 1) / 2);
    if (fasesTablero === 3) return k % 3;
    if (fasesTablero === 2) return k % 2;
    return 0;
  }

  /* ---------- catálogo y reglas de fabricante: breakers, supresores, tableros ---------- */
  const rango = s => { const m = String(s || '').match(/(\d+)\s*-\s*(\d+)/); if (m) return [Number(m[1]), Number(m[2])]; const v = Number(s); return v ? [v, v] : null; };
  const lista = s => String(s || '').split(/[,;]/).map(x => x.trim()).filter(Boolean);
  const reglasB = () => ((K.reglas || {}).breakers || []);
  const reglasT = () => ((K.reglas || {}).tableros || []);
  /** Familia de un breaker según las reglas del fabricante (QOB, EDB, H, PDG2…). */
  function famBreaker(b) {
    const m = String(b.modelo || '').trim();
    return reglasB().find(f => f.marca === b.marca && new RegExp(f.patron, 'i').test(m)) || null;
  }
  /** Familia de un tablero de catálogo (NQ, NF, I-Line, CH, PRL1…). */
  function famTablero(t) { return reglasT().find(f => f.marca === t.fabricante && new RegExp(f.patron, 'i').test(String(t.modelo || ''))) || null; }
  const fasesOk = (f, tipo) => !f || !f.fases || f.fases.split('/').includes(tipo);
  /** SCCR del breaker al voltaje del sistema: las filas duplicadas del Excel son la capacidad a 240 V y a 480 V. */
  const vSccr = b => n(b.vSccr) || ((famBreaker(b) || {}).vMax <= 240 ? 240 : 480);
  /** Número de catálogo a partir de la plantilla de la familia ({p} polos, {a} amperios, {a3} amperios a 3 dígitos, {m3} 3 letras del modelo). */
  function modeloRef(b, polos, amp) {
    const f = famBreaker(b);
    if (f && f.plantilla === '{pdg}' && amp) {
      // Power Defense: PDG + marco + polos + letra de capacidad + amperios (4 dígitos) + disparo + terminales J
      const m = /^PDG(\d)\d([A-Z])/.exec(String(b.modelo || '')); if (!m) return b.modelo || '';
      const px = /PXR(\d+)/.exec(b.modelo || '');
      const trip = px ? (px[1] === '10' ? 'B' : px[1] === '25' ? 'P' : 'E') + (/G$/.test(b.unidad || '') ? '3' : '2') + 'N' : (Number(m[1]) <= 2 ? 'TFF' : 'TFA');
      return 'PDG' + m[1] + polos + m[2] + String(amp).padStart(4, '0') + trip + 'J';
    }
    if (f && /=/.test(f.plantilla || '') && amp) {   // plantilla por tipo de unidad: GFCI=…;AFCI=…
      const m = f.plantilla.split(';').map(x => x.split('=')).find(x => x[0].trim() === (b.unidad || ''));
      if (m) return m[1].trim().replace('{p}', polos).replace('{a3}', String(amp).padStart(3, '0')).replace('{a}', amp);
    }
    if (!f || !f.plantilla || (b.unidad && b.unidad !== 'STD' && !/\{u\}/.test(f.plantilla)) || !amp) return b.modelo || '';
    return f.plantilla.replace('{p}', polos).replace('{a3}', String(amp).padStart(3, '0')).replace('{a}', amp).replace('{m3}', String(b.modelo || '').slice(0, 3)).replace('{u}', b.unidad || '');
  }
  /** Elige el breaker: familias permitidas por el tablero, voltaje, polos, rango de amperios, tipo de unidad y SCCR ≥ Icc (el menor que cumpla). */
  function breaker(marca, id, polos, amp, opt) {
    opt = opt || {};
    const todos = K.breakers.filter(b => b.marca === marca);
    const res = b => (b ? Object.assign({}, b, { fam: (famBreaker(b) || {}).familia || '', ref: modeloRef(b, polos, amp), sccrBajo: !!(opt.iccKA && n(b.sccr) && n(b.sccr) < opt.iccKA) }) : null);
    if (id) return res(todos.find(b => String(b.id) === String(id)));
    const V = n(opt.V) || 240, perm = lista(opt.familias), unidad = opt.unidad || 'STD';
    const okBase = b => { const r = rango(b.amperios), f = famBreaker(b); return Number(b.polos) === Number(polos) && (!r || (amp >= r[0] && amp <= r[1])) && (!f || n(f.vMax) >= V) && (!perm.length || (f && perm.includes(f.familia))); };
    // con 240 V se usa la capacidad a 240 V si existe; con 480 V solo la de 480 V
    let c = todos.filter(okBase).filter(b => (V > 240 ? vSccr(b) >= 480 : true));
    if (V <= 240) c = c.filter(b => !(vSccr(b) >= 480 && c.some(o => o !== b && o.modelo === b.modelo && o.polos === b.polos && o.amperios === b.amperios && o.unidad === b.unidad && vSccr(o) <= 240)));
    const porUnidad = c.filter(b => (b.unidad || 'STD') === unidad);
    if (porUnidad.length) c = porUnidad;
    if (!c.length && perm.length) return breaker(marca, '', polos, amp, Object.assign({}, opt, { familias: '' }));
    const ord = b => { const f = famBreaker(b); const i = f ? perm.indexOf(f.familia) : -1; return i < 0 ? 99 : i; };
    const ok = c.filter(b => !opt.iccKA || n(b.sccr) >= opt.iccKA).sort((a, b) => n(a.sccr) - n(b.sccr) || ord(a) - ord(b));
    if (ok.length) return res(ok[0]);
    return res(c.sort((a, b) => n(b.sccr) - n(a.sccr))[0] || null);
  }
  /** Clasificación en serie (NEC 240.86): combinación listada por el fabricante entre el principal (línea) y el ramal (carga). */
  function serie(bkCarga, bkLinea, ampLinea, V, iccKA) {
    if (!bkCarga || !bkLinea || !iccKA) return null;
    const tablas = (K.seriesRating || {}).combinaciones || [];
    const volt = V >= 440 ? ['480Y/277', '480'] : V === 240 ? ['120/240'] : ['208Y/120', '120/240'];
    // "PDG2xG" = PDG2 + polos + G; el resto son prefijos del modelo (BAB, GHB, HFD…)
    const tok = t => new RegExp('^' + t.trim().replace(/[^A-Za-z0-9-]/g, '').replace(/x/g, '\\d'), 'i');
    const ml = String(bkLinea.ref || bkLinea.modelo || ''), mc = String(bkCarga.modelo || bkCarga.ref || '');
    return tablas.filter(c => (!c.marca || c.marca === bkCarga.marca) && volt.includes(String(c.voltaje)) && n(c.kA) >= iccKA && (!n(c.principalMax) || ampLinea <= n(c.principalMax))
      && lista(c.linea).some(t => tok(t).test(ml)) && lista(c.carga).some(t => tok(t).test(mc)))
      .sort((a, b) => n(a.kA) - n(b.kA))[0] || null;
  }
  const normV = s => String(s || '').split('/').map(Number).filter(x => x).sort((a, b) => a - b).join('/');
  /** kA por fase recomendados para el SPD: acometida 250 kA, tableros de 600–1200 A 120 kA, hasta 400 A 50 kA. */
  const kaSpd = (servicio, barras) => (servicio ? 250 : barras >= 600 ? 120 : 50);
  function supresor(marca, id, sistema, fases, kaMin) {
    const l = K.supresores.filter(b => b.marca === marca);
    if (id) return l.find(b => String(b.id) === String(id)) || null;
    const fs = fases === 3 ? 3 : 1;
    const ok = l.filter(b => normV(b.voltaje) === normV(sistema) && Number(b.fases) === fs).sort((a, b) => n(a.kaLL) - n(b.kaLL));
    return ok.find(b => n(b.kaLL) >= n(kaMin)) || ok[ok.length - 1] || null;
  }
  /** Tablero de catálogo: familias del fabricante válidas para el voltaje y el tipo (1F/3F), barras ≥ protección principal,
      espacios ≥ espacios usados + reserva; la familia más sencilla primero (orden), luego barras y espacios. */
  function tableroCat(tab, prot, espaciosUsados, V, reservaEsp) {
    if (tab.catalogoId) return K.tablerosCat.find(t => String(t.id) === String(tab.catalogoId)) || null;
    const marca = tab.marca || K.marcaDefecto || 'Eaton', tipo = tab.tipo || (Number(tab.fases) === 3 ? '3F' : '1F');
    const nec = Math.max(n(tab.espacios), Math.ceil(espaciosUsados * (1 + (n(reservaEsp) / 100))));
    const sub = tab.clase === 'subestacion', princ = tab.principal === 'zapatas' ? 'Zapatas' : 'Interruptor';
    const cand = K.tablerosCat.filter(t => {
      if (t.fabricante !== marca || n(t.barraFase) < n(prot) || n(t.espacios) < nec) return false;
      const f = famTablero(t), clase = (f && f.clase) || 'tablero';
      if (f && n(f.vMax) < V) return false;
      if (sub ? clase === 'tablero' : clase === 'subestacion') return false;
      if (t.principal && t.principal !== 'Ambos' && t.principal !== princ) return false;
      return fasesOk(t.fases ? { fases: t.fases } : f, tipo);
    });
    const ord = t => (famTablero(t) || { orden: 99 }).orden;
    // con espacios elegidos se prefiere el modelo con exactamente esos espacios
    const exacto = t => (n(tab.espacios) && n(t.espacios) === n(tab.espacios) ? 0 : 1);
    return cand.sort((a, b) => exacto(a) - exacto(b) || ord(a) - ord(b) || n(a.barraFase) - n(b.barraFase) || n(a.espacios) - n(b.espacios))[0] || null;
  }

  /* ---------- validación de ampacidad (310.16 con factores de temperatura y agrupamiento) ---------- */
  const CALIBRES = ['14', '12', '10', '8', '6', '4', '3', '2', '1', '1/0', '2/0', '3/0', '4/0', '250', '300', '350', '400', '500', '600', '700', '750', '800', '900', '1000'];
  /** Temperatura del aislamiento: RHW, THW, TW son de 75 °C; THHN, XHHW-2, RHW-2 y MC de 90 °C. */
  const tempAis = ais => (/^(RHW|THW|TW|USE)$/i.test(String(ais || '').trim()) ? 75 : 90);
  function agrupAuto(fases, hilos) {
    const nc = (fases === 3 ? 3 : fases === 2 ? 2 : 1) + (fases === 3 && hilos === 4 ? 1 : 0);
    return nc <= 3 ? '1-3' : nc <= 6 ? '4-6' : nc <= 9 ? '7-9' : '10-20';
  }
  /** Comprueba un calibre: ampacidad corregida = ampacidad a la temperatura del aislamiento × Ft × Fg, limitada por la columna de los bornes. */
  function ampacidad(cal, mat, nPar, ais, tempAmb, tempBorne, agrup, Ireq, prot) {
    const f = (K.ampacidad31016 || []).find(x => x.cal === calStr(cal));
    if (!f) return null;
    const tA = tempAis(ais), m = mat || 'CU';
    const amp90 = n(f[m + '|' + tA]), term = n(f[m + '|' + (tempBorne || 75)]) || amp90;
    const ft = n(((K.tempFactor || {})[tempAmb || '26-30'] || {})[String(tA)]) || 0;
    const fg = n((K.agrupamiento || {})[agrup || '1-3']) || 1;
    const corr = amp90 * ft * fg, cap = Math.min(corr, term) * nPar;
    const std = (K.protecciones || []).map(x => n(x.prot)).sort((a, b) => a - b);
    const sig = std.find(x => x >= cap);
    const okI = cap >= Ireq - 1e-9, okP = !prot || prot <= cap + 1e-9 || (prot <= 800 && prot === sig);
    return { cal: calStr(cal), amp90, term, ft, fg, corr: corr * nPar, cap, tA, okI, okP, ok: okI && okP && ft > 0 };
  }

  /* ---------- NEC 2020: GFCI (210.8, 422.5, 680.21) y AFCI (210.12) según ubicación y ocupación ---------- */
  const normTxt = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  const LUG_GFCI = /bano|cocina|exterior|lavander|garaj|garage|sotano|piscina|jacuzzi|tina|ducha|bbq|terraza|patio|azotea|fregadero|mojad|humed|ice maker|maquina de hielo|vestidor/;
  const LUG_AFCI = /habitaci|dormitor|recamara|sala|comedor|estudio|pasillo|closet|cuarto|biblioteca|family|sunroom|lavander|cocina/;
  function necUnidad(desc, det, O, P, amp, fases, Vsis) {
    const ocup = K.ocupacion || 'comercial', d = normTxt(desc + ' ' + (det.descripcion || ''));
    const vTierra = Vsis >= 440 ? 277 : 120;                         // 120/208 y 120/240: 120 V a tierra; 277/480: 277 V
    if (vTierra > 150) return null;
    const toma = /toma|receptac|enchuf/.test(d) || [2, 3].includes(Number(det.tipo));
    let gfci = null, afci = null;
    if (/lavaplato|lavavajilla/.test(d)) gfci = '422.5: lavaplatos';
    else if (/piscina|jacuzzi|spa\b/.test(d) && amp <= 60) gfci = '680.21: equipo de piscina';
    else if (ocup === 'vivienda') {
      if (toma && LUG_GFCI.test(d)) gfci = '210.8(A): toma en ' + (d.match(LUG_GFCI) || [''])[0];
      else if (/exterior|condensad|unidad cu|uc-|aire|a\/c/.test(d) && amp <= 50) gfci = '210.8(F): salida exterior ≤ 50 A';
    } else if (toma && LUG_GFCI.test(d) && amp <= (P === 3 ? 100 : 50)) gfci = '210.8(B): toma en ' + (d.match(LUG_GFCI) || [''])[0];
    const af120 = O === 120 && amp <= 20 && (toma || /ilum|luz/.test(d));
    if (af120 && ((ocup === 'vivienda' && LUG_AFCI.test(d)) || ((ocup === 'hotel' || ocup === 'hospital') && /habitaci|dormitor/.test(d)))) afci = '210.12: ' + (d.match(LUG_AFCI) || ['habitación'])[0];
    if (gfci && afci) return { unidad: 'AFCI/GFCI', motivo: gfci + ' · ' + afci };
    if (gfci) return { unidad: 'GFCI', motivo: gfci };
    if (afci) return { unidad: 'AFCI', motivo: afci };
    return null;
  }

  /* ---------- un circuito ramal (fila 16 a 115 del Machote) ---------- */
  function circuito(c, tab, ctx) {
    const det = K.detallesCarga.find(d => String(d.id) === String(c.detalleId)) || {};
    const tiene = v => v !== undefined && v !== null && v !== '';
    const O = n(det.v), P = n(det.fases), Q = n(det.hilos), T = n(det.fp) || 1;
    const R = tiene(c.fu) ? n(c.fu) : (n(det.fd) || 1), S = tiene(c.fdivC) ? (n(c.fdivC) || 1) : (n(det.fdiv) || 1);   // [R16] factor de uso, [S16] diversidad
    const N = ctx.fases, J = n(c.kva);
    const polos = (c.polos || []).map(Number).filter(x => x > 0);
    // [U16:W16] reparto por fases
    const fase = [0, 0, 0];
    if (J && P) {
      if (N === 1) fase[0] = J;
      else polos.forEach(p => { fase[fasePolo(p, N)] = J / P; });
    }
    const dem = fase.map(x => x * R / S);                                      // [X16:Z16]
    const I = fase.map((u, i) => corriente(u, dem[i], O, P));                // [AA16:AC16]
    const AD = c.mult !== undefined && c.mult !== '' && c.mult !== null ? n(c.mult) : (det.mult || 1.25);
    const AE = !J ? null : N === 3 ? Math.max(I[0], I[1], I[2]) * AD : (fase[0] > 0 ? I[0] * AD : fase[1] > 0 ? I[1] * AD : fase[2] > 0 ? I[2] * AD : null);
    const AF = n(c.prot) || (AE !== null ? proteccion(AE) : null);          // protección
    const mat = c.material || 'CU', ais = c.aislamiento || 'THHN';
    const AI = n(c.paralelos) || (AE !== null ? paralelosAuto(AE) : 1);
    const AJ = c.aumento !== undefined && c.aumento !== '' ? n(c.aumento) : 1;
    const AL = AF ? calibre(AF * AJ / AI, mat) : '';
    const pre = AI === 1 ? (P === 3 ? '3#' : (O >= 208 && O <= 240 ? '2#' : ' ')) : AI + 'x' + (P === 3 ? '3#' : (O >= 208 && O <= 240 ? '2#' : ' '));
    const pre2 = AI === 1 ? ' ' : AI + 'x';
    const AN = Q === 0 ? ' ' : (O >= 208 && P === 3 && Q === 4) ? AL : (P === 3 && Q === 3) ? ' ' : (O >= 208 && P === 1 && Q === 3) ? ' ' : AL;
    const AP = AF ? tierra(AF, mat) : '';
    const AR = AL ? conduit(AL, mat, ais) : '';
    const AS = AL ? fac(AL, mat, T) : null, AT = facAjustado(AS, O, P, false);
    const AU = !J ? null : (O === ctx.V ? ctx.vBus : ctx.vBusLN);             // voltaje real de partida
    const Iv = fase[0] > 0 ? I[0] : fase[1] > 0 ? I[1] : I[2];
    const AV = J && AT !== null ? n(c.longitud) * 3.28 * Iv * AT / (10000 * AI) : null;
    const AW = AV !== null ? AU - AV : null, AX = AW !== null ? O - AW : null, AY = AX !== null && O ? AX * 100 / O : null;
    const BE = P === 3 ? 3 : (O >= 208 && O <= 240 ? 2 : P || 1);
    const nec = J && !c.unidad && !c.breakerId ? necUnidad(c.descripcion, det, O, P, AF || 0, ctx.fases, ctx.V) : null;   // vacío = automático según NEC
    const bk = J ? breaker(ctx.marca, c.breakerId, BE, AF, { V: ctx.V, familias: ctx.famRamales, iccKA: ctx.iccKA, unidad: c.unidad || (nec ? nec.unidad : 'STD') }) : null;
    const err = [];
    if (J && !det.id) err.push('Sin detalle de carga');
    if (J && !polos.length) err.push('Sin posición');
    if (J && O && O !== ctx.V && O !== Math.round(vLN(ctx.V)) && !(ctx.V === 240 && O === 120)) err.push('El detalle es de ' + O + ' V y el tablero de ' + ctx.V + ' V: elija un detalle de carga de este voltaje');
    if (J && P && polos.length > P) err.push('El detalle es de ' + P + ' fase(s) y el circuito ocupa ' + polos.length + ' polos: la carga se suma en cada fase. Use un detalle de ' + polos.length + ' fases o deje un solo polo.');
    if (AY !== null && AY > ctx.cvMaxTotal) err.push('Caída total ' + r2(AY) + ' % > ' + ctx.cvMaxTotal + ' %');
    if (bk && bk.sccrBajo) err.push('SCCR ' + bk.sccr + ' kA < Icc ' + r2(ctx.iccKA) + ' kA (no hay breaker de la familia con mayor capacidad)');
    if (J && !bk) err.push('Sin breaker de catálogo para ' + BE + 'P ' + (AF || '') + ' A');
    if (J && AF && !AL) err.push('Sin calibre para ' + AF + ' A');
    return {
      c, det, O, P, Q, R, S, T, L: det.tipo, J, polos, fase, dem, I, AD, AE, AF, AI, AJ, mat, ais, AL, AN, AP, AR, AS, AT, AU, AV, AW, AX, AY,
      fasesTxt: pre + AL, neutroTxt: (AN === ' ' ? '' : pre2.trim() + AN), tierraTxt: AP ? pre2.trim() + AP : '', tuboTxt: AR ? pre2.trim() + AR : '',
      descripcion: c.descripcion || det.descripcion || '', breaker: bk, polosBreaker: BE, err, nec,
    };
  }

  /* ---------- un tablero completo ---------- */
  /** Calcula un tablero. ctx: { vInicio, iccInicioA } que vienen del tablero que lo alimenta (cascada). */
  function tablero(tab, P, ctxIn, altOrigen) {
    const fases = n(tab.fases) || 3, V = n(tab.voltaje) || 208, al = tab.alim || {};
    const marcaTab = tab.catalogoId ? ((K.tablerosCat.find(t => String(t.id) === String(tab.catalogoId)) || {}).fabricante) : (tab.marca || K.marcaDefecto || 'Eaton');

    // --- 1) cargas y fases, sin voltaje (para conocer el alimentador primero)
    // orden físico del tablero: primero el lado impar (1, 3, 5…) y luego el par (2, 4, 6…)
    const llave = c => { const p = Math.min(...((c.polos || []).length ? c.polos.map(Number) : [9999])); return (p % 2 ? 0 : 10000) + p; };
    const circs = (tab.circuitos || []).slice().sort((a, b) => llave(a) - llave(b));
    const pre = circs.map(c => circuito(c, tab, { fases, V, vBus: V, vBusLN: vLN(V), marca: marcaTab, cvMaxTotal: 999 }));
    const suma = pre.filter(x => !x.c.respaldoDe);       // el circuito de bypass ocupa espacios pero no suma carga
    const J116 = suma.reduce((a, x) => a + x.J, 0);
    const U116 = [0, 1, 2].map(i => suma.reduce((a, x) => a + x.fase[i], 0));
    const J117 = J116 / 3, U117 = U116.map(x => (J117 ? x / J117 : 0));
    const mx = fases === 3 ? Math.max(...U116) : fases === 2 ? Math.max(U116[0], U116[1]) : 0;
    const mn = fases === 3 ? Math.min(...U116) : fases === 2 ? Math.min(U116[0], U116[1]) : 0;
    const desbalance = mx ? r2((mx - mn) / mx * 100) : 0;                    // [U118]

    // --- 2) factores de demanda por tipo de carga [filas 123 a 134]
    const reserva = n(tab.reserva);
    const demandaDeriv = P && P.cargaDerivados === 'demandada';
    const derivados = demandaDeriv ? suma.filter(x => x.c.tableroHijoId) : [];
    const tiposCat = K.tiposCarga.concat(derivados.length ? [{ id: 'deriv', nombre: 'TABLEROS DERIVADOS (demanda ya aplicada)', metodo: 'fijo', fd: 1, fdiv: 1 }] : []);
    const tipos = tiposCat.map(t => {
      const cs = t.id === 'deriv' ? derivados : suma.filter(x => String(x.L) === String(t.id) && !derivados.includes(x));
      const conectados = cs.reduce((a, x) => a + x.J, 0), res = conectados * reserva, total = conectados + res;
      const ov = (tab.fd || {})[t.id];
      let fd, demandados;
      const fdiv = n((tab.fdiv || {})[t.id]) || n(t.fdiv) || 1;
      if (ov !== undefined && ov !== '' && ov !== null) { fd = n(ov); demandados = total * fd / fdiv; }
      else if (t.metodo === 'tomas') { demandados = total > 10 ? (total - 10) * 0.5 + 10 : total; fd = total > 0 ? demandados / total : 1; }
      else if (t.metodo === 'cocina') { const nn = cs.length, f = (K.demanda22056 || []).filter(x => nn >= x.n).pop(); fd = f ? f.fd : 1; demandados = total * fd / fdiv; }
      else if (t.metodo === '220.53') { fd = cs.length >= 4 ? 0.75 : 1; demandados = total * fd / fdiv; }
      else { fd = n(t.fd) || 1; demandados = total * fd / fdiv; }
      return { tipo: t, conectados, reserva: res, total, fd, fdiv, demandados, n: cs.length };
    });
    const W128 = tipos.reduce((a, t) => a + t.conectados, 0), W129 = tipos.reduce((a, t) => a + t.reserva, 0);
    const W130 = W128 + W129, J134 = tipos.reduce((a, t) => a + t.demandados, 0);

    // --- 3) alimentador / acometida [fila 139]
    const R139 = W130 ? J134 / W130 : 1, S139 = n(tab.fdivTablero) || 1, L139 = W130 * R139 / S139, T139 = n(al.fp) || 0.9;
    const X139 = U117.map(u => L139 * u / 3);
    const AA139 = X139.map(x => (x > 0 ? corrienteAlim(x, V, fases) : 0));
    const AD139 = al.rated100 ? 1 : (al.mult !== undefined && al.mult !== '' ? n(al.mult) : 1.25);
    const AE139 = W130 > 0 ? (fases === 3 ? Math.max(...AA139) * AD139 : (U116[0] > 0 ? AA139[0] : U116[1] > 0 ? AA139[1] : AA139[2]) * AD139) : null;
    const AF139 = n(al.prot) || (AE139 ? proteccion(AE139) : null);
    const mat = al.material || 'CU', ais = al.aislamiento || 'XHHW-2', hilos = n(tab.hilos) || 4;
    const AJ139 = al.aumento !== undefined && al.aumento !== '' ? n(al.aumento) : 1;
    let AI139 = n(al.paralelos) || (AE139 ? paralelosAuto(AE139) : 1);
    let AL139 = AF139 ? calibre(AF139 * AJ139 / AI139, mat) : '';
    // validación por temperatura y agrupamiento (310.15) y aumento automático de calibre o de paralelos si no cumple
    const agrup = al.agrupamiento || agrupAuto(fases, hilos);
    const calBase = AL139, parBase = AI139;
    let val = AL139 ? ampacidad(AL139, mat, AI139, ais, al.tempAmb, al.tempBorne, agrup, AE139 || 0, AF139) : null, ajustado = false;
    const autoAmp = !(P && P.autoAmpacidad === false) && !al.calibreFijo;
    for (let it = 0; autoAmp && val && !val.ok && val.ft > 0 && it < 14; it++) {
      const i = CALIBRES.indexOf(calStr(AL139));
      const sig = CALIBRES.slice(i + 1).find(c => (K.conduit || []).some(x => x.cal === c) && (K.ampacidad31016 || []).some(x => x.cal === c));
      if (sig && CALIBRES.indexOf(sig) <= CALIBRES.indexOf('500')) AL139 = sig;
      else { AI139++; AL139 = calibre(AF139 * AJ139 / AI139, mat); }
      val = ampacidad(AL139, mat, AI139, ais, al.tempAmb, al.tempBorne, agrup, AE139 || 0, AF139); ajustado = true;
    }
    const pre139 = AI139 === 1 ? (fases === 3 ? '3#' : (V >= 208 && V <= 240 ? '2#' : ' ')) : AI139 + 'x' + (fases === 3 ? '3#' : (V >= 208 && V <= 240 ? '2#' : ' '));
    const preN = AI139 === 1 ? '' : AI139 + 'x';
    const AN139 = hilos === 0 ? '' : (V >= 208 && fases === 3 && hilos === 4) ? AL139 : (fases === 3 && hilos === 3) ? '' : (V >= 208 && fases === 1 && hilos === 3) ? '' : AL139;
    const AP139 = AF139 ? tierra(AF139, mat) : '', AP140 = AL139 ? electrodo(AL139, mat) : '';
    const AR139 = AL139 ? conduit(AL139, mat, ais) : '', tuberia = al.tuberia || 'EMT';
    const AS139 = AL139 ? fac(AL139, mat, T139) : null, AT139 = facAjustado(AS139, V, fases, true);
    const Iv = U116[0] > 0 ? AA139[0] : U116[1] > 0 ? AA139[1] : AA139[2];
    const M139 = n(tab.longitud);

    // --- transformador (opcional) aguas arriba del alimentador: regulación y cortocircuito
    const tr = tab.trafo && tab.trafo.activo && n(tab.trafo.kva) && n(tab.trafo.z) ? tab.trafo : null;
    const Vpadre = ctxIn && ctxIn.Vpadre;
    let vInicio = V, trafo = null;
    if (tr) {
      const kva = n(tr.kva), z = n(tr.z), xr = n(tr.xr) || 4, pR = z / Math.sqrt(1 + xr * xr), pX = pR * xr;
      const sin = Math.sqrt(Math.max(0, 1 - T139 * T139)), carga = kva ? L139 / kva : 0;
      const reg = carga * (pR * T139 + pX * sin);                                  // caída de tensión en el transformador (%)
      vInicio = V * (1 - reg / 100);
      trafo = { kva, z, xr, carga, reg, Vp: n(tr.primario) || Vpadre || null, Vs: V, fases: n(tr.fases) || (fases === 3 ? 3 : 1), nombre: tr.nombre || '' };
    } else if (ctxIn && ctxIn.vInicio && (!Vpadre || Vpadre === V)) vInicio = ctxIn.vInicio;   // cascada [AU139]
    const upsD = tab.ups && tab.ups.activo && n(tab.ups.kva) ? tab.ups : null;
    let ups = null;
    if (upsD) { vInicio = V; ups = { kva: n(upsD.kva), factor: n(upsD.factorIcc) || 2, carga: L139 / n(upsD.kva), nombre: upsD.nombre || 'UPS' }; ups.icc = ups.factor * ups.kva * 1000 / ((fases === 3 ? 1.732 : 1) * V); }

    const AV139 = AT139 !== null && W130 ? M139 * 3.28 * Iv * AT139 / (10000 * AI139) : 0;
    const AW139 = vInicio - AV139, AX139 = V - AW139, AY139 = AX139 * 100 / V;   // bornes, caída total acumulada
    const AU140 = vLN(V), AV140 = AV139 / (V === 240 ? 2 : 1.732), AW140 = AU140 - (V - vInicio) / (V === 240 ? 2 : 1.732) - AV140;

    // verificación de ampacidad como en el Excel [AA10:AJ10] (se conserva para la memoria)
    const tempF = ((K.tempFactor || {})[al.tempAmb || '26-30'] || {})[String(al.tempBorne || 90)] || 1;
    const agrF = (K.agrupamiento || {})[agrup] || 1;
    const ampReq = AE139 ? AE139 / (tempF * agrF) : 0;
    const fAmp = (K.ampacidad31016 || []).find(x => x.cal === calStr(AL139));
    const ampCond = fAmp ? n(fAmp[mat + '|' + (al.tempBorne || 90)]) * AI139 : null;

    // --- 4) cortocircuito (punto a punto) [AQ8:AQ10]
    let iccA = null, iccFuente = '';
    const lmax = n(P && P.iccLongMax) || 0;
    const cIcc = () => {
      const f = (K.constC || []).find(x => x.cal === calStr(AL139)); if (!f) return null;
      const aisC = /XHHW/.test(ais) ? 'XHHW-2' : /RHW/.test(ais) ? 'RHW' : /barra/i.test(ais) ? 'DUCTOBARRA' : 'THHN';
      return n(f[aisC + '|' + (aisC === 'DUCTOBARRA' ? 'DUCTOBARRA' : tuberia)]) || null;
    };
    const p2p = I0 => {
      const C = cIcc(); if (!C || !I0) return I0;
      const Lm = lmax ? Math.min(M139, lmax) : M139, k = fases === 3 ? 1.732 : 2;
      return I0 / (1 + (k * (Lm / 0.3048) * I0) / (C * AI139 * V));
    };
    const p2pCal = (I0, cal, aisX, tub, nPar, L) => {
      const f = (K.constC || []).find(x => x.cal === calStr(cal)); if (!f || !I0) return I0;
      const aC = /XHHW/.test(aisX) ? 'XHHW-2' : /RHW/.test(aisX) ? 'RHW' : /barra/i.test(aisX) ? 'DUCTOBARRA' : 'THHN';
      const C = n(f[aC + '|' + (aC === 'DUCTOBARRA' ? 'DUCTOBARRA' : tub)]); if (!C) return I0;
      const Lm = lmax ? Math.min(L, lmax) : L, k = fases === 3 ? 1.732 : 2;
      return I0 / (1 + (k * (Lm / 0.3048) * I0) / (C * nPar * V));
    };
    const Ip = (ctxIn && ctxIn.iccInicioA) || (!ctxIn && n(P && P.iccRed) ? n(P.iccRed) * 1000 : null);
    if (n(tab.iccManual)) { iccA = n(tab.iccManual) * 1000; iccFuente = 'manual'; }
    else if (trafo) {
      const Vp = trafo.Vp || V, k3 = trafo.fases === 3;
      let Is;
      if (Ip) { const f = (k3 ? Ip * Vp * 1.732 : Ip * Vp) * trafo.z / (100000 * trafo.kva); Is = (Vp / V) * Ip / (1 + f); }
      else Is = trafo.kva * 1000 / ((k3 ? 1.732 : 1) * V * trafo.z / 100);          // red de capacidad infinita
      trafo.iccPrim = Ip; trafo.iccSec = Is; iccA = p2p(Is); iccFuente = 'transformador';
    } else if (tab.transformadorId) {
      const t0 = K.transformadores.find(t => String(t.id) === String(tab.transformadorId));
      const k = t0 ? (V >= 440 ? t0.kacc480 : V >= 230 ? t0.kacc240 : t0.kacc208) : null;
      if (k) { iccA = p2p(n(k)); iccFuente = 'transformador (KACC)'; }
    } else if (ctxIn && ctxIn.iccInicioA && (!Vpadre || Vpadre === V)) { iccA = p2p(ctxIn.iccInicioA); iccFuente = 'cascada'; }
    if (ups) { iccA = p2p(iccA ? Math.min(ups.icc, iccA) : ups.icc); iccFuente = 'UPS (limitado por el inversor)'; }

    // --- segunda acometida (generador o bypass desde otro tablero) con ATS / MTS / interruptor
    const altD = tab.alterna && tab.alterna.activo ? tab.alterna : null;
    let alterna = null;
    if (altD && AF139) {
      const am = altD.material || 'CU', aa = altD.aislamiento || 'XHHW-2', at = altD.tuberia || 'EMT', AL = n(altD.longitud);
      let an = n(altD.paralelos) || AI139, ac = calibre(AF139 * AJ139 / an, am);
      let av = ac ? ampacidad(ac, am, an, aa, al.tempAmb, al.tempBorne, agrup, AE139 || 0, AF139) : null;
      for (let it = 0; autoAmp && av && !av.ok && av.ft > 0 && it < 14; it++) {
        const i = CALIBRES.indexOf(calStr(ac));
        const sig = CALIBRES.slice(i + 1).find(c => (K.conduit || []).some(x => x.cal === c) && (K.ampacidad31016 || []).some(x => x.cal === c));
        if (sig && CALIBRES.indexOf(sig) <= CALIBRES.indexOf('500')) ac = sig; else { an++; ac = calibre(AF139 * AJ139 / an, am); }
        av = ampacidad(ac, am, an, aa, al.tempAmb, al.tempBorne, agrup, AE139 || 0, AF139);
      }
      const af = facAjustado(ac ? fac(ac, am, T139) : null, V, fases, true);
      const avd = af !== null && W130 ? AL * 3.28 * Iv * af / (10000 * an) : 0;
      let v0 = V, iSrc = null, carga = null, origen = '';
      if (altD.tipo === 'generador') {
        const kva = n(altD.kva), xd = n(altD.xd) || 12;
        if (kva) { iSrc = kva * 1000 / ((fases === 3 ? 1.732 : 1) * V * xd / 100); carga = L139 / kva; }
        origen = (altD.nombre || 'Generador') + (kva ? ' ' + kva + ' kVA' : '');
      } else if (altOrigen) {
        if (altOrigen.V === V) { v0 = altOrigen.alim.AW139; iSrc = altOrigen.iccA; }
        origen = altOrigen.nombre;
      }
      const pre = an === 1 ? (fases === 3 ? '3#' : '2#') : an + 'x' + (fases === 3 ? '3#' : '2#'), pN = an === 1 ? '' : an + 'x';
      const iAlt = iSrc ? p2pCal(iSrc, ac, aa, at, an, AL) : null;
      alterna = { equipo: altD.equipo || 'ATS', tipo: altD.tipo || 'generador', origen, nombre: altD.nombre || '', cal: ac, par: an, mat: am, ais: aa, tuberia: at, longitud: AL, val: av,
        fasesTxt: pre + ac, neutroTxt: AN139 ? pN + ac : '', tierraTxt: AP139 ? pN + tierra(AF139, am) : '', tubo: ac ? pN + conduit(ac, am, aa) : '',
        caida: avd, vBornes: v0 - avd, cv: (V - (v0 - avd)) * 100 / V, iccA: iAlt, carga, kva: n(altD.kva), xd: n(altD.xd) || 12 };
      if (iAlt && (!iccA || iAlt > iccA)) { iccA = iAlt; iccFuente = (iccFuente ? iccFuente + ' · ' : '') + 'máx. por ' + (alterna.tipo === 'generador' ? 'generador' : 'bypass'); }
    }
    const iccKA = iccA ? iccA / 1000 : null;

    // --- 5) tablero de catálogo y breakers según las reglas del fabricante
    const espaciosUsados = Math.max(0, ...circs.map(c => Math.max(0, ...(c.polos || []).map(Number))));
    const cat = tableroCat(tab, AF139, espaciosUsados, V, P && P.reservaEspacios);
    const famT = cat ? famTablero(cat) : null;
    const marca = cat ? cat.fabricante : marcaTab;
    const cvMaxTotal = n(P && P.cvMaxTotal) || 5;
    const polosBk = fases === 3 ? 3 : (V >= 208 && V <= 240 ? 2 : fases);
    const zapatas = tab.principal === 'zapatas';
    const bkMain = AF139 && !zapatas ? breaker(marca, al.breakerId, polosBk, AF139, { V, familias: famT ? famT.principales : '', iccKA, unidad: al.unidad || 'STD' }) : null;
    if (bkMain && al.rated100) { bkMain.rated100 = true; if (/^PDG/.test(bkMain.ref)) bkMain.ref = bkMain.ref.replace(/^PDG/, 'PDF'); }
    const rows = circs.map(c => {
      const x = circuito(c, tab, { fases, V, vBus: AW139, vBusLN: AW140, marca, cvMaxTotal, iccKA, famRamales: famT ? famT.ramales : '' });
      if (x.breaker && x.breaker.sccrBajo && bkMain && !bkMain.sccrBajo) {
        const sr = serie(x.breaker, bkMain, AF139, V, iccKA);
        if (sr) { x.breaker.sccrBajo = false; x.breaker.serie = sr; x.err = x.err.filter(e => !/^SCCR/.test(e)); }
      }
      return x;
    });
    const enSerie = rows.filter(x => x.breaker && x.breaker.serie);
    const servicio = !ctxIn || !!trafo;
    const kaSpdMin = kaSpd(!ctxIn, n(cat && cat.barraFase) || n(AF139));
    const spd = tab.sinSupresor ? null : supresor(marca, tab.supresorId, tab.sistema, fases, kaSpdMin);

    // --- revisión NEC 2020 del tablero
    const nec = [];
    if (!ctxIn) nec.push({ art: '110.24', nivel: 'info', txt: 'Rotular en el equipo de acometida la corriente de falla disponible (' + (iccKA ? r2(iccKA) + ' kA' : 'calcular') + ') y la fecha del cálculo.' });
    if (iccKA) nec.push({ art: '408.6', nivel: 'info', txt: 'El SCCR marcado del tablero debe ser ≥ ' + r2(iccKA) + ' kA.' });
    if (!ctxIn && (K.ocupacion === 'vivienda') && !spd) nec.push({ art: '230.67', nivel: 'error', txt: 'Acometida de vivienda: se requiere SPD tipo 1 o 2.' });
    if (spd && n(spd.kaLL) < kaSpdMin) nec.push({ art: '242', nivel: 'aviso', txt: 'SPD de ' + spd.kaLL + ' kA; se recomiendan ≥ ' + kaSpdMin + ' kA por fase para este tablero' + (!ctxIn ? ' (acometida, tipo 1 o 2)' : '') + '.' });
    if (spd) nec.push({ art: '242', nivel: 'info', txt: 'SPD ' + (!ctxIn ? 'tipo 1 o 2 en la acometida' : 'tipo 2 en el lado de carga') + '; MCOV ≥ ' + (V >= 440 ? '320 V L-N (550 V L-L)' : V === 240 ? '150 V L-N (320 V L-L)' : '150 V L-N') + '; In = 20 kA; SCCR del SPD ≥ Icc.' });
    if (AF139 >= 1200) nec.push({ art: '240.87', nivel: 'aviso', txt: 'Interruptor de ' + AF139 + ' A: requiere un método de reducción de energía de arco (ARMS, ZSI, relé diferencial o mitigación activa), ajustado por debajo de la corriente de arco, probado y documentado. Usar disparo con ARMS (p. ej. PXR25 LSIG+ARMS).' });
    if (AF139 >= 1200 && !ctxIn) nec.push({ art: '110.16(B)', nivel: 'info', txt: 'Acometida ≥ 1200 A: etiqueta de arco eléctrico con voltaje, corriente de falla, tiempo de despeje y fecha.' });
    if (alterna && /ATS|MTS/.test(alterna.equipo)) nec.push({ art: '700.5(E) / 702.5', nivel: n(altD.sccr) && iccKA && n(altD.sccr) < iccKA ? 'error' : 'aviso', txt: 'Rotular en campo el SCCR del ' + alterna.equipo + (n(altD.sccr) ? ' (' + altD.sccr + ' kA)' : '') + ' según el dispositivo de protección aguas arriba; debe ser ≥ ' + (iccKA ? r2(iccKA) + ' kA' : 'la corriente de falla') + '.' });
    rows.filter(x => x.nec && !x.c.unidad).forEach(x => nec.push({ art: x.nec.motivo.split(':')[0], nivel: 'info', txt: (x.descripcion || '') + ' [' + x.polos.join(',') + ']: interruptor ' + x.nec.unidad + ' (' + x.nec.motivo + ').' }));
    nec.filter(x => x.nivel === 'error').forEach(x => avisos.push('NEC ' + x.art + ': ' + x.txt));

    const avisos = [];
    if (AY139 > n(P && P.cvMaxAlim || 3)) avisos.push('Caída acumulada en bornes ' + r2(AY139) + ' % (máx. ' + n(P && P.cvMaxAlim || 3) + ' %)');
    if (val && !val.ok) avisos.push('El alimentador no cumple ampacidad corregida: ' + r2(val.cap) + ' A disponibles' + (val.okI ? '' : ' < ' + r2(AE139) + ' A requeridos') + (val.okP ? '' : ', protección ' + AF139 + ' A mayor que la permitida') + (val.ft ? '' : ' (temperatura ambiente fuera del rango del aislamiento)'));
    if (ajustado && val && val.ok) avisos.push('Calibre del alimentador aumentado por temperatura/agrupamiento: ' + (parBase > 1 ? parBase + 'x' : '') + calBase + ' → ' + (AI139 > 1 ? AI139 + 'x' : '') + AL139 + ' AWG/kcmil');
    if (cat && espaciosUsados > n(cat.espacios)) avisos.push('Circuitos usan ' + espaciosUsados + ' espacios; el tablero tiene ' + cat.espacios);
    if (!cat) avisos.push('No hay tablero de catálogo (' + marcaTab + ') para ' + V + ' V, ' + (AF139 || '?') + ' A y ' + espaciosUsados + ' espacios');
    if (bkMain && al.unidad && al.unidad !== 'STD' && bkMain.unidad !== al.unidad) avisos.push('No hay interruptor principal ' + al.unidad + ' para ' + AF139 + ' A en el catálogo de ' + marca + '; se usó ' + (bkMain.unidad || 'STD'));
    if (bkMain && al.rated100 && !/^PDF/.test(bkMain.ref)) avisos.push('Interruptor principal 100 % rated: verifique el modelo con el fabricante (' + (bkMain.ref || bkMain.modelo) + ')');
    if (bkMain && bkMain.sccrBajo) avisos.push('SCCR del interruptor principal ' + bkMain.sccr + ' kA < Icc ' + r2(iccKA) + ' kA');
    if (trafo && trafo.carga > 1) avisos.push('Transformador cargado al ' + r2(trafo.carga * 100) + ' %');
    if (enSerie.length) {
      const k = Math.min(...enSerie.map(x => n(x.breaker.serie.kA)));
      avisos.push('Clasificación en serie (240.86): ' + enSerie.length + ' ramal(es) protegidos en serie con el principal ' + (bkMain.ref || bkMain.modelo) + ' hasta ' + k + ' kA. Rotular el tablero según 110.22(C) ("Precaución – sistema en serie… ' + k + ' kA").');
      if (rows.some(x => x.J && [4, 5, 6, 7, 8, 9, 10].includes(Number(x.L)))) avisos.push('240.86(C): hay motores en el tablero; la serie no aplica si la contribución de motores supera el 1 % del rating del ramal. Verifique.');
    }
    if (ups && ups.carga > 1) avisos.push('UPS cargada al ' + r2(ups.carga * 100) + ' %');
    if (alterna) {
      if (alterna.cv > n(P && P.cvMaxAlim || 3)) avisos.push('Segunda acometida (' + alterna.equipo + '): caída ' + r2(alterna.cv) + ' % (máx. ' + n(P && P.cvMaxAlim || 3) + ' %)');
      if (alterna.val && !alterna.val.ok) avisos.push('Segunda acometida: el conductor no cumple ampacidad corregida');
      if (alterna.carga > 1) avisos.push('Generador cargado al ' + r2(alterna.carga * 100) + ' %');
      if (alterna.tipo !== 'generador' && !altOrigen) avisos.push('Segunda acometida: elija el tablero de origen del bypass');
    }
    if (Vpadre && Vpadre !== V && !trafo) avisos.push('Voltaje distinto al del tablero que lo alimenta: agregue el transformador (diagrama unifilar o memoria)');
    const ocupados = {};
    rows.forEach(x => x.polos.forEach(p => { (ocupados[p] = ocupados[p] || []).push(x); }));
    Object.keys(ocupados).forEach(p => { if (ocupados[p].length > 1) avisos.push('Posición ' + p + ' repetida (' + ocupados[p].map(x => x.descripcion || '?').join(' / ') + ')'); });
    if (desbalance > n(P && P.desbalanceMax || 10) && J116) avisos.push('Desbalance ' + desbalance + ' % (máx. ' + n(P && P.desbalanceMax || 10) + ' %)');

    return {
      tab, nombre: nombreTablero(tab), fases, V, hilos, rows, J116, U116, U117, desbalance, tipos, W128, W129, W130, J134,
      alim: { R139, S139, L139, T139, X139, AA139, AD139, AE139, AF139, mat, ais, AI139, AJ139, AL139, AN139, AP139, AP140, AR139, tuberia, AS139, AT139,
        M139, vInicio, AV139, AW139, AX139, AY139, AU140, AV140, AW140, tempF, agrF, ampReq, ampCond, agrup, val, ajustado, calBase, parBase,
        fasesTxt: pre139 + AL139, neutroTxt: AN139 ? preN + AN139 : '', tierraTxt: AP139 ? preN + AP139 : '', tuboTxt: AR139 ? preN + AR139 : '', preN, preF: pre139.trim() },
      iccA, iccKA, iccFuente, trafo, ups, alterna, zapatas, enSerie, nec, servicio, kaSpdMin, cat, famT, marca, bkMain, spd, espaciosUsados, avisos,
    };
  }

  const nombreTablero = t => ((t.prefijo === undefined ? 'TABLERO' : t.prefijo) + ' ' + (t.nombre || '')).trim();

  /* ---------- proyecto completo (cascada) ---------- */
  function proyecto(P, catalog) {
    K = Object.assign({}, catalog, { marcaDefecto: P.marcaDefecto || 'Eaton', ocupacion: P.ocupacion || 'comercial' });
    const tabs = P.tableros || [], byId = {}; tabs.forEach(t => { byId[t.id] = t; });
    const hijos = {}; tabs.forEach(t => { if (t.padreId && byId[t.padreId]) (hijos[t.padreId] = hijos[t.padreId] || []).push(t); });
    // ciclos: un tablero no puede alimentarse de sí mismo ni de sus derivados
    const ciclo = new Set();
    tabs.forEach(t => { const seen = new Set(); let x = t; while (x && x.padreId) { if (seen.has(x.id)) { ciclo.add(t.id); break; } seen.add(x.id); x = byId[x.padreId]; } });
    const raiz = t => !t.padreId || !byId[t.padreId] || ciclo.has(t.id);

    // kVA de los tableros derivados hacia el circuito que los alimenta (de abajo hacia arriba)
    const total = {}, visit = new Set();
    function cargar(t) {
      if (visit.has(t.id)) return total[t.id] || 0; visit.add(t.id);
      (hijos[t.id] || []).forEach(cargar);
      (t.circuitos || []).forEach(c => {
        const hj = c.tableroHijoId && byId[c.tableroHijoId];
        if (hj && !ciclo.has(hj.id)) { c.kva = r4(total[hj.id] || 0); c.longitud = hj.longitud; c.descripcion = nombreTablero(hj); }
        const rb = c.respaldoDe && byId[c.respaldoDe];
        if (rb) { if (!visit.has(rb.id)) cargar(rb); c.kva = r4(total[rb.id] || 0); c.longitud = (rb.alterna || {}).longitud; c.descripcion = 'BYPASS ' + nombreTablero(rb); }
      });
      const rt = tablero(t, P, null);
      total[t.id] = P.cargaDerivados === 'demandada' ? rt.alim.L139 : rt.W130;
      return total[t.id];
    }
    tabs.forEach(cargar);

    // voltaje y cortocircuito de arriba hacia abajo (segunda pasada si hay bypass desde otro tablero)
    let res = {}, orden = [];
    const correr = prev => {
      res = {}; orden = [];
      const bajar = (t, ctx, nivel) => {
        const alt = t.alterna && t.alterna.activo && t.alterna.tipo !== 'generador' && prev ? prev[t.alterna.origenId] : null;
        const r = tablero(t, P, ctx, alt); r.nivel = nivel; r.padre = ctx && ctx.padre; res[t.id] = r; orden.push(r);
        (hijos[t.id] || []).filter(h => !ciclo.has(h.id)).forEach(h => {
          bajar(h, { vInicio: r.alim.AW139, iccInicioA: r.iccA, Vpadre: r.V, padre: r }, nivel + 1);
        });
      };
      tabs.filter(raiz).forEach(t => bajar(t, null, 0));
    };
    correr(null);
    if (tabs.some(t => t.alterna && t.alterna.activo && t.alterna.tipo !== 'generador' && t.alterna.origenId)) correr(res);
    orden.forEach(r => {
      r.alimentadoDesde = r.padre ? r.padre.nombre : (r.tab.conectadoA || '');
      r.circuitoPadre = r.padre ? (r.padre.rows.find(x => x.c.tableroHijoId === r.tab.id) || null) : null;
      if (r.padre && !r.circuitoPadre) r.avisos.push('El tablero ' + r.padre.nombre + ' no tiene un circuito asignado a este tablero');
      if (ciclo.has(r.tab.id)) r.avisos.push('Alimentación circular: revise "Alimentado desde"');
    });
    return { res, orden, byId, hijos };
  }
  const r4 = v => Math.round(v * 10000) / 10000;

  /* ---------- autobalanceo (hoja BALANCEO) ---------- */
  /** Busca posiciones que reduzcan el desbalance: intercambia circuitos con el mismo número de polos y mueve circuitos
      (de 1, 2 o 3 polos) a espacios libres del mismo lado. Respeta los circuitos fijados (c.fijo).
      opts: { espacios, maxMov, objetivo (%) }. Devuelve los movimientos y los cambios netos de posición (para Revit). */
  function balanceo(tab, catalog, opts) {
    K = catalog; opts = typeof opts === 'number' ? { maxMov: opts } : (opts || {});
    const fases = n(tab.fases) || 3, nf = fases === 3 ? 3 : 2;
    const base = (tab.circuitos || []).map(c => ({ c, polos: (c.polos || []).map(Number).filter(x => x > 0), carga: circuito(c, tab, { fases, V: n(tab.voltaje), vBus: 0, vBusLN: 0, marca: '', cvMaxTotal: 999 }) }));
    const espacios = Math.max(n(opts.espacios), ...base.map(x => Math.max(0, ...x.polos)));
    if (fases === 1 || !espacios) return { movs: [], cambios: [], antes: 0, despues: 0, resultado: [] };
    const carga = arr => { const f = [0, 0, 0]; arr.forEach(x => x.polos.forEach(p => { f[fasePolo(p, fases)] += x.carga.J / (x.carga.P || x.polos.length || 1); })); return f; };
    const des = f => { const v = f.slice(0, nf); const mx = Math.max(...v); return mx ? (mx - Math.min(...v)) / mx * 100 : 0; };
    const lado = p => (p % 2 ? 'impar' : 'par');
    // se minimiza la dispersión entre fases (suma de desviaciones al cuadrado); el % de desbalance se informa
    const disp = f => { const v = f.slice(0, nf), m = v.reduce((a, b) => a + b, 0) / nf; return v.reduce((a, b) => a + (b - m) * (b - m), 0); };
    let cs = base.map(x => Object.assign({}, x));
    const antes = des(carga(cs)), movs = [], objetivo = n(opts.objetivo);
    for (let it = 0; it < (opts.maxMov || 30); it++) {
      const actual = disp(carga(cs)); if (objetivo && des(carga(cs)) <= objetivo) break;
      let best = null;
      const usados = new Set(); cs.forEach(x => x.polos.forEach(p => usados.add(p)));
      const mejor = (d, prueba, txt, dist) => { if (d < actual - 1e-6 && (!best || d < best.d - 1e-9 || (Math.abs(d - best.d) < 1e-9 && dist < best.dist))) best = { d, prueba, txt, dist }; };
      for (let i = 0; i < cs.length; i++) {
        const a = cs[i]; if (!a.carga.J || a.c.fijo || !a.polos.length) continue;
        for (let j = i + 1; j < cs.length; j++) {
          const b = cs[j]; if (b.c.fijo || b.polos.length !== a.polos.length) continue;
          const prueba = cs.map((x, k) => (k === i ? Object.assign({}, x, { polos: b.polos }) : k === j ? Object.assign({}, x, { polos: a.polos }) : x));
          mejor(disp(carga(prueba)), prueba, 'Intercambiar [' + a.polos.join(',') + '] ' + (a.carga.descripcion || '') + ' ↔ [' + b.polos.join(',') + '] ' + (b.carga.descripcion || '(espacio)'), Math.abs(a.polos[0] - b.polos[0]));
        }
        for (let p = 1; p <= espacios; p++) {
          const set = a.polos.map((_, k) => p + 2 * k);
          if (set[set.length - 1] > espacios || set.some(x => usados.has(x) && !a.polos.includes(x))) continue;
          if (set.join() === a.polos.join()) continue;
          const prueba = cs.map((x, k) => (k === i ? Object.assign({}, x, { polos: set }) : x));
          mejor(disp(carga(prueba)), prueba, 'Mover [' + a.polos.join(',') + '] ' + (a.carga.descripcion || '') + ' → [' + set.join(',') + '] (' + lado(p) + ')', Math.abs(a.polos[0] - p) + (lado(p) === lado(a.polos[0]) ? 0 : 50));
        }
      }
      if (!best) break;
      cs = best.prueba; movs.push({ txt: best.txt, d: r2(des(carga(cs))) });
    }
    const cambios = cs.filter((x, i) => x.polos.join(',') !== base[i].polos.join(','))
      .map(x => { const o = base.find(b => b.c === x.c); return { id: x.c.id, descripcion: x.carga.descripcion, de: o.polos, a: x.polos, revit: x.c.circuitoRevit || o.polos.join(',') }; });
    return { movs, cambios, antes: r2(antes), despues: r2(des(carga(cs))), fasesAntes: carga(base), fasesDespues: carga(cs), resultado: cs.map(x => ({ id: x.c.id, polos: x.polos })) };
  }

  /* ---------- importación de Revit: elegir detalle de carga ---------- */
  const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  const CLAVES = [
    [/ilum|lumin|luz|luces|lamp/, 'ilum'], [/ups/, 'ups'], [/toma|tc\b|recept|enchuf|cortiner/, 'toma'], [/ascensor|elevador/, 'elevador'],
    [/microondas|horno/, 'horno'], [/cocina|estufa|plantilla/, 'cocina'], [/lavaplato|lavavajilla/, 'lavaplatos'], [/facp|incendio/, 'incendio'], [/acceso/, 'acceso'],
    [/seguridad|cctv/, 'seguridad'], [/rack|\bti\b|datos/, 'rack'], [/extract/, 'extractor'], [/inyect/, 'inyector'], [/bomba/, 'bomba'],
    [/calentador|tanque agua|termo/, 'calentador'], [/aire|a\/?c|minisplit|split|fan ?coil|condensad|cu-|ah-|ms-|unidad/, 'clima'],
    [/tablero|panel/, 'tablero'], [/secaman/, 'secamanos'], [/cargador/, 'cargador'], [/porton/, 'porton'], [/refri/, 'refrigerador'],
  ];
  function detalleRevit(nombre, voltaje, polos, catalog) {
    K = catalog || K;
    const v = n(voltaje), p = n(polos) || 1, nm = norm(nombre);
    const cand = K.detallesCarga.filter(d => n(d.v) === v && n(d.fases) === p);
    const pool = cand.length ? cand : K.detallesCarga.filter(d => n(d.fases) === p);
    const clave = (CLAVES.find(([re]) => re.test(nm)) || [])[1];
    let best = null, bestS = -1;
    pool.forEach(d => {
      const dn = norm(d.descripcion); let s = 0;
      if (clave && (CLAVES.find(([re, k]) => k === clave && re.test(dn)))) s += 5;
      nm.split(/[^a-z0-9]+/).filter(w => w.length > 3).forEach(w => { if (dn.includes(w)) s += 2; });
      if (n(d.v) === v) s += 1;
      if (s > bestS) { bestS = s; best = d; }
    });
    if (bestS < 2) { // sin coincidencias: carga genérica (tomas en 1 polo, equipos en 2 o 3)
      const gen = p === 1 ? /^tomas$/ : p === 3 ? /^equipos mecanicos$/ : /^(equipos|tomas 208v 1f)$/;
      best = pool.find(d => gen.test(norm(d.descripcion).trim())) || best;
    }
    return { detalle: best, seguro: !!cand.length && bestS >= 5 };
  }

  g.Calc = { proyecto, tablero, circuito, balanceo, detalleRevit, fasePolo, nombreTablero, vLN, famBreaker, famTablero, agrupAuto, setCatalog: c => { K = c; } };
})(typeof window !== 'undefined' ? window : globalThis);
