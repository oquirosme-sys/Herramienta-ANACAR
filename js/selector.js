/* Selector de interruptores de caja moldeada: Eaton Power Defense (PD) y Schneider PowerPacT (PP).
   Reproduce la lógica de los libros Selector_Breakers_PowerDefense.xlsx y Selector_Breakers_Schneider_PowerPacT.xlsx:
     · Selección automática: menor amperaje ≥ corriente requerida → menor marco → menor capacidad interruptiva (kA) que cumpla la Icc a la tensión.
     · Código de catálogo completo.
     · Verificación de bornes contra el cableado (calibre, conductores en paralelo y material); solo avisa, no cambia el cálculo.
   Datos en data/selectores.js (window.SELECTORES). Campos de cada fila:
     PD cat : [id, cert, marco, polos(díg.), cap, amperaje(texto), amperaje(núm), H, familia, disparos, disparo def., kA240, kA480, kA600, CL]
     PD term: [id, marco, amp desde, amp hasta, H, rango, (tipo, cable) × 4 grupos, kits × 4, (n máx, calibre mín, calibre máx) × 4 grupos]
     PD alt : [id, marco, polos, amp, H, cert, grupo, kA240, kA480, kA600, cond. máx, calibre mín, calibre máx, acepta Al, grupo, tipo, rango]
     PP cat : [id, marco, rango marco, polos, cap, amperaje, régimen %, familia, disparos, disparo def., kA240, kA480, kA600, Vmax, díg. tensión, díg. amperaje]
     PP term: [id, marco, amp desde, amp hasta, grupo, nombre grupo, tipo, cond. máx por zapata, calibre mín, calibre máx 1 cond., calibre máx varios, material, rango, kit, nota]
     PP alt : [id, marco, rango, polos, amp, régimen, grupo, fila term, kA240, kA480, kA600, Vmax, cond. máx, calibre mín, calibre máx 1, calibre máx varios, material, grupo, rango, kit]
   Los calibres se indexan 1…21 (14 AWG … 1000 kcmil). */
(function (g) {
  'use strict';
  const D = () => g.SELECTORES || { pd: { cat: [], term: [], alt: [] }, pp: { cat: [], term: [], alt: [] }, calibres: [] };
  const num = v => (v === null || v === undefined || v === '' || isNaN(Number(v)) ? 0 : Number(v));

  /** Índice 1…21 del calibre (700, 800 y 900 kcmil se llevan al siguiente comercial: 750 y 1000). */
  function calIdx(cal) {
    const c = String(cal || '').replace(/\s*(AWG|kcmil|MCM)/i, '').trim();
    const L = D().calibres, i = L.indexOf(c);
    if (i >= 0) return i + 1;
    const x = Number(c); if (!x) return 0;
    const sup = L.findIndex(y => Number(y) >= x && /^\d+$/.test(y));
    return sup >= 0 ? sup + 1 : L.length;
  }
  const vIdx = V => (V <= 240 ? 0 : V <= 480 ? 1 : 2);
  const vTxt = V => (V <= 240 ? 240 : V <= 480 ? 480 : 600);
  const minPor = (lista, clave) => lista.reduce((m, r) => { const k = clave(r); return !m || k < m.k ? { k, r } : m; }, null);

  /* ---------- Power Defense ---------- */
  const PD_DISPARO_FAM = { TM: 'TM', EL: 'EL', MCS: 'MCS', FR: 'FR' };
  function pickPD(o) {
    const d = D().pd, V = num(o.V) || 480, kc = 11 + vIdx(V), kreq = num(o.kA), I = num(o.I), polos = num(o.polos) || 3;
    const cert = o.rated100 ? 'F' : 'G';
    const base = d.cat.filter(r => r[1] === cert && r[3] === polos && r[6] >= I && r[7] === 'N' && (!o.marcos || o.marcos.includes(r[2])));
    const buscar = fam => {
      const l = base.filter(r => r[8] === fam);
      const ok = l.filter(r => num(r[kc]) >= kreq);
      return { ok, l };
    };
    const orden = o.disparo && o.disparo !== 'AUTO' ? [PD_DISPARO_FAM[o.disparo] || 'EL'] : ['TM', 'EL'];
    let hallado = null, sccrBajo = false;
    for (const fam of orden) { const b = buscar(fam), m = minPor(b.ok, r => r[6] * 1e9 + r[2] * 1e8 + num(r[kc]) * 1e5 + r[0]); if (m) { hallado = m.r; break; } }
    if (!hallado) {   // ninguna llega a la Icc: la de mayor capacidad entre las de menor amperaje (se marca SCCR bajo)
      for (const fam of orden) {
        const l = buscar(fam).l; if (!l.length) continue;
        const a = Math.min(...l.map(r => r[6])), m = l.filter(r => r[6] === a).sort((x, y) => num(y[kc]) - num(x[kc]) || x[2] - y[2])[0];
        if (m) { hallado = m; sccrBajo = true; break; }
      }
    }
    if (!hallado) return null;
    const r = hallado, grupo = o.grupo || 1;
    const X = r[2] >= 5 ? 'N' : (d.terminales[grupo - 1] || d.terminales[0])[0];
    const ampTxt = typeof r[5] === 'number' ? String(r[5]).padStart(4, '0') : r[5];
    const marcoAmp = Math.max(...d.cat.filter(x => x[2] === r[2]).map(x => x[6]));
    return {
      tipo: 'pd', marca: 'Eaton', id: r[0], cert, marco: r[2], marcoTxt: 'PD-' + r[2], marcoAmp, polosDig: r[3], cap: r[4], amp: r[6], ampTxt: r[5], H: r[7], fam: r[8], disparo: r[10], terminal: X,
      kA: num(r[kc]), kAV: vTxt(V), CL: r[14] === 'CL', grupo, sccrBajo,
      codigo: 'PD' + cert + r[2] + r[3] + r[4] + ampTxt + r[10] + X,
      desc: 'PD-' + r[2] + ' · ' + (r[3] === 4 ? '4P' : r[3] + 'P') + ' · ' + r[6] + ' A · cap. ' + r[4] + ' · ' + (d.disparos[r[10]] || r[10]),
    };
  }
  function bornesPD(s, c) {
    const d = D().pd, size = calIdx(c.cal), nPar = Math.max(1, num(c.par) || 1), Al = /AL/i.test(c.mat || '');
    if (!size) return null;
    const fila = d.term.filter(t => t[1] === s.marco && t[2] <= s.amp && s.amp <= t[3] && t[4] === s.H).sort((a, b) => a[0] - b[0])[0];
    if (!fila) return { ok: null, msg: 'sin datos de bornes para ' + s.marcoTxt + ' ' + s.amp + ' A' };
    const grupo = c.grupo || s.grupo || 1, nombres = d.grupos;
    const prueba = k => { const n = fila[18 + 3 * (k - 1)], mn = fila[19 + 3 * (k - 1)], mx = fila[20 + 3 * (k - 1)]; return n > 0 && nPar <= n && size >= mn && size <= mx && (!Al || /\(Al\)/.test(fila[6 + 2 * (k - 1)])); };
    const rango = k => fila[6 + 2 * (k - 1)] + ' ' + fila[7 + 2 * (k - 1)];
    const act = fila[18 + 3 * (grupo - 1)] > 0 ? rango(grupo) : 'N/A';
    if (prueba(grupo)) return { ok: true, actual: nombres[grupo - 1] + ': ' + act };
    let rec = null; for (let k = 1; k <= 4; k++) if (prueba(k)) { rec = k; break; }
    const base = nPar + ' × ' + c.cal + ' ' + (Al ? 'Al' : 'Cu') + ' no cabe en los bornes ' + nombres[grupo - 1] + ' de ' + s.marcoTxt + ' ' + s.amp + ' A (' + act + ')';
    if (rec) {
      const letra = (d.terminales[rec - 1] || [])[0];
      return { ok: false, msg: base + '. Use bornes ' + nombres[rec - 1] + ' (' + rango(rec) + ', terminal ' + letra + ')', sugerido: { grupo: rec, terminal: letra } };
    }
    const kc = 7 + vIdx(num(c.V) || 480);
    const alt = minPor(d.alt.filter(a => a[2] === s.polosDig && a[3] === s.amp && a[4] === s.H && a[5] === s.cert && num(a[kc]) >= s.kA && a[10] >= nPar && a[11] <= size && a[12] >= size && (!Al || a[13] === 1)), a => a[1] * 1e5 + a[0]);
    if (alt) return { ok: false, msg: base + '. Cambie a marco PD-' + alt.r[1] + ' (' + alt.r[14] + ': ' + alt.r[15] + ' ' + alt.r[16] + ')', sugerido: { marco: alt.r[1] } };
    return { ok: false, msg: base + '. Ningún bornes del catálogo lo cubre: menos paralelos, calibre mayor con menos paralelos, o barra/conector externo' };
  }

  /* ---------- Schneider PowerPacT ---------- */
  function pickPP(o) {
    const d = D().pp, V = num(o.V) || 480, kc = 10 + vIdx(V), kreq = num(o.kA), I = num(o.I), polos = num(o.polos) || 3, reg = o.rated100 ? 100 : 80;
    const base = d.cat.filter(r => r[6] === reg && r[3] === polos && r[5] >= I && r[13] >= V && (!o.marcos || o.marcos.includes(r[1])));
    const buscar = fam => { const l = base.filter(r => r[7] === fam); return { l, ok: l.filter(r => num(r[kc]) >= kreq) }; };
    const orden = o.disparo && o.disparo !== 'AUTO' ? [o.disparo === 'TM' ? 'TM' : 'EL'] : ['TM', 'EL'];
    let hallado = null, sccrBajo = false;
    for (const fam of orden) { const b = buscar(fam), m = minPor(b.ok, r => r[5] * 1e9 + r[2] * 1e8 + num(r[kc]) * 1e5 + r[0]); if (m) { hallado = m.r; break; } }
    if (!hallado) {
      for (const fam of orden) {
        const l = buscar(fam).l; if (!l.length) continue;
        const a = Math.min(...l.map(r => r[5])), m = l.filter(r => r[5] === a).sort((x, y) => num(y[kc]) - num(x[kc]) || x[2] - y[2])[0];
        if (m) { hallado = m; sccrBajo = true; break; }
      }
    }
    if (!hallado) return null;
    const r = hallado, grupo = o.grupo === 0 ? 0 : (o.grupo || 1);
    const ter = o.terminal || ((r[1] === 'R' || grupo === 0) ? 'F' : (grupo > 1 && r[1] !== 'B') ? 'F' : 'L');
    const suf = r[1] === 'B' && ter === 'L' && grupo === 2 ? 'LH' : r[1] === 'B' && ter === 'L' && grupo === 3 ? 'LC' : '';
    const disp = r[9];
    const marcoAmp = Math.max(...d.cat.filter(x => x[1] === r[1]).map(x => x[5]));
    return {
      tipo: 'pp', marca: 'Square D', id: r[0], marco: r[1], marcoTxt: r[1] + '-frame', marcoAmp, polosDig: r[3], cap: r[4], amp: r[5], reg, fam: r[7], disparo: disp, terminal: ter,
      kA: num(r[kc]), kAV: vTxt(V), grupo: grupo || 1, sccrBajo,
      codigo: r[1] + (r[4] === '-' ? '' : r[4]) + ter + r[3] + (r[14] == null ? '' : r[14]) + (r[15] == null ? '' : r[15]) + (reg === 100 ? 'C' : '') + (disp === 'TM' || disp === 'BASICO' ? '' : disp) + suf,
      desc: 'Marco ' + r[1] + ' · ' + r[3] + 'P · ' + r[5] + ' A · ' + (r[4] === '-' ? 'cap. única' : 'cap. ' + r[4]) + ' · ' + reg + '% · ' + (d.disparos[disp] || disp),
    };
  }
  function bornesPP(s, c) {
    const d = D().pp, size = calIdx(c.cal), nPar = Math.max(1, num(c.par) || 1), Al = /AL/i.test(c.mat || '');
    if (!size) return null;
    const fila = k => d.term.filter(t => t[1] === s.marco && t[2] <= s.amp && s.amp <= t[3] && t[4] === k).sort((a, b) => a[0] - b[0])[0];
    const prueba = t => t && t[7] > 0 && nPar <= t[7] && size >= t[8] && size <= (nPar > 1 ? t[10] : t[9]) && (!Al || t[11] === 'AlCu');
    const grupo = c.grupo || s.grupo || 1, ta = fila(grupo);
    if (!ta) return { ok: null, msg: 'sin datos de bornes del grupo ' + grupo + ' para ' + s.marcoTxt + ' ' + s.amp + ' A' };
    if (prueba(ta)) return { ok: true, actual: ta[5] + ': ' + ta[6] + ' ' + ta[12] };
    const base = nPar + ' × ' + c.cal + ' ' + (Al ? 'Al' : 'Cu') + ' no cabe en las zapatas ' + ta[5] + ' de ' + s.marcoTxt + ' ' + s.amp + ' A (' + ta[12] + ')';
    for (let k = 1; k <= 5; k++) { const t = fila(k); if (prueba(t)) return { ok: false, msg: base + '. Use ' + t[5] + ' (' + t[12] + ') · Cat. No. ' + t[13], sugerido: { grupo: k, kit: t[13] } }; }
    const kc = 8 + vIdx(num(c.V) || 480);
    const alt = minPor(d.alt.filter(a => a[3] === s.polosDig && a[4] === s.amp && a[5] === s.reg && num(a[kc]) >= s.kA && a[11] >= (num(c.V) || 480) && a[12] >= nPar && a[13] <= size && (nPar > 1 ? a[15] : a[14]) >= size && (!Al || a[16] === 'AlCu')), a => a[2] * 1e5 + a[0]);
    if (alt) return { ok: false, msg: base + '. Cambie a marco ' + alt.r[1] + ' (' + alt.r[17] + ': ' + alt.r[18] + ') · Cat. No. ' + alt.r[19], sugerido: { marco: alt.r[1] } };
    return { ok: false, msg: base + '. Ningún kit del catálogo lo cubre: menos paralelos, calibre mayor con menos paralelos, o barra/conector externo' };
  }

  g.Selector = {
    /** o: { V, I (A requeridos), kA (Icc), polos, rated100, disparo ('TM' | 'EL' | 'AUTO'), marcos (opcional), grupo, terminal }. */
    pick(tipo, o) { return tipo === 'pd' ? pickPD(o) : tipo === 'pp' ? pickPP(o) : null; },
    /** c: { cal, par, mat, V }: compara el cableado con el rango de bornes del interruptor elegido. null = no aplica. */
    bornes(s, c) { return !s ? null : s.tipo === 'pd' ? bornesPD(s, c) : bornesPP(s, c); },
    calIdx,
  };
})(typeof window !== 'undefined' ? window : globalThis);
