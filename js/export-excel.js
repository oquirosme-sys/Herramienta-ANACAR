/* Exporta a Excel (.xlsx real, sin librerías externas): TABLA RESUMEN (horizontal), VERTICAL, DU, ORIGINAL
   y un cuadro de cargas por tablero (formato TABLEROS 3F / 1F) para insertar en los planos de Revit. */
(function (g) {
  'use strict';
  const enc = new TextEncoder();
  const xe = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  const colName = n => { let s = ''; for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };

  /* ---------- zip mínimo (sin compresión) ---------- */
  const CRC = (() => { const t = []; for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  function zip(files) {
    const parts = [], central = []; let off = 0;
    files.forEach(f => {
      const name = enc.encode(f.name), data = typeof f.data === 'string' ? enc.encode(f.data) : f.data, crc = crc32(data);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 0, true); lh.setUint16(10, 0, true); lh.setUint16(12, 0x21, true);
      lh.setUint32(14, crc, true); lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true); lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
      parts.push(new Uint8Array(lh.buffer), name, data);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true); ch.setUint16(10, 0, true); ch.setUint16(12, 0, true); ch.setUint16(14, 0x21, true);
      ch.setUint32(16, crc, true); ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true); ch.setUint16(28, name.length, true); ch.setUint32(42, off, true);
      central.push(new Uint8Array(ch.buffer), name);
      off += 30 + name.length + data.length;
    });
    const csize = central.reduce((a, b) => a + b.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true); end.setUint32(12, csize, true); end.setUint32(16, off, true);
    return new Blob(parts.concat(central, [new Uint8Array(end.buffer)]), { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  /* ---------- estilos ---------- */
  function Styles() {
    const fonts = [{ b: 0, sz: 9, color: '' }], fills = [], xfs = [{}], idx = {};
    const fontIdx = o => { const k = JSON.stringify([o.b ? 1 : 0, o.sz || 9, o.color || '']); let i = fonts.findIndex(f => JSON.stringify([f.b, f.sz, f.color]) === k); if (i < 0) { fonts.push({ b: o.b ? 1 : 0, sz: o.sz || 9, color: o.color || '' }); i = fonts.length - 1; } return i; };
    const fillIdx = bg => { if (!bg) return 0; const c = bg.replace('#', '').toUpperCase(); let i = fills.indexOf(c); if (i < 0) { fills.push(c); i = fills.length - 1; } return i + 2; };
    return {
      get(o) { if (!o) return 0; const k = JSON.stringify(o); if (idx[k] !== undefined) return idx[k]; xfs.push({ f: fontIdx(o), fl: fillIdx(o.bg), o }); idx[k] = xfs.length - 1; return idx[k]; },
      xml() {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
          + '<fonts count="' + fonts.length + '">' + fonts.map(f => '<font>' + (f.b ? '<b/>' : '') + '<sz val="' + f.sz + '"/>' + (f.color ? '<color rgb="FF' + f.color.replace('#', '') + '"/>' : '') + '<name val="Arial"/></font>').join('') + '</fonts>'
          + '<fills count="' + (fills.length + 2) + '"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' + fills.map(c => '<fill><patternFill patternType="solid"><fgColor rgb="FF' + c + '"/><bgColor indexed="64"/></patternFill></fill>').join('') + '</fills>'
          + '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF8C8C8C"/></left><right style="thin"><color rgb="FF8C8C8C"/></right><top style="thin"><color rgb="FF8C8C8C"/></top><bottom style="thin"><color rgb="FF8C8C8C"/></bottom><diagonal/></border></borders>'
          + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
          + '<cellXfs count="' + xfs.length + '">' + xfs.map((x, i) => i === 0 ? '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            : '<xf numFmtId="' + (x.o.pct ? 9 : x.o.d2 ? 2 : 0) + '" applyNumberFormat="1" fontId="' + x.f + '" fillId="' + x.fl + '" borderId="' + (x.o.border === false ? 0 : 1) + '" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="' + (x.o.al || 'left') + '" vertical="center"' + (x.o.wrap ? ' wrapText="1"' : '') + (x.o.rot ? ' textRotation="' + x.o.rot + '"' : '') + '/></xf>').join('') + '</cellXfs>'
          + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
      },
    };
  }

  /* ---------- hoja ---------- */
  function Sheet(name, st) {
    const rows = [], merges = [], widths = []; let filter = null;
    const api = {
      name, picture: null,
      set(r, c, v, style) { (rows[r] = rows[r] || [])[c] = { v, s: st.get(style) }; return api; },
      merge(r1, c1, r2, c2) { merges.push(colName(c1) + (r1 + 1) + ':' + colName(c2) + (r2 + 1)); return api; },
      width(c, w) { widths[c] = w; return api; },
      autofilter(r1, c1, r2, c2) { filter = colName(c1) + (r1 + 1) + ':' + colName(c2) + (r2 + 1); return api; },
      xml() {
        const sd = rows.map((row, r) => row ? '<row r="' + (r + 1) + '">' + row.map((c, ci) => {
          if (!c) return ''; const ref = colName(ci) + (r + 1);
          if (c.v === null || c.v === undefined || c.v === '') return '<c r="' + ref + '" s="' + c.s + '"/>';
          if (typeof c.v === 'number') return '<c r="' + ref + '" s="' + c.s + '"><v>' + c.v + '</v></c>';
          return '<c r="' + ref + '" s="' + c.s + '" t="inlineStr"><is><t xml:space="preserve">' + xe(c.v) + '</t></is></c>';
        }).join('') + '</row>' : '').join('');
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
          + '<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews><sheetFormatPr defaultRowHeight="13"/>'
          + (widths.length ? '<cols>' + widths.map((w, i) => w ? '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>' : '').join('') + '</cols>' : '')
          + '<sheetData>' + sd + '</sheetData>' + (filter ? '<autoFilter ref="' + filter + '"/>' : '')
          + (merges.length ? '<mergeCells count="' + merges.length + '">' + merges.map(m => '<mergeCell ref="' + m + '"/>').join('') + '</mergeCells>' : '')
          + '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>' + (api.picture ? '<drawing r:id="rId1"/>' : '') + '</worksheet>';
      },
    };
    return api;
  }


  /* ---------- libro ---------- */
  const H = { b: 1, bg: '#D9D9D9', al: 'center', wrap: true };
  const HG = { b: 1, bg: '#BFBFBF', al: 'center', wrap: true };
  const T = { sz: 14, b: 1, border: false };
  const L = { b: 1, bg: '#EDEDED' };
  const N = { al: 'right', d2: true };
  const val = v => (typeof v === 'number' && !isFinite(v) ? '' : v);

  function build(R, opts) {
    opts = opts || {};
    const st = Styles(), sheets = [], used = {}, P = Store.project;
    const mk = n => { let b = n.replace(/[:\\/?*\[\]]/g, '-').slice(0, 28), nn = b, i = 2; while (used[nn.toLowerCase()]) nn = b.slice(0, 26) + '_' + i++; used[nn.toLowerCase()] = 1; const s = Sheet(nn, st); sheets.push(s); return s; };
    const put = (s, r, c, v, style) => s.set(r, c, val(v), typeof v === 'number' ? Object.assign({}, N, style || {}) : style || {});
    const encabezado = (s, titulo) => {
      s.set(0, 0, titulo, T);
      s.set(1, 0, [P.nombre, P.numero ? 'Proyecto N.º ' + P.numero : '', P.fecha].filter(Boolean).join(' · '), { border: false });
    };
    const lista = Resumen.ordenados(R);

    if (!opts.soloTipo) {
      /* TABLA RESUMEN horizontal: mismas columnas que la hoja del Excel (B:AU) */
      const s = mk('TABLA RESUMEN'); encabezado(s, 'TABLA RESUMEN - TABLEROS ELÉCTRICOS');
      let c = 0;
      Resumen.grupos.forEach(([g, n]) => { s.set(3, c, g, HG); for (let k = 1; k < n; k++) s.set(3, c + k, '', HG); if (n > 1) s.merge(3, c, 3, c + n - 1); else { s.set(4, c, '', HG); s.merge(3, c, 4, c); } c += n; });
      Resumen.sub.forEach((t, i) => { if (t) s.set(4, i, t, H); });
      lista.forEach((r, i) => Resumen.horizontal(r).forEach((v, k) => put(s, 5 + i, k, v, k === 0 || k === 24 ? { b: 1 } : {})));
      Resumen.sub.forEach((t, i) => s.width(i, i === 0 || i === 1 || i === 24 ? 22 : i >= 25 ? 14 : 10));

      const vert = (nombre, titulo, fn) => {
        const v = mk(nombre); encabezado(v, titulo);
        if (!lista.length) return;
        const filas = lista.map(fn);
        v.set(3, 0, 'Tablero / Alimentadores', H); v.width(0, 34);
        lista.forEach((r, j) => { v.set(3, 1 + 2 * j, r.nombre, H); v.set(3, 2 + 2 * j, '', H); v.merge(3, 1 + 2 * j, 3, 2 + 2 * j); v.width(1 + 2 * j, 9); v.width(2 + 2 * j, 9); });
        filas[0].forEach((f, i) => {
          v.set(4 + i, 0, f[0], L);
          filas.forEach((fl, j) => {
            const x = fl[i];
            if (x.length > 2) { put(v, 4 + i, 1 + 2 * j, x[1], { al: 'right' }); put(v, 4 + i, 2 + 2 * j, x[2], { al: 'left' }); }
            else { put(v, 4 + i, 1 + 2 * j, x[1], { al: 'center' }); v.set(4 + i, 2 + 2 * j, '', {}); v.merge(4 + i, 1 + 2 * j, 4 + i, 2 + 2 * j); }
          });
        });
      };
      vert('RESUMEN VERTICAL', 'TABLA RESUMEN (VERTICAL)', Resumen.vertical);
      vert('RESUMEN DU', 'TABLA RESUMEN DU (VERTICAL)', Resumen.du);
      vert('RESUMEN ORIGINAL', 'TABLA RESUMEN ORIGINAL', Resumen.original);
    }

    /* Un cuadro de cargas por tablero (formato TABLEROS 3F / 1F) */
    R.orden.filter(r => !opts.soloTipo || r.tab.tipo === opts.soloTipo).forEach(r => {
      const s = mk(r.tab.tipo + ' ' + (r.tab.nombre || r.nombre)), A = r.alim, nf = r.fases === 3 ? 3 : 2;
      s.set(0, 0, r.nombre, T); s.set(1, 0, 'Alimentado desde: ' + (r.alimentadoDesde || '—'), { border: false });
      const bloque = (fila, titulo, pares) => {
        s.set(fila, 0, titulo, HG); for (let k = 1; k < pares.length; k++) s.set(fila, k, '', HG); s.merge(fila, 0, fila, pares.length - 1);
        pares.forEach(([l, v], k) => { s.set(fila + 1, k, l, H); put(s, fila + 2, k, v === null || v === undefined ? '' : v, { al: 'center' }); });
      };
      bloque(3, 'Datos de cálculos eléctricos', [['kVA conectados', r.W130], ['kVA demandados', A.L139], ['% reserva', (Number(r.tab.reserva) || 0) * 100], ['Factor demanda', A.R139], ['Factor diversidad', A.S139], ['Factor potencia', A.T139], ['Amperios fase A', A.AA139[0]], ['Amperios fase B', A.AA139[1]], ['Amperios fase C', nf === 3 ? A.AA139[2] : ''], ['Icc disponible (kA)', r.iccKA || '']]);
      bloque(7, 'Especificación del tablero', [['Tensión nominal (V)', r.tab.sistema], ['Fases', r.fases], ['Hilos', r.hilos], ['Modelo de referencia', r.cat ? r.cat.modelo : ''], ['Fabricante', r.marca], ['Montaje', r.tab.montaje], ['Barras fase (A)', r.cat ? r.cat.barraFase : ''], ['Barras neutro (A)', r.cat ? r.cat.barraNeutro : ''], ['Barras tierra (A)', r.cat ? r.cat.barraTierra : ''], ['Espacios', r.cat ? r.cat.espacios : ''], ['% desbalance máximo', r.desbalance]]);
      bloque(11, 'Datos del alimentador / acometida', [['Fases (AWG)', A.fasesTxt], ['Neutro (AWG)', A.neutroTxt], ['Tierra (AWG)', A.tierraTxt], ['Material', A.mat], ['Aislamiento', A.ais], ['Tubería (Ø mm)', (A.preN || '') + A.AR139], ['Longitud (m)', A.M139], ['Voltaje (V)', A.AW139], ['Caída voltaje (V)', A.AX139], ['Caída voltaje total (%)', A.AY139]]);
      const bk = r.bkMain || {}, spd = r.spd || {};
      bloque(15, 'Interruptor principal y supresor', [['Modelo interruptor', bk.modelo || ''], ['Amperaje (A)', A.AF139 || ''], ['Marco (A)', Number(bk.marco) ? bk.marco : ''], ['Tipo', bk.unidad || ''], ['Polos', bk.polos || ''], ['SCCR (kA)', bk.sccr || ''], ['Modelo SPD', spd.modelo || ''], ['Capacidad SPD (kA)', spd.kaLL || ''], ['Montaje SPD', spd.montaje || '']]);
      const hdr = ['ID de circuito', 'Descripción de carga', 'kVA conectados', 'Caída voltaje total (%)', 'Modelo', 'Marco (A)', 'Amperios (A)', 'Tipo de unidad', 'Polos', 'SCCR (kA)', 'Fases', 'Neutro', 'Tierra', 'Tipo', 'Aislamiento', 'Tubo EMT (Ø mm)'].concat(['Fase A', 'Fase B', 'Fase C'].slice(0, nf));
      hdr.forEach((t, k) => s.set(19, k, t, H));
      const filas = Cuadro.filasCircuitos(r);
      filas.forEach((x, i) => [x.id, x.desc, x.kva, x.cv === null ? '' : x.cv, x.bkModelo, x.marco, x.amp, x.unidad, x.polosBk, x.sccr, x.fasesTxt, x.neutroTxt, x.tierraTxt, x.mat, x.ais, x.tubo].concat(x.fase.slice(0, nf).map(v => v || '')).forEach((v, k) => put(s, 20 + i, k, v, k === 1 ? { wrap: true } : {})));
      const z = 20 + filas.length;
      s.set(z, 0, 'kVA conectados sin reserva', L); s.set(z, 1, '', L); s.merge(z, 0, z, 1); put(s, z, 2, r.J116, { b: 1 }); r.U116.slice(0, nf).forEach((v, k) => put(s, z, 16 + k, v, { b: 1 }));
      s.set(z + 2, 0, Store.catalog.leyendaUnidades || '', { border: false, wrap: true }); s.merge(z + 2, 0, z + 2, 15 + nf);
      [16, 34, 9, 9, 14, 8, 9, 10, 6, 8, 10, 8, 8, 7, 10, 10, 9, 9, 9].forEach((w, k) => s.width(k, w));
    });

    const files = [
      { name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + sheets.map((s, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') + '</Types>' },
      { name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { name: 'xl/workbook.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + sheets.map((s, i) => '<sheet name="' + xe(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join('') + '</sheets></workbook>' },
      { name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + sheets.map((s, i) => '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('') + '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
    ];
    sheets.forEach((s, i) => files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: s.xml() }));
    files.push({ name: 'xl/styles.xml', data: st.xml() });
    return zip(files);
  }

  function download(R, opts) {
    if (!R || !R.orden.length) { UI.alert('No hay tableros para exportar.'); return; }
    const blob = build(R, opts);
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
    a.download = App.fileBase() + (opts && opts.soloTipo ? ' - tableros ' + opts.soloTipo : ' - tablas resumen') + '.xlsx';
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
    UI.toast('Excel generado', 'ok');
  }
  g.ExportExcel = { build, download };
})(window);
