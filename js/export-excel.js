/* Exporta a Excel (.xlsx real, sin librerías externas) con el mismo diseño de las hojas del Excel original:
   TABLA RESUMEN, (VERTICAL), DU(VERTICAL), ORIGINAL y una hoja TABLEROS 3F / 1F por tablero, con el logo. */
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

  /* ---------- estilos (fuente, relleno, bordes, alineación y formato numérico) ---------- */
  function Styles() {
    const fonts = [{ fn: 'Calibri', sz: 11, b: 0, color: '' }], fills = [], xfs = [{}], idx = {};
    const NUMFMT = { 2: 2, 1: 164, pct: 10 };
    const fontIdx = o => {
      const f = { fn: o.fn || 'Century Gothic', sz: o.sz || 12, b: o.b ? 1 : 0, color: (o.color || '').replace('#', '') };
      const k = JSON.stringify(f); let i = fonts.findIndex(x => JSON.stringify(x) === k); if (i < 0) { fonts.push(f); i = fonts.length - 1; } return i;
    };
    const fillIdx = bg => { if (!bg) return 0; const c = bg.replace('#', '').toUpperCase(); let i = fills.indexOf(c); if (i < 0) { fills.push(c); i = fills.length - 1; } return i + 2; };
    const brd = b => (b === 'none' ? 0 : b === 'medium' ? 2 : 1);
    return {
      get(o) {
        if (!o) return 0; const k = JSON.stringify(o); if (idx[k] !== undefined) return idx[k];
        xfs.push({ f: fontIdx(o), fl: fillIdx(o.bg), b: brd(o.border), nf: NUMFMT[o.nf] || 0, o }); idx[k] = xfs.length - 1; return idx[k];
      },
      xml() {
        const B = s => (s ? '<left style="' + s + '"><color rgb="FF000000"/></left><right style="' + s + '"><color rgb="FF000000"/></right><top style="' + s + '"><color rgb="FF000000"/></top><bottom style="' + s + '"><color rgb="FF000000"/></bottom><diagonal/>' : '<left/><right/><top/><bottom/><diagonal/>');
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
          + '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0"/></numFmts>'
          + '<fonts count="' + fonts.length + '">' + fonts.map(f => '<font>' + (f.b ? '<b/>' : '') + '<sz val="' + f.sz + '"/>' + (f.color ? '<color rgb="FF' + f.color + '"/>' : '') + '<name val="' + xe(f.fn) + '"/></font>').join('') + '</fonts>'
          + '<fills count="' + (fills.length + 2) + '"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' + fills.map(c => '<fill><patternFill patternType="solid"><fgColor rgb="FF' + c + '"/><bgColor indexed="64"/></patternFill></fill>').join('') + '</fills>'
          + '<borders count="3"><border>' + B('') + '</border><border>' + B('thin') + '</border><border>' + B('medium') + '</border></borders>'
          + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
          + '<cellXfs count="' + xfs.length + '">' + xfs.map((x, i) => i === 0 ? '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            : '<xf numFmtId="' + x.nf + '" fontId="' + x.f + '" fillId="' + x.fl + '" borderId="' + x.b + '" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="' + (x.o.al || 'left') + '" vertical="center"' + (x.o.wrap ? ' wrapText="1"' : '') + '/></xf>').join('') + '</cellXfs>'
          + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
      },
    };
  }

  /* ---------- hoja ---------- */
  function Sheet(name, st) {
    const rows = [], merges = [], widths = [], heights = [];
    const api = {
      name, pictures: [],
      set(r, c, v, style) { (rows[r] = rows[r] || [])[c] = { v, s: st.get(style) }; return api; },
      merge(r1, c1, r2, c2) { merges.push(colName(c1) + (r1 + 1) + ':' + colName(c2) + (r2 + 1)); return api; },
      width(c, w) { widths[c] = w; return api; },
      height(r, pt) { heights[r] = pt; return api; },
      widths, heights,
      xml(drawRid) {
        const n = Math.max(rows.length, heights.length), sd = [];
        for (let r = 0; r < n; r++) {
          const row = rows[r] || [];
          const cs = row.map((c, ci) => {
            if (!c) return ''; const ref = colName(ci) + (r + 1);
            if (c.v === null || c.v === undefined || c.v === '') return '<c r="' + ref + '" s="' + c.s + '"/>';
            if (typeof c.v === 'number') return '<c r="' + ref + '" s="' + c.s + '"><v>' + c.v + '</v></c>';
            return '<c r="' + ref + '" s="' + c.s + '" t="inlineStr"><is><t xml:space="preserve">' + xe(c.v) + '</t></is></c>';
          }).join('');
          if (cs || heights[r]) sd.push('<row r="' + (r + 1) + '"' + (heights[r] ? ' ht="' + heights[r] + '" customHeight="1"' : '') + '>' + cs + '</row>');
        }
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
          + '<sheetViews><sheetView workbookViewId="0" showGridLines="0" zoomScale="60" zoomScaleNormal="60"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
          + (widths.length ? '<cols>' + Array.from(widths, (w, i) => (w ? '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>' : '')).join('') + '</cols>' : '')
          + '<sheetData>' + sd.join('') + '</sheetData>'
          + (merges.length ? '<mergeCells count="' + merges.length + '">' + merges.map(m => '<mergeCell ref="' + m + '"/>').join('') + '</mergeCells>' : '')
          + '<pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/><pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>' + (drawRid ? '<drawing r:id="' + drawRid + '"/>' : '') + '</worksheet>';
      },
    };
    return api;
  }

  /** Copia una cuadrícula de Hoja.* a una hoja de Excel (mismos estilos, celdas combinadas, anchos y altos). */
  function volcar(s, G) {
    G.anchos.forEach((w, i) => s.width(i, w));
    G.altos.forEach((a, i) => { if (a) s.height(i, a); });
    G.cells.forEach(c => {
      const base = Hoja.ESTILOS[c.s] || {};
      const st = Object.assign({}, base, typeof c.v === 'number' && c.f !== null && c.f !== undefined ? { nf: c.f } : {});
      for (let y = c.y; y <= c.y2; y++) for (let x = c.x; x <= c.x2; x++) if (y !== c.y || x !== c.x) s.set(y, x, '', st);
      s.set(c.y, c.x, c.s === 'logo' ? '' : c.v, st);
      if (c.y2 > c.y || c.x2 > c.x) s.merge(c.y, c.x, c.y2, c.x2);
      if (c.s === 'logo') s.pictures.push(c);
    });
  }

  /** Convierte el reporte de memoria (encabezados, párrafos, tablas y listas) en una cuadrícula para Excel. */
  function memoriaGrid(r) {
    const G = Hoja.Grid([2, 34].concat(Array(18).fill(13))); let y = 0;
    const num = s => { const t = String(s).trim(); return /^-?[\d.]+(,\d+)?$/.test(t) ? Number(t.replace(/\./g, '').replace(',', '.')) : null; };
    const visitar = el => Array.from(el.children).forEach(ch => {
      const tag = ch.tagName;
      if (/^H[1-4]$/.test(tag)) { if (y) y++; G.put(y, 1, y, 12, ch.textContent, 'hdoc'); G.alto(y, tag === 'H1' || tag === 'H2' ? 22 : 18); y++; }
      else if (tag === 'P') { G.put(y, 1, y, 12, ch.textContent, 'pdoc'); y++; }
      else if (tag === 'UL') { Array.from(ch.children).forEach(li => { G.put(y, 1, y, 12, '• ' + li.textContent, 'pdoc'); y++; }); }
      else if (tag === 'TABLE') {
        const kv = ch.classList.contains('kv2');
        Array.from(ch.querySelectorAll('tr')).forEach(tr => {
          const cs = Array.from(tr.children);
          cs.forEach((c, i) => {
            const v = c.tagName === 'TD' && c.classList.contains('n') ? num(c.textContent) : null;
            if (kv && i === 1) G.put(y, 2, y, 8, c.textContent, 'tdd');
            else G.put(y, 1 + i, y, 1 + i, v !== null ? v : c.textContent, c.tagName === 'TH' ? (kv ? 'tdd' : 'thd') : v !== null ? 'tddn' : 'tdd', v !== null ? 2 : null);
          });
          y++;
        });
        y++;
      } else if (tag !== 'IMG') visitar(ch);
    });
    visitar(Reporte.documento([r]));
    return G;
  }

  /* ---------- libro ---------- */
  function build(R, opts, logo) {
    opts = opts || {};
    const st = Styles(), sheets = [], used = {};
    const mk = n => { let b = n.replace(/[:\\/?*\[\]]/g, '-').slice(0, 28), nn = b, i = 2; while (used[nn.toLowerCase()]) nn = b.slice(0, 26) + '_' + i++; used[nn.toLowerCase()] = 1; const s = Sheet(nn, st); sheets.push(s); return s; };
    const lista = Resumen.ordenados(R);
    // nombres definidos (como la macro AsignarNombresTableros) para vincular las tablas en Revit
    const nombres = [], usadosN = {};
    const ref = (s, G, c1, c2) => "'" + s.name.replace(/'/g, "''") + "'!$" + colName(c1) + '$' + (G.filas() ? 1 : 1) + ':$' + colName(c2) + '$' + G.filas();
    const nombre = (n, s, G, c1, c2) => { let k = n.replace(/[^A-Za-z0-9_ÁÉÍÓÚÑáéíóúñ]/g, '_'); if (/^\d/.test(k)) k = 'T_' + k; let kk = k, i = 2; while (usadosN[kk.toUpperCase()]) kk = k + '_' + i++; usadosN[kk.toUpperCase()] = 1; nombres.push([kk, ref(s, G, c1, c2)]); };
    const limpio = t => String(t || '').replace(/^TABLERO\s+/i, '').trim().replace(/[\s\-./\\]/g, '_').replace(/[^A-Za-z0-9_]/g, '');
    if (opts.memoria) {
      R.orden.filter(r => !opts.ids || opts.ids.includes(r.tab.id)).forEach(r => volcar(mk('MEMORIA ' + (r.tab.nombre || r.nombre)), memoriaGrid(r)));
    } else if (!opts.soloTipo) {
      const sr = mk('TABLA RESUMEN'), gr = Hoja.resumen(lista, Store.project.resumenSimple); volcar(sr, gr);
      nombre('TABLA_RESUMEN__TABLEROS_ELÉCTRICOS', sr, gr, Hoja.col('B'), Hoja.col('Y'));
      if (!Store.project.resumenSimple) { nombre('DATOS_DEL_TABLERO', sr, gr, Hoja.col('Z'), Hoja.col('AG')); nombre('DATOS_DEL_SUPRESOR', sr, gr, Hoja.col('AH'), Hoja.col('AN')); nombre('DATOS_INTERRUPTOR_PRINCIPAL', sr, gr, Hoja.col('AO'), Hoja.col('AU')); }
      volcar(mk('TABLA RESUMEN (VERTICAL)'), Hoja.vertical(lista, Resumen.vertical, 'TABLA RESUMEN - TABLEROS ELÉCTRICOS'));
      volcar(mk('TABLA RESUMEN DU(VERTICAL)'), Hoja.vertical(lista, Resumen.du, 'TABLA RESUMEN DU - TABLEROS ELÉCTRICOS'));
      volcar(mk('TABLA RESUMEN ORIGINAL'), Hoja.vertical(lista, Resumen.original, 'TABLA RESUMEN - TABLEROS ELÉCTRICOS'));
    }
    if (!opts.memoria) R.orden.filter(r => !opts.soloTipo || r.tab.tipo === opts.soloTipo).forEach(r => {
      const s = mk(r.tab.tipo + ' ' + (r.tab.nombre || r.nombre)), G = Hoja.tablero(r); volcar(s, G);
      nombre('TABLERO_' + limpio(r.tab.nombre || r.nombre), s, G, Hoja.col('B'), Hoja.col('AJ'));
    });

    const files = [], ct = [], wbRels = [];
    let nDraw = 0;
    const EMU = 9525, pxW = w => Math.round((w || 8.43) * 7 + 5), pxH = pt => Math.round((pt || 15) * 96 / 72);
    sheets.forEach((s, i) => {
      let rid = '';
      if (logo && s.pictures.length) {
        nDraw++; rid = 'rId1';
        const anchors = s.pictures.map((c, k) => {
          let w = 0, hh = 0; for (let x = c.x; x <= c.x2; x++) w += pxW(s.widths[x]); for (let y = c.y; y <= c.y2; y++) hh += pxH(s.heights[y]);
          const ar = logo.w / logo.h; let iw = w * 0.85, ih = iw / ar; if (ih > hh * 0.85) { ih = hh * 0.85; iw = ih * ar; }
          let ox = (w - iw) / 2, oy = (hh - ih) / 2, cx = c.x, cy = c.y;
          while (cx < c.x2 && ox > pxW(s.widths[cx])) { ox -= pxW(s.widths[cx]); cx++; }
          while (cy < c.y2 && oy > pxH(s.heights[cy])) { oy -= pxH(s.heights[cy]); cy++; }
          return '<xdr:oneCellAnchor><xdr:from><xdr:col>' + cx + '</xdr:col><xdr:colOff>' + Math.round(ox * EMU) + '</xdr:colOff><xdr:row>' + cy + '</xdr:row><xdr:rowOff>' + Math.round(oy * EMU) + '</xdr:rowOff></xdr:from><xdr:ext cx="' + Math.round(iw * EMU) + '" cy="' + Math.round(ih * EMU) + '"/>'
            + '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + (k + 2) + '" name="Logo ' + (k + 1) + '"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
            + '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + Math.round(iw * EMU) + '" cy="' + Math.round(ih * EMU) + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>';
        }).join('');
        files.push({ name: 'xl/drawings/drawing' + nDraw + '.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' + anchors + '</xdr:wsDr>' });
        files.push({ name: 'xl/drawings/_rels/drawing' + nDraw + '.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>' });
        files.push({ name: 'xl/worksheets/_rels/sheet' + (i + 1) + '.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' + nDraw + '.xml"/></Relationships>' });
        ct.push('<Override PartName="/xl/drawings/drawing' + nDraw + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>');
      }
      files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: s.xml(rid) });
      ct.push('<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
      wbRels.push('<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>');
    });
    if (nDraw) files.push({ name: 'xl/media/image1.png', data: logo.bytes });
    files.unshift(
      { name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + ct.join('') + '</Types>' },
      { name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { name: 'xl/workbook.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + sheets.map((s, i) => '<sheet name="' + xe(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join('') + '</sheets>' + (nombres.length ? '<definedNames>' + nombres.map(([k, v]) => '<definedName name="' + xe(k) + '">' + xe(v) + '</definedName>').join('') + '</definedNames>' : '') + '</workbook>' },
      { name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + wbRels.join('') + '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
    );
    files.push({ name: 'xl/styles.xml', data: st.xml() });
    return zip(files);
  }

  let logoCache = null;
  async function cargarLogo() {
    if (logoCache) return logoCache;
    try {
      const bytes = new Uint8Array(await (await fetch('img/logo-sinergia.png')).arrayBuffer());
      const dv = new DataView(bytes.buffer);
      logoCache = { bytes, w: dv.getUint32(16), h: dv.getUint32(20) };
    } catch (e) { logoCache = null; } // abierto como archivo local: se exporta sin logo
    return logoCache;
  }

  async function download(R, opts) {
    if (!R || !R.orden.length) { UI.alert('No hay tableros para exportar.'); return; }
    const blob = build(R, opts, await cargarLogo());
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
    a.download = App.fileBase() + (opts && opts.memoria ? ' - memoria de cálculo' : opts && opts.soloTipo ? ' - tableros ' + opts.soloTipo : ' - tablas resumen y tableros') + '.xlsx';
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
    UI.toast('Excel generado', 'ok');
  }
  g.ExportExcel = { build, download, cargarLogo };
})(window);
