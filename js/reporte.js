/* Reporte de la memoria de cálculo (un tablero o todos): documento para imprimir / PDF, Word (.doc) y Excel.
   Contenido: datos del proyecto y del tablero, circuitos ramales, balance, factores de demanda, alimentador con su
   validación, fuentes (transformador, UPS, segunda acometida), cortocircuito y revisión NEC 2020. */
(function () {
  'use strict';
  const h = U.h, f2 = v => U.fmt(v, 2), f1 = v => U.fmt(v, 1), f0 = v => U.fmt(v, 0);
  const tabla = (cab, filas, cls) => h('table', { class: 'rt ' + (cls || '') }, h('thead', null, h('tr', null, cab.map(c => h('th', null, c)))), h('tbody', null, filas.map(f => h('tr', null, f.map(v => h('td', { class: typeof v === 'number' ? 'n' : '' }, typeof v === 'number' ? f2(v) : (v === null || v === undefined ? '' : v)))))));
  const kv = pares => h('table', { class: 'rt kv2' }, h('tbody', null, pares.filter(Boolean).map(([k, v]) => h('tr', null, h('th', null, k), h('td', null, v === null || v === undefined || v === '' ? '—' : v)))));

  function seccion(r) {
    const t = r.tab, A = r.alim, P = Store.project, nf = r.fases === 3 ? 3 : 2, cat = r.cat || {}, bk = r.bkMain || {};
    const circ = r.rows.map(x => [x.polos.join(','), x.descripcion, (x.det.id || '') + ' ' + (x.det.descripcion || ''), x.J, x.O || '', x.P || '', x.J ? Math.max(...x.I) : '', x.AE || '', x.AF || '',
      x.breaker ? (x.breaker.ref || x.breaker.modelo || '') + ' ' + (x.breaker.unidad || '') + (x.breaker.serie ? ' (serie ' + x.breaker.serie.kA + ' kA)' : '') : '', x.breaker ? x.breaker.sccr : '',
      x.AL ? x.fasesTxt.trim() + (x.neutroTxt ? ' + ' + x.neutroTxt : '') + (x.tierraTxt ? ' + ' + x.tierraTxt + 'T' : '') + ' ' + x.mat + ' ' + x.ais : '', x.tuboTxt, Number(x.c.longitud) || '', x.AY === null ? '' : x.AY]
      .concat(x.fase.slice(0, nf).map(v => v || '')));
    const val = A.val;
    return h('section', { class: 'rep-tab' },
      h('h2', null, 'Memoria de cálculo — ' + r.nombre),
      kv([['Alimentado desde', r.alimentadoDesde], ['Sistema', t.sistema + ' V · ' + r.fases + ' fases · ' + r.hilos + ' hilos'], ['Tipo de equipo', t.clase === 'subestacion' ? 'Subestación / switchboard' : 'Tablero'],
        ['Tablero', (cat.modelo || '—') + ' · ' + (r.marca || '') + (r.famT ? ' (familia ' + r.famT.familia + ')' : '') + ' · barras ' + (cat.barraFase || '—') + ' A · ' + (cat.espacios && cat.espacios < 999 ? cat.espacios + ' espacios' : '')],
        ['Principal', r.zapatas ? 'Zapatas (main lugs)' : (bk.ref || bk.modelo || '—') + ' · ' + (A.AF139 || '') + ' A · ' + (bk.unidad || '') + ' · ' + (bk.sccr || '') + ' kA'],
        ['Supresor (SPD)', r.spd ? r.spd.modelo + ' · ' + r.spd.kaLL + '/' + r.spd.kaLN + ' kA · ' + r.spd.montaje : 'No'], ['Montaje', t.montaje], ['Reserva', f0((Number(t.reserva) || 0) * 100) + ' %']]),
      h('h3', null, '1. Circuitos ramales'),
      tabla(['Pos.', 'Descripción', 'Detalle de carga', 'kVA', 'V', 'F', 'I (A)', 'Amp. req. (A)', 'Prot. (A)', 'Interruptor', 'SCCR (kA)', 'Conductores', 'Tubo (mm)', 'Long. (m)', 'ΔV total (%)'].concat(['kVA A', 'kVA B', 'kVA C'].slice(0, nf)), circ, 'chica'),
      h('h3', null, '2. Balance de cargas'),
      kv([['kVA por fase (A / B / C)', r.U116.slice(0, nf).map(f2).join(' / ')], ['kVA conectados sin reserva', f2(r.J116)], ['Desbalance máximo', f2(r.desbalance) + ' %']]),
      h('h3', null, '3. Factores de demanda y diversidad'),
      tabla(['Tipo de carga', 'Circuitos', 'kVA conectados', 'Reserva', 'kVA totales', 'Factor demanda', 'F. diversidad', 'kVA demandados'],
        r.tipos.filter(x => x.conectados).map(x => [x.tipo.nombre, x.n, x.conectados, x.reserva, x.total, x.fd, x.fdiv, x.demandados]).concat([['Totales', '', r.W128, r.W129, r.W130, A.R139, '', r.J134]])),
      h('h3', null, '4. Alimentador / acometida'),
      kv([['kVA totales / demandados', f2(r.W130) + ' / ' + f2(A.L139) + ' (FD ' + f2(A.R139) + ', FP ' + f2(A.T139) + ')'], ['Corriente por fase', A.AA139.slice(0, nf).map(f1).join(' / ') + ' A'],
        ['Ampacidad requerida (× ' + f2(A.AD139) + ')', f2(A.AE139) + ' A'], ['Protección', (A.AF139 || '—') + ' A'],
        ['Conductores', A.fasesTxt + (A.neutroTxt ? ' + ' + A.neutroTxt + ' N' : '') + (A.tierraTxt ? ' + ' + A.tierraTxt + ' T' : '') + ' AWG/kcmil ' + A.mat + ' ' + A.ais], ['Tubería', (A.preN || '') + A.AR139 + ' mm ' + A.tuberia],
        ['Conductor del electrodo (250.66)', A.AP140 ? A.AP140 + ' AWG' : '—'], ['Longitud', f1(A.M139) + ' m'],
        ['Voltaje de partida / en bornes', f2(A.vInicio) + ' V / ' + f2(A.AW139) + ' V'], ['Caída en el alimentador', f2(A.AV139) + ' V'], ['Caída acumulada', f2(A.AX139) + ' V · ' + f2(A.AY139) + ' %']]),
      val ? h('div', null, h('h4', null, 'Validación por temperatura y agrupamiento (310.15 / 310.16): ' + (val.ok ? 'CUMPLE' : 'NO CUMPLE')),
        kv([['Ampacidad a ' + val.tA + ' °C', f1(val.amp90) + ' A'], ['Factor de temperatura (' + (t.alim.tempAmb || '26-30') + ' °C)', f2(val.ft)], ['Factor de agrupamiento (' + A.agrup + ')', f2(val.fg)],
          ['Ampacidad corregida', f1(val.corr) + ' A'], ['Límite por bornes a ' + (t.alim.tempBorne || 75) + ' °C', f1(val.term * A.AI139) + ' A'], ['Ampacidad disponible', f1(val.cap) + ' A'],
          A.ajustado ? ['Ajuste', 'Calibre aumentado de ' + A.calBase + ' a ' + A.AL139] : null])) : null,
      h('h3', null, '5. Fuentes y cortocircuito'),
      kv([r.trafo ? ['Transformador', (r.trafo.nombre || '') + ' ' + f1(r.trafo.kva) + ' kVA · Z ' + f2(r.trafo.z) + ' % · X/R ' + f1(r.trafo.xr) + ' · ' + (r.trafo.Vp ? f0(r.trafo.Vp) + ' V' : '') + ' / ' + t.sistema + ' V · carga ' + f1(r.trafo.carga * 100) + ' % · ΔV ' + f2(r.trafo.reg) + ' % · Icc sec. ' + f2(r.trafo.iccSec / 1000) + ' kA'] : null,
        r.ups ? ['UPS', r.ups.nombre + ' ' + f1(r.ups.kva) + ' kVA · carga ' + f1(r.ups.carga * 100) + ' % · Icc ' + f2(r.ups.icc / 1000) + ' kA'] : null,
        r.alterna ? ['Segunda acometida (' + r.alterna.equipo + ')', (r.alterna.origen || '') + ' · ' + r.alterna.fasesTxt + ' ' + r.alterna.mat + ' ' + r.alterna.ais + ' · ' + f1(r.alterna.longitud) + ' m · ΔV ' + f2(r.alterna.cv) + ' % · Icc ' + (r.alterna.iccA ? f2(r.alterna.iccA / 1000) + ' kA' : '—')] : null,
        ['Icc en bornes', r.iccKA ? f2(r.iccKA) + ' kA (' + r.iccFuente + ')' : '—'], ['Método', 'Punto a punto (Bussmann); constantes C de la hoja DATOS'],
        r.enSerie && r.enSerie.length ? ['Clasificación en serie (240.86)', r.enSerie.length + ' ramal(es) en serie con el principal; rotular según 110.22(C)'] : null]),
      r.nec && r.nec.length ? h('div', null, h('h3', null, '6. Revisión NEC 2020'), h('ul', null, r.nec.map(x => h('li', null, h('b', null, x.art + ': '), x.txt)))) : null,
      r.avisos.length ? h('div', null, h('h3', null, 'Avisos'), h('ul', null, r.avisos.map(a => h('li', null, a)))) : null);
  }

  function documento(lista) {
    const P = Store.project;
    return h('div', { class: 'rep' },
      h('header', { class: 'rep-h' }, h('img', { src: 'img/logo-sinergia.png', alt: 'Sinergia Ingeniería' }),
        h('div', null, h('h1', null, 'Memoria de cálculo de tableros eléctricos'), h('p', null, [P.nombre, P.numero ? 'Proyecto N.º ' + P.numero : '', P.ubicacion].filter(Boolean).join(' · ')),
          h('p', null, 'Fecha: ' + (P.fecha || '') + (P.elaboro ? ' · Elaboró: ' + P.elaboro : '') + (P.revision ? ' · Revisión: ' + P.revision : '')))),
      h('section', null, h('h3', null, 'Criterios de diseño'),
        kv([['Normativa', 'NEC 2020 (NFPA 70)'], ['Caída de voltaje máxima', 'Alimentador ' + P.cvMaxAlim + ' % · total ' + P.cvMaxTotal + ' %'], ['Desbalance máximo', P.desbalanceMax + ' %'],
          ['Ocupación', P.ocupacion || 'comercial'], ['Conductores', 'Tabla 310.16 con corrección por temperatura 310.15(B) y agrupamiento 310.15(C)(1)' + (P.autoAmpacidad === false ? '' : '; calibre ajustado automáticamente')],
          ['Demanda', 'Tomas: 100 % de los primeros 10 kVA + 50 % del resto (220.44); otros según tipo de carga']])),
      lista.map(seccion));
  }

  /* ---------- exportaciones ---------- */
  const CSS_DOC = 'body{font-family:"Century Gothic",Arial,sans-serif;font-size:9.5pt;color:#000}h1{font-size:16pt;color:#006600;margin:0}h2{font-size:13pt;color:#006600;border-bottom:2px solid #006600;margin-top:18pt}h3{font-size:11pt;margin:12pt 0 4pt}h4{font-size:10pt;margin:8pt 0 3pt}'
    + 'table{border-collapse:collapse;width:100%;margin:2pt 0 6pt}th,td{border:1px solid #888;padding:2pt 4pt;vertical-align:top}th{background:#D9D9D9;text-align:left}td.n{text-align:right}.kv2 th{width:32%;background:#EFEFEF}.chica{font-size:7.5pt}'
    + '.rep-h{display:flex;gap:14pt;align-items:center;border-bottom:3px solid #006600;padding-bottom:6pt}.rep-h img{height:46pt}p{margin:2pt 0}';
  function word(lista) {
    const cuerpo = documento(lista).outerHTML.replace(/src="img\/logo-sinergia\.png"/, 'src="' + new URL('img/logo-sinergia.png', location.href).href + '"');
    const html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>Memoria de cálculo</title>'
      + '<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View></w:WordDocument></xml><![endif]--><style>@page Section1{size:11in 8.5in;mso-page-orientation:landscape;margin:0.6in}div.Section1{page:Section1}' + CSS_DOC + '</style></head><body><div class="Section1">' + cuerpo + '</div></body></html>';
    U.download(App.fileBase() + ' - memoria de cálculo.doc', '﻿' + html, 'application/msword');
  }

  App.views.reporte = function (view, R, id) {
    const lista = id && id !== 'todos' ? R.orden.filter(r => r.tab.id === id) : R.orden;
    view.appendChild(h('div', { class: 'page-h row no-print' }, h('div', null, h('h2', null, 'Reporte de memoria de cálculo'), h('p', { class: 'muted' }, lista.length === 1 ? lista[0].nombre : 'Todos los tableros (' + lista.length + ')')),
      h('div', { class: 'toolbar' },
        UI.btn('Imprimir / PDF', () => window.print(), 'primary small'),
        UI.btn('Descargar Word', () => word(lista), 'small'),
        UI.btn('Descargar Excel', () => ExportExcel.download(R, { memoria: true, ids: lista.map(r => r.tab.id) }), 'small'),
        UI.btn('← Volver', () => history.back(), 'small ghost'))));
    if (!lista.length) { view.appendChild(h('div', { class: 'empty' }, 'Sin tableros.')); return; }
    view.appendChild(h('article', { class: 'hoja doc' }, documento(lista)));
  };
  window.Reporte = { documento, word };
})();
