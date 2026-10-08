/* Prevista para diseño esquemático: carga estimada por tipo de uso y metros cuadrados (iluminación NEC 2020 tabla 220.12,
   tomas y climatización con densidades editables) y generación de circuitos de prevista en el tablero elegido. */
(function () {
  'use strict';
  const h = U.h, f2 = v => U.fmt(v, 2), f1 = v => U.fmt(v, 1);
  const n = v => Number(v) || 0;
  const MAX_ILUM = 1.5, MAX_TOMA = 1.44, MAX_CLIMA = 10;   // kVA por circuito al generar (20 A al 80 % en 120 V; equipos ≤ 10 kVA)

  const uso = id => (Store.catalog.usos || []).find(u => String(u.id) === String(id)) || {};
  const val = (p, k) => (p[k] !== undefined && p[k] !== '' && p[k] !== null ? n(p[k]) : n(uso(p.usoId)[k]));
  function calcular(p) {
    const a = n(p.area);
    const ilum = a * val(p, 'ilum') / 1000, tomas = a * val(p, 'tomas') / 1000, clima = a * val(p, 'clima') / 1000, otros = n(p.otros);
    return { ilum, tomas, clima, otros, total: ilum + tomas + clima + otros };
  }

  /** Reemplaza los circuitos de prevista de esta fila en su tablero. */
  function generar(p) {
    const t = Store.tablero(p.tableroId); if (!t) { UI.alert('Elija el tablero donde se generan los circuitos.'); return; }
    const c = calcular(p), u = uso(p.usoId), nom = (u.nombre || 'Uso') + (p.nombre ? ' ' + p.nombre : '') + ' (' + f1(n(p.area)) + ' m²)';
    t.circuitos = t.circuitos.filter(x => x.previstaId !== p.id);
    const V = n(t.voltaje), det = re => (Store.catalog.detallesCarga.find(d => re.test(d.descripcion) && (V >= 440 ? n(d.v) === 277 || n(d.v) === 480 : n(d.v) <= 240)) || {}).id;
    const dIlum = V >= 440 ? 6 : 1, dToma = 8, dEq = det(/^EQUIPOS MEC/i) || 55;
    const crear = (kva, max, detalle, txt) => {
      if (kva <= 0) return 0;
      const k = Math.max(1, Math.ceil(kva / max)), dd = Store.catalog.detallesCarga.find(x => String(x.id) === String(detalle)) || {};
      const np = n(dd.fases) === 3 ? 3 : n(dd.fases) === 2 ? 2 : 1;
      for (let i = 0; i < k; i++) t.circuitos.push({ id: U.uid(), polos: Store.posicionLibre(t, np), detalleId: detalle, descripcion: 'Prevista ' + txt + ' ' + nom + (k > 1 ? ' ' + (i + 1) + '/' + k : ''), kva: Math.round(kva / k * 1000) / 1000,
        longitud: p.longitud || 25, material: 'CU', aislamiento: 'THHN', mult: '', paralelos: '', aumento: 1, breakerId: '', prot: '', tableroHijoId: '', origen: 'prevista', previstaId: p.id });
      return k;
    };
    const total = crear(c.ilum, MAX_ILUM, dIlum, 'iluminación') + crear(c.tomas, MAX_TOMA, dToma, 'tomas') + crear(c.clima, MAX_CLIMA, dEq, 'climatización') + crear(c.otros, MAX_CLIMA, dEq, 'equipos');
    Store.save(); App.refresh(); UI.toast(total + ' circuitos de prevista en ' + Calc.nombreTablero(t), 'ok');
  }

  App.views.prevista = function (view, R) {
    const P = Store.project, L = P.previstas = P.previstas || [];
    const usos = (Store.catalog.usos || []).map(u => ({ value: u.id, label: u.nombre }));
    const tabs = UI.opts(P.tableros.map(t => ({ value: t.id, label: Calc.nombreTablero(t) })), '— tablero —');
    view.appendChild(h('div', { class: 'page-h row' }, h('div', null, h('h2', null, 'Prevista (diseño esquemático)'), h('p', { class: 'muted' }, 'Carga estimada por tipo de uso y área. Iluminación según NEC 2020 tabla 220.12; tomas y climatización con densidades estimadas (editables por fila o en Administración ▸ Tipos de uso).')),
      h('div', { class: 'toolbar' }, UI.btn('+ Área', () => { L.push({ id: U.uid(), usoId: 1, area: '', nombre: '', tableroId: P.tableros[0] ? P.tableros[0].id : '', longitud: 25 }); Store.save(); App.refresh(); }, 'primary small'))));
    if (!L.length) { view.appendChild(h('div', { class: 'empty' }, 'Agregue las áreas del proyecto (tipo de uso y m²) para estimar la carga y generar circuitos de prevista.')); return; }
    const tot = { ilum: 0, tomas: 0, clima: 0, otros: 0, total: 0 };
    const body = h('tbody', null, L.map(p => {
      const c = calcular(p), k = 'pv:' + p.id + ':', u = uso(p.usoId);
      Object.keys(tot).forEach(x => { tot[x] += c[x]; });
      const nCirc = P.tableros.reduce((a, t) => a + t.circuitos.filter(x => x.previstaId === p.id).length, 0);
      return h('tr', null,
        h('td', null, UI.input(p, 'nombre', { fk: k + 'nom', placeholder: 'Nivel 1, ala norte…', class: 'w-desc' })),
        h('td', null, UI.bind(p, 'usoId', usos, { num: true, fk: k + 'uso' })),
        h('td', null, UI.input(p, 'area', { type: 'num', fk: k + 'area', class: 'w-num' })),
        h('td', null, UI.input(p, 'ilum', { type: 'num', fk: k + 'il', class: 'w-xs', placeholder: f1(u.ilum) })),
        h('td', null, UI.input(p, 'tomas', { type: 'num', fk: k + 'to', class: 'w-xs', placeholder: f1(u.tomas) })),
        h('td', null, UI.input(p, 'clima', { type: 'num', fk: k + 'cl', class: 'w-xs', placeholder: f1(u.clima) })),
        h('td', null, UI.input(p, 'otros', { type: 'num', fk: k + 'ot', class: 'w-xs', placeholder: '0' })),
        h('td', { class: 'r' }, f2(c.ilum)), h('td', { class: 'r' }, f2(c.tomas)), h('td', { class: 'r' }, f2(c.clima + c.otros)), h('td', { class: 'r' }, h('b', null, f2(c.total))),
        h('td', null, UI.bind(p, 'tableroId', tabs, { fk: k + 'tab' })),
        h('td', { class: 'acc' }, UI.btn(nCirc ? 'Regenerar (' + nCirc + ')' : 'Generar circuitos', () => generar(p), 'small'),
          UI.iconBtn('🗑', 'Quitar área y sus circuitos', () => { P.tableros.forEach(t => { t.circuitos = t.circuitos.filter(x => x.previstaId !== p.id); }); L.splice(L.indexOf(p), 1); Store.save(); App.refresh(); }, 'danger')));
    }));
    const sis = P.tableros[0] ? P.tableros[0] : { voltaje: 208, fases: 3 };
    const I = tot.total * 1000 / ((n(sis.fases) === 3 ? 1.732 : 1) * (n(sis.voltaje) || 208));
    view.appendChild(UI.card('Áreas', h('div', { class: 'tbl-wrap' }, h('table', { class: 'tbl' },
      h('thead', null, h('tr', null, ['Área', 'Tipo de uso', 'm²', 'Ilum. VA/m²', 'Tomas VA/m²', 'Clima VA/m²', 'Otros kVA', 'kVA ilum.', 'kVA tomas', 'kVA clima + otros', 'kVA total', 'Tablero', ''].map(x => h('th', null, x)))),
      body, h('tfoot', null, h('tr', null, h('th', { colspan: 7, class: 'r' }, 'Totales'), h('th', { class: 'r' }, f2(tot.ilum)), h('th', { class: 'r' }, f2(tot.tomas)), h('th', { class: 'r' }, f2(tot.clima + tot.otros)), h('th', { class: 'r' }, f2(tot.total)), h('th', { colspan: 2 })))))));
    view.appendChild(h('div', { class: 'kpis' }, UI.kpi('Carga conectada estimada', f2(tot.total) + ' kVA'), UI.kpi('Corriente a ' + (sis.voltaje || 208) + ' V', U.fmt(I, 0) + ' A', 'sin factores de demanda'),
      UI.kpi('Área total', f1(L.reduce((a, p) => a + n(p.area), 0)) + ' m²')));
    view.appendChild(h('p', { class: 'hint' }, 'Al generar, cada fila crea circuitos de prevista en su tablero: iluminación en circuitos de hasta ' + MAX_ILUM + ' kVA, tomas hasta ' + MAX_TOMA + ' kVA y climatización/equipos hasta ' + MAX_CLIMA + ' kVA (3 polos). Los factores de demanda se aplican luego en la memoria de cálculo como a cualquier circuito. Regenerar reemplaza los circuitos de esa fila.'));
  };
})();
