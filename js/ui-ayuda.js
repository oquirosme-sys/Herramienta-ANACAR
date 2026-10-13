/* Módulo Ayuda: guía de llenado, glosario básico de equipos eléctricos (con ilustración o foto) y catálogo
   de tableros y breakers por familia del fabricante. El administrador puede cargar una foto real para cada término
   o familia (se guarda en el catálogo). */
(function () {
  'use strict';
  const h = U.h;

  const GLOSARIO = [
    { k: 'tablero', t: 'Tablero (panelboard)', il: 'tablero', d: 'Gabinete con barras de fase, neutro y tierra donde se conectan los interruptores de los circuitos ramales (NEC 408). Se especifica por voltaje, fases e hilos, amperaje de barras, número de espacios (polos) y capacidad de cortocircuito (SCCR). Ejemplos: Eaton PRL1a/PRL2a/PRL3a, Square D NQ y NF.' },
    { k: 'centroCarga', t: 'Centro de carga', il: 'centroCarga', d: 'Tablero residencial o comercial liviano (240 V máx., barras hasta 225 A) con breakers enchufables de 1 y 2 polos. Ejemplos: Eaton CH, Square D QO.' },
    { k: 'subestacion', t: 'Subestación / switchboard', il: 'subestacion', d: 'Equipo de piso de una o varias secciones que recibe la acometida o el secundario del transformador y distribuye a los tableros con interruptores de caja moldeada o de potencia (600 V, 800–4000 A). Ejemplos: Square D QED-2 e I-Line, Eaton Pow-R-Line 4 y Pow-R-Line Xpert. Se marca en la memoria como "Subestación".' },
    { k: 'principal', t: 'Interruptor principal o zapatas', il: 'principal', d: 'Con interruptor principal el tablero tiene su propia protección general. Con zapatas (main lugs) solo tiene terminales y lo protege el interruptor del tablero que lo alimenta. Cambia el modelo de catálogo del tablero.' },
    { k: 'enchufable', t: 'Breaker enchufable o atornillable (ramal)', il: 'enchufable', d: 'Interruptor termomagnético de 1, 2 o 3 polos para circuitos ramales, de 15 a 125 A. Cada polo ocupa un espacio; los de 2 y 3 polos tienen la palanca unida. La familia depende del tablero: BAB y QBHW (PRL1a), GHB (PRL2a), QOB (NQ), EDB/EGB (NF).' },
    { k: 'cajaMoldeada', t: 'Interruptor de caja moldeada (MCCB)', il: 'cajaMoldeada', d: 'Interruptor de mayor capacidad (hasta 2500 A) para alimentadores y principales. El "marco" es el tamaño físico (p. ej. 250 A) y el "amperaje" el ajuste de disparo (p. ej. 175 A). Ejemplos: Eaton Power Defense PDG2–PDG6, Square D H, J, L, M y P.' },
    { k: 'unidad', t: 'Unidad de disparo: TM, LI, LSI, LSIG · 100 % rated', il: 'cajaMoldeada', o: { electronico: true, etq: 'Unidad electrónica LSIG' }, d: 'TM: termomagnética. LI, LSI y LSIG: electrónicas con ajustes de largo tiempo (L), corto tiempo (S), instantáneo (I) y falla a tierra (G). Desde 1200 A se requiere reducción de energía de arco (NEC 240.87). "100 % rated" permite cargar el interruptor al 100 % de su valor en carga continua (el alimentador se dimensiona sin el 125 %).' },
    { k: 'gfci', t: 'GFCI y AFCI', il: 'gfci', d: 'GFCI: protección contra falla a tierra para personas (NEC 210.8: baños, cocinas, exteriores, lavanderías, piscinas). AFCI: protección contra falla de arco (NEC 210.12: dormitorios y áreas habitables de vivienda, habitaciones de hotel). La herramienta los propone según la descripción del circuito y la ocupación del proyecto.' },
    { k: 'sccr', t: 'SCCR, capacidad interruptiva y clasificación en serie', il: 'cascada', d: 'Cada interruptor y tablero debe soportar la corriente de cortocircuito disponible en sus bornes (Icc). Si un breaker ramal tiene menos capacidad, se puede usar una combinación en serie listada por el fabricante con el principal (NEC 240.86), rotulando el tablero (110.22(C)).' },
    { k: 'spd', t: 'Supresor de picos (SPD)', il: 'spd', d: 'Protege contra sobretensiones transitorias (NEC 242). Tipo 1 o 2 en la acometida (recomendado 250 kA por fase), tipo 2 en tableros (120 kA en tableros de 600–1200 A, 50 kA hasta 400 A). Se elige por sistema (120/208, 277/480…) y kA.' },
    { k: 'transformador', t: 'Transformador', il: 'transformador', d: 'Cambia el voltaje (p. ej. 480 V a 208Y/120 V). Se especifica por kVA, impedancia Z % y relación X/R: con ellos se calcula la caída de voltaje en el transformador y la corriente de cortocircuito del secundario.' },
    { k: 'transferencia', t: 'Transferencia ATS / MTS / interruptor enclavado', il: 'transferencia', d: 'Permite alimentar una carga desde dos fuentes (red y generador, o UPS y bypass). ATS: automática; MTS: manual; IP: interruptores con enclavamiento. Su SCCR debe ser mayor que la corriente de falla (NEC 700.5(E)).' },
    { k: 'ups', t: 'UPS y bypass', il: 'ups', d: 'Sistema de alimentación ininterrumpida. En falla el inversor entrega una corriente limitada (≈ 2 × In), por lo que el cortocircuito aguas abajo es bajo. El bypass alimenta la carga directo desde otro tablero durante mantenimiento.' },
    { k: 'generador', t: 'Generador', il: 'generador', d: "Fuente de respaldo. Su corriente de cortocircuito depende de la reactancia subtransitoria X''d (típica 12 %). Se conecta a la carga por la segunda acometida con un ATS o MTS." },
    { k: 'conductores', t: 'Conductores y calibres', il: 'conductores', d: 'Cobre (CU) o aluminio (AL); aislamiento THHN, XHHW-2 (90 °C) o RHW (75 °C). El calibre (AWG o kcmil) se elige por la protección y se valida con la temperatura ambiente y el agrupamiento (NEC 310.15/310.16); en cargas grandes se usan conductores en paralelo (2x, 3x…). Tierra según 250.122.' },
    { k: 'tuberia', t: 'Canalizaciones', il: 'tuberia', d: 'EMT (tubo metálico) o PVC SCH40; el diámetro sale del número y calibre de conductores (Anexo C). La ductobarra reemplaza cables en alimentadores grandes.' },
    { k: 'barras', t: 'Barras y espacios', il: 'barras', d: 'Amperaje de las barras de fase, neutro y tierra del tablero (debe ser ≥ la protección principal) y número de espacios (polos) disponibles. La herramienta deja la reserva de espacios indicada en Proyecto.' },
    { k: 'factores', t: 'Factores de demanda, uso y diversidad', il: 'factores', d: 'Factor de demanda o de uso: fracción de la carga conectada que se usa a la vez (tomas: 100 % de los primeros 10 kVA y 50 % del resto, NEC 220.44). Factor de diversidad: entre 0 y 1, multiplica la carga para considerar que no todas las cargas coinciden. Reserva: fracción de carga futura (0,1 = 10 %).' },
    { k: 'cascada', t: 'Caída de voltaje y cortocircuito en cascada', il: 'cascada', d: 'Cada tablero parte del voltaje en bornes del tablero que lo alimenta, por lo que la caída es acumulada desde la acometida (máx. 3 % alimentador, 5 % total). La corriente de cortocircuito disminuye en cada alimentador (método punto a punto).' },
  ];

  const GUIA = [
    ['Proyecto', 'Datos del proyecto y criterios (caídas máximas, desbalance, marca, reserva de espacios, ocupación NEC, Icc de la red). Cree los tableros (uno, varios o importándolos de Revit) e indique de cuál se alimenta cada uno.'],
    ['Memoria de cálculo', 'Por cada tablero: datos (tipo, sistema, principal o zapatas, subestación), alimentador (material, aislamiento, unidad de disparo, 100 % rated) y circuitos. Las celdas de color verde son para llenar; lo que aparece en gris dentro de ellas es el valor automático, que se usa mientras la celda quede vacía.'],
    ['Circuitos', 'Posición (polos separados por coma: 1,3,5), detalle de carga, kVA y longitud. Opcional: factor de protección, factor de uso, factor de diversidad (0 a 1), protección, conductor, unidad (GFCI/AFCI) y breaker. Primero van las posiciones impares y luego las pares, como en el tablero físico.'],
    ['Fuentes', 'Transformador aguas arriba (kVA, Z %), UPS, y segunda acometida con ATS/MTS/IP desde un generador o en bypass desde otro tablero.'],
    ['Autobalanceo y Revit', 'El botón Autobalancear reubica circuitos para igualar las fases; los cambios de posición quedan en Revit ▸ Cambios para Revit. En la pestaña Revit se importa la tabla de circuitos o se actualizan cargas y distancias.'],
    ['Tableros 3F / 1F y Tablas resumen', 'Cuadros de carga y tablas para los planos con el diseño del Excel; se imprimen o se exportan a Excel con los nombres definidos para vincular en Revit.'],
    ['Prevista', 'Para diseño esquemático: área y tipo de uso de cada zona y la herramienta genera los circuitos de prevista en el tablero elegido.'],
  ];

  /** Reduce una foto a 640 px de ancho máximo (JPEG) para guardarla en el catálogo. */
  function foto(file) {
    return new Promise((ok, mal) => {
      const fr = new FileReader();
      fr.onload = () => { const im = new Image(); im.onload = () => {
        const k = Math.min(1, 640 / im.width), cv = document.createElement('canvas'); cv.width = Math.round(im.width * k); cv.height = Math.round(im.height * k);
        cv.getContext('2d').drawImage(im, 0, 0, cv.width, cv.height); ok(cv.toDataURL('image/jpeg', 0.82));
      }; im.onerror = mal; im.src = fr.result; };
      fr.onerror = mal; fr.readAsDataURL(file);
    });
  }
  function imagen(clave, il, o) {
    const fotos = Store.catalog.fotos || {}, box = h('div', { class: 'ilus' });
    if (fotos[clave]) box.appendChild(h('img', { src: fotos[clave], alt: clave }));
    else box.innerHTML = Ilus.get(il, o);           // SVG propio (texto fijo de ilustraciones.js)
    if (Auth.isAdmin()) box.appendChild(h('div', { class: 'ilus-acc' },
      UI.btn(fotos[clave] ? 'Cambiar foto' : 'Subir foto', async () => {
        const f = await UI.pickFile('image/*'); if (!f) return;
        try { Store.catalog.fotos = Object.assign({}, Store.catalog.fotos, { [clave]: await foto(f) }); Store.save(); App.refresh(); } catch (e) { UI.alert('No se pudo leer la imagen.'); }
      }, 'small'),
      fotos[clave] ? UI.btn('Quitar', () => { delete Store.catalog.fotos[clave]; Store.save(); App.refresh(); }, 'small') : null));
    return box;
  }

  const ilTablero = f => (!f ? 'tablero' : f.clase === 'subestacion' ? 'subestacion' : /^(CH|NQ)$/.test(f.familia) ? 'centroCarga' : 'tablero');
  const ilBreaker = f => (!f ? 'enchufable' : /GFCI/.test(f.familia) ? 'gfci' : /^(BAB|QBHW|GHB|QOB|QOB-VH)$/.test(f.familia) ? 'enchufable' : 'cajaMoldeada');

  function catalogoTableros(filtro) {
    const C = Store.catalog, grupos = new Map();
    C.tablerosCat.forEach(x => { const f = Calc.famTablero(x), k = f ? f.marca + ' · ' + f.familia : x.fabricante + ' · otros'; if (!grupos.has(k)) grupos.set(k, { f, items: [] }); grupos.get(k).items.push(x); });
    return Array.from(grupos.entries()).filter(([k, g]) => !filtro || (k + ' ' + g.items.map(x => x.modelo).join(' ')).toLowerCase().includes(filtro)).map(([k, g]) => h('article', { class: 'glo' },
      h('div', { class: 'glo-txt' }, h('h4', null, k),
        g.f ? h('p', null, h('b', null, (g.f.clase === 'subestacion' ? 'Subestación / switchboard' : g.f.clase === 'ambos' ? 'Tablero o subestación' : 'Tablero') + ' · ' + g.f.vMax + ' V máx. · ' + (g.f.fases || '3F/1F') + '. '), g.f.nota || '',
          g.f.ramales ? h('span', null, ' Ramales: ' + g.f.ramales + '.') : null, g.f.principales ? h('span', null, ' Principales: ' + g.f.principales + '.') : null) : null,
        h('div', { class: 'tbl-wrap' }, h('table', { class: 'tbl mini' }, h('thead', null, h('tr', null, ['Tipo', 'Modelo', 'Barras F/N/T (A)', 'Espacios', 'Fases', 'Principal'].map(x => h('th', null, x)))),
          h('tbody', null, g.items.map(x => h('tr', null, h('td', null, x.id), h('td', null, x.modelo), h('td', null, x.barraFase + ' / ' + x.barraNeutro + ' / ' + x.barraTierra), h('td', { class: 'r' }, x.espacios >= 999 ? 'secciones' : x.espacios), h('td', null, x.fases || ''), h('td', null, x.principal || ''))))))),
      imagen('fam:' + k, ilTablero(g.f), { etq: k })));
  }
  function catalogoBreakers(filtro) {
    const C = Store.catalog, grupos = new Map();
    C.breakers.forEach(x => { const f = Calc.famBreaker(x), k = f ? f.marca + ' · ' + f.familia : x.marca + ' · otros'; if (!grupos.has(k)) grupos.set(k, { f, items: [] }); grupos.get(k).items.push(x); });
    return Array.from(grupos.entries()).filter(([k, g]) => !filtro || (k + ' ' + g.items.map(x => x.modelo).join(' ')).toLowerCase().includes(filtro)).map(([k, g]) => h('article', { class: 'glo' },
      h('div', { class: 'glo-txt' }, h('h4', null, k),
        g.f ? h('p', null, h('b', null, g.f.vMax + ' V máx. '), g.f.plantilla ? 'Número de catálogo: ' + g.f.plantilla.replace('{p}', '‹polos›').replace('{a3}', '‹amp. 3 dígitos›').replace('{a}', '‹amp.›').replace('{pdg}', 'PDG + marco + polos + capacidad + amperios + disparo') : '') : null,
        h('div', { class: 'tbl-wrap' }, h('table', { class: 'tbl mini' }, h('thead', null, h('tr', null, ['Tipo', 'Modelo', 'Amperios', 'Polos', 'Unidad', 'Marco', 'SCCR'].map(x => h('th', null, x)))),
          h('tbody', null, g.items.map(x => h('tr', null, h('td', null, x.id), h('td', null, x.modelo || ''), h('td', null, x.amperios || ''), h('td', { class: 'r' }, x.polos || ''), h('td', null, x.unidad || 'STD'), h('td', { class: 'r' }, x.marco || ''), h('td', null, (x.sccr || '') + ' kA @ ' + (x.vSccr || '') + ' V'))))))),
      imagen('fam:' + k, ilBreaker(g.f), { etq: k })));
  }

  App.views.ayuda = function (view, R, id) {
    const st = App.ayuda = App.ayuda || { sec: 'guia', q: '' };
    if (id && st.ruta !== id) { st.sec = id; st.ruta = id; }   // #ayuda/glosario, #ayuda/tableros…
    const sec = (id, txt) => h('button', { type: 'button', class: 'chip' + (st.sec === id ? ' active' : ''), onclick: () => { st.sec = id; App.refresh(); } }, txt);
    const q = h('input', { type: 'search', placeholder: 'Buscar familia o modelo…', value: st.q, 'data-fk': 'ayuda:q', oninput: e => { st.q = e.target.value; clearTimeout(st.tm); st.tm = setTimeout(() => App.refresh(), 250); } });
    view.appendChild(h('div', { class: 'page-h' }, h('h2', null, 'Ayuda'), h('p', { class: 'muted' }, 'Guía de llenado, glosario básico de equipos eléctricos y catálogo de tableros y breakers por familia.' + (Auth.isAdmin() ? ' Como administrador puede subir una foto real para cada término o familia.' : ''))));
    view.appendChild(h('nav', { class: 'chips' }, sec('guia', 'Guía de llenado'), sec('glosario', 'Glosario básico'), sec('tableros', 'Catálogo de tableros'), sec('breakers', 'Catálogo de breakers')));
    if (st.sec === 'guia') {
      view.appendChild(UI.card('Guía de llenado', h('div', null,
        h('div', { class: 'leyenda' }, h('span', { class: 'muestra' }, h('input', { type: 'text', value: 'Valor que usted escribe', readOnly: true })), h('span', { class: 'muestra' }, h('input', { type: 'text', placeholder: 'Automático (gris)', readOnly: true })),
          h('span', { class: 'muted' }, 'Las celdas verdes se llenan; si se dejan vacías se usa el valor automático que se ve en gris.')),
        h('ol', { class: 'guia' }, GUIA.map(([a, b]) => h('li', null, h('b', null, a + ': '), b))))));
    } else if (st.sec === 'glosario') {
      view.appendChild(h('div', { class: 'glosario' }, GLOSARIO.map(x => h('article', { class: 'glo' },
        h('div', { class: 'glo-txt' }, h('h4', null, x.t), h('p', null, x.d)), imagen('glo:' + x.k, x.il, x.o)))));
    } else {
      view.appendChild(h('div', { class: 'toolbar' }, q));
      const f = st.q.trim().toLowerCase();
      view.appendChild(h('div', { class: 'glosario' }, st.sec === 'tableros' ? catalogoTableros(f) : catalogoBreakers(f)));
    }
  };
  App.ayudaDatos = { GLOSARIO, GUIA };
})();
