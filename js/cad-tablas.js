/* Tablas de datos del unifilar para AutoCAD / AutoCAD LT, enlazadas por bloques con atributos (como el flujo de TagAcometida con ATTOUT/ATTIN, pero sin Excel):
   1. Una vez por proyecto: se inserta en el dibujo el bloque TABLA_TABLERO (archivo TABLA_TABLERO.dxf) y se carga la rutina ANACAR-TABLAS.lsp (APPLOAD).
   2. Cada vez que cambia el proyecto: ANACAR exporta el .txt de las tablas (con las casillas de los tableros que se quieren) y en CAD se corre ACTUALIZARTABLAS:
      busca cada bloque por el atributo NOMBRE, actualiza sus valores y ofrece insertar los tableros nuevos. */
(function () {
  'use strict';
  const h = U.h;
  /* filas base de la tabla (las 18 primeras de Unifilar.filasDatos): etiqueta fija del bloque y etiqueta (tag) del atributo con el valor; '' = fila sin valor */
  const FILAS = [
    ['EQUIPO', 'NOMBRE'], ['CORRIENTE CORTOCIRCUITO', 'ICC'], ['kVA TOTALES', 'KVA_TOTAL'], ['FACTOR DE DEMANDA', 'FACTOR_DEM'], ['kVA DEMANDADOS', 'KVA_DEM'], ['FACTOR DE POTENCIA', 'FP'],
    ['PROTECCIÓN DEL TABLERO', 'PROTECCION'], ['MODELO / SCCR', 'MODELO'], ['ACOMETIDA', ''], ['FASES', 'FASES'], ['NEUTRO', 'NEUTRO'], ['TIERRA', 'TIERRA'],
    ['AISLAMIENTO / COND.', 'AISLAMIENTO'], ['CANALIZACIÓN', 'CANALIZACION'], ['LONGITUD', 'LONGITUD'], ['VOLTAJE NOMINAL', 'V_NOMINAL'], ['VOLTAJE CALCULADO', 'V_CALCULADO'], ['% CAÍDA DE VOLTAJE', 'CAIDA'],
  ];
  const TAGS = FILAS.map(f => f[1]).filter(Boolean);
  const CAPA = 'C-E-D.UNIFILAR';

  /** Texto a bytes de Windows-1252 (lo que lee AutoLISP de AutoCAD LT); lo que no cabe se escribe «?». */
  const ansi = t => { const s = String(t), b = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); b[i] = c < 256 ? c : 63; } return b; };
  const limpio = v => String(v === undefined || v === null ? '' : v).replace(/[\t\r\n]+/g, ' ').replace(/[—–]/g, '-').trim();

  /** Una línea por tablero: NOMBRE, ICC, … en el orden de TAGS; los valores salen de la misma tabla que se ve en el diagrama. */
  function txt(rs) {
    const lineas = [TAGS.join('\t')];
    rs.forEach(r => { const f = Unifilar.filasDatos(r); lineas.push(FILAS.map((x, i) => (x[1] ? limpio(f[i][1]) : null)).filter(v => v !== null).join('\t')); });
    return ansi(lineas.join('\r\n') + '\r\n');
  }

  /* ---------- TABLA_TABLERO.dxf: el bloque (al insertar el .dxf toma el nombre del archivo) ---------- */
  const esc = t => String(t).replace(/[^\x20-\x7e]/g, c => '\\U+' + c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0'));
  function bloqueDxf() {
    const W1 = 44, W2 = 36, H = 3.6, TH = 2, n = FILAS.length + 1, out = [];   // mm; origen arriba a la izquierda, hacia abajo
    const g = (c, v) => out.push(String(c), String(v));
    const r = v => Math.round(v * 1000) / 1000;
    const linea = (x1, y1, x2, y2) => { g(0, 'LINE'); g(8, CAPA); g(10, r(x1)); g(20, r(y1)); g(30, 0); g(11, r(x2)); g(21, r(y2)); g(31, 0); };
    const texto = (x, y, t) => { g(0, 'TEXT'); g(8, CAPA); g(10, r(x)); g(20, r(y)); g(30, 0); g(40, TH); g(1, esc(t)); };
    for (let i = 0; i <= n; i++) linea(0, -i * H, W1 + W2, -i * H);
    [0, W1, W1 + W2].forEach(x => linea(x, 0, x, -n * H));
    texto(1.2, -H + 1, 'TABLA RESUMEN DEL TABLERO');
    FILAS.forEach(([etq, tag], i) => {
      const y = -(i + 2) * H + 1;
      texto(1.2, y, etq);
      if (tag) { g(0, 'ATTDEF'); g(8, CAPA); g(10, W1 + 1.2); g(20, r(y)); g(30, 0); g(40, TH); g(1, '-'); g(3, tag); g(2, tag); g(70, 0); }
    });
    const cab = [];
    const c = (a, b) => cab.push(String(a), String(b));
    c(0, 'SECTION'); c(2, 'HEADER'); c(9, '$ACADVER'); c(1, 'AC1009'); c(9, '$INSUNITS'); c(70, 4); c(0, 'ENDSEC');
    c(0, 'SECTION'); c(2, 'TABLES');
    c(0, 'TABLE'); c(2, 'LTYPE'); c(70, 1); c(0, 'LTYPE'); c(2, 'CONTINUOUS'); c(70, 0); c(3, 'Solid line'); c(72, 65); c(73, 0); c(40, 0); c(0, 'ENDTAB');
    c(0, 'TABLE'); c(2, 'LAYER'); c(70, 1); c(0, 'LAYER'); c(2, CAPA); c(70, 0); c(62, 7); c(6, 'CONTINUOUS'); c(0, 'ENDTAB');
    c(0, 'TABLE'); c(2, 'STYLE'); c(70, 1); c(0, 'STYLE'); c(2, 'STANDARD'); c(70, 0); c(40, 0); c(41, 1); c(50, 0); c(71, 0); c(42, 2.5); c(3, 'txt'); c(4, ''); c(0, 'ENDTAB');
    c(0, 'ENDSEC'); c(0, 'SECTION'); c(2, 'ENTITIES');
    const pares = [...cab, ...out, '0', 'ENDSEC', '0', 'EOF'];
    const l = []; for (let i = 0; i < pares.length; i += 2) l.push(String(pares[i]).padStart(3, ' '), String(pares[i + 1]));
    return l.join('\r\n') + '\r\n';
  }

  /* ---------- ANACAR-TABLAS.lsp (AutoCAD LT 2024 o más nuevo; solo funciones básicas de AutoLISP, sin vla-/vl-) ---------- */
  const LSP = [
    ';;; ANACAR-TABLAS.lsp - actualiza los bloques TABLA_TABLERO con el .txt que exporta ANACAR (Diagrama unifilar > Tablas a CAD).',
    ';;; Cargar una vez con APPLOAD. Comando: ACTUALIZARTABLAS (AutoCAD LT 2024 o mas nuevo).',
    '(defun anacar:campos (s / i n r ini)',
    '  (setq r nil ini 0 i 0 n (strlen s))',
    '  (while (< i n)',
    '    (if (= (substr s (1+ i) 1) "\\t")',
    '      (setq r (cons (substr s (1+ ini) (- i ini)) r) ini (1+ i)))',
    '    (setq i (1+ i)))',
    '  (reverse (cons (substr s (1+ ini)) r)))',
    '(defun anacar:pares (cab reg / r)',
    '  (setq r nil)',
    '  (while (and cab reg)',
    '    (setq r (cons (cons (strcase (car cab)) (car reg)) r) cab (cdr cab) reg (cdr reg)))',
    '  (reverse r))',
    '(defun anacar:leer (ruta / f l cab res)',
    '  (setq res nil f (open ruta "r"))',
    '  (if f',
    '    (progn',
    '      (setq l (read-line f))',
    '      (if l (setq cab (anacar:campos l)))',
    '      (while (and cab (setq l (read-line f)))',
    '        (if (> (strlen l) 0) (setq res (cons (anacar:pares cab (anacar:campos l)) res))))',
    '      (close f)))',
    '  (reverse res))',
    '(defun anacar:atributos (en / e d r)',
    '  (setq r nil e (entnext en))',
    '  (while (and e (= (cdr (assoc 0 (setq d (entget e)))) "ATTRIB"))',
    '    (setq r (cons (cons (strcase (cdr (assoc 2 d))) e) r) e (entnext e)))',
    '  r)',
    '(defun anacar:poner (en reg / p d v n)',
    '  (setq n 0)',
    '  (foreach p (anacar:atributos en)',
    '    (if (setq v (assoc (car p) reg))',
    '      (progn',
    '        (setq d (entget (cdr p)))',
    '        (entmod (subst (cons 1 (cdr v)) (assoc 1 d) d))',
    '        (entupd (cdr p))',
    '        (setq n (1+ n)))))',
    '  (entupd en)',
    '  n)',
    '(defun anacar:nombre (en / p)',
    '  (setq p (assoc "NOMBRE" (anacar:atributos en)))',
    '  (if p (cdr (assoc 1 (entget (cdr p)))) nil))',
    '(defun anacar:registro (datos nom / r res)',
    '  (setq res nil)',
    '  (foreach r datos',
    '    (if (and (null res) (= (strcase (cdr (assoc "NOMBRE" r))) (strcase nom))) (setq res r)))',
    '  res)',
    '(defun c:ACTUALIZARTABLAS (/ ruta datos ss i en nom reg hechos vistos nuevos pt resp old1 old2 old3 e)',
    '  (setq ruta (getfiled "Tablas exportadas por ANACAR" "" "txt" 4))',
    '  (if (null ruta)',
    '    (princ "\\nCancelado.")',
    '    (progn',
    '      (setq datos (anacar:leer ruta))',
    '      (if (null datos)',
    '        (princ "\\nEl archivo no tiene tableros.")',
    '        (progn',
    '          (setq hechos 0 vistos nil)',
    '          (setq ss (ssget "_X" (quote ((0 . "INSERT") (66 . 1)))))',
    '          (if ss',
    '            (progn',
    '              (setq i 0)',
    '              (while (< i (sslength ss))',
    '                (setq en (ssname ss i) i (1+ i))',
    '                (if (= (strcase (cdr (assoc 2 (entget en)))) "TABLA_TABLERO")',
    '                  (progn',
    '                    (setq nom (anacar:nombre en))',
    '                    (if nom',
    '                      (progn',
    '                        (setq reg (anacar:registro datos nom))',
    '                        (if reg',
    '                          (progn (anacar:poner en reg) (setq hechos (1+ hechos) vistos (cons (strcase nom) vistos)))))))))))',
    '          (setq nuevos nil)',
    '          (foreach reg datos',
    '            (if (not (member (strcase (cdr (assoc "NOMBRE" reg))) vistos))',
    '              (setq nuevos (cons reg nuevos))))',
    '          (setq nuevos (reverse nuevos))',
    '          (princ (strcat "\\nBloques actualizados: " (itoa hechos) ". Tableros del archivo sin bloque en el dibujo: " (itoa (length nuevos)) "."))',
    '          (cond',
    '            ((and nuevos (null (tblsearch "BLOCK" "TABLA_TABLERO")))',
    '              (princ "\\nFalta el bloque TABLA_TABLERO: inserte TABLA_TABLERO.dxf una vez y repita el comando."))',
    '            (nuevos',
    '              (initget "Si No")',
    '              (setq resp (getkword "\\nInsertar los bloques de los tableros nuevos? [Si/No] <Si>: "))',
    '              (if (/= resp "No")',
    '                (progn',
    '                  (setq pt (getpoint "\\nPunto de insercion del primer bloque (esquina superior izquierda): "))',
    '                  (if pt',
    '                    (progn',
    '                      (setq old1 (getvar "ATTREQ") old2 (getvar "OSMODE") old3 (getvar "CMDECHO"))',
    '                      (setvar "ATTREQ" 0) (setvar "OSMODE" 0) (setvar "CMDECHO" 0)',
    '                      (foreach reg nuevos',
    '                        (command "_.-INSERT" "TABLA_TABLERO" pt 1 1 0)',
    '                        (setq e (entlast))',
    '                        (if (= (cdr (assoc 0 (entget e))) "SEQEND") (setq e (cdr (assoc -2 (entget e)))))',
    '                        (anacar:poner e reg)',
    '                        (setq pt (list (+ (car pt) 90.0) (cadr pt) 0.0)))',
    '                      (setvar "ATTREQ" old1) (setvar "OSMODE" old2) (setvar "CMDECHO" old3)',
    '                      (princ (strcat "\\nBloques insertados: " (itoa (length nuevos)) "."))))))))))))',
    '  (princ))',
    '(princ "\\nANACAR-TABLAS cargado: use el comando ACTUALIZARTABLAS.")',
    '(princ)',
  ].join('\r\n') + '\r\n';

  function dialogo(R) {
    const P = Store.project, ya = new Set(P.cadExportados || []), marcas = {};
    const filas = R.orden.map(r => {
      marcas[r.tab.id] = true;
      const cb = h('input', { type: 'checkbox', checked: true, onchange: e => { marcas[r.tab.id] = e.target.checked; } });
      return h('label', { class: 'chk', style: 'display:block;margin:2px 0' }, cb, ' ' + r.nombre, ya.has(r.tab.id) ? '' : h('span', { class: 'muted' }, ' · nuevo desde la última exportación'));
    });
    const todos = v => () => { filas.forEach(l => { l.querySelector('input').checked = v; }); R.orden.forEach(r => { marcas[r.tab.id] = v; }); };
    UI.modal('Tablas del unifilar a AutoCAD', h('div', null,
      h('p', { class: 'muted' }, 'Elija los tableros cuyas tablas quiere exportar. En CAD, el comando ACTUALIZARTABLAS actualiza los bloques TABLA_TABLERO por el nombre del tablero y ofrece insertar los que no existan.'),
      h('div', { class: 'toolbar' }, UI.btn('Marcar todos', todos(true), 'small'), UI.btn('Quitar todos', todos(false), 'small')),
      h('div', { style: 'max-height:260px;overflow:auto;margin:6px 0' }, filas),
      h('details', null, h('summary', null, 'Primera vez en un dibujo'),
        h('ol', { class: 'muted' }, h('li', null, 'Bajar y cargar la rutina (APPLOAD): ANACAR-TABLAS.lsp.'), h('li', null, 'Insertar una vez el bloque TABLA_TABLERO.dxf (comando INSERT › Examinar). Copiarlo para cada tablero o dejar que ACTUALIZARTABLAS lo inserte.')),
        h('div', { class: 'toolbar' },
          UI.btn('Bloque TABLA_TABLERO.dxf', () => U.download('TABLA_TABLERO.dxf', bloqueDxf(), 'application/dxf'), 'small'),
          UI.btn('Rutina ANACAR-TABLAS.lsp', () => U.download('ANACAR-TABLAS.lsp', LSP, 'text/plain'), 'small')))), [
      { label: 'Cerrar' },
      { label: 'Exportar tablas (.txt)', cls: 'primary', onclick: () => {
        const sel = R.orden.filter(r => marcas[r.tab.id]);
        if (!sel.length) { UI.alert('Marque al menos un tablero.'); return false; }
        U.download(App.fileBase() + ' - tablas CAD.txt', txt(sel), 'text/plain');
        P.cadExportados = Array.from(new Set((P.cadExportados || []).concat(sel.map(r => r.tab.id)))); Store.save();
        UI.toast(sel.length + ' tabla(s) exportadas: en AutoCAD use ACTUALIZARTABLAS', 'ok');
      } },
    ], { wide: true });
  }
  window.CadTablas = { dialogo, txt, bloqueDxf, LSP };
})();
