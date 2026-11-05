/* Tablas de datos del unifilar para AutoCAD / AutoCAD LT, enlazadas por bloques con atributos (como el flujo de TagAcometida con ATTOUT/ATTIN, pero sin Excel).
   Dos bloques, como se arma el diagrama en CAD: TABLA_ETIQUETAS (la columna de rótulos, una sola vez) y TABLA_DATOS (la columna de valores, uno por tablero, con atributos).
   1. Una vez por proyecto se carga la rutina ANACAR-TABLAS.lsp (APPLOAD); ella crea los dos bloques en el dibujo si faltan.
   2. Cada vez que cambia el proyecto: ANACAR exporta el .txt de las tablas (con las casillas de los tableros que se quieren) y en CAD se corre ACTUALIZARTABLAS:
      busca cada bloque TABLA_DATOS por el atributo NOMBRE (o por el nombre corto del tablero), actualiza sus valores y ofrece insertar los tableros nuevos. */
(function () {
  'use strict';
  const h = U.h;
  /* filas base de la tabla (las 18 primeras de Unifilar.filasDatos): rótulo de TABLA_ETIQUETAS y etiqueta (tag) del atributo de TABLA_DATOS; '' = fila sin valor */
  const FILAS = [
    ['EQUIPO', 'NOMBRE'], ['CORRIENTE CORTOCIRCUITO', 'ICC'], ['kVA TOTALES', 'KVA_TOTAL'], ['FACTOR DE DEMANDA', 'FACTOR_DEM'], ['kVA DEMANDADOS', 'KVA_DEM'], ['FACTOR DE POTENCIA', 'FP'],
    ['PROTECCIÓN DEL TABLERO', 'PROTECCION'], ['MODELO / SCCR', 'MODELO'], ['ACOMETIDA', ''], ['FASES', 'FASES'], ['NEUTRO', 'NEUTRO'], ['TIERRA', 'TIERRA'],
    ['AISLAMIENTO / COND.', 'AISLAMIENTO'], ['CANALIZACIÓN', 'CANALIZACION'], ['LONGITUD', 'LONGITUD'], ['VOLTAJE NOMINAL', 'V_NOMINAL'], ['VOLTAJE CALCULADO', 'V_CALCULADO'], ['% CAÍDA DE VOLTAJE', 'CAIDA'],
  ];
  const TAGS = FILAS.map(f => f[1]).filter(Boolean);
  const CAPA = 'C-E-D.UNIFILAR';
  const W1 = 44, W2 = 44, H = 3.6, TH = 2, SEP = 10;   // mm: ancho de la columna de rótulos y de la de datos, alto de fila, alto del texto y separación entre bloques de datos

  /** Texto a bytes de Windows-1252 (lo que lee AutoLISP de AutoCAD LT); lo que no cabe se escribe «?». */
  const ansi = t => { const s = String(t), b = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); b[i] = c < 256 ? c : 63; } return b; };
  const limpio = v => String(v === undefined || v === null ? '' : v).replace(/[\t\r\n]+/g, ' ').replace(/[—–]/g, '-').trim();

  /** Una línea por tablero: los valores en el orden de TAGS (NOMBRE, ICC, …) y al final CLAVE (nombre corto del tablero, por si el atributo NOMBRE del bloque lo trae así). */
  function txt(rs) {
    const lineas = [TAGS.concat('CLAVE').join('\t')];
    rs.forEach(r => { const f = Unifilar.filasDatos(r); lineas.push(FILAS.map((x, i) => (x[1] ? limpio(f[i][1]) : null)).filter(v => v !== null).concat(limpio(r.tab.nombre)).join('\t')); });
    return ansi(lineas.join('\r\n') + '\r\n');
  }

  /** Líneas de AutoLISP que crean los bloques en el dibujo (entmake). Texto con tildes como \U+00D3 para no depender de la página de códigos del archivo. */
  function bloquesLsp() {
    const n = FILAS.length + 1, r = v => (Math.round(v * 1000) / 1000).toFixed(3);
    const lsa = t => String(t).replace(/[^\x20-\x7e]/g, c => '\\U+' + c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0'));
    const c8 = '(cons 8 "' + CAPA + '")';
    const o = [
      '(defun anacar:ln (x1 y1 x2 y2) (entmake (list (quote (0 . "LINE")) ' + c8 + ' (list 10 x1 y1 0.0) (list 11 x2 y2 0.0))))',
      '(defun anacar:tx (x y s) (entmake (list (quote (0 . "TEXT")) ' + c8 + ' (list 10 x y 0.0) (cons 40 ' + TH + '.0) (cons 1 s))))',
      '(defun anacar:ad (x y tg) (entmake (list (quote (0 . "ATTDEF")) (quote (100 . "AcDbEntity")) ' + c8 + ' (quote (100 . "AcDbText")) (list 10 x y 0.0) (cons 40 ' + TH + '.0) (cons 1 "-") (quote (100 . "AcDbAttributeDefinition")) (cons 3 tg) (cons 2 tg) (quote (70 . 0)))))',
      '(defun anacar:capa ()',
      '  (if (null (tblsearch "LAYER" "' + CAPA + '"))',
      '    (entmake (list (quote (0 . "LAYER")) (quote (100 . "AcDbSymbolTableRecord")) (quote (100 . "AcDbLayerTableRecord")) (cons 2 "' + CAPA + '") (quote (70 . 0)) (quote (62 . 7)) (quote (6 . "Continuous"))))))',
      '(defun anacar:marco (w / i)',
      '  (setq i 0)',
      '  (while (<= i ' + n + ')',
      '    (anacar:ln 0.0 (* -' + H + ' i) w (* -' + H + ' i))',
      '    (setq i (1+ i)))',
      '  (anacar:ln 0.0 0.0 0.0 ' + r(-n * H) + ')',
      '  (anacar:ln w 0.0 w ' + r(-n * H) + '))',
      '(defun anacar:bloque-datos ()',
      '  (anacar:capa)',
      '  (entmake (list (quote (0 . "BLOCK")) (quote (2 . "TABLA_DATOS")) (quote (70 . 2)) (list 10 0.0 0.0 0.0)))',
      '  (anacar:marco ' + r(W2) + ')',
    ];
    FILAS.forEach(([, tag], i) => { if (tag) o.push('  (anacar:ad 1.200 ' + r(-(i + 2) * H + 1) + ' "' + tag + '")'); });
    o.push('  (entmake (list (quote (0 . "ENDBLK"))))', '  (princ "\\nBloque TABLA_DATOS creado."))',
      '(defun anacar:bloque-etiquetas ()',
      '  (anacar:capa)',
      '  (entmake (list (quote (0 . "BLOCK")) (quote (2 . "TABLA_ETIQUETAS")) (quote (70 . 0)) (list 10 0.0 0.0 0.0)))',
      '  (anacar:marco ' + r(W1) + ')',
      '  (anacar:tx 1.200 ' + r(-H + 1) + ' "TABLA RESUMEN DEL TABLERO")');
    FILAS.forEach(([etq], i) => o.push('  (anacar:tx 1.200 ' + r(-(i + 2) * H + 1) + ' "' + lsa(etq) + '")'));
    o.push('  (entmake (list (quote (0 . "ENDBLK"))))', '  (princ "\\nBloque TABLA_ETIQUETAS creado."))');
    return o;
  }

  /* ---------- ANACAR-TABLAS.lsp (AutoCAD LT 2024 o más nuevo; solo funciones básicas de AutoLISP, sin vla-/vl-) ---------- */
  const LSP = [
    ';;; ANACAR-TABLAS.lsp - actualiza los bloques TABLA_DATOS con el .txt que exporta ANACAR (Diagrama unifilar > Tablas a CAD).',
    ';;; Cargar una vez con APPLOAD. Comandos: ACTUALIZARTABLAS (actualiza e inserta los tableros nuevos) y TABLAETIQUETAS (inserta la columna de rotulos).',
    ';;; AutoCAD LT 2024 o mas nuevo.',
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
    '(defun anacar:registro (datos nom / r res k)',
    '  (setq res nil k (strcase nom))',
    '  (foreach r datos',
    '    (if (and (null res)',
    '             (or (= (strcase (cdr (assoc "NOMBRE" r))) k)',
    '                 (and (assoc "CLAVE" r) (= (strcase (cdr (assoc "CLAVE" r))) k))))',
    '      (setq res r)))',
    '  res)',
    ...bloquesLsp(),
    '(defun c:TABLAETIQUETAS (/ pt old2)',
    '  (if (null (tblsearch "BLOCK" "TABLA_ETIQUETAS")) (anacar:bloque-etiquetas))',
    '  (setq pt (getpoint "\\nPunto de insercion de la columna de etiquetas (esquina superior izquierda): "))',
    '  (if pt',
    '    (progn',
    '      (setq old2 (getvar "OSMODE"))',
    '      (command "_.-INSERT" "TABLA_ETIQUETAS" pt 1 1 0)',
    '      (setvar "OSMODE" old2)))',
    '  (princ))',
    '(defun c:ACTUALIZARTABLAS (/ ruta datos ss i en nom reg hechos vistos nuevos pt pt0 resp old1 old2 old3 e)',
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
    '                (if (= (strcase (cdr (assoc 2 (entget en)))) "TABLA_DATOS")',
    '                  (progn',
    '                    (setq nom (anacar:nombre en))',
    '                    (if nom',
    '                      (progn',
    '                        (setq reg (anacar:registro datos nom))',
    '                        (if reg',
    '                          (progn (anacar:poner en reg) (setq hechos (1+ hechos) vistos (cons (strcase (cdr (assoc "NOMBRE" reg))) vistos)))))))))))',
    '          (setq nuevos nil)',
    '          (foreach reg datos',
    '            (if (not (member (strcase (cdr (assoc "NOMBRE" reg))) vistos))',
    '              (setq nuevos (cons reg nuevos))))',
    '          (setq nuevos (reverse nuevos))',
    '          (princ (strcat "\\nBloques actualizados: " (itoa hechos) ". Tableros del archivo sin bloque en el dibujo: " (itoa (length nuevos)) "."))',
    '          (if nuevos',
    '            (progn',
    '              (initget "Si No")',
    '              (setq resp (getkword "\\nInsertar los bloques de los tableros nuevos? [Si/No] <Si>: "))',
    '              (if (/= resp "No")',
    '                (progn',
    '                  (setq pt (getpoint "\\nPunto de insercion del primer bloque de datos (esquina superior izquierda): "))',
    '                  (if pt',
    '                    (progn',
    '                      (setq pt0 pt old1 (getvar "ATTREQ") old2 (getvar "OSMODE") old3 (getvar "CMDECHO"))',
    '                      (setvar "ATTREQ" 0) (setvar "OSMODE" 0) (setvar "CMDECHO" 0)',
    '                      (if (null (tblsearch "BLOCK" "TABLA_DATOS")) (anacar:bloque-datos))',
    '                      (foreach reg nuevos',
    '                        (command "_.-INSERT" "TABLA_DATOS" pt 1 1 0)',
    '                        (setq e (entlast))',
    '                        (if (= (cdr (assoc 0 (entget e))) "SEQEND") (setq e (cdr (assoc -2 (entget e)))))',
    '                        (anacar:poner e reg)',
    '                        (setq pt (list (+ (car pt) ' + (W2 + SEP) + '.0) (cadr pt) 0.0)))',
    '                      (if (null (ssget "_X" (quote ((0 . "INSERT") (2 . "TABLA_ETIQUETAS")))))',
    '                        (progn',
    '                          (if (null (tblsearch "BLOCK" "TABLA_ETIQUETAS")) (anacar:bloque-etiquetas))',
    '                          (command "_.-INSERT" "TABLA_ETIQUETAS" (list (- (car pt0) ' + W1 + '.0) (cadr pt0) 0.0) 1 1 0)))',
    '                      (setvar "ATTREQ" old1) (setvar "OSMODE" old2) (setvar "CMDECHO" old3)',
    '                      (princ (strcat "\\nBloques insertados: " (itoa (length nuevos)) "."))))))))))))',
    '  (princ))',
    '(princ "\\nANACAR-TABLAS cargado: use ACTUALIZARTABLAS (y TABLAETIQUETAS para la columna de rotulos).")',
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
      h('p', { class: 'muted' }, 'Elija los tableros cuyas tablas quiere exportar. En CAD, ACTUALIZARTABLAS actualiza los bloques TABLA_DATOS (la columna de valores de cada tablero) por el nombre del tablero y ofrece insertar los que no existan; la columna de rótulos (TABLA_ETIQUETAS) va una sola vez.'),
      h('div', { class: 'toolbar' }, UI.btn('Marcar todos', todos(true), 'small'), UI.btn('Quitar todos', todos(false), 'small')),
      h('div', { style: 'max-height:260px;overflow:auto;margin:6px 0' }, filas),
      h('details', null, h('summary', null, 'Primera vez en un dibujo'),
        h('ol', { class: 'muted' }, h('li', null, 'Bajar y cargar la rutina (APPLOAD): ANACAR-TABLAS.lsp.'), h('li', null, 'Correr ACTUALIZARTABLAS: crea los bloques que falten, inserta los tableros nuevos (con la columna de rótulos a la izquierda del primero) y actualiza los que ya existen. No hace falta insertar ningún archivo.'), h('li', null, 'TABLAETIQUETAS inserta otra columna de rótulos donde la necesite.')),
        h('div', { class: 'toolbar' }, UI.btn('Rutina ANACAR-TABLAS.lsp', () => U.download('ANACAR-TABLAS.lsp', LSP, 'text/plain'), 'small')))), [
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
  window.CadTablas = { dialogo, txt, LSP };
})();
