# Prompt para Claude Code — aplicar y verificar los cambios web.11 a web.14

Copie TODO lo que está debajo de la línea y péguelo en Claude Code, abierto en la carpeta del proyecto que se va a actualizar.

---

Eres el responsable de dejar esta herramienta web igual que la versión de referencia 2026.1-web.14 de ANACAR. Trabaja en español. Haz el trabajo completo sin pedirme confirmación en cada paso: inspecciona, aplica, prueba, corrige y verifica. Solo detente y pregunta si algo choca con una decisión de diseño propia de este proyecto que no puedas conciliar.

## 0. Cómo trabajar

1. **Inspecciona antes de tocar nada.** Lee `README.md`, `js/config.js` (versión), `index.html` y los archivos `js/*.js`. Anota qué puntos de la lista de abajo ya existen (total o parcialmente) en este proyecto y cuáles no. No dupliques lo que ya está; completa lo que falte y corrige lo que difiera.
2. **Fuente de verdad.** Si tienes acceso al repositorio de referencia `oquirosme-sys/Herramienta-ANACAR` (commits `71a8e42` a `cf18eb1`), compara con `git diff` o `git show` y adapta los cambios al código de este proyecto. Si no tienes acceso, implementa cada punto según su descripción y su criterio de aceptación.
3. **Reglas del proyecto.** HTML y JavaScript plano, sin compilación. Conserva los saltos de línea que ya tenga cada archivo (CRLF en la referencia): al editar por script lee y escribe con `newline=''`. No edites `js/data/seed.js` a mano. Al terminar cada punto sube la versión (`js/config.js` APP_VERSION) y cambia la etiqueta `?v=` de TODOS los scripts y estilos de `index.html` (si no, el navegador usa archivos viejos). Los proyectos y catálogos ya guardados en `localStorage` deben seguir abriendo: agrega migraciones en `Store.migrate` cuando agregues campos nuevos al catálogo.
4. **Pruebas.** Levanta un servidor local (`python -m http.server 8766`) y abre `tests/calc-test.html` (debe decir 269 de 269) y `tests/ui-test.html` (10 de 10; si no existe, créalo según el punto 1). Si algo falla, no sigas: arréglalo. Revisa también la consola del navegador: no debe haber errores nuevos.
5. **Commits.** Un commit por punto, con rutas específicas (no `git add -A`). No subas libros de prueba ni archivos que no sean del cambio. No cambies la visibilidad del repositorio ni subas nada a un remoto sin que yo lo pida.

## 1. Lista de cambios

### P1. Casillas numéricas que no guardan (web.12)
En `js/ui.js`, en el manejador de cambio de `UI.input`: si el campo es numérico (`opts.type === 'num'`) → `parseNum`; si es NaN, marca `invalid`, avisa «Número no válido» y sal; si hay `opts.rango` y el valor está fuera (≤ mínimo o > máximo), avisa «Debe estar entre X y Y», restablece el valor anterior y sal. Solo si NO es numérico y `opts.trim !== false` se hace `v = v.trim()`. Ese `trim` va en el `else` del bloque numérico, nunca después (el error típico es `v.trim is not a function` en consola).
Crea o completa `tests/ui-test.html` (iframe que simula escribir con coma decimal) con 10 comprobaciones: kVA del generador se guarda; longitud generador→ATS con coma; distancia normal→ATS; distancia ATS→tablero; longitud del alimentador; el kVA del generador no se borra al cambiar otro dato; kVA de un circuito; cambiar la fuente y volver conserva los datos; diversidad fuera de 0–1 se rechaza; 0,8 se guarda.
**Aceptación:** `ui-test.html` da 10 de 10; escribir 100 en «Generador (kVA)» y cambiar otra casilla no lo borra.

### P2. ATS/MTS con las dos acometidas y sus distancias — motor y memoria (web.11)
- `js/calc.js`: `largoAlim(t)` = `distNormal + distAts` cuando el ATS/MTS está activo y `distNormal > 0`; si no, la longitud normal. Úsala en la longitud del alimentador (M139) y en la longitud del circuito del padre. El objeto `alterna` del resultado incluye: `equipo`, `tipo`, `origen`, `origenId`, `distNormal` (si no se escribió: longitud total − `distAts`), `caidaNormal`, `caidaFuente`, `caidaComun`, `distAts`, `iccAtsKA` (la mayor Icc de las dos fuentes en el ATS), `ats` (modelo del catálogo), `wcr`, `wcrOk`, `transicion`, `polosAts`, `fasesTxt`, `mat`, `cv`, `longitud`, `kva`. Declara el arreglo `avisos` antes de usarlo. Avisos: NEC 700.5(E) y «WCR del ATS < Icc en el ATS».
- Catálogo `ats` (equipo ATS/MTS, modelo, amperios, polos 3/4, transición abierta/retardada/cerrada, WCR), con migración y tabla editable en administración.
- `js/ui-memoria.js`, tarjeta «Segunda acometida»: Equipo (ATS/MTS/IP), Fuente (generador/bypass), Tipo de transferencia, Polos, Modelo (catálogo, filtrado), WCR, «Generador → ATS (m)» (`longitud`), «Fuente normal → ATS (m)» (`distNormal`, placeholder Auto) y «Distancia ATS → tablero (m)» (`distAts`). La etiqueta de la fuente normal: Transformador, o el tablero padre, o el texto «Conectado a» del tablero (p. ej. MÓDULO MEDIDORES). La tabla de resultados trae una fila por tramo (metros y ΔV) y «Icc en el ATS / WCR» con ✓/✗.
**Aceptación:** generador 100 kVA a 25 m, normal 12 m, ATS→tablero 8 m → largo total 20 m; aparece el modelo del catálogo; con WCR 1 kA sale el aviso.

### P3. Unifilar: espacio y cableado ATS → tablero (web.13)
En `js/ui-unifilar.js`, `bajada()`: deja 58 px entre el ATS/MTS y la barra (`oy = y1 - 58 - ha`) y escribe en ese tramo, a la izquierda de la línea y alineado a la derecha, «fases+neutro+tierra material aislamiento» y «ATS → tablero X m · tubería Ø tipo · ΔV Y %». Sube el texto del cable normal (`y1 - 176` con segunda acometida) y escribe «Normal → ATS X m · …». Rotula NORMAL con su origen (transformador / nombre del padre / texto «Conectado a») y EMERGENCIA en la entrada alterna; el texto de la otra fuente dice «Gen → ATS …» o «Bypass → ATS …». Como `dibujar(R, D)` sirve a SVG y DXF, el DXF lo hereda.
**Aceptación:** el ATS queda separado de la barra y entre ambos se lee el cableado; el rótulo NORMAL dice de dónde viene.

### P4. Datos de ATS/MTS y UPS en la tabla del diagrama y en las tablas resumen (web.13)
- `tablaDatos(r, x, y)`: con segunda acometida agrega filas (equipo/fuente, modelo y amperios, transición y polos, WCR / Icc en el ATS, normal→ATS, generador→ATS, ATS→tablero, kVA del generador u origen del bypass, calibre y caída); con UPS agrega nombre, kVA e Icc del inversor.
- `js/ui-resumen.js`: `Resumen.conTransf` (se activa en `ordenados` si algún tablero tiene alterna o ups), `Resumen.transf(r)` con 15 valores en este orden: Equipo, Modelo de referencia, Amperios (A), Transición, # polos, WCR (kA), Icc en el ATS (kA), Fuente alterna, Normal → ATS (m), Alterna → ATS (m), ATS → tablero (m), Conductores 2ª acometida, Caída 2ª acometida (%), UPS (kVA), Icc del inversor (kA); `Resumen.subTransf`; grupo «DATOS TRANSFERENCIA (ATS / MTS) Y UPS» en `grupos` y `sub`; concatenado en `horizontal`; filas equivalentes en `vertical` y `du` cuando `conTransf`.
- `js/hoja.js`: anchos AV–BJ, título del bloque y 15 subtítulos; límite BJ si `conTransf`, AU si no.
- `js/export-excel.js`: nombre definido `DATOS_TRANSFERENCIA` sobre AV:BJ si `conTransf`.
**Aceptación:** con un ATS y una UPS, «Tablas resumen» muestra el bloque nuevo a la derecha de «DATOS INTERRUPTOR PRINCIPAL»; sin ATS ni UPS no aparece.

### P5. UPS con MTS: a qué tablero se conecta y tablero aguas abajo — interfaz (web.14)
En `js/ui-memoria.js`, tarjeta «Fuentes de alimentación», cuando la UPS está activa:
- `propio` = `u.propio` si está definido; si no, verdadero cuando el tablero no tiene padre. `antecesores` = tableros de la cadena `padreId` hacia arriba.
- Selector «Dónde se conecta la UPS»: opción `'__propio'` («A la barra de este tablero (la UPS cuelga de él)») y una por cada tablero de `Store.posiblesPadres(t)` («A la barra de X (este tablero queda aguas abajo)»). `'__propio'` → `u.propio = true`. Otro tablero → `u.propio = false`, `Store.cambiarPadre(t, v)` (crea el circuito en ese tablero) y, si la segunda acometida es un MTS de bypass desde tablero, `alterna.origenId = v`.
- Si NO es propio: casilla «La UPS lleva MTS» que fija `alterna = { activo: true, equipo: 'MTS', tipo: 'tablero', origenId: t.padreId, polosAts: 3 }` (o la apaga). `Store.sincronizarBypass` crea en el tablero de arriba el segundo circuito (`respaldoDe`), que ocupa posición pero no suma carga.
- Si es propio: casilla `u.mts` y selector «Tablero conectado aguas abajo de la UPS / MTS» con todos los tableros que no estén en `antecesores` ni sean el actual; al elegir uno: `deUps = true` solo en ese (false en los demás) y, si su padre no es este tablero, `Store.cambiarPadre(d, t.id)`.
- `js/calc.js`: `origenId` en el objeto `alterna` del resultado.
**Aceptación:** en el tablero principal el primer selector ofrece «A la barra de este tablero» y el segundo lista todos los demás; en un hijo, el primero lista al principal y a los demás posibles padres; al elegir uno se crea su circuito en el tablero de arriba.

### P6. UPS con MTS en el unifilar — dibujo (web.14)
En `dibujar(R, D)` de `js/ui-unifilar.js`: `todos(r)`; `propio(r)`; `hijoUps(r)` (hijo con ups no propia); `ramaUps(r)`; `hijoU(r)` (hijoUps, o hijo con `deUps` si es propio); `hijos(r)` = todos sin `hijoU` (así no sale en la fila normal); `prof()` usa `todos`. `medir`: suma `UXW = 520` al ancho de un tablero con rama de UPS (o el ancho del tablero de abajo + 170 si es mayor); `D.tamano` suma una altura `NIVEL` si algún tablero tiene `hijoU`. Elimina el dibujo de la UPS en línea de `bajada()`. `ramaUpsDibujo(r, xu, yb, hu)` (llamada desde `nodo`): en la barra del tablero de origen, interruptor «UPS X kVA» en `xu − 45` (o `xu` sin MTS) → rectángulo UPS (nombre, kVA, Icc del inversor) → símbolo MTS «MTS (bypass UPS)»; con MTS, segundo interruptor «Bypass» en `xu + 45` que entra por el puerto alterno del MTS; devuelve `[x1, yS]`. Sin tablero abajo: interruptor «Salida UPS» y recuadro «Cargas críticas». En `nodo`: `xu = max(centro + W/2, x de hijos) + (hay hijos ? 380 : 90)`, la barra se extiende hasta `xu + 150`, y con `hijoU`: `yh = yS + NIVEL; bajada(x1, yS, yh, hu, hu.circuitoPadre); nodo(hu, x1 − span[hu]/2, yh)` (mismo espaciado que cualquier hijo).
**Aceptación:** la UPS cuelga de la barra elegida con dos interruptores, UPS y MTS; el tablero queda debajo con su formato normal; ningún texto se monta sobre la barra ni sobre otro texto; el lienzo crece lo necesario. Verifica solapes midiendo con `getBBox()` de los textos del SVG, no solo a ojo.

### P7. Zoom y mano en el unifilar (web.14)
En `App.views.unifilar`: variables del módulo `zoomU = 1`, `manoU = true`. Guarda el tamaño original en `svg.dataset.w/h` y aplica ancho/alto × zoom sin redibujar. Botones «−», porcentaje, «+» (×1,25; 10 %–400 %), «100 %», «Ajustar» (cabe el ancho; máx. 100 %) y «✋ Mover»; el zoom mantiene centrada la zona visible; Ctrl + rueda acerca/aleja sobre el cursor (listener `wheel` no pasivo). Mano: arrastrar en `.u-wrap` (pointerdown/move/up) desplaza `scrollLeft/scrollTop`; también con el botón central o la barra espaciadora (no con foco en un campo); un arrastre de más de 4 px no cuenta como clic (captura de `click`); NO uses `setPointerCapture` al empezar sino al detectar movimiento (si no se rompen los clics sobre las barras); no acumules listeners de teclado al refrescar; excluye `.u-add`. CSS: `.u-wrap.mano {cursor: grab}`, `.arrastrando {cursor: grabbing; user-select: none}`; oculta el zoom al imprimir. La descarga SVG usa el tamaño original.
**Aceptación:** los botones cambian tamaño y porcentaje; Ctrl + rueda acerca; arrastrar mueve; un clic sin movimiento sobre una barra abre su memoria.

### P8. Excel con macros (solo si este proyecto tiene la carpeta `excel/`)
- `excel/build.py`: en la hoja de cada tablero agrega BX118–BX132 con `rx()`: equipo, modelo ATS, amperios, transición, polos (usa `VALUE(LEFT(RIGHT($I$22,2),1))`, no `N()`), WCR, Icc en el ATS (kA), fuente alterna, normal→ATS (m) = `MAX(0, BX40 − I23)`, alterna→ATS, ATS→tablero, conductores 2ª acometida, caída (%), UPS (kVA), Icc del inversor (kA); vacías si no hay ATS/UPS. «TABLA RESUMEN»: columnas AV–BJ (títulos, anchos, título del bloque).
- `excel/vba/ModTableros.bas`: `ActualizarResumen` con rangos hasta BJ y refs a BX118…BX132; `ResumenVertical` +15 campos (el nombre `RESUMEN_VERTICAL` llega a la fila 48); `ResumenDU` +2 columnas (nombre `RESUMEN_DU` hasta Q); `NombresDefinidos`: `DATOS_TRANSFERENCIA` sobre AV:BJ; `SoloTablaResumen` oculta Z:BJ.
- `excel/vba/ModDXF.bas`: ATS/MTS 22 unidades sobre la barra, cableado y distancia en `E-ALIMENTADOR`, rótulo NORMAL con origen, 5 filas de ATS y 1 de UPS en la tabla de datos; con ATS/MTS en un principal sube lo de arriba 21 unidades (`dg`).
- Pruebas: `excel/test_ats.py`; `excel/construir.sh` debe dar 238, 125 y 35 comprobaciones sin fallos; DXF validado en AutoCAD (`_.AUDIT`, 0 errores; con Git Bash usa `MSYS_NO_PATHCONV=1`; en AutoCAD en español la respuesta es «S»).
- Trampas de VBA: no distingue mayúsculas (no uses una local `yS` si existe `ys`); evita `Val()` con celdas (usa `NumV`); evita `TEXT()` en fórmulas (usa `FIXED`); en DXF R12 no escribas `66 0` en un INSERT sin atributos.
Si no hay Excel, Windows o AutoCAD disponibles, dilo en el informe final y deja este punto como pendiente; no lo marques como hecho.

## 2. Verificación final obligatoria
Antes de dar por terminado, comprueba y reporta cada línea con ✓ o ✗ y la evidencia:
1. `tests/calc-test.html` → 269 de 269. `tests/ui-test.html` → 10 de 10.
2. Consola del navegador sin errores nuevos al abrir cada pestaña (Proyecto, Unifilar, Memoria, Tablas resumen).
3. Escenario manual (o por script en el navegador): tablero principal + 4 hijos; en el principal activa segunda acometida con ATS (generador 100 kVA a 25 m, normal 12 m, ATS→tablero 8 m) → largo 20 m y aviso de WCR si pones 1 kA.
4. Escenario UPS: en un hijo activa UPS y conéctala al principal → aparece el circuito en el principal; marca MTS → segundo circuito de bypass; el unifilar muestra UPS y MTS colgando de la barra del principal con el hijo debajo. En el principal activa UPS «a la barra de este tablero» y elige un tablero aguas abajo → se dibuja debajo del MTS.
5. «Tablas resumen» con ATS y UPS muestra el bloque AV–BJ; sin ellos no aparece; el Excel exportado trae `DATOS_TRANSFERENCIA`.
6. Zoom y mano funcionan; un clic sin arrastre sigue abriendo la memoria de la barra.
7. Un proyecto guardado antes (con `localStorage` viejo) abre sin errores.
8. Versión y etiquetas `?v=` actualizadas; README con una sección por versión.

Entrega al final un informe corto: qué ya existía, qué cambiaste (archivos), qué pruebas corriste y su resultado, qué quedó pendiente y por qué. No inventes resultados: si algo no pudo probarse, dilo.

## 3. Pendientes conocidos de la referencia (no los «arregles» a menos que se pida)
- El cálculo trata al principal con UPS como si estuviera alimentado por la UPS (limita su Icc y avisa «UPS cargada al 1546 %»).
- Cuando la UPS cuelga de su propia barra no se crea el circuito de la UPS ni se suma su carga.
- Faltan amperios, modelo del MTS y SPD en el dibujo de la UPS.
- La contraseña inicial de administrador está escrita en el código: cámbiela antes de publicar el repositorio.
