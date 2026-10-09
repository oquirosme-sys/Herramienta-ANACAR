# ANACAR — Análisis de cargas de tableros eléctricos (web)

Herramienta HTML basada en `ANACAR 2026.1.xlsm`. Sin dependencias ni compilación: se publica tal cual en **GitHub Pages**.

## Publicar en GitHub Pages
1. Suba todo el contenido de esta carpeta a la raíz de un repositorio (incluye `.github/workflows/static.yml`).
2. *Settings ▸ Pages ▸ Source: GitHub Actions* (o *Deploy from a branch ▸ `main` / `(root)`*).
3. Abra `https://<usuario>.github.io/<repositorio>/`.

## Cómo está organizada
| Pestaña | Qué hace | Equivale en el Excel a |
|---|---|---|
| **Proyecto** | Datos del proyecto, criterios (ΔV máx., desbalance, marca, reserva de espacios, Icc de la red, ajuste automático de calibre), alta de tableros y de cuál se alimenta cada uno. | — (antes: copiar el Machote con Ctrl+T) |
| **Diagrama unifilar** | Diagrama tipo ETAP/SKM: red → transformador → interruptor → alimentador → barra de cada tablero, con voltaje en bornes, caída acumulada y cortocircuito. Desde aquí se agregan transformadores y tableros derivados. | — |
| **Revit** | *Importar circuitos* (.txt/.csv exportado, el Excel *Circuitos revit-excel.xlsm* o filas pegadas) y *Cambios para Revit*: lista de circuitos que cambiaron de posición (autobalanceo o edición) para pasarlos al modelo; se copia o exporta a CSV. | Macro `ImportarCircuitosRevit` |
| **Memoria de cálculo** | Un módulo por tablero: datos, circuitos ramales, factores de demanda, alimentador con validación por temperatura y agrupamiento, transformador, cortocircuito y **autobalanceo**. | Bloque de la hoja `ANACAR` + `BALANCEO` |
| **Tableros 3F / 1F** | Cuadro de cargas con el **mismo diseño de la hoja del Excel** (título verde, logo, barras grises, posiciones impares y luego pares); imprimible y exportable. | `TABLEROS 3F`, `TABLEROS 1F` |
| **Tablas resumen** | Horizontal, vertical, DU y original con el diseño de las hojas del Excel; orden y exclusión de tableros; exportación a Excel (con logo) y CSV. | `TABLA RESUMEN…`, `ORDEN TABLEROS` |
| **Administración** *(solo admin)* | Marcas, tableros, breakers (con SCCR a 240/480 V), supresores, **reglas de fabricante**, tipos y detalles de carga, cableado, tablas NEC, transformadores, respaldo. | `DCARGAS`, `CONDISEÑO`, `DATOS` (ocultas) |

Las hojas informativas u ocultas (`INSTRUCCIONES`, `CÓDIGO VBA`, `FACTORES POR CARGA`, `CARGAS CONTINUAS`, `GEN-ATS-IP`, `UPS-01`, `TRAFOS`) no aparecen: su contenido útil quedó en el catálogo (p. ej. el factor multiplicador y si la carga es continua de cada detalle).

## Cascada entre tableros
- La **carga** de un tablero derivado (kVA totales con reserva) sube automáticamente al circuito que lo alimenta en el tablero padre; la longitud de ese circuito es la del alimentador del derivado.
- El **voltaje en bornes** del padre es el punto de partida del alimentador del derivado, por lo que la **caída de voltaje es acumulada** desde la acometida (en el Excel cada tablero partía del voltaje nominal).
- El **cortocircuito** baja por cada alimentador (método punto a punto con las constantes C de la hoja DATOS). En el tablero principal se indica el Icc o se elige el transformador.
- Si el derivado tiene otro voltaje (transformador de por medio), la caída y el Icc se reinician y se avisa.

## Versión Excel con macros (`ANACAR Web 2026.xlsm`)
Libro de Excel con las mismas funciones que la herramienta en línea, para trabajar sin conexión. Se abre limpio (sin tableros) y hay que **habilitar las macros**.

| Hoja | Qué hace |
|---|---|
| **INICIO** | Botones: nuevo tablero, actualizar proyecto y tablas, importar/actualizar desde Revit, autobalancear (activo o todos), prevista, exportar PDF, unifilar DXF, tablas para Revit, modo administrador. |
| **PROYECTO** | Datos del proyecto y criterios (ΔV máx., desbalance, Icc de la red, marca, reserva de espacios, ocupación NEC, carga de derivados, aumento automático de calibre). Lista de tableros con vínculos, orden y exclusión para las tablas resumen. |
| **MC &lt;tablero&gt;** | Memoria de cálculo con **fórmulas** (como el bloque Machote): datos del tablero, alimentador (TM/LI/LSI/LSIG, 100 % rated), transformador, UPS, segunda acometida (ATS/MTS/IP, generador o bypass), circuitos en filas por posición (impares y luego pares), factores de demanda y diversidad, resultados en la columna BX. *Alimentado desde* conecta el tablero a su padre y crea su circuito (cascada de carga, caída de voltaje y cortocircuito). |
| **3F/1F &lt;tablero&gt;** | Cuadro de cargas con el diseño de la hoja `TABLEROS 3F`; oculta las posiciones que no existen en el tablero elegido. Nombre definido `TABLERO_<tablero>` para Revit. |
| **TABLA RESUMEN / RESUMEN VERTICAL / RESUMEN DU** | Tablas para los planos con sus nombres definidos (`TABLA_RESUMEN__TABLEROS_ELÉCTRICOS`, `DATOS_DEL_TABLERO`, `DATOS_DEL_SUPRESOR`, `DATOS_INTERRUPTOR_PRINCIPAL`…); botón *Solo tabla resumen*. |
| **PREVISTA** | Carga por tipo de uso y m² y generación de circuitos de prevista. |
| **REVIT / CAMBIOS REVIT** | Importar o actualizar desde la tabla de circuitos (archivo o pegado) y lista de cambios de posición del autobalanceo. |

**Autollenado como en línea**: las celdas de datos traen en *gris cursiva* el valor automático (descripción, % de protección, factores de uso y diversidad, material, aislamiento, # en paralelo, protección, unidad GFCI/AFCI, breaker, marca, agrupamiento, tablero de catálogo, SPD, voltaje primario); se escribe encima para cambiarlo y se borra para volver al automático. Detalle de carga, breaker, tablero y SPD se eligen de listas con nombre. Botones: *Crear varios tableros*, *Tablero derivado*, *Duplicar tablero*, *Duplicar circuito*, *Mover circuito* (anota el cambio en CAMBIOS REVIT); en PROYECTO se editan tipo, sistema, alimentado desde y longitud.

Las funciones de selección (breakers, tableros, supresores por reglas de fabricante, clasificación en serie, ampacidad 310.15, punto a punto y GFCI/AFCI) son funciones VBA de hoja (`BREAKER_SEL`, `TABLERO_SEL`, `SPD_SEL`, `CAL_AJUSTE`, `P2P`, `NEC_UNIDAD`…) que leen los catálogos de las hojas `CAT_*` (ocultas; se editan en *Modo administrador*). El libro se genera con los scripts de la carpeta `excel/` a partir de `js/data/seed.js`, y se verifica contra los resultados del Machote del Excel original y contra el motor web (`excel/test_*.py`).

## Novedades 2026.1-web.5
- **Arranca limpio**: cada vez que se abre la herramienta empieza con un proyecto en blanco. Los proyectos con datos quedan en *Archivo ▸ Proyectos guardados…* (y en la pantalla de inicio) para abrirlos o eliminarlos; al recargar la página se sigue con el proyecto abierto.

## Novedades 2026.1-web.4
- **Excel con nombres definidos** para vincular en Revit: `TABLERO_<nombre>` en cada hoja de tablero (rango B1:AJ…, como la macro AsignarNombresTableros) y `TABLA_RESUMEN__TABLEROS_ELÉCTRICOS`, `DATOS_DEL_TABLERO`, `DATOS_DEL_SUPRESOR`, `DATOS_INTERRUPTOR_PRINCIPAL` en la tabla resumen.
- **Porcentajes**: el aumento de calibre (0 % = sin aumento) y el factor de protección (125 %) se escriben en %.
- **Orden físico**: en la memoria, el reporte y el cuadro, los circuitos van primero del lado impar y luego del par.
- **Factores de uso y diversidad**: cada circuito usa el factor de uso y el de diversidad de su detalle de carga (o los propios, columnas *F. uso* y *F. div.*); el tablero aplica la demanda por tipo de carga; el factor de diversidad del tablero **divide** la demanda (en el Excel multiplicaba; con el valor 1 por defecto no cambia nada). Opción para que los tableros derivados suban con sus kVA demandados en lugar de conectados.
- **Prevista (diseño esquemático)**: carga por tipo de uso y m² (iluminación NEC 2020 tabla 220.12; tomas y climatización estimadas, editables) y generación de circuitos de prevista en el tablero.
- **Bypass con su protección**: el tablero de origen lleva automáticamente el circuito del bypass (ocupa espacios, no suma carga); en el DU el ATS/MTS va pegado a la barra y la rama alterna lleva su interruptor (el del circuito del bypass o el del generador).
- **Interruptor principal**: unidad de disparo TM, LI, LSI o LSIG y opción 100 % rated (la protección lleva el 100 % de la carga continua; en Power Defense el modelo pasa a PDF…).

## Novedades 2026.1-web.3
- **Impresión centrada**: cada hoja (tableros, tablas resumen) se ajusta al ancho de la página horizontal al imprimir.
- **Tabla resumen**: casilla *Solo tabla resumen* (sin datos del tablero, del supresor ni del interruptor principal), también en la exportación a Excel.
- **Diagrama unifilar en DXF** (AutoCAD lo abre directamente; *Guardar como* → DWG). Capas E-TABLERO, E-ALIMENTADOR, E-PROTECCION, E-TRAFO, E-EQUIPOS, E-TEXTO, E-TABLA; texto en Century Gothic. Opción *Tabla de datos junto a cada tablero* (equipo, Icc, kVA, FD, FP, acometida, longitud, voltajes y % de caída), como en los DU.
  El formato DWG es cerrado y no se puede escribir desde una página web; el DXF es el formato de intercambio de Autodesk.
- **Por tablero**: interruptor principal o zapatas, espacios (polos) del tablero y tipo de equipo (tablero o subestación/switchboard: QED-2, I-Line, PRL4, Pow-R-Line Xpert); el modelo de catálogo se ajusta solo.
- **Fuentes**: UPS en la alimentación normal (salida regulada, Icc limitado por el inversor) y **segunda acometida** con ATS, MTS o interruptor enclavado, desde un generador (kVA, X''d) o como bypass desde otro tablero; se dimensiona su cable, caída e Icc, y el tablero toma el mayor Icc.
- **Revit ▸ Actualizar desde la tabla**: vuelve a leer la tabla de circuitos y solo actualiza carga, longitud y nombre (conserva detalle de carga, conductores, breakers y posiciones fijadas); agrega los nuevos, avisa los que ya no están y reconoce los cambios de posición ya aplicados en Revit.
- **Memoria de cálculo** como documento: imprimir / PDF, Word (.doc) y Excel (una hoja por tablero).
- **NEC 2020** (según los resúmenes de Eaton, CITEC y CIEMI): GFCI 210.8(A)/(B)/(F), 422.5 y 680.21, AFCI 210.12 por ubicación y ocupación (interruptor automático si el circuito no tiene uno elegido); SPD por artículo 242 (kA recomendados según la posición: acometida 250 kA, 600–1200 A 120 kA, hasta 400 A 50 kA; MCOV; 230.67 en vivienda); 110.24 y 408.6 (rótulo de corriente de falla y SCCR); 240.87 (reducción de energía de arco ≥ 1200 A); 700.5(E)/702.5 (SCCR del ATS/MTS).
- **Clasificación en serie (240.86)**: si un ramal no alcanza el Icc pero hay una combinación listada con el principal, se acepta y se indica el rótulo 110.22(C) y la restricción de motores 240.86(C). Tabla inicial con las combinaciones Eaton de la presentación *Series Rating* (parcial), editable en Administración.
- **Números de catálogo**: Power Defense (PDG + marco + polos + letra de capacidad + amperios + disparo + J), GFCI/AFCI de referencia (Eaton QBGF/QBAF/QBAG, Square D QOB…GFI / QO…CAFI).

## Transformadores y cortocircuito
Cada tablero puede tener un **transformador aguas arriba** (kVA, %Z, X/R, voltaje primario). El Icc en el secundario se calcula por el método punto a punto (Bussmann) a partir del Icc del tablero que lo alimenta o de la red (Icc de la red en Proyecto; vacío = red infinita), y luego baja por el alimentador. También se calcula la caída de tensión en el transformador (regulación) y su porcentaje de carga. Con un transformador de 750 kVA y Z 5,75 % la herramienta reproduce los 36 206 A de la tabla del Excel y el Icc en bornes del Machote (AQ10).

## Reglas de fabricante (selección automática)
- **Tablero**: familias de la marca cuyo voltaje máximo y fases sirven (p. ej. NQ ≤ 240 V, NF 480Y/277, I-Line 600 V; CH/PRL1 ≤ 240 V, PRL2 480 V, PRL3/PRL4 600 V), barras ≥ protección principal y espacios ≥ usados + reserva %.
- **Breakers**: solo de las familias permitidas por el tablero (ramales y principales), con los polos y el rango de amperios correctos y el **menor SCCR que supere el Icc** del tablero, usando la capacidad a 240 V o a 480 V según el sistema. El modelo de referencia se arma con la plantilla de la familia (QOB120, QOB230VH, EDB34020, BAB1020, GHB3020, HDL36100…).
- Todo se edita en *Administración ▸ Reglas*. Siempre se puede elegir un modelo manualmente.

## Validación de alimentadores (310.15 / 310.16)
Ampacidad a la temperatura del aislamiento × factor de temperatura ambiente (tabla 310.15(B)(1) completa) × factor de agrupamiento (310.15(C)(1), automático según fases e hilos), limitada por la columna de temperatura de los bornes; debe superar la ampacidad requerida y la protección (240.4(B)). Si no cumple, el calibre se aumenta automáticamente (o se agrega un conductor en paralelo por encima de 500 kcmil); se puede desactivar por proyecto o fijar el calibre por tablero.

## Autobalanceo
Intercambia circuitos con el mismo número de polos y mueve circuitos a espacios libres para minimizar la diferencia entre fases; no mueve los circuitos marcados como *Fijo*. Cada cambio queda anotado (circuito en Revit → circuito nuevo) en *Revit ▸ Cambios para Revit*.

## Motor de cálculo y verificación
`js/calc.js` replica las fórmulas del bloque *Machote* (reparto por fases de las funciones VBA `POLA…POLE`, corrientes, protección, calibres, neutro, tierra 250.122, electrodo 250.66, tubería C.10, FAC, caída de voltaje, breakers y supresor por marca, factores de demanda con la regla de tomas 10 kVA + 50 %).
`tests/calc-test.html` compara la herramienta con los valores que calculó Excel para el Machote (268 comprobaciones: el Machote del Excel, la cascada, el transformador, las reglas de fabricante, el autobalanceo, la clasificación en serie, GFCI/AFCI, diversidad, 100 % rated y bypass) y prueba la cascada. Ábralo con un servidor local, p. ej. `py -m http.server` en esta carpeta y `http://localhost:8000/tests/calc-test.html`.

Mejoras respecto al Excel (indicadas en pantalla): tablero, interruptor principal, breakers y supresor se eligen **automáticamente** según la marca, la protección y los polos si no se escoge uno; conductores en paralelo del alimentador automáticos (>300 A); avisos de posiciones repetidas, caída de voltaje, SCCR menor que el Icc, ampacidad corregida insuficiente y espacios insuficientes.

## Modo administrador (paso 1)
Contraseña inicial: `sinergia-admin` (cámbiela en *Administración ▸ Seguridad y respaldo*).
En una página estática esto **solo oculta** la edición; la restricción real se implementa en el paso 2 con Supabase Auth + Row Level Security.

## Datos (paso 1)
Los proyectos (varios) y el catálogo se guardan en `localStorage` del navegador; el proyecto abierto en cada pestaña se recuerda en `sessionStorage`. Use *Archivo ▸ Guardar proyecto (.json)* y *Administración ▸ Exportar catálogo* para respaldar o compartir. El catálogo inicial sale del Excel (`js/data/seed.js`, generado automáticamente).

## Paso 2 (pendiente): Supabase
Todo el acceso a datos pasa por `js/store.js`; ahí se reemplaza `localStorage` por Supabase. El modelo ya está separado en:
- **Catálogo** (`tiposCarga`, `detallesCarga`, `marcas`, `tablerosCat`, `breakers`, `supresores`, `transformadores`, tablas NEC): lectura para todos, escritura solo rol `admin`.
- **Proyecto** (`proyecto` → `tableros` (con `padreId`) → `circuitos` (con `tableroHijoId`)): por usuario/proyecto.
- `js/config.js` ya tiene los campos `SUPABASE_URL` y `SUPABASE_ANON_KEY`.

## Estructura
```
index.html
css/styles.css
js/config.js  util.js  calc.js  store.js  auth.js  ui.js  xlsx-read.js  app.js
js/hoja.js             (hojas con el diseño del Excel: pantalla y exportación)
js/reporte.js          (memoria de cálculo: PDF, Word, Excel)
js/ui-prevista.js      (prevista por tipo de uso y m²)
js/ui-proyecto.js  ui-unifilar.js  ui-revit.js  ui-memoria.js  ui-cuadro.js  ui-resumen.js  ui-admin.js  export-excel.js
img/logo-sinergia.png
js/data/seed.js        (catálogo generado del Excel)
tests/calc-test.html   tests/machote.js
```
