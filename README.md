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

## Cambios de la versión SNG v2 — 2026.1-web.16
Se aplicaron sin cambiar colores, estilos ni la integración con la Suite (base de datos, accesos y roles no se tocaron):
- **Fuentes del proyecto** (Proyecto ▸ Fuentes): transformadores y generadores compartidos por varios tableros y ATS; los tableros que cuelgan de una fuente toman su voltaje secundario y la carga de la fuente suma la demanda de todos. «Alimentado desde» ofrece también esas fuentes.
- **Calibres mínimos** de un alimentador: 8 AWG en fases y neutro y 10 AWG en tierra (nunca 12 AWG), y **acometida manual** (calibre de fases, neutro y tierra) con avisos si no cumple.
- **ATS / MTS:** el tramo ATS → tablero es la longitud del tablero; la acometida normal → ATS (`distNormal`) se suma a ella. El campo separado «Distancia ATS → tablero» desapareció.
- **Zoom y mano** en las tablas resumen y en los cuadros 3F/1F (`Hoja.zoomUI`), y barra de desplazamiento superior en la tabla de circuitos.
- **Formato del cuadro de cargas** (Proyecto): «Formato original (Costa Rica)» (por defecto) o «Panamá (formato TL4)» con 9 dispositivos por circuito.
- La unidad del interruptor pedida en el circuito se respeta en las tablas y avisa si no hay interruptor de esa unidad en la familia del tablero.
- Tablas verticales y DU: bloque «UPS / Segunda acometida» con las distancias de cada tramo.
- **Excel:** `I23` es ahora la longitud del tablero (ATS → tablero), alimentador con 8 AWG mínimo y tierra 10 AWG, acometida manual (`F20`–`F22`). Fuentes compartidas, unidad pedida y formato Panamá no se portaron al Excel.

## UPS con MTS — 2026.1-web.14
- **Excel:** nueva entrada «La UPS lleva MTS» (I13) que activa el MTS de bypass desde el tablero de arriba (con su circuito) o, sin tablero de arriba, elige el MTS por la corriente de la UPS y lo muestra en las tablas resumen y en el DXF. En el Excel la UPS sigue dibujada en línea (no colgando de la barra).
- En *Fuentes de alimentación*, al marcar *A través de una UPS* aparece **«La UPS lleva MTS (bypass de mantenimiento)»**: el tablero que alimenta la UPS recibe un segundo circuito con su interruptor (el bypass, que ocupa espacio pero no suma carga) y el unifilar muestra el MTS con las dos acometidas. Requiere que el tablero esté alimentado desde otro tablero.

## Casillas y datos de ATS/UPS en las tablas — 2026.1-web.12 / web.13
- **web.12:** las casillas numéricas (kVA del generador, distancias, etc.) ya no se borran al cambiar otro dato (prueba `tests/ui-test.html`).
- **web.13:** en el unifilar hay espacio entre el ATS/MTS y la barra con el cableado del tramo común; al ATS llegan la acometida normal (transformador, módulo de medidores o tablero padre) y el generador/bypass. La tabla de datos del diagrama y las tablas resumen (horizontal, vertical y DU) incluyen los datos de ATS/MTS y UPS (bloque «DATOS TRANSFERENCIA», columnas AV–BJ; nombre definido `DATOS_TRANSFERENCIA`). El Excel con macros trae lo mismo, incluido el DXF.

## ATS / MTS con las dos acometidas y sus distancias — 2026.1-web.11
- Al marcar *Lleva segunda acometida* con ATS o MTS, al equipo llegan **las dos acometidas**: la **normal** (del transformador o del tablero que alimenta) y la **de emergencia** (generador o bypass). Se indican tres distancias: *Transformador → ATS*, *Generador → ATS* y *ATS → tablero* (tramo común). Con la distancia normal indicada, el largo del alimentador pasa a ser normal + ATS → tablero.
- La memoria muestra por tramo la longitud y la caída de voltaje, la Icc en el ATS (la mayor de las dos fuentes) y la comparación con su WCR. El unifilar rotula las entradas NORMAL y EMERGENCIA y las distancias; el Excel agrega la entrada *Distancia fuente normal → ATS* (I24).

## Unifilar con bloques de los DU — 2026.1-web.10
- El DXF (web y Excel) ahora lleva **bloques de AutoCAD** con la geometría de tus DU (`DU prueba.dwg`, `DU ELÉCTRICO 3.dwg`): `INTERRUPTOR_1P/2P/3P` (contactos con arco y un trazo por polo), `TRAFO_SECO`, `ATS` y `MTS` (normal arriba, alterna a la derecha, carga abajo), `GENERADOR` y `TIERRA`. Se ven igual en la pantalla de la herramienta.
- Cada bloque trae **atributos** (doble clic en AutoCAD o extracción de atributos): interruptor (polos, amperios, unidad, modelo, SCCR), trafo (kVA, Z, voltajes), ATS/MTS (modelo, amperios, WCR, transición), generador (nombre, kVA, X''d). Quedan invisibles para no recargar el dibujo; el texto visible sigue siendo el de antes.
- Las líneas se cortan donde hay símbolos; el tablero de acometida lleva su barra de tierra con el conductor del electrodo.
- Para usar sus propios bloques: en AutoCAD redefina el bloque con el mismo nombre (por ejemplo `BLOCKREPLACE`). La biblioteca vive en `js/data/simbolos.js` y se copia a la hoja oculta `CAT_SIMBOLOS` del Excel.
- Validado abriendo los DXF (web y Excel) en AutoCAD 2026: 0 errores de auditoría.

## Segunda acometida: ubicación y tipo de ATS — 2026.1-web.9
- **Distancias**: *longitud fuente → ATS* y *distancia ATS → tablero* (tramo común con el alimentador normal). La caída de voltaje y la corriente de cortocircuito de la fuente alterna se calculan por tramos, y se obtiene la **Icc en el ATS** (la mayor entre la alimentación normal y la alterna en ese punto).
- **Tipo de transferencia**: transición abierta, retardada (motores), cerrada o con bypass-aislamiento, y MTS; **3 o 4 polos** (neutro conmutado, fuente derivada separadamente 250.30).
- **Modelo de catálogo** (Administración ▸ Transferencias; hoja CAT_ATS en Excel): se elige el de esa transición y polos con amperaje ≥ protección y **WCR ≥ Icc en el ATS**; se puede fijar el modelo o escribir el WCR del fabricante. Los WCR genéricos son típicos: reemplácelos por los del fabricante.
- Aviso y revisión NEC 700.5(E)/702.5 cuando el WCR es menor que la Icc en el ATS; el diagrama unifilar muestra el modelo y la distancia.
- Corregido: un aviso NEC de nivel "error" detenía el cálculo del tablero.

## Detalles de carga (como DCARGAS) — 2026.1-web.8
- Nueva pestaña **Detalles de carga** (web) y hoja **DETALLES DE CARGA** (Excel), editables por todos: voltaje, fases, hilos, tipo de carga, factores, FP, continua, multiplicador y **palabras clave (Revit)**. En la web también se editan los tipos de carga (demanda). No se puede borrar un detalle que use el proyecto.
- **Importación de Revit**: el detalle se elige primero por las palabras clave ("uc-" reconoce "UC-01", "ue" a "UE2-03", "dh-" a "DH-03", "calentador de paso"…). El ⚠ indica un detalle propuesto sin coincidencia segura: al corregirlo se aplica a los circuitos con el mismo nombre y **se recuerda** como palabra clave para la próxima importación. Los alimentadores de tableros toman el detalle de tablero. En el Excel los detalles por revisar quedan en naranja con un comentario.
- **Cargas monofásicas a 208/240 V en 2 polos** (calentador de paso, minisplit, cargador de vehículo…): la carga se reparte entre los 2 polos con la corriente correcta (antes se duplicaba en cada polo). Los circuitos de 2 polos también buscan estos detalles.
- Nuevo detalle 67 **DESHUMIDIFICADOR DH-XX** (120 V).
- Con el proyecto real: de 43 circuitos ⚠ quedan 0 al reconocer las palabras clave, salvo "Other".

## Proceso Revit → ANACAR (sin pérdida de circuitos) — 2026.1-web.7
1. **Revit**: exportar la tabla *Electrical Circuit Schedule* (.csv).
2. **`Circuitos revit-excel.xlsm`** (versión nueva en esta carpeta): botón *IMPORTAR DATOS* → limpia y acomoda la tabla de *Hoja1* (mismas columnas C–J) y deja la hoja **Revision** con líneas leídas, circuitos escritos, filas omitidas con su motivo, kVA/polos por tablero, posiciones repetidas, longitudes faltantes y qué circuito alimenta a cada tablero.
   - Lee UTF-8 (acentos correctos), busca las columnas por nombre (*True Power* o *True Load*), lee números con punto o coma sin depender de la configuración regional y guarda el número de circuito como texto ("24,26").
   - Ya no descarta filas sin voltaje; las filas sin tablero (resumen de Revit: tablero principal, ATS) se listan como omitidas.
   - **Siempre empieza en blanco**: al abrirlo se vacían la tabla y la hoja Revision. El archivo guardado conserva el último resultado para que ANACAR lo importe (ANACAR lo lee sin ejecutar esa macro).
3. **ANACAR** (web o Excel) importa ese libro:
   - Reconoce *True Power* y el circuito como texto.
   - **Enlaza los alimentadores** a su tablero derivado: por el nombre de la carga o, si no coincide, por la misma carga (p. ej. "2PIS1" → 2MS1). La longitud del circuito pasa al alimentador del derivado.
   - Muestra la **conciliación**: circuitos y kVA por tablero en el archivo frente a lo importado, y los movidos de posición.
   - Al reimportar se actualiza sin duplicar.

Probado con el proyecto real: 86 líneas → 84 circuitos en 7 tableros + 2 filas de resumen; 0 perdidos; kVA idénticos por tablero (web y Excel). Los scripts y la macro del formateador están en `excel/revit/`.

## Novedades 2026.1-web.6
- **Pestañas**: Prevista y Revit pasan al final; nueva pestaña **Ayuda** con guía de llenado, glosario básico de equipos (tablero, centro de carga, subestación, principal/zapatas, breakers, caja moldeada, unidades LSI/LSIG, GFCI/AFCI, SCCR y serie, SPD, transformador, ATS/MTS, UPS, generador, conductores, canalizaciones, barras, factores, cascada) con ilustraciones, y catálogo de tableros y breakers por familia. El administrador puede subir una foto real para cada término o familia.
- **Factores de diversidad entre 0 y 1**: multiplican la carga (antes dividían); los valores guardados mayores que 1 se convierten a 1/x al abrir, con el mismo resultado.
- **Celdas para llenar** en verde más intenso con borde verde (web y Excel).
- **Cuadros 3F/1F** ajustados para que quepan completos en pantalla.
- **Excel**: mismos cambios; hoja AYUDA con la guía, el glosario ilustrado y el catálogo; los tableros se insertan antes de PREVISTA / REVIT.

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
