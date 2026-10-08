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
`tests/calc-test.html` compara la herramienta con los valores que calculó Excel para el Machote (256 comprobaciones: el Machote del Excel, la cascada, el transformador, las reglas de fabricante y el autobalanceo) y prueba la cascada. Ábralo con un servidor local, p. ej. `py -m http.server` en esta carpeta y `http://localhost:8000/tests/calc-test.html`.

Mejoras respecto al Excel (indicadas en pantalla): tablero, interruptor principal, breakers y supresor se eligen **automáticamente** según la marca, la protección y los polos si no se escoge uno; conductores en paralelo del alimentador automáticos (>300 A); avisos de posiciones repetidas, caída de voltaje, SCCR menor que el Icc, ampacidad corregida insuficiente y espacios insuficientes.

## Modo administrador (paso 1)
Contraseña inicial: `sinergia-admin` (cámbiela en *Administración ▸ Seguridad y respaldo*).
En una página estática esto **solo oculta** la edición; la restricción real se implementa en el paso 2 con Supabase Auth + Row Level Security.

## Datos (paso 1)
Proyecto y catálogo se guardan en `localStorage` del navegador. Use *Archivo ▸ Guardar proyecto (.json)* y *Administración ▸ Exportar catálogo* para respaldar o compartir. El catálogo inicial sale del Excel (`js/data/seed.js`, generado automáticamente).

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
js/ui-proyecto.js  ui-unifilar.js  ui-revit.js  ui-memoria.js  ui-cuadro.js  ui-resumen.js  ui-admin.js  export-excel.js
img/logo-sinergia.png
js/data/seed.js        (catálogo generado del Excel)
tests/calc-test.html   tests/machote.js
```
