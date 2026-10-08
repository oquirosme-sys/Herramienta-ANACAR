# ANACAR — Análisis de cargas de tableros eléctricos (web)

Herramienta HTML basada en `ANACAR 2026.1.xlsm`. Sin dependencias ni compilación: se publica tal cual en **GitHub Pages**.

## Publicar en GitHub Pages
1. Suba todo el contenido de esta carpeta a la raíz de un repositorio (incluye `.github/workflows/static.yml`).
2. *Settings ▸ Pages ▸ Source: GitHub Actions* (o *Deploy from a branch ▸ `main` / `(root)`*).
3. Abra `https://<usuario>.github.io/<repositorio>/`.

## Cómo está organizada
| Pestaña | Qué hace | Equivale en el Excel a |
|---|---|---|
| **Proyecto** | Datos del proyecto, criterios (ΔV máx., desbalance, marca), alta de tableros (uno o varios), **de cuál se alimenta cada uno** y diagrama de alimentación. | — (antes: copiar el Machote con Ctrl+T) |
| **Importar de Revit** | Lee la tabla de circuitos de Revit (.txt/.csv exportado, el Excel *Circuitos revit-excel.xlsm* o filas pegadas), agrupa por *Panel*, propone el detalle de carga y crea/actualiza los tableros. | Macro `ImportarCircuitosRevit` |
| **Memoria de cálculo** | Un módulo por tablero: datos del tablero, circuitos ramales, factores de demanda/diversidad, alimentador con caída de voltaje acumulada, cortocircuito, verificación de ampacidad y balanceo sugerido. | Bloque de la hoja `ANACAR` + hoja `BALANCEO` |
| **Tableros 3F / 1F** | Cuadro de cargas de cada tablero con el formato de las hojas del Excel y la vista de posiciones; imprimible. | `TABLEROS 3F`, `TABLEROS 1F` |
| **Tablas resumen** | Horizontal, vertical, DU y original, con orden y exclusión de tableros; exportación a Excel/CSV para los planos de Revit. | `TABLA RESUMEN…`, `ORDEN TABLEROS` |
| **Administración** *(solo admin)* | Marcas, tableros, breakers, supresores, tipos y detalles de carga, tipos de cableado, tablas NEC, transformadores, respaldo. | `DCARGAS`, `CONDISEÑO`, `DATOS` (ocultas) |

Las hojas informativas u ocultas (`INSTRUCCIONES`, `CÓDIGO VBA`, `FACTORES POR CARGA`, `CARGAS CONTINUAS`, `GEN-ATS-IP`, `UPS-01`, `TRAFOS`) no aparecen: su contenido útil quedó en el catálogo (p. ej. el factor multiplicador y si la carga es continua de cada detalle).

## Cascada entre tableros
- La **carga** de un tablero derivado (kVA totales con reserva) sube automáticamente al circuito que lo alimenta en el tablero padre; la longitud de ese circuito es la del alimentador del derivado.
- El **voltaje en bornes** del padre es el punto de partida del alimentador del derivado, por lo que la **caída de voltaje es acumulada** desde la acometida (en el Excel cada tablero partía del voltaje nominal).
- El **cortocircuito** baja por cada alimentador (método punto a punto con las constantes C de la hoja DATOS). En el tablero principal se indica el Icc o se elige el transformador.
- Si el derivado tiene otro voltaje (transformador de por medio), la caída y el Icc se reinician y se avisa.

## Motor de cálculo y verificación
`js/calc.js` replica las fórmulas del bloque *Machote* (reparto por fases de las funciones VBA `POLA…POLE`, corrientes, protección, calibres, neutro, tierra 250.122, electrodo 250.66, tubería C.10, FAC, caída de voltaje, breakers y supresor por marca, factores de demanda con la regla de tomas 10 kVA + 50 %).
`tests/calc-test.html` compara la herramienta con los valores que calculó Excel para el Machote (249 comprobaciones) y prueba la cascada. Ábralo con un servidor local, p. ej. `py -m http.server` en esta carpeta y `http://localhost:8000/tests/calc-test.html`.

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
js/ui-proyecto.js  ui-revit.js  ui-memoria.js  ui-cuadro.js  ui-resumen.js  ui-admin.js  export-excel.js
js/data/seed.js        (catálogo generado del Excel)
tests/calc-test.html   tests/machote.js
```
