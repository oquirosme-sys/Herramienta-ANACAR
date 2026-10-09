#!/bin/bash
# Genera "ANACAR Web 2026.xlsm" a partir del catálogo de la herramienta web (js/data/seed.js) y lo prueba.
# Requiere Excel (Windows) con "Confiar en el acceso al modelo de objetos de VBA" y Python con openpyxl, pywin32, pillow y pymupdf.
set -e
cd "$(dirname "$0")"; R=..
python build.py "$R/js/data/seed.js" plantilla.xlsx "$R/img/logo-sinergia.png"
python importar.py plantilla.xlsx "$R/ANACAR Web 2026.xlsm" vba
python test_machote.py "$R/ANACAR Web 2026.xlsm" "$R/tests/machote.js" res.json     # 238 comparaciones con el Machote del Excel
python test_escenarios.py "$R/ANACAR Web 2026.xlsm"                               # 125 comparaciones con el motor web
python test_flujo.py "$R/ANACAR Web 2026.xlsm" revit_prueba.txt salida             # Revit, prevista, balanceo, DXF, PDF, tablas
