"""Importa los módulos VBA a la plantilla, agrega botones y guarda como .xlsm."""
import sys, os, glob, time
from xlcom import abrir_excel, donde_error, LOG

SRC, OUT, VBA = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2]), os.path.abspath(sys.argv[3])
TMP = os.path.join(os.path.dirname(SRC), "_vba_cp1252"); os.makedirs(TMP, exist_ok=True)

xl, pid, vig = abrir_excel()
try:
    wb = xl.Workbooks.Open(SRC)
    proj = wb.VBProject
    for f in sorted(glob.glob(os.path.join(VBA, "*.bas"))):
        txt = open(f, encoding="utf8").read().replace("\r\n", "\n").replace("\n", "\r\n")
        dst = os.path.join(TMP, os.path.basename(f))
        open(dst, "w", encoding="cp1252", newline="").write(txt)
        proj.VBComponents.Import(dst)
        print("importado", os.path.basename(f))
    tw = None
    for c in proj.VBComponents:
        if c.Type == 100 and c.Name in ("ThisWorkbook", "EsteLibro"): tw = c
    if tw is None:
        for c in proj.VBComponents:
            if c.Type == 100 and c.Properties("Name").Value == wb.Name: pass
        tw = proj.VBComponents(wb.CodeName)
    cm = tw.CodeModule
    if cm.CountOfLines: cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(open(os.path.join(VBA, "ThisWorkbook.txt"), encoding="utf8").read())
    wb.Names.Add("ADMIN_CLAVE", '="sinergia-admin"', False)

    def boton(ws, x, y, w, h, texto, macro, color=0x3A7A2E):
        s = ws.Shapes.AddShape(5, x, y, w, h)
        s.Fill.ForeColor.RGB = color; s.Line.Visible = False
        tf = s.TextFrame2; tf.TextRange.Text = texto
        tf.TextRange.Font.Size = 10; tf.TextRange.Font.Bold = True; tf.TextRange.Font.Name = "Century Gothic"
        tf.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
        tf.VerticalAnchor = 3; tf.TextRange.ParagraphFormat.Alignment = 2
        s.OnAction = macro; s.Placement = 3
        try: ws.Shapes(s.Name).DrawingObject.PrintObject = False
        except Exception as e: print("print", e)
        return s

    VERDE, AZUL, GRIS, ROJO = 0x3A7A2E, 0x9C5B1F, 0x606060, 0x2B2BB0
    ini = wb.Worksheets("INICIO")
    botones = [("Nuevo tablero", "NuevoTablero", VERDE), ("Crear varios tableros", "CrearVarios", VERDE), ("Tablero derivado del activo", "NuevoDerivado", VERDE), ("Duplicar tablero activo", "DuplicarTablero", VERDE),
               ("Duplicar circuito (fila activa)", "DuplicarCircuito", AZUL), ("Mover circuito (fila activa)", "MoverCircuito", AZUL), ("Actualizar proyecto y tablas", "ActualizarProyecto", VERDE), ("Importar / actualizar desde Revit (archivo)", "ImportarArchivoRevit", AZUL),
               ("Importar lo pegado en REVIT", "ImportarPegado", AZUL), ("Autobalancear tablero activo", "Autobalancear", AZUL), ("Autobalancear todos", "AutobalancearTodos", AZUL),
               ("Generar circuitos de prevista", "GenerarPrevista", AZUL), ("Exportar PDF (cuadros, resumen y memorias)", "ExportarPDF", GRIS), ("Exportar unifilar DXF", "ExportarDXF", GRIS),
               ("Exportar tablas para Revit (.xlsx)", "ExportarTablasRevit", GRIS), ("Solo tabla resumen (sí/no)", "SoloTablaResumen", GRIS), ("Eliminar tablero activo", "EliminarTablero", ROJO),
               ("Modo administrador", "ModoAdministrador", GRIS), ("Salir de administrador", "SalirAdministrador", GRIS), ("Cambiar contraseña", "CambiarClave", GRIS), ("Recalcular todo", "RecalcularTodo", GRIS)]
    top0 = ini.Range("B17").Top
    for i, (t, m, c) in enumerate(botones):
        boton(ini, ini.Range("B1").Left + (i % 4) * 215, top0 + (i // 4) * 42, 205, 34, t, m, c)
    for nm, lista in [("PROYECTO", [("Nuevo tablero", "NuevoTablero"), ("Crear varios", "CrearVarios"), ("Actualizar proyecto y tablas", "ActualizarProyecto"), ("Importar desde Revit", "ImportarArchivoRevit")]),
                      ("REVIT", [("Importar archivo", "ImportarArchivoRevit"), ("Importar lo pegado (desde B5)", "ImportarPegado")]),
                      ("PREVISTA", [("Generar circuitos de prevista", "GenerarPrevista")]),
                      ("TABLA RESUMEN", [("Actualizar tablas", "ActualizarProyecto"), ("Solo tabla resumen (sí/no)", "SoloTablaResumen"), ("Exportar tablas para Revit", "ExportarTablasRevit")]),
                      ("RESUMEN VERTICAL", [("Actualizar tablas", "ActualizarProyecto")]), ("RESUMEN DU", [("Actualizar tablas", "ActualizarProyecto")]),
                      ("MC_MACHOTE", [("Autobalancear", "Autobalancear"), ("Actualizar proyecto", "ActualizarProyecto"), ("+ Tablero derivado", "NuevoDerivado"), ("Duplicar tablero", "DuplicarTablero"),
                                      ("Duplicar circuito (fila activa)", "DuplicarCircuito"), ("Mover circuito (fila activa)", "MoverCircuito"), ("Ir a PROYECTO", "IrAProyecto"), ("Eliminar tablero", "EliminarTablero")])]:
        ws = wb.Worksheets(nm)
        if nm == "MC_MACHOTE":
            x0, y0 = ws.Range("T3").Left, ws.Range("T3").Top
            for i, (t, m) in enumerate(lista): boton(ws, x0 + (i % 2) * 190, y0 + (i // 2) * 32, 180, 27, t, m, ROJO if "Eliminar" in t else VERDE)
        elif nm == "TABLA RESUMEN":
            x0, y0 = ws.Range("AH2").Left, ws.Range("AH2").Top
            for i, (t, m) in enumerate(lista): boton(ws, x0 + i * 260, y0, 250, 40, t, m)
        else:
            col = {"PROYECTO": "H3", "REVIT": "L1", "PREVISTA": "L1", "RESUMEN VERTICAL": "F1", "RESUMEN DU": "F1"}[nm]
            x0, y0 = ws.Range(col).Left, ws.Range(col).Top
            for i, (t, m) in enumerate(lista): boton(ws, x0 + i * 170, y0, 160, 30, t, m)
    for ws in wb.Worksheets:
        if ws.Name.startswith("CAT_") or ws.Name in ("MC_MACHOTE", "VISTA_MACHOTE"): ws.Visible = 2   # xlSheetVeryHidden
    ini.Activate(); ini.Range("A1").Select()
    if os.path.exists(OUT): os.remove(OUT)
    wb.SaveAs(OUT, 52)
    print("guardado", OUT)
    wb.Close(False)
finally:
    vig.stop = True
    try: xl.Quit()
    except Exception: pass
    import subprocess; time.sleep(1); subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
print("dialogos:", LOG)
