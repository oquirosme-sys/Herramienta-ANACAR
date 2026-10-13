"""Crea la versión nueva de 'Circuitos revit-excel.xlsm' (a partir de una COPIA del original), la prueba con el CSV
y guarda: la plantilla limpia (DESTINO) y un ejemplo ya formateado (EJEMPLO) para probar la importación en ANACAR."""
import sys, os, json, time, shutil, subprocess
sys.path.insert(0, os.path.join(os.environ["TEMP"], "anacar", "xlsm"))
from xlcom import abrir_excel, donde_error, LOG

ORIG, CSV, BAS, DESTINO, EJEMPLO, VERDAD = [os.path.abspath(x) for x in sys.argv[1:7]]
tmp = os.path.join(os.path.dirname(BAS), "_trabajo.xlsm"); shutil.copyfile(ORIG, tmp)
xl, pid, vig = abrir_excel(); res = {}
try:
    wb = xl.Workbooks.Open(tmp)
    proj = wb.VBProject
    for c in list(proj.VBComponents):
        if c.Type == 1: proj.VBComponents.Remove(c)                       # Module1 (macro repetida)
        elif c.Type == 100 and c.CodeModule.CountOfLines: c.CodeModule.DeleteLines(1, c.CodeModule.CountOfLines)   # ThisWorkbook / Hoja1
    txt = open(BAS, encoding="utf8").read().replace("\r\n", "\n").replace("\n", "\r\n")
    b2 = os.path.join(os.path.dirname(BAS), "_ModFormato.bas"); open(b2, "w", encoding="cp1252", newline="").write(txt)
    proj.VBComponents.Import(b2)
    for c in proj.VBComponents:
        if c.Type == 100 and c.Name in (wb.CodeName,):
            c.CodeModule.AddFromString("Private Sub Workbook_Open()" + chr(13) + chr(10) + "    ModFormato.Limpiar" + chr(13) + chr(10) + "End Sub")
    ws = wb.Worksheets("Hoja1")
    for s in ws.Shapes:
        try: s.OnAction = "ImportarCircuitosRevit"
        except Exception as e: print("boton", e)
    ws.Range("F1").Value = "Paso 1: exporte de Revit 'Electrical Circuit Schedule' (.csv).  Paso 2: IMPORTAR DATOS y revise la hoja Revision.  Paso 3: importe este libro en ANACAR."
    ws.Range("F1").Font.Italic = True; ws.Range("F1").Font.Size = 9
    # guardar la plantilla limpia
    xl.Run("ModFormato.ImportarCircuitosRevit", CSV)
    sh = wb.Worksheets("Hoja1"); u = sh.Cells(sh.Rows.Count, 5).End(-4162).Row
    filas = [[sh.Cells(r, c).Value for c in range(3, 11)] for r in range(6, u + 1)]
    rv = wb.Worksheets("Revision")
    res["revision"] = [[rv.Cells(r, c).Text for c in range(1, 8)] for r in range(3, 30)]
    res["tabla"] = sh.ListObjects(1).Range.Address
    res["filas"] = filas
    if os.path.exists(EJEMPLO): os.remove(EJEMPLO)
    wb.SaveCopyAs(EJEMPLO)
    # plantilla limpia: tabla con una fila vacía
    sh.Range("C6:J%d" % max(6, u)).ClearContents(); sh.ListObjects(1).Resize(sh.Range("C5:J6")); sh.Range("F3").ClearContents()
    rv.Delete() if False else rv.Cells.Clear()
    rv.Range("A1").Value = "La revisión aparece aquí después de IMPORTAR DATOS."
    sh.Activate(); sh.Range("C6").Select()
    xl.Run("ModFormato.Limpiar")
    if os.path.exists(DESTINO): os.remove(DESTINO)
    wb.SaveAs(DESTINO, 52)
    wb.Close(False)
    # prueba: el ejemplo guardado CON datos se abre en blanco (con macros) pero los datos siguen en el archivo
    xl.EnableEvents = True
    w2 = xl.Workbooks.Open(EJEMPLO); w2.RunAutoMacros(1) if False else None
    res["abre_vacio"] = (w2.Worksheets("Hoja1").Range("E6").Value, w2.Worksheets("Hoja1").ListObjects(1).Range.Address)
    w2.Close(False)
finally:
    vig.stop = True
    try: xl.Quit()
    except Exception: pass
    time.sleep(1); subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
# comparación con la verdad (CSV leído por Python)
verdad = json.load(open(VERDAD, encoding="utf8"))
tot = {}
for f in res["filas"]:
    p = f[2]; t = tot.setdefault(p, {"n": 0, "kva": 0.0, "polos": 0}); t["n"] += 1; t["kva"] += f[1] or 0; t["polos"] += int(f[6] or 0)
mal = []
for p, v in verdad.items():
    t = tot.get(p)
    if not t or t["n"] != v["n"] or abs(t["kva"] - v["kva"]) > 0.01 or t["polos"] != v["polos"]: mal.append((p, v, t))
print("tableros", len(tot), "circuitos", sum(t["n"] for t in tot.values()), "diferencias", mal)
print("circuitos ejemplo:", [f[3] for f in res["filas"][:8]], "nombres:", [f[5] for f in res["filas"][:3]])
for r in res["revision"]:
    if any(r): print(" | ".join(r))
print("al abrir el ejemplo con datos:", res.get("abre_vacio"))
import openpyxl
print("datos guardados en el ejemplo (lo que lee ANACAR):", openpyxl.load_workbook(EJEMPLO, read_only=True)["Hoja1"]["E6"].value)
print("dialogos", LOG)
