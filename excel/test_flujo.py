"""Flujo completo: importar Revit (archivo de ejemplo), conectar tableros, prevista, balanceo, DXF, PDF y tablas para Revit."""
import sys, os, json, time, subprocess, shutil
from xlcom import abrir_excel, donde_error, LOG
XLSM = os.path.abspath(sys.argv[1]); REV = os.path.abspath(sys.argv[2]); OUTD = os.path.abspath(sys.argv[3]); os.makedirs(OUTD, exist_ok=True)
xl, pid, vig = abrir_excel(); r = {}
def run(m, *a):
    t = time.time()
    try: v = xl.Run(m, *a); print("ok", m, round(time.time() - t, 1), "s"); return v
    except Exception as ex: print("ERROR", m, ex, donde_error(xl)); r.setdefault("errores", []).append(m + ": " + str(ex))
try:
    wb = xl.Workbooks.Open(XLSM)
    run("ModRevit.ImportarArchivoRevit", REV)
    mcs = [s.Name for s in wb.Worksheets if s.Name.startswith("MC ")]
    r["tableros"] = mcs
    for n in mcs[:40]:
        ws = wb.Worksheets(n)
        r[n] = [ws.Range(a).Text for a in ("C5", "C6", "BX10", "BX14", "BX24", "BX63", "BX72", "BX97")]
    # actualizar desde Revit por segunda vez: no debe duplicar
    tot1 = sum(wb.Worksheets(n).Range("BX10").Value or 0 for n in mcs)
    run("ModRevit.ImportarArchivoRevit", REV)
    tot2 = sum(wb.Worksheets(n).Range("BX10").Value or 0 for n in mcs)
    r["reimport_igual"] = (round(tot1, 4), round(tot2, 4), len([s for s in wb.Worksheets if s.Name.startswith("MC ")]))
    # conectar el segundo tablero al primero
    if len(mcs) >= 2:
        h = wb.Worksheets(mcs[1]); h.Range("C10").Value = mcs[0][3:]
        xl.CalculateFull()
        p = wb.Worksheets(mcs[0])
        r["conexion"] = [(rr, p.Range("F%d" % rr).Text, p.Range("G%d" % rr).Value, p.Range("C%d" % rr).Value, p.Range("D%d" % rr).Value) for rr in range(25, 125) if p.Range("S%d" % rr).Value == mcs[1][3:]]
        r["hijo_v"] = (h.Range("BX47").Value, p.Range("BX49").Value)
    # prevista
    pv = wb.Worksheets("PREVISTA"); pv.Range("C5").Value = "Oficinas"; pv.Range("D5").Value = 250; pv.Range("I5").Value = mcs[0][3:]
    try:
        usos = [c.Value for c in wb.Worksheets("CAT_USOS").Range("B2:B20") if c.Value]; pv.Range("C5").Value = usos[0]
    except Exception as e: print(e)
    run("ModPrevista.GenerarPrevista")
    p = wb.Worksheets(mcs[0])
    r["prevista"] = [(p.Range("B%d" % rr).Value, p.Range("F%d" % rr).Text, p.Range("G%d" % rr).Value) for rr in range(25, 125) if str(p.Range("V%d" % rr).Value).startswith("prevista")]
    r["prevista_R5"] = pv.Range("R5").Value
    run("ModBalanceo.AutobalancearTodos")
    cm = wb.Worksheets("CAMBIOS REVIT"); r["cambios"] = [[cm.Cells(i, c).Text for c in range(1, 7)] for i in range(2, 8)]
    run("ModTableros.ActualizarProyecto")
    run("ModDXF.ExportarDXF", os.path.join(OUTD, "unifilar.dxf"))
    run("ModTableros.ExportarPDF", os.path.join(OUTD, "tableros.pdf"), True)
    run("ModTableros.ExportarTablasRevit", os.path.join(OUTD, "tablas.xlsx"))
    r["proyecto"] = [[wb.Worksheets("PROYECTO").Cells(i, c).Text for c in range(2, 15)] for i in range(17, 22)]
    xl.DisplayAlerts = False
    if os.path.exists(os.path.join(OUTD, "flujo.xlsm")): os.remove(os.path.join(OUTD, "flujo.xlsm"))
    wb.SaveAs(os.path.join(OUTD, "flujo.xlsm"), 52)
    wb.Close(False)
finally:
    vig.stop = True
    try: xl.Quit()
    except Exception: pass
    time.sleep(1); subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
r["dialogos"] = LOG
json.dump(r, open(os.path.join(OUTD, "flujo.json"), "w", encoding="utf8"), ensure_ascii=False, indent=1, default=str)
print(json.dumps({k: v for k, v in r.items() if not k.startswith("MC ")}, ensure_ascii=False, default=str)[:4000])
import pymupdf
d = pymupdf.open(os.path.join(OUTD, "tableros.pdf"))
print("PDF", len(d), [(round(p.rect.width), round(p.rect.height)) for p in d][:8])
for i in range(min(len(d), 6)): d[i].get_pixmap(dpi=70).save(os.path.join(OUTD, "p%d.png" % i))
