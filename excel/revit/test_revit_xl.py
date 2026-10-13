import sys, os, time, subprocess, json
sys.path.insert(0, os.path.join(os.environ["TEMP"], "anacar", "xlsm"))
from xlcom import abrir_excel, donde_error, LOG
XLSM, FORM, CSV = [os.path.abspath(x) for x in sys.argv[1:4]]
def estado(wb):
    out = {}
    for s in wb.Worksheets:
        if s.Name.startswith("MC "):
            n = sum(1 for r in range(25, 125) if s.Range("E%d" % r).Formula or s.Range("G%d" % r).Formula)
            kva = sum((s.Range("G%d" % r).Value or 0) for r in range(25, 125) if not s.Range("S%d" % r).Value and isinstance(s.Range("G%d" % r).Value, float))
            out[s.Name[3:]] = (n, round(kva, 2), s.Range("C10").Value, s.Range("C12").Value, sum(1 for r in range(25, 125) if s.Range("S%d" % r).Value))
    return out
for etapa, archivos in [("formateado", [FORM, FORM]), ("csv directo", [CSV])]:
    xl, pid, vig = abrir_excel()
    try:
        wb = xl.Workbooks.Open(XLSM)
        for i, f in enumerate(archivos):
            t = time.time()
            try: xl.Run("ModRevit.ImportarArchivoRevit", f)
            except Exception as e: print("ERROR", e, donde_error(xl))
            rv = wb.Worksheets("REVIT")
            conc = [[rv.Cells(r, c).Text for c in range(12, 19)] for r in range(7, 26) if rv.Cells(r, 12).Text]
            print("==", etapa, "importación", i + 1, round(time.time() - t, 1), "s")
            for c in conc: print("  ", " | ".join(c))
            print("  tableros (circuitos, kVA propios, alimentado desde, long., enlaces):", json.dumps(estado(wb), ensure_ascii=False))
        wb.Close(False)
    finally:
        vig.stop = True
        try: xl.Quit()
        except Exception: pass
        time.sleep(1); subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
print("dialogos:"); [print("  ", x[:400].replace("\r\n", " / ")) for x in LOG]
