import sys, os, time, subprocess
from xlcom import abrir_excel, donde_error, LOG
xl, pid, vig = abrir_excel()
try:
    wb = xl.Workbooks.Open(os.path.abspath("ANACAR.xlsm"))
    xl.Run("ModTableros.CrearTablero", "GEN-T", "3F", "120/208", True)
    ws = wb.Worksheets("MC GEN-T"); xl.EnableEvents = False
    for k, v in {"C14": 10, "C12": 8, "I14": "SI", "I15": "ATS", "I16": "Generador", "I18": 100, "I20": 25, "I22": "Retardada 4P", "I24": 12, "E25": 55, "G25": 40, "H25": 10, "C25": 3, "D25": 5}.items(): ws.Range(k).Value = v
    xl.EnableEvents = True; xl.CalculateFull()
    g = lambda a: ws.Range(a).Value
    print("modelo", g("BX115"), "amp", g("BX116"), "wcr", g("BX117"), "Icc ATS kA", round(g("BX113") or 0, 2), "Icc bornes kA", round((g("BX88") or 0) / 1000, 2), "dV%", round(g("BX93") or 0, 2), "prot", g("BX24"))
    print("largo total BX40 (esperado 20):", g("BX40"))
    print("panel K19:", ws.Range("L19").Text); print("avisos:", g("BX97"))
    print("BX118..132:", [g("BX%d" % i) for i in range(118, 133)])
    ws.Range("I10").Value = "SI"; ws.Range("I11").Value = 20; xl.CalculateFull()
    xl.Run("ModTableros.ActualizarProyecto")
    tr = wb.Worksheets("TABLA RESUMEN"); print("resumen AV..BJ:", [tr.Cells(11, c).Value for c in range(48, 63)])
    rv = wb.Worksheets("RESUMEN VERTICAL"); print("vertical:", [(rv.Cells(r, 2).Value, rv.Cells(r, 3).Value) for r in range(34, 49)])
    du = wb.Worksheets("RESUMEN DU"); print("du Q,R:", du.Range("Q4").Value, "|", du.Range("R4").Value)
    xl.Run("ModDXF.ExportarDXF", os.path.abspath("ats.dxf"))
    ws.Range("I10").Value = "NO"
    xl.Run("ModTableros.CrearTablero", "UPS-T", "3F", "120/208", True)
    w2 = wb.Worksheets("MC UPS-T"); xl.EnableEvents = False
    for k, v in {"I10": "SI", "I11": 8, "I13": "SI", "E25": 55, "G25": 40, "H25": 10, "C25": 3, "D25": 5}.items(): w2.Range(k).Value = v
    xl.EnableEvents = True; xl.CalculateFull()
    g2 = lambda a: w2.Range(a).Value
    print("UPS+MTS sin padre -> BX133/134:", g2("BX133"), g2("BX134"), "| 118..125:", [g2("BX%d" % i) for i in range(118, 126)], "| UPS:", g2("BX131"), g2("BX132"))
    w2.Range("C10").Value = "GEN-T"; xl.CalculateFull()
    w2.Range("I13").Value = "NO"; w2.Range("I13").Value = "SI"; xl.CalculateFull()
    print("UPS+MTS con padre -> I14..I17:", [g2("I%d" % i) for i in (14, 15, 16, 17)], "| 118..125:", [g2("BX%d" % i) for i in range(118, 126)])
    print("circuito de bypass en GEN-T:", [ws.Cells(r, 20).Value for r in range(25, 40) if ws.Cells(r, 20).Value])
    xl.Run("ModDXF.ExportarDXF", os.path.abspath("ats2.dxf"))
    ws.Range("I21").Value = 1; xl.Calculate(); print("con WCR 1 kA:", g("BX97"))
    wb.Close(False)
finally:
    vig.stop = True
    try: xl.Quit()
    except Exception: pass
    time.sleep(1); subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
print(LOG)
