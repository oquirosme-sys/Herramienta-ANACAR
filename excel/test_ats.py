import sys, os, time, subprocess
from xlcom import abrir_excel, donde_error, LOG
xl, pid, vig = abrir_excel()
try:
    wb = xl.Workbooks.Open(os.path.abspath("ANACAR.xlsm"))
    xl.Run("ModTableros.CrearTablero", "GEN-T", "3F", "120/208", True)
    ws = wb.Worksheets("MC GEN-T"); xl.EnableEvents = False
    for k, v in {"C14": 10, "I14": "SI", "I15": "ATS", "I16": "Generador", "I18": 100, "I20": 25, "I22": "Retardada 4P", "I23": 8, "I24": 12, "E25": 55, "G25": 40, "H25": 10, "C25": 3, "D25": 5}.items(): ws.Range(k).Value = v
    xl.EnableEvents = True; xl.CalculateFull()
    g = lambda a: ws.Range(a).Value
    print("modelo", g("BX115"), "amp", g("BX116"), "wcr", g("BX117"), "Icc ATS kA", round(g("BX113") or 0, 2), "Icc bornes kA", round((g("BX88") or 0) / 1000, 2), "dV%", round(g("BX93") or 0, 2), "prot", g("BX24"))
    print("largo total BX40 (esperado 20):", g("BX40"))
    print("panel K19:", ws.Range("L19").Text); print("avisos:", g("BX97"))
    ws.Range("I21").Value = 1; xl.Calculate(); print("con WCR 1 kA:", g("BX97"))
    wb.Close(False)
finally:
    vig.stop = True
    try: xl.Quit()
    except Exception: pass
    time.sleep(1); subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
print(LOG)
