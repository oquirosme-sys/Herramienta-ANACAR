import sys, os, time, subprocess
from xlcom import abrir_excel, LOG
xl, pid, vig = abrir_excel()
try:
    wb = xl.Workbooks.Open(os.path.abspath("ANACAR.xlsm"))
    xl.Run("ModTableros.CrearTablero", "CAL-1", "3F", "120/208", True)
    ws = wb.Worksheets("MC CAL-1"); xl.EnableEvents = False
    for k, v in {"C12": 10, "E25": 55, "G25": 3, "H25": 10, "C25": 3, "D25": 5}.items(): ws.Range(k).Value = v
    xl.EnableEvents = True; xl.CalculateFull()
    g = lambda a: ws.Range(a).Value
    print("auto: prot", g("BX24"), "cal BX31", g("BX31"), "neutro", g("BX33"), "tierra BX34", g("BX34"), "| req 250.122:", g("BX135"), "| avisos:", g("BX97"))
    xl.EnableEvents = False; ws.Range("F20").Value = "4"; ws.Range("F21").Value = "6"; ws.Range("F22").Value = "12"; xl.EnableEvents = True; xl.CalculateFull()
    print("manual: cal", g("BX31"), "neutro", g("BX33"), "tierra", g("BX34"), "| val", g("BX29"), "| avisos:", g("BX97"))
    print("resumen fases/neutro/tierra:", g("BX99"), g("BX100"), g("BX101"))
    wb.Close(False)
finally:
    vig.stop = True
    try: xl.Quit()
    except Exception: pass
    time.sleep(1); subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
print(LOG)
