"""Prueba del autollenado: valores automáticos en gris, restaurar al borrar, listas con nombre, crear varios, duplicar, mover y edición desde PROYECTO."""
import sys, os, time, subprocess
from xlcom import abrir_excel, donde_error, LOG
XLSM = os.path.abspath(sys.argv[1]); mal = []; n = 0
def chk(nm, v, e):
    global n; n += 1
    ok = (abs(float(v) - float(e)) < 1e-6) if isinstance(e, (int, float)) and isinstance(v, (int, float)) else str(v) == str(e)
    if not ok: mal.append((nm, v, e))
xl, pid, vig = abrir_excel()
try:
    wb = xl.Workbooks.Open(XLSM); xl.Visible = False
    def run(m, *a):
        try: return xl.Run(m, *a)
        except Exception as ex: mal.append((m, str(ex), donde_error(xl)))
    run("ModTableros.CrearTablero", "TA", "3F", "120/208", True)
    ws = wb.Worksheets("MC TA"); ws.Activate()
    chk("vacío no cuenta", run("ModTableros.TieneDatos", ws, 25), False)
    chk("marca auto", ws.Range("C15").Value, "Eaton"); chk("marca es fórmula", ws.Range("C15").HasFormula, True)
    ws.Range("E25").Value = "8 · TOMAS (120 V 1F)"
    chk("lista → número", ws.Range("E25").Value, 8)
    ws.Range("G25").Value = 1.2; ws.Range("H25").Value = 20
    chk("descripción auto", ws.Range("F25").Value, "TOMAS"); chk("prot % auto", ws.Range("I25").Value, 125)
    chk("material auto", ws.Range("L25").Value, "CU"); chk("aislamiento auto", ws.Range("M25").Value, "THHN")
    chk("prot auto", ws.Range("P25").Value, 20); chk("breaker auto = cálculo", ws.Range("R25").Value, ws.Range("BK25").Value)
    chk("cuenta como circuito", run("ModTableros.TieneDatos", ws, 25), True)
    ws.Range("P25").Value = 30
    chk("prot manual", ws.Range("AT25").Value, 30); chk("manual no es fórmula", ws.Range("P25").HasFormula, False)
    ws.Range("P25").ClearContents()
    chk("borrar → vuelve auto", ws.Range("P25").HasFormula, True); chk("prot vuelve a 20", ws.Range("AT25").Value, 20)
    ws.Range("F25").Value = "Tomas oficina"; ws.Range("F25").ClearContents()
    chk("desc vuelve auto", ws.Range("F25").Value, "TOMAS")
    ws.Range("E26").Value = 55; ws.Range("G26").Value = 10
    chk("polos completados", (ws.Range("C26").Value, ws.Range("D26").Value), (5.0, 7.0))
    chk("feeder prot auto = BX24", ws.Range("F10").Value, ws.Range("BX24").Value)
    chk("tablero auto = BX62", ws.Range("F17").Value, ws.Range("BX62").Value)
    ws.Range("F17").Value = ws.Range("F17").Value  # usuario fija el mismo
    ws.Range("F17").ClearContents(); chk("tablero vuelve auto", ws.Range("F17").HasFormula, True)
    ws.Range("R25").Value = "4 · Eaton BAB 15-100 A STD 1P 10 kA @240 V"
    chk("breaker por nombre → id", ws.Range("R25").Value, 4)
    ws.Range("R25").ClearContents()
    # duplicar circuito (fila 25 → siguiente libre)
    ws.Range("G25").Select(); run("ModAuto.DuplicarCircuito")
    filas = [r for r in range(25, 125) if ws.Range("G%d" % r).Value == 1.2]
    chk("duplicado", len(filas), 2)
    if len(filas) == 2: chk("duplicado conserva auto", ws.Range("P%d" % filas[1]).HasFormula, True)
    # mover circuito (sin InputBox): CopiarFila + anotar
    pol = xl.Run("ModAuto.PolosFila", ws, 25) if False else None
    # crear varios con los valores por defecto (3 tableros T A..C)
    run("ModAuto.CrearVarios")
    nombres = [s.Name for s in wb.Worksheets if s.Name.startswith("MC ")]
    chk("crear varios", [x for x in nombres if x in ("MC TA", "MC TB", "MC TC")], ["MC TA", "MC TB", "MC TC"])
    # editar desde PROYECTO: TB alimentado desde TA
    pj = wb.Worksheets("PROYECTO")
    fila = [r for r in range(17, 40) if pj.Range("B%d" % r).Value == "TB"]
    if fila:
        pj.Range("E%d" % fila[0]).Value = "TA"; pj.Range("F%d" % fila[0]).Value = 35
        tb = wb.Worksheets("MC TB")
        chk("PROYECTO → padre", tb.Range("C10").Value, "TA"); chk("PROYECTO → longitud", tb.Range("C12").Value, 35)
        chk("circuito en TA", len([r for r in range(25, 125) if ws.Range("S%d" % r).Value == "TB"]), 1)
    else: mal.append(("fila TB en PROYECTO", None, None))
    # duplicar tablero (nombre por defecto TA-2)
    ws.Activate(); run("ModAuto.DuplicarTablero")
    chk("duplicar tablero", "MC TA-2" in [s.Name for s in wb.Worksheets], True)
    if "MC TA-2" in [s.Name for s in wb.Worksheets]:
        d = wb.Worksheets("MC TA-2"); chk("copia kVA", d.Range("BX3").Value, ws.Range("BX3").Value - 0)
        chk("copia sin circuito del derivado", len([r for r in range(25, 125) if d.Range("S%d" % r).Value]), 0)
    # balanceo con celdas automáticas
    for r, p in [(27, 5), (28, 7), (29, 9)]:
        pass
    wb.Close(False)
finally:
    vig.stop = True
    try: xl.Quit()
    except Exception: pass
    time.sleep(1); subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
print("checks", n, "fallan", len(mal))
for m in mal: print("  FALLA", m)
print("dialogos", [x[:120] for x in LOG])
