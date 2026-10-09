"""Prueba del .xlsm: crea el tablero Machote con los datos del Excel ANACAR y compara con los resultados esperados;
luego prueba la cascada (tablero derivado), el balanceo y la selección automática. Guarda los resultados en JSON."""
import sys, os, time, json, time
from xlcom import abrir_excel, donde_error, LOG

XLSM, MACH, OUTJ = os.path.abspath(sys.argv[1]), sys.argv[2], sys.argv[3]
s = open(MACH, encoding="utf8").read(); D = json.loads(s[s.index("{"):s.rindex("}") + 1])
T, E = D["tablero"], D["esperado"]
fila = lambda p: 25 + (p - 1) // 2 if p % 2 else 75 + (p - 2) // 2
res = {"checks": [], "errores": []}
def chk(nombre, v, e, tol=1e-3):
    if isinstance(e, (int, float)) and not isinstance(e, bool):
        try: ok = abs(float(v) - e) <= tol * max(1, abs(e))
        except Exception: ok = False
    else: ok = str(v).strip() == str(e).strip()
    res["checks"].append((nombre, v, e, ok))

xl, pid, vig = abrir_excel()
try:
    wb = xl.Workbooks.Open(XLSM)
    xl.Calculation = -4105
    def run(m, *a):
        try: return xl.Run(m, *a)
        except Exception as ex:
            res["errores"].append("%s: %s | %s | %s" % (m, ex, donde_error(xl), LOG[-1:] )); print("ERROR", m, ex, donde_error(xl)); return None
    run("ModTableros.CrearTablero", "Machote", "3F", "120/208", True)
    ws = wb.Worksheets("MC Machote")
    xl.EnableEvents = False
    al = T["alim"]
    vals = {"C4": "TABLERO", "C11": T["conectadoA"], "C12": T["longitud"], "C13": T["reserva"], "C14": T["iccManual"], "F17": T["catalogoId"], "F18": T["supresorId"],
            "F3": al["material"], "F4": al["aislamiento"], "F5": al["tuberia"], "F6": al["paralelos"], "F7": (al["aumento"] - 1) * 100, "F8": al["fp"], "F9": al["mult"] * 100,
            "F16": al["breakerId"], "F11": al["tempAmb"], "F12": al["tempBorne"], "F13": al["agrupamiento"]}
    for k, v in vals.items(): ws.Range(k).Value = v
    for c in T["circuitos"]:
        r = fila(c["polos"][0])
        ws.Range("E%d" % r).Value = c["detalleId"]; ws.Range("G%d" % r).Value = c["kva"]; ws.Range("H%d" % r).Value = c["longitud"]
        ws.Range("I%d" % r).Value = c["mult"] * 100; ws.Range("L%d" % r).Value = c["material"]; ws.Range("M%d" % r).Value = c["aislamiento"]
        if c.get("paralelos"): ws.Range("N%d" % r).Value = c["paralelos"]
        ws.Range("O%d" % r).Value = round((c["aumento"] - 1) * 100, 6)
        if len(c["polos"]) > 1: ws.Range("C%d" % r).Value = c["polos"][1]
        if len(c["polos"]) > 2: ws.Range("D%d" % r).Value = c["polos"][2]
    xl.EnableEvents = True
    xl.CalculateFull()
    g = lambda a: ws.Range(a).Value
    for i, c in enumerate(T["circuitos"]):
        r, e = fila(c["polos"][0]), E["circuitos"][i]
        for k, col in [("U", "AJ"), ("V", "AK"), ("W", "AL"), ("AA", "AP"), ("AE", "AS"), ("AF", "AT"), ("AL", "AW"), ("AP", "AZ"), ("AR", "BA"), ("AS", "BB"), ("AT", "BC"), ("AU", "BD"), ("AV", "BE"), ("AY", "BH")]:
            v = g("%s%d" % (col, r))
            ee = e[k]
            if k in ("AL", "AP") and isinstance(ee, (int, float)): ee = str(int(ee)) if float(ee).is_integer() else str(ee)
            if k in ("AL", "AP"): v = str(v).replace(".0", "") if v is not None else v
            chk("c%d %s" % (i, k), v, ee)
    t = E["totales"]
    m = {"J116": "BX3", "U116": "BX4", "V116": "BX5", "W116": "BX6", "U118": "BX7", "W128": "BX8", "W129": "BX9", "W130": "BX10", "J134": "BX11", "L139": "BX14", "R139": "BX12",
         "AA139": "BX19", "AB139": "BX20", "AC139": "BX21", "AE139": "BX23", "AF139": "BX24", "AL139": "BX31", "AN139": "BX33", "AP139": "BX34", "AP140": "BX35", "AR139": "BX36",
         "AS139": "BX37", "AT139": "BX38", "AV139": "BX48", "AW139": "BX49", "AX139": "BX50", "AY139": "BX51", "AW140": "BX54"}
    for k, a in m.items():
        v = g(a); e = t[k]
        if k in ("AL139", "AN139", "AP139", "AP140"): v = str(v).replace(".0", ""); e = str(e)
        chk(k, v, e)
    chk("K131 (fd tomas)", g("J129"), t["K131"])
    chk("BA139 modelo principal", xl.Run("ModUDF.BREAKER_DATO", g("BX60"), g("BX71"), "modelo"), t["BA139"])
    chk("BJ11 tablero", g("BX63"), t["BJ11"]); chk("BJ16 SPD", g("BX79"), t["BJ16"])
    res["machote_bx"] = {"BX%d" % r: ws.Range("BX%d" % r).Value for r in range(2, 105)}
    res["machote_rows"] = [[ws.Range(c + str(fila(cc["polos"][0]))).Value for c in ["B", "W", "AT", "AW", "BI", "BJ", "BK", "BL", "BM", "BN", "BU"]] for cc in T["circuitos"]]
    # selección automática (sin tablero/principal/SPD manuales)
    for a in ("F16", "F17", "F18"): ws.Range(a).Value = None
    xl.CalculateFull()
    res["auto"] = {k: ws.Range(k).Value for k in ("BX60", "BX62", "BX63", "BX67", "BX68", "BX69", "BX70", "BX71", "BX72", "BX73", "BX74", "BX78", "BX79", "BX80", "BX97")}
    res["auto_rows"] = [[ws.Range(c + str(fila(cc["polos"][0]))).Value for c in ["B", "AT", "BI", "BJ", "BK", "BL", "BM", "BN", "BU"]] for cc in T["circuitos"]]
    # cascada
    run("ModTableros.CrearTablero", "TB", "3F", "120/208", True)
    h = wb.Worksheets("MC TB")
    xl.EnableEvents = False
    h.Range("C12").Value = 30; h.Range("C13").Value = 0; h.Range("F4").Value = "THHN"
    h.Range("E25").Value = 8; h.Range("G25").Value = 1.5; h.Range("H25").Value = 15
    h.Range("E75").Value = 8; h.Range("G75").Value = 1.5; h.Range("H75").Value = 15
    xl.EnableEvents = True
    h.Range("C10").Value = "Machote"           # dispara la conexión (evento)
    xl.CalculateFull()
    fr = [r for r in range(25, 125) if ws.Range("S%d" % r).Value == "TB"]
    res["cascada_fila"] = fr
    chk("Cascada: circuito creado en el padre", len(fr), 1)
    if fr:
        chk("Cascada: kVA del circuito = kVA del hijo", ws.Range("X%d" % fr[0]).Value, h.Range("BX10").Value)
        res["cascada_circ"] = [ws.Range(c + str(fr[0])).Formula for c in "BEFGHS"]
    chk("Cascada: inicio del hijo = bornes del padre", h.Range("BX47").Value, ws.Range("BX49").Value)
    chk("Cascada: caída hijo > padre", h.Range("BX51").Value > ws.Range("BX51").Value, True)
    chk("Cascada: Icc hijo < Icc padre", (h.Range("BX56").Value or 0) < (ws.Range("BX56").Value or 0) and (h.Range("BX56").Value or 0) > 0, True)
    chk("Cascada: nombre del padre", h.Range("BX98").Value, "TABLERO Machote")
    # renombrar
    h.Range("C3").Value = "TB2"
    xl.CalculateFull()
    names = [x.Name for x in wb.Worksheets]
    chk("Renombrar: hoja MC", "MC TB2" in names and "3F TB2" in names, True)
    chk("Renombrar: circuito del padre", ws.Range("S%d" % fr[0]).Value if fr else None, "TB2")
    # proyecto, tablas y nombres
    run("ModTableros.ActualizarProyecto")
    pj = wb.Worksheets("PROYECTO")
    res["proyecto"] = [[pj.Cells(r, c).Value for c in range(2, 15)] for r in range(17, 20)]
    tr = wb.Worksheets("TABLA RESUMEN")
    res["resumen"] = [[tr.Cells(r, c).Value for c in range(2, 48)] for r in range(11, 13)]
    res["nombres"] = [(n.Name, n.RefersTo) for n in wb.Names if n.Name.startswith(("TABLERO_", "TABLA_", "DATOS_", "RESUMEN_"))]
    vs = wb.Worksheets("3F Machote")
    res["vista"] = [[vs.Range(c + str(y)).Text for c in ["B", "D", "E", "G", "I", "J", "K", "N", "O", "R", "S", "T", "W", "AC", "AF", "AI"]] for y in (6, 9, 13, 14, 22, 63)]
    res["vista_ocultas"] = sum(1 for y in range(13, 113) if vs.Rows(y).Hidden)
    # balanceo
    antes = ws.Range("BX7").Value
    txt = run("ModBalanceo.BalancearTablero", ws, False)
    xl.CalculateFull()
    res["balanceo"] = {"antes": antes, "despues": ws.Range("BX7").Value, "txt": txt, "cambios": [[wb.Worksheets("CAMBIOS REVIT").Cells(r, c).Text for c in range(1, 7)] for r in range(2, 12)]}
    chk("Balanceo mejora", (ws.Range("BX7").Value or 0) < antes, True)
    chk("Balanceo conserva kVA", ws.Range("BX3").Value, t["J116"] + (h.Range("BX10").Value or 0))
    # errores de fórmula en las hojas
    errs = []
    for sh in (ws, h, vs, pj, tr):
        try:
            rg = sh.UsedRange.SpecialCells(-4123, 16)
            errs.append((sh.Name, rg.Address, rg.Cells(1, 1).Formula))
        except Exception: pass
    res["errores_formula"] = errs
    wb.Close(False)
finally:
    vig.stop = True
    try: xl.Quit()
    except Exception: pass
    import subprocess; time.sleep(1); subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
res["dialogos"] = LOG
json.dump(res, open(OUTJ, "w", encoding="utf8"), ensure_ascii=False, indent=1, default=str)
mal = [c for c in res["checks"] if not c[3]]
print("checks:", len(res["checks"]), "fallan:", len(mal))
for c in mal: print("  FALLA", c)
print("errores:", res["errores"]); print("errores de fórmula:", errs); print("diálogos:", LOG)
