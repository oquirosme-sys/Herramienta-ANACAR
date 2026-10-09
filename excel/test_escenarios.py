import sys, os, json, time, subprocess
from xlcom import abrir_excel, donde_error, LOG
XLSM = os.path.abspath(sys.argv[1]); SC = json.load(open("escenarios.json", encoding="utf8")); EX = json.load(open("esperado_sc.json", encoding="utf8"))
fila = lambda p: 25 + (p - 1) // 2 if p % 2 else 75 + (p - 2) // 2
mal, n = [], 0
def chk(nm, v, e):
    global n; n += 1
    if isinstance(e, (int, float)) and e is not None and not isinstance(e, bool):
        try: ok = abs(float(v) - e) <= 2e-3 * max(1, abs(e))
        except Exception: ok = False
    else: ok = (v in (None, "") and e in (None, "")) or str(v).strip() == str(e).strip()
    if not ok: mal.append((nm, v, e))
xl, pid, vig = abrir_excel()
try:
    wb = xl.Workbooks.Open(XLSM)
    for s, e in zip(SC, EX):
        t = s["t"]; wb.Worksheets("PROYECTO").Range("F9").Value = s["P"]["ocupacion"]
        xl.Run("ModTableros.CrearTablero", t["nombre"], t["tipo"], t["sistema"], True)
        ws = wb.Worksheets("MC " + t["nombre"]); xl.EnableEvents = False
        al = t["alim"]
        v = {"C12": t["longitud"], "C13": t["reserva"], "C14": t["iccManual"], "F3": al["material"], "F4": al["aislamiento"], "F5": al["tuberia"], "F8": al["fp"]}
        if t.get("marca"): v["C15"] = t["marca"]
        if t.get("principal") == "zapatas": v["C16"] = "Zapatas"
        if t.get("clase") == "subestacion": v["C20"] = "Subestación"
        if al.get("tempAmb"): v["F11"] = al["tempAmb"]
        if al.get("tempBorne"): v["F12"] = al["tempBorne"]
        if al.get("unidad"): v["F14"] = al["unidad"]
        if al.get("rated100"): v["F15"] = "SI"
        for k, x in v.items(): ws.Range(k).Value = x
        for c in t["circuitos"]:
            r = fila(c["polos"][0])
            ws.Range("E%d" % r).Value = c["detalleId"]; ws.Range("F%d" % r).Value = c["descripcion"]; ws.Range("G%d" % r).Value = c["kva"]; ws.Range("H%d" % r).Value = c["longitud"]
            if len(c["polos"]) > 1: ws.Range("C%d" % r).Value = c["polos"][1]
            if len(c["polos"]) > 2: ws.Range("D%d" % r).Value = c["polos"][2]
        xl.EnableEvents = True; xl.CalculateFull()
        g = lambda a: ws.Range(a).Value
        nm = t["nombre"]
        chk(nm + " cat", g("BX63"), e["cat"]); chk(nm + " main", g("BX72") if g("BX72") != "ZAPATAS" else None, e["main"]); chk(nm + " sccr", g("BX74"), e["mainSccr"])
        chk(nm + " unidad", g("BX73"), e["unidad"]); chk(nm + " spd", g("BX79"), e["spd"]); chk(nm + " prot", g("BX24"), e["prot"]); chk(nm + " cal", g("BX99"), e["cal"])
        chk(nm + " cv", g("BX51"), e["cv"]); chk(nm + " kva", g("BX14"), e["kva"])
        for row in e["rows"]:
            r = fila(row[0])
            for k, col in zip(range(1, 6), ["AT", "BL", "BO", "BN", "BH"]): chk("%s c%d %s" % (nm, row[0], col), g("%s%d" % (col, r)), row[k])
        print(nm, "avisos:", g("BX97"), "| circ:", [g("BU%d" % fila(rr[0])) for rr in e["rows"]])
    wb.Close(False)
finally:
    vig.stop = True
    try: xl.Quit()
    except Exception: pass
    time.sleep(1); subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
print("checks", n, "fallan", len(mal))
for m in mal: print("  FALLA", m)
print("dialogos", LOG)
