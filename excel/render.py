"""Convierte cada ilustración de js/ilustraciones.js a PNG (480×360) con Chrome sin pantalla, para la hoja AYUDA del Excel."""
import sys, os, subprocess
ROOT, OUT = sys.argv[1], sys.argv[2]; os.makedirs(OUT, exist_ok=True)
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
claves = ["tablero", "centroCarga", "subestacion", "principal", "enchufable", "cajaMoldeada", "cajaMoldeadaE", "gfci", "spd", "transformador", "transferencia", "ups", "generador", "conductores", "tuberia", "barras", "factores", "cascada"]
js = "file:///" + os.path.join(ROOT, "js", "ilustraciones.js").replace("\\", "/")
for k in claves:
    il, o = (k, "{}") if k != "cajaMoldeadaE" else ("cajaMoldeada", "{electronico:true, etq:'Unidad electrónica LSIG'}")
    html = os.path.join(OUT, k + ".html")
    open(html, "w", encoding="utf8").write('<!doctype html><meta charset="utf-8"><style>html,body{margin:0;padding:0;overflow:hidden;background:#fff}svg{display:block;width:240px;height:180px}</style><body><div id="d"></div><script src="%s"></script><script>document.getElementById("d").innerHTML = Ilus.get("%s", %s);</script>' % (js, il, o))
    subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--allow-file-access-from-files", "--hide-scrollbars", "--force-device-scale-factor=2",
                    "--screenshot=" + os.path.join(OUT, k + ".png"), "--window-size=240,180", "file:///" + html.replace("\\", "/")], capture_output=True)
    print(k, os.path.getsize(os.path.join(OUT, k + ".png")))

# ---- simbolos.json: la biblioteca de símbolos de js/data/simbolos.js para la hoja CAT_SIMBOLOS del Excel ----
import re, html, json
vol = os.path.join(OUT, "simbolos_dump.html")
open(vol, "w", encoding="utf8").write('<!doctype html><meta charset="utf-8"><pre id="o"></pre><script src="file:///%s"></script><script>document.getElementById("o").textContent = "JSONINI" + JSON.stringify(window.SIMB) + "JSONFIN";</script>' % os.path.join(ROOT, "js", "data", "simbolos.js").replace("\\", "/"))
r = subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--allow-file-access-from-files", "--dump-dom", "file:///" + vol.replace("\\", "/")], capture_output=True, text=True, encoding="utf8")
m = re.search(r"JSONINI(.*?)JSONFIN", html.unescape(r.stdout), re.S)
open(os.path.join(OUT, "simbolos.json"), "w", encoding="utf8").write(m.group(1)); print("simbolos.json", len(m.group(1)), "bytes,", len(json.loads(m.group(1))), "simbolos")
