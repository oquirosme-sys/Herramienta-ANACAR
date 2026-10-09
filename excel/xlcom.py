"""Utilidades COM para Excel con vigilante de diálogos (errores de compilación VBA, MsgBox)."""
import threading, time, os
import win32com.client as w32, pythoncom, win32gui, win32process, win32con, win32api

LOG = []


def textos(hwnd):
    out = []
    def cb(h, _):
        t = win32gui.GetWindowText(h)
        if t: out.append(t)
        return True
    try: win32gui.EnumChildWindows(hwnd, cb, None)
    except Exception: pass
    return out


def boton_txt(h, txt):
    res = []
    def cb(c, _):
        if win32gui.GetWindowText(c).replace("&", "") == txt: res.append(c)
        return True
    try: win32gui.EnumChildWindows(h, cb, None)
    except Exception: pass
    return res[0] if res else 0


class Vigilante(threading.Thread):
    def __init__(self, pid, stream=None):
        super().__init__(daemon=True); self.pid = pid; self.stop = False; self.stream = stream
    def run(self):
        pythoncom.CoInitialize()
        xl = None
        if self.stream is not None:
            try: xl = w32.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(self.stream, pythoncom.IID_IDispatch))
            except Exception as e: print("marshal", e)
        while not self.stop:
            def cb(h, _):
                try:
                    if win32gui.GetClassName(h) == "#32770" and win32gui.IsWindowVisible(h):
                        _, p = win32process.GetWindowThreadProcessId(h)
                        if p == self.pid:
                            msg = win32gui.GetWindowText(h) + " :: " + " | ".join(textos(h))
                            ok_ = boton_txt(h, "Aceptar") or boton_txt(h, "OK")
                            if ok_ and xl is not None and "error" in msg.lower():
                                win32api.PostMessage(ok_, win32con.BM_CLICK, 0, 0); time.sleep(1.5)
                                msg += " @ " + str(donde_error(xl))
                                try: xl.VBE.CommandBars.FindControl(Id=228).Execute()
                                except Exception as e: msg += " (reset: %s)" % e
                                LOG.append(msg); print("[DIALOGO]", msg, flush=True)
                                return True
                            dbg = boton_txt(h, "Debug") or boton_txt(h, "Depurar")
                            if dbg and xl is not None:
                                win32api.PostMessage(dbg, win32con.BM_CLICK, 0, 0); time.sleep(1.5)
                                msg += " @ " + str(donde_error(xl))
                                try: xl.VBE.CommandBars.FindControl(Id=228).Execute()
                                except Exception as e: msg += " (reset: %s)" % e
                                LOG.append(msg); print("[DIALOGO]", msg, flush=True)
                                return True
                            LOG.append(msg); print("[DIALOGO]", msg, flush=True)
                            # Sí / Aceptar
                            for t in ("Sí", "Yes", "Aceptar", "OK"):
                                b = boton_txt(h, t)
                                if b: win32api.PostMessage(b, win32con.BM_CLICK, 0, 0); break
                            else:
                                win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
                            time.sleep(0.8)
                except Exception as e:
                    pass
                return True
            try: win32gui.EnumWindows(cb, None)
            except Exception: pass
            time.sleep(0.7)


def abrir_excel():
    pythoncom.CoInitialize()
    xl = w32.DispatchEx("Excel.Application")
    xl.Visible = False; xl.DisplayAlerts = False; xl.ScreenUpdating = False
    try: xl.AutomationSecurity = 1
    except Exception: pass
    _, pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
    st = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, xl._oleobj_)
    v = Vigilante(pid, st); v.start()
    return xl, pid, v


def donde_error(xl):
    try:
        cp = xl.VBE.ActiveCodePane
        if cp is None: return None
        sel = cp.GetSelection(0, 0, 0, 0)
        cm = cp.CodeModule
        ln = sel[0]
        return "%s línea %d: %s" % (cm.Parent.Name, ln, cm.Lines(ln, 1).strip())
    except Exception as e:
        return "sin ubicación (%s)" % e
