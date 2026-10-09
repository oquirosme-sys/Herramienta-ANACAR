Attribute VB_Name = "ModRevit"
' Importar / actualizar desde la tabla de circuitos de Revit (Electrical Circuit Schedule):
' .txt/.csv exportado, el Excel "Circuitos revit-excel" o filas pegadas en la hoja REVIT.
' Tableros nuevos: se crean con sus circuitos. Tableros existentes: se ACTUALIZAN carga, longitud y nombre
' (por numero de circuito de Revit, posicion, cambio pendiente o nombre) conservando los demas ajustes.
Option Explicit
Option Compare Text

Private Function LeerTexto(ruta As String) As String
    Dim st As Object, t As String
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open: st.LoadFromFile ruta
    t = st.ReadText: st.Close
    If InStr(t, ChrW(&HFFFD)) > 0 Then
        st.Charset = "windows-1252": st.Open: st.LoadFromFile ruta
        t = st.ReadText: st.Close
    End If
    LeerTexto = t
End Function

Private Function TextoAFilas(txt As String) As Variant
    Dim lineas As Variant, sep As String, i As Long, n As Long, a As Variant, j As Long, maxc As Long, filas() As Variant, nPc As Long
    txt = Replace(txt, vbCrLf, vbLf): txt = Replace(txt, vbCr, vbLf)
    lineas = Split(txt, vbLf)
    For i = 0 To UBound(lineas)
        If InStr(lineas(i), vbTab) > 0 Then sep = vbTab: Exit For
        If InStr(lineas(i), ";") > 0 Then nPc = nPc + 1
    Next i
    If sep = "" Then sep = IIf(nPc > (UBound(lineas) + 1) / 2, ";", ",")
    ReDim filas(0 To UBound(lineas))
    For i = 0 To UBound(lineas)
        If Trim$(lineas(i)) <> "" Then
            If sep = "," Then a = PartirCSV(CStr(lineas(i))) Else a = Split(lineas(i), sep)
            For j = 0 To UBound(a)
                a(j) = Trim$(a(j))
                If Left$(a(j), 1) = """" And Right$(a(j), 1) = """" And Len(a(j)) >= 2 Then a(j) = Mid$(a(j), 2, Len(a(j)) - 2)
            Next j
            filas(n) = a: n = n + 1
            If UBound(a) + 1 > maxc Then maxc = UBound(a) + 1
        End If
    Next i
    Dim m() As Variant
    If n = 0 Then TextoAFilas = Empty: Exit Function
    ReDim m(1 To n, 1 To maxc)
    For i = 0 To n - 1
        For j = 0 To UBound(filas(i))
            m(i + 1, j + 1) = filas(i)(j)
        Next j
    Next i
    TextoAFilas = m
End Function

Private Function PartirCSV(l As String) As Variant
    Dim out() As String, k As Long, cur As String, q As Boolean, i As Long, ch As String
    ReDim out(0 To 0)
    For i = 1 To Len(l)
        ch = Mid$(l, i, 1)
        If ch = """" Then
            q = Not q
        ElseIf ch = "," And Not q Then
            out(k) = cur: cur = "": k = k + 1: ReDim Preserve out(0 To k)
        Else
            cur = cur & ch
        End If
    Next i
    out(k) = cur
    PartirCSV = out
End Function

Public Sub ImportarArchivoRevit(Optional ruta As String = "")
    Dim f As Variant, wb As Workbook, ws As Worksheet, datos As Variant, hoja As Worksheet, r As Long, c As Long, encontrada As Boolean
    f = ruta
    If f = "" Then f = Application.GetOpenFilename("Tabla de circuitos (*.txt;*.csv;*.tsv;*.xlsx;*.xlsm;*.xls),*.txt;*.csv;*.tsv;*.xlsx;*.xlsm;*.xls", , "Tabla de circuitos de Revit")
    If f = False Then Exit Sub
    If LCase$(f) Like "*.xls*" Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Set wb = Workbooks.Open(Filename:=f, ReadOnly:=True, UpdateLinks:=0)
        Application.EnableEvents = True
        For Each ws In wb.Worksheets
            If ws.Visible = xlSheetVisible And Not encontrada Then
                For r = 1 To 30
                    For c = 1 To 15
                        If CStr(ws.Cells(r, c).Value) Like "*panel*" Then encontrada = True: Set hoja = ws: Exit For
                    Next c
                    If encontrada Then Exit For
                Next r
            End If
        Next ws
        If hoja Is Nothing Then Set hoja = wb.Worksheets(1)
        datos = hoja.UsedRange.Value
        wb.Close SaveChanges:=False
        Application.ScreenUpdating = True
    Else
        datos = TextoAFilas(LeerTexto(CStr(f)))
    End If
    If IsEmpty(datos) Then MsgBox "El archivo esta vacio.", vbExclamation: Exit Sub
    Procesar datos, CStr(f)
End Sub

Public Sub ImportarPegado()
    Dim ws As Worksheet, rg As Range
    Set ws = ThisWorkbook.Worksheets("REVIT")
    Set rg = ws.Range("B5").CurrentRegion
    If Application.WorksheetFunction.CountA(rg) = 0 Then MsgBox "Pegue la tabla de circuitos (con encabezados) desde la celda B5 de la hoja REVIT.", vbInformation: Exit Sub
    Procesar rg.Value, "hoja REVIT"
End Sub

Private Function Num(v As Variant, ByRef pies As Boolean) As Variant
    Dim s As String, i As Long, ch As String, t As String, p As Long, pulg As Double
    pies = False
    If IsEmpty(v) Then Num = Empty: Exit Function
    If VarType(v) = vbDouble Or VarType(v) = vbInteger Or VarType(v) = vbLong Or VarType(v) = vbCurrency Then Num = CDbl(v): Exit Function
    s = Trim$(CStr(v))
    If s = "" Then Num = Empty: Exit Function
    p = InStr(s, "'")
    If p > 0 Then
        pies = True
        t = Replace(Trim$(Left$(s, p - 1)), ",", ".")
        pulg = 0
        If InStr(p, s, """") > 0 Then pulg = Val(Replace(Replace(Replace(Mid$(s, p + 1), """", ""), "-", ""), ",", "."))
        Num = Val(t) + pulg / 12
        Exit Function
    End If
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[0-9.,-]" Then
            t = t & ch
        ElseIf t <> "" Then
            Exit For
        End If
    Next i
    If t = "" Then Num = Empty: Exit Function
    If InStr(t, ",") > 0 And InStr(t, ".") > 0 Then
        If InStrRev(t, ",") > InStrRev(t, ".") Then t = Replace(Replace(t, ".", ""), ",", ".") Else t = Replace(t, ",", "")
    ElseIf InStr(t, ",") > 0 Then
        t = Replace(t, ",", ".")
    End If
    Num = Val(t)
    If (s Like "* VA*" Or s Like "*[0-9]VA*") And Not (s Like "*kVA*") Then Num = Num / 1000
End Function

Private Function NormPanel(s As String) As String
    s = UCase$(Trim$(s))
    If Left$(s, 8) = "TABLERO " Then s = Mid$(s, 9)
    NormPanel = Replace(s, " ", "")
End Function

Private Function BuscarTablero(panel As String) As Worksheet
    Dim o As Variant, ws As Worksheet
    For Each o In Tableros()
        Set ws = o
        If NormPanel(IdDe(ws)) = NormPanel(panel) Or NormPanel(CStr(ws.Range("C4").Value) & " " & IdDe(ws)) = NormPanel(panel) Then Set BuscarTablero = ws: Exit Function
    Next o
End Function

' ---- detalle de carga segun el nombre (mismas claves que la herramienta web)
Private Function Clave(nm As String) As String
    Dim k As Variant, g As Variant
    k = Array("ilum,lumin,luz,luces,lamp|ilum", "ups|ups", "toma,tc ,recept,enchuf,cortiner|toma", "ascensor,elevador|elevador", "microondas,horno|horno", "cocina,estufa,plantilla|cocina", _
              "lavaplato,lavavajilla|lavaplatos", "facp,incendio|incendio", "acceso|acceso", "seguridad,cctv|seguridad", "rack, ti ,datos|rack", "extract|extractor", "inyect|inyector", "bomba|bomba", _
              "calentador,tanque agua,termo|calentador", "aire,a/c,ac ,minisplit,split,fan coil,fancoil,condensad,cu-,ah-,ms-,unidad|clima", "tablero,panel|tablero", "secaman|secamanos", "cargador|cargador", "porton|porton", "refri|refrigerador")
    For Each g In k
        If HayPalabra(" " & nm & " ", CStr(Split(g, "|")(0))) Then Clave = Split(g, "|")(1): Exit Function
    Next g
End Function

Private Function HayPalabra(d As String, lista As String) As Boolean
    Dim p As Variant
    For Each p In Split(lista, ",")
        If InStr(1, d, p, vbTextCompare) > 0 Then HayPalabra = True: Exit Function
    Next p
End Function

Private Function SinAc(s As String) As String
    Dim a As Variant, b As Variant, i As Long
    s = LCase$(s)
    a = Array("á", "é", "í", "ó", "ú", "ü", "ñ"): b = Array("a", "e", "i", "o", "u", "u", "n")
    For i = 0 To UBound(a): s = Replace(s, a(i), b(i)): Next i
    SinAc = s
End Function

Public Function DetalleRevit(nombre As String, voltaje As Double, polos As Long) As Variant
    Dim t As Variant, i As Long, s As Long, bestS As Long, best As Long, nm As String, dn As String, cl As String, w As Variant, hayV As Boolean
    t = ThisWorkbook.Names("T_DET").RefersToRange.Value
    nm = SinAc(nombre): cl = Clave(nm)
    If polos < 1 Then polos = 1
    For i = 1 To UBound(t, 1)
        If NumV(t(i, 4)) = voltaje And NumV(t(i, 5)) = polos Then hayV = True: Exit For
    Next i
    bestS = -1
    For i = 1 To UBound(t, 1)
        If NumV(t(i, 5)) = polos And (Not hayV Or NumV(t(i, 4)) = voltaje) Then
            dn = SinAc(CStr(t(i, 3))): s = 0
            If cl <> "" Then
                If Clave(dn) = cl Then s = s + 5
            End If
            For Each w In Split(Replace(Replace(Replace(nm, "-", " "), "_", " "), "/", " "), " ")
                If Len(w) > 3 Then
                    If InStr(1, dn, w) > 0 Then s = s + 2
                End If
            Next w
            If NumV(t(i, 4)) = voltaje Then s = s + 1
            If s > bestS Then bestS = s: best = i
        End If
    Next i
    If bestS < 2 Then
        For i = 1 To UBound(t, 1)
            If NumV(t(i, 5)) = polos And (Not hayV Or NumV(t(i, 4)) = voltaje) Then
                dn = Trim$(SinAc(CStr(t(i, 3))))
                If (polos = 1 And dn = "tomas") Or (polos = 3 And dn = "equipos mecanicos") Or (polos = 2 And (dn = "equipos" Or dn = "tomas 208v 1f")) Then best = i: Exit For
            End If
        Next i
    End If
    If best = 0 Then DetalleRevit = "" Else DetalleRevit = t(best, 1)
End Function

' ---- procesamiento
Private Sub Procesar(datos As Variant, origen As String)
    Dim ini As Long, cV As Long, cN As Long, cK As Long, cP As Long, cC As Long, cPo As Long, cL As Long, r As Long, c As Long, txt As String
    Dim paneles As Object, panel As String, k As Variant, it As Variant, enPies As Boolean, pies As Boolean
    ' encabezados
    For r = 1 To Application.WorksheetFunction.Min(30, UBound(datos, 1))
        cV = 0: cN = 0: cK = 0: cP = 0: cC = 0: cPo = 0: cL = 0
        For c = 1 To UBound(datos, 2)
            txt = CStr(datos(r, c))
            If cV = 0 And txt Like "*volt*" Then
                cV = c
            ElseIf cN = 0 And (txt Like "*load name*" Or txt Like "*nombre*" Or txt Like "*descrip*") Then
                cN = c
            ElseIf cK = 0 And (txt Like "*true load*" Or txt Like "*apparent load*" Or txt Like "*carga*" Or txt Like "*load*" Or txt Like "*kva*" Or txt Like "*potencia*") Then
                cK = c
            ElseIf cP = 0 And (txt Like "*panel*" Or txt Like "*tablero*") Then
                cP = c
            ElseIf cC = 0 And (txt Like "*circuit*" Or txt Like "*circuito*") Then
                cC = c
            ElseIf cPo = 0 And (txt Like "*pole*" Or txt Like "*polo*") Then
                cPo = c
            ElseIf cL = 0 And (txt Like "*length*" Or txt Like "*longitud*" Or txt Like "*largo*") Then
                cL = c
            End If
        Next c
        If cP > 0 And cC > 0 Then ini = r + 1: Exit For
    Next r
    If ini = 0 Then ini = 1: cV = 1: cK = 2: cP = 3: cC = 4: cPo = 5: cN = 6: cL = 7     ' orden de la macro ImportarCircuitosRevit
    Set paneles = CreateObject("Scripting.Dictionary")
    For r = ini To UBound(datos, 1)
        panel = ""
        If cP <= UBound(datos, 2) Then panel = Reparar(Trim$(CStr(datos(r, cP))))
        If panel <> "" And Not (panel Like "*electrical*" Or panel Like "*schedule*" Or panel = "panel") Then
            Dim cir As Object, np As Long, v As Double, lon As Variant
            Set cir = CreateObject("Scripting.Dictionary")
            np = NumV(Num(Celda(datos, r, cPo), pies)): If np < 1 Then np = 1
            v = NumV(Num(Celda(datos, r, cV), pies))
            lon = Num(Celda(datos, r, cL), pies)
            If pies Then enPies = True: cir("pies") = True Else cir("pies") = False
            cir("nombre") = Reparar(Trim$(CStr(Celda(datos, r, cN))))
            cir("v") = v: cir("np") = np
            cir("kva") = NumV(Num(Celda(datos, r, cK), pies))
            cir("lon") = lon
            cir("polos") = PolosDe(Celda(datos, r, cC), np)
            cir("num") = Join(cir("polos"), ",")
            cir("det") = DetalleRevit(CStr(cir("nombre")), v, IIf(np = 2 And v >= 208, 2, np))
            If Not paneles.Exists(panel) Then paneles.Add panel, New Collection
            paneles(panel).Add cir
        End If
    Next r
    If paneles.Count = 0 Then MsgBox "No se encontraron circuitos. Verifique que la tabla tenga las columnas Panel y Circuit Number.", vbExclamation: Exit Sub
    Dim factor As Double, resp As VbMsgBoxResult, lista As String, existentes As Long
    factor = 1
    If enPies Then factor = 0.3048
    For Each k In paneles.Keys
        lista = lista & vbCrLf & "  " & k & " (" & paneles(k).Count & " circ.)" & IIf(BuscarTablero(CStr(k)) Is Nothing, " - nuevo", " - ACTUALIZAR")
        If Not BuscarTablero(CStr(k)) Is Nothing Then existentes = existentes + 1
    Next k
    resp = MsgBox("Origen: " & origen & vbCrLf & "Tableros encontrados:" & lista & vbCrLf & vbCrLf & _
                  "Los existentes se actualizan (carga, longitud y nombre) conservando sus demas datos; los nuevos se crean." & vbCrLf & _
                  IIf(enPies, "Longitudes en pies: se convierten a metros." & vbCrLf, "") & vbCrLf & "Continuar?", vbYesNo + vbQuestion, "Importar desde Revit")
    If resp <> vbYes Then Exit Sub
    Dim res As String, nNuevos As Long, nAct As Long, nCirc As Long, faltan As String, ws As Worksheet, sis As String, tipo As String, vmax As Double, tres As Boolean
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    For Each k In paneles.Keys
        Set ws = BuscarTablero(CStr(k))
        If ws Is Nothing Then
            vmax = 0: tres = False
            For Each it In paneles(k)
                If it("v") > vmax Then vmax = it("v")
                If it("np") = 3 Then tres = True
            Next it
            sis = IIf(vmax >= 440, "277/480", IIf(vmax = 240, "120/240", "120/208"))
            tipo = IIf(tres Or sis = "120/208", "3F", "1F")
            Set ws = CrearTablero(LimpiarId(CStr(k)), tipo, sis, True)
            If ws Is Nothing Then GoTo siguiente
            nNuevos = nNuevos + 1
            Application.EnableEvents = False
        Else
            nAct = nAct + 1
        End If
        Actualizar ws, paneles(k), factor, nCirc, faltan
siguiente:
    Next k
    Application.EnableEvents = True
    ActualizarProyecto
    Application.ScreenUpdating = True
    MsgBox "Tableros nuevos: " & nNuevos & vbCrLf & "Tableros actualizados: " & nAct & vbCrLf & "Circuitos nuevos o con cambios: " & nCirc & _
           IIf(faltan <> "", vbCrLf & vbCrLf & "Circuitos importados antes de Revit que ya no estan en la tabla (revise y borre si corresponde):" & faltan, ""), vbInformation, "Importar desde Revit"
End Sub

Private Function Celda(d As Variant, r As Long, c As Long) As Variant
    If c < 1 Or c > UBound(d, 2) Then Celda = Empty Else Celda = d(r, c)
End Function

Private Function LimpiarId(s As String) As String
    Dim ch As Variant
    s = Trim$(s)
    If UCase$(Left$(s, 8)) = "TABLERO " Then s = Trim$(Mid$(s, 9))
    For Each ch In Array("[", "]", ":", "*", "?", "/", "\", "'")
        s = Replace(s, ch, "-")
    Next ch
    LimpiarId = Left$(s, 27)
End Function

Private Function PolosDe(v As Variant, n As Long) As Variant
    Dim s As String, a() As String, k As Long, i As Long, cur As String, ch As String
    ReDim a(1 To 6)
    If VarType(v) = vbDouble Then
        k = 1: a(1) = CStr(Int(v))     ' "20,22" llega como 20.22: se toma el primero y se completan los demas
    Else
        s = CStr(v) & " "
        For i = 1 To Len(s)
            ch = Mid$(s, i, 1)
            If ch Like "#" Then
                cur = cur & ch
            ElseIf cur <> "" Then
                If k < 6 Then k = k + 1: a(k) = cur
                cur = ""
            End If
        Next i
    End If
    If k = 0 Then k = 1: a(1) = "1"
    If n < k Then n = k
    For i = k + 1 To n
        a(i) = CStr(Val(a(1)) + 2 * (i - 1))
    Next i
    Dim r() As String
    ReDim r(0 To n - 1)
    For i = 1 To n: r(i - 1) = a(i): Next i
    PolosDe = r
End Function

' repara textos UTF-8 leidos como ANSI ("HabitaciÃ³n" -> "Habitación")
Private Function Reparar(s As String) As String
    Dim st As Object, b() As Byte, i As Long, ok As Boolean
    Reparar = s
    If InStr(s, "Ã") = 0 And InStr(s, "Â") = 0 Then Exit Function
    On Error GoTo fin
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "windows-1252": st.Open: st.WriteText s
    st.Position = 0: st.Type = 1: b = st.Read: st.Close
    st.Type = 1: st.Open: st.Write b: st.Position = 0: st.Type = 2: st.Charset = "utf-8"
    Reparar = st.ReadText: st.Close
    If InStr(Reparar, ChrW(&HFFFD)) > 0 Then Reparar = s
fin:
End Function

Private Sub Actualizar(ws As Worksheet, cs As Collection, factor As Double, ByRef nCirc As Long, ByRef faltan As String)
    Dim pl As Variant
    Dim usados As Object, it As Variant, r As Long, fila As Long, p As Long, lon As Variant, cm As Worksheet, rr As Long, np As Long, libre As Boolean, oc As Variant, x As Variant
    Set usados = CreateObject("Scripting.Dictionary")
    Set cm = ThisWorkbook.Worksheets("CAMBIOS REVIT")
    For Each it In cs
        fila = 0: pl = it("polos")
        If IsEmpty(it("lon")) Then lon = Empty Else lon = Round(it("lon") * factor, 2)
        ' 2) por un cambio de posicion pendiente (balanceo) ya pasado a Revit
        If fila = 0 Then
            For rr = 2 To cm.Cells(cm.Rows.Count, 1).End(xlUp).Row
                If cm.Cells(rr, 2).Value = IdDe(ws) And CStr(cm.Cells(rr, 5).Value) = it("num") And UCase$(CStr(cm.Cells(rr, 7).Value)) <> "X" Then
                    p = Val(Split(CStr(cm.Cells(rr, 5).Value), ",")(0))
                    If p > 0 Then
                        If Not usados.Exists(FilaPos(p)) And TieneDatosR(ws, FilaPos(p)) Then fila = FilaPos(p): cm.Cells(rr, 7).Value = "X": Exit For
                    End If
                End If
            Next rr
        End If
        If fila = 0 Then
            ' 1) por numero de circuito de Revit o por la posicion
            For r = FILA1 To FILAN
                If Not usados.Exists(r) And Trim$(CStr(ws.Cells(r, 19).Value)) = "" And TieneDatosR(ws, r) Then
                    If CStr(ws.Cells(r, 21).Value) = it("num") Or (CStr(ws.Cells(r, 21).Value) = "" And Join(PolosTxt(ws, r), ",") = it("num")) Then fila = r: Exit For
                End If
            Next r
        End If
        ' 3) por el nombre
        If fila = 0 Then
            For r = FILA1 To FILAN
                If Not usados.Exists(r) And Trim$(CStr(ws.Cells(r, 19).Value)) = "" And TieneDatosR(ws, r) Then
                    If LCase$(Trim$(CStr(ws.Cells(r, 6).Value))) = LCase$(it("nombre")) And it("nombre") <> "" Then fila = r: Exit For
                End If
            Next r
        End If
        If fila > 0 Then
            If NumV(ws.Cells(fila, 7).Value) <> it("kva") Or CStr(ws.Cells(fila, 8).Value) <> CStr(lon) Or CStr(ws.Cells(fila, 6).Value) <> it("nombre") Then nCirc = nCirc + 1
        Else
            ' circuito nuevo: en su posicion si esta libre, si no en la primera libre
            np = UBound(pl) + 1
            oc = Ocupacion(ws): libre = True
            For Each x In pl
                If Val(x) < 1 Or Val(x) > 100 Then libre = False Else If oc(Val(x)) Then libre = False
            Next x
            If libre Then p = Val(pl(0)) Else p = PosicionLibre(ws, np)
            If p = 0 Then GoTo sig
            fila = FilaPos(p)
            LimpiarFila ws, fila
            ws.Cells(fila, 5).Value = it("det")
            If libre And np >= 2 Then ws.Cells(fila, 3).Value = Val(pl(1))
            If libre And np >= 3 Then ws.Cells(fila, 4).Value = Val(pl(2))
            If Not libre And np >= 2 Then ws.Cells(fila, 3).Value = p + 2
            If Not libre And np >= 3 Then ws.Cells(fila, 4).Value = p + 4
            nCirc = nCirc + 1
        End If
        usados(fila) = True
        ws.Cells(fila, 6).Value = it("nombre")
        ws.Cells(fila, 7).Value = it("kva")
        ws.Cells(fila, 8).Value = lon
        ws.Cells(fila, 21).Value = it("num")
        ws.Cells(fila, 22).Value = "revit"
sig:
    Next it
    For r = FILA1 To FILAN
        If Not usados.Exists(r) And CStr(ws.Cells(r, 22).Value) = "revit" And TieneDatosR(ws, r) Then faltan = faltan & vbCrLf & "  " & IdDe(ws) & " [" & ws.Cells(r, 2).Value & "] " & ws.Cells(r, 6).Value
    Next r
End Sub

Private Function TieneDatosR(ws As Worksheet, r As Long) As Boolean
    TieneDatosR = TieneDatos(ws, r)
End Function

Private Function PolosTxt(ws As Worksheet, r As Long) As Variant
    Dim a As Variant, s() As String, i As Long
    a = PolosFila(ws, r)
    ReDim s(0 To UBound(a) - 1)
    For i = 1 To UBound(a): s(i - 1) = CStr(a(i)): Next i
    PolosTxt = s
End Function
