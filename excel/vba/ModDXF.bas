Attribute VB_Name = "ModDXF"
' Diagrama unifilar en DXF (AutoCAD R12): acometida / transformador / UPS, barras de cada tablero,
' interruptores, alimentadores con su calibre, segunda acometida (ATS/MTS/IP) y tabla de datos junto a cada tablero.
Option Explicit

Private buf As String, blq As String
Private sk As Object, sca As Object, sal As Object, spx As Object, spy As Object, sat As Object
Private Const ESC As Double = 1.15          ' escala de los simbolos en el dibujo (los bloques traen la geometria de los DU)
Private Const DX As Double = 190      ' separacion horizontal entre tableros
Private Const DY As Double = 150      ' separacion vertical entre niveles
Private xs As Object, ys As Object, hoja As Object, hijos As Object, sigX As Double

Private Function T(v As Variant) As String
    ' texto DXF: caracteres fuera de ASCII como \U+XXXX
    Dim s As String, i As Long, c As Long, r As String
    s = CStr(v)
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)): If c < 0 Then c = c + 65536
        If c > 126 Then r = r & "\U+" & Right$("0000" & Hex$(c), 4) Else r = r & Mid$(s, i, 1)
    Next i
    T = r
End Function

Private Function F(x As Double) As String
    F = Replace(Format$(x, "0.0###"), ",", ".")
End Function

Private Sub Ln(capa As String, x1 As Double, y1 As Double, x2 As Double, y2 As Double, Optional color As Long = 256, Optional tipo As String = "")
    buf = buf & "0" & vbCrLf & "LINE" & vbCrLf & "8" & vbCrLf & capa & vbCrLf
    If tipo <> "" Then buf = buf & "6" & vbCrLf & tipo & vbCrLf
    If color <> 256 Then buf = buf & "62" & vbCrLf & color & vbCrLf
    buf = buf & "10" & vbCrLf & F(x1) & vbCrLf & "20" & vbCrLf & F(y1) & vbCrLf & "11" & vbCrLf & F(x2) & vbCrLf & "21" & vbCrLf & F(y2) & vbCrLf
End Sub

Private Sub Tx(capa As String, x As Double, y As Double, h As Double, s As Variant, Optional alin As Long = 0, Optional rot As Double = 0)
    If CStr(s) = "" Then Exit Sub
    buf = buf & "0" & vbCrLf & "TEXT" & vbCrLf & "8" & vbCrLf & capa & vbCrLf & "7" & vbCrLf & "SINERGIA" & vbCrLf & "10" & vbCrLf & F(x) & vbCrLf & "20" & vbCrLf & F(y) & vbCrLf & _
          "40" & vbCrLf & F(h) & vbCrLf & "1" & vbCrLf & T(s) & vbCrLf
    If rot <> 0 Then buf = buf & "50" & vbCrLf & F(rot) & vbCrLf
    If alin <> 0 Then buf = buf & "72" & vbCrLf & alin & vbCrLf & "11" & vbCrLf & F(x) & vbCrLf & "21" & vbCrLf & F(y) & vbCrLf
End Sub

Private Sub Circ(capa As String, x As Double, y As Double, r As Double)
    buf = buf & "0" & vbCrLf & "CIRCLE" & vbCrLf & "8" & vbCrLf & capa & vbCrLf & "10" & vbCrLf & F(x) & vbCrLf & "20" & vbCrLf & F(y) & vbCrLf & "40" & vbCrLf & F(r) & vbCrLf
End Sub

Private Sub Rect(capa As String, x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    Ln capa, x1, y1, x2, y1: Ln capa, x2, y1, x2, y2: Ln capa, x2, y2, x1, y2: Ln capa, x1, y2, x1, y1
End Sub

Private Function G(codigo As String, valor As Variant) As String
    G = codigo & vbCrLf & CStr(valor) & vbCrLf
End Function

Private Function N(v As Variant) As String
    N = Replace(Format$(CDbl(v), "0.0###"), ",", ".")
End Function

' lee la hoja CAT_SIMBOLOS y arma las definiciones de bloque (las mismas del diagrama de la herramienta en linea)
Private Sub CargarSimbolos()
    Dim ws As Worksheet, d As Variant, i As Long, nm As String, ult As String, nAtt As Object, b As String, u As Long, tp As String
    Set ws = ThisWorkbook.Worksheets("CAT_SIMBOLOS")
    u = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    d = ws.Range("A2:K" & u).Value
    Set sk = CreateObject("Scripting.Dictionary"): Set sca = CreateObject("Scripting.Dictionary"): Set sal = CreateObject("Scripting.Dictionary")
    Set spx = CreateObject("Scripting.Dictionary"): Set spy = CreateObject("Scripting.Dictionary"): Set sat = CreateObject("Scripting.Dictionary")
    Set nAtt = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(d, 1)
        nm = CStr(d(i, 1))
        sca(nm) = CStr(d(i, 2)): sal(nm) = CDbl(d(i, 4))
        If CStr(d(i, 5)) = "AD" Then nAtt(nm) = nAtt(nm) + 1
        If CStr(d(i, 5)) = "PT" Then spx(nm & "|" & d(i, 9)) = CDbl(d(i, 6)): spy(nm & "|" & d(i, 9)) = CDbl(d(i, 7))
    Next i
    blq = ""
    For i = 1 To UBound(d, 1)
        nm = CStr(d(i, 1)): tp = CStr(d(i, 5))
        If nm <> ult Then
            If ult <> "" Then blq = blq & G("0", "ENDBLK") & G("8", "0")
            blq = blq & G("0", "BLOCK") & G("8", "0") & G("2", nm) & G("70", IIf(nAtt.Exists(nm), 2, 0)) & G("10", 0) & G("20", 0) & G("30", 0) & G("3", nm)
            ult = nm
        End If
        Select Case tp
            Case "L": blq = blq & G("0", "LINE") & G("8", "0") & G("10", N(d(i, 6))) & G("20", N(d(i, 7))) & G("30", 0) & G("11", N(d(i, 8))) & G("21", N(d(i, 9))) & G("31", 0)
            Case "C": blq = blq & G("0", "CIRCLE") & G("8", "0") & G("10", N(d(i, 6))) & G("20", N(d(i, 7))) & G("30", 0) & G("40", N(d(i, 8)))
            Case "F"   ' punto relleno: polilinea cerrada de dos vertices con bulge 1 y ancho = radio
                blq = blq & G("0", "POLYLINE") & G("8", "0") & G("66", 1) & G("10", 0) & G("20", 0) & G("30", 0) & G("40", N(d(i, 8))) & G("41", N(d(i, 8))) & G("70", 1) & _
                      G("0", "VERTEX") & G("8", "0") & G("10", N(d(i, 6) - d(i, 8) / 2)) & G("20", N(d(i, 7))) & G("30", 0) & G("42", 1) & _
                      G("0", "VERTEX") & G("8", "0") & G("10", N(d(i, 6) + d(i, 8) / 2)) & G("20", N(d(i, 7))) & G("30", 0) & G("42", 1) & G("0", "SEQEND") & G("8", "0")
            Case "A": blq = blq & G("0", "ARC") & G("8", "0") & G("10", N(d(i, 6))) & G("20", N(d(i, 7))) & G("30", 0) & G("40", N(d(i, 8))) & G("50", N(d(i, 9))) & G("51", N(d(i, 10)))
            Case "T": blq = blq & G("0", "TEXT") & G("8", "0") & G("7", "SINERGIA") & G("10", N(d(i, 6))) & G("20", N(d(i, 7))) & G("30", 0) & G("40", N(d(i, 8))) & G("1", T(d(i, 9))) & G("72", 1) & G("73", 2) & G("11", N(d(i, 6))) & G("21", N(d(i, 7))) & G("31", 0)
            Case "AD": blq = blq & G("0", "ATTDEF") & G("8", "0") & G("10", N(d(i, 6))) & G("20", N(d(i, 7))) & G("30", 0) & G("40", N(d(i, 8))) & G("1", " ") & G("3", T(d(i, 10))) & G("2", d(i, 9)) & G("70", IIf(CStr(d(i, 11)) = "1", 0, 1)) & G("7", "SINERGIA")
        End Select
        If tp = "AD" Then
            If Not sat.Exists(nm) Then sat.Add nm, New Collection
            sat(nm).Add Array(CStr(d(i, 9)), CDbl(d(i, 6)), CDbl(d(i, 7)), CDbl(d(i, 8)), CStr(d(i, 11)) = "1")
        End If
    Next i
    If ult <> "" Then blq = blq & G("0", "ENDBLK") & G("8", "0")
End Sub

' inserta un simbolo (bloque) con sus atributos: vals = Array("TAG", "valor", "TAG2", "valor2", ...)
Private Sub Simbolo(nombre As String, x As Double, y As Double, Optional vals As Variant)
    Dim a As Variant, i As Long, v As String, j As Long, conAtt As Boolean
    conAtt = sat.Exists(nombre)
    buf = buf & G("0", "INSERT") & G("8", "E-" & sca(nombre)) & IIf(conAtt, G("66", 1), "") & G("2", nombre) & G("10", N(x)) & G("20", N(y)) & G("30", 0) & G("41", N(ESC)) & G("42", N(ESC)) & G("43", N(ESC)) & G("50", 0)
    If conAtt Then
        For Each a In sat(nombre)
            v = " "
            If Not IsMissing(vals) Then
                For j = 0 To UBound(vals) Step 2
                    If UCase$(CStr(vals(j))) = UCase$(CStr(a(0))) Then If CStr(vals(j + 1)) <> "" Then v = CStr(vals(j + 1))
                Next j
            End If
            buf = buf & G("0", "ATTRIB") & G("8", "E-" & sca(nombre)) & G("10", N(x + a(1) * ESC)) & G("20", N(y + a(2) * ESC)) & G("30", 0) & G("40", N(a(3) * ESC)) & G("1", T(v)) & G("7", "SINERGIA") & G("2", a(0)) & G("70", IIf(a(4), 0, 1))
        Next a
        buf = buf & G("0", "SEQEND") & G("8", "E-" & sca(nombre))
    End If
End Sub

Private Function AltoS(nombre As String) As Double
    AltoS = sal(nombre) * ESC
End Function

' interruptor (bloque de 1, 2 o 3 polos) con el origen en su terminal superior; devuelve el borde inferior
Private Function Breaker(x As Double, ysup As Double, polos As Long, amp As Variant, unidad As String, modelo As String, sccr As Variant, txt As String) As Double
    Dim np As Long
    np = polos: If np < 1 Then np = 3
    If np > 3 Then np = 3
    Simbolo "INTERRUPTOR_" & np & "P", x, ysup, Array("POLOS", np & "p", "AMPERIOS", IIf(CStr(amp) = "", "", amp & "A"), "UNIDAD", unidad, "MODELO", modelo, "SCCR", CStr(sccr))
    Tx "E-PROTECCION", x + 7, ysup - AltoS("INTERRUPTOR_3P") / 2, 2.5, txt
    Breaker = ysup - AltoS("INTERRUPTOR_3P")
End Function

' linea vertical de y0 (arriba) a y1 (abajo) salteando los huecos de los simbolos: hu(i,0)=borde superior, hu(i,1)=borde inferior (ordenados de arriba hacia abajo)
Private Sub LineaV(x As Double, y0 As Double, y1 As Double, hu() As Double, n As Long, capa As String)
    Dim a As Double, i As Long
    a = y0
    For i = 1 To n
        If hu(i, 0) < a Then Ln capa, x, a, x, hu(i, 0)
        If hu(i, 1) < a Then a = hu(i, 1)
    Next i
    If y1 < a Then Ln capa, x, a, x, y1
End Sub

Private Sub Posiciones(id As String, nivel As Long)
    Dim c As Collection, k As Variant, x0 As Double, x1 As Double, n As Long
    Set c = hijos(id)
    ys(id) = -nivel * DY
    If c.Count = 0 Then
        xs(id) = sigX: sigX = sigX + DX
    Else
        For Each k In c
            Posiciones CStr(k), nivel + 1
            n = n + 1
            If n = 1 Then x0 = xs(k)
            x1 = xs(k)
        Next k
        xs(id) = (x0 + x1) / 2
    End If
End Sub

Public Sub ExportarDXF(Optional destino As String = "")
    Dim ruta As Variant, o As Variant, ws As Worksheet, id As String, padre As String, k As Variant, raiz As Collection
    ActualizarProyecto
    ruta = destino
    If ruta = "" Then ruta = Application.GetSaveAsFilename(InitialFileName:=Replace(ThisWorkbook.Name, ".xlsm", "") & " - unifilar.dxf", FileFilter:="AutoCAD DXF (*.dxf), *.dxf")
    If ruta = False Then Exit Sub
    Set xs = CreateObject("Scripting.Dictionary"): Set ys = CreateObject("Scripting.Dictionary")
    Set hoja = CreateObject("Scripting.Dictionary"): Set hijos = CreateObject("Scripting.Dictionary")
    Set raiz = New Collection
    For Each o In Orden()
        Set ws = o: id = IdDe(ws)
        hoja.Add id, ws: hijos.Add id, New Collection
    Next o
    For Each k In hoja.Keys
        padre = Trim$(CStr(hoja(k).Range("C10").Value))
        If padre <> "" And hoja.Exists(padre) Then hijos(padre).Add CStr(k) Else raiz.Add CStr(k)
    Next k
    sigX = 0
    For Each k In raiz
        Posiciones CStr(k), 0
        sigX = sigX + DX * 0.3
    Next k
    buf = ""
    CargarSimbolos
    For Each k In hoja.Keys
        Dibujar CStr(k)
    Next k
    Tx "E-TEXTO", 0, 70, 6, "DIAGRAMA UNIFILAR - " & ThisWorkbook.Names("P_NOMBRE").RefersToRange.Value
    Tx "E-TEXTO", 0, 60, 3, "ANACAR " & Format$(Date, "yyyy-mm-dd") & " - corriente de cortocircuito por el metodo punto a punto; caidas de voltaje acumuladas."
    Escribir CStr(ruta)
    If destino = "" Then MsgBox "Diagrama unifilar exportado:" & vbCrLf & ruta, vbInformation
End Sub

Private Function PolosB(ws As Worksheet) As Long
    If NumV(ws.Range("C8").Value) = 3 Then
        PolosB = 3
    ElseIf NumV(ws.Range("C7").Value) >= 208 And NumV(ws.Range("C7").Value) <= 240 Then
        PolosB = 2
    Else
        PolosB = 1
    End If
End Function

Private Sub Dibujar(id As String)
    Dim ws As Worksheet, x As Double, y As Double, padre As String, c As Collection, k As Variant, x0 As Double, x1 As Double, yb As Double
    Dim dg As Double, filas As Variant, i As Long, hu() As Double, nh As Long, hb As Double, yBk As Double, yT As Double, yTop As Double, yBot As Double
    Dim sym As String, yA As Double, ya2 As Double, xa As Double, gx As Double, yBa As Double, yG As Double, ha As Double
    Dim und As String, mod_ As String, sccr As Variant
    Set ws = hoja(id): x = xs(id): y = ys(id)
    Set c = hijos(id)
    ' barra del tablero (se extiende a sus derivados)
    x0 = x - 25: x1 = x + 25
    For Each k In c
        If xs(k) - 5 < x0 Then x0 = xs(k) - 5
        If xs(k) + 5 > x1 Then x1 = xs(k) + 5
    Next k
    yb = y - 20
    Ln "E-TABLERO", x0, yb, x1, yb, 1
    Ln "E-TABLERO", x0, yb - 1, x1, yb - 1, 1
    Tx "E-TABLERO", x - 24, yb + 2, 4, ws.Range("BX2").Value
    ReDim hu(1 To 8, 0 To 1): nh = 0
    hb = AltoS("INTERRUPTOR_3P")
    und = IIf(OrDefault(ws.Range("F14").Value) = "TM", "TM", CStr(ws.Range("F14").Value))
    mod_ = CStr(ws.Range("BX72").Value): sccr = ws.Range("BX74").Value
    padre = Trim$(CStr(ws.Range("C10").Value))
    ' con ATS/MTS se deja espacio entre el equipo y la barra: el resto de la bajada sube lo mismo
    dg = IIf(ws.Range("I14").Value = "SI" And ws.Range("I15").Value <> "IP", 21, 0)
    ' ---- alimentacion normal: interruptor (bloque), transformador seco, UPS ----
    If padre <> "" And hoja.Exists(padre) Then
        yTop = ys(padre) - 21
        yBk = ys(padre) - 26
        Breaker x, yBk, PolosB(ws), ws.Range("BX24").Value, und, mod_, sccr, NumV(ws.Range("BX24").Value) & " A " & PolosB(ws) & "P " & und
        nh = nh + 1: hu(nh, 0) = yBk: hu(nh, 1) = yBk - hb
        Tx "E-ALIMENTADOR", x - 3, ys(padre) - 45, 2.2, ws.Range("BX99").Value & IIf(ws.Range("BX100").Value <> "", " + " & ws.Range("BX100").Value & "N", "") & IIf(ws.Range("BX101").Value <> "", " + " & ws.Range("BX101").Value & "T", "") & " " & ws.Range("F3").Value & " - " & Format$(NumV(ws.Range("C12").Value), "0.0") & " m", 0, 90
    Else
        yTop = y + 100 + dg
        Tx "E-TEXTO", x + 3, yTop + 2, 2.5, IIf(CStr(ws.Range("C11").Value) <> "", ws.Range("C11").Value, "ACOMETIDA")
    End If
    If ws.Range("I3").Value = "SI" Then
        yT = y + 84 + dg
        Simbolo "TRAFO_SECO", x, yT, Array("KVA", ws.Range("I4").Value & " kVA", "IMPEDANCIA", "Z " & ws.Range("I5").Value & " %", "VOLTAJES", ws.Range("I7").Value & " / " & ws.Range("C6").Value & " V")
        nh = nh + 1: hu(nh, 0) = yT + 3 * ESC: hu(nh, 1) = yT - AltoS("TRAFO_SECO") - 3 * ESC
        Tx "E-TRAFO", x + 13, yT - 6, 2.5, IIf(CStr(ws.Range("I8").Value) <> "", ws.Range("I8").Value & " ", "") & ws.Range("I4").Value & " kVA  Z " & ws.Range("I5").Value & " %"
    End If
    If ws.Range("I10").Value = "SI" Then
        Rect "E-EQUIPOS", x - 7, y + 44 + dg, x + 7, y + 56 + dg
        Tx "E-EQUIPOS", x, y + 48 + dg, 3, "UPS", 1
        Tx "E-EQUIPOS", x + 9, y + 48 + dg, 2.5, ws.Range("I11").Value & " kVA"
        nh = nh + 1: hu(nh, 0) = y + 56 + dg: hu(nh, 1) = y + 44 + dg
    End If
    If padre = "" And ws.Range("C16").Value <> "Zapatas" Then
        yBk = y + 34 + dg
        Breaker x, yBk, PolosB(ws), ws.Range("BX24").Value, und, mod_, sccr, NumV(ws.Range("BX24").Value) & " A " & PolosB(ws) & "P " & und
        nh = nh + 1: hu(nh, 0) = yBk: hu(nh, 1) = yBk - hb
    End If
    ' ---- segunda acometida: ATS / MTS pegado a la barra, interruptor y generador (o bypass) ----
    If ws.Range("I14").Value = "SI" Then
        sym = IIf(ws.Range("I15").Value = "MTS", "MTS", "ATS")
        If ws.Range("I15").Value = "IP" Then
            Rect "E-EQUIPOS", x - 7, yb + 4, x + 7, yb + 14
            Tx "E-EQUIPOS", x, yb + 7.5, 3, "IP", 1
            nh = nh + 1: hu(nh, 0) = yb + 14: hu(nh, 1) = yb + 4
            xa = x + 7: ya2 = yb + 9
        Else
            ha = AltoS(sym): yA = yb + 22 + ha   ' espacio entre el ATS/MTS y la barra para el cableado del tramo comun
            Simbolo sym, x, yA, Array("MODELO", ws.Range("BX115").Value, "AMPERIOS", ws.Range("BX116").Value, "WCR", ws.Range("BX117").Value, "TRANSICION", ws.Range("I22").Value)
            nh = nh + 1: hu(nh, 0) = yA: hu(nh, 1) = yA - ha
            xa = x + spx(sym & "|alterna") * ESC: ya2 = yA + spy(sym & "|alterna") * ESC
            Tx "E-EQUIPOS", x - 14, yA - 8, 2.4, ws.Range("I15").Value & IIf(CStr(ws.Range("BX115").Value) <> "", " " & ws.Range("BX115").Value & "  WCR " & ws.Range("BX117").Value & " kA", ""), 2
            If NumV(ws.Range("I23").Value) > 0 Then Tx "E-EQUIPOS", x - 14, yA - 12, 2, Format$(NumV(ws.Range("I23").Value), "0.0") & " m al tablero", 2
            Tx "E-ALIMENTADOR", x - 3, yb + 17, 2, ws.Range("BX99").Value & IIf(ws.Range("BX100").Value <> "", " + " & ws.Range("BX100").Value & "N", "") & IIf(ws.Range("BX101").Value <> "", " + " & ws.Range("BX101").Value & "T", "") & " " & ws.Range("F3").Value, 2
            Tx "E-ALIMENTADOR", x - 3, yb + 13, 2, ws.Range("I15").Value & " - tablero " & Format$(NumV(ws.Range("I23").Value), "0.0") & " m  " & ws.Range("BX102").Value & " mm", 2
            Tx "E-EQUIPOS", x + 3, yA + 2, 2.2, "NORMAL" & IIf(ws.Range("I3").Value = "SI", " (transformador)", IIf(padre <> "", " (" & padre & ")", IIf(CStr(ws.Range("C11").Value) <> "", " (" & ws.Range("C11").Value & ")", ""))) & IIf(NumV(ws.Range("I24").Value) > 0, "  " & Format$(NumV(ws.Range("I24").Value), "0.0") & " m", "")
            Tx "E-EQUIPOS", x + spx(sym & "|alterna") * ESC + 3, yA + spy(sym & "|alterna") * ESC - 4, 2.2, "EMERGENCIA"
        End If
        gx = x + 48
        Ln "E-ALIMENTADOR", xa, ya2, gx, ya2
        yBa = ya2 + 26
        Ln "E-ALIMENTADOR", gx, ya2, gx, yBa - hb
        Breaker gx, yBa, PolosB(ws), ws.Range("BX24").Value, "TM", "", "", NumV(ws.Range("BX24").Value) & " A " & PolosB(ws) & "P"
        yG = ya2 + 46
        Ln "E-ALIMENTADOR", gx, yBa, gx, yG
        If ws.Range("I16").Value = "Generador" Then
            Simbolo "GENERADOR", gx, yG, Array("NOMBRE", IIf(CStr(ws.Range("I18").Value) <> "", "Generador", ""), "KVA", ws.Range("I18").Value & " kVA", "XD", ws.Range("I19").Value & " %")
            Tx "E-EQUIPOS", gx + 9, yG + 5, 2.5, "Generador " & ws.Range("I18").Value & " kVA  X''d " & ws.Range("I19").Value & " %"
        Else
            Tx "E-EQUIPOS", gx + 4, yG, 2.5, "BYPASS desde " & ws.Range("I17").Value
            ' el bypass sube hasta la barra del tablero que lo alimenta
            If padre <> "" And hoja.Exists(padre) And Trim$(CStr(ws.Range("I17").Value)) = padre Then Ln "E-ALIMENTADOR", gx, yG, gx, ys(padre) - 20
        End If
        Tx "E-ALIMENTADOR", gx + 4, ya2 + 3, 2, ws.Range("BX95").Value & "  " & Format$(NumV(ws.Range("I20").Value), "0.0") & " m"
        If NumV(ws.Range("BX113").Value) > 0 Then Tx "E-ALIMENTADOR", gx + 4, ya2 - 1, 2, "Icc en el " & ws.Range("I15").Value & " " & Format$(NumV(ws.Range("BX113").Value), "0.00") & " kA"
        Ln "E-ALIMENTADOR", x, hu(nh, 1), x, yb
        LineaV x, yTop, hu(nh, 1), hu, nh, "E-ALIMENTADOR"
    Else
        LineaV x, yTop, yb, hu, nh, "E-ALIMENTADOR"
    End If
    ' barra de tierra del tablero de acometida (conductor del electrodo, 250.66)
    If padre = "" Then
        Ln "E-TABLERO", x0 + 4, yb, x0 + 4, yb - 6
        Simbolo "TIERRA", x0 + 4, yb - 6
        Tx "E-TABLERO", x0 + 9, yb - 13, 2, "Tierra " & IIf(CStr(ws.Range("BX35").Value) <> "", ws.Range("BX35").Value & " AWG", "")
    End If
    ' tabla de datos junto al tablero
    filas = Array("kVA conectados: " & Format$(NumV(ws.Range("BX10").Value), "0.00") & "   demandados: " & Format$(NumV(ws.Range("BX14").Value), "0.00"), _
                  "Sistema: " & ws.Range("C6").Value & " V  " & ws.Range("C5").Value & "   Proteccion: " & ws.Range("BX24").Value & " A " & ws.Range("F14").Value & IIf(ws.Range("F15").Value = "SI", " 100%", ""), _
                  "Tablero: " & ws.Range("BX63").Value & " (" & ws.Range("BX60").Value & ")  " & ws.Range("BX67").Value & " esp.", _
                  "Principal: " & ws.Range("BX72").Value & "  SCCR " & ws.Range("BX74").Value & " kA", _
                  "Alimentador: " & ws.Range("BX99").Value & " " & ws.Range("F3").Value & " " & ws.Range("F4").Value & "  tubo " & ws.Range("BX102").Value & " mm", _
                  "V bornes: " & Format$(NumV(ws.Range("BX49").Value), "0.00") & " V   " & ChrW(916) & "V acum.: " & Format$(NumV(ws.Range("BX51").Value), "0.00") & " %", _
                  "Icc: " & Format$(NumV(ws.Range("BX58").Value), "0.00") & " kA   SPD: " & ws.Range("BX79").Value)
    If CStr(ws.Range("BX118").Value) <> "" Then
        ReDim Preserve filas(UBound(filas) + IIf(ws.Range("BX84").Value = True, 4, 1))
        filas(UBound(filas) - IIf(ws.Range("BX84").Value = True, 3, 0)) = ws.Range("BX118").Value & ": " & ws.Range("BX119").Value & " " & ws.Range("BX120").Value & " A  " & ws.Range("BX121").Value & " " & ws.Range("BX122").Value & "P  WCR " & ws.Range("BX123").Value & " / Icc " & Format$(NumV(ws.Range("BX124").Value), "0.00") & " kA"
        If ws.Range("BX84").Value = True Then
        filas(UBound(filas) - 2) = "Normal-" & ws.Range("BX118").Value & ": " & Format$(NumV(ws.Range("BX126").Value), "0.0") & " m   " & ws.Range("BX125").Value & "-" & ws.Range("BX118").Value & ": " & Format$(NumV(ws.Range("BX127").Value), "0.0") & " m"
        filas(UBound(filas) - 1) = ws.Range("BX118").Value & "-tablero: " & Format$(NumV(ws.Range("BX128").Value), "0.0") & " m   2a acom.: " & ws.Range("BX129").Value
        filas(UBound(filas)) = "Caida 2a acometida: " & Format$(NumV(ws.Range("BX130").Value), "0.00") & " %"
        End If
    End If
    If CStr(ws.Range("BX131").Value) <> "" Then
        ReDim Preserve filas(UBound(filas) + 1)
        filas(UBound(filas)) = "UPS: " & Format$(NumV(ws.Range("BX131").Value), "0.0") & " kVA   Icc inversor " & Format$(NumV(ws.Range("BX132").Value), "0.00") & " kA"
    End If
    Rect "E-TABLA", x + 27, yb - 3, x + 27 + 95, yb - 3 - 4.2 * (UBound(filas) + 1) - 2
    For i = 0 To UBound(filas)
        Tx "E-TABLA", x + 29, yb - 7 - 4.2 * i, 2.2, filas(i)
    Next i
End Sub

Private Function OrDefault(v As Variant) As String
    If IsError(v) Then OrDefault = "": Exit Function
    If IsEmpty(v) Then OrDefault = "TM" Else If CStr(v) = "" Then OrDefault = "TM" Else OrDefault = CStr(v)
End Function

Private Sub Escribir(ruta As String)
    Dim h As String, fn As Integer, capas As Variant, colores As Variant, i As Long
    capas = Array("E-TABLERO", "E-ALIMENTADOR", "E-PROTECCION", "E-TRAFO", "E-EQUIPOS", "E-TEXTO", "E-TABLA")
    colores = Array(1, 3, 2, 5, 6, 7, 8)
    h = "0" & vbCrLf & "SECTION" & vbCrLf & "2" & vbCrLf & "HEADER" & vbCrLf & "9" & vbCrLf & "$ACADVER" & vbCrLf & "1" & vbCrLf & "AC1009" & vbCrLf & "0" & vbCrLf & "ENDSEC" & vbCrLf
    h = h & "0" & vbCrLf & "SECTION" & vbCrLf & "2" & vbCrLf & "TABLES" & vbCrLf
    h = h & "0" & vbCrLf & "TABLE" & vbCrLf & "2" & vbCrLf & "LTYPE" & vbCrLf & "70" & vbCrLf & "2" & vbCrLf
    h = h & "0" & vbCrLf & "LTYPE" & vbCrLf & "2" & vbCrLf & "CONTINUOUS" & vbCrLf & "70" & vbCrLf & "0" & vbCrLf & "3" & vbCrLf & "Solid line" & vbCrLf & "72" & vbCrLf & "65" & vbCrLf & "73" & vbCrLf & "0" & vbCrLf & "40" & vbCrLf & "0.0" & vbCrLf
    h = h & "0" & vbCrLf & "LTYPE" & vbCrLf & "2" & vbCrLf & "DASHED" & vbCrLf & "70" & vbCrLf & "0" & vbCrLf & "3" & vbCrLf & "Dashed __ __ __" & vbCrLf & "72" & vbCrLf & "65" & vbCrLf & "73" & vbCrLf & "2" & vbCrLf & "40" & vbCrLf & "3.0" & vbCrLf & "49" & vbCrLf & "2.0" & vbCrLf & "49" & vbCrLf & "-1.0" & vbCrLf
    h = h & "0" & vbCrLf & "ENDTAB" & vbCrLf
    h = h & "0" & vbCrLf & "TABLE" & vbCrLf & "2" & vbCrLf & "LAYER" & vbCrLf & "70" & vbCrLf & (UBound(capas) + 1) & vbCrLf
    For i = 0 To UBound(capas)
        h = h & "0" & vbCrLf & "LAYER" & vbCrLf & "2" & vbCrLf & capas(i) & vbCrLf & "70" & vbCrLf & "0" & vbCrLf & "62" & vbCrLf & colores(i) & vbCrLf & "6" & vbCrLf & "CONTINUOUS" & vbCrLf
    Next i
    h = h & "0" & vbCrLf & "ENDTAB" & vbCrLf
    h = h & "0" & vbCrLf & "TABLE" & vbCrLf & "2" & vbCrLf & "STYLE" & vbCrLf & "70" & vbCrLf & "1" & vbCrLf
    h = h & "0" & vbCrLf & "STYLE" & vbCrLf & "2" & vbCrLf & "SINERGIA" & vbCrLf & "70" & vbCrLf & "0" & vbCrLf & "40" & vbCrLf & "0.0" & vbCrLf & "41" & vbCrLf & "1.0" & vbCrLf & "50" & vbCrLf & "0.0" & vbCrLf & "71" & vbCrLf & "0" & vbCrLf & "42" & vbCrLf & "2.5" & vbCrLf & "3" & vbCrLf & "GOTHIC.TTF" & vbCrLf & "4" & vbCrLf & "" & vbCrLf
    h = h & "0" & vbCrLf & "ENDTAB" & vbCrLf & "0" & vbCrLf & "ENDSEC" & vbCrLf
    h = h & "0" & vbCrLf & "SECTION" & vbCrLf & "2" & vbCrLf & "BLOCKS" & vbCrLf & blq & "0" & vbCrLf & "ENDSEC" & vbCrLf
    h = h & "0" & vbCrLf & "SECTION" & vbCrLf & "2" & vbCrLf & "ENTITIES" & vbCrLf & buf & "0" & vbCrLf & "ENDSEC" & vbCrLf & "0" & vbCrLf & "EOF" & vbCrLf
    fn = FreeFile
    Open ruta For Output As #fn
    Print #fn, h;
    Close #fn
End Sub
