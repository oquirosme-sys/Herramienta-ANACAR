Attribute VB_Name = "ModDXF"
' Diagrama unifilar en DXF (AutoCAD R12): acometida / transformador / UPS, barras de cada tablero,
' interruptores, alimentadores con su calibre, segunda acometida (ATS/MTS/IP) y tabla de datos junto a cada tablero.
Option Explicit

Private buf As String
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

Private Sub Breaker(x As Double, y As Double, txt As String)
    ' interruptor: arco simplificado (dos lineas) y texto con amperios
    Ln "E-PROTECCION", x, y + 6, x + 4, y - 2
    Ln "E-PROTECCION", x - 2, y + 6, x + 2, y + 6
    Tx "E-PROTECCION", x + 6, y, 2.5, txt
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
    For Each k In hoja.Keys
        Dibujar CStr(k)
    Next k
    Tx "E-TEXTO", 0, 70, 6, "DIAGRAMA UNIFILAR - " & ThisWorkbook.Names("P_NOMBRE").RefersToRange.Value
    Tx "E-TEXTO", 0, 60, 3, "ANACAR " & Format$(Date, "yyyy-mm-dd") & " - corriente de cortocircuito por el metodo punto a punto; caidas de voltaje acumuladas."
    Escribir CStr(ruta)
    If destino = "" Then MsgBox "Diagrama unifilar exportado:" & vbCrLf & ruta, vbInformation
End Sub

Private Sub Dibujar(id As String)
    Dim ws As Worksheet, x As Double, y As Double, padre As String, c As Collection, k As Variant, x0 As Double, x1 As Double, yb As Double
    Dim filas As Variant, i As Long, ytop As Double
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
    ' alimentacion
    ytop = y + 30
    padre = Trim$(CStr(ws.Range("C10").Value))
    If padre <> "" And hoja.Exists(padre) Then
        Ln "E-ALIMENTADOR", x, ys(padre) - 21, x, yb
        Breaker x, ys(padre) - 32, NumV(ws.Range("BX24").Value) & " A"
        Tx "E-ALIMENTADOR", x - 3, ys(padre) - 45, 2.2, ws.Range("BX99").Value & IIf(ws.Range("BX100").Value <> "", " + " & ws.Range("BX100").Value & "N", "") & IIf(ws.Range("BX101").Value <> "", " + " & ws.Range("BX101").Value & "T", "") & " " & ws.Range("F3").Value & " - " & Format$(NumV(ws.Range("C12").Value), "0.0") & " m", 0, 90
    Else
        Ln "E-ALIMENTADOR", x, ytop, x, yb
        Breaker x, y + 2, IIf(ws.Range("C16").Value = "Zapatas", "ZAPATAS", NumV(ws.Range("BX24").Value) & " A")
        Tx "E-TEXTO", x + 3, ytop + 2, 2.5, IIf(CStr(ws.Range("C11").Value) <> "", ws.Range("C11").Value, "ACOMETIDA")
    End If
    If ws.Range("I3").Value = "SI" Then
        Circ "E-TRAFO", x, y + 18, 5: Circ "E-TRAFO", x, y + 11, 5
        Tx "E-TRAFO", x + 8, y + 14, 2.5, IIf(CStr(ws.Range("I8").Value) <> "", ws.Range("I8").Value & " ", "") & ws.Range("I4").Value & " kVA  Z " & ws.Range("I5").Value & " %"
    End If
    If ws.Range("I10").Value = "SI" Then
        Rect "E-EQUIPOS", x - 7, y + 6, x + 7, y + 18
        Tx "E-EQUIPOS", x, y + 10, 3, "UPS", 1
        Tx "E-EQUIPOS", x + 9, y + 10, 2.5, ws.Range("I11").Value & " kVA"
    End If
    If ws.Range("I14").Value = "SI" Then
        ' segunda acometida con equipo de transferencia pegado a la barra
        Rect "E-EQUIPOS", x + 12, yb + 4, x + 26, yb + 14
        Tx "E-EQUIPOS", x + 19, yb + 7.5, 3, ws.Range("I15").Value, 1
        Ln "E-ALIMENTADOR", x + 19, yb + 4, x + 19, yb
        Ln "E-ALIMENTADOR", x + 26, yb + 9, x + 45, yb + 9, , "DASHED"
        If ws.Range("I16").Value = "Generador" Then
            Circ "E-EQUIPOS", x + 50, yb + 9, 5
            Tx "E-EQUIPOS", x + 50, yb + 7.5, 3, "G", 1
            Tx "E-EQUIPOS", x + 56, yb + 9, 2.5, ws.Range("I18").Value & " kVA"
        Else
            Tx "E-EQUIPOS", x + 46, yb + 9, 2.5, "BYPASS desde " & ws.Range("I17").Value
        End If
        Breaker x + 35, yb + 9, NumV(ws.Range("BX24").Value) & " A"
        Tx "E-ALIMENTADOR", x + 28, yb + 15, 2, ws.Range("BX95").Value
    End If
    ' tabla de datos junto al tablero
    filas = Array("kVA conectados: " & Format$(NumV(ws.Range("BX10").Value), "0.00") & "   demandados: " & Format$(NumV(ws.Range("BX14").Value), "0.00"), _
                  "Sistema: " & ws.Range("C6").Value & " V  " & ws.Range("C5").Value & "   Proteccion: " & ws.Range("BX24").Value & " A " & ws.Range("F14").Value & IIf(ws.Range("F15").Value = "SI", " 100%", ""), _
                  "Tablero: " & ws.Range("BX63").Value & " (" & ws.Range("BX60").Value & ")  " & ws.Range("BX67").Value & " esp.", _
                  "Principal: " & ws.Range("BX72").Value & "  SCCR " & ws.Range("BX74").Value & " kA", _
                  "Alimentador: " & ws.Range("BX99").Value & " " & ws.Range("F3").Value & " " & ws.Range("F4").Value & "  tubo " & ws.Range("BX102").Value & " mm", _
                  "V bornes: " & Format$(NumV(ws.Range("BX49").Value), "0.00") & " V   " & ChrW(916) & "V acum.: " & Format$(NumV(ws.Range("BX51").Value), "0.00") & " %", _
                  "Icc: " & Format$(NumV(ws.Range("BX58").Value), "0.00") & " kA   SPD: " & ws.Range("BX79").Value)
    Rect "E-TABLA", x + 27, yb - 3, x + 27 + 95, yb - 3 - 4.2 * (UBound(filas) + 1) - 2
    For i = 0 To UBound(filas)
        Tx "E-TABLA", x + 29, yb - 7 - 4.2 * i, 2.2, filas(i)
    Next i
    ' derivados: interruptor en la barra
    For Each k In c
        Ln "E-ALIMENTADOR", xs(k), yb, xs(k), ys(k) - 21
    Next k
End Sub

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
    h = h & "0" & vbCrLf & "SECTION" & vbCrLf & "2" & vbCrLf & "ENTITIES" & vbCrLf & buf & "0" & vbCrLf & "ENDSEC" & vbCrLf & "0" & vbCrLf & "EOF" & vbCrLf
    fn = FreeFile
    Open ruta For Output As #fn
    Print #fn, h;
    Close #fn
End Sub
