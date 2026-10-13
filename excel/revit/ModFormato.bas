Attribute VB_Name = "ModFormato"
' Paso 2 del proceso Revit -> ANACAR:
'   1. Revit exporta "Electrical Circuit Schedule.csv" (separado por tabulaciones, UTF-8).
'   2. Este libro lo limpia y lo acomoda en la tabla de Hoja1 (Tabla1, columnas C a J) SIN PERDER FILAS,
'      y deja en la hoja "Revision" el control: lineas leidas, circuitos escritos, filas omitidas con su motivo,
'      totales por tablero, posiciones repetidas, longitudes faltantes y alimentadores de tableros.
'   3. ANACAR (web o Excel) importa este libro.
' Los numeros se leen con punto o coma decimal sin depender de la configuracion regional y el numero de
' circuito se guarda como TEXTO ("24,26"), para que no se convierta en 24,26 o 2426.
Option Explicit
Option Compare Text

Private Const HOJA As String = "Hoja1"
Private Const FILA_ENC As Long = 5          ' encabezados de Tabla1
Private Const COL0 As Long = 3              ' columna C

Public Sub ImportarCircuitosRevit(Optional ruta As String = "")
    Dim f As Variant
    f = ruta
    If f = "" Then
        f = Application.GetOpenFilename("Tabla de circuitos de Revit (*.csv;*.txt;*.tsv),*.csv;*.txt;*.tsv,Todos los archivos (*.*),*.*", , "Seleccionar archivo CSV exportado de Revit")
        If f = False Then Exit Sub
    End If
    Formatear CStr(f), (ruta = "")
End Sub

' ---------- lectura ----------
Private Function LeerTexto(ruta As String) As String
    Dim st As Object, b() As Byte, t As String, cs As String
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open: st.LoadFromFile ruta
    If st.Size >= 2 Then
        b = st.Read(3)
        If UBound(b) >= 1 Then
            If b(0) = &HFF And b(1) = &HFE Then cs = "unicode"
            If b(0) = &HFE And b(1) = &HFF Then cs = "unicodeFFFE"
        End If
        If UBound(b) >= 2 And cs = "" Then
            If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then cs = "utf-8"
        End If
    End If
    st.Close
    If cs = "" Then cs = "utf-8"
    st.Type = 2: st.Charset = cs: st.Open: st.LoadFromFile ruta
    t = st.ReadText: st.Close
    If cs = "utf-8" And InStr(t, ChrW(&HFFFD)) > 0 Then          ' no era UTF-8: ANSI (Windows-1252)
        st.Type = 2: st.Charset = "windows-1252": st.Open: st.LoadFromFile ruta
        t = st.ReadText: st.Close
    End If
    If Len(t) > 0 Then
        If AscW(Left$(t, 1)) = &HFEFF Then t = Mid$(t, 2)
    End If
    LeerTexto = t
End Function

Private Function Partir(linea As String, sep As String) As Variant
    Dim out() As String, k As Long, cur As String, q As Boolean, i As Long, ch As String
    If sep <> "," Then
        Dim a As Variant, j As Long
        a = Split(linea, sep)
        For j = 0 To UBound(a)
            a(j) = Trim$(Replace(Replace(a(j), """", ""), vbTab, " "))
        Next j
        Partir = a
        Exit Function
    End If
    ReDim out(0 To 0)
    For i = 1 To Len(linea)
        ch = Mid$(linea, i, 1)
        If ch = """" Then
            q = Not q
        ElseIf ch = "," And Not q Then
            out(k) = Trim$(cur): cur = "": k = k + 1: ReDim Preserve out(0 To k)
        Else
            cur = cur & ch
        End If
    Next i
    out(k) = Trim$(cur)
    Partir = out
End Function

' numero sin depender de la configuracion regional: "1.08 kVA", "1,08", "1.234,5", "12' - 6"" (pies), "Not Computed"
Public Function NumTxt(ByVal s As String, Optional ByRef ok As Boolean, Optional ByRef pies As Boolean) As Double
    Dim t As String, i As Long, ch As String, p As Long, pulg As Double
    ok = False: pies = False
    s = Trim$(s)
    If s = "" Then Exit Function
    p = InStr(s, "'")
    If p > 0 Then
        pies = True
        NumTxt = NumTxt(Left$(s, p - 1), ok)
        If InStr(p, s, """") > 0 Then pulg = NumTxt(Replace(Replace(Mid$(s, p + 1), """", ""), "-", ""))
        NumTxt = NumTxt + pulg / 12
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
    If t = "" Or t = "-" Then Exit Function
    If InStr(t, ",") > 0 And InStr(t, ".") > 0 Then
        If InStrRev(t, ",") > InStrRev(t, ".") Then t = Replace(Replace(t, ".", ""), ",", ".") Else t = Replace(t, ",", "")
    ElseIf InStr(t, ",") > 0 Then
        t = Replace(t, ",", ".")
    End If
    NumTxt = Val(t)                      ' Val siempre usa punto decimal
    ok = True
    If (s Like "*[0-9] VA*" Or s Like "*[0-9]VA*") And Not (s Like "*kVA*") Then NumTxt = NumTxt / 1000
End Function

' "24, 26" / "24.26" / "<unnamed>" -> "24,26" (solo digitos separados por coma)
Private Function Circuito(s As String) As String
    Dim i As Long, ch As String, cur As String, r As String
    s = s & " "
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then
            cur = cur & ch
        ElseIf cur <> "" Then
            r = r & IIf(r = "", "", ",") & CStr(CLng(cur)): cur = ""
        End If
    Next i
    Circuito = r
End Function

' ---------- proceso ----------
Private Sub Formatear(ruta As String, avisar As Boolean)
    Dim ws As Worksheet, rv As Worksheet, txt As String, lineas As Variant, sep As String, i As Long, j As Long, c As Variant
    Dim iV As Long, iK As Long, iP As Long, iC As Long, iNom As Long, iPol As Long, iLen As Long, enc As Long, h As String
    Dim fila As Long, n As Long, leidas As Long, omit As Collection, pan As Object, k As Variant
    Set ws = ThisWorkbook.Worksheets(HOJA)
    Set rv = HojaRevision()
    txt = Replace(Replace(LeerTexto(ruta), vbCrLf, vbLf), vbCr, vbLf)
    lineas = Split(txt, vbLf)
    ' separador
    For i = 0 To Application.WorksheetFunction.Min(UBound(lineas), 10)
        If InStr(lineas(i), vbTab) > 0 Then sep = vbTab: Exit For
    Next i
    If sep = "" Then
        For i = 0 To Application.WorksheetFunction.Min(UBound(lineas), 10)
            If InStr(lineas(i), ";") > 0 Then sep = ";": Exit For
        Next i
    End If
    If sep = "" Then sep = ","
    ' encabezados por nombre (no por posicion)
    enc = -1
    For i = 0 To Application.WorksheetFunction.Min(UBound(lineas), 30)
        c = Partir(CStr(lineas(i)), sep)
        iV = -1: iK = -1: iP = -1: iC = -1: iNom = -1: iPol = -1: iLen = -1
        For j = 0 To UBound(c)
            h = LCase$(Trim$(c(j)))
            If iV < 0 And h Like "*volt*" Then
                iV = j
            ElseIf iNom < 0 And (h Like "*load name*" Or h Like "*nombre*" Or h Like "*descrip*") Then
                iNom = j
            ElseIf iK < 0 And (h Like "*true power*" Or h Like "*true load*" Or h Like "*apparent*" Or h Like "*potencia*" Or h Like "*carga*" Or h Like "*kva*" Or h = "load" Or h Like "*power*") Then
                iK = j
            ElseIf iP < 0 And (h Like "*panel*" Or h Like "*tablero*") Then
                iP = j
            ElseIf iC < 0 And (h Like "*circuit*" Or h Like "*circuito*") And Not (h Like "*name*") Then
                iC = j
            ElseIf iPol < 0 And (h Like "*pole*" Or h Like "*polo*") Then
                iPol = j
            ElseIf iLen < 0 And (h Like "*length*" Or h Like "*longitud*" Or h Like "*largo*") Then
                iLen = j
            End If
        Next j
        If iP >= 0 And iC >= 0 Then enc = i: Exit For
    Next i
    If enc < 0 Then
        MsgBox "No se encontraron los encabezados Panel y Circuit Number en el archivo:" & vbCrLf & ruta, vbExclamation, "Importar circuitos"
        Exit Sub
    End If
    If iK < 0 Then MsgBox "Aviso: no se encontro la columna de carga (True Power / True Load). Las cargas quedaran en blanco.", vbExclamation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ' limpiar la tabla anterior
    Dim lo As ListObject, u As Long
    Set lo = ws.ListObjects(1)
    u = ws.Cells(ws.Rows.Count, COL0).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, COL0 + 2).End(xlUp).Row > u Then u = ws.Cells(ws.Rows.Count, COL0 + 2).End(xlUp).Row
    If u > FILA_ENC Then ws.Range(ws.Cells(FILA_ENC + 1, COL0), ws.Cells(u, COL0 + 7)).ClearContents
    ws.Range(ws.Cells(FILA_ENC + 1, COL0 + 3), ws.Cells(FILA_ENC + 2000, COL0 + 3)).NumberFormat = "@"
    Set omit = New Collection
    Set pan = CreateObject("Scripting.Dictionary")
    fila = FILA_ENC + 1
    Dim v As Double, okV As Boolean, kva As Double, okK As Boolean, lon As Double, okL As Boolean, pies As Boolean
    Dim panel As String, circ As String, polos As Long, nom As String, sinLong As Long, enPies As Long
    For i = enc + 1 To UBound(lineas)
        If Trim$(Replace(lineas(i), sep, "")) <> "" Then
            leidas = leidas + 1
            c = Partir(CStr(lineas(i)), sep)
            panel = Celda(c, iP): circ = Celda(c, iC): nom = Celda(c, iNom)
            If panel = "" Then
                omit.Add Array(i + 1, "Sin tablero (fila de resumen de Revit)", CStr(lineas(i)))
            Else
                v = NumTxt(Celda(c, iV), okV)
                kva = NumTxt(Celda(c, iK), okK)
                lon = NumTxt(Celda(c, iLen), okL, pies)
                If pies Then lon = lon * 0.3048: enPies = enPies + 1
                polos = CLng(NumTxt(Celda(c, iPol)))
                ws.Cells(fila, COL0).Value = IIf(okV, v, Empty)
                ws.Cells(fila, COL0 + 1).Value = IIf(okK, kva, Empty)
                ws.Cells(fila, COL0 + 2).Value = panel
                ws.Cells(fila, COL0 + 3).NumberFormat = "@"
                ws.Cells(fila, COL0 + 3).Value = Circuito(circ)
                ws.Cells(fila, COL0 + 5).Value = nom
                ws.Cells(fila, COL0 + 6).Value = IIf(polos > 0, polos, Empty)
                ws.Cells(fila, COL0 + 7).Value = IIf(okL, Round(lon, 2), Empty)
                If Not okL Then sinLong = sinLong + 1
                If Not pan.Exists(panel) Then pan.Add panel, New Collection
                pan(panel).Add Array(Circuito(circ), kva, polos, nom, okL, okK, okV, fila)
                fila = fila + 1: n = n + 1
            End If
        End If
    Next i
    ' formula PAR / IMPAR / MEZCLADO en la columna G
    If n > 0 Then
        ws.Range(ws.Cells(FILA_ENC + 1, COL0 + 4), ws.Cells(FILA_ENC + n, COL0 + 4)).Formula = _
            "=IFERROR(IF(ISEVEN(VALUE(LEFT(F6&"","",FIND("","",F6&"","")-1)))=ISEVEN(VALUE(TRIM(RIGHT(SUBSTITUTE(F6,"","",REPT("" "",20)),20)))),IF(ISEVEN(VALUE(LEFT(F6&"","",FIND("","",F6&"","")-1))),""PAR"",""IMPAR""),""MEZCLADO""),"""")"
    End If
    lo.Resize ws.Range(ws.Cells(FILA_ENC, COL0), ws.Cells(FILA_ENC + IIf(n > 0, n, 1), COL0 + 7))
    ws.Cells(4, COL0).Value = "Electrical Circuit Schedule"
    ws.Cells(3, COL0 + 5).Value = "Archivo: " & ruta & "  (" & Format$(Now, "yyyy-mm-dd hh:nn") & ")"
    Revision rv, ruta, leidas, n, omit, pan, sinLong, enPies
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If avisar Then
        MsgBox "Importacion completada." & vbCrLf & vbCrLf & _
               "Lineas de datos leidas: " & leidas & vbCrLf & _
               "Circuitos escritos en la tabla: " & n & vbCrLf & _
               "Filas sin tablero (resumen de Revit): " & omit.Count & vbCrLf & _
               "Perdidas: " & (leidas - n - omit.Count) & vbCrLf & vbCrLf & _
               "Revise la hoja 'Revision' (totales por tablero y avisos) y luego importe este libro en ANACAR.", vbInformation, "Importar circuitos"
    End If
End Sub

Private Function Celda(c As Variant, i As Long) As String
    If i < 0 Or i > UBound(c) Then Exit Function
    Celda = Trim$(Replace(CStr(c(i)), """", ""))
End Function

Private Function HojaRevision() As Worksheet
    On Error Resume Next
    Set HojaRevision = ThisWorkbook.Worksheets("Revision")
    On Error GoTo 0
    If HojaRevision Is Nothing Then
        Set HojaRevision = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        HojaRevision.Name = "Revision"
    End If
End Function

Private Sub Revision(rv As Worksheet, ruta As String, leidas As Long, n As Long, omit As Collection, pan As Object, sinLong As Long, enPies As Long)
    Dim r As Long, k As Variant, it As Variant, tot As Double, pol As Long, rep As String, sinL As Long, usados As Object, p As Variant, nn As Long, alim As String, k2 As Variant
    Dim totales As Object
    rv.Cells.Clear
    rv.Range("A1").Value = "REVISION DE LA IMPORTACION (paso 2 de 3: Revit -> este libro -> ANACAR)"
    rv.Range("A1").Font.Bold = True: rv.Range("A1").Font.Size = 14
    rv.Range("A3").Value = "Archivo": rv.Range("B3").Value = ruta
    rv.Range("A4").Value = "Fecha": rv.Range("B4").Value = Format$(Now, "yyyy-mm-dd hh:nn")
    rv.Range("A5").Value = "Lineas de datos leidas": rv.Range("B5").Value = leidas
    rv.Range("A6").Value = "Circuitos escritos en Hoja1": rv.Range("B6").Value = n
    rv.Range("A7").Value = "Filas sin tablero (resumen)": rv.Range("B7").Value = omit.Count
    rv.Range("A8").Value = "Perdidas": rv.Range("B8").Value = leidas - n - omit.Count
    rv.Range("A9").Value = "Circuitos sin longitud": rv.Range("B9").Value = sinLong
    rv.Range("A10").Value = "Longitudes convertidas de pies a m": rv.Range("B10").Value = enPies
    rv.Range("A3:A10").Font.Bold = True
    ' totales por tablero
    Set totales = CreateObject("Scripting.Dictionary")
    r = 12
    rv.Cells(r, 1).Resize(1, 7).Value = Array("Tablero", "Circuitos", "kVA (suma)", "Polos", "Posiciones repetidas", "Sin longitud", "Alimentado por (circuito en otro tablero)")
    rv.Cells(r, 1).Resize(1, 7).Font.Bold = True: rv.Cells(r, 1).Resize(1, 7).Interior.Color = RGB(217, 217, 217)
    For Each k In pan.Keys
        tot = 0
        For Each it In pan(k): tot = tot + it(1): Next it
        totales(k) = tot
    Next k
    For Each k In pan.Keys
        r = r + 1: tot = 0: pol = 0: rep = "": sinL = 0
        Set usados = CreateObject("Scripting.Dictionary")
        For Each it In pan(k)
            tot = tot + it(1): pol = pol + it(2)
            If Not it(4) Then sinL = sinL + 1
            For Each p In Split(it(0), ",")
                If p <> "" Then
                    If usados.Exists(p) Then rep = rep & IIf(rep = "", "", ", ") & p Else usados.Add p, True
                End If
            Next p
        Next it
        ' quien lo alimenta: circuito de otro tablero con el mismo nombre o la misma carga
        alim = ""
        For Each k2 In pan.Keys
            If k2 <> k Then
                For Each it In pan(k2)
                    If UCase$(Replace(it(3), " ", "")) = UCase$(Replace(k, " ", "")) Then
                        alim = k2 & " [" & it(0) & "] por nombre"
                    ElseIf alim = "" And Abs(it(1) - totales(k)) < 0.005 And totales(k) > 0 And it(2) >= 2 Then
                        alim = k2 & " [" & it(0) & "] '" & it(3) & "' por carga (" & Format$(it(1), "0.00") & " kVA) - revise el nombre en Revit"
                    End If
                Next it
            End If
        Next k2
        rv.Cells(r, 1).Value = k: rv.Cells(r, 2).Value = pan(k).Count: rv.Cells(r, 3).Value = Round(tot, 3): rv.Cells(r, 4).Value = pol
        rv.Cells(r, 5).Value = rep: rv.Cells(r, 6).Value = sinL: rv.Cells(r, 7).Value = alim
        If rep <> "" Then rv.Cells(r, 5).Interior.Color = RGB(255, 199, 206)
        If sinL > 0 Then rv.Cells(r, 6).Interior.Color = RGB(255, 235, 156)
        If InStr(alim, "por carga") > 0 Then rv.Cells(r, 7).Interior.Color = RGB(255, 235, 156)
        nn = nn + pan(k).Count
    Next k
    r = r + 1
    rv.Cells(r, 1).Value = "TOTAL": rv.Cells(r, 2).Value = nn: rv.Cells(r, 1).Resize(1, 3).Font.Bold = True
    ' filas omitidas
    r = r + 2
    rv.Cells(r, 1).Value = "FILAS NO ESCRITAS EN Hoja1 (no son circuitos de un tablero)": rv.Cells(r, 1).Font.Bold = True
    r = r + 1
    rv.Cells(r, 1).Resize(1, 3).Value = Array("Linea del archivo", "Motivo", "Contenido")
    rv.Cells(r, 1).Resize(1, 3).Font.Bold = True: rv.Cells(r, 1).Resize(1, 3).Interior.Color = RGB(217, 217, 217)
    For Each it In omit
        r = r + 1
        rv.Cells(r, 1).Value = it(0): rv.Cells(r, 2).Value = it(1): rv.Cells(r, 3).Value = Replace(it(2), vbTab, " | ")
    Next it
    rv.Columns("A").ColumnWidth = 30: rv.Columns("B").ColumnWidth = 18: rv.Columns("C:F").ColumnWidth = 14: rv.Columns("G").ColumnWidth = 70
End Sub
