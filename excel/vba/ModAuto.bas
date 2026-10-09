Attribute VB_Name = "ModAuto"
' Autollenado como en la herramienta en linea:
' - las celdas de datos traen el valor automatico (formula en gris); al borrar una celda vuelve el automatico,
' - las listas con nombre ("8 · TOMAS (120 V 1F)") dejan solo el numero,
' - crear varios tableros, tablero derivado, duplicar tablero / circuito y mover un circuito (anotado para Revit),
' - editar tipo, sistema, alimentado desde y longitud desde la lista de la hoja PROYECTO.
Option Explicit

Public Const ZONA_AUTO As String = "C15,F6,F10,F13,F16:F18,I7,F25:F124,I25:R124"
Public Const ZONA_LISTAS As String = "E25:E124,R25:R124,F16:F18"

Public Function Plantilla() As Worksheet
    Set Plantilla = ThisWorkbook.Worksheets("MC_MACHOTE")
End Function

' formula automatica de la plantilla para esa celda ("" si no tiene)
Public Function FormulaAuto(c As Range) As String
    Dim f As String
    f = Plantilla().Range(c.Address).Formula
    If Left$(f, 1) = "=" Then FormulaAuto = f
End Function

Public Function EsAuto(c As Range) As Boolean
    Dim f As String
    f = FormulaAuto(c)
    EsAuto = (f <> "" And c.Formula = f)
End Function

' al borrar celdas de la zona automatica se restaura la formula
Public Sub RestaurarAuto(ws As Worksheet, rng As Range)
    Dim c As Range, z As Range, f As String
    Set z = Intersect(rng, ws.Range(ZONA_AUTO))
    If z Is Nothing Then Exit Sub
    If z.Cells.Count > 3000 Then Exit Sub
    For Each c In z.Cells
        If Len(c.Formula) = 0 Then
            f = FormulaAuto(c)
            If f <> "" Then c.Formula = f
        End If
    Next c
End Sub

' "8 · TOMAS (120 V 1F)" -> 8
Public Sub ConvertirEtiquetas(ws As Worksheet, rng As Range)
    Dim c As Range, z As Range, s As String, p As Long
    Set z = Intersect(rng, ws.Range(ZONA_LISTAS))
    If z Is Nothing Then Exit Sub
    If z.Cells.Count > 3000 Then Exit Sub
    For Each c In z.Cells
        If VarType(c.Value) = vbString And Not c.HasFormula Then
            s = c.Value: p = InStr(s, " " & ChrW(183) & " ")
            If p > 1 Then
                If IsNumeric(Left$(s, p - 1)) Then c.Value = CDbl(Left$(s, p - 1))
            End If
        End If
    Next c
End Sub

' ---------- mover / duplicar circuitos ----------
' copia los datos del usuario de la fila 'de' a la fila 'a' (las celdas automaticas quedan automaticas en la nueva fila)
Public Sub CopiarFila(ws As Worksheet, de As Long, a As Long, polos As Variant)
    Dim col As Long, f() As String, auto() As Boolean
    ReDim f(1 To 22): ReDim auto(1 To 22)
    For col = 1 To 22
        f(col) = ws.Cells(de, col).Formula
        auto(col) = EsAuto(ws.Cells(de, col))
    Next col
    LimpiarFila ws, a
    For col = 1 To 22
        If col <> 2 And Not auto(col) And Len(f(col)) > 0 Then ws.Cells(a, col).Formula = f(col)
    Next col
    ' polos 2 y 3 segun la nueva posicion
    If Len(f(3)) > 0 Then ws.Cells(a, 3).Value = polos(2) Else ws.Cells(a, 3).ClearContents
    If Len(f(4)) > 0 Then ws.Cells(a, 4).Value = polos(3) Else ws.Cells(a, 4).ClearContents
End Sub

Private Function PolosEn(p As Long) As Variant
    Dim r(1 To 3) As Long
    r(1) = p: r(2) = p + 2: r(3) = p + 4
    PolosEn = r
End Function

Public Sub DuplicarCircuito()
    Dim ws As Worksheet, r As Long, np As Long, p As Long
    Set ws = MCActiva(): If ws Is Nothing Then Exit Sub
    r = ActiveCell.Row
    If r < FILA1 Or r > FILAN Or Not TieneDatos(ws, r) Then MsgBox "Seleccione una celda de la fila del circuito a duplicar.", vbInformation: Exit Sub
    np = UBound(PolosFila(ws, r))
    p = PosicionLibre(ws, np)
    If p = 0 Then MsgBox "No hay espacio libre.", vbExclamation: Exit Sub
    Congelar
    CopiarFila ws, r, FilaPos(p), PolosEn(p)
    ws.Cells(FilaPos(p), 19).ClearContents: ws.Cells(FilaPos(p), 20).ClearContents: ws.Cells(FilaPos(p), 21).ClearContents
    Descongelar
    ws.Cells(FilaPos(p), 7).Select
End Sub

' mueve el circuito de la fila activa a otra posicion y lo anota en CAMBIOS REVIT
Public Sub MoverCircuito()
    Dim ws As Worksheet, r As Long, v As String, p As Long, np As Long, k As Long, oc As Variant, de As String, viejo As Variant, i As Long
    Set ws = MCActiva(): If ws Is Nothing Then Exit Sub
    r = ActiveCell.Row
    If r < FILA1 Or r > FILAN Or Not TieneDatos(ws, r) Then MsgBox "Seleccione una celda de la fila del circuito a mover.", vbInformation: Exit Sub
    viejo = PolosFila(ws, r): np = UBound(viejo)
    For i = 1 To np: de = de & IIf(i > 1, ",", "") & viejo(i): Next i
    v = InputBox("Circuito en [" & de & "] " & ws.Cells(r, 23).Value & vbCrLf & "Nueva posicion (primer polo):", "Mover circuito")
    If v = "" Then Exit Sub
    p = CLng(Val(v))
    If p < 1 Or p + 2 * (np - 1) > 100 Then MsgBox "Posicion no valida.", vbExclamation: Exit Sub
    oc = Ocupacion(ws)
    For k = 0 To np - 1
        If oc(p + 2 * k) Then
            Dim propio As Boolean: propio = False
            For i = 1 To np
                If viejo(i) = p + 2 * k Then propio = True
            Next i
            If Not propio Then MsgBox "La posicion " & (p + 2 * k) & " esta ocupada.", vbExclamation: Exit Sub
        End If
    Next k
    Congelar
    Dim tmp As Long
    tmp = 0
    If FilaPos(p) <> r Then
        CopiarFila ws, r, FilaPos(p), PolosEn(p)
        LimpiarFila ws, r
    End If
    Anotar ws, CStr(ws.Cells(FilaPos(p), 23).Value), IIf(CStr(ws.Cells(FilaPos(p), 21).Value) <> "", CStr(ws.Cells(FilaPos(p), 21).Value), de), Join(Array(p, p + 2, p + 4), ","), np, "manual"
    Descongelar
    ActualizarProyecto
    ws.Activate: ws.Cells(FilaPos(p), 7).Select
End Sub

Public Sub Anotar(ws As Worksheet, desc As String, de As String, a As String, np As Long, origen As String)
    Dim cm As Worksheet, u As Long, partes As Variant, i As Long, txt As String
    partes = Split(a, ",")
    For i = 0 To np - 1
        txt = txt & IIf(i > 0, ",", "") & partes(i)
    Next i
    Set cm = ThisWorkbook.Worksheets("CAMBIOS REVIT")
    u = cm.Cells(cm.Rows.Count, 1).End(xlUp).Row + 1
    cm.Cells(u, 1).Value = Date
    cm.Cells(u, 2).Value = IdDe(ws)
    cm.Cells(u, 3).Value = desc
    cm.Cells(u, 4).Value = "'" & de
    cm.Cells(u, 5).Value = "'" & txt
    cm.Cells(u, 6).Value = origen
End Sub

' ---------- tableros ----------
Public Sub CrearVarios()
    Dim txt As String, nombres As Variant, n As Long, pref As String, i As Long, tipo As String, sis As String, padre As String, id As String, ws As Worksheet, creados As Long
    txt = InputBox("Nombres de los tableros separados por coma (TA, TB, 2G1, 2G2...)" & vbCrLf & "o la CANTIDAD para nombrarlos con un prefijo + letra:", "Crear varios tableros", "3")
    If Trim$(txt) = "" Then Exit Sub
    If IsNumeric(txt) Then
        n = CLng(txt): If n < 1 Or n > 60 Then Exit Sub
        pref = InputBox("Prefijo del nombre (se agrega A, B, C...):", "Crear varios tableros", "T")
        ReDim nombres(0 To n - 1)
        For i = 0 To n - 1
            nombres(i) = pref & Chr$(65 + (i Mod 26)) & IIf(i >= 26, CStr(i \ 26), "")
        Next i
    Else
        nombres = Split(Replace(txt, ";", ","), ",")
    End If
    tipo = UCase$(Trim$(InputBox("Tipo (3F o 1F):", "Crear varios tableros", "3F")))
    If tipo <> "1F" Then tipo = "3F"
    sis = Trim$(InputBox("Sistema (V): 120/208, 120/240, 277/480...", "Crear varios tableros", IIf(tipo = "1F", "120/240", "120/208")))
    If sis = "" Then sis = "120/208"
    padre = Trim$(InputBox("Alimentados desde (ID de un tablero existente; vacio = acometida / externo):", "Crear varios tableros", ""))
    Congelar
    For i = 0 To UBound(nombres)
        id = Trim$(nombres(i))
        If id <> "" Then
            Set ws = CrearTablero(id, tipo, sis, True)
            If Not ws Is Nothing Then
                creados = creados + 1
                If padre <> "" Then ws.Range("C10").Value = padre: ConectarTablero ws, True
            End If
        End If
    Next i
    Descongelar
    ActualizarProyecto
    MsgBox creados & " tablero(s) creados.", vbInformation
End Sub

Public Sub NuevoDerivado()
    Dim padre As Worksheet, id As String, ws As Worksheet
    Set padre = MCActiva(): If padre Is Nothing Then Exit Sub
    id = Trim$(InputBox("Nombre del tablero derivado de " & IdDe(padre) & ":", "Nuevo tablero derivado"))
    If id = "" Then Exit Sub
    Congelar
    Set ws = CrearTablero(id, CStr(padre.Range("C5").Value), CStr(padre.Range("C6").Value))
    If Not ws Is Nothing Then
        ws.Range("C10").Value = IdDe(padre)
        ConectarTablero ws, True
    End If
    Descongelar
    If ws Is Nothing Then Exit Sub
    ActualizarProyecto
    ws.Activate: ws.Range("C12").Select
End Sub

Public Sub DuplicarTablero()
    Dim o As Worksheet, ws As Worksheet, id As String, r As Long, c As Range, rng As Variant
    Set o = MCActiva(): If o Is Nothing Then Exit Sub
    id = Trim$(InputBox("Nombre de la copia de " & IdDe(o) & ":", "Duplicar tablero", IdDe(o) & "-2"))
    If id = "" Then Exit Sub
    Congelar
    Set ws = CrearTablero(id, CStr(o.Range("C5").Value), CStr(o.Range("C6").Value), True)
    If ws Is Nothing Then Descongelar: MsgBox "No se pudo crear " & id & " (nombre no valido o repetido).", vbExclamation: Exit Sub
    For Each rng In Array("C4", "C6:C6", "C11:C20", "F3:F19", "I3:I21", "A25:V124")
        For Each c In o.Range(rng).Cells
            If c.Column <> 2 Or c.Row < FILA1 Then
                If Not EsAuto(c) Then ws.Range(c.Address).Formula = c.Formula
            End If
        Next c
    Next rng
    ws.Range("C3").Value = id
    ws.Range("C10").Value = o.Range("C10").Value
    For r = FILA1 To FILAN
        If Len(ws.Cells(r, 19).Formula) > 0 Or Len(ws.Cells(r, 20).Formula) > 0 Then LimpiarFila ws, r
    Next r
    ConectarTablero ws, True
    Descongelar
    ActualizarProyecto
    ws.Activate
End Sub

' ---------- edicion desde la lista de PROYECTO ----------
Public Sub CambioProyecto(Target As Range)
    Dim c As Range, id As String, ws As Worksheet, pj As Worksheet
    Set pj = ThisWorkbook.Worksheets("PROYECTO")
    If Intersect(Target, pj.Range("C17:F116")) Is Nothing Then Exit Sub
    For Each c In Intersect(Target, pj.Range("C17:F116")).Cells
        id = Trim$(CStr(pj.Cells(c.Row, 2).Value))
        If id <> "" Then
            Set ws = HojaMC(id)
            If Not ws Is Nothing Then
                Select Case c.Column
                    Case 3: ws.Range("C5").Value = c.Value: RenombrarTablero ws
                    Case 4: ws.Range("C6").Value = c.Value: AjustarVista ws
                    Case 5: ws.Range("C10").Value = c.Value: ConectarTablero ws
                    Case 6: ws.Range("C12").Value = c.Value
                End Select
            End If
        End If
    Next c
    ActualizarLista
End Sub
