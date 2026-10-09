Attribute VB_Name = "ModTableros"
' Tableros del proyecto: cada tablero es una hoja "MC <ID>" (memoria de calculo, copia de MC_MACHOTE)
' y una hoja "3F <ID>" o "1F <ID>" (cuadro de cargas con el diseno del Excel original, copia de VISTA_MACHOTE).
' La conexion entre tableros (cascada) se hace con "Alimentado desde" (celda C10 de la MC).
Option Explicit
Option Compare Text

Public Const FILA1 As Long = 25
Public Const FILAN As Long = 124
Public enProceso As Boolean

' ---------------- utilidades ----------------
Public Function NumV(v As Variant) As Double
    ' numero de una celda sin depender de la configuracion regional (Val trunca "5,4")
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then NumV = CDbl(v)
End Function

' congelar Excel durante operaciones masivas (sin recalculo, eventos ni pantalla) y restaurar
Public Sub Congelar()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub Descongelar()
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' escribe la formula solo si cambio (evita recalculos en cascada)
Public Sub SetF(c As Range, f As String)
    If c.Formula <> f Then c.Formula = f
End Sub

Public Function FilaPos(p As Long) As Long
    If p Mod 2 = 1 Then FilaPos = FILA1 + (p - 1) \ 2 Else FilaPos = 75 + (p - 2) \ 2
End Function

Public Function EsMC(ws As Object) As Boolean
    EsMC = (Left$(ws.Name, 3) = "MC ")
End Function

Public Function HojaMC(id As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If EsMC(ws) Then
            If Trim$(CStr(ws.Range("C3").Value)) = Trim$(id) Then Set HojaMC = ws: Exit Function
        End If
    Next ws
End Function

Public Function HojaVista(id As String) As Worksheet
    On Error Resume Next
    Set HojaVista = ThisWorkbook.Worksheets("3F " & id)
    If HojaVista Is Nothing Then Set HojaVista = ThisWorkbook.Worksheets("1F " & id)
End Function

Public Function IdDe(ws As Worksheet) As String
    IdDe = Trim$(CStr(ws.Range("C3").Value))
End Function

Public Function Tableros() As Collection
    Dim ws As Worksheet, c As New Collection
    For Each ws In ThisWorkbook.Worksheets
        If EsMC(ws) Then c.Add ws
    Next ws
    Set Tableros = c
End Function

Private Function IdValido(id As String) As Boolean
    Dim ch As Variant
    If Len(id) = 0 Or Len(id) > 27 Then Exit Function
    For Each ch In Array("[", "]", ":", "*", "?", "/", "\", "'")
        If InStr(id, ch) > 0 Then Exit Function
    Next ch
    IdValido = True
End Function

Public Function RefMC(id As String) As String
    RefMC = "'MC " & id & "'!"
End Function

Public Function TieneDatos(ws As Worksheet, r As Long) As Boolean
    ' fila con circuito: detalle, kVA, descripcion propia o enlace a otro tablero (las celdas automaticas no cuentan)
    If Len(ws.Cells(r, 5).Formula) + Len(ws.Cells(r, 7).Formula) + Len(ws.Cells(r, 19).Formula) + Len(ws.Cells(r, 20).Formula) > 0 Then TieneDatos = True: Exit Function
    If Len(ws.Cells(r, 6).Formula) > 0 Then TieneDatos = Not ModAuto.EsAuto(ws.Cells(r, 6))
End Function

Public Function FasesDet(det As Variant) As Long
    On Error Resume Next
    FasesDet = CLng(Application.WorksheetFunction.VLookup(det, ThisWorkbook.Names("T_DET").RefersToRange, 5, False))
End Function

' polos que ocupa el circuito de la fila r
Public Function PolosFila(ws As Worksheet, r As Long) As Variant
    ' polos escritos en la fila (posicion y polos 2 y 3), como en el Excel ANACAR
    Dim a(1 To 3) As Long, k As Long, res() As Long, i As Long
    a(1) = ws.Cells(r, 2).Value: k = 1
    If NumV(ws.Cells(r, 3).Value) > 0 Then k = k + 1: a(k) = NumV(ws.Cells(r, 3).Value)
    If NumV(ws.Cells(r, 4).Value) > 0 Then k = k + 1: a(k) = NumV(ws.Cells(r, 4).Value)
    ReDim res(1 To k)
    For i = 1 To k: res(i) = a(i): Next i
    PolosFila = res
End Function

' completa los polos 2 y 3 segun las fases del detalle de carga (si estan vacios)
Public Sub CompletarPolos(ws As Worksheet, r As Long)
    Dim nf As Long, p As Long
    nf = FasesDet(ws.Cells(r, 5).Value): p = ws.Cells(r, 2).Value
    If nf >= 2 And Len(ws.Cells(r, 3).Formula) = 0 And p + 2 <= 100 Then ws.Cells(r, 3).Value = p + 2
    If nf >= 3 And Len(ws.Cells(r, 4).Formula) = 0 And p + 4 <= 100 Then ws.Cells(r, 4).Value = p + 4
    If nf < 3 And nf > 0 And NumV(ws.Cells(r, 4).Value) > 0 And Len(ws.Cells(r, 4).Formula) > 0 Then ws.Cells(r, 4).ClearContents
    If nf < 2 And nf > 0 And NumV(ws.Cells(r, 3).Value) > 0 And Len(ws.Cells(r, 3).Formula) > 0 Then ws.Cells(r, 3).ClearContents
End Sub

Public Function Ocupacion(ws As Worksheet) As Variant
    Dim oc(1 To 110) As Boolean, r As Long, p As Variant, pp As Variant
    For r = FILA1 To FILAN
        If TieneDatos(ws, r) Then
            pp = PolosFila(ws, r)
            For Each p In pp
                If p >= 1 And p <= 110 Then oc(p) = True
            Next p
        End If
    Next r
    Ocupacion = oc
End Function

' primera posicion libre para un circuito de np polos (p, p+2, p+4)
Public Function PosicionLibre(ws As Worksheet, np As Long) As Long
    Dim oc As Variant, p As Long, k As Long, ok As Boolean
    oc = Ocupacion(ws)
    For p = 1 To 100
        ok = True
        For k = 0 To np - 1
            If p + 2 * k > 100 Then ok = False: Exit For
            If oc(p + 2 * k) Then ok = False: Exit For
        Next k
        If ok Then PosicionLibre = p: Exit Function
    Next p
End Function

Public Sub LimpiarFila(ws As Worksheet, r As Long)
    ' deja la fila como en la plantilla (sin datos y con los valores automaticos)
    ws.Range("A" & r).ClearContents
    ws.Range("C" & r & ":V" & r).Formula = ModAuto.Plantilla().Range("C" & r & ":V" & r).Formula
End Sub

' detalle de carga para el circuito que alimenta a un tablero (o a un transformador si cambia el voltaje)
Public Function DetalleTablero(fasesHijo As Long, vHijo As Double, vPadre As Double) As Variant
    Dim t As Variant, i As Long, f As Long
    t = ThisWorkbook.Names("T_DET").RefersToRange.Value
    f = IIf(fasesHijo = 3, 3, 2)
    If vPadre <> vHijo Then
        For i = 1 To UBound(t, 1)
            If CStr(t(i, 3)) Like "*transformador*" And NumV(t(i, 4)) = vPadre Then DetalleTablero = t(i, 1): Exit Function
        Next i
    End If
    For i = 1 To UBound(t, 1)
        If CStr(t(i, 3)) Like "*tablero*" And NumV(t(i, 5)) = f And NumV(t(i, 4)) = vPadre Then DetalleTablero = t(i, 1): Exit Function
    Next i
    For i = 1 To UBound(t, 1)
        If CStr(t(i, 3)) Like "*tablero*" And NumV(t(i, 4)) = vPadre Then DetalleTablero = t(i, 1): Exit Function
    Next i
    For i = 1 To UBound(t, 1)
        If CStr(t(i, 3)) Like "*tablero*" And NumV(t(i, 5)) = f Then DetalleTablero = t(i, 1): Exit Function
    Next i
    DetalleTablero = t(1, 1)
End Function

' ---------------- crear / eliminar / renombrar ----------------
Public Sub NuevoTablero()
    Dim id As String
    id = Trim$(InputBox("Nombre o ID del tablero (por ejemplo 2G2, PRINCIPAL, UPS-1):", "Nuevo tablero"))
    If id = "" Then Exit Sub
    If Not CrearTablero(id) Is Nothing Then
        HojaMC(id).Activate
        HojaMC(id).Range("C4").Select
    End If
End Sub

Public Function CrearTablero(id As String, Optional tipo As String = "3F", Optional sistema As String = "120/208", Optional silencio As Boolean = False) As Worksheet
    Dim mc As Worksheet, vs As Worksheet, ult As Worksheet, ev As Boolean, sc As Boolean, calc As Long
    id = Trim$(id)
    If Not IdValido(id) Then
        If Not silencio Then MsgBox "Nombre no valido (maximo 27 caracteres, sin [ ] : * ? / \ ').", vbExclamation
        Exit Function
    End If
    If Not HojaMC(id) Is Nothing Then
        If Not silencio Then MsgBox "Ya existe el tablero " & id & ".", vbExclamation
        Set CrearTablero = HojaMC(id)
        Exit Function
    End If
    ev = Application.EnableEvents: sc = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    calc = Application.Calculation: Application.Calculation = xlCalculationManual
    Dim w As Worksheet
    For Each w In ThisWorkbook.Worksheets
        If w.Visible = xlSheetVisible Then Set ult = w
    Next w
    ThisWorkbook.Worksheets("MC_MACHOTE").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("VISTA_MACHOTE").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("MC_MACHOTE").Copy After:=ult
    Set mc = ThisWorkbook.Worksheets("MC_MACHOTE (2)")
    mc.Name = "MC " & id
    ThisWorkbook.Worksheets("VISTA_MACHOTE").Copy After:=mc
    Set vs = ThisWorkbook.Worksheets("VISTA_MACHOTE (2)")
    vs.Name = tipo & " " & id
    ThisWorkbook.Worksheets("MC_MACHOTE").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("VISTA_MACHOTE").Visible = xlSheetVeryHidden
    mc.Visible = xlSheetVisible: vs.Visible = xlSheetVisible
    vs.Cells.Replace What:="MC_MACHOTE!", Replacement:=RefMC(id), LookAt:=xlPart, MatchCase:=False
    mc.Range("C3").Value = id
    mc.Range("C5").Value = tipo
    mc.Range("C6").Value = sistema
    mc.Range("C10").ClearContents
    mc.Tab.Color = RGB(198, 224, 180)
    vs.Tab.Color = RGB(221, 235, 247)
    Application.Calculation = calc
    AjustarVista mc
    Application.EnableEvents = ev: Application.ScreenUpdating = sc
    Set CrearTablero = mc
End Function

Public Sub EliminarTablero()
    Dim ws As Worksheet, id As String
    Set ws = MCActiva(): If ws Is Nothing Then Exit Sub
    id = IdDe(ws)
    If MsgBox("Eliminar el tablero " & id & " (memoria y cuadro de cargas)?" & vbCrLf & "Los tableros que alimenta quedan sin 'Alimentado desde'.", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    BorrarTablero id
    ActualizarProyecto
End Sub

Public Sub BorrarTablero(id As String)
    Dim ws As Worksheet, r As Long, o As Variant
    Application.EnableEvents = False
    For Each o In Tableros()
        Set ws = o
        If IdDe(ws) <> id Then
            For r = FILA1 To FILAN
                If Trim$(CStr(ws.Cells(r, 19).Value)) = id Or Trim$(CStr(ws.Cells(r, 20).Value)) = id Then LimpiarFila ws, r
            Next r
            If Trim$(CStr(ws.Range("C10").Value)) = id Then ws.Range("C10").ClearContents
            If Trim$(CStr(ws.Range("I17").Value)) = id Then ws.Range("I17").ClearContents
            FormulasEnlace ws
        End If
    Next o
    Application.DisplayAlerts = False
    If Not HojaVista(id) Is Nothing Then HojaVista(id).Delete
    If Not HojaMC(id) Is Nothing Then HojaMC(id).Delete
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Public Function MCActiva() As Worksheet
    Dim ws As Worksheet
    If TypeName(ActiveSheet) <> "Worksheet" Then GoTo nada
    Set ws = ActiveSheet
    If EsMC(ws) Then Set MCActiva = ws: Exit Function
    If Left$(ws.Name, 3) = "3F " Or Left$(ws.Name, 3) = "1F " Then
        Set MCActiva = HojaMC(Mid$(ws.Name, 4))
        If Not MCActiva Is Nothing Then Exit Function
    End If
nada:
    MsgBox "Active primero la hoja MC o el cuadro (3F/1F) de un tablero.", vbInformation
End Function

' el usuario cambio el ID (C3) o el tipo (C5) en la hoja MC
Public Sub RenombrarTablero(mc As Worksheet)
    Dim viejo As String, nuevo As String, vs As Worksheet, o As Variant, ws As Worksheet, r As Long, tipo As String
    viejo = Mid$(mc.Name, 4): nuevo = IdDe(mc)
    tipo = IIf(CStr(mc.Range("C5").Value) = "1F", "1F", "3F")
    Set vs = HojaVista(viejo)
    If nuevo = viejo Then
        If Not vs Is Nothing Then
            If Left$(vs.Name, 2) <> tipo Then vs.Name = tipo & " " & nuevo
        End If
        AjustarVista mc
        Exit Sub
    End If
    If Not IdValido(nuevo) Or (Not HojaMC(nuevo) Is mc) Then
        MsgBox "Nombre no valido o repetido: " & nuevo, vbExclamation
        Application.EnableEvents = False: mc.Range("C3").Value = viejo: Application.EnableEvents = True
        Exit Sub
    End If
    Application.EnableEvents = False
    mc.Name = "MC " & nuevo
    If Not vs Is Nothing Then vs.Name = tipo & " " & nuevo
    For Each o In Tableros()
        Set ws = o
        If Trim$(CStr(ws.Range("C10").Value)) = viejo Then ws.Range("C10").Value = nuevo
        If Trim$(CStr(ws.Range("I17").Value)) = viejo Then ws.Range("I17").Value = nuevo
        For r = FILA1 To FILAN
            If Trim$(CStr(ws.Cells(r, 19).Value)) = viejo Then ws.Cells(r, 19).Value = nuevo
            If Trim$(CStr(ws.Cells(r, 20).Value)) = viejo Then ws.Cells(r, 20).Value = nuevo
        Next r
    Next o
    With ThisWorkbook.Worksheets("PREVISTA")
        For r = 5 To 54
            If Trim$(CStr(.Cells(r, 9).Value)) = viejo Then .Cells(r, 9).Value = nuevo
        Next r
    End With
    Application.EnableEvents = True
    ActualizarLista
End Sub

' ---------------- conexiones (cascada) ----------------
Private Function EsDescendiente(posible As String, id As String) As Boolean
    ' True si 'posible' se alimenta (directa o indirectamente) de 'id'
    Dim ws As Worksheet, k As Long, actual As String
    actual = posible
    Do While actual <> "" And k < 100
        If actual = id Then EsDescendiente = True: Exit Function
        Set ws = HojaMC(actual)
        If ws Is Nothing Then Exit Function
        actual = Trim$(CStr(ws.Range("C10").Value))
        k = k + 1
    Loop
End Function

Public Sub ConectarTablero(hijo As Worksheet, Optional silencio As Boolean = False)
    Dim id As String, padreId As String, padre As Worksheet, o As Variant, ws As Worksheet, r As Long, fila As Long, p As Long, det As Variant, np As Long
    Dim ev As Boolean
    ev = Application.EnableEvents: Application.EnableEvents = False
    id = IdDe(hijo): padreId = Trim$(CStr(hijo.Range("C10").Value))
    If padreId <> "" Then
        Set padre = HojaMC(padreId)
        If padre Is Nothing Or padreId = id Or EsDescendiente(padreId, id) Then
            If Not silencio Then MsgBox "No se puede alimentar " & id & " desde " & padreId & " (no existe o es el mismo tablero o uno de sus derivados).", vbExclamation
            hijo.Range("C10").ClearContents: padreId = "": Set padre = Nothing
        End If
    End If
    ' quitar el circuito de los otros tableros
    For Each o In Tableros()
        Set ws = o
        If Not ws Is padre Then
            For r = FILA1 To FILAN
                If Trim$(CStr(ws.Cells(r, 19).Value)) = id Then LimpiarFila ws, r
            Next r
        End If
    Next o
    If Not padre Is Nothing Then
        fila = 0
        For r = FILA1 To FILAN
            If Trim$(CStr(padre.Cells(r, 19).Value)) = id Then fila = r: Exit For
        Next r
        If fila = 0 Then
            det = DetalleTablero(IIf(hijo.Range("C5").Value = "1F", 1, 3), NumV(hijo.Range("C7").Value), NumV(padre.Range("C7").Value))
            np = FasesDet(det): If np < 2 Then np = 2
            p = PosicionLibre(padre, np)
            If p = 0 Then
                MsgBox "No hay espacio libre en " & padreId & " para el circuito de " & id & ".", vbExclamation
            Else
                fila = FilaPos(p)
                LimpiarFila padre, fila
                padre.Cells(fila, 5).Value = det
                CompletarPolos padre, fila
            End If
        End If
        If fila > 0 Then
            SetF padre.Cells(fila, 6), "=" & RefMC(id) & "$BX$2"
            SetF padre.Cells(fila, 7), "=IF(CARGA_DERIV=""demandada""," & RefMC(id) & "$BX$14," & RefMC(id) & "$BX$10)"
            SetF padre.Cells(fila, 8), "=" & RefMC(id) & "$C$12"
            If CStr(padre.Cells(fila, 19).Value) <> id Then padre.Cells(fila, 19).Value = id
            If CStr(padre.Cells(fila, 22).Value) <> "auto" Then padre.Cells(fila, 22).Value = "auto"
        End If
    End If
    SincronizarBypass hijo
    FormulasEnlace hijo
    Application.EnableEvents = ev
End Sub

' referencias directas al tablero padre y al de origen del bypass (sin INDIRECT, que recalcula todo el libro)
Public Sub FormulasEnlace(ws As Worksheet)
    Dim p As String, o As String, R0 As String
    p = Trim$(CStr(ws.Range("C10").Value)): o = Trim$(CStr(ws.Range("I17").Value))
    If p <> "" And p <> IdDe(ws) And Not HojaMC(p) Is Nothing Then
        R0 = RefMC(p)
        SetF ws.Range("BX105"), "=" & R0 & "$C$7"
        SetF ws.Range("BX106"), "=" & R0 & "$BX$49"
        SetF ws.Range("BX107"), "=" & R0 & "$BX$56"
        SetF ws.Range("BX108"), "=" & R0 & "$BX$2"
    Else
        SetF ws.Range("BX105"), "=""""": SetF ws.Range("BX106"), "=""""": SetF ws.Range("BX107"), "=""""": SetF ws.Range("BX108"), "="""""
    End If
    If o <> "" And o <> IdDe(ws) And Not HojaMC(o) Is Nothing Then
        R0 = RefMC(o)
        SetF ws.Range("BX109"), "=" & R0 & "$BX$55"
        SetF ws.Range("BX110"), "=" & R0 & "$BX$49"
    Else
        SetF ws.Range("BX109"), "=""""": SetF ws.Range("BX110"), "="""""
    End If
End Sub

' bypass / segunda acometida desde otro tablero: circuito de respaldo en el tablero de origen (no suma carga)
Public Sub SincronizarBypass(hijo As Worksheet)
    Dim id As String, origenId As String, origen As Worksheet, o As Variant, ws As Worksheet, r As Long, fila As Long, p As Long, det As Variant, np As Long
    id = IdDe(hijo)
    If hijo.Range("I14").Value = "SI" And hijo.Range("I16").Value = "Tablero" Then origenId = Trim$(CStr(hijo.Range("I17").Value))
    If origenId <> "" Then Set origen = HojaMC(origenId)
    For Each o In Tableros()
        Set ws = o
        If Not ws Is origen Then
            For r = FILA1 To FILAN
                If Trim$(CStr(ws.Cells(r, 20).Value)) = id Then LimpiarFila ws, r
            Next r
        End If
    Next o
    If origen Is Nothing Or origenId = id Then Exit Sub
    For r = FILA1 To FILAN
        If Trim$(CStr(origen.Cells(r, 20).Value)) = id Then fila = r: Exit For
    Next r
    If fila = 0 Then
        det = DetalleTablero(IIf(hijo.Range("C5").Value = "1F", 1, 3), NumV(hijo.Range("C7").Value), NumV(origen.Range("C7").Value))
        np = FasesDet(det): If np < 2 Then np = 2
        p = PosicionLibre(origen, np)
        If p = 0 Then Exit Sub
        fila = FilaPos(p)
        LimpiarFila origen, fila
        origen.Cells(fila, 5).Value = det
        CompletarPolos origen, fila
    End If
    SetF origen.Cells(fila, 6), "=""BYPASS ""&" & RefMC(id) & "$BX$2"
    SetF origen.Cells(fila, 7), "=" & RefMC(id) & "$BX$14"
    SetF origen.Cells(fila, 8), "=N(" & RefMC(id) & "$I$20)"
    If CStr(origen.Cells(fila, 20).Value) <> id Then origen.Cells(fila, 20).Value = id
    If CStr(origen.Cells(fila, 22).Value) <> "auto" Then origen.Cells(fila, 22).Value = "auto"
End Sub

' ---------------- vista (cuadro de cargas) ----------------
Public Sub AjustarVista(mc As Worksheet)
    Dim vs As Worksheet, esp As Long, y As Long, p As Long, id As String
    id = IdDe(mc)
    Set vs = HojaVista(id): If vs Is Nothing Then Exit Sub
    On Error Resume Next
    esp = NumV(mc.Range("BX67").Value)
    If esp < NumV(mc.Range("BX61").Value) Then esp = NumV(mc.Range("BX61").Value)
    If esp = 0 Then esp = 42
    If esp > 100 Then esp = 100
    Dim ocultar As Range, mostrar As Range, su As Boolean
    For y = 13 To 112
        p = vs.Cells(y, 39).Value       ' columna AM = posicion
        If (p > esp) <> vs.Rows(y).Hidden Then
            If p > esp Then
                If ocultar Is Nothing Then Set ocultar = vs.Rows(y) Else Set ocultar = Union(ocultar, vs.Rows(y))
            Else
                If mostrar Is Nothing Then Set mostrar = vs.Rows(y) Else Set mostrar = Union(mostrar, vs.Rows(y))
            End If
        End If
    Next y
    su = Application.ScreenUpdating: Application.ScreenUpdating = False
    If Not ocultar Is Nothing Then ocultar.EntireRow.Hidden = True
    If Not mostrar Is Nothing Then mostrar.EntireRow.Hidden = False
    If vs.Columns("AI:AJ").Hidden <> (mc.Range("C8").Value < 3) Then vs.Columns("AI:AJ").Hidden = (mc.Range("C8").Value < 3)
    Application.ScreenUpdating = su
    If vs.PageSetup.PrintArea <> "$A$1:$AJ$114" Then vs.PageSetup.PrintArea = "$A$1:$AJ$114"
End Sub

' ---------------- proyecto: lista, tablas resumen y nombres ----------------
Public Function Orden() As Collection
    ' tableros en orden de arbol (raices primero y luego sus derivados)
    Dim c As New Collection, todos As Collection, o As Variant, ws As Worksheet
    Set todos = Tableros()
    For Each o In todos
        Set ws = o
        If Trim$(CStr(ws.Range("C10").Value)) = "" Or HojaMC(Trim$(CStr(ws.Range("C10").Value))) Is Nothing Then Agregar c, todos, ws
    Next o
    For Each o In todos
        If Not Esta(c, o) Then c.Add o
    Next o
    Set Orden = c
End Function

Private Function Esta(c As Collection, ws As Variant) As Boolean
    Dim o As Variant
    For Each o In c
        If o Is ws Then Esta = True: Exit Function
    Next o
End Function

Private Sub Agregar(c As Collection, todos As Collection, ws As Worksheet)
    Dim o As Variant, h As Worksheet
    If Esta(c, ws) Then Exit Sub
    c.Add ws
    For Each o In todos
        Set h = o
        If Trim$(CStr(h.Range("C10").Value)) = IdDe(ws) Then Agregar c, todos, h
    Next o
End Sub

Public Sub ActualizarLista()
    Dim pj As Worksheet, c As Collection, o As Variant, ws As Worksheet, r As Long, id As String, R0 As String, ev As Boolean
    Dim excl As Object, ordn As Object
    ev = Application.EnableEvents: Application.EnableEvents = False
    Set pj = ThisWorkbook.Worksheets("PROYECTO")
    Set excl = CreateObject("Scripting.Dictionary"): Set ordn = CreateObject("Scripting.Dictionary")
    For r = 17 To 116
        id = Trim$(CStr(pj.Cells(r, 2).Value))
        If id <> "" Then excl(id) = pj.Cells(r, 16).Value: ordn(id) = pj.Cells(r, 15).Value
    Next r
    pj.Range("B17:P116").Clear
    Set c = Orden()
    r = 17
    For Each o In c
        Set ws = o: id = IdDe(ws): R0 = RefMC(id)
        pj.Cells(r, 2).Value = id
        pj.Hyperlinks.Add Anchor:=pj.Cells(r, 2), Address:="", SubAddress:=R0 & "A1", TextToDisplay:=id
        pj.Cells(r, 3).Value = ws.Range("C5").Value
        pj.Cells(r, 4).Value = ws.Range("C6").Value
        pj.Cells(r, 5).Value = ws.Range("C10").Value
        pj.Cells(r, 6).Value = ws.Range("C12").Value
        pj.Cells(r, 7).Formula = "=" & R0 & "$BX$10"
        pj.Cells(r, 8).Formula = "=" & R0 & "$BX$14"
        pj.Cells(r, 9).Formula = "=" & R0 & "$BX$24"
        pj.Cells(r, 10).Formula = "=" & R0 & "$BX$49"
        pj.Cells(r, 11).Formula = "=" & R0 & "$BX$51"
        pj.Cells(r, 12).Formula = "=" & R0 & "$BX$58"
        pj.Cells(r, 13).Formula = "=" & R0 & "$BX$63&"""""
        pj.Cells(r, 14).Formula = "=" & R0 & "$BX$97&"""""
        If ordn.Exists(id) Then pj.Cells(r, 15).Value = ordn(id)
        If excl.Exists(id) Then pj.Cells(r, 16).Value = excl(id)
        r = r + 1
    Next o
    With pj.Range("B16:P" & Application.WorksheetFunction.Max(17, r - 1))
        .Borders.LineStyle = xlContinuous
        .Font.Name = "Century Gothic"
    End With
    pj.Range("B17:P116").Font.Size = 9
    pj.Range("G17:K116").NumberFormat = "0.00"
    pj.Range("L17:L116").NumberFormat = "0.00"
    pj.Range("O17:P" & Application.WorksheetFunction.Max(17, r - 1)).Interior.Color = RGB(230, 244, 225)
    pj.Range("C17:F" & Application.WorksheetFunction.Max(17, r - 1)).Interior.Color = RGB(230, 244, 225)
    On Error Resume Next
    pj.Range("C17:C116").Validation.Delete: pj.Range("C17:C116").Validation.Add Type:=xlValidateList, Formula1:="=L_TIPO"
    pj.Range("D17:D116").Validation.Delete: pj.Range("D17:D116").Validation.Add Type:=xlValidateList, Formula1:="=L_SISTEMAS"
    pj.Range("E17:E116").Validation.Delete: pj.Range("E17:E116").Validation.Add Type:=xlValidateList, Formula1:="=LISTA_TABLEROS"
    pj.Range("E17:E116").Validation.ShowError = False
    On Error GoTo 0
    Application.EnableEvents = ev
End Sub

' boton "Actualizar proyecto": conexiones, calculo, lista, vistas, tablas resumen y nombres
Public Sub ActualizarProyecto()
    Dim o As Variant, ws As Worksheet, sc As Boolean
    sc = Application.ScreenUpdating
    Congelar
    ModUDF.RecargarCatalogo
    For Each o In Tableros()
        Set ws = o
        ConectarTablero ws, True
    Next o
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.Calculation = xlCalculationManual
    ActualizarLista
    For Each o In Tableros()
        Set ws = o
        AjustarVista ws
    Next o
    ActualizarResumen
    NombresDefinidos
    Descongelar
End Sub

Public Function ListaResumen() As Collection
    ' tableros para las tablas resumen: orden de la columna "Orden resumen" (vacio = orden del arbol), sin los excluidos
    Dim pj As Worksheet, r As Long, c As New Collection, i As Long, j As Long, n As Long
    Dim ids() As String, ords() As Double, t As String, td As Double
    Set pj = ThisWorkbook.Worksheets("PROYECTO")
    ReDim ids(1 To 100): ReDim ords(1 To 100)
    For r = 17 To 116
        If Trim$(CStr(pj.Cells(r, 2).Value)) <> "" And UCase$(Trim$(CStr(pj.Cells(r, 16).Value))) <> "X" Then
            n = n + 1: ids(n) = Trim$(CStr(pj.Cells(r, 2).Value))
            If IsNumeric(pj.Cells(r, 15).Value) And Not IsEmpty(pj.Cells(r, 15).Value) Then ords(n) = pj.Cells(r, 15).Value Else ords(n) = 1000 + r
        End If
    Next r
    For i = 1 To n - 1
        For j = i + 1 To n
            If ords(j) < ords(i) Then t = ids(i): ids(i) = ids(j): ids(j) = t: td = ords(i): ords(i) = ords(j): ords(j) = td
        Next j
    Next i
    For i = 1 To n
        If Not HojaMC(ids(i)) Is Nothing Then c.Add HojaMC(ids(i))
    Next i
    Set ListaResumen = c
End Function

Public Sub ActualizarResumen()
    Dim tr As Worksheet, c As Collection, o As Variant, ws As Worksheet, r As Long, R0 As String, i As Long
    Dim cols As Variant, refs As Variant, fm As String
    Set tr = ThisWorkbook.Worksheets("TABLA RESUMEN")
    tr.Range("B11:AU300").UnMerge
    tr.Range("B11:AU300").Clear
    Set c = ListaResumen()
    cols = Array("B", "C", "D", "E", "F", "G", "H", "I", "K", "M", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU")
    refs = Array("$BX$2", "$BX$98", "$BX$10", "$BX$14", "$BX$12", "$BX$13", "$BX$15", "$BX$99", "$BX$100", "$BX$101", "$F$3", "$F$4", "$BX$30", "$BX$36", "$F$5", "$C$12", "$BX$49", "$BX$50", "$BX$51", "$BX$58", "$BX$24", _
                 "$BX$2", "$BX$63", "$BX$60", "$BX$64", "$BX$65", "$BX$66", "$BX$67", "$C$18", "$BX$60", "$BX$79", "$BX$82", "$BX$80", "$BX$81", "$C$6", "$C$8", "$BX$60", "$BX$72", "$BX$75", "$BX$24", "$BX$73", "$BX$76", "$BX$74")
    r = 11
    For Each o In c
        Set ws = o: R0 = RefMC(IdDe(ws))
        For i = 0 To UBound(cols)
            tr.Range(cols(i) & r).Formula = "=IF(" & R0 & refs(i) & "="""",""""," & R0 & refs(i) & ")"
        Next i
        tr.Range("I" & r & ":J" & r).Merge: tr.Range("K" & r & ":L" & r).Merge: tr.Range("M" & r & ":N" & r).Merge
        r = r + 1
    Next o
    If r > 11 Then
        With tr.Range("B11:AU" & r - 1)
            .Borders.LineStyle = xlContinuous
            .Font.Name = "Century Gothic": .Font.Size = 12: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
            .RowHeight = 30
        End With
        tr.Range("D11:H" & r - 1).NumberFormat = "0.00"
        tr.Range("U11:X" & r - 1).NumberFormat = "0.00"
    End If
    tr.PageSetup.PrintArea = "$B$1:$AU$" & Application.WorksheetFunction.Max(11, r - 1)
    ResumenVertical c
    ResumenDU c
End Sub

Private Sub ResumenVertical(c As Collection)
    Dim w As Worksheet, campos As Variant, refs As Variant, i As Long, j As Long, o As Variant, ws As Worksheet, R0 As String
    Set w = ThisWorkbook.Worksheets("RESUMEN VERTICAL")
    w.Range("B3:ZZ80").Clear
    campos = Array("Tablero / Equipo", "Alimentado desde", "kVA totales", "kVA demandados", "Factor demanda", "Factor diversidad", "Factor potencia", "Fases (AWG)", "Neutro (AWG)", "Tierra (AWG)", "Material", "Aislamiento", "Tuberia (mm)", _
                   "Longitud (m)", "Voltaje bornes (V)", "Caida total (V)", "Caida total (%)", "Icc disponible (kA)", "Proteccion (A)", "Tablero: modelo", "Fabricante", "Barras F/N/T (A)", "Espacios", "Montaje", _
                   "SPD: modelo", "SPD kA L-L / L-N", "Principal: modelo", "Principal: marco (A)", "Principal: unidad", "Principal: polos", "Principal: SCCR (kA)")
    refs = Array("$BX$2", "$BX$98", "$BX$10", "$BX$14", "$BX$12", "$BX$13", "$BX$15", "$BX$99", "$BX$100", "$BX$101", "$F$3", "$F$4", "$BX$102", "$C$12", "$BX$49", "$BX$50", "$BX$51", "$BX$58", "$BX$24", "$BX$63", "$BX$60", _
                 "#BARRAS", "$BX$67", "$C$18", "$BX$79", "#SPD", "$BX$72", "$BX$75", "$BX$73", "$BX$76", "$BX$74")
    For i = 0 To UBound(campos)
        w.Cells(3 + i, 2).Value = campos(i)
    Next i
    j = 3
    For Each o In c
        Set ws = o: R0 = RefMC(IdDe(ws))
        For i = 0 To UBound(refs)
            If refs(i) = "#BARRAS" Then
                w.Cells(3 + i, j).Formula = "=" & R0 & "$BX$64&"" / ""&" & R0 & "$BX$65&"" / ""&" & R0 & "$BX$66"
            ElseIf refs(i) = "#SPD" Then
                w.Cells(3 + i, j).Formula = "=" & R0 & "$BX$80&"" / ""&" & R0 & "$BX$81"
            Else
                w.Cells(3 + i, j).Formula = "=IF(" & R0 & refs(i) & "="""",""""," & R0 & refs(i) & ")"
            End If
        Next i
        w.Columns(j).ColumnWidth = 24
        j = j + 1
    Next o
    With w.Range(w.Cells(3, 2), w.Cells(3 + UBound(campos), Application.WorksheetFunction.Max(3, j - 1)))
        .Borders.LineStyle = xlContinuous: .Font.Name = "Century Gothic": .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
    End With
    w.Range(w.Cells(3, 2), w.Cells(3 + UBound(campos), 2)).Interior.Color = RGB(217, 217, 217)
    w.Range(w.Cells(3, 2), w.Cells(3 + UBound(campos), 2)).Font.Bold = True
    w.Range(w.Cells(5, 3), w.Cells(9, 200)).NumberFormat = "0.00"
    w.Range(w.Cells(17, 3), w.Cells(20, 200)).NumberFormat = "0.00"
    w.Range(w.Cells(3, 3), w.Cells(3, 200)).Font.Bold = True
End Sub

Private Sub ResumenDU(c As Collection)
    Dim w As Worksheet, cab As Variant, i As Long, r As Long, o As Variant, ws As Worksheet, R0 As String, f As Variant
    Set w = ThisWorkbook.Worksheets("RESUMEN DU")
    w.Range("B3:P300").Clear
    cab = Array("Tablero", "Alimentado desde", "kVA demandados", "Proteccion (A)", "Conductores", "Tuberia", "Long. (m)", "Caida total (%)", "Icc (kA)", "Segunda acometida", "Conductores 2a acometida", "Caida 2a acometida (%)", "Transformador / UPS", "Avisos")
    For i = 0 To UBound(cab)
        w.Cells(3, 2 + i).Value = cab(i)
    Next i
    r = 4
    For Each o In c
        Set ws = o: R0 = RefMC(IdDe(ws))
        f = Array("=" & R0 & "$BX$2", "=" & R0 & "$BX$98&""""", "=" & R0 & "$BX$14", "=" & R0 & "$BX$24&""""", "=" & R0 & "$BX$99&"" ""&" & R0 & "$BX$100&"" N ""&" & R0 & "$BX$101&"" T""", "=" & R0 & "$BX$102&""""", _
                  "=" & R0 & "$C$12", "=" & R0 & "$BX$51", "=" & R0 & "$BX$58", _
                  "=IF(" & R0 & "$I$14=""SI""," & R0 & "$I$15&"" - ""&IF(" & R0 & "$I$16=""Generador"",""Generador ""&" & R0 & "$I$18&"" kVA"",""Bypass desde ""&" & R0 & "$I$17),"""")", _
                  "=" & R0 & "$BX$95&""""", "=IF(" & R0 & "$BX$93="""",""""," & R0 & "$BX$93)", _
                  "=TRIM(IF(" & R0 & "$I$3=""SI"",""Transformador ""&" & R0 & "$I$4&"" kVA Z ""&" & R0 & "$I$5&"" % "","""")&IF(" & R0 & "$I$10=""SI"",""UPS ""&" & R0 & "$I$11&"" kVA"",""""))", _
                  "=" & R0 & "$BX$97&""""")
        For i = 0 To UBound(f)
            w.Cells(r, 2 + i).Formula = f(i)
        Next i
        r = r + 1
    Next o
    With w.Range(w.Cells(3, 2), w.Cells(Application.WorksheetFunction.Max(4, r - 1), 2 + UBound(cab)))
        .Borders.LineStyle = xlContinuous: .Font.Name = "Century Gothic": .Font.Size = 10
        .VerticalAlignment = xlCenter: .WrapText = True
    End With
    With w.Range(w.Cells(3, 2), w.Cells(3, 2 + UBound(cab)))
        .Interior.Color = RGB(217, 217, 217): .Font.Bold = True: .HorizontalAlignment = xlCenter
    End With
    w.Range("D4:D300").NumberFormat = "0.00": w.Range("I4:J300").NumberFormat = "0.00": w.Range("M4:M300").NumberFormat = "0.00"
    w.Columns("C:O").ColumnWidth = 18: w.Columns("P").ColumnWidth = 60
End Sub

Private Function Limpio(s As String) As String
    Dim i As Long, ch As String, r As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9_]" Then r = r & ch Else r = r & "_"
    Next i
    If r Like "#*" Then r = "_" & r
    Limpio = r
End Function

' nombres definidos para vincular en Revit: TABLERO_<id> (cuadro de cargas) y las tablas resumen
Public Sub NombresDefinidos()
    Dim nm As Name, o As Variant, ws As Worksheet, vs As Worksheet, u As Long, n As Long
    On Error Resume Next
    For Each nm In ThisWorkbook.Names
        If Left$(nm.Name, 8) = "TABLERO_" Then nm.Delete
    Next nm
    For Each o In Tableros()
        Set ws = o: Set vs = HojaVista(IdDe(ws))
        If Not vs Is Nothing Then
            u = 114
            Do While u > 13 And vs.Rows(u).Hidden
                u = u - 1
            Loop
            ThisWorkbook.Names.Add Name:="TABLERO_" & Limpio(IdDe(ws)), RefersTo:="='" & vs.Name & "'!$B$1:$AJ$114"
        End If
    Next o
    n = ListaResumen().Count
    u = Application.WorksheetFunction.Max(11, 10 + n)
    ThisWorkbook.Names.Add Name:="TABLA_RESUMEN__TABLEROS_ELÉCTRICOS", RefersTo:="='TABLA RESUMEN'!$B$1:$Y$" & u
    ThisWorkbook.Names.Add Name:="DATOS_DEL_TABLERO", RefersTo:="='TABLA RESUMEN'!$Z$1:$AG$" & u
    ThisWorkbook.Names.Add Name:="DATOS_DEL_SUPRESOR", RefersTo:="='TABLA RESUMEN'!$AH$1:$AN$" & u
    ThisWorkbook.Names.Add Name:="DATOS_INTERRUPTOR_PRINCIPAL", RefersTo:="='TABLA RESUMEN'!$AO$1:$AU$" & u
    ThisWorkbook.Names.Add Name:="RESUMEN_VERTICAL", RefersTo:="='RESUMEN VERTICAL'!$B$1:$" & Split(ThisWorkbook.Worksheets("RESUMEN VERTICAL").Cells(1, 2 + n).Address, "$")(1) & "$33"
    ThisWorkbook.Names.Add Name:="RESUMEN_DU", RefersTo:="='RESUMEN DU'!$B$1:$O$" & (3 + n)
End Sub

Public Sub SoloTablaResumen()
    ' alterna: mostrar solo la tabla resumen (sin datos del tablero, del supresor ni del interruptor principal)
    With ThisWorkbook.Worksheets("TABLA RESUMEN")
        .Columns("Z:AU").Hidden = Not .Columns("Z:AU").Hidden
        If .Columns("Z:AU").Hidden Then .PageSetup.PrintArea = "$B$1:$Y$" & Application.WorksheetFunction.Max(11, 10 + ListaResumen().Count) _
        Else .PageSetup.PrintArea = "$B$1:$AU$" & Application.WorksheetFunction.Max(11, 10 + ListaResumen().Count)
    End With
End Sub

' ---------------- PDF y exportacion para Revit ----------------
Public Sub ExportarPDF(Optional destino As String = "", Optional conMemorias As Boolean = False)
    Dim ruta As Variant, hojas() As String, n As Long, o As Variant, ws As Worksheet, vs As Worksheet, conMC As VbMsgBoxResult
    ActualizarProyecto
    If destino = "" Then
        conMC = MsgBox("Incluir las memorias de calculo (hojas MC)?", vbYesNoCancel + vbQuestion, "Exportar PDF")
        If conMC = vbCancel Then Exit Sub
    Else
        conMC = IIf(conMemorias, vbYes, vbNo)
    End If
    ruta = destino
    If ruta = "" Then ruta = Application.GetSaveAsFilename(InitialFileName:=Replace(ThisWorkbook.Name, ".xlsm", "") & " - tableros.pdf", FileFilter:="PDF (*.pdf), *.pdf")
    If ruta = False Then Exit Sub
    ReDim hojas(1 To 300)
    n = n + 1: hojas(n) = "TABLA RESUMEN"
    For Each o In Orden()
        Set ws = o: Set vs = HojaVista(IdDe(ws))
        If Not vs Is Nothing Then n = n + 1: hojas(n) = vs.Name
        If conMC = vbYes Then
            PrepararImpresionMC ws
            n = n + 1: hojas(n) = ws.Name
        End If
    Next o
    ReDim Preserve hojas(1 To n)
    Pagina ThisWorkbook.Worksheets("TABLA RESUMEN"), ThisWorkbook.Worksheets("TABLA RESUMEN").PageSetup.PrintArea, False, False
    For Each o In Orden()
        Set ws = o: Set vs = HojaVista(IdDe(ws))
        If Not vs Is Nothing Then Pagina vs, "$A$1:$AJ$114", 1, False
    Next o
    ThisWorkbook.Worksheets(hojas).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ruta, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=(destino = "")
    ThisWorkbook.Worksheets("INICIO").Select
    If conMC = vbYes Then
        For Each o In Orden()
            Set ws = o: RestaurarMC ws
        Next o
    End If
End Sub

Public Sub PrepararImpresionMC(ws As Worksheet)
    ' memoria imprimible: datos, panel de resultados, circuitos (columnas principales) y factores de demanda
    ws.Columns("A:CB").Hidden = False
    Dim c As Variant
    For Each c In Split("A:A,I:J,S:V,Y:Y,AA:AA,AD:AE,AG:AO,AU:AV,AX:BB,BF:BG,BI:BI,BK:BK,BN:BP,BV:CB", ",")
        ws.Columns(CStr(c)).Hidden = True
    Next c
    Dim r As Long, oc As Range
    For r = FILA1 To FILAN
        If Not TieneDatos(ws, r) Then
            If oc Is Nothing Then Set oc = ws.Rows(r) Else Set oc = Union(oc, ws.Rows(r))
        End If
    Next r
    If Not oc Is Nothing Then oc.EntireRow.Hidden = True
    Pagina ws, "$B$1:$BU$146", False, False
End Sub

Public Sub RestaurarMC(ws As Worksheet)
    ws.Columns("A:BU").Hidden = False
    ws.Columns("BV").Hidden = True
    ws.Columns("BY:CB").Hidden = True
    ws.Rows(FILA1 & ":" & FILAN).Hidden = False
End Sub

Public Sub Pagina(ws As Worksheet, area As String, alto As Variant, vertical As Boolean)
    On Error Resume Next
    With ws.PageSetup
        .PrintArea = area
        .Orientation = IIf(vertical, xlPortrait, xlLandscape)
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = alto
        .CenterHorizontally = True
        .LeftMargin = Application.CentimetersToPoints(0.8): .RightMargin = Application.CentimetersToPoints(0.8)
        .TopMargin = Application.CentimetersToPoints(1): .BottomMargin = Application.CentimetersToPoints(1.2)
        .HeaderMargin = Application.CentimetersToPoints(0.4): .FooterMargin = Application.CentimetersToPoints(0.5)
        .CenterFooter = "&8&A - p" & ChrW(225) & "gina &P de &N"
    End With
End Sub

' copia de las tablas como valores (con los nombres TABLERO_x) para vincular en Revit
Public Sub ExportarTablasRevit(Optional destino As String = "")
    Dim wb As Workbook, ruta As Variant, o As Variant, ws As Worksheet, vs As Worksheet, nm As Name, h As Worksheet
    ActualizarProyecto
    ruta = destino
    If ruta = "" Then ruta = Application.GetSaveAsFilename(InitialFileName:=Replace(ThisWorkbook.Name, ".xlsm", "") & " - tablas Revit.xlsx", FileFilter:="Excel (*.xlsx), *.xlsx")
    If ruta = False Then Exit Sub
    Application.ScreenUpdating = False
    ThisWorkbook.Worksheets(Array("TABLA RESUMEN", "RESUMEN VERTICAL", "RESUMEN DU")).Copy
    Set wb = ActiveWorkbook
    For Each o In Orden()
        Set ws = o: Set vs = HojaVista(IdDe(ws))
        If Not vs Is Nothing Then vs.Copy After:=wb.Worksheets(wb.Worksheets.Count)
    Next o
    For Each h In wb.Worksheets
        h.UsedRange.Value = h.UsedRange.Value
    Next h
    On Error Resume Next
    For Each nm In wb.Names
        nm.Delete
    Next nm
    On Error GoTo 0
    For Each nm In ThisWorkbook.Names
        If Left$(nm.Name, 8) = "TABLERO_" Or nm.Name Like "TABLA_RESUMEN*" Or nm.Name Like "DATOS_*" Or nm.Name Like "RESUMEN_*" Then
            wb.Names.Add Name:=nm.Name, RefersTo:=nm.RefersTo
        End If
    Next nm
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=ruta, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If destino = "" Then MsgBox "Tablas exportadas: " & ruta, vbInformation
End Sub

Public Sub IrAProyecto()
    ThisWorkbook.Worksheets("PROYECTO").Activate
End Sub
