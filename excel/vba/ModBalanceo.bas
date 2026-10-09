Attribute VB_Name = "ModBalanceo"
' Autobalanceo de fases del tablero activo (mismo algoritmo que la herramienta web):
' mueve o intercambia circuitos (no los marcados como Fijo) minimizando la dispersion entre fases
' y anota cada cambio de posicion en la hoja CAMBIOS REVIT para pasarlo al modelo.
Option Explicit

Private Type Circ
    fila As Long
    np As Long
    pol(1 To 3) As Long
    kva As Double
    nf As Double
    fijo As Boolean
    desc As String
    revit As String
End Type

Private cs() As Circ, nC As Long, fasesT As Long

Private Function FaseDe(p As Long) As Long
    Dim k As Long
    k = (p - 1) \ 2
    If fasesT = 3 Then FaseDe = k Mod 3 Else FaseDe = k Mod 2
End Function

Private Sub Cargas(a() As Circ, f() As Double)
    Dim i As Long, j As Long
    f(0) = 0: f(1) = 0: f(2) = 0
    For i = 1 To nC
        For j = 1 To a(i).np
            f(FaseDe(a(i).pol(j))) = f(FaseDe(a(i).pol(j))) + a(i).kva / a(i).nf
        Next j
    Next i
End Sub

Private Function Disp(a() As Circ) As Double
    Dim f(0 To 2) As Double, m As Double, i As Long, nf As Long
    Cargas a, f
    nf = IIf(fasesT = 3, 3, 2)
    For i = 0 To nf - 1: m = m + f(i): Next i
    m = m / nf
    For i = 0 To nf - 1: Disp = Disp + (f(i) - m) ^ 2: Next i
End Function

Private Function Desb(a() As Circ) As Double
    Dim f(0 To 2) As Double, mx As Double, mn As Double, i As Long, nf As Long
    Cargas a, f
    nf = IIf(fasesT = 3, 3, 2)
    mx = f(0): mn = f(0)
    For i = 1 To nf - 1
        If f(i) > mx Then mx = f(i)
        If f(i) < mn Then mn = f(i)
    Next i
    If mx > 0 Then Desb = (mx - mn) / mx * 100
End Function

Private Function PolTxt(c As Circ) As String
    Dim j As Long
    For j = 1 To c.np
        PolTxt = PolTxt & IIf(j > 1, ",", "") & c.pol(j)
    Next j
End Function

Public Sub Autobalancear()
    Dim ws As Worksheet
    Set ws = MCActiva(): If ws Is Nothing Then Exit Sub
    BalancearTablero ws, True
End Sub

Public Sub AutobalancearTodos()
    Dim o As Variant, txt As String, ws As Worksheet
    If MsgBox("Autobalancear todos los tableros y anotar los cambios en CAMBIOS REVIT?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    For Each o In Tableros()
        Set ws = o
        txt = txt & BalancearTablero(ws, False)
    Next o
    ActualizarProyecto
    MsgBox IIf(txt = "", "Sin cambios.", txt), vbInformation, "Autobalanceo"
End Sub

Public Function BalancearTablero(ws As Worksheet, preguntar As Boolean) As String
    Dim r As Long, pp As Variant, j As Long, esp As Long, base() As Circ, it As Long, actual As Double, d As Double
    Dim i As Long, k As Long, p As Long, bestD As Double, bestDist As Long, hay As Boolean, bestA() As Circ, prueba() As Circ, ok As Boolean
    Dim usados(1 To 200) As Boolean, movs As String, antes As Double, maxMov As Long, dist As Long, tmp As Circ
    ws.Calculate
    fasesT = NumV(ws.Range("C8").Value)
    If fasesT < 2 Then BalancearTablero = "": If preguntar Then MsgBox "Tablero monofasico: no hay fases que balancear.", vbInformation
    If fasesT < 2 Then Exit Function
    nC = 0: ReDim cs(1 To 100)
    For r = FILA1 To FILAN
        If (Len(ws.Cells(r, 5).Formula) + Len(ws.Cells(r, 6).Formula) + Len(ws.Cells(r, 7).Formula)) > 0 Then
            nC = nC + 1
            With cs(nC)
                .fila = r
                pp = PolosFila(ws, r)
                .np = UBound(pp)
                For j = 1 To .np: .pol(j) = pp(j): Next j
                .kva = NumV(ws.Cells(r, 24).Value)
                .nf = NumV(ws.Cells(r, 26).Value): If .nf = 0 Then .nf = .np
                .fijo = (UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = "SI")
                .desc = CStr(ws.Cells(r, 23).Value)
                .revit = CStr(ws.Cells(r, 21).Value)
                If .revit = "" Then .revit = PolTxt(cs(nC))
            End With
            For j = 1 To cs(nC).np
                If cs(nC).pol(j) > esp Then esp = cs(nC).pol(j)
            Next j
        End If
    Next r
    If nC = 0 Then Exit Function
    If NumV(ws.Range("BX67").Value) > esp Then esp = NumV(ws.Range("BX67").Value)
    If esp > 100 Then esp = 100
    ReDim Preserve cs(1 To nC)
    base = cs
    antes = Desb(cs)
    maxMov = 30
    For it = 1 To maxMov
        actual = Disp(cs)
        hay = False
        Erase usados
        For i = 1 To nC
            For j = 1 To cs(i).np: usados(cs(i).pol(j)) = True: Next j
        Next i
        For i = 1 To nC
            If cs(i).kva <> 0 And Not cs(i).fijo Then
                ' intercambiar con otro circuito del mismo numero de polos
                For k = i + 1 To nC
                    If Not cs(k).fijo And cs(k).np = cs(i).np Then
                        prueba = cs
                        For j = 1 To 3: prueba(i).pol(j) = cs(k).pol(j): prueba(k).pol(j) = cs(i).pol(j): Next j
                        d = Disp(prueba): dist = Abs(cs(i).pol(1) - cs(k).pol(1))
                        If d < actual - 0.000001 And (Not hay Or d < bestD - 0.000000001 Or (Abs(d - bestD) < 0.000000001 And dist < bestDist)) Then
                            hay = True: bestD = d: bestDist = dist: bestA = prueba
                        End If
                    End If
                Next k
                ' mover a posiciones libres
                For p = 1 To esp
                    ok = (p + 2 * (cs(i).np - 1) <= esp)
                    If ok Then
                        For j = 0 To cs(i).np - 1
                            If usados(p + 2 * j) Then
                                Dim propio As Boolean, q As Long
                                propio = False
                                For q = 1 To cs(i).np
                                    If cs(i).pol(q) = p + 2 * j Then propio = True
                                Next q
                                If Not propio Then ok = False: Exit For
                            End If
                        Next j
                    End If
                    If ok And p <> cs(i).pol(1) Then
                        prueba = cs
                        For j = 1 To cs(i).np: prueba(i).pol(j) = p + 2 * (j - 1): Next j
                        d = Disp(prueba)
                        dist = Abs(cs(i).pol(1) - p) + IIf((p Mod 2) = (cs(i).pol(1) Mod 2), 0, 50)
                        If d < actual - 0.000001 And (Not hay Or d < bestD - 0.000000001 Or (Abs(d - bestD) < 0.000000001 And dist < bestDist)) Then
                            hay = True: bestD = d: bestDist = dist: bestA = prueba
                        End If
                    End If
                Next p
            End If
        Next i
        If Not hay Then Exit For
        cs = bestA
    Next it
    ' cambios
    Dim n As Long
    For i = 1 To nC
        If PolTxt(cs(i)) <> PolTxt(base(i)) Then
            n = n + 1
            movs = movs & vbCrLf & "  [" & PolTxt(base(i)) & "] -> [" & PolTxt(cs(i)) & "] " & cs(i).desc
        End If
    Next i
    If n = 0 Then
        If preguntar Then MsgBox "El tablero ya esta balanceado (" & Format$(antes, "0.00") & " %).", vbInformation
        Exit Function
    End If
    If preguntar Then
        If MsgBox("Desbalance: " & Format$(antes, "0.00") & " % -> " & Format$(Desb(cs), "0.00") & " %" & vbCrLf & "Cambios (" & n & "):" & movs & vbCrLf & vbCrLf & "Aplicar y anotar en CAMBIOS REVIT?", vbYesNo + vbQuestion, "Autobalanceo " & IdDe(ws)) <> vbYes Then Exit Function
    End If
    Congelar
    Aplicar ws, base
    Descongelar
    BalancearTablero = IdDe(ws) & ": " & Format$(antes, "0.00") & " % -> " & Format$(Desb(cs), "0.00") & " % (" & n & " cambios)" & vbCrLf
    If preguntar Then ActualizarProyecto
End Function

Private Sub Aplicar(ws As Worksheet, base() As Circ)
    Dim i As Long, f() As Variant, cm As Worksheet, u As Long, ev As Boolean, nf As Long, explic2 As Boolean, explic3 As Boolean
    ev = Application.EnableEvents: Application.EnableEvents = False
    Set cm = ThisWorkbook.Worksheets("CAMBIOS REVIT")
    ReDim f(1 To nC)
    For i = 1 To nC
        f(i) = ws.Range("A" & base(i).fila & ":V" & base(i).fila).Formula
    Next i
    For i = 1 To nC
        If PolTxt(cs(i)) <> PolTxt(base(i)) Then LimpiarFila ws, base(i).fila
    Next i
    For i = 1 To nC
        If PolTxt(cs(i)) <> PolTxt(base(i)) Then
            Dim r As Long, v As Variant
            r = FilaPos(cs(i).pol(1))
            v = f(i)
            v(1, 2) = cs(i).pol(1)
            If Len(CStr(v(1, 3))) > 0 And cs(i).np >= 2 Then v(1, 3) = cs(i).pol(2)
            If Len(CStr(v(1, 4))) > 0 And cs(i).np >= 3 Then v(1, 4) = cs(i).pol(3)
            ws.Range("A" & r & ":V" & r).Formula = v
            u = cm.Cells(cm.Rows.Count, 1).End(xlUp).Row + 1
            cm.Cells(u, 1).Value = Date
            cm.Cells(u, 2).Value = IdDe(ws)
            cm.Cells(u, 3).Value = cs(i).desc
            cm.Cells(u, 4).Value = "'" & base(i).revit
            cm.Cells(u, 5).Value = "'" & PolTxt(cs(i))
            cm.Cells(u, 6).Value = "balanceo"
        End If
    Next i
    Application.EnableEvents = ev
End Sub
