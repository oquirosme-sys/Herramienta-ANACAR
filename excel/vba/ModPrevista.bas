Attribute VB_Name = "ModPrevista"
' Prevista para diseno esquematico: genera circuitos de prevista en el tablero de cada area (hoja PREVISTA)
' iluminacion hasta 1.5 kVA por circuito, tomas hasta 1.44 kVA y climatizacion/equipos hasta 10 kVA.
Option Explicit

Private Const MAX_ILUM As Double = 1.5
Private Const MAX_TOMA As Double = 1.44
Private Const MAX_CLIMA As Double = 10

Public Sub GenerarPrevista()
    Dim pv As Worksheet, r As Long, n As Long, tot As Long, sinTab As Long
    Set pv = ThisWorkbook.Worksheets("PREVISTA")
    pv.Calculate
    Congelar
    For r = 5 To 54
        QuitarPrevista r
        If Trim$(CStr(pv.Cells(r, 3).Value)) <> "" Then
            If HojaMC(Trim$(CStr(pv.Cells(r, 9).Value))) Is Nothing Then
                sinTab = sinTab + 1
                pv.Cells(r, 18).Value = "Elija el tablero"
            Else
                n = GenerarFila(r)
                pv.Cells(r, 18).Value = n
                tot = tot + n
            End If
        Else
            pv.Cells(r, 18).ClearContents
        End If
    Next r
    ActualizarProyecto
    Application.ScreenUpdating = True
    MsgBox tot & " circuitos de prevista generados." & IIf(sinTab > 0, vbCrLf & sinTab & " area(s) sin tablero.", ""), vbInformation, "Prevista"
End Sub

Private Sub QuitarPrevista(r As Long)
    Dim o As Variant, ws As Worksheet, f As Long
    For Each o In Tableros()
        Set ws = o
        For f = FILA1 To FILAN
            If CStr(ws.Cells(f, 22).Value) = "prevista:" & r Then LimpiarFila ws, f
        Next f
    Next o
End Sub

Private Function DetEquipos(V As Double) As Variant
    Dim t As Variant, i As Long
    t = ThisWorkbook.Names("T_DET").RefersToRange.Value
    For i = 1 To UBound(t, 1)
        If UCase$(CStr(t(i, 3))) Like "EQUIPOS MEC*" Then
            If (V >= 440 And (NumV(t(i, 4)) = 277 Or NumV(t(i, 4)) = 480)) Or (V < 440 And NumV(t(i, 4)) <= 240 And NumV(t(i, 4)) > 0) Then DetEquipos = t(i, 1): Exit Function
        End If
    Next i
    DetEquipos = 55
End Function

Private Function GenerarFila(r As Long) As Long
    Dim pv As Worksheet, ws As Worksheet, V As Double, nom As String, lon As Double
    Set pv = ThisWorkbook.Worksheets("PREVISTA")
    Set ws = HojaMC(Trim$(CStr(pv.Cells(r, 9).Value)))
    V = NumV(ws.Range("C7").Value)
    nom = CStr(pv.Cells(r, 3).Value) & IIf(Trim$(CStr(pv.Cells(r, 2).Value)) <> "", " " & pv.Cells(r, 2).Value, "") & " (" & Format$(NumV(pv.Cells(r, 4).Value), "0.0") & " m" & ChrW(178) & ")"
    lon = NumV(pv.Cells(r, 10).Value): If lon = 0 Then lon = 25
    GenerarFila = Crear(ws, NumV(pv.Cells(r, 14).Value), MAX_ILUM, IIf(V >= 440, 6, 1), "iluminaci" & ChrW(243) & "n", nom, lon, r) + _
                  Crear(ws, NumV(pv.Cells(r, 15).Value), MAX_TOMA, 8, "tomas", nom, lon, r) + _
                  Crear(ws, NumV(pv.Cells(r, 16).Value) - NumV(pv.Cells(r, 8).Value), MAX_CLIMA, DetEquipos(V), "climatizaci" & ChrW(243) & "n", nom, lon, r) + _
                  Crear(ws, NumV(pv.Cells(r, 8).Value), MAX_CLIMA, DetEquipos(V), "equipos", nom, lon, r)
End Function

Private Function Crear(ws As Worksheet, kva As Double, maximo As Double, det As Variant, txt As String, nom As String, lon As Double, r As Long) As Long
    Dim k As Long, i As Long, np As Long, p As Long, f As Long
    If kva <= 0.0000001 Then Exit Function
    k = Application.WorksheetFunction.RoundUp(kva / maximo, 0): If k < 1 Then k = 1
    np = FasesDet(det): If np < 1 Then np = 1
    For i = 1 To k
        p = PosicionLibre(ws, np)
        If p = 0 Then Exit For
        f = FilaPos(p)
        LimpiarFila ws, f
        ws.Cells(f, 5).Value = det
        CompletarPolos ws, f
        ws.Cells(f, 6).Value = "Prevista " & txt & " " & nom & IIf(k > 1, " " & i & "/" & k, "")
        ws.Cells(f, 7).Value = Round(kva / k, 3)
        ws.Cells(f, 8).Value = lon
        ws.Cells(f, 22).Value = "prevista:" & r
        Crear = Crear + 1
    Next i
End Function
