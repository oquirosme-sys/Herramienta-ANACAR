Attribute VB_Name = "ModUDF"
' Funciones de hoja (UDF) que usan las memorias de calculo "MC <tablero>".
' Replican el motor de la herramienta web (js/calc.js): seleccion de breakers, tableros y supresores
' por reglas de fabricante, clasificacion en serie, ampacidad 310.15, punto a punto y GFCI/AFCI NEC 2020.
Option Explicit
Option Compare Text

Private cBK As Variant, cTB As Variant, cSPD As Variant, cRT As Variant, cRB As Variant, cSE As Variant
Private cAMP As Variant, cTEMP As Variant, cAGR As Variant, cCC As Variant, cPROT As Variant, cCOND As Variant, cDET As Variant
Private cAmpCU As Variant, cAmpAL As Variant
Private cargado As Boolean
Private famBK() As Long, famTB() As Long, vscBK() As Double

Public Sub RecargarCatalogo()
    cargado = False
End Sub

Private Function Tabla(hoja As String, nCols As Long) As Variant
    Dim ws As Worksheet, u As Long
    Set ws = ThisWorkbook.Worksheets(hoja)
    u = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If u < 2 Then u = 2
    Tabla = ws.Range(ws.Cells(2, 1), ws.Cells(u, nCols)).Value
End Function

Private Sub Cargar()
    If cargado Then Exit Sub
    cBK = Tabla("CAT_BREAKERS", 10)
    cTB = Tabla("CAT_TABLEROS", 10)
    cSPD = Tabla("CAT_SPD", 9)
    cRT = Tabla("CAT_REGLAS_T", 10)
    cRB = Tabla("CAT_REGLAS_B", 5)
    cSE = Tabla("CAT_SERIE", 7)
    cAMP = ThisWorkbook.Names("T_AMP").RefersToRange.Value
    cTEMP = ThisWorkbook.Names("T_TEMP").RefersToRange.Value
    cAGR = ThisWorkbook.Names("T_AGRUP").RefersToRange.Value
    cCC = ThisWorkbook.Names("T_CONSTC").RefersToRange.Value
    cPROT = ThisWorkbook.Names("T_PROT").RefersToRange.Value
    cCOND = ThisWorkbook.Names("T_CONDUIT").RefersToRange.Value
    cAmpCU = ThisWorkbook.Names("T_AMPCU").RefersToRange.Value
    cAmpAL = ThisWorkbook.Names("T_AMPAL").RefersToRange.Value
    cDET = Tabla("CAT_DET", 12)
    Dim i As Long
    ReDim famBK(1 To UBound(cBK, 1)): ReDim vscBK(1 To UBound(cBK, 1)): ReDim famTB(1 To UBound(cTB, 1))
    For i = 1 To UBound(cBK, 1)
        famBK(i) = FamB0(Sx(cBK(i, 1)), Sx(cBK(i, 3)))
        vscBK(i) = Nm(cBK(i, 9))
        If vscBK(i) = 0 Then
            vscBK(i) = 480
            If famBK(i) > 0 Then
                If Nm(cRB(famBK(i), 4)) <= 240 Then vscBK(i) = 240
            End If
        End If
    Next i
    For i = 1 To UBound(cTB, 1)
        famTB(i) = FamT0(Sx(cTB(i, 3)), Sx(cTB(i, 2)))
    Next i
    cargado = True
End Sub

Private Function Nm(v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsNumeric(v) And Not IsEmpty(v) Then
        If VarType(v) = vbString Then
            If Trim$(v) = "" Then Exit Function
        End If
        Nm = CDbl(v)
    End If
End Function

Private Function Sx(v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDouble Then
        If v = Int(v) Then Sx = CStr(CLng(v)) Else Sx = CStr(v)
    Else
        Sx = Trim$(CStr(v))
    End If
End Function

' --- patrones de las reglas: lista de patrones Like separada por coma; "(vacio)" = modelo vacio
Private Function Cumple(modelo As String, patrones As String) As Boolean
    Dim p As Variant
    If patrones = "(vacio)" Then Cumple = (Trim$(modelo) = ""): Exit Function
    For Each p In Split(patrones, ",")
        If Trim$(p) <> "" Then
            If UCase$(Trim$(modelo)) Like UCase$(Trim$(p)) Then Cumple = True: Exit Function
        End If
    Next p
End Function

Private Function EnLista(x As String, lista As String) As Long
    ' posicion (1..n) de x en la lista separada por coma o punto y coma; 0 si no esta
    Dim a As Variant, i As Long
    a = Split(Replace(lista, ";", ","), ",")
    For i = 0 To UBound(a)
        If Trim$(a(i)) <> "" And Trim$(a(i)) = x Then EnLista = i + 1: Exit Function
    Next i
End Function

Private Function ListaVacia(lista As String) As Boolean
    ListaVacia = (Trim$(Replace(Replace(lista, ",", ""), ";", "")) = "")
End Function

' regla de familia de un breaker (indice en cRB, 0 si no hay)
Private Function FamB0(marca As String, modelo As String) As Long
    Dim i As Long
    For i = 1 To UBound(cRB, 1)
        If Sx(cRB(i, 1)) = marca Then
            If Cumple(modelo, Sx(cRB(i, 3))) Then FamB0 = i: Exit Function
        End If
    Next i
End Function

Private Function FamT0(fabricante As String, modelo As String) As Long
    Dim i As Long
    For i = 1 To UBound(cRT, 1)
        If Sx(cRT(i, 1)) = fabricante Then
            If Cumple(modelo, Sx(cRT(i, 3))) Then FamT0 = i: Exit Function
        End If
    Next i
End Function

Private Function VSccrB(i As Long) As Double
    VSccrB = vscBK(i)
End Function

Private Function EnRango(rangoTxt As Variant, amp As Double) As Boolean
    Dim s As String, p As Long
    s = Sx(rangoTxt)
    If s = "" Then EnRango = True: Exit Function
    p = InStr(s, "-")
    If p > 0 Then
        EnRango = (amp >= Val(Left$(s, p - 1)) And amp <= Val(Mid$(s, p + 1)))
    ElseIf Val(s) > 0 Then
        EnRango = (amp = Val(s))
    Else
        EnRango = True
    End If
End Function

Private Function IdxBreaker(marca As String, id As Variant) As Long
    Dim i As Long
    For i = 1 To UBound(cBK, 1)
        If Sx(cBK(i, 1)) = marca And Sx(cBK(i, 2)) = Sx(id) Then IdxBreaker = i: Exit Function
    Next i
    ' sin la marca: por tipo
    For i = 1 To UBound(cBK, 1)
        If Sx(cBK(i, 2)) = Sx(id) Then IdxBreaker = i: Exit Function
    Next i
End Function

' ================== BREAKERS ==================
' Devuelve el "Tipo" (columna B de CAT_BREAKERS) del breaker elegido.
Public Function BREAKER_SEL(marca As String, polos As Variant, amp As Variant, V As Variant, iccKA As Variant, Optional familias As Variant = "", Optional unidad As Variant = "STD") As Variant
    On Error GoTo fallo
    Cargar
    BREAKER_SEL = SelBk(marca, Nm(polos), Nm(amp), Nm(V), Nm(iccKA), Sx(familias), Sx(unidad))
    Exit Function
fallo:
    BREAKER_SEL = ""
End Function

Private Function SelBk(marca As String, polos As Double, amp As Double, V As Double, icc As Double, fam As String, unidad As String) As Variant
    Dim n As Long, i As Long, j As Long, f As Long, c() As Long, k As Long, okU As Boolean
    Dim mejor As Long, mejorKA As Double, ord As Long, mejorOrd As Long, ka As Double
    If V = 0 Then V = 240
    If unidad = "" Then unidad = "STD"
    n = UBound(cBK, 1): ReDim c(1 To n)
    For i = 1 To n
        If Sx(cBK(i, 1)) = marca And Nm(cBK(i, 7)) = polos And EnRango(cBK(i, 5), amp) Then
            f = famBK(i)
            If f = 0 Or Nm(cRB(IIf(f = 0, 1, f), 4)) >= V Then
                If ListaVacia(fam) Or (f > 0 And EnLista(Sx(cRB(IIf(f = 0, 1, f), 2)), fam) > 0) Then
                    If V > 240 Then
                        If VSccrB(i) >= 480 Then k = k + 1: c(k) = i
                    Else
                        k = k + 1: c(k) = i
                    End If
                End If
            End If
        End If
    Next i
    If V <= 240 And k > 0 Then
        ' si hay la fila de 240 V del mismo modelo, se descarta la de 480 V
        Dim c2() As Long, k2 As Long, dup As Boolean
        ReDim c2(1 To k)
        For i = 1 To k
            dup = False
            If VSccrB(c(i)) >= 480 Then
                For j = 1 To k
                    If j <> i Then
                        If Sx(cBK(c(j), 3)) = Sx(cBK(c(i), 3)) And Sx(cBK(c(j), 7)) = Sx(cBK(c(i), 7)) And Sx(cBK(c(j), 5)) = Sx(cBK(c(i), 5)) And Sx(cBK(c(j), 6)) = Sx(cBK(c(i), 6)) And VSccrB(c(j)) <= 240 Then dup = True: Exit For
                    End If
                Next j
            End If
            If Not dup Then k2 = k2 + 1: c2(k2) = c(i)
        Next i
        k = k2
        For i = 1 To k: c(i) = c2(i): Next i
    End If
    ' por tipo de unidad
    Dim kU As Long
    For i = 1 To k
        If IIf(Sx(cBK(c(i), 6)) = "", "STD", Sx(cBK(c(i), 6))) = unidad Then kU = kU + 1
    Next i
    If k = 0 And Not ListaVacia(fam) Then SelBk = SelBk(marca, polos, amp, V, icc, "", unidad): Exit Function
    If k = 0 Then SelBk = "": Exit Function
    mejor = 0: mejorKA = 1E+99: mejorOrd = 999
    For i = 1 To k
        okU = (kU = 0) Or (IIf(Sx(cBK(c(i), 6)) = "", "STD", Sx(cBK(c(i), 6))) = unidad)
        If okU Then
            ka = Nm(cBK(c(i), 8))
            If icc = 0 Or ka >= icc Then
                f = famBK(c(i))
                ord = 99
                If f > 0 Then ord = EnLista(Sx(cRB(f, 2)), fam): If ord = 0 Then ord = 99
                If ka < mejorKA Or (ka = mejorKA And ord < mejorOrd) Then mejor = c(i): mejorKA = ka: mejorOrd = ord
            End If
        End If
    Next i
    If mejor = 0 Then
        mejorKA = -1
        For i = 1 To k
            okU = (kU = 0) Or (IIf(Sx(cBK(c(i), 6)) = "", "STD", Sx(cBK(c(i), 6))) = unidad)
            If okU And Nm(cBK(c(i), 8)) > mejorKA Then mejor = c(i): mejorKA = Nm(cBK(c(i), 8))
        Next i
    End If
    If mejor = 0 Then SelBk = "" Else SelBk = cBK(mejor, 2)
End Function

' Numero de catalogo con la plantilla de la familia ({p} polos, {a} amperios, {a3} 3 digitos, {m3} 3 letras del modelo, {u} unidad, {pdg} Power Defense)
Public Function BREAKER_REF(marca As String, id As Variant, polos As Variant, amp As Variant, Optional rated100 As Variant = False) As Variant
    On Error GoTo fallo
    Cargar
    Dim i As Long, f As Long, modelo As String, pl As String, un As String, a As Long, p As Long, r As String
    i = IdxBreaker(marca, id): If i = 0 Then BREAKER_REF = "": Exit Function
    modelo = Sx(cBK(i, 3)): un = Sx(cBK(i, 6)): a = CLng(Nm(amp)): p = CLng(Nm(polos))
    f = famBK(i)
    If f > 0 Then pl = Sx(cRB(f, 5))
    r = modelo
    If f > 0 And pl = "{pdg}" And a > 0 And Left$(modelo, 3) = "PDG" And Len(modelo) >= 6 Then
        Dim marco As String, letra As String, trip As String, px As Long, dig As String
        marco = Mid$(modelo, 4, 1): letra = Mid$(modelo, 6, 1)
        px = InStr(modelo, "PXR")
        If px > 0 Then
            dig = Mid$(modelo, px + 3, 2)
            If dig = "10" Then trip = "B" Else If dig = "25" Then trip = "P" Else trip = "E"
            trip = trip & IIf(Right$(un, 1) = "G", "3", "2") & "N"
        Else
            trip = IIf(Val(marco) <= 2, "TFF", "TFA")
        End If
        r = "PDG" & marco & p & letra & Format$(a, "0000") & trip & "J"
    ElseIf f > 0 And InStr(pl, "=") > 0 And a > 0 Then
        Dim it As Variant, kv As Variant
        For Each it In Split(pl, ";")
            kv = Split(it, "=")
            If UBound(kv) >= 1 Then
                If Trim$(kv(0)) = un Then r = Replace(Replace(Replace(Trim$(kv(1)), "{p}", p), "{a3}", Format$(a, "000")), "{a}", a): Exit For
            End If
        Next it
    ElseIf f > 0 And pl <> "" And a > 0 And (un = "" Or un = "STD" Or InStr(pl, "{u}") > 0) Then
        r = Replace(Replace(Replace(Replace(Replace(pl, "{p}", p), "{a3}", Format$(a, "000")), "{a}", a), "{m3}", Left$(modelo, 3)), "{u}", un)
    End If
    If CBool(rated100) And Left$(r, 3) = "PDG" Then r = "PDF" & Mid$(r, 4)
    BREAKER_REF = r
    Exit Function
fallo:
    BREAKER_REF = ""
End Function

Public Function BREAKER_DATO(marca As String, id As Variant, campo As String) As Variant
    On Error GoTo fallo
    Cargar
    Dim i As Long, f As Long
    i = IdxBreaker(marca, id): If i = 0 Then BREAKER_DATO = "": Exit Function
    Select Case LCase$(campo)
        Case "modelo": BREAKER_DATO = Sx(cBK(i, 3))
        Case "marco": BREAKER_DATO = cBK(i, 4)
        Case "amperios": BREAKER_DATO = cBK(i, 5)
        Case "unidad": BREAKER_DATO = IIf(Sx(cBK(i, 6)) = "", "STD", Sx(cBK(i, 6)))
        Case "polos": BREAKER_DATO = cBK(i, 7)
        Case "sccr": BREAKER_DATO = cBK(i, 8)
        Case "familia": f = famBK(i): If f > 0 Then BREAKER_DATO = Sx(cRB(f, 2)) Else BREAKER_DATO = ""
        Case Else: BREAKER_DATO = ""
    End Select
    If IsEmpty(BREAKER_DATO) Then BREAKER_DATO = ""
    Exit Function
fallo:
    BREAKER_DATO = ""
End Function

' Clasificacion en serie (NEC 240.86): kA de la combinacion listada principal (linea) + ramal (carga); "" si no hay
Public Function SERIE_KA(marca As String, modeloCarga As String, refLinea As String, ampLinea As Variant, V As Variant, iccKA As Variant) As Variant
    On Error GoTo fallo
    Cargar
    Dim i As Long, vv As Double, mejor As Double, okV As Boolean, sv As String
    vv = Nm(V): mejor = 0
    If modeloCarga = "" Or refLinea = "" Or Nm(iccKA) = 0 Then SERIE_KA = "": Exit Function
    For i = 1 To UBound(cSE, 1)
        sv = Sx(cSE(i, 2))
        If vv >= 440 Then
            okV = (sv = "480Y/277" Or sv = "480")
        ElseIf vv = 240 Then
            okV = (sv = "120/240")
        Else
            okV = (sv = "208Y/120" Or sv = "120/240")
        End If
        If okV And (Sx(cSE(i, 1)) = "" Or Sx(cSE(i, 1)) = marca) And Nm(cSE(i, 3)) >= Nm(iccKA) Then
            If Nm(cSE(i, 4)) = 0 Or Nm(ampLinea) <= Nm(cSE(i, 4)) Then
                If Token(refLinea, Sx(cSE(i, 5))) And Token(modeloCarga, Sx(cSE(i, 6))) Then
                    If mejor = 0 Or Nm(cSE(i, 3)) < mejor Then mejor = Nm(cSE(i, 3))
                End If
            End If
        End If
    Next i
    If mejor = 0 Then SERIE_KA = "" Else SERIE_KA = mejor
    Exit Function
fallo:
    SERIE_KA = ""
End Function

Private Function Token(modelo As String, lista As String) As Boolean
    ' "PDG2xG" = PDG2 + digito + G (x = digito); el resto son prefijos del modelo
    Dim t As Variant, p As String, i As Long, ch As String, q As String
    For Each t In Split(Replace(lista, ";", ","), ",")
        p = Trim$(t): q = ""
        For i = 1 To Len(p)
            ch = Mid$(p, i, 1)
            If StrComp(ch, "x", vbBinaryCompare) = 0 Then
                q = q & "#"
            ElseIf ch Like "[A-Za-z0-9-]" Then
                q = q & ch
            End If
        Next i
        If q <> "" Then
            If UCase$(modelo) Like UCase$(q) & "*" Then Token = True: Exit Function
        End If
    Next t
End Function

' ================== TABLEROS DE CATALOGO ==================
Public Function TABLERO_SEL(marca As String, V As Variant, tipo As String, prot As Variant, espUsados As Variant, reservaEsp As Variant, Optional clase As Variant = "Tablero", Optional principal As Variant = "Interruptor", Optional espacios As Variant = 0) As Variant
    On Error GoTo fallo
    Cargar
    Dim i As Long, f As Long, nec As Double, sub_ As Boolean, cl As String, pr As String, best As Long, kb As Double, kn As Double
    Dim ord As Double, exacto As Long, bOrd As Double, bEx As Long, ok As Boolean, fs As String
    nec = Application.WorksheetFunction.Max(Nm(espacios), Application.WorksheetFunction.RoundUp(Nm(espUsados) * (1 + Nm(reservaEsp) / 100), 0))
    sub_ = (Sx(clase) Like "Sub*")
    pr = IIf(Sx(principal) = "Zapatas", "Zapatas", "Interruptor")
    best = 0
    For i = 1 To UBound(cTB, 1)
        If Sx(cTB(i, 3)) = marca And Nm(cTB(i, 4)) >= Nm(prot) And Nm(cTB(i, 7)) >= nec Then
            f = famTB(i)
            ok = True
            If f > 0 Then
                If Nm(cRT(f, 4)) < Nm(V) Then ok = False
                cl = Sx(cRT(f, 9)): If cl = "" Then cl = "tablero"
            Else
                cl = "tablero"
            End If
            If sub_ And cl = "tablero" Then ok = False
            If Not sub_ And cl = "subestacion" Then ok = False
            If Sx(cTB(i, 9)) <> "" And Sx(cTB(i, 9)) <> "Ambos" And Sx(cTB(i, 9)) <> pr Then ok = False
            fs = Sx(cTB(i, 8))
            If fs = "" And f > 0 Then fs = Sx(cRT(f, 5))
            If fs <> "" Then
                If EnLista(tipo, Replace(fs, "/", ",")) = 0 Then ok = False
            End If
            If ok Then
                ord = 99: If f > 0 Then ord = Nm(cRT(f, 8))
                exacto = 1: If Nm(espacios) > 0 And Nm(cTB(i, 7)) = Nm(espacios) Then exacto = 0
                If best = 0 Then
                    best = i: bOrd = ord: bEx = exacto
                ElseIf exacto < bEx Or (exacto = bEx And (ord < bOrd Or (ord = bOrd And (Nm(cTB(i, 4)) < Nm(cTB(best, 4)) Or (Nm(cTB(i, 4)) = Nm(cTB(best, 4)) And Nm(cTB(i, 7)) < Nm(cTB(best, 7))))))) Then
                    best = i: bOrd = ord: bEx = exacto
                End If
            End If
        End If
    Next i
    If best = 0 Then TABLERO_SEL = "" Else TABLERO_SEL = cTB(best, 1)
    Exit Function
fallo:
    TABLERO_SEL = ""
End Function

Public Function CAT_DATO(id As Variant, campo As String) As Variant
    On Error GoTo fallo
    Cargar
    Dim i As Long, f As Long
    CAT_DATO = ""
    If Sx(id) = "" Then Exit Function
    For i = 1 To UBound(cTB, 1)
        If Sx(cTB(i, 1)) = Sx(id) Then
            f = famTB(i)
            Select Case LCase$(campo)
                Case "modelo": CAT_DATO = Sx(cTB(i, 2))
                Case "fabricante": CAT_DATO = Sx(cTB(i, 3))
                Case "barraf": CAT_DATO = cTB(i, 4)
                Case "barran": CAT_DATO = cTB(i, 5)
                Case "barrat": CAT_DATO = cTB(i, 6)
                Case "espacios": CAT_DATO = cTB(i, 7)
                Case "familia": If f > 0 Then CAT_DATO = Sx(cRT(f, 2))
                Case "ramales": If f > 0 Then CAT_DATO = Sx(cRT(f, 6))
                Case "principales": If f > 0 Then CAT_DATO = Sx(cRT(f, 7))
                Case "clase": If f > 0 Then CAT_DATO = Sx(cRT(f, 9))
            End Select
            If IsEmpty(CAT_DATO) Then CAT_DATO = ""
            Exit Function
        End If
    Next i
    Exit Function
fallo:
    CAT_DATO = ""
End Function

' ================== SUPRESORES ==================
Private Function NormV(s As String) As String
    Dim a As Variant, v() As Double, k As Long, i As Long, j As Long, t As Double
    a = Split(s, "/")
    ReDim v(0 To UBound(a) + 1)
    For i = 0 To UBound(a)
        If IsNumeric(Trim$(a(i))) Then
            If CDbl(Trim$(a(i))) > 0 Then v(k) = CDbl(Trim$(a(i))): k = k + 1
        End If
    Next i
    For i = 0 To k - 2
        For j = i + 1 To k - 1
            If v(j) < v(i) Then t = v(i): v(i) = v(j): v(j) = t
        Next j
    Next i
    For i = 0 To k - 1
        NormV = NormV & IIf(i > 0, "/", "") & v(i)
    Next i
End Function

Public Function SPD_SEL(marca As String, sistema As String, fases As Variant, kaMin As Variant) As Variant
    On Error GoTo fallo
    Cargar
    Dim i As Long, fs As Long, best As Long, mayor As Long, sv As String
    fs = IIf(Nm(fases) = 3, 3, 1): sv = NormV(sistema)
    For i = 1 To UBound(cSPD, 1)
        If Sx(cSPD(i, 1)) = marca And NormV(Sx(cSPD(i, 7))) = sv And Nm(cSPD(i, 8)) = fs Then
            If Nm(cSPD(i, 5)) >= Nm(kaMin) Then
                If best = 0 Then best = i Else If Nm(cSPD(i, 5)) < Nm(cSPD(best, 5)) Then best = i
            End If
            If mayor = 0 Then mayor = i Else If Nm(cSPD(i, 5)) >= Nm(cSPD(mayor, 5)) Then mayor = i
        End If
    Next i
    If best = 0 Then best = mayor
    If best = 0 Then SPD_SEL = "" Else SPD_SEL = cSPD(best, 2)
    Exit Function
fallo:
    SPD_SEL = ""
End Function

Public Function SPD_DATO(marca As String, id As Variant, campo As String, Optional sistema As String = "") As Variant
    ' los tipos de SPD se repiten por voltaje: se busca el del sistema del tablero y, si no hay, el primero
    On Error GoTo fallo
    Cargar
    Dim i As Long, k As Long, pasada As Long
    SPD_DATO = ""
    For pasada = 1 To 2
    For i = 1 To UBound(cSPD, 1)
        If Sx(cSPD(i, 2)) = Sx(id) And (Sx(cSPD(i, 1)) = marca Or marca = "") And (pasada = 2 Or sistema = "" Or NormV(Sx(cSPD(i, 7))) = NormV(sistema)) Then
            Select Case LCase$(campo)
                Case "modelo": SPD_DATO = Sx(cSPD(i, 3))
                Case "montaje": SPD_DATO = Sx(cSPD(i, 4))
                Case "kall": SPD_DATO = cSPD(i, 5)
                Case "kaln": SPD_DATO = cSPD(i, 6)
            End Select
            Exit Function
        End If
    Next i
    Next pasada
    Exit Function
fallo:
    SPD_DATO = ""
End Function

' ================== CONDUCTORES ==================
Public Function PARTE(texto As Variant, i As Long) As Variant
    Dim a As Variant
    a = Split(Sx(texto), "|")
    If i - 1 <= UBound(a) And i >= 1 Then PARTE = a(i - 1) Else PARTE = ""
End Function

Public Function PARTEN(texto As Variant, i As Long) As Variant
    ' parte numerica (independiente de la configuracion regional)
    Dim p As Variant
    p = PARTE(texto, i)
    If IsNumeric(p) And p <> "" Then PARTEN = CDbl(p) Else PARTEN = 0
End Function

Private Function FilaCal(t As Variant, cal As String) As Long
    Dim i As Long
    For i = 1 To UBound(t, 1)
        If Sx(t(i, 1)) = cal Then FilaCal = i: Exit Function
    Next i
End Function

Private Function CalibrePorAmp(amp As Double, mat As String) As String
    Dim t As Variant, i As Long
    If mat = "AL" Or mat = "AL-MC" Then t = cAmpAL Else t = cAmpCU
    For i = 1 To UBound(t, 1)
        If Nm(t(i, 1)) <= amp Then CalibrePorAmp = Sx(t(i, 2)) Else Exit For
    Next i
End Function

Private Function Ampac(cal As String, mat As String, nPar As Double, ais As String, tempAmb As String, tempBorne As Double, agrup As String, Ireq As Double, prot As Double, _
                       ByRef ok As Boolean, ByRef cap As Double, ByRef amp90 As Double, ByRef ft As Double, ByRef fg As Double, ByRef term As Double) As Boolean
    Dim f As Long, tA As Long, mi As Long, ti As Long, i As Long, corr As Double, sig As Double, okI As Boolean, okP As Boolean
    f = FilaCal(cAMP, cal): If f = 0 Then Exit Function
    Select Case UCase$(Trim$(ais))
        Case "RHW", "THW", "TW", "USE": tA = 75
        Case Else: tA = 90
    End Select
    Select Case mat
        Case "CU-MC": mi = 1
        Case "AL": mi = 2
        Case "AL-MC": mi = 3
        Case Else: mi = 0
    End Select
    ti = IIf(tA = 60, 0, IIf(tA = 75, 1, 2))
    amp90 = Nm(cAMP(f, 2 + mi * 3 + ti))
    ti = IIf(tempBorne = 60, 0, IIf(tempBorne = 90, 2, 1))
    term = Nm(cAMP(f, 2 + mi * 3 + ti)): If term = 0 Then term = amp90
    ft = 0
    For i = 1 To UBound(cTEMP, 1)
        If Sx(cTEMP(i, 1)) = tempAmb Then ft = Nm(cTEMP(i, IIf(tA = 60, 2, IIf(tA = 75, 3, 4))))
    Next i
    fg = 1
    For i = 1 To UBound(cAGR, 1)
        If Sx(cAGR(i, 1)) = agrup Then fg = Nm(cAGR(i, 2))
    Next i
    If fg = 0 Then fg = 1
    corr = amp90 * ft * fg
    cap = IIf(corr < term, corr, term) * nPar
    sig = 0
    For i = 1 To UBound(cPROT, 1)
        If Nm(cPROT(i, 2)) >= cap - 0.000000001 Then
            If sig = 0 Or Nm(cPROT(i, 2)) < sig Then sig = Nm(cPROT(i, 2))
        End If
    Next i
    okI = cap >= Ireq - 0.000000001
    okP = (prot = 0) Or (prot <= cap + 0.000000001) Or (prot <= 800 And prot = sig)
    ok = okI And okP And ft > 0
    Ampac = True
End Function

' Valida el calibre del alimentador (310.15: temperatura y agrupamiento) y lo aumenta si no cumple.
' Devuelve "paralelos|calibre|cumple(1/0)|capacidad|amp90|Ft|Fg|bornes"
Public Function CAL_AJUSTE(cal As Variant, mat As String, nPar As Variant, ais As String, tempAmb As String, tempBorne As Variant, agrup As String, Ireq As Variant, prot As Variant, Optional auto As Variant = True, Optional aumento As Variant = 1) As Variant
    On Error GoTo fallo
    Cargar
    Dim c As String, n As Double, ok As Boolean, cap As Double, a9 As Double, ft As Double, fg As Double, tm As Double, it As Long
    Dim CALS As Variant, i As Long, j As Long, sig As String
    CALS = Array("14", "12", "10", "8", "6", "4", "3", "2", "1", "1/0", "2/0", "3/0", "4/0", "250", "300", "350", "400", "500", "600", "700", "750", "800", "900", "1000")
    c = Sx(cal): n = Nm(nPar): If n < 1 Then n = 1
    If mat = "" Then mat = "CU"
    If Not Ampac(c, mat, n, ais, tempAmb, Nm(tempBorne), agrup, Nm(Ireq), Nm(prot), ok, cap, a9, ft, fg, tm) Then CAL_AJUSTE = n & "|" & c & "|1|0|0|0|0|0": Exit Function
    Do While CBool(auto) And Not ok And ft > 0 And it < 14
        i = -1
        For j = 0 To UBound(CALS)
            If CALS(j) = c Then i = j
        Next j
        sig = ""
        For j = i + 1 To UBound(CALS)
            If FilaCal(cCOND, CStr(CALS(j))) > 0 And FilaCal(cAMP, CStr(CALS(j))) > 0 Then sig = CALS(j): Exit For
        Next j
        If sig <> "" And j <= 17 Then
            c = sig
        Else
            n = n + 1: c = CalibrePorAmp(Nm(prot) * Nm(aumento) / n, mat)
        End If
        Ampac c, mat, n, ais, tempAmb, Nm(tempBorne), agrup, Nm(Ireq), Nm(prot), ok, cap, a9, ft, fg, tm
        it = it + 1
    Loop
    CAL_AJUSTE = n & "|" & c & "|" & IIf(ok, 1, 0) & "|" & Round(cap, 3) & "|" & a9 & "|" & ft & "|" & fg & "|" & tm
    Exit Function
fallo:
    CAL_AJUSTE = ""
End Function

' Corriente de cortocircuito al final de un alimentador (metodo punto a punto, constantes C).
Public Function P2P(I0 As Variant, cal As Variant, ais As String, tub As String, nPar As Variant, L As Variant, V As Variant, fases As Variant, Optional lmax As Variant = 0) As Variant
    On Error GoTo fallo
    Cargar
    Dim f As Long, aC As String, col As Long, C As Double, Lm As Double, k As Double, x0 As Double
    x0 = Nm(I0)
    If x0 = 0 Then P2P = 0: Exit Function
    f = FilaCal(cCC, Sx(cal))
    If f = 0 Then P2P = x0: Exit Function
    If InStr(1, ais, "XHHW") > 0 Then
        aC = "XHHW-2"
    ElseIf InStr(1, ais, "RHW") > 0 Then
        aC = "RHW"
    ElseIf InStr(1, ais, "barra") > 0 Then
        aC = "DUCTOBARRA"
    Else
        aC = "THHN"
    End If
    Select Case aC
        Case "THHN": col = 2
        Case "RHW": col = 4
        Case "XHHW-2": col = 6
        Case Else: col = 8
    End Select
    If aC <> "DUCTOBARRA" And tub = "EMT" Then col = col + 1
    C = Nm(cCC(f, col))
    If C = 0 Then P2P = x0: Exit Function
    Lm = Nm(L): If Nm(lmax) > 0 And Nm(lmax) < Lm Then Lm = Nm(lmax)
    k = IIf(Nm(fases) = 3, 1.732, 2)
    P2P = x0 / (1 + (k * (Lm / 0.3048) * x0) / (C * IIf(Nm(nPar) < 1, 1, Nm(nPar)) * Nm(V)))
    Exit Function
fallo:
    P2P = Nm(I0)
End Function

' ================== NEC 2020: GFCI / AFCI ==================
Private Function SinAcentos(s As String) As String
    Dim a As Variant, b As Variant, i As Long
    s = LCase$(s)
    a = Array("á", "é", "í", "ó", "ú", "ü", "ñ")
    b = Array("a", "e", "i", "o", "u", "u", "n")
    For i = 0 To UBound(a): s = Replace(s, a(i), b(i)): Next i
    SinAcentos = s
End Function

Private Function Hay(d As String, palabras As String) As String
    Dim p As Variant
    For Each p In Split(palabras, ",")
        If InStr(1, d, p, vbBinaryCompare) > 0 Then Hay = p: Exit Function
    Next p
End Function

Public Function NEC_UNIDAD(desc As String, det As Variant, Vdet As Variant, Fdet As Variant, amp As Variant, Vsis As Variant, Optional ocup As String = "comercial") As Variant
    On Error GoTo fallo
    Cargar
    Dim d As String, i As Long, tipo As Double, toma As Boolean, g As String, a As String, w As String
    Const LG As String = "bano,cocina,exterior,lavander,garaj,garage,sotano,piscina,jacuzzi,tina,ducha,bbq,terraza,patio,azotea,fregadero,mojad,humed,ice maker,maquina de hielo,vestidor"
    Const LA As String = "habitaci,dormitor,recamara,sala,comedor,estudio,pasillo,closet,cuarto,biblioteca,family,sunroom,lavander,cocina"
    NEC_UNIDAD = "STD"
    If Nm(Vsis) >= 440 Then Exit Function
    d = desc
    For i = 1 To UBound(cDET, 1)
        If Sx(cDET(i, 1)) = Sx(det) Then d = d & " " & Sx(cDET(i, 3)): tipo = Nm(cDET(i, 2)): Exit For
    Next i
    d = SinAcentos(d)
    toma = (Hay(d, "toma,receptac,enchuf") <> "") Or tipo = 2 Or tipo = 3
    If Hay(d, "lavaplato,lavavajilla") <> "" Then
        g = "1"
    ElseIf (Hay(d, "piscina,jacuzzi") <> "" Or d Like "*spa" Or d Like "*spa[!a-z]*") And Nm(amp) <= 60 Then
        g = "1"
    ElseIf ocup = "vivienda" Then
        If toma And Hay(d, LG) <> "" Then
            g = "1"
        ElseIf Hay(d, "exterior,condensad,unidad cu,uc-,aire,a/c") <> "" And Nm(amp) <= 50 Then
            g = "1"
        End If
    ElseIf toma And Hay(d, LG) <> "" And Nm(amp) <= IIf(Nm(Fdet) = 3, 100, 50) Then
        g = "1"
    End If
    If Nm(Vdet) = 120 And Nm(amp) <= 20 And (toma Or Hay(d, "ilum,luz") <> "") Then
        If (ocup = "vivienda" And Hay(d, LA) <> "") Or ((ocup = "hotel" Or ocup = "hospital") And Hay(d, "habitaci,dormitor") <> "") Then a = "1"
    End If
    If g <> "" And a <> "" Then
        NEC_UNIDAD = "AFCI/GFCI"
    ElseIf g <> "" Then
        NEC_UNIDAD = "GFCI"
    ElseIf a <> "" Then
        NEC_UNIDAD = "AFCI"
    End If
    Exit Function
fallo:
    NEC_UNIDAD = "STD"
End Function
