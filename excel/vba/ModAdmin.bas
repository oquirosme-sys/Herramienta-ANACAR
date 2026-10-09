Attribute VB_Name = "ModAdmin"
' Modo administrador: muestra las hojas de catalogo (CAT_*) para editar marcas, tableros, breakers,
' supresores, detalles y tipos de carga, conductores, reglas de fabricante, clasificacion en serie y tipos de uso.
Option Explicit

Private Function Clave() As String
    On Error Resume Next
    Clave = Mid$(ThisWorkbook.Names("ADMIN_CLAVE").RefersTo, 3, Len(ThisWorkbook.Names("ADMIN_CLAVE").RefersTo) - 3)
    If Clave = "" Then Clave = "sinergia-admin"
End Function

Public Sub ModoAdministrador()
    Dim c As String, ws As Worksheet
    c = InputBox("Contrasena de administrador:", "Modo administrador")
    If c = "" Then Exit Sub
    If c <> Clave() Then MsgBox "Contrasena incorrecta.", vbExclamation: Exit Sub
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 4) = "CAT_" Or ws.Name = "MC_MACHOTE" Or ws.Name = "VISTA_MACHOTE" Then ws.Visible = xlSheetVisible
    Next ws
    ThisWorkbook.Worksheets("CAT_BREAKERS").Activate
    MsgBox "Hojas de catalogo visibles (CAT_*). MC_MACHOTE y VISTA_MACHOTE son las plantillas de los tableros nuevos." & vbCrLf & _
           "Reglas: en 'Patrones del modelo' use comodines de Excel/VBA (* ? #) separados por coma; (vacio) = modelo vacio." & vbCrLf & _
           "Al terminar use 'Salir de administrador' para ocultarlas y recalcular.", vbInformation
End Sub

Public Sub SalirAdministrador()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 4) = "CAT_" Or ws.Name = "MC_MACHOTE" Or ws.Name = "VISTA_MACHOTE" Then ws.Visible = xlSheetVeryHidden
    Next ws
    ModUDF.RecargarCatalogo
    Application.CalculateFull
    ThisWorkbook.Worksheets("INICIO").Activate
End Sub

Public Sub CambiarClave()
    Dim a As String, n As String
    a = InputBox("Contrasena actual:", "Cambiar contrasena")
    If a = "" Then Exit Sub
    If a <> Clave() Then MsgBox "Contrasena incorrecta.", vbExclamation: Exit Sub
    n = InputBox("Nueva contrasena:", "Cambiar contrasena")
    If n = "" Then Exit Sub
    ThisWorkbook.Names.Add Name:="ADMIN_CLAVE", RefersTo:="=""" & Replace(n, """", "") & """", Visible:=False
    MsgBox "Contrasena cambiada.", vbInformation
End Sub

Public Sub RecalcularTodo()
    ModUDF.RecargarCatalogo
    Application.CalculateFull
End Sub
