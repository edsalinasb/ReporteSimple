Attribute VB_Name = "DllRSimple_Excel"
Option Explicit
' =============================================================================
' Exportacion a Excel del Libro Diario Simplificado (Formato 5.2)
' - Consume el modelo normalizado que llena PreparaDetalledelDiarioSimplificado
'   (modulo DllRSimple_Diario52); no vuelve a consultar tablas temporales.
' - Late binding (CreateObject): no requiere referencia a Excel en el proyecto.
' - No guarda archivos, no genera CSV y no imprime: deja Excel abierto y visible.
' - Si falla, cierra sin guardar el libro y la instancia de Excel que creo.
' =============================================================================

Private Const xlLandscape As Long = 2
Private Const xlCenter As Long = -4108
Private Const xlLeft As Long = -4131
Private Const xlRight As Long = -4152
Private Const xlEdgeTop As Long = 8
Private Const xlEdgeBottom As Long = 9
Private Const xlContinuous As Long = 1
Private Const xlThin As Long = 2
Private Const xlMedium As Long = -4138
Private Const xlPaperA4 As Long = 9

Private Const kCOL_FECHA As Integer = 1
Private Const kCOL_CAR As Integer = 2
Private Const kCOL_GLOSA As Integer = 3
Private Const kCOLS_FIJAS As Integer = 3
Private Const kFILA_PRIMER_BLOQUE As Long = 4
Private Const kFMT_MONTO As String = "#,##0.00;-#,##0.00;"""""
Private Const kFMT_FECHA As String = "dd/mm/yyyy"
Private Const kFMT_TEXTO As String = "@"
Private Const kCOLOR_EMPRESA As Long = &H808000      ' RGB(0, 128, 128) verde azulado
Private Const kCOLOR_CABECERA As Long = &HF2F2F2      ' gris claro

Public Const kERR_EXCEL_CAPACIDAD As Long = vbObjectError + 5301

' Genera la hoja. Devuelve True si Excel quedo abierto con el reporte, False si
' no hay filas. Cualquier error se propaga al llamador despues de cerrar Excel.
Public Function ExcelDiarioSimplificadoF52(ByVal sEmpresa As String, ByVal sRUC As String, _
                                           ByVal sTitulo As String, ByVal sPeriodo As String) As Boolean
    Dim xlApp As Object, xlWb As Object, xlWs As Object
    Dim lErr As Long, sSrc As String, sDesc As String
    Dim lCab As Long, k As Long, lFila As Long, lFilasNecesarias As Long
    Dim iUltCol As Integer

    On Error GoTo Fallo
    If Not D52_Listo() Then
        Err.Raise D52_ERR_ESTADO, "ExcelDiarioSimplificadoF52", _
                  "El Diario Simplificado no se preparó correctamente; no se exporta."
    End If
    If Not D52_FechasTipadas() Then
        Err.Raise D52_ERR_ESTADO, "ExcelDiarioSimplificadoF52", _
                  "El campo Fecha_Documento no es de tipo Fecha/Hora; no se puede garantizar el orden cronológico."
    End If

    D52_ConstruirFilas
    If D52_NumFilas() = 0 Then Exit Function

    iUltCol = kCOLS_FIJAS + D52_COLUMNAS
    ' Por bloque: titulo + cabecera de columnas + filas + total + linea en blanco
    lFilasNecesarias = kFILA_PRIMER_BLOQUE + D52_NumFilas() + 4 * D52_NumBloques()

    Set xlApp = CreateObject("Excel.Application")
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False
    Set xlWb = xlApp.Workbooks.Add
    Do While xlWb.Worksheets.Count > 1
        xlWb.Worksheets(xlWb.Worksheets.Count).Delete
    Loop
    Set xlWs = xlWb.Worksheets(1)
    If lFilasNecesarias > xlWs.Rows.Count Then
        Err.Raise kERR_EXCEL_CAPACIDAD, "ExcelDiarioSimplificadoF52", _
                  "El reporte requiere " & lFilasNecesarias & " filas y esta versión de Excel admite " & xlWs.Rows.Count & "."
    End If
    xlWs.Name = "Formato 5.2"
    xlWs.Cells.Font.Name = "Arial"
    xlWs.Cells.Font.Size = 8
    PrepararColumnas xlWs, iUltCol

    EscribirEncabezado xlWs, iUltCol, sEmpresa, sRUC, sTitulo, sPeriodo

    lFila = kFILA_PRIMER_BLOQUE
    k = 1
    For lCab = 1 To D52_NumBloques()
        lFila = EscribirBloque(xlWs, lCab, k, lFila, iUltCol) + 1
    Next

    ConfigurarPagina xlWs, iUltCol, lFila - 1

    xlWs.Activate
    xlWs.Range("A1").Select
    xlApp.ScreenUpdating = True
    xlApp.DisplayAlerts = True
    xlApp.Visible = True
    xlApp.UserControl = True
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    ExcelDiarioSimplificadoF52 = True
    Exit Function

Fallo:
    lErr = Err.Number: sSrc = Err.Source: sDesc = Err.Description
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = False
        xlApp.Quit
    End If
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    On Error GoTo 0
    Err.Raise lErr, sSrc, sDesc
End Function

Private Sub PrepararColumnas(xlWs As Object, ByVal iUltCol As Integer)
    Dim i As Integer
    xlWs.Columns(kCOL_FECHA).ColumnWidth = 10
    xlWs.Columns(kCOL_CAR).ColumnWidth = 24
    xlWs.Columns(kCOL_GLOSA).ColumnWidth = 30
    For i = kCOLS_FIJAS + 1 To iUltCol
        xlWs.Columns(i).ColumnWidth = 12
    Next
    ' CAR y glosa como texto: conserva ceros a la izquierda y evita formulas
    xlWs.Columns(kCOL_CAR).NumberFormat = kFMT_TEXTO
    xlWs.Columns(kCOL_GLOSA).NumberFormat = kFMT_TEXTO
End Sub

Private Sub EscribirEncabezado(xlWs As Object, ByVal iUltCol As Integer, ByVal sEmpresa As String, _
                              ByVal sRUC As String, ByVal sTitulo As String, ByVal sPeriodo As String)
    With xlWs
        .Range(.Cells(1, 1), .Cells(2, iUltCol)).NumberFormat = kFMT_TEXTO
        .Cells(1, 1).Value = sEmpresa
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 11
        .Cells(1, 1).Font.Color = kCOLOR_EMPRESA
        .Cells(1, iUltCol).Value = "RUC: " & sRUC
        .Cells(1, iUltCol).HorizontalAlignment = xlRight
        .Cells(1, iUltCol).Font.Bold = True
        With .Range(.Cells(1, 1), .Cells(1, iUltCol)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        .Cells(2, 1).Value = sTitulo
        .Cells(2, 1).Font.Bold = True
        .Cells(2, 1).Font.Size = 10
        .Cells(2, iUltCol).Value = sPeriodo
        .Cells(2, iUltCol).HorizontalAlignment = xlRight
        .Cells(2, iUltCol).Font.Bold = True
        With .Range(.Cells(2, 1), .Cells(2, iUltCol)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End With
End Sub

' Escribe un bloque (seccion o su continuacion) y devuelve la ultima fila usada.
' k es la posicion en las filas ordenadas; avanza mientras pertenezcan al bloque.
Private Function EscribirBloque(xlWs As Object, ByVal lCab As Long, ByRef k As Long, _
                                ByVal lFilaIni As Long, ByVal iUltCol As Integer) As Long
    Dim aDatos() As Variant
    Dim kIni As Long, nFil As Long, nTot As Long, r As Long, f As Long
    Dim i As Integer, nCtas As Integer, iSec As Integer
    Dim cMonto As Currency, sTitulo As String, vFecha As Variant

    kIni = k
    Do While k <= D52_NumFilas()
        If D52_FilaCabecera(D52_FilaOrdenada(k)) <> lCab Then Exit Do
        k = k + 1
    Loop
    nFil = k - kIni
    nTot = nFil + 3
    nCtas = D52_BloqueNumCuentas(lCab)
    iSec = D52_BloqueSeccion(lCab)

    ReDim aDatos(1 To nTot, 1 To iUltCol)
    sTitulo = D52_NombreSeccion(iSec)
    If D52_EsContinuacion(lCab) Then sTitulo = sTitulo & " (continuación)"
    aDatos(1, 1) = sTitulo

    aDatos(2, kCOL_FECHA) = "FECHA"
    aDatos(2, kCOL_CAR) = "CAR"
    aDatos(2, kCOL_GLOSA) = "GLOSA"
    For i = 0 To nCtas - 1
        aDatos(2, kCOLS_FIJAS + 1 + i) = D52_BloqueCuenta(lCab, i)
    Next

    For r = 1 To nFil
        f = D52_FilaOrdenada(kIni + r - 1)
        vFecha = D52_FilaFecha(f)
        If VarType(vFecha) = vbDate Then aDatos(2 + r, kCOL_FECHA) = vFecha
        aDatos(2 + r, kCOL_CAR) = D52_FilaCAR(f)
        aDatos(2 + r, kCOL_GLOSA) = D52_FilaGlosa(f)
        For i = 0 To nCtas - 1
            cMonto = D52_FilaMonto(f, i)
            If cMonto <> 0 Then aDatos(2 + r, kCOLS_FIJAS + 1 + i) = CDbl(cMonto)
        Next
    Next

    aDatos(nTot, kCOL_GLOSA) = "TOTAL " & D52_NombreSeccion(iSec)
    For i = 0 To nCtas - 1
        cMonto = D52_TotalColumna(lCab, i)
        If cMonto <> 0 Then aDatos(nTot, kCOLS_FIJAS + 1 + i) = CDbl(cMonto)
    Next

    With xlWs
        ' Texto antes de asignar valores (conserva ceros a la izquierda de las cuentas)
        .Range(.Cells(lFilaIni, 1), .Cells(lFilaIni + 1, iUltCol)).NumberFormat = kFMT_TEXTO
        .Range(.Cells(lFilaIni, 1), .Cells(lFilaIni + nTot - 1, iUltCol)).Value = aDatos
        ' Fecha e importes despues, para que Excel no aplique formatos propios
        If nFil > 0 Then
            .Range(.Cells(lFilaIni + 2, kCOL_FECHA), .Cells(lFilaIni + 1 + nFil, kCOL_FECHA)).NumberFormat = kFMT_FECHA
        End If
        .Range(.Cells(lFilaIni + 2, kCOLS_FIJAS + 1), .Cells(lFilaIni + nTot - 1, iUltCol)).NumberFormat = kFMT_MONTO

        With .Cells(lFilaIni, 1).Font
            .Bold = True
            .Size = 9
        End With
        With .Range(.Cells(lFilaIni + 1, 1), .Cells(lFilaIni + 1, iUltCol))
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = kCOLOR_CABECERA
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
        If nFil > 0 Then
            .Range(.Cells(lFilaIni + 2, kCOL_FECHA), .Cells(lFilaIni + 1 + nFil, kCOL_FECHA)).HorizontalAlignment = xlCenter
            .Range(.Cells(lFilaIni + 2, kCOL_CAR), .Cells(lFilaIni + 1 + nFil, kCOL_GLOSA)).HorizontalAlignment = xlLeft
        End If
        With .Range(.Cells(lFilaIni + nTot - 1, 1), .Cells(lFilaIni + nTot - 1, iUltCol))
            .Font.Bold = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
        End With
    End With
    EscribirBloque = lFilaIni + nTot - 1
End Function

' La configuracion de pagina depende del controlador de impresora; si no hay
' impresora instalada Excel produce errores aqui que no afectan los datos.
Private Sub ConfigurarPagina(xlWs As Object, ByVal iUltCol As Integer, ByVal lUltimaFila As Long)
    On Error Resume Next
    With xlWs.PageSetup
        .PrintArea = xlWs.Range(xlWs.Cells(1, 1), xlWs.Cells(lUltimaFila, iUltCol)).Address
        .PrintTitleRows = "$1:$2"
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = xlWs.Application.InchesToPoints(0.3)
        .RightMargin = xlWs.Application.InchesToPoints(0.3)
        .TopMargin = xlWs.Application.InchesToPoints(0.4)
        .BottomMargin = xlWs.Application.InchesToPoints(0.5)
        .CenterFooter = "Página &P de &N"
    End With
    Err.Clear
End Sub
