Attribute VB_Name = "DllRSimple_Excel"
Option Explicit

' ===== Constantes de Excel (late binding: no requiere referencia a Excel) =====
Private Const xlPortrait As Long = 1
Private Const xlLandscape As Long = 2
Private Const xlCenter As Long = -4108
Private Const xlLeft As Long = -4131
Private Const xlRight As Long = -4152
Private Const xlEdgeTop As Long = 8
Private Const xlEdgeBottom As Long = 9
Private Const xlInsideHorizontal As Long = 12
Private Const xlInsideVertical As Long = 11
Private Const xlEdgeLeft As Long = 7
Private Const xlEdgeRight As Long = 10
Private Const xlContinuous As Long = 1
Private Const xlDouble As Long = -4119
Private Const xlThin As Long = 2
Private Const xlMedium As Long = -4138
Private Const xlPaperA4 As Long = 9

Private Const kFMT_MONTO As String = "#,##0.00;(#,##0.00);"""""
Private Const kFMT_FECHA As String = "dd/mm/yyyy"

' ===== Tabla de agrupación de cuentas =====
Public Type TipoSeccion
    Nombre As String
    Inicio As Integer
    Final As Integer
End Type

Private aSecciones() As TipoSeccion
Private nSecciones As Integer

' ===== Estructuras internas =====
Private Type TipoColumna
    Cuenta As String
    Indice As Integer
End Type

Private Type TipoFila
    Fecha As String
    Correlativo As String
    Glosa As String
    Cuenta As String
    Montos() As Currency
    EsSubtotalCuenta As Boolean
    EsSubtotalSeccion As Boolean
    EsTotalGeneral As Boolean
    TextoEncabezado As String
End Type

' ----- Inicializa las secciones (llamar una sola vez) -----
Public Sub InitSecciones()
    nSecciones = 10
    ReDim aSecciones(1 To nSecciones)
    
    aSecciones(1).Nombre = "Activo Corriente": aSecciones(1).Inicio = 10: aSecciones(1).Final = 29
    aSecciones(2).Nombre = "Activo No Corriente": aSecciones(2).Inicio = 30: aSecciones(2).Final = 39
    aSecciones(3).Nombre = "Pasivo Corriente": aSecciones(3).Inicio = 40: aSecciones(3).Final = 46
    aSecciones(4).Nombre = "Pasivo No Corriente": aSecciones(4).Inicio = 47: aSecciones(4).Final = 49
    aSecciones(5).Nombre = "Patrimonio": aSecciones(5).Inicio = 50: aSecciones(5).Final = 59
    aSecciones(6).Nombre = "Gestión Egresos": aSecciones(6).Inicio = 60: aSecciones(6).Final = 69
    aSecciones(7).Nombre = "Gestión Ingresos": aSecciones(7).Inicio = 70: aSecciones(7).Final = 79
    aSecciones(8).Nombre = "Saldos Intermedios": aSecciones(8).Inicio = 80: aSecciones(8).Final = 89
    aSecciones(9).Nombre = "Analítica de Explotación": aSecciones(9).Inicio = 90: aSecciones(9).Final = 99
    aSecciones(10).Nombre = "Cuentas de Orden": aSecciones(10).Inicio = 0: aSecciones(10).Final = 0
End Sub

' ----- Obtiene la sección de una cuenta -----
Private Function ObtenerSeccion(sCuenta As String) As Integer
    Dim iCta As Integer, i As Integer
    On Error Resume Next
    iCta = CInt(Left$(sCuenta, 2))
    For i = 1 To nSecciones
        If aSecciones(i).Inicio = 0 And aSecciones(i).Final = 0 Then ' Cuentas de Orden
            If sCuenta = "00" Or iCta = 0 Then
                ObtenerSeccion = i
                Exit Function
            End If
        ElseIf iCta >= aSecciones(i).Inicio And iCta <= aSecciones(i).Final Then
            ObtenerSeccion = i
            Exit Function
        End If
    Next
    ObtenerSeccion = 0
End Function

' ----- Abre Excel con late binding -----
Public Function ExcelAbrir(xlApp As Object, xlWb As Object, xlWs As Object, _
                           ByVal NombreHoja As String) As Boolean
    On Error GoTo ErrHandler
    Set xlApp = CreateObject("Excel.Application")
    xlApp.ScreenUpdating = False
    Set xlWb = xlApp.Workbooks.Add
    Do While xlWb.Worksheets.Count > 1
        xlApp.DisplayAlerts = False
        xlWb.Worksheets(xlWb.Worksheets.Count).Delete
        xlApp.DisplayAlerts = True
    Loop
    Set xlWs = xlWb.Worksheets(1)
    xlWs.Name = Left$(NombreHoja, 31)
    xlWs.Cells.Font.Name = "Calibri"
    xlWs.Cells.Font.Size = 9
    ExcelAbrir = True
    Exit Function
ErrHandler:
    MsgBox "No se pudo iniciar Microsoft Excel." & vbCrLf & Err.Description, vbCritical, "Calcum - Excel"
    ExcelAbrir = False
End Function

' ----- Muestra Excel al usuario -----
Public Sub ExcelMostrar(xlApp As Object, xlWb As Object, xlWs As Object)
    On Error Resume Next
    xlWs.Activate
    xlWs.Range("A1").Select
    xlApp.ScreenUpdating = True
    xlApp.Visible = True
    xlApp.UserControl = True
    xlApp.WindowState = -4137
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
End Sub

' ----- Encabezado estándar -----
Public Function ExcelEncabezado(xlWs As Object, ByVal Titulo As String, _
                                ByVal Periodo As String, ByVal NumColumnas As Integer) As Long
    With xlWs
        .Cells(1, 1).Value = Titulo
        .Range(.Cells(1, 1), .Cells(1, NumColumnas)).Merge
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 11
        .Cells(1, 1).Font.Color = RGB(0, 102, 153)
        .Cells(1, 1).HorizontalAlignment = xlCenter

        .Cells(2, 1).Value = "LIBRO DIARIO DE FORMATO SIMPLIFICADO"
        .Range(.Cells(2, 1), .Cells(2, NumColumnas)).Merge
        .Cells(2, 1).Font.Bold = True
        .Cells(2, 1).Font.Size = 9
        .Cells(2, 1).HorizontalAlignment = xlCenter

        .Cells(4, 1).Value = "RUC:"
        .Cells(4, 2).Value = cRUC
        .Cells(4, 3).Value = "PERIODO:"
        .Cells(4, 4).Value = Periodo
        .Range(.Cells(4, 1), .Cells(4, 4)).Font.Size = 8

    End With
    ExcelEncabezado = 5
End Function

' ----- Títulos de columnas -----
Public Sub ExcelTitulosColumnasF52(xlWs As Object, ByVal Fila As Long, aCuentas() As TipoColumna, NumCuentas As Integer)
    Dim i As Integer
    Dim col As Integer
    
    With xlWs
        .Cells(Fila, 1).Value = "CUENTA"
        .Cells(Fila, 2).Value = "FECHA"
        .Cells(Fila, 3).Value = "CÑR"
        .Cells(Fila, 4).Value = "GLOSA"
        
        col = 5
        For i = 1 To NumCuentas
            .Cells(Fila, col).Value = aCuentas(i).Cuenta
            col = col + 1
        Next
        
        With .Range(.Cells(Fila, 1), .Cells(Fila, col - 1))
            .Font.Bold = True
            .Font.Size = 8
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Interior.Color = RGB(200, 200, 200)
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        End With
        .Rows(Fila).RowHeight = 25
    End With
End Sub

' ----- Configuración de página -----
Public Sub ExcelConfigurarPagina(xlWs As Object, ByVal Orientacion As Integer, _
                                 ByVal FilaTitulos As Long, ByVal NumColumnas As Integer)
    On Error Resume Next
    With xlWs.PageSetup
        .Orientation = IIf(Orientacion = 1, xlLandscape, xlPortrait)
        .PaperSize = xlPaperA4
        .PrintTitleRows = "$1:$" & FilaTitulos
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = xlWs.Application.InchesToPoints(0.3)
        .RightMargin = xlWs.Application.InchesToPoints(0.3)
        .TopMargin = xlWs.Application.InchesToPoints(0.4)
        .BottomMargin = xlWs.Application.InchesToPoints(0.5)
        .CenterFooter = "Página &P de &N"
    End With
    xlWs.Activate
    xlWs.Cells(FilaTitulos + 1, 1).Select
    On Error Resume Next
    xlWs.Application.ActiveWindow.FreezePanes = True
End Sub

' =============================================================================
' FORMATO 5.2 - LIBRO DIARIO SIMPLIFICADO CON COLUMNAS DINÁMICAS Y AGRUPACIÓN
' Orden: Cuenta → Fecha Documento → Correlativo
' =============================================================================
Public Function ExcelDiarioSimplificadoF52(ByVal sTablaDetalle As String, _
                                           ByVal sTablaSubtotales As String, _
                                           ByVal sPeriodo As String) As Boolean
    Dim xlApp As Object, xlWb As Object, xlWs As Object
    Dim rsDetalle As DAO.Recordset, rsSubtot As DAO.Recordset
    Dim aCuentas() As TipoColumna, aFilas() As TipoFila
    Dim nCuentas As Integer, nFilas As Long, fila As Long
    Dim iCol As Integer, i As Long, lFilaIni As Long
    Dim sCuentaAnt As String, iSeccionAnt As Integer, iSeccionActual As Integer
    Dim cTotDebe As Currency, cTotHaber As Currency
    Dim cSubCuentaDebe As Currency, cSubCuentaHaber As Currency
    Dim cSubSeccionDebe As Currency, cSubSeccionHaber As Currency
    Dim col As Integer, val As Currency
    
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    
    ' Inicializa secciones
    Call InitSecciones
    
    ' Abre detalle ordenado por Cuenta → Fecha → Correlativo
    Set rsDetalle = dbWorkArea.OpenRecordset( _
        "SELECT Fecha, Correlativo, Glosa, Cuenta, Debe, Haber FROM " & sTablaDetalle & _
        " ORDER BY Cuenta, Fecha, Correlativo", dbOpenSnapshot)
    
    If rsDetalle.EOF Then
        Screen.MousePointer = vbArrow
        MsgBox "No hay información para exportar.", vbInformation, "Calcum - Excel"
        Exit Function
    End If
    
    ' Obtiene lista única de cuentas
    Set rsSubtot = dbWorkArea.OpenRecordset( _
        "SELECT DISTINCT Cuenta FROM " & sTablaSubtotales & " ORDER BY Cuenta", dbOpenSnapshot)
    
    nCuentas = 0
    Do While Not rsSubtot.EOF
        nCuentas = nCuentas + 1
        ReDim Preserve aCuentas(1 To nCuentas)
        aCuentas(nCuentas).Cuenta = rsSubtot!Cuenta & ""
        aCuentas(nCuentas).Indice = nCuentas
        rsSubtot.MoveNext
    Loop
    rsSubtot.Close
    
    ' Arma arreglo de filas
    rsDetalle.MoveLast: nFilas = rsDetalle.RecordCount: rsDetalle.MoveFirst
    ReDim aFilas(1 To nFilas * 3 + 100) ' Con espacio para separadores y subtotales
    ReDim aFilas(1).Montos(1 To nCuentas)
    
    fila = 0
    sCuentaAnt = ""
    iSeccionAnt = 0
    
    Do While Not rsDetalle.EOF
        iSeccionActual = ObtenerSeccion(rsDetalle!Cuenta & "")
        
        ' Si cambió sección y no es la primera, agrega subtotal de sección
        If iSeccionActual <> iSeccionAnt And iSeccionAnt <> 0 Then
            GoSub PonerSubtotalSeccion
        End If
        
        ' Si cambió sección, agrega encabezado
        If iSeccionActual <> iSeccionAnt Then
            fila = fila + 1
            aFilas(fila).EsSubtotalSeccion = False
            aFilas(fila).EsSubtotalCuenta = False
            aFilas(fila).TextoEncabezado = aSecciones(iSeccionActual).Nombre
            iSeccionAnt = iSeccionActual
        End If
        
        ' Si cambió cuenta, agrega subtotal anterior
        If rsDetalle!Cuenta & "" <> sCuentaAnt And sCuentaAnt <> "" Then
            GoSub PonerSubtotalCuenta
        End If
        
        ' Agrega fila de detalle
        fila = fila + 1
        If fila > UBound(aFilas) Then ReDim Preserve aFilas(1 To fila + 50)
        ReDim aFilas(fila).Montos(1 To nCuentas)
        
        aFilas(fila).Fecha = rsDetalle!Fecha & ""
        aFilas(fila).Correlativo = rsDetalle!Correlativo & ""
        aFilas(fila).Cuenta = rsDetalle!Cuenta & ""
        aFilas(fila).Glosa = rsDetalle!Glosa & ""
        
        ' Busca el índice de columna para esta cuenta
        iCol = BuscaIndiceCuenta(aCuentas, nCuentas, rsDetalle!Cuenta & "")
        If iCol > 0 Then
            If rsDetalle!Debe & "" <> "" Then
                aFilas(fila).Montos(iCol) = aFilas(fila).Montos(iCol) + CCur(rsDetalle!Debe)
                cSubCuentaDebe = cSubCuentaDebe + CCur(rsDetalle!Debe)
                cSubSeccionDebe = cSubSeccionDebe + CCur(rsDetalle!Debe)
                cTotDebe = cTotDebe + CCur(rsDetalle!Debe)
            End If
            If rsDetalle!Haber & "" <> "" Then
                aFilas(fila).Montos(iCol) = aFilas(fila).Montos(iCol) - CCur(rsDetalle!Haber)
                cSubCuentaHaber = cSubCuentaHaber + CCur(rsDetalle!Haber)
                cSubSeccionHaber = cSubSeccionHaber + CCur(rsDetalle!Haber)
                cTotHaber = cTotHaber + CCur(rsDetalle!Haber)
            End If
        End If
        
        sCuentaAnt = rsDetalle!Cuenta & ""
        rsDetalle.MoveNext
    Loop
    
    ' Último subtotal de cuenta
    If sCuentaAnt <> "" Then GoSub PonerSubtotalCuenta
    ' Último subtotal de sección
    If iSeccionAnt <> 0 Then GoSub PonerSubtotalSeccion
    
    ' Agrega total general
    fila = fila + 1
    If fila > UBound(aFilas) Then ReDim Preserve aFilas(1 To fila + 10)
    ReDim aFilas(fila).Montos(1 To nCuentas)
    aFilas(fila).EsTotalGeneral = True
    aFilas(fila).Montos(1) = cTotDebe
    If nCuentas > 1 Then aFilas(fila).Montos(2) = cTotHaber
    
    nFilas = fila
    
    ' Crea Excel
    If Not ExcelAbrir(xlApp, xlWb, xlWs, "Formato 5.2") Then GoTo Salir
    
    lFilaIni = ExcelEncabezado(xlWs, "TEXTILES TBM S.A.C.", sPeriodo, nCuentas + 4)
    ExcelTitulosColumnasF52 xlWs, lFilaIni, aCuentas, nCuentas
    
    ' Escribe datos en Excel
    fila = 0
    For i = 1 To nFilas
        fila = fila + 1
        With xlWs
            If aFilas(i).TextoEncabezado <> "" Then
                ' Encabezado de sección
                .Cells(lFilaIni + fila, 1).Value = aFilas(i).TextoEncabezado
                .Range(.Cells(lFilaIni + fila, 1), .Cells(lFilaIni + fila, nCuentas + 4)).Merge
                .Cells(lFilaIni + fila, 1).Font.Bold = True
                .Cells(lFilaIni + fila, 1).Font.Size = 9
                .Cells(lFilaIni + fila, 1).Interior.Color = RGB(220, 230, 240)
                .Rows(lFilaIni + fila).RowHeight = 18
            ElseIf aFilas(i).EsSubtotalCuenta Then
                ' Subtotal de cuenta
                .Cells(lFilaIni + fila, 4).Value = "SUBTOTAL"
                .Cells(lFilaIni + fila, 4).Font.Bold = True
                For col = 1 To nCuentas
                    If aFilas(i).Montos(col) <> 0 Then
                        .Cells(lFilaIni + fila, 4 + col).Value = aFilas(i).Montos(col)
                        .Cells(lFilaIni + fila, 4 + col).NumberFormat = kFMT_MONTO
                        .Cells(lFilaIni + fila, 4 + col).Font.Bold = True
                    End If
                Next
                .Rows(lFilaIni + fila).RowHeight = 16
            ElseIf aFilas(i).EsSubtotalSeccion Then
                ' Subtotal de sección
                .Cells(lFilaIni + fila, 4).Value = "TOTAL SECCION"
                .Cells(lFilaIni + fila, 4).Font.Bold = True
                .Cells(lFilaIni + fila, 4).Font.Color = RGB(255, 255, 255)
                .Range(.Cells(lFilaIni + fila, 1), .Cells(lFilaIni + fila, nCuentas + 4)).Interior.Color = RGB(100, 100, 100)
                For col = 1 To nCuentas
                    If aFilas(i).Montos(col) <> 0 Then
                        .Cells(lFilaIni + fila, 4 + col).Value = aFilas(i).Montos(col)
                        .Cells(lFilaIni + fila, 4 + col).NumberFormat = kFMT_MONTO
                        .Cells(lFilaIni + fila, 4 + col).Font.Bold = True
                        .Cells(lFilaIni + fila, 4 + col).Font.Color = RGB(255, 255, 255)
                    End If
                Next
                .Rows(lFilaIni + fila).RowHeight = 18
            ElseIf aFilas(i).EsTotalGeneral Then
                ' Total general
                .Cells(lFilaIni + fila, 4).Value = "TOTAL GENERAL"
                .Cells(lFilaIni + fila, 4).Font.Bold = True
                .Range(.Cells(lFilaIni + fila, 1), .Cells(lFilaIni + fila, nCuentas + 4)).Interior.Color = RGB(0, 102, 153)
                .Range(.Cells(lFilaIni + fila, 1), .Cells(lFilaIni + fila, nCuentas + 4)).Font.Color = RGB(255, 255, 255)
                For col = 1 To nCuentas
                    If aFilas(i).Montos(col) <> 0 Then
                        .Cells(lFilaIni + fila, 4 + col).Value = aFilas(i).Montos(col)
                        .Cells(lFilaIni + fila, 4 + col).NumberFormat = kFMT_MONTO
                        .Cells(lFilaIni + fila, 4 + col).Font.Bold = True
                        .Cells(lFilaIni + fila, 4 + col).Font.Color = RGB(255, 255, 255)
                    End If
                Next
                .Rows(lFilaIni + fila).RowHeight = 18
            Else
                ' Fila de detalle
                .Cells(lFilaIni + fila, 1).Value = aFilas(i).Cuenta
                .Cells(lFilaIni + fila, 2).Value = aFilas(i).Fecha
                .Cells(lFilaIni + fila, 2).NumberFormat = kFMT_FECHA
                .Cells(lFilaIni + fila, 3).Value = aFilas(i).Correlativo
                .Cells(lFilaIni + fila, 4).Value = aFilas(i).Glosa
                .Cells(lFilaIni + fila, 1).HorizontalAlignment = xlCenter
                .Cells(lFilaIni + fila, 2).HorizontalAlignment = xlCenter
                .Cells(lFilaIni + fila, 3).HorizontalAlignment = xlCenter
                For col = 1 To nCuentas
                    If aFilas(i).Montos(col) <> 0 Then
                        .Cells(lFilaIni + fila, 4 + col).Value = aFilas(i).Montos(col)
                        .Cells(lFilaIni + fila, 4 + col).NumberFormat = kFMT_MONTO
                    End If
                Next
            End If
        End With
    Next
    
    ExcelConfigurarPagina xlWs, 1, lFilaIni, nCuentas + 4
    ExcelMostrar xlApp, xlWb, xlWs
    ExcelDiarioSimplificadoF52 = True

Salir:
    On Error Resume Next
    rsDetalle.Close: Set rsDetalle = Nothing
    rsSubtot.Close: Set rsSubtot = Nothing
    Screen.MousePointer = vbArrow
    Exit Function

PonerSubtotalCuenta:
    fila = fila + 1
    If fila > UBound(aFilas) Then ReDim Preserve aFilas(1 To fila + 50)
    ReDim aFilas(fila).Montos(1 To nCuentas)
    aFilas(fila).EsSubtotalCuenta = True
    For col = 1 To nCuentas
        aFilas(fila).Montos(col) = 0
    Next
    ' Llena solo la primera columna de montos si es Debe/Haber
    aFilas(fila).Montos(1) = cSubCuentaDebe
    If nCuentas > 1 Then aFilas(fila).Montos(2) = cSubCuentaHaber
    cSubCuentaDebe = 0: cSubCuentaHaber = 0
    Return

PonerSubtotalSeccion:
    fila = fila + 1
    If fila > UBound(aFilas) Then ReDim Preserve aFilas(1 To fila + 50)
    ReDim aFilas(fila).Montos(1 To nCuentas)
    aFilas(fila).EsSubtotalSeccion = True
    aFilas(fila).Montos(1) = cSubSeccionDebe
    If nCuentas > 1 Then aFilas(fila).Montos(2) = cSubSeccionHaber
    cSubSeccionDebe = 0: cSubSeccionHaber = 0
    Return

ErrHandler:
    MsgBox "Error al generar Excel: " & Err.Description & vbCrLf & "Línea " & Erl, vbCritical, "Calcum - Excel"
    If Not xlApp Is Nothing Then ExcelMostrar xlApp, xlWb, xlWs
    Resume Salir
End Function

Private Function BuscaIndiceCuenta(aCuentas() As TipoColumna, nCuentas As Integer, sCuenta As String) As Integer
    Dim i As Integer
    For i = 1 To nCuentas
        If aCuentas(i).Cuenta = sCuenta Then
            BuscaIndiceCuenta = i
            Exit Function
        End If
    Next
    BuscaIndiceCuenta = 0
End Function

' =============================================================================
' Exportación genérica para otros reportes
' =============================================================================
Public Function ExcelDesdeConsulta(dbOrigen As DAO.Database, ByVal sSql As String, _
                                   ByVal Titulo As String, ByVal Periodo As String, _
                                   ByVal Orientacion As Integer) As Boolean
    Dim xlApp As Object, xlWb As Object, xlWs As Object
    Dim rs As DAO.Recordset
    Dim aTit() As Variant, aAnc() As Variant, aDatos() As Variant
    Dim nCol As Integer, c As Integer, f As Long, nRegs As Long, lFila As Long

    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    Set rs = dbOrigen.OpenRecordset(sSql, dbOpenSnapshot)
    If rs.EOF Then
        Screen.MousePointer = vbArrow
        MsgBox "No hay información para exportar.", vbInformation, "Calcum - Excel"
        Exit Function
    End If
    rs.MoveLast: nRegs = rs.RecordCount: rs.MoveFirst
    nCol = rs.Fields.Count

    ReDim aTit(1 To nCol): ReDim aAnc(1 To nCol)
    For c = 1 To nCol
        aTit(c) = UCase$(Replace(rs.Fields(c - 1).Name, "_", " "))
        Select Case rs.Fields(c - 1).Type
            Case dbCurrency, dbDouble, dbSingle, dbDecimal: aAnc(c) = 14
            Case dbDate: aAnc(c) = 11
            Case Else: aAnc(c) = IIf(rs.Fields(c - 1).Size > 30, 40, 14)
        End Select
    Next

    If Not ExcelAbrir(xlApp, xlWb, xlWs, "Reporte") Then GoTo Salir
    lFila = ExcelEncabezado(xlWs, Titulo, Periodo, nCol)

    ReDim aDatos(1 To nRegs, 1 To nCol)
    Do While Not rs.EOF
        f = f + 1
        For c = 1 To nCol
            If rs.Fields(c - 1).Type = dbText Then
                aDatos(f, c) = "'" & rs.Fields(c - 1).Value
            Else
                aDatos(f, c) = rs.Fields(c - 1).Value
            End If
        Next
        rs.MoveNext
    Loop
    With xlWs
        .Range(.Cells(lFila + 1, 1), .Cells(lFila + nRegs, nCol)).Value = aDatos
        For c = 1 To nCol
            Select Case rs.Fields(c - 1).Type
                Case dbCurrency, dbDouble, dbSingle, dbDecimal
                    .Range(.Cells(lFila + 1, c), .Cells(lFila + nRegs, c)).NumberFormat = kFMT_MONTO
                Case dbDate
                    .Range(.Cells(lFila + 1, c), .Cells(lFila + nRegs, c)).NumberFormat = kFMT_FECHA
            End Select
        Next
    End With

    ExcelConfigurarPagina xlWs, Orientacion, lFila, nCol
    ExcelMostrar xlApp, xlWb, xlWs
    ExcelDesdeConsulta = True
Salir:
    On Error Resume Next
    rs.Close: Set rs = Nothing
    Screen.MousePointer = vbArrow
    Exit Function
ErrHandler:
    MsgBox "Error al generar Excel: " & Err.Description, vbCritical, "Calcum - Excel"
    If Not xlApp Is Nothing Then ExcelMostrar xlApp, xlWb, xlWs
    Resume Salir
End Function
