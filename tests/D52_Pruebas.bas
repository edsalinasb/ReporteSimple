Attribute VB_Name = "D52_Pruebas"
Option Explicit
' Pruebas de regresion de DllRSimple_Diario52 (logica pura del Formato 5.2).
' Ejecutables:
'   - Linux/CI: python3 tests/run_tests.py  (LibreOffice Basic, Option VBASupport 1)
'   - VB6: agregar este modulo y RSimple_Diario52.bas a un proyecto EXE de prueba y
'          ejecutar  Debug.Print D52_EjecutarPruebas()
' No ejecutan DAO, Crystal ni Excel.

Private mFallas As Long
Private mPruebas As Long
Private mLog As String

Private Sub Verificar(ByVal bOk As Boolean, ByVal sNombre As String)
    mPruebas = mPruebas + 1
    If bOk Then
        mLog = mLog & "OK    " & sNombre & vbLf
    Else
        mFallas = mFallas + 1
        mLog = mLog & "FALLA " & sNombre & vbLf
    End If
End Sub

Private Function Cta(ByVal sPrefijo As String, ByVal n As Integer) As String
    Cta = sPrefijo & Right$("00" & CStr(n), 2)
End Function

Public Function D52_EjecutarPruebas() As String
    mFallas = 0: mPruebas = 0: mLog = ""
    PruebaClasificacion
    PruebaBloquesMasDe14
    PruebaRevisitaCuenta
    PruebaLimitesSeccion
    PruebaOrdenInvalido
    PruebaCuentaInvalida
    PruebaMontosYCAR
    PruebaCombinacionYTotales
    PruebaOrdenFilas
    PruebaReinicio
    D52_EjecutarPruebas = mLog & "RESUMEN: " & mPruebas & " pruebas, " & mFallas & " fallas" & vbLf
End Function

Private Sub PruebaClasificacion()
    Verificar D52_ClasificarCuenta("10") = 1, "10 Activo Corriente"
    Verificar D52_ClasificarCuenta("2999") = 1, "29 Activo Corriente"
    Verificar D52_ClasificarCuenta("30") = 2, "30 Activo No Corriente"
    Verificar D52_ClasificarCuenta("39") = 2, "39 Activo No Corriente"
    Verificar D52_ClasificarCuenta("40111") = 3, "40 Pasivo Corriente"
    Verificar D52_ClasificarCuenta("46") = 3, "46 Pasivo Corriente"
    Verificar D52_ClasificarCuenta("47") = 4, "47 Pasivo No Corriente"
    Verificar D52_ClasificarCuenta("49") = 4, "49 Pasivo No Corriente"
    Verificar D52_ClasificarCuenta("50") = 5, "50 Patrimonio"
    Verificar D52_ClasificarCuenta("59") = 5, "59 Patrimonio"
    Verificar D52_ClasificarCuenta("60") = 6, "60 Gestion Egresos"
    Verificar D52_ClasificarCuenta("69") = 6, "69 Gestion Egresos"
    Verificar D52_ClasificarCuenta("7911101") = 7, "79 Gestion Ingresos"
    Verificar D52_ClasificarCuenta("80") = 8, "80 Saldos Intermedios"
    Verificar D52_ClasificarCuenta("89") = 8, "89 Saldos Intermedios"
    Verificar D52_ClasificarCuenta("90") = 9, "90 Analitica"
    Verificar D52_ClasificarCuenta("99") = 9, "99 Analitica"
    Verificar D52_ClasificarCuenta("0111") = 10, "0111 Cuentas de Orden"
    Verificar D52_ClasificarCuenta("09") = 10, "09 Cuentas de Orden"
    Verificar D52_ClasificarCuenta("00") = 10, "00 Cuentas de Orden"
    Verificar D52_ClasificarCuenta("") = 0, "vacia invalida"
    Verificar D52_ClasificarCuenta("1") = 0, "un digito invalida"
    Verificar D52_ClasificarCuenta(" 10") = 0, "espacio inicial invalida"
    Verificar D52_ClasificarCuenta("A1") = 0, "letra invalida"
    Verificar D52_ClasificarCuenta(Null) = 0, "Null invalida"
    Verificar D52_NombreSeccion(1) = "Activo Corriente", "nombre seccion 1"
    Verificar D52_NombreSeccion(9) = "Analitica de Explotacion", "nombre seccion 9"
    Verificar D52_NombreSeccion(10) = "Cuentas de Orden", "nombre seccion 10 (ultima)"
End Sub

Private Sub PruebaBloquesMasDe14()
    Dim i As Integer, iCol As Integer, lCab As Long, bOk As Boolean
    D52_Iniciar
    bOk = True
    For i = 1 To 20
        lCab = D52_UbicarCuenta(Cta("101", i), iCol)
        If i <= 14 Then
            If lCab <> 1 Or iCol <> i - 1 Then bOk = False
        Else
            If lCab <> 2 Or iCol <> i - 15 Then bOk = False
        End If
    Next
    Verificar bOk, ">14 cuentas: columnas 0..13 y bloque de continuacion"
    Verificar D52_NumBloques() = 2, ">14 cuentas: 2 bloques"
    Verificar D52_BloqueNumCuentas(1) = 14 And D52_BloqueNumCuentas(2) = 6, ">14 cuentas: 14 + 6"
    Verificar D52_EsContinuacion(2) And Not D52_EsContinuacion(1), "bloque 2 es continuacion"
    Verificar D52_BloqueCuenta(2, 0) = "10115", "cabecera conserva codigo"
End Sub

Private Sub PruebaRevisitaCuenta()
    Dim i As Integer, iCol As Integer, lCab As Long
    D52_Iniciar
    For i = 1 To 16
        lCab = D52_UbicarCuenta(Cta("121", i), iCol)
    Next
    lCab = D52_UbicarCuenta("12103", iCol)
    Verificar lCab = 1 And iCol = 2, "revisita en bloque anterior usa columna real"
    lCab = D52_UbicarCuenta("12115", iCol)
    Verificar lCab = 2 And iCol = 0, "revisita en bloque actual usa columna real"
    Verificar D52_NumBloques() = 2 And D52_BloqueNumCuentas(2) = 2, "revisita no agrega columnas"
End Sub

Private Sub PruebaLimitesSeccion()
    Dim iCol As Integer, lCab As Long
    D52_Iniciar
    lCab = D52_UbicarCuenta("2911", iCol)
    lCab = D52_UbicarCuenta("3011", iCol)
    Verificar lCab = 2 And iCol = 0, "29 -> 30 abre bloque nuevo"
    lCab = D52_UbicarCuenta("4699", iCol)
    lCab = D52_UbicarCuenta("4711", iCol)
    Verificar lCab = 4 And D52_BloqueSeccion(4) = 4, "46 -> 47 abre bloque nuevo"
    lCab = D52_UbicarCuenta("0111", iCol)
    Verificar D52_BloqueSeccion(lCab) = 10, "cuentas de orden al final"
    Verificar Not D52_EsContinuacion(lCab), "seccion nueva no es continuacion"
End Sub

Private Sub PruebaOrdenInvalido()
    Dim iCol As Integer, lCab As Long, lErr As Long
    D52_Iniciar
    lCab = D52_UbicarCuenta("6011", iCol)
    On Error Resume Next
    lCab = D52_UbicarCuenta("1011", iCol)
    lErr = Err.Number
    On Error GoTo 0
    Verificar lErr = D52_ERR_ORDEN, "seccion hacia atras produce error"
End Sub

Private Sub PruebaCuentaInvalida()
    Dim iCol As Integer, lCab As Long, lErr As Long
    D52_Iniciar
    On Error Resume Next
    lCab = D52_UbicarCuenta("   ", iCol)
    lErr = Err.Number
    On Error GoTo 0
    Verificar lErr = D52_ERR_CUENTA_INVALIDA, "cuenta en blanco produce error"
    Verificar D52_NumBloques() = 0, "cuenta invalida no crea bloque"
End Sub

Private Sub PruebaMontosYCAR()
    Verificar D52_Monto(100, Null) = 100, "monto debe con haber Null"
    Verificar D52_Monto(Null, 25.5) = -25.5, "monto haber negativo"
    Verificar D52_Monto(Null, Null) = 0, "monto Null/Null = 0"
    Verificar D52_Monto(0.1, 0.3) = -0.2, "monto Currency exacto"
    Verificar D52_SI("001000000") = "1" And D52_SI("005000123") = "2", "SI saldo inicial"
    Verificar D52_CAR("701000045", "01", " F001 ", "00000123", "", "20123456789") = "2012345678901F00100000123", "CAR ventas"
    Verificar D52_CAR("601000045", "01", "E001", "00000077", "20999999991", "20123456789") = "2099999999101E00100000077", "CAR compras"
    Verificar D52_CAR("005000001", "01", "E001", "1", "X", "20123456789") = "005000001", "CAR otros conserva TipoNumero"
    Verificar D52_Texto(Null) = "" And D52_Texto("0012") = "0012", "texto Null-safe y ceros"
End Sub

Private Sub PruebaCombinacionYTotales()
    Dim iCol As Integer, lCab As Long, dF As Date, k As Long, f As Long
    Dim cSuma As Currency, i As Integer
    D52_Iniciar
    dF = DateSerial(2026, 3, 15)
    ' Voucher 005000001: debe 6011 100 y 6311 50 (mismo bloque), haber 4011 -150 (otra seccion)
    lCab = D52_UbicarCuenta("4011", iCol)
    D52_AgregarDetalle lCab, iCol, "4011", dF, dF, "2", "005000001", "Pago", -150
    lCab = D52_UbicarCuenta("6011", iCol)
    D52_AgregarDetalle lCab, iCol, "6011", dF, dF, "2", "005000001", "Pago", 100
    lCab = D52_UbicarCuenta("6311", iCol)
    D52_AgregarDetalle lCab, iCol, "6311", dF, dF, "2", "005000001", "Pago", 50
    ' Misma cuenta, mismo voucher: credito se separa del debito
    D52_AgregarDetalle lCab, iCol, "6311", dF, dF, "2", "005000001", "Pago", -20
    D52_AgregarDetalle lCab, iCol, "6311", dF, dF, "2", "005000001", "Pago", 20
    D52_ConstruirFilas
    Verificar D52_NumFilas() = 3, "combinacion: debitos juntos, creditos aparte, por bloque"
    cSuma = 0
    For k = 1 To D52_NumFilas()
        f = D52_FilaOrdenada(k)
        For i = 0 To D52_COLUMNAS - 1
            cSuma = cSuma + D52_FilaMonto(f, i)
        Next
    Next
    Verificar cSuma = D52_TotalDetalle() And cSuma = 0, "cuadre: filas = detalle = 0 (sin duplicar)"
    Verificar D52_TotalColumna(2, 0) = 100 And D52_TotalColumna(2, 1) = 50, "total por columna con signo"
    Verificar D52_TotalColumna(1, 0) = -150, "total columna credito negativo"
    f = D52_FilaOrdenada(2)
    Verificar D52_FilaMonto(f, 0) = 100 And D52_FilaMonto(f, 1) = 70, "fila debito combina 2 cuentas"
End Sub

Private Sub PruebaOrdenFilas()
    Dim iCol As Integer, lCab As Long, f As Long
    D52_Iniciar
    ' Cuenta 1011 (col 0) y 1211 (col 1); fechas en distinto anio (orden textual dd/mm fallaria)
    lCab = D52_UbicarCuenta("1011", iCol)
    D52_AgregarDetalle lCab, iCol, "1011", DateSerial(2026, 1, 5), DateSerial(2026, 3, 1), "2", "B", "g1", 10
    D52_AgregarDetalle lCab, iCol, "1011", DateSerial(2025, 12, 20), DateSerial(2026, 3, 1), "2", "C", "g2", 10
    D52_AgregarDetalle lCab, iCol, "1011", DateSerial(2026, 2, 28), DateSerial(2026, 2, 28), "1", "001000000", "Saldo Inicial", 5
    D52_AgregarDetalle lCab, iCol, "1011", DateSerial(2026, 1, 5), DateSerial(2026, 3, 1), "2", "A", "g3", 10
    lCab = D52_UbicarCuenta("1211", iCol)
    D52_AgregarDetalle lCab, iCol, "1211", DateSerial(2024, 1, 1), DateSerial(2026, 3, 1), "2", "Z", "g4", 7
    D52_AgregarDetalle lCab, iCol, "1211", Null, DateSerial(2026, 3, 1), "2", "Y", "g5", 7
    D52_ConstruirFilas
    Verificar D52_NumFilas() = 6, "orden: 6 filas"
    f = D52_FilaOrdenada(1): Verificar D52_FilaSI(f) = "1", "orden: saldo inicial primero en la cuenta"
    f = D52_FilaOrdenada(2): Verificar D52_FilaCAR(f) = "C", "orden: fecha 20/12/2025 antes que 05/01/2026"
    f = D52_FilaOrdenada(3): Verificar D52_FilaCAR(f) = "A", "orden: desempate por CAR (A)"
    f = D52_FilaOrdenada(4): Verificar D52_FilaCAR(f) = "B", "orden: desempate por CAR (B)"
    f = D52_FilaOrdenada(5): Verificar D52_FilaCAR(f) = "Y", "orden: cuenta siguiente, fecha Null primero"
    f = D52_FilaOrdenada(6): Verificar D52_FilaCAR(f) = "Z" And D52_FilaCabecera(f) = 1, "orden: segunda cuenta despues aunque fecha menor"
End Sub

Private Sub PruebaReinicio()
    Dim iCol As Integer, lCab As Long
    D52_Iniciar
    lCab = D52_UbicarCuenta("1011", iCol)
    D52_AgregarDetalle lCab, iCol, "1011", Null, Null, "2", "A", "x", 1
    D52_MarcarListo True
    D52_Iniciar
    Verificar D52_NumBloques() = 0 And D52_NumDetalles() = 0 And Not D52_Listo(), "reinicio limpia el modelo (exportacion repetida)"
    lCab = D52_UbicarCuenta("1011", iCol)
    Verificar lCab = 1 And iCol = 0, "reinicio: numeracion de cabecera desde 1"
End Sub
