Attribute VB_Name = "DllRSimple_Diario52"
Option Explicit
' =============================================================================
' Libro Diario Simplificado - Formato 5.2
' Logica pura (sin DAO, sin Excel, sin formularios) compartida por:
'   - frmRSimple.PreparaDetalledelDiarioSimplificado (tablas para Crystal)
'   - DllRSimple_Excel.ExcelDiarioSimplificadoF52 (vista normalizada en memoria)
' Se usan arreglos paralelos (no UDT) para que el modulo pueda ejecutarse
' tambien en el arnes de pruebas (tests/) sin VB6.
' =============================================================================

Public Const D52_NUM_SECCIONES As Integer = 10
Public Const D52_COLUMNAS As Integer = 14            ' Cuentas por bloque (S1..S14)
Public Const D52_SECCION_ORDEN As Integer = 10        ' Cuentas de Orden (clase 0)
Public Const D52_TIPONUMERO_SALDO_INICIAL As String = "001000000"

Public Const D52_ERR_CUENTA_INVALIDA As Long = vbObjectError + 5201
Public Const D52_ERR_ORDEN As Long = vbObjectError + 5202
Public Const D52_ERR_ESTADO As Long = vbObjectError + 5203

' ----- Bloques (cabeceras de hasta 14 cuentas) -----
Private mNumBloques As Long
Private mBloqueSeccion() As Integer
Private mBloqueNumCtas() As Integer
Private mBloqueCtas() As String          ' (0 To 13, 1 To n)

' ----- Detalle normalizado (una fila por linea contable) -----
Private mNumDet As Long
Private mDetCab() As Long
Private mDetCol() As Integer
Private mDetCta() As String
Private mDetFecDoc() As Variant
Private mDetFecOp() As Variant
Private mDetSI() As String
Private mDetCAR() As String
Private mDetGlosa() As String
Private mDetMonto() As Currency

' ----- Filas combinadas para Excel -----
Private mNumFil As Long
Private mFilCab() As Long
Private mFilMinCol() As Integer
Private mFilSI() As String
Private mFilFecDoc() As Variant
Private mFilCAR() As String
Private mFilGlosa() As String
Private mFilSeq() As Long
Private mFilMontos() As Currency         ' (0 To 13, 1 To n)
Private mFilOrden() As Long               ' indices ordenados

Private mListo As Boolean
Private mFechasTipadas As Boolean
Private mDiagnostico As String

' =============================================================================
' Secciones (rango de los dos primeros digitos de la cuenta)
' =============================================================================
Public Function D52_NombreSeccion(ByVal iSeccion As Integer) As String
    Select Case iSeccion
        Case 1: D52_NombreSeccion = "Activo Corriente"
        Case 2: D52_NombreSeccion = "Activo No Corriente"
        Case 3: D52_NombreSeccion = "Pasivo Corriente"
        Case 4: D52_NombreSeccion = "Pasivo No Corriente"
        Case 5: D52_NombreSeccion = "Patrimonio"
        Case 6: D52_NombreSeccion = "Gestión Egresos"
        Case 7: D52_NombreSeccion = "Gestión Ingresos"
        Case 8: D52_NombreSeccion = "Saldos Intermedios"
        Case 9: D52_NombreSeccion = "Analitica de Explotacion"
        Case 10: D52_NombreSeccion = "Cuentas de Orden"
        Case Else: D52_NombreSeccion = ""
    End Select
End Function

Private Function EsDigito(ByVal s As String) As Boolean
    EsDigito = (Len(s) = 1 And s >= "0" And s <= "9")
End Function

' Devuelve 1..10 en el orden de presentacion, o 0 si la cuenta no es valida.
' Regla: los dos primeros caracteres deben ser digitos (se conservan ceros a la
' izquierda). Clase 0 (01..09) = Cuentas de Orden; 10..99 segun rangos.
Public Function D52_ClasificarCuenta(ByVal vCuenta As Variant) As Integer
    Dim s As String, n As Integer
    D52_ClasificarCuenta = 0
    If IsNull(vCuenta) Or IsEmpty(vCuenta) Then Exit Function
    s = CStr(vCuenta)
    If Len(s) < 2 Then Exit Function
    If Not EsDigito(Mid$(s, 1, 1)) Or Not EsDigito(Mid$(s, 2, 1)) Then Exit Function
    If Mid$(s, 1, 1) = "0" Then
        D52_ClasificarCuenta = D52_SECCION_ORDEN
        Exit Function
    End If
    n = CInt(Mid$(s, 1, 2))
    Select Case n
        Case 10 To 29: D52_ClasificarCuenta = 1
        Case 30 To 39: D52_ClasificarCuenta = 2
        Case 40 To 46: D52_ClasificarCuenta = 3
        Case 47 To 49: D52_ClasificarCuenta = 4
        Case 50 To 59: D52_ClasificarCuenta = 5
        Case 60 To 69: D52_ClasificarCuenta = 6
        Case 70 To 79: D52_ClasificarCuenta = 7
        Case 80 To 89: D52_ClasificarCuenta = 8
        Case 90 To 99: D52_ClasificarCuenta = 9
    End Select
End Function

' =============================================================================
' Utilitarios de campos
' =============================================================================
Public Function D52_Texto(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        D52_Texto = ""
    Else
        D52_Texto = CStr(v)
    End If
End Function

Public Function D52_Monto(ByVal vDebe As Variant, ByVal vHaber As Variant) As Currency
    Dim cD As Currency, cH As Currency
    If Not (IsNull(vDebe) Or IsEmpty(vDebe)) Then cD = CCur(vDebe)
    If Not (IsNull(vHaber) Or IsEmpty(vHaber)) Then cH = CCur(vHaber)
    D52_Monto = cD - cH
End Function

Public Function D52_SI(ByVal sTipoNumero As String) As String
    If sTipoNumero = D52_TIPONUMERO_SALDO_INICIAL Then
        D52_SI = "1"
    Else
        D52_SI = "2"
    End If
End Function

' CAR imprimible (mismas reglas del legado):
'   Ventas  (TipoNumero 7*): RUC empresa (11) + Tipo + Serie + Numero de documento
'   Compras (TipoNumero 6*): Auxiliar del voucher + Tipo + Serie + Numero
'   Otros: TipoNumero original
Public Function D52_CAR(ByVal sTipoNumero As String, ByVal sTipoDoc As String, _
                        ByVal sSerie As String, ByVal sNumDoc As String, _
                        ByVal sAuxCompra As String, ByVal sRUC As String) As String
    Select Case Left$(sTipoNumero, 1)
        Case "7"
            D52_CAR = Right$(sRUC, 11) & sTipoDoc & Trim$(sSerie) & sNumDoc
        Case "6"
            D52_CAR = sAuxCompra & sTipoDoc & Trim$(sSerie) & sNumDoc
        Case Else
            D52_CAR = sTipoNumero
    End Select
End Function

' =============================================================================
' Estado
' =============================================================================
Public Sub D52_Iniciar()
    mNumBloques = 0
    Erase mBloqueSeccion, mBloqueNumCtas, mBloqueCtas
    mNumDet = 0
    Erase mDetCab, mDetCol, mDetCta, mDetFecDoc, mDetFecOp, mDetSI, mDetCAR, mDetGlosa, mDetMonto
    mNumFil = 0
    Erase mFilCab, mFilMinCol, mFilSI, mFilFecDoc, mFilCAR, mFilGlosa, mFilSeq, mFilMontos, mFilOrden
    mListo = False
    mFechasTipadas = True
    mDiagnostico = ""
End Sub

Public Sub D52_MarcarListo(ByVal bListo As Boolean)
    mListo = bListo
End Sub

Public Function D52_Listo() As Boolean
    D52_Listo = mListo
End Function

Public Sub D52_EstablecerFechasTipadas(ByVal bTipadas As Boolean)
    mFechasTipadas = bTipadas
End Sub

Public Function D52_FechasTipadas() As Boolean
    D52_FechasTipadas = mFechasTipadas
End Function

Public Sub D52_EstablecerDiagnostico(ByVal sTexto As String)
    mDiagnostico = sTexto
End Sub

Public Function D52_Diagnostico() As String
    D52_Diagnostico = mDiagnostico
End Function

' =============================================================================
' Bloques de cabecera
' =============================================================================
' Ubica la cuenta en su bloque (Cabecera) y columna 0..13. Crea bloque nuevo al
' cambiar de seccion o cuando el bloque actual ya tiene 14 cuentas. Una cuenta
' ya vista en la seccion actual reutiliza su columna real (no la ultima usada).
Public Function D52_UbicarCuenta(ByVal sCuenta As String, ByRef iColumna As Integer) As Long
    Dim iSec As Integer, l As Long, i As Integer

    iSec = D52_ClasificarCuenta(sCuenta)
    If iSec = 0 Then
        Err.Raise D52_ERR_CUENTA_INVALIDA, "D52_UbicarCuenta", _
                  "Cuenta contable inválida o vacía: '" & sCuenta & "'."
    End If
    If mNumBloques > 0 Then
        If iSec < mBloqueSeccion(mNumBloques) Then
            Err.Raise D52_ERR_ORDEN, "D52_UbicarCuenta", _
                      "Detalle no ordenado por sección/cuenta (cuenta '" & sCuenta & "')."
        End If
        l = mNumBloques
        Do While l >= 1
            If mBloqueSeccion(l) <> iSec Then Exit Do
            For i = 0 To mBloqueNumCtas(l) - 1
                If mBloqueCtas(i, l) = sCuenta Then
                    iColumna = i
                    D52_UbicarCuenta = l
                    Exit Function
                End If
            Next
            l = l - 1
        Loop
    End If

    If mNumBloques = 0 Then
        NuevoBloque iSec
    ElseIf mBloqueSeccion(mNumBloques) <> iSec Or mBloqueNumCtas(mNumBloques) >= D52_COLUMNAS Then
        NuevoBloque iSec
    End If
    iColumna = mBloqueNumCtas(mNumBloques)
    mBloqueCtas(iColumna, mNumBloques) = sCuenta
    mBloqueNumCtas(mNumBloques) = iColumna + 1
    D52_UbicarCuenta = mNumBloques
End Function

Private Sub NuevoBloque(ByVal iSec As Integer)
    mNumBloques = mNumBloques + 1
    If mNumBloques = 1 Then
        ReDim mBloqueSeccion(1 To 1)
        ReDim mBloqueNumCtas(1 To 1)
        ReDim mBloqueCtas(0 To D52_COLUMNAS - 1, 1 To 1)
    Else
        ReDim Preserve mBloqueSeccion(1 To mNumBloques)
        ReDim Preserve mBloqueNumCtas(1 To mNumBloques)
        ReDim Preserve mBloqueCtas(0 To D52_COLUMNAS - 1, 1 To mNumBloques)
    End If
    mBloqueSeccion(mNumBloques) = iSec
    mBloqueNumCtas(mNumBloques) = 0
End Sub

Public Function D52_NumBloques() As Long
    D52_NumBloques = mNumBloques
End Function

Public Function D52_BloqueSeccion(ByVal lCab As Long) As Integer
    D52_BloqueSeccion = mBloqueSeccion(lCab)
End Function

Public Function D52_BloqueNumCuentas(ByVal lCab As Long) As Integer
    D52_BloqueNumCuentas = mBloqueNumCtas(lCab)
End Function

Public Function D52_BloqueCuenta(ByVal lCab As Long, ByVal iColumna As Integer) As String
    D52_BloqueCuenta = mBloqueCtas(iColumna, lCab)
End Function

' Primer bloque de la seccion (para rotular "continuacion")
Public Function D52_EsContinuacion(ByVal lCab As Long) As Boolean
    If lCab > 1 Then D52_EsContinuacion = (mBloqueSeccion(lCab - 1) = mBloqueSeccion(lCab))
End Function

' =============================================================================
' Detalle normalizado
' =============================================================================
Public Sub D52_AgregarDetalle(ByVal lCab As Long, ByVal iColumna As Integer, ByVal sCuenta As String, _
                              ByVal vFecDoc As Variant, ByVal vFecOp As Variant, ByVal sSI As String, _
                              ByVal sCAR As String, ByVal sGlosa As String, ByVal cMonto As Currency)
    Dim lCap As Long
    If lCab < 1 Or lCab > mNumBloques Or iColumna < 0 Or iColumna >= D52_COLUMNAS Then
        Err.Raise D52_ERR_ESTADO, "D52_AgregarDetalle", "Cabecera/columna fuera de rango."
    End If
    mNumDet = mNumDet + 1
    If mNumDet = 1 Then
        lCap = 256
        ReDim mDetCab(1 To lCap): ReDim mDetCol(1 To lCap): ReDim mDetCta(1 To lCap)
        ReDim mDetFecDoc(1 To lCap): ReDim mDetFecOp(1 To lCap): ReDim mDetSI(1 To lCap)
        ReDim mDetCAR(1 To lCap): ReDim mDetGlosa(1 To lCap): ReDim mDetMonto(1 To lCap)
    ElseIf mNumDet > UBound(mDetCab) Then
        lCap = UBound(mDetCab) * 2
        ReDim Preserve mDetCab(1 To lCap): ReDim Preserve mDetCol(1 To lCap): ReDim Preserve mDetCta(1 To lCap)
        ReDim Preserve mDetFecDoc(1 To lCap): ReDim Preserve mDetFecOp(1 To lCap): ReDim Preserve mDetSI(1 To lCap)
        ReDim Preserve mDetCAR(1 To lCap): ReDim Preserve mDetGlosa(1 To lCap): ReDim Preserve mDetMonto(1 To lCap)
    End If
    mDetCab(mNumDet) = lCab
    mDetCol(mNumDet) = iColumna
    mDetCta(mNumDet) = sCuenta
    mDetFecDoc(mNumDet) = vFecDoc
    mDetFecOp(mNumDet) = vFecOp
    mDetSI(mNumDet) = sSI
    mDetCAR(mNumDet) = sCAR
    mDetGlosa(mNumDet) = sGlosa
    mDetMonto(mNumDet) = cMonto
End Sub

Public Function D52_NumDetalles() As Long
    D52_NumDetalles = mNumDet
End Function

' Total con signo (Debe - Haber) de una cuenta/columna dentro de un bloque.
Public Function D52_TotalColumna(ByVal lCab As Long, ByVal iColumna As Integer) As Currency
    Dim l As Long, c As Currency
    For l = 1 To mNumDet
        If mDetCab(l) = lCab And mDetCol(l) = iColumna Then c = c + mDetMonto(l)
    Next
    D52_TotalColumna = c
End Function

' Total con signo de todo el detalle (control de cuadre).
Public Function D52_TotalDetalle() As Currency
    Dim l As Long, c As Currency
    For l = 1 To mNumDet
        c = c + mDetMonto(l)
    Next
    D52_TotalDetalle = c
End Function

' =============================================================================
' Filas combinadas (misma regla que la tabla de Crystal cfFile7):
' se agrupan lineas del mismo bloque con igual SI, CAR, fechas, glosa y signo
' (debe > 0 / haber <= 0). Orden: Cabecera, primera cuenta (columna) de la fila,
' SI (saldo inicial primero), Fecha_Documento, CAR, orden de llegada.
' =============================================================================
Private Function ClaveFecha(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        ClaveFecha = "N"
    ElseIf VarType(v) = vbDate Then
        ClaveFecha = "D" & CStr(CDbl(v))
    Else
        ClaveFecha = "T" & CStr(v)
    End If
End Function

Public Sub D52_ConstruirFilas()
    Dim colClaves As Collection
    Dim l As Long, f As Long, sClave As String

    mNumFil = 0
    Erase mFilCab, mFilMinCol, mFilSI, mFilFecDoc, mFilCAR, mFilGlosa, mFilSeq, mFilMontos, mFilOrden
    If mNumDet = 0 Then Exit Sub

    ReDim mFilCab(1 To mNumDet): ReDim mFilMinCol(1 To mNumDet): ReDim mFilSI(1 To mNumDet)
    ReDim mFilFecDoc(1 To mNumDet): ReDim mFilCAR(1 To mNumDet): ReDim mFilGlosa(1 To mNumDet)
    ReDim mFilSeq(1 To mNumDet): ReDim mFilMontos(0 To D52_COLUMNAS - 1, 1 To mNumDet)

    Set colClaves = New Collection
    For l = 1 To mNumDet
        sClave = CStr(mDetCab(l)) & Chr$(1) & mDetSI(l) & Chr$(1) & mDetCAR(l) & Chr$(1) & _
                 ClaveFecha(mDetFecDoc(l)) & Chr$(1) & ClaveFecha(mDetFecOp(l)) & Chr$(1) & _
                 mDetGlosa(l) & Chr$(1) & IIf(mDetMonto(l) > 0, "D", "H")
        f = BuscarClave(colClaves, sClave)
        If f = 0 Then
            mNumFil = mNumFil + 1
            f = mNumFil
            colClaves.Add f, sClave
            mFilCab(f) = mDetCab(l)
            mFilMinCol(f) = mDetCol(l)
            mFilSI(f) = mDetSI(l)
            mFilFecDoc(f) = mDetFecDoc(l)
            mFilCAR(f) = mDetCAR(l)
            mFilGlosa(f) = mDetGlosa(l)
            mFilSeq(f) = l
        ElseIf mDetCol(l) < mFilMinCol(f) Then
            mFilMinCol(f) = mDetCol(l)
        End If
        mFilMontos(mDetCol(l), f) = mFilMontos(mDetCol(l), f) + mDetMonto(l)
    Next
    Set colClaves = Nothing

    ReDim mFilOrden(1 To mNumFil)
    For f = 1 To mNumFil
        mFilOrden(f) = f
    Next
    OrdenarFilas 1, mNumFil
End Sub

Private Function BuscarClave(colClaves As Collection, ByVal sClave As String) As Long
    On Error GoTo NoExiste
    BuscarClave = colClaves.Item(sClave)
    Exit Function
NoExiste:
    BuscarClave = 0
End Function

Private Function ValorFecha(ByVal v As Variant) As Double
    If VarType(v) = vbDate Then
        ValorFecha = CDbl(v)
    Else
        ValorFecha = -1E+30   ' Null/sin fecha: primero
    End If
End Function

' -1 si a va antes que b, 1 si despues, 0 si igual
Private Function CompararFilas(ByVal a As Long, ByVal b As Long) As Integer
    Dim dA As Double, dB As Double
    If mFilCab(a) <> mFilCab(b) Then CompararFilas = IIf(mFilCab(a) < mFilCab(b), -1, 1): Exit Function
    If mFilMinCol(a) <> mFilMinCol(b) Then CompararFilas = IIf(mFilMinCol(a) < mFilMinCol(b), -1, 1): Exit Function
    If mFilSI(a) <> mFilSI(b) Then CompararFilas = IIf(mFilSI(a) < mFilSI(b), -1, 1): Exit Function
    dA = ValorFecha(mFilFecDoc(a)): dB = ValorFecha(mFilFecDoc(b))
    If dA <> dB Then CompararFilas = IIf(dA < dB, -1, 1): Exit Function
    If mFilCAR(a) <> mFilCAR(b) Then CompararFilas = IIf(mFilCAR(a) < mFilCAR(b), -1, 1): Exit Function
    If mFilSeq(a) <> mFilSeq(b) Then CompararFilas = IIf(mFilSeq(a) < mFilSeq(b), -1, 1): Exit Function
    CompararFilas = 0
End Function

' Ordenamiento por mezcla (estable, O(n log n)) sobre mFilOrden
Private Sub OrdenarFilas(ByVal lIni As Long, ByVal lFin As Long)
    Dim aTmp() As Long
    If lFin <= lIni Then Exit Sub
    ReDim aTmp(lIni To lFin)
    MezclarRango lIni, lFin, aTmp
End Sub

Private Sub MezclarRango(ByVal lIni As Long, ByVal lFin As Long, aTmp() As Long)
    Dim lMed As Long, i As Long, j As Long, k As Long
    If lFin <= lIni Then Exit Sub
    lMed = (lIni + lFin) \ 2
    MezclarRango lIni, lMed, aTmp
    MezclarRango lMed + 1, lFin, aTmp
    i = lIni: j = lMed + 1: k = lIni
    Do While i <= lMed And j <= lFin
        If CompararFilas(mFilOrden(j), mFilOrden(i)) < 0 Then
            aTmp(k) = mFilOrden(j): j = j + 1
        Else
            aTmp(k) = mFilOrden(i): i = i + 1
        End If
        k = k + 1
    Loop
    Do While i <= lMed
        aTmp(k) = mFilOrden(i): i = i + 1: k = k + 1
    Loop
    Do While j <= lFin
        aTmp(k) = mFilOrden(j): j = j + 1: k = k + 1
    Loop
    For k = lIni To lFin
        mFilOrden(k) = aTmp(k)
    Next
End Sub

Public Function D52_NumFilas() As Long
    D52_NumFilas = mNumFil
End Function

' Indice de la fila combinada en la posicion k (1..D52_NumFilas) ya ordenada
Public Function D52_FilaOrdenada(ByVal k As Long) As Long
    D52_FilaOrdenada = mFilOrden(k)
End Function

Public Function D52_FilaCabecera(ByVal f As Long) As Long
    D52_FilaCabecera = mFilCab(f)
End Function

Public Function D52_FilaFecha(ByVal f As Long) As Variant
    D52_FilaFecha = mFilFecDoc(f)
End Function

Public Function D52_FilaCAR(ByVal f As Long) As String
    D52_FilaCAR = mFilCAR(f)
End Function

Public Function D52_FilaGlosa(ByVal f As Long) As String
    D52_FilaGlosa = mFilGlosa(f)
End Function

Public Function D52_FilaSI(ByVal f As Long) As String
    D52_FilaSI = mFilSI(f)
End Function

Public Function D52_FilaMonto(ByVal f As Long, ByVal iColumna As Integer) As Currency
    D52_FilaMonto = mFilMontos(iColumna, f)
End Function
