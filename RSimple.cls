VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'variables locales para almacenar los valores de las propiedades
Private RSimple1 As frmRSimple
Private mvarNewHDC As Variant 'copia local
Private mvarFormato As Variant 'copia local
Private mvartcAnoPro As String 'copia local
Private mvartcWhere As String 'copia local
Private mvartcEmpPro As String 'copia local
Private mvartcMesPro As String 'copia local
Private mvartcCia As String 'copia local
Private mvartiLongCta As Integer 'copia local
Private mvartiReporte As Integer 'copia local
Private mvartcFechSys As String 'copia local
Private mvartdb As Variant 'copia local
Private mvartdbEmpresa As Variant 'copia local
Private mvartdbWork As Variant 'copia local
Private mvartcPath As String 'copia local
Private mvartpsUserName As String 'copia local
Private mvartcRUC As String 'copia local
Private mvartpsUserPasswordDBase As String 'copia local
Private mvarLocalizacionExe As String 'copia local
Private mvarCaption As String 'copia local
Private mvarFileCrystal As String 'copia local
Private mvarSQLCrystal As String 'copia local
Private mvarSortCrystal As String 'copia local

Public Property Let LocalizacionExe(ByVal vData As String)
    mvarLocalizacionExe = vData
End Property

Public Property Let Caption(ByVal vData As String)
    mvarCaption = vData
End Property

Public Property Let FileCrystal(ByVal vData As String)
    mvarFileCrystal = vData
End Property

Public Property Let SqlCrystal(ByVal vData As String)
    mvarSQLCrystal = vData
End Property

Public Property Let SortCrystal(ByVal vData As String)
    mvarSortCrystal = vData
End Property

Public Property Let tpsUserPasswordDBase(ByVal vData As String)
    mvartpsUserPasswordDBase = vData
End Property

Public Property Let tcPath(ByVal vData As String)
    mvartcPath = vData
End Property

Public Property Let tdb(ByVal vData As DAO.Database)
    Set mvartdb = vData
End Property

Public Property Set tdb(ByVal vData As DAO.Database)
    Set mvartdb = vData
End Property

Public Property Get tdb() As DAO.Database
    Set tdb = mvartdb
End Property

Public Property Let tdbWork(ByVal vData As DAO.Database)
    Set mvartdbWork = vData
End Property

Public Property Set tdbWork(ByVal vData As DAO.Database)
    Set mvartdbWork = vData
End Property

Public Property Get tdbWork() As DAO.Database
    Set tdbWork = mvartdbWork
End Property

Public Property Let tdbEmpresa(ByVal vData As DAO.Database)
    Set mvartdbEmpresa = vData
End Property

Public Property Set tdbEmpresa(ByVal vData As DAO.Database)
    Set mvartdbEmpresa = vData
End Property

Public Property Get tdbEmpresa() As DAO.Database
    Set tdbEmpresa = mvartdbEmpresa
End Property

Public Property Let tcFechSys(ByVal vData As String)
    mvartcFechSys = vData
End Property

Public Property Let tpsUserName(ByVal vData As String)
    mvartpsUserName = vData
End Property

Public Property Let tcRUC(ByVal vData As String)
    mvartcRUC = vData
End Property

Public Property Let tiLongCta(ByVal vData As Integer)
    mvartiLongCta = vData
End Property

Public Property Let tiReporte(ByVal vData As Integer)
    mvartiReporte = vData
End Property

Public Property Let tcCia(ByVal vData As String)
    mvartcCia = vData
End Property

Public Property Let tcEmpPro(ByVal vData As String)
    mvartcEmpPro = vData
End Property

Public Property Let tcMesPro(ByVal vData As String)
    mvartcMesPro = vData
End Property

Public Property Let tcAnoPro(ByVal vData As String)
    mvartcAnoPro = vData
End Property

Public Property Let tcWhere(ByVal vData As String)
    mvartcWhere = vData
End Property

Public Property Let NewHDC(ByVal vData As Variant)
    Set mvarNewHDC = vData
    Set mvarFormMain = vData
End Property

Public Property Set NewHDC(ByVal vData As Variant)
    Set mvarNewHDC = vData
    Set mvarFormMain = vData
End Property

Public Property Let frmInterfase(ByVal vData As Variant)
    Set mvarFormato = vData
End Property

Public Property Set frmInterfase(ByVal vData As Variant)
    Set mvarFormato = vData
End Property

Public Property Get frmInterfase() As Variant
    Set frmInterfase = mvarFormato
End Property

Public Function RSimple_Execute()
    Set db = mvartdb
    Set dbEmpresa = mvartdbEmpresa
    Set dbWorkArea = mvartdbWork
    Set Formato = frmInterfase
    cAnoPro = mvartcAnoPro
    cEmpPro = mvartcEmpPro
    cMesPro = mvartcMesPro
    cCia = mvartcCia
    psUserName = mvartpsUserName
    cRUC = mvartcRUC
    psUserPasswordDBase = mvartpsUserPasswordDBase
    cFechSys = mvartcFechSys
    cPath = mvartcPath
    cWhere = mvartcWhere
    cLocalizacionExe = mvarLocalizacionExe
    iLongCta = mvartiLongCta
    Load RSimple1
    With RSimple1
        .iReporte = mvartiReporte
        Select Case mvartiReporte
            Case kFormatos_5_2
                .Caption = "Diario Simplificado Formato 5.2"
            Case kFormatos_5_3
                .Caption = "Movimiento Contable"
            Case kFormatos_3_18
                .Caption = "Flujo de Efectivo Formato 3.18"
                .opOrientacion(0).Value = True
            Case kVoucherIngresados, kCheckVouchers, kCheckVouchersContinuos
                .Caption = "Reporte de Vouchers"
                .opOrientacion(0).Value = True
            Case kAnalisisdeAuxiliares
                .Caption = "Análisis de Auxiliares"
                .opOrientacion(0).Value = True
            Case Else
                .Caption = mvarCaption
                .FileCrystal = mvarFileCrystal
                .SqlCrystal = mvarSQLCrystal
                .SortCrystal = mvarSortCrystal
                .opOrientacion(0).Value = True
        End Select
        .Show vbModal, mvarNewHDC
    End With
End Function

Private Sub Class_Initialize()
    Set RSimple1 = New frmRSimple
End Sub

Private Sub Class_Terminate()
    Set mvarFormato = Nothing
    Set mvartdb = Nothing
    Set mvartdbEmpresa = Nothing
    Set mvartdbWork = Nothing
    Set RSimple1 = Nothing
End Sub


