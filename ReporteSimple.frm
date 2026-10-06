VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{604A59D5-2409-101D-97D5-46626B63EF2D}#1.0#0"; "TDBNumbr.ocx"
Begin VB.Form frmRSimple 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte Simple"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   ControlBox      =   0   'False
   HelpContextID   =   15
   Icon            =   "ReporteSimple.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmImpresora 
      Caption         =   "Impresora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   5535
      Begin VB.CheckBox chkSansSerif 
         Caption         =   "Imprimir en &hojas continuas"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ComboBox cmbImpresoras 
         Height          =   315
         Left            =   1040
         Style           =   2  'Dropdown List
         TabIndex        =   0
         TabStop         =   0   'False
         Tag             =   "&Nombre"
         Top             =   280
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "I&mpresora:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   0
         Left            =   1020
         TabIndex        =   23
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Estado:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   1
         Left            =   1020
         TabIndex        =   21
         Top             =   1050
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Ubicación:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   1050
         Width           =   855
      End
   End
   Begin VB.Frame frmOrientacion 
      Caption         =   "Orientación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   18
      Top             =   840
      Width           =   2655
      Begin VB.OptionButton opOrientacion 
         Caption         =   "&Horizontal"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton opOrientacion 
         Caption         =   "&Vertical"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frmCopias 
      Caption         =   "Copias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   16
      Top             =   2040
      Width           =   4095
      Begin VB.CheckBox chkIntercalar 
         Caption         =   "&Empaste"
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   270
         Width           =   975
      End
      Begin TDBNumberCtrl.TDBNumber txtCopias 
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   556
         _Version        =   65537
         AlignHorizontal =   1
         ClipMode        =   0
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         HighlightText   =   0   'False
         ZeroAllowed     =   -1  'True
         MinusColor      =   255
         MaxValue        =   99999
         MinValue        =   0
         Value           =   0
         SelStart        =   1
         SelLength       =   0
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         KeyThreeZero    =   ""
         SepDecimal      =   "."
         SepThousand     =   ","
         Text            =   "0"
         Format          =   "####0"
         DisplayFormat   =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "Número de Copia&s"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   96
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ReporteSimple.frx":000C
         MousePointer    =   0
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   2200
         TabIndex        =   17
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtCopias"
         BuddyDispid     =   196627
         OrigLeft        =   2400
         OrigTop         =   360
         OrigRight       =   2595
         OrigBottom      =   675
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   28
         Enabled         =   -1  'True
      End
   End
   Begin VB.Frame frmIntervalo 
      Caption         =   "Intervalo de Páginas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   15
      Top             =   2040
      Width           =   4095
      Begin VB.OptionButton opTodo 
         Caption         =   "&Todas"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton opIntervalo 
         Caption         =   "Pá&gina desde"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   630
         Width           =   1335
      End
      Begin VB.CheckBox chkArchivo 
         Alignment       =   1  'Right Justify
         Caption         =   "&Imprimir en Archivo"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin TDBNumberCtrl.TDBNumber txtHasta 
         Height          =   315
         Left            =   2640
         TabIndex        =   9
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   65537
         AlignHorizontal =   1
         ClipMode        =   0
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         HighlightText   =   -1  'True
         ZeroAllowed     =   -1  'True
         MinusColor      =   255
         MaxValue        =   99999
         MinValue        =   0
         Value           =   0
         SelStart        =   1
         SelLength       =   0
         KeyClear        =   "{F2}"
         KeyNext         =   "{ENTER}"
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         KeyThreeZero    =   ""
         SepDecimal      =   "."
         SepThousand     =   ","
         Text            =   "0"
         Format          =   "####0"
         DisplayFormat   =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "&hasta"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   32
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ReporteSimple.frx":0028
         MousePointer    =   0
      End
      Begin TDBNumberCtrl.TDBNumber txtDesde 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   65537
         AlignHorizontal =   1
         ClipMode        =   0
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         HighlightText   =   -1  'True
         ZeroAllowed     =   -1  'True
         MinusColor      =   255
         MaxValue        =   99999
         MinValue        =   0
         Value           =   0
         SelStart        =   1
         SelLength       =   0
         KeyClear        =   "{F2}"
         KeyNext         =   "{ENTER}"
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         KeyThreeZero    =   ""
         SepDecimal      =   "."
         SepThousand     =   ","
         Text            =   "0"
         Format          =   "####0"
         DisplayFormat   =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   ""
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   0
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ReporteSimple.frx":0044
         MousePointer    =   0
      End
   End
   Begin VB.CommandButton cmdPreliminar 
      Caption         =   "&Preliminar"
      Height          =   400
      Left            =   6720
      TabIndex        =   12
      Top             =   1080
      Width           =   1200
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "E&xcel"
      Height          =   400
      Left            =   6720
      TabIndex        =   27
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   400
      Left            =   7440
      TabIndex        =   3
      Top             =   2970
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   400
      Left            =   5760
      TabIndex        =   2
      Top             =   2970
      Width           =   1200
   End
   Begin VB.Frame frmResumen 
      Height          =   615
      Left            =   6000
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
      Begin VB.OptionButton opResumen 
         Caption         =   "&Analítico"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   260
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton opResumen 
         Caption         =   "&Resumido"
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   25
         Top             =   260
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmRSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cfFile As String
Private cfFile1 As String
Private cfFile2 As String
Private cfFile3 As String
Private cfFile4 As String
Private cfFile5 As String
Private cfFile6 As String
Private cfFile7 As String
Private cfFile8 As String
Private cfFile9 As String
Private cfFile10 As String
Private cfFile11 As String
Private cfFile14 As String
Private cfFile15 As String
Public iReporte As Integer
Public FileCrystal As String
Public SqlCrystal As String
Public SortCrystal As String
Private lRegistros As Long
Private rs4 As DAO.Recordset
Private oViewSaldos As VerSaldosCTB
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const kMAX_REINTENTOS_BLOQUEO As Integer = 5
Private Const kCTA_CONTRA_COSTOS As String = "7911101     "
Private Const kTITULO_FORMATO_5_2 As String = "FORMATO 5.2: LIBRO DIARIO DE FORMATO SIMPLIFICADO"

Private Sub Form_Activate()
    Select Case iReporte
        Case kFormatos_3_18
            opResumen(0).Caption = "&Del Mes"
            opResumen(1).Caption = "Ac&umulado"
            cmdPreliminar.Top = 1200
            frmResumen.Visible = True
            
        Case kFormatos_5_2
            frmResumen.Visible = True
            frmResumen.Caption = "Compras/Ventas"
            cmdExcel.Visible = True
            opResumen(0).Value = True
            opResumen(1).Value = False
            
        Case Else
            cmdPreliminar.Top = frmImpresora.Top + 200
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass

    Set oViewSaldos = New VerSaldosCTB
    oViewSaldos.tcMesPro = cMesPro
'    LoadStatus 1, 3
    Call CenterMDI(Me, mvarFormMain)
    Call Archivos
    Set oImprimir = New clsViewReport
    Set oArchivo = New Archivodb
    Set oString = New MString
    cfFile = oArchivo.FileTemp(dbWorkArea)
    cfFile1 = oArchivo.FileTemp(dbWorkArea)
    cfFile2 = oArchivo.FileTemp(dbWorkArea)
    cfFile3 = oArchivo.FileTemp(dbWorkArea)
    cfFile4 = oArchivo.FileTemp(dbWorkArea)
    cfFile5 = oArchivo.FileTemp(dbWorkArea)
    cfFile6 = oArchivo.FileTemp(dbWorkArea)
    cfFile7 = oArchivo.FileTemp(dbWorkArea)
    cfFile8 = oArchivo.FileTemp(dbWorkArea)
    cfFile9 = oArchivo.FileTemp(dbWorkArea)
    cfFile10 = oArchivo.FileTemp(dbWorkArea)
    cfFile11 = oArchivo.FileTemp(dbWorkArea)
    cfFile14 = oArchivo.FileTemp(dbWorkArea)
    cfFile15 = oArchivo.FileTemp(dbWorkArea)
    With oImprimir
        .CBImpresoras = cmbImpresoras
        .cmdAceptar = cmdAceptar
        .cmdCancelar = cmdCancelar
        .cmdPreliminar = cmdPreliminar
        .opIntervalo = opIntervalo
        .opTodo = opTodo
        .txtDesde = txtDesde
        .txtHasta = txtHasta
        .LocalizacionExe = cLocalizacionExe
        .LocalizacionData = cPath
        .tcEmpPro = cEmpPro
        .tcMesPro = cMesPro
        .tcAnoPro = cAnoPro
        .tcPath = cPath
        .sFilePrefix = "Dmm"
        .Formato = Me
        .tpsUserName = psUserName
        .tpsUserPasswordDBase = psUserPassword
        .NewHDC = mvarFormMain
        .ClearPrint
    End With
    Screen.MousePointer = vbArrow
'    LoadStatus 1, 1
End Sub

Public Function PrepararReporte() As Boolean
    PrepararReporte = True
    With oImprimir
        .sFilePrefix = "Dmm"
        .Orientacion = 0
        Select Case iReporte
            Case kFormatos_5_2
                .Orientacion = 1
                .crArchivo = "FORMATO5_2.rpt"
                .sSql = "SELECT Ctb.Cuenta, Ctb.Detalle, Ctb.TipoNumero, Ctb.DebeMN, Ctb.HaberMN, Ctb.Fecha_Operacion, Ctb.Nombre From " & cfFile14 & " Ctb"
                If Not PrepararDiarioSimplificadoSeguro() Then
                    PrepararReporte = False
                    Exit Function
                End If
                .LoadPrint
                Call PreparaDiarioSimplificadoAnalitico
                .Files "Ctb", cfFile7
                .Files "Ctb1", cfFile11
                
            Case kFormatos_5_3
                .sFilePrefix = kPrefijoFileContabilidad
                .Orientacion = 0
                .crArchivo = "FORMATO5_3.rpt"
                .sSql = "SELECT Ctb.Cuenta, Ctb.Detalle, Ctb.TipoNumero, Ctb.DebeMN, Ctb.HaberMN, Ctb.Fecha_Operacion, Ctb.Nombre From " & cfFile14 & " Ctb"
                .LoadPrint
                Call .Parameters1("CIA", cCia, PE_PF_STRING, 0)
                Call .Parameters1("RUC", cRUC, PE_PF_STRING, 1)
                .Files "Pl", cfPla
                .Files "Do", cfDoc
                .Files "Au", cfAux
                .Files "Tr", cfTra
                
            Case kAnalisisdeAuxiliares
                If Formato.cmbplan.Text = "" Then
                    PrepararReporte = False
                    Exit Function
                End If
                .sFilePrefix = kPrefijoFileContabilidad
                .crArchivo = "ConsultaSaldosdeAuxiliar.rpt"
                .sSql = "SELECT Sa.Cuenta, Sa.Auxiliar, Sa.[Apertura_Debe], Sa.[Apertura_Haber], Sa.[Enero_Debe], Sa.[Enero_Haber], Sa.[Febrero_Debe], Sa.[Febrero_Haber], Sa.[Marzo_Debe], Sa.[Marzo_Haber], Sa.[Abril_Debe], Sa.[Abril_Haber], Sa.[Mayo_Debe], Sa.[Mayo_Haber], Sa.[Junio_Debe], Sa.[Junio_Haber], Sa.[Julio_Debe], Sa.[Julio_Haber], Sa.[Agosto_Debe], Sa.[Agosto_Haber], Sa.[Setiembre_Debe], Sa.[Setiembre_Haber], Sa.[Octubre_Debe], Sa.[Octubre_Haber], Sa.[Noviembre_Debe], Sa.[Noviembre_Haber], Sa.[Diciembre_Debe], Sa.[Diciembre_Haber], Sa.[Cierre_Debe], Sa.[Cierre_Haber], Pl.Nombre, Au.Nombre " + _
                        "FROM (" & IIf(Formato.Check1 = 0, cfSa1, cfSa2) & " Sa INNER JOIN " & cfPla & " Pl ON Sa.Cuenta = Pl.Cuenta) " & _
                        "INNER JOIN " & cfAux & " Au ON Sa.Auxiliar = Au.Auxiliar"
                .sWhere = "{Sa.Cuenta} = '" & Formato.cmbplan.Text & "' And {Sa.Tipo} = '02'"
                .sSort = "Sa.Auxiliar"
                .LoadPrint
                Call PreparaAnalisisdeAuxiliares
                .Files "Sa", IIf(Formato.Check1 = 0, cfSa1, cfSa2)
                .Files "Pl", cfPla
                .Files "Au", cfAux
        
            Case kVoucherIngresados
                .sFilePrefix = kPrefijoFileContabilidad
                .crArchivo = "FileTransacciones.rpt"
                .sSql = "SELECT Tr.Cuenta, Tr.Auxiliar, Tr.Tipo_Documento, Tr.Numero_Documento, Tr.Fecha_Documento, Tr.Detalle, Tr.Debe, Tr.Haber, Tr.Moneda, Tr.Tipo_de_Cambio, Tr.Tipo, Tr.Numero, Tr.TipoNumero, Tr.Grupo, " + _
                        "Vo.Nombre, Pl.Nombre, Au.Nombre " + _
                        "From ((" & cfTra & " Tr INNER JOIN " & cfPla & " Pl ON Tr.Cuenta = Pl.Cuenta) LEFT JOIN " & cfAux & " Au ON Tr.Auxiliar = Au.Auxiliar) LEFT JOIN " & cfVou & " Vo ON Tr.Tipo = Vo.Codigo"
                .LoadPrint
                If Not PreparaVoucherIngresados() Then
                    MsgBox "No se pueden seleccionar mas de 10 Voucher a Imprimir", vbCritical, "Calcum - Error"
                    PrepararReporte = False
                    Exit Function
                End If
                .Files "Tr", cfTra
                .Files "Pl", cfPla
                .Files "Au", cfAux
                .Files "Vo", cfVou
           
            Case kCheckVouchers
                .sFilePrefix = kPrefijoFileContabilidad
                .crArchivo = "CheckVouchers.rpt"
                .sSql = "SELECT Tr.Cuenta, Tr.Auxiliar, Tr.[Centro_de_Costo], Tr.[Tipo_Documento], Tr.Abreviatura, Tr.[Numero_Documento], " + _
                        "Tr.[Fecha_Documento], Tr.Vencimiento, Tr.Debe, Tr.Haber, Tr.[Tipo_de_Cambio], Tr.Moneda, Vo.Nombre " + _
                        "FROM " & cfTra & " Tr INNER JOIN " & cfVou & " Vo ON Tr.Tipo = Vo.Codigo"
                .sSort = ""
                .sWhere = ""
                .LoadPrint
                Call PreparaCheckVouchers
                .Files "Tr", cfTra
                .Files "Vo", cfVou
                
            Case kCheckVouchersContinuos
                .sFilePrefix = kPrefijoFileContabilidad
                .crArchivo = "CheckVouchers1.rpt"
                .sSql = "SELECT Tr.Cuenta, Tr.Auxiliar, Tr.[Centro_de_Costo], Tr.[Tipo_Documento], Tr.Abreviatura, Tr.[Numero_Documento], " + _
                        "Tr.[Fecha_Documento], Tr.Vencimiento, Tr.Debe, Tr.Haber, Tr.[Tipo_de_Cambio], Tr.Moneda, Vo.Nombre " + _
                        "FROM " & cfTra & " Tr INNER JOIN " & cfVou & " Vo ON Tr.Tipo = Vo.Codigo"
                .sSort = ""
                .sWhere = ""
                .LoadPrint
                Call PreparaCheckVouchers1
                .Files "Tr", cfTra
                .Files "Vo", cfVou
                
            Case kFormatos_3_18
                .sFilePrefix = kPrefijoFileContabilidad
                .crArchivo = "FORMATO3_18.rpt"
                .sSql = "SELECT xCtb.Nombre, Fl.Codigo, Fl.MainGroup, " + _
                        "Fl.Grupo, Fl.SubGrupo, Fl.Nombre, Fl.Apertura, Fl.Enero, " + _
                        "Fl.Febrero, Fl.Marzo, Fl.Abril, Fl.Mayo, Fl.Junio, Fl.Julio, " + _
                        "Fl.Agosto, Fl.Setiembre, Fl.Octubre, Fl.Noviembre, Fl.Diciembre, Fl.Cierre, " + _
                        "Fl.AperturaME, Fl.EneroME, Fl.FebreroME, Fl.MarzoME, Fl.AbrilME, Fl.MayoME, Fl.JunioME, Fl.JulioME, Fl.AgostoME, Fl.SetiembreME, Fl.OctubreME, Fl.NoviembreME, Fl.DiciembreME, Fl.CierreME " + _
                        "FROM xCtbFlujo xCtb INNER JOIN " + cfFlu + " Fl ON xCtb.Codigo = Fl.Grupo"
                .sWhere = "{Fl.Prefijo} = '" & kQUERY_TIPO_FLUJO_EFECTIVO_DIRECTO & "'"
                .sSort = "Fl.MainGroup, Fl.Grupo, Fl.SubGrupo, Fl.Codigo"
                .LoadPrint
                Call PreparaFlujodeEfectivo
                .Files "xCtb", "xCtbFlujo"
                .Files "Fl", cfFlu
                
            Case kIngresos
                .sFilePrefix = kPrefijoFileContabilidad
                .crArchivo = FileCrystal
                .sSql = SqlCrystal
                .sSort = SortCrystal
                .LoadPrint
                Call PreparaCaseElse
                .Files "Pr", cfEle
                
            Case kEgresos
                .sFilePrefix = kPrefijoFileContabilidad
                .crArchivo = FileCrystal
                .sSql = SqlCrystal
                .sSort = SortCrystal
                .LoadPrint
                Call PreparaCaseElse
                .Files "Pr", cfPre
                
            Case kGrupos
                .sFilePrefix = kPrefijoFileContabilidad
                .crArchivo = FileCrystal
                .sSql = SqlCrystal
                .sSort = SortCrystal
                .sWhere = "{Gr.Prefijo} = '" & cWhere & "'"
                .LoadPrint
                Call PreparaCaseElse
                .Files "xCtb", "xCtbGrupos"
                .Files "Gr", cfGrp
                
            Case Else
                .sFilePrefix = kPrefijoFileContabilidad
                .crArchivo = FileCrystal
                .sSql = SqlCrystal
                .sSort = SortCrystal
                .LoadPrint
                Call PreparaCaseElse
                
        End Select
    End With
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call dbWorkArea.Execute("DROP TABLE " + cfFile + ";")
    Call dbWorkArea.Execute("DROP TABLE " + cfFile1 + ";")
    Call dbWorkArea.Execute("DROP TABLE " + cfFile2 + ";")
    Call dbWorkArea.Execute("DROP TABLE " + cfFile3 + ";")
    Call dbWorkArea.Execute("DROP TABLE " + cfFile4 + ";")
    Call dbWorkArea.Execute("DROP TABLE " + cfFile5 + ";")
    Call dbWorkArea.Execute("DROP TABLE " + cfFile6 + ";")
    Call dbWorkArea.Execute("DROP TABLE " + cfFile7 + ";")
    Call dbWorkArea.Execute("DROP TABLE " + cfFile8 + ";")
    Call dbWorkArea.Execute("DROP TABLE " + cfFile9 + ";")
    Call dbWorkArea.Execute("DROP TABLE " + cfFile10 + ";")
    Call dbWorkArea.Execute("DROP TABLE " + cfFile11 + ";")
    Call dbWorkArea.Execute("DROP TABLE " + cfFile14 + ";")
    Call dbWorkArea.Execute("DROP TABLE " + cfFile15 + ";")
    Set oViewSaldos = Nothing
    Set oImprimir = Nothing
    Set oArchivo = Nothing
    Set oString = Nothing
    Set db = Nothing
    Set dbEmpresa = Nothing
    Set dbWorkArea = Nothing
    Set Formato = Nothing
End Sub

Private Sub PreparaAnalisisdeAuxiliares()
    Dim objTrueCombo As TrueDBList50.TDBCombo
    Set objTrueCombo = Formato.cmbplan
    With oImprimir
        Call .Parameters1("MESPRO", Val(cMesPro), PE_PF_NUMBER, 0)
        Call .Parameters1("CIA", cCia, PE_PF_STRING, 1)
        Call .Parameters1("LONGCTA", iLongCta, PE_PF_NUMBER, 2)
        Call .Parameters1("MESANO", "ANALISIS POR AUXILIAR A " + UCase(oString.Periodo(CInt(Val(cMesPro) - 1))) & " " & cAnoPro, PE_PF_STRING, 3)
        Call .Parameters1("FECSYS", cFechSys, PE_PF_STRING, 4)
        Call .Parameters1("USER", psUserName, PE_PF_STRING, 5)
        Call .Parameters1("MONEDA", IIf(Formato.Check1 = 0, "M.Nacional", "M.Extranjera"), PE_PF_STRING, 6)
        Call .Parameters1("NOMBRE", objTrueCombo.Columns(0).Text & " " & objTrueCombo.Columns(1).Text, PE_PF_STRING, 7)
        Call .Parameters1("RUC", cRUC, PE_PF_STRING, 8)
    End With
    Set objTrueCombo = Nothing
End Sub

Private Function PreparaVoucherIngresados() As Boolean
    Dim objTrueDBGrid As TrueOleDBGrid60.TDBGrid
    Dim sImprimeString As String
    Dim iIndex As Integer
    PreparaVoucherIngresados = True
    
    Set objTrueDBGrid = Formato.TDBGrid1
    If objTrueDBGrid.SelBookmarks.Count > 0 Then
        If objTrueDBGrid.SelBookmarks.Count > 10 Then
            PreparaVoucherIngresados = False
            Exit Function
        End If
        For iIndex = 0 To objTrueDBGrid.SelBookmarks.Count - 1
            objTrueDBGrid.Bookmark = objTrueDBGrid.SelBookmarks(iIndex)
            sImprimeString = sImprimeString + "{Tr.TipoNumero} = '" & objTrueDBGrid.Columns("TipoNumero").Value & "' Or "
        Next
        oImprimir.sWhere = Mid(sImprimeString, 1, Len(sImprimeString) - 4)
       Else
        oImprimir.sWhere = "{Tr.TipoNumero} = '" & objTrueDBGrid.Columns("TipoNumero").Value & "'"
    End If
    With oImprimir
        Call .Parameters1("CIA", cCia, PE_PF_STRING, 0)
        Call .Parameters1("MESANO", UCase(oString.Periodo(CInt(Val(cMesPro) - 1))) & " DEL " & cAnoPro, PE_PF_STRING, 1)
        Call .Parameters1("FECSYS", cFechSys, PE_PF_STRING, 2)
        Call .Parameters1("USER", psUserName, PE_PF_STRING, 3)
        Call .Parameters1("RUC", cRUC, PE_PF_STRING, 4)
        .sSort = "Tr.Registro"
    End With
    Set objTrueDBGrid = Nothing
End Function

Private Sub PreparaCheckVouchers()
    Dim objTrueDBGrid As TrueOleDBGrid60.TDBGrid
    Dim adRs As DAO.Recordset
    Dim sConcepto As String, sAnombre As String, sBnombre As String
    Dim sNroCta As String, sFDocumento As String, sNDocumento As String
    Set objTrueDBGrid = Formato.TDBGrid1
    oImprimir.sWhere = "{Tr.TipoNumero} = '" & objTrueDBGrid.Columns("TipoNumero").Text & "'"
    oImprimir.sSort = "Registro"
    If oArchivo.MakeQuery(dbEmpresa, "Select Tr.Fecha_Documento, Tr.Numero_Documento, Tr.Detalle, Au.Nombre AS ANombre, Ba.Nombre AS BNombre, Ba.Numero_de_Cuenta FROM (" + cfTra + " Tr INNER JOIN " & cfAux & " Au ON Tr.Auxiliar = Au.Auxiliar) INNER JOIN " & cfBan & " Ba ON Tr.Cuenta = Ba.Cuenta WHERE Tr.Cuenta LIKE '10*' And Tr.TipoNumero = '" & objTrueDBGrid.Columns("TipoNumero") & "' And (Tr.Tipo_Documento >= '" & kDOCUMENTOS_TRANSFERENCIA_BANCARIA & "' And Tr.Tipo_Documento <= '" & kDOC_TARJETA_DE_CREDITO & "' Or Tr.Tipo_Documento = '" & KDOCUMENTOS_DOCEMITIDOBCOFINANSEGURO & "')", adRs) Then
        sConcepto = oString.ReadField(adRs!Detalle)
        sAnombre = oString.ReadField(adRs!ANombre)
        sBnombre = oString.ReadField(adRs!BNombre)
        sNroCta = oString.ReadField(adRs![Numero_de_Cuenta])
        sFDocumento = Format(adRs![Fecha_Documento], kFORMAT_TO_DISPLAY_DATE)
        sNDocumento = oString.ReadField(adRs![Numero_Documento])
    End If
    adRs.Close
    Set adRs = Nothing
    With oImprimir
        Call .Parameters1("CIA", cCia, PE_PF_STRING, 0)
        Call .Parameters1("MESANO", UCase(oString.Periodo(CInt(Val(cMesPro) - 1))) & " DEL " & cAnoPro, PE_PF_STRING, 1)
        Call .Parameters1("FECSYS", cFechSys, PE_PF_STRING, 2)
        Call .Parameters1("USER", psUserName, PE_PF_STRING, 3)
        Call .Parameters1("ANOMBRE", sAnombre, PE_PF_STRING, 4)
        Call .Parameters1("CONCEPTO", sConcepto, PE_PF_STRING, 5)
        Call .Parameters1("BNOMBRE", sBnombre, PE_PF_STRING, 6)
        Call .Parameters1("NROCTA", sNroCta, PE_PF_STRING, 7)
        Call .Parameters1("FDOCUMENTO", sFDocumento, PE_PF_STRING, 8)
        Call .Parameters1("NDOCUMENTO", sNDocumento, PE_PF_STRING, 9)
        Call .Parameters1("RUC", cRUC, PE_PF_STRING, 10)
    End With
    Set objTrueDBGrid = Nothing
End Sub

Private Sub PreparaCheckVouchers1()
    Dim objTrueDBGrid As TrueOleDBGrid60.TDBGrid
    Dim adRs As DAO.Recordset
    Dim sConcepto As String, sAnombre As String, cMonto As Currency, sFDocumento As String
    Set objTrueDBGrid = Formato.TDBGrid1
    oImprimir.sWhere = "{Tr.TipoNumero} = '" & objTrueDBGrid.Columns("TipoNumero").Text & "'"
    oImprimir.sSort = "Registro"
    If oArchivo.MakeQuery(dbEmpresa, "Select Tr.Fecha_Documento, Tr.Numero_Documento, Tr.Detalle, Tr.Haber, Au.Nombre AS ANombre, Ba.Nombre AS BNombre, Ba.Numero_de_Cuenta FROM (" + cfTra + " Tr INNER JOIN " & cfAux & " Au ON Tr.Auxiliar = Au.Auxiliar) INNER JOIN " & cfBan & " Ba ON Tr.Cuenta = Ba.Cuenta WHERE Tr.Cuenta LIKE '10*' And Tr.TipoNumero = '" & objTrueDBGrid.Columns("TipoNumero") & "' And (Tr.Tipo_Documento >= '" & kDOCUMENTOS_TRANSFERENCIA_BANCARIA & "' And Tr.Tipo_Documento <= '" & kDOC_TARJETA_DE_CREDITO & "' Or Tr.Tipo_Documento = '" & KDOCUMENTOS_DOCEMITIDOBCOFINANSEGURO & "')", adRs) Then
        sConcepto = oString.ReadField(adRs!Detalle)
        sAnombre = oString.ReadField(adRs!ANombre)
        sFDocumento = Format(adRs![Fecha_Documento], kFORMAT_TO_SAVE_DATE)
        cMonto = oString.ReadField(adRs!Haber)
    End If
    adRs.Close
    Set adRs = Nothing
    With oImprimir
        Call .Parameters1("ANOMBRE", sAnombre, PE_PF_STRING, 0)
        Call .Parameters1("CONCEPTO", sConcepto, PE_PF_STRING, 1)
        Call .Parameters1("FDOCUMENTO", Mid(sFDocumento, 1, 2) + "      " + Mid(sFDocumento, 4, 2) + "      " + Mid(sFDocumento, 7), PE_PF_STRING, 2)
        Call .Parameters1("MONTO", cMonto, PE_PF_CURRENCY, 3)
        Call .Parameters1("FECHA", sFDocumento, PE_PF_STRING, 4)
    End With
    Set objTrueDBGrid = Nothing
End Sub

Private Sub PreparaCaseElse()
    With oImprimir
        Call .Parameters1("CIA", cCia, PE_PF_STRING, 0)
        Call .Parameters1("USER", psUserName, PE_PF_STRING, 1)
        Call .Parameters1("RUC", cRUC, PE_PF_STRING, 2)
        Call .Parameters1("MESANO", Me.Caption, PE_PF_STRING, 3)
    End With
End Sub

Private Sub PreparaFlujodeEfectivo()
    With oImprimir
        Call .Parameters1("CIA", cCia, PE_PF_STRING, 0)
        Call .Parameters1("MESPRO", Val(cMesPro), PE_PF_NUMBER, 1)
        Call .Parameters1("MESANO", IIf(opResumen(1).Value, "ACUMULADO A", "MES DE") & " " & UCase(oString.Periodo(CInt(Val(cMesPro) - 1))) & " " & cAnoPro, PE_PF_STRING, 2)
        Call .Parameters1("USER", psUserName, PE_PF_STRING, 3)
        Call .Parameters1("RUC", cRUC, PE_PF_STRING, 4)
        Call .Parameters1("Resumen", opResumen(1).Value, PE_PF_BOOLEAN, 5)
    End With
    Call Formato.CargaResumen
End Sub

Private Sub PreparaDiarioSimplificadoAnalitico()
    With oImprimir
        Call .Parameters1("CIA", cCia, PE_PF_STRING, 0)
        Call .Parameters1("MESANO", UCase(oString.Periodo(CInt(Val(cMesPro) - 1))) & " " & cAnoPro, PE_PF_STRING, 1)
        Call .Parameters1("RUC", cRUC, PE_PF_STRING, 2)
    End With
    Call oArchivo.Delay(2)
End Sub

Private Sub VerificarArreglo(aDatos() As Currency, TipoNumero As String, Fecha As String, Detalle As String, Corrector As Integer, Grupo As Integer, Actualizar As Boolean)
    Dim iIndex As Integer
    For iIndex = 0 To 9
        If aDatos(iIndex) <> 0 Then
            Call PonerRegistro(TipoNumero, Fecha, Detalle, aDatos(iIndex), iIndex + Corrector, Grupo, Actualizar)
        End If
    Next
End Sub

Private Sub PonerRegistro(TipoNumero As String, Fecha As String, Detalle As String, Valor As Currency, Columna As Integer, Grupo As Integer, Actualizar As Boolean)
    On Error GoTo TraperErrors
    Dim iIndice As Integer
    Dim intLockCount As Integer
    Dim intRndCount As Integer

    With rs4
        If Not .EditMode <> 0 Then .AddNew
        !Grupo = Grupo
        !TipoNumero = TipoNumero
        !Fecha_Documento = IIf(Fecha = "01/01/1980", Format(oString.LastDayofMonth("01/" + oString.PadLeft(Str(Val(cMesPro) - 1), 2, "0") + "/" + cAnoPro), kFORMAT_TO_SAVE_DATE), Fecha)
        !Detalle = Detalle
        Select Case Columna
            Case 0
                !Columna1 = Valor
            Case 1
                !Columna2 = Valor
            Case 2
                !Columna3 = Valor
            Case 3
                !Columna4 = Valor
            Case 4
                !Columna5 = Valor
            Case 5
                !Columna6 = Valor
            Case 6
                !Columna7 = Valor
            Case 7
                !Columna8 = Valor
            Case 8
                !Columna9 = Valor
            Case 9
                !Columna10 = Valor
            Case 10
                !Columna11 = Valor
            Case 11
                !Columna12 = Valor
        End Select
        If Actualizar Then .Update
    End With
CleanExit:
    Exit Sub
    
TraperErrors:
    Select Case Err.Number
        Case 3260, 3197
            intLockCount = intLockCount + 1
            If intLockCount > 3 Then
                If MsgBox(Err.Description & " ¿Desea reintentar?", vbOKCancel, "Cuidado") = vbYes Then
                    intLockCount = 1
                Else
                    Resume CleanExit
                End If
            End If
            DoEvents
            intRndCount = intLockCount ^ 2 * Int(Rnd * 3000 + 1000)
            For iIndice = 1 To intRndCount: DoEvents: Next iIndice
            Resume
    
        Case -2147467259, -2147217864
            Err.Clear
            Resume
            
        Case Else ' Error inesperado.
            MsgBox Err.Description, vbCritical, "Error"
            Resume CleanExit
    End Select
End Sub

Private Function aCeros(aData() As Currency)
    Dim iIndex As Byte
    For iIndex = 0 To 9
        aData(iIndex) = 0
    Next
End Function

Private Function PrepararDiarioSimplificadoSeguro() As Boolean
    Dim iPuntero As Integer, sDesc As String
    iPuntero = Screen.MousePointer
    On Error GoTo Fallo
    Screen.MousePointer = vbHourglass
    PreparaDetalledelDiarioSimplificado
    Screen.MousePointer = iPuntero
    PrepararDiarioSimplificadoSeguro = True
    Exit Function
Fallo:
    sDesc = Err.Description
    Screen.MousePointer = iPuntero
    MsgBox "No se pudo preparar el Libro Diario Simplificado (Formato 5.2)." & vbCrLf & vbCrLf & sDesc, vbCritical, "Formato 5.2"
End Function

' Exporta el Formato 5.2 a Excel (sin guardar ni imprimir; Excel queda abierto).
Private Sub cmdExcel_Click()
    Dim iPuntero As Integer, sDesc As String, bAbierto As Boolean
    If iReporte <> kFormatos_5_2 Then Exit Sub
    iPuntero = Screen.MousePointer
    On Error GoTo Fallo
    cmdExcel.Enabled = False
    Screen.MousePointer = vbHourglass
    PreparaDetalledelDiarioSimplificado
    If D52_NumDetalles() > 0 Then
        bAbierto = ExcelDiarioSimplificadoF52(cCia, cRUC, kTITULO_FORMATO_5_2, UCase(oString.Periodo(CInt(Val(cMesPro) - 1))) & " " & cAnoPro)
    End If
    Screen.MousePointer = iPuntero
    cmdExcel.Enabled = True
    If Not bAbierto Then MsgBox "No hay movimientos para el período seleccionado.", vbInformation, "Formato 5.2"
    Exit Sub
Fallo:
    sDesc = Err.Description
    Screen.MousePointer = iPuntero
    cmdExcel.Enabled = True
    MsgBox "No se pudo generar el Excel del Formato 5.2." & vbCrLf & vbCrLf & sDesc, vbCritical, "Formato 5.2"
End Sub

' Inserta una transaccion sintetica (saldo inicial o resumen de compras/ventas).
' Debe/Haber se registran igual que antes (con los importes MN). Reintenta un
' numero acotado de veces ante bloqueos (3260/3197) y propaga cualquier otro
' error para no producir reportes contables incompletos en silencio.
Private Sub ActualizaTransacciones(rs4 As DAO.Recordset, Tipo As String, Numero As String, Cuenta As String, Fecha As String, Debe As Currency, Haber As Currency, DebeMN As Currency, HaberMN As Currency, DebeME As Currency, HaberME As Currency, Detalle As String)
    Dim iIntentos As Integer
    Dim lErr As Long, sSrc As String, sDesc As String
    On Error GoTo TraperErrors

    With rs4
        .AddNew
        !Tipo = Tipo
        !Numero = Numero
        !Cuenta = Cuenta
        !Auxiliar = Space(5)
        !Centro_de_Costo = Space(iLongCta)
        !Tipo_Documento = "00"
        !Abreviatura = "Var"
        !Numero_Documento = "Varios Doc."
        !Fecha_Documento = Format(Fecha, kFORMAT_TO_SAVE_DATE)
        !Vencimiento = Format(Fecha, kFORMAT_TO_SAVE_DATE)
        !Fecha_Operacion = Format(Fecha, kFORMAT_TO_SAVE_DATE)
        !Detalle = Detalle
        !Moneda = "N"
        !TipoNumero = Tipo + Numero
        !Cambio = 0
        !Tipo_de_Cambio = 0
        !Debe = DebeMN
        !Haber = HaberMN
        !DebeMN = DebeMN
        !HaberMN = HaberMN
        !DebeME = DebeME
        !HaberME = HaberME
        .Update
    End With
    Exit Sub

TraperErrors:
    Select Case Err.Number
        Case 3260, 3197
            iIntentos = iIntentos + 1
            If iIntentos <= kMAX_REINTENTOS_BLOQUEO Then
                Sleep 200 * iIntentos
                Resume
            End If
    End Select
    lErr = Err.Number: sSrc = Err.Source: sDesc = Err.Description
    On Error Resume Next
    If rs4.EditMode <> dbEditNone Then rs4.CancelUpdate
    On Error GoTo 0
    Err.Raise lErr, sSrc, "No se pudo registrar la transacción de la cuenta " & Cuenta & ": " & sDesc
End Sub

' Ejecuta una sentencia en dbWorkArea dentro de una transaccion propia
' (DBEngine.Workspaces(0), el mismo espacio de trabajo usado antes). Si falla,
' revierte solo esa transaccion y propaga el error con la sentencia.
Private Sub EjecutarSQL(ByVal sSql As String)
    Dim bTrans As Boolean
    Dim lErr As Long, sSrc As String, sDesc As String
    On Error GoTo Fallo
    DBEngine.Workspaces(0).BeginTrans
    bTrans = True
    dbWorkArea.Execute sSql, dbFailOnError
    DBEngine.Workspaces(0).CommitTrans
    bTrans = False
    Exit Sub
Fallo:
    lErr = Err.Number: sSrc = Err.Source: sDesc = Err.Description
    If bTrans Then
        On Error Resume Next
        DBEngine.Workspaces(0).Rollback
        On Error GoTo 0
    End If
    Err.Raise lErr, sSrc, sDesc & vbCrLf & "SQL: " & Left$(sSql, 400)
End Sub

Private Sub BorrarTabla(ByVal sTabla As String)
    If oArchivo.IsTableName(dbWorkArea, sTabla) Then
        dbWorkArea.Execute "DROP TABLE " & sTabla, dbFailOnError
    End If
End Sub

' Limpieza en rutas de error: no debe ocultar el error original.
Private Sub BorrarTablaSilencioso(ByVal sTabla As String)
    On Error Resume Next
    If oArchivo.IsTableName(dbWorkArea, sTabla) Then dbWorkArea.Execute "DROP TABLE " & sTabla
    Err.Clear
End Sub

Private Sub CerrarRecordset(rs As DAO.Recordset)
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Err.Clear
End Sub

Private Function NzMonto(ByVal v As Variant) As Currency
    If Not (IsNull(v) Or IsEmpty(v)) Then NzMonto = CCur(v)
End Function

' Preparacion compartida por el modo Analitico y Resumido:
' copia cfTra en sDestino (Orden = 1) y agrega
'   - lineas con centro de costo, con Cuenta = Centro_de_Costo (Orden = 4)
'   - su contrapartida en 7911101 con debe/haber invertidos
'     (MesPro = 'XX' y Orden = 3 en las que pasaron de haber a debe).
' Mismas sentencias y mismo orden que la version anterior (no cambia signos).
' Una fila con Debe > 0 y Haber > 0 conserva el comportamiento previo: su
' contrapartida queda solo en el haber por el importe del Debe.
' Nunca modifica cfTra; sTmpCC y sTmpContra son tablas de trabajo.
Private Sub PreparaTransaccionesConCostos(ByVal sDestino As String, ByVal sTmpCC As String, ByVal sTmpContra As String)
    Dim sCtaContra As String
    sCtaContra = Mid(kCTA_CONTRA_COSTOS, 1, iLongCta)

    BorrarTabla sTmpCC
    BorrarTabla sTmpContra
    BorrarTabla sDestino
    EjecutarSQL "SELECT * INTO " & sTmpCC & " FROM [" & dbEmpresa.Name & "]." & cfTra & " WHERE TRIM(Centro_de_Costo) <> ''"
    EjecutarSQL "UPDATE " & sTmpCC & " SET Cuenta = Centro_de_Costo"
    EjecutarSQL "UPDATE " & sTmpCC & " SET Centro_de_Costo = '" & Space(iLongCta) & "', Orden = 4"
    EjecutarSQL "SELECT * INTO " & sTmpContra & " FROM " & sTmpCC
    EjecutarSQL "UPDATE " & sTmpContra & " SET Cuenta = '" & sCtaContra & "', Haber = Debe, HaberMN = DebeMN, HaberME = DebeME WHERE Debe > 0"
    EjecutarSQL "UPDATE " & sTmpContra & " SET Debe = 0, DebeMN = 0, DebeME = 0 WHERE Debe > 0"
    EjecutarSQL "UPDATE " & sTmpContra & " SET Cuenta = '" & sCtaContra & "', Debe = Haber, DebeMN = HaberMN, DebeME = HaberME, Mespro = 'XX' WHERE Haber > 0 And Cuenta <> '" & sCtaContra & "'"
    EjecutarSQL "UPDATE " & sTmpContra & " SET Haber = 0, HaberMN = 0, HaberME = 0, Orden = 3 WHERE MesPro = 'XX'"
    EjecutarSQL "SELECT * INTO " & sDestino & " FROM [" & dbEmpresa.Name & "]." & cfTra
    EjecutarSQL "UPDATE " & sDestino & " SET ORDEN = 1"
    EjecutarSQL "INSERT INTO " & sDestino & " SELECT * FROM " & sTmpCC
    EjecutarSQL "INSERT INTO " & sDestino & " SELECT * FROM " & sTmpContra
End Sub

' Valida que todas las cuentas del diario se puedan clasificar por rango
' (dos digitos iniciales; clase 0 = Cuentas de Orden). Si no, informa cuales.
Private Sub ValidarCuentasDiario(ByVal sTabla As String)
    Dim rs As DAO.Recordset
    Dim sLista As String, lInvalidas As Long
    Dim lErr As Long, sSrc As String, sDesc As String
    On Error GoTo Fallo
    Set rs = dbWorkArea.OpenRecordset("SELECT Tr.Cuenta, Count(*) AS Lineas FROM " & sTabla & " Tr GROUP BY Tr.Cuenta ORDER BY Tr.Cuenta", dbOpenSnapshot)
    Do While Not rs.EOF
        If D52_ClasificarCuenta(rs!Cuenta) = 0 Then
            lInvalidas = lInvalidas + 1
            If lInvalidas <= 10 Then sLista = sLista & vbCrLf & "   '" & D52_Texto(rs!Cuenta) & "' (" & rs!Lineas & " línea(s))"
        End If
        rs.MoveNext
    Loop
    CerrarRecordset rs
    If lInvalidas > 0 Then
        Err.Raise D52_ERR_CUENTA_INVALIDA, "ValidarCuentasDiario", _
                  "Hay " & lInvalidas & " cuenta(s) vacías o que no empiezan con dos dígitos; " & _
                  "no se pueden agrupar (Activo Corriente ... Cuentas de Orden):" & sLista
    End If
    Exit Sub
Fallo:
    lErr = Err.Number: sSrc = Err.Source: sDesc = Err.Description
    CerrarRecordset rs
    Err.Raise lErr, sSrc, sDesc
End Sub

' Auxiliar (proveedor) de cada voucher de compras (TipoNumero 6*), para el CAR.
' Se toma el menor Auxiliar no vacio del voucher: independiente del orden de
' lectura (antes dependia de la primera linea leida de cada voucher).
Private Function AuxiliaresDeCompras(ByVal sTabla As String) As Collection
    Dim rs As DAO.Recordset, colAux As Collection
    Dim lErr As Long, sSrc As String, sDesc As String
    On Error GoTo Fallo
    Set colAux = New Collection
    Set rs = dbWorkArea.OpenRecordset("SELECT Ctb.TipoNumero, Min(Ctb.Auxiliar) AS AuxCompra FROM " & sTabla & " Ctb " & _
                                      "WHERE Left(Ctb.TipoNumero, 1) = '6' AND Ctb.Auxiliar Is Not Null AND Trim(Ctb.Auxiliar) <> '' " & _
                                      "GROUP BY Ctb.TipoNumero", dbOpenSnapshot)
    Do While Not rs.EOF
        If D52_Texto(rs!TipoNumero) <> "" Then colAux.Add D52_Texto(rs!AuxCompra), D52_Texto(rs!TipoNumero)
        rs.MoveNext
    Loop
    CerrarRecordset rs
    Set AuxiliaresDeCompras = colAux
    Exit Function
Fallo:
    lErr = Err.Number: sSrc = Err.Source: sDesc = Err.Description
    CerrarRecordset rs
    Err.Raise lErr, sSrc, sDesc
End Function

Private Function AuxiliarDeCompra(colAux As Collection, ByVal sTipoNumero As String) As String
    On Error Resume Next
    AuxiliarDeCompra = colAux(sTipoNumero)
    Err.Clear
End Function

' Prepara el Formato 5.2:
'   cfFile10 (detalle por cuenta/columna), cfFile11 (cabeceras de 14 cuentas) y
'   cfFile7 (filas agrupadas) para Crystal, y el modelo en memoria
'   (DllRSimple_Diario52) que usa la exportacion a Excel.
' Orden: seccion (Activo Corriente ... Analitica, Cuentas de Orden al final),
' Cuenta, saldo inicial, Fecha_Documento, TipoNumero, Fecha_Operacion.
' Ante cualquier error se cierran recordsets, el modelo queda "no listo" y se
' propaga el error al llamador.
Private Sub PreparaDetalledelDiarioSimplificado()
    Dim rs1 As DAO.Recordset, rs4 As DAO.Recordset, rsSaldos As DAO.Recordset
    Dim sSql As String, sLastDate As String
    Dim cDebeMN As Currency, cHaberMN As Currency, cResultado As Currency
    Dim colAux As Collection
    Dim lCab As Long, iCol As Integer, iSec As Integer, iIndex As Integer
    Dim sCuenta As String, sTipoNumero As String, sCAR As String, sSI As String
    Dim cMonto As Currency, bTipoTexto As Boolean, iLargoNombre As Integer
    Dim sCampos As String, sSumas As String, sMonto As String
    Dim lErr As Long, sSrc As String, sDesc As String

    On Error GoTo Fallo
    D52_Iniciar

    BorrarTabla cfFile14
    BorrarTabla cfFile10
    BorrarTabla cfFile11
    BorrarTabla cfFile
    BorrarTabla cfFile7
    BorrarTabla cfFile8
    BorrarTabla cfFile9

    Call oArchivo.CopyStructCTB(db, dbWorkArea, "xCtbDiario", cfFile10, False, iLongCta)
    Call oArchivo.CopyStructCTB(db, dbWorkArea, "xCtbDiario1", cfFile11, False, iLongCta)
    dbWorkArea.TableDefs.Refresh    ' No borrar
    If opResumen(0).Value Then
            Call AgregadoCtasCostos(cfFile)
        Else
            Call PreparaDiarioResumido
    End If
    BorrarTabla cfFile8
    BorrarTabla cfFile9
    EjecutarSQL "SELECT * INTO " & cfFile8 & " FROM [" & dbEmpresa.Name & "]." & cfPla
    EjecutarSQL "SELECT * INTO " & cfFile9 & " FROM [" & dbEmpresa.Name & "].xCtbBalance"

    ' Saldos iniciales: mismo criterio de periodo que la version anterior
    ' (fecha = ultimo dia del mes cMesPro - 2, saldo acumulado 1..cMesPro - 1).
    ' Ver README: la convencion de cMesPro la define el sistema llamador.
    Set rs4 = dbWorkArea.OpenRecordset(cfFile)
    Set rsSaldos = dbEmpresa.OpenRecordset(cfSa1)
    rsSaldos.Index = "Busqueda"
    If Val(cMesPro) > 1 Then
        Set rs1 = dbEmpresa.OpenRecordset("Select Sa.Cuenta, Pl.Analisis From " & cfSa1 & " Sa INNER JOIN " & cfPla & " Pl ON Sa.Cuenta = Pl.Cuenta WHERE Pl.Analisis AND Sa.Tipo = '01' ORDER BY Sa.Cuenta", dbOpenSnapshot)
        sLastDate = Format(oString.LastDayofMonth("01/" + oString.PadLeft(Str(Val(cMesPro) - 2), 2, "0") + "/" + cAnoPro), kFORMAT_TO_SAVE_DATE)
        With rs1
            Do While Not .EOF
                cDebeMN = 0
                cHaberMN = 0
                cResultado = oViewSaldos.ViewSaldo(!Cuenta, Space(11), Space(iLongCta), rsSaldos, 1, Val(cMesPro) - 1)
                If cResultado > 0 Then
                        cDebeMN = cResultado
                    Else
                        cHaberMN = cResultado * -1
                End If
                If cResultado <> 0 Then
                    ActualizaTransacciones rs4, "001", "000000", D52_Texto(!Cuenta), sLastDate, cDebeMN, cHaberMN, cDebeMN, cHaberMN, 0, 0, "Saldo Inicial"
                End If
                .MoveNext
            Loop
        End With
        CerrarRecordset rs1
    End If
    CerrarRecordset rs4
    CerrarRecordset rsSaldos

    ValidarCuentasDiario cfFile

    ' LEFT JOIN: una cuenta sin plan o sin tipo de balance ya no se descarta.
    sSql = "SELECT Tr.Cuenta, Tr.Fecha_Documento, Tr.Detalle AS Detalle, Tr.TipoNumero, Tr.Tipo_Documento, Tr.Serie_Documento, Tr.Numero_Documento, Tr.Auxiliar, Tr.DebeMN, Tr.HaberMN, Tr.Fecha_Operacion, Pl.Tipo, xCtb.Nombre INTO " & cfFile14 & " " & _
           "FROM (" & cfFile & " Tr LEFT JOIN " & cfFile8 & " Pl ON Tr.Cuenta = Pl.Cuenta) LEFT JOIN " & cfFile9 & " xCtb ON Pl.Tipo = xCtb.Id"
    EjecutarSQL sSql

    Set colAux = AuxiliaresDeCompras(cfFile14)

    ' Orden explicito del recordset (no se depende del orden fisico de SELECT INTO)
    Set rs1 = dbWorkArea.OpenRecordset("SELECT Ctb.* FROM " & cfFile14 & " Ctb ORDER BY IIf(Left(Ctb.Cuenta, 1) = '0', 1, 0), Ctb.Cuenta, " & _
                                       "IIf(Ctb.TipoNumero = '" & D52_TIPONUMERO_SALDO_INICIAL & "', 0, 1), Ctb.Fecha_Documento, Ctb.TipoNumero, Ctb.Fecha_Operacion", dbOpenSnapshot)
    D52_EstablecerFechasTipadas (rs1.Fields("Fecha_Documento").Type = dbDate)
    Set rsSaldos = dbWorkArea.OpenRecordset(cfFile11)
    Set rs4 = dbWorkArea.OpenRecordset(cfFile10)
    bTipoTexto = (rs4.Fields("Tipo").Type = dbText)
    iLargoNombre = 0
    If rs4.Fields("Nombre").Type = dbText Then iLargoNombre = rs4.Fields("Nombre").Size

    Do While Not rs1.EOF
        sCuenta = D52_Texto(rs1!Cuenta)
        lCab = D52_UbicarCuenta(sCuenta, iCol)
        iSec = D52_BloqueSeccion(lCab)
        sTipoNumero = D52_Texto(rs1!TipoNumero)
        sSI = D52_SI(sTipoNumero)
        sCAR = D52_CAR(sTipoNumero, D52_Texto(rs1!Tipo_Documento), D52_Texto(rs1!Serie_Documento), _
                       D52_Texto(rs1!Numero_Documento), AuxiliarDeCompra(colAux, sTipoNumero), cRUC)
        cMonto = D52_Monto(rs1!DebeMN, rs1!HaberMN)
        With rs4
            .AddNew
            !Cabecera = lCab
            !SI = sSI
            If bTipoTexto Then
                !Tipo = Format(iSec, "00")
            Else
                !Tipo = iSec
            End If
            If iLargoNombre > 0 Then
                !Nombre = Left$(D52_NombreSeccion(iSec), iLargoNombre)
            Else
                !Nombre = D52_NombreSeccion(iSec)
            End If
            !Cuenta = rs1!Cuenta
            !Fecha_Documento = rs1!Fecha_Documento
            !Detalle = rs1!Detalle
            !TipoNumero = sCAR
            !Fecha_Operacion = rs1!Fecha_Operacion
            .Fields("S" & CStr(iCol + 1)).Value = cMonto
            .Update
        End With
        D52_AgregarDetalle lCab, iCol, sCuenta, rs1!Fecha_Documento, rs1!Fecha_Operacion, sSI, sCAR, D52_Texto(rs1!Detalle), cMonto
        rs1.MoveNext
    Loop
    CerrarRecordset rs1
    CerrarRecordset rs4

    ' Cabeceras posicionales (Cabecera + 14 codigos de cuenta) para Crystal
    For lCab = 1 To D52_NumBloques()
        rsSaldos.AddNew
        rsSaldos!Cabecera = lCab
        For iIndex = 0 To D52_COLUMNAS - 1
            If iIndex < D52_BloqueNumCuentas(lCab) Then
                rsSaldos.Fields(iIndex + 2).Value = D52_BloqueCuenta(lCab, iIndex)
            Else
                rsSaldos.Fields(iIndex + 2).Value = ""
            End If
        Next
        rsSaldos.Update
    Next
    CerrarRecordset rsSaldos

    ' cfFile7: filas de debe y de haber por separado (criterio "Milton Flores").
    ' Cada linea de cfFile10 tiene un solo importe (S1..S14), por lo que la suma
    ' de las 14 columnas es el importe de la linea: particion exacta, sin duplicar
    ' sumas aunque las columnas no usadas valgan 0 en lugar de Null.
    sMonto = "(IIf(S1 Is Null, 0, S1)"
    For iIndex = 2 To D52_COLUMNAS
        sMonto = sMonto & " + IIf(S" & iIndex & " Is Null, 0, S" & iIndex & ")"
    Next
    sMonto = sMonto & ")"
    sSumas = "SUM(S1) AS S1a"
    For iIndex = 2 To D52_COLUMNAS
        sSumas = sSumas & ", SUM(S" & iIndex & ") AS S" & iIndex & "a"
    Next
    sCampos = "Cabecera, Tipo, Nombre, SI, TipoNumero, Fecha_Operacion, Detalle"
    BorrarTabla cfFile7
    EjecutarSQL "SELECT " & sCampos & ", " & sSumas & ", Fecha_Documento INTO " & cfFile7 & " FROM " & cfFile10 & _
                " WHERE " & sMonto & " > 0 GROUP BY " & sCampos & ", Fecha_Documento ORDER BY Cabecera, SI, Fecha_Documento, TipoNumero"
    EjecutarSQL "INSERT INTO " & cfFile7 & " SELECT " & sCampos & ", " & sSumas & ", Fecha_Documento FROM " & cfFile10 & _
                " WHERE " & sMonto & " <= 0 GROUP BY " & sCampos & ", Fecha_Documento ORDER BY Cabecera, SI, Fecha_Documento, TipoNumero"

    lRegistros = oArchivo.DAOExecuteReturn(dbWorkArea, "Select Count(*) From " & cfFile7, 0)
    D52_MarcarListo True
    Exit Sub

Fallo:
    lErr = Err.Number: sSrc = Err.Source: sDesc = Err.Description
    CerrarRecordset rs1
    CerrarRecordset rs4
    CerrarRecordset rsSaldos
    D52_MarcarListo False
    BorrarTablaSilencioso cfFile7
    Err.Raise lErr, sSrc, sDesc
End Sub

Private Sub AgregadoCtasCostos(File_Name As String)
    Dim sTmpCC As String, sTmpContra As String
    Dim lErr As Long, sSrc As String, sDesc As String
    On Error GoTo Fallo

    sTmpCC = oArchivo.FileTemp(dbWorkArea)
    sTmpContra = oArchivo.FileTemp(dbWorkArea)
    PreparaTransaccionesConCostos File_Name, sTmpCC, sTmpContra
    BorrarTabla sTmpCC
    BorrarTabla sTmpContra
    dbWorkArea.TableDefs.Refresh    ' No borrar
    Exit Sub

Fallo:
    lErr = Err.Number: sSrc = Err.Source: sDesc = Err.Description
    BorrarTablaSilencioso sTmpCC
    BorrarTablaSilencioso sTmpContra
    Err.Raise lErr, sSrc, sDesc
End Sub

' Modo Resumido: los vouchers de compras/ventas (Tipo 600..799) se reemplazan por
' un resumen por Tipo y Cuenta al cierre del mes. Las lineas se reparten en
' grupos excluyentes para conservar el saldo:
'   1) Debe > 0                      (igual que antes)
'   2) Debe <= 0 y Haber > 0         (antes una fila con Debe y Haber > 0 se sumaba dos veces)
'   3) Debe <= 0 y Haber <= 0        (antes se perdian, p.ej. importes negativos)
Private Sub PreparaDiarioResumido()
    Dim rs4 As DAO.Recordset
    Dim sLastDate As String
    Dim lErr As Long, sSrc As String, sDesc As String
    Const sDEBE As String = "IIf(Tr.Debe Is Null, 0, Tr.Debe)"
    Const sHABER As String = "IIf(Tr.Haber Is Null, 0, Tr.Haber)"
    On Error GoTo Fallo

    dbWorkArea.TableDefs.Refresh    ' No borrar
    BorrarTabla cfFile7
    BorrarTabla cfFile8
    BorrarTabla cfFile9
    PreparaTransaccionesConCostos cfFile, cfFile7, cfFile8
    EjecutarSQL "SELECT * INTO " & cfFile9 & " FROM " & cfFile & " WHERE Tipo BETWEEN '600' And '799'"
    EjecutarSQL "DELETE FROM " & cfFile & " WHERE Tipo BETWEEN '600' And '799'"

    Set rs4 = dbWorkArea.OpenRecordset(cfFile)
    sLastDate = Format(oString.LastDayofMonth("01/" + oString.PadLeft(Str(Val(cMesPro) - 1), 2, "0") + "/" + cAnoPro), kFORMAT_TO_SAVE_DATE)
    AgregarResumenComprasVentas rs4, sDEBE & " > 0", sLastDate
    AgregarResumenComprasVentas rs4, sDEBE & " <= 0 AND " & sHABER & " > 0", sLastDate
    AgregarResumenComprasVentas rs4, sDEBE & " <= 0 AND " & sHABER & " <= 0", sLastDate
    CerrarRecordset rs4
    lRegistros = oArchivo.DAOExecuteReturn(dbWorkArea, "Select Count(*) From " & cfFile, 0)
    Exit Sub

Fallo:
    lErr = Err.Number: sSrc = Err.Source: sDesc = Err.Description
    CerrarRecordset rs4
    Err.Raise lErr, sSrc, sDesc
End Sub

Private Sub AgregarResumenComprasVentas(rs4 As DAO.Recordset, ByVal sFiltro As String, ByVal sLastDate As String)
    Dim rs1 As DAO.Recordset
    Dim cDebe As Currency, cHaber As Currency, cDebeMN As Currency, cHaberMN As Currency, cDebeME As Currency, cHaberME As Currency
    Dim lErr As Long, sSrc As String, sDesc As String
    On Error GoTo Fallo

    Set rs1 = dbWorkArea.OpenRecordset("Select Tr.Tipo, Tr.Cuenta, SUM(Tr.Debe) AS Debe, SUM(Tr.Haber) As Haber, SUM(Tr.DebeMN) As DebeMN, SUM(Tr.HaberMN) As HaberMN, SUM(Tr.DebeME) As DebeME, SUM(Tr.HaberME) As HaberME From " & cfFile9 & " Tr " & _
                                       "WHERE " & sFiltro & " Group by Tr.Tipo, Tr.Cuenta ORDER BY Tr.Tipo, Tr.Cuenta", dbOpenSnapshot)
    With rs1
        Do While Not .EOF
            cDebe = NzMonto(!Debe): cHaber = NzMonto(!Haber)
            cDebeMN = NzMonto(!DebeMN): cHaberMN = NzMonto(!HaberMN)
            cDebeME = NzMonto(!DebeME): cHaberME = NzMonto(!HaberME)
            If cDebe <> 0 Or cHaber <> 0 Or cDebeMN <> 0 Or cHaberMN <> 0 Or cDebeME <> 0 Or cHaberME <> 0 Then
                ActualizaTransacciones rs4, D52_Texto(!Tipo), "000000", D52_Texto(!Cuenta), sLastDate, cDebe, cHaber, cDebeMN, cHaberMN, cDebeME, cHaberME, "POR LAS " & IIf(Mid(D52_Texto(!Tipo), 1, 1) = "6", "COMPRAS", "VENTAS") & " DEL MES"
            End If
            .MoveNext
        Loop
    End With
    CerrarRecordset rs1
    Exit Sub

Fallo:
    lErr = Err.Number: sSrc = Err.Source: sDesc = Err.Description
    CerrarRecordset rs1
    Err.Raise lErr, sSrc, sDesc
End Sub
