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
                Call PreparaDetalledelDiarioSimplificado
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

Private Function ActualizaTransacciones(rs4 As DAO.Recordset, Tipo As String, Numero As String, Cuenta As String, Fecha As String, Debe As Currency, Haber As Currency, DebeMN As Currency, HaberMN As Currency, DebeME As Currency, HaberME As Currency, Detalle As String)
    On Error GoTo TraperErrors
    Dim iIndice As Integer
    Dim intLockCount As Integer
    Dim intRndCount As Integer
     
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
    
CleanExit:
    Exit Function
    
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
End Function

Private Sub PreparaDetalledelDiarioSimplificado()
    Dim rs1 As DAO.Recordset, rs4 As DAO.Recordset, rsSaldos As DAO.Recordset
    Dim sSql As String, sLastDate As String, lRegistros As Long
    
    If oArchivo.IsTableName(dbWorkArea, cfFile14) Then
        Call dbWorkArea.Execute("Drop Table " & cfFile14)
    End If
    
    If oArchivo.IsTableName(dbWorkArea, cfFile10) Then
        Call dbWorkArea.Execute("Drop Table " & cfFile10)
    End If
    
    If oArchivo.IsTableName(dbWorkArea, cfFile11) Then
        Call dbWorkArea.Execute("Drop Table " & cfFile11)
    End If
    
    If oArchivo.IsTableName(dbWorkArea, cfFile) Then
        Call dbWorkArea.Execute("Drop Table " & cfFile)
    End If

    If oArchivo.IsTableName(dbWorkArea, cfFile7) Then
        Call dbWorkArea.Execute("Drop Table " & cfFile7)
    End If
    
    If oArchivo.IsTableName(dbWorkArea, cfFile8) Then
        Call dbWorkArea.Execute("Drop Table " & cfFile8)
    End If
    If oArchivo.IsTableName(dbWorkArea, cfFile9) Then
        Call dbWorkArea.Execute("Drop Table " & cfFile9)
    End If
    
    Call oArchivo.CopyStructCTB(db, dbWorkArea, "xCtbDiario", cfFile10, False, iLongCta)
    Call oArchivo.CopyStructCTB(db, dbWorkArea, "xCtbDiario1", cfFile11, False, iLongCta)
    dbWorkArea.TableDefs.Refresh    ' No borrar
    If opResumen(0).Value Then
            Call AgregadoCtasCostos(cfFile)
        Else
            Call PreparaDiarioResumido
    End If
    If oArchivo.IsTableName(dbWorkArea, cfFile8) Then
        Call dbWorkArea.Execute("Drop Table " & cfFile8)
    End If
    If oArchivo.IsTableName(dbWorkArea, cfFile9) Then
        Call dbWorkArea.Execute("Drop Table " & cfFile9)
    End If
    sSql = "SELECT * INTO " & cfFile8 & " FROM [" & dbEmpresa.Name & "]." & cfPla
    dbWorkArea.Execute sSql
    sSql = "SELECT * INTO " & cfFile9 & " FROM [" & dbEmpresa.Name & "].xCtbBalance"
    dbWorkArea.Execute sSql

    Set rs4 = dbWorkArea.OpenRecordset(cfFile)
    Set rsSaldos = dbEmpresa.OpenRecordset(cfSa1)
    rsSaldos.Index = "Busqueda"
    If Val(cMesPro) > 1 Then
        Set rs1 = dbEmpresa.OpenRecordset("Select Sa.Cuenta, Pl.Analisis From " & cfSa1 & " Sa INNER JOIN " & cfPla & " Pl ON Sa.Cuenta = Pl.Cuenta WHERE Pl.Analisis AND Sa.Tipo = '01'")
        If Val(cMesPro) > 1 Then
                sLastDate = Format(oString.LastDayofMonth("01/" + oString.PadLeft(Str(Val(cMesPro) - 2), 2, "0") + "/" + cAnoPro), kFORMAT_TO_SAVE_DATE)
            Else
                sLastDate = Format(oString.LastDayofMonth("31/12/" + Trim(Str(Val(cAnoPro) - 1))), kFORMAT_TO_SAVE_DATE)
        End If
        
        Dim cDebeMN As Currency, cHaberMN As Currency, cResultado As Currency
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
                    ActualizaTransacciones rs4, "001", "000000", !Cuenta, sLastDate, cDebeMN, cHaberMN, cDebeMN, cHaberMN, 0, 0, "Saldo Inicial"
                End If
                .MoveNext
            Loop
            .Close
        End With
    End If
    rs4.Close
    rsSaldos.Close
    Set rs1 = Nothing
    Set rs4 = Nothing
    Set rsSaldos = Nothing
    
    sSql = "SELECT Tr.Cuenta, Tr.Fecha_Documento, Tr.Detalle AS Detalle, Tr.TipoNumero, Tr.Tipo_Documento, Tr.Serie_Documento, Tr.Numero_Documento, Tr.Auxiliar, Tr.DebeMN, Tr.HaberMN, Tr.Fecha_Operacion, Pl.Tipo, xCtb.Nombre INTO " & cfFile14 & " " & _
           "From (" & cfFile & " Tr INNER JOIN " & cfFile8 & " Pl ON Tr.Cuenta = Pl.Cuenta) INNER JOIN " & cfFile9 & " xCtb ON Pl.Tipo = xCtb.Id ORDER BY Pl.Tipo ASC, Tr.Cuenta ASC, Tr.Fecha_Operacion ASC, Tr.TipoNumero ASC"
    dbWorkArea.Execute sSql
    If oArchivo.IsTableName(dbWorkArea, cfFile7) Then dbWorkArea.Execute ("DROP TABLE " & cfFile7)
    Set rs1 = dbWorkArea.OpenRecordset(cfFile14)
    Set rsSaldos = dbWorkArea.OpenRecordset(cfFile11)
    Set rs4 = dbWorkArea.OpenRecordset(cfFile10)
    Dim X As Integer, ay(14) As String, iIndex As Integer, iCabecera As Integer, iTipo As Integer
    X = -1
    iCabecera = 1
    Dim sTipoNumero As String, sAuxiliar As String
    With rs4
        Do While Not rs1.EOF
            iTipo = rs1!Tipo
            Do While Not rs1.EOF
                If iTipo <> rs1!Tipo Then Exit Do
                If oString.ASearch(ay, rs1!Cuenta) = -1 Then
                        If X >= 13 Then
                                rsSaldos.AddNew
                                rsSaldos!Cabecera = iCabecera
                                For iIndex = 0 To 13
                                    rsSaldos.Fields(iIndex + 2).Value = ay(iIndex)
                                    ay(iIndex) = ""
                                Next
                                rsSaldos.Update
                                iCabecera = iCabecera + 1
                                X = 0
                            Else
                                X = X + 1
                        End If
                        ay(X) = rs1!Cuenta
                End If
                If Mid(rs1!TipoNumero, 1, 1) = "6" Then
                    If sTipoNumero <> rs1!TipoNumero Then
                        sTipoNumero = rs1!TipoNumero
                        sAuxiliar = ""
                        If rs1!Auxiliar <> "" Then
                            sAuxiliar = rs1!Auxiliar
                        End If
                    End If
                End If
                .AddNew
                !Cabecera = iCabecera
                !SI = IIf(rs1!TipoNumero = "001000000", "1", "2")
                !Tipo = rs1!Tipo
                !Nombre = rs1!Nombre
                !Cuenta = rs1!Cuenta
                !Fecha_Documento = rs1!Fecha_Documento
                !Detalle = rs1!Detalle
                If Mid(rs1!TipoNumero, 1, 1) = "7" Then
                        !TipoNumero = Right(cRUC, 11) & rs1!Tipo_Documento + Trim(rs1!Serie_Documento) + rs1!Numero_Documento
                    ElseIf Mid(rs1!TipoNumero, 1, 1) = "6" Then
                        !TipoNumero = sAuxiliar & rs1!Tipo_Documento + Trim(rs1!Serie_Documento) + rs1!Numero_Documento
                    Else
                        !TipoNumero = rs1!TipoNumero
                End If
                !Fecha_Operacion = rs1!Fecha_Operacion
                .Fields(X + 10).Value = rs1!DebeMN - rs1!HaberMN
                .Update
                rs1.MoveNext
            Loop
            rsSaldos.AddNew
            rsSaldos!Cabecera = iCabecera
            For iIndex = 0 To 13
                rsSaldos.Fields(iIndex + 2).Value = ay(iIndex)
                ay(iIndex) = ""
            Next
            rsSaldos.Update
            iCabecera = iCabecera + 1
            X = -1
        Loop
        .Close
    End With
    rsSaldos.Close
    rs1.Close
    DBEngine.BeginTrans
    ' Original
'    Call dbWorkArea.Execute("SELECT Cabecera, Tipo, Nombre, SI, TipoNumero, Fecha_Operacion, Detalle, SUM(S1) AS S1a, SUM(S2) AS S2a, SUM(S3) AS S3a, SUM(S4) AS S4a, SUM(S5) AS S5a, SUM(S6) AS S6a, SUM(S7) AS S7a, SUM(S8) AS S8a, SUM(S9) AS S9a, SUM(S10) AS S10a, SUM(S11) AS S11a, SUM(S12) AS S12a, SUM(S13) AS S13a, SUM(S14) AS S14a INTO " & cfFile7 & " " & _
           "From " & cfFile10 & " GROUP BY Cabecera, Tipo, Nombre, SI, TipoNumero, Fecha_Operacion, Detalle ORDER BY Cabecera, Tipo, Nombre, SI, TipoNumero, Fecha_Operacion, Detalle")
    
'   Modificado para Milton Flores
    Call dbWorkArea.Execute("SELECT Cabecera, Tipo, Nombre, SI, TipoNumero, Fecha_Operacion, Detalle, SUM(S1) AS S1a, SUM(S2) AS S2a, SUM(S3) AS S3a, SUM(S4) AS S4a, SUM(S5) AS S5a, SUM(S6) AS S6a, SUM(S7) AS S7a, SUM(S8) AS S8a, SUM(S9) AS S9a, SUM(S10) AS S10a, SUM(S11) AS S11a, SUM(S12) AS S12a, SUM(S13) AS S13a, SUM(S14) AS S14a INTO " & cfFile7 & " " & _
           "From " & cfFile10 & " WHERE S1 > 0 OR S2 > 0 OR S3 > 0 OR S4 > 0 OR S5 > 0 OR S6 > 0 OR S7 > 0 OR S8 > 0 OR S9 > 0 OR S10 > 0 OR S11 > 0 OR S12 > 0 OR S13 > 0 OR S14 > 0 GROUP BY Cabecera, Tipo, Nombre, SI, TipoNumero, Fecha_Operacion, Detalle ORDER BY Cabecera, Tipo, Nombre, SI, TipoNumero, Fecha_Operacion, Detalle")
           
    Call dbWorkArea.Execute("INSERT INTO " & cfFile7 & " SELECT Cabecera, Tipo, Nombre, SI, TipoNumero, Fecha_Operacion, Detalle, SUM(S1) AS S1a, SUM(S2) AS S2a, SUM(S3) AS S3a, SUM(S4) AS S4a, SUM(S5) AS S5a, SUM(S6) AS S6a, SUM(S7) AS S7a, SUM(S8) AS S8a, SUM(S9) AS S9a, SUM(S10) AS S10a, SUM(S11) AS S11a, SUM(S12) AS S12a, SUM(S13) AS S13a, SUM(S14) AS S14a " & _
           "From " & cfFile10 & " WHERE S1 <= 0 OR S2 <= 0 OR S3 <= 0 OR S4 <= 0 OR S5 <= 0 OR S6 <= 0 OR S7 <= 0 OR S8 <= 0 OR S9 <= 0 OR S10 <= 0 OR S11 <= 0 OR S12 <= 0 OR S13 <= 0 OR S14 <= 0 GROUP BY Cabecera, Tipo, Nombre, SI, TipoNumero, Fecha_Operacion, Detalle ORDER BY Cabecera, Tipo, Nombre, SI, TipoNumero, Fecha_Operacion, Detalle")
           
    DBEngine.CommitTrans
    lRegistros = oArchivo.DAOExecuteReturn(dbWorkArea, "Select Count(*) From " & cfFile7, 0)
End Sub

Private Sub AgregadoCtasCostos(File_Name As String)
    Dim sSql As String
    Dim sfFile7 As String, sfFile8 As String, sfFile9 As String
    
    sfFile7 = oArchivo.FileTemp(dbWorkArea)
    sfFile8 = oArchivo.FileTemp(dbWorkArea)
    sfFile9 = oArchivo.FileTemp(dbWorkArea)
    
    sSql = "SELECT * INTO " & sfFile7 & " FROM [" & dbEmpresa.Name & "]." & cfTra & " WHERE TRIM(Centro_de_Costo) <> ''"
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "UPDATE " & sfFile7 & " SET Cuenta = Centro_de_Costo"
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "UPDATE " & sfFile7 & " SET Centro_de_Costo = '" & Space(iLongCta) & "', Orden = 4"
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "SELECT * INTO " & sfFile8 & " FROM " & sfFile7
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "UPDATE " & sfFile8 & " SET Cuenta = '" & Mid("7911101     ", 1, iLongCta) & "', Haber = Debe, HaberMN = DebeMN, HaberME = DebeME WHERE Debe > 0"
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "UPDATE " & sfFile8 & " SET Debe = 0, DebeMN = 0, DebeME = 0 WHERE Debe > 0"
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "UPDATE " & sfFile8 & " SET Cuenta = '" & Mid("7911101     ", 1, iLongCta) & "', Debe = Haber, DebeMN = HaberMN, DebeME = HaberME, Mespro = 'XX' WHERE Haber > 0 And Cuenta <> '" & Mid("7911101     ", 1, iLongCta) & "'"
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "UPDATE " & sfFile8 & " SET Haber = 0, HaberMN = 0, HaberME = 0, Orden = 3 WHERE MesPro = 'XX'"
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "SELECT * INTO " & File_Name & " FROM [" & dbEmpresa.Name & "]." & cfTra
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "UPDATE " & File_Name & " SET ORDEN = 1"
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "INSERT INTO " & File_Name & " SELECT * FROM " & sfFile7
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "INSERT INTO " & File_Name & " SELECT * FROM " & sfFile8
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    dbWorkArea.TableDefs.Refresh    ' No borrar
    If oArchivo.IsTableName(dbWorkArea, sfFile7) Then
        DBEngine.BeginTrans
        Call dbWorkArea.Execute("Drop Table " & sfFile7)
        DBEngine.CommitTrans
    End If
    If oArchivo.IsTableName(dbWorkArea, sfFile8) Then
        DBEngine.BeginTrans
        Call dbWorkArea.Execute("Drop Table " & sfFile8)
        DBEngine.CommitTrans
    End If
    If oArchivo.IsTableName(dbWorkArea, sfFile9) Then
        DBEngine.BeginTrans
        Call dbWorkArea.Execute("Drop Table " & sfFile9)
        DBEngine.CommitTrans
    End If
    dbWorkArea.TableDefs.Refresh    ' No borrar
End Sub

Private Sub PreparaDiarioResumido()
    Dim rs1 As DAO.Recordset
    Dim rs4 As DAO.Recordset
    Dim sSql As String, sLastDate As String, lRegistros As Long
    
    dbWorkArea.TableDefs.Refresh    ' No borrar
    If oArchivo.IsTableName(dbWorkArea, cfFile) Then
        DBEngine.BeginTrans
        Call dbWorkArea.Execute("Drop Table " & cfFile)
        DBEngine.CommitTrans
    End If
    If oArchivo.IsTableName(dbWorkArea, cfFile7) Then
        DBEngine.BeginTrans
        Call dbWorkArea.Execute("Drop Table " & cfFile7)
        DBEngine.CommitTrans
    End If
    If oArchivo.IsTableName(dbWorkArea, cfFile8) Then
        DBEngine.BeginTrans
        Call dbWorkArea.Execute("Drop Table " & cfFile8)
        DBEngine.CommitTrans
    End If
    If oArchivo.IsTableName(dbWorkArea, cfFile9) Then
        DBEngine.BeginTrans
        Call dbWorkArea.Execute("Drop Table " & cfFile9)
        DBEngine.CommitTrans
    End If

    sSql = "SELECT * INTO " & cfFile7 & " FROM [" & dbEmpresa.Name & "]." & cfTra & " WHERE TRIM(Centro_de_Costo) <> ''"
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "UPDATE " & cfFile7 & " SET Cuenta = Centro_de_Costo"
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "UPDATE " & cfFile7 & " SET Centro_de_Costo = '" & Space(iLongCta) & "', Orden = 4"
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "SELECT * INTO " & cfFile8 & " FROM " & cfFile7
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "UPDATE " & cfFile8 & " SET Cuenta = '" & Mid("7911101     ", 1, iLongCta) & "', Haber = Debe, HaberMN = DebeMN, HaberME = DebeME WHERE Debe > 0"
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "UPDATE " & cfFile8 & " SET Debe = 0, DebeMN = 0, DebeME = 0 WHERE Debe > 0"
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "UPDATE " & cfFile8 & " SET Cuenta = '" & Mid("7911101     ", 1, iLongCta) & "', Debe = Haber, DebeMN = HaberMN, DebeME = HaberME, Mespro = 'XX' WHERE Haber > 0 And Cuenta <> '" & Mid("7911101     ", 1, iLongCta) & "'"
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "UPDATE " & cfFile8 & " SET Haber = 0, HaberMN = 0, HaberME = 0, Orden = 3 WHERE MesPro = 'XX'"
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "SELECT * INTO " & cfFile & " FROM [" & dbEmpresa.Name & "]." & cfTra
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "UPDATE " & cfFile & " SET ORDEN = 1"
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "INSERT INTO " & cfFile & " SELECT * FROM " & cfFile7
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    sSql = "INSERT INTO " & cfFile & " SELECT * FROM " & cfFile8
    DBEngine.BeginTrans
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    DBEngine.BeginTrans
    sSql = "SELECT * INTO " & cfFile9 & " FROM " & cfFile & " WHERE Tipo BETWEEN '600' And '799'"
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    DBEngine.BeginTrans
    sSql = "DELETE FROM " & cfFile & " WHERE Tipo BETWEEN '600' And '799'"
    dbWorkArea.Execute sSql
    DBEngine.CommitTrans
    
    Set rs4 = dbWorkArea.OpenRecordset(cfFile)
    Set rs1 = dbWorkArea.OpenRecordset("Select Tr.Tipo, Tr.Cuenta, SUM(Tr.Debe) AS Debe, SUM(Tr.Haber) As Haber, SUM(Tr.DebeMN) As DebeMN, SUM(Tr.HaberMN) As HaberMN, SUM(Tr.DebeME) As DebeME, SUM(Tr.HaberME) As HaberME From " & cfFile9 & " Tr WHERE Tr.Debe > 0 Group by Tipo, Cuenta")
    sLastDate = Format(oString.LastDayofMonth("01/" + oString.PadLeft(Str(Val(cMesPro) - 1), 2, "0") + "/" + cAnoPro), kFORMAT_TO_SAVE_DATE)
    With rs1
        Do While Not .EOF
            ActualizaTransacciones rs4, !Tipo, "000000", !Cuenta, sLastDate, !Debe, !Haber, !DebeMN, !HaberMN, !DebeME, !HaberME, "POR LAS " & IIf(Mid(!Tipo, 1, 1) = "6", "COMPRAS", "VENTAS") & " DEL MES"
            .MoveNext
        Loop
        .Close
    End With
    Set rs1 = dbWorkArea.OpenRecordset("Select Tr.Tipo, Tr.Cuenta, SUM(Tr.Debe) AS Debe, SUM(Tr.Haber) As Haber, SUM(Tr.DebeMN) As DebeMN, SUM(Tr.HaberMN) As HaberMN, SUM(Tr.DebeME) As DebeME, SUM(Tr.HaberME) As HaberME From " & cfFile9 & " Tr WHERE Tr.Haber > 0 Group by Tipo, Cuenta")
    With rs1
        Do While Not .EOF
            ActualizaTransacciones rs4, !Tipo, "000000", !Cuenta, sLastDate, !Debe, !Haber, !DebeMN, !HaberMN, !DebeME, !HaberME, "POR LAS " & IIf(Mid(!Tipo, 1, 1) = "6", "COMPRAS", "VENTAS") & " DEL MES"
            .MoveNext
        Loop
        .Close
    End With
    rs4.Close
    Set rs1 = Nothing
    Set rs4 = Nothing
    lRegistros = oArchivo.DAOExecuteReturn(dbWorkArea, "Select Count(*) From " & cfFile, 0)
End Sub

