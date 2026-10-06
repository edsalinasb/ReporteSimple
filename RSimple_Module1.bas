Attribute VB_Name = "DllRSimple_Module1"
Option Explicit

Public cAnoPro As String
Public cMesPro As String
Public cEmpPro As String
Public cCia As String
Public psUserPassword As String
Public psUserPasswordDBase As String
Public cLocalizacionExe As String
Public iLongCta As Integer
Public cFechSys As String
Public psUserName As String
Public cRUC As String
Public cPath As String
Public cWhere As String
Public db As DAO.Database
Public dbEmpresa As DAO.Database
Public dbWorkArea As DAO.Database
Public Formato As Form
Public oArchivo As Archivodb
Public oString As MString
Public oImprimir As clsViewReport
Public mvarFormMain As Form
