VERSION 5.00
Object = "{A30B2DDF-C00F-469F-A23C-D6177608A128}#10.5#0"; "crviewer.dll"
Begin VB.Form preview 
   Caption         =   "Form1"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14580
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10305
   ScaleWidth      =   14580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd‘ÈÏÔÎ¸„ÈÔ”Â 
      Caption         =   "‘ÈÏÔÎ¸„ÈÔ ÛÂ ≈ÍÙı˘Ùﬁ"
      Height          =   360
      Left            =   5880
      TabIndex        =   3
      Top             =   0
      Width           =   2655
   End
   Begin VB.CommandButton cmd‘ÈÏÔÎÔ„ÈÔ”Â 
      Caption         =   "‘ÈÏÔÎÔ„ÈÔ ÛÂ PDF"
      Height          =   360
      Left            =   1800
      TabIndex        =   2
      Top             =   0
      Width           =   2655
   End
   Begin CrystalActiveXReportViewerLib105Ctl.CrystalActiveXReportViewer CrystalActiveXReportViewer1 
      Height          =   11415
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   14295
      _cx             =   25215
      _cy             =   20135
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
   Begin VB.Label lblID_NUM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   9960
      Width           =   45
   End
End
Attribute VB_Name = "preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Application As New CRAXdDRT.Application
Public REPORT As New CRAXdDRT.REPORT



'Public Sub printingCR9(ByVal mfile As String)

Sub PREVIEW()

Dim mdsn As String
Dim mdatabase As String
Dim musr As String
Dim mpwd As String


Open "C:\mercvb\reportconn.TXT" For Input As #1
    Line Input #1, mdsn
    Line Input #1, mdatabase
    Line Input #1, musr
    Line Input #1, mpwd

Close #1


mfile = "C:\MERCVB\REPORTS\TIMOL51.RPT"

Dim crxApp As CRAXdDRT.Application   ' .Application




Dim REPORT As CRAXdDRT.REPORT '     .Report
Set crxApp = New CRAXdDRT.Application
Set REPORT = crxApp.OpenReport(mfile)
'Declare a Connection Info Object
Dim ConnectionInfo As CRAXdDRT.ConnectionProperties
'Set the Connection Info to Connection Properties of ëthe table object
'Set ConnectionInfo = crxRpt.Database.Tables(1).ConnectionProperties


Set ConnectionInfo = REPORT.Database.Tables(1).ConnectionProperties
ConnectionInfo.Item("Use DSN Default Properties") = True '"MERCHP"



ConnectionInfo.Item("DSN") = "mercsql"
'ConnectionInfo.Item("DATABASE") = "MERCURY" '"62.103.69.140,65527" 'DELLAGAKIS\SQL17"
'ConnectionInfo.Item("Initial Catalog") = "EMP"
ConnectionInfo.Item("User ID") = "sa"
ConnectionInfo.Item("Password") = "12345678"


Dim Q As String
Q = "SELECT EGGTIM.POSO, EGGTIM.MONA, EGGTIM.KODE, EGGTIM.ERGO, EGGTIM.EKPT, EGGTIM.KAU_AJIA, EGGTIM.ONOMA, EGGTIM.KOLA, EGGTIM.PROELEYSH, EGGTIM.EKPT2, EGGTIM.XVRA, EGGTIM.LITRA, PEL.EPO, PEL.POL, PEL.EPA, PEL.TYP, TIM.ID_NUM "
Q = Q + " FROM  EGGTIM INNER JOIN  TIM ON EGGTIM.ID_NUM=TIM.ID_NUM "
Q = Q + " INNER JOIN  PEL ON (TIM.EIDOS=PEL.EIDOS) AND (TIM.KPE=PEL.KOD) "
Q = Q + " INNER JOIN PARASTAT ON LEFT(TIM.ATIM,1)=PARASTAT.EIDOS"
Q = Q + " Where TIM.ID_NUM = " + lblID_NUM.Caption + " ORDER BY TIM.ID_NUM"

REPORT.SQLQueryString = Q

CrystalActiveXReportViewer1.ReportSource = REPORT
CrystalActiveXReportViewer1.ViewReport




'ÂÓ·„˘„Á ÛÂ pdf
' ExportReportToPDF REPORT, "c:\mercvb\Beds.pdf", "Beds Held"





'typvnei to report
'REPORT.PrintOut False, 1 'NoCopies

End Sub






Private Sub ExportReportToPDF(ReportObject As CRAXdDRT.REPORT, ByVal FILENAME As String, ByVal ReportTitle As String)
    
    Dim objExportOptions As CRAXdDRT.ExportOptions
 
    ReportObject.ReportTitle = ReportTitle
    
    With ReportObject
        .EnableParameterPrompting = False
        .MorePrintEngineErrorMessages = True
    End With
    
    Set objExportOptions = ReportObject.ExportOptions
    
    With objExportOptions
        .DestinationType = crEDTDiskFile
        .DiskFileName = FILENAME
        .FormatType = crEFTPortableDocFormat
        .PDFExportAllPages = True
    End With
 
    ReportObject.Export False
 
End Sub






Private Sub cmd‘ÈÏÔÎÔ„ÈÔ”Â_Click()


'ÂÓ·„˘„Á ÛÂ pdf
 ExportReportToPDF REPORT, "c:\mercvb\Beds.pdf", "Beds Held"





'typvnei to report
'REPORT.PrintOut False, 1 'NoCopies

End Sub

Private Sub cmd‘ÈÏÔÎ¸„ÈÔ”Â_Click()


'ÂÓ·„˘„Á ÛÂ pdf
' ExportReportToPDF REPORT, "c:\mercvb\Beds.pdf", "Beds Held"





'typvnei to report
REPORT.PrintOut False, 1 'NoCopies

End Sub

'Set m_Prop = crxRpt.Database.Tables(1).ConnectionProperties
'
'With m_Prop
'.Item("Use DSN Default Properties") = False '  "MERCSQL" ' MYODBCî
'.Item("User ID") = "sa" ' :ìMYUSERNAMEî
'.Item("Password") = "12345678"  'ìMYPASSWORDî
''.Item("ServerName") = "aaa" '"Connection String") = "dsn=mercsql;"
'
'
'End With




'Set the ODBC DSN
'ConnectionInfo.Item("Use DSN Default Properties") = True '"MERCHP"
'Set the database name
'ConnectionInfo.Item("ase") = mdatabase ' "MERCURY"
'Set the user name


'ConnectionInfo.Item("User ID") = musr '"sa"


'.ConnectionInfo.Item("Server") = "DELLAGAKIS\SQL17"
'ConnectionInfo.Item("DatabaseName") = "EMP"


'Set the password





'Set m_Prop = crxRpt.Database.Tables(1).ConnectionProperties
'ë
'With m_Prop
'.Item("DSN") = "MERCSQL" ' MYODBCî
'.Item("User ID") = "sa" ' :ìMYUSERNAMEî
'.Item("Password") = "12345678"  'ìMYPASSWORDî
'End With
'With CRViewer91
'.ReportSource = m_Report
'.ViewReport
'End With








''Dim crxApp As CRAXDRT.Application
''Dim crxRpt As CRAXDRT.Report
'crxApp.LogOnServer "p2sodbc.DLL", mLogonServer, mLogonDB, mlogonUserID, mpwd
'Set crxApp = New CRAXDRT.Application
'Set crxRpt = crxApp.OpenReport(mfile)
'crxRpt.ReadRecords
''crxRpt.PrinterName = "Bullzip PDF Printer"
'crxRpt.PrintOut False, NoCopies
'
'crxApp.LogOnServer "P2SSQL.DLL", "HPSRV\SQLEXPRESS", "mercury", "sa", "sa"
'crxRpt.Database = "mercury"
'Dim datcmd1 As ADODB.Command
'    ' Create a new instance of an ADO command object
'    Set datcmd1 = New ADODB.Command
'    Set datcmd1.ActiveConnection = Gdb
'    datcmd1.CommandText = "EGGTIMREP"
'    datcmd1.CommandType = adCmdText
'
'crxRpt.Database.AddADOCommand Gdb, datcmd1
'
''crxRpt.DriverName = "mercury"
''.ReadRecords
'crxApp.LogOnServer "P2SSQL.DLL", "MERCHP", "mercury", "sa", "sa"









Private Sub Form_Load()
  
  
End Sub

Private Sub Form_Paint()

   PREVIEW
   
End Sub

Private Sub Form_Resize()
   CrystalActiveXReportViewer1.Width = Me.Width
   CrystalActiveXReportViewer1.Height = Me.Height
End Sub
