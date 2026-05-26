VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form preview 
   Caption         =   "Form1"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15195
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   15195
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox CRViewer91 
      Height          =   8655
      Left            =   240
      ScaleHeight     =   8595
      ScaleWidth      =   14835
      TabIndex        =   5
      Top             =   600
      Width           =   14895
   End
   Begin VB.CommandButton cmdPDF 
      Caption         =   "PDF"
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      Top             =   0
      Width           =   990
   End
   Begin VB.CommandButton cmd≈ ‘’–Ÿ”« 
      Caption         =   "≈ ‘’–Ÿ”«"
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   360
      Left            =   2880
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   990
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   11640
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EMAIL"
      Height          =   372
      Left            =   5640
      TabIndex        =   0
      Top             =   0
      Width           =   1212
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   7440
      TabIndex        =   4
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim CRXApplication As New CRAXDRT.Application
'Dim CRXReport As CRAXDRT.Report

Dim mEMAIL As String
Dim fRep_File As String



Private Sub cmdPDF_Click()
   ' ExportReportToPDF CRXReport, "c:\mercvb\timologio.pdf", "‘…ÃœÀœ√…œ"
End Sub

'Private Sub ExportReportToPDF(ReportObject As CRAXDRT.Report, ByVal FILENAME As String, ByVal ReportTitle As String)
'
'    Dim objExportOptions As CRAXDRT.ExportOptions
'
'    ReportObject.ReportTitle = ReportTitle
'
'    With ReportObject
'        .EnableParameterPrompting = False
'        .MorePrintEngineErrorMessages = True
'    End With
'
'    Set objExportOptions = ReportObject.ExportOptions
'
'    With objExportOptions
'        .DestinationType = crEDTDiskFile
'        .DiskFileName = FILENAME
'        .FormatType = crEFTPortableDocFormat
'        .PDFExportAllPages = True
'    End With
'
'    ReportObject.Export False
 
'End Sub












Private Sub cmd≈ ‘’–Ÿ”«_Click()

'ok typonei
'CRXReport.PrintOut

End Sub

Private Sub Command1_Click()
 ' ExportReportToPDF CRXReport, "c:\mercvb\timologio.pdf", "‘…ÃœÀœ√…œ"
  Main_Form.ListBox1.AddItem "c:\mercvb\timologio.pdf"
  Main_Form.emailParal = mEMAIL
  Main_Form.SHOW 1
  
End Sub

Sub pelates()

'Dim Connection As New ADODB.Connection
'Dim RS As New ADODB.Recordset
'Connection.Open "DSN=magazi;uid=sa;pwd=12345678"
''Obtain a Recordset object from the Customers table of the Xtreme database.
'RS.Open "select * from PEL WHERE ID=11680 ", _
'Connection, adOpenDynamic, adLockPessimistic, adCmdText
'Dim CRXApplication As New CRAXDRT.Application
'Dim CRXReport As CRAXDRT.REPORT
''Obtain a Report object by opening the report file you created earlier. This example uses the file Customer.RPT.
'Set CRXReport = CRXApplication.OpenReport("c:\MERCVB\reports\CUSTOMERS.rpt", 1)
'Dim CRXDatabase As CRAXDRT.Database
'Set CRXDatabase = CRXReport.Database
''Once you have a Database object for the Report object, you can pass the Active data source to the Report object using the SetDataSource method. This method requires three parameters. The first is the data source itself. The second parameter is a value indicating that the data source you are passing to the report is an ActiveX data source. This value must be 3. The third parameter is the table you are passing the data source to. Since you should only have one table defining the structure of the recordset, this should always be 1. For example:
'CRXDatabase.SetDataSource RS, 3, 1
'CRXReport.SQLQueryString = "select * from PEL WHERE ID=11680   "
'CRViewer91.ReportSource = CRXReport
'CRViewer91.ViewReport
''ok typonei
'CRXReport.PrintOut

End Sub
'
 'SELECT "HME", "APA", "ATIM", "PEL"."EPO", "PEL"."DIE", "PEL"."EPA", "PEL"."TYP", "KOD", "XREOSI", "PISTOSI", "ID", "AIT"
 'FROM   "MERCURY"."dbo"."EGG" "EGG" INNER JOIN "MERCURY"."dbo"."PEL" "PEL" ON ("EIDOS"="PEL"."EIDOS") AND ("KOD"="PEL"."KOD")
Private Sub Command2_Click()
   '  showtim fRep_File
End Sub


Public Sub showtim(ByVal REP_FILE As String)
''=============================================================
'If Len(Dir(REP_FILE, vbNormal)) = 0 Then
''   MsgBox "‰ÂÌ ı‹Ò˜ÂÈ ÙÔ ·Ò˜ÂﬂÔ " + REP_FILE
''   Exit Sub
'End If
'
'
'
'
'Dim Connection As New ADODB.Connection
'Dim rs As New ADODB.Recordset
''Connection.Open "DSN=mercsql"  '
'
'Connection.Open gConnect ' "DSN=magazi;uid=sa;pwd=12345678"
''Obtain a Recordset object from the Customers table of the Xtreme database.
'Dim sql As String
'
'sql = "select * from TIMOLOGIA WHERE ID_NUM=" + Label1.Caption
''sql = "SELECT HME, APA, ATIM, PEL.EPO, PEL.DIE, PEL.EPA, PEL.TYP, EGG.KOD, XREOSI, PISTOSI, EGG.ID, AIT  "
''sql = sql + "FROM   EGG INNER JOIN PEL ON (EGG.EIDOS=PEL.EIDOS) AND (EGG.KOD=PEL.KOD)" ' WHERE EGG.ID=93630 "
'
'rs.Open sql, Connection, adOpenDynamic, adLockPessimistic, adCmdText
'
'If IsNull(rs!eMail) Then
'   mEMAIL = ""
'   Else
'   mEMAIL = rs!eMail
'   End If
''Dim CRXApplication As New CRAXDRT.Application
''Dim CRXReport As CRAXDRT.Report
''Obtain a Report object by opening the report file you created earlier. This example uses the file Customer.RPT.
'Set CRXReport = CRXApplication.OpenReport(REP_FILE, 1)
'Dim CRXDatabase As CRAXDRT.Database
'Set CRXDatabase = CRXReport.Database
''Once you have a Database object for the Report object, you can pass the Active data source to the Report object using the SetDataSource method. This method requires three parameters. The first is the data source itself. The second parameter is a value indicating that the data source you are passing to the report is an ActiveX data source. This value must be 3. The third parameter is the table you are passing the data source to. Since you should only have one table defining the structure of the recordset, this should always be 1. For example:
'CRXDatabase.SetDataSource rs, 3, 1
'CRXReport.SQLQueryString = sql
''CRXReport.RecordSelectionFormula = " {TIMOLOGIA.ATIM}='T000181' "
'
'
'CRViewer91.ReportSource = CRXReport
'CRViewer91.ViewReport
'

End Sub

Private Sub Form_Load()
  fRep_File = "c:\MERCVB\reports\TIMOL519.rpt"
End Sub

Private Sub Form_Paint()

  
'  showtim fRep_File
End Sub

