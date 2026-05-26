VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form bohu6 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   7680
   ClientLeft      =   225
   ClientTop       =   765
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   13245
   WindowState     =   2  'Maximized
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "bohu6XORIS.frx":0000
      Height          =   3165
      Left            =   8115
      OleObjectBlob   =   "bohu6XORIS.frx":0014
      TabIndex        =   25
      Top             =   195
      Visible         =   0   'False
      Width           =   6780
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7740
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7335
      Width           =   2220
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Πολλά barc σε ένα κωδικό"
      Height          =   330
      Left            =   4635
      TabIndex        =   24
      Top             =   3450
      Width           =   1365
   End
   Begin VB.CheckBox Check1 
      Caption         =   "EXCEL"
      Height          =   195
      Left            =   7605
      TabIndex        =   23
      Top             =   4905
      Width           =   975
   End
   Begin VB.CheckBox ODBC 
      Caption         =   "ODBC"
      Height          =   255
      Left            =   12735
      TabIndex        =   22
      Top             =   4905
      Width           =   1440
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Μεταφορά αρχείου ΕΙΔΩΝ σε νέα χρονιά"
      Height          =   420
      Left            =   195
      TabIndex        =   21
      ToolTipText     =   "Πήγαινε πρώτα παρουσίαση ειδών για να υπολογίσει τα τελευταία υπόλοιπα"
      Top             =   3720
      Width           =   2625
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Εκτύπωση"
      Height          =   255
      Left            =   6105
      TabIndex        =   20
      Top             =   4875
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3360
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Αποθήκευση"
      Height          =   255
      Left            =   3360
      TabIndex        =   19
      Top             =   3360
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   3360
      TabIndex        =   18
      Top             =   1800
      Width           =   1695
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Bindings        =   "bohu6XORIS.frx":09E7
      Height          =   3105
      Left            =   5910
      TabIndex        =   17
      Top             =   195
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   5477
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DataView        =   2
      GroupByCaption  =   "Σύρετε τον τίτλο της στήλης για ομαδοποίηση"
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton APOTSETIM 
      Caption         =   "ενημερωση ΤΙΜ με αποθήκη"
      Height          =   390
      Left            =   2055
      TabIndex        =   16
      Top             =   6165
      Width           =   1890
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Εξοδος"
      Height          =   615
      Left            =   12735
      TabIndex        =   15
      Top             =   6210
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Mεταφορά από sql server σε DBF αρχεία"
      Height          =   615
      Left            =   4440
      TabIndex        =   14
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Μεταφορά αρχείου ΕΙΔΩΝ σε αρχείο C:\EID.DBF"
      Height          =   975
      Left            =   3360
      TabIndex        =   13
      ToolTipText     =   "Πήγαινε πρώτα παρουσίαση ειδών για να υπολογίσει τα τελευταία υπόλοιπα"
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "μεταφορά σε FALCON"
      Height          =   495
      Left            =   4440
      TabIndex        =   12
      ToolTipText     =   "Πήγαινε πρώτα παρουσίαση ειδών για να υπολογίσει τα τελευταία υπόλοιπα"
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Μεταφορά F5 Aποθήκης"
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ΜΕΤΑΦΟΡΑ F10"
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   4440
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   6600
      Visible         =   0   'False
      Width           =   10815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Μεταφορά Αρχ.Υπολ. σε κίνηση απογραφής"
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXECUTE SQL"
      Height          =   495
      Left            =   12720
      TabIndex        =   7
      Top             =   5190
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SQL"
      Height          =   495
      Left            =   6135
      TabIndex        =   6
      Top             =   5175
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   1455
      Left            =   6075
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3360
      Width           =   8100
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Text            =   "C:\LAGEURO"
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "dBASE III;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Μηδενισμός αρχείων SERVER και μεταφορά  από dbf σε SERVER"
      Height          =   735
      Left            =   2040
      TabIndex        =   2
      Top             =   4200
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
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
      Caption         =   "SQL SERVER"
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
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   840
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   6375
      Top             =   6240
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
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
      Caption         =   "SQL SERVER"
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
   Begin VB.Label Label4 
      Caption         =   "Θέση dbf αρχείων"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "bohu6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F_xre_APOU
Dim F_axre_APOU
Dim F_pis_APOU
Dim F_apis_APOU
Dim col As TrueOleDBGrid80.Column

Dim cols As TrueOleDBGrid80.Columns

 

Private Sub cmdClearFilter_Click()

    'Clears filter from grid

    For Each col In TDBGrid.Columns

        col.FilterText = ""

    Next col

    Adodc2.Recordset.Filter = adFilterNone

End Sub

 

Private Sub Command11_Click()
  CD1.InitDir = "c:\mercvb\queries"
  
  CD1.ShowSave
  
  Dim f
  f = CD1.FILENAME
  
Open f For Output As #5
    Print #5, Text2.Text
Close #5
  
  
  '+ ".txt"
  
  
  
  
  
End Sub

Private Sub Command12_Click()

' With TDBGrid.PrintInfo
'        ' Set the page header
'      .PageHeaderFont.Italic = True
'        .PageHeader = "Composers table"
'
'         'Column headers will be on every page
'        .RepeatColumnHeaders = True
'
'        ' Display page numbers (centered)
'        .PageFooter = "\tσελίδα: \p"
'        ' Invoke Print Preview
'        .PrintPreview
'  End With
If Check1 Then
   Print4_Excel UCase(Text2.Text), "11111111" 'sum_str
Else
  print3_xar UCase(Text2.Text), "11111111" 'sum_str
End If

End Sub

Private Sub Command13_Click()
' METAFORA YPOLOIPOY EIDON
Dim arxeio
Dim a As Long
Dim DBF As Database
Dim sql As New ADODB.Connection
Dim rSQLEGGT As New ADODB.Recordset
Dim rSQL As New ADODB.Recordset

Dim conSQL As String
Dim K As Long
Dim Fname As String
Dim hm As String

hm = InputBox("δωσε ημερομηνία απογραφής", , "01/01/2006")
If IsDate(hm) Then
  'ok
Else
   MsgBox "λαθος ημερομηνία"
End If


a = GetTickCount()
Me.Caption = "ENHMEΡΩΣΗ ΜΕ ΤΕΛΕΥΤΑΙΑ ΥΠΟΛΟΙΠΑ"
 UPDATE_YPOLOIPA2
arxeio = "EID"
sql.Open gConnect
On Error GoTo 0
rSQL.Open "SELECT * FROM EID WHERE POS>0 ", Gdb, adOpenForwardOnly, adLockReadOnly
rSQLEGGT.Open "SELECT * FROM EGGTIM ", Gdb, adOpenDynamic, adLockBatchOptimistic

rSQL.MoveFirst
Dim Z
Z = 0
On Error GoTo WRITEERROR ' Resume Next

Do While Not rSQL.EOF

 If Not IsNull(rSQL("POS01")) Then
    If rSQL("POS01") > 0 Then
     rSQLEGGT.AddNew
     rSQLEGGT("hme") = CDate(hm)
     rSQLEGGT("ATIM") = "λ00001"
     rSQLEGGT("KODE") = rSQL("KOD")
     rSQLEGGT("ONOMA") = rSQL("ONO")
     rSQLEGGT("TIMM") = rSQL("XTI")
     rSQLEGGT("FPA") = rSQL("FPA")
     rSQLEGGT("POS0") = rSQL("POS01")
     rSQLEGGT("APOT") = 1
     rSQLEGGT.Update
    End If
  End If
  If Not IsNull(rSQL("POS02")) Then
    If rSQL("POS02") > 0 Then
     rSQLEGGT.AddNew
     rSQLEGGT("hme") = CDate(hm)
     rSQLEGGT("ATIM") = "λ00001"
     rSQLEGGT("KODE") = rSQL("KOD")
     rSQLEGGT("ONOMA") = rSQL("ONO")
     rSQLEGGT("TIMM") = rSQL("XTI")
     rSQLEGGT("FPA") = rSQL("FPA")
     rSQLEGGT("POS0") = rSQL("POS01")
     rSQLEGGT("APOT") = 2
     rSQLEGGT.Update
    End If
  End If
    Z = Z + 1
    If Z Mod 10 = 0 Then
      Me.Caption = Z
      DoEvents
    End If
    rSQL.MoveNext
Loop

rSQL.Close
rSQLEGGT.Close

Me.Caption = (GetTickCount() - a) / 1000




Exit Sub


WRITEERROR:
List2.Visible = True

List2.AddItem Err.Description
Resume Next

End Sub

Private Sub File1_Click()
  Dim a
  a = 0

Dim f As String

f = File1.FILENAME


Dim ss As String
Dim b As String

ss = ""
Open File1.Path + "\" + f For Input As #1
Do While Not EOF(1)
    
    Line Input #1, b
    ss = ss + b + " "
Loop
Close #1


Text2.Text = ss



End Sub

Private Sub List1_DblClick()
 Dim K As Integer
   For K = 0 To List1.ListCount - 1
     List1.Selected(K) = True
   Next
End Sub

Private Sub TDBGrid_FilterChange()

'Gets called when an action is performed on the filter bar

On Error GoTo ErrHandler

 

Set cols = TDBGrid.Columns

Dim c As Integer

c = TDBGrid.col

TDBGrid.HoldFields

Adodc2.Recordset.Filter = getFilter()

TDBGrid.col = c

TDBGrid.EditActive = True

Exit Sub

 

ErrHandler:

    MsgBox Err.Source & ":" & vbCrLf & Err.Description

    Call cmdClearFilter_Click

End Sub

 

Private Function getFilter() As String
    'Creates the SQL statement in adodc1.recordset.filter
    'and only filters text currently. It must be modified to filter other data types.
    
    Dim tmp As String
    Dim n As Integer
    For Each col In cols
        If Trim(col.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If
           
           
          If Adodc2.Recordset(col.ColIndex).Type = 5 Then '
            If InStr(">< >= <=  = ", left(col.FilterText, 1)) > 0 And Len(col.FilterText) > 1 And IsNumeric(Right(col.FilterText, 1)) Then
                tmp = tmp & col.DataField & col.FilterText
            Else
             tmp = tmp & col.DataField & " >  " & Val(col.FilterText)
            End If
            
          Else
             tmp = tmp & col.DataField & " LIKE '" & col.FilterText & "*'"
          End If
            
        End If
    Next col
                
    getFilter = tmp
End Function


Private Sub APOTSETIM_Click()

Gdb.Execute "CREATE VIEW dbo.[UPD_TIM] as " _
& " SELECT    dbo.TIM.B_N1 , dbo.EGGTIM.APOT " _
& "  FROM         dbo.TIM LEFT JOIN" _
& "  dbo.EGGTIM ON dbo.TIM.HME = dbo.EGGTIM.HME AND dbo.TIM.ATIM = dbo.EGGTIM.ATIM"





Gdb.Execute "UPDATE UPD_TIM  SET B_N1=APOT"

End Sub

Private Sub Command1_Click()
Dim DBF As New ADODB.Connection
Dim sql As New ADODB.Connection
Dim rDBF As New ADODB.Recordset
Dim rSQL As New ADODB.Recordset



Dim conDBF As String
Dim conSQL As String
On Error Resume Next

Me.MousePointer = vbHourglass

If ODBC.Value = vbChecked Then
   DBGrid1.Visible = True
   Data2.Connect = "ODBC;" + gConnect
   Data2.RecordSource = Text2.Text
   Data2.Refresh
Else
   DBGrid1.Visible = False
   Adodc2.ConnectionString = gConnect
   Adodc2.RecordSource = Text2.Text
   Adodc2.Refresh
End If


TDBGrid.AlternatingRowStyle = True

TDBGrid.OddRowStyle.BackColor = &H8000000F   ' GRI   vbCyan
TDBGrid.EvenRowStyle.BackColor = &HFFFFC0   'OYRANI     &H8000000F  ' GRI


Me.MousePointer = vbNormal

End Sub

Private Sub update_sql_from_dbf(arxeio As String)
Dim DBF As Database
Dim sql As New ADODB.Connection
Dim rDBF As Recordset
Dim rSQL As New ADODB.Recordset
Dim conDBF As String
Dim conSQL As String
Dim K As Long
Dim Fname As String
G2 = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP4;Data Source=PLATO"
'2004 G2 = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP4;Data Source=PLATO"
conSQL = gConnect ' "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP;Data Source=LAGAKIS"
conDBF = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dBASE Files;Initial Catalog=C:\LAGEURO"
'2004 conDBF = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dBASE Files;Initial Catalog=C:\LAGEURO\04"

sql.Open conSQL
'SET DBF=DATA.Open conDBF

Data1.DatabaseName = Text1.Text

Data1.RecordSource = arxeio
Data1.Refresh

Set rDBF = Data1.Recordset

Label1.Caption = arxeio + " " + Format(rDBF.RecordCount, "######")
sql.Execute "DELETE FROM " + arxeio

rSQL.Open "SELECT *FROM " + arxeio, sql, adOpenDynamic, adLockOptimistic
' rDBF.Open "SELECT *FROM " + arxeio, DBF, adOpenDynamic, adLockOptimistic

rDBF.MoveFirst

Dim Z
Z = 0
On Error GoTo WRITEERROR ' Resume Next

Do While Not rDBF.EOF
  
  rSQL.AddNew
  
  For K = 0 To rDBF.FIELDS.Count - 1
      Fname = rDBF.FIELDS(K).Name ' p.x. FNAME=epo    R(0).NAME
      
      If IsNull(rDBF(K)) Then
      Else
         rSQL(Fname) = rDBF(K) ' rsqk("epo")=r(0)
      End If
     ' ΑΥΤΑ ΠΟΥ ΕΙΝΑΙ ΝULL ΤΑ ΓΕΜΙΖΕΙ
     If IsNull(rDBF(K)) Then
        If rDBF(K).Type = dbText Then '10=STRING
            rSQL(Fname) = " "
        ElseIf rDBF(K).Type = dbDate Then
            rSQL(Fname) = CDate("01/01/1990")
        ElseIf rDBF(K).Type = dbBoolean Then
            rSQL(Fname) = False
        ElseIf rDBF(K).Type = dbDouble Then
            rSQL(Fname) = 0
        Else
             conSQL = ""
        End If
     End If
  Next
  
  
  ' ΜΕΤΑΤΡΕΠΕΙ ΤΗΝ "ΧRΕOSI" ΣΕ "XRE" KAI "PISTOSI"
  If arxeio = "EGG" Then
     rSQL("PISTOSI") = 0: rSQL("XREOSI") = 0
     If Not IsNull(rSQL("XRE")) Then
       If rSQL("EIDOS") = "e" Then
          If rSQL("XRE") > 0 Then rSQL("XREOSI") = rSQL("XRE")
          If rSQL("XRE") < 0 Then rSQL("PISTOSI") = -rSQL("XRE")
       Else
          If rSQL("XRE") > 0 Then rSQL("PISTOSI") = rSQL("XRE")
          If rSQL("XRE") < 0 Then rSQL("XREOSI") = -rSQL("XRE")
       End If
     End If
  End If


  
  ' METΑΤΡΕΠΕΙ ΤΟ 'POSO' ΣΕ 'XRE' KAI 'PIS'
  If arxeio = "EGGTIM" Then
     rSQL("PIS") = 0: rSQL("XRE") = 0
     If Not IsNull(rSQL("POSO")) Then
        If InStr(F_xre_APOU + F_axre_APOU, left(rSQL("ATIM"), 1)) > 0 Then
            rSQL("XRE") = rSQL("POSO")
        End If
        If InStr(F_pis_APOU + F_apis_APOU, left(rSQL("ATIM"), 1)) > 0 Then
            rSQL("PIS") = rSQL("POSO")
        End If
     End If
  End If
  
  
  
  
  
  rSQL.Update
    Z = Z + 1
    If Z Mod 10 = 0 Then
      Me.Caption = Z
    End If
  rDBF.MoveNext
  DoEvents
  
Loop

rSQL.Close
rDBF.Close
Exit Sub


WRITEERROR:
List2.Visible = True

List2.AddItem Err.Description
Resume Next






End Sub

Private Sub Command10_Click()
   Unload Me


End Sub

Private Sub Command2_Click()
Dim K As Integer
Dim R As New ADODB.Recordset
ANS = MsgBox("Προσοχή θα σβηστούν τα αρχεία του SERVER. Είσαι σίγουρος;", vbYesNo)
If ANS = vbNo Then Exit Sub
Dim mkod As String

If Not ANS = vbYes Then
   MsgBox " ΑΚΥΡΏΘΗΚΕ Η ΕΡΓΑΣΊΑ"
   Exit Sub
End If

  mkod = InputBox("Αν είσαι σίγουρος δώσε κωδικό ")
If Val(mkod) = 12345 Then
'ok
Else
  Exit Sub
End If



R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic


F_xre_APOU = ""
F_axre_APOU = ""
F_pis_APOU = ""
F_apis_APOU = ""
  
R.MoveFirst
  Do While Not R.EOF
     If R("POS_APOU") = 1 Then
        F_xre_APOU = F_xre_APOU + R("eidos")
     ElseIf R("POS_APOU") = 2 Then
        F_axre_APOU = F_axre_APOU + R("eidos")
     ElseIf R("POS_APOU") = 3 Then
        F_pis_APOU = F_pis_APOU + R("eidos")
     ElseIf R("POS_APOU") = 4 Then
        F_apis_APOU = F_apis_APOU + R("eidos")
     End If
       
       R.MoveNext
  Loop
R.Close


For K = 0 To List1.ListCount - 1
 If List1.Selected(K) Then
    update_sql_from_dbf List1.List(K)
 End If
 
Next

End Sub

Private Sub Command3_Click()
Dim DBF As New ADODB.Connection
Dim sql As New ADODB.Connection
Dim rDBF As New ADODB.Recordset
Dim rSQL As New ADODB.Recordset



Dim conDBF As String
Dim conSQL As String
Dim LO As Long
Dim db As Database



LO = 0
Me.MousePointer = vbHourglass
   On Error GoTo latos  'On Error Resume Next


If ODBC.Value = Checked Then

   Set db = OpenDatabase("", False, False, gConnect)
   db.Execute Text2.Text
   LO = db.RecordsAffected
Else
   Gdb.Execute Text2.Text, LO
   
End If
   MsgBox Str(LO) + " ΕΓΓΡΑΦΕΣ ΕΝΗΜΕΡΩΘΗΚΑΝ"
   Me.MousePointer = vbNormal


Exit Sub

latos:
MsgBox Err.Description
Resume Next

End Sub

Private Sub Command4_Click()
Dim sql As New ADODB.Connection
Dim rDBF As New ADODB.Recordset
Dim rSQL As New ADODB.Recordset

rSQL.Open "SELECT * FROM EID WHERE APOS>0", Gdb, adOpenDynamic, adLockOptimistic




a = InputBox("Δώσε κωδικό παραστατικού", , "λ")

Gdb.Execute "DELETE  FROM EGGTIM WHERE LEFT(ATIM,1)='" + a + "';"

Dim K As Long
rSQL.MoveFirst
K = 0
Do While Not rSQL.EOF
   If Not IsNull(rSQL("APOS")) Then
      K = K + 1
      Gdb.Execute "INSERT INTO EGGTIM (EIDOS,PELKOD,HME,XRE,KODE,ATIM) VALUES ('e','0000','01/01/2005'," + Str(rSQL("APOS")) + ",'" + rSQL("KOD") + "','" + a + Format(K, "00000") + "')"
   End If
   rSQL.MoveNext
   If K Mod 10 = 0 Then DoEvents: Me.Caption = Str(K)
Loop
Gdb.Execute "UPDATE EID SET APOS=0"
End Sub

Private Sub Command5_Click()
Dim DBF As Database
Dim sql As New ADODB.Connection
Dim rDBF As Recordset
Dim rSQL As New ADODB.Recordset
Dim conDBF As String
Dim conSQL As String
Dim LO As Long
Dim r2 As DAO.Recordset

Set DBF = OpenDatabase(Text1.Text, False, False, "dBase III;")
   Set r2 = DBF.OpenRecordset("select LEFT(memo_key,5) AS KLEIDI,COUNT(*) from memotext where left(memo_key,1) in ('e','r') group by  LEFT(memo_key,5) order by  LEFT(memo_key,5) ")

rSQL.Open "select *from PEL", Gdb, adOpenDynamic, adLockOptimistic


Dim a, b, c, d, e, f, G

r2.MoveFirst
Do While Not r2.EOF
   a = left(r2(0), 5)
   Set R = DBF.OpenRecordset("select *from memotext where left(memo_key,5)='" + a + "' order by memo_key ")
   If R.RecordCount > 0 Then
      Me.Caption = a
      DoEvents
      R.MoveFirst
      b = ""
      Do While Not R.EOF
         b = b + R(1) + Chr$(13) + Chr$(10) 'memoline
         R.MoveNext
      Loop
      
      If Not IsNull(b) Then
         b = Replace(b, "'", "-")
         If Not IsNull(b) Then Gdb.Execute "update PEL SET MEMO='" + b + "' WHERE EIDOS='" + left(a, 1) + "' AND KOD='" + Mid$(a, 2, 4) + "'"
      End If
      DoEvents
   End If
   r2.MoveNext
 Loop








End Sub

Private Sub Command6_Click()
'F5 ΑΠΟΘΗΚΗς
Dim DBF As Database
Dim sql As New ADODB.Connection
Dim rDBF As Recordset
Dim rSQL As New ADODB.Recordset
Dim conDBF As String
Dim conSQL As String
Dim LO As Long

Set DBF = OpenDatabase(Text1.Text, False, False, "dBase III;")


rSQL.Open "select *from EID", Gdb, adOpenDynamic, adLockOptimistic


Dim a, b, c, d, e, f, G
G = 0
rSQL.MoveFirst
Do While Not rSQL.EOF
   G = G + 1
   a = "a" + rSQL!kod
   
   Set R = DBF.OpenRecordset("select *from memotext where left(memo_key,15)='" + a + "' order by memo_key ")
   If R.RecordCount > 0 Then
      
      R.MoveFirst
      b = ""
      Do While Not R.EOF
         b = b + R(1) + Chr$(13) 'memoline
         R.MoveNext
      Loop
      rSQL!Memo = b
      rSQL.Update
      DoEvents
   End If
   G = G + 1
   If G Mod 20 = 0 Then
      Me.Caption = a
      DoEvents
   End If
   rSQL.MoveNext
 Loop

End Sub

Private Sub Command7_Click()
'------ FALCON ----------

Dim sql As New ADODB.Connection
Dim rDBF As New ADODB.Recordset
Dim rSQL As New ADODB.Recordset
If Check2.Value = vbChecked Then
   rSQL.Open "SELECT BARCODES.*,EID.KOD,EID.ONO,EID.LTI5,EID.POS01,EID.POS02," _
   & " EID.PROM,EID.XTI,EID.KODLOG,EID.AEG" _
   & " FROM BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD" _
   & "  ", Gdb, adOpenDynamic, adLockOptimistic
Else
   rSQL.Open "SELECT * FROM EID where LEFT(KOD,5) NOT IN ('91393')", Gdb, adOpenDynamic, adLockOptimistic
End If

Dim db As Database, R As Recordset

Set db = OpenDatabase("C:\LAGEURO\MHXANAKI", False, False, "dBase III;")

On Error Resume Next

Kill "C:\LAGEURO\MHXANAKI\OLA.DBF"

On Error GoTo 0
db.Execute "SELECT TOP 1 OLAA.* INTO OLA FROM OLAA"
Set R = db.OpenRecordset("OLA")

On Error Resume Next


Dim K As Long
rSQL.MoveFirst
K = 0
Do While Not rSQL.EOF
    R.AddNew
    
If Not Check2.Value = vbChecked Then
    
    
    
    
    R("BARCODE") = left(rSQL("KOD"), 13)
    R("LTI") = rSQL("LTI5")
    R("NAME") = left(rSQL("ONO"), 25)
    R("POS01") = rSQL("POS01")
    R("POS02") = rSQL("POS02")
    R("PROM") = left(rSQL("PROM"), R("PROM").Size)
    R("XTI") = rSQL("XTI")
    R("KODLOG") = rSQL("KODLOG")
Else

    R("BARCODE") = left(rSQL("ERG"), 13)
    R("LTI") = rSQL("LTI5")
    R("XTI") = rSQL("XTI")
    R("CODE") = rSQL("KOD")
    R("NAME") = left(rSQL("ONO"), 25)
    R("POS01") = rSQL("POS01")
    R("POS02") = rSQL("POS02")
    R("PROM") = left(rSQL("PROM"), R("PROM").Size)
    R("XTI") = rSQL("XTI")
    R("KODLOG") = rSQL("KODLOG")



End If
    
    R.Update
    K = K + 1
    If K Mod 100 = 0 Then
     Me.Caption = K
     DoEvents
    End If
   
   'Me.Caption = rSQL("kod")
   
   rSQL.MoveNext

Loop
Gdb.Execute "UPDATE EID SET APOS=0"

'FileCopy "c:\lageuro\olaa.dbf", "c:\lageuro\mhxanaki\ola.dbf"
Dim DUM

'            C:\lageuro\MHXANAKI\SEND.BAT
On Error Resume Next

DUM = Shell("C:\LAGEURO\MHXANAKI\send.bat", vbHide)




'dum = Shell("C:\LAGEURO\MHXANAKI\pkzip c:\lageuro\mhxanaki\o c:\lageuro\mhxanaki\ola.dbf", vbHide)


'dum = Shell("C:\LAGEURO\pkzip  c:\lageuro\mhxanaki\o c:\lageuro\mhxanaki\ola.dbf", vbMaximizedFocus)








End Sub

Private Sub Command8_Click()
' METAFORA SE ARXEIO DBF
Dim arxeio
Dim a As Long
a = GetTickCount()
Me.Caption = "ENHMEΡΩΣΗ ΜΕ ΤΕΛΕΥΤΑΙΑ ΥΠΟΛΟΙΠΑ"

UPDATE_YPOLOIPA2

arxeio = "EID"



Dim DBF As Database
Dim sql As New ADODB.Connection
Dim rDBF As Recordset
Dim rSQL As New ADODB.Recordset

Dim conSQL As String
Dim K As Long
Dim Fname As String





Dim db As DAO.Database
'Set DB = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
On Error Resume Next
Kill "C:\LAGEURO\EID.DBF"
Kill "C:\LAGEURO\BARCODES.DBF"
On Error GoTo 0

Dim LO As Long, LO2 As Long

Me.Caption = "ΔΗΜΙΟΥΡΓΙΑ ΑΡΧΕΙΟΥ ΕΙΔΩΝ"
   Set db = OpenDatabase("", False, False, gConnect)
   db.Execute "SELECT KOD,ONO,ERG,LTI,LTI5,PONTOI,FPA,POS01," _
   & " POS02,PROM,XTI,KODLOG,POS_KERD,POS_KERD2,SPA,MON" _
   & " INTO EID IN 'C:\LAGEURO' 'dBase III;' FROM EID"
   LO = db.RecordsAffected

Me.Caption = "ΔΗΜΙΟΥΡΓΙΑ ΑΡΧΕΙΟΥ BARCODE"
db.Execute "SELECT * " _
   & " INTO BARCODES IN 'C:\LAGEURO' 'dBase III;' FROM BARCODES "
   LO2 = db.RecordsAffected





   MsgBox "Εστάλησαν " + Str(LO) + " είδη και " + Str(LO2) + " barcodes "
   Exit Sub
   




sql.Open gConnect
db.Execute "SELECT TOP 1 EID0.* INTO EID  FROM EID0"


Set rDBF = db.OpenRecordset("EID")





'Label1.Caption = ARXEIO + " " + Format(rDBF.RecordCount, "######")

'SQL.Execute "DELETE FROM " + arxeio

rSQL.Open "SELECT * FROM EID ", Gdb, adOpenForwardOnly, adLockReadOnly

' rDBF.Open "SELECT *FROM " + arxeio, DBF, adOpenDynamic, adLockOptimistic

rSQL.MoveFirst

Dim Z
Z = 0
On Error GoTo WRITEERROR ' Resume Next

Do While Not rSQL.EOF
  
  rDBF.AddNew
     rDBF("KOD") = rSQL("KOD")
     rDBF("ONO") = rSQL("ONO")
     rDBF("ERG") = rSQL("ERG")
     rDBF("LTI") = rSQL("LTI")
     rDBF("LTI5") = rSQL("LTI5")
     rDBF("pontoi") = rSQL("pontoi")
     rDBF("FPA") = rSQL("FPA")
     rDBF("POS01") = rSQL("POS01")
     rDBF("POS02") = rSQL("POS02")
     
     rDBF("prom") = rSQL("prom")
     rDBF("xti") = rSQL("xti")
     rDBF("KODLOG") = rSQL("KODLOG")
     
     
     
     rDBF("pOS_KERD") = rSQL("pOS_KERD")
     rDBF("POS_KERD2") = rSQL("POS_KERD2")
     rDBF("SPA") = rSQL("SPA")
     rDBF("MON") = rSQL("MON")
     
     
     

     
     
     
     
     
  rDBF.Update
  
    Z = Z + 1
    If Z Mod 10 = 0 Then
      Me.Caption = Z
    End If
  rSQL.MoveNext
  DoEvents
  
Loop

rSQL.Close
rDBF.Close

Me.Caption = (GetTickCount() - a) / 1000




Exit Sub


WRITEERROR:
List2.Visible = True

List2.AddItem Err.Description
Resume Next







End Sub

Private Sub Command9_Click()
' apo SQL SERVER ΣΕ DBF ΑΡΧΕΙΑ ΠΟΥ ΗΔΗ ΥΠΑΡΧΟΥΝ
Dim K As Integer
Dim R As New ADODB.Recordset
ANS = MsgBox("Προσοχή θα σβηστούν τα αρχεία DBF. Είσαι σίγουρος;", vbYesNo)
If ANS = vbNo Then Exit Sub
Dim mkod As Integer
mkod = InputBox("Αν είσαι σίγουρος δώσε κωδικό ")
If Val(mkod) = 12345 Then
'ok
Else
  Exit Sub
End If



R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic


F_xre_APOU = ""
F_axre_APOU = ""
F_pis_APOU = ""
F_apis_APOU = ""
  
R.MoveFirst
  Do While Not R.EOF
     If R("POS_APOU") = 1 Then
        F_xre_APOU = F_xre_APOU + R("eidos")
     ElseIf R("POS_APOU") = 2 Then
        F_axre_APOU = F_axre_APOU + R("eidos")
     ElseIf R("POS_APOU") = 3 Then
        F_pis_APOU = F_pis_APOU + R("eidos")
     ElseIf R("POS_APOU") = 4 Then
        F_apis_APOU = F_apis_APOU + R("eidos")
     End If
       
       R.MoveNext
  Loop
R.Close


For K = 0 To List1.ListCount - 1
 If List1.Selected(K) Then
    update_DBF_from_SQL List1.List(K)
 End If
 
Next

End Sub

Private Sub Form_Load()
 Me.Picture = LoadPicture(gPicture)
List1.AddItem "PEL"
List1.AddItem "EID"
List1.AddItem "EGG"
List1.AddItem "EGGTIM"
List1.AddItem "TIM"
List1.AddItem "GRA"

On Error Resume Next
File1.Path = "C:\mercvb\queries"

End Sub

Private Sub update_DBF_from_SQL(arxeio As String)

Dim DBF As Database
Dim sql As New ADODB.Connection
Dim rDBF As Recordset
Dim rSQL As New ADODB.Recordset
Dim conDBF As String
Dim conSQL As String
Dim K As Long
Dim Fname As String

Dim db As DAO.Database
Set db = OpenDatabase(Text1.Text, False, False, "dBase III;")
db.Execute "delete * from " + arxeio
Set rDBF = db.OpenRecordset(arxeio)

Label1.Caption = arxeio + " " + Format(rDBF.RecordCount, "######")



rSQL.Open "SELECT *FROM " + arxeio, sql, adOpenDynamic, adLockOptimistic
' rDBF.Open "SELECT *FROM " + arxeio, DBF, adOpenDynamic, adLockOptimistic

rSQL.MoveFirst

Dim Z
Z = 0
On Error GoTo WRITEERROR ' Resume Next

Do While Not rSQL.EOF
  
  rDBF.AddNew
  
  For K = 0 To rDBF.FIELDS.Count - 1
      Fname = rDBF.FIELDS(K).Name ' p.x. FNAME=epo    R(0).NAME
      
'      If IsNull(rDBF(k)) Then
'      Else
         rDBF(Fname) = rSQL(Fname) ' rsqk("epo")=r(0)
'      End If
  Next
  rDBF.Update
    Z = Z + 1
    If Z Mod 10 = 0 Then
      Me.Caption = Z
    End If
  rSQL.MoveNext
  DoEvents
  
Loop

rSQL.Close
rDBF.Close
Exit Sub


WRITEERROR:
List2.Visible = True

List2.AddItem Err.Description
Resume Next




End Sub

Private Sub TDBGrid_GroupColMove(ByVal Position As Integer, ByVal ColIndex As Integer, Cancel As Integer)
   Dim strSort As String

    Dim col As TrueOleDBGrid80.Column

 

' Loop through GroupColumns collection and construct

' the sort string for the Sort property of the Recordset

    For Each col In TDBGrid.GroupColumns

        If strSort <> vbNullString Then

            strSort = strSort & ", "

        End If

        strSort = strSort & "[" & col.DataField & "]"

    Next col

 

    TDBGrid.HoldFields

    Adodc2.Recordset.Sort = strSort


End Sub
