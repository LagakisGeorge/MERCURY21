VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form bohu6 
   BackColor       =   &H00FF0000&
   Caption         =   "Εξαγωγή Αρχείων"
   ClientHeight    =   9135
   ClientLeft      =   225
   ClientTop       =   765
   ClientWidth     =   15555
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   15555
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6000
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\mercvb\reports\timol1.rpt"
   End
   Begin VB.TextBox Text2 
      Height          =   1335
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   52
      Top             =   3360
      Width           =   9495
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Τυπωνω Grid"
      Height          =   255
      Left            =   10440
      TabIndex        =   51
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CheckBox CheckReport 
      Caption         =   "Reports.mdb"
      Height          =   255
      Left            =   13680
      TabIndex        =   50
      Top             =   5280
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H000080FF&
      Height          =   540
      Left            =   1440
      TabIndex        =   49
      Top             =   120
      Width           =   4425
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0080C0FF&
      Caption         =   "METAΦOREΣ AΠO DOS"
      Height          =   3375
      Left            =   240
      TabIndex        =   35
      Top             =   3360
      Width           =   5535
      Begin VB.CommandButton cmdWebservice 
         Caption         =   "webservice"
         Height          =   360
         Left            =   3480
         TabIndex        =   53
         Top             =   2880
         Width           =   990
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1560
         TabIndex        =   48
         Text            =   "C:\LAGEURO"
         Top             =   1140
         Width           =   1710
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Μηδενισμός αρχείων SERVER και μεταφορά  από dbf σε SERVER"
         Height          =   735
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   3150
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Μεταφορά αρχείου ΕΙΔΩΝ σε νέα χρονιά"
         Height          =   420
         Left            =   120
         TabIndex        =   43
         ToolTipText     =   "Πήγαινε πρώτα παρουσίαση ειδών για να υπολογίσει τα τελευταία υπόλοιπα"
         Top             =   2160
         Width           =   3150
      End
      Begin VB.CommandButton Command5 
         Caption         =   "ΜΕΤΑΦΟΡΑ F10"
         Height          =   360
         Left            =   3360
         TabIndex        =   42
         Top             =   2430
         Width           =   2040
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Μεταφορά F5 Aποθήκης"
         Height          =   495
         Left            =   3360
         TabIndex        =   41
         Top             =   1410
         Width           =   2040
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Mεταφορά από sql server σε DBF αρχεία"
         Height          =   495
         Left            =   3360
         TabIndex        =   40
         Top             =   885
         Width           =   2040
      End
      Begin VB.CommandButton APOTSETIM 
         Caption         =   "ενημερωση ΤΙΜ με αποθήκη"
         Height          =   495
         Left            =   3360
         TabIndex        =   39
         Top             =   360
         Width           =   2040
      End
      Begin VB.CommandButton Command14 
         Caption         =   "ΦΟΡΤΩΣΗ ΦΩΤΟΓΡΑΦΙΩΝ"
         Height          =   465
         Left            =   3360
         TabIndex        =   38
         Top             =   1935
         Width           =   2040
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Μεταφορά Αρχ.Υπολ. σε κίνηση απογραφής"
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         Width           =   3150
      End
      Begin VB.Label Label4 
         Caption         =   "Θέση dbf αρχείων"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   1125
         Width           =   1395
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H000080FF&
      Height          =   2895
      Left            =   13095
      TabIndex        =   16
      Top             =   5640
      Width           =   2295
      Begin VB.TextBox STORED 
         Height          =   330
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   1860
      End
      Begin VB.TextBox PAR1 
         Height          =   330
         Left            =   240
         TabIndex        =   27
         Top             =   1440
         Width           =   1860
      End
      Begin VB.TextBox PAR2 
         Height          =   330
         Left            =   240
         TabIndex        =   26
         Top             =   1800
         Width           =   1860
      End
      Begin VB.CommandButton Command18 
         Caption         =   "STORED PROCEDURE"
         Height          =   360
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   1860
      End
      Begin VB.Label Label2 
         Caption         =   "PROCEDURE"
         Height          =   300
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Height          =   3375
      Left            =   5880
      TabIndex        =   15
      Top             =   5640
      Width           =   7215
      Begin VB.CommandButton CmdupdateStructures 
         Caption         =   "cmdupdateStructures"
         Height          =   360
         Left            =   5400
         TabIndex        =   60
         Top             =   2880
         Width           =   990
      End
      Begin VB.CommandButton cmdΤΗΛΕΦ 
         Caption         =   "ΤΗΛΕΦ"
         Height          =   360
         Left            =   2520
         TabIndex        =   59
         Top             =   2880
         Width           =   990
      End
      Begin VB.CommandButton CMPUpdateMon 
         Caption         =   "Ολα τα είδη σε ΜΟΝ μετ από λιστα"
         Height          =   360
         Left            =   3720
         TabIndex        =   58
         Top             =   720
         Width           =   2895
      End
      Begin VB.CommandButton Command27 
         Caption         =   "INSERT BLANK ENTRIES IN TIM"
         Height          =   360
         Left            =   1440
         TabIndex        =   57
         Top             =   2880
         Width           =   990
      End
      Begin VB.CommandButton CrystalToPDF 
         Caption         =   "CrystalToPDF"
         Height          =   360
         Left            =   360
         TabIndex        =   56
         Top             =   2880
         Width           =   990
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Παραμετροποιηση Παρόχου"
         Height          =   255
         Left            =   3720
         TabIndex        =   55
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton SELECTTIM 
         Caption         =   "SELECTTIM"
         Height          =   360
         Left            =   285
         TabIndex        =   54
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton Dataentry 
         Caption         =   "Γρηγορη Εισαγωγή"
         Height          =   360
         Left            =   3720
         TabIndex        =   24
         Top             =   2880
         Width           =   1440
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Δημιουργία εγγραφής Calendar"
         Height          =   435
         Left            =   285
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton IOSIFIDIS 
         Caption         =   "Αποστολή Αρχείου ASCII  σε Λογιστική"
         Height          =   435
         Left            =   3720
         TabIndex        =   22
         Top             =   1335
         Width           =   2880
      End
      Begin VB.CommandButton Command20 
         Caption         =   "ΜΕΤΑΤΡΟΠΗ ΚΙΝΗΣΕΩΝ ΤΑΜΕΙΑΚΗΣ ΣΕ ΠΑΡΑΣΤΑΤΙΚΑ"
         Height          =   435
         Left            =   285
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2385
         Width           =   2895
      End
      Begin VB.CommandButton Command21 
         Caption         =   "ΕΝΗΜΕΡΩΣΗ ΑΠΟ ΑΛΛΟ ΕΜΠΟΡΙΚΟ"
         Height          =   435
         Left            =   3720
         TabIndex        =   20
         Top             =   1860
         Width           =   2880
      End
      Begin VB.CommandButton Command23 
         Caption         =   "συγχρονισμός τοπικού από απομακρυσμένο  πίνακα"
         Height          =   435
         Left            =   3720
         TabIndex        =   19
         Top             =   2385
         Width           =   2880
      End
      Begin VB.CommandButton Command24 
         Caption         =   "ΜΕΤΑΤΡΟΠΗ ΣΕ 13 ΨΗΦΙΑ"
         Height          =   435
         Left            =   285
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1845
         Width           =   2895
      End
      Begin VB.CommandButton PERP 
         Caption         =   "ΠΑΓΩΤΑ ΒΟΗΘΗΤΙΚΑ ΕΝΗΜΕΡΩΣΗ EGGTIM  ΜΕ ΚΩΔΙΚΟ ΠΕΛΑΤΗ"
         Height          =   600
         Left            =   285
         TabIndex        =   17
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H80000010&
      Caption         =   "Εξοδος"
      Height          =   435
      Left            =   13095
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8580
      Width           =   2295
   End
   Begin VB.CommandButton Command22 
      Caption         =   "EXTERNAL DB"
      Height          =   255
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5280
      Width           =   3015
   End
   Begin VB.CommandButton Command16 
      Caption         =   "TRANSACTIONS SAMPLE "
      Height          =   435
      Left            =   7680
      TabIndex        =   11
      Top             =   11760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox mactext 
      Height          =   360
      Left            =   15240
      TabIndex        =   10
      Top             =   10560
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   12000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox ODBC 
      Caption         =   "ODBC"
      Height          =   240
      Left            =   12495
      TabIndex        =   9
      Top             =   5280
      Width           =   1080
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Εκτύπωση"
      Height          =   255
      Left            =   9000
      TabIndex        =   8
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Αποθήκευση"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   3045
      Width           =   4425
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H0080C0FF&
      Height          =   2235
      Left            =   1440
      TabIndex        =   6
      Top             =   600
      Width           =   4425
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Bindings        =   "bohu6.frx":0000
      Height          =   3105
      Left            =   5880
      TabIndex        =   5
      Top             =   195
      Width           =   9465
      _ExtentX        =   16695
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
      Splits(0)._SavedRecordSelectors=   -1  'True
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
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   13695
      TabIndex        =   4
      Top             =   10620
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXECUTE SQL"
      Height          =   315
      Left            =   12495
      TabIndex        =   3
      Top             =   4920
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SQL"
      Height          =   615
      Left            =   5880
      TabIndex        =   2
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "dBASE III;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   11880
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3000
      Top             =   12000
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Height          =   2205
      Left            =   270
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1800
      Top             =   11880
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "ΑΠΟΣΤΟΛΕΣ"
      Height          =   2280
      Left            =   240
      TabIndex        =   12
      Top             =   6735
      Width           =   5535
      Begin VB.CheckBox BONUS 
         BackColor       =   &H000080FF&
         Caption         =   "BONUS"
         Height          =   255
         Left            =   3240
         MaskColor       =   &H0080C0FF&
         TabIndex        =   46
         Top             =   550
         Width           =   1575
      End
      Begin VB.CheckBox TIMES 
         BackColor       =   &H000080FF&
         Caption         =   "TIMEΣ"
         Height          =   255
         Left            =   3240
         TabIndex        =   45
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command17 
         Caption         =   "ΕΝΗΜ.POS"
         Height          =   375
         Left            =   615
         TabIndex        =   44
         Top             =   1680
         Width           =   2445
      End
      Begin VB.CommandButton Command7 
         Caption         =   "μεταφορά σε FALCON"
         Height          =   420
         Left            =   615
         TabIndex        =   34
         ToolTipText     =   $"bohu6.frx":0015
         Top             =   1245
         Width           =   2445
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Μεταφορά αρχείου ΕΙΔΩΝ σε αρχείο C:\EID.DBF (ΥΠΟΚ)"
         Height          =   525
         Left            =   615
         TabIndex        =   33
         ToolTipText     =   "Πήγαινε πρώτα παρουσίαση ειδών για να υπολογίσει τα τελευταία υπόλοιπα"
         Top             =   660
         Width           =   2445
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H000080FF&
         Caption         =   "Πολλά barc σε ένα κωδικό"
         Height          =   330
         Left            =   3240
         TabIndex        =   32
         Top             =   1275
         Width           =   1365
      End
      Begin VB.CommandButton Command15 
         Caption         =   "ΦΟΡΗΤΗ ΤΙΜΟΛΟΓΗΣΗ"
         Height          =   390
         Left            =   615
         TabIndex        =   31
         ToolTipText     =   "F33 ΕΙΔΗ F31 ΠΕΛΑΤΕΣ"
         Top             =   240
         Width           =   2445
      End
      Begin VB.CheckBox XorisTimiAgoras 
         BackColor       =   &H000080FF&
         Caption         =   "Xωρίς την τιμή αγοράς"
         Height          =   495
         Left            =   3240
         TabIndex        =   30
         Top             =   720
         Width           =   2355
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   120
      Top             =   9240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "utilities"
      Height          =   540
      Left            =   270
      TabIndex        =   1
      Top             =   120
      Width           =   1095
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

Dim F_DOROKARTA As String

Dim Col         As TrueOleDBGrid80.Column
Dim ColS As TrueOleDBGrid80.columns




Dim fSCR        As Object    ' SCRIPT CONTROL

Dim F928




Dim F_POS1_FOLDER As String
Dim F_POS2_FOLDER As String
Dim F_POS3_FOLDER As String

Dim F_MHXANI_FOLDER As String

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)
'  mForm_Load Me, fh, fw, ft, fl


Private Sub cmdClearFilter_Click()

        '<EhHeader>
        On Error GoTo cmdClearFilter_Click_Err

        '</EhHeader>

        'Clears filter from grid

100     For Each Col In TDBGrid.columns

110         Col.FILTERTEXT = ""

120     Next Col

130     Adodc2.Recordset.Filter = adFilterNone

        '<EhFooter>
        Exit Sub

cmdClearFilter_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.cmdClearFilter_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.cmdClearFilter_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub CmdupdateStructures_Click()
Update_Structures
End Sub

Private Sub cmdWebservice_Click()
  Dim c As DOMDocument
  '192.168.1.204
  
  c = HttpGetRequest("http://localhost:11180/webservice/checkreceipt?recno=3215118&recSeries=0011F&recType=174&format=xml")
  
  
  
End Sub

Private Sub cmdΤΗΛΕΦ_Click()
YPOSTHRIXI.SHOW


End Sub

Private Sub CMPUpdateMon_Click()
 Dim N As Integer
  N = GGET_NVALUE("select count(*) from EID  WHERE MON NOT IN (SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=2) ")
  MsgBox "ΥΠΑΡΧΟΥΝ " + str(N) + " ΕΙΔΗ ΠΟΥ ΔΕΝ ΕΙΝΑΙ ΣΤΟΝ ΚΑΤΑΛΟΓΟ"
  Text2.Text = "select count(*) as [ΠΛΗΘΟΣ],MON from EID  WHERE MON NOT IN (SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=2) GROUP BY MON"
  runQuery
  
  If N > 0 Then
     ' DMJ 24000229
  
     ANS = MsgBox("ΘΕΛΕΙΣ ΤΑ ΠΑΡΑΠΑΝΩ ΝΑ ΠΡΟΣΤΕΘΟΥΝ ΣΤΗΝ ΛΙΣΤΑ ΤΩΝ ΜΟΝΑΔΩΝ(ΘΑ ΛΕΙΠΟΥΝ ΜΟΝΟ ΤΑ UBL)", vbYesNo)
     If ANS = vbYes Then
         Gdb.Execute "INSERT INTO PINAKES(TYPOS,N1,PERIGRAFH) select 2 AS SS,count(*) as N1,MON from EID  WHERE MON NOT IN (SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=2) GROUP BY MON"
     End If
  End If
  
  N = GGET_NVALUE("select count(*) from EID  WHERE len(ISNULL(MON,''))=0 ")
  ' Dim ANS As Integer
  If N > 0 Then
    
     ANS = MsgBox("ΥΠΑΡΧΟΥΝ ΚΑΙ " + str(N) + " ΕΙΔΗ ΧΩΡΙΣ ΜΟΝΑΔΑ ΜΕΤΡΗΣΗΣ. ΝΑ ΕΜΦΑΝΙΣΤΟΥΝ;", vbYesNo)
     If ANS = vbYes Then
             Text2.Text = "select KOD,ONO,MON from EID  WHERE LEN(isnull(MON,''))=0 "
             runQuery
             
             Dim D As String
             D = InputBox("ΠΡΟΤΕΙΝΕ ΜΟΝΑΔΑ ΜΕΤΡΗΣΗΣ ΓΙΑ ΝΑ ΜΠΕΙ ΣΕ ΟΛΑ ΑΥΤΑ : ", "", "ΤΕΜ")
             ANS = MsgBox("ΝΑ ΜΠΟΥΝ ΣΕ ΟΛΑ ΑΥΤΑ ΣΑΝ ΜΟΝΑΔΑ ΜΕΤΡΗΣΗΣ ΤΟ " + D, vbYesNo)
             If ANS = vbYes Then
                Gdb.Execute "UPDATE EID SET MON='" + D + "'  WHERE LEN(isnull(MON,''))=0 "
             End If
             
             
     End If
     
  Else
     
     
     
     
  
  End If
  
End Sub

Private Sub Command11_Click()

        '<EhHeader>
        On Error GoTo Command11_Click_Err

        '</EhHeader>

100     cd1.InitDir = "c:\mercvb\queries"

110     cd1.ShowSave

        Dim F

120     F = cd1.FILENAME

        On Error Resume Next

130     Open F For Output As #5
140     Print #5, Text2.Text
150     Close #5

File1.Refresh


        '+ ".txt"

        '<EhFooter>
        Exit Sub

Command11_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command11_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command11_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command12_Click()

        'typoma me to componentone
        ' With TDBGrid.PrintInfo
        '         '.PageSetup
        '        ' Set the page header
        '      .PageHeaderFont.Italic = True
        '        .PageHeader = "Εκτύπωση " + Date$
        '
        '
        '         'Column headers will be on every page
        '        .RepeatColumnHeaders = True
        '
        '        ' Display page numbers (centered)
        '        .PageFooter = "\tσελίδα: \p"
        '        ' Invoke Print Preview
        '        .PrintPreview
        '  End With
        '<EhHeader>
        On Error GoTo Command12_Click_Err

        '</EhHeader>
100     If Check1 Then
            ' print7_excel UCase(Text2.Text), "11111111" 'sugm_str
        Else
110         print3_xar Text2.Text, "11111111", "", 0    'sugm_str
        End If

        '<EhFooter>
        Exit Sub

Command12_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command12_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command12_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command13_Click()

        ' METAFORA YPOLOIPOY EIDON
        '<EhHeader>
        On Error GoTo Command13_Click_Err

        '</EhHeader>
        Dim arxeio

        Dim a        As Long

        Dim DBF      As Database

        Dim sql      As New ADODB.Connection

        Dim rSQLEGGT As New ADODB.Recordset

        Dim rSQL     As New ADODB.Recordset

        Dim conSQL   As String

        Dim k        As Long

        Dim Fname    As String

        Dim hM       As String

100     hM = InputBox("δωσε ημερομηνία απογραφής", , "01/01/2006")

110     If IsDate(hM) Then
            'ok
        Else
120         MsgBox "λαθος ημερομηνία"
        End If

130     a = GetTickCount()
140     Me.Caption = "ENHMEΡΩΣΗ ΜΕ ΤΕΛΕΥΤΑΙΑ ΥΠΟΛΟΙΠΑ"
150     UPDATE_YPOLOIPA2
160     arxeio = "EID"
170     sql.Open gConnect

        On Error GoTo Command13_Click_Err

180     rSQL.Open "SELECT * FROM EID WHERE POS>0 ", Gdb, adOpenForwardOnly, adLockReadOnly
190     rSQLEGGT.Open "SELECT * FROM EGGTIM ", Gdb, adOpenDynamic, adLockBatchOptimistic

200     rSQL.MoveFirst

        Dim Z

210     Z = 0

        On Error GoTo printERROR    ' Resume Next

220     Do While Not rSQL.EOF

230         If Not IsNull(rSQL("POS01")) Then
240             If rSQL("POS01") > 0 Then
250                 rSQLEGGT.AddNew
260                 rSQLEGGT("hme") = CDate(hM)
270                 rSQLEGGT("ATIM") = "λ00001"
280                 rSQLEGGT("KODE") = rSQL("KOD")
290                 rSQLEGGT("ONOMA") = rSQL("ONO")
300                 rSQLEGGT("TIMM") = rSQL("XTI")
310                 rSQLEGGT("FPA") = rSQL("FPA")
320                 rSQLEGGT("POS0") = rSQL("POS01")
330                 rSQLEGGT("APOT") = 1
340                 rSQLEGGT.Update
                End If
            End If

350         If Not IsNull(rSQL("POS02")) Then
360             If rSQL("POS02") > 0 Then
370                 rSQLEGGT.AddNew
380                 rSQLEGGT("hme") = CDate(hM)
390                 rSQLEGGT("ATIM") = "λ00001"
400                 rSQLEGGT("KODE") = rSQL("KOD")
410                 rSQLEGGT("ONOMA") = rSQL("ONO")
420                 rSQLEGGT("TIMM") = rSQL("XTI")
430                 rSQLEGGT("FPA") = rSQL("FPA")
440                 rSQLEGGT("POS0") = rSQL("POS01")
450                 rSQLEGGT("APOT") = 2
460                 rSQLEGGT.Update
                End If
            End If

470         Z = Z + 1

480         If Z Mod 10 = 0 Then
490             Me.Caption = Z

500             DoEvents
            End If

510         rSQL.MoveNext
        Loop

520     rSQL.Close
530     rSQLEGGT.Close

540     Me.Caption = (GetTickCount() - a) / 1000

        Exit Sub

printERROR:
550     List2.Visible = True

560     List2.AddItem Err.Description

570     Resume Next

        '<EhFooter>
        Exit Sub

Command13_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command13_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command13_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command14_Click()

        '<EhHeader>
        On Error GoTo Command14_Click_Err

        '</EhHeader>

100     bohu62.SHOW

        '<EhFooter>
        Exit Sub

Command14_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command14_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command14_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command15_Click()

        '<EhHeader>
        On Error GoTo Command15_Click_Err

        '</EhHeader>

        Dim a

100     F928 = 1

110     Gdb.Execute "UPDATE PEL SET TYP=0 WHERE EIDOS='e' and TYP IS NULL"

120     Adodc2.ConnectionString = gConnect
130     Adodc2.RecordSource = "select * FROM PEL WHERE  EIDOS='e'  AND KOD<'9900' ORDER BY KOD"
140     Adodc2.Refresh

150     Set fSCR = CreateObject("MSScriptControl.ScriptControl")
160     fSCR.language = "vbscript"
170     fSCR.addObject "ADODC2", Adodc2
180     fSCR.addObject "mactext", MACtEXT

        ' Me.Caption = mac("Adodc2.Recordset(0)")  ok 6-12-06

190     a = toascii2("C:\MERCVB\F31.TXT", "C:\pel.TXT")

200     Adodc2.RecordSource = "select * FROM EID ORDER BY KOD"
210     Adodc2.Refresh

220     a = toascii2("C:\MERCVB\F33.TXT", "C:\EID.TXT")

230     Adodc2.RecordSource = "select * FROM EGG WHERE EIDOS='e'"
240     Adodc2.Refresh

250     a = toascii2("C:\MERCVB\F35.TXT", "C:\EGG.TXT")

        'Connect
        '    If RapiConnect Then
        '        Me.Caption = "Status: Connected!"
        '    Else
        '        Me.Caption = "Status: Disconnected!"
        '    End If
        '
        '
        '
        '
        '    If Not RapiIsConnected Then
        '        MsgBox "Device is not connected. Please connect first."
        '        Exit Sub
        '    End If
        '    If Not FileExists("C:\pel.TXT") Then
        '        MsgBox "ΔΕΝ ΒΡΙΣΚΩ ΤΟ ΑΡΧΕΙΟ ΤΩΝ ΠΕΛΑΤΩΝ ΓΙΑ ΑΠΟΣΤΟΛΗ", vbInformation
        '        Exit Sub
        '    End If
        '    If Not FileExists("C:\EID.TXT") Then
        '        MsgBox "ΔΕΝ ΒΡΙΣΚΩ ΤΟ ΑΡΧΕΙΟ ΤΩΝ ΕΙΔΩΝ ΓΙΑ ΑΠΟΣΤΟΛΗ", vbInformation
        '        Exit Sub
        '    End If
        '
        '    Call RAPICopyPCFileToCE("C:\PEL.TXT", "\PEL.TXT")
        '    Call RAPICopyPCFileToCE("C:\EID.TXT", "\EID.TXT")
        '

        On Error Resume Next

260     Shell "c:\sendtoce.bat", vbMinimizedFocus

270     MsgBox "Τέλος Αποστολής"

280     Me.Caption = mac("Adodc2.Recordset(3)")

290     Set fSCR = Nothing

        '<EhFooter>
        Exit Sub

Command15_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command15_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command15_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command17_Click()

'TeST git

If times.Value = vbChecked Then
    TIMES_UPD
End If

If BONUS.Value = vbChecked Then
    LOAD_PELATES
End If

'TeSt   KL


'tttttttt
'ppppppppppppppppppppppppppp


End Sub


Sub TIMES_UPD()
        ' If Len(Dir("C:\MERCVB\TEST.EXE", vbNormal)) > 0 And cXRONIES.Value = vbChecked Then

'OK ----00 START=021     LEN=13    FIX=                            BARCODE                                                                 A
'OK----01       036         30                                    ΠΕΡΙΓΡΑΦΗ                                                              G
'OK----03       087         08                                    ΤΙΜΗ                                                                D,###.00
'OK ----10       095         01                                    ΦΠΑ 1=6,2=13,3=24,5=0                                                                A
'OK'' ----12       107         02                                    ΜΟΝΑΔΑ ΜΕΤΡΗΣΗΣ 1=ΤΕΜ 2=ΚΙΛΑ                                                    A
'OK----17       095         01                                    ΤΜΗΜΑ                                                              I
'----20       140         01                                    ΕΛΕΥΘΕΡΗ ΤΙΜΗ 0=ΟΧΙ 1=ΝΑΙ                                                        I,S0~1
'----21       137         01        0                           ΖΥΓΙΖΟΜΕΝΟ ΕΙΔΟΣ 0=ΟΧΙ 1=ΝΑΙ                                                              I
'----24       131         04        0                           ΠΟΝΤΟΙ ΕΙΔΟΥΣ                                                           L
'OK  ----25       00115                                    Β ΚΩΔΙΚΟΣ




Dim R22 As New ADODB.Recordset


            Open "C:\MERCVB\ERRORSPOSFILE.TXT" For Output As #2

            Open "C:\MERCVB\POSFILE.TXT" For Output As #1
             ' ena eidos
              R22.Open "SELECT  EID.KOD,BARCODES.ERG,ONO,LTI5,( CASE WHEN MON IS NULL THEN  '1' ELSE '1' END )  AS MON,( CASE WHEN FPA=4 THEN 1 ELSE ( CASE WHEN FPA=1 THEN 2 ELSE 3 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID RIGHT JOIN BARCODES ON EID.KOD=BARCODES.KOD where  NOT ( EID.KOD LIKE '913%' ) ORDER BY EID.KOD ", Gdb, adOpenDynamic, adLockOptimistic  ' WHERE EID.KOD='" + Text1(0) + "'"
             
             ' ola ta eidh
             'R22.Open "SELECT KOD,ONO,LTI5,( CASE WHEN MON IS NULL THEN  'TEM' ELSE MON END )  AS MON,( CASE WHEN FPA=1 THEN 13 ELSE ( CASE WHEN FPA=5 THEN 0 ELSE 23 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID  ", Gdb, adOpenDynamic, adLockOptimistic
                'On Error GoTo iliadis_error


Dim mLTI5 As Single
Dim MONO As String
Dim MKOD As String
Dim mbar As String
Dim MFPA As Integer
Dim MBARC As String


Dim nc As Long


            Do While Not R22.EOF
                ' Open "C:\POSFILE.TXT" For Output As #1
                
             If IsNull(R22(0)) Then
                MKOD = "."
             Else
                MKOD = R22(0)
             End If
                
                
              If MKOD = "005.429" Then
                 nc = nc + 1
              End If
                
             MKOD = Replace(MKOD, Chr(10), " ")
             MKOD = Replace(MKOD, Chr(13), " ")
                
                
             CC = Left(MKOD + Space(21), 20)  ' KVDIKOS
             
             
             
             'R22 ("ONO")
             If IsNull(R22("ono")) Then
                MONO = "."
             Else
                MONO = R22("ono")
             End If
             MONO = Replace(MONO, Chr(10), " ")
             MONO = Replace(MONO, Chr(13), " ")
             
             'R22 ("ONO")
             If IsNull(R22(1)) Then
                MBARC = "0000"
             Else
                MBARC = R22(1)
             End If
             
             MBARC = Replace(MBARC, Chr(10), " ")
             MBARC = Replace(MBARC, Chr(13), " ")

             
             
             CC = CC + Left(MBARC + Space(21), 15) ' BARCODE
             
             
             
             
             
             CC = CC + mID(MONO + String(30, " "), 1, 30) + String(21, " ") ' Space(21)
             If IsNull(R22("LTI5")) Then
                mLTI5 = 0
             Else
                mLTI5 = R22("LTI5")
             End If
             
             CC = CC + Replace(Format(mLTI5, "00000.00"), ".", ",") '94   '+ String(F_DEK_LIANIKIS, "0"))
     
                '  If IsNull(R22("mon")) Then
                '     CC = CC + Space(3)
                '   Else
                '      CC = CC + R22("mon") '+ Space(3), 3)
                ' End If
     
             
             
             If IsNull(R22("FPA1")) Then
                MFPA = 3
             Else
                MFPA = R22("FPA1")
             End If
             
             
             DoEvents
             
             Me.Caption = nc
             nc = nc + 1
             
             CC = CC + Format(MFPA, "0") + Space(10)
             CC = CC + " 1"    ' monada Format(R22("mon")
             
             
   '          On Error Resume Next
             
            ' CC = CC + Space(18) + Format(R22("tmhma"), "0.00")
             
             CC = CC + "                       00    0  0      "   ' ποντοι ζυγιζομενα
             CC = Replace(CC, Chr(13), "")
             If Len(CC) < 140 Then
                 Print #2, CC
             Else
                 Print #1, CC
             End If
'123456789012340078900230567890
                  
                R22.MoveNext
            Loop






             Close #1
             
             Close #2
             
             
             
             R22.Close
             
             
             
Dim KL As Integer
Dim ckl As String
             
             
             On Error Resume Next
             Dim ANS As Integer
             ANS = MsgBox("να αποσταλούν στα pos? ", vbYesNo)
             If ANS = vbYes Then
             

             
                 If Len(Trim(F_POS1_FOLDER)) > 1 Then
                     If FolderExists(F_POS1_FOLDER) Then
                          FileCopy "C:\MERCVB\POSFILE.TXT", F_POS1_FOLDER + "\POSFILE.TXT"
                     Else
                          MsgBox "δεν υπάρχει ο φάκελος" + F_POS1_FOLDER
                     End If
                     
                 Else
                 
                     
                    For KL = 1 To 3
                         ckl = Format(KL, "0")
                         If FolderExists("\\POS" + ckl + "\TEC_POS\DATA") Then
                            FileCopy "C:\MERCVB\POSFILE.TXT", "\\POS" + ckl + "\TEC_POS\DATA\POSFILE.TXT"
                          Else
                            MsgBox "δεν υπάρχει ο φάκελος" + "\\POS" + ckl + "\TEC_POS\DATA\POSFILE.TXT"
                          End If
                    Next
                    ' FileCopy "C:\MERCVB\POSFILE.TXT", "\\POS2\TEC_POS\DATA\POSFILE.TXT"
                    ' FileCopy "C:\MERCVB\POSFILE.TXT", "\\POS3\TEC_POS\DATA\POSFILE.TXT"
                 End If
             End If
            
            
' LOAD_PELATES




                On Error Resume Next
                
               ' Exit Function

         '   End If

End Sub
Sub LOAD_PELATES()
        ' If Len(Dir("C:\MERCVB\TEST.EXE", vbNormal)) > 0 And cXRONIES.Value = vbChecked Then

'OK ----00 START=021     LEN=13    FIX=                            BARCODE                                                                 A
'OK----01       036         30                                    ΠΕΡΙΓΡΑΦΗ                                                              G
'OK----03       087         08                                    ΤΙΜΗ                                                                D,###.00
'OK ----10       095         01                                    ΦΠΑ 1=6,2=13,3=24,5=0                                                                A
'OK'' ----12       107         02                                    ΜΟΝΑΔΑ ΜΕΤΡΗΣΗΣ 1=ΤΕΜ 2=ΚΙΛΑ                                                    A
'OK----17       095         01                                    ΤΜΗΜΑ                                                              I
'----20       140         01                                    ΕΛΕΥΘΕΡΗ ΤΙΜΗ 0=ΟΧΙ 1=ΝΑΙ                                                        I,S0~1
'----21       137         01        0                           ΖΥΓΙΖΟΜΕΝΟ ΕΙΔΟΣ 0=ΟΧΙ 1=ΝΑΙ                                                              I
'----24       131         04        0                           ΠΟΝΤΟΙ ΕΙΔΟΥΣ                                                           L
'OK  ----25       00115                                    Β ΚΩΔΙΚΟΣ




Dim R22 As New ADODB.Recordset


            Open "C:\MERCVB\points.upd" For Output As #2

            Open "C:\MERCVB\customer.upd" For Output As #1
             ' ena eidos
              R22.Open "SELECT  * FROM EID  where  KOD LIKE '913%'  ORDER BY EID.KOD ", Gdb, adOpenDynamic, adLockOptimistic  ' WHERE EID.KOD='" + Text1(0) + "'"
             
             ' ola ta eidh
             'R22.Open "SELECT KOD,ONO,LTI5,( CASE WHEN MON IS NULL THEN  'TEM' ELSE MON END )  AS MON,( CASE WHEN FPA=1 THEN 13 ELSE ( CASE WHEN FPA=5 THEN 0 ELSE 23 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID  ", Gdb, adOpenDynamic, adLockOptimistic
                'On Error GoTo iliadis_error


Dim mLTI5 As Single
Dim MONO As String
Dim MKOD As String
Dim mbar As String
Dim MFPA As Integer
Dim MBARC As String

Dim cp As String
Dim nc As Long




            Do While Not R22.EOF
                CC = Space(250)
                cp = Space(89)
                
                
                ' Open "C:\POSFILE.TXT" For Output As #1
                
             If IsNull(R22("kod")) Then
                MKOD = "."
             Else
                MKOD = R22("kod")
             End If
                
                
'              If mkod = "005.429" Then
'                 nc = nc + 1
'              End If
                
             MKOD = Replace(MKOD, Chr(10), " ")
             MKOD = Replace(MKOD, Chr(13), " ")
                
                
              Mid(CC, 1, 1) = "0"  'Left(mkod + Space(15), 15)
              '  Mid(CC, 3, 15) = Left(mkod + Space(16), 15)   '      Left(mID(mkod, 8, 6) + Space(15), 15)
              Mid(CC, 3, 15) = Left(MKOD + Space(15), 15)
              
              Mid(CC, 19, 20) = Left("ΠΕΛΑΤΗΣ ΛΙΑΝΙΚΗΣ" + Space(15), 20)
              
              
              
              Mid(CC, 246, 2) = "00"
              Mid(CC, 244, 1) = "1"
              Mid(CC, 201, 16) = Left(MKOD + Space(16), 16)
              
              
              
              
              Mid(cp, 1, 1) = "0"  'Left(mkod + Space(15), 15)
              Mid(cp, 3, 15) = Left(MKOD + Space(15), 15)
              Mid(cp, 19, 20) = Left("ΠΕΛΑΤΗΣ ΛΙΑΝΙΚΗΣ" + Space(15), 20)
              
              Mid(cp, 60, 1) = "0"
              
              If IsNull(R22("pontoi")) Then
              
                  Mid(cp, 66, 5) = "0"
              Else
                  Mid(cp, 66, 5) = Right("      " + str(R22("pontoi")), 5)
              
              End If
              
              'Mid(cp, 66, 5) = "0"
              
              
              'Mid(CC, 201, 16) = Left(mkod + Space(16), 16)
              
              
              
              
              
              
 '             0=0   3->15  kodikos  customer.upd   data
'60=>0      66 =>5


              
              
              
             
             DoEvents
             
             Me.Caption = nc
             nc = nc + 1
             
             
             
            ' If Len(CC) < 140 Then
                 Print #2, cp
             'Else
                 Print #1, CC
             'End If

                  
                R22.MoveNext
                
                
                ' If nc > 10 Then Exit Do
                
            Loop






             Close #1
             
             Close #2
             
             
             
             R22.Close
             
  Dim ANS As Integer
   
ANS = MsgBox("να αποσταλούν στα pos? ", vbYesNo)
If ANS = vbYes Then



                 If Len(Trim(F_POS1_FOLDER)) > 1 Then
                     If FolderExists(F_POS1_FOLDER) Then
                          FileCopy "C:\MERCVB\customer.upd", F_POS1_FOLDER + "\customer.upd"
                          FileCopy "C:\MERCVB\points.upd", F_POS1_FOLDER + "\points.upd"
                          
                     Else
                          MsgBox "δεν υπάρχει ο φάκελος" + F_POS1_FOLDER
                     End If
                     
                 Else  ' points.upd
                 
                     
                    For KL = 1 To 3
                         ckl = Format(KL, "0")
                         If FolderExists("\\POS" + ckl + "\TEC_POS\DATA") Then
                              FileCopy "C:\MERCVB\customer.upd", "\\POS" + ckl + "\TEC_POS\DATA\customer.upd"
                              FileCopy "C:\MERCVB\points.upd", "\\POS" + ckl + "\TEC_POS\DATA\points.upd"
                          Else
                              MsgBox "δεν υπάρχει ο φάκελος" + "\\POS" + ckl + "\TEC_POS\DATA"
                          End If
                    Next
                    ' FileCopy "C:\MERCVB\POSFILE.TXT", "\\POS2\TEC_POS\DATA\POSFILE.TXT"
                    ' FileCopy "C:\MERCVB\POSFILE.TXT", "\\POS3\TEC_POS\DATA\POSFILE.TXT"
                 End If




















  
'  If Len(Trim(F_POS1_FOLDER)) > 1 Then
'                     FileCopy "C:\MERCVB\customer.upd", F_POS1_FOLDER + "\customer.upd"
'                     FileCopy "C:\MERCVB\points.upd", F_POS1_FOLDER + "\points.upd"
'  Else
'                     FileCopy "C:\MERCVB\UPDATE.CUS", "\\POS1\TEC_POS\FILES\customer.upd"
'                     FileCopy "C:\MERCVB\UPDATE.CUS", "\\POS2\TEC_POS\FILES\customer.upd"
'                     FileCopy "C:\MERCVB\UPDATE.CUS", "\\POS3\TEC_POS\FILES\customer.upd"
'                     FileCopy "C:\MERCVB\points.upd", "\\POS1\TEC_POS\FILES\points.upd"
'                      FileCopy "C:\MERCVB\UPDATE.CUS", "\\POS2\TEC_POS\FILES\points.upd"
'                     FileCopy "C:\MERCVB\UPDATE.CUS", "\\POS3\TEC_POS\FILES\points.upd"
'   End If
   

End If


                On Error Resume Next
                
                Exit Sub

         '   End If

End Sub










Private Sub Command21_Click()

        '<EhHeader>
        On Error GoTo Command21_Click_Err

        '</EhHeader>

100     Gdb.Execute "DELETE FROM EID"

        Dim R As New ADODB.Recordset

        Dim mk, mO, mF, mL

        Dim mERG, mP

110     GCONNECT2 = gConnect + "DATABASE=super;"

        Dim DB2 As New ADODB.Connection

120     DB2.Open GCONNECT2
130     R.Open "SELECT * FROM EIDH3 ORDER BY KOD", DB2, adOpenDynamic, adLockOptimistic

140     Do While Not R.EOF
150         DoEvents
160         Me.Caption = R("KOD")

170         If IsNull(mk) Then mk = " " Else mk = Left(R("KOD"), 14)
180         If IsNull(mF) Then mF = "2" Else mF = Format(R("FPA"), "##")
190         If IsNull(mERG) Then mERG = " " Else mERG = R("BARCODE")
200         If IsNull(mO) Then mO = " " Else mO = Replace(Left(R("ONO"), 35), "'", "-")
210         If IsNull(mL) Then mL = "0" Else mL = Format(R("LTI5"), "####0.##")
220         If IsNull(mP) Then mP = " " Else mP = R("PROM")
230         Gdb.Execute "INSERT INTO EID (KOD,ONO,FPA,LTI5,ERG,PROM) VALUES ( '" + mk + "','" + mO + "'," + mF + "," + mL + ",'" + mERG + "','" + mP + "')"

240         R.MoveNext

        Loop

250     MsgBox "ΤΕΛΟΣ ΑΠΟΡΡΟΦΗΣΗΣ"

        '<EhFooter>
        Exit Sub

Command21_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command21_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command21_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command22_Click()

        '<EhHeader>
        On Error GoTo Command22_Click_Err

        '</EhHeader>

        On Error Resume Next

100     Gdb.Close
110     gConnect = InputBox("ΔΩΣΕ ΝΕΟ CONNECTION STRING(Π.X. DSN=REPORTS ")

120     Gdb.Open gConnect

        '<EhFooter>
        Exit Sub

Command22_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command22_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command22_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command23_Click()

        '<EhHeader>
        On Error GoTo Command23_Click_Err

        '</EhHeader>
        Dim DUM

100     DUM = updateSQL2SQL("EGGTIM", "DSN=IPMAGAZIOY;UID=sa;pwd=epsilonsa")
        'SELECT * INTO ARITMISI  IN 'C:\' 'Text;'  FROM ARITMISI

        '<EhFooter>
        Exit Sub

Command23_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command23_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command23_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command24_Click()

        '<EhHeader>
        On Error GoTo Command24_Click_Err

        '</EhHeader>

        Dim R As New ADODB.Recordset

100     R.Open "SELECT * FROM EID WHERE LEN(KOD)=12 AND LEFT(KOD,1)<>'0'", Gdb, adOpenDynamic, adLockOptimistic

        Dim a As String

        Dim B As String

110     R.MoveFirst

120     Do While Not R.EOF
130         B = R("KOD")
140         a = B + ChDigEAN13(B)    'timtext3.Text

            On Error Resume Next

150         Gdb.Execute "UPDATE EID SET KOD='" + a + "'  WHERE KOD='" + B + "'", k
160         R.MoveNext
        Loop

170     R.Close

        '<EhFooter>
        Exit Sub

Command24_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command24_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command24_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub



Private Sub Command25_Click()
 With TDBGrid.PrintInfo
        ' Set the page header
        .PageHeaderFont.Italic = True
        .PageHeader = Text2.Text ' "Composers table"
        
        ' Column headers will be on every page
        .RepeatColumnHeaders = True
        
        ' Display page numbers (centered)
        .PageFooter = "\tPage: \p"
        ' Invoke Print Preview
        .PrintPreview
    End With
End Sub

Private Sub Command26_Click()
On Error Resume Next
Dim sql As String
        sql = "update PINAKES SET C2='S' WHERE TYPOS=1;update PINAKES SET C2='E' WHERE TYPOS=1 AND TIMH=0;"
sql = sql + "update PINAKES SET C2='H87' WHERE TYPOS=2 AND  PERIGRAFH LIKE '%TEM%';"
sql = sql + "update PINAKES SET C2='MTR' WHERE TYPOS=2 AND  PERIGRAFH LIKE '%MET%';"
sql = sql + "update PINAKES SET C2='KGM' WHERE TYPOS=2 AND  PERIGRAFH LIKE '%ΚΙΛ%';UPDATE  PARASTAT SET UBL=(CASE  WHEN EIDOS IN ('P','p') then '381' else '380' end)   WHERE ISMYDATA=1 "
        Gdb.Execute sql

        Gdb.Execute "alter table PINAKES ADD PRIMARY KEY (ID)"

        Gdb.Execute "alter table PARAMETROI ADD PRIMARY KEY (ID)"
End Sub

Private Sub Command27_Click()
Dim N As Integer
Dim B As String

B = InputBox("ARXIKO GRAMMA?")

For N = 1 To 1000
   FF = B + LTrim(str(N))
   Gdb.Execute "INSERT INTO TIM (KLEIDI) VALUES ('" + FF + "')"
Next
   
End Sub

Private Sub CrystalToPDF_Click()
' ?p???t??µe ?t? ??e?? ??a Crystal Report Control (p.?. CrystalReport1)
'CrystalReport1.ReportFileName = "C:\mercvb\reports\timol1.rpt"
CrystalReport1.Destination = crptToFile
CrystalReport1.PrintFileType = crptPDF
CrystalReport1.PrintFileName = "C:\mercvb\reports\timol1.pdf"
CrystalReport1.Action = 1
End Sub

Private Sub Dataentry_Click()
   bohu61.SHOW
   
End Sub

Private Sub Dir1_Change()

100     File1.Path = Dir1.Path
110     File1.Refresh
End Sub

Private Sub energ_pel_Click(index As Integer)
       make_energoi "e"
       
End Sub
Private Sub make_energoi(buff As String)
   Dim c As String: c = GGET_CVALUE("SELECT DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS   WHERE COLUMN_NAME='ENERGOS' AND TABLE_NAME='PEL'  ")
   If UCase(c) = "BIT" Then
      Gdb.Execute "ALTER TABLE PEL ALTER COLUMN ENERGOS INTEGER NULL"
   End If
   
        Gdb.Execute "UPDATE PEL SET ENERGOS=1 WHERE EIDOS='" + buff + "'"  'ISNULL(ENERGOS,0)=0  AND

End Sub

Private Sub energ_prom_Click(index As Integer)
     make_energoi "r"
End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 15465, 10545, fh, fw, ft, fl

End Sub

Private Sub IOSIFIDIS_Click()

        '<EhHeader>
        On Error GoTo IOSIFIDIS_Click_Err

        '</EhHeader>
        Dim a

100     F928 = 1

110     Adodc2.ConnectionString = gConnect
120     Adodc2.RecordSource = "select  TOP 10 * FROM PEL "    'WHERE  EIDOS='e' ORDER BY KOD"
130     Adodc2.Refresh

140     Set fSCR = CreateObject("MSScriptControl.ScriptControl")
150     fSCR.language = "vbscript"
160     fSCR.addObject "R", Adodc2.Recordset
170     fSCR.addObject "mactext", MACtEXT

180     F928 = 0

190     a = ToAsciiMultiLine("C:\MERCVB\F38.TXT", "C:\EID.TXT")
200     Set fSCR = Nothing

        Dim mm2 As String
        Dim ETOS As String
        ETOS = InputBox("ETOΣ ")
210     MM = InputBox("ΑΠΟ ΜΗΝΑ ")
220     mm2 = InputBox("ΕΩΣ ΜΗΝΑ ")
230     L = 0

240     Adodc2.RecordSource = "select  * FROM TIM WHERE NOT KPE IS NULL AND LEFT(ATIM,1) in ('T','t') AND MONTH(HME)>=" + MM + " AND MONTH(HME)<=" + mm2 + " AND YEAR(HME)=" + ETOS + "  order by HME" '
250     Adodc2.Refresh

        '   Adodc2.RecordSource = "select * FROM EID ORDER BY KOD"
        '  Adodc2.Refresh

        ' A = toascii2("C:\MERCVB\F33.TXT", "C:\EID.TXT")

        ' ME ASCII
        'Set fSCR = CreateObject("MSScriptControl.ScriptControl")
        'fSCR.language = "vbscript"
        'fSCR.AddObject "R", Adodc2.Recordset
        'fSCR.AddObject "mactext", mactext
        ' A = ToAsciiMultiLine("C:\MERCVB\F39.TXT", "C:\LOG.TXT")

        ' MsgBox "Τέλος Αποστολής"

260     Me.Caption = mac("Adodc2.Recordset(3)")

270     Set fSCR = Nothing

        ' ME ASCII
        'Set fSCR = CreateObject("MSScriptControl.ScriptControl")
        'fSCR.language = "vbscript"
        'fSCR.AddObject "R", Adodc2.Recordset
        'fSCR.AddObject "mactext", mactext
        ' A = ToAsciiMultiLine("C:\MERCVB\F39.TXT", "C:\LOG.TXT")

        '
        ' karfoto sthn dbIII
        Dim DB  As Database

        Dim RDB As Recordset

        Dim R   As New ADODB.Recordset

        '
        '
        '
        '  Set db = OpenDatabase("c:\mercvb\iosifidis", False, False, "dBase III;")  '"FoxPro 3.0;") '
        ' Set RDB = db.OpenRecordset("select max(a_a) from kinped1")
        'If RDB.EOF Then
        '  L = 0
        'Else
        '  L = 0
        '  On Error Resume Next
        '  L = RDB(0)
        'End If

        Dim AIT As String

        '

        Dim W   As String

280     Open "c:\log.txt" For Output As #1

290     Adodc2.Recordset.MoveFirst

300     Do While Not Adodc2.Recordset.EOF

310         R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic
320         R.MoveFirst

330         Do While Not R.EOF

340             If Left(Adodc2.Recordset("ATIM"), 1) = R("EIDOS") Then
350                 AIT = to437(R("TITLOS"))
                End If

360             R.MoveNext
            Loop

370         R.Close

380         W = ""

390         L = 0
400         L = L + 1

            Dim TEL As Single

410         TEL = Adodc2.Recordset("AJ1") + Adodc2.Recordset("AJ2") + Adodc2.Recordset("AJ3") + Adodc2.Recordset("AJ4") + Adodc2.Recordset("AJ5") + Adodc2.Recordset("AJ6") + Adodc2.Recordset("AJ7")
420         TEL = TEL + Adodc2.Recordset("fpa1") + Adodc2.Recordset("fpa2") + Adodc2.Recordset("fpa3") + Adodc2.Recordset("fpa4") + Adodc2.Recordset("fpa6") + Adodc2.Recordset("fpa7")

430         W = W + Format(L, "00") + " "
440         W = W + "30.00.0" + Adodc2.Recordset("kpe") + Space(9)
450         W = W + Left(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 6) + Right(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 2)
460         W = W + " 70 "
470         W = W + Left(AIT + Space(15), 41)
480         W = W + mID(Adodc2.Recordset("ATIM"), 2, 7)
490         W = W + Format(TEL, "000000000000000.00")
500         W = W + Space(8) + "1"
510         Print #1, W

520         If Adodc2.Recordset("AJ1") > 0 Then
530             W = ""
540             L = L + 1
550             W = W + Format(L, "00") + " "
560             W = W + "70.00.00013" + Space(9)
570             W = W + Left(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 6) + Right(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 2)
580             W = W + " 70 "
590             W = W + Left(AIT + Space(15), 41)
600             W = W + mID(Adodc2.Recordset("ATIM"), 2, 7)
610             W = W + Format(Adodc2.Recordset("AJ1"), "000000000000000.00")
620             W = W + Space(8) + "0"
630             Print #1, W
            End If

640         If Adodc2.Recordset("AJ2") + Adodc2.Recordset("AJ6") > 0 Then
650             W = ""
660             L = L + 1
670             W = W + Format(L, "00") + " "
680             W = W + "70.00.00024" + Space(9)
690             W = W + Left(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 6) + Right(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 2)
700             W = W + " 70 "
710             W = W + Left(AIT + Space(15), 41)
720             W = W + mID(Adodc2.Recordset("ATIM"), 2, 7)
730             W = W + Format(Adodc2.Recordset("AJ2") + Adodc2.Recordset("AJ6"), "000000000000000.00")
740             W = W + Space(8) + "0"
750             Print #1, W
            End If

760         If Adodc2.Recordset("AJ3") > 0 Then
770             W = ""
780             L = L + 1
790             W = W + Format(L, "00") + " "
800             W = W + "70.00.00017" + Space(9)
810             W = W + Left(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 6) + Right(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 2)
820             W = W + " 70 "
830             W = W + Left(AIT + Space(15), 41)
840             W = W + mID(Adodc2.Recordset("ATIM"), 2, 7)
850             W = W + Format(Adodc2.Recordset("AJ3"), "000000000000000.00")
860             W = W + Space(8) + "0"
870             Print #1, W
            End If

880         If Adodc2.Recordset("AJ4") + Adodc2.Recordset("AJ7") > 0 Then

890             W = ""
900             L = L + 1
910             W = W + Format(L, "00") + " "
920             W = W + "70.00.00009" + Space(9)
930             W = W + Left(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 6) + Right(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 2)
940             W = W + " 70 "
950             W = W + Left(AIT + Space(15), 41)
960             W = W + mID(Adodc2.Recordset("ATIM"), 2, 7)
970             W = W + Format(Adodc2.Recordset("AJ4") + Adodc2.Recordset("AJ7"), "000000000000000.00")
980             W = W + Space(8) + "0"
990             Print #1, W
            End If

1000        If Adodc2.Recordset("fpa1") > 0 Then
1010            W = ""
1020            L = L + 1
1030            W = W + Format(L, "00") + " "
1040            W = W + "85.70.00013" + Space(9)
1050            W = W + Left(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 6) + Right(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 2)
1060            W = W + " 70 "
1070            W = W + Left(AIT + Space(15), 41)
1080            W = W + mID(Adodc2.Recordset("ATIM"), 2, 7)
1090            W = W + Format(Adodc2.Recordset("fpa1"), "000000000000000.00")
1100            W = W + Space(8) + "0"
1110            Print #1, W
            End If

1120        If Adodc2.Recordset("fpa2") + Adodc2.Recordset("fpa6") > 0 Then
1130            W = ""
1140            L = L + 1
1150            W = W + Format(L, "00") + " "
1160            W = W + "85.70.00024" + Space(9)
1170            W = W + Left(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 6) + Right(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 2)
1180            W = W + " 70 "
1190            W = W + Left(AIT + Space(15), 41)
1200            W = W + mID(Adodc2.Recordset("ATIM"), 2, 7)
1210            W = W + Format(Adodc2.Recordset("fpa2") + Adodc2.Recordset("fpa6"), "000000000000000.00")
1220            W = W + Space(8) + "0"
1230            Print #1, W
            End If

1240        If Adodc2.Recordset("fpa3") > 0 Then
1250            W = ""
1260            L = L + 1
1270            W = W + Format(L, "00") + " "
1280            W = W + "85.70.00016" + Space(9)
1290            W = W + Left(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 6) + Right(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 2)
1300            W = W + " 70 "
1310            W = W + Left(AIT + Space(15), 41)
1320            W = W + mID(Adodc2.Recordset("ATIM"), 2, 7)
1330            W = W + Format(Adodc2.Recordset("fpa3"), "000000000000000.00")
1340            W = W + Space(8) + "0"
1350            Print #1, W
            End If

1360        If Adodc2.Recordset("fpa4") > 0 Then
1370            W = ""
1380            L = L + 1
1390            W = W + Format(L, "00") + " "
1400            W = W + "85.70.00009" + Space(9)
1410            W = W + Left(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 6) + Right(Format(Adodc2.Recordset("hme"), "DD/MM/YYYY"), 2)
1420            W = W + " 70 "
1430            W = W + Left(AIT + Space(15), 41)
1440            W = W + mID(Adodc2.Recordset("ATIM"), 2, 7)
1450            W = W + Format(Adodc2.Recordset("fpa4") + Adodc2.Recordset("fpa7"), "000000000000000.00")
1460            W = W + Space(8) + "0"
1470            Print #1, W
            End If

1480        Adodc2.Recordset.MoveNext

        Loop

1490    Close #1

1500    MsgBox "Τέλος Αποστολής"

1510    Me.Caption = mac("Adodc2.Recordset(3)")

1520    Set fSCR = Nothing

        '<EhFooter>
        Exit Sub

IOSIFIDIS_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.IOSIFIDIS_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.IOSIFIDIS_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function ToAsciiMultiLine(ByVal arxeio As String, moutput As String)

        ' arxeio π.χ. f90.txt  η φόρμα της εκτύπωσης    host.txt to arxeio me ejagogi ascii
        '<EhHeader>
        On Error GoTo ToAsciiMultiLine_Err

        '</EhHeader>

        Dim DUM

        Dim eField(1 To 50), ePict(1 To 50), eProw(1 To 50), ePcol(1 To 50), eFonts(1 To 50)

        Dim mField(1 To 50), mPict(1 To 50), mProw(1 To 50), mPcol(1 To 50), mFonts(1 To 50)

        Dim sField(1 To 50), sPict(1 To 50), sProw(1 To 50), sPcol(1 To 50), sFonts(1 To 50)

        Dim m_No_of_seir, i, k, m, c

        Dim m_seir_synol, m_seir_eid, sf

        Dim a(1 To 250), ar_ped(1 To 250)

        Dim xa(1 To 250), m_m, npic

        Dim m1

        ' dim gm_str(1 To 250)
        Dim DB     As Database

        Dim TIM    As Recordset

        Dim PEL    As Recordset

        Dim EGGTIM As Recordset

        Dim m_entol_ektyp

        Dim xart11, kod_tim

        Dim ejodos

        Dim e, MM, s

        Dim ektypoths

        Dim m_syn_row, syn_row

        Dim CDOK

100     Adodc2.Recordset.MoveFirst
110     DUM = load2_forma(arxeio, 1)

120     Open moutput For Output As #1

        Dim S2 As String

130     Adodc2.Recordset.MoveFirst

140     Do While Not Adodc2.Recordset.EOF
150         S2 = ""

160         For k = 1 To 90

170             If gm_r(k) = 1 And k > 1 Then
180                 Print #1, S2
190                 S2 = ""              ' : Print #1, Chr(13) 'allagh seiras
                End If

200             If Len(Trim(gm_f(k))) > 0 Then
210                 MM = mac(Trim(Trim(gm_f(k))))                   ': m_m = ExecuteLine("gvar=" + MM)
220                 S2 = S2 + ALIGNFORMAT(MM, gpic(k))    '
                End If

            Next

230         Print #1, S2
240         Adodc2.Recordset.MoveNext
        Loop

250     Close #1
260     Adodc2.Recordset.Close

        '<EhFooter>
        Exit Function

ToAsciiMultiLine_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.ToAsciiMultiLine " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.ToAsciiMultiLine " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function toascii2(ByVal arxeio As String, moutput As String)

        ' arxeio π.χ. f90.txt  η φόρμα της εκτύπωσης    host.txt to arxeio me ejagogi ascii
        '<EhHeader>
        On Error GoTo toascii2_Err

        '</EhHeader>

        Dim DUM

        Dim eField(1 To 50), ePict(1 To 50), eProw(1 To 50), ePcol(1 To 50), eFonts(1 To 50)

        Dim mField(1 To 50), mPict(1 To 50), mProw(1 To 50), mPcol(1 To 50), mFonts(1 To 50)

        Dim sField(1 To 50), sPict(1 To 50), sProw(1 To 50), sPcol(1 To 50), sFonts(1 To 50)

        Dim m_No_of_seir, i, k, m, c

        Dim m_seir_synol, m_seir_eid, sf

        Dim a(1 To 250), ar_ped(1 To 250)

        Dim xa(1 To 250), m_m, npic

        Dim m1

        ' dim gm_str(1 To 250)
        Dim DB     As Database

        Dim TIM    As Recordset

        Dim PEL    As Recordset

        Dim EGGTIM As Recordset

        Dim m_entol_ektyp

        Dim xart11, kod_tim

        Dim ejodos

        Dim e, MM, s

        Dim ektypoths

        Dim m_syn_row, syn_row

        Dim CDOK

100     Adodc2.Recordset.MoveFirst
110     DUM = load2_forma(arxeio, 1)

120     Open moutput For Output As #1

        Dim S2 As String

130     Adodc2.Recordset.MoveFirst

140     Do While Not Adodc2.Recordset.EOF
150         S2 = ""

160         For k = 1 To 8
170             MM = mac(Trim(Trim(gm_f(k))))                   ': m_m = ExecuteLine("gvar=" + MM)
180             S2 = S2 + ALIGNFORMAT(MM, gpic(k))    '
            Next

190         Print #1, S2
200         Adodc2.Recordset.MoveNext

210         DoEvents
        Loop

220     Close #1
230     Adodc2.Recordset.Close

        '<EhFooter>
        Exit Function

toascii2_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.toascii2 " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.toascii2 " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function ALIGNFORMAT(MM, PIC)

        ' RIGHT ALIGNMENT
        '<EhHeader>
        On Error GoTo ALIGNFORMAT_Err

        '</EhHeader>
100     If mID(PIC, 2, 1) = "#" Then
110         ALIGNFORMAT = Right(Space(Len(PIC)) + Replace(MM, ",", "."), Len(PIC))
        Else
            ' LEFT ALIGNMENT
120         ALIGNFORMAT = Left(MM + Space(Len(PIC)), Len(PIC))
        End If

        '<EhFooter>
        Exit Function

ALIGNFORMAT_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.ALIGNFORMAT " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.ALIGNFORMAT " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub Command16_Click()

        '<EhHeader>
        On Error GoTo Command16_Click_Err

        '</EhHeader>

100     Gdb.BeginTrans

        On Error GoTo Update_Error

        ' Populate it with backup data
110     strSQL = "update EID SET POS=999 " & "WHERE POS>=0 "
120     Gdb.Execute strSQL

        ' Update the new salary information
130     strSQL = "UPDATE EID " & "SET POS=998 " & "WHERE POS>=0 "
140     Gdb.Execute strSQL

150     MsgBox "Completed Successfully!" & vbCrLf & "Committing Changes.", , "Update"

160     Gdb.CommitTrans

        Exit Sub

Update_Error:

170     MsgBox "Error occured - " & Err.Description & vbCrLf & "Rolling Back the last transactions.", , "Update"

180     Gdb.RollbackTrans

        '<EhFooter>
        Exit Sub

Command16_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command16_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command16_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub Command17_Click()
'    Dim R As New ADODB.Recordset
'    Dim XX
'    Dim rEID As New ADODB.Recordset
'
'
'
'    R.Open "select * from EGGTIM WHERE MONTH(HME)=12 AND YEAR(HME)=2006 AND LEFT(ATIM,1) in ('G','g')", Gdb, adOpenDynamic, adLockOptimistic
'
'    R.MoveFirst
'
'
'    Do While Not R.EOF
'
'
'        FF = R("KODE")
'
'        rEID.Open "select * from EID WHERE KOD='" + FF + "'", Gdb, adOpenDynamic, adLockOptimistic
'        If rEID.EOF Then
'            XX = "INSERT INTO EID (KOD,ONO,FPA) VALUES ('" + R("kode")
'
'
'            XX = XX + "','" + R("ONOma") + "' ,"
'            XX = XX + Str(R("FPA")) + ")"
'            Gdb.Execute XX
'
'        End If
'        rEID.Close
'
'        R.MoveNext
'    Loop
'

'
'
'
'End Sub

Private Sub Command18_Click()
'δειγμα πως δουλευουν οι stored procedures


        '<EhHeader>
        On Error GoTo Command18_Click_Err

        '</EhHeader>
        Dim OB As ADODB.Command

        Dim N  As Integer

100     N = 1

        Dim PR() As Variant

110     ReDim PR(N)

        On Error Resume Next

120     Set OB = New ADODB.Command
130     OB.ActiveConnection = Gdb
140     OB.CommandType = adCmdStoredProc
150     OB.CommandText = STORED.Text

160     PR(0) = (PAR1.Text)    ' 01/01/2006

170     PR(1) = (PAR2.Text)    ' 12/31/2006

        'PR(2) = "'T','t'"
        'PR(3) = "'p','P'"
        'PR(4) = "'g','G'"
        'PR(5) = "'D'"

        Dim R As ADODB.Recordset

        'ΑΝ ΘΕΛΩ ΝΑ ΤΡΑΒΑΩ ΣΕ RECORDSET
180     Set R = OB.Execute(, PR)

190     R.Close   ' 2o SELECT

200     PR(0) = (PAR1.Text)    ' 01/01/2006
210     PR(1) = (PAR2.Text)    ' 12/31/2006
220     Set R = OB.Execute(, PR)

        Dim TT As Long

230     TT = GetCurrentTime()

        On Error GoTo ERRORS

        '   OB.Execute , PR
240     MsgBox GetCurrentTime() - TT

250     Set OB = Nothing

        '
        'CREATE PROCEDURE [APOT1]
        '@PAR1  VARCHAR(255) , @PAR2  VARCHAR(255)
        ' AS
        '
        'DROP TABLE DOKEGGT1
        'SELECT  KODE,
        'SUM( CASE APOT  WHEN 1  THEN XRE ELSE 0  END ) AS S1X,
        'SUM(CASE APOT  WHEN 1  THEN PIS ELSE 0  END ) AS S1P,
        'SUM( CASE APOT  WHEN 2 THEN XRE ELSE 0  END ) AS S2X,
        'SUM( CASE APOT WHEN 2 THEN PIS ELSE 0 END ) AS S2P
        ' Into DOKEGGT1
        'FROM EGGTIM
        'where HME>=@PAR1   AND HME < @PAR2
        'GROUP BY KODE
        'GO
        Exit Sub

ERRORS:
260     MsgBox Err.Description + Chr(13) + "ΛΑΘΟΣ ΣΤΟΝ ΤΥΠΟ"

        'YPO ELEGXO
        'CREATE PROCEDURE [APOT1]
        '@PAR1  VARCHAR(255) , @PAR2  VARCHAR(255),
        '@PAR3 VARCHAR(255),@PAR4  VARCHAR(255) ,  /* POL ,  POLEPIS    'A','V'   */
        '@PAR5  VARCHAR(255),@PAR6 VARCHAR(255)    /* AGO , AGOEPIS   */
        '
        '
        ' AS
        '
        'DECLARE @X    VARCHAR(255)
        'SELECT @X=" 'T','t'  "
        '
        '
        'DROP TABLE DOKEGGT1
        '
        'SELECT  KODE,
        'SUM( CASE WHEN LEFT(ATIM,1)  IN  ( @X )  THEN  PIS   ELSE 0  END )   AS SP,
        'SUM( CASE WHEN LEFT(ATIM,1)  IN ( '+@PAR4+')   THEN   PIS    ELSE 0  END ) AS SEP,
        'SUM(  CASE WHEN LEFT(ATIM,1)  IN ( @PAR5)   THEN   XRE   ELSE 0  END ) AS SG,
        'SUM(  CASE WHEN LEFT(ATIM,1)  IN ( @PAR6)   THEN   XRE   ELSE 0  END ) AS SEG
        ' Into DOKEGGT1
        '
        'FROM EGGTIM
        '
        'where HME>=@PAR1   AND HME < @PAR2
        '
        'GROUP BY KODE
        'GO

        '<EhFooter>
        Exit Sub

Command18_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command18_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command18_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command19_Click()

        '<EhHeader>
        On Error GoTo Command19_Click_Err

        '</EhHeader>
        Dim mXML

100     mXML = "<Calendar CompactMode="

110     mXML = mXML + """"
120     mXML = mXML + "1"
130     mXML = mXML + """"
140     mXML = mXML + "/>"

        Dim strSQL As String

150     strSQL = "INSERT INTO CalendarEvents ("
160     strSQL = strSQL + "Subject, "
170     strSQL = strSQL + "Location, "
180     strSQL = strSQL + "Body, "
190     strSQL = strSQL + "StartDateTime,EndDateTime,"
200     strSQL = strSQL + "RecurrenceState,RemainderSoundFile,BusyStatus,IsAllDayEvent,ImportanceLevel,"
210     strSQL = strSQL + "IsPrivate,IsMeeting,LabelID,RExceptionEndTimeOrig,RExceptionStartTimeOrig,RecurrencePatternID,CustomPropertiesXMLData,IsReminder,ReminderMinutesBeforeStart)  "
220     strSQL = strSQL + "VALUES('Subject Synantisi4','','','"

230     strSQL = strSQL + Format(Now, "mm/dd/yyyy hh:mm") + "','" + Format(DateAdd("h", 1, Now), "mm/dd/yyyy hh:mm") + "',"
240     strSQL = strSQL + "0,'',2,0,0,"
250     strSQL = strSQL + "0,0,3,'00:00:00','00:00:00',0,'" + mXML + "',0,15)"

        '   strSQL = strSQL & MakeSQLstr(pEvent.CustomProperties.SaveToString) & "); " & vbCrLf
260     Gdb.Execute strSQL

        ' <Calendar CompactMode="1"/>

        '    strSQL = "INSERT INTO CalendarEvents ("
        '    strSQL = strSQL + "Subject, "
        '    strSQL = strSQL + "Location, "
        '    strSQL = strSQL + "Body, "
        '
        '''''''''''''    strSQL = strSQL + "IsMeeting, "
        ''''''''''''''''    strSQL = strSQL + "IsPrivate,IsMeeting "
        '
        ''''''''    strSQL = strSQL + "LabelID, "
        '''''''''''''''''''    strSQL = strSQL + "BusyStatus, "
        '''''''''''''''''    strSQL = strSQL + "ImportanceLevel, "
        '
        '''''''    strSQL = strSQL + "StartDateTime, "
        '''''''''    strSQL = strSQL + "EndDateTime, "
        '
        ''''''''''    strSQL = strSQL + "IsAllDayEvent, "
        '
        '    strSQL = strSQL + "IsReminder, "
        '''''''''''''''''''    strSQL = strSQL + "IsReminder,ReminderMinutesBeforeStart, "
        '''''''''''''''    strSQL = strSQL + "RemainderSoundFile, "
        '
        '''''''''''''''''''''''''    strSQL = strSQL + "RecurrenceState, "
        '''''''''''''''''''''''    strSQL = strSQL + "RecurrenceState,RecurrencePatternID, "
        '
        '    strSQL = strSQL + " "
        ''''''''''''''''    strSQL = strSQL + "RExceptionEndTimeOrig,RExceptionStartTimeOrig, "
        '    strSQL = strSQL + "ISRecurrenceExceptionDeleted, "
        '
        '    strSQL = strSQL + "CustomPropertiesXMLData) " & vbCrLf
        '
        '    strSQL = strSQL + "VALUES("
        '
        '    strSQL = strSQL + MakeSQLstr(pEvent.Subject) & ", " & vbCrLf
        '    strSQL = strSQL + MakeSQLstr(pEvent.location) & ", " & vbCrLf
        '    strSQL = strSQL + MakeSQLstr(pEvent.body) & ", " & vbCrLf
        '
        '    strSQL = strSQL & IIf(pEvent.MeetingFlag, 1, 0) & ", " & vbCrLf
        '    strSQL = strSQL & IIf(pEvent.PrivateFlag, 1, 0) & ", " & vbCrLf
        '
        '    strSQL = strSQL & pEvent.Label & ", " & vbCrLf
        '    strSQL = strSQL & pEvent.BusyStatus & ", " & vbCrLf
        '    strSQL = strSQL & pEvent.Importance & ", " & vbCrLf
        '
        '    strSQL = strSQL & MakeSQLDateTime(pEvent.StartTime) & ", " & vbCrLf
        '    strSQL = strSQL & MakeSQLDateTime(pEvent.EndTime) & ", " & vbCrLf
        '
        '    strSQL = strSQL & IIf(pEvent.AllDayEvent, 1,

        '<EhFooter>
        Exit Sub

Command19_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command19_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command19_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command20_Click()

        '<EhHeader>
        On Error GoTo Command20_Click_Err

        '</EhHeader>

100     Gdb.Execute "SELECT SUM(POSO*TIMM) AS AJ2, SUM(MIK_AJIA) AS AJI ,ATIM,HME,APOT INTO EGGTIML FROM EGGTIM WHERE LEFT(ATIM,1)='L' GROUP BY ATIM,HME,APOT"
110     Gdb.Execute "INSERT INTO TIM (ATIM,HME,AJ2,AJI,B_N1) SELECT ATIM,HME,AJ2,AJI,APOT FROM EGGTIML"
120     Gdb.Execute "UPDATE TIM SET KPE='1111',EIDOS='e',AJ1=0,AJ3=0,AJ5=0,AJ4=0   WHERE LEFT(ATIM,1)='L'"

        '<EhFooter>
        Exit Sub

Command20_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command20_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command20_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub File1_Click()

        '<EhHeader>
        On Error GoTo File1_Click_Err

        '</EhHeader>
        Dim a

100     a = 0

        Dim F As String

110     F = File1.FILENAME

        Dim SS As String

        Dim B  As String

120     SS = ""

130     Open File1.Path + "\" + F For Input As #1

140     Do While Not EOF(1)

150         Line Input #1, B
160         SS = SS + B + Chr(13)
        Loop

170     Close #1

        'Text2.width = 4800
180     Text2.Text = SS

        '<EhFooter>
        Exit Sub

File1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.File1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.File1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.SHOW 1

        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Form_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then Command10_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Label1_Click()
'  Call Me.PopupMenu(ENERGA)

Dim ANS As Integer
ANS = MsgBox("Να γίνουν ολοι οι πελάτες ενεργοί;", vbYesNo)
If ANS = vbYes Then
   make_energoi "e"
End If
ANS = MsgBox("Να γίνουν ολοι οι προμηθευτές ενεργοί;", vbYesNo)
If ANS = vbYes Then
   make_energoi "r"
End If



End Sub

Private Sub List1_DblClick()

        '<EhHeader>
        On Error GoTo List1_DblClick_Err

        '</EhHeader>
        Dim k As Integer

100     For k = 0 To List1.ListCount - 1
110         List1.Selected(k) = True
        Next

        '<EhFooter>
        Exit Sub

List1_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.List1_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.List1_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub PERP_Click()

'===================================================================================================


'1   EPANAFERO TA DELTIA APOSTOLHS
Dim R As New ADODB.Recordset
'R.Open "SELECT * FROM EGGTIM WHERE ATIM>='T000048' AND ATIM<='T000460' ORDER BY ATIM,ID", Gdb, adOpenDynamic, adLockOptimistic
'Dim MATIM As String
'Dim ARXID As Long
'Do While Not R.EOF
'  MATIM = R!Atim
'  ARXID = R!ID
'  Do While MATIM = R!Atim
'     If R!ID - ARXID > 500 Then
'         Gdb.Execute "UPDATE EGGTIM SET ATIM2='A'+SUBSTRING(ATIM,2,6) WHERE ID=" + str(R!ID)
'     End If
'     R.MoveNext
'  Loop
'Loop
'META EDVSA
'UPDATE  EGGTIM SET  ATIM=ATIM2    WHERE ATIM2>='A000048'
'




'2  BAZO HMEROMHNIA STA DELTIA
'Dim R2 As New ADODB.Recordset
'Dim RT As New ADODB.Recordset
'R.Open "SELECT * FROM EGGTIM WHERE ATIM>='A000048' AND ATIM<='A000460' ORDER BY ATIM,ID", Gdb, adOpenDynamic, adLockOptimistic
'Dim MATIM As String
'Dim ARXID As Long
'
'Do While Not R.EOF
'  MATIM = R!Atim
'  ARXID = R!ID
'  R2.Open "SELECT * FROM TIM WHERE ATIM='" + MATIM + "'", Gdb, adOpenDynamic, adLockOptimistic
'
'  Do While MATIM = R!Atim And Not R.EOF
'     If Not R2.EOF Then
'         Gdb.Execute "UPDATE EGGTIM SET HME='" + Format(R2!hme, "MM/DD/YYYY") + "'  WHERE ID=" + str(R!ID)
'     End If
'     R.MoveNext
'  Loop
'  R2.Close
'
'
'DoEvents
'
'
'
'
'Loop
'
'
''3 BAZO TIS EKPTOSEIS POY LEIPOYN
'
'Dim RPEL As New ADODB.Recordset
'Dim REID As New ADODB.Recordset
'Dim RTIMOK As New ADODB.Recordset
'Dim e As String
'R.Open "SELECT * FROM EGGTIM WHERE LEFT(ATIM,1)='τ' AND HME>'03/18/2014' AND HME<'07/25/2014' ORDER BY ATIM,ID", Gdb, adOpenDynamic, adLockOptimistic
'Dim m As String
'
'Dim ARXID As Long
'
'Do While Not R.EOF
'     e = R!KODE
'     If IsNull(R!PELKOD) Then
'        RTIMOK.Open "SELECT * FROM TIM WHERE ATIM='" + R!Atim + "' AND HME='" + Format(R!hme, "MM/DD/YYYY") + "'", Gdb, adOpenDynamic, adLockOptimistic
'         m = RTIMOK!KPE
'         RTIMOK.Close
'
'     Else
'         m = R!PELKOD
'     End If
'     RPEL.Open "SELECT * FROM PEL WHERE EIDOS='e' AND KOD='" + m + "'", Gdb, adOpenDynamic, adLockOptimistic
'     If RPEL.EOF Then
'        MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ Ο ΠΕΛΑΤΗΣ " + m
'     End If
'
'     REID.Open "SELECT * FROM EID WHERE KOD='" + e + "'", Gdb, adOpenDynamic, adLockOptimistic
'
'     If REID.EOF Then
'        MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ to eidos " + e
'     End If
'
'     If IsNull(RPEL!ch4) Then
'        EP = "ΩΩΩ"
'     Else
'        EP = Trim$(RPEL!ch4)
'     End If
'
'     If IsNull(REID!CH1) Then
'        ee = "ΩΩΩ"
'     Else
'        ee = Trim$(REID!CH1)
'     End If
'
'
'
'
'     RTIMOK.Open "select * FROM TIMOKAT WHERE TIMOKPEL='" + EP + "' AND TIMOKEID='" + ee + "'", Gdb, adOpenDynamic, adLockOptimistic
'
'     If Not RTIMOK.EOF And R!EKPT = 0 Then
'           Gdb.Execute "UPDATE EGGTIM SET EKPT=" + Format(RTIMOK!EKPT, "000.00") + " , ATIM2='" + Format(RTIMOK!EKPT, "000.00") + "' WHERE ID=" + str(R!ID)
'     End If
'
'
'     RPEL.Close
'     REID.Close
'     RTIMOK.Close
'
'
'     DoEvents


'     R.MoveNext
'Loop
''META EDVSA
''UPDATE  EGGTIM SET  ATIM=ATIM2    WHERE ATIM2>='A000048'
''


'Exit Sub

'4 BAZO Tον πελατη (pelkod sto EGGTIM
Me.Caption = "4 BAZO Tον πελατη (pelkod sto EGGTIM"
Dim rPEL As New ADODB.Recordset
Dim REID As New ADODB.Recordset
Dim RTIMOK As New ADODB.Recordset
Dim e As String

R.Open "SELECT * FROM EGGTIM where PELKOD IS NULL ORDER BY ATIM", Gdb, adOpenDynamic, adLockOptimistic
    'LEFT(ATIM,1)='T' AND HME>'03/18/2014' AND HME<'07/25/2014' ORDER BY ATIM,ID"
Dim m As String

Dim ARXID As Long
Dim PELKODOK As Boolean
R.MoveFirst
Do While Not R.EOF
     PELKODOK = True
     e = R!KODE
     If IsNull(R!PELKOD) Then
        RTIMOK.Open "SELECT * FROM TIM WHERE ATIM='" + R!ATIM + "' AND HME='" + Format(R!hme, "MM/DD/YYYY") + "'", Gdb, adOpenDynamic, adLockOptimistic
        PELKODOK = False
         m = RTIMOK!KPE
         RTIMOK.Close

     Else
         m = R!PELKOD
     End If

     'RPEL.Open "SELECT * FROM PEL WHERE EIDOS='e' AND KOD='" + m + "'", Gdb, adOpenDynamic, adLockOptimistic
     'If RPEL.EOF Then
     '   MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ Ο ΠΕΛΑΤΗΣ " + m
     'End If

      REID.Open "SELECT * FROM EID WHERE KOD='" + e + "'", Gdb, adOpenDynamic, adLockOptimistic

     If REID.EOF Then
        ' MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ to eidos " + e
        GoTo HHH
     End If

    ' If IsNull(RPEL!ch4) Then
    '    EP = "ΩΩΩ"
    ' Else
    '    EP = Trim$(RPEL!ch4)
    ' End If

     If IsNull(REID!CH1) Then
        ee = "ΩΩΩ"
     Else
        ee = Trim(REID!CH1)
     End If




     'RTIMOK.Open "select * FROM TIMOKAT WHERE TIMOKPEL='" + EP + "' AND TIMOKEID='" + ee + "'", Gdb, adOpenDynamic, adLockOptimistic

    ' If Not RTIMOK.EOF And R!EKPT = 0 Then

     If PELKODOK = False Then
           Gdb.Execute "UPDATE EGGTIM SET PELKOD='" + m + "' , ATIM2='" + ee + "' WHERE ID=" + str(R!ID)
     Else
           Gdb.Execute "UPDATE EGGTIM SET  ATIM2='" + ee + "' WHERE ID=" + str(R!ID)
     End If


     ' RPEL.Close

     'RTIMOK.Close
HHH:
     Me.Caption = R!ATIM
     DoEvents

     REID.Close
     R.MoveNext
Loop


MsgBox "OK"












End Sub

Private Sub SELECTTIM_Click()
   Text2.Text = "SELECT TOP 10 HME,ATIM,AJI,ENTITYMARK,ENTITYUID,AKYROMENO,TRP,SKOPOS,JWT,ID_NUM FROM TIM ORDER BY ID_NUM DESC"
End Sub

Private Sub TDBGrid_HeadClick(ByVal ColIndex As Integer)

        '--------------------------------------------------------------------------------------=
        '<EhHeader>
        On Error GoTo TDBGrid_HeadClick_Err

        '</EhHeader>
        Dim sumes(100) As String

100     For k = 0 To Adodc2.Recordset.FIELDS.Count - 1
110         sumes(k) = TDBGrid.Splits(0).columns(k).FooterText    '  = Format(SUMES(k), "######0.00")
        Next

120     If Adodc2.Recordset.Sort = "[" & TDBGrid.columns(ColIndex).DataField & "] asc" Then   ' strSort
130         Adodc2.Recordset.Sort = "[" & TDBGrid.columns(ColIndex).DataField & "] desc"    ' strSort
        Else
140         Adodc2.Recordset.Sort = "[" & TDBGrid.columns(ColIndex).DataField & "] asc"    ' strSort
        End If

150     For k = 0 To Adodc2.Recordset.FIELDS.Count - 1
160         TDBGrid.Splits(0).columns(k).FooterText = sumes(k)   '  = Format(SUMES(k), "######0.00")
        Next

        '<EhFooter>
        Exit Sub

TDBGrid_HeadClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.TDBGrid_HeadClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.TDBGrid_HeadClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub TDBGrid_FilterChange()

        '<EhHeader>
        On Error GoTo TDBGrid_FilterChange_Err

        '</EhHeader>

        'Gets called when an action is performed on the filter bar

        On Error GoTo errHandler

100     Set ColS = TDBGrid.columns

        Dim c As Integer

110     c = TDBGrid.Col

120     TDBGrid.HoldFields

130     Adodc2.Recordset.Filter = getFilter()

140     TDBGrid.Col = c

150     TDBGrid.EditActive = True

        Exit Sub

errHandler:

160     MsgBox Err.Source & ":" & vbCrLf & Err.Description

170     Call cmdClearFilter_Click

        '<EhFooter>
        Exit Sub

TDBGrid_FilterChange_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.TDBGrid_FilterChange " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.TDBGrid_FilterChange " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Function getFilter() As String

        'Creates the SQL statement in adodc1.recordset.filter
        'and only filters text currently. It must be modified to filter other data types.
        '<EhHeader>
        On Error GoTo getFilter_Err

        '</EhHeader>

        Dim tmp As String

        Dim N   As Integer

100     For Each Col In ColS

110         If Trim(Col.FILTERTEXT) <> "" Then
120             N = N + 1

130             If N > 1 Then
140                 tmp = tmp & " AND "
                End If

150             If Adodc2.Recordset(Col.ColIndex).Type = 5 Then    '
160                 If InStr(">< >= <=  = ", Left(Col.FILTERTEXT, 1)) > 0 And Len(Col.FILTERTEXT) > 1 And IsNumeric(Right(Col.FILTERTEXT, 1)) Then
170                     tmp = tmp & Col.DataField & Col.FILTERTEXT
                    Else
180                     tmp = tmp & Col.DataField & " >  " & Val(Col.FILTERTEXT)
                    End If

                Else
190                 tmp = tmp & Col.DataField & " LIKE '" & Col.FILTERTEXT & "*'"
                End If

            End If

200     Next Col

210     getFilter = tmp

        '<EhFooter>
        Exit Function

getFilter_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.getFilter " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.getFilter " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub APOTSETIM_Click()

        '<EhHeader>
        On Error GoTo APOTSETIM_Click_Err

        '</EhHeader>

100     Gdb.Execute "CREATE VIEW dbo.[UPD_TIM] as " & " SELECT    dbo.TIM.B_N1 , dbo.EGGTIM.APOT " & "  FROM         dbo.TIM LEFT JOIN" & "  dbo.EGGTIM ON dbo.TIM.HME = dbo.EGGTIM.HME AND dbo.TIM.ATIM = dbo.EGGTIM.ATIM"

110     Gdb.Execute "UPDATE UPD_TIM  SET B_N1=APOT"

        '<EhFooter>
        Exit Sub

APOTSETIM_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.APOTSETIM_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.APOTSETIM_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command1_Click()
         runQuery
End Sub
         
         
     Private Sub runQuery()
        
        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>
        Dim DBF  As New ADODB.Connection

        Dim sql  As New ADODB.Connection

        Dim rDBF As New ADODB.Recordset

        Dim rSQL As New ADODB.Recordset

100     Open "C:\MERCVB\BOHU6LOG.TXT" For Append As #1
110     Write #1, Text2.Text

        Dim conDBF As String

        Dim conSQL As String

        On Error Resume Next

120     Me.MousePointer = vbHourglass

        Dim N As Long

130     N = GetCurrentTime()

140     If ODBC.Value = vbChecked Then

           If CheckReport.Value = vbChecked Then
               On Error Resume Next
           
              Dim strSource, strProvider As String
              strProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
              strSource = "Data Source=C:\MERCVB\REPORTS\REPORTS.mdb;" & "Persist Security Info=False"
             ' DATA2.Connect = strProvider + strSource
              
              Adodc2.ConnectionString = strProvider + strSource  ' gConnect
              Adodc2.RecordSource = Text2.Text
              Adodc2.Refresh
              
              
              
           
           Else
              ' DBGrid1.Visible = True
150             data2.Connect = "ODBC;" + gConnect
                data2.RecordSource = Text2.Text
170             data2.Refresh
           End If
           
160
        Else
            ' DBGrid1.Visible = False
180         Adodc2.ConnectionString = gConnect
190         Adodc2.RecordSource = Text2.Text
200         Adodc2.Refresh
        End If

210     Me.Caption = (GetCurrentTime() - N) / 1000

220     TDBGrid.AlternatingRowStyle = True

230     TDBGrid.OddRowStyle.BackColor = &H8000000F   ' GRI   vbCyan
240     TDBGrid.EvenRowStyle.BackColor = &HFFFFC0   'OYRANI     &H8000000F  ' GRI

250     Me.MousePointer = vbNormal

260     Close #1

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub update_sql_from_dbf(arxeio As String)

        '<EhHeader>
        On Error GoTo update_sql_from_dbf_Err

        '</EhHeader>
        Dim DBF    As Database

        Dim sql    As New ADODB.Connection

        Dim rDBF   As Recordset

        Dim rSQL   As New ADODB.Recordset

        Dim conDBF As String

        Dim conSQL As String

        Dim k      As Long

        Dim Fname  As String

100    ' G2 = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP4;Data Source=PLATO"
        '2004 G2 = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP4;Data Source=PLATO"
110     conSQL = gConnect    ' "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP;Data Source=LAGAKIS"
120     conDBF = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dBASE Files;Initial Catalog=C:\LAGEURO"
        '2004 conDBF = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dBASE Files;Initial Catalog=C:\LAGEURO\04"
'On Error GoTo 0
130     sql.Open conSQL
        'SET DBF=DATA.Open conDBF

140     data1.DatabaseName = Text1.Text

150     data1.RecordSource = arxeio
160     data1.Refresh

170     Set rDBF = data1.Recordset

180     Label1.Caption = arxeio + " " + Format(rDBF.RecordCount, "######")
190     sql.Execute "DELETE FROM " + arxeio

200     rSQL.Open "SELECT *FROM " + arxeio, sql, adOpenDynamic, adLockOptimistic
        ' rDBF.Open "SELECT *FROM " + arxeio, DBF, adOpenDynamic, adLockOptimistic


210     rDBF.MoveFirst

        Dim Z

220     Z = 0

        On Error GoTo printERROR    ' Resume Next
'On Error GoTo 0
230     Do While Not rDBF.EOF

240         rSQL.AddNew

250         For k = 0 To rDBF.FIELDS.Count - 1
260             Fname = rDBF.FIELDS(k).Name    ' p.x. FNAME=epo    R(0).NAME

270             If IsNull(rDBF(k)) Then
                Else
280                 rSQL(Fname) = rDBF(k)    ' rsqk("epo")=r(0)
                End If

                ' ΑΥΤΑ ΠΟΥ ΕΙΝΑΙ ΝULL ΤΑ ΓΕΜΙΖΕΙ
290             If IsNull(rDBF(k)) Then
300                 If rDBF(k).Type = dbText Then    '10=STRING
310                     rSQL(Fname) = " "
320                 ElseIf rDBF(k).Type = dbDate Then
330                     rSQL(Fname) = CDate("01/01/1990")
340                 ElseIf rDBF(k).Type = dbBoolean Then
350                     rSQL(Fname) = False
360                 ElseIf rDBF(k).Type = dbDouble Then
370                     rSQL(Fname) = 0
                    Else
380                     conSQL = ""
                    End If
                End If

            Next

            ' ΜΕΤΑΤΡΕΠΕΙ ΤΗΝ "ΧRΕOSI" ΣΕ "XRE" KAI "PISTOSI"
390         If arxeio = "EGG" Then
400             rSQL("PISTOSI") = 0: rSQL("XREOSI") = 0

410             If Not IsNull(rSQL("XRE")) Then
420                 If rSQL("EIDOS") = "e" Then
430                     If rSQL("XRE") > 0 Then rSQL("XREOSI") = rSQL("XRE")
440                     If rSQL("XRE") < 0 Then rSQL("PISTOSI") = -rSQL("XRE")
                    Else

450                     If rSQL("XRE") > 0 Then rSQL("PISTOSI") = rSQL("XRE")
460                     If rSQL("XRE") < 0 Then rSQL("XREOSI") = -rSQL("XRE")
                    End If
                End If
            End If

            ' METΑΤΡΕΠΕΙ ΤΟ 'POSO' ΣΕ 'XRE' KAI 'PIS'
470         If arxeio = "EGGTIM" Then
480             rSQL("PIS") = 0: rSQL("XRE") = 0

490             If Not IsNull(rSQL("POSO")) Then
500                 If InStr(F_xre_APOU + F_axre_APOU, Left(rSQL("ATIM"), 1)) > 0 Then
510                     rSQL("XRE") = rSQL("POSO")
                    End If

520                 If InStr(F_pis_APOU + F_apis_APOU, Left(rSQL("ATIM"), 1)) > 0 Then
530                     rSQL("PIS") = rSQL("POSO")
                    End If
                End If
            End If

540         rSQL.Update
550         Z = Z + 1

560         If Z Mod 10 = 0 Then
570             Me.Caption = Z
            End If

580         rDBF.MoveNext

590         DoEvents

        Loop

600     rSQL.Close
610     rDBF.Close


Gdb.Execute "UPDATE EGGTIM SET APOT=1, ID_NUM = (SELECT TOP 1 ID_NUM FROM TIM WHERE ATIM=EGGTIM.ATIM AND HME=EGGTIM.HME) WHERE ID_NUM IS NULL"
Gdb.Execute "UPDATE TIM SET B_N1=1 WHERE B_N1 IS NULL"
Gdb.Execute "UPDATE TIM SET TRP='1.ΜΕΤΡ' WHERE TRP LIKE 'ΜΕ%' "
Gdb.Execute "UPDATE TIM SET TRP='2.ΠΙΣΤ' WHERE TRP LIKE 'ΠΙ%' "


        Exit Sub

printERROR:
620     List2.Visible = True

630     List2.AddItem Err.Description

640     Resume Next

        '<EhFooter>
        Exit Sub

update_sql_from_dbf_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.update_sql_from_dbf " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.update_sql_from_dbf " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command10_Click()

        '<EhHeader>
        On Error GoTo Command10_Click_Err

        '</EhHeader>
    
        'Dim ANS
        'ANS = MsgBox("Να αποθηκευτούν οι αλλαγές μόνιμα;", vbYesNo)
        'If ANS = vbYes Then
        '    Gdb.CommitTrans
        '    MsgBox "ΟΚ Αποθηκεύτηκαν"
        ''    ANS = MsgBox("Οι εγγραφές που δημιουργήθηκαν είναι επίσης αποθηκευμένες στο " + mfil + Chr(13) + "Θέλετε να το ανοίξετε;", vbYesNo)
        ' '   If ANS = vbYes Then
        '  '      DUM = Shell("notepad.exe " + mfil, vbMaximizedFocus)
        '   ' End If
        'Else
        '    Gdb.RollbackTrans
        '    MsgBox "Δεν Αποθηκεύτηκαν"
        '
        'End If
        '
        '
    
100     Unload Me

110     If gApoMenu = True Then
120         SendKeys "%"
130         SendKeys "{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

Command10_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command10_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command10_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>
        Dim k As Integer

        Dim R As New ADODB.Recordset

100     ANS = MsgBox("Προσοχή θα σβηστούν τα αρχεία του SERVER. Είσαι σίγουρος;", vbYesNo)

110     If ANS = vbNo Then Exit Sub

        Dim MKOD As String

120     If Not ANS = vbYes Then
130         MsgBox " ΑΚΥΡΏΘΗΚΕ Η ΕΡΓΑΣΊΑ"

            Exit Sub

        End If

140     MKOD = InputBox("Αν είσαι σίγουρος δώσε κωδικό ")

150     If Val(MKOD) = 12345 Then
            'ok
        Else

            Exit Sub

        End If

160     R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic

170     F_xre_APOU = ""
180     F_axre_APOU = ""
190     F_pis_APOU = ""
200     F_apis_APOU = ""

210     R.MoveFirst

220     Do While Not R.EOF

230         If R("POS_APOU") = 1 Then
240             F_xre_APOU = F_xre_APOU + R("eidos")
250         ElseIf R("POS_APOU") = 2 Then
260             F_axre_APOU = F_axre_APOU + R("eidos")
270         ElseIf R("POS_APOU") = 3 Then
280             F_pis_APOU = F_pis_APOU + R("eidos")
290         ElseIf R("POS_APOU") = 4 Then
300             F_apis_APOU = F_apis_APOU + R("eidos")
            End If

310         R.MoveNext
        Loop

320     R.Close

330     For k = 0 To List1.ListCount - 1

340         If List1.Selected(k) Then
350             update_sql_from_dbf List1.List(k)
            End If

        Next

        Gdb.Execute "UPDATE TIM SET B_N1=1"



        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>
        Dim DBF As New ADODB.Connection

        Dim sql As New ADODB.Connection

        Dim rDBF As New ADODB.Recordset

        Dim rSQL As New ADODB.Recordset

        Dim Entolh As String

        '100 Open "C:\MERCVB\BOHU6LOG.TXT" For Append As #1
        '110 Write #1, Text2.Text
        Entolh = Text2.Text
        Gdb.Execute "INSERT INTO LOGGING (PROCED,ENTOLH,HME) VALUES ('bohu6','" + Replace(Entolh, "'", "''") + "',GETDATE())"

        Dim conDBF As String

        Dim conSQL As String

        Dim lo As Long

        Dim DB As Database

        Dim SEIRES(30)

        Dim MLINE As String, k

        Dim N As Long

120     For k = 1 To 30: SEIRES(k) = "": Next

130     lo = 0
140     Me.MousePointer = vbHourglass

        On Error GoTo LATOS  'On Error Resume Next

150     If ODBC.Value = Checked Then


          If CheckReport.Value = vbChecked Then
          
                Dim cnn           As ADODB.Connection

                Dim RDB           As ADODB.Recordset

                Dim strConnection As String, strProvider As String, strSource As String

                strProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
                strSource = "Data Source=C:\MERCVB\REPORTS\REPORTS.mdb;" '& App.Path &

                Set cnn = New ADODB.Connection
                strConnection = strProvider & strSource & "Persist Security Info=False"
                cnn.Open strConnection
                cnn.Execute Text2.Text
          
          
              ' DBF.Open "DSN=REPORTS"
              ' DBF.Execute Text2.Text
            

              ' DBGrid1.Visible = True         'Data2.Connect = "ODBC;" + gConnect
              'Data2.RecordSource = TEXT2.Text
              ' Data2.Refresh
          Else

160             Set DB = OpenDatabase("", False, False, gConnect)
170             DB.Execute Text2.Text
180             lo = DB.RecordsAffected
           End If

 
              
   









        Else
190         MLINE = Text2.Text

200         DUM = FETES2_DELIM(MLINE, SEIRES)

210         N = GetCurrentTime()

            ' Exit Sub
220         For k = 1 To 30

230             If Len(SEIRES(k)) > 2 Then
240                 Gdb.Execute Trim(SEIRES(k)), lo
                End If

            Next

250         Me.Caption = (GetCurrentTime() - N) / 1000
        End If

260     MsgBox str(lo) + " ΕΓΓΡΑΦΕΣ ΕΝΗΜΕΡΩΘΗΚΑΝ"
270     Me.MousePointer = vbNormal

280     ' Write #1, Str(lo) + " ΕΓΓΡΑΦΕΣ ΕΝΗΜΕΡΩΘΗΚΑΝ"
290     ' Close #1

        Exit Sub

LATOS:
300     MsgBox Err.Description

310     Resume Next

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function FETES2_DELIM(LINE, ELEM)

        '<EhHeader>
        On Error GoTo FETES2_DELIM_Err

        '</EhHeader>
        Dim KL, KE

        On Error GoTo MHNYMA

        '  DIABAZO SE PINAKA OLA TA STOIXEIA THS GRAMHS
100     For KE = 1 To 30
110         ELEM(KE) = ""
        Next

120     KL = 0   ' metraei xaraktires
130     KE = 0   ' metritis toy pinaka  ELEMENT

        Do
140         KE = KE + 1  ' metritis toy pinaka  ELEMENT
150         KL = KL + 1  ' metraei xaraktires

160         Do While mID$(LINE, KL, 1) <> ";"    ' tab
170             ELEM(KE) = ELEM(KE) + mID$(LINE, KL, 1)
180             KL = KL + 1  ' metraei xaraktires

190             If KL > Len(LINE) Then Exit Do
            Loop

200         KL = KL + 1    ' για να περασει το chr(10)

210         If KL > Len(LINE) Then Exit Do

220     Loop Until KL >= Len(LINE)    'OLO TO MHKOS THS GRAMMHS

230     FETES_DELIM = 0

        Exit Function

MHNYMA:
240     HandleError "Par1:Fetesdelim"

250     Resume Next

        '<EhFooter>
        Exit Function

FETES2_DELIM_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.FETES2_DELIM " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.FETES2_DELIM " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub Command4_Click()

        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>
        Dim sql  As New ADODB.Connection

        Dim rDBF As New ADODB.Recordset

        Dim rSQL As New ADODB.Recordset

100     rSQL.Open "SELECT * FROM EID WHERE APOS>0", Gdb, adOpenDynamic, adLockOptimistic
110     a = InputBox("Δώσε κωδικό παραστατικού", , "λ")

120     Gdb.Execute "DELETE  FROM EGGTIM WHERE LEFT(ATIM,1)='" + a + "';"

        Dim k As Long

130     rSQL.MoveFirst
140     k = 0

150     Do While Not rSQL.EOF

160         If Not IsNull(rSQL("APOS")) Then
170             k = k + 1
180             Gdb.Execute "INSERT INTO EGGTIM (EIDOS,PELKOD,HME,XRE,POSO,KODE,ATIM) VALUES ('e','0000','01/01/" + Format(Year(Now), "0000") + "'," + str(rSQL("APOS")) + "," + str(rSQL("APOS")) + ",'" + rSQL("KOD") + "','" + a + Format(k, "00000") + "')"
            End If

190         rSQL.MoveNext

200         If k Mod 10 = 0 Then DoEvents: Me.Caption = str(k)
        Loop

210     Gdb.Execute "UPDATE EID SET APOS=0"

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command5_Click()

        '<EhHeader>
        On Error GoTo Command5_Click_Err

        '</EhHeader>
        Dim DBF    As Database

        Dim sql    As New ADODB.Connection

        Dim rDBF   As Recordset

        Dim rSQL   As New ADODB.Recordset

        Dim conDBF As String

        Dim conSQL As String

        Dim lo     As Long

        Dim r2     As DAO.Recordset

100     Set DBF = OpenDatabase(Text1.Text, False, False, "dBase III;")
110     Set r2 = DBF.OpenRecordset("select LEFT(memo_key,5) AS KLEIDI,COUNT(*) from memotext where left(memo_key,1) in ('e','E','r') group by  LEFT(memo_key,5) order by  LEFT(memo_key,5) ")

120     rSQL.Open "select *from PEL", Gdb, adOpenDynamic, adLockOptimistic

        Dim a, B, c, D, e, F, G

130     r2.MoveFirst

140     Do While Not r2.EOF
150         a = Left(r2(0), 5)
160         Set R = DBF.OpenRecordset("select *from memotext where left(memo_key,5)='" + a + "' order by memo_key ")

170         If R.RecordCount > 0 Then
180             Me.Caption = a

190             DoEvents
200             R.MoveFirst
210             B = ""

220             Do While Not R.EOF
230                 B = B + R(1) + Chr$(13) + Chr$(10)    'memoline
240                 R.MoveNext
                Loop

250             If Not IsNull(B) Then
260                 B = Replace(B, "'", "-")

270                 If Not IsNull(B) Then Gdb.Execute "update PEL SET MEMO='" + B + "' WHERE EIDOS='" + Left(a, 1) + "' AND KOD='" + mID$(a, 2, 4) + "'"
                End If

280             DoEvents
            End If

290         r2.MoveNext
        Loop

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command6_Click()

        'F5 ΑΠΟΘΗΚΗς
        '<EhHeader>
        On Error GoTo Command6_Click_Err

        '</EhHeader>
        Dim DBF    As Database

        Dim sql    As New ADODB.Connection

        Dim rDBF   As Recordset

        Dim rSQL   As New ADODB.Recordset

        Dim conDBF As String

        Dim conSQL As String

        Dim lo     As Long

100     Set DBF = OpenDatabase(Text1.Text, False, False, "dBase III;")

110     rSQL.Open "select *from EID", Gdb, adOpenDynamic, adLockOptimistic

        Dim a, B, c, D, e, F, G

120     G = 0
130     rSQL.MoveFirst

140     Do While Not rSQL.EOF
150         G = G + 1
160         a = "a" + rSQL!kod

170         Set R = DBF.OpenRecordset("select *from memotext where left(memo_key,15)='" + a + "' order by memo_key ")

180         If R.RecordCount > 0 Then

190             R.MoveFirst
200             B = ""

210             Do While Not R.EOF
220                 B = B + R(1) + Chr$(13)    'memoline
230                 R.MoveNext
                Loop

240             rSQL!memo = B
250             rSQL.Update

260             DoEvents
            End If

270         G = G + 1

280         If G Mod 20 = 0 Then
290             Me.Caption = a

300             DoEvents
            End If

310         rSQL.MoveNext
        Loop

        '<EhFooter>
        Exit Sub

Command6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command7_Click()

        '------ FALCON ----------
        '<EhHeader>
        On Error GoTo Command7_Click_Err

        '</EhHeader>
        Dim CC   As String

        Dim lo   As Long

        Dim sql  As New ADODB.Connection

        Dim rDBF As New ADODB.Recordset

        Dim rSQL As New ADODB.Recordset

100     If Check2.Value = vbChecked Then
110         CC = "SELECT BARCODES.*,EID.KOD,EID.ERG AS ERG2,EID.ONO,EID.LTI5,EID.POS01,EID.POS02," & " EID.PROM,EID.XTI,EID.KODLOG,EID.AEG,EID.LTI,EID.POS03 " & " FROM BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD" & " WHERE BARCODES.ERG NOT LIKE '" + LTrim(str(F_DOROKARTA)) + "%'"

120         rSQL.Open CC, Gdb, adOpenDynamic, adLockOptimistic

130         If ODBC.Value = Checked Then

                Dim DB2 As Database

140             Set DB2 = OpenDatabase("", False, False, gConnect)

                On Error Resume Next

150             Kill "C:\OLA3"

                On Error GoTo Command7_Click_Err

                '          CC = "SELECT BARCODES.*,EID.ONO,EID.LTI5,EID.POS01,EID.POS02," _
                '          & " EID.PROM,EID.XTI,EID.KODLOG,EID.AEG,EID.LTI,EID.POS03 into OLA2 in 'C:\' 'Text;' " _
                '          & " FROM BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD" _
                '          & " WHERE BARCODES.ERG NOT LIKE '" + LTrim(Str(F_DOROKARTA)) + "%'"

160             CC = "SELECT BARCODES.KOD,BARCODES.ERG,EID.ONO,EID.LTI5,EID.POS " & "  into OLA4 in 'C:\' 'Text;' " & " FROM BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD" & " WHERE BARCODES.ERG NOT LIKE '" + LTrim(str(F_DOROKARTA)) + "%'"

                On Error Resume Next

170             Kill "C:\OLA4"
180             DB2.Execute CC
190             lo = DB2.RecordsAffected

200             MILSEC 2000

210             If Len(Dir("C:\OLA3.*")) = 0 Then
220                 MsgBox "ΔΕΝ ΒΡΙΣΚΩ ΤΟ ΑΡΧΕΙΟ ΤΩΝ ΕΙΔΩΝ ΓΙΑ ΑΠΟΣΤΟΛΗ", vbInformation

                    Exit Sub

                End If

230             MsgBox "ΔΗΜΙΟΥΡΓΗΘΗΚΕ ΤΟ ΑΡΧΕΙΟ C:\OLA3"
                ' Call RAPICopyPCFileToCE("C:\OLA3", "\OLA3")
                '     MsgBox "ΑΝΤΙΓΡΑΦΗΚΕ ΣΤΟ PDA  ΤΟ ΑΡΧΕΙΟ C:\OLA3"

                Exit Sub

            End If

        Else
240         rSQL.Open "SELECT * FROM EID where LEFT(KOD,5) NOT IN ('91393')", Gdb, adOpenDynamic, adLockOptimistic
        End If

        Dim DB As Database, R As Recordset

250     Set DB = OpenDatabase(F_MHXANI_FOLDER, False, False, "dBase III;")

        On Error Resume Next

        'On Error GoTo Command7_Click_Err
260     Kill Trim(F_MHXANI_FOLDER) + "\OLA.DBF"

270     DB.Execute "SELECT TOP 1 OLAA.* INTO OLA FROM OLAA"
280     Set R = DB.OpenRecordset("OLA")

        On Error Resume Next

        Dim k As Long

290     rSQL.MoveFirst
300     k = 0

        Dim OK As Boolean

310     For k = 0 To R.FIELDS.Count - 1

320         If R.FIELDS(k).Name = "XONDR" Then
330             OK = True
            End If

        Next

340     Do While Not rSQL.EOF
350         R.AddNew

360         If Not Check2.Value = vbChecked Then

370             R("BARCODE") = Left(rSQL("KOD"), 13)
380             R("LTI") = rSQL("LTI5")
390             R("NAME") = Left(rSQL("ONO"), 25)
400             R("POS01") = rSQL("POS01")
410             R("POS02") = rSQL("POS02")
420             R("PROM") = Left(rSQL("PROM"), R("PROM").Size)
430             R("XTI") = rSQL("XTI")
440             R("KODLOG") = rSQL("KODLOG")
450             R("CODE") = rSQL("ERG2")

460             If OK Then
470                 R("XONDR") = rSQL("LTI")
                End If

            Else

480             R("BARCODE") = Left(rSQL("ERG"), 13)
490             R("LTI") = rSQL("LTI5")
500             R("XTI") = rSQL("XTI")
510             R("CODE") = rSQL("ERG2")
520             R("NAME") = Left(rSQL("ONO"), 25)
530             R("POS01") = rSQL("POS01")
540             R("POS02") = rSQL("POS02")
550             R("POS03") = rSQL("POS03")

560             R("PROM") = Left(rSQL("PROM"), R("PROM").Size)
570             R("XTI") = rSQL("XTI")
580             R("KODLOG") = rSQL("KODLOG")

590             If OK Then
600                 R("XONDR") = rSQL("LTI")
                End If
            End If

610         R.Update
620         k = k + 1

630         If k Mod 100 = 0 Then
640             Me.Caption = k

650             DoEvents
            End If

            'Me.Caption = rSQL("kod")

660         rSQL.MoveNext

        Loop

670     Gdb.Execute "UPDATE EID SET APOS=0"

680     R.Close
690     DB.Close

700     DoEvents

        Dim DUM

        On Error Resume Next

710     DUM = Shell(F_MHXANI_FOLDER + "\send.bat", vbHide)

        '<EhFooter>
        Exit Sub

Command7_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command7_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command7_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command8_Click()

        ' METAFORA SE ARXEIO DBF
        '<EhHeader>
        On Error GoTo Command8_Click_Err

        '</EhHeader>
        Dim arxeio

        Dim a As Long

100     a = GetTickCount()
110     Me.Caption = "ENHMEΡΩΣΗ ΜΕ ΤΕΛΕΥΤΑΙΑ ΥΠΟΛΟΙΠΑ"

120     UPDATE_YPOLOIPA2

130     arxeio = "EID"

        Dim DBF    As Database

        Dim sql    As New ADODB.Connection

        Dim rDBF   As Recordset

        Dim rSQL   As New ADODB.Recordset

        Dim r2     As New ADODB.Recordset

        Dim conSQL As String

        Dim k      As Long

        Dim Fname  As String

        Dim DB     As DAO.Database

        'Set DB = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
        On Error Resume Next

140     Kill "C:\LAGEURO\EID.DBF"
150     Kill "C:\LAGEURO\BARCODES.DBF"

        On Error GoTo Command8_Click_Err

        Dim lo As Long, LO2 As Long

160     Me.Caption = "ΔΗΜΙΟΥΡΓΙΑ ΑΡΧΕΙΟΥ ΕΙΔΩΝ"

        Dim F_EXO_XROMATA, TEX As String

170     F_EXO_XROMATA = Val(FINDPARAMETROI(1, "PAR1", "F_EXO_XROMATA", "0", "Γ.1 ΧΡΗΣΙΜΟΠΟΙΩ ΜΕΓΕΘΗ-ΧΡΩΜΑΤΑ=1"))   'posa psifia tha exei h kathe seira

180     If F_EXO_XROMATA = 1 Then
190         rSQL.Open "select * FROM EID", Gdb, adOpenDynamic, adLockOptimistic
200         rSQL.MoveFirst

210         Do While Not rSQL.EOF

220             If rSQL("KOD") = "110002" Then
230                 LO2 = LO2   'DEBUG

                End If

240             Fname = "SELECT KODE,LEFT(PROELEYSH,5) AS PROEL ,APOT,(select PERIGRAFH FROM PINAKES WHERE TYPOS=4 AND AYJON=APOT) AS APOTHIKI,ROUND(SUM(XRE),0) AS [ΕΙΣ],ROUND(SUM(PIS),0) AS [ΕΞΑΓ]," & " round( (CASE WHEN SUM(XRE)IS NULL THEN 0 ELSE SUM(XRE) END )-" & "  (CASE WHEN SUM(PIS)IS NULL THEN 0 ELSE SUM(PIS) END ),1) AS YPOL "
250             Fname = Fname + " FROM EGGTIM  WHERE KODE='" + CNull(rSQL("KOD")) + "' GROUP BY APOT,KODE,LEFT(PROELEYSH,5) ORDER BY APOT,KODE,LEFT(PROELEYSH,5)"
260             r2.Open Fname, Gdb, adOpenDynamic, adLockOptimistic

                'r2.MoveFirst
270             TEX = ""

280             Do While Not r2.EOF

290                 If r2("KODE") = "110002" Then
300                     LO2 = LO2    'DEBUG
                    End If

310                 TEX = TEX + Left(CNull(r2("APOTHIKI")), 4) + " / " + Format(r2("YPOL"), "##0") + " / " + Left(CNull(r2("PROEL")), 5) + Space(44) + Chr(13)

320                 r2.MoveNext
                Loop

330             Gdb.Execute "UPDATE EID SET MEMO='" + TEX + "' WHERE KOD='" + CNull(rSQL("KOD")) + "'"
340             r2.Close

350             DoEvents

360             rSQL.MoveNext

            Loop

        End If

370     Set DB = OpenDatabase("", False, False, gConnect)

380     If XorisTimiAgoras = vbChecked Then
390         DB.Execute "SELECT KOD,ONO,ERG,LTI,LTI5,PONTOI,FPA,POS01," & " POS02,POS03,POS04,PROM,0 AS XTI,KODLOG,POS_KERD,POS_KERD2,SPA,MON,MEMO,KATHGORIA " & " INTO EID IN 'C:\LAGEURO' 'dBase III;' FROM EID"
        Else
400         DB.Execute "SELECT KOD,ONO,ERG,LTI,LTI5,PONTOI,FPA,POS01," & " POS02,POS03,POS04,PROM,XTI,KODLOG,POS_KERD,POS_KERD2,SPA,MON,MEMO,KATHGORIA " & " INTO EID IN 'C:\LAGEURO' 'dBase III;' FROM EID"
        End If

410     lo = DB.RecordsAffected

420     Me.Caption = "ΔΗΜΙΟΥΡΓΙΑ ΑΡΧΕΙΟΥ BARCODE"
430     DB.Execute "SELECT * " & " INTO BARCODES IN 'C:\LAGEURO' 'dBase III;' FROM BARCODES "
440     LO2 = DB.RecordsAffected

450     If Len(Dir("C:\LAGEURO\EIDFTP.bat")) > 1 Then
460         MILSEC 1000
470         DUM = Shell("C:\LAGEURO\EIDFTP.bat", vbMaximizedFocus)

        End If

480     MsgBox "Εστάλησαν " + str(lo) + " είδη και " + str(LO2) + " barcodes "

        Exit Sub

        '----------------------------------------------------------------------------------------------

490     sql.Open gConnect
500     DB.Execute "SELECT TOP 1 EID0.* INTO EID  FROM EID0"

510     Set rDBF = DB.OpenRecordset("EID")

        'Label1.Caption = ARXEIO + " " + Format(rDBF.RecordCount, "######")

        'SQL.Execute "DELETE FROM " + arxeio

520     rSQL.Open "SELECT * FROM EID ", Gdb, adOpenForwardOnly, adLockReadOnly

        ' rDBF.Open "SELECT *FROM " + arxeio, DBF, adOpenDynamic, adLockOptimistic

530     rSQL.MoveFirst

        Dim Z

540     Z = 0

        On Error GoTo printERROR    ' Resume Next

550     Do While Not rSQL.EOF

560         rDBF.AddNew
570         rDBF("KOD") = rSQL("KOD")
580         rDBF("ONO") = rSQL("ONO")
590         rDBF("ERG") = rSQL("ERG")
600         rDBF("LTI") = rSQL("LTI")
610         rDBF("LTI5") = rSQL("LTI5")
620         rDBF("pontoi") = rSQL("pontoi")
630         rDBF("FPA") = rSQL("FPA")
640         rDBF("POS01") = rSQL("POS01")
650         rDBF("POS02") = rSQL("POS02")

660         rDBF("prom") = rSQL("prom")
670         rDBF("xti") = rSQL("xti")
680         rDBF("KODLOG") = rSQL("KODLOG")
690         rDBF("KATHGORIA") = rSQL("KATHGORIA")

700         rDBF("pOS_KERD") = rSQL("pOS_KERD")
710         rDBF("POS_KERD2") = rSQL("POS_KERD2")
720         rDBF("SPA") = rSQL("SPA")
730         rDBF("MON") = rSQL("MON")

740         rDBF.Update

750         Z = Z + 1

760         If Z Mod 10 = 0 Then
770             Me.Caption = Z
            End If

780         rSQL.MoveNext

790         DoEvents

        Loop

800     rSQL.Close
810     rDBF.Close

820     Me.Caption = (GetTickCount() - a) / 1000

        Exit Sub

printERROR:
830     List2.Visible = True

840     List2.AddItem Err.Description

850     Resume Next

        '<EhFooter>
        Exit Sub

Command8_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command8_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command8_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command9_Click()

        ' apo SQL SERVER ΣΕ DBF ΑΡΧΕΙΑ ΠΟΥ ΗΔΗ ΥΠΑΡΧΟΥΝ
        '<EhHeader>
        On Error GoTo Command9_Click_Err

        '</EhHeader>
        Dim k As Integer

        Dim R As New ADODB.Recordset

100     ANS = MsgBox("Προσοχή θα σβηστούν τα αρχεία DBF. Είσαι σίγουρος;", vbYesNo)

110     If ANS = vbNo Then Exit Sub

        Dim MKOD As Integer

120     MKOD = InputBox("Αν είσαι σίγουρος δώσε κωδικό ")

130     If Val(MKOD) = 12345 Then
            'ok
        Else

            Exit Sub

        End If

140     R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic

150     F_xre_APOU = ""
160     F_axre_APOU = ""
170     F_pis_APOU = ""
180     F_apis_APOU = ""

190     R.MoveFirst

200     Do While Not R.EOF

210         If R("POS_APOU") = 1 Then
220             F_xre_APOU = F_xre_APOU + R("eidos")
230         ElseIf R("POS_APOU") = 2 Then
240             F_axre_APOU = F_axre_APOU + R("eidos")
250         ElseIf R("POS_APOU") = 3 Then
260             F_pis_APOU = F_pis_APOU + R("eidos")
270         ElseIf R("POS_APOU") = 4 Then
280             F_apis_APOU = F_apis_APOU + R("eidos")
            End If

290         R.MoveNext
        Loop

300     R.Close

310     For k = 0 To List1.ListCount - 1

320         If List1.Selected(k) Then
330             update_DBF_from_SQL List1.List(k)
            End If

        Next

        '<EhFooter>
        Exit Sub

Command9_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Command9_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Command9_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
'ENERGA.Visible = False

100     mForm_Load Me, fh, fw, ft, fl
110     Me.Picture = LoadPicture(gPicture)
120     List1.AddItem "PEL"
130     List1.AddItem "EID"
140     List1.AddItem "EGG"
150     List1.AddItem "EGGTIM"
160     List1.AddItem "TIM"
170     List1.AddItem "GRA"
180     List1.AddItem "GEFYRES"
 List1.AddItem "PARASTAT"
        On Error Resume Next

190     F_DOROKARTA = Val(FINDPARAMETROI(1, "BOHU6", "F_DOROKARTA", "9139393", "ΟΙ ΔΩΡΟΚΑΡΤΕΣ ΑΡΧΙΖΟΥΝ ΑΠΟ "))

200     F_MHXANI_FOLDER = FINDPARAMETROI(1, "BOHU6", "F_MHXANI_FOLDER", "C:\LAGEURO\MHXANAKI", "ΦΑΚΕΛΟΣ PDA ")

        '   'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""

        '    'SkinFramework1.ApplyWindow Me.hWnd
        '    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

        ' Gdb.BeginTrans


F_POS1_FOLDER = Trim(FINDPARAMETROI(1, "BOHU6", "F_POS1_FOLDER", "", "ΦΑΚΕΛΟΣ POS1 "))
F_POS2_FOLDER = Trim(FINDPARAMETROI(1, "BOHU6", "F_POS2_FOLDER", "", "ΦΑΚΕΛΟΣ POS2 "))
F_POS3_FOLDER = Trim(FINDPARAMETROI(1, "BOHU6", "F_POS3_FOLDER", "", "ΦΑΚΕΛΟΣ POS3"))


Dir1.Path = "c:\mercvb\queries"


 

210     Me.KeyPreview = True

        On Error Resume Next

220     File1.Path = "C:\mercvb\queries"

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub update_DBF_from_SQL(arxeio As String)

        '<EhHeader>
        On Error GoTo update_DBF_from_SQL_Err

        '</EhHeader>

        Dim DBF    As Database

        Dim sql    As New ADODB.Connection

        Dim rDBF   As Recordset

        Dim rSQL   As New ADODB.Recordset

        Dim conDBF As String

        Dim conSQL As String

        Dim k      As Long

        Dim Fname  As String

        Dim DB     As DAO.Database

100     Set DB = OpenDatabase(Text1.Text, False, False, "dBase III;")
110     DB.Execute "delete * from " + arxeio
120     Set rDBF = DB.OpenRecordset(arxeio)

130     Label1.Caption = arxeio + " " + Format(rDBF.RecordCount, "######")

140     rSQL.Open "SELECT *FROM " + arxeio, sql, adOpenDynamic, adLockOptimistic
        ' rDBF.Open "SELECT *FROM " + arxeio, DBF, adOpenDynamic, adLockOptimistic

150     rSQL.MoveFirst

        Dim Z

160     Z = 0

        On Error GoTo printERROR    ' Resume Next

170     Do While Not rSQL.EOF

180         rDBF.AddNew

190         For k = 0 To rDBF.FIELDS.Count - 1
200             Fname = rDBF.FIELDS(k).Name    ' p.x. FNAME=epo    R(0).NAME

                '      If IsNull(rDBF(k)) Then
                '      Else
210             rDBF(Fname) = rSQL(Fname)    ' rsqk("epo")=r(0)
                '      End If
            Next

220         rDBF.Update
230         Z = Z + 1

240         If Z Mod 10 = 0 Then
250             Me.Caption = Z
            End If

260         rSQL.MoveNext

270         DoEvents

        Loop

280     rSQL.Close
290     rDBF.Close

        Exit Sub

printERROR:
300     List2.Visible = True

310     List2.AddItem Err.Description

320     Resume Next

        '<EhFooter>
        Exit Sub

update_DBF_from_SQL_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.update_DBF_from_SQL " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.update_DBF_from_SQL " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub TDBGrid_GroupColMove(ByVal Position As Integer, _
                                 ByVal ColIndex As Integer, _
                                 Cancel As Integer)

        '<EhHeader>
        On Error GoTo TDBGrid_GroupColMove_Err

        '</EhHeader>
        Dim strSort As String

        Dim Col     As TrueOleDBGrid80.Column

        ' Loop through GroupColumns collection and construct

        ' the sort string for the Sort property of the Recordset

100     For Each Col In TDBGrid.GroupColumns

110         If strSort <> vbNullString Then

120             strSort = strSort & ", "

            End If

130         strSort = strSort & "[" & Col.DataField & "]"

140     Next Col

150     TDBGrid.HoldFields

160     Adodc2.Recordset.Sort = strSort

        '<EhFooter>
        Exit Sub

TDBGrid_GroupColMove_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.TDBGrid_GroupColMove " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.TDBGrid_GroupColMove " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function mac(ByVal s As String) As String

        'δινω ενα στρινγκ και μου επιστρέφει μία τιμή
        '<EhHeader>
        On Error GoTo mac_Err

        '</EhHeader>
        Dim X

        On Error GoTo mac_Err

100     MACtEXT.Text = "   "

        On Error GoTo err2

110     X = "MACtEXT.Text = " + s
120     fSCR.ExecuteStatement X

130     If F928 = 1 Then
140         mac = MACtEXT.Text
        Else
150         mac = to437(MACtEXT.Text)
        End If

        Exit Function

err2:
160     Me.Caption = X
170     mac = "  "

180     Resume Next

        '<EhFooter>
        Exit Function

mac_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.mac " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.mac " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function load2_forma(arxeio As String, spacing)

        'spacing=1 klassiko poy ta ypologizei me +1 thn stili toy pedioy (x1)
        'spacing=0 kanonika
        '*********************************************************
        '-----------    GLOBAL ΜΕΤΑΒΛΗΤΕΣ  ----------------
        'gm_str() οι σειρές του f99.txt
        'gpic(i) το στρινγκ του πεδίου
        'gm_r(i),gm_c(i) σειρά και στήλη του πεδίου
        'gm_f(i) το όνομα του πεδίου
        '<EhHeader>
        On Error GoTo load2_forma_Err

        '</EhHeader>

        'gm_str(6)="     ^XXXXXXXXXXXXXXXXXXXXXXXXXXX               ^xxxxxxxx    @XXXXXXXX  ~ono_par  ~SUBS(tim->atim,2,5)  ~TIM->HME
        '        gpic(2)='XXXXXXXXXXXXXXXXXXXXXXXXXXX
        'gm_r(2)=5  gm_c(2)=6    gm_f(2)=ono_par
        '========================================================
        'Dim gm_str(1 To 250) As String
        Dim ar_ped(1 To 250) As Integer

        Dim k                As Integer, i As Integer, lastseir As Integer

        Dim npic             As Integer, m_npic As Integer, L1 As Integer, N As Integer

        Dim a(1 To 250)      As Integer

        Dim u                As Integer, mhk_seir As Integer

        Dim X1               As Integer, X2 As Integer

        Dim xa(1 To 30)      As Integer

        '-----------------------------------
        On Error GoTo load2_forma_Err

        'Close #1

        'Close #1
100     Open arxeio For Input As #1

110     For k = 1 To 250
120         gm_str(k) = ""
130         gm_f(k) = ""
140         gpic(k) = " "
150         gm_r(k) = 0
160         gm_c(k) = 0
170         ar_ped(k) = 0
        Next

180     i = 1

190     Do While Not EOF(1)
200         Line Input #1, gm_str(i)
210         i = i + 1
        Loop

220     Close #1
        '----------------------------------

230     lastseir = 0
240     npic = 0
        ' ΔΙΑΒΑΖΩ ΤΑ PICTURES ΚΑΙ ΤΗΝ ΘΕΣΗ ΤΟΥΣ

        '*-------------------------------------------------------------------
250     For k = 1 To 250
260         L1 = InStr(gm_str(k), "^")

270         For N = 1 To 30
280             a(N) = 0
            Next

            ' * δεν έχει καθόλου ^
290         If L1 = 0 Then
300             ar_ped(k) = 0
310         ElseIf Left(gm_str(k), 3) = "^^^" Then   'SXOLIA
320             npic = npic + 1
330             ar_ped(k) = 0
            Else
                '* βρέθηκε σειρά με ^ , ψάχνω μήπως έχει και άλλα
340             a(1) = L1    '&& a[]  θέσεις όπου βρέθηκαν τα ^

350             For u = 2 To 30
360                 a(u) = InStr(mID$(gm_str(k), a(u - 1) + 1, Len(gm_str(k))), "^")

370                 If a(u) = 0 Then

                        Exit For

                    End If

380                 a(u) = a(u) + a(u - 1)
                Next

                ' μηκος σειράς
390             mhk_seir = InStr(gm_str(k), "~")

400             If mhk_seir = 0 Then
410                 MsgBox "δεν έχω το σημάδι ~ στην σειρά" + str(k)

                    Exit Function

                End If

420             ar_ped(k) = u - 1

                'ΣΕ ΑΥΤΗΝ ΤΗΝ ΣΕΙΡΑ ΒΡΕΘΗΚΑΝ U-1 ΠΕΔΙΑ

430             For i = 1 To u - 1
440                 npic = npic + 1  'ΑΥΞΑΝΕΙ Ο ΑΡΙΘΜΟΣ ΤΩΝ ΠΕΔΙΩΝ

                    'If spacing = 1 Then  ' bgazei +1  (λανθασμένο κρατείται για συμβατότητα)
                    '    x1 = A(i) + 1  '&&  if ( i=1,1,a(i))
                    '    x2 = IIf(A(i + 1) = 0, (mhk_seir - 1) - x1 + 1, A(i + 1) - 1 - x1)
                    'Else
450                 X1 = a(i)   '&&  if ( i=1,1,a(i))
460                 X2 = IIf(a(i + 1) = 0, (mhk_seir - 1) - X1 + 1, a(i + 1) - X1)
                    'End If

                    'το PICTURE του πεδίου
470                 gpic(npic) = mID$(gm_str(k), X1, X2)

                    'η σειρά του πεδίου
480                 gm_r(npic) = IIf(i = 1, k - lastseir, 0)    'αφου είναι στην ίδια σειρά να μην προσθέτει σειρές

                    'η στήλη του πεδίου
490                 gm_c(npic) = a(i)
                Next

500             lastseir = k
            End If

        Next

510     m_npic = npic

520     npic = 0

530     For k = 1 To 250
540         L1 = InStr(gm_str(k), "~")

550         For N = 1 To 30
560             xa(N) = 0
            Next

570         If Left(gm_str(k), 3) = "^^^" Then
580             npic = npic + 1
590         ElseIf L1 = 0 Then        ' &&  ›¤ β®  ΅ζΆ¦¬ ~

            Else
                'βρέθηκε σειρά με ~ , ψάχνω μήπως έχει και άλλα
600             xa(1) = L1    'a[]  θέσεις όπου βρέθηκαν τα ~

610             For u = 2 To 30
620                 xa(u) = InStr(mID$(gm_str(k), xa(u - 1) + 1, Len(gm_str(k))), "~")

630                 If xa(u) = 0 Then

                        Exit For

                    End If

640                 xa(u) = xa(u) + xa(u - 1)
                Next

                'μηκος σειράς
650             mhk_seir = Len(gm_str(k))

                'ΣΕ ΑΥΤΗΝ ΤΗΝ ΣΕΙΡΑ ΒΡΕΘΗΚΑΝ U-1 ΠΕΔΙΑ

660             For i = 1 To u - 1
670                 npic = npic + 1  'ΑΥΞΑΝΕΙ Ο ΑΡΙΘΜΟΣ ΤΩΝ ΠΕΔΙΩΝ
680                 X1 = xa(i) + 1  '&&  if ( i=1,1,a(i))
690                 X2 = IIf(xa(i + 1) = 0, (mhk_seir) - X1 + 1, xa(i + 1) - 1 - X1)

                    '* o titlow του πεδίου
700                 gm_f(npic) = mID$(gm_str(k), X1, X2)

                Next

710             If ar_ped(k) <> u - 1 Then
720                 If ar_ped(k) > u - 1 Then
730                     MsgBox "στην σειρά " + Format(k, "##") + Left(gm_str(k), 40) + ".... εχω παραπάνω  " + str(ar_ped(k) - (u - 1)) + "^  από  ~"
                    Else
740                     MsgBox " στην σειρά " + Format(k, "##") + " " + Left(gm_str(k), 40) + ".... έχω παραπάνω " + str(-ar_ped(k) + (u - 1)) + " ~ από ^ "
                    End If
                End If

750             lastseir = k
            End If

        Next

        '<EhFooter>
        Exit Function

load2_forma_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.load2_forma " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.load2_forma " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

