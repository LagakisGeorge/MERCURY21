VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form apot3 
   BackColor       =   &H00FF0000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Παρουσ.Ειδών"
   ClientHeight    =   11610
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   19260
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11610
   ScaleWidth      =   19260
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   42
      Top             =   3240
      Width           =   2385
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "A/A στην εκτύπωση"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2040
      TabIndex        =   41
      Top             =   2880
      Width           =   1905
   End
   Begin VB.CommandButton YPOLOGISMOS 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Υπολογισμός"
      Height          =   330
      Left            =   1950
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Για να υπολογισθούν σωστά τα Υπόλοιπα πρέπει οι ημερομηνίες να είναι από την αρχή της χρονιάς μέχρι σήμερα"
      Top             =   3720
      Width           =   2355
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Εξοδος"
      Height          =   330
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3720
      Width           =   2000
   End
   Begin VB.ListBox List11 
      Height          =   255
      Left            =   4200
      MultiSelect     =   1  'Simple
      TabIndex        =   29
      Top             =   1080
      Width           =   2772
   End
   Begin VB.ListBox AGNOSTOI 
      Height          =   255
      Left            =   3600
      TabIndex        =   28
      Top             =   10320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Ενημέρωση με τελευταίο προμηθευτη"
      Height          =   300
      Left            =   13440
      TabIndex        =   27
      Top             =   2880
      Width           =   3465
   End
   Begin VB.CommandButton aritmos_pelaton 
      Caption         =   "Αριθμός Πελατών"
      Height          =   300
      Left            =   13440
      TabIndex        =   26
      Top             =   1080
      Width           =   3465
   End
   Begin VB.CommandButton ISOZYGIO 
      BackColor       =   &H8000000A&
      Caption         =   "ΙΣΟΖΥΓΙΟ ΑΠΟ ΕΩΣ"
      Height          =   300
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   720
      Width           =   3465
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Μέση Τιμ.Αγοράς"
      Height          =   300
      Left            =   13440
      TabIndex        =   24
      Top             =   1800
      Width           =   3465
   End
   Begin VB.CommandButton telkin 
      Caption         =   "Τελευταία Κίνηση"
      Height          =   300
      Left            =   13440
      TabIndex        =   22
      Top             =   2520
      Width           =   3465
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ενημ.τελ.τιμή αγοράς&.ημερ.αγοράς"
      Height          =   300
      Left            =   13440
      TabIndex        =   21
      Top             =   2160
      Width           =   3465
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H8000000A&
      Caption         =   "EKTYΠΩΣΗ ΟΛΩΝ ΤΩΝ ΑΠΟΘΗΚΩΝ"
      Height          =   300
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   360
      Width           =   3465
   End
   Begin VB.CommandButton Command9 
      Caption         =   "ΕΚΤΥΠΩΣΗ ΚΕΡΔΩΝ ΧΩΡΙΣ ΤΗΝ ΠΥ"
      Height          =   435
      Left            =   7320
      TabIndex        =   19
      Top             =   9750
      Visible         =   0   'False
      Width           =   2730
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3600
      Top             =   9360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Αποθήκευση Report"
      Height          =   435
      Left            =   11505
      TabIndex        =   18
      Top             =   1965
      Width           =   1530
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   8835
      TabIndex        =   17
      Top             =   480
      Width           =   4185
   End
   Begin VB.TextBox Text4 
      Height          =   1005
      Left            =   8880
      MultiLine       =   -1  'True
      TabIndex        =   16
      ToolTipText     =   "Για σύνολα στα πεδία 4,5 βάζω στην 1η σειρά :   /* ΣΥΝΟΛΑ 000110  */   "
      Top             =   2460
      Width           =   4185
   End
   Begin VB.CheckBox rCheck 
      Caption         =   "Χρησιμοποιώ το επιλεγμενο report"
      Height          =   405
      Left            =   9330
      TabIndex        =   15
      Top             =   1995
      Width           =   2115
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "apot3-orig.frx":0000
      Left            =   7320
      List            =   "apot3-orig.frx":000D
      TabIndex        =   14
      Top             =   9405
      Visible         =   0   'False
      Width           =   2730
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8040
      TabIndex        =   13
      Text            =   "99"
      Top             =   2280
      Width           =   315
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Είδη σε έλλειψη"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   12720
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   9840
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Εκτύπωση"
      Height          =   255
      Left            =   10170
      TabIndex        =   8
      Top             =   9915
      Visible         =   0   'False
      Width           =   1935
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4800
      Top             =   9840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\mercvb\reports\dokeggt3.rpt"
      UserName        =   "sa"
   End
   Begin VB.ComboBox ORDERCOMBO 
      Height          =   315
      ItemData        =   "apot3-orig.frx":0044
      Left            =   7080
      List            =   "apot3-orig.frx":0054
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ΕΚΤΥΠΩΣΗ ΚΑΤΑ ΑΠΟΘΗΚΗ"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      ToolTipText     =   "Εκτύπωση όλων τωναποθηκών (αγορές,πωλήσεις,υπόλοιπα ανά αποθήκη) . Διαλέγω τα κριτήρια αυτήν την στιγμή (όχι αποθηκευμένα) . "
      Top             =   10320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc sqlcommand 
      Height          =   330
      Left            =   0
      Top             =   9360
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "sqlcommand"
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
   Begin VB.CommandButton Command5 
      Caption         =   "Κριτήρια επιλογής"
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      Top             =   3000
      Width           =   3735
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      ItemData        =   "apot3-orig.frx":008C
      Left            =   240
      List            =   "apot3-orig.frx":008E
      MultiSelect     =   1  'Simple
      TabIndex        =   3
      Top             =   960
      Width           =   3696
   End
   Begin MSAdodcLib.Adodc egg 
      Height          =   330
      Left            =   0
      Top             =   9240
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   240
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Εκτύπωση"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1200
   End
   Begin VB.ComboBox Combo2 
      Height          =   288
      ItemData        =   "apot3-orig.frx":0090
      Left            =   240
      List            =   "apot3-orig.frx":00AC
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   360
      Width           =   3696
   End
   Begin TrueOleDBGrid80.TDBGrid GR2 
      Bindings        =   "apot3-orig.frx":0143
      Height          =   5400
      Left            =   45
      TabIndex        =   7
      Top             =   4095
      Width           =   16995
      _ExtentX        =   29977
      _ExtentY        =   9525
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
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   49
      Columns(5)._MaxComboItems=   5
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).AllowColMove=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(5).Button=1"
      Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(26)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(30)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).NoClipping=   -1  'True
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos(0).NeedTotalPageCount=   -1  'True
      PrintInfos.Count=   1
      ColumnFooters   =   -1  'True
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DataView        =   2
      GroupByCaption  =   ""
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=66,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=63,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=64,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=65,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
      _StyleDefs(68)  =   "Named:id=33:Normal"
      _StyleDefs(69)  =   ":id=33,.parent=0"
      _StyleDefs(70)  =   "Named:id=34:Heading"
      _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   ":id=34,.wraptext=-1"
      _StyleDefs(73)  =   "Named:id=35:Footing"
      _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=36:Selected"
      _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=37:Caption"
      _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(79)  =   "Named:id=38:HighlightRow"
      _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=39:EvenRow"
      _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(83)  =   "Named:id=40:OddRow"
      _StyleDefs(84)  =   ":id=40,.parent=33"
      _StyleDefs(85)  =   "Named:id=41:RecordSelector"
      _StyleDefs(86)  =   ":id=41,.parent=34"
      _StyleDefs(87)  =   "Named:id=42:FilterBar"
      _StyleDefs(88)  =   ":id=42,.parent=33"
   End
   Begin MSComCtl2.DTPicker apo 
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   360
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   308019201
      CurrentDate     =   38814
   End
   Begin MSComCtl2.DTPicker eos 
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   720
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   308019201
      CurrentDate     =   38814
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ΜΕΣΗ ΤΙΜΗ ΠΩΛΗΣΗΣ"
      Height          =   300
      Left            =   13440
      TabIndex        =   23
      Top             =   1440
      Width           =   3465
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ετοιμες Εκτυπώσεις"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   8760
      TabIndex        =   31
      Top             =   120
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Εκτυπώσεις - Υπολογισμοί"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   13440
      TabIndex        =   32
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton cmdΙΣΟΖΥΓΙΟΜΕ 
         Caption         =   "ΙΣΟΖΥΓΙΟ ΜΕ ΔΙΑΧΩΡΙΣΜΕΝΕΣ ΧΡΟΝΙΕΣ"
         Height          =   360
         Left            =   0
         TabIndex        =   44
         Top             =   3120
         Width           =   3495
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Φίλτρα Εκτύπωσης"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   4080
      TabIndex        =   33
      Top             =   120
      Width           =   4575
      Begin VB.CheckBox meKinisi 
         Caption         =   "Ενεργά μόνο"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Εως"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   38
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Από"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   37
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Αποθήκη (99=όλες μαζί)"
         Height          =   255
         Left            =   480
         TabIndex        =   36
         Top             =   2160
         Width           =   3375
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ταξινόμηση "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   4032
      Begin VB.CheckBox synola 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Σύνολα σε αριθμ.πεδία"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   2760
         Width           =   1785
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Στήλες εκτύπωσης"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   150
         TabIndex        =   39
         Top             =   520
         Width           =   2172
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Τίτλος"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   3120
         Width           =   855
      End
   End
   Begin VB.Shape Shape2 
      Height          =   3015
      Left            =   8880
      Top             =   360
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      Height          =   1305
      Left            =   4320
      Top             =   9015
      Visible         =   0   'False
      Width           =   5730
   End
End
Attribute VB_Name = "apot3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim f_arxiko As String
Dim f_PALIAXRONIA

Dim fupd_eid ' upd_eid =1 'σημαινει υπολογιζει τα υπολ. των ειδών και τα αποθηκευει

'Dim F_PROTHFORA As Integer
Dim FGDB   As New ADODB.Connection

Dim F_RUN  As Integer

Dim Fsql   As String, FIsHidden As Boolean

Dim F_COL  As TrueOleDBGrid80.Column

Dim F_epan As Integer

Dim F_ALIG(70), F_mWidth(50), F_sum(70)
Attribute F_ALIG.VB_VarUserMemId = 1073938437

Dim f_list1(70, 4)
Attribute f_list1.VB_VarUserMemId = 1073938440

Dim F_COLS          As TrueOleDBGrid80.columns

Dim F_UPDYPOL       As Integer

Dim F_TXT(20)  ' ΚΕΙΜΕΝΟ ΑΠΟ ΕΤΟΙΜΑ ΡΕΠΟΡΤ


Dim F_PLATOS_FORMAS As Long

Dim SEL(35)    ' TA PEDIA POY EPILEGV
Attribute SEL.VB_VarUserMemId = 1073938442

Dim fList1ARX As Integer  ' αρχικη επιλογή του λιστ1

Dim F_LHM1    As String

Dim F_LHM2    As String

Dim f_lHM3    As String

Dim f_lab1    As String

Dim f_lab2    As String

Dim f_lab3    As String

Dim f_lab4    As String

Dim f_lab5    As String

Dim f_lab6    As String

Dim f_num1    As String

Dim f_num2    As String

Dim f_num3    As String

Dim sumes(40) As Single




Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)
Attribute fh.VB_VarUserMemId = 1073938456

Private Sub aritmos_pelaton_Click()

        '<EhHeader>
        On Error GoTo aritmos_pelaton_Click_Err

        '</EhHeader>

        Dim X       As String

        Dim pol     As String

        Dim polepis As String

        Dim k       As Long

        Dim R       As New ADODB.Recordset

100     R.Open UCase("select pol,eidos,ajia_apou,SHM1 from parastat"), Gdb, adOpenDynamic, adLockOptimistic

110     pol = ""
120     R.MoveFirst

130     Do While Not R.EOF

140         If Not IsNull(R("eidos") + R("pol") + R("SHM1")) Then
150             If R("pol") = "1" Then     ' παροχη And Left(r("SHM1"), 1) = "2" Then
160                 pol = pol + "'" + R("eidos") + "',"
                End If
            End If

170         R.MoveNext
        Loop

180     pol = Left(pol, Len(pol) - 1)

190     X = " select PARASTAT.TITLOS,count(*) AS [ΔΙΑΦ.ΕΙΔΗ],count(DISTINCT ATIM) AS [ΠΕΛΑΤΕΣ] "
200     X = X + " FROM EGGTIM INNER JOIN PARASTAT ON LEFT(ATIM,1)=PARASTAT.EIDOS "
210     X = X + " WHERE LEFT(ATIM,1) IN (" + pol + ") and "
220     X = X + " HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'"
230     X = X + " GROUP BY PARASTAT.TITLOS"

240     Data1.ConnectionString = gConnect
250     Data1.RecordSource = X
260     Data1.Refresh

270     If Check1 Then

            'typos2
            On Error GoTo aritmos_pelaton_Click_Err

280         print3_xar UCase(X), "0111111111", "", 0
        End If

        '<EhFooter>
        Exit Sub

aritmos_pelaton_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.aritmos_pelaton_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.aritmos_pelaton_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Check3_Click()

        '<EhHeader>
        On Error GoTo Check3_Click_Err

        '</EhHeader>
100     If Check3.Value = vbChecked Then
110         If Len(Trim(Text1.Text)) > 0 Then

120             Text1.Text = Text1.Text + " AND  EID.POS <= EID.SPA AND EID.EPIUYP>0 "
            Else
130             Text1.Text = " EID.POS <= EID.SPA AND EID.EPIUYP>0 "
            End If

        End If

        '<EhFooter>
        Exit Sub

Check3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Check3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Check3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdΙΣΟΖΥΓΙΟΜΕ_Click()
     ISOZ2026
     


End Sub

Private Sub ISOZ2026()


Dim ALLTIME As String
ALLTIME = " and HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND HME<='" + Format(gLhjh, "MM/DD/YYYY") + "'"

Dim PROHG As String
PROHG = " and HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND HME<'" + Format(apo.Value, "MM/DD/YYYY") + "'"

Dim Period As String


 On Error GoTo EXO


Dim nc As Long, Nid As Long, NNID As Long, KK As Integer
nc = GGET_NVALUE("SELECT COUNT(*) FROM EGGTIM WHERE LEFT(ATIM,1)='λ'")
If nc > 0 Then
   Nid = GGET_NVALUE("SELECT COUNT(*) FROM EGGTIM WHERE LEFT(ATIM,1)='λ' and isnull(ID_NUM,0)>0")
   If Nid < nc Then ' ΔΕΝ ΕΧΕΙ Η ΑΠΟΓΡΑΦΗ ΑΝΤΙΣΤΟΙΧΟ RECORD STO TIM
      On Error Resume Next
       Gdb.Execute ("INSERT INTO TIM (KLEIDI,ATIM,HME ) SELECT top 1 ATIM AS KL,ATIM,HME FROM EGGTIM WHERE LEFT(ATIM,1)='λ' and isnull(ID_NUM,0)=0"), KK
         On Error GoTo EXO
        'If KK > 0 Then
             NNID = GGET_NVALUE("SELECT MAX(ID_NUM)  FROM TIM WHERE LEFT(ATIM,1)='λ'")
             Gdb.Execute "UPDATE EGGTIM SET ID_NUM=" + str(NNID) + " WHERE LEFT(ATIM,1)='λ' and  isnull(ID_NUM,0)=0"
        'End If
        
        
   
   End If
End If




Dim R As New ADODB.Recordset
Dim PosAg, PosAgEp, PosPol, PosPolEp
Dim AjAg, AjAgEp, AjPol, AjPolEp

PosAg = " ": PosAgEp = " ": PosPol = " ": PosPolEp = " "
 AjAg = " ": AjAgEp = " ": AjPol = " ": AjPolEp = " "


R.Open "select * from PARASTAT WHERE EIDOS<>'λ' ", Gdb, adOpenDynamic, adLockOptimistic
Dim e As String
Do While Not R.EOF
 e = R("eidos")
 Select Case R!POS_APOU
    Case "1"
        PosAg = PosAg + "'" + e + "',"
    Case "2"
       PosAgEp = PosAgEp + "'" + e + "',"
    Case "3"
        PosPol = PosPol + "'" + e + "',"
    Case "4"
       PosPolEp = PosPolEp + "'" + e + "',"
 End Select
 Select Case R!AJIA_APOU
    Case "1"
        AjAg = AjAg + "'" + e + "',"
    Case "2"
       AjAgEp = AjAgEp + "'" + e + "',"
    Case "3"
        AjPol = AjPol + "'" + e + "',"
    Case "4"
       AjPolEp = AjPolEp + "'" + e + "',"
 End Select

 R.MoveNext
Loop


PosAg = Left(PosAg, Len(PosAg) - 1)
 PosAgEp = Left(PosAgEp, Len(PosAgEp) - 1)
 
 PosPolEp = Left(PosPolEp, Len(PosPolEp) - 1)
 PosPol = Left(PosPol, Len(PosPol) - 1)
 
 AjAg = Left(AjAg, Len(AjAg) - 1)
 AjAgEp = Left(AjAgEp, Len(AjAgEp) - 1)
 
 AjPolEp = Left(AjPolEp, Len(AjPolEp) - 1)
 AjPol = Left(AjPol, Len(AjPol) - 1)


'

R.Close

Dim Q As String
Dim anal As Integer
anal = gVal(InputBox("0=KANONIKA 1.με αναλυση ανα τιμολογιο;2.kATA MHNA 3.KATA HMEPA", , ""))
If anal = 1 Then
    Q = "select  G.ATIM,KODE,D.ONO ,"
ElseIf anal = 2 Then
    Q = "select  MONTH(G.HME),KODE,D.ONO ,"
ElseIf anal = 3 Then
    Q = "select  DAY(G.HME),KODE,D.ONO ,"
Else
    Q = "select  KODE,D.ONO ,"
End If
Q = Q + "sum(IIF(LEFT(G.ATIM,1)='λ' ,POSO,0) )                      AS [ΠΟΣ.ΑΠΟΓ] ,"
Q = Q + "SUM(IIF(LEFT(G.ATIM,1)='λ' ,(XRE-PIS)*TIMM,0) )          AS [ΑΞΙΑ ΑΠΟΓ] ,"

'ΑΓΟΡΕΣ ΠΟΣΟΤΗΤΑ & ΑΞΙΑ
Q = Q + "sum(IIF(LEFT(G.ATIM,1) in (" + PosAg + "," + PosAgEp + ") , IIF(LEFT(G.ATIM,1) IN (" + PosAgEp + ")  ,-POSO,POSO),0) )            AS [ΠΟΣ.ΑΓΟΡ] ,"
Q = Q + "SUM(IIF (charindex(  LEFT(G.ATIM,1) ,'" + Replace(AjAg + AjAgEp, "'", "") + "')>0  ,IIF(LEFT(G.ATIM,1) in (" + AjAgEp + ")" + ",-1,1)*(POSO)*TIMM*(100-G.EKPT)*(100-G.EKPT2)/10000,0 ) ) AS [ΑΞΙΑ AΓΟΡ], "

''ΠΩΛΗΣΕΙΣ ΠΟΣΟΤΗΤΑ & ΑΞΙΑ
Q = Q + "SUM(IIF (charindex(  LEFT(G.ATIM,1) ,'" + Replace(PosPol + PosPolEp, "'", "") + "')>0  ,IIF(LEFT(G.ATIM,1) IN (" + PosPolEp + ")  ,-POSO,POSO) ,0 ) ) AS [ΠΩΛΗΘ.ΠΟΣ],"
Q = Q + "SUM(IIF (charindex(  LEFT(G.ATIM,1) ,'" + Replace(AjPol + AjPolEp, "'", "") + "')>0  ,IIF(LEFT(G.ATIM,1) IN (" + AjPolEp + ")  ,-POSO,POSO)*TIMM*(100-G.EKPT)*(100-G.EKPT2)/10000,0 ) ) AS [ΑΞ.ΠΩΛ],"
'ΥΠΟΛΟΙΠΟ
Q = Q + "sum(IIF(LEFT(G.ATIM,1)='λ' ,XRE-PIS,0) ) + sum(IIF(LEFT(G.ATIM,1) in (" + PosAg + ") , XRE-PIS,0) )"
Q = Q + "-SUM(IIF (charindex(  LEFT(G.ATIM,1) ,'" + Replace(PosPol, "'", "") + "')>0  ,IIF(LEFT(G.ATIM,1) IN (" + PosPolEp + ")  ,-POSO,POSO) ,0 ) )  AS [ΥΠΟΛ.ΠΟΣ]"
'


Q = Q + ",   sum (IIF(LEFT(G.ATIM,1)='λ' ,POSO*D.XTI,0) ) + sum(IIF(LEFT(G.ATIM,1) in (" + PosAg + ") , (XRE-PIS)*D.XTI,0)   )  "
Q = Q + "- SUM(  "
Q = Q + "       IIF (charindex(  LEFT(G.ATIM,1) ,'" + Replace(PosPol, "'", "") + "')>0  ,"
Q = Q + "              IIF(LEFT(G.ATIM,1) IN (" + PosPolEp + ")  ,-POSO*D.XTI,POSO*D.XTI)"
 Q = Q + "      ,0 ) "
 Q = Q + ") AS [ΑΞΙΑ ΥΠΟΛ]   "


Q = Q + " FROM EGGTIM  G  inner  join TIM T ON G.ID_NUM=T.ID_NUM "
Q = Q + "INNER JOIN EID D ON D.KOD=G.KODE "


'Dim mALL As String
'mALL = PosAg + PosAgEp + PosPol + PosPolEp + AjAg + AjAgEp + AjPol + AjPolEp

Q = Q + "WHERE T.AKYROMENO=0 AND charindex(  LEFT(G.ATIM,1) ,'λ," + Replace(PosAg + PosAgEp + PosPol + PosPolEp + AjAg + AjAgEp + AjPol + AjPolEp, "'", "") + "')>0"

      Period = " and G.HME>='" + Format(apo.Value, "MM/DD/YYYY") + "' AND G.HME<='" + Format(eos.Value, "MM/DD/YYYY") + "'"
Q = Q + Period '                 "  AND G.HME>=@x1 and G.HME<@x2"

If anal = 1 Then
   Q = Q + "GROUP BY   G.ATIM,KODE,D.ONO"
ElseIf anal = 2 Then
   Q = Q + "GROUP BY   MONTH(G.HME),KODE,D.ONO"
ElseIf anal = 3 Then
   Q = Q + "GROUP BY   DAY(G.HME),KODE,D.ONO"
    'MONTH(G.HME)
Else

   Q = Q + "GROUP BY   KODE,D.ONO"
End If




530     Data1.ConnectionString = gConnect
540     Data1.RecordSource = Q
550  On Error GoTo EXO
Data1.Refresh

Dim sumes(39)
Dim k As Integer


         Do While Not Data1.Recordset.EOF
              
              For k = 0 To Data1.Recordset.FIELDS.Count - 1
                 If IsNumeric(Data1.Recordset.FIELDS(k).Value) And Data1.Recordset.FIELDS(k).Type <> 202 Then
                     sumes(k) = sumes(k) + Data1.Recordset.FIELDS(k).Value
                 End If
              Next
              Data1.Recordset.MoveNext
         Loop

Dim sum_str
sum_str = "0011111111111111111"
         For k = 1 To Data1.Recordset.FIELDS.Count - 1

             If mID(sum_str, k, 1) = "1" Then
                 GR2.Splits(0).columns(k).FooterText = Format(sumes(k), "######0.00")
             End If

         Next

     GR2.AlternatingRowStyle = True
     GR2.OddRowStyle.BackColor = &H8000000F   ' GRI   vbCyan
     GR2.EvenRowStyle.BackColor = &HFFFFC0   'OYRANI     &H8000000F  ' GRI











560     If Check1 Then

            'typos2
          '  On Error GoTo Command8_Click_Err

570         print3_xar UCase(Q), sum_str, "", 0
        End If


Exit Sub
EXO:
MsgBox Err.Description

End Sub





Private Sub GR2_AfterColEdit(ByVal ColIndex As Integer)


If ColIndex = 0 Then
  If Len(f_arxiko) > 0 Then
      GR2.columns(0).Value = f_arxiko
      f_arxiko = ""
  End If
  End If
  
End Sub

Private Sub GR2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)

     If ColIndex = 0 Then
         f_arxiko = GR2.columns(0)
          Exit Sub
         '   Else
          '      GR2.AllowUpdate = True
     End If


End Sub



'

Private Sub YPOLOGISMOS_Click()

' upd_eid =1 'σημαινει υπολογιζει τα υπολ. των ειδών και τα αποθηκευει
' 0  =  σε προσωρινό πινακα
'Dim fupd_eid






        'ΥΠΟΛΟΓΙΖΕΙ ΤΟ ΡΕΠΟΡΤ ΕΙΤΕ
        ' ΑΠΟ ΑΥΤΑ ΠΟΥ ΕΠΕΛΕΞΑ ΕΙΤΕ
        ' ΑΠΟ ΤΟ ΕΤΟΙΜΟ ΡΕΠΟΡΤ ΠΟΥ ΕΧΩ ΕΠΙΛΕΞΕΙ
        '<EhHeader>
        On Error GoTo YPOLOGISMOS_Click_Err

        '</EhHeader>

        Dim X  As String, order As String

        Dim k  As Integer, R As New ADODB.Recordset    ', sum As Single

        Dim mc As Integer

        Dim EPIK(38), alig(38), mWidth(38)

        Dim Sum(30) As Single

        Dim TIM     As Long

        Dim SUgm_str  'που θα κανει sumes

        Dim Kinhseis As Boolean

        Dim a

100     Kinhseis = False

        ' για να βγαζει combo στο grid
        Dim gm_fpa As Integer

        Dim m_kat  As Integer

110     m_kat = 0    ' krato thn stili tis katigorias
120     gm_fpa = 0    ' krato thn stili toy FPA

130     Me.MousePointer = vbHourglass
140     TIM = GetCurrentTime()

150     For k = 0 To 9: Sum(k) = 0

        Next
       For k = 1 To 18: mWidth(k) = 1: alig(k) = 0: EPIK(k) = "":
       Next
170     F_RUN = 1


           


180     Data1.ConnectionString = gConnect



  '    Gdb.BeginTrans




            Dim mnow As String
            
            If fupd_eid = 1 Then
                mnow = "EID"
            Else
            
               mnow = "DOKEID" + Format(Now, "YYMMDDHHmm")
               Gdb.Execute "DROP TABLE " + mnow
               Gdb.Execute "SELECT * into " + mnow + "  FROM EID"
            End If
            
           



190     If F_UPDYPOL = 1 Then
200         UPDATE_YPOLOIPA3 mnow, List11, apo, eos
        End If
    
210     Me.Caption = GetCurrentTime() - TIM

220     mc = 2

230     If Check2 Then
240         SUgm_str = "000"
250         X = "select  '@@@@@' AS [A/A],LEFT(KOD,20) AS [Κωδικός],LEFT(ONO,35) as [Περιγραφή]"    '
        Else
260         SUgm_str = "00"
270         X = "select KOD as [Κωδικός],ONO as [Περιγραφή]"    '
        End If

        Dim KK

        Dim temp

280     For KK = 0 To List1.ListCount - 1

290         If List1.Selected(KK) Then

                'ο πραγματικός δείκτης (γιατί μπορεί να μετακινήθηκε)
300             temp = Val(Right(List1.List(KK), 2))
310             X = X + f_list1(temp, 0)
320             mc = mc + 1
                'x = x + f_list1(Val(Right(List1.List(KK), 2)), 0)
330             F_sum(mc) = f_list1(temp, 1)
340             F_ALIG(mc) = f_list1(temp, 2)
350             F_mWidth(mc) = f_list1(temp, 3)

360             If f_list1(temp, 1) = 1 Then
370                 SUgm_str = SUgm_str + "1"
                Else
380                 SUgm_str = SUgm_str + "0"
                End If
            End If

        Next
X = X + " ,'' AS [`] " ' ΓΙΑ ΝΑ ΜΗΝ ΚΟΛΛΑΕΙ ΔΕΞΙΑ ΤΟ ΤΕΛΕΥΤΑΙΟ ΠΕΔΙΟ
        '

        'For K = 0 To List1.ListCount - 1
        '   If List1.Selected(K) Then
        '      x = x + f_list1(K, 0)
        '      mc = mc + 1
        '      F_sum(mc) = f_list1(K, 1)
        '      F_ALIG(mc) = f_list1(K, 2)
        '      F_mWidth(mc) = f_list1(K, 3)
        '   End If
        'Next

390     Kinhseis = True

        ' Kinhseis = False 'debug

        '-------------------------------- NEO ARXH ===============================
400     If Kinhseis Then
    
       On Error Resume Next

410         Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1


On Error Resume Next





420         Gdb.Execute "UPDATE EGGTIM SET EKPT=0 WHERE EKPT IS NULL"
        
430         Gdb.Execute "UPDATE EGGTIM SET XRE=0 WHERE XRE IS NULL"
        
440         Gdb.Execute "UPDATE EGGTIM SET PIS=0 WHERE PIS IS NULL"
        
450         Gdb.Execute "UPDATE EGGTIM SET TIMM=0 WHERE TIMM IS NULL"
        
460         Gdb.Execute "UPDATE EGGTIM SET APOT=1 WHERE APOT IS NULL"

            '----------------------------------------------------------------------------------------
            'ΠΑΙΡΝΕΙ ΜΟΝΟ ΑΠΟ ΤΟ ΤΟ ΛΙΣΤΒΟΧ ΤΙΣ ΚΙΝΗΣΕΙΣ
            Dim pol

            Dim KKE

470         pol = ""

480         For KKE = 0 To List11.ListCount - 1

490             If List11.Selected(KKE) = True Then
500                 pol = pol + LTrim(str(Asc(Right(List11.List(KKE), 1)))) + ","
                End If

            Next

510         If Len(pol) - 1 > 0 Then
520             pol = Left(pol, Len(pol) - 1)
            Else
530             pol = "' '"
            End If

            '----------------------------------------------------------------------------------------

540         a = "SELECT KODE," & " SUM( CASE APOT  WHEN 1  THEN XRE ELSE 0  END ) AS S1X,"
                   a = a + " SUM(CASE APOT  WHEN 1  THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS1X,"
                   a = a + " SUM(CASE APOT  WHEN 1  THEN PIS ELSE 0  END ) AS S1P,"
                   a = a + " SUM(CASE APOT  WHEN 1  THEN PIS*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS1P,"
                   a = a + " SUM( CASE WHEN APOT=1 AND LEFT(ATIM,1)='λ'  THEN XRE ELSE 0  END )-SUM(CASE   WHEN APOT=1 AND LEFT(ATIM,1)='λ'  THEN PIS ELSE 0  END ) AS APOG01,"
                   a = a + " SUM( CASE WHEN APOT=2 AND LEFT(ATIM,1)='λ'  THEN XRE ELSE 0  END )-SUM(CASE   WHEN APOT=2 AND LEFT(ATIM,1)='λ'  THEN PIS ELSE 0  END ) AS APOG02,"
                   a = a + " SUM( CASE WHEN APOT=3 AND LEFT(ATIM,1)='λ'  THEN XRE ELSE 0  END )-SUM(CASE   WHEN APOT=3 AND LEFT(ATIM,1)='λ'  THEN PIS ELSE 0  END ) AS APOG03,"
                   a = a + " SUM( CASE WHEN APOT=4 AND LEFT(ATIM,1)='λ'  THEN XRE ELSE 0  END )-SUM(CASE   WHEN APOT=4 AND LEFT(ATIM,1)='λ'  THEN PIS ELSE 0  END ) AS APOG04,"

                   a = a + " SUM( CASE WHEN APOT=1 AND LEFT(ATIM,1)='λ'  THEN XRE*TIMM ELSE 0  END )-SUM(CASE   WHEN APOT=1 AND LEFT(ATIM,1)='λ'  THEN PIS*TIMM ELSE 0  END ) AS KOSAPOG01,"
                   a = a + " SUM( CASE WHEN APOT=2 AND LEFT(ATIM,1)='λ'  THEN XRE*TIMM ELSE 0  END )-SUM(CASE   WHEN APOT=2 AND LEFT(ATIM,1)='λ'  THEN PIS*TIMM ELSE 0  END ) AS KOSAPOG02,"
                   a = a + " SUM( CASE WHEN APOT=3 AND LEFT(ATIM,1)='λ'  THEN XRE*TIMM ELSE 0  END )-SUM(CASE   WHEN APOT=3 AND LEFT(ATIM,1)='λ'  THEN PIS*TIMM ELSE 0  END ) AS KOSAPOG03,"
                   a = a + " SUM( CASE WHEN APOT=4 AND LEFT(ATIM,1)='λ'  THEN XRE*TIMM ELSE 0  END )-SUM(CASE   WHEN APOT=4 AND LEFT(ATIM,1)='λ'  THEN PIS*TIMM ELSE 0  END ) AS KOSAPOG04,"









            a = a + " SUM( CASE APOT  WHEN 1  THEN XRE ELSE 0  END )-SUM(CASE APOT  WHEN 1  THEN PIS ELSE 0  END ) AS XPOS01," _
               & " SUM( CASE APOT  WHEN 2 THEN XRE ELSE 0  END ) AS S2X," _
               & " SUM( CASE APOT  WHEN 2  THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS2X," _
               & " SUM(CASE APOT  WHEN 2  THEN PIS*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS2P," _
               & " SUM( CASE APOT WHEN 2 THEN PIS ELSE 0 END ) AS S2P," _
               & " SUM( CASE APOT  WHEN 2  THEN XRE ELSE 0  END )-SUM(CASE APOT  WHEN 2  THEN PIS ELSE 0  END ) AS XPOS02," _
               & " SUM( CASE APOT  WHEN 3 THEN XRE ELSE 0  END ) AS S3X," _
               & " SUM( CASE APOT  WHEN 3  THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS3X," _
               & " SUM(CASE APOT  WHEN 3  THEN PIS*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS3P," _
               & " SUM( CASE APOT WHEN 3 THEN PIS ELSE 0 END ) AS S3P," _
               & " SUM( CASE APOT  WHEN 3  THEN XRE ELSE 0  END )-SUM(CASE APOT  WHEN 3  THEN PIS ELSE 0  END ) AS XPOS03," _
               & " SUM( CASE APOT  WHEN 4 THEN XRE ELSE 0  END ) AS S4X," _
               & " SUM( CASE APOT  WHEN 4  THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS4X," _
               & " SUM(CASE APOT  WHEN 4  THEN PIS*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS4P," _
               & " SUM( CASE APOT WHEN 4 THEN PIS ELSE 0 END ) AS S4P, " _
               & " SUM( CASE APOT  WHEN 4  THEN XRE ELSE 0  END )-SUM(CASE APOT  WHEN 4  THEN PIS ELSE 0  END ) AS XPOS04," _
               & " SUM(CASE  WHEN APOT=1   AND ascii(left(ATIM,1)) in (" + pol + ") THEN PIS ELSE 0  END ) AS ES1P," _
               & " SUM(CASE  WHEN APOT=2   AND ascii(left(ATIM,1)) in (" + pol + ") THEN PIS ELSE 0  END ) AS ES2P," _
               & " SUM(CASE  WHEN APOT=3   AND ascii(left(ATIM,1)) in (" + pol + ") THEN PIS ELSE 0  END ) AS ES3P," _
               & " SUM(CASE  WHEN APOT=4   AND ascii(left(ATIM,1)) in (" + pol + ") THEN PIS ELSE 0  END ) AS ES4P  INTO DOKEGGT1 "
          
          
          
          
          
550         a = a + " FROM EGGTIM WHERE  (NOT APOT IS NULL) AND (NOT EKPT IS NULL) AND (NOT XRE IS NULL) AND (NOT PIS IS NULL) AND (NOT TIMM IS NULL) AND "
            a = a + " ( LEFT(ATIM,1)='λ' or ( HME>='" + Format(apo, "MM/DD/YYYY") + "' AND   HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' ) )  GROUP BY KODE"

            '   On Error GoTo Command1_Click_Err

            'Gdb.Execute "UPDATE EGGTIM SET PIS=0 WHERE PIS IS NULL"
560         Gdb.Execute "UPDATE EGGTIM SET EKPT=0 WHERE PIS IS NULL" ' afoimagrioti@gmail.com

570         Gdb.Execute a, k

580         X = X + " FROM " + mnow + " E  left JOIN DOKEGGT1  ON DOKEGGT1.KODE=E.KOD "
        Else
590         X = X + " from " + mnow + " E  left JOIN DOKEGGT1  ON DOKEGGT1.KODE=E.KOD"
        End If
       
           Dim mEnerga As String
           If meKinisi.Value = vbChecked Then
              mEnerga = " ENERGO=1 and "
           Else
              mEnerga = ""
           End If
       
       
600     If Len(Text1.Text) > 0 Then
           If Len(mEnerga) > 1 Then
610           X = X + " WHERE  ENERGO=1 and " + Text1.Text  ' Fsql
           Else
            X = X + " WHERE " + Text1.Text  ' Fsql
           End If

        Else
           If Len(mEnerga) > 1 Then
              X = X + " WHERE  ENERGO=1 "   ' Fsql
           End If
        End If

620     order = "kod"

630     For k = 0 To 2

640         If Combo2.Text = Combo2.List(0) Then
650             order = "E.KOD"
660         ElseIf Combo2.Text = Combo2.List(1) Then
670             order = "ONO"
680         ElseIf Combo2.Text = Combo2.List(2) Then
690             order = "POS"
700         ElseIf Combo2.Text = Combo2.List(3) Then
710             order = "PROM"
720         ElseIf Combo2.Text = Combo2.List(4) Then
730             order = "XTI"
740         ElseIf Combo2.Text = Combo2.List(5) Then
750             order = "XTI*POS"

            End If

        Next

760     X = X + " order by " + order
             
770     If rCheck.Value = vbChecked Then    ' AN PAIRNV TO EPILEGMENO REPORT
780         X = Text4.Text
            If InStr(F_TXT(1), "ΣΥΝΟΛΑ") > 0 And InStr(F_TXT(1), "/*") > 0 Then
                SUgm_str = LTrim(Trim(mID(F_TXT(1), InStr(F_TXT(1), "ΣΥΝΟΛΑ") + 6, 20)))
            End If

        Else
790         Text4.Text = X
        End If

800     Data1.RecordSource = X

        On Error GoTo lathos

810     Data1.Refresh

        On Error Resume Next

        Dim xd As New TrueOleDBGrid80.ValueItem

        ' On Error GoTo Command1_Click_Err

820     If gm_fpa > 0 Then
830         xd.Value = 1
840         xd.DisplayValue = "9%"
850         GR2.columns(gm_fpa).ValueItems.ADD xd

860         xd.Value = 2
870         xd.DisplayValue = "19%"
880         GR2.columns(gm_fpa).ValueItems.ADD xd

890         GR2.columns(gm_fpa).ValueItems.translate = True
900         GR2.columns(gm_fpa).ValueItems.Presentation = dbgComboBox
        End If

910     If m_kat > 0 Then

            Dim r7 As New ADODB.Recordset

920         r7.Open "SELECT *FROM PINAKES WHERE TYPOS=3 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

930         Do While Not r7.EOF

940             If r7("typos") = 3 Then
950                 xd.Value = Format(r7("ayjon"), "000")
960                 xd.DisplayValue = r7("PERIGRAFH")
970                 GR2.columns(m_kat).ValueItems.ADD xd
                Else
                

                    Exit Do

                End If

980             r7.MoveNext
            Loop

            ' kathgories

990         R.Close

            'xd.Value = 1
            'xd.DisplayValue = "9%"
            'GR2.Columns(gm_fpa).ValueItems.Add xd

1000        GR2.columns(m_kat).ValueItems.translate = True
1010        GR2.columns(m_kat).ValueItems.Presentation = dbgComboBox
        End If

        'GR2.Columns(0).Text = "Κωδικός"
        'GR2.Columns(1).Text = "Περιγραφή"
        'GR2.Columns(2).Text = "Ποσότητα"

1020    For k = 5 To 12

1030        If Len(EPIK(k)) > 1 Then
                '      GR2.Columns(K).Text = epik(K)
1040            GR2.columns(k).alignment = alig(k)
1050            GR2.columns(k).Width = mWidth(k)

            End If

        Next
        Dim nT As Long
        For k = 2 To Data1.Recordset.FIELDS.Count - 1
          nT = Data1.Recordset(k).Type
          If nT = 7 Or nT = 5 Or nT = 3 Or nT = 131 Then
             GR2.columns(k).NumberFormat = "######0.00"
          End If
        Next

1060    GR2.columns(0).Width = 2000
1070    GR2.columns(1).Width = 6000
1080    GR2.columns(2).Width = 1000

        GR2.Font.Bold = True
       'KANEI TON ARIUMO ME 2 ΔΕΚΑΔΙΚΑ Π.Χ. "11,30"  GR2.columns(3).NumberFormat = "######0.00"

1090    Me.MousePointer = vbNormal

1100    If Check1 Then

            'typos2
            On Error GoTo YPOLOGISMOS_Click_Err

1110        If Len(SUgm_str) < 3 Then SUgm_str = "0111111111111111"
1120        print3_xar UCase(X), SUgm_str, Text2.Text, 0   ' RR.RecordSource
        End If

        'If Check2 Then
        ' Print4_Excel UCase(x), SUgm_str
        'End If
        
        'On Error GoTo 0

Dim SUMES2(0 To 50)

If synola.Value = vbChecked Then
         Do While Not Data1.Recordset.EOF

             For k = 0 To Data1.Recordset.FIELDS.Count - 1

                 If IsNumeric(Data1.Recordset.FIELDS(k).Value) And Data1.Recordset.FIELDS(k).Type <> 202 Then
                     SUMES2(k) = SUMES2(k) + Data1.Recordset.FIELDS(k).Value
                    End If

                Next
               ' DoEvents

                Data1.Recordset.MoveNext

                '  Exit Do
            Loop

         For k = 0 To Data1.Recordset.FIELDS.Count - 1

             If SUMES2(k) > 0 Then
                 GR2.Splits(0).columns(k).FooterText = Format(SUMES2(k), "######0.00")
             End If

         Next
End If

















On Error Resume Next

'ΣΒΗΝΕΙ ΤΟΝ ΠΡΟΣΩΡΙΝΟ ΠΙΝΑΚΑ
If Left(mnow, 3) = "DOK" Then
   Gdb.Execute "drop table " + mnow
End If

   ' Gdb.RollbackTrans
   
    



        Exit Sub

lathos:
1130    MsgBox "λάθος " + Err.Description

1140    Resume Next

        '<EhFooter>
        Exit Sub

YPOLOGISMOS_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'---------------------------------------------
'Private Sub Typose()
'    Dim R, sum, N
'    Set R = data1.Recordset
'
'
'
'    Printer.FONTSIZE = 8
'    Printer.FontName = "Courier New"
'    Printer.Font.Charset = 161
'    Printer.FontBold = True
'
'
'    Printer.Print Tab(4); "Κωδικός"; Tab(10); "Επωνυμία"; Tab(36); "Υπόλοιπο"; Tab(48); "Τηλέφωνο"
'    R.MoveFirst
'    sum = 0
'    N = 0
'    Do While Not R.EOF
'        N = N + 1
'        Printer.Print Format(N, "###"); Tab(4); R("KOD"); Tab(10); Left(R("EPO"), 25); Tab(36); Format(R("TYP"), "###,###.00"); Tab(48); R("thl")
'        sum = sum + R("typ")
'        R.MoveNext
'    Loop
'    Printer.Print
'    Printer.Print Tab(36); Format(sum, "###,###.00")
'End Sub
'Sub typos2()
'    Dim R, sum, N, db As Database, db2 As Database
'
'    Dim X As String, order As String
'    Dim k As Integer
'
'
'
'
'
'    F_RUN = 1
'
'    'data1.DatabaseName = gDir
'    DATA1.ConnectionString = gConnect
'
'    Set db2 = OpenDatabase("c:\mercvb\reports\reports.mdb")    ', False, False, gConnect)
'
'    On Error Resume Next
'    db2.Execute "drop table dok2 "
'    db2.Close
'
'
'
'    X = "select kod,ono,pos,erg,PROM into dokeid in 'c:\mercvb\reports\reports.mdb' from eid  "
'
'    'If IsNumeric(apo.Text) And IsNumeric(eos) Then
'    '   x = x + " WHERE pos>=" + apo.Text + " and pos<=" + eos.Text
'    'End If
'
'    order = "kod"
'    For k = 0 To 2
'        If Combo2.Text = Combo2.List(0) Then
'            order = "kod"
'        ElseIf Combo2.Text = Combo2.List(1) Then
'            order = "ono"
'        ElseIf Combo2.Text = Combo2.List(2) Then
'            order = "pos"
'        ElseIf Combo2.Text = Combo2.List(3) Then
'            order = "prom"
'        End If
'    Next
'    X = X + " order by " + order
'    db.Execute X
'
'    'CrystalReport1.Action = 1
'
'
'End Sub
Private Sub Command10_Click()


        '======== ??S? ???? a????? S? ???????????? ???S???? ===========================================================================
        '<EhHeader>
        On Error GoTo Command10_Click_Err

        '</EhHeader>
        Dim arx_arit, tel_arit, buff, kod_tim, ono_par, p1, progr, k

        Dim m_rec, m_synt, synt1, synt2

        Dim SUMAEID As New ADODB.Recordset

100     arx_arit = 0: tel_arit = 0

        Dim polhs As String

        Dim PolXParoxh As String

        Dim Paroxh As String

        Dim AGOEPIS As String

        Dim polepis As String, pol As String, ago As String, APIS

        Dim a

110     a = Get2_AJ(polhs, polepis, ago, AGOEPIS, PolXParoxh, Paroxh)

        Dim EGGTIM As New ADODB.Recordset

        Dim X

        Dim sum_str

        Dim SUgm_str
Dim synt As String

120     X = "SELECT KODE,ONOMA,SUM(POSO) as [ΠΟΣ],SUM(POSO*TIMM*(1-EKPT/100)) AS  [ΑΞΙΑ],(CASE WHEN SUM(POSO)=0 THEN 0 ELSE SUM(POSO*TIMM*(1-EKPT/100))/SUM(POSO) END ) AS [MES.ΤΙΜΗ ΑΓΟΡΑΣ] FROM EGGTIM WHERE LEFT(ATIM,1) IN ('λ'," + ago + ") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' group by KODE,ONOMA"

130     EGGTIM.Open X, Gdb, adOpenDynamic, adLockOptimistic
        Gdb.Execute "drop TABLE DOKMESXTI"
        
         synt = " LEFT(ATIM,1) IN ('λ'," + ago + AGOEPIS + ") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' "
        
         Gdb.Execute "UPDATE EGGTIM SET MIKTA=0 WHERE  " + synt
        
         '?G???S ?O??S ??S?O????
         Gdb.Execute "UPDATE EGGTIM SET KOLA=POSO, MIKTA=POSO*TIMM*(1-EKPT/100) WHERE  " + " LEFT(ATIM,1) IN ('λ'," + ago + ") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' "
         
         '??S?O???? ?G??O?  STA PISTOTIKA BAZO MONO AXIES OXI POSOTHTES
         Gdb.Execute "UPDATE EGGTIM SET MIKTA=-POSO*TIMM*(1-EKPT/100) WHERE  " + " LEFT(ATIM,1) IN (" + AGOEPIS + ") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' "
         
         
         
        
'140     Gdb.Execute "SELECT KODE,(CASE WHEN SUM(POSO)=0 THEN 0 ELSE SUM(POSO*TIMM*(1-EKPT/100))/SUM(POSO) END ) AS MESXTI  INTO DOKMESXTI FROM EGGTIM WHERE LEFT(ATIM,1) IN ('?'," + ago + ") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' group by KODE"
'150     Gdb.Execute "UPDATE EID SET EID.MESXTI=DOKMESXTI.MESXTI FROM DOKMESXTI LEFT JOIN EID ON EID.KOD=DOKMESXTI.KODE  "

'140     Gdb.Execute "SELECT KODE,(CASE WHEN SUM(KOLA)=0 THEN 0 ELSE SUM(MIKTA)/SUM(KOLA) END ) AS MESXTI  INTO DOKMESXTI FROM EGGTIM WHERE LEFT(ATIM,1) IN ('?'," + AGOEPIS + ago + ") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' group by KODE"



150     Dim sql As String
sql = "UPDATE EID SET MESXTI=(SELECT (CASE WHEN SUM(KOLA)=0 THEN 0 ELSE SUM(MIKTA)/SUM(KOLA) END ) From EGGTIM "
sql = sql + " WHERE KODE=EID.KOD AND  HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "') "


Gdb.Execute sql



160     If Check1 Then

            'typos2
            On Error GoTo Command10_Click_Err

170         SUgm_str = "0111111111111111"
            'Dim ff As String
            'DD = X

180         print3_xar X, SUgm_str, Format(apo, "DD/MM/YYYY") + " ?OS " + Format(eos, "DD/MM/YYYY"), 0        ' RR.RecordSource"
        End If

        '<EhFooter>
        Exit Sub

Command10_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Command10_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Command10_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>





























'        '======== ΜΕΣΗ ΤΙΜΗ αγοράς ΣΕ ΗΜΕΡΟΜΗΝΙΑΚΟ ΔΙΑΣΤΗΜΑ ===========================================================================
'        '<EhHeader>
'        On Error GoTo Command10_Click_Err
'
'        '</EhHeader>
'        Dim arx_arit, tel_arit, BUFF, kod_tim, ono_par, p1, progr, k
'
'        Dim m_rec, m_synt, synt1, synt2
'
'        Dim SUMAEID As New ADODB.Recordset
'
'100     arx_arit = 0: tel_arit = 0
'
'        Dim polhs As String
'
'        Dim PolXParoxh As String
'
'        Dim Paroxh As String
'
'        Dim AGOEPIS As String
'
'        Dim polepis As String, pol As String, ago As String, APIS
'
'        Dim a
'
'110     a = Get2_AJ(polhs, polepis, ago, AGOEPIS, PolXParoxh, Paroxh)
'
'        Dim EGGTIM As New ADODB.Recordset
'
'        Dim x
'
'        Dim SUM_STR
'
'        Dim SUgm_str
'Dim synt As String
'
'120     x = "SELECT KODE,ONOMA,SUM(POSO) as [ΣΥΝ.ΠΟΣΟΤ],SUM(POSO*TIMM*(1-EKPT/100)) AS  [ΣΥΝ.ΑΞΙΑΣ],(CASE WHEN SUM(POSO)=0 THEN 0 ELSE SUM(POSO*TIMM*(1-EKPT/100))/SUM(POSO) END ) AS [MEΣ.ΤΙΜΗ ΑΓΟΡΑΣ] FROM EGGTIM WHERE LEFT(ATIM,1) IN ('λ'," + ago + ") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' group by KODE,ONOMA"
'
'130     EGGTIM.Open x, Gdb, adOpenDynamic, adLockOptimistic
'        Gdb.Execute "drop TABLE DOKMESXTI"
'
'         synt = " LEFT(ATIM,1) IN ('λ'," + ago + AGOEPIS + ") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' "
'
'         Gdb.Execute "UPDATE EGGTIM SET MIKTA=0 WHERE  " + synt
'
'         'ΑΓΟΡΕΣ ΧΩΡΙΣ ΠΙΣΤΩΤΙΚΑ
'         Gdb.Execute "UPDATE EGGTIM SET KOLA=POSO, MIKTA=POSO*TIMM*(1-EKPT/100) WHERE  " + " LEFT(ATIM,1) IN ('λ'," + ago + ") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' "
'
'         'ΠΙΣΤΩΤΙΚΑ ΑΓΟΡΩΝ  STA PISTOTIKA BAZO MONO AXIES OXI POSOTHTES
'         Gdb.Execute "UPDATE EGGTIM SET MIKTA=-POSO*TIMM*(1-EKPT/100) WHERE  " + " LEFT(ATIM,1) IN (" + AGOEPIS + ") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' "
'
'
'
'
''140     Gdb.Execute "SELECT KODE,(CASE WHEN SUM(POSO)=0 THEN 0 ELSE SUM(POSO*TIMM*(1-EKPT/100))/SUM(POSO) END ) AS MESXTI  INTO DOKMESXTI FROM EGGTIM WHERE LEFT(ATIM,1) IN ('λ'," + ago + ") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' group by KODE"
''150     Gdb.Execute "UPDATE EID SET EID.MESXTI=DOKMESXTI.MESXTI FROM DOKMESXTI LEFT JOIN EID ON EID.KOD=DOKMESXTI.KODE  "
'
'140     Gdb.Execute "SELECT KODE,(CASE WHEN SUM(KOLA)=0 THEN 0 ELSE SUM(MIKTA)/SUM(KOLA) END ) AS MESXTI  INTO DOKMESXTI FROM EGGTIM WHERE LEFT(ATIM,1) IN ('λ'," + AGOEPIS + ago + ") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' group by KODE"
'
'150     Gdb.Execute "UPDATE EID SET EID.MESXTI=DOKMESXTI.MESXTI FROM DOKMESXTI LEFT JOIN EID ON EID.KOD=DOKMESXTI.KODE  "
'
'
'
'
'
'160     If Check1 Then
'
'            'typos2
'            On Error GoTo Command10_Click_Err
'
'170         SUgm_str = "0111111111111111"
'            'Dim ff As String
'            'DD = X
'
'180         print3_xar x, SUgm_str, Format(apo, "DD/MM/YYYY") + " ΕΩΣ " + Format(eos, "DD/MM/YYYY"), 0        ' RR.RecordSource"
'        End If
'
'        '<EhFooter>
'        Exit Sub
'
'Command10_Click_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.apot3.Command10_Click " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Command10_Click " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>

End Sub

Private Sub Command11_Click()

        '<EhHeader>
        On Error GoTo Command11_Click_Err

        '</EhHeader>
100     If Len(Dir("c:\mercvb\queries\EID", vbDirectory)) = 0 Then
110         MkDir "c:\mercvb\queries\EID"
        End If

120     cd1.InitDir = "c:\mercvb\queries\EID"
130     cd1.ShowSave

        Dim F

140     F = cd1.FILENAME

150     Open F For Output As #5
160     Print #5, Text4.Text
170     Close #5

        '<EhFooter>
        Exit Sub

Command11_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Command11_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Command11_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command12_Click()

        '<EhHeader>
        On Error GoTo Command12_Click_Err

        '</EhHeader>

100     f_PALIAXRONIA = Trim(FINDPARAMETROI(1, "PAR1", "F_PALIAXRONIA", "2005", "Γ.18 DSN ΠΑΛΙΑΣ ΧΡΟΝΙΑΣ"))

        On Error Resume Next
      
110     If Len(f_PALIAXRONIA) > 0 Then
120         FGDB.Open gConnect + ";DATABASE=" + f_PALIAXRONIA
130         TelProm
140         FGDB.Close
      
        End If

150     FGDB.Open gConnect
160     TelProm

170     FGDB.Close

        '<EhFooter>
        Exit Sub

Command12_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Command12_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Command12_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>

100     Unload Me

110     If gApoMenu = True Then
120         SendKeys "%"
130         SendKeys "{RIGHT}+{DOWN}+{DOWN}+{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>

100     TelHmeAgoras apo.Value, eos.Value


        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command4_Click()

        '======== ΜΕΣΗ ΤΙΜΗ ΠΩΛΗΣΗΣ ΣΕ ΗΜΕΡΟΜΗΝΙΑΚΟ ΔΙΑΣΤΗΜΑ ===========================================================================
        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>
        Dim arx_arit, tel_arit, buff, kod_tim, ono_par, p1, progr, k

        Dim m_rec, m_synt, synt1, synt2

        Dim SUMAEID As New ADODB.Recordset

100     arx_arit = 0: tel_arit = 0

        ' sele TIM
        'SET index to ARTHME

        Dim polhs As String

        Dim PolXParoxh As String

        Dim Paroxh As String

        Dim AGOEPIS As String

        Dim polepis As String, pol As String, ago As String, APIS

        Dim a

110     a = Get2_AJ(polhs, polepis, ago, AGOEPIS, PolXParoxh, Paroxh)

        Dim EGGTIM As New ADODB.Recordset

        Dim X

        Dim sum_str

        Dim SUgm_str

120     X = "SELECT KODE,ONOMA,SUM(POSO) as [ΣΥΝ.ΠΟΣΟΤ],SUM(POSO*TIMM*(1-EKPT/100)) AS  [ΣΥΝ.ΑΞΙΑΣ],SUM(POSO*TIMM*(1-EKPT/100))/SUM(POSO) AS [MEΣΗ ΤΙΜΗ ΠΩΛΗΣΗΣ] FROM EGGTIM WHERE LEFT(ATIM,1) IN ('T','t','L') AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' group by KODE,ONOMA"

130     EGGTIM.Open X, Gdb, adOpenDynamic, adLockOptimistic

140     If Check1 Then

            'typos2
            On Error GoTo Command4_Click_Err

150         SUgm_str = "0111111111111111"
160         print3_xar UCase(X), SUgm_str, Format(apo, "DD/MM/YYYY") + " ΕΩΣ " + Format(eos, "DD/MM/YYYY"), 0        ' RR.RecordSource"
        End If

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    Dim m As Long
    
    If F_PLATOS_FORMAS < 0 Then
        ResizeFormFor Me
    Else
        '       If F_PLATOS_FORMAS < MDIForm1.Width Then   ' Screen.Width Then
        '          m = MDIForm1.Width '  Screen.Width
        '       Else
        m = F_PLATOS_FORMAS
        '       End If
        '       'εαν το κουμπι εξόδου βγαίνει εξω να το μαζευει
        '       If Command2.Left + Command2.Width + 100 > m Then
        '          m = Command2.Left + Command2.Width + 100
        '       End If
        '
 'If F_PROTHFORA = 1 Then
  '   F_PROTHFORA = 0
 'Else
        mForm_Resize Me, 8, m, 9855, fh, fw, ft, fl
' End If
        
       
    End If
    
    '    '</EhHeader>
    '    If F_PLATOS_FORMAS < 0 Then
    '       ResizeFormFor Me
    '    Else
    '       If F_PLATOS_FORMAS < Screen.Width Then
    '          m = Screen.Width
    '       Else
    '          m = F_PLATOS_FORMAS
    '       End If
    '
    '       mForm_Resize Me, 8, m, 9915, fh, fw, ft, fl
    '    End If

End Sub

Sub neo_isoz()

Dim R As New ADODB.Recordset
Dim PosAg, PosAgEp, PosPol, PosPolEp
Dim AjAg, AjAgEp, AjPol, AjPolEp

PosAg = " ": PosAgEp = " ": PosPol = " ": PosPolEp = " "
 AjAg = " ": AjAgEp = " ": AjPol = " ": AjPolEp = " "


R.Open "select * from PARASTAT WHERE EIDOS<>'λ' ", Gdb, adOpenDynamic, adLockOptimistic
Dim e As String
Do While Not R.EOF
 e = R("eidos")
 Select Case R!POS_APOU
    Case "1"
        PosAg = PosAg + "'" + e + "',"
    Case "2"
       PosAgEp = PosAgEp + "'" + e + "',"
    Case "3"
        PosPol = PosPol + "'" + e + "',"
    Case "4"
       PosPolEp = PosPolEp + "'" + e + "',"
 End Select
 Select Case R!AJIA_APOU
    Case "1"
        AjAg = AjAg + "'" + e + "',"
    Case "2"
       AjAgEp = AjAgEp + "'" + e + "',"
    Case "3"
        AjPol = AjPol + "'" + e + "',"
    Case "4"
       AjPolEp = AjPolEp + "'" + e + "',"
 End Select

 R.MoveNext
Loop


PosAg = Left(PosAg, Len(PosAg) - 1)
 PosAgEp = Left(PosAgEp, Len(PosAgEp) - 1)
 
 PosPolEp = Left(PosPolEp, Len(PosPolEp) - 1)
 PosPol = Left(PosPol, Len(PosPol) - 1)
 
 AjAg = Left(AjAg, Len(AjAg) - 1)
 AjAgEp = Left(AjAgEp, Len(AjAgEp) - 1)
 
 AjPolEp = Left(AjPolEp, Len(AjPolEp) - 1)
 AjPol = Left(AjPol, Len(AjPol) - 1)




R.Close






Dim sql As String
Dim EIDISOZ As String
EIDISOZ = "DOKISOZ"

Dim ALLTIME As String
ALLTIME = " and HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND HME<='" + Format(gLhjh, "MM/DD/YYYY") + "'"

Dim PROHG As String
PROHG = " and HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND HME<'" + Format(apo.Value, "MM/DD/YYYY") + "'"

Dim Period As String
Period = " and HME>='" + Format(apo.Value, "MM/DD/YYYY") + "' AND HME<='" + Format(eos.Value, "MM/DD/YYYY") + "'"







On Error Resume Next
Gdb.Execute "drop table " + EIDISOZ
On Error GoTo neo_isoz_error

sql = " SELECT KOD,ONO,XTI,LTI,POS AS APOGPOS,POS AS AGPRPOS,POS AS AGPERPOS,POS AS POLPRPOS,POS AS POLPERPOS,POS AS YPOLPOS, "
sql = sql + "POS AS APOGAJ,POS AS AGPRAJ,POS AS AGPERAJ,POS As POLPRAJ, POS As POLPERAJ, POS As YPOLAJ"
sql = sql + " INTO " + EIDISOZ + " FROM EID  WHERE KOD IN (SELECT KODE FROM EGGTIM) "
Gdb.Execute sql


Gdb.Execute "UPDATE " + EIDISOZ + " set APOGPOS=0, AGPRPOS=0,AGPERPOS=0,POLPRPOS=0,POLPERPOS=0,YPOLPOS=0, "
sql = sql + "APOGAJ=0,AGPRAJ=0,AGPERAJ=0,POLPRAJ=0,POLPERAJ=0,YPOLAJ=0"
Gdb.Execute sql
'-----------  ΑΠΟΓΡΑΦΕΣ  -----------------------------------------------
'Gdb.Execute sql

sql = "UPDATE " + EIDISOZ + " SET APOGPOS=(SELECT SUM(POSO) FROM EGGTIM "
sql = sql + "WHERE KODE=" + EIDISOZ + ".KOD AND  LEFT(ATIM,1)='λ' " + ALLTIME + ")"
  
Gdb.Execute sql

'Gdb.Execute sql
sql = "UPDATE " + EIDISOZ + " SET APOGAJ=(SELECT SUM(POSO*TIMM*(100-EKPT)/100 ) FROM EGGTIM "
sql = sql + "WHERE KODE=" + EIDISOZ + ".KOD AND  LEFT(ATIM,1)='λ' " + ALLTIME + ")"
Gdb.Execute sql

'-------------  ΠΡΟΗΓΟΥΜΕΝΕΣ ΑΓΟΡΕΣ  -----------------------------------------------
sql = "UPDATE " + EIDISOZ + " SET AGPRPOS=(SELECT SUM( case when left(ATIM,1) IN (" + PosAg + ") then POSO ELSE -POSO END ) FROM EGGTIM "
sql = sql + " WHERE KODE=" + EIDISOZ + ".KOD  " + PROHG + " AND LEFT(ATIM,1) IN (" + PosAg + "," + PosAgEp + ") ) "
 Gdb.Execute sql
 

sql = "UPDATE " + EIDISOZ + " SET AGPRAJ=(SELECT SUM( case when left(ATIM,1) IN (" + AjAg + ") then POSO*TIMM*(100-EKPT)/100  ELSE -POSO*TIMM*(100-EKPT)/100  END ) FROM EGGTIM "
sql = sql + " WHERE KODE=" + EIDISOZ + ".KOD  " + PROHG + " AND LEFT(ATIM,1) IN (" + AjAg + "," + AjAgEp + ") ) "
 Gdb.Execute sql
 
'-------------  ΠΡΟΗΓΟΥΜΕΝΕΣ ΠΩΛΗΣΕΙΣ  -----------------------------------------------
'Gdb.Execute sql
sql = "UPDATE " + EIDISOZ + " SET POLPRPOS=(SELECT SUM( case when left(ATIM,1) IN (" + PosPol + ") then POSO ELSE -POSO END ) FROM EGGTIM "
sql = sql + " WHERE KODE=" + EIDISOZ + ".KOD  " + PROHG + " AND LEFT(ATIM,1) IN (" + PosPol + IIf(Len(PosPolEp) > 1, "," + PosPolEp, "") + ") ) "
 Gdb.Execute sql
 

sql = "UPDATE " + EIDISOZ + " SET POLPRAJ=(SELECT SUM( case when left(ATIM,1) IN (" + AjPol + ") then POSO*TIMM*(100-EKPT)/100  ELSE -POSO*TIMM*(100-EKPT)/100  END ) FROM EGGTIM "
sql = sql + " WHERE KODE=" + EIDISOZ + ".KOD  " + PROHG + " AND LEFT(ATIM,1) IN (" + AjPol + "," + AjPolEp + ") ) "
 Gdb.Execute sql
 
 '-------------  trexoyses ΑΓΟΡΕΣ  -----------------------------------------------
sql = "UPDATE " + EIDISOZ + " SET AGPERPOS=(SELECT SUM( case when left(ATIM,1) IN (" + PosAg + ") then POSO ELSE -POSO END ) FROM EGGTIM "
sql = sql + " WHERE KODE=" + EIDISOZ + ".KOD  " + Period + " AND LEFT(ATIM,1) IN (" + PosAg + IIf(Len(PosAgEp) > 1, "," + PosAgEp, "") + ") )"
 Gdb.Execute sql
 
sql = "UPDATE " + EIDISOZ + " SET AGPERAJ=(SELECT SUM( case when left(ATIM,1) IN (" + AjAg + ") then POSO*TIMM*(100-EKPT)/100  ELSE -POSO*TIMM*(100-EKPT)/100  END ) FROM EGGTIM "
sql = sql + " WHERE KODE=" + EIDISOZ + ".KOD  " + Period + " AND LEFT(ATIM,1) IN (" + AjAg + IIf(Len(AjAgEp) > 1, "," + AjAgEp, "") + ") )"
 Gdb.Execute sql
 
'-------------  trexoyses ΠΩΛΗΣΕΙΣ  -----POLPERPOS------------------------------------------
sql = "UPDATE " + EIDISOZ + " SET POLPERPOS=(SELECT SUM( case when left(ATIM,1) IN (" + PosPol + ") then POSO ELSE -POSO END ) FROM EGGTIM "
sql = sql + " WHERE KODE=" + EIDISOZ + ".KOD  " + Period + " AND LEFT(ATIM,1) IN (" + PosPol + IIf(Len(PosPolEp) > 1, "," + PosPolEp, "") + ") )"
 Gdb.Execute sql
 
' Gdb.Execute sql
 
 sql = "UPDATE " + EIDISOZ + " SET POLPERAJ=(SELECT SUM( case when left(ATIM,1) IN (" + AjPol + ") then POSO*TIMM*(100-EKPT)/100  ELSE -POSO*TIMM*(100-EKPT)/100  END ) FROM EGGTIM "
sql = sql + " WHERE KODE=" + EIDISOZ + ".KOD  " + Period + " AND LEFT(ATIM,1) IN (" + AjPol + IIf(Len(AjPolEp) > 1, "," + AjPolEp, "") + ") ) "
 Gdb.Execute sql
 


'-------------  υπολοιπα  -----------------------------------------------

 

 
sql = "UPDATE " + EIDISOZ + " SET YPOLPOS=ISNULL(APOGPOS,0)+ISNULL(AGPRPOS,0)+ISNULL(AGPERPOS,0) - ISNULL(POLPRPOS,0) - ISNULL(POLPERPOS,0) "
 Gdb.Execute sql
 
 sql = "UPDATE " + EIDISOZ + " SET YPOLAJ=YPOLPOS*(APOGAJ+AGPRAJ+AGPERAJ)/(APOGPOS+AGPRPOS+AGPERPOS ) WHERE APOGPOS+AGPRPOS+AGPERPOS >0"
 Gdb.Execute sql

Dim ANS2 As Integer, X As String

490     ANS2 = MsgBox("ΜΟΝΟ ΑΥΤΑ ΠΟΥ ΚΙΝΗΘΗΚΑΝ", vbYesNo)
        X = "SELECT KOD,ONO,APOGPOS AS [ΑΠΟΓ.ΠΟΣ],APOGAJ AS [ΑΠΟΓ.ΑΞΙΑ],AGPRPOS AS [ΑΓΟΡ.ΠΡΟΗΓ],AGPRAJ AS [ΑΞ.ΑΓ.ΠΡΟΗΓ],POLPRPOS AS [ΠΩΛ.ΠΡΟΗΓ],POLPRAJ AS [ΑΞ.ΠΩΛ.ΠΡΟΗΓ] ,AGPERPOS AS [ΑΓ.ΠΕΡΙΟΔ],AGPERAJ AS [AΞ.ΑΓ.ΠΕΡ],POLPERPOS AS [ΠΩΛ.ΠΕΡ],POLPERAJ AS [ΑΞ.ΠΩΛ.ΠΕΡ],YPOLPOS AS [ΥΠΟΛ.ΠΟΣ],YPOLAJ AS [ΥΠΟΛ.ΑΞΙΑ] FROM   " + EIDISOZ
500     If ANS2 = vbYes Then
510         X = X + "  WHERE AGPERPOS>0 OR POLPERPOS>0  "
        Else
520
        End If








530     Data1.ConnectionString = gConnect
540     Data1.RecordSource = X
550     Data1.Refresh

Dim sumes(39)
Dim k As Integer


         Do While Not Data1.Recordset.EOF
              
              For k = 0 To Data1.Recordset.FIELDS.Count - 1
                 If IsNumeric(Data1.Recordset.FIELDS(k).Value) And Data1.Recordset.FIELDS(k).Type <> 202 Then
                     sumes(k) = sumes(k) + Data1.Recordset.FIELDS(k).Value
                 End If
              Next
              Data1.Recordset.MoveNext
         Loop

Dim sum_str
sum_str = "0011111111111111111"
         For k = 1 To Data1.Recordset.FIELDS.Count - 1

             If mID(sum_str, k, 1) = "1" Then
                 GR2.Splits(0).columns(k).FooterText = Format(sumes(k), "######0.00")
             End If

         Next

     GR2.AlternatingRowStyle = True
     GR2.OddRowStyle.BackColor = &H8000000F   ' GRI   vbCyan
     GR2.EvenRowStyle.BackColor = &HFFFFC0   'OYRANI     &H8000000F  ' GRI











560     If Check1 Then

            'typos2
          '  On Error GoTo Command8_Click_Err

570         print3_xar UCase(X), sum_str, "", 0
        End If

        'If Check2 Then
        '  Print4_Excel UCase(x), "00001111111111111111"
        'End If




















neo_isoz_error:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.neo_isoz " & "at line " & Erl

        Resume Next



End Sub





















Private Sub GR2_HeadClick(ByVal ColIndex As Integer)

        Dim sumes(100) As String
Dim k As Integer

100     For k = 0 To Data1.Recordset.FIELDS.Count - 1
110         sumes(k) = GR2.Splits(0).columns(k).FooterText    '  = Format(SUMES(k), "######0.00")
        Next

120     If Data1.Recordset.sort = "[" & GR2.columns(ColIndex).DataField & "] asc" Then   ' strSort
130         Data1.Recordset.sort = "[" & GR2.columns(ColIndex).DataField & "] desc"    ' strSort
        Else
140         Data1.Recordset.sort = "[" & GR2.columns(ColIndex).DataField & "] asc"    ' strSort
        End If

150     For k = 0 To Data1.Recordset.FIELDS.Count - 1
160         GR2.Splits(0).columns(k).FooterText = sumes(k)   '  = Format(SUMES(k), "######0.00")
        Next

        '    Adodc2.Recordset.Sort = "[" & TDBGrid.Columns(ColIndex).DataField & "] desc" ' strSort

        '<EhFooter>
        Exit Sub








End Sub

Private Sub ISOZYGIO_Click()


neo_isoz
Exit Sub







        '<EhHeader>
        On Error GoTo ISOZYGIO_Click_Err

        '</EhHeader>

        'ΕΚΤΥΠΩΣΗ ΣΤΟ ΣΥΝΟΛΟ ΤΩΝ ΑΠΟΘΗΚΩΝ

        Dim a As String

        Dim B As String

        Dim D As Long

100     Me.MousePointer = vbHourglass



   Dim mnow As String
   
   On Error Resume Next
   mnow = "EID"
   If Left(mnow, 3) = "DOK" Then
            mnow = "DOKEID" + Format(Now, "YYMMDDHHmm")
            Gdb.Execute "DROP TABLE " + mnow
            Gdb.Execute "SELECT * into " + mnow + "  FROM EID"
    End If
    





        ' UPDATE_YPOLOIPA
110     UPDATE_YPOLOIPA3 mnow, List11, apo, eos

120     Fsql = Text1.Text

        Dim ANS

130     If F_epan = 1 Then

            '   ANS = MsgBox("Να υπολογίσω το κέρδος ξανά;", vbYesNo)
            '  If ANS = vbYes Then
            '    ypol_kerd
            '  End If
        End If

        Dim polhs As String

        Dim AGOEPIS As String

        Dim polepis As String, pol As String, ago As String, APIS

140     a = Get_AJ(polhs, polepis, ago, AGOEPIS)

        'polhs = "'L','t','T','y','Y','υ'"

        Dim apothkes As String

150     If Val(Text3.Text) = 99 Then
160         apothkes = ""
        Else
170         apothkes = " and APOT=" + Text3.Text
        End If

        '
        ' On Error Resume Next    AS1P>0

        On Error GoTo ISOZYGIO_Click_Err

        '-------------------------------- NEO ARXH ===============================

180     Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1

190     a = "SELECT KODE," & " SUM( CASE   WHEN LEFT(ATIM,1)='λ'  THEN XRE ELSE 0  END ) AS S0X," & " SUM( CASE   WHEN LEFT(ATIM,1) IN ('G','g')  THEN XRE ELSE 0  END ) AS S1X," & " SUM( CASE WHEN PIS IS NULL  OR LEFT(ATIM,1) NOT IN (" + polhs + ")  THEN 0 ELSE PIS END) AS S1P," & " SUM( CASE WHEN XRE*TIMM IS NULL THEN 0 ELSE XRE*TIMM*(1- (CASE WHEN EKPT IS NULL THEN 0 ELSE EKPT END)/100) END  ) AS AS1X," & " SUM( CASE WHEN POSO*TIMM IS NULL OR LEFT(ATIM,1) NOT IN (" + polhs + ")  THEN 0 ELSE POSO*TIMM*(1-(CASE WHEN EKPT IS NULL THEN 0 ELSE EKPT END)/100) END ) AS AS1P, " & " SUM( CASE WHEN KERDOS IS NULL OR LEFT(ATIM,1) NOT IN (" + polhs + ")  THEN 0 ELSE KERDOS END ) AS KERD1 INTO DOKEGGT1 "
200     a = a + " FROM EGGTIM WHERE  left(ATIM,1) NOT IN ('σ','Σ') AND  " + " HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' " + apothkes + "   GROUP BY KODE"

        ' ' & " SUM(XRE)-SUM(PIS) AS YP," _

            Gdb.Execute a, D

210     If D = 0 Then
220         MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"
230         Me.MousePointer = vbNormal

            Exit Sub

        End If

        '
240     a = "SELECT LEFT(KOD,14)AS [ΚΩΔ],LEFT(EID.PROM,10) AS [ΠΡΟΜΗΘ],(SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=11 AND AYJON=EID.AEG) AS [ΟΙΚΟΓ] ,"

250     a = a + "(SELECT TOP 1 PERIGRAFH FROM PINAKES WHERE TYPOS=3 AND TIMH=EID.AEG AND CAST(KODLOG AS INT)=AYJON ) AS [ΥΠΟ-ΟΙΚ],LEFT(EID.ONO,35) AS [ΠΕΡΙΓΡΑΦΗ],"

260     a = a + "S0X AS [ΑΠΟΓΡ], S1X AS [ΑΓΟΡΕΣ],S1P AS [ΠΩΛΗΣΕΙΣ],S0X+S1X-S1P AS [ΥΠΟΛ], "
270     a = a + "AS1X AS [ΑΞ.ΑΓΟΡ],AS1P AS [ΑΞ.ΠΩΛ] "

280     a = a + " INTO DOKEGGT3 FROM EID  LEFT JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD  "

290     If Len(Fsql) > 1 Then
300         a = a + " where " + Fsql
        End If

310     If ORDERCOMBO.Text = ORDERCOMBO.List(0) Then
320         a = a + " ORDER BY EID.ONO"
        End If

330     If ORDERCOMBO.Text = ORDERCOMBO.List(1) Then
340         a = a + " ORDER BY KOD"
        End If

350     If ORDERCOMBO.Text = ORDERCOMBO.List(2) Then
360         a = a + " ORDER BY EID.ERG"
        End If

370     If ORDERCOMBO.Text = ORDERCOMBO.List(3) Then
380         a = a + " ORDER BY AS1P-S1P*EID.XTI "  ' [KEPΔΟΣ]"
        End If

        ' Exit Sub

        On Error Resume Next

390     Gdb.Execute "DROP TABLE DOKEGGT3"  'SOUMES ΑΠΟΘΗΚΗΣ 1

        On Error GoTo LATOS

400     Gdb.Execute a, D
410     Me.MousePointer = vbNormal

420     If D = 0 Then
430         MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT3"
            GoTo mexit2
            Exit Sub

        End If

        Dim X As String

        Dim ANS2

        Dim synt As String

440     ANS2 = MsgBox("ΜΟΝΟ ΑΥΤΑ ΠΟΥ ΚΙΝΗΘΗΚΑΝ", vbYesNo)

450     If ANS2 = vbYes Then
460         synt = " WHERE [ΑΓΟΡΕΣ]>0 OR [ΠΩΛΗΣΕΙΣ]>0  OR [ΑΞ.ΠΩΛ]>0"
            '    X = "SELECT * FROM DOKEGGT3 WHERE [ΑΓΟΡΕΣ]>0 OR [ΠΩΛΗΣΕΙΣ]>0  OR [ΑΞ.ΠΩΛ]>0 ORDER BY [ΟΙΚΟΓ],[ΥΠΟ-ΟΙΚ]"
        Else
470         synt = ""
            '    X = "SELECT * FROM DOKEGGT3 ORDER BY [ΟΙΚΟΓ],[ΥΠΟ-ΟΙΚ]"
        End If

480     ANS2 = MsgBox("ΣΥΝΟΛΑ ΑΝΑ ΟΙΚΟΓΕΝΕΙΑ;", vbYesNo)

490     If ANS2 = vbYes Then

500         X = "SELECT  [ΟΙΚΟΓ],[ΥΠΟ-ΟΙΚ],SUM([ΑΓΟΡΕΣ]) AS [ΑΓΟΡΕΣ],SUM([ΠΩΛΗΣΕΙΣ]) AS [ΠΩΛΗΣΕΙΣ], SUM([ΑΞ.ΑΓΟΡ]) AS [ΑΞ.ΑΓΟΡΩΝ],SUM([ΑΞ.ΠΩΛ]) AS [ΑΞ.ΠΩΛΗΣ]  FROM DOKEGGT3 " + synt + " GROUP BY  [ΟΙΚΟΓ],[ΥΠΟ-ΟΙΚ]  ORDER BY [ΟΙΚΟΓ],[ΥΠΟ-ΟΙΚ]"
        Else
510         X = "SELECT * FROM DOKEGGT3 " + synt + " ORDER BY [ΟΙΚΟΓ],[ΥΠΟ-ΟΙΚ]"
        End If

520     Data1.ConnectionString = gConnect
530     Data1.RecordSource = X
540     Data1.Refresh

550     If Check1 Then

            'typos2
            On Error GoTo ISOZYGIO_Click_Err

560         print3_xar UCase(X), "000011111221111111111", "", 0
        End If

        'If Check2 Then
        '  Print4_Excel UCase(x), "00001111111111111111"
        'End If


mexit2:
If Left(mnow, 3) = "DOK" Then
   Gdb.Execute "drop TABLE " + mnow
End If



        Exit Sub

LATOS:
570     MsgBox Err.Description

580     Resume Next

        'End Sub

        ' ΙΣΟΖΥΓΙΟ ΜΗΝΟΣ ΜΕ ΠΟΣΟΤΗΤΕΣ ΚΑΙ ΑΞΙΕΣ ΚΑΙ  ΜΕ ΠΡΟΗΓΟΥΜΕΝΑ ΥΠΟΛΟΙΠΑ
        ' M01   M02 APOGRAFH
        'M03 M04  AGORES PROHGOYMENES
        'M05 M06 POLISEIS PROHGOYMENES
        'M07 M08  AGORES PERIODOY
        'M09  M10 POLISEIS PERIODOY
        'M11  M12  YPOLOIPA

        'ΥΠΟΛΟΓΙΖΕΙ ΤΟ ΡΕΠΟΡΤ ΕΙΤΕ
        ' ΑΠΟ ΑΥΤΑ ΠΟΥ ΕΠΕΛΕΞΑ ΕΙΤΕ
        ' ΑΠΟ ΤΟ ΕΤΟΙΜΟ ΡΕΠΟΡΤ ΠΟΥ ΕΧΩ ΕΠΙΛΕΞΕΙ

        'Dim X As String, order As String
        'Dim k As Integer, R As New ADODB.Recordset ', sum As Single
        'Dim mc As Integer
        'Dim EPIK(38), alig(38), mWidth(38)
        'Dim sum(30) As Single
        'Dim TIM As Long
        'Dim SUgm_str  'που θα κανει sumes
        'Dim Kinhseis As Boolean
        'Dim A
        '
        'Kinhseis = False
        '
        '' για να βγαζει combo στο grid
        'Dim gm_fpa As Integer
        'Dim m_kat As Integer
        'm_kat = 0 ' krato thn stili tis katigorias
        'gm_fpa = 0 ' krato thn stili toy FPA
        '
        '
        '
        'Me.MousePointer = vbHourglass
        'TIM = GetCurrentTime()
        '
        'For k = 0 To 9: sum(k) = 0: Next
        '
        '
        'For k = 1 To 18: mWidth(k) = 1: alig(k) = 0: EPIK(k) = "": Next
        'F_RUN = 1
        '
        'data1.ConnectionString = gConnect
        'UPDATE_YPOLOIPA3
        '
        'Me.Caption = GetCurrentTime() - TIM
        '
        'mc = 2
        '
        '  SUgm_str = "000"
        '  X = "select  '@@@@@' AS [A/A],KOD AS [Κωδικός],LEFT(ONO,35) as [Περιγραφή],ERG AS [BARCODE],AEG AS [OIK],KODLOG AS[YΠΟΟΙΚ]," '
        'X = X + "DM01 AS [ΑΠΟΓ],DM02 AS [ΑΞΙΑ ΑΠΟΓΡΑΦΗΣ],"
        'X = X + "DM03 AS [ΑΓ.ΠΡ],DM04 AS [ΑΞΙΑ ΑΓΟΡΩΝ],"
        'X = X + "DM05 AS [ΠΩΛ],DM06 AS [ΑΞΙΑ ΠΩΛΗΣΕΩΝ],"
        'X = X + "DM07 AS [ΑΓ.ΠΕΡ],DM08 AS [ΑΞΙΑ ΑΓΟΡΩΝ ΠΕΡΙΟΔΟΥ],"
        'X = X + "DM09 AS [ΠΩΛ.ΠΕΡ],DM10 AS [ΑΞΙΑ ΠΩΛΗΣΕΩΝ ΠΕΡΙΟΔΟΥ],"
        'X = X + "DM01+DM03-DM05+DM07-DM09 AS [ΥΠΟΛ],(DM01+DM03-DM05+DM07-DM09)*EID.XTI AS [YΠΟΛΟΙΠΟ ΑΞΙΑΣ]  "
        '
        'Dim KK
        '
        '
        '
        'Dim temp
        ''
        '
        '
        ' Kinhseis = True
        '
        ''-------------------------------- NEO ARXH ===============================
        'If Kinhseis Then
        '
        '    Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1
        '    Gdb.Execute "UPDATE EGGTIM SET EKPT=0 WHERE EKPT IS NULL"
        '
        '
        'A = "SELECT KODE, SUM( CASE WHEN MONTH(HME)=1 AND DAY(HME)=1  THEN XRE ELSE 0  END ) AS DM01," _
        '& " SUM( CASE WHEN MONTH(HME)=1 AND DAY(HME)=1  THEN XRE*EID.XTI ELSE 0  END ) AS DM02," _
        '& " SUM( CASE  WHEN HME>='01/02/" + Format(Year(apo), "0000") + "' AND HME<'" + Format(apo, "MM/DD/YYYY") + "'  THEN XRE ELSE 0  END ) AS DM03," _
        '& " SUM( CASE  WHEN HME>='01/02/" + Format(Year(apo), "0000") + "' AND HME<'" + Format(apo, "MM/DD/YYYY") + "'  THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS DM04," _
        '& " SUM( CASE  WHEN HME>='01/02/" + Format(Year(apo), "0000") + "' AND HME<'" + Format(apo, "MM/DD/YYYY") + "'  THEN PIS ELSE 0  END ) AS DM05," _
        '& " SUM( CASE  WHEN HME>='01/02/" + Format(Year(apo), "0000") + "' AND HME<'" + Format(apo, "MM/DD/YYYY") + "'  THEN PIS*TIMM*(1-EKPT/100) ELSE 0  END ) AS DM06," _
        '& " SUM( CASE  WHEN HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'  THEN XRE ELSE 0  END ) AS DM07," _
        '& " SUM( CASE  WHEN HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'  THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS DM08," _
        '& " SUM( CASE  WHEN HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'  THEN PIS ELSE 0  END ) AS DM09," _
        '& " SUM( CASE  WHEN HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'  THEN PIS*TIMM*(1-EKPT/100) ELSE 0  END ) AS DM10 " _
        '& " INTO DOKEGGT1 "
        'A = A + " FROM EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD WHERE (NOT APOT IS NULL) " _
        '& "  GROUP BY KODE"
        '
        ''   On Error GoTo ISOZYGIO_Click_Err
        '
        '  'Gdb.Execute "UPDATE EGGTIM SET PIS=0 WHERE PIS IS NULL"
        '  Gdb.Execute "UPDATE EGGTIM SET EKPT=0 WHERE PIS IS NULL"
        '
        '   Gdb.Execute A, k
        '
        '
        '
        '
        '
        '   X = X + " FROM EID RIGHT JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD "
        'Else
        '   X = X + " from EID  inner JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD"
        'End If
        '
        '
        'If Len(Text1.Text) > 0 Then
        '   X = X + " WHERE DM01+DM02+DM03+DM04+DM05+DM06+DM07+DM08+DM09+DM10>0 " + " and " + Fsql
        'End If
        '
        '
        'order = "kod"
        'For k = 0 To 2
        '  If Combo2.Text = Combo2.List(0) Then
        '     order = "EID.KOD"
        '  ElseIf Combo2.Text = Combo2.List(1) Then
        '     order = "ONO"
        '  ElseIf Combo2.Text = Combo2.List(2) Then
        '     order = "POS"
        '  ElseIf Combo2.Text = Combo2.List(3) Then
        '     order = "PROM"
        '  ElseIf Combo2.Text = Combo2.List(4) Then
        '     order = "XTI"
        '  ElseIf Combo2.Text = Combo2.List(5) Then
        '     order = "XTI*POS"
        ' ElseIf Combo2.Text = Combo2.List(6) Then
        '     order = "FPA,ONO"
        ' ElseIf Combo2.Text = Combo2.List(7) Then
        '     order = "AEG,KODLOG,ONO"
        '
        '
        '
        '
        '
        '  End If
        'Next
        '
        '
        '
        '
        'X = X + " order by " + order
        '
        '
        'If rCheck.value = vbChecked Then ' AN PAIRNV TO EPILEGMENO REPORT
        '   X = Text4.Text
        'Else
        '   Text4.Text = X
        'End If
        '
        '
        '
        'Gdb.Execute "UPDATE DOKEGGT1 SET DM01=0 WHERE DM01 IS NULL"
        'Gdb.Execute "UPDATE DOKEGGT1 SET DM02=0 WHERE DM02 IS NULL"
        'Gdb.Execute "UPDATE DOKEGGT1 SET DM03=0 WHERE DM03 IS NULL"
        'Gdb.Execute "UPDATE DOKEGGT1 SET DM04=0 WHERE DM04 IS NULL"
        'Gdb.Execute "UPDATE DOKEGGT1 SET DM05=0 WHERE DM05 IS NULL"
        'Gdb.Execute "UPDATE DOKEGGT1 SET DM06=0 WHERE DM06 IS NULL"
        'Gdb.Execute "UPDATE DOKEGGT1 SET DM07=0 WHERE DM07 IS NULL"
        'Gdb.Execute "UPDATE DOKEGGT1 SET DM08=0 WHERE DM08 IS NULL"
        'Gdb.Execute "UPDATE DOKEGGT1 SET DM09=0 WHERE DM09 IS NULL"
        'Gdb.Execute "UPDATE DOKEGGT1 SET DM10=0 WHERE DM10 IS NULL"
        '
        '
        '
        '
        '
        '
        'data1.RecordSource = X
        '
        'On Error GoTo lathos
        '   data1.Refresh
        'On Error Resume Next
        '
        '
        '
        '    Dim xd As New TrueOleDBGrid80.ValueItem
        '' On Error GoTo ISOZYGIO_Click_Err
        '
        '
        'If gm_fpa > 0 Then
        '     xd.value = 1
        '     xd.DisplayValue = "11%"
        '     GR2.Columns(gm_fpa).ValueItems.Add xd
        '
        '     xd.value = 2
        '     xd.DisplayValue = "23%"
        '     GR2.Columns(gm_fpa).ValueItems.Add xd
        '
        '    GR2.Columns(gm_fpa).ValueItems.Translate = True
        '     GR2.Columns(gm_fpa).ValueItems.Presentation = dbgComboBox
        'End If
        '
        'If m_kat > 0 Then
        '    Dim r7 As New ADODB.Recordset
        '    r7.Open "SELECT *FROM PINAKES WHERE TYPOS=3 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
        '    Do While Not r7.EOF
        '       If r7("typos") = 3 Then
        '          xd.value = Format(r7("ayjon"), "000")
        '          xd.DisplayValue = r7("PERIGRAFH")
        '          GR2.Columns(m_kat).ValueItems.Add xd
        '       Else
        '          Exit Do
        '       End If
        '       r7.MoveNext
        '    Loop
        '    ' kathgories
        '
        '    R.Close
        '
        '
        '
        '
        '
        '
        '     'xd.Value = 1
        '     'xd.DisplayValue = "9%"
        '     'GR2.Columns(gm_fpa).ValueItems.Add xd
        '
        '
        '    GR2.Columns(m_kat).ValueItems.Translate = True
        '     GR2.Columns(m_kat).ValueItems.Presentation = dbgComboBox
        'End If
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        ''GR2.Columns(0).Text = "Κωδικός"
        ''GR2.Columns(1).Text = "Περιγραφή"
        ''GR2.Columns(2).Text = "Ποσότητα"
        '
        '
        '
        'For k = 5 To 12
        '    If Len(EPIK(k)) > 1 Then
        ' '      GR2.Columns(K).Text = epik(K)
        '       GR2.Columns(k).Alignment = alig(k)
        '       GR2.Columns(k).width = mWidth(k)
        '
        '     End If
        'Next
        '
        '       GR2.Columns(0).width = 1200
        '       GR2.Columns(1).width = 2000
        '       GR2.Columns(2).width = 1000
        '
        'Me.MousePointer = vbNormal
        '
        'If Check1 Then
        '  'typos2
        '  SUgm_str = "0111111111111111"
        '  On Error GoTo ISOZYGIO_Click_Err
        '  If Len(SUgm_str) < 3 Then SUgm_str = "0111111111111111"
        '  Text2.Text = "ΙΣΟΖΥΓΙΟ ΑΠΟΘΗΚΗΣ ΑΠΟ " + Format(apo, "DD/MM/YYYY") + " ΕΩΣ " + Format(eos, "DD/MM/YYYY")
        '  print3_xar UCase(X), SUgm_str, Text2.Text, 0   ' RR.RecordSource
        'End If
        '
        '
        ''If Check2 Then
        ' ' Print4_Excel UCase(x), SUgm_str
        ''End If
        '
        'Exit Sub
        '
        'lathos:
        '
        'MsgBox "λάθος " + Err.Description
        '
        '
        'Resume Next

        '<EhFooter>
        Exit Sub

ISOZYGIO_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.ISOZYGIO_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.ISOZYGIO_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command5_Click()

        '<EhHeader>
        On Error GoTo Command5_Click_Err

        '</EhHeader>
        Dim N As Integer

100     N = 0
110     Check3.Value = vbUnchecked

120     If Not FIsHidden Then

130         With KRITHRIA.Grid1

140             N = N + 1: .TextMatrix(N, 0) = "Περιγραφή": .TextMatrix(N, 3) = "ONO": .TextMatrix(N, 4) = "c"
150             N = N + 1: .TextMatrix(N, 0) = "Οικογένεια": .TextMatrix(N, 3) = "AEG": .TextMatrix(N, 4) = "n"
160             N = N + 1: .TextMatrix(N, 0) = "Υποοικογένεια": .TextMatrix(N, 3) = "KODLOG": .TextMatrix(N, 4) = "n"

170             N = N + 1: .TextMatrix(N, 0) = "Υπόλοιπο": .TextMatrix(N, 3) = "isnull(POS,0)": .TextMatrix(N, 4) = "n"
180             N = N + 1: .TextMatrix(N, 0) = "Υπόλ.Αποθ1": .TextMatrix(N, 3) = "POS01": .TextMatrix(N, 4) = "n"
190             N = N + 1: .TextMatrix(N, 0) = "Υπόλ.Αποθ2": .TextMatrix(N, 3) = "POS02": .TextMatrix(N, 4) = "n"

200             N = N + 1: .TextMatrix(N, 0) = "Υπόλ.Αποθ3": .TextMatrix(N, 3) = "POS03": .TextMatrix(N, 4) = "n"
210             N = N + 1: .TextMatrix(N, 0) = "Υπόλ.Αποθ4": .TextMatrix(N, 3) = "POS04": .TextMatrix(N, 4) = "n"
            
220             N = N + 1: .TextMatrix(N, 0) = "Πωλ.Ποσ.Αποθ1": .TextMatrix(N, 3) = "S1P": .TextMatrix(N, 4) = "n"
230             N = N + 1: .TextMatrix(N, 0) = "Πωλ.Ποσ.Αποθ2": .TextMatrix(N, 3) = "S22": .TextMatrix(N, 4) = "n"
240             N = N + 1: .TextMatrix(N, 0) = "Πωλ.Ποσ.Αποθ3": .TextMatrix(N, 3) = "S3P": .TextMatrix(N, 4) = "n"
250             N = N + 1: .TextMatrix(N, 0) = "Πωλ.Ποσ.Αποθ4": .TextMatrix(N, 3) = "S4P": .TextMatrix(N, 4) = "n"

260             N = N + 1: .TextMatrix(N, 0) = "Προμηθευτής": .TextMatrix(N, 3) = "PROM": .TextMatrix(N, 4) = "c"
270             N = N + 1: .TextMatrix(N, 0) = "BARCODE": .TextMatrix(N, 3) = "ERG": .TextMatrix(N, 4) = "c"
280             N = N + 1: .TextMatrix(N, 0) = "Κωδικός": .TextMatrix(N, 3) = "KOD": .TextMatrix(N, 4) = "c"

290             N = N + 1: .TextMatrix(N, 0) = "Τιμή Αγοράς": .TextMatrix(N, 3) = "XTI": .TextMatrix(N, 4) = "n"
300             N = N + 1: .TextMatrix(N, 0) = "Τιμή Λιανικής": .TextMatrix(N, 3) = "LTI5": .TextMatrix(N, 4) = "n"

310             N = N + 1: .TextMatrix(N, 0) = "Τιμή Χονδρικής": .TextMatrix(N, 3) = "LTI": .TextMatrix(N, 4) = "n"

320             N = N + 1: .TextMatrix(N, 0) = "Κατ.ΦΠΑ": .TextMatrix(N, 3) = "FPA": .TextMatrix(N, 4) = "n"

330             N = N + 1: .TextMatrix(N, 0) = "Πόντοι": .TextMatrix(N, 3) = "PONTOI": .TextMatrix(N, 4) = "n"

340             N = N + 1: .TextMatrix(N, 0) = "ΣτάθμηΠαραγγ": .TextMatrix(N, 3) = "SPA": .TextMatrix(N, 4) = "n"
350             N = N + 1: .TextMatrix(N, 0) = "Επιθ.Υπόλ": .TextMatrix(N, 3) = "EPIUYP": .TextMatrix(N, 4) = "n"

360             N = N + 1: .TextMatrix(N, 0) = F_LHM1: .TextMatrix(N, 3) = "HM1": .TextMatrix(N, 4) = "d"
370             N = N + 1: .TextMatrix(N, 0) = F_LHM2: .TextMatrix(N, 3) = "HM2": .TextMatrix(N, 4) = "d"
380             N = N + 1: .TextMatrix(N, 0) = f_lHM3: .TextMatrix(N, 3) = "HM3": .TextMatrix(N, 4) = "d"

390             N = N + 1: .TextMatrix(N, 0) = f_lab1: .TextMatrix(N, 3) = "CH1": .TextMatrix(N, 4) = "c"
400             N = N + 1: .TextMatrix(N, 0) = f_lab2: .TextMatrix(N, 3) = "CH2": .TextMatrix(N, 4) = "c"
410             N = N + 1: .TextMatrix(N, 0) = f_lab3: .TextMatrix(N, 3) = "CH3": .TextMatrix(N, 4) = "c"
420             N = N + 1: .TextMatrix(N, 0) = f_lab4: .TextMatrix(N, 3) = "CH4": .TextMatrix(N, 4) = "c"

430             N = N + 1: .TextMatrix(N, 0) = f_lab5: .TextMatrix(N, 3) = "CH5": .TextMatrix(N, 4) = "c"
440             N = N + 1: .TextMatrix(N, 0) = f_lab6: .TextMatrix(N, 3) = "CH6": .TextMatrix(N, 4) = "c"

450             N = N + 1: .TextMatrix(N, 0) = f_num1: .TextMatrix(N, 3) = "NUM1": .TextMatrix(N, 4) = "N"
460             N = N + 1: .TextMatrix(N, 0) = f_num2: .TextMatrix(N, 3) = "NUM2": .TextMatrix(N, 4) = "N"
470             N = N + 1: .TextMatrix(N, 0) = f_num3: .TextMatrix(N, 3) = "NUM3": .TextMatrix(N, 4) = "N"

480             N = N + 1: .TextMatrix(N, 0) = "Περιγραφή 2η": .TextMatrix(N, 3) = "ONO": .TextMatrix(N, 4) = "c"

            End With

490         FIsHidden = True

        End If

500     KRITHRIA.SHOW 1
510     Fsql = MDIForm1.StatusBar1.Panels(6).Text
520     Text1.Text = Fsql

530     MDIForm1.StatusBar1.Panels(6).Text = ""

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command6_Click()

        '<EhHeader>
        On Error GoTo Command6_Click_Err

        '</EhHeader>

        'ΕΚΤΥΠΩΣΗ ΚΑΤΑ ΑΠΟΘΗΚΗ
        Dim a As String

        Dim B As String

        Dim D As Long

        Dim NAPOT

        Dim R As New ADODB.Recordset

100     R.Open "SELECT COUNT(*) AS DDD FROM PINAKES WHERE TYPOS=4 ", Gdb, adOpenDynamic, adLockOptimistic
110     NAPOT = R(0)
120     R.Close

        ' UPDATE_YPOLOIPA
130     UPDATE_YPOLOIPA3 "EID", List11, apo, eos

        '
        ' On Error Resume Next

        On Error GoTo Command6_Click_Err

        '-------------------------------- NEO ARXH ===============================

140     Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1

150     a = "SELECT KODE," _
           & " SUM( CASE APOT  WHEN 1  THEN XRE ELSE 0  END ) AS S1X," _
           & " SUM(CASE APOT  WHEN 1  THEN PIS ELSE 0  END ) AS S1P," _
           & " SUM( CASE APOT  WHEN 2 THEN XRE ELSE 0  END ) AS S2X," _
           & " SUM( CASE APOT WHEN 2 THEN PIS ELSE 0 END ) AS S2P," _
           & " SUM( CASE APOT  WHEN 3  THEN XRE ELSE 0  END ) AS S3X," _
           & " SUM(CASE APOT  WHEN 3  THEN PIS ELSE 0  END ) AS S3P," _
           & " SUM( CASE APOT  WHEN 4 THEN XRE ELSE 0  END ) AS S4X," _
           & " SUM( CASE APOT WHEN 4 THEN PIS ELSE 0 END ) AS S4P," _
           & " SUM( CASE APOT  WHEN 1  THEN XRE*TIMM ELSE 0  END ) AS AS1X," _
           & " SUM(CASE APOT  WHEN 1  THEN PIS*TIMM ELSE 0  END ) AS AS1P," _
           & " SUM( CASE APOT  WHEN 2 THEN XRE*TIMM ELSE 0  END ) AS AS2X," _
           & " SUM( CASE APOT WHEN 2 THEN PIS*TIMM ELSE 0 END ) AS AS2P," _
           & " SUM( CASE APOT  WHEN 3 THEN XRE*TIMM ELSE 0  END ) AS AS3X," _
           & " SUM( CASE APOT WHEN 3 THEN PIS*TIMM ELSE 0 END ) AS AS3P, " _
           & " SUM( CASE APOT  WHEN 4 THEN XRE*TIMM ELSE 0  END ) AS AS4X," _
           & " SUM( CASE APOT WHEN 4 THEN PIS*TIMM ELSE 0 END ) AS AS4P " _
           & " INTO DOKEGGT1 "
160     a = a + " FROM EGGTIM WHERE " + " HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' GROUP BY KODE"
        ' Exit Sub
        'Exit Sub
170     Gdb.Execute a, D

180     If D = 0 Then
190         MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"

            Exit Sub

        End If

        '
200     a = "SELECT KOD,LEFT(EID.PROM,20) AS [ΠΡΟΜΗΘ],EID.ERG AS [BARCODE],LEFT(EID.ONO,35) AS [ΠΕΡΙΓΡΑΦΗ],"
210     a = a + " S1X AS [ΑΓ.ΠΟΣ1],S1P AS [ΠΩΛ.ΠΟΣ1],POS01 AS [ΥΠΟΛ.1],"
220     a = a + " S2X AS [ΑΓ.ΑΠ2] ,S2P AS [ΠΩΛ.ΠΟΣ2],POS02 AS [ΥΠ.ΑΠ2],"

230     If NAPOT > 2 Then    'FAKIS
240         a = a + " S3X AS [ΑΓ.ΠΟΣ3],S3P AS [ΠΩΛ.ΠΟΣ3],POS03 AS [ΥΠΟΛ.3],"
250         a = a + " S4X AS [ΑΓ.ΑΠ4] ,S4P AS [ΠΩΛ.ΠΟΣ4],POS04 AS [ΥΠ.ΑΠ4],"
        End If

260     a = a + " AS1X AS [ΑΞ.ΑΓ1],AS1P AS [ΑΞ.ΠΩΛ.ΑΠ1],AS2X AS [ΑΞ.ΑΓ2],AS2P AS [ΑΞ.ΠΩΛ.ΑΠ2] "

270     If NAPOT > 2 Then    ' FAKIS
280         a = a + " ,AS3X AS [ΑΞ.ΑΓ3],AS3P AS [ΑΞ.ΠΩΛ.ΑΠ3],AS4X AS [ΑΞ.ΑΓ4],AS4P AS [ΑΞ.ΠΩΛ.ΑΠ4] "
        End If

290     a = a + " INTO DOKEGGT3 FROM EID  LEFT JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD"
        '  ,(AS2P+AS1P)-(S2P+S1P)*EID.XTI AS [ΚΕΡΔΟΣ]

300     If Len(Fsql) > 1 Then
310         a = a + " WHERE " + Fsql
        End If

320     If ORDERCOMBO.Text = ORDERCOMBO.List(0) Then
330         a = a + " ORDER BY EID.ONO"
        End If

340     If ORDERCOMBO.Text = ORDERCOMBO.List(1) Then
350         a = a + " ORDER BY KOD"
        End If

360     If ORDERCOMBO.Text = ORDERCOMBO.List(3) Then
370         a = a + " ORDER BY EID.ERG"
        End If

        ' Exit Sub

        On Error Resume Next

380     Gdb.Execute "DROP TABLE DOKEGGT3"  'SOUMES ΑΠΟΘΗΚΗΣ 1

        On Error GoTo Command6_Click_Err

390     Gdb.Execute a, D

400     If D = 0 Then
410         MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT3"

            Exit Sub

        End If

        Dim X As String

420     X = "SELECT * FROM DOKEGGT3"

430     If Check1 Then
            'typos2
440         print3_xar UCase(X), "00001111111111111111", "", 0
        End If

        'If Check2 Then
        '  Print4_Excel UCase(x), "00001111111111111111"
        'End If

        Exit Sub

        'CrystalReport1.Connect = gConnect

        'CrystalReport1.Action = 1
        Dim DUM

        'DUM = Shell("bcp emp.dbo.DOKEGGT3 out region.txt -S PLATO -T -c  -C RAW", vbMinimizedFocus)
        ' On Error Resume Next
450     FileCopy "C:\LAGEURO\DOKEGGT3SOS.DBF", "C:\LAGEURO\DOKEGGT3.DBF"

460     update_DBF_from_SQL "SELECT *FROM DOKEGGT3", "DOKEGGT3"

470     MsgBox "ΑΡΧΕΙΟ EXCEL:" + Chr(13) + "C:\LAGEURO\DOKEGGT3.DBF"

        Exit Sub

480     D = D + 1
        '-------------------------------- NEO TELOS ===============================

        '---------------------- ΚΙΝΗΣΕΙΣ ---------------------------------------------
        On Error Resume Next

490     Gdb.Execute "DROP TABLE DOKEGGT1"
500     Gdb.Execute "DROP TABLE DOKEGGT2"
510     Gdb.Execute "DROP TABLE DOKEGGT3"

520     a = "SELECT KODE,SUM(XRE) AS S1X,SUM(PIS) AS S1P,0 AS S31,0 AS S32 INTO DOKEGGT1 "
530     a = a + " FROM EGGTIM WHERE APOT=1 AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' GROUP BY KODE"

540     B = "SELECT KODE,SUM(XRE) AS S2X,SUM(PIS) AS S2P INTO DOKEGGT2 "
550     B = B + " FROM EGGTIM WHERE APOT=2  AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'  GROUP BY KODE"

560     Gdb.Execute a, D

570     If D = 0 Then
580         MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"

            Exit Sub

        End If

590     D = 0

600     Gdb.Execute B, D

610     If D = 0 Then
620         MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT2"

            Exit Sub

        End If

630     Gdb.Execute "UPDATE DOKEGGT1 SET S1X=0 WHERE S1X IS NULL"

640     Gdb.Execute "UPDATE DOKEGGT1 SET S1P=0 WHERE S1P IS NULL"

650     Gdb.Execute "UPDATE DOKEGGT2 SET S2X=0 WHERE S2X IS NULL"

660     Gdb.Execute "UPDATE DOKEGGT2 SET S2P=0 WHERE S2P IS NULL"

        'DOKEGGT12 VIEW =  ΕΝΩΝΕΙ ΤΑ DOKEGGT1 KAI DOKEGGT2
        ' "SELECT DOKEGGT1.*,S2X,S2P FROM DOKEGGT1 LEFT JOIN DOKEGGT2 ON DOKEGGT1.KODE=DOKEGGT2.KODE "

670     D = 0

        On Error GoTo Command6_Click_Err

680     a = "SELECT KOD,EID.PROM,EID.ERG,EID.ONO,S1X,S1P,POS01,EXPR1,EXPR2,POS02 "
690     a = a + " INTO DOKEGGT3 FROM DOKEGGT12 INNER JOIN EID ON DOKEGGT12.KODE=EID.KOD"

700     If Len(Fsql) > 1 Then
710         a = a + " WHERE " + Fsql
        End If

720     If ORDERCOMBO.Text = ORDERCOMBO.List(0) Then
730         a = a + " ORDER BY EID.ONO"
        End If

740     If ORDERCOMBO.Text = ORDERCOMBO.List(1) Then
750         a = a + " ORDER BY KOD"
        End If

760     If ORDERCOMBO.Text = ORDERCOMBO.List(3) Then
770         a = a + " ORDER BY EID.ERG"
        End If

        On Error GoTo Command6_Click_Err

780     Gdb.Execute a, D

790     If D = 0 Then
800         MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT3"

            Exit Sub

        End If

        ' Dim X As String
810     X = "SELECT * FROM DOKEGGT3"

820     If Check1 Then
            'typos2
830         print3_xar UCase(X), "000000000", "", 0
        End If

        'If Check2 Then
        '  Print4_Excel UCase(x), "000000000"
        'End If

        Exit Sub

840     CrystalReport1.Connect = gConnect

850     CrystalReport1.Action = 1
        'Dim DUM

        'DUM = Shell("bcp emp.dbo.DOKEGGT3 out region.txt -S PLATO -T -c  -C RAW", vbMinimizedFocus)
        On Error Resume Next

860     FileCopy "C:\LAGEURO\DOKEGGT3SOS.DBF", "C:\LAGEURO\DOKEGGT3.DBF"

870     update_DBF_from_SQL "SELECT *FROM DOKEGGT3", "DOKEGGT3"

880     MsgBox "ΑΡΧΕΙΟ EXCEL:" + Chr(13) + "C:\LAGEURO\DOKEGGT3.DBF"

        '<EhFooter>
        Exit Sub

Command6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Command6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Command6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command7_Click()

        'GR2.PrintInfo.Draft = True
        '<EhHeader>
        On Error GoTo Command7_Click_Err

        '</EhHeader>

100     GR2.PrintInfo.NoClipping = True
110     GR2.Styles(6).BackColor = vbWhite

120     With GR2.PrintInfo

            ' Set the page header
130         .PageHeaderFont.Italic = True
140         .PageHeader = "Αποθήκη " + Format(Date$, "mm/dd/YYYY")

            'Column headers will be on every page
150         .RepeatColumnHeaders = True

            ' Display page numbers (centered)
160         .PageFooter = "\tσελίδα: \p"
            ' Invoke Print Preview
170         .PrintPreview
        End With

        '<EhFooter>
        Exit Sub

Command7_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Command7_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Command7_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command8_Click()

        'ΕΚΤΥΠΩΣΗ ΣΤΟ ΣΥΝΟΛΟ ΤΩΝ ΑΠΟΘΗΚΩΝ
        '<EhHeader>
        On Error GoTo Command8_Click_Err

        '</EhHeader>

        Dim a As String

        Dim B As String

        Dim D As Long

100     Me.MousePointer = vbHourglass




            Dim mnow As String
            mnow = "EID"
            If mnow = "EID" Then
            Else
                 mnow = "DOKEID" + Format(Now, "YYMMDDHHmm")
                 Gdb.Execute "DROP TABLE " + mnow
                 Gdb.Execute "SELECT * into " + mnow + "  FROM EID"
            End If
            



        ' UPDATE_YPOLOIPA
110     UPDATE_YPOLOIPA3 mnow, List11, apo, eos

120     Fsql = Text1.Text

        Dim ANS
            
130     If F_epan = 1 Then

140         ANS = MsgBox("Να υπολογίσω το κέρδος ξανά;", vbYesNo)

150         If ANS = vbYes Then
160             ypol_kerd
            End If
        End If

        Dim polhs As String

        Dim AGOEPIS As String

        Dim polepis As String, pol As String, ago As String, APIS

170     a = Get_AJ(polhs, polepis, ago, AGOEPIS)

        'polhs = "'L','t','T','y','Y','υ'"

        Dim apothkes As String

180     If Val(Text3.Text) = 99 Then
190         apothkes = ""
            'Gdb.Execute "UPDATE DOKEGGT1 SET  S1P=S1P+S2P+S3P+S4P,S1X=S1X+S2X+S3X+S4X"
        Else
200         apothkes = " and APOT=" + Text3.Text
        End If

        '
        ' On Error Resume Next    AS1P>0

        On Error GoTo Command8_Click_Err

        '-------------------------------- NEO ARXH ===============================

210     Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1

220     a = "SELECT KODE," & " SUM( CASE   WHEN LEFT(ATIM,1)='λ'  THEN XRE ELSE 0  END ) AS S0X," & " SUM( CASE   WHEN LEFT(ATIM,1) IN ('G','g')  THEN POSO ELSE 0  END ) AS S1X," & " SUM( CASE WHEN PIS IS NULL  OR LEFT(ATIM,1) NOT IN (" + polhs + ")  THEN 0 ELSE POSO END) AS S1P," & " SUM( CASE WHEN XRE*TIMM IS NULL THEN 0 ELSE XRE*TIMM*(1- (CASE WHEN EKPT IS NULL THEN 0 ELSE EKPT END)/100) END  ) AS AS1X," & " SUM( CASE WHEN POSO*TIMM IS NULL OR LEFT(ATIM,1) NOT IN (" + polhs + ")  THEN 0 ELSE POSO*TIMM*(1-(CASE WHEN EKPT IS NULL THEN 0 ELSE EKPT END)/100) END ) AS AS1P, " & " SUM( CASE WHEN KERDOS IS NULL OR LEFT(ATIM,1) NOT IN (" + polhs + ")  THEN 0 ELSE KERDOS END ) AS KERD1 INTO DOKEGGT1 "
230     a = a + " FROM EGGTIM WHERE  left(ATIM,1) NOT IN ('σ','Σ') AND  " + " HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' " + apothkes + "   GROUP BY KODE"

        ' ' & " SUM(XRE)-SUM(PIS) AS YP," _

            Gdb.Execute a, D

240     If D = 0 Then
250         MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"
260         Me.MousePointer = vbNormal

            Exit Sub

        End If

        '
270     a = "SELECT LEFT(KOD,14)AS [ΚΩΔ],LEFT(EID.PROM,10) AS [ΠΡΟΜΗΘ],LEFT(EID.ONO,35) AS [ΠΕΡΙΓΡΑΦΗ],"
280     a = a + "S0X AS [ΑΠΟΓΡ], S1X AS [ΑΓΟΡΕΣ],S1P AS [ΠΩΛΗΣΕΙΣ],S0X+S1X-S1P AS [ΥΠΟΛ],"

290     a = a + "(CASE WHEN S1P>0 THEN AS1P/S1P ELSE 0  END ) AS [MEΣΤΙΜ.ΠΩΛ],XTI AS [TIM.AΓ],"

300     If F_epan = 1 Then
310         a = a + "(CASE WHEN AS1P*S1P*XTI>0 THEN (AS1P/(XTI*S1P)-1)*100 ELSE 0  END ) AS [M.UP],(CASE WHEN AS1P*S1P*XTI>0 THEN (1-(XTI*S1P)/AS1P)*100 ELSE 0  END ) AS [MG],AS1X AS [ΑΞ.ΑΓΟΡ],AS1P AS [ΑΞ.ΠΩΛ],(case WHEN AS1P>0 THEN AS1P-S1P*EID.XTI ELSE 0 END) AS [KEΡΔΟΣ],KERD1 "
        Else
320         a = a + "(CASE WHEN AS1P*S1P*XTI>0 THEN (AS1P/(XTI*S1P)-1)*100 ELSE 0  END ) AS [M.UP],(CASE WHEN AS1P*S1P*XTI>0 THEN (1-(XTI*S1P)/AS1P)*100 ELSE 0  END ) AS [MG],AS1X AS [ΑΞ.ΑΓΟΡ],AS1P AS [ΑΞ.ΠΩΛ],(case WHEN AS1P>0 THEN AS1P-S1P*EID.XTI ELSE 0 END) AS [KEΡΔΟΣ] "
        End If

330     a = a + " INTO DOKEGGT3 FROM " + mnow + " EID  LEFT JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD  "
        '  ,(AS2P+AS1P)-(S2P+S1P)*EID.XTI AS [ΚΕΡΔΟΣ]

340     If Len(Fsql) > 1 Then
350         a = a + " where " + Fsql
        End If

360     If ORDERCOMBO.Text = ORDERCOMBO.List(0) Then
370         a = a + " ORDER BY EID.ONO"
        End If

380     If ORDERCOMBO.Text = ORDERCOMBO.List(1) Then
390         a = a + " ORDER BY KOD"
        End If

400     If ORDERCOMBO.Text = ORDERCOMBO.List(2) Then
410         a = a + " ORDER BY EID.ERG"
        End If

420     If ORDERCOMBO.Text = ORDERCOMBO.List(3) Then
430         a = a + " ORDER BY AS1P-S1P*EID.XTI "  ' [KEPΔΟΣ]"
        End If

        ' Exit Sub

        On Error Resume Next

440     Gdb.Execute "DROP TABLE DOKEGGT3"  'SOUMES ΑΠΟΘΗΚΗΣ 1

        On Error GoTo LATOS

450     Gdb.Execute a, D
460     Me.MousePointer = vbNormal

470     If D = 0 Then
480         MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT3"
            GoTo mexit
            Exit Sub

        End If

        Dim X As String

        Dim ANS2

490     ANS2 = MsgBox("ΜΟΝΟ ΑΥΤΑ ΠΟΥ ΚΙΝΗΘΗΚΑΝ", vbYesNo)

500     If ANS2 = vbYes Then
510         X = "SELECT * FROM DOKEGGT3 WHERE [ΑΓΟΡΕΣ]>0 OR [ΠΩΛΗΣΕΙΣ]>0 OR [ΑΞ.ΠΩΛ]>0 "
        Else
520         X = "SELECT * FROM DOKEGGT3"
        End If

530     Data1.ConnectionString = gConnect
540     Data1.RecordSource = X
550     Data1.Refresh













560     If Check1 Then

            'typos2
            On Error GoTo Command8_Click_Err

570         print3_xar UCase(X), "000011111221111111111", "", 0
        End If

        'If Check2 Then
        '  Print4_Excel UCase(x), "00001111111111111111"
        'End If

mexit:


On Error Resume Next
If Left$(mnow, 3) = "DOK" Then
   Gdb.Execute "drop table " + mnow
End If





        Exit Sub

LATOS:
580     MsgBox Err.Description

590     Resume Next

        '<EhFooter>
        Exit Sub

Command8_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Command8_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Command8_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub


Sub sumes_grid(Sthles As Integer)


'         Do While Not ado.Recordset.EOF
'
'            For k = 0 To ado.Recordset.FIELDS.Count - 1
'
'                 If IsNumeric(ado.Recordset.FIELDS(k).Value) And ado.Recordset.FIELDS(k).Type <> 202 Then
'                     sumes(k) = sumes(k) + Adodc2.Recordset.FIELDS(k).Value
'                 End If
'
'             Next
'
'             Adodc2.Recordset.MoveNext
'
'                '  Exit Do
'            Loop
Dim k As Long

         For k = 0 To Sthles

             If sumes(k) > 0 Then
                 GR2.Splits(0).columns(k).FooterText = Format(sumes(k), "######0.00")
             End If

         Next

     GR2.AlternatingRowStyle = True
     GR2.OddRowStyle.BackColor = &H8000000F   ' GRI   vbCyan
     GR2.EvenRowStyle.BackColor = &HFFFFC0   'OYRANI     &H8000000F  ' GRI

End Sub





















Function Get3_AJ(ByRef pol As String, _
                 polepis As String, _
                 ago As String, _
                 AGOEPIS As String, _
                 ByRef PolXParoxh As String, _
                 Paroxh As String, _
                 ByRef PolAJ As String, _
                 AgoAJ As String) As Boolean

        '<EhHeader>
        On Error GoTo Get3_AJ_Err

        '</EhHeader>

        Dim DB As Database

        Dim R  As New ADODB.Recordset, X As String

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

100     R.Open "select POL,EIDOS,AJIA_APOU,SHM1 from PARASTAT", Gdb, adOpenDynamic, adLockOptimistic

110     pol = ""
120     PolXParoxh = ""

130     PolAJ = ""
140     AgoAJ = ""
150     Paroxh = ""

160     R.MoveFirst

170     Do While Not R.EOF

180         If Not IsNull(R("eidos") + R("pol") + R("ajia_apou")) Then
190             If R("pol") = "1" And R("ajia_apou") = "3" Then
200                 pol = pol + "'" + R("eidos") + "',"
                End If

210             If R("ajia_apou") = "3" Then
220                 PolAJ = PolAJ + "'" + R("eidos") + "',"
                End If
            
230             If R("pol") = "1" And R("ajia_apou") = "3" Then
240                 If Not IsNull(R("shm1")) And Val(Left(R("shm1"), 2)) <> 2 Then
250                     PolXParoxh = PolXParoxh + "'" + R("eidos") + "',"
                    End If
                End If
            
260             If R("pol") = "1" And R("ajia_apou") = "3" Then
270                 If Not IsNull(R("shm1")) And Val(Left(R("shm1"), 2)) = 2 Then
280                     Paroxh = Paroxh + "'" + R("eidos") + "',"
                    End If
                End If

290             If R("pol") = "1" And R("ajia_apou") = "4" Then
300                 polepis = polepis + "'" + R("eidos") + "',"
                End If

310             If R("pol") = "2" And R("ajia_apou") = "1" Then
320                 ago = ago + "'" + R("eidos") + "',"
                End If
            
330             If R("ajia_apou") = "1" Then
340                 AgoAJ = AgoAJ + "'" + R("eidos") + "',"
                End If
            
350             If R("pol") = "2" And R("ajia_apou") = "2" Then
360                 AGOEPIS = AGOEPIS + "'" + R("eidos") + "',"
                End If
            End If

370         R.MoveNext
        Loop

380     pol = Left(pol, Len(pol) - 1)
    
390     PolAJ = Left(PolAJ, Len(PolAJ) - 1)
400     AgoAJ = Left(AgoAJ, Len(AgoAJ) - 1)
    
410     PolXParoxh = Left(PolXParoxh, Len(PolXParoxh) - 1)

        On Error Resume Next

420     Paroxh = Left(Paroxh, Len(Paroxh) - 1)

430     If Len(polepis) > 0 Then
440         polepis = Left(polepis, Len(polepis) - 1)
        Else
450         polepis = ""
        End If

460     AGOEPIS = Left(AGOEPIS, Len(AGOEPIS) - 1)
470     ago = Left(ago, Len(ago) - 1)
480     Get3_AJ = True

        '<EhFooter>
        Exit Function

Get3_AJ_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Get3_AJ " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Get3_AJ " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

'Private Sub ypol_kerd_FAST()
'ypologismos kerdoys
'************************ ???d?? ************************************
'DROP TABLE DOKTIMHAG;
'SELECT EID.KOD,(SELECT TOP 1 TIMM*(100-EKPT)/100 FROM EGGTIM WHERE LEFT(ATIM,1) IN('G','Ξ') AND KODE=EID.KOD ORDER BY HME DESC) AS TIMHAG  INTO DOKTIMHAG  FROM EID
'End Sub

Private Sub ypol_kerd()

        'ypologismos kerdoys
        '************************ ???d?? ************************************
        'DROP TABLE DOKTIMHAG;
        'SELECT EID.KOD,(SELECT TOP 1 TIMM*(100-EKPT)/100 FROM EGGTIM WHERE LEFT(ATIM,1) IN('G','Ξ') AND KODE=EID.KOD ORDER BY HME DESC) AS TIMHAG  INTO DOKTIMHAG  FROM EID
        '<EhHeader>
        On Error GoTo ypol_kerd_Err

        '</EhHeader>

        Dim arx_arit, tel_arit, buff, kod_tim, ono_par, p1, progr, k

        Dim m_rec, m_synt, synt1, synt2

        '* ?????? ??a ??? a??e??  sumaeid.dbf µe fields kod (??d???? e?d???) pos (p?s?t?ta)
        '*          ?a? µe index sumaeid.ntx ?at? ??d???

        'do while .t.   &&  -------------  loop p?????µµat??
        '  *a????? t? a??e?? eggtim µe index a???µ? t?µ???????
        '  DO OPEN_DATA WITH .f.,3,'EID','TIM','EGGTIM'

        ' sele 9
        Dim SUMAEID As New ADODB.Recordset

        'Gdb.Execute "DELETE FROM SUMAEID"

        'SUMAEID.Open "SELECT * FROM SUMAEID", Gdb, adOpenDynamic, adLockOptimistic

100     arx_arit = 0: tel_arit = 0

        ' sele TIM
        'SET index to ARTHME

        Dim polhs As String

        Dim AgoAJ As String

        Dim PolAJ As String

        Dim PolXParoxh As String

        Dim Paroxh As String

        Dim AGOEPIS As String

        Dim polepis As String, pol As String, ago As String, APIS

        Dim a

110     a = Get3_AJ(polhs, polepis, ago, AGOEPIS, PolXParoxh, Paroxh, PolAJ, AgoAJ)

120     Gdb.Execute "UPDATE TIM SET KERDOS=0 where HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'"

130     Gdb.Execute "UPDATE EGGTIM SET KERDOS=0 where HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'"

        Dim EGGTIM As New ADODB.Recordset

140     EGGTIM.Open "SELECT * FROM EGGTIM WHERE LEFT(ATIM,1) IN (" + PolAJ + ") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' ORDER BY HME", Gdb, adOpenDynamic, adLockOptimistic

        Dim R As New ADODB.Recordset

        Dim EID As New ADODB.Recordset

150     AGNOSTOI.AddItem "ΔΕΝ ΒΡΕΘΗΚΑΝ ΟΙ ΚΩΔΙΚΟΙ:"

160     EGGTIM.MoveFirst

        Dim MATIM, mHME, MM

        Dim B, D

170     Do While Not EGGTIM.EOF

            'ΒΡΙΣΚΩ ΤΗΝ ΤΕΛΕΥΤΑΙΑ ΤΙΜΗ ΑΓΟΡΑΣ
180         a = EGGTIM("KODE")

            ' If Left(a, 10) = "158-60-ΑΠΨ" Then
            '    a = a
            ' End If

190         D = EGGTIM("HME")
200         k = 0
210         R.Open "SELECT top 1 * FROM EGGTIM WHERE KODE='" + a + "' AND LEFT(ATIM,1) IN ('G','Ξ','g','λ') AND HME<='" + Format(D, "MM/DD/YYYY") + "' ORDER BY HME DESC", Gdb, adOpenDynamic, adLockOptimistic

220         If R.EOF Then
                'PAIRNO THN TIMH APO TO EID
230             EID.Open "SELECT XTI FROM EID WHERE KOD='" + a + "'", Gdb, adOpenForwardOnly, adLockReadOnly

240             If EID.EOF Then
250                 k = MsgBox("ΔΕΝ ΥΠΑΡΧΕΙ ΣΤΑ ΕΙΔΗ Ο ΚΩΔΙΚΟΣ " + a + Chr(13) + "Συνέχεια Ν/Ο", vbYesNo)

260                 If k = vbNo Then

                        Exit Sub

                    End If

                Else
270                 B = EID(0)
                End If

280             EID.Close
            Else

290             If IsNull(R("EKPT")) Then
300                 B = R("TIMM")
                Else
310                 B = R("TIMM") * (1 - R("EKPT") / 100)
                End If

320             Gdb.Execute "UPDATE EID SET HPAR='" + Format(R("HME"), "MM/DD/YYYY") + "' where KOD='" + a + "'", k

            End If

330         If IsNull(B) Or B = 0 Then
                'PAIRNO THN TIMH APO TO EID
340             EID.Open "SELECT XTI FROM EID WHERE KOD='" + a + "'", Gdb, adOpenForwardOnly, adLockReadOnly

350             If EID.EOF Then
360                 AGNOSTOI.Visible = True
370                 AGNOSTOI.Refresh
380                 AGNOSTOI.AddItem a
                Else
390                 B = EID(0)

                    ' εαν δεν βρηκε τιμή αγοράς βάλε το 80% της τιμής πώλησης
400                 If B = 0 Then
410                     B = 0.8 * EGGTIM("timm") * (1 - EGGTIM("EKPT") / 100)
                    End If

                End If

420             EID.Close
            End If

            '
430         R.Close

440         DoEvents
450         Me.Caption = Format(D, "dd/mm/yyyy hh:mm") + " " + Format(k, "####")

460         If IsNull(EGGTIM("EKPT")) Then
470             EGGTIM("KERDOS") = (EGGTIM("TIMM") - B) * EGGTIM("POSO")
            Else
                ' TIMH POLHSHS                            - AGORAS X POSOTHTA
480             EGGTIM("KERDOS") = (EGGTIM("TIMM") * (1 - EGGTIM("EKPT") / 100) - B) * EGGTIM("POSO")
            End If

490         MATIM = EGGTIM("atim")
500         mHME = Format(EGGTIM("hme"), "MM/DD/YYYY")
510         MM = EGGTIM("kerdos")
520         EGGTIM.Update
530         EGGTIM.MoveNext

540         If IsNull(MM) Then MM = 0
550         Gdb.Execute "UPDATE TIM SET KERDOS=KERDOS + " + str(MM) + " WHERE HME='" + mHME + "' AND ATIM='" + MATIM + "'"
        Loop

560     EGGTIM.Close

        '<EhFooter>
        Exit Sub

ypol_kerd_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.ypol_kerd " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.ypol_kerd " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Public Sub TelHmeAgoras()
'
'        '===================================================================================
'        '<EhHeader>
'        On Error GoTo TelHmeAgoras_Err
'
'        '</EhHeader>
'        Dim arx_arit, tel_arit, BUFF, kod_tim, ono_par, p1, progr, K
'
'        Dim m_rec, m_synt, synt1, synt2
'
'        Dim SUMAEID As New ADODB.Recordset
'
'100     arx_arit = 0: tel_arit = 0
'
'        ' sele TIM
'        'SET index to ARTHME
'
'        Dim polhs As String
'
'        Dim PolXParoxh As String
'
'        Dim Paroxh As String
'
'        Dim AGOEPIS As String
'
'        Dim polepis As String, pol As String, ago As String, APIS
'
'        Dim a
'
'110     a = Get2_AJ(polhs, polepis, ago, AGOEPIS, PolXParoxh, Paroxh)
'
'        Dim EGGTIM As New ADODB.Recordset
'
'        ''E','α',
'120     EGGTIM.Open "SELECT * FROM EGGTIM WHERE LEFT(ATIM,1) IN ('E','α','λ'," + ago + ") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' ORDER BY HME", Gdb, adOpenDynamic, adLockOptimistic
'
'        Dim R As New ADODB.Recordset
'
'        Dim EID As New ADODB.Recordset
'
'130     If EGGTIM.EOF Then
'
'            Exit Sub
'
'        End If
'
'140     EGGTIM.MoveFirst
'
'        Dim b, D, ATIM
'
'150     K = 0
'
'        Dim N
'
'160     Do While Not EGGTIM.EOF
'            'ΒΡΙΣΚΩ ΤΗΝ ΤΕΛΕΥΤΑΙΑ ΤΙΜΗ ΑΓΟΡΑΣ
'170         a = EGGTIM("KODE")
'180         D = EGGTIM("HME")
'190         ATIM = EGGTIM("atim")
'200         K = K + 1
'
'210         If Not IsNull(EGGTIM("kode")) Then
'220             If IsNull(EGGTIM("TIMM")) Then
'                Else
'
'230                 If IsNull(EGGTIM("EKPT")) Then
'                    Else
'
'240                     If EGGTIM("TIMM") > 0 Then
'250                         Gdb.Execute "UPDATE EID SET XTI=" + str(nNull(EGGTIM("TIMM") * (1 - EGGTIM("EKPT") / 100))) + ",CH1='" + ATIM + "',HPAR='" + Format(D, "MM/DD/YYYY") + "',PAR=" + str(EGGTIM("POSO")) + "  where KOD='" + a + "'", N
'                        Else
'260                         Gdb.Execute "UPDATE EID SET CH1='" + ATIM + "',HPAR='" + Format(D, "MM/DD/YYYY") + "',PAR=" + str(EGGTIM("POSO")) + "  where KOD='" + a + "'", N
'                        End If
'                    End If
'
'                End If
'            End If
'
'270         If N > 0 Then
'280             MDIForm1.Caption = Format(K, "####") + " - " + Format(D, "dd/mm/yyyy hh:mm") + " "
'            End If
'
'290         If (K Mod 10 = 0) Then DoEvents
'
'300         DoEvents
'
'310         EGGTIM.MoveNext
'        Loop
'
'320     EGGTIM.Close
'
'        '<EhFooter>
'        Exit Sub
'
'TelHmeAgoras_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.apot3.TelHmeAgoras " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.TelHmeAgoras " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>
'
'End Sub

Private Sub TelProm()

        '===================================================================================
        '<EhHeader>
        On Error GoTo TelProm_Err

        '</EhHeader>
        Dim arx_arit, tel_arit, buff, kod_tim, ono_par, p1, progr, k

        Dim m_rec, m_synt, synt1, synt2

        Dim SUMAEID As New ADODB.Recordset

100     arx_arit = 0: tel_arit = 0

        ' sele TIM
        'SET index to ARTHME

        Dim polhs As String

        Dim PolXParoxh As String

        Dim Paroxh As String

        Dim AGOEPIS As String

        Dim polepis As String, pol As String, ago As String, APIS

        Dim a

110     a = Get2_AJ(polhs, polepis, ago, AGOEPIS, PolXParoxh, Paroxh)

        Dim EGGTIM As New ADODB.Recordset

        ' EGGTIM.Open "SELECT * FROM EGGTIM WHERE LEFT(ATIM,1) IN ("" + ago + "") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' ORDER BY HME", Gdb, adOpenDynamic, adLockOptimistic
120     EGGTIM.Open " select PEL.EPO,EGGTIM.* FROM EGGTIM INNER JOIN PEL ON EGGTIM.EIDOS=PEL.EIDOS AND PEL.KOD=PELKOD         WHERE LEFT(ATIM,1) IN ('G','g') order by HME", FGDB, adOpenDynamic, adLockOptimistic

        Dim R As New ADODB.Recordset

        Dim EID As New ADODB.Recordset

130     If EGGTIM.EOF Then

            Exit Sub

        End If

140     EGGTIM.MoveFirst

        Dim B, D, ATIM

150     k = 0

        Dim N

160     Do While Not EGGTIM.EOF
            'ΒΡΙΣΚΩ ΤΗΝ ΤΕΛΕΥΤΑΙΑ ΤΙΜΗ ΑΓΟΡΑΣ
170         a = EGGTIM("KODE")
180         D = EGGTIM("HME")
190         ATIM = EGGTIM("atim")
200         k = k + 1

210         If Not IsNull(EGGTIM("kode")) Then
220             Gdb.Execute "UPDATE EID SET PROM='" + Left(EGGTIM("EPO"), 30) + "'  where KOD='" + a + "'", N
            End If

230         If N > 0 Then
240             MDIForm1.Caption = Format(k, "####") + " - " + Format(D, "dd/mm/yyyy hh:mm") + " "
            End If

250         If (k Mod 10 = 0) Then DoEvents

260         DoEvents

270         EGGTIM.MoveNext
        Loop

280     EGGTIM.Close

        '<EhFooter>
        Exit Sub

TelProm_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.TelProm " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.TelProm " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command9_Click()

        'ektypvsh kerdvn
        'ΕΚΤΥΠΩΣΗ ΣΤΟ ΣΥΝΟΛΟ ΤΩΝ ΑΠΟΘΗΚΩΝ
        '<EhHeader>
        On Error GoTo Command9_Click_Err

        '</EhHeader>

        Dim a As String

        Dim B As String

        Dim D As Long

        ' UPDATE_YPOLOIPA
100     UPDATE_YPOLOIPA3 "EID", List11, apo, eos

110     Fsql = Text1.Text

        Dim ANS

120     If F_epan = 1 Then

130         ANS = MsgBox("Να υπολογίσω το κέρδος ξανά;", vbYesNo)

140         If ANS = vbYes Then
150             ypol_kerd
            End If
        End If

        Dim polhs As String

        Dim PolXParoxh As String

        Dim Paroxh As String

        Dim AGOEPIS As String

        Dim polepis As String, pol As String, ago As String, APIS

160     a = Get2_AJ(polhs, polepis, ago, AGOEPIS, PolXParoxh, Paroxh)

170     If Combo1.ListIndex = 0 Then
180         polhs = PolXParoxh
        End If

190     If Combo1.ListIndex = 2 Then
200         polhs = Paroxh
        End If

        'polhs = "'L','t','T','y','Y','υ'"

        Dim apothkes As String

210     If Val(Text3.Text) = 99 Then
220         apothkes = ""
        Else
230         apothkes = " and APOT=" + Text3.Text
        End If

        '
        ' On Error Resume Next    AS1P>0

        On Error GoTo Command9_Click_Err

        '-------------------------------- NEO ARXH ===============================

240     Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1

250     a = "SELECT KODE," & " SUM( CASE   WHEN LEFT(ATIM,1)='λ'  THEN XRE ELSE 0  END ) AS S0X," & " SUM( CASE   WHEN LEFT(ATIM,1) IN ('G','g')  THEN XRE ELSE 0  END ) AS S1X," & " SUM( CASE WHEN PIS IS NULL  OR LEFT(ATIM,1) NOT IN (" + polhs + ")  THEN 0 ELSE PIS END) AS S1P," & " SUM( CASE WHEN XRE*TIMM IS NULL THEN 0 ELSE XRE*TIMM*(1- (CASE WHEN EKPT IS NULL THEN 0 ELSE EKPT END)/100) END  ) AS AS1X," & " SUM( CASE WHEN POSO*TIMM IS NULL OR LEFT(ATIM,1) NOT IN (" + polhs + ")  THEN 0 ELSE POSO*TIMM*(1-(CASE WHEN EKPT IS NULL THEN 0 ELSE EKPT END)/100) END ) AS AS1P, " & " SUM( CASE WHEN KERDOS IS NULL OR LEFT(ATIM,1) NOT IN (" + polhs + ")  THEN 0 ELSE KERDOS END ) AS KERD1 INTO DOKEGGT1 "
260     a = a + " FROM EGGTIM WHERE  left(ATIM,1) NOT IN ('σ','Σ') AND left(ATIM,1) IN (" + polhs + ") AND  " + " HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' " + apothkes + "   GROUP BY KODE"

        ' ' & " SUM(XRE)-SUM(PIS) AS YP," _

            Gdb.Execute a, D

270     If D = 0 Then
280         MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"

            Exit Sub

        End If

        '
290     a = "SELECT LEFT(KOD,14)AS [ΚΩΔ],LEFT(EID.PROM,10) AS [ΠΡΟΜΗΘ],LEFT(EID.ONO,35) AS [ΠΕΡΙΓΡΑΦΗ],"
300     a = a + "S0X AS [ΑΠΟΓΡ], S1X AS [ΑΓΟΡΕΣ],S1P AS [ΠΩΛΗΣΕΙΣ],S0X+S1X-S1P AS [ΥΠΟΛ],"

310     a = a + "(CASE WHEN S1P>0 THEN AS1P/S1P ELSE 0  END ) AS [MEΣΤΙΜ.ΠΩΛ],XTI AS [TIM.AΓ],"

320     If F_epan = 1 Then
330         a = a + "(CASE WHEN AS1P*S1P*XTI>0 THEN (AS1P/(XTI*S1P)-1)*100 ELSE 0  END ) AS [M.UP],(CASE WHEN AS1P*S1P*XTI>0 THEN (1-(XTI*S1P)/AS1P)*100 ELSE 0  END ) AS [MG],AS1X AS [ΑΞ.ΑΓΟΡ],AS1P AS [ΑΞ.ΠΩΛ],(case WHEN AS1P>0 THEN AS1P-S1P*EID.XTI ELSE 0 END) AS [KEΡΔΟΣ],KERD1 "
        Else
340         a = a + "(CASE WHEN AS1P*S1P*XTI>0 THEN (AS1P/(XTI*S1P)-1)*100 ELSE 0  END ) AS [M.UP],(CASE WHEN AS1P*S1P*XTI>0 THEN (1-(XTI*S1P)/AS1P)*100 ELSE 0  END ) AS [MG],AS1X AS [ΑΞ.ΑΓΟΡ],AS1P AS [ΑΞ.ΠΩΛ],(case WHEN AS1P>0 THEN AS1P-S1P*EID.XTI ELSE 0 END) AS [KEΡΔΟΣ] "
        End If

350     a = a + " INTO DOKEGGT3 FROM EID  LEFT JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD  "
        '  ,(AS2P+AS1P)-(S2P+S1P)*EID.XTI AS [ΚΕΡΔΟΣ]

360     If Len(Fsql) > 1 Then
370         a = a + " where " + Fsql
        End If

380     If ORDERCOMBO.Text = ORDERCOMBO.List(0) Then
390         a = a + " ORDER BY EID.ONO"
        End If

400     If ORDERCOMBO.Text = ORDERCOMBO.List(1) Then
410         a = a + " ORDER BY KOD"
        End If

420     If ORDERCOMBO.Text = ORDERCOMBO.List(2) Then
430         a = a + " ORDER BY EID.ERG"
        End If

440     If ORDERCOMBO.Text = ORDERCOMBO.List(3) Then
450         a = a + " ORDER BY AS1P-S1P*EID.XTI "  ' [KEPΔΟΣ]"
        End If

        ' Exit Sub

        On Error Resume Next

460     Gdb.Execute "DROP TABLE DOKEGGT3"  'SOUMES ΑΠΟΘΗΚΗΣ 1

        On Error GoTo LATOS

470     Gdb.Execute a, D

480     If D = 0 Then
490         MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT3"

            Exit Sub

        End If

        Dim X As String

        Dim ANS2

500     ANS2 = MsgBox("ΜΟΝΟ ΑΥΤΑ ΠΟΥ ΚΙΝΗΘΗΚΑΝ", vbYesNo)

510     If ANS2 = vbYes Then
520         X = "SELECT * FROM DOKEGGT3 WHERE [ΑΓΟΡΕΣ]>0 OR [ΠΩΛΗΣΕΙΣ]>0 OR [ΑΞ.ΠΩΛ]>0 "
        Else
530         X = "SELECT * FROM DOKEGGT3"
        End If

540     Data1.ConnectionString = gConnect
550     Data1.RecordSource = X
560     Data1.Refresh

570     If Check1 Then

            'typos2
            On Error GoTo Command9_Click_Err

580         print3_xar UCase(X), "000011111111111111111", "", 0
        End If

        'If Check2 Then
        '  Print4_Excel UCase(x), "00001111111111111111"
        'End If

        Exit Sub

LATOS:
590     MsgBox Err.Description

600     Resume Next

        '<EhFooter>
        Exit Sub

Command9_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Command9_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Command9_Click " & "at line " & Erl

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
        Dim N As Integer
        N = 0
130     Open File1.Path + "\" + F For Input As #1

140     Do While Not EOF(1)
            N = N + 1
150         Line Input #1, B: F_TXT(N) = B
160         SS = SS + B + " "
        Loop

170     Close #1

180     Text4.Text = SS


'130     Open File1.Path + "\" + f For Input As #1
'
'140     Do While Not EOF(1)
'
'150         Line Input #1, b
'155         n = n + 1: txt(n) = b
'160         ss = ss + b + Chr(13)
'        Loop
'
'170     Close #1
'        'label3=par1
'        'label4=par2
'        'label5=par3
       






190     rCheck.Value = vbChecked

        '<EhFooter>
        Exit Sub

File1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.File1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.File1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If g_Stop = 1 Then  'entos loop
110         g_Stop = 2  'ektos loop

            Exit Sub

        End If

120     If KeyCode = 27 Then Command2_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'**
' basic form-load
'@rem          remarks
'**
'
'@see
'@rem
Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err
' F_PROTHFORA = 1
        '</EhHeader>
'If F_PROTHFORA = 1 Then
'Else
100     mForm_Load Me, fh, fw, ft, fl

'End If



        fupd_eid = FINDPARAMETROI(1, "APOT3", "fupd_eid", "1", "=1 σημαινει υπολογιζει τα υπολ. των ειδών και τα αποθηκευει")


        F_PLATOS_FORMAS = FINDPARAMETROI(1, "APOT3", "F_PLATOS_FORMAS", "15625", "Πλάτος φόρμας")

        'ειανι η καλυτερη διόρθωση με το δεξιό άκρο του δεξιότερου control (εξοδος =command2)
        If F_PLATOS_FORMAS > 0 Then F_PLATOS_FORMAS = Command2.Left + Command2.Width + 100

        On Error Resume Next

110     Combo1.Text = Combo1.List(0)
120     Combo2.Text = Combo2.List(1)
130     F_RUN = 0
140     Fsql = ""
150     FIsHidden = False
160     Me.KeyPreview = True

170     F_UPDYPOL = Val(FINDPARAMETROI(1, "APOT3", "F_UPDYPOL", "1", "1=ΝΑ ΑΠΟΘΗΚΕΥΟΝΤΑΙ ΤΑ ΥΠΟΛΟΙΠΑ ΣΤΟ MASTER ΕΙΔΩΝ ΟΧΙ=0"))

180     F_epan = Val(FINDPARAMETROI(1, "APOT3", "F_EPAN", "1", "EΠΑΝΥΠΟΛΟΓΙΣΜΟΣ ΠΑΡ/ΚΩΝ=1 ΟΧΙ=0"))

        '   'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        '    'SkinFramework1.ApplyWindow apot3.hWnd
        '    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

190     Me.Picture = LoadPicture(gPicture)

200     cd1.InitDir = "c:\mercvb\queries\EID"
        'On Error Resume Next
        ' File1.Path = LCase("c:\mercvb\queries\EID")

210     If Len(UCase(Dir("c:\mercvb\queries\eid", vbDirectory))) = 0 Then
220         MkDir "c:\mercvb\queries\eid"
        End If

230     File1.Path = "c:\mercvb\queries\eid"

240     File1.Refresh

250     F_LHM1 = FINDPARAMETROI(1, "APOT2", "F_lhm1", "Ημερ1", "Ετικέτα ημερ.1")
260     F_LHM2 = FINDPARAMETROI(1, "APOT2", "F_lhm2", "Ημερ2", "Ετικέτα ημερ.2")
270     f_lHM3 = FINDPARAMETROI(1, "APOT2", "F_lhm3", "Ημερ3", "Ετικέτα ημερ.3")

280     f_lab1 = FINDPARAMETROI(1, "APOT2", "F_lab1", "Κειμ1", "Ετικέτα κειμ.1")
290     f_lab2 = FINDPARAMETROI(1, "APOT2", "F_lab2", "Κειμ2", "Ετικέτα κειμ.2")
300     f_lab3 = FINDPARAMETROI(1, "APOT2", "F_lab3", "Κειμ3", "Ετικέτα κειμ.3")
310     f_lab4 = FINDPARAMETROI(1, "APOT2", "F_lab4", "Κειμ4", "Ετικέτα κειμ.4")
320     f_lab5 = FINDPARAMETROI(1, "APOT2", "F_lab5", "Κειμ5", "Ετικέτα κειμ.5")
330     f_lab6 = FINDPARAMETROI(1, "APOT2", "F_lab6", "Κειμ6", "Ετικέτα κειμ.6")

340     f_num1 = FINDPARAMETROI(1, "APOT2", "F_num1", "Αριθ1", "Ετικέτα κειμ.1")
350     f_num2 = FINDPARAMETROI(1, "APOT2", "F_num2", "Αριθ2", "Ετικέτα κειμ.2")
360     f_num3 = FINDPARAMETROI(1, "APOT2", "F_num3", "Αριθ3", "Ετικέτα κειμ.3")

        'ΣΤΗΛΕΣ ΠΙΝΑΚΑ F_LIST1
        ' (n,0)=sql string  (n,1)=sum  (n,2)=alignment  (n,3)=width  (n,4)=listbox TEXT

        Dim R As Integer

370     R = 0 + 0: f_list1(R, 0) = ",LEFT(MON,5) AS [Mον.Μετ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Μον.Μέτρ."
380     R = R + 1: f_list1(R, 0) = ",convert(decimal(10,2),LTI) as [Χονδ.Τιμή]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "Τιμή Xoνδρικής"

'  R = R + 1: f_list1(R, 0) = ",convert(decimal(10,2),LTI5) as [Χονδ.Τιμή]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "Τιμή Xoνδρικής"

390     R = R + 1: f_list1(R, 0) = ",convert(decimal(10,2),LTI5) as [Λιαν.Τιμή]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "Τιμή Λιανικής"
400     R = R + 1: f_list1(R, 0) = ",(XPOS01+XPOS02+XPOS03+XPOS04) AS [Υπόλοιπο]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Υπόλοιπο"

410     R = R + 1: f_list1(R, 0) = ",FPA as [Kατ.ΦΠΑ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Κατ.ΦΠΑ"
420     R = R + 1: f_list1(R, 0) = ",convert(decimal(10,2),XTI) AS [Tιμή Αγ.]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "Τιμή αγοράς"

430     R = R + 1: f_list1(R, 0) = ",XTI*(XPOS01+XPOS02+XPOS03+XPOS04) AS [Αξ.Απογρ.]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Αξία Απογραφής"

440     R = R + 1: f_list1(R, 0) = ",PONTOI AS [Πόντοι]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Πόντοι"

450     R = R + 1: f_list1(R, 0) = ",ERG as [Βοηθ.Κωδ.] ": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Βarcode"
460     R = R + 1: f_list1(R, 0) = ",PROM AS [Προμηθευτής]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Προμηθευτής"
470     R = R + 1: f_list1(R, 0) = ",KODLOG as [Κατηγορία]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Κατηγορία"

        R = R + 1: f_list1(R, 0) = ",AEG as [Aρ.Οικογ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Aρ.Οικογ]"
                
                
        Dim query_oik As String
        query_oik = "(SELECT ISNULL ( (SELECT TOP 1 PERIGRAFH  FROM PINAKES WHERE TYPOS=11 AND AYJON=ISNULL(EID.AEG,0)  ) , '---') ) as [Oικογένεια] "
                R = R + 1: f_list1(R, 0) = "," + query_oik + "": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Οικογένεια"
        
        '  R = R + 1: .TextMatrix(N, 0) = "Οικογένεια": .TextMatrix(N, 3) = "SELECT ISNULL ( (SELECT TOP 1 PERIGRAFH  FROM PINAKES WHERE TYPOS=11 AND AYJON=ISNULL(EID.AEG,0)  ) , '---') AS DD ": .TextMatrix(N, 4) = "n"
  
480     R = R + 1: f_list1(R, 0) = ",CONVERT(CHAR(8),HPAR,3)  as [Ημ.Τελ.Αγοράς]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Ημερ.Τελ.Αγοράς"
490     R = R + 1: f_list1(R, 0) = ",PAR as [Παραγγελ.Ποσ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Παραγγ.Ποσότητα"

500     R = R + 1: f_list1(R, 0) = ",SPA as [ΣτάθμηΠαραγγ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Στάθμη Παραγγελίας"
510     R = R + 1: f_list1(R, 0) = ",EPIUYP as [Επιθ.Υπόλ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Επιθυμητό Υπόλοιπο"

520     R = R + 1: f_list1(R, 0) = ",SYSKMAX as [Συσκευασία]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Συσκευασία"
530     R = R + 1: f_list1(R, 0) = ",SYSKMIN as [ΥποΣυσκευασία]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "YποΣυσκευασία"

'======================================================================================
532     R = R + 1: f_list1(R, 0) = ",(CASE APOG01 WHEN NULL THEN 0 ELSE APOG01 END) AS [Απογ.Απ1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ1.Απογραφή"
        R = R + 1: f_list1(R, 0) = ",(CASE KOSAPOG01 WHEN NULL THEN 0 ELSE KOSAPOG01 END) AS [Αξ.Περ.Απογ.Απ1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ1.Αξ.Περ.Απογραφής"

540     R = R + 1: f_list1(R, 0) = ",(CASE S1X WHEN NULL THEN 0 ELSE S1X-APOG01 END) AS [Εισ1.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ1.Εισαγ.Περιόδου"
        
550     R = R + 1: f_list1(R, 0) = ",(CASE AS1X WHEN NULL THEN 0 ELSE AS1X-KOSAPOG01 END)  AS [Αξ1.Αγ.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ1.Αξ.Εισαγ.Περιόδου"
560     R = R + 1: f_list1(R, 0) = ",(CASE S1P WHEN NULL THEN 0 ELSE S1P END)  AS [Εξ1.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ1.Εξ.Περιόδου"

570     R = R + 1: f_list1(R, 0) = ",(CASE AS1P WHEN NULL THEN 0 ELSE AS1P END)  AS [Αξ1.Εξ.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ1.Αξ.Εξ.Περιόδου"

580     R = R + 1: f_list1(R, 0) = ",XPOS01 as [Υπ.Απ1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1300: f_list1(R, 4) = "Υπολ.Απ1"
590     R = R + 1: f_list1(R, 0) = ",XTI*(XPOS01) AS [Αξ.Απογρ.απ.1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Αξία Απογραφής απ1"


'=======================================================================================
        R = R + 1: f_list1(R, 0) = ",(CASE APOG02 WHEN NULL THEN 0 ELSE APOG02 END) AS [Απογ.Απ2]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ2.Απογραφή"
        R = R + 1: f_list1(R, 0) = ",(CASE KOSAPOG02 WHEN NULL THEN 0 ELSE KOSAPOG02 END) AS [Αξ.Περ.Απογ.Απ2]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ2.Αξια Περ.Απογραφής"
        
        
600     R = R + 1: f_list1(R, 0) = ",(CASE S2X WHEN NULL THEN 0 ELSE S2X-APOG02 END)  AS [Εισ2.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ2.Εισαγ.Περιόδου"
610     R = R + 1: f_list1(R, 0) = ",(CASE AS2X WHEN NULL THEN 0 ELSE AS2X-KOSAPOG02 END) AS [Αξ2.Αγ.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ2.Αξ.Εισαγ.Περιόδου"
620     R = R + 1: f_list1(R, 0) = ",(CASE S2P WHEN NULL THEN 0 ELSE S2P END)  AS [Εξ2.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ2.Εξ.Περιόδου"
630     R = R + 1: f_list1(R, 0) = ",(CASE AS2P WHEN NULL THEN 0 ELSE AS2P END)  AS [Αξ2.Εξ.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ2.Αξ.Εξ.Περιόδου"
640     R = R + 1: f_list1(R, 0) = ",XPOS02 as [Υπ.Απ2]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "Υπολ.Απ.2"
650     R = R + 1: f_list1(R, 0) = ",XTI*(XPOS02) AS [Αξ.Απογρ.απ.2]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Αξία Απογραφής απ2"


'===========================================================================================
        R = R + 1: f_list1(R, 0) = ",(CASE APOG03 WHEN NULL THEN 0 ELSE APOG03 END) AS [Απογ.Απ1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ3.Απογραφή"
        R = R + 1: f_list1(R, 0) = ",(CASE KOSAPOG03 WHEN NULL THEN 0 ELSE KOSAPOG03 END) AS [Αξ.Περ.Απογ.Απ3]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ3.Αξια Περ.Απογραφής"
660     R = R + 1: f_list1(R, 0) = ",(CASE S3X WHEN NULL THEN 0 ELSE S3X-APOG03 END)  AS [Εισ3.Περιόδ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ3.Εισαγ.Περιόδου"
670     R = R + 1: f_list1(R, 0) = ",(CASE AS3X WHEN NULL THEN 0 ELSE AS3X-KOSAPOG03 END) AS [Αξ3.Αγ.Περιόδ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ3.Αξ.Εισαγ.Περιόδου"
680     R = R + 1: f_list1(R, 0) = ",(CASE S3P WHEN NULL THEN 0 ELSE S3P END)  AS [Εξ3.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ3.Εξ.Περιόδου"
690     R = R + 1: f_list1(R, 0) = ",(CASE AS3P WHEN NULL THEN 0 ELSE AS3P END)  AS [Αξ3.Αγ.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ3.Αξ.Εξ.Περιόδου"
700     R = R + 1: f_list1(R, 0) = ",XPOS03 as [Υπ.Απ3]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Υπολ.Απ3"
710     R = R + 1: f_list1(R, 0) = ",XTI*(XPOS03) AS [Αξ.Απογρ.απ.3]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Αξία Απογραφής απ3"
      

'===========================================================================================
        R = R + 1: f_list1(R, 0) = ",(CASE APOG04 WHEN NULL THEN 0 ELSE APOG04 END) AS [Απογ.Απ1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ4.Απογραφή"
        R = R + 1: f_list1(R, 0) = ",(CASE KOSAPOG04 WHEN NULL THEN 0 ELSE KOSAPOG04 END) AS [Αξ.Περ.Απογ.Απ2]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ4.Αξια Περ.Απογραφής"
        
720     R = R + 1: f_list1(R, 0) = ",(CASE S4X WHEN NULL THEN 0 ELSE S4X-APOG04 END)  AS [Εισ4.Περιόδ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ4.Εισαγ.Περιόδου"
730     R = R + 1: f_list1(R, 0) = ",(CASE AS4X WHEN NULL THEN 0 ELSE AS4X-KOSAPOG04 END) AS [Αξ4.Αγ.Περιόδ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ4.Αξ.Εισαγ.Περιόδου"
740     R = R + 1: f_list1(R, 0) = ",(CASE S4P WHEN NULL THEN 0 ELSE S4P END)  AS [Εξ4.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ4.Εξ.Περιόδου"
750     R = R + 1: f_list1(R, 0) = ",(CASE AS4P WHEN NULL THEN 0 ELSE AS4P END)  AS [Αξ4.Αγ.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ4.Αξ.Εξ.Περιόδου"
760     R = R + 1: f_list1(R, 0) = ",XPOS04 as [Υπ.Απ4]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Υπολ.Απ4"
770     R = R + 1: f_list1(R, 0) = ",XTI*(XPOS04) AS [Αξ.Απογρ.απ4]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Αξία Απογραφής απ4"

        'R = R + 1: f_list1(R, 0) = ",POS04 as [Υπ.Απ4]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Υπολ.Απ4"

780     R = R + 1: f_list1(R, 0) = ", CONVERT(CHAR(12),HM1,3) AS [" + F_LHM1 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = F_LHM1

790     R = R + 1: f_list1(R, 0) = ", CONVERT(CHAR(12),HM2,3) AS [" + F_LHM2 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = F_LHM2

800     R = R + 1: f_list1(R, 0) = ", CONVERT(CHAR(12),HM3,3) AS [" + f_lHM3 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lHM3

810     R = R + 1: f_list1(R, 0) = ", CH1 AS [" + f_lab1 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab1
820     R = R + 1: f_list1(R, 0) = ", CH2 AS [" + f_lab2 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab2
830     R = R + 1: f_list1(R, 0) = ", CH3 AS [" + f_lab3 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab3
840     R = R + 1: f_list1(R, 0) = ", CH4 AS [" + f_lab4 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab4

850     R = R + 1: f_list1(R, 0) = ", CH5 AS [" + f_lab5 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab5
860     R = R + 1: f_list1(R, 0) = ", CH6 AS [" + f_lab6 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab6

870     R = R + 1: f_list1(R, 0) = ", ONO2 AS [Περιγραφή 2η]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Περιγραφή 2η"

880     R = R + 1: f_list1(R, 0) = ", str(NUM1,10,2) AS [" + f_num1 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_num1
890     R = R + 1: f_list1(R, 0) = ", str(NUM2,10,2) AS [" + f_num2 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_num2
900     R = R + 1: f_list1(R, 0) = ", str(NUM3,10,2) AS [" + f_num3 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_num3

910     R = R + 1: f_list1(R, 0) = ",(CASE ES1P WHEN NULL THEN 0 ELSE ES1P END) AS [ΠΩΛ.ΑΠ1.ΕΠΙΛ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ1.ΠΩΛ.ΕΠΙΛ.Περιόδου"
920     R = R + 1: f_list1(R, 0) = ",(CASE ES2P WHEN NULL THEN 0 ELSE ES2P END) AS [ΠΩΛ.ΑΠ2.ΕΠΙΛ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ2.ΠΩΛ.ΕΠΙΛ.Περιόδου"
930     R = R + 1: f_list1(R, 0) = ",(CASE ES3P WHEN NULL THEN 0 ELSE ES3P END) AS [ΠΩΛ.ΑΠ3.ΕΠΙΛ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ3.ΠΩΛ.ΕΠΙΛ.Περιόδου"
940     R = R + 1: f_list1(R, 0) = ",(CASE ES4P WHEN NULL THEN 0 ELSE ES4P END) AS [ΠΩΛ.ΑΠ4.ΕΠΙΛ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ4.ΠΩΛ.ΕΠΙΛ.Περιόδου"
    
950     R = R + 1: f_list1(R, 0) = ",MESXTI AS [ΜΕΣΗ ΤΙΜΗ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "ΜΕΣΗ ΤΙΜΗ"

        Dim k As Integer

956     For k = 0 To UBound(f_list1)

960         If Len(f_list1(k, 4)) > 0 Then
970             List1.AddItem f_list1(k, 4) + Space(180) + Format(k, "00")
            End If

        Next

980     List1.ToolTipText = "Με click επιλέγετε το πεδία" + " Με σύρσιμο μπορείτε να αλλάξετε την σειρά"
        List1.Selected(3) = True
        Dim r3 As New ADODB.Recordset

        'GEMISMA MULTI SELECT LISTBOX
990     r3.Open "select POL,TITLOS,EIDOS,AJIA_APOU,POS_APOU from PARASTAT where (POL=1 OR POL=2) AND POS_APOU<>0 ORDER BY POL,TITLOS;", Gdb, adOpenDynamic, adLockOptimistic
1000    List11.Clear
1010    r3.MoveFirst

1020    k = 0

1030    Do While Not r3.EOF

1040        If Not IsNull(r3("TITLOS")) And Left(r3("TITLOS"), 1) <> "-" Then
1050            List11.AddItem r3("TITLOS") + Space(1) + r3("eidos")
1060            List11.Selected(k) = True
1070            k = k + 1
                'mEidos_Par(K) = R("eidos")
            End If

1080        r3.MoveNext
        Loop

1090    r3.Close

        'List1.AddItem "Μον.Μέτρ."
        '
        'List1.AddItem "Τιμή Xoνδρικής"
        'List1.AddItem "Τιμή Λιανικής"
        '
        'List1.AddItem "Υπόλοιπο"
        'List1.AddItem "Υπολ.Απ.1"
        'List1.AddItem "Υπολ.Απ.2"

        'List1.AddItem "Κατ.ΦΠΑ"
        '
        'List1.AddItem "Τιμή αγοράς"
        'List1.AddItem "Αξία Απογραφής απ1"
        'List1.AddItem "Αξία Απογραφής απ2"
        'List1.AddItem "Αξία Απογραφής"
        'List1.AddItem "Πόντοι"

        'List1.AddItem "Απ1.Εισαγ.Περιόδου"
        'List1.AddItem "Απ1.Αξ.Εισαγ.Περιόδου"

        'List1.AddItem "Απ2.Εισαγ.Περιόδου"
        'List1.AddItem "Απ2.Αξ.Εισαγ.Περιόδου"

        'List1.AddItem "Απ1.Εξ.Περιόδου"
        'List1.AddItem "Απ1.Αξ.Εξ.Περιόδου"
        '
        'List1.AddItem "Απ2.Εξ.Περιόδου"
        'List1.AddItem "Απ2.Αξ.Εξ.Περιόδου"

        'List1.AddItem "Βarcode"
        'List1.AddItem "Προμηθευτής"
        'List1.AddItem "Κατηγορία"
        '
        'List1.AddItem "Ημερ.Τελ.Αγοράς"
        'List1.AddItem "Παραγγ.Ποσότητα"
        '
        '
        'List1.AddItem "Στάθμη Παραγγελίας"
        'List1.AddItem "Επιθυμητό Υπόλοιπο"

        'List1.AddItem "Συσκευασία"
        ' List1.AddItem "Υποσυσκευασία"

        'List1.AddItem "Απ3.Εισαγ.Περιόδου"
        'List1.AddItem "Απ3.Αξ.Εισαγ.Περιόδου"
        '
        'List1.AddItem "Απ3.Εξ.Περιόδου"
        'List1.AddItem "Απ3.Αξ.Εξ.Περιόδου"
        '
        '
        'List1.AddItem "Απ4.Εισαγ.Περιόδου"
        'List1.AddItem "Απ4.Αξ.Εισαγ.Περιόδου"
        '
        'List1.AddItem "Απ4.Εξ.Περιόδου"
        'List1.AddItem "Απ4.Αξ.Εξ.Περιόδου"

        'List1.AddItem "Υπολ.Απ3"
        'List1.AddItem "Υπολ.Απ4"

        'DATA2.Recordset("SYSKMIN")

1100    sqlcommand.ConnectionString = gConnect
        'sqlcommand.RecordSource = "SELECT  *FROM SQLCOMMAND"
1110    sqlcommand.RecordSource = "SELECT TOP 10 *FROM EID"
1120    sqlcommand.Refresh

        'If sqlcommand.Recordset.RecordCount < 100 Then
        '  sqlcommand.Recordset.AddNew
        ' sqlcommand.Recordset(0) = ""
        ' sqlcommand.Recordset(1) = "..."
        ' sqlcommand.Recordset(2) = "0000"

        ' sqlcommand.Recordset.Update
        'End If

        'DBGrid1.Columns(0).width = 1020
        'DBGrid1.Columns(1).width = 3495
        'DBGrid1.Columns(2).width = 1320
        'DBGrid1.Columns(3).width = 1095
        '
        'DBGrid3.Columns(0).width = 1020
        'DBGrid3.Columns(1).width = 3500
        'DBGrid3.Columns(2).width = 1320
        'DBGrid3.Columns(3).width = 1095

1130    apo = gEnarjh
1140    eos = gLhjh

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "APOT3"
110     PARAMETROI.SHOW 1

        'F_TAB = Val(FindParametroi(1,"APOT2", "F_TAB", "3", "Σε ποιό TAB είναι σταματημένο"))
        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Form_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub OLE1_Updated(Code As Integer)
'
'
'End Sub

'--------------------------------------------------------------------------------

Private Sub update_DBF_from_SQL(arxeio As String, ARXEIOdbf As String)

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

100     Set DB = OpenDatabase("C:\LAGEURO\", False, False, "dBase III;")

110     Set rDBF = DB.OpenRecordset(ARXEIOdbf)

120     rSQL.Open arxeio, Gdb, adOpenDynamic, adLockOptimistic
        ' rDBF.Open "SELECT *FROM " + arxeio, DBF, adOpenDynamic, adLockOptimistic

130     rSQL.MoveFirst

        Dim Z

140     Z = 0

        On Error GoTo WRITEERROR    ' Resume Next

150     Do While Not rSQL.EOF

160         rDBF.AddNew

170         For k = 0 To rDBF.FIELDS.Count - 1
180             Fname = rDBF.FIELDS(k).Name    ' p.x. FNAME=epo    R(0).NAME
190             rDBF(Fname) = rSQL(Fname)    ' rsqk("epo")=r(0)
            Next

200         rDBF.Update
210         Z = Z + 1

220         If Z Mod 10 = 0 Then
230             Me.Caption = Z
            End If

240         rSQL.MoveNext

250         DoEvents

        Loop

260     rSQL.Close
270     rDBF.Close

        Exit Sub

WRITEERROR:
        'List2.Visible = True

        'List2.AddItem Err.Description
280     Resume Next

        '<EhFooter>
        Exit Sub

update_DBF_from_SQL_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.update_DBF_from_SQL " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.update_DBF_from_SQL " & "at line " & Erl

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

100     For Each F_COL In F_COLS

110         If Trim(F_COL.FILTERTEXT) <> "" Then
120             N = N + 1

130             If N > 1 Then
140                 tmp = tmp & " AND "
                End If

150             If Data1.Recordset(F_COL.ColIndex).Type = 5 Then    '
160                 If InStr(">< >= <=  = ", Left(F_COL.FILTERTEXT, 1)) > 0 And Len(F_COL.FILTERTEXT) > 1 And IsNumeric(Right(F_COL.FILTERTEXT, 1)) Then
170                     tmp = tmp & F_COL.DataField & F_COL.FILTERTEXT
                    Else
180                     tmp = tmp & F_COL.DataField & " >  " & Val(F_COL.FILTERTEXT)
                    End If

                Else
190                 tmp = tmp & F_COL.DataField & " LIKE '" & F_COL.FILTERTEXT & "*'"
                End If

            End If

200     Next F_COL

210     getFilter = tmp

        '<EhFooter>
        Exit Function

getFilter_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.getFilter " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.getFilter " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub GR2_DblClick()

        '<EhHeader>
        On Error GoTo GR2_DblClick_Err

        '</EhHeader>

        Dim F

100     F = GR2.Text

        '  TIMText3.Visible = True
        ' kartella
110     GR2.Col = 0
120     apot21.Text1.Text = GR2.Text

130     GR2.Col = 1

140     apot21.Text2.Text = GR2.Text

150     apot21.Label1.Caption = "1"    ' Left(APOTH1.Text, 1)

160     apot21.Label2.Caption = "0"

        'Me.Hide

170     apot21.SHOW

        '<EhFooter>
        Exit Sub

GR2_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.GR2_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.GR2_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub GR2_FilterChange()

        'Private Sub DATAGRID1_FilterChange()
        '<EhHeader>
        On Error GoTo GR2_FilterChange_Err

        '</EhHeader>

        'Gets called when an action is performed on the filter bar

        On Error GoTo errHandler

100     Set F_COLS = GR2.columns

        Dim c As Integer

110     c = GR2.Col

120     GR2.HoldFields

130     Data1.Recordset.Filter = getFilter()

140     GR2.Col = c

150     GR2.EditActive = True

        Exit Sub

errHandler:

160     MsgBox Err.Source & ":" & vbCrLf & Err.Description

        '    Call cmdClearFilter_Click

170     Resume Next

        '<EhFooter>
        Exit Sub

GR2_FilterChange_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.GR2_FilterChange " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.GR2_FilterChange " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub GR2_GroupColMove(ByVal Position As Integer, _
                             ByVal ColIndex As Integer, _
                             Cancel As Integer)

        '<EhHeader>
        On Error GoTo GR2_GroupColMove_Err

        '</EhHeader>

        Dim strSort As String

        Dim Col     As TrueOleDBGrid80.Column

        ' Loop through GroupColumns collection and construct

        ' the sort string for the Sort property of the Recordset

100     For Each Col In GR2.GroupColumns

110         If strSort <> vbNullString Then

120             strSort = strSort & ", "

            End If

130         strSort = strSort & "[" & Col.DataField & "]"

140     Next Col

150     GR2.HoldFields

160     Data1.Recordset.sort = strSort

        '<EhFooter>
        Exit Sub

GR2_GroupColMove_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.GR2_GroupColMove " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.GR2_GroupColMove " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub List1_Click()

        '<EhHeader>
        On Error GoTo List1_Click_Err

        '</EhHeader>
        Dim k As Integer

        Dim SELE

100     SELE = List1.ListIndex

110     For k = 0 To 35

120         If SEL(k) = SELE Then

                Exit For

            End If

130         If IsEmpty(SEL(k)) Then
140             SEL(k) = SELE

                Exit For

            End If

        Next

        'SEL(1) = 2
        'Dim KK
        'For Each KK In SEL()

        'Next

        '<EhFooter>
        Exit Sub

List1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.List1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.List1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

        '<EhHeader>
        On Error GoTo List1_MouseDown_Err

        '</EhHeader>

100     fList1ARX = List1.ListIndex

        '<EhFooter>
        Exit Sub

List1_MouseDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.List1_MouseDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.List1_MouseDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

        '<EhHeader>
        On Error GoTo List1_MouseUp_Err

        '</EhHeader>
        Dim TEL As Integer, ARX As Integer

        Dim s   As String

100     TEL = List1.ListIndex
110     ARX = fList1ARX

        ' Me.Caption = List1.ListIndex
120     If ARX = TEL Then  'ΠΑΡΕΜΕΙΝΕ ΣΤΟ ΙΔΙΟ ΚΕΛΙ
            'ΑΣΤΟ
        Else

130         If TEL > ARX Then
140             s = List1.List(ARX)
150             List1.RemoveItem ARX

160             List1.AddItem s, TEL
170             List1.Selected(TEL) = True

            Else
180             List1.AddItem List1.List(ARX), TEL
190             List1.Selected(TEL) = True
200             List1.RemoveItem ARX + 1
            End If

        End If

        '<EhFooter>
        Exit Sub

List1_MouseUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.List1_MouseUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.List1_MouseUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub List1_DblClick()

    '<EhHeader>
    On Error GoTo List1_DblClick_Err

    '</EhHeader>
    ''        Dim k As Integer
    Dim nc As Integer

    Dim k  As Integer
        
    nc = 0

80  For k = 0 To List1.ListCount - 1

90      If List1.Selected(k) = True Then
            nc = nc + 1
        End If

    Next

    If nc >= List1.ListCount - 1 Then

        For k = 0 To List1.ListCount - 1
            List1.Selected(k) = False
        Next

    Else

        For k = 0 To List1.ListCount - 1
            List1.Selected(k) = True
        Next

    End If

    '<EhFooter>
    Exit Sub

List1_DblClick_Err:
    'MsgBox Err.Description & vbCrLf & _
     "in ADOMERCNEW.apot3.List1_DblClick " & _
     "at line " & Erl, _
     vbExclamation + vbOKOnly, "Application Error"
    SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.List1_DblClick " & "at line " & Erl

    Resume Next

    '</EhFooter>

End Sub



'Function Get2_AJ(ByRef pol As String, _
'                 polepis As String, _
'                 ago As String, _
'                 AGOEPIS As String, _
'                 ByRef PolXParoxh As String, _
'                 Paroxh As String) As Boolean
'
'        '<EhHeader>
'        On Error GoTo Get2_AJ_Err
'
'        '</EhHeader>
'
'        Dim DB As Database
'
'        Dim R  As New ADODB.Recordset, x As String
'
'        'If gConnect = "Access" Then
'        '   Set db = OpenDatabase(gDir, False, False)
'        'Else
'        '   Set db = OpenDatabase(gDir, False, False, gConnect)
'        'End If
'
'100     R.Open "select POL,EIDOS,AJIA_APOU,SHM1 from PARASTAT", Gdb, adOpenDynamic, adLockOptimistic
'
'110     pol = ""
'120     PolXParoxh = ""
'
'130     Paroxh = ""
'
'140     R.MoveFirst
'
'150     Do While Not R.EOF
'
'160         If Not IsNull(R("eidos") + R("pol") + R("ajia_apou")) Then
'170             If R("pol") = "1" And R("ajia_apou") = "3" Then
'180                 pol = pol + "'" + R("eidos") + "',"
'                End If
'
'190             If R("pol") = "1" And R("ajia_apou") = "3" Then
'200                 If Not IsNull(R("shm1")) And Val(Left(R("shm1"), 2)) <> 2 Then
'210                     PolXParoxh = PolXParoxh + "'" + R("eidos") + "',"
'                    End If
'                End If
'
'220             If R("pol") = "1" And R("ajia_apou") = "3" Then
'230                 If Not IsNull(R("shm1")) And Val(Left(R("shm1"), 2)) = 2 Then
'240                     Paroxh = Paroxh + "'" + R("eidos") + "',"
'                    End If
'                End If
'
'250             If R("pol") = "1" And R("ajia_apou") = "4" Then
'260                 polepis = polepis + "'" + R("eidos") + "',"
'                End If
'
'270             If R("pol") = "2" And R("ajia_apou") = "1" Then
'280                 ago = ago + "'" + R("eidos") + "',"
'                End If
'
'290             If R("pol") = "2" And R("ajia_apou") = "2" Then
'300                 AGOEPIS = AGOEPIS + "'" + R("eidos") + "',"
'                End If
'            End If
'
'310         R.MoveNext
'        Loop
'
'320     pol = Left(pol, Len(pol) - 1)
'330     PolXParoxh = Left(PolXParoxh, Len(PolXParoxh) - 1)
'
'        On Error Resume Next
'
'340     Paroxh = Left(Paroxh, Len(Paroxh) - 1)
'
'350     If Len(polepis) > 0 Then
'360         polepis = Left(polepis, Len(polepis) - 1)
'        Else
'370         polepis = ""
'        End If
'
'380     AGOEPIS = Left(AGOEPIS, Len(AGOEPIS) - 1)
'390     ago = Left(ago, Len(ago) - 1)
'400     Get2_AJ = True
'
'        '<EhFooter>
'        Exit Function
'
'Get2_AJ_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.apot3.Get2_AJ " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Get2_AJ " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>
'
'End Function

Private Sub List11_DblClick()

        '<EhHeader>
        On Error GoTo List11_DblClick_Err

        '</EhHeader>
        Dim k As Integer

100     For k = 0 To List11.ListCount - 1
110         List11.Selected(k) = False
        Next

        '<EhFooter>
        Exit Sub

List11_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.List11_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.List11_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub telkin_Click()

        '===================================================================================
        '<EhHeader>
        On Error GoTo telkin_Click_Err

        '</EhHeader>
        Dim arx_arit, tel_arit, buff, kod_tim, ono_par, p1, progr, k

        Dim m_rec, m_synt, synt1, synt2

        Dim SUMAEID As New ADODB.Recordset

100     arx_arit = 0: tel_arit = 0

        Dim polhs As String

        Dim PolXParoxh As String

        Dim Paroxh As String

        Dim AGOEPIS As String

        Dim polepis As String, pol As String, ago As String, APIS

        Dim a

110     a = Get2_AJ(polhs, polepis, ago, AGOEPIS, PolXParoxh, Paroxh)

        Dim EGGTIM As New ADODB.Recordset

120     EGGTIM.Open "SELECT * FROM EGGTIM WHERE LEFT(ATIM,1) IN ('l','L') AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' ORDER BY HME", Gdb, adOpenDynamic, adLockOptimistic

        Dim R As New ADODB.Recordset

        Dim EID As New ADODB.Recordset

130     EGGTIM.MoveFirst

        Dim B, D, ATIM

140     k = 0

        Dim N

150     Do While Not EGGTIM.EOF
            'ΒΡΙΣΚΩ ΤΗΝ ΤΕΛΕΥΤΑΙΑ ΤΙΜΗ ΑΓΟΡΑΣ
160         a = EGGTIM("KODE")
170         D = EGGTIM("HME")
180         ATIM = EGGTIM("atim")
190         k = k + 1

200         If Not IsNull(EGGTIM("kode")) Then
210             Gdb.Execute "UPDATE EID SET HM3='" + Format(D, "MM/DD/YYYY") + "' where KOD='" + a + "'", N
            End If

220         If N > 0 Then
230             MDIForm1.Caption = Format(k, "####") + " - " + Format(D, "dd/mm/yyyy hh:mm") + " "
            End If

240         If (k Mod 10 = 0) Then DoEvents

250         DoEvents
260         EGGTIM.MoveNext
        Loop

270     EGGTIM.Close

280     UPDATE_YPOLOIPA3 "EID", List11, apo, eos

        On Error Resume Next

290     Gdb.Execute "DROP TABLE EIDFOTOS"

300     Gdb.Execute "SELECT * INTO EIDFOTOS FROM EID WHERE POS<=0 AND HM3<'" + Format(DateAdd("D", -90, Now), "MM/DD/YYYY") + "'"

        On Error GoTo telkin_Click_Err

310     SUMAEID.Open "SELECT * FROM EIDFOTOS", Gdb, adOpenDynamic, adLockBatchOptimistic

320     Do While Not SUMAEID.EOF

330         If Len(Dir("C:\MERCVB\IMAGES\" + SUMAEID("KOD") + ".JPG")) > 0 Then
                'FileCopy "C:\MERCVB\IMAGES\" + SUMAEID("KOD") + ".JPG", "C:\DEL\" + SUMAEID("KOD") + ".JPG"
                'If Len(Dir("C:\DEL\" + SUMAEID("KOD") + ".JPG")) > 0 Then
                '    Kill "C:\MERCVB\IMAGES\" + SUMAEID("KOD") + ".JPG"
                '    DoEvents
                '    Me.Caption = SUMAEID("KOD")

                'End If
            End If

340         SUMAEID.MoveNext

        Loop

350     SUMAEID.Close

        '<EhFooter>
        Exit Sub

telkin_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.telkin_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.telkin_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text4_GotFocus()

        '<EhHeader>
        On Error GoTo Text4_GotFocus_Err

        '</EhHeader>

100     'Text4.Width = 6795

        Text4.Left = Frame1.Left
        Text4.Width = Frame1.Width
        




        '<EhFooter>
        Exit Sub

Text4_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Text4_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Text4_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text4_LostFocus()

        '<EhHeader>
        On Error GoTo Text4_LostFocus_Err

        '</EhHeader>

100     Text4.Width = 3705

        '<EhFooter>
        Exit Sub

Text4_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Text4_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Text4_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
