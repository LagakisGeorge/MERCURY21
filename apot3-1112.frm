VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form apot3param 
   BackColor       =   &H00FF0000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
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
   Begin VB.CommandButton cmdKataMhna 
      Caption         =   "Κινήσεις Κατά Μήνα"
      Height          =   360
      Left            =   15000
      TabIndex        =   32
      Top             =   2400
      Width           =   2000
   End
   Begin VB.TextBox TEXT1 
      Height          =   285
      Left            =   8520
      TabIndex        =   31
      Top             =   3840
      Visible         =   0   'False
      Width           =   150
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
      Left            =   15000
      TabIndex        =   29
      Top             =   1680
      Width           =   2000
   End
   Begin VB.TextBox sumes 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   5880
      TabIndex        =   28
      Top             =   10200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Rep_Fields 
      BackColor       =   &H00C0C0FF&
      Height          =   285
      Left            =   5880
      TabIndex        =   26
      Top             =   9720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox WHERE_QUERY 
      Height          =   285
      Left            =   360
      TabIndex        =   25
      Top             =   9720
      Width           =   3855
   End
   Begin VB.CommandButton YPOLOGISMOS 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Υπολογισμός"
      Height          =   360
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Για να υπολογισθούν σωστά τα Υπόλοιπα πρέπει οι ημερομηνίες να είναι από την αρχή της χρονιάς μέχρι σήμερα"
      Top             =   480
      Width           =   2000
   End
   Begin VB.CommandButton CMDeXIT 
      BackColor       =   &H0080FF80&
      Caption         =   "Εξοδος"
      Height          =   450
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3360
      Width           =   2000
   End
   Begin VB.ListBox Parastatika 
      Height          =   645
      ItemData        =   "apot3-1112.frx":0000
      Left            =   4440
      List            =   "apot3-1112.frx":0002
      MultiSelect     =   1  'Simple
      TabIndex        =   16
      Top             =   1440
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2400
      Top             =   10440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton SaveReport 
      Caption         =   "Αποθήκευση Report"
      Height          =   285
      Left            =   12840
      TabIndex        =   15
      Top             =   2400
      Width           =   1770
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   10155
      TabIndex        =   14
      Top             =   720
      Width           =   4425
   End
   Begin VB.TextBox Text4 
      Height          =   1005
      Left            =   10080
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   2820
      Width           =   4545
   End
   Begin VB.CheckBox rCheck 
      Caption         =   "Χρησιμοποιώ το επιλεγμενο report"
      Height          =   285
      Left            =   10170
      TabIndex        =   12
      Top             =   2400
      Width           =   2595
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6360
      TabIndex        =   11
      Text            =   "99"
      Top             =   2640
      Width           =   315
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Είδη σε έλλειψη"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6720
      TabIndex        =   10
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00E0E0E0&
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
      Height          =   270
      Left            =   1320
      TabIndex        =   9
      Top             =   3120
      Width           =   2385
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   3480
      Width           =   2385
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2880
      Top             =   10440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\mercvb\reports\dokeggt3.rpt"
      UserName        =   "sa"
   End
   Begin VB.ComboBox ORDERCOMBO 
      Height          =   315
      ItemData        =   "apot3-1112.frx":0004
      Left            =   7680
      List            =   "apot3-1112.frx":0014
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc sqlcommand 
      Height          =   330
      Left            =   240
      Top             =   10560
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
   Begin VB.CommandButton KRITIRIA 
      Caption         =   "Κριτήρια επιλογής ειδών"
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Top             =   2280
      Width           =   5055
   End
   Begin VB.ListBox FieldNames 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      ItemData        =   "apot3-1112.frx":004C
      Left            =   210
      List            =   "apot3-1112.frx":004E
      MultiSelect     =   1  'Simple
      TabIndex        =   2
      Top             =   1200
      Width           =   3570
   End
   Begin MSAdodcLib.Adodc egg 
      Height          =   330
      Left            =   240
      Top             =   10920
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
      Top             =   10080
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
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "apot3-1112.frx":0050
      Left            =   240
      List            =   "apot3-1112.frx":006C
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   720
      Width           =   3690
   End
   Begin TrueOleDBGrid80.TDBGrid GR2 
      Bindings        =   "apot3-1112.frx":0103
      Height          =   5400
      Left            =   45
      TabIndex        =   5
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
      Height          =   255
      Left            =   6600
      TabIndex        =   6
      Top             =   720
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   450
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   275120129
      CurrentDate     =   38814
   End
   Begin MSComCtl2.DTPicker eos 
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   1080
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   450
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   275120129
      CurrentDate     =   38814
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ετοιμες Εκτυπώσεις"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   10080
      TabIndex        =   18
      Top             =   480
      Width           =   4575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Φίλτρα Εκτύπωσης"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4080
      TabIndex        =   19
      Top             =   480
      Width           =   5895
      Begin VB.ComboBox POLHTHS 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   2520
         Width           =   2535
      End
      Begin VB.CommandButton cmdkritiriaPelaton 
         BackColor       =   &H00FFFF80&
         Caption         =   "Κριτήρια επιλογής Πελατών"
         Height          =   360
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2880
         Width           =   5055
      End
      Begin VB.Label LAB9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ΠΩΛΗΤΗΣ"
         Height          =   195
         Left            =   1680
         TabIndex        =   34
         Top             =   2520
         Width           =   750
      End
      Begin VB.Label Label4 
         Caption         =   "Εως"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Από"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Αποθήκη (99=όλες μαζί)"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   2160
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ταξινόμηση & Στήλες"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   3855
      Begin VB.Label Label1 
         Caption         =   "Τίτλος"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3000
         Width           =   855
      End
   End
   Begin VB.Label Label_titlos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   27
      Top             =   0
      Width           =   75
   End
   Begin VB.Shape Shape2 
      Height          =   3015
      Left            =   10200
      Top             =   720
      Width           =   4455
   End
End
Attribute VB_Name = "apot3param"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim f_PALIAXRONIA


Dim F_KRITHRIA As New KRITHRIA

Dim Feid_KRITHRIA As New KRITHRIA

Dim FGDB   As New ADODB.Connection

Dim F_RUN  As Integer

Dim Fsql   As String, FIsHidden As Boolean
Dim Fsql2   As String, F2IsHidden As Boolean

Dim F_COL  As TrueOleDBGrid80.Column

Dim F_epan As Integer

Dim F_ALIG(70), F_mWidth(50), F_sum(70)
Attribute F_ALIG.VB_VarUserMemId = 1073938437

Dim f_FieldNames(70, 4)
Attribute f_FieldNames.VB_VarUserMemId = 1073938440

Dim F_COLS          As TrueOleDBGrid80.columns

Dim F_UPDYPOL       As Integer

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


'PELATON
Dim FP_LHM1          As String

Dim FP_LHM2          As String

Dim fP_lHM3          As String

Dim fP_lHM4          As String

Dim fP_lHM5          As String

Dim fP_lHM6          As String

Dim fP_lab1          As String

Dim fP_lab2          As String

Dim fP_lab3          As String

Dim fP_lab4          As String

Dim fP_lab5          As String

Dim fP_lab6          As String

Dim fP_num1          As String

Dim fP_num2          As String

Dim fP_num3          As String

Dim fP_num4          As String


Dim fP_CO1          As String

Dim fP_CO2          As String

Dim fP_CO3          As String

Dim fP_CO4          As String

Dim fP_CO5          As String




















































Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)
Attribute fh.VB_VarUserMemId = 1073938456
'Dim fkrithria As New KRITHRIA


Private Sub Check3_Click()

        '<EhHeader>
        On Error GoTo Check3_Click_Err

        '</EhHeader>
100     If Check3.Value = vbChecked Then
110         If Len(Trim(WHERE_QUERY.Text)) > 0 Then

120             WHERE_QUERY.Text = WHERE_QUERY.Text + " AND  EID.POS <= EID.SPA AND EID.EPIUYP>0 "
            Else
130             WHERE_QUERY.Text = " EID.POS <= EID.SPA AND EID.EPIUYP>0 "
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

Private Sub cmdKataMhna_Click()
'ΥΠΟΛΟΓΙΖΕΙ ΤΟ ΡΕΠΟΡΤ ΕΙΤΕ
        ' ΑΠΟ ΑΥΤΑ ΠΟΥ ΕΠΕΛΕΞΑ ΕΙΤΕ
        ' ΑΠΟ ΤΟ ΕΤΟΙΜΟ ΡΕΠΟΡΤ ΠΟΥ ΕΧΩ ΕΠΙΛΕΞΕΙ
        '<EhHeader>
       ' On Error GoTo Command1_Click_Err

        '</EhHeader>

        Dim x  As String, order As String
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

150     For k = 0 To 9: Sum(k) = 0: Next

160     For k = 1 To 18: mWidth(k) = 1: alig(k) = 0: EPIK(k) = "": Next
170     F_RUN = 1

180     Data1.ConnectionString = gConnect
     
190     If F_UPDYPOL = 1 Then
200         'UPDATE_YPOLOIPA3 "EID", FieldNames, apo, eos
        End If
    
210     Me.Caption = GetCurrentTime() - TIM

220     mc = 2

230     If Check2 Then
240         SUgm_str = "000"

            If InStr(Rep_Fields.Text, "select '@@@@@'") = 0 Then
250             Rep_Fields.Text = Replace(Rep_Fields.Text, "select ", "select '@@@@@' AS [A/A],")  '
            End If
        Else
260         SUgm_str = "00"
270      '   Rep_Fields.Text = "select KOD as [Κωδικός],LEFT(ONO,35) as [Περιγραφή]"    '
        End If

        x = Rep_Fields.Text  'ξεκινάω με το αρχικό

        Dim KK

        Dim temp

280     For KK = 0 To FieldNames.ListCount - 1

290         If FieldNames.Selected(KK) Then
                'ο πραγματικός δείκτης (γιατί μπορεί να μετακινήθηκε)
300             temp = Val(Right(FieldNames.List(KK), 2))
310             x = x + f_FieldNames(temp, 0)
320             mc = mc + 1
                'x = x + f_list1(Val(Right(List1.List(KK), 2)), 0)
330             F_sum(mc) = f_FieldNames(temp, 1)
340             F_ALIG(mc) = f_FieldNames(temp, 2)
350             F_mWidth(mc) = f_FieldNames(temp, 3)
360             If f_FieldNames(temp, 1) = 1 Then
370                 SUgm_str = SUgm_str + "1"
                Else
380                 SUgm_str = SUgm_str + "0"
                End If
            End If

        Next

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
           ' LOAD_KINHSEIS_EIDON
580        ' x = x + " FROM EID INNER JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD "
        Else
590        ' x = x + " from EID  inner JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD"
        End If





Dim mWhere As String
mWhere = ""

600     If Len(WHERE_QUERY.Text) > 0 Then
610        mWhere = " KODE IN (SELECT KOD FROM EID WHERE " + WHERE_QUERY.Text + ") " ' Fsql
        End If

620     order = "KOD"

630     For k = 0 To 2

640         If Combo2.Text = Combo2.List(0) Then
650             order = "EID.KOD"
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

'760     x = x + " order by " + order





 Dim query_pelaton As String
            query_pelaton = ""
            If Len(Fsql2) > 0 Then
                 query_pelaton = " PELKOD IN ( select KOD from PEL where EIDOS='e' AND  " + Fsql2 + ")  "
                 If Len(mWhere) > 0 Then
                     mWhere = "WHERE " + query_pelaton + " AND " + mWhere
                 Else
                      mWhere = "WHERE " + query_pelaton
                 End If
            
            Else
                If Len(mWhere) > 0 Then
                    mWhere = "WHERE " + mWhere
                Else
                    
                End If
                    
                
            
            End If

If Val(Left(POLHTHS, 1)) = 0 Then
Else
   If Len(mWhere) = 0 Then
      mWhere = "WHERE " + " ID_NUM IN (SELECT ID_NUM FROM TIM WHERE LEFT(EIDPAR,1)='" + Left(POLHTHS, 1) + "' )"
   Else
      mWhere = mWhere + "AND ID_NUM IN (SELECT ID_NUM FROM TIM WHERE LEFT(EIDPAR,1)='" + Left(POLHTHS, 1) + "' )"
   End If
End If

'If Len(mWhere) = 0 Then
'
'    mWhere = " WHERE  ascii(LEFT(ATIM,1)) IN(" + POL + ")  "
'Else
'    mWhere = mWhere + " AND  ascii(LEFT(ATIM,1)) IN(" + POL + ")  "
'End If
'


        Dim pol As String, polepis As String, ago As String, AGOEPIS As String
        pol = "": polepis = "": ago = "": AGOEPIS = ""
        Dim DUM As String
        
        DUM = Get_AJ_ASCII(pol, polepis, ago, AGOEPIS)





x = "SELECT KODE+'    '+ONOMA,[1] AS [ΙΑΝ],[2] AS [ΦΕΒ],[3] AS [MAP],[4] AS [ΑΠΡ],[5] AS [MAI],[6] AS [IOYN],[7] AS [IOYΛ],[8] AS [AYΓ],[9] AS [ΣΕΠ],[10] AS [OKT],[11] AS [NOE] ,[12] AS [ΔΕΚ]"




x = x + " From ("
x = x + "  SELECT KODE,ONOMA,"
x = x + "    CASE WHEN ascii(LEFT(ATIM,1)) IN(" + pol + ") THEN  POSO ELSE 0 END AS MPOSO,"     'LEFT(ATIM,1) IN ('T','L')

x = x + "  MONTH(HME) As DD "
x = x + "  From EGGTIM " + mWhere + " ) G "   '" order by " + order +
x = x + " Pivot("
x = x + "  Sum (MPOSO)"
x = x + "  FOR DD IN ( [1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12]  )"
x = x + ") p "

x = x + " Union SELECT KODE+' ΑΞΙΑ ',[1] AS [ΙΑΝ],[2] AS [ΦΕΒ],[3] AS [MAP],[4] AS [ΑΠΡ],[5] AS [MAI],[6] AS [IOYN],[7] AS [IOYΛ],[8] AS [AYΓ],[9] AS [ΣΕΠ],[10] AS [OKT],[11] AS [NOE] ,[12] AS [ΔΕΚ] From "
x = x + " (  SELECT KODE,ONOMA,    CASE WHEN ascii(LEFT(ATIM,1)) IN(" + pol + ")  THEN ROUND(POSO*TIMM*(100-EKPT)/100,2)  ELSE 0 END AS MPOSO,  MONTH(HME) As DD   From EGGTIM " + mWhere + " ) G "
x = x + "   Pivot(  Sum (MPOSO)  FOR DD IN ( [1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12]  )) p2;"









770     If rCheck.Value = vbChecked Then    ' AN PAIRNV TO EPILEGMENO REPORT
780         x = Text4.Text
        Else
790         Text4.Text = x
        End If

800     Data1.RecordSource = x

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

1060    GR2.columns(0).Width = 1200
1070    GR2.columns(1).Width = 2000
1080    GR2.columns(2).Width = 1000

1090    Me.MousePointer = vbNormal

1100    If Check1 Then

            'typos2
           ' On Error GoTo Command1_Click_Err

1110        If Len(SUgm_str) < 3 Then SUgm_str = "0111111111111111"
            If Len(Trim(sumes.Text)) > 0 Then SUgm_str = Trim(sumes.Text)

1120        print3_xar UCase(x), SUgm_str, Text2.Text, 0   ' RR.RecordSource
        End If

        'If Check2 Then
        ' Print4_Excel UCase(x), SUgm_str
        'End If

        Exit Sub

lathos:
1130    MsgBox "λάθος " + Err.Description

1140    Resume Next

End Sub

Private Sub cmdkritiriaPelaton_Click()

        '<EhHeader>
        On Error GoTo cmdKRITHRIA_Click_Err

        '</EhHeader>
        Dim N As Integer

100     N = 0

110     If Not F2IsHidden Then

120         With F_KRITHRIA.Grid1
130             N = N + 1: .TextMatrix(N, 0) = "Κωδικός": .TextMatrix(N, 3) = "KOD": .TextMatrix(N, 4) = "c"
140             N = N + 1: .TextMatrix(N, 0) = "Επωνυμία": .TextMatrix(N, 3) = "EPO": .TextMatrix(N, 4) = "c"
150             N = N + 1: .TextMatrix(N, 0) = "Διεύθυνση": .TextMatrix(N, 3) = "DIE": .TextMatrix(N, 4) = "c"
160             N = N + 1: .TextMatrix(N, 0) = "Επάγγελμα": .TextMatrix(N, 3) = "EPA": .TextMatrix(N, 4) = "c"
170             N = N + 1: .TextMatrix(N, 0) = "Υπόλοιπο": .TextMatrix(N, 3) = "TYP": .TextMatrix(N, 4) = "n"
180             N = N + 1: .TextMatrix(N, 0) = "ΧΡΕΩΣΗ ΕΤΟΥΣ": .TextMatrix(N, 3) = "XRESYN": .TextMatrix(N, 4) = "n"
190             N = N + 1: .TextMatrix(N, 0) = "ΠΙΣΤΩΣΗ ΕΤΟΥΣ": .TextMatrix(N, 3) = "PISSYN": .TextMatrix(N, 4) = "n"
200             N = N + 1: .TextMatrix(N, 0) = "Πόλη": .TextMatrix(N, 3) = "POL": .TextMatrix(N, 4) = "c"
210             N = N + 1: .TextMatrix(N, 0) = "Hμ.Τελ.Ενημέρωσης": .TextMatrix(N, 3) = "LASTUPDT": .TextMatrix(N, 4) = "d"
220             N = N + 1: .TextMatrix(N, 0) = "Κατηγορία-1": .TextMatrix(N, 3) = "KODGAL": .TextMatrix(N, 4) = "c"
230             N = N + 1: .TextMatrix(N, 0) = "Τρόπος Πληρωμής": .TextMatrix(N, 3) = "ARPARAG": .TextMatrix(N, 4) = "c"
240             N = N + 1: .TextMatrix(N, 0) = "Τηλέφωνο": .TextMatrix(N, 3) = "THL": .TextMatrix(N, 4) = "c"

250             N = N + 1: .TextMatrix(N, 0) = "Εκπτωση": .TextMatrix(N, 3) = "PEK": .TextMatrix(N, 4) = "n"
260             N = N + 1: .TextMatrix(N, 0) = "Πλαφόν": .TextMatrix(N, 3) = "PLAFON": .TextMatrix(N, 4) = "n"

270             N = N + 1: .TextMatrix(N, 0) = "Καθ.ΦΠΑ": .TextMatrix(N, 3) = "TYPOS": .TextMatrix(N, 4) = "c"
280             N = N + 1: .TextMatrix(N, 0) = "Μεταφορική": .TextMatrix(N, 3) = "PLAISIO": .TextMatrix(N, 4) = "c"
290             N = N + 1: .TextMatrix(N, 0) = "Τρόπος Πληρωμής": .TextMatrix(N, 3) = "ARPARAG": .TextMatrix(N, 4) = "c"
300             N = N + 1: .TextMatrix(N, 0) = "Κατηγορία": .TextMatrix(N, 3) = "KODGAL": .TextMatrix(N, 4) = "c"
310             N = N + 1: .TextMatrix(N, 0) = "Πωλητής": .TextMatrix(N, 3) = "PVLHTHS": .TextMatrix(N, 4) = "c"

                'TYPOS   KAUESTVS  FPA   CHAR
                'PLAISIO  =METAFORIKH  CHAR
                'ARPARAG  TROPOS PLHRVMHS   CHAR
                'KODGAL  KATHGORIA   CHAR

320             N = N + 1: .TextMatrix(N, 0) = FP_LHM1: .TextMatrix(N, 3) = "HM1": .TextMatrix(N, 4) = "d"
330             N = N + 1: .TextMatrix(N, 0) = FP_LHM2: .TextMatrix(N, 3) = "HM2": .TextMatrix(N, 4) = "d"
340             N = N + 1: .TextMatrix(N, 0) = fP_lHM3: .TextMatrix(N, 3) = "HM3": .TextMatrix(N, 4) = "d"
350             N = N + 1: .TextMatrix(N, 0) = fP_lHM4: .TextMatrix(N, 3) = "HM4": .TextMatrix(N, 4) = "d"

360             N = N + 1: .TextMatrix(N, 0) = fP_lab1: .TextMatrix(N, 3) = "CH1": .TextMatrix(N, 4) = "c"
370             N = N + 1: .TextMatrix(N, 0) = fP_lab2: .TextMatrix(N, 3) = "CH2": .TextMatrix(N, 4) = "c"
380             N = N + 1: .TextMatrix(N, 0) = fP_lab3: .TextMatrix(N, 3) = "CH3": .TextMatrix(N, 4) = "c"
390             N = N + 1: .TextMatrix(N, 0) = fP_lab4: .TextMatrix(N, 3) = "CH4": .TextMatrix(N, 4) = "c"

400             N = N + 1: .TextMatrix(N, 0) = fP_num1: .TextMatrix(N, 3) = "NUM1": .TextMatrix(N, 4) = "N"
410             N = N + 1: .TextMatrix(N, 0) = fP_num2: .TextMatrix(N, 3) = "NUM2": .TextMatrix(N, 4) = "N"
420             N = N + 1: .TextMatrix(N, 0) = fP_num3: .TextMatrix(N, 3) = "NUM3": .TextMatrix(N, 4) = "N"
430             N = N + 1: .TextMatrix(N, 0) = fP_num4: .TextMatrix(N, 3) = "NUM4": .TextMatrix(N, 4) = "N"


435             N = N + 1: .TextMatrix(N, 0) = fP_CO1: .TextMatrix(N, 3) = "COMB1": .TextMatrix(N, 4) = "c"
436             N = N + 1: .TextMatrix(N, 0) = fP_CO2: .TextMatrix(N, 3) = "COMB2": .TextMatrix(N, 4) = "c"
437             N = N + 1: .TextMatrix(N, 0) = fP_CO3: .TextMatrix(N, 3) = "COMB3": .TextMatrix(N, 4) = "c"
439             N = N + 1: .TextMatrix(N, 0) = fP_CO4: .TextMatrix(N, 3) = "COMB4": .TextMatrix(N, 4) = "c"
                N = N + 1: .TextMatrix(N, 0) = fP_CO5: .TextMatrix(N, 3) = "COMB5": .TextMatrix(N, 4) = "c"




            End With

440         F2IsHidden = True

        End If

450     F_KRITHRIA.SHOW 1
460     Fsql2 = MDIForm1.StatusBar1.Panels(6).Text
470     MDIForm1.StatusBar1.Panels(6).Text = ""
480     Text1.Text = Fsql

        '<EhFooter>
        Exit Sub

cmdKRITHRIA_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3param.cmdKRITIRIAPELATON " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'



Public Sub YPOLOGISMOS_Click()
      ypologizo_Ypoloipa False
End Sub
   
Public Sub ypologizo_Ypoloipa(ByVal mkinhseis As Boolean)

'ΥΠΟΛΟΓΙΖΕΙ ΤΟ ΡΕΠΟΡΤ ΕΙΤΕ
        ' ΑΠΟ ΑΥΤΑ ΠΟΥ ΕΠΕΛΕΞΑ ΕΙΤΕ
        ' ΑΠΟ ΤΟ ΕΤΟΙΜΟ ΡΕΠΟΡΤ ΠΟΥ ΕΧΩ ΕΠΙΛΕΞΕΙ
        '<EhHeader>
        ' ΓΙΑ ΝΑ ΠΕΤΑΕΙ ΕΚΤΟΣ ΑΚΥΡΩΜΕΝΑ ΚΑΙ ΑΚΥΡΩΤΙΚΑ : " and LEFT(ATIM,1) NOT IN (SELECT AKYR FROM PARASTAT WHERE AKYR IS NOT NULL ) AND ISNULL(AKYROMENO,0)=0   "
        On Error GoTo Command1_Click_Err

        '</EhHeader>

        Dim x  As String, order As String
        Dim k  As Integer, R As New ADODB.Recordset    ', sum As Single
        Dim mc As Integer
        Dim EPIK(38), alig(38), mWidth(38)
        Dim Sum(30) As Single
        Dim TIM     As Long
        Dim SUgm_str  'που θα κανει sumes
        Dim Kinhseis As Boolean
        Dim a














100     Kinhseis = mkinhseis 'False

        ' για να βγαζει combo στο grid
        Dim gm_fpa As Integer

        Dim m_kat  As Integer

110     m_kat = 0    ' krato thn stili tis katigorias
120     gm_fpa = 0    ' krato thn stili toy FPA

130     Me.MousePointer = vbHourglass
140     TIM = GetCurrentTime()

150     For k = 0 To 9: Sum(k) = 0: Next

160     For k = 1 To 18: mWidth(k) = 1: alig(k) = 0: EPIK(k) = "": Next
170     F_RUN = 1

180     Data1.ConnectionString = gConnect
     
190     If F_UPDYPOL = 1 Then
200         UPDATE_YPOLOIPA3 "EID", FieldNames, apo, eos
        End If
    
210     Me.Caption = GetCurrentTime() - TIM

220     mc = 2

230     If Check2 Then
240         SUgm_str = "000"

            If InStr(Rep_Fields.Text, "select '@@@@@'") = 0 Then
250             Rep_Fields.Text = Replace(Rep_Fields.Text, "select ", "select '@@@@@' AS [A/A],")  '
            End If
        Else
260         SUgm_str = "00"
270      '   Rep_Fields.Text = "select KOD as [Κωδικός],LEFT(ONO,35) as [Περιγραφή]"    '
        End If

        x = Rep_Fields.Text  'ξεκινάω με το αρχικό

        Dim KK

        Dim temp

280     For KK = 0 To FieldNames.ListCount - 1

290         If FieldNames.Selected(KK) Then
                'ο πραγματικός δείκτης (γιατί μπορεί να μετακινήθηκε)
300             temp = Val(Right(FieldNames.List(KK), 2))
310             x = x + f_FieldNames(temp, 0)
320             mc = mc + 1
                'x = x + f_list1(Val(Right(List1.List(KK), 2)), 0)
330             F_sum(mc) = f_FieldNames(temp, 1)
340             F_ALIG(mc) = f_FieldNames(temp, 2)
350             F_mWidth(mc) = f_FieldNames(temp, 3)
360             If f_FieldNames(temp, 1) = 1 Then
370                 SUgm_str = SUgm_str + "1"
                Else
380                 SUgm_str = SUgm_str + "0"
                End If
            End If

        Next

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
            LOAD_KINHSEIS_EIDON
580         x = x + " FROM EID INNER JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD "
        Else
590         x = x + " from EID  inner JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD"
        End If







600     If Len(WHERE_QUERY.Text) > 0 Then
610         x = x + " WHERE " + WHERE_QUERY.Text    ' Fsql
        End If

620     order = "KOD"

630     For k = 0 To 2

640         If Combo2.Text = Combo2.List(0) Then
650             order = "EID.KOD"
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

760     x = x + " order by " + order

770     If rCheck.Value = vbChecked Then    ' AN PAIRNV TO EPILEGMENO REPORT
780         x = Text4.Text
        Else
790         Text4.Text = x
        End If

800     Data1.RecordSource = x

        On Error GoTo lathos

810     Data1.Refresh

        On Error Resume Next

        Dim xd As New TrueOleDBGrid80.ValueItem

        ' On Error GoTo Command1_Click_Err

820     If gm_fpa > 0 Then
830         xd.Value = 1
840         xd.DisplayValue = "13%"
850         GR2.columns(gm_fpa).ValueItems.ADD xd

860         xd.Value = 2
870         xd.DisplayValue = "24%"
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

1060    GR2.columns(0).Width = 1200
1070    GR2.columns(1).Width = 2000
1080    GR2.columns(2).Width = 1000

1090    Me.MousePointer = vbNormal

1100    If Check1 Then

            'typos2
            On Error GoTo Command1_Click_Err

1110        If Len(SUgm_str) < 3 Then SUgm_str = "0111111111111111"
            If Len(Trim(sumes.Text)) > 0 Then SUgm_str = Trim(sumes.Text)

1120        print3_xar UCase(x), SUgm_str, Text2.Text, 0   ' RR.RecordSource
        End If

        'If Check2 Then
        ' Print4_Excel UCase(x), SUgm_str
        'End If

        Exit Sub

lathos:
1130    MsgBox "λάθος " + Err.Description

1140    Resume Next

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
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

Private Sub SaveReport_Click()

        '<EhHeader>
        On Error GoTo SaveReport_Click_Err

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

SaveReport_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Command11_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Command11_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub




Public Sub CMDEXIT_Click()

        '<EhHeader>
        On Error GoTo CMDEXIT_Click_Err

        '</EhHeader>

100     Unload Me

110    ' If gApoMenu = True Then
120    '     SendKeys "%"
130    '     SendKeys "{RIGHT}+{DOWN}+{DOWN}+{DOWN}"
       ' End If

        '<EhFooter>
        Exit Sub

CMDEXIT_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Command2_Click " & "at line " & Erl

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

        mForm_Resize Me, 8, m, 10455, fh, fw, ft, fl
       
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


Private Sub KRITIRIA_Click()

        '<EhHeader>
        On Error GoTo KRITIRIA_Click_Err

        '</EhHeader>
        Dim N As Integer

100     N = 0
110     Check3.Value = vbUnchecked

120     If Not FIsHidden Then

130         'With MDIForm1.fkrithria.Grid1
              gemisma_Krithrion


490         FIsHidden = True

        End If
        'KRITHRIA.Visible = True
        Fsql = ""
        Feid_KRITHRIA.SHOW 1
        
'500     MDIForm1.fkrithria.SHOW 1
510     Fsql = MDIForm1.StatusBar1.Panels(6).Text
520     WHERE_QUERY.Text = Fsql

530     MDIForm1.StatusBar1.Panels(6).Text = ""

        '<EhFooter>
        Exit Sub

KRITIRIA_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

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

        Dim R  As New ADODB.Recordset, x As String

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

150    ' AGNOSTOI.AddItem "ΔΕΝ ΒΡΕΘΗΚΑΝ ΟΙ ΚΩΔΙΚΟΙ:"

160     EGGTIM.MoveFirst

        Dim MATIM, mHME, MM

        Dim B, d

170     Do While Not EGGTIM.EOF

            'ΒΡΙΣΚΩ ΤΗΝ ΤΕΛΕΥΤΑΙΑ ΤΙΜΗ ΑΓΟΡΑΣ
180         a = EGGTIM("KODE")

            ' If Left(a, 10) = "158-60-ΑΠΨ" Then
            '    a = a
            ' End If

190         d = EGGTIM("HME")
200         k = 0
210         R.Open "SELECT top 1 * FROM EGGTIM WHERE KODE='" + a + "' AND LEFT(ATIM,1) IN ('G','Ξ','g','λ') AND HME<='" + Format(d, "MM/DD/YYYY") + "' ORDER BY HME DESC", Gdb, adOpenDynamic, adLockOptimistic

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
360                ' AGNOSTOI.Visible = True
370                 'AGNOSTOI.Refresh
380                 'AGNOSTOI.AddItem A
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
450         Me.Caption = Format(d, "dd/mm/yyyy hh:mm") + " " + Format(k, "####")

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
160         SS = SS + B + " "
        Loop

170     Close #1

180     Text4.Text = SS

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
  ' n JHNJNJKJI
120     If KeyCode = 27 Then CMDEXIT_Click

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
Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     mForm_Load Me, fh, fw, ft, fl

        F_PLATOS_FORMAS = FINDPARAMETROI(1, "APOT3PARAM", "F_PLATOS_FORMAS", "12625", "Πλάτος φόρμας")

        'ειανι η καλυτερη διόρθωση με το δεξιό άκρο του δεξιότερου control (εξοδος =command2)
        If F_PLATOS_FORMAS > 0 Then F_PLATOS_FORMAS = CmdExit.Left + CmdExit.Width + 100

        On Error Resume Next

110   '  Combo1.Text = Combo1.List(0)
120    ' Combo2.Text = Combo2.List(1)
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




load_pedia_XRHSTH


'ΠΕΔΙΑ ΧΡΗΣΤΗ ΠΕΛΑΤΗ
     FP_LHM1 = FINDPARAMETROI(1, "PELAT2", "F_lhm1", "Ημερ1", "Ετικέτα ημερ.1")
     FP_LHM2 = FINDPARAMETROI(1, "PELAT2", "F_lhm2", "Ημερ2", "Ετικέτα ημερ.2")
     fP_lHM3 = FINDPARAMETROI(1, "PELAT2", "F_lhm3", "Ημερ3", "Ετικέτα ημερ.3")
     fP_lHM4 = FINDPARAMETROI(1, "PELAT2", "F_lhm4", "Ημερ4", "Ετικέτα ημερ.4")
     fP_lHM5 = FINDPARAMETROI(1, "PELAT2", "F_lhm5", "Ημερ5", "Ετικέτα ημερ.5")
     fP_lHM6 = FINDPARAMETROI(1, "PELAT2", "F_lhm6", "Ημερ6", "Ετικέτα ημερ.6")

     fP_lab1 = FINDPARAMETROI(1, "PELAT2", "F_lab1", "Κειμ1", "Ετικέτα κειμ.1")
     fP_lab2 = FINDPARAMETROI(1, "PELAT2", "F_lab2", "Κειμ2", "Ετικέτα κειμ.2")
     fP_lab3 = FINDPARAMETROI(1, "PELAT2", "F_lab3", "Κειμ3", "Ετικέτα κειμ.3")
     fP_lab4 = FINDPARAMETROI(1, "PELAT2", "F_lab4", "Κειμ4", "Ετικέτα κειμ.4")
     fP_lab5 = FINDPARAMETROI(1, "PELAT2", "F_lab5", "Κειμ5", "Ετικέτα κειμ.5")
     fP_lab6 = FINDPARAMETROI(1, "PELAT2", "F_lab6", "Κειμ6", "Ετικέτα κειμ.6")

     fP_num1 = FINDPARAMETROI(1, "PELAT2", "F_num1", "Αριθ1", "Ετικέτα κειμ.1")
     fP_num2 = FINDPARAMETROI(1, "PELAT2", "F_num2", "Αριθ2", "Ετικέτα κειμ.2")
     fP_num3 = FINDPARAMETROI(1, "PELAT2", "F_num3", "Αριθ3", "Ετικέτα κειμ.3")
     fP_num4 = FINDPARAMETROI(1, "PELAT2", "F_num4", "Αριθ4", "Ετικέτα κειμ.4")


        fP_CO1 = FINDPARAMETROI(52, "PELAT2", "f_lbcb1", "", "Επιλογές π.χ. Εμπορος;Ιδιώτης;Υπάλληλος")
        fP_CO2 = FINDPARAMETROI(54, "PELAT2", "f_lbcb2", "", "Επιλογές π.χ. Εμπορος;Ιδιώτης;Υπάλληλος")
        fP_CO3 = FINDPARAMETROI(56, "PELAT2", "f_lbcb3", "", "Επιλογές π.χ. Εμπορος;Ιδιώτης;Υπάλληλος")
        fP_CO4 = FINDPARAMETROI(58, "PELAT2", "f_lbcb4", "", "Επιλογές π.χ. Εμπορος;Ιδιώτης;Υπάλληλος")
        fP_CO5 = FINDPARAMETROI(60, "PELAT2", "f_lbcb5", "", "Επιλογές π.χ. Εμπορος;Ιδιώτης;Υπάλληλος")










        'ΣΤΗΛΕΣ ΠΙΝΑΚΑ
        ' (n,0)=sql string  (n,1)=sum  (n,2)=alignment  (n,3)=width  (n,4)=listbox

' load_sthles_EIDON

        Dim k As Integer


        Dim r3 As New ADODB.Recordset

        'GEMISMA MULTI SELECT LISTBOX
990     r3.Open "select POL,TITLOS,EIDOS,AJIA_APOU,POS_APOU from PARASTAT where (POL=1 OR POL=2) AND POS_APOU<>0 ORDER BY POL,TITLOS;", Gdb, adOpenDynamic, adLockOptimistic
1000    Parastatika.Clear
1010    r3.MoveFirst

1020    k = 0

1030    Do While Not r3.EOF

1040        If Not IsNull(r3("TITLOS")) And Left(r3("TITLOS"), 1) <> "-" Then
1050            Parastatika.AddItem r3("TITLOS") + Space(1) + r3("eidos")
1060            Parastatika.Selected(k) = True
1070            k = k + 1
                'mEidos_Par(K) = R("eidos")
            End If

1080        r3.MoveNext
        Loop

1090    r3.Close




'ΦΟΡΤΩΜΑ ΠΩΛΗΤΩΝ
  Dim R As New ADODB.Recordset
            POLHTHS.Clear
            R.Open "SELECT *FROM PINAKES WHERE TYPOS=5 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

            Do While Not R.EOF
                POLHTHS.AddItem LTrim(str(R("AYJON"))) + "." + R("PERIGRAFH")
                R.MoveNext
            Loop
            
            'ΟΤΑΝ ΔΕΝ ΕΧΕΙ ΠΩΛΗΤΗ
            POLHTHS.AddItem "   "

            R.Close










1100    sqlcommand.ConnectionString = gConnect
1110    sqlcommand.RecordSource = "SELECT TOP 10 *FROM EID"
1120    sqlcommand.Refresh


1130    apo = gEnarjh
1140    eos = gLhjh

        Exit Sub

Form_Load_Err:
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Form_Load " & "at line " & Erl
        Resume Next
End Sub

Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "APOT3PARAM"
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

'Private Sub update_DBF_from_SQL(arxeio As String, ARXEIOdbf As String)
'
'        '<EhHeader>
'        On Error GoTo update_DBF_from_SQL_Err
'
'        '</EhHeader>
'
'        Dim DBF    As Database
'
'        Dim sql    As New ADODB.Connection
'
'        Dim rDBF   As Recordset
'
'        Dim rSQL   As New ADODB.Recordset
'
'        Dim conDBF As String
'
'        Dim conSQL As String
'
'        Dim K      As Long
'
'        Dim Fname  As String
'
'        Dim DB     As DAO.Database
'
'100     Set DB = OpenDatabase("C:\LAGEURO\", False, False, "dBase III;")
'
'110     Set rDBF = DB.OpenRecordset(ARXEIOdbf)
'
'120     rSQL.Open arxeio, Gdb, adOpenDynamic, adLockOptimistic
'        ' rDBF.Open "SELECT *FROM " + arxeio, DBF, adOpenDynamic, adLockOptimistic
'
'130     rSQL.MoveFirst
'
'        Dim Z
'
'140     Z = 0
'
'        On Error GoTo WRITEERROR    ' Resume Next
'
'150     Do While Not rSQL.EOF
'
'160         rDBF.AddNew
'
'170         For K = 0 To rDBF.FIELDS.Count - 1
'180             Fname = rDBF.FIELDS(K).Name    ' p.x. FNAME=epo    R(0).NAME
'190             rDBF(Fname) = rSQL(Fname)    ' rsqk("epo")=r(0)
'            Next
'
'200         rDBF.Update
'210         Z = Z + 1
'
'220         If Z Mod 10 = 0 Then
'230             Me.Caption = Z
'            End If
'
'240         rSQL.MoveNext
'
'250         DoEvents
'
'        Loop
'
'260     rSQL.Close
'270     rDBF.Close
'
'        Exit Sub
'
'WRITEERROR:
'        'List2.Visible = True
'
'        'List2.AddItem Err.Description
'280     Resume Next
'
'        '<EhFooter>
'        Exit Sub
'
'update_DBF_from_SQL_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.apot3.update_DBF_from_SQL " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.update_DBF_from_SQL " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>
'
'End Sub

Private Function getFilter() As String

        'Creates the SQL statement in adodc1.recordset.filter
        'and only filters text currently. It must be modified to filter other data types.
        '<EhHeader>
        On Error GoTo getFilter_Err

        '</EhHeader>

        Dim tmp As String

        Dim N   As Integer

100     For Each F_COL In F_COLS

110         If Trim(F_COL.FilterText) <> "" Then
120             N = N + 1

130             If N > 1 Then
140                 tmp = tmp & " AND "
                End If

150             If Data1.Recordset(F_COL.ColIndex).Type = 5 Then    '
160                 If InStr(">< >= <=  = ", Left(F_COL.FilterText, 1)) > 0 And Len(F_COL.FilterText) > 1 And IsNumeric(Right(F_COL.FilterText, 1)) Then
170                     tmp = tmp & F_COL.DataField & F_COL.FilterText
                    Else
180                     tmp = tmp & F_COL.DataField & " >  " & Val(F_COL.FilterText)
                    End If

                Else
190                 tmp = tmp & F_COL.DataField & " LIKE '" & F_COL.FilterText & "*'"
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

        On Error GoTo ErrHandler

100     Set F_COLS = GR2.columns

        Dim C As Integer

110     C = GR2.Col

120     GR2.HoldFields

130     Data1.Recordset.Filter = getFilter()

140     GR2.Col = C

150     GR2.EditActive = True

        Exit Sub

ErrHandler:

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

Private Sub FieldNames_Click()

        '<EhHeader>
        On Error GoTo FieldNames_Click_Err

        '</EhHeader>
        Dim k As Integer

        Dim SELE

100     SELE = FieldNames.ListIndex

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

FieldNames_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.List1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.List1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub FieldNames_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

        '<EhHeader>
        On Error GoTo FieldNames_MouseDown_Err

        '</EhHeader>

100     fList1ARX = FieldNames.ListIndex

        '<EhFooter>
        Exit Sub

FieldNames_MouseDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.List1_MouseDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.List1_MouseDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub FieldNames_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

        '<EhHeader>
        On Error GoTo FieldNames_MouseUp_Err

        '</EhHeader>
        Dim TEL As Integer, ARX As Integer

        Dim s   As String

100     TEL = FieldNames.ListIndex
110     ARX = fList1ARX

        ' Me.Caption = List1.ListIndex
120     If ARX = TEL Then  'ΠΑΡΕΜΕΙΝΕ ΣΤΟ ΙΔΙΟ ΚΕΛΙ
            'ΑΣΤΟ
        Else

130         If TEL > ARX Then
140             s = FieldNames.List(ARX)
150             FieldNames.RemoveItem ARX

160             FieldNames.AddItem s, TEL
170             FieldNames.Selected(TEL) = True

            Else
180             FieldNames.AddItem FieldNames.List(ARX), TEL
190             FieldNames.Selected(TEL) = True
200             FieldNames.RemoveItem ARX + 1
            End If

        End If

        '<EhFooter>
        Exit Sub

FieldNames_MouseUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.List1_MouseUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.List1_MouseUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub FieldNames_DblClick()

    '<EhHeader>
    On Error GoTo FieldNames_DblClick_Err

    '</EhHeader>
    ''        Dim k As Integer
    Dim nc As Integer

    Dim k  As Integer
        
    nc = 0

80  For k = 0 To FieldNames.ListCount - 1

90      If FieldNames.Selected(k) = True Then
            nc = nc + 1
        End If

    Next

    If nc >= FieldNames.ListCount - 1 Then

        For k = 0 To FieldNames.ListCount - 1
            FieldNames.Selected(k) = False
        Next

    Else

        For k = 0 To FieldNames.ListCount - 1
            FieldNames.Selected(k) = True
        Next

    End If

    '<EhFooter>
    Exit Sub

FieldNames_DblClick_Err:
    'MsgBox Err.Description & vbCrLf & _
     "in ADOMERCNEW.apot3.List1_DblClick " & _
     "at line " & Erl, _
     vbExclamation + vbOKOnly, "Application Error"
    SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.List1_DblClick " & "at line " & Erl

    Resume Next

    '</EhFooter>

End Sub

'Sub oldUPDATE_YPOLOIPA3()
'        'into DOKEGGT1
'        'B = "CREATE VIEW dbo.[EIDT3]" _
'
'        '<EhHeader>
'        On Error GoTo UPDATE_YPOLOIPA3_Err
'
'        '</EhHeader>
'
'        Dim a As String
'
'        Dim b As String
'
'        Dim D As Long
'
'        'Dim krit As String
'
'        ' ΒΡΙΣΚΩ ΜΙΑ ΦΟΡΑ ΤΑ ΥΠΟΛΟΙΠΑ ΑΝΑ ΑΠΟΘΗΚΗ
'
'        On Error Resume Next
'
'        ' On Error GoTo UPDATE_YPOLOIPA3_Err
'
'        '----------------------------------------------------------------------------------------
'        'ΠΑΙΡΝΕΙ ΜΟΝΟ ΑΠΟ ΤΟ ΤΟ ΛΙΣΤΒΟΧ ΤΙΣ ΚΙΝΗΣΕΙΣ
'        Dim pol
'
'        Dim k
'
'100     pol = ""
'
'110     For k = 0 To Parastatika.ListCount - 1
'
'120         If Parastatika.Selected(k) = True Then
'130             pol = pol + LTrim(str(Asc(Right(Parastatika.List(k), 1)))) + ","
'            End If
'
'        Next
'
'140     If Len(pol) - 1 > 0 Then
'150         pol = Left(pol, Len(pol) - 1)
'        Else
'160         pol = "' '"
'        End If
'
'        '----------------------------------------------------------------------------------------
'
'170     Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1
'
'180     a = "SELECT KODE," & " SUM(CASE APOT  WHEN 1  THEN XRE ELSE 0  END ) AS S1X," & " SUM(CASE APOT  WHEN 1  THEN PIS ELSE 0  END ) AS S1P," & " SUM(CASE APOT  WHEN 2  THEN XRE ELSE 0  END ) AS S2X," & " SUM(CASE APOT  WHEN 2  THEN PIS ELSE 0  END ) AS S2P," & " SUM(CASE APOT  WHEN 3  THEN XRE ELSE 0  END ) AS S3X," & " SUM(CASE APOT  WHEN 3  THEN PIS ELSE 0  END ) AS S3P," & " SUM(CASE APOT  WHEN 4  THEN XRE ELSE 0  END ) AS S4X," & " SUM(CASE APOT  WHEN 4  THEN PIS ELSE 0  END ) AS S4P  "
'190     a = a + " INTO DOKEGGT1 FROM EGGTIM  "
'200     a = a + "where HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'  GROUP BY KODE"   'AND ascii(left(ATIM,1)) in (" + pol + ")
'
'        Dim TT As Long
'
'210     TT = GetCurrentTime()
'220     Gdb.Execute a, D
'
'        'MsgBox GetCurrentTime() - TT
'
'230     If D = 0 Then
'240         MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"
'
'            Exit Sub
'
'        End If
'
'250     Gdb.Execute "UPDATE EID SET POS01=0,POS02=0,POS03=0,POS04=0,POS=0", D
'
'        'b = "CREATE VIEW dbo.[EIDT3]" _
'        '& "AS " _
'        '& " SELECT     dbo.EID.*, (dbo.DOKEGGT1.S1X-dbo.DOKEGGT1.S1P) AS EXPR1, (dbo.DOKEGGT1.S2X-dbo.DOKEGGT1.S2P) AS EXPR2 " _
'        '& " FROM         dbo.DOKEGGT1 INNER JOIN " _
'        '& " dbo.EID ON dbo.DOKEGGT1.KODE = dbo.EID.KOD"
'        'Gdb.Execute b, d
'
'260     Gdb.Execute "update EID SET" & " POS01=(CASE WHEN S1X IS NULL THEN 0 ELSE S1X END) - (CASE WHEN S1P IS NULL THEN 0 ELSE S1P END)," & " POS02=(CASE WHEN S2X IS NULL THEN 0 ELSE S2X END) - (CASE WHEN S2P IS NULL THEN 0 ELSE S2P END)," & " POS03=(CASE WHEN S3X IS NULL THEN 0 ELSE S3X END) - (CASE WHEN S3P IS NULL THEN 0 ELSE S3P END)," & " POS04=(CASE WHEN S4X IS NULL THEN 0 ELSE S4X END) - (CASE WHEN S4P IS NULL THEN 0 ELSE S4P END)" & " FROM EID INNER JOIN DOKEGGT1 ON EID.KOD=DOKEGGT1.KODE ", D
'
'270     Gdb.Execute "UPDATE EID SET  POS = POS01 + POS02+POS03+POS04", D
'
'280     D = 0
'        'Gdb.Execute "update EIDT3 SET POS01=EXPR1,POS02=EXPR2,POS=EXPR1+EXPR2 ", d
'
'290     D = D + 1
'
'        '<EhFooter>
'        Exit Sub
'
'UPDATE_YPOLOIPA3_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.apot3.UPDATE_YPOLOIPA3 " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.UPDATE_YPOLOIPA3 " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>
'
'End Sub

Function Get2_AJ(ByRef pol As String, _
                 polepis As String, _
                 ago As String, _
                 AGOEPIS As String, _
                 ByRef PolXParoxh As String, _
                 Paroxh As String) As Boolean

        '<EhHeader>
        On Error GoTo Get2_AJ_Err

        '</EhHeader>

        Dim DB As Database

        Dim R  As New ADODB.Recordset, x As String

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

100     R.Open "select POL,EIDOS,AJIA_APOU,SHM1 from PARASTAT", Gdb, adOpenDynamic, adLockOptimistic

110     pol = ""
120     PolXParoxh = ""

130     Paroxh = ""

140     R.MoveFirst

150     Do While Not R.EOF

160         If Not IsNull(R("eidos") + R("pol") + R("ajia_apou")) Then
170             If R("pol") = "1" And R("ajia_apou") = "3" Then
180                 pol = pol + "'" + R("eidos") + "',"
                End If

190             If R("pol") = "1" And R("ajia_apou") = "3" Then
200                 If Not IsNull(R("shm1")) And Val(Left(R("shm1"), 2)) <> 2 Then
210                     PolXParoxh = PolXParoxh + "'" + R("eidos") + "',"
                    End If
                End If

220             If R("pol") = "1" And R("ajia_apou") = "3" Then
230                 If Not IsNull(R("shm1")) And Val(Left(R("shm1"), 2)) = 2 Then
240                     Paroxh = Paroxh + "'" + R("eidos") + "',"
                    End If
                End If

250             If R("pol") = "1" And R("ajia_apou") = "4" Then
260                 polepis = polepis + "'" + R("eidos") + "',"
                End If

270             If R("pol") = "2" And R("ajia_apou") = "1" Then
280                 ago = ago + "'" + R("eidos") + "',"
                End If

290             If R("pol") = "2" And R("ajia_apou") = "2" Then
300                 AGOEPIS = AGOEPIS + "'" + R("eidos") + "',"
                End If
            End If

310         R.MoveNext
        Loop

320     pol = Left(pol, Len(pol) - 1)
330     PolXParoxh = Left(PolXParoxh, Len(PolXParoxh) - 1)

        On Error Resume Next

340     Paroxh = Left(Paroxh, Len(Paroxh) - 1)

350     If Len(polepis) > 0 Then
360         polepis = Left(polepis, Len(polepis) - 1)
        Else
370         polepis = ""
        End If

380     AGOEPIS = Left(AGOEPIS, Len(AGOEPIS) - 1)
390     ago = Left(ago, Len(ago) - 1)
400     Get2_AJ = True

        '<EhFooter>
        Exit Function

Get2_AJ_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.Get2_AJ " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.Get2_AJ " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub Parastatika_DblClick()

        '<EhHeader>
        On Error GoTo Parastatika_DblClick_Err

        '</EhHeader>
        Dim k As Integer

100     For k = 0 To Parastatika.ListCount - 1
110         Parastatika.Selected(k) = False
        Next

        '<EhFooter>
        Exit Sub

Parastatika_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.List11_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot3.List11_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub


Private Sub Text4_GotFocus()

        '<EhHeader>
        On Error GoTo Text4_GotFocus_Err

        '</EhHeader>

100     Text4.Width = 6795

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



Sub load_pedia_XRHSTH()


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


End Sub

Public Sub load_sthles_EIDON()
'========================================================================
        Dim R As Integer
Dim k As Integer

370     R = 0 + 0: f_FieldNames(R, 0) = ",LEFT(MON,5) AS [Mον.Μετ]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 3000: f_FieldNames(R, 4) = "Μον.Μέτρ."
380     R = R + 1: f_FieldNames(R, 0) = ",str(LTI,10,2) as [Χονδ.Τιμή]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 6: f_FieldNames(R, 3) = 1300: f_FieldNames(R, 4) = "Τιμή Xoνδρικής"
390     R = R + 1: f_FieldNames(R, 0) = ",str(LTI5,10,2) as [Λιαν.Τιμή]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 6: f_FieldNames(R, 3) = 1300: f_FieldNames(R, 4) = "Τιμή Λιανικής"
400     R = R + 1: f_FieldNames(R, 0) = ",(XPOS01+XPOS02+XPOS03+XPOS04) AS [Υπόλοιπο]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Υπόλοιπο"

410     R = R + 1: f_FieldNames(R, 0) = ",FPA as [Kατ.ΦΠΑ]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Κατ.ΦΠΑ"
420     R = R + 1: f_FieldNames(R, 0) = ",str(XTI,10,2) AS [Tιμή Αγ.]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 6: f_FieldNames(R, 3) = 1300: f_FieldNames(R, 4) = "Τιμή αγοράς"

430     R = R + 1: f_FieldNames(R, 0) = ",XTI*(XPOS01+XPOS02+XPOS03+XPOS04) AS [Αξ.Απογρ.]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Αξία Απογραφής"

440     R = R + 1: f_FieldNames(R, 0) = ",PONTOI AS [Πόντοι]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Πόντοι"

450     R = R + 1: f_FieldNames(R, 0) = ",ERG as [Βοηθ.Κωδ.] ": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Βarcode"
460     R = R + 1: f_FieldNames(R, 0) = ",PROM AS [Προμηθευτής]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Προμηθευτής"
470     R = R + 1: f_FieldNames(R, 0) = ",KODLOG as [Κατηγορία]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Κατηγορία"



       R = R + 1: f_FieldNames(R, 0) = ",AEG as [Aρ.Οικογ]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Aρ.Οικογ]"
                
                
        Dim query_oik As String
        query_oik = "(SELECT ISNULL ( (SELECT TOP 1 PERIGRAFH  FROM PINAKES WHERE TYPOS=11 AND AYJON=ISNULL(EID.AEG,0)  ) , '---') ) as [Oικογένεια] "
                R = R + 1: f_FieldNames(R, 0) = "," + query_oik + "": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Οικογένεια"





480     R = R + 1: f_FieldNames(R, 0) = ",CONVERT(CHAR(8),HPAR,3)  as [Ημ.Τελ.Αγοράς]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Ημερ.Τελ.Αγοράς"
490     R = R + 1: f_FieldNames(R, 0) = ",PAR as [Παραγγελ.Ποσ]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Παραγγ.Ποσότητα"

500     R = R + 1: f_FieldNames(R, 0) = ",SPA as [ΣτάθμηΠαραγγ]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Στάθμη Παραγγελίας"
510     R = R + 1: f_FieldNames(R, 0) = ",EPIUYP as [Επιθ.Υπόλ]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Επιθυμητό Υπόλοιπο"

520     R = R + 1: f_FieldNames(R, 0) = ",SYSKMAX as [Συσκευασία]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Συσκευασία"
530     R = R + 1: f_FieldNames(R, 0) = ",SYSKMIN as [ΥποΣυσκευασία]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "YποΣυσκευασία"

'======================================================================================
532     R = R + 1: f_FieldNames(R, 0) = ",(CASE APOG01 WHEN NULL THEN 0 ELSE APOG01 END) AS [Απογ.Απ1]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ1.Απογραφή"
        R = R + 1: f_FieldNames(R, 0) = ",(CASE KOSAPOG01 WHEN NULL THEN 0 ELSE KOSAPOG01 END) AS [Αξ.Περ.Απογ.Απ1]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ1.Αξ.Περ.Απογραφής"

540     R = R + 1: f_FieldNames(R, 0) = ",(CASE S1X WHEN NULL THEN 0 ELSE S1X-APOG01 END) AS [Εισ1.Περιόδ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ1.Εισαγ.Περιόδου"
        
550     R = R + 1: f_FieldNames(R, 0) = ",(CASE AS1X WHEN NULL THEN 0 ELSE AS1X-KOSAPOG01 END)  AS [Αξ1.Αγ.Περιόδ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ1.Αξ.Εισαγ.Περιόδου"
560     R = R + 1: f_FieldNames(R, 0) = ",(CASE S1P WHEN NULL THEN 0 ELSE S1P END)  AS [Εξ1.Περιόδ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ1.Εξ.Περιόδου"

570     R = R + 1: f_FieldNames(R, 0) = ",(CASE AS1P WHEN NULL THEN 0 ELSE AS1P END)  AS [Αξ1.Εξ.Περιόδ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ1.Αξ.Εξ.Περιόδου"

580     R = R + 1: f_FieldNames(R, 0) = ",XPOS01 as [Υπ.Απ1]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1300: f_FieldNames(R, 4) = "Υπολ.Απ1"
590     R = R + 1: f_FieldNames(R, 0) = ",XTI*(XPOS01) AS [Αξ.Απογρ.απ.1]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Αξία Απογραφής απ1"


'=======================================================================================
        R = R + 1: f_FieldNames(R, 0) = ",(CASE APOG02 WHEN NULL THEN 0 ELSE APOG02 END) AS [Απογ.Απ2]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ2.Απογραφή"
        R = R + 1: f_FieldNames(R, 0) = ",(CASE KOSAPOG02 WHEN NULL THEN 0 ELSE KOSAPOG02 END) AS [Αξ.Περ.Απογ.Απ2]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ2.Αξια Περ.Απογραφής"
        
        
600     R = R + 1: f_FieldNames(R, 0) = ",(CASE S2X WHEN NULL THEN 0 ELSE S2X-APOG02 END)  AS [Εισ2.Περιόδ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ2.Εισαγ.Περιόδου"
610     R = R + 1: f_FieldNames(R, 0) = ",(CASE AS2X WHEN NULL THEN 0 ELSE AS2X-KOSAPOG02 END) AS [Αξ2.Αγ.Περιόδ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ2.Αξ.Εισαγ.Περιόδου"
620     R = R + 1: f_FieldNames(R, 0) = ",(CASE S2P WHEN NULL THEN 0 ELSE S2P END)  AS [Εξ2.Περιόδ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ2.Εξ.Περιόδου"
630     R = R + 1: f_FieldNames(R, 0) = ",(CASE AS2P WHEN NULL THEN 0 ELSE AS2P END)  AS [Αξ2.Εξ.Περιόδ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ2.Αξ.Εξ.Περιόδου"
640     R = R + 1: f_FieldNames(R, 0) = ",XPOS02 as [Υπ.Απ2]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 6: f_FieldNames(R, 3) = 1300: f_FieldNames(R, 4) = "Υπολ.Απ.2"
650     R = R + 1: f_FieldNames(R, 0) = ",XTI*(XPOS02) AS [Αξ.Απογρ.απ.2]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Αξία Απογραφής απ2"


'===========================================================================================
        R = R + 1: f_FieldNames(R, 0) = ",(CASE APOG03 WHEN NULL THEN 0 ELSE APOG03 END) AS [Απογ.Απ1]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ3.Απογραφή"
        R = R + 1: f_FieldNames(R, 0) = ",(CASE KOSAPOG03 WHEN NULL THEN 0 ELSE KOSAPOG03 END) AS [Αξ.Περ.Απογ.Απ3]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ3.Αξια Περ.Απογραφής"
660     R = R + 1: f_FieldNames(R, 0) = ",(CASE S3X WHEN NULL THEN 0 ELSE S3X-APOG03 END)  AS [Εισ3.Περιόδ]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ3.Εισαγ.Περιόδου"
670     R = R + 1: f_FieldNames(R, 0) = ",(CASE AS3X WHEN NULL THEN 0 ELSE AS3X-KOSAPOG03 END) AS [Αξ3.Αγ.Περιόδ]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ3.Αξ.Εισαγ.Περιόδου"
680     R = R + 1: f_FieldNames(R, 0) = ",(CASE S3P WHEN NULL THEN 0 ELSE S3P END)  AS [Εξ3.Περιόδ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ3.Εξ.Περιόδου"
690     R = R + 1: f_FieldNames(R, 0) = ",(CASE AS3P WHEN NULL THEN 0 ELSE AS3P END)  AS [Αξ3.Αγ.Περιόδ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ3.Αξ.Εξ.Περιόδου"
700     R = R + 1: f_FieldNames(R, 0) = ",XPOS03 as [Υπ.Απ3]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Υπολ.Απ3"
710     R = R + 1: f_FieldNames(R, 0) = ",XTI*(XPOS03) AS [Αξ.Απογρ.απ.3]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Αξία Απογραφής απ3"
      

'===========================================================================================
        R = R + 1: f_FieldNames(R, 0) = ",(CASE APOG04 WHEN NULL THEN 0 ELSE APOG04 END) AS [Απογ.Απ1]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ4.Απογραφή"
        R = R + 1: f_FieldNames(R, 0) = ",(CASE KOSAPOG04 WHEN NULL THEN 0 ELSE KOSAPOG04 END) AS [Αξ.Περ.Απογ.Απ2]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ4.Αξια Περ.Απογραφής"
        
720     R = R + 1: f_FieldNames(R, 0) = ",(CASE S4X WHEN NULL THEN 0 ELSE S4X-APOG04 END)  AS [Εισ4.Περιόδ]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ4.Εισαγ.Περιόδου"
730     R = R + 1: f_FieldNames(R, 0) = ",(CASE AS4X WHEN NULL THEN 0 ELSE AS4X-KOSAPOG04 END) AS [Αξ4.Αγ.Περιόδ]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ4.Αξ.Εισαγ.Περιόδου"
740     R = R + 1: f_FieldNames(R, 0) = ",(CASE S4P WHEN NULL THEN 0 ELSE S4P END)  AS [Εξ4.Περιόδ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ4.Εξ.Περιόδου"
750     R = R + 1: f_FieldNames(R, 0) = ",(CASE AS4P WHEN NULL THEN 0 ELSE AS4P END)  AS [Αξ4.Αγ.Περιόδ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ4.Αξ.Εξ.Περιόδου"
760     R = R + 1: f_FieldNames(R, 0) = ",XPOS04 as [Υπ.Απ4]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Υπολ.Απ4"
770     R = R + 1: f_FieldNames(R, 0) = ",XTI*(XPOS04) AS [Αξ.Απογρ.απ4]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Αξία Απογραφής απ4"

780     R = R + 1: f_FieldNames(R, 0) = ", CONVERT(CHAR(12),HM1,3) AS [" + F_LHM1 + "]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = F_LHM1

790     R = R + 1: f_FieldNames(R, 0) = ", CONVERT(CHAR(12),HM2,3) AS [" + F_LHM2 + "]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = F_LHM2
800     R = R + 1: f_FieldNames(R, 0) = ", CONVERT(CHAR(12),HM3,3) AS [" + f_lHM3 + "]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = f_lHM3
810     R = R + 1: f_FieldNames(R, 0) = ", CH1 AS [" + f_lab1 + "]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = f_lab1
820     R = R + 1: f_FieldNames(R, 0) = ", CH2 AS [" + f_lab2 + "]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = f_lab2
830     R = R + 1: f_FieldNames(R, 0) = ", CH3 AS [" + f_lab3 + "]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = f_lab3
840     R = R + 1: f_FieldNames(R, 0) = ", CH4 AS [" + f_lab4 + "]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = f_lab4
850     R = R + 1: f_FieldNames(R, 0) = ", CH5 AS [" + f_lab5 + "]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = f_lab5
860     R = R + 1: f_FieldNames(R, 0) = ", CH6 AS [" + f_lab6 + "]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = f_lab6
870     R = R + 1: f_FieldNames(R, 0) = ", ONO2 AS [Περιγραφή 2η]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Περιγραφή 2η"
880     R = R + 1: f_FieldNames(R, 0) = ", str(NUM1,10,2) AS [" + f_num1 + "]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = f_num1
890     R = R + 1: f_FieldNames(R, 0) = ", str(NUM2,10,2) AS [" + f_num2 + "]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = f_num2
900     R = R + 1: f_FieldNames(R, 0) = ", str(NUM3,10,2) AS [" + f_num3 + "]": f_FieldNames(R, 1) = 0: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = f_num3
910     R = R + 1: f_FieldNames(R, 0) = ",(CASE ES1P WHEN NULL THEN 0 ELSE ES1P END) AS [ΠΩΛ.ΑΠ1.ΕΠΙΛ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ1.ΠΩΛ.ΕΠΙΛ.Περιόδου"
920     R = R + 1: f_FieldNames(R, 0) = ",(CASE ES2P WHEN NULL THEN 0 ELSE ES2P END) AS [ΠΩΛ.ΑΠ2.ΕΠΙΛ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ2.ΠΩΛ.ΕΠΙΛ.Περιόδου"
930     R = R + 1: f_FieldNames(R, 0) = ",(CASE ES3P WHEN NULL THEN 0 ELSE ES3P END) AS [ΠΩΛ.ΑΠ3.ΕΠΙΛ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ3.ΠΩΛ.ΕΠΙΛ.Περιόδου"
940     R = R + 1: f_FieldNames(R, 0) = ",(CASE ES4P WHEN NULL THEN 0 ELSE ES4P END) AS [ΠΩΛ.ΑΠ4.ΕΠΙΛ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "Απ4.ΠΩΛ.ΕΠΙΛ.Περιόδου"
950     R = R + 1: f_FieldNames(R, 0) = ",MESXTI AS [ΜΕΣΗ ΤΙΜΗ]": f_FieldNames(R, 1) = 1: f_FieldNames(R, 2) = 1: f_FieldNames(R, 3) = 1600: f_FieldNames(R, 4) = "ΜΕΣΗ ΤΙΜΗ"

956     For k = 0 To UBound(f_FieldNames)
960         If Len(f_FieldNames(k, 4)) > 0 Then
970             FieldNames.AddItem f_FieldNames(k, 4) + Space(40) + Format(k, "00")
            End If
        Next

980     FieldNames.ToolTipText = "Με click επιλέγετε το πεδία" + " Με σύρσιμο μπορείτε να αλλάξετε την σειρά"

End Sub



Public Sub LOAD_KINHSEIS_EIDON()
'--------------------------------------------------------------------------------------------------------------------------------------------------------
            On Error Resume Next
            
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
            
            
            
Gdb.Execute "UPDATE EGGTIM SET AJAGOR=0,AJPOL=0"
   On Error GoTo 0
            
Gdb.Execute "UPDATE EGGTIM SET AJAGOR=ABS(POSO)*TIMM*(100-EKPT)/100 WHERE LEFT(ATIM,1) IN (" + AjAg + ")"

Gdb.Execute "UPDATE EGGTIM SET AJAGOR=-ABS(POSO)*TIMM*(100-EKPT)/100 WHERE LEFT(ATIM,1) IN (" + AjAgEp + ")"
            
Gdb.Execute "UPDATE EGGTIM SET AJPOL=ABS(POSO)*TIMM*(100-EKPT)/100 WHERE LEFT(ATIM,1) IN (" + AjPol + ")"

Gdb.Execute "UPDATE EGGTIM SET AJPOL=-ABS(POSO)*TIMM*(100-EKPT)/100 WHERE LEFT(ATIM,1) IN (" + AjPolEp + ")"
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
Dim a As String
Dim k As Long

410         Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1

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

480         For KKE = 0 To Parastatika.ListCount - 1

490             If Parastatika.Selected(KKE) = True Then
500                 pol = pol + LTrim(str(Asc(Right(Parastatika.List(KKE), 1)))) + ","
                End If

            Next

510         If Len(pol) - 1 > 0 Then
520             pol = Left(pol, Len(pol) - 1)
            Else
530             pol = "' '"
            End If
            
            
            Dim query_pelaton As String
            query_pelaton = ""
            If Len(Fsql2) > 0 Then
                 query_pelaton = "PELKOD IN ( select KOD from PEL where EIDOS='e' AND  " + Fsql2 + ") and "
            
            
            End If
            
            
            
            If Val(Left(POLHTHS, 1)) = 0 Then
            Else
                If Len(query_pelaton) = 0 Then
                    query_pelaton = " e.ID_NUM IN (SELECT ID_NUM FROM TIM WHERE LEFT(EIDPAR,1)='" + Left(POLHTHS, 1) + "' ) AND "
                Else
                    query_pelaton = query_pelaton + " e.ID_NUM IN (SELECT ID_NUM FROM TIM WHERE LEFT(EIDPAR,1)='" + Left(POLHTHS, 1) + "' ) AND "
                End If
            End If

            
            

            '----------------------------------------------------------------------------------------

540         a = "SELECT KODE," & " SUM( CASE APOT  WHEN 1  THEN XRE ELSE 0  END ) AS S1X,"
                   a = a + " SUM(CASE APOT  WHEN 1  THEN AJAGOR ELSE 0  END ) AS AS1X,"
                   a = a + " SUM(CASE APOT  WHEN 1  THEN PIS ELSE 0  END ) AS S1P,"
                   a = a + " SUM(CASE APOT  WHEN 1  THEN AJPOL ELSE 0  END ) AS AS1P,"
                   a = a + " SUM( CASE WHEN APOT=1 AND LEFT(e.ATIM,1)='λ'  THEN XRE ELSE 0  END )-SUM(CASE   WHEN APOT=1 AND LEFT(e.ATIM,1)='λ'  THEN PIS ELSE 0  END ) AS APOG01,"
                   a = a + " SUM( CASE WHEN APOT=2 AND LEFT(e.ATIM,1)='λ'  THEN XRE ELSE 0  END )-SUM(CASE   WHEN APOT=2 AND LEFT(e.ATIM,1)='λ'  THEN PIS ELSE 0  END ) AS APOG02,"
                   a = a + " SUM( CASE WHEN APOT=3 AND LEFT(e.ATIM,1)='λ'  THEN XRE ELSE 0  END )-SUM(CASE   WHEN APOT=3 AND LEFT(e.ATIM,1)='λ'  THEN PIS ELSE 0  END ) AS APOG03,"
                   a = a + " SUM( CASE WHEN APOT=4 AND LEFT(e.ATIM,1)='λ'  THEN XRE ELSE 0  END )-SUM(CASE   WHEN APOT=4 AND LEFT(e.ATIM,1)='λ'  THEN PIS ELSE 0  END ) AS APOG04,"

                   a = a + " SUM( CASE WHEN APOT=1 AND LEFT(e.ATIM,1)='λ'  THEN XRE*TIMM ELSE 0  END )-SUM(CASE   WHEN APOT=1 AND LEFT(e.ATIM,1)='λ'  THEN PIS*TIMM ELSE 0  END ) AS KOSAPOG01,"
                   a = a + " SUM( CASE WHEN APOT=2 AND LEFT(e.ATIM,1)='λ'  THEN XRE*TIMM ELSE 0  END )-SUM(CASE   WHEN APOT=2 AND LEFT(e.ATIM,1)='λ'  THEN PIS*TIMM ELSE 0  END ) AS KOSAPOG02,"
                   a = a + " SUM( CASE WHEN APOT=3 AND LEFT(e.ATIM,1)='λ'  THEN XRE*TIMM ELSE 0  END )-SUM(CASE   WHEN APOT=3 AND LEFT(e.ATIM,1)='λ'  THEN PIS*TIMM ELSE 0  END ) AS KOSAPOG03,"
                   a = a + " SUM( CASE WHEN APOT=4 AND LEFT(e.ATIM,1)='λ'  THEN XRE*TIMM ELSE 0  END )-SUM(CASE   WHEN APOT=4 AND LEFT(e.ATIM,1)='λ'  THEN PIS*TIMM ELSE 0  END ) AS KOSAPOG04,"


            a = a + " SUM( CASE APOT  WHEN 1  THEN XRE ELSE 0  END )-SUM(CASE APOT  WHEN 1  THEN PIS ELSE 0  END ) AS XPOS01," _
               & " SUM( CASE APOT  WHEN 2 THEN XRE ELSE 0  END ) AS S2X," _
               & " SUM( CASE APOT  WHEN 2  THEN XRE*TIMM*(1-e.EKPT/100) ELSE 0  END ) AS AS2X," _
               & " SUM(CASE APOT  WHEN 2  THEN PIS*TIMM*(1-e.EKPT/100) ELSE 0  END ) AS AS2P," _
               & " SUM( CASE APOT WHEN 2 THEN PIS ELSE 0 END ) AS S2P," _
               & " SUM( CASE APOT  WHEN 2  THEN XRE ELSE 0  END )-SUM(CASE APOT  WHEN 2  THEN PIS ELSE 0  END ) AS XPOS02," _
               & " SUM( CASE APOT  WHEN 3 THEN XRE ELSE 0  END ) AS S3X," _
               & " SUM( CASE APOT  WHEN 3  THEN XRE*TIMM*(1-e.EKPT/100) ELSE 0  END ) AS AS3X," _
               & " SUM(CASE APOT  WHEN 3  THEN PIS*TIMM*(1-e.EKPT/100) ELSE 0  END ) AS AS3P," _
               & " SUM( CASE APOT WHEN 3 THEN PIS ELSE 0 END ) AS S3P," _
               & " SUM( CASE APOT  WHEN 3  THEN XRE ELSE 0  END )-SUM(CASE APOT  WHEN 3  THEN PIS ELSE 0  END ) AS XPOS03," _
               & " SUM( CASE APOT  WHEN 4 THEN XRE ELSE 0  END ) AS S4X," _
               & " SUM( CASE APOT  WHEN 4  THEN XRE*TIMM*(1-e.EKPT/100) ELSE 0  END ) AS AS4X," _
               & " SUM(CASE APOT  WHEN 4  THEN PIS*TIMM*(1-e.EKPT/100) ELSE 0  END ) AS AS4P," _
               & " SUM( CASE APOT WHEN 4 THEN PIS ELSE 0 END ) AS S4P, " _
               & " SUM( CASE APOT  WHEN 4  THEN XRE ELSE 0  END )-SUM(CASE APOT  WHEN 4  THEN PIS ELSE 0  END ) AS XPOS04," _
               & " SUM(CASE  WHEN APOT=1   AND ascii(left(e.ATIM,1)) in (" + pol + ") THEN PIS ELSE 0  END ) AS ES1P," _
               & " SUM(CASE  WHEN APOT=2   AND ascii(left(e.ATIM,1)) in (" + pol + ") THEN PIS ELSE 0  END ) AS ES2P," _
               & " SUM(CASE  WHEN APOT=3   AND ascii(left(e.ATIM,1)) in (" + pol + ") THEN PIS ELSE 0  END ) AS ES3P," _
               & " SUM(CASE  WHEN APOT=4   AND ascii(left(e.ATIM,1)) in (" + pol + ") THEN PIS ELSE 0  END ) AS ES4P  INTO DOKEGGT1 "
          
550         a = a + " FROM EGGTIM e left join TIM t on e.ID_NUM=t.ID_NUM WHERE ISNULL(t.AKYROMENO,0)=0 AND " + query_pelaton
            a = a + " (NOT APOT IS NULL) AND (NOT e.EKPT IS NULL) AND (NOT XRE IS NULL) AND (NOT PIS IS NULL) AND (NOT TIMM IS NULL) AND "
            a = a + " ( LEFT(e.ATIM,1)='λ' or ( e.HME>='" + Format(apo, "MM/DD/YYYY") + "' AND   e.HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' ) )  GROUP BY KODE"

            '   On Error GoTo Command1_Click_Err

            'Gdb.Execute "UPDATE EGGTIM SET PIS=0 WHERE PIS IS NULL"
560         Gdb.Execute "UPDATE EGGTIM SET EKPT=0 WHERE PIS IS NULL"

570         Gdb.Execute a, k

End Sub


 
Public Sub gemisma_Krithrion()
'==================================================================================================
Dim N As Integer
N = 0
With Feid_KRITHRIA.Grid1   'MDIForm1.fkrithria.Grid1

140             N = N + 1: .TextMatrix(N, 0) = "Περιγραφή": .TextMatrix(N, 3) = "EID.ONO": .TextMatrix(N, 4) = "c"
150             N = N + 1: .TextMatrix(N, 0) = "Οικογένεια": .TextMatrix(N, 3) = "SELECT ISNULL ( (SELECT TOP 1 PERIGRAFH  FROM PINAKES WHERE TYPOS=11 AND AYJON=ISNULL(EID.AEG,0)  ) , '---') AS DD ": .TextMatrix(N, 4) = "n"
160             N = N + 1: .TextMatrix(N, 0) = "Υποοικογένεια": .TextMatrix(N, 3) = "EID.KODLOG": .TextMatrix(N, 4) = "c"

170             N = N + 1: .TextMatrix(N, 0) = "Υπόλοιπο": .TextMatrix(N, 3) = "EID.POS": .TextMatrix(N, 4) = "n"
180             N = N + 1: .TextMatrix(N, 0) = "Υπόλ.Αποθ1": .TextMatrix(N, 3) = "EID.POS01": .TextMatrix(N, 4) = "n"
190             N = N + 1: .TextMatrix(N, 0) = "Υπόλ.Αποθ2": .TextMatrix(N, 3) = "EID.POS02": .TextMatrix(N, 4) = "n"

200             N = N + 1: .TextMatrix(N, 0) = "Υπόλ.Αποθ3": .TextMatrix(N, 3) = "EID.POS03": .TextMatrix(N, 4) = "n"
210             N = N + 1: .TextMatrix(N, 0) = "Υπόλ.Αποθ4": .TextMatrix(N, 3) = "EID.POS04": .TextMatrix(N, 4) = "n"
            
220             N = N + 1: .TextMatrix(N, 0) = "Πωλ.Ποσ.Αποθ1": .TextMatrix(N, 3) = "S1P": .TextMatrix(N, 4) = "n"
230             N = N + 1: .TextMatrix(N, 0) = "Πωλ.Ποσ.Αποθ2": .TextMatrix(N, 3) = "S22": .TextMatrix(N, 4) = "n"
240             N = N + 1: .TextMatrix(N, 0) = "Πωλ.Ποσ.Αποθ3": .TextMatrix(N, 3) = "S3P": .TextMatrix(N, 4) = "n"
250             N = N + 1: .TextMatrix(N, 0) = "Πωλ.Ποσ.Αποθ4": .TextMatrix(N, 3) = "S4P": .TextMatrix(N, 4) = "n"

260             N = N + 1: .TextMatrix(N, 0) = "Προμηθευτής": .TextMatrix(N, 3) = "EID.PROM": .TextMatrix(N, 4) = "c"
270             N = N + 1: .TextMatrix(N, 0) = "BARCODE": .TextMatrix(N, 3) = "EID.ERG": .TextMatrix(N, 4) = "c"
280             N = N + 1: .TextMatrix(N, 0) = "Κωδικός": .TextMatrix(N, 3) = "EID.KOD": .TextMatrix(N, 4) = "c"

290             N = N + 1: .TextMatrix(N, 0) = "Τιμή Αγοράς": .TextMatrix(N, 3) = "EID.XTI": .TextMatrix(N, 4) = "n"
300             N = N + 1: .TextMatrix(N, 0) = "Τιμή Λιανικής": .TextMatrix(N, 3) = "EID.LT5": .TextMatrix(N, 4) = "n"

310             N = N + 1: .TextMatrix(N, 0) = "Τιμή Χονδρικής": .TextMatrix(N, 3) = "EID.LTI": .TextMatrix(N, 4) = "n"

320             N = N + 1: .TextMatrix(N, 0) = "Κατ.ΦΠΑ": .TextMatrix(N, 3) = "EID.FPA": .TextMatrix(N, 4) = "n"

330             N = N + 1: .TextMatrix(N, 0) = "Πόντοι": .TextMatrix(N, 3) = "EID.PONTOI": .TextMatrix(N, 4) = "n"

340             N = N + 1: .TextMatrix(N, 0) = "ΣτάθμηΠαραγγ": .TextMatrix(N, 3) = "EID.SPA": .TextMatrix(N, 4) = "n"
350             N = N + 1: .TextMatrix(N, 0) = "Επιθ.Υπόλ": .TextMatrix(N, 3) = "EID.EPIUYP": .TextMatrix(N, 4) = "n"

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

480             N = N + 1: .TextMatrix(N, 0) = "Περιγραφή 2η": .TextMatrix(N, 3) = "EID.ONO": .TextMatrix(N, 4) = "c"

            End With
            
            
End Sub
