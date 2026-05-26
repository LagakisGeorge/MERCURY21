VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form apot32 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9555
   ScaleWidth      =   14925
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Iσοζύγιο όλων των αποθηκών"
      Height          =   495
      Left            =   10680
      TabIndex        =   28
      Top             =   3720
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5085
      Top             =   8550
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Αποθήκευση Report"
      Height          =   255
      Left            =   10680
      TabIndex        =   27
      Top             =   7440
      Width           =   3540
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   10680
      TabIndex        =   26
      Top             =   4200
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      Height          =   1305
      Left            =   10680
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   6000
      Width           =   3465
   End
   Begin VB.CheckBox rCheck 
      Caption         =   "Χρησιμοποιώ το επιλεγμενο report"
      Height          =   270
      Left            =   10680
      TabIndex        =   24
      Top             =   5640
      Width           =   3465
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "apot32.frx":0000
      Left            =   10755
      List            =   "apot32.frx":000D
      TabIndex        =   23
      Top             =   1845
      Width           =   2640
   End
   Begin VB.CommandButton Command9 
      Caption         =   "ΕΚΤΥΠΩΣΗ ΚΕΡΔΩΝ ΧΩΡΙΣ ΤΗΝ ΠΥ"
      Height          =   315
      Left            =   10800
      TabIndex        =   22
      Top             =   3000
      Width           =   3330
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   12915
      TabIndex        =   20
      Text            =   "99"
      Top             =   1470
      Width           =   315
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6270
      TabIndex        =   19
      Top             =   885
      Width           =   2475
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3540
      TabIndex        =   18
      Top             =   900
      Width           =   2430
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4905
      TabIndex        =   16
      Top             =   1155
      Width           =   5160
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   3435
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   345
      Width           =   3255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "EKTYΠΩΣΗ ΟΛΩΝ ΤΩΝ ΑΠΟΘΗΚΩΝ"
      Height          =   285
      Left            =   10800
      TabIndex        =   14
      Top             =   2280
      Width           =   3285
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Εκτύπωση"
      Enabled         =   0   'False
      Height          =   255
      Left            =   12600
      TabIndex        =   11
      Top             =   6840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   13440
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\mercvb\reports\dokeggt3.rpt"
      UserName        =   "sa"
   End
   Begin VB.ComboBox ORDERCOMBO 
      Height          =   315
      ItemData        =   "apot32.frx":0044
      Left            =   12720
      List            =   "apot32.frx":0054
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ΕΚΤΥΠΩΣΗ ΚΑΤΑ ΑΠΟΘΗΚΗ"
      Height          =   495
      Left            =   12840
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc sqlcommand 
      Height          =   330
      Left            =   6240
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
      Height          =   255
      Left            =   3450
      TabIndex        =   5
      Top             =   45
      Width           =   3270
   End
   Begin VB.ListBox List1 
      Height          =   1185
      ItemData        =   "apot32.frx":008C
      Left            =   240
      List            =   "apot32.frx":008E
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   360
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc egg 
      Height          =   330
      Left            =   2880
      Top             =   9360
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
      Left            =   8520
      Top             =   9360
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
   Begin VB.CommandButton Command2 
      Caption         =   "Εξοδος"
      Height          =   615
      Left            =   12840
      TabIndex        =   3
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10800
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Υπολογισμός"
      Enabled         =   0   'False
      Height          =   195
      Left            =   12000
      TabIndex        =   1
      Top             =   9240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "apot32.frx":0090
      Left            =   240
      List            =   "apot32.frx":00A6
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   0
      Width           =   2655
   End
   Begin TrueOleDBGrid80.TDBGrid GR2 
      Bindings        =   "apot32.frx":010E
      Height          =   7785
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   13732
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
      Left            =   11445
      TabIndex        =   12
      Top             =   165
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   450
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   75759617
      CurrentDate     =   38814
   End
   Begin MSComCtl2.DTPicker eos 
      Height          =   255
      Left            =   11445
      TabIndex        =   13
      Top             =   495
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   450
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   75759617
      CurrentDate     =   38814
   End
   Begin VB.Shape Shape2 
      Height          =   720
      Left            =   3435
      Top             =   825
      Width           =   7020
   End
   Begin VB.Shape Shape1 
      Height          =   2265
      Left            =   10680
      Top             =   1440
      Width           =   3450
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Αποθήκη (99=όλες μαζί)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   10710
      TabIndex        =   21
      Top             =   1530
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Τιτλος εκτύπωσης"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   240
      Left            =   3510
      TabIndex        =   17
      Top             =   1140
      Width           =   1380
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ΕΩΣ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10800
      TabIndex        =   8
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ΑΠΟ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10800
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "apot32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim F_RUN As Integer

Dim Fsql As String, FIsHidden As Boolean

Dim F_COL As TrueOleDBGrid80.Column

Dim F_epan As Integer

Dim F_ALIG(60), F_mWidth(60), F_sum(60)

Dim f_list1(60, 4)

Dim F_COLS As TrueOleDBGrid80.Columns


Dim SEL(35) ' TA PEDIA POY EPILEGV



Dim F_LHM1 As String
Dim F_LHM2 As String
Dim f_lHM3 As String

Dim f_lab1 As String
Dim f_lab2 As String
Dim f_lab3 As String
Dim f_lab4 As String
Dim f_lab5 As String
Dim f_lab6 As String

Dim f_num1 As String
Dim f_num2 As String
Dim f_num3 As String




Private Sub Check3_Click()
  If Check3.value = vbChecked Then
    If Len(Trim(Text1.Text)) > 0 Then
    
       Text1.Text = Text1.Text + " AND  EID.POS <= EID.SPA AND EID.EPIUYP>0 "
    Else
       Text1.Text = " EID.POS <= EID.SPA AND EID.EPIUYP>0 "
    End If
    
    
  End If


End Sub

'


Private Sub Command1_Click()
'ΥΠΟΛΟΓΙΖΕΙ ΤΟ ΡΕΠΟΡΤ ΕΙΤΕ
' ΑΠΟ ΑΥΤΑ ΠΟΥ ΕΠΕΛΕΞΑ ΕΙΤΕ
' ΑΠΟ ΤΟ ΕΤΟΙΜΟ ΡΕΠΟΡΤ ΠΟΥ ΕΧΩ ΕΠΙΛΕΞΕΙ

Dim X As String, order As String
Dim k As Integer, R As New ADODB.Recordset ', sum As Single
Dim mc As Integer
Dim EPIK(38), alig(38), mWidth(38)
Dim sum(30) As Single
Dim TIM As Long
Dim SUgm_str  'που θα κανει sumes
Dim Kinhseis As Boolean
Dim A

Kinhseis = False

' για να βγαζει combo στο grid
Dim gm_fpa As Integer
Dim m_kat As Integer
m_kat = 0 ' krato thn stili tis katigorias
gm_fpa = 0 ' krato thn stili toy FPA



Me.MousePointer = vbHourglass
TIM = GetCurrentTime()

For k = 0 To 9: sum(k) = 0: Next


For k = 1 To 18: mWidth(k) = 1: alig(k) = 0: EPIK(k) = "": Next
F_RUN = 1

DATA1.ConnectionString = gConnect



'Dim R As New ADODB.Recordset

 UPDATE_YPOLOIPA3
'data1.RecordSource = "SELECT KODE,EID.ONO," _
'& " SUM ( CASE APOT  WHEN 1 THEN XRE ELSE 0  END ) AS XRE1," _
'& " SUM ( CASE APOT  WHEN 1 THEN PIS ELSE 0  END )  AS PIS1," _
'& " SUM ( CASE APOT  WHEN 2 THEN XRE ELSE 0  END )  AS XRE2," _
'& " SUM ( CASE APOT  WHEN 2 THEN PIS ELSE 0  END )  AS PIS2 " _
'& " FROM EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD GROUP BY KODE,EID.ONO ORDER BY XRE2" ' , Gdb, adOpenForwardOnly, adLockReadOnly


'data1.Refresh

Me.Caption = GetCurrentTime() - TIM
'Exit Sub

'data1.DatabaseName = gDir

'MSFlexGrid1.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "ΑΡΧΙΚΟ ΥΠΟΛΟΙΠΟ" & Chr(9) & Format(ayp, "###,###.00"), 1


mc = 2
'X = "select KOD as [Κωδικός],LEFT(ONO,35) as [Περιγραφή]"

If Check2 Then
  SUgm_str = "000"
  X = "select  '@@@@@' AS [A/A],KOD AS [Κωδικός],LEFT(ONO,35) as [Περιγραφή]" '
Else
   SUgm_str = "00"
  X = "select KOD as [Κωδικός],LEFT(ONO,35) as [Περιγραφή]" '
End If



'Dim SEL(35)
'SEL(0) = 3
'SEL(1) = 2
Dim KK


For Each KK In SEL
   If Not IsEmpty(KK) And KK >= 0 Then
      X = X + f_list1(KK, 0)
      mc = mc + 1
      F_sum(mc) = f_list1(KK, 1)
      F_ALIG(mc) = f_list1(KK, 2)
      F_mWidth(mc) = f_list1(KK, 3)
      If f_list1(KK, 1) = 1 Then
          SUgm_str = SUgm_str + "1"
      Else
          SUgm_str = SUgm_str + "0"
      End If
      
      
      
   End If
   
Next


'For K = 0 To List1.ListCount - 1
'   If List1.Selected(K) Then
'      x = x + f_list1(K, 0)
'      mc = mc + 1
'      F_sum(mc) = f_list1(K, 1)
'      F_ALIG(mc) = f_list1(K, 2)
'      F_mWidth(mc) = f_list1(K, 3)
'   End If
'Next
 
 
 Kinhseis = True



' Kinhseis = False 'debug


'-------------------------------- NEO ARXH ===============================
If Kinhseis Then
     
      
     Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1

A = "SELECT KODE,SUM( CASE APOT  WHEN 1  THEN XRE ELSE 0  END ) AS S1X," _
& " SUM(CASE APOT  WHEN 1  THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS1X," _
& " SUM(CASE APOT  WHEN 1  THEN PIS ELSE 0  END ) AS S1P," _
& " SUM(CASE APOT  WHEN 1  THEN PIS*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS1P," _
& " SUM( CASE APOT  WHEN 2 THEN XRE ELSE 0  END ) AS S2X," _
& " SUM( CASE APOT  WHEN 2  THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS2X," _
& " SUM(CASE APOT  WHEN 2  THEN PIS*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS2P," _
& " SUM( CASE APOT WHEN 2 THEN PIS ELSE 0 END ) AS S2P," _
& " SUM( CASE APOT  WHEN 3 THEN XRE ELSE 0  END ) AS S3X," _
& " SUM( CASE APOT  WHEN 3  THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS3X," _
& " SUM(CASE APOT  WHEN 3  THEN PIS*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS3P," _
& " SUM( CASE APOT WHEN 3 THEN PIS ELSE 0 END ) AS S3P," _
& " SUM( CASE APOT  WHEN 4 THEN XRE ELSE 0  END ) AS S4X," _
& " SUM( CASE APOT  WHEN 4  THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS4X," _
& " SUM(CASE APOT  WHEN 4  THEN PIS*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS4P," _
& " SUM( CASE APOT WHEN 4 THEN PIS ELSE 0 END ) AS S4P" _
& " INTO DOKEGGT1 "
A = A + " FROM EGGTIM WHERE " _
+ " HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' GROUP BY KODE"
   
'   On Error GoTo 0
   Gdb.Execute A, k
   
   X = X + " FROM EID LEFT JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD "
Else
   X = X + " from EID  inner JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD"
End If


If Len(Text1.Text) > 0 Then
   X = X + " WHERE " + Text1.Text ' Fsql
End If


order = "kod"
For k = 0 To 2
  If Combo2.Text = Combo2.List(0) Then
     order = "EID.KOD"
  ElseIf Combo2.Text = Combo2.List(1) Then
     order = "ONO"
  ElseIf Combo2.Text = Combo2.List(2) Then
     order = "POS"
  ElseIf Combo2.Text = Combo2.List(3) Then
     order = "PROM"
  ElseIf Combo2.Text = Combo2.List(4) Then
     order = "XTI"
  ElseIf Combo2.Text = Combo2.List(5) Then
     order = "XTI*POS"
     
     
     
     
     
  End If
Next




X = X + " order by " + order


If rCheck.value = vbChecked Then
   X = Text4.Text
Else
   Text4.Text = X
End If










DATA1.RecordSource = X

On Error GoTo lathos
   DATA1.Refresh
On Error Resume Next



    Dim xd As New TrueOleDBGrid80.ValueItem
' On Error GoTo 0
     
     
If gm_fpa > 0 Then
     xd.value = 1
     xd.DisplayValue = "9%"
     GR2.Columns(gm_fpa).ValueItems.Add xd
     
     xd.value = 2
     xd.DisplayValue = "19%"
     GR2.Columns(gm_fpa).ValueItems.Add xd
    
    GR2.Columns(gm_fpa).ValueItems.Translate = True
     GR2.Columns(gm_fpa).ValueItems.Presentation = dbgComboBox
End If

If m_kat > 0 Then
    Dim r7 As New ADODB.Recordset
    r7.Open "SELECT *FROM PINAKES WHERE TYPOS=3 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
    Do While Not r7.EOF
       If r7("typos") = 3 Then
          xd.value = Format(r7("ayjon"), "000")
          xd.DisplayValue = r7("PERIGRAFH")
          GR2.Columns(m_kat).ValueItems.Add xd
       Else
          Exit Do
       End If
       r7.MoveNext
    Loop
    ' kathgories

    R.Close






     'xd.Value = 1
     'xd.DisplayValue = "9%"
     'GR2.Columns(gm_fpa).ValueItems.Add xd
     
    
    GR2.Columns(m_kat).ValueItems.Translate = True
     GR2.Columns(m_kat).ValueItems.Presentation = dbgComboBox
End If












'GR2.Columns(0).Text = "Κωδικός"
'GR2.Columns(1).Text = "Περιγραφή"
'GR2.Columns(2).Text = "Ποσότητα"



For k = 5 To 12
    If Len(EPIK(k)) > 1 Then
 '      GR2.Columns(K).Text = epik(K)
       GR2.Columns(k).Alignment = alig(k)
       GR2.Columns(k).width = mWidth(k)
       
     End If
Next

       GR2.Columns(0).width = 1200
       GR2.Columns(1).width = 2000
       GR2.Columns(2).width = 1000

Me.MousePointer = vbNormal

If Check1 Then
  'typos2
  On Error GoTo 0
  If IsEmpty(SUgm_str) Then SUgm_str = "0111111111111111"
  print3_xar UCase(X), SUgm_str, Text2.Text, 0   ' RR.RecordSource
End If


'If Check2 Then
 ' Print4_Excel UCase(x), SUgm_str
'End If

Exit Sub

lathos:
MsgBox "λάθος " + Err.Description


Resume Next






End Sub

'---------------------------------------------
Private Sub Typose()
Dim R, sum, n
Set R = DATA1.Recordset
 
 
 
 Printer.FONTSIZE = 8
 Printer.FontName = "Courier New"
 Printer.Font.Charset = 161
 Printer.FontBold = True


Printer.Print Tab(4); "Κωδικός"; Tab(10); "Επωνυμία"; Tab(36); "Υπόλοιπο"; Tab(48); "Τηλέφωνο"
R.MoveFirst
sum = 0
n = 0
Do While Not R.EOF
    n = n + 1
    Printer.Print Format(n, "###"); Tab(4); R("KOD"); Tab(10); Left(R("EPO"), 25); Tab(36); Format(R("TYP"), "###,###.00"); Tab(48); R("thl")
    sum = sum + R("typ")
    R.MoveNext
Loop
Printer.Print
Printer.Print Tab(36); Format(sum, "###,###.00")
End Sub
Sub typos2()
Dim R, sum, n, db As Database, db2 As Database

Dim X As String, order As String
Dim k As Integer





F_RUN = 1

'data1.DatabaseName = gDir
DATA1.ConnectionString = gConnect

Set db2 = OpenDatabase("c:\mercvb\reports\reports.mdb") ', False, False, gConnect)

On Error Resume Next
db2.Execute "drop table dok2 "
db2.Close



X = "select kod,ono,pos,erg,PROM into dokeid in 'c:\mercvb\reports\reports.mdb' from eid  "

'If IsNumeric(apo.Text) And IsNumeric(eos) Then
'   x = x + " WHERE pos>=" + apo.Text + " and pos<=" + eos.Text
'End If

order = "kod"
For k = 0 To 2
  If Combo2.Text = Combo2.List(0) Then
     order = "kod"
  ElseIf Combo2.Text = Combo2.List(1) Then
     order = "ono"
  ElseIf Combo2.Text = Combo2.List(2) Then
     order = "pos"
  ElseIf Combo2.Text = Combo2.List(3) Then
     order = "prom"
  End If
Next
X = X + " order by " + order
db.Execute X

'CrystalReport1.Action = 1


End Sub



Private Sub Command11_Click()
 If Len(Dir("c:\mercvb\queries\EID", vbDirectory)) = 0 Then
     MkDir "c:\mercvb\queries\EID"
 End If
 
 
 
  Cd1.InitDir = "c:\mercvb\queries\EID"
  Cd1.ShowSave
  
  Dim f
  f = Cd1.FILENAME
  
  Open f For Output As #5
     Print #5, Text4.Text
  Close #5
  
End Sub

Private Sub Command2_Click()
    Unload Me
    
If gApoMenu = True Then
    SendKeys "%"
    SendKeys "{RIGHT}+{DOWN}+{DOWN}+{DOWN}"
End If


End Sub





Private Sub Command3_Click()

'ΕΚΤΥΠΩΣΗ ΟΛΩΝ  ΤΩΝ ΑΠΟΘΗΚΩΝ

Dim A As String
Dim b As String
Dim d As Long

Me.MousePointer = vbHourglass



' UPDATE_YPOLOIPA
UPDATE_YPOLOIPA3

Fsql = Text1.Text


Dim ANS







Dim polhs As String
Dim AGOEPIS As String
Dim polepis As String, pol As String, ago As String, APIS

A = Get_AJ(polhs, polepis, ago, AGOEPIS)




'polhs = "'L','t','T','y','Y','υ'"


Dim apothkes As String
If Val(TEXT3.Text) = 99 Then
    apothkes = ""
Else
    apothkes = " and APOT=" + TEXT3.Text
End If





'
' On Error Resume Next    AS1P>0

On Error GoTo 0
'-------------------------------- NEO ARXH ===============================

Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1





Gdb.Execute "UPDATE EGGTIM SET XRE=0,PIS=POSO WHERE LEFT(ATIM,1) IN(" + polhs + ") and  HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' "
Gdb.Execute "UPDATE EGGTIM SET XRE=0,PIS=-ABS(POSO) WHERE LEFT(ATIM,1) IN(" + polepis + ") and  HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' "

Gdb.Execute "UPDATE EGGTIM SET PIS=0,XRE=POSO WHERE LEFT(ATIM,1) IN(" + ago + ") and  HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' "
Gdb.Execute "UPDATE EGGTIM SET PIS=0,XRE=-ABS(POSO) WHERE LEFT(ATIM,1) IN(" + AGOEPIS + ") and  HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' "


Gdb.Execute "UPDATE EGGTIM SET EKPT=0 WHERE EKPT IS NULL and  HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' "
Gdb.Execute "UPDATE EGGTIM SET PROOD=XRE*TIMM*(100-EKPT)/100 WHERE  HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' "
Gdb.Execute "UPDATE EGGTIM SET PROOD_AJ=PIS*TIMM*(100-EKPT)/100 WHERE  HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' "


A = "SELECT KODE," _
& " " _
& " SUM( XRE ) AS  S1X," _
& " SUM( PIS ) AS S1P," _
& " SUM (PROOD) AS AS1X," _
& " SUM( PROOD_AJ ) AS AS1P  INTO DOKEGGT1 "
A = A + " FROM EGGTIM WHERE  left(ATIM,1) NOT IN ('σ','Σ') AND  " _
+ " HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' " + apothkes + "   GROUP BY KODE"



 Gdb.Execute A, d

If d = 0 Then
   MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"
   Me.MousePointer = vbNormal
   
   Exit Sub
End If

'
A = "SELECT LEFT(KOD,14)AS [ΚΩΔ],LEFT(EID.PROM,10) AS [ΠΡΟΜΗΘ],LEFT(EID.ONO,35) AS [ΠΕΡΙΓΡΑΦΗ],"
A = A + "S1X AS [ΑΓΟΡΕΣ],AS1X AS [ΑΞ.ΑΓΟΡ],S1P AS [ΠΩΛΗΣΕΙΣ],AS1P AS [ΑΞ.ΠΩΛ]"

A = A + " INTO DOKEGGT3 FROM EID  LEFT JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD  "
'  ,(AS2P+AS1P)-(S2P+S1P)*EID.XTI AS [ΚΕΡΔΟΣ]



If Len(Fsql) > 1 Then
     A = A + " where " + Fsql
End If

If ORDERCOMBO.Text = ORDERCOMBO.List(0) Then
  A = A + " ORDER BY EID.ONO"
End If

If ORDERCOMBO.Text = ORDERCOMBO.List(1) Then
  A = A + " ORDER BY KOD"
End If

If ORDERCOMBO.Text = ORDERCOMBO.List(2) Then
  A = A + " ORDER BY EID.ERG"
End If

If ORDERCOMBO.Text = ORDERCOMBO.List(3) Then
  A = A + " ORDER BY AS1P-S1P*EID.XTI "  ' [KEPΔΟΣ]"
End If


' Exit Sub

On Error Resume Next
Gdb.Execute "DROP TABLE DOKEGGT3"  'SOUMES ΑΠΟΘΗΚΗΣ 1
On Error GoTo LATOS
Gdb.Execute A, d
Me.MousePointer = vbNormal

If d = 0 Then
   MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT3"
   Exit Sub
End If

Dim X As String
Dim ANS2

ANS2 = MsgBox("ΜΟΝΟ ΑΥΤΑ ΠΟΥ ΚΙΝΗΘΗΚΑΝ", vbYesNo)
If ANS2 = vbYes Then
    X = "SELECT * FROM DOKEGGT3 WHERE [ΑΓΟΡΕΣ]<>0 OR [ΠΩΛΗΣΕΙΣ]<>0 OR [ΑΞ.ΠΩΛ]<>0 "
Else
    X = "SELECT * FROM DOKEGGT3"
End If


DATA1.ConnectionString = gConnect
DATA1.RecordSource = X
DATA1.Refresh











If Check1 Then
  'typos2
  On Error GoTo 0
  print3_xar UCase(X), "00111111221111111111", "", 0
End If


'If Check2 Then
'  Print4_Excel UCase(x), "00001111111111111111"
'End If

Exit Sub


LATOS:
MsgBox Err.Description
Resume Next



End Sub

Private Sub Command5_Click()
Dim n As Integer
n = 0
Check3.value = vbUnchecked

If Not FIsHidden Then
 With KRITHRIA.Grid1
 
 
'R = R + 1: f_list1(R, 0) = ",S1X AS [ΑΓ.ΠΟΣ1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αγορές Αποθ1"
'R = R + 1: f_list1(R, 0) = ",S1P AS [ΠΩΛ.ΠΟΣ1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Πωλήσ.Αποθ1"
'R = R + 1: f_list1(R, 0) = ", AS1X AS [ΑΞ.ΑΓ1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αξία Αγορών Αποθ1"
'R = R + 1: f_list1(R, 0) = ",AS1P AS [ΑΞ.ΠΩΛ.ΑΠ1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αξία Πωλήσεων Αποθ1"
'
'R = R + 1: f_list1(R, 0) = ",S2X AS [ΑΓ.ΠΟΣ2]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αγορές Αποθ2"
'R = R + 1: f_list1(R, 0) = ",S2P AS [ΠΩΛ.ΠΟΣ2]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Πωλήσ.Αποθ2"
'R = R + 1: f_list1(R, 0) = ", AS2X AS [ΑΞ.ΑΓ2]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αξία Αγορών Αποθ2"
'R = R + 1: f_list1(R, 0) = ",AS2P AS [ΑΞ.ΠΩΛ.ΑΠ2]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αξία Πωλήσεων Αποθ2"
'
'R = R + 1: f_list1(R, 0) = ",S3X AS [ΑΓ.ΠΟΣ3]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αγορές Αποθ3"
'R = R + 1: f_list1(R, 0) = ",S3P AS [ΠΩΛ.ΠΟΣ3]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Πωλήσ.Αποθ3"
'R = R + 1: f_list1(R, 0) = ", AS3X AS [ΑΞ.ΑΓ3]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αξία Αγορών Αποθ3"
'R = R + 1: f_list1(R, 0) = ",AS3P AS [ΑΞ.ΠΩΛ.ΑΠ3]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αξία Πωλήσεων Αποθ3"
'
'R = R + 1: f_list1(R, 0) = ",S4X AS [ΑΓ.ΠΟΣ4]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αγορές Αποθ4"
'R = R + 1: f_list1(R, 0) = ",S4P AS [ΠΩΛ.ΠΟΣ4]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Πωλήσ.Αποθ4"
'R = R + 1: f_list1(R, 0) = ", AS4X AS [ΑΞ.ΑΓ4]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αξία Αγορών Αποθ4"
'R = R + 1: f_list1(R, 0) = ",AS4P AS [ΑΞ.ΠΩΛ.ΑΠ4]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αξία Πωλήσεων Αποθ4"
 
n = n + 1: .TextMatrix(n, 0) = "Ποσ.Αγορ.Αποθ1": .TextMatrix(n, 3) = " S1X ": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Ποσ.Πωλ.Αποθ1": .TextMatrix(n, 3) = " S1P ": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Αξ.Αγορ.Αποθ1": .TextMatrix(n, 3) = " AS1X ": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Αξ.Πωλ.Αποθ1": .TextMatrix(n, 3) = " AS1P ": .TextMatrix(n, 4) = "n"

n = n + 1: .TextMatrix(n, 0) = "Ποσ.Αγορ.Αποθ2": .TextMatrix(n, 3) = " S2X ": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Ποσ.Πωλ.Αποθ2": .TextMatrix(n, 3) = " S2P ": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Αξ.Αγορ.Αποθ2": .TextMatrix(n, 3) = " AS2X ": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Αξ.Πωλ.Αποθ2": .TextMatrix(n, 3) = " AS2P ": .TextMatrix(n, 4) = "n"

n = n + 1: .TextMatrix(n, 0) = "Ποσ.Αγορ.Αποθ3": .TextMatrix(n, 3) = " S3X ": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Ποσ.Πωλ.Αποθ3": .TextMatrix(n, 3) = " S3P ": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Αξ.Αγορ.Αποθ3": .TextMatrix(n, 3) = " AS3X ": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Αξ.Πωλ.Αποθ3": .TextMatrix(n, 3) = " AS3P ": .TextMatrix(n, 4) = "n"

n = n + 1: .TextMatrix(n, 0) = "Ποσ.Αγορ.Αποθ4": .TextMatrix(n, 3) = " S4X ": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Ποσ.Πωλ.Αποθ4": .TextMatrix(n, 3) = " S4P ": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Αξ.Αγορ.Αποθ4": .TextMatrix(n, 3) = " AS4X ": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Αξ.Πωλ.Αποθ4": .TextMatrix(n, 3) = " AS4P ": .TextMatrix(n, 4) = "n"








 
n = n + 1: .TextMatrix(n, 0) = "Περιγραφή": .TextMatrix(n, 3) = "EID.ONO": .TextMatrix(n, 4) = "c"
n = n + 1: .TextMatrix(n, 0) = "Οικογένεια": .TextMatrix(n, 3) = "EID.AEG": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Υποοικογένεια": .TextMatrix(n, 3) = "EID.KODLOG": .TextMatrix(n, 4) = "c"

n = n + 1: .TextMatrix(n, 0) = "Υπόλοιπο": .TextMatrix(n, 3) = "EID.POS": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Υπόλ.Αποθ1": .TextMatrix(n, 3) = "EID.POS01": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Υπόλ.Αποθ2": .TextMatrix(n, 3) = "EID.POS02": .TextMatrix(n, 4) = "n"

n = n + 1: .TextMatrix(n, 0) = "Υπόλ.Αποθ3": .TextMatrix(n, 3) = "EID.POS03": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Υπόλ.Αποθ4": .TextMatrix(n, 3) = "EID.POS04": .TextMatrix(n, 4) = "n"

n = n + 1: .TextMatrix(n, 0) = "Προμηθευτής": .TextMatrix(n, 3) = "EID.PROM": .TextMatrix(n, 4) = "c"
n = n + 1: .TextMatrix(n, 0) = "BARCODE": .TextMatrix(n, 3) = "EID.ERG": .TextMatrix(n, 4) = "c"
n = n + 1: .TextMatrix(n, 0) = "Κωδικός": .TextMatrix(n, 3) = "EID.KOD": .TextMatrix(n, 4) = "c"

n = n + 1: .TextMatrix(n, 0) = "Τιμή Αγοράς": .TextMatrix(n, 3) = "EID.XTI": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Τιμή Λιανικής": .TextMatrix(n, 3) = "EID.LT5": .TextMatrix(n, 4) = "n"
   
n = n + 1: .TextMatrix(n, 0) = "Τιμή Χονδρικής": .TextMatrix(n, 3) = "EID.LTI": .TextMatrix(n, 4) = "n"
   
n = n + 1: .TextMatrix(n, 0) = "Κατ.ΦΠΑ": .TextMatrix(n, 3) = "EID.FPA": .TextMatrix(n, 4) = "n"
   
n = n + 1: .TextMatrix(n, 0) = "Πόντοι": .TextMatrix(n, 3) = "EID.PONTOI": .TextMatrix(n, 4) = "n"

n = n + 1: .TextMatrix(n, 0) = "ΣτάθμηΠαραγγ": .TextMatrix(n, 3) = "EID.SPA": .TextMatrix(n, 4) = "n"
n = n + 1: .TextMatrix(n, 0) = "Επιθ.Υπόλ": .TextMatrix(n, 3) = "EID.EPIUYP": .TextMatrix(n, 4) = "n"


       
        n = n + 1: .TextMatrix(n, 0) = F_LHM1: .TextMatrix(n, 3) = "HM1": .TextMatrix(n, 4) = "d"
        n = n + 1: .TextMatrix(n, 0) = F_LHM2: .TextMatrix(n, 3) = "HM2": .TextMatrix(n, 4) = "d"
        n = n + 1: .TextMatrix(n, 0) = f_lHM3: .TextMatrix(n, 3) = "HM3": .TextMatrix(n, 4) = "d"
        
        n = n + 1: .TextMatrix(n, 0) = f_lab1: .TextMatrix(n, 3) = "CH1": .TextMatrix(n, 4) = "c"
        n = n + 1: .TextMatrix(n, 0) = f_lab2: .TextMatrix(n, 3) = "CH2": .TextMatrix(n, 4) = "c"
        n = n + 1: .TextMatrix(n, 0) = f_lab3: .TextMatrix(n, 3) = "CH3": .TextMatrix(n, 4) = "c"
        n = n + 1: .TextMatrix(n, 0) = f_lab4: .TextMatrix(n, 3) = "CH4": .TextMatrix(n, 4) = "c"
        
        
        n = n + 1: .TextMatrix(n, 0) = f_lab5: .TextMatrix(n, 3) = "CH5": .TextMatrix(n, 4) = "c"
        n = n + 1: .TextMatrix(n, 0) = f_lab6: .TextMatrix(n, 3) = "CH6": .TextMatrix(n, 4) = "c"
        n = n + 1: .TextMatrix(n, 0) = f_num1: .TextMatrix(n, 3) = "NUM1": .TextMatrix(n, 4) = "N"
        n = n + 1: .TextMatrix(n, 0) = f_num2: .TextMatrix(n, 3) = "NUM2": .TextMatrix(n, 4) = "N"
        n = n + 1: .TextMatrix(n, 0) = f_num3: .TextMatrix(n, 3) = "NUM3": .TextMatrix(n, 4) = "N"
   






























  
 End With
 FIsHidden = True
 
End If


    
    KRITHRIA.SHOW 1
    Fsql = MDIForm1.StatusBar1.Panels(6).Text
    Text1.Text = Fsql
    
    
    MDIForm1.StatusBar1.Panels(6).Text = ""
    
End Sub

Private Sub Command6_Click()
'ΕΚΤΥΠΩΣΗ ΚΑΤΑ ΑΠΟΘΗΚΗ
Dim A As String
Dim b As String
Dim d As Long
Dim NAPOT
Dim R As New ADODB.Recordset

 R.Open "SELECT COUNT(*) AS DDD FROM PINAKES WHERE TYPOS=4 ", Gdb, adOpenDynamic, adLockOptimistic
 NAPOT = R(0)
R.Close

UPDATE_YPOLOIPA


'
' On Error Resume Next

On Error GoTo 0
'-------------------------------- NEO ARXH ===============================


'ΥΠΟΛΟΓΙΣΜΟΣ ΙΣΟΖΥΓΙΟΥ


Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1

A = "SELECT KODE," _
& " SUM( CASE APOT  WHEN 1  THEN XRE ELSE 0  END ) AS S1X," _
& " SUM(CASE APOT  WHEN 1  THEN PIS ELSE 0  END ) AS S1P," _
& " SUM( CASE APOT  WHEN 2 THEN XRE ELSE 0  END ) AS S2X," _
& " SUM( CASE APOT WHEN 2 THEN PIS ELSE 0 END ) AS S2P," _
& " SUM( CASE APOT  WHEN 3  THEN XRE ELSE 0  END ) AS S3X," _
& " SUM(CASE APOT  WHEN 3  THEN PIS ELSE 0  END ) AS S3P," _
& " SUM( CASE APOT  WHEN 4 THEN XRE ELSE 0  END ) AS S4X," _
& " SUM( CASE APOT WHEN 4 THEN PIS ELSE 0 END ) AS S4P," _
& " SUM( CASE APOT  WHEN 1  THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS1X," _
& " SUM(CASE APOT  WHEN 1  THEN PIS*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS1P," _
& " SUM( CASE APOT  WHEN 2 THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS2X," _
& " SUM( CASE APOT WHEN 2 THEN PIS*TIMM*(1-EKPT/100) ELSE 0 END ) AS AS2P," _
& " SUM( CASE APOT  WHEN 3 THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS3X," _
& " SUM( CASE APOT WHEN 3 THEN PIS*TIMM*(1-EKPT/100) ELSE 0 END ) AS AS3P, " _
& " SUM( CASE APOT  WHEN 4 THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS4X," _
& " SUM( CASE APOT WHEN 4 THEN PIS*TIMM*(1-EKPT/100) ELSE 0 END ) AS AS4P " _
& " INTO DOKEGGT1 "
A = A + " FROM EGGTIM WHERE " _
+ " HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' GROUP BY KODE"

Gdb.Execute A, d
'
A = "SELECT KOD,LEFT(EID.PROM,20) AS [ΠΡΟΜΗΘ],EID.ERG AS [BARCODE],LEFT(EID.ONO,35) AS [ΠΕΡΙΓΡΑΦΗ],"
A = A + " S1X,S1P,POS01,"
A = A + " S2X,S2P,POS02,"

'If NAPOT > 2 Then 'FAKIS
   A = A + " S3X,S3P,POS03,"
   A = A + " S4X,S4P,POS04,"
'End If

A = A + " AS1X,AS1P,AS2X,AS2P"

'If NAPOT > 2 Then ' FAKIS
   A = A + " ,AS3X,AS3P,AS4X,AS4P"
'End If
A = A + " INTO DOKEGGT3 FROM EID  LEFT JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD"
If Len(Fsql) > 1 Then
'     A = A + " WHERE " + Fsql   κανει παρακατω φίλτρο
End If
If ORDERCOMBO.Text = ORDERCOMBO.List(0) Then
  A = A + " ORDER BY EID.ONO"
End If
If ORDERCOMBO.Text = ORDERCOMBO.List(1) Then
  A = A + " ORDER BY KOD"
End If
If ORDERCOMBO.Text = ORDERCOMBO.List(3) Then
  A = A + " ORDER BY EID.ERG"
End If
On Error Resume Next
Gdb.Execute "DROP TABLE DOKEGGT3"  'SOUMES ΑΠΟΘΗΚΗΣ 1
On Error GoTo 0
Gdb.Execute A, d


If d = 0 Then
   MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT3"
   Exit Sub
End If



Dim X As String
' x = "SELECT * FROM DOKEGGT3"


Dim SUgm_str, mc
'υπολογιζω τα πεδια που θα τυπωθουν
If Check2 Then
  SUgm_str = "000"
  X = "select  '@@@@@' AS [A/A],EID.KOD AS [Κωδικός],LEFT(EID.ONO,35) as [Περιγραφή]" '
Else
   SUgm_str = "00"
  X = "select EID.KOD as [Κωδικός],LEFT(EID.ONO,35) as [Περιγραφή]" '
End If
Dim KK
For Each KK In SEL
   If KK >= 0 Then
      X = X + f_list1(KK, 0)
      mc = mc + 1
      F_sum(mc) = f_list1(KK, 1)
      F_ALIG(mc) = f_list1(KK, 2)
      F_mWidth(mc) = f_list1(KK, 3)
      If f_list1(KK, 1) = 1 Then
          SUgm_str = SUgm_str + "1"
      Else
          SUgm_str = SUgm_str + "0"
      End If
   End If
Next


 X = X + " FROM DOKEGGT3  LEFT JOIN EID  ON DOKEGGT3.KOD=EID.KOD "

If Len(Text1.Text) > 0 Then
   X = X + " WHERE " + Text1.Text ' Fsql
End If



If Check1 Then
  'typos2
  print3_xar UCase(X), SUgm_str, "", 0
End If


'If Check2 Then
'  Print4_Excel UCase(x), "00001111111111111111"
'End If

Exit Sub


'
'
'
'
'
''CrystalReport1.Connect = gConnect
'
''CrystalReport1.Action = 1
'Dim DUM
'
''DUM = Shell("bcp emp.dbo.DOKEGGT3 out region.txt -S PLATO -T -c  -C RAW", vbMinimizedFocus)
'' On Error Resume Next
'FileCopy "C:\LAGEURO\DOKEGGT3SOS.DBF", "C:\LAGEURO\DOKEGGT3.DBF"
'
' update_DBF_from_SQL "SELECT *FROM DOKEGGT3", "DOKEGGT3"
'
'MsgBox "ΑΡΧΕΙΟ EXCEL:" + Chr(13) + "C:\LAGEURO\DOKEGGT3.DBF"
'
'
'
'Exit Sub
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
'
'd = d + 1
''-------------------------------- NEO TELOS ===============================
'
'
'
''---------------------- ΚΙΝΗΣΕΙΣ ---------------------------------------------
'On Error Resume Next
'
'Gdb.Execute "DROP TABLE DOKEGGT1"
'Gdb.Execute "DROP TABLE DOKEGGT2"
'Gdb.Execute "DROP TABLE DOKEGGT3"
'
'A = "SELECT KODE,SUM(XRE) AS S1X,SUM(PIS) AS S1P,0 AS S31,0 AS S32 INTO DOKEGGT1 "
'A = A + " FROM EGGTIM WHERE APOT=1 AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' GROUP BY KODE"
'
'B = "SELECT KODE,SUM(XRE) AS S2X,SUM(PIS) AS S2P INTO DOKEGGT2 "
'B = B + " FROM EGGTIM WHERE APOT=2  AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'  GROUP BY KODE"
'
'Gdb.Execute A, d
'If d = 0 Then
'   MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"
'   Exit Sub
'End If
'd = 0
'
'
'
'Gdb.Execute B, d
'If d = 0 Then
'   MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT2"
'   Exit Sub
'End If
'
'
'
'Gdb.Execute "UPDATE DOKEGGT1 SET S1X=0 WHERE S1X IS NULL"
'
'Gdb.Execute "UPDATE DOKEGGT1 SET S1P=0 WHERE S1P IS NULL"
'
'
'Gdb.Execute "UPDATE DOKEGGT2 SET S2X=0 WHERE S2X IS NULL"
'
'Gdb.Execute "UPDATE DOKEGGT2 SET S2P=0 WHERE S2P IS NULL"
'
'
'
'
'
''DOKEGGT12 VIEW =  ΕΝΩΝΕΙ ΤΑ DOKEGGT1 KAI DOKEGGT2
'' "SELECT DOKEGGT1.*,S2X,S2P FROM DOKEGGT1 LEFT JOIN DOKEGGT2 ON DOKEGGT1.KODE=DOKEGGT2.KODE "
'
'
'
'd = 0
'On Error GoTo 0
'
'A = "SELECT KOD,EID.PROM,EID.ERG,EID.ONO,S1X,S1P,POS01,EXPR1,EXPR2,POS02 "
'A = A + " INTO DOKEGGT3 FROM DOKEGGT12 INNER JOIN EID ON DOKEGGT12.KODE=EID.KOD"
'
'If Len(Fsql) > 1 Then
'     A = A + " WHERE " + Fsql
'End If
'
'If ORDERCOMBO.Text = ORDERCOMBO.List(0) Then
'  A = A + " ORDER BY EID.ONO"
'End If
'
'If ORDERCOMBO.Text = ORDERCOMBO.List(1) Then
'  A = A + " ORDER BY KOD"
'End If
'
'If ORDERCOMBO.Text = ORDERCOMBO.List(3) Then
'  A = A + " ORDER BY EID.ERG"
'End If
'
'
'
'
'
'On Error GoTo 0
'
'Gdb.Execute A, d
'
'If d = 0 Then
'   MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT3"
'   Exit Sub
'End If
'
'' Dim X As String
'x = "SELECT * FROM DOKEGGT3"
'
'If Check1 Then
'  'typos2
'  print3_xar UCase(x), "000000000", "", 0
'End If
'
'
''If Check2 Then
''  Print4_Excel UCase(x), "000000000"
''End If
'
'Exit Sub
'
'
'CrystalReport1.Connect = gConnect
'
'CrystalReport1.Action = 1
''Dim DUM
'
''DUM = Shell("bcp emp.dbo.DOKEGGT3 out region.txt -S PLATO -T -c  -C RAW", vbMinimizedFocus)
'On Error Resume Next
'FileCopy "C:\LAGEURO\DOKEGGT3SOS.DBF", "C:\LAGEURO\DOKEGGT3.DBF"
'
' update_DBF_from_SQL "SELECT *FROM DOKEGGT3", "DOKEGGT3"
'
'MsgBox "ΑΡΧΕΙΟ EXCEL:" + Chr(13) + "C:\LAGEURO\DOKEGGT3.DBF"

End Sub

Private Sub Command7_Click()
'GR2.PrintInfo.Draft = True
GR2.PrintInfo.NoClipping = True
GR2.Styles(6).BackColor = vbWhite
  With GR2.PrintInfo
       
        ' Set the page header
      .PageHeaderFont.Italic = True
        .PageHeader = "Αποθήκη " + Format(Date$, "mm/dd/YYYY")
        
         'Column headers will be on every page
        .RepeatColumnHeaders = True
        
        ' Display page numbers (centered)
        .PageFooter = "\tσελίδα: \p"
        ' Invoke Print Preview
        .PrintPreview
    End With
 

End Sub

Private Sub Command8_Click()
'ΕΚΤΥΠΩΣΗ ΣΤΟ ΣΥΝΟΛΟ ΤΩΝ ΑΠΟΘΗΚΩΝ

Dim A As String
Dim b As String
Dim d As Long

Me.MousePointer = vbHourglass



' UPDATE_YPOLOIPA
UPDATE_YPOLOIPA3

Fsql = Text1.Text


Dim ANS

If F_epan = 1 Then

   ANS = MsgBox("Να υπολογίσω το κέρδος ξανά;", vbYesNo)
  If ANS = vbYes Then
    ypol_kerd
  End If
End If






Dim polhs As String
Dim AGOEPIS As String
Dim polepis As String, pol As String, ago As String, APIS

A = Get_AJ(polhs, polepis, ago, AGOEPIS)




'polhs = "'L','t','T','y','Y','υ'"


Dim apothkes As String
If Val(TEXT3.Text) = 99 Then
    apothkes = ""
Else
    apothkes = " and APOT=" + TEXT3.Text
End If





'
' On Error Resume Next    AS1P>0

On Error GoTo 0
'-------------------------------- NEO ARXH ===============================

Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1

A = "SELECT KODE," _
& " SUM( CASE   WHEN LEFT(ATIM,1)='λ'  THEN XRE ELSE 0  END ) AS S0X," _
& " SUM( CASE   WHEN LEFT(ATIM,1) IN ('G','g')  THEN XRE ELSE 0  END ) AS S1X," _
& " SUM( CASE WHEN PIS IS NULL  OR LEFT(ATIM,1) NOT IN (" + polhs + ")  THEN 0 ELSE PIS END) AS S1P," _
& " SUM( CASE WHEN XRE*TIMM IS NULL THEN 0 ELSE XRE*TIMM*(1- (CASE WHEN EKPT IS NULL THEN 0 ELSE EKPT END)/100) END  ) AS AS1X," _
& " SUM( CASE WHEN POSO*TIMM IS NULL OR LEFT(ATIM,1) NOT IN (" + polhs + ")  THEN 0 ELSE POSO*TIMM*(1-(CASE WHEN EKPT IS NULL THEN 0 ELSE EKPT END)/100) END ) AS AS1P, " _
& " SUM( CASE WHEN KERDOS IS NULL OR LEFT(ATIM,1) NOT IN (" + polhs + ")  THEN 0 ELSE KERDOS END ) AS KERD1 INTO DOKEGGT1 "
A = A + " FROM EGGTIM WHERE  left(ATIM,1) NOT IN ('σ','Σ') AND  " _
+ " HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' " + apothkes + "   GROUP BY KODE"

' ' & " SUM(XRE)-SUM(PIS) AS YP," _


 Gdb.Execute A, d

If d = 0 Then
   MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"
   Me.MousePointer = vbNormal
   
   Exit Sub
End If

'
A = "SELECT LEFT(KOD,14)AS [ΚΩΔ],LEFT(EID.PROM,10) AS [ΠΡΟΜΗΘ],LEFT(EID.ONO,35) AS [ΠΕΡΙΓΡΑΦΗ],"
A = A + "S0X AS [ΑΠΟΓΡ], S1X AS [ΑΓΟΡΕΣ],S1P AS [ΠΩΛΗΣΕΙΣ],S0X+S1X-S1P AS [ΥΠΟΛ],"

A = A + "(CASE WHEN S1P>0 THEN AS1P/S1P ELSE 0  END ) AS [MEΣΤΙΜ.ΠΩΛ],XTI AS [TIM.AΓ],"

If F_epan = 1 Then
    A = A + "(CASE WHEN AS1P*S1P*XTI>0 THEN (AS1P/(XTI*S1P)-1)*100 ELSE 0  END ) AS [M.UP],(CASE WHEN AS1P*S1P*XTI>0 THEN (1-(XTI*S1P)/AS1P)*100 ELSE 0  END ) AS [MG],AS1X AS [ΑΞ.ΑΓΟΡ],AS1P AS [ΑΞ.ΠΩΛ],(case WHEN AS1P>0 THEN AS1P-S1P*EID.XTI ELSE 0 END) AS [KEΡΔΟΣ],KERD1 "
Else
    A = A + "(CASE WHEN AS1P*S1P*XTI>0 THEN (AS1P/(XTI*S1P)-1)*100 ELSE 0  END ) AS [M.UP],(CASE WHEN AS1P*S1P*XTI>0 THEN (1-(XTI*S1P)/AS1P)*100 ELSE 0  END ) AS [MG],AS1X AS [ΑΞ.ΑΓΟΡ],AS1P AS [ΑΞ.ΠΩΛ],(case WHEN AS1P>0 THEN AS1P-S1P*EID.XTI ELSE 0 END) AS [KEΡΔΟΣ] "
End If



A = A + " INTO DOKEGGT3 FROM EID  LEFT JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD  "
'  ,(AS2P+AS1P)-(S2P+S1P)*EID.XTI AS [ΚΕΡΔΟΣ]



If Len(Fsql) > 1 Then
     A = A + " where " + Fsql
End If

If ORDERCOMBO.Text = ORDERCOMBO.List(0) Then
  A = A + " ORDER BY EID.ONO"
End If

If ORDERCOMBO.Text = ORDERCOMBO.List(1) Then
  A = A + " ORDER BY KOD"
End If

If ORDERCOMBO.Text = ORDERCOMBO.List(2) Then
  A = A + " ORDER BY EID.ERG"
End If

If ORDERCOMBO.Text = ORDERCOMBO.List(3) Then
  A = A + " ORDER BY AS1P-S1P*EID.XTI "  ' [KEPΔΟΣ]"
End If


' Exit Sub

On Error Resume Next
Gdb.Execute "DROP TABLE DOKEGGT3"  'SOUMES ΑΠΟΘΗΚΗΣ 1
On Error GoTo LATOS
Gdb.Execute A, d
Me.MousePointer = vbNormal

If d = 0 Then
   MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT3"
   Exit Sub
End If

Dim X As String
Dim ANS2

ANS2 = MsgBox("ΜΟΝΟ ΑΥΤΑ ΠΟΥ ΚΙΝΗΘΗΚΑΝ", vbYesNo)
If ANS2 = vbYes Then
    X = "SELECT * FROM DOKEGGT3 WHERE [ΑΓΟΡΕΣ]>0 OR [ΠΩΛΗΣΕΙΣ]>0 OR [ΑΞ.ΠΩΛ]>0 "
Else
    X = "SELECT * FROM DOKEGGT3"
End If


DATA1.ConnectionString = gConnect
DATA1.RecordSource = X
DATA1.Refresh











If Check1 Then
  'typos2
  On Error GoTo 0
  print3_xar UCase(X), "000011111221111111111", "", 0
End If


'If Check2 Then
'  Print4_Excel UCase(x), "00001111111111111111"
'End If

Exit Sub


LATOS:
MsgBox Err.Description
Resume Next





End Sub

Private Sub ypol_kerd()
'ypologismos kerdoys
' ' εαν δεν βρηκε τιμή αγοράς βάλε το 80% της τιμής πώλησης

'************************ ???d?? ************************************
Dim arx_arit, tel_arit, BUFF, kod_tim, ono_par, p1, progr, k
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



arx_arit = 0: tel_arit = 0

 ' sele TIM
  'SET index to ARTHME



Dim polhs As String
Dim PolXParoxh As String
Dim Paroxh As String
Dim AGOEPIS As String
Dim polepis As String, pol As String, ago As String, APIS
Dim A
 A = Get2_AJ(polhs, polepis, ago, AGOEPIS, PolXParoxh, Paroxh)






'On Error GoTo LATOS








Dim EGGTIM As New ADODB.Recordset
Dim ANS3 As Integer

ANS3 = MsgBox("Μόνο τα είδη που δεν έχουν κέρδος;", vbYesNo)

If ANS3 = vbYes Then
  EGGTIM.Open "SELECT * FROM EGGTIM WHERE (KERDOS IS NULL OR KERDOS=0) AND LEFT(ATIM,1) IN (" + polhs + ") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' ORDER BY HME", Gdb, adOpenDynamic, adLockOptimistic
Else
  EGGTIM.Open "SELECT * FROM EGGTIM WHERE LEFT(ATIM,1) IN (" + polhs + ") AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' ORDER BY HME", Gdb, adOpenDynamic, adLockOptimistic
End If


Dim R As New ADODB.Recordset
Dim EID As New ADODB.Recordset





EGGTIM.MoveFirst

Dim Anyparktoi(100), COUNT_aNYPARKTOI As Long

COUNT_aNYPARKTOI = 0

Dim FF As New UDialog

FF.OKButton.Visible = False
FF.List1.Visible = False

FF.CancelButton.Caption = "ΔΙΑΚΟΠΗ"
FF.CancelButton.Top = 120
FF.CancelButton.Left = 120
FF.CancelButton.width = 2895
FF.CancelButton.height = 495


FF.width = 3210
FF.height = 810
FF.Caption = "ΔΙΑΔΙΚΑΣΙΑ ΥΠΟΛΟΓΙΣΜΟΥ"







FF.Top = MDIForm1.Top + (MDIForm1.height) / 2 ' Command2.Top
FF.Left = MDIForm1.Left + (MDIForm1.width) / 2 ' Command2.Left









FF.SHOW




Dim b, d
Do While Not EGGTIM.EOF
  
  If FF.CancelButton.Enabled = False Then
      Exit Do
  End If
  
  
  
     
  'ΒΡΙΣΚΩ ΤΗΝ ΤΕΛΕΥΤΑΙΑ ΤΙΜΗ ΑΓΟΡΑΣ
  A = EGGTIM("KODE")
  d = EGGTIM("HME")
  
  
  
  
  k = 0
  R.Open "SELECT * FROM EGGTIM WHERE KODE='" + A + "' AND LEFT(ATIM,1) IN ('G','g','λ','Ξ') AND HME<='" + Format(d, "MM/DD/YYYY") + "' ORDER BY HME DESC", Gdb, adOpenDynamic, adLockOptimistic
  If R.EOF Then
    'PAIRNO THN TIMH APO TO EID
     EID.Open "SELECT XTI FROM EID WHERE KOD='" + A + "'", Gdb, adOpenForwardOnly, adLockReadOnly
     If EID.EOF Then
        COUNT_aNYPARKTOI = COUNT_aNYPARKTOI + 1
        If COUNT_aNYPARKTOI < 100 Then Anyparktoi(COUNT_aNYPARKTOI) = A
        ' MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ ΣΤΑ ΕΙΔΗ Ο ΚΩΔΙΚΟΣ " + A
     Else
       b = EID(0)
     End If
     EID.Close
   Else
      If IsNull(R("EKPT")) Then
        b = R("TIMM")
      Else
         b = R("TIMM") * (1 - R("EKPT") / 100)
      End If
      Gdb.Execute "UPDATE EID SET HPAR='" + Format(R("HME"), "MM/DD/YYYY") + "' where KOD='" + A + "'", k
      
      
   End If
   If A = "230-73-VRA38X2" Then
      arx_arit = 1
   End If

  If IsNull(b) Or b = 0 Then
    'PAIRNO THN TIMH APO TO EID
     EID.Open "SELECT XTI FROM EID WHERE KOD='" + A + "'", Gdb, adOpenForwardOnly, adLockReadOnly
     If EID.EOF Then
        'MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ ΣΤΑ ΕΙΔΗ Ο ΚΩΔΙΚΟΣ " + A
        
         COUNT_aNYPARKTOI = COUNT_aNYPARKTOI + 1
        If COUNT_aNYPARKTOI < 100 Then Anyparktoi(COUNT_aNYPARKTOI) = A
        
     Else
       b = EID(0)
       
      
      ' εαν δεν βρηκε τιμή αγοράς βάλε το 80% της τιμής πώλησης
      If b = 0 Then
         'ΑΝ ΕΙΝΑΙ ΠΑΡΟΧΗ ΑΣΤΟ 0
         If InStr(Paroxh, Left(EGGTIM("atim"), 1)) > 0 Then
         Else
            b = 0.8 * EGGTIM("timm") * (1 - EGGTIM("EKPT") / 100)
         End If
      End If
       
       
     End If
     EID.Close
   End If



'
   R.Close
   DoEvents
   Me.Caption = Format(d, "dd/mm/yyyy hh:mm") + " " + Format(k, "####")
   
   If IsNull(EGGTIM("EKPT")) Then
        EGGTIM("KERDOS") = (EGGTIM("TIMM") - b) * EGGTIM("POSO")
      Else
                       ' TIMH POLHSHS                            - AGORAS X POSOTHTA
         EGGTIM("KERDOS") = (EGGTIM("TIMM") * (1 - EGGTIM("EKPT") / 100) - b) * EGGTIM("POSO")
      End If
   EGGTIM.Update
   EGGTIM.MoveNext








Loop
EGGTIM.Close


Unload FF




If COUNT_aNYPARKTOI > 0 Then
   Dim ff2 As New UDialog
   Dim mM As Long
   If COUNT_aNYPARKTOI > 100 Then mM = 100 Else mM = COUNT_aNYPARKTOI
   For k = 1 To mM
     On Error Resume Next
     ff2.List1.AddItem Anyparktoi(k)
   Next
   ff2.Caption = "ανυπαρκτοι κωδικοί"
   ff2.SHOW 1
   
End If


Exit Sub

LATOS:
Resume Next





   
End Sub

Private Sub Command9_Click()
'ektypvsh kerdvn
'ΕΚΤΥΠΩΣΗ ΣΤΟ ΣΥΝΟΛΟ ΤΩΝ ΑΠΟΘΗΚΩΝ

Dim A As String
Dim b As String
Dim d As Long

' UPDATE_YPOLOIPA
UPDATE_YPOLOIPA3

Fsql = Text1.Text


Dim ANS

If F_epan = 1 Then

   ANS = MsgBox("Να υπολογίσω το κέρδος ξανά;", vbYesNo)
  If ANS = vbYes Then
    ypol_kerd
  End If
End If






Dim polhs As String

Dim PolXParoxh As String
Dim Paroxh As String

Dim AGOEPIS As String
Dim polepis As String, pol As String, ago As String, APIS

A = Get2_AJ(polhs, polepis, ago, AGOEPIS, PolXParoxh, Paroxh)

If Combo1.ListIndex = 0 Then
   polhs = PolXParoxh
End If
If Combo1.ListIndex = 2 Then
   polhs = Paroxh
End If


'polhs = "'L','t','T','y','Y','υ'"


Dim apothkes As String
If Val(TEXT3.Text) = 99 Then
    apothkes = ""
Else
    apothkes = " and APOT=" + TEXT3.Text
End If





'
' On Error Resume Next    AS1P>0

On Error GoTo 0
'-------------------------------- NEO ARXH ===============================

Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1

A = "SELECT KODE," _
& " SUM( CASE   WHEN LEFT(ATIM,1)='λ'  THEN XRE ELSE 0  END ) AS S0X," _
& " SUM( CASE   WHEN LEFT(ATIM,1) IN ('G','g')  THEN XRE ELSE 0  END ) AS S1X," _
& " SUM( CASE WHEN PIS IS NULL  OR LEFT(ATIM,1) NOT IN (" + polhs + ")  THEN 0 ELSE PIS END) AS S1P," _
& " SUM( CASE WHEN XRE*TIMM IS NULL THEN 0 ELSE round(XRE*TIMM*(1- (CASE WHEN EKPT IS NULL THEN 0 ELSE EKPT END)/100),2) END  ) AS AS1X," _
& " SUM( CASE WHEN POSO*TIMM IS NULL OR LEFT(ATIM,1) NOT IN (" + polhs + ")  THEN 0 ELSE round(POSO*TIMM*(1-(CASE WHEN EKPT IS NULL THEN 0 ELSE EKPT END)/100),2) END ) AS AS1P, " _
& " SUM( CASE WHEN KERDOS IS NULL OR LEFT(ATIM,1) NOT IN (" + polhs + ")  THEN 0 ELSE round(KERDOS,2) END ) AS KERD1 INTO DOKEGGT1 "
A = A + " FROM EGGTIM WHERE  left(ATIM,1) NOT IN ('σ','Σ') AND left(ATIM,1) IN (" + polhs + ") AND  " _
+ " HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' " + apothkes + "   GROUP BY KODE"

' ' & " SUM(XRE)-SUM(PIS) AS YP," _


 Gdb.Execute A, d

If d = 0 Then
   MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"
   Exit Sub
End If

'
A = "SELECT LEFT(KOD,14)AS [ΚΩΔ],LEFT(EID.PROM,10) AS [ΠΡΟΜΗΘ],LEFT(EID.ONO,35) AS [ΠΕΡΙΓΡΑΦΗ],"
A = A + "S0X AS [ΑΠΟΓΡ], S1X AS [ΑΓΟΡΕΣ],S1P AS [ΠΩΛΗΣΕΙΣ],S0X+S1X-S1P AS [ΥΠΟΛ],"

A = A + "(CASE WHEN S1P>0 THEN AS1P/S1P ELSE 0  END ) AS [MEΣΤΙΜ.ΠΩΛ],XTI AS [TIM.AΓ],"

If F_epan = 1 Then
    A = A + "(CASE WHEN AS1P*S1P*XTI>0 THEN round((AS1P/(XTI*S1P)-1)*100,2) ELSE 0  END ) AS [M.UP],(CASE WHEN AS1P*S1P*XTI>0 THEN round((1-(XTI*S1P)/AS1P)*100,2) ELSE 0  END ) AS [MG],AS1X AS [ΑΞ.ΑΓΟΡ],AS1P AS [ΑΞ.ΠΩΛ],(case WHEN AS1P>0 THEN round(AS1P-S1P*EID.XTI,2) ELSE 0 END) AS [KEΡΔΟΣ],KERD1 "
Else
    A = A + "(CASE WHEN AS1P*S1P*XTI>0 THEN (AS1P/(XTI*S1P)-1)*100 ELSE 0  END ) AS [M.UP],(CASE WHEN AS1P*S1P*XTI>0 THEN (1-(XTI*S1P)/AS1P)*100 ELSE 0  END ) AS [MG],AS1X AS [ΑΞ.ΑΓΟΡ],AS1P AS [ΑΞ.ΠΩΛ],(case WHEN AS1P>0 THEN AS1P-S1P*EID.XTI ELSE 0 END) AS [KEΡΔΟΣ] "
End If



A = A + " INTO DOKEGGT3 FROM EID  LEFT JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD  "
'  ,(AS2P+AS1P)-(S2P+S1P)*EID.XTI AS [ΚΕΡΔΟΣ]



If Len(Fsql) > 1 Then
     A = A + " where " + Fsql
End If

If ORDERCOMBO.Text = ORDERCOMBO.List(0) Then
  A = A + " ORDER BY EID.ONO"
End If

If ORDERCOMBO.Text = ORDERCOMBO.List(1) Then
  A = A + " ORDER BY KOD"
End If

If ORDERCOMBO.Text = ORDERCOMBO.List(2) Then
  A = A + " ORDER BY EID.ERG"
End If

If ORDERCOMBO.Text = ORDERCOMBO.List(3) Then
  A = A + " ORDER BY AS1P-S1P*EID.XTI "  ' [KEPΔΟΣ]"
End If


' Exit Sub

On Error Resume Next
Gdb.Execute "DROP TABLE DOKEGGT3"  'SOUMES ΑΠΟΘΗΚΗΣ 1
On Error GoTo LATOS
Gdb.Execute A, d


If d = 0 Then
   MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT3"
   Exit Sub
End If

Dim X As String
Dim ANS2

ANS2 = MsgBox("ΜΟΝΟ ΑΥΤΑ ΠΟΥ ΚΙΝΗΘΗΚΑΝ", vbYesNo)
If ANS2 = vbYes Then
    X = "SELECT * FROM DOKEGGT3 WHERE [ΑΓΟΡΕΣ]>0 OR [ΠΩΛΗΣΕΙΣ]>0 OR [ΑΞ.ΠΩΛ]>0 "
Else
    X = "SELECT * FROM DOKEGGT3"
End If


DATA1.ConnectionString = gConnect
DATA1.RecordSource = X
DATA1.Refresh











If Check1 Then
  'typos2
  On Error GoTo 0
  print3_xar UCase(X), "000011111111111111111", "", 0
End If


'If Check2 Then
'  Print4_Excel UCase(x), "00001111111111111111"
'End If

Exit Sub


LATOS:
MsgBox Err.Description
Resume Next

End Sub

Private Sub File1_Click()
 Dim A
  A = 0

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


Text4.Text = ss

rCheck.value = vbChecked




End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then Command2_Click
  
End Sub

'**
' basic form-load
'@rem          remarks
'**
'
'@see
'@rem
Private Sub Form_Load()
   
   Combo1.Text = Combo1.List(0)
   Combo2.Text = Combo2.List(1)
   F_RUN = 0
   Fsql = ""
   FIsHidden = False
Me.KeyPreview = True

F_epan = Val(FindParametroi("APOT3", "F_EPAN", "1", "EΠΑΝΥΠΟΛΟΓΙΣΜΟΣ ΠΑΡ/ΚΩΝ=1 ΟΧΙ=0"))

   'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
   ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
    'SkinFramework1.ApplyWindow Me.hWnd
    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
Dim k As Integer
For k = 0 To 35
    SEL(k) = -1
Next

 
 Me.Picture = LoadPicture(gPicture)


Cd1.InitDir = "c:\mercvb\queries\EID"
'On Error Resume Next
' File1.Path = LCase("c:\mercvb\queries\EID")






If Len(UCase(Dir("c:\mercvb\queries\eid", vbDirectory))) = 0 Then
    MkDir "c:\mercvb\queries\eid"
End If




File1.Path = "c:\mercvb\queries\eid"






File1.Refresh

 F_LHM1 = FindParametroi("APOT2", "F_lhm1", "Ημερ1", "Ετικέτα ημερ.1")
 F_LHM2 = FindParametroi("APOT2", "F_lhm2", "Ημερ2", "Ετικέτα ημερ.2")
 f_lHM3 = FindParametroi("APOT2", "F_lhm3", "Ημερ3", "Ετικέτα ημερ.3")
 
 f_lab1 = FindParametroi("APOT2", "F_lab1", "Κειμ1", "Ετικέτα κειμ.1")
 f_lab2 = FindParametroi("APOT2", "F_lab2", "Κειμ2", "Ετικέτα κειμ.2")
 f_lab3 = FindParametroi("APOT2", "F_lab3", "Κειμ3", "Ετικέτα κειμ.3")
 f_lab4 = FindParametroi("APOT2", "F_lab4", "Κειμ4", "Ετικέτα κειμ.4")
 f_lab5 = FindParametroi("APOT2", "F_lab5", "Κειμ5", "Ετικέτα κειμ.5")
 f_lab6 = FindParametroi("APOT2", "F_lab6", "Κειμ6", "Ετικέτα κειμ.6")
 
 f_num1 = FindParametroi("APOT2", "F_num1", "Αριθ1", "Ετικέτα κειμ.1")
 f_num2 = FindParametroi("APOT2", "F_num2", "Αριθ2", "Ετικέτα κειμ.2")
 f_num3 = FindParametroi("APOT2", "F_num3", "Αριθ3", "Ετικέτα κειμ.3")

'ΣΤΗΛΕΣ ΠΙΝΑΚΑ
' 0=sql string  1=sum  2=alignment  3=width  4=listbox
Dim R As Integer
'A = "SELECT KOD,LEFT(EID.PROM,20) AS [ΠΡΟΜΗΘ],EID.ERG AS [BARCODE],LEFT(EID.ONO,35) AS [ΠΕΡΙΓΡΑΦΗ],"
'A = A + " S1X AS [ΑΓ.ΠΟΣ1],S1P AS [ΠΩΛ.ΠΟΣ1],POS01 AS [ΥΠΟΛ.1],"
'A = A + " S2X AS [ΑΓ.ΑΠ2] ,S2P AS [ΠΩΛ.ΠΟΣ2],POS02 AS [ΥΠ.ΑΠ2],"
'If NAPOT > 2 Then 'FAKIS
'   A = A + " S3X AS [ΑΓ.ΠΟΣ3],S3P AS [ΠΩΛ.ΠΟΣ3],POS03 AS [ΥΠΟΛ.3],"
'   A = A + " S4X AS [ΑΓ.ΑΠ4] ,S4P AS [ΠΩΛ.ΠΟΣ4],POS04 AS [ΥΠ.ΑΠ4],"
'End If
'A = A + " AS1X AS [ΑΞ.ΑΓ1],AS1P AS [ΑΞ.ΠΩΛ.ΑΠ1],AS2X AS [ΑΞ.ΑΓ2],AS2P AS [ΑΞ.ΠΩΛ.ΑΠ2] "
'If NAPOT > 2 Then ' FAKIS
'   A = A + " ,AS3X AS [ΑΞ.ΑΓ3],AS3P AS [ΑΞ.ΠΩΛ.ΑΠ3],AS4X AS [ΑΞ.ΑΓ4],AS4P AS [ΑΞ.ΠΩΛ.ΑΠ4] "
'End If
'A = A + " INTO DOKEGGT3 FROM EID  LEFT JOIN DOKEGGT1  ON DOKEGGT1.KODE=EID.KOD"
R = -1
R = R + 1: f_list1(R, 0) = ",S1X AS [ΑΓ.ΠΟΣ1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αγορές Αποθ1"
R = R + 1: f_list1(R, 0) = ",S1P AS [ΠΩΛ.ΠΟΣ1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Πωλήσ.Αποθ1"
R = R + 1: f_list1(R, 0) = ", AS1X AS [ΑΞ.ΑΓ1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αξία Αγορών Αποθ1"
R = R + 1: f_list1(R, 0) = ",AS1P AS [ΑΞ.ΠΩΛ.ΑΠ1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αξία Πωλήσεων Αποθ1"

R = R + 1: f_list1(R, 0) = ",S2X AS [ΑΓ.ΠΟΣ2]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αγορές Αποθ2"
R = R + 1: f_list1(R, 0) = ",S2P AS [ΠΩΛ.ΠΟΣ2]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Πωλήσ.Αποθ2"
R = R + 1: f_list1(R, 0) = ", AS2X AS [ΑΞ.ΑΓ2]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αξία Αγορών Αποθ2"
R = R + 1: f_list1(R, 0) = ",AS2P AS [ΑΞ.ΠΩΛ.ΑΠ2]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αξία Πωλήσεων Αποθ2"

R = R + 1: f_list1(R, 0) = ",S3X AS [ΑΓ.ΠΟΣ3]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αγορές Αποθ3"
R = R + 1: f_list1(R, 0) = ",S3P AS [ΠΩΛ.ΠΟΣ3]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Πωλήσ.Αποθ3"
R = R + 1: f_list1(R, 0) = ", AS3X AS [ΑΞ.ΑΓ3]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αξία Αγορών Αποθ3"
R = R + 1: f_list1(R, 0) = ",AS3P AS [ΑΞ.ΠΩΛ.ΑΠ3]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αξία Πωλήσεων Αποθ3"

R = R + 1: f_list1(R, 0) = ",S4X AS [ΑΓ.ΠΟΣ4]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αγορές Αποθ4"
R = R + 1: f_list1(R, 0) = ",S4P AS [ΠΩΛ.ΠΟΣ4]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Πωλήσ.Αποθ4"
R = R + 1: f_list1(R, 0) = ", AS4X AS [ΑΞ.ΑΓ4]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αξία Αγορών Αποθ4"
R = R + 1: f_list1(R, 0) = ",AS4P AS [ΑΞ.ΠΩΛ.ΑΠ4]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Αξία Πωλήσεων Αποθ4"





R = R + 1: f_list1(R, 0) = ",LEFT(EID.MON,5) AS [Mον.Μετ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Μον.Μέτρ."
R = R + 1: f_list1(R, 0) = ",str(EID.LTI,10,2) as [Χονδ.Τιμή]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "Τιμή Xoνδρικής"
R = R + 1: f_list1(R, 0) = ",str(EID.LTI5,10,2) as [Λιαν.Τιμή]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "Τιμή Λιανικής"
R = R + 1: f_list1(R, 0) = ",EID.POS AS [Υπόλοιπο]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Υπόλοιπο"
R = R + 1: f_list1(R, 0) = ",EID.POS01 as [Υπ.Απ1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1300: f_list1(R, 4) = "Υπολ.Απ.1"
R = R + 1: f_list1(R, 0) = ",EID.POS02 as [Υπ.Απ2]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "Υπολ.Απ.2"
R = R + 1: f_list1(R, 0) = ",EID.FPA as [Kατ.ΦΠΑ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Κατ.ΦΠΑ"
R = R + 1: f_list1(R, 0) = ",str(EID.XTI,10,2) AS [Tιμή Αγ.]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "Τιμή αγοράς"
R = R + 1: f_list1(R, 0) = ",EID.XTI*EID.POS01 AS [Αξ.Απογρ.απ.1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Αξία Απογραφής απ1"
R = R + 1: f_list1(R, 0) = ",EID.XTI*EID.POS02 AS [Αξ.Απογρ.απ.2]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Αξία Απογραφής απ2"

R = R + 1: f_list1(R, 0) = ",EID.XTI*EID.POS AS [Αξ.Απογρ.]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Αξία Απογραφής"
R = R + 1: f_list1(R, 0) = ",EID.PONTOI AS [Πόντοι]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Πόντοι"
R = R + 1: f_list1(R, 0) = ",EID.ERG as [Βοηθ.Κωδ.] ": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Βarcode"
R = R + 1: f_list1(R, 0) = ",EID.PROM AS [Προμηθευτής]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Προμηθευτής"
R = R + 1: f_list1(R, 0) = ",EID.KODLOG as [Κατηγορία]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Κατηγορία"
R = R + 1: f_list1(R, 0) = ",EID.HPAR as [Ημ.Τελ.Αγοράς]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Ημερ.Τελ.Αγοράς"
R = R + 1: f_list1(R, 0) = ",EID.PAR as [Παραγγελ.Ποσ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Παραγγ.Ποσότητα"
R = R + 1: f_list1(R, 0) = ",EID.SPA as [ΣτάθμηΠαραγγ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Στάθμη Παραγγελίας"
R = R + 1: f_list1(R, 0) = ",EID.EPIUYP as [Επιθ.Υπόλ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Επιθυμητό Υπόλοιπο"
R = R + 1: f_list1(R, 0) = ",EID.SYSKMAX as [Συσκευασία]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Συσκευασία"
R = R + 1: f_list1(R, 0) = ",EID.SYSKMIN as [ΥποΣυσκευασία]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "YποΣυσκευασία"
R = R + 1: f_list1(R, 0) = ",EID.POS03 as [Υπ.Απ3]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Υπολ.Απ3"
R = R + 1: f_list1(R, 0) = ",EID.POS04 as [Υπ.Απ4]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Υπολ.Απ4"
R = R + 1: f_list1(R, 0) = ", CONVERT(CHAR(12),EID.HM1,3) AS [" + F_LHM1 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = F_LHM1
R = R + 1: f_list1(R, 0) = ", CONVERT(CHAR(12),EID.HM2,3) AS [" + F_LHM2 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = F_LHM2
R = R + 1: f_list1(R, 0) = ", CONVERT(CHAR(12),EID.HM3,3) AS [" + f_lHM3 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lHM3
R = R + 1: f_list1(R, 0) = ", EID.CH1 AS [" + f_lab1 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab1
R = R + 1: f_list1(R, 0) = ", EID.CH2 AS [" + f_lab2 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab2
R = R + 1: f_list1(R, 0) = ", EID.CH3 AS [" + f_lab3 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab3
R = R + 1: f_list1(R, 0) = ", EID.CH4 AS [" + f_lab4 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab4
R = R + 1: f_list1(R, 0) = ", EID.CH5 AS [" + f_lab5 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab5
R = R + 1: f_list1(R, 0) = ", EID.CH6 AS [" + f_lab6 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab6
R = R + 1: f_list1(R, 0) = ", str(EID.NUM1,10,2) AS [" + f_num1 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_num1
R = R + 1: f_list1(R, 0) = ", str(EID.NUM2,10,2) AS [" + f_num2 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_num2
R = R + 1: f_list1(R, 0) = ", str(EID.NUM3,10,2) AS [" + f_num3 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_num3

























'R = 0: f_list1(R, 0) = ",LEFT(MON,5) AS [Mον.Μετ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "Μον.Μέτρ."
'R = R + 1: f_list1(R, 0) = ",str(LTI,10,2) as [Χονδ.Τιμή]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "Τιμή Xoνδρικής"
'R = R + 1: f_list1(R, 0) = ",str(LTI5,10,2) as [Λιαν.Τιμή]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "Τιμή Λιανικής"
'R = R + 1: f_list1(R, 0) = ",POS AS [Υπόλοιπο]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Υπόλοιπο"
'R = R + 1: f_list1(R, 0) = ",POS01 as [Υπ.Απ1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1300: f_list1(R, 4) = "Υπολ.Απ.1"
'R = R + 1: f_list1(R, 0) = ",POS02 as [Υπ.Απ2]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "Υπολ.Απ.2"
'R = R + 1: f_list1(R, 0) = ",FPA as [Kατ.ΦΠΑ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Κατ.ΦΠΑ"
'R = R + 1: f_list1(R, 0) = ",str(XTI,10,2) AS [Tιμή Αγ.]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "Τιμή αγοράς"
'R = R + 1: f_list1(R, 0) = ",XTI*POS01 AS [Αξ.Απογρ.απ.1]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Αξία Απογραφής απ1"
'R = R + 1: f_list1(R, 0) = ",XTI*POS02 AS [Αξ.Απογρ.απ.2]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Αξία Απογραφής απ2"

'R = R + 1: f_list1(R, 0) = ",XTI*POS AS [Αξ.Απογρ.]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Αξία Απογραφής"
'R = R + 1: f_list1(R, 0) = ",PONTOI AS [Πόντοι]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Πόντοι"
'R = R + 1: f_list1(R, 0) = ",(CASE S1X WHEN NULL THEN 0 ELSE S1X END) AS [Εισ1.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ1.Εισαγ.Περιόδου"
'R = R + 1: f_list1(R, 0) = ",(CASE AS1X WHEN NULL THEN 0 ELSE AS1X END)  AS [Αξ1.Αγ.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ1.Αξ.Εισαγ.Περιόδου"
'R = R + 1: f_list1(R, 0) = ",(CASE S2X WHEN NULL THEN 0 ELSE S2X END)  AS [Εισ2.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ2.Εισαγ.Περιόδου"
'R = R + 1: f_list1(R, 0) = ",(CASE AS2X WHEN NULL THEN 0 ELSE AS2X END) AS [Αξ2.Αγ.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ2.Αξ.Εισαγ.Περιόδου"
'R = R + 1: f_list1(R, 0) = ",(CASE S1P WHEN NULL THEN 0 ELSE S1P END)  AS [Εξ1.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ1.Εξ.Περιόδου"
'R = R + 1: f_list1(R, 0) = ",(CASE AS1P WHEN NULL THEN 0 ELSE AS1P END)  AS [Αξ2.Αγ.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ1.Αξ.Εξ.Περιόδου"
'R = R + 1: f_list1(R, 0) = ",(CASE S2P WHEN NULL THEN 0 ELSE S2P END)  AS [Εξ2.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ2.Εξ.Περιόδου"
'R = R + 1: f_list1(R, 0) = ",(CASE AS2P WHEN NULL THEN 0 ELSE AS2P END)  AS [Αξ2.Εξ.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ2.Αξ.Εξ.Περιόδου"
'R = R + 1: f_list1(R, 0) = ",ERG as [Βοηθ.Κωδ.] ": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Βarcode"
'R = R + 1: f_list1(R, 0) = ",PROM AS [Προμηθευτής]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Προμηθευτής"
'R = R + 1: f_list1(R, 0) = ",KODLOG as [Κατηγορία]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Κατηγορία"
'R = R + 1: f_list1(R, 0) = ",HPAR as [Ημ.Τελ.Αγοράς]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Ημερ.Τελ.Αγοράς"
'R = R + 1: f_list1(R, 0) = ",PAR as [Παραγγελ.Ποσ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Παραγγ.Ποσότητα"
'R = R + 1: f_list1(R, 0) = ",SPA as [ΣτάθμηΠαραγγ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Στάθμη Παραγγελίας"
'R = R + 1: f_list1(R, 0) = ",EPIUYP as [Επιθ.Υπόλ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Επιθυμητό Υπόλοιπο"
'R = R + 1: f_list1(R, 0) = ",SYSKMAX as [Συσκευασία]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Συσκευασία"
'R = R + 1: f_list1(R, 0) = ",SYSKMIN as [ΥποΣυσκευασία]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "YποΣυσκευασία"
'R = R + 1: f_list1(R, 0) = ",(CASE S3X WHEN NULL THEN 0 ELSE S3X END)  AS [Εισ3.Περιόδ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ3.Εισαγ.Περιόδου"
'R = R + 1: f_list1(R, 0) = ",(CASE AS3X WHEN NULL THEN 0 ELSE AS3X END) AS [Αξ3.Αγ.Περιόδ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ3.Αξ.Εισαγ.Περιόδου"
'R = R + 1: f_list1(R, 0) = ",(CASE S3P WHEN NULL THEN 0 ELSE S3P END)  AS [Εξ3.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ3.Εξ.Περιόδου"
'R = R + 1: f_list1(R, 0) = ",(CASE AS3P WHEN NULL THEN 0 ELSE AS3P END)  AS [Αξ3.Αγ.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ3.Αξ.Εξ.Περιόδου"
'R = R + 1: f_list1(R, 0) = ",(CASE S4X WHEN NULL THEN 0 ELSE S4X END)  AS [Εισ4.Περιόδ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ4.Εισαγ.Περιόδου"
'R = R + 1: f_list1(R, 0) = ",(CASE AS4X WHEN NULL THEN 0 ELSE AS4X END) AS [Αξ4.Αγ.Περιόδ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ4.Αξ.Εισαγ.Περιόδου"
'R = R + 1: f_list1(R, 0) = ",(CASE S4P WHEN NULL THEN 0 ELSE S4P END)  AS [Εξ4.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ4.Εξ.Περιόδου"
'R = R + 1: f_list1(R, 0) = ",(CASE AS4P WHEN NULL THEN 0 ELSE AS4P END)  AS [Αξ4.Αγ.Περιόδ]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Απ4.Αξ.Εξ.Περιόδου"
'R = R + 1: f_list1(R, 0) = ",POS03 as [Υπ.Απ3]": f_list1(R, 1) = 1: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Υπολ.Απ3"
'R = R + 1: f_list1(R, 0) = ",POS04 as [Υπ.Απ4]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Υπολ.Απ4"
'R = R + 1: f_list1(R, 0) = ", CONVERT(CHAR(12),HM1,3) AS [" + F_LHM1 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = F_LHM1
'R = R + 1: f_list1(R, 0) = ", CONVERT(CHAR(12),HM2,3) AS [" + F_LHM2 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = F_LHM2
'R = R + 1: f_list1(R, 0) = ", CONVERT(CHAR(12),HM3,3) AS [" + f_lHM3 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lHM3
'R = R + 1: f_list1(R, 0) = ", CH1 AS [" + f_lab1 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab1
'R = R + 1: f_list1(R, 0) = ", CH2 AS [" + f_lab2 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab2
'R = R + 1: f_list1(R, 0) = ", CH3 AS [" + f_lab3 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab3
'R = R + 1: f_list1(R, 0) = ", CH4 AS [" + f_lab4 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab4
'R = R + 1: f_list1(R, 0) = ", CH5 AS [" + f_lab5 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab5
'R = R + 1: f_list1(R, 0) = ", CH6 AS [" + f_lab6 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab6
'R = R + 1: f_list1(R, 0) = ", str(NUM1,10,2) AS [" + f_num1 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_num1
'R = R + 1: f_list1(R, 0) = ", str(NUM2,10,2) AS [" + f_num2 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_num2
'R = R + 1: f_list1(R, 0) = ", str(NUM3,10,2) AS [" + f_num3 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_num3


For k = 0 To UBound(f_list1)
   If Len(f_list1(k, 4)) > 0 Then
       List1.AddItem f_list1(k, 4)
   End If
Next



































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





sqlcommand.ConnectionString = gConnect
'sqlcommand.RecordSource = "SELECT  *FROM SQLCOMMAND"
sqlcommand.RecordSource = "SELECT TOP 10 *FROM EID"
sqlcommand.Refresh





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
   
apo = gEnarjh
eos = gLhjh

   
   
   
End Sub

Private Sub Form_DblClick()
   PARAMETROI.PARAM.Caption = "APOT3"
   PARAMETROI.SHOW 1
    'F_TAB = Val(FindParametroi("APOT2", "F_TAB", "3", "Σε ποιό TAB είναι σταματημένο"))
End Sub

Private Sub OLE1_Updated(Code As Integer)


End Sub





'--------------------------------------------------------------------------------


Private Sub update_DBF_from_SQL(arxeio As String, ARXEIOdbf As String)

Dim DBF As Database
Dim sql As New ADODB.Connection
Dim rDBF As Recordset
Dim rSQL As New ADODB.Recordset
Dim conDBF As String
Dim conSQL As String
Dim k As Long
Dim Fname As String

Dim db As DAO.Database
Set db = OpenDatabase("C:\LAGEURO\", False, False, "dBase III;")

Set rDBF = db.OpenRecordset(ARXEIOdbf)





rSQL.Open arxeio, Gdb, adOpenDynamic, adLockOptimistic
' rDBF.Open "SELECT *FROM " + arxeio, DBF, adOpenDynamic, adLockOptimistic

rSQL.MoveFirst

Dim Z
Z = 0
On Error GoTo WRITEERROR ' Resume Next

Do While Not rSQL.EOF
  
  rDBF.AddNew
  
  For k = 0 To rDBF.FIELDS.Count - 1
      Fname = rDBF.FIELDS(k).Name ' p.x. FNAME=epo    R(0).NAME
      rDBF(Fname) = rSQL(Fname) ' rsqk("epo")=r(0)
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
'List2.Visible = True

'List2.AddItem Err.Description
Resume Next




End Sub




















'
Private Function getFilter() As String
    'Creates the SQL statement in adodc1.recordset.filter
    'and only filters text currently. It must be modified to filter other data types.
    
    Dim tmp As String
    Dim n As Integer
    For Each F_COL In F_COLS
        If Trim(F_COL.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If
           
           
          If DATA1.Recordset(F_COL.ColIndex).Type = 5 Then '
            If InStr(">< >= <=  = ", Left(F_COL.FilterText, 1)) > 0 And Len(F_COL.FilterText) > 1 And IsNumeric(Right(F_COL.FilterText, 1)) Then
                tmp = tmp & F_COL.DataField & F_COL.FilterText
            Else
             tmp = tmp & F_COL.DataField & " >  " & Val(F_COL.FilterText)
            End If
            
          Else
             tmp = tmp & F_COL.DataField & " LIKE '" & F_COL.FilterText & "*'"
          End If
            
        End If
    Next F_COL
                
    getFilter = tmp
End Function

Private Sub GR2_FilterChange()
'Private Sub DATAGRID1_FilterChange()

'Gets called when an action is performed on the filter bar

On Error GoTo ErrHandler

 

Set F_COLS = GR2.Columns

Dim C As Integer

C = GR2.Col

GR2.HoldFields

DATA1.Recordset.Filter = getFilter()

GR2.Col = C

GR2.EditActive = True

Exit Sub

 

ErrHandler:

    MsgBox Err.Source & ":" & vbCrLf & Err.Description

'    Call cmdClearFilter_Click

Resume Next


End Sub



Private Sub GR2_GroupColMove(ByVal Position As Integer, ByVal ColIndex As Integer, Cancel As Integer)



   Dim strSort As String

    Dim Col As TrueOleDBGrid80.Column

 

' Loop through GroupColumns collection and construct

' the sort string for the Sort property of the Recordset

    For Each Col In GR2.GroupColumns

        If strSort <> vbNullString Then

            strSort = strSort & ", "

        End If

        strSort = strSort & "[" & Col.DataField & "]"

    Next Col

 

    GR2.HoldFields

    DATA1.Recordset.Sort = strSort



End Sub

Private Sub List1_Click()
 Dim k As Integer
 Dim SELE
 SELE = List1.ListIndex
  If List1.Selected(SELE) Then
     SEL(SELE) = SELE
  Else
     SEL(SELE) = -1
  End If
  Exit Sub
 
 
 
 
 For k = 0 To 35
   If SELE = 0 Then
      If SEL(k) = SELE Then
        Exit For
      End If
      If SEL(k) = -1 Then
        SEL(k) = SELE
        Exit For
      End If
    
   Else
      If SEL(k) = SELE Then
        Exit For
      End If
      If SEL(k) = -1 Then  'αν ηταν αμαρκάριστο το δίνω τιμη>0   =list1.listindex
        SEL(k) = SELE
        Exit For
      End If
   End If
 Next
 
'SEL(1) = 2
'Dim KK
'For Each KK In SEL()


'Next


End Sub

Private Sub List1_DblClick()
   Dim k As Integer
   For k = 0 To List1.ListCount - 1
     List1.Selected(k) = True
   Next

End Sub
Sub UPDATE_YPOLOIPA3()
'into DOKEGGT1
'B = "CREATE VIEW dbo.[EIDT3]" _


Dim A As String
Dim b As String
Dim d As Long
'Dim krit As String

' ΒΡΙΣΚΩ ΜΙΑ ΦΟΡΑ ΤΑ ΥΠΟΛΟΙΠΑ ΑΝΑ ΑΠΟΘΗΚΗ

On Error Resume Next

Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1



A = "SELECT KODE," _
& " SUM(CASE APOT  WHEN 1  THEN XRE ELSE 0  END ) AS S1X," _
& " SUM(CASE APOT  WHEN 1  THEN PIS ELSE 0  END ) AS S1P," _
& " SUM(CASE APOT  WHEN 2  THEN XRE ELSE 0  END ) AS S2X," _
& " SUM(CASE APOT  WHEN 2  THEN PIS ELSE 0  END ) AS S2P," _
& " SUM(CASE APOT  WHEN 3  THEN XRE ELSE 0  END ) AS S3X," _
& " SUM(CASE APOT  WHEN 3  THEN PIS ELSE 0  END ) AS S3P," _
& " SUM(CASE APOT  WHEN 4  THEN XRE ELSE 0  END ) AS S4X," _
& " SUM(CASE APOT  WHEN 4  THEN PIS ELSE 0  END ) AS S4P  "
A = A + " INTO DOKEGGT1 FROM EGGTIM  "
A = A + "where HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME < '" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' GROUP BY KODE"

Dim TT As Long
TT = GetCurrentTime()
Gdb.Execute A, d

'MsgBox GetCurrentTime() - TT

If d = 0 Then
   MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"
   Exit Sub
End If

Gdb.Execute "UPDATE EID SET POS01=0,POS02=0,POS03=0,POS04=0,POS=0"

'b = "CREATE VIEW dbo.[EIDT3]" _
'& "AS " _
'& " SELECT     dbo.EID.*, (dbo.DOKEGGT1.S1X-dbo.DOKEGGT1.S1P) AS EXPR1, (dbo.DOKEGGT1.S2X-dbo.DOKEGGT1.S2P) AS EXPR2 " _
'& " FROM         dbo.DOKEGGT1 INNER JOIN " _
'& " dbo.EID ON dbo.DOKEGGT1.KODE = dbo.EID.KOD"
'Gdb.Execute b, d

Gdb.Execute "update EID SET" _
& " POS01=(CASE WHEN S1X IS NULL THEN 0 ELSE S1X END) - (CASE WHEN S1P IS NULL THEN 0 ELSE S1P END)," _
& " POS02=(CASE WHEN S2X IS NULL THEN 0 ELSE S2X END) - (CASE WHEN S2P IS NULL THEN 0 ELSE S2P END)," _
& " POS03=(CASE WHEN S3X IS NULL THEN 0 ELSE S3X END) - (CASE WHEN S3P IS NULL THEN 0 ELSE S3P END)," _
& " POS04=(CASE WHEN S4X IS NULL THEN 0 ELSE S4X END) - (CASE WHEN S4P IS NULL THEN 0 ELSE S4P END)" _
& " FROM EID INNER JOIN DOKEGGT1 ON EID.KOD=DOKEGGT1.KODE ", d


Gdb.Execute "UPDATE EID SET  POS = POS01 + POS02+POS03+POS04", d



d = 0
'Gdb.Execute "update EIDT3 SET POS01=EXPR1,POS02=EXPR2,POS=EXPR1+EXPR2 ", d


d = d + 1
End Sub


Function Get2_AJ(ByRef pol As String, polepis As String, ago As String, AGOEPIS As String, ByRef PolXParoxh As String, Paroxh As String) As Boolean

Dim db As Database
Dim R As New ADODB.Recordset, X As String

'If gConnect = "Access" Then
'   Set db = OpenDatabase(gDir, False, False)
'Else
'   Set db = OpenDatabase(gDir, False, False, gConnect)
'End If

R.Open "select POL,EIDOS,AJIA_APOU,SHM1 from PARASTAT", Gdb, adOpenDynamic, adLockOptimistic

  pol = ""
  PolXParoxh = ""
  
  Paroxh = ""
  
  R.MoveFirst
  
  Do While Not R.EOF
     If Not IsNull(R("eidos") + R("pol") + R("ajia_apou")) Then
       If R("pol") = "1" And R("ajia_apou") = "3" Then
          pol = pol + "'" + R("eidos") + "',"
       End If
       
       If R("pol") = "1" And R("ajia_apou") = "3" Then
          If Not IsNull(R("shm1")) And Val(Left(R("shm1"), 2)) <> 2 Then
              PolXParoxh = PolXParoxh + "'" + R("eidos") + "',"
          End If
       End If
       
       If R("pol") = "1" And R("ajia_apou") = "3" Then
          If Not IsNull(R("shm1")) And Val(Left(R("shm1"), 2)) = 2 Then
              Paroxh = Paroxh + "'" + R("eidos") + "',"
          End If
       End If
       
       
       
       
       
       If R("pol") = "1" And R("ajia_apou") = "4" Then
          polepis = polepis + "'" + R("eidos") + "',"
       End If
       
       If R("pol") = "2" And R("ajia_apou") = "1" Then
          ago = ago + "'" + R("eidos") + "',"
       End If
       If R("pol") = "2" And R("ajia_apou") = "2" Then
          AGOEPIS = AGOEPIS + "'" + R("eidos") + "',"
       End If
       
       
       
       
     End If
     R.MoveNext
  Loop
  pol = Left(pol, Len(pol) - 1)
  PolXParoxh = Left(PolXParoxh, Len(PolXParoxh) - 1)
  Paroxh = Left(Paroxh, Len(Paroxh) - 1)
  If Len(polepis) > 0 Then
     polepis = Left(polepis, Len(polepis) - 1)
  Else
     polepis = ""
  End If
     
  AGOEPIS = Left(AGOEPIS, Len(AGOEPIS) - 1)
  ago = Left(ago, Len(ago) - 1)
Get2_AJ = True



End Function



