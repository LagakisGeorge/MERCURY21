VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form par4 
   BackColor       =   &H00FFFF00&
   Caption         =   "Τζίροι-Παραστατικά "
   ClientHeight    =   10410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10410
   ScaleWidth      =   16905
   WindowState     =   2  'Maximized
   Begin VB.CheckBox akyromena 
      BackColor       =   &H00FF0000&
      Caption         =   "Να μην φαίνονται ακυρωτικα & ακυρωμενα"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   7920
      TabIndex        =   38
      Top             =   480
      Width           =   2985
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Προεπισκόπηση"
      Height          =   255
      Left            =   7920
      TabIndex        =   37
      Top             =   720
      Width           =   2295
   End
   Begin VB.ComboBox DATES 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "par4.frx":0000
      Left            =   4800
      List            =   "par4.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton restore_width 
      Caption         =   "Επαναφορά"
      Height          =   375
      Left            =   11400
      TabIndex        =   34
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin TrueOleDBGrid80.TDBGrid DBGrid1 
      Bindings        =   "par4.frx":008D
      Height          =   3874
      Left            =   117
      TabIndex        =   12
      Top             =   0
      Width           =   7683
      _ExtentX        =   13573
      _ExtentY        =   6826
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
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
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
      ColumnFooters   =   -1  'True
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTipsWidth   =   0
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=780,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=780,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=780,.italic=0"
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
   Begin VB.ComboBox metrhta 
      Height          =   288
      Left            =   6600
      TabIndex        =   32
      Top             =   3960
      Width           =   1332
   End
   Begin VB.TextBox chart 
      Height          =   288
      Left            =   13080
      TabIndex        =   30
      Text            =   "1"
      Top             =   4920
      Width           =   612
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3420
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9120
      Visible         =   0   'False
      Width           =   2970
   End
   Begin VB.TextBox macTEXT 
      Height          =   285
      Left            =   5475
      TabIndex        =   15
      Top             =   8655
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      Caption         =   "αποστολη σε ASCII"
      Height          =   285
      Left            =   5475
      TabIndex        =   14
      Top             =   8325
      Width           =   1680
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFLEXGRID2 
      Bindings        =   "par4.frx":00A1
      Height          =   3692
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   7475
      _ExtentX        =   13176
      _ExtentY        =   6509
      _Version        =   393216
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Κριτήρια"
      Height          =   375
      Left            =   11400
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
      Height          =   3614
      Left            =   117
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   15
      Cols            =   18
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   18
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2040
      Top             =   8400
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "data1"
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
      Height          =   288
      ItemData        =   "par4.frx":00B5
      Left            =   8040
      List            =   "par4.frx":00BF
      TabIndex        =   9
      Top             =   3960
      Width           =   2895
   End
   Begin MSChart20Lib.MSChart G 
      Height          =   3885
      Left            =   120
      OleObjectBlob   =   "par4.frx":00EC
      TabIndex        =   10
      Top             =   4320
      Width           =   12855
   End
   Begin VB.CommandButton exit 
      Caption         =   "Εξοδος"
      Height          =   495
      Left            =   9150
      TabIndex        =   4
      Top             =   8355
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc DOK 
      Height          =   330
      Left            =   150
      Top             =   9165
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
      Caption         =   "DOK"
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   600
      TabIndex        =   18
      Top             =   3960
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   582
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   16777088
      CalendarTrailingForeColor=   16711680
      Format          =   308936705
      CurrentDate     =   38814
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   3000
      TabIndex        =   19
      Top             =   3960
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   582
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   16777088
      CalendarTrailingForeColor=   16711680
      Format          =   308936705
      CurrentDate     =   38814
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par4.frx":30C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par4.frx":3619
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par4.frx":3B5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par4.frx":409D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par4.frx":492F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par4.frx":4E81
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par4.frx":4F93
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par4.frx":54E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par4.frx":5A37
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par4.frx":5F89
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par4.frx":651B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "par4.frx":6AAD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton ShowPar 
      Caption         =   "Αναλυτική Προβολή"
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton SumHmeras 
      Caption         =   "Σύνολα κατα ημέρα"
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FF0000&
      Caption         =   "Κατά Αποθήκη"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   8415
      TabIndex        =   2
      Top             =   1440
      Width           =   1785
   End
   Begin VB.ComboBox APOT 
      Height          =   273
      ItemData        =   "par4.frx":6FFF
      Left            =   8340
      List            =   "par4.frx":7001
      TabIndex        =   5
      Text            =   "APOT"
      Top             =   2970
      Width           =   1875
   End
   Begin VB.CheckBox ekt 
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
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   8430
      MaskColor       =   &H00FF0000&
      TabIndex        =   13
      Top             =   2760
      Width           =   1785
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Σύνολα ανά Παραστατικό"
      Height          =   375
      Left            =   7935
      TabIndex        =   16
      Top             =   1770
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Σύνολα Καταστημάτων"
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      Top             =   2400
      Width           =   2295
   End
   Begin MSComctlLib.ImageCombo Combo1 
      Height          =   330
      Left            =   8340
      TabIndex        =   33
      Top             =   3360
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   12648384
      Text            =   "Combo1"
      ImageList       =   "ImageList1"
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Πωλήσεις"
      Height          =   255
      Left            =   11445
      TabIndex        =   20
      Top             =   3225
      Value           =   1  'Checked
      Width           =   1860
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Αγορές"
      Height          =   195
      Left            =   11445
      TabIndex        =   21
      Top             =   3585
      Value           =   1  'Checked
      Width           =   1860
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Παροχή Υπηρεσιών"
      Height          =   240
      Left            =   11445
      TabIndex        =   22
      Top             =   3915
      Value           =   1  'Checked
      Width           =   1860
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      ItemData        =   "par4.frx":7003
      Left            =   11400
      List            =   "par4.frx":7005
      TabIndex        =   23
      Text            =   "APOT"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   11040
      MultiSelect     =   1  'Simple
      TabIndex        =   26
      Top             =   360
      Width           =   2415
   End
   Begin VB.CheckBox propersi 
      BackColor       =   &H00FF0000&
      Caption         =   "Συγκριση με πρόπερσι"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   11040
      TabIndex        =   27
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF0000&
      Caption         =   "Συγκριση με περυσι"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   11040
      TabIndex        =   28
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Σύνολα κατα μήνα"
      Height          =   255
      Left            =   11040
      TabIndex        =   29
      ToolTipText     =   "Xωρίς ΦΠΑ"
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Παρ/κό"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   299
      Left            =   7839
      TabIndex        =   35
      Top             =   3393
      Width           =   351
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Τύπος Γραφικής Παράστασης"
      Height          =   372
      Left            =   12960
      TabIndex        =   31
      Top             =   4440
      Width           =   1692
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Εως"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Από"
      Height          =   255
      Left            =   -120
      TabIndex        =   7
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Αποθ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   299
      Left            =   7839
      TabIndex        =   24
      Top             =   2990
      Width           =   351
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   10920
      Top             =   45
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Πωλητής"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   11385
      TabIndex        =   25
      Top             =   2520
      Width           =   1890
   End
End
Attribute VB_Name = "par4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Public Col As TrueOleDBGrid80.Column

Public ColS As TrueOleDBGrid80.columns



Dim FIsHidden As Boolean

Dim fSCR      As Object    ' SCRIPT CONTROL

Dim F928      As Integer

Dim f_PALIAXRONIA   ' = Val(FindParametroi(1,"PAR1", "f_PALIAXRONIA", "2005", "Σειρές τιμ.Πώλησης"))

Dim f_PROPERSI      ' = Val(FindParametroi(1,"PAR1", "f_PALIAXRONIA", "2005", "Σειρές τιμ.Πώλησης"))

Dim f_TimMeFPA      ' αν τα τιμολόγια έχουν ΦΠΑ ΣΤΟ ΣΎΝΟΛΟ ΚΑΤΑΣΤΗΜ’ΤΩΝ

Dim Fsql As String


Dim f_arxiko_platos_Grid1 As Long



Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Private Sub chart_Change()
 On Error Resume Next
 G.ChartType = chart.Text
  
End Sub

Private Sub Combo2_Click()

    '<EhHeader>
    On Error GoTo Combo2_Click_Err

    '</EhHeader>

    '---
    Dim where, From, FIELDS, titfields, taxinomisi

    Dim DUM, DB As Database, a

    Dim R As Recordset

    '
    'DUM = OpenD(DB)
    '
    '
    '
    'If Combo2.ListIndex = 0 Then 'ΔΕΛΤΙΑ
    '
    '
    '   a = InputBox("0=ΣΕ ΕΚΚΡΕΜΟΤΗΤΑ 1=ΤΙΜΟΛΟΓΗΘΕΝΤΑ", "ΕΠΙΛΟΓΗ ΠΑΡΑΣΤΑΤΙΚΩΝ", 0)
    '   WHERE = " left(atim,1)='A' and hme>=#" + Format(DTPicker1.Value, "mm/dd/yyyy") + "# and hme<=#" + Format(DTPicker2.Value, "mm/dd/yyyy") + "# "
    '   If a = "0" Then
    '      WHERE = WHERE + " AND ( ART IS NULL OR LEFT(ART,1)=' ' ) "
    '   Else
    '      WHERE = WHERE + " AND LEFT(ART,1)='" + a + "'"
    '   End If
    '   FROM = " tim inner join pel on tim.eidos=pel.eidos and tim.kpe=pel.kod"
    '   FIELDS = "HME,atim,pel.epo,SXETIKO,LEFT(skopos,25) "
    '   titfields = " hme,atim "
    '   taxinomisi = " hme"
    'End If
    '
    'If Combo2.ListIndex = 1 Then  ' ΤΙΜΟΛΟΓΙΑ
    '   'a = InputBox("0=ΣΕ ΕΚΚΡΕΜΟΤΗΤΑ 1=ΤΙΜΟΛΟΓΗΘΕΝΤΑ", "ΕΠΙΛΟΓΗ ΠΑΡΑΣΤΑΤΙΚΩΝ", 0)
    '   WHERE = "LEFT(ATIM,1)='" + Right(Combo1.Text, 1) + "' AND  hme>=#" + Format(DTPicker1.Value, "mm/dd/yyyy") + "# and hme<=#" + Format(DTPicker2.Value, "mm/dd/yyyy") + "# "
    '   FROM = " tim inner join pel on tim.eidos=pel.eidos and tim.kpe=pel.kod"
    '   FIELDS = "HME,atim,PEL.EPO,AJI,AJ1,AJ2,AJ3,AJ4,FPA1+FPA2+FPA3+FPA4 "
    '   titfields = " hme,atim "
    '   taxinomisi = " hme"
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
    'Set R = DB.OpenRecordset("select *from vbrepeid where aa=99")
    'R.Edit
    'R!WHERE = WHERE
    'R!Table = FROM
    'R!FIELDS = FIELDS
    'R!titfields = titfields
    'R!taxinomisi = taxinomisi
    'R.Update
    '
    ''db.Execute "update vbrepeid set vrepeid.from='" + FROM + "' ,fields='" + FIELDS + "', titfields='" + titfields + "', taxinomisi='" + taxinomisi + "' where aa=99"
    'g_ektyp 99
    '
    '
    '
    '
    '

    '<EhFooter>
    Exit Sub

Combo2_Click_Err:
    'MsgBox Err.Description & vbCrLf & _
     "in ADOMERCNEW.par4.Combo2_Click " & _
     "at line " & Erl, _
     vbExclamation + vbOKOnly, "Application Error"
    SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.Combo2_Click " & "at line " & Erl

    Resume Next

    '</EhFooter>

End Sub

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>
        Dim a

        '    F928 = 0

100     a = toascii("C:\MERCVB\F31.TXT", "C:\digicom\HOST.TXT")

110     MsgBox "Τέλος Αποστολής"

        Dim DUM

120     DUM = Shell("c:\mercvb\NOTEPAd.exe c:\DIGICOM\HOST.txt ", vbNormalFocus)

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function toascii(ByVal arxeio As String, moutput As String)

        'a = toascii("C:\MERCVB\F31.TXT", "C:\LIBRA\HOST.TXT")
        ' arxeio π.χ. f31.txt  η φόρμα της εκτύπωσης
        ' Π.Χ.   ^@@@@@@ ^@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ ^#@@@@@@@@ ^@@@@@  ~"0"+MID(DOK.recordset("ERG"),3,5)  ~DOK.recordset("ONO") ~FORMATNUMBER(100*DOK.recordset("LTI5"),0)+"0"  ~LEFT(DOK.recordset("ERG"),7)
        '  host.txt to arxeio me ejagogi ascii
        'ΧΡΕΙΑΖΕΤΑΙ ΜΙΑ ΜΕΤΑΒΛΗΤΗ STATIC GIA OLO το module fSCR
        ' Dim fSCR As Object ' SCRIPT CONTROL
        ' ΚΑΙ ΕΝΑ ADO control p.x. DOK
        '<EhHeader>
        On Error GoTo toascii_Err

        '</EhHeader>

        ' Set fSCR = CreateObject("MSScriptControl.ScriptControl")
        ' DOK.ConnectionString = gConnect
        ' DOK.RecordSource = "select *from EID  WHERE LEFT(ERG,3)='270' "
        ' DOK.Refresh

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

100     Set fSCR = CreateObject("MSScriptControl.ScriptControl")

110     DATA2.Connect = "ODBC;" + gConnect

        'DOK.ConnectionString = gConnect
        'DOK.RecordSource = "select TIM.*,PEL.* from TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD   WHERE MONTH(HME)=1 "
        'DOK.Refresh

        Dim pol As String, polepis As String, ago As String, AGOEPIS As String

        Dim xxx

120     xxx = Get_AJ(pol, polepis, ago, AGOEPIS)

        Dim B

130     B = InputBox("Πωλησεις=1, Αγορές=2   Ζυγιζομενα=3")

140     If B = 1 Then s = pol + "," + polepis
150     If B = 2 Then s = ago + "," + AGOEPIS

        'If b = 3 Then s = ago
        'If b = 4 Then s = agoepis
160     If B < 3 Then
170         s = InputBox("ΚΩΔΙΚΟΙ ΠΑΡΑΣΤΑΤΙΚΩΝ", "", s)
        End If

180     If B < 3 Then
190         DATA2.RecordSource = "select FORMAT(HME,'DDMMYY') AS SHME,iif(ucase(left(TIM.ATIM,1))='P','5','1') AS SHME2,FORMAT(AJ4,'00000.00') AS SAJ4," & " FORMAT(AJ2,'00000.00') AS SAJ2,FORMAT(FPA2,'00000.00') AS SFPA2," & "FORMAT(FPA4,'00000.00') AS SFPA4, TIM.*,PEL.* from TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD" & " WHERE B_N1=1 AND  HME>=#" + Format(DTPicker1.Value, "mm/dd/yyyy") + "# and HME<#" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "#  and  LEFT(ATIM,1) IN (" + s + ") "
        Else
200         DATA2.RecordSource = "select  ONO,LTI5,BARCODES.ERG,EID.KOD,MON,CH1 FROM EID INNER JOIN BARCODES ON EID.KOD=BARCODES.KOD  WHERE LTI5>=0 AND LEFT(BARCODES.ERG,2) in('21','29') "
        End If

210     DATA2.Refresh

220     fSCR.language = "vbscript"
        ' fSCR.AddObject "cTIM", CTIM
230     fSCR.addObject "CDOK", DATA2    'cEGGTIM
240     fSCR.addObject "mactext", MACtEXT
        'CDOK.Recordset.MoveFirst

250     DUM = load_forma(arxeio, 0)

        '-----------    GLOBAL ΜΕΤΑΒΛΗΤΕΣ  ----------------
        'gm_str() οι σειρές του f99.txt
        'gpic(i) το στρινγκ του πεδίου
        'gm_r(i),gm_c(i) σειρά και στήλη του πεδίου
        'gm_f(i) το όνομα του πεδίου

        'gm_str(6)="     ^XXXXXXXXXXXXXXXXXXXXXXXXXXX               ^xxxxxxxx    @XXXXXXXX  ~ono_par  ~SUBS(tim->atim,2,5)  ~TIM->HME
        '        gpic(2)='XXXXXXXXXXXXXXXXXXXXXXXXXXX
        'gm_r(2)=5  gm_c(2)=6    gm_f(2)=ono_par
        '=================================

260     Open moutput For Output As #1

        Dim S2 As String

270     DATA2.Recordset.MoveFirst

280     Do While Not DATA2.Recordset.EOF
290         S2 = ""

300         For k = 1 To 80

310             If Len(gm_f(k)) > 0 Then
320                 MM = mac(Trim(Trim(gm_f(k))))                   ': m_m = ExecuteLine("gvar=" + MM)
330                 S2 = S2 + gAlignFormat(MM, gpic(k))    '
                End If

            Next

340         Print #1, S2
            'Exit Do
350         DATA2.Recordset.MoveNext
        Loop

360     Close #1
370     DATA2.Recordset.Close

380     Set fSCR = Nothing

        '<EhFooter>
        Exit Function

toascii_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.toascii " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.toascii " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub Command2_Click()

        '===========================================
        ' κατα μήνα πωλήσεις αγορές
        '===========================================
        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>
        Dim D                      As String

        Dim pol                    As String

        Dim polepis                As String

        Dim ago                    As String

        Dim AGOEPIS                As String

        Dim a                      As String

        Dim Var                    As String

        Dim Stat(1 To 30, 1 To 12) As Single

        Dim SYN(1 To 12)

        Dim s  As Integer

        Dim R  As Recordset

        Dim k  As Integer, L As Integer, addit As Integer, addit2 As Integer

        Dim d1 As Date, d2 As Date

        Dim SY

        Dim SYNF(12), SYNOLOF

        Dim sy0
        If Not (Check1.Value = vbChecked) Then
              f_arxiko_platos_Grid1 = DBGrid1.Width
              restore_width.Visible = True
              ' f_arxiko_platos_Grid1 = DBGrid1.Width
        End If
        

        'GIA NA BAPSO TO BACKGROUND
        With G.Plot.Backdrop.fill
            .Style = VtFillStyleBrush '<--- You need this
            .Brush.Style = VtBrushStyleSolid '<--- And this
            .Brush.FillColor.Set 200, 200, 200 '<--- Then your color values.
        End With

100     G.ChartType = chart.Text

110     Gdb.Execute "UPDATE TIM SET FPA1=0 WHERE FPA1 IS NULL"
120     Gdb.Execute "UPDATE TIM SET FPA2=0 WHERE FPA2 IS NULL"
130     Gdb.Execute "UPDATE TIM SET FPA3=0 WHERE FPA3 IS NULL"
140     Gdb.Execute "UPDATE TIM SET FPA4=0 WHERE FPA4 IS NULL"
        'Gdb.Execute "UPDATE TIM SET FPA5=0 WHERE FPA1 IS NULL"
150     Gdb.Execute "UPDATE TIM SET FPA6=0 WHERE FPA6 IS NULL"
160     Gdb.Execute "UPDATE TIM SET FPA7=0 WHERE FPA7 IS NULL"


     Gdb.Execute "UPDATE TIM SET AJ1=0 WHERE AJ1 IS NULL"
     Gdb.Execute "UPDATE TIM SET AJ2=0 WHERE AJ2 IS NULL"
     Gdb.Execute "UPDATE TIM SET AJ3=0 WHERE AJ3 IS NULL"
     Gdb.Execute "UPDATE TIM SET AJ4=0 WHERE AJ4 IS NULL"
     Gdb.Execute "UPDATE TIM SET AJ5=0 WHERE AJ5 IS NULL"
     Gdb.Execute "UPDATE TIM SET AJ6=0 WHERE AJ6 IS NULL"
     Gdb.Execute "UPDATE TIM SET AJ7=0 WHERE AJ7 IS NULL"






170     MSFlexGrid1.Visible = False
180     MSFlexGrid2.Visible = False

190     MSFlexGrid1.ColS = 16
200     MSFlexGrid1.rows = 16

        'MSFlexGrid1.DataSource = Null

210     If Check1 Then    ' palia xronia
            ' Data1.DatabaseName = gDir + "\" + Right(Format(Year(Now) - 1, "0000"), 2)
220         Data1.ConnectionString = gConnect + ";DATABASE=" + f_PALIAXRONIA  ' "DSN=" + LTrim(Str(f_PALIAXRONIA)) ' = Val(FindParametroi(1,"PAR1", "f_PALIAXRONIA", "2005", "Σειρές τιμ.Πώλησης"))

            On Error GoTo noLastYear

230         Data1.Refresh
240         d1 = DTPicker1.Value - 365
250         d2 = DTPicker2.Value - 365
260         addit = 1    'gia to Chart
270         addit2 = 7    ' για το grid
            Command3.Visible = False
            Command4.Visible = False
            akyromena.Visible = False
            restore_width.Visible = False
            
            
            
            On Error Resume Next

280         MSFlexGrid1.row = 0: MSFlexGrid1.Col = 7: MSFlexGrid1.Text = "Π.Πωλήσεις"

290     ElseIf propersi Then    'ΠΡΟΠΕΡΣΙΝΗ
            ' Data1.DatabaseName = gDir + "\" + Right(Format(Year(Now) - 1, "0000"), 2)
300         Data1.ConnectionString = gConnect + ";DATABASE=" + Trim(f_PROPERSI)  ' "DSN=" + LTrim(Str(f_PALIAXRONIA)) ' = Val(FindParametroi(1,"PAR1", "f_PALIAXRONIA", "2005", "Σειρές τιμ.Πώλησης"))

            On Error GoTo noLastYear

310         Data1.Refresh
320         d1 = DTPicker1.Value - 365 - 365
330         d2 = DTPicker2.Value - 365 - 365
340         addit = 1    'gia to Chart
350         addit2 = 7    ' για το grid
            Command3.Visible = False
            On Error Resume Next

360         MSFlexGrid1.row = 0: MSFlexGrid1.Col = 7: MSFlexGrid1.Text = "Π.Πωλήσεις"

        Else  ' τρέχουσα χρονιά

            'Data1.DatabaseName = gDir
370         Data1.ConnectionString = gConnect
380         Data1.RecordSource = "select top 1 *from TIM"
390         Data1.Refresh
400         d1 = DTPicker1.Value
410         d2 = DTPicker2.Value
420         addit = 0    'gia to Chart
430         addit2 = 0    ' για το grid
440         MSFlexGrid1.row = 0: MSFlexGrid1.Col = 7: MSFlexGrid1.Text = "Παρ.Υπηρεσιών"
            'MSFlexGrid1.TextMatrix(0, addit2 + 7) = "Παρ.Υπηρεσιών"

            Dim j, i

450         For j = 1 To 12
460             G.Column = j

470             For i = 1 To 6
480                 G.row = i: G.Data = 0
                Next
            Next

        End If

        On Error Resume Next
Dim platos As Long
platos = 1100
490     DBGrid1.Visible = False
500     MSFlexGrid1.Visible = True

510     MSFlexGrid1.row = 0: MSFlexGrid1.Col = 1: MSFlexGrid1.Text = "Πωλήσεις": MSFlexGrid1.ColWidth(MSFlexGrid1.Col) = platos
520     MSFlexGrid1.row = 0: MSFlexGrid1.Col = 2: MSFlexGrid1.Text = "Επιστρ.Πωλ":  MSFlexGrid1.ColWidth(MSFlexGrid1.Col) = platos

530     MSFlexGrid1.row = 0: MSFlexGrid1.Col = 3: MSFlexGrid1.Text = "Καθ.Πωλ.":  MSFlexGrid1.ColWidth(MSFlexGrid1.Col) = platos

540     MSFlexGrid1.row = 0: MSFlexGrid1.Col = 4: MSFlexGrid1.Text = "Αγορές":  MSFlexGrid1.ColWidth(MSFlexGrid1.Col) = platos
550     MSFlexGrid1.row = 0: MSFlexGrid1.Col = 5: MSFlexGrid1.Text = "Επιστ.Αγ.": MSFlexGrid1.ColWidth(MSFlexGrid1.Col) = platos
560     MSFlexGrid1.row = 0: MSFlexGrid1.Col = 6: MSFlexGrid1.Text = "Καθ.Αγορ.":  MSFlexGrid1.ColWidth(MSFlexGrid1.Col) = platos
570     MSFlexGrid1.row = 0: MSFlexGrid1.Col = 7: MSFlexGrid1.Text = "Παρ.Υπηρεσιών"
580    ' MSFlexGrid1.row = 0: MSFlexGrid1.Col = 8: MSFlexGrid1.Text = "ΕΞΟΔΑ"

590     If addit2 > 0 Then
             MSFlexGrid1.row = 0: MSFlexGrid1.Col = 8: MSFlexGrid1.Text = "Π.Πωλήσεις" 'E4JODA

            '  MSFlexGrid1.row = 0:  MSFlexGrid1.col = 8:  MSFlexGrid1.Text = "Π.Επισ.Πωλ"
            '  MSFlexGrid1.row = 0:  MSFlexGrid1.col = 9:  MSFlexGrid1.Text = "Π.Καθ.Πωλ."
            '
            '  MSFlexGrid1.row = 0:  MSFlexGrid1.col = 10:  MSFlexGrid1.Text = "Π.Αγορές"
            '  MSFlexGrid1.row = 0:  MSFlexGrid1.col = 11:  MSFlexGrid1.Text = "Π.Επ.Αγ."
            '  MSFlexGrid1.row = 0:  MSFlexGrid1.col = 13:  MSFlexGrid1.Text = "Καθ.Πωλ."
            'MSFlexGrid1.row = 0: MSFlexGrid1.Col = 0 + addit2: MSFlexGrid1.Text = "Π.Πωλήσεις"
600         MSFlexGrid1.row = 0: MSFlexGrid1.Col = 1 + addit2: MSFlexGrid1.Text = "Π.Πωλήσεις"
610         MSFlexGrid1.row = 0: MSFlexGrid1.Col = 2 + addit2: MSFlexGrid1.Text = "Π.Επιστρ.Πωλ"

620         MSFlexGrid1.row = 0: MSFlexGrid1.Col = 3 + addit2: MSFlexGrid1.Text = "Π.Καθ.Πωλ."

630         MSFlexGrid1.row = 0: MSFlexGrid1.Col = 4 + addit2: MSFlexGrid1.Text = "Π.Αγορές"
640         MSFlexGrid1.row = 0: MSFlexGrid1.Col = 5 + addit2: MSFlexGrid1.Text = "Π.Επιστ.Αγ."
650         MSFlexGrid1.row = 0: MSFlexGrid1.Col = 6 + addit2: MSFlexGrid1.Text = "Π.Καθ.Αγορ."
660         MSFlexGrid1.row = 0: MSFlexGrid1.Col = 7 + addit2: MSFlexGrid1.Text = "Π.Παρ.Υπηρεσιών"
670        ' MSFlexGrid1.row = 0: MSFlexGrid1.Col = 8: MSFlexGrid1.Text = "ΕΞΟΔΑ"

680       '  MSFlexGrid1.row = 0: MSFlexGrid1.Col = 9 + addit2: MSFlexGrid1.Text = "% ΔΙΑΦΟΡΑ"
        Else
            MSFlexGrid1.row = 0: MSFlexGrid1.Col = 8: MSFlexGrid1.Text = "Εξοδα"


        End If

690     MSFlexGrid1.row = 13: MSFlexGrid1.Col = 0: MSFlexGrid1.Text = "Σύνολα"
700     MSFlexGrid1.row = 14: MSFlexGrid1.Col = 0: MSFlexGrid1.Text = "Φ.Π.Α."

710     MSFlexGrid1.row = 15: MSFlexGrid1.Col = 0: MSFlexGrid1.Text = "ΠΛΗΡΩΤΕΟ"

        'PAIRNEI OLES TIS ΠΩΛΗΣΕΙΣ ΑΠΟ ΤΟ ΠΑΡΑΣΤΑΤ
        '    a = Get_AJ(pol, polepis, ago, AGOEPIS)

720     a = Get_AJ_ASCII(pol, polepis, ago, AGOEPIS)

        Dim mPOL, mPOLEPIS, mAGO, mAGOEPIS

730     mPOL = pol: mPOLEPIS = polepis: mAGO = ago: mAGOEPIS = AGOEPIS

        '    'ΠΑΙΡΝΕΙ ΜΟΝΟ ΑΠΟ ΤΟ ΤΟ ΛΙΣΤΒΟΧ ΤΙΣ ΚΙΝΗΣΕΙΣ
        '    pol = ""
        '    polepis = ""
        '    For k = 0 To List1.ListCount - 1
        '
        '        If List1.Selected(k) = True Then
        '            If InStr(mPOL, Right(List1.List(k), 1)) > 0 Then
        '                pol = pol + "'" + Right(List1.List(k), 1) + "',"
        '            ElseIf InStr(mPOLEPIS, Right(List1.List(k), 1)) > 0 Then
        '                polepis = polepis + "'" + Right(List1.List(k), 1) + "',"
        '            End If
        '        End If
        '
        '    Next
        '    If Len(pol) - 1 > 0 Then
        '        pol = Left(pol, Len(pol) - 1)
        '    Else
        '        pol = "' '"
        '    End If
        '
        '    If Len(polepis) - 1 > 0 Then
        '        polepis = Left(polepis, Len(polepis) - 1)
        '    Else
        '        polepis = "' '"
        '    End If
        '

        '-- Πωλήσεις
740     Var = pol: s = 1 + addit2: GoSub findD

        '-- Επιστροφές Πωλήσεων
750     Var = polepis

760     If Len(Var) > 0 Then
770         s = 2 + addit2: GoSub findD
        End If

        On Error GoTo Command2_Click_Err

780     SY = 0  'poliseis

790     For k = 1 To 12
            
800         MSFlexGrid1.TextMatrix(k, addit2 + 3) = Format(Stat(k, addit2 + 1) - Stat(k, addit2 + 2), "###,###.00")
810         SY = SY + (Stat(k, addit2 + 1) - Stat(k, addit2 + 2))
820         MSFlexGrid1.row = k: MSFlexGrid1.Col = addit2 + 3

830         If addit2 = 0 Then    ' fetines poliseis
840             MSFlexGrid1.CellBackColor = vbRed    ' vbYellow
            Else
850             MSFlexGrid1.CellBackColor = vbGreen    ' vbYellow

860             If MSFlexGrid1.TextMatrix(k, 3 + addit2) > 0 Then
870                 MSFlexGrid1.TextMatrix(k, 8 + addit2) = Format((MSFlexGrid1.TextMatrix(k, 3) / MSFlexGrid1.TextMatrix(k, 3 + addit2) - 1) * 100, "##0.00")
                End If
            End If

        Next
        
        
        

'         ' Η ΔΙΑΦΟΡΑ ΜΕ ΠΕΡΥΣΙ ΣΤΟ ΣΥΝΟΛΟ
'        If Check1 Then
'                MSFlexGrid1.Refresh
'                 MSFlexGrid1.TextMatrix(0, 15) = "% ΔΙΑΦ"
'                MSFlexGrid1.TextMatrix(13, 15) = Format((MSFlexGrid1.TextMatrix(13, 3) / MSFlexGrid1.TextMatrix(13, 3 + addit2) - 1) * 100, "##0.00")
'        End If

        
        
        
        
        
        
        For k = 0 To MSFlexGrid1.ColS - 1
          MSFlexGrid1.ColAlignment(k) = 7
        Next
        
        

880     MSFlexGrid1.TextMatrix(13, addit2 + 3) = Format(SY, "###,###.00")

890     SY = gVal(MSFlexGrid1.TextMatrix(14, addit2 + 1)) - gVal(MSFlexGrid1.TextMatrix(14, addit2 + 2))
900     MSFlexGrid1.TextMatrix(14, addit2 + 3) = Format(SY, "###,###.00")
910     sy0 = SY

920     If MSFlexGrid1.TextMatrix(13, 3 + addit2) > 0 Then
930         MSFlexGrid1.TextMatrix(13, 8 + addit2) = Format((MSFlexGrid1.TextMatrix(k, 3) / MSFlexGrid1.TextMatrix(k, 3 + addit2) - 1) * 100, "##0.00")
        End If

        '---- Aγορές
940     Var = ago: s = 4 + addit2: GoSub findD

        '---- Επιστροφές αγορών
950     Var = AGOEPIS: s = 5 + addit2: GoSub findD

960     SY = 0

970     For k = 1 To 12
980         MSFlexGrid1.TextMatrix(k, addit2 + 6) = Format(Stat(k, addit2 + 4) - Stat(k, addit2 + 5), "###,###.00")
990         SY = SY + (Stat(k, addit2 + 4) - Stat(k, addit2 + 5))
1000        MSFlexGrid1.row = k: MSFlexGrid1.Col = addit2 + 6
1010        MSFlexGrid1.CellBackColor = vbCyan 'vbYellow

1020        If addit2 = 0 Then    ' fetines agores
1030            MSFlexGrid1.CellBackColor = vbCyan    ' vbYellow
            Else
1040            MSFlexGrid1.CellBackColor = vbYellow
            End If

        Next

1050    MSFlexGrid1.TextMatrix(13, addit2 + 6) = Format(SY, "###,###.00")

1060    SY = MSFlexGrid1.TextMatrix(14, addit2 + 4) - MSFlexGrid1.TextMatrix(14, addit2 + 5)
1070    MSFlexGrid1.TextMatrix(14, addit2 + 6) = Format(SY, "###,###.00")

        ' φπα πληρωτεο

1080    MSFlexGrid1.TextMatrix(15, addit2 + 3) = Format(sy0 - SY, "###,###.00")

1090    SY = MSFlexGrid1.TextMatrix(14, addit2 + 1) - MSFlexGrid1.TextMatrix(14, addit2 + 2)
1100    MSFlexGrid1.TextMatrix(14, addit2 + 3) = Format(SY, "###,###.00")




  ' Η ΔΙΑΦΟΡΑ ΜΕ ΠΕΡΥΣΙ ΣΤΟ ΣΥΝΟΛΟ
        If Check1 Then
                 
                  MSFlexGrid1.Width = Shape1.Left + Shape1.Width ' - MSFlexGrid1.Left
                 MSFlexGrid1.TextMatrix(0, 15) = "% ΔΙΑΦ"
                MSFlexGrid1.TextMatrix(13, 15) = Format((MSFlexGrid1.TextMatrix(13, 3) / MSFlexGrid1.TextMatrix(13, 3 + addit2) - 1) * 100, "##0.00")
                
                 MSFlexGrid1.Refresh
                
                 
        End If
        
       











1110    For k = 1 To 12
1120        MSFlexGrid1.row = k
1130        MSFlexGrid1.Col = 0
1140        MSFlexGrid1.Text = k
        Next

        ' Γραφική παράσταση
1150    s = 0

        ' Dim PY
1160    Var = Get_PY()
        '-- PAROXH YPIRESION
1170    s = 3 + addit2

If Len(Var) = 0 Then
  Var = " '' "
End If

 GoSub findD2

        ' Ε Ξ Ο Δ Α
1180    Var = "'Ξ'"
1190    s = 8 + addit2: GoSub findD2

        Dim Sexoda

1200    Sexoda = 0

1210    SY = 0

1220    For k = 1 To 12
1230        SY = SY + Stat(k, addit2 + 3)    'Stat(K, addit2 + 7)
1240        MSFlexGrid1.row = k: MSFlexGrid1.Col = addit2 + 7
            '  MSFlexGrid1.CellBackColor = vbMagenta ' vbYellow

1250        If addit2 = 0 Then    ' fetines paroxes
1260            MSFlexGrid1.CellBackColor = vbMagenta    ' vbYellow
            Else
1270            MSFlexGrid1.CellBackColor = vbCyan    ' vbYellow
            End If

1280        MSFlexGrid1.TextMatrix(k, addit2 + 7) = Stat(k, addit2 + 3)

            'ejoda
1290        Sexoda = Sexoda + Stat(k, addit2 + 8)
1300        MSFlexGrid1.row = k: MSFlexGrid1.Col = addit2 + 8
1310        MSFlexGrid1.CellBackColor = vbMagenta    ' vbYellow
1320        MSFlexGrid1.TextMatrix(k, addit2 + 8) = Stat(k, addit2 + 8)

        Next

1330    MSFlexGrid1.TextMatrix(13, addit2 + 7) = Format(SY, "###,###.00")
        MSFlexGrid1.TextMatrix(13, addit2 + 8) = Format(Sexoda, "###,###.00")

1340    For L = 1 To 12

            ' Exit Sub
1350        If Check3 Then    ' πωλησεις
1360            G.row = 1 + addit: G.Column = L: G.Data = Stat(L, addit2 + 1) - Stat(L, addit2 + 2)

1370            If addit = 1 Then
1380                G.RowLabel = "Περ.Πωλήσεις"
                Else
1390                G.RowLabel = "Πωλησεις"
                End If
            End If

            'agores
1400        If Check4 Then  ' αγορές
1410            G.row = 3 + addit: G.Column = L: G.Data = Stat(L, addit2 + 4) - Stat(L, addit2 + 5)

1420            If addit = 1 Then
1430                G.RowLabel = "Περ.Αγορές"
                Else
1440                G.RowLabel = "Αγορές"
                End If
            End If

            'paroxh

1450        If Check5 Then
1460            G.row = 5 + addit: G.Column = L: G.Data = Stat(L, addit2 + 3)

1470            If addit = 1 Then
1480                G.RowLabel = "Περ.Παρ.Υπηρ"
                Else
1490                G.RowLabel = "Παρ.Υπηρ"
                End If

            End If

        Next

1500    G.row = 5 + addit: G.RowLabel = " "

1510    MSFlexGrid1.Visible = True

1520    Me.MousePointer = vbNormal

        Exit Sub

findD:

1530    D = "select MONTH(HME) as shme,sum(AJ1+AJ2+AJ3+AJ4+AJ5+(CASE WHEN AJ6 IS NULL THEN 0 ELSE AJ6 END) +"
1540    D = D + "(CASE WHEN AJ7 IS NULL THEN 0 ELSE AJ7 END) ) as saji,sum(FPA1+FPA2+FPA3+FPA4+FPA6+FPA7) as SFPA   "
1550    D = D + "from TIM  where  ascii(LEFT(ATIM,1)) IN(" + Var + ")  "

       If akyromena.Value = vbChecked Then
            D = D + " and LEFT(ATIM,1) NOT IN (SELECT AKYR FROM PARASTAT WHERE AKYR IS NOT NULL ) AND ISNULL(AKYROMENO,0)=0   "
        End If




1560    If gVal(apot.Text) < 99 Then
1570        D = D + " and B_N1=" + Left(apot.Text, 1)    ' Str(1 + apot.ListIndex)
        End If

1580    D = D + " and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2 + 1, "mm/dd/yyyy") + "' "
1590    D = D + "group by MONTH(HME)  ;"  ' collate Greek_CS_AS  ;"
1600    Data1.RecordSource = D
1610    Data1.Refresh

        On Error Resume Next

1620    For k = 1 To 12: SYNF(k) = 0: Next

1630    Data1.Recordset.MoveFirst

1640    Do While Not Data1.Recordset.EOF
1650        Stat(gVal(Data1.Recordset("shme")), s) = Data1.Recordset("saji")
1660        SYNF(gVal(Data1.Recordset("shme"))) = Data1.Recordset("sfpa")    '
1670        Data1.Recordset.MoveNext

1680        DoEvents
        Loop

1690    SYN(s) = 0
1700    SYNOLOF = 0

1710    For k = 1 To 12
1720        MSFlexGrid1.row = k
1730        MSFlexGrid1.Col = s
1740        MSFlexGrid1.Text = Format(Stat(k, s), "###,###.00")
1750        SYN(s) = SYN(s) + Stat(k, s)
1760        SYNOLOF = SYNOLOF + SYNF(k)
        Next

1770    MSFlexGrid1.row = 13
1780    MSFlexGrid1.Col = s
1790    MSFlexGrid1.Text = Format(SYN(s), "###,###.00")

1800    MSFlexGrid1.row = 14
1810    MSFlexGrid1.Col = s
1820    MSFlexGrid1.Text = Format(SYNOLOF, "###,###.00")

1830    Return

findD2:

1840    D = "select MONTH(HME) as shme,sum(AJ1+AJ2+AJ3+AJ4+AJ5+AJ6) as saji "
1850    D = D + "from TIM  where  left(ATIM,1) IN (" + Var + ")  "
1860    D = D + "and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2 + 1, "mm/dd/yyyy") + "' "
1870    D = D + "group by MONTH(HME) ;" '  collate Greek_CS_AS ;"
1880    Data1.RecordSource = D
1890    Data1.Refresh

        On Error Resume Next

1900    Data1.Recordset.MoveFirst

1910    Do While Not Data1.Recordset.EOF
1920        Stat(gVal(Data1.Recordset("shme")), s) = Data1.Recordset("saji")
1930        Data1.Recordset.MoveNext

        Loop

1940    Return

noLastYear:
1950    MsgBox "δεν υπάρχει η παλιά χρονιά"

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function Get_PY()

        '<EhHeader>
        On Error GoTo Get_PY_Err

        '</EhHeader>

        Dim DB As Database

        Dim R  As New ADODB.Recordset, X As String

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

100     R.Open UCase("select pol,eidos,ajia_apou,SHM1 from parastat"), Gdb, adOpenDynamic, adLockOptimistic

        Dim pol As String

110     pol = ""
120     R.MoveFirst

130     Do While Not R.EOF

140         If Not IsNull(R("eidos") + R("pol") + R("SHM1")) Then
150             If R("pol") = "1" And Left(R("SHM1"), 1) = "2" Then
160                 pol = pol + "'" + R("eidos") + "',"
                End If
            End If

170         R.MoveNext
        Loop

180     Get_PY = Left(pol, Len(pol) - 1)

        '<EhFooter>
        Exit Function

Get_PY_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.Get_PY " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.Get_PY " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function
'Private Sub Command3_Click()
'   G.Top = 4680  'arxika
'   G.height = 4485
'
'   G.Top = 46
'   G.height = 8485
'
'
'
'End Sub

Private Sub Command3_Click()
With DBGrid1.PrintInfo
        ' Set the page header
        .PageHeaderFont.Italic = True
        .PageHeader = "Παραστατικά"
        .PreviewInitZoom = 0.5
        ' Column headers will be on every page
         .RepeatColumnHeaders = True
        .SettingsOrientation = 2
        ' Display page numbers (centered)
        .PageFooter = "\tPage: \p"
        ' Invoke Print Preview
        .PrintPreview
    End With
End Sub

Private Sub Command4_Click()

        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>

        Dim N As Integer

100     N = 0
        'KRITHRIA.Grid1.Clear

110     If Not FIsHidden Then
120         N = 0

130         With KRITHRIA.Grid1
140             N = N + 1: .TextMatrix(N, 0) = "Επωνυμία Συναλασσόμενου": .TextMatrix(N, 3) = "LEFT(PEL.EPO,25)": .TextMatrix(N, 4) = "C"
150             N = N + 1: .TextMatrix(N, 0) = "Κωδικός": .TextMatrix(N, 3) = "KPE": .TextMatrix(N, 4) = "C"
160             N = N + 1: .TextMatrix(N, 0) = "Πίστ/Μετρ": .TextMatrix(N, 3) = "TRP": .TextMatrix(N, 4) = "C"
170             N = N + 1: .TextMatrix(N, 0) = "Αξία": .TextMatrix(N, 3) = "AJI": .TextMatrix(N, 4) = "n"
180             N = N + 1: .TextMatrix(N, 0) = "Πωλητής": .TextMatrix(N, 3) = "EIDPAR": .TextMatrix(N, 4) = "c"
190             N = N + 1: .TextMatrix(N, 0) = "ΣΚΟΠΟΣ π.χ. %ΠΩΛ": .TextMatrix(N, 3) = "SKOPOS": .TextMatrix(N, 4) = "C"

                '[Συναλασσόμενος]

            End With

200         FIsHidden = True
        End If

210     KRITHRIA.SHOW 1
        ' Fsql = MDIForm1.StatusBar1.Panels(6).Text
        ' MDIForm1.StatusBar1.Panels(6).Text = ""

220     Fsql = MDIForm1.StatusBar1.Panels(6).Text
230     MDIForm1.StatusBar1.Panels(6).Text = ""

        ' With KRITHRIA.Grid1
        '    .TextMatrix(1, 0) = "Κωδικός":  .TextMatrix(1, 3) = "KPE": .TextMatrix(1, 4) = "c"
        '    .TextMatrix(2, 0) = "Πίστ/Μετρ": .TextMatrix(2, 3) = "TRP": .TextMatrix(2, 4) = "C"
        '    .TextMatrix(3, 0) = "Αξία": .TextMatrix(3, 3) = "AJI": .TextMatrix(3, 4) = "n"
        '    .TextMatrix(4, 0) = "Πωλητής": .TextMatrix(4, 3) = "EIDPAR": .TextMatrix(4, 4) = "c"
        '
        ' End With
        '
        '    KRITHRIA.SHOW 1
        '    Fsql = MDIForm1.StatusBar1.Panels(6).Text
        '    MDIForm1.StatusBar1.Panels(6).Text = ""

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub NEW_ANA_PARASTATIKO()
  '-------------  ana parastatiko  ---------------
        '<EhHeader>
        On Error GoTo Command5_Click_Err

        '</EhHeader>

        Dim D As String

        Dim MSYNT As String

        Dim SA As Single

        Dim SP As Single

        Dim k As Long

        Dim mPOL As String

        Dim mAGO As String

        Dim R As New ADODB.Recordset

        Dim SP2 As Single

        Dim a

        Dim AGOEPIS As String

        Dim polepis As String, pol As String, ago As String, APIS

100     pol = ""
110     a = Get_AJ_ASCII(pol, polepis, ago, AGOEPIS)

120     mPOL = pol
130     mAGO = ago

        'mPOL = mPOL + "''"
        'mAGO = mAGO + "''"
    
        Dim poliseis As Integer

140     poliseis = MsgBox("ΠΩΛΗΣΕΙΣ=ΝΑΙ  ΟΧΙ=ΑΓΟΡΕΣ", vbYesNo)
    
150     If poliseis = vbNo Then
160         pol = ago
170         polepis = AGOEPIS
        End If
        If Len(polepis) = 0 Then polepis = "0"  ' ΓΙΑ ΝΑ ΜΗΝ ΚΟΛΛΑΕΙ
180     DBGrid1.Visible = True
190     MSFlexGrid1.Visible = False

        On Error Resume Next

200     Me.MousePointer = vbHourglass

210     If apot.ListIndex = -1 Then
220         MSYNT = "AKYROMENO<>1 AND ascii(LEFT(ATIM,1)) IN (" + pol + "," + polepis + ") AND "
        Else
230         MSYNT = "AKYROMENO<>1 AND  ascii(LEFT(ATIM,1)) IN (" + pol + "," + polepis + ") AND  APOT=" + str(apot.ListIndex) + " AND "
        End If

        Dim sql As String


Dim r3 As New ADODB.Recordset

r3.Open "SELECT COUNT(*), ASCII(EIDOS) FROM PARASTAT GROUP BY ASCII(EIDOS) HAVING COUNT(*)>1 ", Gdb, adOpenDynamic, adLockOptimistic
If Not r3.EOF Then
   If r3(0) > 1 Then
      MsgBox "Προσοχή υπάρχει δύο φορές το παραστατικό " + Chr(r3(1))
      MsgBox "διορθώστε το στις παραμέτρους παραστατικών"
     ' Exit Sub
   End If
  
   
   
End If
r3.Close


Gdb.Execute "UPDATE TIM SET AJ6=0 WHERE AJ6 IS NULL"


Gdb.Execute "SELECT ATIM INTO DOKTIMATIM  from TIM  WHERE " _
           + MSYNT + "  HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "' ORDER BY ATIM "


             ''left(PARASTAT.EIDOS+' '+PARASTAT.TITLOS+'                     ',25) AS [Παραστατικό],"
240     sql = "SELECT LEFT(ATIM,1) AS [ΚΩΔ.ΠΑΡΑΣΤ] " _
           & " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN AJ2 ELSE -(AJ2)  END )) AS [" + Format(g_Fpa(2), "##") + "%], " _
           & " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN AJ3 ELSE -(AJ3)  END )) AS [" + Format(g_Fpa(3), "##") + "%], " _
           & " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN AJ4 ELSE -(AJ4)  END )) AS [" + Format(g_Fpa(4), "##") + "%], " _
           & " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN AJ1 ELSE -(AJ1)  END )) AS [" + Format(g_Fpa(1), "##") + "%], " _
           & " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN AJ5 ELSE -(AJ5)  END )) AS [0%], " _
           & " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN AJ6 ELSE -(AJ6)  END )) AS [" + Format(g_Fpa(6), "##") + "%], " _
           & " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN AJ7 ELSE -(AJ7)  END )) AS [" + Format(g_Fpa(7), "##") + "%], "
           
           
           sql = sql + " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7 ELSE -(AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7)  END )) AS [ΣΥΝΟΛΟ], "
           If poliseis = vbYes Then
             ' sql = sql + " (SELECT substring(min(ATIM),2,6) FROM DOKTIMATIM WHERE LEFT(ATIM,1)=PARASTAT.EIDOS) AS [ΑΠΟ ΠΑΡ.], (SELECT substring(MAX(ATIM),2,6) FROM DOKTIMATIM WHERE LEFT(ATIM,1)=PARASTAT.EIDOS) AS [ΤΕΛ.ΠΑΡ],"
           End If
           
           'INNER JOIN PARASTAT ON ascii(LEFT(ATIM,1))=ascii(PARASTAT.EIDOS) WHERE "
          sql = sql + " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN KERDOS ELSE 0  END )) AS [ΚΕΡΔΟΣ] " _
           & " from TIM WHERE " _
           + MSYNT + "  HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "' " _
           & "  GROUP BY LEFT(ATIM,1)"
           'PARASTAT.EIDOS,PARASTAT.TITLOS " '  collate Greek_CS_AS"
           
           
           
250     Data1.RecordSource = sql
260     Data1.Refresh







270     If ekt Then
            'typos2
280         print3_xar sql, "0111111111111", "Από " + Format(DTPicker1.Value, "mm/dd/yyyy") + " έως " + Format(DTPicker2.Value, "mm/dd/yyyy"), 0

        End If
        
  Gdb.Execute "DROP TABLE DOKTIMATIM "
        

290     DBGrid1.Visible = False
300     MSFlexGrid1.Visible = False
310     MSFlexGrid2.Visible = True

320     MSFlexGrid2.AddItem " " & Chr(9)
330     MSFlexGrid2.ColAlignment(1) = 6    'RIGHT
340     SA = 0
350     SP = 0
360     k = 0
370     SP2 = 0

        Dim SEP
       Dim SUMA(30) As Single
       Dim KOL As Integer
       
380     SEP = 0

390     MSFlexGrid2.ColWidth(0) = 4000



For KOL = 1 To MSFlexGrid2.ColS - 1
        SUMA(KOL) = 0
        For k = 1 To MSFlexGrid2.rows - 2
          'SA = SA + gVal(MSFlexGrid2.TextMatrix(k, KOL))
            'SP = SP + Val(MSFLEXGRID2.TextMatrix(K, 3))
            SUMA(KOL) = SUMA(KOL) + gVal(MSFlexGrid2.TextMatrix(k, KOL))
        Next
        If KOL = 10 Or KOL = 9 Then  ' να μην κανει σουμα το πρωτο και τελειταιο παραστατικο"
        Else
            MSFlexGrid2.TextMatrix(k, KOL) = Format(SUMA(KOL), "###,###,##0.00")
        End If
        
        MSFlexGrid2.ColWidth(KOL) = 1200
        MSFlexGrid2.ColAlignment(KOL) = 6
Next



420     ' MSFlexGrid2.TextMatrix(k, 1) = SA
        'MSFLEXGRID2.TextMatrix(K, 3) = SP
        'MSFLEXGRID2.TextMatrix(K, 4) = SP2
        'MSFLEXGRID2.TextMatrix(K, 5) = SEP
430     MSFlexGrid2.TopRow = MSFlexGrid2.rows - 10

440     Me.MousePointer = vbNormal

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub





Private Sub Command5_Click()

 NEW_ANA_PARASTATIKO
 Exit Sub
 
 



        '-------------  ana parastatiko  ---------------
        '<EhHeader>
        On Error GoTo Command5_Click_Err

        '</EhHeader>

        Dim D As String

        Dim MSYNT As String

        Dim SA As Single

        Dim SP As Single

        Dim k As Long

        Dim mPOL As String

        Dim mAGO As String

        Dim R As New ADODB.Recordset

        Dim SP2 As Single

        Dim a

        Dim AGOEPIS As String

        Dim polepis As String, pol As String, ago As String, APIS

100     pol = ""
110     a = Get_AJ_ASCII(pol, polepis, ago, AGOEPIS)

120     mPOL = pol
130     mAGO = ago

        'mPOL = mPOL + "''"
        'mAGO = mAGO + "''"
    
        Dim poliseis As Integer

140     poliseis = MsgBox("ΠΩΛΗΣΕΙΣ=ΝΑΙ  ΟΧΙ=ΑΓΟΡΕΣ", vbYesNo)
    
150     If poliseis = vbNo Then
160         pol = ago
170         polepis = AGOEPIS
        End If
        If Len(polepis) = 0 Then polepis = "0"  ' ΓΙΑ ΝΑ ΜΗΝ ΚΟΛΛΑΕΙ
180     DBGrid1.Visible = True
190     MSFlexGrid1.Visible = False

        On Error Resume Next

200     Me.MousePointer = vbHourglass

210     If apot.ListIndex = -1 Then
220         MSYNT = "AKYROMENO<>1 AND ascii(LEFT(ATIM,1)) IN (" + pol + "," + polepis + ") AND "
        Else
230         MSYNT = "AKYROMENO<>1 AND  ascii(LEFT(ATIM,1)) IN (" + pol + "," + polepis + ") AND  APOT=" + str(apot.ListIndex) + " AND "
        End If

        Dim sql As String


Dim r3 As New ADODB.Recordset

r3.Open "SELECT COUNT(*), ASCII(EIDOS) FROM PARASTAT GROUP BY ASCII(EIDOS) HAVING COUNT(*)>1 ", Gdb, adOpenDynamic, adLockOptimistic
If Not r3.EOF Then
   If r3(0) > 1 Then
      MsgBox "Προσοχή υπάρχει δύο φορές το παραστατικό " + Chr(r3(1))
      MsgBox "διορθώστε το στις παραμέτρους παραστατικών"
     ' Exit Sub
   End If
  
   
   
End If
r3.Close


Gdb.Execute "UPDATE TIM SET AJ6=0 WHERE AJ6 IS NULL"


Gdb.Execute "SELECT ATIM INTO DOKTIMATIM  from TIM  WHERE " _
           + MSYNT + "  HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "' ORDER BY ATIM "



240     sql = "SELECT left(PARASTAT.EIDOS+' '+PARASTAT.TITLOS+'                     ',25) AS [Παραστατικό]," _
           & " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN AJ2 ELSE -(AJ2)  END )) AS [" + Format(g_Fpa(2), "##") + "%], " _
           & " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN AJ3 ELSE -(AJ3)  END )) AS [" + Format(g_Fpa(3), "##") + "%], " _
           & " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN AJ4 ELSE -(AJ4)  END )) AS [" + Format(g_Fpa(4), "##") + "%], " _
           & " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN AJ1 ELSE -(AJ1)  END )) AS [" + Format(g_Fpa(1), "##") + "%], " _
           & " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN AJ5 ELSE -(AJ5)  END )) AS [0%], " _
           & " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN AJ6 ELSE -(AJ6)  END )) AS [" + Format(g_Fpa(6), "##") + "%], " _
           & " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN AJ7 ELSE -(AJ7)  END )) AS [" + Format(g_Fpa(7), "##") + "%], "
           
           
           sql = sql + " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7 ELSE -(AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7)  END )) AS [ΣΥΝΟΛΟ], "
           If poliseis = vbYes Then
              sql = sql + " (SELECT substring(min(ATIM),2,6) FROM DOKTIMATIM WHERE LEFT(ATIM,1)=PARASTAT.EIDOS) AS [ΑΠΟ ΠΑΡ.], (SELECT substring(MAX(ATIM),2,6) FROM DOKTIMATIM WHERE LEFT(ATIM,1)=PARASTAT.EIDOS) AS [ΤΕΛ.ΠΑΡ],"
           End If
           
           
          sql = sql + " CONVERT(DECIMAL(10,2),sum(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + pol + ") THEN KERDOS ELSE 0  END )) AS [ΚΕΡΔΟΣ] " _
           & " from TIM INNER JOIN PARASTAT ON ascii(LEFT(ATIM,1))=ascii(PARASTAT.EIDOS) WHERE " _
           + MSYNT + "  HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "' " _
           & "  GROUP BY PARASTAT.EIDOS,PARASTAT.TITLOS " '  collate Greek_CS_AS"
           
           
           
250     Data1.RecordSource = sql
260     Data1.Refresh







270     If ekt Then
            'typos2
280         print3_xar sql, "0111111111111", "Από " + Format(DTPicker1.Value, "mm/dd/yyyy") + " έως " + Format(DTPicker2.Value, "mm/dd/yyyy"), 0

        End If
        
  Gdb.Execute "DROP TABLE DOKTIMATIM "
        

290     DBGrid1.Visible = False
300     MSFlexGrid1.Visible = False
310     MSFlexGrid2.Visible = True

320     MSFlexGrid2.AddItem " " & Chr(9)
330     MSFlexGrid2.ColAlignment(1) = 6    'RIGHT
340     SA = 0
350     SP = 0
360     k = 0
370     SP2 = 0

        Dim SEP
       Dim SUMA(30) As Single
       Dim KOL As Integer
       
380     SEP = 0

390     MSFlexGrid2.ColWidth(0) = 4000



For KOL = 1 To MSFlexGrid2.ColS - 1
        SUMA(KOL) = 0
        For k = 1 To MSFlexGrid2.rows - 2
          'SA = SA + gVal(MSFlexGrid2.TextMatrix(k, KOL))
            'SP = SP + Val(MSFLEXGRID2.TextMatrix(K, 3))
            SUMA(KOL) = SUMA(KOL) + gVal(MSFlexGrid2.TextMatrix(k, KOL))
        Next
        If KOL = 10 Or KOL = 9 Then  ' να μην κανει σουμα το πρωτο και τελειταιο παραστατικο"
        Else
            MSFlexGrid2.TextMatrix(k, KOL) = Format(SUMA(KOL), "###,###,##0.00")
        End If
        
        MSFlexGrid2.ColWidth(KOL) = 1200
        MSFlexGrid2.ColAlignment(KOL) = 6
Next



420     ' MSFlexGrid2.TextMatrix(k, 1) = SA
        'MSFLEXGRID2.TextMatrix(K, 3) = SP
        'MSFLEXGRID2.TextMatrix(K, 4) = SP2
        'MSFLEXGRID2.TextMatrix(K, 5) = SEP
430     MSFlexGrid2.TopRow = MSFlexGrid2.rows - 10

440     Me.MousePointer = vbNormal

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command6_Click()

        ' Σύνολα Καταστημάτων
        '-----------------------------------------------------
        '<EhHeader>
        On Error GoTo Command6_Click_Err

        '</EhHeader>
        Dim D As String

        Dim MSYNT As String

        Dim SA As Single

        Dim SP As Single

        Dim k As Long

        Dim mPOL As String

        Dim mAGO As String

        Dim R As New ADODB.Recordset

        Dim SP2 As Single

        Dim a

        Dim AGOEPIS As String

        Dim polepis As String, pol As String, ago As String, APIS

100     pol = ""
110     a = Get_AJ(pol, polepis, ago, AGOEPIS)

120     mPOL = pol
130     mAGO = ago

140     mPOL = mPOL + "''"
150     mAGO = mAGO + "''"
160     DBGrid1.Visible = True
170     MSFlexGrid1.Visible = False

        On Error Resume Next

        If Check1 Then    ' palia xronia

            On Error GoTo 0

            YPOL_PALIA_TAMEIAKES
         
            '         data1.ConnectionString = gConnect + ";DATABASE=" + f_PALIAXRONIA  ' "DSN=" + LTrim(Str(f_PALIAXRONIA)) ' = Val(FindParametroi(1,"PAR1", "f_PALIAXRONIA", "2005", "Σειρές τιμ.Πώλησης"))
            '        D1 = DTPicker1.Value - 365
            '       D2 = DTPicker2.Value - 365
     
        End If

180     Me.MousePointer = vbHourglass

190     If Check2 Then

200         If apot.ListIndex = -1 Then
210             MSYNT = ""
            Else
220             MSYNT = " APOT=" + str(1 + apot.ListIndex) + " AND "
            End If

            '  data1.RecordSource = "SELECT  CONVERT(CHAR(12),HME,3) ," _
            '  & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=1 AND LEFT(ATIM,1) IN ('L') THEN PIS*TIMM ELSE 0 END )) AS [TAM.1]," _
            '  & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=1 AND LEFT(ATIM,1) IN ('T','t') THEN PIS*TIMM*(1-EKPT/100)* (CASE WHEN FPA=4 THEN 1.04 ELSE 1.19 END ) ELSE 0 END)) AS [TIM.1]," _
            '  & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=2 AND LEFT(ATIM,1) IN ('L') THEN PIS*TIMM ELSE 0 END)) AS [TAM.2]," _
            '  & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=2 AND LEFT(ATIM,1) IN ('T','t') THEN POSO*TIMM*(1-EKPT/100)*(CASE WHEN FPA=4 THEN 1.04 ELSE 1.19 END ) ELSE 0 END )) AS [TIM.2]" _
            '  & " from EGGTIM WHERE " _
            '  + MSYNT + "  HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "' " _
            '  & "  GROUP BY CONVERT(CHAR(12),HME,3) ORDER BY CONVERT(CHAR(12),HME,3)"
            '  data1.Refresh
            On Error Resume Next

230         Gdb.Execute "DROP TABLE DOKSYNOLA"
240         Gdb.Execute "UPDATE EGGTIM SET EKPT=0 WHERE EKPT IS NULL"

            On Error GoTo Command6_Click_Err

250         If f_TimMeFPA = 1 Then

260             Gdb.Execute "SELECT ROW_NUMBER() OVER( ORDER BY CONVERT(CHAR(12),HME,3) ) AS AA, CONVERT(CHAR(12),HME,3) AS HMER ," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=1 AND LEFT(ATIM,1) IN ('L') THEN POSO*TIMM*(1-EKPT/100) ELSE 0 END )) AS [TAM.1]," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=1 AND LEFT(ATIM,1) IN ('T','t') THEN MIK_AJIA ELSE 0 END)) AS [TIM.1]," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=1 AND LEFT(ATIM,1) IN ('l','R','μ') THEN MIK_AJIA  ELSE 0 END)) AS [in.1]," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=2 AND LEFT(ATIM,1) IN ('L') THEN MIK_AJIA  ELSE 0 END)) AS [TAM.2]," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=2 AND LEFT(ATIM,1) IN ('l','R','μ') THEN MIK_AJIA  ELSE 0 END)) AS [in.2]," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=2 AND LEFT(ATIM,1) IN ('T','t') THEN MIK_AJIA ELSE 0 END )) AS [TIM.2]," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=3 AND LEFT(ATIM,1) IN ('L') THEN MIK_AJIA  ELSE 0 END)) AS [TAM.3]," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=3 AND LEFT(ATIM,1) IN ('l','R','μ') THEN MIK_AJIA  ELSE 0 END)) AS [in.3]," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=3 AND LEFT(ATIM,1) IN ('T','t') THEN MIK_AJIA ELSE 0 END )) AS [TIM.3],0 AS [ΣΥΝΟΛΑ]," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=4 AND (PELKOD IS NULL) AND  LEFT(ATIM,1) IN ('L') THEN POSO*TIMM*(100-EKPT)/100  ELSE 0 END)) AS [tam-4]," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=4 AND LEFT(ATIM,1) IN ('R') THEN POSO*TIMM*(100-EKPT)/100 ELSE 0 END)) AS [int-4]," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=4 AND LEFT(ATIM,1) IN ('T') THEN 1.24*POSO*TIMM*(100-EKPT)/100  ELSE 0 END)) AS [tim-4] " _
                   & " INTO DOKSYNOLA from EGGTIM WHERE " _
                   + MSYNT + "  HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "' " _
                   & "  GROUP BY CONVERT(CHAR(12),HME,3) ORDER BY  CONVERT(CHAR(12),HME,3)"

            Else
270             Gdb.Execute "SELECT ROW_NUMBER() OVER( ORDER BY CONVERT(CHAR(12),HME,3) ) AS AA, CONVERT(CHAR(12),HME,3) AS HMER ," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=1 AND LEFT(ATIM,1) IN ('L') THEN PIS*TIMM ELSE 0 END )) AS [TAM.1]," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=1 AND LEFT(ATIM,1) IN ('T','t') THEN PIS*TIMM*(1-EKPT/100) ELSE 0 END)) AS [TIM.1]," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=3 AND LEFT(ATIM,1) IN ('L') THEN PIS*TIMM ELSE 0 END)) AS [TAM.3]," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=3 AND LEFT(ATIM,1) IN ('T','t') THEN POSO*TIMM*(1-EKPT/100) ELSE 0 END )) AS [TIM.3]," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=2 AND LEFT(ATIM,1) IN ('L') THEN PIS*TIMM ELSE 0 END)) AS [TAM.2]," _
                   & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=2 AND LEFT(ATIM,1) IN ('T','t') THEN POSO*TIMM*(1-EKPT/100) ELSE 0 END )) AS [TIM.2],0 AS [ΣΥΝΟΛΑ]" _
                   & " INTO DOKSYNOLA from EGGTIM WHERE " _
                   + MSYNT + "  HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "' " _
                   & "  GROUP BY CONVERT(CHAR(12),HME,3) ORDER BY CONVERT(CHAR(12),HME,3)"
            End If

            Dim DB As Database

280         Gdb.Execute "UPDATE DOKSYNOLA SET [ΣΥΝΟΛΑ]=[TAM.1]+[TIM.1]+[TAM.2]+[TIM.2]+[TAM.3]+[TIM.3]+[in.1]+[in.2]+[in.3]+[tam-4]+[int-4]+[tim-4]"

            If Check1 Then    ' palia xronia
            
                On Error GoTo 0

                Data1.RecordSource = "SELECT F.HMER,F.[TAM.1],P.[TAM.1],F.[TIM.1],P.[TIM.1],F.[TAM.1]+F.[TIM.1] AS [ΔΡΑΜΑ],P.[TAM.1]+P.[TIM.1] AS [Π.ΔΡΑΜΑ],F.[TAM.3],P.[TAM.3],F.[TIM.3],P.[TIM.3],F.[TAM.3]+F.[TIM.3] AS [ΞΑΝΘΗ],P.[TAM.3]+P.[TIM.3] AS [Π.ΞΑΝΘΗ] ,F.[in.1],P.[in.1],F.[in.3],P.[in.3],F.[in.1]+F.[in.3] AS [ΙΝΤΕΡΝΕΤ],P.[in.1]+P.[in.3] AS [Π.ΙΝΤΕΡΝΕΤ]  FROM DOKSYNOLA F INNER JOIN " + f_PALIAXRONIA + ".dbo.DOKSYNOLA P ON F.AA=P.AA   ORDER BY F.HMER "
                Data1.Refresh
            Else
290             Data1.RecordSource = "SELECT HMER,[TAM.1],[TIM.1],[TAM.1]+[TIM.1] AS [ΔΡΑΜΑ],[TAM.3],[TIM.3],[TAM.3]+[TIM.3] AS [ΞΑΝΘΗ] ,[in.1],[in.3],[in.1]+[in.3] AS [ΙΝΤΕΡΝΕΤ],[tam-4],[int-4],[tim-4]  FROM DOKSYNOLA order by HMER "
300             Data1.Refresh

            End If

310         If ekt Then

                'typos2
                On Error GoTo Command6_Click_Err

320            ' print3_xar "SELECT HMER,[TAM.1],[TIM.1],[TAM.1]+[TIM.1] AS [ΔΡΑΜΑ],[TAM.3],[TIM.3],[TAM.3]+[TIM.3] AS [ΞΑΝΘΗ] ,[in.1],[in.3],[in.1]+[in.3] AS [ΙΝΤΕΡΝΕΤ]  FROM DOKSYNOLA", "11111111", "Σύνολα ανά κατάστημα", 0

                print3_xar Data1.RecordSource, "0111111111111", "Σύνολα ανά κατάστημα", 0


            End If

            'Dim DB As Database
            '  Set DB = OpenDatabase("", False, False, gConnect)

            On Error GoTo Command6_Click_Err ' Resume Next

            '        Kill "C:\SYNOLA.DBF"

            '       db.Execute "SELECT * INTO SYNOLA IN 'C:\' 'dBASE III;' FROM DOKSYNOLA"

            Dim DD

            ' *(WHEN FPA=4 THEN 1.04 ELSE 1.19 END )
330         DBGrid1.Visible = False
340         MSFlexGrid1.Visible = False
350         MSFlexGrid2.Visible = True
            MSFlexGrid2.Col = 2
            MSFlexGrid2.row = 2
            ' .row = 0
            '.Col = 0
            '.ColSel = .ColS - 1
            ' = .BackColorFixed
            ' MSFlexGrid1.CellBackColor = MSFlexGrid1.TextMatrix(4, 2) = vbRed
            ' MSFlexGrid2.CellBackColor = vbYellow
360         MSFlexGrid2.AddItem " " & Chr(9)
370         SA = 0
380         SP = 0
390         k = 0
400         SP2 = 0

            Dim SEP

410         SEP = 0

            Dim SYN

            Dim SXAN1, SXAN2
        
420         SXAN1 = 0
430         SXAN2 = 0

            Dim DS(30) As Single

            Dim L3 As Long

440         For k = 1 To MSFlexGrid2.rows - 2
    
450             For L3 = 1 To MSFlexGrid2.ColS - 1
460                 DS(L3) = DS(L3) + Val(MSFlexGrid2.TextMatrix(k, L3))
470             Next L3

                'ΓΙΑ ΝΑ ΒΓΑΖΕΙ ΤΑ ΣΥΝΟΛΑ ΜΕ ΔΕΚΑΔΙΚΑ
                '    MSFlexGrid2.TextMatrix(k, MSFlexGrid2.ROWS - 1) = Format(MSFlexGrid2.TextMatrix(k, MSFlexGrid2.ROWS - 1), "####,##0.00")    'SEP
480         Next k

            '        For k = 1 To MSFlexGrid2.ROWS - 2
            '
            '            SEP = SEP + Val(MSFlexGrid2.TextMatrix(k, 1))
            '            SA = SA + Val(MSFlexGrid2.TextMatrix(k, 2))
            '            SP = SP + Val(MSFlexGrid2.TextMatrix(k, 3))
            '            SP2 = SP2 + Val(MSFlexGrid2.TextMatrix(k, 4))
            '
            '            SXAN1 = SXAN1 + Val(MSFlexGrid2.TextMatrix(k, 5))
            '            SXAN2 = SXAN2 + Val(MSFlexGrid2.TextMatrix(k, 6))
            '
            '
            '            SYN = SYN + Val(MSFlexGrid2.TextMatrix(k, 7))
            '        Next
            Dim L

490         For L = 1 To MSFlexGrid2.ColS - 1
500             MSFlexGrid2.ColAlignment(L) = 6    'RIGHT
510             MSFlexGrid2.TextMatrix(k, L) = Format(DS(L), "####,##0.00")    'SEP
            Next

            '        MSFlexGrid2.TextMatrix(k, 1) = Format(SEP, "####,##0.00")    'SEP
            '        MSFlexGrid2.TextMatrix(k, 2) = Format(SA, "####,##0.00")    'SA
            '        MSFlexGrid2.TextMatrix(k, 3) = Format(SP, "####,##0.00")    'SP
            '        MSFlexGrid2.TextMatrix(k, 4) = Format(SP2, "####,##0.00")    'SP2
            '
            '        MSFlexGrid2.TextMatrix(k, 5) = Format(SXAN1, "####,##0.00")    'SP2
            '        MSFlexGrid2.TextMatrix(k, 6) = Format(SXAN2, "####,##0.00")
            '        MSFlexGrid2.TextMatrix(k, 7) = Format(SYN, "####,##0.00")

            On Error Resume Next

520         MSFlexGrid2.TopRow = MSFlexGrid2.rows - 10 + 1

            'Else
            '  d = "select HME as shme,sum(AJI) as saji "
            '  d = d + "from TIM  where '" + Right(Combo1.Text, 1) + "'=left(ATIM,1) "
            '
            '
            '  d = d + "and HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "' "
            '  d = d + "group by HME;"
            '  data1.RecordSource = d
            '
            '  data1.Refresh
            '  DBGrid1.Columns(1).Caption = "Σύνολο με ΦΠΑ"
            '  DBGrid1.Columns(0).Caption = "Ημερομηνία"
            '  DBGrid1.Columns(1).NumberFormat = "##,###,##0.00"
            '
            '

        End If

        Dim i As Long

        Dim j As Long

        If Check1 Then    ' palia xronia

            For i = 0 To Data1.Recordset.RecordCount - 1
                For j = 1 To MSFlexGrid2.ColS Step 2
                    MSFlexGrid2.row = i + 1
                    MSFlexGrid2.Col = j
                    MSFlexGrid2.CellBackColor = vbYellow
                Next
            Next

        End If

530     Me.MousePointer = vbNormal

        '<EhFooter>
        Exit Sub

Command6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.Command6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.Command6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub YPOL_PALIA_TAMEIAKES()

        Dim d1   As Date

        Dim d2   As Date

        Dim GDS  As New ADODB.Connection

        Dim CONN As String

        On Error GoTo 0

        CONN = gConnect + ";DATABASE=" + f_PALIAXRONIA ' "DSN=" + LTrim(Str(f_PALIAXRONIA)) ' = Val(FindParametroi(1,"PAR1", "f_PALIAXRONIA", "2005", "Σειρές τιμ.Πώλησης"))
        d1 = DTPicker1.Value - 365
        d2 = DTPicker2.Value - 365

        GDS.Open CONN

        Dim MSYNT As String

        '190     If Check2 Then

200     If apot.ListIndex = -1 Then
210         MSYNT = ""
        Else
220         MSYNT = " APOT=" + str(1 + apot.ListIndex) + " AND "
        End If

        On Error Resume Next

230     GDS.Execute "DROP TABLE DOKSYNOLA"
240     GDS.Execute "UPDATE EGGTIM SET EKPT=0 WHERE EKPT IS NULL"
        'On Error GoTo Command6_Click_Err

250     If f_TimMeFPA = 1 Then

            Dim nc As Long

260         GDS.Execute "SELECT  ROW_NUMBER() OVER( ORDER BY CONVERT(CHAR(12),HME,3) ) AS AA, CONVERT(CHAR(12),HME,3) AS HMER ," _
               & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=1 AND LEFT(ATIM,1) IN ('L') THEN POSO*TIMM*(1-EKPT/100) ELSE 0 END )) AS [TAM.1]," _
               & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=1 AND LEFT(ATIM,1) IN ('T','t') THEN MIK_AJIA ELSE 0 END)) AS [TIM.1]," _
               & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=1 AND LEFT(ATIM,1) IN ('l','R','μ') THEN MIK_AJIA  ELSE 0 END)) AS [in.1]," _
               & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=2 AND LEFT(ATIM,1) IN ('L') THEN MIK_AJIA  ELSE 0 END)) AS [TAM.2]," _
               & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=2 AND LEFT(ATIM,1) IN ('l','R','μ') THEN MIK_AJIA  ELSE 0 END)) AS [in.2]," _
               & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=2 AND LEFT(ATIM,1) IN ('T','t') THEN MIK_AJIA ELSE 0 END )) AS [TIM.2]," _
               & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=3 AND LEFT(ATIM,1) IN ('L') THEN MIK_AJIA  ELSE 0 END)) AS [TAM.3]," _
               & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=3 AND LEFT(ATIM,1) IN ('l','R','μ') THEN MIK_AJIA  ELSE 0 END)) AS [in.3]," _
               & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=3 AND LEFT(ATIM,1) IN ('T','t') THEN MIK_AJIA ELSE 0 END )) AS [TIM.3],0 AS [ΣΥΝΟΛΑ]" _
               & " INTO DOKSYNOLA from EGGTIM WHERE " _
               + MSYNT + "  HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2 + 1, "mm/dd/yyyy") + "' " _
               & "  GROUP BY CONVERT(CHAR(12),HME,3) ORDER BY CONVERT(CHAR(12),HME,3)", nc

        Else
270         GDS.Execute "SELECT ROW_NUMBER() OVER( ORDER BY CONVERT(CHAR(12),HME,3) ) AS AA, CONVERT(CHAR(12),HME,3) AS HMER ," _
               & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=1 AND LEFT(ATIM,1) IN ('L') THEN PIS*TIMM ELSE 0 END )) AS [TAM.1]," _
               & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=1 AND LEFT(ATIM,1) IN ('T','t') THEN PIS*TIMM*(1-EKPT/100) ELSE 0 END)) AS [TIM.1]," _
               & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=3 AND LEFT(ATIM,1) IN ('L') THEN PIS*TIMM ELSE 0 END)) AS [TAM.3]," _
               & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=3 AND LEFT(ATIM,1) IN ('T','t') THEN POSO*TIMM*(1-EKPT/100) ELSE 0 END )) AS [TIM.3]," _
               & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=2 AND LEFT(ATIM,1) IN ('L') THEN PIS*TIMM ELSE 0 END)) AS [TAM.2]," _
               & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN APOT=2 AND LEFT(ATIM,1) IN ('T','t') THEN POSO*TIMM*(1-EKPT/100) ELSE 0 END )) AS [TIM.2],0 AS [ΣΥΝΟΛΑ]" _
               & " INTO DOKSYNOLA from EGGTIM WHERE " _
               + MSYNT + "  HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2 + 1, "mm/dd/yyyy") + "' " _
               & "  GROUP BY CONVERT(CHAR(12),HME,3) ORDER BY CONVERT(CHAR(12),HME,3)"
        End If

280     GDS.Execute "UPDATE DOKSYNOLA SET [ΣΥΝΟΛΑ]=[TAM.1]+[TIM.1]+[TAM.2]+[TIM.2]+[TAM.3]+[TIM.3]+[in.1]+[in.2]+[in.3]"

End Sub




Sub PAR2SHOWING()

        '<EhHeader>
        On Error GoTo PAR2SHOWING_Err

        '</EhHeader>

        Dim FF As PAR2

100
110     Set FF = New PAR2

120

130     DBGrid1.Col = 1
140     FF.DTPicker1 = CDate(DBGrid1.Text)
150     FF.DTPicker2 = CDate(DBGrid1.Text)

 DBGrid1.Col = 0
  
160   FF.Text1 = Trim(Left(DBGrid1.Text, 9))

FF.CallFromOther.Caption = "0"

        'GIA OLA TA PARASTATIKA
170     FF.Combo2.Text = "*"
        
       FF.Command1_Click
       
      FF.SHOW_TIMOL

180     FF.SHOW

 FF.TIM.Move 0
        FF.SHOW_TIMOL
190     FF.bohuitiko.SetFocus

200     DoEvents

210     FF.Command1.SetFocus
220    ' PAR1.LASTTIMOL.Refresh

        'On Error GoTo 0

230     '   SendKeys "{enter}"    'ENTER  ~

        '<EhFooter>
        Exit Sub

PAR2SHOWING_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.PAR2SHOWING " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.PAR2SHOWING " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub




Private Sub DATES_Click()

Dim a As Integer
'0 Σήμερα
'1 Χθές
'2 Τρέχ.Εβδομάδα
'3 Τρέχ.Μήνας
'4 Τρέχ.Τρίμηνο
'5 Προηγ.Μήνας
'6  Προηγ.3Μήνας
'7 τρεχ ετος
'8 ΠΡΟΗΓ ΕΤΟΣ

 a = DATES.ListIndex
 Dim d1 As DTPicker, d2 As DTPicker
  Dim MHNAS As Integer
 Select Case a
     Case 0
       DTPicker1 = gWorkDay 'Format(gWorkDay, "dd/MM/yyyy")
       DTPicker2 = gWorkDay ' Format(gWorkDay, "dd/MM/yyyy")
     Case 1
       DTPicker1 = DateAdd("d", -1, gWorkDay) ' Format(DateAdd("d", -1, gWorkDay), "dd/MM/yyyy")
       DTPicker2 = DateAdd("d", -1, gWorkDay) ' Format(DateAdd("d", -1, gWorkDay), "dd/MM/yyyy")
     Case 2
       DTPicker1 = Format(DateAdd("d", -Weekday(gWorkDay) + 1, gWorkDay), "dd/MM/yyyy")
       DTPicker2 = Format(DateAdd("d", -1, gWorkDay), "dd/MM/yyyy")
     
     Case 3  '3 Τρέχ.Μήνας
     
        DTPicker1 = Format(DateSerial(Year(gWorkDay), Month(gWorkDay), 1), "dd/MM/yyyy")
        
        
        
        If Month(gWorkDay) = 12 Then
           DTPicker2 = Format(DateSerial(Year(gWorkDay), Month(gWorkDay), 31), "dd/MM/yyyy")
        Else
           DTPicker2 = Format(DateSerial(Year(gWorkDay), Month(gWorkDay) + 1, 0), "dd/MM/yyyy")
        End If
        
     
     Case 4 ' 4 Τρέχ.Τρίμηνο
         
          MHNAS = Month(gWorkDay)
          If MHNAS < 3 Then MHNAS = 3
          If MHNAS > 3 And MHNAS <= 6 Then MHNAS = 6
          If MHNAS > 6 And MHNAS <= 9 Then MHNAS = 9
          If MHNAS > 9 Then MHNAS = 12
          
          DTPicker1 = Format(DateSerial(Year(gWorkDay), MHNAS - 2, 1), "dd/MM/yyyy")
          
          DTPicker2 = Format(DateSerial(Year(gWorkDay), IIf(MHNAS = 12, 12, MHNAS + 1), 0), "dd/MM/yyyy")
         
     
     
     Case 5 '5 Προηγ.Μήνας
     
        DTPicker2 = Format(DateSerial(Year(gWorkDay), Month(gWorkDay), 0), "dd/MM/yyyy")
        DTPicker1 = Format(DTPicker2 - Day(DTPicker2) + 1, "dd/MM/yyyy")
     
     
     ' DateSerial(Year(gWorkDay), Month(gWorkDay), 0) = 28 / 2 / 2013
     ' DateSerial(Year(gWorkDay), 1, 0) = 31 / 12 / 2012
     
     
     Case 6 ' 4 Τρέχ.Τρίμηνο
          
          MHNAS = Month(gWorkDay)
          If MHNAS < 3 Then
             MHNAS = 12
             
             DTPicker1 = Format(DateSerial(Year(gWorkDay) - 1, MHNAS - 2, 1), "dd/MM/yyyy")
          
             DTPicker2 = Format(DateSerial(Year(gWorkDay) - 1, IIf(MHNAS = 12, 12, MHNAS + 1), 31), "dd/MM/yyyy")
          
          
          
          
          Else
          
             If MHNAS > 3 And MHNAS <= 6 Then MHNAS = 3
             If MHNAS > 6 And MHNAS <= 9 Then MHNAS = 6
             If MHNAS > 9 Then MHNAS = 9
          
             DTPicker1 = Format(DateSerial(Year(gWorkDay), MHNAS - 2, 1), "dd/MM/yyyy")
          
             DTPicker2 = Format(DateSerial(Year(gWorkDay), IIf(MHNAS = 12, 12, MHNAS + 1), 0), "dd/MM/yyyy")
          End If
          
     Case 7 ' 7Τρέχ.Ετος
          
     
          
             DTPicker1 = Format(DateSerial(Year(gWorkDay), 1, 1), "dd/MM/yyyy")
             DTPicker2 = Format(DateSerial(Year(gWorkDay), 12, 31), "dd/MM/yyyy")
          
             'DTPicker2 = gWorkDay 'Format(gWorkDay, "dd/MM/yyyy")
        
      Case 8 ' 7Τρέχ.Ετος
          
     
          
             DTPicker1 = Format(DateSerial(Year(gWorkDay) - 1, 1, 1), "dd/MM/yyyy")
             DTPicker2 = Format(DateSerial(Year(gWorkDay) - 1, 12, 31), "dd/MM/yyyy")
          
             'DTPicker2 = gWorkDay 'Format(gWorkDay, "dd/MM/yyyy")
     
     
 
 End Select


End Sub

Private Sub DBGrid1_DblClick()
Dim a As String
a = ""
PAR2SHOWING


End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)

 '<EhHeader>
        On Error GoTo dbgrid1_HeadClick_Err

        '</EhHeader>

        '
        '  Dim strSort As String
        '
        '  Dim Col As TrueOleDBGrid80.Column
        '
        '
        '
        '' Loop through GroupColumns collection and construct
        '
        '' the sort string for the Sort property of the Recordset
        '
        '    For Each Col In dbgrid1.Columns
        '
        '        If strSort <> vbNullString Then
        '
        '            strSort = strSort & ", "
        '
        '        End If
        '
        '        strSort = strSort & "[" & Col.DataField & "]"
        '
        '    Next Col
        '
        '

        'dbgrid1.HoldFields
        Dim sumes(100) As String, k As Integer
        

100     For k = 0 To Data1.Recordset.FIELDS.Count - 1
110         sumes(k) = DBGrid1.Splits(0).columns(k).FooterText    '  = Format(SUMES(k), "######0.00")
        Next

120     If Data1.Recordset.sort = "[" & DBGrid1.columns(ColIndex).DataField & "] asc" Then   ' strSort
130         Data1.Recordset.sort = "[" & DBGrid1.columns(ColIndex).DataField & "] desc"    ' strSort
        Else
140         Data1.Recordset.sort = "[" & DBGrid1.columns(ColIndex).DataField & "] asc"    ' strSort
        End If

150     For k = 0 To Data1.Recordset.FIELDS.Count - 1
160         DBGrid1.Splits(0).columns(k).FooterText = sumes(k)   '  = Format(SUMES(k), "######0.00")
        Next

        '    data1.Recordset.Sort = "[" & dbgrid1.Columns(ColIndex).DataField & "] desc" ' strSort

        '<EhFooter>
        Exit Sub

dbgrid1_HeadClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.dbgrid1_HeadClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.dbgrid1_HeadClick " & "at line " & Erl

        Resume Next



End Sub

Private Sub exit_Click()

        '<EhHeader>
        On Error GoTo exit_Click_Err

        '</EhHeader>

100     Unload Me
110     SendKeys "%"
120     SendKeys "{RIGHT}+{RIGHT}+{RIGHT}+{DOWN}+{DOWN}+{DOWN}+{DOWN}"

        '<EhFooter>
        Exit Sub

exit_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.exit_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.exit_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "PAR4"
110     PARAMETROI.SHOW 1

        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.Form_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then exit_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
        Dim DB As Database

        Dim R  As New ADODB.Recordset

100     mForm_Load Me, fh, fw, ft, fl
    
110     F928 = Val(FINDPARAMETROI(1, "PAR1", "F928", "0", "Α.11 Εκτύπωση φόρμας 1=928 0=437"))

        ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        '   ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        '    'SkinFramework1.ApplyWindow Me.hWnd
        '    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

        'tropos plhromhs
        'R.Close
        Dim k As Long
        
        
        
        f_arxiko_platos_Grid1 = DBGrid1.Width
        
       
        
        

120     R.Open "SELECT *FROM PINAKES WHERE TYPOS=12 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
130     k = 0

140     Do While Not R.EOF

150         If R("typos") = 12 Then
160             metrhta.AddItem str(R("AYJON")) + "." + R("PERIGRAFH")
                'F_LOGPLHR(K) = Format(R("TIMH"), "0000")
170             k = k + 1
            End If

180         R.MoveNext
        Loop

190     R.Close

        'ΦΟΡΤΩΜΑ ΑΠΟΘΗΚΩΝ
200     R.Open "SELECT *FROM PINAKES WHERE TYPOS=4 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

210     Do While Not R.EOF
220         apot.AddItem LTrim(str(R("AYJON"))) + "." + R("PERIGRAFH")
230         R.MoveNext
        Loop

240     R.Close
250     apot.Text = apot.List(0)

        'ΦΟΡΤΩΜΑ ΠΩΛΗΤΩΝ
260     R.Open "SELECT *FROM PINAKES WHERE TYPOS=5 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

270     Do While Not R.EOF
280         Combo8.AddItem LTrim(str(R("AYJON"))) + "." + R("PERIGRAFH")
290         R.MoveNext
        Loop

300     R.Close

310     Combo8.Text = ""

320     Me.KeyPreview = True

330     Me.Picture = LoadPicture(gPicture)

        'Set db = OpenDatabase(gDir, False, False, gConnect)

340     R.Open UCase("select pol,titlos,eidos from parastat ORDER BY POL,TITLOS;"), Gdb, adOpenDynamic, adLockOptimistic

350     Combo1.ComboItems.Clear

        'Combo1.AddItem "Πωλήσεις"
        'Combo1.AddItem "Αγορές"

360     R.MoveFirst
370     k = 0
        Dim combo1stop As String: combo1stop = ""
380     Do While Not R.EOF

390         If Not IsNull(R("TITLOS")) And Not IsNull(R("pol")) Then
        
400             If R("POL") = 2 Then  ' αγορες
410                 Combo1.ComboItems.ADD k + 1, "", Format(k + 1, "##") + "---" + " " + R("TITLOS") + Space(1) + R("eidos"), 6
                    ' Combo2.ComboItems(k).Image = 6
                Else
             
420                 Combo1.ComboItems.ADD k + 1, "", Format(k + 1, "##") + "---" + " " + R("TITLOS") + Space(1) + R("eidos"), 3
                    ' Combo2.ComboItems(k).Image = 3
                    If InStr(R("TITLOS"), "Π ω λ") > 0 Then
                       combo1stop = Format(k + 1, "##") + "---" + " " + R("TITLOS") + Space(1) + R("eidos")
                    End If
                    
                    
                    
                    
                End If

430             k = k + 1
                '           Combo1.AddItem R("TITLOS") + Space(30) + R("eidos")
            End If

440         R.MoveNext
        Loop

450     R.Close

        Combo1.Text = combo1stop
        


        'GEMISMA MULTI SELECT LISTBOX
460     R.Open "select POL,TITLOS,EIDOS,AJIA_APOU from PARASTAT where POL=1 ORDER BY POL,TITLOS;", Gdb, adOpenDynamic, adLockOptimistic
470     List1.Clear
480     R.MoveFirst

490     k = 0

500     Do While Not R.EOF

510         If Not IsNull(R("TITLOS")) And Left(R("TITLOS"), 1) <> "-" And (R("AJIA_APOU") = "3" Or R("AJIA_APOU") = "4") Then
520             If R("POL") = 1 Then
530                 List1.AddItem R("TITLOS") + Space(1) + R("eidos")
540                 List1.Selected(k) = True
                End If

550             k = k + 1
                'mEidos_Par(K) = R("eidos")
            End If

560         R.MoveNext
        Loop

570     R.Close

580     Combo1.Enabled = True
590     'Combo1.Text = Combo1.ComboItems(geidos_paras + 1)

        'Data1.DatabaseName = gDir
600     Data1.ConnectionString = gConnect
610     DTPicker1.Value = CDate("01/01/" + str(Year(Now)))
620     DTPicker2.Value = Now
630     apot.Text = apot.List(0)

        ' f_PALIAXRONIA = Val(FindParametroi(1,"PAR1", "F_PALIAXRONIA", "DELETE", "1=ΕΛΕΓΧOΣ 2=ΑΠΑΓΟΡΕΥΣΗ 0=ΤΙΠΟΤΑ"))

640     f_PALIAXRONIA = FINDPARAMETROI(1, "PAR1", "F_PALIAXRONIA", "2005", "Γ.18 ΒΑΣΗ ΠΑΛΙΑΣ ΧΡΟΝΙΑΣ")
650     f_PROPERSI = FINDPARAMETROI(1, "PAR1", "F_PROPERSI", "2005", "Γ.19 ΒΑΣΗ ΠΡΟΠΕΡΣΙΝΗΣ ΧΡΟΝΙΑΣ")
660     f_TimMeFPA = FINDPARAMETROI(1, "PAR4", "F_TIMMEFPA", "0", "Στα στατιστικά Τα ΤΙΜ.ΠΩΛΗΣΗΣ εχουν φπα=1 ")




        DATES.Text = DATES.List(2)
        DATES_Click
        



        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 13950, 9750, fh, fw, ft, fl
    f_arxiko_platos_Grid1 = DBGrid1.Width
    
    Dim L As Long
    L = DATES.Height
    DTPicker1.Height = L
    DTPicker2.Height = L
    DATES.Height = L
    metrhta.Height = L
    Combo2.Height = L
    
    
    
    

End Sub

Private Sub Label3_Click()

        '<EhHeader>
        On Error GoTo Label3_Click_Err

        '</EhHeader>

100     UPDATE_PINAKES 4, Label3, apot

        '<EhFooter>
        Exit Sub

Label3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.Label3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.Label3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Label4_Click()

        '<EhHeader>
        On Error GoTo Label4_Click_Err

        '</EhHeader>

100     UPDATE_PINAKES 5, Label4, Combo8

        '<EhFooter>
        Exit Sub

Label4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.Label4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.Label4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub List1_DblClick()

        '<EhHeader>
        On Error GoTo List1_DblClick_Err

        '</EhHeader>
        Dim k As Integer

100     For k = 0 To List1.ListCount - 1
110         List1.Selected(k) = False
        Next

        '<EhFooter>
        Exit Sub

List1_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.List1_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.List1_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub restore_width_Click()
   DBGrid1.Width = f_arxiko_platos_Grid1
End Sub

Private Sub ShowPar_Click()

        '<EhHeader>
        On Error GoTo ShowPar_Click_Err

        '</EhHeader>

        Dim D As String, DB As Database

        Dim R As New ADODB.Recordset, pol As String, ago As String

        Dim mChoose, mSaji2

        Dim a, polepis As String, AGOEPIS As String

        Dim synt

        Dim D0 As String

100     a = Get_AJ_ASCII(pol, polepis, ago, AGOEPIS)

110     DBGrid1.Visible = True
120     MSFlexGrid1.Visible = False

        'If Combo1.SelectedItem.Index = -1 Then Combo1.SelectedItem.Index = 0

130     If InStr(Combo1.Text, "Π ω λ ή σ ε ι ς") > 0 Then
140         mChoose = pol + "," + polepis
150     ElseIf InStr(Combo1.Text, "Α γ ο ρ έ ς") > 0 Then  ' Α γ ο ρ έ ς -------- θ
160         mChoose = ago + "," + AGOEPIS
        Else
170         mChoose = str(Asc(Right(Combo1.Text, 1)))
        End If

180     Me.MousePointer = vbHourglass
190     D = "select LEFT(ATIM,8)+' '+(SELECT top 1 LEFT(TITLOS,15) FROM PARASTAT WHERE ASCII(PARASTAT.EIDOS)=ASCII(LEFT(ATIM,1))) AS [ΠΑΡ/ΚΟ],"
        D = D + " CONVERT(CHAR(8),HME,3) as [Ημερ/νία]," & " LEFT((case when AKYROMENO=1 THEN '**' ELSE '' END)+PEL.EPO,25) as [Συναλασσόμενος],"
        D = D + "(CASE WHEN ascii(LEFT(ATIM,1)) IN (" + polepis + "," + AGOEPIS + ") THEN -AJI ELSE ROUND(AJI,2) END) as [Συν.Αξία]," '4
        D = D + " (CASE WHEN ASCII(LEFT(ATIM,1)) IN (" + polepis + "," + AGOEPIS + ") THEN -AJ1 ELSE ROUND(AJ1,2) END) as [Καθ(1) " + Format(g_Fpa(1), "##") + "%]," '5
        D = D + "(CASE WHEN ASCII(LEFT(ATIM,1)) IN (" + polepis + "," + AGOEPIS + ") THEN -AJ2 ELSE ROUND(AJ2,2) END) as  [Καθ(2) " + Format(g_Fpa(2), "##") + "%]," '6
        
        D = D + "(CASE WHEN ASCII(LEFT(ATIM,1)) IN (" + polepis + "," + AGOEPIS + ") THEN -AJ3 ELSE ROUND(AJ3,2) END) as  [Καθ(3)" + Format(g_Fpa(3), "##") + "%],"  '7
        D = D + "(CASE WHEN ASCII(LEFT(ATIM,1)) IN (" + polepis + "," + AGOEPIS + ") THEN -AJ4 ELSE ROUND(AJ4,2) END) as  [Καθ(4)" + Format(g_Fpa(4), "##") + "%]," '8
        
        D = D + "(CASE WHEN ASCII(LEFT(ATIM,1)) IN (" + polepis + "," + AGOEPIS + ") THEN -AJ5 ELSE round(AJ5,2) END) as  [Καθ(5)" + Format(g_Fpa(5), "##") + "%]," '9
        
         D = D + "(CASE WHEN ASCII(LEFT(ATIM,1)) IN (" + polepis + "," + AGOEPIS + ") THEN -AJ6 ELSE round(AJ6,2) END) as  [Καθ(6)" + Format(g_Fpa(6), "##") + "%]," '10
         
          D = D + "(CASE WHEN ASCII(LEFT(ATIM,1)) IN (" + polepis + "," + AGOEPIS + ") THEN -AJ7 ELSE round(AJ7,2) END) as  [Καθ(7)" + Format(g_Fpa(7), "##") + "%],"  '11
        
          
        'D = D + " AJ3  AS [Κ.Αξ." + Format(g_Fpa(3), "##") + "%],AJ4 AS [κ.Αξ." + Format(g_Fpa(4), "##.#") + "%],AJ5 as [Κ.Αξ.0%],"
        'D = D + " AJ6 as [K.Αξ." + Format(g_Fpa(6), "##") + "%],"
        'D = D + " AJ7 as [K.Αξ." + Format(g_Fpa(7), "##") + "%],"
        
        D = D + " round(FPA1+FPA2+FPA3+FPA4+FPA6+FPA7,2)  as [ΣΥΝ.ΦΠΑ],SKOPOS as [Σκοπός],"   '12  ,13
        D = D + " SXETIKO AS [ΣΧΕΤΙΚΟ],round(KERDOS,2) AS [ΚΕΡΔΟΣ],(case when left(TRP,1)='1' THEN AJI ELSE 0 end ) AS [ΜΕΤΡΗΤ],"
        D = D + " (case when left(TRP,1)='1' THEN 0 ELSE AJI end) AS [ΠΙΣΤΩΣΗ] ,ENTITYMARK AS MAPK " 'into temp" + LTrim(Str(gXEIRISTHS))
200     D = D + " from TIM LEFT join PEL on TIM.EIDOS=PEL.EIDOS and TIM.KPE=PEL.KOD "
210     D0 = "select SUM(CASE WHEN ascii(LEFT(ATIM,1)) IN(" + polepis + "," + AGOEPIS + ") THEN -AJI ELSE round(AJI,2) END ) AS SAJI,"
         D0 = D0 + "SUM(CASE WHEN ASCII(LEFT(ATIM,1)) IN(" + polepis + "," + AGOEPIS + ") THEN -AJ1 ELSE round(AJ1,2) END) AS SAJ1,"
         D0 = D0 + "SUM(CASE WHEN ASCII(LEFT(ATIM,1)) IN(" + polepis + "," + AGOEPIS + ") THEN -AJ2 ELSE round(AJ2,2) END) AS SAJ2,"
         D0 = D0 + "SUM(CASE WHEN ASCII(LEFT(ATIM,1)) IN(" + polepis + "," + AGOEPIS + ") THEN -AJ3 ELSE round(AJ3,2) END) AS SAJ3,"
         D0 = D0 + "SUM(CASE WHEN ASCII(LEFT(ATIM,1)) IN(" + polepis + "," + AGOEPIS + ") THEN -AJ4 ELSE round(AJ4,2) END) AS SAJ4,"
         D0 = D0 + "SUM(CASE WHEN ASCII(LEFT(ATIM,1)) IN(" + polepis + "," + AGOEPIS + ") THEN -AJ5 ELSE round(AJ5,2) END) AS SAJ5,"
         D0 = D0 + "SUM(CASE WHEN ASCII(LEFT(ATIM,1)) IN(" + polepis + "," + AGOEPIS + ") THEN -AJ6 ELSE round(AJ6,2) END) AS SAJ6,"
         D0 = D0 + "SUM(CASE WHEN ASCII(LEFT(ATIM,1)) IN(" + polepis + "," + AGOEPIS + ") THEN -AJ7 ELSE round(AJ7,2) END) AS SAJ7,"
         D0 = D0 + "SUM(CASE WHEN ASCII(LEFT(ATIM,1)) IN(" + polepis + "," + AGOEPIS + ") THEN -ABS(FPA1+FPA2+FPA3+FPA4+FPA6+FPA7)  ELSE round(ABS(FPA1+FPA2+FPA3+FPA4+FPA6+FPA7),2) END) AS SFPA,"
         D0 = D0 + "round(sum(KERDOS),2) AS SK,"
         D0 = D0 + "SUM(case when TRP like '1%' THEN round(AJI ,2) ELSE 0 END ) AS METR,"
         D0 = D0 + "SUM(case when left(TRP,1)<>'1' THEN round(AJI,2) ELSE 0 END ) AS PIST  "
         D0 = D0 + "FROM TIM "

220     synt = ""

230     If Combo2.Text = Combo2.List(0) Then
            ' synt = "ART IS NULL AND "   '   OR LEFT(ART,1)=' '"
240         synt = " ( ART IS NULL OR ART='') AND "
        End If

250     If Combo2.Text = Combo2.List(1) Then
260         synt = "LEFT(ART,1)='1' AND "
        End If

270     If Len(Combo8.Text) > 0 Then
280         synt = synt + " LEFT(EIDPAR,1)='" + Left(Combo8, 1) + "' AND "
        End If

        If akyromena.Value = vbChecked Then
            synt = synt + " LEFT(ATIM,1) NOT IN (SELECT AKYR FROM PARASTAT WHERE AKYR IS NOT NULL ) AND ISNULL(AKYROMENO,0)=0  AND "
        End If






290     If Len(metrhta.Text) > 0 Then
300         If metrhta.Text = metrhta.List(0) Then
310             synt = synt + " TRP='ΜΕ' " + " AND "
            Else
320             synt = synt + " TRP<>'ΜΕ' " + " AND "
            End If
        End If

330     'If APOT.ListIndex <= 0 Then


Dim Q_APOT As String

If Val(Left(apot.Text, 2)) = 0 Or Val(Left(apot.Text, 2)) = 99 Then
    Q_APOT = " "
Else
     Q_APOT = " B_N1=" + Left(apot.Text, 1) + " AND "
End If


        'Else
340     If Len(synt) = 0 Then
350         synt = Q_APOT
        Else
360         synt = synt + Q_APOT '  "  B_N1=" + Left(apot.Text, 1) + " AND "
        End If

        'End If

370     If Fsql = "" Then
380         D = D + "where " + synt + " ascii(left(ATIM,1)) IN (" + mChoose + ") and HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "' order by HME,ATIM;"
390         D0 = D0 + "where " + synt + " ascii(left(ATIM,1)) IN (" + mChoose + ") and HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "'"
        Else
400         D = D + "where " + synt + " ascii(left(ATIM,1)) IN (" + mChoose + ") and HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "' and " + Fsql + " order by HME,ATIM;"
410         D0 = D0 + "where " + synt + " ascii(left(ATIM,1)) IN (" + mChoose + ") and HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "' and " + Fsql + " "

        End If

        'db.Execute d

        On Error Resume Next

        'Gdb.Execute "drop table TEMP" + LTrim(Str(gXEIRISTHS))
        'Gdb.Execute d

        On Error GoTo ShowPar_Click_Err
        
        
         R.Open D0, Gdb, adOpenForwardOnly, adLockReadOnly
        
        
Dim ll As Integer

For ll = 4 To 10
   'DBGrid1.columns(ll).FooterText = Format(R("SAJ" + Format(ll - 3, "0")), "######0.00")
   If R("SAJ" + Format(ll - 3, "0")) = 0 Then
        DBGrid1.columns(ll).Width = 0
   Else
        DBGrid1.columns(ll).Width = 800
        
   End If
   
   
Next
        
     DBGrid1.Refresh
     
     
        

420     Data1.RecordSource = D    ' "select sum(AJI) as saji2 from TEMP" + LTrim(Str(gXEIRISTHS))

        On Error Resume Next

430     Data1.Refresh

        'mSaji2 = data1.Recordset("saji2")

        'data1.RecordSource = "select *from TEMP" + LTrim(Str(gXEIRISTHS))

        'data1.Refresh
        'data1.Recordset.AddNew
        'data1.Recordset("aji") = mSaji2
        'data1.Recordset("epo") = "Σ Υ Ν Ο Λ Α "
        'data1.Recordset.Update

        ' Dim R As New ADODB.Recordset
        On Error Resume Next

440    ' R.Open D0, Gdb, adOpenForwardOnly, adLockReadOnly

        'data1.Refresh
        On Error Resume Next

450     DBGrid1.columns(3).alignment = dbgRight
460     DBGrid1.columns(3).FooterText = Format(R!SAJI, "######0.00")

For ll = 4 To 10
   
        DBGrid1.columns(ll).FooterText = Format(R("SAJ" + Format(ll - 3, "0")), "######0.00")
   
   
   
Next
     


'470     DBGrid1.columns(4).FooterText = Format(R!SAJ1, "######0.00")
'480     DBGrid1.columns(5).FooterText = Format(R!SAJ2, "######0.00")
'490     DBGrid1.columns(6).FooterText = Format(R!SAJ3, "######0.00")
'500     DBGrid1.columns(7).FooterText = Format(R!SAJ4, "######0.00")
'510     DBGrid1.columns(8).FooterText = Format(R!SAJ5, "######0.00")
'520     DBGrid1.columns(9).FooterText = Format(R!SAJ6, "######0.00")
'530     DBGrid1.columns(10).FooterText = Format(R!SAJ7, "######0.00")
531     DBGrid1.columns(11).FooterText = Format(R!sfpa, "######0.00")


532     DBGrid1.columns(14).FooterText = Format(R!SK, "######0.00")
534     DBGrid1.columns(15).FooterText = Format(R!METR, "######0.00")
536     DBGrid1.columns(16).FooterText = Format(R!pist, "######0.00")

540    ' DBGRID1.columns(13).FooterText = Format(R(9), "######0.00")
550    ' DBGRID1.columns(14).FooterText = Format(R(8), "######0.00")

560     DBGrid1.columns(0).Caption = "Παραστατικό"
570     DBGrid1.columns(1).Caption = "Ημερομηνία"
580     DBGrid1.columns(2).Caption = "Πελάτης /Προμηθευτής"
590     DBGrid1.columns(3).Caption = "Αξία με ΦΠΑ"

600     DBGrid1.columns(0).Width = 1000
610     DBGrid1.columns(1).Width = 1000
620     DBGrid1.columns(2).Width = 3800
630     DBGrid1.columns(3).Width = 800
640     DBGrid1.columns(3).NumberFormat = "##,###,##0.00"
650     Me.MousePointer = vbNormal

660     If ekt Then
            'typos2
670         print3_xar D, "001111111111111111", Combo1.Text, 0

        End If

        '<EhFooter>
        Exit Sub

ShowPar_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.ShowPar_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.ShowPar_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub OLD_ShowPar_Click()
'
''    Dim d As String, db As Database
''    Dim R As New ADODB.Recordset, pol As String, ago As String
''    Dim mChoose, mSaji2
''    Dim a, polepis As String, AGOEPIS As String
''    Dim synt
''    Dim D0 As String
''
''    a = Get_AJ(pol, polepis, ago, AGOEPIS)
''
''    DBGrid1.Visible = True
''    MSFlexGrid1.Visible = False
''
''
''    If Combo1.ListIndex = -1 Then Combo1.ListIndex = 0
''
''
''
''    If Combo1.ListIndex = 0 Then
''        mChoose = pol + "," + polepis
''    ElseIf Combo1.ListIndex = 1 Then
''        mChoose = ago
''    Else
''        mChoose = "'" + Right(Combo1.Text, 1) + "'"
''    End If
''
''
''
''
''    Me.MousePointer = vbHourglass
''    d = "select LEFT(ATIM,8)+' '+(SELECT LEFT(TITLOS,15) FROM PARASTAT WHERE ASCII(PARASTAT.EIDOS)=ASCII(LEFT(ATIM,1))),CONVERT(CHAR(8),HME,3) as [Ημερ/νία]," _
''      & " LEFT(PEL.EPO,25) as [Συναλασσόμενος],(CASE WHEN LEFT(ATIM,1) IN (" + polepis + "," + AGOEPIS + ") THEN -AJI ELSE AJI END) as [Συν.Αξία]," _
''      & " (CASE WHEN LEFT(ATIM,1) IN (" + polepis + "," + AGOEPIS + ") THEN -AJ1 ELSE AJ1 END) as [Κ.Αξ.10%],(CASE WHEN LEFT(ATIM,1) IN (" + polepis + "," + AGOEPIS + ") THEN -AJ2 ELSE AJ2 END) as  [Κ.Αξ." + Format(g_Fpa(2), "##") + "%],AJ3  AS [Κ.Αξ." + Format(g_Fpa(3), "##") + "%],AJ4 AS [κ.Αξ." + Format(g_Fpa(4), "##.#") + "%],AJ5 as [Κ.Αξ.0%],AJ6 as [Κ.Αξ.19%],SKOPOS as [Σκοπός],SXETIKO AS [ΣΧΕΤΙΚΟ],KERDOS AS [ΚΕΡΔΟΣ],(case when TRP='ΜΕ' THEN AJI ELSE 0 end ) AS [ΜΕΤΡΗΤ],(case when TRP<>'ΜΕ' THEN AJI ELSE 0 end) AS [ΠΙΣΤΩΣΗ] "    'into temp" + LTrim(Str(gXEIRISTHS))
''    d = d + " from TIM inner join PEL on TIM.EIDOS=PEL.EIDOS and TIM.KPE=PEL.KOD "
''    D0 = "select SUM(CASE WHEN LEFT(ATIM,1) IN(" + polepis + "," + AGOEPIS + ") THEN -AJI ELSE AJI END ),SUM(CASE WHEN LEFT(ATIM,1) IN(" + polepis + "," + AGOEPIS + ") THEN -AJ1 ELSE AJ1 END),SUM(CASE WHEN LEFT(ATIM,1) IN(" + polepis + "," + AGOEPIS + ") THEN -AJ2 ELSE AJ2 END),SUM(AJ3),SUM(AJ4),SUM(AJ5),sum(AJ6),sum(KERDOS) AS SK,SUM(case when TRP='ΜΕ' THEN AJI ELSE 0 END ) AS METR,SUM(case when TRP<>'ΜΕ' THEN AJI ELSE 0 END ) AS PIST FROM TIM "
''
''
''    synt = ""
''
''    If Combo2.Text = Combo2.List(0) Then
''        ' synt = "ART IS NULL AND "   '   OR LEFT(ART,1)=' '"
''        synt = " ( ART IS NULL OR ART='') AND "
''    End If
''    If Combo2.Text = Combo2.List(1) Then
''        synt = "LEFT(ART,1)='1' AND "
''    End If
''
''    If Len(Combo8.Text) > 0 Then
''        synt = synt + " LEFT(EIDPAR,1)='" + Left(Combo8, 1) + "' AND "
''    End If
''
''
''    If Len(metrhta.Text) > 0 Then
''        If metrhta.Text = metrhta.List(0) Then
''            synt = synt + " TRP='ΜΕ' " + " AND "
''        Else
''            synt = synt + " TRP<>'ΜΕ' " + " AND "
''        End If
''    End If
''
''
''
''
''
''    If APOT.ListIndex <= 0 Then
''
''    Else
''        If Len(synt) = 0 Then
''            synt = " B_N1=" + Str(APOT.ListIndex + 1) + " AND "
''        Else
''            synt = synt + "  B_N1=" + Str(APOT.ListIndex + 1) + " AND "
''        End If
''    End If
''
''
''
''    If Fsql = "" Then
''        d = d + "where " + synt + " left(ATIM,1) IN (" + mChoose + ") and HME>='" + Format(DTPicker1.value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.value + 1, "mm/dd/yyyy") + "' order by HME,ATIM;"
''        D0 = D0 + "where " + synt + " left(ATIM,1) IN (" + mChoose + ") and HME>='" + Format(DTPicker1.value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.value + 1, "mm/dd/yyyy") + "'"
''    Else
''        d = d + "where " + synt + " left(ATIM,1) IN (" + mChoose + ") and HME>='" + Format(DTPicker1.value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.value + 1, "mm/dd/yyyy") + "' and " + Fsql + " order by HME,ATIM;"
''        D0 = D0 + "where " + synt + " left(ATIM,1) IN (" + mChoose + ") and HME>='" + Format(DTPicker1.value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.value + 1, "mm/dd/yyyy") + "' and " + Fsql + " "
''
''    End If
''    'db.Execute d
''
''    On Error Resume Next
''    'Gdb.Execute "drop table TEMP" + LTrim(Str(gXEIRISTHS))
''    'Gdb.Execute d
''
''
''    On Error GoTo 0
''    data1.RecordSource = d    ' "select sum(AJI) as saji2 from TEMP" + LTrim(Str(gXEIRISTHS))
''
''  On Error Resume Next
''  data1.Refresh
''
''    'mSaji2 = data1.Recordset("saji2")
''
''    'data1.RecordSource = "select *from TEMP" + LTrim(Str(gXEIRISTHS))
''
''
''    'data1.Refresh
''    'data1.Recordset.AddNew
''    'data1.Recordset("aji") = mSaji2
''    'data1.Recordset("epo") = "Σ Υ Ν Ο Λ Α "
''    'data1.Recordset.Update
''
''    ' Dim R As New ADODB.Recordset
''    On Error Resume Next
''    R.Open D0, Gdb, adOpenForwardOnly, adLockReadOnly
''
''
''    'data1.Refresh
''    On Error Resume Next
''    DBGrid1.Columns(3).Alignment = dbgRight
''    DBGrid1.Columns(3).FooterText = Format(R(0), "######0.00")
''    DBGrid1.Columns(4).FooterText = Format(R(1), "######0.00")
''    DBGrid1.Columns(5).FooterText = Format(R(2), "######0.00")
''    DBGrid1.Columns(6).FooterText = Format(R(3), "######0.00")
''    DBGrid1.Columns(7).FooterText = Format(R(4), "######0.00")
''    DBGrid1.Columns(8).FooterText = Format(R(5), "######0.00")
''    DBGrid1.Columns(9).FooterText = Format(R(6), "######0.00")
''    DBGrid1.Columns(12).FooterText = Format(R(7), "######0.00")
''
''    DBGrid1.Columns(13).FooterText = Format(R(8), "######0.00")
''    DBGrid1.Columns(14).FooterText = Format(R(9), "######0.00")
''
''
''    DBGrid1.Columns(0).Caption = "Παραστατικό"
''    DBGrid1.Columns(1).Caption = "Ημερομηνία"
''    DBGrid1.Columns(2).Caption = "Πελάτης /Προμηθευτής"
''    DBGrid1.Columns(3).Caption = "Αξία με ΦΠΑ"
''
''    DBGrid1.Columns(0).width = 1000
''    DBGrid1.Columns(1).width = 1000
''    DBGrid1.Columns(2).width = 3800
''    DBGrid1.Columns(3).width = 800
''    DBGrid1.Columns(3).NumberFormat = "##,###,##0.00"
''    Me.MousePointer = vbNormal
''
''
''    If ekt Then
''        'typos2
''        print3_xar d, "001111111111111111", Combo1.Text, 0
''
''    End If
''
''
''
''
''
''
''
''
''
''
''
''
''
''
''
''
''
''
'
'
'End Sub

Private Sub SumHmeras_Click()

        '<EhHeader>
        On Error GoTo SumHmeras_Click_Err

        '</EhHeader>

        Dim D As String

        Dim MSYNT As String

        Dim SA As Single

        Dim SP As Single

        Dim k As Long

        Dim mPOL As String

        Dim mAGO As String

        Dim R As New ADODB.Recordset

        Dim SP2 As Single

        Dim a

        Dim AGOEPIS As String

        Dim polepis As String, pol As String, ago As String, APIS

100     pol = ""
110     a = Get_AJ(pol, polepis, ago, AGOEPIS)

120     mPOL = pol
130     mAGO = ago

        'R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic

        ' mpol = ""
        'mago = ""

        'R.MoveFirst
        'Do While Not R.EOF
        '  If Not IsNull(R("POL")) Then
        '   If R("pol") = 1 Then ' 1=POLISEIS 2=AGORES
        '     mpol = mpol + "'" + R("EIDOS") + "',"
        '  Else
        '     mago = mago + "'" + R("EIDOS") + "',"
        '  End If
        ' End If
        '  R.MoveNext
        '  Loop
        ' R.Close

140     mPOL = mPOL + ",''"
150     mAGO = mAGO + ",''"
160     DBGrid1.Visible = True
170     MSFlexGrid1.Visible = False

        On Error Resume Next

180     Me.MousePointer = vbHourglass

190     If Check2 Then

200         If apot.ListIndex = -1 Then
210             MSYNT = ""
            Else
220             MSYNT = " APOT=" + str(1 + apot.ListIndex) + " AND "
            End If

            Dim ANS As Integer

230         ANS = MsgBox("Μόνο τις λιανικές πωλήσεις;", vbYesNo)

240         If ANS = vbYes Then
250             pol = "'l','L'"
            End If

260         Data1.RecordSource = "SELECT  CONVERT(CHAR(12),HME,3) ,APOT," & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN LEFT(ATIM,1) IN (" + mAGO + ") THEN XRE*TIMM ELSE 0 END )) AS [ΑΓΟΡΕΣ]," & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN LEFT(ATIM,1) IN (" + pol + ") THEN MIK_AJIA ELSE 0 END)) AS [ΠΩΛΗΣΕΙΣ+εκπτ]," & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN LEFT(ATIM,1) IN (" + pol + ") THEN PIS*TIMM*(1-EKPT/100) ELSE 0 END)) AS [ΠΩΛΗΣΕΙΣ]," & " CONVERT(DECIMAL(10,2),SUM(CASE WHEN LEFT(ATIM,1) IN (" + polepis + ") THEN POSO*TIMM*(1-EKPT/100) ELSE 0 END )) AS [ΕΠΙΣΤΡ]" & " from EGGTIM WHERE " + MSYNT + "  HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "' " & "  GROUP BY CONVERT(CHAR(12),HME,3),APOT ORDER BY APOT,CONVERT(CHAR(12),HME,3)"

270         Data1.Refresh

280         DBGrid1.Visible = False
290         MSFlexGrid1.Visible = False
300         MSFlexGrid2.Visible = True

310         MSFlexGrid2.AddItem " " & Chr(9)
320         MSFlexGrid2.ColAlignment(3) = 6    'RIGHT
330         MSFlexGrid2.ColAlignment(2) = 6    'RIGHT
340         MSFlexGrid2.ColAlignment(4) = 6    'RIGHT
350         MSFlexGrid2.ColAlignment(5) = 6    'RIGHT
360         SA = 0
370         SP = 0
380         k = 0
390         SP2 = 0

            Dim SEP

400         SEP = 0

410         For k = 1 To MSFlexGrid2.rows - 2
420             SA = SA + Val(MSFlexGrid2.TextMatrix(k, 2))
430             SP = SP + Val(MSFlexGrid2.TextMatrix(k, 3))
440             SP2 = SP2 + Val(MSFlexGrid2.TextMatrix(k, 4))
450             SEP = SEP + Val(MSFlexGrid2.TextMatrix(k, 5))
            Next

460         MSFlexGrid2.TextMatrix(k, 2) = SA
470         MSFlexGrid2.TextMatrix(k, 3) = SP
480         MSFlexGrid2.TextMatrix(k, 4) = SP2
490         MSFlexGrid2.TextMatrix(k, 5) = SEP
500         MSFlexGrid2.TopRow = MSFlexGrid2.rows - 10

        Else
510         D = "select HME as shme,sum(AJI) as saji "
520         D = D + "from TIM  where '" + Right(Combo1.Text, 1) + "'=left(ATIM,1) "

530         D = D + "and HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "' "
540         D = D + "group by HME;"
550         Data1.RecordSource = D

560         Data1.Refresh
570         DBGrid1.columns(1).Caption = "Σύνολο με ΦΠΑ"
580         DBGrid1.columns(0).Caption = "Ημερομηνία"
590         DBGrid1.columns(1).NumberFormat = "##,###,##0.00"

        End If

600     If ekt Then
            'typos2
            ' On Error GoTo SumHmeras_Click_Err ' NA SBHSTEI STO TELOS
610         D = Data1.RecordSource
620         print3_xar D, "0111111111", Combo1.Text, 0

        End If

630     Me.MousePointer = vbNormal

        '<EhFooter>
        Exit Sub

SumHmeras_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par4.SumHmeras_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.SumHmeras_Click " & "at line " & Erl

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
         "in ADOMERCNEW.par4.mac " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par4.mac " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

'Function Get_AJANTIGRAFO(ByRef pol As String, polepis As String, ago As String, AGOEPIS As String) As Boolean
'
'    Dim db As Database
'    Dim R As New ADODB.Recordset, X As String
'
'    'If gConnect = "Access" Then
'    '   Set db = OpenDatabase(gDir, False, False)
'    'Else
'    '   Set db = OpenDatabase(gDir, False, False, gConnect)
'    'End If
'
'    R.Open UCase("select pol,eidos,ajia_apou from parastat"), Gdb, adOpenDynamic, adLockOptimistic
'
'    pol = ""
'    R.MoveFirst
'    Do While Not R.EOF
'        If Not IsNull(R("eidos") + R("pol") + R("ajia_apou")) Then
'            If R("pol") = "1" And R("ajia_apou") = "3" Then
'                pol = pol + "'" + R("eidos") + "',"
'            End If
'
'            If R("pol") = "1" And R("ajia_apou") = "4" Then
'                polepis = polepis + "'" + R("eidos") + "',"
'            End If
'
'            If R("pol") = "2" And R("ajia_apou") = "1" Then
'                ago = ago + "'" + R("eidos") + "',"
'            End If
'            If R("pol") = "2" And R("ajia_apou") = "2" Then
'                AGOEPIS = AGOEPIS + "'" + R("eidos") + "',"
'            End If
'
'
'
'
'        End If
'        R.MoveNext
'    Loop
'    pol = Left(pol, Len(pol) - 1)
'    If Len(polepis) > 0 Then
'        polepis = Left(polepis, Len(polepis) - 1)
'    Else
'        polepis = ""
'    End If
'
'    AGOEPIS = Left(AGOEPIS, Len(AGOEPIS) - 1)
'    ago = Left(ago, Len(ago) - 1)
'    Get_AJANTIGRAFO = True
'
'
'
'End Function


 

Private Sub cmdClearFilter_Click()

    'Clears filter from grid

    For Each Col In DBGrid1.columns

        Col.FILTERTEXT = ""

    Next Col

    Data1.Recordset.Filter = adFilterNone

End Sub

 

Private Sub DBGrid1_FilterChange()

'Gets called when an action is performed on the filter bar

On Error GoTo errHandler

 'Dim cols As TrueOleDBGrid80.columns

Set ColS = DBGrid1.columns

Dim c As Integer

c = DBGrid1.Col

DBGrid1.HoldFields

Data1.Recordset.Filter = getFilter()

DBGrid1.Col = c

DBGrid1.EditActive = True

Exit Sub

 

errHandler:

    MsgBox Err.Source & ":" & vbCrLf & Err.Description

    Call cmdClearFilter_Click

End Sub

 

Private Function getFilter() As String

    'Creates the SQL statement in adodc1.recordset.filter

    'and only filters text currently. It must be modified to
    'filter other data types.

   ' Dim cols As TrueOleDBGrid80.columns
'Dim col As TrueOleDBGrid80.Column
    Dim tmp As String

    Dim N As Integer

    For Each Col In ColS

        If Trim(Col.FILTERTEXT) <> "" Then

            N = N + 1

            If N > 1 Then

                tmp = tmp & " AND "

            End If

            tmp = tmp & Col.DataField & " LIKE '" & Col.FILTERTEXT & "*'"

        End If

    Next Col

                

    getFilter = tmp

End Function


