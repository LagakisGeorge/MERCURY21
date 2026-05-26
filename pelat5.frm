VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form pelat5 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   11085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18615
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11085
   ScaleWidth      =   18615
   WindowState     =   2  'Maximized
   Begin VB.ComboBox FILTRO 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   1200
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   10680
      Top             =   6840
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
   Begin VB.TextBox c3 
      Height          =   285
      Left            =   7605
      TabIndex        =   18
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox c2 
      Height          =   285
      Left            =   7605
      TabIndex        =   17
      ToolTipText     =   "@c2 ð.÷. CH2 LIKE    '@c2%'  Þ  Á=@C2"
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox c1 
      Height          =   285
      Left            =   7605
      TabIndex        =   16
      ToolTipText     =   "@c1 ð.÷. CH2 LIKE    '@c1%'  Þ  Á=@C1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000010&
      Caption         =   "Åîïäïò"
      Height          =   615
      Left            =   8445
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7920
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5280
      Top             =   8760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command11 
      Caption         =   "ÁðïèÞêåõóç Report"
      Height          =   435
      Left            =   2820
      TabIndex        =   14
      Top             =   8100
      Width           =   2310
   End
   Begin VB.CheckBox rCheck 
      Caption         =   "×ñçóéìïðïéþ ôï åðéëåãìåíï report"
      Height          =   435
      Left            =   2820
      TabIndex        =   13
      Top             =   7515
      Width           =   2310
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   45
      TabIndex        =   12
      Top             =   7470
      Width           =   2715
   End
   Begin VB.TextBox Text2 
      Height          =   705
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   6765
      Width           =   5085
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Top             =   6840
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "åêôýðùóç"
      Height          =   285
      Left            =   3435
      TabIndex        =   8
      Top             =   810
      Width           =   2235
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
      Bindings        =   "pelat5.frx":0000
      Height          =   375
      Left            =   17160
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      FillStyle       =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin MSAdodcLib.Adodc EGG 
      Height          =   375
      Left            =   -45
      Top             =   10740
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   375
      Left            =   120
      Top             =   7680
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSComCtl2.DTPicker apo 
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "@X1  ð.÷.  HME>@X1"
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   166068225
      CurrentDate     =   38358
   End
   Begin VB.PictureBox CrystalReport1 
      Height          =   480
      Left            =   3360
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   8760
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Õðïëïãéóìüò"
      Height          =   615
      Left            =   6405
      TabIndex        =   1
      Top             =   7920
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "pelat5.frx":0014
      Left            =   240
      List            =   "pelat5.frx":001E
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker eos 
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      ToolTipText     =   "@X2  ð.÷.  HME>@X2"
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   166068225
      CurrentDate     =   38358
   End
   Begin TrueOleDBGrid80.TDBGrid GR2 
      Bindings        =   "pelat5.frx":0038
      Height          =   4920
      Left            =   120
      TabIndex        =   23
      Top             =   1680
      Width           =   16395
      _ExtentX        =   28919
      _ExtentY        =   8678
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ößëôñï óôçí áéôéïëïãßá"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5400
      TabIndex        =   22
      Top             =   6960
      Width           =   1785
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÐáñÜìåôñïò c3"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   6000
      TabIndex        =   21
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÐáñÜìåôñïò c2"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   6000
      TabIndex        =   20
      Top             =   480
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÐáñÜìåôñïò c1"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   6000
      TabIndex        =   19
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Aéôéïëïãßá"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   6045
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Åùò"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Áðü"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "pelat5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim F_RUN As Integer


Dim F_COL  As TrueOleDBGrid80.Column



Dim F_COLS          As TrueOleDBGrid80.columns







'
Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Private Sub Command1_Click()

        '
        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>
        Dim x      As String, order As String

        Dim k      As Integer, R As New ADODB.Recordset, Sum As Single

        'ÂÑÉÓÊÙ ÔÏÕÓ ËÏÃÁÑÉÁÓÌÏÕÓ ÐÏÕ ÁÍÁÖÅÑÏÍÔÁÉ ÓÔÏ ÔÁÌÅÉÏ,ÐÉÓÔÙÔÉÊÅÓ ÃÉÁ ÍÁ ÔÏÕÓ ÅÎÁÉÑÅÓÙ ÁÐÏ ÔÉÓ ÊÉÍÇÓÅÉÓ
        Dim TAMEIA As String

100     TAMEIA = ""
110     R.Open "SELECT *FROM PINAKES WHERE TYPOS=12 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

120     Do While Not R.EOF

130         If R("typos") = 12 Then
140             TAMEIA = TAMEIA + "'" + Format(R("TIMH"), "0000") + "',"
            End If

150         R.MoveNext
        Loop

160     TAMEIA = Left(TAMEIA, Len(TAMEIA) - 1)
170     R.Close

180     F_RUN = 1

        Dim TT

190     TT = GetCurrentTime()

        Dim FILTERTEXT As String

200     If Len(Trim(Text1.Text)) > 0 Then
210         FILTERTEXT = " and EGG.AIT LIKE '" + Text1.Text + "%' "
        Else
220         FILTERTEXT = ""
        End If
        
        If Left$(FILTRO.Text, 2) = "**" Then
        
        Else
        
           FILTERTEXT = FILTERTEXT + " AND EIDXPI=" + mID$(FILTRO.Text, 51, 5)
        End If
        
        
        
        

        'If gXEIRISTHS = 2 Then
        '
        '
        '  FILTERTEXT = " and  KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%' " + FILTERTEXT
        '
        '
        '
        'End If

        'CONVERT(CHAR(8),HME,3)
230     x = "select CONVERT(CHAR(8),HME,3) as [Hìåñ/íßá],EGG.ATIM ,LEFT(EGG.AIT,25) AS [Áéôéïëïãßá],ROUND(EGG.XREOSI,2) as [×ñÝùóç],ROUND(EGG.PISTOSI,2) as [Ðßóôùóç],PEL.EPO AS [Åðùíõìßá]  from EGG "
        'x = "select egg.hme,egg.atim,egg.ait,iif(egg.xre>0,format(egg.XRE,'###,###.00'),'') as [×ñÝùóç],iif(egg.xre<0,format(-egg.XRE,'###,###.00'),'') as [Ðßóôùóç],kod,eidos  into dok5  in 'c:\mercvb\reports\reports.mdb' from egg "
240     x = x + " inner join PEL on EGG.EIDOS=PEL.EIDOS and EGG.KOD=PEL.KOD "

        'x = x + " inner join pel on egg.eidos+egg.kod=pel.eidos+pel.kod  "
250  '   X = X + " where EGG.KOD NOT IN (" + TAMEIA + ") AND EGG.HME>='" + Format(apo.Value, "mm/dd/yyyy") + "'  and EGG.HME<'" + Format(DateAdd("d", 1, eos.Value), "mm/dd/yyyy") + "' and EGG.EIDOS='" + IIf(Combo1.Text = Combo1.List(0), "e", "r") + "' " + FILTERTEXT + " order by EGG.ID"
        x = x + " where  EGG.HME>='" + Format(apo.Value, "mm/dd/yyyy") + "'  and EGG.HME<'" + Format(DateAdd("d", 1, eos.Value), "mm/dd/yyyy") + "' and EGG.EIDOS='" + IIf(Combo1.Text = Combo1.List(0), "e", "r") + "' " + FILTERTEXT + " order by EGG.ID"
        Dim DUM, DB As Database, dbacc As Database

        Dim NX, NP
    
260     NX = 3: NP = 4

270     If rCheck.Value = vbChecked Then
280         x = PutDates(Text2.Text)
        End If

        'dum = OpenD(db)

290     Data1.ConnectionString = gConnect
        'Data1.DatabaseName = gDir
300     Data1.RecordSource = x ' "select * FROM EGG"   'X    ' "egg"

310     Data1.Refresh
320     Text2.Text = x





    ADODC2.ConnectionString = gConnect
         ADODC2.RecordSource = x
         ADODC2.Refresh
Dim SUMES2(0 To 50)

'If synola.Value = vbChecked Then
         Do While Not ADODC2.Recordset.EOF

             For k = 0 To ADODC2.Recordset.FIELDS.Count - 1

                 If IsNumeric(ADODC2.Recordset.FIELDS(k).Value) And ADODC2.Recordset.FIELDS(k).Type <> 202 Then
                     SUMES2(k) = SUMES2(k) + ADODC2.Recordset.FIELDS(k).Value
                    End If

                Next
               ' DoEvents

                ADODC2.Recordset.MoveNext

                '  Exit Do
            Loop

         For k = 0 To ADODC2.Recordset.FIELDS.Count - 1

             If SUMES2(k) > 0 Then
                 GR2.Splits(0).columns(k).FooterText = Format(SUMES2(k), "######0.00")
             End If

         Next
'End If



        GR2.Refresh
        
        
        
        
        
        
        
        

'
'        'On Error Resume Next
'        '
'        'Set dbacc = OpenDatabase("c:\mercvb\reports\reports.mdb")
'        'dbacc.Execute "drop table dok5"
'        'On Error GoTo Command1_Click_Err
'        '
'        ''Kill gDir + "\temp.dbf"
'        '
'        'db.Execute x
'        '
'        'dbacc.Execute "drop table pel5"
'        'db.Execute "select eidos,kod,epo into pel5 in 'c:\mercvb\reports\reports.mdb' from pel where eidos='" + IIf(Combo1.Text = Combo1.List(0), "e", "r") + "'"
'        '
'        'Data1.DatabaseName = "c:\mercvb\reports\reports.mdb"
'        'Data1.Connect = "Access"
'        'Data1.RecordSource = "select dok5.*,pel5.epo from dok5 inner join pel5  on dok5.eidos+dok5.kod=pel5.eidos+pel5.kod"
'        'Data1.Refresh
'        '
'330     MDIForm1.Caption = GetCurrentTime() - TT
'
'        'Me.Caption = "--"
'        'Set R = Data1.Recordset
'       ' On Error Resume Next
'
'        'R.MoveFirst
'
'        Dim mc, sum2
'
'340     mc = 0
'
'350     Sum = 0
'360     sum2 = 0
'
'370     For mc = 1 To MSFlexGrid1.ROWS - 1
'380         Sum = Sum + mval(MSFlexGrid1.TextMatrix(mc, 3))
'390         sum2 = sum2 + mval(MSFlexGrid1.TextMatrix(mc, 4))
'
'400         If mc Mod 2 = 0 Then
'410             MSFlexGrid1.row = mc
'
'420             For k = 0 To MSFlexGrid1.ColS - 1: MSFlexGrid1.Col = k: MSFlexGrid1.CellBackColor = &HC0E0FF: Next
'            End If
'
'430         MSFlexGrid1.TextMatrix(mc, 0) = Format(MSFlexGrid1.TextMatrix(mc, 0), "dd/mm/yyyy")
'
'440         MSFlexGrid1.TextMatrix(mc, NX) = Format(Val(MSFlexGrid1.TextMatrix(mc, NX)), "######.##")
'450         MSFlexGrid1.TextMatrix(mc, NP) = Format(Val(MSFlexGrid1.TextMatrix(mc, NP)), "######.##")
'
'            ' If Val(MSFlexGrid1.TextMatrix(mc, NX)) < 0 Then
'            '      MSFlexGrid1.TextMatrix(mc, NX) = ""
'            '      MSFlexGrid1.TextMatrix(mc, NP) = Format(-Val(MSFlexGrid1.TextMatrix(mc, NP)), "######.00")
'            ' Else
'            '      MSFlexGrid1.TextMatrix(mc, NP) = ""
'            '      MSFlexGrid1.TextMatrix(mc, NX) = Format(Val(MSFlexGrid1.TextMatrix(mc, NX)), "######.00")
'            ' End If
'
'460         MSFlexGrid1.TextMatrix(mc, 0) = Format(MSFlexGrid1.TextMatrix(mc, 0), "dd/mm/yyyy")
'
'        Next
'
'470     MSFlexGrid1.ColWidth(5) = 2500
'480     MSFlexGrid1.ColWidth(4) = 1200    'ÁÐÁ
'490     MSFlexGrid1.ColWidth(3) = 1200    'ÁÐÁ
'500     MSFlexGrid1.ColWidth(2) = 2500    'ÅÉÄ
'
'        MSFlexGrid1.AllowUserResizing = flexResizeColumns
'
'510     MSFlexGrid1.AddItem ("Óýíïëá" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(Sum, "###,###.00")) & Chr(9) & Format(sum2, "###,###.00")
'
'520     MSFlexGrid1.TopRow = MSFlexGrid1.ROWS - 10
'530     MSFlexGrid1.ColAlignment(4) = 6    ' RIGHT
'
'540     MSFlexGrid1.ColAlignment(3) = 6    ' RIGHT
'        'MSFlexGrid1.ColAlignment(7) = 6 ' RIGHT
'
'        'DBGrid3.Columns(2).Caption = Format(sum, "###,###.00")
'        'DBGrid3.ReBind
'        'DBGrid3.Columns(0).Caption = Format(mc, "###,###")

        'If Check1 Then typos2

550     If Check1 Then

            Dim PELA

560         PELA = IIf(Combo1.Text = Combo1.List(0), "ÐÅËÁÔÅÓ", "ÐÑÏÌÇÈÅÕÔÅÓ")

            On Error GoTo Command1_Click_Err    ' Resume Next

570         print3_xar x, "000110", PELA, 0    'typos2
        End If

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat5.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat5.Command1_Click " & "at line " & Erl

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

130     ADODC2.Recordset.Filter = getFilter()

140     GR2.Col = c

150     GR2.EditActive = True

        Exit Sub

errHandler:

160    ' MsgBox Err.Source & ":" & vbCrLf & Err.Description

        '    Call cmdClearFilter_Click

170     Resume Next

        '<EhFooter>
        Exit Sub

GR2_FilterChange_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot3.GR2_FilterChange " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PELAT5.GR2_FilterChange " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub





Function PutDates2(ByVal x As String)

        '<EhHeader>
        On Error GoTo PutDates2_Err

        '</EhHeader>

        Dim X1 As String, X2 As String

100     X1 = "'" + Format(apo.Value, "MM/DD/YYYY") + "'"
110     X2 = "'" + Format(DateAdd("H", 23.99, eos.Value), "MM/DD/YYYY") + "'"

        Dim SQL2 As String

120     SQL2 = x    ' TEXT2.Text

130     SQL2 = Replace(SQL2, "@X1", X1)
140     SQL2 = Replace(SQL2, "@X2", X2)

150     PutDates2 = SQL2

        '<EhFooter>
        Exit Function

PutDates2_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat5.PutDates2 " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat5.PutDates2 " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

'Private Sub Typose()
'End Sub

Private Sub Command11_Click()

        '<EhHeader>
        On Error GoTo Command11_Click_Err

        '</EhHeader>

100     If Len(Dir("c:\mercvb\queries\pelKIN", vbDirectory)) = 0 Then
110         MkDir "c:\mercvb\queries\pelKIN"
        End If

120     cd1.InitDir = "c:\mercvb\queries\pel"
130     cd1.ShowSave

        Dim F

140     F = cd1.FILENAME

150     Open F For Output As #5
160     Print #5, Text2.Text
170     Close #5

        '<EhFooter>
        Exit Sub

Command11_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat5.Command11_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat5.Command11_Click " & "at line " & Erl

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
130         SendKeys "{ENTER}+{DOWN}+{DOWN}+{DOWN}+{DOWN} "
        End If

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat5.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat5.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'
'End Sub

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

180     Text2.Text = SS

        '<EhFooter>
        Exit Sub

File1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat5.File1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat5.File1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then Command2_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat5.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat5.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function PutDates(ByVal x As String)

        '<EhHeader>
        On Error GoTo PutDates_Err

        '</EhHeader>

        Dim X1 As String, X2 As String

100     X1 = "'" + Format(apo, "MM/DD/YYYY") + "'"
110     X2 = "'" + Format(DateAdd("H", 23.99, eos), "MM/DD/YYYY") + "'"

        Dim SQL2 As String

120     SQL2 = x    ' TEXT2.Text

130     SQL2 = Replace(SQL2, "@X1", X1)
140     SQL2 = Replace(SQL2, "@X2", X2)
150     SQL2 = Replace(SQL2, "@x1", X1)
160     SQL2 = Replace(SQL2, "@x2", X2)

170     SQL2 = Replace(SQL2, "@c1", C1.Text)
180     SQL2 = Replace(SQL2, "@c2", c2.Text)
190     SQL2 = Replace(SQL2, "@c3", c3.Text)
200     SQL2 = Replace(SQL2, "@C1", C1.Text)
210     SQL2 = Replace(SQL2, "@C2", c2.Text)
220     SQL2 = Replace(SQL2, "@C3", c3.Text)

230     PutDates = SQL2

        '<EhFooter>
        Exit Function

PutDates_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.PutDates " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.PutDates " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function




Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     mForm_Load Me, fh, fw, ft, fl

110     Me.KeyPreview = True

        'If gUserId = 2 Then
        ' Command1.Enabled = False
        '
        'End If

        MSFlexGrid1.AllowUserResizing = flexResizeColumns

        'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

120     Combo1.Text = Combo1.List(0)
        'Combo2.Text = Combo2.List(1)
130     F_RUN = 0

140     Me.Picture = LoadPicture(gPicture)

150     apo.Value = Now
160     eos.Value = Now

170     If Len(Dir("c:\mercvb\queries\pel", vbDirectory)) = 0 Then
180         MkDir "c:\mercvb\queries\pel"
        End If

190     File1.Path = "c:\mercvb\queries\pel"

200     cd1.InitDir = "c:\mercvb\queries\pel"










Dim r2 As New ADODB.Recordset


  r2.Open "select * from XREOPIS where PEL='" + IIf(Combo1.Text = Combo1.List(0), "e", "r") + "'", Gdb, adOpenDynamic, adLockOptimistic
  
  r2.MoveFirst
        '  Els
        '     Adodc1.RecordSource = "select * from XREOPIS where RIGHT(TITLOS,1)<>'.' AND PEL ='" + gBuff + "'"
        ' End If

FILTRO.AddItem "*** OLA ***"

340     Do While Not r2.EOF
350         FILTRO.AddItem Left(r2("titlos") + Space$(50), 50) + Format(r2!EIDOS, "000")
360         r2.MoveNext
        Loop




FILTRO.Text = FILTRO.List(0)



















        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat5.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat5.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function mval(s As String)

        '<EhHeader>
        On Error GoTo mval_Err

        '</EhHeader>
        Dim k, L

100     For k = 1 To Len(s)
110         L = InStr(s, ",")

120         If L > 0 Then
130             Mid$(s, L, 1) = " "
            End If

        Next

140     mval = Val(s)

        '<EhFooter>
        Exit Function

mval_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat5.mval " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat5.mval " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 6.3, 10545, 9585, fh, fw, ft, fl
End Sub

Private Sub GR2_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
        Dim sumes(100) As String
Dim k As Integer

100     For k = 0 To ADODC2.Recordset.FIELDS.Count - 1
110         sumes(k) = GR2.Splits(0).columns(k).FooterText    '  = Format(SUMES(k), "######0.00")
        Next

120     If ADODC2.Recordset.sort = "[" & GR2.columns(ColIndex).DataField & "] asc" Then   ' strSort
130         ADODC2.Recordset.sort = "[" & GR2.columns(ColIndex).DataField & "] desc"    ' strSort
        Else
140         ADODC2.Recordset.sort = "[" & GR2.columns(ColIndex).DataField & "] asc"    ' strSort
        End If

150     For k = 0 To ADODC2.Recordset.FIELDS.Count - 1
160         GR2.Splits(0).columns(k).FooterText = sumes(k)   '  = Format(SUMES(k), "######0.00")
        Next

        '    Adodc2.Recordset.Sort = "[" & TDBGrid.Columns(ColIndex).DataField & "] desc" ' strSort

        '<EhFooter>
        Exit Sub


End Sub

