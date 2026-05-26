VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form epit3 
   BackColor       =   &H00FF0000&
   Caption         =   "–·ÒÔıÛﬂ·ÛÁ ≈ÈÙ·„˛Ì"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   15150
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "epit3.frx":0000
      Left            =   7980
      List            =   "epit3.frx":0002
      TabIndex        =   23
      Top             =   5985
      Width           =   3525
   End
   Begin VB.ComboBox PLAFON 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "epit3.frx":0004
      Left            =   1590
      List            =   "epit3.frx":0006
      TabIndex        =   21
      Top             =   7080
      Width           =   1635
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "epit3.frx":0008
      Left            =   5760
      List            =   "epit3.frx":0012
      TabIndex        =   20
      Text            =   "1.«Ã.À«Œ≈Ÿ”"
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1590
      TabIndex        =   15
      Top             =   6540
      Width           =   1635
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "epit3.frx":0033
      Left            =   1590
      List            =   "epit3.frx":0035
      TabIndex        =   10
      Top             =   6015
      Width           =   1635
   End
   Begin VB.CheckBox katamhna 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8055
      TabIndex        =   9
      Top             =   5220
      Width           =   180
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1590
      TabIndex        =   7
      Top             =   5610
      Width           =   1635
   End
   Begin VB.CheckBox ekt 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   8040
      TabIndex        =   4
      Top             =   5520
      Width           =   195
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   1080
      Top             =   8160
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      BackColor       =   &H00FF0000&
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   8055
      TabIndex        =   3
      Top             =   4935
      Width           =   195
   End
   Begin VB.CommandButton Command2 
      Caption         =   "≈ÓÔ‰ÔÚ"
      Height          =   495
      Left            =   9840
      TabIndex        =   2
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "–ÒÔ‚ÔÎﬁ"
      Height          =   495
      Left            =   7980
      TabIndex        =   1
      Top             =   6480
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "epit3.frx":0037
      Left            =   4485
      List            =   "epit3.frx":0059
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   5505
      Width           =   3180
   End
   Begin MSComCtl2.DTPicker d1 
      Height          =   345
      Left            =   1185
      TabIndex        =   5
      Top             =   4950
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   582
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   176160769
      CurrentDate     =   38814
   End
   Begin MSComCtl2.DTPicker d2 
      Height          =   345
      Left            =   3435
      TabIndex        =   6
      Top             =   4905
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   120913921
      CurrentDate     =   38814
   End
   Begin TrueOleDBGrid80.TDBGrid MSFlexGrid1 
      Bindings        =   "epit3.frx":0151
      Height          =   4545
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   8017
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
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2778"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2778"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2778"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2778"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2778"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=2778"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(5).Button=1"
      Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(26)=   "Column(6).Width=2778"
      Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(30)=   "Column(7).Width=2778"
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
      AllowUpdate     =   0   'False
      ColumnFooters   =   -1  'True
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      GroupByCaption  =   "”˝ÒÂÙÂ ÙÔÌ ÙﬂÙÎÔ ÙÁÚ ÛÙﬁÎÁÚ „È· ÔÏ·‰ÔÔﬂÁÛÁ"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=1200,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=1200,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=1200,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38,.ellipsis=-1"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42,.locked=0"
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
      _StyleDefs(88)  =   ":id=42,.parent=33,.locked=-1"
   End
   Begin VB.Label HME_METAB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4680
      TabIndex        =   25
      Top             =   6360
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label PlafonText 
      BackStyle       =   0  'Transparent
      Caption         =   "ƒ·ÌÂÈ·Í.ÀÔ„"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   7110
      Width           =   1050
   End
   Begin VB.Label Label9 
      Caption         =   "≈–ŸÕ.–—œÃ«»≈’‘«"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   7800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000004&
      Height          =   1080
      Left            =   7980
      Top             =   4860
      Width           =   3555
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000004&
      Height          =   2115
      Left            =   225
      Top             =   5490
      Width           =   4170
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      Height          =   585
      Left            =   225
      Top             =   4875
      Width           =   5520
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "≈˘Ú"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2955
      TabIndex        =   18
      Top             =   4935
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "¡¸"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   630
      TabIndex        =   17
      Top             =   4965
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "¡ÒÈË.≈ÈÙ·„ﬁÚ"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   435
      TabIndex        =   16
      Top             =   6555
      Width           =   1170
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "≈ÍÙ˝˘ÛÁ"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   8280
      TabIndex        =   14
      Top             =   5490
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "”ı„Í.Í·Ù· ÏﬁÌ·"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   8310
      TabIndex        =   13
      Top             =   5205
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "”ı„ÍÂÌÙÒ˘ÙÈÍ‹"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   8295
      TabIndex        =   12
      Top             =   4905
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "‘Ò‹ÂÊ·"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   6045
      Width           =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "≈˘ÌıÏﬂ·"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   495
      TabIndex        =   8
      Top             =   5625
      Width           =   930
   End
End
Attribute VB_Name = "epit3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim F_PINAKIO As Integer

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

'Dim FN_EIS '= "≈…”–—¡Œ«"
'Dim FN_PLI ' = "–À«—ŸÃ«"
'Dim FN_PLA ' = "”≈ –À¡÷œÕ"
'Dim FN_DIA ' = "ƒ…¡√—¡÷«"
'Dim FN_APO  '= "¡–œ”‘œÀ« ”≈ –—œÃ«»≈’‘«"
'Dim FN_EPI '  = "≈–…”‘—œ÷« ≈–…‘¡√«”"
'Dim FN_EPA  ' = "≈–¡Õ¡÷œ—¡ ”‘¡ ◊≈—…¡ Ã¡”"
'Dim FN_DIO ' "ƒ…œ—»Ÿ”«"
'Dim FN_EXO ' = "≈Œœƒœ”"

Const FN_EIS = "≈…”–—¡Œ«"

Const FN_PLI = "–À«—ŸÃ«"

Const FN_PLA = "”≈ –À¡÷œÕ"

Const FN_DIA = "ƒ…¡√—¡÷«"

Const FN_EPIS = "≈–…”‘—.≈–…‘.–≈À¡‘œ’ ¡–œ –—œÃ«»≈’‘«"

Const FN_APO = "¡–œ”‘œÀ« ”≈ –—œÃ«»≈’‘«"

Const FN_EPI = "≈–…”‘—œ÷« ≈–…‘¡√«”"

Const FN_EPA = "≈–¡Õ¡÷œ—¡ ”‘¡ ◊≈—…¡ Ã¡”"

Const FN_DIO = "ƒ…œ—»Ÿ”«"

Const FN_EXO = "≈Œœƒœ”"

Private Sub Combo1_Click()

        '<EhHeader>
        On Error GoTo Combo1_Click_Err

        '</EhHeader>

100     Command1_Click

        '<EhFooter>
        Exit Sub

Combo1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.epit3.Combo1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.epit3.Combo1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

        Dim synt  As String

        Dim eis   As String

        Dim pl    As String

        Dim Okeis As String

        Dim Okpl  As String

        Dim platos, k

        Dim NEIS, NPL, nOKeIS, NOKPL, NARX

        Dim L As Integer

        Dim sql

        On Error Resume Next

100     MSFlexGrid1.RowHeight = 450

        'For k = 1 To MSFlexGrid1.RowHeight - 1
        '   MSFlexGrid1.RowHeight(k) = 450   ' 900
        'Next
 
        'MSFlexGrid1.RowHeightMin = 200 ' 400

110     MSFlexGrid1.columns(0).Width = 2500   'hme
120     MSFlexGrid1.columns(1).Width = 3800   'hme
    
130     MSFlexGrid1.columns(1).Width = 3800   '
        '    MSFlexGrid1.Columns(2).width = 1200   '
        '   MSFlexGrid1.Columns(3).width = 1200   '
    
        ' MSFlexGrid1.ColWidth(1) = 3800  'epo

        ' MSFlexGrid1.ColWidth(2) = 1200  'epo
        ' MSFlexGrid1.ColWidth(3) = 1200  'epo

140     If Check1 Or katamhna Then
150         sygk

            Exit Sub

        End If

        ' Data1.DatabaseName = gDir
160     Data1.ConnectionString = gConnect

170     MSFlexGrid1.Visible = False

        Dim SUgm_str

180     If Combo1.ListIndex = 2 Then    '-----------  OLES MAZI
190         MSFlexGrid1.columns(0).Width = 2100   '
200         MSFlexGrid1.columns(3).Width = 1500   '
210         MSFlexGrid1.columns(4).Width = 1500   '
         
            ' MSFlexGrid1.ColWidth(0) = 2100  'hme
           
            '        MSFlexGrid1.ColWidth(3) = 1500  'hme
            '       MSFlexGrid1.ColWidth(4) = 1500  'hme

            ' MSFlexGrid1.Font.Name = "Arial"

220         If Left(Combo3.Text, 1) = "1" Then    '«Ã≈—.À«Œ≈Ÿ”
230             synt = "HMEL>='" + Format(d1, "mm/dd/yy") + "' and HMEL<='" + Format(d2, "mm/dd/yy") + "' AND ARX<3"
            Else
240             synt = "HMEE>='" + Format(d1, "mm/dd/yy") + "' and HMEE<='" + Format(d2, "mm/dd/yy") + "' AND ARX<3"
            End If

250         NEIS = 2: eis = "STR((CASE ARX WHEN 1 THEN POS ELSE 0 END),10,2) as [≈ﬂÛÒ·ÓÁ],"

260         NPL = 3: pl = "STR((CASE ARX WHEN 2 THEN POS ELSE 0 END),10,2)  as [–ÎÁÒ˘Ïﬁ],"
270         nOKeIS = 4: Okeis = "STR((CASE ARX WHEN 5 THEN POS ELSE 0 END),10,2) as [≈ÈÛÒ·˜ËÂﬂÛÂÚ],"
280         NOKPL = 5: Okpl = "STR((CASE ARX WHEN 6 THEN POS ELSE 0 END),10,2)  as [–ÎÁÒ˘Ï›ÌÂÚ]"
290         NARX = 6

            On Error GoTo Command1_Click_Err

300         sql = "select HMEL AS [«Ï_ÀﬁÓÂ˘Ú],EPO as [≈˘ÌıÏﬂ·]," + eis + pl + Okeis + Okpl
310         sql = sql + ", CONVERT(CHAR(12),HMEE,3) as [HÏ.≈ÈÛ·„],ARX,ONO AS [Ã≈‘¡¬…¬],ID from GRA where " + synt + " order by HMEL;"
320         Data1.RecordSource = sql
330         Data1.Refresh
340         SUgm_str = "001111"

350         If MSFlexGrid1.columns.Count < 8 Then MSFlexGrid1.columns.ADD (0)
360         If MSFlexGrid1.columns.Count < 8 Then MSFlexGrid1.columns.ADD (0)
370         If MSFlexGrid1.columns.Count < 8 Then MSFlexGrid1.columns.ADD (0)
380         If MSFlexGrid1.columns.Count < 8 Then MSFlexGrid1.columns.ADD (0)
        
            '        For k = 1 To MSFlexGrid1.ROWS - 1
            '           MSFlexGrid1.TextMatrix(k, 0) = Format(k, "00") + "." + Format(MSFlexGrid1.TextMatrix(k, 0), "DDD DD/MM/YYYY")
            '          If k Mod 2 = 0 Then
            '             For L = 0 To MSFlexGrid1.ColS - 1
            '                MSFlexGrid1.row = k: MSFlexGrid1.Col = L: MSFlexGrid1.CellBackColor = vbYellow    '&HC0E0FF
            '           Next
            '      End If
            ' Next

390         MSFlexGrid1.columns(1).Width = 2600

400         MSFlexGrid1.columns(4).alignment = dbgRight ' 7     'right
410         MSFlexGrid1.columns(2).alignment = 1     'right
420         MSFlexGrid1.columns(3).alignment = 1     'right
430         MSFlexGrid1.columns(5).alignment = 1     'right
        
            'MSFlexGrid1.ColAlignment(2) = 7    'right
            'MSFlexGrid1.ColAlignment(3) = 7    'right
            'MSFlexGrid1.ColAlignment(5) = 7    'right

440         eis = Format(Calc_Suma(NEIS), "###,###.00")
450         pl = Format(Calc_Suma(NPL), "###,###.00")

460         Okeis = Format(Calc_Suma(nOKeIS), "###,###.00")
470         Okpl = Format(Calc_Suma(NOKPL), "###,###.00")

            Dim s1 As Single, S2 As Single

480         s1 = 0
490         S2 = 0

            On Error Resume Next

500         Data1.Recordset.MoveFirst

510         Do While Not Data1.Recordset.EOF
520             s1 = s1 + Val(Data1.Recordset(3))
530             S2 = S2 + Val(Data1.Recordset(4))
540             Data1.Recordset.MoveNext
            Loop

550         MSFlexGrid1.Splits(0).columns(3).FooterText = Format(s1, "######0.00")
560         MSFlexGrid1.Splits(0).columns(4).FooterText = Format(S2, "######0.00")

            ' ÒÔÛË›Ù˘ ÛÂÈÒ‹ „È· Ù· Û˝ÌÔÎ·

            ' MSFlexGrid1.AddItem ("" & Chr(9) & "”˝ÌÔÎÔ" & Chr(9) & eis & Chr(9) & pl & Chr(9) & Okeis & Chr(9) + Okpl)

            'MSFlexGrid1.AddItem ("" & Chr(9) & "ƒÈ·ˆÔÒ‹" & Chr(9) & Format(Calc_Suma(NEIS) - Calc_Suma(NPL), "###,###.00"))

            ' ÍÔÍÍÈÌÈÊ˘ Ù· Û˝ÌÔÎ·

            '        MSFlexGrid1.row = MSFlexGrid1.ROWS - 2
            '       MSFlexGrid1.Col = 0
            '      MSFlexGrid1.RowSel = MSFlexGrid1.ROWS - 1
            '     MSFlexGrid1.Col = 2
            '    MSFlexGrid1.CellBackColor = vbRed
            '   MSFlexGrid1.Col = 3
            '  MSFlexGrid1.CellBackColor = vbRed
            ' MSFlexGrid1.ColAlignment(3) = 7

570     ElseIf Combo1.ListIndex = 8 Then    '-----------  ÂÈÛÙÒ.ÂÎ·Ù
    
580         If Left(Combo3.Text, 1) = "2" Then    ' .«Ã.≈…”¡√Ÿ√«”
590             sql = "select HMEL as [«Ï_ÀﬁÓÁÚ],EPO as [E˘ÌıÏﬂ·]," & "str(POS,10,2) as [¡Óﬂ·],PAR AS [A— ≈–…‘],HMEE AS [«Ã.≈…”],(SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=6 AND AYJON=KOD_BANK) AS [‘—¡–≈∆¡] ,PINAKIO AS [–…Õ¡ …œ],ONO+'  ' AS [Ã≈‘¡¬…¬],ID from GRA where HMEE>='" + Format(d1, "Mm/dd/yy") + "' and HMEE<'" + Format(DateAdd("D", 1, d2), "mm/dd/yy") + "' and ARX=11"

            Else

600             sql = "select HMEL as [«Ï_ÀﬁÓÁÚ],EPO as [E˘ÌıÏﬂ·]," & "str(POS,10,2) as [¡Óﬂ·],PAR AS [A— ≈–…‘],HMEE AS [«Ã.≈…”],(SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=6 AND AYJON=KOD_BANK) AS [‘—¡–≈∆¡] ,PINAKIO AS [–…Õ¡ …œ],ONO+'  ' AS [Ã≈‘¡¬…¬],ID from GRA where HMEL>='" + Format(d1, "Mm/dd/yy") + "' and HMEL<='" + Format(DateAdd("D", 1, d2), "mm/dd/yy") + "' and ARX=11"
            End If
    
610         Data1.RecordSource = sql
620         Data1.Refresh
630         SUgm_str = "001111"
    
640     ElseIf Combo1.ListIndex = 9 Then    '-----------  epistrafeises prom
    
650         If Left(Combo3.Text, 1) = "2" Then    ' .«Ã.≈…”¡√Ÿ√«”
660             sql = "select HMEL as [«Ï_ÀﬁÓÁÚ],EPO as [E˘ÌıÏﬂ·]," & "str(POS,10,2) as [¡Óﬂ·],PAR AS [A— ≈–…‘],HMEE AS [«Ã.≈…”],(SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=6 AND AYJON=KOD_BANK) AS [‘—¡–≈∆¡] ,PINAKIO AS [–…Õ¡ …œ],ONO+'  ' AS [Ã≈‘¡¬…¬],ID from GRA where HMEE>='" + Format(d1, "Mm/dd/yy") + "' and HMEE<'" + Format(DateAdd("D", 1, d2), "mm/dd/yy") + "' and ARX=12"

            Else

670             sql = "select HMEL as [«Ï_ÀﬁÓÁÚ],EPO as [E˘ÌıÏﬂ·]," & "str(POS,10,2) as [¡Óﬂ·],PAR AS [A— ≈–…‘],HMEE AS [«Ã.≈…”],(SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=6 AND AYJON=KOD_BANK) AS [‘—¡–≈∆¡] ,PINAKIO AS [–…Õ¡ …œ],ONO+'  ' AS [Ã≈‘¡¬…¬],ID from GRA where HMEL>='" + Format(d1, "Mm/dd/yy") + "' and HMEL<='" + Format(DateAdd("D", 1, d2), "mm/dd/yy") + "' and ARX=12"
            End If
    
680         Data1.RecordSource = sql
690         Data1.Refresh
700         SUgm_str = "001111"

        Else  '----------------------- ÔÎÂÛ Ï·ÊÈ

            Dim DD

710         DD = str(Combo1.ListIndex + 1)

720         If Combo1.ListIndex = 7 Then    ' diagramenes epitages
730             DD = "0"
            End If

740         If Val(DD) = 5 Then
750             DD = "3"
            End If

            ' ÂÎ·Ù˛Ì-ÒÔÏÁËÂıÙ˛Ì

760         If Left(Combo3.Text, 1) = "2" Then    ' .«Ã.≈…”¡√Ÿ√«”
770             sql = "select HMEL as [«Ï_ÀﬁÓÁÚ],EPO as [E˘ÌıÏﬂ·]," & "str(POS,10,2) as [¡Óﬂ·],PAR AS [A— ≈–…‘],HMEE AS [«Ã.≈…”],(SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=6 AND AYJON=KOD_BANK) AS [‘—¡–≈∆¡] ,PINAKIO AS [–…Õ¡ …œ],ONO+'  ' AS [Ã≈‘¡¬…¬],ID from GRA where HMEE>='" + Format(d1, "Mm/dd/yy") + "' and HMEE<'" + Format(DateAdd("D", 1, d2), "mm/dd/yy") + "' and ARX=" + DD

            Else

780             sql = "select HMEL as [«Ï_ÀﬁÓÁÚ],EPO as [E˘ÌıÏﬂ·]," & "str(POS,10,2) as [¡Óﬂ·],PAR AS [A— ≈–…‘],HMEE AS [«Ã.≈…”],(SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=6 AND AYJON=KOD_BANK) AS [‘—¡–≈∆¡] ,PINAKIO AS [–…Õ¡ …œ],ONO+'  ' AS [Ã≈‘¡¬…¬],ID from GRA where HMEL>='" + Format(d1, "Mm/dd/yy") + "' and HMEL<='" + Format(DateAdd("D", 1, d2), "mm/dd/yy") + "' and ARX=" + DD
            End If

790         If Len(Text1.Text) > 0 Then
800             sql = sql + " and EPO LIKE '" + Text1.Text + "%' "
            End If

810         If Val(Combo1.ListIndex) < 2 Then  '

820             If Val(Combo2.Text) > 0 Then    'TRAPEZA
830                 sql = sql + " AND KOD_BANK LIKE '%" + Format(Val(Combo2.Text), "#0") + "%'"
                End If

            End If

            '  If Val(Combo2.Text) = 2 Then  'TRAPEZA

840         If Combo1.ListIndex = 3 Then
                '    If Val(Left(Combo2.Text, 2)) > 0 Then
850             sql = sql + " AND PINAKIO=" + Format(Val(Left(PLAFON.Text, 2)), "0")
                '   End If

            End If

860         If Len(Text2.Text) > 0 Then  ' ariumow epitaghw
870             sql = sql + " AND PAR LIKE '%" + Trim(Text2.Text) + "%'"
                ' sql = sql + " AND left(PAR," + Str(Len(Trim(Text2.Text))) + ") ='" + Trim(Text2.Text) + "'"
            End If

880         If Left(Combo4.Text, 1) = "2" Then
890             sql = sql + "  ORDER BY PAR"
900         ElseIf Left(Combo4.Text, 1) = "3" Then
910             sql = sql + "  ORDER BY ID"

            Else
920             sql = sql + "  ORDER BY HMEL"
            End If

930         SUgm_str = "0010"

            Dim nn

940         Gdb.Execute "UPDATE GRA SET KOD_BANK='0' WHERE LEFT(KOD_BANK,1) NOT IN ('0','1','2','3','4','5','6','7','8','9')", nn

            On Error Resume Next

950         Data1.RecordSource = sql

            'MSFlexGrid1.ColAlignment(2) = 7
960         Data1.Refresh
970         MSFlexGrid1.columns(1).Width = 2600

            Dim X As Single

980         X = 0
            'For k = 1 To MSFlexGrid1.ROWS - 1
            '  MSFlexGrid1.TextMatrix(k, 0) = Format(k, "##") + ". " + Format(MSFlexGrid1.TextMatrix(k, 0), "DDD DD/MM/YYYY")
            '   MSFlexGrid1.TextMatrix(k, 5) = Format(MSFlexGrid1.TextMatrix(k, 5), "DD/MM/YYYY")
            '   MSFlexGrid1.TextMatrix(k, 1) = Format(MSFlexGrid1.TextMatrix(k, 1), "DD/MM/YYYY")
            '  X = X + Val(MSFlexGrid1.TextMatrix(k, 2))
            '           MSFlexGrid1.TextMatrix(k, 2) = Format(MSFlexGrid1.TextMatrix(k, 2), "###,###.00")
            '
            '          If k Mod 2 = 0 Then
            '             For L = 0 To MSFlexGrid1.ColS - 1
            '                MSFlexGrid1.row = k: MSFlexGrid1.Col = L: MSFlexGrid1.CellBackColor = &HC0E0FF    '&HC0E0FF
            '           Next
            '      End If
            '     MSFlexGrid1.TextMatrix(k, 7) = k

            'Next

            '  MSFlexGrid1.AddItem ("”˝ÌÔÎ·" & Chr(9) & "" & Chr(9) & Format(X, "###,###.00"))
            ' MSFlexGrid1.Col = 2
            ' MSFlexGrid1.row = MSFlexGrid1.ROWS - 1
            ' MSFlexGrid1.CellBackColor = vbRed

            ' MSFlexGrid1.ColAlignment(2) = 7    'de??? st????s?

            Dim SS1, SS2

990         SS1 = 0
1000        SS2 = 0
1010        Data1.Recordset.MoveFirst

1020        Do While Not Data1.Recordset.EOF
1030            SS1 = SS1 + Val(Data1.Recordset(2))
                ' SS2 = S2 + DATA1.Recordset(4)
1040            Data1.Recordset.MoveNext
            Loop

1050        MSFlexGrid1.Splits(0).columns(2).FooterText = Format(SS1, "######0.00")
            'MSFlexGrid1.Splits(0).Columns(4).FooterText = Format(SS2, "######0.00")

        End If

1060    platos = 1000

1070    platos = 0
        'For k = 0 To MSFlexGrid1.ColS - 1
        '    platos = platos + MSFlexGrid1.ColWidth(k)
        'Next

        '  MSFlexGrid1.width = K * 460 + platos

1080    '   MSFlexGrid1.width = 11175
        'For k = 0 To MSFlexGrid1.ColS - 1
        '    MSFlexGrid1.ColWidth(k) = MSFlexGrid1.ColWidth(k) * (MSFlexGrid1.width / platos) * 0.75 + 100
        'Next

        'If MSFlexGrid1.ROWS > 7 Then
        '    MSFlexGrid1.TopRow = MSFlexGrid1.ROWS - 7
        'End If

        'MSFlexGrid1.ColWidth(0) = 2100  '«Ã≈—
        'MSFlexGrid1.ColWidth(2) = 2000  'epo
        ' MSFlexGrid1.ColWidth(3) = 2000  '¡—.≈–…‘
        ' MSFlexGrid1.ColWidth(4) = 1400  '‘—¡–≈∆¡
        ' MSFlexGrid1.ColWidth(5) = 1400  '
        ' MSFlexGrid1.ColWidth(6) = 2000  'epo
        ' MSFlexGrid1.ColWidth(7) = 1400  '

1090    MSFlexGrid1.Visible = True

        'MSFlexGrid1.Refresh

1100    If ekt Then
            'typos2
1110        print3_xar UCase(sql), SUgm_str, "≈ÈÙ·„›Ú " + Combo1.Text + " ·¸ " + Format(d1, "dd/mm/yyyy") + " ›˘Ú " + Format(d2, "dd/mm/yyyy"), 0
        End If

        'If Excel Then
        '  Print4_Excel UCase(sql), "111111111111111"
        'End If

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.epit3.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.epit3.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub sygk()

        '********************************  ”’√ ≈Õ‘—Ÿ‘… ¡ *********************************
        '<EhHeader>
        On Error GoTo sygk_Err

        '</EhHeader>
        Dim synt  As String

        Dim eis   As String

        Dim pl    As String

        Dim Okeis As String

        Dim Okpl  As String

        Dim platos, k

        Dim dB  As Database, db2 As Database

        Dim X   As Single

        Dim X2  As Single

        Dim sql As String

        Dim s1  As Single, S2 As Single

        '  MSFlexGrid1.ColWidth(0) = 1500  'hme
        '   MSFlexGrid1.ColWidth(1) = 1000  'epo

100     If Combo1.ListIndex = 2 Then

110         synt = " HMEL>='" + Format(d1, "mm/dd/yy") + "' and HMEL<='" + Format(d2, "mm/dd/yy") + "'"

120         eis = " ( CASE ARX WHEN 1 THEN POS ELSE 0  END) AS EIS , "  '"iif(arx=1,pos,0) as eis ,"
130         pl = " ( CASE  WHEN ARX=2 THEN POS ELSE 0  END) AS PL , "  ' "POS as PL,"   ' "iif(arx=2,pos,0) as pl,"
140         Okeis = " ( CASE ARX WHEN 6 THEN POS ELSE 0  END) AS HDHE , "    ' "POS as HDHE,"   '"iif(arx=6,pos,0) as hdhe,"
150         Okpl = " ( CASE ARX WHEN 7 THEN POS ELSE 0  END) AS HDHP  "    ' "POS as HDHP" ' "iif(arx=7,pos,0) as hdhp"

            'Set db = OpenDatabase(gDir, False, False, gConnect)

160         Set db2 = OpenDatabase("c:\mercvb\reports\reports.mdb", False, False)

            On Error Resume Next

170         Gdb.Execute "drop table EPIT"

            On Error GoTo sygk_Err

180         Gdb.Execute "select HMEL," + eis + pl + Okeis + Okpl + ",ARX into EPIT from GRA WHERE " + synt + " order by HMEL;"

            ' Gdb.Execute "insert into epit in 'c:\mercvb\reports\reports.mdb' select hmel," + eis + pl + Okeis + Okpl + " from gra where " + synt + " order by hmel;"

            'Data1.DatabaseName = "c:\mercvb\reports\reports.mdb"
190         Data1.ConnectionString = gConnect

200         If katamhna Then
210             sql = "select month(HMEL) as [ÏÁÌ_ÀﬁÓÁÚ] ,STR(sum(EIS),10,2) as [≈ÈÛÒ‹ÓÂÈÚ],STR(sum(PL),10,2) as [–ÎÁÒ˘Ï›Ú],ARX from EPIT where " + synt + " group by ARX,month(HMEL) HAVING SUM(EIS)>0 OR SUM(PL)>0 order by month(HMEL);"
            Else
220             sql = "select HMEL as [«ÏÂÒ_ÀﬁÓÁÚ] ,STR(sum(EIS),10,2) as [≈ÈÛÒ‹ÓÂÈÚ],STR(sum(PL),10,2) as [–ÎÁÒ˘Ï›Ú],ARX from EPIT where " + synt + " group by HMEL,ARX order by HMEL;"
            End If

230         Data1.RecordSource = sql
240         Data1.Refresh
            'MSFlexGrid1.ColAlignment(4) = 7 'right
            ' MSFlexGrid1.ColAlignment(1) = 7    'right
            '  MSFlexGrid1.ColAlignment(2) = 7    'right

250         s1 = 0
260         S2 = 0
270         Data1.Recordset.MoveFirst

280         Do While Not Data1.Recordset.EOF
290             s1 = s1 + Data1.Recordset(1)
300             S2 = S2 + Data1.Recordset(2)
310             Data1.Recordset.MoveNext
            Loop

320         MSFlexGrid1.Splits(0).columns(1).FooterText = Format(s1, "######0.00")
330         MSFlexGrid1.Splits(0).columns(2).FooterText = Format(S2, "######0.00")

340         X = 0
            '        For k = 1 To MSFlexGrid1.ROWS - 1
            '            If MSFlexGrid1.TextMatrix(k, 3) = 2 Then MSFlexGrid1.TextMatrix(k, 1) = ""
            '            If MSFlexGrid1.TextMatrix(k, 3) = 1 Then MSFlexGrid1.TextMatrix(k, 2) = ""
            '
            '
            '            If katamhna Then
            '            Else
            '
            '                '  MSFlexGrid1.TextMatrix(K, 0) = Format(MSFlexGrid1.TextMatrix(K, 0), "dd/mm/yyyy")
            '
            '                MSFlexGrid1.TextMatrix(k, 0) = Format(MSFlexGrid1.TextMatrix(k, 0), "DDD DD/MM/YYYY")
            '
            '            End If
            '
            '            X = X + Val(MSFlexGrid1.TextMatrix(k, 1))
            '            X2 = X2 + Val(MSFlexGrid1.TextMatrix(k, 2))
            '         '   MSFlexGrid1.ColWidth(3) = 0
            '
            '

            '       Next

            '  MSFlexGrid1.AddItem "”˝ÌÔÎ·" & Chr(9) & Format(X, "###,###.00") & Chr(9) & Format(X2, "###,###.00")
            '  MSFlexGrid1.Col = 2
            '  MSFlexGrid1.row = MSFlexGrid1.ROWS - 1
            '  MSFlexGrid1.CellBackColor = vbRed

            'MSFlexGrid1.ColAlignment(5) = 7 'right

            'eis = Format(Calc_Suma(Data1.Recordset, "??sp???e??"), "###,###.00")
            'pl = Format(Calc_Suma(Data1.Recordset, "?????µ??"), "###,###.00")

            'Okeis = Format(Calc_Suma(Data1.Recordset, "?d???sp?a??e?se?"), "###,###.00")
            'Okpl = Format(Calc_Suma(Data1.Recordset, "?d??????µ??e?"), "###,###.00")

            ' p??s??t? se??? ??a ta s????a

            'MSFlexGrid1.AddItem ("" & Chr(9) & "S????a" & Chr(9) & eis & Chr(9) & pl & Chr(9) & Okeis & Chr(9) + Okpl)

            'MSFlexGrid1.AddItem ("" & Chr(9) & "??af???" & Chr(9) & Format(Calc_Suma(Data1.Recordset, "?????µ??") - Calc_Suma(Data1.Recordset, "??sp???e??"), "###,###.00"))

            ' ????????? ta s????a

            'MSFlexGrid1.Row = MSFlexGrid1.Rows - 2
            'MSFlexGrid1.Col = 0
            'MSFlexGrid1.RowSel = MSFlexGrid1.Rows - 1
            'MSFlexGrid1.Col = 2
            'MSFlexGrid1.CellBackColor = vbRed
            'MSFlexGrid1.Col = 3
            'MSFlexGrid1.CellBackColor = vbRed
            'MSFlexGrid1.ColAlignment(3) = 7

        Else

            ' MSFlexGrid1.ColWidth(1) = 2000
350         If katamhna Then
360             sql = "select MONTH(HMEL) as [MHN_ÀﬁÓÁÚ],str(sum(POS),10,2) as [”’ÕœÀ¡] from GRA where HMEL>='" + Format(d1, "mm/dd/yy") + "' and HMEL<='" + Format(d2, "mm/dd/yy") + "' and ARX=" + str(Combo1.ListIndex + 1) + " group by MONTH(HMEL) order by MONTH(HMEL)"
            Else
370             sql = "select CONVERT(DATETIME,HMEL,3) as [«Ï_ÀﬁÓÁÚ],str(sum(POS),10,2) as [”’ÕœÀ¡] from GRA where HMEL>='" + Format(d1, "mm/dd/yy") + "' and HMEL<='" + Format(d2, "mm/dd/yy") + "' and ARX=" + str(Combo1.ListIndex + 1) + " group by HMEL order by HMEL"
            End If

380         Data1.RecordSource = sql
            'MSFlexGrid1.ColAlignment(2) = 7
390         Data1.Refresh
            'MSFlexGrid1.ColAlignment(2) = 7
400         MSFlexGrid1.columns(1).alignment = dbgRight ' 7     'right
410         X = 0
            '        For k = 1 To MSFlexGrid1.ROWS - 1
            '            If katamhna Then
            '            Else
            '                MSFlexGrid1.TextMatrix(k, 0) = Format(MSFlexGrid1.TextMatrix(k, 0), "dd/mm/yyyy")
            '            End If
            '
            '            X = X + Val(MSFlexGrid1.TextMatrix(k, 1))
            '        Next

            '    MSFlexGrid1.AddItem ("”˝ÌÔÎ·" & Chr(9) & Format(X, "###,###.00"))
            '  MSFlexGrid1.Col = 2
            '     MSFlexGrid1.row = MSFlexGrid1.ROWS - 1
            '     MSFlexGrid1.CellBackColor = vbRed

            '  MSFlexGrid1.ColAlignment(1) = 7    'de??? st????s?

        End If

420     platos = 0
        '    For k = 0 To MSFlexGrid1.ColS - 1
        '        platos = platos + MSFlexGrid1.ColWidth(k)
        '    Next

        '  MSFlexGrid1.width = k * 150 + platos + 100    ' + 2720

430     s1 = 0
440     S2 = 0
450     Data1.Recordset.MoveFirst

460     Do While Not Data1.Recordset.EOF
470         s1 = s1 + Data1.Recordset(1)
            '   S2 = S2 + DATA1.Recordset(2)
480         Data1.Recordset.MoveNext
        Loop

490     MSFlexGrid1.Splits(0).columns(1).FooterText = Format(s1, "######0.00")
        ' MSFlexGrid1.Splits(0).Columns(2).FooterText = Format(S2, "######0.00")

        '    If MSFlexGrid1.ROWS > 7 Then
        '        MSFlexGrid1.TopRow = MSFlexGrid1.ROWS - 7
        '    End If

500     If ekt Then
            'typos2
510         print3_xar UCase(sql), "1111111111111111", "", 0
        End If

        'If Excel Then
        '  Print4_Excel UCase(sql), "111111111111111"
        'End If

        'MSFlexGrid1.Refresh

        '<EhFooter>
        Exit Sub

sygk_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.epit3.sygk " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.epit3.sygk " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>

100     Unload Me

110     If gApoMenu Then
120         SendKeys "%"
130         SendKeys "{RIGHT}+{RIGHT}+{DOWN}+{DOWN}+{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.epit3.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.epit3.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub EPIT31_Click(Index As Integer)
'    Dim a
'    a = 0
'
'
'
'End Sub


Sub epit312_Click(Index As Integer)
       
       
       Dim a   As String, mc As Integer

        Dim k   As Integer

        Dim R   As New ADODB.Recordset

        Dim r2  As New ADODB.Recordset

        Dim RR  As New ADODB.Recordset

        Dim kod As String

        Dim sql As String

        Dim B

        Dim M_TRAPEZA

        Dim KOD_TRAPEZAS As String

        Dim c            As String

        Dim hmee         As Date

        Dim mXreosi      As String

        Dim mPistosi     As String

        Dim m_ID          As Long

        Dim mmID         As Long

        Dim mHMEL        As Date

        Dim hmel         As String

        Dim mXre, mPis

        Dim mmXre, mmPis
        
        Dim mARX    As String

        Dim meidos  As String

        Dim mmHMEL  As Date

        Dim mmPoso  As Single

        Dim mmParat As String
 
        Dim mPoso   As Single

        Dim mParat  As String

100     MSFlexGrid1.Col = 1
110     B = MSFlexGrid1.Text

120     MSFlexGrid1.Col = 3
130     a = MSFlexGrid1.Text
   
140     MSFlexGrid1.Col = 0
150     hmel = MSFlexGrid1.Text '        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0)
     
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
      ' If MDIForm1.epit311(Index).Caption = FN_EPIS Then     ' FN_DIO = "ƒ…œ—»Ÿ”«"
    
    
     '   End If


End Sub






Sub epit311_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo epit311_Click_Err

        '</EhHeader>

        Dim a   As String, mc As Integer

        Dim k   As Integer

        Dim R   As New ADODB.Recordset

        Dim r2  As New ADODB.Recordset

        Dim RR  As New ADODB.Recordset

        Dim kod As String

        Dim sql As String

        Dim B

        Dim M_TRAPEZA

        Dim KOD_TRAPEZAS As String

        Dim c            As String

        Dim hmee         As Date

        Dim mXreosi      As String

        Dim mPistosi     As String

        Dim m_ID          As Long

        Dim mmID         As Long

        Dim mHMEL        As Date

        Dim hmel         As String

        Dim mXre, mPis

        Dim mmXre, mmPis
        
        Dim mARX    As String

        Dim meidos  As String

        Dim mmHMEL  As Date

        Dim mmPoso  As Single

        Dim mmParat As String
 
        Dim mPoso   As Single

        Dim mParat  As String

100     MSFlexGrid1.Col = 1
110     B = MSFlexGrid1.Text

120     MSFlexGrid1.Col = 3
130     a = MSFlexGrid1.Text
   
140     MSFlexGrid1.Col = 0
150     hmel = MSFlexGrid1.Text '        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0)

        On Error GoTo MHNYMA
    
        'Const FN_EIS = "≈…”–—¡Œ«"
        'Const FN_PLI = "–À«—ŸÃ«"
        'Const FN_PLA = "”≈ –À¡÷œÕ"
        'Const FN_DIA = "ƒ…¡√—¡÷«"
        'Const FN_APO = "¡–œ”‘œÀ« ”≈ –—œÃ«»≈’‘«"
        'Const FN_EPI = "≈–…”‘—œ÷« ≈–…‘¡√«”"
        'Const FN_EPA = "≈–¡Õ¡÷œ—¡ ”‘¡ ◊≈—…¡ Ã¡”"
        'Const FN_DIO = "ƒ…œ—»Ÿ”«"
        'Const FN_EXO = "≈Œœƒœ”"
    
    
    
    
    
        
        
        
        
        
        
        
    
160     If MDIForm1.epit311(Index).Caption = FN_APO Then    ' FN_APO = "¡–œ”‘œÀ« ”≈ –—œÃ«»≈’‘«"
170         epit3200.SHOW 1
180         k = MsgBox("Õ· ÏÂÙ·ˆÂÒËÂﬂ Á ÂÈÙ·„ﬁ " + Chr(13) + B + Chr(13) + a + Chr(13) + Label9.Caption, vbYesNo)

            'Me.Caption = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
190         If k = vbYes Then
200             Gdb.Execute "UPDATE GRA SET HME_METAB='" + HME_METAB.Caption + "', ONO='" + Left(Label9.Caption + Space$(6), 6) + Me.Caption + "',ARX=3 WHERE LEFT(PAR,10)='" + Left(a, 10) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", mc
            Else

                Exit Sub

            End If

210         R.Open "SELECT * FROM GRA WHERE LEFT(PAR,8)='" + Left(a, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

            '≈Õ«Ã≈—ŸÕŸ ‘œÕ ÒÔÏÁËÂıÙﬁ

220         sql = "INSERT INTO EGG (HME,XRE,AIT,KOD,EIDOS,APA,XREOSI,PISTOSI,EID,USERID ) VALUES "
230         sql = sql + "('" + Format(Now, "MM/dd/YYYY")
240         sql = sql + "',-" + Replace(LTrim(R("pos")), ",", ".") + ",'"
250         sql = sql + Left("≈–…‘." + R("EPO"), 30) + "','"
260         sql = sql + Me.Caption + "','"
270         sql = sql + "r" + "',"
280         sql = sql + Format(Val(Left(R("PAR"), 9)), "##0000000") + ","

290         sql = sql + Replace(Format(R("POS"), "000000.00"), ",", ".") + ","
300         sql = sql + "0,"

310         sql = sql + "'E–',"
320         sql = sql + str(gUserId) + " )"
330         k = 0
340         Gdb.Execute sql, k
            '       Gdb.Execute sql, k
            '             Gdb.Execute sql, k

            '                   Gdb.Execute sql, k
350         If k > 0 Then
360             MsgBox "≈Õ«Ã≈—Ÿ»« ≈ «  ¡—‘≈À¡ ‘œ’ –—œÃ«»≈’‘«"
            Else
370             MsgBox "–—œ”œ◊« !!! ƒ≈Õ ≈Õ«Ã≈—Ÿ»« ≈ «  ¡—‘≈À¡ ‘œ’ –—œÃ«»≈’‘«"
            End If
        End If

380     If MDIForm1.epit311(Index).Caption = FN_PLI Or MDIForm1.epit311(Index).Caption = "≈…”–—¡Œ«" Then      ' FN_PLI = "–À«—ŸÃ«"

390         k = MsgBox("Õ· –À«—Ÿ»≈… Á ÂÈÙ·„ﬁ " + Chr(13) + B + Chr(13) + a, vbYesNo)
400         Me.Caption = MSFlexGrid1.Text ' .TextMatrix(MSFlexGrid1.row, 1)

410         If k = vbYes Then
420             Gdb.Execute "UPDATE GRA SET ARX=ARX+5 WHERE LEFT(PAR,8)='" + Left(a, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", mc
            End If

430         R.Open "SELECT * FROM GRA WHERE LEFT(PAR,8)='" + Left(a, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

            '≈Õ«Ã≈—ŸÕŸ  ¡… ‘œÕ ¡ÀÀ«Àœ◊—≈œ –…”‘ŸÕœÕ‘¡” ‘œÕ
440         If Len(Dir("C:\KAGI.TXT")) > 0 And R("ARX") = 9 Then  ' –À«—ŸÃ≈Õ≈” ‘œ’ –À¡÷œÕ  ********************************
                '–À«—ŸÃ« ≈–…‘¡√«” ¡–O –À¡÷œÕ
450             sql = "INSERT INTO EGG (HME,XRE,AIT,KOD,EIDOS,APA,XREOSI,PISTOSI,EID,USERID ) VALUES "
460             sql = sql + "('" + Format(Now, "MM/dd/YYYY")
                'BRISKO ‘œÕ  Ÿƒ… œ ‘œ’ –…Õ¡ …œ’

470             r2.Open "SELECT * FROM PINAKES WHERE TYPOS=18 AND AYJON=" + str(R("PINAKIO")), Gdb, adOpenForwardOnly, adLockReadOnly

480             KOD_TRAPEZAS = Format(r2("TIMH"), "0000")

490             sql = sql + "',-" + LTrim(R("pos")) + ",'"
500             sql = sql + Left("≈–…‘." + R("EPO"), 30) + "','"
510             sql = sql + KOD_TRAPEZAS + "','"
520             sql = sql + "e" + "',"
                'sql = sql + Left(R("PAR"), 9) + ","
530             sql = sql + Format(Val(Left(R("PAR"), 9)), "##0000000") + ","

540             sql = sql + "0,"
550             sql = sql + Format(R("POS"), "000000.00") + ","
560             sql = sql + "'E–…',"
570             sql = sql + str(gUserId) + " )"
580             k = 0

590             If Val(KOD_TRAPEZAS) > 0 Then
600                 Gdb.Execute sql, k
                End If

610             If k > 0 Then
620                 MsgBox "≈Õ«Ã≈—Ÿ»« ≈ œ AÀÀ«Àœ◊—≈œ”"
                End If
            End If

            '≈Õ«Ã≈—ŸÕŸ  ¡… ‘œ œÿ≈Ÿ” –…”‘ŸÕœÕ‘¡” ‘œ
630         If Len(Dir("C:\KAGI.TXT")) > 0 And R("ARX") = 7 Then  ' –À«—ŸÃ≈Õ≈” –—œÃ«»≈’‘ŸÕ  ********************************
                '–À«—ŸÃ« ≈–…‘¡√«” ¡–O œÿ≈Ÿ”
                '      R.Open "SELECT * FROM GRA WHERE LEFT(PAR,8)='" + Left(A, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockReadOnly
640             sql = "INSERT INTO EGG (HME,XRE,AIT,KOD,EIDOS,APA,XREOSI,PISTOSI,EID,USERID ) VALUES "
650             sql = sql + "('" + Format(Now, "MM/dd/YYYY")

660             For k = 0 To Combo2.ListCount - 1

670                 If Val(Left(Combo2.List(k), 2)) = R("KOD_BANK") Then
680                     KOD_TRAPEZAS = Right(Combo2.List(k), 4)
                    End If

                Next

690             sql = sql + "',-" + LTrim(R("pos")) + ",'"
700             sql = sql + Left("≈–…‘." + R("EPO"), 30) + "','"
710             sql = sql + KOD_TRAPEZAS + "','"
720             sql = sql + "e" + "',"

730             sql = sql + Format(Val(Left(R("PAR"), 9)), "##0000000") + ","
                ' sql = sql + Left(R("PAR"), 9) + ","
740             sql = sql + "0,"
750             sql = sql + Format(R("POS"), "000000.00") + ","
760             sql = sql + "'E–…',"
770             sql = sql + str(gUserId) + " )"
780             k = 0

790             If Val(KOD_TRAPEZAS) > 0 Then
800                 Gdb.Execute sql, k
                End If

810             If k > 0 Then
820                 MsgBox "≈Õ«Ã≈—Ÿ»« ≈ œ œÿ≈Ÿ”"
                End If
            End If

        End If

830     If MDIForm1.epit311(Index).Caption = FN_PLA Then     ' FN_PLa = "PLAFON"
            'If Index = 1 Then
            ' On Error GoTo epit311_Click_Err
840         F_PINAKIO = 1

850         epit3100.SHOW 1

860         Gdb.Execute "UPDATE GRA SET ARX=4,PINAKIO=" + Left(Me.Caption, 2) + " WHERE LEFT(PAR,8)='" + Left(a, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", k

870         If k > 0 Then
880             MsgBox "ÃÂÙÂˆ›ÒËÁ"
            End If

890         If Len(Dir("C:\KAGI.TXT")) > 0 Then    ' ********************************

                'ENHMERVSH ¡ÀÀ«Àœ◊—≈œ’ (–À¡÷œÕ) ‘—¡–≈∆¡”
900             R.Open "SELECT * FROM GRA WHERE LEFT(PAR,8)='" + Left(a, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

910             sql = "INSERT INTO EGG (HME,XRE,AIT,KOD,EIDOS,APA,XREOSI,PISTOSI,EID,USERID ) VALUES "
920             sql = sql + "('" + Format(Now, "MM/dd/YYYY")

930             KOD_TRAPEZAS = Right(Me.Caption, 4)

940             sql = sql + "',-" + LTrim(R("pos")) + ",'"
950             sql = sql + Left("≈–…‘." + R("EPO"), 30) + "','"

960             sql = sql + KOD_TRAPEZAS + "','"
970             sql = sql + "e" + "',"
980             sql = sql + Format(Val(Left(R("PAR"), 9)), "##0000000") + ","

990             sql = sql + Format(R("POS"), "000000.00") + ","

1000            sql = sql + "0,"

1010            sql = sql + "'E–',"
1020            sql = sql + str(gUserId) + " )"
1030            k = 0

1040            If Val(KOD_TRAPEZAS) > 0 And R("ARX") = 4 Then    ' SE PLAFON
1050                Gdb.Execute sql, k
                End If

1060            If k > 0 Then
1070                MsgBox "≈Õ«Ã≈—Ÿ»« ≈ œ ¡ÀÀ«Àœ◊—≈œ”"
                End If

            End If

            ' On Error GoTo epit311_Click_Err

        End If

1080    If MDIForm1.epit311(Index).Caption = FN_DIA Then     ' FN_DIA = "ƒ…¡√—¡÷« ≈–…‘¡√«”

            'If Index = 2 Then    ' ++++++++++++++
1090        If Len(Dir("C:\KAGI.TXT")) > 0 Then    ' ********************************
1100            k = MsgBox("Õ· ‰È·„Ò·ˆÂﬂ Á ÂÈÙ·„ﬁ " + Chr(13) + B + Chr(13) + a, vbYesNo)
1110            Me.Caption = MSFlexGrid1.Text ' TextMatrix(MSFlexGrid1.row, 1)

1120            If k = vbYes Then
1130                Gdb.Execute "UPDATE GRA SET ARX=0 WHERE LEFT(PAR,8)='" + Left(a, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", mc

1140                If mc > 0 Then    ' ------------
1150                    MsgBox "ƒ…≈√—¡÷« « ≈–…‘¡√«." + Chr(13) + "–—œ”œ◊« !!!!!  ‘¡ ‘œ–œ…«”‘≈ ‘«Õ  ¡—‘≈À¡ ‘œ’ –≈À¡‘«/–—œÃ«»≈’‘«"
                    Else
1160                    MsgBox "ƒ≈Õ ƒ…≈√—¡÷« « ≈–…‘¡√«. "
                    End If    '------------------

                End If

            Else    '' **************************************************************

1170            k = MsgBox("Õ· ‰È·„Ò·ˆÂﬂ Á ÂÈÙ·„ﬁ " + a, vbYesNo)

1180            If k = vbYes Then  ' \\\\\\\\\\\\\\\\\\\\\\\
1190
                    '‚ÒÈÛÍ˘ ÙÔ ID ‘«” ≈–…‘¡√«”
                    MSFlexGrid1.Col = MSFlexGrid1.columns.Count - 1
                    'MSFlexGrid1.Text   ID EPITAGHS

                    'R.Open "SELECT *FROM GRA WHERE ID=" + MSFlexGrid1.Text + "", Gdb, adOpenForwardOnly, adLockReadOnly  'PAR='" + a + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'"
1200                R.Open "SELECT *FROM GRA WHERE ID=" + MSFlexGrid1.Text + "", Gdb, adOpenForwardOnly, adLockReadOnly  'PAR='" + a + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'"

                    If Not R.EOF Then    ' =================
1210                    kod = R("KOD")
1220                    hmee = R("HMEE")
                        'MSFlexGrid1.Columns.Count
                    
1230                    Gdb.Execute "DELETE FROM GRA WHERE ID=" + MSFlexGrid1.Text, mc
                        RR.Open "SELECT count(*) FROM GRA WHERE ID=" + MSFlexGrid1.Text, Gdb, adOpenForwardOnly, adLockReadOnly  'PAR='" + a + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'"

1240                    If RR(0) > 1 Then
1250                        MsgBox "ƒ≈Õ ƒ…≈√—¡÷« « ≈–…‘¡√«. "

                            Exit Sub

                        End If

1260                    RR.Close

                        '  Gdb.BeginTrans
1270

1280                    If mc > 0 Then    '---------
1290                        sql = "DELETE FROM EGG WHERE IDGRA=" + MSFlexGrid1.Text ' , mc ' --------------KOD='" + kod + "' AND LEFT(AIT,7)='" + Left(a, 7) + "' AND HME ='" + Format(Right(hmee, 10), "MM/DD/YYYY") + "'"
1300                        ' Gdb.Execute sql, mc
1310                        Gdb.Execute sql, mc

                            If mc = 1 Then
                                '         Gdb.CommitTrans
1320                            MsgBox "ƒ…≈√—¡÷« « ≈–…‘¡√«  ¡… «  …Õ«”« ‘œ’ –≈À¡‘«"
                            Else
1330                            MsgBox "ƒ≈Õ ƒ…≈√—¡÷« «  …Õ«”« ‘œ’ Àœ√¡—…¡”Ãœ’ " + Chr(13) + " ƒ…¡√—¡ÿ‘≈ ‘«Õ ¡–œ ‘«Õ  …Õ«”«"
                                '       Gdb.RollbackTrans
                            End If
                        End If    ' -----------
                    End If    ' ================
                End If    ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\

            End If    ' **********************************************************

        End If    ' ++++++++++++++ ƒ…¡√—¡÷« ≈–…‘¡√«”

1340    If MDIForm1.epit311(Index).Caption = FN_DIO Then     ' FN_DIO = "ƒ…œ—»Ÿ”«"
            ' If Index = 5 Then
        
1350        If Combo1.ListIndex < 3 Then    ' –—œÃ«»≈’‘«Ú / PELATHS

                'ENHMERVSH ¡ÀÀ«Àœ◊—≈œ’ (–À¡÷œÕ) ‘—¡–≈∆¡”
1360            R.Open "SELECT * FROM GRA WHERE LEFT(PAR,10)='" + Left(a, 10) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

1370            If R.EOF Then

                    Exit Sub

                End If
           
                ' Left("≈–…‘." + R("EPO"), 30)
           
1380            mHMEL = R("HMEL")
1390            mPoso = R("POS")
1400            mParat = R("PAR")

1410            kod = R("KOD")
1420            hmee = R("HMEE")
            
1430            m_ID = R("id")

1440            mHMEL = InputBox("Õ≈¡ «Ã≈—.À«Œ≈Ÿ” ", "≈–…‘¡√« " + a, CDate(mHMEL))
1450            mPoso = InputBox("Õ≈O –œ”œ ", "≈–…‘¡√« " + a, str(mPoso))
1460            mParat = InputBox("Õ≈œ” ¡—…»Ãœ” ", "≈–…‘¡√« " + a, mParat)

                '      Gdb.Execute "UPDATE GRA SET HMEL='" + Format(Right(mHMEL, 10), "MM/DD/YYYY") + "'," _
                '      & "POS=" + Str(mPoso) + ",PAR='" + mParat + "' WHERE ID=" + Str(mID)

                '         R("HMEL") = Format(mHMEL, "MM/DD/YYYY")
                '         R("POSO") = mPoso
                '         R("PARAT") = mParat
                '         R.Update

1470            If IsNull(R("id")) Then
1480                Gdb.Execute "UPDATE GRA SET HMEL='" + Format(Right(mHMEL, 10), "MM/DD/YYYY") + "'," & "POS=" + str(mPoso) + ",PAR='" + mParat + "'   WHERE LEFT(PAR,10)='" + Left(a, 10) + "' AND HMEL='" + Format(Right(mHMEL, 10), "MM/DD/YYYY") + "'"
                          
                Else
1490                Gdb.Execute "UPDATE GRA SET HMEL='" + Format(Right(mHMEL, 10), "MM/DD/YYYY") + "'," & "POS=" + str(mPoso) + ",PAR='" + mParat + "' WHERE ID=" + str(m_ID)
                
1500                If R("ARX") = 1 Then
1510                    mXre = 0: mPis = mPoso
                    Else
1520                    mPis = 0: mXre = mPoso
                    End If
                
1530                Gdb.Execute "UPDATE EGG SET XREOSI=" + str(mXre) + ",PISTOSI=" + str(mPis) + ",XRE=" + str(-mPoso) + ",AIT='" + Left(R("PAR") + Space(10), 10) + " À«Œ" + Format(mHMEL, "DD/MM/YY") + "'   WHERE IDGRA=" + str(R("ID"))
                    '
                          
                End If

1540            R.Close

                '           Gdb.Execute "UPDATE EGG SET XRE=" + Str(mPoso) + "  WHERE KOD='" + kod + "' AND LEFT(AIT,7)='" + Left(mParat, 7) + "' AND HME ='" + Format(Right(hmee, 10), "MM/DD/YYYY") + "'"

            Else    ' SE XERIA PROMHUEYTH

1550            k = MsgBox("Õ· ≈–¡Õ≈—»≈… ”‘¡ ◊≈—…¡ Ã¡”  Á ÂÈÙ·„ﬁ " + a, vbYesNo)

1560            If k = vbYes Then  ' \\\\\\\\\\\\\\\\\\\\\\\
1570                Gdb.Execute "UPDATE GRA  SET ARX=1 WHERE LEFT(PAR,8)='" + Left(a, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", mc

1580                If mc = 1 Then
1590                    MsgBox "ok Â·ÌﬁÒËÂ"
                    Else
1600                    MsgBox "–ÒÔÛÔ˜ﬁ ‰ÂÌ  Â·ÌﬁÒËÂ"
                    End If
                End If    ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\
            End If

        End If    ' ++++++++++++++ DIORUVSH ≈–…‘¡√«”
    
1610    If MDIForm1.epit311(Index).Caption = FN_EPI Then     ' FN_EPI =  +++EPISTROFH EPITAGHS (APLHROTH)
            ' If Index = 6 Then    '

            On Error GoTo epit311_Click_Err

1620        If Combo1.ListIndex < 3 Then    ' –—œÃ«»≈’‘«Ú / PELATHS

                Dim mHM6 As String

                Dim HM6  As Date

1630            mHM6 = InputBox("«ÏÂÒÔÏÁÌﬂ· ≈ÈÛÙÒÔˆﬁÚ ")

1640            If Not IsDate(mHM6) Then
1650                MsgBox "À¡»œ” «Ã≈—œÃ«Õ…¡"

                    Exit Sub

                End If

1660            HM6 = CDate(mHM6)
           
1670            If HM6 < gEnarjh Or HM6 > gLhjh Then
1680                MsgBox "«Ã≈—œÃ«Õ…¡ ≈ ‘œ” œ—…ŸÕ ◊—«”«”"

                    Exit Sub

                End If

                'ENHMERVSH ¡ÀÀ«Àœ◊—≈œ’ (–À¡÷œÕ) ‘—¡–≈∆¡”
1690            R.Open "SELECT * FROM GRA WHERE LEFT(PAR,10)='" + Left(a, 10) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

1700            If R.EOF Then

                    Exit Sub

                End If
           
1710            If R("ARX") > 2 Then
1720                MsgBox "ƒ≈Õ ≈–…”‘—≈÷≈‘¡… .  ¡Õ‘≈ ‘…”  …Õ«”≈…” ¡–œ ‘… ◊—≈Ÿ–…”‘Ÿ”≈…”"
                End If
            
1730            If R("ARX") = 2 Then 'ÂÈÙ. ÛÂ ÒÔÏÁËÂıÙÁ
1740                mARX = "12"
1750                meidos = "r"
1760                mPistosi = Replace(Format(R("POS"), "000000.00"), ",", ".")
1770                mXreosi = "0"
                Else  ' ÂÈÙ·„Á ÂÎ·ÙÔı
1780                mARX = "11"
1790                meidos = "e"
1800                mXreosi = Replace(Format(R("POS"), "000000.00"), ",", ".")
1810                mPistosi = "0"
                End If
           
                ' Left("≈–…‘." + R("EPO"), 30)
           
1820            mmHMEL = R("HMEL")
1830            mmPoso = R("POS")
1840            mmParat = R("PAR")

1850            kod = R("KOD")
1860            hmee = R("HMEE")
           
1870            mmID = R("id")

                ' mHMEL = InputBox("Õ≈¡ «Ã≈—.À«Œ≈Ÿ” ", "≈–…‘¡√« " + a, CDate(mHMEL))
                ' mPoso = InputBox("Õ≈O –œ”œ ", "≈–…‘¡√« " + a, Str(mPoso))
                ' mParat = InputBox("Õ≈œ” ¡—…»Ãœ” ", "≈–…‘¡√« " + a, mParat)

                '      Gdb.Execute "UPDATE GRA SET HMEL='" + Format(Right(mHMEL, 10), "MM/DD/YYYY") + "'," _
                '      & "POS=" + Str(mPoso) + ",PAR='" + mParat + "' WHERE ID=" + Str(mID)

                '         R("HMEL") = Format(mHMEL, "MM/DD/YYYY")
                '         R("POSO") = mPoso
                '         R("PARAT") = mParat
                '         R.Update

1880            If IsNull(R("id")) Then
1890                Gdb.Execute "UPDATE GRA SET ARX=" + mARX + "" & "  WHERE LEFT(PAR,10)='" + Left(a, 10) + "' AND HMEL='" + Format(Right(mmHMEL, 10), "MM/DD/YYYY") + "'"
                          
                Else
1900                Gdb.Execute "UPDATE GRA SET " & " ARX=" + mARX + "  WHERE ID=" + str(R("id"))
                End If
                
1910            sql = "INSERT INTO EGG (HME,XRE,AIT,KOD,EIDOS,APA,XREOSI,PISTOSI,EID,USERID ) VALUES "
1920            sql = sql + "('" + Format(HM6, "MM/dd/YYYY")

1930            sql = sql + "',-" + Replace(LTrim(R("pos")), ",", ".") + ",'"
1940            sql = sql + Left("≈–…”.≈–." + a, 30) + "','"

1950            sql = sql + kod + "','"
1960            sql = sql + meidos + "',"
1970            sql = sql + Format(Val(Left(R("PAR"), 9)), "##0000000") + ","

1980            sql = sql + mXreosi + ","

1990            sql = sql + mPistosi + ","

2000            sql = sql + "'E–',"
2010            sql = sql + str(gUserId) + " )"
                
2020            Gdb.Execute sql, k
                '

2030            R.Close

2040            If k > 0 Then
2050                MsgBox "≈Õ«Ã≈—Ÿ»« ≈ «  ¡—‘≈À¡ ‘œ’" + IIf(meidos = "e", "–≈À¡‘«", "–—œÃ«»≈’‘«")
                End If
            
                '           Gdb.Execute "UPDATE EGG SET XRE=" + Str(mPoso) + "  WHERE KOD='" + kod + "' AND LEFT(AIT,7)='" + Left(mParat, 7) + "' AND HME ='" + Format(Right(hmee, 10), "MM/DD/YYYY") + "'"

            Else    ' SE XERIA PROMHUEYTH

2060            k = MsgBox("Õ· ≈–¡Õ≈—»≈… ”‘¡ ◊≈—…¡ Ã¡”  Á ÂÈÙ·„ﬁ " + a, vbYesNo)

2070            If k = vbYes Then  ' \\\\\\\\\\\\\\\\\\\\\\\
2080                Gdb.Execute "UPDATE GRA  SET ARX=1 WHERE LEFT(PAR,8)='" + Left(a, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", mc

2090                If mc = 1 Then
2100                    MsgBox "ok Â·ÌﬁÒËÂ"
                    Else
2110                    MsgBox "–ÒÔÛÔ˜ﬁ ‰ÂÌ  Â·ÌﬁÒËÂ"
                    End If
                End If    ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\
            End If

        End If    ' ++++++++++++++ ≈–…”‘—œ÷« ≈–…‘¡√«” ==============================================

2120    Command1_Click

        Exit Sub

MHNYMA:
2130    HandleError "EPIT3:EPIT311_click"

2140    Resume Next

        '<EhFooter>
        Exit Sub

epit311_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.epit3.epit311_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.epit3.epit311_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub d1_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo d1_KeyUp_Err

        '</EhHeader>
100     If KeyCode = vbKeyReturn Then
110         KeyCode = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
        End If

        '<EhFooter>
        Exit Sub

d1_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.epit3.d1_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.epit3.d1_KeyUp " & "at line " & Erl

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
         "in ADOMERCNEW.epit3.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.epit3.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     Me.KeyPreview = True
110     mForm_Load Me, fh, fw, ft, fl

        'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

120     Combo4.AddItem "1.KATA HMEP"
130     Combo4.AddItem "2.KATA API—»MO"
140     Combo4.AddItem "2.KATA ”≈…—¡  ¡‘¡◊Ÿ—«”«”"

150     Me.Picture = LoadPicture(gPicture)
160     Combo1.ListIndex = 0
170     MSFlexGrid1.MarqueeStyle = 4
    
180     MSFlexGrid1.Splits(0).ExtendRightColumn = True
190     MSFlexGrid1.ColumnFooters = True
        ' MSFlexGrid1.Splits(0).AlternatingRowStyle = True
200     MSFlexGrid1.EmptyRows = True
210     MSFlexGrid1.FilterBar = True
    
        'Data1.DatabaseName = gDir
220     Data1.ConnectionString = gConnect
230     d1 = Now
240     d2 = DateAdd("m", 6, Now)
250     Check1.BackColor = epit3.BackColor

        Dim R As New ADODB.Recordset

260     R.Open "SELECT *FROM PINAKES WHERE TYPOS=6 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

        ' trapezes
270     Do While Not R.EOF

280         If R("typos") = 6 Then
                ' Combo2.AddItem Str(R("AYJON")) + "  " + R("PERIGRAFH")
290             Combo2.AddItem Format(R("AYJON"), "#0") + " " + R("PERIGRAFH") + Right(Space(30) + Format(R("TIMH"), "0000"), 30)
            End If

300         R.MoveNext
        Loop

310     R.Close

320     R.Open "SELECT *FROM PINAKES WHERE TYPOS=18 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

        ' trapezes
330     Do While Not R.EOF

340         If R("typos") = 18 Then
                ' Combo2.AddItem Str(R("AYJON")) + "  " + R("PERIGRAFH")
350             PLAFON.AddItem Format(R("AYJON"), "#0") + " " + R("PERIGRAFH") + Right(Space(30) + Format(R("TIMH"), "0000"), 30)
            End If

360         R.MoveNext
        Loop

370     R.Close

        'pelatew promitheytes
380     Combo1.Text = Combo1.List(2)
390     Combo1.ListIndex = 2

400     MSFlexGrid1.columns(0).Width = 1500
410     MSFlexGrid1.columns(1).Width = 3400
420     MSFlexGrid1.columns(2).Width = 1400
430     MSFlexGrid1.columns(5).Width = 1400
440     MSFlexGrid1.columns(4).Width = 1400
450     MSFlexGrid1.columns(3).Width = 1400
460     MSFlexGrid1.columns(6).Width = 1400

        ' MSFlexGrid1.Font = "Arial"

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.epit3.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.epit3.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function Calc_Suma(ByVal CO As Long) As Single

        '<EhHeader>
        On Error GoTo Calc_Suma_Err

        '</EhHeader>
        Dim k As Long

        On Error Resume Next

100     Calc_Suma = 0
        ' For k = 0 To MSFlexGrid1.ROWS - 1
        ' Calc_Suma = Calc_Suma + Val(MSFlexGrid1.TextMatrix(k, CO))
        ' Next

        '<EhFooter>
        Exit Function

Calc_Suma_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.epit3.Calc_Suma " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.epit3.Calc_Suma " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 11865, 9405, fh, fw, ft, fl

End Sub

Private Sub Label2_Click()

        '<EhHeader>
        On Error GoTo Label2_Click_Err

        '</EhHeader>

100     UPDATE_PINAKES 6, Label2, Combo2

        ''----------- OIKOGENEIES ---------------------------
        '  Dim f As New DataEntry2, DD
        '  Dim R As New ADODB.Recordset
        ''-----------------------------------------------------------------------------------------
        ''ÃœÕ¡ƒ≈” Ã≈‘—«”«”
        '  'sql ÔÔı Ë· „ﬂÌÂÙ·È ÙÔ update ﬁ ÙÔ addnew
        ' '  f.SHOW
        '  Load f
        '
        '  Gdb.Execute "update PINAKES SET AYJON =0 WHERE AYJON IS NULL AND TYPOS=6"
        '  f.ParamGrid.TextMatrix(0, 1) = "SELECT * FROM PINAKES WHERE TYPOS=6 " '    sql
        ''Ù· Â‰ﬂ· Ôı Ë· Ï·ﬂÌÔıÌ sto ParamGrid Í·È ÛÙ· textboxes
        'f.ParamGrid.TextMatrix(0, 2) = "SELECT AYJON,PERIGRAFH " _
        '& " FROM PINAKES WHERE TYPOS=6 ORDER BY AYJON " 'ParamGrid   sql2
        '
        ''ÏÂ ÔÈ‹ ÂÌÙÔÎﬁ Ë· Í‹ÌÂÈ ÙÔ delete (ÙÔ where ÂÈÌ·È ÙÔ kod=...
        'f.ParamGrid.TextMatrix(0, 3) = "delete from PINAKES where TYPOS=6 AND " 'ParamGrid              DelSQL
        '
        ''ÔÈ¸ update query Ë· ‰ÁÏÈÔıÒ„ÂﬂÙ·È (ÒÔ·ÈÒÂÙÈÍ‹ )
        ''·Ì ‰ÂÌ Ë›Î˘ update Ù¸ÙÂ ‚‹Ê˘ f.ParamGrid.TextMatrix(0, 4) = ""
        ''                    f.ParamGrid.TextMatrix(0, 4) = "  ""UPDATE ARITMISI2 SET ARITMISI= "" + t12 +"" where ID=""+t11 "   'UpdSQL
        ' DD = " UPDATE PINAKES SET TYPOS=6 where TYPOS IS NULL AND AYJON="" + t1   "
        '' AddNewScript
        '  DD = "Sub Main()" & vbCrLf & _
        '         " DIM mCon  " & vbCrLf & _
        '         " Set mCon=CreateObject(""ADODB.Connection"")" & vbCrLf & _
        '         " mCon.Open """ + gConnect + " "" " & vbCrLf & _
        '         " mCon.execute """ + DD + " " & vbCrLf & _
        '        "Set r=Nothing" & vbCrLf & _
        '        "Set mCon=Nothing" & vbCrLf & _
        '        "End Sub"
        'f.ParamGrid.TextMatrix(0, 4) = DD
        'f.ParamGrid.TextMatrix(1, 1) = "A/A"
        'f.ParamGrid.TextMatrix(2, 1) = "‘…‘Àœ” "
        ''Ã≈√¡ÀŸÕŸ À…√œ ‘¡ LABELS
        'f.Labels(0).width = f.Labels(0).width + 1000 '* TextWidth("A")
        'f.Textboxes(0).Left = f.Textboxes(0).Left + 1000 ' * TextWidth("A")
        'f.Left = Label2.Left
        'f.Top = Label2.Top + MDIForm1.Top
        '
        '
        '
        'f.SHOW 1
        'Combo2.Clear
        ' R.Open "SELECT *FROM PINAKES WHERE TYPOS=6 ORDER BY AYJON", Gdb, adOpenDynamic, _
        '     adLockOptimistic
        'Do While Not R.EOF
        '  Combo2.AddItem R("PERIGRAFH")
        '  R.MoveNext
        'Loop
        'R.Close
        'Combo2.Refresh

        '<EhFooter>
        Exit Sub

Label2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.epit3.Label2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.epit3.Label2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub MSFlexGrid1_1Click()
    
Private Sub MSFlexGrid1_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

        '<EhHeader>
        On Error GoTo MSFlexGrid1_MouseDown_Err

        '</EhHeader>
    
        Dim a As String, mc As Integer

        Dim k As Integer

        Dim L As Integer

        ' ƒ≈Õ ¬√¡∆≈… Ã≈Õœ’ ¡Õ ≈◊Ÿ –≈À¡‘ŸÕ+–—œÃ«»≈’‘ŸÕ Ã¡∆…
100     If Combo1.ListIndex = 2 Or Combo1.ListIndex > 4 Then

            Exit Sub

        End If

        ' ”‘¡ ”’√ ≈Õ‘—Ÿ‘… ¡ Õ¡ Ã«Õ ¬√¡∆≈… Ã≈Õœ’
110     If Check1 Then

            Exit Sub

        End If

        ' ·ÒÈÛÙÂÒÔ ÍÎÈÍ  (Ï Â‰ÂÓﬂ ‚„‹ÊÂÈ ÏÂÌÔı button=2)
120     If Button = 1 Then

            Exit Sub

        End If

        ' MSFlexGrid1.row = MSFlexGrid1.row - 1

        On Error Resume Next

130     MSFlexGrid1.Col = 2

140     a = MSFlexGrid1.Text

        On Error Resume Next

150     MSFlexGrid1.Col = 0

160     If IsDate(Right(MSFlexGrid1.Text, 10)) Then
            ' OK
        Else

            Exit Sub

        End If

        Dim CurRow As Integer

170     CurRow = MSFlexGrid1.row
        '  Command1_Click

        '    For k = 1 To MSFlexGrid1.ROWS - 1
        '       If MSFlexGrid1.RowIsVisible(k) And MSFlexGrid1.CellBackColor = vbRed Then
        '
        '           If k Mod 2 = 0 Then
        '              For L = 0 To MSFlexGrid1.ColS - 1
        '                 MSFlexGrid1.row = k: MSFlexGrid1.Col = L: MSFlexGrid1.CellBackColor = &HC0E0FF
        '              Next
        '        End If
        '       If k Mod 2 = 1 Then
        '               For L = 0 To MSFlexGrid1.ColS - 1
        '              MSFlexGrid1.row = k: MSFlexGrid1.Col = L: MSFlexGrid1.CellBackColor = vbYellow
        '         Next
        '    End If
        '        End If

        '    Next

180     MSFlexGrid1.row = CurRow

        ' For k = 0 To MSFlexGrid1.ColS - 1
        '    MSFlexGrid1.Col = k
        '   MSFlexGrid1.CellBackColor = vbRed
        '   Next

        'FN_EIS = "≈…”–—¡Œ«"
        'FN_PLI = "–À«—ŸÃ«"
        'FN_PLA = "”≈ –À¡÷œÕ"
        'FN_DIA = "ƒ…¡√—¡÷«"
        'FN_APO = "¡–œ”‘œÀ« ”≈ –—œÃ«»≈’‘«"
        'FN_EPI = "≈–…”‘—œ÷« ≈–…‘¡√«”"
        'FN_EPA = "≈–¡Õ¡÷œ—¡ ”‘¡ ◊≈—…¡ Ã¡”"
        'FN_DIO = "ƒ…œ—»Ÿ”«"
        'FN_EXO = "≈Œœƒœ”"
        '

        ' On Error GoTo MSFlexGrid1_MouseDown_Err


' ≈–…”‘—œ÷« ≈–…‘¡√«” –≈À¡‘œ’ ¡–œ –—œÃ«»≈’‘«
 If Combo1.ListIndex = 4 Then


        mc = 0
           
      '  Load MDIForm1.epit311 '(mc)
       '   MDIForm1.epit311.Caption = FN_EPIS ' "≈–…”‘—œ÷«"
       
        
'         mc = mc + 1
'
'    MDIForm1.epit312(mc).Caption = FN_EXO '"≈Œœƒœ”"
'       Load MDIForm1.epit312(mc)
        '  Call Me.PopupMenu(MDIForm1.par1menu)

     ' Call MDIForm1.PopupMenu(MDIForm1.epit312)
        
        
        
        
       ' Exit Sub
Dim sql As String

Dim b2 As String
Dim a2 As String
Dim hmel As String

     MSFlexGrid1.Col = 1
     b2 = MSFlexGrid1.Text

     MSFlexGrid1.Col = 3
     a2 = MSFlexGrid1.Text
   
     MSFlexGrid1.Col = 0
    hmel = MSFlexGrid1.Text '        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0)

Dim m_ID As String
MSFlexGrid1.Col = 8
     m_ID = MSFlexGrid1.Text



'+ Label9.Caption
    
            k = MsgBox("Õ· ≈–…”‘—¡÷≈…  Á ÂÈÙ·„ﬁ " + Chr(13) + b2 + Chr(13) + a2 + Chr(13) + "«Ã≈—.À«Œ." + hmel, vbYesNo)

            'Me.Caption = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
            If k = vbYes Then
                 Gdb.Execute "UPDATE GRA SET ARX=1 WHERE ID=" + m_ID, mc
            Else
               Exit Sub
            End If
            
               
    
    
                '≈Õ«Ã≈—ŸÕŸ ‘œÕ ÒÔÏÁËÂıÙﬁ
             Dim R As New ADODB.Recordset

          R.Open "SELECT * FROM GRA WHERE LEFT(PAR,8)='" + Left(a2, 8) + "' AND HMEL='" + Format(Right(hmel, 10), "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockReadOnly
          
          sql = "INSERT INTO EGG (HME,XRE,AIT,KOD,EIDOS,APA,XREOSI,PISTOSI,EID,USERID ) VALUES "
            sql = sql + "('" + Format(Now, "MM/dd/YYYY")
            sql = sql + "',-" + Replace(LTrim(R("pos")), ",", ".") + ",'"
            sql = sql + Left("≈–…”‘—.≈–…‘." + R("EPO"), 30) + "','"
            
            Dim M_K As String
            M_K = R!ono
            M_K = Trim(mID(M_K, 7, 18))
            
            
            sql = sql + M_K + "','"
            sql = sql + "r" + "',"
            sql = sql + Format(Val(Left(R("PAR"), 9)), "##0000000") + ","

            sql = sql + Replace(Format(-R("POS"), "#######.00"), ",", ".") + ","
            sql = sql + "0,"

            sql = sql + "'E–',"
            sql = sql + str(gUserId) + " )"
            k = 0
            Gdb.Execute sql, k
            If k > 0 Then
                MsgBox "≈Õ«Ã≈—Ÿ»« ≈ «  ¡—‘≈À¡ ‘œ’ –—œÃ«»≈’‘«"
            Else
                MsgBox "–—œ”œ◊« !!! ƒ≈Õ ≈Õ«Ã≈—Ÿ»« ≈ «  ¡—‘≈À¡ ‘œ’ –—œÃ«»≈’‘«"
            End If





            Exit Sub























 End If





190     mc = 0

200     If Combo1.Text = Combo1.List(0) Then
210         MDIForm1.epit311(0).Caption = FN_EIS '"≈…”–—¡Œ«"
        Else
220         MDIForm1.epit311(0).Caption = FN_PLI '"–À«—ŸÃ«"
        End If

230     mc = mc + 1
240     Load MDIForm1.epit311(mc)
250     MDIForm1.epit311(mc).Caption = FN_PLA '"”≈ –À¡÷œÕ"
    
260     MDIForm1.epit311(mc).Enabled = True

270     If Combo1.ListIndex = 1 Then    ' –—œÃ«»≈’‘«Ú
280         MDIForm1.epit311(mc).Enabled = False
        End If

290     If Combo1.ListIndex = 3 Then    ' EINAI HDH STO PLAFON
300         MDIForm1.epit311(mc).Enabled = False
        End If

310     mc = mc + 1
320     Load MDIForm1.epit311(mc)
330     MDIForm1.epit311(mc).Caption = FN_DIA ' "ƒ…¡√—¡÷«"

340     mc = mc + 1
350     Load MDIForm1.epit311(mc)
360     MDIForm1.epit311(mc).Enabled = True
370     MDIForm1.epit311(mc).Caption = FN_APO '"¡–œ”‘œÀ« ”≈ –—œÃ«»≈’‘«"

380     If Combo1.ListIndex = 1 Then    ' –—œÃ«»≈’‘«Ú
390         MDIForm1.epit311(mc).Enabled = False
        End If
    
400     mc = mc + 1
410     Load MDIForm1.epit311(mc)
420     MDIForm1.epit311(mc).Caption = FN_EPI ' "≈–…”‘—œ÷« ≈–…‘¡√«”"
    
        '—… œ 1130L  MARIA SINTOY 25220-21100

430     If Combo1.ListIndex = 3 Then    ' EINAI HDH STO PLAFON

440         mc = mc + 1
450         Load MDIForm1.epit311(mc)
460         MDIForm1.epit311(mc).Caption = FN_EPA ' "≈–¡Õ¡÷œ—¡ ”‘¡ ◊≈—…¡ Ã¡”"

470         MDIForm1.epit311(1).Enabled = False
480         MDIForm1.epit311(2).Enabled = False
490         MDIForm1.epit311(3).Enabled = False
500         MDIForm1.epit311(0).Enabled = False

        End If

510     If Combo1.ListIndex < 3 Then    ' –—œÃ«»≈’‘«Ú / PELATHS

520         mc = mc + 1
530         Load MDIForm1.epit311(mc)
540         MDIForm1.epit311(mc).Caption = FN_DIO ' "ƒ…œ—»Ÿ”«"
      
        End If



 
 













550     mc = mc + 1
560     Load MDIForm1.epit311(mc)
570     MDIForm1.epit311(mc).Caption = FN_EXO '"≈Œœƒœ”"

        '  Call Me.PopupMenu(MDIForm1.par1menu)

580     Call Me.PopupMenu(MDIForm1.EPITAG311)

        '<EhFooter>
        Exit Sub

MSFlexGrid1_MouseDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.epit3.MSFlexGrid1_MouseDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.epit3.MSFlexGrid1_MouseDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim a
'a = MSFlexGrid1.Text

'End Sub

Private Sub plafontext_Click()

        '<EhHeader>
        On Error GoTo plafontext_Click_Err

        '</EhHeader>

100     UPDATE_PINAKES 18, PlafonText, PLAFON

        '<EhFooter>
        Exit Sub

plafontext_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.epit3.plafontext_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.epit3.plafontext_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
