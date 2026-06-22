VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form apot21 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form1"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   11520
   Begin VB.CommandButton cmdKODIKOI 
      Caption         =   "¡–œ  Ÿƒ ≈Ÿ”  Ÿƒ"
      Height          =   360
      Left            =   9000
      TabIndex        =   23
      Top             =   6960
      Width           =   2415
   End
   Begin VB.TextBox EOS 
      Height          =   285
      Left            =   10320
      TabIndex        =   22
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox APO 
      Height          =   285
      Left            =   9000
      TabIndex        =   21
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton OLES 
      Caption         =   "œÀ≈” ‘…”  …Õ«”≈…”"
      Height          =   360
      Left            =   8982
      TabIndex        =   20
      Top             =   3360
      Width           =   2505
   End
   Begin VB.CommandButton CmdPrintKarteles 
      Caption         =   "cmdPrintKarteles KRIKRI"
      Height          =   360
      Left            =   480
      TabIndex        =   19
      Top             =   7080
      Width           =   3495
   End
   Begin VB.CheckBox CHECK_PROOD 
      Caption         =   "–—œœƒ≈’‘… ¡ ’–œÀœ…–¡"
      Height          =   255
      Left            =   8982
      TabIndex        =   18
      Top             =   780
      Value           =   1  'Checked
      Width           =   2505
   End
   Begin VB.CheckBox cXRONIES 
      Caption         =   "‘≈À≈’‘¡…≈”5 ◊—œÕ…≈”"
      Height          =   255
      Left            =   8982
      TabIndex        =   16
      Top             =   1080
      Width           =   2505
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000010&
      Caption         =   "≈ÓÔ‰ÔÚ"
      Height          =   324
      Left            =   8982
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6000
      Width           =   2505
   End
   Begin VB.CheckBox Check4 
      Caption         =   "–¡À…¡ ◊—œÕ…¡"
      Height          =   270
      Left            =   8982
      TabIndex        =   13
      Top             =   5235
      Width           =   2505
   End
   Begin VB.CheckBox Check3 
      Caption         =   "≈…”¡√Ÿ√≈”"
      Height          =   255
      Left            =   8982
      TabIndex        =   12
      Top             =   4320
      Value           =   1  'Checked
      Width           =   2505
   End
   Begin VB.CheckBox Check2 
      Caption         =   " ≈Œ¡√Ÿ√≈”"
      Height          =   255
      Left            =   8982
      TabIndex        =   11
      Top             =   4800
      Value           =   1  'Checked
      Width           =   2505
   End
   Begin VB.CheckBox Check1 
      Caption         =   "EKTY–Ÿ”«"
      Height          =   255
      Left            =   8982
      TabIndex        =   10
      Top             =   3810
      Width           =   2505
   End
   Begin VB.CommandButton Command2 
      Caption         =   "’ÔÎÔ„ÈÛÏ¸Ú"
      Height          =   315
      Left            =   8982
      TabIndex        =   9
      Top             =   2895
      Width           =   2505
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   8982
      TabIndex        =   5
      Top             =   1980
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   450
      _Version        =   393216
      Format          =   294912001
      CurrentDate     =   38947
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   3255
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   375
      Left            =   1320
      Top             =   6720
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   255
      Left            =   8982
      TabIndex        =   6
      Top             =   2400
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   450
      _Version        =   393216
      Format          =   294912001
      CurrentDate     =   38947
   End
   Begin TrueOleDBGrid80.TDBGrid G1 
      Bindings        =   "apot21.frx":0000
      Height          =   5520
      Left            =   315
      TabIndex        =   17
      Top             =   780
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   9737
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
      Splits(0).PartialRightColumn=   0   'False
      Splits(0).MarqueeStyle=   4
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
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   3
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTipsWidth   =   0
      DataView        =   2
      GroupByCaption  =   " ‹ÌÙÂ click ÛÙÔÌ ÙﬂÙÎÔ ÛÙﬁÎÁÚ „È· Ù·ÓÈÌ¸ÏÁÛÁ ﬁ ”˝ÒÙÂ ÙÔÌ „È· ÔÏ·‰ÔÔﬂÁÛÁ"
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
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5,.bgpicMode=2,.bgbmp=1"
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
      _StyleDefs(51)  =   ":id=35,.appearance=3,.borderColor=&H80FF&"
      _StyleDefs(52)  =   "Named:id=36:Selected"
      _StyleDefs(53)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(54)  =   "Named:id=37:Caption"
      _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(56)  =   "Named:id=38:HighlightRow"
      _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(58)  =   "Named:id=39:EvenRow"
      _StyleDefs(59)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(60)  =   "Named:id=40:OddRow"
      _StyleDefs(61)  =   ":id=40,.parent=33"
      _StyleDefs(62)  =   "Named:id=41:RecordSelector"
      _StyleDefs(63)  =   ":id=41,.parent=34"
      _StyleDefs(64)  =   "Named:id=42:FilterBar"
      _StyleDefs(65)  =   ":id=42,.parent=33"
      _StyleDefs(66)  =   "bmp(0):id=1,KAAAAEAAAABAAAAAAQAYAAAAAAAAMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(67)  =   "bmp(1):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(68)  =   "bmp(2):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(69)  =   "bmp(3):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwAAAAICAgAAAAICAgAAAAICAgMDAwICAgAAA"
      _StyleDefs(70)  =   "bmp(4):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(71)  =   "bmp(5):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(72)  =   "bmp(6):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwAAA"
      _StyleDefs(73)  =   "bmp(7):id=1,AAAAAAAAAAAAAAAAAICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(74)  =   "bmp(8):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///4CAgAAA"
      _StyleDefs(75)  =   "bmp(9):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(76)  =   "bmp(10):id=1,AAAAAAAAAAAAAAAAAICAgICAgMDAwAAAAAAAAAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(77)  =   "bmp(11):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(78)  =   "bmp(12):id=1,AAAAAICAgMDAwICAgMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(79)  =   "bmp(13):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgICAgAAAAAAA"
      _StyleDefs(80)  =   "bmp(14):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(81)  =   "bmp(15):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///wAAAAAAAICAgAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(82)  =   "bmp(16):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgAAAAAAAAAAA"
      _StyleDefs(83)  =   "bmp(17):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(84)  =   "bmp(18):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICA"
      _StyleDefs(85)  =   "bmp(19):id=1,gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(86)  =   "bmp(20):id=1,AAAAAAAAAICAgMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(87)  =   "bmp(21):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(88)  =   "bmp(22):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(89)  =   "bmp(23):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(90)  =   "bmp(24):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(91)  =   "bmp(25):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(92)  =   "bmp(26):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(93)  =   "bmp(27):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(94)  =   "bmp(28):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(95)  =   "bmp(29):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwMDAwICAgAAA"
      _StyleDefs(96)  =   "bmp(30):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(97)  =   "bmp(31):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(98)  =   "bmp(32):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(99)  =   "bmp(33):id=1,AAAAAAAAAICAgP///8DAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(100) =   "bmp(34):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(101) =   "bmp(35):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(102) =   "bmp(36):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(103) =   "bmp(37):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(104) =   "bmp(38):id=1,AAAAAICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(105) =   "bmp(39):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(106) =   "bmp(40):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(107) =   "bmp(41):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///4CAgAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(108) =   "bmp(42):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(109) =   "bmp(43):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(110) =   "bmp(44):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///8DAwP//"
      _StyleDefs(111) =   "bmp(45):id=1,/4CAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(112) =   "bmp(46):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(113) =   "bmp(47):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(114) =   "bmp(48):id=1,AAAAAICAgMDAwMDAwMDAwMDAwMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(115) =   "bmp(49):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///4CAgAAAAAAA"
      _StyleDefs(116) =   "bmp(50):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(117) =   "bmp(51):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///4CAgAAAAAAAAAAAAICAgAAAAAAAAAAAAAAA"
      _StyleDefs(118) =   "bmp(52):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(119) =   "bmp(53):id=1,AICAgMDAwICAgMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(120) =   "bmp(54):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwAAA"
      _StyleDefs(121) =   "bmp(55):id=1,AAAAAICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(122) =   "bmp(56):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///wAAAAAAAICAgAAAAICAgICAgICAgAAAAAAAAAAA"
      _StyleDefs(123) =   "bmp(57):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(124) =   "bmp(58):id=1,AAAAAAAAAAAAAAAAAAAAAICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(125) =   "bmp(59):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgAAA"
      _StyleDefs(126) =   "bmp(60):id=1,AICAgP///8DAwP///4CAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(127) =   "bmp(61):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(128) =   "bmp(62):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(129) =   "bmp(63):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwICAgMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(130) =   "bmp(64):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(131) =   "bmp(65):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(132) =   "bmp(66):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///wAAAAAAAICA"
      _StyleDefs(133) =   "bmp(67):id=1,gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(134) =   "bmp(68):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(135) =   "bmp(69):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(136) =   "bmp(70):id=1,AAAAAAAAAAAAAICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(137) =   "bmp(71):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(138) =   "bmp(72):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(139) =   "bmp(73):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgICAgICA"
      _StyleDefs(140) =   "bmp(74):id=1,gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(141) =   "bmp(75):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(142) =   "bmp(76):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(143) =   "bmp(77):id=1,AICAgICAgP///8DAwP///8DAwP///4CAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(144) =   "bmp(78):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(145) =   "bmp(79):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(146) =   "bmp(80):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///8DAwP///8DAwP///8DAwP///8DAwICAgAAAAAAA"
      _StyleDefs(147) =   "bmp(81):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(148) =   "bmp(82):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(149) =   "bmp(83):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///8DAwP///8DAwP//"
      _StyleDefs(150) =   "bmp(84):id=1,/8DAwP///8DAwP///8DAwICAgAAAAAAAAAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(151) =   "bmp(85):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(152) =   "bmp(86):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(153) =   "bmp(87):id=1,AICAgP///8DAwP///8DAwP///8DAwP///8DAwP///8DAwP///4CAgAAAAAAAAMDAwMDAwICAgICA"
      _StyleDefs(154) =   "bmp(88):id=1,gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(155) =   "bmp(89):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(156) =   "bmp(90):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///8DAwP///8DAwP///8DAwP///8DAwP///8DA"
      _StyleDefs(157) =   "bmp(91):id=1,wP///4CAgAAAAMDAwP///8DAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(158) =   "bmp(92):id=1,AAAAAAAAAAAAAICAgICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(159) =   "bmp(93):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///8DAwP///8DA"
      _StyleDefs(160) =   "bmp(94):id=1,wP///8DAwP///8DAwP///8DAwP///8DAwICAgAAAAAAAAMDAwMDAwAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(161) =   "bmp(95):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgP///8DAwP///8DAwICAgICAgAAAAAAAAAAA"
      _StyleDefs(162) =   "bmp(96):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(163) =   "bmp(97):id=1,AAAAAICAgMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwICAgAAAAAAAAAAA"
      _StyleDefs(164) =   "bmp(98):id=1,AAAAAAAAAAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///8DAwP//"
      _StyleDefs(165) =   "bmp(99):id=1,/8DAwP///8DAwP///4CAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DA"
      _StyleDefs(166) =   "bmp(100):id=1,wP///4CAgAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///4CAgP///4CAgAAAAICAgAAAAICAgAAA"
      _StyleDefs(167) =   "bmp(101):id=1,AICAgAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAMDAwMDAwICAgICAgAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(168) =   "bmp(102):id=1,AAAAAAAAAICAgP///8DAwP///8DAwP///8DAwP///8DAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(169) =   "bmp(103):id=1,AAAAAAAAAAAAAICAgP///8DAwP///8DAwP///4CAgAAAAAAAAAAAAAAAAAAAAICAgP///4CAgP//"
      _StyleDefs(170) =   "bmp(104):id=1,/4CAgAAAAICAgAAAAICAgAAAAICAgAAAAICAgMDAwICAgAAAAAAAAAAAAAAAAAAAAMDAwP///8DA"
      _StyleDefs(171) =   "bmp(105):id=1,wICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///8DAwP///8DAwP///8DAwP///8DA"
      _StyleDefs(172) =   "bmp(106):id=1,wICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwMDAwMDAwMDAwMDAwICAgAAAAAAA"
      _StyleDefs(173) =   "bmp(107):id=1,AAAAAAAAAAAAAAAAAICAgP///4CAgAAAAICAgAAAAICAgAAAAAAAAAAAAAAAAAAAAICAgAAAAAAA"
      _StyleDefs(174) =   "bmp(108):id=1,AAAAAAAAAAAAAAAAAAAAAMDAwMDAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwMDAwMDA"
      _StyleDefs(175) =   "bmp(109):id=1,wMDAwMDAwMDAwMDAwMDAwMDAwMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP//"
      _StyleDefs(176) =   "bmp(110):id=1,/4CAgAAAAAAAAAAAAICAgAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(177) =   "bmp(111):id=1,AAAAAAAAAAAAAMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(178) =   "bmp(112):id=1,AAAAAAAAAAAAAICAgP///4CAgP///4CAgAAAAICAgAAAAICAgAAAAICAgICAgAAAAAAAAAAAAAAA"
      _StyleDefs(179) =   "bmp(113):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwAAAAAAAAICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(180) =   "bmp(114):id=1,AICAgMDAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(181) =   "bmp(115):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///4CAgAAAAICAgAAAAICAgAAA"
      _StyleDefs(182) =   "bmp(116):id=1,AICAgMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgAAAAAAA"
      _StyleDefs(183) =   "bmp(117):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgMDAwAAAAAAAAAAAAAAAAICAgICAgAAAAAAA"
      _StyleDefs(184) =   "bmp(118):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICA"
      _StyleDefs(185) =   "bmp(119):id=1,gMDAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(186) =   "bmp(120):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICA"
      _StyleDefs(187) =   "bmp(121):id=1,gICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(188) =   "bmp(122):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwAAAAAAAAAAAAAAAAAAAAAAAAICAgAAAAAAAAAAA"
      _StyleDefs(189) =   "bmp(123):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(190) =   "bmp(124):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(191) =   "bmp(125):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgMDAwAAAAAAA"
      _StyleDefs(192) =   "bmp(126):id=1,AAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(193) =   "bmp(127):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(194) =   "bmp(128):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(195) =   "bmp(129):id=1,AAAAAAAAAAAAAAAAAICAgICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(196) =   "bmp(130):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(197) =   "bmp(131):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(198) =   "bmp(132):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgICAgAAAAAAA"
      _StyleDefs(199) =   "bmp(133):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(200) =   "bmp(134):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(201) =   "bmp(135):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICA"
      _StyleDefs(202) =   "bmp(136):id=1,gMDAwP///8DAwP///8DAwICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(203) =   "bmp(137):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(204) =   "bmp(138):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(205) =   "bmp(139):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///8DAwP///8DAwICAgAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(206) =   "bmp(140):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(207) =   "bmp(141):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(208) =   "bmp(142):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///8DAwP///8DA"
      _StyleDefs(209) =   "bmp(143):id=1,wP///8DAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(210) =   "bmp(144):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(211) =   "bmp(145):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(212) =   "bmp(146):id=1,AICAgMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(213) =   "bmp(147):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(214) =   "bmp(148):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(215) =   "bmp(149):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///4CAgMDAwICAgAAAAICAgAAAAICAgICAgAAAAAAA"
      _StyleDefs(216) =   "bmp(150):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(217) =   "bmp(151):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(218) =   "bmp(152):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwAAAAICAgAAA"
      _StyleDefs(219) =   "bmp(153):id=1,AICAgAAAAICAgMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(220) =   "bmp(154):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(221) =   "bmp(155):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(222) =   "bmp(156):id=1,AAAAAAAAAAAAAICAgMDAwAAAAAAAAAAAAAAAAAAAAICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(223) =   "bmp(157):id=1,AAAAAAAAAAAAAICAgICAgICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(224) =   "bmp(158):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(225) =   "bmp(159):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgMDAwAAAAAAAAAAAAICAgICAgAAA"
      _StyleDefs(226) =   "bmp(160):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///8DAwP///8DAwP///8DAwICAgAAAAAAA"
      _StyleDefs(227) =   "bmp(161):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(228) =   "bmp(162):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(229) =   "bmp(163):id=1,AICAgICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///8DAwP//"
      _StyleDefs(230) =   "bmp(164):id=1,/8DAwP///8DAwP///8DAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(231) =   "bmp(165):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(232) =   "bmp(166):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(233) =   "bmp(167):id=1,AAAAAAAAAMDAwP///8DAwP///8DAwP///8DAwP///8DAwP///8DAwAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(234) =   "bmp(168):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(235) =   "bmp(169):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(236) =   "bmp(170):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///8DAwP///8DAwP///8DAwP///8DA"
      _StyleDefs(237) =   "bmp(171):id=1,wP///4CAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(238) =   "bmp(172):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(239) =   "bmp(173):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///8DA"
      _StyleDefs(240) =   "bmp(174):id=1,wP///8DAwP///8DAwP///8DAwP///8DAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(241) =   "bmp(175):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(242) =   "bmp(176):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(243) =   "bmp(177):id=1,AAAAAAAAAAAAAICAgMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwICAgAAAAAAAAAAA"
      _StyleDefs(244) =   "bmp(178):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(245) =   "bmp(179):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(246) =   "bmp(180):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///4CAgMDAwICAgAAAAICAgAAA"
      _StyleDefs(247) =   "bmp(181):id=1,AICAgAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(248) =   "bmp(182):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(249) =   "bmp(183):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDA"
      _StyleDefs(250) =   "bmp(184):id=1,wICAgAAAAICAgAAAAICAgAAAAICAgAAAAICAgMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(251) =   "bmp(185):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(252) =   "bmp(186):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(253) =   "bmp(187):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwAAAAAAAAICAgAAAAAAAAAAAAAAAAAAAAICAgAAAAAAA"
      _StyleDefs(254) =   "bmp(188):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(255) =   "bmp(189):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(256) =   "bmp(190):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwICAgAAAAAAAAAAA"
      _StyleDefs(257) =   "bmp(191):id=1,AAAAAAAAAAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(258) =   "bmp(192):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(259) =   "bmp(193):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(260) =   "bmp(194):id=1,AAAAAAAAAICAgMDAwMDAwAAAAAAAAAAAAAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(261) =   "bmp(195):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(262) =   "bmp(196):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(263) =   "bmp(197):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgICAgICAgAAAAAAAAAAA"
      _StyleDefs(264) =   "bmp(198):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICA"
      _StyleDefs(265) =   "bmp(199):id=1,gICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(266) =   "bmp(200):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(267) =   "bmp(201):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(268) =   "bmp(202):id=1,AAAAAAAAAAAAAAAAAICAgMDAwP///8DAwP///8DAwICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(269) =   "bmp(203):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(270) =   "bmp(204):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(271) =   "bmp(205):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///8DAwP///8DAwICA"
      _StyleDefs(272) =   "bmp(206):id=1,gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(273) =   "bmp(207):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(274) =   "bmp(208):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP//"
      _StyleDefs(275) =   "bmp(209):id=1,/8DAwP///8DAwP///8DAwP///8DAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(276) =   "bmp(210):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(277) =   "bmp(211):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(278) =   "bmp(212):id=1,AAAAAAAAAAAAAAAAAAAAAICAgMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwICAgAAAAAAAAAAAAAAA"
      _StyleDefs(279) =   "bmp(213):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(280) =   "bmp(214):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(281) =   "bmp(215):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgAAAAICAgMDAwICAgAAAAICA"
      _StyleDefs(282) =   "bmp(216):id=1,gAAAAICAgICAgAAAAAAAAA=="
   End
   Begin VB.Label filtroPEL 
      Height          =   255
      Left            =   2262
      TabIndex        =   15
      Top             =   1440
      Width           =   2505
   End
   Begin VB.Label Label4 
      Caption         =   "≈˘Ú"
      Height          =   255
      Left            =   9000
      TabIndex        =   8
      Top             =   2400
      Width           =   780
   End
   Begin VB.Label Label3 
      Caption         =   "¡¸"
      Height          =   255
      Left            =   9000
      TabIndex        =   7
      Top             =   1980
      Width           =   780
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000004&
      Height          =   540
      Left            =   10380
      TabIndex        =   4
      Top             =   225
      Width           =   825
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000004&
      Height          =   525
      Left            =   8985
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "apot21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fFILE As String
Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Dim F_xre_APOU, F_axre_APOU, F_pis_APOU, F_apis_APOU
Attribute F_axre_APOU.VB_VarUserMemId = 1073938432
Attribute F_pis_APOU.VB_VarUserMemId = 1073938432
Attribute F_apis_APOU.VB_VarUserMemId = 1073938432

Dim PROTH_FORA As Boolean

Private Sub cmdKODIKOI_Click()


Dim R As New ADODB.Recordset

130     R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic

140     F_xre_APOU = ""
150     F_axre_APOU = ""
160     F_pis_APOU = ""
170     F_apis_APOU = ""

Dim f_AJEIS  As String
Dim f_AJEX  As String

Dim f_AAJEIS  As String
Dim f_AAJEX  As String

Dim s1 As String



f_AJEIS = 0
f_AJEX = 0

f_AAJEIS = 0
f_AAJEX = 0

180     R.MoveFirst

190     Do While Not R.EOF

200         If R("POS_APOU") = 1 Then
210             F_xre_APOU = F_xre_APOU + R("eidos")
220         ElseIf R("POS_APOU") = 2 Then
230             F_axre_APOU = F_axre_APOU + R("eidos")
240         ElseIf R("POS_APOU") = 3 Then
250             F_pis_APOU = F_pis_APOU + R("eidos")
260         ElseIf R("POS_APOU") = 4 Then
270             F_apis_APOU = F_apis_APOU + R("eidos")
            End If

              If R("AJIA_APOU") = 1 Then
                  f_AJEIS = f_AJEIS + R("eidos")
              ElseIf R("AJIA_APOU") = 2 Then
                  f_AAJEIS = f_AAJEIS + R("eidos")
              ElseIf R("AJIA_APOU") = 3 Then
                  f_AJEX = f_AJEX + R("eidos")
              ElseIf R("AJIA_APOU") = 4 Then
                  f_AAJEX = f_AAJEX + R("eidos")
              End If

280         R.MoveNext
        Loop

290     R.Close











Dim FFILE2 As String
Dim EGGTIM As String
  FFILE2 = "TEMPPROS" + Format(Now, "YYMMDDHHmm")


 filtroPEL.Caption = "APOT>=1 "
EGGTIM = " EGGTIM "



     s = "select E.KODE,CONVERT(CHAR(12),E.HME,3) as [HÏÂÒ/Ìﬂ·],E.ATIM+LEFT(PARASTAT.TITLOS,10) as [¡Ò –·Ò/ÍÔ˝],E.XRE AS [≈ÈÛ·„˘„›Ú],E.PIS AS [≈Ó·„˘„›Ú]"
     s = s + " ,PROOD as [’¸ÎÔÈÔ],E.TIMM*(100-E.EKPT)/100*(100-ISNULL(E.EKPT2,0))/100 AS [TEÀ… .TIMH], PEL.EPO,"
     s = s + "(CASE WHEN CHARINDEX(LEFT(ATIM,1),'" + f_AJEIS + "' )>0 THEN ROUND(E.POSO*E.TIMM*(100-E.EKPT)/100*(100-ISNULL(E.EKPT2,0))/100,2) ELSE (CASE WHEN CHARINDEX(LEFT(ATIM,1),'" + f_AAJEIS + "' )>0 THEN ROUND(-E.POSO*E.TIMM*(100-E.EKPT)/100*(100-ISNULL(E.EKPT2,0))/100,2) ELSE 0 END ) END ) AS [¡Œ…¡ EI”¡√],"
     s = s + "(CASE WHEN CHARINDEX(LEFT(ATIM,1),'" + f_AJEX + "' )>0 THEN ROUND(E.POSO*E.TIMM*(100-E.EKPT)/100*(100-ISNULL(E.EKPT2,0))/100,2) ELSE(CASE WHEN CHARINDEX(LEFT(ATIM,1),'" + f_AAJEX + "' )>0 THEN ROUND(-E.POSO*E.TIMM*(100-E.EKPT)/100*(100-ISNULL(E.EKPT2,0))/100,2) ELSE 0 END ) END ) AS [¡Œ…¡ EŒ¡√], "
   
     s1 = "E.TIMM as [TÈÏ ÃÔÌ],CONVERT(DECIMAL(10,2),E.EKPT) AS [≈Í%],CONVERT(DECIMAL(10,2),ISNULL(E.EKPT2,0)) AS [2Á ≈Í%],E.KERDOS AS [ ≈—ƒœ”],E.POSO AS [–œ”œ‘«‘¡],CONVERT(DECIMAL(10,2),E.FCURRENCY) AS [Œ≈Õœ ÕœÃ],E.ID  "
     s1 = s1 + " INTO " + FFILE2 + " FROM  " + EGGTIM + " E LEFT join PEL on E.EIDOS=PEL.EIDOS and E.PELKOD=PEL.KOD INNER JOIN PARASTAT ON PARASTAT.EIDOS=LEFT(E.ATIM,1) "
     s1 = s1 + " where " + filtroPEL.Caption + " AND  E.HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND E.HME<='" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "'  " + " and E.KODE>='" + apo.Text + "' AND E.KODE<='" + eos.Text + "'"
     s1 = s1 + " order by E.KODE,convert(char(10),E.HME,2),E.ATIM ;"
     
 Gdb.Execute s + s1
 


380    ' Data1.RecordSource = s + s1  ' "SELECT *FROM EGGTIM"

390  On Error GoTo 0
Dim sql As String
sql = "SELECT * FROM " + FFILE2 '+ " order by KODE,[HÏÂÒ/Ìﬂ·]"
data1.RecordSource = sql

data1.Refresh

If data1.Recordset.EOF Then
      
      
   Gdb.Execute "DROP TABLE " + FFILE2
   Exit Sub


End If


On Error GoTo 0

data1.Recordset.MoveFirst
Dim mk As String
mk = data1.Recordset("KODE")
Dim SSUM As Single
SSUM = 0
Do While Not data1.Recordset.EOF
  '[≈ÈÛ·„˘„›Ú],E.PIS AS [≈Ó·„˘„›Ú]
  If mk = data1.Recordset("KODE") Then ' STO IDIO EIDOS
       'œ 
  Else ' ¡ÀÀ¡Œ¡ ≈…ƒœ”
       mk = data1.Recordset("KODE")
       SSUM = 0 ' Œ¡Õ¡Ã«ƒ≈Õ…∆Ÿ ‘œ Õ≈œ ≈…ƒœ”
  End If
  
     SSUM = SSUM + data1.Recordset("≈ÈÛ·„˘„›Ú") - data1.Recordset("≈Ó·„˘„›Ú")
     Gdb.Execute "UPDATE " + FFILE2 + " SET [’¸ÎÔÈÔ] =" + Replace(Format(SSUM, "###########0.00"), ",", ".") + " WHERE ID=" + str(data1.Recordset("ID"))
  
  data1.Recordset.MoveNext
  


Loop

data1.Refresh




'MsgBox "OK"

data1.Refresh

If data1.Recordset.EOF Then
   
End If
   


  print3_xar "SELECT * FROM " + FFILE2, "00000000000000000000", m_ep, 0
  
  
 Gdb.Execute "DROP TABLE " + FFILE2
  
  Exit Sub




For j = 1 To 10
  g1.Splits(0).columns(j).Width = 700
Next
g1.Splits(0).columns(6).Width = 1800
g1.Splits(0).columns(1).Width = 1300
g1.Refresh

'On Error GoTo OUT

         data1.Recordset.MoveFirst
'          TDBGrid.Splits(0).columns(1).Width = 3500
         Dim sumes(0 To 20) As Single

         For k = 0 To data1.Recordset.FIELDS.Count - 1
            sumes(k) = 0
         Next



           ' On Error GoTo ektos
         Do While Not data1.Recordset.EOF
             For k = 0 To data1.Recordset.FIELDS.Count - 1
                 If IsNumeric(data1.Recordset.FIELDS(k).Value) And data1.Recordset.FIELDS(k).Type <> 202 Then

                     sumes(k) = sumes(k) + nNull(data1.Recordset.FIELDS(k).Value)
                    End If
                Next
                data1.Recordset.MoveNext
         Loop
         For k = 0 To data1.Recordset.FIELDS.Count - 1

               If sumes(k) > 0 Then

                  If k = 4 Or k = 5 Or k = 11 Or k = 12 Then
                     '„È· Ì· ÏÁÌ ÛÔıÏ·ÒÂÈ ıÔÎÔÈ·,ÙÈÏÂÚ,ÌÔÏÈÛÏ·Ù· Í·È id ·ÌÙÈÛÙÔÈ˜·
                  Else
                     g1.Splits(0).columns(k).FooterText = Format(sumes(k), "######0.00")
                  End If

                End If

         Next
          g1.AlternatingRowStyle = True

     g1.OddRowStyle.BackColor = &H8000000F   ' GRI   vbCyan
     g1.EvenRowStyle.BackColor = &HFFFFC0   'OYRANI     &H8000000F  ' GRI


777


Me.MousePointer = vbNormal

























'Dim R5 As New ADODB.Recordset
'R5.Open "SELECT * FROM EID WHERE KOD>='" + APO.Text + "' AND KOD<='" + EOS.Text + "'", Gdb, adOpenDynamic, adLockOptimistic
'
'
'
'Do While Not R5.EOF
'   Text1.Text = R5!kod
'   DoEvents
'    paint_grid
'    R5.MoveNext
'Loop


End Sub

Private Sub CmdPrintKarteles_Click()

Dim rrd As New ADODB.Recordset
rrd.Open "select * from EID WHERE AEG=2 AND KODLOG='01' ", Gdb, adOpenDynamic, adLockOptimistic

Do While Not rrd.EOF
    If IsNull(rrd!kod) Then
    Else
      Text1.Text = rrd!kod
      TEXT2.Text = CNull(rrd!ono)
      paint_grid
    End If
    rrd.MoveNext
Loop



End Sub

Private Sub Command1_Click()

        ' APOT2.SHOW
        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

100     Unload Me

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot21.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot21.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()
   paint_grid
End Sub '   Exit Sub
   
'        'Form_Activate
'        '<EhHeader>
'        On Error GoTo Command2_Click_Err
'
'        '</EhHeader>
'
'        ' STELNEI STO FIND_PARASTAT NA BREI TON A/A KAI PAIRNEI TIS PARAMETROYS APO TO PARASTAT
'        Dim k As Integer
'
'        Dim DB As Database
'
'        Dim R As New ADODB.Recordset, x As String
'
'        Dim s, SX, SP
'
'100     Me.Picture = LoadPicture(gPicture)
'
'        'If gConnect = "Access" Then
'        '   Set db = OpenDatabase(gDir, False, False)
'        'Else
'        '   Set db = OpenDatabase(gDir, False, False, gConnect)
'        'End If
'
'        Dim a
'
'        Dim AGOEPIS As String
'
'        Dim polepis As String, pol As String, ago As String, APIS
'
'110     pol = ""
'120     a = Get_AJ(pol, polepis, ago, AGOEPIS)
'
'130     R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic
'
'140     F_xre_APOU = ""
'150     F_axre_APOU = ""
'160     F_pis_APOU = ""
'170     F_apis_APOU = ""
'
'180     R.MoveFirst
'
'190     Do While Not R.EOF
'
'200         If R("POS_APOU") = 1 Then
'210             F_xre_APOU = F_xre_APOU + R("eidos")
'220         ElseIf R("POS_APOU") = 2 Then
'230             F_axre_APOU = F_axre_APOU + R("eidos")
'240         ElseIf R("POS_APOU") = 3 Then
'250             F_pis_APOU = F_pis_APOU + R("eidos")
'260         ElseIf R("POS_APOU") = 4 Then
'270             F_apis_APOU = F_apis_APOU + R("eidos")
'            End If
'
'280         R.MoveNext
'        Loop
'
'290     R.Close
'
'300     f_PALIAXRONIA = Trim(FINDPARAMETROI(1, "PAR1", "F_PALIAXRONIA", "2005", "√.18 DSN –¡À…¡” ◊—œÕ…¡”"))
'
'        Dim mConnect As String
'
'        '===========  TO PARAKATO EINAI GIA 2 XRONIES
'310     If Check4.Value = vbChecked Then    '   <> gDir Then ' palia xronia
'320         mConnect = gConnect + ";DATABASE=" + f_PALIAXRONIA
'
'        Else
'330         mConnect = gConnect
'        End If
'
'        '========================================================================================================================================
'340     s = "select '@@@' AS AA, CONVERT(CHAR(12),HME,3) as [HÏÂÒ/Ìﬂ·],ATIM as [¡Ò –·Ò/ÍÔ˝],EGGTIM.XRE AS [≈ÈÛ·„],EGGTIM.PIS AS [≈Ó·„˘„]"
'        's = s + ",iif(PATINDEX ( '%'+" + left(atim, 1) + "+'%' ,'" + F_xre_APOU + "')>0,POSO,iif(PATINDEX ( '%" + left(atim, 1) + "%' ,'" + F_axre_APOU + "')>0,-POSO,0)) AS [XREOSI] "
'        's = s + "iif(PATINDEX ( '%" + left(atim, 1) + "%' ,'" + F_pis_APOU + "')>0,POSO,iif(PATINDEX ( '%" + left(atim, 1) + "%' ,'" + F_apis_APOU + "')>0,-POSO,0)) AS [PISTOSI] "
'        '      instr('" + F_xre_APOU + "',)>0,poso,iif(instr('" + F_axre_APOU + "',left(atim,1))>0,-poso,0)) as [◊Ò›˘ÛÁ],"
'        '      s = s + "iif(instr('" + F_pis_APOU + "',left(atim,1))>0,poso,iif(instr('" + F_apis_APOU + "',left(atim,1))>0,-poso,0)) as [–ﬂÛÙ˘ÛÁ] "
'350     s = s + " ,PROOD as [’¸ÎÔÈÔ],TIMM as [TÈÏ ÃÔÌ],CONVERT(DECIMAL(10,2),EKPT) AS [≈Í%],PEL.EPO,CONVERT(DECIMAL(10,2),KERDOS) AS [ ≈—ƒœ”],ROUND(POSO*TIMM*(1-EKPT/100),2) AS [¡Œ…¡],EGGTIM.ID from "
'360     s = s + " EGGTIM LEFT join PEL on " + UCase("eggtim.eidos=pel.eidos and eggtim.pelKOD=pel.kod ")
'
'370     s = s + "where  HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "'  AND APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' "    '    order by HME;"
'
'380     If Check2.Value = 0 Then    ' ƒ≈Õ »≈ÀŸ –ŸÀ«”≈…”
'390         s = s + " AND  NOT LEFT(ATIM,1) IN (" + pol + ") "
'        End If
'
'400     If Check3.Value = 0 Then    ' ƒ≈Õ »≈ÀŸ ¡√œ—≈”
'410         s = s + " AND  NOT LEFT(ATIM,1) IN (" + ago + ") "
'        End If
'
'420     s = s + " ORDER BY HME "
'
'430     R.Open "select * FROM EGGTIM where  HME<'" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "'  AND APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' ORDER BY HME ", Gdb, adOpenDynamic
'
'440     If R.EOF Then Exit Sub
'
'450     R.MoveFirst
'
'        Dim yp As Single
'
'460     yp = 0
'
'        'Do While Not R.EOF
'        '
'        '   If Not IsNull(R("XRE")) Then yp = yp + R("XRE")
'        '   If Not IsNull(R("PIS")) Then yp = yp - R("PIS")
'        '
'        '   R("PROOD") = yp
'        '    On Error Resume Next
'        '   R.Update
'        '   R.MoveNext
'        '
'        '
'        '   Me.Caption = R("HME")
'        '
'        '
'        '
'        'Loop
'
'        On Error Resume Next
'
'470     DBGRID1.columns(2).NumberFormat = "###,###.##"
'480     DBGRID1.columns(3).NumberFormat = "###,###.##"
'490     DBGRID1.columns(4).NumberFormat = "###,###.##"
'500     DBGRID1.columns(5).NumberFormat = "###,###.##"
'
'510     G1.Visible = False
'
'520     DATA1.ConnectionString = mConnect
'530     DATA1.RecordSource = s    ' "SELECT *FROM EGGTIM"
'
'540     DATA1.Refresh
'
'
'
'
'   DATA1.Recordset.MoveFirst
'   Dim SEIS As Single, SEX As Single
'   SEIS = 0: SEX = 0
'   Do While Not DATA1.Recordset.EOF
'        SEIS = SEIS + DATA1.Recordset("≈ÈÛ·„")
'        SEX = SEX + DATA1.Recordset("≈Ó·„˘„")
'        Gdb.Execute "UPDATE EGGTIM SET PROOD=" + Replace(Format(SEIS - SEX, "######.00"), ",", ".") + " WHERE ID=" + str(DATA1.Recordset("ID"))
'        DATA1.Recordset.MoveNext
'   Loop
'  Me.Caption = "≈…”¡√Ÿ√≈” " + Format(SEIS, "######.00") + "  ≈Œ¡√Ÿ√≈” " + Format(SEX, "######.00")
'
'
'
'
'
'
'
'
'        Dim TITLOS
'
'550     If Check1 Then
'            'typos2
'560         TITLOS = apot21.Text1.Text + "  " + apot21.Text2.Text + "     ¡–œ " + Format(DTPicker1, "DD/MM/YYYY") + " ≈Ÿ” " + Format(DTPicker2, "DD/MM/YYYY")
'570         print3_xar UCase(s), "00111000000", TITLOS, 0
'        End If
'
'        'Set R = Data1.Recordset
'        'Data1.Refresh
'        'Exit Sub
'
'580     If DATA1.Recordset.RecordCount = 0 Then Exit Sub
'
'590     DATA1.Recordset.MoveFirst
'
'600     SU = Label2.Caption    '
'
'610     SX = 0
'620     SP = 0
'
'        Dim nATIM As Integer, nHME As Integer, NX As Integer, NP As Integer, NY As Integer, nT As Integer
'
'630     nATIM = 1: nHME = 0: NX = 3: NP = 4: NY = 5: nT = 6
'
'640     G1.Visible = False
'
'        Dim l As Long
'
'        Dim S_AJIA As Single
'
'650     S_AJIA = 0
'660       = 0
'670     G1.row = k
'680     G1.Col = NP: G1.CellBackColor = vbGreen
'690     G1.Col = NX: G1.CellBackColor = vbGreen
'700     G1.Col = NY: G1.CellBackColor = vbGreen
'
'710     For k = 1 To G1.ROWS - 1
'720         D = Left(G1.TextMatrix(k, nATIM), 1)    'LEFT(¡‘…Ã,1)
'730         G1.TextMatrix(k, 0) = k
'
'740         G1.TextMatrix(k, 10) = Format(Val(G1.TextMatrix(k, 10)), "####.00")
'
'750         If Val(G1.TextMatrix(k, NX)) <> 0 Then
'760             G1.TextMatrix(k, NP) = ""
'770             G1.TextMatrix(k, NX) = Format(Val(G1.TextMatrix(k, NX)), "####.00")
'
'            End If
'
'780         If Val(G1.TextMatrix(k, NP)) <> 0 Then
'790             G1.TextMatrix(k, NX) = ""
'800             G1.TextMatrix(k, NP) = Format(Val(G1.TextMatrix(k, NP)), "#####.00")
'
'            End If
'
'810         S_AJIA = S_AJIA + Val(G1.TextMatrix(k, 10))
'820         SX = SX + Val(G1.TextMatrix(k, NX))
'830         SP = SP + Val(G1.TextMatrix(k, NP))
'            'SU = sx - sp
'840         G1.TextMatrix(k, NY) = Val(SU) + SX - SP
'850         G1.TextMatrix(k, nT) = Format(Val(G1.TextMatrix(k, nT)), "######.00")
'
'860         If k Mod 2 = 0 Then
'
'870             For l = 0 To G1.ColS - 1
'880                 G1.row = k: G1.Col = l: G1.CellBackColor = &HC0E0FF
'                Next
'
'            End If
'
'        Next
'
'890     G1.ColWidth(8) = 3000  'epo
'
'900     G1.AddItem Chr(9) + Chr(9) + Chr(9) + Format(SX, "####0.00") + Chr(9) + Format(SP, "####0.00") + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Format(S_AJIA, "####0.00")
'910     G1.Visible = True
'
'        '<EhFooter>
'        Exit Sub
'
'Command2_Click_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.apot21.Command2_Click " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot21.Command2_Click " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>

Private Sub cXronies_Click()
  If cXronies.Value = vbChecked Then
      paint_grid
  End If
End Sub

'End Sub

Private Sub Form_Activate()
    If PROTH_FORA = False Then Exit Sub
     paint_grid
End Sub
     


Sub paint_grid()
        '<EhHeader>
        On Error GoTo Form_Activate_Err

        '</EhHeader>

        ' STELNEI STO FIND_PARASTAT NA BREI TON A/A KAI PAIRNEI TIS PARAMETROYS APO TO PARASTAT
        Dim k  As Integer

        Dim DB As Database

        Dim R  As New ADODB.Recordset, X As String

        Dim s, SX, SP
        
        If Len(Trim(filtroPEL.Caption)) = 0 Then
           filtroPEL.Caption = " APOT=" + Label1.Caption + " "
        End If
        
        
        

Dim GTGDB As New ADODB.Connection
f_PALIAXRONIA = Trim(FINDPARAMETROI(1, "PAR1", "F_PALIAXRONIA", "2005", "√.18 DSN –¡À…¡” ◊—œÕ…¡”"))

     If Check4.Value = vbChecked Then    '   <> gDir Then ' palia xronia
        GTGDB.Open gConnect + ";DATABASE=" + f_PALIAXRONIA
        data1.ConnectionString = gConnect + ";DATABASE=" + f_PALIAXRONIA
        
     Else
        data1.ConnectionString = gConnect
        GTGDB.Open gConnect
     End If






100    ' If PROTH_FORA = False Then Exit Sub

110     PROTH_FORA = False

120     Me.Picture = LoadPicture(gPicture)


   g1.Splits(0).ExtendRightColumn = True
   g1.ColumnFooters = True


        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

130     R.Open "SELECT *FROM PARASTAT", GTGDB, adOpenDynamic, adLockOptimistic

140     F_xre_APOU = ""
150     F_axre_APOU = ""
160     F_pis_APOU = ""
170     F_apis_APOU = ""

Dim f_AJEIS  As String
Dim f_AJEX  As String

Dim f_AAJEIS  As String
Dim f_AAJEX  As String

Dim s1 As String



f_AJEIS = 0
f_AJEX = 0

f_AAJEIS = 0
f_AAJEX = 0

180     R.MoveFirst

190     Do While Not R.EOF

200         If R("POS_APOU") = 1 Then
210             F_xre_APOU = F_xre_APOU + R("eidos")
220         ElseIf R("POS_APOU") = 2 Then
230             F_axre_APOU = F_axre_APOU + R("eidos")
240         ElseIf R("POS_APOU") = 3 Then
250             F_pis_APOU = F_pis_APOU + R("eidos")
260         ElseIf R("POS_APOU") = 4 Then
270             F_apis_APOU = F_apis_APOU + R("eidos")
            End If

              If R("AJIA_APOU") = 1 Then
                  f_AJEIS = f_AJEIS + R("eidos")
              ElseIf R("AJIA_APOU") = 2 Then
                  f_AAJEIS = f_AAJEIS + R("eidos")
              ElseIf R("AJIA_APOU") = 3 Then
                  f_AJEX = f_AJEX + R("eidos")
              ElseIf R("AJIA_APOU") = 4 Then
                  f_AAJEX = f_AAJEX + R("eidos")
              End If





280         R.MoveNext
        Loop

290     R.Close


        's = s + ",iif(PATINDEX ( '%'+" + left(atim, 1) + "+'%' ,'" + F_xre_APOU + "')>0,POSO,iif(PATINDEX ( '%" + left(atim, 1) + "%' ,'" + F_axre_APOU + "')>0,-POSO,0)) AS [XREOSI] "
        's = s + "iif(PATINDEX ( '%" + left(atim, 1) + "%' ,'" + F_pis_APOU + "')>0,POSO,iif(PATINDEX ( '%" + left(atim, 1) + "%' ,'" + F_apis_APOU + "')>0,-POSO,0)) AS [PISTOSI] "
        '      instr('" + F_xre_APOU + "',)>0,poso,iif(instr('" + F_axre_APOU + "',left(atim,1))>0,-poso,0)) as [◊Ò›˘ÛÁ],"
        '      s = s + "iif(instr('" + F_pis_APOU + "',left(atim,1))>0,poso,iif(instr('" + F_apis_APOU + "',left(atim,1))>0,-poso,0)) as [–ﬂÛÙ˘ÛÁ] "
If CHECK_PROOD.Value = vbChecked Then

     On Error Resume Next
     GTGDB.Execute "DROP TABLE " + fFILE
     
     On Error GoTo exitcode
     If cXronies.Value = vbChecked Then
        Dim SQL11 As String
        SQL11 = "SELECT * INTO " + fFILE + " FROM EGGTIM  where " + filtroPEL.Caption + " AND HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME<='" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "'  " + " and KODE='" + Text1.Text + "' "
     
        SQL11 = SQL11 + " UNION SELECT *  FROM " + f_PALIAXRONIA + ".dbo.EGGTIM  where " + filtroPEL.Caption + " AND HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME<='" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "'  " + " and KODE='" + Text1.Text + "' "
        Dim mPROP As String
        mPROP = FINDPARAMETROI(1, "PAR1", "F_PROPERSI", "", "√.19 ¬¡”« –—œ–≈—”…Õ«” ◊—œÕ…¡”")
        SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP + ".dbo.EGGTIM  where " + filtroPEL.Caption + " " + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
        Dim mPROP2 As String
        
        
        
        mPROP3 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI2", "", "√.19 ¬¡”« ◊—œÕ…¡” –—…Õ 3 XPONIA")
        If Len(mPROP3) > 0 Then
           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP3 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " AND  KODE='" + Text1.Text + "' "  '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
        End If
        
        Dim mPROP4 As String
        mPROP4 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI3", "", "√.19 ¬¡”« ◊—œÕ…¡” –—…Õ 4 XPONIA")
        If Len(mPROP4) > 0 Then
           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP4 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " " + " and KODE='" + Text1.Text + "' "  '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
        End If
        
        
          Dim mPROP5 As String
        mPROP5 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI5", "", "√.19 ¬¡”« ◊—œÕ…¡” –—œ 5≈‘…¡”")
        If Len(mPROP5) > 0 Then
           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP5 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " " + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
        End If
       
          Dim mPROP6 As String
        mPROP6 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI6", "", "√.19 ¬¡”« ◊—œÕ…¡” –—œ 6≈‘…¡”")
        If Len(mPROP6) > 0 Then
           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP6 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " " + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
        End If
        
           Dim mPROP7 As String
        mPROP7 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI7", "", "√.19 ¬¡”« ◊—œÕ…¡” –—œ 7≈‘…¡”")
        If Len(mPROP7) > 0 Then
           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP7 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " " + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
        End If
        
           Dim mPROP8 As String
        mPROP8 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI8", "", "√.19 ¬¡”« ◊—œÕ…¡” –—œ 8≈‘…¡”")
        If Len(mPROP8) > 0 Then
           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP8 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " " + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
        End If
        
        Dim mPROP9 As String
        mPROP9 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI9", "", "√.19 ¬¡”« ◊—œÕ…¡” –—œ 9≈‘…¡”")
        If Len(mPROP9) > 0 Then
           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP9 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " " + " and KODE='" + Text1.Text + "' "  '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
        End If
        
        
        Dim mPROP10 As String
        mPROP10 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI10", "", "√.19 ¬¡”« ◊—œÕ…¡” –—œ 10≈‘…¡”")
        If Len(mPROP10) > 0 Then
           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP10 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " " + " and KODE='" + Text1.Text + "' "  '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
        End If
        
        
        
        
        
        
        
        
        
        
        
        Gdb.Execute SQL11
    Else
    
        On Error GoTo exitcode
        GTGDB.Execute "SELECT * INTO " + fFILE + " FROM EGGTIM  where " + filtroPEL.Caption + " AND  HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME<='" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "'  " + " and KODE='" + Text1.Text + "' "  '  + "  Order by convert(char(10),HME,2),ATIM "
     
     
     
     
     End If
     
     
     
         
    
     
     
     
     
     
     Dim r3 As New ADODB.Recordset
     Dim SS As Single
     SS = 0
     r3.Open "SELECT * FROM " + fFILE + "  Order by convert(char(10),HME,2),ATIM", GTGDB, adOpenDynamic, adLockOptimistic
     Dim F_AUTO As String
     If r3.FIELDS("ID").Properties("IsAutoIncrement").Value = True Then
        ' F_AUTO = "ID"
    ' Else
       '  F_AUTO = "COL"   ' MONO SE MPOYGOYDH
     End If
      
      F_AUTO = "ID"
      
      
      'R3.Open "SELECT * FROM EGGTIM   where " + filtroPEL.Caption + " HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME<='" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "'  AND APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " + "  Order by convert(char(10),HME,2),ATIM ", Gdb, adOpenDynamic, adLockOptimistic
     
     Do While Not r3.EOF
         'Exit Do
         SS = SS + nNull(r3!XRE) - nNull(r3!PIS)
         GTGDB.Execute "UPDATE " + fFILE + " SET PROOD=" + str(SS) + " WHERE " + F_AUTO + "= " + str(r3(F_AUTO)) ' COL=ID"
         r3.MoveNext
     Loop
     
     s = "select CONVERT(CHAR(12),E.HME,3) as [HÏÂÒ/Ìﬂ·],E.ATIM+LEFT(PARASTAT.TITLOS,10) as [¡Ò –·Ò/ÍÔ˝],E.XRE AS [≈ÈÛ·„˘„›Ú],E.PIS AS [≈Ó·„˘„›Ú]"
     s = s + " ,PROOD as [’¸ÎÔÈÔ],E.TIMM*(100-E.EKPT)/100*(100-ISNULL(E.EKPT2,0))/100 AS [TEÀ… .TIMH], PEL.EPO,"
     '  s = s + " E.XRE*E.TIMM*(100-E.EKPT)/100 AS [¡Œ.≈…”],"
     
    ' s = s + "(CASE WHEN CHARINDEX(LEFT(ATIM,1),'" + f_AJEIS + "' )>0 THEN ROUND(E.POSO*E.TIMM*(100-E.EKPT)/100,2) ELSE 0 END ) AS [¡Œ…¡ EI”¡√],"
    ' s = s + "(CASE WHEN CHARINDEX(LEFT(ATIM,1),'" + f_AJEX + "' )>0 THEN ROUND(E.POSO*E.TIMM*(100-E.EKPT)/100,2) ELSE 0 END ) AS [¡Œ…¡ EŒ¡√ ],"
     
     
     s = s + "(CASE WHEN CHARINDEX(LEFT(ATIM,1),'" + f_AJEIS + "' )>0 THEN ROUND(E.POSO*E.TIMM*(100-E.EKPT)/100*(100-ISNULL(E.EKPT2,0))/100,2) ELSE (CASE WHEN CHARINDEX(LEFT(ATIM,1),'" + f_AAJEIS + "' )>0 THEN ROUND(-E.POSO*E.TIMM*(100-E.EKPT)/100*(100-ISNULL(E.EKPT2,0))/100,2) ELSE 0 END ) END ) AS [¡Œ…¡ EI”¡√],"
     s = s + "(CASE WHEN CHARINDEX(LEFT(ATIM,1),'" + f_AJEX + "' )>0 THEN ROUND(E.POSO*E.TIMM*(100-E.EKPT)/100*(100-ISNULL(E.EKPT2,0))/100,2) ELSE(CASE WHEN CHARINDEX(LEFT(ATIM,1),'" + f_AAJEX + "' )>0 THEN ROUND(-E.POSO*E.TIMM*(100-E.EKPT)/100*(100-ISNULL(E.EKPT2,0))/100,2) ELSE 0 END ) END ) AS [¡Œ…¡ EŒ¡√], "
   
     
     
     
     
     s1 = "E.TIMM as [TÈÏ ÃÔÌ],CONVERT(DECIMAL(10,2),E.EKPT) AS [≈Í%],CONVERT(DECIMAL(10,2),ISNULL(E.EKPT2,0)) AS [2Á ≈Í%],E.KERDOS AS [ ≈—ƒœ”],E.POSO AS [–œ”œ‘«‘¡],CONVERT(DECIMAL(10,2),E.FCURRENCY) AS [Œ≈Õœ ÕœÃ],E.ID  "
     s1 = s1 + " FROM  " + fFILE + " E LEFT join PEL on E.EIDOS=PEL.EIDOS and E.PELKOD=PEL.KOD INNER JOIN PARASTAT ON PARASTAT.EIDOS=LEFT(E.ATIM,1) "
     's = s + " inner join EGGTIM G on E.APOT=G.APOT AND E.KODE=G.KODE AND convert(char(10),E.HME,2)+E.ATIM >=convert(char(10),G.HME,2) +G.ATIM "
     s1 = s1 + " where " + filtroPEL.Caption + " AND  E.HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND E.HME<='" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "'  " + " and E.KODE='" + Text1.Text + "' "
    ' s = s + " Group By  CONVERT(CHAR(12),E.HME,3) ,E.ATIM ,E.XRE ,E.PIS ,E.POSO,E.TIMM,E.EKPT,PEL.EPO,E.KERDOS,E.FCURRENCY,E.ID,convert(char(10),E.HME,2)
      s1 = s1 + " order by convert(char(10),E.HME,2),E.ATIM ;"
     
     
     
     
     

'     s = "select CONVERT(CHAR(12),E.HME,3) as [HÏÂÒ/Ìﬂ·],E.ATIM as [¡Ò –·Ò/ÍÔ˝],E.XRE AS [≈ÈÛ·„˘„›Ú],E.PIS AS [≈Ó·„˘„›Ú]"
'     s = s + " ,SUM(G.XRE-G.PIS) as [’¸ÎÔÈÔ],E.TIMM as [TÈÏ ÃÔÌ],CONVERT(DECIMAL(10,2),E.EKPT) AS [≈Í%],PEL.EPO,"
'     s = s + " E.XRE*E.TIMM*(100-E.EKPT)/100 AS [¡Œ.≈…”],E.PIS*E.TIMM*(100-E.EKPT)/100 AS [¡Œ…¡ ≈Œ¡√],E.KERDOS AS [ ≈—ƒœ”],E.POSO AS [–œ”œ‘«‘¡],CONVERT(DECIMAL(10,2),E.FCURRENCY) AS [Œ≈Õœ ÕœÃ],E.ID  "
'     s = s + " FROM  DOKEIDKIN E LEFT join PEL on E.EIDOS=PEL.EIDOS and E.PELKOD=PEL.KOD "
'     s = s + " inner join DOKEIDKIN G on E.APOT=G.APOT AND E.KODE=G.KODE AND convert(char(10),E.HME,2)+E.ATIM >=convert(char(10),G.HME,2) +G.ATIM "
'     s = s + " where " + filtroPEL.Caption + " E.HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND E.HME<='" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "'  AND E.APOT=" + Label1.Caption + " and E.KODE='" + Text1.Text + "' "
'     s = s + " Group By  CONVERT(CHAR(12),E.HME,3) ,E.ATIM ,E.XRE ,E.PIS ,E.POSO,E.TIMM,E.EKPT,PEL.EPO,E.KERDOS,E.FCURRENCY,E.ID,convert(char(10),E.HME,2) order by convert(char(10),E.HME,2),E.ATIM ;"
        
Else

     s = "select CONVERT(CHAR(12),E.HME,3) as [HÏÂÒ/Ìﬂ·],E.ATIM as [¡Ò –·Ò/ÍÔ˝],E.XRE AS [≈ÈÛ·„˘„›Ú],E.PIS AS [≈Ó·„˘„›Ú]"
     s = s + " ,'     ' as [’¸ÎÔÈÔ],E.TIMM*(100-E.EKPT)/100*(100-ISNULL(E.EKPT2,0))/100 AS [TEÀ… .TIMH], PEL.EPO,"
     
     s = s + "(CASE WHEN CHARINDEX(LEFT(ATIM,1),'" + f_AJEIS + "' )>0 THEN ROUND(E.POSO*E.TIMM*(100-E.EKPT)/100*(100-ISNULL(E.EKPT2,0))/100,2) ELSE (CASE WHEN CHARINDEX(LEFT(ATIM,1),'" + f_AAJEIS + "' )>0 THEN ROUND(-E.POSO*E.TIMM*(100-E.EKPT)/100*(100-ISNULL(E.EKPT2,0))/100,2) ELSE 0 END ) END ) AS [¡Œ…¡ EI”¡√],"
     s = s + "(CASE WHEN CHARINDEX(LEFT(ATIM,1),'" + f_AJEX + "' )>0 THEN ROUND(E.POSO*E.TIMM*(100-E.EKPT)/100*(100-ISNULL(E.EKPT2,0))/100,2) ELSE(CASE WHEN CHARINDEX(LEFT(ATIM,1),'" + f_AAJEX + "' )>0 THEN ROUND(-E.POSO*E.TIMM*(100-E.EKPT)/100*(100-ISNULL(E.EKPT2,0))/100,2) ELSE 0 END ) END ) AS [¡Œ…¡ EŒ¡√], "
   
     
     
     
     s1 = " E.TIMM as [TÈÏ ÃÔÌ],CONVERT(DECIMAL(10,2),E.EKPT) AS [≈Í%],CONVERT(DECIMAL(10,2),ISNULL(E.EKPT2,0)) AS [2Á ≈Í%],E.KERDOS AS [ ≈—ƒœ”],E.POSO AS [–œ”œ‘«‘¡],CONVERT(DECIMAL(10,2),E.FCURRENCY) AS [Œ≈Õœ ÕœÃ],E.ID  "
     s1 = s1 + " FROM  EGGTIM E LEFT join PEL on E.EIDOS=PEL.EIDOS and E.PELKOD=PEL.KOD "
     's = s + " inner join EGGTIM G on E.APOT=G.APOT AND E.KODE=G.KODE AND convert(char(10),E.HME,2)+E.ATIM >=convert(char(10),G.HME,2) +G.ATIM "
     s1 = s1 + " where " + filtroPEL.Caption + " AND  E.HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND E.HME<='" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "'  " + " and E.KODE='" + Text1.Text + "' "
    ' s = s + " Group By  CONVERT(CHAR(12),E.HME,3) ,E.ATIM ,E.XRE ,E.PIS ,E.POSO,E.TIMM,E.EKPT,PEL.EPO,E.KERDOS,E.FCURRENCY,E.ID,convert(char(10),E.HME,2)
      s1 = s1 + " order by convert(char(10),E.HME,2),E.ATIM ;"


End If
        
        
        
'         select CONVERT(CHAR(12),t1.HME,3) as [HÏÂÒ/Ìﬂ·],t1.ATIM as [¡Ò –·Ò/ÍÔ˝],t1.XRE AS [≈ÈÛ·„˘„›Ú],t1.PIS AS [≈Ó·„˘„›Ú] ,
'
'
'SUM(T2.XRE-T2.PIS) AS YPOL ,
'
't1.TIMM as [TÈÏ ÃÔÌ],CONVERT(DECIMAL(10,2),t1.EKPT) AS [≈Í%],PEL.EPO,CONVERT(DECIMAL(10,2),t1.KERDOS) AS [ ≈—ƒœ”],t1.POSO AS [–œ”œ‘«‘¡],CONVERT(DECIMAL(10,2),t1.FCURRENCY) AS kerd ,t1.ID
'
'FROM  EGGTIM t1
'LEFT join PEL on t1.EIDOS=PEL.EIDOS AND t1.PELKOD=PEL.KOD
'INNER JOIN EGGTIM T2 ON t1.KODE=T2.KODE AND t1.ID>=T2.ID
'
'
'
'
'where  t1.HME>='01/01/2000' AND t1.HME<='01/01/2017'  AND t1.APOT= 1 and t1.KODE='7318826570663'
'Group By
' CONVERT(CHAR(12),E.HME,3) ,E.ATIM ,E.XRE ,E.PIS ,E.POSO,E.TIMM,E.EKPT,PEL.EPO,E.KERDOS,E.FCURRENCY,E.ID
'order by t1.ID

        
        
        
        
        
        
        
        
        
        
        
        
        'DBGrid1.Columns(2).NumberFormat = "###,###.##"
        'DBGrid1.Columns(3).NumberFormat = "###,###.##"
        'DBGrid1.Columns(4).NumberFormat = "###,###.##"
        'DBGrid1.Columns(5).NumberFormat = "###,###.##"
        'Data1.DatabaseName = gDir
340    'G1.Visible = False

'350     If APOT2.cXRONIES.Value = vbChecked Then
'360         DATA1.ConnectionString = gConnect + ";DATABASE=" + Trim(mID(APOT2.XRONIES.Text, 21, 20))
'        Else
'370         DATA1.ConnectionString = gConnect
'        End If


'cXRONIES

'        If cXRONIES.Value = vbChecked Then
'            DATA1.ConnectionString = gConnect + ";DATABASE=" + Trim(mID(APOT2.XRONIES.Text, 21, 20))
'        Else
'            DATA1.ConnectionString = gConnect
'        End If
Dim j As Integer
'Me.MousePointer = vbHourglass
DoEvents





'        '===========  TO PARAKATO EINAI GIA 2 XRONIES
'     If Check4.Value = vbChecked Then    '   <> gDir Then ' palia xronia
'        DATA1.ConnectionString = gConnect + ";DATABASE=" + f_PALIAXRONIA
'
'     Else
'        DATA1.ConnectionString = gConnect
'     End If


380     data1.RecordSource = s + s1  ' "SELECT *FROM EGGTIM"

390     data1.Refresh

data1.Refresh

If data1.Recordset.EOF Then
   Exit Sub
   
End If
   

For j = 1 To 10
  g1.Splits(0).columns(j).Width = 700
Next
g1.Splits(0).columns(6).Width = 1800
g1.Splits(0).columns(1).Width = 1300
g1.Refresh

'On Error GoTo OUT

         data1.Recordset.MoveFirst
'          TDBGrid.Splits(0).columns(1).Width = 3500
         Dim sumes(0 To 20) As Single
         
         For k = 0 To data1.Recordset.FIELDS.Count - 1
            sumes(k) = 0
         Next
         
         
         
           ' On Error GoTo ektos
         Do While Not data1.Recordset.EOF
             For k = 0 To data1.Recordset.FIELDS.Count - 1
                 If IsNumeric(data1.Recordset.FIELDS(k).Value) And data1.Recordset.FIELDS(k).Type <> 202 Then
                 
                     sumes(k) = sumes(k) + nNull(data1.Recordset.FIELDS(k).Value)
                    End If
                Next
                data1.Recordset.MoveNext
         Loop
         For k = 0 To data1.Recordset.FIELDS.Count - 1

               If sumes(k) > 0 Then
                  
                  If k = 4 Or k = 5 Or k = 11 Or k = 12 Then
                     '„È· Ì· ÏÁÌ ÛÔıÏ·ÒÂÈ ıÔÎÔÈ·,ÙÈÏÂÚ,ÌÔÏÈÛÏ·Ù· Í·È id ·ÌÙÈÛÙÔÈ˜·
                  Else
                     g1.Splits(0).columns(k).FooterText = Format(sumes(k), "######0.00")
                  End If
                  
                End If

         Next
          g1.AlternatingRowStyle = True

     g1.OddRowStyle.BackColor = &H8000000F   ' GRI   vbCyan
     g1.EvenRowStyle.BackColor = &HFFFFC0   'OYRANI     &H8000000F  ' GRI


777


Me.MousePointer = vbNormal

























   
'   DATA1.Recordset.MoveFirst
'   Dim SEIS As Single, SEX As Single
'   SEIS = 0: SEX = 0
'   Do While Not DATA1.Recordset.EOF
'        SEIS = SEIS + DATA1.Recordset("≈ÈÛ·„˘„›Ú")
'        SEX = SEX + DATA1.Recordset("≈Ó·„˘„›Ú")
'        Gdb.Execute "UPDATE EGGTIM SET PROOD=" + Replace(Format(SEIS - SEX, "######.00"), ",", ".") + " WHERE ID=" + str(DATA1.Recordset("ID"))
'        DATA1.Recordset.MoveNext
'   Loop
'  Me.Caption = "≈…”¡√Ÿ√≈” " + Format(SEIS, "######.00") + "  ≈Œ¡√Ÿ√≈” " + Format(SEX, "######.00")

 
'550     If Check1 Then
'            'typos2
'560         TITLOS = apot21.Text1.Text + "  " + apot21.Text2.Text + "     ¡–œ " + Format(DTPicker1, "DD/MM/YYYY") + " ≈Ÿ” " + Format(DTPicker2, "DD/MM/YYYY")
'570         print3_xar UCase(s), "00111000000", TITLOS, 0
'        End If



Dim m_ep As String
m_ep = Text1.Text + " " + TEXT2.Text + "  …Õ«”≈…” ¡–œ " + Format(DTPicker1.Value, "DD/MM/YYYY") + " ≈Ÿ” " + Format(DTPicker2.Value, "DD/MM/YYYY")


400     If Check1 Then
            'typos2
             On Error GoTo 0
410         print3_xar s + s1, "001111111111111111111", m_ep, 0
        End If

        'Set R = Data1.Recordset
        'Data1.Refresh
        'Exit Sub

420   '  If DATA1.Recordset.RecordCount = 0 Then Exit Sub

430    ' DATA1.Recordset.MoveFirst

440     SU = Label2.Caption    '
   
   
   
   
   
   
   
   
   
         Exit Sub
   
   
exitcode:
   MsgBox "–ÒÔ‚ÎÁÏ· ÛÙÁÌ ÂÏˆ‹ÌÈÛÁ " + Err.Description
   Exit Sub
   
   
   
   
   
   
   
450     SX = 0
460     SP = 0

        Dim nATIM   As Integer, nHME As Integer, NX As Integer, NP As Integer, NY As Integer, nT As Integer

        Dim nkERDOS As Integer

        Dim sKerdos As Single

470     sKerdos = 0

480     nATIM = 1: nHME = 0: NX = 2: NP = 3: NY = 4: nT = 5
490     nkERDOS = 8
500     g1.Visible = False

        Dim L As Long

510       = 0
520    On Error Resume Next
g1.row = k
530     g1.Col = NP: 'G1.CellBackColor = vbGreen
540     g1.Col = NX: 'G1.CellBackColor = vbGreen
550     g1.Col = NY: 'G1.CellBackColor = vbGreen

'560     For k = 1 To G1.ROWS - 1
'570         D = Left(G1.TextMatrix(k, nATIM), 1)    'LEFT(¡‘…Ã,1)
'
'580         If Val(G1.TextMatrix(k, NX)) <> 0 Then
'590             G1.TextMatrix(k, NP) = ""
'600             G1.TextMatrix(k, NX) = Format(Val(G1.TextMatrix(k, NX)), "####.00")
'
'            End If
'
'610         If Val(G1.TextMatrix(k, NP)) <> 0 Then
'620             G1.TextMatrix(k, NX) = ""
'630             G1.TextMatrix(k, NP) = Format(Val(G1.TextMatrix(k, NP)), "#####.00")
'
'            End If
'
'640         SX = SX + Val(Replace(G1.TextMatrix(k, NX), ",", "."))
'650         SP = SP + Val(Replace(G1.TextMatrix(k, NP), ",", "."))
'660         sKerdos = sKerdos + Val(G1.TextMatrix(k, nkERDOS))
'            'SU = sx - sp
'670         G1.TextMatrix(k, NY) = Val(SU) + SX - SP
'680         G1.TextMatrix(k, nT) = Format(Val(G1.TextMatrix(k, nT)), "######.00")
'
'            'Format(Val(MSFlexGrid1.TextMatrix(k, NX)), "#####.00")
'
'690         G1.TextMatrix(k, NY) = Format(Val(Replace(G1.TextMatrix(k, NY), ",", ".")), "#####.00")
'
'700         If k Mod 2 = 0 Then
'
'710             For l = 0 To G1.ColS - 1
'720                 G1.row = k: G1.Col = l: G1.CellBackColor = &HC0E0FF
'                Next
'
'            End If
'
'        Next
'
730    ' G1.ColAlignment(8) = 7

740    ' G1.ColWidth(7) = 3000  'epo

750    ' G1.AddItem Chr(9) + Chr(9) + Format(SX, "####0.00") + Chr(9) + Format(SP, "####0.00") + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Format(sKerdos, "####0.00")
760     g1.Visible = True
770     Command1.SetFocus

        '<EhFooter>
        Exit Sub

Form_Activate_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot21.Form_Activate " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot21.Form_Activate " & "at line " & Erl

        Resume Next

        '</EhFooter>


OUT:
'OUT



End Sub

'Function F_XRE(MATIM, POSO)
'    If InStr(F_xre_APOU, Left(Atim, 1)) > 0 Then
'        F_XRE = POSO
'    End If
'    If InStr(F_axre_APOU, Left(Atim, 1)) > 0 Then
'        F_XRE = -POSO
'    End If
'    If InStr(F_pis_APOU, Left(Atim, 1)) > 0 Then
'        F_XRE = 0
'    End If
'    If InStr(F_apis_APOU, Left(Atim, 1)) > 0 Then
'        F_XRE = 0
'    End If
'End Function

'Function f_pis(MATIM, POSO)
'    Dim F_XRE
'    If InStr(F_xre_APOU, Left(Atim, 1)) > 0 Then
'        F_XRE = 0
'    End If
'    If InStr(F_axre_APOU, Left(Atim, 1)) > 0 Then
'        F_XRE = 0
'    End If
'    If InStr(F_pis_APOU, Left(Atim, 1)) > 0 Then
'        F_XRE = POSO
'    End If
'    If InStr(F_apis_APOU, Left(Atim, 1)) > 0 Then
'        F_XRE = -POSO
'    End If
'End Function


Private Sub Form_Load()
    fFILE = "TEMP" + Format(Now, "YYMMDDHHmm")
        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
        
           mForm_Load Me, fh, fw, ft, fl

100     PROTH_FORA = True

110     DTPicker1.Value = gEnarjh
120     DTPicker2.Value = gLhjh

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot21.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot21.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()
   mForm_Resize Me, 8, 11790, 7785, fh, fw, ft, fl     ' h=10545
' 11790, 9885
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
Gdb.Execute "DROP TABLE " + fFILE + ""


End Sub

Private Sub OLES_Click()
   filtroPEL.Caption = ""
End Sub
