VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form PELAT42 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   11415
   ClientLeft      =   -10695
   ClientTop       =   -1905
   ClientWidth     =   18960
   FillStyle       =   7  'Diagonal Cross
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11415
   ScaleWidth      =   18960
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000010&
      Caption         =   "≈ÓÔ‰ÔÚ"
      Height          =   735
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton deltia 
      BackColor       =   &H00FFFFC0&
      Caption         =   "≈À≈√◊œ” ƒ≈À‘…ŸÕ-‘…ÃœÀœ√…ŸÕ"
      Height          =   735
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CheckBox anapar 
      BackColor       =   &H00FF0000&
      Caption         =   "¡Ì· ·Ò·ÛÙ·ÙÈÍ¸"
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   9720
      TabIndex        =   10
      Top             =   7080
      Width           =   2460
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF0000&
      Caption         =   "≈ÍÙ˝˘ÛÁ"
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   9720
      TabIndex        =   4
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "YÔÎÔ„ÈÛÏ¸Ú"
      Height          =   540
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   2520
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   6240
      TabIndex        =   1
      Top             =   7440
      Width           =   2505
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
      Bindings        =   "pelat42.frx":0000
      Height          =   1095
      Left            =   9840
      TabIndex        =   0
      Top             =   9960
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1931
      _Version        =   393216
      BackColor       =   8421631
      Cols            =   10
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   9240
      Top             =   11640
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   270
      Left            =   4245
      TabIndex        =   6
      Top             =   7080
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   476
      _Version        =   393216
      Format          =   190578689
      CurrentDate     =   38854
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   270
      Left            =   4230
      TabIndex        =   7
      Top             =   7440
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   476
      _Version        =   393216
      Format          =   190578689
      CurrentDate     =   38854
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Bindings        =   "pelat42.frx":0014
      Height          =   5970
      Left            =   360
      TabIndex        =   13
      Top             =   600
      Width           =   15555
      _ExtentX        =   27437
      _ExtentY        =   10530
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
   Begin VB.Label lblConn 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   13080
      TabIndex        =   18
      Top             =   8280
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Label LAB_nc 
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   10680
      Width           =   3975
   End
   Begin VB.Label PEL_EIDOS 
      Height          =   375
      Left            =   480
      TabIndex        =   16
      Top             =   10080
      Width           =   3975
   End
   Begin VB.Label PEL_KOD 
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   9480
      Width           =   3975
   End
   Begin VB.Label PEL_EPO 
      Height          =   375
      Left            =   480
      TabIndex        =   14
      Top             =   8880
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      FillStyle       =   6  'Cross
      Height          =   1215
      Left            =   360
      Top             =   6600
      Width           =   15615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "·¸"
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   7110
      Width           =   525
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "›˘Ú"
      Height          =   240
      Left            =   3720
      TabIndex        =   8
      Top             =   7470
      Width           =   465
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   135
      Width           =   15555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ã¸ÌÔ ÙÔıÚ Í˘‰ÈÍÔ˝Ú"
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   6360
      TabIndex        =   2
      Top             =   7080
      Width           =   2475
   End
End
Attribute VB_Name = "PELAT42"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sumes(40) As Single

Dim ColS      As TrueOleDBGrid80.columns
Dim Col      As TrueOleDBGrid80.Column

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Private Sub Command1_Click()


     Dim dummy, PolAJ As String, ePolAj As String, PolPos As String, ePolPos As String, AgoAJ As String, eAgoAJ As String, AgoPos As String, eAgoPos As String


     dummy = Get_EIDPAR(PolAJ, ePolAj, PolPos, ePolPos, AgoAJ, eAgoAJ, AgoPos, eAgoPos)





        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

        Dim x As String

        Dim k

100     Label2.Caption = PEL_EPO + " ¡–œ " + Format(DTPicker1, "DD/MM/YYYY") + "  ≈Ÿ” " + Format(DTPicker2, "DD/MM/YYYY")

        Dim N  As Integer

 '       Dim nc As String

'110     nc = ""
'
'120     For N = 0 To pelat4.List1.ListCount - 1
'
'            'On Error Resume Next
'130         If pelat4.List1.Selected(N) Then
'140             nc = nc + "'" + Right(pelat4.List1.List(N), 1) + "',"
'            End If
'
'        Next
'
'150     If Len(nc) > 0 Then
'160         nc = Left(nc, Len(nc) - 1)
'170         nc = " left (ATIM,1) IN (" + nc + ") AND "
'        End If
 Dim mGDB As New ADODB.Connection
                mGDB.Open lblConn.Caption
                
                
180     If anapar Then
190         If Len(Text1.Text) > 0 Then
200             x = "select LEFT(ATIM,1),KODE as [ ˘‰ÈÍ¸Ú],LEFT(ONOMA,30) as [–ÂÒÈ„Ò·ˆﬁ],STR(SUM(POSO),10,2) as [–ÔÛ¸ÙÁÙ·],STR(TIMM,10,2) AS [TIMH],str(sum(TIMM*POSO),10,2) as [¡Óﬂ· ÒÔ ≈ –‘],str(sum(TIMM*POSO*(100-EKPT)/100),10,2) as [¡Óﬂ· Ã≈‘¡ ‘«Õ ≈ –‘],AVG(TIMM*(100-EKPT)/100) AS [‘…Ã.Ã≈‘¡ ‘«Õ ≈ –‘Ÿ”«] FROM EGGTIM where " + LAB_nc.Caption + " KODE IN (" + Text1.Text + ") AND HME >='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME <='" + Format(DTPicker2, "MM/DD/YYYY") + "' AND EIDOS='" + PEL_EIDOS + "' AND PELKOD='" + PEL_KOD + "'  group by KODE,LEFT(ATIM,1),ONOMA,TIMM;"
            Else
210             x = "select LEFT(ATIM,1),KODE as [ ˘‰ÈÍ¸Ú],LEFT(ONOMA,30) as [–ÂÒÈ„Ò·ˆﬁ],STR(SUM(POSO),10,2) as [–ÔÛ¸ÙÁÙ·],STR(TIMM,10,2) AS TIMH,str(sum(TIMM*POSO),10,2) as [¡Óﬂ· –—œ ≈ –‘],str(sum(TIMM*POSO*(100-EKPT)/100),10,2) as [¡Óﬂ· Ã≈‘¡ ‘«Õ ≈ –‘],AVG(TIMM*(100-EKPT)/100) AS [‘…Ã.Ã≈‘¡ ‘«Õ ≈ –‘Ÿ”«] FROM EGGTIM where " + LAB_nc.Caption + " HME >='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME <='" + Format(DTPicker2, "MM/DD/YYYY") + "' AND EIDOS='" + PEL_EIDOS + "' AND PELKOD='" + PEL_KOD + "'  group by KODE,LEFT(ATIM,1),ONOMA,TIMM;"
            End If

        Else
          
             If PEL_EIDOS = "e" Then
                 mGDB.Execute "UPDATE EGGTIM SET AJPOL =TIMM*(-XRE+PIS)*(100-EKPT)/100*(100+( SELECT TOP 1 TIMH  FROM PINAKES WHERE TYPOS=1 AND AYJON=EGGTIM.FPA)  ) /100 "
             Else
                 mGDB.Execute "UPDATE EGGTIM SET AJPOL =TIMM*(XRE-PIS)*(100-EKPT)/100*(100+( SELECT TOP 1 TIMH  FROM PINAKES WHERE TYPOS=1 AND AYJON=EGGTIM.FPA)  ) /100 "
             End If
             


220         If Len(Text1.Text) > 0 Then
230             x = "select KODE as [ ˘‰ÈÍ¸Ú],LEFT(ONOMA,30) as [–ÂÒÈ„Ò·ˆﬁ], SUM(XRE) AS [¡√œ—≈”] , SUM(PIS) as [–ŸÀ«”≈…”],STR(TIMM,10,2) AS [TIMH],str(sum(TIMM*POSO),10,2) as [¡Óﬂ·] FROM EGGTIM where  KODE IN (" + Text1.Text + ") AND HME >='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME <='" + Format(DTPicker2, "MM/DD/YYYY") + "' AND EIDOS='" + PEL_EIDOS + "' AND PELKOD='" + PEL_KOD + "'  group by KODE,ONOMA,TIMM;"
            Else

                On Error Resume Next
                
               
                

240             mGDB.Execute "DROP TABLE DOKPELAT42"
250             mGDB.Execute "SELECT DISTINCT KODE INTO DOKPELAT42 FROM EGGTIM WHERE EIDOS='" + PEL_EIDOS + "' AND PELKOD='" + PEL_KOD + "'"
               ' Gdb.Execute "UPDATE DOKPELAT42 SET AJIPOL =TIMM*POSO*EKPT/100*(100+( SELECT TIMH  FROM PINAKES WHERE AYJON=FPA)  ) /100 "
                
                
                
                ' x = "select KODE as [ ˘‰ÈÍ¸Ú],LEFT(ONOMA,30) as [–ÂÒÈ„Ò·ˆﬁ],SUM(XRE) AS [¡√œ—≈”] , SUM(PIS) as [–ŸÀ«”≈…”],STR(TIMM,10,2) AS TIMH,str(sum(TIMM*POSO),10,2) as [¡Óﬂ·] FROM EGGTIM where  HME >='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME <='" + Format(DTPicker2, "MM/DD/YYYY") + "' AND EIDOS='" + pelat4.data1.Recordset("EIDOS") + "' AND PELKOD='" + pelat4.data1.Recordset("kod") + "' GROUP BY KODE,ONOMA,TIMM;"

                If PEL_EIDOS = "e" Then
                    x = "select KODE AS [ Ÿƒ… œ”],LEFT(EID.ONO,30) as [–ÂÒÈ„Ò·ˆﬁ], ROUND(SUM(PIS),2) as [–ŸÀ.–œ”œ‘],"
                    x = x + "round(sum(TIMM*(-XRE+PIS)),2) AS [¡Œ…¡ –—œ ;∏ –‘],round(sum(TIMM*(-XRE+PIS)*EKPT/100),2) as [¡Œ.≈ –],"

                    x = x + "round(sum(TIMM*(-XRE+PIS)*(100-EKPT)/100),2) as [¡Óﬂ· –˘Î.],ROUND(SUM(XRE),2) as [≈–…”.–œ”],"
                    x = x + "round(sum(AJPOL  ),2) as [¡Œ.ME ÷–¡],"
                    
                     '  x = x + "round( sum( CASE WHEN LEFT(ATIM,1) IN ('P') THEN TIMM*POSO*(100-EKPT)/100 ELSE 0 END ),2) as [¡Óﬂ· ≈ÈÛ],"
                     x = x + "round( sum( TIMM*XRE*(100-EKPT)/100 ),2) as [¡Óﬂ· ≈ÈÛ],"
                    
                    x = x + "ROUND(XTI,2) AS [TIM.A√œ—],round(AVG(LTI),2) AS [MO XON.TIMH],round(AVG(EKPT),2) AS [MO EK–‘] FROM EGGTIM  INNER JOIN EID ON EGGTIM.KODE=EID.KOD"
                    x = x + " where EIDOS='" + PEL_EIDOS + "' AND PELKOD='" + PEL_KOD + "' AND HME >='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME <='" + Format(DTPicker2, "MM/DD/YYYY") + "' AND  KODE IN (SELECT KODE FROM DOKPELAT42) GROUP BY EID.ONO,KODE,XTI"
                Else
                    x = "select KODE AS [ Ÿƒ… œ”],LEFT(EID.ONO,30) as [–ÂÒÈ„Ò·ˆﬁ],ROUND(SUM(XRE),2) AS [¡√œ—≈”] ,str(sum((XRE-PIS)*TIMM),10,2) AS [¡Œ…¡ –—œ ;∏ –‘],str(sum(TIMM*(XRE-PIS)*EKPT/100),10,2) as [¡Œ.≈ –], str(sum(TIMM*(XRE-PIS)*(100-EKPT)/100),10,2) as [¡Óﬂ·],ROUND(XTI,2) AS [TIM.A√œ—],AVG(LTI) AS [MO XON.TIMH],AVG(EKPT) AS [MO EK–‘], "
                    x = x + " round(sum(AJPOL  ),2) as [¡Œ.ME ÷–¡]    FROM EGGTIM  INNER JOIN EID ON EGGTIM.KODE=EID.KOD"
                    x = x + " where  PELKOD='" + PEL_KOD + "' AND HME >='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME <='" + Format(DTPicker2, "MM/DD/YYYY") + "' AND  KODE IN (SELECT KODE FROM DOKPELAT42) GROUP BY EID.ONO,KODE,XTI"
                End If
            End If
        End If

        On Error GoTo Command1_Click_Err

280     data1.ConnectionString = lblConn.Caption

290     data1.RecordSource = x

300     data1.Refresh


data1.Refresh



'310     If anapar Then
'320         MSFlexGrid1.ColS = 9
'330         MSFlexGrid1.ColWidth(0) = 1000
'340         MSFlexGrid1.ColWidth(2) = 3000
'
'350         MSFlexGrid1.ColAlignment(4) = 7
'            '   MSFlexGrid1.col = 5
'360         MSFlexGrid1.ColWidth(5) = 3000
'        Else
'370         MSFlexGrid1.ColWidth(0) = 1000
'380         MSFlexGrid1.ColS = 9
'390         MSFlexGrid1.ColWidth(1) = 3000
'400         MSFlexGrid1.ColAlignment(3) = 7
'410         MSFlexGrid1.Col = 4
'420         MSFlexGrid1.ColWidth(4) = 3000
'        End If

       

430   '  MSFlexGrid1.Refresh

         
         
         
'‚·ÊÂÈ ÙÈÛ ÛÔıÏÂÚ Í·Ù˘
         data1.Recordset.MoveFirst
          TDBGrid.Splits(0).columns(1).Width = 3500
         
         For k = 0 To data1.Recordset.FIELDS.Count - 1
            sumes(k) = 0
         Next
         
         
         
            On Error GoTo ektos
         Do While Not data1.Recordset.EOF
             For k = 1 To data1.Recordset.FIELDS.Count - 1
                 If IsNumeric(data1.Recordset.FIELDS(k).Value) And data1.Recordset.FIELDS(k).Type <> 202 Then
                     sumes(k) = sumes(k) + Val(Replace(data1.Recordset.FIELDS(k).Value, ",", "."))
                    End If
                Next
                data1.Recordset.MoveNext
         Loop
         For k = 0 To data1.Recordset.FIELDS.Count - 1

             If sumes(k) > 0 Then
                 TDBGrid.Splits(0).columns(k).FooterText = Format(sumes(k), "######0.00")
                End If

         Next
            
            
          TDBGrid.AlternatingRowStyle = True

     TDBGrid.OddRowStyle.BackColor = &H8000000F   ' GRI   vbCyan
     TDBGrid.EvenRowStyle.BackColor = &HFFFFC0   'OYRANI     &H8000000F  ' GRI
            

 

 
        
        









440     If Check1 Then


            Dim gr As Integer, omadop As Integer

            gr = MsgBox("Õ· „ﬂÌÂÙ·È ÔÏ·‰ÔÔﬂÁÛÁ ÛÙÔ 1Ô Â‰ﬂÔ;", vbYesNo)

            If gr = vbYes Then
               omadop = 1
            Else
               omadop = 0
            End If
            Dim sum_s As String
'            If Len(sumstring.Tag) > 2 Then
'               sum_s = sumstring.Tag
'            Else
               sum_s = "00111111111111111111111111"
'            End If

          Dim gg As String
          gg = "SELECT E.AEG ,D.* FROM (" + x + ") D INNER JOIN EID E ON E.KOD=D.[ Ÿƒ… œ”] order by E.AEG"



            print3_xar gg, sum_s, "", omadop    'sugm_str

450         'print3_xar x, "1111111111", Label2.Caption, 0    'typos2
        End If

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.PELAT42.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PELAT42.Command1_Click " & "at line " & Erl

        Resume Next
        
        
        
ektos:
        Exit Sub

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>

100     Unload Me

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.PELAT42.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PELAT42.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub deltia_Click()

        '<EhHeader>
        On Error GoTo deltia_Click_Err

        '</EhHeader>
        Dim dummy, PolAJ As String, ePolAj As String, PolPos As String, ePolPos As String, AgoAJ As String, eAgoAJ As String, AgoPos As String, eAgoPos As String

        Dim ANS As Integer

100     ANS = MsgBox("Ã¸ÌÔ ·ıÙ‹ Ôı ‰ÂÌ ÛıÏˆ˘ÌÂﬂ Á ÙÈÏÔÎÔ„ÁËÂﬂÛ· ÔÛ¸ÙÁÙ· ÏÂ ÙÁÌ ·„ÔÒ·ÛËÂﬂÛ·;", vbYesNo)

110     dummy = Get_EIDPAR(PolAJ, ePolAj, PolPos, ePolPos, AgoAJ, eAgoAJ, AgoPos, eAgoPos)

        Dim x As String

120     x = "select KODE as [ ˘‰ÈÍ¸Ú],LEFT(EID.ONO,30) as [–ÂÒÈ„Ò·ˆﬁ],"

130     x = x + " SUM(XRE) AS [–œ”œ‘«‘≈” ¡√œ—ŸÕ] ,"
140     x = x + " SUM(case when left(ATIM,1) IN (" + AgoAJ + ") THEN POSO ELSE 0  END ) as [‘…ÃœÀ.–œ”œ‘«‘≈”],"
150     x = x + "  str(sum(TIMM*POSO),10,2) as [¡Óﬂ·],CH1 AS [‘≈À.–¡—¡”],HPAR AS [‘≈À.¡√œ—¡] "
160     x = x + " FROM EGGTIM  INNER JOIN EID ON EGGTIM.KODE=EID.KOD"
170     x = x + " where  PELKOD='" + PEL_KOD
180     x = x + "' AND HME >='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME <='" + Format(DTPicker2, "MM/DD/YYYY") + "'"
190     x = x + " AND  KODE IN (SELECT KODE FROM DOKPELAT42) GROUP BY EID.ONO,KODE,EID.CH1,EID.HPAR"

200     If ANS = vbYes Then
210         x = x + " having SUM(XRE) - SUM(case when left(ATIM,1) IN (" + AgoAJ + ") THEN POSO ELSE 0  END ) <>0 "
        End If

220     data1.ConnectionString = gConnect

230     data1.RecordSource = x
240     data1.Refresh

250     If Check1 Then
260         print3_xar x, "1111111111", Label2.Caption, 0    'typos2
        End If

        '<EhFooter>
        Exit Sub

deltia_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.PELAT42.deltia_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PELAT42.deltia_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Activate()




  TDBGrid.Splits(0).ExtendRightColumn = True
     TDBGrid.ColumnFooters = True

        '<EhHeader>
        On Error GoTo Form_Activate_Err

        '</EhHeader>

        Dim s(1 To 10) As Single

        Dim R          As Recordset

        Dim k          As Integer

100     Me.Picture = LoadPicture(gPicture)

        'data1.DatabaseName = gDir

110     data1.ConnectionString = gConnect
120     Command1_Click
       ' Command1_Click
        
        
        
        Exit Sub

        'Data1.RecordSource = "select kode,onoma,sum(poso) as sposo,sum(timm*poso) as sajia,(sajia/sposo) as mtim from eggtim where EIDOS='" + pelat4.Data1.Recordset("EIDOS") + "' AND pelkod='" + pelat4.Data1.Recordset("kod") + "' group by kode,ONOMA"
        Dim x

130     x = "select KODE as [ ˘‰ÈÍ¸Ú],ONOMA as [–ÂÒÈ„Ò·ˆﬁ],STR(SUM(POSO),10,2) as [–ÔÛ¸ÙÁÙ·],str(sum(TIMM*POSO),10,2) as [¡Óﬂ·],LEFT(ATIM,1) FROM EGGTIM where HME >='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME <='" + Format(DTPicker2, "MM/DD/YYYY") + "' AND EIDOS='" + PEL_EIDOS + "' AND PELKOD='" + PEL_KOD + "'  group by LEFT(ATIM,1),KODE,ONOMA;"
140     data1.RecordSource = x

150     data1.Refresh

160     MSFlexGrid1.ColS = 5
170     MSFlexGrid1.ColWidth(1) = 3000

180     MSFlexGrid1.ColAlignment(3) = 7

190     MSFlexGrid1.Col = 4
200     MSFlexGrid1.ColWidth(4) = 3000

210     For k = 1 To MSFlexGrid1.rows - 1
220         MSFlexGrid1.row = k
230         MSFlexGrid1.Text = FindPar(MSFlexGrid1.Text)
        Next

240     MSFlexGrid1.Refresh

250     Label2.Caption = PEL_EPO.Caption + " ¡–œ " + Format(DTPicker1, "DD/MM/YYYY") + "  ≈Ÿ” " + Format(DTPicker2, "DD/MM/YYYY")

        Exit Sub

        'MSFlexGrid1.Columns(1).Caption = "œÌÔÏ·"
        '

        'MSFlexGrid1.Columns(2).Caption = "–ÔÛ¸ÙÁÙ·"

        'MSFlexGrid1.Columns(3).Caption = "ÃÂÛ.‘ÈÏ."
        'MSFlexGrid1.Columns(3).DataField = "mtim"
        'MSFlexGrid1.Columns(3).NumberFormat = "###,###.00"

        'MSFlexGrid1.Columns(4).Caption = "¡Óﬂ·"
        'MSFlexGrid1.Columns(4).NumberFormat = "###,###.00"

260     s(2) = 0
270     s(3) = 0

280     Set R = data1.Recordset

290     data1.Refresh

300     data1.Recordset.MoveFirst

310     Do While Not data1.Recordset.EOF
320         s(2) = s(2) + data1.Recordset(2)
330         s(3) = s(3) + data1.Recordset(3)
340         data1.Recordset.MoveNext
        Loop

350     data1.Refresh

        '<EhFooter>
        Exit Sub

Form_Activate_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.PELAT42.Form_Activate " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PELAT42.Form_Activate " & "at line " & Erl

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

130     data1.Recordset.Filter = getFilter()

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
         "in ADOMERCNEW.apot7.TDBGrid_FilterChange " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.TDBGrid_FilterChange " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub


Private Sub cmdClearFilter_Click()

        '<EhHeader>
        On Error GoTo cmdClearFilter_Click_Err

        '</EhHeader>

        'Clears filter from grid

100     For Each Col In TDBGrid.columns

110         Col.FILTERTEXT = ""

120     Next Col

130     data1.Recordset.Filter = adFilterNone

        '<EhFooter>
        Exit Sub

cmdClearFilter_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.cmdClearFilter_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.cmdClearFilter_Click " & "at line " & Erl

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

150             If data1.Recordset(Col.ColIndex).Type = 5 Then    '
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
         "in ADOMERCNEW.apot7.getFilter " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.getFilter " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function


Private Sub TDBGrid_HeadClick(ByVal ColIndex As Integer)

        '<EhHeader>
        On Error GoTo TDBGrid_HeadClick_Err
Dim k As Integer

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
        '    For Each Col In TDBGrid.Columns
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

        'TDBGrid.HoldFields
        Dim sumes(100) As String

100     For k = 0 To data1.Recordset.FIELDS.Count - 1
110         sumes(k) = TDBGrid.Splits(0).columns(k).FooterText    '  = Format(SUMES(k), "######0.00")
        Next

120     If data1.Recordset.Sort = "[" & TDBGrid.columns(ColIndex).DataField & "] asc" Then   ' strSort
130         data1.Recordset.Sort = "[" & TDBGrid.columns(ColIndex).DataField & "] desc"    ' strSort
        Else
140         data1.Recordset.Sort = "[" & TDBGrid.columns(ColIndex).DataField & "] asc"    ' strSort
        End If

150     For k = 0 To data1.Recordset.FIELDS.Count - 1
160         TDBGrid.Splits(0).columns(k).FooterText = sumes(k)   '  = Format(SUMES(k), "######0.00")
        Next

        '    data1.Recordset.Sort = "[" & TDBGrid.Columns(ColIndex).DataField & "] desc" ' strSort

        '<EhFooter>
        Exit Sub

TDBGrid_HeadClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.TDBGrid_HeadClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.TDBGrid_HeadClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

