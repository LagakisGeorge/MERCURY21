VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form HMEROL2 
   Caption         =   "«ÏÂÒﬁÛÈÔ –Ò¸„Ò·ÏÏ·"
   ClientHeight    =   9105
   ClientLeft      =   1275
   ClientTop       =   1620
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   14055
   Begin VB.TextBox Text8 
      Height          =   420
      Left            =   10800
      TabIndex        =   4
      Top             =   1560
      Width           =   1980
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   7320
      TabIndex        =   2
      Top             =   7440
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   163119105
      CurrentDate     =   38814
   End
   Begin VB.CommandButton Command1 
      Caption         =   "≈ ‘’–Ÿ”«"
      Height          =   420
      Left            =   9120
      TabIndex        =   1
      Top             =   7440
      Width           =   1770
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   8400
      Width           =   7095
      _ExtentX        =   12515
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
      Caption         =   "Customer Contact"
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
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Bindings        =   "hmerol2.frx":0000
      Height          =   7320
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   12912
      _LayoutType     =   4
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "–≈À¡‘«”"
      Columns(0).DataField=   "PELATHS"
      Columns(0).DataWidth=   35
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "≈—√¡”…¡"
      Columns(1).DataField=   "ERGASIA"
      Columns(1).DataWidth=   80
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   129
      Columns(2)._MaxComboItems=   5
      Columns(2).ValueItems(0)._DefaultItem=   0
      Columns(2).ValueItems(0).Value=   "ƒ«Ã«‘—«”"
      Columns(2).ValueItems(0).Value.vt=   8
      Columns(2).ValueItems(0).DisplayValue=   "ƒ«Ã«‘—«”"
      Columns(2).ValueItems(0).DisplayValue.vt=   8
      Columns(2).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(1)._DefaultItem=   0
      Columns(2).ValueItems(1).Value=   " ¡œ’–"
      Columns(2).ValueItems(1).Value.vt=   8
      Columns(2).ValueItems(1).DisplayValue=   " ¡œ’–"
      Columns(2).ValueItems(1).DisplayValue.vt=   8
      Columns(2).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(2)._DefaultItem=   0
      Columns(2).ValueItems(2).Value=   "Ã…◊¡À«”"
      Columns(2).ValueItems(2).Value.vt=   8
      Columns(2).ValueItems(2).DisplayValue=   "Ã…◊¡À«”"
      Columns(2).ValueItems(2).DisplayValue.vt=   8
      Columns(2).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(3)._DefaultItem=   0
      Columns(2).ValueItems(3).Value=   "”¡ «”"
      Columns(2).ValueItems(3).Value.vt=   8
      Columns(2).ValueItems(3).DisplayValue=   "”¡ «”"
      Columns(2).ValueItems(3).DisplayValue.vt=   8
      Columns(2).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(4)._DefaultItem=   0
      Columns(2).ValueItems(4).Value=   "»¡Õœ”"
      Columns(2).ValueItems(4).Value.vt=   8
      Columns(2).ValueItems(4).DisplayValue=   "»¡Õœ”"
      Columns(2).ValueItems(4).DisplayValue.vt=   8
      Columns(2).ValueItems(4)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(5)._DefaultItem=   0
      Columns(2).ValueItems(5).Value=   "À¡√¡ «”"
      Columns(2).ValueItems(5).Value.vt=   8
      Columns(2).ValueItems(5).DisplayValue=   "À¡√¡ «”"
      Columns(2).ValueItems(5).DisplayValue.vt=   8
      Columns(2).ValueItems(5)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(6)._DefaultItem=   0
      Columns(2).ValueItems(6).Value=   "–¡Õ¡√…Ÿ‘«”"
      Columns(2).ValueItems(6).Value.vt=   8
      Columns(2).ValueItems(6).DisplayValue=   "–¡Õ¡√…Ÿ‘«”"
      Columns(2).ValueItems(6).DisplayValue.vt=   8
      Columns(2).ValueItems(6)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(7)._DefaultItem=   0
      Columns(2).ValueItems(7).Value=   "–¡Õ‘≈À«”"
      Columns(2).ValueItems(7).Value.vt=   8
      Columns(2).ValueItems(7).DisplayValue=   "–¡Õ‘≈À«”"
      Columns(2).ValueItems(7).DisplayValue.vt=   8
      Columns(2).ValueItems(7)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(8)._DefaultItem=   0
      Columns(2).ValueItems(8).Value=   " œ”Ã…ƒ«”"
      Columns(2).ValueItems(8).Value.vt=   8
      Columns(2).ValueItems(8).DisplayValue=   " œ”Ã…ƒ«”"
      Columns(2).ValueItems(8).DisplayValue.vt=   8
      Columns(2).ValueItems(8)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems.Count=   9
      Columns(2).Caption=   "TEXNITHS"
      Columns(2).DataField=   "TEXNITHS"
      Columns(2).DataWidth=   14
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "ORES"
      Columns(3).DataField=   "ORES"
      Columns(3).DataWidth=   23
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   1
      Columns(4)._MaxComboItems=   5
      Columns(4).ValueItems(0)._DefaultItem=   0
      Columns(4).ValueItems(0).Value=   " À«”« ¡À‘≈ "
      Columns(4).ValueItems(0).Value.vt=   8
      Columns(4).ValueItems(0).DisplayValue=   " À«”« ¡À‘≈ "
      Columns(4).ValueItems(0).DisplayValue.vt=   8
      Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems(1)._DefaultItem=   0
      Columns(4).ValueItems(1).Value=   " À«”« ¡ÀÀ«"
      Columns(4).ValueItems(1).Value.vt=   8
      Columns(4).ValueItems(1).DisplayValue=   " À«”« ¡ÀÀ«"
      Columns(4).ValueItems(1).DisplayValue.vt=   8
      Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems(2)._DefaultItem=   0
      Columns(4).ValueItems(2).Value=   "≈Õ‘œ” ≈√√’«”«”"
      Columns(4).ValueItems(2).Value.vt=   8
      Columns(4).ValueItems(2).DisplayValue=   "≈Õ‘œ” ≈√√’«”«”"
      Columns(4).ValueItems(2).DisplayValue.vt=   8
      Columns(4).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems(3)._DefaultItem=   0
      Columns(4).ValueItems(3).Value=   "≈√ ¡‘¡”‘¡”« Àœ√…”Ã… œ’"
      Columns(4).ValueItems(3).Value.vt=   8
      Columns(4).ValueItems(3).DisplayValue=   "≈√ ¡‘¡”‘¡”« Àœ√…”Ã… œ’"
      Columns(4).ValueItems(3).DisplayValue.vt=   8
      Columns(4).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems.Count=   4
      Columns(4).Caption=   "KODLOG"
      Columns(4).DataField=   "KODLOG"
      Columns(4).DataWidth=   3
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   4
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "DONE"
      Columns(5).DataField=   "DONE"
      Columns(5).DataWidth=   23
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "POS"
      Columns(6).DataField=   "POS"
      Columns(6).DataWidth=   23
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "HME"
      Columns(7).DataField=   "HME"
      Columns(7).DataWidth=   19
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).AllowColMove=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=5847"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5768"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3784"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3704"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2355"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2275"
      Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Button=1"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=1111"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1032"
      Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(3)._AlignLeft=0"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=3069"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2990"
      Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(27)=   "Column(4).Button=1"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=1217"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=1138"
      Splits(0)._ColumnProps(32)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(34)=   "Column(5)._AlignLeft=0"
      Splits(0)._ColumnProps(35)=   "Column(6).Width=1429"
      Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=1349"
      Splits(0)._ColumnProps(38)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(40)=   "Column(6)._AlignLeft=0"
      Splits(0)._ColumnProps(41)=   "Column(7).Width=2593"
      Splits(0)._ColumnProps(42)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(7)._WidthInPix=2514"
      Splits(0)._ColumnProps(44)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(45)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(46)=   "Column(7)._AlignLeft=0"
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
      AnimateWindow   =   1
      AnimateWindowClose=   1
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bgcolor=&HFF8000&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=31,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35,.bgcolor=&HFFFFFF&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36,.bgcolor=&HC0C0C0&"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=93,.parent=2,.namedParent=95"
      _StyleDefs(23)  =   "FilterBarStyle:id=96,.parent=1,.namedParent=98"
      _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1,.bgcolor=&HFF8000&,.bold=0,.fontsize=825"
      _StyleDefs(25)  =   ":id=11,.italic=0,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(26)  =   ":id=11,.fontname=MS Sans Serif"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=94,.parent=93"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=97,.parent=96"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=44,.parent=11"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=41,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=42,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=43,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=48,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=45,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=47,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=40,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=37,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=38,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=39,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=28,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=60,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=57,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=58,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=59,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=52,.parent=11"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=49,.parent=12"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=50,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=51,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=56,.parent=11"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=53,.parent=12"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=54,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=55,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=24,.parent=11"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=21,.parent=12"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=22,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=23,.parent=15"
      _StyleDefs(70)  =   "Named:id=29:Normal"
      _StyleDefs(71)  =   ":id=29,.parent=0"
      _StyleDefs(72)  =   "Named:id=30:Heading"
      _StyleDefs(73)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(74)  =   ":id=30,.wraptext=-1"
      _StyleDefs(75)  =   "Named:id=31:Footing"
      _StyleDefs(76)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(77)  =   "Named:id=32:Selected"
      _StyleDefs(78)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(79)  =   "Named:id=33:Caption"
      _StyleDefs(80)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(81)  =   "Named:id=34:HighlightRow"
      _StyleDefs(82)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(83)  =   "Named:id=35:EvenRow"
      _StyleDefs(84)  =   ":id=35,.parent=29,.bgcolor=&H808080&,.fgcolor=&H0&,.bold=0,.fontsize=900"
      _StyleDefs(85)  =   ":id=35,.italic=0,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(86)  =   ":id=35,.fontname=Times New Roman"
      _StyleDefs(87)  =   "Named:id=36:OddRow"
      _StyleDefs(88)  =   ":id=36,.parent=29,.bgcolor=&H4080FF&"
      _StyleDefs(89)  =   "Named:id=95:RecordSelector"
      _StyleDefs(90)  =   ":id=95,.parent=30"
      _StyleDefs(91)  =   "Named:id=98:FilterBar"
      _StyleDefs(92)  =   ":id=98,.parent=29"
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
      Bindings        =   "hmerol2.frx":0015
      Height          =   1545
      Left            =   10800
      TabIndex        =   3
      Top             =   2040
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   2725
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
      Splits(0).RecordSelectorWidth=   794
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
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
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   345
      Left            =   10800
      Top             =   3840
      Visible         =   0   'False
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   609
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
End
Attribute VB_Name = "HMEROL2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RSClone As ADODB.Recordset

Dim Col     As TrueOleDBGrid80.Column

Dim ColS    As TrueOleDBGrid80.columns

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>
        Dim a, b, C

100     a = TDBGrid1.BackColor

110     b = TDBGrid1.OddRowStyle.BackColor
120     C = TDBGrid1.EvenRowStyle.BackColor

        'TDBGrid1.Columns(2).ValueItems(3).Value = "AAAA"

        'TDBGrid1.Columns(2).Value
        'TDBGrid1.Columns(2).CellValue(1) = "11111"
        'TDBGrid1.Columns(2).CellValue(2) = "22222"
        'TDBGrid1.Columns(2).CellValue(3) = "33333"

130     Adodc1.ConnectionString = gConnect
140     Adodc1.RecordSource = "SELECT PELATHS,ERGASIA,TEXNITHS,ORES,DONE FROM HMEROL2 WHERE  PELATHS IS NOT NULL  AND HME='" + Format(Now, "MM/DD/YYYY") + "'  ORDER BY TEXNITHS,PELATHS"

        'On Error GoTo latos
150     Adodc1.Refresh

160     TDBGrid1.clearFields

170     TDBGrid1.Rebind
180     TDBGrid1.columns(1).Width = 4000
190     TDBGrid1.columns(3).Width = 1000
200     TDBGrid1.columns(4).Width = 1000

210     With TDBGrid1.PrintInfo

220         TDBGrid1.OddRowStyle.BackColor = vbWhite
230         TDBGrid1.EvenRowStyle.BackColor = vbWhite

            ' Set the page header
240         .PageHeaderFont.Italic = True
250         .PageHeader = "Composers table"

            'Column headers will be on every page
260         .RepeatColumnHeaders = True

            ' Display page numbers (centered)
270         .PageFooter = "\tÛÂÎﬂ‰·: \p"
            ' Invoke Print Preview
280         .PrintPreview
290         TDBGrid1.BackColor = a
        End With

300     TDBGrid1.OddRowStyle.BackColor = b
310     TDBGrid1.EvenRowStyle.BackColor = C

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.HMEROL2.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.HMEROL2.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DTPicker1_Change()

        '<EhHeader>
        On Error GoTo DTPicker1_Change_Err

        '</EhHeader>

100     Adodc1.RecordSource = "SELECT * FROM HMEROL2 WHERE  HME='" + Format(dtpicker1, "MM/DD/YYYY") + "'"
110     Adodc1.Refresh

        '<EhFooter>
        Exit Sub

DTPicker1_Change_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.HMEROL2.DTPicker1_Change " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.HMEROL2.DTPicker1_Change " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
        Dim K

100     dtpicker1.Value = Now

        Dim R As New ADODB.Recordset

        Dim N

110     R.Open "SELECT *FROM PINAKES WHERE TYPOS=14 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
120     N = 0

130     Do While Not R.EOF
140         TDBGrid1.columns(2).ValueItems(N).Value = R("PERIGRAFH")
150         R.MoveNext
160         N = N + 1
        Loop

170     R.Close

180     Adodc1.ConnectionString = gConnect
190     Adodc1.RecordSource = "SELECT * FROM HMEROL2 WHERE  HME='" + Format(Now, "MM/DD/YYYY") + "'"

        On Error GoTo LATOS

200     Adodc1.Refresh

        On Error GoTo Form_Load_Err

        ' Adodc2.Refresh
        ' Adodc3.Refresh

210     R.Open "SELECT COUNT(*) FROM HMEROL2 WHERE HME='" + Format(Now, "MM/DD/YYYY") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

220     If R(0) < 20 Then
230         Gdb.Execute "DELETE FROM HMEROL2 WHERE ERGASIA IS NULL AND PELATHS IS NULL"

240         For K = 1 To 20
250             Gdb.Execute "INSERT INTO HMEROL2 (HME,TEXNITHS,POS,KODLOG) VALUES ('" + Format(Now, "MM/DD/YYYY") + "',''," + str(K) + ",'" + Format(Now, "DD/MM") + "')"
            Next

260         Gdb.Execute "UPDATE HMEROL2 SET HME='" + Format(Now, "MM/DD/YYYY") + "' WHERE DONE IS NULL "

270         Adodc1.RecordSource = "SELECT * FROM HMEROL2 WHERE  HME='" + Format(Now, "MM/DD/YYYY") + "'"
280         Adodc1.Refresh
        End If

290     TDBGrid1.columns(0).Width = 2000
300     TDBGrid1.columns(1).Width = 4000

310     TDBGrid1.columns(4).Width = 1000

        Exit Sub

LATOS:
320     Gdb.Execute "SELECT TOP 1 HPAR AS HME,XTI AS ORES,KOD AS TEXNITHS,ONO AS ERGASIA,ONO2 AS PELATHS,LTI AS DONE,POS,KODLOG INTO HMEROL2 FROM EID"

        ' Set RSClone = Adodc1.Recordset.Clone()
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.HMEROL2.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.HMEROL2.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub LTDBGrid1_Error(ByVal DataError As Integer, Response As Integer)
'    Dim a
'    a = DataError
'    Response = 0
'
'    Dim C
'    C = TDBGrid1.Col
'    a = Adodc1.Recordset.Bookmark
'    Adodc1.Refresh
'    Adodc1.Recordset.Bookmark = a
'    TDBGrid1.Col = C
'    'Resume Next
'
'End Sub

Private Sub TDBGrid1_FilterChange()

        '<EhHeader>
        On Error GoTo TDBGrid1_FilterChange_Err

        '</EhHeader>

        'Gets called when an action is performed on the filter bar

        On Error GoTo ErrHandler

100     Set ColS = TDBGrid1.columns

        Dim C As Integer

110     C = TDBGrid1.Col

120     TDBGrid1.HoldFields

130     Adodc1.Recordset.Filter = getFilter()

140     TDBGrid1.Col = C

150     TDBGrid1.EditActive = True

        Exit Sub

ErrHandler:

160     MsgBox Err.Source & ":" & vbCrLf & Err.Description

        'Call cmdClearFilter_Click

        '<EhFooter>
        Exit Sub

TDBGrid1_FilterChange_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.HMEROL2.TDBGrid1_FilterChange " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.HMEROL2.TDBGrid1_FilterChange " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub TDBGrid1_GroupColMove(ByVal Position As Integer, _
                                  ByVal ColIndex As Integer, _
                                  Cancel As Integer)

        '<EhHeader>
        On Error GoTo TDBGrid1_GroupColMove_Err

        '</EhHeader>
        Dim strSort As String

        Dim Col     As TrueOleDBGrid80.Column

        ' Loop through GroupColumns collection and construct

        ' the sort string for the Sort property of the Recordset

100     For Each Col In TDBGrid1.GroupColumns

110         If strSort <> vbNullString Then

120             strSort = strSort & ", "

            End If

130         strSort = strSort & "[" & Col.DataField & "]"

140     Next Col

150     TDBGrid1.HoldFields

160     Adodc1.Recordset.sort = strSort

        '<EhFooter>
        Exit Sub

TDBGrid1_GroupColMove_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.HMEROL2.TDBGrid1_GroupColMove " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.HMEROL2.TDBGrid1_GroupColMove " & "at line " & Erl

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

110         If Trim(Col.FilterText) <> "" Then
120             N = N + 1

130             If N > 1 Then
140                 tmp = tmp & " AND "
                End If

150             If Adodc1.Recordset(Col.ColIndex).Type = 5 Then    '
160                 If InStr(">< >= <=  = ", Left(Col.FilterText, 1)) > 0 And Len(Col.FilterText) > 1 And IsNumeric(Right(Col.FilterText, 1)) Then
170                     tmp = tmp & Col.DataField & Col.FilterText
                    Else
180                     tmp = tmp & Col.DataField & " >  " & Val(Col.FilterText)
                    End If

                Else
190                 tmp = tmp & Col.DataField & " LIKE '" & Col.FilterText & "*'"
                End If

            End If

200     Next Col

210     getFilter = tmp

        '<EhFooter>
        Exit Function

getFilter_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.HMEROL2.getFilter " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.HMEROL2.getFilter " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub TDBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo TDBGrid1_KeyUp_Err

        '</EhHeader>
        Dim a

100     If KeyCode = 13 Then

110         a = 1

        End If

        '<EhFooter>
        Exit Sub

TDBGrid1_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.HMEROL2.TDBGrid1_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.HMEROL2.TDBGrid1_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text8_LostFocus()

        '<EhHeader>
        On Error GoTo Text8_LostFocus_Err

        '</EhHeader>

100     Adodc2.RecordSource = "SELECT EPO,DIE,* FROM PEL WHERE EPO LIKE '" + (SameLetters(Text8.Text)) + "%'"
110     Adodc2.Refresh
120     TDBGrid1.columns(0).Width = 3000
130     TDBGrid1.columns(1).Width = 3000
140     TDBGrid1.Splits(0).AlternatingRowStyle = True
150     TDBGrid1.Splits(0).AnchorRightColumn = True

160     If Adodc1.Recordset.EOF Then

            Exit Sub

        End If

        Dim EPO, DIE, EPA, THL

170     EPO = Adodc2.Recordset("EPO")
180     DIE = CNull(Adodc2.Recordset("DIE"))
190     EPA = CNull(Adodc2.Recordset("EPA"))
200     THL = CNull(Adodc2.Recordset("THL"))

        ' ono.Text = ono.List(0)

        '<EhFooter>
        Exit Sub

Text8_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.HMEROL2.Text8_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.HMEROL2.Text8_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
