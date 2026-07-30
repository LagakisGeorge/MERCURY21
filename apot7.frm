VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form apot7 
   BackColor       =   &H00FF0000&
   Caption         =   "—ıËÏÈÊ¸ÏÂÌÂÚ ≈ÍÙı˛ÛÂÈÚ"
   ClientHeight    =   9900
   ClientLeft      =   225
   ClientTop       =   765
   ClientWidth     =   19110
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9900
   ScaleWidth      =   19110
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd≈…ƒ«¡Õ¡ 
      Caption         =   "≈…ƒ« ¡Õ¡ –≈À¡‘«"
      Height          =   360
      Left            =   13800
      TabIndex        =   42
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdREPORT 
      Caption         =   "REPORT"
      Height          =   240
      Left            =   4800
      TabIndex        =   41
      Top             =   1200
      Width           =   990
   End
   Begin VB.CommandButton Unilever 
      Caption         =   "Unilever"
      Height          =   360
      Left            =   13680
      TabIndex        =   40
      Top             =   1560
      Width           =   2052
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   7440
      TabIndex        =   39
      Top             =   9480
      Width           =   735
   End
   Begin VB.CommandButton gfk 
      Caption         =   "GFK ¡–œ”‘œÀ«"
      Height          =   375
      Left            =   11760
      TabIndex        =   38
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton TSENTIDIS 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5880
      TabIndex        =   37
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CMDpREVIEW 
      Caption         =   "–ÒÔÂÈÛÍ¸ÁÛÁ"
      Height          =   360
      Left            =   8040
      TabIndex        =   36
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox TITLOS 
      Height          =   375
      Left            =   1200
      TabIndex        =   31
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton CALC_YPOLOIPA_APOTHIKIS 
      Caption         =   "’ÔÎÔ„ÈÛÏ¸Ú ıÔÎÔÈ˘Ì"
      Height          =   330
      Left            =   3840
      TabIndex        =   30
      Top             =   3150
      Width           =   2100
   End
   Begin VB.CommandButton CMDRETAIL 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CSV RETAIL@LINK"
      Height          =   360
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3588
      Width           =   1728
   End
   Begin VB.TextBox sumstring 
      Height          =   285
      Left            =   8040
      TabIndex        =   28
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CheckBox Check2 
      Caption         =   "ƒÈ¸ÒË˘ÛÁ"
      Height          =   255
      Left            =   5964
      TabIndex        =   27
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H80000010&
      Caption         =   "≈ÓÔ‰ÔÚ"
      Height          =   375
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4110
      Width           =   1815
   End
   Begin VB.CheckBox epan 
      Caption         =   "≈·ÌÂÌÁÏ›Ò˘ÛÁ"
      Height          =   255
      Left            =   13680
      TabIndex        =   25
      Top             =   720
      Visible         =   0   'False
      Width           =   2676
   End
   Begin VB.CommandButton Command6 
      Caption         =   "OLD NOT USED √≈÷’—¡ ”≈ EPSILON"
      Height          =   375
      Left            =   14520
      TabIndex        =   24
      Top             =   9480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   450
      Left            =   16425
      ScaleHeight     =   19.5
      ScaleMode       =   2  'Point
      ScaleWidth      =   39.75
      TabIndex        =   23
      Top             =   7335
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Height          =   3855
      Left            =   13800
      ScaleHeight     =   3795
      ScaleWidth      =   5055
      TabIndex        =   22
      Top             =   4680
      Visible         =   0   'False
      Width           =   5115
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ÃÂÙ·ˆÔÒ‹ ˆ˘ÙÔ„Ò·ˆÈ˛Ì ÛÂ c:\par"
      Height          =   375
      Left            =   13680
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.CommandButton Command4 
      Caption         =   "≈ÍÙ˝˘ÛÁ ‘ﬂÙÎÔı"
      Height          =   375
      Left            =   13680
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "≈ÍÙ˝˘ÛÁ ¬È‚ÎﬂÔı"
      Height          =   375
      Left            =   3945
      TabIndex        =   19
      Top             =   4110
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox Check1EKTYP 
      Caption         =   "≈ÍÙ˝˘ÛÁ"
      Height          =   210
      Left            =   4320
      TabIndex        =   18
      Top             =   2040
      Width           =   1590
   End
   Begin VB.TextBox c3 
      Height          =   285
      Left            =   1800
      TabIndex        =   16
      ToolTipText     =   "@c3"
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox c2 
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      ToolTipText     =   "@c2"
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox c1 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      ToolTipText     =   "@c1 .˜. CH2 LIKE    '@c1%'  ﬁ  ¡=@C1"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   3705
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9360
      Visible         =   0   'False
      Width           =   2220
   End
   Begin RichTextLib.RichTextBox TEXT2 
      Height          =   3480
      Left            =   5964
      TabIndex        =   6
      Top             =   12
      Width           =   7608
      _ExtentX        =   13414
      _ExtentY        =   6138
      _Version        =   393217
      Enabled         =   0   'False
      TextRTF         =   $"apot7.frx":0000
   End
   Begin VB.CheckBox ODBC 
      Caption         =   "ODBC"
      Height          =   240
      Left            =   3840
      TabIndex        =   5
      Top             =   1200
      Width           =   840
   End
   Begin VB.CommandButton CMDAPOTHIKEYSI 
      Caption         =   "¡ÔËﬁÍÂıÛÁ"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   1560
      Width           =   2085
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   3735
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Bindings        =   "apot7.frx":0082
      Height          =   4650
      Left            =   0
      TabIndex        =   2
      Top             =   4680
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   8202
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
      Columns(5)._VlistStyle=   0
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
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
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
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(29)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(34)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(39)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(44)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(49)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(51)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(54)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   2
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
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
      _StyleDefs(80)  =   "Named:id=33:Normal"
      _StyleDefs(81)  =   ":id=33,.parent=0"
      _StyleDefs(82)  =   "Named:id=34:Heading"
      _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(84)  =   ":id=34,.wraptext=-1"
      _StyleDefs(85)  =   "Named:id=35:Footing"
      _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(87)  =   ":id=35,.appearance=3,.borderColor=&H80FF&"
      _StyleDefs(88)  =   "Named:id=36:Selected"
      _StyleDefs(89)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(90)  =   "Named:id=37:Caption"
      _StyleDefs(91)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(92)  =   "Named:id=38:HighlightRow"
      _StyleDefs(93)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(94)  =   "Named:id=39:EvenRow"
      _StyleDefs(95)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(96)  =   "Named:id=40:OddRow"
      _StyleDefs(97)  =   ":id=40,.parent=33"
      _StyleDefs(98)  =   "Named:id=41:RecordSelector"
      _StyleDefs(99)  =   ":id=41,.parent=34"
      _StyleDefs(100) =   "Named:id=42:FilterBar"
      _StyleDefs(101) =   ":id=42,.parent=33"
      _StyleDefs(102) =   "bmp(0):id=1,KAAAAEAAAABAAAAAAQAYAAAAAAAAMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(103) =   "bmp(1):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(104) =   "bmp(2):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(105) =   "bmp(3):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwAAAAICAgAAAAICAgAAAAICAgMDAwICAgAAA"
      _StyleDefs(106) =   "bmp(4):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(107) =   "bmp(5):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(108) =   "bmp(6):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwAAA"
      _StyleDefs(109) =   "bmp(7):id=1,AAAAAAAAAAAAAAAAAICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(110) =   "bmp(8):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///4CAgAAA"
      _StyleDefs(111) =   "bmp(9):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(112) =   "bmp(10):id=1,AAAAAAAAAAAAAAAAAICAgICAgMDAwAAAAAAAAAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(113) =   "bmp(11):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(114) =   "bmp(12):id=1,AAAAAICAgMDAwICAgMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(115) =   "bmp(13):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgICAgAAAAAAA"
      _StyleDefs(116) =   "bmp(14):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(117) =   "bmp(15):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///wAAAAAAAICAgAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(118) =   "bmp(16):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgAAAAAAAAAAA"
      _StyleDefs(119) =   "bmp(17):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(120) =   "bmp(18):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICA"
      _StyleDefs(121) =   "bmp(19):id=1,gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(122) =   "bmp(20):id=1,AAAAAAAAAICAgMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(123) =   "bmp(21):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(124) =   "bmp(22):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(125) =   "bmp(23):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(126) =   "bmp(24):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(127) =   "bmp(25):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(128) =   "bmp(26):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(129) =   "bmp(27):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(130) =   "bmp(28):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(131) =   "bmp(29):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwMDAwICAgAAA"
      _StyleDefs(132) =   "bmp(30):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(133) =   "bmp(31):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(134) =   "bmp(32):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(135) =   "bmp(33):id=1,AAAAAAAAAICAgP///8DAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(136) =   "bmp(34):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(137) =   "bmp(35):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(138) =   "bmp(36):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(139) =   "bmp(37):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(140) =   "bmp(38):id=1,AAAAAICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(141) =   "bmp(39):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(142) =   "bmp(40):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(143) =   "bmp(41):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///4CAgAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(144) =   "bmp(42):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(145) =   "bmp(43):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(146) =   "bmp(44):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///8DAwP//"
      _StyleDefs(147) =   "bmp(45):id=1,/4CAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(148) =   "bmp(46):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(149) =   "bmp(47):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(150) =   "bmp(48):id=1,AAAAAICAgMDAwMDAwMDAwMDAwMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(151) =   "bmp(49):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///4CAgAAAAAAA"
      _StyleDefs(152) =   "bmp(50):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(153) =   "bmp(51):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///4CAgAAAAAAAAAAAAICAgAAAAAAAAAAAAAAA"
      _StyleDefs(154) =   "bmp(52):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(155) =   "bmp(53):id=1,AICAgMDAwICAgMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(156) =   "bmp(54):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwAAA"
      _StyleDefs(157) =   "bmp(55):id=1,AAAAAICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(158) =   "bmp(56):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///wAAAAAAAICAgAAAAICAgICAgICAgAAAAAAAAAAA"
      _StyleDefs(159) =   "bmp(57):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(160) =   "bmp(58):id=1,AAAAAAAAAAAAAAAAAAAAAICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(161) =   "bmp(59):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgAAA"
      _StyleDefs(162) =   "bmp(60):id=1,AICAgP///8DAwP///4CAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(163) =   "bmp(61):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(164) =   "bmp(62):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(165) =   "bmp(63):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwICAgMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(166) =   "bmp(64):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(167) =   "bmp(65):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(168) =   "bmp(66):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///wAAAAAAAICA"
      _StyleDefs(169) =   "bmp(67):id=1,gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(170) =   "bmp(68):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(171) =   "bmp(69):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(172) =   "bmp(70):id=1,AAAAAAAAAAAAAICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(173) =   "bmp(71):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(174) =   "bmp(72):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(175) =   "bmp(73):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgICAgICA"
      _StyleDefs(176) =   "bmp(74):id=1,gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(177) =   "bmp(75):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(178) =   "bmp(76):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(179) =   "bmp(77):id=1,AICAgICAgP///8DAwP///8DAwP///4CAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(180) =   "bmp(78):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(181) =   "bmp(79):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(182) =   "bmp(80):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///8DAwP///8DAwP///8DAwP///8DAwICAgAAAAAAA"
      _StyleDefs(183) =   "bmp(81):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(184) =   "bmp(82):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(185) =   "bmp(83):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///8DAwP///8DAwP//"
      _StyleDefs(186) =   "bmp(84):id=1,/8DAwP///8DAwP///8DAwICAgAAAAAAAAAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(187) =   "bmp(85):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(188) =   "bmp(86):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(189) =   "bmp(87):id=1,AICAgP///8DAwP///8DAwP///8DAwP///8DAwP///8DAwP///4CAgAAAAAAAAMDAwMDAwICAgICA"
      _StyleDefs(190) =   "bmp(88):id=1,gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(191) =   "bmp(89):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(192) =   "bmp(90):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///8DAwP///8DAwP///8DAwP///8DAwP///8DA"
      _StyleDefs(193) =   "bmp(91):id=1,wP///4CAgAAAAMDAwP///8DAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(194) =   "bmp(92):id=1,AAAAAAAAAAAAAICAgICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(195) =   "bmp(93):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///8DAwP///8DA"
      _StyleDefs(196) =   "bmp(94):id=1,wP///8DAwP///8DAwP///8DAwP///8DAwICAgAAAAAAAAMDAwMDAwAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(197) =   "bmp(95):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgP///8DAwP///8DAwICAgICAgAAAAAAAAAAA"
      _StyleDefs(198) =   "bmp(96):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(199) =   "bmp(97):id=1,AAAAAICAgMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwICAgAAAAAAAAAAA"
      _StyleDefs(200) =   "bmp(98):id=1,AAAAAAAAAAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///8DAwP//"
      _StyleDefs(201) =   "bmp(99):id=1,/8DAwP///8DAwP///4CAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DA"
      _StyleDefs(202) =   "bmp(100):id=1,wP///4CAgAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///4CAgP///4CAgAAAAICAgAAAAICAgAAA"
      _StyleDefs(203) =   "bmp(101):id=1,AICAgAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAMDAwMDAwICAgICAgAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(204) =   "bmp(102):id=1,AAAAAAAAAICAgP///8DAwP///8DAwP///8DAwP///8DAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(205) =   "bmp(103):id=1,AAAAAAAAAAAAAICAgP///8DAwP///8DAwP///4CAgAAAAAAAAAAAAAAAAAAAAICAgP///4CAgP//"
      _StyleDefs(206) =   "bmp(104):id=1,/4CAgAAAAICAgAAAAICAgAAAAICAgAAAAICAgMDAwICAgAAAAAAAAAAAAAAAAAAAAMDAwP///8DA"
      _StyleDefs(207) =   "bmp(105):id=1,wICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///8DAwP///8DAwP///8DAwP///8DA"
      _StyleDefs(208) =   "bmp(106):id=1,wICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwMDAwMDAwMDAwMDAwICAgAAAAAAA"
      _StyleDefs(209) =   "bmp(107):id=1,AAAAAAAAAAAAAAAAAICAgP///4CAgAAAAICAgAAAAICAgAAAAAAAAAAAAAAAAAAAAICAgAAAAAAA"
      _StyleDefs(210) =   "bmp(108):id=1,AAAAAAAAAAAAAAAAAAAAAMDAwMDAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwMDAwMDA"
      _StyleDefs(211) =   "bmp(109):id=1,wMDAwMDAwMDAwMDAwMDAwMDAwMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP//"
      _StyleDefs(212) =   "bmp(110):id=1,/4CAgAAAAAAAAAAAAICAgAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(213) =   "bmp(111):id=1,AAAAAAAAAAAAAMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(214) =   "bmp(112):id=1,AAAAAAAAAAAAAICAgP///4CAgP///4CAgAAAAICAgAAAAICAgAAAAICAgICAgAAAAAAAAAAAAAAA"
      _StyleDefs(215) =   "bmp(113):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwAAAAAAAAICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(216) =   "bmp(114):id=1,AICAgMDAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(217) =   "bmp(115):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///4CAgAAAAICAgAAAAICAgAAA"
      _StyleDefs(218) =   "bmp(116):id=1,AICAgMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgAAAAAAA"
      _StyleDefs(219) =   "bmp(117):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgMDAwAAAAAAAAAAAAAAAAICAgICAgAAAAAAA"
      _StyleDefs(220) =   "bmp(118):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICA"
      _StyleDefs(221) =   "bmp(119):id=1,gMDAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(222) =   "bmp(120):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICA"
      _StyleDefs(223) =   "bmp(121):id=1,gICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(224) =   "bmp(122):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwAAAAAAAAAAAAAAAAAAAAAAAAICAgAAAAAAAAAAA"
      _StyleDefs(225) =   "bmp(123):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(226) =   "bmp(124):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(227) =   "bmp(125):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgMDAwAAAAAAA"
      _StyleDefs(228) =   "bmp(126):id=1,AAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(229) =   "bmp(127):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(230) =   "bmp(128):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(231) =   "bmp(129):id=1,AAAAAAAAAAAAAAAAAICAgICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(232) =   "bmp(130):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(233) =   "bmp(131):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(234) =   "bmp(132):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgICAgAAAAAAA"
      _StyleDefs(235) =   "bmp(133):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(236) =   "bmp(134):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(237) =   "bmp(135):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICA"
      _StyleDefs(238) =   "bmp(136):id=1,gMDAwP///8DAwP///8DAwICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(239) =   "bmp(137):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(240) =   "bmp(138):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(241) =   "bmp(139):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///8DAwP///8DAwICAgAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(242) =   "bmp(140):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(243) =   "bmp(141):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(244) =   "bmp(142):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///8DAwP///8DA"
      _StyleDefs(245) =   "bmp(143):id=1,wP///8DAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(246) =   "bmp(144):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(247) =   "bmp(145):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(248) =   "bmp(146):id=1,AICAgMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(249) =   "bmp(147):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(250) =   "bmp(148):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(251) =   "bmp(149):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///4CAgMDAwICAgAAAAICAgAAAAICAgICAgAAAAAAA"
      _StyleDefs(252) =   "bmp(150):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(253) =   "bmp(151):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(254) =   "bmp(152):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwAAAAICAgAAA"
      _StyleDefs(255) =   "bmp(153):id=1,AICAgAAAAICAgMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(256) =   "bmp(154):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(257) =   "bmp(155):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(258) =   "bmp(156):id=1,AAAAAAAAAAAAAICAgMDAwAAAAAAAAAAAAAAAAAAAAICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(259) =   "bmp(157):id=1,AAAAAAAAAAAAAICAgICAgICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(260) =   "bmp(158):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(261) =   "bmp(159):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgMDAwAAAAAAAAAAAAICAgICAgAAA"
      _StyleDefs(262) =   "bmp(160):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///8DAwP///8DAwP///8DAwICAgAAAAAAA"
      _StyleDefs(263) =   "bmp(161):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(264) =   "bmp(162):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(265) =   "bmp(163):id=1,AICAgICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///8DAwP//"
      _StyleDefs(266) =   "bmp(164):id=1,/8DAwP///8DAwP///8DAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(267) =   "bmp(165):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(268) =   "bmp(166):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(269) =   "bmp(167):id=1,AAAAAAAAAMDAwP///8DAwP///8DAwP///8DAwP///8DAwP///8DAwAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(270) =   "bmp(168):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(271) =   "bmp(169):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(272) =   "bmp(170):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///8DAwP///8DAwP///8DAwP///8DA"
      _StyleDefs(273) =   "bmp(171):id=1,wP///4CAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(274) =   "bmp(172):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(275) =   "bmp(173):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///8DA"
      _StyleDefs(276) =   "bmp(174):id=1,wP///8DAwP///8DAwP///8DAwP///8DAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(277) =   "bmp(175):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(278) =   "bmp(176):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(279) =   "bmp(177):id=1,AAAAAAAAAAAAAICAgMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwICAgAAAAAAAAAAA"
      _StyleDefs(280) =   "bmp(178):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(281) =   "bmp(179):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(282) =   "bmp(180):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwP///4CAgMDAwICAgAAAAICAgAAA"
      _StyleDefs(283) =   "bmp(181):id=1,AICAgAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(284) =   "bmp(182):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(285) =   "bmp(183):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDA"
      _StyleDefs(286) =   "bmp(184):id=1,wICAgAAAAICAgAAAAICAgAAAAICAgAAAAICAgMDAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(287) =   "bmp(185):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(288) =   "bmp(186):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(289) =   "bmp(187):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwAAAAAAAAICAgAAAAAAAAAAAAAAAAAAAAICAgAAAAAAA"
      _StyleDefs(290) =   "bmp(188):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(291) =   "bmp(189):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(292) =   "bmp(190):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgMDAwICAgAAAAAAAAAAA"
      _StyleDefs(293) =   "bmp(191):id=1,AAAAAAAAAAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(294) =   "bmp(192):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(295) =   "bmp(193):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(296) =   "bmp(194):id=1,AAAAAAAAAICAgMDAwMDAwAAAAAAAAAAAAAAAAICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(297) =   "bmp(195):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(298) =   "bmp(196):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(299) =   "bmp(197):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgICAgICAgICAgICAgAAAAAAAAAAA"
      _StyleDefs(300) =   "bmp(198):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICA"
      _StyleDefs(301) =   "bmp(199):id=1,gICAgICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(302) =   "bmp(200):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(303) =   "bmp(201):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(304) =   "bmp(202):id=1,AAAAAAAAAAAAAAAAAICAgMDAwP///8DAwP///8DAwICAgICAgAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(305) =   "bmp(203):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(306) =   "bmp(204):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(307) =   "bmp(205):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP///8DAwP///8DAwP///8DAwICA"
      _StyleDefs(308) =   "bmp(206):id=1,gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(309) =   "bmp(207):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(310) =   "bmp(208):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgP//"
      _StyleDefs(311) =   "bmp(209):id=1,/8DAwP///8DAwP///8DAwP///8DAwICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(312) =   "bmp(210):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(313) =   "bmp(211):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(314) =   "bmp(212):id=1,AAAAAAAAAAAAAAAAAAAAAICAgMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwICAgAAAAAAAAAAAAAAA"
      _StyleDefs(315) =   "bmp(213):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(316) =   "bmp(214):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      _StyleDefs(317) =   "bmp(215):id=1,AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgAAAAICAgMDAwICAgAAAAICA"
      _StyleDefs(318) =   "bmp(216):id=1,gAAAAICAgICAgAAAAAAAAA=="
   End
   Begin VB.CommandButton CMDEXECUTE 
      Caption         =   "≈ÍÙ›ÎÂÛÁ SQL"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   915
      Width           =   2085
   End
   Begin VB.CommandButton CMDPROVOLI 
      BackColor       =   &H00C0FFC0&
      Caption         =   "–ÒÔ‚ÔÎﬁ Â„„Ò·ˆ˛Ì (SQL)"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   390
      Width           =   2085
   End
   Begin MSComCtl2.DTPicker apo 
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      ToolTipText     =   "@x1  .˜.  HME>@x1"
      Top             =   2280
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   456982529
      CurrentDate     =   38814
   End
   Begin MSComCtl2.DTPicker eos 
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      ToolTipText     =   "@x2"
      Top             =   2670
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   456982529
      CurrentDate     =   38814
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   2280
      Top             =   9360
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
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label9 
      Caption         =   "”ÙﬁÎÂÚ ÏÂ Û˝ÌÔÎ·"
      Height          =   255
      Left            =   5964
      TabIndex        =   35
      Top             =   4230
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "‘…‘Àœ”"
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
      Left            =   0
      TabIndex        =   34
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "¡–œ"
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
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "¡–œ"
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
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "–·Ò‹ÏÂÙÒÔÚ 3"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4230
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "–·Ò‹ÏÂÙÒÔÚ 2"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "–·Ò‹ÏÂÙÒÔÚ 1"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "¡–œ"
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
      Left            =   3840
      TabIndex        =   10
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "≈Ÿ”"
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
      Left            =   3870
      TabIndex        =   9
      Top             =   2790
      Width           =   615
   End
End
Attribute VB_Name = "apot7"
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

Dim fSCR        As Object    ' CONTROLSCRIPT

Dim F928

Dim FDSN     As String

Dim F_EIKONA As Integer  ' = Val(FindParametroi(1,"APOT2", "F_EIKONA", "1", "1=EIKONA=> Ÿƒ… œ” 2=>Ã≈ ¬œ«». À≈…ƒ… 11= VDIKOS.JPG 12=BOH.JPG 0=œ◊… ≈… œÕ≈”"))

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Dim ColS      As TrueOleDBGrid80.columns

Dim f_pelMERC As String
Dim f_gfk As Long





'Dim m_plath(20) As Integer
'
'
' Property Get logot() As String
'    logot = m_fontname
'End Property
'
' Property Let logot(ByVal New_fontname As Integer)
'    m_fontname = New_fontname
'    PropertyChanged "fontname"
'End Property


'Private v() As Double
Dim m_plath(150) As Integer
Public Property Get Plath_Grid(ByVal index As Long) As Double
    Plath_Grid = m_plath(index)
End Property
Public Property Let Plath_Grid(ByVal index As Long, MyValue As Double)
    m_plath(index) = MyValue
End Property





Function PutDates(ByVal X As String)

        '<EhHeader>
        On Error GoTo PutDates_Err

        '</EhHeader>

        Dim X1 As String, X2 As String

100     X1 = "'" + Format(APO, "MM/DD/YYYY") + "'"
110     X2 = "'" + Format(DateAdd("H", 23.99, EOS), "MM/DD/YYYY") + "'"

        Dim SQL2 As String

120     SQL2 = X    ' TEXT2.Text

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
         "in ADOMERCNEW.apot7.cmdClearFilter_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.cmdClearFilter_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub CALC_YPOLOIPA_APOTHIKIS_Click()

Dim FF As New apot3param
 FF.Label_titlos = "’¸ÎÔÈ· ≈È‰˛Ì"

FF.load_sthles_EIDON
FF.APO = APO.Value
FF.EOS = EOS.Value

'FF.LOAD_KINHSEIS_EIDON
FF.Rep_Fields.Text = "select KOD as [ ˘‰ÈÍ¸Ú],LEFT(ONO,35) as [–ÂÒÈ„Ò·ˆﬁ],POS AS [’¸ÎÔÈÔ]"
FF.gemisma_Krithrion
'fkrithria.Grid1.TextMatrix(4, 1) = 0.1
'fkrithria.Grid1.TextMatrix(4, 2) = 999999999
'FF.KRITIRIA.≈◊…‘_Click
FF.sumes.Text = "001"
FF.WHERE_QUERY.Text = "POS>=0.1 AND POS<=999999999" ' StatusBar1.Panels(6).Text"

FF.ypologizo_Ypoloipa (False)

FF.CMDEXIT_Click











End Sub

Private Sub Check2_Click()

    If Check2.Value = vbChecked Then
        Text2.Enabled = True
        CMDAPOTHIKEYSI.Enabled = True
    Else
        Text2.Enabled = False
        CMDAPOTHIKEYSI.Enabled = False
     
    End If

End Sub

Private Sub CMDAPOTHIKEYSI_Click()

        '<EhHeader>
        On Error GoTo CMDAPOTHIKEYSI_Click_Err

        '</EhHeader>

100     cd1.InitDir = File1.Path

110     cd1.ShowSave

        Dim F

120     F = cd1.FILENAME

130     Open F For Output As #5
140     Print #5, Text2.Text
150     Close #5

        '+ ".txt"

        '<EhFooter>
        Exit Sub

CMDAPOTHIKEYSI_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.Command11_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.Command11_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdREPORT_Click()
    Shell "EXPLORER.EXE " + "C:\MERCVB\REPORTS\TAMEIO3.RPT", vbMaximizedFocus
End Sub

Private Sub cmd≈…ƒ«¡Õ¡_Click()
Dim txt(100) As String

'‘œ SQL ≈…Õ¡… ‘œ ¡ œÀœ’»œ
txt(1) = "/*#¡–œ  Ÿƒ… œ     */" + Chr(13)
txt(2) = "/*#≈Ÿ”  Ÿƒ… œ     */" + Chr(13)
Text2.Text = txt(1) + txt(2) + " select D.ONO,SUM(G.XRE) AS SYNEIS,SUM(G.PIS) AS SYNEX,EPO,AFM,KODE FROM EGGTIM G INNER JOIN PEL  P "
Text2.Text = Text2.Text + " ON G.EIDOS=P.EIDOS AND G.PELKOD=P.KOD  INNER JOIN EID D  ON G.KODE=D.KOD "
Text2.Text = Text2.Text + " WHERE HME>=@x1 and HME<=@x2 AND KODE >='@c1'  and KODE<='@c2' group by D.ONO,EPO,AFM,KODE  ORDER BY D.ONO"


        If InStr(txt(1), "#") > 0 And InStr(txt(1), "/*") > 0 Then
            Label3.Caption = mID(txt(1), InStr(txt(1), "#") + 1, 15)
        End If

        If InStr(txt(2), "#") > 0 And InStr(txt(2), "/*") > 0 Then
            Label4.Caption = mID(txt(2), InStr(txt(2), "#") + 1, 15)
        End If

'        If InStr(txt(3), "#") > 0 And InStr(txt(3), "/*") > 0 Then
'            Label5.Caption = mID(txt(3), InStr(txt(3), "#") + 1, 15)
'        End If
'









End Sub

Private Sub Command1_Click()
   CREATE_MONADES
   
End Sub

Public Sub CREATE_MONADES()
Dim a As String


a = " SET ANSI_NULLS ON"
a = a + ";SET QUOTED_IDENTIFIER ON"
a = a + ";"
a = a + "CREATE TABLE [dbo].[MONADES]([Code value] [nvarchar](255) NULL, [Name] [nvarchar](255) NULL,[Description] [nvarchar](max) NULL) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'1N', N'Count', N'Count')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'1S', N'Milligram per Kilogram (mg/kg)', N'Milligram per Kilogram (mg/kg)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'23', N'Grams Per Cubic Centimetre', N'Grams Per Cubic Centimetre')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'28', N'Kilogram per square metre', N'A unit of pressure equal to 9.80665*10-05 Bar')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'2L', N'Cubic Foot Per Minute', N'Cubic foot per minute')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'2M', N'Centimetre Per Second', N'The number of centimetre per second the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'2N', N'Decibel', N'A measurement for sound in air and other gases, relative to 20 micropascals (µPa) = 2◊10-5 Pa, the quietest sound a human can hear. This is roughly the sound of a mosquito flying 3 metres away. This is often abbreviated to just dBa however the correct abbreviation is dB(SPL), indicating decibel for Sound Pressure Level.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'2P', N'Kilobyte', N'A unit of information equal to 10≥ (1000) bytes.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'2Q', N'Kilo Becquerel', N'kBq is 10≥ Bq')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'2X', N'Metre Per Minute', N'Metre Per Minute')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'4G', N'Microlitre', N'A microlitre is one millionth of a litre')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'4H', N'Micrometre', N'A micrometre is one millionth of a metre, also termed Micron.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'4L', N'Megabyte', N'A unit of information equal to 10 (1000000) bytes.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'4N', N'Megabecquerel', N'Megabecquerel: 106 Bq1 Bq is defined as the activity of a quantity of radioactive material in which one nucleus decays per second.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'58', N'Net kilogram', N'A unit of mass defining the total number of kilograms after deductions.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'59', N'Part per million', N'A unit of proportion equal to 10-6 (ppm).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'5B', N'Batch', N'A unit of count defining the number of batches (batch: quantity of material produced in one operation or number of animals or persons coming at once).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'64', N'Pound per square inch - Gauge', N'Psig (pound-force per square inch gauge) is a unit of pressure relative to the surrounding atmosphere. At sea level, Earth''s atmosphere actually exerts a pressure of 14.7 psi. Humans do not feel this pressure because internal pressure of liquid in their bodies matches the external pressure. If a pressure gauge is calibrated to read zero in space, then at sea level on Earth it would read 14.7 psi. Thus a reading of 30 psig, on Earth, on a tire gauge represents an absolute pressure of 44.7 psi (lb/in≤).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'80', N'Pound per square inch - Absolute', N'Psia (pound-force per square inch absolute)is a unit of pressure pressure relative to a vacuum (such as that in space). At sea level, Earth''s atmosphere actually exerts a pressure of 14.7 psi. Humans do not feel this pressure because internal pressure of liquid in their bodies matches the external pressure. If a pressure gauge is calibrated to read zero in space, then at sea level on Earth it would read 14.7 psi. Thus a reading of 30 psig, on Earth, on a tire gauge represents an absolute pressure of 44.7 psi (lb/in≤).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'/L', N'per litre (/L)', N'Per Litre [/L] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'/mL', N'per millilitre (/mL)', N'Per millilitre [/mL] (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'/mmol', N'per millimole (/mmol)', N'Per millimole [/mmol] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'[APL''U]', N'IgA phospholipid units [APL''U]', N'IgA phospholipid units - biologic activity of anticardiolipin [APL''U] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'[beth''U]', N'Bethesda unit [beth''U]', N'Bethesda unit - biologic activity of factor VIII inhibitor [beth''U] (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'[GPL''U]', N'IgG phospholipid units [GPL''U]', N'IgG phospholipid units - biologic activity of anticardiolipin IgG [GPL''U] (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'[HPF]', N'per high power field [HPF]', N'Per high power field - view area in microscope [/[HPF]] (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'[lb_av]/[ft_i]3', N'Pound per cubic foot', N'Pound per cubic foot, [lb_av]/[ft_i]3 (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'[lb_av]/[in_i]3', N'Pound per cubic inch', N'Pound per cubic inch, [lb_av]/[in_i]3 (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'[iU]/d', N'International unit per day [iU/d]', N'International unit per day [iU/d] (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'[iU]/L', N'International unit per litre [iU/L]', N'International unit per litre [iU/L] (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'[iU]/mL', N'International unit per millilitre [iU/mL]', N'International unit per millilitre [iU/mL] (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'[LPF]', N'Per low power field [LPF]', N'Per low power field - view area in microscope [/[LPF]] (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'[MPL''U]/L', N'IgM phospholipid units [MPL''U] per litre', N'IgM phospholipid units - biologic activity of anticardiolipin IgM ñ MPL unit per litre [[MPL''U]/L] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'A11', N'Angstrom', N'A unit of length equal to one hundred-millionth of a centimetre, 10ñ10 metre, used mainly to express wavelengths and interatomic distances.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'A24', N'Candela per Square Metre', N'Candela per Square Metre')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'A43', N'Deadweight Tonnage', N'A unit of mass defining the difference between the weight of a ship when completely empty and its weight when completely loaded, expressed as the number of tons.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'A71', N'Femtometre', N'Femtometre- a quadrillionth of a metre (10 to the negative 15th power)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'A86', N'Gigahertz', N'A unit of frenquecy equal to 109 Hertz')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'AD', N'Byte', N'A unit of information equal to 8 bits.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'AMH', N'Ampere Hour', N'A unit of electric charge defining the amount of charge accumulated by a steady flow of one ampere for one hour.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'AMP', N'Ampere', N'Ampere')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ANN', N'Year', N'Unit of time equal to 365,25 days.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'APZ', N'Troy ounce or apothecary ounce', N'The troy ounce is a unit of imperial measure. In the present day it is most commonly used to gauge the weight and therefore the price of precious metals. One troy ounce equals 480 grains or 31.1035 grams.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'AS', N'Assortment', N'A unit of count defining the number of assortments (assortment: set of items grouped in a mixed collection).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ATM', N'Standard Atmosphere', N'Standard Atmosphere')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ATT', N'Technical Atmosphere', N'Technical Atmosphere')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'AWG', N'Gauge', N'Gauge- A number referring to the outer diameter of hypodermic or suture needles. Smaller gauge numbers indicate larger outer diameters. Inner diameter depends on both gauge and wall thickness. An increasing needle-gauge size corresponds to a smaller diameter needle. This is contrary to French Gauge where an increasing gauge size corresponds to a larger external diameter.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'AXU', N'Anti XA Unit', N'A unit of measure for blood potency. Units for the anti XA activity which is a measure to the anti coagulating effect at low molecular heparins.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'B', N'Bel - level [B]', N'Bel - level [B] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'B10', N'Bit per second', N'In telecommunications and computing, bitrate (sometimes written bit rate, data rate or as a variable R or fb) is the number of bits that are conveyed or processed per unit of time. The bit rate is quantified using the bits per second (bit/s or bps) unit.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'B37', N'Kilogram Force', N'The gravitational force of a kilogram weight or a one-kilogram mass multiplied by the acceleration of standard Earth gravity, equal to 9.8 newtons.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'B47', N'Kilonewton', N'1000 Newtons or amount needed to accelerate 1 kilo of mass at the rate of 1 metre per second squared multiplied by 1000 (Used to measure force of an object where the value exceeds 1000 newtons).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'B60', N'Lumens per Square Meter', N'Lumens per Square Meter')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'B61', N'Lumens Per Watt', N'The number of Lumens expressed per watt.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'B62', N'Lumen Seconds', N'Lumen seconds is the SI derived unit of luminous energy. It is based on the lumen, the SI unit of luminous flux, and the second, the SI base unit of time. The lumen second is sometimes called the Talbot (symbol T).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'B64', N'Lux Seconds', N'Exposure is measured in lux seconds, and can be computed from exposure value (EV) and scene luminance in a specified region.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'BAR', N'Bar (unit of pressure)', N'The bar is widely used in descriptions of pressurea= 1 bar = 100 kilopascals 0.987 atmospheres.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'BB', N'Base box', N'A unit of area of 112 sheets of tin mil products (tin plate, tin free steel or black plate) 14 by 20 inches, or 31,360 square inches.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'BFT', N'Board Foot', N'A specialized unit of measure for the volume of rough lumber (before drying and planing with no adjustments) or planed/surfaced lumber. It is the volume of a one-foot length of a board one foot wide and one inch thick. Some countries utilize the synonym super foot or superficial foot.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'BLL', N'Barrel US', N'There are varying standards for barrel for some specific commodities, including 31 gal for beer, 40 gal for whiskey or kerosene, and 42 gal for petroleum. The general standard for liquids is 31.5 gal or half a hogsheada= the general standard for dry contents is 7,056 Cubic Inches.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'BP', N'Hundred board foot', N'A unit of volume equal to one hundred board foot.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'BPM', N'Beats Per Minute', N'Beats Per Minute')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'BQL', N'Becquerel', N'The becquerel (symbol Bq) is the SI derived unit of radioactivity. One Bq is defined as the activity of a quantity of radioactive material in which one nucleus decays per second. SI uses the becquerel rather than the second for the unit of activity measure to avoid dangerous mistakes: a measurement in becquerels is proportional to activity, and thus a more dangerous source of radiation gives a higher reading. A measurement in seconds is inversely proportional.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'BTU', N'British thermal unit', N'The British thermal unit (BTU or Btu) is a traditional unit of energy. It is approximately the amount of energy needed to heat one pound of water one degree Fahrenheit. One Btu is equal to about 1.06 kilojoules. It is used in the power, steam generation, heating and air conditioning industries.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'BUA', N'Bushel (US)', N'A bushel is an imperial and U.S. customary unit of dry volume, equivalent in each of these systems to 4 pecks or 8 gallons. It is used for volumes of dry commodities (not liquids), most often in agriculture')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'BUI', N'Bushel (UK)', N'A bushel is an imperial and U.S. customary unit of dry volume, equivalent in each of these systems to 4 pecks or 8 gallons. It is used for volumes of dry commodities (not liquids), most often in agriculture')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'C', N'Coulomb', N'Coulomb - electric charge [C] (Unifed Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'C16', N'Millimetre Per Second', N'The number of millimetre per second the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'C18', N'Millimole', N'a millimole is one thousandth of a mole.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'C26', N'Millisecond', N'A millisecond (from milli- and seconda= abbreviation: ms) is a thousandth (1/1000) of a second.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'C34', N'Mole', N'The mole (symbol mol) is the SI base unit of amount of substancea= one of a few units used to measure this physical quantity. A mole will possess mass exactly equal to the substance''s molecular or atomic weight in grams. That is to say, a substance''s atomic or molecular mass in atomic mass units is the same as its molar mass in grams. Because of this, one can measure the number of moles in a pure substance by weighing it and comparing the result to its molecular or atomic weight')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'C45', N'Nanometre', N'Nanometre- a billionth of a metre (10 to the negative 9th power)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'C47', N'Nano Seconds', N'A nanosecond is an SI unit of time equal to one thousand-millionth of a second (or one billionth of a second), that is, 10-9 seconds.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'C52', N'Picometre', N'Picometre- a trillionth of a metre (10 to the negative 12th power)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'C65', N'Pascal Seconds', N'Pascal second ( N s / m 2 ) is the unit of dynamic viscosity which is also known as the absolute viscosity of the fluid. It is the fluids internal resistance to flow . e.g.: water has dynamic viscosity of 1.787 at 273K or 0 degrees Celsius. and 1.002 Pa-second at 293K i.e. 20degrees Celsius.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'C75', N'Picowatt', N'Picowatt is a derived metric SI (System International) measurement unit of power. The picowatt is equal to one trillionth of a watt (10-12W).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'cd', N'Candela', N'Candela [cd] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'CEL', N'Degree Celsius', N'Celsius (also historically known as centigrade) is a temperature scale, the freezing point of water is 0 degrees Celsius (∞C) and the boiling point 100 ∞C (at standard atmospheric pressure), placing the boiling and freezing points of water exactly 100 degrees apart.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'CFU', N'Colony Forming Units', N'Colony Forming Units: is a rough estimate of the number of viable bacteria or fungal cells in a sample.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'CG', N'Card', N'A unit of count defining the number of units of card (card: thick stiff paper or cardboard).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'CGM', N'Centigram', N'A centigram is one hundredth (1/100) of a gram')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'CLT', N'Centilitre', N'A centilitre is one hundredth (1/100) of a litre')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'CMK', N'Square centimetre', N'A square centimetre is an area of a square whose sides are exactly 1 centimetre in length.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'CMQ', N'Cubic centimetre', N'A cubic centimetre is the volume of a cube of side length one centimetre (0.01 m) equal to a millilitre.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'CMT', N'Centimetre', N'A centimetre is equal to one hundredth of a metre.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'CTM', N'Metric Carat', N'Metric Carat.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'CWA', N'Hundred pound (cwt) / hundred weight (US)', N'A unit of weight in the U.S. Customary System equal to 100 pounds (45.36 kilograms) also called cental.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'CWI', N'Hundred weight (UK)', N'A unit of weight in the British Imperial System equal to 112 pounds (50.80 kilograms) also called quintal.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'D03', N'Kilowatt / hour', N'A unit of accumulated energy of a thousand watts over a period of one hour.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'D19', N'Square Metre Kelvin Per Watt', N'Square metre kelvin per watt .The definition is the thermal resistance (or Rd value) of the material, which is measured by the thickness of the material layer divided by the thermal conductivity of the material. and is expressed in m≤K/W.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'D29', N'Terahertz', N'A unit of frenquecy equal to 1012 Hertz')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'D30', N'Terajoule', N'A terajoule is 10π≤ joules')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'D32', N'Terawatt hour', N'A terawatt hour is 109 * kilowat hour or 3.6 petajoules.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'D40', N'Thousand Litre', N'References SI system of units of measure')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'D43', N'Atomic Mass Units (AMU)', N'Atomic Mass Units')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'D53', N'Watt per metre Kelvin', N'Specifies the thermal conductivity (lambda value): a measure of the product''s ability to conduct heat.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'D5', N'Kilogram per square centimetre', N'A kilogram-force per square centimeter (kgf/cm2), often just kilogram per square centimeter (kg/cm2), or kilopond per square centimeter is a unit of pressure using metric units. Its use is now deprecateda= it is not a part of the International System of Units (SI), the modern metric system. The unit is similar to the English unit psi (lbf/in2).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'D55', N'Heat Transfer Coefficient', N'Heat Transfer Coefficient')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'D63', N'Book', N'A unit of count defining the number of books (book: set of items bound together or written document of a material whole).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'D70', N'Calorie - International Table (IT)', N'A calorie is 1/100 of the amount of energy required to warm one gram of air-free water from 0 ∞C to 100 ∞C at standard atmospheric pressurea= this is about 4.190 J. Its use is archaic, having been replaced by the SI unit of energy, the joule. However, in many countries it remains in common use as a unit of food energy. In the context of nutrition, and especially food labelling, the calorie is approximately equal to 4.1868 joules (J), and energy values are normally quoted in kilojoules (kJ) and kilocalories (kcal).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'DAY', N'Days', N'A day is one three hundreds and sixty fifth (1/365) of a year')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'DD', N'Degree (Unit of Angle)', N'A measurement of plane angle, representing 1/360 of a full rotationa= one degree is equivalent to p/180 radians.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'DG', N'Decigram', N'A decigram is one tenth (1/10) of a gram.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'diop', N'Dioptre - refraction of a lens [diop]', N'Dioptre - refraction of a lens [diop] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'DLT', N'Decilitre', N'A decilitre is one tenth (1/10) of a litre.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'DMK', N'Square decimetre', N'A square deciimetre is an area of a square whose sides are exactly 1 deciimetre in length.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'DMQ', N'Cubic decimetre', N'A cubic decimetre is the volume of a cube of side length one decimetre (0.1 m)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'DMT', N'Decimetre', N'A decimetre is equal to one tenth of a metre.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'DRA', N'Dram (US)', N'The dram (archaic spelling drachm) was historically both a coin and a weight. Currently it is both a small mass in the Apothecaries'' system of weights and a small unit of volume. This unit is called more correctly fluid dram or in contraction also fluidram. The term also refers to the fluid dram, a measure of capacity equal 1/8 of a fluid ounce, which means it is exactly equal to 3.696 691 195 312 5 mL in the United States.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'DRI', N'Dram (UK)', N'The dram (archaic spelling drachm) was historically both a coin and a weight. Currently it is both a small mass in the Apothecaries'' system of weights and a small unit of volume. This unit is called more correctly fluid dram or in contraction also fluidram. The fluid dram is defined as 1/8 of a fluid ounce, which means it is exactly equal to 3.551 632 812 500 0 mL in the Commonwealth and Ireland. In England dram came to mean a small draught of cordial or alcohola= hence the term dram-house for the taverns where one could purchase a dram.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'DZN', N'Dozen', N'A unit of count defining the number of units in multiples of 12.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'E09', N'Milliampere hour', N'Milliampere hour')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'E14', N'Kilocalorie', N'A unit of energy equal to 1000 calories.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'E20', N'106 (1000000) bits (binary digits) per second', N'A unit of information equal to 106 (1000000) bits (binary digits) per second.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'E27', N'Dose', N'A unit of count defining the number of doses (dose: a definite quantity of a medicine or drug).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'E32', N'Litre Per Hour', N'Litre Per Hour')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'E34', N'Gigabyte', N'A unit of information equal to 109 bytes.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'E35', N'Terabyte', N'A unit of information equal to 10π≤ bytes.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'E36', N'Petabyte', N'A unit of information equal to 10π5 bytes. 1000 Terabytes = 1 Petabyte')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'E37', N'Pixel', N'A unit of count defining the number of pixels (pixel: picture element).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'E39', N'Dots per inch', N'A unit of count defining the number of dots per linear inch as a measure of the resolution or sharpness of a graphic image.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'E4', N'Gross kilogram', N'A unit of mass defining the total number of kilograms before deductions.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'E55', N'Use', N'A unit of count defining the number of times an object is used.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'EA', N'Each', N'A unit of count defining the number of items regarded as separate units.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ELU', N'ELISA Units', N'Enzyme-linked immunosorbent assay unit, is always associated with a product and a method.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'eV', N'Electronvolt', N'Electronvolt - energy [eV] (Unifed code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'F', N'Farad', N'Farad - electric capacitance [F] (unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'F27', N'Gram Per Hour', N'Gram Per Hour')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'F79', N'Inch of Mercury', N'Inch of mercury conversion of 3,386 39 ◊ 10≥ kg x m?π x s?≤ OR 3386.389 pascals.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'FAH', N'Degree Fahrenheit', N'The Fahrenheit temperature scale, the freezing point of water is 32 degrees Fahrenheit (∞F) and the boiling point 212 ∞F (at standard atmospheric pressure), placing the boiling and freezing points of water exactly 180 degrees apart.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'FH', N'Micromole', N'One millionth (10 -6 ) of a mole.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'FJ', N'Sizing Factor', N'Sizing Factor')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'fmol/L', N'Femtomole per litre [fmol/L]', N'Femtomole per litre [fmol/L] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'FOT', N'Foot', N'The international foot is defined to be equal to 0.3048 meters.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'FP', N'Pound per square foot', N'A non SI unit of Pressure approximately equal to 47.88025 PASCAL''s.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'FR', N'Foot Per Minute', N'The number of foot per minute the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'FS', N'Foot Per Second', N'The number of foot per second the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'FTK', N'Square foot', N'A square foot is an area of a square whose sides are exactly 1 foot in length.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'FTQ', N'Cubic foot', N'A cubic foot is the volume of a cube of side length one foot (0.3048 m) .')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'g(48.h)', N'Gram per 48 hour [g/48 hr]', N'Gram per 48 hour [g/48 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'g/(12.h)', N'Gram per 12 hour [g/12 hr]', N'Gram per 12 hour [g/12 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'g/(4.h)', N'Gram per 4 hour [g/4 hr]', N'Gram per 4 hour [g/4 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'g/(6.h)', N'Gram per 6 hour [g/6 hr]', N'Gram per 6 hour [g/6 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'g/(72.h)', N'Gram per 72 hour [g/72 hr]', N'Gram per 72 hour [g/72 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'g/d', N'Gram per day [g/d]', N'Gram per day [g/d] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'g/g', N'Gram per gram [g/g]', N'Gram per gram [g/g] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'G2', N'US Gallon Per Minute', N'US gallon per minute')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'G21', N'Cup (US)', N'Cup (US)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'G23', N'Peck', N'A peck is an imperial and U.S. customary unit of dry volume, equivalent in each of these systems to 2 gallons, 8 dry quarts, or 16 dry pints.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'G24', N'Tablespoon', N'Tablespoon. 1/2 fluid ounces, 3 teaspoons, 15 millilitres')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'G25', N'Teaspoon', N'Teaspoon. 1/6 fluid ounces or 5 millilitres')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'G26', N'Stere', N'The stere or stËre is a unit of volume in the original metric system equal to one cubic metre. The stËre is typically used for measuring large quantities of firewood or other cut wood,')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'GM', N'Gram Per Square Metre', N'Gram per square metre: In the metric system, the density of all types of paper, paperboard, and fabric, is expressed in terms of grams per square meter (g/m≤). This quantity is commonly called grammage both in English and French (ISO 536), though many English-speaking countries still refer to the weight. The term density here is used somewhat incorrectly, as density is mass by volume. More precisely, it is a measure of the area density, areal density, or surface density.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'GBQ', N'Gigabecquerel', N'A unit of activity equal to 109 becquerels.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'GFI', N'Gram of Fissile Isotope', N'A unit of mass defining the number of grams of a fissile isotope (fissile isotope: an isotope whose nucleus is able to be split when irradiated with low energy neutrons).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'GL', N'Gram Per Litre', N'A unit of measurement of mass concentration that shows how many grams of a certain substance are present in one litre.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'GLI', N'Gallon (UK)', N'The imperial (UK) gallon was legally defined as 4.54609 litres.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'GLL', N'Gallon (US)', N'The U.S. liquid gallon is legally defined as 231 cubic inches, and is equal to exactly 3.785411784 litres or about 0.133680555 cubic feet.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'GRM', N'Gram', N'A gram is defined as one one-thousandth of the kilogram (1◊10-3 kg).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'GRN', N'Grain', N'A grain or troy grain is precisely 64.79891 milligrams. Exactly 7,000 grains per avoirdupois pound.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'GRO', N'Gross', N'A unit of count defining the number of units in multiples of 144 (12 x 12).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'GWH', N'Gigawatt hour', N'A gigaawatt hour is 109 kilowat hour or 3.6 terajoules.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'Gy', N'Gray [Gy]', N'Gray - energy dose [Gy] (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'H', N'Henry [H]', N'Henry - inductance [H] (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'H49', N'Centimetre Per Hour', N'The number of centimetre per hour the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'H67', N'Millimetre Per Hour', N'The number of millimetre per hour the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'H79', N'French gauge', N'The French scale (most correctly abbreviated as Fr, but also often abbreviated as FR or F) is commonly used to measure the catheter size (Circumference is in millimeters), in which 1 Fr = 0.33 mm in diameter. In the French Gauge system as it is also known, the diameter in millimeters of the catheter can be determined by dividing the French size by 3, thus an increasing French size corresponds with a larger diameter catheter. The following equations summarize the relationships: D(mm) = Fr/3 or Fr = D(mm)*3')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'H81', N'Millimetre Per Minute', N'The number of millimetre per minute the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'H83', N'Litre per kilogram', N'Litre per kilogram.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'H87', N'Piece', N'A unit of count defining the number of pieces (piece: a single item, article or exemplar).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'har', N'Hecto - Hectare [ha]', N'Hectare: Hecto-are and hectare are equivalent units [har] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'HC', N'Hundred count', N'A unit of count defining the number of units counted in multiples of 100.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'HD', N'Half dozen', N'A unit of count defining the number of units in multiplt of six (6).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'HEP', N'Histamine Equivalent Prick', N'Histamine equivalent prick testing for allergen.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'HGM', N'Hectogram', N'A hectogram is one hundred (100) grams')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'HLT', N'Hectolitre', N'A hectolitre is one hundred (100) litres.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'HM', N'Mile Per Hour (statute mile)', N'The number of mile per hour the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'HN', N'Conventional millimetre of mercury', N'Conventional millimetre of mercury mm Hg.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'HTZ', N'Hertz', N'A unit of frequency defined as the number of complete cycles per seconda= it is the basic unit of frequency in the International System of Units (SI).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'HUR', N'Hour', N'An hour is a unit of measurement of time of the duration of 60 minutes, or 3600 seconds. It is 1/24 of a median Earth day.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'INH', N'Inches', N'An international inch is defined to be equal to 25.4 millimeters.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'INK', N'Square inch', N'A square inch is an area of a square whose sides are exactly 1 inch in length.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'INQ', N'Cubic inch', N'A cubic inch is the volume of a cube of side length one inch (0.254 m).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'IU', N'Inch Per Second', N'The number of inch per second the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'JOU', N'Joule', N'A joule is the energy exerted by a force of one newton acting to move an object through a distance of one metre.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'K14', N'Foot Per Hour', N'The number of foot per hour the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'K30', N'Gallon (US liquid) Per Second', N'Gallon (US liquid) per second')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'K43', N'Horsepower', N'Horsepower (electric) - Automotive horsepower is defined as a unit of power equal to 550 foot-pounds per second (745.7 watts).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'K6', N'Kilolitre', N'A kilolitre is one thousand (1000) litres.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'kat', N'Katal - catalytic activity [kat]', N'Katal - catalytic activity [kat] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KDW', N'Kilogram drained net weight', N'Kilogram drained net weight.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KGM', N'Kilogram', N'A unit of mass equal to one thousand grams.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'kg/m3', N'Kilogram per cubic metre', N'Kilogram per cubic metre, kg/m3 (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KHY', N'Kilogram of hydrogen peroxide', N'Kilogram of hydrogen peroxide.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KHZ', N'Kilohertz', N'A unit of frenquecy equal to 103 Hertz')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KIU', N'Kallikrein inactivator unit.', N'Kallikrein Inactivator Unit per Milliliter definition: An arbitrary unit of a kallikrein inactivator concentration equal to the concentration at which one milliliter of the mixture contains one unit of the kallikrein inactivator')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KJO', N'Kilojoule', N'A kilojoule is 1000 joules')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KMA', N'Kilogram of methylamine', N'Kilogram of methylamine')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KMH', N'Kilometre Per Hour', N'The number of kilometre per hour the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KMT', N'Kilometre', N'A kilometre is one thousand (1000) metres')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KNI', N'Kilogram of nitrogen', N'Kilogram of nitrogen')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KNM', N'Kilonewton Per Square Metre', N'Kilonewton Per Square Metre')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KNT', N'Knot', N'The number of knots the trade item or part of the trade item moves. A Knot is a unit of speed equal to one nautical mile (1.852 km) per hour, approximately 1.151 mph.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KO', N'The milliequivalence caustic potash per gram of product', N'The milliequivalence caustic potash per gram of product. A unit of count defining the number of milligrams of potassium hydroxide per gram of product as a measure of the concentration of potassium hydroxide in the product.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KPA', N'Kilopascal- a thousand pascals (10 to the 3rd power)', N'Kilopascal- a thousand pascals (10 to the 3rd power)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KPH', N'Kilogram of potassium hydroxide (caustic potash)', N'Kilogram of potassium hydroxide (caustic potash)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KPO', N'Kilogram of potassium oxide', N'Kilogram of potassium oxide')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KPP', N'Kilogram of phosphorus pentoxide (phosphoric anhydride)', N'Kilogram of phosphorus pentoxide (phosphoric anhydride).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KSD', N'Kilogram of substance 90% dry', N'Kilogram of substance 90% dry')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KSH', N'Kilogram of sodium hydroxide (caustic soda)', N'Kilogram of sodium hydroxide (caustic soda)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KT', N'Kit', N'A unit of count defining the number of kits (kit: tub, barrel or pail).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KUR', N'Kilogram of Uranium', N'A unit of mass equal to one thousand grams of uranium.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'kU/L', N'Kilo unit per litre [kU/L]', N'Kilo unit per litre [kU/L] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KWH', N'Kilowatt hour', N'A kilowatt hour is a unit of energy equal to 3.6 megajoules. It is also a common commercial unit of electric energy representing the amount of energy delivered at a rate of 1,000 watts over a period of one hour.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'KWT', N'Kilowatt', N'A kilowatt is one thousand (1000) watts')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'L/min', N'Litre per minute', N'Litre per minute [L/min] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'L/L', N'Litre per litre [L/L]', N'Litre per litre [L/L] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'LBR', N'Pound', N'The international avoirdupois pound of exactly 0.45359237 kilogram.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'LD', N'Litre / Day', N'Litre per day.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'LF', N'Linear foot', N'A unit of count defining the number of feet (12inch) in length of a uniform width object.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'LK', N'Link', N'A unit of distance equal to 0.01 chain.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'LM', N'Linear metre', N'A unit of count defining the number of metres in length of a uniform width object.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'LPA', N'Litre of pure alcohole', N'Litre of pure alcohol.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'LR', N'Layer', N'A unit of count defining the number of layers.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'LTN', N'Ton (UK) or long ton (US)', N'Ton (UK) = 1016 Kg or 2240 Lb.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'LTR', N'Litre', N'A litre is defined as a special name for a cubic decimetre (1 L = 1 dm3 = 103 cm3).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'LUX', N'Lux', N'The International System unit of illumination, equal to one lumen per square meter.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'm[iU]/L', N'Milliinternational units per litre [m[iU]/L]', N'Milliinternational units per litre [m[iU]/L] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'M57', N'Mile Per Minute', N'The number of mile per minute the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'M58', N'Mile Per Second', N'The number of mile per second the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'M60', N'Metre Per Hour', N'The number of meter per hour the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'M62', N'Kilometre Per Second', N'The number of kilometre per second the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'M63', N'Inch Per Minute', N'The number of inch per minute the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'M64', N'Yard Per Second', N'The number of yard per second the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'M65', N'Yard Per Minute', N'The number of yard per minute the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'M66', N'Yard Per Hour', N'The number of yard per hour the trade item or part of the trade item moves.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MAW', N'Megawatt', N'A unit of power defining the rate of energy transferred or consumed when a current of 1000 amperes flows due to a potential of 1000 volts at unity power factor.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MBR', N'Millibar', N'Millibar- a hundred pascals (10 to the 2nd power)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MC', N'Microgram', N'A microgram is one millionth of a gram (0.000001)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MEQ', N'Milliequivalents', N'mEq or milliequivalents, the measure is used in terms of milliequivalents of solute per liter of solvent (or milliNormal where mEq/L = mN). This is especially common for measurement of compounds in biological fluidsa= for instance, the healthy level of potassium in the blood of a human is defined between 3.5 and 5.0 mEq/L. To better be able to denote the units of measure for nutrients as used in business today.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mg/(12.h)', N'Milligram per 12 hour [mg/12 hr]', N'Milligram per 12 hour [mg/12 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mg/d', N'Milligram per day [mg/d]', N'Milligram per day [mg/d] (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mg/dL', N'Milligram per decilitre (mg/dL)', N'Milligram per decilitre (mg/dL) (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mg/g', N'Milligram per gram [mg/g]', N'Milligram per gram [mg/g] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mg/L', N'Milligram per litre [mg/L]', N'Milligram per litre [mg/L] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mg/mg', N'Milligram per milligram [mg/mg]', N'Milligram per milligram [mg/mg] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mg/mL', N'Milligram per millilitre [mg/mL]', N'Milligram per millilitre [mg/mL] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MGM', N'Milligram', N'A milligram is one thousandth of a gram (0.001)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MHZ', N'Megahertz', N'A unit of frenquecy equal to 106 Hertz')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MIK', N'Square mile', N'A square mile is an area of a square whose sides are exactly 1 mile in length.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MIN', N'Minute (unit of time)', N'A minute is a unit of time equal to 1/60th of an hour or 60 seconds')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MIU', N'Million International Unit (NIE)', N'A unit of count defining the number of international units in multiples of 106.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mL/(10.h)', N'Millilitre per 10 hour [mL/10 hr]', N'Millilitre per 10 hour [mL/10 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mL/(12.h)', N'Millilitre per 12 hour [mL/12 hr]', N'Millilitre per 12 hour [mL/12 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mL/(2.h)', N'Millilitre per 2 hour [mL/2 hr]', N'Millilitre per 2 hour [mL/2 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mL/(4.h)', N'Millilitre per 4 hour [mL/4 hr]', N'Millilitre per 4 hour [mL/4 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mL/(5.h)', N'Millilitre per 5 hour [mL/5 hr]', N'Millilitre per 5 hour [mL/5 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mL/(6.h)', N'Millilitre per 6 hour [mL/6 hr]', N'Millilitre per 6 hour [mL/6 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mL/(72.h)', N'Millilitre per 72 hour [mL/72 hr]', N'Millilitre per 72 hour [mL/72 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mL/(8.h)', N'Millilitre per 8 hour [mL/8 hr]', N'Millilitre per 8 hour [mL/8 hr] (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mL/d', N'Millilitre per day [mL/d]', N'Millilitre per day [mL/d] (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mL/h', N'Millilitre per hour [mL/hr]', N'Millilitre per hour [mL/hr] (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mL/min', N'Millilitre per minute', N'Millilitre per minute [mL/min] (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mL/s', N'Millilitre per second', N'Millilitre per second [mL/s] (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MLT', N'Millilitre', N'A millilitre is one thousandth of a litre (0.001)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mm[Hg]', N'Millimetre of mercury [pressure]', N'Millimetre of mercury [pressure] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MMK', N'Square millimetre', N'A square millimetre is an area of a square whose sides are exactly 1 millimetre in length.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mmol/(12.h)', N'Millimole per 12 hour [mmol/12 hr]', N'Millimole per 12 hour [mmol/12 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mmol/(5.h)', N'Millimole per 5 hour [mmol/5 hr]', N'Millimole per 5 hour [mmol/5 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mmol/(6.h)', N'Millimole per 6 hour [mmol/6 hr]', N'Millimole per 6 hour [mmol/6 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mmol/d', N'Millimole per day [mmol/d]', N'Millimole per day [mmol/d] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mmol/g', N'Millimole per gram [mmol/g]', N'Millimole per gram [mmol/g] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mmol/kg', N'Millimole per kilogram [mmol/kg]', N'Millimole per kilogram [mmol/kg] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mmol/kg[H2O]', N'Millimole per kilogram [H2O] [mmol/kg Water]', N'Millimole per kilogram [H2O] [mmol/kg Water] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mmol/L', N'Millimole per litre [mmol/L]', N'Millimole per litre [mmol/L] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mmol/mmol', N'Millimole per millimole [mmol/mmol]', N'Millimole per millimole [mmol/mmol] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MMQ', N'Cubic millimetre', N'A cubic millimetre is the volume of a cube of side length one milliimetre (0.001 m)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MMT', N'Millimetre', N'A millimetre is one thousandth of a metre (0.001)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MON', N'Month', N'Unit of time equal to 1/12 of a year of 365,25 days')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mosm/L', N'Milliosmole per litre [mosm/L]', N'Milliosmole per litre [mosm/L] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MPN', N'Most Probable Number', N'Most Probable Number: is a method of getting quantitative data on concentrations of discrete items from positive/negative (incidence) data..')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MQH', N'Cubic Metre Per Hour', N'Cubic Metre Per Hour')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MTK', N'Square metre', N'A square metre is an area of a square whose sides are exactly 1 metre in length.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MTQ', N'Cubic metre', N'A cubic metre is the volume of a cube of side length one metre.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MTR', N'Metre', N'The metre is the basic unit of length in the International System of Units (SI).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MTS', N'Metre Per Second', N'Defines the speed defined by distance in metres divided by time in seconds.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mU', N'MilliUnit [mU]', N'MilliUnit [mU] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'mU/L', N'MilliUnit per litre [mU/L]', N'MilliUnit per litre [mU/L] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'MWH', N'Megawatt hour (1000 kW.h)', N'A unit of energy defining the total amount of bulk energy transferred or consumed.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'N16', N'Inch of Mercury (32 ∫F)', N'Non SI-conforming unit of pressure according to the Anglo-American and Imperial system for units, whereas the value of 1 inHg meets the static pressure, which is generated by a mercury at a temperature of 32∞F with a height of 1 inch.3,386 38 x 10≥ Pa OR 3386.38 pascals.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'N17', N'Inch of Mercury (60 ∫F)', N'Non SI-conforming unit of pressure according to the Anglo-American and Imperial system for units, whereas the value of 1 inHg meets the static pressure, which is generated by a mercury at a temperature of 60∞F with a height of 1 inch.3,376 85 x 10≥ Pa OR 3376.85 pascals.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'NCL', N'Number of cells', N'Number of cells')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'NEW', N'Newton', N'Newton')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ng/d', N'Nanogram per day [ng/d]', N'Nanogram per day [ng/d] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ng/g', N'Nanogram per gram [ng/g]', N'Nanogram per gram [ng/g] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ng/L', N'Nanogram per litre [ng/L]', N'Nanogram per litre [ng/L] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ng/mL', N'Nanogram per millilitre [ng/mL]', N'Nanogram per millilitre [ng/mL] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'NIU', N'Number of International Units', N'A unit of count defining the number of international units. The International Unit is a unit of measurement for the amount of a substance, based on measured biological activity or effect. The unit is used for vitamins, hormones, some medications, vaccines, blood products, and similar biologically active substances')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'nmol/d', N'Nanomole per day [nmol/d]', N'Nanomole per day [nmol/d] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'nmol/g', N'Nanomole per gram [nmol/g]', N'Nanomole per gram [nmol/g] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'nmol/h/mL', N'Nanomole per hour per millilitre [nmol/hr/mL]', N'Nanomole per hour per millilitre [nmol/h/mL] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'nmol/L', N'Nanomole per litre [nmol/L]', N'Nanomole per litre [nmol/L] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'nmol/mmol', N'Nanomole per millimole [nmol/mmol]', N'Nanomole per millimole [nmol/mmol] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'nmol/nmol', N'Nanomole per nanomole [nmol/nmol]', N'Nanomole per nanomole [nmol/nmol] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'Np', N'Neper', N'Neper - level [Np] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'NPR', N'Number of pairs', N'Number of pairs')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'NTU', N'Nephelometric turbidity unit', N'A unit used for measurement†of concentration of†enzymatic†compounds†through†optical analysis.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'NU', N'Newton Metre', N'Newton Metre')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'OHM', N'Ohm', N'A unit for electrical impedance, the ratio of the voltage phasor to the electric current phasor.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ON', N'Ounces per square yard', N'The weight of one square yard of the material expressed in ounces. Commonly used to express the density or weight of all types of paper, paperboard, and fabric, e.g. 20 OZ or 20 Weight denim has an area density of 20 oz/yd2. The term density here is used somewhat incorrectly, as density is mass by volume. More precisely, it is a measure of the area density, areal density, or surface density.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ONZ', N'Ounce', N'A unit of mass with several definitions, the most commonly used of which are equal to approximately 30 grams')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'OPM', N'Oscillations Per Minute', N'Oscillations Per Minute')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'OZA', N'Fluid ounce (US)', N'A fluid ounce (US) is equal to one sixteenth (1/16) of a US pint or 29.5735295625 millilitres .')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'OZI', N'Fluid ounce (UK)', N'A fluid ounce (UK) is equal to one twentieth (1/20) of a pint (UK) or 28.4130625 millilitres.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'P1', N'Percent', N'A unit of proportion equal to 0.01.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'PAL', N'Pascal', N'The pascal (symbol: Pa) is the SI derived unit of pressure, stress, Young''s modulus and tensile strength. It is a measure of force per unit area, defined as one newton per square metre.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'PD', N'Pad', N'A unit of count defining the number of pads (pad: block of paper sheets fastened together at one end).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'PFU', N'Plaque Forming unit(s)', N'Plaque Forming unit(s)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'pg', N'Picogram [pg]', N'Picogram [pg] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'pg/mL', N'Picogram per millilitre [pg/mL]', N'Picogram per millilitre [pg/mL] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'pmol/d', N'Picomole per day [pmol/d]', N'Picomole per day [pmol/d] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'pmol/g', N'Picomole per gram [pmol/g]', N'Picomole per gram [pmol/g] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'pmol/h/mg', N'Picomole per hour per milligram [pmol/hr/mg]', N'Picomole per hour per milligram [pmol/hr/mg] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'pmol/h/mL', N'Picomole per hour per millilitre [pmol/hr/mL]', N'Picomole per hour per millilitre [pmol/hr/mL] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'pmol/L', N'Picomole per litre [pmol/L]', N'Picomole per litre [pmol/L] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'pmol/mmol', N'Picomole per millimole [pmol/mmol]', N'Picomole per millimole [pmol/mmol] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'PNT', N'Point', N'A single unit on a scale of measurement as part of an incentive program or pricing structure used as a means of making a quantitative evaluation.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'PR', N'Pair', N'A unit of count defining the number of pairs (pair: item described by two''s).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'PRS', N'Potential Renal Solute Load', N'Potential Renal Solute Load')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'PS', N'Pound-force per square inch', N'The pound-force per square inch (symbol: psi or lbf/in2 or lbf/in2) is a unit of pressure or of stress based on avoirdupois units. It is the pressure resulting from a force of one pound-force applied to an area of one square inch. Other abbreviations are used that append a modifier to psi. However, the US National Institute of Standards and Technology recommends that, to avoid confusion, any modifiers be instead applied to the quantity being measured rather than the unit of measure[1] For example, Pg = 100 psi rather than P = 100 psig.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'PTD', N'Dry Pint (US)', N'The United States dry pint is equal one eighth of a US dry gallon or one half US dry quarts. It is used in the United States but is not as common as the liquid pint.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'PTI', N'Pint (UK)', N'A pint (UK) is equal to 1/8 Gallon (UK)a= used primarly as a measure for beer and cider when sold by the glass.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'PTL', N'Liquid pint (US)', N'The US liquid pint is equal one eighth of a United States liquid gallon.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'PTN', N'Portion:', N'Portion')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'Q30', N'pH (potential of Hydrogen)', N'The activity of the (solvated) hydrogen ion (a logarithmic measure used to state the acidity or alkalinity of a chemical solution).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'Q32', N'Femtolitre', N'Femtolitre- Femtolitre is the metric unit of volume equal to one thousand trillionth (European) or one quadrillionth (American) litre. One femtolitre is the same as one cubic micrometre (µm3), , which can be represented numerically as 0.00000000000001/liter.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'Q33', N'Picolitre', N'Picoliter is the metric unit of volume equal to a trillionth (one millionth of a millionth) of a liter, which can be represented numerically as 0.000000000001/liter. just as the prefix nano denotes a billionth part')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'Q34', N'Nanolitre', N'Nanolitre is the metric unit of volume equal to one billionth of a litre, which can be represented numerically as 0.000000001/liter.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'QB', N'Page - hardcopy', N'A unit of count defining the number of hardcopy pages (hardcopy page: a page rendered as printed or written output on paper, film, or other permanent medium).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'QTD', N'Quart (US dry)', N'A US dry quart is equal to 1/32 of a US bushel, exactly 1.101220942715 litres.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'QTI', N'Quart (UK)', N'A quart (UK) is equal to 1/4 (one quarter) of a gallon (UK) or 40 fluid ounces (UK) or 1.1365225 litres.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'QTL', N'Liquid quart (US)', N'A US liquid quart exactly equals 57.75 cubic inches, which is exactly equal to 0.946352946 litres.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'R9', N'Thousand cubic metre', N'Thousand cubic metre')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'RPM', N'Revolutions Per Minute', N'Revolutions Per Minute')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'S', N'Siemens - electric conductance [S]', N'Siemens - electric conductance [S] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'S4', N'Square Metre / Second', N'Metre squared per second.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'SEC', N'Second (unit of time)', N'A second is a unit of time equal to 1/60th of an minute.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'SET', N'Set', N'A unit of count defining the number of sets (set: a number of objects grouped together).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'SMI', N'Mile (statute mile)', N'A statute mile of 5,280 feet (exactly 1,609.344 meters).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'SQE', N'SQ-E', N'Number of allergens based on the SQ-E unit')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'STN', N'Ton (US) or short ton (UK)', N'Ton (US) = 2000 Lb or 907 Kg')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'Sv', N'Sievert [Sv]', N'Sievert - dose equivalent [Sv] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'SX', N'Shipment', N'A unit of count defining the number of shipments (shipment: an amount of a=ods shipped or transported).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'T', N'Tesla [T]', N'Tesla - magnetic flux density [T] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'T3', N'Thousand piece', N'Thousand piece')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'TNE', N'Tonne', N'Metric ton = 1000 Kg')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'TPI', N'Teeth Per Inch', N'Teeth Per Inch')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'u', N'Unified atomic mass unit', N'Unified atomic mass unit [u] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'U/(12.h)', N'Unit per 12 hour [U/12 hr]', N'Unit per 12 hour [U/12 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'U/(2.h)', N'Unit per 2 hour [U/2 hr]', N'Unit per 2 hour [U/2 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'U/d', N'Unit per day [U/d]', N'Unit per day [U/d] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'U/g', N'Unit per gram [U/g]', N'Unit per gram [U/g] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'U/h', N'Unit per 1 hour [U/hr]', N'Unit per 1 hour [U/hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'U/kg', N'Unit per kilogram [U/kg]', N'Unit per kilogram [U/kg] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'U/mL', N'Unit per millilitre [U/mL]', N'Unit per millilitre [U/mL] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'u[iU]/mL', N'Micro international unit per millilitre (u[iU]/mL)', N'Micro international unit per millilitre (u[iU]/mL) (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'U2', N'Tablet', N'A unit of count defining the number of tablets (tablet: a small flat or compressed solid object).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'UA', N'Torr', N'Torr')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ug/d', N'Microgram per day [ug/d] aka [mcg/d]', N'Microgram per day [ug/d] (Unified code) aka [mcg/d]')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ug/dL', N'Microgram per decilitre (ug/dL)', N'Microgram per decilitre (ug/dL) (Unified code) aka [mcg/dL]')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ug/g', N'Microgram per gram [ug/g] aka [mcg/g]', N'Microgram per gram [ug/g] (Unified code) aka [mcg/g]')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ug/h', N'Microgram per hour [ug/h] aka [mcg/h]', N'Microgram per hour [ug/h] aka [mcg/h] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ug/L', N'Microgram per litre [ug/L] aka [mcg/L]', N'Microgram per litre [ug/L] (Unified code) aka [mcg/L]')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ug/min', N'Microgram per minute [ug/min] aka [mcg/min]', N'Microgram per minute [ug/min] (Unified code) aka [mcg/min]')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'ug/mL', N'Microgram per millilitre [ug/mL] aka [mcg/mL]', N'Microgram per millilitre [ug/mL] (Unified code) aka [mcg/mL]')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'um/s', N'Micrometre per second [um/s]', N'Micrometre per second [um/s] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'umol/(2.h)', N'Micromole per 2 hour [umol/2 hr]', N'Micromole per 2 hour [umol/2 hr] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'umol/d', N'Micromole per day [umol/d]', N'Micromole per day [umol/d] (Unifed Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'umol/g', N'Micromole per gram [umol/g]', N'Micromole per gram [umol/g] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'umol/L', N'Micromole per litre [umol/L]', N'Micromole per litre [umol/L] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'umol/min', N'Micromole per minute (umol/min)', N'Micromole per minute (umol/min) (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'umol/min/L', N'Enzyme unit [micromoles/minute] per litre [umol/min/L]', N'Enzyme unit [micromoles/minute] per litre [umol/min/L] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'umol/mmol', N'Micromole per millimole [umol/mmol]', N'Micromole per millimole [umol/mmol] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'umol/umol', N'Micromole per micromole [umol/umol]', N'Micromole per micromole [umol/umol] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'VLT', N'Volt', N'Volt')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'Wb', N'Weber - magnetic flux [Wb]', N'Weber - magnetic flux [Wb] (Unified code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'W/m2', N'Watt per square metre', N'Specifies the maximum heat output of a product. It is the SI unit (International System of Units) for radiative and other energy fluxes. (Unified Code)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'WEE', N'Week', N'A week is a time unit equal to seven days.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'WHR', N'Watt hour', N'The watt-hour is a unit of energy equivalent to one watt of power expended for one hour of timea= it is equal to 3.6 kilojoules. The watt-hour is rarely used to express energy in any form other than electrical.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'WTT', N'Watt', N'A watt is a derived unit of powera= one watt is equivalent to 1 joule (J) of energy per second.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'XRE', N'Retinol Equivalent (RE)', N'A unit used for quantifying the vitamin A value of sources of vitamin A, including both preformed retinoids and precursor carotenoids. Examples include: . One RE is equivalent to 1 mcg of retinol, 6 mcg of Beta-carotene, or 12 mcg of other pro-vitamin A carotenoids. . One RE is equivalent to 0.3 International Units of retinol or 0.1 International Units of Beta-carotene.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'XRO', N'Roll', N'A unit of count defining the number of rolls where Roll is the name of the code value in UN/ECE Recommendation 21, pluralized as appropriate.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_SER', N'Serving', N'A unit of count defining the number of servings being referenced. There is a validation limiting the usage of this UoM.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'XST', N'Sheet', N'A unit of count defining the number of ìSheetsî where ìSheetî is the name of the code value in UN/ECE Recommendation 21, pluralized as appropriate.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_CCA', N'Cold Cramp Amp', N'A measurement of the number of amps a battery can deliver at 0 ∞ F for 30 seconds and not drop below 7.2 volts.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_CFG', N'Colony Forming Units per gram (CFU/g)', N'Colony Forming Units per gram (CFU/g). See CFU for definition of Colony Forming Unit')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_CFP', N'Colony Forming Units per Pound (CFU/lb)', N'Colony Forming Units per Pound (CFU/lb) See CFU for definition of Colony Forming Unit')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_CHD', N'Centisimal Hahnemannian Dilution (CH)', N'A count of attenuation steps or dilution levels representing the homeopathic potency of a substance using the Hahnemannian (CH) method of attenuationa= commonly denoted as CH1, CH2, CH3, etc. Each centesimal attenuation step represents one part source material combined with 99 parts dilution mediuma= commonly denoted as C1, C2, C3, etc.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_DBA', N'Decibar', N'Decibar')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_DWT', N'Penny Weight', N'A pennyweight (abbreviated dwt or denarius weight) is a unit of mass that is equal to 24 grains, 1/20 of a troy ounce, 1/240 of a troy pound, and exactly 1.55517384 grams. (Usage- The pennyweight is the common weight used in the valuation and measurement of precious metals. Jewellers use the pennyweight in calculating the amount and cost of precious metals used in fabricating or casting jewellery. Similarly, dentists and dental labs still use the pennyweight as the measure of precious metals in dental crowns and inlays.)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_HIN', N'Hundredths of an Inch', N'One Hundredth of an inch or 0.01. (Usage- Measurements of Papers, Films, Circuit boards and the like)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_IUK', N'International Units per Kilogram (IU/kg)', N'International Units per Kilogram (IU/kg) a unit of activity or potency for vitamins, hormones, or other substances, defined individually for each substance in terms of the activity of a standard quantity or preparation.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_KVN', N'Korsakovian (K)', N'A count of attenuation steps or dilution levels representing the homeopathic potency of a substance using the Korsakovian (K) method of attenuationa= commonly denoted as CK1, CK2, CK3, etc. Each centesimal attenuation step represents one part source material combined with 99 parts dilution mediuma= commonly denoted as C1, C2, C3, etc.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_MLM', N'Millesimai (LM)', N'A count of attenuation steps or dilution levels representing the homeopathic potency of a substance where each attenuation step represents one part source material combined with 49,999 parts dilution mediuma= commonly denoted as LM1, LM2, LM3, etc.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_MPG', N'Miles Per Gallon', N'Fuel usage by the number of miles the trade item can a= on 1 US gallon of fuel.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_MTC', N'Mother tincture (Dry material)', N'A count of a dry crud medical substance Mother tincture, when used for homeopathic preparations, are liquid preparations obtained by the solvent action of a suitable vehicle upon raw materials. The raw materials (medical substance) are usually in the fresh form but may be dried. Mother tinctures for homeopathic preparations may also be obtained from plant juices, with, or without the addition of a vehicle.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_NGM', N'Nanogram', N'A nanogram is 10-9 gram or a billionth of a gram.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_PPC', N'Pixel per centimetre', N'A unit of count defining the number of pixels per linear centimetre as a measurement of the resolution of devices in various contextsa= typically computer displays, image scanners or digital camera image sensors.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_PPI', N'Pixel per inch', N'A unit of count defining the number of pixels per linear inch (PPI) as a measurement of the resolution of devices in various contextsa= typically computer displays, image scanners or digital camera image sensors.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_RAE', N'Retinol Activity Equivalents', N'A measure of vitamin A activity based on the capacity of the body to convert provitamin carotenoids containing at least one unsubstituted ionone ring to retinaldehyde. 1 microgram RAE = 1 mg retinol = 12 mg ﬂ-carotene = 24 mg other vitamin A precursor carotenoids.')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_SIN', N'Thirty Seconds of an Inch', N'Thirty Seconds of an inch (1/32 or 0.3125). (Usage when measuring using some measuring tapes. Commonly used in model making and some hobbies for scaling to full size. Also used in railway applications.)')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_SPS', N'Sample per second', N'A unit of count defining the number of samplings takes during a period of time')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'X_UIN', N'Ten Thousandths of an Inch', N'One Ten Thousandths of an inch or 0.0001 often called a tenth in machinery. (Usage- Measurements of some machine parts, gaps in spark plugs).')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'YDK', N'Square Yard', N'A square yard is the area of a square with sides of one yard (three feet, thirty-six inches, 0.9144 metres) in length')"
a = a + ";"
a = a + "insert [dbo].[MONADES] ([Code value], [Name], [Description]) VALUES (N'YRD', N'Yard', N'A yard is It is equal to 3 feet or 36 inches or 0.9144 meter.')"
'apot7.Text2.Text = a

Gdb.Execute a


End Sub

Private Sub Unilever_Click()
  
 'csv_send Now, Now, 0, filesql, filecsv
  
  Dim APO, EOS As Date
APO = Now
EOS = Now

    csv_send APO, EOS, 0, "C:\MERCVB\UNILEVER\SQL\ELAIS10", "c:\mercvb\UNILEVER\010-066159748-" + Format(Now, "yyyyMMddhhmm") + ".csv" 'TSIATSIARIMARIA@GMAIL.COM
    csv_send APO, EOS, 0, "C:\MERCVB\UNILEVER\SQL\ELAIS20", "c:\mercvb\UNILEVER\020-066159748-" + Format(Now, "yyyyMMddhhmm") + ".csv"
    csv_send APO, EOS, 0, "C:\MERCVB\UNILEVER\SQL\ELAIS30", "c:\mercvb\UNILEVER\030-066159748-" + Format(Now, "yyyyMMddhhmm") + ".csv"
    csv_send APO, EOS, 0, "C:\MERCVB\UNILEVER\SQL\ELAIS40", "c:\mercvb\UNILEVER\040-066159748-" + Format(Now, "yyyyMMddhhmm") + ".csv"

  
  
End Sub

Public Sub csv_send(ByVal APO As Date, ByVal EOS As Date, weeks As Integer, filesql As String, filecsv As String)

'NCFTPPUT -u GFK004998 -p 15KLS.wZYaRQ2OYgrrZ1 gimftp.gfk.com /  c:\mercvb\mercpath.txt
Dim DUM

   Dim sql As String
   Dim mfile2 As String
   Dim mfile As String
   Dim a As String
   
   '---------------------------------HEADER  ----------------------------------
   Dim mpath As String
   mpath = "c:\mercvb\queries\"
   
    mfile = "c:\mercvb\QUERIES\ret1path.txt"
    Dim mAA As String  ' mAA epipleon sto onoma na katalabainei an einai to proto h to deytero zeygari
    
    mAA = ""
 

   
   sql = ""
   
  
    mfile = filesql '  "C:\mercvb\queries\RYTMIZOMENES\GFK"
    If Len(Dir(mfile, vbNormal)) > 0 Then
       Open mfile For Input As #1
     
         Do While Not EOF(1)
           Line Input #1, a
           sql = sql + a
         Loop

       Close #1
     Else
        sql = "" ' " select KODE,ONOMA,POSO,TIMM,MONA,CONVERT(CHAR(10),HME,103),ATIM AS HMEPOM from EGGTIM WHERE LEFT(ATIM,1) IN ('T','L','Y')"
     End If
     
       
       
       
    
  
    'mfile = "c:\mercvb\QUERIES\ret1"
   
'   If weeks = 1 Then
'      sql = sql + " and  DATEPART(week, HME)= DATEPART(week, GETDATE())-1  "
'   Else
'      sql = sql + " AND HME>='" + Format(apo, "MM/dd/yyyy") + "' AND HME<='" + Format(eos, "MM/dd/yyyy") + "'"
'   End If
   
   
   mfile2 = filecsv '  mpath + Format(Now, "yyyyMMdd") + ".CSV"




    PRINT_CSV sql, "0000000000000", "1", 0, 0, mfile2
    
On Error Resume Next
FileCopy "c:\print", mfile2
    
    
    
'  Open "c:\mercvb\ss.bat" For Output As #1
'  Print #1, "c:\mercvb\NCFTPPUT.exe -u GFK004998 -p 15KLS.wZYaRQ2OYgrrZ1 gimftp.gfk.com /  " + mfile2
'  Close #1
'  DUM = Shell("cmd.exe /c c:\mercvb\ss.bat", vbMinimizedFocus)

End Sub

Public Sub gfk_send(APO As Date, EOS As Date, weeks As Integer)

'NCFTPPUT -u GFK004998 -p 15KLS.wZYaRQ2OYgrrZ1 gimftp.gfk.com /  c:\mercvb\mercpath.txt
Dim DUM

   Dim sql As String
   Dim mfile2 As String
   Dim mfile As String
   Dim a As String
   
   '---------------------------------HEADER  ----------------------------------
   Dim mpath As String
   mpath = "c:\mercvb\queries\"
   
    mfile = "c:\mercvb\QUERIES\ret1path.txt"
    Dim mAA As String  ' mAA epipleon sto onoma na katalabainei an einai to proto h to deytero zeygari
    
    mAA = ""
 

   
   sql = ""
   
  
    mfile = "C:\mercvb\queries\RYTMIZOMENES\GFK"
    If Len(Dir(mfile, vbNormal)) > 0 Then
       Open mfile For Input As #1
     
         Do While Not EOF(1)
           Line Input #1, a
           sql = sql + a
         Loop

       Close #1
     Else
        sql = " select KODE,ONOMA,POSO,TIMM,MONA,CONVERT(CHAR(10),HME,103),ATIM AS HMEPOM from EGGTIM WHERE LEFT(ATIM,1) IN ('T','L','Y')"
     End If
     
       
       
       
    
  
    'mfile = "c:\mercvb\QUERIES\ret1"
   
   If weeks = 1 Then
      sql = sql + " and  DATEPART(week, HME)= DATEPART(week, GETDATE())-1  "
      mfile2 = mpath + "PyxidaData" + Format(DatePart("ww", Now), "00") + Format(Now, "yyyy") + ".CSV"
      
      
       
   Else
       sql = sql + " AND HME>='" + Format(APO, "MM/dd/yyyy") + "' AND HME<='" + Format(EOS, "MM/dd/yyyy") + "'"
       mfile2 = mpath + Format(Now, "yyyyMMdd") + ".CSV"
   End If
   
   
  ' mfile2 = mpath + Format(Now, "yyyyMMdd") + ".CSV"

'PyxidaData252026 „È· ÙÁÌ Â‚‰ÔÏ‹‰· 25 ÙÔı 2026.


    PRINT_CSV sql, "0000000000000", "", 0, 0, mfile2
    
   ' Shell
    
Open "c:\mercvb\ss.bat" For Output As #1
  Print #1, "c:\mercvb\NCFTPPUT.exe -u GFK004998 -p 15KLS.wZYaRQ2OYgrrZ1 gimftp.gfk.com /  " + mfile2
  Close #1
 ' Run "c:\mercvb\ss.bat"
  
'  DUM = Shell("c:\TYP.BAT ", vbMinimizedNoFocus)
 DUM = Shell("cmd.exe /c c:\mercvb\ss.bat", vbMinimizedFocus)

End Sub

'Private Sub Command12_Click()
'
'' With TDBGrid.PrintInfo
''         '.PageSetup
''        ' Set the page header
''      .PageHeaderFont.Italic = True
''        .PageHeader = "≈ÍÙ˝˘ÛÁ " + Date$
''
''
''         'Column headers will be on every page
''        .RepeatColumnHeaders = True
''
''        ' Display page numbers (centered)
''        .PageFooter = "\tÛÂÎﬂ‰·: \p"
''        ' Invoke Print Preview
''        .PrintPreview
''  End With
'
'
'
'End Sub

'Function ALIGNFORMAT(MM, pic)
'' RIGHT ALIGNMENT
'    If Left(pic, 1) = "#" Then
'        ALIGNFORMAT = Right(Space(Len(pic)) + MM, Len(pic))
'    Else
'        ' LEFT ALIGNMENT
'        ALIGNFORMAT = Left(MM + Space(Len(pic)), Len(pic))
'    End If
'
'
'End Function

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>

100     Adodc2.Recordset.MoveFirst

        Dim par As String

110     Do While Not Adodc2.Recordset.EOF

120         If Left(Adodc2.Recordset("ATIM"), 1) = "A" Then
130             par = "ƒ¡" + mID(Adodc2.Recordset("ATIM"), 2, 6) + "¬"
140         ElseIf Left(Adodc2.Recordset("ATIM"), 1) = "L" Then
150             par = "ƒ¡" + mID(Adodc2.Recordset("ATIM"), 2, 6) + "√"

160         ElseIf Left(Adodc2.Recordset("ATIM"), 1) = "E" Then
170             par = "ƒ≈" + mID(Adodc2.Recordset("ATIM"), 2, 6)
            Else
180             par = Adodc2.Recordset("ATIM")
            End If

190         STR_EKT = Format(Adodc2.Recordset("HME"), "DD/MM/YYYY") + " " + Left(par + Space(10), 10) + " " + Left(Adodc2.Recordset("KODE") + Space(10), 10) + " " + Left(Adodc2.Recordset("ONOMA") + Space(25), 25) + " " + " …¬" + " " + Left(Format(Adodc2.Recordset("TIMM"), "###") + Space(4), 4) + " TEM " + " " + Right("        " + Format(Adodc2.Recordset("XRE"), "###0"), 6) + " " + Right("       " + Format(Adodc2.Recordset("PIS"), "###0"), 6)

200         MILSEC 100

210         Open "LPT1" For Output As 1
220         Print #1, Chr(15) + to437(STR_EKT)
230         Close 1

240         Adodc2.Recordset.MoveNext

        Loop

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command4_Click()

        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>

100     STR_EKT = "HMEP/NIA  " + " " + "–¡—/ œ    " + " " + " Ÿƒ… œ”   " + " " + "–≈—…√—¡÷«                " + " " + "”’”" + " " + "     ÃœÕ " + " " + "≈…”¡√ " + " " + "≈Œ¡√Ÿ√"

110     Open "LPT1" For Output As 1
120     Print #1, Chr(15) + to437(STR_EKT)
130     Close 1

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command5_Click()

        ' ME EXCEL
        ' TO C:\PAR.XLS STELNEI TIS EIKONES APO TO MERCVB\IMAGES  STO C:\PAR
        'PREPEI H 1H STHLH NA EXEI 'p  kai h 3h sthlh ton kodiko
        '<EhHeader>
        On Error GoTo Command5_Click_Err

        '</EhHeader>

        Dim xl As New excel.Application

        Dim xlsheet As excel.Worksheet

        Dim xlsheet3 As excel.Worksheet

        Dim xlwbook As excel.workbook

100     Set xlwbook = xl.Workbooks.Open("C:\EKTYP.XLS")
110     Set xlsheet = xlwbook.Sheets.Item(1)

        Dim k

        Dim a$, B, X

        Dim D

        Dim FPA As Integer

        Dim rSQL As New ADODB.Recordset

        Dim kod, per, MON, POS, TIMH, TIMOL

        Dim F As String, RECS As Integer

120     k = 5

        Dim CH1, CH2, CH3

        Dim XTI, lti

        Dim s As String

        Dim nn As Long

        Dim LL2 As Long

130     Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF

140         If IsNull(xlsheet.cells(k, 1)) Then

                Exit Do

            End If

150         If IsEmpty(xlsheet.cells(k, 1)) Then

                Exit Do

            End If

            On Error Resume Next

160         kod = LTrim(str(xlsheet.cells(k + 1, 3)))    '

170         If UCase(LTrim(str(xlsheet.cells(k + 1, 1)))) = "P" Then
180             FileCopy "c:\mercvb\images\" + kod + ".JPG", "c:\par\" + kod + ".JPG"
            End If

190         k = k + 1
        Loop

200     Set xlsheet = Nothing
210     Set xlwbook = Nothing
220     xl.Quit

230     MsgBox "œ "

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command6_Click()
'
'        '√≈÷’—¡ ”≈ EPSILON
'        '<EhHeader>
'        On Error GoTo Command6_Click_Err
'
'        '</EhHeader>
'        Dim R    As New ADODB.Recordset
'
'        Dim Rp   As New ADODB.Recordset
'
'        Dim mait As String, mAit2 As String
'
'        Dim MAFM, mEpo
'
'        Dim synal, OK, i, Msynal(500) As String
'
'        Dim nSYN
'
'        nSYN = 0
'
'        Dim AJIA_KEPYO As Single
'
'        On Error GoTo SHOW_ERROR
'
'        Dim mfil
'
'
'
'Dim M_PK  ' Àœ√¡—…¡”Ãœ” –…”‘. ¡—‘œÕ
'M_PK = "30-02-00-0000"
'
'100     mfil = "C:\KINHS" + Format(Now, "YYYYMMDDHHmm")
'110     Open mfil For Output As #1
'
'        Dim m70, m54, m54b, m30, m38
'
'120     SS = Get_GEF()
'
'130     SS = InputBox(" Ÿƒ… œ… –ŸÀ«”≈ŸÕ", , SS)
'
'        Dim synt1 As String, synt2 As String
'
'140     If epan.Value = vbChecked Then
'150         synt1 = ""
'160         synt2 = ""
'        Else
'170         synt1 = " (B_C1 is null or LEFT(B_C1,1)<>'*') and "
'180         synt2 = " (ATIM2 IS NULL OR LEFT(ATIM2,1)<>'*') AND "
'        End If
'
'190     R.Open "SELECT TIM.*,PEL.EPO,PEL.AFM FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD WHERE " + synt1 + " HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' AND LEFT(ATIM,1) in (" + SS + ")", Gdb, adOpenDynamic, adLockOptimistic
'200     mKPE = "0004450"
'        'mAit = "–ŸÀ«”≈…” ≈Ã–œ—≈’Ã¡‘ŸÕ 23%"
'        ' mAit = "–ŸÀ«”≈…” ≈Ã–œ—≈’Ã¡‘ŸÕ 13%"
'        ' mAit = "–ŸÀ«”≈…” ≈Ã–œ—≈’Ã¡‘ŸÕ 9%"
'210     mXPI_X = "0" ' ˜ÒÂ˘ÛÁ
'220     mXPI_P = "1" ' –…”‘Ÿ”«
'        'm30 = "30-00-00-0000"
'230     m38 = "38-00-00-0000"
'
'        Dim aji
'
'        Dim MATIM
'
'        Dim m_aj2 As Single
'
'
'        Dim mPOL      As String
'
'        Dim AJ_KAU    As Single
'
'        Dim kepyo_fpa As Single
'
'        Dim G         As String
'
'        Dim isKepyo   As String
'
'240     N = 1
'
'
'
'        Dim M_ELGA As Single
'250     Do While Not R.EOF
'
'            AJ_KAU = R!aj1 + R!aj2 + R!aj3 + R!aj4 + R!aj5 + R!aj7 + R!aj6
'            kepyo_fpa = R!fpa1 + R!fpa2 + R!fpa3 + R!FPA4
'            AJIA_KEPYO = AJ_KAU
'
'
'
'                   If IsNull(R!ELGA) Then
'                       M_ELGA = 0
'                   Else
'                       M_ELGA = R!ELGA
'                   End If
'
'
'
'            If Left(R!ATIM, 1) = "p" Or Left(R!ATIM, 1) = "P" Then
'                AJIA_KEPYO = -AJ_KAU
'                kepyo_fpa = -kepyo_fpa
'            End If
'
'260         aji = R!aj1 + R!aj2 + R!aj3 + R!aj4 + R!aj5 + R!fpa1 + R!fpa2 + R!fpa3 + R!FPA4 + R!aj7 + R!aj6 + R!FPA6
'270         Rp.Open "SELECT * FROM PARASTAT WHERE EIDOS='" + Left(R!ATIM, 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
'
'280         If Rp!pol = 1 Then ' –ŸÀ«”≈…”
'290             mAit2 = to437(R("epo")) '"–ŸÀ«”≈…” "
'300             mPOL = "13"   '  "1310"
'                isKepyo = "1"
'            Else
'310             mAit2 = to437(R("epo")) ' "¡√œ—≈” "
'320             mPOL = "12"    '"1220"
'                isKepyo = "2"
'            End If
'
'330         If IsNull(Rp("XREPEL")) Or Len(Trim(Rp!xrepel)) = 0 Then ' –ŸÀ«”≈…”
'340             m30 = Rp!PISPEL
'350             mXPI_X = "1" ' –…”‘Ÿ”«    '¡√œ—≈”  –…”‘Ÿ‘… ¡
'360             mXPI_P = "0" ' ◊—≈Ÿ”«
'            Else
'370             m30 = Rp!xrepel
'380             mXPI_X = "0" ' XREOSI          ' –ŸÀ«”≈…”
'390             mXPI_P = "1" ' PISTOSI
'            End If
'
'            If Left(m30, 5) = "30-01" Or Left(m30, 5) = "50-01" Then ' ≈Õƒœ œ…Õœ‘… ≈”
'                kepyo_fpa = 0
'                isKepyo = "0"
'                AJ_KAU = 0
'                mKPE = "0000000"
'
'                If Len(R!AFM) < 9 Then  ' KAI DEN EINAI ENDOKOINOTIKH
'                    MAFM = "000000000"
'                Else
'                    MAFM = R("afm")
'                End If
'
'                GoSub ADDSYNAL
'            Else
'
'                If InStr("lLpZ", Left(R!ATIM, 1)) > 0 Then
'
'                    If IsNull(R("afm")) Then
'                            mEpo = "ÎÈ·ÌÈÍÁ"
'                            MAFM = "000000000"
'                            mKPE = "0000000"
'                    Else
'                       If Len(Trim(R!AFM)) < 9 Then
'                            mEpo = "ÎÈ·ÌÈÍÁ"
'                            MAFM = "000000000"
'                            mKPE = "0000000"
'                       Else
'                            mKPE = mID(R("afm"), 3, 7)
'
'                            MAFM = R("afm")
'
'                            mEpo = mID(R("EPO"), 1, 30)
'                       End If
'                    End If
'
'
'                    GoSub ADDSYNAL
'                Else
'                    MAFM = R("afm")
'                    mEpo = mID(R("EPO"), 1, 30)
'
'                    If IsNull(R!AFM) Then    ' KAI DEN EINAI ENDOKOINOTIKH
'                        MsgBox "–—œ”œ◊« ƒ≈Õ ≈◊≈… ¡÷Ã ”‘œ " + R!ATIM + " " + Format(R!hme, "DD/MM/YYYY")
'                        MAFM = "000000000"
'                    Else
'
'                        If Len(R!AFM) < 9 Then  ' KAI DEN EINAI ENDOKOINOTIKH
'                            MsgBox "–—œ”œ◊« ƒ≈Õ ≈◊≈… ¡÷Ã ”‘œ " + R!ATIM + " " + Format(R!hme, "DD/MM/YYYY") + Chr(13) + " ‘œ–œ»≈‘«»« ≈ ¡÷Ã=000000000"
'                            MAFM = "000000000"
'                            mKPE = "0000000"
'                        Else
'                            mKPE = mID(R("afm"), 3, 7)
'                        End If
'                    End If
'
'                    GoSub ADDSYNAL
'                End If
'            End If
'
'400         mait = to437(R("epo"))
'
'410         If IsNull(Rp("synt_titl")) Then
'420             MsgBox "‰ÂÌ ı‹Ò˜ÂÈ ÛıÌÙÔÏÔ„Ò·ˆﬂ· „È· ÙÔ " + Rp("titlos") + "== ƒ…¡ œ–«"
'
'                Exit Sub
'
'            End If
'
'430         MATIM = to437(Rp("synt_titl") + mID(R("ATIM"), 2, 6))
'
'            mPOL = mPOL + isKepyo + "0"
'
'            '----------------------------------30 ------------------ 50 --------------------------------------
'
'            G = "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(46), 46) + Replace(Right(Space(14) + Format(kepyo_fpa, "######0.00"), 14), ".", ",") + Replace(Right(Space(14) + Format(AJIA_KEPYO, "######0.00"), 14), ".", ",") + Left(m30 + Space(7), 20) + mXPI_X + Replace(Format(aji - M_ELGA, "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'
'            If InStr("lLp", Left(R!ATIM, 1)) > 0 Then
'                If Len(G) < 200 Then
'                    G = G + String(200 - Len(G), " ")
'                End If
'
'                Mid(G, 200, 1) = "1"
'            End If
'
'950         Print #1, G
'
'440         If R("FPA1") > 0 Then
'450             If IsNull(Rp("LAJ1")) Or IsNull(Rp("LFPA1")) Then
'460                 MsgBox "ƒ≈Õ ’–¡—◊≈… « –¡—¡Ã≈‘—œ” √…¡ ‘œ –¡—¡”‘¡‘… œ " + Rp("TITLOS") + " √…¡ ‘œ ÷–¡ 13%"
'                Else
'470                 m70 = Rp("LAJ1")
'480                 m54 = Rp("LFPA1")
'490                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(R("aj1"), "00000000000.00"), ".", ",") + Left(m70 + Space(7), 20) + mXPI_P + Replace(Format(R("aj1"), "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'
'500                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m54 + Space(7), 20) + mXPI_P + Replace(Format(R("FPA1"), "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'                End If
'            End If
'
'510         If R("FPA2") > 0 Then
'520             If IsNull(Rp("LAJ2")) Or IsNull(Rp("LFPA2")) Then
'530                 MsgBox "ƒ≈Õ ’–¡—◊≈… « –¡—¡Ã≈‘—œ” √…¡ ‘œ –¡—¡”‘¡‘… œ " + Rp("TITLOS") + " √…¡ ‘œ ÷–¡ 23%"
'                Else
'540                 m70 = Rp("LAJ2")
'550                 m54 = Rp("LFPA2")
'
'
'                    If IsNull(R!Parakratisi) Then
'                        m_aj2 = R!aj2
'                    Else
'                        m_aj2 = R!aj2 - R!Parakratisi
'                    End If
'
'
'
'
'560                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(R("aj2"), "00000000000.00"), ".", ",") + Left(m70 + Space(7), 20) + mXPI_P + Replace(Format(m_aj2, "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'570                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m54 + Space(7), 20) + mXPI_P + Replace(Format(R("FPA2"), "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'                End If
'
'            End If
'
'580         If R("FPA3") > 0 Then
'590             If IsNull(Rp("LAJ3")) Or IsNull(Rp("LFPA3")) Then
'600                 MsgBox "ƒ≈Õ ’–¡—◊≈… « –¡—¡Ã≈‘—œ” √…¡ ‘œ –¡—¡”‘¡‘… œ " + Rp("TITLOS") + " √…¡ ‘œ ÷–¡ 19%"
'                Else
'610                 m70 = Rp("LAJ3")
'620                 m54 = Rp("LFPA3")
'630                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(R("aj3"), "00000000000.00"), ".", ",") + Left(m70 + Space(7), 20) + mXPI_P + Replace(Format(R("aj3"), "00000000000.00"), ".", ",") + Left(ATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'640                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m54 + Space(7), 20) + mXPI_P + Replace(Format(R("FPA3"), "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'                End If
'
'            End If
'
'650         If R("FPA4") > 0 Then
'660             If IsNull(Rp("LAJ4")) Or IsNull(Rp("LFPA4")) Then
'670                 MsgBox "ƒ≈Õ ’–¡—◊≈… « –¡—¡Ã≈‘—œ” √…¡ ‘œ –¡—¡”‘¡‘… œ " + Rp("TITLOS") + " √…¡ ‘œ ÷–¡ 9%"
'                Else
'680                 m70 = Rp("LAJ4")
'690                 m54 = Rp("LFPA4")
'700                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(R("aj4"), "00000000000.00"), ".", ",") + Left(m70 + Space(7), 20) + mXPI_P + Replace(Format(R("aj4"), "00000000000.00"), ".", ",") + Left(ATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'710                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m54 + Space(7), 20) + mXPI_P + Replace(Format(R("FPA4"), "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'                End If
'
'            End If
'
'720         If R("AJ5") > 0 Then
'730             If IsNull(Rp("LAJ5")) Then
'740                 MsgBox "ƒ≈Õ ’–¡—◊≈… « –¡—¡Ã≈‘—œ” √…¡ ‘œ –¡—¡”‘¡‘… œ " + Rp("TITLOS") + " √…¡ ‘œ ÷–¡ 0%"
'                Else
'
'
'
'750                    m70 = Rp("LAJ5")
'760                    Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(R("aj5"), "00000000000.00"), ".", ",") + Left(m70 + Space(7), 20) + mXPI_P + Replace(Format(R("aj5"), "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'
'                    If M_ELGA > 0 Then
'                       m54 = Rp("LPARAKRAT")
'                       Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m54 + Space(7), 20) + mXPI_X + Replace(Format(R("ELGA"), "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'                    End If
'                    'End If
'                End If
'            End If
'
'
'
'' ≈‘”… «‘¡Õ –—…Õ √…Õ≈… ¡ÀÀ¡√« ‘œ’ ÷–¡ .  ‘Ÿ—¡ –À≈œÕ ◊—«”…Ãœ–œ…Ÿ ‘œ 24% ”≈ ¡’‘«Õ ‘« »≈”«
''770         If R("AJ6") > 0 Then
''780             If IsNull(Rp("LAJ6")) Then
''790                 MsgBox "ƒ≈Õ ’–¡—◊≈… « –¡—¡Ã≈‘—œ” √…¡ ‘œ –¡—¡”‘¡‘… œ " + Rp("TITLOS") + " √…¡ ‘œ ÷–¡ KAT 6 0%"
''                Else
''800                 m70 = Rp("LAJ6")
''810                 m54 = Rp("LFPA6")
''820                 m54b = Rp("LFPA16")
''
''830                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(R("aj6"), "00000000000.00"), ".", ",") + Left(m70 + Space(7), 20) + mXPI_P + Replace(Format(R("aj6"), "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
''840                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(Round(R("aj6") * 0.13, 2), "00000000000.00"), ".", ",") + Left(m54 + Space(7), 20) + mXPI_P + Replace(Format(Round(R("aj6") * 0.13, 2), "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
''
''850                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(Round(R("aj6") * 0.13, 2), "00000000000.00"), ".", ",") + Left(m54b + Space(7), 20) + mXPI_X + Replace(Format(Round(R("aj6") * 0.13, 2), "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
''
''                End If
''            End If
''
'
'
'
'770         If R("AJ6") > 0 Then
'780             If IsNull(Rp("LAJ6")) Then
'790                 MsgBox "ƒ≈Õ ’–¡—◊≈… « –¡—¡Ã≈‘—œ” √…¡ ‘œ –¡—¡”‘¡‘… œ " + Rp("TITLOS") + " √…¡ ‘œ ÷–¡ KAT 24%"
'                Else
'800                 m70 = Rp("LAJ6")
'810                 m54 = Rp("LFPA6")
'820                 'm54b = Rp("LFPA16")
'
'
'                    If IsNull(R!Parakratisi) Then
'                        m_aj2 = R!aj6
'                    Else
'                        m_aj2 = R!aj6 - R!Parakratisi
'                    End If
'
'
'
'830                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(R("aj6"), "00000000000.00"), ".", ",") + Left(m70 + Space(7), 20) + mXPI_P + Replace(Format(m_aj2, "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'840                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(Round(R("aj6") * 0.13, 2), "00000000000.00"), ".", ",") + Left(m54 + Space(7), 20) + mXPI_P + Replace(Format(Round(R("FPA6"), 2), "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'
''850                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(Round(R("aj6") * 0.13, 2), "00000000000.00"), ".", ",") + Left(m54b + Space(7), 20) + mXPI_X + Replace(Format(Round(R("aj6") * 0.13, 2), "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'
'                End If
'            End If
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
'
'
'
'
'
'
'860         If R("AJ7") > 0 Then
'870             If IsNull(Rp("LAJ6")) Then
'880                 MsgBox "ƒ≈Õ ’–¡—◊≈… « –¡—¡Ã≈‘—œ” √…¡ ‘œ –¡—¡”‘¡‘… œ " + Rp("TITLOS") + " √…¡ ‘œ ÷–¡ KAT 6 0%"
'                Else
'
'890                 m54 = Rp("LFPA7")
'900                 m70 = Rp("LAJ7")
'910                 m54b = Rp("LFPA17")
'
'920                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(R("aj7"), "00000000000.00"), ".", ",") + Left(m70 + Space(7), 20) + mXPI_P + Replace(Format(R("aj7"), "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'
'930                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(Round(R("aj7") * 0.23, 2), "00000000000.00"), ".", ",") + Left(m54 + Space(7), 20) + mXPI_P + Replace(Format(Round(R("aj7") * 0.23, 2), "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'
'940                 Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(Round(R("aj7") * 0.23, 2), "00000000000.00"), ".", ",") + Left(m54b + Space(7), 20) + mXPI_X + Replace(Format(Round(R("aj7") * 0.23, 2), "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'
'                End If
'            End If
'
'            If R("PARAKRATISI") > 0 Then
'
'                If IsNull(Rp("LPARAKRAT")) Then
'                   MsgBox "ƒ≈Õ ’–¡—◊≈… « –¡—¡Ã≈‘—œ” √…¡ ‘œ –¡—¡”‘¡‘… œ " + Rp("TITLOS") + " √…¡ ‘œ ÷–¡ KAT 6 0%"
'                Else
'                   m54 = Rp("LPARAKRAT")
'                   aji = aji - R("PARAKRATISI")
'                   Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(Round(0, 2), "00000000000.00"), ".", ",") + Left(m54 + Space(7), 20) + mXPI_X + Replace(Format(Round(R("PARAKRATISI"), 2), "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'                End If
'            End If
'
'
'
'
'
'
'
'
'            '30
'            '950    Print #1, "4" + Format(n, "00000") + mPOL + mKPE + Left(MATIM + Space(18), _
'            '             18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(46), 46) + _
'            '             Replace(Format(kepyo_fpa, "00000000000.00"), ".", ",") + Replace(Format(AJ_KAU, "00000000000.00"), ".", ",") + Left(m30 + Space(7), 20) + _
'            '             mXPI_X + Replace(Format(aji, "00000000000.00"), ".", ",") + Left(MATIM + _
'            '             Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'
'            'METRHTA
'960         If R("TRP") = "Ã≈" Then
'970             Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m30 + Space(7), 20) + mXPI_P + Replace(Format(aji - M_ELGA, "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'
'980             Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m38 + Space(7), 20) + mXPI_X + Replace(Format(aji - M_ELGA, "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'            End If
'
'
'
'            If R("TRP") = "– " Then
'            'MPK ANTI M38
'                Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m30 + Space(7), 20) + mXPI_P + Replace(Format(aji, "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'
'                Print #1, "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(R("hme"), "DD/MM/YYYY") + Left(to437(mait) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(mPK + Space(7), 20) + mXPI_X + Replace(Format(aji, "00000000000.00"), ".", ",") + Left(MATIM + Space(20), 20) + Left(to437(mAit2) + Space(60), 60)
'            End If
'
'         Rp.Close
'
'
'
'
'
'
'
'
'1000        Gdb.Execute "UPDATE TIM SET B_C1= '*'+convert(CHAR(10),GETDATE(),3) WHERE ID_NUM=" + str(R!ID_NUM)
'1010        N = N + 1
'1020        R.MoveNext
'        Loop
'
'        '=========================================================================================================================================
'
'1030    R.Close
'
'        Dim GEF As Integer  ' 1=OK GEFYRONO    0=DEN GEFYRONO
'
'        Dim sql As String
'
'        '1250  sql = _
'        '          "SELECT EGG.*,PEL.EPO FROM EGG INNER JOIN PEL ON EGG.EIDOS=PEL.EIDOS AND EGG.KOD=PEL.KOD WHERE (ATIM IS NULL OR ATIM='' OR (ISNUMERIC(ATIM)=1)) AND " + synt2 + " HME>='" _
'        '          + Format(apo, "MM/DD/YYYY") + "' AND HME<='" + Format(DateAdd("D", 1, eos), _
'        '          "MM/DD/YYYY") + "' "
'1040    sql = "SELECT EGG.*,PEL.EPO,PEL.AFM FROM EGG INNER JOIN PEL ON EGG.EIDOS=PEL.EIDOS AND EGG.KOD=PEL.KOD WHERE AAXREOPIS>0 AND " + synt2 + " HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' "
'
'        '      '===================  EGG ===============================================
'1050    R.Open sql, Gdb, adOpenDynamic, adLockOptimistic
'1060    mKPE = "0004450"
'
'1070    mXPI_X = "0" ' ˜ÒÂ˘ÛÁ
'1080    mXPI_P = "1" ' –…”‘Ÿ”«
'
'        Dim R23         As New ADODB.Recordset
'
'        Dim SYMCIFISMOS As Integer
'
'1090    Do While Not R.EOF
'            GEF = 1 ' 1=OK GEFYRONO    0=DEN GEFYRONO
'            SYMCIFISMOS = 0
'            MAFM = R("afm")
'            mEpo = mID(R("EPO"), 1, 30)
'
'            If IsNull(R("AFM")) Then
'                mKPE = "0000000"
'            Else
'
'                If Len(R("AFM")) < 9 Then
'                    mKPE = "0000000"
'                Else
'                    mKPE = mID(R("AFM"), 3, 7)
'                End If
'            End If
'
'            GoSub ADDSYNAL
'
'1100        R23.Open "SELECT KODLOG,EPIT FROM XREOPIS WHERE EIDOS=" + str(R!AAXREOPIS), Gdb, adOpenDynamic, adLockOptimistic
'
'1110        If IsNull(R23!KODLOG) Then
'1120            m38 = "38-00-00-0000"
'            Else
'1130            m38 = Trim(R23!KODLOG)
'            End If
'
'            If IsNull(R23!epit) Then
'                SYMCIFISMOS = 0
'            Else
'
'                If R23!epit = 2 Then
'                    SYMCIFISMOS = 1
'                End If
'            End If
'
'1140        R23.Close
'
'1150        mait = R("epo")
'
'1160        hme = R("hme")
'1170        mAit2 = R!AIT
'
'1180        If R("EIDOS") = "e" Then
'1190            m30 = "30-00-00-0000"
'1200            mPOL = "1300"
'
'                If IsNull(R!m30) Then  ' ¡Õ ‘¡ ≈√—¡ÿ≈ ÃœÕœ” ‘œ’ œ –≈À¡‘«” ‘œ’” Àœ√¡—…¡”Ãœ’”  —¡‘«”≈ ‘œ’”
'                Else
'                    m30 = R!m30
'                End If
'
'                If IsNull(R!m38) Then  ' ¡Õ ‘¡ ≈√—¡ÿ≈ ÃœÕœ” ‘œ’ œ –≈À¡‘«” ‘œ’” Àœ√¡—…¡”Ãœ’”  —¡‘«”≈ ‘œ’”
'                Else
'                    m38 = R!m38
'                End If
'
'                If Left(m38, 2) = "00" And Left(m30, 2) = "00" Then   'ƒ≈Õ √≈÷’—ŸÕŸ
'                    GEF = 0
'                End If
'
'                If SYMCIFISMOS = 1 Then
'                    mPOL = "1100"
'                End If
'
'                '1370              If InStr(mAit2, "À«Œ") > 0 Then
'                '1380                 m38 = "33-90-00-0000"
'                '1390              Else
'                '1400                 m38 = "38-00-00-0000"
'                '1410              End If
'                '
'
'                '≈…”–—¡Œ« ¡–œ –≈À¡‘«  – 30  X 38
'1210            If R("PISTOSI") > 0 And GEF = 1 Then
'1220                MATIM = to437("¡–.≈…”." + Format(R!APA, "#####0"))
'1230                ajia = R("pistosi")
'
'1240                Print #1, prints(N, mKPE, MATIM, hme, mait, 0, m30, "1", ajia, ATIM, mAit2, mPOL)
'1250                Print #1, prints(N, mKPE, MATIM, hme, mait, 0, m38, "0", ajia, ATIM, mAit2, mPOL)
'
'                End If
'
'                '–À«—ŸÃ« –≈À¡‘«
'1260            If R("XREOSI") > 0 And GEF = 1 Then
'1270                MATIM = to437("¡–.ÎÁÒ." + Format(R!APA, "#####0"))
'1280                ajia = R("xreosi")
'1290                Print #1, prints(N, mKPE, MATIM, hme, mait, 0, m30, "0", ajia, ATIM, mAit2, mPOL)
'1300                Print #1, prints(N, mKPE, MATIM, hme, mait, 0, m38, "1", ajia, ATIM, mAit2, mPOL)
'                End If
'
'            Else  ' agores
'1310            mPOL = "1200"
'
'                If SYMCIFISMOS = 1 Then mPOL = "1100"
'                '1560              If InStr(mAit2, "À«Œ") > 0 Then
'                '1570                 m38 = "53-90-00-0000"
'                '1580              Else
'                '1590                 m38 = "38-00-00-0000"
'                '1600              End If
'
'                '–—œÃ«»≈’‘≈”
'1320            m30 = "50-00-00-0000"
'
'                If IsNull(R!m30) Then  ' ¡Õ ‘¡ ≈√—¡ÿ≈ ÃœÕœ” ‘œ’ œ –≈À¡‘«” ‘œ’” Àœ√¡—…¡”Ãœ’”  —¡‘«”≈ ‘œ’”
'                Else
'                    m30 = R!m30
'                End If
'
'                If IsNull(R!m38) Then  ' ¡Õ ‘¡ ≈√—¡ÿ≈ ÃœÕœ” ‘œ’ œ –≈À¡‘«” ‘œ’” Àœ√¡—…¡”Ãœ’”  —¡‘«”≈ ‘œ’”
'                Else
'                    m38 = R!m38
'                End If
'
'                If Left(m38, 2) = "00" And Left(m30, 2) = "00" Then   'ƒ≈Õ √≈÷’—ŸÕŸ
'                    GEF = 0
'                End If
'
'                '≈…”–—¡Œ« ¡–œ –—œÃ«»≈’‘«  ◊ 38  – 50
'1330            If R("PISTOSI") > 0 And GEF = 1 Then
'1340                MATIM = to437("¡–.≈…”." + Format(R!APA, "#####0"))
'1350                ajia = R("pistosi")
'1360                Print #1, prints(N, mKPE, MATIM, hme, mait, 0, m30, "1", ajia, ATIM, mAit2, mPOL)
'1370                Print #1, prints(N, mKPE, MATIM, hme, mait, 0, m38, "0", ajia, ATIM, mAit2, mPOL)
'                End If
'
'                '–À«—ŸÃ« –—œÃ«»≈’‘«
'1380            If R("XREOSI") > 0 And GEF = 1 Then
'1390                ajia = R("xreosi")
'1400                MATIM = to437("¡.–ÎÁÒ." + Format(R!APA, "#####0"))
'1410                Print #1, prints(N, mKPE, MATIM, hme, mait, 0, m30, "0", ajia, ATIM, mAit2, mPOL)
'1420                Print #1, prints(N, mKPE, MATIM, hme, mait, 0, m38, "1", ajia, ATIM, mAit2, mPOL)
'                End If
'            End If
'
'1430        Gdb.Execute "UPDATE EGG SET ATIM2 = '*' WHERE ID=" + str(R!ID) '+convert(CHAR(10),GETDATE(),3)
'1440           N = N + 1
'1450        R.MoveNext
'
'        Loop
'
'        Close #1
'
'        FileCopy mfil, "C:\LAG\KINHS.TXT"
'
'        mfil = Replace(mfil, "KINHS", "SYNAL")
'
'        Open mfil For Output As #1
'
'        For k = 1 To nSYNAL
'            Print #1, to437(Msynal(k))
'        Next
'
'        Close #1
'
'        FileCopy mfil, "C:\LAG\SYNAL.TXT"
'
'1490    MsgBox "ƒ«Ã…œ’—«»« ≈ ‘œ ¡—◊≈…œ " + mfil
'
'        Exit Sub
'
'SHOW_ERROR:
'1500    MsgBox str(Erl) + "**" + Err.Description
'
'1510    Resume Next
'
'        '<EhFooter>
'        Exit Sub
'
'        Exit Sub
'
'ADDSYNAL:
'
'        '”’Õ¡ÀÀ¡””œÃ≈Õœ”
'        synal = "2" + mKPE + mID(mEpo + Space(60), 1, 60) ' 2+ ID SYNALLASOMENOY +epo=60
'        synal = synal + mID(MAFM + Space(9), 1, 9) ' afm=9
'        synal = synal + mID(Space(40), 1, 40)  ' epa=40
'        synal = synal + "    1" + mID(Space(60), 1, 60)  ' kvdikos DOY=4, ¡»≈”‘ŸÚ ˆ·=1 , die=60
'        synal = synal + Space(10) 'ÙÍ
'        synal = synal + mID(Space(40), 1, 40)  'POLH
'        synal = synal + mID(Space(40), 1, 40)  'THL1
'        synal = synal + mID(Space(40), 1, 40)  'THL2
'        synal = synal + mID(Space(40), 1, 40)  'THL3
'        synal = synal + mID(Space(40), 1, 40)  'FAX1
'        synal = synal + mID(Space(40), 1, 40)  'FAX2
'        synal = synal + mID(Space(40), 1, 40)  'EMAIL
'        synal = synal + "1000"  '
'        synal = synal + mID(Space(60), 1, 60)  'ƒ…≈’».Àœ√…”‘«—…œ’
'        synal = synal + mID(Space(60), 1, 60)  'ƒ…≈’».¡–œ»« «”
'        synal = synal + mID(Space(60), 1, 60)  '’–≈’».¡–œ»« «”
'        synal = synal + mID(Space(60), 1, 60)  '‘—¡–≈∆¡ 1
'        synal = synal + mID(Space(60), 1, 60)  'Àœ√.‘—¡–≈∆«” 1
'        synal = synal + mID(Space(60), 1, 60)  '‘—¡–≈∆¡ 2
'        synal = synal + mID(Space(60), 1, 60)  'Àœ√.‘—¡–≈∆ 2
'        synal = synal + mID(Space(60), 1, 40)  '◊Ÿ—¡ ≈Õƒœ œ…Õœ‘… ŸÕ
'        synal = synal + mID(Space(2), 1, 2)  '–—œ».◊Ÿ—¡ ≈Õƒœ œ…Õœ‘… ŸÕ
'        synal = synal + mID(Space(12), 1, 12)  '¡—.Ã«‘—.÷–¡
'
'        synal = synal + mID("AT" + Space(12), 1, 10) '≈…ƒœ” ‘¡’‘œ‘«‘¡”
'        synal = synal + mID(Space(20), 1, 20)  '¡—.‘¡’‘œ‘«‘¡”
'
'        OK = 0
'
'        For i = 1 To nSYNAL
'
'            If Left(synal, 8) = Left(Msynal(i), 8) Then
'                OK = 1
'            End If
'
'        Next
'
'        If OK = 0 Then
'            nSYNAL = nSYNAL + 1
'            Msynal(nSYNAL) = synal
'        End If
'
'        Return
'
'Command6_Click_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.apot7.Command6_Click " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.Command6_Click " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>

End Sub

Function prints(ByVal N As Long, _
                ByVal mKPE As String, _
                ByVal MATIM As String, _
                ByVal hme As Date, _
                ByVal Ait16 As String, _
                ByVal AJIA_KEPYO As Single, _
                ByVal log As String, _
                ByVal xpi As String, _
                ByVal ajia As Single, _
                ByVal ATIM As String, _
                ByVal mait As String, _
                ByVal mPOL As String)

        '<EhHeader>
        On Error GoTo prints_Err

        '</EhHeader>
    
100     prints = "4" + Format(N, "00000") + mPOL + mKPE + Left(MATIM + Space(18), 18) + Format(hme, "DD/MM/YYYY") + Left(to437(Ait16) + Space(60), 60) + Replace(Format(AJIA_KEPYO, "00000000000.00"), ".", ",") + Left(log + Space(7), 20) + xpi + Replace(Format(ajia, "00000000000.00"), ".", ",") + Left(ATIM + Space(20), 20) + Left(to437(mait) + Space(60), 60)

        '<EhFooter>
        Exit Function

prints_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.prints " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.prints " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function Get_GEF()

        '<EhHeader>
        On Error GoTo Get_GEF_Err

        '</EhHeader>

        Dim DB  As Database

        Dim R   As New ADODB.Recordset, X As String

        Dim GEF As String
    
100     R.Open "select XREPEL,PISPEL,EIDOS from PARASTAT", Gdb, adOpenDynamic, adLockOptimistic

        Dim pol As String

110     GEF = ""
120     R.MoveFirst

130     Do While Not R.EOF

140         If IsNull(R("XREPEL")) And IsNull(R("PISPEL")) Then
                ' ADIAFORO
            Else

150             If IsNull(R("XREPEL")) Then
160                 If Len(Trim(R("PISPEL"))) > 4 Then
170                     GEF = GEF + "'" + R("eidos") + "',"
                    End If
                End If

180             If IsNull(R("PISPEL")) Then
190                 If Len(Trim(R("XREPEL"))) > 4 Then
200                     GEF = GEF + "'" + R("eidos") + "',"
                    End If
                End If

210             If (Not IsNull(R("PISPEL"))) And (Not IsNull(R("XREPEL"))) Then
220                 If Len(Trim(R("PISPEL"))) > 4 Then
230                     GEF = GEF + "'" + R("eidos") + "',"
                    Else

240                     If Len(Trim(R("XREPEL"))) > 4 Then
250                         GEF = GEF + "'" + R("eidos") + "',"
                        End If
                    End If
                End If
            End If

260         R.MoveNext
        Loop

        On Error Resume Next

270     If Len(GEF) = 0 Then
280         Get_GEF = ""
        Else
290         Get_GEF = Left(GEF, Len(GEF) - 1)
        End If

        '<EhFooter>
        Exit Function

Get_GEF_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.Get_GEF " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.Get_GEF " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub CMDRETAIL_Click()
Dim DUM

   Dim sql As String
   Dim mfile2 As String
   Dim mfile As String
   Dim a As String
   
   '---------------------------------HEADER  ----------------------------------
   Dim mpath As String
   
    mfile = "c:\mercvb\QUERIES\ret1path.txt"
    Dim mAA As String  ' mAA epipleon sto onoma na katalabainei an einai to proto h to deytero zeygari
    
    mAA = ""
   If Len(Dir(mfile, vbNormal)) > 0 Then
     Open mfile For Input As #1
       Do While Not EOF(1)
         Line Input #1, a
         mpath = mpath + a
         If Not EOF(1) Then
            Line Input #1, mAA
         End If
       Loop
     Close #1
    Else
    
        MsgBox "œÒﬂÛÙÂ ÙÁÌ ‰È·‰ÒÔÏﬁ Ù˘Ì ·Ò·„¸ÏÂÌ˘Ì ·Ò˜Âﬂ˘Ì ÛÙÔ " + mfile
    
        MsgBox "–ÒÔÛ˘ÒÈÌ‹ ·ÔËÁÍÂ˝ÔÌÙ·È ÛÙÔ  c:\mercvb\QUERIES\"
         mpath = "c:\mercvb\"
    
    End If
    
   
   
  'Gdb.Execute " UPDATE TIM SET AJ1= ( SELECT SUM(ROUND(POSO*TIMM,2) )  FROM EGGTIM WHERE  HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<='" + Format(eos, "MM/DD/YYYY") + "' AND FPA=1 AND  ID_NUM=TIM.ID_NUM) "
   
  'Gdb.Execute " UPDATE TIM SET AJ2= ( SELECT SUM(ROUND(POSO*TIMM,2) )  FROM EGGTIM WHERE  HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<='" + Format(eos, "MM/DD/YYYY") + "' AND FPA=2 AND  ID_NUM=TIM.ID_NUM) "
   
   
   sql = ""
   
  
    mfile = "c:\mercvb\QUERIES\ret1"
    Open mfile For Input As #1
     
       Do While Not EOF(1)
          Line Input #1, a
          sql = sql + a
       Loop

    Close #1
   
   sql = sql + " AND HME>='" + Format(APO, "MM/DD/YYYY") + "' AND HME<='" + Format(EOS, "MM/DD/YYYY") + "'"
   
   
   
   mfile2 = mpath + "INV_HEADER_" + Trim(mAA) + Format(Now, "YYYYMMDD") + ".CSV"

Text2.Text = Text2.Text + mfile2 + Chr(13)


    PRINT_CSV sql, "0000000000000", "", 0, 0, mfile2
   
'---------------------------------LINES ----------------------------------
 sql = ""
   
   mfile = "c:\mercvb\QUERIES\ret2"
   Open mfile For Input As #1
     
      Do While Not EOF(1)
        Line Input #1, a
        sql = sql + a
      Loop

   Close #1
   
    sql = sql + " AND HME>='" + Format(APO, "MM/DD/YYYY") + "' AND HME<='" + Format(EOS, "MM/DD/YYYY") + "'"
   
   
   
   mfile2 = mpath + "INV_LINES_" + Trim(mAA) + Format(Now, "YYYYMMDD") + ".CSV"

Text2.Text = Text2.Text + mfile2 + Chr(13)
    
    
    PRINT_CSV sql, "0000000000000", "", 0, 0, mfile2

End Sub



'Private Sub TSENTIDIS_Click()
'Dim DUM
'
'   Dim sql As String
'   Dim mfile2 As String
'   Dim mfile As String
'   Dim a As String
'
'   '---------------------------------HEADER  ----------------------------------
'   Dim mpath As String
'
'    mfile = "c:\mercvb\QUERIES\ret1path.txt"
'    Dim mAA As String  ' mAA epipleon sto onoma na katalabainei an einai to proto h to deytero zeygari
'
'    mAA = ""
'   If Len(Dir(mfile, vbNormal)) > 0 Then
'     Open mfile For Input As #1
'       Do While Not EOF(1)
'         Line Input #1, a
'         mpath = mpath + a
'         If Not EOF(1) Then
'            Line Input #1, mAA
'         End If
'       Loop
'     Close #1
'    Else
'
'        MsgBox "œÒﬂÛÙÂ ÙÁÌ ‰È·‰ÒÔÏﬁ Ù˘Ì ·Ò·„¸ÏÂÌ˘Ì ·Ò˜Âﬂ˘Ì ÛÙÔ " + mfile
'
'        MsgBox "–ÒÔÛ˘ÒÈÌ‹ ·ÔËÁÍÂ˝ÔÌÙ·È ÛÙÔ  c:\mercvb\QUERIES\"
'         mpath = "c:\mercvb\"
'
'    End If
'
'
'
'  'Gdb.Execute " UPDATE TIM SET AJ1= ( SELECT SUM(ROUND(POSO*TIMM,2) )  FROM EGGTIM WHERE  HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<='" + Format(eos, "MM/DD/YYYY") + "' AND FPA=1 AND  ID_NUM=TIM.ID_NUM) "
'
'  'Gdb.Execute " UPDATE TIM SET AJ2= ( SELECT SUM(ROUND(POSO*TIMM,2) )  FROM EGGTIM WHERE  HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<='" + Format(eos, "MM/DD/YYYY") + "' AND FPA=2 AND  ID_NUM=TIM.ID_NUM) "
'
'
'   sql = ""
'
'
'    mfile = "c:\mercvb\QUERIES\ret1"
'    Open mfile For Input As #1
'
'       Do While Not EOF(1)
'          Line Input #1, a
'          sql = sql + a
'       Loop
'
'    Close #1
'
'   sql = sql + " AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<='" + Format(eos, "MM/DD/YYYY") + "'"
'
'
'
'   mfile2 = mpath + "INV_HEADER_" + Trim(mAA) + Format(Now, "YYYYMMDD") + ".CSV"
'
'Text2.Text = Text2.Text + mfile2 + Chr(13)
'
'
'    PRINT_CSV sql, "0000000000000", "", 0, 0, mfile2
'
''---------------------------------LINES ----------------------------------
' sql = ""
'
'   mfile = "c:\mercvb\QUERIES\ret2"
'   Open mfile For Input As #1
'
'      Do While Not EOF(1)
'        Line Input #1, a
'        sql = sql + a
'      Loop
'
'   Close #1
'
'    sql = sql + " AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<='" + Format(eos, "MM/DD/YYYY") + "'"
'
'
'
'   mfile2 = mpath + "INV_LINES_" + Trim(mAA) + Format(Now, "YYYYMMDD") + ".CSV"
'
'Text2.Text = Text2.Text + mfile2 + Chr(13)
'
'
'    PRINT_CSV sql, "0000000000000", "", 0, 0, mfile2
''Shell ("c:\windows\NCFTPPUT.exe  -u toyshopftp -p  Vykq819*  176.9.148.246 /httpdocs/images/ " + F_FAKEL_KOINOS + "\images\" + PIC(k).Caption)
'
'End Sub













Private Sub cmdPreview_Click()

 With TDBGrid.PrintInfo
        ' Set the page header
        .PageHeaderFont.Italic = True
        .PageHeader = "Composers table"
        
        ' Column headers will be on every page
         .RepeatColumnHeaders = True
        
        ' Display page numbers (centered)
        .PageFooter = "\tPage: \p"
        ' Invoke Print Preview
        .PrintPreview
    End With





End Sub

' „ÂˆıÒ· Ôı ‰ÔıÎÂıÂÈ ÏÂ ˘ÎÁÛÂÈÚ ÔÍ  29-1-2013

'Private Sub Command6_Click()
''√≈÷’—¡ ”≈ EPSILON
'Dim R As New ADODB.Recordset
'Dim Rp As New ADODB.Recordset
'Dim ait23, ait13, ait09
'Open "C:\SSS" For Output As #1
'Dim m70, m54, m30, m38
'
'ss = "'T','t'"
'
'ss = InputBox(" Ÿƒ… œ… –ŸÀ«”≈ŸÕ", , ss)
'R.Open "SELECT * FROM TIM WHERE HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<='" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' AND LEFT(ATIM,1) in (" + ss + ")", Gdb, adOpenDynamic, adLockOptimistic
'  mKPE = "0004450"
'  ait23 = "–ŸÀ«”≈…” ≈Ã–œ—≈’Ã¡‘ŸÕ 23%"
'  ait13 = "–ŸÀ«”≈…” ≈Ã–œ—≈’Ã¡‘ŸÕ 13%"
'  ait09 = "–ŸÀ«”≈…” ≈Ã–œ—≈’Ã¡‘ŸÕ 9%"
'  mXPI_X = "0" ' ˜ÒÂ˘ÛÁ
'  mXPI_P = "1" ' –…”‘Ÿ”«
'  m30 = "30-00-00-0000"
'  m38 = "38-00-00-0000"
'  Dim aji
'
'
'N = 1
'Do While Not R.EOF
'  aji = R!aj1 + R!aj2 + R!aj3 + R!aj4 + R!aj5 + R!fpa1 + R!fpa2 + R!fpa3 + R!fpa4
'  Rp.Open "SELECT * FROM PARASTAT WHERE EIDOS='" + Left(R("ATIM"), 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
'
'
'  If R("FPA1") > 0 Then
'     If IsNull(Rp("LAJ1")) Or IsNull(Rp("LFPA1")) Then
'       MsgBox "ƒ≈Õ ’–¡—◊≈… « –¡—¡Ã≈‘—œ” √…¡ ‘œ –¡—¡”‘¡‘… œ " + Rp("TITLOS") + " √…¡ ‘œ ÷–¡ 13%"
'     Else
'       m70 = Rp("LAJ1")
'       m54 = Rp("LFPA1")
'       Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait16) + Space(60), 60) + Replace(Format(R("aj1"), "00000000000.00"), ".", ",") + Left(m70 + Space(7), 20) + mXPI_P + Replace(Format(R("aj1"), "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'       Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait16) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m54 + Space(7), 20) + mXPI_P + Replace(Format(R("FPA1"), "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'     End If
'  End If
'
'  If R("FPA2") > 0 Then
'     If IsNull(Rp("LAJ2")) Or IsNull(Rp("LFPA2")) Then
'       MsgBox "ƒ≈Õ ’–¡—◊≈… « –¡—¡Ã≈‘—œ” √…¡ ‘œ –¡—¡”‘¡‘… œ " + Rp("TITLOS") + " √…¡ ‘œ ÷–¡ 23%"
'     Else
'       m70 = Rp("LAJ2")
'       m54 = Rp("LFPA2")
'
'       Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait23) + Space(60), 60) + Replace(Format(R("aj2"), "00000000000.00"), ".", ",") + Left(m70 + Space(7), 20) + mXPI_P + Replace(Format(R("aj2"), "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'       Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait23) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m54 + Space(7), 20) + mXPI_P + Replace(Format(R("FPA2"), "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'     End If
'
'  End If
'
'  If R("FPA3") > 0 Then
'     If IsNull(Rp("LAJ3")) Or IsNull(Rp("LFPA3")) Then
'       MsgBox "ƒ≈Õ ’–¡—◊≈… « –¡—¡Ã≈‘—œ” √…¡ ‘œ –¡—¡”‘¡‘… œ " + Rp("TITLOS") + " √…¡ ‘œ ÷–¡ 19%"
'     Else
'       m70 = Rp("LAJ3")
'       m54 = Rp("LFPA3")
'       Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait16) + Space(60), 60) + Replace(Format(R("aj3"), "00000000000.00"), ".", ",") + Left(m70 + Space(7), 20) + mXPI_P + Replace(Format(R("aj3"), "00000000000.00"), ".", ",") + Left(Atim + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'       Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait16) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m54 + Space(7), 20) + mXPI_P + Replace(Format(R("FPA3"), "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'     End If
'
'  End If
'
'  If R("FPA4") > 0 Then
'     If IsNull(Rp("LAJ4")) Or IsNull(Rp("LFPA4")) Then
'       MsgBox "ƒ≈Õ ’–¡—◊≈… « –¡—¡Ã≈‘—œ” √…¡ ‘œ –¡—¡”‘¡‘… œ " + Rp("TITLOS") + " √…¡ ‘œ ÷–¡ 9%"
'     Else
'       m70 = Rp("LAJ4")
'       m54 = Rp("LFPA4")
'       Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait16) + Space(60), 60) + Replace(Format(R("aj4"), "00000000000.00"), ".", ",") + Left(m70 + Space(7), 20) + mXPI_P + Replace(Format(R("aj4"), "00000000000.00"), ".", ",") + Left(Atim + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'       Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait16) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m54 + Space(7), 20) + mXPI_P + Replace(Format(R("FPA4"), "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'     End If
'
'  End If
'
'  If R("AJ5") > 0 Then
'     If IsNull(Rp("LAJ5")) Then
'       MsgBox "ƒ≈Õ ’–¡—◊≈… « –¡—¡Ã≈‘—œ” √…¡ ‘œ –¡—¡”‘¡‘… œ " + Rp("TITLOS") + " √…¡ ‘œ ÷–¡ 0%"
'     Else
'        m70 = Rp("AJ5")
'        Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait23) + Space(60), 60) + Replace(Format(R("aj5"), "00000000000.00"), ".", ",") + Left(m70 + Space(7), 20) + mXPI_P + Replace(Format(R("aj5"), "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'     End If
'  End If
'
'  '30
'  Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait23) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m30 + Space(7), 20) + mXPI_X + Replace(Format(aji, "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'
'  'METRHTA
'  If R("TRP") = "Ã≈" Then
'     Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait23) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m30 + Space(7), 20) + mXPI_P + Replace(Format(aji, "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'     Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait23) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m38 + Space(7), 20) + mXPI_X + Replace(Format(aji, "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'  End If
'  Rp.Close
'
'
'  N = N + 1
'  R.MoveNext
'Loop
'
'
'  Print #1, to437("20004450–≈À¡‘«” ◊œÕƒ—… «”                                           000000000...............                         42311BEROIAS 21                                                  53423     UESSALONIKH                             2310451230                              2310461238                              2310461230                              2310461122                              2310461123                              aitsios otenet.gr                       1000PAPAGOY 233                                                 KARABAGGELH 4                                               GEVRGIOS MPOYTOS                                            EUROBANK                                                    7878788787844                                               EUNIKH                                                      45562548154487                                                                                                    AT        AZ976431")
'  Print #1, to437("20004423–≈À¡‘«” À…¡Õ… «”                                            000000001.........................               42111PATROKLOY 33                                                52366     UESSALONIKH                             2310963320                              2310963321                              2310963322                              2310963323                              2310963324                              ppolyxroniadis yahoo.gr                 1000ERATYRAS 5                                                  PAYLOY MELA 287                                             BASILEIOS DROYTSAS                                          AGROTIKH                                                    8745645424545                                               PROBANK                                                     74545121545454                                                                                                    AT         P802814")
'
'
'Close #1
'
'
'
'
'MsgBox "ƒ«Ã…œ’—«»« ≈ ‘œ ¡—◊≈…œ C:\SSS"
'End Sub

'Sub gef_eps(ByVal Atim As String, ByVal mKPE As String, ait23 As String, ait16 As String, m70 As String)
'
'
'
''√≈÷’—¡ ”≈ EPSILON
'Dim R As New ADODB.Recordset
'R.Open "SELECT * FROM TIM WHERE HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<='" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' AND LEFT(ATIM,1)='T'", Gdb, adOpenDynamic, adLockOptimistic
''mKPE = mKPE
'
'Open "C:\SSS" For Output As #1
'
'N = 0
'Do While Not R.EOF
'  N = N + 1
'  '                                                                                                                                                                                       AJIA KEPYO
'
'If R("AJ5") > 0 Then
'  Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait23) + Space(60), 60) + Replace(Format(R("aj2"), "00000000000.00"), ".", ",") + Left(m70 + Space(7), 20) + "1" + Replace(Format(R("aj2"), "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'  Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait23) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m547057 + Space(7), 20) + "1" + Replace(Format(R("FPA2"), "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'End If
'
'
'
'
'If R("FPA2") > 0 Then
'  Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait23) + Space(60), 60) + Replace(Format(R("aj2"), "00000000000.00"), ".", ",") + Left(m70 + Space(7), 20) + "1" + Replace(Format(R("aj2"), "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'  Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait23) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m547057 + Space(7), 20) + "1" + Replace(Format(R("FPA2"), "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'End If
'
'
'If R("FPA3") > 0 Then
'  Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait16) + Space(60), 60) + Replace(Format(R("aj1"), "00000000000.00"), ".", ",") + Left(m700054 + Space(7), 20) + "1" + Replace(Format(R("aj2"), "00000000000.00"), ".", ",") + Left(Atim + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'  Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait16) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m547054 + Space(7), 20) + "1" + Replace(Format(R("FPA2"), "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'End If
'
'
'
'  Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait23) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m30 + Space(7), 20) + "0" + Replace(Format(R("ajI"), "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'
'
'
'
'If R("TRP") = "Ã≈" Then
'  Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait23) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m30 + Space(7), 20) + "1" + Replace(Format(R("aji"), "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'  Print #1, "4" + Format(N, "00000") + "1310" + mKPE + Left(R("Atim") + Space(20), 20) + Format(R("hme"), "DDMMYYYY") + Left(to437(ait23) + Space(60), 60) + Replace(Format(0, "00000000000.00"), ".", ",") + Left(m38 + Space(7), 20) + "0" + Replace(Format(R("aji"), "00000000000.00"), ".", ",") + Left(R("Atim") + Space(20), 20) + Left(to437(ait23) + Space(60), 60)
'End If
'
'
'
'
'
'
'R.MoveNext
'Loop
'
'
'
'
'
'
'
'
'  Print #1, "2mKPE–≈À¡‘«” ◊œÕƒ—… «”                                           000000000...............                         42311BEROIAS 21                                                  53423     UESSALONIKH                             2310451230                              2310461238                              2310461230                              2310461122                              2310461123                              aitsios otenet.gr                       1000PAPAGOY 233                                                 KARABAGGELH 4                                               GEVRGIOS MPOYTOS                                            EUROBANK                                                    7878788787844                                               EUNIKH                                                      45562548154487                                                                                                    AT        AZ976431"
'  Print #1, "20004423–≈À¡‘«” À…¡Õ… «”                                            000000001.........................               42111PATROKLOY 33                                                52366     UESSALONIKH                             2310963320                              2310963321                              2310963322                              2310963323                              2310963324                              ppolyxroniadis yahoo.gr                 1000ERATYRAS 5                                                  PAYLOY MELA 287                                             BASILEIOS DROYTSAS                                          AGROTIKH                                                    8745645424545                                               PROBANK                                                     74545121545454                                                                                                    AT         P802814"
'
'
'Close #1
''4000011220mKPETDA9899             04112012AGORA EMPOREYMATVN ME 19%                                   00000003445,3820-00-00-0077       000000003445,38TDA9899             AGORA EMPOREYMATVN ME 19%
''4000011220mKPETDA9899             04112012AGORA EMPOREYMATVN ME 19%                                   0000000000000054-00-20-0077       000000002554,62TDA9899             AGORA EMPOREYMATVN ME 19%
''4000011220mKPETDA9899             04112012AGORA EMPOREYMATVN ME 19%                                   0000000000000050-00-00-0088       100000006000,00TDA9899             AGORA EMPOREYMATVN ME 19%
''40000213100004423TDA1220             04112012PVLHSH EMPOREYMATVN ME 19%                                  00000008487,3970-00-00-0077       100000008487,39TDA1220             PVLHSH EMPORETMATVN ME 19%
''40000213100004423TDA1220             04112012PVLHSH EMPOREYMATVN ME 19%                                  0000000000000054-00-70-0077       100000003512,61TDA1220             PVLHSH EMPOREYMATVN ME 19%
''40000213100004423TDA1220             04112012PVLHSH EMPOREYMATVN ME 19%                                  0000000000000030-00-00-0021       000000012000,00TDA1220             PVLHSH EMPOREYMATVN ME 19%
'MsgBox "ƒ«Ã…œ’—«»« ≈ ‘œ ¡—◊≈…œ C:\SSS"
'End Sub

Private Sub Dir1_Change()

        '<EhHeader>
        On Error GoTo Dir1_Change_Err

        '</EhHeader>

100     File1.Path = Dir1.Path
110     File1.Refresh

        '<EhFooter>
        Exit Sub

Dir1_Change_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.Dir1_Change " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.Dir1_Change " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Dir1_Click()

        '<EhHeader>
        On Error GoTo Dir1_Click_Err

        '</EhHeader>

100     File1.Path = Dir1.Path
110     File1.Refresh

        '<EhFooter>
        Exit Sub

Dir1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.Dir1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.Dir1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub File1_Click()
'shmadeyv to report
        '<EhHeader>
        On Error GoTo File1_Click_Err

        '</EhHeader>
        Dim a

100     a = 0

105     Dim N

106     Dim txt(100) As String

        Dim F        As String

110     F = File1.FILENAME

        Dim SS As String

        Dim B  As String

120     SS = ""
125     N = 0
130     Open File1.Path + "\" + F For Input As #1

140     Do While Not EOF(1)

150         Line Input #1, B
155         N = N + 1: txt(N) = B
160         SS = SS + B + Chr(13)
        Loop

170     Close #1
        'label3=par1
        'label4=par2
        'label5=par3
       
        If InStr(txt(1), "”’ÕœÀ¡") > 0 And InStr(txt(4), "/*") > 0 Then
            sumstring.Tag = Trim(mID(txt(5), InStr(txt(5), "”’ÕœÀ¡") + 1, 20))
        End If
       
       
       
       
       
        If InStr(txt(1), "#") > 0 And InStr(txt(1), "/*") > 0 Then
            Label3.Caption = mID(txt(1), InStr(txt(1), "#") + 1, 15)
        End If

        If InStr(txt(2), "#") > 0 And InStr(txt(2), "/*") > 0 Then
            Label4.Caption = mID(txt(2), InStr(txt(2), "#") + 1, 15)
        End If

        If InStr(txt(3), "#") > 0 And InStr(txt(3), "/*") > 0 Then
            Label5.Caption = mID(txt(3), InStr(txt(3), "#") + 1, 15)
        End If
       
        If InStr(txt(4), "#") > 0 And (InStr(UCase(txt(4)), "ID") > 0 Or InStr(UCase(txt(4)), "DSN") > 0) Then
            Dim LA As Integer, FI As Integer
            FI = InStr(txt(4), "#")
            LA = InStr(10, txt(4), "#")
            FDSN = Trim(mID(txt(4), InStr(txt(4), "#") + 1, LA - FI - 1))
            Me.Caption = FDSN
        End If
       
        If InStr(txt(5), "#") > 0 And InStr(txt(4), "/*") > 0 Then
            sumstring.Text = Trim(mID(txt(5), InStr(txt(5), "#") + 1, 20))
        End If
       
       
       
       
       
180     Text2.Text = SS

        ' ·ÎÎ·Ê˘ ÙÔÌ ÙÈÙÎÔ "–·Ò·ÏÂÙÒÔÚ 1,2,3  Í·È 4Á ÛÂÈÒ‹ ·ÎÎ·„ﬁ dsn
        ' /*  #BARCODE   */
        '/*  #ONOMA      */
        '/*  #THLEFONO   */
        '/*  #DSN=MERURY;uid=sa;pwd=p@ssw0rd  /
        '
        'update EID SET ONO='@c2' ,ERG='@c3'  where KOD='@c1'
        '

        '<EhFooter>
        Exit Sub

File1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.File1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.File1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>
 PARAMETROI.PARAM.Caption = "APOT7"
100     PARAMETROI.SHOW 1

        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.Form_DblClick " & "at line " & Erl

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
         "in ADOMERCNEW.apot7.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 15240, 10125, fh, fw, ft, fl
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
         "in ADOMERCNEW.apot7.TDBGrid_FilterChange " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.TDBGrid_FilterChange " & "at line " & Erl

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
         "in ADOMERCNEW.apot7.getFilter " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.getFilter " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Sub CMDPROVOLI_Click()


        '<EhHeader>
        On Error GoTo CMDPROVOLI_Click_Err

        '</EhHeader>
        Dim DBF    As New ADODB.Connection

        Dim sql    As New ADODB.Connection

        Dim rDBF   As New ADODB.Recordset

        Dim rSQL   As New ADODB.Recordset

        Dim conDBF As String

        Dim conSQL As String

        On Error Resume Next

100     Me.MousePointer = vbHourglass

        Dim XR        As String

        Dim sumes(150) As Single

        On Error GoTo EXO

110     If ODBC.Value = vbChecked Then
            ' DBGrid1.Visible = True
120         DATA2.Connect = "ODBC;" + gConnect
130         DATA2.RecordSource = PutDates(Text2.Text)
140         DATA2.Refresh
        Else

            ' DBGrid1.Visible = False
            If Len(FDSN) > 1 Then
                'gdb2.Open FDSN
                Adodc2.ConnectionString = FDSN
            Else
170             'Gdb.Execute PutDates(TEXT2.Text), lo
                Adodc2.ConnectionString = gConnect
            End If

            Adodc2.RecordSource = PutDates(Text2.Text)
            Adodc2.Refresh

180         Adodc2.Recordset.MoveFirst

           ' On Error Resume Next

            On Error GoTo ektos
             Dim KL As Long
             
             
             
             
'------------------------------------- SHOW CANCEL -------------------------------
     Dim FF As New UDialog
     FF.SHOW
     FF.OKButton.Visible = False
     FF.List1.Visible = False
     FF.CancelButton.Caption = "ƒ…¡ œ–«"
     FF.CancelButton.Top = 120
     FF.CancelButton.Left = 120
     FF.CancelButton.Width = 2895
     FF.CancelButton.Height = 495
     FF.Width = 3210
     FF.Height = 810
     FF.Left = Screen.Width / 2 - FF.Width / 2
     FF.Top = Screen.Height / 2 - FF.Height / 2
     FF.Caption = "ƒ…¡ƒ… ¡”…¡ ’–œÀœ√…”Ãœ’"
     MDIForm1.MousePointer = vbHourglass
'------------------------------------- SHOW CANCEL -------------------------------

             
             
             
             
             
             
             
             
            
190         Do While Not Adodc2.Recordset.EOF
                 If FF.CancelButton.Enabled = False Then
                          FF.Hide
                          Set FF = Nothing
                          Exit Do
                 End If

200             For k = 0 To Adodc2.Recordset.FIELDS.Count - 1
                  If Adodc2.Recordset.FIELDS(k).Type = 200 Then
                  Else
210                 If IsNumeric(Adodc2.Recordset.FIELDS(k).Value) And Adodc2.Recordset.FIELDS(k).Type <> 202 Then
220                     sumes(k) = sumes(k) + nNull(Adodc2.Recordset.FIELDS(k).Value)
                    End If
                  End If
                Next

                DoEvents
                KL = KL + 1: Me.Caption = KL
230             Adodc2.Recordset.MoveNext

                '  Exit Do
            Loop

                          FF.Hide
                          Set FF = Nothing
 MDIForm1.MousePointer = vbNormal



240         For k = 0 To Adodc2.Recordset.FIELDS.Count - 1

250             If sumes(k) > 0 Then
260                 TDBGrid.Splits(0).columns(k).FooterText = Format(sumes(k), "######0.00")
                    

                End If
                
                If Plath_Grid(k) > 0 Then
                    TDBGrid.Splits(0).columns(k).Width = Plath_Grid(k)
                End If
                
               
            Next
            
            
            
          '  Dim CX As TrueDBGrid80.Column

''Set CX = TDBGrid.columns.ADD(k)

 

' Initialize the new Column 0

'With CX

  '  .Visible = True        ' Make it visible

   ' .DataField = "LAST"    ' Set the column's database field

   ' .Caption = "Last Name" ' Set the column's caption

'End With

 


            
            
            
            
            
            
            
             TDBGrid.Refresh
             
             
             
        End If

270     TDBGrid.AlternatingRowStyle = True

280     TDBGrid.OddRowStyle.BackColor = &H8000000F   ' GRI   vbCyan
290     TDBGrid.EvenRowStyle.BackColor = &HFFFFC0   'OYRANI     &H8000000F  ' GRI

        'Dim k As Long

        'If Check1 Then
        '   print7_excel UCase(PutDates(Text2.Text)), "11111111" 'sugm_str
        'Else
        Dim gr As Integer, omadop As Integer

'        GR = MsgBox("Õ· „ﬂÌÂÙ·È ÔÏ·‰ÔÔﬂÁÛÁ ÛÙÔ 1Ô Â‰ﬂÔ;", vbYesNo)
'
'        If GR = vbYes Then
'            omadop = 1
'        Else
'            omadop = 0
'        End If
'
        
        
        
        
        Dim sum_s As String
        If Len(sumstring.Text) > 2 Then  ' tag
           sum_s = sumstring.Text 'tag
        Else
           sum_s = "00111111111111111111111111"
        End If
        
        
300     print3_xar PutDates(Text2.Text), sum_s, TITLOS.Text, 0   'omadop     'sugm_str
        'End If

310     Me.MousePointer = vbNormal

        Exit Sub

ektos:
MsgBox Err.Description
Me.MousePointer = vbNormal
        Exit Sub

EXO:
320     MsgBox Err.Description

330     Me.MousePointer = vbNormal

340     'Resume Next

        Exit Sub

        '<EhFooter>
        Exit Sub

CMDPROVOLI_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command10_Click()

        '<EhHeader>
        On Error GoTo Command10_Click_Err

        '</EhHeader>

100     Unload Me

110     If gApoMenu = True Then
120         SendKeys "%"
130         SendKeys "{RIGHT}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

Command10_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.Command10_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.Command10_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub CMDEXECUTE_Click()

        '<EhHeader>
        On Error GoTo CMDEXECUTE_Click_Err

        '</EhHeader>
        Dim DBF As New ADODB.Connection

        Dim sql As New ADODB.Connection

        Dim rDBF As New ADODB.Recordset

        Dim rSQL As New ADODB.Recordset

        Dim msql As String

        Dim conDBF As String

        Dim conSQL As String

        Dim lo As Long

        Dim DB As Database

        Dim SEIRES(30)

        Dim MLINE As String, k

100     For k = 1 To 30: SEIRES(k) = "": Next

110     lo = 0
120     Me.MousePointer = vbHourglass

        On Error GoTo LATOS  'On Error Resume Next

130     If ODBC.Value = Checked Then

140         Set DB = OpenDatabase("", False, False, gConnect)
150         DB.Execute PutDates(Text2.Text)
160         lo = DB.RecordsAffected
        Else

            If Len(FDSN) > 1 Then

                Dim gdb2 As New ADODB.Connection
                On Error GoTo LATOS2
                gdb2.Open FDSN
                Me.Caption = FDSN
                
                msql = PutDates(Text2.Text)
                gdb2.Execute msql, lo
                gdb2.Close
                 
            Else
170             Gdb.Execute PutDates(Text2.Text), lo
            End If

            ' to ekana gia na typonei polaples entoles xorismenes me to ; alla den xreiazetai
            '   MLINE = TEXT2.Text
            '  DUM = FETES2_DELIM(MLINE, SEIRES)

            ' For k = 1 To 30
            '  If Len(SEIRES(k)) > 2 Then
            '       Trim(SEIRES(k)), LO
            '  End If
            ' Next
        End If

180     MsgBox str(lo) + " ≈√√—¡÷≈” ≈Õ«Ã≈—Ÿ»« ¡Õ"
190     Me.MousePointer = vbNormal

        Exit Sub
        
LATOS2:
  MsgBox "‰ÂÌ Â„ÈÌÂ Á Û˝Ì‰ÂÛÁ"
  On Error Resume Next
   gdb2.Close
   
Exit Sub

        
        

LATOS:
200     MsgBox Err.Description
         
210     Resume Next

        '<EhFooter>
        Exit Sub

CMDEXECUTE_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'

'
Private Sub Form_Load()

    '<EhHeader>
    On Error GoTo Form_Load_Err

    '</EhHeader>

98  FDSN = ""
100     mForm_Load Me, fh, fw, ft, fl

110     Me.Picture = LoadPicture(gPicture)

        On Error Resume Next

120     F_DOROKARTA = Val(FINDPARAMETROI(1, "BOHU6", "F_DOROKARTA", "9139393", "œ… ƒŸ—œ ¡—‘≈” ¡—◊…∆œ’Õ ¡–œ "))


f_gfk = FINDPARAMETROI(1, "BOHU8", "F_GFK", 0, "GFK ¡–œ”‘œÀ« ¡—◊≈…ŸÕ 0=OXI ¡ÀÀœ…Ÿ”  Ÿƒ… œ”")
If f_gfk = 1000 Then
  gfk.Visible = True
End If



        '   'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""

        '    'SkinFramework1.ApplyWindow Me.hWnd
        '    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

        ' F_DOROKARTA = Val(FindParametroi(1,"BOHU6", "F_DOROKARTA", "9139393", "œ… ƒŸ—œ ¡—‘≈” ¡—◊…∆œ’Õ ¡–œ "))

130     f_pelMERC = FINDPARAMETROI(1, "PAR1", "F_PELMERC", "..", "◊—«”‘«” MERCURY")

        Text2.ToolTipText = "·ÎÎ·Ê˘ ÙÔÌ ÙÈÙÎÔ –·Ò·ÏÂÙÒÔÚ 1,2,3 " + Chr(13) + " Í·È 4Á ÛÂÈÒ‹ ·ÎÎ·„ﬁ dsn " + Chr(13) + " .˜. 1Á ÛÂÈÒ· /*#BARCODE*/ ...4Á ÛÂÈÒ‹ /* #DSN=MERURY;... */  5Á ÛÂÈÒ‹ 00111 ·ËÒÔﬂÛÏ·Ù· ÛÙÁÎ˛Ì"

140     If Left(f_pelMERC, 3) = "SOK" Then
150         Command5.Visible = True
        End If

160     F_EIKONA = Val(FINDPARAMETROI(1, "APOT2", "F_EIKONA", "1", "1=EIKONA=> Ÿƒ… œ” 2=>Ã≈ ¬œ«». À≈…ƒ… 11= VDIKOS.JPG 12=BOH.JPG 0=œ◊… ≈… œÕ≈”"))

        'On Error Resume Next
        ' File1.Path = LCase("c:\mercvb\queries\EID")

170     If Len(UCase(Dir("c:\DELTA.TXT", vbDirectory))) > 0 Then
180         Command2.Visible = True
190         Command4.Visible = True
        End If

200     If Len(UCase(Dir("c:\mercvb\queries\RYTMIZOMENES", vbDirectory))) = 0 Then
210         MkDir "c:\mercvb\queries\RYTMIZOMENES"
        End If

220     Dir1.Path = "C:\mercvb\queries\RYTMIZOMENES"

230     File1.Refresh

240     cd1.InitDir = "c:\mercvb\queries\RYTMIZOMENES"

250     DoEvents

260     File1.Refresh
270     APO = Now
280     EOS = Now

290     TDBGrid.Splits(0).ExtendRightColumn = True
300     TDBGrid.ColumnFooters = True

310     Me.KeyPreview = True

        On Error Resume Next

        'File1.Path = "C:\mercvb\queries"

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.Form_Load " & "at line " & Erl

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

160     Adodc2.Recordset.sort = strSort

        '<EhFooter>
        Exit Sub

TDBGrid_GroupColMove_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.TDBGrid_GroupColMove " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.TDBGrid_GroupColMove " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub TDBGrid_HeadClick(ByVal ColIndex As Integer)

        '<EhHeader>
        On Error GoTo TDBGrid_HeadClick_Err

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

100     For k = 0 To Adodc2.Recordset.FIELDS.Count - 1
110         sumes(k) = TDBGrid.Splits(0).columns(k).FooterText    '  = Format(SUMES(k), "######0.00")
        Next

120     If Adodc2.Recordset.sort = "[" & TDBGrid.columns(ColIndex).DataField & "] asc" Then   ' strSort
130         Adodc2.Recordset.sort = "[" & TDBGrid.columns(ColIndex).DataField & "] desc"    ' strSort
        Else
140         Adodc2.Recordset.sort = "[" & TDBGrid.columns(ColIndex).DataField & "] asc"    ' strSort
        End If

150     For k = 0 To Adodc2.Recordset.FIELDS.Count - 1
160         TDBGrid.Splits(0).columns(k).FooterText = sumes(k)   '  = Format(SUMES(k), "######0.00")
        Next

        '    Adodc2.Recordset.Sort = "[" & TDBGrid.Columns(ColIndex).DataField & "] desc" ' strSort

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

Private Sub TDBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

        '<EhHeader>
        On Error GoTo TDBGrid_RowColChange_Err

        '</EhHeader>

        Dim PI

100     TDBGrid.Col = 0

110     If F_EIKONA > 0 Then

            On Error Resume Next

120         If F_EIKONA = 2 Then    '‚ÔÁËÁÙÈÍÔÛ Í˘‰ÈÍÔÛ
130             PI = "C:\MERCVB\IMAGES\" + TFBGrid.Text
140         ElseIf F_EIKONA = 1 Then    ' kvdikos
150             PI = "C:\MERCVB\IMAGES\" + TDBGrid.Text
160         ElseIf F_EIKONA = 12 Then    '‚ÔÁËÁÙÈÍÔÛ Í˘‰ÈÍÔÛ +.jpg
170             PI = "C:\MERCVB\IMAGES\" + TDBGrid.Text + ".JPG"
180         ElseIf F_EIKONA = 11 Then    ' kodikos .jpg
190             PI = "C:\MERCVB\IMAGES\" + TDBGrid.Text + ".JPG"
            End If

200         If Len(Dir(PI)) > 0 Then
210             PICTURE1.Picture = LoadPicture(PI)
220             Picture2.PaintPicture PICTURE1.Picture, 0, 0, Picture2.Width, Picture2.Height, 0, 0, PICTURE1.Width, PICTURE1.Height
            Else
230             PICTURE1.Picture = LoadPicture()
240             Picture2.Picture = LoadPicture()
            End If

        End If

        '<EhFooter>
        Exit Sub

TDBGrid_RowColChange_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.apot7.TDBGrid_RowColChange " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.apot7.TDBGrid_RowColChange " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

