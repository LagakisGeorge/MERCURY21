VERSION 5.00
Object = "{CCA214C0-DFEB-4C91-9F0D-2665F77F6E23}#1.2#0"; "IDAutomationLinear.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form paralabh 
   BackColor       =   &H00FF0000&
   Caption         =   "–¡—¡À¡¬« ¡Õ‘… ≈…Ã≈ÕŸÕ √…¡ ≈–…” ≈’«"
   ClientHeight    =   10875
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   16230
   ForeColor       =   &H80000014&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10875
   ScaleWidth      =   16230
   WindowState     =   2  'Maximized
   Begin VB.Frame FramePelath 
      BackColor       =   &H00C0FFC0&
      Caption         =   "¡Ì·ÊﬁÙÁÛÁ –ÂÎ‹ÙÁ - ƒÂÎÙﬂÔ ¡ÔÛÙÔÎﬁÚ"
      Height          =   5535
      Left            =   8640
      TabIndex        =   41
      Top             =   4200
      Width           =   6975
      Begin VB.TextBox Text8 
         Height          =   420
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   1980
      End
      Begin VB.CommandButton daPelath 
         Caption         =   "ƒ¡ –≈À¡‘«"
         Height          =   555
         Left            =   120
         TabIndex        =   47
         Top             =   2760
         Width           =   1515
      End
      Begin VB.TextBox da 
         Height          =   285
         Left            =   1680
         TabIndex        =   45
         Top             =   3825
         Width           =   1320
      End
      Begin VB.CommandButton Command4 
         Caption         =   "◊—≈Ÿ”« ≈–…” ≈’¡”Ã≈Õœ’"
         Height          =   615
         Left            =   2520
         TabIndex        =   44
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Caption         =   "¡ÒﬂËÏÁÛÁ"
         Height          =   255
         Left            =   3600
         TabIndex        =   43
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CommandButton cmdE·ÌÂÍÙı˘ÛÁ≈ÙÈÍ›Ù·Ú 
         Caption         =   "E·ÌÂÍÙı˘ÛÁ ÂÙÈÍ›Ù·Ú"
         Height          =   495
         Left            =   120
         TabIndex        =   42
         Top             =   4200
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   1710
         TabIndex        =   46
         Top             =   3375
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Format          =   293076993
         CurrentDate     =   39310
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
         Bindings        =   "paralabh.frx":0000
         Height          =   1785
         Left            =   120
         TabIndex        =   48
         Top             =   840
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   3149
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¡ÒÈËÏ¸Ú ƒ¡"
         Height          =   315
         Left            =   120
         TabIndex        =   51
         Top             =   3810
         Width           =   1575
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«Ï.ƒ¡"
         Height          =   330
         Left            =   120
         TabIndex        =   50
         Top             =   3360
         Width           =   1635
      End
   End
   Begin VB.Frame FrameApokatastasi 
      BackColor       =   &H00C0FFC0&
      Caption         =   "¡ÔÍ·Ù‹ÛÙ·ÛÁ"
      Height          =   5535
      Left            =   8640
      TabIndex        =   32
      Top             =   120
      Width           =   6975
      Begin VB.TextBox TEXNIKOS 
         Height          =   285
         Left            =   5400
         TabIndex        =   54
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "–·ÒÂÎﬁˆËÁ"
         Height          =   255
         Left            =   2520
         TabIndex        =   53
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox kostos 
         Height          =   285
         Left            =   2520
         TabIndex        =   34
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox ono2 
         Height          =   285
         Left            =   240
         TabIndex        =   33
         Top             =   2160
         Width           =   6495
      End
      Begin MSComCtl2.DTPicker hm1 
         Height          =   330
         Left            =   2520
         TabIndex        =   35
         Top             =   1080
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Format          =   292749313
         CurrentDate     =   39310
      End
      Begin MSComCtl2.DTPicker hm3 
         Height          =   330
         Left            =   2520
         TabIndex        =   36
         Top             =   3240
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   292749313
         CurrentDate     =   39310
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‘≈◊Õ… œ” :1=GAUB 2=BEN 3=À¡√ 4=»œƒ"
         Height          =   195
         Left            =   3480
         TabIndex        =   55
         Top             =   240
         Width           =   3300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EÒ„·ÛﬂÂÚ Ôı ›„ÈÌ·Ì"
         Height          =   195
         Left            =   240
         TabIndex        =   52
         Top             =   1800
         Width           =   1590
      End
      Begin VB.Label lbl6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ¸ÛÙÔÚ"
         Height          =   195
         Left            =   360
         TabIndex        =   40
         Top             =   480
         Width           =   555
      End
      Begin VB.Label lbl«ÏÂÒ–·Ò·Î·‚ﬁÚ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«ÏÂÒ. –·Ò·Î·‚ﬁÚ ·Ô ÂÎ‹ÙÁ"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   3360
         Width           =   2235
      End
      Begin VB.Label lbl«ÏÂÒ¡ÔÂÒ‹Ù˘ÛÁÚ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«ÏÂÒ.¡ÔÂÒ‹Ù˘ÛÁÚ"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   38
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lbl ¸ÛÙÔÚ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ¸ÛÙÔÚ"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   0
         TabIndex        =   37
         Top             =   120
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   " ·Ù·˜˛ÒÁÛÁ ƒÈ¸ÒË˘ÛÁÚ"
      Enabled         =   0   'False
      Height          =   240
      Left            =   6240
      TabIndex        =   31
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox Noumero 
      BackColor       =   &H0000FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   7560
      TabIndex        =   29
      Top             =   720
      Width           =   1095
   End
   Begin VB.PictureBox foto2 
      Height          =   2655
      Left            =   5280
      ScaleHeight     =   2595
      ScaleWidth      =   3315
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CommandButton cmdEikona 
      BackColor       =   &H00FFFF00&
      Caption         =   "÷Ÿ‘œ√—¡÷…¡"
      Height          =   315
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CheckBox chkTSANTA 
      Caption         =   "‘Û‹ÌÙ·"
      Height          =   255
      Left            =   3480
      TabIndex        =   24
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CheckBox chkTROF 
      Caption         =   "TÒÔˆÔ‰ÔÙÈÍ¸"
      Height          =   255
      Left            =   1560
      TabIndex        =   23
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "≈·ÌÂÍÙ˝˘ÛÁ ÛÂ ‚È‚ÎﬂÔ"
      Height          =   615
      Left            =   12240
      TabIndex        =   20
      Top             =   9960
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8280
      Top             =   9960
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.ComboBox ONO 
      Height          =   315
      ItemData        =   "paralabh.frx":0015
      Left            =   1560
      List            =   "paralabh.frx":0017
      TabIndex        =   0
      ToolTipText     =   "–ﬂÌ·Í·Ú 41"
      Top             =   1440
      Width           =   3735
   End
   Begin VB.CommandButton Nea_eggrafh 
      Caption         =   "Õ›· ≈„„Ò·ˆﬁ"
      Height          =   240
      Left            =   2640
      TabIndex        =   19
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "≈ÈÍÂˆ·Îﬂ‰· ¬È‚ÎﬂÔı"
      Height          =   615
      Left            =   10800
      TabIndex        =   18
      Top             =   10080
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   13560
      Top             =   10080
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
   Begin VB.ComboBox SKOPOS 
      Height          =   315
      ItemData        =   "paralabh.frx":0019
      Left            =   1560
      List            =   "paralabh.frx":002C
      TabIndex        =   16
      Text            =   "≈–…” ≈’«"
      Top             =   2280
      Width           =   2745
   End
   Begin VB.ComboBox MON 
      Height          =   315
      ItemData        =   "paralabh.frx":0069
      Left            =   6480
      List            =   "paralabh.frx":0076
      TabIndex        =   15
      Text            =   "‘≈Ã"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox POSO 
      Height          =   300
      Left            =   5520
      TabIndex        =   14
      Text            =   "1"
      Top             =   1440
      Width           =   660
   End
   Begin VB.CommandButton KATAXORHSH 
      Caption         =   " ·Ù·˜˛ÒÁÛÁ Õ›·Ú"
      Enabled         =   0   'False
      Height          =   240
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox EPA 
      Height          =   435
      Left            =   1440
      TabIndex        =   5
      Top             =   9720
      Width           =   5175
   End
   Begin VB.TextBox DIE 
      Height          =   435
      Left            =   7320
      TabIndex        =   4
      Top             =   9720
      Width           =   5175
   End
   Begin VB.CheckBox DATA 
      BackColor       =   &H00FF8080&
      Caption         =   "‘· ÛÙÔÈ˜Âﬂ· ˜ÒÂÈ‹ÊÔÌÙ·È- Õ· ÏÁÌ Û‚ÁÛÙÔıÌ"
      ForeColor       =   &H80000014&
      Height          =   345
      Left            =   840
      TabIndex        =   11
      Top             =   10200
      Width           =   3855
   End
   Begin VB.TextBox BLABH 
      Height          =   1020
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2640
      Width           =   3615
   End
   Begin VB.TextBox THL 
      Height          =   390
      Left            =   1560
      TabIndex        =   6
      Top             =   960
      Width           =   2730
   End
   Begin VB.TextBox EPO 
      Height          =   390
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   5130
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7800
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Bindings        =   "paralabh.frx":0089
      Height          =   3570
      Left            =   120
      TabIndex        =   30
      Top             =   5640
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   6297
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
   Begin VB.Label lbl–œ”º‘«‘¡ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "–œ”º‘«‘¡"
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   5520
      TabIndex        =   26
      Top             =   1200
      Width           =   1485
   End
   Begin VB.Label lbl¡ÎÎ· 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¡ÎÎ·"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5400
      TabIndex        =   25
      Top             =   1920
      Width           =   1680
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode1 
      Height          =   660
      Left            =   0
      TabIndex        =   21
      Top             =   5520
      Visible         =   0   'False
      Width           =   2085
      _cx             =   3678
      _cy             =   1164
      Enabled         =   -1  'True
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      SymbologyId     =   13
      DataToEncode    =   "!10001"
      Orientation     =   0
      BarHeight       =   1
      NarrowBarWidth  =   0,03
      Wide2NarrowRatio=   2
      AddCheckDigit   =   1
      AddCheckDigitToText=   1
      Code128CharSet  =   1
      UPCESystem      =   0
      EANUPCSupplement=   0
      ShowText        =   1
      CodabarStartCharacter=   "A"
      CodabarStopCharacter=   "B"
      LeftMarginCM    =   0,2
      TopMarginCM     =   0,2
      SupplementToEncode=   ""
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "–¡—¡À¡¬« ¡Õ‘… ≈…Ã≈ÕŸÕ √…¡ ≈–…” ≈’«"
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
      Height          =   330
      Left            =   195
      TabIndex        =   17
      Top             =   45
      Width           =   7035
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "≈‹„„ÂÎÏ·"
      ForeColor       =   &H80000014&
      Height          =   330
      Left            =   1440
      TabIndex        =   13
      Top             =   9600
      Width           =   1710
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ƒÈÂıËıÌÛÁ"
      ForeColor       =   &H80000014&
      Height          =   330
      Left            =   7440
      TabIndex        =   12
      Top             =   9480
      Width           =   1710
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "¬Î‹‚Á /≈Ì›Ò„ÂÈÚ"
      ForeColor       =   &H80000014&
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   1365
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "”ıÛÍÂıﬁ"
      ForeColor       =   &H80000014&
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "‘ÁÎ›ˆ˘ÌÔ"
      ForeColor       =   &H80000014&
      Height          =   390
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "œÌÔÏ·ÙÂ˛ÌıÏÔ"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   1200
   End
End
Attribute VB_Name = "paralabh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim F_ArPerEtik

Dim F_BIBLIO_EIS As Integer
Dim F_ENHM_EGGTIM As Integer
Dim F_XBARCODE, F_YBARCODE
Attribute F_YBARCODE.VB_VarUserMemId = 1073938432



Dim Col         As TrueOleDBGrid80.Column
Dim ColS As TrueOleDBGrid80.columns






Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)
Attribute fh.VB_VarUserMemId = 1073938434

Private Sub BLABH_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo BLABH_KeyUp_Err

        '</EhHeader>
100     If KeyCode = 13 Then
110         KATAXORHSH.SetFocus
        End If

        '<EhFooter>
        Exit Sub

BLABH_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.BLABH_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.BLABH_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdCommand1_Click()

End Sub

Private Sub Check1_Click()

    If Check1.Value = vbChecked Then
        HM1.Enabled = True
    Else
        HM1.Enabled = False
    End If
    


End Sub

Private Sub cmdEikona_Click()
      '</EhHeader>
        Dim F_EIKONA As Long

100     F_EIKONA = 11
110     CommonDialog1.FILENAME = "*.jpg"
120     CommonDialog1.ShowOpen

        Dim m As String

        On Error Resume Next

130     m = "" + Noumero.Text


140    ' If F_EIKONA = 11 Then
150         m = m + ".JPG"
       ' End If

160     FileCopy CommonDialog1.FILENAME, "c:\mercvb\images\" + m
    
170     foto2.Picture = LoadPicture("c:\mercvb\images\" + m)
            
180     StretchSourcePictureFromPicture Me.foto2.Picture, Me.foto2
    
        '<EhFooter>
        Exit Sub
End Sub

Private Sub cmdUpdate_Click()

Gdb.Execute "update EID SET HM3='" + Format(HM3.Value, "MM/DD/YYYY") + "' WHERE KOD='" + Noumero.Text + "'"
If HM1.Enabled = True Then
    Gdb.Execute "update EID SET HM1='" + Format(HM1.Value, "MM/DD/YYYY") + "' WHERE KOD='" + Noumero.Text + "'"
End If


Gdb.Execute "update EID SET NUM3=" + Replace(kostos.Text, ",", ".") + " WHERE KOD='" + Noumero.Text + "'"

Gdb.Execute "update EID SET ONO2='" + ONO2.Text + "' WHERE KOD='" + Noumero.Text + "'"

Gdb.Execute "update EID SET NUM2=" + Replace(kostos.Text, ",", ".") + " WHERE KOD='" + Noumero.Text + "'"

If Val(TEXNIKOS.Text) > 0 Then
   Gdb.Execute "update EID SET NUM3=" + Replace(TEXNIKOS.Text, ",", ".") + " WHERE KOD='" + Noumero.Text + "'"
End If

End Sub

Private Sub KATAXORHSH_Click()

        'ƒ«Ã…œ’—√œ’Ã≈ ≈Õ¡ RECORD ”‘œ EGGTIM KAI STO EID
        ' KAI ‘’–ŸÕœ’Ã≈
        '<EhHeader>
        On Error GoTo KATAXORHSH_Click_Err
        
        
        If Nea_eggrafh.Enabled Then
           MsgBox "·ÙÁÛÙÂ Õ›· ≈„„Ò·ˆﬁ"
           Exit Sub
        End If
        
        
        KATAXORHSH.Enabled = False
        
        

        '</EhHeader>
        Dim x As String, sql As String

        Dim N As Integer

        Dim PELATHS

        Dim PR, FOUND

        Dim DUM As String

        Dim R   As New ADODB.Recordset

130     F_XBARCODE = Val(FINDPARAMETROI(1, "PARALABH", "F_XBARCODE", "348", "barcode ◊ "))   'posa psifia tha exei h kathe seira
140     F_YBARCODE = Val(FINDPARAMETROI(1, "PARALABH", "F_YBARCODE", "824", "barcode Y "))   'posa psifia tha exei h kathe seira

        'CSBmk_CS <If Len(da.Text) > 0>
100     If Len(da.Text) > 0 Then
110         MsgBox "ƒ«ÀŸ”¡‘≈ ƒ¡ –≈À¡‘«. –¡‘«”‘≈ ‘œ ¡ÀÀœ –À« ‘—œ"

            Exit Sub

        'CSBmk_CS <End If>
        End If

        '120  FOUND = False
        '130  For Each PR In Printers
        '140      If InStr(PR.DeviceName, "POSPARALABH") > 0 Then
150     FOUND = True

        '             Exit For
        '         End If
        '     Next
        'CSBmk_CS <If Not FOUND>
160     If Not FOUND Then
170         MsgBox "ƒ≈Õ ’–¡—◊≈… ”≈ ¡’‘«Õ ‘«Õ »≈”« œ ≈ ‘’–Ÿ‘«” √…¡ ‘…” –¡—¡À¡¬≈”"
180         Unload Me

            Exit Sub

        'CSBmk_CS <End If>
        End If

        'CSBmk_CS <If EPO.Text = Adodc1.Recordset("EPO")>
190     If EPO.Text = Adodc1.Recordset("EPO") Then
200         PELATHS = Adodc1.Recordset("KOD")
        'CSBmk_CS <Else>
        Else
210         PELATHS = "1112"
        'CSBmk_CS <End If>
        End If

        'CSBmk_CS <If Len(Trim(EPO.Text)) = 0>
220     If Len(Trim(EPO.Text)) = 0 Then
230         MsgBox "ƒÂÌ ÛıÏÎÁÒ˛Û·ÙÂ ¸ÌÔÏ·"

            Exit Sub

        'CSBmk_CS <End If>
        End If

        'CSBmk_CS <If Len(Trim(ono.Text)) = 0>
240     If Len(Trim(ono.Text)) = 0 Then
250         MsgBox "ƒÂÌ ÛıÏÎÁÒ˛Û·ÙÂ Âﬂ‰ÔÚ"

            Exit Sub

        'CSBmk_CS <End If>
        End If

        On Error Resume Next

        'ƒ«Ã…œ’—√…¡ ‘œ’  Ÿƒ… œ’ ¡–œ»« «”

260     x = FIND_NUM(0, 0, 0)    ' ¬—…” Ÿ ‘œ ‘≈À≈’‘¡…œ Õœ’Ã≈—œ KAI
        'ƒ…¡¬¡∆Ÿ ‘«Õ –—œ«√œ’Ã≈Õ« «Ã≈—œÃ«Õ…¡ √…¡ Õ¡ Ã«Õ ≈◊Ÿ «Ã≈—œÃ«Õ…¡ Ã… —œ‘≈—« ‘«” ‘≈À≈’‘¡…¡”  ¡‘¡◊Ÿ—«”«”

270     R.Open "SELECT * FROM EID WHERE KOD='" + x + "'", Gdb, adOpenDynamic, adLockOptimistic

        'CSBmk_CS <If Not R.EOF>
280     If Not R.EOF Then
290         MsgBox "’‹Ò˜ÂÈ Ô Í˘‰ÈÍ¸Ú " + x

            Exit Sub

        'CSBmk_CS <End If>
        End If
        
        'CSBmk_CS <If chkTSANTA.Value = vbChecked>
        If chkTSANTA.Value = vbChecked Then
        
           BLABH.Text = BLABH.Text + " ;≈˜ÂÈ ‘Û‹ÌÙ· "
        
        
        'CSBmk_CS <End If>
        End If
        
        
         'CSBmk_CS <If chkTROF.Value = vbChecked>
         If chkTROF.Value = vbChecked Then
        
           BLABH.Text = BLABH.Text + " ;≈˜ÂÈ ÙÒÔˆÔ‰ÔÙÈÍ¸ "
        
        
        'CSBmk_CS <End If>
        End If
        
        
        
        
        
        

300     sql = "INSERT INTO EID (KOD,ONO,FPA,MEMO,MON,CH1,NUM1,CH2,CH3,CH4,CH5,CH6,HM2) VALUES (" & "'" + x + " ','" + Replace(ono.Text, "'", "-") + "',2,'" + Replace(BLABH.Text, "'", "-") + "'," & "'" + MON.Text + "','" + SKOPOS.Text + "'," + POSO.Text + "," & "'" + Left(EPO.Text, 20) + "','" + Left(DIE.Text, 20) + "','" + EPA.Text + "','" + THL.Text + "','" + PELATHS + "'," & "GETDATE())"

310     Gdb.Execute sql, N

        'CSBmk_CS <If N = 0>
320     If N = 0 Then
330         MsgBox " ‰ÂÌ Í·Ù·˜˘ÒﬁËÁÍÂ "

            Exit Sub

        'CSBmk_CS <End If>
        End If

340     ' DUM = FIND_NUM(1, 0, 0)    ' Í·Ù·˜˘Ò˛ ÙÔ Â¸ÏÂÌÔ ÌÔ˝ÏÂÒÔ

        'ƒ«Ã…œ’—√…¡ ‘«”  …Õ«”«” ¡–œ»« «” (≈GGTIM)
        
        
        
       If F_ENHM_EGGTIM = 1 Then

350     sql = "INSERT INTO EGGTIM (HME,KODE,ONOMA,FPA,EIDOS,APOT,POSO,XRE,PELKOD,ATIM) VALUES (" & "'" + Format(Now, "MM/DD/YYYY") + "','" + x + " ','" + ono.Text + "',2,'e',1," + POSO.Text + "," + POSO.Text + ",'" + PELATHS + "','S" + mID(x, 2, 6) + "')"
360     Gdb.Execute sql, N

       End If


        On Error Resume Next

        Dim STR_EKT As String

370     STR_EKT = x + " " + Format(Now, "DD/MM/YYYY") + " " + Left(EPO.Text + Space(25), 25) + " " + Left(EPA.Text + Space(20), 20) + " " + Left(DIE.Text + Space(22), 22) + " " + Left(ono.Text + Space(25), 25) + " " + Left(MON.Text + Space(4), 3) + " " + Format(POSO, "#0") + " " + Left(SKOPOS.Text + Space(15), 15)
    
        'CSBmk_CS <If F_BIBLIO_EIS = 1>
380     If F_BIBLIO_EIS = 1 Then
390         Open "LPT1" For Output As 1
400         Print #1, Chr(15) + to437(STR_EKT)
410         Close 1

        'CSBmk_CS <End If>
        End If

        'CSBmk_CS <For Each PR In Printers>
420     For Each PR In Printers

            'CSBmk_CS <    If InStr(PR.DeviceName, "SATO") > 0>
430         If InStr(PR.DeviceName, "SATO") > 0 Then
440             Set Printer = PR

                Exit For

            'CSBmk_CS <    End If>
            End If

        'CSBmk_CS <Next>
        Next

450     MsgBox " ¡‘¡◊Ÿ—«»« ≈ Ã≈ ‘œÕ  Ÿƒ… œ " + x

        Dim r2 As New ADODB.Recordset

460     r2.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic

        Dim ZEBRA As Integer

470     ZEBRA = 0

        '480  If Len(Dir("C:\MERCVB\ZEBRA.TXT")) > 1 Then
490     ZEBRA = 1
        '    End If

        Dim c As Integer

500     c = F_ArPerEtik

        Dim W As Integer

        'CSBmk_CS <For W = 1 To 2>
510     For W = 1 To 2

            'CSBmk_CS <    If ZEBRA = 1>
520         If ZEBRA = 1 Then

530             Label9.Caption = "--"
540             Printer.CurrentY = F_YBARCODE

550             BarCode1.SymbologyID = CODE128    'CODE39

560             BarCode1.DataToEncode = x
570             Printer.PaintPicture BarCode1.Picture, F_XBARCODE, F_YBARCODE, , 300    ' ORIZ,KATHETA
            'CSBmk_CS <    Else>
            Else
580             Printer.FontName = "128"
590             Printer.FontSize = 6
600             Printer.CurrentY = 10
610             Printer.Print Tab(c + 3); x
            'CSBmk_CS <    End If>
            End If

620         Printer.CurrentY = 50

630         Printer.FontBold = False    '  True
640         Printer.FontSize = 6
650         Printer.FontUnderline = False    'True
660         Printer.FontName = "Arial Greek"

670         Printer.FontSize = 7
680         Printer.Print Tab(c + 1); r2("PELONO")  ' "À·„‹ÍÁÚ Computers"
690         Printer.FontSize = 7
700         Printer.Print Tab(c + 1); "ÙÁÎ." + r2("PELEPA")  '25210-22225"   'CommonDialog1'.FontName

710         Printer.FontUnderline = False

720         Printer.CurrentY = 360

730         Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
740         Printer.FontBold = True

750         Printer.FontSize = 10
760         Printer.Print Tab(c); x;

770         Printer.CurrentY = 390
780         Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
790         Printer.FontBold = False

800         Printer.FontSize = 7
810         Printer.Print Tab(c + 30); Now

820         Printer.CurrentY = 530

830         Printer.FontSize = 8
840         Printer.FontBold = True
850         Printer.Print Tab(c); Left(EPO.Text, 22) + " " + THL.Text

860         Printer.FontName = "Courier New"
870         Printer.FontSize = 7
880         Printer.Print Tab(c); Left(ono, 40)
890         Printer.Print Tab(c); Left(BLABH.Text, 35)
900         Printer.Print Tab(c); mID(BLABH.Text, 36, 40)

            ' Printer.Print
            ' Printer.Print Tab(c); Now
910         Printer.EndDoc

        'CSBmk_CS <Next W>
920     Next W

        'GoTo 10

930     r2.Close

940     Printer.EndDoc

950     Unload Me

        '<EhFooter>
        Exit Sub

KATAXORHSH_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function FIND_NUM(ByVal DUM As Long, _
                  ByVal DELTIO As Integer, _
                  ByVal KARFOTO As Integer) As String


'      ElseIf KARFOTO = -1 Then
            ' NA MHN ANEBAZEI KAI KANEI UPDATE




        'DELTIO=1 ƒ¡ –≈À¡‘«   DELTIO=0 ƒ≈À‘…œ –œ”.–¡—¡À¡¬«”
        'dum=0.5 tote  n=n+1
        'dum>=1 bazo ton metrhth sto dum
        '************************
        '<EhHeader>
        On Error GoTo FIND_NUM_Err

        '</EhHeader>
        Dim rec  As ADODB.Recordset

        'rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
        Dim REC2 As New ADODB.Recordset

100     REC2.Open "SELECT * FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
110     REC2.MoveFirst

120     If DELTIO = 1 Then
130         REC2.Move 8  ' ‰· ÂÎ·Ù˘Ì
        Else
140         REC2.Move 9    ' ‰ÂÎ.ÔÛ.·Ò·Î·‚ﬁÚ
        End If

        Dim F

        F = REC2("ID")
        
      
      If KARFOTO = 1 Then
            Gdb.Execute "UPDATE MEM SET EPEL=" + str(DUM - 1) + " WHERE ID=" + str(F)
      
      
      ElseIf KARFOTO = -1 Then
            ' NA MHN ANEBAZEI
      
      Else

         If IsNull(REC2("epel")) Then
                Gdb.Execute "UPDATE MEM SET EPEL=1 WHERE ID=" + str(F)
         Else
             Gdb.Execute "UPDATE MEM SET EPEL=EPEL+1 WHERE ID=" + str(F)
         End If
      End If





      If DUM > 0 Then
         REC2.Update
      End If

        'REC2.Close

        'ƒ…¡¬¡∆Ÿ ‘«Õ –—œ«√œ’Ã≈Õ« «Ã≈—œÃ«Õ…¡ √…¡ Õ¡ Ã«Õ ≈◊Ÿ «Ã≈—œÃ«Õ…¡ Ã… —œ‘≈—« ‘«” ‘≈À≈’‘¡…¡”  ¡‘¡◊Ÿ—«”«”

        'REC2.Close

        Dim REC3 As New ADODB.Recordset

220     REC3.Open "SELECT HM2 FROM EID WHERE KOD='" + "!" + Right("00000" + LTrim(str(REC2("EPEL") - 1)), 5) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

230     If REC3.EOF Then

            Dim ANS As Integer

240         ANS = MsgBox("–—œ”œ◊« ƒ≈Õ ’–¡—◊≈… « –—œ«√œ’Ã≈Õ«  ¡‘¡◊Ÿ—«”« TOY ¬…¬À…œ’. NA ”’Õ≈◊…”Ÿ;", vbYesNo)

250         If ANS = vbNo Then

260             End

            End If

        Else

270         If REC3(0) > Now Then
280             MsgBox "« ”«Ã≈—…Õ« «Ã≈—œÃ«Õ…¡ ≈…Õ¡… Ã… —œ‘≈—« ¡–œ ‘«Õ ‘≈À≈’‘¡…¡  ¡‘¡◊Ÿ—«”«"

290             End

            End If
        End If

300     FIND_NUM = "!" + Right("00000" + LTrim(str(REC2("EPEL"))), 5)

        '<EhFooter>
        Exit Function

FIND_NUM_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.FIND_NUM " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.FIND_NUM " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub Command2_Click()

        'HEADER
        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>
        Dim STR_EKT As String

        Dim ANS

100     ANS = MsgBox("≈…”¡… ”…√œ’—œ”;", vbYesNo, "≈ ‘’–Ÿ”« ≈–… ≈÷¡À…ƒ¡”")

110     If ANS = vbNo Then Exit Sub

120     STR_EKT = "¡/¡  HME—œÃ«Õ…¡ " + " " + Left("E–ŸÕ’Ã…¡ –≈À¡‘«" + Space(25), 25) + " " + Left("≈–¡√√≈ÀÃ¡" + Space(20), 20) + " " + Left("ƒ…≈’»’Õ”«" + Space(22), 22) + " " + Left("–≈—…√—¡÷« Ã«◊/‘œ”" + Space(25), 25) + " " + Left("MON" + Space(4), 3) + " " + Format("–œ”", "#0") + " " + "” œ–.–¡—¡À¡¬«”"

130     Printer.FontName = "Courier New"
140     Printer.FontSize = 8

150     Open "LPT1" For Output As 1
160     Print #1, Chr(15) + to437(STR_EKT)
170     Close 1

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Nea_eggrafh_Click()

        Nea_eggrafh.Enabled = False
   
FramePelath.Visible = True
FrameApokatastasi.Visible = False
KATAXORHSH.Enabled = True



        '<EhHeader>
        On Error GoTo Nea_eggrafh_Click_Err

        '</EhHeader>

100     EPO.Text = ""
110     EPA.Text = ""
120     DIE.Text = " "
130     THL.Text = " "

        '<EhFooter>
        Exit Sub

Nea_eggrafh_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command4_Click()

        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>
        Dim C1, c2

        Dim ANS

100     ANS = MsgBox("≈ﬂÛ·È Ûﬂ„ÔıÒÔÚ;", vbYesNoCancel)

110     If ANS = vbYes Then
            'ok
        Else

            Exit Sub

        End If

120     C1 = InputBox("‰˘ÛÂ ÙÔÌ Í˘‰ÈÍ¸")

'130     Adodc2.ConnectionString = gConnect
'140     Adodc2.RecordSource = "select KOD,CH2,ONO,FPA,MEMO,MON,CH1,NUM1,CH3,CH4,CH5,CH6,HM1 AS [HM.≈–…” ],HM2,HM3 AS [«Ã.≈–…”‘—] from EID WHERE KOD='" + c1 + "'"
'150     Adodc2.Refresh

         mrefresh

160     If Adodc2.Recordset.EOF Then
170         MsgBox "ƒ≈Õ ’–¡—◊≈… œ  Ÿƒ… œ”"

            Exit Sub

        End If

180     If Not IsNull(Adodc2.Recordset("«Ã.≈–…”‘—")) Then
190         MsgBox "≈◊≈… –¡—¡ƒœ»≈… ‘œ Ã«◊¡Õ«Ã¡ ”‘…” " + Format(Adodc2.Recordset("«Ã.≈–…”‘—"), "DD/MM/YYYY")

            Exit Sub

        End If

200     c2 = InputBox(Adodc2.Recordset("CH2") + "**" + Adodc2.Recordset("CH1"), "ƒŸ”≈ ◊—≈Ÿ”«")

210     Gdb.Execute "UPDATE EID SET HM1=GETDATE(), LTI5=" + str(Val(c2)) + " WHERE KOD='" + C1 + "'"

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command5_Click()

        '<EhHeader>
        On Error GoTo Command5_Click_Err

        '</EhHeader>

        Dim STR_EKT As String

100     STR_EKT = Adodc2.Recordset("kod") + " " + Format(Adodc2.Recordset("hm2"), "DD/MM/YYYY") + " "

110     STR_EKT = STR_EKT + Left(Adodc2.Recordset("ch2") + Space(25), 25) + " " + Left(Adodc2.Recordset("ch4") + Space(20), 20)

120     STR_EKT = STR_EKT + " " + Left(Adodc2.Recordset("ch3") + Space(22), 22) + " " + Left(Adodc2.Recordset("ono") + Space(25), 25) + " " + Left(Adodc2.Recordset("mon") + Space(4), 3) + " " + Format(1, "#0") + " " + Left(Adodc2.Recordset("ch1") + Space(15), 15)

        'Printer.FontName = "Courier New"
        'Printer.FontSize = 8

130     Open "LPT1" For Output As 1
140     Print #1, Chr(15) + to437(STR_EKT)
150     Close 1

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command6_Click()

        '<EhHeader>
        On Error GoTo Command6_Click_Err

        '</EhHeader>
        Dim N

        Dim L

        '-------------------------------  ‰.ÔÛÔÙÈÍÁÚ ·Ò·Î·‚ÁÚ -------------------------------------
100     N = FIND_NUM(0, 0, 0)
110     N = mID(N, 2, 6)
120     N = InputBox("ƒ˛ÛÙÂ ÙÔÌ Â¸ÏÂÌÔ ·ÒÈËÏ¸ ƒ.–œ”œ‘.–¡—¡À¡¬«” Ôı Ë· Ùı˘ËÂÈ", , N)
130     L = N
140     N = FIND_NUM(L, 0, 1)    'KARFOTO TO NOYMERO

        '-------------------------------  ‰· -------------------------------------

150     N = FIND_NUM(0, 1, 0)
160     N = mID(N, 2, 6)
170     N = InputBox("ƒ˛ÛÙÂ ÙÔÌ Â¸ÏÂÌÔ ·ÒÈËÏ¸ „È· Ù· ƒ¡ –≈À¡‘ŸÕ", , N)
180     L = N
190     N = FIND_NUM(L, 1, 1)    'KARFOTO TO NOYMERO

        '<EhFooter>
        Exit Sub

Command6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.Command6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.Command6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command7_Click()

        '<EhHeader>
        On Error GoTo Command7_Click_Err

        '</EhHeader>


     F_XBARCODE = Val(FINDPARAMETROI(1, "PARALABH", "F_XBARCODE", "348", "barcode ◊ "))   'posa psifia tha exei h kathe seira
     F_YBARCODE = Val(FINDPARAMETROI(1, "PARALABH", "F_YBARCODE", "824", "barcode Y "))   'posa psifia tha exei h kathe seira




        Dim r2 As New ADODB.Recordset

100     r2.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic

110     C1 = InputBox("‰˘ÛÂ ÙÔÌ Í˘‰ÈÍ¸")



        mrefresh
        

150     If Adodc2.Recordset.EOF Then
160         MsgBox "ƒ≈Õ ’–¡—◊≈… œ  Ÿƒ… œ”"
170         r2.Close

            Exit Sub

        End If

180     For Each PR In Printers

190         If InStr(PR.DeviceName, "SATO") > 0 Then
200             Set Printer = PR

                Exit For

            End If

        Next

        'R("

        '       Print #1, R("pelono")
        '      Print #1, R("pelepa")

        Dim ZEBRA As Integer

210     ZEBRA = 1

220     If Len(Dir("C:\MERCVB\ZEBRA.TXT")) > 1 Then
230         ZEBRA = 1
        End If

        Dim c As Integer

240     c = F_ArPerEtik

        Dim W As Integer

250     For W = 1 To 2

260         If ZEBRA = 1 Then

270             Label9.Caption = "--"
280             Printer.CurrentY = F_YBARCODE

290             BarCode1.SymbologyID = CODE128    'CODE39

300             BarCode1.DataToEncode = Trim(C1)

310             Printer.PaintPicture BarCode1.Picture, F_XBARCODE, F_YBARCODE, , 300  ' ORIZ,KATHETA
            Else
320             Printer.FontName = "128"
330             Printer.FontSize = 6
340             Printer.CurrentY = 10
350             Printer.Print Tab(c + 3); Adodc2.Recordset("kod")
            End If

360         Printer.CurrentY = 50

370         Printer.FontBold = False    '  True
380         Printer.FontSize = 6
390         Printer.FontUnderline = False    'True
400         Printer.FontName = "Arial Greek"

410         Printer.FontSize = 7
420         Printer.Print Tab(c + 1); r2("PELONO")  ' "À·„‹ÍÁÚ Computers"
430         Printer.FontSize = 7
440         Printer.Print Tab(c + 1); "ÙÁÎ." + r2("PELEPA")  '25210-22225"   'CommonDialog1'.FontName

450         Printer.FontUnderline = False

460         Printer.CurrentY = 360

470         Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
480         Printer.FontBold = True

490         Printer.FontSize = 10
500         Printer.Print Tab(c); Adodc2.Recordset("kod");

510         Printer.CurrentY = 390
520         Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
530         Printer.FontBold = False

540         Printer.FontSize = 7
550         Printer.Print Tab(c + 30); Adodc2.Recordset("HM2")

560         Printer.CurrentY = 530  '580

570         Printer.FontSize = 8
580         Printer.FontBold = True
590         Printer.Print Tab(c); Left(Adodc2.Recordset("ch2"), 22) + " " + Adodc2.Recordset("ch5")

600         Printer.FontName = "Courier New"
610         Printer.FontSize = 7
620         Printer.Print Tab(c); Left(Adodc2.Recordset("ono"), 40)
630         Printer.Print Tab(c); Left(Adodc2.Recordset("memo"), 35)
640         Printer.Print Tab(c); mID(Adodc2.Recordset("memo"), 36, 40)

            ' Printer.Print
            ' Printer.Print Tab(c); Now
650         Printer.EndDoc

660     Next W

        'GoTo 10

670     r2.Close

680     Printer.EndDoc

        '<EhFooter>
        Exit Sub

Command7_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.Command7_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.Command7_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub daPelath_Click()

        '---- ƒ¡ –≈À¡‘« -------------------------------------------
        'ƒ«Ã…œ’—√œ’Ã≈ ≈Õ¡ RECORD ”‘œ EGGTIM KAI STO EID
        ' KAI ‘’–ŸÕœ’Ã≈
        '<EhHeader>
        On Error GoTo daPelath_Click_Err

        '</EhHeader>

        Dim x As String, sql As String

        Dim N As Integer

        Dim PR, FOUND

100     FOUND = False

110     For Each PR In Printers

120         If InStr(PR.DeviceName, "POSPARALABH") > 0 Then
130             FOUND = True

                Exit For

            End If

        Next

140     If Not FOUND Then
150         MsgBox "ƒ≈Õ ’–¡—◊≈… ”≈ ¡’‘«Õ ‘«Õ »≈”« œ ≈ ‘’–Ÿ‘«” √…¡ ‘…” –¡—¡À¡¬≈”"
            '   Unload Me
            '  Exit Sub
        End If

        Dim PELATHS

160     If EPO.Text = Adodc1.Recordset("EPO") Then
170         PELATHS = Adodc1.Recordset("KOD")
        Else
180         MsgBox "–—≈–≈… Õ¡ ’–¡—◊≈… ”‘œ ¡—◊≈…œ œ –≈À¡‘«”"

            Exit Sub

        End If

190     If Len(Trim(PELATHS)) < 4 Then
200         MsgBox "ƒ≈Õ ¬—…” ≈‘¡… œ  Ÿƒ… œ” ‘œ’ –≈À¡‘«"

            Exit Sub

        End If

210     If Len(Trim(EPO.Text)) = 0 Then
220         MsgBox "ƒÂÌ ÛıÏÎÁÒ˛Û·ÙÂ ¸ÌÔÏ·"

            Exit Sub

        End If

230     If Len(Trim(ono.Text)) = 0 Then
240         MsgBox "ƒÂÌ ÛıÏÎÁÒ˛Û·ÙÂ Âﬂ‰ÔÚ"

            Exit Sub

        End If

250     If Len(Trim(da.Text)) = 0 Then
260         MsgBox "ƒÂÌ ÛıÏÎÁÒ˛Û·ÙÂ ·ÒÈË.‰ÂÎÙﬂÔı"

            Exit Sub

        End If

        'ƒ«Ã…œ’—√…¡ ‘œ’  Ÿƒ… œ’ ¡–œ»« «”

270     x = FIND_NUM(0, 1, 0)    ' ¬—…” Ÿ ‘œ ‘≈À≈’‘¡…œ Õœ’Ã≈—œ KAI
        'ƒ…¡¬¡∆Ÿ ‘«Õ –—œ«√œ’Ã≈Õ« «Ã≈—œÃ«Õ…¡ √…¡ Õ¡ Ã«Õ ≈◊Ÿ «Ã≈—œÃ«Õ…¡ Ã… —œ‘≈—« ‘«” ‘≈À≈’‘¡…¡”  ¡‘¡◊Ÿ—«”«”

        Dim R As New ADODB.Recordset

        'R.Open "SELECT KOD FROM EID WHERE KOD LIKE '" + Left(X, 2) + "%' ORDER BY KOD", Gdb, adOpenDynamic, adLockOptimistic
        'R.MoveLast
        'X = Left(R(0), 2) + Format(Val(mID(R(0), 3, 5)) + 1, "0000")

280     sql = "INSERT INTO EID (KOD,ONO,FPA,MEMO,MON,CH1,NUM1,CH2,CH3,CH4,CH5,CH6,HM2) VALUES (" & "'" + x + " ','" + ono.Text + "',2,'" + BLABH.Text + "'," & "'" + MON.Text + "','" + SKOPOS.Text + "'," + POSO.Text + "," & "'" + Left(EPO.Text, 20) + "','" + Left(DIE.Text, 20) + "','" + EPA.Text + "','" + THL.Text + "','" + PELATHS + "'," & "'" + Format(DTPicker1, "mm/dd/yyyy") + "')"

        On Error GoTo DAERR

290     Gdb.Execute sql, N

300     If N = 0 Then
310         MsgBox " ‰ÂÌ Í·Ù·˜˘ÒﬁËÁÍÂ "

            Exit Sub

        End If

        Dim DUM As String

320     DUM = FIND_NUM(1, 1, 0)    ' Í·Ù·˜˘Ò˛ ÙÔ Â¸ÏÂÌÔ ÌÔ˝ÏÂÒÔ

        'ƒ«Ã…œ’—√…¡ ‘«”  …Õ«”«” ¡–œ»« «” (≈GGTIM)

330     sql = "INSERT INTO EGGTIM (HME,KODE,ONOMA,FPA,EIDOS,APOT,POSO,XRE,PELKOD,ATIM) VALUES (" & "'" + Format(DTPicker1.Value, "MM/DD/YYYY") + "','" + x + " ','" + ono.Text + "',2,'e',1," + POSO.Text + "," + POSO.Text + ",'" + PELATHS + "','·" + Format(da, "000000") + "')"
340     Gdb.Execute sql, N

350     sql = "INSERT INTO TIM (HME,EIDOS,ATIM,KPE,B_N1,AJ1,AJ2,AJ3,AJ4,AJ5,KLEIDI) VALUES (" & "'" + Format(DTPicker1.Value, "MM/DD/YYYY") + "','e','·" + Format(da, "000000") + "','" + PELATHS + "',1,0,0,0,0,0,'·" + Format(da, "000000") + "-" + PELATHS + "')"
360     Gdb.Execute sql, N

370     sql = "INSERT INTO EGG (HME,EIDOS,AIT,KOD,ATIM,XRE) VALUES (" & "'" + Format(DTPicker1.Value, "MM/DD/YYYY") + "','e','ƒ¡ –≈À¡‘«','" + PELATHS + "','·" + Format(da, "000000") + "',0)"
380     Gdb.Execute sql, N

        'rt.AddNew
        '  rt("hme") = DTPicker1.Value
        '  rt("atim") = "Œ" + m_tim
        '
        '
        '  rt("aj2") = Val(POSO.Text) / 1.19
        '  rt("aj1") = 0
        '
        '  rt("aj3") = 0
        '  rt("aj4") = 0
        '  rt("aj5") = 0
        '
        '
        '
        '  rt("fpa2") = Val(POSO.Text) - Val(POSO.Text) / 1.19
        '  rt("kpe") = M_KOD
        '  rt("aji") = Val(POSO.Text)
        '  rt("eidos") = "r"
        '  rt("EIDPAR") = ""
        '  rt("B_N1") = 1
        'rt.Update

        On Error Resume Next

        Dim STR_EKT As String

        'STR_EKT = x + " " + Format(Now, "DD/MM/YYYY") _
        '+ " " + Left(EPO.Text + Space(25), 25) _
        '+ " " + Left(EPA.Text + Space(20), 20) + _
        '" " + Left(DIE.Text + Space(22), 22) + _
        '" " + Left(ONO.Text + Space(25), 25) + _
        '" " + Left(MON.Text + Space(4), 3) + _
        '" " + Format(POSO, "#0") + _
        '" " + Left(SKOPOS.Text + Space(15), 15)
        '
        '
        '
        '
        'Open "LPT1" For Output As 1
        ' Print #1, Chr(15) + to437(STR_EKT)
        'Close 1
        '

        '
        '
        '
        '
        '
        'Printer.height = 100
        'Printer.Print STR_EKT
        ''Printer.Print STR_EKT
        ''
        'Printer.EndDoc
        ''
        '

390     For Each PR In Printers

400         If InStr(PR.DeviceName, "SATO") > 0 Then
410             Set Printer = PR

                Exit For

            End If

        Next

        '   BarCode1.ShowText = NO
        '     BarCode1.NarrowBarWidth = 0.025  '0.03
        '    BarCode1.DataToEncode = datagrid2.Text
        '    Printer.PaintPicture BarCode1.Picture, -18, 124   '-88,124

420     MsgBox " ¡‘¡◊Ÿ—«»« ≈ Ã≈ ‘œÕ  Ÿƒ… œ " + x

        Dim c As Integer

430     c = F_ArPerEtik

        Dim W As Integer

SSS:

440     For W = 1 To 1

450         Printer.FontName = "128"
460         Printer.FontSize = 6
470         Printer.Print Tab(8 + c); x

480         Printer.CurrentY = 100

490         Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
500         Printer.FontBold = True
510         Printer.FontSize = 10
520         Printer.Print Tab(c); x

530         Printer.CurrentY = 600

540         Printer.FontSize = 8
550         Printer.FontBold = True
560         Printer.Print Tab(c); EPO.Text + " " + THL.Text
            'Printer.Print

570         Printer.FontName = "Courier New"
580         Printer.FontSize = 8
590         Printer.Print Tab(c); Left(BLABH.Text, 40)
            ' Printer.Print Tab(c); Mid(BLABH.Text, 41, 40)

            'Printer.Print
600         Printer.Print Tab(c); "‰· " + da.Text + " " + Format(DTPicker1, "dd/mm/yyyy")
610         Printer.EndDoc

        Next

        'GoTo SSS

620     Printer.EndDoc

630     Unload Me

        Exit Sub

DAERR:
640     MsgBox "¡‰ıÌ·Ïﬂ· Í·Ù·˜˛ÒÁÛÁÚ . " + Err.Description

650     Resume Next

        '<EhFooter>
        Exit Sub

daPelath_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.daPelath_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.daPelath_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub ÔÎ‰TDBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim m As String


      m = TDBGrid.columns(0).Text
      DoEvents
      
           m = TDBGrid.columns(0).Text
      
      
      
'
'
'   'DataGrid2.CellBackColor = vbBlue
'
'
'
'    If datagrid2.row = 1 Then
'
'
'        If Adodc2.Recordset.sort = "[" & datagrid2.TextMatrix(0, datagrid2.Col) & "] asc" Then   ' strSort
'           Adodc2.Recordset.sort = "[" & datagrid2.TextMatrix(0, datagrid2.Col) & "] desc"   ' strSo
'        Else
'           Adodc2.Recordset.sort = "[" & datagrid2.TextMatrix(0, datagrid2.Col) & "] asc"   ' strSort
'        End If
'        Adodc2.Refresh
'
'        datagrid2.Refresh
'
'
'    End If
'
     
        
        
        
       
       
       Dim R As New ADODB.Recordset
       R.Open "select top 1 * from EID WHERE KOD = '" + Trim(m) + "'", Gdb, adOpenDynamic, adLockOptimistic
       
       If R.EOF Then
        'TIPOTA
       Else
           BLABH.Text = R!memo
           EPO.Text = CNull(R!CH2)
           THL.Text = CNull(R!ch5)
           ono.Text = CNull(R!ono)
           Noumero.Text = CNull(R!kod)
           
           
           
       End If
       
        m = Trim(m) + ".jpg"
       
       On Error GoTo tipota ' Resume Next
       foto2.Picture = LoadPicture("c:\mercvb\images\" + m)
            
       StretchSourcePictureFromPicture Me.foto2.Picture, Me.foto2




Exit Sub




tipota:

foto2.Picture = LoadPicture()


End Sub

Private Sub Form_Click()

        '<EhHeader>
        On Error GoTo Form_Click_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "PARALABH"
110     PARAMETROI.SHOW 1

        '<EhFooter>
        Exit Sub

Form_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.Form_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.Form_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>


'FrameApokatastasi.Visible = True
FramePelath.Top = FrameApokatastasi.Top


100    '·Ì·ÎÔ„ÈÍÁ ÔËÔÌÁ
        mForm_Load Me, fh, fw, ft, fl

110     DTPicker1.Value = Now

120     paralabh.Picture = LoadPicture(gPicture)


 cmdUpdate.Enabled = False



 F_ENHM_EGGTIM = Val(FINDPARAMETROI(1, "PARALABH", "F_ENHM_EGGTIM", "0", "1=ƒ«Ã…œ’—√≈…  …Õ«”« Ã≈ ‘«Õ  ¡‘¡◊Ÿ—«”« ¬À¡¬«” "))    '

130     F_XBARCODE = Val(FINDPARAMETROI(1, "PARALABH", "F_XBARCODE", "348", "barcode ◊ "))   'posa psifia tha exei h kathe seira
140     F_YBARCODE = Val(FINDPARAMETROI(1, "PARALABH", "F_YBARCODE", "824", "barcode Y "))   'posa psifia tha exei h kathe seira

        Dim R As New ADODB.Recordset

150     R.Open "SELECT *FROM PINAKES WHERE TYPOS=41 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

160     Do While Not R.EOF

170         If R("typos") = 41 Then
180             ono.AddItem R("PERIGRAFH")
            End If

190         R.MoveNext
        Loop

200     R.Close

210     F_ArPerEtik = Val(FINDPARAMETROI(1, "PARALABH", "F_ArPerEtik", "2", "¡ÒÈÛÙÂÒ¸ ÂÒÈË˛ÒÈÔ ÂÙÈÍ›ÙÙ·Ú")) 'posa psifia tha exei h kathe seira

220     F_BIBLIO_EIS = Val(FINDPARAMETROI(1, "PARALABH", "F_BIBLIO_EIS", "0", "Õ¡ ‘’–ŸÕ≈… ”‘œ ¬…¬À…œ=1 œ◊…=0")) 'posa psifia tha exei h kathe seira

        'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", "NormalOffice2007.ini"
        '     ' SkinFramework.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

        'Dim R As New ADODB.Recordset, ll
        'R.Open "SELECT TOP 1 * FROM PEL", Gdb, adOpenForwardOnly, adLockReadOnly
        'll = R.FIELDS("KOD").DefinedSize
        'R.Close
        ' If ll < 13 Then
        '
        '   Gdb.Execute " "
        '   Gdb.Execute "ALTER TABLE PEL ALTER COLUMN KOD CHAR(13)"
        ' End If
        '
        'R.Open "SELECT TOP 1 * FROM EGG", Gdb, adOpenForwardOnly, adLockReadOnly
        'll = R.FIELDS("KOD").DefinedSize
        'R.Close
        ' If ll < 13 Then
        '   Gdb.Execute "ALTER TABLE EGG ALTER COLUMN KOD CHAR(13)"
        ' End If
        '
        'R.Open "SELECT TOP 1 * FROM EGGTIM", Gdb, adOpenForwardOnly, adLockReadOnly
        'll = R.FIELDS("PELKOD").DefinedSize
        'R.Close
        ' If ll < 13 Then
        '   Gdb.Execute "ALTER TABLE EGGTIM ALTER COLUMN PELKOD CHAR(13)"
        ' End If
        '
        '
        'R.Open "SELECT TOP 1 * FROM TIM", Gdb, adOpenForwardOnly, adLockReadOnly
        'll = R.FIELDS("KPE").DefinedSize
        'R.Close
        ' If ll < 13 Then
        '   Gdb.Execute "ALTER TABLE TIM ALTER COLUMN KPE CHAR(13)"
        ' End If

230     Adodc1.ConnectionString = gConnect
240     Adodc1.RecordSource = "select top 10 * from PEL"
250     Adodc1.Refresh



mrefresh
   
   
   
   
     'ƒ«Ã…œ’—√…¡ ‘œ’  Ÿƒ… œ’ ¡–œ»« «”

     x = FIND_NUM(0, 0, -1)    ' ¬—…” Ÿ ‘œ ‘≈À≈’‘¡…œ Õœ’Ã≈—œ XVRIS NA ENHMERVSV
        'ƒ…¡¬¡∆Ÿ ‘«Õ –—œ«√œ’Ã≈Õ« «Ã≈—œÃ«Õ…¡ √…¡ Õ¡ Ã«Õ ≈◊Ÿ «Ã≈—œÃ«Õ…¡ Ã… —œ‘≈—« ‘«” ‘≈À≈’‘¡…¡”  ¡‘¡◊Ÿ—«”«”

     R.Open "SELECT * FROM EID WHERE KOD='" + x + "'", Gdb, adOpenDynamic, adLockOptimistic

     If Not R.EOF Then
         MsgBox "’‹Ò˜ÂÈ Ô Í˘‰ÈÍ¸Ú " + x

            Exit Sub

        End If
   
    Noumero.Text = x
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 14235, 9990, fh, fw, ft, fl

End Sub

Private Sub ono_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo ONO_KeyUp_Err

        '</EhHeader>
100     If KeyCode = 13 Then
110         BLABH.SetFocus
        End If

        '<EhFooter>
        Exit Sub

ONO_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.ONO_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.ONO_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub TDBGrid_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

 m = TDBGrid.columns(0).Text




End Sub

Private Sub TDBGrid_Click()
   m = TDBGrid.columns(0).Text
      
FramePelath.Visible = False
FrameApokatastasi.Visible = True

End Sub

Private Sub TDBGrid_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo TDBGrid1_KeyUp_Err

        '</EhHeader>

100     If KeyCode = 13 Then
110         ono.SetFocus
        End If

        '<EhFooter>
        Exit Sub

TDBGrid1_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.TDBGrid1_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.TDBGrid1_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub TDBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

        '<EhHeader>
        On Error GoTo TDBGrid1_RowColChange_Err

        '</EhHeader>
        
        cmdUpdate.Enabled = True
        
        
        
        

100     If Adodc2.Recordset.EOF Then

            Exit Sub

        End If
        
        
        m = TDBGrid.columns(0).Text

110     EPO = CNull(Adodc2.Recordset("CH2"))

        'If IsNull(Adodc1.Recordset("DIE")) Then
        '   DIE = ""
        '  Else
120    ' DIE = CNull(Adodc1.Recordset("DIE"))

130    ' EPA = CNull(Adodc1.Recordset("EPA"))
140    ' THL = CNull(Adodc1.Recordset("THL"))

         If IsNull(Adodc2.Recordset("HM1")) Then
             HM1 = Now
         Else
             HM1 = Adodc2.Recordset("HM1")
         End If
         
          
         If IsNull(Adodc2.Recordset("HM3")) Then
             HM3 = Now
         Else
             HM3 = Adodc2.Recordset("HM3")
         End If
         
         kostos.Text = nNull(Adodc2.Recordset("num2"))

'         If IsNull(Adodc1.Recordset("ONO2")) Then
'            ono2.Text = ""
'         Else
            ONO2.Text = CNull(Adodc2.Recordset("ONO2"))
         'End If
         
         
         
         
         TEXNIKOS.Text = nNull(Adodc2.Recordset("num3"))
         
         
         
         
         
         
         
         
   '         Dim R As New ADODB.Recordset
     '  R.Open "select top 1 * from EID WHERE KOD = '" + Trim(m) + "'", Gdb, adOpenDynamic, adLockOptimistic
         
       
       
     '  If R.EOF Then
     '   'TIPOTA
     '  Else
           BLABH.Text = R!memo
           EPO.Text = CNull(Adodc2.Recordset("CH2"))
           THL.Text = CNull(Adodc2.Recordset("CH5")) 'CNull(R!ch5)
           ono.Text = CNull(Adodc2.Recordset("ONO")) ' CNull(R!ONO)
           Noumero.Text = CNull(Adodc2.Recordset("KOD")) ' CNull(R!kod)
           
           
           
    '   End If
       
        m = Trim(CNull(Adodc2.Recordset("KOD"))) + ".jpg"
       
        ' Resume Next
        
        On Error GoTo NOFOTO
        
       foto2.Picture = LoadPicture("c:\mercvb\images\" + m)
            
       StretchSourcePictureFromPicture Me.foto2.Picture, Me.foto2
         
        Exit Sub
NOFOTO:
         
       foto2.Picture = LoadPicture()
         
         
         
         
         
         
         
         
         
         
         
         

        '<EhFooter>
        Exit Sub

TDBGrid1_RowColChange_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.TDBGrid1_RowColChange " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.TDBGrid1_RowColChange " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Text8_KeyUp_Err

        '</EhHeader>
100     If KeyCode = 13 Then
110         TDBGrid1.SetFocus
        End If

        '<EhFooter>
        Exit Sub

Text8_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.Text8_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.Text8_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text8_LostFocus()

        '<EhHeader>
        On Error GoTo Text8_LostFocus_Err

        '</EhHeader>

100     Adodc1.RecordSource = "SELECT TOP 10 EPO,DIE,* FROM PEL WHERE EPO LIKE '" + (SameLetters(Text8.Text)) + "%'"
110     Adodc1.Refresh
120     TDBGrid1.columns(0).Width = 3000
130     TDBGrid1.columns(1).Width = 3000
140     TDBGrid1.Splits(0).AlternatingRowStyle = True
150     TDBGrid1.Splits(0).AnchorRightColumn = True

160     If Adodc1.Recordset.EOF Then

            Exit Sub

        End If

170     EPO = Adodc1.Recordset("EPO")
180     DIE = CNull(Adodc1.Recordset("DIE"))
190     EPA = CNull(Adodc1.Recordset("EPA"))
200     THL = CNull(Adodc1.Recordset("THL"))

210     ono.Text = ono.List(0)

        'SELECT * FROM  EID WHERE CH2='"+

        ' Adodc2.ConnectionString = gConnect
'220     Adodc2.RecordSource = "select top 200 KOD,CH2,ONO,HM2,MEMO,MON,CH1,NUM1,CH3,CH4,CH5,CH6 from EID WHERE CH2 LIKE '" + (SameLetters(Text8.Text)) + "%' ORDER BY HM2 DESC "
'230     Adodc2.Refresh


mrefresh


        '<EhFooter>
        Exit Sub

Text8_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.paralabh.Text8_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.paralabh.Text8_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Sub LAG_ETIK()
'
'End Sub



Sub mrefresh()

'120     Adodc2.ConnectionString = gConnect
'130     Adodc2.RecordSource = "select KOD,CH2,ONO,FPA,MEMO,MON,CH1,NUM1,CH3,CH4,CH5,CH6,HM1 AS [HM.≈–…” ],HM2,HM3 AS [«Ã.≈–…”‘—] from EID WHERE KOD='" + c1 + "'"
'140     Adodc2.Refresh

260     Adodc2.ConnectionString = gConnect
270     Adodc2.RecordSource = "select top 200  KOD,CH2 AS [–≈À¡‘«”],HM2 AS [«Ã≈—],CH5 AS [‘«À≈÷ŸÕœ],ONO AS [”’” ≈’«],MEMO AS [¬À¡¬«],ONO2 AS [≈—√¡”…¡],HM1 AS [HM.≈–…” ],HM3 AS [«Ã.≈–…”‘—],CH1,MON,NUM1,CH3,CH4,CH6,NUM2,NUM3,ONO2,CH2,ONO from EID ORDER BY HM2 DESC "
280     Adodc2.Refresh

  TDBGrid.columns(0).Width = 900

  TDBGrid.columns(1).Width = 900
  
  TDBGrid.columns(2).Width = 2000 'ONOMA
   TDBGrid.columns(3).Width = 1800    'HME
   TDBGrid.columns(4).Width = 1420 'THL
   
   TDBGrid.columns(5).Width = 2300  'PERIGR
   TDBGrid.columns(6).Width = 6200 ' BLABH
   
   
   
   'DataGrid2.ColWidth(6) = 420
   
   
   
   
  ' DataGrid2.ColWidth(5) = 420


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

110         If Trim(Col.FilterText) <> "" Then
120             N = N + 1

130             If N > 1 Then
140                 tmp = tmp & " AND "
                End If

150             If Adodc2.Recordset(Col.ColIndex).Type = 5 Then    '
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

Private Sub cmdClearFilter_Click()

        '<EhHeader>
        On Error GoTo cmdClearFilter_Click_Err

        '</EhHeader>

        'Clears filter from grid

100     For Each Col In TDBGrid.columns

110         Col.FilterText = ""

120     Next Col

130     Adodc2.Recordset.Filter = adFilterNone

        '<EhFooter>
        Exit Sub

cmdClearFilter_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.cmdClearFilter_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.cmdClearFilter_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

