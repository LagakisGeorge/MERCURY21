VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form psifiakopelat 
   BackColor       =   &H00C0C000&
   Caption         =   "ÿÁˆÈ·Í¸ –ÂÎ·ÙÔÎ¸„ÈÔ"
   ClientHeight    =   12525
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   21690
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12525
   ScaleWidth      =   21690
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "¡Õ‘…”‘œ…◊…”« ‘…ÃœÀ.Ã≈ Ã¡— "
      Height          =   375
      Left            =   0
      TabIndex        =   50
      Top             =   9360
      Width           =   3375
   End
   Begin VB.CommandButton UploadCar 
      Caption         =   "¡ÌÂ‚·ÛÏ· ·ıÙÔÍÈÌﬁÙÔı ÛÂ ¯ÁˆÈ·Í¸"
      Height          =   360
      Left            =   1560
      TabIndex        =   44
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton autoClose 
      Caption         =   "¡ıÙ¸Ï·ÙÔ  ÎÂﬂÛÈÏÔ"
      Height          =   375
      Left            =   9360
      TabIndex        =   43
      Top             =   8160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox chk¬À≈–ŸœÀœ À«—ŸÃ≈Õ¡ 
      Caption         =   "¬À≈–Ÿ œÀœ À«—ŸÃ≈Õ¡"
      Height          =   375
      Left            =   6120
      TabIndex        =   40
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdOLD 
      Caption         =   "–¡À…≈” ≈√√—¡÷≈”"
      Height          =   360
      Left            =   20040
      TabIndex        =   39
      Top             =   8520
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txSXOLIA 
      Height          =   375
      Left            =   9240
      TabIndex        =   37
      Top             =   1560
      Width           =   4215
   End
   Begin VB.CommandButton cmdNEA 
      Caption         =   "Õ≈¡ ≈√√—¡÷«"
      Height          =   375
      Left            =   360
      TabIndex        =   36
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox afmonoma 
      BackColor       =   &H000080FF&
      Height          =   495
      Left            =   18150
      TabIndex        =   34
      Top             =   840
      Width           =   1695
   End
   Begin VB.ListBox ListEpo 
      Height          =   3570
      Left            =   16710
      TabIndex        =   33
      Top             =   1440
      Width           =   3135
   End
   Begin VB.CommandButton search 
      Caption         =   "≈ıÒÂÛÁ"
      Height          =   360
      Left            =   17310
      TabIndex        =   32
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton ccmdCloseInvoice 
      Caption         =   "œÀœ À«—Ÿ”« Ã≈ ‘…ÃœÀœ√…œ"
      Height          =   375
      Left            =   4200
      TabIndex        =   29
      Top             =   8160
      Width           =   4095
   End
   Begin VB.ComboBox InvoiceKind 
      Height          =   315
      ItemData        =   "psifiakopelat.frx":0000
      Left            =   7320
      List            =   "psifiakopelat.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   2350
      Width           =   2535
   End
   Begin VB.ComboBox providedServiceCategory 
      Height          =   315
      ItemData        =   "psifiakopelat.frx":0051
      Left            =   360
      List            =   "psifiakopelat.frx":0053
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   2350
      Width           =   6495
   End
   Begin VB.CommandButton cmdAPOXML 
      Caption         =   "APO XML"
      Height          =   360
      Left            =   10800
      TabIndex        =   24
      Top             =   13080
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton Command3 
      Caption         =   "≈Õ«Ã≈—Ÿ”« ≈–…” ≈’«” "
      Height          =   360
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2350
      Width           =   2172
   End
   Begin VB.CommandButton Command2 
      Caption         =   "INSERT INTO CARRELAT"
      Height          =   360
      Left            =   6240
      TabIndex        =   22
      Top             =   12600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "??"
      Height          =   360
      Left            =   9600
      TabIndex        =   21
      Top             =   12960
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton cmdUpdateClient 
      Caption         =   "updateClient"
      Height          =   360
      Left            =   8160
      TabIndex        =   20
      Top             =   12960
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton cmdRequestClients 
      Caption         =   "RequestClients"
      Height          =   360
      Left            =   11190
      TabIndex        =   19
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton DIAGRAFI 
      Caption         =   "ƒ…¡√—¡÷« œ◊«Ã¡‘œ”"
      Height          =   360
      Left            =   12360
      TabIndex        =   18
      Top             =   480
      Width           =   3270
   End
   Begin VB.CommandButton ananeosi 
      Caption         =   "¡Ì·Ì›˘ÛÁ"
      Height          =   360
      Left            =   8400
      TabIndex        =   17
      Top             =   480
      Width           =   3270
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Frame1"
      Height          =   2295
      Left            =   12720
      TabIndex        =   10
      Top             =   5880
      Width           =   7215
      Begin VB.TextBox TXTPOSO 
         Height          =   375
         Left            =   4800
         TabIndex        =   30
         Text            =   " "
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox TAMFOROL 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdFIM 
         Caption         =   "¡Õ‘…”‘œ…◊…”«"
         Height          =   360
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox chkTameiakhForologikos 
         Caption         =   "‘¡Ã≈…¡ «/÷œ—œÀœ√… œ”"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.TextBox FIMAA 
         Height          =   285
         Left            =   4800
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker FIMDATE 
         Height          =   285
         Left            =   4800
         TabIndex        =   13
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         _Version        =   393216
         Format          =   456785921
         CurrentDate     =   45743
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "–œ”œ ‘¡Ã≈…¡ «”/÷œ—œÀ"
         Height          =   195
         Left            =   2760
         TabIndex        =   31
         Top             =   1320
         Width           =   1950
      End
      Begin VB.Label lbl«Ã≈—¡–œƒ≈…Œ«” 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã≈—.¡–œƒ≈…Œ«”"
         Height          =   195
         Left            =   2760
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblAA¡–œƒ≈…Œ«” 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AA ¡–œƒ≈…Œ«”"
         Height          =   195
         Left            =   2760
         TabIndex        =   11
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdCLOSEcLIENT 
      Caption         =   "œÀœ À«—Ÿ”« ≈–…” ≈’«” ÷ÔÒÔÎ./‘·ÏÂÈ·Íﬁ/ƒ˘ÒÂ‹Ì"
      Height          =   360
      Left            =   12630
      TabIndex        =   9
      Top             =   8160
      Width           =   7215
   End
   Begin VB.ComboBox VFACT 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "psifiakopelat.frx":0055
      Left            =   6600
      List            =   "psifiakopelat.frx":0057
      TabIndex        =   8
      Top             =   1560
      Width           =   2535
   End
   Begin VB.ComboBox VCAT 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "psifiakopelat.frx":0059
      Left            =   3960
      List            =   "psifiakopelat.frx":005B
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdKatax 
      Caption         =   "≈ÌÁÏ›Ò˘ÛÁ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13560
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox ARKYKL 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc cars 
      Height          =   495
      Left            =   7320
      Top             =   13440
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
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
      Caption         =   "cars"
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
   Begin MSAdodcLib.Adodc timologia 
      Height          =   495
      Left            =   10200
      Top             =   13440
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
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
      BackColor       =   12648447
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
      Caption         =   "timologia"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
      Height          =   855
      Left            =   0
      TabIndex        =   27
      Top             =   2000
      Width           =   16575
      Begin VB.TextBox sxolia2 
         Height          =   372
         Left            =   9960
         TabIndex        =   46
         Top             =   360
         Width           =   4092
      End
   End
   Begin TrueOleDBGrid80.TDBGrid CARGrid 
      Bindings        =   "psifiakopelat.frx":005D
      Height          =   2800
      Left            =   0
      TabIndex        =   38
      Top             =   2955
      Width           =   16545
      _ExtentX        =   29184
      _ExtentY        =   4948
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
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2752"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
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
   Begin MSAdodcLib.Adodc AdoCarsOLD 
      Height          =   495
      Left            =   12960
      Top             =   13440
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
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
      Caption         =   "carsOLD"
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
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Bindings        =   "psifiakopelat.frx":0070
      Height          =   2000
      Left            =   0
      TabIndex        =   45
      Top             =   6240
      Visible         =   0   'False
      Width           =   19785
      _ExtentX        =   34899
      _ExtentY        =   3545
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
      DataView        =   2
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
   Begin VB.CommandButton METAXRONOLOGHMENO 
      BackColor       =   &H0080C0FF&
      Caption         =   "≈Œœƒœ” œ◊«Ã¡‘œ” & –¡—¡ÃœÕ« ”≈ ≈  —≈Ãœ‘«‘¡ Ã≈‘¡◊—œÕœÀ«√«Ã≈Õœ ‘…ÃœÀœ√…œ ≈Ÿ” 15 ≈–œÃ≈Õœ’ Ã«Õ¡"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   8760
      Width           =   5052
   End
   Begin VB.CommandButton SEEKREMOTHTA 
      BackColor       =   &H0080C0FF&
      Caption         =   "œ◊«Ã¡‘¡ ”≈ ≈  —≈Ãœ‘«‘¡"
      Height          =   492
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   8760
      Width           =   1452
   End
   Begin VB.CommandButton ANTISTTIMOL 
      BackColor       =   &H0080C0FF&
      Caption         =   "¡Õ‘…”‘œ…◊…”« Ã≈ Ã¡— "
      Height          =   492
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   8760
      Width           =   1572
   End
   Begin TrueOleDBGrid80.TDBGrid TIMGrid 
      Bindings        =   "psifiakopelat.frx":0089
      Height          =   2100
      Left            =   0
      TabIndex        =   5
      Top             =   5760
      Width           =   19728
      _ExtentX        =   34793
      _ExtentY        =   3704
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
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2778"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2778"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=136,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HC0FFFF&,.bold=0,.fontsize=1200"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1200,.italic=0"
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
   Begin VB.CommandButton cmdSYSX 
      Caption         =   "ANTI”TOIXI”H ME –¡—¡”‘¡‘… œ"
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Top             =   8160
      Width           =   3135
   End
   Begin VB.Label lbl¡— ’ À 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "”◊œÀ…¡"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   1
      Left            =   9360
      TabIndex        =   42
      Top             =   1200
      Width           =   870
   End
   Begin VB.Label cURRTIM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   360
      TabIndex        =   41
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   10200
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label LSEARCH 
      BackStyle       =   0  'Transparent
      Caption         =   "œÌÔÏ·/¡÷Ã"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   16680
      TabIndex        =   35
      Top             =   840
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   120
      Top             =   840
      Width           =   15495
   End
   Begin VB.Label lblVFACT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ã¡— ¡ ¡’‘."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   6720
      TabIndex        =   4
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label lblVCAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ¡‘«√œ—…¡"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   4320
      TabIndex        =   3
      Top             =   1200
      Width           =   1305
   End
   Begin VB.Label lbl¡— ’ À 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¡—. ’ À"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   960
   End
End
Attribute VB_Name = "psifiakopelat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f_parking_auto As Integer '= Val(FINDPARAMETROI(3, "PAR1", "f_parking_auto", "0", "ÂÈÌ·È parking ÏÂ ·ıÙ¸Ï·ÙÔ ¯ÁˆÈ·ÍÔ ÂÎ·ÙÔÎ¸„ÈÔ=1 /Ô˜È=0"))
Dim F_MESS As Integer
Dim f_name_xml   As String ' „È· Ì· ‚·Ê˘ ÔÌÔÏ· ÛÙÔ xml

'Dim gf_Bearer As String
Dim f_YPOK       As String

Dim Col          As TrueOleDBGrid80.Column

Dim ColS         As TrueOleDBGrid80.columns

Dim F_demo       As Integer

Dim PAROX_URLADD As String

Dim PAROX_USER   As String

Dim PAROX_PASS   As String

Dim PAROX_KEY    As String

Dim f_pelono, f_peldie, f_pelarit, f_pelpol, f_peltk As String

Dim F_PAROCHOS As Integer

'F_PAROCHOS = Val(FINDPARAMETROI(1, "PAR7MYDATA", "f_parochos", "0", "0=MyData 1=–·ÒÔ˜ÔÚ"))

Dim antFPA(10)

Dim timhFPA(10)        As Single
Dim F_artam As String

Dim F_TXTFILES         As String

Dim F_ID_NUMS(5000)    As Long ' ¡–œ»« ≈’≈… ‘¡ ID_NUM √…¡ Õ¡ ‘¡ »’Ã¡‘¡… œ‘¡Õ ≈—»≈… « ¡–¡Õ‘«”« ¡–œ ¡¡ƒ≈

Dim F_dokimi           As Integer ' ·ÒÈËÏ¸Ú ‰ÔÍÈÏ·ÛÙÈÍÔı user

' –Ÿ” ”‘≈ÀÕŸ ◊¡—¡ ‘«—…”Ãœ ≈”œƒŸÕ
'imports System
'imports System.Net.Http.headers
'imports System.Text
'imports System.Net.Http
'imports System.Web
'
'Namespace CSHttpClientSample
'    Module Program
'        Private Sub Main()
'            MakeRequest()
'            Console.WriteLine ("Hit ENTER to exit...")
'            Console.ReadLine()
'        End Sub
'
'        Private Async Sub MakeRequest()
'            Dim client = New HttpClient()
'            Dim queryString = HttpUtility.ParseQueryString(String.Empty)
'            client.DefaultRequestHeaders.Add("{aade-user-id}", "{aade-user-id}")
'            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "{subscription key}")
'            Dim uri = "https://mydata-dev.azure-api.net/SendExpensesClassification?" & queryString
'            Dim response As HttpResponseMessage
'            Dim byteData As Byte() = Encoding.UTF8.GetBytes("{body}")
'
'            Using content = New ByteArrayContent(byteData)
'                content.Headers.ContentType = New MediaTypeHeaderValue("< your content type, i.e. application/json >")
'                response = Await client.PostAsync(uri, content)
'            End Using
'        End Sub
'    End Module
'End Namespace

'        <taxesTotals>
'            <taxes>
'                <taxType>1</taxType>
'                <taxCategory>2</taxCategory>
'                <underlyingValue>1000.00</underlyingValue>
'                <taxAmount>200.00</taxAmount>
'            </taxes>
'        </taxesTotals>

'–Â‰ﬂÔ           ‘˝ÔÚ        ’Ô˜ÒÂ˘ÙÈÍ¸    –ÂÒÈ„Ò·ˆﬁ                  ¡Ô‰ÂÍÙ›Ú ÙÈÏ›Ú
'taxType         xs:byte        Õ·È         ≈ﬂ‰ÔÚ ÷¸ÒÔı                  ÀﬂÛÙ· ÙÈÏ˛Ì:
'                                                                    1 = –·Ò·ÍÒ·ÙÔ˝ÏÂÌÔÚ ÷¸ÒÔÚ
'                                                                    2 = ‘›ÎÁ
'                                                                    3 = ÀÔÈÔﬂ ÷¸ÒÔÈ
'                                                                    4 = ◊·ÒÙ¸ÛÁÏÔ
'                                                                    5 =  Ò·ÙﬁÛÂÈÚ
'taxCategory     xs:byte       º˜È         ·ÙÁ„ÔÒﬂ· ÷¸ÒÔı                 ≈Î‹˜ÈÛÙÁ ÙÈÏﬁ = 1
'underlyingValue xs:decimal    º˜È        ’ÔÍÂﬂÏÂÌÁ ¡Óﬂ·            ≈Î‹˜ÈÛÙÁ ÙÈÏﬁ = 0
'                                                                    ƒÂÍ·‰ÈÍ‹ ¯Áˆﬂ· = 2
'taxAmount       xs:decimal    Õ·È          –ÔÛ¸ ÷¸ÒÔı               ≈Î‹˜ÈÛÙÁ ÙÈÏﬁ = 0
'                                                                     ƒÂÍ·‰ÈÍ‹ ¯Áˆﬂ· = 2
'id              xs:byte       º˜È      ¡˝Ó˘Ì ·ÒÈËÏ¸Ú „Ò·ÏÏﬁÚ
'–·Ò·ÙÁÒﬁÛÂÈÚ:
'1) ‘Ô Â‰ﬂÔ taxCategory ÏÔÒÂﬂ Ì· ‹ÒÂÈ Í‹ËÂ ˆÔÒ‹ ÔÔÈ·‰ﬁÔÙÂ ÙÈÏﬁ ·¸ ÙÔÌ ·ÌÙﬂÛÙÔÈ˜Ô ﬂÌ·Í· ÙÔı –·Ò·ÒÙﬁÏ·ÙÔÚ ÙÔı ˆ¸ÒÔı Ôı ·Ì·ˆ›ÒÂÙ·È ÛÙÔ Â‰ﬂÔ taxType
'2) ‘Ô Â‰ﬂÔ underlyingValue ıÔ‰ÁÎ˛ÌÂÈ ÙÁÌ ·Óﬂ· ÛÙÁÌ ÔÔﬂ· Âˆ·ÒÏ¸ÊÂÙ·È Ô Ûı„ÍÂÍÒÈÏ›ÌÔÚ ˆ¸ÒÔÚ

Dim f_par1ypok         As String ' ·Ò·ÛÙ·ÙÈÍ‹ Ôı ·ÌﬁÍÔıÌ ÛÙÔ ıÔÍ/Ï· 1

Dim f_par2ypok         As String ' ·Ò·ÛÙ·ÙÈÍ‹ Ôı ·ÌﬁÍÔıÌ ÛÙÔ ıÔÍ/Ï· 1

Dim fSynt              As String

Dim fKatEsod(10)       As String

Dim fKatEXod(10)       As String

Dim FkatE3Esod(10)     As String

Dim F_E3_APOKATHG_EID  As Integer

Dim sql                As String

'Dim gdb As New ADODB.Connection
Dim MARK               As String

Dim fMydataFromEID1    As Integer

Dim P_USER             As String

Dim P_KEY              As String

Dim afmCompany         As String

Dim UrlAADE            As String

Dim url2               As String

'Dim sqldt As New ADODB.Recordset
Dim AFM                As String ' –≈À¡‘«” COUNTERPART

Dim AFM0               As String ' ≈ ƒœ‘«Ú ISSUER

Dim issueDate          As String

Dim AA                 As String

Dim invoiceType        As String

Dim f_apoDOS           As Integer

Dim paytype            As String

Dim payaji             As String

Dim totalNetValue      As String

Dim totalVatAmount     As String

Dim totalGrossValue    As String

Dim docStock           As MSXML2.DOMDocument

Dim classificationType As String

Const MinDate          As Date = #1/1/1900# '-- set MinDate as you need

'-- 2^31 is used as an offset for 32-bit-sign-integer

Dim fLong              As Long  ' ÙÔ ˜ÒÁÛÈÏÔÔÈ˘ Û·Ì ‰ÂÈÍÙÁ Ì· ÓÂÒ˘ ÙÔ update tim ÙÈ Ë· ÂÌÁÏÂÒ˘ÌÂÈ






























Dim F_ISCAR As Integer





Function fupdateclient(Message As Integer, ByVal clientID As String, _
                              ByVal ServiceType As String, _
                              ByVal completionDateTime As String, providedServiceCategory As Integer, InvoiceKind As Integer, POSO As String) As String '"202502221033
   
        '  Dim docStock As MSXML2.DOMDocument
        Dim elemRoot  As MSXML2.IXMLDOMElement

        Dim elemStock As MSXML2.IXMLDOMElement

        Dim elemField As MSXML2.IXMLDOMElement
        Dim elem1Field As MSXML2.IXMLDOMElement
        Dim elem2Field As MSXML2.IXMLDOMElement
        
        
        
        
        

        Dim i         As Integer
    
100     Set docStock = New MSXML2.DOMDocument

102     With docStock
104         .appendChild .createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
106         Set elemRoot = .createElement("dcrudt:UpdateClientDoc")
108         elemRoot.setAttribute "xmlns:dcrudt", "https://www.aade.gr/myDATA/dcrudt/v1.0"
110         elemRoot.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
112         elemRoot.setAttribute "xsi:schemaLocation", "https://www.aade.gr/myDATA/dcrudt/v1.0 updateClientType-v1.0.xsd updateClient-v1.0.xsd"

114         Set elemStock = docStock.createElement("dcrudt:updateClient")
            
            
116          Set elem2Field = docStock.createElement("dcrudt:initialDclId"): elem2Field.Text = clientID: elemStock.appendChild elem2Field
118         Set elem2Field = docStock.createElement("dcrudt:clientServiceType"): elem2Field.Text = ServiceType: elemStock.appendChild elem2Field
120         Set elem2Field = docStock.createElement("dcrudt:entryCompletion"): elem2Field.Text = "false": elemStock.appendChild elem2Field
             
             If InvoiceKind > 0 Then
                Set elem2Field = docStock.createElement("dcrudt:nonIssueInvoice"): elem2Field.Text = "false": elemStock.appendChild elem2Field
                 Set elem2Field = docStock.createElement("dcrudt:amount"): elem2Field.Text = POSO: elemStock.appendChild elem2Field
                  Set elemField = docStock.createElement("dcrudt:providedServiceCategory"): elemField.Text = providedServiceCategory: elemStock.appendChild elemField
             
             Else
                Set elem2Field = docStock.createElement("dcrudt:nonIssueInvoice"): elem2Field.Text = "true": elemStock.appendChild elem2Field
                
                 Set elemField = docStock.createElement("dcrudt:providedServiceCategory"): elemField.Text = providedServiceCategory: elemStock.appendChild elemField
                
                Set elem2Field = docStock.createElement("dcrudt:reasonNonIssueType"): elem2Field.Text = 1: elemStock.appendChild elem2Field
                 Set elem2Field = docStock.createElement("dcrudt:comments"): elem2Field.Text = sxolia2.Text: elemStock.appendChild elem2Field
                
               ' Set elem2Field = docStock.createElement("dcrudt:invoiceCounterparty"): elem2Field.Text = "000000000": elemStock.appendChild elem2Field
               ' reasonNonIssueType
             End If
             
            
             
           ' Set elem2Field = docStock.createElement("dcrudt:amount"): elem2Field.Text = 10#: elemStock.appendChild elem2Field
             
           '   Set elem2Field = docStock.createElement("dcrudt:amount"): elem2Field.Text = POSO: elemStock.appendChild elem2Field
            '  Set elemField = docStock.createElement("dcrudt:providedServiceCategory"): elemField.Text = providedServiceCategory: elemStock.appendChild elemField
            
           ' invoiceKind
            If InvoiceKind > 0 Then
                Set elem2Field = docStock.createElement("dcrudt:invoiceKind"): elem2Field.Text = InvoiceKind: elemStock.appendChild elem2Field
                 Set elem2Field = docStock.createElement("dcrudt:comments"): elem2Field.Text = sxolia2.Text: elemStock.appendChild elem2Field
            End If
            
            
           
    
            
122                        elemRoot.appendChild elemStock
  
            '  End With  /////////////////////////////////////////////////////////////////////////////////////////////
124         Set .documentElement = elemRoot
   
            
            
            
            
            
            
            
            ' With elemStock----------------------------------  100000000006453
         '   Set elem2Field = docStock.createElement("dcrudt:completionDateTime"): elem2Field.Text = "2025-05-26T18:58:53.555": elemStock.appendChild elem2Field
            'isDiffVehReturnLocation
         '    Set elem2Field = docStock.createElement("dcrudt:isDiffVehReturnLocation"): elem2Field.Text = "false": elemStock.appendChild elem2Field
            
            

            
            
            
            
            
            
            
          '  Set elem2Field = docStock.createElement("dcrudt:exitDateTime"): elem2Field.Text = "2025-05-26T18:58:53.555": elemStock.appendChild elem2Field

'118         Set elem2Field = docStock.createElement("dcrudt:clientServiceType"): elem2Field.Text = ServiceType: elemStock.appendChild elem2Field
'120         Set elem2Field = docStock.createElement("dcrudt:entryCompletion"): elem2Field.Text = "true": elemStock.appendChild elem2Field


            
          
             
130        ' Set elemField = docStock.createElement("dcrudt:providedServiceCategory"): elemField.Text = "2"

                       
        ' Set elem2Field = docStock.createElement("dcrudt:ProvidedCategoryServiceType"): elem2Field.Text = 3: elemField.appendChild elem2Field
                          

                       
        ' elemStock.appendChild elemField
           

'  Set elem2Field = docStock.createElement("dcrudt:amount"): elem2Field.Text = 10#: elemStock.appendChild elem2Field

 '  Set elem2Field = docStock.createElement("dcrudt:invoiceKind"): elem2Field.Text = "3": elemStock.appendChild elem2Field
            
'                 '------------------------------------------------------------------------------
'                  Set elemField = docStock.createElement("dcrudt:providedServiceCategory")
'
'
'                         Set elem2Field = docStock.createElement("dcrudt:ProvidedCategoryServiceType"): elem2Field.Text = "3": elemField.appendChild elem2Field
'
'                 elemStock.appendChild elemField ' elemRoot.appendChild elem1Field
'
'
'
'
'
'             elemStock.appendChild elem2Field
'
            
      '      Set elem2Field = docStock.createElement("dcrudt:amount"): elem2Field.Text = 10#: elemStock.appendChild elem2Field
            'Set elem2Field = docStock.createElement("dcrudt:completionDateTime"): elem2Field.Text = completionDateTime: elemStock.appendChild elem2Field
            
'            'ProvidedCategoryServiceType
'             Set elem2Field = docStock.createElement("dcrudt:providedServiceCategory"): elem2Field.Text = "3": elemStock.appendChild elem2Field
'             Set elem2Field = docStock.createElement("dcrudt:amount"): elem2Field.Text = 10#: elemStock.appendChild elem2Field
'             Set elem2Field = docStock.createElement("dcrudt:completionDateTime"): elem2Field.Text = completionDateTime: elemStock.appendChild elem2Field
'            ' elemStock.appendChild elemField
             '---------------------------------------------------
                    
'                providedServiceCategory
'ProvidedCategoryServiceType
'º˜È
' ·ÙÁ„ÔÒﬂ· –·ÒÂ˜¸ÏÂÌ˘Ì ’ÁÒÂÛÈ˛Ì
'1 = ∏ÎÂ„˜ÔÚ/ÂÍÙﬂÏÁÛÁ ÊÁÏÈ‹Ú
'2 = ≈Ò„·ÛﬂÂÚ ÙÈÏÔÍ·Ù·Î¸„Ôı
'3 = ≈Ò„·Ûﬂ· Í·ÙíÔÈÌ
'Digital Client REST API 7
'ÛıÏˆ˘Ìﬂ·Ú
'4 = ƒ˘ÒÂ‹Ì ıÁÒÂÛﬂ·
'5 = ÀÔÈ‹
                    
                 
                    
                    
'   <dcrudt:UpdateClientDoc xsi:schemaLocation="https://www.aade.gr/myDATA/dcrudt/v1.0 updateClient-v1.0.xsd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcrudt="https://www.aade.gr/myDATA/dcrudt/v1.0">
'    <dcrudt:updateClient>
'        <dcrudt:initialDclId></dcrudt:initialDclId>
'        <dcrudt:clientServiceType>3</dcrudt:clientServiceType>
'        <dcrudt:entryCompletion>true</dcrudt:entryCompletion>
'        <dcrudt:amount>10</dcrudt:amount>
'        <dcrudt:comments>TEST UPDATE GARAGE CLIENT SERVICE</dcrudt:comments>
'    </dcrudt:updateClient>
'</dcrudt:UpdateClientDoc>


            On Error Resume Next
            
126         FormatXmlDocument docStock ' ‚·ÊÂÈ ÍÂÌ· Ì· ÂÈÌ·È ÂıÍÔÎÔ ÛÙÔ ‰È·‚·ÛÏ·

            Dim ANS2 As Integer
            
            
        If Message = 1 Then
            
            

128           ANS2 = MsgBox(docStock.XML, vbYesNo)

             If ANS2 = vbNo Then

                 Exit Function

             End If
            
      End If
            
            
         Kill "C:\txtfiles\" + clientID + ".xml"

            On Error GoTo cmdDOMDocumentUTF8_Click_Err

         .save "C:\txtfiles\" + clientID + ".xml"
        End With

     FileCopy "C:\txtfiles\" + clientID + ".xml", "c:\" + F_TXTFILES + "\sendinv\update-aytok" + "-" + clientID + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"

        ' TO XML EINAI docStock.xml

        '--------------- ¡–œ”‘œÀ« INVOICES -----------------------------------------------------
     
        Dim MARK As String

        '    MARK = InputBox("‰˘ÛÂ Ùo mark", "Ï·ÒÍ")
     
     url2 = url2 + MARKTIM
     
     url2 = UrlAADE + "DCL/UpdateClient"

        'initialize
     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
    XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
148     XMLServer.Option(9) = 2048
150     XMLServer.Option(6) = True
    
152     XMLServer.Open "POST", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
154     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
156     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"

158     XMLServer.send docStock.XML 'TXML 'strData  '”‘≈ÀÕŸ ‘œ STRING TXML STO MYDATA  'XMLServer.send TXML '  ¡Õ »≈ÀŸ Õ¡ ‘œ ”‘≈…ÀŸ ¡–œ ¡—◊≈…œ ‘œ‘≈ ¬¡∆Ÿ strData
160     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
162     v = XMLServer.responseText
        '--------------- ¡–œ”‘œÀ« INVOICES -----------------------------------------------------
 If Message = 1 Then
164     MsgBox v
  End If
        Dim CC As String

166     CC = takefromxml(v)
168     fupdateclient = CC

170     TIMOL = Replace(TIMOL, "/", "-")

        Dim mfile As String:

172     mfile = "c:\" + F_TXTFILES + "\sendinv\apant-update-aytok" + "-" + clientID + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"
  
        ' Dim mfileAsk As String
        ' mfileAsk = "C:\txtfiles\" + clientID + ".xml"
174     Debug.Print v
176     Open mfile For Output As #8
178     Print #8, url2 + "--" + v
180     Close #8
    
        '  FileCopy mfile, "C:\" + F_TXTFILES + "\apantupdate.XML"
    
        Dim MARK2 As String
    
        '  READ_CANCEL v, MARK2

        '  400001829337674

        '    CANCEL_INVOICE = MARK2
    
        '<EhFooter>
        Exit Function

cmdDOMDocumentUTF8_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.cmdDOMDocumentUTF8_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function




Function fupdateclientFIM(Message As Integer, ByVal clientID As String, ByVal ServiceType As String, ByVal completionDateTime As String, POSO As String, servC As String) As String
',ByVal completionDateTime As String, POSO As String, InvoiceKind As String, providedServiceCategory As String) As String '"202502221033

   
        '  Dim docStock As MSXML2.DOMDocument
        Dim elemRoot  As MSXML2.IXMLDOMElement

        Dim elemStock As MSXML2.IXMLDOMElement

        Dim elemField As MSXML2.IXMLDOMElement
        Dim elem1Field As MSXML2.IXMLDOMElement
        Dim elem2Field As MSXML2.IXMLDOMElement
        
        
        
        
        

        Dim i         As Integer
    
100     Set docStock = New MSXML2.DOMDocument

102     With docStock
104         .appendChild .createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
106         Set elemRoot = .createElement("dcrudt:UpdateClientDoc")
108         elemRoot.setAttribute "xmlns:dcrudt", "https://www.aade.gr/myDATA/dcrudt/v1.0"
110         elemRoot.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
112         elemRoot.setAttribute "xsi:schemaLocation", "https://www.aade.gr/myDATA/dcrudt/v1.0 updateClientType-v1.0.xsd updateClient-v1.0.xsd"

114         Set elemStock = docStock.createElement("dcrudt:updateClient")
            
           ' Dim nposo As String: nposo = Replace(Format(POSO, "####0.00"), ",", ".")
            
            
116          Set elem2Field = docStock.createElement("dcrudt:initialDclId"): elem2Field.Text = clientID: elemStock.appendChild elem2Field
118         Set elem2Field = docStock.createElement("dcrudt:clientServiceType"): elem2Field.Text = ServiceType: elemStock.appendChild elem2Field

               If Val(servC) = 4 Or (F_ISCAR = 2 And Val(servC) = 9) Then 'dorean
                       
               End If


120         Set elem2Field = docStock.createElement("dcrudt:entryCompletion"): elem2Field.Text = "true": elemStock.appendChild elem2Field
             'Set elem2Field = docStock.createElement("dcrudt:nonIssueInvoice"): elem2Field.Text = "false": elemStock.appendChild elem2Field
'<<<<<<< .mine
             
             If Val(servC) = 4 Or (F_ISCAR = 2 And Val(servC) = 9) Then 'dorean
                 Set elem2Field = docStock.createElement("dcrudt:nonIssueInvoice"): elem2Field.Text = "true": elemStock.appendChild elem2Field
                 
                 
                 
                 
             Else
             
             
             
             
             
                 Set elem2Field = docStock.createElement("dcrudt:nonIssueInvoice"): elem2Field.Text = "false": elemStock.appendChild elem2Field
             
                 Set elem2Field = docStock.createElement("dcrudt:amount"): elem2Field.Text = POSO: elemStock.appendChild elem2Field
                 
                 
             End If
                 
'=======
            
'            If Val(InvoiceKind) > 0 Then
'
'                    Set elem2Field = docStock.createElement("dcrudt:amount"): elem2Field.Text = POSO: elemStock.appendChild elem2Field


'            End If
'


          If Val(servC) = 4 Or (F_ISCAR = 2 And Val(servC) = 9) Then 'dorean
          
          Else

             Set elem2Field = docStock.createElement("dcrudt:isDiffVehReturnLocation"): elem2Field.Text = "false": elemStock.appendChild elem2Field
             
          End If
             
            ' Set elemField = docStock.createElement("dcrudt:providedServiceCategory"): elemField.Text = "4"
134        ' Set elem2Field = docStock.createElement("dcrudt:ProvidedCategoryServiceType"): elem2Field.Text = 3: elemField.appendChild elem2Field
'elemStock.appendChild elemField
'<<<<<<< .mine
 Set elemField = docStock.createElement("dcrudt:providedServiceCategory"): elemField.Text = servC: elemStock.appendChild elemField
 
 
  If Val(servC) = 4 Or (F_ISCAR = 2 And Val(servC) = 9) Then 'dorean  'If Val(servC) = 4 Then 'dorean
      Set elem2Field = docStock.createElement("dcrudt:reasonNonIssueType"): elem2Field.Text = 1: elemStock.appendChild elem2Field
 Else
    Set elemField = docStock.createElement("dcrudt:invoiceKind"): elemField.Text = 3: elemStock.appendChild elemField
 End If
 
'=======
' Set elemField = docStock.createElement("dcrudt:providedServiceCategory"): elemField.Text = providedServiceCategory: elemStock.appendChild elemField
 
 
 'If Val(InvoiceKind) > 0 Then
 '       Set elemField = docStock.createElement("dcrudt:invoiceKind"): elemField.Text = InvoiceKind: elemStock.appendChild elemField
' End If



'>>>>>>> .theirs
 '<dcrudt:providedServiceCategory>2</dcrudt:providedServiceCategory>
  '  <dcrudt:invoiceKind>2</dcrudt:invoiceKind>
 
 
 
 
 Set elem2Field = docStock.createElement("dcrudt:comments"): elem2Field.Text = sxolia2.Text: elemStock.appendChild elem2Field
22          elemRoot.appendChild elemStock
  
 
124         Set .documentElement = elemRoot

            On Error Resume Next
            
126         FormatXmlDocument docStock ' ‚·ÊÂÈ ÍÂÌ· Ì· ÂÈÌ·È ÂıÍÔÎÔ ÛÙÔ ‰È·‚·ÛÏ·

            Dim ANS2 As Integer
If Message = 1 Then
128         ANS2 = MsgBox(docStock.XML, vbYesNo)

         If ANS2 = vbNo Then

                Exit Function

            End If
            
End If
            
         Kill "C:\txtfiles\" + clientID + ".xml"

            On Error GoTo cmdDOMDocumentUTF8_Click_Err

         .save "C:\txtfiles\" + clientID + ".xml"
        End With

     FileCopy "C:\txtfiles\" + clientID + ".xml", "c:\" + F_TXTFILES + "\sendinv\update-aytok" + "-" + clientID + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"
    
        Dim MARK As String

        '    MARK = InputBox("‰˘ÛÂ Ùo mark", "Ï·ÒÍ")
     
     url2 = url2 + MARKTIM
     
     url2 = UrlAADE + "DCL/UpdateClient"

        'initialize
     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
    XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
148     XMLServer.Option(9) = 2048
150     XMLServer.Option(6) = True
    
152     XMLServer.Open "POST", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
154     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
156     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"

158     XMLServer.send docStock.XML 'TXML 'strData  '”‘≈ÀÕŸ ‘œ STRING TXML STO MYDATA  'XMLServer.send TXML '  ¡Õ »≈ÀŸ Õ¡ ‘œ ”‘≈…ÀŸ ¡–œ ¡—◊≈…œ ‘œ‘≈ ¬¡∆Ÿ strData
160     Debug.Print XMLServer.Status

162     v = XMLServer.responseText
        '--------------- ¡–œ”‘œÀ« INVOICES -----------------------------------------------------
 If Message = 1 Then
164     MsgBox v
End If
        Dim CC As String

166     CC = takefromxml(v)
168     fupdateclientFIM = CC

170     TIMOL = Replace(TIMOL, "/", "-")

        Dim mfile As String:

172     mfile = "c:\" + F_TXTFILES + "\sendinv\apant-update-aytok" + "-" + clientID + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"
  
        ' Dim mfileAsk As String
        ' mfileAsk = "C:\txtfiles\" + clientID + ".xml"
174     Debug.Print v
176     Open mfile For Output As #8
178     Print #8, v
180     Close #8
        Dim MARK2 As String
        Exit Function
        
        
'---------------------------------------ƒŸ—≈¡Õ -----------------------------------------
'        <dcrudt:updateClient>
'    <dcrudt:initialDclId>100000000038724</dcrudt:initialDclId>
'    <dcrudt:clientServiceType>3</dcrudt:clientServiceType>
'    <dcrudt:entryCompletion>true</dcrudt:entryCompletion>
'    <dcrudt:nonIssueInvoice>false</dcrudt:nonIssueInvoice>
'    <dcrudt:amount>10</dcrudt:amount>
'    <dcrudt:isDiffVehReturnLocation>false</dcrudt:isDiffVehReturnLocation>
'    <dcrudt:providedServiceCategory>4</dcrudt:providedServiceCategory>
'    <dcrudt:comments>----</dcrudt:comments>
'  </dcrudt:updateClient>
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        

cmdDOMDocumentUTF8_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.cmdDOMDocumentUTF8_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next
End Function

Function fupdateclientTIM(Message As Integer, ByVal clientID As String, _
                              ByVal ServiceType As String, _
                              ByVal completionDateTime As String, providedServiceCategory As Integer, InvoiceKind As Integer, ByVal olokl As String, AFM As String, POSO As String) As String '"202502221033
   
        '  Dim docStock As MSXML2.DOMDocument
        Dim elemRoot  As MSXML2.IXMLDOMElement

        Dim elemStock As MSXML2.IXMLDOMElement

        Dim elemField As MSXML2.IXMLDOMElement
        Dim elem1Field As MSXML2.IXMLDOMElement
        Dim elem2Field As MSXML2.IXMLDOMElement
        
        
        
        
        

        Dim i         As Integer
    
100     Set docStock = New MSXML2.DOMDocument

         POSO = Replace(POSO, ",", ".")
         If gVal(POSO) = 0 Then POSO = "0"



102     With docStock
104         .appendChild .createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
106         Set elemRoot = .createElement("dcrudt:UpdateClientDoc")
108         elemRoot.setAttribute "xmlns:dcrudt", "https://www.aade.gr/myDATA/dcrudt/v1.0"
110         elemRoot.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
112         elemRoot.setAttribute "xsi:schemaLocation", "https://www.aade.gr/myDATA/dcrudt/v1.0 updateClientType-v1.0.xsd updateClient-v1.0.xsd"

114         Set elemStock = docStock.createElement("dcrudt:updateClient")
            
            
116          Set elem2Field = docStock.createElement("dcrudt:initialDclId"): elem2Field.Text = clientID: elemStock.appendChild elem2Field
118         Set elem2Field = docStock.createElement("dcrudt:clientServiceType"): elem2Field.Text = ServiceType: elemStock.appendChild elem2Field
120         Set elem2Field = docStock.createElement("dcrudt:entryCompletion"): elem2Field.Text = olokl: elemStock.appendChild elem2Field
             Set elem2Field = docStock.createElement("dcrudt:nonIssueInvoice"): elem2Field.Text = "false": elemStock.appendChild elem2Field
            Set elem2Field = docStock.createElement("dcrudt:amount"): elem2Field.Text = POSO: elemStock.appendChild elem2Field
             
             
             Set elemField = docStock.createElement("dcrudt:providedServiceCategory"): elemField.Text = providedServiceCategory: elemStock.appendChild elemField
            
           ' invoiceKind
            
            Set elem2Field = docStock.createElement("dcrudt:invoiceKind"): elem2Field.Text = InvoiceKind: elemStock.appendChild elem2Field
            
            
          If Len(AFM) > 0 Then
            
              Set elem2Field = docStock.createElement("dcrudt:invoiceCounterparty"): elem2Field.Text = AFM: elemStock.appendChild elem2Field
         End If



           Set elem2Field = docStock.createElement("dcrudt:invoiceCounterpartyCountry"): elem2Field.Text = "GR": elemStock.appendChild elem2Field
             
            
            
            
            
          '
            
          ' ƒ≈Õ ‘œ ƒ≈◊≈‘¡… Set elem2Field = docStock.createElement("dcrudt:comments"): elem2Field.Text = sxolia2.Text: elemStock.appendChild elem2Field
    
            
122                        elemRoot.appendChild elemStock
  
            '  End With  /////////////////////////////////////////////////////////////////////////////////////////////
124         Set .documentElement = elemRoot
   
            




            On Error Resume Next
            
126         FormatXmlDocument docStock ' ‚·ÊÂÈ ÍÂÌ· Ì· ÂÈÌ·È ÂıÍÔÎÔ ÛÙÔ ‰È·‚·ÛÏ·

            Dim ANS2 As Integer
            If Message = 1 Then

128         ANS2 = MsgBox(docStock.XML, vbYesNo)

         If ANS2 = vbNo Then

                Exit Function

            End If
            
            End If
            
         Kill "C:\txtfiles\" + clientID + ".xml"

            On Error GoTo cmdDOMDocumentUTF8_Click_Err

         .save "C:\txtfiles\" + clientID + ".xml"
        End With

     FileCopy "C:\txtfiles\" + clientID + ".xml", "c:\" + F_TXTFILES + "\sendinv\update-aytok" + "-" + clientID + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"

        ' TO XML EINAI docStock.xml

        '--------------- ¡–œ”‘œÀ« INVOICES -----------------------------------------------------
     
        Dim MARK As String

        '    MARK = InputBox("‰˘ÛÂ Ùo mark", "Ï·ÒÍ")
     
     url2 = url2 + MARKTIM
     
     url2 = UrlAADE + "DCL/UpdateClient"

        'initialize
     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
    XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
148     XMLServer.Option(9) = 2048
150     XMLServer.Option(6) = True
    
152     XMLServer.Open "POST", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
154     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
156     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"

158     XMLServer.send docStock.XML 'TXML 'strData  '”‘≈ÀÕŸ ‘œ STRING TXML STO MYDATA  'XMLServer.send TXML '  ¡Õ »≈ÀŸ Õ¡ ‘œ ”‘≈…ÀŸ ¡–œ ¡—◊≈…œ ‘œ‘≈ ¬¡∆Ÿ strData
160     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
162     v = XMLServer.responseText
        '--------------- ¡–œ”‘œÀ« INVOICES -----------------------------------------------------
     If Message = 1 Then
164     MsgBox v
     End If
        Dim CC As String

166     CC = takefromxml(v)
168     fupdateclientTIM = CC

170     TIMOL = Replace(TIMOL, "/", "-")

        Dim mfile As String:

172     mfile = "c:\" + F_TXTFILES + "\sendinv\apant-update-aytok" + "-" + clientID + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"
  
        ' Dim mfileAsk As String
        ' mfileAsk = "C:\txtfiles\" + clientID + ".xml"
174     Debug.Print v
176     Open mfile For Output As #8
178     Print #8, v
180     Close #8
    
        '  FileCopy mfile, "C:\" + F_TXTFILES + "\apantupdate.XML"
    
        Dim MARK2 As String
    
        '  READ_CANCEL v, MARK2

        '  400001829337674

        '    CANCEL_INVOICE = MARK2
    
        '<EhFooter>
        Exit Function

cmdDOMDocumentUTF8_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.cmdDOMDocumentUTF8_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function

Public Sub loadpar72()

        '  gdb.Open "DSN=MERCSQL;DATABASE=EMPMYDATA"
        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     F_TXTFILES = Trim(FINDPARAMETROI(1, "PAR7MYDATA", "F_TXTFILES", "TXTFILES", "÷¡ ≈Àœ” ¡–œ”‘ŸÀŸÕ "))

102     F_E3_APOKATHG_EID = Val(FINDPARAMETROI(1, "PAR7MYDATA", "F_E3_APOKATHG_EID", "0", "E3 ·¸ Í·ÙÁ„ÔÒﬂ· ÂÈ‰˛Ì=1 ¡Ô –·Ò·ÛÙ·ÙÈÍ‹=0 "))

104     F_dokimi = Val(FINDPARAMETROI(1, "PAR7MYDATA", "F_dokimi", "1", "1=‰ÔÍÈÏÁ 1ÔÚ user 2=‰ÔÍÏÁ 2ÔÚ user"))

106     f_YPOK = FINDPARAMETROI(1, "PAR7MYDATA", "f_YPOK", "0", "¡—…»Ãœ” ’–œ ¡‘¡”‘«Ã¡‘œ” 0= ≈Õ‘—… œ ")
108     f_YPOK = Left$(f_YPOK, 1)

110     If InStr("0123", f_YPOK) = 0 Then
112         f_YPOK = "0"
        End If

        'Dim f_par1ypok As String ' ·Ò·ÛÙ·ÙÈÍ‹ Ôı ·ÌﬁÍÔıÌ ÛÙÔ ıÔÍ/Ï· 1
114     f_par1ypok = FINDPARAMETROI(1, "PAR7MYDATA", "f_par1ypok", "", "·Ò·ÛÙ·ÙÈÍ‹ Ôı ·ÌﬁÍÔıÌ ÛÙÔ ıÔÍ/Ï· 1 .˜.  ·‚  ")

116     f_par2ypok = FINDPARAMETROI(1, "PAR7MYDATA", "f_par2ypok", "", "·Ò·ÛÙ·ÙÈÍ‹ Ôı ·ÌﬁÍÔıÌ ÛÙÔ ıÔÍ/Ï· 2 .˜.  gW  ")

118     F_PAROCHOS = Val(FINDPARAMETROI(1, "PAR7MYDATA", "f_parochos", "0", "0=MyData 1=–·ÒÔ˜ÔÚ"))

120     If F_PAROCHOS = 1 Then
122        ' MYDPAR.Visible = True
124        ' MYDPAR.Text = MYDPAR.List(1)
        End If

126     f_apoDOS = Val(FINDPARAMETROI(1, "PAR7MYDATA", "f_apoDOS", "0", "Ã–œ—Ÿ Õ¡ ‘—¡¬«ŒŸ ¡–œ DOS C:\LAGEURO=1 OXI=0 "))

128     If f_apoDOS = 1 Then
130         'APODOS.Visible = True
   
        End If

        Dim R As New ADODB.Recordset

132     R.Open "select * from MEM", Gdb, adOpenDynamic, adLockOptimistic

134     f_pelono = R("pelono")
136     f_peldie = R("peldie")
138     f_pelarit = R("PELARIT")
140     f_pelpol = R("PELPOL")
142     f_peltk = R("PELTK")

144     F_demo = 0

146     If R!mydataIsDemo = 1 Then
148         PAROX_URLADD = "beta-"

150         If F_dokimi = 1 Then
152             afmCompany = "028783755"
154             P_USER = "glagakis22"   ' "555bc57c80634243958f62b629316aaa"
156             P_KEY = "534044b048f4023151f37c2a44282200"  ' "555bc57c80634243958f62b629316aaa" ' CNull(R!MYDATAPASW)
                 
158             PAROX_USER = "glagakis@gmail.com"
160             PAROX_PASS = "y39213921!!Y"
162             PAROX_KEY = "32657FD7929B483AABBD1C633401E945"
           ElseIf F_dokimi = 2025 Then
             afmCompany = "028783755"
             P_USER = "glagakis@gmail.com"   ' "555bc57c80634243958f62b629316aaa"
             P_KEY = "c155dc156c6322ce6f2156e4c1228e0c"   ' "555bc57c80634243958f62b629316aaa" ' CNull(R!MYDATAPASW)
          
            ElseIf F_dokimi = 1000 Then  'TSIATS
             afmCompany = "302009116"
             P_USER = "TSIATSIARIMARIA"   '
             P_KEY = "c863dd9a3f458795f290dad92af9e899"   ' "555bc57c80634243958f62b629316aaa" ' CNull(R!MYDATAPASW)
           
           
            Else
164             afmCompany = "150144455"
166             P_USER = "plagakis"
168             P_KEY = "ed99b37f3e99cc46ba284186185b5cdc"
       
170             PAROX_USER = "glagakis@gmail.com"
172             PAROX_PASS = "y39213921!!Y"
174             PAROX_KEY = "32657FD7929B483AABBD1C633401E945"
       
            End If
       
176         F_demo = 1
178         UrlAADE = "https://mydataapidev.aade.gr/" '"https://mydata-dev.azure-api.net/"
        
            '  XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
            '  XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"
       
180         MsgBox "ÒÔÛÔ˜ﬁ ‰ÔÍÈÏ·ÛÙÈÍÔ ÂÒÈ‚·ÎÎÔÌ " + IIf(F_PAROCHOS = 1, "·ÒÔ˜Ôı", "myData")
     
        Else
182         UrlAADE = "https://mydatapi.aade.gr/"  'SendInvoices
184         PAROX_URLADD = ""
            ' PAROX_URLADD = "beta-" ' ÒÔÛÔ˜Á debug only Ì· Û‚ÁÛÙÂÈ ÛÙÁÌ ·Ò·„˘„ﬁ

186         P_USER = CNull(R!MYDATAUSER)
188         P_KEY = CNull(R!MYDATAPASW)

190         afmCompany = CNull(R!AFMPEL)

192         If IsNull(R!AFMPEL) Then
194             MsgBox "‰ÂÌ Â˜ÂÈ ‰ÁÎ˘ËÂÈ  ÙÔ ¡÷Ã (¬ÔÁËÁÙÈÍ· ·Ò‹ÏÂÙÒÔÈ ÒÔ„Ò·ÏÏ·ÙÔÚ)"
196             Unload Me

                Exit Sub

            End If
   
198         If IsNull(R!MYDATAUSER) Then
200             MsgBox "‰ÂÌ Â˜ÂÈ ‰ÁÎ˘ËÂÈ  ÙÔ USER NAME √…¡ ‘œ MYDATA.(¬ÔÁËÁÙÈÍ· ·Ò‹ÏÂÙÒÔÈ ÒÔ„Ò·ÏÏ·ÙÔÚ)"
202             Unload Me

                Exit Sub

            End If

204         If IsNull(R!MYDATAPASW) Then
206             MsgBox "‰ÂÌ Â˜ÂÈ ‰ÁÎ˘ËÂÈ Ô  Ÿƒ… œ” √…¡ ‘œ MYDATA.(¬ÔÁËÁÙÈÍ· ·Ò‹ÏÂÙÒÔÈ ÒÔ„Ò·ÏÏ·ÙÔÚ)"
208             Unload Me

                Exit Sub

            End If


210         PAROX_USER = CNull(R.Fields("PAROXUSER").Value) '"glagakis@gmail.com"
212         PAROX_PASS = CNull(R.Fields("PAROXPASS").Value) '"y39213921!!Y"
214         PAROX_KEY = CNull(R.Fields("PAROXKEY").Value) '"32657FD7929B483AABBD1C633401E945"

        End If
 
216     R.Close
 
218     R.Open "select ISNULL(AYJON,0) AS AYJON,ISNULL(TIMH,0) AS TIMH,ISNULL(N1,0) AS N1  FROM PINAKES WHERE TYPOS=1", Gdb, adOpenDynamic, adLockOptimistic

        Dim lp As Integer

220     Do While Not R.EOF
222         antFPA(R!AYJON) = R!n1
224         timhFPA(R!AYJON) = R!TIMH
    
226         If R!n1 = 0 Or R!n1 > 8 Then
228             MsgBox "À¡»œ” ”‘…” ¡Õ‘…”‘œ…◊«”≈…” ÷–¡"
230             apo.Enabled = False
232             eos.Enabled = False
234             Command2.Enabled = False
236             monoApestalmena.Enabled = False
            End If

238         R.MoveNext
        Loop
 
240     R.Close
  updateScreen
  
  
'242     APO.Value = Now
'244     EOS.Value = Now
'
'        Dim sql As String
'
'246     sql = FIND_QUERY()
'
'248     If sql = "0" Then
'
'250         APO.Enabled = False
'252         EOS.Enabled = False
'254         Command2.Enabled = False
'256         monoApestalmena.Enabled = False
'
'258         GoTo lab1000
'        End If
'
'260     If Len(where.Caption) > 0 Then  ' ONLINE ƒ≈Õ ≈◊—≈…¡∆≈‘¡… GRID
'
'        Else
'
'262         If gisOnlineMydata = 0 Then
'264             Adodc1.ConnectionString = gConnect  '1
'266             Adodc1.RecordSource = sql
'268             Adodc1.Refresh
'            End If
'
'        End If
'
'        Dim R56 As New ADODB.Recordset
'
'270     R56.Open "select MYDATAFROMEID1 FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
'272     fMydataFromEID1 = nNull(R56(0))
'274     R56.Close
'
'        Dim n50 As Integer
'
'276     If fMydataFromEID1 = 1 Then '  '≈Õ«Ã≈—Ÿ”« ¡–œ  ¡‘«√œ—…¡ ≈…ƒŸÕ
'278         R.Open "SELECT AYJON,C1,LOG1,ISNULL(C3,'') AS C3 FROM PINAKES WHERE TYPOS=15", Gdb, adOpenDynamic, adLockOptimistic
'280         n50 = 0
'
'282         Do While Not R.EOF
'284             n50 = n50 + 1
'286             fKatEsod(n50) = Trim(CNull(R!c1))
'288             FkatE3Esod(n50) = Trim(CNull(R!c3)) 'E3-561-1 ¡Õ ‘—¡¬¡≈… ‘œ ≈3 ¡–œ ‘«Õ  ¡‘«√œ—…¡ ≈…ƒœ’”  ¡… œ◊… ¡–œ ‘œ –¡—¡”‘¡‘… œ
'290             fKatEXod(n50) = Trim(CNull(R!log1))
'292             R.MoveNext
'            Loop
'
'294         R.Close
'
'296         If n50 = 1 Then
'298             MsgBox "‰ÂÌ Â˜ÔıÌ ‰ÁÎ˘ËÂÈ ÔÈ Í·ÙÁ„ÔÒﬂÂÚ ÂÈ‰˛Ì"
'
'                Exit Sub
'
'            End If
'
'        End If
'
        '174  sql = " CREATE TABLE [dbo].[APESTALMENA](" _
        '       & "[AFM] [nchar](10) NULL," _
        '       & "[ISSUEDATE] [nchar](10) NULL," _
        '       & "[AA] [nchar](10) NULL," _
        '       & "[TYPOS] [nchar](10) NULL," _
        '       & "[PAYTYPE] [nchar](10) NULL," _
        '       & "[PAYAJI] [float] NULL," _
        '       & "[TOTALNETVALUE] [float] NULL," _
        '       & "[TOTALVATAMOUNT] [float] NULL," _
        '       & "[TOTALGROSSVALUE] [float] NULL, " _
        '       & "[CLASSIFICATIONTYPE] [nchar](10) NULL," _
        '       & "[MARK] [nchar](20) NULL," _
        '       & "[ID] [int] IDENTITY(1,1) NOT NULL" _
        '    & ") ON [PRIMARY]"
        '
        '
        '
        '176       R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'APESTALMENA'", Gdb, adOpenDynamic, adLockOptimistic
        '        'On Error Resume Next
        '
        '178     If R(0) = 0 Then
        '180         Gdb.Execute sql
        '        End If
        '
        '
        '182   dum2 = ADD_FIELD("APESTALMENA", "CLASSTYPE11", "[nchar](10) NULL")
        '184   dum2 = ADD_FIELD("APESTALMENA", "CLASSTYPE12", "[nchar](10) NULL")
        '186   dum2 = ADD_FIELD("APESTALMENA", "CLASSCAT11", "[nchar](10) NULL")
        '188   dum2 = ADD_FIELD("APESTALMENA", "CLASSCAT12", "[nchar](10) NULL")
        '
        '190   dum2 = ADD_FIELD("APESTALMENA", "CLASSTYPE21", "[nchar](10) NULL")
        '192   dum2 = ADD_FIELD("APESTALMENA", "CLASSTYPE22", "[nchar](10) NULL")
        '194   dum2 = ADD_FIELD("APESTALMENA", "CLASSCAT21", "[nchar](10) NULL")
        '196   dum2 = ADD_FIELD("APESTALMENA", "CLASSCAT22", "[nchar](10) NULL")
        '
        '198   dum2 = ADD_FIELD("APESTALMENA", "MARK2", "[nchar](20) NULL")
        '
        '200    dum2 = ADD_FIELD("APESTALMENA", "AMOUNT1", "float NULL")
        '202    dum2 = ADD_FIELD("APESTALMENA", "AMOUNT2", "float NULL")
        '
        '204    dum2 = ADD_FIELD("APESTALMENA", "ISANTIST", "INT NULL")
        '206    dum2 = ADD_FIELD("APESTALMENA", "ATIM", "NVARCHAR(12) NULL ")
        '
        '207    dum2 = ADD_FIELD("APESTALMENA", "DEMO", "INT NULL ")
        '

        'paris user           lagakis2
        'paris key            c10c55456c3a411b866cd0f5e69155dc

        'XMLServer.setRequestHeader "aade-user-id", "glagakis2"
        'XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", "555bc57c80634243958f62b629316aaa"

        '<EhFooter>
lab1000:
       Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.Form_Load " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Public Function newclient2(Message As Integer, ByVal garaz3epark2 As Integer, _
                          ark As String, _
                          vFACTORY As String, _
                          vCategory As String) As String

100     If Val(garaz3epark2) = 2 Then
102         TYPOS = "dcrnew:parkingcarwash"
        Else
104         TYPOS = "dcrnew:garage"
        End If

'DEBUG !!!!!!!
''  TYPOS = "dcrnew:parkingcarwash"

       ' Dim varStock  As Variant

        '  Dim docStock As MSXML2.DOMDocument
        Dim elemRoot  As MSXML2.IXMLDOMElement

        Dim elemStock As MSXML2.IXMLDOMElement

        Dim elemField As MSXML2.IXMLDOMElement

        Dim i         As Integer
    
106    ' varStock = Array(Array("≈’—Ÿ12345", 10.32), Array("•45632", 5.43)) 'Yen sign used here to show Unicode.  '
    
108     Set docStock = New MSXML2.DOMDocument

110     With docStock
112         .appendChild .createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
114         Set elemRoot = .createElement("dcrnew:NewDigitalClientDoc")

116         elemRoot.setAttribute "xmlns:dcrnew", "http://www.aade.gr/myDATA/dcrnew/v1.0"
118         elemRoot.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
120         elemRoot.setAttribute "xsi:schemaLocation", "http://www.aade.gr/myDATA/dcrnew/v1.0 SendClient-v1.0.xsd"

122         Set elemStock = docStock.createElement("dcrnew:newDigitalClient")
            ' With elemStock----------------------------------2002-05-30T09:00:00
                   
124         Set elem2Field = docStock.createElement("dcrnew:clientServiceType"): elem2Field.Text = garaz3epark2: elemStock.appendChild elem2Field
                      
            'HTAN OK ALLA EBGALE MHNYMA (NEXT LINE)
            'Set elem2Field = docStock.createElement("dcrnew:creationDateTime"): elem2Field.Text = Format(Now, "yyyy-MM-ddTHH:mm:ss"): elemStock.appendChild elem2Field
            '<message>creationDateTime is allowed only when transmissionFailure has value</message>
                  
126         Set elem2Field = docStock.createElement("dcrnew:entityVatNumber"): elem2Field.Text = afmCompany: elemStock.appendChild elem2Field
128         Set elem2Field = docStock.createElement("dcrnew:branch"): elem2Field.Text = "0": elemStock.appendChild elem2Field
               
            Set elem2Field = docStock.createElement("dcrnew:comments"): elem2Field.Text = txSXOLIA.Text: elemStock.appendChild elem2Field
               '<dcrnew:comments>TEST_GARAGE_CASE</dcrnew:comments>

           '------------------------------------------------------------------------------
130         Set elemField = docStock.createElement("dcrnew:useCase")


            '========================================================
132         Set elem1Field = docStock.createElement(TYPOS)
                              
134         Set elem2Field = docStock.createElement("dcrnew:vehicleRegistrationNumber"): elem2Field.Text = ark: elem1Field.appendChild elem2Field
136         Set elem2Field = docStock.createElement("dcrnew:vehicleCategory"): elem2Field.Text = vCategory: elem1Field.appendChild elem2Field
138         Set elem2Field = docStock.createElement("dcrnew:vehicleFactory"): elem2Field.Text = vFACTORY: elem1Field.appendChild elem2Field
                              
140         elemField.appendChild elem1Field
           '==========================================================
                       
142         elemStock.appendChild elemField

            '--------------------------------------------------------------------------------
                    
144         elemRoot.appendChild elemStock
  
            '  End With  /////////////////////////////////////////////////////////////////////////////////////////////
146         Set .documentElement = elemRoot

            On Error Resume Next
            
148         FormatXmlDocument docStock ' ‚·ÊÂÈ ÍÂÌ· Ì· ÂÈÌ·È ÂıÍÔÎÔ ÛÙÔ ‰È·‚·ÛÏ·

If Message = 1 Then
150         MsgBox docStock.XML
End If
 
152         Kill "C:\txtfiles\" + ARKYKL + ".xml"

            On Error GoTo cmdDOMDocumentUTF8_Click_Err

154         .save "C:\txtfiles\" + ARKYKL + ".xml"
        End With
  
156     url2 = UrlAADE + "CancelInvoice?mark=" '/RequestDocs"
     
        Dim MARK As String

        '    MARK = InputBox("‰˘ÛÂ Ùo mark", "Ï·ÒÍ")
     
158     url2 = url2 + MARKTIM
     
      If F_demo = 1 Then
160      url2 = "https://mydataapidev.aade.gr/DCL/SendClient"
      Else
      
      url2 = "https://mydatapi.aade.gr/DCL/SendClient"

        ' url2 = UrlAADE + "DCL/SendClient"
        'initialize
      End If
        
        
        
        
        
        
        
162     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
164     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
166     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
168     XMLServer.Option(9) = 2048
170     XMLServer.Option(6) = True
    
172     XMLServer.Open "POST", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
174     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
176     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"
178     XMLServer.send docStock.XML 'TXML 'strData  '”‘≈ÀÕŸ ‘œ STRING TXML STO MYDATA  'XMLServer.send TXML '  ¡Õ »≈ÀŸ Õ¡ ‘œ ”‘≈…ÀŸ ¡–œ ¡—◊≈…œ ‘œ‘≈ ¬¡∆Ÿ strData
180     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
  
182     v = XMLServer.responseText

        '--------------- ¡–œ”‘œÀ« INVOICES -----------------------------------------------------
        Dim CC As String

184     CC = takefromxml(v)
186     newclient2 = CC


If Message = 1 Then
188     MsgBox v
End If

190     TIMOL = Replace(TIMOL, "/", "-")

        Dim mfile As String:

192     mfile = "c:\" + F_TXTFILES + "\sendinv\aytok" + "-" + ARKYKL + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"
194     Debug.Print v
196     Open mfile For Output As #8
198     Print #8, v
200     Close #8
    
202     FileCopy mfile, "C:\" + F_TXTFILES + "\apantSendCANC.XML"
    
        Dim MARK2 As String
    
        '  READ_CANCEL v, MARK2

        '  400001829337674

        '    CANCEL_INVOICE = MARK2
    
        '<EhFooter>
        Exit Function

cmdDOMDocumentUTF8_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.cmdDOMDocumentUTF8_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

        '–Â‰ﬂÔ '‘˝ÔÚ'’Ô˜ÒÂ˘ÙÈÍ¸'–ÂÒÈ„Ò·ˆﬁ'¡Ô‰ÂÍÙ›Ú ÙÈÏ›Ú
        'idDcl'xs:long'º˜È'ÃÔÌ·‰ÈÍ¸Ú ¡ÒÈËÏ¸Ú ÿÁˆÈ·ÍÔ˝ –ÂÎ·ÙÔÎÔ„ﬂÔı'”ıÏÎÁÒ˛ÌÂÙ·È ·¸ ÙÁÌ ’ÁÒÂÛﬂ·
        'ClientServiceType'ClientServiceType'Õ·È'‘˝ÔÚ –ÂÎ·ÙÔÎÔ„ﬂÔı'1 = ≈ÌÔÈÍﬂ·ÛÁ'2 = –‹ÒÍÈÌ„Í/'Digital Client REST API 13'–ÎıÌÙﬁÒÈ·'3 = ”ıÌÂÒ„Âﬂ·
        'creationDateTime'xs: DateTime'º˜È'«ÏÂÒÔÏÁÌﬂ· Í·È ˛Ò· ‰ÁÏÈÔıÒ„ﬂ·Ú Â„„Ò·ˆﬁÚ ÛÙÔ ÿÁˆÈ·Í¸ –ÂÎ·ÙÔÎ¸„ÈÔ
        'entityVatNumber'xs:string'º˜È'¡÷Ã œÌÙ¸ÙÁÙ·Ú ¡Ì·ˆÔÒ‹Ú'entityVatNumber
        'branch'xs:int'º˜È'¡Ò.≈„Í·Ù‹ÛÙ·ÛÁÚ
        'recurringService'xs:boolean'º˜È'≈·Ì·Î·Ï‚·Ì¸ÏÂÌÁ ’ÁÒÂÛﬂ·
        'continuousServic'xs:boolean'º˜È'ƒÈ·ÒÍﬁÚ ’ÁÒÂÛﬂ·
        'fromAgreedPeriodDate'xs: Date'º˜È'¡¸ ”ıÏˆ˘ÌÁÏ›ÌÁ –ÂÒﬂÔ‰Ô
        'toAgreedPeriodDate'xs: Date'º˜È'∏˘Ú ”ıÏˆ˘ÌÁÏ›ÌÁ –ÂÒﬂÔ‰Ô
        'mixedService'xs:boolean'º˜È'ÃÂÈÍÙﬁ ’ÁÒÂÛﬂ·
        'customerVatNumber xs:string º˜È ¡÷Ã –ÂÎ‹ÙÁ
        'customerCountry'CountryType'º˜È'◊˛Ò· –ÂÎ‹ÙÁ
        'transmissionFailure'xs:byte'º˜È'¡‰ıÌ·Ïﬂ· ≈ÈÍÔÈÌ˘Ìﬂ·Ú ﬁ ¡‰ıÌ·Ïﬂ· ‰È·‚ﬂ‚·ÛÁÚ'œÈ ÙÈÏ›Ú {1,2} ÂﬂÌ·È ÂÈÙÒÂÙ›Ú
        'correlatedDclId'xs:long'º˜È'”ıÛ˜ÂÙÈÊ¸ÏÂÌ· DC ID
        'Comments'Comments'º˜È'”˜¸ÎÈ·/ –·Ò·ÙÁÒﬁÛÂÈÚ
        'useCase'ServiceType'Õ·È'–ÂÒÈÙ˛ÛÂÈÚ ˜ÒﬁÛÁÚ

End Function




Public Sub NEWPARKING(ARKYKLText As String, MTPARK2SYN3 As Integer, VFACTText As String, VCATText As String, txSXOLIAText As String, MATIM As String, mPoso As String)
 
 
 Dim nn As Integer
   nn = GGET_NVALUE("SELECT COUNT(*) FROM CARS WHERE VEHICLEREGISTRATIONNUMBER='" + ARKYKLText + "' AND ISNULL(ENTRYCOMPLETION,0)=0")
If nn > 0 Then
   MsgBox "’–¡—◊≈… «ƒ« ‘œ " + ARKYKL.Text + " ”≈ ≈  —≈Ãœ‘«‘¡"
   Exit Sub
End If

' 3=synergeia
   ' Dim MT As Integer
    'If F_ISCAR = 1 Then MT = 3 Else MT = 2

Dim vid As String
loadpar72
vid = newclient2(0, MTPARK2SYN3, ARKYKLText, VFACTText, VCATText)
Me.Caption = vid

 If Len(vid) > 10 Then
     'If F_ISCAR >= 1 Then
     
'   UPDATE_PEL "VEHICLEREGISTRATIONNUMBER", Left(ARKYKL.Text, Adodc1.Recordset("VEHICLEREGISTRATIONNUMBER").DefinedSize), 0, x, MKOD
'   UPDATE_PEL "VEHICLECATEGORY", Left(VCAT.Text, Adodc1.Recordset("VEHICLECATEGORY").DefinedSize), 0, x, MKOD
'   UPDATE_PEL "VEHICLEFACTORY", Left(VFACT.Text, Adodc1.Recordset("VEHICLEFACTORY").DefinedSize), 0, x, MKOD
'   UPDATE_PEL "NEWDIGITALCLIENTID", Left(vid, Adodc1.Recordset("NEWDIGITALCLIENTID").DefinedSize), 0, x, MKOD
   Gdb.Execute "INSERT INTO CARS(COMMENTS,VEHICLEREGISTRATIONNUMBER,VEHICLECATEGORY,VEHICLEFACTORY,NEWDIGITALCLIENTID,HME) VALUES ('" + txSXOLIAText + "','" + ARKYKLText + "','" + VCATText + "','" + VFACTText + "','" + vid + "',GETDATE() )"
   
   
   
    '4 = ƒ˘ÒÂ‹Ì ıÁÒÂÛﬂ· (–‹ÒÍÈÌ„Í/ –ÎıÌÙﬁÒÈ·,”ıÌÂÒ„Âﬂ·)
        '5 = ÀÔÈ‹ (–‹ÒÍÈÌ„Í/ –ÎıÌÙﬁÒÈ·,”ıÌÂÒ„Âﬂ·)
        '6 = ¡ÔÊÁÏﬂ˘ÛÁ –·ÒÔ˜ﬁÚ ≈„„˝ÁÛÁÚ
        '7 = ’ÁÒÂÛﬂ· ¬‹ÛÂÈ ‘ÈÏÔÍ·Ù·Î¸„Ôı (–‹ÒÍÈÌ„Í/ –ÎıÌÙﬁÒÈ·)
        '8 = ’ÁÒÂÛﬂ·  ·Ù¸ÈÌ ”ıÏˆ˘Ìﬂ·Ú (–‹ÒÍÈÌ„Í/ –ÎıÌÙﬁÒÈ·)
        '9 = …‰È¸˜ÒÁÛÁ (–‹ÒÍÈÌ„Í/ –ÎıÌÙﬁÒÈ·,”ıÌÂÒ„Âﬂ·)
   
   
   '2 ----------------- STEP ----------
   Dim VID2 As String
    VID2 = fupdateclient(0, vid, LTrim(str(MTPARK2SYN3)), Format(Now, "yyyy-MM-ddTHH:mm:ss"), 7, 3, "0.0")
    If Len(VID2) > 10 Then
        Gdb.Execute "UPDATE CARS SET AFM='" + AFM + "',PROVIDEDSERVICECATEGORY=7,INVOICEKIND=3 where NEWDIGITALCLIENTID='" + vid + "'"
    End If
    
    
    
    '3 ---------------- STEP ----------
    Dim VID3 As String
     ' Dim AA As String, DD As String
    'AA = "0000" ' ¡—…»Ãœ” ¡–œƒ≈…Œ«”
     
     Dim artam As String: artam = FINDPARAMETROI(1, "PELAT1", "F_ARTAM", ";", "¡—…»Ãœ” SN 1ÁÚTAMEIAKH”;2ÁÚÙ·Ï/÷œ—OÀO√IKOY")
   ' DD = Format(Now, "yyyy-MM-dd")
    
    
 
    
    VID3 = RELATECLIEND(0, "", vid, 2, MATIM, Format(Now, "yyyy-MM-dd"), artam)
    
        ' MPOSO = TXTPOSOText
        mPoso = Replace(mPoso, ",", ".")
        If Val(mPoso) = 0 Then
           mPoso = "0.0"
        End If
    Gdb.Execute "UPDATE CARS SET  POSO=" + mPoso + ",  SXETIKA='" + MATIM + "'+' " + Format(Now, "yyyy-MM-ddTHH:mm:ss") + "'+' " + "" + " //'  WHERE NEWDIGITALCLIENTID='" + vid + "'"
        'ISNULL(SXETIKA,'')
         '  Dim c As String

'''' Gdb.Execute "UPDATE CARS SET   SXETIKA='" + MARK + "'  where ID=" + str(ID)
  MARK = ""
    Gdb.Execute "INSERT INTO CARRELATIONS(FIMISSUEDATE,CORRELATEDDCLIDS,CORRELATEDID,FIMAA) VALUES(CONVERT(CHAR(10),GETDATE(),103),'" + vid + "','" + VID3 + "','" + MATIM + "')"
    

    
   
 End If
 
'  cars.Refresh
'ARKYKL.Text = ""
'txSXOLIA.Text = ""

'cmdKatax.Enabled = False '   ¯Ï‰Í·Ù·˜.ÂÌ·‚ÎÂ‰ = ˆ·ÎÛÂ


' ElseIf iInvoiceKind = 3 Then 'FIM
       




End Sub





























 Function RELATECLIEND(Message As Integer, ByVal MARK As String, ByVal CORID As String, mark1fim2 As Integer, mfimaa As String, mfimdate As String, ByVal artam As String) As String
        'CORID =100000000007161#
        'MARK=4000000000
   
        '  Dim docStock As MSXML2.DOMDocument
        Dim elemRoot  As MSXML2.IXMLDOMElement

        Dim elemStock As MSXML2.IXMLDOMElement

        Dim elemField As MSXML2.IXMLDOMElement

        Dim i         As Integer
    
100     Set docStock = New MSXML2.DOMDocument

102     With docStock
104         .appendChild .createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
106         Set elemRoot = .createElement("dcrudtcor:ClientCorrelationDoc")
108         elemRoot.setAttribute "xmlns:dcrudtcor", "http://www.aade.gr/myDATA/dcrudtcor/v1.0"
110         elemRoot.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
112         elemRoot.setAttribute "xsi:schemaLocation", "http://www.aade.gr/myDATA/dcrudtcor/v1.0 SimpleTypes-v1.0.xsd clientCorrelationType-v1.0.xsd"

114         Set elemStock = docStock.createElement("dcrudtcor:clientCorrelation")

116         Set elem2Field = docStock.createElement("dcrudtcor:entityVatNumber"): elem2Field.Text = afmCompany: elemStock.appendChild elem2Field
                                         
118        ' Set ELEM2FIELD = docStock.createElement("dcrudtcor:mark"): ELEM2FIELD.Text = MARK: elemStock.appendChild ELEM2FIELD
            '' Set elem2Field = docStock.createElement("dcrudt:completionDateTime"): elem2Field.Text = "202502221033: elemStock.appendChild elem2Field"

          If mark1fim2 = 1 Then  ' me mark
              Set elem2Field = docStock.createElement("dcrudtcor:mark"): elem2Field.Text = MARK: elemStock.appendChild elem2Field
          Else
            Set elem2Field = docStock.createElement("dcrudtcor:FIM")
            

                    Set elem22field = docStock.createElement("dcrudtcor:FIMNumber"): elem22field.Text = artam: elem2Field.appendChild elem22field
                    Set elem22field = docStock.createElement("dcrudtcor:FIMAA"): elem22field.Text = mfimaa: elem2Field.appendChild elem22field
                    Set elem22field = docStock.createElement("dcrudtcor:FIMIssueDate"): elem22field.Text = mfimdate: elem2Field.appendChild elem22field
                    Set elem22field = docStock.createElement("dcrudtcor:FIMIssueTime"): elem22field.Text = "00:00:00": elem2Field.appendChild elem22field

              elemStock.appendChild elem2Field
           End If




















120         Set elem2Field = docStock.createElement("dcrudtcor:correlatedDCLids"): elem2Field.Text = CORID: elemStock.appendChild elem2Field
                       

                    
122         elemRoot.appendChild elemStock
  
            '  End With  /////////////////////////////////////////////////////////////////////////////////////////////
124         Set .documentElement = elemRoot

            On Error Resume Next
            
126         FormatXmlDocument docStock ' ‚·ÊÂÈ ÍÂÌ· Ì· ÂÈÌ·È ÂıÍÔÎÔ ÛÙÔ ‰È·‚·ÛÏ·

            Dim ANS2 As Integer

If Message = 1 Then


128         ANS2 = MsgBox(docStock.XML, vbYesNo)

130         If ANS2 = vbNo Then

                Exit Function

            End If
            
End If



132         Kill "C:\txtfiles\" + ARKYKL + "r.xml"

            On Error GoTo cmdDOMDocumentUTF8_Click_Err

134         .save "C:\txtfiles\" + ARKYKL + "r.xml"
        End With

        ' TO XML EINAI docStock.xml

        '--------------- ¡–œ”‘œÀ« INVOICES -----------------------------------------------------
     
        ' Dim MARK As String
        '    MARK = InputBox("‰˘ÛÂ Ùo mark", "Ï·ÒÍ")
     
136     url2 = url2 + MARKTIM
     
138     url2 = UrlAADE + "DCL/ClientCorrelations"
     url2 = UrlAADE + "DCL/ClientCorrelations"
        'initialize
140     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
142     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
144     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
146     XMLServer.Option(9) = 2048
148     XMLServer.Option(6) = True
    
150     XMLServer.Open "POST", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
152     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
154     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"

156     XMLServer.send docStock.XML 'TXML 'strData  '”‘≈ÀÕŸ ‘œ STRING TXML STO MYDATA  'XMLServer.send TXML '  ¡Õ »≈ÀŸ Õ¡ ‘œ ”‘≈…ÀŸ ¡–œ ¡—◊≈…œ ‘œ‘≈ ¬¡∆Ÿ strData
158     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
160     v = XMLServer.responseText
        '--------------- ¡–œ”‘œÀ« INVOICES -----------------------------------------------------
 
        Dim CC As String

162     CC = takefromxml(v)
 
164     RELATECLIEND = CC '  <clientCorrelationID>100000000008565</clientCorrelationID>
166     CC = takefromxml(v)
        'CC = takefromxml(v)
 If Message = 1 Then
168     MsgBox v
 End If
170     TIMOL = Replace(TIMOL, "/", "-")

        Dim mfile As String:

172     mfile = "c:\" + F_TXTFILES + "\sendinv\aytok" + "-" + ARKYKL + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"
174     Debug.Print v
176     Open mfile For Output As #8
178     Print #8, v
180     Close #8
    
182     FileCopy mfile, "C:\" + F_TXTFILES + "\apantSendCANC.XML"
    
        Dim MARK2 As String
    
    
    
    RELATECLIEND = CC
    
    
        '  READ_CANCEL v, MARK2

        '  400001829337674

        '    CANCEL_INVOICE = MARK2
    
        '<EhFooter>
        Exit Function

cmdDOMDocumentUTF8_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.RELATE_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function


Function takefromxml(ByVal v As String) As String

        On Error GoTo TAKEFROMXML_Err

100    ' List1.Clear

        Dim RECS As Integer
        Dim CCV  As String
        Dim C1, c2, c3, c4, CURL, SXOLIA As String
        Dim objxml As New MSXML2.DOMDocument
102     objxml.async = True
104     objxml.loadXML v
        Dim nLathosApost As Long
106     nLathosApost = 0
        Dim nApost As Long, nAll As Long
        
108     nApost = 0
110     nAll = 0

        Dim nodeList As IXMLDOMNodeList

        Dim Node     As IXMLDOMNode
    
112     Set nodeList = objxml.SelectNodes("ResponseDoc/response")

114     For Each Node In nodeList
        
            Dim r2       As New ADODB.Recordset

            Dim xmlNode  As IXMLDOMNode

            Dim xml2Node As IXMLDOMNode

            Dim xml3Node As IXMLDOMNode
        
116         If Node.hasChildNodes Then
                'List1.AddItem node.childNodes.Item(0)
118             C1 = "": c2 = "": c3 = ""

120             For Each xmlNode In Node.childNodes

                    ' List1.AddItem xmlNode.nodeName
            
122                 If xmlNode.nodeName = "index" Then
124                     C1 = xmlNode.Text
                        ' List1.AddItem c1
                    End If
            
126                 If xmlNode.nodeName = "invoiceUid" Then
128                     c4 = xmlNode.Text
                        'List1.AddItem c2
                    End If
            
130                 If xmlNode.nodeName = "newClientDclID" Then
132                     c2 = xmlNode.Text
                        'List1.AddItem c2
                    End If

                    'clientCorrelationID
134                 If xmlNode.nodeName = "clientCorrelationID" Then
136                     c2 = xmlNode.Text
                        'List1.AddItem c2
                    End If

                    'updatedClientDclID
138                 If xmlNode.nodeName = "updatedClientDclID" Then
140                     c2 = xmlNode.Text
                        'List1.AddItem c2
                    End If
                '<cancellationID
                If xmlNode.nodeName = "cancellationID" Then
                        c2 = xmlNode.Text
                        'List1.AddItem c2
                    End If
                
                
                    'cURL
142                 If xmlNode.nodeName = "qrUrl" Then
144                     CURL = xmlNode.Text
                        'List1.AddItem c2
                    End If
                
146                 CCV = str(F_ID_NUMS(Val(C1)))
                
                    ' ·Ì ÂÒ˜ÂÙ·È ·¸ ÙÁÌ Â·Ì·ÔÛÙÔÎﬁ ·¸ ·Ò˜ÂﬂÔ Ù¸ÙÂ Ì· ·ﬂÒÌÂÈ ÙÔ id_num Ôı ÎÁÍÙÒÔÎ¸„ÁÛ·
148                 If Val(file_idnum) > 0 Then
150                     CCV = file_idnum
                    End If
            
152                 If xmlNode.nodeName = "errors" Then
            
154                     If xmlNode.hasChildNodes Then  'CHILD=ERROR
                            'List1.AddItem XMLnode.childNodes.Item(0)
                  
156                         For Each xml2Node In xmlNode.childNodes  'ERROR
                        
158                             If xml2Node.hasChildNodes Then 'CHILD=MESSAGE

160                                 For Each xml3Node In xml2Node.childNodes

162                                     If xml3Node.nodeName = "message" Then
164                                         SXOLIA = xml3Node.Text
                                         
166                                         SXOLIA = Left(xml3Node.Text, 40)
168                                         SXOLIA = Replace(SXOLIA, "'", "`")

                                             On Error GoTo TAKEFROMXML_Err

                                            ' GDBEXE "UPDATETIMWITHMARK", "UPDATE TIM SET ENTITY=0,ENTITYUID='" + SXOLIA + "'  WHERE   ID_NUM=" + CCV   ' ENTLINEN=" + str(fLong) + " and ENTITY=" + C1
170                                         'List1.AddItem "  " + SXOLIA
                                        End If

172                                     If xml3Node.nodeName = "code" Then
174                                         SXOLIA = xml3Node.Text + " ************* À¡»œ” ****************"
176                                        ' List1.AddItem "  " + SXOLIA
                                        End If

                                    Next
                        
                                End If

                            Next

                        End If
               
                    End If
            
178                 nAll = nAll + 1

                    On Error Resume Next

                    On Error GoTo TAKEFROMXML_Err

180             Next xmlNode
      
            End If
        
        Next

        '1111
        'DIAXOR txmlapost

182     If nApost = 1 Then
            '  GDBEXE "UPDATETIMWITHMARK", "UPDATE TIM SET XMLTEXT='" + txmlapost + "' WHERE ID_NUM=" + CCV
            ' Gdb.Execute "UPDATE TIM SET XMLMYDATA='" + txmlapost + "' WHERE ID_NUM=" + CCV

        End If

        'GoTo 1111

184     'List1.AddItem "–ﬁÒ·Ì Ã¡—  " + str(nApost)

186     'List1.AddItem "À‹ËÔÚ , ˜˘ÒﬂÚ  Ã¡—  " + str(nLathosApost)

188     takefromxml = c2

        '<EhFooter>
        Exit Function

TAKEFROMXML_Err:

        If noask = 0 Then
            MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.UpdateTIMwithMARK " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
        End If

        Resume Next

        '</EhFooter>

End Function



Private Sub FormatXmlDocument(ByVal xml_doc As DOMDocument)

        '<EhHeader>
        On Error GoTo FormatXmlDocument_Err

        '</EhHeader>
100     FormatXmlNode xml_doc.documentElement, 0

        '<EhFooter>
        Exit Sub

FormatXmlDocument_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.FormatXmlDocument " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

' Add formatting to this element. Indent it and add a
' carriage return before its children. Then recursively
' format the children with increased indentation.
Private Sub FormatXmlNode(ByVal Node As IXMLDOMNode, ByVal indent As Integer)

        '<EhHeader>
        On Error GoTo FormatXmlNode_Err

        '</EhHeader>
        Dim child     As IXMLDOMNode

        Dim text_only As Boolean

        ' Do nothing if this is a text node.
100     If TypeOf Node Is IXMLDOMText Then Exit Sub

        ' See if this node contains only text.
102     text_only = True

104     If Node.hasChildNodes Then

106         For Each child In Node.childNodes

108             If Not (TypeOf child Is IXMLDOMText) Then
110                 text_only = False

                    Exit For

                End If

112         Next child

        End If

        ' Process child nodes.
114     If Node.hasChildNodes Then

            ' Add a carriage return before the children.
116         If Not text_only Then
118             Node.insertBefore Node.ownerDocument.createTextNode(vbCrLf), Node.firstChild
            End If

            ' Format the children.
120         For Each child In Node.childNodes

122             FormatXmlNode child, indent + 2
124         Next child

        End If

        ' Format this element.
126     If indent > 0 Then
            ' Indent before this element.
128         Node.parentNode.insertBefore Node.ownerDocument.createTextNode(Space$(indent)), Node

            ' Indent after the last child node.
130         If Not text_only Then Node.appendChild Node.ownerDocument.createTextNode(Space$(indent))

            ' Add a carriage return after this node.
132         If Node.nextSibling Is Nothing Then
134             Node.parentNode.appendChild Node.ownerDocument.createTextNode(vbCrLf)
            Else
136             Node.parentNode.insertBefore Node.ownerDocument.createTextNode(vbCrLf), Node.nextSibling
            End If
        End If

        '<EhFooter>
        Exit Sub

FormatXmlNode_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.FormatXmlNode " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub



Private Function killClient(ByVal park2garage3 As Integer, ByVal clientID As String) As String
     Dim ANS    As String, TYPOS As String



104     If park2garage3 = 2 Then
106         TYPOS = "dcrnew:parkingcarwash"
        Else
108         TYPOS = "dcrnew:garage"
        End If
   
     

        Dim i         As Integer
    


        Dim MARK As String

        '    MARK = InputBox("‰˘ÛÂ Ùo mark", "Ï·ÒÍ")
     
146     url2 = url2 + MARKTIM
     
148     url2 = UrlAADE + "DCL/CancelClient?dclID=" + clientID ' "100000000007161"

        'initialize
150     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
152     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
154     XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
156     XMLServer.Option(9) = 2048
158     XMLServer.Option(6) = True
    
160     XMLServer.Open "POST", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
162     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
164     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"

166     XMLServer.send 'docStock.XML 'TXML 'strData  '”‘≈ÀÕŸ ‘œ STRING TXML STO MYDATA  'XMLServer.send TXML '  ¡Õ »≈ÀŸ Õ¡ ‘œ ”‘≈…ÀŸ ¡–œ ¡—◊≈…œ ‘œ‘≈ ¬¡∆Ÿ strData
168     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
170     v = XMLServer.responseText
        '--------------- ¡–œ”‘œÀ« INVOICES -----------------------------------------------------
 
172     MsgBox v

174     TIMOL = Replace(TIMOL, "/", "-")

        Dim mfile As String:

176     mfile = "c:\" + F_TXTFILES + "\sendinv\aytok" + "-" + ARKYKL + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"
178     Debug.Print v
180     Open mfile For Output As #8
182     Print #8, v
184     Close #8
    
186     FileCopy mfile, "C:\" + F_TXTFILES + "\apantSendCANC.XML"
    
           Dim CC As String

    CC = takefromxml(v)
  
    
    
    
    killClient = CC   ' 1
    
    
        
        Exit Function

'cmdDOMDocumentUTF8_Click_Err:
'        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.PSIFIAKO " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
'
'        Resume Next

        '</EhFooter>

End Function
  
Private Sub LOADFORM()
      F_MESS = 1 ' MSGBOX STIS APOSTOLES
        FIMDATE.Value = Now
        
            CARGrid.MarqueeStyle = 3  'SEIRA MARKARISMENH
            TIMGrid.MarqueeStyle = 3  'SEIRA MARKARISMENH
  
        '   Set elemStock = docStock.createElement("dcrnew:newDigitalClient")
        '                   ' With elemStock----------------------------------2002-05-30T09:00:00
        '
        '
        '
        '                      Set elem2Field = docStock.createElement("dcrnew:clientServiceType"): elem2Field.Text = garaz3epark2: elemStock.appendChild elem2Field
        '
        '                  'HTAN OK ALLA EBGALE MHNYMA (NEXT LINE)     Set elem2Field = docStock.createElement("dcrnew:creationDateTime"): elem2Field.Text = Format(Now, "yyyy-MM-ddTHH:mm:ss"): elemStock.appendChild elem2Field
        '                  '<message>creationDateTime is allowed only when transmissionFailure has value</message>
        '
        '
        '                      Set elem2Field = docStock.createElement("dcrnew:entityVatNumber"): elem2Field.Text = afmCompany: elemStock.appendChild elem2Field
        '                      Set elem2Field = docStock.createElement("dcrnew:branch"): elem2Field.Text = "0": elemStock.appendChild elem2Field
        '
        '
        '
        '
        '118                  Set elemField = docStock.createElement("dcrnew:useCase")
        '                         Set elem1Field = docStock.createElement(TYPOS)
        '
        '                              Set elem2Field = docStock.createElement("dcrnew:vehicleRegistrationNumber"): elem2Field.Text = ark: elem1Field.appendChild elem2Field
        '                              Set elem2Field = docStock.createElement("dcrnew:vehicleCategory"): elem2Field.Text = vCategory: elem1Field.appendChild elem2Field
        '                              Set elem2Field = docStock.createElement("dcrnew:vehicleFactory"): elem2Field.Text = vFACTORY: elem1Field.appendChild elem2Field
        '
                 
        F_ISCAR = Val(FINDPARAMETROI(1, "PELAT1", "F_ISCAR", "0", "”YNER√EIA=1 ÎıÌÙÁÒÈ·=2 œ◊…=0")) '
        F_artam = FINDPARAMETROI(1, "PELAT1", "F_ARTAM", ";", "¡—…»Ãœ” SN 1ÁÚTAMEIAKH”;2ÁÚÙ·Ï/÷œ—OÀO√IKOY")
        Dim error As Integer
        TAMFOROL.Clear
        If InStr(F_artam, ";") > 0 Then
           TAMFOROL.AddItem Split(F_artam, ";")(0)    '''''''''''''''''''
           TAMFOROL.AddItem Split(F_artam, ";")(1)
        Else
          TAMFOROL.AddItem F_artam
        
        End If
        
        On Error Resume Next
        TAMFOROL.Text = TAMFOROL.List(0)
        
        'products_id,options_id,options_values_id

        '================================ PRODUCT_ATTRIBUTES ================================
      
        Dim Q     As String
        Dim R     As New ADODB.Recordset
        R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'CARS'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next

        If R(0) = 0 Then
            
            Q = "CREATE TABLE CARS ( NEWDIGITALCLIENTID VARCHAR(16) NOT NULL,CLIENTSERVICETYPE INT,ENTITYVATNUMBER VARCHAR(10),ID  INT IDENTITY(1,1),"
            Q = Q + "VEHICLEREGISTRATIONNUMBER VARCHAR(10),VEHICLECATEGORY VARCHAR(20) ,VEHICLEFACTORY VARCHAR(20) ) "
            Gdb.Execute Q
        End If

        R.Close

        R.Open "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'CARRELATIONS'", Gdb, adOpenDynamic, adLockOptimistic
        'On Error Resume Next
        Q = "create TABLE CARRELATIONS (CORRELATEDID NVARCHAR(15),ID INT IDENTITY(1,1),MARK NVARCHAR(15),FIMNUMBER NVARCHAR(11),FIMAA INT,FIMISSUEDATE NVARCHAR(25),CORRELATEDDCLIDS NVARCHAR(15)      )"

        If R(0) = 0 Then
            ' Dim Q As String
            Gdb.Execute Q
        End If

        R.Close

        ADD_FIELD "CARS", "SXETIKA", "text" '"VARCHAR(250)"
        ADD_FIELD "CARS", "ENTRYCOMPLETION", "INT DEFAULT (0)"
        ADD_FIELD "CARS", "COMPLETIONDATETIME", "DATETIME"
        ADD_FIELD "CARS", "CANCELED", "INT DEFAULT (0)"
        ADD_FIELD "CARS", "CANCELEDID", "VARCHAR(15)"
        ADD_FIELD "CARS", "HME", "DATETIME"
        ADD_FIELD "CARS", "COMMENTS", "VARCHAR(120)"
        
        
        ADD_FIELD "CARS", "AFM", "VARCHAR(10)"
         ADD_FIELD "CARS", "COMPLETIONID", "VARCHAR(15)"
        ADD_FIELD "CARS", "PROVIDEDSERVICECATEGORY", "INT DEFAULT(0)"
        ADD_FIELD "CARS", "INVOICEKIND", "INT DEFAULT (0)"
        ADD_FIELD "CARS", "STATUS", "INT DEFAULT (1)" ' 1=”≈ ≈  —≈Ãœ‘«‘¡   2=≈–…” ≈’¡”‘« ≈   3=–«√≈ ”≈ ¡ÀÀœÕ
        ADD_FIELD "CARS", "POSO", "NUMERIC(10,2)"
        'clientServiceType"): elem2Field.Text = "2": elemStock.appendChild elem2Field
        '                      Set elem2Field = docStock.createElement("dcrudt:entryCompletion"): elem2Field.Text = "true": elemStock.appendChild elem2Field
        '                      Set elem2Field = docStock.createElement("dcrudt:completionDateTime"): elem2Field.Text = "202502221033: "

f_parking_auto = Val(FINDPARAMETROI(3, "PAR1", "f_parking_auto", "0", "ÂÈÌ·È parking ÏÂ ·ıÙ¸Ï·ÙÔ ¯ÁˆÈ·ÍÔ ÂÎ·ÙÔÎ¸„ÈÔ=1 /Ô˜È=0"))

If f_parking_auto = 1 Then
   autoClose.Visible = True
   

End If



        cars.ConnectionString = gConnect
        
        Dim MTYPOS As Integer
        If F_ISCAR = 1 Then MTYPOS = 52 Else MTYPOS = 53
        

        cars.RecordSource = "SELECT VEHICLEREGISTRATIONNUMBER AS [¡—. ’ À] ,VEHICLECATEGORY AS [KATH√],VEHICLEFACTORY AS [≈—√œ”‘],PINAKES.PERIGRAFH ," _
        & "(CASE INVOICEKIND WHEN 1 THEN '¡À–' WHEN 2  THEN '‘…Ã' WHEN 3  THEN '÷«Ã' ELSE ' ' END) AS [–¡—¡”‘], " _
        & "ISNULL(ENTRYCOMPLETION,0) AS [œ /1 ≈  /0],COMMENTS AS [”◊œÀ…¡],* FROM CARS LEFT JOIN PINAKES ON PINAKES.TYPOS=" + str(MTYPOS) + " AND PROVIDEDSERVICECATEGORY=TIMH " _
        & " WHERE ISNULL(CANCELED,0)=0 AND ISNULL(ENTRYCOMPLETION,0)=0 ORDER BY CARS.ID DESC"
        'TIMGrid.columns(0).DataWidth = 3600
        ' = 3600
        cars.Refresh
        CARGrid.columns(1).Width = 3600

        TIMGrid.Refresh

        If F_ISCAR >= 1 Then
            ARKYKL.Visible = True
            VCAT.Visible = True
            VFACT.Visible = True
            cmdKatax.Visible = True
            'LARKYKL.Visible = True
            'lblCAT.Visible = True
            'Label3.Visible = True
       
        End If

        timologia.ConnectionString = gConnect
        Dim sql As String
        sql = "SELECT top 5000 HME,ATIM,AJI AS [AŒ…¡],PEL.EPO AS [≈–ŸÕ’Ã…¡],PEL.KOD AS [ Ÿƒ… œ”],PARAT AS [”◊œÀ…¡],SXETIKO AS [”◊≈‘.–¡—],ID_NUM,ENTITYMARK,C3,INCMARK AS [MARK ¡ ’—Ÿ”«”] FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD  ORDER BY ID_NUM DESC"
        timologia.RecordSource = sql '"SELECT HME,ATIM,ENTITYMARK FROM TIM ORDER BY ID_NUM DESC"
        timologia.Refresh

794     R.Open "SELECT *FROM PINAKES WHERE TYPOS=50 ORDER BY PERIGRAFH", Gdb, adOpenDynamic, adLockOptimistic

796     Do While Not R.EOF
798         VCAT.AddItem R("PERIGRAFH") ' Right(" " + Format(R!AYJON, "##"), 2) + "." + R("PERIGRAFH")
800         R.MoveNext
        Loop

802     If VCAT.ListCount < 1 Then

804         Gdb.Execute "DELETE FROM PINAKES WHERE TYPOS=50"

806         Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (1,'¡’‘œ …Õ«‘œ',50)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (2,'÷œ—‘«√¡ …',50)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (3,'‘—¡ ‘≈—',50)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (4,'‘—≈…À≈—',50)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (5,'Ãœ‘œ”’ À≈‘¡',50)"
             
        End If
        '       ¡’‘œ …Õ«‘œ
        '÷œ—‘«√¡ …
        '‘—¡ ‘≈—
        '—≈…À≈—
        ' ¡—œ‘”¡
        R.Close

        R.Open "SELECT *FROM PINAKES WHERE TYPOS=51 ORDER BY PERIGRAFH", Gdb, adOpenDynamic, adLockOptimistic

1796    Do While Not R.EOF
1798        VFACT.AddItem R("PERIGRAFH") ' Right(" " + Format(R!AYJON, "##"), 2) + "." + R("PERIGRAFH")
1800        R.MoveNext
        Loop

1802    If VFACT.ListCount < 1 Then

1804        Gdb.Execute "DELETE FROM PINAKES WHERE TYPOS=51"

1806        Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (1,'CITROEN',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (2,'RENAULT',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (3,'PEUGEOT',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (4,'OPEL',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (5,'TOYOTA',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (6,'NISSAN',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (7,'HYUNDAI',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (8,'FORD',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (9,'SUZUKI',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (10,'ALFA ROMEO',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (11,'AUDI',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (12,'CHEVROLET',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (13,'DAIHATSU',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (14,'FIAT',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (15,'HONDA',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (16,'JAGUAR',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (17,'JEEP',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (18,'KIA',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (19,'MAZDA',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (20,'MERCEDES-BENZ',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (21,'MITSUBISHI',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (22,'SEAT',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (23,'SKODA',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (24,'SMART',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (25,'SUBARU',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (26,'VOLKSWAGEN',51)"
            Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (27,'VOLVO',51)"
            
            
            
            
            
            
            
            
            
            
        End If

        R.Close

        'CITROEN
        'RENAULT
        'PEUGEOT
        'OPEL
        'TOYOTA
        'NISSAN
        'HUYNDAI
        'FORD
        'HONDA
        'MITSUBISHI
        '..

        'ÛıÌÂÒ„ÂÈ·
        '1 = ≈Ò„·Ûﬂ· ÏÂ ˜ÒﬁÛÁ ·ÌÙ·ÎÎ·ÍÙÈÍ˛Ì (”ıÌÂÒ„Âﬂ·)
        '2 = ≈Ò„·Ûﬂ· ÏÂ ˜ÒﬁÛÁ ·ÌÙ·ÎÎ·ÍÙÈÍ˛Ì ÂÎ‹ÙÁ (”ıÌÂÒ„Âﬂ·)
        '3 = ≈Ò„·Ûﬂ· ˜˘ÒﬂÚ ˜ÒﬁÛÁ ·ÌÙ·ÎÎ·ÍÙÈÍ˛Ì (”ıÌÂÒ„Âﬂ·)
        '4 = ƒ˘ÒÂ‹Ì ıÁÒÂÛﬂ· (–‹ÒÍÈÌ„Í/ –ÎıÌÙﬁÒÈ·,”ıÌÂÒ„Âﬂ·)
        '5 = ÀÔÈ‹ (–‹ÒÍÈÌ„Í/ –ÎıÌÙﬁÒÈ·,”ıÌÂÒ„Âﬂ·)
        '6 = ¡ÔÊÁÏﬂ˘ÛÁ –·ÒÔ˜ﬁÚ ≈„„˝ÁÛÁÚ
        '7 = ’ÁÒÂÛﬂ· ¬‹ÛÂÈ ‘ÈÏÔÍ·Ù·Î¸„Ôı (–‹ÒÍÈÌ„Í/ –ÎıÌÙﬁÒÈ·)
        '8 = ’ÁÒÂÛﬂ·  ·Ù¸ÈÌ ”ıÏˆ˘Ìﬂ·Ú (–‹ÒÍÈÌ„Í/ –ÎıÌÙﬁÒÈ·)
        '9 = …‰È¸˜ÒÁÛÁ (–‹ÒÍÈÌ„Í/ –ÎıÌÙﬁÒÈ·,”ıÌÂÒ„Âﬂ·)
        If F_ISCAR = 1 Then
160         R.Open "SELECT * FROM PINAKES WHERE TYPOS=52 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
            Dim i As Integer
            i = 0
170         Do While Not R.EOF
                i = i + 1
180             providedServiceCategory.AddItem str(R!AYJON) + ";" + R!PERIGRAFH
            
190             R.MoveNext
            Loop
            If i = 0 Then
                Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (52,1,'1 ≈Ò„·Ûﬂ· ÏÂ ·ÌÙ·ÎÎ·ÍÙÈÍ‹',1 )"
                Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (52,2,'2 ≈Ò„·Ûﬂ· ÏÂ ·ÌÙ·ÎÎ·ÍÙÈÍ‹ ÂÎ‹ÙÁ',2 )"
                Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (52,3,'3 ≈Ò„·Ûﬂ· ˜˘ÒﬂÚ ·ÌÙ·ÎÎ·ÍÙÈÍ‹ ',3 )"
                Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (52,4,'4 ƒ˘ÒÂ‹Ì ıÁÒÂÛﬂ· ',4 )"
                Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (52,5,'5 ÎÔÈ· ',5)"
                Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (52,6,'6 ¡ÔÊÁÏﬂ˘ÛÁ ≈„„˝ÁÛÁÚ ',6 )"
                Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (52,7,'9 …‰È¸˜ÒÁÛÁ ',9 )"
           
            End If
            R.Close
        ElseIf F_ISCAR = 2 Then

            'ÎıÌÙÁÒÈ·
        

1160        R.Open "SELECT * FROM PINAKES WHERE TYPOS=53 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
           
            i = 0
1170        Do While Not R.EOF
                i = i + 1
1180            providedServiceCategory.AddItem str(R!AYJON) + ";" + R!PERIGRAFH
            
1190            R.MoveNext
            Loop
            If i = 0 Then
                Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (53,1,'4 ƒ˘ÒÂ‹Ì ıÁÒÂÛﬂ· ',4 )"
                Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (53,2,'5 ÀÔÈ· ',5 )"
                Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (53,3,'7 ‘ÈÏÔÍ·Ù‹ÎÔ„ÔÚ ',7 )"
                Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (53,4,'8 ”ıÏˆ˘Ìﬂ· ',8)"
                Gdb.Execute "insert into PINAKES (TYPOS,AYJON,PERIGRAFH,TIMH) VALUES (53,4,'9 …‰È¸˜ÒÁÛÁ ',9 )"
        
            End If
            R.Close

        End If
updateScreen
End Sub
  
  
Private Sub READXML(ByVal XX As String)
  

'ƒ…¡¬¡∆≈… ‘œ REQUEST KAI RIXNEI TO CARRELATIONS
       
       Dim objxml As MSXML2.DOMDocument

154     Set objxml = New MSXML2.DOMDocument
    
156     objxml.async = True
158     objxml.loadXML v
    
        Dim r0       As New ADODB.Recordset

        Dim mFound   As Integer

        Dim nodeList As IXMLDOMNodeList

        Dim Node     As IXMLDOMNode

        Dim m_ID     As String

        'Gdb.Open "DSN=MERCSQL;"
160     Set nodeList = objxml.SelectNodes("RequestedDoc/clientcorrelationsRequestsDoc/ClientCorrelationsRequest") '"RequestedDoc/clientsDoc/DigitalClient")  'clientcorrelationsRequestsDoc/ClientCorrelationsRequest")

162   On Error GoTo 0

     For Each Node In nodeList

            ' Print Node.nodeName  ' this works'
164         List1.AddItem Node.nodeName + "*******" + Node.Text  ' INVOICE
       
            '       Print node.n
166        ' Call printNode(Node)     'here is the problem explained below'
    
168         If Len(Trim(payaji)) = 0 Then payaji = "0"
170         If Len(Trim(totalGrossValue)) = 0 Then totalGrossValue = "0"
    
172         r0.Open "select * from APESTALMENA WHERE AFM='" + AFM0 + "' AND ISSUEDATE='" + issueDate + "' AND TYPOS='" + invoiceType + "' AND AA='" + AA + "'", Gdb, adOpenDynamic, adLockOptimistic
174         FOUND = 0

176         If r0.EOF Then
        
            Else

178             If IsNull(r0("ID")) Then
          
                Else
180                 FOUND = 1
182                 m_ID = str(r0("ID"))
                End If
            End If
        
184         If FOUND = 0 Then
186             sql = "INSERT INTO APESTALMENA  ([MARK],[AFM],[ISSUEDATE],[AA],[TYPOS],[PAYTYPE],[PAYAJI],[TOTALNETVALUE],[TOTALVATAMOUNT],[TOTALGROSSVALUE],[CLASSIFICATIONTYPE])"
188             sql = sql + " Values ('" + MARK + "','" + AFM0 + "','" + issueDate + "','" + AA + "','" + invoiceType + "','" + paytype + "'," + Replace(payaji, ",", ".") + "," + Replace(totalNetValue, ",", ".") + "," + Replace(totalVatAmount, ",", ".") + "," + Replace(totalGrossValue, ",", ".") + ",'" + classificationType + "')"
190             Gdb.Execute sql
            Else
192             Gdb.Execute "update APESTALMENA SET MARK='" + MARK + "',PAYTYPE='" + paytype + "',PAYAJI=" + Replace(payaji, ",", ".") + ",TOTALNETVALUE=" + Replace(totalNetValue, ",", ".") + ",TOTALVATAMOUNT=" + Replace(totalVatAmount, ",", ".") + ",TOTALGROSSVALUE=" + Replace(totalGrossValue, ",", ".") + ",CLASSIFICATIONTYPE='" + classificationType + "' WHERE ID=" + m_ID
        
            End If
        
194         r0.Close
        
196     Next Node
    
End Sub


Private Sub take2fromxml(ByVal v As String, ByRef MARK, ByRef ID, ByRef c2)

        On Error GoTo TAKEFROMXML_Err

100    ' List1.Clear

        Dim RECS As Integer
        Dim CCV  As String
        Dim C1, c3, c4, CURL, SXOLIA As String
        Dim objxml As New MSXML2.DOMDocument
102     objxml.async = True
104     objxml.loadXML v
        Dim nLathosApost As Long
106     nLathosApost = 0
        Dim nApost As Long, nAll As Long
        
108     nApost = 0
110     nAll = 0

        Dim nodeList As IXMLDOMNodeList

        Dim Node     As IXMLDOMNode
    
112     Set nodeList = objxml.SelectNodes("ResponseDoc/response")

114     For Each Node In nodeList
        
            Dim r2       As New ADODB.Recordset

            Dim xmlNode  As IXMLDOMNode

            Dim xml2Node As IXMLDOMNode

            Dim xml3Node As IXMLDOMNode
        
116         If Node.hasChildNodes Then
                'List1.AddItem node.childNodes.Item(0)
118             C1 = "": c2 = "": c3 = ""

120             For Each xmlNode In Node.childNodes

                    ' List1.AddItem xmlNode.nodeName
            
122                 If xmlNode.nodeName = "index" Then
124                     C1 = xmlNode.Text
                        ' List1.AddItem c1
                    End If
            
126                 If xmlNode.nodeName = "invoiceUid" Then
128                     c4 = xmlNode.Text
                        'List1.AddItem c2
                    End If
            
130                 If xmlNode.nodeName = "newClientDclID" Then
132                     c2 = xmlNode.Text
                        'List1.AddItem c2
                    End If

                    'clientCorrelationID
134                 If xmlNode.nodeName = "clientCorrelationID" Then
136                     c2 = xmlNode.Text
                        'List1.AddItem c2
                    End If

                    'updatedClientDclID
138                 If xmlNode.nodeName = "updatedClientDclID" Then
140                     c2 = xmlNode.Text
                        'List1.AddItem c2
                    End If
                
                    'cURL
142                 If xmlNode.nodeName = "qrUrl" Then
144                     CURL = xmlNode.Text
                        'List1.AddItem c2
                    End If
                
146                 CCV = str(F_ID_NUMS(Val(C1)))
                
                    ' ·Ì ÂÒ˜ÂÙ·È ·¸ ÙÁÌ Â·Ì·ÔÛÙÔÎﬁ ·¸ ·Ò˜ÂﬂÔ Ù¸ÙÂ Ì· ·ﬂÒÌÂÈ ÙÔ id_num Ôı ÎÁÍÙÒÔÎ¸„ÁÛ·
148                 If Val(file_idnum) > 0 Then
150                     CCV = file_idnum
                    End If
            
152                 If xmlNode.nodeName = "errors" Then
            
154                     If xmlNode.hasChildNodes Then  'CHILD=ERROR
                            'List1.AddItem XMLnode.childNodes.Item(0)
                  
156                         For Each xml2Node In xmlNode.childNodes  'ERROR
                        
158                             If xml2Node.hasChildNodes Then 'CHILD=MESSAGE

160                                 For Each xml3Node In xml2Node.childNodes

162                                     If xml3Node.nodeName = "message" Then
164                                         SXOLIA = xml3Node.Text
                                         
166                                         SXOLIA = Left(xml3Node.Text, 40)
168                                         SXOLIA = Replace(SXOLIA, "'", "`")

                                             On Error GoTo TAKEFROMXML_Err

                                            ' GDBEXE "UPDATETIMWITHMARK", "UPDATE TIM SET ENTITY=0,ENTITYUID='" + SXOLIA + "'  WHERE   ID_NUM=" + CCV   ' ENTLINEN=" + str(fLong) + " and ENTITY=" + C1
170                                         'List1.AddItem "  " + SXOLIA
                                        End If

172                                     If xml3Node.nodeName = "code" Then
174                                         SXOLIA = xml3Node.Text + " ************* À¡»œ” ****************"
176                                        ' List1.AddItem "  " + SXOLIA
                                        End If

                                    Next
                        
                                End If

                            Next

                        End If
               
                    End If
            
178                 nAll = nAll + 1

                    On Error Resume Next

                    On Error GoTo TAKEFROMXML_Err

180             Next xmlNode
      
            End If
        
        Next

        '1111
        'DIAXOR txmlapost

182     If nApost = 1 Then
            '  GDBEXE "UPDATETIMWITHMARK", "UPDATE TIM SET XMLTEXT='" + txmlapost + "' WHERE ID_NUM=" + CCV
            ' Gdb.Execute "UPDATE TIM SET XMLMYDATA='" + txmlapost + "' WHERE ID_NUM=" + CCV

        End If

        'GoTo 1111

184     'List1.AddItem "–ﬁÒ·Ì Ã¡—  " + str(nApost)

186     'List1.AddItem "À‹ËÔÚ , ˜˘ÒﬂÚ  Ã¡—  " + str(nLathosApost)

188     'takefromxml = C2

        '<EhFooter>
        Exit Sub

TAKEFROMXML_Err:

        If noask = 0 Then
            MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.UpdateTIMwithMARK " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
        End If

        Resume Next

        '</EhFooter>

End Sub
  
  
  
Private Sub updateScreen()
   On Error Resume Next
   If CARGrid.columns("INVOICEKIND") = 3 Then '‘¡Ã≈…¡ «  ÷œ—œÀœ√… œ” CARGrid.Text  '153
      Frame1.Visible = True
      TIMGrid.Visible = False
      cmdSYSX.Visible = False
      ccmdCloseInvoice.Visible = False
      cmdCLOSEcLIENT.Visible = True
      
   ElseIf CARGrid.columns("INVOICEKIND") = 2 Then  ' TIMOLOGIA
      cmdCLOSEcLIENT.Visible = False
      
      Frame1.Visible = False
      TIMGrid.Visible = True
      cmdSYSX.Visible = True
      ccmdCloseInvoice.Visible = True
      Dim AFM As String, mKPE As String, synt As String
      AFM = CARGrid.columns("AFM")
     ' MKPE=GGET_CvAlUE("SELECT * FROM PEL WHERE AFM
      'SELECT * FROM TIM WHERE KPE IN (SELECT KOD FROM PEL WHERE EIDOS='e' AND AFM='"+AFM+"') ORDER BY ID_NUM DESC"
      synt = " WHERE KPE IN (SELECT KOD FROM PEL WHERE EIDOS='e' AND AFM='" + AFM + "') "
      Dim sql As String
        sql = "SELECT top 20 HME,ATIM,AJI AS [AŒ…¡],PEL.EPO AS [≈–ŸÕ’Ã…¡],PEL.KOD AS [ Ÿƒ… œ”],PARAT AS [”◊œÀ…¡],SXETIKO AS [”◊≈‘.–¡—],ID_NUM,ENTITYMARK,C3,INCMARK AS [MARK ¡ ’—Ÿ”«”] FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD " + synt + "  ORDER BY ID_NUM DESC"
        
        
     If Len(cURRTIM.Caption) > 2 Then
        ' DIALEJE EIDH MHN TOY XALAS TO FOCUS
     Else
        
        timologia.RecordSource = sql '"SELECT HME,ATIM,ENTITYMARK FROM TIM ORDER BY ID_NUM DESC"
        timologia.Refresh
     End If
      
   
   Else
   
   
   
   
      Frame1.Visible = False
      TIMGrid.Visible = False
      cmdSYSX.Visible = False
      If providedServiceCategory = 4 Then
         ccmdCloseInvoice.Visible = True
      Else
          ccmdCloseInvoice.Visible = False
      End If
     
   End If
   
   
   
End Sub

       



Private Sub afmonoma_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       search_Click
    End If
       
       
End Sub

Private Sub ananeosi_Click()
    LOADFORM
    cars.Refresh
    
    
End Sub

Private Sub ANTISTTIMOL_Click()

On Error Resume Next


   Dim cATIM, DIDCL, e As String
   Dim MARK As String
 
 
        TIMGrid.Col = 1
        cATIM = TIMGrid.Text  'Y000123 PARASTATIKO
      
      If Len(Trim$(cATIM)) = 0 Then
      
          MsgBox " ƒ≈Õ ’–¡—◊≈… –¡—¡”‘¡‘… œ"
          Exit Sub
          
      End If
   Dim c3 As String
   
   c3 = GGET_CVALUE("SELECT TOP 1 C3 FROM TIM WHERE ATIM='" + cATIM + "'")
       'c3 = TIMGrid.columns("C3")  DEN GINOTAN ONLINE
       
       If Len(c3) > 0 Then
          MsgBox "«ƒ« ”’”◊≈‘…”‘« ≈"
          Exit Sub
         
       End If
     
     
        TIMGrid.Col = 8
      MARK = TIMGrid.Text ' 4000001023334
      If Left(MARK, 3) = "ƒœ " Then
          MARK = mID(MARK, 4, 15)
      End If
      
      Dim ID As String
         
         ' CARGrid.Col = 3
        ID = CARGrid.columns("ID") 'CARGrid.Text  '153
         
         
         
         
         
       '  CARGrid.Col = 0
        DIDCL = CARGrid.columns("NEWDIGITALCLIENTID") ' CARGrid.Text  '100000000008565
        
      
        
       ' TIMGrid.Col = 9
    '  MARK = TIMGrid.Text
        
        
        loadpar72
vid = RELATECLIEND(F_MESS, MARK, DIDCL, 1, "", "", "")
 'vid = "100000000008565"
        
        Dim mPoso As String
        mPoso = Format(TIMGrid.columns(2), "####0.00")
        mPoso = Replace(mPoso, ",", ".")
        
 If Len(Trim$(vid)) > 10 Then
      Gdb.Execute "UPDATE TIM SET  C3='" + ID + "'  where ATIM='" + cATIM + "'"
        
    Gdb.Execute "UPDATE CARS SET  ENTRYCOMPLETION=1, POSO=" + mPoso + ", SXETIKA='" + MARK + "'  where ID=" + str(ID)
    
    Gdb.Execute "INSERT INTO CARRELATIONS(FIMISSUEDATE,CORRELATEDDCLIDS,CORRELATEDID,MARK) VALUES(CONVERT(CHAR(10),GETDATE(),103),'" + DIDCL + "','" + vid + "','" + MARK + "')"
    TIMGrid.Refresh
    Else
    
    MsgBox "ƒ≈Õ ¡Õ‘…”‘œ…◊«»« ≈"
 End If
      'Gdb.Execute "UPDATE CARS SET   SXETIKA=ISNULL(SXETIKA,'')+'" + MARK + "'+' " + cATIM + "'+' " + str(ID) + " ' where ID=" + str(ID)
         '  Dim c As String

      '  PAR1.TDBGrid1.Col = 2
     '   c = Replace(PAR1.TDBGrid1.Text, ",", ".")
        
     ' cars.Refresh
      
End Sub

Private Sub autoClose_Click()


Dim ANS As Integer, synt As String
ANS = MsgBox("ÃœÕœ ‘¡ ◊»≈”…Õ¡;", vbYesNo)
If ANS = vbYes Then
  synt = " AND DAY(HME)<DAY(GETDATE()) "
Else
   synt = " "
End If


 cars.RecordSource = "SELECT VEHICLEREGISTRATIONNUMBER AS [¡—. ’ À] ,VEHICLECATEGORY AS [KATH√],VEHICLEFACTORY AS [≈—√œ”‘],PINAKES.PERIGRAFH ," _
        & "(CASE INVOICEKIND WHEN 1 THEN '¡À–' WHEN 2  THEN '‘…Ã' WHEN 3  THEN '÷«Ã' ELSE ' ' END) AS [–¡—¡”‘], " _
        & "ISNULL(ENTRYCOMPLETION,0) AS [œ /1 ≈  /0],COMMENTS AS [”◊œÀ…¡],* FROM CARS LEFT JOIN PINAKES ON PINAKES.TYPOS=" + str(MTYPOS) + " AND PROVIDEDSERVICECATEGORY=TIMH " _
        & " WHERE INVOICEKIND=3  AND  ISNULL(CANCELED,0)=0 AND ISNULL(ENTRYCOMPLETION,0)=0 " + synt + " ORDER BY CARS.ID DESC"
        'TIMGrid.columns(0).DataWidth = 3600
        ' = 3600
        cars.Refresh
        
        

 Dim cATIM, DIDCL, e As String
   Dim MARK As String
 

      
      Dim ID As String


          
      loadpar72


  Dim l2 As Integer
  l2 = cars.Recordset.RecordCount
  
  For L = 1 To l2
         'OI METABLHTES –œ’ ◊—«”…Ãœ–œ…œ’Õ‘¡… :DIDCL,ID,INVC,servC,SX
         
         CARGrid.row = 0
   
         ID = CARGrid.columns("ID")
         
        ' CARGrid.Col = 0
          DIDCL = CARGrid.columns("NEWDIGITALCLIENTID")
       Dim INVC As String
      INVC = CARGrid.columns("INVOICEKIND")
      
      
      ' Dim providedServiceCategory As String
       '    providedServiceCategory = CARGrid.columns("providedServiceCategory")
       
        Dim servC As String: servC = CARGrid.columns("providedServiceCategory")
       
       
       
       
              
       Dim SX As String
        '  CARGrid.Col = 5
        'SX = CARGrid.Text  '153
     SX = CARGrid.columns("SXETIKA")
     Dim nc As Integer
      nc = GGET_NVALUE("SELECT COUNT(*) FROM CARRELATIONS WHERE CORRELATEDDCLIDS='" + DIDCL + "'")
      
      
        '¡Õ ƒ≈Õ ≈…Õ¡… ƒŸ—≈¡Õ Õ¡ ≈◊≈… ”’”◊≈‘…”«
        If nc = 0 And Val(INVC) > 0 Then ' Len(SX) = 0 Then
        
          MsgBox ("ƒ≈Õ ”’”◊≈‘…”‘« ≈ Ã≈ Ã¡—  ‘œ œ◊«Ã¡")
          Exit Sub
        
        
        
        End If
        
        
        Dim CPOSO As String
           CPOSO = GGET_CVALUE("SELECT POSO FROM CARS WHERE ID='" + ID + "'")
           CPOSO = Replace$(CPOSO, ",", ".")
           
  ' 3=synergeia
    Dim MT As String
    If F_ISCAR = 1 Then MT = "3" Else MT = "2"
             vid = fupdateclientFIM(0, DIDCL, MT, Format(Now, "yyyy-MM-ddTHH:mm:ss"), CPOSO, servC) '"yyyyMMddThhmmss"))

        
   If Len(vid) > 10 Then
              Gdb.Execute "UPDATE CARS SET ENTRYCOMPLETION=1,COMPLETIONID='" + vid + "'  where ID=" + ID '  SXETIKA='**'+ISNULL(SXETIKA,'')
   Else
   
          Gdb.Execute "UPDATE CARS SET ENTRYCOMPLETION=3  where ID=" + ID '  SXETIKA='**'+ISNULL(SXETIKA,'')
          ListEpo.Visible = True
          ListEpo.AddItem CARGrid.columns(0) + "*****OXI "
             ' MsgBox "ƒ≈Õ ≈ À≈…”≈ « ≈√√—¡÷«"
   
   End If
   
  ' cars.Recordset.MoveNext
    cars.Refresh
Next

'¡’‘¡ –œ’ ƒ≈Õ  –≈—¡”¡Õ ‘¡ Œ¡Õ¡÷≈—ÕŸ ”≈ ≈  —≈Ãœ‘«‘¡
Gdb.Execute "UPDATE CARS SET ENTRYCOMPLETION=0 WHERE ENTRYCOMPLETION=3"


 cars.RecordSource = "SELECT VEHICLEREGISTRATIONNUMBER AS [¡—. ’ À] ,VEHICLECATEGORY AS [KATH√],VEHICLEFACTORY AS [≈—√œ”‘],PINAKES.PERIGRAFH ," _
        & "(CASE INVOICEKIND WHEN 1 THEN '¡À–' WHEN 2  THEN '‘…Ã' WHEN 3  THEN '÷«Ã' ELSE ' ' END) AS [–¡—¡”‘], " _
        & "ISNULL(ENTRYCOMPLETION,0) AS [œ /1 ≈  /0],COMMENTS AS [”◊œÀ…¡],* FROM CARS LEFT JOIN PINAKES ON PINAKES.TYPOS=" + str(MTYPOS) + " AND PROVIDEDSERVICECATEGORY=TIMH " _
        & " WHERE ISNULL(CANCELED,0)=0 AND ISNULL(ENTRYCOMPLETION,0)=0 ORDER BY CARS.ID DESC"
        'TIMGrid.columns(0).DataWidth = 3600
        ' = 3600
        cars.Refresh

      



End Sub

Private Sub CARGrid_BeforeRowColChange(Cancel As Integer)
     updateScreen

End Sub

Private Sub CARGrid_Click()
  
  updateScreen
  
  
  End Sub

Private Sub CARGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    updateScreen
End Sub

Private Sub ccmdCloseInvoice_Click()



   ' DIABAZO TIS METABLHTES APO TO GRID-------------------------------------------



  Dim cATIM, DIDCL, e As String
   Dim MARK As String
      Dim ID As String
          CARGrid.Col = 3
         ID = CARGrid.columns("ID")
         CARGrid.Col = 0
          DIDCL = CARGrid.columns("NEWDIGITALCLIENTID")
          
          Dim MAFM As String
          MAFM = CARGrid.columns("AFM")
          
          
          Dim mINVOICEKIND As String
          mINVOICEKIND = CARGrid.columns("INVOICEKIND")
          
           Dim mprovidedServiceCategory As String
          mprovidedServiceCategory = CARGrid.columns("providedServiceCategory")
           Dim iprovidedServiceCategory As Integer: iprovidedServiceCategory = Val(mprovidedServiceCategory)
      Dim iInvoiceKind As Integer: iInvoiceKind = Val(mINVOICEKIND)
       'SXETIKA
       
       
      '----------------- ¬À≈–Ÿ ¡Õ ≈◊Ÿ  ¡Õ≈… ”’”◊≈‘…”≈…” ------------------------------------------------
      
       
       Dim SX As String
     SX = CARGrid.columns("SXETIKA")
     Dim nc As Integer
      nc = GGET_NVALUE("SELECT COUNT(*) FROM CARRELATIONS WHERE CORRELATEDDCLIDS='" + DIDCL + "'")
        If nc = 0 Then ' Len(SX) = 0 Then
          Me.Caption = "ƒ≈Õ ”’”◊≈‘…”‘« ≈ Ã≈ Ã¡—  ‘œ œ◊«Ã¡"
          MsgBox "ƒ≈Õ ”’”◊≈‘…”‘« ≈ Ã≈ Ã¡—  ‘œ œ◊«Ã¡"
          Exit Sub
        End If
      loadpar72
     
      ' 3=synergeia
      
      
      
'--------------------------  OLOKLHRVNV THN ≈–…” ≈’« -----------------------------------------------

      
        Dim mPoso As String
     ' MPOSO = InputBox("ƒŸ”≈ ‘œ –œ”œ", "”≈ ≈’—Ÿ ¡–œƒ≈…Œ«/–œ”")
     ' MPOSO = Replace$(MPOSO, ",", ".")
      mPoso = "0.0"
      
      
      Dim MT As String
    If F_ISCAR = 1 Then MT = "3" Else MT = "2"
      
      
   If iInvoiceKind = 2 Then 'TIMOLOGIO
   
     Dim MPOSO2 As String: MPOSO2 = CARGrid.columns("POSO")
      vid = fupdateclientTIM(F_MESS, DIDCL, MT, Format(Now, "yyyy-MM-ddTHH:mm:ss"), iprovidedServiceCategory, iInvoiceKind, "true", MAFM, MPOSO2)
   ElseIf iInvoiceKind = 3 Then 'FIM
     
     
     'AYTO EJYPHRETEITE APO TO fupdateclient
     ' MPOSO = InputBox("ƒŸ”≈ ‘œ –œ”œ", "”≈ ≈’—Ÿ ¡–œƒ≈…Œ«/–œ”")
     ' MPOSO = Replace$(MPOSO, ",", ".")
     ' vid = fupdateclient(DIDCL, "3", Format(Now, "yyyy-MM-ddTHH:mm:ss"), iprovidedServiceCategory, iInvoiceKind, MPOSO)
   
   
   
   
  ElseIf iInvoiceKind = 1 Then 'FIM
      mPoso = InputBox("ƒŸ”≈ ‘œ –œ”œ", "”≈ ≈’—Ÿ ¡–œƒ≈…Œ«/–œ”")
      mPoso = Replace$(mPoso, ",", ".")
      vid = fupdateclient(F_MESS, DIDCL, MT, Format(Now, "yyyy-MM-ddTHH:mm:ss"), iprovidedServiceCategory, iInvoiceKind, mPoso)
 
   
   
   
   End If
   
   
   
   
  If Len(vid) > 10 Then
              Gdb.Execute "UPDATE CARS SET ENTRYCOMPLETION=1,PROVIDEDSERVICECATEGORY=" + str(iprovidedServiceCategory) + ",INVOICEKIND=" + str(iInvoiceKind) + " where ID=" + ID
   Else
              MsgBox "ƒ≈Õ ≈ À≈…”≈ « ≈√√—¡÷«"
   
   End If
      cars.Refresh





' GIA NA JEMPLOKAREI TO GRID
  cURRTIM.Caption = ""
  












End Sub

Private Sub chk¬À≈–ŸœÀœ À«—ŸÃ≈Õ¡_Click()
  If chk¬À≈–ŸœÀœ À«—ŸÃ≈Õ¡.Value = vbChecked Then
       TDBGrid.Height = 3800
       TDBGrid.Visible = True
         AdoCarsOLD.RecordSource = "SELECT VEHICLEREGISTRATIONNUMBER AS [¡—. ’ À] ,VEHICLECATEGORY AS [KATH√],VEHICLEFACTORY AS [≈—√œ”‘],PINAKES.PERIGRAFH ," _
        & "(CASE INVOICEKIND WHEN 1 THEN '¡À–' WHEN 2  THEN '‘…Ã' WHEN 3  THEN '÷«Ã' ELSE ' ' END) AS [–¡—¡”‘], " _
        & "ISNULL(ENTRYCOMPLETION,0) AS [œ /1 ≈  /0],* FROM CARS LEFT JOIN PINAKES ON PINAKES.TYPOS=" + str(MTYPOS) + " AND PROVIDEDSERVICECATEGORY=TIMH " _
        & " WHERE (ISNULL(CANCELED,0)=1 OR ISNULL(ENTRYCOMPLETION,0)=1) ORDER BY CARS.ID DESC"
        'TIMGrid.columns(0).DataWidth = 3600
        ' = 3600
        
        ccmdCloseInvoice.Visible = False
        cmdCLOSEcLIENT.Visible = False
        cmdRequestClients.Visible = False
        
        'AdoCarsOLD.RecordSource = "SELECT * FROM CARS"
        
         AdoCarsOLD.ConnectionString = gConnect
         AdoCarsOLD.Refresh
        'ADOCARGrid.columns(1).Width = 3600
  
  Else
        TDBGrid.Visible = False
        ccmdCloseInvoice.Visible = True
        cmdCLOSEcLIENT.Visible = True
        cmdRequestClients.Visible = True
         TDBGrid.Height = 2100
  
  End If
  



End Sub

Private Sub cmdAPOXML_Click()
   loadpar72
Dim TXML As String

        Dim objStream, strData

124     Set objStream = CreateObject("ADODB.Stream")

126     objStream.Charset = "utf-8"
128     objStream.Open
130     objStream.LoadFromFile ("c:\TXTFILES\inv2.xml")

132     strData = objStream.ReadText()

134     objStream.Close


138         TXML = strData

  




  Dim MARK As String

        '    MARK = InputBox("‰˘ÛÂ Ùo mark", "Ï·ÒÍ")
     
   '   url2 = url2 + MARKTIM
     
   '  url2 = UrlAADE + "DCL/UpdateClient"
     
     url2 = url2 + MARKTIM
     
     url2 = UrlAADE + "DCL/UpdateClient"

        'initialize
     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
    XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
        'force TLS 1.2
148     XMLServer.Option(9) = 2048
150     XMLServer.Option(6) = True
    
152 ' Exit Sub


XMLServer.Open "POST", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
154     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
156     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"


'Exit Sub


158     XMLServer.send TXML  'docStock.XML 'TXML 'strData  '”‘≈ÀÕŸ ‘œ STRING TXML STO MYDATA  'XMLServer.send TXML '  ¡Õ »≈ÀŸ Õ¡ ‘œ ”‘≈…ÀŸ ¡–œ ¡—◊≈…œ ‘œ‘≈ ¬¡∆Ÿ strData
160     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
162     v = XMLServer.responseText
        '--------------- ¡–œ”‘œÀ« INVOICES -----------------------------------------------------
 
164     MsgBox v


























End Sub

Private Sub cmdCLOSEcLIENT_Click()
   Dim cATIM, DIDCL, e As String
   Dim MARK As String
 

   ' DIABAZO TIS METABLHTES APO TO GRID-------------------------------------------
      Dim ID As String
         
        '  CARGrid.Col = 3
   
         ID = CARGrid.columns("ID")
         
        ' CARGrid.Col = 0
          DIDCL = CARGrid.columns("NEWDIGITALCLIENTID")
       Dim INVC As String
      INVC = CARGrid.columns("INVOICEKIND")
       Dim providedServiceCategory As String
       providedServiceCategory = CARGrid.columns("providedServiceCategory")
       
       Dim servC As String: servC = CARGrid.columns("providedServiceCategory")
           
              
                   
           
           
       
      '----------------- ¬À≈–Ÿ ¡Õ ≈◊Ÿ  ¡Õ≈… ”’”◊≈‘…”≈…” ------------------------------------------------
       'SXETIKA
       Dim SX As String
        '  CARGrid.Col = 5
        'SX = CARGrid.Text  '153
     SX = CARGrid.columns("SXETIKA")
     Dim nc As Integer
      nc = GGET_NVALUE("SELECT COUNT(*) FROM CARRELATIONS WHERE CORRELATEDDCLIDS='" + DIDCL + "'")
      
      
        '¡Õ ƒ≈Õ ≈…Õ¡… ƒŸ—≈¡Õ Õ¡ ≈◊≈… ”’”◊≈‘…”«
        If nc = 0 And Val(INVC) > 0 Then ' Len(SX) = 0 Then
        
          MsgBox ("ƒ≈Õ ”’”◊≈‘…”‘« ≈ Ã≈ Ã¡—  ‘œ œ◊«Ã¡")
          Exit Sub
        
        
        
        End If
        
        
        Dim CPOSO As String
           CPOSO = GGET_CVALUE("SELECT POSO FROM CARS WHERE ID='" + ID + "'")
           CPOSO = Replace$(CPOSO, ",", ".")
           
        
        
        
        
       ' TIMGrid.Col = 9
    '  MARK = TIMGrid.Text
        
      '====================================  ÷œ—‘ŸÕŸ PARAMETROYS ==================================================
      loadpar72
      
 'DVREAN
'      <dcrudt:updateClient>
'    <dcrudt:initialDclId>100000000038724</dcrudt:initialDclId>
'    <dcrudt:clientServiceType>3</dcrudt:clientServiceType>
'    <dcrudt:entryCompletion>true</dcrudt:entryCompletion>
'    <dcrudt:nonIssueInvoice>false</dcrudt:nonIssueInvoice>
'    <dcrudt:amount>10</dcrudt:amount>
'    <dcrudt:isDiffVehReturnLocation>false</dcrudt:isDiffVehReturnLocation>
'    <dcrudt:providedServiceCategory>4</dcrudt:providedServiceCategory>
'    <dcrudt:comments>----</dcrudt:comments>
'  </dcrudt:updateClient>
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
'      Dim mPoso As String
'      mPoso = InputBox("ƒŸ”≈ ‘œ –œ”œ", "”≈ ≈’—Ÿ ¡–œƒ≈…Œ«/–œ”")
'      mPoso = Replace$(mPoso, ",", ".")
'
'      mPoso = CARGrid.columns("POSO")
'




'--------------------------  OLOKLHRVNV THN ≈–…” ≈’« -----------------------------------------------

  ' 3=synergeia
    Dim MT As String
    If F_ISCAR = 1 Then MT = "3" Else MT = "2"
     
vid = fupdateclientFIM(F_MESS, DIDCL, MT, Format(Now, "yyyy-MM-ddTHH:mm:ss"), CPOSO, servC) '"yyyyMMddThhmmss"))
'vid = fupdateclientFIM(DIDCL, "3", Format(Now, "yyyy-MM-ddTHH:mm:ss"), CPOSO, INVC, providedServiceCategory) '"yyyyMMddThhmmss"))
 'vid = "100000000008565"
        
  If Len(vid) > 10 Then
              Gdb.Execute "UPDATE CARS SET ENTRYCOMPLETION=1,COMPLETIONID='" + vid + "'  where ID=" + ID '  SXETIKA='**'+ISNULL(SXETIKA,'')
   Else
              MsgBox "ƒ≈Õ ≈ À≈…”≈ « ≈√√—¡÷«"
   
   End If

      cars.Refresh
      
      
      
      
      
End Sub

Private Sub cmdFIM_Click()
  
   If Len(FIMAA.Text) = 0 Then
      MsgBox ("ƒ≈Õ ¬¡À¡‘≈ ¡—…»Ãœ ¡–œƒ≈…Œ«”")
      FIMAA.SetFocus
      Exit Sub
   End If
   
   
   If gVal(TXTPOSO.Text) = 0 Then
      MsgBox ("ƒ≈Õ ¬¡À¡‘≈ –œ”œ")
      TXTPOSO.SetFocus
      Exit Sub
   End If
   
  
  
  
  
  
  
  Dim cATIM, DIDCL, e As String
   Dim MARK As String
 
 
'        TIMGrid.Col = 1
 '       cATIM = TIMGrid.Text  'Y000123 PARASTATIKO
      
     
     
  '      TIMGrid.Col = 8
   '   MARK = TIMGrid.Text ' 4000001023334
      
      
      
      
      Dim ID As String
         
          CARGrid.Col = 3
          ID = CARGrid.columns("ID")
        'ID = CARGrid.Text  '153
         
         
         CARGrid.Col = 0
        'DIDCL = CARGrid.Text  '100000000008565
         DIDCL = CARGrid.columns("NEWDIGITALCLIENTID")
      
        
       ' TIMGrid.Col = 9
    '  MARK = TIMGrid.Text
    Dim AA As String, DD As String
    AA = FIMAA.Text
    DD = Format(FIMDATE, "yyyy-MM-dd")
    
        
        
        
        
        'Par7MyData.
        loadpar72
vid = RELATECLIEND(F_MESS, MARK, DIDCL, 2, AA, DD, F_artam)
 'vid = "100000000008565"
        mPoso = TXTPOSO.Text
        mPoso = Replace(mPoso, ",", ".")
        If Val(mPoso) = 0 Then
           mPoso = "0.0"
        End If
    Gdb.Execute "UPDATE CARS SET  POSO=" + mPoso + ",  SXETIKA='" + AA + "'+' " + DD + "'+' " + "" + " //' where ID=" + ID
        'ISNULL(SXETIKA,'')
         '  Dim c As String

'''' Gdb.Execute "UPDATE CARS SET   SXETIKA='" + MARK + "'  where ID=" + str(ID)
  MARK = ""
    Gdb.Execute "INSERT INTO CARRELATIONS(FIMISSUEDATE,CORRELATEDDCLIDS,CORRELATEDID,FIMAA) VALUES(CONVERT(CHAR(10),GETDATE(),103),'" + DIDCL + "','" + vid + "','" + FIMAA + "')"
    




      '  PAR1.TDBGrid1.Col = 2
     '   c = Replace(PAR1.TDBGrid1.Text, ",", ".")
        
      cars.Refresh
      
        
        
End Sub







Private Sub cmdKatax_Click()


 Dim nn As Integer
   nn = GGET_NVALUE("SELECT COUNT(*) FROM CARS WHERE VEHICLEREGISTRATIONNUMBER='" + ARKYKL.Text + "' AND ISNULL(ENTRYCOMPLETION,0)=0")
If nn > 0 Then
   MsgBox "’–¡—◊≈… «ƒ« ‘œ " + ARKYKL.Text + " ”≈ ≈  —≈Ãœ‘«‘¡"
   Exit Sub
End If

' 3=synergeia
    Dim MT As Integer
    If F_ISCAR = 1 Then MT = 3 Else MT = 2

Dim vid As String
loadpar72
vid = newclient2(F_MESS, MT, ARKYKL.Text, VFACT.Text, VCAT.Text)
Me.Caption = vid

If Len(vid) > 10 Then
   ' If F_ISCAR >= 1 Then
     
'   UPDATE_PEL "VEHICLEREGISTRATIONNUMBER", Left(ARKYKL.Text, Adodc1.Recordset("VEHICLEREGISTRATIONNUMBER").DefinedSize), 0, x, MKOD
'   UPDATE_PEL "VEHICLECATEGORY", Left(VCAT.Text, Adodc1.Recordset("VEHICLECATEGORY").DefinedSize), 0, x, MKOD
'   UPDATE_PEL "VEHICLEFACTORY", Left(VFACT.Text, Adodc1.Recordset("VEHICLEFACTORY").DefinedSize), 0, x, MKOD
'   UPDATE_PEL "NEWDIGITALCLIENTID", Left(vid, Adodc1.Recordset("NEWDIGITALCLIENTID").DefinedSize), 0, x, MKOD
   Gdb.Execute "INSERT INTO CARS(COMMENTS,VEHICLEREGISTRATIONNUMBER,VEHICLECATEGORY,VEHICLEFACTORY,NEWDIGITALCLIENTID,HME) VALUES ('" + txSXOLIA.Text + "','" + ARKYKL.Text + "','" + VCAT.Text + "','" + VFACT.Text + "','" + vid + "',GETDATE() )"
   
 End If
 
  cars.Refresh
ARKYKL.Text = ""
txSXOLIA.Text = ""

cmdKatax.Enabled = False '   ¯Ï‰Í·Ù·˜.ÂÌ·‚ÎÂ‰ = ˆ·ÎÛÂ


End Sub

Private Sub cmdNEA_Click()
  cmdKatax.Enabled = True
  
End Sub

Private Sub cmdOLD_Click()

    AdoCarsOLD.RecordSource = "SELECT VEHICLEREGISTRATIONNUMBER AS [¡—. ’ À] ,VEHICLECATEGORY AS [KATH√],VEHICLEFACTORY AS [≈—√œ”‘],PINAKES.PERIGRAFH ," _
        & "(CASE INVOICEKIND WHEN 1 THEN '¡À–' WHEN 2  THEN '‘…Ã' WHEN 3  THEN '÷«Ã' ELSE ' ' END) AS [–¡—¡”‘], " _
        & "ISNULL(ENTRYCOMPLETION,0) AS [œ /1 ≈  /0],* FROM CARS LEFT JOIN PINAKES ON PINAKES.TYPOS=" + str(MTYPOS) + " AND PROVIDEDSERVICECATEGORY=TIMH " _
        & " WHERE ISNULL(CANCELED,0)=0 AND ISNULL(ENTRYCOMPLETION,0)=0 ORDER BY CARS.ID DESC"
        'TIMGrid.columns(0).DataWidth = 3600
        ' = 3600
        AdoCarsOLD.Refresh
        'ADOCARGrid.columns(1).Width = 3600


End Sub

Private Sub cmdRequestClients_Click()
loadpar72
Dim URL22 As String

      If F_demo = 1 Then
160      URL22 = "https://mydataapidev.aade.gr/DCL/"
      Else
      
          URL22 = "https://mydatapi.aade.gr/DCL/"

      End If

   '=============== request client ======================================
        Dim ANS    As String, TYPOS As String

        Dim ARXID As String

100     ANS = InputBox("2=PARKING 3=GARAGE", "")

102     ARXID = InputBox("ID", , "100000000000000")

104     If Val(ANS) = 2 Then
106         TYPOS = "dcrnew:parkingcarwash"
        Else
108         TYPOS = "dcrnew:garage"
        End If
   
        '  Dim docStock As MSXML2.DOMDocument
        Dim elemRoot  As MSXML2.IXMLDOMElement

        Dim elemStock As MSXML2.IXMLDOMElement

        Dim elemField As MSXML2.IXMLDOMElement

        Dim i         As Integer
    
        ''102     Set docStock = New MSXML2.DOMDocument
        ''104     With docStock
        ''106         .appendChild .createProcessingInstruction("xml", _
        ''                                                      "version=""1.0"" encoding=""utf-8""")
        ''108         Set elemRoot = .createElement("dcrudtcor:ClientCorrelationDoc")
        ''                elemRoot.setAttribute "xmlns:dcrudtcor", "http://www.aade.gr/myDATA/dcrudtcor/v1.0"
        ''                elemRoot.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
        ''                elemRoot.setAttribute "xsi:schemaLocation", "http://www.aade.gr/myDATA/dcrudtcor/v1.0 SimpleTypes-v1.0.xsd clientCorrelationType-v1.0.xsd"
        ''
        ''116                 Set elemStock = docStock.createElement("dcrudtcor:clientCorrelation")
        ''
        ''                       Set elem2Field = docStock.createElement("dcrudtcor:entityVatNumber"): elem2Field.Text = "028783755": elemStock.appendChild elem2Field
        ''
        ''                      Set elem2Field = docStock.createElement("dcrudtcor:mark"): elem2Field.Text = "400000000000000": elemStock.appendChild elem2Field
        ''                      ' Set elem2Field = docStock.createElement("dcrudt:completionDateTime"): elem2Field.Text = "202502221033: elemStock.appendChild elem2Field"
        ''                        Set elem2Field = docStock.createElement("dcrudtcor:correlatedDCLids"): elem2Field.Text = 100000000007161#: elemStock.appendChild elem2Field
        ''
        ''126
        ''
        ''130                 elemRoot.appendChild elemStock
        ''
        ''          '  End With  /////////////////////////////////////////////////////////////////////////////////////////////
        ''132         Set .documentElement = elemRoot
        ''            On Error Resume Next
        ''
        ''             FormatXmlDocument docStock ' ﬂa?e? ?e?a ?a e??a? e????? st? d?aﬂasµa
        ''            Dim ANS2 As Integer
        ''            ANS2 = MsgBox(docStock.XML, vbYesNo)
        ''            If ANS2 = vbNo Then
        ''                Exit Sub
        ''            End If
        ''
        ''
        ''134         Kill "C:\txtfiles\" + ARKYKL + "r.xml"
        ''            On Error GoTo cmdDOMDocumentUTF8_Click_Err
        ''136         .save "C:\txtfiles\" + ARKYKL + "r.xml"
        ''        End With
        ''
        ''

        ' TO XML EINAI docStock.xml

        '--------------- ???S???? INVOICES -----------------------------------------------------
     
        Dim MARK      As String

        '    MARK = InputBox("d?se to mark", "µa??")
     
110     url2 = url2 + MARKTIM
     
112     url2 = URL22 + "RequestClients?DCLID=" + ARXID '100000000007663" ''''''100000000007739"


      If F_demo = 1 Then
         ' url2 = "https://mydataapidev.aade.gr/DCL/SendClient"
            url2 = "https://mydataapidev.aade.gr/DCL/RequestClients?DCLID=" + ARXID '100000000007663" ''''''100000000007739"
'                                                                         100000000834957
      Else
      
             ' url2 = "https://mydatapi.aade.gr/DCL/SendClient"
             url2 = "https://mydatapi.aade.gr/DCL/RequestClients?DCLID=100000000834950&maxdclid=100000000834960"  '[&maxdclid= + "&entityVatNumber =302009116 " '+ AFMPEL '"   '100000000007663" ''''''100000000007739"
        '                                                                                       100000000834957
      End If





        'initialize  100000000007664
114     Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
116     Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
118     XMLServer.setTimeouts 5000, 60000, 10000, 10000

        '130     XMLServer.Open "GET", url2 & "?mark=" + MARK2, False
    
        'force TLS 1.2
120     XMLServer.Option(9) = 2048
122     XMLServer.Option(6) = True
    
124     XMLServer.Open "GET", url2, False
        'XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
126     XMLServer.setRequestHeader "aade-user-id", P_USER '"glagakis2"
128     XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", P_KEY '  "555bc57c80634243958f62b629316aaa"

130   On Error Resume Next
        XMLServer.send 'docStock.XML 'TXML 'strData  'S????O ?? STRING TXML STO MYDATA  'XMLServer.send TXML '  ?? T??O ?? ?? S????O ??? ?????? ???? ???O strData
132     Debug.Print XMLServer.Status
        '  debugPrint XMLServer.responseText
134     v = XMLServer.responseText
        '--------------- ???S???? INVOICES -----------------------------------------------------
 
136     MsgBox v

138     TIMOL = Replace(TIMOL, "/", "-")

        Dim mfile As String:

140     mfile = "c:\" + F_TXTFILES + "\sendinv\OLAaytok" + "-" + ARXID + "-" + Format(Now, "yyyyddMMHHmm") + ".xml"
142     Debug.Print v
144     Open mfile For Output As #8
146     Print #8, v
148     Close #8
    
150     FileCopy mfile, "C:\" + F_TXTFILES + "\apantSendCANC.XML"
    
    DUM = Shell("c:\mercvb\NOTEPAd.exe " + mfile, vbNormalFocus)

        Dim MARK2 As String
        
       ' Dim ANS
       ' ANS = MsgBox("NA GINEI EISAGVGH", vbYesNo)
       ' If ANS = vbYes Then
           ' READXML v
        'End If
       
        
        
        ' -------------- À«÷»≈…”¡ ¡–¡Õ‘«”« ------------------https://mydataapidev.aade.gr/DCL/RequestClients?DCLID=100000000007664"
        '<?xml version="1.0" encoding="utf-8"?>
        '<RequestedDoc xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns="http://www.aade.gr/myDATA/dcr/v1.0">
        '  <entityVatNumber>028783755</entityVatNumber>
        '  <clientsDoc>
        '    <DigitalClient>
        '      <InitialClientData>
        '        <idDcl>100000000007739</idDcl>
        '        <clientServiceType>3</clientServiceType>
        '        <creationDateTime xsi:nil="true" />
        '        <entityVatNumber>028783755</entityVatNumber>
        '        <branch>0</branch>
        '        <useCase>
        '          <garage>
        '            <vehicleRegistrationNumber>RRR8787</vehicleRegistrationNumber>
        '            <vehicleCategory>AYTOK</vehicleCategory>
        '            <vehicleFactory>CITROEN</vehicleFactory>
        '          </garage>
        '        </useCase>
        '      </InitialClientData>
        '    </DigitalClient>
        '  </clientsDoc>
        '  <updateclientRequestsDoc />
        '  <clientcorrelationsRequestsDoc />
        '  <cancelClientRequestsDoc />
        '</RequestedDoc>

        '===================================== https://mydataapidev.aade.gr/DCL/RequestClients?DCLID=100000000007663"
        '<?xml version="1.0" encoding="utf-8"?>
        '<RequestedDoc xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns="http://www.aade.gr/myDATA/dcr/v1.0">
        '  <entityVatNumber>028783755</entityVatNumber>
        '  <clientsDoc>
        '    <DigitalClient>
        '      <InitialClientData>
        '        <idDcl>100000000007664</idDcl>
        '        <clientServiceType>3</clientServiceType>
        '        <creationDateTime xsi:nil="true" />
        '        <entityVatNumber>028783755</entityVatNumber>
        '        <branch>0</branch>
        '        <useCase>
        '          <garage>
        '            <vehicleRegistrationNumber>RMK5555</vehicleRegistrationNumber>
        '            <vehicleCategory>FORTHGAKI</vehicleCategory>
        '            <vehicleFactory>FIAT</vehicleFactory>
        '          </garage>
        '        </useCase>
        '      </InitialClientData>
        '    </DigitalClient>
        '    <DigitalClient>
        '      <InitialClientData>
        '        <idDcl>100000000007739</idDcl>
        '        <clientServiceType>3</clientServiceType>
        '        <creationDateTime xsi:nil="true" />
        '        <entityVatNumber>028783755</entityVatNumber>
        '        <branch>0</branch>
        '        <useCase>
        '          <garage>
        '            <vehicleRegistrationNumber>RRR8787</vehicleRegistrationNumber>
        '            <vehicleCategory>AYTOK</vehicleCategory>
        '            <vehicleFactory>CITROEN</vehicleFactory>
        '          </garage>
        '        </useCase>
        '      </InitialClientData>
        '    </DigitalClient>
        '  </clientsDoc>
        '  <updateclientRequestsDoc />
        '  <clientcorrelationsRequestsDoc />
        '  <cancelClientRequestsDoc />
        '</RequestedDoc>
    
        '<EhFooter>
        Exit Sub

cmdDOMDocumentUTF8_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ADOMERCNEW.Par7MyData.RELATE_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>






End Sub

Private Sub cmdSYSX_Click()
On Error Resume Next


   Dim cATIM, DIDCL, e As String
   Dim MARK As String
 
 
        TIMGrid.Col = 1
        cATIM = TIMGrid.Text  'Y000123 PARASTATIKO
      
      If Len(Trim$(cATIM)) = 0 Then
      
          MsgBox " ƒ≈Õ ’–¡—◊≈… –¡—¡”‘¡‘… œ"
          Exit Sub
          
      End If
   Dim c3 As String
   
   c3 = GGET_CVALUE("SELECT TOP 1 C3 FROM TIM WHERE ATIM='" + cATIM + "'")
       'c3 = TIMGrid.columns("C3")  DEN GINOTAN ONLINE
       
       If Len(c3) > 0 Then
          MsgBox "«ƒ« ”’”◊≈‘…”‘« ≈"
          Exit Sub
         
       End If
     
     
        TIMGrid.Col = 8
      MARK = TIMGrid.Text ' 4000001023334
      If Left(MARK, 3) = "ƒœ " Then
          MARK = mID(MARK, 4, 15)
      End If
      
      Dim ID As String
         
         ' CARGrid.Col = 3
        ID = CARGrid.columns("ID") 'CARGrid.Text  '153
         
         
         
         
         
       '  CARGrid.Col = 0
        DIDCL = CARGrid.columns("NEWDIGITALCLIENTID") ' CARGrid.Text  '100000000008565
        
      
        
       ' TIMGrid.Col = 9
    '  MARK = TIMGrid.Text
        
        
        loadpar72
vid = RELATECLIEND(F_MESS, MARK, DIDCL, 1, "", "", "")
 'vid = "100000000008565"
        
        Dim mPoso As String
        mPoso = Format(TIMGrid.columns(2), "####0.00")
        mPoso = Replace(mPoso, ",", ".")
        
 If Len(Trim$(vid)) > 10 Then
      Gdb.Execute "UPDATE TIM SET  C3='" + ID + "'  where ATIM='" + cATIM + "'"
        
    Gdb.Execute "UPDATE CARS SET  POSO=" + mPoso + ", SXETIKA='" + MARK + "'  where ID=" + str(ID)
    
    Gdb.Execute "INSERT INTO CARRELATIONS(FIMISSUEDATE,CORRELATEDDCLIDS,CORRELATEDID,MARK) VALUES(CONVERT(CHAR(10),GETDATE(),103),'" + DIDCL + "','" + vid + "','" + MARK + "')"
    TIMGrid.Refresh
    Else
    
    MsgBox "ƒ≈Õ ¡Õ‘…”‘œ…◊«»« ≈"
 End If
      'Gdb.Execute "UPDATE CARS SET   SXETIKA=ISNULL(SXETIKA,'')+'" + MARK + "'+' " + cATIM + "'+' " + str(ID) + " ' where ID=" + str(ID)
         '  Dim c As String

      '  PAR1.TDBGrid1.Col = 2
     '   c = Replace(PAR1.TDBGrid1.Text, ",", ".")
        
     ' cars.Refresh
      
        
        
        
End Sub

Private Sub cmdUpdateClient_Click()
 '<?xml version="1.0" encoding="utf-8"?>
        '<dcrudt:UpdateClientDoc xmlns:dcrudt="http://www.aade.gr/myDATA/dcrudt/v1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.aade.gr/myDATA/dcrudt/updateClientType-v1.0.xsd">
        '  <dcrudt:updateClient>
        '    <dcrudt:initialDclId>2</dcrudt:initialDclId>
        '
        '  </dcrudt:updateClient>
        '</dcrudt:UpdateClientDoc>

        Dim ANS    As String, TYPOS As String

        Dim ARKYKL As String

100     ANS = InputBox("2=PARKING 3=GARAGE", "")

102     ARKYKL = InputBox("·Ò.ÍıÍÎÔˆÔÒÈ·Ú", "")

104     If Val(ANS) = 2 Then
106         TYPOS = "dcrnew:parkingcarwash"
        Else
108         TYPOS = "dcrnew:garage"
        End If

        Dim DUM As String

110    ' DUM = fupdateclient("100000000006453", "2", "202502221033")

End Sub

Private Sub Command1_Click()
    MsgBox ("≈…”¡√Ÿ√« –≈À¡‘« –œ—‘œ ¡À… œ»œÕ«->¡—…»Ãœ” ‘¡Ã≈…¡ «” ‘’–œ” ”’Õ≈—√≈…œ’/–¡— …Õ√ ")

End Sub

Private Sub Command2_Click()
  
  READXML "AA"
  End Sub

Private Sub Command3_Click()
    '=======================================  update ·¸‰ÂÈÓÁ =============================
    Dim cATIM, DIDCL, e As String
    Dim MARK As String
    Dim ID   As String
    CARGrid.Col = 3
    ID = CARGrid.columns("ID")
    CARGrid.Col = 0
    DIDCL = CARGrid.columns("NEWDIGITALCLIENTID")
    
    Dim iInvoiceKind As Integer: iInvoiceKind = Val(Left$(InvoiceKind.Text, 2))
    
    If iInvoiceKind = 2 Then 'TIMOLOGIO
       If Len(afmonoma.Text) = 9 And Val(afmonoma) > 1000 Then
           'OK
       Else
          MsgBox "”’Ã–À«—Ÿ”‘≈ ‘œ ¡÷Ã." + Chr(13) + "ƒ≈Õ ≈√…Õ≈ «  ¡‘◊Ÿ—«”«"
          Exit Sub
       End If
       If afmonoma.Visible = False Then
          MsgBox "”’Ã–À«—Ÿ”‘≈ ‘œ ¡÷Ã." + Chr(13) + "ƒ≈Õ ≈√…Õ≈ «  ¡‘◊Ÿ—«”«"
          Exit Sub
       
       End If
       
    End If
    
    
    
    If f_parking_auto = 1 Then
        F_MESS = 0
    End If
    
    
    
    
    'SXETIKA
    Dim SX As String
    SX = CARGrid.columns("SXETIKA")
    Dim nc As Integer
    nc = GGET_NVALUE("SELECT COUNT(*) FROM CARRELATIONS WHERE CORRELATEDDCLIDS='" + DIDCL + "'")
    
    
    
    
    
    
    If nc = 0 Then ' Len(SX) = 0 Then
        Me.Caption = "ƒ≈Õ ”’”◊≈‘…”‘« ≈ Ã≈ Ã¡—  ‘œ œ◊«Ã¡"
        ' Exit Sub
    End If
    loadpar72
    Dim CC As String
    CC = providedServiceCategory.Text
    CC = Split(CC, ";")(1)
    
    
    
    Dim iprovidedServiceCategory As Integer: iprovidedServiceCategory = Val(Left(CC, 2))
    
    ' 3=synergeia
    Dim MT As String
    If F_ISCAR = 1 Then MT = "3" Else MT = "2"
    
    
    
      
    Dim AFM                      As String: AFM = ""
    If iInvoiceKind = 2 Then 'TIMOLOGIO
     
        AFM = afmonoma.Text 'InputBox("¡÷Ã –≈À¡‘«")
      
        vid = fupdateclientTIM(F_MESS, DIDCL, MT, Format(Now, "yyyy-MM-ddTHH:mm:ss"), iprovidedServiceCategory, iInvoiceKind, "false", AFM, "0.0")
    ElseIf iInvoiceKind = 3 Then 'FIM
        vid = fupdateclient(F_MESS, DIDCL, MT, Format(Now, "yyyy-MM-ddTHH:mm:ss"), iprovidedServiceCategory, iInvoiceKind, "0.0")
    ElseIf iInvoiceKind = 1 Then
    
        'AFM = InputBox("¡÷Ã –≈À¡‘«")
      
        vid = fupdateclientTIM(F_MESS, DIDCL, MT, Format(Now, "yyyy-MM-ddTHH:mm:ss"), iprovidedServiceCategory, iInvoiceKind, "false", "000000000", "0.0")
    ElseIf iInvoiceKind = 0 Then 'ƒŸ—≈¡Õ
        
        
        vid = fupdateclient(F_MESS, DIDCL, MT, Format(Now, "yyyy-MM-ddTHH:mm:ss"), iprovidedServiceCategory, iInvoiceKind, "0.0")
        
    
    End If
   
    If Len(vid) > 10 Then
        Gdb.Execute "UPDATE CARS SET AFM='" + AFM + "',PROVIDEDSERVICECATEGORY=" + str(iprovidedServiceCategory) + ",INVOICEKIND=" + str(iInvoiceKind) + " where ID=" + ID
        afmonoma.Text = ""
    Else
        '  MsgBox "ƒ≈Õ ≈ À≈…”≈ « ≈√√—¡÷«"
   
    End If
    
    If f_parking_auto = 1 Then
       'NO REFRESH ' F_MESS = 0
    Else
       cars.Refresh
    End If
    
     afmonoma.Text = ""
     afmonoma.Visible = False
     ListEpo.Visible = False
     LSEARCH.Visible = False
     search.Visible = False
     
    
    
    
    
End Sub

Private Sub Command4_Click()
On Error Resume Next


   Dim cATIM, DIDCL, e As String
   Dim MARK As String
 
 
       ' TIMGrid.Col = 1
       ' cATIM = TIMGrid.Text  'Y000123 PARASTATIKO
      
    '  If Len(Trim$(cATIM)) = 0 Then
      
        '  MsgBox " ƒ≈Õ ’–¡—◊≈… –¡—¡”‘¡‘… œ"
        '  Exit Sub
          
    '  End If
   Dim c3 As String
   
 '  c3 = GGET_CVALUE("SELECT TOP 1 C3 FROM TIM WHERE ATIM='" + cATIM + "'")
     
       
  '     If Len(c3) > 0 Then
   '       MsgBox "«ƒ« ”’”◊≈‘…”‘« ≈"
    '      Exit Sub
         
    '   End If
     
     
     '   TIMGrid.Col = 8
      
      MARK = InputBox("ƒŸ”≈ ‘œ Ã¡—  ‘œ’ ‘…ÃœÀœ√…œ’", "Ã¡— ;", "400") '   TIMGrid.Text ' 4000001023334
      
      If Left(MARK, 3) = "ƒœ " Then
          MARK = mID(MARK, 4, 15)
      End If
      
      Dim ID As String
         
         ' CARGrid.Col = 3
        ID = CARGrid.columns("ID") 'CARGrid.Text  '153
         
         
         
         
         
       '  CARGrid.Col = 0
        DIDCL = CARGrid.columns("NEWDIGITALCLIENTID") ' CARGrid.Text  '100000000008565
        
      
        
       ' TIMGrid.Col = 9
    '  MARK = TIMGrid.Text
        
        
        loadpar72
vid = RELATECLIEND(F_MESS, MARK, DIDCL, 1, "", "", "")
 'vid = "100000000008565"
        
        Dim mPoso As String
        'mPoso = Format(TIMGrid.columns(2), "####0.00")
        mPoso = InputBox("ƒŸ”≈ ‘œ –œ”œ", "–œ”œ;", "0")
        mPoso = Replace(mPoso, ",", ".")
        
 If Len(Trim$(vid)) > 10 Then
     ' Gdb.Execute "UPDATE TIM SET  C3='" + ID + "'  where ATIM='" + cATIM + "'"
        
    Gdb.Execute "UPDATE CARS SET  ENTRYCOMPLETION=1, POSO=" + mPoso + ", SXETIKA='" + MARK + "'  where ID=" + str(ID)
    
    Gdb.Execute "INSERT INTO CARRELATIONS(FIMISSUEDATE,CORRELATEDDCLIDS,CORRELATEDID,MARK) VALUES(CONVERT(CHAR(10),GETDATE(),103),'" + DIDCL + "','" + vid + "','" + MARK + "')"
    TIMGrid.Refresh
    Else
    
    MsgBox "ƒ≈Õ ¡Õ‘…”‘œ…◊«»« ≈"
 End If
      
      





























End Sub

Private Sub DIAGRAFI_Click()
 Dim s As Integer
 
   Dim ID As String
         
          'CARGrid.Col = 3
        ID = CARGrid.columns("ID")   ' CARGrid.Text  '153
         
         
         'CARGrid.Col = 0
        DIDCL = CARGrid.columns("NEWDIGITALCLIENTID") ' CARGrid.Text  '100000000008565
 
 
    
loadpar72
 
 
 
 Dim CANCID As String
 CANCID = killClient(3, DIDCL)
 
  Gdb.Execute "UPDATE CARS SET   ENTRYCOMPLETION=1, CANCELED=1,CANCELEDID='" + CANCID + "' where ID=" + ID
 
  cars.Refresh
 
 
End Sub








Private Sub Form_Load()
  LOADFORM
  End Sub

Private Sub InvoiceKind_Click()
If Val(InvoiceKind.Text) = 0 And F_ISCAR = 1 Then 'MHDENIKO
     providedServiceCategory.Text = providedServiceCategory.List(3)
End If

If Val(InvoiceKind.Text) = 2 Then 'TIMOLOGIO
     
     afmonoma.Visible = True
     ListEpo.Visible = True
     LSEARCH.Visible = True
     search.Visible = True
     sxolia2.Text = ""
     sxolia2.Visible = False
     
 Else
 
     afmonoma.Visible = False
     ListEpo.Visible = False
     LSEARCH.Visible = False
     search.Visible = False
     sxolia2.Visible = True
     
End If




End Sub

Private Sub InvoiceKind_LostFocus()

'If Val(InvoiceKind.Text) = 0 Then 'MHDENIKO
'
'     providedServiceCategory.Text = providedServiceCategory.List(3)
'
'
'
'
'End If





End Sub

Private Sub lblVCAT_Click()
     UPDATE_PINAKES 50, lblVCAT, VCAT
End Sub

Private Sub lblVFACT_Click()
     UPDATE_PINAKES 51, lblVFACT, VFACT
End Sub


Private Sub ListEpo_Click()

  'If afmonoma.index Then
  
 ' End If
Dim c As String
c = ListEpo.List(ListEpo.ListIndex)
afmonoma.Text = Split(c, ";")(1)


End Sub

Private Sub METAXRONOLOGHMENO_Click()


   ' DIABAZO TIS METABLHTES APO TO GRID-------------------------------------------



  Dim cATIM, DIDCL, e As String
   Dim MARK As String
      Dim ID As String
          CARGrid.Col = 3
         ID = CARGrid.columns("ID")
         CARGrid.Col = 0
          DIDCL = CARGrid.columns("NEWDIGITALCLIENTID")
          
          Dim MAFM As String
          MAFM = CARGrid.columns("AFM")
          
          
          Dim mINVOICEKIND As String
          mINVOICEKIND = CARGrid.columns("INVOICEKIND")
          
           Dim mprovidedServiceCategory As String
          mprovidedServiceCategory = CARGrid.columns("providedServiceCategory")
           Dim iprovidedServiceCategory As Integer: iprovidedServiceCategory = Val(mprovidedServiceCategory)
      Dim iInvoiceKind As Integer: iInvoiceKind = Val(mINVOICEKIND)
       'SXETIKA
       
       
      '----------------- ¬À≈–Ÿ ¡Õ ≈◊Ÿ  ¡Õ≈… ”’”◊≈‘…”≈…”. AN OXI –—œ◊Ÿ—Ÿ  ¡ÕœÕ… ¡  ¡… ¬¡∆Ÿ ENTRYCOMPLETION= 2 ------------------------------------------------
      

       Dim SX As String
     SX = CARGrid.columns("SXETIKA")
     Dim nc As Integer
      nc = GGET_NVALUE("SELECT COUNT(*) FROM CARRELATIONS WHERE CORRELATEDDCLIDS='" + DIDCL + "'")
        If nc = 0 Then ' Len(SX) = 0 Then
          Me.Caption = "ƒ≈Õ ”’”◊≈‘…”‘« ≈ Ã≈ Ã¡—  ‘œ œ◊«Ã¡"
          MsgBox "ƒ≈Õ ”’”◊≈‘…”‘« ≈ Ã≈ Ã¡—  ‘œ œ◊«Ã¡. »¡ ”’”◊≈‘…”‘≈… ¡—√œ‘≈—¡ ≈Ÿ” ‘…” 15 ≈–œÃ≈Õœ’ Ã«Õ¡"
         ' Exit Sub
         Else
           MsgBox "«ƒ« ”’”◊‘≈‘…”‘« ≈. –¡‘«”‘≈ œÀœ À«—Ÿ”«"
           Exit Sub
        End If
      loadpar72
     
      ' 3=synergeia
      
      
      
'--------------------------  OLOKLHRVNV THN ≈–…” ≈’« -----------------------------------------------

      
        Dim mPoso As String
     ' MPOSO = InputBox("ƒŸ”≈ ‘œ –œ”œ", "”≈ ≈’—Ÿ ¡–œƒ≈…Œ«/–œ”")
     ' MPOSO = Replace$(MPOSO, ",", ".")
      mPoso = "0.0"
      
      
      Dim MT As String
    If F_ISCAR = 1 Then MT = "3" Else MT = "2"
      
      
   If iInvoiceKind = 2 Then 'TIMOLOGIO
   
     Dim MPOSO2 As String: MPOSO2 = CARGrid.columns("POSO")
      vid = fupdateclientTIM(F_MESS, DIDCL, MT, Format(Now, "yyyy-MM-ddTHH:mm:ss"), iprovidedServiceCategory, iInvoiceKind, "true", MAFM, MPOSO2)
   ElseIf iInvoiceKind = 3 Then 'FIM
     
     
     'AYTO EJYPHRETEITE APO TO fupdateclient
     ' MPOSO = InputBox("ƒŸ”≈ ‘œ –œ”œ", "”≈ ≈’—Ÿ ¡–œƒ≈…Œ«/–œ”")
     ' MPOSO = Replace$(MPOSO, ",", ".")
     ' vid = fupdateclient(DIDCL, "3", Format(Now, "yyyy-MM-ddTHH:mm:ss"), iprovidedServiceCategory, iInvoiceKind, MPOSO)
   
   
   
   
  ElseIf iInvoiceKind = 1 Then 'FIM
      mPoso = InputBox("ƒŸ”≈ ‘œ –œ”œ", "”≈ ≈’—Ÿ ¡–œƒ≈…Œ«/–œ”")
      mPoso = Replace$(mPoso, ",", ".")
      vid = fupdateclient(F_MESS, DIDCL, MT, Format(Now, "yyyy-MM-ddTHH:mm:ss"), iprovidedServiceCategory, iInvoiceKind, mPoso)
 
   
   
   
   End If
   
   
   
   
  If Len(vid) > 10 Then
              Gdb.Execute "UPDATE CARS SET ENTRYCOMPLETION=2,PROVIDEDSERVICECATEGORY=" + str(iprovidedServiceCategory) + ",INVOICEKIND=" + str(iInvoiceKind) + " where ID=" + ID
   Else
              MsgBox "ƒ≈Õ ≈ À≈…”≈ « ≈√√—¡÷«"
   
   End If
      cars.Refresh





' GIA NA JEMPLOKAREI TO GRID
  cURRTIM.Caption = ""
  












End Sub

Private Sub providedServiceCategory_LostFocus()


If F_ISCAR = 1 And Val(Left(providedServiceCategory, 2)) = 4 Then
   InvoiceKind.Text = InvoiceKind.List(0)
   InvoiceKind.Enabled = False
 Else
   'InvoiceKind.Text = InvoiceKind.List(3)
   InvoiceKind.Enabled = True

End If

  



End Sub

Private Sub search_Click()
On Error Resume Next

ListEpo.Clear
'If Val(afmonoma.Text) = 0 Then

   Dim R As New ADODB.Recordset
   R.Open "select LEFT(EPO,25)+' '+LEFT(ISNULL(EPA,'           '),12)+';'+ISNULL(AFM,'') AS LISTA FROM PEL WHERE EIDOS='e' and (EPO LIKE '%" + afmonoma.Text + "%'  or AFM  LIKE '%" + Trim(afmonoma.Text) + "%' ) ", Gdb, adOpenDynamic, adLockOptimistic
   Do While Not R.EOF
       ListEpo.AddItem R(0)
     
   
   
     R.MoveNext
   Loop
   R.Close
   
  ' End If
   
   
   
   
   
End Sub


Private Sub SEEKREMOTHTA_Click()
 cars.ConnectionString = gConnect
        
        Dim MTYPOS As Integer
        If F_ISCAR = 1 Then MTYPOS = 52 Else MTYPOS = 53
        

        cars.RecordSource = "SELECT VEHICLEREGISTRATIONNUMBER AS [¡—. ’ À] ,VEHICLECATEGORY AS [KATH√],VEHICLEFACTORY AS [≈—√œ”‘],PINAKES.PERIGRAFH ," _
        & "(CASE INVOICEKIND WHEN 1 THEN '¡À–' WHEN 2  THEN '‘…Ã' WHEN 3  THEN '÷«Ã' ELSE ' ' END) AS [–¡—¡”‘], " _
        & "ISNULL(ENTRYCOMPLETION,0) AS [œ /1 ≈  /0],COMMENTS AS [”◊œÀ…¡],* FROM CARS LEFT JOIN PINAKES ON PINAKES.TYPOS=" + str(MTYPOS) + " AND PROVIDEDSERVICECATEGORY=TIMH " _
        & " WHERE ISNULL(CANCELED,0)=0 AND ISNULL(ENTRYCOMPLETION,0)=2 ORDER BY CARS.ID DESC"
        'TIMGrid.columns(0).DataWidth = 3600
        ' = 3600
        cars.Refresh
        CARGrid.columns(1).Width = 3600
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

130     AdoCarsOLD.Recordset.Filter = getFilter()

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
         "in ADOMERCNEW.bohu6.TDBGrid_FilterChange " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.TDBGrid_FilterChange " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub TIMGrid_Click()
  
    TIMGrid.Col = 1
        cATIM = TIMGrid.Text  'Y000123 PARASTATIKO
  
     cURRTIM.Caption = cATIM
  
End Sub

Private Sub TDBGrid_HeadClick(ByVal ColIndex As Integer)

        '--------------------------------------------------------------------------------------=
        '<EhHeader>
        On Error GoTo TDBGrid_HeadClick_Err

        '</EhHeader>
        Dim sumes(100) As String

100     For k = 0 To Adodc2.Recordset.Fields.Count - 1
110         sumes(k) = TDBGrid.Splits(0).columns(k).FooterText    '  = Format(SUMES(k), "######0.00")
        Next

120     If AdoCarsOLD.Recordset.Sort = "[" & TDBGrid.columns(ColIndex).DataField & "] asc" Then   ' strSort
130         AdoCarsOLD.Recordset.Sort = "[" & TDBGrid.columns(ColIndex).DataField & "] desc"    ' strSort
        Else
140         AdoCarsOLD.Recordset.Sort = "[" & TDBGrid.columns(ColIndex).DataField & "] asc"    ' strSort
        End If

150     For k = 0 To Adodc2.Recordset.Fields.Count - 1
160         TDBGrid.Splits(0).columns(k).FooterText = sumes(k)   '  = Format(SUMES(k), "######0.00")
        Next

        '<EhFooter>
        Exit Sub

TDBGrid_HeadClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.TDBGrid_HeadClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.TDBGrid_HeadClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub



Private Sub UPDATE_PEL(ByVal PEDIO As String, ByVal TIMH As String, ByVal ARITMOS As Integer, ByVal EIDOS As String, ByVal kod As String)
   
   If ARITMOS = 1 Then
            Gdb.Execute "UPDATE CARS SET " + PEDIO + "=" + TIMH + "   WHERE EIDOS='" + EIDOS + "' AND KOD='" + kod + "'"
   Else
            Gdb.Execute "UPDATE CARS SET " + PEDIO + "='" + TIMH + "' WHERE EIDOS='" + EIDOS + "' AND KOD='" + kod + "'"

   End If



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

150             If AdoCarsOLD.Recordset(Col.ColIndex).Type = 5 Then    '
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

 getFilter = tmp

        '<EhFooter>
        Exit Function

getFilter_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.getFilter " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PSIFIAKO.getFilter " & "at line " & Erl

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

130     AdoCarsOLD.Recordset.Filter = adFilterNone

        '<EhFooter>
        Exit Sub

cmdClearFilter_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu6.cmdClearFilter_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.PSIFIAKO.cmdClearFilter_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

