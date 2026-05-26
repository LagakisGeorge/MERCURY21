VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form APOT6 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1:  True OLEDBGrid Pro 8.0 - Tutorial 8"
   ClientHeight    =   8550
   ClientLeft      =   1275
   ClientTop       =   1620
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   10995
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000010&
      Caption         =   "≈ÓÔ‰ÔÚ"
      Height          =   735
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton YPOL 
      Caption         =   "’–œÀœ√…”Ãœ”"
      Height          =   375
      Left            =   7080
      TabIndex        =   23
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox DORO 
      Height          =   495
      Left            =   7680
      TabIndex        =   22
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox eosk 
      Height          =   495
      Left            =   3000
      TabIndex        =   16
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox apok 
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox notsql 
      Height          =   360
      Left            =   4515
      TabIndex        =   11
      Top             =   15810
      Visible         =   0   'False
      Width           =   3090
   End
   Begin VB.CommandButton Command2 
      Caption         =   "’ÔÎÔ„ÈÛÏ¸Ú"
      Height          =   375
      Left            =   7740
      TabIndex        =   10
      Top             =   15435
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.CommandButton Command1000 
      Caption         =   "¡ÎÎ·„ﬁ"
      Height          =   465
      Index           =   0
      Left            =   4560
      TabIndex        =   9
      Top             =   16980
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.TextBox ypooik 
      Height          =   390
      Left            =   5940
      TabIndex        =   6
      Top             =   16500
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox oik 
      Height          =   345
      Left            =   4560
      TabIndex        =   5
      Top             =   16500
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox sql 
      Height          =   360
      Left            =   4515
      TabIndex        =   3
      Top             =   15435
      Visible         =   0   'False
      Width           =   3090
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   6825
      Top             =   12045
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      DataSourceName  =   "MERCCOMPAQ"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "YPOKATHG"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   480
      Top             =   8880
      Visible         =   0   'False
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
      DataSourceName  =   "MERCCOMPAQ"
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
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDown1 
      Bindings        =   "APOT6.frx":0000
      Height          =   1200
      Left            =   5985
      TabIndex        =   1
      Top             =   10845
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2117
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "AYJON"
      Columns(0).DataField=   "AYJON"
      Columns(0).DataWidth=   23
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "PERIGRAFH"
      Columns(1).DataField=   "PERIGRAFH"
      Columns(1).DataWidth=   20
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1164"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1085"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._AlignLeft=0"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   1
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   2
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   0   'False
      ListField       =   "PERIGRAFH"
      DataField       =   "AYJON"
      IntegralHeight  =   -1  'True
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   12632256
      ValueTranslate  =   0   'False
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=161"
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
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(23)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=11"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=11"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=12"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=15"
      _StyleDefs(44)  =   "Named:id=29:Normal"
      _StyleDefs(45)  =   ":id=29,.parent=0"
      _StyleDefs(46)  =   "Named:id=30:Heading"
      _StyleDefs(47)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=30,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=31:Footing"
      _StyleDefs(50)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=32:Selected"
      _StyleDefs(52)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=33:Caption"
      _StyleDefs(54)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(55)  =   "Named:id=34:HighlightRow"
      _StyleDefs(56)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(57)  =   "Named:id=35:EvenRow"
      _StyleDefs(58)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=36:OddRow"
      _StyleDefs(60)  =   ":id=36,.parent=29"
      _StyleDefs(61)  =   "Named:id=39:RecordSelector"
      _StyleDefs(62)  =   ":id=39,.parent=30"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=29"
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Bindings        =   "APOT6.frx":0015
      Height          =   4965
      Left            =   480
      TabIndex        =   0
      Top             =   9360
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8758
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "AEG"
      Columns(0).DataField=   "AEG"
      Columns(0).DataWidth=   23
      Columns(0).DropDown=   "TDBDropDown2"
      Columns(0).DropDown.vt=   8
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "KODLOG"
      Columns(1).DataField=   "KODLOG"
      Columns(1).DataWidth=   3
      Columns(1).DropDown=   "TDBDropDown1"
      Columns(1).DropDown.vt=   8
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ONO"
      Columns(2).DataField=   "ONO"
      Columns(2).DataWidth=   35
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "XTI"
      Columns(3).DataField=   "XTI"
      Columns(3).DataWidth=   23
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "LTI"
      Columns(4).DataField=   "LTI"
      Columns(4).DataWidth=   23
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "LTI2"
      Columns(5).DataField=   "LTI2"
      Columns(5).DataWidth=   23
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "LTI3"
      Columns(6).DataField=   "LTI3"
      Columns(6).DataWidth=   23
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "LTI4"
      Columns(7).DataField=   "LTI4"
      Columns(7).DataWidth=   23
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "LTI5"
      Columns(8).DataField=   "LTI5"
      Columns(8).DataWidth=   23
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "APOS"
      Columns(9).DataField=   "APOS"
      Columns(9).DataWidth=   23
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "EISSYN"
      Columns(10).DataField=   "EISSYN"
      Columns(10).DataWidth=   23
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "EXSYN"
      Columns(11).DataField=   "EXSYN"
      Columns(11).DataWidth=   23
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "POS"
      Columns(12).DataField=   "POS"
      Columns(12).DataWidth=   23
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "EISMHN"
      Columns(13).DataField=   "EISMHN"
      Columns(13).DataWidth=   23
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "EXMHN"
      Columns(14).DataField=   "EXMHN"
      Columns(14).DataWidth=   23
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "DESMIA"
      Columns(15).DataField=   "DESMIA"
      Columns(15).DataWidth=   23
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "APOS01"
      Columns(16).DataField=   "APOS01"
      Columns(16).DataWidth=   23
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "APOS02"
      Columns(17).DataField=   "APOS02"
      Columns(17).DataWidth=   23
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "APOS03"
      Columns(18).DataField=   "APOS03"
      Columns(18).DataWidth=   23
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "APOS04"
      Columns(19).DataField=   "APOS04"
      Columns(19).DataWidth=   23
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "POS01"
      Columns(20).DataField=   "POS01"
      Columns(20).DataWidth=   23
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "POS02"
      Columns(21).DataField=   "POS02"
      Columns(21).DataWidth=   23
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "POS03"
      Columns(22).DataField=   "POS03"
      Columns(22).DataWidth=   23
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "POS04"
      Columns(23).DataField=   "POS04"
      Columns(23).DataWidth=   23
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "MON"
      Columns(24).DataField=   "MON"
      Columns(24).DataWidth=   3
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "SPA"
      Columns(25).DataField=   "SPA"
      Columns(25).DataWidth=   23
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).Caption=   "FPA"
      Columns(26).DataField=   "FPA"
      Columns(26).DataWidth=   23
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).Caption=   "PAR"
      Columns(27).DataField=   "PAR"
      Columns(27).DataWidth=   23
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).Caption=   "HPAR"
      Columns(28).DataField=   "HPAR"
      Columns(28).DataWidth=   19
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).Caption=   "AEG"
      Columns(29).DataField=   "AEG"
      Columns(29).DataWidth=   23
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(30)._VlistStyle=   0
      Columns(30)._MaxComboItems=   5
      Columns(30).Caption=   "UES"
      Columns(30).DataField=   "UES"
      Columns(30).DataWidth=   4
      Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(31)._VlistStyle=   0
      Columns(31)._MaxComboItems=   5
      Columns(31).Caption=   "M01"
      Columns(31).DataField=   "M01"
      Columns(31).DataWidth=   23
      Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(32)._VlistStyle=   0
      Columns(32)._MaxComboItems=   5
      Columns(32).Caption=   "M02"
      Columns(32).DataField=   "M02"
      Columns(32).DataWidth=   23
      Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(33)._VlistStyle=   0
      Columns(33)._MaxComboItems=   5
      Columns(33).Caption=   "M03"
      Columns(33).DataField=   "M03"
      Columns(33).DataWidth=   23
      Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(34)._VlistStyle=   0
      Columns(34)._MaxComboItems=   5
      Columns(34).Caption=   "M04"
      Columns(34).DataField=   "M04"
      Columns(34).DataWidth=   23
      Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(35)._VlistStyle=   0
      Columns(35)._MaxComboItems=   5
      Columns(35).Caption=   "M05"
      Columns(35).DataField=   "M05"
      Columns(35).DataWidth=   23
      Columns(35)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(36)._VlistStyle=   0
      Columns(36)._MaxComboItems=   5
      Columns(36).Caption=   "M06"
      Columns(36).DataField=   "M06"
      Columns(36).DataWidth=   23
      Columns(36)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(37)._VlistStyle=   0
      Columns(37)._MaxComboItems=   5
      Columns(37).Caption=   "M07"
      Columns(37).DataField=   "M07"
      Columns(37).DataWidth=   23
      Columns(37)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(38)._VlistStyle=   0
      Columns(38)._MaxComboItems=   5
      Columns(38).Caption=   "M08"
      Columns(38).DataField=   "M08"
      Columns(38).DataWidth=   23
      Columns(38)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(39)._VlistStyle=   0
      Columns(39)._MaxComboItems=   5
      Columns(39).Caption=   "M09"
      Columns(39).DataField=   "M09"
      Columns(39).DataWidth=   23
      Columns(39)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(40)._VlistStyle=   0
      Columns(40)._MaxComboItems=   5
      Columns(40).Caption=   "M10"
      Columns(40).DataField=   "M10"
      Columns(40).DataWidth=   23
      Columns(40)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(41)._VlistStyle=   0
      Columns(41)._MaxComboItems=   5
      Columns(41).Caption=   "M11"
      Columns(41).DataField=   "M11"
      Columns(41).DataWidth=   23
      Columns(41)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(42)._VlistStyle=   0
      Columns(42)._MaxComboItems=   5
      Columns(42).Caption=   "M12"
      Columns(42).DataField=   "M12"
      Columns(42).DataWidth=   23
      Columns(42)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(43)._VlistStyle=   0
      Columns(43)._MaxComboItems=   5
      Columns(43).Caption=   "G01"
      Columns(43).DataField=   "G01"
      Columns(43).DataWidth=   23
      Columns(43)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(44)._VlistStyle=   0
      Columns(44)._MaxComboItems=   5
      Columns(44).Caption=   "G02"
      Columns(44).DataField=   "G02"
      Columns(44).DataWidth=   23
      Columns(44)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(45)._VlistStyle=   0
      Columns(45)._MaxComboItems=   5
      Columns(45).Caption=   "G03"
      Columns(45).DataField=   "G03"
      Columns(45).DataWidth=   23
      Columns(45)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(46)._VlistStyle=   0
      Columns(46)._MaxComboItems=   5
      Columns(46).Caption=   "G04"
      Columns(46).DataField=   "G04"
      Columns(46).DataWidth=   23
      Columns(46)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(47)._VlistStyle=   0
      Columns(47)._MaxComboItems=   5
      Columns(47).Caption=   "G05"
      Columns(47).DataField=   "G05"
      Columns(47).DataWidth=   23
      Columns(47)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(48)._VlistStyle=   0
      Columns(48)._MaxComboItems=   5
      Columns(48).Caption=   "G06"
      Columns(48).DataField=   "G06"
      Columns(48).DataWidth=   23
      Columns(48)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(49)._VlistStyle=   0
      Columns(49)._MaxComboItems=   5
      Columns(49).Caption=   "G07"
      Columns(49).DataField=   "G07"
      Columns(49).DataWidth=   23
      Columns(49)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(50)._VlistStyle=   0
      Columns(50)._MaxComboItems=   5
      Columns(50).Caption=   "G08"
      Columns(50).DataField=   "G08"
      Columns(50).DataWidth=   23
      Columns(50)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(51)._VlistStyle=   0
      Columns(51)._MaxComboItems=   5
      Columns(51).Caption=   "G09"
      Columns(51).DataField=   "G09"
      Columns(51).DataWidth=   23
      Columns(51)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(52)._VlistStyle=   0
      Columns(52)._MaxComboItems=   5
      Columns(52).Caption=   "G10"
      Columns(52).DataField=   "G10"
      Columns(52).DataWidth=   23
      Columns(52)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(53)._VlistStyle=   0
      Columns(53)._MaxComboItems=   5
      Columns(53).Caption=   "G11"
      Columns(53).DataField=   "G11"
      Columns(53).DataWidth=   23
      Columns(53)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(54)._VlistStyle=   0
      Columns(54)._MaxComboItems=   5
      Columns(54).Caption=   "G12"
      Columns(54).DataField=   "G12"
      Columns(54).DataWidth=   23
      Columns(54)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(55)._VlistStyle=   0
      Columns(55)._MaxComboItems=   5
      Columns(55).Caption=   "POS_KERD2"
      Columns(55).DataField=   "POS_KERD2"
      Columns(55).DataWidth=   23
      Columns(55)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(56)._VlistStyle=   0
      Columns(56)._MaxComboItems=   5
      Columns(56).Caption=   "POS_KERD"
      Columns(56).DataField=   "POS_KERD"
      Columns(56).DataWidth=   23
      Columns(56)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(57)._VlistStyle=   0
      Columns(57)._MaxComboItems=   5
      Columns(57).Caption=   "POS_EKPT"
      Columns(57).DataField=   "POS_EKPT"
      Columns(57).DataWidth=   23
      Columns(57)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(58)._VlistStyle=   0
      Columns(58)._MaxComboItems=   5
      Columns(58).Caption=   "KODFIAL"
      Columns(58).DataField=   "KODFIAL"
      Columns(58).DataWidth=   23
      Columns(58)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(59)._VlistStyle=   0
      Columns(59)._MaxComboItems=   5
      Columns(59).Caption=   "KODLOG"
      Columns(59).DataField=   "KODLOG"
      Columns(59).DataWidth=   3
      Columns(59)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(60)._VlistStyle=   0
      Columns(60)._MaxComboItems=   5
      Columns(60).Caption=   "AJIAPO"
      Columns(60).DataField=   "AJIAPO"
      Columns(60).DataWidth=   23
      Columns(60)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(61)._VlistStyle=   0
      Columns(61)._MaxComboItems=   5
      Columns(61).Caption=   "AJIPOL"
      Columns(61).DataField=   "AJIPOL"
      Columns(61).DataWidth=   23
      Columns(61)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(62)._VlistStyle=   0
      Columns(62)._MaxComboItems=   5
      Columns(62).Caption=   "AJIAGO"
      Columns(62).DataField=   "AJIAGO"
      Columns(62).DataWidth=   23
      Columns(62)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(63)._VlistStyle=   0
      Columns(63)._MaxComboItems=   5
      Columns(63).Caption=   "MESXTI"
      Columns(63).DataField=   "MESXTI"
      Columns(63).DataWidth=   23
      Columns(63)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(64)._VlistStyle=   0
      Columns(64)._MaxComboItems=   5
      Columns(64).Caption=   "LIT_SYNOD"
      Columns(64).DataField=   "LIT_SYNOD"
      Columns(64).DataWidth=   23
      Columns(64)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(65)._VlistStyle=   0
      Columns(65)._MaxComboItems=   5
      Columns(65).Caption=   "LTI_SYNOD"
      Columns(65).DataField=   "LTI_SYNOD"
      Columns(65).DataWidth=   23
      Columns(65)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(66)._VlistStyle=   0
      Columns(66)._MaxComboItems=   5
      Columns(66).Caption=   "SYSKEYASIA"
      Columns(66).DataField=   "SYSKEYASIA"
      Columns(66).DataWidth=   23
      Columns(66)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(67)._VlistStyle=   0
      Columns(67)._MaxComboItems=   5
      Columns(67).Caption=   "PROMHU"
      Columns(67).DataField=   "PROMHU"
      Columns(67).DataWidth=   23
      Columns(67)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(68)._VlistStyle=   0
      Columns(68)._MaxComboItems=   5
      Columns(68).Caption=   "KODLOGAG"
      Columns(68).DataField=   "KODLOGAG"
      Columns(68).DataWidth=   12
      Columns(68)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(69)._VlistStyle=   0
      Columns(69)._MaxComboItems=   5
      Columns(69).Caption=   "PROM"
      Columns(69).DataField=   "PROM"
      Columns(69).DataWidth=   30
      Columns(69)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(70)._VlistStyle=   0
      Columns(70)._MaxComboItems=   5
      Columns(70).Caption=   "SYNAGO"
      Columns(70).DataField=   "SYNAGO"
      Columns(70).DataWidth=   23
      Columns(70)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(71)._VlistStyle=   0
      Columns(71)._MaxComboItems=   5
      Columns(71).Caption=   "SYNPOL"
      Columns(71).DataField=   "SYNPOL"
      Columns(71).DataWidth=   23
      Columns(71)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(72)._VlistStyle=   0
      Columns(72)._MaxComboItems=   5
      Columns(72).Caption=   "EISSYN0"
      Columns(72).DataField=   "EISSYN0"
      Columns(72).DataWidth=   23
      Columns(72)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(73)._VlistStyle=   0
      Columns(73)._MaxComboItems=   5
      Columns(73).Caption=   "EXSYN0"
      Columns(73).DataField=   "EXSYN0"
      Columns(73).DataWidth=   23
      Columns(73)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(74)._VlistStyle=   0
      Columns(74)._MaxComboItems=   5
      Columns(74).Caption=   "SYNAGO0"
      Columns(74).DataField=   "SYNAGO0"
      Columns(74).DataWidth=   23
      Columns(74)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(75)._VlistStyle=   0
      Columns(75)._MaxComboItems=   5
      Columns(75).Caption=   "SYNPOL0"
      Columns(75).DataField=   "SYNPOL0"
      Columns(75).DataWidth=   23
      Columns(75)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(76)._VlistStyle=   0
      Columns(76)._MaxComboItems=   5
      Columns(76).Caption=   "EISMHN0"
      Columns(76).DataField=   "EISMHN0"
      Columns(76).DataWidth=   23
      Columns(76)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(77)._VlistStyle=   0
      Columns(77)._MaxComboItems=   5
      Columns(77).Caption=   "EXMHN0"
      Columns(77).DataField=   "EXMHN0"
      Columns(77).DataWidth=   23
      Columns(77)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(78)._VlistStyle=   0
      Columns(78)._MaxComboItems=   5
      Columns(78).Caption=   "POS_KERD3"
      Columns(78).DataField=   "POS_KERD3"
      Columns(78).DataWidth=   23
      Columns(78)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(79)._VlistStyle=   0
      Columns(79)._MaxComboItems=   5
      Columns(79).Caption=   "MEMO"
      Columns(79).DataField=   "MEMO"
      Columns(79).DataWidth=   65535
      Columns(79)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(80)._VlistStyle=   0
      Columns(80)._MaxComboItems=   5
      Columns(80).Caption=   "PONTOI"
      Columns(80).DataField=   "PONTOI"
      Columns(80).DataWidth=   11
      Columns(80)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   81
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=81"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=979"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=900"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._AlignLeft=0"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1058"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=979"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=7514"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=7435"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(3)._AlignLeft=0"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(28)=   "Column(4)._AlignLeft=0"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(32)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(34)=   "Column(5)._AlignLeft=0"
      Splits(0)._ColumnProps(35)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(38)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(40)=   "Column(6)._AlignLeft=0"
      Splits(0)._ColumnProps(41)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(42)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(44)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(45)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(46)=   "Column(7)._AlignLeft=0"
      Splits(0)._ColumnProps(47)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(50)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(51)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(52)=   "Column(8)._AlignLeft=0"
      Splits(0)._ColumnProps(53)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(54)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(56)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(57)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(58)=   "Column(9)._AlignLeft=0"
      Splits(0)._ColumnProps(59)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(60)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(62)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(63)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(64)=   "Column(10)._AlignLeft=0"
      Splits(0)._ColumnProps(65)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(66)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(68)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(69)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(70)=   "Column(11)._AlignLeft=0"
      Splits(0)._ColumnProps(71)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(72)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(74)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(75)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(76)=   "Column(12)._AlignLeft=0"
      Splits(0)._ColumnProps(77)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(78)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(79)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(80)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(81)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(82)=   "Column(13)._AlignLeft=0"
      Splits(0)._ColumnProps(83)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(84)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(85)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(86)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(87)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(88)=   "Column(14)._AlignLeft=0"
      Splits(0)._ColumnProps(89)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(90)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(92)=   "Column(15)._EditAlways=0"
      Splits(0)._ColumnProps(93)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(94)=   "Column(15)._AlignLeft=0"
      Splits(0)._ColumnProps(95)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(96)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(97)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(98)=   "Column(16)._EditAlways=0"
      Splits(0)._ColumnProps(99)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(100)=   "Column(16)._AlignLeft=0"
      Splits(0)._ColumnProps(101)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(102)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(103)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(104)=   "Column(17)._EditAlways=0"
      Splits(0)._ColumnProps(105)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(106)=   "Column(17)._AlignLeft=0"
      Splits(0)._ColumnProps(107)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(108)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(109)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(110)=   "Column(18)._EditAlways=0"
      Splits(0)._ColumnProps(111)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(112)=   "Column(18)._AlignLeft=0"
      Splits(0)._ColumnProps(113)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(114)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(116)=   "Column(19)._EditAlways=0"
      Splits(0)._ColumnProps(117)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(118)=   "Column(19)._AlignLeft=0"
      Splits(0)._ColumnProps(119)=   "Column(20).Width=2725"
      Splits(0)._ColumnProps(120)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(121)=   "Column(20)._WidthInPix=2646"
      Splits(0)._ColumnProps(122)=   "Column(20)._EditAlways=0"
      Splits(0)._ColumnProps(123)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(124)=   "Column(20)._AlignLeft=0"
      Splits(0)._ColumnProps(125)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(126)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(127)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(128)=   "Column(21)._EditAlways=0"
      Splits(0)._ColumnProps(129)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(130)=   "Column(21)._AlignLeft=0"
      Splits(0)._ColumnProps(131)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(132)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(133)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(134)=   "Column(22)._EditAlways=0"
      Splits(0)._ColumnProps(135)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(136)=   "Column(22)._AlignLeft=0"
      Splits(0)._ColumnProps(137)=   "Column(23).Width=2725"
      Splits(0)._ColumnProps(138)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(139)=   "Column(23)._WidthInPix=2646"
      Splits(0)._ColumnProps(140)=   "Column(23)._EditAlways=0"
      Splits(0)._ColumnProps(141)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(142)=   "Column(23)._AlignLeft=0"
      Splits(0)._ColumnProps(143)=   "Column(24).Width=873"
      Splits(0)._ColumnProps(144)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(145)=   "Column(24)._WidthInPix=794"
      Splits(0)._ColumnProps(146)=   "Column(24)._EditAlways=0"
      Splits(0)._ColumnProps(147)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(148)=   "Column(25).Width=2725"
      Splits(0)._ColumnProps(149)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(150)=   "Column(25)._WidthInPix=2646"
      Splits(0)._ColumnProps(151)=   "Column(25)._EditAlways=0"
      Splits(0)._ColumnProps(152)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(153)=   "Column(25)._AlignLeft=0"
      Splits(0)._ColumnProps(154)=   "Column(26).Width=2725"
      Splits(0)._ColumnProps(155)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(156)=   "Column(26)._WidthInPix=2646"
      Splits(0)._ColumnProps(157)=   "Column(26)._EditAlways=0"
      Splits(0)._ColumnProps(158)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(159)=   "Column(26)._AlignLeft=0"
      Splits(0)._ColumnProps(160)=   "Column(27).Width=2725"
      Splits(0)._ColumnProps(161)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(162)=   "Column(27)._WidthInPix=2646"
      Splits(0)._ColumnProps(163)=   "Column(27)._EditAlways=0"
      Splits(0)._ColumnProps(164)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(165)=   "Column(27)._AlignLeft=0"
      Splits(0)._ColumnProps(166)=   "Column(28).Width=2593"
      Splits(0)._ColumnProps(167)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(168)=   "Column(28)._WidthInPix=2514"
      Splits(0)._ColumnProps(169)=   "Column(28)._EditAlways=0"
      Splits(0)._ColumnProps(170)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(171)=   "Column(28)._AlignLeft=0"
      Splits(0)._ColumnProps(172)=   "Column(29).Width=2725"
      Splits(0)._ColumnProps(173)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(174)=   "Column(29)._WidthInPix=2646"
      Splits(0)._ColumnProps(175)=   "Column(29)._EditAlways=0"
      Splits(0)._ColumnProps(176)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(177)=   "Column(29)._AlignLeft=0"
      Splits(0)._ColumnProps(178)=   "Column(30).Width=794"
      Splits(0)._ColumnProps(179)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(180)=   "Column(30)._WidthInPix=714"
      Splits(0)._ColumnProps(181)=   "Column(30)._EditAlways=0"
      Splits(0)._ColumnProps(182)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(183)=   "Column(31).Width=2725"
      Splits(0)._ColumnProps(184)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(185)=   "Column(31)._WidthInPix=2646"
      Splits(0)._ColumnProps(186)=   "Column(31)._EditAlways=0"
      Splits(0)._ColumnProps(187)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(188)=   "Column(31)._AlignLeft=0"
      Splits(0)._ColumnProps(189)=   "Column(32).Width=2725"
      Splits(0)._ColumnProps(190)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(191)=   "Column(32)._WidthInPix=2646"
      Splits(0)._ColumnProps(192)=   "Column(32)._EditAlways=0"
      Splits(0)._ColumnProps(193)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(194)=   "Column(32)._AlignLeft=0"
      Splits(0)._ColumnProps(195)=   "Column(33).Width=2725"
      Splits(0)._ColumnProps(196)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(197)=   "Column(33)._WidthInPix=2646"
      Splits(0)._ColumnProps(198)=   "Column(33)._EditAlways=0"
      Splits(0)._ColumnProps(199)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(200)=   "Column(33)._AlignLeft=0"
      Splits(0)._ColumnProps(201)=   "Column(34).Width=2725"
      Splits(0)._ColumnProps(202)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(203)=   "Column(34)._WidthInPix=2646"
      Splits(0)._ColumnProps(204)=   "Column(34)._EditAlways=0"
      Splits(0)._ColumnProps(205)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(206)=   "Column(34)._AlignLeft=0"
      Splits(0)._ColumnProps(207)=   "Column(35).Width=2725"
      Splits(0)._ColumnProps(208)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(209)=   "Column(35)._WidthInPix=2646"
      Splits(0)._ColumnProps(210)=   "Column(35)._EditAlways=0"
      Splits(0)._ColumnProps(211)=   "Column(35).Order=36"
      Splits(0)._ColumnProps(212)=   "Column(35)._AlignLeft=0"
      Splits(0)._ColumnProps(213)=   "Column(36).Width=2725"
      Splits(0)._ColumnProps(214)=   "Column(36).DividerColor=0"
      Splits(0)._ColumnProps(215)=   "Column(36)._WidthInPix=2646"
      Splits(0)._ColumnProps(216)=   "Column(36)._EditAlways=0"
      Splits(0)._ColumnProps(217)=   "Column(36).Order=37"
      Splits(0)._ColumnProps(218)=   "Column(36)._AlignLeft=0"
      Splits(0)._ColumnProps(219)=   "Column(37).Width=2725"
      Splits(0)._ColumnProps(220)=   "Column(37).DividerColor=0"
      Splits(0)._ColumnProps(221)=   "Column(37)._WidthInPix=2646"
      Splits(0)._ColumnProps(222)=   "Column(37)._EditAlways=0"
      Splits(0)._ColumnProps(223)=   "Column(37).Order=38"
      Splits(0)._ColumnProps(224)=   "Column(37)._AlignLeft=0"
      Splits(0)._ColumnProps(225)=   "Column(38).Width=2725"
      Splits(0)._ColumnProps(226)=   "Column(38).DividerColor=0"
      Splits(0)._ColumnProps(227)=   "Column(38)._WidthInPix=2646"
      Splits(0)._ColumnProps(228)=   "Column(38)._EditAlways=0"
      Splits(0)._ColumnProps(229)=   "Column(38).Order=39"
      Splits(0)._ColumnProps(230)=   "Column(38)._AlignLeft=0"
      Splits(0)._ColumnProps(231)=   "Column(39).Width=2725"
      Splits(0)._ColumnProps(232)=   "Column(39).DividerColor=0"
      Splits(0)._ColumnProps(233)=   "Column(39)._WidthInPix=2646"
      Splits(0)._ColumnProps(234)=   "Column(39)._EditAlways=0"
      Splits(0)._ColumnProps(235)=   "Column(39).Order=40"
      Splits(0)._ColumnProps(236)=   "Column(39)._AlignLeft=0"
      Splits(0)._ColumnProps(237)=   "Column(40).Width=2725"
      Splits(0)._ColumnProps(238)=   "Column(40).DividerColor=0"
      Splits(0)._ColumnProps(239)=   "Column(40)._WidthInPix=2646"
      Splits(0)._ColumnProps(240)=   "Column(40)._EditAlways=0"
      Splits(0)._ColumnProps(241)=   "Column(40).Order=41"
      Splits(0)._ColumnProps(242)=   "Column(40)._AlignLeft=0"
      Splits(0)._ColumnProps(243)=   "Column(41).Width=2725"
      Splits(0)._ColumnProps(244)=   "Column(41).DividerColor=0"
      Splits(0)._ColumnProps(245)=   "Column(41)._WidthInPix=2646"
      Splits(0)._ColumnProps(246)=   "Column(41)._EditAlways=0"
      Splits(0)._ColumnProps(247)=   "Column(41).Order=42"
      Splits(0)._ColumnProps(248)=   "Column(41)._AlignLeft=0"
      Splits(0)._ColumnProps(249)=   "Column(42).Width=2725"
      Splits(0)._ColumnProps(250)=   "Column(42).DividerColor=0"
      Splits(0)._ColumnProps(251)=   "Column(42)._WidthInPix=2646"
      Splits(0)._ColumnProps(252)=   "Column(42)._EditAlways=0"
      Splits(0)._ColumnProps(253)=   "Column(42).Order=43"
      Splits(0)._ColumnProps(254)=   "Column(42)._AlignLeft=0"
      Splits(0)._ColumnProps(255)=   "Column(43).Width=2725"
      Splits(0)._ColumnProps(256)=   "Column(43).DividerColor=0"
      Splits(0)._ColumnProps(257)=   "Column(43)._WidthInPix=2646"
      Splits(0)._ColumnProps(258)=   "Column(43)._EditAlways=0"
      Splits(0)._ColumnProps(259)=   "Column(43).Order=44"
      Splits(0)._ColumnProps(260)=   "Column(43)._AlignLeft=0"
      Splits(0)._ColumnProps(261)=   "Column(44).Width=2725"
      Splits(0)._ColumnProps(262)=   "Column(44).DividerColor=0"
      Splits(0)._ColumnProps(263)=   "Column(44)._WidthInPix=2646"
      Splits(0)._ColumnProps(264)=   "Column(44)._EditAlways=0"
      Splits(0)._ColumnProps(265)=   "Column(44).Order=45"
      Splits(0)._ColumnProps(266)=   "Column(44)._AlignLeft=0"
      Splits(0)._ColumnProps(267)=   "Column(45).Width=2725"
      Splits(0)._ColumnProps(268)=   "Column(45).DividerColor=0"
      Splits(0)._ColumnProps(269)=   "Column(45)._WidthInPix=2646"
      Splits(0)._ColumnProps(270)=   "Column(45)._EditAlways=0"
      Splits(0)._ColumnProps(271)=   "Column(45).Order=46"
      Splits(0)._ColumnProps(272)=   "Column(45)._AlignLeft=0"
      Splits(0)._ColumnProps(273)=   "Column(46).Width=2725"
      Splits(0)._ColumnProps(274)=   "Column(46).DividerColor=0"
      Splits(0)._ColumnProps(275)=   "Column(46)._WidthInPix=2646"
      Splits(0)._ColumnProps(276)=   "Column(46)._EditAlways=0"
      Splits(0)._ColumnProps(277)=   "Column(46).Order=47"
      Splits(0)._ColumnProps(278)=   "Column(46)._AlignLeft=0"
      Splits(0)._ColumnProps(279)=   "Column(47).Width=2725"
      Splits(0)._ColumnProps(280)=   "Column(47).DividerColor=0"
      Splits(0)._ColumnProps(281)=   "Column(47)._WidthInPix=2646"
      Splits(0)._ColumnProps(282)=   "Column(47)._EditAlways=0"
      Splits(0)._ColumnProps(283)=   "Column(47).Order=48"
      Splits(0)._ColumnProps(284)=   "Column(47)._AlignLeft=0"
      Splits(0)._ColumnProps(285)=   "Column(48).Width=2725"
      Splits(0)._ColumnProps(286)=   "Column(48).DividerColor=0"
      Splits(0)._ColumnProps(287)=   "Column(48)._WidthInPix=2646"
      Splits(0)._ColumnProps(288)=   "Column(48)._EditAlways=0"
      Splits(0)._ColumnProps(289)=   "Column(48).Order=49"
      Splits(0)._ColumnProps(290)=   "Column(48)._AlignLeft=0"
      Splits(0)._ColumnProps(291)=   "Column(49).Width=2725"
      Splits(0)._ColumnProps(292)=   "Column(49).DividerColor=0"
      Splits(0)._ColumnProps(293)=   "Column(49)._WidthInPix=2646"
      Splits(0)._ColumnProps(294)=   "Column(49)._EditAlways=0"
      Splits(0)._ColumnProps(295)=   "Column(49).Order=50"
      Splits(0)._ColumnProps(296)=   "Column(49)._AlignLeft=0"
      Splits(0)._ColumnProps(297)=   "Column(50).Width=2725"
      Splits(0)._ColumnProps(298)=   "Column(50).DividerColor=0"
      Splits(0)._ColumnProps(299)=   "Column(50)._WidthInPix=2646"
      Splits(0)._ColumnProps(300)=   "Column(50)._EditAlways=0"
      Splits(0)._ColumnProps(301)=   "Column(50).Order=51"
      Splits(0)._ColumnProps(302)=   "Column(50)._AlignLeft=0"
      Splits(0)._ColumnProps(303)=   "Column(51).Width=2725"
      Splits(0)._ColumnProps(304)=   "Column(51).DividerColor=0"
      Splits(0)._ColumnProps(305)=   "Column(51)._WidthInPix=2646"
      Splits(0)._ColumnProps(306)=   "Column(51)._EditAlways=0"
      Splits(0)._ColumnProps(307)=   "Column(51).Order=52"
      Splits(0)._ColumnProps(308)=   "Column(51)._AlignLeft=0"
      Splits(0)._ColumnProps(309)=   "Column(52).Width=2725"
      Splits(0)._ColumnProps(310)=   "Column(52).DividerColor=0"
      Splits(0)._ColumnProps(311)=   "Column(52)._WidthInPix=2646"
      Splits(0)._ColumnProps(312)=   "Column(52)._EditAlways=0"
      Splits(0)._ColumnProps(313)=   "Column(52).Order=53"
      Splits(0)._ColumnProps(314)=   "Column(52)._AlignLeft=0"
      Splits(0)._ColumnProps(315)=   "Column(53).Width=2725"
      Splits(0)._ColumnProps(316)=   "Column(53).DividerColor=0"
      Splits(0)._ColumnProps(317)=   "Column(53)._WidthInPix=2646"
      Splits(0)._ColumnProps(318)=   "Column(53)._EditAlways=0"
      Splits(0)._ColumnProps(319)=   "Column(53).Order=54"
      Splits(0)._ColumnProps(320)=   "Column(53)._AlignLeft=0"
      Splits(0)._ColumnProps(321)=   "Column(54).Width=2725"
      Splits(0)._ColumnProps(322)=   "Column(54).DividerColor=0"
      Splits(0)._ColumnProps(323)=   "Column(54)._WidthInPix=2646"
      Splits(0)._ColumnProps(324)=   "Column(54)._EditAlways=0"
      Splits(0)._ColumnProps(325)=   "Column(54).Order=55"
      Splits(0)._ColumnProps(326)=   "Column(54)._AlignLeft=0"
      Splits(0)._ColumnProps(327)=   "Column(55).Width=2725"
      Splits(0)._ColumnProps(328)=   "Column(55).DividerColor=0"
      Splits(0)._ColumnProps(329)=   "Column(55)._WidthInPix=2646"
      Splits(0)._ColumnProps(330)=   "Column(55)._EditAlways=0"
      Splits(0)._ColumnProps(331)=   "Column(55).Order=56"
      Splits(0)._ColumnProps(332)=   "Column(55)._AlignLeft=0"
      Splits(0)._ColumnProps(333)=   "Column(56).Width=2725"
      Splits(0)._ColumnProps(334)=   "Column(56).DividerColor=0"
      Splits(0)._ColumnProps(335)=   "Column(56)._WidthInPix=2646"
      Splits(0)._ColumnProps(336)=   "Column(56)._EditAlways=0"
      Splits(0)._ColumnProps(337)=   "Column(56).Order=57"
      Splits(0)._ColumnProps(338)=   "Column(56)._AlignLeft=0"
      Splits(0)._ColumnProps(339)=   "Column(57).Width=2725"
      Splits(0)._ColumnProps(340)=   "Column(57).DividerColor=0"
      Splits(0)._ColumnProps(341)=   "Column(57)._WidthInPix=2646"
      Splits(0)._ColumnProps(342)=   "Column(57)._EditAlways=0"
      Splits(0)._ColumnProps(343)=   "Column(57).Order=58"
      Splits(0)._ColumnProps(344)=   "Column(57)._AlignLeft=0"
      Splits(0)._ColumnProps(345)=   "Column(58).Width=2725"
      Splits(0)._ColumnProps(346)=   "Column(58).DividerColor=0"
      Splits(0)._ColumnProps(347)=   "Column(58)._WidthInPix=2646"
      Splits(0)._ColumnProps(348)=   "Column(58)._EditAlways=0"
      Splits(0)._ColumnProps(349)=   "Column(58).Order=59"
      Splits(0)._ColumnProps(350)=   "Column(58)._AlignLeft=0"
      Splits(0)._ColumnProps(351)=   "Column(59).Width=1402"
      Splits(0)._ColumnProps(352)=   "Column(59).DividerColor=0"
      Splits(0)._ColumnProps(353)=   "Column(59)._WidthInPix=1323"
      Splits(0)._ColumnProps(354)=   "Column(59)._EditAlways=0"
      Splits(0)._ColumnProps(355)=   "Column(59).Order=60"
      Splits(0)._ColumnProps(356)=   "Column(60).Width=2725"
      Splits(0)._ColumnProps(357)=   "Column(60).DividerColor=0"
      Splits(0)._ColumnProps(358)=   "Column(60)._WidthInPix=2646"
      Splits(0)._ColumnProps(359)=   "Column(60)._EditAlways=0"
      Splits(0)._ColumnProps(360)=   "Column(60).Order=61"
      Splits(0)._ColumnProps(361)=   "Column(60)._AlignLeft=0"
      Splits(0)._ColumnProps(362)=   "Column(61).Width=2725"
      Splits(0)._ColumnProps(363)=   "Column(61).DividerColor=0"
      Splits(0)._ColumnProps(364)=   "Column(61)._WidthInPix=2646"
      Splits(0)._ColumnProps(365)=   "Column(61)._EditAlways=0"
      Splits(0)._ColumnProps(366)=   "Column(61).Order=62"
      Splits(0)._ColumnProps(367)=   "Column(61)._AlignLeft=0"
      Splits(0)._ColumnProps(368)=   "Column(62).Width=2725"
      Splits(0)._ColumnProps(369)=   "Column(62).DividerColor=0"
      Splits(0)._ColumnProps(370)=   "Column(62)._WidthInPix=2646"
      Splits(0)._ColumnProps(371)=   "Column(62)._EditAlways=0"
      Splits(0)._ColumnProps(372)=   "Column(62).Order=63"
      Splits(0)._ColumnProps(373)=   "Column(62)._AlignLeft=0"
      Splits(0)._ColumnProps(374)=   "Column(63).Width=2725"
      Splits(0)._ColumnProps(375)=   "Column(63).DividerColor=0"
      Splits(0)._ColumnProps(376)=   "Column(63)._WidthInPix=2646"
      Splits(0)._ColumnProps(377)=   "Column(63)._EditAlways=0"
      Splits(0)._ColumnProps(378)=   "Column(63).Order=64"
      Splits(0)._ColumnProps(379)=   "Column(63)._AlignLeft=0"
      Splits(0)._ColumnProps(380)=   "Column(64).Width=2725"
      Splits(0)._ColumnProps(381)=   "Column(64).DividerColor=0"
      Splits(0)._ColumnProps(382)=   "Column(64)._WidthInPix=2646"
      Splits(0)._ColumnProps(383)=   "Column(64)._EditAlways=0"
      Splits(0)._ColumnProps(384)=   "Column(64).Order=65"
      Splits(0)._ColumnProps(385)=   "Column(64)._AlignLeft=0"
      Splits(0)._ColumnProps(386)=   "Column(65).Width=2725"
      Splits(0)._ColumnProps(387)=   "Column(65).DividerColor=0"
      Splits(0)._ColumnProps(388)=   "Column(65)._WidthInPix=2646"
      Splits(0)._ColumnProps(389)=   "Column(65)._EditAlways=0"
      Splits(0)._ColumnProps(390)=   "Column(65).Order=66"
      Splits(0)._ColumnProps(391)=   "Column(65)._AlignLeft=0"
      Splits(0)._ColumnProps(392)=   "Column(66).Width=2725"
      Splits(0)._ColumnProps(393)=   "Column(66).DividerColor=0"
      Splits(0)._ColumnProps(394)=   "Column(66)._WidthInPix=2646"
      Splits(0)._ColumnProps(395)=   "Column(66)._EditAlways=0"
      Splits(0)._ColumnProps(396)=   "Column(66).Order=67"
      Splits(0)._ColumnProps(397)=   "Column(66)._AlignLeft=0"
      Splits(0)._ColumnProps(398)=   "Column(67).Width=2725"
      Splits(0)._ColumnProps(399)=   "Column(67).DividerColor=0"
      Splits(0)._ColumnProps(400)=   "Column(67)._WidthInPix=2646"
      Splits(0)._ColumnProps(401)=   "Column(67)._EditAlways=0"
      Splits(0)._ColumnProps(402)=   "Column(67).Order=68"
      Splits(0)._ColumnProps(403)=   "Column(67)._AlignLeft=0"
      Splits(0)._ColumnProps(404)=   "Column(68).Width=1799"
      Splits(0)._ColumnProps(405)=   "Column(68).DividerColor=0"
      Splits(0)._ColumnProps(406)=   "Column(68)._WidthInPix=1720"
      Splits(0)._ColumnProps(407)=   "Column(68)._EditAlways=0"
      Splits(0)._ColumnProps(408)=   "Column(68).Order=69"
      Splits(0)._ColumnProps(409)=   "Column(69).Width=2725"
      Splits(0)._ColumnProps(410)=   "Column(69).DividerColor=0"
      Splits(0)._ColumnProps(411)=   "Column(69)._WidthInPix=2646"
      Splits(0)._ColumnProps(412)=   "Column(69)._EditAlways=0"
      Splits(0)._ColumnProps(413)=   "Column(69).Order=70"
      Splits(0)._ColumnProps(414)=   "Column(70).Width=2725"
      Splits(0)._ColumnProps(415)=   "Column(70).DividerColor=0"
      Splits(0)._ColumnProps(416)=   "Column(70)._WidthInPix=2646"
      Splits(0)._ColumnProps(417)=   "Column(70)._EditAlways=0"
      Splits(0)._ColumnProps(418)=   "Column(70).Order=71"
      Splits(0)._ColumnProps(419)=   "Column(70)._AlignLeft=0"
      Splits(0)._ColumnProps(420)=   "Column(71).Width=2725"
      Splits(0)._ColumnProps(421)=   "Column(71).DividerColor=0"
      Splits(0)._ColumnProps(422)=   "Column(71)._WidthInPix=2646"
      Splits(0)._ColumnProps(423)=   "Column(71)._EditAlways=0"
      Splits(0)._ColumnProps(424)=   "Column(71).Order=72"
      Splits(0)._ColumnProps(425)=   "Column(71)._AlignLeft=0"
      Splits(0)._ColumnProps(426)=   "Column(72).Width=2725"
      Splits(0)._ColumnProps(427)=   "Column(72).DividerColor=0"
      Splits(0)._ColumnProps(428)=   "Column(72)._WidthInPix=2646"
      Splits(0)._ColumnProps(429)=   "Column(72)._EditAlways=0"
      Splits(0)._ColumnProps(430)=   "Column(72).Order=73"
      Splits(0)._ColumnProps(431)=   "Column(72)._AlignLeft=0"
      Splits(0)._ColumnProps(432)=   "Column(73).Width=2725"
      Splits(0)._ColumnProps(433)=   "Column(73).DividerColor=0"
      Splits(0)._ColumnProps(434)=   "Column(73)._WidthInPix=2646"
      Splits(0)._ColumnProps(435)=   "Column(73)._EditAlways=0"
      Splits(0)._ColumnProps(436)=   "Column(73).Order=74"
      Splits(0)._ColumnProps(437)=   "Column(73)._AlignLeft=0"
      Splits(0)._ColumnProps(438)=   "Column(74).Width=2725"
      Splits(0)._ColumnProps(439)=   "Column(74).DividerColor=0"
      Splits(0)._ColumnProps(440)=   "Column(74)._WidthInPix=2646"
      Splits(0)._ColumnProps(441)=   "Column(74)._EditAlways=0"
      Splits(0)._ColumnProps(442)=   "Column(74).Order=75"
      Splits(0)._ColumnProps(443)=   "Column(74)._AlignLeft=0"
      Splits(0)._ColumnProps(444)=   "Column(75).Width=2725"
      Splits(0)._ColumnProps(445)=   "Column(75).DividerColor=0"
      Splits(0)._ColumnProps(446)=   "Column(75)._WidthInPix=2646"
      Splits(0)._ColumnProps(447)=   "Column(75)._EditAlways=0"
      Splits(0)._ColumnProps(448)=   "Column(75).Order=76"
      Splits(0)._ColumnProps(449)=   "Column(75)._AlignLeft=0"
      Splits(0)._ColumnProps(450)=   "Column(76).Width=2725"
      Splits(0)._ColumnProps(451)=   "Column(76).DividerColor=0"
      Splits(0)._ColumnProps(452)=   "Column(76)._WidthInPix=2646"
      Splits(0)._ColumnProps(453)=   "Column(76)._EditAlways=0"
      Splits(0)._ColumnProps(454)=   "Column(76).Order=77"
      Splits(0)._ColumnProps(455)=   "Column(76)._AlignLeft=0"
      Splits(0)._ColumnProps(456)=   "Column(77).Width=2725"
      Splits(0)._ColumnProps(457)=   "Column(77).DividerColor=0"
      Splits(0)._ColumnProps(458)=   "Column(77)._WidthInPix=2646"
      Splits(0)._ColumnProps(459)=   "Column(77)._EditAlways=0"
      Splits(0)._ColumnProps(460)=   "Column(77).Order=78"
      Splits(0)._ColumnProps(461)=   "Column(77)._AlignLeft=0"
      Splits(0)._ColumnProps(462)=   "Column(78).Width=2725"
      Splits(0)._ColumnProps(463)=   "Column(78).DividerColor=0"
      Splits(0)._ColumnProps(464)=   "Column(78)._WidthInPix=2646"
      Splits(0)._ColumnProps(465)=   "Column(78)._EditAlways=0"
      Splits(0)._ColumnProps(466)=   "Column(78).Order=79"
      Splits(0)._ColumnProps(467)=   "Column(78)._AlignLeft=0"
      Splits(0)._ColumnProps(468)=   "Column(79).Width=2725"
      Splits(0)._ColumnProps(469)=   "Column(79).DividerColor=0"
      Splits(0)._ColumnProps(470)=   "Column(79)._WidthInPix=2646"
      Splits(0)._ColumnProps(471)=   "Column(79)._EditAlways=0"
      Splits(0)._ColumnProps(472)=   "Column(79).Order=80"
      Splits(0)._ColumnProps(473)=   "Column(80).Width=1535"
      Splits(0)._ColumnProps(474)=   "Column(80).DividerColor=0"
      Splits(0)._ColumnProps(475)=   "Column(80)._WidthInPix=1455"
      Splits(0)._ColumnProps(476)=   "Column(80)._EditAlways=0"
      Splits(0)._ColumnProps(477)=   "Column(80).Order=81"
      Splits(0)._ColumnProps(478)=   "Column(80)._AlignLeft=0"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bgcolor=&H4080FF&,.bold=0,.fontsize=825"
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
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=93,.parent=2,.namedParent=95"
      _StyleDefs(23)  =   "FilterBarStyle:id=96,.parent=1,.namedParent=98"
      _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=94,.parent=93"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=97,.parent=96"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=40,.parent=11"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=37,.parent=12"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=38,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=39,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=68,.parent=11"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=65,.parent=12"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=66,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=67,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=72,.parent=11"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=69,.parent=12"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=70,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=71,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=76,.parent=11"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=73,.parent=12"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=74,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=75,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=80,.parent=11"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=77,.parent=12"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=78,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=79,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=84,.parent=11"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=81,.parent=12"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=82,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=83,.parent=15"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=88,.parent=11"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=85,.parent=12"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=86,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=87,.parent=15"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=92,.parent=11"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=89,.parent=12"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=90,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=91,.parent=15"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=102,.parent=11"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=99,.parent=12"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=100,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=101,.parent=15"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=106,.parent=11"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=103,.parent=12"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=104,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=105,.parent=15"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=110,.parent=11"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=107,.parent=12"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=108,.parent=13"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=109,.parent=15"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=114,.parent=11"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=111,.parent=12"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=112,.parent=13"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=113,.parent=15"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=118,.parent=11"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=115,.parent=12"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=116,.parent=13"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=117,.parent=15"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=122,.parent=11"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=119,.parent=12"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=120,.parent=13"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=121,.parent=15"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=126,.parent=11"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=123,.parent=12"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=124,.parent=13"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=125,.parent=15"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=130,.parent=11"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=127,.parent=12"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=128,.parent=13"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=129,.parent=15"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=134,.parent=11"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=131,.parent=12"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=132,.parent=13"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=133,.parent=15"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=138,.parent=11"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=135,.parent=12"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=136,.parent=13"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=137,.parent=15"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=142,.parent=11"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=139,.parent=12"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=140,.parent=13"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=141,.parent=15"
      _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=146,.parent=11"
      _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=143,.parent=12"
      _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=144,.parent=13"
      _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=145,.parent=15"
      _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=150,.parent=11"
      _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=147,.parent=12"
      _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=148,.parent=13"
      _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=149,.parent=15"
      _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=154,.parent=11"
      _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=151,.parent=12"
      _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=152,.parent=13"
      _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=153,.parent=15"
      _StyleDefs(132) =   "Splits(0).Columns(24).Style:id=158,.parent=11"
      _StyleDefs(133) =   "Splits(0).Columns(24).HeadingStyle:id=155,.parent=12"
      _StyleDefs(134) =   "Splits(0).Columns(24).FooterStyle:id=156,.parent=13"
      _StyleDefs(135) =   "Splits(0).Columns(24).EditorStyle:id=157,.parent=15"
      _StyleDefs(136) =   "Splits(0).Columns(25).Style:id=162,.parent=11"
      _StyleDefs(137) =   "Splits(0).Columns(25).HeadingStyle:id=159,.parent=12"
      _StyleDefs(138) =   "Splits(0).Columns(25).FooterStyle:id=160,.parent=13"
      _StyleDefs(139) =   "Splits(0).Columns(25).EditorStyle:id=161,.parent=15"
      _StyleDefs(140) =   "Splits(0).Columns(26).Style:id=166,.parent=11"
      _StyleDefs(141) =   "Splits(0).Columns(26).HeadingStyle:id=163,.parent=12"
      _StyleDefs(142) =   "Splits(0).Columns(26).FooterStyle:id=164,.parent=13"
      _StyleDefs(143) =   "Splits(0).Columns(26).EditorStyle:id=165,.parent=15"
      _StyleDefs(144) =   "Splits(0).Columns(27).Style:id=170,.parent=11"
      _StyleDefs(145) =   "Splits(0).Columns(27).HeadingStyle:id=167,.parent=12"
      _StyleDefs(146) =   "Splits(0).Columns(27).FooterStyle:id=168,.parent=13"
      _StyleDefs(147) =   "Splits(0).Columns(27).EditorStyle:id=169,.parent=15"
      _StyleDefs(148) =   "Splits(0).Columns(28).Style:id=174,.parent=11"
      _StyleDefs(149) =   "Splits(0).Columns(28).HeadingStyle:id=171,.parent=12"
      _StyleDefs(150) =   "Splits(0).Columns(28).FooterStyle:id=172,.parent=13"
      _StyleDefs(151) =   "Splits(0).Columns(28).EditorStyle:id=173,.parent=15"
      _StyleDefs(152) =   "Splits(0).Columns(29).Style:id=178,.parent=11"
      _StyleDefs(153) =   "Splits(0).Columns(29).HeadingStyle:id=175,.parent=12"
      _StyleDefs(154) =   "Splits(0).Columns(29).FooterStyle:id=176,.parent=13"
      _StyleDefs(155) =   "Splits(0).Columns(29).EditorStyle:id=177,.parent=15"
      _StyleDefs(156) =   "Splits(0).Columns(30).Style:id=182,.parent=11"
      _StyleDefs(157) =   "Splits(0).Columns(30).HeadingStyle:id=179,.parent=12"
      _StyleDefs(158) =   "Splits(0).Columns(30).FooterStyle:id=180,.parent=13"
      _StyleDefs(159) =   "Splits(0).Columns(30).EditorStyle:id=181,.parent=15"
      _StyleDefs(160) =   "Splits(0).Columns(31).Style:id=186,.parent=11"
      _StyleDefs(161) =   "Splits(0).Columns(31).HeadingStyle:id=183,.parent=12"
      _StyleDefs(162) =   "Splits(0).Columns(31).FooterStyle:id=184,.parent=13"
      _StyleDefs(163) =   "Splits(0).Columns(31).EditorStyle:id=185,.parent=15"
      _StyleDefs(164) =   "Splits(0).Columns(32).Style:id=190,.parent=11"
      _StyleDefs(165) =   "Splits(0).Columns(32).HeadingStyle:id=187,.parent=12"
      _StyleDefs(166) =   "Splits(0).Columns(32).FooterStyle:id=188,.parent=13"
      _StyleDefs(167) =   "Splits(0).Columns(32).EditorStyle:id=189,.parent=15"
      _StyleDefs(168) =   "Splits(0).Columns(33).Style:id=194,.parent=11"
      _StyleDefs(169) =   "Splits(0).Columns(33).HeadingStyle:id=191,.parent=12"
      _StyleDefs(170) =   "Splits(0).Columns(33).FooterStyle:id=192,.parent=13"
      _StyleDefs(171) =   "Splits(0).Columns(33).EditorStyle:id=193,.parent=15"
      _StyleDefs(172) =   "Splits(0).Columns(34).Style:id=198,.parent=11"
      _StyleDefs(173) =   "Splits(0).Columns(34).HeadingStyle:id=195,.parent=12"
      _StyleDefs(174) =   "Splits(0).Columns(34).FooterStyle:id=196,.parent=13"
      _StyleDefs(175) =   "Splits(0).Columns(34).EditorStyle:id=197,.parent=15"
      _StyleDefs(176) =   "Splits(0).Columns(35).Style:id=202,.parent=11"
      _StyleDefs(177) =   "Splits(0).Columns(35).HeadingStyle:id=199,.parent=12"
      _StyleDefs(178) =   "Splits(0).Columns(35).FooterStyle:id=200,.parent=13"
      _StyleDefs(179) =   "Splits(0).Columns(35).EditorStyle:id=201,.parent=15"
      _StyleDefs(180) =   "Splits(0).Columns(36).Style:id=206,.parent=11"
      _StyleDefs(181) =   "Splits(0).Columns(36).HeadingStyle:id=203,.parent=12"
      _StyleDefs(182) =   "Splits(0).Columns(36).FooterStyle:id=204,.parent=13"
      _StyleDefs(183) =   "Splits(0).Columns(36).EditorStyle:id=205,.parent=15"
      _StyleDefs(184) =   "Splits(0).Columns(37).Style:id=210,.parent=11"
      _StyleDefs(185) =   "Splits(0).Columns(37).HeadingStyle:id=207,.parent=12"
      _StyleDefs(186) =   "Splits(0).Columns(37).FooterStyle:id=208,.parent=13"
      _StyleDefs(187) =   "Splits(0).Columns(37).EditorStyle:id=209,.parent=15"
      _StyleDefs(188) =   "Splits(0).Columns(38).Style:id=214,.parent=11"
      _StyleDefs(189) =   "Splits(0).Columns(38).HeadingStyle:id=211,.parent=12"
      _StyleDefs(190) =   "Splits(0).Columns(38).FooterStyle:id=212,.parent=13"
      _StyleDefs(191) =   "Splits(0).Columns(38).EditorStyle:id=213,.parent=15"
      _StyleDefs(192) =   "Splits(0).Columns(39).Style:id=218,.parent=11"
      _StyleDefs(193) =   "Splits(0).Columns(39).HeadingStyle:id=215,.parent=12"
      _StyleDefs(194) =   "Splits(0).Columns(39).FooterStyle:id=216,.parent=13"
      _StyleDefs(195) =   "Splits(0).Columns(39).EditorStyle:id=217,.parent=15"
      _StyleDefs(196) =   "Splits(0).Columns(40).Style:id=222,.parent=11"
      _StyleDefs(197) =   "Splits(0).Columns(40).HeadingStyle:id=219,.parent=12"
      _StyleDefs(198) =   "Splits(0).Columns(40).FooterStyle:id=220,.parent=13"
      _StyleDefs(199) =   "Splits(0).Columns(40).EditorStyle:id=221,.parent=15"
      _StyleDefs(200) =   "Splits(0).Columns(41).Style:id=226,.parent=11"
      _StyleDefs(201) =   "Splits(0).Columns(41).HeadingStyle:id=223,.parent=12"
      _StyleDefs(202) =   "Splits(0).Columns(41).FooterStyle:id=224,.parent=13"
      _StyleDefs(203) =   "Splits(0).Columns(41).EditorStyle:id=225,.parent=15"
      _StyleDefs(204) =   "Splits(0).Columns(42).Style:id=230,.parent=11"
      _StyleDefs(205) =   "Splits(0).Columns(42).HeadingStyle:id=227,.parent=12"
      _StyleDefs(206) =   "Splits(0).Columns(42).FooterStyle:id=228,.parent=13"
      _StyleDefs(207) =   "Splits(0).Columns(42).EditorStyle:id=229,.parent=15"
      _StyleDefs(208) =   "Splits(0).Columns(43).Style:id=234,.parent=11"
      _StyleDefs(209) =   "Splits(0).Columns(43).HeadingStyle:id=231,.parent=12"
      _StyleDefs(210) =   "Splits(0).Columns(43).FooterStyle:id=232,.parent=13"
      _StyleDefs(211) =   "Splits(0).Columns(43).EditorStyle:id=233,.parent=15"
      _StyleDefs(212) =   "Splits(0).Columns(44).Style:id=238,.parent=11"
      _StyleDefs(213) =   "Splits(0).Columns(44).HeadingStyle:id=235,.parent=12"
      _StyleDefs(214) =   "Splits(0).Columns(44).FooterStyle:id=236,.parent=13"
      _StyleDefs(215) =   "Splits(0).Columns(44).EditorStyle:id=237,.parent=15"
      _StyleDefs(216) =   "Splits(0).Columns(45).Style:id=242,.parent=11"
      _StyleDefs(217) =   "Splits(0).Columns(45).HeadingStyle:id=239,.parent=12"
      _StyleDefs(218) =   "Splits(0).Columns(45).FooterStyle:id=240,.parent=13"
      _StyleDefs(219) =   "Splits(0).Columns(45).EditorStyle:id=241,.parent=15"
      _StyleDefs(220) =   "Splits(0).Columns(46).Style:id=246,.parent=11"
      _StyleDefs(221) =   "Splits(0).Columns(46).HeadingStyle:id=243,.parent=12"
      _StyleDefs(222) =   "Splits(0).Columns(46).FooterStyle:id=244,.parent=13"
      _StyleDefs(223) =   "Splits(0).Columns(46).EditorStyle:id=245,.parent=15"
      _StyleDefs(224) =   "Splits(0).Columns(47).Style:id=250,.parent=11"
      _StyleDefs(225) =   "Splits(0).Columns(47).HeadingStyle:id=247,.parent=12"
      _StyleDefs(226) =   "Splits(0).Columns(47).FooterStyle:id=248,.parent=13"
      _StyleDefs(227) =   "Splits(0).Columns(47).EditorStyle:id=249,.parent=15"
      _StyleDefs(228) =   "Splits(0).Columns(48).Style:id=254,.parent=11"
      _StyleDefs(229) =   "Splits(0).Columns(48).HeadingStyle:id=251,.parent=12"
      _StyleDefs(230) =   "Splits(0).Columns(48).FooterStyle:id=252,.parent=13"
      _StyleDefs(231) =   "Splits(0).Columns(48).EditorStyle:id=253,.parent=15"
      _StyleDefs(232) =   "Splits(0).Columns(49).Style:id=258,.parent=11"
      _StyleDefs(233) =   "Splits(0).Columns(49).HeadingStyle:id=255,.parent=12"
      _StyleDefs(234) =   "Splits(0).Columns(49).FooterStyle:id=256,.parent=13"
      _StyleDefs(235) =   "Splits(0).Columns(49).EditorStyle:id=257,.parent=15"
      _StyleDefs(236) =   "Splits(0).Columns(50).Style:id=262,.parent=11"
      _StyleDefs(237) =   "Splits(0).Columns(50).HeadingStyle:id=259,.parent=12"
      _StyleDefs(238) =   "Splits(0).Columns(50).FooterStyle:id=260,.parent=13"
      _StyleDefs(239) =   "Splits(0).Columns(50).EditorStyle:id=261,.parent=15"
      _StyleDefs(240) =   "Splits(0).Columns(51).Style:id=266,.parent=11"
      _StyleDefs(241) =   "Splits(0).Columns(51).HeadingStyle:id=263,.parent=12"
      _StyleDefs(242) =   "Splits(0).Columns(51).FooterStyle:id=264,.parent=13"
      _StyleDefs(243) =   "Splits(0).Columns(51).EditorStyle:id=265,.parent=15"
      _StyleDefs(244) =   "Splits(0).Columns(52).Style:id=270,.parent=11"
      _StyleDefs(245) =   "Splits(0).Columns(52).HeadingStyle:id=267,.parent=12"
      _StyleDefs(246) =   "Splits(0).Columns(52).FooterStyle:id=268,.parent=13"
      _StyleDefs(247) =   "Splits(0).Columns(52).EditorStyle:id=269,.parent=15"
      _StyleDefs(248) =   "Splits(0).Columns(53).Style:id=274,.parent=11"
      _StyleDefs(249) =   "Splits(0).Columns(53).HeadingStyle:id=271,.parent=12"
      _StyleDefs(250) =   "Splits(0).Columns(53).FooterStyle:id=272,.parent=13"
      _StyleDefs(251) =   "Splits(0).Columns(53).EditorStyle:id=273,.parent=15"
      _StyleDefs(252) =   "Splits(0).Columns(54).Style:id=278,.parent=11"
      _StyleDefs(253) =   "Splits(0).Columns(54).HeadingStyle:id=275,.parent=12"
      _StyleDefs(254) =   "Splits(0).Columns(54).FooterStyle:id=276,.parent=13"
      _StyleDefs(255) =   "Splits(0).Columns(54).EditorStyle:id=277,.parent=15"
      _StyleDefs(256) =   "Splits(0).Columns(55).Style:id=282,.parent=11"
      _StyleDefs(257) =   "Splits(0).Columns(55).HeadingStyle:id=279,.parent=12"
      _StyleDefs(258) =   "Splits(0).Columns(55).FooterStyle:id=280,.parent=13"
      _StyleDefs(259) =   "Splits(0).Columns(55).EditorStyle:id=281,.parent=15"
      _StyleDefs(260) =   "Splits(0).Columns(56).Style:id=286,.parent=11"
      _StyleDefs(261) =   "Splits(0).Columns(56).HeadingStyle:id=283,.parent=12"
      _StyleDefs(262) =   "Splits(0).Columns(56).FooterStyle:id=284,.parent=13"
      _StyleDefs(263) =   "Splits(0).Columns(56).EditorStyle:id=285,.parent=15"
      _StyleDefs(264) =   "Splits(0).Columns(57).Style:id=290,.parent=11"
      _StyleDefs(265) =   "Splits(0).Columns(57).HeadingStyle:id=287,.parent=12"
      _StyleDefs(266) =   "Splits(0).Columns(57).FooterStyle:id=288,.parent=13"
      _StyleDefs(267) =   "Splits(0).Columns(57).EditorStyle:id=289,.parent=15"
      _StyleDefs(268) =   "Splits(0).Columns(58).Style:id=294,.parent=11"
      _StyleDefs(269) =   "Splits(0).Columns(58).HeadingStyle:id=291,.parent=12"
      _StyleDefs(270) =   "Splits(0).Columns(58).FooterStyle:id=292,.parent=13"
      _StyleDefs(271) =   "Splits(0).Columns(58).EditorStyle:id=293,.parent=15"
      _StyleDefs(272) =   "Splits(0).Columns(59).Style:id=298,.parent=11"
      _StyleDefs(273) =   "Splits(0).Columns(59).HeadingStyle:id=295,.parent=12"
      _StyleDefs(274) =   "Splits(0).Columns(59).FooterStyle:id=296,.parent=13"
      _StyleDefs(275) =   "Splits(0).Columns(59).EditorStyle:id=297,.parent=15"
      _StyleDefs(276) =   "Splits(0).Columns(60).Style:id=302,.parent=11"
      _StyleDefs(277) =   "Splits(0).Columns(60).HeadingStyle:id=299,.parent=12"
      _StyleDefs(278) =   "Splits(0).Columns(60).FooterStyle:id=300,.parent=13"
      _StyleDefs(279) =   "Splits(0).Columns(60).EditorStyle:id=301,.parent=15"
      _StyleDefs(280) =   "Splits(0).Columns(61).Style:id=306,.parent=11"
      _StyleDefs(281) =   "Splits(0).Columns(61).HeadingStyle:id=303,.parent=12"
      _StyleDefs(282) =   "Splits(0).Columns(61).FooterStyle:id=304,.parent=13"
      _StyleDefs(283) =   "Splits(0).Columns(61).EditorStyle:id=305,.parent=15"
      _StyleDefs(284) =   "Splits(0).Columns(62).Style:id=310,.parent=11"
      _StyleDefs(285) =   "Splits(0).Columns(62).HeadingStyle:id=307,.parent=12"
      _StyleDefs(286) =   "Splits(0).Columns(62).FooterStyle:id=308,.parent=13"
      _StyleDefs(287) =   "Splits(0).Columns(62).EditorStyle:id=309,.parent=15"
      _StyleDefs(288) =   "Splits(0).Columns(63).Style:id=314,.parent=11"
      _StyleDefs(289) =   "Splits(0).Columns(63).HeadingStyle:id=311,.parent=12"
      _StyleDefs(290) =   "Splits(0).Columns(63).FooterStyle:id=312,.parent=13"
      _StyleDefs(291) =   "Splits(0).Columns(63).EditorStyle:id=313,.parent=15"
      _StyleDefs(292) =   "Splits(0).Columns(64).Style:id=318,.parent=11"
      _StyleDefs(293) =   "Splits(0).Columns(64).HeadingStyle:id=315,.parent=12"
      _StyleDefs(294) =   "Splits(0).Columns(64).FooterStyle:id=316,.parent=13"
      _StyleDefs(295) =   "Splits(0).Columns(64).EditorStyle:id=317,.parent=15"
      _StyleDefs(296) =   "Splits(0).Columns(65).Style:id=322,.parent=11"
      _StyleDefs(297) =   "Splits(0).Columns(65).HeadingStyle:id=319,.parent=12"
      _StyleDefs(298) =   "Splits(0).Columns(65).FooterStyle:id=320,.parent=13"
      _StyleDefs(299) =   "Splits(0).Columns(65).EditorStyle:id=321,.parent=15"
      _StyleDefs(300) =   "Splits(0).Columns(66).Style:id=326,.parent=11"
      _StyleDefs(301) =   "Splits(0).Columns(66).HeadingStyle:id=323,.parent=12"
      _StyleDefs(302) =   "Splits(0).Columns(66).FooterStyle:id=324,.parent=13"
      _StyleDefs(303) =   "Splits(0).Columns(66).EditorStyle:id=325,.parent=15"
      _StyleDefs(304) =   "Splits(0).Columns(67).Style:id=330,.parent=11"
      _StyleDefs(305) =   "Splits(0).Columns(67).HeadingStyle:id=327,.parent=12"
      _StyleDefs(306) =   "Splits(0).Columns(67).FooterStyle:id=328,.parent=13"
      _StyleDefs(307) =   "Splits(0).Columns(67).EditorStyle:id=329,.parent=15"
      _StyleDefs(308) =   "Splits(0).Columns(68).Style:id=334,.parent=11"
      _StyleDefs(309) =   "Splits(0).Columns(68).HeadingStyle:id=331,.parent=12"
      _StyleDefs(310) =   "Splits(0).Columns(68).FooterStyle:id=332,.parent=13"
      _StyleDefs(311) =   "Splits(0).Columns(68).EditorStyle:id=333,.parent=15"
      _StyleDefs(312) =   "Splits(0).Columns(69).Style:id=338,.parent=11"
      _StyleDefs(313) =   "Splits(0).Columns(69).HeadingStyle:id=335,.parent=12"
      _StyleDefs(314) =   "Splits(0).Columns(69).FooterStyle:id=336,.parent=13"
      _StyleDefs(315) =   "Splits(0).Columns(69).EditorStyle:id=337,.parent=15"
      _StyleDefs(316) =   "Splits(0).Columns(70).Style:id=342,.parent=11"
      _StyleDefs(317) =   "Splits(0).Columns(70).HeadingStyle:id=339,.parent=12"
      _StyleDefs(318) =   "Splits(0).Columns(70).FooterStyle:id=340,.parent=13"
      _StyleDefs(319) =   "Splits(0).Columns(70).EditorStyle:id=341,.parent=15"
      _StyleDefs(320) =   "Splits(0).Columns(71).Style:id=346,.parent=11"
      _StyleDefs(321) =   "Splits(0).Columns(71).HeadingStyle:id=343,.parent=12"
      _StyleDefs(322) =   "Splits(0).Columns(71).FooterStyle:id=344,.parent=13"
      _StyleDefs(323) =   "Splits(0).Columns(71).EditorStyle:id=345,.parent=15"
      _StyleDefs(324) =   "Splits(0).Columns(72).Style:id=350,.parent=11"
      _StyleDefs(325) =   "Splits(0).Columns(72).HeadingStyle:id=347,.parent=12"
      _StyleDefs(326) =   "Splits(0).Columns(72).FooterStyle:id=348,.parent=13"
      _StyleDefs(327) =   "Splits(0).Columns(72).EditorStyle:id=349,.parent=15"
      _StyleDefs(328) =   "Splits(0).Columns(73).Style:id=354,.parent=11"
      _StyleDefs(329) =   "Splits(0).Columns(73).HeadingStyle:id=351,.parent=12"
      _StyleDefs(330) =   "Splits(0).Columns(73).FooterStyle:id=352,.parent=13"
      _StyleDefs(331) =   "Splits(0).Columns(73).EditorStyle:id=353,.parent=15"
      _StyleDefs(332) =   "Splits(0).Columns(74).Style:id=358,.parent=11"
      _StyleDefs(333) =   "Splits(0).Columns(74).HeadingStyle:id=355,.parent=12"
      _StyleDefs(334) =   "Splits(0).Columns(74).FooterStyle:id=356,.parent=13"
      _StyleDefs(335) =   "Splits(0).Columns(74).EditorStyle:id=357,.parent=15"
      _StyleDefs(336) =   "Splits(0).Columns(75).Style:id=362,.parent=11"
      _StyleDefs(337) =   "Splits(0).Columns(75).HeadingStyle:id=359,.parent=12"
      _StyleDefs(338) =   "Splits(0).Columns(75).FooterStyle:id=360,.parent=13"
      _StyleDefs(339) =   "Splits(0).Columns(75).EditorStyle:id=361,.parent=15"
      _StyleDefs(340) =   "Splits(0).Columns(76).Style:id=366,.parent=11"
      _StyleDefs(341) =   "Splits(0).Columns(76).HeadingStyle:id=363,.parent=12"
      _StyleDefs(342) =   "Splits(0).Columns(76).FooterStyle:id=364,.parent=13"
      _StyleDefs(343) =   "Splits(0).Columns(76).EditorStyle:id=365,.parent=15"
      _StyleDefs(344) =   "Splits(0).Columns(77).Style:id=370,.parent=11"
      _StyleDefs(345) =   "Splits(0).Columns(77).HeadingStyle:id=367,.parent=12"
      _StyleDefs(346) =   "Splits(0).Columns(77).FooterStyle:id=368,.parent=13"
      _StyleDefs(347) =   "Splits(0).Columns(77).EditorStyle:id=369,.parent=15"
      _StyleDefs(348) =   "Splits(0).Columns(78).Style:id=374,.parent=11"
      _StyleDefs(349) =   "Splits(0).Columns(78).HeadingStyle:id=371,.parent=12"
      _StyleDefs(350) =   "Splits(0).Columns(78).FooterStyle:id=372,.parent=13"
      _StyleDefs(351) =   "Splits(0).Columns(78).EditorStyle:id=373,.parent=15"
      _StyleDefs(352) =   "Splits(0).Columns(79).Style:id=378,.parent=11"
      _StyleDefs(353) =   "Splits(0).Columns(79).HeadingStyle:id=375,.parent=12"
      _StyleDefs(354) =   "Splits(0).Columns(79).FooterStyle:id=376,.parent=13"
      _StyleDefs(355) =   "Splits(0).Columns(79).EditorStyle:id=377,.parent=15"
      _StyleDefs(356) =   "Splits(0).Columns(80).Style:id=382,.parent=11"
      _StyleDefs(357) =   "Splits(0).Columns(80).HeadingStyle:id=379,.parent=12"
      _StyleDefs(358) =   "Splits(0).Columns(80).FooterStyle:id=380,.parent=13"
      _StyleDefs(359) =   "Splits(0).Columns(80).EditorStyle:id=381,.parent=15"
      _StyleDefs(360) =   "Named:id=29:Normal"
      _StyleDefs(361) =   ":id=29,.parent=0"
      _StyleDefs(362) =   "Named:id=30:Heading"
      _StyleDefs(363) =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(364) =   ":id=30,.wraptext=-1"
      _StyleDefs(365) =   "Named:id=31:Footing"
      _StyleDefs(366) =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(367) =   "Named:id=32:Selected"
      _StyleDefs(368) =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(369) =   "Named:id=33:Caption"
      _StyleDefs(370) =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(371) =   "Named:id=34:HighlightRow"
      _StyleDefs(372) =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(373) =   "Named:id=35:EvenRow"
      _StyleDefs(374) =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(375) =   "Named:id=36:OddRow"
      _StyleDefs(376) =   ":id=36,.parent=29"
      _StyleDefs(377) =   "Named:id=95:RecordSelector"
      _StyleDefs(378) =   ":id=95,.parent=30"
      _StyleDefs(379) =   "Named:id=98:FilterBar"
      _StyleDefs(380) =   ":id=98,.parent=29"
   End
   Begin TrueOleDBGrid80.TDBDropDown TDBDropDown2 
      Bindings        =   "APOT6.frx":002A
      Height          =   1200
      Left            =   375
      TabIndex        =   2
      Top             =   13725
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2117
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "AYJON"
      Columns(0).DataField=   "AYJON"
      Columns(0).DataWidth=   23
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "PERIGRAFH"
      Columns(1).DataField=   "PERIGRAFH"
      Columns(1).DataWidth=   20
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1164"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1085"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._AlignLeft=0"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   1
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   2
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   0   'False
      ListField       =   "PERIGRAFH"
      DataField       =   "AYJON"
      IntegralHeight  =   -1  'True
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   12632256
      ValueTranslate  =   0   'False
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=161"
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
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(23)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=11"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=11"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=12"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=15"
      _StyleDefs(44)  =   "Named:id=29:Normal"
      _StyleDefs(45)  =   ":id=29,.parent=0"
      _StyleDefs(46)  =   "Named:id=30:Heading"
      _StyleDefs(47)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=30,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=31:Footing"
      _StyleDefs(50)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=32:Selected"
      _StyleDefs(52)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=33:Caption"
      _StyleDefs(54)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(55)  =   "Named:id=34:HighlightRow"
      _StyleDefs(56)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(57)  =   "Named:id=35:EvenRow"
      _StyleDefs(58)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=36:OddRow"
      _StyleDefs(60)  =   ":id=36,.parent=29"
      _StyleDefs(61)  =   "Named:id=39:RecordSelector"
      _StyleDefs(62)  =   ":id=39,.parent=30"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=29"
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   2805
      Top             =   14355
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      DataSourceName  =   "MERCCOMPAQ"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "OIKOGENEIES"
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
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   2040
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   188350465
      CurrentDate     =   38814
   End
   Begin MSComCtl2.DTPicker eos 
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   2640
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   188350465
      CurrentDate     =   38814
   End
   Begin VB.Shape Shape 
      Height          =   3255
      Left            =   720
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "% ‰˘ÒÔÂÈÙ·„ÁÚ"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   6120
      TabIndex        =   21
      Top             =   1200
      Width           =   1335
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
      Index           =   3
      Left            =   2160
      TabIndex        =   20
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label5 
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
      Left            =   2160
      TabIndex        =   19
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "≈Ÿ”  Ÿƒ… œ "
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   14
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "¡–œ  Ÿƒ… œ "
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   13
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Í·È Ì· ÏÁÌ ÂÒÈ›˜ÂÈ ÙÁÌ Î›ÓÁ :"
      Height          =   285
      Left            =   1410
      TabIndex        =   12
      Top             =   15840
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.Label Label3 
      Caption         =   "’ÔÔÈÍÔ„›ÌÂÈ·"
      Height          =   285
      Left            =   5850
      TabIndex        =   8
      Top             =   16185
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label2 
      Caption         =   "OÈÍÔ„›ÌÂÈ·"
      Height          =   285
      Left            =   4545
      TabIndex        =   7
      Top             =   16155
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "‘Ô ¸ÌÔÏ· Ì· ›˜ÂÈ ÙÁÌ Î›ÓÁ :"
      Height          =   285
      Index           =   0
      Left            =   1425
      TabIndex        =   4
      Top             =   15465
      Visible         =   0   'False
      Width           =   3045
   End
End
Attribute VB_Name = "APOT6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RSClone As ADODB.Recordset

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

'Private Sub Command1_Click()
'    On Error GoTo LATOS
'    Gdb.Execute "UPDATE   EID SET AEG=" + oik + " , KODLOG='" + ypooik.Text + "'  where   ONO LIKE '" + sql + "%' "
'    Exit Sub
'LATOS:
'    MsgBox Err.Description
'    Resume Next
'
'
'End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>

100     If Len(sql) = 0 Then
110         Adodc1.RecordSource = "SELECT  AEG,KODLOG,ONO  FROM EID where AEG IS NULL OR AEG=0 ORDER BY ONO    "
        Else

120         If Len(notsql.Text) > 0 Then
130             Adodc1.RecordSource = "SELECT  AEG,KODLOG,ONO  FROM EID where   ONO LIKE '" + sql + "%' and not ONO LIKE '" + notsql + "%'  ORDER BY ONO    "

            Else
140             Adodc1.RecordSource = "SELECT  AEG,KODLOG,ONO  FROM EID where   ONO LIKE '" + sql + "%'  ORDER BY ONO    "
            End If
        End If

150     Adodc1.Refresh

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT6.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT6.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>

100     Unload Me
110     SendKeys "%"

120     SendKeys "{RIGHT}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}"

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT6.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT6.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then
110         Command3_Click
        End If

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT6.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT6.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     mForm_Load Me, fh, fw, ft, fl

110     Me.KeyPreview = True

120     Adodc1.ConnectionString = gConnect
130     Adodc2.ConnectionString = gConnect
140     Adodc3.ConnectionString = gConnect

150     Adodc2.RecordSource = "SELECT * FROM PINAKES WHERE TYPOS=3 ORDER BY PERIGRAFH"
160     Adodc3.RecordSource = "SELECT * FROM PINAKES WHERE TYPOS=11 ORDER BY PERIGRAFH "
170     Adodc1.RecordSource = "SELECT  AEG,KODLOG,ONO,XTI  FROM EID where POS>0 AND (AEG IS NULL OR AEG=0) ORDER BY ONO    "

180     Adodc1.Refresh

190     Adodc2.Refresh
200     Adodc3.Refresh

        Dim R As New ADODB.Recordset

210     R.Open "SELECT COUNT(*) FROM EID WHERE   POS>0 AND ( AEG IS NULL OR AEG=0)  ", Gdb, adOpenForwardOnly, adLockReadOnly
220     Me.Caption = str(R(0)) + " ◊Ÿ—…”  ¡‘«√œ—…¡"

        ' Set RSClone = Adodc1.Recordset.Clone()
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT6.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT6.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 10980, 9615, fh, fw, ft, fl
End Sub

Private Sub TDBDropDown1_DropDownOpen()

        '<EhHeader>
        On Error GoTo TDBDropDown1_DropDownOpen_Err

        '</EhHeader>

100     Adodc2.RecordSource = "SELECT * FROM PINAKES WHERE TYPOS=3 AND TIMH= " + str(Adodc1.Recordset("AEG")) + "  ORDER BY PERIGRAFH"
110     Adodc2.Refresh
120     Me.Caption = str(Adodc1.Recordset("AEG"))

130     TDBDropDown1.Rebind
140     TDBDropDown1.Refresh

        '<EhFooter>
        Exit Sub

TDBDropDown1_DropDownOpen_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT6.TDBDropDown1_DropDownOpen " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT6.TDBDropDown1_DropDownOpen " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub YPOL_Click()

        '<EhHeader>
        On Error GoTo YPOL_Click_Err

        '</EhHeader>
        Dim k As Long

100     Gdb.Execute "UPDATE EID SET NUM1=" + str(DORO) + " , HM1='" + Format(apo, "MM/DD/YYYY") + "' AND HM2='" + Format(eos, "MM/DD/YYYY") + "' WHERE KOD>='" + apok.Text + "' AND KOD<='" + eosk.Text + "'", k

110     MsgBox "≈Õ«Ã≈—Ÿ»« ¡Õ " + str(k) + "  Ÿƒ… œ…"

        '<EhFooter>
        Exit Sub

YPOL_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.APOT6.YPOL_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.APOT6.YPOL_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
