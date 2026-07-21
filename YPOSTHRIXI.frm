VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form YPOSTHRIXI 
   Caption         =   "Form2"
   ClientHeight    =   13140
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   21885
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
   ScaleHeight     =   13140
   ScaleWidth      =   21885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox XRONIKA 
      Height          =   315
      ItemData        =   "YPOSTHRIXI.frx":0000
      Left            =   11400
      List            =   "YPOSTHRIXI.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdEYRESHPEL 
      Caption         =   "EYRESHPEL"
      Height          =   360
      Left            =   19320
      TabIndex        =   11
      Top             =   7320
      Width           =   990
   End
   Begin VB.ListBox LISTPEL 
      Height          =   1620
      Left            =   14760
      TabIndex        =   10
      Top             =   7920
      Width           =   4935
   End
   Begin VB.TextBox MEPO 
      Height          =   375
      Left            =   16440
      TabIndex        =   8
      Top             =   7320
      Width           =   2775
   End
   Begin VB.CommandButton cmdDiort 
      Caption         =   "diort"
      Height          =   360
      Left            =   9480
      TabIndex        =   7
      Top             =   7320
      Width           =   1695
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Bindings        =   "YPOSTHRIXI.frx":0038
      Height          =   6855
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   24015
      _ExtentX        =   42360
      _ExtentY        =   12091
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
      Splits(0).RecordSelectorWidth=   979
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   15790320
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
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   15790320
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
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
      _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
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
   Begin VB.CommandButton cmd¡Õ¡Õ≈Ÿ”« 
      BackColor       =   &H0080FF80&
      Caption         =   "¡Õ¡Õ≈Ÿ”«"
      Height          =   360
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   1935
   End
   Begin VB.ComboBox Com 
      Height          =   315
      ItemData        =   "YPOSTHRIXI.frx":004D
      Left            =   6360
      List            =   "YPOSTHRIXI.frx":0060
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   7320
      Width           =   2535
   End
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "ÕÂ· ≈„„Ò·ˆﬁ"
      Height          =   360
      Left            =   2760
      TabIndex        =   0
      Top             =   7200
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   840
      Top             =   12840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1085
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
         Name            =   "Tahoma"
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
      Bindings        =   "YPOSTHRIXI.frx":0091
      Height          =   3615
      Left            =   960
      TabIndex        =   6
      Top             =   7680
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   6376
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
      Splits(0).RecordSelectorWidth=   979
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   15790320
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
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=161,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   15790320
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
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
      _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=161"
      _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
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
   Begin MSAdodcLib.Adodc AdodcSX 
      Height          =   615
      Left            =   9000
      Top             =   12840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1085
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
      Caption         =   "AdodcSX"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lbl–≈À¡‘«” 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "–≈À¡‘«”"
      Height          =   195
      Left            =   14640
      TabIndex        =   9
      Top             =   7440
      Width           =   690
   End
   Begin VB.Label lblÃMERCURY_ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   585
      Left            =   9120
      TabIndex        =   3
      Top             =   8520
      Width           =   1425
   End
   Begin VB.Label lblUSER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USER"
      Height          =   195
      Left            =   5520
      TabIndex        =   2
      Top             =   7320
      Width           =   390
   End
End
Attribute VB_Name = "YPOSTHRIXI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Col         As TrueOleDBGrid80.Column
Dim ColS As TrueOleDBGrid80.columns

Private Sub cmdCommand1_Click()
If gVal(Com.Text) > 0 Then
   Gdb.Execute "INSERT INTO THLEFONA(HME,XRHSTHS) VALUES (GETDATE()," + str(gVal(Com.Text)) + ")"
   ' Adodc2.Refresh
    GRIDREFRESH
   
Else
   MsgBox "DIALEJE ◊—«”‘«"
End If


'If StatusBar1.Panels(1).Text = Panel Then
'     R.Open " SELECT TOP 20  THLEFONA.THL,HME,PEL.EPO  FROM THLEFONA INNER JOIN PEL ON PEL.EIDOS=THLEFONA.EIDOS AND PEL.KOD=THLEFONA.KOD ORDER BY HME DESC", Gdb, adOpenForwardOnly, adLockReadOnly
'
'     mc = 0
'     runt(0).Caption = R(0)
'     ETAIR(0) = R(0)
'
'     Do While Not R.EOF
'         mc = mc + 1
'         Load runt(mc)
'         runt(mc).Caption = R(0) + "*" + Format(R!hme, "HH:MM") + "*" + R!EPO
'         ETAIR(mc) = Trim(R(0))
'         R.MoveNext
'     Loop
'
'     PopupMenu RUNTIME(0)
'
'        Exit Sub
'
'
'
'
'
'End If


End Sub

Private Sub cmdDiort_Click()

   TDBGrid.Col = 2
   Dim MKOD As String: MKOD = TDBGrid.Text
   


        '<EhHeader>
       ' On Error GoTo Command6_Click_Err

        '</EhHeader>
    
        Dim fff As New pelat2

        'fff.ffbuff = f_buff + data1.Recordset("kod") ' Text1(0).Text
   
100     gApoMenu = False
        'fff.COMMAND5.Enabled = False
    
102    ' If POLAGOR.ListIndex = 0 Then
104         fff.ffbuff = "e" + MKOD 'data1.Recordset("kod")    ' Text1(0).Text

        
 
142     fff.SHOW
      
144     gApoMenu = True


End Sub

Private Sub cmdEYRESHPEL_Click()
Dim R As New ADODB.Recordset
   R.Open "SELECT KOD,EPO FROM PEL WHERE KOD LIKE '%" + Trim$(MEPO.Text) + "%' OR EPO LIKE '%" + Trim(MEPO.Text) + "%' ", Gdb, adOpenDynamic, adLockOptimistic
   Dim I As Integer: I = 0
   Do While Not R.EOF
       LISTPEL.AddItem R!kod + ";" + R!EPO
       R.MoveNext
       I = I + 1
       If I > 300 Then Exit Do
   Loop
   R.Close
   
   
   
   
   
   
End Sub

Private Sub cmd¡Õ¡Õ≈Ÿ”«_Click()
  'Adodc2.Refresh
  GRIDREFRESH
End Sub

Private Sub Form_Load()
        XRONIKA.Text = XRONIKA.List(0)
        TDBGrid.FilterBar = True
        TDBGrid.Splits(0).AlternatingRowStyle = True
        TDBGrid.Splits(0).ExtendRightColumn = True
        
    'TDBGrid1.Splits(0).columns(0).Locked = True
        
        
        
              Adodc2.ConnectionString = gConnect
              'Adodc2.RecordSource = "SELECT TOP 100 THL,HME,KOD,EPO,SXOLIA,APOK,XRHSTHS,TYPOS,ID  FROM THLEFONA where DATEDIFF(HOUR,HME,GETDATE())<=14 ORDER BY ID DESC" ' Text2.Text
              GRIDREFRESH
End Sub

Sub GRIDREFRESH()

synt = " where DATEDIFF(HOUR,HME,GETDATE())<=14 "
If Val(Left(XRONIKA.Text, 1)) = 1 Then
      synt = " where DATEDIFF(HOUR,HME,GETDATE())<=14 AND DAY(HME)=DAY(GETDATE()) "
ElseIf Val(Left(XRONIKA.Text, 1)) = 2 Then
        synt = " where DATEDIFF(DAY,HME,GETDATE())=1 "
Else
      synt = ""
End If
             Adodc2.ConnectionString = gConnect

              Adodc2.RecordSource = "SELECT TOP 100 THL,HME,KOD,EPO,SXOLIA,APOK,XRHSTHS,TYPOS,ID  FROM THLEFONA " + synt + " ORDER BY ID DESC" ' Text2.Text
        
              Adodc2.Refresh
              
               TDBGrid.columns(0).Width = 1300  '‘«À
                  TDBGrid.columns(1).Width = 2300 '«Ã≈
         TDBGrid.Splits(0).columns(3).Width = 3400 '≈–ŸÕ
          TDBGrid.Splits(0).columns(4).Width = 5000 '”◊œÀ…¡
           TDBGrid.Splits(0).columns(5).Width = 5100 '¡–œ 
            TDBGrid.Splits(0).columns(7).Width = 2000
              
    
    AdodcSX.ConnectionString = gConnect
    AdodcSX.RecordSource = "SELECT  DATEPART(HOUR, HME) AS [Ÿ—¡] , COUNT(*) AS [¡—…».‘«À≈÷] FROM THLEFONA  " + synt + "  GROUP BY DATEPART(HOUR, HME) ORDER BY DATEPART(HOUR, HME) "
             
     AdodcSX.Refresh
             
             
             
              
End Sub









Private Sub TDBGrid_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If ColIndex = 0 Then
        Cancel = True ' ??????e? t?? epe?e??as?a
    End If
End Sub

Private Sub TDBGrid_Click()
Dim MCOL As Integer
'
MCOL = TDBGrid.Col
'
'TDBGrid.Refresh
'
TDBGrid.Col = 0
    Dim MT As String
    MT = TDBGrid.Text
'
     AdodcSX.RecordSource = "SELECT TOP 100 HME,SXOLIA,APOK,XRHSTHS,TYPOS,ID  FROM THLEFONA where THL='" + Trim(MT) + "' AND DATEDIFF(DAY,HME,GETDATE())<=30 ORDER BY ID DESC" ' Text2.Text
'
    AdodcSX.ConnectionString = gConnect '
    AdodcSX.Refresh
     TDBGrid1.columns(0).Width = 2500 '«Ã≈
TDBGrid.Col = MCOL
'TDBGrid.SetFocus





End Sub




'Private Sub TDBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
'    ' ?? ? ???st?? p?e? ?a epe?e??aste? t?? p??t? st??? (Index 0)
'    If ColIndex = 0 Then
'        Cancel = True ' ??????e? t?? epe?e??as?a
'    End If
'End Sub

Private Sub TDBGrid_DblClick()

'TDBGrid.Col = 0
'    Dim MT As String
'    MT = TDBGrid.Text
'
'     AdodcSX.RecordSource = "SELECT TOP 100 HME,SXOLIA,APOK,XRHSTHS,TYPOS,ID  FROM THLEFONA where THL='" + Trim(MT) + "' AND DATEDIFF(DAY,HME,GETDATE())<=30 ORDER BY ID DESC" ' Text2.Text
'
'    AdodcSX.ConnectionString = gConnect '
'    AdodcSX.Refresh
'     TDBGrid1.columns(0).Width = 2500 '«Ã≈
'

End Sub

Private Sub TDBGrid_HeadClick(ByVal ColIndex As Integer)
           'Adodc2.Refresh
           GRIDREFRESH
        '--------------------------------------------------------------------------------------=
        '<EhHeader>
        On Error GoTo TDBGrid_HeadClick_Err

        '</EhHeader>
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
         "in ADOMERCNEW.bohu6.getFilter " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.getFilter " & "at line " & Erl

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
         "in ADOMERCNEW.bohu6.cmdClearFilter_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.cmdClearFilter_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub TDBGrid_HeadClick(ByVal ColIndex As Integer)
'
'        '<EhHeader>
'        On Error GoTo TDBGrid_HeadClick_Err
'
'        '</EhHeader>
'
'        '
'        '  Dim strSort As String
'        '
'        '  Dim Col As TrueOleDBGrid80.Column
'        '
'        '
'        '
'        '' Loop through GroupColumns collection and construct
'        '
'        '' the sort string for the Sort property of the Recordset
'        '
'        '    For Each Col In TDBGrid.Columns
'        '
'        '        If strSort <> vbNullString Then
'        '
'        '            strSort = strSort & ", "
'        '
'        '        End If
'        '
'        '        strSort = strSort & "[" & Col.DataField & "]"
'        '
'        '    Next Col
'        '
'        '
'
'        'TDBGrid.HoldFields
'        Dim sumes(100) As String
'
'        Dim k          As Long
'
'        On Error Resume Next
'
'        'edo parakato exei latos 15/8/2009
'
'100     For k = 0 To Adodc2.Recordset.FIELDS.Count - 1
'110         sumes(k) = TDBGrid.Splits(0).columns(k).FooterText    '  = Format(SUMES(k), "######0.00")
'        Next
'
'120     If Adodc2.Recordset.Sort = "[" & TDBGrid.columns(ColIndex).DataField & "] asc" Then   ' strSort
'130         Adodc2.Recordset.Sort = "[" & TDBGrid.columns(ColIndex).DataField & "] desc"    ' strSort
'        Else
'140         Adodc2.Recordset.Sort = "[" & TDBGrid.columns(ColIndex).DataField & "] asc"    ' strSort
'        End If
'
'150     For k = 0 To Adodc2.Recordset.FIELDS.Count - 1
'160         TDBGrid.Splits(0).columns(k).FooterText = sumes(k)   '  = Format(SUMES(k), "######0.00")
'        Next
'
'        '    Adodc2.Recordset.Sort = "[" & TDBGrid.Columns(ColIndex).DataField & "] desc" ' strSort
'
'        '<EhFooter>
'        Exit Sub
'
'TDBGrid_HeadClick_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.pelat3.TDBGrid_HeadClick " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.TDBGrid_HeadClick " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>
'
'End Sub


'Private Sub cmdClearFilter_Click()
'
'        '<EhHeader>
'        On Error Resume Next  ' NEXT '  GoTo cmdClearFilter_Click_Err
'
'        '</EhHeader>
'
'        'Clears filter from grid
'
'100     For Each Col In TDBGrid.columns
'
'110         Col.FILTERTEXT = ""
'
'120     Next Col
'
'130     Adodc2.Recordset.Filter = adFilterNone
'
'        '<EhFooter>
'        Exit Sub
'
'cmdClearFilter_Click_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.bohu6.cmdClearFilter_Click " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu6.cmdClearFilter_Click " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>
'
'End Sub
'




Private Sub TDBGrid_FilterChange()


        '<EhHeader>
        On Error GoTo TDBGrid_FilterChange_Err

        '</EhHeader>

        'Gets called when an action is performed on the filter bar

        On Error Resume Next 'GoTo errHandler

100     Set ColS = TDBGrid.columns

        Dim c As Integer

110     c = TDBGrid.Col

120     TDBGrid.HoldFields
        Dim CC As String
        CC = getFilter()
130     Adodc2.Recordset.Filter = CC '  getFilter()

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

Private Sub TDBGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' TDBGrid1.Col = 1
   ' Dim MT As String
    'MT = TDBGrid1.Text
   
End Sub

Private Sub TDBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'
'Dim MCOL As Integer
'MCOL = tdbgrid1.Col
'
'TDBGrid.Refresh
'
'TDBGrid.Col = 0
'    Dim MT As String
'    MT = TDBGrid.Text
'
'     AdodcSX.RecordSource = "SELECT TOP 100 HME,SXOLIA,APOK,XRHSTHS,TYPOS,ID  FROM THLEFONA where THL='" + Trim(MT) + "' AND DATEDIFF(DAY,HME,GETDATE())<=30 ORDER BY ID DESC" ' Text2.Text
'
'    AdodcSX.ConnectionString = gConnect '
'    AdodcSX.Refresh
'    If MCOL < 4 Then
'          TDBGrid.Col = 4 'MCOL
'    Else
'         tdbgrid1.Col = MCOL
'
'    End If
    
End Sub

Private Sub XRONIKA_Change()
    GRIDREFRESH
End Sub

Private Sub XRONIKA_Click()
 GRIDREFRESH
End Sub
