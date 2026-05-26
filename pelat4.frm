VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form pelat4 
   BackColor       =   &H00FF0000&
   Caption         =   "Καρτέλα"
   ClientHeight    =   12525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17880
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12525
   ScaleWidth      =   17880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdDeltia 
      Caption         =   "Δ.Α. εκκρεμή"
      Height          =   405
      Left            =   11685
      TabIndex        =   55
      Top             =   6120
      Width           =   1530
   End
   Begin VB.CheckBox cXronies 
      Caption         =   "5 τελ.χρονιές"
      Height          =   315
      Left            =   5040
      TabIndex        =   54
      Top             =   6120
      Width           =   1815
   End
   Begin TrueOleDBGrid80.TDBGrid tdbgrid1 
      Bindings        =   "pelat4.frx":0000
      Height          =   3855
      Left            =   120
      TabIndex        =   53
      Top             =   2280
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   6800
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
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
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
      AllowUpdate     =   0   'False
      ColumnFooters   =   -1  'True
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
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
   Begin VB.CommandButton cmdPreview 
      Caption         =   "ΠΡΟΕΠΙΣΚΟΠΗΣΗ"
      Height          =   360
      Left            =   6480
      TabIndex        =   52
      Top             =   840
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3720
      Top             =   9960
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VB.CheckBox CH_SYGK 
      BackColor       =   &H00FF0000&
      Caption         =   "Συγκ.Καρτέλα"
      ForeColor       =   &H8000000E&
      Height          =   340
      Left            =   240
      TabIndex        =   51
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Αποθηκευση"
      Height          =   255
      Left            =   14040
      TabIndex        =   47
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox kentroAdyn 
      Height          =   315
      ItemData        =   "pelat4.frx":0015
      Left            =   120
      List            =   "pelat4.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CMDEXIT 
      BackColor       =   &H80000010&
      Caption         =   "Εξοδος"
      Height          =   495
      Left            =   9885
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton Command12 
      Caption         =   "PDF ΤΙΜΟΛΟΓΙΟ "
      Enabled         =   0   'False
      Height          =   405
      Left            =   6720
      TabIndex        =   40
      Top             =   10200
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.ComboBox XRONIES 
      Height          =   315
      Left            =   6960
      TabIndex        =   39
      Text            =   "XRONIES"
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton THLEFONA 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ληφθέντα τηλέφωνα"
      Height          =   405
      Left            =   11685
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   8040
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.CommandButton cmdKARTELAseEXCEL 
      Caption         =   "Καρτέλλα σε Excel"
      Height          =   405
      Left            =   11685
      TabIndex        =   37
      Top             =   5640
      Width           =   1530
   End
   Begin VB.CheckBox Check2 
      Caption         =   "ΠΑΛΙΕΣ ΧΡ"
      Height          =   255
      Left            =   9465
      TabIndex        =   36
      Top             =   6150
      Width           =   1155
   End
   Begin VB.TextBox timtext3 
      Height          =   240
      Left            =   9120
      TabIndex        =   33
      Text            =   "Text4"
      Top             =   6150
      Width           =   195
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Eκτυπώσεις Καρτελλών"
      Height          =   405
      Left            =   13890
      TabIndex        =   31
      Top             =   5640
      Width           =   1530
   End
   Begin VB.CommandButton cmdPRINTMEEIDH 
      Caption         =   "Εκτύπωση με είδη"
      Height          =   495
      Left            =   5910
      TabIndex        =   30
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdGraph 
      Caption         =   "Γραφική Παράσταση"
      Height          =   495
      Left            =   7875
      TabIndex        =   29
      Top             =   6480
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3900
      TabIndex        =   27
      Top             =   8940
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2460
      TabIndex        =   25
      Top             =   8940
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DbGrid2 
      Bindings        =   "pelat4.frx":0032
      Height          =   2175
      Left            =   8160
      TabIndex        =   14
      Top             =   105
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "EPO"
         Caption         =   "ONOMATEΠΩΝΥΜΟ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "THL"
         Caption         =   "ΤΗΛEΦΩΝΟ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "DIE"
         Caption         =   "ΔΙΕΥΘΥΝΣΗ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   3
         Size            =   184
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3480
      Top             =   10560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Εκτύπωση Παραστατικού"
      Height          =   405
      Left            =   6720
      TabIndex        =   24
      Top             =   9960
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.ListBox List1 
      Height          =   1185
      Left            =   11685
      Style           =   1  'Checkbox
      TabIndex        =   23
      Top             =   6600
      Width           =   3735
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   270
      Left            =   5385
      TabIndex        =   18
      Top             =   135
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   476
      _Version        =   393216
      Format          =   154468353
      CurrentDate     =   38854
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Διόρθωση"
      Height          =   345
      Left            =   6480
      TabIndex        =   17
      Top             =   1245
      Width           =   1470
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ΠΕΡΥΣΙ"
      Height          =   255
      Left            =   10680
      TabIndex        =   16
      Top             =   6150
      Width           =   900
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   3360
      Top             =   10200
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
      Bindings        =   "pelat4.frx":0046
      Height          =   3852
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   6800
      _Version        =   393216
      BackColor       =   4227327
      Cols            =   10
      FixedCols       =   0
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin MSAdodcLib.Adodc egg 
      Height          =   330
      Left            =   0
      Top             =   5490
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   495
      Left            =   3480
      Top             =   10080
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.CommandButton Command5 
      Caption         =   "Παλιά χρονιά"
      Height          =   495
      Left            =   15720
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdXREOPISTOSEIS 
      Caption         =   "&Χρεωπιστώσεις"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3915
      TabIndex        =   9
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdEIDHSYNOLIKA 
      Caption         =   "Εί&δη συνολικά"
      Enabled         =   0   'False
      Height          =   495
      Left            =   75
      TabIndex        =   8
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdPRINTKARTELA 
      Caption         =   "Εκ&τύπωση"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1995
      TabIndex        =   7
      Top             =   6480
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Height          =   495
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   1485
      Begin VB.ComboBox PELPROM 
         Height          =   315
         Left            =   0
         TabIndex        =   34
         Top             =   120
         Width           =   1455
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   270
      Left            =   5370
      TabIndex        =   19
      Top             =   495
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   476
      _Version        =   393216
      Format          =   308281345
      CurrentDate     =   38854
   End
   Begin MSChart20Lib.MSChart graph1 
      Height          =   3855
      Left            =   0
      OleObjectBlob   =   "pelat4.frx":0058
      TabIndex        =   28
      Top             =   2280
      Width           =   7215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid DBGRID22 
      Bindings        =   "pelat4.frx":1E61
      Height          =   1632
      Left            =   12
      TabIndex        =   32
      Top             =   7080
      Width           =   11568
      _ExtentX        =   20399
      _ExtentY        =   2884
      _Version        =   393216
      BackColor       =   4227327
      Rows            =   300
      Cols            =   10
      FixedCols       =   0
      ScrollBars      =   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin ADOMERCNEW.msEditor2 msEditor1 
      Height          =   3255
      Left            =   11685
      TabIndex        =   45
      Top             =   2280
      Width           =   3735
      _extentx        =   6588
      _extenty        =   5741
      backcolor       =   12648447
      font            =   "pelat4.frx":1E75
      open_visible    =   0   'False
      save_visible    =   0   'False
      backcolor       =   12648447
      enabled         =   0   'False
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
      Bindings        =   "pelat4.frx":1EA1
      Height          =   1452
      Left            =   120
      TabIndex        =   56
      Top             =   7200
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   2566
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
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
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
      AllowUpdate     =   0   'False
      ColumnFooters   =   -1  'True
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
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
   Begin VB.Label lbInformation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F3 :Διόρθωση Εγγραφής  Διπλό Κλίκ: Διαγραφή εγγραφής"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   180
      TabIndex        =   50
      Top             =   6210
      Width           =   4290
   End
   Begin VB.Label ypolposo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14520
      TabIndex        =   49
      Top             =   9000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label ypolmhx 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14520
      TabIndex        =   48
      Top             =   8520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Φιλτρο Παραστατικών"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   12120
      TabIndex        =   46
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label symf 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Υπολ.Ποσό"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   1
      Left            =   13560
      TabIndex        =   43
      Top             =   9000
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label symf 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Υπόλ.μηχαν"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   0
      Left            =   13560
      TabIndex        =   42
      Top             =   8520
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label ffbuff 
      Height          =   165
      Left            =   5055
      TabIndex        =   35
      Top             =   1380
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Φίλτρο"
      ForeColor       =   &H80000014&
      Height          =   225
      Left            =   1635
      TabIndex        =   26
      Top             =   8970
      Width           =   825
   End
   Begin VB.Label APOALLO 
      Caption         =   "0"
      Height          =   300
      Left            =   6540
      TabIndex        =   22
      Top             =   8310
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "έως"
      ForeColor       =   &H80000005&
      Height          =   240
      Left            =   4860
      TabIndex        =   21
      Top             =   525
      Width           =   465
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "από"
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   4860
      TabIndex        =   20
      Top             =   165
      Width           =   525
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   120
      Top             =   1650
      Width           =   9975
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Caption         =   ".."
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   1770
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   ".."
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1770
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   ".."
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   1770
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A.Φ.Μ."
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Επωνυμία"
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
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Κωδικός"
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
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "pelat4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim f_resizeCount
Dim F_PIST As Integer

'Private El As New FORMA
Dim PROTH, fflag
Attribute fflag.VB_VarUserMemId = 1073938432

Dim F_EKTYPOTHS As String '= FINDPARAMETROI(1, "PAR1", "F_EKTYPOTHS", "..", "Εκτυπωτής παραστατικών(με .. προεπιλεγμένος)") '

Dim PROTH_FORA
Attribute PROTH_FORA.VB_VarUserMemId = 1073938434
 Dim F_PLATOS_FORMAS As Long, F_YCOS_FORMAS As Long
 
 
 Dim F_NEA_KRTELLA As Integer
Dim LAST_TIM As Long

Dim f_direct  ' τυπωνει απευθειας χωρις να ρωτα εκτυπωτη =1
Attribute f_direct.VB_VarUserMemId = 1073938436

Dim PELPROM_STRING As String

Dim f_telephone    As Integer

Dim f_text1(3)
Attribute f_text1.VB_VarUserMemId = 1073938439

Dim f_buff As String    'πελατης ή προμηθευτής

Dim F_CanDelete    ' = FindParametroi(1,"BOHU11", "F_CanDelete", "1", "μπορω να σβήνω κινήσεις")
Attribute F_CanDelete.VB_VarUserMemId = 1073938441

Dim f_PALIAXRONIA   ' = Val(FindParametroi(1,"PAR1", "f_PALIAXRONIA", "2005", "Σειρές τιμ.Πώλησης"))
Attribute f_PALIAXRONIA.VB_VarUserMemId = 1073938442

Dim f_PROPERSI    ' = FindParametroi(1,"PAR1", "F_PROPERSI", "2005", "Γ.19 ΒΑΣΗ ΠΡΟΠΕΡΣΙΝΗΣ ΧΡΟΝΙΑΣ")
Attribute f_PROPERSI.VB_VarUserMemId = 1073938443

Dim f_resize
Attribute f_resize.VB_VarUserMemId = 1073938444

Dim f_KENTRA_ADYNATISMATOS As Integer
Dim f_2000MORE_ROWS As Integer  '
Dim F_MONOENERGOS As Integer



Dim ButtonFlag As Integer

Dim Prospective As New TrueDBGrid80.Style

Dim Distributors As New TrueDBGrid80.Style
Dim RSClone As ADODB.Recordset







Dim f_msynt                As String  ' filtro gia ta kentra adynatismatos
' Dim F_FIRSTRESIZE As Integer

Dim f_LabAFM  As String

Dim f_pedioAFM As String  ' = FINDPARAMETROI(1, "PELAT2", "f_pedioAFM", "CH1", "Ποιό πεδίο θα αναζητάται μαζί με το ΑΦΜ")

Dim FCONNECTIONPALIA       As New ADODB.Connection

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)
Attribute fh.VB_VarUserMemId = 1073938445


'-------------------------------------------
Private Type ControlInfo_type
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
  End Type
  Dim ControlInfos() As ControlInfo_type
'--------------------------------------------------




Public Property Let resizing(strTable As Integer)

        '<EhHeader>
        On Error GoTo resizing_Err

        '</EhHeader>
100     If Not IsNumeric(strTable) Then
110         Err.Raise 30000, "Data Table Editor", "Table name appears to be invalid."

            Exit Property

        End If

120     f_resize = strTable
130     Me.Visible = True
140     Me.ZOrder 0

        '<EhFooter>
        Exit Property

resizing_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.resizing " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.resizing " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Property


 
















Private Sub CH_SYGK_Click()
   If CH_SYGK.Value = vbChecked Then
   
        If Check1.Value = vbChecked Or Check2.Value = vbChecked Then
           MsgBox "Συγκεντρωτική καρτέλλα μόνο στην τρέχουσα χρήση"
           Check1.Value = vbUnchecked
           Check1.Value = vbUnchecked
        End If
        
        cmdXREOPISTOSEIS.Enabled = False
    
   
   End If
End Sub

Private Sub Check1_Click()

        '<EhHeader>
        On Error GoTo Check1_Click_Err

        '</EhHeader>
100     If Check1.Value = vbChecked Then
110         DBGrid2_RowColChange 0, 0
        End If
    
        '<EhFooter>
        Exit Sub

Check1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Check1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Check1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub CmdDeltia_Click()

On Error Resume Next
Dim R As New ADODB.Recordset
DBGRID22.Clear
List1.Clear

    R.Open "select METAF,ATIM,CONVERT(CHAR(10),HME,103) AS [Ημερομηνία],ID_NUM from TIM where ISNULL(AKYROMENO,0)=0 AND left(ATIM,1) in (SELECT EIDOS FROM PARASTAT WHERE STADIO=3 AND POL=1) and EIDOS='" + data1.Recordset("EIDOS") + "' AND KPE='" + data1.Recordset("kod") + "' and ART IS null ORDER BY HME;", Gdb, adOpenForwardOnly, adLockReadOnly

     'If Not r.EOF Then

      '   DBGRID22.AddItem " "
       '  DBGRID22.TextMatrix(DBGRID22.ROWS - 1, 0) = "Σκοπός:"
        ' DBGRID22.TextMatrix(DBGRID22.ROWS - 1, 1) = r("SKOPOS")

    ' End If
    
    
  On Error GoTo LATOS
     Dim k As Integer
     k = 1
     Do While Not R.EOF
        ' DBGRID22.TextMatrix(k, 3) = R!ATIM
        ' DBGRID22.TextMatrix(k, 2) = CNull(R(2))
         List1.AddItem R!ATIM + " " + CNull(R(2))
     
       R.MoveNext
       k = k + 1
     Loop

R.Close

Exit Sub

LATOS:
MsgBox "ΛΑΘΟΣ ΣΤΟ :" + R!ATIM + "   " + str(k)
Resume Next


End Sub




Sub PAR2SHOWING(ByVal k As Integer)

        '<EhHeader>
        On Error GoTo PAR2SHOWING_Err

        '</EhHeader>

        Dim FF As PAR2

100
110     Set FF = New PAR2

120

130    'DBGrid1 .Col = 1
140     FF.DTPicker1 = CDate(DBGRID22.TextMatrix(k, 2))
150     FF.DTPicker2 = CDate(DBGRID22.TextMatrix(k, 2))

 'DBGrid1.Col = 0
  
160   FF.Text1 = DBGRID22.TextMatrix(k, 3)    '  Trim(Left(DBGrid1.Text, 9))

FF.CallFromOther.Caption = "0"

        'GIA OLA TA PARASTATIKA
170     FF.Combo2.Text = "*"
        
       FF.Command1_Click
       
      FF.SHOW_TIMOL

180     FF.SHOW

 FF.TIM.Move 0
        FF.SHOW_TIMOL
190     FF.bohuitiko.SetFocus

200     DoEvents

210     FF.Command1.SetFocus
220    ' PAR1.LASTTIMOL.Refresh

        'On Error GoTo 0

230     '   SendKeys "{enter}"    'ENTER  ~

        '<EhFooter>
        Exit Sub

PAR2SHOWING_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.par1.PAR2SHOWING " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.par1.PAR2SHOWING " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub




























Private Sub cXronies_Click()


'    If cXronies.Value = vbChecked Then
'        Dim SQL11 As String
'        SQL11 = "SELECT * INTO " + fFILE + " FROM EGGTIM  where " + filtroPEL.Caption + " HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "'  AND APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' "
'
'        SQL11 = SQL11 + " UNION SELECT *  FROM " + f_PALIAXRONIA + ".dbo.EGGTIM  where " + filtroPEL.Caption + " HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "'  AND APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' "
'        Dim mPROP As String
'        mPROP = FINDPARAMETROI(1, "PAR1", "F_PROPERSI", "", "Γ.19 ΒΑΣΗ ΠΡΟΠΕΡΣΙΝΗΣ ΧΡΟΝΙΑΣ")
'        SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        Dim mPROP2 As String
'
'
'
'        mPROP3 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI2", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΙΝ 3 XPONIA")
'        If Len(mPROP3) > 0 Then
'           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP3 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        End If
'
'        Dim mPROP4 As String
'        mPROP4 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI3", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΙΝ 4 XPONIA")
'        If Len(mPROP4) > 0 Then
'           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP4 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        End If
'
'
'          Dim mPROP5 As String
'        mPROP5 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI5", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΟ 5ΕΤΙΑΣ")
'        If Len(mPROP5) > 0 Then
'           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP5 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        End If
'
'          Dim mPROP6 As String
'        mPROP6 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI6", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΟ 6ΕΤΙΑΣ")
'        If Len(mPROP6) > 0 Then
'           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP6 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        End If
'
'           Dim mPROP7 As String
'        mPROP7 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI7", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΟ 7ΕΤΙΑΣ")
'        If Len(mPROP7) > 0 Then
'           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP7 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        End If
'
'           Dim mPROP8 As String
'        mPROP8 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI8", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΟ 8ΕΤΙΑΣ")
'        If Len(mPROP8) > 0 Then
'           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP8 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        End If
'
'        Dim mPROP9 As String
'        mPROP9 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI9", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΟ 9ΕΤΙΑΣ")
'        If Len(mPROP9) > 0 Then
'           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP9 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        End If
'
'
'        Dim mPROP10 As String
'        mPROP10 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI10", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΟ 10ΕΤΙΑΣ")
'        If Len(mPROP10) > 0 Then
'           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP10 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        End If
'
'
'
'        Gdb.Execute SQL11
'    Else
'
'        On Error GoTo exitcode
'        GTGDB.Execute "SELECT * INTO " + fFILE + " FROM EGGTIM  where " + filtroPEL.Caption + " HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME<='" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "'  AND APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '  + "  Order by convert(char(10),HME,2),ATIM "
'
'
'
'
'     End If
'
'

End Sub

Private Sub cmdPreview_Click()

  
neakartela 0


'  TDBGrid1.columns(3).Width = 100
'  TDBGrid1.columns(4).Width = 100
 
  
  
    
             '  Adodc1.RecordSource = "SELECT convert(char(10),HME,103) as [HME],ATIM,AIT,XREOSI AS [ΧΡΕΟΣΗ],PISTOSI AS [ΠΙΣΤΩΣΗ],ROUND( " + myp2 + "+ sum(" + mxreo + ") over(order by  HME,ID ) ,2)    as [ΥΠΟΛΟΙΠΟ]  FROM EGG WHERE  EIDOS='" + gBuff + "' AND KOD='" + Data1.Recordset("KOD") + "' and   HME>='" + Format(dtpicker1.Value, "MM/DD/YYYY") + "' AND HME<='" + Format(dtpicker2.Value, "MM/DD/YYYY") + "'"
               'Adodc1.Recordset.MoveLast
             '  Adodc1.Refresh
               Dim COMPANY As String
               COMPANY = GGET_CVALUE("SELECT TOP 1 PELONO FROM MEM")
  
  With tdbgrid1.PrintInfo

        ' Set the page header

        .PageHeaderFont.Italic = True

        .PageHeader = "\t  " + COMPANY + ":      " + IIf(gBuff = "e", "Πελατης:", "Προμηθευτής: ") + data1.Recordset("EPO")

        .SettingsOrientation = 1 '1 - Portrait or 2 - Landscape orientation.

 


        ' Column headers will be on every page

        .RepeatColumnHeaders = True

        

        ' Display page numbers (centered)


      '

        .PageFooter = "\tΣελίδα: \p"

        ' Invoke Print Preview

        .PrintPreview

    End With


End Sub

Private Sub CMDEXIT_Click()
      gApoMenu = True  ' αν τυχον μπηκα στην διορθωση να αλλαξω το ΦΛΑΓ
        '<EhHeader>
        On Error GoTo CMDEXIT_Click_Err

        '</EhHeader>
        Dim x As Integer

100     x = Val(APOALLO)

        'ME.HIDE
110     Unload Me

120     If x = 0 Then
130         If gApoMenu = True Then
140             SendKeys "%"
150             SendKeys "{ENTER}+{DOWN}+{DOWN}+{DOWN}"
            End If

        Else
            ' pelat2.SHOW
160         Unload Me

        End If

        '<EhFooter>
        Exit Sub

CMDEXIT_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command10_Click()

        '<EhHeader>
        On Error GoTo Command10_Click_Err

        '</EhHeader>
        Dim a1, a2

100     a1 = InputBox("Από Κωδικό")
110     a2 = InputBox("Εως Κωδικό")

        'Data1.RecordSource = "select EPO,THL,DIE,* from PEL where EIDOS='e' and KOD >= '" + a1 + "' and KOD<='" + a2 + " ';"
        'Data1.Refresh

120     data1.RecordSource = "select EPO,THL,DIE,* from PEL where EIDOS='e' and KOD >= '" + a1 + "' and KOD<='" + a2 + " ';"
130     data1.Refresh

        f_direct = 1

        Dim ANS As Integer

140     ANS = MsgBox("Είσαι σίγουρος;", vbYesNo)

150     If ANS = vbNo Then

            Exit Sub

        End If

        '   f_direct = 1
160     utDialog.Command3.Enabled = False
170     utDialog.Command4.Enabled = False
180     utDialog.SHOW 1
190     PROTH_FORA = False

200     data1.Recordset.MoveFirst

        On Error Resume Next

       '  f_direct = 0 ' τυπωνει απευθειας χωρις να ρωτα εκτυπωτη =1
210     Do While Not data1.Recordset.EOF

220         MILSEC 2500
230         DBGrid2_RowColChange 0, 0

240         cmdPRINTKARTELA_Click
250         data1.Recordset.MoveNext
260         MILSEC 2500

        Loop

        '<EhFooter>
        Exit Sub

Command10_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Command10_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Command10_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdKARTELAseEXCEL_Click()

   ' On Error Resume Next

'Dim oExcelApp As Object, wb As Object
'
'Set oExcelApp = CreateObject("Excel.Application")
'
'Set wb = oExcelApp.Workbooks.Open("C:\File.xls")

        '--------- excel  kartella --------------------------------
        '<EhHeader>
        On Error GoTo cmdKARTELAseEXCEL_Click_Err

        '</EhHeader>
'
        Dim excel  As Object:
        Set excel = CreateObject("Excel.Application")

        Dim workbook As Object: Set workbook = excel.Workbooks.Add '  As New excel.workbook

        Dim myXL As Object: Set myXL = workbook.ActiveSheet '  '    As excel.Worksheet

100     Set excel = New excel.Application
        ' Excel.Visible = True
102     Set workbook = excel.Workbooks.Add

       ' On Error Resume Next

104     workbook.Activate

106     Set myXL = workbook.ActiveSheet
        
        
        
        
108 With myXL.range("A1:G1")
110     .HorizontalAlignment = xlCenter
112     .VerticalAlignment = xlBottom
114     .WrapText = False
116     .Orientation = 0
118     .AddIndent = False
120     .IndentLevel = 0
122     .shrinkToFit = False
124     .ReadingOrder = xlContext
126     .MergeCells = False
    End With
    
128  myXL.columns(5).HorizontalAlignment = xlRight
130  myXL.columns(6).HorizontalAlignment = xlRight
132   myXL.columns(7).HorizontalAlignment = xlRight
    
134  myXL.columns(3).HorizontalAlignment = xlLeft
    
    
136 myXL.range("a1", "g1").Merge
    
138    With myXL.range("A1:G1").Interior
140     .Pattern = xlSolid
142     .PatternColorIndex = xlAutomatic
144     .ThemeColor = xlThemeColorDark2
146     .TintAndShade = -9.99786370433668E-02
148     .PatternTintAndShade = 0
    End With
        
        
150    '1O  myXL.rows(0).HorizontalAlignment = xlRight
        
   
   

152 myXL.range("e8:G35").HorizontalAlignment = xlRight
    

    'range("A1:G1").Select
154 With myXL.range("A1:G1")
156     .HorizontalAlignment = xlCenter
158     .VerticalAlignment = xlBottom
160     .WrapText = False
162     .Orientation = 0
164     .AddIndent = False
166     .IndentLevel = 0
168     .shrinkToFit = False
170     .ReadingOrder = xlContext
172     .MergeCells = True
    End With
        
       ' range("A3:G5").Select
174 myXL.range("A3:G5").Borders(xlDiagonalDown).LineStyle = xlNone
176 myXL.range("A3:G5").Borders(xlDiagonalUp).LineStyle = xlNone
    
178 With myXL.range("A3:G5").Borders(xlEdgeLeft)
180     .LineStyle = xlContinuous
182     .ColorIndex = 0
184     .TintAndShade = 0
186     .Weight = xlMedium
    End With
    
188 With myXL.range("A3:G5").Borders(xlEdgeTop)
190     .LineStyle = xlContinuous
192     .ColorIndex = 0
194     .TintAndShade = 0
196     .Weight = xlMedium
    End With
198 With myXL.range("A3:G5").Borders(xlEdgeBottom)
200     .LineStyle = xlContinuous
202     .ColorIndex = 0
204     .TintAndShade = 0
206     .Weight = xlMedium
    End With
208 With myXL.range("A3:G5").Borders(xlEdgeRight)
210     .LineStyle = xlContinuous
212     .ColorIndex = 0
214     .TintAndShade = 0
216     .Weight = xlMedium
    End With
218 myXL.range("A3:G5").Borders(xlInsideVertical).LineStyle = xlNone
220 myXL.range("A3:G5").Borders(xlInsideHorizontal).LineStyle = xlNone
222 With myXL.range("A3:G5").Interior
224     .Pattern = xlSolid
226     .PatternColorIndex = xlAutomatic
228     .ThemeColor = xlThemeColorAccent3
230     .TintAndShade = 0.599993896298105
232     .PatternTintAndShade = 0
    End With
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
    '        myXL.range("a1", "g1").Merge

        Dim sxre, spis, DB As Database, dbac As Database

        Dim MKOD, mEpo

       ' On Error Resume Next

       ' On Error GoTo cmdKARTELAseEXCEL_Click_Err

        Dim k, YPARXEI(5)

        Dim T As String

234     MKOD = data1.Recordset("kod")
236     sxre = 0
238     spis = 0

        Dim R

240     egg.RecordSource = "SELECT * FROM KARTELA"
242     Set R = egg.Recordset
        ' Printer.FontName = "COURIER"

        Dim ra As New ADODB.Recordset

        'ra.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
        '        k = 1
        '        With myXL
        '            .cells(1, 1) = ra("pelono")
        '            '.cells(, 1) = ra("pelepa")
        '            .cells(1, 3) = Now
        '            .cells(2, 1) = "Kαρτέλλα Πελάτη " ' EPIKEF
        '        End With
        '
        '  On Error GoTo Command11_Click_Err
        '
        '
        '    '----------------------- ΕΠΙΚΕΦΑΛΙΔΑ ----------------------------
        '   ' For k = 0 To n - 1
        '    '    myXL.cells(5, k + 1) = F_T(k).Name
        '   ' Next
        '
        '' myXL.cells(5, 1) = "Hμερ/νία"
        '' myXL.cells(6, 2) = "Αριθμός"
        '' myXL.cells(5, 3) = "Hμερ/νία"
        '' myXL.cells(5, 2) = "Hμερ/νία"
        '' myXL.cells(5, 4) = "Χρέωση"
        '' myXL.cells(5, 5) = "Πίστωση"
        '' myXL.cells(5, 6) = "Υπόλοιπο"
        '' myXL.cells(5, 7) = "Χρήστης"
        '
        '
        '    myXL.cells(3, 1) = Data1.Recordset("epo") + "   Κωδ." + Data1.Recordset("kod")
        '    myXL.cells(4, 1) = Data1.Recordset("die")
        '   myXL.cells(5, 1) = "Tηλέφωνο " + Data1.Recordset("thl")

244     myXL.rows(6).Font.Size = 14
246     myXL.rows(6).Font.fontStyle = 12
        
        
       ' On Error Resume Next

        Dim mait As String
        
248      For k = 1 To 7
250          myXL.cells(6, k) = "'" + MSFlexGrid1.TextMatrix(0, k)     ' , "DD/MM/YYYY") ' 01/10/2005
         Next
252      myXL.cells(6, 3) = "Αριθ.Παρ"
254      myXL.cells(6, 2) = "  "




256     For k = 1 To MSFlexGrid1.rows - 1
258         myXL.rows(k).Font.Size = 10
260         sxre = sxre + gVal(MSFlexGrid1.TextMatrix(k, 5)) ' R(5) '"Χρέωση")
262         spis = spis + gVal(MSFlexGrid1.TextMatrix(k, 6)) 'R(6) '"Πίστωση")

264         myXL.cells(6 + k, 1) = "'" + MSFlexGrid1.TextMatrix(k, 0)    ' , "DD/MM/YYYY") ' 01/10/2005
266         myXL.cells(6 + k, 2) = LTrim(tostC(MSFlexGrid1.TextMatrix(k, 1), 10))    ' ;  ' a14609B

268         myXL.cells(6 + k, 3) = LTrim(tostC(MSFlexGrid1.TextMatrix(k, 3), 11))    ' " " '  1600 ar.par

270         mait = MSFlexGrid1.TextMatrix(k, 4)
272         myXL.cells(6 + k, 4) = Left(mait + Space(25), 25)  ' da pelath

274         If k = 0 Then
276             myXL.cells(6 + k, 5) = TOST(Val(Replace(MSFlexGrid1.TextMatrix(k, 5), ",", ".")), 10)  '  xreosi
278             myXL.cells(6 + k, 6) = TOST(Val(Replace(MSFlexGrid1.TextMatrix(k, 6), ",", ".")), 10)  '  pistosi
280             myXL.cells(6 + k, 7) = TOST(Val(Replace(MSFlexGrid1.TextMatrix(k, 7), ",", ".")), 10) '  ypoloipo
            Else
282             myXL.cells(6 + k, 5) = TOST(Val(Replace(MSFlexGrid1.TextMatrix(k, 5), ",", ".")), 10) '   xreosi
284             myXL.cells(6 + k, 6) = TOST(Val(Replace(MSFlexGrid1.TextMatrix(k, 6), ",", ".")), 10) ' pistosi
286             myXL.cells(6 + k, 7) = TOST(Val(Replace(MSFlexGrid1.TextMatrix(k, 7), ",", ".")), 10) 'ypoloipo
            End If

          ' myXL.cells(6 + K, 8) = TOST(Val(MSFlexGrid1.TextMatrix(K, 8)), 12)   ' Format(Val(MSFlexGrid1.TextMatrix(K, 8)), "##") 'ΧΡΗΣΤΗς
        Next

       ' On Error GoTo cmdKARTELAseEXCEL_Click_Err

         ' myXL.columns("B:K").Select
288  myXL.columns.AutoFit
    '
    ' myXL.columns("E:K").Select
    'myXL.columns.HorizontalAlignment = xlRight
    '
    ' myXL.columns("A:D").Select
    'myXL.columns.HorizontalAlignment = xlLeft
    '
    '
    ' myXL.columns("E:e").Select
    'myXL.columns("e").HorizontalAlignment = xlRight
    '
    '
    ' myXL.columns("f:f").Select
    'myXL.columns("f").HorizontalAlignment = xlRight
    '
    ' myXL.columns("E:G").Select
    'myXL.columns.HorizontalAlignment = xlRight

        '----------------------EPIKEFALIDES STO TELOS GIA NA  MHN XALANE TO AUTO FIT
290     ra.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
292     k = 1

294     With myXL
296         .cells(1, 1) = ra("pelono")

298         myXL.rows(1).Font.Size = 14
300         myXL.rows(1).Font.fontStyle = 12

            '.cells(, 1) = ra("pelepa")
302         .cells(2, 4) = Now

304         .cells(2, 1) = "Kαρτέλλα Πελάτη "    ' EPIKEF
        End With

     '   On Error GoTo cmdKARTELAseEXCEL_Click_Err

306     myXL.cells(3, 1) = data1.Recordset("epo") + "   Κωδ." + data1.Recordset("kod")
308     myXL.rows(3).Font.Size = 14
310     myXL.rows(3).Font.fontStyle = 12

312     myXL.cells(4, 1) = data1.Recordset("die")
314     myXL.cells(5, 1) = "Tηλέφωνο " + data1.Recordset("thl")


    '         myXL.cells(5, 1) = "Hμερ/νία"
    '         myXL.cells(5, 2) = "Αριθμός"
    '         myXL.cells(5, 3) = "Hμερ/νία"
    '         myXL.cells(5, 2) = "Hμερ/νία"
    '         myXL.cells(5, 4) = "Χρέωση"
    '         myXL.cells(5, 5) = "Πίστωση"
    '         myXL.cells(5, 6) = "Υπόλοιπο"



'316 myXL.columns("5:8").Select



'318 myXL.columns.HorizontalAlignment = xlRight

'320 myXL.columns("3:3").Select
'322 myXL.columns.HorizontalAlignment = xlLeft




    '
    '
    '



324 ' myXL.rows("2:6").Select


326 ' myXL.rows.HorizontalAlignment = xlLeft
    '
    '
    ' On Error GoTo 0


    '    range("D1").Select
    '    selection.Cut Destination:=range("E2")
    '    range("E2").Select
    '
    '
    '    columns("E:E").columnWidth = 10.14
    '    selection.Cut Destination:=range("D2")
    '    columns("D:D").Select
    'myXL.range("A1:G1").Select


    
 
'328        myXL.cells(1, 1).BackColor = vbRed
             
     
   
       ' .Pattern = xlSolid
       ' .PatternColorIndex = xlAutomatic
       ' .ThemeColor = xlThemeColorDark2
      '  .TintAndShade = -9.99786370433668E-02
      '  .PatternTintAndShade = 0






330     MDIForm1.MousePointer = vbNormal


       ' On Error Resume Next



332 myXL.rows("6:600").HorizontalAlignment = xlRight

    'myXL.ROWS.HorizontalAlignment = xlRight

    'myXL.ROWS("1:1").Select
334 myXL.rows(1).HorizontalAlignment = xlCenter


    'myXL.ROWS("2:6").Select
336 myXL.rows("2:6").HorizontalAlignment = xlLeft



    Dim mRange As String
338 mRange = "a6:g" + Format(MSFlexGrid1.rows - 1 + 6, "000")

       ' myXL.range(mRange).Select
340    myXL.range(mRange).Borders(xlDiagonalDown).LineStyle = xlNone
342     myXL.range(mRange).Borders(xlDiagonalUp).LineStyle = xlNone
344    With myXL.range(mRange).Borders(xlEdgeLeft)
346        .LineStyle = xlContinuous
348        .ColorIndex = 0
350         .TintAndShade = 0
352        .Weight = xlThin
        End With
354    With myXL.range(mRange).Borders(xlEdgeTop)
356        .LineStyle = xlContinuous
358        .ColorIndex = 0
360        .TintAndShade = 0
362        .Weight = xlThin
       End With
364    With myXL.range(mRange).Borders(xlEdgeBottom)
366        .LineStyle = xlContinuous
368        .ColorIndex = 0
370        .TintAndShade = 0
372        .Weight = xlThin
       End With
374    With myXL.range(mRange).Borders(xlEdgeRight)
376        .LineStyle = xlContinuous
378        .ColorIndex = 0
380        .TintAndShade = 0
382        .Weight = xlThin
       End With
384 With myXL.range(mRange).Borders(xlInsideVertical)
386     .LineStyle = xlContinuous
388     .ColorIndex = 0
390     .TintAndShade = 0
392     .Weight = xlThin
    End With
394 With myXL.range(mRange).Borders(xlInsideHorizontal)
396     .LineStyle = xlContinuous
398     .ColorIndex = 0
400     .TintAndShade = 0
402     .Weight = xlThin
    End With
404 myXL.range("A6:G6").Select
406 With selection.Interior
408     .Pattern = xlSolid
410     .PatternColorIndex = xlAutomatic
412     .Color = 65535
414     .TintAndShade = 0
416     .PatternTintAndShade = 0
    End With
    
    
    



418  Kill "C:\EKTYP.XLS"

420     DoEvents

422     myXL.SaveAs "C:\EKTYP.XLS"

         Dim ANS3 As Long
         
424       ANS3 = MsgBox("Εκτυπώνω;", vbYesNo)
          
426   If ANS3 = vbYes Then
428        myXL.PrintOut 1, 1, 1, False
      End If


       

430     ANS3 = MsgBox("Κλείνω το EXCEL", vbYesNo)

432     If ANS3 = vbYes Then
434         Call workbook.Close(False)
436         excel.Quit
438         Set excel = Nothing
        End If
        
    On Error Resume Next
    
    
440            excel.Visible = True

        '<EhFooter>
        Exit Sub

cmdKARTELAseEXCEL_Click_Err:
        MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Command11_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Command11_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command12_Click()

        Dim ANS3 As Long

        'ans3=msgbox("Να τυπωθεί στον προεπιλεγμένο εκτυπωτή ν/ο"

        '<EhHeader>
        On Error GoTo Command12_Click_Err

        '</EhHeader>
 
        Dim sql  As String, fDB As Database, D1 As Date, d2 As Date

        Dim synt As String, DB As Database

        Dim DUM

        Dim f_psifia_atim

100     f_psifia_atim = 6

        Dim x As String

110     DBGRID22.Col = 1

        'ΑΝ ΔΕΝ ΣΗΜΑΔΕΥΕΙ ΠΑΡΑΣΤΑΤΙΚΟ
120     If DBGRID22.row = -1 Then Exit Sub

        On Error Resume Next

130     D1 = CDate(Format(DBGRID22.Text, "dd/mm/yyyy"))

        '   d2 = DateAdd("d", 1, d1)
        Dim MATIM As String, file As String

140     DBGRID22.Col = 0
150     MATIM = DBGRID22.Text
 
        '20120903T015014 -1.pdf
 
160     file = "C:\PDF\" + Format(D1, "YYYYMMDD") + Trim(MATIM) + " -1.PDF"

170     If Len(Dir(file)) > 2 Then
 
            Dim OpenFile As String, ARL As String, fl As String

180         ARL = "C:\PDF\ACROWRAP.EXE  " ', vbNormalFocus
190         fl = file
            'fl = "20130719τ01808 -1.PDF"
200         OpenFile = ARL & fl
210         Shell OpenFile, vbNormalFocus
            ' AcroPDF1.LoadFile FILE
      
220         file = "C:\PDF\" + Format(D1, "YYYYMMDD") + Trim(MATIM) + " -2.PDF"
230         ' ARL = "C:\MERCVB\READER\AcroRd32.exe " ', vbNormalFocus
240         fl = file
            ' fl = "20130719τ01808 -2.PDF"
250         OpenFile = ARL & fl
260         Shell OpenFile, vbNormalFocus
            ' AcroPDF1.LoadFile FILE
      
        Else
270         MsgBox "ΔΕΝ ΒΡΕΘΗΚΕ ΤΟ ΑΡΧΕΙΟ"
        End If

        '<EhFooter>
        Exit Sub

Command12_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Command12_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Command12_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command13_Click()
    data1.Recordset.Move 0
End Sub

Private Sub cmdPRINTKARTELA_Click()

        '     Data1.Recordset.Delete
        '<EhHeader>
        On Error GoTo cmdPRINTKARTELA_Click_Err

        '</EhHeader>
        Dim R, sxre, spis, DB As Database, dbac As Database

        Dim MKOD, mEpo

        'Set db = OpenDatabase(gDir, False, False, gConnect)
        'Set dbac = OpenDatabase("c:\mercvb\reports\reports.mdb")
        On Error Resume Next

        'Gdb.Execute "drop table kartela" + xeirisths
        ' Gdb.Execute "drop table kartelap"
        On Error GoTo cmdPRINTKARTELA_Click_Err

        ' ΕΛΕΓΧΩ ΑΝ ΥΠΑΡΧΕΙ ΤΟ TABLE  "MEM" ΓΙΑ ΝΑ ΤΟ ΣΒΗΣΩ
        Dim k, YPARXEI(5)

        'YPARXEI(1) = 0
        'For K = 0 To dbac.TableDefs.Count - 1
        '   If UCase(dbac.TableDefs(K).Name) = "MEM" Then
        '      YPARXEI(1) = 1
        '   End If
        'Next
        'If YPARXEI(1) = 1 Then
        'dbac.Execute "drop table mem"

100     egg.RecordSource = "SELECT * FROM KARTELA"

110     If f_direct = 2 Then
            '  On Error Resume Next
            
            
            Dim t2 As String
       If data1.Recordset("eidos") = "e" Then
          t2 = "select ATIM as [Αρ_Παρ],HME as [Ημερ/νία],EID as [Είδ],"
          t2 = t2 + "APA as [Αριθ_Παρ],AIT as [Αιτιολογία],ISNULL(XREOSI,0) as [Χρέωση],ISNULL(PISTOSI,0) as [Πίστωση],"
          t2 = t2 + "PROOD as [Υπόλοιπο] into KARTELADOK   from EGG "
          t2 = t2 + "where EIDOS='e' and KOD='" + data1.Recordset("kod") + "' order by HME,ATIM;"
          Gdb.Execute "DROP TABLE  KARTELAP "
          
          Gdb.Execute "DROP TABLE  KARTELADOK "
          Gdb.Execute t2
          Gdb.Execute "DELETE FROM KARTELA"
          Gdb.Execute "INSERT INTO KARTELA SELECT * FROM KARTELADOK"
          
          
         ' Gdb.Execute "select * into kartelap in 'c:\mercvb\reports\reports.mdb' from pel where kod='" + MKOD + "' and eidos='e'"
          Gdb.Execute "select * into KARTELAP  from PEL where KOD='" + data1.Recordset("kod") + "' and EIDOS='e'"
       Else
          Gdb.Execute "select ATIM as [Αρ_Παρ],FORMAT(hme,'DD/MM/YY') as [Ημερ/νία],EID as [Είδ],APA as [Αριθ_Παρ],AIT as [Αιτιολογία],iif(xre<0,-xre,0) as [Χρέωση],iif(xre>0,xre,0) as [Πίστωση],FORMAT(PROOD,'###,###.00') as [Υπόλοιπο] into kartela in 'c:\mercvb\reports\reports.mdb' from egg where eidos='r' and kod='" + data1.Recordset("kod") + "' order by hme,atim;"
          Gdb.Execute "select * into KARTELAP  from PEL where KOD='" + MKOD + "' and EIDOS='r'"
       
       '   Gdb.Execute "select * into kartelap in 'c:\mercvb\reports\reports.mdb' from pel where kod='" + MKOD + "' and eidos='r'"
       End If
       'CrystalReport1.Action = 1

            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            

120         CrystalReport1.Connect = gConnect
130         CrystalReport1.ReportFileName = "c:\mercvb\reports\kartela2.rpt"

140       '  CrystalReport1.SelectionFormula = "{EGG.KOD}='" + DATA1.Recordset("KOD") + "' and  {EGG.EIDOS}='" + DATA1.Recordset("EIDOS") + "'"

150        ' CrystalReport1.SortFields(0) = "+{EGG.HME}"
160         CrystalReport1.Action = 1

            Exit Sub

        End If

170     If f_direct = 0 Then
180         utDialog.Command3.Enabled = False
190         utDialog.Command4.Enabled = False
200         utDialog.SHOW 1
        End If

        'db.Execute "select top 1 * into mem in 'c:\mercvb\reports\reports.mdb'from mem "

        Dim T As String

210     MKOD = data1.Recordset("kod")
        'If Option1(1) Then
        '  t = "select ATIM as [Αρ_Παρ],HME as [Ημερ/νία],EID as [Είδ],"
        '  t = t + "APA as [Αριθ_Παρ],AIT as [Αιτιολογία],XREOSI as [Χρέωση],PISTOSI as [Πίστωση],"
        '  t = t + "PROOD as [Υπόλοιπο] into KARTELA" + xeirisths + "  from EGG "
        '  t = t + "where EIDOS='e' and KOD='" + Data1.Recordset("kod") + "' order by HME,ATIM;"
        '  Gdb.Execute t
        'Gdb.Execute "select * into kartelap in 'c:\mercvb\reports\reports.mdb' from pel where kod='" + mkod + "' and eidos='e'"
        'Else
        'Gdb.Execute "select ATIM as [Αρ_Παρ],FORMAT(hme,'DD/MM/YY') as [Ημερ/νία],EID as [Είδ],APA as [Αριθ_Παρ],AIT as [Αιτιολογία],iif(xre<0,-xre,0) as [Χρέωση],iif(xre>0,xre,0) as [Πίστωση],FORMAT(PROOD,'###,###.00') as [Υπόλοιπο] into kartela in 'c:\mercvb\reports\reports.mdb' from egg where eidos='r' and kod='" + Data1.Recordset("kod") + "' order by hme,atim;"
        'Gdb.Execute "select * into kartelap in 'c:\mercvb\reports\reports.mdb' from pel where kod='" + mkod + "' and eidos='r'"
        'End If
        'CrystalReport1.Action = 1

        'Exit Sub

220     sxre = 0
230     spis = 0
    
240     Set R = egg.Recordset
        ' Printer.FontName = "COURIER"

        On Error Resume Next

250     Printer.FontSize = 10
        'Printer.FontName = "Courier New Greek"

        ' Printer.FontName = "Courier New"
        ' Printer.Font.Charset = 161
        ' Printer.FontBold = True
        ' Printer.Print Tab(34); Data1.Recordset("EPO")
        ' Printer.Print

        '    Printer.Print "Ημερ/νία ";
        '    Printer.Print "Παρ/κό  ";
        '    Printer.Print left("Αιτιολογία              " + Space(40), 25);
        '    Printer.Print "      Χρέωση   ";
        '    Printer.Print "  Πίστωση  ";
        '    Printer.Print "  Υπόλοιπο  "

        'Printer.FontBold = False

        ' R.MoveFirst

        'On Error GoTo Command2_Click_Err

        'Do While Not R.EOF

        ' select ATIM as [Αρ_Παρ],hme as [Ημερ/νία],EID as [Είδ],APA as [Αριθ_Παρ],AIT as [Αιτιολογία],
        ' iif(xre>0,xre,0) as [Χρέωση],iif(xre<0,-xre,0) as [Πίστωση],PROOD as [Υπόλοιπο] from egg where eidos='" + x2 + "' and kod='" + s + "' order by hme,atim;"

260     Printer.FontSize = 9
270     Printer.FontName = "Courier New"
280     Printer.Font.Charset = 161
        ' Printer.FontBold = True

290     Printer.Orientation = vbPRORPortrait    ' vbPRORLandscape

        Dim R22 As New ADODB.Recordset

300     R22.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic

310     Printer.Print R22("pelono")
320     Printer.Print R22("pelepa")

330     Printer.Print Now
340     Printer.Print "------------------------------------------------------------------------------------------"
350     R22.Close

360     If data1.Recordset("EIDOS") = "e" Then
370         Printer.Print "Πελάτης     "
        Else
380         Printer.Print "Προμηθευτής "
        End If

        'Open "C:\KARTELA.TXT" For Output As #1
390     Printer.Print data1.Recordset("epo") + "   Κωδ." + data1.Recordset("kod")
400     Printer.Print data1.Recordset("die")
410     Printer.Print "Tηλέφωνο " + data1.Recordset("thl")
420     Printer.Print

        ' Printer.Print "Χρέωση   Πίστωση    Υπόλοιπο";

        Dim mait As String

430     For k = 0 To MSFlexGrid1.rows - 1

440    '     sxre = sxre + R("[Χρέωση]")
450     '    spis = spis + R("[Πίστωση]")

460         If k = MSFlexGrid1.rows - 1 Then
470             Printer.Print
            End If

480         Printer.Print tostC(MSFlexGrid1.TextMatrix(k, 0), 10);  ' a14609B
490         Printer.Print tostC(MSFlexGrid1.TextMatrix(k, 1), 5);  ' 01/10/2005
            ' Printer.Print " " + tostC(MSFlexGrid1.TextMatrix(k, 2), 3);  ' py

500         Printer.Print tostC(MSFlexGrid1.TextMatrix(k, 3), 11); " ";    '  1600 ar.par

510         mait = MSFlexGrid1.TextMatrix(k, 4)
520         Printer.Print Left(mait + Space(20), 20);  ' da pelath

530         If k = 0 Then   ' τιτλοι
540             Printer.Print Right(Space(15) + MSFlexGrid1.TextMatrix(k, 5), 15);  '   xreosi
550             Printer.Print Right(Space(15) + MSFlexGrid1.TextMatrix(k, 6), 15);  ' pistosi
560             Printer.Print Right(Space(15) + MSFlexGrid1.TextMatrix(k, 7), 15);  'ypoloipo
            
            Else
570             Printer.Print TOST(Val(Replace(MSFlexGrid1.TextMatrix(k, 5), ",", ".")), 15);  '   xreosi
580             Printer.Print TOST(Val(Replace(MSFlexGrid1.TextMatrix(k, 6), ",", ".")), 15);  ' pistosi
590             Printer.Print TOST(Val(Replace(MSFlexGrid1.TextMatrix(k, 7), ",", ".")), 15);  ' YPOLOIPO
                'Printer.Print TOST(Val(MSFlexGrid1.TextMatrix(k, 7)), 12);  'ypoloipo
            End If

600        Printer.Print "  " '+ MSFlexGrid1.TextMatrix(K, 8)  ' Format(Val(MSFlexGrid1.TextMatrix(K, 8)), "##") 'ΧΡΗΣΤΗς
    
610         If k > 0 And Val(MSFlexGrid1.TextMatrix(k, 9)) > 0 Then Gdb.Execute "UPDATE EGG SET PROOD=" + MSFlexGrid1.TextMatrix(k, 7) + " WHERE ID=" + MSFlexGrid1.TextMatrix(k, 9)
    
        Next

        ' PRINTER.PRINT, "Σύνολα"; Tab(43); TOST(sxre, 12);
        ' PRINTER.PRINT, TOST(spis, 12);

        'Close #1
        Dim DUM

        'dum = Shell("c:\mercvb\NOTEPAd.exe c:\kartela.txt ", vbNormalFocus)
620     Printer.EndDoc

        Exit Sub

        '<EhFooter>
        Exit Sub

cmdPRINTKARTELA_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function TOST(ByVal s As Single, ByVal N As Integer) As String

        '<EhHeader>
        On Error GoTo TOST_Err

        '</EhHeader>
If s = 0 Then
   TOST = Space$(N)
Else
   TOST = Right(Space(30) + Format(s, "###,###.00"), N)
End If
        '<EhFooter>
        Exit Function

TOST_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.TOST " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.TOST " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function tostC(ByVal s, ByVal N As Integer) As String

        '<EhHeader>
        On Error GoTo tostC_Err

        '</EhHeader>

100     If IsNull(s) Then
110         tostC = Space(N)
        Else
120         tostC = Left(s + Space(N), N)
        End If

        '<EhFooter>
        Exit Function

tostC_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.tostC " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.tostC " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub cmdEIDHSYNOLIKA_Click()

        '<EhHeader>
        On Error GoTo cmdEIDHSYNOLIKA_Click_Err

        '</EhHeader>
        Dim k As Integer

   
          Dim nc As String

     nc = ""
Dim N As Integer
120     For N = 0 To List1.ListCount - 1

            'On Error Resume Next
130         If List1.Selected(N) Then
140             nc = nc + "'" + Right(List1.List(N), 1) + "',"
            End If

        Next

150     If Len(nc) > 0 Then
160         nc = Left(nc, Len(nc) - 1)
170         nc = " left (ATIM,1) IN (" + nc + ") AND "
        End If
   
   
   
        'If fflag = 0 Then Exit Do
    Dim FPELAT42 As New PELAT42
    
100     FPELAT42.DTPicker1 = DTPicker1
110     FPELAT42.DTPicker2 = DateAdd("d", 1, DTPicker2)



   FPELAT42.PEL_EPO = data1.Recordset("EPO")
   FPELAT42.PEL_KOD = data1.Recordset("KOD")
   FPELAT42.PEL_EIDOS = data1.Recordset("EIDOS")
   
   
   
   
FPELAT42.LAB_nc = nc



' ΑΝ ΘΑ ΥΠΟΛΟΓΙΖΕΙ ΦΕΤΟΣ ή ΠΕΡΥΣΙ
 Dim mCon As String
320     If Check1.Value = vbChecked Then    '   <> gDir Then ' palia xronia
330        mCon = gConnect + ";DATABASE=" + f_PALIAXRONIA     '"DSN=" + LTrim(Trim(Str(f_PALIAXRONIA)))
340     ElseIf Check2.Value = vbChecked And Len(XRONIES.Text) > 0 Then  '   <> gDir Then ' palia xronia
350          mCon = gConnect + ";DATABASE=" + f_PROPERSI  ' "DSN=" + LTrim(Str(f_PALIAXRONIA))
        Else
360          mCon = gConnect
        End If

     FPELAT42.lblConn = mCon
     FPELAT42.SHOW

        '<EhFooter>
        Exit Sub

cmdEIDHSYNOLIKA_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdXREOPISTOSEIS_Click()

Dim RR As New ADODB.Recordset
 RR.Open "SELECT ISNULL(ENABLED,1) AS ENERGO FROM ACTIVE_MENUS WHERE NAME='MENUXREOPISFORCHECK' AND USER_ID=" + str(gUserId), Gdb, adOpenDynamic, adLockOptimistic
 If RR.EOF Then
    'OK
 Else
    If RR!energo = 0 Then
        Exit Sub
        RR.Close
    End If
 End If
RR.Close






        '<EhHeader>
        On Error GoTo cmdXREOPISTOSEIS_Click_Err

        '</EhHeader>
        Dim DUM

        Dim x As String
        
        x = mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1)

        gBuff = x
        
100     pelat43.PELDATA1.ConnectionString = data1.ConnectionString
110     pelat43.PELDATA1.RecordSource = "select EPO,THL,DIE,* from PEL where EIDOS='" + data1.Recordset("EIDOS") + "' and KOD='" + data1.Recordset("KOD") + "'"
120     pelat43.PELDATA1.Refresh
130     pelat43.Left = 0
140     pelat43.Top = 0
150     pelat43.Caption = Label2.Caption

160     pelat43.SHOW 1
170

        If f_KENTRA_ADYNATISMATOS = 1 Then
   
           ' DATA1.Refresh
    
            If IsNull(data1.Recordset("num5")) Then
                ypolmhx.Caption = data1.Recordset("num2")
            Else
                ypolmhx.Caption = data1.Recordset("num2") - data1.Recordset("num5")
            End If
        End If
If F_NEA_KRTELLA = 0 Then
190     DUM = KARTELLA(gDir)
Else
        neakartela 1
End If
        
        '  egg.Refresh
        '  egg.Recordset.MoveLast
        '<EhFooter>
        Exit Sub

cmdXREOPISTOSEIS_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command5_Click()

        '<EhHeader>
        On Error GoTo Command5_Click_Err

        '</EhHeader>
        Dim mgdir, DUM

100     LAST_TIM = -1    ' μετρητης grid

110     If Left(Command5.Caption, 1) = "Τ" Then
120         Command5.Caption = "ΠΡΟΗΓΟΥΜΕΝΗ ΧΡΗΣΗ"
130         mgdir = gDir
        Else
140         mgdir = gDir + "\" + Right(Format(Year(Now) - 1, "0000"), 2)
            'd1 = DTPicker1.Value - 365
            'd2 = DTPicker2.Value - 365
150         Command5.Caption = "ΤΡΕΧΟΥΣΑ ΧΡΗΣΗ"
        End If

160     DUM = KARTELLA(mgdir)

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdPRINTMEEIDH_Click()

        '<EhHeader>
        On Error GoTo cmdPRINTMEEIDH_Click_Err

        '</EhHeader>
        Dim a, B, c, D

        Dim x                As String

        Dim polepis          As String, AGOEPIS As String, pol As String, ago As String

        'sql με συνδυασμο egg & ειδη
        'sql = "SELECT  CONVERT(CHAR(12),EGG.HME,3),EGG.ATIM,EGG.AIT,EGGTIM.KODE,EGGTIM.ONOMA,EGGTIM.POSO,EGGTIM.TIMM,EGG.XREOSI,EGG.PISTOSI  FROM EGG INNER JOIN EGGTIM  ON EGG.ATIM=EGGTIM.ATIM  AND EGG.KOD=EGGTIM.PELKOD  WHERE EGG.XREOSI>0  AND EGG.KOD='2877'  ORDER BY EGG.HME,EGG.ATIM"

        Dim SELE_PARASTATIKA As String

        Dim N                As Long

100     SELE_PARASTATIKA = ""

110     For N = 0 To List1.ListCount - 1

            'On Error Resume Next
120         If List1.Selected(N) Then
130             SELE_PARASTATIKA = SELE_PARASTATIKA + "'" + Right(List1.List(N), 1) + "',"
            End If

        Next

        'If Len(SELE_PARASTATIKA) > 0 Then
        '   SELE_PARASTATIKA = Left(SELE_PARASTATIKA, Len(SELE_PARASTATIKA) - 1)
        '   SELE_PARASTATIKA = " left (ATIM,1) IN (" + SELE_PARASTATIKA + ") AND "
        'End If

        Dim MGYRNA_FETOS As Integer

140     MGYRNA_FETOS = 0

        '===========  TO PARAKATO EINAI GIA 2 XRONIES
150     If Check1.Value = vbChecked Then    '   <> gDir Then ' palia xronia
160         MGYRNA_FETOS = 1
170         egg.ConnectionString = gConnect + ";DATABASE=" + f_PALIAXRONIA    ' "DSN=" + LTrim(Str(f_PALIAXRONIA))
180         Gdb.Close
190         Gdb.Open gConnect + ";DATABASE=" + f_PALIAXRONIA    ' "DSN=" + LTrim(Str(f_PALIAXRONIA))
        
200     ElseIf Check2.Value = vbChecked And Len(XRONIES.Text) > 2 Then  '   <> gDir Then ' palia xronia

            ' XRONIES.AddItem Left(R("TITLOS") + Space(20), 20) + R("DSN")
210         f_PROPERSI = Trim(mID(XRONIES.Text, 21, 20))
220         egg.ConnectionString = gConnect + ";DATABASE=" + f_PROPERSI    ' "DSN=" + LTrim(Str(f_PALIAXRONIA))
230         Gdb.Close
240         Gdb.Open gConnect + ";DATABASE=" + f_PROPERSI    ' "DSN=" + LTrim(Str(f_PALIAXRONIA))
250         MGYRNA_FETOS = 2

        Else
260         egg.ConnectionString = gConnect
        End If

270     a = Get_AJ(pol, polepis, ago, AGOEPIS)

        Dim ANS As Integer

280     ANS = MsgBox("Με στήλη κέρδος;", vbYesNo)

290     a = data1.Recordset("kod")

300     x = "select ATIM,CONVERT(CHAR(10),HME,103)+' ' as [Ημερoμηνία],KODE as [Κωδικός],LEFT(ONOMA,30) as [Περιγραφή],POSO as [Ποσότητα],"
310     x = x + " str(TIMM*POSO,10,2) as [Αξία],EKPT AS [ΕΚΠΤΩΣΗ],str(TIMM*POSO*(1-EKPT/100),10,2) as [Αξία μετά την εκπτωση],LEFT(ATIM,1)"

320     If ANS = vbYes Then
330         x = x + " ,STR(CASE WHEN LEFT(ATIM,1) IN (" + pol + ") THEN POSO*(TIMM-EID.XTI) ELSE 0 END ,10,2) AS [Κέρδος],EID.XTI AS [ΤΙΜ.ΑΓΟΡ] "
        End If

340     x = x + " FROM EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD "
350     x = x + " where EIDOS='" + data1.Recordset("EIDOS") + "'"

360     If Len(SELE_PARASTATIKA) > 0 Then
370         SELE_PARASTATIKA = Left(SELE_PARASTATIKA, Len(SELE_PARASTATIKA) - 1)
380         x = x + " AND left (ATIM,1) IN (" + SELE_PARASTATIKA + ")  "
        End If

390     x = x + "  AND PELKOD='" + data1.Recordset("kod") + "' ORDER BY HME,ATIM,PROOD_AJ;"

        'On Error GoTo Command6_Click_Err
400     egg.RecordSource = x

410     egg.Refresh

420     a = print3_xar(x, "0001101", data1.Recordset("epo") + "   Κωδ." + data1.Recordset("kod"), 1)    'typos2

430     If MGYRNA_FETOS > 0 Then

440         Gdb.Close
450         Gdb.Open gConnect

        End If

        '<EhFooter>
        Exit Sub

cmdPRINTMEEIDH_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Command6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Command6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command7_Click()

        '<EhHeader>
        On Error GoTo Command7_Click_Err

        '</EhHeader>

        Dim fff As New pelat2

100     gApoMenu = False

110     fff.ffbuff = f_buff + data1.Recordset("kod") ' Text1(0).Text

120    ' fff.Command5.Enabled = False
   
        ' fff.ffbuff = f_buff + DATA1.Recordset("kod")    ' Text1(0).Text
         fff.PELPROM.Text = PELPROM.Text
         fff.Text2(0).Text = Left(data1.Recordset!kod, 5)
          fff.Text2(1).Text = ""
        fff.Text2_LostFocus 0
        
        

        If f_KENTRA_ADYNATISMATOS = 1 Then
            If kentroAdyn.ListIndex >= 0 Then
                fff.Text2(0).Text = data1.Recordset("KOD")
                fff.data1.ConnectionString = gConnect
                fff.kentroAdyn.ListIndex = kentroAdyn.ListIndex ' Comb
            End If
   
        End If

130     fff.SHOW
140     'gApoMenu = True

        Exit Sub

        'left(data1.Recordset!kod, 5) + data1.Recordset!epo
        On Error GoTo LATOS

150     pelat2.PELPROM.Text = PELPROM.Text
160     pelat2.Text2(0).Text = Left(data1.Recordset!kod, 5)

170     pelat2.SHOW

        Exit Sub

LATOS:

        '    pelat2.Text2(0).SetFocus

        '<EhFooter>
        Exit Sub

Command7_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Command7_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Command7_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command8_Click()

        '------------------------------------------------------------------------
        '<EhHeader>
        On Error GoTo Command8_Click_Err

        '</EhHeader>
        Dim sql  As String, fDB As Database, D1 As Date, d2 As Date

        Dim synt As String, DB As Database

        Dim DUM

        Dim f_psifia_atim

100     f_psifia_atim = 6

        Dim x As String

110     DBGRID22.Col = 1

        'ΑΝ ΔΕΝ ΣΗΜΑΔΕΥΕΙ ΠΑΡΑΣΤΑΤΙΚΟ
120     If DBGRID22.row = -1 Then Exit Sub

        On Error Resume Next
        tdbgrid1.Col = 0
        D1 = CDate(Format(tdbgrid1.Text, "dd/MM/yyyy"))
        
130   '  d1 = CDate(Format(DBGRID22.Text, "dd/MM/yyyy"))
140     d2 = DateAdd("d", 1, D1)

150     DBGRID22.Col = 0

Dim M_ATIM As String
   tdbgrid1.Col = 1
   M_ATIM = tdbgrid1.Text
   



160     x = "select KODE,EID.ERG,ATIM,left(ONOMA,35) as [Περιγραφή],POSO as [Ποσότητα],TIMM as [Tιμ_Μον]," & " EKPT as [Εκπτωση],str(KAU_AJIA,10,2) as [K_Αξία],str(MIK_AJIA,10,2) as [ΜΕΦΠΑ_Αξία] from EGGTIM INNER JOIN EID  ON EGGTIM.KODE=EID.KOD where" & " left(ATIM," + str(f_psifia_atim + 1) + ")='" + M_ATIM + "'"
170     x = x + " and HME>='" + Format(D1, "MM/dd/yyyy") + "' and HME<'" + Format(d2, "MM/dd/yyyy") + "'"
180     x = x + " order by PROOD_AJ;"

        Dim TIT As String

190     TIT = DbGrid2.Text + " " + DBGRID22.Text + Format(D1, "dd/mm/yyyy")

200     f_PROPERSI = FINDPARAMETROI(1, "PAR1", "F_PROPERSI", "2005", "Γ.19 ΒΑΣΗ ΠΡΟΠΕΡΣΙΝΗΣ ΧΡΟΝΙΑΣ")
210     f_PALIAXRONIA = Trim(FINDPARAMETROI(1, "PAR1", "F_PALIAXRONIA", "2005", "Γ.18 DSN ΠΑΛΙΑΣ ΧΡΟΝΙΑΣ"))

        '===========
        Dim mgconnect As String

220     mgconnect = gConnect    '  As String

230     If Check1.Value = vbChecked Then    '   <> gDir Then ' palia xronia
240         gConnect = gConnect + ";DATABASE=" + f_PALIAXRONIA   ' "DSN=" + LTrim(Str(f_PALIAXRONIA))
250         Gdb.Close
260         Gdb.Open gConnect

        End If
        
        
        
        If Check2.Value = vbChecked And Len(XRONIES.Text) > 2 Then  '   <> gDir Then ' palia xronia

            ' XRONIES.AddItem Left(R("TITLOS") + Space(20), 20) + R("DSN")
             f_PROPERSI = Trim(mID(XRONIES.Text, 21, 20))
        

270     '  If Check2.Value = vbChecked Then    '   <> gDir Then ' palia xronia
280         gConnect = gConnect + ";DATABASE=" + f_PROPERSI   ' "DSN=" + LTrim(Str(f_PALIAXRONIA))
290         Gdb.Close
            On Error GoTo checkdb
300         Gdb.Open gConnect

        End If

310     DUM = print3_xar(x, "111111111", TIT, 0)
        
320     If Check2.Value = vbChecked Or Check1.Value = vbChecked Then    '   <> gDir Then ' palia xronia
330         gConnect = mgconnect
340         Gdb.Close
            Gdb.Open gConnect
        End If

        '<EhFooter>
        Exit Sub

checkdb:
    MsgBox Err.Description
    Resume Next
    





Command8_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Command8_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Command8_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdGraph_Click()

        '<EhHeader>
        On Error GoTo cmdGraph_Click_Err

        '</EhHeader>
        Dim R As New ADODB.Recordset

        Dim k As Integer

100     If data1.Recordset.EOF Then Exit Sub
110     graph1.Visible = True

120     MSFlexGrid1.Visible = False
        tdbgrid1.Visible = False
        

130     R.Open "SELECT MONTH(HME),SUM(XREOSI),SUM(PISTOSI) FROM EGG WHERE KOD='" + data1.Recordset("kod") + "' AND EIDOS='" + data1.Recordset("EIDOS") + "' GROUP BY MONTH(HME) ", Gdb, adOpenForwardOnly, adLockReadOnly

        Dim A_XRE(12)   As Single

        Dim A_PIS(12)   As Single

        Dim A_XREAJ(12) As Single

        Dim A_PISAJ(12) As Single

140     For k = 1 To 12: A_XRE(k) = 0: A_PIS(k) = 0: A_XREAJ(k) = 0: A_PISAJ(k) = 0: Next

        ' Exit Sub
150     If Not R.EOF Then

            On Error Resume Next

160         R.MoveFirst

170         Do While Not R.EOF
180             A_XRE(R(0)) = R(1): A_XREAJ(R(0)) = R(3)
190             A_PIS(R(0)) = R(2): A_PISAJ(R(0)) = R(4)
200             R.MoveNext
            Loop

        End If

        'graph1.top = 4680  'arxika
        '  graph1.height = 4485

        ' graph1.top = 46

210     graph1.row = 2
220     graph1.RowLabel = "ΑΓΟΡΕΣ"

230     graph1.row = 1
240     graph1.RowLabel = "ΠΛΗΡΩΜΕΣ"

250     For k = 1 To 12
            '  G = "G" + Format(k, "00")
            '  M = "M" + Format(k, "00")
            '  GA = "g" + Format(k, "00")
            '  MA = "M" + Format(k, "00")

260         graph1.row = 1: graph1.Column = k: graph1.DATA = A_XRE(k)    ' DATA2.Recordset(G)
270         graph1.row = 2: graph1.Column = k: graph1.DATA = A_PIS(k)    'DATA2.Recordset(M)
280         graph1.ColumnLabel = k
290         graph1.Refresh
        Next

300     MsgBox "ΟΚ"
310     MSFlexGrid1.Visible = True

320     graph1.Visible = False
        tdbgrid1.Visible = True
        
        '<EhFooter>
        Exit Sub

cmdGraph_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Command9_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Command9_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DbGrid2_Click()
    DBGrid2_RowColChange 0, 0
    
     
End Sub

Private Sub DBGrid2_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo DBGrid2_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = 13 Then
110         KeyAscii = 13
            '
        
120         DBGrid2_RowColChange 0, 0
            '  Command4.SetFocus
        End If

        '<EhFooter>
        Exit Sub

DBGrid2_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.DBGrid2_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.DBGrid2_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DbGrid2_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo DbGrid2_KeyUp_Err

        '</EhHeader>

100     If KeyCode = 38 Or KeyCode = 40 Then
110         KeyCode = 13
            '

120         DBGrid2_RowColChange 0, 0
            '  Command4.SetFocus
        End If

        '<EhFooter>
        Exit Sub

DbGrid2_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.DbGrid2_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.DbGrid2_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub neakartela(ByVal withID As Integer)

Dim myp2 As String

               tdbgrid1.Visible = True
               tdbgrid1.FilterBar = False
               
              

'    TDBGrid1.columns(6).Width = 1544.882
'    TDBGrid1.columns(5).Width = 2610.142
'    TDBGrid1.columns(4).Width = 480.189
'    TDBGrid1.columns(3).Width = 3550.142
'    TDBGrid1.columns(2).Width = 929.7639
'    TDBGrid1.columns(1).Width = 900.2835
'    TDBGrid1.columns(0).Width = 1049.953



               
               
               
               
               
               Dim mxreo As String
               
               If gBuff = "r" Then
                  'myp2 = str(Data1.Recordset("ayp"))
                  mxreo = "-isnull(XREOSI,0)+isnull(PISTOSI,0)"
               Else
                 ' myp2 = str(Data1.Recordset("ayp"))
                  mxreo = "isnull(XREOSI,0)-isnull(PISTOSI,0)"
               End If
              ' myp2 = Replace(myp2, ",", ".")
              
              
              
'240     If Check1.Value = vbChecked Then    '   <> gDir Then ' palia xronia
'
'250         egg.ConnectionString = gConnect + ";DATABASE=" + f_PALIAXRONIA    ' "DSN=" + LTrim(Str(f_PALIAXRONIA))
'            '   gConnect ";DATABASE=" + f_PALIAXRONIA
'260         R.Open "select AYP FROM PEL WHERE EIDOS='" + gBuff + "' AND KOD='" + s + "'", gConnect + ";DATABASE=" + f_PALIAXRONIA, adOpenForwardOnly, adLockReadOnly
'
'270         If R.EOF Then
'280             R.Close
'
'                Exit Function
'
'            End If
'
'290         ayp = R(0)
'300     ElseIf Check2.Value = vbChecked And Len(XRONIES.Text) > 2 Then  '   <> gDir Then ' palia xronia
'
'            ' XRONIES.AddItem Left(R("TITLOS") + Space(20), 20) + R("DSN")
'310         f_PROPERSI = Trim(mID(XRONIES.Text, 21, 20))
'
'            '    If Len(f_PROPERSI) = 0 Then
'            '       MsgBox "ΔΙΑΛΕΞΤΕ ΧΡΟΝΙΑ"
'            '      Exit Function
'            '  End If
'
'320         egg.ConnectionString = gConnect + ";DATABASE=" + f_PROPERSI    ' "DSN=" + LTrim(Str(f_PALIAXRONIA))
'
'            '   gConnect ";DATABASE=" + f_PALIAXRONIA
'            On Error Resume Next
'
'330         R.Open "select AYP FROM PEL WHERE EIDOS='" + gBuff + "' AND KOD='" + s + "'", gConnect + ";DATABASE=" + f_PROPERSI, adOpenForwardOnly, adLockReadOnly
'
'340         If R.EOF Then
'350             R.Close
'
'                Exit Function
'
'            End If
'
'360         ayp = R(0)
'        Else
'
'        End If
    Dim FDB2 As New ADODB.Connection
     Dim r0x, r1p, r2, ayp As Single
     
     
     
 '    If cXRONIES.Value = vbChecked Then
     
     
     
     
'        Dim SQL11 As String
'        SQL11 = "SELECT * INTO " + fFILE + " FROM EGGTIM  where " + filtroPEL.Caption + " HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "'  AND APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' "
'
'        SQL11 = SQL11 + " UNION SELECT *  FROM " + f_PALIAXRONIA + ".dbo.EGGTIM  where " + filtroPEL.Caption + " HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "'  AND APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' "
'        Dim mPROP As String
'        mPROP = FINDPARAMETROI(1, "PAR1", "F_PROPERSI", "", "Γ.19 ΒΑΣΗ ΠΡΟΠΕΡΣΙΝΗΣ ΧΡΟΝΙΑΣ")
'        SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        Dim mPROP2 As String
'
'
'
'        mPROP3 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI2", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΙΝ 3 XPONIA")
'        If Len(mPROP3) > 0 Then
'           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP3 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        End If
'
'        Dim mPROP4 As String
'        mPROP4 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI3", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΙΝ 4 XPONIA")
'        If Len(mPROP4) > 0 Then
'           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP4 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        End If
'
'
'          Dim mPROP5 As String
'        mPROP5 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI5", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΟ 5ΕΤΙΑΣ")
'        If Len(mPROP5) > 0 Then
'           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP5 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        End If
'
'          Dim mPROP6 As String
'        mPROP6 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI6", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΟ 6ΕΤΙΑΣ")
'        If Len(mPROP6) > 0 Then
'           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP6 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        End If
'
'           Dim mPROP7 As String
'        mPROP7 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI7", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΟ 7ΕΤΙΑΣ")
'        If Len(mPROP7) > 0 Then
'           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP7 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        End If
'
'           Dim mPROP8 As String
'        mPROP8 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI8", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΟ 8ΕΤΙΑΣ")
'        If Len(mPROP8) > 0 Then
'           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP8 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        End If
'
'        Dim mPROP9 As String
'        mPROP9 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI9", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΟ 9ΕΤΙΑΣ")
'        If Len(mPROP9) > 0 Then
'           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP9 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        End If
'
'
'        Dim mPROP10 As String
'        mPROP10 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI10", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΟ 10ΕΤΙΑΣ")
'        If Len(mPROP10) > 0 Then
'           SQL11 = SQL11 + " UNION SELECT *  FROM " + mPROP10 + ".dbo.EGGTIM  where " + filtroPEL.Caption + " APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' " '      + "  Order by HME,ATIM "  convert(char(10),HME,2),ATIM "
'        End If
       ' Gdb.Execute SQL11
  ' Else
    
         If Check1.Value = vbChecked Then    '   <> gDir Then ' palia xronia
             Adodc1.ConnectionString = gConnect + ";DATABASE=" + f_PALIAXRONIA    ' "DSN=" + LTrim(Str(f_PALIAXRONIA))
             FDB2.Open gConnect + ";DATABASE=" + f_PALIAXRONIA
          ElseIf Check2.Value = vbChecked And Len(XRONIES.Text) > 2 Then  '   <> gDir Then ' palia xronia
            ' XRONIES.AddItem Left(R("TITLOS") + Space(20), 20) + R("DSN")
             f_PROPERSI = Trim(mID(XRONIES.Text, 21, 20))
             Adodc1.ConnectionString = gConnect + ";DATABASE=" + f_PROPERSI     ' "DSN=" + LTrim(Str(f_PALIAXRONIA))
             FDB2.Open gConnect + ";DATABASE=" + f_PROPERSI
          Else
             Adodc1.ConnectionString = gConnect
             FDB2.Open gConnect
          End If
     
     
     
    'End If
     
     
 
     
     
     
     
     
     
        
              
              
              
              
              
             Dim R As New ADODB.Recordset
             R.Open "select AYP FROM PEL WHERE EIDOS='" + gBuff + "' AND KOD='" + data1.Recordset("KOD") + "'", FDB2, adOpenForwardOnly, adLockReadOnly

             If R.EOF Then
                R.Close
                Exit Sub
             End If
                ayp = nNull(R(0))
               R.Close
               
               
              Dim SQL2 As String: SQL2 = "select sum(ISNULL(XREOSI,0)),sum(ISNULL(PISTOSI,0)) from EGG WHERE  HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND HME<'" + Format(DTPicker1.Value, "MM/DD/YYYY") + "' AND EIDOS='" + gBuff + "' and KOD='" + data1.Recordset("KOD") + "'"

               

420             Dim R22 As New ADODB.Recordset: R22.Open SQL2, FDB2, adOpenForwardOnly, adLockReadOnly

430             If IsNull(R22(0)) Then r0x = 0 Else r0x = R22(0)
440             If IsNull(R22(1)) Then r1p = 0 Else r1p = R22(1)

450             If gBuff = "e" Then
460                 ayp = ayp + r0x - r1p
                Else
470                 ayp = ayp + r1p - r0x
                End If
               
               myp2 = Replace(str(ayp), ",", ".")
               
               
               
               
               Dim cIDq As String
               If withID = 1 Then
                   cIDq = " ,ID,USERID AS[ΧΡΗΣΤΗΣ]  "
               Else
                   cIDq = " "
               End If
               
               
               
         Dim SQLL As String
           Dim fFILE As String
           
             fFILE = "TEMP" + Format(Now, "YYMMDDHHmm")
             
               
         If cXronies.Value = vbChecked Then
         
'                    R.Open "select AYP FROM PEL WHERE EIDOS='" + gBuff + "' AND KOD='" + Data1.Recordset("KOD") + "'", FDB2, adOpenForwardOnly, adLockReadOnly
'
'             If R.EOF Then
'                R.Close
'                Exit Sub
'             End If
'                ayp = nNull(R(0))
'               R.Close
         
         
         
         
             
             Dim SQL11 As String
             
                   SQLL = "SELECT convert(char(10),HME,103) as [HME],ATIM,AIT,isnull(XREOSI,0) AS [ΧΡΕΟΣΗ],isnull(PISTOSI,0) AS [ΠΙΣΤΩΣΗ],ROUND( " + myp2 + "+ sum(" + mxreo + ") over(order by  HME,ID ) ,2)    as [ΥΠΟΛΟΙΠΟ] " + cIDq + " into " + fFILE + "  FROM EGG WHERE  EIDOS='" + gBuff + "' AND KOD='" + data1.Recordset("KOD") + "' and   HME>='" + Format(DTPicker1.Value, "MM/DD/YYYY") + "' AND HME<='" + Format(DTPicker2.Value, "MM/DD/YYYY") + "'  "
             
             
             SQL11 = SQLL

             SQL11 = SQL11 + " UNION SELECT convert(char(10),HME,103) as [HME],ATIM,AIT,XREOSI AS [ΧΡΕΟΣΗ],PISTOSI AS [ΠΙΣΤΩΣΗ],ROUND( " + myp2 + "+ sum(" + mxreo + ") over(order by  HME,ID ) ,2)    as [ΥΠΟΛΟΙΠΟ] " + cIDq + " FROM  " + f_PALIAXRONIA + ".dbo.EGG WHERE  EIDOS='" + gBuff + "' AND KOD='" + data1.Recordset("KOD") + "' and   HME>='" + Format(DTPicker1.Value, "MM/DD/YYYY") + "' AND HME<='" + Format(DTPicker2.Value, "MM/DD/YYYY") + "'  "
             ' SELECT *  FROM " + f_PALIAXRONIA + ".dbo.EGGTIM  where " + filtroPEL.Caption + " HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "'  AND APOT=" + Label1.Caption + " and KODE='" + Text1.Text + "' "
             Dim mPROP As String
             mPROP = FINDPARAMETROI(1, "PAR1", "F_PROPERSI", "", "Γ.19 ΒΑΣΗ ΠΡΟΠΕΡΣΙΝΗΣ ΧΡΟΝΙΑΣ")
             SQL11 = SQL11 + " UNION " + " SELECT convert(char(10),HME,103) as [HME],ATIM,AIT,XREOSI AS [ΧΡΕΟΣΗ],PISTOSI AS [ΠΙΣΤΩΣΗ],ROUND( " + myp2 + "+ sum(" + mxreo + ") over(order by  HME,ID ) ,2)    as [ΥΠΟΛΟΙΠΟ] " + cIDq + "  FROM " + mPROP + ".dbo.EGG WHERE  EIDOS='" + gBuff + "' AND KOD='" + data1.Recordset("KOD") + "' and   HME>='" + Format(DTPicker1.Value, "MM/DD/YYYY") + "' AND HME<='" + Format(DTPicker2.Value, "MM/DD/YYYY") + "'  "
             Dim mPROP2 As String


             Dim mPROP3 As String
             mPROP3 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI2", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΙΝ 3 XPONIA")
             If Len(mPROP3) > 0 Then
                  SQL11 = SQL11 + " UNION " + " SELECT convert(char(10),HME,103) as [HME],ATIM,AIT,XREOSI AS [ΧΡΕΟΣΗ],PISTOSI AS [ΠΙΣΤΩΣΗ],ROUND( " + myp2 + "+ sum(" + mxreo + ") over(order by  HME,ID ) ,2)    as [ΥΠΟΛΟΙΠΟ] " + cIDq + "  FROM " + mPROP3 + ".dbo.EGG WHERE  EIDOS='" + gBuff + "' AND KOD='" + data1.Recordset("KOD") + "' and   HME>='" + Format(DTPicker1.Value, "MM/DD/YYYY") + "' AND HME<='" + Format(DTPicker2.Value, "MM/DD/YYYY") + "'  "
             End If

             Dim mPROP4 As String
             mPROP4 = FINDPARAMETROI(1, "PAR1", "F_PROPERSI3", "", "Γ.19 ΒΑΣΗ ΧΡΟΝΙΑΣ ΠΡΙΝ 4 XPONIA")
             If Len(mPROP4) > 0 Then
                SQL11 = SQL11 + " UNION " + " SELECT convert(char(10),HME,103) as [HME],ATIM,AIT,XREOSI AS [ΧΡΕΟΣΗ],PISTOSI AS [ΠΙΣΤΩΣΗ],ROUND( " + myp2 + "+ sum(" + mxreo + ") over(order by  HME,ID ) ,2)    as [ΥΠΟΛΟΙΠΟ] " + cIDq + "  FROM " + mPROP4 + ".dbo.EGG WHERE  EIDOS='" + gBuff + "' AND KOD='" + data1.Recordset("KOD") + "' and   HME>='" + Format(DTPicker1.Value, "MM/DD/YYYY") + "' AND HME<='" + Format(DTPicker2.Value, "MM/DD/YYYY") + "'  "
             End If
             
             Gdb.Execute SQL11
                 SQLL = "select * from " + fFILE + "   Order by SUBSTRING(HME,7,4)+SUBSTRING(HME,4,2)+SUBSTRING(HME,1,2) "
           Else
                  SQLL = "SELECT convert(char(10),HME,103) as [HME],ATIM,AIT,XREOSI AS [ΧΡΕΟΣΗ],PISTOSI AS [ΠΙΣΤΩΣΗ],ROUND( " + myp2 + "+ sum(" + mxreo + ") over(order by  HME,ID ) ,2)    as [ΥΠΟΛΟΙΠΟ],ATIM2 AS [ΠΛΗΡ] " + cIDq + " FROM EGG WHERE  EIDOS='" + gBuff + "' AND KOD='" + data1.Recordset("KOD") + "' and   HME>='" + Format(DTPicker1.Value, "MM/DD/YYYY") + "' AND HME<='" + Format(DTPicker2.Value, "MM/DD/YYYY") + "'  "
           End If
            Adodc1.RecordSource = SQLL
           
           
           
               'Adodc1.Recordset.MoveLast
               Adodc1.Refresh
               

Dim Col As TrueDBGrid80.Column
'With TDBGrid1
'
'    Set Col = .columns.ADD(8)
'
'    Col.Visible = True
'
'    Col.Caption = "Unbound"
'
'    .HoldFields
'
'    .Rebind
'
'End With


               
                   
     
     'Adodc1.Refresh
    Set RSClone = Adodc1.Recordset.Clone()
     
     
               
               
               
               
               
               
               On Error GoTo 0
               ' TDBGRID1.columns(0).FooterText = "Footer0"USERID as [Χρήστης]
               tdbgrid1.MoveLast
               ' TDBGRID1.columns(1).FooterText = "Footer1"
                   
                Dim r7 As New ADODB.Recordset
                r7.Open "SELECT SUM(ISNULL(XREOSI,0)) AS SX,SUM(ISNULL(PISTOSI,0) ) AS SP FROM EGG  WHERE  EIDOS='" + gBuff + "' AND KOD='" + data1.Recordset("KOD") + "' and   HME>='" + Format(DTPicker1.Value, "MM/DD/YYYY") + "' AND HME<='" + Format(DTPicker2.Value, "MM/DD/YYYY") + "'", FDB2, adOpenDynamic, adLockOptimistic
                Dim sxre As Single, spis As Single
                sxre = nNull(r7(0))
                spis = nNull(r7(1))
                Dim l2 As Integer
                 With tdbgrid1
                      .HeadLines = 2
                       tdbgrid1.HeadFont.Size = CMDEXIT.FontSize
                       tdbgrid1.FooterFont.Size = CMDEXIT.FontSize
                      
                             .columns(0).Width = 1100
                             .columns(2).Width = 3100
                             .columns(0).Caption = "ΗΜΕΡ/ΝΙΑ" + vbCr + " "
                             .columns(1).Caption = "ΠΑΡ/ΚΟ" + vbCr + " "
               
                     If gBuff = "r" Then
                          
                             .columns(2).Caption = "AITIOΛΟΓΙΑ" + vbCr + "ΠΡΟΗΓ.ΥΠΟΛ"
                             .columns(3).Caption = "ΧΡΕΩΣΗ" + vbCr + Format(0, "######0.00")
                             .columns(4).Caption = "ΠΙΣΤΩΣΗ" + vbCr + Format(Val(myp2), "######0.00")
                 
                     Else
                             .columns(2).Caption = "AITIOΛΟΓΙΑ" + vbCr + "ΠΡΟΗΓ.ΥΠΟΛ"
                             .columns(3).Caption = "ΧΡΕΩΣΗ" + vbCr + Format(Val(myp2), "######0.00")
                             .columns(4).Caption = "ΠΙΣΤΩΣΗ" + vbCr + Format(0, "######0.00")
                     End If
                      .columns(5).Caption = "ΥΠΟΛΟΙΠΟ" + vbCr + " "
                      '.columns(6).Caption = "ΧΡΗΣΤΗΣ" + vbCr + " "
                      '.columns(6).Caption = "ΙD" + vbCr + " "
                 End With
                 
                 
                
                 tdbgrid1.columns(3).FooterText = Format(sxre, "######0.00")
                 tdbgrid1.columns(4).FooterText = Format(spis, "######0.00")
                 
                 
                  tdbgrid1.columns(1).Width = 1200
                  tdbgrid1.columns(0).Width = 1200
                  
                   If withID = 1 Then
                  
               Else
                  tdbgrid1.columns(2).Width = 2500
                  tdbgrid1.columns(3).Width = 1200
                  tdbgrid1.columns(4).Width = 1200
                  
               End If
               tdbgrid1.MoveLast
               Dim YPOL5 As Single
             ' If gVal(TDBGrid1.columns(5).Text) > 0 Then
                  Dim YpolApl As Single
                    YpolApl = gVal(tdbgrid1.columns(5).Text)
                    Dim napl As Integer
                    Dim NNN As Integer
                    For napl = tdbgrid1.row To 0 Step -1
                      
                        If data1.Recordset("eidos") = "r" Then NNN = 4 Else NNN = 5
                            YPOL5 = YpolApl
                            YpolApl = YpolApl - gVal(tdbgrid1.columns(NNN).Text)
                       ' If Left(TDBGrid1.columns(6).Text, 1) <> "*" Then
                              ' ENHMERVNEI THN DATABASE
                             
                        'Else
                              
                        'End If
                        If YpolApl <= 0 Then
                           If YPOL5 > 0 Then
                              tdbgrid1.columns(6).Text = Format(YPOL5, "###0") ' " "
                           Else
                              tdbgrid1.columns(6).Text = ""
                           End If
                              
                        Else
                              tdbgrid1.columns(6).Text = "*" ' + TDBGrid1.columns(8).Text
                        End If
                        tdbgrid1.MovePrevious
                    Next
                     tdbgrid1.MoveLast
             ' End If
If cXronies.Value = vbChecked Then
   Gdb.Execute "DROP TABLE " + fFILE

End If
 
                  
                  
                 
                'TDBGRID1.columns(3).FooterFont.Size = 12 ' = Format(SYNOLA(k), "######0.00")

End Sub

Private Sub DBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

        '<EhHeader>
        On Error GoTo DBGrid2_RowColChange_Err

        '</EhHeader>
        Dim DUM

100     If PROTH_FORA Then Exit Sub

        '-------------------------------------
           Dim AFMPEL As String
           Dim myp2 As String
          AFMPEL = GGET_CVALUE("SELECT TOP 1 AFMPEL FROM MEM")
          
          
          
          
          
           If F_NEA_KRTELLA = 1 Then
                tdbgrid1.Visible = True
                neakartela 1
                'δεν κανω εξοδο για να σχεδιασει και το flexgrid γιατι δεν δουλευει η καρτελλα σε excel
                'Exit Sub
              Else
              
              tdbgrid1.Visible = False
                'Exit Sub
                
                
                
           End If
                
                
                
                
                
                
                
                
        '-------------------------------------
                

        On Error Resume Next

110     Label2.Caption = Left(data1.Recordset!kod, 5) + " " + data1.Recordset!EPO
120     Label3.Caption = data1.Recordset!DIE
130     Label4.Caption = data1.Recordset!THL

        On Error GoTo DBGrid2_RowColChange_Err

140     Command5.Caption = "ΠΡΟΗΓΟΥΜΕΝΗ ΧΡΗΣΗ"

        If f_KENTRA_ADYNATISMATOS = 1 Then
        
            If gXEIRISTHS = 1 Then
               Exit Sub
            End If
            
        
   
            ' data1.Refresh
            Dim R5  As New ADODB.Recordset

            Dim NR5 As Long
   
            If IsNull(data1.Recordset("HM1")) Then
                NR5 = 0
            Else
                R5.Open "SELECT COUNT(*) FROM EGG WHERE  XREOSI+PISTOSI=0 AND HME>='" + Format(data1.Recordset("HM1"), "MM/DD/YYYY") + "' AND EIDOS='" + gBuff + "' AND KOD='" + data1.Recordset("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic
                NR5 = R5(0)
                R5.Close

                On Error Resume Next

                If Year(data1.Recordset("HM1")) < Year(gLhjh) Then
                    FCONNECTIONPALIA.Open gConnect + ";DATABASE=" + f_PALIAXRONIA
                    R5.Open "SELECT COUNT(*) FROM EGG WHERE  XREOSI+PISTOSI=0 AND HME>='" + Format(data1.Recordset("HM1"), "MM/DD/YYYY") + "' AND EIDOS='" + gBuff + "' AND KOD='" + data1.Recordset("KOD") + "'", FCONNECTIONPALIA, adOpenDynamic, adLockOptimistic
                    NR5 = NR5 + R5(0)
                    R5.Close
                    FCONNECTIONPALIA.Close
                End If
      
            End If
   
            If IsNull(data1.Recordset("num2")) Then
                data1.Recordset("num2") = 0
            End If

            ypolmhx.Caption = data1.Recordset("num2") - NR5 ' DATA1.Recordset("num5")
   
            If IsNull(data1.Recordset("num4")) Then
                ypolposo.Caption = data1.Recordset("num1")
            Else
                ypolposo.Caption = data1.Recordset("num1") - data1.Recordset("num4")
            End If

        End If

150     DUM = KARTELLA(gDir)
        'Command5.Enabled = True
        ' Command4.SetFocus

        '<EhFooter>
        Exit Sub

DBGrid2_RowColChange_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.DBGrid2_RowColChange " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
                            SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.DBGrid2_RowColChange " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function KARTELLA(mgdir)

        '===============
        'Καρτελλα πελάτη
        '===============
        ' DBGRid2.height = 1215
        '<EhHeader>
        On Error GoTo KARTELLA_Err

        '</EhHeader>

        Dim X2 As String, s As String

        Dim yp As Single

        Dim SX As Double

        Dim SP As Double

        Dim DB As Database

        'Dim R As Recordset
        Dim XREOSI, PISTOSI
        Dim MAFM As String
        MAFM = Trim(Text2(2).Text)

        Dim ayp As Single

        Dim k   As Long

        Dim sql

        Dim L As Long, R As New ADODB.Recordset

        ' If Option1(1) Then X2 = "e" Else X2 = "r"
120     X2 = mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1)

        'If gXEIRISTHS = 8 Or gXEIRISTHS = 9 Or gXEIRISTHS = 99 Then   'auanasia blepei kinhseis apo 8
        '   If X2 = "e" Then
        '      X2 = "e','E"
        '   Else
        '      X2 = "r','R"
        '   End If
        'End If

130     If data1.Recordset.RecordCount = 0 Then Exit Function
140     If IsNull(data1.Recordset("kod")) Then Exit Function

        On Error GoTo KARTELLA_Err

        'ΥΠΟΛΟΓΙΖΕΙ ΤΑ ΠΡΟΟΔΕΥΤΙΚΑ
150     s = data1.Recordset("kod")

        'EGG.DatabaseName = mgdir
        
          If CH_SYGK.Value = vbChecked Then
             If Check1.Value = vbChecked Or Check2.Value = vbChecked Then
                MsgBox "Συγκεντρωτική καρτέλλα μόνο στην τρέχουσα χρήση"
                Exit Function
             End If
          End If
          

160     If gUserId = 2 Then    'ταμίας καταστήματος
170         If s = "9990" Then
180             R.Open "select * FROM EGG ORDER BY ID", Gdb, adOpenDynamic, adLockReadOnly

190             R.MoveLast
200             DTPicker1.Value = R("HME")
210             R.Close

220             If DTPicker1 < Now Then
230                 DTPicker1 = Now
                End If
            End If
        End If

        '===========  TO PARAKATO EINAI GIA 2 XRONIES
240     If Check1.Value = vbChecked Then    '   <> gDir Then ' palia xronia

250         egg.ConnectionString = gConnect + ";DATABASE=" + f_PALIAXRONIA    ' "DSN=" + LTrim(Str(f_PALIAXRONIA))
            '   gConnect ";DATABASE=" + f_PALIAXRONIA
260         R.Open "select AYP FROM PEL WHERE EIDOS='" + gBuff + "' AND KOD='" + s + "'", gConnect + ";DATABASE=" + f_PALIAXRONIA, adOpenForwardOnly, adLockReadOnly

270         If R.EOF Then
280             R.Close

                Exit Function

            End If

290         ayp = R(0)
300     ElseIf Check2.Value = vbChecked And Len(XRONIES.Text) > 2 Then  '   <> gDir Then ' palia xronia

            ' XRONIES.AddItem Left(R("TITLOS") + Space(20), 20) + R("DSN")
310         f_PROPERSI = Trim(mID(XRONIES.Text, 21, 20))
    
            '    If Len(f_PROPERSI) = 0 Then
            '       MsgBox "ΔΙΑΛΕΞΤΕ ΧΡΟΝΙΑ"
            '      Exit Function
            '  End If
    
320         egg.ConnectionString = gConnect + ";DATABASE=" + f_PROPERSI    ' "DSN=" + LTrim(Str(f_PALIAXRONIA))

            '   gConnect ";DATABASE=" + f_PALIAXRONIA
            On Error Resume Next

330         R.Open "select AYP FROM PEL WHERE EIDOS='" + gBuff + "' AND KOD='" + s + "'", gConnect + ";DATABASE=" + f_PROPERSI, adOpenForwardOnly, adLockReadOnly

340         If R.EOF Then
350             R.Close

                Exit Function

            End If

360         ayp = R(0)
        Else
370         egg.ConnectionString = gConnect
380         yp = IIf(IsNull(data1.Recordset("AYP")), 0, data1.Recordset("AYP"))

            If CH_SYGK.Value = vbChecked Then
                 If Len(MAFM) < 9 Then
                      MsgBox "Συμπληρώστε το ΑΦΜ"
                      Exit Function
                 End If
                 
                 yp = GGET_NVALUE("select sum(ISNULL(AYP,0)) FROM PEL WHERE AFM='" + MAFM + "' AND EIDOS='" + gBuff + "'")
            End If
390         ayp = yp

            Dim r0, r1

400         If DTPicker1 = gEnarjh Then    ' CDate("01/01/" + Str(Year(Now))) Then
            Else                                                    '"01/01/" + LTrim(Str(Year(Now)))
410             sql = "select sum(XREOSI),sum(PISTOSI) from EGG WHERE  HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND HME<'" + Format(DTPicker1.Value, "MM/DD/YYYY") + "' AND EIDOS in ('" + X2 + "') and KOD='" + s + "'"

                On Error GoTo ERRORDATA

420             R.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

430             If IsNull(R(0)) Then r0 = 0 Else r0 = R(0)
440             If IsNull(R(1)) Then r1 = 0 Else r1 = R(1)

450             If InStr(1, X2, "e") > 0 Then
460                 ayp = ayp + r0 - r1
                Else
470                 ayp = ayp + r1 - r0
                End If

            End If

            On Error Resume Next

480         R.Close

        End If

        '===========  TO PARAKATO EINAI GIA 2 XRONIES

        'yp = IIf(IsNull(data1.Recordset("AYP")), 0, data1.Recordset("AYP"))

490     MSFlexGrid1.Visible = False

        ' EGG.RecordSource = "select *from EGG where EIDOS='" + x2 + "' and KOD='" + s + "' order by HME,ATIM; "
        ' EGG.Refresh

        'Set R = EGG.Recordset
        'GoTo 200

        '  If EGG.Recordset.RecordCount = 0 Then Exit Function
        '
        '  EGG.Recordset.MoveFirst
        '  On Error GoTo KARTELLA_Err
        '
        '
        '   SX = 0
        '   SP = 0
        '  'ENHMEΡΩΝΩ TA ΠΡΟΟΔΕΥΤΙΚΑ
        'Do While Not EGG.Recordset.EOF
        '  yp = yp + EGG.Recordset("XREOSI")
        '  yp = yp - EGG.Recordset("PISTOSI")
        '
        '  SX = SX + EGG.Recordset("XREOSI")
        '  SP = SP + EGG.Recordset("PISTOSI")
        '  EGG.Recordset("prood") = yp
        '  EGG.Recordset.Update
        '  EGG.Recordset.MoveNext
        'Loop
        '
        '

        'Exit Sub
        'If DATA1.Recordset.RecordCount > 0 Then
        '  If IsNull(DATA1.Recordset("kod")) Then Exit Sub

500     s = data1.Recordset("kod")
        'If x2 = "e" Then

        Dim N                As Integer

        Dim SELE_PARASTATIKA As String

510     SELE_PARASTATIKA = ""

520     For N = 0 To List1.ListCount - 1

            'On Error Resume Next
530         If List1.Selected(N) Then
540             SELE_PARASTATIKA = SELE_PARASTATIKA + "'" + Right(List1.List(N), 1) + "',"
            End If

        Next

550     If Len(SELE_PARASTATIKA) > 0 Then
560         SELE_PARASTATIKA = Left(SELE_PARASTATIKA, Len(SELE_PARASTATIKA) - 1)
570         SELE_PARASTATIKA = " left (ATIM,1) IN (" + SELE_PARASTATIKA + ") AND "
        End If

        Dim FILTRO As String
    
580     If Len(Trim(Text1.Text)) > 0 Then
590         FILTRO = " ATIM LIKE '" + Text1.Text + "%' AND "
        Else
600         FILTRO = ""
        End If

610     If Len(Trim(Text3.Text)) > 0 Then
620         FILTRO = FILTRO + " AIT LIKE '" + Text3.Text + "%' AND "
        Else
630         FILTRO = FILTRO + ""
        End If

640     If Check1.Value = vbChecked Then    '   <> gDir Then ' palia xronia
650         egg.RecordSource = "select CONVERT(CHAR(10),HME,3) as [Ημερoμηνία], EID as [Είδ], APA as [Αριθ_Παρ],ATIM as [Πελατ],AIT as [Αιτιολογία],XREOSI as [Χρέωση],PISTOSI as [Πίστωση],PROOD as [Υπόλοιπο],USERID as [Χρήστης],ID from EGG where " + FILTRO + SELE_PARASTATIKA + " EIDOS in ('" + X2 + "') and KOD='" + s + "' order by HME,ID;"
660     ElseIf Check2.Value = vbChecked Then    '   <> gDir Then ' palia xronia
670         egg.RecordSource = "select CONVERT(CHAR(10),HME,3) as [Ημερoμηνία], EID as [Είδ], APA as [Αριθ_Παρ],ATIM as [Πελατ],AIT as [Αιτιολογία],XREOSI as [Χρέωση],PISTOSI as [Πίστωση],PROOD as [Υπόλοιπο],USERID as [Χρήστης],ID from EGG where " + FILTRO + SELE_PARASTATIKA + " EIDOS in ('" + X2 + "') and KOD='" + s + "' order by HME,ID;"

        Else
            If CH_SYGK.Value = vbChecked Then
               egg.RecordSource = "select CONVERT(CHAR(10),HME,3) as [Ημερoμηνία], EID as [Είδ], APA as [Αριθ_Παρ],ATIM as [Πελατ],AIT as [Αιτιολογία],XREOSI as [Χρέωση],PISTOSI as [Πίστωση],PROOD as [Υπόλοιπο]," & "USERID as [Χρήστης],ID from EGG where XREOSI+PISTOSI<>0 AND " + FILTRO + SELE_PARASTATIKA + " HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME<='" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "' AND EIDOS in ('" + X2 + "') and KOD in (select KOD FROM PEL WHERE AFM='" + MAFM + "' AND  EIDOS in ('" + X2 + "')  )  order by HME,ID;"
            Else
680            egg.RecordSource = "select CONVERT(CHAR(10),HME,3) as [Ημερoμηνία], EID as [Είδ], APA as [Αριθ_Παρ],ATIM as [Πελατ],AIT as [Αιτιολογία],XREOSI as [Χρέωση],PISTOSI as [Πίστωση],PROOD as [Υπόλοιπο]," & "USERID as [Χρήστης],ID from EGG where " + FILTRO + SELE_PARASTATIKA + " HME>='" + Format(DTPicker1, "MM/DD/YYYY") + "' AND HME<='" + Format(DateAdd("d", 1, DTPicker2), "MM/DD/YYYY") + "' AND EIDOS in ('" + X2 + "') and KOD='" + s + "' order by HME,ID;"
            End If
        End If

        'Else
        '  EGG.RecordSource = "select ATIM as [Αρ_Παρ],HME as [Ημερ/νία],EID as [Είδ],APA as [Αριθ_Παρ],AIT as [Αιτιολογία],XRE as [Χρέωση],XRE as [Πίστωση],PROOD as [Υπόλοιπο] from EGG where EIDOS='" + x2 + "' and KOD='" + s + "' order by HME,ATIM;"
        'End If
        'EGG.DatabaseName = mgdir

        '  EGG.Recordset.Close
690     egg.Refresh


'EMFANIZE ANOMALIA STIN KARTELA KAI TO EKANA PROAIRETIKO
If f_2000MORE_ROWS = 0 Then
   GoTo 700
End If




'ΟΙ ΠΑΡΑΚΑΤΩ 12 ΣΕΙΡΕΣ ΕΙΝΑΙ ΓΙΑ ΝΑ ΒΓΑΖΕΙ ΠΑΝΩ ΑΠΟ 2000 ΣΕΙΡΕΣ ΤΟ GRID
'================================================================================
'' Assuming that rs is your ADO recordset  MSFlexGrid1
   MSFlexGrid1.rows = egg.Recordset.RecordCount + 1
Dim rsVar As Variant
   rsVar = egg.Recordset.GetString(adClipString, egg.Recordset.RecordCount)
'
   MSFlexGrid1.ColS = egg.Recordset.FIELDS.Count
'
'   ' Set column names in the grid
   Dim i As Long
   For i = 0 To egg.Recordset.FIELDS.Count - 1

       MSFlexGrid1.TextMatrix(0, i) = egg.Recordset.FIELDS(i).Name

   Next
'
   MSFlexGrid1.row = 1
   MSFlexGrid1.Col = 0
'
'   ' Set range of cells in the grid
   MSFlexGrid1.RowSel = MSFlexGrid1.rows - 1
   MSFlexGrid1.ColSel = MSFlexGrid1.ColS - 1
   MSFlexGrid1.clip = rsVar
'================================================================================






        '  yp = 0
700     MSFlexGrid1.Refresh

        On Error Resume Next

        'MSFlexGrid1.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "ΠΡΟΗΓ.ΥΠΟΛΟΙΠΟ" & Chr(9) & Chr(9) & Chr(9) & Format(ayp, "######.00"), 1

710     MSFlexGrid1.ColWidth(4) = 3000
720     MSFlexGrid1.ColWidth(3) = 1500    'ΑΠΑ

730     MSFlexGrid1.ColWidth(2) = 600    'ΕΙΔ

740     MSFlexGrid1.ColWidth(1) = 600

        Dim SU, NX, NP, NY

750     NX = 5: NP = 6: NY = 7

        Dim synt As Single

760     If InStr(X2, "r") > 0 Then
770         If F_PIST = 0 Then
780             SU = ayp
            Else
790             SU = -ayp
            End If
      
800         synt = -1    ' για να μην βγαζει αρνητικα στα πιστωτικά υπόλοιπα σε προμηθευτές
    
            MSFlexGrid1.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "ΠΡΟΗΓ.ΥΠΟΛΟΙΠΟ" & Chr(9) & Chr(9) & Chr(9) & Format(SU, "######.00"), 1
    
        Else
810         SU = ayp
820         synt = 1
    
830         MSFlexGrid1.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "ΠΡΟΗΓ.ΥΠΟΛΟΙΠΟ" & Chr(9) & Chr(9) & Chr(9) & Format(synt * SU, "######.00"), 1
        End If

840     SX = 0: SP = 0

        On Error GoTo KARTELLA_Err ' debug

        Dim sxr As Single, spis As Single

        'MSFlexGrid1.Clear
850     For k = 2 To MSFlexGrid1.rows - 1

860         MSFlexGrid1.TextMatrix(k, 1) = Format(MSFlexGrid1.TextMatrix(k, 1), "dd/mm/yyyy")

870         If Val(MSFlexGrid1.TextMatrix(k, NX)) <> 0 Then
880             MSFlexGrid1.TextMatrix(k, NP) = ""
890             MSFlexGrid1.TextMatrix(k, NX) = Format(gVal(MSFlexGrid1.TextMatrix(k, NX)), "#####.00")
            End If

900         If Val(MSFlexGrid1.TextMatrix(k, NP)) <> 0 Then
910             MSFlexGrid1.TextMatrix(k, NP) = Format(gVal(MSFlexGrid1.TextMatrix(k, NP)), "#####.00")
920             MSFlexGrid1.TextMatrix(k, NX) = " "
            End If

930         MSFlexGrid1.TextMatrix(k, NY) = Format(synt * gVal(MSFlexGrid1.TextMatrix(k, NY)), "#####.00")
940         SX = SX + Val(DOT(MSFlexGrid1.TextMatrix(k, NX)))
950         SP = SP + Val(DOT(MSFlexGrid1.TextMatrix(k, NP)))
            ' SU = SU + SX - SP
            If f_KENTRA_ADYNATISMATOS = 1 Then
                MSFlexGrid1.TextMatrix(k, NY + 1) = k - 1
                
            End If
            
960         If Len(SELE_PARASTATIKA) = 0 Then

970             If X2 = "r" Then
980                 MSFlexGrid1.TextMatrix(k, NY) = Format((SU - SX + SP), "#####.00")
                Else
990                 MSFlexGrid1.TextMatrix(k, NY) = Format(synt * (SU + SX - SP), "#####.00")
                End If

            End If       'MSFLEXGRID1.TextMatrix(K, nT) = Format(Val(MSFLEXGRID1.TextMatrix(K, nT)), "######.00")

            ' ΚΑΝΩ ΤΟ GRID ΡΙΓΕ
1000        If k Mod 2 = 0 Then

1010            For L = 0 To MSFlexGrid1.ColS - 1
1020                MSFlexGrid1.row = k: MSFlexGrid1.Col = L: MSFlexGrid1.CellBackColor = &HC0E0FF
                Next

            End If

        Next

            If f_KENTRA_ADYNATISMATOS = 1 Then
                MSFlexGrid1.TextMatrix(0, NY + 1) = "Α/Α"
                
            End If




1030    MSFlexGrid1.AddItem ("" & Chr(9) & "Σύνολα" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(SX, "######.00") & Chr(9) & Format(SP, "######.00") & Chr(9) & MSFlexGrid1.TextMatrix(k - 1, NY))
        ' MSFlexGrid1.TextMatrix(k + 1, NY) = Format(synt * Val(MSFlexGrid1.TextMatrix(k, NY)), "#####.00")

1040    MSFlexGrid1.TopRow = Max(1, MSFlexGrid1.rows - 10)
1050    MSFlexGrid1.ColAlignment(5) = 6    ' RIGHT
1060    MSFlexGrid1.ColAlignment(6) = 6    ' RIGHT
1070    MSFlexGrid1.ColAlignment(7) = 6    ' RIGHT
1080    MSFlexGrid1.Visible = True
        'End If

        Exit Function

ERRORDATA:
1090    MsgBox "αδυνατη ανεύρεση"

        '<EhFooter>
        Exit Function

KARTELLA_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.KARTELLA " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.KARTELLA " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub DBGRID22_Click()

' ΒΓΑΖΕΙ ΤΕΧΤΒΟΧ
 '<EhHeader>
      '  On Error GoTo DBGRID22_Click_Err

        '</EhHeader>

100     timtext3.Text = DBGRID22.Text
110     timtext3.BackColor = DBGRID22.BackColor
120     timtext3.Visible = False

        ' ΤΟ ΤΕΧΤ3 ΠΑΡΑΛΑΜΒΑΝΕΙ ΤIΣ ΔΙΑΣΤΑΣΕΙΣ ΑΠΟ ΤΟ GRID
130     timtext3.Top = DBGRID22.Top + DBGRID22.CellTop
140     timtext3.Left = DBGRID22.Left + DBGRID22.CellLeft
150     timtext3.Width = DBGRID22.CellWidth - 15
160     timtext3.Height = DBGRID22.CellHeight - 25
170     timtext3.Visible = True

        DBGRID22.Col = 2




      
        '</EhFooter>

End Sub

Private Sub DBGRID22_DblClick()
 
       
        
       
        '<EhFooter>
        Exit Sub

DBGRID22_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.DBGRID22_Click " & _
         "at line " & Erl, _
         vbExclamation vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.DBGRID22_Click " & "at line " & Erl

        Resume Next

'
'        '<EhHeader>
'        On Error GoTo DBGRID22_DblClick_Err
'
'        '</EhHeader>
'
'100     DBGRID22.Col = 2
'110     APOT2.kod.Text = DBGRID22.Text   ' "10"
'120     APOT2.SHOW
'130     APOT2.kod.SetFocus
'
'        '<EhFooter>
'        Exit Sub
'
'DBGRID22_DblClick_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.pelat4.DBGRID22_DblClick " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.DBGRID22_DblClick " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>

End Sub

Private Sub DBGRID22_LostFocus()
    '  If TIMText3.SetFocus Then
    '  TIMText3.Visible = False
End Sub

Private Sub DBGRID22_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             Y As Single)

        '<EhHeader>
        On Error GoTo DBGRID22_MouseUp_Err
       DBGRID22.Col = 2
        '</EhHeader>
        
        
       'ΑΝ ΠΡΟΚΕΙΤΑΙ ΓΙΑ ΔΕΛΤΙΑ ΑΠΟΣΤΟΛΗΣ
'       If IsDate(DBGRID22.Text) Then
 '         PAR2SHOWING DBGRID22.row
  '        Exit Sub
   '    Else
       
100     If Button = 2 Then
110

            Dim FF As New APOT2

120         FF.kod.Text = DBGRID22.Text   ' "10"
130         FF.SHOW
140         FF.kod.SetFocus
        End If
    '   End If
       
        '<EhFooter>
        Exit Sub

DBGRID22_MouseUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.DBGRID22_MouseUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.DBGRID22_MouseUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Activate()
     f_resizeCount = 0
     
        '<EhHeader>
        On Error GoTo Form_Activate_Err

        '</EhHeader>

100     LAST_TIM = -1
        ' data1.DatabaseName = gDir

        ' egg.DatabaseName = gDir
        'egg.RecordSource = "select hme,EID,APA,AIT,iif(xre>0,xre,0) as xreosi,iif(xre<0,-xre,0) as pistosi,PROOD FROM EGG"

        'data1.ConnectTIONSTRING = gConnect

        ' Data1.RecordSource = "SELECT *FROM PEL"
110     Text2(0).SetFocus



    Set Prospective = tdbgrid1.Styles.Add("Prospective")

    Prospective.Font.Italic = True

    Prospective.Font.Bold = True

    Prospective.ForeColor = vbBlue

    

    Set Distributors = tdbgrid1.Styles.Add("Distributors")

    Distributors.BackColor = vbRed

    Distributors.ForeColor = vbWhite


'---------------------------------------------------------
Dim ThisControl As Control
  
  ReDim Preserve ControlInfos(0 To 0)
  ControlInfos(0).Width = Me.Width
  ControlInfos(0).Height = Me.Height
  For Each ThisControl In Me.Controls
    ReDim Preserve ControlInfos(0 To UBound(ControlInfos) + 1)
    On Error Resume Next  ' hack to bypass controls with no size or position properties
    With ControlInfos(UBound(ControlInfos))
      .Left = ThisControl.Left
      .Top = ThisControl.Top
      .Width = ThisControl.Width
      .Height = ThisControl.Height
      .FontSize = ThisControl.FontSize
    End With
    On Error GoTo 0
  Next
'-----------------------------------------------------------------








        '<EhFooter>
        Exit Sub

Form_Activate_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Form_Activate " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Form_Activate " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Sub KARTELLA2()
'
''============================================================================
'''                          Καρτελλα πελάτη
'''============================================================================
''WWW.DELLPARTNERDIRECT.COM / MARKETING AND SALES TOOL / DELL CAMPAIGN BUILDER
''WEB BANNER /
''Dim x2 As String, s As String
''Dim yp As Single
''Dim SX As Single
''Dim SP As Single
''  27342  karakizos
''If Data1.Recordset.RecordCount > 0 Then
''
''  If IsNull(Data1.Recordset("kod")) Then Exit Sub
''   Me.MousePointer = vbHourglass
''   s = Data1.Recordset("kod")
''  If Option1(1) Then x2 = "e" Else x2 = "r"
''  EGG.RecordSource = "select ATIM as [Αρ_Παρ],hme as [Ημερ/νία],EID as [Είδ],APA as [Αριθ_Παρ],AIT as [Αιτιολογία],iif(xre>0,xre,0) as [Χρέωση],iif(xre<0,-xre,0) as [Πίστωση],PROOD as [Υπόλοιπο] from egg where eidos='" + x2 + "' and kod='" + s + "' order by hme,atim;"
''  EGG.Refresh
''
''
''  yp = 0
''  On Error Resume Next
''
''  EGG.Recordset.MoveFirst
''  SX = Data1.Recordset("ayp")
''  SP = 0
''Do While Not EGG.Recordset.EOF
''  yp = yp + EGG.Recordset("Χρέωση")
''  yp = yp - EGG.Recordset("Πίστωση")
''
''  SX = SX + EGG.Recordset("Χρέωση")
''  SP = SP + EGG.Recordset("Πίστωση")
''
''  EGG.Recordset.Edit
''  EGG.Recordset("Υπόλοιπο") = yp
''  EGG.Recordset.Update
''
''  EGG.Recordset.MoveNext
''Loop
''
''  EGG.Refresh
''MSFlexGrid1.ColWidth(4) = 2000
''MSFlexGrid1.ColWidth(3) = 600 'ΑΠΑ
''MSFlexGrid1.ColWidth(2) = 600 'ΕΙΔ
''
''
''MSFlexGrid1.AddItem ("" & Chr(9) & "Σύνολα" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(SX, "###,###.00") & Chr(9) & Format(SP, "###,###.00"))
''
''
''MSFlexGrid1.row = 1
''MSFlexGrid1.Col = 1
''MSFlexGrid1.Text = "ΑΡΧΙΚΟ ΥΠΟΛΟΙΠΟ"
''
''MSFlexGrid1.Col = 5
''MSFlexGrid1.Text = Data1.Recordset("AYP")
''
''
''
''MSFlexGrid1.TopRow = MSFlexGrid1.Rows - 10
''
''
''End If
''Me.MousePointer = vbNormal
''
'End Sub

Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "PELAT4"
110     PARAMETROI.SHOW 1

        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Form_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then CMDEXIT_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyUp_Err

        '</EhHeader>
100     If KeyCode = 27 Then
110         Text2(0).SelStart = 0
120         Text2(0).SelLength = Len(Trim(Text2(0).Text))
130         Text2(0).SetFocus
        End If

        '<EhFooter>
        Exit Sub

Form_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Form_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Form_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
        
        
      '  F_FIRSTRESIZE = 1

        mForm_Load Me, fh, fw, ft, fl

        Dim k As Integer

        'El.mForm_Load Me
        
        tdbgrid1.MarqueeStyle = 3  'SEIRA MARKARISMENH
          tdbgrid1.Splits(0).AlternatingRowStyle = True    'ENALAX XROMATA SEIRON
          tdbgrid1.Splits(0).ExtendRightColumn = True
          tdbgrid1.FilterBar = False
 F_EKTYPOTHS = FINDPARAMETROI(1, "PAR1", "F_EKTYPOTHS", "..", "Εκτυπωτής παραστατικών(με .. προεπιλεγμένος)") '
        
        F_MONOENERGOS = Val(FINDPARAMETROI(1, "PAR1", "F_MONOENERGOS", "0", "ΔΕΙΧΝΕΙ ΜΟΝΟ ΤΟΥΣ ΕΝΕΡΓΟΥΣ=1 ΟΛΟΥΣ=0")) '
        
        
       
        
          F_PLATOS_FORMAS = FINDPARAMETROI(1, "PELAT4", "F_PLATOS_FORMAS", "16000", "Πλάτος φόρμας")
     
        
        F_YCOS_FORMAS = FINDPARAMETROI(1, "PELAT4", "F_YCOS_FORMAS", "9000", "Υψος φόρμας")
        
        
        
        
        f_pedioAFM = FINDPARAMETROI(1, "PELAT2", "f_pedioAFM", "CH1", "Ποιό πεδίο θα αναζητάται μαζί με το ΑΦΜ")
        
           f_LabAFM = FINDPARAMETROI(1, "PELAT2", "f_LabAFM", "ΑΦΜ /κειμ1", "Tίτλος ετικέτας ΑΦΜ")
      Label1(2).Caption = f_LabAFM
 
        

100     PELPROM.AddItem "ΠΕΛΑΤΗΣ"
110     PELPROM.AddItem "ΠΡΟΜΗΘΕΥΤΗΣ"
120     PELPROM_STRING = "er"

        f_msynt = ""
        f_PROPERSI = FINDPARAMETROI(1, "PAR1", "F_PROPERSI", "2005", "Γ.19 ΒΑΣΗ ΠΡΟΠΕΡΣΙΝΗΣ ΧΡΟΝΙΑΣ")
        f_PALIAXRONIA = Trim(FINDPARAMETROI(1, "PAR1", "F_PALIAXRONIA", "2005", "Γ.18 DSN ΠΑΛΙΑΣ ΧΡΟΝΙΑΣ"))

        'If ffbuff.Caption = "" Then
130     PELPROM.ListIndex = 0

        ' pelat4.SHOW
        'Me.height = 9260

        ' El.Init pelat4

140     f_telephone = Val(FINDPARAMETROI(1, "MDIFORM1", "F_TELEPHONE", "0", "ΕΜΦΑΝΙΣΗ ΚΑΛΟΥΝΤΩΝ ΤΗΛΕΦΩΝΩΝ=1"))

150     If f_telephone = 1 Then
160         THLEFONA.Visible = True
        End If

        Dim R As New ADODB.Recordset

170     R.Open "select DSN,TITLOS FROM ETAIREIES ORDER BY TITLOS DESC", Gdb, adOpenDynamic, adLockOptimistic

180     If Not R.EOF Then
185         R.MoveFirst
190         XRONIES.Clear

200         Do While Not R.EOF

210             If Not IsNull(R("TITLOS")) And (Not IsNull(R("DSN"))) Then
220                 XRONIES.AddItem Left(R("TITLOS") + Space(20), 20) + R("DSN")
225             End If

230             R.MoveNext
232         Loop

235     End If

240     R.Close

        'PARAMETROI.PARAM.Caption = "PELAT2"
        '  PARAMETROI.SHOW 1

250     Me.KeyPreview = True
260     f_direct = 0
    
270     f_direct = Val(FINDPARAMETROI(1, "PELAT4", "F_DIRECT", "0", "0=ΕΚΤΥΠΩΣΗ ΑΠΛΗ  2=ΕΚΤΥΠΩΣΗ ME CRYSTAL REPORTS"))
    
280     F_PIST = Val(FINDPARAMETROI(1, "PELAT4", "F_PIST", "0", "0=to υπολοιπο των προμηθευτών είναι πιστωτικό 1=χρεωστικό"))
290     F_CanDelete = Val(FINDPARAMETROI(1, "BOHU11", "F_CanDelete", "1", "μπορω να σβήνω κινήσεις"))

        Dim dum2

300     f_buff = "e"

        'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

        ' On Error GoTo Form_Load_Err
310     dum2 = ADD_FIELD("EGG", "USERID", "INT")
320     dum2 = ADD_FIELD("EGG", "ID", "INT IDENTITY(1,1)")
330     dum2 = ADD_FIELD("TIM", "KERDOS", "SINGLE")

340     Me.Picture = LoadPicture(gPicture)

350     For k = 0 To 2: f_text1(k) = "         ": Next

360     PROTH = 1
370     PROTH_FORA = True
380     fflag = 0
        ' DTPicker1.Value = CDate("01/01/" + Str(Year(Now)))

390     DTPicker1.Value = gEnarjh
        'eos.Value = gLhjh
        
        F_NEA_KRTELLA = Val(FINDPARAMETROI(1, "PELAT4", "F_NEA_KRTELLA", "0", "NEA ΚΑΡΤΕΛΑ ΠΕΛΑΤΗ=1 ΠΑΛΙΑ=0"))

        f_KENTRA_ADYNATISMATOS = Val(FINDPARAMETROI(1, "PELAT2", "f_KENTRA_ADYNATISMATOS", "0", "Αν είναι κέντρο αδυνατίσματος=1"))

f_2000MORE_ROWS = Val(FINDPARAMETROI(1, "PELAT4", "f_2000MORE_ROWS", "1", "Ενεργοποιηση πολλων σειρων>2000 σε καρτελλα"))



        If f_KENTRA_ADYNATISMATOS = 1 Then
            kentroAdyn.Visible = True
            kentroAdyn.ListIndex = 0 ' Combo3.List(0)
            ypolposo.Visible = True
            ypolmhx.Visible = True
            symf(0).Visible = True
            symf(1).Visible = True
    
        End If

400     DTPicker2.Value = Now

        'Dim R As New ADODB.Recordset
410     R.Open "select POL,TITLOS,EIDOS from PARASTAT where POL=1 ORDER BY POL,TITLOS;", Gdb, adOpenDynamic, adLockOptimistic
420     List1.Clear

430     If R.RecordCount = 0 Then

            Exit Sub

        End If

        Dim L As Integer, FirstTime As Boolean

440     L = 0
450     FirstTime = False
460     R.MoveFirst
470     k = 0

480     Do While Not R.EOF

490         If Not IsNull(R("TITLOS")) And Left(R("TITLOS"), 1) <> "-" Then
500             If R("POL") = 2 Then
510                 If FirstTime = False Then L = k: FirstTime = True

520                 List1.AddItem "----" + Format(k, "##") + " " + R("TITLOS") + Space(1) + R("eidos")
                Else
530                 List1.AddItem Format(k, "##") + " " + R("TITLOS") + Space(1) + R("eidos")
                End If

540             k = k + 1
                'mEidos_Par(K) = R("eidos")
            End If

550         R.MoveNext
        Loop

560     R.Close





    tdbgrid1.columns(6).Width = 1544.882
    tdbgrid1.columns(5).Width = 2610.142
    tdbgrid1.columns(4).Width = 480.189
    tdbgrid1.columns(3).Width = 3550.142
    tdbgrid1.columns(2).Width = 929.7639
    tdbgrid1.columns(1).Width = 900.2835
    tdbgrid1.columns(0).Width = 1049.953


















        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Paint()

        '<EhHeader>
        On Error GoTo Form_Paint_Err

        '</EhHeader>
100     If PROTH = 1 Then
110         PROTH = 2
120         data1.ConnectionString = gConnect
130         data1.RecordSource = "select top 10 *from PEL WHERE EIDOS='z';"
140         data1.Refresh
150         egg.ConnectionString = gConnect

160         If Len(Text2(0).Text) > 0 Then
170             data1.RecordSource = "select EPO,THL,DIE,* from PEL where EIDOS='" + mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1) + "' and KOD like '" + Text2(0).Text + "%';"
180             data1.Refresh
190             Text2_LostFocus 0
            End If

200         If Len(Trim(ffbuff.Caption)) > 0 Then
210             If Left(ffbuff.Caption, 1) = "e" Then
220                 PELPROM.ListIndex = 0
                Else
230                 PELPROM.ListIndex = 1
                End If

240             Text2(0).Text = Trim(mID(ffbuff, 2, 10))
250             Text2_LostFocus 0
260             SendKeys "~"    'ENTER

            End If
        End If

        '<EhFooter>
        Exit Sub

Form_Paint_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Form_Paint " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Form_Paint " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next





'-----------------------νεο resize υπο εξεταση ---αν βγαλω τα ρεμ παιζει κανονικα 0.95 * 0.75 δικα μου----------------------------------------------------------
'Dim ThisControl As Control, HorizRatio As Single, VertRatio As Single, Iter As Integer
'  If Me.WindowState = vbMinimized Then Exit Sub
'  HorizRatio = Me.Width / ControlInfos(0).Width
'  VertRatio = 0.95 * Me.Height / ControlInfos(0).Height
'
'  Iter = 0
'  For Each ThisControl In Me.Controls
'    Iter = Iter + 1
'    On Error Resume Next  ' hack to bypass controls
'    With ThisControl
'      .Left = ControlInfos(Iter).Left * HorizRatio
'      .Top = ControlInfos(Iter).Top * VertRatio
'      .Width = ControlInfos(Iter).Width * HorizRatio
'      .Height = ControlInfos(Iter).Height * VertRatio
'      .FontSize = ControlInfos(Iter).FontSize * HorizRatio * 0.7
'    End With
'    On Error GoTo 0
'  Next
'  Exit Sub
'------------------------------------------------------------------------------------































   
    '</EhHeader>
    If f_resize = 1 Then
        '  mForm_Resize pelat4, 15000, 10000, fh, fw, ft, fl
        
        f_resizeCount = f_resizeCount + 1

'        If f_resizeCount = 1 Then
'           Me.Hide
'        End If
    
        ' F_PLATOS_FORMAS = 16000 '20000 OK
        ' F_YCOS_FORMAS = 9000 ' 12000 OK
        If F_PLATOS_FORMAS < 0 Then
            ResizeFormFor Me
        Else
            mForm_Resize Me, 8, F_PLATOS_FORMAS, F_YCOS_FORMAS, fh, fw, ft, fl
        End If
'        If f_resizeCount = 2 Then Me.SHOW
    
    End If

    tdbgrid1.HeadFont.Size = CMDEXIT.FontSize
    tdbgrid1.FooterFont.Size = CMDEXIT.FontSize
    
    ' El.FormResize Me
    ' El.mForm_Resize Me

End Sub

Private Sub KentroAdyn_Change()

    Dim a

    a = 1
End Sub

Private Sub kentroAdyn_Click()

    Dim a

    a = 1

    On Error Resume Next

    If f_KENTRA_ADYNATISMATOS = 1 Then
        If data1.Recordset("NUM3") = 0 Then ' HTAN STO SVMA
            If kentroAdyn.Text = "ΣΩΜΑ" Then  ' PARAMENEI STO SOMA
                'MHN KANEIS TIPOTA
                Text2(0).Text = data1.Recordset("kod")
                f_msynt = "NUM3=0 AND "
            Else  ' TO GYRIZEI STO PROSOPO
                Text2(0).Text = "10" + data1.Recordset("KOD")
                f_msynt = "NUM3=1 AND "
            End If
        End If
   
        If data1.Recordset("NUM3") = 1 Then ' HTAN STO PROSOPO
            If kentroAdyn.Text <> "ΣΩΜΑ" Then  ' PARAMENEI STO PROSOPO
                'MHN KANEIS TIPOTA
                Text2(0).Text = data1.Recordset("kod")
                f_msynt = "NUM3=1 AND "
            Else  ' TO GYRIZEI STO SOMA
                Text2(0).Text = mID(data1.Recordset("KOD"), 3, 4)
                f_msynt = "NUM3=1 AND "
          
            End If
        End If
    
        Text2(1).Text = ""
        Text2(0).SetFocus
    
        Text2(1).SetFocus
   
    End If

End Sub

Private Sub List1_Click()


Dim a As String
a = List1.List(List1.ListIndex)
Dim C1, c2 As String
C1 = Split(a, " ")(2)
c2 = Split(a, " ")(0)


mpelat41_show gDir, C1, c2 '   d1 As String, L1 As String)




End Sub

Private Sub MSFlexGrid1_Click()

        '<EhHeader>
        On Error GoTo MSFlexGrid1_Click_Err

        '</EhHeader>

        Dim a, mgdir, DUM

100     timtext3.Visible = False

110     If MSFlexGrid1.rows < 3 Then Exit Sub

        ' MSFlexGrid1.height = 1215
        ' If left(Command5.Caption, 1) = "Π" Then
        
        
       Dim CurRow As Long
        Dim L1 As String
        Dim k As Long
        Dim L As Long
        Dim D1 As String
        
             MSFlexGrid1.Col = 0
        ' βρισκω την σειρά που έκανα κλικ
        'If MSFlexGrid1.row + 1 < MSFlexGrid1.Rows - 1 Then
        '    MSFlexGrid1.row = MSFlexGrid1.row + 1
        'End If
     CurRow = MSFlexGrid1.row    '+ 1

120     L1 = MSFlexGrid1.TextMatrix(CurRow, 3)    ' ATIM



130     Me.Caption = L1

        ' αποχρωματιζω τις όλες τις σειρές
140     For k = 1 To MSFlexGrid1.rows - 1

150         If MSFlexGrid1.RowIsVisible(k) Then
160             MSFlexGrid1.row = k

                ' αν ειναι πράσινο επανέφερε το
170             If MSFlexGrid1.CellBackColor = vbGreen Then

180                 For L = 0 To MSFlexGrid1.ColS - 1
190                     MSFlexGrid1.Col = L
200                     MSFlexGrid1.CellBackColor = IIf(k Mod 2 = 1, MSFlexGrid1.BackColor, &HC0E0FF)
                    Next

                End If
            End If

        Next

        'χρωματίζω πράσινη την τρέχουσα σειρά
210     MSFlexGrid1.row = CurRow

220     For k = 0 To MSFlexGrid1.ColS - 1
230         Me.Caption = "Σ " + str(MSFlexGrid1.row)
240         MSFlexGrid1.Col = k: MSFlexGrid1.CellBackColor = vbGreen

        Next

250     MSFlexGrid1.Col = 0: MSFlexGrid1.CellFontBold = True
260     MSFlexGrid1.Col = 0    'ΗΜΕΡΟΜΗΝΙΑ
270     MSFlexGrid1.Refresh

280     If Not IsDate(MSFlexGrid1.Text) Then Exit Sub

290     D1 = Trim(MSFlexGrid1.Text)

300   '  d2 = DateAdd("d", 1, d1)
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
     DUM = mpelat41_show(gDir, D1, L1)
        ' Else
        '    mgdir = gDir + "\" + Right(Format(Year(Now) - 1, "0000"), 2)
        '   DUM = mpelat41_show(mgdir)
        ' End If

        '<EhFooter>
        Exit Sub

MSFlexGrid1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.MSFlexGrid1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.MSFlexGrid1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub MSFlexGrid1_DblClick()
'-------------------- διαγραφη

170     If MSFlexGrid1.CellBackColor = vbGreen Then
            ' ok σημαδεύει δωστά
        Else

            Exit Sub

        End If
        
         Dim m_hme, m_ait, m_xreosi, m_pistosi, m_apa

       

110     MSFlexGrid1.row = MSFlexGrid1.row    '+ 1
120     MSFlexGrid1.Col = 1: m_hme = MSFlexGrid1.Text
    
130     MSFlexGrid1.Col = 3: m_apa = MSFlexGrid1.Text

140     MSFlexGrid1.Col = 4: m_ait = MSFlexGrid1.Text

150     MSFlexGrid1.Col = 5: m_xreosi = MSFlexGrid1.Text
160     MSFlexGrid1.Col = 6: m_pistosi = MSFlexGrid1.Text
        
        
        Dim m_ID As String
        MSFlexGrid1.Col = 9: m_ID = MSFlexGrid1.Text

   DiagEggraf m_hme, m_ait, m_xreosi, m_pistosi, m_apa, m_ID

End Sub


Sub DiagEggraf(m_hme, m_ait, m_xreosi, m_pistosi, m_apa, m_ID As String)



        '<EhHeader>
        On Error GoTo MSFlexGrid1_DblClick_Err

        '</EhHeader>
        Dim a, mgdir, DUM

        Dim FF

100     If Len(Dir("c:\mercvb\ll.txt")) > 0 Then

            Exit Sub  ' DEN EPITREPONTAI OI DIAGRAFES

            ' End If

        End If
 Dim RECS
        ' a = MSFlexGrid1.Text
       
    
        'If m_xreosi + m_pistosi = 0 Then   ' ηταν προηγ υπολοιπο
        '    Exit Sub
        'End If
    

    
        Dim mID

        'MSFlexGrid1.Col = 8: MID = MSFlexGrid1.Text

        'F_CanDelete = FindParametroi(1,"BOHU11", "F_CanDelete", "1", "μπορω να σβήνω κινήσεις")

180    ' MSFlexGrid1.Col = 9: mID = MSFlexGrid1.Text
        'MSFlexGrid1.Col = 8: MID = MSFlexGrid1.Text

        'δεν συμαδευει πραγματικη εγγραφή
190     If Val(m_ID) = 0 Then

            Exit Sub

        End If

200     If F_CanDelete = 0 Then

210         FF = InputBox("Η ΔIAΓΡΑΦΗ ΔΕΝ ΕΠΙΤΡΕΠΕΤΑΙ")

220         If FF = "2304603921" Then
            Else

                Exit Sub

            End If

        End If

230     DUM = MsgBox("Να διαγραφεί η εγγραφή; " + Chr(13) + m_hme + Chr(13) + m_ait + Chr(13) + m_xreosi + m_pistosi, vbYesNo)

240     If DUM = vbNo Then
250         MsgBox "Ακυρώθηκε η διαγραφή"

            Exit Sub

        End If
Dim ll As Integer



      '  ll = Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)))
260    ' If Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)) = 0 And Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)) > 0 Then
270        ' MsgBox "δεν διεγράφη επειδή σχετίζεται με ένα παραστατικό"

           ' Exit Sub

        'End If

    ' If Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)) = 0 And Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)) = 0 Then
           ' MsgBox "δεν διεγράφη επειδή σχετίζεται με ΕΠΙΤΑΓΗ"

           ' Exit Sub

       ' End If

            
            
            ' ΒΛΕΠΩ ΑΝ ΕΧΕΙ ΣΥΝΔΕΟΜΕΝΗ ΕΓΓΡΑΦΗ
            Dim IDN As Long, RR As New ADODB.Recordset
            RR.Open "SELECT * FROM EGG WHERE ID=" + m_ID, Gdb, adOpenDynamic, adLockOptimistic
            
             If IsNull(RR("IDGRA")) And IsNull(RR("IDTIM")) Then
                  'OK
             
             Else
                If Not IsNull(RR("IDGRA")) Then
                    MsgBox "δεν διεγράφη επειδή σχετίζεται με ΕΠΙΤΑΓΗ"
                    Exit Sub
                End If
                
                If Not IsNull(RR("IDTIM")) Then
                    
                    
                    MsgBox "δεν διεγράφη επειδή σχετίζεται με ΠΑΡΑΣΤΑΤΙΚΟ"
                    Exit Sub
                End If
              End If
              
                 
                 ' OK ΔΕΝ ΣΧΕΤΙΖΕΤΑΙ ΜΕ ΕΠΙΤΑΓΗ ΜΠΟΡΩ ΝΑ ΤΗΝ ΣΒΗΣΩ
                 
            
            
            
            
            'RR.Open "SELECT IDEGGSYND FROM EGG WHERE ID=" + mID, Gdb, adOpenDynamic, adLockOptimistic
            If IsNull(RR("IDEGGSYND")) Then
               IDN = 0
            Else
               IDN = RR(0)
            End If
            RR.Close
            


280     a = "delete  FROM EGG WHERE ID=" + m_ID
        '+ " AND KOD='" + DATA1.Recordset("kod") + "' AND EIDOS='" + DATA1.Recordset("EIDOS") + "'"
        'a = a + " AND CONVERT(CHAR(10),HME,103)='" + Format(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0), "MM/DD/YYYY") + "'" '  AND AIT like '" + m_ait + "%'"
        'If Val(m_xreosi) = 0 Then
        '   a = a + " AND PISTOSI=" + Str(Val(m_pistosi))
        'Else
        '   a = a + " AND XREOSI=" + Str(Val(m_xreosi)) ' m_xreosi
        'End If
290     RECS = 0

        On Error Resume Next
        'Gdb.BeginTrans
        
300     Gdb.Execute a, RECS

310     If RECS > 0 Then
            If IDN > 0 Then ' ΕΧΕΙ ΣΥΝΔΕΟΜΕΝΗ ΕΓΓΡΑΦΗ
                Gdb.Execute "delete  FROM EGG WHERE ID=" + str(IDN)
            End If
            
320         MsgBox "ΔΙΕΓΡΑΦΗΣΑΝ " + str(RECS) + " ΕΓΓΡΑΦΗ/ΕΣ"
        Else
330         MsgBox "ΔΕΝ ΔΙΕΓΡΑΦΗ "
        End If

340     DUM = KARTELLA(gDir)

        '<EhFooter>
        Exit Sub

MSFlexGrid1_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.MSFlexGrid1_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.MSFlexGrid1_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then  'F3
100     MSFlexGrid1.Col = 0
        ' βρισκω την σειρά που έκανα κλικ
        'If MSFlexGrid1.row + 1 < MSFlexGrid1.Rows - 1 Then
        '    MSFlexGrid1.row = MSFlexGrid1.row + 1
        'End If
        
        Dim CurRow As Integer
        
        
110     CurRow = MSFlexGrid1.row    '+ 1


    Dim c, D, d2 As String
    Dim ID As String
    ID = MSFlexGrid1.TextMatrix(CurRow, 9)

120      c = MSFlexGrid1.TextMatrix(CurRow, 4)    ' ATIM


  
   
    D = MSFlexGrid1.TextMatrix(CurRow, 2)    ' APA
    
    d2 = MSFlexGrid1.TextMatrix(CurRow, 3)    ' ATIM
Dim POSO As Single
If gVal(MSFlexGrid1.TextMatrix(CurRow, 5)) > 0 Then
    POSO = gVal(MSFlexGrid1.TextMatrix(CurRow, 5))
Else
    POSO = -gVal(MSFlexGrid1.TextMatrix(CurRow, 6))
End If


     DiorEggraf ID, c, D, d2, POSO
     
      Dim DUM
    DUM = KARTELLA(gDir)
     
     
     
End If
     
End Sub

Sub DiorEggraf(ID, c, D, d2, POSO As Single)


 c = InputBox("διόρθωση Αιτιολογίας", , c)

    
    
    Dim IDIO As Integer
    IDIO = 0
    If d2 = D Then
        IDIO = 1 ' Gdb.Execute "update EGG SET ATIM=" + D + " WHERE ID=" + ID
    
    End If
    
    
    
    Dim CPOSO As String
     Gdb.Execute "update EGG SET AIT='" + Replace(c, "'", "`") + "' WHERE ID=" + ID
     
   Dim NCC As Long: NCC = GGET_NVALUE("SELECT ISNULL(IDTIM,0) FROM EGG WHERE ID=" + ID)
   If NCC > 0 Then
       Exit Sub
   End If
    
    
    
    
    If POSO > 0 Then
         CPOSO = InputBox("Διόρθωση ποσού", , str(POSO))
        Gdb.Execute "update EGG SET PISTOSI=0,XREOSI=" + Replace(CPOSO, ",", ".") + ",XRE=" + Replace(CPOSO, ",", ".") + " WHERE ID=" + ID
    Else
          CPOSO = InputBox("Διόρθωση ποσού", , str(-POSO))
        Gdb.Execute "update EGG SET XREOSI=0,PISTOSI=" + Replace(CPOSO, ",", ".") + ",XRE=-" + Replace(CPOSO, ",", ".") + " WHERE ID=" + ID
    End If
    
    
    D = InputBox("Διόρθωση Αριθμού", , D)
    

       
    
        Gdb.Execute "update EGG SET ATIM='" + D + "' WHERE ID=" + ID
   
   
   If IDIO = 1 Then
        Gdb.Execute "update EGG SET ATIM='" + D + "' WHERE ID=" + ID
    
    End If
   
   
    
   
    





End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                Y As Single)
                                
                                
           rePrint Button
                                
End Sub


Sub rePrint(Button As Integer) ' 2=dexi klik

        '<EhHeader>
        On Error GoTo MSFlexGrid1_MouseUp_Err

        '</EhHeader>
        Dim RR As Recordset

        Dim DB As DAO.Database

        Dim sql

        Dim R As New ADODB.Recordset

        On Error Resume Next

100     If Button = 2 Then


Dim D1 As String


110         Set DB = OpenDatabase("C:\MERCVB\REPORTS\REPORTS.MDB")
120         Set RR = DB.OpenRecordset("mem")


             Dim MDATE, md1 As String
            

















130        '' RR.Edit
           If F_NEA_KRTELLA = 0 Then ' palia
140   '        RR!n1 = Val(Replace(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6), ",", "."))  ' Val(Text3.Text) ' poso
150    '       RR!par10 = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)    ' Text1.Text ' ar.parastatikoy
160     '      RR!hme = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0)  'DTPicker1.value
         '     RR!par12 = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4)    'AITIOLOGIA
         
         
               MDATE = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0)
                md1 = MDATE
               MDATE = Split(MDATE, "/")(1) + "/" + Split(MDATE, "/")(0) + "/" + Split(MDATE, "/")(2)
                 
         
         
         
         
         
               DB.Execute "update MEM SET N1=" + str(gVal((Replace(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6), ",", "."))))
               
               DB.Execute "update MEM SET MEMO='" + Olografos(gVal((Replace(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6), ",", ".")))) + "'"
               
               
               DB.Execute "update MEM SET PAR10=" + MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
               DB.Execute "update MEM SET HME=#" + md1 + "#"
               DB.Execute "update MEM SET par12='" + MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) + "'"
                    DB.Execute "update MEM SET PAR9='" + md1 + "'"
         
         
         
           Else
              'TDBGrid1.Col = 6: d1 = TDBGrid1.Text
162  '         RR!n1 = Val(Replace(d1, ",", "."))  ' Val(Text3.Text) ' poso
              
              Dim DATE1 As String, L1 As String
      
                If tdbgrid1.row < 0 Then Exit Sub
                    tdbgrid1.Col = 1
         
                  L1 = tdbgrid1.Text
                      tdbgrid1.Col = 0
                      DATE1 = tdbgrid1.Text
              
              
              
              tdbgrid1.Col = 4: D1 = tdbgrid1.Text
              
              
              
              
              
              DB.Execute "update MEM SET N1=" + str(gVal(Replace(D1, ",", ".")))

               DB.Execute "update MEM SET MEMO='" + Olografos(gVal(Replace(D1, ",", "."))) + "'"




              tdbgrid1.Col = 2: D1 = tdbgrid1.Text
164          ' RR!par10 = d1   ' Text1.Text ' ar.parastatikoy
              tdbgrid1.Col = 1: D1 = tdbgrid1.Text
              DB.Execute "update MEM SET PAR10=" + D1

166           tdbgrid1.Col = 0: D1 = tdbgrid1.Text
              MDATE = D1
               MDATE = Split(MDATE, "/")(1) + "/" + Split(MDATE, "/")(0) + "/" + Split(MDATE, "/")(2)
              'RR!hme = d1  'DTPicker1.value
 DB.Execute "update MEM SET HME=#" + MDATE + "#"
            
              tdbgrid1.Col = 3: D1 = tdbgrid1.Text
168         '  RR!par12 = d1    'AITIOLOGIA

              DB.Execute "update MEM SET PAR9='" + DATE1 + "'"
               
              tdbgrid1.Col = 2: D1 = tdbgrid1.Text: DB.Execute "update MEM SET par11='" + D1 + "'"
              
             'TDBGrid1.Col = 3: d1 = TDBGrid1.Text: DB.Execute "update MEM SET par11='" + d1 + "'"

                

           End If



170     '    RR!PAR2 = data1.Recordset("kod")
'180         RR!C1 = data1.Recordset("epo")
'190         RR!c2 = data1.Recordset("epa")
'200         RR!c3 = data1.Recordset("die")
'210         RR!c4 = data1.Recordset("afm")
220       '  RR!PAR1 = data1.Recordset("DOY")







        
         
DB.Execute "update MEM SET par2='" + data1.Recordset("kod") + "'"

DB.Execute "update MEM SET C1='" + data1.Recordset("epO") + "'"
DB.Execute "update MEM SET C2='" + data1.Recordset("epA") + "'"
DB.Execute "update MEM SET C3='" + Left(data1.Recordset("die"), 25) + "'"
DB.Execute "update MEM SET C4='" + data1.Recordset("afm") + "'"














230

240         sql = "select sum(XREOSI) AS XX0,sum(PISTOSI) AS PP0 from EGG WHERE  HME>='" + "01/01/" + LTrim(str(Year(Now))) + "' AND HME<='" + Format(DTPicker1.Value, "MM/DD/YYYY") + "' AND EIDOS ='" + data1.Recordset("eidos") + "' and KOD='" + data1.Recordset("KOD") + "'"

250         R.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly
260       '  RR!par3 = Format(R(0) + data1.Recordset("ayp") - R(1), "###,000.00")
                DB.Execute "update MEM SET par3='" + Format(R(0) + data1.Recordset("ayp") - R(1), "###,000.00") + "'"


270         R.Close

            '  RR!PAR11 = Left(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3), RR("PAR11").Size)
            ' DB.Execute "update MEM SET par11='" + Left(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3), RR("PAR11").Size) + "'"
            
280         'RR!memo = Olografos(Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6)))
           
290        ' RR.Update

300         If data1.Recordset("eidos") = "e" Then
                'If MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = "ΜΕ" Then
310             CrystalReport1.ReportFileName = "C:\MERCVB\REPORTS\EISPR.RPT"
                'Else
                '   CrystalReport1.ReportFileName = "C:\MERCVB\REPORTS\sympsif.RPT"
                'End If

            Else

                '   If MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = "ΜΕ" Then
320             CrystalReport1.ReportFileName = "C:\MERCVB\REPORTS\plir.RPT"
                '  Else
                '   CrystalReport1.ReportFileName = "C:\MERCVB\REPORTS\sympsif2.RPT"
                '  End If
            End If
            Ucr9print.printingCR9 F_EKTYPOTHS, CrystalReport1.ReportFileName
            
330       ' CrystalReport1.Destination = crptToWindow
340       ' CrystalReport1.Action = 1

        End If

        '<EhFooter>
        Exit Sub

MSFlexGrid1_MouseUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.MSFlexGrid1_MouseUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.MSFlexGrid1_MouseUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub Option1_Click(Index As Integer)
'f_text1(1) = " "
'Text2_LostFocus 1
'
'Dim R As New ADODB.Recordset
'
''If Option1(1) Then x = "e" Else x = "r"
'Dim X As String
'
'If Option1(1) Then
'   X = "1"
'   f_buff = "e"
'Else
'   X = "2"
'   f_buff = "r"
'End If
'
'
'R.Open "select POL,TITLOS,EIDOS from PARASTAT where POL=" + X + " ORDER BY POL,TITLOS;", Gdb, adOpenDynamic, adLockOptimistic
' List1.Clear
' If R.RecordCount = 0 Then
'    Exit Sub
' End If
'
'
' Dim k As Integer
'
'
' Dim L As Integer, FirstTime As Boolean
' L = 0
' FirstTime = False
' R.MoveFirst
'  k = 0
'  Do While Not R.EOF
'     If Not IsNull(R("TITLOS")) And Left(R("TITLOS"), 1) <> "-" Then
'        If R("POL") = 2 Then
'           If FirstTime = False Then L = k: FirstTime = True
'
'           List1.AddItem "----" + Format(k, "##") + " " + R("TITLOS") + Space(1) + R("eidos")
'        Else
'           List1.AddItem Format(k, "##") + " " + R("TITLOS") + Space(1) + R("eidos")
'        End If
'        k = k + 1
'        'mEidos_Par(K) = R("eidos")
'     End If
'     R.MoveNext
'  Loop
'
'
'  If Len(Trim(Text2(1))) > 0 Then 'eponumo
'      data1.RecordSource = "select EPO,THL,DIE,* from PEL where EIDOS='" + f_buff + "' and EPO like '" + (SameLetters(Text2(1).Text)) + "%';"
'   End If
'   If Len(Trim(Text2(0))) > 0 Then 'kodikos
'      data1.RecordSource = "select EPO,THL,DIE,* from PEL where EIDOS='" + f_buff + "' and KOD like '" + Text2(0).Text + "%';"
'   End If
'   If Len(Trim(Text2(2))) > 0 Then 'afm
'       data1.RecordSource = "select EPO,THL,DIE,* from PEL where EIDOS='" + f_buff + "' and AFM like '" + Text2(2).Text + "%';"
'   End If
'
'data1.Refresh
'
'
'End Sub

Private Sub PELPROM_Change()

        '<EhHeader>
        On Error GoTo PELPROM_Change_Err

        '</EhHeader>

100     f_buff = mID(PELPROM_STRING, IIf(PELPROM.ListIndex = -1, 1, PELPROM.ListIndex + 1), 1)

110     f_text1(1) = " "
120     Text2_LostFocus 1

        Dim R As New ADODB.Recordset

        'If Option1(1) Then x = "e" Else x = "r"
        Dim x As String

130     If f_buff = "r" Then
140         x = "2"
        Else
150         x = "1"
        End If

160     R.Open "select POL,TITLOS,EIDOS from PARASTAT where POL=" + x + " ORDER BY POL,TITLOS;", Gdb, adOpenDynamic, adLockOptimistic
170     List1.Clear

180     If R.RecordCount = 0 Then

            Exit Sub

        End If

        Dim k As Integer

        Dim L As Integer, FirstTime As Boolean

190     L = 0
200     FirstTime = False
210     R.MoveFirst
220     k = 0

230     Do While Not R.EOF

240         If Not IsNull(R("TITLOS")) And Left(R("TITLOS"), 1) <> "-" Then
250             If R("POL") = 2 Then
260                 If FirstTime = False Then L = k: FirstTime = True

270                 List1.AddItem "----" + Format(k, "##") + " " + R("TITLOS") + Space(1) + R("eidos")
                Else
280                 List1.AddItem Format(k, "##") + " " + R("TITLOS") + Space(1) + R("eidos")
                End If

290             k = k + 1
                'mEidos_Par(K) = R("eidos")
            End If

300         R.MoveNext
        Loop

310     If Len(Trim(Text2(1))) > 0 Then    'eponumo
320         data1.RecordSource = "select EPO,THL,DIE,* from PEL where EIDOS='" + f_buff + "' and EPO like '" + (SameLetters(Text2(1).Text)) + "%';"
        End If

330     If Len(Trim(Text2(0))) > 0 Then    'kodikos
340         data1.RecordSource = "select EPO,THL,DIE,* from PEL where EIDOS='" + f_buff + "' and KOD like '" + Text2(0).Text + "%';"
        End If

350     If Len(Trim(Text2(2))) > 0 Then    'afm
360         data1.RecordSource = "select EPO,THL,DIE,* from PEL where EIDOS='" + f_buff + "' and AFM like '" + Text2(2).Text + "%';"
        End If

370     data1.Refresh

        '<EhFooter>
        Exit Sub

PELPROM_Change_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.PELPROM_Change " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.PELPROM_Change " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub PELPROM_LostFocus()

        '<EhHeader>
        On Error GoTo PELPROM_LostFocus_Err

        '</EhHeader>

100     PELPROM_Change

        '<EhFooter>
        Exit Sub

PELPROM_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.PELPROM_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.PELPROM_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub




Sub SHOWTIMOL()



        Dim a, mgdir, DUM

100     timtext3.Visible = False

 



110   ' If TDBGrid1.row < 1 Then Exit Sub
     
         tdbgrid1.Col = 1
         Dim D1 As String, L1 As String
         
     
     
        ' MSFlexGrid1.height = 1215
        ' If left(Command5.Caption, 1) = "Π" Then
120     L1 = tdbgrid1.Text
         tdbgrid1.Col = 0
         D1 = tdbgrid1.Text

If Len(L1) > 0 And Len(D1) > 0 Then DUM = mpelat41_show(gDir, D1, L1)
        ' Else
        '    mgdir = gDir + "\" + Right(Format(Year(Now) - 1, "0000"), 2)
        '   DUM = mpelat41_show(mgdir)
        ' End If




End Sub




Private Sub TDBGrid1_DblClick()
'διαγραφη εγγραφης-----------------------------------------------
      Dim D1 As String, L1 As String
      
110      If tdbgrid1.row < 0 Then Exit Sub
         tdbgrid1.Col = 1
         
120     L1 = tdbgrid1.Text
         tdbgrid1.Col = 0
         D1 = tdbgrid1.Text
      
        Dim m_hme, m_ait, m_xreosi, m_pistosi, m_apa

       


     tdbgrid1.Col = 0: m_hme = tdbgrid1.Text
    
130     tdbgrid1.Col = 1: m_apa = tdbgrid1.Text

140     tdbgrid1.Col = 2: m_ait = tdbgrid1.Text

150     tdbgrid1.Col = 3: m_xreosi = tdbgrid1.Text

160     tdbgrid1.Col = 4: m_pistosi = tdbgrid1.Text
      
      If Val(m_xreosi) = 0 Then m_xreosi = ""
      If Val(m_pistosi) = 0 Then m_pistosi = ""
      
      Dim m_ID As String
      
        tdbgrid1.Col = 7: m_ID = tdbgrid1.Text
      
     DiagEggraf m_hme, m_ait, m_xreosi, m_pistosi, m_apa, m_ID
      
      neakartela 1
      
      
      
      
      
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)


'RSClone.Bookmark = Bookmark

 '  [ΧΡΕΟΣΗ]

'If ButtonFlag = 0 And RSClone("ΧΡΕΟΣΗ").Value > 0 Then

   ' RowStyle = Prospective

'End If

 

'If ButtonFlag = 0 And RSClone("ΧΡΕΟΣΗ").Value > 0 Then

   ' RowStyle = Distributors

'End If























End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

'   Dim d1 As String, L1 As String
'
'110      If TDBGrid1.row < 1 Then Exit Sub
'         TDBGrid1.Col = 1
'
'120     L1 = TDBGrid1.Text
'         TDBGrid1.Col = 0
'         d1 = TDBGrid1.Text


If KeyCode = 114 Then  'F3
100
        ' βρισκω την σειρά που έκανα κλικ
        


    Dim cait, d2atim, dapa As String
    Dim ID As String
    tdbgrid1.Col = 7
    ID = tdbgrid1.Text  'MSFlexGrid1.TextMatrix(CurRow, 9)

120    '  c = MSFlexGrid1.TextMatrix(CurRow, 4)    ' ATIM

    tdbgrid1.Col = 2
    cait = tdbgrid1.Text


  ' c = InputBox("διόρθωση Αιτιολογίας", , c)
   
  '  d = MSFlexGrid1.TextMatrix(CurRow, 2)    ' APA
    
    
    
    tdbgrid1.Col = 1
    dapa = tdbgrid1.Text
    
    
    tdbgrid1.Col = 1
    d2atim = tdbgrid1.Text
    
    
   ' d2 = MSFlexGrid1.TextMatrix(CurRow, 3)    ' ATIM


Dim POSO As Single

tdbgrid1.Col = 3
If tdbgrid1.Text > 0 Then
    POSO = gVal(tdbgrid1.Text)
Else
    tdbgrid1.Col = 4
    POSO = -gVal(tdbgrid1.Text)
End If




     DiorEggraf ID, cait, dapa, d2atim, POSO
     
     neakartela 1
End If



End Sub



Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
     rePrint Button
   

End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     SHOWTIMOL
End Sub

Private Sub Text2_Change(index As Integer)
    '   Text2_LostFocus Index
    '   SendKeys "{enter}"
    '   SendKeys "{enter}"
    '   TEXT2(Index).SetFocus
    '
End Sub

Private Sub Text2_GotFocus(index As Integer)

        '<EhHeader>
        On Error GoTo Text2_GotFocus_Err

        '</EhHeader>

100     Text2(index).BackColor = vbYellow

        '<EhFooter>
        Exit Sub

Text2_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Text2_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Text2_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text2_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Text2_KeyUp_Err

        '</EhHeader>

100     If KeyCode = 13 Then
110         If index <= 1 Then Text2(index + 1).SetFocus
120         If index = 2 Then DbGrid2.SetFocus    ' Text2(1).SetFocus
        End If

130     If KeyCode = 38 Then
140         If index > 0 Then Text2(index - 1).SetFocus

        End If

        '<EhFooter>
        Exit Sub

Text2_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Text2_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Text2_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text2_LostFocus(index As Integer)

        '<EhHeader>
        On Error GoTo Text2_LostFocus_Err

        '</EhHeader>
        Dim x As String

100     Text2(index).BackColor = vbWhite

        Dim mBUFF As String

110     mBUFF = mID(PELPROM_STRING, IIf(PELPROM.ListIndex = -1, 1, PELPROM.ListIndex + 1), 1)

120     If f_text1(index) = Text2(index).Text Then
130         If mBUFF = f_buff Then Exit Sub
        End If

140     f_text1(index) = Text2(index).Text    ' το κρατω σε μεταβλητή για να βλέπω αν άλλαξε
150     LAST_TIM = -1

160     If Len(Trim(Text2(index))) = 0 Then Exit Sub
170     PROTH_FORA = False
180     Me.MousePointer = vbHourglass
190     x = mID(PELPROM_STRING, PELPROM.ListIndex + 1, 1)
200     gBuff = x

        If f_KENTRA_ADYNATISMATOS = 1 Then
              If gXEIRISTHS = 1 Then
                 msEditor1.Visible = False
              End If
            If x = "e" Then
                If kentroAdyn.List(0) = kentroAdyn.Text Then
                    f_msynt = "NUM3=0 AND "
                Else
                    f_msynt = "NUM3=1 AND "
                End If
            End If
        End If

280     If index = 1 Then  'eponumo
            If F_MONOENERGOS = 0 Then
               data1.RecordSource = "select EPO,THL,DIE,* from PEL where " + f_msynt + " EIDOS='" + x + "'  and EPO like '" + (SameLetters(Text2(1).Text)) + "%';"
            Else
            
290            data1.RecordSource = "select EPO,THL,DIE,* from PEL where " + f_msynt + " EIDOS='" + x + "' AND ISNULL(ENERGOS,1)=1 and EPO like '" + (SameLetters(Text2(1).Text)) + "%';"
            End If
        End If

300     If index = 0 Then    'kodikos
310         data1.RecordSource = "select EPO,THL,DIE,* from PEL where " + f_msynt + " EIDOS='" + x + "' and KOD like '" + Text2(0).Text + "%';"
        End If

320     If index = 2 Then    'afm
330         data1.RecordSource = "select EPO,THL,DIE,* from PEL where " + f_msynt + " EIDOS='" + x + "' and (AFM like '%" + Text2(2).Text + "%' OR " + f_pedioAFM + " like '%" + Text2(2).Text + "%' );"
        End If

340     data1.Refresh
  
  
  
  
  
        Set msEditor1.mDataSource = data1  ' for code you have to use the 'adoNotes' instead of the Adodc1
        msEditor1.MaxLength = 8000          ' Optional maximum length.
        '
        msEditor1.mDataField = "MEMO" '  "nNotes"     ' The field of your text
        msEditor1.Enabled = True

        ' HMESYMF.Visible = True

        If f_KENTRA_ADYNATISMATOS = 1 Then
   
            ' data1.Refresh
   
            If IsNull(data1.Recordset("num5")) Then
                ypolmhx.Caption = data1.Recordset("num2")
            Else
                ypolmhx.Caption = data1.Recordset("num2") - data1.Recordset("num5")
            End If
   
            If IsNull(data1.Recordset("num4")) Then
                ypolposo.Caption = data1.Recordset("num1")
            Else
                ypolposo.Caption = data1.Recordset("num1") - data1.Recordset("num4")
            End If

        End If

350     DbGrid2.columns(0).Width = 4000

360     cmdPRINTKARTELA.Enabled = True
370     cmdEIDHSYNOLIKA.Enabled = True
380     cmdXREOPISTOSEIS.Enabled = True

390     x = ""

400     If MSFlexGrid1.rows > 1 Then MSFlexGrid1.TopRow = IIf(MSFlexGrid1.rows - 10 <= 0, 1, MSFlexGrid1.rows - 10)
410     Me.MousePointer = vbNormal
420     DbGrid2.SetFocus

        DBGrid2_RowColChange 0, 0

        Exit Sub

Text2_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.Text2_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.Text2_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function mpelat41_show(mgdir, ByVal D1 As String, ByVal L1 As String)

Dim d2 As Date

 If IsDate(D1) Then d2 = DateAdd("d", 1, D1)
 
        '<EhHeader>
        On Error GoTo mpelat41_show_Err

        '</EhHeader>
      

        Dim MLAST, CurRow As Long

        Dim L As Integer

        Dim k As Integer

        On Error GoTo mpelat41_show_Err    'Resume Next

       ' On Error Resume Next

'100     MSFlexGrid1.Col = 0
'        ' βρισκω την σειρά που έκανα κλικ
'        'If MSFlexGrid1.row + 1 < MSFlexGrid1.Rows - 1 Then
'        '    MSFlexGrid1.row = MSFlexGrid1.row + 1
'        'End If
'110     CurRow = MSFlexGrid1.row    '+ 1
'
'120     L1 = MSFlexGrid1.TextMatrix(CurRow, 3)    ' ATIM
'
'
'
'130     Me.Caption = L1
'
'        ' αποχρωματιζω τις όλες τις σειρές
'140     For k = 1 To MSFlexGrid1.ROWS - 1
'
'150         If MSFlexGrid1.RowIsVisible(k) Then
'160             MSFlexGrid1.row = k
'
'                ' αν ειναι πράσινο επανέφερε το
'170             If MSFlexGrid1.CellBackColor = vbGreen Then
'
'180                 For L = 0 To MSFlexGrid1.ColS - 1
'190                     MSFlexGrid1.Col = L
'200                     MSFlexGrid1.CellBackColor = IIf(k Mod 2 = 1, MSFlexGrid1.BackColor, &HC0E0FF)
'                    Next
'
'                End If
'            End If
'
'        Next
'
'        'χρωματίζω πράσινη την τρέχουσα σειρά
'210     MSFlexGrid1.row = CurRow
'
'220     For k = 0 To MSFlexGrid1.ColS - 1
'230         Me.Caption = "Σ " + str(MSFlexGrid1.row)
'240         MSFlexGrid1.Col = k: MSFlexGrid1.CellBackColor = vbGreen
'
'        Next
'
'250     MSFlexGrid1.Col = 0: MSFlexGrid1.CellFontBold = True
'260     MSFlexGrid1.Col = 0    'ΗΜΕΡΟΜΗΝΙΑ
'270     MSFlexGrid1.Refresh
'
'280     If Not IsDate(MSFlexGrid1.Text) Then Exit Function
'
'290     d1 = Trim(MSFlexGrid1.Text)
'
'300     d2 = DateAdd("d", 1, d1)
   Dim SQL2 As String
   SQL2 = "select ATIM,HME as SHME,KODE,ONOMA,CONVERT(decimal(10,2),POSO) as [ΠΟΣΟΤ]," & " CONVERT(decimal(10,4),TIMM)AS TIMH,CONVERT(decimal(10,2),TIMM*POSO) as [AΞIA ΠΡΟ ΕΚΠΤ],CONVERT(decimal(10,2),EKPT)AS [ΕΚΠΤ],CONVERT(decimal(10,2),TIMM*POSO*(100-EKPT)/100) as [AΞIA ΜΕΤΑ ΕΚΠΤ],APOT AS [Αποθ],CONVERT(decimal(10,2),POSO*KERDOS) AS [KΡΔ],EID.ERG from EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD where PELKOD='" + data1.Recordset!kod + "' AND  ATIM='" + L1 + "' and DAY(HME)=" + str(Day(D1)) + " and month(HME)=" + str(Month(D1)) + " order by PROOD_AJ;"
  ' write_query SQL2
310     data2.RecordSource = SQL2
   Dim mGDB As New ADODB.Connection
   Dim mCon As String
320     If Check1.Value = vbChecked Then    '   <> gDir Then ' palia xronia
330        mCon = gConnect + ";DATABASE=" + f_PALIAXRONIA     '"DSN=" + LTrim(Trim(Str(f_PALIAXRONIA)))
340     ElseIf Check2.Value = vbChecked And Len(XRONIES.Text) > 0 Then  '   <> gDir Then ' palia xronia
350          mCon = gConnect + ";DATABASE=" + f_PROPERSI  ' "DSN=" + LTrim(Str(f_PALIAXRONIA))
        Else
360          mCon = gConnect
        End If

data2.ConnectionString = mCon
mGDB.Open mCon




        'Data2.Refresh
370     data2.Refresh

        Dim R As New ADODB.Recordset

380     R.Open "select SUM(ISNULL(TIMM,0)*ISNULL(POSO,0) ) as [ΑΞΙΑ ΠΡΟ ΕΚΠ],SUM(ISNULL(POSO,0)*ISNULL(KERDOS,0) ) AS [KΡΔ],SUM(ISNULL(TIMM,0)*ISNULL(POSO,0)*(100-ISNULL(EKPT,0))/100) as [AΞIA META EKΠ],ATIM,HME" & " from EGGTIM where ATIM='" + L1 + "' and HME>='" + Format(D1, "MM/DD/yyyy") + "' and HME<'" + Format(d2, "MM/DD/yyyy") + "' GROUP BY ATIM,HME  ;", mGDB, adOpenForwardOnly, adLockReadOnly

390     DBGRID22.TextMatrix(0, 0) = "ΠΑΡ/ΚΟ"
400     DBGRID22.ColWidth(0) = 0
        'DBGRID22.ColAlignment(0) = 7

410     DBGRID22.Height = 2500
420     DBGRID22.TextMatrix(0, 1) = "ΗΜΕΡ/ΝΙΑ"
430     DBGRID22.ColWidth(1) = 0

440     DBGRID22.TextMatrix(0, 2) = "ΚΩΔ.ΕΙΔΟΥΣ"
450     DBGRID22.ColWidth(2) = 1400

460     DBGRID22.TextMatrix(0, 3) = "ΠΕΡΙΓΡΑΦΗ ΕΙΔΟΥΣ"
470     DBGRID22.ColWidth(3) = TextWidth("A") * 35

480     DBGRID22.TextMatrix(0, 4) = "ΠΟΣΟΤΗΤΑ"
490     DBGRID22.ColWidth(4) = 900
500     DBGRID22.ColAlignment(4) = 7

510     DBGRID22.ColAlignment(5) = 7
         DBGRID22.ColWidth(5) = 1200


520     DBGRID22.ColAlignment(6) = 7
530     DBGRID22.ColAlignment(7) = 7
540     DBGRID22.ColAlignment(8) = 7
550     DBGRID22.ColAlignment(9) = 7

      '  DBGRID22.ColWidth(5) = 700
        DBGRID22.ColWidth(7) = 600
        DBGRID22.ColWidth(9) = 600
        DBGRID22.ColWidth(10) = 800

560     DBGRID22.AddItem " "

570     DBGRID22.TextMatrix(DBGRID22.rows - 1, 6) = Format(R(0), "#####.00")
        DBGRID22.row = DBGRID22.rows - 1: DBGRID22.Col = 8: DBGRID22.CellBackColor = vbGreen
580     DBGRID22.TextMatrix(DBGRID22.rows - 1, 8) = Format(R(2), "#####.00")

585     DBGRID22.TextMatrix(DBGRID22.rows - 1, 10) = Format(R(1), "#####.00")

        'βαζω σκοπο, παράδοση
        'Dim R As New ADODB.Recordset
590     R.Close

600     R.Open "select ATIM,PARAT,PROOR,SKOPOS,B_N1,B_N2,FPA1,FPA2,FPA3,FPA4,FPA6,FPA7,AJI,AJ1,AJ2,AJ3,AJ4,AJ5,AJ6,AJ7 from TIM where" & " ATIM='" + L1 + "' and HME>='" + Format(D1, "MM/DD/yyyy") + "' and HME<'" + Format(d2, "MM/DD/yyyy") + "';", mGDB, adOpenForwardOnly, adLockReadOnly

610     If Not R.EOF Then


            TDBGrid2.columns(0).FooterText = "Σκοπός:"
            TDBGrid2.columns(1).FooterText = R("SKOPOS")

            TDBGrid2.columns(3).FooterText = "Προορισμός:"
            TDBGrid2.columns(4).FooterText = R("proor")

            TDBGrid2.columns(6).FooterText = "Σχόλια:"
            TDBGrid2.columns(7).FooterText = R("PARAT")


620         DBGRID22.AddItem " "
630         DBGRID22.TextMatrix(DBGRID22.rows - 1, 0) = "Σκοπός:"
640         DBGRID22.TextMatrix(DBGRID22.rows - 1, 1) = R("SKOPOS")

            '  DBGRID22.TextMatrix(2, 2) = "Σκοπός:"

            DBGRID22.TextMatrix(DBGRID22.rows - 1, 8) = Format(R(2), "#####.00")

            DBGRID22.TextMatrix(DBGRID22.rows - 1, 6) = "ΑΞΙΑ ΦΠΑ"
             DBGRID22.row = DBGRID22.rows - 1: DBGRID22.Col = 8: DBGRID22.CellBackColor = vbYellow
            DBGRID22.TextMatrix(DBGRID22.rows - 1, 8) = Format(R!FPA1 + R!fpa2 + R!fpa3 + R!FPA4 + R!FPA6 + R!FPA7, "#####.00")

650         DBGRID22.AddItem " "
660         DBGRID22.TextMatrix(DBGRID22.rows - 1, 2) = "Προορισμός:"
670         DBGRID22.TextMatrix(DBGRID22.rows - 1, 3) = R("proor")
            
 DBGRID22.row = DBGRID22.rows - 1: DBGRID22.Col = 8: DBGRID22.CellBackColor = vbGreen
            DBGRID22.TextMatrix(DBGRID22.rows - 1, 6) = "ΣΥΝ.ΑΞΙΑ"
            DBGRID22.TextMatrix(DBGRID22.rows - 1, 8) = Format(R!aji, "#####.00")

680         If IsNull(R("parat")) Then
            Else
690             DBGRID22.AddItem " "
700             DBGRID22.TextMatrix(DBGRID22.rows - 1, 2) = "Σχόλια:"
710             DBGRID22.TextMatrix(DBGRID22.rows - 1, 3) = R("parat")
            End If
        
            Dim r2 As New ADODB.Recordset
        
720         If IsNull(R("B_N1")) Then
            Else
730             r2.Open "SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=4  AND AYJON=" + str(R("B_N1")) + " ORDER BY AYJON", mGDB, adOpenDynamic, adLockOptimistic
740             DBGRID22.AddItem " "
750             DBGRID22.TextMatrix(DBGRID22.rows - 1, 0) = "Αποθ:"

760             If Not IsNull(R("PERIGRAFH")) Then
770                 DBGRID22.TextMatrix(DBGRID22.rows - 1, 1) = r2("PERIGRAFH")
                End If
            
            End If

        End If
        TDBGrid2.columns(8).FooterText = Format(R!aj1 + R!aj2 + R!aj3 + R!aj4 + R!aj5 + R!aj6 + R!AJ7, "#####.00") ' "---"
        TDBGrid2.columns(3).Width = 4500
        TDBGrid2.columns(0).Width = 0
780     R.Close

        'DBGRID22.Columns(1).DataField = "atim"

        'DBGRID22.Columns(1).width = 800
        'DBGRID22.Columns(1).DataField = "shme"
        'DBGRID22.Columns(1).Caption = "ΗΜΕΡ/ΝΙΑ"
        '
        'DBGRID22.Columns(2).Caption = "ΚΩΔ.ΕΙΔΟΥΣ"
        'DBGRID22.Columns(2).DataField = "kode"
        '
        'DBGRID22.Columns(3).width = 2600
        'DBGRID22.Columns(3).Caption = "ΠΕΡΙΓ.ΕΙΔΟΥΣ"
        'DBGRID22.Columns(3).DataField = "onoma"
        '
        'DBGRID22.Columns(4).width = 900
        'DBGRID22.Columns(4).Caption = "ΠΟΣΟΤΗΤΑ"
        'DBGRID22.Columns(4).DataField = "poso"
        'DBGRID22.Columns(4).NumberFormat = "#####.00"
        'DBGRID22.Columns(4).Alignment = dbgRight
        '

        '
        'DBGRID22.Columns(5).width = 800
        'DBGRID22.Columns(5).Caption = "ΤΙΜ.ΜΟΝ"
        'DBGRID22.Columns(5).DataField = "timm"
        'DBGRID22.Columns(5).NumberFormat = "#####.00"
        'DBGRID22.Columns(5).Alignment = dbgRight
        '
        '
        'DBGRID22.Columns(6).width = 800
        'DBGRID22.Columns(6).Caption = "ΑΞΙΑ"
        'DBGRID22.Columns(6).DataField = "ajia"
        'DBGRID22.Columns(6).NumberFormat = "#####.00"
        'DBGRID22.Columns(6).Alignment = dbgRight
        '
        '
        ''On Error Resume Next
        '
        'DBGRID22.Columns(7).width = 800
        'DBGRID22.Columns(7).Caption = "% ΕΚΠΤ"
        'DBGRID22.Columns(7).DataField = "EKPT"
        'DBGRID22.Columns(7).NumberFormat = "###"
        'DBGRID22.Columns(7).Alignment = dbgCenter
        '
        '
        'DBGRID22.Columns(8).Alignment = dbgCenter

        '<EhFooter>
        Exit Function

mpelat41_show_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.mpelat41_show " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.mpelat41_show " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function



Private Sub THLEFONA_Click()

        '<EhHeader>
        On Error GoTo THLEFONA_Click_Err

        '</EhHeader>

Dim ANS As Integer
ANS = MsgBox("Αναλυτικά;", vbYesNo)
100     data2.ConnectionString = gConnect
If ANS = vbYes Then
110     data2.RecordSource = " SELECT * FROM THLEFONA WHERE KOD='" + data1.Recordset("KOD") + "' AND EIDOS='e' order by HME DESC"
Else
        data2.RecordSource = " SELECT count(*) as AR,MONTH(HME) AS [MHNAΣ],YEAR(HME) AS [ETOΣ] FROM THLEFONA WHERE KOD='" + data1.Recordset("KOD") + "' AND EIDOS='e' GROUP BY YEAR(HME),MONTH(HME) ORDER by YEAR(HME) DESC,MONTH(HME) DESC"
End If

120     data2.Refresh

        '<EhFooter>
        Exit Sub

THLEFONA_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.THLEFONA_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.THLEFONA_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub TIMText3_DblClick()


 Dim FF As New APOT2
        FF.kod.Text = timtext3.Text   ' "10"
        FF.CMDEXIT.Enabled = False
        FF.anazitisi 0
        FF.SHOW
        'MsgBox "ok"
       ' APOT2.kod.SetFocus
180   '  SendKeys "{ENTER}"




End Sub

Private Sub TIMText3_LostFocus()

        '<EhHeader>
        On Error GoTo TIMText3_LostFocus_Err

        '</EhHeader>

100     timtext3.Visible = False

        '<EhFooter>
        Exit Sub

TIMText3_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.TIMText3_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.TIMText3_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub XRONIES_Change()

    '<EhHeader>
    On Error GoTo XRONIES_Change_Err

    '</EhHeader>
    Dim DUM

    ' DUM = KARTELLA(gDir)
    ' DBGrid2_RowColChange 0, 0

    '<EhFooter>
    Exit Sub

XRONIES_Change_Err:
    'MsgBox Err.Description & vbCrLf & _
     "in ADOMERCNEW.pelat4.XRONIES_Change " & _
     "at line " & Erl, _
     vbExclamation + vbOKOnly, "Application Error"
    SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.XRONIES_Change " & "at line " & Erl

    Resume Next

    '</EhFooter>

End Sub

Private Sub XRONIES_Click()

        '<EhHeader>
        On Error GoTo XRONIES_Click_Err

        '</EhHeader>

100     DBGrid2_RowColChange 0, 0

        '<EhFooter>
        Exit Sub

XRONIES_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat4.XRONIES_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat4.XRONIES_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub write_query(sql As String)


 On Error Resume Next

    'SAVE_ERROR Err.Description & " in Project1.Form1.cmdCommand2_Click " & " at line " & Erl
    Dim F As Integer

    F = FreeFile
    Open "C:\MERCVB\ERR.TXT" For Append As #F
    Write #F, Format(Now, "DD/MM/YYYY HH:MM") + sql

    Close #F
End Sub
