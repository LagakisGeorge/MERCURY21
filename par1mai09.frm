VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form par1 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   1590
   ClientWidth     =   14925
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9570
   ScaleMode       =   0  'User
   ScaleWidth      =   15904.33
   WindowState     =   2  'Maximized
   Begin VB.TextBox TIMText3 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   45
      TabIndex        =   23
      Top             =   4815
      Width           =   75
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   3615
      Left            =   60
      TabIndex        =   66
      Top             =   1245
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   6376
      _Version        =   393216
      BackColor       =   16744576
      Rows            =   300
      Cols            =   20
      BackColorSel    =   16744576
      _NumberOfBands  =   1
      _Band(0).Cols   =   20
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Bindings        =   "par1.frx":0000
      Height          =   1755
      Left            =   -15
      TabIndex        =   65
      Top             =   8085
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3096
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
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   13160660
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
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
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
   Begin MSAdodcLib.Adodc LastTimol 
      Height          =   330
      Left            =   7575
      Top             =   10305
      Visible         =   0   'False
      Width           =   3630
      _ExtentX        =   6403
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
      Caption         =   "LastTimol"
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
   Begin VB.CommandButton Command6 
      Caption         =   "Διόρθωση"
      Height          =   210
      Left            =   6060
      TabIndex        =   64
      Top             =   690
      Width           =   915
   End
   Begin VB.CheckBox metasx 
      BackColor       =   &H00FF0000&
      Caption         =   "Μετασχ.Παραστατικού"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8835
      TabIndex        =   63
      ToolTipText     =   "Με διόρθωση μπορώ να μετατρέψω το παραστατικό χωρίς να χαθεί το αρχικό παραστατικό"
      Top             =   8085
      Width           =   2100
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Επανεκτύπωση Τελευταίου Παραστατικού"
      Height          =   615
      Left            =   8775
      TabIndex        =   58
      Top             =   9180
      Width           =   1485
   End
   Begin VB.CommandButton Command4 
      Height          =   240
      Left            =   7035
      TabIndex        =   57
      Top             =   990
      Width           =   165
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kαρτέλα"
      Height          =   210
      Left            =   3210
      TabIndex        =   56
      Top             =   720
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid GridEidon 
      Bindings        =   "par1.frx":0018
      Height          =   2715
      Left            =   15
      TabIndex        =   29
      Top             =   4845
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   4789
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   11064319
      HeadLines       =   1
      RowHeight       =   20
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
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox isotimia 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   1200
      TabIndex        =   54
      Text            =   "1"
      Top             =   7635
      Width           =   690
   End
   Begin VB.TextBox metaf 
      Height          =   240
      Left            =   4935
      TabIndex        =   50
      Top             =   7605
      Width           =   825
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   2
      Left            =   5265
      TabIndex        =   11
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   7185
      TabIndex        =   12
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton selectPEL 
      Caption         =   "Πελάτης"
      Height          =   360
      Left            =   6000
      TabIndex        =   4
      Top             =   240
      Width           =   870
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   1425
      TabIndex        =   10
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   105
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox POLAGOR 
      Height          =   315
      ItemData        =   "par1.frx":002A
      Left            =   120
      List            =   "par1.frx":0037
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.ComboBox PARAS 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "par1.frx":005B
      Left            =   1560
      List            =   "par1.frx":0068
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.ComboBox SEIRA 
      BackColor       =   &H000000FF&
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      Left            =   8160
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox APOTH1 
      Height          =   315
      Left            =   9105
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox APOTH2 
      Height          =   315
      Left            =   9105
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   555
      Width           =   1215
   End
   Begin VB.TextBox parat 
      Height          =   285
      Left            =   6945
      TabIndex        =   38
      Top             =   7215
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ΖΥΓΟΣ"
      Height          =   255
      Left            =   -165
      TabIndex        =   37
      Top             =   9465
      Width           =   780
   End
   Begin VB.ComboBox AYTOKINHTO 
      Height          =   315
      Left            =   7335
      TabIndex        =   36
      Top             =   6570
      Width           =   1485
   End
   Begin VB.CheckBox kostologhsh 
      Caption         =   "Κοστολόγηση"
      Height          =   255
      Left            =   8865
      TabIndex        =   35
      Top             =   6960
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Λήψη Τιμ.ASCII"
      Height          =   240
      Left            =   8865
      TabIndex        =   34
      Top             =   4860
      Width           =   1470
   End
   Begin VB.CheckBox lianikh 
      Caption         =   "Λιανική"
      Height          =   255
      Left            =   8865
      TabIndex        =   33
      Top             =   6720
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox ektyp 
      Caption         =   "Εκτύπωση"
      Height          =   255
      Left            =   8865
      TabIndex        =   32
      Top             =   6480
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   7350
      TabIndex        =   30
      Top             =   6915
      Width           =   1485
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2475
      Top             =   9585
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      Destination     =   1
      PrintFileType   =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlxGridDeltion 
      Bindings        =   "par1.frx":008C
      Height          =   2655
      Left            =   240
      TabIndex        =   28
      Top             =   4920
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4683
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc DOK 
      Height          =   330
      Left            =   3675
      Top             =   9585
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "DOK"
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
   Begin MSAdodcLib.Adodc EGGTIM 
      Height          =   375
      Left            =   2115
      Top             =   9105
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "EGGTIM"
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
   Begin MSAdodcLib.Adodc PEL 
      Height          =   375
      Left            =   75
      Top             =   9105
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "PEL"
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
   Begin MSAdodcLib.Adodc EID 
      Height          =   330
      Left            =   -45
      Top             =   9465
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "EID"
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
   Begin VB.Data abat2 
      Caption         =   "Data1"
      Connect         =   "dBASE III;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5610
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data abat 
      Caption         =   "Data1"
      Connect         =   "dBASE III;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6555
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10530
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data temp 
      Caption         =   "Data1"
      Connect         =   "dBASE III;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7890
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9450
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox MACtEXT 
      Height          =   285
      Left            =   6795
      TabIndex        =   26
      Text            =   "Text5"
      Top             =   9705
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Data CMEM 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\mercvb\reports\reports.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6555
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "mem"
      Top             =   9225
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data cEGGTIM 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\mercvb\reports\reports.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7755
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9225
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data CTIM 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\mercvb\reports\reports.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5235
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9225
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CheckBox ApoParaggelia 
      Caption         =   "Από παραγγελία"
      Height          =   255
      Left            =   6825
      TabIndex        =   24
      Top             =   4905
      Width           =   1995
   End
   Begin VB.ComboBox PROORISMOS 
      Height          =   315
      ItemData        =   "par1.frx":009E
      Left            =   6285
      List            =   "par1.frx":00A0
      TabIndex        =   20
      Top             =   6240
      Width           =   2535
   End
   Begin VB.ComboBox FORTOSH 
      Height          =   315
      ItemData        =   "par1.frx":00A2
      Left            =   6285
      List            =   "par1.frx":00A4
      TabIndex        =   19
      Top             =   5880
      Width           =   2535
   End
   Begin VB.ComboBox SKOPOS 
      Height          =   315
      ItemData        =   "par1.frx":00A6
      Left            =   6285
      List            =   "par1.frx":00A8
      TabIndex        =   18
      Top             =   5520
      Width           =   2535
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "par1.frx":00AA
      Left            =   6285
      List            =   "par1.frx":00AC
      TabIndex        =   17
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton metasxhmatismos 
      Caption         =   "Μετασχηματισμός"
      Height          =   420
      Left            =   5880
      TabIndex        =   16
      Top             =   7560
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.CommandButton b_akyr 
      Caption         =   "Εξοδος"
      Height          =   420
      Left            =   9180
      TabIndex        =   22
      Top             =   7560
      Width           =   1140
   End
   Begin VB.CommandButton b_Katax 
      Caption         =   "Καταχώρηση"
      Height          =   420
      Left            =   7680
      TabIndex        =   21
      Top             =   7560
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   4740
      TabIndex        =   3
      Top             =   240
      Width           =   1035
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   6915
      TabIndex        =   5
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   70582273
      CurrentDate     =   38294
   End
   Begin MSDataGridLib.DataGrid GridPelaton 
      Bindings        =   "par1.frx":00AE
      Height          =   3600
      Left            =   45
      TabIndex        =   27
      Top             =   1245
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   6350
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   6
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Προορ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   5805
      TabIndex        =   62
      Top             =   6255
      Width           =   450
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Φορτ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   5790
      TabIndex        =   61
      Top             =   5880
      Width           =   450
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Σκοπό"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   5820
      TabIndex        =   60
      Top             =   5550
      Width           =   450
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Πληρ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   5820
      TabIndex        =   59
      Top             =   5220
      Width           =   450
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   315
      Left            =   30
      Top             =   7590
      Width           =   5745
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   10440
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Ισοτιμία"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   60
      TabIndex        =   55
      Top             =   7620
      Width           =   1065
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "F5:Τελευταία Τιμή     F8:Ποσότητα από Αξία      F9:Τιμή χωρίς ΦΠΑ"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   45
      TabIndex        =   53
      Top             =   7905
      Width           =   8775
   End
   Begin VB.Label Label15 
      Caption         =   ".."
      Height          =   255
      Left            =   5880
      TabIndex        =   52
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Μεταφορικά και άλλα έξοδα"
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
      Height          =   225
      Left            =   2025
      TabIndex        =   51
      Top             =   7620
      Width           =   2880
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Σειρά"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3915
      TabIndex        =   49
      Top             =   -15
      Width           =   630
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Αριθμός"
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
      Left            =   4740
      TabIndex        =   48
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Παραστατικό"
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
      Left            =   1560
      TabIndex        =   47
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Είδος"
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
      Left            =   120
      TabIndex        =   46
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF0000&
      Caption         =   "Αποθήκη"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9120
      TabIndex        =   45
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      Caption         =   "Πωλητής"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8160
      TabIndex        =   44
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Διεύθυνση"
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
      Left            =   7170
      TabIndex        =   43
      Top             =   690
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Α.Φ.Μ."
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
      Left            =   5295
      TabIndex        =   42
      Top             =   675
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Επώνυμο"
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
      Left            =   1560
      TabIndex        =   41
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label10 
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
      Left            =   240
      TabIndex        =   40
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Σχόλια"
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
      Left            =   5760
      TabIndex        =   39
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Γεν.Εκπτωση"
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
      Left            =   6030
      TabIndex        =   31
      Top             =   6945
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Αρ.Αυτοκ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6075
      TabIndex        =   25
      Top             =   6630
      Width           =   1215
   End
   Begin VB.Label SynoloG 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8850
      TabIndex        =   15
      Top             =   6000
      Width           =   1470
   End
   Begin VB.Label sfpa 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8850
      TabIndex        =   14
      Top             =   5520
      Width           =   1470
   End
   Begin VB.Label Synolo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8850
      TabIndex        =   13
      Top             =   5100
      Width           =   1470
   End
   Begin VB.Menu menupopup 
      Caption         =   "menupopup"
      Visible         =   0   'False
      Begin VB.Menu par2show 
         Caption         =   "Προβολή Παραστατικού"
      End
      Begin VB.Menu diortosis 
         Caption         =   "Διόρθωση/Ακυρωτικό"
      End
   End
End
Attribute VB_Name = "par1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' FORMA 41
Option Explicit
' σταθερές για τις στήλες της τιμολόγησης


Const f_k = 1
Const f_o = 2
Const f_p = 3
Const F_T = 4
Const f_e = 5
Const f_f = 6
Const f_a = 7
Const F_M = 8
Const F_mefpa = 9


Dim fbuff As String

Dim XROMATA(100, 12, 12) 'κρατάει σε πίνακα τα χρώματα



Dim F_TRANS As Long ' 1=BEGINTRANS   - ROLLBACK      0=NOTHING
Dim F_EIKONA As Long '
Dim f_mono_apoParaggelia As Integer ' mono apo Paraggelia TIMOLOGO

Dim f_SYMPYKNVSH As Integer ' ΣΥΜΠΥΚΝΩΣΗ ΚΩΔΙΚΩΝ


Dim F_SHMADEMENA_DELTIA As String ' ΔΕΛΤΙΑ ΠΟΥ ΘΑ ΤΙΜΟΛΟΓΗΘΟΥΝ
Dim F_MONOMEYPOLOIPO As Integer

Dim f_barcode As Integer '1=diabazei apeyueias ta barcode
Dim f_LIANIKHCHECK As Integer 'ΤΣΕΚΑΡΙΣΜΕΝΗ Η ΛΙΑΝΙΚΗ
Dim f_TIMOKAT As Integer ' τιμοκατάλογος
Dim F_PARAS As String ' ΚΕΙΜΕΝΟ ΤΟΥ ΠΑΡΑΣΤΑΤΙΚΟΥ
Dim f_apeyueias  As String
Dim F_add_barc
Dim f_BARCODE_STOTIM As Integer   ' 1=ΒΑΖΕΙ ΤΟ ERG ΣΤΟ ΤΙΜΟΛΟΓΙΟ 0=DEFAULT ΤΟΝ ΚΩΔΙΚΟ

Dim F_falcon_Polla_tim  As Integer ' pairnei polla timologia apo to falcon
Dim f_PALIAXRONIA
Dim f_FORM_EKTY 'parastat!FORM_EKTY  φορμα εκτύπωσης=99  1=crystal -> timol1 timol1an
Dim f_SeiresTimologioy 'POSA EIDH MPORO NA PERASO

Dim f_SynSeiresTimologioy 'γραμμενες+κενες σειρές

Dim f_MaxSeiresTimologioy 'POSA EIDH MPORO NA PERASO Max TIMOLOGIA Agoras
Dim f_MinSeiresTimologioy 'POSA EIDH MPORO NA PERASO Min TIMOLOGIA polisis

Dim f_AutoChangeLine As Long 'def="1"  "1=μπορω να αλάξω έκπτωση γραμμής 0=δεν μπορώ
Dim f_CanChangeEkptosi As Long 'def="1"  "1=μπορω να αλάξω έκπτωση γραμμής 0=δεν μπορώ
Dim f_CanChangeFpa As Long 'def="1"  "1=μπορω να αλάξω το φπα γραμμής 0=δεν μπορώ

Dim f_checkYpol As Long 'def="0"  "1=ΕΛΕΓΧOΣ 2=ΑΠΑΓΟΡΕΥΣΗ 0=ΤΙΠΟΤΑ")) ελεγχος υπολοιπου
Dim f_psifia_atim   As Long 'DEF=5  T12345A   KALA EINAI NA EXO 6 T123456A
Dim f_PARAGGMARKUP 'paraggelia  pelatoy  0=xondriki   1=me mark up
Dim f_psifiaAjias 'posa psifia tha exei h kathe seira
Dim f_1choice_polhseis, f_1choice_agores ' οταν διαλέγει αγορές ή πωλήσεις σε ποιό παραστατικό να κάθεται

Dim f_dior_mAtim, f_dior_HME ' το τιμολόγιο που θα πάω να διορθώσω (ενδέχεται να το αλλάξω αριθμό ή ημερομηνία και έτσι χάνω το αρχικό που θα διόρθωνα )

Dim F928 As Integer


Dim F_STOPDELTIA  As Integer  'ΝΑ ΜΠΛΟΚΑΡΕΙ ΑΝ ΕΚΚΡΕΜΟΥΝ ΔΕΛΤΙΑ
Dim F_2CHFIA_ZYGIZOMENA  As String   '2 ΠΡΩΤΑ ΨΗΦΙΑ ΒΟΗΘ.ΚΩΔ.ΤΩΝ ΖΥΓΙΖΟΜΕΝΩ



Dim f_mc As Integer ' σειρες grid γεφυρωσης
Dim f_gkod_pel
Dim SumFpa(9)  ' συνολο ΦΠΑ ΚΑΤΑ ΣΥΝΤΕΛΕΣΤΗ
Dim f_AKYR As Boolean

Dim F_DIOR ' ΑΝ ΕΙΝΑΙ ΔΙΟΡΘΩΣΗ =1    ΑΝ ΟΧΙ ΔΙΟΡΘΩΣΗ=0
Dim f_error_gefyr As Integer  ' αν είναι 0 δεν έχει λάθος 1=λάθος στον κωδικό πελάτη
                              ' 2=λάθος στον κωδικό του είδους
Dim f_gDek_ajia
Dim f_sajia(300), f_fpa(300) ' κρατά στην μνήμη τοσύνολο κάθε σειράς και το ΦΠΑ
Dim fSCR As Object ' SCRIPT CONTROL
Dim FlagGridEidon
Dim f_problebgef
Dim f_matim As String
Dim f_checkPos
Dim f_dat As Date
Dim F_POLAGOR_LastValue

Dim f_use_Sysk  As Integer ' χρησιμοποιω συσκευασια = 1

Dim F_LOGPLHR(20) As String ' ΛΟΓΑΡΙΑΣΜΟΙ ΤΡΟΠΟΩΝ ΠΛΗΡΩΜΗΣ


Dim f_mfpa(1 To 9) As Single 'οριζονται στο find_parastat

Dim f_kau(1 To 9) As Single 'καθαρή αξία κατά φπα
Dim f_gfkod(20) As String, f_gfaj(20) As Single

'METABLHTES EIDOYS PARASTATIKOY APO PARASTAT
Dim F_TITLOS, F_POS_APOU, F_STADIO, F_AJIA_APOU, F_XONDR, f_pel, f_metasx, f_pol, F_rec_p As Integer, F_rec_m As Integer 'παράμετροι του παραστατικού

Dim f_deltia(1 To 30) 'δελτια που σημαδεύω
Dim f_IDdeltia(1 To 30) As Long  'ID δελτιων  που σημαδεύω για να μπορώ εύκολα να τα μαρκάρω

Dim F_File_Timologioy As String ' το όνομα του αρχείου που γίνεται η εκτύπωση π.χ. c:\print

Dim F_Type_File_Timologioy As Integer '1=typvnei to arxeio 0=den to typonei

Dim f_F12 As Integer, f_F1 As Integer  ' ΠΡΟΓΡΑΜΜΑΤΙΣΜΟΣ ΠΛΗΚΤΡΩΝ

Dim F_EXO_XROMATA As Integer  '1=ΧΡΗΣΙΜΟΠΟΙΩ ΤΑ ΤΡΩΜΑΤΑ ΜΕΓΕΘΗ

Dim F_PAR_MONO_ME_YPOLOIPO As Integer  '
Dim f_Fkey ' το πληκτρο που κάνει update

Dim f_Desmia As Integer ' χρησιμοποιω τις παραγγελίες για έλεγχο υπολοίπου  'βλεπε mdesmia

Dim f_suma ' teliko poso timologioy
Dim f_ColUpdate  ' flag αν ενημερώθηκε η κολόνα του dbgrid1
'Dim f_epitr_seir
Dim fDB As Database  ' ανοιγω γενικά την database και με την open_data ανοίγω τον πίνακα
Dim ProthFora As Boolean
Dim fSHMADEMENA As String
Dim mEidos_Par(50) 'ειδος παραστατικού . π.χ.  atim=T


'dim constant a

'

Private Sub b_akyr_Click()
Dim ANS As Integer

ANS = vbYes
If Len(Grid1.TextMatrix(1, 1)) > 0 Then
  ANS = MsgBox("Να χαθούν οι αλλαγές;", vbYesNo)
End If

 If ANS = vbYes Then
    
      Unload Me
      SendKeys "%"
      SendKeys "{RIGHT}+{RIGHT}+{RIGHT}+{DOWN}"
 End If
  
  
End Sub
Private Sub b_Katax_Click()
'====================================================


Dim wrkDefault As Workspace
Dim R As New ADODB.Recordset
Dim X, DUM
Dim f_atim
Dim K, L
Dim eidoskpe
Dim m_tim, S_POSOTHTA As Single
Dim par As New ADODB.Recordset, m_r
Dim m_m
Dim sql As String
Dim s1 As Single





'βρισκει το είδος του παραστατικού
'γιατί πολλές φορές το πουλάει σαν λιανική ενώ είναι χονδρική πώληση
 find_eid_parastat




' X = mEidos_Par(1 + PARAS.ListIndex)
 
 




f_AKYR = False





On Error GoTo MHNYMA
Dim Mf_matim  As String
Dim ans0 As Integer

' bazo +1 gia na me αφηνει την τελευταία μέρα να γράφω
If DTPicker1.value < gEnarjh Or DTPicker1.value > gLhjh + 1 Then
    MsgBox "Ημερομηνία εκτός ορίων χρήσης"
    Exit Sub
End If

If DTPicker1.value > Now Then
   ans0 = MsgBox("Προσοχή  ημερομηνία μεγαλυτερη της σημερινής.  Συνεχίζω;", vbYesNo)
   If ans0 <> vbYes Then
      Exit Sub
   End If
End If

If DateDiff("d", DTPicker1.value, Now) < -1 Then
   ans0 = MsgBox("Προσοχή  ημερομηνία μικρότερη της σημερινής.  Συνεχίζω;", vbYesNo)
   If ans0 <> vbYes Then
      Exit Sub
   End If
End If




F_PARAS = PARAS.Text




Dim RR As New ADODB.Recordset, PARAST As String
'ελεγχω μηπως υπάρχει μεγαλύτερο νουμερο απο αυτό
'If f_FORM_EKTY > 0 And F_DIOR = 0 Then ' εκτυπωνεται
'   PARAST = Right(PARAS.Text, 1) + Right("00000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text
'   RR.Open "select * FROM TIM WHERE ATIM>'" + PARAST + "' AND LEFT(ATIM,1)='" + Right(PARAS.Text, 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
'   If Not RR.EOF Then
'      MsgBox "ΥΠΑΡΧΕΙ ΠΑΡΑΣΤΑΤΙΚΟ ME ΜΕΓΑΛΥΤΕΡΟ AΡΙΘΜΟ"
'      RR.Close
'      Exit Sub
'   End If
'   RR.Close
'
'
'End If











'ελεγχος για άδειο παραστατικο
'x = Open_Data("select sum(poso) as ss from timedit")

Dim db As Database
'If gConnect = "Access" Then
 '  Set db = OpenDatabase(gDir, False, False)
'Else
 '  Set db = OpenDatabase(gDir, False, False, gConnect)
'End If

10


S_POSOTHTA = 0

For K = 1 To Grid1.ROWS - 1
   S_POSOTHTA = S_POSOTHTA + Abs(Val(Grid1.TextMatrix(K, f_p)))
Next

If S_POSOTHTA = 0 Then  ' Val(Synolo.Caption)   ' 202329
   MsgBox "Δεν περάστηκαν είδη"
   POLAGOR.SetFocus
   Exit Sub
End If
m_tim = GetCurrentTime()



If Label1.Caption = "" Then
    MsgBox "Δεν Διαλέξατε " + IIf(fbuff = "e", "Πελάτη", "Προμηθευτή")
    Exit Sub
End If


If Val(Text1.Text) = 0 Then
    MsgBox "Δεν δωσατε αριθμό παραστατικού "
    Exit Sub
End If


If F_STADIO = 5 And APOTH1.Text = APOTH2.Text Then '' 1=ΠΡΟΣΦ 2=ΠΑΡΑΓΓ 3=ΔΑ 4=ΤΙΜ.ΠΟΛ 5=ΕΝΔΟΔΙΑΚΙΝΙΣΙ
    MsgBox "Οι Αποθήκες φόρτωσης και αποστολής πρέπει να διαφέρουν "
    Exit Sub
End If

' προσοχή τα δελτία και χωρίς αξία πρέπει να περνάνε
' If Val(Synolo.Caption) = 0 Then Exit Sub

Me.MousePointer = vbHourglass

metasxhmatismos.Enabled = True


f_dat = DTPicker1.value


f_matim = Right(PARAS.Text, 1) + Right("00000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text
20




par.Open "SELECT *FROM PARASTAT WHERE EIDOS='" + Left(f_matim, 1) + "';", Gdb, adOpenDynamic, adLockOptimistic
If IsNull(par!FORM_EKTYP) Then f_FORM_EKTY = 0 Else f_FORM_EKTY = Val(par!FORM_EKTYP)

'καταργει την διορθωση
If metasx.value = vbChecked Then F_DIOR = 0


If F_DIOR = 1 Then ' ΔΙΟΡΘΩΣΗ
   Dim ANS As Integer
   ANS = MsgBox("Να αποθηκευθεί η διόρθωση;", vbYesNo)
   If ANS = vbNo Then GoTo 145 'Exit Sub
   
        
       ANS = vbNo
      If IsNull(par!FORM_EKTYP) Then f_FORM_EKTY = 0 Else f_FORM_EKTY = Val(par!FORM_EKTYP)
      
       If f_FORM_EKTY > 0 Then '   fBuff = "e" Then
          ANS = MsgBox("Να ΕΚΔΟΘΕΙ ΑΚΥΡΩΤΙΚΟ;", vbYesNo)
       End If
       If ANS = vbYes Then
       

          
       
       
       
          f_AKYR = True
          parat.Text = "Σχ.Παρ. " + f_matim
          F_PARAS = "ΑΚΥΡΩΤΙΚΟ ΣΤΟΙΧΕΙΟ  κ"
          
          If Len(Dir("c:\BSEIRA", vbDirectory)) > 2 Then
             Mf_matim = "k" + Format(find_parastat(1, "k", 0), String(f_psifia_atim, "0"))
          Else
             Mf_matim = "κ" + Format(find_parastat(1, "κ", 0), String(f_psifia_atim, "0"))
          End If
          
          ' ELEGXOS AN YPARXEI HDH TO AKYROTIKO
          R.Open "select ATIM from TIM where ATIM='" + Mf_matim + "'", Gdb, adOpenDynamic, adLockOptimistic
          If Not R.EOF Then
             MsgBox "Υπάρχει ήδη  TO ΑΚΥΡΩΤΙΚΟ ΜΕ  ίδιο αριθμό. " + Mf_matim
             MsgBox "ΔΙΑΓΡΑΨΤΕ ΤΟ ΠΑΛΙΟ ΑΚΥΡΩΤΙΚΟ ή ΑΛΛΑΞΤΕ ΤΗΝ ΑΡΙΘΜΗΣΗ ΤΟΥ ΑΚΥΡΩΤΙΚΟΥ ΑΠΟ ΤΟ ΠΑΡΑΜΕΤΡΟΙ ΠΑΡΑΣΤΑΤΙΚΩΝ "
             Me.MousePointer = vbNormal
             Exit Sub
          End If
          
       Else
         f_AKYR = False
       End If
       
       
       
       'σβηνω το παλιό παραστατικό
       sql = "delete  "
       sql = sql + " from EGGTIM where left(ATIM," + Str(Len(f_dior_mAtim)) + ")='" + f_dior_mAtim + "' and  HME >= '" + Format(f_dior_HME, "mm/dd/yyyy") + "' and  HME <= '" + Format(DateAdd("D", 1, f_dior_HME), "mm/dd/yyyy") + "';"
       Gdb.Execute sql, L
       

       sql = "delete  "
       sql = sql + " from TIM where left(ATIM," + Str(Len(f_dior_mAtim)) + ")='" + f_dior_mAtim + "' and  HME >= '" + Format(f_dior_HME, "mm/dd/yyyy") + "' and  HME <= '" + Format(DateAdd("D", 1, f_dior_HME), "mm/dd/yyyy") + "';"
       Gdb.Execute sql, L


       sql = "delete  "
       sql = sql + " from EGG where left(ATIM," + Str(Len(f_dior_mAtim)) + ")='" + f_dior_mAtim + "' and  HME >= '" + Format(f_dior_HME, "mm/dd/yyyy") + "' and  HME <= '" + Format(DateAdd("D", 1, f_dior_HME), "mm/dd/yyyy") + "';"
       Gdb.Execute sql, L
       
       
       ' SAN MATIM EXO TO AKYROTIKO KAI BAZO SIMERINI HMEROMHNIA
       If f_AKYR = True Then
          f_matim = Mf_matim
          Text1.Text = Val(mID$(Mf_matim, 2, f_psifia_atim))
          f_dat = Now
          ' ΜΗΔΕΝΙΖΩ ΑΥΤΕΣ ΤΙΣ ΠΑΡΑΜΕΤΡΟΥΣ ΓΙΑΤΙ
          'ΘΑ ΠΑΡΕΙ ΤΙΣ ΠΑΡΑΜΕΤΡΟΥΣ ΤΟΥ ΑΡΧΙΚΟΥ ΠΑΡΑΣΤΑΤΙΚΟΥ ΚΑΙ ΘΑ ΧΑΛΑΣΕΙ ΤΗΝ ΑΠΟΘΗΚΗ
          F_POS_APOU = "0"
          F_AJIA_APOU = "0"
          f_pel = "0"

       End If
       
Else
 
 
  'diadikasia egkrishs
   If f_FORM_EKTY > 0 And Combo4.Text = Combo4.List(1) And Len(Dir("c:\mercvb\egkr.txt")) > 1 Then '<> METRHTOIS ή κάρτα
       Dim ccc As String, nnn, date1, Topos
       ccc = Right(Format(GetCurrentTime(), "#######0"), 4)
       If InputBox("Αριθμός Εγκρισης για " + ccc) = egkrisi(ccc) Then
            ' ok
       Else
           
           If InputBox("Eπανάληψη:Αριθμός Εγκρισης για " + ccc) = egkrisi(ccc) Then
             'ok
           Else
              MsgBox "Λάθος αριθμός έγκρισης"
              Exit Sub
           End If
       End If
       
       date1 = InputBox("Ημερομηνία Εξόφλησης " + Format(DateAdd("d", 30, Now), "dd/MM/YYYY"))
       Topos = InputBox("Τρόπος Εξόφλησης ")
       If Not IsDate(date1) Then
          date1 = DateAdd("d", 30, Now)
       Else
          date1 = DateAdd("h", Hour(Now), date1)
          
       End If
       
      nnn = add_schedule(date1, Text2(1).Text, Topos)
      Dim SQLMEMO
      If IsNull(PEL.Recordset("MEMO")) Then
        SQLMEMO = Format(date1, "DD/MM/YYYY") + " " + Topos + "' WHERE EIDOS='e' and KOD='" + PEL.Recordset("KOD")
      Else
        SQLMEMO = PEL.Recordset("MEMO") + Chr(13) + Format(date1, "DD/MM/YYYY") + " " + Topos + "' WHERE EIDOS='e' and KOD='" + PEL.Recordset("KOD")
      End If
    
     Gdb.Execute "UPDATE PEL SET MEMO='" + SQLMEMO + "'"
  End If
 
      'ενημερωση με τον τελευταίο αριθμό
   ' γιατί μπορεί ενδιάμεσα να κόπηκε και αλλο παραστατικό
   If F_DIOR = 0 And f_FORM_EKTY > 0 Then  ' οχι ΔΙΟΡΘΩΣΗ
      Text1.Text = find_parastat(1, Right(PARAS.Text, 1), 0)
      f_matim = Right(PARAS.Text, 1) + Right("00000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text
   End If
   If Check_Dipla = 0 Then Exit Sub


   
   
   
End If

30
   ' Get default Workspace.
  ' Set wrkDefault = DBEngine.Workspaces(0)
' Start of outer transaction.
If F_TRANS = 1 Then
  Gdb.BeginTrans
End If

If Len(f_deltia(1)) > 0 Then
   Dim S As Integer
   K = 1
   Do While f_IDdeltia(K) > 0
     Gdb.Execute "update TIM set ART='1',SXETIKO='" + f_matim + " " + Format(f_dat, "dd/mm/yyyy") + "' where ID_NUM=" + Str(f_IDdeltia(K))
     K = K + 1
   Loop
   
End If


'δημιουργία eggtim και ΤΙΜ
40 enhm_eggtim


If f_AKYR = False Then  ' αν δεν ειναι ακυρωτικό

'ενημέρωση αποθήκης
50 enhm_apot

'ενημέρωση πελάτη
60 enhm_pel

'--------------   ενημέρωση μετρητων  ----------------------
Dim Flag_Plir As Boolean
Flag_Plir = False
If Val(Left(Combo4.Text, 2)) > 0 Then
   If F_LOGPLHR(Val(Left(Combo4.Text, 2)) - 1) <> "0000" Then ' αν εχει ορισθει λογαριασμοσ πιστωτικων καρτων
      Flag_Plir = True
   End If
End If
   


If Combo4.Text = Combo4.List(0) Or Flag_Plir Then ' METRHTOIS ή κάρτα
   enhm_met
End If

End If

Me.Caption = Str(Abs(m_tim - GetCurrentTime()))


70
   
   

'ΕΝΗΜΕΡΩΣΗ ΑΡΙΘΜΟΥ ΠΑΡΑΣΤΑΤΙΚΟΥ
   
   
   
   
'   par.Open "SELECT *FROM PARASTAT WHERE EIDOS='" + left(f_matim, 1) + "';", Gdb, adOpenDynamic, adLockOptimistic
   
   ' If IsNull(par!FORM_EKTYP) Then f_FORM_EKTY = 0 Else f_FORM_EKTY = Val(par!FORM_EKTYP)
  

'---------------γεφυρώσεις-------------------------------------
If Len(gdirlog) > 2 And f_error_gefyr = 0 Then
  eidoskpe = PEL.Recordset("eidos") + PEL.Recordset("kod")
  'ΒΡΙΣΚΩ ΑΠΟ ΤΙΣ GEFYRES.DBF ΤΗΝ ΚΑΤΑΛΛΗΛΗ ΕΓΓΡΑΦΗ
         
       f_mc = 0 ' ΣΕΙΡΕς GRID ΦΕΦΥΡΩΣΗΣ
   
  
   'η χρέωση
'EIDOS     C  1  0N
'TITLOS    C 30  0N
'POL       C  1  0N
'POS_APOU  C  1  0N
'AJIA_APOU C  1  0N
'PEL       C  1  0N
'SYNT_TITL C  2  0N
'EKTYP     C  1  0N
'DESMIA    C  1  0N
'FORM_EKTYPC  2  0N
'STADIO    C  1  0N
'ETIK      C  1  0N
'METASX    C  1  0N
'XONDR     C  1  0N
'GEF_P     N  3  0N
'GEF_M     N  3  0N
'ARITMISI  N  5  0N
'SHM1      C 60  0N

   
   

   m_r = par("gef_p")
   m_m = par("gef_m")
    
    DUM = Gefyroma(f_dat, f_matim, eidoskpe, f_gfkod, f_gfaj, m_r, "Σ")
   ' τα μετρητά
   m_r = m_m
   DUM = Gefyroma(f_dat, f_matim, eidoskpe, f_gfkod, f_gfaj, m_r, "Ε")
   
   db.Close
   
    
End If


'ΜΑΡΚΆΡΩ ΤΑ ΔΕΛΤΙΑ ΠΟΥ ΧΡΗΣΙΜΟΠΟΙΗΣΑ
If Len(fSHMADEMENA) > 5 Then
  Gdb.Execute "UPDATE TIM set ART='1' where ATIM+CONVERT(CHAR(10),hme,103)  in (" + fSHMADEMENA + ");"
End If


            
If F_TRANS = 1 Then

        ' Ask if the user wants to commit to all the changes
        ' made above.
        
        
        If MsgBox("Nα αποθηκευθεί το παραστατικό", vbYesNo) = vbYes Then
             Gdb.CommitTrans
        Else
            Gdb.RollbackTrans
            Me.MousePointer = vbNormal
            Exit Sub
        End If

End If





90
'εκτύπωση
If ektyp.value = 1 Then


If f_FORM_EKTY >= 90 Then ' φορμα τύπου f90.txt
100      DUM = printCrystal(f_matim, CDate(DTPicker1.value)) 'δημιουργει τα προσωρινά αρχεία
110      DUM = print_timol(f_matim, CDate(DTPicker1.value))
Else ' crystal reports timol1.rpt(σφραγίδα)   timol1an.rpt (αντίγραφο)
   
   If f_FORM_EKTY > 0 Then '      PEL.Recordset("eidos") = "e" Then


        ' On Error GoTo 0 ' na sbhstei 15-6-2006
      If f_FORM_EKTY = 15 Then     '  ME SQL SERVER H EKTYPOSI
         DUM = PrintSqlCrystal(f_matim, CDate(f_dat)) 'DTPicker1.Value)) ''δημιουργει τα προσωρινά αρχεία
      Else  ' ME MDB
120      DUM = printCrystal(f_matim, CDate(f_dat)) 'DTPicker1.Value)) ''δημιουργει τα προσωρινά αρχεία
      End If

      '   On Error Resume Next
      
     If DUM < -900 Then
         MsgBox "αδυνατη η εκτύπωση" + Chr(13) + "Διορθώστε το παραστατικό για να ξανατυπωθεί"
         Exit Sub
     End If
     
      
      
      
      
      
    Dim xxx As Printer

    For Each xxx In Printers
       If InStr(UCase(xxx.DeviceName), "DESKJET") > 0 Then
           ' Set printer as system default.
           Set Printer = xxx
           ' Stop looking for a printer.
           Exit For
        End If
     Next
      
      
      
      
    If f_FORM_EKTY = 12 Then     '     λιανικη
       CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol2.rpt"
    
      ElseIf f_FORM_EKTY = 13 Then     '     λιανικη
       CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol3.rpt"
    
    
    
    ElseIf f_FORM_EKTY = 15 Then     '     λιανικη
       CrystalReport1.ReportFileName = "C:\Mercvb\reports\timolsql.rpt"
       CrystalReport1.Connect = gConnect
    
    Else
       CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
    End If
     
    If f_FORM_EKTY = 10 Then
       CrystalReport1.Action = 1
    ElseIf f_FORM_EKTY = 20 Then
    
       CrystalReport1.Destination = crptToWindow
       CrystalReport1.Action = 1
    ElseIf f_FORM_EKTY = 18 Then
       CrystalReport1.Action = 1
       CrystalReport1.Action = 1
       CrystalReport1.Action = 1
    Else
       CrystalReport1.Action = 1
       CrystalReport1.Action = 1
    End If
 
 
 End If   ' If f_FORM_EKTY > 0 Then '
   
End If  ' If ektyp.Value = 1 Then



End If



'ENHMERVSH ARIUMOY PARASTATIKOY
If f_AKYR Then ' ΑΝΗΜΕΡΩΣΗ ΑΚΥΡΩΤΙΚΟΥ
     DUM = find_parastat(1, Left(f_matim, 1), 1)
Else
   If F_DIOR = 1 Then ' ΔΙΟΡΘΩΣΗ   DEN ΕΝΗΜΕΡΩΝΕΙ ΑΡΙΘΜΗΤΗΡΑ
   Else
        Upd_Parastat
   End If
End If





'ENHMERVSH ARIUMOY PARASTATIKOY
'Upd_Parastat
'    b_Katax.Enabled = False
    
             
'-----------------------------------------------------------------
'--------------------  initialization ----------------------------
'-----------------------------------------------------------------
145
Synolo.Caption = ""
sfpa.Caption = ""
SynoloG.Caption = ""
         
For K = 1 To 9: f_kau(K) = 0: SumFpa(K) = 0: Next ' συνολο ΦΠΑ ΚΑΤΑ ΣΥΝΤΕΛΕΣΤΗ

For K = 1 To 100: f_fpa(K) = 0: f_sajia(K) = 0: Next ' κρατά στην μνήμη τοσύνολο κάθε σειράς και το ΦΠΑ

For K = 1 To 30: f_IDdeltia(K) = 0: f_deltia(K) = "": Next 'δελτια που σημαδεύω


' μηδενιζω το grid1
For L = 1 To Grid1.ColS - 1
    'Grid1.Col = l
    For K = 1 To Grid1.ROWS - 1
       'Grid1.Row = K
       'Grid1.Text = ""
       Grid1.TextMatrix(K, L) = ""
    Next
Next

  Text2(0).Enabled = True
  Text2(1).Enabled = True
  Text2(2).Enabled = True
  

Text2(0).Text = "" ' κωδ
Text2(1).Text = "" 'επωνυμια
Text2(2).Text = "" 'αφμ
parat.Text = "" 'ΠΑΡΑΤΗΡΗΣΕΙΣ


timText3.Top = 18000
APOTH2.Enabled = True
F_DIOR = 0 ' ΔΙΟΡΘΩΣΗ

LastTimol.Refresh




180
POLAGOR.SetFocus
'PARAS.SetFocus


'kostologhsh  aploma othonis
'Gdb.Execute "DROP TABLE KOSTOL" + gXEIRISTHS
'Gdb.Execute "select EID.KOD,EID.ONO,EID.XTI,EID.POS_KERD,EID.LTI,EID.LTI5,EGGTIM.TIMM  INTO KOSTOL" + gXEIRISTHS + " from EGGTIM inner join EID on EGGTIM.KODE=EID.KOD where left(ATIM,7)='" + f_matim + "'"
'EID.RecordSource = "SELECT * FROM  KOSTOL" + gXEIRISTHS
'EID.Refresh

Me.MousePointer = vbNormal


If Val(f_apeyueias) > 0 Then
  PARAS_LostFocus
  Text2(0) = f_apeyueias
  SendKeys "{ENTER} {ENTER}"
End If









Exit Sub

MHNYMA:
HandleError "Par1:b_katax_click"
Resume Next
            

End Sub
Function egkrisi(ByVal ccc As String) As Long
     egkrisi = Val(mID(ccc, 1, 1)) + Val(mID(ccc, 2, 1)) + Val(mID(ccc, 3, 1)) + Val(mID(ccc, 4, 1))
     egkrisi = Int(egkrisi / 2) + 135
End Function





Sub elegxoi()
'================================================================
'=======   DEMO TA  TRANSACTIONS  ===============================
'================================================================

   Dim strName As String
   Dim strMessage As String
   Dim wrkDefault As Workspace
   Dim dbsNorthwind As Database
   Dim rstEmployees As Recordset

   ' Get default Workspace.
   Set wrkDefault = DBEngine.Workspaces(0)
   Set dbsNorthwind = OpenDatabase("Northwind.mdb")
   Set rstEmployees = _
      dbsNorthwind.OpenRecordset("Employees")

   ' Start of outer transaction.
   wrkDefault.BeginTrans
   ' Start of main transaction.
   wrkDefault.BeginTrans

   With rstEmployees

      ' Loop through recordset and ask user if she wants to
      ' change the title for a specified employee.
      Do Until .EOF
         If !title = "Sales Representative" Then
            strName = !LastName & ", " & !FirstName
            strMessage = "Employee: " & strName & vbCr & _
               "Change title to Account Executive?"

            ' Change the title for the specified employee.
            If MsgBox(strMessage, vbYesNo) = vbYes Then
               .Edit
               !title = "Account Executive"
               .Update
            End If
         End If

         .MoveNext
      Loop

      ' Ask if the user wants to commit to all the changes
      ' made above.
      If MsgBox("Save all changes?", vbYesNo) = vbYes Then
         wrkDefault.CommitTrans
      Else
         wrkDefault.Rollback
      End If

      ' Print current data in recordset.
      .MoveFirst
      Do While Not .EOF
         Debug.Print !LastName & ", " & !FirstName & _
            " - " & !title
         .MoveNext
      Loop

      ' Roll back any changes made by the user since this is
      ' a demonstration.
      wrkDefault.Rollback
      .Close
   End With

   dbsNorthwind.Close

'End Sub



'1.να μην υπάρχουν ΦΠΑ>9 και <=0
End Sub
Sub Upd_Parastat()

Dim K As Integer
Dim db As Database
Dim R As Recordset, X As String
Dim DUM As Integer


For K = 0 To SEIRA.ListCount
   If SEIRA.Text = SEIRA.List(K) Then
      Exit For
   End If
Next


If K <= SEIRA.ListCount Then
   DUM = find_parastat(K, Right(F_PARAS, 1), 1)
End If

End Sub
Function Check_Dipla()
'************************************************************************
' αν υπάρχει ήδη ίδιο παραστατικό  με ίδιο αριθμό  γυρνάει πίσω
' στους πελάτες αν υπάρχει ό ίδιος αριθμός με παλιά ημερομηνία μας ενημερώνει
Dim db As Database, re As New ADODB.Recordset, rt As New ADODB.Recordset
Dim ANS As Integer


     Check_Dipla = 1

'If gConnect = "Access" Then
 '  Set db = OpenDatabase(gDir, False, False)
'Else
'   Set db = OpenDatabase(gDir, False, False, gConnect)
'End If
  
  If fbuff = "e" And f_FORM_EKTY > 0 Then  ' φορμα τύπου f90.txt Then
      
      re.Open "select ATIM,HME from TIM where ATIM='" + f_matim + "'", Gdb, adOpenDynamic, adLockOptimistic
      If Not re.EOF Then
          re.MoveFirst
          Do While Not re.EOF
             If re("hme") = f_dat Then
               MsgBox "Υπάρχει ήδη  παραστατικό με ίδιο αριθμό και ημερομηνία " '        Format(re("hme"), "dd/mm/yyyy")
               Check_Dipla = 0
               Exit Function
             End If
             re.MoveNext
          Loop
          
          re.MoveFirst
          ANS = MsgBox("Να συνεχισθεί η καταχώριση;", vbYesNo, "Υπάρχει ήδη  παραστατικό με ίδιο αριθμό στις " + Format(re("hme"), "dd/mm/yyyy"))
          
          If ANS = vbYes Then
              Check_Dipla = 1
          Else
              Check_Dipla = 0
          End If
          Exit Function
      End If
      re.Close
      
  
'ΕΛΕΓΧΟΣ ΑΝ ΥΠΑΡΧΕΙ ΤΟ ΠΡΟΗΓΟΥΜΕΝΟ ΠΑΡΑΣΤΑΤΙΚΟ
      
      Dim reco As String
      reco = "select * from TIM where ATIM LIKE '" + Left(f_matim, 1) + Format(Val(mID$(f_matim, 2, f_psifia_atim) - 1), String(f_psifia_atim, "0")) + mID$(f_matim, 8, 1) + "%'"
        'SELECT * FROM TIM WHERE ATIM LIKE 'T0005%'
      re.Open reco, Gdb, adOpenForwardOnly, adLockReadOnly
      If re.EOF Then
          reco = "Δεν υπάρχει ο προηγούμενος αριθμός. " + Format(Val(mID$(f_matim, 2, f_psifia_atim)) - 1, String(f_psifia_atim, "0")) + " Να συνεχισθεί η καταχώριση;"
          ANS = MsgBox(reco, vbYesNo)
          
          If ANS = vbYes Then
              Check_Dipla = 1
          Else
              Check_Dipla = 0
          End If
          
          Exit Function
      Else ' YPARXEI ALLA DEN PREPEI NA EXO MIKROTERI HMEROMHNIA APO TO PROHGOYMENO
      
          If DTPicker1.value < re("HME") Then
             MsgBox "η ημερομηνια είναι μικρότερη από την ημερομηνία του  προηγούμενου παραστατικού "
             Check_Dipla = 0
          End If
      End If
  Else 'ΠΡΟΜΗΘΕΥΤΕΣ
        re.Open "select ATIM from TIM where HME='" + Format(f_dat, "mm/dd/yyyy") + "' and ATIM='" + f_matim + "'", Gdb, adOpenDynamic, adLockOptimistic
        If re.RecordCount > 0 Then
           MsgBox "Υπάρχει ήδη  παραστατικό με ίδια ημερομηνία και ίδιο αριθμό"
           Check_Dipla = 0
        End If
  End If
  
End Function


Sub enhm_eggtim()
'=========================================
'δημιουργία eggtim και ΤΙΜ
'=========================================

Dim K, L, m, n As Integer
Dim MHKOS As Integer

'Dim Fdb As New ADODB.Connection
'Dim Re As New ADODB.Recordset
'Dim Rt As New ADODB.Recordset

Dim db As Database, re As New ADODB.Recordset, rt As New ADODB.Recordset
Dim meid As New ADODB.Recordset
Dim FEGG(10)
'If gConnect = "Access" Then
'   Set db = OpenDatabase(gDir, False, False)
'Else
'   Set db = OpenDatabase(gDir, False, False, gConnect)
'End If



On Error GoTo MHNYMA







For K = 1 To 9: f_kau(K) = 0: Next
  
  
    'DB.Open gDir
'  End If
10
  re.Open "SELECT  *FROM EGGTIM", Gdb, adOpenDynamic, adLockOptimistic
  rt.Open "SELECT  *FROM TIM", Gdb, adOpenDynamic, adLockOptimistic
  
  
Dim SYN_ME_FPA As Single
5 re("MIK_AJIA") = 0

For K = 1 To 10: FEGG(K) = 0: Next
'------------------------------eggtim -----------------------------------------
For K = 1 To Grid1.ROWS - 1

   If Len(Grid1.TextMatrix(K, f_k)) > 0 Then

        meid.Open "SELECT * FROM EID WHERE KOD='" + Grid1.TextMatrix(K, f_k) + "'", Gdb, adOpenDynamic, adLockOptimistic 'adOpenForwardOnly, adLockReadOnly

    
    'ΠΡΟΣΘΕΤΕΙ ΜΙΑ ΕΓΓΡΑΦΗ ΓΙΑ ΚΑΘΕ ΧΡΩΜΑ ή ΜΕΓΕΘΟΣ
    If F_EXO_XROMATA = 0 Then
        GoSub ADD_RECORD
    Else
         Dim K2 As Integer, k3 As Integer
         For K2 = 1 To 12: For k3 = 1 To 12
           If Val(XROMATA(K, K2, k3)) <> 0 Then
              GoSub ADD_RECORD
              XROMATA(K, K2, k3) = 0
           End If
         Next: Next
              
    End If


'20      re.AddNew
'        re("prood_aj") = k
'        re("hme") = f_dat
'        re("atim") = f_matim
'        ' re("mona") = n
'        re("poso") = Val(Grid1.TextMatrix(k, f_p))
'        re("timm") = Val(Grid1.TextMatrix(k, F_T))
'        ' ΕΝΗΜΕΡΏΝΕΙ ΤΟ ΚΈΡΔΟς
'        ' meid.Open "SELECT * FROM EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) + "'", Gdb, adOpenDynamic, adLockOptimistic 'adOpenForwardOnly, adLockReadOnly
'         If PEL.Recordset("EIDOS") = "e" And Val(Grid1.TextMatrix(k, F_T)) > 0 Then
'            re("KERDOS") = Val(Grid1.TextMatrix(k, F_T)) - meid("XTI")
'         End If
'         re("kode") = Grid1.TextMatrix(k, f_k)
'         re("onoma") = Left(Grid1.TextMatrix(k, f_o), 35)
'        ' ΒΑΖΕΙ ΤΗΝ ΣΥΣΚΕΥΑΣΙΑ ΜΕΣΑ ΣΤΟ ΟΝΟΜΑ ΓΙΑ ΤΙΜΟΛΟΓΗΣΗ
'       If f_use_Sysk = 1 Then
'         If meid("SYSKMAX") > 1 Then
'           MHKOS = Len(Trim(Format(meid("SYSKMAX"), "####") + "X" + Format(re("poso") / meid("SYSKMAX"), "####")))
'           re("onoma") = Left(Grid1.TextMatrix(k, f_o) + Space(35), 34 - MHKOS) + " " + Format(meid("SYSKMAX"), "####") + "X" + Format(re("poso") / meid("SYSKMAX"), "####")
'          End If
'        End If
'        re("mona") = Grid1.TextMatrix(k, F_M)  ' Right(Grid1.TextMatrix(K, f_o), 3)
'        re("kau_ajia") = Round(Val(Grid1.TextMatrix(k, f_a)), f_psifiaAjias)
'        re("MIK_AJIA") = Round(Val(Grid1.TextMatrix(k, f_a)), f_psifiaAjias) * (1 + g_Fpa(Val(Grid1.TextMatrix(k, f_f))) / 100)
'        SYN_ME_FPA = SYN_ME_FPA + re("MIK_AJIA")
'        re("fpa") = Val(Grid1.TextMatrix(k, f_f))
'        FEGG(re("FPA")) = FEGG(re("FPA")) + re("MIK_AJIA")
'        re("ekpt") = Val(Grid1.TextMatrix(k, f_e))
'        re("pelkod") = PEL.Recordset("kod")
'        re("eidos") = PEL.Recordset("eidos")
'
'        re("APOT") = Val(Left(APOTH1.Text, 2))
'         re("XRE") = 0: re("PIS") = 0 'ΑΠΟΦΥΓΗ NULL
'
'        Select Case Val(F_POS_APOU)
'        Case 1 ' +EISAGOGES
'           re("XRE") = re("POSO")
'        Case 2 ' -EISAGOGES
'           re("XRE") = -re("POSO")
'        Case 3 ' +EXAGOGES
'           re("PIS") = re("POSO")
'        Case 4 ' +EISAGOGES
'           re("PIS") = -re("POSO")
'        End Select
'        re.Update
'---------------- ενδοδιακίνηση  --------------------------------------------
 
' If F_STADIO = 5 Then ' ενδοδιακινηση  Left(f_matim, 1) = "σ" Then
'30      re.AddNew
'        re("hme") = f_dat
'        re("atim") = f_matim
'        re("poso") = Val(Grid1.TextMatrix(k, f_p))
'        re("timm") = Val(Grid1.TextMatrix(k, F_T))
'        re("kode") = Grid1.TextMatrix(k, f_k)
'        re("onoma") = Left(Grid1.TextMatrix(k, f_o), 30)
'        re("mona") = Right(Grid1.TextMatrix(k, f_o), 3)
'        re("kau_ajia") = Val(Grid1.TextMatrix(k, f_a))
'        re("fpa") = Val(Grid1.TextMatrix(k, f_f))
'        re("ekpt") = Val(Grid1.TextMatrix(k, f_e))
'        re("pelkod") = PEL.Recordset("kod")
'        re("eidos") = PEL.Recordset("eidos")
'
'        re("APOT") = Val(Left(APOTH2.Text, 2))
'        re("XRE") = 0: re("PIS") = 0 'ΑΠΟΦΥΓΗ NULL
'        re("XRE") = re("POSO")
'        re.Update
'      End If
'
       meid.Close
   End If
  
  
  
   
  
  Next
'------------------------------ TIM -----------------------------------------
're.Close
40
'rt.AddNew
'  rt("hme") = f_dat '
'  rt("atim") = f_matim
'  f_suma = 0
'  For k = 1 To 9
'     rt("aj" + LTrim(Str(k))) = Round(SumFpa(k), f_psifiaAjias)
'     f_suma = f_suma + rt("aj" + LTrim(Str(k)))
'     If k <> 5 Then
'        rt("fpa" + LTrim(Str(k))) = Round(SumFpa(k) * g_Fpa(k) / 100, f_psifiaAjias)
'        f_suma = f_suma + rt("fpa" + LTrim(Str(k)))
'     End If
'  Next
'  rt("kpe") = PEL.Recordset("kod")
'  rt("aji") = Round(SYN_ME_FPA, f_psifiaAjias)  ' Val(SynoloG.Caption)
'  If F_XONDR = 3 Then
'      rt("ELGA") = Round(SYN_ME_FPA * 3 / 100, f_psifiaAjias)
'      rt("aji") = rt("aji") - rt("ELGA")
'      f_suma = rt("AJI")
'      SynoloG.Caption = f_suma
'  Else
'      rt("ELGA") = 0
'  End If
'  rt("eidos") = PEL.Recordset("eidos")
'  If Combo4.Text = Combo4.List(0) Then ' METRHTOIS
'     rt("trp") = "ΜΕ"
'  Else
'     rt("trp") = "ΠΙ"
'  End If
'  rt("EIDPAR") = Left(Combo8.Text, 3)
'  rt("B_N1") = Val(Left(APOTH1.Text, 2)) ' αποθηκη
'  rt("B_N2") = Val(Left(APOTH2.Text, 2)) ' αποθηκη
'  rt("parat") = parat.Text
'  rt("metaf") = Val(Replace(metaf.Text, ",", "."))
'  rt("SKOPOS") = SKOPOS.Text
'  rt("PROOR") = PROORISMOS.Text
'  rt("FORTOSH") = FORTOSH.Text
'  rt("AYTOK") = AYTOKINHTO.Text
'  rt.Update
Dim SQLTIM As String

SQLTIM = "Insert INTO TIM (HME,ATIM,AJ1,AJ2,AJ3,AJ4,AJ5,AJ6,AJ7,AJ8,AJ9,"
SQLTIM = SQLTIM + "FPA1,FPA2,FPA3,FPA4,FPA6,FPA7,FPA8,FPA9,"
SQLTIM = SQLTIM + "KPE,AJI,ELGA,EIDOS,TRP,EIDPAR,B_N1,B_N2,"
SQLTIM = SQLTIM + "PARAT,METAF,SKOPOS,PROOR,FORTOSH,AYTOK) VALUES ("
SQLTIM = SQLTIM + "'" + Format(f_dat, "MM/DD/YYYY") + "',"  ' HME
SQLTIM = SQLTIM + "'" + f_matim + "'," 'ATIM
f_suma = 0
For K = 1 To 9
   SQLTIM = SQLTIM + Str(Round(SumFpa(K), f_psifiaAjias)) + "," 'AJ1,aj2..aj9
   f_suma = f_suma + Round(SumFpa(K), f_psifiaAjias)
Next
For K = 1 To 9
   If K <> 5 Then
     SQLTIM = SQLTIM + Str(Round(SumFpa(K) * g_Fpa(K) / 100, f_psifiaAjias)) + "," 'fpa1,fpa2,fpa3..fpa9
     f_suma = f_suma + Round(SumFpa(K) * g_Fpa(K) / 100, f_psifiaAjias)
   End If
Next
SQLTIM = SQLTIM + "'" + PEL.Recordset("kod") + "'," 'KPE

Dim F_ELGA As Single
F_ELGA = 0
If F_XONDR = 3 Then
    F_ELGA = Round(SYN_ME_FPA * 3 / 100, f_psifiaAjias)
End If
SQLTIM = SQLTIM + Str(Round(SYN_ME_FPA - F_ELGA, f_psifiaAjias)) + "," 'AJI
SQLTIM = SQLTIM + Str(Round(F_ELGA, f_psifiaAjias)) + "," 'ELGA
SQLTIM = SQLTIM + "'" + PEL.Recordset("EIDOS") + "'," 'EIDOS
If Combo4.Text = Combo4.List(0) Then ' METRHTOIS
   SQLTIM = SQLTIM + "'ΜΕ',"  'ΤRP
Else
   SQLTIM = SQLTIM + "'ΠΙ',"  'ΤRP
End If
SQLTIM = SQLTIM + "'" + Left(Combo8.Text, 3) + "'," 'ΕΙDPAR
SQLTIM = SQLTIM + Left(APOTH1.Text, 2) + "," ' B_N1
SQLTIM = SQLTIM + Left(APOTH2.Text, 2) + "," ' B_N2
SQLTIM = SQLTIM + "'" + parat.Text + "'," ' PARAT
SQLTIM = SQLTIM + Str(Val(Replace(metaf.Text, ",", "."))) + "," ' METAF
SQLTIM = SQLTIM + "'" + SKOPOS.Text + "'," ' SKOPOS
SQLTIM = SQLTIM + "'" + PROORISMOS.Text + "'," ' PROORISMOS
SQLTIM = SQLTIM + "'" + FORTOSH.Text + "'," ' FORTOSH
SQLTIM = SQLTIM + "'" + AYTOKINHTO.Text + "')" ' AYTOKINHTO

Dim NTIM As Long

Gdb.Execute SQLTIM, NTIM




rt.Close
re.Close

For K = 1 To f_SeiresTimologioy
   gm_str(K) = ""
Next



Exit Sub

'----------------------------------------------------------------------------
ADD_RECORD:
20      re.AddNew
        re("prood_aj") = K
        re("hme") = f_dat
        re("atim") = f_matim
        ' re("mona") = n
        re("ekpt") = Val(Grid1.TextMatrix(K, f_e))
        re("timm") = Val(Grid1.TextMatrix(K, F_T))
         If F_EXO_XROMATA = 0 Then
            re("poso") = Val(Grid1.TextMatrix(K, f_p))
            re("kau_ajia") = Round(Val(Grid1.TextMatrix(K, f_a)), f_psifiaAjias)
            re("MIK_AJIA") = Round(Val(Grid1.TextMatrix(K, f_a)), f_psifiaAjias) * (1 + g_Fpa(Val(Grid1.TextMatrix(K, f_f))) / 100)
         Else
         
         
        If XROMATA(K, 0, 0) = -0.01 Then    ' ΣΗΜΑΔΙ ΟΤΙ ΧΡΗΣΙΜΟΠΟΙΗΣΕ ΤΟ F6
               re("poso") = XROMATA(K, K2, k3)
        Else
               re("poso") = Val(Grid1.TextMatrix(K, f_p))
        End If
         
         
         
         
         
            re("PROELEYSH") = Left(XROMATA(K, K2, 0), 3) + Left(XROMATA(K, 0, k3), 2) + "*" + Format(K2, "00") + "-" + Format(k3, "00")
            
            re("kau_ajia") = Round(re("POSO") * re("TIMM") * (100 - re("EKPT")) / 100, f_psifiaAjias)
            re("MIK_AJIA") = Round(re("POSO") * re("TIMM") * (100 - re("EKPT")) / 100, f_psifiaAjias) * (1 + g_Fpa(Val(Grid1.TextMatrix(K, f_f))) / 100)
         End If
        
         
        ' ΕΝΗΜΕΡΏΝΕΙ ΤΟ ΚΈΡΔΟς
        ' meid.Open "SELECT * FROM EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) + "'", Gdb, adOpenDynamic, adLockOptimistic 'adOpenForwardOnly, adLockReadOnly
         If PEL.Recordset("EIDOS") = "e" And Val(Grid1.TextMatrix(K, F_T)) > 0 Then
            re("KERDOS") = Val(Grid1.TextMatrix(K, F_T)) - meid("XTI")
         End If
         re("kode") = Grid1.TextMatrix(K, f_k)
         re("onoma") = Left(Grid1.TextMatrix(K, f_o), 35)
        ' ΒΑΖΕΙ ΤΗΝ ΣΥΣΚΕΥΑΣΙΑ ΜΕΣΑ ΣΤΟ ΟΝΟΜΑ ΓΙΑ ΤΙΜΟΛΟΓΗΣΗ
        If f_use_Sysk = 1 Then
          If meid("SYSKMAX") > 1 Then
            MHKOS = Len(Trim(Format(meid("SYSKMAX"), "####") + "X" + Format(re("poso") / meid("SYSKMAX"), "####")))
            re("onoma") = Left(Grid1.TextMatrix(K, f_o) + Space(35), 34 - MHKOS) + " " + Format(meid("SYSKMAX"), "####") + "X" + Format(re("poso") / meid("SYSKMAX"), "####")
          End If
        End If
        
        re("mona") = Grid1.TextMatrix(K, F_M)  ' Right(Grid1.TextMatrix(K, f_o), 3)
        
        
        SYN_ME_FPA = SYN_ME_FPA + re("MIK_AJIA")
        re("fpa") = Val(Grid1.TextMatrix(K, f_f))
        FEGG(re("FPA")) = FEGG(re("FPA")) + re("MIK_AJIA")
        
        re("pelkod") = PEL.Recordset("kod")
        re("eidos") = PEL.Recordset("eidos")
               
        re("APOT") = Val(Left(APOTH1.Text, 2))
         re("XRE") = 0: re("PIS") = 0 'ΑΠΟΦΥΓΗ NULL
         
        Select Case Val(F_POS_APOU)
        Case 1 ' +EISAGOGES
           re("XRE") = re("POSO")
        Case 2 ' -EISAGOGES
           re("XRE") = -re("POSO")
        Case 3 ' +EXAGOGES
           re("PIS") = re("POSO")
        Case 4 ' +EISAGOGES
           re("PIS") = -re("POSO")
        End Select
        re.Update
        
        
      If F_STADIO = 5 Then ' ενδοδιακινηση  Left(f_matim, 1) = "σ" Then
30      re.AddNew
        re("hme") = f_dat
        re("atim") = f_matim
        
         
         If F_EXO_XROMATA = 0 Then
            re("poso") = Val(Grid1.TextMatrix(K, f_p))
         Else
            re("poso") = XROMATA(K, K2, k3)
            re("PROELEYSH") = Left(XROMATA(K, K2, 0), 3) + Left(XROMATA(K, 0, k3), 2) + "*" + Format(K2, "00") + "-" + Format(k3, "00")
         End If
         
'         If F_EXO_XROMATA = 0 Then
 '           re("poso") = Val(Grid1.TextMatrix(K, f_p))
  '       Else
   '         re("poso") = XROMATA(K, K2, k3)
    '        re("PROELEYSH") = Format(K2, "00") + "-" + Format(k3, "00")
     '    End If
        
        
'        re("poso") = Val(Grid1.TextMatrix(k, f_p))
        
        
        re("timm") = Val(Grid1.TextMatrix(K, F_T))
        re("kode") = Grid1.TextMatrix(K, f_k)
        re("onoma") = Left(Grid1.TextMatrix(K, f_o), 30)
        re("mona") = Right(Grid1.TextMatrix(K, f_o), 3)
        re("kau_ajia") = Val(Grid1.TextMatrix(K, f_a))
        re("fpa") = Val(Grid1.TextMatrix(K, f_f))
        re("ekpt") = Val(Grid1.TextMatrix(K, f_e))
        re("pelkod") = PEL.Recordset("kod")
        re("eidos") = PEL.Recordset("eidos")
        
        re("APOT") = Val(Left(APOTH2.Text, 2))
        re("XRE") = 0: re("PIS") = 0 'ΑΠΟΦΥΓΗ NULL
        re("XRE") = re("POSO")
        re.Update
      End If
   
   
 
        
        
        
        
        
        
        
        
        
        
        
        
        

Return


MHNYMA:
HandleError "Par1:enhm_eggtim"
Resume Next
            



End Sub

'ενημέρωση αποθήκης
Sub enhm_apot()

Dim G, m, GA, MA, P
'Dim DB As New ADODB.Connection
'Dim e As New ADODB.Recordset
'Dim R As New ADODB.Recordset, K
Dim db As Database, R As New ADODB.Recordset, K, L As Integer
Dim r2 As New ADODB.Recordset
Dim e As New ADODB.Recordset
Dim NeaTim
Dim gm_f_gfkod As String
Erase f_gfkod
Erase f_gfaj


On Error GoTo MHNYMA

Dim m2, gm_f_gfFPA
Dim m_m As Boolean

Dim A_fpa As Single, a As Single 'αξιες γεφυρώσεων
'
'If gConnect = "Access" Then
'   Set db = OpenDatabase(gDir, False, False)
'Else
'   Set db = OpenDatabase(gDir, False, False, gConnect)
'End If

If Len(gdirlog) > 2 Then
   'r2.Open "SELECT * FROM GEFYRES", Gdb
   r2.Open "SELECT * FROM GEFYRES", Gdb, adOpenKeyset, adLockOptimistic
   r2.MoveFirst
   r2.Move F_rec_p - 1
End If




G = "G" + Format(Month(f_dat), "00")
m = "M" + Format(Month(f_dat), "00")
GA = "GA" + Format(Month(f_dat), "00")
MA = "MA" + Format(Month(f_dat), "00")



Set fSCR = CreateObject("MSScriptControl.ScriptControl")
fSCR.language = "vbscript"
fSCR.AddObject "eid", EID
fSCR.AddObject "mactext", mactext











'R.MoveFirst
'e.Index = "EIDko"
'  DB.Open gDir
For K = 1 To Grid1.ROWS - 1 'Do While Not R.EOF
  
  'e.Seek "=", R("KODE")


  
'    e.Close
If Len(Grid1.TextMatrix(K, f_k)) > 0 Then
    'Set gEid = db.OpenRecordset("SELECT *FROM EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) + "'") ' , DB, adOpenKeyset, adLockOptimistic
  EID.RecordSource = "SELECT *FROM EID WHERE KOD='" + Grid1.TextMatrix(K, f_k) _
      + "'" ' , DB, adOpenKeyset, adLockOptimistic
  EID.Refresh
  
  If EID.Recordset.RecordCount = 1 Then
    P = Val(Grid1.TextMatrix(K, f_p)) '("POSO")
    NeaTim = Val(Grid1.TextMatrix(K, F_T)) * (100 - Val(Grid1.TextMatrix(K, _
        f_e))) / 100
    a = P * NeaTim
    'EID.Recordset.Edit
    
   
If F_POS_APOU = "1" Or F_AJIA_APOU = "1" Then ' EISAGOGES=1
         EID.Recordset("POS") = EID.Recordset("POS") + P
         EID.Recordset(G) = EID.Recordset(G) + P
         ' EID.Recordset() = EID.Recordset(G) + P
         
         If F_AJIA_APOU = "1" Then
            EID.Recordset("prom") = Left(PEL.Recordset("EPO"), 30)
         End If
    
         If IsNull(EID.Recordset("POS_KERD")) Then EID.Recordset("POS_KERD") = 0
' 455




'------------------------- ΚΟΣΤΟΛΟΓΗΣΗ ------------------------------------
'--------------------------------------------------------------------------
If kostologhsh And F_STADIO = 4 Then  ' ' 1=ΠΡΟΣΦ 2=ΠΑΡΑΓΓ 3=ΔΑ 4=ΤΙΜ.ΠΟΛ 5=ΕΝΔΟΔΙΑΚΙΝΙΣΙ
         PAR11.Label10 = EID.Recordset("XTI")
         PAR11.label11 = NeaTim
         If Val(Synolo.Caption) <> 0 Then
            PAR11.Label4 = 100 * Val(metaf) / Val(Synolo)
         End If
         PAR11.POSO.Caption = P
         PAR11.Text1(2) = Round(EID.Recordset("POS_KERD"), 3)
         
         If IsNull(EID.Recordset("LTI5")) Then
             EID.Recordset("LTI5") = 0
         End If
          PAR11.Text1(4) = EID.Recordset("LTI5")
          PAR11.Text1(7) = EID.Recordset("LTI")
        If IsNull(EID.Recordset("pos_kerd2")) Then
           EID.Recordset("pos_kerd2") = 0
        End If
           
           PAR11.Text1(6) = EID.Recordset("pos_kerd2")
         
         If IsNull(EID.Recordset("lti2")) Then
           EID.Recordset("lti2") = 0
        End If
     '   PAR11.Text1(7) = EID.Recordset("LTI2")
         
         
         'PAR11.Text1(2).SetFocus
         PAR11.Label1(1).Caption = EID.Recordset("kod") + "-" + _
             EID.Recordset("ono")
         PAR11.SHOW 1
    
        
    If kostologhsh Then  ' μην τυχον και ακυρωσε την κοστολογηση εν τω μεταξύ
    
       If Val(Synolo.Caption) <> 0 Then
           EID.Recordset("POS_KERD3") = 100 * Val(metaf) / Val(Synolo)
       End If

    
    
    
         EID.Recordset("XTI") = NeaTim

'         GoTo 455
         '  EID.Recordset("POS_KERD") = PAR11.Text1(2)
         
         EID.Recordset("LTI5") = Val(Replace(PAR11.Text1(4), ",", "."))  ' PAR11.Text1(4)
         
         EID.Recordset("LTI") = Val(Replace(PAR11.Text1(7), ",", "."))
         
         EID.Recordset("POS_KERD") = (100 * EID.Recordset("LTI5")) / ((100 + _
             g_Fpa(EID.Recordset("FPA"))) / 100 * EID.Recordset("XTI")) - 100
    
         EID.Recordset("POS_KERD2") = (100 * EID.Recordset("LTI")) / _
         EID.Recordset("XTI") - 100 '  EID.Recordset("XTI") - 100
      End If
      
End If
'------------------------- ΚΟΣΤΟΛΟΓΗΣΗ ------------------------------------
'--------------------------------------------------------------------------
End If

    
    
    
 Select Case F_POS_APOU
    Case "2" ' -EISAGOGES
         EID.Recordset("POS") = EID.Recordset("POS") - P
         EID.Recordset(G) = EID.Recordset(G) - P
    Case "3" ' EXAGOGES=1
         EID.Recordset("POS") = EID.Recordset("POS") - P
         EID.Recordset(m) = EID.Recordset(m) + P
    Case "4" ' -EXAGOGES
         EID.Recordset("POS") = EID.Recordset("POS") + P
         EID.Recordset(m) = EID.Recordset(m) - P
    End Select
    
    Select Case F_AJIA_APOU
    Case "1" ' EISAGOGES=1
        ' eid.recordset(GA) = eid.recordset(GA) + a
    Case "2" ' -EISAGOGES
'         eid.recordset(GA) = eid.recordset(GA) - a
    Case "3" ' EXAGOGES=1
'         eid.recordset(MA) = eid.recordset(MA) + a
    Case "4" ' -EXAGOGES
       '  eid.recordset(MA) = eid.recordset(MA) - a
    End Select
    
    
    
    
    EID.Recordset.Update
     ' e.Close
  
  
If Len(gdirlog) > 2 Then
   
   ' μαζεύω τις σούμες για τις γεφυρώσεις
   '====================================================
    If f_pol = "1" Then ' οταν πρόκειται για πωλήσεις
      
      If IsNull(EID.Recordset("kodsynod")) Then
         MsgBox "Δεν έχω κωδικό λογιστικής για τον κωδικό " + _
             EID.Recordset("kod")
         f_error_gefyr = 2
         GoTo 1220
         
      Else
         gm_f_gfkod = EID.Recordset("kodsynod")
      End If
    Else ' agores
       gm_f_gfkod = EID.Recordset("kodlogag")
    End If
    
    ' left(eid.recordset("kodsynod')+"018"
    gm_f_gfkod = IIf(IsNull(r2("kau" + Grid1.TextMatrix(1, f_f))), "", r2("kau" _
        + Grid1.TextMatrix(1, f_f)))
    
    gvar = mac(gm_f_gfkod)
    gm_f_gfkod = gvar ' 7000018
           
    ' "5400"+left(eid.recordset("kodynod"),2)+"04"
    gm_f_gfFPA = IIf(IsNull(r2("kFPA" + Grid1.TextMatrix(1, f_f))), "", _
        r2("kfpa" + Grid1.TextMatrix(1, f_f)))
    gvar = mac(gm_f_gfFPA)
    
    
    gm_f_gfFPA = gvar ' 54007004
    A_fpa = a * g_Fpa(Val(Grid1.TextMatrix(1, f_f))) / 100 ' 500 * 18 /100
    
    
    
    ' τακτοποιώ τις αξίες και τα ΦΠΑ στους πίνακες f_gfkod()  kai f_gfaj()
    ' Π.Χ.  F_GFKOD(1)="7000018"   F_gfaj(1)=100
    '       F_GFKOD(2)="54007004"   F_gfaj(2)=18
    ' τακτοποιώ την καθαρή αξία
    For L = 1 To 20
     If Len(f_gfkod(L)) = 0 Then 'βρήκα άδειο , σημαίνει ότι δεν βρήκα στα γεμάτα κάποιο να ταιριάζει
        f_gfkod(L) = gm_f_gfkod 'den ypirxe to anoigei tora (70-00-00-00)
        f_gfaj(L) = a
        Exit For
     End If
     If f_gfkod(L) = gm_f_gfkod Then 'βρήκα όμοιο στα γεμάτα
        f_gfaj(L) = f_gfaj(L) + a ' soymarei axies
        Exit For
     End If
    Next
   
    ' τακτοποιώ το ΦΠΑ
    For L = 1 To 20
     If Len(f_gfkod(L)) = 0 Then
        f_gfkod(L) = gm_f_gfFPA 'den ypirxe to anoigei tora (54-00-70-04)
        f_gfaj(L) = A_fpa
        Exit For
     End If
     If f_gfkod(L) = gm_f_gfFPA Then
        f_gfaj(L) = f_gfaj(L) + A_fpa ' soymarei axies
        Exit For
     End If
    
    
1220
    
    
    Next
   
   
   
   
   
   
   
   
   
   
   
   
  End If 'If Len(gdirlog) Then
End If  ' If eid.recordset.RecordCount = 1 Then
  
  
  EID.Recordset.Close
End If 'len( Grid1.TextMatrix(K, f_k) )>0
Next

Set fSCR = Nothing

Exit Sub

MHNYMA:
HandleError "Par1:enhm_apot"
Resume Next
            

End Sub
'ενημέρωση πελάτη
Sub enhm_pel()
'----------------------------------------------------------------------
' f_pel = "1"  XREOSI
' f_pel = "2"  -XREOSI
' f_pel = "3"  PISTOSI
' f_pel = "4"  -PISTOSI

Dim X As String

'Dim DB As New ADODB.Connection
'Dim E As New ADODB.Recordset
'Dim R As New ADODB.Recordset, K
'DB.Open gDir

Dim db As Database, e As Recordset, R As New ADODB.Recordset, K

'If gConnect = "Access" Then
'   Set db = OpenDatabase(gDir, False, False)
'Else
''   Set db = OpenDatabase(gDir, False, False, gConnect)
'End If

On Error GoTo MHNYMA

R.Open "SELECT TOP 1 * FROM EGG", Gdb, adOpenDynamic, adLockOptimistic
'If F_XONDR = 3 Then  'ELGA
'Else
   f_suma = Round(Val(SynoloG.Caption), 2)
'End If

''PEL.Recordset.Edit
'If IsNull(PEL.Recordset!ayp) Then PEL.Recordset!ayp = 0
'If IsNull(PEL.Recordset!typ) Then PEL.Recordset!typ = 0
'' F_PEL=1 ΧΡΕΩΣΗ  2= -ΧΡΕΩΣΗ     3=ΠΙΣΤΩΣΗ   4=-ΠΙΣΤΩΣΗ
'If PEL.Recordset("eidos") = "e" Then
'   If f_pel = "1" Or f_pel = "4" Then
'      PEL.Recordset("typ") = PEL.Recordset("typ") + f_suma
'   End If
'   If f_pel = "2" Or f_pel = "3" Then
'      PEL.Recordset("typ") = PEL.Recordset("typ") - f_suma
'   End If
'End If
'' ΠΡΟΜΗΘΕΥΤΕΣ   F_PEL=3 ΧΡΕΩΣΗ  4= -ΧΡΕΩΣΗ     1=ΠΙΣΤΩΣΗ   2=-ΠΙΣΤΩΣΗ
'If PEL.Recordset("eidos") = "r" Then
'   If f_pel = "1" Or f_pel = "4" Then
'      PEL.Recordset("typ") = PEL.Recordset("typ") + f_suma
'   End If
'   If f_pel = "2" Or f_pel = "3" Then
'      PEL.Recordset("typ") = PEL.Recordset("typ") - f_suma
'   End If
'End If
'PEL.Recordset.Update


R.AddNew
   R("hme") = CDate(Format(DTPicker1.value, "DD/MM/YYYY")) ' DTPicker1.Value
   R("UserID") = gUserId
   
   
  ' R("atim") = f_matim
   
   R("apa") = Val(mID$(f_matim, 2, 6))
   
   
   R("atim") = f_matim ' Left(f_matim, 1 + f_psifia_atim)
   
   
   R("ait") = Left(F_PARAS, 20)



      R("XREOSI") = 0: R("PISTOSI") = 0 ' ΑΠΟΦΥΓΗ ΝULL
      
      If f_pel = "1" Or f_pel = "4" Then ' XREVSH
         R("xre") = f_suma
         If f_pel = "1" Then R("XREOSI") = f_suma
         If f_pel = "4" Then R("PISTOSI") = -f_suma
      Else
         R("xre") = -f_suma
         If f_pel = "2" Then R("XREOSI") = -f_suma
         If f_pel = "3" Then R("PISTOSI") = f_suma
      End If
      
   
   
   R("eidos") = PEL.Recordset("eidos")
   R("kod") = PEL.Recordset("kod")

On Error GoTo MHNYMA





   R("id") = 0
R.Update

On Error Resume Next



R.Close


R.Open "SELECT SUM(XREOSI) AS SX,SUM(PISTOSI) AS SP FROM EGG" _
& " WHERE EIDOS='" + PEL.Recordset("EIDOS") + "' AND KOD='" + PEL.Recordset("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly
If IsNull(PEL.Recordset!ayp) Then
    Gdb.Execute "UPDATE PEL SET AYP=0,TYP=" + Str(R("SX")) + " - " + Str(R("SP")) + "" _
    & " WHERE EIDOS='" + PEL.Recordset("EIDOS") + "' AND KOD='" + PEL.Recordset("KOD") + "'"
Else
    Gdb.Execute "UPDATE PEL SET TYP=AYP + " + Str(R("SX")) + " - " + Str(R("SP")) + "" _
    & " WHERE EIDOS='" + PEL.Recordset("EIDOS") + "' AND KOD='" + PEL.Recordset("KOD") + "'"
End If
' PEL.Refresh










Exit Sub

MHNYMA:
HandleError "Par1:enhm_pel"
Resume Next
            



End Sub


Sub enhm_met()
'
'ενημέρωση μετρητών
'ΚΑΙ ΤΑΜΕΙΟ ΚΑΤΑΣΤΗΜΑΤΟΣ ------------------------------------------------
'----------------------------------------------------------------------
' Dim DB As New ADODB.Connection
' Dim E As New ADODB.Recordset
' Dim R As New ADODB.Recordset, K
Dim db As Database, e As New ADODB.Recordset, R As New ADODB.Recordset, K
Dim m_suma As Single, r2 As New ADODB.Recordset
'If gConnect = "Access" Then
'   Set db = OpenDatabase(gDir, False, False)
'Else
'   Set db = OpenDatabase(gDir, False, False, gConnect)
'End If

On Error GoTo MHNYMA
Dim m_tameio

'DB.Open gDir
R.Open "SELECT TOP 1 * FROM EGG", Gdb, adOpenDynamic, adLockOptimistic

m_suma = Round(Val(SynoloG.Caption), 2)

' PEL.Recordset.Edit


'3-5-2009  REM TIS EPOMENES 20 GRAMMES GIATI DEN XREIAZONTAI
'' F_PEL=1 ΧΡΕΩΣΗ  2= -ΧΡΕΩΣΗ     3=ΠΙΣΤΩΣΗ   4=-ΠΙΣΤΩΣΗ
'If PEL.Recordset("eidos") = "e" Then
'   If f_pel = "1" Or f_pel = "4" Then
'      PEL.Recordset("typ") = PEL.Recordset("typ") - m_suma
'   End If
'   If f_pel = "2" Or f_pel = "3" Then
'      PEL.Recordset("typ") = PEL.Recordset("typ") + m_suma
'   End If
'End If
'
'
'' ΠΡΟΜΗΘΕΥΤΕΣ   F_PEL=3 ΧΡΕΩΣΗ  4= -ΧΡΕΩΣΗ     1=ΠΙΣΤΩΣΗ   2=-ΠΙΣΤΩΣΗ
'If PEL.Recordset("eidos") = "r" Then
'   If f_pel = "1" Then
'      PEL.Recordset("typ") = PEL.Recordset("typ") - m_suma
'   End If
'   If f_pel = "2" Then
'      PEL.Recordset("typ") = PEL.Recordset("typ") + m_suma
'   End If
'End If
'PEL.Recordset.Update
'

R.AddNew
   R("hme") = DTPicker1.value
   
   R("atim") = f_matim
   R("apa") = Val(mID$(f_matim, 2, 6))
   
   R("ait") = "ΜΕΤΡΗΤΑ"

'   If PEL.Recordset("eidos") = "e" Then
'      If f_pel = "1" Or f_pel = "4" Then ' XREVSH
'         R("xre") = Round(m_suma, 2)
'
'      Else  ' 2 , 3
'         R("xre") = Round(-m_suma, 2)
'
'      End If
'   Else ' ΠΡΟΜΗΘΕΥΤΕΣ ' ΠΡΟΜΗΘΕΥΤΕΣ   F_PEL=3 ΧΡΕΩΣΗ  4= -ΧΡΕΩΣΗ     1=ΠΙΣΤΩΣΗ   2=-ΠΙΣΤΩΣΗ
'      If f_pel = "2" Then
'         R("xre") = Round(-m_suma, 2)
'         R("XPI") = "Χ"
'      Else ' 2 , 3
'         R("xre") = Round(m_suma, 2)
'         R("XPI") = "Χ"
'      End If
'   End If
       
       R("XREOSI") = 0: R("PISTOSI") = 0 ' ΑΠΟΦΥΓΗ ΝULL
       ' F_PEL=1 ΧΡΕΩΣΗ  2= -ΧΡΕΩΣΗ     3=ΠΙΣΤΩΣΗ   4=-ΠΙΣΤΩΣΗ
      
   ' If PEL.Recordset("eidos") = "e" Then
       If f_pel = "1" Then  '  TIMOLOGIO , LIANIKH
             R("PISTOSI") = Abs(f_suma)
       ElseIf f_pel = "2" Then  ' PISTOTIKO TIMOLOGIO
             R("PISTOSI") = -Abs(f_suma)
       ElseIf f_pel = "3" Then  ' tim.agoras
             R("XREOSI") = Abs(f_suma)
       ElseIf f_pel = "3" Then  ' PISTOTIKO TIMOLOGIO apo promitheyti
             R("XREOSI") = -Abs(f_suma)
       End If
'    ElseIf PEL.Recordset("eidos") = "r" Then
'       If f_pel = "1" Then  '  TIMOLOGIO , LIANIKH
'             R("PISTOSI") = Abs(f_suma)
'       ElseIf f_pel = "2" Then  ' PISTOTIKO TIMOLOGIO
'             R("PISTOSI") = -Abs(f_suma)
'       ElseIf f_pel = "3" Then  ' PISTOTIKO TIMOLOGIO
'             R("XREOSI") = Abs(f_suma)
'       ElseIf f_pel = "3" Then  ' PISTOTIKO TIMOLOGIO
'             R("XREOSI") = -Abs(f_suma)
'       End If
'    End If
    
   R("eidos") = PEL.Recordset("eidos")
   R("kod") = PEL.Recordset("kod")
R.Update



'------------------ ΤΑΜΕΙΟ ΚΑΤΑΣΤΗΜΑΤΟΣ ------------------------------------------------
If Val(Left(Combo4.Text, 2)) > 0 Then
   m_tameio = F_LOGPLHR(Val(Left(Combo4.Text, 2)) - 1)
End If


' = FindParametroi("PELAT43", "F_TAMEIO", "9990", "ΛΟΓΑΡΙΑΣΜΟΣ ΤΑΜΕΙΟ")


r2.Open "SELECT * FROM PEL WHERE KOD='" + m_tameio + "' AND EIDOS='e'", Gdb, adOpenForwardOnly, adLockReadOnly
If r2.EOF Then
  Gdb.Execute "insert into PEL (EIDOS,KOD,EPO) VALUES ('e','" + m_tameio + "','" + Combo4.Text + "')"
End If
r2.Close
R.AddNew
   R("hme") = DTPicker1.value
   R("atim") = f_matim
   R("apa") = Val(mID$(f_matim, 2, 6))
   R("ait") = "ΜΕΤΡΗΤΑ"
   
' F_PEL=1 ΧΡΕΩΣΗ  2= -ΧΡΕΩΣΗ     3=ΠΙΣΤΩΣΗ   4=-ΠΙΣΤΩΣΗ
' GIA TAMEIO SHMAINEI
'  XREOSI=F_SUMA , XREOSI=-F_SUMA , PISTOSI=F_SUMA , PISTOSI=-F_SUMA

   
      R("XREOSI") = 0: R("PISTOSI") = 0 ' ΑΠΟΦΥΓΗ ΝULL
      If PEL.Recordset("eidos") = "e" Then ' εισπραξη
            If f_pel = "1" Then R("XREOSI") = f_suma '
            If f_pel = "2" Then R("XREOSI") = -f_suma
            If f_pel = "3" Then R("PISTOSI") = f_suma
            If f_pel = "4" Then R("XREOSI") = -f_suma
      Else
          R("PISTOSI") = f_suma  'πληρωμή
      End If
   R("eidos") = "e" ' PEL.Recordset("eidos")
   R("kod") = m_tameio ' PEL.Recordset("kod")
R.Update
'------------------ ΤΑΜΕΙΟ ΚΑΤΑΣΤΗΜΑΤΟΣ ------------------------------------------------














R.Close


Exit Sub

MHNYMA:
HandleError "Par1:enhm_met"
Resume Next
            




End Sub

'εκτύπωση
Function print_timol(ByVal Atim As String, ByVal hme As Date)
Dim DUM

    Me.Caption = "ektyp_forma"
    
If f_FORM_EKTY = 99 Then
    DUM = ektyp_forma("c:\mercvb\f90.txt", Atim, hme)
Else
    DUM = ektyp_forma("c:\mercvb\f" + Format(f_FORM_EKTY, "00") + ".txt", Atim, hme)
End If

'
'GoTo 10

End Function


Private Sub b_Katax_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     KeyAscii = 0
  End If
 
End Sub

Private Sub Combo8_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
     Text2(0).Enabled = True: Text2(0).SetFocus
 End If
End Sub

Private Sub Command1_Click()
  
  Dim db As DAO.Database
  Dim R As DAO.Recordset
  Dim K
  Dim a$, b, X
  Dim d
  Dim rSQL As New ADODB.Recordset
  Dim rSQL2 As New ADODB.Recordset
  
  
  
 If Len(Dir("c:\TIMOLOGIA", vbDirectory)) > 2 Then
   ' LOISIR
   loisir_import
    Exit Sub
 End If
  
 If Len(Dir("C:\SYMBOL", vbNormal)) > 2 Then
   ' LOISIR
    SYMBOL_import
    Exit Sub
 End If
  
  
  
  
  
  
  
  ' LAMBANEI ENA MONO PARASTAR\TIKO
If F_falcon_Polla_tim = 2 Then    ' mpoygoydis
  
  
    ' μπουγουδης
   Set db = OpenDatabase("C:\LAGEURO\MHXANAKI", False, False, "dBase III;")
   Set R = db.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM=" + Left(Text1.Text, 5) + ";")
   If R.RecordCount = 0 Then
      MsgBox "δεν υπάρχουν εγγραφές"
      Exit Sub
   End If


   R.MoveFirst
   K = 0
   Do While Not R.EOF
      K = K + 1
      If K > f_SeiresTimologioy Then Exit Do
      If IsNull(R("BARCODE")) Then
         Grid1.TextMatrix(K, f_k) = " "
      Else
         Grid1.TextMatrix(K, f_k) = R("BARCODE")
      End If
      R.MoveNext
   Loop
   R.Close
   db.Close
   Exit Sub
End If
 
  
  
  
  
  
  
  
  ' LAMBANEI ENA MONO PARASTAR\TIKO
  If F_falcon_Polla_tim = 0 Or Len(Dir("c:\DIONISIS.EXE", vbDirectory)) > 2 Then
    
   '
   On Error Resume Next
   Set db = OpenDatabase("C:\LAGEURO\MHXANAKI", False, False, "dBase III;")
   Set R = db.OpenRecordset("SELECT * FROM TIMAGOR where POS<>0") '  WHERE ATIM=" + left(Text1.Text, 5))
   If R.RecordCount = 0 Then
      MsgBox "δεν υπάρχουν εγγραφές"
      Exit Sub
   End If


   R.MoveFirst
   K = 0
   Dim M_CODE
   Do While Not R.EOF
      K = K + 1
      If K > f_SeiresTimologioy Then Exit Do
      If IsNull(R("BARCODE")) Or R("pos") = 0 Then
         Grid1.TextMatrix(K, f_k) = " "
      Else
      
      
         rSQL.Open "SELECT * FROM BARCODES WHERE left(ERG,13)='" + (Left(R("BARCODE"), 13)) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
        If rSQL.EOF And Left(R("BARCODE"), 1) = "0" Then
           rSQL.Close
           
           rSQL.Open "SELECT * FROM BARCODES WHERE ERG LIKE '" + Trim(mID(R("BARCODE"), 2, 12)) + "%'", Gdb, adOpenForwardOnly, adLockReadOnly
        End If
        
        If rSQL.EOF Then
            M_CODE = " "
            Grid1.TextMatrix(K, f_k) = M_CODE
               '       OK = False
            '       List2.AddItem rDBF("BARCODE") + ": ΔΕΝ ΥΠΑΡΧΕΙ ΣΤΟ ΑΡΧΕΙΟ ΕΙΔΩΝ"
        Else
            M_CODE = rSQL("KOD")
            Grid1.TextMatrix(K, f_k) = M_CODE
          '   Grid1.TextMatrix(K, F_T) = rSQL("LTI")
            Grid1.TextMatrix(K, f_p) = R("POS")
            rSQL.Close
            rSQL.Open "SELECT * FROM EID WHERE KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
            Grid1.TextMatrix(K, F_T) = Round(rSQL("LTI"), 2)
            Grid1.TextMatrix(K, f_o) = rSQL("ONO")
            Grid1.TextMatrix(K, f_f) = rSQL("FPA")
            
            If f_TIMOKAT > 0 And Not IsNull(PEL.Recordset("pek")) Then
               rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + Str(PEL.Recordset("pek")) + " AND KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
               If rSQL2.EOF Then
                  Grid1.TextMatrix(K, f_e) = 0
               Else
                  Grid1.TextMatrix(K, f_e) = rSQL2(0)
               End If
               rSQL2.Close
               
               'ΔΟΥΛΕΥΕΙ ΜΕ ΚΑΡΦΩΤΗ ΤΙΜΗ
             ElseIf f_TIMOKAT < 0 And Not IsNull(PEL.Recordset("pek")) Then
               rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + Str(PEL.Recordset("pek")) + " AND KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
               If rSQL2.EOF Then
                  ' Grid1.TextMatrix(k, F_T) = 0
               Else
                  Grid1.TextMatrix(K, F_T) = rSQL2(0)
               End If
               rSQL2.Close
               
               
               
               
               
               
               
               
            Else
              Grid1.TextMatrix(K, f_e) = PEL.Recordset("pek")
              ' Grid1.Text =
            End If
            
            Grid1.row = K
            FindSynolo
                     
        End If
        rSQL.Close
        
      End If
      R.Edit
        R("pos") = 0
      R.Update
      
      
      
      R.MoveNext
   Loop
   R.Close
   db.Close
    
   Exit Sub
  
  
End If
   

   '
   On Error Resume Next
   Set db = OpenDatabase("C:\LAGEURO\MHXANAKI", False, False, "dBase III;")
   Set R = db.OpenRecordset("SELECT * FROM TIMAGOR where ATIM=" + Left(Text1.Text, 5))
   If R.RecordCount = 0 Then
      MsgBox "δεν υπάρχουν εγγραφές"
      Exit Sub
   End If


   R.MoveFirst
   K = 0
   'Dim M_CODE
   Do While Not R.EOF
      K = K + 1
      If K > f_SeiresTimologioy Then Exit Do
      If IsNull(R("BARCODE")) Or R("pos") = 0 Then
         Grid1.TextMatrix(K, f_k) = " "
      Else
      
      
        rSQL.Open "SELECT * FROM BARCODES WHERE LEFT(ERG,13)='" + Left(R("BARCODE"), 13) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
        
       If rSQL.EOF And Left(R("BARCODE"), 2) = "00" Then
         rSQL.Close
           
           rSQL.Open "SELECT * FROM BARCODES WHERE ERG LIKE '" + mID(R("BARCODE"), 2, 12) + "%'", Gdb, adOpenForwardOnly, adLockReadOnly
        End If
        
        
        
        
        If rSQL.EOF Then
            M_CODE = " "
            Grid1.TextMatrix(K, f_k) = M_CODE
        Else
            M_CODE = rSQL("KOD")
            Grid1.TextMatrix(K, f_k) = Trim(M_CODE)
            Grid1.TextMatrix(K, f_p) = R("POS")
            rSQL.Close
            
            rSQL.Open "SELECT * FROM EID WHERE KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
            Grid1.TextMatrix(K, F_T) = Round(rSQL("LTI"), 2)
            Grid1.TextMatrix(K, f_o) = Trim(rSQL("ONO"))
            Grid1.TextMatrix(K, f_f) = rSQL("FPA")
            
            If f_TIMOKAT > 0 And Not IsNull(PEL.Recordset("pek")) Then
               rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + Str(PEL.Recordset("pek")) + " AND KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
               If rSQL2.EOF Then
                  Grid1.TextMatrix(K, f_e) = 0
                  
                     'ΔΟΥΛΕΥΕΙ ΜΕ ΚΑΡΦΩΤΗ ΤΙΜΗ
             ElseIf f_TIMOKAT < 0 And Not IsNull(PEL.Recordset("pek")) Then
               rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + Str(PEL.Recordset("pek")) + " AND KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
               If rSQL2.EOF Then
                  ' Grid1.TextMatrix(k, F_T) = 0
               Else
                  Grid1.TextMatrix(K, F_T) = rSQL2(0)
               End If
               rSQL2.Close
                 
                  
                  
               Else
                  Grid1.TextMatrix(K, f_e) = rSQL2(0)
               End If
               rSQL2.Close
            Else
              Grid1.TextMatrix(K, f_e) = PEL.Recordset("pek")
              ' Grid1.Text = PEL.Recordset("pek")
            End If
            
            Grid1.row = K
            FindSynolo
                     
        End If
        rSQL.Close
        
      End If
     ' R.Edit
      '  R("pos") = 0
     ' R.Update
      
      
      
      R.MoveNext
   Loop
   R.Close
   db.Close





End Sub
Sub loisir2_import()

  Dim db As DAO.Database
  Dim R As DAO.Recordset
  Dim K
  Dim a$, b, X
  Dim d
  Dim rSQL As New ADODB.Recordset
  
  
  
  
 If Len(Dir("c:\TIMOLOGIA", vbDirectory)) > 2 Then
   ' LOISIR



Open "c:\timologia\" + Text1.Text + ".asc" For Binary Access Read As #1

'Do While Not EOF(1)
' Line Input #1, A$
' Exit Do
'Loop
'Close #1
'Exit Sub

X = 1


Dim kod, per, MON, pos, TIMH
Dim f As String, RECS As Integer

K = 0

Do While Not EOF(1) 'True
  On Error Resume Next
    Line Input #1, a$
   If EOF(1) Then Exit Do
    
   f = Chr$(10) + Chr$(27) + Chr$(15)
   
'   B = InStr(x, A$, f)
'   If B = 0 Then
 '    GoTo 555
'     Exit Do
  ' End If
   
   
      
      
      
X = 1
   
     f = Chr$(27) + Chr$(15) 'Chr$(10) + Chr$(27) + Chr$(15)
   
   b = InStr(X, a$, f)
 
   If b = 0 Then
     GoTo 555
   End If
   
   
   
   
   
   
   
   
   
   d = mID(a$, b + 2, 180)
   kod = Left(d, 12)
   
   
   
   
   
   
   
   per = mID$(d, 30, 38)
   pos = mID$(d, 76, 10)
   
   MON = mID$(d, 90, 10)
   TIMH = Trim(mID$(d, 104, 10))
   TIMH = Replace(TIMH, ",", ".")
   
   'List1.AddItem d
   
   

rSQL.Open "SELECT * FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
   If rSQL.EOF Then
 '      On Error Resume Next
      Gdb.Execute "INSERT INTO EID (KOD,ONO,FPA,XTI,MON) VALUES " _
      & "('" + Trim(kod) + "','" + Replace(Left(Trim(per), 35), "'", "~") + "',2," + TIMH + ",'" + Trim(MON) + "')", RECS
      If RECS = 0 Then
         MsgBox "Δεν υπάρχει το είδος " + kod
      Else
         MsgBox "Ανοίχθηκε το είδος " + kod
      End If
   End If
   
rSQL.Close
   
   X = b + 3


  K = K + 1
  If K > f_SeiresTimologioy Then Exit Do
     Grid1.TextMatrix(K, f_k) = kod
     Grid1.TextMatrix(K, f_o) = per
     Grid1.TextMatrix(K, f_p) = pos
     Grid1.TextMatrix(K, F_T) = TIMH
     Grid1.TextMatrix(K, f_f) = 2
     Grid1.row = K
     FindSynolo
     
555
Loop
Close #1

  
 Else
 ' μπουγουδης
Set db = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
Set R = db.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM=" + Left(Text1.Text, 5))
If R.RecordCount = 0 Then
   MsgBox "δεν υπάρχουν εγγραφές"
   Exit Sub
End If


R.MoveFirst
K = 0
Do While Not R.EOF
  K = K + 1
  If K > f_SeiresTimologioy Then Exit Do
  If IsNull(R("BARCODE")) Then
     Grid1.TextMatrix(K, f_k) = " "
  Else
     Grid1.TextMatrix(K, f_k) = R("BARCODE")
  End If
  R.MoveNext
Loop
R.Close
db.Close

End If



End Sub

Sub loisir_import()

  Dim db As DAO.Database
  Dim R As DAO.Recordset
  Dim K
  Dim a$, b, X
  Dim d
  Dim rSQL As New ADODB.Recordset
  
  
  
  
 If Len(Dir("c:\TIMOLOGIA", vbDirectory)) > 2 Then
   ' LOISIR



Open "c:\timologia\" + Text1.Text + ".asc" For Binary Access Read As #1

'Do While Not EOF(1)
' Line Input #1, A$
' Exit Do
'Loop
'Close #1
'Exit Sub

X = 1


Dim kod, per, MON, pos, TIMH
Dim f As String, RECS As Integer

K = 0

Do While Not EOF(1) 'True
  On Error Resume Next
    Line Input #1, a$
   If EOF(1) Then Exit Do
    
   f = Chr$(10) + Chr$(27) + Chr$(15)
   
'   B = InStr(x, A$, f)
'   If B = 0 Then
 '    GoTo 555
'     Exit Do
  ' End If
   
   
      
      
      
X = 1
   
    ' f = Chr$(27) + Chr$(15) 'Chr$(10) + Chr$(27) + Chr$(15)
   f = Chr$(15)  'Chr$(10) + Chr$(27) + Chr$(15)
   
   b = InStr(X, a$, f)
 
   If b = 0 Then
     GoTo 555
   End If
   
   
   
   
   
   
   
   
   
   d = mID(a$, b + 2, 180)
   kod = mID(d, 3, 11) ' Left(d, 12)
   
   
   
   
   
   
   
   per = mID$(d, 30, 38)
   pos = mID$(d, 74, 10)
   
   MON = mID$(d, 85, 10)
   TIMH = Trim(mID$(d, 92, 10))
   TIMH = Replace(TIMH, ",", ".")
   
   'List1.AddItem d
   
   

rSQL.Open "SELECT * FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
   If rSQL.EOF Then
 '      On Error Resume Next
      Gdb.Execute "INSERT INTO EID (KOD,ONO,FPA,XTI,MON) VALUES " _
      & "('" + Trim(kod) + "','" + Replace(Left(Trim(per), 35), "'", "~") + "',2," + TIMH + ",'" + Trim(MON) + "')", RECS
      If RECS = 0 Then
         MsgBox "Δεν υπάρχει το είδος " + kod
      Else
         MsgBox "Ανοίχθηκε το είδος " + kod
      End If
   End If
   
rSQL.Close
   
   X = b + 3


  K = K + 1
  If K > f_SeiresTimologioy Then Exit Do
     Grid1.TextMatrix(K, f_k) = kod
     Grid1.TextMatrix(K, f_o) = per
     Grid1.TextMatrix(K, f_p) = pos
     Grid1.TextMatrix(K, F_T) = TIMH
     Grid1.TextMatrix(K, f_f) = 2
     Grid1.row = K
     FindSynolo
     
555
Loop
Close #1

  
 Else
 ' μπουγουδης
Set db = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
Set R = db.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM=" + Left(Text1.Text, 5))
If R.RecordCount = 0 Then
   MsgBox "δεν υπάρχουν εγγραφές"
   Exit Sub
End If


R.MoveFirst
K = 0
Do While Not R.EOF
  K = K + 1
  If K > f_SeiresTimologioy Then Exit Do
  If IsNull(R("BARCODE")) Then
     Grid1.TextMatrix(K, f_k) = " "
  Else
     Grid1.TextMatrix(K, f_k) = R("BARCODE")
  End If
  R.MoveNext
Loop
R.Close
db.Close

End If







End Sub

Sub SYMBOL_import()

  Dim db As DAO.Database
  Dim R As DAO.Recordset
  Dim K
  Dim a$, b, X
  Dim d
  Dim FPA As Integer
  
  Dim rSQL As New ADODB.Recordset
Open "C:\APOG.TXT" For Input As #1

X = 1


Dim kod, per, MON, pos, TIMH
Dim f As String, RECS As Integer

K = 0

Do While Not EOF(1) 'True
  On Error Resume Next
    Line Input #1, a$

   X = 1
   kod = mID(a$, 1, 13)
   
   pos = Val(mID$(a$, 15, 20))
   
   rSQL.Open "SELECT KOD FROM BARCODES WHERE ERG='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
   If rSQL.EOF Then
        MsgBox "Δεν υπάρχει το είδος " + kod
   Else
        kod = rSQL(0)
        rSQL.Close
        
        rSQL.Open "SELECT ONO,LTI5,FPA FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
        per = rSQL("ONO")
        TIMH = rSQL("LTI5")
        FPA = rSQL("FPA")
        TIMH = Round(rSQL("LTI5") / ((100 + g_Fpa(FPA)) / 100), 2)
        
        
        
   End If
   
   rSQL.Close
   
    K = K + 1
  If K > f_SeiresTimologioy Then Exit Do
     Grid1.TextMatrix(K, f_k) = kod
     Grid1.TextMatrix(K, f_o) = per
     Grid1.TextMatrix(K, f_p) = pos
     Grid1.TextMatrix(K, F_T) = TIMH
     Grid1.TextMatrix(K, f_f) = FPA
     Grid1.row = K
     FindSynolo
     
If EOF(1) Then Exit Do
     
     

Loop
Close #1

  








End Sub



Private Sub Command2_Click()
  Dim a
  
  
  a = toascii("C:\MERCVB\F31.TXT", "C:\LIBRA\HOST.TXT")
  
  MsgBox "Τέλος Αποστολής"
  
On Error Resume Next
a = Shell("C:\LIBRA\LIBRA.EXE 1", vbMaximizedFocus)
  
End Sub

Private Sub Command3_Click()
Dim fpelat4 As New pelat4
'If Index = 0 Then
     fpelat4.SHOW
     fpelat4.Text2(0).Text = Text2(0).Text
   If POLAGOR.ListIndex = 0 Then
      fpelat4.PELPROM.ListIndex = 1
   Else
       fpelat4.PELPROM.ListIndex = 0
     'fpelat4.Option1(1).value = True
   End If
     fpelat4.Data1.ConnectionString = gConnect
     fpelat4.Refresh
     gApoMenu = False
     
     'pelat4.Text2(0).SetFocus
     
     'fpelat4.SHOW
     fpelat4.Text2(0).SetFocus
     SendKeys "+{ENTER}"
Dim DUM
     DUM = fpelat4.KARTELLA(gDir)
'End If
   Set fpelat4 = Nothing


End Sub

Private Sub Command4_Click()
'ΔΗΜΙΟΥΡΓΕΙ ΑΥΤΟΜΑΤΑ ΠΑΡΑΣΤΑΤΙΚΟ ΤΠΥ,ΑΠΥ,ΔΑ
'ΠΑΙΡΝΟΝΤΑΣ ΤΑ ΣΤΟΙΧΕΙΑ ΑΠΟ ΤΟ ΕΙΔΟΣ ΠΟΥ ΕΧΟΥΜΕ ΠΕΡΑΣΕΙ ΜΕΣΑ ΑΠΟ ΤΗΝ ΦΟΡΜΑ ΠΑΡΑΛΑΒΗ


Dim ko As String, ajia As String, n As Integer, TYPOS As Integer




ko = InputBox("δωσε το barcode του είδους ", , "")
ajia = InputBox("δωσε αξία ", , "")
TYPOS = Val(InputBox("Αποδ.ΠΥ=6 Δ.Α=9 ΤΠΥ=16 ", , ""))

If TYPOS = 6 Or TYPOS = 9 Or TYPOS = 16 Then
   ' OK
Else
    MsgBox "Λάθος τύπος παραστατικού"
    Exit Sub
End If



Dim R As New ADODB.Recordset, sql As String

'ΒΡΙΣΚΩ ΤΟ ΕΙΔΟΣ
R.Open "SELECT * FROM EID WHERE LEFT(KOD,6)='" + Left(ko, 6) + "'", Gdb, adOpenDynamic, adLockOptimistic

If R.EOF Then
    MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ Ο ΚΩΔΙΚΟΣ"
    Exit Sub
End If





SendKeys ""

Text2(0).Text = Trim(R("CH6")) ' ΚΩΔΙΚΟΣ ΠΕΛΑΤΗ

If Left(R("CH6"), 4) = "1112" Then
    'ΕΝΗΜΕΡΩΝΩ  ΤΟΝ ΠΕΛΑΤΗ ΛΙΑΝΙΚΗΣ 1112 ΜΕ ΤΑ ΣΤΟΙΧΕΙΑ ΤΟΥ
    If TYPOS = 16 Then ' tpy
        MsgBox "δεν μπορεις να κόψεις ΤΠΥ χωρίς ΑΦΜ"
    Else
        'sql = "UPDATE PEL SET EPO='" + R("CH2") + "',DIE='" + R("CH3") + "',EPA='" + R("CH4") + "',THL='" + R("CH5") + "' WHERE EIDOS='e' AND KOD='" + Trim(R("CH6")) + "'"
        'Gdb.Execute sql, N
    
     '      sql = "UPDATE PEL SET EPO='" + Trim(R("CH2")) + "',DIE='" + Trim(R("CH3")) + "',EPA='" + Trim(R("CH4")) + "',THL='" + Left(Trim(R("CH5")), 10) + "' WHERE EIDOS='e' AND KOD='" + Trim(R("CH6")) + "'"
    '     Gdb.Execute sql, N
    
      sql = "UPDATE PEL SET EPO='" + Trim(R("CH2")) + "',DIE='" + Trim(R("CH3")) + "',EPA='" + Trim(R("CH4")) + "',THL='" + Left(Trim(R("CH5")), 10) + "' WHERE EIDOS='e' AND KOD='" + Trim(R("CH6")) + "'"
         Gdb.Execute sql, n
    
    
    
    
    End If
    
    
End If
 
 
'μπαινω απ'ευθειας στο παραστατικό
  PARAS.Text = PARAS.List(TYPOS - 1)
  
  Dim ll As Integer
  Dim gg As String
  
  gg = IIf(TYPOS = 6, "y", IIf(TYPOS = 9, "A", "Y"))
  
  For ll = 0 To PARAS.ListCount - 1
      If Right(PARAS.List(ll), 1) = gg Then
          PARAS.Text = PARAS.List(ll)
      End If
  Next
  
  
  
  
  PARAS_LostFocus
  SendKeys "{ENTER}{ENTER}"
  SendKeys "{ENTER}{ENTER}"
  
 Dim K As Integer
 
 
 
'ΔΗΜΙΟΥΡΓΩ ΤΙΣ ΣΕΙΡΕΣ ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ
If TYPOS = 9 Then 'δελτιο αποστολής
  K = 1
  Grid1.TextMatrix(K, f_k) = ko
  Grid1.TextMatrix(K, f_p) = 1
  Grid1.TextMatrix(K, f_o) = R("ONO")
  Grid1.TextMatrix(K, F_T) = 0
  Grid1.TextMatrix(K, f_f) = 2
  Grid1.TextMatrix(K, f_a) = 0
  Grid1.row = K: Grid1.Col = 0:  FindSynolo
  SKOPOS.Text = "ΑΠΟ ΕΠΙΣΚΕΥΗ"
  Gdb.Execute "UPDATE EID SET HM3=GETDATE(),NUM1=" + Text1.Text + " WHERE KOD='" + ko + "'"
Else
  K = 1
  Grid1.TextMatrix(K, f_k) = "223344"
  timText3 = "223344"
  Grid1.TextMatrix(K, f_p) = 1
  Grid1.TextMatrix(K, f_o) = "ΕΠΙΣΚΕΥΗ"
  Grid1.TextMatrix(K, F_T) = Round(Val(ajia) / 1.19, 2)
  Grid1.TextMatrix(K, f_f) = 2
  Grid1.TextMatrix(K, f_a) = Round(Val(ajia) / 1.19, 2)
  Grid1.row = K: Grid1.Col = 0:  FindSynolo
  K = 2
  Grid1.TextMatrix(K, f_k) = ko
  Grid1.TextMatrix(K, f_p) = 0
  Grid1.TextMatrix(K, f_o) = R("ONO")
  Grid1.TextMatrix(K, F_T) = 0
  Grid1.TextMatrix(K, f_f) = 2
  Grid1.TextMatrix(K, f_a) = 0
  Grid1.row = K: Grid1.Col = 0:  FindSynolo
  Gdb.Execute "UPDATE EID SET HM3=GETDATE(),NUM2=" + Text1.Text + " WHERE KOD='" + ko + "'"
End If






End Sub

Private Sub Command5_Click()
       CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
       CrystalReport1.Destination = crptToWindow
       CrystalReport1.Action = 1


End Sub

Private Sub Command6_Click()
Dim fff As New pelat2

   ' fff.ffbuff = f_buff + Text2(0).Text ' Text1(0).Text
   If POLAGOR.ListIndex = 0 Then
      fff.PELPROM.ListIndex = 1
       fff.ffbuff = "r" + Text2(0).Text ' Text1(0).Text

   Else
       fff.PELPROM.ListIndex = 0
       fff.ffbuff = "e" + Text2(0).Text ' Text1(0).Text
   End If
   fff.Command5.Enabled = False
   gApoMenu = False
   pelat2.SHOW
    
End Sub

Private Sub diortosis_Click()
 
  Dim FF As PAR2, MKOD As String
  
Dim Atim
  
TDBGrid1.Col = 0
DTPicker1.value = CDate(TDBGrid1.Text)
   
TDBGrid1.Col = 4
MKOD = TDBGrid1.Text
   
   
   
TDBGrid1.Col = 1
Atim = TDBGrid1.Text


Dim R As New ADODB.Recordset

  R.Open "SELECT *FROM PARASTAT WHERE EIDOS='" + Left(Atim, 1) + "'", Gdb, adOpenDynamic, adLockOptimistic
  If R.EOF Then Exit Sub
  If R("pol") > 3 Or R("pol") < 1 Then Exit Sub ' δεν διορθωνω  asxeta

  
  If R("pol") = 1 Then 'pvlhseis
     POLAGOR.ListIndex = 1
  ElseIf R("pol") = 2 Then
     POLAGOR.ListIndex = 0 ' agores
  ElseIf R("pol") = 3 Then
     POLAGOR.ListIndex = 2 ' κινησεις αποθηκης
  End If
  
  
  
'  PARAS.Text = R("titlos") + Space(1) + Left(Atim, 1)
 POLAGOR_LostFocus
'POLAGOR.SetFocus
'PARAS.SetFocus
Dim K As Integer
For K = 0 To PARAS.ListCount - 1
   PARAS.ListIndex = K
   If Right(PARAS.Text, 1) = Left(Atim, 1) Then
      Exit For
   End If
Next

If K = PARAS.ListCount Then Exit Sub




  
  
  
  ' βρισκω τις παραμετρους του παραστατικού
  find_eid_parastat

  Text1.Text = Val(mID(Atim, 2, 6))

  Text2(0).Text = MKOD

Text2(0).Enabled = True
Text2(0).SetFocus
  
SendKeys "{ENTER} {ENTER}"


End Sub

Private Sub DTPicker1_GotFocus()
   DTPicker1.ToolTipText = "."
End Sub

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
       KeyCode = 0 'suppress the beep
       
       If DTPicker1.ToolTipText = "." Then
          DTPicker1.ToolTipText = ""
       Else
          keybd_event VK_TAB, 0, 0, 0 'send a tab
       End If
    End If
End Sub

Private Sub Form_DblClick()
   PARAMETROI.PARAM.Caption = "PAR1"
    
    PARAMETROI.SHOW 1
   

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then b_akyr_Click
  
  
' If KeyCode = 38 Then 'PANO BELOS KANEI SHIFT+TAB
'    SendKeys "+{TAB}"
' End If
  
  
  
  
  

End Sub

Private Sub Grid1_DblClick()
  timText3.Visible = True
' kartella
  apot21.Text1.Text = Grid1.TextMatrix(Grid1.row, 1)
  
  apot21.Text2.Text = Grid1.TextMatrix(Grid1.row, 2)
  apot21.Label1.Caption = Left(APOTH1.Text, 1)
  apot21.Label2.Caption = "0"
  
  
  
  
  'Me.Hide
  
  apot21.SHOW
 





End Sub

Private Sub Grid1_Scroll()
   timText3.Visible = False
End Sub

Private Sub GridEidon_GotFocus()
  ' ARXIKO Grid1.width = 5775
  GridEidon.width = 10320

End Sub

Private Sub GridEidon_LostFocus()
   GridEidon.width = 5775
End Sub

Private Sub isotimia_LostFocus()
   
   Dim K As Integer
   Dim mISOT
   
   mISOT = gVal(isotimia.Text)
   
   If mISOT <> 1 Then
      
      isotimia.Text = mISOT
      
      For K = 1 To Grid1.ROWS - 1
         If Val(Grid1.TextMatrix(K, f_p)) <> 0 Then
            Grid1.TextMatrix(K, F_T) = Val(Grid1.TextMatrix(K, F_T)) / mISOT
            Grid1.row = K: FindSynolo
         End If
      Next
   End If


End Sub

Private Sub Label18_Click()
 UPDATE_PINAKES 12, Label18, Combo4
End Sub

Private Sub Label19_Click()
 UPDATE_PINAKES 7, Label9, SKOPOS
End Sub

Private Sub Label20_Click()
    UPDATE_PINAKES 8, Label20, FORTOSH
End Sub

Private Sub Label21_Click()
   UPDATE_PINAKES 9, Label21, PROORISMOS
End Sub

Private Sub Label5_Click()
 UPDATE_PINAKES 10, Label5, AYTOKINHTO
End Sub

Private Sub Label6_Click()
    UPDATE_PINAKES 5, Label6, Combo8
End Sub

Private Sub Label7_Click()
 UPDATE_PINAKES 4, Label7, APOTH1
End Sub









Private Sub par2show_Click()
    par2showing
End Sub

Private Sub POLAGOR_Change()
  PARAS.Enabled = True
  
End Sub

Private Sub POLAGOR_Click()
   PARAS.Enabled = True
End Sub

Private Sub POLAGOR_GotFocus()
   POLAGOR.BackColor = vbYellow
     F_POLAGOR_LastValue = POLAGOR.Text
     'Me.KeyPreview = False
   
End Sub

Private Sub POLAGOR_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   PARAS.Enabled = True
   PARAS.SetFocus
 End If
 
End Sub

Private Sub POLAGOR_LostFocus()
' ANALOGA ME TI DIALEJA ΑΝΟΙΓΕΙ ΤΟ PARASTAT KAI GEMIZEI TO PARAS


Dim X As String, K As Integer
'Dim db As Database
 Dim R As New ADODB.Recordset
 
 
' db.Open gDir


'If gConnect = "Access" Then
   'Set db = OpenDatabase(gDir, False, False)
'Else
   'Set db = OpenDatabase(gDir, False, False, gConnect)
'End If
POLAGOR.BackColor = vbWhite
If POLAGOR.Text = POLAGOR.List(0) Then
    X = "2": fbuff = "r" 'προμηθευτές
    
    f_SeiresTimologioy = f_MaxSeiresTimologioy 'POSA EIDH MPORO NA PERASO Max TIMOLOGIA Agoras
     'f_SeiresTimologioy = 30 OTAN DIALEGO  POLISEIS
    
    Combo4.Text = Combo4.List(1) ' PISTVSH
  ElseIf POLAGOR.Text = POLAGOR.List(1) Then
    X = "1": fbuff = "e" 'pelates
    f_SeiresTimologioy = f_MinSeiresTimologioy 'POSA EIDH MPORO NA PERASO Max TIMOLOGIA Agoras  '25
  Else
    X = "3": fbuff = "l" 'λογαριασμοί
End If

 If Val(X) = 0 Then X = "1"
 
'Set R = OpenRecordset("select pol,titlos,eidos from parastat where pol like '" + x + "%';")


' θα κανειrefresh τα παραστατικά μόνο άν άλλαξε
' ή βρισκόμαστε στην πρώτη φορά

If F_POLAGOR_LastValue <> POLAGOR.Text Or Len(PARAS.Text) = 0 Then ' MONO THN 1H FORA NA ALLAZEI
' If Left(APOTH1.Text) = "1" Then
    R.Open "select POL,TITLOS,EIDOS from PARASTAT where POL='" + X + "'  AND YP" + Left(APOTH1.Text, 1) + "=1 order by TITLOS", Gdb, adOpenDynamic, adLockOptimistic

PARAS.Clear
 If R.RecordCount = 0 Then
    Exit Sub
 End If
 
 R.MoveFirst
  K = 0
  Do While Not R.EOF
     If Not IsNull(R("TITLOS")) And R("pol") = X Then
       If F_STOPDELTIA = 99 Then
          If R("eidos") = "t" Or R("eidos") = "l" Then
             PARAS.AddItem R("TITLOS") + Space(1) + R("eidos")
             K = K + 1
          End If
       Else
          PARAS.AddItem R("TITLOS") + Space(1) + R("eidos")
             K = K + 1
      End If
        mEidos_Par(K) = R("eidos")
     End If
     R.MoveNext
  Loop


 
   If X = "2" Then 'agores
     PARAS.Text = PARAS.List(f_1choice_agores)
   ElseIf X = "1" Then 'polhseis
     PARAS.Text = PARAS.List(f_1choice_polhseis)
   Else
     PARAS.Text = PARAS.List(1)
   End If
   
   
   
End If

DTPicker1.value = Now


PARAS.Enabled = True

F_POLAGOR_LastValue = POLAGOR.Text


' Me.KeyPreview = True

 X = mEidos_Par(1 + PARAS.ListIndex)
 Text1.Text = find_parastat(1, X, 0)


'db.Close

End Sub

Private Sub PARAS_GotFocus()
   PARAS.BackColor = vbYellow
End Sub

Private Sub PARAS_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Text1.SetFocus   ' SEIRA.SetFocus
End Sub

Private Sub PARAS_LostFocus()
' STELNEI STO FIND_PARASTAT NA BREI TON A/A KAI PAIRNEI TIS PARAMETROYS APO TO PARASTAT

Dim K As Integer

'Dim db As Database
Dim R As New ADODB.Recordset
 
Dim X As String

On Error GoTo MHNYMA


PARAS.BackColor = vbWhite
'db.Open gDir
'R.Open "select *from parastat;", db, adOpenKeyset, adLockReadOnly ', adCmdTableDirect

'If gConnect = "Access" Then
'   Set db = OpenDatabase(gDir, False, False)
'Else
'   Set db = OpenDatabase(gDir, False, False, gConnect)
'End If
R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic



  R.MoveFirst
  Do While Not R.EOF
     If Right(PARAS.Text, 1) = R("EIDOS") Then
       F_TITLOS = R("TITLOS")
       F_POS_APOU = R("POS_APOU")
       F_AJIA_APOU = R("AJIA_APOU")
       f_pel = R("pel")
       f_pol = R("pol") ' 1=POLISEIS 2=AGORES
       F_STADIO = R("STADIO") ' 1=ΠΡΟΣΦ 2=ΠΑΡΑΓΓ 3=ΔΑ 4=ΤΙΜ.ΠΟΛ 5=ΕΝΔΟΔΙΑΚΙΝΙΣΙ
       If IsNull(R("GEF_P")) Then
           F_rec_p = 0
       Else
           F_rec_p = R("GEF_P") ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
       End If
       
       If IsNull(R("XONDR")) Then
           F_XONDR = 0
       Else
           F_XONDR = R("XONDR") ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
       End If
       
       
       
       
       
'       F_rec_m = R("GEF_M") ' ........  GIA METRITA


       If IsNull(R("GEF_M")) Then
           F_rec_m = 0
       Else
           F_rec_m = R("GEF_M") ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
       End If
       





       f_metasx = R("metasx") ' απο ποιό παραστατικό μετασχηματίζεται
     End If
     R.MoveNext
  Loop

  If F_STADIO = 5 Then  '' 1=ΠΡΟΣΦ 2=ΠΑΡΑΓΓ 3=ΔΑ 4=ΤΙΜ.ΠΟΛ 5=ΕΝΔΟΔΙΑΚΙΝΙΣΙ
     APOTH2.Enabled = True
  Else
     APOTH2.Enabled = False
  End If

'PARAS.Clear
geidos_paras = PARAS.ListIndex


If F_STADIO = 2 Then
   Label14.Caption = "MarkUp %"
Else
   Label14.Caption = "Μεταφορικά και άλλα έξοδα"
End If




 X = mEidos_Par(1 + PARAS.ListIndex)
 Text1.Text = find_parastat(K, X, 0)


  If Val(Text1.Text) > 0 Then
      selectPEL.SetFocus
  End If






PARAS.Enabled = False



Exit Sub

MHNYMA:
HandleError "Par1:paras_lostfocus"
Resume Next
            



End Sub

Private Sub SEIRA_GotFocus()
   SEIRA.BackColor = vbYellow
End Sub

Private Sub SEIRA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Text1.SetFocus
End Sub

Private Sub SEIRA_LostFocus()
' STELNEI STO FIND_PARASTAT NA BREI TON A/A KAI PAIRNEI TIS PARAMETROYS APO TO PARASTAT

Dim K As Integer

Dim X As String

 ' Dim db As Database
 Dim R As New ADODB.Recordset
 On Error GoTo MHNYMA
 

SEIRA.BackColor = vbWhite



For K = 0 To SEIRA.ListCount
   If SEIRA.Text = SEIRA.List(K) Then
      Exit For
   End If
Next


If K <= SEIRA.ListCount Then
  ' Text1.Text = find_parastat(k, Right(PARAS.Text, 1), 0)
    X = mEidos_Par(1 + PARAS.ListIndex)
    Text1.Text = find_parastat(K, X, 0)
    
End If



'If gConnect = "Access" Then
'   Set db = OpenDatabase(gDir, False, False)
'Else
'   Set db = OpenDatabase(gDir, False, False, gConnect)
'End If

' db.Open gDir
' R.Open "SELECT *FROM parastat", db, adOpenKeyset, adLockReadOnly ', adCmdTableDirect


R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic


  R.MoveFirst
  Do While Not R.EOF
     If Right(PARAS.Text, 1) = R("EIDOS") Then
       F_TITLOS = R("TITLOS")
       F_POS_APOU = R("POS_APOU")
       F_AJIA_APOU = R("AJIA_APOU")
     End If
     R.MoveNext
  Loop
  
  R.Close
  
  If Val(Text1.Text) > 0 Then
      selectPEL.SetFocus
  End If
  
  
  
  
  
  
  
Exit Sub

MHNYMA:
HandleError "Par1:FINDSYNOLO"
Resume Next
              
  
  
'PARAS.Clear
End Sub




Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   b_Katax.SetFocus
End If

End Sub


Private Sub DBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    If f_ColUpdate Then
      f_ColUpdate = False
      GridEidon.SetFocus
    End If
 End If
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If LastCol = 5 Then
       
      ' DBGrid1.Row = DBGrid1.Row + 1: DBGrid1.Col = 0
   End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       KeyAscii = 0 'suppress the beep
       keybd_event VK_TAB, 0, 0, 0 'send a tab
    End If
End Sub

Private Sub SKOPOS_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       KeyAscii = 0 'suppress the beep
       keybd_event VK_TAB, 0, 0, 0 'send a tab
    End If
End Sub

Private Sub FORTOSH_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
       KeyAscii = 0 'suppress the beep
       keybd_event VK_TAB, 0, 0, 0 'send a tab
    End If
End Sub

Private Sub PROORISMOS_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       KeyAscii = 0 'suppress the beep
       keybd_event VK_TAB, 0, 0, 0 'send a tab
    End If
End Sub



Private Sub GridEidon_Click()
Dim DELT As Recordset


'Grid1.SetFocus

End Sub



Private Sub GridEidon_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Dim m_a As String
   Dim K As Integer
   Dim MKOD(30) As String
   Dim mPoso(30) As Single
   Dim f5 As String
   Dim elem(30) As String
   Dim DUM As Integer
   Dim sPOSO As Single
   Dim arSET As Single
   Dim mRow
   Dim R As New ADODB.Recordset
   Dim Sql2 As New ADODB.Recordset
   On Error GoTo MHNYMA
   
   For K = 1 To 30: MKOD(K) = "": mPoso(K) = 0: Next
   
   
   'ΕΑΝ Ο ΚΩΔΙΚΟΣ ΕΧΕΙ ΜΑΖΙ ΤΟΥ ΟΜΑΔΑ ΕΙΔΩΝ ΤΟΤΕ
   'ΦΟΡΤΩΝΕ ΣΕ ΠΙΝΑΚΑ ΤΑ ΕΙΔΗ
   
  
  If EID.Recordset.EOF Then Exit Sub
  '------------------------------------------------------------------------------------------
  If IsNull(EID.Recordset("MMEMO")) Then f5 = "" Else f5 = EID.Recordset("MMEMO")
   If Left(f5, 1) = "#" Then
       'Dim r As New ADODB.Recordset
        R.Open "SELECT MEMO FROM EID WHERE KOD='" + EID.Recordset("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic
        f5 = R(0)
        R.Close
       DUM = FETES_DELIM(f5, elem)
       
       ' BΡΙΣΚΩ ΠΟΙΟΥΣ ΚΩΔΙΚΟΥΣ MKOD() ΚΑΙ ΤΙ ΠΟΣΟΤΗΤΕΣ ΕΧΩ MPOSO()
       sPOSO = 0
       For K = 1 To 30
          If Left(elem(K), 1) = "#" Then
            DUM = InStr(elem(K), ",")
            If DUM - 2 > 0 Then 'AN EXEI KOMMA
               MKOD(K) = mID$(elem(K), 2, DUM - 2)
               mPoso(K) = Val(mID$(elem(K), DUM + 1, Len(elem(K)) - DUM))
               sPOSO = sPOSO + mPoso(K)
            End If
          End If
       Next
       arSET = Val(InputBox("Δώστε συνολ.τεμάχια" + Chr(13) + "To σετ αυτό έχει " + Format(sPOSO, "##") + " τεμάχια" + Chr(13) + "αν θέλετε 2 σετ τότε " + Chr(13) + " πληκτρολογήστε " + Format(2 * sPOSO, "##"), "Τεμάχια ανά σετ : " + Format(sPOSO, "##")))
       If sPOSO > 0 Then arSET = arSET / sPOSO
       
   'sxediazo to grid
   DUM = Grid1.row
   For K = 1 To 30
     If mPoso(K) > 0 Then
         EID.RecordSource = "select KOD,ONO,XTI,LTI,LTI5,FPA,MON,MEMO from EID WHERE KOD ='" + MKOD(K) + "';"
         EID.Refresh
         If Not EID.Recordset.EOF Then
            Grid1.TextMatrix(DUM + K - 1, f_k) = MKOD(K)
            Grid1.TextMatrix(DUM + K - 1, f_p) = mPoso(K) * arSET
            If IsNull(EID.Recordset("mon")) Then m_a = Space(3) Else m_a = EID.Recordset("mon")
            Grid1.TextMatrix(DUM + K - 1, f_o) = Left(EID.Recordset("ono") + Space(60), 60)
            Grid1.TextMatrix(DUM + K - 1, F_M) = m_a
            If F_XONDR = 3 Then 'elga
               Grid1.TextMatrix(DUM + K - 1, f_f) = 5
            Else
               Grid1.TextMatrix(DUM + K - 1, f_f) = EID.Recordset("FPA")
            End If
            Grid1.TextMatrix(DUM + K - 1, f_e) = PEL.Recordset("ekpt")
          If fbuff = "r" Then
              Grid1.TextMatrix(DUM + K - 1, F_T) = EID.Recordset("xti")
          Else
              If lianikh Then
                 Grid1.TextMatrix(DUM + K - 1, F_T) = EID.Recordset("LTI5") / ((100 + g_Fpa(EID.Recordset("FPA"))) / 100)
                 Grid1.TextMatrix(DUM + K - 1, F_mefpa) = EID.Recordset("LTI5")
              Else
                 Grid1.TextMatrix(DUM + K - 1, F_T) = EID.Recordset("LTI")
              End If
          End If
            
            
            Grid1.row = DUM + K - 1
            FindSynolo
         End If
     End If
   Next
  '------------------------------------------------------------------------------------------
Else
   
   
   
   
   
   
   
   Grid1.Col = f_k
   
  ' ΓΡΑΦΩ ΤΟΝ ΚΩΔΙΚΟ
   Grid1.Text = Left(EID.Recordset("kod"), 16)
   Grid1.Col = f_o
   
   If IsNull(EID.Recordset("mon")) Then m_a = Space(3) Else m_a = Left(EID.Recordset("mon"), 3)
   'ΓΡΑΦΩ ΤΗΝ ΠΕΡΙΓΡΑΦΗ ΤΟΥ ΕΙΔΟΥΣ
   
   Grid1.Col = F_M
   Grid1.Text = m_a
   Grid1.Col = f_o
   
   Grid1.Text = Left(EID.Recordset("ono") + Space(60), 60)  '+ m_a
   
   
   
   
   If Left(EID.Recordset("kod"), 1) = "*" Then
     Grid1.Text = InputBox("Δώσε περιγραφή", "", Grid1.Text)
   End If
   
   
   
   
   
   
   Grid1.Col = f_p  '11-3-07
   Grid1.Text = 1 '11-3-07
   
   
   
   
   
   
   Grid1.Col = f_e
   


If f_TIMOKAT > 0 And Not IsNull(PEL.Recordset("pek")) Then
   R.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + Str(PEL.Recordset("pek")) + " AND KOD='" + Trim(Left(EID.Recordset("kod"), 13)) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
   If R.EOF Then
      Grid1.Text = ""
   Else
      Grid1.Text = R(0)
   End If
  'ΔΟΥΛΕΥΕΙ ΜΕ ΚΑΡΦΩΤΗ ΤΙΜΗ
 ElseIf f_TIMOKAT < 0 And Not IsNull(PEL.Recordset("pek")) Then
        
        R.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + Str(PEL.Recordset("pek")) + " AND KOD='" + Trim(Left(EID.Recordset("kod"), 13)) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
        If R.EOF Then
           ' Grid1.TextMatrix(k, F_T) = 0
        Else
           Grid1.TextMatrix(K, F_T) = R(0)
        End If
        R.Close
              
   
   
   
Else
   Grid1.Text = PEL.Recordset("pek")
End If
   'ΓΡΑΦΩ ΤΟ ΦΠΑ ΤΟΥ ΕΙΔΟΥΣ
   Grid1.Col = f_f
   If Not IsNull(PEL.Recordset("TYPOS")) Then
    If Val(PEL.Recordset("TYPOS")) = 1 Then  'ΝΗΣΙΑ
      If EID.Recordset("FPA") = 2 Then
         Grid1.Text = 3
     ' ElseIf EID.Recordset("FPA") = 4 Then
      '   Grid1.Text = EID.Recordset("FPA")
      ElseIf EID.Recordset("FPA") = 1 Then
         Grid1.Text = 4
      End If
     Else
        
            If F_XONDR = 3 Then 'elga
               Grid1.Text = 5
            Else
               Grid1.Text = EID.Recordset("FPA")
            End If
                  
        
        
            ' Grid1.Text = EID.Recordset("FPA")
     End If
         
   Else
      Grid1.Text = EID.Recordset("FPA")
   End If
   
   'ΓΡΑΦΩ ΤΗΝ ΤΙΜΗ ΤΟΥ ΕΙΔΟΥΣ
   If fbuff = "r" Then
      Grid1.Col = F_T
      Grid1.Text = Format(EID.Recordset("xti"), "####0." + String(f_psifiaAjias, "0"))
    Else
              Grid1.Col = F_T
              
              If lianikh Then
                 
                 Grid1.Col = F_mefpa
                 Grid1.Text = EID.Recordset("LTI5")
                 Grid1.Col = F_T
                 Grid1.Text = Format(EID.Recordset("LTI5") / (1 + g_Fpa(EID.Recordset("FPA")) / 100), "####0." + String(f_psifiaAjias, "0"))
                 
                 
              Else
                 
                 Grid1.Text = Format(EID.Recordset("LTI"), "####0." + String(f_psifiaAjias, "0"))
              
              End If
              
          If F_STADIO = 2 Then  ' ΠΑΡΑΓΓΕΛΊΑ
            
            If Val(f_PARAGGMARKUP) = 0 Then
            Else
               Grid1.Text = Format((100 + Val(metaf.Text)) / 100 * EID.Recordset("XTI"), "####0.00")
            End If
            
          End If
     
          If F_STADIO = 3 Or F_STADIO = 5 Then ' deltio apostolis   Ή ΕΝΔΟΔΙΑΚΙΝΗΣΗ
               Grid1.Text = 0
          End If
     
     
     
     
     
     
     
     
     End If
 End If  '="#"
   
   ' ΤΟ ΤΕΧΤ3 ΠΑΡΑΛΑΜΒΑΝΕΙ ΤΗΝ ΤΙΜΗ ΑΠΟ ΤΟ GRID
   Grid1.Col = f_p
   timText3.Text = Grid1.Text
   timText3.Visible = False
   
   ' ΤΟ ΤΕΧΤ3 ΠΑΡΑΛΑΜΒΑΝΕΙ ΤIΣ ΔΙΑΣΤΑΣΕΙΣ ΑΠΟ ΤΟ GRID
   timText3.Top = Grid1.Top + Grid1.CellTop
   timText3.Left = Grid1.Left + Grid1.CellLeft
   timText3.width = Grid1.CellWidth - 15
   timText3.height = Grid1.CellHeight - 25
   timText3.Visible = True
   timText3.SetFocus
 End If  ' =CHR(13)

Dim PI


' PI = "C:\MERCVB\IMAGES\" + CNull(EID.Recordset("kod"))

If F_EIKONA > 0 Then
  On Error Resume Next
  If F_EIKONA = 2 Then 'βοηθητικοσ κωδικοσ
     PI = "C:\MERCVB\IMAGES\" + CNull(EID.Recordset("ERG"))
  ElseIf F_EIKONA = 1 Then ' kvdikos
     PI = "C:\MERCVB\IMAGES\" + CNull(EID.Recordset("kod"))
  ElseIf F_EIKONA = 12 Then 'βοηθητικοσ κωδικοσ +.jpg
     PI = "C:\MERCVB\IMAGES\" + CNull(EID.Recordset("ERG")) + ".JPG"
  ElseIf F_EIKONA = 11 Then ' kodikos .jpg
     PI = "C:\MERCVB\IMAGES\" + CNull(EID.Recordset("kod")) + ".JPG"
  End If
  
 
 If Len(Dir(PI)) > 0 Then
    Image1.Picture = LoadPicture(PI)
    '.PaintPicture Picture1.Picture, 0, 0, Picture2.width, Picture2.height, 0, 0, Picture1.width, Picture1.height
 Else
    Image1.Picture = LoadPicture()
    ' Picture2.Picture = LoadPicture()
 End If
   
End If











'If Len(Dir(PI)) > 0 Then
'   On Error Resume Next
'   Image1.Picture = LoadPicture(PI)
'Else
'  On Error Resume Next
'   PI = "C:\MERCVB\IMAGES\" + EID.Recordset("kod") + ".JPG"
'   If Len(Dir(PI)) > 0 Then
'      Image1.Picture = LoadPicture(PI)
'   Else
'      Image1.Picture = LoadPicture()
'   End If
'End If

Exit Sub

MHNYMA:
HandleError "Par1:GridEidon_KeyUp"
Resume Next
            

End Sub
Function FETES_DELIM(LINE, elem)
Dim KL, KE

On Error GoTo MHNYMA
 '  DIABAZO SE PINAKA OLA TA STOIXEIA THS GRAMHS
     For KE = 1 To 30
       elem(KE) = ""
     Next
     
     KL = 0   ' metraei xaraktires
     KE = 0   ' metritis toy pinaka  ELEMENT
     Do
        KE = KE + 1  ' metritis toy pinaka  ELEMENT
        KL = KL + 1  ' metraei xaraktires
        
        Do While mID$(LINE, KL, 1) <> Chr(13) ' tab
            elem(KE) = elem(KE) + mID$(LINE, KL, 1)
            KL = KL + 1  ' metraei xaraktires
            If KL > Len(LINE) Then Exit Do
        Loop
        KL = KL + 1 ' για να περασει το chr(10)
        
            If KL > Len(LINE) Then Exit Do
     
     Loop Until KL >= Len(LINE) 'OLO TO MHKOS THS GRAMMHS

     FETES_DELIM = 0
Exit Function

MHNYMA:
HandleError "Par1:Fetesdelim"
Resume Next
            
End Function
Private Sub GridPelaton_DblClick()
Dim a, e
   
On Error GoTo MHNYMA
10



If PEL.Recordset.EOF Then
   Text2(1).SelStart = 0
   Text2(1).SelLength = Len(Text2(1))
  
   Text2(0).Text = ""
   Text2(1).SetFocus
      
      
      Exit Sub
End If

f_gkod_pel = fbuff + PEL.Recordset("KOD")
   
   'A = pel.Recordset("KOD")
   'E = Data1.Recordset("EIDOS")
   

GridPelaton.Visible = False
20
PEL.RecordSource = "select *from PEL where EIDOS='" + Left(f_gkod_pel, 1) + "' and KOD='" + mID$(f_gkod_pel, 2, Len(f_gkod_pel) - 1) + "'"
PEL.Refresh
30
If PEL.Recordset.RecordCount = 0 Then
  Text2(1).SetFocus
  Exit Sub
End If
40
If Len(gdirlog) > 0 Then
     If IsNull(PEL.Recordset("shm2")) Then
         f_error_gefyr = 1
     Else
        If Len(PEL.Recordset("shm2")) < 4 Then
            f_error_gefyr = 1
        End If
     End If
End If
50
   Text2(0).Text = PEL.Recordset("KOD")
   Text2(1).Text = PEL.Recordset("EPO")
   Text2(2).Text = PEL.Recordset("afm")
   Text2(3).Text = PEL.Recordset("die")
   
   If Not IsNull(PEL.Recordset("ARPARAG")) Then
      Combo4.Text = PEL.Recordset("ARPARAG")
   End If
   
   
   ' If Check1 Then f_metasx = "a"
60

If Len(Trim(f_metasx)) = 0 And ApoParaggelia = Checked Then
      f_metasx = "a"
End If

   FlxGridDeltion.Clear
   FlxGridDeltion.Visible = False
   GridEidon.Visible = True
   


   ' αν είναι τιμολόγιο του δείχνει τα δελτία σε εκκρεμότητα
If Len(Trim(f_metasx)) > 0 Then
'
   


 EID.RecordSource = "select METAF,ATIM,CONVERT(CHAR(10),HME,103) AS [Ημερομηνία],ID_NUM from TIM where left(ATIM,1) in ('" + Get_meKomma(f_metasx) + "') and KPE='" + PEL.Recordset("kod") + "' and ART IS null;"
 EID.Refresh
70
  metasxhmatismos.Visible = True
  FlagGridEidon = False
  If EID.Recordset.RecordCount > 0 Then
80        Set FlxGridDeltion.DataSource = EID
     GridEidon.Visible = False
     FlxGridDeltion.Visible = True
     FlxGridDeltion.SetFocus
  End If
End If
  
   
90
   
   'Grid1.Col = 0
   Grid1.row = 1
  'Grid1.SetFocus
  
  Text2(0).Enabled = False
  Text2(1).Enabled = False
  Text2(2).Enabled = False
  
  Grid1.Col = 1
100
   timText3.Top = Grid1.Top + Grid1.CellTop
   timText3.Left = Grid1.Left + Grid1.CellLeft
   timText3.width = Grid1.CellWidth - 15
   timText3.height = Grid1.CellHeight - 25
   
110
   
'- --------------- ΔΙΟΡΘΩΣΗ ΤΙΜΟΛΟΓΙΟΥ ----------------------------------
Dim re As New ADODB.Recordset, ANS As Integer

f_dat = DTPicker1.value
f_matim = Right(PARAS.Text, 1) + Right("0000000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text
ANS = vbNo
re.Open "select * from TIM where  HME>='" + Format(f_dat, "mm/dd/yyyy") + "' AND HME<='" + Format(DateAdd("D", 1, f_dat), "mm/dd/yyyy") + "'  and ATIM='" + f_matim + "'", Gdb, adOpenDynamic, adLockOptimistic
        If Not re.EOF Then
           ANS = MsgBox("Υπάρχει ήδη  παραστατικό με ίδια ημερομηνία και ίδιο αριθμό. " + Chr(13) + "Να διορθωθεί;", vbYesNo)
        End If
  'ans = vbYes
120






If ANS = vbYes Then

'βαζει τον αρχικό πελάτη
   PEL.RecordSource = "select *from PEL where EIDOS='" + Left(f_gkod_pel, 1) + "' and KOD='" + re("kpe") + "'"
   PEL.Refresh
   Text2(0).Text = PEL.Recordset("KOD")
   Text2(1).Text = PEL.Recordset("EPO")
   Text2(2).Text = PEL.Recordset("afm")
   Text2(3).Text = PEL.Recordset("die")

   f_dior_mAtim = f_matim ' ΤΑ ΚΡΑΤΑΩ ΜΗΝ ΤΥΧΟΝ ΤΑ ΑΛΛΑΞΕΙ ΣΤΗΝ ΦΆΣΗ ΤΗΣ ΔΙΟΡΘΩΣΗΣ
   f_dior_HME = f_dat 'ΓΙΑ ΝΑ ΜΠΟΡΕΣΩ ΝΑ ΣΒΗΣΩ ΤΟ ΑΡΧΙΚΟ ΤΙΜΟΛΟΓΙΟ


''το είχα έτσι αλλά έχανε τα είδη  αν  εβαζα λάθος τις αποθήκες
   ' EID.RecordSource = "select EGGTIM.* from EGGTIM  WHERE HME >= '" + Format(f_dat, "mm/dd/yyyy") + "' AND HME <= '" + Format(DateAdd("D", 1, f_dat), "mm/dd/yyyy") + "' and ATIM='" + f_matim + "' AND APOT=" + Str(Val(APOTH1)) + " order by PROOD_AJ"
   EID.RecordSource = "select EGGTIM.* from EGGTIM  WHERE HME >= '" + Format(f_dat, "mm/dd/yyyy") + "' AND HME <= '" + Format(DateAdd("D", 1, f_dat), "mm/dd/yyyy") + "' and ATIM='" + f_matim + "'  order by PROOD_AJ"
   
   EID.Refresh
   '   On Error Resume Next
   EID.Recordset.MoveFirst
   Dim K: K = 1
130

Dim DUMY
   Do While Not EID.Recordset.EOF
      
      XROMATA(K, 1, 0) = Left(EID.Recordset("PROELEYSH"), 3)
      XROMATA(K, 0, 1) = mID(EID.Recordset("PROELEYSH"), 4, 2)
        
      
      Grid1.TextMatrix(K, f_k) = EID.Recordset("kode")
      Grid1.TextMatrix(K, f_p) = EID.Recordset("POSO")
      Grid1.TextMatrix(K, f_e) = EID.Recordset("ekpt")
      Grid1.TextMatrix(K, f_o) = EID.Recordset("onoma")
      Grid1.TextMatrix(K, F_T) = EID.Recordset("TIMM")
      Grid1.TextMatrix(K, F_M) = EID.Recordset("MONA")
      Grid1.TextMatrix(K, f_f) = EID.Recordset("FPA")
      Grid1.TextMatrix(K, f_a) = Round(EID.Recordset("POSO") * EID.Recordset("TIMM"), 2)
      Grid1.row = K: Grid1.Col = 0:  FindSynolo
      
140
      K = K + 1
      EID.Recordset.MoveNext
   Loop
   Grid1.row = 1
   If K > 1 Then
      timText3.Text = Grid1.TextMatrix(1, f_k)
   End If
   
   F_DIOR = 1
   Grid1.Col = f_k  ' για να μην ζητάει να ανοιξει νέα καρτέλλα
End If
   

   
   
   
   
   
   
150
   

   
   
   
   
   
   
   
   
   
   
   
 timText3.Text = Grid1.Text
 timText3.SetFocus
   
  
  Exit Sub

MHNYMA:
HandleError "Par1:GridPelAton_dblclick"
Resume Next
            
  
  
  
  
  
End Sub

Private Sub GridPelaton_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
       KeyAscii = 0 'suppress the beep
       ' keybd_event VK_TAB, 0, 0, 0 'send a tab
    End If
End Sub

Private Sub GridPelaton_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    PAR1.KeyPreview = False
    
    GridPelaton_DblClick

End If
End Sub

Private Sub metasxhmatismos_Click()
'===================================
'συγκεντρώνω τα δελτία που σημάδεψε ο χρήστης
'===================================
Dim DELT As Recordset, K As Integer, L As Integer
Dim e, S As String, d As String, t As Recordset
Dim X, C, Eidh As Recordset
Dim R As Recordset
On Error GoTo MHNYMA





K = 0
'Set delt = EID.Recordset
10
If EID.Recordset.RecordCount = 0 Then
   Exit Sub
End If
' ΣΥΓΚΕΝΤΡΩΝΩ ΤΑ ΣΗΜΑΔΕΜΕΝΑ ΣΤΟ S
20
S = ""
Dim SX As String
SX = ""
For L = 1 To FlxGridDeltion.ROWS - 1
  If Val(FlxGridDeltion.TextMatrix(L, 1)) = 1 Then
    K = K + 1
    f_deltia(K) = FlxGridDeltion.TextMatrix(L, 2) + FlxGridDeltion.TextMatrix(L, 3)
    f_IDdeltia(K) = FlxGridDeltion.TextMatrix(L, 4)
    S = S + "'" + f_deltia(K) + "'" + ","
    SX = SX + Format(Val(mID(f_deltia(K), 2, 6)), "######") + ","
  End If
Next
S = S + "''"

If f_metasx = "a" Then
   parat.Text = "Αρ.Παραγγ." + SX
Else
   parat.Text = "Σχετ.Δελτ." + SX
End If
fSHMADEMENA = S

If K = 0 Then  ' DEN SHMADECE TIPOTA
   FlxGridDeltion.Visible = False
   GridEidon.Visible = True
   Exit Sub
End If
d = EID.RecordSource
30

' σημάδεψε δελτια // ψάχνω να βρώ τα είδη
If K = 1 Then
 '"order by PROOD_AJ"
   EID.RecordSource = "select POSO as SPOSO,KODE,TIMM from EGGTIM where ATIM+CONVERT(CHAR(10),HME,103)  IN (" + S + ") order by PROOD_AJ;"
   EID.Refresh

Else ' polla deltia
   EID.RecordSource = "select sum(POSO) as SPOSO,KODE,TIMM from EGGTIM where ATIM+CONVERT(CHAR(10),HME,103)  IN (" + S + ") group by KODE,TIMM;"
   EID.Refresh
End If

F_SHMADEMENA_DELTIA = S



If EID.Recordset.RecordCount > 500 Then
   MsgBox "Πάνω από 500 είδη στο τιμολόγιο. Διαλέξτε λιγότερα δελτία"
   EID.RecordSource = d
   EID.Refresh
   Exit Sub
End If


40

'Set fDB = OpenDatabase(gDir, False, False, gConnect)

Dim EID2 As New ADODB.Recordset


'EID2.Open "SELECT * FROM EID", Gdb, adOpenDynamic, adLockOptimistic


'If EID.Recordset.RecordCount = 0 Then
'   EID.RecordSource = "SELECT *FROM EID"
'   EID.Refresh
'   Exit Sub
'End If



'ΒΡΙΣΚΕΙ ΤΑ ΕΙΔΗ ΚΑΙ ΒΑΖΕΙ ΠΕΡΙΓΡΑΦΕΣ ΚΑΙ ΤΑ ΑΠΛΩΝΕΙ ΣΤΟ GRID
EID.Recordset.MoveFirst
K = 1
Do While Not EID.Recordset.EOF
  
  EID2.Open "SELECT *FROM EID WHERE KOD='" + EID.Recordset("kode") + "'", Gdb, adOpenDynamic, adLockOptimistic
  Grid1.TextMatrix(K, f_k) = EID.Recordset("kode")
  Grid1.TextMatrix(K, f_p) = EID.Recordset("SPOSO")
  Grid1.TextMatrix(K, f_o) = EID2("ONO")
  If EID.Recordset("TIMM") = 0 Then
      Grid1.TextMatrix(K, F_T) = EID2("LTI")
  Else
      Grid1.TextMatrix(K, F_T) = EID.Recordset("TIMM") ' ΠΑΡΑΓΓΕΛΙΑ
  End If
  
  Grid1.TextMatrix(K, f_f) = EID2("FPA")
  Grid1.TextMatrix(K, f_a) = Round(EID.Recordset("SPOSO") * Grid1.TextMatrix(K, F_T), 2)
  
  Grid1.row = K: Grid1.Col = 0:  FindSynolo
  
  
  
  EID2.Close
  K = K + 1
  EID.Recordset.MoveNext
Loop

50

metasxhmatismos.Enabled = False
'EID.RecordSource = "SELECT *FROM EID"
'EID.Refresh
FindSynolo
   Grid1.row = 1
  'Grid1.SetFocus
  
  Text2(0).Enabled = False
  Text2(1).Enabled = False
  Text2(2).Enabled = False
  
   Grid1.Col = 1
   timText3.Top = Grid1.Top + Grid1.CellTop
   timText3.Left = Grid1.Left + Grid1.CellLeft
   timText3.width = Grid1.CellWidth - 15
   timText3.height = Grid1.CellHeight - 25
   timText3.Text = Grid1.Text
   timText3.SetFocus
   GridPelaton.Visible = False
   FlxGridDeltion.Visible = False
   GridEidon.Visible = True
   
   
   Exit Sub

MHNYMA:
HandleError "Par1:metasxhmatismos"
Resume Next
            
   
   
   
End Sub

Private Sub selectPEL_Click()
' Διαλέγει πελάτη και δείχνει δελτία σε εκκρεμότητα
'===================================================
 ' Διαλέγει πελάτη και δείχνει δελτία σε εκκρεμότητα
'===================================================
  '   GridPelaton.Visible = True
  
  Text2(0).Enabled = True
  Text2(1).Enabled = True
  Text2(2).Enabled = True
  Text2(0).SetFocus
  SendKeys ""


' αν είναι τιμολόγιο του δείχνει τα δελτία σε εκκρεμότητα
If Len(f_metasx) > 0 Then
'
'  EID.RecordSource = "select metaf,atim,hme from tim where left(atim,1)='" + f_metasx + "' and kpe='" + PEL.Recordset("kod") + "' and art=null;"
 ' EID.Refresh
  
 ' metasxhmatismos.Visible = True
 ' FlagGridEidon = False
End If
  
  
  
  
  
  
End Sub

Private Sub GridEidon_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'==================================
'επιλογή είδους ή δελτίων αποστολής
'==================================
Dim DELT As Recordset

  On Error GoTo MHNYMA

If Grid1.Enabled = False Then ' φάση δελτίων αποστολής
 
 If FlagGridEidon Then
   ' FlagGridEidon = False
    Set DELT = EID.Recordset
    If IsNull(DELT("metaf")) Or DELT("metaf") = 0 Then
        DELT("metaf") = 1: DELT.Update
    Else
        DELT("metaf") = 0: DELT.Update
    End If
 End If
  
Else

'   TIMEDIT.Recordset.Edit
'   TIMEDIT.Recordset("kode") = EID.Recordset("kod")
'   TIMEDIT.Recordset("onoma") = EID.Recordset("ono")
'   If fBuff = "r" Then
'      TIMEDIT.Recordset("timm") = EID.Recordset("xti")
'   Else
'      TIMEDIT.Recordset("timm") = EID.Recordset("lti")
'   End If
'   TIMEDIT.Recordset("fpa") = EID.Recordset("fpa")
'   TIMEDIT.Recordset.Update
'   DBGrid1.Refresh
   
End If
   

   
   
Exit Sub

MHNYMA:
HandleError "Par1:GridEidon_RowColChange"
Resume Next

End Sub

Private Sub Form_Load()
'==============================
'==============================
'==============================
'==============================

Dim K, L As Integer, X As String, db
Dim R As New ADODB.Recordset
 
' par1.SHOW
 
 F_add_barc = Val(FindParametroi("PAR1", "add_barc", "1", "Γ.16 Barcode έιναι 1.Κωδ 2.Βοηθ.Κλειδί"))
 PAR1.KeyPreview = True
 
 f_PALIAXRONIA = Trim(FindParametroi("PAR1", "F_PALIAXRONIA", "2005", "Γ.18 DSN ΠΑΛΙΑΣ ΧΡΟΝΙΑΣ"))
 F_EXO_XROMATA = Val(FindParametroi("PAR1", "F_EXO_XROMATA", "0", "Γ. 1 ΧΡΗΣΙΜΟΠΟΙΩ ΜΕΓΕΘΗ-ΧΡΩΜΑΤΑ=1")) 'posa psifia tha exei h kathe seira
  
  
  f_Fkey = Trim(FindParametroi("PAR1", "F_FKEY", "112", "ΤΟ ΠΛΗΚΤΡΟ ΠΟΥ ΧΡΗΣΙΜΟΠΟΙΩ ΓΙΑ UPDATE"))
 
 F_EIKONA = Val(FindParametroi("APOT2", "F_EIKONA", "1", "1=EIKONA=>ΚΩΔΙΚΟΣ 2=>ΜΕ ΒΟΗΘ.ΚΛΕΙΔΙ 11=ΚVDIKOS.JPG 12=BOH.JPG 0=ΟΧΙ ΕΙΚΟΝΕΣ"))

 F_PAR_MONO_ME_YPOLOIPO = Val(FindParametroi("PAR1", "F_PAR_MONO_ME_YPOLOIPO", "0", "Γ. 30 ΠΑΡΑΓΓΕΛΙΑ ΠΕΛΑΤΗ ΜΟΝΟ ΑΝ ΥΠΑΡΧΕΙ ΥΠΟΛΟΙΠΟ"))

 F_STOPDELTIA = Val(FindParametroi("PAR1", "F_STOPDELTIA", "0", "Α. 2.Αποτρέπει την τιμολόγηση αν εκκρεμούν δελτία=1")) 'posa psifia tha exei h kathe seira

 F_2CHFIA_ZYGIZOMENA = FindParametroi("PAR1", "F_2CHFIA_ZYGIZOMENA", "21", "Γ. 2. ΠΡΩΤΑ ΨΗΦΙΑ ΒΟΗΘ.ΚΩΔ.ΤΩΝ ΖΥΓΙΖΟΜΕΝΩΝ") '2 ΠΡΩΤΑ ΨΗΦΙΑ ΒΟΗΘ.ΚΩΔ.ΤΩΝ ΖΥΓΙΖΟΜΕΝΩ

f_Desmia = Val(FindParametroi("PAR1", "F_DESMIA", "0", "Γ. 3.ΟΙ ΠΑΡΑΓΓΕΛΙΕΣ ΥΠΟΛΟΓΙΖΟΝΤΑΙ ΣΤΟ ΥΠΟΛΟΙΠΟ=1 ")) ' χρησιμοποιω τις παραγγελίες για έλεγχο υπολοίπου  'βλεπε mdesmia
 
 
f_mono_apoParaggelia = Val(FindParametroi("PAR1", "f_mono_apoParaggelia", "0", "Γ. 3.5 ΤΙΜΟΛΟΓΩ ΜΟΝΟ ΑΝ ΥΠΑΡΧΕΙ ΠΑΡΑΓΓΕΛΙΑ "))  '

f_SYMPYKNVSH = Val(FindParametroi("PAR1", "f_SYMPYKNVSH", "0", "Γ. 3.6 ΣΥΜΠΥΚΝΩΣΗ ΣΤΟΝ ΙΔΙΟ ΚΩΔΙΚΟ "))   '

F_TRANS = Val(FindParametroi("PAR1", "F_TRANS", "0", "Γ. 3.7 ΔΟΥΛΕΥΕΙ TRANSACTIONS=1  "))   '


F_MONOMEYPOLOIPO = Val(FindParametroi("PAR1", "F_MONOMEYPOLOIPO", "0", "Γ. 9.ΕΜΦΑΝΙΣΗ ΜΟΝΟ ΤΩΝ ΕΙΔΩΝ ΜΕ ΥΠΟΛΟΙΠΟ=1")) '



'  'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
 
'    'SkinFramework1.ApplyWindow Me.hWnd
'    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

    Dim DELT As String
    Dim nDelt As Integer
 
Dim sql As String
   sql = "SELECT convert(CHAR(10),HME,3) AS [ΗΜΕΡ/ΝΙΑ],ATIM,SKOPOS,EPO" _
   & " FROM TIM INNER JOIN PEL ON PEL.EIDOS=TIM.EIDOS AND PEL.KOD=TIM.KPE" _
   & " WHERE ART IS NULL AND LEFT(ATIM,1)='A' and SKOPOS LIKE '%ΩΛ%' AND" _
   & " HME<'" + Format(DateAdd("D", -25, Now), "MM/DD/YYYY") + "' and HME>'" + Format(DateAdd("D", -31, Now), "MM/DD/YYYY") + "' ORDER BY HME"
   
 R.Open sql, Gdb, adOpenStatic, adLockReadOnly
 If R.EOF Then
   'OK
 Else
    nDelt = 0
    DELT = "Τα παρακάτω δελτία είναι ατιμολόγητα " + Chr(13)
    Do While Not R.EOF
     If InStr(R("skopos"), "ΩΛ") > 0 Then
        DELT = DELT + R("atim") + " " + R(0) + " " + R("skopos") + " " + R("epo") + Chr(13)
        nDelt = nDelt + 1
     End If
      R.MoveNext
    Loop
 
    If nDelt > 0 Then
       MsgBox DELT
    End If
    
    
  End If
  R.Close

 
 If F_STOPDELTIA = 1 Then
 
   sql = "SELECT convert(CHAR(10),HME,3) AS [ΗΜΕΡ/ΝΙΑ],ATIM,SKOPOS,EPO" _
   & " FROM TIM INNER JOIN PEL ON PEL.EIDOS=TIM.EIDOS AND PEL.KOD=TIM.KPE" _
   & " WHERE ART IS NULL AND LEFT(ATIM,1)='A' and SKOPOS LIKE '%ΩΛ%' AND" _
   & " HME<'" + Format(DateAdd("D", -27, Now), "MM/DD/YYYY") + "' and HME>'" + Format(DateAdd("D", -31, Now), "MM/DD/YYYY") + "' ORDER BY HME"
   R.Open sql, Gdb, adOpenStatic, adLockReadOnly
   If R.EOF Then
      'ok
   Else
    
    nDelt = 0
    DELT = "Θα σταματήσει η τιμολόγηση επειδή " + Chr(13) + "Τα παρακάτω δελτία είναι ατιμολόγητα " + Chr(13)
    Do While Not R.EOF
     If InStr(R("skopos"), "ΩΛ") > 0 Then
        DELT = DELT + R("atim") + " " + R(0) + " " + R("skopos") + " " + R("epo") + Chr(13)
        nDelt = nDelt + 1
     End If
      R.MoveNext
    Loop
    If nDelt > 0 Then
       MsgBox DELT, vbCritical
    End If
    
    
  End If
  R.Close
  F_STOPDELTIA = 99
End If


 
 
 
 
 
 
 
 On Error GoTo MHNYMA
10
geidos_paras = 1
' On Error GoTo 0
'tropos plhromhs
R.Close
R.Open "SELECT *FROM PINAKES WHERE TYPOS=12 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
   K = 0
Do While Not R.EOF
   If R("typos") = 12 Then
      Combo4.AddItem Str(R("AYJON")) + "." + R("PERIGRAFH")
      F_LOGPLHR(K) = Format(R("TIMH"), "0000")
      K = K + 1
   End If
   R.MoveNext
Loop
 R.Close
K = 0





F_DIOR = 0
 
 'ΦΟΡΤΩΜΑ ΑΠΟΘΗΚΩΝ
 R.Open "SELECT *FROM PINAKES WHERE TYPOS=4 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
Do While Not R.EOF
      APOTH1.AddItem LTrim(Str(R("AYJON"))) + "." + R("PERIGRAFH")
      APOTH2.AddItem LTrim(Str(R("AYJON"))) + "." + R("PERIGRAFH")
   R.MoveNext
Loop
R.Close

20
If Len(Dir("C:\LAGEURO\NOPONTOI.TXT", vbNormal)) > 0 Then
    Command2.Visible = True
Else
    Command2.Visible = False
End If


f_use_Sysk = Val(FindParametroi("PAR1", "F_USE_SYSK", "0", "Γ. 4. Χρησιμοποιώ συσκευασία=1")) ' πολαπλασιάζει την ποσότητα επί την συσκευασία

F_falcon_Polla_tim = Val(FindParametroi("PAR1", "F_FALCON_POLLA_TIM", "0", "Γ. 5. Πολλαπλά τιμολόγια από PDA=1")) 'posa psifia tha exei h kathe seira

f_apeyueias = Trim(FindParametroi("PAR1", "f_APEYUEIAS", "0", "Α. 3. Κανονικα=0/Απ ευθείας στο παρ/κό με κωδ.πελάτη..")) '

f_1choice_agores = Val(FindParametroi("PAR1", "f_1choice_agores", "2", "Α. 4. Πρώτη επιλογή αγορών")) 'posa psifia tha exei h kathe seira
f_1choice_polhseis = Val(FindParametroi("PAR1", "f_1choice_polhseis", "2", "Α. 5. Πρώτη επιλογή πωλήσεων")) 'posa psifia tha exei h kathe seira

f_AutoChangeLine = Val(FindParametroi("PAR1", "f_AutoChangeLine", "0", "Α. 6 Αλλάζει αυτόματα σειρά=1"))  'def="1"  "1=μπορω να αλλάξω έκπτωση γραμμής 0=δεν μπορώ

f_CanChangeEkptosi = Val(FindParametroi("PAR1", "f_CanChangeEkptosi", "1", "Α. 7. Αλλάζω έκπτωση γραμμής=1 ")) 'def="1"  "1=μπορω να αλλάξω έκπτωση γραμμής 0=δεν μπορώ
f_CanChangeFpa = Val(FindParametroi("PAR1", "f_CanChangeFpa", "1", "Α. 8 Αλλάξω ΦΠΑ γραμμής=1")) 'def="1"  "1=μπορω να αλλάξω  γραμμής 0=δεν μπορώ



f_psifiaAjias = Val(FindParametroi("PAR1", "f_psifiaAjias", "2", "Γ. 6. Δεκαδικά Ψηφία Αξίας Σειρών τιμολογιου")) 'posa psifia tha exei h kathe seira

f_PARAGGMARKUP = Val(FindParametroi("PAR1", "f_PARAGGMARKUP", "0", "Γ. 7. Στην Παραγγελία Πελ.0=Χονδ 1=Markup"))

f_MinSeiresTimologioy = Val(FindParametroi("PAR1", "F_MINSEIRESTIMOLOGIOY", "25", "Β. 1 Σειρές τιμ.Πώλησης"))

f_MaxSeiresTimologioy = Val(FindParametroi("PAR1", "F_MAXSEIRESTIMOLOGIOY", "80", "Β. 2 Σειρές Τιμ.Αγοράς"))
f_SynSeiresTimologioy = Val(FindParametroi("PAR1", "f_SynSeiresTimologioy", "25", "Β. 3 Σειρές ΤΙΜ.ΠΩΛ ΚΕΝΕΣ+ΓΡΑΜ"))

Grid1.ROWS = f_MaxSeiresTimologioy

f_checkPos = Val(FindParametroi("PAR1", "F_CHECKPOS", "DELETE", " 1=ΕΛΕΓΧOΣ 2=ΑΠΑΓΟΡΕΥΣΗ 0=ΤΙΠΟΤΑ"))


f_checkYpol = Val(FindParametroi("PAR1", "F_checkYpol", "0", "Α. 8.ΥΠΟΛΟΙΠΟ ΕΙΔΟΥΣ ΕΛΕΓΧOΣ/ΑΠΑΓΟΡΕΥΣΗ/ΤΙΠΟΤΑ 1/2/0"))

f_TIMOKAT = Val(FindParametroi("PELAT2", "F_TIMOKAT", "0", "Γ. 8. 0=ΓΕΝ.ΕΚΠΤΩΣΗ ΠΕΛΑΤΗ  1..100= Νο ΤΙΜΟΚΑΤΑΛΟΓΟΥ"))

f_psifia_atim = Val(FindParametroi("PAR1", "F_PSIFIA_ATIM", "5", "Γ. 9. ΨΗΦΙΑ ΑΡΙΘΜΟΥ ΤΙΜΟΛΟΓΙΟΥ (5)"))
If f_psifia_atim < 5 Then f_psifia_atim = 5


'Dim f_barcode As Integer '1=diabazei apeyueias ta barcode
f_barcode = Val(FindParametroi("PAR1", "F_BARCODE", "0", "Α. 9 EIΔΗ: 0=ΑΠ ΕΥΘΕΙΑΣ Ο ΚΩΔΙΚΟΣ 1=ΤΟ BARCODE"))

f_LIANIKHCHECK = Val(FindParametroi("PAR1", "f_LIANIKHCHECK", "1", "Α.10  1=ΛΙΑΝΙΚΗ ΜΑΡΚΑΡΙΣΜΕΝΗ 0=ΟΧΙ"))
If f_LIANIKHCHECK = 0 Then lianikh.value = vbUnchecked Else lianikh.value = vbChecked
   

F928 = Val(FindParametroi("PAR1", "F928", "0", "Α.11 Εκτύπωση φόρμας 1=928 0=437"))

F_File_Timologioy = FindParametroi("PAR1", "F_File_Timologioy", "c:\print", "Α.12 Αρχείο εκτύπωσης παραστατικού")
F_Type_File_Timologioy = Val(FindParametroi("PAR1", "F_Type_File_Timologioy", "1", "Α.13 1=εκτύπωση αρχείου παρ/κού 0=Οχι"))

f_F12 = Val(FindParametroi("PAR1", "F_F12", "-1", "Α.14 F12:Τύπος Παραστατικού που ενημερώνεται"))
f_F1 = Val(FindParametroi("PAR1", "F_F1", "-1", "Α.15 F1:Τύπος Παραστατικού που ενημερώνεται"))

f_BARCODE_STOTIM = Val(FindParametroi("PAR1", "F_BARCODE_STOTIM", "0", "Α.16  1=ΒΑΖΕΙ ΤΟ ERG ΣΤΟ ΤΙΜΟΛΟΓΙΟ 0=ΤΟΝ ΚΩΔΙΚΟ(def)"))






Dim PROOR As String
Dim APOSTOLH As String

'PROOR = Val(FindParametroi("PAR1", "PROOR", "1", "ΕΔΡΑ ΠΕΛΑΤΗ,Γ.ΖΕΡΒΟΥ 35"))
'APOSTOLH = Val(FindParametroi("PAR1", "APOSTOLH", "1", "ΕΔΡΑ ΜΑΣ,Κ.ΠΑΛΑΙΟΛΟΓΟΥ 21"))



'ΦΟΡΤΩΜΑ ΣΚΟΠΩΝ
 R.Open "SELECT *FROM PINAKES WHERE TYPOS=7 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
Do While Not R.EOF
      SKOPOS.AddItem R("PERIGRAFH")
   R.MoveNext
Loop
If SKOPOS.ListCount = 0 Then
  Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (1,'ΠΩΛΗΣΗ',7)"
  Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (2,'ΕΠΙΣΚΕΥΗ',7)"
  Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (3,'ΑΠΟ ΕΠΙΣΚΕΥΗ',7)"
  Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (4,'ΕΠΙΣΤΡΟΦΗ',7)"
  Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (5,'ΑΝΤΙΚΑΤΑΣΤΑΣΗ',7)"
  Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (6,'ΜΕΤΑΦΟΡΑ',7)"
End If
R.Close


'ΦΟΡΤΩΜΑ FORTOSHS
 R.Open "SELECT *FROM PINAKES WHERE TYPOS=8 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
Do While Not R.EOF
      FORTOSH.AddItem R("PERIGRAFH")
   R.MoveNext
Loop
If FORTOSH.ListCount = 0 Then
  Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (1,'ΕΔΡΑ ΜΑΣ',8)"
End If
R.Close

'ΦΟΡΤΩΜΑ ΠΡΟΟΡΙΣΜΟΥ
 R.Open "SELECT *FROM PINAKES WHERE TYPOS=9 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
 
 ' If Not R.EOF Then R.MoveFirst
 
Do While Not R.EOF
      PROORISMOS.AddItem R("PERIGRAFH")
   R.MoveNext
Loop

If PROORISMOS.ListCount = 0 Then
  Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (1,'ΕΔΡΑ ΠΕΛΑΤΗ',9)"
End If
R.Close


'ΦΟΡΤΩΜΑ AYTOKINHTVN
 R.Open "SELECT *FROM PINAKES WHERE TYPOS=10 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
Do While Not R.EOF
      AYTOKINHTO.AddItem R("PERIGRAFH")
   R.MoveNext
Loop
If AYTOKINHTO.ListCount = 0 Then
  Gdb.Execute "INSERT INTO PINAKES (AYJON,PERIGRAFH,TYPOS) VALUES (1,'   ',10)"
End If
R.Close









'On Error Resume Next

SKOPOS.Text = SKOPOS.List(0)
 FORTOSH.Text = FORTOSH.List(0)
 PROORISMOS.Text = PROORISMOS.List(0)
 



' AN EINAI YPOKATASTHMA
If Len(Dir("C:\KENTRIKO.TXT")) = 0 Or Len(Dir("C:\LAGEURO\MPOYGET.EXE", vbNormal)) = 0 Then
   APOTH1.Text = APOTH1.List(1)
   APOTH2.Text = APOTH2.List(0)
'   FORTOSH.Text = FORTOSH.List(1)
Else
   APOTH1.Text = APOTH1.List(0)
   APOTH2.Text = APOTH2.List(1)
End If

If Len(Dir("C:\KENTRIKO.TXT")) > 0 Then
   APOTH1.Text = APOTH1.List(0)
   APOTH2.Text = APOTH2.List(1)
'   FORTOSH.Text = FORTOSH.List(1)
End If

If Len(Dir("C:\YPOK2.TXT")) > 0 Then
   APOTH1.Text = APOTH1.List(2)
   APOTH2.Text = APOTH2.List(0)
End If

If Len(Dir("C:\YPOK3.TXT")) > 0 Then
   APOTH1.Text = APOTH1.List(3)
   APOTH2.Text = APOTH2.List(0)
End If

'ΦΟΡΤΩΜΑ ΠΩΛΗΤΩΝ
 R.Open "SELECT *FROM PINAKES WHERE TYPOS=5 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
Do While Not R.EOF
      Combo8.AddItem LTrim(Str(R("AYJON"))) + "." + R("PERIGRAFH")
   R.MoveNext
Loop
R.Close
30
Combo8.Text = Combo8.List(0)



f_SeiresTimologioy = 80 'f_SeiresTimologioy = 30 OTAN DIALEGO  POLISEIS

f_gDek_ajia = "###,###.00"
f_ColUpdate = False
f_error_gefyr = 0



ProthFora = False


' πινακας γεφυρώσεων
FlxGridDeltion.Clear
FlxGridDeltion.Visible = False

40


Dim fDB As Database
 'Dim R As Recordset
 
 
 
 
Grid1.ColS = 10

Grid1.ColWidth(0) = 300
Grid1.ColWidth(f_k) = 1200
Grid1.ColWidth(f_o) = 3800
Grid1.ColWidth(f_p) = 1000
Grid1.ColWidth(F_T) = 1000
Grid1.ColWidth(f_e) = 800
Grid1.ColWidth(f_f) = 800

Grid1.Col = 0
Grid1.row = 1
Grid1.Text = " "
Grid1.Col = 0
For K = 1 To Grid1.ROWS - 1
      
   Grid1.row = K
   Grid1.Text = Format(K, "##")
   Grid1.RowHeight(K) = 250
Next
Grid1.row = 0
Grid1.Col = f_k
Grid1.Text = "Κωδικός"
Grid1.Col = f_o
Grid1.Text = "Περιγραφή"
Grid1.Col = f_p
Grid1.Text = "Ποσότητα"
Grid1.Col = F_T
Grid1.Text = "Τιμή Μονάδος"
Grid1.Col = f_e
Grid1.Text = "Εκπτωση"
Grid1.Col = f_a
Grid1.Text = "Αξία"

Grid1.Col = f_f
Grid1.Text = "Kατ.ΦΠΑ"


GridEidon.Columns(0).width = 18000

50









 
 'data1.DATABASENAME = gDir
 'data1.RecordSource = "SELECT *FROM PINAKES"
 
 
'Grid1.Row = 20
'Fdb.Open gDir
 'R.Open "SELECT *FROM mem", db, adOpenKeyset, adLockReadOnly ', adCmdTableDirect
 
 
'If gConnect = "Access" Then
 ' Set fDB = OpenDatabase(gDir, False, False)
'Else
'  Set fDB = OpenDatabase(gDir, False, False, gConnect)
'End If

Synolo = " " ' ΓΙΑ ΝΑ ΜΗΝ ΒΓΑΖΕΙ NULL


'
'If gConnect = "Access" Then
'  Set DB = OpenDatabase(gDir, False, False)
'Else
'  Set DB = OpenDatabase(gDir, False, False, gConnect)
'End If
'Set r = DB.OpenRecordset("select *from mem")
'r.MoveFirst
'r.Move 3
'
''r.MoveNext
'gEPITR_SEIR = r("c3")
'r.Close


'For k = 1 To Len(gEPITR_SEIR)
'    POLAGOR.AddItem Mid$(gEPITR_SEIR, k, 1)
'Next


'TIMEDIT.DatabaseName = gDir
' TIMEDIT.RecordSource = "SELECT *FROM TIMEDIT"
' TIMEDIT.ConnectionString = gConnect

'EID.DatabaseName = gDir
EID.RecordSource = "SELECT TOP 1 *FROM EID"
EID.ConnectionString = gConnect

'PEL.DatabaseName = gDir
PEL.RecordSource = "SELECT *FROM PEL"
PEL.ConnectionString = gConnect

EGGTIM.ConnectionString = gConnect
EGGTIM.RecordSource = "SELECT top 10 * FROM EGGTIM"
    
DOK.ConnectionString = gConnect
DOK.RecordSource = "SELECT *FROM DOK"
    
    
LastTimol.ConnectionString = gConnect
LastTimol.RecordSource = "SELECT top 50 HME,ATIM,AJI AS [AΞΙΑ],PEL.EPO AS [ΕΠΩΝΥΜΙΑ],PEL.KOD AS [ΚΩΔΙΚΟΣ] FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD ORDER BY ID_NUM DESC"
LastTimol.Refresh
TDBGrid1.MarqueeStyle = 3  'SEIRA MARKARISMENH
TDBGrid1.Splits(0).AlternatingRowStyle = True 'ENALAX XROMATA SEIRON
TDBGrid1.Splits(0).ExtendRightColumn = True


    
    
    
'TIMEDIT.Enabled = True
EID.Enabled = True
PEL.Enabled = True
    
DTPicker1.value = Now

    
PAR1.Picture = LoadPicture(gPicture)
    
    
'TIMEDIT.Recordset.MoveFirst


'TIMEDIT.Refresh
 
 
'R.Open "SELECT *FROM TIMEDIT", Fdb, adOpenKeyset, adLockOptimistic

 
 
' If R.RecordCount < 40 Then
'    For K = 1 To 40
'      R.AddNew
'      R.Update
'    Next
' Else
'    For K = 1 To 40
      'TIMEDIT.Recordset.Edit
 '     For l = 0 To R.Fields.Count - 1
  '            R(l) = Null
   '   Next
'      TIMEDIT.Recordset.Update
      'TIMEDIT.Recordset(0) = "."
      
'      TIMEDIT.Recordset.MoveLast
      
 '     TIMEDIT.Recordset.Move K - 1
      
'      R.Update
'      R.MoveNext
'    Next
 'End If
 'R.MoveFirst
  
 
60
 
 
 SEIRA.Clear
 X = find_seires(0)
 For K = 1 To Len(X)
    SEIRA.AddItem mID$(X, K, 1)
 Next
SEIRA.Text = SEIRA.List(0)
70
' Combo4.Clear
' Combo4.AddItem "ΜΕΤΡΗΤΟΙΣ"
' Combo4.AddItem "ΠΙΣΤΩΣΗ"
' Combo4.AddItem "ΠΙΣΤ.ΚΑΡΤΑ"
 Combo4.Text = Combo4.List(0)
80
 
 
 POLAGOR.Text = POLAGOR.List(1)
POLAGOR_LostFocus
90
PARAS.Text = "" ' PARAS.List(geidos_paras)
 
 
 'DbGrid1.Columns(2).NumberFormat = gDek_pos 'δεκαδικά ποσοτήτων
 'DbGrid1.Columns(3).NumberFormat = gDek_ajia 'δεκαδικά αξίας
 'DbGrid1.Columns(6).NumberFormat = gDek_ajia 'δεκαδικά αξίας
 
  
 'PARAS.SetFocus
100
  PARAS.Enabled = True
'On Error GoTo 0
'PARAS.SetFocus
 
 
 
 
 
'μπαινω απ'ευθειας στο παραστατικό
If Val(f_apeyueias) > 0 Then
  PARAS_LostFocus
  Text2(0) = f_apeyueias
  SendKeys "{ENTER} {ENTER}{ENTER} {ENTER}"
End If
 
 
 
 
 
 
 
 
 
 
 
 
 
 Exit Sub

MHNYMA:
HandleError "Par1:form_load"
Resume Next
            
 
 
 
End Sub

Function find_parastat(recno As Integer, par As String, ReadWrite As Integer) As Long
'ReadWrite=0   Βρίσκει τον αριθμο παραστατικού          -------------------------
'ReadWrite=1   Aποθηκεύει τον αριθμο παραστατικού + 1   -------------------------
'Βρίσκει τις κατηγορίες φπα

Dim K As Integer
Dim X, Bohuhtiko As String
On Error GoTo MHNYMA


'Dim db As Database
 Dim R As New ADODB.Recordset
  R.Open "SELECT * FROM PARASTAT WHERE EIDOS='" + par + "'", Gdb, adOpenDynamic, adLockOptimistic

If IsNull(R("ARITMISI")) Then
   X = 0
Else
  X = R("ARITMISI")
End If
R.Close

'MEM..MoveFirst
'For k = 1 To 9: f_mfpa(k) = R("fpa" + LTrim(Str(k))): Next


If X = 0 Then
   find_parastat = 0
   Exit Function
End If

If ReadWrite = 0 Then
    R.Open "SELECT * FROM ARITMISI WHERE ID=" + Str(X), Gdb, adOpenDynamic, adLockOptimistic
    If IsNull(R("ARITMISI")) Then
       find_parastat = 0
    Else
       find_parastat = R("ARITMISI") + 1
    End If
    R.Close
    
   
Else
   Gdb.Execute "UPDATE ARITMISI SET ARITMISI=" + Text1.Text + "  WHERE  ID=" + Str(X)
   
End If

Exit Function

MHNYMA:
HandleError "Par1:FIND_Parastat"
Resume Next
            



End Function
Function find_seires(DUM As Integer) As String

'Βρίσκει τις υπάρχουσες σειρές
Dim db As Database
Dim X As String

'If gConnect = "Access" Then
'   Set db = OpenDatabase(gDir, False, False)
'Else
'   Set db = OpenDatabase(gDir, False, False, gConnect)
'End If

'Dim db As Database
 Dim R As New ADODB.Recordset
 
'If gConnect = "Access" Then
 ' Set db = OpenDatabase(gDir, False, False)
'Else
 ' Set db = OpenDatabase(gDir, False, False, gConnect)
'End If
 
 
 
 
 
' Data1.DatabaseName = gDir
' Data1.RecordSource = "SELECT *FROM PINAKES"
 
 
 'Fdb.Open gDir
 'R.Open "SELECT c3 FROM mem", Fdb, adOpenStatic, adLockReadOnly
 
 'adLockReadOnly ', adCmdTableDirect





R.Open "select C3 from MEM", Gdb, adOpenDynamic, adLockOptimistic
R.MoveFirst
R.Move 3

If IsNull(R("c3")) Then
   find_seires = " "
Else
   find_seires = RTrim(R("c3"))
End If
R.Close
End Function

Function Open_Data(ByVal mtable As String) As Recordset

 Set Open_Data = fDB.OpenRecordset(mtable)
End Function

Private Sub FlxGridDeltion_Click()
'If Grid1.Enabled = False Then ' φάση δελτίων αποστολής
   FlagGridEidon = True
'  If ProthFora = False Then
 '    ProthFora = True
     ' Set delt = EID.Recordset
    
    DoEvents
    
    If Val(FlxGridDeltion.TextMatrix(FlxGridDeltion.row, 1)) = 0 Then ' If IsNull(EID.Recordset("metaf")) Or EID.Recordset("metaf") = 0 Then
         FlxGridDeltion.TextMatrix(FlxGridDeltion.row, 1) = 1
         FlxGridDeltion.Col = 2
         FlxGridDeltion.CellBackColor = vbYellow
         'EID.Recordset("metaf") = 1: EID.Recordset.Update
     Else
         'EID.Recordset("metaf") = 0: EID.Recordset.Update
         FlxGridDeltion.TextMatrix(FlxGridDeltion.row, 1) = 0
         
         FlxGridDeltion.Col = 2
         FlxGridDeltion.CellBackColor = vbWhite
         
     End If
     FlxGridDeltion.Refresh
     
     
  '   Grid1.Enabled = True
  'End If
'End If
'EID.Refresh
End Sub

Private Sub TDBGrid1_Click()
Dim MC
'MC = 0
'  Load LAST1(MC)
'  LAST1(MC).Caption = "ΕΜΦΑΝΙΣΗ ΠΑΡΑΣΤΑΤΙΚΟΥ"
''   Load LAST1(MC)
'
'   MC = MC + 1
'      Load LAST1(MC)
'   LAST1(MC).Caption = "ΔΙΟΡΘΩΣΗ/ΑΚΥΡΩΤΙΚΟ"
'
'
'  'MC = MC + 1
'  ' LAST1(MC).Caption = "ΕΜΦΑΝΙΣΗ ΠΑΡΑΣΤΑΤΙΚΟΥ"
'
''   Load LAST1(MC)
'
''   LAST1(MC).Caption = "ΔΙΟΡΘΩΣΗ/ΑΚΥΡΩΤΙΚΟ"
Call Me.PopupMenu(menupopup)

'
  
  
  
  
End Sub

Private Sub par2showing()

 
  Dim FF As PAR2
   
  TDBGrid1.Col = 1
Set FF = New PAR2

FF.Text1 = TDBGrid1.Text

  TDBGrid1.Col = 0
FF.DTPicker1 = CDate(TDBGrid1.Text)
FF.DTPicker2 = CDate(TDBGrid1.Text)

FF.CallFromOther.Caption = "1"

'GIA OLA TA PARASTATIKA
FF.Combo2.Text = "*"



  FF.SHOW
  

FF.bohuitiko.SetFocus
DoEvents

FF.Command1.SetFocus

  
  
  

  
  
'  SendKeys "~" 'ENTER
'  SendKeys "~" 'enter  {DOWN}"
' SendKeys "~" 'ENTER
 SendKeys "~" 'ENTER
''FF.Command1_Click






End Sub

Private Sub Text1_GotFocus()
     Text1.BackColor = vbYellow
     Text1.SelStart = 0
     Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       KeyAscii = 0
       If POLAGOR.Text = POLAGOR.List(0) Then
          DTPicker1.SetFocus
       Else
          selectPEL.SetFocus
       End If
  End If
End Sub
Private Sub Grid1_EnterCell()
'  If Grid1.MouseRow = 0 Then
'    timtext3.Visible = False
'    Exit Sub
'   End If

On Error GoTo MHNYMA

   If Grid1.row > f_SeiresTimologioy Then
    Grid1.row = f_SeiresTimologioy
   End If
   
   timText3.Text = Grid1.Text
   timText3.Visible = False
   timText3.Top = Grid1.Top + Grid1.CellTop
   timText3.Left = Grid1.Left + Grid1.CellLeft + Grid1.Col * 100
   
   timText3.width = Grid1.CellWidth  ' - 15
   timText3.height = Grid1.CellHeight - 25
   
   timText3.SelStart = 0
   timText3.SelLength = Len(timText3.Text)
   

   
   
   
   
   timText3.Visible = True
   timText3.SetFocus
   
   
   Exit Sub

MHNYMA:
HandleError "Par1:EnterCell"
Resume Next
            
   
   
   
   
End Sub

Private Sub Grid1_LeaveCell()
'Grid1.Text = timtext3.Text
End Sub



Private Sub Text1_LostFocus()
Text1.BackColor = vbWhite
End Sub

Private Sub Text2_GotFocus(Index As Integer)
   Text2(Index).BackColor = vbYellow
End Sub

Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 
'  On Error GoTo MHNYMA
'
' If KeyCode = 13 Then
'
'   KeyCode = 0
'   If Index = 0 Then
'
'     If Len(Trim(Text2(0).Text)) = 0 Then
'        Text2(1).SetFocus ' δώσε όνομα
'     Else
'        GridPelaton.Columns(0).width = 5000
'        GridPelaton.Visible = True
'        GridPelaton.ZOrder 0
'        GridPelaton.SetFocus 'διαλεξε πελάτη
'     End If 'len
'
'   End If 'index=0
'
'   If Index = 1 Then
'
'     If Len(Trim(Text2(1).Text)) = 0 Then
'
'        Text2(2).SetFocus ' δώσε ΑΦΜ
'
'     Else
'        GridPelaton.Visible = True
'        GridPelaton.ZOrder 0
'        GridPelaton.SetFocus 'διαλεξε πελάτη
'
'     End If 'len
'
'   End If 'index=0
'
'   If Index = 2 Then
'        GridPelaton.Visible = True
'        GridPelaton.ZOrder 0
'        GridPelaton.SetFocus 'διαλεξε πελάτη
'
'   End If 'index=0
'
'End If ' keyascci=13
'
'
'Exit Sub
'
'MHNYMA:
'HandleError "Par1:text2_keyUP"
'Resume Next
'
'
'




End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
        KeyAscii = 0
       ' selectPEL.SetFocus
   End If
End Sub

Private Sub Text2_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

  On Error GoTo MHNYMA

 If KeyCode = 13 Then

   KeyCode = 0
   If Index = 0 Then

     If Len(Trim(Text2(0).Text)) = 0 Then
        Text2(1).SetFocus ' δώσε όνομα
     Else
        GridPelaton.Columns(0).width = 5000
        GridPelaton.Visible = True
        GridPelaton.ZOrder 0
        GridPelaton.SetFocus 'διαλεξε πελάτη
     End If 'len

   End If 'index=0

   If Index = 1 Then

     If Len(Trim(Text2(1).Text)) = 0 Then

        Text2(2).SetFocus ' δώσε ΑΦΜ

     Else
        GridPelaton.Visible = True
        GridPelaton.ZOrder 0
        GridPelaton.SetFocus 'διαλεξε πελάτη

     End If 'len

   End If 'index=0

   If Index = 2 Then
        GridPelaton.Visible = True
        GridPelaton.ZOrder 0
        GridPelaton.SetFocus 'διαλεξε πελάτη

   End If 'index=0

End If ' keyascci=13


Exit Sub

MHNYMA:
HandleError "Par1:text2_keyUP"
Resume Next


End Sub

Private Sub Text2_LostFocus(Index As Integer)

Dim X As String
On Error GoTo MHNYMA


Text2(Index).BackColor = vbWhite

If Len(Trim(Text2(Index))) = 0 Then Exit Sub
If fbuff = "l" Then fbuff = "r"


X = fbuff 'If Option1(1) Then x = "e" Else x = "r"


If Index = 1 Then  'eponumo
    PEL.RecordSource = "select LEFT(left(EPO+'-'+DIE+'-',56)+KOD,60) as xx,DIE,THL,KOD  from PEL where EIDOS='" + X + "' and EPO like '" + UCase(SameLetters(Text2(1).Text)) + "%' order by EPO;"
End If

If Index = 0 Then 'kodikos
    PEL.RecordSource = "select  LEFT(left(EPO+'-'+DIE+'-',56)+KOD,60) as xx,DIE,THL,KOD  from PEL where EIDOS='" + X + "' and KOD like '" + SameLetters(Text2(0).Text) + "%' order by KOD;"
End If

If Index = 2 Then 'afm
    PEL.RecordSource = "select  LEFT(left(EPO+'-'+DIE+'-',56)+KOD,60) AS XX,DIE,THL,KOD  from PEL where EIDOS='" + X + "' and AFM like '" + Text2(2).Text + "%' order by AFM;"
End If

GridPelaton.Columns(0).width = 6000
GridPelaton.Columns(0).DataField = "XX"

PEL.Refresh
GridPelaton.Columns(0).width = 6000
GridPelaton.Visible = True
GridPelaton.SetFocus

Exit Sub

MHNYMA:
HandleError "Par1:FINDSYNOLO"
Resume Next
            
End Sub


Private Sub TIMText3_GotFocus()
     
     timText3.BackColor = vbYellow
     timText3.SelStart = 0
     timText3.SelLength = Len(timText3.Text)


End Sub

Private Sub timText3_KeyDown(KeyCode As Integer, Shift As Integer)
'ΕΝΤΡΥ ΤΙΜΟΛΟΓΙΟΥ
Dim DUM

If KeyCode = 13 Then
   KeyCode = 0
   DUM = 0
End If

If KeyCode > 95 Then
  Me.Caption = KeyCode
  DUM = 0
End If

End Sub

Private Sub timText3_KeyPress(KeyAscii As Integer)
'ΕΝΤΡΥ ΤΙΜΟΛΟΓΙΟΥ
Dim DUM
If KeyAscii = 13 Then
   KeyAscii = 0
   DUM = 0
End If

If KeyAscii = 27 Then
   KeyAscii = 0
   DUM = 0
End If



End Sub

Private Sub timText3_KeyUp(KeyCode As Integer, Shift As Integer)
'ΕΝΤΡΥ ΤΙΜΟΛΟΓΙΟΥ  750   1650  4mhnh
   Dim P, t, e, a, C
   Dim KOMA, mText
   Dim R0 As New ADODB.Recordset
   Dim XX As String

   Dim intRecord As Integer
   Dim intField As Integer

Dim Sql2 As String
Dim HDH_FORTOMENA As Single, LK0 As Long, n_HDH_FORTOMENA

 HDH_FORTOMENA = 0



On Error Resume Next






On Error GoTo MHNYMA
timText3.Text = Replace(timText3.Text, "'", "~")

If KeyCode = 46 And Grid1.Col = f_k Then ' DELETE ΣΒΗΝΩ ΤΗΝ ΣΕΙΡΑ
               Grid1.TextMatrix(Grid1.row, f_k) = ""
               Grid1.TextMatrix(Grid1.row, f_o) = ""
               Grid1.TextMatrix(Grid1.row, f_p) = ""
               Grid1.TextMatrix(Grid1.row, F_T) = ""
               Grid1.TextMatrix(Grid1.row, f_e) = ""
               Grid1.TextMatrix(Grid1.row, f_f) = ""
               Grid1.TextMatrix(Grid1.row, f_a) = ""
               Exit Sub
End If

  If KeyCode = 27 Then 'esc τελος εισαγωγής
       KeyCode = 0
       Combo4.SetFocus
  End If




If KeyCode = 121 And Shift = 0 Then ' --------------  F10 ----------------------
  b_Katax_Click
  Exit Sub

End If

 ' --------------  F6 -BGAZEI GRID ----- xroma-megetos -----------------------------------
If KeyCode = 117 And Shift = 0 Then   ' --------------  F6 ------ xroma-megetos
 ' UGridEdit.Label1.Caption = 0
  UGridEdit.Label1.Caption = Grid1.row
  
  Dim ColS As Integer, ROWS As Integer
  ColS = 0: ROWS = 0
 
 'ΦΟΡΤΩΝΩ ΤΙΣ ΕΠΙΚΕΦΑΛΙΔΕΣ ΤΟΥ GRID ΧΡΩΜΑΤΩΝ
 Dim MEGETHOS As String, XROMA As String
  XROMA = Format(29 + (2 * EID.Recordset("KATHGORIA") - 1), "00")
  MEGETHOS = Format(29 + (2 * EID.Recordset("KATHGORIA")), "00")
 If Val(XROMA) = 0 Then Exit Sub
 If Val(MEGETHOS) = 0 Then Exit Sub
 
 
 R0.Open "SELECT *FROM PINAKES WHERE TYPOS=" + XROMA + " ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
 Do While Not R0.EOF
     UGridEdit.Grid1.TextMatrix(R0("AYJON"), 0) = R0("PERIGRAFH")
'     Exit Sub
     ROWS = ROWS + 1
     XROMATA(Grid1.row, ROWS, 0) = R0("PERIGRAFH")
     R0.MoveNext
 Loop
 R0.Close
 'ΦΟΡΤΩΝΩ ΤΙΣ ΕΠΙΚΕΦΑΛΙΔΕΣ ΤΟΥ GRID ΧΡΩΜΑΤΩΝ
 R0.Open "SELECT *FROM PINAKES WHERE TYPOS=" + MEGETHOS + " ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
 Do While Not R0.EOF
     UGridEdit.Grid1.TextMatrix(0, R0("AYJON")) = R0("PERIGRAFH")
     ColS = ColS + 1
     ' Exit Sub
     XROMATA(Grid1.row, 0, ColS) = R0("PERIGRAFH")
     R0.MoveNext
 Loop
 R0.Close
  
  UGridEdit.Grid1.ROWS = ROWS + 1
  UGridEdit.Grid1.ColS = ColS + 1
  
  UGridEdit.XTI.Caption = EID.Recordset("XTI")
  UGridEdit.LTI5.Caption = EID.Recordset("LTI5")
  UGridEdit.kod.Caption = EID.Recordset("KOD")
  UGridEdit.ono.Caption = EID.Recordset("ONO")
  UGridEdit.CH1.Caption = EID.Recordset("CH1")
  UGridEdit.hme.Caption = Format(DTPicker1.value, "DD/MM/YYYY")
  
  UGridEdit.SHOW 1
    
  
  
  'ΦΟΡΤΩΝΩ ΤΑ ΧΡΩΜΑΤΑ ΣΕ ΠΙΝΑΚΑ
  Dim K0 As Long, K2 As Long, r2 As Long, c2 As Long
  Dim SUMA_XROMATA As Long
  SUMA_XROMATA = 0
  K2 = 1
  For r2 = 1 To ROWS
    For c2 = 1 To ColS
       K0 = InStr(K2, gm_str(Grid1.row), ";")
       If K0 = 0 Then
         XROMATA(Grid1.row, r2, c2) = ""
       Else
         XROMATA(Grid1.row, r2, c2) = mID(gm_str(Grid1.row), K2, K0 - K2)
         SUMA_XROMATA = SUMA_XROMATA + Val(XROMATA(Grid1.row, r2, c2))
       End If
       K2 = K0 + 1
    Next
  Next
If Grid1.Col = f_p Then 'AN PATOYSE PANV STHN POSOTHTA
   timText3.Text = SUMA_XROMATA
Else
   Grid1.TextMatrix(Grid1.row, f_p) = SUMA_XROMATA
End If
If SUMA_XROMATA > 0 Then
    XROMATA(Grid1.row, 0, 0) = -0.01 ' ΣΗΜΑΔΙ ΟΤΙ ΧΡΗΣΙΜΟΠΟΙΗΣΕ ΤΟ F6
End If

  
  
  Exit Sub
End If


' --------------  F6 ------ xroma-megetos -----------------------------------






If KeyCode = Val(Right(f_Fkey, 3)) And Shift = Val(Left(f_Fkey, 1)) Then ' --------------  F1 ----------------------
  If f_F1 > -1 Then
     PARAS.Text = PARAS.List(f_F1)
     PARAS_LostFocus
     b_Katax_Click
     Exit Sub
  End If
End If

If KeyCode = 123 And Shift = 0 Then  ' --------------  F12 ----------------------
  If f_F12 > -1 Then
     PARAS.Text = PARAS.List(f_F12)
     PARAS_LostFocus
     b_Katax_Click
     Exit Sub
  End If
End If








If KeyCode = 13 Then   'enter ------------------------------------------------
     KeyCode = 0
     
   If Len(Trim(timText3.Text)) = 0 Then   ' enter me keno koytaki  ***************
      Grid1.Text = timText3.Text ' κρατάει το grid το κείμενο
      ' μιά σειρά παρακάτω
      If Grid1.Col < Grid1.ColS - 4 Then  ' πάει στην επόμενη κολόνα
           Grid1.Col = Grid1.Col + 1
           Grid1_EnterCell
           Exit Sub
      End If

   End If '                                                     ****************
   
           '   GridEidon.width = 10000
              GridEidon.Columns(0).width = 1000
             GridEidon.Columns(1).width = 5000
              GridEidon.Columns(3).width = 1000
             
             
             GridEidon.Columns(2).width = 1000
     
     
     ' kod ή όνομα <2 τότε ψάχνουμε
     
     '******************************************************************************
     
     If Grid1.Col < f_p And (Grid1.Text <> timText3.Text Or Val(Grid1.TextMatrix(Grid1.row, f_p)) = 0) Then  ' αλλαξε ο κωδικός άρα ψάξε στο αρχείο
         
          ' ψαχνει κατά κωδικό   %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
10       If Grid1.Col = f_k Then
            timText3.Text = Trim(Left(timText3.Text, 16))
            
            
            If Left(timText3.Text, 1) = "." Then ' $$$$$$$
                 'EID.RecordSource = "select SUM(XRE)-SUM(PIS) AS [ΥΠΟΛ],EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO from EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD WHERE ERG  LIKE '" + Trim(Mid(timtext3.Text, 2, 14)) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) ;"
                 
                ' ψαχνει κατά BOHUHTIKO κωδικό  ERG %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                 EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)" _
                 & "-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," _
                 & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                 & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" _
                 & " WHERE ERG  LIKE '" + SameLetters(Trim(mID(timText3.Text, 2, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1 "
               
                 If F_MONOMEYPOLOIPO = 1 And f_pol = 1 Then
                   EID.RecordSource = EID.RecordSource + " HAVING [ΥΠΟΛ]>0"
                 End If
                 
                 
                 
                 
                 
               
               
' & " WHERE EGGTIM.APOT=" + left(APOTH1.Text, 1) + " AND ERG  LIKE '" + SameLetters(Trim(Mid(TIMText3.Text, 2, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) ;"
               
              ' ψαχνει κατά κωδικό  Εργοστασίου  KODERG %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            ElseIf Left(timText3.Text, 1) = "+" Then ' $$$$$$$
                 'EID.RecordSource = "select SUM(XRE)-SUM(PIS) AS [ΥΠΟΛ],EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO from EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD WHERE ERG  LIKE '" + Trim(Mid(timtext3.Text, 2, 14)) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) ;"
               
                 EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)" _
                 & "-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," _
                 & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                 & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" _
                 & " WHERE KODERG = '" + Trim(mID(timText3.Text, 2, 14)) + "' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1 ;"
               
' & " WHERE EGGTIM.APOT=" + left(APOTH1.Text, 1) + " AND ERG  LIKE '" + SameLetters(Trim(Mid(TIMText3.Text, 2, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) ;"
               
               
               
                
            '----------------kata kodiko otan einai gyrismeno sta barcode ---------------------------------
            ElseIf f_barcode = 1 And Left(timText3.Text, 1) = "/" Then ' $$$$$$  KATA BARCODE
                EID.RecordSource = "select SUM(CASE WHEN APOT<>" + Left(APOTH1.Text, 1) + " OR  XRE IS NULL THEN 0 ELSE XRE END)" _
                 & "-SUM(CASE WHEN  APOT<>" + Left(APOTH1.Text, 1) + " OR  PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," _
                 & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                 & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" _
                 & " WHERE   KOD  LIKE '" + SameLetters(Trim(mID(timText3.Text, 2, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1 ;"
            
            '----------------kata  barcode ---------------------------------
            ElseIf f_barcode = 1 Then ' $$$$$$  KATA BARCODE
               
               If F_EXO_XROMATA = 0 Then
                  EID.RecordSource = "select EID.KOD,EID.ONO,XTI,LTI,LTI5,FPA,MON,MEMO,POS_KERD,EID.ERG,EID.SYSKMAX from BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD WHERE BARCODES.ERG LIKE '" + SameLetters(Trim(mID(timText3.Text, 1, 13))) + "%';"
                  EID.Refresh
                  
                  'ψαχνω κατα βοηθητικο αν δεν βρεθηκε
                  If EID.Recordset.EOF Then
                         ' ψαχνει κατά BOHUHTIKO κωδικό  ERG %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                         EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)" _
                         & "-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," _
                         & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                         & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" _
                        & " WHERE ERG  LIKE '" + SameLetters(Trim(mID(timText3.Text, 1, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1 "
                        EID.Refresh
                        
               
                 If F_MONOMEYPOLOIPO = 1 And f_pol = 1 Then
                   EID.RecordSource = EID.RecordSource + " HAVING [ΥΠΟΛ]>0"
                 End If
                  
                  
                  
                  
                  End If
                
                  
               Else
                  EID.RecordSource = "select EID.KOD,EID.ONO,XTI,LTI,LTI5,FPA,MON,MEMO,POS_KERD,EID.ERG,EID.SYSKMAX,PROELEYSH,KATHGORIA,CH1 from BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD WHERE BARCODES.ERG LIKE '" + SameLetters(Trim(mID(timText3.Text, 1, 13))) + "%';"
                  EID.Refresh
                  XROMATA(Grid1.row, 1, 0) = Left(EID.Recordset("PROELEYSH"), 3)
                  XROMATA(Grid1.row, 0, 1) = mID(EID.Recordset("PROELEYSH"), 4, 2)
               End If
               
            '----------------kata  barcode  οταν είναι γυρισμένο κατα κωδικό ---------------------------------
            ElseIf Left(timText3.Text, 1) = "/" Then ' $$$$$$  KATA BARCODE
               EID.RecordSource = "select EID.KOD,EID.ONO,XTI,LTI,LTI5,FPA,MON,MEMO,POS_KERD,EID.ERG,EID.SYSKMAX,EID.KATHGORIA,CH1 from BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD WHERE BARCODES.ERG LIKE '" + SameLetters(Trim(mID(timText3.Text, 2, 14))) + "%';"
            
            '-------------------ΓΕΝΝΗΤΡΙΑ ΚΩΔΙΚΩΝ -------------------------------------------------
            ElseIf Left(timText3.Text, 1) = "%" Then   '$$$$$$$
               If fbuff = "r" Then '------------------------
                  Dim GEN As String
                  
                  If Len(Dir("C:\DIONISIS.EXE")) > 0 Then
                       GEN = "5201001%"
                       R0.Open "select max(ERG) FROM EID WHERE ERG LIKE '" + GEN + "'", Gdb, adOpenForwardOnly, adLockReadOnly
                       XX = Format(1 + Val(Left(R0(0), 12)), "000000000000")
                       R0.Close
                    APOT1.flag.Caption = "1"
                       APOT1.Text1(4).Text = XX + ChDigEAN13(XX) 'timtext3.Text
                                         
                  Else
                      GEN = "5221275%"
                      R0.Open "select max(KOD) FROM EID WHERE KOD LIKE '" + GEN + "'", Gdb, adOpenForwardOnly, adLockReadOnly
                      XX = Format(1 + Val(Left(R0(0), 12)), "000000000000")
                      R0.Close
                      APOT1.flag.Caption = "1"
                      APOT1.Text1(0).Text = XX + ChDigEAN13(XX) 'timtext3.Text
                  End If
                  
                  
                  
                  gApoMenu = False
                  APOT1.SHOW 1
                  EID.RecordSource = "select KOD,ONO,XTI,LTI,LTI5,FPA,MON,MEMO,POS_KERD,ERG from EID WHERE KOD LIKE '" + XX + "%';"
               End If              '------------------------
               
20          Else   ' $$$$$$$
                 
                 '----KATA KΩΔIKO ΓΥΡΙΣΜΕΝΟ ΣΤΟ ΚΑΤΑ ΚΩΔΙΚΟ
                EID.RecordSource = "select SUM(CASE WHEN APOT<>" + Left(APOTH1.Text, 1) + " OR  XRE IS NULL THEN 0 ELSE XRE END)" _
                 & "-SUM(CASE WHEN  APOT<>" + Left(APOTH1.Text, 1) + " OR  PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," _
                 & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                 & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" _
                 & " WHERE   KOD  LIKE '" + SameLetters(timText3.Text) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1 ;"
               

            End If   ' $$$$$$$
            EID.Refresh
            
              
              GridEidon.Columns(0).width = 500
           GridEidon.Columns(1).width = 1800
           GridEidon.Columns(2).width = 5000
             
             
             'ψάχνει κατά όνομα
         ElseIf Grid1.Col = f_o Then  '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            
             GridEidon.Columns(0).width = 500
             GridEidon.Columns(1).width = 5000
            'ME * NA MHN PSAXNEI KATA ONOMA
            If Left(Grid1.TextMatrix(Grid1.row, f_k), 1) = "*" Then ' ----------
                Grid1.Text = timText3.Text
                ' μιά κολονα παρακάτω
                If Grid1.Col < Grid1.ColS - 3 Then  ' πάει στην επόμενη κολόνα
                   Grid1.Col = Grid1.Col + 1
                   Grid1_EnterCell
                   Exit Sub
                End If
            Else  ' -------------
              ' SELECT SUM(XRE-PIS),KODE,EID.ONO FROM  EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD WHERE KODE LIKE '01%'  GROUP BY KODE,EID.ONO
              ' EID.RecordSource = "select SUM(XRE-PIS) AS [ΥΠΟΛ],EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO from EID LEFT JOIN EGGTIM ON EGGTIM.KODE=EID.KOD WHERE EID.ONO LIKE '" + UCase(timtext3.Text) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) ;"
               Dim ONOM As String
               Dim KODI As String
               Dim DOL As Long
               'SYNDYASMENH ANAZHTHSH ME $  DHLADH   ΣΚΙΑ$99 ΟΛΕΣ ΤΙΣ "ΣΚΙΕΣ" ΜΕ ΚΩΔΙΚΟ ΠΟΥ ΑΡΧΙΖΕΙ ΑΠΟ "99"
               ONOM = timText3.Text
               DOL = InStr(1, ONOM, "$")
               KODI = ""
               If DOL > 1 Then
                  KODI = Trim(mID(ONOM, DOL + 1, 14))
                  ONOM = Left(ONOM, DOL - 1)
                  
                 EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)" _
                 & "-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ]," _
                 & " EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                 & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" _
                 & " WHERE   EID.ONO  LIKE '" + SameLetters(ONOM) + "%' AND KOD LIKE '" + KODI + "%' and EGGTIM.APOT=" + Left(APOTH1.Text, 1) + "  GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1 ;"
               Else
                 
                 EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL OR APOT<>" + Left(APOTH1.Text, 1) + " THEN 0 ELSE XRE END)" _
                 & "-SUM(CASE WHEN PIS IS NULL OR APOT<>" + Left(APOTH1.Text, 1) + " THEN 0 ELSE PIS END) AS [ΥΠΟΛ]," _
                 & " EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                 & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" _
                 & " WHERE   EID.ONO  LIKE '" + SameLetters(ONOM) + "%'" + "  GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1 "
                 ' and EGGTIM.APOT=" + left(APOTH1.Text, 1)
                 
                 
                 If F_MONOMEYPOLOIPO = 1 And f_pol = 1 Then
                   EID.RecordSource = EID.RecordSource + " HAVING SUM(CASE WHEN XRE IS NULL OR APOT<>" + Left(APOTH1.Text, 1) + " THEN 0 ELSE XRE END)" _
                   & "-SUM(CASE WHEN PIS IS NULL OR APOT<>" + Left(APOTH1.Text, 1) + " THEN 0 ELSE PIS END)>0"
                 End If

                
                
                End If
                
              EID.Refresh
              
       GridEidon.width = 10000
              GridEidon.Columns(0).width = 1000
             GridEidon.Columns(1).width = 5000
              GridEidon.Columns(3).width = 1000
             
             GridEidon.Columns(2).width = 1000
              

              
              
              
              
              
              
              
              
              
              
              If EID.Recordset.EOF Then
              
                 EID.RecordSource = "select EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX " _
                  & " from EID " _
                  & " WHERE   EID.ONO  LIKE '" + SameLetters(timText3.Text) + "%' ORDER BY ONO ;"
                  EID.Refresh
              
              
              
              End If
              If EID.Recordset.RecordCount = 0 Then   ' $$$$$$$$$$$$$$$$$$$$$$$$$$
                ' μιά κολονα παρακάτω
                If Grid1.Col < Grid1.ColS - 3 Then  ' πάει στην επόμενη κολόνα
                   Grid1.Col = Grid1.Col + 1
                   Grid1_EnterCell
                   Exit Sub
                End If
              End If   ' $$$$$$$$$$$$$$$$$$$$$$$$$$
            
            End If '----------
            
             
         End If   '        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
         
         If EID.Recordset.EOF Then   '%%%%%%%%%%%%%
            
            If fbuff = "r" Then  ' $$$$$$$$$$$$$$$
               APOT1.Text1(0).Text = timText3.Text
               APOT1.flag.Caption = "1"
               APOT1.SHOW
            Else   ' $$$$$$$$$$$
            
                ' If EID.Recordset.RecordCount = 0 Then
                ' μιά κολονα παρακάτω
                 If Grid1.Col < Grid1.ColS - 3 Then  ' πάει στην επόμενη κολόνα
                    Grid1.Col = Grid1.Col + 1
                    Grid1_EnterCell
                    Exit Sub
                 End If
            End If   ' $$$$$$$$$$$
            
          Else  ' %%%%%%%%%%%%%%%%%%
            
            GridEidon.SetFocus

               If EID.Recordset.RecordCount = 1 Then
                       If f_AutoChangeLine = 1 Then
                            GridEidon.SetFocus
                            GridEidon_KeyUp 13, 0
                            SendKeys "{enter}{enter}"
                       Else
                            GridEidon.SetFocus
                            GridEidon_KeyUp 13, 0
                       
                       End If
               End If
            
            
            
            Me.Caption = "127"
          End If  ' %%%%%%%%%%%%%%%%%
          
       intRecord = 10 ' EID.Recordset.RecordCount
       intField = EID.Recordset.FIELDS.Count
       'call the procedure here...
       
       Call AdjustDataGridColumns(intRecord, intField, True)
          
     
     
     ' kod ή όνομα <2 τότε ψάχνουμε
     '******************************************************************************
30   Else ' ************************************************************************
         
               ' EKTOS KOD ,ONO
         
         
         
         
         'ΥΠΟΛΟΓΊΖΕΙ ΑΞΊΑ
         If Grid1.Col = f_p Then  ' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            Dim R As New ADODB.Recordset, mXre As Single, mPIS As Single, mDesmia As Single
            
            
            If f_use_Sysk = 1 Then
               If EID.Recordset("SYSKMAX") > 0 Then
                  timText3.Text = timText3.Text * EID.Recordset("SYSKMAX")
               End If
            End If
            
            If (fbuff = "e" And (F_POS_APOU = 2 Or F_POS_APOU = 3) And f_checkYpol > 0 And F_STADIO <> 2) Or (F_PAR_MONO_ME_YPOLOIPO = 1 And F_STADIO = 2) Then ' F_STADIO<>2 ΔΕΝ ΕΙΝΑΙ ΠΑΡΑΓΓΕΛΙΑ
               ' 13-4-2006 R.Open "SELECT SUM(XRE),SUM(PIS) FROM EGGTIM WHERE KODE='" + EID.Recordset("KOD") + "' ", Gdb, adOpenForwardOnly, adLockReadOnly
               R.Open "SELECT SUM(XRE),SUM(PIS) FROM EGGTIM WHERE KODE='" + Grid1.TextMatrix(Grid1.row, 1) + "' ", Gdb, adOpenForwardOnly, adLockReadOnly
               
               If R.EOF Then
               Else
                  mXre = IIf(IsNull(R(0)), 0, R(0)): mPIS = IIf(IsNull(R(1)), 0, R(1))
               End If
               R.Close
               
               'ποσοτητα παραγγελιών για δέσμευση
               
               If IsEmpty(F_SHMADEMENA_DELTIA) Or Len(Trim(F_SHMADEMENA_DELTIA)) <= 2 Then
                    F_SHMADEMENA_DELTIA = "''"
                   'ΣΑΝ ΔΕΣΜΕΥΜΕΝΑ ΝΟΟΥΝΤΑΙ ΟΛΑ ΤΩΝ ΠΑΡΑΓΓΕΛΙΩΝ ΠΛΗΝ ΑΥΤΟΥ ΤΟΥ ΠΕΛΑΤΗ
                   Sql2 = "SELECT SUM(POSO) FROM EGGTIM left JOIN TIM ON EGGTIM.ATIM=TIM.ATIM AND EGGTIM.HME=TIM.HME  WHERE" _
                   & " (TIM.ART is null)  AND LEFT(EGGTIM.ATIM,1)='a' and  KODE='" + Grid1.TextMatrix(Grid1.row, 1) + "'  and PELKOD<>'" + Text2(0).Text + "'"
               Else
                   'ΣΑΝ ΔΕΣΜΕΥΜΕΝΑ ΝΟΟΥΝΤΑΙ ΟΛΑ ΤΩΝ ΠΑΡΑΓΓΕΛΙΩΝ ΠΛΗΝ ΤΩΝ ΠΑΡΑΣΤΑΤΙΚΩΝ ΠΟΥ ΤΣΙΜΠΗΣΑ
                   Sql2 = "SELECT SUM(POSO) FROM EGGTIM left JOIN TIM ON EGGTIM.ATIM=TIM.ATIM AND EGGTIM.HME=TIM.HME  WHERE" _
                   & " (   NOT EGGTIM.ATIM+CONVERT(CHAR(10),EGGTIM.HME,103)   IN (" + F_SHMADEMENA_DELTIA + ") ) AND (TIM.ART is null)  AND LEFT(EGGTIM.ATIM,1)='a' and  KODE='" + Grid1.TextMatrix(Grid1.row, 1) + "' "
               End If
               
               R.Open Sql2, Gdb, adOpenForwardOnly, adLockReadOnly
               If R.EOF Or f_Desmia = 0 Then ' ΔΕΝ ΒΡΗΚΕ ή ΔΕΝ ΚΑΝΕΙ ΕΛΕΓΧΟ
                  mDesmia = 0
               Else
                  mDesmia = IIf(IsNull(R(0)), 0, R(0))
               End If
               
               
               'ελεγχος αν αυτά που στέλνω είναι μέσα σε αυτά που παρήγγειλε o πελάτης
               If f_mono_apoParaggelia = 1 And F_STADIO <> 2 Then 'ΔΕΝ ΕΙΝΑΙ ΠΑΡΑΓΓΕΛΙΑ
                  
                    On Error GoTo MHNYMA
                  

                    HDH_FORTOMENA = 0
                    For LK0 = 1 To Grid1.ROWS - 1
                        If LK0 <> Grid1.row And Grid1.TextMatrix(LK0, 1) = Grid1.TextMatrix(Grid1.row, 1) Then
                              HDH_FORTOMENA = HDH_FORTOMENA + Grid1.TextMatrix(LK0, 3)
                        End If
                    Next
                    '    Sql2 = "SELECT SUM(POSO) FROM EGGTIM left JOIN TIM ON EGGTIM.ATIM=TIM.ATIM AND EGGTIM.HME=TIM.HME  WHERE" _
                    '    & " (   NOT EGGTIM.ATIM+CONVERT(CHAR(10),EGGTIM.HME,103)   IN (" + F_SHMADEMENA_DELTIA + ") ) AND (TIM.ART is null)  AND LEFT(EGGTIM.ATIM,1)='a' and  KODE='" + Grid1.TextMatrix(Grid1.row, 1) + "' and PELKOD='" + Text2(0).Text + "'"
                   
                   
                   
                   'Else
                       Sql2 = "SELECT SUM(POSO) FROM EGGTIM left JOIN TIM ON EGGTIM.ATIM=TIM.ATIM AND EGGTIM.HME=TIM.HME  WHERE" _
                       & " (   NOT EGGTIM.ATIM+CONVERT(CHAR(10),EGGTIM.HME,103)   IN (" + F_SHMADEMENA_DELTIA + ") ) AND (TIM.ART is null)  AND LEFT(EGGTIM.ATIM,1)='a' and  KODE='" + Grid1.TextMatrix(Grid1.row, 1) + "' and PELKOD='" + Text2(0).Text + "'"
                   'End If
                   
                   R.Close
                   R.Open Sql2, Gdb, adOpenForwardOnly, adLockReadOnly
                   If R(0) = 0 Or IsNull(R(0)) Then
                   
                    MHNYMA2.Timer1.Interval = 50
                    MHNYMA2.SHOW 1
                   
                        'Dim ans As Integer
                        MsgBox "ΔΕΝ ΕΚΚΡΕΜΕΙ ΤΕΤΟΙΟ ΕΙΔΟΣ ΣΤΟΝ ΠΕΛΑΤΗ"
                        timText3.Text = 0
                   End If
                   
                   If HDH_FORTOMENA + Val(timText3.Text) > R(0) Then
                      MHNYMA2.Timer1.Interval = 50
                      MHNYMA2.SHOW 1
                   
                      MsgBox "Η παραγγελθείσα ποσότητα είναι " + Str(R(0)) + Chr(13) + "και έχετε ήδη τιμολογήσει " + Format(HDH_FORTOMENA + Val(timText3.Text), "#####.00")
                      timText3.Text = 0
                   End If
                   
                   
               End If
               
               
               
               
               
               If Val(timText3.Text) > (mXre - mPIS) - mDesmia Then
                    MHNYMA2.Timer1.Interval = 50
                      MHNYMA2.SHOW 1
                  
                  MsgBox "υπάρχουσα ποσότητα " + Format(mXre - mPIS, "###0") + " ΔΕΣΜ:" + Format(mDesmia, "####")
                  
                  If f_checkYpol = 2 Then
                     timText3.Text = 0
                  End If
               End If
            End If
            If F_EXO_XROMATA = 1 Then ' ΒΑΖΩ ΤΗΝ ΠΟΣΟΤΗΤΑ ΣΤΟΝ ΠΙΝΑΚΑ ΓΙΑ ΝΑ ΤΟΝ ΤΡΑΒΗΞΕΙ ΣΤΟ ENHM_EGGTIM
                Dim L9 As Long, M9 As Long
                Dim SUMA9 As Single
                SUMA9 = 0
                For L9 = 1 To 12
                   For M9 = 1 To 12
                      SUMA9 = SUMA9 + Val(XROMATA(Grid1.row, L9, M9))
                   Next
                Next
                If SUMA9 = 0 Then
                   XROMATA(Grid1.row, 1, 1) = timText3.Text
                End If
            End If
         End If     ' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        mText = timText3.Text
        KOMA = InStr(mText, ","): If KOMA > 0 Then Mid$(mText, KOMA, 1) = "."
        ' μετακίνηση σε grid
        
        Grid1.Text = mText  ' κρατάει το grid το κείμενο
        
        
         If Grid1.Col = f_p Or Grid1.Col = F_T Or Grid1.Col = f_e Or Grid1.Col = f_f Then
             FindSynolo
         End If
        
        
        ' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        ' f_CanChangeEkptosi
      
        

40      If Grid1.Col <= Grid1.ColS - 3 Then  ' πάει στην επόμενη κολόνα
           Grid1.Col = Grid1.Col + 1
           'εαν δεν επιτρέπεται τότε ξανά στην επόμενη
           If Grid1.Col = f_e And f_CanChangeEkptosi = 0 Then 'EKPTOSI
               Grid1.Col = Grid1.Col + 1
           End If
           If Grid1.Col = f_f And f_CanChangeFpa = 0 Then 'FPA
           
           
           
                 HDH_FORTOMENA = 0
                 
              '  If f_mono_apoParaggelia = 1 Then
              If f_SYMPYKNVSH = 1 Then
                 For LK0 = 1 To Grid1.row - 1
                    'brika idion kodiko
                    If Grid1.TextMatrix(LK0, 1) = Grid1.TextMatrix(Grid1.row, 1) Then
                          Grid1.TextMatrix(LK0, 3) = Val(Grid1.TextMatrix(LK0, 3)) + Val(Grid1.TextMatrix(Grid1.row, 3))
                           
                          HDH_FORTOMENA = 1
                    End If
                  Next
              End If
            
              If HDH_FORTOMENA = 0 Then
                 Grid1.row = Grid1.row + 1
              Else
                 Grid1.TextMatrix(Grid1.row, 3) = ""
                 Grid1.TextMatrix(Grid1.row, 1) = ""
                 Grid1.TextMatrix(Grid1.row, 2) = ""
                 Grid1.TextMatrix(Grid1.row, 4) = ""
                 Grid1.TextMatrix(Grid1.row, 5) = ""
                 
              End If
              
              
              Grid1.Col = 1
           End If
           Me.Caption = "125"
        End If
        
        
        
        
        
        If Grid1.Col > Grid1.ColS - 3 Then  ' πάει στην επόμενη κολόνα
           'πάει στην παρακάτω σειρά
           
           
           
           
           
           
           
           
           If Grid1.row < Grid1.ROWS - 1 And Grid1.row < f_SeiresTimologioy Then
              Grid1.row = Grid1.row + 1
              Grid1.Col = 1
               Me.Caption = "124"
           End If
            Me.Caption = "123"
        End If
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        
        Grid1_EnterCell
       
       
     End If   ' ************************************************
     
End If ' KEYCODE=13 -------------------------------------------------------

 If KeyCode = 116 And Shift = 0 Then ' f5 teleytaia timh
     Me.Caption = 30
     R0.Open "select TOP 1 TIMM,EKPT  from EGGTIM where PELKOD='" + Text2(0) + "' and KODE='" + Grid1.TextMatrix(Grid1.row, 1) + "' ORDER BY HME DESC", Gdb, adOpenForwardOnly, adLockReadOnly
     If Not R0.EOF Then
        'r0.MoveLast
        Me.Caption = "ΤΕΛ.ΤΙΜΗ:" + Format(R0(0), "###0.00") + " EΚΤΠΩΣΗ :" + Format(R0(1), "###0.00")
        Grid1.TextMatrix(Grid1.row, F_T) = Format(R0(0), "###0.00")
        Grid1.TextMatrix(Grid1.row, f_e) = Format(R0(1), "###0.00")
        If Grid1.Col = F_T Then
           timText3.Text = Grid1.TextMatrix(Grid1.row, F_T)
        End If
        If Grid1.Col = f_e Then
           timText3.Text = Grid1.TextMatrix(Grid1.row, f_e)
        End If
     End If
 End If
 
50 If KeyCode = 38 Then   ' ΠΑΝΩ ΒΕΛΟΣ
      
      Grid1.Text = timText3.Text ' κρατάει το grid το κείμενο
      
      ' πάει μία κολόνα αριστερά
      If Grid1.Col > 1 Then
         Grid1.Col = Grid1.Col - 1
      Else
        'π'αει μία σειρά επάνω
        If Grid1.row > 1 Then
           Grid1.row = Grid1.row - 1
           Grid1.Col = 1
        End If
      End If
      Grid1_EnterCell
   End If
   
 If KeyCode = 120 And Shift = 0 Then ' f9 ypologizo τιμη χορισ φπα
     KOMA = 1
     Dim DD As Single, SYNTF As Single
     
     DD = Val(Grid1.TextMatrix(Grid1.row, F_T))
     SYNTF = (100 + g_Fpa(2)) / 100
     If DD > 0 Then
       'timtext3.Text = Format(DD / (100 + g_Fpa(2)) / 100, "###0.000")
       Grid1.TextMatrix(Grid1.row, F_T) = Round(DD / SYNTF, 2)
     End If
   
 End If
   
   
 If KeyCode = 119 And Shift = 0 Then ' f8 ypologizo posotita apo ajia
    Dim m_aj, MR
    
60  m_aj = InputBox("Δώσε αξία ", Str(Grid1.row))
    
    If Val(Grid1.TextMatrix(Grid1.row, F_T)) > 0 Then
      Grid1.TextMatrix(Grid1.row, f_p) = Format(Val(m_aj) / Val(Grid1.TextMatrix(Grid1.row, F_T)) / (1 + g_Fpa(Val(Grid1.TextMatrix(Grid1.row, f_f))) / 100), "###0.00000")
     ' Me.Caption = Val(m_aj) / Val(Grid1.TextMatrix(Grid1.row, F_T))
      
     ' FindSynolo
    End If
 
 End If
   
70
   If KeyCode = 40 Then   ' KATΩ ΒΕΛΟΣ
      Grid1.Text = timText3.Text ' κρατάει το grid το κείμενο
      ' μιά σειρά παρακάτω
      If Grid1.row < Grid1.ROWS - 1 Then
         FindSynolo
         Grid1.row = Grid1.row + 1
               Grid1_EnterCell
      End If

   End If
   

'  If KeyCode = 27 Then 'esc τελος εισαγωγής
'       KeyCode = 0
'       Combo4.SetFocus
'  End If
Exit Sub

MHNYMA:
HandleError "Par1:timtext3_KeyUp"
Resume Next

End Sub





Private Static Sub FindSynolo()
   Dim C, R, P, t, e, K, S, aj, f, SUMPOS
On Error GoTo MHNYMA

            For K = 1 To 9: SumFpa(K) = 0: Next
            
            R = Grid1.row
            C = Grid1.Col  ' απομνημονεύω την τρέχουσα στήλη kai σειρά
            
10          Grid1.Col = f_p: P = Grid1.Text
            Grid1.Col = F_T: t = Grid1.Text
            Grid1.Col = f_e: e = Grid1.Text
            Grid1.Col = f_f: f = Grid1.Text
            
            
            
            'αξια με ΦΠΑ
            Grid1.Col = F_mefpa: aj = Round(Val(P) * Val(t) * (100 - Val(e)) / 100 * (100 + g_Fpa(Val(Grid1.TextMatrix(R, f_f)))) / 100, f_psifiaAjias)
            Grid1.Text = Replace(Format(aj, "####0.000"), ",", ".")
            
            
            
            Grid1.Col = f_a: aj = Round(Val(P) * Val(t) * (100 - Val(e)) / 100, f_psifiaAjias)
            Grid1.Text = Replace(Format(aj, "####0.000"), ",", ".")
            
            
            
            f_sajia(Grid1.row) = aj
            f_fpa(Grid1.row) = f  ' ΚΑΤΗΓΟΡΊΑ ΦΠΑ
20            S = 0
              SUMPOS = 0
            For K = 1 To Grid1.ROWS - 1
              If Len(Grid1.TextMatrix(K, f_k)) > 0 Then
                f_sajia(K) = Val(Grid1.TextMatrix(K, f_a)) '  Val(Grid1.TextMatrix(K, f_p)) * Val(Grid1.TextMatrix(K, F_T)) '
                
                SUMPOS = SUMPOS + Val(Grid1.TextMatrix(K, f_p))
                S = S + f_sajia(K)
                ' ΣΟΥΜΆΡΕΙ ΣΕ ΚΆΘΕ ΚΑΤΗΓΟΡΊΑ φπα ΤΗΝ ΑΞΙΑ
30                If f_fpa(K) < 9 Then
                      SumFpa(f_fpa(K)) = SumFpa(f_fpa(K)) + f_sajia(K)
                  End If
              End If
                  
            Next
            
            Label15.Caption = SUMPOS
            S = Round(S, f_psifiaAjias)
            Synolo.Caption = Format(S, "####0.00")
            f = 0: For K = 1 To 9: f = f + SumFpa(K) * g_Fpa(K) / 100: Next
            
            f = Round(f, 2)
40            sfpa.Caption = Format(f, "####0.00")
            SynoloG.Caption = Format(f + S, "####0.00")
            Grid1.Col = C ' το επαναφέρω εκεί που το βρήκα
            Grid1.row = R
            
            Grid1.Refresh
            
Exit Sub

MHNYMA:
HandleError "Par1:FINDSYNOLO"
Resume Next
            
            
            
            
End Sub
Public Function ExecuteL(Scode As String, Optional fcheckonly As Boolean) As Boolean
  Dim a$
  'ExecuteL = EbExecuteLine(StrPtr(Scode + ":a$=gvar"), 0&, 0&, Abs(fcheckonly)) = 0
  
'  παράδειγμα χρησιμοποίησης μακροεντολής
'  --------------------------------------
  
'   Private Sub Command1_Click()
'   Dim res As Boolean, var As Single
  a$ = "GVAR=1"
  a$ = Left(a$, 2)
  
 ' ExecuteL = EbExecuteLine(StrPtr(A$), 0&, 0&, Abs(fcheckonly)) = 0

'
'   res = ExecuteLine("var=2+3*(5+6):a$=var")
'
'   End Sub
End Function
Function Gefyroma(g_hme, g_atim, g_eidoskpe, gfkod, gfaj, m_r, mgre)
'εδώ δημιουργούνται οι εγγραφές για το dirlog\temp.dbf
'και ενημερώνονται οι ανώτεροι λογαριασμοί στο abat και ανοίγονται αν δεν υπάρχουν
'*********************** γεφυρωμα ****************************
Dim xre_pis(20), xre_Pis_rec(20), xreosis(20), GetList(20)
Dim DUM

Dim JJ, JJK, m_ait As String
Dim kodiklog(9), typ(9), xp(9), kau(9), a As String, mnq As String, eid_par As String
' Dim mgre As String
Dim K As Integer

Dim atr_xre, atr_pis, LW, DIAFORETIKA
atr_xre = 0: atr_pis = 0
Dim ok_gefyr
Dim db As Database, dbL As Database
Dim r2 As Recordset, gefyres As New ADODB.Recordset
Dim bbs As String
Dim anoigma As Integer
' Dim PEL As Recordset
Dim kod_pel As String
Dim mbat_pel As String
' Dim abat2 As Recordset
Dim LO As Integer
Dim JJ_ROBLEM
Dim mmhkos, MHKOS(5) As Integer

Dim MHNAS
Dim sql, K2 As Integer, logar(5) As Integer


On Error GoTo MHNYMA











'If gConnect = "Access" Then
'   Set db = OpenDatabase(gDir, False, False)
'   Set dbL = OpenDatabase(gdirlog, False, False)
'Else
'   Set db = OpenDatabase(gDir, False, False, gConnect)
'   Set dbL = OpenDatabase(gdirlog, False, False, gConnect)
'End If


Set db = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
Set dbL = OpenDatabase(gdirlog, False, False, "dBASE III;")





Dim par As Recordset
'ΒΡΙΣΚΩ ΑΠΟ ΤΙΣ GEFYRES.DBF ΤΗΝ ΚΑΤΑΛΛΗΛΗ ΕΓΓΡΑΦΗ
'Set par = db.OpenRecordset("SELECT *FROM PARASTAT WHERE EIDOS='" + left(g_atim, 1) + "';")

444


gefyres.Open "SELECT * FROM GEFYRES", Gdb, adOpenDynamic, adLockBatchOptimistic

'Set gefyres = db.OpenRecordset("GEFYRES")
'Gefyres.Move par("GEF_P") - 1  ' RECORD ΓΕΦΥΡΏΣΕΩΝ ΕΠΙ ΠΙΣΤΩΣΕΙ
gefyres.Move m_r - 1  ' RECORD ΓΕΦΥΡΏΣΕΩΝ ΕΠΙ ΠΙΣΤΩΣΕΙ


  ok_gefyr = 1
  'a = ctim(0)
  If g_hme <= ghme_orist Then
     
     MsgBox " ΔΕΝ ΕΠΙΤΡΕΠΟΝΤΑΙ ΚΙΝΗΣΕΙΣ  ΣΕ ΟΡΙΣΤΙΚΟΠΟΙΗΜΕΝΕΣ ΕΓΓΡΑΦΕΣ"
     Exit Function
  End If
'
'  GoTo 444
  If InStr("oOοΟ", gefyres("gefyrono")) Then
     ok_gefyr = 0
     Exit Function
  End If
  Dim EIDPAR
  
  
  
  eid_par = "  "
  If gefyres("mEID_PAR") <> "  " Then
     eid_par = Left(gefyres("mEID_PAR"), 5) + "-" + g_atim ' Τ.πωλ-Τ00012Γ
  End If
  
'  If Gefyres("mmGRE") <> " " Then
'     mgre = Gefyres("mmGRE") ' Σ
'  End If
  m_ait = gefyres("parathrhs") 'π.χ.  ΠΩΛΗΣΕΙΣ
'
'
'
  For K = 1 To 9
      a = LTrim(Str(K))
      xp(K) = gefyres("xpi" + a)
      kodiklog(K) = gefyres("Vkod" + a)  ' 3000****
      typ(K) = gefyres("Vtypos" + a)
  Next
  gefyres.Close
  
  
  
  
  
  
  
  
  '*** ψαχνω για διπλοεγγραφές
  gvar = "Select *from temp where eidhme='1' and hme=#"
  gvar = gvar + Format(g_hme, "mm/dd/yyyy") + "# and apa='"
  gvar = gvar + eid_par + "'"

  
  'EID.DatabaseName = gdirlog
  'EID.RecordSource = gvar
  
  'EID.Refresh

 '  Set temp = dbL.OpenRecordset(gvar)
 
 temp.DatabaseName = gdirlog
 temp.Connect = "dBase III;"
 temp.RecordSource = gvar '  "select * from temp" ' gvar
 
 abat.DatabaseName = gdirlog
 abat.Connect = "dBase III;"
 
 abat2.DatabaseName = gdirlog
 abat2.Connect = "dBase III;"




temp.Refresh
If temp.Recordset.RecordCount > 0 Then
         MsgBox ("ΥΠΑΡΧΕΙ ΙΔΙΟΣ ΑΡΙΘΜΟΣ & ΗΜΕΡ/ΝΙΑ ΣΤΗΝ ΓΕΝ.ΛΟΓΙΣΤΙΚΗ ΜΕ ΧΡΕΩΣΗ " + Format(temp.Recordset("XRE"), "###,###,###,###"))
         MsgBox ("ΑΔΥΝΑΤΗ Η ΓΕΦΥΡΩΣΗ ΜΕ ΤΗΝ ΓΕΝΙΚΗ ΛΟΓΙΣΤΙΚΗ ")
         
         'temp.Close
         'abat.Close
         Exit Function
  End If













' Set gTim = db.OpenRecordset("select *from tim where atim='" + g_atim + "' and hme>=#" + Format(g_hme, "mm/dd/yyyy") + "#;")




Set fSCR = CreateObject("MSScriptControl.ScriptControl")



'cEGGTIM.RecordSource = "select *from eggtim"
'cEGGTIM.Refresh

fSCR.language = "vbscript"


'DOK.RecordSource = "select * from TIM where ATIM='" + g_atim + "' and HME>='" + Format(g_hme, "mm/dd/yyyy") + "';"


'DOK.ConnectionString = gConnect
'DOK.Refresh


' DOK.RecordSource = "select *from tim where atim='" + g_atim + "' and hme>='" + Format(g_hme, "mm/dd/yyyy") + "';"
DOK.RecordSource = "select * from TIM where ATIM='" + g_atim + "' and HME>='" + Format(g_hme, "mm/dd/yyyy") + "';"

'CTIM.DatabaseName = gDir
DOK.ConnectionString = gConnect
DOK.Refresh


fSCR.AddObject "cTIM", DOK ' ADO  dok einai to pragmatiko antikeimeno  ctim to antistoixo sthn VBSCRIPT


'fSCR.AddObject "DOK", DOK ' ADO

fSCR.AddObject "temp", temp
fSCR.AddObject "abat", abat
fSCR.AddObject "abat2", abat2

'fSCR.AddObject "cMEM", CMEM
fSCR.AddObject "mactext", mactext









Dim debug1
' εδώ παίρνει το vkod1,vtypos1 έως το vkod9,typos9
 For K = 1 To 9
       If kodiklog(K) <> " " Then
            '            sele TEMP
            
            
            debug1 = DOK.Recordset("aji")
            a = typ(K) '
            'dum = ExecuteLine("gvar=" + a)
            a = mac(a)
                    
            
          If Val(a) <> 0 Then    ' γγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγ
'
'             ' EAN EIMASTE STHN ANALYSH EIDVN
             If mID$(kodiklog(K), 3, 2) = "**" Then ' κωδικός λογιστικής
                 DIAFORETIKA = 20  ' για νά κάνει πολλά loops ένα για κάθε είδος ΑΠΟΘΗΚΗΣ
             Else
                 DIAFORETIKA = 1   ' περνάει μόνο μία φορά
             End If

             For LW = 1 To DIAFORETIKA   ' MONO STHN PERIPTOSH TON EIDVN=20 ALLIOS =1
1010
            '* αν δεν έχει ανάλυση να προχωράει
               If LW > DIAFORETIKA Then Exit For
               If gfaj(LW) = 0 And DIAFORETIKA > 1 Then
                  LW = LW + 1
                  GoTo 1010
               End If
'
'               sele TEMP
'
                  temp.Recordset.AddNew
                  'dum = ExecuteLine("
                  
                  
             'dum = ExecuteLine("gvar=" + Trim(kodiklog(K)))

             
             'dum = ExecuteLine("gvar=" + gvar)
             
             
             
             gvar = Trim(kodiklog(K))
             bbs = gvar
             If InStr("30-50", Left(bbs, 2)) > 0 And InStr(bbs, "***") > 0 Then ' κωδικός πελάτου
                      If InStr("30-50", Left(bbs, 2)) Then '
                         'sele Abat
                         anoigma = 0
                        '******** εχει ΚΩΔΙΚΟ ΛΟΓΙΣΤΙΚΗΣ *******************
                          kod_pel = Trim(PEL.Recordset("shm2"))

                         '**  έχει κωδικό λογιστικής
                           mbat_pel = "3"
                           If Len(Trim(PEL.Recordset("shm2"))) = gDIG_4 Then
                              mbat_pel = "4"
                           End If
                           If Len(kod_pel) = 0 Then
                              MsgBox " δεν έχω κωδικό λογιστικής " + Format(g_hme, "dd/mm/yyyy") + " " + g_atim + " " + g_eidoskpe
                              ok_gefyr = 0
                              'abat.Close
                              'temp.Close
                              Set fSCR = Nothing
                              Exit Function
                           End If
'                              seek mbat_pel + kod_pel
                              abat.RecordSource = "select *from abat where bat='" + mbat_pel + "' and kod='" + kod_pel + "'"
                              
                              abat.Refresh
                              If abat.Recordset.RecordCount > 0 Then
                                 anoigma = 1
                              Else
                                 ' αλλά δεν υπάρχει στην λογιστική
                                 '** βρίσκω τον τελευταίο κωδικό
                                 abat.Recordset.AddNew
                                 abat.Recordset("bat") = mbat_pel
                                 abat.Recordset("kod") = kod_pel
                                 abat.Recordset("ono") = PEL.Recordset("epo")
                                 abat.Recordset("afm") = PEL.Recordset("afm")
                                 kod_pel = abat.Recordset("kod")
                                 abat.Recordset.Update
                                 
                                 'PEL.Edit
                                 'PEL("SHM2") = abat.recordset("KOD")
                                 'anoigma = 1
                              End If         '  eof()
'
                      End If '  left(kodiklog(k),2) $ '30 50'
'
'
    '                      Temp.Edit
                         temp.Recordset("KOD") = kod_pel

                         'If DRX_TO_EURO = 1 Then
                         '    temp.recordset("XRE") = eround(a / 340.75, 2)
                         'Else
                             temp.Recordset("XRE") = a
                         'End If




               ElseIf mID$(bbs, 1, 4) = "****" Then ' κωδικός λογιστικής ΑΠΟΘΗΚΗΣ
                      temp.Recordset("KOD") = Trim(gfkod(LW))
                      temp.Recordset("XRE") = gfaj(LW)  ' ΤΟ ΔΙΑΧΩΡΙΖΩ ΠΑΡΑΠΑΝΩ ΔΡΧ("ΕΥΡΩ

               Else
               
             
               
               
               
               
               
                   temp.Recordset("kod") = bbs
                   temp.Recordset("XRE") = a

               End If


               temp.Recordset("APA") = eid_par
               temp.Recordset("EIDHME") = "1"
               temp.Recordset("AIT") = m_ait
               temp.Recordset("HME") = g_hme
               temp.Recordset("GRE") = mgre
               temp.Recordset("XPI") = xp(K)
               temp.Recordset("PROOD") = 0
               temp.Recordset("ARIU") = 0
               temp.Recordset("SYGK") = ""
               JJ = Len(Trim(temp.Recordset("kod")))
               temp.Recordset.Update
               
 
              Next     ' LW =1 TO DIAFORETIKA

            End If       '     IF &a > 0 γγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγ
         End If ' if kodiklog(k) <> ' '
'
'
     Next  ' K = 1 To 9
      
      
      
      If Abs(atr_xre - atr_pis) >= 0.01 Or JJ_ROBLEM = 1 Then
     Else
        temp.Recordset.MoveFirst
       
       FlxGridDeltion.Visible = True
       ' ενημερώνω τα αρχείο λογαριασμών abat
       Do While Not temp.Recordset.EOF ' and  eidhme+dtos(hme)+left(gre,1)+apa =  '1' + dtos(g_hme) + mgre  + EID_PAR
           f_mc = f_mc + 1 '
           FlxGridDeltion.TextMatrix(f_mc, 0) = temp.Recordset(0)
           FlxGridDeltion.TextMatrix(f_mc, 1) = temp.Recordset("apa")
           FlxGridDeltion.TextMatrix(f_mc, 2) = temp.Recordset("ait")
           If temp.Recordset("xpi") = "Π" Then FlxGridDeltion.TextMatrix(f_mc, 4) = Format(temp.Recordset("xre"), "######0.00")
           If temp.Recordset("xpi") = "Χ" Then FlxGridDeltion.TextMatrix(f_mc, 3) = Format(temp.Recordset("xre"), "######0.00")
               
               
               
               mmhkos = Len(Trim(temp.Recordset("kod")))
               If temp.Recordset("eidHME") = "3" Then ' απογραφή
                   MHNAS = "00"
               Else
                   MHNAS = find_mhnas(temp.Recordset("hme"))
               End If
                sql = "select *from abat WHERE "
                     
                      If mmhkos = gDIG_4 Then sql = sql + "(bat='4' and left(kod," + Str(gDIG_4) + ")='" + temp.Recordset("kod") + "') or"
                     If mmhkos >= gDIG_3 Then sql = sql + "(bat='3' and left(kod," + Str(gDIG_3) + ")='" + Left(temp.Recordset("kod"), gDIG_3) + "') or "
                          If mmhkos >= 4 Then sql = sql + "(bat='2' and left(kod,4)='" + Left(temp.Recordset("kod"), 4) + "') or "
                          If mmhkos >= 2 Then sql = sql + "(bat='1' and left(kod,2)='" + Left(temp.Recordset("kod"), 2) + "') "
                      
                      'συγκεντρώνω τους λογαριασμους που με ενδιαφέρουν
                      abat2.RecordSource = sql
                      ' εαν λειπουν τους ανοιγω
                      MHKOS(1) = 2: MHKOS(2) = 4: MHKOS(3) = gDIG_3: MHKOS(4) = gDIG_4
                      If mmhkos = 2 Then LO = 1
                      If mmhkos = 4 Then LO = 2
                      If mmhkos = gDIG_3 Then LO = 3
                      If mmhkos = gDIG_4 Then LO = 4
                      abat2.Refresh
                      If abat2.Recordset.RecordCount < LO Then
                          For K2 = 1 To LO: logar(K2) = 0: Next
                          On Error Resume Next
                          abat2.Recordset.MoveFirst
                          Do While Not abat2.Recordset.EOF
                            logar(Val(abat2.Recordset("bat"))) = 1
                            abat2.Recordset.MoveNext
                            DoEvents
                          Loop
                          For K2 = 1 To LO
                             If logar(K2) = 0 Then
                                abat2.Recordset.AddNew
                                abat2.Recordset("bat") = LTrim(Str(K2))
                                abat2.Recordset("kod") = Left(temp.Recordset("kod"), MHKOS(K2))
                                abat2.Recordset.Update
                             End If
                          Next
                      End If
                      abat2.Recordset.MoveFirst
                      Do While Not abat2.Recordset.EOF
                            abat2.Recordset.Edit
                            If temp.Recordset("xpi") = "Π" Then
                               If IsNull(abat2.Recordset("ppros" + MHNAS)) Then abat2.Recordset("ppros" + MHNAS) = temp.Recordset("xre") Else abat2.Recordset("ppros" + MHNAS) = abat2.Recordset("ppros" + MHNAS) + temp.Recordset("xre")
                            Else
                               If IsNull(abat2.Recordset("ppros" + MHNAS)) Then abat2.Recordset("xpros" + MHNAS) = temp.Recordset("xre") Else abat2.Recordset("xpros" + MHNAS) = abat2.Recordset("xpros" + MHNAS) + temp.Recordset("xre")
                            End If
                            abat2.Recordset.Update
                            abat2.Recordset.MoveNext
                       Loop
                       abat2.Recordset.Close
               temp.Recordset.MoveNext
       Loop
     End If
      
          
      
Set fSCR = Nothing
      
      Exit Function
               


MHNYMA:
HandleError "Par1:gefyroma"
Resume Next
                                 
                     




End Function




Function printCrystal(MATIM, char_date)
'------------------------------------------------------------------------
'δημιουργεί τα αρχεία TIM  , EGGTIM , MEM    ΣΤΟ  c:\mercvb\reports\reports.mdb"
'ΓΙΑ ΤΗΝ ΕΚΤΥΠΩΣΗ ΤΟΥ ΤΙΜΟΛΟΓΙΟΥ

Dim sql As String, fDB As Database, d1 As Date, d2 As Date
Dim synt As String, db As Database
Dim DUM
Dim REGGTIM As New ADODB.Recordset
Dim rtim As New ADODB.Recordset
Dim R As Recordset
Dim K As Integer
Dim kn As Integer
Dim R3 As New ADODB.Recordset

'-----------  next lines are OK --------------------------

On Error GoTo MHNYMA

printCrystal = 1 ' αν ειναι dum < -900 τοτε δεν τυπωνει
If Len(f_dior_mAtim) > 0 Then
   d1 = Now
   d2 = DateAdd("d", 1, d1)
Else
   d1 = DTPicker1.value
   d2 = DateAdd("d", 1, d1)
End If

Set db = OpenDatabase("c:\mercvb\reports\reports.mdb")

 ADD_FIELDDAO "TIM", "AJ6", "SINGLE", db
 ADD_FIELDDAO "TIM", "FPA6", "SINGLE", db
 ADD_FIELDDAO "TIM", "DEH", "CHAR(12)", db
 ADD_FIELDDAO "TIM", "C1", "CHAR(30)", db
 ADD_FIELDDAO "TIM", "C2", "CHAR(30)", db


'If gConnect = "Access" Then
'  Set fDB = OpenDatabase(gDir, False, False)
'
'Else
'  Set fDB = OpenDatabase(gDir, False, False, gConnect)
'End If
'                  arn_enhm_apot

'=====   EGGTIM  ======


'On Error Resume Next
db.Execute "delete *from eggtim"
10




'TO   round(TIMM*POSO*(100-EKPT)/100,2)  ANTIKATASTAUHKE ME TO KAU_AJIA  4-1-2006


'sql = sql + "MONA,EKPT "
'sql = sql + " from EGGTIM where APOT=" + Str(Val(left(APOTH1.Text, 2))) + " and  left(ATIM," + Str(F_PSIFIA_ATIM + 1) + ")='" + left(mAtim, F_PSIFIA_ATIM + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "' ORDER BY PROOD_AJ;"
'REGGTIM.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

'copy to access table eggtim
'Set R = db.OpenRecordset("eggtim")
'REGGTIM.MoveFirst
'Do While Not REGGTIM.EOF
  
 '  R.AddNew
  ' For K = 0 To REGGTIM.FIELDS.Count - 1
   '  R(K) = REGGTIM(K)
   'Next
   'R.Update
   'REGGTIM.MoveNext
'Loop

sql = "select FPA,ATIM,HME as [Ημερομ],KODE as [Kωδ_Είδ],ONOMA as [Περιγραφή],"


If F_XONDR = 0 Then 'LIANIKH
   sql = sql + "POSO as [Ποσότητα],TIMM as [Tιμ_Μον],MIK_AJIA as [Αξία],"
Else
   sql = sql + "POSO as [Ποσότητα],TIMM as [Tιμ_Μον],KAU_AJIA as [Αξία],"
End If
sql = sql + "MONA,EKPT "

sql = sql + " from EGGTIM where APOT=" + Str(Val(Left(APOTH1.Text, 2))) + " and  left(ATIM," + Str(f_psifia_atim + 1) + ")='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "' ORDER BY PROOD_AJ;"


'Gdb.Execute SQL
'GoTo 10
REGGTIM.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

'copy to access table eggtim
Set R = db.OpenRecordset("eggtim")
REGGTIM.MoveFirst

On Error Resume Next

Do While Not REGGTIM.EOF
  
   R.AddNew
   For K = 0 To REGGTIM.FIELDS.Count - 1
     R(K) = REGGTIM(K)
   Next
   R.Update
   REGGTIM.MoveNext
Loop

Dim sum_mikta As Long, sum_posothta As Long

sum_mikta = 0
sum_posothta = 0

REGGTIM.Close

R.MoveFirst
Do While Not R.EOF
  If Not IsNull(R("fpa")) Then
     R.Edit
     
     If f_BARCODE_STOTIM = 1 Then
        R3.Open "select ERG FROM EID WHERE KOD='" + R("[Kωδ_Είδ]") + "'", Gdb, adOpenDynamic, adLockPessimistic
        If Not R3.EOF Then
          R("[Kωδ_Είδ]") = R3(0)
        End If
        R3.Close
        
     End If
     
     R("fpa") = g_Fpa(R("fpa"))
     sum_mikta = sum_mikta + R("[Ποσότητα]") * R("[Tιμ_Μον]")
     sum_posothta = sum_posothta + R("[Ποσότητα]")
     R.Update
   End If
   R.MoveNext
Loop








'db.Execute "update EGGTIM set FPA=" + Str(g_Fpa(6)) + " where FPA=6"

'db.Execute "update EGGTIM set FPA=" + Str(g_Fpa(1)) + " where FPA=1"
'db.Execute "update EGGTIM set FPA=" + Str(g_Fpa(2)) + " where FPA=2"
'db.Execute "update EGGTIM set FPA=" + Str(g_Fpa(3)) + " where FPA=3"
'db.Execute "update EGGTIM set FPA=" + Str(g_Fpa(4)) + " where FPA=4"

'db.Execute "update EGGTIM set FPA=0 where FPA=5 "

kn = f_SynSeiresTimologioy - R.RecordCount  'f_SeiresTimologioy - R.RecordCount
For K = 1 To kn
   R.AddNew: R.Update
Next
R.Close








20
synt = " left(ATIM," + Str(f_psifia_atim + 1) + ")='" + Left(MATIM, f_psifia_atim + 1) + "'"


' αυτο το κομματι δουλευei με odbc  ================================
'Dim db9 As Database
'db.Execute "drop table tim"
'Set db9 = OpenDatabase("", False, False, gConnect)
'SQL = "select AJI,AJ1,AJ2,AJ3,FPA1,FPA2,FPA3,PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.TYP,PEL.DOY,PEL.KOD,AJ4,AJ5,FPA4,PEL.THL,PEL.POL "
'SQL = SQL + " into tim in 'c:\mercvb\reports\reports.mdb'  from TIM inner join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>=#" + Format(d1, "mm/dd/yyyy") + "# and HME<#" + Format(d2, "mm/dd/yyyy") + "#;"
'db9.Execute SQL
'GoTo 22
' αυτο το κομματι δουλευei με odbc   ==================================


'========  TIM ==========
sql = "select AJI,AJ1,AJ2,AJ3,FPA1,FPA2,FPA3,PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.TYP,PEL.DOY,PEL.KOD,AJ4,AJ5,FPA4,PEL.THL,PEL.POL,AJ6,FPA6,ELGA,PEL.CH1 AS DEH "
sql = sql + "  from TIM inner join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';"
  db.Execute "delete *from tim"
rtim.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

'copy to access table tim
Set R = db.OpenRecordset("tim")
On Error Resume Next


'
'Dim N As Integer
'On Error GoTo 0
'For N = 0 To R.FIELDS.Count - 1
'   If R.FIELDS(N).Name = "DEH" Then
'      OK = 1
'   End If
'Next
'db.Execute "ALTER TABLE tim add column DEH CHAR(12) "
'



rtim.MoveFirst
   
' select ATIM,
' HME as shme,AJI as ajia,PEL.TYP,
   
   
   
'   On Error GoTo 0
   R.AddNew
   
On Error Resume Next

Dim ISPIST As Boolean


   
         R("SHME") = rtim("SHME")
         R("TYP") = rtim("TYP")
   
   
         R("EPO") = rtim("EPO")
         R("EPA") = Left(rtim("EPA"), 16)
         R("DIE") = rtim("DIE")
         R("AFM") = rtim("AFM")
         
         R("DOY") = rtim("DOY")
         R("KOD") = rtim("KOD")
         R("THL") = rtim("THL")
         R("POL") = rtim("POL")
   
   
         R("AJIA") = rtim("AJIA")
         R("AJI") = rtim("AJI")
         R("AJ1") = rtim("AJ1")
         R("AJ2") = rtim("AJ2") '
         R("AJ3") = rtim("AJ3")
         R("AJ4") = rtim("AJ4")
         R("AJ5") = rtim("AJ5")
         R("AJ6") = rtim("AJ6")
         
         R("FPA1") = rtim("FPA1")
         R("FPA2") = rtim("FPA2") '
         R("FPA3") = rtim("FPA3")
         R("FPA4") = rtim("FPA4")
         R("FPA6") = rtim("FPA6")
         R("ELGA") = rtim("ELGA")
         
         R("ATIM") = rtim("ATIM")
         
         
         R("DEH") = Left(rtim("DEH"), 12)
         
         'BLEPEI AN EINAI PISTOTIKO
         If Left(R("ATIM"), 1) = "P" Then
            ISPIST = True
         Else
            ISPIST = False
         End If
         
         
On Error GoTo LATOS
   R.Update
On Error Resume Next
   
rtim.Close

If IsNull(R("shme")) Or IsNull(R("epo")) Then
   printCrystal = -1000
End If





















'rtim.MoveFirst
'Do While Not rtim.EOF
'   R.AddNew
'   For K = 0 To R.FIELDS.Count - 1
'     If UCase(left(R(K).Name, 3)) = "EPA" Then
'         R(K) = left(rtim(R(K).Name), 16)
'     Else
'         R(K) = rtim(rtim(K).Name)
'     End If
'
'   Next
'   R.Update
'   rtim.MoveNext
'Loop
'rtim.Close



22


'=============  KENES SEIRES NA GEMISEI EGGTIM  ===================

'Set meggtim = db.OpenRecordset("eggtim")
Dim SAJIA As Single


Set R = db.OpenRecordset("SELECT SUM([Αξία]) FROM EGGTIM")
SAJIA = R(0)
R.Close

Set R = db.OpenRecordset("SELECT * FROM TIM")

   
If F_XONDR = 0 Then 'LIANIKH
    db.Execute "UPDATE TIM  SET AJI=" + Str(SAJIA) + " , AJIA=" + Str(SAJIA)
   
    db.Execute "UPDATE EGGTIM SET [Tιμ_Μον]=[Tιμ_Μον]*(1+fpa/100)"
'  db.Execute "UPDATE EGGTIM SET [Αξία]=round([Αξία],2)"
  
    db.Execute "UPDATE TIM Set aj1=0,aj2=0,aj3=0,aj4=0,aj5=0,AJ6=0,fpa1=0,fpa2=0,fpa3=0,fpa4=0"
   
Else
     If Abs(SAJIA - (R("AJ1") + R("AJ2") + R("AJ3") + R("AJ4") + R("AJ5") + R("AJ6"))) > 0.005 Then
       MsgBox "Aσυμφωνία παραστατικού. Διαγράψτε το παραστατικό και ξαναπεράστε το"
       End
     End If
   
End If



'=======================ΕΝΗΜΕΡΩΝΩ ΤΟ MEM ΜΕ ΤΑ ΣΤΑΘΕΡΑ ΣΤΟΙΧΕΙΑ ========
Dim fmem As Recordset
Set fmem = db.OpenRecordset("mem")
fmem.Edit
fmem("c1") = Left(F_PARAS, Len(F_PARAS) - 1) 'parastatiko
fmem("c2") = SKOPOS.Text 'skopos
fmem("c3") = FORTOSH.Text 'fortosi
fmem("c4") = PROORISMOS.Text 'paradosi
fmem("banks") = Combo4.Text 'τροποσ πληρωμης
fmem("pict") = Time$
fmem("par1") = Text1.Text
fmem("par2") = SEIRA.Text
fmem("par3") = AYTOKINHTO.Text
fmem("par19") = PEL.Recordset("PLAISIO") 'METAFORIKH ETAIREIA PELATH
fmem("par6") = Format(DTPicker1.value, "dd/mm/yyyy")
fmem("n1") = sum_mikta

fmem("n2") = sum_posothta




fmem("MEMO") = parat.Text
If Combo4.Text = Combo4.List(0) Then ' METRITOIS
   fmem("par4") = Format(PEL.Recordset("TYP"), "#####.00")
Else
   fmem("par4") = Format(PEL.Recordset("TYP") - Val(SynoloG.Caption), "#####.00")
End If

If ISPIST = True Then fmem("PAR4") = " "



fmem.Update

If IsNull(R("hme")) Or IsNull(R("epo")) Then
'   printCrystal = -1000
End If





Exit Function

MHNYMA:
  HandleError "Par1:PrintCrystal"
  printCrystal = -1000
Resume Next



LATOS:
  printCrystal = -1000


Resume Next
            


End Function

Function PrintSqlCrystal(MATIM, char_date)
'------------------------------------------------------------------------
'δημιουργεί τα αρχεία TIM  , EGGTIM , MEM    ΣΤΟ  c:\mercvb\reports\reports.mdb"
'ΓΙΑ ΤΗΝ ΕΚΤΥΠΩΣΗ ΤΟΥ ΤΙΜΟΛΟΓΙΟΥ

Dim sql As String, fDB As Database, d1 As Date, d2 As Date
Dim synt As String, db As Database
Dim DUM
Dim REGGTIM As New ADODB.Recordset
Dim rtim As New ADODB.Recordset
Dim R As New ADODB.Recordset
Dim K As Integer
Dim kn As Integer

'-----------  next lines are OK --------------------------

On Error GoTo MHNYMA

d1 = DTPicker1.value
d2 = DateAdd("d", 1, d1)

'Set db = OpenDatabase("c:\mercvb\reports\reports.mdb")
'If gConnect = "Access" Then
'  Set fDB = OpenDatabase(gDir, False, False)
'
'Else
'  Set fDB = OpenDatabase(gDir, False, False, gConnect)
'End If
'                  arn_enhm_apot

'=====   EGGTIM  ======


'On Error Resume Next
'db.Execute "delete *from eggtim"
10




'TO   round(TIMM*POSO*(100-EKPT)/100,2)  ANTIKATASTAUHKE ME TO KAU_AJIA  4-1-2006


'sql = sql + "MONA,EKPT "
'sql = sql + " from EGGTIM where APOT=" + Str(Val(left(APOTH1.Text, 2))) + " and  left(ATIM," + Str(F_PSIFIA_ATIM + 1) + ")='" + left(mAtim, F_PSIFIA_ATIM + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "' ORDER BY PROOD_AJ;"
'REGGTIM.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

'copy to access table eggtim
'Set R = db.OpenRecordset("eggtim")
'REGGTIM.MoveFirst
'Do While Not REGGTIM.EOF
  
 '  R.AddNew
  ' For K = 0 To REGGTIM.FIELDS.Count - 1
   '  R(K) = REGGTIM(K)
   'Next
   'R.Update
   'REGGTIM.MoveNext
'Loop

sql = "select FPA,ATIM,HME as [Ημερομ],KODE as [Kωδ_Είδ],ONOMA as [Περιγραφή],"

If F_XONDR = 0 Then 'LIANIKH
   sql = sql + "POSO as [Ποσότητα],TIMM as [Tιμ_Μον],MIK_AJIA as [Αξία],"
Else
   sql = sql + "POSO as [Ποσότητα],TIMM as [Tιμ_Μον],KAU_AJIA as [Αξία],"
End If
sql = sql + "MONA,EKPT "

sql = sql + " INTO EGGTIMREP from EGGTIM where APOT=" + Str(Val(Left(APOTH1.Text, 2))) + " and  left(ATIM," + Str(f_psifia_atim + 1) + ")='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "' ORDER BY PROOD_AJ;"

On Error Resume Next
Gdb.Execute "DROP TABLE TIMREP"
Gdb.Execute "DROP TABLE EGGTIMREP"
Gdb.Execute sql
'GoTo 10
On Error GoTo 0


On Error Resume Next
R.Open "SELECT * FROM EGGTIMREP", Gdb, adOpenForwardOnly, adLockReadOnly
'
''copy to access table eggtim
'Set R = db.OpenRecordset("eggtim")
'REGGTIM.MoveFirst
'Do While Not REGGTIM.EOF
'
'   R.AddNew
'   For K = 0 To REGGTIM.FIELDS.Count - 1
'     R(K) = REGGTIM(K)
'   Next
'   R.Update
'   REGGTIM.MoveNext
'Loop
'
'
'
'REGGTIM.Close

Gdb.Execute "update EGGTIMREP set FPA=" + Str(g_Fpa(1)) + " where FPA=1"
Gdb.Execute "update EGGTIMREP set FPA=" + Str(g_Fpa(2)) + " where FPA=2"
Gdb.Execute "update EGGTIMREP set FPA=" + Str(g_Fpa(3)) + " where FPA=3"
Gdb.Execute "update EGGTIMREP set FPA=" + Str(g_Fpa(4)) + " where FPA=4"
Gdb.Execute "update EGGTIMREP set FPA=0 where FPA=5 "
R.Open "SELECT COUNT(*) FROM EGGTIMREP", Gdb, adOpenForwardOnly, adLockReadOnly
kn = f_SeiresTimologioy - R(0)  'f_SeiresTimologioy - R.RecordCount
R.Close





'On Error GoTo 0

'R.Close
R.Open "SELECT * FROM EGGTIMREP", Gdb, adOpenDynamic, adLockOptimistic

'R.Open "SELECT * FROM EGGTIMREP", Gdb, adOpenForwardOnly, adLockReadOnly



' On Error GoTo 0
For K = 1 To kn
   R.AddNew: R.Update
Next
R.Close








20
synt = " left(ATIM," + Str(f_psifia_atim + 1) + ")='" + Left(MATIM, f_psifia_atim + 1) + "'"


' αυτο το κομματι δουλευei με odbc  ================================
'Dim db9 As Database
'db.Execute "drop table tim"
'Set db9 = OpenDatabase("", False, False, gConnect)
'SQL = "select AJI,AJ1,AJ2,AJ3,FPA1,FPA2,FPA3,PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.TYP,PEL.DOY,PEL.KOD,AJ4,AJ5,FPA4,PEL.THL,PEL.POL "
'SQL = SQL + " into tim in 'c:\mercvb\reports\reports.mdb'  from TIM inner join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>=#" + Format(d1, "mm/dd/yyyy") + "# and HME<#" + Format(d2, "mm/dd/yyyy") + "#;"
'db9.Execute SQL
'GoTo 22
' αυτο το κομματι δουλευei με odbc   ==================================

Gdb.Execute "DROP TABLE TIMREP2"
'========  TIM ==========
sql = "select AJI,AJ1,AJ2,AJ3,FPA1,FPA2,FPA3,PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.TYP,PEL.DOY,PEL.KOD,AJ4,AJ5,FPA4,PEL.THL,PEL.POL "
sql = sql + " INTO TIMREP2 from TIM inner join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';"

Gdb.Execute sql


'=============  KENES SEIRES NA GEMISEI EGGTIM  ===================

'Set meggtim = db.OpenRecordset("eggtim")
Dim SAJIA As Single


R.Open "SELECT SUM([Αξία]) FROM EGGTIMREP", Gdb, adOpenForwardOnly, adLockReadOnly


SAJIA = R(0)
R.Close

'Set R = db.OpenRecordset("SELECT * FROM TIM")
R.Open "SELECT * FROM TIMREP2", Gdb, adOpenForwardOnly, adLockReadOnly
   
Dim MATIM2
MATIM2 = R("ATIM")

   
   
If F_XONDR = 0 Then 'LIANIKH
    db.Execute "UPDATE TIMREP2  SET AJI=" + Str(SAJIA) + " , AJIA=" + Str(SAJIA)
   
    db.Execute "UPDATE EGGTIMREP SET [Tιμ_Μον]=[Tιμ_Μον]*(1+fpa/100)"
'  db.Execute "UPDATE EGGTIM SET [Αξία]=round([Αξία],2)"
  
    db.Execute "UPDATE TIMREP2 Set aj1=0,aj2=0,aj3=0,aj4=0,aj5=0,fpa1=0,fpa2=0,fpa3=0,fpa4=0"
   
Else
     If Abs(SAJIA - (R("AJ1") + R("AJ2") + R("AJ3") + R("AJ4") + R("AJ5"))) > 0.005 Then
       MsgBox "Aσυμφωνία παραστατικού. Διαγράψτε το παραστατικό και ξαναπεράστε το"
       End
     End If
   
End If



'=======================ΕΝΗΜΕΡΩΝΩ ΤΟ MEM ΜΕ ΤΑ ΣΤΑΘΕΡΑ ΣΤΟΙΧΕΙΑ ========

On Error GoTo 0
Dim fmem As New ADODB.Recordset
Gdb.Execute "DROP TABLE MEMREP"
Gdb.Execute "SELECT TOP 1  C1,C2,C3,C4,BANKS,PICT,PAR1,PAR2,PAR3,PAR4,PAR5,MEMO  INTO MEMREP FROM MEM"

 fmem.Open "SELECT COUNT(*) FROM MEMREP", Gdb, adOpenDynamic, adLockOptimistic
If fmem(0) = 0 Then
    fmem.AddNew
    fmem.Update
End If

fmem.Close


fmem.Open "SELECT * FROM MEMREP", Gdb, adOpenDynamic, adLockOptimistic
'fmem.Edit
fmem("c1") = Left(F_PARAS, Len(F_PARAS) - 1) 'parastatiko
fmem("c2") = SKOPOS.Text 'skopos
fmem("c3") = FORTOSH.Text 'fortosi
fmem("c4") = PROORISMOS.Text 'paradosi
fmem("banks") = Combo4.Text 'τροποσ πληρωμης
fmem("pict") = Time$
fmem("par1") = Text1.Text
fmem("par2") = SEIRA.Text
fmem("par3") = AYTOKINHTO.Text
fmem("par5") = MATIM2
fmem("par6") = Format(DTPicker1.value, "dd/mm/yyyy")

fmem("MEMO") = parat.Text
If Combo4.Text = Combo4.List(0) Then ' METRITOIS
   fmem("par4") = Format(PEL.Recordset("TYP"), "#####.00")
Else
   fmem("par4") = Format(PEL.Recordset("TYP") - Val(SynoloG.Caption), "#####.00")
End If
fmem.Update




Gdb.Execute "SELECT TIMREP2.* , MEMREP.*  INTO TIMREP FROM TIMREP2,MEMREP"

Exit Function

MHNYMA:
HandleError "Par1:PrintCrystal"
Resume Next
            


End Function




Function ektyp_forma(ByVal arxeio As String, ByVal Atim As String, ByVal hme As Date)
' arxeio π.χ. f90.txt  η φόρμα της εκτύπωσης

' atim αριθμός παραστατικού
' hme  ημερομηνία
 
 


Dim DUM
 Dim eField(1 To 50), ePict(1 To 50), eProw(1 To 50), ePcol(1 To 50), eFonts(1 To 50)
 Dim mField(1 To 50), mPict(1 To 50), mProw(1 To 50), mPcol(1 To 50), mFonts(1 To 50)
 Dim sField(1 To 50), sPict(1 To 50), sProw(1 To 50), sPcol(1 To 50), sFonts(1 To 50)
 Dim m_No_of_seir, I, K, m, C
 Dim m_seir_synol, m_seir_eid, sf
 Dim a(1 To 250), ar_ped(1 To 250)
 Dim xa(1 To 250), m_m, npic
 Dim m1
' dim gm_str(1 To 250)
Dim db As Database
Dim TIM As Recordset
Dim PEL As Recordset
Dim EGGTIM As Recordset
Dim m_entol_ektyp
Dim xart11, kod_tim
Dim ejodos
Dim e, mM, S
Dim ektypoths
Dim SELIDA
Dim m_syn_row, syn_row

On Error GoTo MHNYMA



'cEGGTIM.RecordSource = "eggtim"
Set fSCR = CreateObject("MSScriptControl.ScriptControl")



cEGGTIM.RecordSource = "select *from eggtim"
cEGGTIM.Refresh

CTIM.Connect = "Access"
CTIM.DatabaseName = "C:\MERCVB\REPORTS\REPORTS.MDB"
CTIM.RecordSource = "select *from tim"
CTIM.Refresh


fSCR.language = "vbscript"
fSCR.AddObject "cTIM", CTIM
fSCR.AddObject "cEGGTIM", cEGGTIM
fSCR.AddObject "cMEM", CMEM
fSCR.AddObject "mactext", mactext

'Set db = OpenDatabase(gDir, False, False, gConnect)
'fSCR.ExecuteStatement "CPEL.DATABASENAME=" + Chr$(34) + gDir + Chr$(34)
'fSCR.ExecuteStatement "CTIM.DATABASENAME=" + Chr$(34) + gDir + Chr$(34)
'fSCR.ExecuteStatement "CEGGTIM.DATABASENAME=" + Chr$(34) + gDir + Chr$(34)
'fSCR.ExecuteStatement "CMEM.DATABASENAME=" + Chr$(34) + gDir + Chr$(34)
'fSCR.ExecuteStatement "eid.DATABASENAME=" + Chr$(34) + gDir + Chr$(34)



'fSCR.ExecuteStatement "CTIM.RECORDSOURCE=" + Chr$(34) + "select *from tim where atim='" + atim + "' and hme>=#" + Format(hme, "mm/dd/yyyy") + "#;" + Chr$(34)

'fSCR.ExecuteStatement "CPEL.RECORDSOURCE=" + Chr$(34) + "select *from pel where eidos='" + "e" + "' and kod='" + CTIM("kpe") + "';" + Chr$(34)
'fSCR.ExecuteStatement "CEGGTIM.RECORDSOURCE=" + Chr$(34) + "select *from eggtim where atim='" + atim + "' and hme>=#" + Format(hme, "mm/dd/yyyy") + "#;" + Chr$(34)
'fSCR.ExecuteStatement "CMEM.RECORDSOURCE=" + Chr$(34) + "SELECT *FROM MEM" + Chr$(34)

'fSCR.ExecuteStatement "CMEM.RECORDSOURCE=" + Chr$(34) + "SELECT *FROM eid" + Chr$(34)

'Set gEggtim = db.OpenRecordset("select *from eggtim where atim='" + atim + "' and hme>=#" + Format(hme, "mm/dd/yyyy") + "#;")
cEGGTIM.Recordset.MoveFirst

'Set gTim = db.OpenRecordset("select *from tim where atim='" + atim + "' and hme>=#" + Format(hme, "mm/dd/yyyy") + "#;")
'Set gmem = db.OpenRecordset("select *from mem")
CMEM.Recordset.MoveFirst
'CMEM.Recordset.Move 90


CTIM.Recordset.MoveFirst

'Set gPel = db.OpenRecordset("select *from pel where eidos='" + "e" + "' and kod='" + gTim("kpe") + "';")
'Set gEid = db.OpenRecordset("eid")

'On Error GoTo create_index
'gEid.Index = "eidko"



'gEid.Seek "=", "ωωωω"




    ' arxeio = "c:\mercvb\f90.txt"
    'c=ltrim(str(select(),2))  '&& ΅¨«ι «¤ §¨ ¦®γ ©«γ¤ ¦§¦ε ™¨ ©΅ζ£¦¬¤

    'load_forma (arxeio)

'Dim DUM
     DUM = load_forma(arxeio, 1)
 'Me.Caption = "load_forma=OK"
      
  For I = 1 To 250
    If Left$(gm_str(I), 10) = "**********" Then
       Exit For
    End If
  Next


    '* ΣΥΝΟΛΙΚΕΣ ΣΕΙΡΕΣ
         m_No_of_seir = Val(gm_str(I + 1))

    '* ΣΕΙΡΑ ΠΟΥ ΑΡΧΙΖΟΥΝ ΤΑ ΕΙΔΗ
         m_seir_eid = Val(gm_str(I + 2))

    '* ΣΕΙΡΑ ΠΟΥ ΑΡΧΙΖΟΥΝ ΤΑ ΣΥΝΟΛΑ
         m_seir_synol = Val(gm_str(I + 3))
    
    '* entolh gia ton εκτυπωτή
         m_entol_ektyp = gm_str(I + 4)

    '**************
    ' PRINTER.
    Dim cSelida
    If F928 = 1 Then
       cSelida = "Σελίδα "
    Else
       cSelida = to437("Σελίδα ")
    End If
    
    
    
    
    
    
    
    
    SELIDA = 1
10 Open F_File_Timologioy For Output As #1
    
    '---------------- επικεφαλίδες
20  ' On Error GoTo errorl
30      e = 0
40      Do While Len(Trim(gm_f(e + 1))) > 0
         e = e + 1
         'MM = Trim(gm_f(e))
'
         'm_m = ExecuteLine("gvar=" + MM)
         'eField(e) = gvar
50         ePict(e) = Trim(gpic(e))
60         eProw(e) = gm_r(e)
70         ePcol(e) = gm_c(e)
80         eFonts(e) = ""  '     && trim( Fonts )
        ' Printer.Print gm_f(e) + Format(eProw(e), "0000") + "=" + Format(ePcol(e), "0000") + "="
90      Loop
      '**************** ΣΕΙΡΕΣ ΜΕ ΕΙΔΗ ****************************************
100      m = 0
110    Do While Len(Trim(gm_f(e + m + 2))) > 0
120         m = m + 1

         'MM = Trim(gm_f(M + e + 1))

         'm_m = ExecuteLine("gvar=" + MM)

         'mField(M) = gvar
130         mPict(m) = Trim(gpic(m + e + 1))
140         mProw(m) = gm_r(m + e + 1)
150         mPcol(m) = gm_c(m + e + 1)
         'mFonts(M) = "" '     && trim( Fonts )
         'MM = mPict(M + e + 1)
160      Loop


Dim fcheckonly As Boolean


      '****************  ΣΥΝΟΛΑ  ****************************************
      S = 0
170      Do While Len(Trim(gm_f(e + m + S + 3))) > 0
         S = S + 1
'         MM = Trim(gm_f(M + e + s + 2))
'         MM = "gvar=" + MM
'         m_m = ExecuteLine("gvar=" + MM)
'
'         sField(s) = gvar
180         sPict(S) = Trim(gpic(m + e + S + 2))
190         sProw(S) = gm_r(m + e + S + 2)
200         sPcol(S) = gm_c(m + e + S + 2)
'         sFonts(s) = ""  '   && trim( Fonts )
'         MM = sPict(M + e + s + 2)
'         'Debug.Print Format(sField(s), MM)
210      Loop


 'Printer.FontSize = 10
' Printer.FontName = "Courier New"
' Printer.Font.Charset = 161
 





       If Len(Trim(m_entol_ektyp)) > 0 Then
           'Printer.Print  ' &_entol_ektyp
       End If



'===================================== ΕΠΙΚΕΦΑΛΙΔΕΣ ===================
       syn_row = 0
       ' ΕΚΤΥΠΩΣΗ ΕΠΙΚΕΦΑΛΙΔΩΝ
       
' On Error GoTo errorl
       m1 = 0
208       For K = 1 To e
             
             mM = ePict(K)
             'For i = 1 To eProw(K): Printer.Print: Next
             For I = 1 To eProw(K)
                 If K = 1 And m_No_of_seir = 1000 Then
                    ' ΜΗΝ ΒΑΖΕΙΣ ΚΕΝΗ ΣΕΙΡΑ ΣΤΗΝ ΑΡΧΗ
                 Else
                     Print #1, " "
                 End If
                 
             Next
             syn_row = syn_row + eProw(K) 'συγκεντρώνω τις σειρές
             
220           mM = mac(Trim(Trim(gm_f(K)))) ': m_m = ExecuteLine("gvar=" + MM)
           'Printer.Print Tab(ePcol(K)); Format(gvar, ePict(K))
           
      If eProw(K + 1) > eProw(K) Then 'αν πρόκειται να τυπώσει παρακάτω μην τυπώνεις δίπλα μου
           '   Print #1, Tab(ePcol(K)); Format(mM, ePict(K))
           Print #1, Tab(ePcol(K)); Format(Left(mM, Len(ePict(K))), ePict(K));
      Else
           Print #1, Tab(ePcol(K)); Format(Left(mM, Len(ePict(K))), ePict(K));
      End If
             
           m1 = eProw(K)
             
230        Next
      
      
240   For K = 1 To m_seir_eid - syn_row
          Print #1, Space(80) 'Printer.Print:
      Next 'βρισκω την σειρά  των ειδών
       
       
250       syn_row = m_seir_eid
     '  Printer.Print
       'On Error GoTo 0
260       cEGGTIM.Recordset.MoveFirst
'=========================================== ΕΙΔΗ ====================
270       Do While Not cEGGTIM.Recordset.EOF
         

         DUM = "select *from eid where kod='" + cEGGTIM.Recordset("Kωδ_Είδ") + "'"
         If IsNull(cEGGTIM.Recordset("Kωδ_Είδ")) Then Exit Do
         
         'Set gEid = db.OpenRecordset(dum)
         For K = 1 To m
           
280           mM = mac(Trim(Trim(gm_f(e + K + 1))))                 ': m_m = ExecuteLine("gvar=" + MM)
           '   Printer.Print Format(gvar, gpic(e + K + 1));
290           Print #1, Tab(mPcol(K)); GFORMAT(mM, gpic(e + K + 1));
         Next
310         cEGGTIM.Recordset.MoveNext
            syn_row = syn_row + 1

            
            ' ΑΝ ΔΕΝ ΧΩΡΑΝΕ ΟΙ ΣΕΙΡΕΣ ΣΤΟ ΤΙΜΟΛΟΓΙΟ ( Η ΣΕΙΡΑ ΠΟΥ ΤΥΠΏΝΕΤΑΙ ΞΕΠΕΡΝΑ ΤΗΝ ΣΕΙΡΑ ΤΩΝ ΣΥΝΟΛΩΝ)
            If syn_row > m_seir_synol Then
                Print #1, ""
                Print #1, cSelida + Format(SELIDA, "##") ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
                SELIDA = SELIDA + 1
                Print #1, Chr(12) ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
                'If SELIDA > 1 Then
                 '   Print #1, "Σελίδα " + Format(SELIDA, "##"); ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
                'End If
                GoSub TYPOSE_HEADER

            End If




       Loop

       
  npic = Space(120)
  
330    For K = 1 To m_seir_synol - syn_row: Print #1,: Next  'βρισκω την σειρά  των synolon
       
       For K = 1 To S
340         mM = Trim(gm_f(m + e + K + 2))
350        ' m_m = ExecuteLine("gvar=" + MM)
         
                 mM = mac(Trim(gm_f(m + e + K + 2)))
         
         
360         For I = 1 To sProw(K): Print #1,: Next
         'Printer.Print Format(gvar, gpic(M + e + K + 2));
370         Print #1, Tab(sPcol(K)); GFORMAT(mM, gpic(m + e + K + 2));
       Next
       
'       If SELIDA > 1 Then
'                Print #1, "Σελίδα " + Format(SELIDA, "##") ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
'       End If
       
       
       
400 Close #1

380 cEGGTIM.Recordset.Close

Set fSCR = Nothing
If F_Type_File_Timologioy = 1 Then
   If F928 = 1 Then
      If Len(Dir("c:\marietos.txt")) > 1 Then
         ' kanei "type c:\print"  meso dos
         DUM = Shell("c:\mercvb\notedos.exe /p " + F_File_Timologioy, vbMaximizedFocus)
       ElseIf Len(Dir("c:\AGIOS.txt")) > 1 Then
         ' kanei "type c:\print"  meso dos
         DUM = Shell("c:\mercvb\ff.exe", vbMaximizedFocus)
         
         ElseIf Len(Dir("c:\KAMPOUR.txt")) > 1 Then
         ' kanei "type c:\print"  meso dos
            DUM = EKTYP_ME_PRINT(F_File_Timologioy)
            ' DUM = Shell("c:\mercvb\ff.exe", vbMaximizedFocus)
         
      Else
         DUM = Shell("c:\mercvb\notepad.exe /p " + F_File_Timologioy, vbMaximizedFocus)
      End If
   Else
      If Len(Dir("c:\FAKKAV.TXT")) > 1 Then
         DUM = EKTYP_ME_PRINT(F_File_Timologioy)
      Else
         DUM = Shell("C:\MERCVB\PRINT.EXE " + F_File_Timologioy, vbMinimizedFocus)
      End If
   End If
Else

      If Len(Dir("c:\DION2.EXE")) > 1 Then
        ' DUM = EKTYP_ME_PRINT(F_File_Timologioy)
         MILSEC 5000
         FileCopy "C:\KENO.TXT", "C:\FOROL\FOROL.TXT"
      ElseIf Len(Dir("c:\KAMPOUR.txt")) > 1 Then
          'MILSEC 8000
          Do While Len(Dir("c:\FOROL\FOROL.txt")) > 1
            MILSEC 500
          Loop
         FileCopy "C:\KENO.TXT", "C:\FOROL\FOROL.TXT"
         
         ' kanei "type c:\print"  meso dos
'            DUM = EKTYP_ME_PRINT(F_File_Timologioy)
         
         
      End If
   
   
End If






Exit Function



create_index:

'db.Execute "CREATE INDEX eidko ON eid(kod);"
Resume Next



errorl:
    MsgBox Str(Erl) + "MM=" + mM 'RECOVER USING objErrorInfo
      'sfalma = 1
      'MsgBox "Ά¦ª ©«¦ §›ε¦ " + Field + "  £.§›ε¦¬ " + Str(i)
'    Resume Next
 Close #1
 'Resume Next
 
 
 

 
 
 Exit Function
     

'
'User-Defined String Formats (Format Function)
'
'
'You can use any of the following characters to create a formatexpression for strings:
'
'Character Description
'@ Character placeholder. Display a character or a space. If the string has a character in the position where the at symbol (@) appears in the format string, display it; otherwise, display a space in that position. Placeholders are filled from right to left unless there is an exclamation point character (!) in the format string.
'& Character placeholder. Display a character or nothing. If the string has a character in the position where the ampersand (&) appears, display it; otherwise, display nothing. Placeholders are filled from right to left unless there is an exclamation point character (!) in the format string.
'< Force lowercase. Display all characters in lowercase format.
'> Force uppercase. Display all characters in uppercase format.
'! Force left to right fill of placeholders. The default is to fill placeholders from right to left.





TYPOSE_HEADER:
       syn_row = 0
       ' ΕΚΤΥΠΩΣΗ ΕΠΙΚΕΦΑΛΙΔΩΝ
       
       m1 = 0
       For K = 1 To e
             mM = ePict(K)
             For I = 1 To eProw(K)
                 If K = 1 And m_No_of_seir = 1000 Then
                    ' ΜΗΝ ΒΑΖΕΙΣ ΚΕΝΗ ΣΕΙΡΑ ΣΤΗΝ ΑΡΧΗ
                 Else
                     Print #1, " "
                 End If
             Next
             syn_row = syn_row + eProw(K) 'συγκεντρώνω τις σειρές
             mM = mac(Trim(Trim(gm_f(K)))) ': m_m = ExecuteLine("gvar=" + MM)
      If eProw(K + 1) > eProw(K) Then 'αν πρόκειται να τυπώσει παρακάτω μην τυπώνεις δίπλα μου
           Print #1, Tab(ePcol(K)); Format(Left(mM, Len(ePict(K))), ePict(K));
      Else
           Print #1, Tab(ePcol(K)); Format(Left(mM, Len(ePict(K))), ePict(K));
      End If
      m1 = eProw(K)
    Next
                If SELIDA > 1 Then
                    Print #1, cSelida + Format(SELIDA, "##") ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
                End If
    
    For K = 1 To m_seir_eid - syn_row - 1
          Print #1, Space(80) 'Printer.Print:
    Next 'βρισκω την σειρά  των ειδών
    syn_row = m_seir_eid
Return















MHNYMA:
HandleError "Par1:ektypforma"
Resume Next
            



End Function
Function EKTYP_ME_PRINT(F_File_Timologioy)
Dim a$
  
Printer.FontName = "Arial"
Printer.FONTSIZE = 12
Printer.FontBold = True

  Open F_File_Timologioy For Input As #1
  Do While Not EOF(1)
     Input #1, a$
     Printer.Print a$
  Loop
  Close #1
  Printer.EndDoc
End Function





Function mac(ByVal S As String) As String
'δινω ενα στρινγκ και μου επιστρέφει μία τιμή
Dim X
'On Error GoTo 0
mactext.Text = "   "
On Error GoTo err2
X = "MACtEXT.Text = " + S
fSCR.ExecuteStatement X

If F928 = 1 Then
    mac = mactext.Text
Else
    mac = to437(mactext.Text)
End If

Exit Function

err2:
Me.Caption = X
mac = "  "
Resume Next
End Function

Private Sub AYTOKINHTO_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       KeyAscii = 0 'suppress the beep
       keybd_event VK_TAB, 0, 0, 0 'send a tab
    End If
End Sub

Private Sub Text5_LostFocus()
   Dim K As Integer
   Dim genEkpt
   genEkpt = gVal(Text5.Text)
   If genEkpt > 0 And genEkpt < 100 Then
      
      Text5.Text = genEkpt
      For K = 1 To Grid1.ROWS - 1
         If Val(Grid1.TextMatrix(K, f_p)) <> 0 Then
            Grid1.TextMatrix(K, f_e) = (genEkpt + Val(Grid1.TextMatrix(K, f_e))) - (genEkpt * Val(Grid1.TextMatrix(K, f_e))) / 100
            Grid1.row = K: FindSynolo
         End If
      Next
   End If
   
   Dim telPOSO As Single, log As Single
   If Left(LTrim(Text5.Text), 1) = "*" Then
      telPOSO = gVal(mID(LTrim(Text5.Text), 2, 10))
      
      log = telPOSO / SynoloG  'συντελεστης σμίκρυνσης
      
      Dim arxSynoloG As Single
      arxSynoloG = SynoloG
      
      
      For K = 1 To Grid1.ROWS - 1
         If Val(Grid1.TextMatrix(K, F_T)) <> 0 Then
            Grid1.TextMatrix(K, F_T) = Round(log * Val(Grid1.TextMatrix(K, F_T)), 2)
            Grid1.row = K: FindSynolo
         End If
      Next
      
    If Abs(telPOSO - SynoloG) > 0.01 Then
      log = telPOSO / SynoloG  'epanypologismos συντελεστη σμίκρυνσης
      For K = 1 To Grid1.ROWS - 1
         If Val(Grid1.TextMatrix(K, F_T)) <> 0 Then
            Grid1.TextMatrix(K, F_T) = Round(log * Val(Grid1.TextMatrix(K, F_T)), 3)
            Grid1.row = K: FindSynolo
         End If
      Next
    End If
    
      
      
      
      Text5.Text = "    "
      
      
      
      
   End If
   
   
   
   
   
   
   
   
   
   
   
   
End Sub
Function gVal(a As String) As Single
  Dim KOMA
  ' KOMA = InStr(A, ","): If KOMA > 0 Then Mid$(A, KOMA, 1) = "."
   a = Replace(a, ",", ".")
  gVal = Val(a)
End Function


Function toascii(ByVal arxeio As String, moutput As String)
' arxeio π.χ. f90.txt  η φόρμα της εκτύπωσης    host.txt to arxeio me ejagogi ascii

Dim DUM
 Dim eField(1 To 50), ePict(1 To 50), eProw(1 To 50), ePcol(1 To 50), eFonts(1 To 50)
 Dim mField(1 To 50), mPict(1 To 50), mProw(1 To 50), mPcol(1 To 50), mFonts(1 To 50)
 Dim sField(1 To 50), sPict(1 To 50), sProw(1 To 50), sPcol(1 To 50), sFonts(1 To 50)
 Dim m_No_of_seir, I, K, m, C
 Dim m_seir_synol, m_seir_eid, sf
 Dim a(1 To 250), ar_ped(1 To 250)
 Dim xa(1 To 250), m_m, npic
 Dim m1
' dim gm_str(1 To 250)
Dim db As Database
Dim TIM As Recordset
Dim PEL As Recordset
Dim EGGTIM As Recordset
Dim m_entol_ektyp
Dim xart11, kod_tim
Dim ejodos
Dim e, mM, S
Dim ektypoths
Dim m_syn_row, syn_row
Dim CDOK

Set fSCR = CreateObject("MSScriptControl.ScriptControl")

DOK.ConnectionString = gConnect
DOK.RecordSource = "select *from EID  WHERE LEFT(ERG,2)='" + F_2CHFIA_ZYGIZOMENA + " ' ORDER BY ONO " '    LEFT(ERG,3)='210' "
DOK.Refresh


CTIM.Connect = "Access"
CTIM.DatabaseName = "C:\MERCVB\REPORTS\REPORTS.MDB"
CTIM.RecordSource = "select *from tim"
CTIM.Refresh


fSCR.language = "vbscript"
fSCR.AddObject "cTIM", CTIM
fSCR.AddObject "DOK", DOK 'cEGGTIM
fSCR.AddObject "cMEM", CMEM
fSCR.AddObject "mactext", mactext
DOK.Recordset.MoveFirst

CMEM.Recordset.MoveFirst

CTIM.Recordset.MoveFirst

'Dim DUM
     DUM = load_forma(arxeio, 1)
      
'-----------    GLOBAL ΜΕΤΑΒΛΗΤΕΣ  ----------------
'gm_str() οι σειρές του f99.txt
'gpic(i) το στρινγκ του πεδίου
'gm_r(i),gm_c(i) σειρά και στήλη του πεδίου
'gm_f(i) το όνομα του πεδίου

'gm_str(6)="     ^XXXXXXXXXXXXXXXXXXXXXXXXXXX               ^xxxxxxxx    @XXXXXXXX  ~ono_par  ~SUBS(tim->atim,2,5)  ~TIM->HME
'        gpic(2)='XXXXXXXXXXXXXXXXXXXXXXXXXXX
'gm_r(2)=5  gm_c(2)=6    gm_f(2)=ono_par
'=================================
    
10 Open moutput For Output As #1
    Dim S2 As String
    
  DOK.Recordset.MoveFirst
 Do While Not DOK.Recordset.EOF
         S2 = ""
         For K = 1 To 4
             mM = mac(Trim(Trim(gm_f(K))))                   ': m_m = ExecuteLine("gvar=" + MM)
             S2 = S2 + ALIGNFORMAT(mM, gpic(K)) '
         Next
         Print #1, S2
         DOK.Recordset.MoveNext
       Loop
 Close #1
 DOK.Recordset.Close

Set fSCR = Nothing

End Function

Public Function ALIGNFORMAT(mM, pic)
' RIGHT ALIGNMENT
If Left(pic, 1) = "#" Then
      ALIGNFORMAT = Right(Space(Len(pic)) + mM, Len(pic))
Else
' LEFT ALIGNMENT
      ALIGNFORMAT = Left(mM + Space(Len(pic)), Len(pic))
End If


End Function



Sub AdjustDataGridColumns(intRecord As Integer, _
           intField As Integer, _
           Optional AccForHeaders As Boolean)

Exit Sub


'This procedure will adjust DataGrids column width
'based on longest field in underlying source

'GRIDEIDON = DataGrid
'EID = Adodc control
'intRecord = Number of record
'intField = Number of field
'AccForHeaders = True or False

    Dim row As Long, Col As Long
    Dim width As Single, maxWidth As Single
    Dim saveFont As StdFont, saveScaleMode As Integer
    Dim cellText As String
    
    'If number of records = 0 then exit from the sub
    If intRecord = 0 Then Exit Sub
    'Save the form's font for DataGrid's font
    'We need this for form's TextWidth method
    Set saveFont = GridEidon.Parent.Font
    Set GridEidon.Parent.Font = GridEidon.Font
    'Adjust ScaleMode to vbTwips for the form (parent).
    saveScaleMode = GridEidon.Parent.ScaleMode
    GridEidon.Parent.ScaleMode = vbTwips
    'Always from first record...
    EID.Recordset.MoveFirst
    maxWidth = 0
    'We begin from the first column until the last column
    Dim mM
    If intField - 1 > 3 Then mM = 3 Else mM = intField - 1
    
    Dim FF
    If intRecord - 1 > 20 Then FF = 20 Else FF = intRecord - 1
    
    
    
    For Col = 0 To mM 'intField - 1
        EID.Recordset.MoveFirst
        'Optional param, if true, set maxWidth to
        'width of GRIDEIDON.Parent
        If AccForHeaders Then
            maxWidth = GridEidon.Parent.TextWidth(GridEidon.Columns(Col).Text) + 200
        End If
        'Repeat from first record again after we have
        'finished process the last record in
        'former column...
        EID.Recordset.MoveFirst
        For row = 0 To FF ' intRecord - 1
            'Get the text from the DataGrid's cell
            If intField = 1 Then
            Else  'If number of field more than one
                cellText = GridEidon.Columns(Col).Text
            End If
            'Fix the border...
            'Not for "multiple-line text"...
            width = GridEidon.Parent.TextWidth(cellText) + 200
            'Update the maximum width if we found
            'the wider string...
            If width > maxWidth Then
               maxWidth = width
               GridEidon.Columns(Col).width = maxWidth
            End If
            'Process next record...
            EID.Recordset.MoveNext
        Next row
        'Change the column width...
        GridEidon.Columns(Col).width = maxWidth 'kolom terakhir!
    Next Col
    'Change the DataGrid's parent property
    Set GridEidon.Parent.Font = saveFont
    GridEidon.Parent.ScaleMode = saveScaleMode
    'If finished, then move pointer to first record again
    EID.Recordset.MoveFirst
End Sub  'End of AdjustDataGridColumns



Sub find_eid_parastat()

  Dim R As New ADODB.Recordset

  R.Open "SELECT *FROM PARASTAT", Gdb, adOpenDynamic, adLockOptimistic
  R.MoveFirst
  Do While Not R.EOF
     If Right(PARAS.Text, 1) = R("EIDOS") Then
       F_TITLOS = R("TITLOS")
       F_POS_APOU = R("POS_APOU")
       F_AJIA_APOU = R("AJIA_APOU")
       f_pel = R("pel")
       f_pol = R("pol") ' 1=POLISEIS 2=AGORES
       F_STADIO = R("STADIO") ' 1=ΠΡΟΣΦ 2=ΠΑΡΑΓΓ 3=ΔΑ 4=ΤΙΜ.ΠΟΛ 5=ΕΝΔΟΔΙΑΚΙΝΙΣΙ
       If IsNull(R("GEF_P")) Then
           F_rec_p = 0
       Else
           F_rec_p = R("GEF_P") ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
       End If
       If IsNull(R("XONDR")) Then
           F_XONDR = 0
       Else
           F_XONDR = R("XONDR") ' RECORD APO GEFYRES.DBF OPOY EINAI OI KODIKOI GEFIROSEON
       End If
     End If
     R.MoveNext
  Loop
  
  
End Sub

Function Get_meKomma(ByVal X As String) As String
' δινεις το αβγ  και το κανει      α','β','γ

Dim pol As String

  pol = ""
Dim K As Integer

For K = 1 To Len(X)
  pol = pol + "'" + mID$(X, K, 1) + "',"
Next

pol = Left(pol, Len(pol) - 2)
pol = mID$(pol, 2, Len(pol) - 1)

Get_meKomma = pol

End Function


