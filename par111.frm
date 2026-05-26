VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form par1 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12075
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10335
   ScaleWidth      =   12075
   WindowState     =   2  'Maximized
   Begin VB.CommandButton kerdos 
      Height          =   255
      Left            =   10440
      TabIndex        =   69
      Top             =   9360
      Width           =   255
   End
   Begin VB.CheckBox proepiskopisi 
      Caption         =   "Προεπισκόπηση"
      Height          =   255
      Left            =   10440
      TabIndex        =   68
      Top             =   8040
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2535
      Left            =   10440
      TabIndex        =   67
      Top             =   4560
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   12
      Cols            =   3
   End
   Begin MSAdodcLib.Adodc LASTTIMOL 
      Height          =   336
      Left            =   8520
      Top             =   10080
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
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Bindings        =   "par111.frx":0000
      Height          =   1545
      Left            =   45
      TabIndex        =   66
      Top             =   8025
      Width           =   10305
      _ExtentX        =   18177
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
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
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
   Begin VB.CommandButton Command6 
      Caption         =   "Διόρθωση"
      Height          =   210
      Left            =   6060
      TabIndex        =   65
      Top             =   690
      Width           =   915
   End
   Begin VB.CheckBox metasx 
      BackColor       =   &H00FF0000&
      Caption         =   "Μετασχ.Παραστατικού"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   10395
      TabIndex        =   64
      ToolTipText     =   "Με διόρθωση μπορώ να μετατρέψω το παραστατικό χωρίς να χαθεί το αρχικό παραστατικό"
      Top             =   7230
      Width           =   2100
   End
   Begin VB.CommandButton COMMAND5 
      Caption         =   "Επανεκτύπωση Τελευταίου Παραστατικού"
      Enabled         =   0   'False
      Height          =   372
      Left            =   10920
      TabIndex        =   59
      Top             =   9840
      Visible         =   0   'False
      Width           =   2868
   End
   Begin VB.CommandButton Command4 
      Height          =   240
      Left            =   7035
      TabIndex        =   58
      Top             =   990
      Width           =   165
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kαρτέλα"
      Height          =   210
      Left            =   3210
      TabIndex        =   57
      Top             =   720
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid GridEidon 
      Bindings        =   "par111.frx":0018
      Height          =   2715
      Left            =   15
      TabIndex        =   30
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
      TabIndex        =   55
      Text            =   "1"
      Top             =   7635
      Width           =   690
   End
   Begin VB.TextBox TIMText3 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   45
      TabIndex        =   23
      Top             =   4815
      Width           =   30
   End
   Begin VB.TextBox metaf 
      Height          =   270
      Left            =   4890
      TabIndex        =   51
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
      Caption         =   "Προμηθευτής"
      Height          =   360
      Left            =   5640
      TabIndex        =   4
      Top             =   240
      Width           =   1230
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
      ItemData        =   "par111.frx":002A
      Left            =   120
      List            =   "par111.frx":0037
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.ComboBox PARAS 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "par111.frx":005B
      Left            =   1560
      List            =   "par111.frx":0068
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
      TabIndex        =   39
      Top             =   7215
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ΖΥΓΟΣ"
      Height          =   255
      Left            =   10560
      TabIndex        =   38
      Top             =   8520
      Width           =   1140
   End
   Begin VB.ComboBox AYTOKINHTO 
      Height          =   315
      Left            =   7335
      TabIndex        =   37
      Top             =   6570
      Width           =   1485
   End
   Begin VB.CheckBox kostologhsh 
      Caption         =   "Κοστολόγηση"
      Height          =   255
      Left            =   8865
      TabIndex        =   36
      Top             =   6960
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Λήψη Τιμ.ASCII"
      Height          =   240
      Left            =   8865
      TabIndex        =   35
      Top             =   4860
      Width           =   1470
   End
   Begin VB.CheckBox lianikh 
      Caption         =   "Λιανική"
      Height          =   255
      Left            =   8865
      TabIndex        =   34
      Top             =   6720
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox ektyp 
      Caption         =   "Εκτύπωση"
      Height          =   255
      Left            =   8865
      TabIndex        =   33
      Top             =   6480
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   7320
      TabIndex        =   31
      Top             =   6915
      Width           =   1485
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2625
      Top             =   9600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      Destination     =   1
      PrintFileType   =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlxGridDeltion 
      Bindings        =   "par111.frx":008C
      Height          =   2655
      Left            =   240
      TabIndex        =   29
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
      Left            =   3825
      Top             =   9600
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
      Left            =   3195
      Top             =   9270
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
      Left            =   1155
      Top             =   9270
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
      Left            =   105
      Top             =   9480
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
      Left            =   9480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
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
      Left            =   9345
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
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
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9465
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox MACtEXT 
      Height          =   285
      Left            =   6945
      TabIndex        =   27
      Text            =   "Text5"
      Top             =   9720
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
      Left            =   7635
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "mem"
      Top             =   9390
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
      Left            =   8835
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9390
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
      Left            =   6315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9390
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CheckBox ApoParaggelia 
      Caption         =   "Από παραγγελία"
      Height          =   255
      Left            =   6825
      TabIndex        =   25
      Top             =   4905
      Width           =   1995
   End
   Begin VB.ComboBox PROORISMOS 
      Height          =   315
      ItemData        =   "par111.frx":009E
      Left            =   6285
      List            =   "par111.frx":00A0
      TabIndex        =   20
      Top             =   6240
      Width           =   2535
   End
   Begin VB.ComboBox FORTOSH 
      Height          =   315
      ItemData        =   "par111.frx":00A2
      Left            =   6285
      List            =   "par111.frx":00A4
      TabIndex        =   19
      Top             =   5880
      Width           =   2535
   End
   Begin VB.ComboBox SKOPOS 
      Height          =   315
      ItemData        =   "par111.frx":00A6
      Left            =   6285
      List            =   "par111.frx":00A8
      TabIndex        =   18
      Top             =   5520
      Width           =   2535
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "par111.frx":00AA
      Left            =   6285
      List            =   "par111.frx":00AC
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
      Left            =   4690
      TabIndex        =   3
      Top             =   240
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   3615
      Left            =   105
      TabIndex        =   24
      Top             =   1245
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   6376
      _Version        =   393216
      BackColor       =   16744576
      Rows            =   300
      Cols            =   20
      BackColorSel    =   16744576
      _NumberOfBands  =   1
      _Band(0).Cols   =   20
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
      Format          =   127205377
      CurrentDate     =   38294
   End
   Begin MSDataGridLib.DataGrid GridPelaton 
      Bindings        =   "par111.frx":00AE
      Height          =   3600
      Left            =   45
      TabIndex        =   28
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
      TabIndex        =   63
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
      TabIndex        =   62
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
      TabIndex        =   61
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
      TabIndex        =   60
      Top             =   5220
      Width           =   450
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   285
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
      TabIndex        =   56
      Top             =   7620
      Width           =   1065
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "F3=SN F5:Τελευταία Τιμή     F8:Ποσότητα από Αξία      F9:Τιμή χωρίς ΦΠΑ"
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   -15
      TabIndex        =   54
      Top             =   7860
      Width           =   8775
   End
   Begin VB.Label Label15 
      Caption         =   ".."
      Height          =   255
      Left            =   5880
      TabIndex        =   53
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
      Height          =   210
      Left            =   2025
      TabIndex        =   52
      Top             =   7620
      Width           =   2790
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
      TabIndex        =   50
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
      TabIndex        =   49
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
      TabIndex        =   48
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
      TabIndex        =   47
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
      TabIndex        =   46
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
      TabIndex        =   45
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
      TabIndex        =   44
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
      TabIndex        =   43
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
      TabIndex        =   42
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
      TabIndex        =   41
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
      TabIndex        =   40
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
      TabIndex        =   32
      Top             =   6840
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
      Left            =   6030
      TabIndex        =   26
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label SynoloG 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         Size            =   9.75
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
         Size            =   9.75
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
Const F_mikta = 10

Dim f_CHECK_PARAGG As Integer

Dim fbuff As String

Dim XROMATA(100, 25, 25) 'κρατάει σε πίνακα τα χρώματα
Dim f_YPOK As Integer ' 0=kentriko  1=1o ypok  2= 2o ypok  3= 3o ypok
Dim F_YPOL_GEN  ' NA EMFANIZEI TA EIDH ASXETA AN EXOYN KINHΘΕΙ
Dim f_loadpar1 As Integer '1 FORTONEI THN FORMA SYNEXEIA
Dim F_TRANS As Long ' 1=BEGINTRANS   - ROLLBACK      0=NOTHING
Dim F_EIKONA As Long '
Dim f_mono_apoParaggelia As Integer ' mono apo Paraggelia TIMOLOGO
Dim F_POLHTES As Integer
Dim f_SYMPYKNVSH As Integer ' ΣΥΜΠΥΚΝΩΣΗ ΚΩΔΙΚΩΝ

Dim F_METRHTA As Integer ' ZHTA METPHTA
Dim F_METRHTAposo As Single ' ZHTA METPHTA

Dim F_DEF_CAR As Long ' DEFAULT AYTOKINHTO



Dim f_AkyrParas As String
Dim F_AR_parametroi(200, 2) ' ΦΟΡΤΩΝΩ ΤΙΣ ΠΑΡΑΜΕΤΡΟΥΣ ΣΕ ΠΙΝΑΚΑ ΓΙΑ ΝΑ ΜΗΝ ΑΡΓΕΙ ΣΤΟ ΞΕΚΙΝΗΜΑ ΤΗΣ ΦΟΡΜΑΣ

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

Dim f_Timol2_SeiresTimologioy 'POSA EIDH MPORO NA PERASO sthn forma timol2



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
Dim f_Akyr As Boolean

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

Dim F_akyr_Delete '  σβηνω το παραστατικό που ακυρώνεται

Dim f_use_Sysk  As Integer ' χρησιμοποιω συσκευασια = 1

Dim F_LOGPLHR(20) As String ' ΛΟΓΑΡΙΑΣΜΟΙ ΤΡΟΠΟΩΝ ΠΛΗΡΩΜΗΣ


Dim f_mfpa(1 To 9) As Single 'οριζονται στο find_parastat

Dim f_kau(1 To 9) As Single 'καθαρή αξία κατά φπα
Dim f_gfkod(20) As String, f_gfaj(20) As Single

'METABLHTES EIDOYS PARASTATIKOY APO PARASTAT
Dim F_TITLOS, F_POS_APOU, F_STADIO, F_AJIA_APOU, F_XONDR, f_pel, f_metasx, f_pol, F_rec_p As Integer, F_rec_m As Integer 'παράμετροι του παραστατικού

Dim f_deltia(1 To 100) 'δελτια που σημαδεύω
Dim f_IDdeltia(1 To 100) As Long  'ID δελτιων  που σημαδεύω για να μπορώ εύκολα να τα μαρκάρω

Dim F_File_Timologioy As String ' το όνομα του αρχείου που γίνεται η εκτύπωση π.χ. c:\print

Dim F_Type_File_Timologioy As Integer '1=typvnei to arxeio 0=den to typonei 2=me entolh print 3=c:\typ.bat

Dim F_File2_Timologioy As String ' το όνομα του αρχείου που γίνεται η εκτύπωση π.χ. c:\print
Dim F_Type2_File_Timologioy As Integer '1=typvnei to arxeio 0=den to typonei

Dim f_F12 As Integer, f_F1 As Integer  ' ΠΡΟΓΡΑΜΜΑΤΙΣΜΟΣ ΠΛΗΚΤΡΩΝ

Dim F_EXO_XROMATA As Integer  '1=ΧΡΗΣΙΜΟΠΟΙΩ ΤΑ ΤΡΩΜΑΤΑ ΜΕΓΕΘΗ

Dim F_PAR_MONO_ME_YPOLOIPO As Integer  '
'Dim F_EIDH_APOT As Integer ' να δειχνει μόνο τα είδη που εχουν κινηθεί στο υποκατάστημα

Dim f_Fkey ' το πληκτρο που κάνει update

Dim f_Desmia As Integer ' χρησιμοποιω τις παραγγελίες για έλεγχο υπολοίπου  'βλεπε mdesmia


Dim f_pelMERC As String ' ποιος πελατης του mercury χρησιμοποιει το πρόγραμμα



Dim f_suma ' teliko poso timologioy
Dim f_ColUpdate  ' flag αν ενημερώθηκε η κολόνα του dbgrid1
'Dim f_epitr_seir
Dim fDB As Database  ' ανοιγω γενικά την database και με την open_data ανοίγω τον πίνακα
Dim ProthFora As Boolean
Dim fSHMADEMENA As String
Dim mEidos_Par(50) 'ειδος παραστατικού . π.χ.  atim=T



Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)


'dim constant a

'

Private Sub b_akyr_Click()
Dim ANS As Integer
Dim k
Dim L

ANS = vbYes
If Len(Grid1.TextMatrix(1, 1)) > 0 Then
  ANS = MsgBox("Να χαθούν οι αλλαγές;", vbYesNo)
End If

 If ANS = vbYes Then
 

  '--------------------  initialization ----------------------------
  Synolo.Caption = ""
  sfpa.Caption = ""
  SynoloG.Caption = ""
  For k = 1 To 9: f_kau(k) = 0: SumFpa(k) = 0: Next ' συνολο ΦΠΑ ΚΑΤΑ ΣΥΝΤΕΛΕΣΤΗ
  For k = 1 To 100: f_fpa(k) = 0: f_sajia(k) = 0: Next ' κρατά στην μνήμη τοσύνολο κάθε σειράς και το ΦΠΑ
  For k = 1 To 30: f_IDdeltia(k) = 0: f_deltia(k) = "": Next 'δελτια που σημαδεύω
  ' μηδενιζω το grid1
  For L = 1 To Grid1.ColS - 1
     'Grid1.Col = l
     For k = 1 To Grid1.ROWS - 1
        'Grid1.Row = K
        'Grid1.Text = ""
        Grid1.TextMatrix(k, L) = ""
     Next
  Next

  Text2(0).Enabled = True
  Text2(1).Enabled = True
  Text2(2).Enabled = True
  Text2(0).Text = "" ' κωδ
  Text2(1).Text = "" 'επωνυμια
  Text2(2).Text = "" 'αφμ
  parat.Text = "" 'ΠΑΡΑΤΗΡΗΣΕΙΣ
  TIMText3.Top = 18000
  F_DIOR = 0 ' ΔΙΟΡΘΩΣΗ
  metasxhmatismos.Enabled = True
      f_loadpar1 = Val(FindParametroi("PAR1", "F_LOADPAR1", "1", "ΦΟΡΤΩΝΩ ΚΑΘΕ ΦΟΡΑ ΤΗΝ ΦΟΡΜΑ ΠΑΡΑΣΤΑΤΙΚΩΝ=1"))
    If f_loadpar1 = 1 Then
      Unload Me
    Else
      Me.Hide
    End If
    
      
      SendKeys "%"
      SendKeys "{RIGHT}+{RIGHT}+{RIGHT}+{DOWN}"
 End If
  
  
End Sub
Private Sub b_Katax_Click()
'====================================================


Dim wrkDefault As Workspace
Dim R As New ADODB.Recordset
Dim X, DUM
Dim F_ATIM
Dim k, L
Dim eidoskpe
Dim m_tim, S_POSOTHTA As Single
Dim par As New ADODB.Recordset, m_r
Dim m_m
Dim sql As String
Dim s1 As Single





'βρισκει το είδος του παραστατικού
'γιατί πολλές φορές το πουλάει σαν λιανική ενώ είναι χονδρική πώληση
 find_eid_parastat


 

 'X = mEidos_Par(1 + PARAS.ListIndex)
'Text1.Text = find_parastat(k, X, 0)





f_Akyr = False





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

For k = 1 To Grid1.ROWS - 1
   S_POSOTHTA = S_POSOTHTA + Abs(Val(Grid1.TextMatrix(k, f_p)))
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



If F_TRANS = 1 Then
  Gdb.BeginTrans
End If








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
    

          
       
       
       
          f_Akyr = True
          parat.Text = "Σχ.Παρ. " + f_matim
          F_PARAS = "ΑΚΥΡΩΤΙΚΟ ΣΤΟΙΧΕΙΟ  κ"
          
          
          'αν δουλεύει με terminal services και είναι το υποκατάστημα
          'τότε να κοβει ακυρωτικό από αλλη σειρά
          'If f_YPOK = 1 And Len(Dir("c:\BSEIRA", vbDirectory)) > 2 Then
          '   Mf_matim = "k" + Format(find_parastat(1, "k", 0), String(f_psifia_atim, "0"))
          'Else
             Mf_matim = f_AkyrParas + Format(find_parastat(1, f_AkyrParas, 0), String(f_psifia_atim, "0"))
         ' End If
          
          ' ELEGXOS AN YPARXEI HDH TO AKYROTIKO
          R.Open "select ATIM from TIM where ATIM='" + Mf_matim + "'", Gdb, adOpenDynamic, adLockOptimistic
          If Not R.EOF Then
             MsgBox "Υπάρχει ήδη  TO ΑΚΥΡΩΤΙΚΟ ΜΕ  ίδιο αριθμό. " + Mf_matim
             MsgBox "ΔΙΑΓΡΑΨΤΕ ΤΟ ΠΑΛΙΟ ΑΚΥΡΩΤΙΚΟ ή ΑΛΛΑΞΤΕ ΤΗΝ ΑΡΙΘΜΗΣΗ ΤΟΥ ΑΚΥΡΩΤΙΚΟΥ ΑΠΟ ΤΟ ΠΑΡΑΜΕΤΡΟΙ ΠΑΡΑΣΤΑΤΙΚΩΝ "
             Me.MousePointer = vbNormal
             Exit Sub
          End If
          
       Else
         f_Akyr = False
       End If
       
       
       
       
       
       If f_Akyr = True And F_akyr_Delete = 0 Then
       'den σβηνω το  παραστατικό που ακυρωνεται
       
       Else
       
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
       
       End If
       
       
       ' SAN MATIM EXO TO AKYROTIKO KAI BAZO SIMERINI HMEROMHNIA
       If f_Akyr = True Then
          f_matim = Mf_matim
          Text1.Text = Val(mID$(Mf_matim, 2, f_psifia_atim))
          f_dat = Now
          ' ΜΗΔΕΝΙΖΩ ΑΥΤΕΣ ΤΙΣ ΠΑΡΑΜΕΤΡΟΥΣ ΓΙΑΤΙ
          'ΘΑ ΠΑΡΕΙ ΤΙΣ ΠΑΡΑΜΕΤΡΟΥΣ ΤΟΥ ΑΡΧΙΚΟΥ ΠΑΡΑΣΤΑΤΙΚΟΥ ΚΑΙ ΘΑ ΧΑΛΑΣΕΙ ΤΗΝ ΑΠΟΘΗΚΗ
          F_POS_APOU = "0"
          F_AJIA_APOU = "0"
          f_pel = "0"

       End If
       
Else  'DIOR=0 ΔΕΝ ΕΙΝΑΙ ΔΙΟΡΘΩΣΗ
 
 
  'diadikasia egkrishs
   '----------------------------------------------------------------
   If f_FORM_EKTY > 0 And Val(SynoloG) > 20 And Val(SynoloG) < 10000 And Combo4.Text = Combo4.List(1) And Len(Dir("c:\mercvb\egkr.txt")) > 1 Then '<> METRHTOIS ή κάρτα
       Dim ccc As String, NNN, date1, Topos
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
       
       NNN = add_schedule(date1, Text2(1).Text, Topos)
       Dim SQLMEMO
    If IsNull(PEL.Recordset("MEMO")) Then
       SQLMEMO = Format(date1, "DD/MM/YYYY") + " " + Topos + "' WHERE EIDOS='e' and KOD='" + PEL.Recordset("KOD")
    Else
       SQLMEMO = PEL.Recordset("MEMO") + Chr(13) + Format(date1, "DD/MM/YYYY") + " " + Topos + "' WHERE EIDOS='e' and KOD='" + PEL.Recordset("KOD")
    End If
    
    Gdb.Execute "UPDATE PEL SET MEMO='" + SQLMEMO + "'"
    
'      'ενημερωση με τον τελευταίο αριθμό
'   ' γιατί μπορεί ενδιάμεσα να κόπηκε και αλλο παραστατικό
'    If F_DIOR = 0 And f_FORM_EKTY > 0 Then  ' οχι ΔΙΟΡΘΩΣΗ
'      'ΓΙΑ ΝΑ ΠΙΑΝΕΙ ΤΟΝΣΩΣΤΟ ΑΡΙΘΜΟ ΑΛΛΑ ΔΗΜΙΟΥΡΓΗΣΕ ΠΡΟΒΛΗΜΑ  5-10-2009 ΒΑΖΟΝΤΑΣ ΤΟΝ ΙΔΙΟ ΑΡΙΘΜΟ ΣΕ ΠΟΛΛΑ ΠΑΡΑΣΤΑΤΙΚΑ
'      '30/7/2010
'       Text1.Text = find_parastat(1, Right(PARAS.Text, 1), 0)
'      f_matim = Right(PARAS.Text, 1) + Right("00000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text
'    End If
'
'    If Check_Dipla = 0 Then
'        If F_TRANS = 1 Then
'           Gdb.RollbackTrans
'        End If
'        Exit Sub
'    End If
   
 End If  '---------------------------------------------------------------
 
         'ενημερωση με τον τελευταίο αριθμό
   ' γιατί μπορεί ενδιάμεσα να κόπηκε και αλλο παραστατικό
    If F_DIOR = 0 And f_FORM_EKTY > 0 Then  ' οχι ΔΙΟΡΘΩΣΗ
      'ΓΙΑ ΝΑ ΠΙΑΝΕΙ ΤΟΝΣΩΣΤΟ ΑΡΙΘΜΟ ΑΛΛΑ ΔΗΜΙΟΥΡΓΗΣΕ ΠΡΟΒΛΗΜΑ  5-10-2009 ΒΑΖΟΝΤΑΣ ΤΟΝ ΙΔΙΟ ΑΡΙΘΜΟ ΣΕ ΠΟΛΛΑ ΠΑΡΑΣΤΑΤΙΚΑ
      '30/7/2010
       Text1.Text = find_parastat(1, Right(PARAS.Text, 1), 0)
      f_matim = Right(PARAS.Text, 1) + Right("00000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text
    End If
   
    If Check_Dipla = 0 Then
        If F_TRANS = 1 Then
           Gdb.RollbackTrans
        End If
        Exit Sub
    End If
 
    
   
   
End If

30
   ' Get default Workspace.
  ' Set wrkDefault = DBEngine.Workspaces(0)
' Start of outer transaction.

'MEXRI TIS 21/6/2010 DEN HTAN REM OI PARAKATO 3 SEIRES
'META TIS 21/6/2010  METAFERUHKAN OI SEIRES STHN ARXH
'If F_TRANS = 1 Then
'  Gdb.BeginTrans
'End If

If Len(f_deltia(1)) > 0 Then
   Dim S As Integer
   k = 1
   Do While f_IDdeltia(k) > 0
     Gdb.Execute "update TIM set ART='1',SXETIKO='" + f_matim + " " + Format(f_dat, "dd/mm/yyyy") + "' where ID_NUM=" + Str(f_IDdeltia(k))
     k = k + 1
   Loop
   
End If


'δημιουργία eggtim και ΤΙΜ
40 enhm_eggtim





If F_TRANS = -1 Then  'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
   F_TRANS = 1 ' για να μπορει να ξαναπαίξει
   Me.MousePointer = vbNormal
   Exit Sub
End If





If f_Akyr = False Then  ' αν δεν ειναι ακυρωτικό

'enhm_eggtim

        'ενημέρωση αποθήκης
        
50     enhm_apot

If F_TRANS = -1 Then  'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
   F_TRANS = 1 ' για να μπορει να ξαναπαίξει
   Me.MousePointer = vbNormal
   Exit Sub
End If


        'ενημέρωση πελάτη
60         enhm_pel

If F_TRANS = -1 Then  'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
   F_TRANS = 1 ' για να μπορει να ξαναπαίξει
   Me.MousePointer = vbNormal
   Exit Sub
End If

F_METRHTAposo = 0
If F_METRHTA = 1 Then
   F_METRHTAposo = InputBox("ΜΕΤΡΗΤΑ")
End If

   







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
If F_TRANS = -1 Then  'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
   F_TRANS = 1 ' για να μπορει να ξαναπαίξει
   Me.MousePointer = vbNormal
   Exit Sub
End If




Dim rTemp As New ADODB.Recordset
Dim SQLt As String, TYP As Single
Dim N22 As Long
'ENHMEROSI PEL ME TYP
SQLt = "SELECT KOD,"
SQLt = SQLt + "SUM(CASE WHEN XREOSI IS NULL  THEN 0 ELSE XREOSI  END )  -"
 SQLt = SQLt + " SUM(CASE WHEN PISTOSI IS NULL  THEN 0 ELSE PISTOSI  END )   AS YP"
 SQLt = SQLt + " FROM EGG  where HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND EIDOS='" + PEL.Recordset("EIDOS") + "' AND KOD='" + PEL.Recordset("kod") + "' GROUP BY KOD"
rTemp.Open SQLt, Gdb, adOpenDynamic, adLockOptimistic
If IsNull(rTemp("YP")) Then
   TYP = 0
Else
   TYP = rTemp("YP")
End If
Gdb.Execute "UPDATE PEL SET TYP=AYP+ " + Format(TYP, "###########0.00") + " WHERE  EIDOS='" + PEL.Recordset("EIDOS") + "' AND KOD='" + PEL.Recordset("kod") + "'", N22

























If F_TRANS = -1 Then  'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
   F_TRANS = 1 ' για να μπορει να ξαναπαίξει
   Me.MousePointer = vbNormal
   Exit Sub
End If





End If  ' If f_AKYR = False Then  ' αν δεν ειναι ακυρωτικό


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


'ΜΑΡΚ’ΡΩ ΤΑ ΔΕΛΤΙΑ ΠΟΥ ΧΡΗΣΙΜΟΠΟΙΗΣΑ
If Len(fSHMADEMENA) > 5 Then
  Gdb.Execute "UPDATE TIM set ART='1' where ATIM+CONVERT(CHAR(10),hme,103)  in (" + fSHMADEMENA + ");"
End If



'If F_TRANS = 1 And (APOTH1.Text <> APOTH1.List(0)) Then

        ' Ask if the user wants to commit to all the changes
        ' made above.
      '  If MsgBox("Nα αποθηκευθεί το παραστατικό", vbYesNo) = vbYes Then
             
'       On Error GoTo commitok
       
 '            Gdb.CommitTrans
      '  Else
           ' Gdb.RollbackTrans
           ' Me.MousePointer = vbNormal
           ' Exit Sub
      '  End If
   '    On Error Resume Next
  '
'End If

If F_TRANS = 1 Then
   Gdb.CommitTrans
End If



90


   If Left(f_pelMERC, 3) = "SOK" Then
       
       Dim KSOK
       
      'Kill "C:\KENO.TXT"
     ' MILSEC 1000
       
      ' Open "c:\KENO.txt" For Append As #112
          
          
       '    Print #112, "   "
        '   Print #112, "   "
         '  Print #112, "   "
           
   '     For KSOK = 1 To 10
    '       Print #112, Grid1.TextMatrix(KSOK, f_k)  '   GRAFEI TON KODIKO GIA EPISTROFES
     '   Next
   
      ' Close #112
   
   
   
   End If







'εκτύπωση
If ektyp.value = 1 Then ' ---------------------------------------------


If f_FORM_EKTY >= 90 Then ' φορμα τύπου f90.txt  '====================================
100    DUM = printCrystal(f_matim, CDate(DTPicker1.value)) 'δημιουργει τα προσωρινά αρχεία
110   DUM = print_timol(f_matim, CDate(DTPicker1.value))



   










Else ' crystal reports timol1.rpt(σφραγίδα)   timol1an.rpt (αντίγραφο)
   
   If f_FORM_EKTY > 0 Then '      PEL.Recordset("eidos") = "e" Then  ' EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE


 'ΑΝ ΧΡΗΣΙΜΟΠΟΙΕΙ ΤΗΝ ΦΟΡΜΑ ΤΙΜΟΛ2 TOTE NA DINEI ALLES SEIRES APO TO F_FORM2_SEIRESTIMOL
       If f_FORM_EKTY = 14 Or f_FORM_EKTY = 12 Then
           f_SeiresTimologioy = f_Timol2_SeiresTimologioy
           f_SynSeiresTimologioy = f_Timol2_SeiresTimologioy
       End If








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
    
      ElseIf f_FORM_EKTY = 14 Then     '    1 ENTYPO
       CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol2.rpt"
    
    
    
    
    ElseIf f_FORM_EKTY = 15 Then     '     λιανικη
       CrystalReport1.ReportFileName = "C:\Mercvb\reports\timolsql.rpt"
       CrystalReport1.Connect = gConnect
    
    Else
       CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
    End If
     
    
    
Dim nnf As Integer

    
If proepiskopisi.value = vbChecked Then
  ' nnf = crptToWindow
   CrystalReport1.Destination = crptToWindow
Else

   CrystalReport1.Destination = crptToPrinter
End If
    
    
    
    
    If f_FORM_EKTY = 10 Then
       CrystalReport1.Action = 1
    
    
    ElseIf f_FORM_EKTY = 14 Then
       CrystalReport1.Action = 1
    ElseIf f_FORM_EKTY = 20 Then
    
       CrystalReport1.Destination = crptToWindow
       CrystalReport1.Action = 1
    ElseIf f_FORM_EKTY = 18 Then
       CrystalReport1.Action = 1
       CrystalReport1.Action = 1
       CrystalReport1.Action = 1
    ElseIf f_FORM_EKTY = 13 Then
        CrystalReport1.Action = 1
    
    ElseIf f_FORM_EKTY = 23 Then
       CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
       CrystalReport1.Action = 1
       
       CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol2.rpt"
       CrystalReport1.Action = 1
    
    
    Else
       CrystalReport1.Action = 1
       CrystalReport1.Action = 1
    End If
 
 
 End If   ' If f_FORM_EKTY > 0 Then '  EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
   
End If  ' If ektyp.Value = 1 Then  '========================================



End If  '-------------------------------------------------------------



'ENHMERVSH ARIUMOY PARASTATIKOY
If f_Akyr Then ' ΑΝΗΜΕΡΩΣΗ ΑΚΥΡΩΤΙΚΟΥ
     DUM = find_parastat(1, Left(f_matim, 1), 1)
Else
   If F_DIOR = 1 Then ' ΔΙΟΡΘΩΣΗ   DEN ΕΝΗΜΕΡΩΝΕΙ ΑΡΙΘΜΗΤΗΡΑ
   Else
        Upd_Parastat
   End If
End If






   If Left(f_pelMERC, 3) = "SOK" Then
       If f_FORM_EKTY = -1 Then ' φο
        DUM = Shell("c:\TYP.BAT ", vbMinimizedNoFocus)
            
       End If
   End If





















'ENHMERVSH ARIUMOY PARASTATIKOY
'Upd_Parastat
'    b_Katax.Enabled = False
    
             
'-----------------------------------------------------------------
'--------------------  initialization ----------------------------
'-----------------------------------------------------------------
145
MIDENIZO_GRID
metasx.value = vbUnchecked
F_DIOR = 0 ' ΔΙΟΡΘΩΣΗ

LASTTIMOL.Refresh




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
            
            
commitok:
     Gdb.RollbackTrans
     MsgBox "προσοχη δεν αποθηκεύθηκε το παραστατικό"
     Me.MousePointer = vbNormal
     Exit Sub
            
           
            
            
            
            

End Sub
Function egkrisi(ByVal ccc As String) As Long
     egkrisi = Val(mID(ccc, 1, 1)) + Val(mID(ccc, 2, 1)) + Val(mID(ccc, 3, 1)) + Val(mID(ccc, 4, 1))
     egkrisi = Int(egkrisi / 2) + 135
End Function

Sub MIDENIZO_GRID()
'===================================================================
Dim k As Long
Dim L As Long

Synolo.Caption = ""
sfpa.Caption = ""
SynoloG.Caption = ""
         


         
For k = 1 To 9: f_kau(k) = 0: SumFpa(k) = 0: Next ' συνολο ΦΠΑ ΚΑΤΑ ΣΥΝΤΕΛΕΣΤΗ

For k = 1 To 100: f_fpa(k) = 0: f_sajia(k) = 0: Next ' κρατά στην μνήμη τοσύνολο κάθε σειράς και το ΦΠΑ

For k = 1 To 30: f_IDdeltia(k) = 0: f_deltia(k) = "": Next 'δελτια που σημαδεύω


' μηδενιζω το grid1
For L = 1 To Grid1.ColS - 1
    'Grid1.Col = l
    For k = 1 To Grid1.ROWS - 1
       'Grid1.Row = K
       'Grid1.Text = ""
       Grid1.TextMatrix(k, L) = ""
    Next
Next

  Text2(0).Enabled = True
  Text2(1).Enabled = True
  Text2(2).Enabled = True
  

Text2(0).Text = "" ' κωδ
Text2(1).Text = "" 'επωνυμια
Text2(2).Text = "" 'αφμ
parat.Text = "" 'ΠΑΡΑΤΗΡΗΣΕΙΣ


TIMText3.Top = 18000
APOTH2.Enabled = True
'===================================================================
End Sub



Sub elegxoi()
'================================================================
'=======   DEMO TA  TRANSACTIONS  ===============================
'================================================================
'
'   Dim strName As String
'   Dim strMessage As String
'   Dim wrkDefault As Workspace
'   Dim dbsNorthwind As Database
'   Dim rstEmployees As Recordset
'
'   ' Get default Workspace.
'   Set wrkDefault = DBEngine.Workspaces(0)
'   Set dbsNorthwind = OpenDatabase("Northwind.mdb")
'   Set rstEmployees = _
'      dbsNorthwind.OpenRecordset("Employees")
'
'   ' Start of outer transaction.
'   wrkDefault.BeginTrans
'   ' Start of main transaction.
'   wrkDefault.BeginTrans
'
'   With rstEmployees
'
'      ' Loop through recordset and ask user if she wants to
'      ' change the title for a specified employee.
'      Do Until .EOF
'         If !title = "Sales Representative" Then
'            strName = !LastName & ", " & !FirstName
'            strMessage = "Employee: " & strName & vbCr & _
'               "Change title to Account Executive?"
'
'            ' Change the title for the specified employee.
'            If MsgBox(strMessage, vbYesNo) = vbYes Then
'               .Edit
'               !title = "Account Executive"
'               .Update
'            End If
'         End If
'
'         .MoveNext
'      Loop
'
'      ' Ask if the user wants to commit to all the changes
'      ' made above.
'      If MsgBox("Save all changes?", vbYesNo) = vbYes Then
'         wrkDefault.CommitTrans
'      Else
'         wrkDefault.Rollback
'      End If
'
'      ' Print current data in recordset.
'      .MoveFirst
'      Do While Not .EOF
'         Debug.Print !LastName & ", " & !FirstName & _
'            " - " & !title
'         .MoveNext
'      Loop
'
'      ' Roll back any changes made by the user since this is
'      ' a demonstration.
'      wrkDefault.Rollback
'      .Close
'   End With
'
'   dbsNorthwind.Close
'
''End Sub
'
'

'1.να μην υπάρχουν ΦΠΑ>9 και <=0
End Sub
Sub Upd_Parastat()

Dim k As Integer
Dim db As Database
Dim R As Recordset, X As String
Dim DUM As Integer


For k = 0 To SEIRA.ListCount
   If SEIRA.Text = SEIRA.List(k) Then
      Exit For
   End If
Next


If k <= SEIRA.ListCount Then
   DUM = find_parastat(k, Right(F_PARAS, 1), 1)
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

Dim k, L, M, N As Integer
Dim MHKOS As Integer
Dim mPoso, mKAU_AJIA, mMIK_AJIA, mProeleysh
Dim mXre, mPIS, mONOMA, mKERDOS
Dim sql As String
Dim mLenOnoma


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







For k = 1 To 9: f_kau(k) = 0: Next
  
  
    'DB.Open gDir
'  End If
10
 re.Open "SELECT top 1 *FROM EGGTIM", Gdb, adOpenDynamic, adLockOptimistic
 rt.Open "SELECT top 1 *FROM TIM", Gdb, adOpenDynamic, adLockOptimistic
  
Dim mLenPROORISMOS, mLenAYTOKINHTO, mLenSKOPOS, mLenFORTOSH, mLenPARAT
mLenPROORISMOS = rt("proor").DefinedSize
mLenAYTOKINHTO = rt("aytok").DefinedSize
mLenSKOPOS = rt("skopos").DefinedSize

mLenFORTOSH = rt("fortosh").DefinedSize
mLenPARAT = rt("parat").DefinedSize
mLenOnoma = re("onoma").DefinedSize


  
  
  
  
  
  
  
Dim SYN_ME_FPA As Single
5 re("MIK_AJIA") = 0

For k = 1 To 10: FEGG(k) = 0: Next
'------------------------------eggtim -----------------------------------------
For k = 1 To Grid1.ROWS - 1

   If Len(Grid1.TextMatrix(k, f_k)) > 0 Then

        meid.Open "SELECT * FROM EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) + "'", Gdb, adOpenDynamic, adLockOptimistic 'adOpenForwardOnly, adLockReadOnly

    
    'ΠΡΟΣΘΕΤΕΙ ΜΙΑ ΕΓΓΡΑΦΗ ΓΙΑ ΚΑΘΕ ΧΡΩΜΑ ή ΜΕΓΕΘΟΣ
    If F_EXO_XROMATA = 0 Then
        GoSub ADD_RECORD
    Else
         Dim K2 As Integer, k3 As Integer
'         For K2 = 1 To 25: For k3 = 1 To 25
'           If Val(XROMATA(k, K2, k3)) <> 0 And Val(XROMATA(k, K2, k3)) <> -0.01 Then
'              GoSub ADD_RECORD
'              XROMATA(k, K2, k3) = 0
'           End If
'         Next: Next
         
         
         'ΣΤΗΝ ΔΙΟΡΘΩΣΗ ΔΙΟΡΩΝΕΙ RECORD-RECORD
         If F_DIOR = 1 Then
         
                GoSub ADD_RECORD
         
         Else
           For K2 = 1 To 25: For k3 = 1 To 25
               If Val(XROMATA(k, K2, k3)) <> 0 And Val(XROMATA(k, K2, k3)) <> -0.01 Then
                  GoSub ADD_RECORD
                  XROMATA(k, K2, k3) = 0
               End If
           Next: Next
         End If
         
         
         
         
              
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
SQLTIM = SQLTIM + "PARAT,METAF,SKOPOS,PROOR,FORTOSH,AYTOK,KERDOS,KLEIDI) VALUES ("
SQLTIM = SQLTIM + "'" + Format(f_dat, "MM/DD/YYYY") + "',"  ' HME
SQLTIM = SQLTIM + "'" + Left(f_matim, f_psifia_atim + 2) + "'," 'ATIM
f_suma = 0
For k = 1 To 9
   SQLTIM = SQLTIM + Str(Round(SumFpa(k), f_psifiaAjias)) + "," 'AJ1,aj2..aj9
   f_suma = f_suma + Round(SumFpa(k), f_psifiaAjias)
Next
For k = 1 To 9
   If k <> 5 Then
     SQLTIM = SQLTIM + Str(Round(SumFpa(k) * g_Fpa(k) / 100, f_psifiaAjias)) + "," 'fpa1,fpa2,fpa3..fpa9
     f_suma = f_suma + Round(SumFpa(k) * g_Fpa(k) / 100, f_psifiaAjias)
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


'On Error GoTo 0
If Len(parat.Text) > mLenPARAT Then parat.Text = Left(parat.Text, mLenPARAT)
If Len(SKOPOS.Text) > mLenSKOPOS Then parat.Text = Left(SKOPOS.Text, mLenSKOPOS)
If Len(FORTOSH.Text) > mLenFORTOSH Then FORTOSH.Text = Left(FORTOSH.Text, mLenFORTOSH)
If Len(PROORISMOS.Text) > mLenPROORISMOS Then PROORISMOS.Text = Left(PROORISMOS.Text, mLenPROORISMOS)
If Len(AYTOKINHTO.Text) > mLenAYTOKINHTO Then AYTOKINHTO.Text = Left(AYTOKINHTO.Text, mLenAYTOKINHTO)



SQLTIM = SQLTIM + "'" + Replace(parat.Text, "'", "`") + "'," ' PARAT
SQLTIM = SQLTIM + Str(Val(Replace(metaf.Text, ",", "."))) + "," ' METAF
SQLTIM = SQLTIM + "'" + Replace(SKOPOS.Text, "'", "`") + "'," ' SKOPOS
SQLTIM = SQLTIM + "'" + Replace(PROORISMOS.Text, "'", "`") + "'," ' PROORISMOS
SQLTIM = SQLTIM + "'" + Replace(FORTOSH.Text, "'", "`") + "'," ' FORTOSH
SQLTIM = SQLTIM + "'" + Replace(AYTOKINHTO.Text, "'", "`") + "'," ' AYTOKINHTO
SQLTIM = SQLTIM + Str(get_kerdos) + ","  ' KERDOS

If f_pol = "1" Then
   SQLTIM = SQLTIM + "'" + Left(f_matim, f_psifia_atim + 2) + "')" ' KEY
Else
   SQLTIM = SQLTIM + "'" + Left(f_matim, f_psifia_atim + 2) + PEL.Recordset("kod") + "')" ' KEY
End If


Dim NTIM As Long




Dim rrr As New ADODB.Recordset

      rrr.Open "select COUNT(*)  from TIM where ATIM='" + f_matim + "' and HME='" + Format(f_dat, "MM/DD/YYYY") + "'", Gdb, adOpenDynamic, adLockOptimistic
  
  If F_TRANS = 1 And rrr(0) > 0 Then
     Gdb.RollbackTrans
     MsgBox "Υπάρχει ήδη το παραστατικό " + f_matim + Chr(13) + " δεν κατεχωρήθη"
     F_TRANS = -1 'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
     Exit Sub
  End If
     
     
Gdb.Execute SQLTIM, NTIM


If F_TRANS = 1 And NTIM = 0 Then
    Gdb.RollbackTrans
    MsgBox "Λάθος στην επικεφαλίδα του παραστατικού" + Chr(13) + " δεν κατεχωρήθη"
     F_TRANS = -1 'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
     Exit Sub
End If


rt.Close
re.Close

For k = 1 To f_SeiresTimologioy
   gm_str(k) = ""
Next



Exit Sub

'----------------------------------------------------------------------------
ADD_RECORD:

'20      re.AddNew
'        re("prood_aj") = k
'        re("hme") = f_dat
'        re("atim") = f_matim
'        ' re("mona") = n
'        re("ekpt") = Val(Grid1.TextMatrix(k, f_e))
'        re("timm") = Val(Grid1.TextMatrix(k, F_T))
'         If F_EXO_XROMATA = 0 Then
'            re("poso") = Val(Grid1.TextMatrix(k, f_p))
'            re("kau_ajia") = Round(Val(Grid1.TextMatrix(k, f_a)), f_psifiaAjias)
'            re("MIK_AJIA") = Round(Val(Grid1.TextMatrix(k, f_a)), f_psifiaAjias) * (1 + g_Fpa(Val(Grid1.TextMatrix(k, f_f))) / 100)
'         Else
'
'
'        If XROMATA(k, 0, 0) = -0.01 Then    ' ΣΗΜΑΔΙ ΟΤΙ ΧΡΗΣΙΜΟΠΟΙΗΣΕ ΤΟ F6
'               re("poso") = XROMATA(k, K2, k3)
'               re("PROELEYSH") = Left(XROMATA(k, K2, 0), 3) + Replace(Left(XROMATA(k, 0, k3), 2), "*", "") + "*" + Format(K2, "00") + "-" + Format(k3, "00")
'        Else
'               re("poso") = Val(Grid1.TextMatrix(k, f_p))
'               'ΑΥΤΑ ΠΟΥ ΕΡΧΟΝΤΑΙ ΜΕ BARCODE TRABANE TIS STHLES ΑΠ ΕΥΘΕΙΑΣ ΑΠΟ ΤΟ TABLE BARCODES TO OPOIO TO TSIMPAO OTAN TO BREI KAI TO BAZO STO XROMATA(k, K2, 0)
'               re("PROELEYSH") = Left(XROMATA(k, K2, 0), 3) + Replace(Left(XROMATA(k, 0, k3), 2), "*", "") + "*" + mID(XROMATA(k, K2, 0), 4, 4)   '+ Format(K2, "00") + "-" + Format(k3, "00")
'        End If
'
'            re("kau_ajia") = Round(re("POSO") * re("TIMM") * (100 - re("EKPT")) / 100, f_psifiaAjias)
'            re("MIK_AJIA") = Round(re("POSO") * re("TIMM") * (100 - re("EKPT")) / 100, f_psifiaAjias) * (1 + g_Fpa(Val(Grid1.TextMatrix(k, f_f))) / 100)
'         End If
'
'
'        ' ΕΝΗΜΕΡΏΝΕΙ ΤΟ ΚΈΡΔΟς
'        ' meid.Open "SELECT * FROM EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) + "'", Gdb, adOpenDynamic, adLockOptimistic 'adOpenForwardOnly, adLockReadOnly
'         If PEL.Recordset("EIDOS") = "e" And Val(Grid1.TextMatrix(k, F_T)) > 0 Then
'            re("KERDOS") = Val(Grid1.TextMatrix(k, F_T)) - meid("XTI")
'         End If
'         re("kode") = Grid1.TextMatrix(k, f_k)
'         re("onoma") = Left(Grid1.TextMatrix(k, f_o), 35)
'        ' ΒΑΖΕΙ ΤΗΝ ΣΥΣΚΕΥΑΣΙΑ ΜΕΣΑ ΣΤΟ ΟΝΟΜΑ ΓΙΑ ΤΙΜΟΛΟΓΗΣΗ


'        If f_use_Sysk = 1 Then
'          If meid("SYSKMAX") > 1 Then
'            MHKOS = Len(Trim(Format(meid("SYSKMAX"), "####") + "X" + Format(re("poso") / meid("SYSKMAX"), "####")))
'            re("onoma") = Left(Grid1.TextMatrix(k, f_o) + Space(35), 34 - MHKOS) + " " + Format(meid("SYSKMAX"), "####") + "X" + Format(re("poso") / meid("SYSKMAX"), "####")
'          End If
'        End If
'
'        re("mona") = Grid1.TextMatrix(k, F_M)  ' Right(Grid1.TextMatrix(K, f_o), 3)
'
'
'        SYN_ME_FPA = SYN_ME_FPA + re("MIK_AJIA")
'        re("fpa") = Val(Grid1.TextMatrix(k, f_f))
'        FEGG(re("FPA")) = FEGG(re("FPA")) + re("MIK_AJIA")
'
'        re("pelkod") = PEL.Recordset("kod")
'        re("eidos") = PEL.Recordset("eidos")
'
'
'         'ΤΙΜΗ ΜΟΝΑΔΟΣ ΣΕ ΞΕΝΟ ΝΟΜΙΣΜΑ
'        re("FCURRENCY") = gVal(isotimia.Text) * re("TIMM")
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
'--------------------------------------------------------------
mPoso = 0: mKAU_AJIA = 0: mMIK_AJIA = 0: mProeleysh = ""

If F_EXO_XROMATA = 0 Then
                
         mPoso = Val(Grid1.TextMatrix(k, f_p))
         mKAU_AJIA = Round(Val(Grid1.TextMatrix(k, f_a)), f_psifiaAjias)
         mMIK_AJIA = Round(Val(Grid1.TextMatrix(k, f_a)), f_psifiaAjias) * (1 + g_Fpa(Val(Grid1.TextMatrix(k, f_f))) / 100)
Else
        If XROMATA(k, 0, 0) = -0.01 Then    ' ΣΗΜΑΔΙ ΟΤΙ ΧΡΗΣΙΜΟΠΟΙΗΣΕ ΤΟ F6
              '  mPoso = XROMATA(k, K2, k3)
              '  mProeleysh = Left(XROMATA(k, K2, 0), 3) + Replace(Left(XROMATA(k, 0, k3), 2), "*", "") + "*" + Format(K2, "00") + "-" + Format(k3, "00")
              
              If F_DIOR = 1 Then
                  mPoso = Val(Grid1.TextMatrix(k, f_p)) 'XROMATA(k, K2, k3)
                  mProeleysh = Left(XROMATA(k, 1, 0), 3) + Replace(Left(XROMATA(k, 0, 1), 2), "*", "") + "*" + Format(K2, "00") + "-" + Format(k3, "00")
              Else
                  mPoso = XROMATA(k, K2, k3)
                  mProeleysh = Left(XROMATA(k, K2, 0), 3) + Replace(Left(XROMATA(k, 0, k3), 2), "*", "") + "*" + Format(K2, "00") + "-" + Format(k3, "00")
              End If
        Else
        
'            If F_DIOR = 1 Then
'               mPoso = Val(Grid1.TextMatrix(k, f_p))
'               'ΑΥΤΑ ΠΟΥ ΕΡΧΟΝΤΑΙ ΜΕ BARCODE TRABANE TIS STHLES ΑΠ ΕΥΘΕΙΑΣ ΑΠΟ ΤΟ TABLE BARCODES TO OPOIO TO TSIMPAO OTAN TO BREI KAI TO BAZO STO XROMATA(k, K2, 0)
'               mProeleysh = Left(XROMATA(k, K2, 0), 3) + Replace(Left(XROMATA(k, 0, k3), 2), "*", "") + "*" + mID(XROMATA(k, K2, 0), 4, 4)   '+ Format(K2, "00") + "-" + Format(k3, "00")
'
'
'
'            Else
'               mPoso = Val(Grid1.TextMatrix(k, f_p))
'               'ΑΥΤΑ ΠΟΥ ΕΡΧΟΝΤΑΙ ΜΕ BARCODE TRABANE TIS STHLES ΑΠ ΕΥΘΕΙΑΣ ΑΠΟ ΤΟ TABLE BARCODES TO OPOIO TO TSIMPAO OTAN TO BREI KAI TO BAZO STO XROMATA(k, K2, 0)
'               mProeleysh = Left(XROMATA(k, K2, 0), 3) + Replace(Left(XROMATA(k, 0, k3), 2), "*", "") + "*" + mID(XROMATA(k, K2, 0), 4, 4)   '+ Format(K2, "00") + "-" + Format(k3, "00")
'            End If
            
            
            If F_DIOR = 1 Then
               mPoso = Val(Grid1.TextMatrix(k, f_p))
               'ΑΥΤΑ ΠΟΥ ΕΡΧΟΝΤΑΙ ΜΕ BARCODE TRABANE TIS STHLES ΑΠ ΕΥΘΕΙΑΣ ΑΠΟ ΤΟ TABLE BARCODES TO OPOIO TO TSIMPAO OTAN TO BREI KAI TO BAZO STO XROMATA(k, K2, 0)
               mProeleysh = Left(XROMATA(k, 1, 0), 3) + Replace(Left(XROMATA(k, 0, 1), 2), "*", "") + "*" + mID(XROMATA(k, K2, 0), 4, 4)   '+ Format(K2, "00") + "-" + Format(k3, "00")
            Else
               mPoso = Val(Grid1.TextMatrix(k, f_p))
               'ΑΥΤΑ ΠΟΥ ΕΡΧΟΝΤΑΙ ΜΕ BARCODE TRABANE TIS STHLES ΑΠ ΕΥΘΕΙΑΣ ΑΠΟ ΤΟ TABLE BARCODES TO OPOIO TO TSIMPAO OTAN TO BREI KAI TO BAZO STO XROMATA(k, K2, 0)
               mProeleysh = Left(XROMATA(k, K2, 0), 3) + Replace(Left(XROMATA(k, 0, k3), 2), "*", "") + "*" + mID(XROMATA(k, K2, 0), 4, 4)   '+ Format(K2, "00") + "-" + Format(k3, "00")
            End If
        End If
         
         mKAU_AJIA = Round(mPoso * Val(Grid1.TextMatrix(k, F_T)) * (100 - Val(Grid1.TextMatrix(k, f_e))) / 100, f_psifiaAjias)
         mMIK_AJIA = mKAU_AJIA * (1 + g_Fpa(Val(Grid1.TextMatrix(k, f_f))) / 100)
End If












      ' ΒΑΖΕΙ ΤΗΝ ΣΥΣΚΕΥΑΣΙΑ ΜΕΣΑ ΣΤΟ ΟΝΟΜΑ ΓΙΑ ΤΙΜΟΛΟΓΗΣΗ
              If mLenOnoma > 35 Then
                 mONOMA = Left(Grid1.TextMatrix(k, f_o), 65) ' 24/8/2011  ήταν 45 και το εκανα 65 για ΟΣΣΑ ΡΩΞΑΝΗ
              Else
                 mONOMA = Left(Grid1.TextMatrix(k, f_o), 35)
              End If

        
        
        '   mONOMA = Left(Grid1.TextMatrix(k, f_o), 35)
        If f_use_Sysk = 1 Then
        
        
        
        
        
'        If f_use_Sysk = 1 Then
'         If meid("SYSKMAX") > 1 Then
'           MHKOS = Len(Trim(Format(meid("SYSKMAX"), "####") + "X" + Format(re("poso") / meid("SYSKMAX"), "####")))
'           re("onoma") = Left(Grid1.TextMatrix(k, f_o) + Space(35), 34 - MHKOS) + " " + Format(meid("SYSKMAX"), "####") + "X" + Format(re("poso") / meid("SYSKMAX"), "####")
'          End If
'        End If
        
        
        
        
          If meid("SYSKMAX") > 1 Then
            MHKOS = Len(Trim(Format(meid("SYSKMAX"), "####") + "X" + Format(mPoso / meid("SYSKMAX"), "####")))
              mONOMA = Left(Grid1.TextMatrix(k, f_o) + Space(35), 34 - MHKOS) + " " + Format(meid("SYSKMAX"), "####") + "X" + Format(mPoso / meid("SYSKMAX"), "####")
              If Left(f_pelMERC, 3) = "OLY" Then
                  mONOMA = Left(Grid1.TextMatrix(k, f_o), 35)
              End If
              
              
           Else
              If mLenOnoma > 35 Then
                 mONOMA = Left(Grid1.TextMatrix(k, f_o), 45)
              Else
                 mONOMA = Left(Grid1.TextMatrix(k, f_o), 35)
              End If
          End If
          
        End If
        mONOMA = Replace(mONOMA, "'", """")

        mXre = 0: mPIS = 0
  
  
If F_EXO_XROMATA = 0 Then
        Select Case Val(F_POS_APOU)
        Case 1 ' +EISAGOGES
           mXre = Val(Grid1.TextMatrix(k, f_p))  're("POSO")
        Case 2 ' -EISAGOGES
           mXre = -Val(Grid1.TextMatrix(k, f_p))  ' re("POSO")
        Case 3 ' +EXAGOGES
           mPIS = Val(Grid1.TextMatrix(k, f_p)) ' re("POSO")
        Case 4 ' +EISAGOGES
           mPIS = -Val(Grid1.TextMatrix(k, f_p))  ' re("POSO")
        End Select
 Else
        Select Case Val(F_POS_APOU)
        Case 1 ' +EISAGOGES
           mXre = mPoso
        Case 2 ' -EISAGOGES
           mXre = -mPoso
        Case 3 ' +EXAGOGES
           mPIS = mPoso
        Case 4 ' +EISAGOGES
           mPIS = -mPoso
        End Select
 End If
 

        ' ΕΝΗΜΕΡΏΝΕΙ ΤΟ ΚΈΡΔΟς

         mKERDOS = 0
          meid.Close
           meid.Open "SELECT * FROM EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) + "'", Gdb, adOpenDynamic, adLockOptimistic 'adOpenForwardOnly, adLockReadOnly


         If PEL.Recordset("EIDOS") = "e" And Val(Grid1.TextMatrix(k, F_T)) > 0 Then
            If IsNull(meid("XTI")) Then
               mKERDOS = Val(Grid1.TextMatrix(k, F_T))
            Else
               mKERDOS = Val(Grid1.TextMatrix(k, F_T)) - meid("XTI")
            End If
            
         End If
sql = ""
sql = "INSERT INTO EGGTIM (PROOD_AJ,HME,ATIM,MONA,EKPT,TIMM,"
sql = sql + "POSO,KAU_AJIA,MIK_AJIA, PROELEYSH,"

sql = sql + "FPA,PELKOD,EIDOS,FCURRENCY,APOT,"
sql = sql + "ONOMA,XRE,PIS,KODE,KERDOS,MIKTA"
sql = sql + ") values ("
sql = sql + Str(k) + "," ' prood_aj
sql = sql + "'" + Format(f_dat, "MM/DD/YYYY") + "',"  ' HME
sql = sql + "'" + Left(f_matim, f_psifia_atim + 2) + "'," 'ATIM
sql = sql + "'" + Grid1.TextMatrix(k, F_M) + "',"  'mona
sql = sql + Str(Val(Grid1.TextMatrix(k, f_e))) + ","  'ekpt
sql = sql + Str(Val(Grid1.TextMatrix(k, F_T))) + ","  'timm
sql = sql + Str(mPoso) + "," ' POSO
sql = sql + Str(mKAU_AJIA) + "," ' KAU_aj
sql = sql + Str(mMIK_AJIA) + "," ' MIK_ajIA
sql = sql + "'" + Left(mProeleysh, 10) + "'," ' PROELEYSH

sql = sql + Str(Val(Grid1.TextMatrix(k, f_f))) + ","  'FPA
sql = sql + "'" + PEL.Recordset("kod") + "'," 'PELKOD
sql = sql + "'" + PEL.Recordset("EIDOS") + "'," 'EIDOS
sql = sql + Str(gVal(isotimia.Text) * Val(Grid1.TextMatrix(k, F_T))) + "," 'FCURRENCY
sql = sql + Left(APOTH1.Text, 2) + ","  'APOT
sql = sql + "'" + mONOMA + "'," 'ONOMA
sql = sql + Str(mXre) + ","  'XRE
sql = sql + Str(mPIS) + ","  'PIS
sql = sql + "'" + Grid1.TextMatrix(k, f_k) + "',"  'KODE
sql = sql + Str(mKERDOS) + ", "  'kerdos
sql = sql + Str(Val(Grid1.TextMatrix(k, F_mikta))) + ") "      'mikta
N = 0
Gdb.Execute sql, N


If F_TRANS = 1 And N = 0 Then
    Gdb.RollbackTrans
    F_TRANS = -1 'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
    MsgBox "Λάθος στην σειρά " + Format(k, "###") + " στoν κωδικό " + Grid1.TextMatrix(k, f_k) + Chr(13) + " δεν κατεχωρήθη"
    Exit Sub
End If













        SYN_ME_FPA = SYN_ME_FPA + mMIK_AJIA
        FEGG(Val(Grid1.TextMatrix(k, f_f))) = FEGG(Val(Grid1.TextMatrix(k, f_f))) + mMIK_AJIA


        
'------------------------------------------------------------------------------------------------
 If F_STADIO = 5 Then ' ενδοδιακινηση  Left(f_matim, 1) = "σ" Then
30      re.AddNew
        re("hme") = f_dat
        re("atim") = f_matim
         
         
         If F_EXO_XROMATA = 0 Then
            re("poso") = Val(Grid1.TextMatrix(k, f_p))
         Else
            re("poso") = XROMATA(k, K2, k3)
            re("PROELEYSH") = Left(XROMATA(k, K2, 0), 3) + Left(XROMATA(k, 0, k3), 2) + "*" + Format(K2, "00") + "-" + Format(k3, "00")
         End If
        
       
        re("timm") = Val(Grid1.TextMatrix(k, F_T))
        re("kode") = Grid1.TextMatrix(k, f_k)
        re("onoma") = Left(Grid1.TextMatrix(k, f_o), 30)
        re("mona") = Right(Grid1.TextMatrix(k, f_o), 3)
        re("kau_ajia") = Val(Grid1.TextMatrix(k, f_a))
        re("fpa") = Val(Grid1.TextMatrix(k, f_f))
        re("ekpt") = Val(Grid1.TextMatrix(k, f_e))
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

Dim G, M, GA, MA, P
'Dim DB As New ADODB.Connection
'Dim e As New ADODB.Recordset
'Dim R As New ADODB.Recordset, K
Dim db As Database, R As New ADODB.Recordset, k, L As Integer
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
M = "M" + Format(Month(f_dat), "00")
GA = "GA" + Format(Month(f_dat), "00")
MA = "MA" + Format(Month(f_dat), "00")



Set fSCR = CreateObject("MSScriptControl.ScriptControl")
fSCR.language = "vbscript"
fSCR.AddObject "eid", EID
fSCR.AddObject "mactext", MACtEXT











'R.MoveFirst
'e.Index = "EIDko"
'  DB.Open gDir
For k = 1 To Grid1.ROWS - 1 'Do While Not R.EOF
  
  'e.Seek "=", R("KODE")


  
'    e.Close
If Len(Grid1.TextMatrix(k, f_k)) > 0 Then
    'Set gEid = db.OpenRecordset("SELECT *FROM EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) + "'") ' , DB, adOpenKeyset, adLockOptimistic
  EID.RecordSource = "SELECT *FROM EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) _
      + "'" ' , DB, adOpenKeyset, adLockOptimistic
  EID.Refresh
  
  If EID.Recordset.RecordCount = 1 Then
    P = Val(Grid1.TextMatrix(k, f_p)) '("POSO")
    NeaTim = Val(Grid1.TextMatrix(k, F_T)) * (100 - Val(Grid1.TextMatrix(k, _
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
         PAR11.Label11 = NeaTim
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
         EID.Recordset(M) = EID.Recordset(M) + P
    Case "4" ' -EXAGOGES
         EID.Recordset("POS") = EID.Recordset("POS") + P
         EID.Recordset(M) = EID.Recordset(M) - P
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

Dim rTemp As New ADODB.Recordset
Dim SQLt As String, TYP As Single
Dim N As Long







Dim X As String


Dim db As Database, e As Recordset, R As New ADODB.Recordset, k


On Error GoTo MHNYMA



Dim SQL22 As String


'ENHMERVSH ME PROHGOYMENO YPOLOIPO
'----------------------------------
SQLt = "SELECT KOD,"
SQLt = SQLt + "SUM(CASE WHEN XREOSI IS NULL  THEN 0 ELSE XREOSI  END )  -"
SQLt = SQLt + " SUM(CASE WHEN PISTOSI IS NULL  THEN 0 ELSE PISTOSI  END )   AS YP"
SQLt = SQLt + " FROM EGG  where HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND EIDOS='" + PEL.Recordset("EIDOS") + "' AND KOD='" + PEL.Recordset("kod") + "' GROUP BY KOD"
rTemp.Open SQLt, Gdb, adOpenDynamic, adLockOptimistic
If IsNull(rTemp("YP")) Then
   TYP = 0
Else
   TYP = rTemp("YP")
End If
SQL22 = "UPDATE TIM  SET EKPT5 =" + Format(PEL.Recordset("AYP"), "###########.00") + "+" + Format(TYP, "###########0.00")
SQL22 = SQL22 + " WHERE  HME='" + Format(DTPicker1.value, "MM/DD/YYYY") + "' AND ATIM='" + f_matim + "'"
Gdb.Execute SQL22

R.Open "SELECT TOP 1 * FROM EGG", Gdb, adOpenDynamic, adLockOptimistic
   f_suma = Round(Val(SynoloG.Caption), 2)


Dim MXREOSI, MPISTOSI
Dim mXre
Dim sql
Dim n1
   MXREOSI = 0: MPISTOSI = 0 ' ΑΠΟΦΥΓΗ ΝULL

      If f_pel = "1" Or f_pel = "4" Then ' XREVSH
         mXre = f_suma
         If f_pel = "1" Then MXREOSI = f_suma
         If f_pel = "4" Then MPISTOSI = -f_suma
      Else
         mXre = -f_suma
         If f_pel = "2" Then MXREOSI = -f_suma
         If f_pel = "3" Then MPISTOSI = f_suma
      End If
sql = "INSERT INTO EGG (EIDOS,KOD,HME,USERID,APA,ATIM,AIT,XRE,XREOSI,PISTOSI) VALUES ("
sql = sql + "'" + PEL.Recordset("eidos") + "'," 'EIDOS
sql = sql + "'" + PEL.Recordset("kod") + "'," 'KOD

sql = sql + "'" + Format(DTPicker1.value, "MM/DD/YYYY") + "',"  'HME
sql = sql + Str(gUserId) + ","  'USERID
sql = sql + mID$(f_matim, 2, 6) + "," 'APA
sql = sql + "'" + f_matim + "',"  'ATIM
sql = sql + "'" + Left(F_PARAS, 20) + "',"  'AIT
sql = sql + Str(mXre) + ","  'XRE
sql = sql + Str(MXREOSI) + ","  'XREOSI
sql = sql + Str(MPISTOSI) + ")"  'PISTOSI
n1 = 0
On Error Resume Next
Gdb.Execute sql, n1

   If F_TRANS = 1 And n1 = 0 Then
      Gdb.RollbackTrans
      MsgBox "Λάθος στην ενημέρωση λογαριασμού " + Chr(13) + " δεν κατεχωρήθη" + Chr(13) + Err.Description
      
      F_TRANS = -1 'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
      Exit Sub
   End If











On Error Resume Next



'R.Close


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
            

ROLL:

   If F_TRANS = 1 Then
      Gdb.RollbackTrans
      MsgBox "Λάθος στην ενημέρωση λογαριασμού " + Chr(13) + " δεν κατεχωρήθη"
      F_TRANS = -1 'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
      Exit Sub
   Else
      MsgBox "Προσοχή δεν πέρασε στην καρτέλα του λογαριασμού "
   End If
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
Dim db As Database, e As New ADODB.Recordset, R As New ADODB.Recordset, k
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
Dim TEL_YPOL As Single
'
On Error Resume Next

Dim MXREOSI, MPISTOSI
Dim mXre
Dim sql
Dim n1
Dim mXPI
MXREOSI = 0: MPISTOSI = 0 ' ΑΠΟΦΥΓΗ ΝULL
      
      
       If f_pel = "1" Then  '  TIMOLOGIO , LIANIKH
             MPISTOSI = Abs(f_suma)
             mXPI = "Π"
       ElseIf f_pel = "2" Then  ' PISTOTIKO TIMOLOGIO
             MPISTOSI = -Abs(f_suma)
             mXPI = "Π"
       ElseIf f_pel = "3" Then  ' tim.agoras
             MXREOSI = Abs(f_suma)
             mXPI = "Χ"
       ElseIf f_pel = "3" Then  ' PISTOTIKO TIMOLOGIO apo promitheyti
             MXREOSI = -Abs(f_suma)
             mXPI = "Χ"
       End If





   If PEL.Recordset("eidos") = "e" Then
      If f_pel = "1" Or f_pel = "4" Then ' XREVSH
         mXre = Round(m_suma, 2)
      Else  ' 2 , 3
         mXre = Round(-m_suma, 2)
      End If
   Else ' ΠΡΟΜΗΘΕΥΤΕΣ ' ΠΡΟΜΗΘΕΥΤΕΣ   F_PEL=3 ΧΡΕΩΣΗ  4= -ΧΡΕΩΣΗ     1=ΠΙΣΤΩΣΗ   2=-ΠΙΣΤΩΣΗ
      If f_pel = "2" Then
         mXre = Round(-m_suma, 2)
      Else ' 2 , 3
         mXre = Round(m_suma, 2)
      End If
   End If
       
       
       
       
       
       
       
       
       
       
sql = "INSERT INTO EGG (EIDOS,KOD,HME,USERID,APA,ATIM,AIT,XRE,XREOSI,PISTOSI) VALUES ("
sql = sql + "'" + PEL.Recordset("eidos") + "'," 'EIDOS
sql = sql + "'" + PEL.Recordset("kod") + "'," 'KOD
sql = sql + "'" + Format(DTPicker1.value, "MM/DD/YYYY") + "',"  'HME
sql = sql + Str(gUserId) + ","  'USERID
sql = sql + mID$(f_matim, 2, 6) + "," 'APA
sql = sql + "'" + f_matim + "',"  'ATIM
sql = sql + "'" + Left(Combo4.Text, 20) + "',"   'AIT
sql = sql + Str(mXre) + ","  'XRE
sql = sql + Str(MXREOSI) + ","  'XREOSI
sql = sql + Str(MPISTOSI) + ")"  'PISTOSI
n1 = 0
On Error Resume Next
Gdb.Execute sql, n1
If F_TRANS = 1 And n1 = 0 Then
      Gdb.RollbackTrans
      MsgBox "Λάθος στην ενημέρωση λογαριασμού " + Chr(13) + " δεν κατεχωρήθη" + Chr(13) + Err.Description
      F_TRANS = -1 'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
      Exit Sub
End If






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




MXREOSI = 0: MPISTOSI = 0 ' ΑΠΟΦΥΓΗ ΝULL
If PEL.Recordset("eidos") = "e" Then ' εισπραξη
            If f_pel = "1" Then MXREOSI = f_suma: mXre = f_suma '
            If f_pel = "2" Then MXREOSI = -f_suma: mXre = -f_suma '
            If f_pel = "3" Then MPISTOSI = f_suma: mXre = f_suma '
            If f_pel = "4" Then MXREOSI = -f_suma: mXre = -f_suma '
Else
            MPISTOSI = f_suma  'πληρωμή
End If
sql = "INSERT INTO EGG (EIDOS,KOD,HME,USERID,APA,ATIM,AIT,XRE,XREOSI,PISTOSI) VALUES ("
sql = sql + "'e'," 'EIDOS
sql = sql + "'" + m_tameio + "'," 'KOD
sql = sql + "'" + Format(DTPicker1.value, "MM/DD/YYYY") + "',"  'HME
sql = sql + Str(gUserId) + ","  'USERID
sql = sql + mID$(f_matim, 2, 6) + "," 'APA
sql = sql + "'" + f_matim + "',"  'ATIM
sql = sql + "'" + Left("Μετρητά", 20) + "',"   'AIT
sql = sql + Str(mXre) + ","  'XRE
sql = sql + Str(MXREOSI) + ","  'XREOSI
sql = sql + Str(MPISTOSI) + ")"  'PISTOSI
n1 = 0
On Error Resume Next
Gdb.Execute sql, n1


















'------------------ ΤΑΜΕΙΟ ΚΑΤΑΣΤΗΜΑΤΟΣ ------------------------------------------------














R.Close


Exit Sub

MHNYMA:
HandleError "Par1:enhm_met"
Resume Next
            

ROLL:
   If F_TRANS = 1 Then
      Gdb.RollbackTrans
      MsgBox "Λάθος στην ενημέρωση λογαριασμού " + Chr(13) + " δεν κατεχωρήθη"
      F_TRANS = -1 'FLAG GIA NA MHN SYNEXISEI ENHMEROSI PELATI
      Exit Sub
   Else
      MsgBox "Προσοχή δεν πέρασαν τα μετρητά  στην καρτέλα του λογαριασμού "
   End If
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
  Dim k
  Dim a$, b, X
  Dim d
  Dim rSQL As New ADODB.Recordset
  Dim rSQL2 As New ADODB.Recordset
  
  
  
 If Len(Dir("c:\TIMOLOGIA", vbDirectory)) > 2 Then
   ' LOISIR
   loisir_import
    Exit Sub
 End If
  
 
 
 If Len(Dir("c:\WESTNET", vbDirectory)) > 2 Then
   ' WESTNET
   WESTNET_import
    Exit Sub
 End If
 
 
  
 
 
  If Len(Dir("c:\SOKRATIS", vbDirectory)) > 2 Then
   ' SOKRATIS
      SOKRATIS_import
    Exit Sub
 End If
 
 
 
 
 
 
 If Len(Dir("C:\SYMBOL", vbNormal)) > 2 Then
   ' GAZOROS
    SYMBOL_import
    Exit Sub
 End If
  
  
  
  
  
  
  
  ' LAMBANEI ENA MONO PARASTAR\TIKO
If F_falcon_Polla_tim = 2 Then    ' mpoygoydis
  Dim ANS As Integer
  ANS = MsgBox("θΈΛΩ ΚΑΙ ΠΟΣΟΤΗΤΕΣ;", vbYesNo)
  If ANS = vbNo Then
  
    ' μπουγουδης
   Set db = OpenDatabase("C:\LAGEURO\MHXANAKI", False, False, "dBase III;")
   Set R = db.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM=" + Left(Text1.Text, 5) + ";")
   If R.RecordCount = 0 Then
      MsgBox "δεν υπάρχουν εγγραφές"
      Exit Sub
   End If


   R.MoveFirst
   k = 0
   Do While Not R.EOF
      k = k + 1
      If k > f_SeiresTimologioy Then Exit Do
      If IsNull(R("BARCODE")) Then
         Grid1.TextMatrix(k, f_k) = " "
      Else
         Grid1.TextMatrix(k, f_k) = R("BARCODE")
      End If
      R.MoveNext
   Loop
   R.Close
   db.Close
  
   Exit Sub
  
  End If

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
   k = 0
   Dim M_CODE
   Do While Not R.EOF
      k = k + 1
      If k > f_SeiresTimologioy Then Exit Do
      If IsNull(R("BARCODE")) Or R("pos") = 0 Then
         Grid1.TextMatrix(k, f_k) = " "
      Else
      
      
         rSQL.Open "SELECT * FROM BARCODES WHERE left(ERG,13)='" + (Left(R("BARCODE"), 13)) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
        If rSQL.EOF And Left(R("BARCODE"), 1) = "0" Then
           rSQL.Close
           
           rSQL.Open "SELECT * FROM BARCODES WHERE ERG LIKE '" + Trim(mID(R("BARCODE"), 2, 12)) + "%'", Gdb, adOpenForwardOnly, adLockReadOnly
        End If
        
        If rSQL.EOF Then
            M_CODE = " "
            Grid1.TextMatrix(k, f_k) = M_CODE
               '       OK = False
            '       List2.AddItem rDBF("BARCODE") + ": ΔΕΝ ΥΠΑΡΧΕΙ ΣΤΟ ΑΡΧΕΙΟ ΕΙΔΩΝ"
        Else
            M_CODE = rSQL("KOD")
            Grid1.TextMatrix(k, f_k) = M_CODE
          '   Grid1.TextMatrix(K, F_T) = rSQL("LTI")
            Grid1.TextMatrix(k, f_p) = R("POS")
            rSQL.Close
            rSQL.Open "SELECT * FROM EID WHERE KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
            Grid1.TextMatrix(k, F_T) = Round(rSQL("LTI"), 2)
            Grid1.TextMatrix(k, f_o) = rSQL("ONO")
            Grid1.TextMatrix(k, f_f) = rSQL("FPA")
            
            If f_TIMOKAT > 0 And Not IsNull(PEL.Recordset("pek")) Then
               rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + Str(PEL.Recordset("pek")) + " AND KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
               If rSQL2.EOF Then
                  Grid1.TextMatrix(k, f_e) = 0
               Else
                  Grid1.TextMatrix(k, f_e) = rSQL2(0)
               End If
               rSQL2.Close
               
               'ΔΟΥΛΕΥΕΙ ΜΕ ΚΑΡΦΩΤΗ ΤΙΜΗ
             ElseIf f_TIMOKAT < 0 And Not IsNull(PEL.Recordset("pek")) Then
               rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + Str(PEL.Recordset("pek")) + " AND KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
               If rSQL2.EOF Then
                  ' Grid1.TextMatrix(k, F_T) = 0
               Else
                  Grid1.TextMatrix(k, F_T) = rSQL2(0)
               End If
               rSQL2.Close
               
               
               
               
               
               
               
               
            Else
              Grid1.TextMatrix(k, f_e) = PEL.Recordset("pek")
              ' Grid1.Text =
            End If
            
            Grid1.row = k
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
   k = 0
   'Dim M_CODE
   Do While Not R.EOF
      k = k + 1
      If k > f_SeiresTimologioy Then Exit Do
      If IsNull(R("BARCODE")) Or R("pos") = 0 Then
         Grid1.TextMatrix(k, f_k) = " "
      Else
      
      
        rSQL.Open "SELECT * FROM BARCODES WHERE LEFT(ERG,13)='" + Left(R("BARCODE"), 13) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
        
       If rSQL.EOF And Left(R("BARCODE"), 2) = "00" Then
         rSQL.Close
           
           rSQL.Open "SELECT * FROM BARCODES WHERE ERG LIKE '" + mID(R("BARCODE"), 2, 12) + "%'", Gdb, adOpenForwardOnly, adLockReadOnly
        End If
        
        
        
        
        If rSQL.EOF Then
            M_CODE = " "
            Grid1.TextMatrix(k, f_k) = M_CODE
        Else
            M_CODE = rSQL("KOD")
            Grid1.TextMatrix(k, f_k) = Trim(M_CODE)
            Grid1.TextMatrix(k, f_p) = R("POS")
            rSQL.Close
            
            rSQL.Open "SELECT * FROM EID WHERE KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
            Grid1.TextMatrix(k, F_T) = Round(rSQL("LTI"), 2)
            Grid1.TextMatrix(k, f_o) = Trim(rSQL("ONO"))
            Grid1.TextMatrix(k, f_f) = rSQL("FPA")
            
            If f_TIMOKAT > 0 And Not IsNull(PEL.Recordset("pek")) Then
               rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + Str(PEL.Recordset("pek")) + " AND KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
               If rSQL2.EOF Then
                  Grid1.TextMatrix(k, f_e) = 0
                  
                     'ΔΟΥΛΕΥΕΙ ΜΕ ΚΑΡΦΩΤΗ ΤΙΜΗ
             ElseIf f_TIMOKAT < 0 And Not IsNull(PEL.Recordset("pek")) Then
               rSQL2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + Str(PEL.Recordset("pek")) + " AND KOD='" + M_CODE + "'", Gdb, adOpenForwardOnly, adLockReadOnly
               If rSQL2.EOF Then
                  ' Grid1.TextMatrix(k, F_T) = 0
               Else
                  Grid1.TextMatrix(k, F_T) = rSQL2(0)
               End If
               rSQL2.Close
                 
                  
                  
               Else
                  Grid1.TextMatrix(k, f_e) = rSQL2(0)
               End If
               rSQL2.Close
            Else
              Grid1.TextMatrix(k, f_e) = PEL.Recordset("pek")
              ' Grid1.Text = PEL.Recordset("pek")
            End If
            
            Grid1.row = k
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
'Sub LOISIR2_import()
'
'  Dim db As DAO.Database
'  Dim R As DAO.Recordset
'  Dim k
'  Dim A$, b, X
'  Dim d
'  Dim rSQL As New ADODB.Recordset
'
'
'
'
' If Len(Dir("c:\TIMOLOGIA", vbDirectory)) > 2 Then
'   ' LOISIR
'
'
'
'Open "c:\timologia\" + Text1.Text + ".asc" For Binary Access Read As #1
'
''Do While Not EOF(1)
'' Line Input #1, A$
'' Exit Do
''Loop
''Close #1
''Exit Sub
'
'X = 1
'
'
'Dim kod, per, MON, pos, TIMH
'Dim f As String, RECS As Integer
'
'k = 0
'
'Do While Not EOF(1) 'True
'  On Error Resume Next
'    Line Input #1, A$
'   If EOF(1) Then Exit Do
'
'   f = Chr$(10) + Chr$(27) + Chr$(15)
'
''   B = InStr(x, A$, f)
''   If B = 0 Then
' '    GoTo 555
''     Exit Do
'  ' End If
'
'
'
'
'
'X = 1
'
'     f = Chr$(27) + Chr$(15) 'Chr$(10) + Chr$(27) + Chr$(15)
'
'   b = InStr(X, A$, f)
'
'   If b = 0 Then
'     GoTo 555
'   End If
'
'
'
'
'
'
'
'
'
'   d = mID(A$, b + 2, 180)
'   kod = Left(d, 12)
'
'
'
'
'
'
'
'   per = mID$(d, 30, 38)
'   pos = mID$(d, 76, 10)
'
'   MON = mID$(d, 90, 10)
'   TIMH = Trim(mID$(d, 104, 10))
'   TIMH = Replace(TIMH, ",", ".")
'
'   'List1.AddItem d
'
'
'
'rSQL.Open "SELECT * FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
'   If rSQL.EOF Then
' '      On Error Resume Next
'      Gdb.Execute "INSERT INTO EID (KOD,ONO,FPA,XTI,MON) VALUES " _
'      & "('" + Trim(kod) + "','" + Replace(Left(Trim(per), 35), "'", "~") + "',2," + TIMH + ",'" + Trim(MON) + "')", RECS
'      If RECS = 0 Then
'         MsgBox "Δεν υπάρχει το είδος " + kod
'      Else
'         MsgBox "Ανοίχθηκε το είδος " + kod
'      End If
'   End If
'
'rSQL.Close
'
'   X = b + 3
'
'
'  k = k + 1
'  If k > f_SeiresTimologioy Then Exit Do
'     Grid1.TextMatrix(k, f_k) = kod
'     Grid1.TextMatrix(k, f_o) = per
'     Grid1.TextMatrix(k, f_p) = pos
'     Grid1.TextMatrix(k, F_T) = TIMH
'     Grid1.TextMatrix(k, f_f) = 2
'     Grid1.row = k
'     FindSynolo
'
'555
'Loop
'Close #1
'
'
' Else
' ' μπουγουδης
'Set db = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
'Set R = db.OpenRecordset("SELECT * FROM TIMAGOR WHERE ATIM=" + Left(Text1.Text, 5))
'If R.RecordCount = 0 Then
'   MsgBox "δεν υπάρχουν εγγραφές"
'   Exit Sub
'End If
'
'
'R.MoveFirst
'k = 0
'Do While Not R.EOF
'  k = k + 1
'  If k > f_SeiresTimologioy Then Exit Do
'  If IsNull(R("BARCODE")) Then
'     Grid1.TextMatrix(k, f_k) = " "
'  Else
'     Grid1.TextMatrix(k, f_k) = R("BARCODE")
'  End If
'  R.MoveNext
'Loop
'R.Close
'db.Close
'
'End If
'
'
'
'End Sub

Sub loisir_import()

  Dim db As DAO.Database
  Dim R As DAO.Recordset
  Dim k
  Dim a$, b, X
  Dim d
  Dim rSQL As New ADODB.Recordset
  
  Dim ANS
  ANS = MsgBox("ΜΕ EXCEL", vbYesNo)
  If ANS = vbYes Then
      LOISIR2_import
      Exit Sub
  End If
  
  
  
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
Dim F As String, RECS As Integer

k = 0

Do While Not EOF(1) 'True
  On Error Resume Next
    Line Input #1, a$
   If EOF(1) Then Exit Do
    
   F = Chr$(10) + Chr$(27) + Chr$(15)
   
'   B = InStr(x, A$, f)
'   If B = 0 Then
 '    GoTo 555
'     Exit Do
  ' End If
   
   
      
      
      
X = 1
   
    ' f = Chr$(27) + Chr$(15) 'Chr$(10) + Chr$(27) + Chr$(15)
   F = Chr$(15)  'Chr$(10) + Chr$(27) + Chr$(15)
   
   b = InStr(X, a$, F)
 
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


  k = k + 1
  If k > f_SeiresTimologioy Then Exit Do
     Grid1.TextMatrix(k, f_k) = kod
     Grid1.TextMatrix(k, f_o) = per
     Grid1.TextMatrix(k, f_p) = pos
     Grid1.TextMatrix(k, F_T) = TIMH
     Grid1.TextMatrix(k, f_f) = 2
     Grid1.row = k
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
k = 0
Do While Not R.EOF
  k = k + 1
  If k > f_SeiresTimologioy Then Exit Do
  If IsNull(R("BARCODE")) Then
     Grid1.TextMatrix(k, f_k) = " "
  Else
     Grid1.TextMatrix(k, f_k) = R("BARCODE")
  End If
  R.MoveNext
Loop
R.Close
db.Close

End If







End Sub

Sub WESTNET_import()

  Dim db As DAO.Database
  Dim R As DAO.Recordset
    Dim r2 As New ADODB.Recordset
  Dim k
  Dim a$, b, X
  Dim d
  Dim rSQL As New ADODB.Recordset
  
  Dim ANS
  
  Dim MONO As String
  
  
  If Len(Dir("c:\WESTNET\*.*", vbNormal)) < 3 Then
     MsgBox "δεν υπάρχει αρχείο για τιμολόγηση"
  End If

   ' R2.Open "SELECT COUNT(*) FROM EID WHERE KOD='00106983'", Gdb, adOpenDynamic, adLockOptimistic


   
  POLAGOR.Text = POLAGOR.List(0)
  POLAGOR_LostFocus
  
  PARAS.Text = PARAS.List(8)
  
  Text2(0).Text = "0064"
  fbuff = "r"
 ' Text2_KeyUp 2, 13, 1
  ' GridPelaton.SetFocus 'διαλεξε πελάτη
 ' SendKeys "{ENTER}{ENTER}"

  
  
Open "c:\WESTNET\" + Dir("c:\WESTNET\*.*", vbNormal) For Input As #1
X = 1


Dim kod, per, MON, pos, TIMH
Dim F As String, RECS As Integer

k = 0

Do While Not EOF(1) 'True
  On Error Resume Next
    Line Input #1, a$
   If EOF(1) Then
       Exit Do
   End If
   
    Line Input #1, a$
    Line Input #1, a$
    Line Input #1, a$
    Line Input #1, a$
    Line Input #1, a$
    Line Input #1, a$
    Line Input #1, a$
    Line Input #1, a$
    Line Input #1, a$
    Line Input #1, a$  'ΔΙΑΒΑΖΕΙ ΔΙΕΥΘΥΝΣΗ
    Text1.Text = mID$(a$, 55, 6)
    
    DTPicker1.value = Replace(mID$(a$, 87, 26), ".", "/")
    
    
    Line Input #1, a$
    parat.Text = mID$(a$, 62, 14)
    
    Line Input #1, a$
    Line Input #1, a$
    Line Input #1, a$
    Line Input #1, a$
    Line Input #1, a$
    Line Input #1, a$
    
    Line Input #1, a$
    Line Input #1, a$
    Line Input #1, a$
    Line Input #1, a$
    Line Input #1, a$
    
    Line Input #1, a$
    Line Input #1, a$
    Line Input #1, a$
    
 Dim N As Integer
 
 
 For k = 1 To 20
 
    Line Input #1, a$ '1o EIDOS
   
'    If Len(Trim(Left(kod, 5))) = 0 Then
 '      Exit For
  '  End If
    
     If Len(Trim(Left(a$, 5))) = 0 Then 'EISFORA ANAKYKLOSHS 'H KENO
        If Len(Trim(Left(a$, 18))) = 0 Then 'KENO
           'Line Input #1, A$ '1o EIDOS
           Exit For
        Else
           Line Input #1, a$ '1o EIDOS
        End If
    End If
    
    kod = Trim(mID$(a$, 4, 8))
    
    per = mID$(a$, 15, 38)
    pos = Val(Replace(mID(a$, 67, 5), ",", "."))
    
    pos = Val(Replace(mID(a$, 63, 6), ",", "."))
    TIMH = Val(Replace(mID(a$, 75, 9), ",", "."))
    
    
    rSQL.Close
    r2.Close
    
    
    rSQL.Open "SELECT * FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenDynamic, adLockOptimistic
 
    r2.Open "SELECT COUNT(*) FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenDynamic, adLockOptimistic
    
    
    If r2(0) = 0 Then
       Gdb.Execute "INSERT INTO EID (KOD,ONO,FPA,XTI,MON) VALUES " _
       & "('" + Trim(kod) + "','" + Replace(Left(Trim(per), 35), "'", "~") + "',2," + Str(TIMH) + ",'" + Trim(MON) + "')", RECS
       If RECS = 0 Then
         MsgBox "Δεν υπάρχει το είδος " + kod
       Else
         MsgBox "Ανοίχθηκε το είδος " + kod
       End If
    End If
     
    If k > f_SeiresTimologioy Then
        Exit For
    End If
    
     Grid1.TextMatrix(k, f_k) = kod
     Grid1.TextMatrix(k, f_o) = per
     Grid1.TextMatrix(k, f_p) = pos
     Grid1.TextMatrix(k, F_T) = TIMH
     Grid1.TextMatrix(k, f_f) = 2
     Grid1.row = k
     FindSynolo
Next



rSQL.Close
    
    
    
    
    
    
    
    
    
     Line Input #1, a$
     Line Input #1, a$
     Line Input #1, a$
             
             
     Line Input #1, a$
     Line Input #1, a$
     Line Input #1, a$
             
             
             
  Line Input #1, a$
     Line Input #1, a$
     Line Input #1, a$
             
   Line Input #1, a$
     Line Input #1, a$
     Line Input #1, a$
 Line Input #1, a$
     Line Input #1, a$
     Line Input #1, a$
                                    
 Line Input #1, a$
     Line Input #1, a$
     Line Input #1, a$
             
   'Line Input #1, A$
     
For k = 1 To 9

   Text5.Text = "*" + LTrim((Replace(mID(a$, 60, 20), ",", ".")))
  If Val(mID(Text5.Text, 2, 10)) > 0 Then
      Exit For
  End If
  
   Line Input #1, a$
Next
 
 Exit Do
             
             
             
             
             Loop
             
             Close #1
   
Kill "c:\WESTNET\" + Dir("c:\WESTNET\*.*", vbNormal)






End Sub






Sub LOISIR2_import()
' ME EXCEL
Dim xl As New excel.Application
Dim xlsheet As excel.Worksheet
Dim xlsheet3 As excel.Worksheet
Dim xlwbook As excel.workbook


Set xlwbook = xl.Workbooks.Open("c:\TIMOLOGIA\" + Text1.Text + ".XLS")
Set xlsheet = xlwbook.Sheets.Item(1)







  Dim k
  Dim a$, b, X
  Dim d
  Dim FPA As Integer
  
  Dim rSQL As New ADODB.Recordset

Dim kod, per, MON, pos, TIMH, TIMOL
Dim F As String, RECS As Integer

k = 1


Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF
     If IsNull(xlsheet.cells(k, 1)) Then
         Exit Do
     End If
     
     If IsEmpty(xlsheet.cells(k, 1)) Then
         Exit Do
     End If
     On Error Resume Next
     kod = xlsheet.cells(k, 1)
  
   pos = xlsheet.cells(k, 3)
   TIMH = xlsheet.cells(k, 4)
   
   

   rSQL.Open "SELECT * FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
   If rSQL.EOF Then
        MsgBox "Δεν υπάρχει το είδος " + kod
   Else
        kod = rSQL(0)
        'rSQL.Close
        
        'rSQL.Open "SELECT ONO,XTI,FPA FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
        per = rSQL("ONO")
       ' TIMH = rSQL("XTI")
        FPA = rSQL("FPA")
        'TIMH = Round(rSQL("XTI"), 2) '  / ((100 + g_Fpa(FPA)) / 100)
        
        
        
   End If
   
   rSQL.Close
   

     If k > f_SeiresTimologioy Then Exit Do
     Grid1.TextMatrix(k, f_k) = kod
     Grid1.TextMatrix(k, f_o) = per
     Grid1.TextMatrix(k, f_p) = pos
     Grid1.TextMatrix(k, F_T) = TIMH
     Grid1.TextMatrix(k, f_f) = FPA
     Grid1.row = k
     FindSynolo
     
    k = k + 1
     
'If EOF(1) Then Exit Do
     
     

Loop
'Close #1

  

Set xlsheet = Nothing
Set xlwbook = Nothing






End Sub


Sub SOKRATIS_import()
' ME EXCEL
Dim xl As New excel.Application
Dim xlsheet As excel.Worksheet
Dim xlsheet3 As excel.Worksheet
Dim xlwbook As excel.workbook


Set xlwbook = xl.Workbooks.Open("c:\SOKRATIS\SOK.XLS")
Set xlsheet = xlwbook.Sheets.Item(1)


 Dim k
 Dim a$, b, X
 Dim d
 Dim FPA As Integer
  
 Dim rSQL As New ADODB.Recordset

Dim kod, per, MON, pos, TIMH, TIMOL
Dim F As String, RECS As Integer

k = 1
Dim CH1, CH2, CH3
Dim xti, LTI
Dim S As String
Dim nn As Long
Dim LL2 As Long


Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF
     If IsNull(xlsheet.cells(k, 1)) Then
         Exit Do
     End If
     
     If IsEmpty(xlsheet.cells(k, 1)) Then
         Exit Do
     End If
     On Error Resume Next
     kod = LTrim(Str(xlsheet.cells(k + 1, 1))) '
     CH1 = xlsheet.cells(k + 1, 2) ' PROSFORA
     CH2 = xlsheet.cells(k + 1, 3) ' XEIMONIATIKO
     CH3 = xlsheet.cells(k + 1, 6) 'PROM
     xti = xlsheet.cells(k + 1, 4) ' , ",", ".")
     LTI = xlsheet.cells(k + 1, 5) ', ",", ".")

   rSQL.Open "SELECT * FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
   If rSQL.EOF Then
        S = "INSERT INTO EID (KOD,CH1,CH2,CH3,XTI,LTI5,KATHGORIA,FPA) VALUES ("
        
        
        On Error GoTo 0
        
        
        S = S + "'" + kod + "','" + Str(CH1) + "','" + Str(CH2) + "','" + Str(CH3) + "'," + Str(xti) + "," + Str(LTI) + ",1,2)"
        
        Gdb.Execute S, nn
        
        
        
   Else

        S = "UPDATE  EID  SET XTI=" + Str(xti) + ",LTI5=" + Str(LTI) + "  WHERE KOD='" + Trim(kod) + "'"
        
        Gdb.Execute S, nn
        
   End If
   
   rSQL.Close
   

   If k > f_SeiresTimologioy Then Exit Do
     XROMATA(k, 0, 0) = -0.01     ' ΣΗΜΑΔΙ ΟΤΙ ΧΡΗΣΙΜΟΠΟΙΗΣΕ ΤΟ F6
     pos = 0
     gm_str(k) = ""
     For LL2 = 7 To 19
       XROMATA(k, LL2 - 6, 1) = xlsheet.cells(k + 1, LL2) '
       gm_str(k) = gm_str(k) + Str(XROMATA(k, LL2 - 6, 1)) + ";"
      ' pos = pos + xlsheet.cells(k + 1, LL2)
     Next
     
     Grid1.TextMatrix(k, f_k) = kod
     Grid1.TextMatrix(k, f_o) = ""
     Grid1.TextMatrix(k, f_p) = pos
     Grid1.TextMatrix(k, F_T) = xti
     Grid1.TextMatrix(k, f_f) = FPA
   
     

'For R = 1 To Grid1.ROWS - 1
'  For C = 1 To Grid1.ColS - 1
   ' Grid1.TextMatrix(R, C) = Replace(Grid1.TextMatrix(R, C), ";", " ")
  '  gm_str(Val(Label1.Caption)) = gm_str(Val(Label1.Caption)) + Grid1.TextMatrix(R, C) + ";"
  'Next
'Next


     
     
     
     
     
     
     
     
     
     
     
     
     
     Grid1.row = k
     FindSynolo
     
    k = k + 1
     
'If EOF(1) Then Exit Do
     
     

Loop
'Close #1

  

Set xlsheet = Nothing
Set xlwbook = Nothing
xl.Quit






End Sub






















Sub SYMBOL_import()

  Dim db As DAO.Database
  Dim R As DAO.Recordset
  Dim k
  Dim a$, b, X
  Dim d
  Dim FPA As Integer
  
  Dim rSQL As New ADODB.Recordset
Open "C:\APOG.TXT" For Input As #1

X = 1


Dim kod, per, MON, pos, TIMH, TIMOL
Dim F As String, RECS As Integer

k = 0

Do While Not EOF(1) 'True
  On Error Resume Next
    Line Input #1, a$

   X = 1
   kod = mID(a$, 1, 13)
   
   pos = Val(mID$(a$, 15, 5))
   
   TIMOL = Val(mID$(a$, 20, 6))
If TIMOL = Val(Text1.Text) And Val(kod) > 0 Then
   rSQL.Open "SELECT KOD FROM BARCODES WHERE ERG='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
   If rSQL.EOF Then
        MsgBox "Δεν υπάρχει το είδος " + kod
   Else
        kod = rSQL(0)
        rSQL.Close
        
        rSQL.Open "SELECT ONO,XTI,FPA FROM EID WHERE KOD='" + Trim(kod) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
        per = rSQL("ONO")
        TIMH = rSQL("XTI")
        FPA = rSQL("FPA")
        TIMH = Round(rSQL("XTI"), 2) '  / ((100 + g_Fpa(FPA)) / 100)
        
        
        
   End If
   
   rSQL.Close
   
    k = k + 1
     If k > f_SeiresTimologioy Then Exit Do
     Grid1.TextMatrix(k, f_k) = kod
     Grid1.TextMatrix(k, f_o) = per
     Grid1.TextMatrix(k, f_p) = pos
     Grid1.TextMatrix(k, F_T) = TIMH
     Grid1.TextMatrix(k, f_f) = FPA
     Grid1.row = k
     FindSynolo
     
End If ' TIMOL=TEXT1
     
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
     fpelat4.DATA1.ConnectionString = gConnect
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


Dim ko As String, ajia As String, N As Integer, TYPOS As Integer




ko = InputBox("δωσε το barcode του είδους ", , "")
ajia = InputBox("δωσε αξία ", , "")
TYPOS = Val(InputBox("Αποδ.ΠΥ=6 Δ.Α=9 ΤΠΥ=16 ", , ""))

If TYPOS = 6 Or TYPOS = 9 Or TYPOS = 16 Then
   ' OK
Else
    MsgBox "Λάθος τύπος παραστατικού"
    Exit Sub
End If


POLAGOR.Text = POLAGOR.List(1)

POLAGOR_LostFocus





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
         Gdb.Execute sql, N
    
    
    
    
    End If
    
    
End If
 
 
'μπαινω απ'ευθειας στο παραστατικό
  PARAS.Text = PARAS.List(TYPOS - 1)
  
  Dim ll As Integer
  Dim gg As String
  
  gg = IIf(TYPOS = 6, "y", IIf(TYPOS = 9, "A", "Y"))
  
  
  If gg = "y" And Val(mID(ko, 2, 6)) >= 3398 Then
     gg = "w"
     MsgBox "υποκαταστηματος"
  End If
  
  
  
  For ll = 0 To PARAS.ListCount - 1
      If Right(PARAS.List(ll), 1) = gg Then
          PARAS.Text = PARAS.List(ll)
      End If
  Next
  
  
  
  
  PARAS_LostFocus
  SendKeys "{ENTER}{ENTER}"
  SendKeys "{ENTER}{ENTER}"
  
 Dim k As Integer
 
 
 
'ΔΗΜΙΟΥΡΓΩ ΤΙΣ ΣΕΙΡΕΣ ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ
If TYPOS = 9 Then 'δελτιο αποστολής
  k = 1
  Grid1.TextMatrix(k, f_k) = ko
  Grid1.TextMatrix(k, f_p) = 1
  Grid1.TextMatrix(k, f_o) = R("ONO")
  Grid1.TextMatrix(k, F_T) = 0
  Grid1.TextMatrix(k, f_f) = 2
  Grid1.TextMatrix(k, f_a) = 0
  Grid1.row = k: Grid1.Col = 0:  FindSynolo
  SKOPOS.Text = "ΑΠΟ ΕΠΙΣΚΕΥΗ"
  Gdb.Execute "UPDATE EID SET HM3=GETDATE(),NUM1=" + Text1.Text + " WHERE KOD='" + ko + "'"
Else
  k = 1
  Grid1.TextMatrix(k, f_k) = "223344"
  TIMText3 = "223344"
  Grid1.TextMatrix(k, f_p) = 1
  Grid1.TextMatrix(k, f_o) = "ΕΠΙΣΚΕΥΗ"
  Grid1.TextMatrix(k, F_T) = Round(Val(ajia) / (1 + g_Fpa(2) / 100), 2)
  Grid1.TextMatrix(k, f_f) = 2
  Grid1.TextMatrix(k, f_a) = Round(Val(ajia) / (1 + g_Fpa(2) / 100), 2)
  Grid1.row = k: Grid1.Col = 0:  FindSynolo
  k = 2
  Grid1.TextMatrix(k, f_k) = ko
  Grid1.TextMatrix(k, f_p) = 0
  Grid1.TextMatrix(k, f_o) = R("ONO")
  Grid1.TextMatrix(k, F_T) = 0
  Grid1.TextMatrix(k, f_f) = 2
  Grid1.TextMatrix(k, f_a) = 0
  Grid1.row = k: Grid1.Col = 0:  FindSynolo
  Gdb.Execute "UPDATE EID SET HM3=GETDATE(),NUM2=" + Text1.Text + " WHERE KOD='" + ko + "'"
End If






End Sub

Sub Command5_Click()
'       CrystalReport1.ReportFileName = "C:\Mercvb\reports\timol1.rpt"
'       CrystalReport1.Destination = crptToWindow
'       CrystalReport1.Action = 1

Dim DUM

Dim par As New ADODB.Recordset

f_dat = DTPicker1.value


f_matim = Right(PARAS.Text, 1) + Right("00000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text
20




par.Open "SELECT *FROM PARASTAT WHERE EIDOS='" + Left(f_matim, 1) + "';", Gdb, adOpenDynamic, adLockOptimistic
If IsNull(par!FORM_EKTYP) Then f_FORM_EKTY = 0 Else f_FORM_EKTY = Val(par!FORM_EKTYP)




F_PARAS = PARAS.Text

'εκτύπωση
If ektyp.value = 1 Then ' ---------------------------------------------


If f_FORM_EKTY >= 90 Then ' φορμα τύπου f90.txt  '====================================
100    DUM = printCrystal(f_matim, CDate(DTPicker1.value)) 'δημιουργει τα προσωρινά αρχεία
110   DUM = print_timol(f_matim, CDate(DTPicker1.value))
Else ' crystal reports timol1.rpt(σφραγίδα)   timol1an.rpt (αντίγραφο)
   
   If f_FORM_EKTY > 0 Then '      PEL.Recordset("eidos") = "e" Then  ' EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE


       'ΑΝ ΧΡΗΣΙΜΟΠΟΙΕΙ ΤΗΝ ΦΟΡΜΑ ΤΙΜΟΛ2 TOTE NA DINEI ALLES SEIRES APO TO F_FORM2_SEIRESTIMOL
       If f_FORM_EKTY = 14 Or f_FORM_EKTY = 12 Then
           f_SeiresTimologioy = f_Timol2_SeiresTimologioy
           f_SynSeiresTimologioy = f_Timol2_SeiresTimologioy
       End If




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
      
    Dim ANS As Integer
    ANS = MsgBox("Σε οθόνη το τιμολόγιο;", vbYesNo)
    If ANS = vbYes Then
        CrystalReport1.Destination = crptToWindow
    End If
    
      
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
 
 
 End If   ' If f_FORM_EKTY > 0 Then '  EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
   
End If  ' If ektyp.Value = 1 Then  '========================================



End If  '-------------------------------------------------------------


























End Sub

Private Sub Command6_Click()
Dim fff As New pelat2

   ' fff.ffbuff = f_buff + Text2(0).Text ' Text1(0).Text
   
   If POLAGOR.ListIndex = 0 Then
      fff.PELPROM.ListIndex = 0
       fff.ffbuff = "r" + Text2(0).Text ' Text1(0).Text

   Else
       fff.PELPROM.ListIndex = 1
       fff.ffbuff = "e" + Text2(0).Text ' Text1(0).Text
   End If
   fff.COMMAND5.Enabled = False
   gApoMenu = False
   pelat2.SHOW
    
End Sub

Sub diortosis()
 
  Dim FF As PAR2, mkod As String
  
Dim Atim

MIDENIZO_GRID
   
TDBGrid1.Col = 4
mkod = TDBGrid1.Text
   
   
   
TDBGrid1.Col = 1
Atim = TDBGrid1.Text



Dim k As Integer


'ΧΡΕΙΑΖΕΤΑΙ 2 ΦΟΡΕΣ ΓΙΑ ΝΑ ΚΑΝΕΙ ΔΙΟΡΘΒ=ΩΣΗ ΜΕ ΤΗΝ ΠΡΩΤΗ
'------------------------------------------------------------------------------
'For k = 0 To PARAS.ListCount - 1
'   PARAS.ListIndex = k
'   If Right(PARAS.Text, 1) = Left(atim, 1) Then
'      Exit For
'   End If
'Next
''  find_eid_parastat
'  Text1.Text = Val(mID(atim, 2, 6))
'  Text1.SetFocus
''MILSEC 3500
''DoEvents
''  Text2(0).Text = mkod
''      PEL.RecordSource = "select  LEFT(left(EPO+'-'+DIE+'-',56)+KOD,60) as xx,DIE,THL,KOD  from PEL where EIDOS='" + X + "' and KOD like '" + SameLetters(Text2(0).Text) + "%' order by KOD;"
'''End If
''
''
''
'GridPelaton.Columns(0).width = 6000
'GridPelaton.Columns(0).DataField = "XX"
'
'PEL.Refresh
'GridPelaton.Columns(0).width = 6000
'GridPelaton.Visible = True
'GridPelaton.SetFocus
'SendKeys "{ENTER}{ENTER} {ENTER} "




'------------------------------------------------------------------------------





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
'Dim k As Integer
For k = 0 To PARAS.ListCount - 1
   PARAS.ListIndex = k
   If Right(PARAS.Text, 1) = Left(Atim, 1) Then
      Exit For
   End If
Next

  
TDBGrid1.Col = 0
DTPicker1.value = CDate(TDBGrid1.Text)

If k = PARAS.ListCount Then Exit Sub
  ' βρισκω τις παραμετρους του παραστατικού
  find_eid_parastat

  Text1.Text = Val(mID(Atim, 2, 6))

  Text2(0).Text = mkod

Text2(0).Enabled = True
Text2(0).SetFocus

  
SendKeys "{ENTER}{ENTER} {ENTER} "


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

Private Sub Form_Resize()
   mForm_Resize PAR1, 13860, 10830, fh, fw, ft, fl
 
End Sub

Private Sub Grid1_DblClick()
Dim F33 As New apot21


  TIMText3.Visible = True
' kartella
  F33.Text1.Text = Grid1.TextMatrix(Grid1.row, 1)
  
  F33.Text2.Text = Grid1.TextMatrix(Grid1.row, 2)
  F33.Label1.Caption = Left(APOTH1.Text, 1)
  F33.Label2.Caption = "0"
  
  
  
  
  'Me.Hide
  
  F33.SHOW
 
APOT2.Hide





End Sub

Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 If Button = 2 Then
    

    ' APOT2.kod.Text = MSFlexGrid1.Text   ' "10"
    
    
    APOT2.kod.Text = Grid1.TextMatrix(Grid1.row, 1)
    
    APOT2.SHOW
    APOT2.kod.SetFocus
    SendKeys "{ENTER}"
End If







End Sub

Private Sub Grid1_Scroll()
   TIMText3.Visible = False
End Sub

Private Sub GridEidon_GotFocus()
  ' ARXIKO Grid1.width = 5775
  GridEidon.width = 10320

End Sub

Private Sub GridEidon_LostFocus()
   GridEidon.width = 5775
End Sub

Private Sub isotimia_LostFocus()
   
   Dim k As Integer
   Dim mISOT
   
   mISOT = gVal(isotimia.Text)
   
   If mISOT <> 1 Then
      
      isotimia.Text = mISOT
      
      For k = 1 To Grid1.ROWS - 1
         If Val(Grid1.TextMatrix(k, f_p)) <> 0 Then
            Grid1.TextMatrix(k, F_T) = Val(Grid1.TextMatrix(k, F_T)) / mISOT
            Grid1.row = k: FindSynolo
         End If
      Next
   End If


End Sub

Private Sub kerdos_Click()
    MsgBox get_kerdos()
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









Private Sub MSFlexGrid1_Click()
  If Left(f_pelMERC, 3) = "SOK" Then
     MSFlexGrid1.Col = 1
     TIMText3.Text = MSFlexGrid1.Text
     TIMText3.SetFocus
     
  End If
End Sub

'Private Sub par2show_Click()
'    PAR2SHOWING
'End Sub

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


Dim X As String, k As Integer
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
    SKOPOS.Text = SKOPOS.List(3)
    selectPEL.Caption = "Προμηθευτής"
    f_SeiresTimologioy = f_MaxSeiresTimologioy 'POSA EIDH MPORO NA PERASO Max TIMOLOGIA Agoras
     'f_SeiresTimologioy = 30 OTAN DIALEGO  POLISEIS
    
    Combo4.Text = Combo4.List(1) ' PISTVSH
  ElseIf POLAGOR.Text = POLAGOR.List(1) Then
    selectPEL.Caption = "Πελάτης"
    X = "1": fbuff = "e" 'pelates
    f_SeiresTimologioy = f_MinSeiresTimologioy 'POSA EIDH MPORO NA PERASO Max tim .polisis
    SKOPOS.Text = SKOPOS.List(0) ' ΠΩΛΗΣΗ
    
  Else
    X = "3": fbuff = "l" 'λογαριασμοί
    f_SeiresTimologioy = f_MaxSeiresTimologioy 'POSA EIDH MPORO NA PERASO Max TIMOLOGIA Agoras
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
  k = 0
  Do While Not R.EOF
     If Not IsNull(R("TITLOS")) And R("pol") = X Then
       If F_STOPDELTIA = 99 Then
          If R("eidos") = "t" Or R("eidos") = "l" Then
             PARAS.AddItem R("TITLOS") + Space(1) + R("eidos")
             k = k + 1
          End If
       Else
          PARAS.AddItem R("TITLOS") + Space(1) + R("eidos")
             k = k + 1
      End If
        mEidos_Par(k) = R("eidos")
     End If
     R.MoveNext
  Loop

PARAS.Enabled = True

On Error Resume Next
   If X = "2" Then 'agores
     PARAS.Text = PARAS.List(f_1choice_agores)
   ElseIf X = "1" Then 'polhseis
                  ' PARAS.Text = PARAS.List(0)
     PARAS.Text = PARAS.List(f_1choice_polhseis)
                    
   Else
     PARAS.Text = PARAS.List(1)
   End If
   
   
   
End If

DTPicker1.value = Now


'PARAS.Enabled = True

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

Dim k As Integer

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
       
       'ΑΝ ΧΡΗΣΙΜΟΠΟΙΕΙ ΤΗΝ ΦΟΡΜΑ ΤΙΜΟΛ2 TOTE NA DINEI ALLES SEIRES APO TO F_FORM2_SEIRESTIMOL
       If Val(CNull(R("FORM_EKTYP"))) = 14 Or Val(CNull(R("FORM_EKTYP"))) = 12 Then
           f_SeiresTimologioy = f_Timol2_SeiresTimologioy
           f_SynSeiresTimologioy = f_Timol2_SeiresTimologioy
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
 Text1.Text = find_parastat(k, X, 0)


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

Dim k As Integer

Dim X As String

 ' Dim db As Database
 Dim R As New ADODB.Recordset
 On Error GoTo MHNYMA
 

SEIRA.BackColor = vbWhite



For k = 0 To SEIRA.ListCount
   If SEIRA.Text = SEIRA.List(k) Then
      Exit For
   End If
Next


If k <= SEIRA.ListCount Then
  ' Text1.Text = find_parastat(k, Right(PARAS.Text, 1), 0)
    X = mEidos_Par(1 + PARAS.ListIndex)
    Text1.Text = find_parastat(k, X, 0)
    
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




Private Sub DBGRID1_KeyDown(KeyCode As Integer, Shift As Integer)
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
   Dim k As Integer
   Dim mkod(30) As String
   Dim mPoso(30) As Single
   Dim f5 As String
   Dim ELEM(30) As String
   Dim DUM As Integer
   Dim sPOSO As Single
   Dim arSET As Single
   Dim mRow
   Dim R As New ADODB.Recordset
   Dim Sql2 As New ADODB.Recordset
   On Error GoTo MHNYMA
   
   For k = 1 To 30: mkod(k) = "": mPoso(k) = 0: Next
   
   
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
       DUM = FETES_DELIM(f5, ELEM)
       
       ' BΡΙΣΚΩ ΠΟΙΟΥΣ ΚΩΔΙΚΟΥΣ MKOD() ΚΑΙ ΤΙ ΠΟΣΟΤΗΤΕΣ ΕΧΩ MPOSO()
       sPOSO = 0
       For k = 1 To 30
          If Left(ELEM(k), 1) = "#" Then
            DUM = InStr(ELEM(k), ",")
            If DUM - 2 > 0 Then 'AN EXEI KOMMA
               mkod(k) = mID$(ELEM(k), 2, DUM - 2)
               mPoso(k) = Val(mID$(ELEM(k), DUM + 1, Len(ELEM(k)) - DUM))
               sPOSO = sPOSO + mPoso(k)
            End If
          End If
       Next
       arSET = Val(InputBox("Δώστε συνολ.τεμάχια" + Chr(13) + "To σετ αυτό έχει " + Format(sPOSO, "##") + " τεμάχια" + Chr(13) + "αν θέλετε 2 σετ τότε " + Chr(13) + " πληκτρολογήστε " + Format(2 * sPOSO, "##"), "Τεμάχια ανά σετ : " + Format(sPOSO, "##")))
       If sPOSO > 0 Then arSET = arSET / sPOSO
       
   'sxediazo to grid
   DUM = Grid1.row
   For k = 1 To 30
     If mPoso(k) > 0 Then
         EID.RecordSource = "select KOD,ONO,XTI,LTI,LTI5,FPA,MON,MEMO from EID WHERE KOD ='" + mkod(k) + "';"
         EID.Refresh
         If Not EID.Recordset.EOF Then
            Grid1.TextMatrix(DUM + k - 1, f_k) = mkod(k)
            Grid1.TextMatrix(DUM + k - 1, f_p) = mPoso(k) * arSET
            If IsNull(EID.Recordset("mon")) Then m_a = Space(3) Else m_a = EID.Recordset("mon")
            Grid1.TextMatrix(DUM + k - 1, f_o) = Left(EID.Recordset("ono") + Space(60), 60)
            Grid1.TextMatrix(DUM + k - 1, F_M) = m_a
            If F_XONDR = 3 Then 'elga
               Grid1.TextMatrix(DUM + k - 1, f_f) = 5
            Else
               Grid1.TextMatrix(DUM + k - 1, f_f) = EID.Recordset("FPA")
            End If
            Grid1.TextMatrix(DUM + k - 1, f_e) = PEL.Recordset("ekpt")
          If fbuff = "r" Then
              Grid1.TextMatrix(DUM + k - 1, F_T) = EID.Recordset("xti")
          Else
              If lianikh Then
                 Grid1.TextMatrix(DUM + k - 1, F_T) = EID.Recordset("LTI5") / ((100 + g_Fpa(EID.Recordset("FPA"))) / 100)
                 Grid1.TextMatrix(DUM + k - 1, F_mefpa) = EID.Recordset("LTI5")
              Else
                 Grid1.TextMatrix(DUM + k - 1, F_T) = EID.Recordset("LTI")
              End If
          End If
            
            
            Grid1.row = DUM + k - 1
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
   
   
   
   If Left(EID.Recordset("kod"), 1) = "*" Then
   Else
        Grid1.Text = Left(EID.Recordset("ono") + Space(60), 60)  '+ m_a
   End If
   
   
   
   If Left(EID.Recordset("kod"), 1) = "*" Then
     
     Grid1.Text = InputBox("Δώσε περιγραφή", "", Grid1.Text)
   End If
   
   

         
  '   Grid1.Text = InputBox("Δώσε περιγραφή", "", Grid1.Text)
   'End If
   
   
   
   
   
   
   
   
   
   Grid1.Col = f_p  '11-3-07
   Grid1.Text = 1 '11-3-07
   
   If Len(Trim(F_2CHFIA_ZYGIZOMENA)) > 1 And Left(F_2CHFIA_ZYGIZOMENA, 2) = Left(TIMText3.Text, 2) Then
       Grid1.Text = Val(mID(TIMText3.Text, 8, 5)) / 1000
   End If
   
   
   
   
   
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
           Grid1.TextMatrix(Grid1.row, F_T) = R(0)
           
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
    Else  'If fbuff = "r" Then
              Grid1.Col = F_T
              
              If lianikh Then
                 
                 Grid1.Col = F_mefpa
                 Grid1.Text = EID.Recordset("LTI5")
                 Grid1.Col = F_T
                 Grid1.Text = Format(EID.Recordset("LTI5") / (1 + g_Fpa(EID.Recordset("FPA")) / 100), "####0." + String(f_psifiaAjias, "0"))
                 
                 
              Else
                 
                If f_TIMOKAT < 0 Then
                
                Else
                   Grid1.Text = Format(EID.Recordset("LTI"), "####0." + String(f_psifiaAjias, "0"))
                End If
                
              
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
     End If   ' If fbuff = "r" Then
     FindSynolo
 End If  '="#"
   
   ' ΤΟ ΤΕΧΤ3 ΠΑΡΑΛΑΜΒΑΝΕΙ ΤΗΝ ΤΙΜΗ ΑΠΟ ΤΟ GRID
   Grid1.Col = f_p
   TIMText3.Text = Grid1.Text
   TIMText3.Visible = False
   
   ' ΤΟ ΤΕΧΤ3 ΠΑΡΑΛΑΜΒΑΝΕΙ ΤIΣ ΔΙΑΣΤΑΣΕΙΣ ΑΠΟ ΤΟ GRID
   TIMText3.Top = Grid1.Top + Grid1.CellTop
   TIMText3.Left = Grid1.Left + Grid1.CellLeft
   TIMText3.width = Grid1.CellWidth - 15
   TIMText3.height = Grid1.CellHeight - 25
   TIMText3.Visible = True
   TIMText3.SetFocus
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
 ElseIf F_EIKONA = 111 Then ' kodikos .jpg
     PI = "C:\MERCVB\IMAGES\" + Trim(CNull(EID.Recordset("kod"))) + "*" + ".JPG"
     
     
     
     
  End If
  
 
 If Len(Dir(PI)) > 0 Then
    
    Image1.Picture = LoadPicture("C:\MERCVB\IMAGES\" + Dir(PI))
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

If Err.Description = "Automation error" Then
   'Gdb.Close
   'Gdb.Open gConnect
End If

Resume Next
            

End Sub
Function FETES_DELIM(LINE, ELEM)
Dim KL, KE

On Error GoTo MHNYMA
 '  DIABAZO SE PINAKA OLA TA STOIXEIA THS GRAMHS
     For KE = 1 To 30
       ELEM(KE) = ""
     Next
     
     KL = 0   ' metraei xaraktires
     KE = 0   ' metritis toy pinaka  ELEMENT
     Do
        KE = KE + 1  ' metritis toy pinaka  ELEMENT
        KL = KL + 1  ' metraei xaraktires
        
        Do While mID$(LINE, KL, 1) <> Chr(13) ' tab
            ELEM(KE) = ELEM(KE) + mID$(LINE, KL, 1)
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
   
   
 If F_POLHTES = 1 Then
       If Not IsNull(PEL.Recordset("PVLHTHS")) Then
         On Error Resume Next
         ' Combo8.ListIndex = Val(PEL.Recordset("NUM4"))
          Combo8.ListIndex = Val(PEL.Recordset("PVLHTHS")) - 1
       End If
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
   


 EID.RecordSource = "select METAF,ATIM,CONVERT(CHAR(10),HME,103) AS [Ημερομηνία],ID_NUM from TIM where left(ATIM,1) in ('" + Get_meKomma(f_metasx) + "') and KPE='" + PEL.Recordset("kod") + "' and ART IS null ORDER BY HME;"
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
   TIMText3.Top = Grid1.Top + Grid1.CellTop
   TIMText3.Left = Grid1.Left + Grid1.CellLeft
   TIMText3.width = Grid1.CellWidth - 15
   TIMText3.height = Grid1.CellHeight - 25
   
110
   
'- --------------- ΔΙΟΡΘΩΣΗ ΤΙΜΟΛΟΓΙΟΥ ----------------------------------
Dim re As New ADODB.Recordset, ANS As Integer
Dim rEgg As New ADODB.Recordset

f_dat = DTPicker1.value
f_matim = Right(PARAS.Text, 1) + Right("0000000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text
ANS = vbNo
re.Open "select * from TIM where  HME>='" + Format(f_dat, "mm/dd/yyyy") + "' AND HME<='" + Format(DateAdd("D", 1, f_dat), "mm/dd/yyyy") + "'  and ATIM='" + f_matim + "'", Gdb, adOpenDynamic, adLockOptimistic
        If Not re.EOF Then
           ANS = MsgBox("Υπάρχει ήδη  παραστατικό με ίδια ημερομηνία και ίδιο αριθμό. " + Chr(13) + "Να διορθωθεί;", vbYesNo)
        Else
             'ΣΒΗΝΩ ΤΥΧΟΝ ΥΠΟΛΕΙΜΑΤΑ
             Gdb.Execute "DELETE  from EGGTIM where  HME>='" + Format(f_dat, "mm/dd/yyyy") + "' AND HME<='" + Format(DateAdd("D", 1, f_dat), "mm/dd/yyyy") + "'  and ATIM='" + f_matim + "'"
        End If
  
120






If ANS = vbYes Then

'βαζει τον αρχικό πελάτη
   PEL.RecordSource = "select *from PEL where EIDOS='" + Left(f_gkod_pel, 1) + "' and KOD='" + re("kpe") + "'"
   PEL.Refresh
   Text2(0).Text = PEL.Recordset("KOD")
   Text2(1).Text = PEL.Recordset("EPO")
   Text2(2).Text = PEL.Recordset("afm")
   Text2(3).Text = PEL.Recordset("die")


parat.Text = re("parat")
AYTOKINHTO.Text = re("aytok")

PROORISMOS.Text = re("PROORISMOS")
SKOPOS.Text = re("SKOPOS")
FORTOSH.Text = re("FORTOSH")
' For k = 0 To Combo4.ListCount - 1
'    If InStr(Combo4.List(k), re("TRP")) > 0 Then
'       Combo4.Text = Combo4.List(k)
'       Exit For
'    End If
' Next

If re("TRP") = "ΜΕ" Then ' METRHTOIS
      Combo4.Text = Combo4.List(0)
Else
      Combo4.Text = Combo4.List(1)
End If







   f_dior_mAtim = f_matim ' ΤΑ ΚΡΑΤΑΩ ΜΗΝ ΤΥΧΟΝ ΤΑ ΑΛΛΑΞΕΙ ΣΤΗΝ Φ’ΣΗ ΤΗΣ ΔΙΟΡΘΩΣΗΣ
   f_dior_HME = f_dat 'ΓΙΑ ΝΑ ΜΠΟΡΕΣΩ ΝΑ ΣΒΗΣΩ ΤΟ ΑΡΧΙΚΟ ΤΙΜΟΛΟΓΙΟ


''το είχα έτσι αλλά έχανε τα είδη  αν  εβαζα λάθος τις αποθήκες
   ' EID.RecordSource = "select EGGTIM.* from EGGTIM  WHERE HME >= '" + Format(f_dat, "mm/dd/yyyy") + "' AND HME <= '" + Format(DateAdd("D", 1, f_dat), "mm/dd/yyyy") + "' and ATIM='" + f_matim + "' AND APOT=" + Str(Val(APOTH1)) + " order by PROOD_AJ"
   EID.RecordSource = "select EGGTIM.* from EGGTIM  WHERE HME >= '" + Format(f_dat, "mm/dd/yyyy") + "' AND HME <= '" + Format(DateAdd("D", 1, f_dat), "mm/dd/yyyy") + "' and ATIM='" + f_matim + "'  order by PROOD_AJ"
   
   EID.Refresh
   '   On Error Resume Next
   EID.Recordset.MoveFirst
   Dim k: k = 1
130

Dim DUMY
Dim M_APOT

' ΟΤΑΝ ΕΧΩ ΕΝΔΟΔΙΑΚΙΝΗΣΗ ΝΑ ΜΗΝ  ΦΕΡΝΕΙ ΔΙΠΛΑ ΤΑ ΕΙΔΗ ΑΛΛΑ ΝΑ ΦΕΡΝΕΙ ΜΟΝΟ ΑΠΟ ΤΗΝ ΜΙΑ ΑΠΟΘΗΚΗ
M_APOT = EID.Recordset("APOT")

   Do While Not EID.Recordset.EOF
      
     If M_APOT = EID.Recordset("APOT") Then
      
      XROMATA(k, 1, 0) = Left(EID.Recordset("PROELEYSH"), 3)
      XROMATA(k, 0, 1) = mID(EID.Recordset("PROELEYSH"), 4, 2)
        
      
      Grid1.TextMatrix(k, f_k) = EID.Recordset("kode")
      Grid1.TextMatrix(k, f_p) = EID.Recordset("POSO")
      Grid1.TextMatrix(k, f_e) = EID.Recordset("ekpt")
      Grid1.TextMatrix(k, f_o) = EID.Recordset("onoma")
      Grid1.TextMatrix(k, F_T) = EID.Recordset("TIMM")
      Grid1.TextMatrix(k, F_M) = EID.Recordset("MONA")
      Grid1.TextMatrix(k, f_f) = EID.Recordset("FPA")
      
      Grid1.TextMatrix(k, F_mikta) = EID.Recordset("mikta") ' χρειαζονται για το συγκεντροτικο
      
      
      Grid1.TextMatrix(k, f_a) = Round(EID.Recordset("POSO") * EID.Recordset("TIMM"), 2)
      
      
      Grid1.row = k: Grid1.Col = 0:  FindSynolo
      
140
      k = k + 1
      
     End If
     
      EID.Recordset.MoveNext
   Loop
   Grid1.row = k
   If k > 1 Then
      TIMText3.Text = Grid1.TextMatrix(k + 1, f_k)
   End If
   
   F_DIOR = 1
   Grid1.Col = f_k  ' για να μην ζητάει να ανοιξει νέα καρτέλλα
End If
   

   
   
   
   
   
   
150
   

   
   
   
   
   
   
   
   
   
   
   
 TIMText3.Text = Grid1.Text
 TIMText3.SetFocus
   
  
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
Dim DELT As Recordset, k As Integer, L As Integer
Dim e, S As String, d As String, t As Recordset
Dim X, C, Eidh As Recordset
Dim R As Recordset
On Error GoTo MHNYMA
Dim r2 As New ADODB.Recordset




k = 0
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
    k = k + 1
    f_deltia(k) = FlxGridDeltion.TextMatrix(L, 2) + FlxGridDeltion.TextMatrix(L, 3)
    f_IDdeltia(k) = FlxGridDeltion.TextMatrix(L, 4)
    S = S + "'" + f_deltia(k) + "'" + ","
    SX = SX + Format(Val(mID(f_deltia(k), 2, 6)), "######") + ","
  End If
Next
S = S + "''"

If f_metasx = "a" Then

    If Len(Dir("C:\HRAKLHS.TXT")) > 1 Then
       parat.Text = "ΔΕΛΤ.ΠΑΡΑΛ. " + SX
    Else
       parat.Text = "Αρ.Παραγγ." + SX
    End If

Else
   parat.Text = "Σχετ.Δελτ." + SX
End If
fSHMADEMENA = S

If k = 0 Then  ' DEN SHMADECE TIPOTA
   FlxGridDeltion.Visible = False
   GridEidon.Visible = True
   Exit Sub
End If
d = EID.RecordSource
30

' σημάδεψε δελτια // ψάχνω να βρώ τα είδη
If k = 1 Then
 '"order by PROOD_AJ"
   EID.RecordSource = "select POSO as SPOSO,KODE,TIMM,ONOMA from EGGTIM where ATIM+CONVERT(CHAR(10),HME,103)  IN (" + S + ") order by PROOD_AJ;"
   EID.Refresh

Else ' polla deltia
   EID.RecordSource = "select sum(POSO) as SPOSO,KODE,TIMM,ONOMA from EGGTIM where ATIM+CONVERT(CHAR(10),HME,103)  IN (" + S + ") group by KODE,TIMM,ONOMA;"
   EID.Refresh
End If

'EID.RecordSource = "select sum(POSO) as SPOSO,KODE,TIMM,ONOMA from EGGTIM where ATIM+CONVERT(CHAR(10),HME,103)  IN (" + S + ") group by KODE,TIMM,ONOMA;"
  ' EID.Refresh


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
k = 1
Do While Not EID.Recordset.EOF
  
  EID2.Open "SELECT *FROM EID WHERE KOD='" + EID.Recordset("kode") + "'", Gdb, adOpenDynamic, adLockOptimistic
  Grid1.TextMatrix(k, f_k) = EID.Recordset("kode")
  Grid1.TextMatrix(k, f_p) = EID.Recordset("SPOSO")
  
  If Left(Grid1.TextMatrix(k, f_k), 1) = "*" Then
     Grid1.TextMatrix(k, f_o) = EID.Recordset("ONOma")
  Else
     Grid1.TextMatrix(k, f_o) = EID2("ONO")
  End If
  Grid1.TextMatrix(k, F_M) = EID2("MON")
  
  If EID.Recordset("TIMM") = 0 Then
      Grid1.TextMatrix(k, F_T) = EID2("LTI")
  Else
      Grid1.TextMatrix(k, F_T) = EID.Recordset("TIMM") ' ΠΑΡΑΓΓΕΛΙΑ
  End If
  
  
'******************   15-6-2010  ************************************
'ΕΚΠΤΩΣΕΙΣ
Grid1.Col = f_e
If f_TIMOKAT > 0 And Not IsNull(PEL.Recordset("pek")) Then
   r2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + Str(PEL.Recordset("pek")) + " AND KOD='" + Grid1.TextMatrix(k, f_k) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
   If r2.EOF Then
      Grid1.Text = ""
   Else
    '  Grid1.Text = r2(0)
    Grid1.TextMatrix(k, f_e) = r2(0)
    End If
   r2.Close
   
  'ΔΟΥΛΕΥΕΙ ΜΕ ΚΑΡΦΩΤΗ ΤΙΜΗ
 ElseIf f_TIMOKAT < 0 And Not IsNull(PEL.Recordset("pek")) Then
        r2.Open "select EKPT FROM TIMOKAT WHERE TIMOK=" + Str(PEL.Recordset("pek")) + " AND KOD='" + Trim(Left(EID.Recordset("kod"), 13)) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
        If r2.EOF Then
           ' Grid1.TextMatrix(k, F_T) = 0
        Else
           Grid1.TextMatrix(k, F_T) = R(0)
        End If
        r2.Close
Else
   Grid1.TextMatrix(k, f_e) = PEL.Recordset("pek")
   r2.Close
   
End If
'******************   15-6-2010  ************************************
  
  
  
  
  Grid1.TextMatrix(k, f_f) = EID2("FPA")
  Grid1.TextMatrix(k, f_a) = Round(EID.Recordset("SPOSO") * Grid1.TextMatrix(k, F_T), 2) * (100 - Val(Grid1.TextMatrix(k, f_e))) / 100
  
  Grid1.row = k: Grid1.Col = 0:  FindSynolo
  
  
  
  EID2.Close
  k = k + 1
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
   TIMText3.Top = Grid1.Top + Grid1.CellTop
   TIMText3.Left = Grid1.Left + Grid1.CellLeft
   TIMText3.width = Grid1.CellWidth - 15
   TIMText3.height = Grid1.CellHeight - 25
   TIMText3.Text = Grid1.Text
   TIMText3.SetFocus
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

'    GetLocation Me
'    CenterForm Me
'    ResizeForm Me

mForm_Load PAR1, fh, fw, ft, fl



Dim k, L As Integer, X As String, db
Dim R As New ADODB.Recordset
 
 On Error GoTo MHNYMA
  
' R.Open "SELECT top 1 *FROM TIM", Gdb, adOpenDynamic, adLockOptimistic
 
' R.Close
 
' par1.SHOW



'R.Open "select * FROM PARAMETROI", Gdb, adOpenDynamic, adLockBatchOptimistic
'
'For K = 1 To 200
'   F_AR_parametroi(K, 0) = R("FORMA")
'   F_AR_parametroi(K, 1) = R("VAR")
'   F_AR_parametroi(K, 2) = R("TIMH")
'   R.MoveNext
'   If R.EOF() Then
'     Exit For
'   End If
'Next
'
'R.Close


If gXEIRISTHS < 2 Then
   TDBGrid1.Visible = False
End If


f_pelMERC = FindParametroi("PAR1", "F_PELMERC", "..", "ΧΡΗΣΤΗΣ MERCURY")


If Left(f_pelMERC, 3) = "SOK" Then
'   Text14.Visible = True
   MSFlexGrid1.Visible = True
   MSFlexGrid1.row = 0
   
   MSFlexGrid1.Col = 0
   MSFlexGrid1.Text = "ΧΡΩΜΑ"
   
   MSFlexGrid1.Col = 1
   MSFlexGrid1.Text = "BARCODE"
   
   MSFlexGrid1.Col = 2
   MSFlexGrid1.Text = "ΥΠΟΛΟΙΠΟ"
   
   
   
   
   
End If
  
  
 F_akyr_Delete = Val(FindParametroi("PAR1", "F_akyr_Delete", "1", "Γ.31 Σβήνω το παραστατικό που ακυρώνεται=1")) ' δεν ακυρωνω το παραστατικό που ακυρώνεται
 
 F_add_barc = Val(FindParametroi("PAR1", "add_barc", "1", "Γ.16 Barcode έιναι 1.Κωδ 2.Βοηθ.Κλειδί"))
 
 PAR1.KeyPreview = True
 
 f_PALIAXRONIA = Trim(FindParametroi("PAR1", "F_PALIAXRONIA", "2005", "Γ.18 DSN ΠΑΛΙΑΣ ΧΡΟΝΙΑΣ"))
 F_EXO_XROMATA = Val(FindParametroi("PAR1", "F_EXO_XROMATA", "0", "Γ. 1 ΧΡΗΣΙΜΟΠΟΙΩ ΜΕΓΕΘΗ-ΧΡΩΜΑΤΑ=1")) 'posa psifia tha exei h kathe seira
  
  
  F_POLHTES = Val(FindParametroi("PAR1", "F_POLHTES", "0", "Δ. 1=ΠΩΛΗΤΗΣ ΠΕΛΑΤΗ->ΤΙΜΟΛΟΓΙΟ  0=ΧΩΡΙΣ ΠΩΛΗΤΕΣ"))
  
  f_Fkey = Trim(FindParametroi("PAR1", "F_FKEY", "112", "ΤΟ ΠΛΗΚΤΡΟ ΠΟΥ ΧΡΗΣΙΜΟΠΟΙΩ ΓΙΑ UPDATE"))
 
 F_EIKONA = Val(FindParametroi("APOT2", "F_EIKONA", "1", "1=EIKONA=>ΚΩΔΙΚΟΣ 2=>ΜΕ ΒΟΗΘ.ΚΛΕΙΔΙ 11=ΚVDIKOS.JPG 12=BOH.JPG 0=ΟΧΙ ΕΙΚΟΝΕΣ"))

 F_PAR_MONO_ME_YPOLOIPO = Val(FindParametroi("PAR1", "F_PAR_MONO_ME_YPOLOIPO", "0", "Γ. 30 ΠΑΡΑΓΓΕΛΙΑ ΠΕΛΑΤΗ ΜΟΝΟ ΑΝ ΥΠΑΡΧΕΙ ΥΠΟΛΟΙΠΟ"))

 F_STOPDELTIA = Val(FindParametroi("PAR1", "F_STOPDELTIA", "0", "Α. 2.Αποτρέπει την τιμολόγηση αν εκκρεμούν δελτία=1")) 'posa psifia tha exei h kathe seira

 F_2CHFIA_ZYGIZOMENA = FindParametroi("PAR1", "F_2CHFIA_ZYGIZOMENA", "  ", "Γ. 2. ΠΡΩΤΑ ΨΗΦΙΑ ΒΟΗΘ.ΚΩΔ.ΤΩΝ ΖΥΓΙΖΟΜΕΝΩΝ") '2 ΠΡΩΤΑ ΨΗΦΙΑ ΒΟΗΘ.ΚΩΔ.ΤΩΝ ΖΥΓΙΖΟΜΕΝΩ

 If Len(Trim(F_2CHFIA_ZYGIZOMENA)) > 1 Then
   Grid1.ColWidth(7) = 10
   Grid1.ColWidth(8) = 10
 
 End If


f_Desmia = Val(FindParametroi("PAR1", "F_DESMIA", "0", "Γ. 3.ΟΙ ΠΑΡΑΓΓΕΛΙΕΣ ΥΠΟΛΟΓΙΖΟΝΤΑΙ ΣΤΟ ΥΠΟΛΟΙΠΟ=1 ")) ' χρησιμοποιω τις παραγγελίες για έλεγχο υπολοίπου  'βλεπε mdesmia
 
 
f_mono_apoParaggelia = Val(FindParametroi("PAR1", "f_mono_apoParaggelia", "0", "Γ. 3.5 ΤΙΜΟΛΟΓΩ ΜΟΝΟ ΑΝ ΥΠΑΡΧΕΙ ΠΑΡΑΓΓΕΛΙΑ "))  '
f_SYMPYKNVSH = Val(FindParametroi("PAR1", "f_SYMPYKNVSH", "0", "Γ. 3.6 ΣΥΜΠΥΚΝΩΣΗ ΣΤΟΝ ΙΔΙΟ ΚΩΔΙΚΟ "))   '

F_METRHTA = Val(FindParametroi("PAR1", "f_METRHTA", "0", "Γ. 3.9 ZHTA METPHTA=1 OXI=0 "))   '

F_TRANS = Val(FindParametroi("PAR1", "F_TRANS", "0", "Γ. 3.7 ΔΟΥΛΕΥΕΙ TRANSACTIONS=1  "))   '


F_MONOMEYPOLOIPO = Val(FindParametroi("PAR1", "F_MONOMEYPOLOIPO", "0", "Γ. 9.ΕΜΦΑΝΙΣΗ ΜΟΝΟ ΤΩΝ ΕΙΔΩΝ ΜΕ ΥΠΟΛΟΙΠΟ=1")) '
'  F_EIDH_APOT = Val(FindParametroi("PAR1", "F_EIDH_APOT", "0", "Γ. 9B. ΕΜΦΑΝΙΣΗ ΜΟΝΟ ΤΩΝ ΕΙΔΩΝ ΜΕ ΚΙΝΗΣΗ ΣΤΟ ΥΠΟΚΑΤΑΣΤΗΜΑ=1")) '
F_YPOL_GEN = Val(FindParametroi("PAR1", "F_YPOL_GEN", "1", "Γ. 10.ΕΜΦΑΝΙΣΗ ΟΛΩΝ ΤΩΝ ΕΙΔΩΝ =1 ΜΟΝΟ ΜΕΚΙΝΗΣΗ=0")) ' NA EMFANIZEI TA EIDH ASXETA AN EXOYN KINHΘΕΙ


F_DEF_CAR = Val(FindParametroi("PAR1", "F_DEF_CAR", "0", "Γ. 11.ΑΡΙΘ.ΠΡΟΕΠΙΛΕΓΜΕΝΟΥ ΑΥΤΟΚ/ΤΟΥ 0=ΚΑΝΕΝΑ")) '




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


 
 
 
 
 
 
 
' On Error GoTo MHNYMA
10
geidos_paras = 1
' On Error GoTo 0
'tropos plhromhs
'R.Close
R.Open "SELECT *FROM PINAKES WHERE TYPOS=12 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
   k = 0
Do While Not R.EOF
   If R("typos") = 12 Then
      Combo4.AddItem Str(R("AYJON")) + "." + R("PERIGRAFH")
      F_LOGPLHR(k) = Format(R("TIMH"), "0000")
      k = k + 1
   End If
   R.MoveNext
Loop
 R.Close
k = 0





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

f_Timol2_SeiresTimologioy = Val(FindParametroi("PAR1", "f_TIMOL2_SEIRESTIMOLOGIOY", Str(f_MinSeiresTimologioy), "Β. 4 Σειρές τιμ.Πώλησης φορμας TIMOL2"))


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
F_File2_Timologioy = FindParametroi("PAR1", "F_File2_Timologioy", "c:\print", "Α.12b Αρχείο εκτύπωσης παραστ.Υποκ")

F_Type_File_Timologioy = Val(FindParametroi("PAR1", "F_Type_File_Timologioy", "1", "Α.13 1=εκτύπωση αρχείου παρ/κού 0=Οχι 2=Print 3=c:\typ.bat"))
F_Type2_File_Timologioy = Val(FindParametroi("PAR1", "F_Type2_File_Timologioy", "1", "Α.13b 1=εκτύπ.αρχείου παρ/κού υποκ. 0=Οχι 2=Print 3=c:\typ.bat"))



f_F12 = Val(FindParametroi("PAR1", "F_F12", "-1", "Α.14 F12:Τύπος Παραστατικού που ενημερώνεται"))
f_F1 = Val(FindParametroi("PAR1", "F_F1", "-1", "Α.15 F1:Τύπος Παραστατικού που ενημερώνεται"))

f_BARCODE_STOTIM = Val(FindParametroi("PAR1", "F_BARCODE_STOTIM", "0", "Α.16  1=ΒΑΖΕΙ ΤΟ ERG ΣΤΟ ΤΙΜΟΛΟΓΙΟ 0=ΤΟΝ ΚΩΔΙΚΟ(def)"))


f_CHECK_PARAGG = Val(FindParametroi("PAR1", "F_CHECK_PARAGG", "0", "Α.17  1=<<ΤΣΕΚΑΡΙΣΜΕΝΟ ΑΠΟ ΠΑΡΑΓΓΕΛΙΑ>> 0=ΟΧΙ"))

If f_CHECK_PARAGG = 1 Then
   ApoParaggelia.value = vbChecked
End If


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
 If F_DEF_CAR > 0 Then
      AYTOKINHTO.Text = AYTOKINHTO.List(F_DEF_CAR)
 End If
 

' AN EINAI 1o YPOKATASTHMA
If Len(Dir("C:\KENTRIKO.TXT")) = 0 Or Len(Dir("C:\LAGEURO\MPOYGET.EXE", vbNormal)) = 0 Then
   APOTH1.Text = APOTH1.List(1)
   APOTH2.Text = APOTH2.List(0)
   f_YPOK = 1
'   FORTOSH.Text = FORTOSH.List(1)
Else
   APOTH1.Text = APOTH1.List(0)
   APOTH2.Text = APOTH2.List(1)
   f_YPOK = 0
End If

If Len(Dir("C:\KENTRIKO.TXT")) > 0 Or gUserId = 1 Then
   APOTH1.Text = APOTH1.List(0)
   APOTH2.Text = APOTH2.List(1)
   f_YPOK = 0
End If

If Len(Dir("C:\YPOK2.TXT")) > 0 Then
   APOTH1.Text = APOTH1.List(2)
   APOTH2.Text = APOTH2.List(0)
   f_YPOK = 2
End If

If Len(Dir("C:\YPOK3.TXT")) > 0 Then
   APOTH1.Text = APOTH1.List(3)
   APOTH2.Text = APOTH2.List(0)
   f_YPOK = 3
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
 
 
 
 
Grid1.ColS = 11

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
For k = 1 To Grid1.ROWS - 1
      
   Grid1.row = k
   Grid1.Text = Format(k, "##")
   Grid1.RowHeight(k) = 250
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
    
    
LASTTIMOL.ConnectionString = gConnect
LASTTIMOL.RecordSource = "SELECT top 550 HME,ATIM,AJI AS [AΞΙΑ],PEL.EPO AS [ΕΠΩΝΥΜΙΑ],PEL.KOD AS [ΚΩΔΙΚΟΣ],PARAT AS [ΣΧΟΛΙΑ],SXETIKO AS [ΣΧΕΤ.ΠΑΡ] FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD ORDER BY ID_NUM DESC"
LASTTIMOL.Refresh
TDBGrid1.MarqueeStyle = 3  'SEIRA MARKARISMENH
TDBGrid1.Splits(0).AlternatingRowStyle = True 'ENALAX XROMATA SEIRON
TDBGrid1.Splits(0).ExtendRightColumn = True

TDBGrid1.Columns(6).width = 1544.882
TDBGrid1.Columns(5).width = 2610.142
TDBGrid1.Columns(4).width = 480.189
TDBGrid1.Columns(3).width = 2550.142
TDBGrid1.Columns(2).width = 929.7639
TDBGrid1.Columns(1).width = 900.2835
TDBGrid1.Columns(0).width = 1049.953
    
    
    
'TIMEDIT.Enabled = True
EID.Enabled = True
PEL.Enabled = True
    
DTPicker1.value = Now

    
PAR1.Picture = LoadPicture(gPicture)
    
    

 
 
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
 For k = 1 To Len(X)
    SEIRA.AddItem mID$(X, k, 1)
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


PARAS.Enabled = True
' PARAS.ListIndex = geidos_paras  ' ""
 
 
 'DbGrid1.Columns(2).NumberFormat = gDek_pos 'δεκαδικά ποσοτήτων
 'DbGrid1.Columns(3).NumberFormat = gDek_ajia 'δεκαδικά αξίας
 'DbGrid1.Columns(6).NumberFormat = gDek_ajia 'δεκαδικά αξίας
 
  
 'PARAS.SetFocus
100
  
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

Dim k As Integer
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



 Sub PAR2SHOWING()

 
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
 PAR1.LASTTIMOL.Refresh
  
  
  

  
  
'  SendKeys "~" 'ENTER
'  SendKeys "~" 'enter  {DOWN}"
' SendKeys "~" 'ENTER
 SendKeys "~" 'ENTER
''FF.Command1_Click






End Sub





Private Sub SynoloG_Click()
Dim C As String
    C = InputBox("τελικό ποσό", "Αλλαγή τελικού ποσού", SynoloG.Caption)
    Text5.Text = "*" + C
    Text5_LostFocus
b_Katax.SetFocus
    
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
'================================================================================
Dim SUMES(100) As String
Dim k As Long

 On Error Resume Next
'
' 'edo parakato exei latos 15/8/2009
'
' For k = 0 To Data1.Recordset.FIELDS.Count - 1
'        SUMES(k) = TDBGrid.Splits(0).Columns(k).FooterText '  = Format(SUMES(k), "######0.00")
'  Next
'
'
'
    
 If LASTTIMOL.Recordset.Sort = "[" & TDBGrid1.Columns(ColIndex).DataField & "] asc" Then   ' strSort
    LASTTIMOL.Recordset.Sort = "[" & TDBGrid1.Columns(ColIndex).DataField & "] desc" ' strSort
 Else
    LASTTIMOL.Recordset.Sort = "[" & TDBGrid1.Columns(ColIndex).DataField & "] asc" ' strSort
 End If
  
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then

'MDIForm1.runt.Item(0).Caption = "Παρουσίαση"
'MDIForm1.runt.Item(1).Caption = "Διόρθωση"
On Error Resume Next

   Call Me.PopupMenu(MDIForm1.par1menu)
  ' LASTTIMOL.Refresh
   
   
' μεσω drop down menu
'MdiForm1
   
   
   
'ΕΚΤΥΠΩΣΗ
'   par1.diortosis
'  DoEvents
'  par1.Command5_Click
   
   
'ΔΙΟΡΘΩΣΗ
 'par1.diortosis
   
'ΠΑΡΟΥΣΙΑΣΗ
'par1.PAR2SHOWING
   
'ΜΕΤΑΣΧΗΜΑΤΙΣΜΟΣ
'par1.metasx.value = vbChecked
'   par1.diortosis
   
   
   
   
End If

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
   
   TIMText3.Text = Grid1.Text
   TIMText3.Visible = False
   TIMText3.Top = Grid1.Top + Grid1.CellTop
   TIMText3.Left = Grid1.Left + Grid1.CellLeft
   TIMText3.width = Grid1.CellWidth - 15
   TIMText3.height = Grid1.CellHeight - 25
   
   TIMText3.SelStart = 0
   TIMText3.SelLength = Len(TIMText3.Text)
   

   
   
   
   
   TIMText3.Visible = True
   TIMText3.SetFocus
   
   
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
    PEL.RecordSource = "select LEFT(left(EPO+'-'+DIE+'-',56)+KOD,60) as xx,EPA,THL,DIE,KOD  from PEL where EIDOS='" + X + "' and EPO like '" + UCase(SameLetters(Text2(1).Text)) + "%' order by EPO;"
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
     
     TIMText3.BackColor = vbYellow
     TIMText3.SelStart = 0
     TIMText3.SelLength = Len(TIMText3.Text)


End Sub

Private Sub timText3_KeyDown(KeyCode As Integer, Shift As Integer)
'ΕΝΤΡΥ ΤΙΜΟΛΟΓΙΟΥ
Dim DUM
Dim NNN As Long

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
Dim snc1, snc2, snN1
Dim Synt5
Dim RSOK As New ADODB.Recordset
   Dim P, t, e, a, C
   Dim KOMA, mText
   Dim R0 As New ADODB.Recordset
   Dim XX As String
Dim DD2 As String

   Dim intRecord As Integer
   Dim intField As Integer
Dim ELEM(200) As String
Dim DUM


Dim Sql2 As String
Dim HDH_FORTOMENA As Single, LK0 As Long, n_HDH_FORTOMENA

 HDH_FORTOMENA = 0

Dim NNN


On Error Resume Next

If Len(Text2(1).Text) = 0 Then
   MsgBox "Δεν επιλέξατε Πελάτη / Προμηθευτή"
   Text2(1).SetFocus
   Exit Sub
End If





On Error GoTo MHNYMA
TIMText3.Text = Replace(TIMText3.Text, "'", "~")

If KeyCode = 46 And Grid1.Col = f_k Then ' DELETE ΣΒΗΝΩ ΤΗΝ ΣΕΙΡΑ
               Grid1.TextMatrix(Grid1.row, f_k) = ""
               Grid1.TextMatrix(Grid1.row, f_o) = ""
               Grid1.TextMatrix(Grid1.row, f_p) = ""
               Grid1.TextMatrix(Grid1.row, F_T) = ""
               Grid1.TextMatrix(Grid1.row, f_e) = ""
               Grid1.TextMatrix(Grid1.row, f_f) = ""
               Grid1.TextMatrix(Grid1.row, f_a) = ""
               Grid1.TextMatrix(Grid1.row, F_mefpa) = ""
               Exit Sub
End If

  If KeyCode = 27 Then 'esc τελος εισαγωγής
       KeyCode = 0
       Combo4.SetFocus
  End If



If KeyCode = 114 And Shift = 0 Then ' --------------  F3 καταχωρηση SN ----------------------
  'snc1, snc2, snN1
  If Grid1.TextMatrix(Grid1.row, f_k) Then
     On Error Resume Next
     snc2 = ""
     Do While True
        snc1 = InputBox(snc2 + "δωσε το sn")
        If Len(Trim(snc1)) = 0 Then
           Exit Do
        Else
           a = "insert INTO SN (ARIU,KOD,HME,KODPEL,ATIM) VALUES ("
           a = a + "'" + snc1 + "',"
           a = a + "'" + Grid1.TextMatrix(Grid1.row, f_k) + "',"
           a = a + "'" + Format(DTPicker1.value, "MM/DD/YYYY") + "',"
           a = a + "'" + Text2(0).Text + "',"
           a = a + "'" + Right(PARAS.Text, 1) + Right("00000" + Trim(Text1.Text), f_psifia_atim) + SEIRA.Text + "')"
           Gdb.Execute a, snN1
           If snN1 = 1 Then
               snc2 = snc2 + snc1 + Chr(13)
           Else
               MsgBox "δεν κατεχωρήθη"
           End If
        End If
     Loop
     
  End If
  
  
  Exit Sub
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
   
   
  UGridEdit.LTI.Caption = EID.Recordset("LTI")
  UGridEdit.xti.Caption = EID.Recordset("XTI")
  UGridEdit.LTI5.Caption = EID.Recordset("LTI5")
  UGridEdit.kod.Caption = EID.Recordset("KOD")
  UGridEdit.ono.Caption = EID.Recordset("ONO")
  UGridEdit.CH1.Caption = EID.Recordset("CH1")
  UGridEdit.hme.Caption = Format(DTPicker1.value, "DD/MM/YYYY")
  
  UGridEdit.SHOW 1
    
  '000000027311
' 27311
  
  
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
   TIMText3.Text = SUMA_XROMATA
Else
   Grid1.TextMatrix(Grid1.row, f_p) = SUMA_XROMATA
End If
If SUMA_XROMATA > 0 Then
    XROMATA(Grid1.row, 0, 0) = -0.01 ' ΣΗΜΑΔΙ ΟΤΙ ΧΡΗΣΙΜΟΠΟΙΗΣΕ ΤΟ F6
End If

  
  
  Exit Sub
End If


' --------------  F6 ------ xroma-megetos -----------------------------------

If KeyCode = Val(Right(f_Fkey, 3)) And Shift = 0 Then ' Val(Left(f_Fkey, 1)) Then ' --------------  F1 ----------------------
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
     
   
   
   
   
   If Len(Trim(TIMText3.Text)) = 0 Then   ' enter me keno koytaki  ***************
      Grid1.Text = TIMText3.Text ' κρατάει το grid το κείμενο
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
     
     If Grid1.Col < f_p And (Grid1.Text <> TIMText3.Text Or Val(Grid1.TextMatrix(Grid1.row, f_p)) = 0) Then  ' αλλαξε ο κωδικός άρα ψάξε στο αρχείο
         
          ' ψαχνει κατά κωδικό   %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
10       If Grid1.Col = f_k Then
            TIMText3.Text = Trim(Left(TIMText3.Text, 16))
            
            
            If Left(TIMText3.Text, 1) = "." Then ' $$$$$$$
                 'EID.RecordSource = "select SUM(XRE)-SUM(PIS) AS [ΥΠΟΛ],EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO from EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD WHERE ERG  LIKE '" + Trim(Mid(timtext3.Text, 2, 14)) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) ;"
                 
                ' ψαχνει κατά BOHUHTIKO κωδικό  ERG %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                 EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)" _
                 & "-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," _
                 & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                 & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" _
                 & " WHERE EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND ERG  LIKE '" + SameLetters(Trim(mID(TIMText3.Text, 2, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1 "
               
                 If F_MONOMEYPOLOIPO = 1 And f_pol = 1 Then
                   EID.RecordSource = EID.RecordSource + " HAVING [ΥΠΟΛ]>0"
                 End If
                                   
                 EID.Refresh
                'ΑΝ ΔΕΝ ΕΧΕΙ ΚΙΝΗΘΕΙ ΚΑΘΟΛΟΥ ΝΑ ΜΠΟΡΕΙ ΝΑ ΔΕΙΞΕΙ
                If EID.Recordset.EOF Then
                  EID.RecordSource = "SELECT KOD AS [ΚΩΔ]," _
                  & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                  & " from EID " _
                  & " WHERE  ERG  LIKE '" + SameLetters(Trim(mID(TIMText3.Text, 2, 14))) + "%'  "
                End If
                 
                 
                 
               
               
' & " WHERE EGGTIM.APOT=" + left(APOTH1.Text, 1) + " AND ERG  LIKE '" + SameLetters(Trim(Mid(TIMText3.Text, 2, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) ;"
               
              ' ψαχνει κατά κωδικό  Εργοστασίου  KODERG %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            ElseIf Left(TIMText3.Text, 1) = "+" Then ' $$$$$$$
                 'EID.RecordSource = "select SUM(XRE)-SUM(PIS) AS [ΥΠΟΛ],EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO from EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD WHERE ERG  LIKE '" + Trim(Mid(timtext3.Text, 2, 14)) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) ;"
               
                 EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)" _
                 & "-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," _
                 & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                 & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" _
                 & " WHERE EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND KODERG = '" + Trim(mID(TIMText3.Text, 2, 14)) + "' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1 ;"
                 
                 EID.Refresh
                'ΑΝ ΔΕΝ ΕΧΕΙ ΚΙΝΗΘΕΙ ΚΑΘΟΛΟΥ ΝΑ ΜΠΟΡΕΙ ΝΑ ΔΕΙΞΕΙ
                If EID.Recordset.EOF Then
                  EID.RecordSource = "SELECT KOD AS [ΚΩΔ]," _
                  & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                  & " from EID " _
                  & " WHERE  KODERG = '" + Trim(mID(TIMText3.Text, 2, 14)) + "'  ;"
                 End If
                 
                 
                 
                 
                 
                 
                 
               
' & " WHERE EGGTIM.APOT=" + left(APOTH1.Text, 1) + " AND ERG  LIKE '" + SameLetters(Trim(Mid(TIMText3.Text, 2, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) ;"
               
               
               
                
            '----------------kata kodiko otan einai gyrismeno sta barcode ---------------------------------
            ElseIf f_barcode = 1 And Left(TIMText3.Text, 1) = "/" Then ' $$$$$$  KATA BARCODE
                EID.RecordSource = "select SUM(CASE WHEN APOT<>" + Left(APOTH1.Text, 1) + " OR  XRE IS NULL THEN 0 ELSE XRE END)" _
                 & "-SUM(CASE WHEN  APOT<>" + Left(APOTH1.Text, 1) + " OR  PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," _
                 & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                 & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" _
                 & " WHERE EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND   KOD  LIKE '" + SameLetters(Trim(mID(TIMText3.Text, 2, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1;"
            
                EID.Refresh
                
                  If Left(f_pelMERC, 3) = "SOK" And (Not EID.Recordset.EOF) Then
                       'Dim RSOK As New ADODB.Recordset

                       'Dim RSOK As New ADODB.Recordset
                       'Dim RSOK2 As New ADODB.Recordset
                       
                       NNN = 0
                       RSOK.Open "select MEMO from EID WHERE KOD='" + EID.Recordset("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic
                      ' Text14.Text = RSOK("MEMO")
                                              'Text14.Text = RSOK(0)
                                             DD2 = CNull(RSOK(0))
                                              RSOK.Close
                       
                       
                       RSOK.Open "select BARCODES.*,(SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=30 AND AYJON= CAST(SUBSTRING(BARCODES.PROELEYSH,7,2) AS INT ) ) AS XX from BARCODES  WHERE BARCODES.KOD='" + EID.Recordset("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic
                       ' NNN = RSOK.RecordCount
                       
                       
                       
                       MSFlexGrid1.row = 0
                       'DD2 = Text14.Text
                       
                       DUM = FETES_DELIM(DD2, ELEM)
                       
                       'If NNN > 0 Then
                       
                     '  DUM = FETES_DELIM(DD2, ELEM)
                       
                       Do While Not RSOK.EOF
                          MSFlexGrid1.row = MSFlexGrid1.row + 1
                          
                          MSFlexGrid1.Col = 0
                          MSFlexGrid1.Text = CNull(RSOK("XX"))
                         ' RSOK.MoveNext
                          
                          MSFlexGrid1.Col = 1
                          MSFlexGrid1.Text = RSOK("ERG")
                          
                          MSFlexGrid1.Col = 2
                          'MSFlexGrid1.Text = RSOK("PROELEYSH")
                         
                          
                          For NNN = 1 To 30
                              If InStr(ELEM(NNN), Left(CNull(RSOK("XX")), 3)) > 0 Then
                                   MSFlexGrid1.Col = 2
                                   
                                   If mID(ELEM(NNN), 6, 1) = "/" Then
                                       MSFlexGrid1.Text = mID(ELEM(NNN), 7, 3)
                                   Else
                                       MSFlexGrid1.Text = mID(ELEM(NNN), 6, 3)
                                   End If
                                   
                                   Exit For
                               End If
                          Next
                          
                          
                          
                          RSOK.MoveNext
                       Loop
                       RSOK.Close
                    'End If
                  
                  End If
                
                
                
                'ΑΝ ΔΕΝ ΕΧΕΙ ΚΙΝΗΘΕΙ ΚΑΘΟΛΟΥ ΝΑ ΜΠΟΡΕΙ ΝΑ ΔΕΙΞΕΙ
                If EID.Recordset.EOF Then
                  EID.RecordSource = "SELECT KOD AS [ΚΩΔ]," _
                   & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                   & " from EID " _
                   & " WHERE  KOD  LIKE '" + SameLetters(Trim(mID(TIMText3.Text, 2, 14))) + "%' ;"
                 End If
            
            ElseIf f_barcode = 1 And Left(TIMText3.Text, 1) = "%" Then ' $$$$$$  KATA BARCODE
            
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
            
            '----------------kata  barcode ---------------------------------
            ElseIf f_barcode = 1 Then ' $$$$$$  KATA BARCODE
               
               If F_EXO_XROMATA = 0 Then
                  Dim MPOSOTHTA2 As String
                  MPOSOTHTA2 = "EID.POS"
                  If Val(Left(APOTH1.Text, 1)) = 1 Then MPOSOTHTA2 = "EID.POS01"
                  If Val(Left(APOTH1.Text, 1)) = 2 Then MPOSOTHTA2 = "EID.POS02"
                  If Val(Left(APOTH1.Text, 1)) = 3 Then MPOSOTHTA2 = "EID.POS03"
                  If Val(Left(APOTH1.Text, 1)) = 4 Then MPOSOTHTA2 = "EID.POS04"
                  
                  If Len(Trim(F_2CHFIA_ZYGIZOMENA)) > 1 And Left(F_2CHFIA_ZYGIZOMENA, 2) = Left(TIMText3.Text, 2) Then
                     EID.RecordSource = "select " + MPOSOTHTA2 + " AS [ΥΠΟΛ],EID.KOD,EID.ONO,XTI,LTI,LTI5,FPA,MON,MEMO,POS_KERD,EID.ERG,EID.SYSKMAX from BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD WHERE BARCODES.ERG LIKE '" + SameLetters(Trim(mID(TIMText3.Text, 3, 5))) + "%';"
                  Else
                                   
                     EID.RecordSource = "select " + MPOSOTHTA2 + " AS [ΥΠΟΛ],EID.KOD,EID.ONO,XTI,LTI,LTI5,FPA,MON,MEMO,POS_KERD,EID.ERG,EID.SYSKMAX from BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD WHERE BARCODES.ERG LIKE '" + SameLetters(Trim(mID(TIMText3.Text, 1, 13))) + "%';"
                  End If
                  
                  EID.Refresh
                  'ψαχνω κατα βοηθητικο αν δεν βρεθηκε
                  If EID.Recordset.EOF Then
                         ' ψαχνει κατά BOHUHTIKO κωδικό  ERG %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                         EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)" _
                         & "-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," _
                         & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                         & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" _
                         & " WHERE  EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND ERG LIKE '" + SameLetters(Trim(mID(TIMText3.Text, 1, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1 "
                         EID.Refresh
                         If F_MONOMEYPOLOIPO = 1 And f_pol = 1 Then
                             EID.RecordSource = EID.RecordSource + " HAVING [ΥΠΟΛ]>0"
                         End If
                   End If
               Else 'XROMATA=1
               On Error GoTo 0
               
                   EID.RecordSource = "select top 10 EID.KOD,EID.ONO,LTI5,FPA,MON,XTI,LTI,MEMO,POS_KERD,EID.ERG,EID.SYSKMAX,PROELEYSH,KATHGORIA,CH1 from BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD WHERE BARCODES.ERG LIKE '" + SameLetters(Trim(mID(TIMText3.Text, 1, 13))) + "%';"
                  
                        ' ψαχνει κατά BOHUHTIKO κωδικό  ERG %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                  '       EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)" _
                   '      & "-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," _
                    '     & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                     '    & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" _
                      '   & " WHERE ERG  LIKE '" + SameLetters(Trim(mID(TIMText3.Text, 1, 14))) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1 "
                       '  EID.Refresh
                   
                  EID.Refresh
                 If Not EID.Recordset.EOF Then
                     XROMATA(Grid1.row, 1, 0) = EID.Recordset("PROELEYSH")
                     XROMATA(Grid1.row, 0, 1) = Replace(mID(EID.Recordset("PROELEYSH"), 4, 2), "*", "")
                 End If
                 
                  
                  
                  
                  If Left(f_pelMERC, 3) = "SOK" And (Not EID.Recordset.EOF) Then
                       NNN = 0
'                       RSOK.Open "select MEMO from EID WHERE KOD='" + EID.Recordset("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic
                        DD2 = CNull(EID.Recordset("MEMO"))
                        
'                                              RSOK.Close
                       RSOK.Open "select BARCODES.*,(SELECT PERIGRAFH FROM PINAKES WHERE TYPOS=30 AND AYJON= CAST(SUBSTRING(BARCODES.PROELEYSH,7,2) AS INT ) ) AS XX from BARCODES  WHERE BARCODES.KOD='" + EID.Recordset("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic
                       MSFlexGrid1.row = 0
                       DUM = FETES_DELIM(DD2, ELEM)
                       Do While Not RSOK.EOF
                          MSFlexGrid1.row = MSFlexGrid1.row + 1
                          
                          MSFlexGrid1.Col = 0
                          MSFlexGrid1.Text = CNull(RSOK("XX"))
                         ' RSOK.MoveNext
                          
                          MSFlexGrid1.Col = 1
                          MSFlexGrid1.Text = RSOK("ERG")
                          
                          MSFlexGrid1.Col = 2
                          For NNN = 1 To 30
                              If InStr(ELEM(NNN), Left(CNull(RSOK("XX")), 3)) > 0 Then
                                   MSFlexGrid1.Col = 2
                                   
                                   If mID(ELEM(NNN), 6, 1) = "/" Then
                                       MSFlexGrid1.Text = mID(ELEM(NNN), 7, 3)
                                   Else
                                       MSFlexGrid1.Text = mID(ELEM(NNN), 6, 3)
                                   End If
                                   
                                   Exit For
                               End If
                          Next
                          RSOK.MoveNext
                       Loop
                       RSOK.Close
                  End If
                 
                 
               End If
               
            '----------------kata  barcode  οταν είναι γυρισμένο κατα κωδικό ---------------------------------
            ElseIf Left(TIMText3.Text, 1) = "/" Then ' $$$$$$  KATA BARCODE
               EID.RecordSource = "select EID.KOD,EID.ONO,XTI,LTI,LTI5,FPA,MON,MEMO,POS_KERD,EID.ERG,EID.SYSKMAX,EID.KATHGORIA,CH1 from BARCODES INNER JOIN EID ON BARCODES.KOD=EID.KOD WHERE BARCODES.ERG LIKE '" + SameLetters(Trim(mID(TIMText3.Text, 2, 14))) + "%';"
            
            '-------------------ΓΕΝΝΗΤΡΙΑ ΚΩΔΙΚΩΝ -------------------------------------------------
            ElseIf Left(TIMText3.Text, 1) = "%" Then   '$$$$$$$
               If fbuff = "r" Then '------------------------
                  Dim GEN2 As String
                  
                  If Len(Dir("C:\DIONISIS.EXE")) > 0 Then
                       GEN2 = "5201001%"
                       R0.Open "select max(ERG) FROM EID WHERE ERG LIKE '" + GEN2 + "'", Gdb, adOpenForwardOnly, adLockReadOnly
                       XX = Format(1 + Val(Left(R0(0), 12)), "000000000000")
                       R0.Close
                    APOT1.flag.Caption = "1"
                       APOT1.Text1(4).Text = XX + ChDigEAN13(XX) 'timtext3.Text
                                         
                  Else
                      GEN2 = "5221275%"
                      R0.Open "select max(KOD) FROM EID WHERE KOD LIKE '" + GEN2 + "'", Gdb, adOpenForwardOnly, adLockReadOnly
                      XX = Format(1 + Val(Left(R0(0), 12)), "000000000000")
                      R0.Close
                      APOT1.flag.Caption = "1"
                      APOT1.Text1(0).Text = XX + ChDigEAN13(XX) 'timtext3.Text
                  End If
                  
                  
                  
                  gApoMenu = False
                  APOT1.SHOW 1
                  TIMText3.SetFocus
                  EID.RecordSource = "select KOD,ONO,XTI,LTI,LTI5,FPA,MON,MEMO,POS_KERD,ERG from EID WHERE KOD LIKE '" + XX + "%';"
               End If              '------------------------
               
20          Else   ' $$$$$$$
                 
                 '----KATA KΩΔIKO ΓΥΡΙΣΜΕΝΟ ΣΤΟ ΚΑΤΑ ΚΩΔΙΚΟ
                
                 'εαν ειναι ρυθμισμένο να δείχνει όλα τα είδη ή είναι αγορά (οπότε πρέπει να τα δείχνει όλα)
'                 If F_YPOL_GEN = 1 Or PEL.Recordset("EIDOS") = "r" Then
'                     If Val(Left(APOTH1.Text, 2)) = 1 Then
'                         C = "POS01"
'                     ElseIf Val(Left(APOTH1.Text, 2)) = 2 Then
'                          C = "POS02"
'                     ElseIf Val(Left(APOTH1.Text, 2)) = 3 Then
'                          C = "POS03"
'                     ElseIf Val(Left(APOTH1.Text, 2)) = 4 Then
'                          C = "POS04"
'                     Else
'                          C = "POS"
'                     End If
'
'                     EID.RecordSource = "select " + C + " AS [ΥΠΟΛ], KOD AS [ΚΩΔ]," _
'                     & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
'                     & " from EID " _
'                     & " WHERE  ONO  LIKE '" + SameLetters(TIMText3.Text) + "%' ;"
'                     EID.Refresh
'                 Else
                  If F_YPOL_GEN = 1 Or PEL.Recordset("EIDOS") = "r" Then
                       Synt5 = ""
                  Else
                       Synt5 = " EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND "
                  End If
                  
                  
                  
                     EID.RecordSource = "select SUM(CASE WHEN APOT<>" + Left(APOTH1.Text, 1) + " OR  XRE IS NULL THEN 0 ELSE XRE END)" _
                     & "-SUM(CASE WHEN  APOT<>" + Left(APOTH1.Text, 1) + " OR  PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ],KOD AS [ΚΩΔ]," _
                     & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                     & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" _
                     & " WHERE " + Synt5 + "  KOD  LIKE '" + SameLetters(TIMText3.Text) + "%' GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1 ;"
                     
                     EID.Refresh
                      ' End If
                 
                 
               If EID.Recordset.EOF Then
                  EID.RecordSource = "select KOD AS [ΚΩΔ]," _
                 & " EID.ONO,XTI,LTI,LTI5,EID.FPA,MON,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                 & " from EID " _
                 & " WHERE  KOD  LIKE '" + SameLetters(TIMText3.Text) + "%' ;"
               End If
             
             
             
             End If   ' $$$$$$$
             EID.Refresh
            
              
              GridEidon.Columns(0).width = 500
           GridEidon.Columns(1).width = 1800
           GridEidon.Columns(2).width = 5000
           
             GridEidon.SetFocus
             
             'ψάχνει κατά όνομα
         ElseIf Grid1.Col = f_o Then  '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            
             GridEidon.Columns(0).width = 500
             GridEidon.Columns(1).width = 5000
            'ME * NA MHN PSAXNEI KATA ONOMA
            If Left(Grid1.TextMatrix(Grid1.row, f_k), 1) = "*" Then ' ----------
                Grid1.Text = TIMText3.Text
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
               Dim MpOS
               'SYNDYASMENH ANAZHTHSH ME $  DHLADH   ΣΚΙΑ$99 ΟΛΕΣ ΤΙΣ "ΣΚΙΕΣ" ΜΕ ΚΩΔΙΚΟ ΠΟΥ ΑΡΧΙΖΕΙ ΑΠΟ "99"
               ONOM = TIMText3.Text
               DOL = InStr(1, ONOM, "$")
               KODI = ""
               If DOL > 1 Then
                  KODI = Trim(mID(ONOM, DOL + 1, 14))
                  ONOM = Left(ONOM, DOL - 1)
                  
                 EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)" _
                 & "-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS [ΥΠΟΛ]," _
                 & " EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                 & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" _
                 & " WHERE  EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND  EID.ONO  LIKE '" + SameLetters(ONOM) + "%' AND KOD LIKE '" + KODI + "%' and EGGTIM.APOT=" + Left(APOTH1.Text, 1) + "  GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1 ;"
               Else
                 
                 'εαν ειναι ρυθμισμένο να δείχνει όλα τα είδη ή είναι αγορά (οπότε πρέπει να τα δείχνει όλα)
'                 If F_YPOL_GEN = 1 Or PEL.Recordset("EIDOS") = "r" Then
'                     If Val(Left(APOTH1.Text, 2)) = 1 Then
'                         MpOS = "POS01"
'                     ElseIf Val(Left(APOTH1.Text, 2)) = 2 Then
'                          MpOS = "POS02"
'                     ElseIf Val(Left(APOTH1.Text, 2)) = 3 Then
'                          MpOS = "POS03"
'                     ElseIf Val(Left(APOTH1.Text, 2)) = 4 Then
'                          MpOS = "POS04"
'                     Else
'                          MpOS = "POS"
'                     End If
'
'                     EID.RecordSource = "select " + MpOS + " AS [ΥΠΟΛ], KOD AS [ΚΩΔ]," _
'                     & " EID.ONO,LTI5,EID.FPA,MON,XTI,LTI,KOD,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
'                     & " from EID " _
'                     & " WHERE  ONO  LIKE '" + SameLetters(TIMText3.Text) + "%' ;"
'                 Else
                 
                 
                 
                    'ΠΡΟΣΟΧΗ ΑΝ ΕΧΩ
                    ' WHERE  EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND  EID.ONO  LIKE '" + SameLetters(ONOM) + "%'"
                    'ΒΓΑΖΕΙ ΜΟΝΟ ΤΑ ΕΙΔΗ ΠΟΥ ΕΧΟΥΝ ΚΙΝΗΘΕΙ ΣΤΗΝ ΑΠΟΘΗΚΗ 1 Π.Χ.
                     'ΕΝΩ ΤΟ
                     ' WHERE EID.ONO  LIKE '" + SameLetters(ONOM) + "%'"
                     ' ΒΓΑΖΕΙ ΟΛΑ ΤΑ ΕΙΔΗ,ΚΑΙ ΑΥΤΑ ΠΟΥ ΔΕΝ ΕΧΟΥΝ ΚΙΝΗΘΕΙ
                  If F_YPOL_GEN = 1 Or PEL.Recordset("EIDOS") = "r" Then
                       Synt5 = ""
                  Else
                       Synt5 = " EGGTIM.APOT=" + Left(APOTH1.Text, 1) + " AND "
                  End If
                     
                     
                     
                     
                     EID.RecordSource = "select SUM(CASE WHEN XRE IS NULL OR APOT<>" + Left(APOTH1.Text, 1) + " THEN 0 ELSE XRE END)" _
                     & "-SUM(CASE WHEN PIS IS NULL OR APOT<>" + Left(APOTH1.Text, 1) + " THEN 0 ELSE PIS END) AS [ΥΠΟΛ]," _
                     & " EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX,KATHGORIA,CH1 " _
                     & " from EID LEFT JOIN EGGTIM  ON EID.KOD=EGGTIM.KODE" _
                     & " WHERE " + Synt5 + " EID.ONO  LIKE '" + SameLetters(ONOM) + "%'" + "  GROUP BY ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14),SYSKMAX,KATHGORIA,CH1 "
                     ' and EGGTIM.APOT=" + left(APOTH1.Text, 1)
                 
                 
                     If F_MONOMEYPOLOIPO = 1 And f_pol = 1 Then
                         EID.RecordSource = EID.RecordSource + " HAVING SUM(CASE WHEN XRE IS NULL OR APOT<>" + Left(APOTH1.Text, 1) + " THEN 0 ELSE XRE END)" _
                         & "-SUM(CASE WHEN PIS IS NULL OR APOT<>" + Left(APOTH1.Text, 1) + " THEN 0 ELSE PIS END)>0"
                      End If
                 'End If
                  
                
                
                End If
                
              EID.Refresh
              
    '   GridEidon.width = 10000
              GridEidon.Columns(0).width = 1000
             GridEidon.Columns(1).width = 5000
              GridEidon.Columns(3).width = 1000
             
             GridEidon.Columns(2).width = 1000
              

              
              
              
              
              
              
              
              
              
              
              If EID.Recordset.EOF Then
              
                 EID.RecordSource = "select EID.ONO,KOD,XTI,LTI,LTI5,EID.FPA,MON,POS_KERD,ERG,SUBSTRING(MEMO,1,14) AS MMEMO,SYSKMAX " _
                  & " from EID " _
                  & " WHERE   EID.ONO  LIKE '" + SameLetters(TIMText3.Text) + "%' ORDER BY ONO ;"
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
               APOT1.Text1(0).Text = TIMText3.Text
               APOT1.flag.Caption = "1"
               APOT1.SHOW 1
               TIMText3.SetFocus
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
          
       intRecord = min(10, EID.Recordset.RecordCount)
       intField = EID.Recordset.FIELDS.Count
       'call the procedure here...
    If intRecord > 1 Then
      Call AdjustDataGridColumns(intRecord, intField, True)
      EID.Recordset.MoveFirst
      GridEidon.SetFocus
       
    End If
    
'          GridEidon.SetFocus
      
   
     
     ' kod ή όνομα <2 τότε ψάχνουμε
     '******************************************************************************
30   Else ' ************************************************************************
         
               ' EKTOS KOD ,ONO
         
         
         
         
         'ΥΠΟΛΟΓΊΖΕΙ ΑΞΊΑ
         If Grid1.Col = f_p Then  ' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            Dim R As New ADODB.Recordset, mXre As Single, mPIS As Single, mDesmia As Single
            
            
            If f_use_Sysk = 1 Then
               If EID.Recordset("SYSKMAX") > 0 Then
                  TIMText3.Text = TIMText3.Text * EID.Recordset("SYSKMAX")
               End If
            End If
            
            If (fbuff = "e" And (F_POS_APOU = 2 Or F_POS_APOU = 3) And f_checkYpol > 0 And F_STADIO <> 2) Or (F_PAR_MONO_ME_YPOLOIPO = 1 And F_STADIO = 2) Then ' F_STADIO<>2 ΔΕΝ ΕΙΝΑΙ ΠΑΡΑΓΓΕΛΙΑ
               ' 13-4-2006 R.Open "SELECT SUM(XRE),SUM(PIS) FROM EGGTIM WHERE KODE='" + EID.Recordset("KOD") + "' ", Gdb, adOpenForwardOnly, adLockReadOnly
               R.Open "SELECT SUM(XRE),SUM(PIS) FROM EGGTIM WHERE  APOT=" + Left(APOTH1.Text, 1) + " and KODE='" + Grid1.TextMatrix(Grid1.row, 1) + "' ", Gdb, adOpenForwardOnly, adLockReadOnly
               
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
                        TIMText3.Text = 0
                   End If
                   
                   If HDH_FORTOMENA + Val(TIMText3.Text) > R(0) Then
                      MHNYMA2.Timer1.Interval = 50
                      MHNYMA2.SHOW 1
                   
                      MsgBox "Η παραγγελθείσα ποσότητα είναι " + Str(R(0)) + Chr(13) + "και έχετε ήδη τιμολογήσει " + Format(HDH_FORTOMENA + Val(TIMText3.Text), "#####.00")
                      TIMText3.Text = 0
                   End If
                   
                   
               End If
               
               
               
               
               
               If Val(TIMText3.Text) > (mXre - mPIS) - mDesmia Then
                    MHNYMA2.Timer1.Interval = 50
                      MHNYMA2.SHOW 1
                  
                  MsgBox "υπάρχουσα ποσότητα " + Format(mXre - mPIS, "###0") + " ΔΕΣΜ:" + Format(mDesmia, "####")
                  
                  If f_checkYpol = 2 Then
                     TIMText3.Text = 0
                  End If
               End If
            End If
            If F_EXO_XROMATA = 1 Then ' ΒΑΖΩ ΤΗΝ ΠΟΣΟΤΗΤΑ ΣΤΟΝ ΠΙΝΑΚΑ ΓΙΑ ΝΑ ΤΟΝ ΤΡΑΒΗΞΕΙ ΣΤΟ ENHM_EGGTIM
                Dim L9 As Long, M9 As Long
                Dim SUMA9 As Single
                SUMA9 = 0
                For L9 = 1 To 25
                   For M9 = 1 To 25
                      SUMA9 = SUMA9 + Val(XROMATA(Grid1.row, L9, M9))
                   Next
                Next
                If SUMA9 = 0 Then
                   XROMATA(Grid1.row, 1, 1) = TIMText3.Text
                End If
            End If
         End If     ' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        mText = TIMText3.Text
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
              
              
              If f_SYMPYKNVSH = 1 Then  '  And Left(timText3.Text, 2) <> Left(F_2CHFIA_ZYGIZOMENA, 2) Then
               
                 
                 
                 For LK0 = 1 To Grid1.row - 1
                    'brika idion kodiko
                    If Grid1.TextMatrix(LK0, 1) = Grid1.TextMatrix(Grid1.row, 1) Then
                          Grid1.TextMatrix(LK0, f_p) = Val(Grid1.TextMatrix(LK0, f_p)) + Val(Grid1.TextMatrix(Grid1.row, f_p))
                          Grid1.TextMatrix(LK0, f_a) = Round(Val(Grid1.TextMatrix(LK0, f_p)) * Val(Grid1.TextMatrix(LK0, F_T)) * (100 - Val(Grid1.TextMatrix(LK0, f_e))) / 100, f_psifiaAjias)
                          ' Round(Val(P) * Val(t) * (100 - Val(e)) / 100 * (100 + g_Fpa(Val(Grid1.TextMatrix(R, f_f)))) / 100, f_psifiaAjias)
                             
                          If LK0 <> Grid1.row Then
                              Grid1.TextMatrix(Grid1.row, f_a) = 0
                              Grid1.TextMatrix(Grid1.row, f_p) = 0
                              Grid1.TextMatrix(Grid1.row, F_T) = 0
                          End If
                          
                          
                          HDH_FORTOMENA = 1
                    End If
                  Next
                              
                              
                              
                              
                             ' Grid1.TextMatrix(Grid1.row, f_a) = 0
                             ' Grid1.TextMatrix(Grid1.row, f_p) = 0
                             ' Grid1.TextMatrix(Grid1.row, F_T) = 0
                              
                              
                              
                              Grid1.Refresh
                              
                              
                  FindSynolo
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
           TIMText3.Text = Grid1.TextMatrix(Grid1.row, F_T)
        End If
        If Grid1.Col = f_e Then
           TIMText3.Text = Grid1.TextMatrix(Grid1.row, f_e)
        End If
     End If
 End If
 
50 If KeyCode = 38 Then   ' ΠΑΝΩ ΒΕΛΟΣ
      
      Grid1.Text = TIMText3.Text ' κρατάει το grid το κείμενο
      
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
     SYNTF = (100 + g_Fpa(Grid1.TextMatrix(Grid1.row, f_f))) / 100
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
    End If
 
 End If
   
70
   If KeyCode = 40 Then   ' KATΩ ΒΕΛΟΣ
      Grid1.Text = TIMText3.Text ' κρατάει το grid το κείμενο
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
'If Err.Description = "Automation error" Then
'   Gdb.Close
'   Gdb.Open gConnect

'End If





Resume Next


RECONNECT:




Resume Next

End Sub





Private Static Sub FindSynolo()
   Dim C, R, P, t, e, k, S, aj, F, SUMPOS
On Error GoTo MHNYMA

            For k = 1 To 9: SumFpa(k) = 0: Next
            
            R = Grid1.row
            C = Grid1.Col  ' απομνημονεύω την τρέχουσα στήλη kai σειρά
            
10          Grid1.Col = f_p: P = Grid1.Text
            Grid1.Col = F_T: t = Grid1.Text
            Grid1.Col = f_e: e = Grid1.Text
            Grid1.Col = f_f: F = Grid1.Text
            
            
            
            'αξια με ΦΠΑ
            Grid1.Col = F_mefpa: aj = Round(Val(P) * Val(t) * (100 - Val(e)) / 100 * (100 + g_Fpa(Val(Grid1.TextMatrix(R, f_f)))) / 100, f_psifiaAjias)
            Grid1.Text = Replace(Format(aj, "####0.000"), ",", ".")
            
            
            
            Grid1.Col = f_a: aj = Round(Val(P) * Val(t) * (100 - Val(e)) / 100, f_psifiaAjias)
            Grid1.Text = Replace(Format(aj, "####0.000"), ",", ".")
            
            
            
            f_sajia(Grid1.row) = aj
            f_fpa(Grid1.row) = F  ' ΚΑΤΗΓΟΡΊΑ ΦΠΑ
20            S = 0
              SUMPOS = 0
            For k = 1 To Grid1.ROWS - 1
              If Len(Grid1.TextMatrix(k, f_k)) > 0 Then
                f_sajia(k) = Val(Grid1.TextMatrix(k, f_a)) '  Val(Grid1.TextMatrix(K, f_p)) * Val(Grid1.TextMatrix(K, F_T)) '
                
                SUMPOS = SUMPOS + Val(Grid1.TextMatrix(k, f_p))
                S = S + f_sajia(k)
                ' ΣΟΥΜ’ΡΕΙ ΣΕ Κ’ΘΕ ΚΑΤΗΓΟΡΊΑ φπα ΤΗΝ ΑΞΙΑ
30                If f_fpa(k) < 9 Then
                      SumFpa(f_fpa(k)) = SumFpa(f_fpa(k)) + f_sajia(k)
                  End If
              End If
                  
            Next
            
            Label15.Caption = SUMPOS
            
            S = Round(S, f_psifiaAjias)
            Synolo.Caption = Format(S, "####0.00")
            F = 0: For k = 1 To 9: F = F + SumFpa(k) * g_Fpa(k) / 100: Next
            
            F = Round(F, f_psifiaAjias)
40            sfpa.Caption = Format(F, "####0.00")
            SynoloG.Caption = Format(F + S, "####0.00")
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
Dim kodiklog(9), TYP(9), xp(9), kau(9), a As String, mnq As String, eid_par As String
' Dim mgre As String
Dim k As Integer

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
Dim lo As Integer
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
  For k = 1 To 9
      a = LTrim(Str(k))
      xp(k) = gefyres("xpi" + a)
      kodiklog(k) = gefyres("Vkod" + a)  ' 3000****
      TYP(k) = gefyres("Vtypos" + a)
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
fSCR.AddObject "mactext", MACtEXT









Dim debug1
' εδώ παίρνει το vkod1,vtypos1 έως το vkod9,typos9
 For k = 1 To 9
       If kodiklog(k) <> " " Then
            '            sele TEMP
            
            
            debug1 = DOK.Recordset("aji")
            a = TYP(k) '
            'dum = ExecuteLine("gvar=" + a)
            a = mac(a)
                    
            
          If Val(a) <> 0 Then    ' γγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγγ
'
'             ' EAN EIMASTE STHN ANALYSH EIDVN
             If mID$(kodiklog(k), 3, 2) = "**" Then ' κωδικός λογιστικής
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
             
             
             
             gvar = Trim(kodiklog(k))
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
               temp.Recordset("XPI") = xp(k)
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
                      If mmhkos = 2 Then lo = 1
                      If mmhkos = 4 Then lo = 2
                      If mmhkos = gDIG_3 Then lo = 3
                      If mmhkos = gDIG_4 Then lo = 4
                      abat2.Refresh
                      If abat2.Recordset.RecordCount < lo Then
                          For K2 = 1 To lo: logar(K2) = 0: Next
                          On Error Resume Next
                          abat2.Recordset.MoveFirst
                          Do While Not abat2.Recordset.EOF
                            logar(Val(abat2.Recordset("bat"))) = 1
                            abat2.Recordset.MoveNext
                            DoEvents
                          Loop
                          For K2 = 1 To lo
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
Dim Rtim As New ADODB.Recordset
Dim R As Recordset
Dim k As Integer
Dim KN As Integer
Dim R3 As New ADODB.Recordset
Dim mTimh, mPoso
Dim PROHG





'-----------  next lines are OK --------------------------

On Error GoTo MHNYMA

printCrystal = 1 ' αν ειναι dum < -900 τοτε δεν τυπωνει



'ΑΝ ΜΕΝ ΕΙΝΑΙ ΔΙΟΡΘΩΣΗ ΠΡΕΠΕΙ ΝΑ ΚΡΑΤΑ ΤΗΝ ΠΑΛΙΑ ΗΜΕΡΟΜΗΝΙΑ
'ΑΝ ΔΕ ΕΙΝΑΙ ΑΚΥΡΩΤΙΚΟ ΠΡΕΠΕΙ ΝΑ ΕΧΕΙ ΤΗΝ ΣΗΜΕΡΙΝΗ ΗΜΕΡΟΜΗΝΙΑ
'TA EKANA REM STIS 26/5/2009
'If Len(f_dior_mAtim) > 0 Then
'   d1 = Now
'   d2 = DateAdd("d", 1, d1)
'Else
'   d1 = DTPicker1.value
'   d2 = DateAdd("d", 1, d1)
'End If



   d1 = char_date
   d2 = DateAdd("d", 1, d1)



Set db = OpenDatabase("c:\mercvb\reports\reports.mdb")

 ADD_FIELDDAO "TIM", "AJ6", "SINGLE", db
 ADD_FIELDDAO "TIM", "FPA6", "SINGLE", db
 ADD_FIELDDAO "TIM", "DEH", "CHAR(12)", db
 ADD_FIELDDAO "TIM", "C1", "CHAR(30)", db
 ADD_FIELDDAO "TIM", "C2", "CHAR(30)", db
 ADD_FIELDDAO "EGGTIM", "PROELEYSH", "CHAR(10)", db
 ADD_FIELDDAO "EGGTIM", "MIKTA", "SINGLE", db

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
sql = sql + "MONA,EKPT,PROELEYSH,(CASE WHEN MIKTA IS NULL THEN 0 ELSE MIKTA END) AS MIKTA "

sql = sql + " from EGGTIM where APOT=" + Str(Val(Left(APOTH1.Text, 2))) + " and  left(ATIM," + Str(f_psifia_atim + 1) + ")='" + Left(MATIM, f_psifia_atim + 1) + "' and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "' ORDER BY PROOD_AJ;"


'Gdb.Execute SQL
'GoTo 10
'On Error GoTo 0

'On Error GoTo 0
REGGTIM.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly

'copy to access table eggtim
Set R = db.OpenRecordset("eggtim")
'On Error Resume Next

'REGGTIM.Close
'REGGTIM.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly




REGGTIM.MoveFirst

On Error Resume Next
'On Error GoTo 0
Do While Not REGGTIM.EOF
  
   R.AddNew
'On Error GoTo 0
   
   
   For k = 0 To REGGTIM.FIELDS.Count - 1
     
     On Error GoTo LATOS
     
     
     If REGGTIM(k).Type = 202 Or REGGTIM(k).Type = 200 Then ' STRING
        'R(REGGTIM(k).Name) = Left(REGGTIM(k), R(REGGTIM(k).Name).Size)
        R(k) = Left(REGGTIM(k), R(REGGTIM(k).Name).Size)
     Else
         R(k) = REGGTIM(k)
        'R(REGGTIM(k).Name) = REGGTIM(k)
     End If
     
   Next
   
   mTimh = REGGTIM("Tιμ_Μον")
   mPoso = REGGTIM("Ποσότητα")
   
'   If F_XONDR = 0 Then 'LIANIKH
'   Else
'        r("Αξία") = r("Ποσότητα") * r("Tιμ_Μον") * (100 - r("EKPT")) / 100
'   End If
   
   
   R.Update
   
   REGGTIM.MoveNext
Loop

Dim sum_mikta As Long, sum_posothta As Single

sum_mikta = 0
sum_posothta = 0

REGGTIM.Close


Dim SAJIA2 As Single   ' Η ΑΞΙΑ ΥΠΟΛΟΓΙΖΟΜΕΝΗ




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
     
    
     
     
     If R("fpa") = 0 Then
          R("FPA") = 2
          MsgBox "Προσοχή βρέθηκε ΦΠΑ=0 και μετετράπη σε κατηγ. 2"
     Else
          R("fpa") = g_Fpa(R("fpa"))
     End If
     
     
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

KN = f_SynSeiresTimologioy - R.RecordCount  'f_SeiresTimologioy - R.RecordCount
For k = 1 To KN
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

'
' EINAI OK ALLA TO EBALA STO ENHM_PEL
'Dim rTemp As New ADODB.Recordset
'Dim SQLt As String, TYP As Single
'SQLt = "SELECT EIDOS,KOD,"
'SQLt = SQLt + "SUM(CASE WHEN XREOSI IS NULL  THEN 0 ELSE XREOSI  END )  -"
' SQLt = SQLt + " SUM(CASE WHEN PISTOSI IS NULL  THEN 0 ELSE PISTOSI  END )   AS YP"
' SQLt = SQLt + " FROM EGG  where HME>='" + Format(gEnarjh, "MM/DD/YYYY") + "' AND EIDOS='" + PEL.Recordset("EIDOS") + "' AND KOD='" + PEL.Recordset("kod") + "' GROUP BY KOD"
'rTemp.Open SQLt, Gdb, adOpenDynamic, adLockOptimistic
'If IsNull(rTemp("YP")) Then
'   TYP = 0
'Else
'   TYP = rTemp("YP")
'End If
'Gdb.Execute "UPDATE PEL SET TYP=AYP+ " + Format(TYP, "###########0.00")








'========  TIM ==========
sql = "select AJI,AJ1,AJ2,AJ3,FPA1,FPA2,FPA3,PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.TYP,PEL.DOY,PEL.KOD,AJ4,AJ5,FPA4,PEL.THL,PEL.POL,AJ6,FPA6,ELGA,PEL.CH1 AS DEH,EKPT5 AS PROHGYPOL "
sql = sql + "  from TIM inner join PEL on TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD where " + synt + " and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2, "mm/dd/yyyy") + "';"
  db.Execute "delete *from tim"
Rtim.Open sql, Gdb, adOpenForwardOnly, adLockReadOnly


PROHG = Rtim("prohgYPOL")

'copy to access table tim
Set R = db.OpenRecordset("tim")
On Error GoTo LATOS


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



Rtim.MoveFirst
   
' select ATIM,
' HME as shme,AJI as ajia,PEL.TYP,
   
   
   
'   On Error GoTo 0
   R.AddNew
   
On Error Resume Next

'Dim ISPIST As Boolean


   
         R("SHME") = Rtim("SHME")
         R("TYP") = Rtim("TYP")
   
   
         R("EPO") = Left(CNull(Rtim("EPO")), Rtim("epo").ActualSize)
         R("EPA") = Left(CNull(Rtim("EPA")), Rtim("epa").ActualSize)
         R("DIE") = Left(CNull(Rtim("DIE")), Rtim("die").ActualSize)
         R("AFM") = Left(CNull(Rtim("AFM")), Rtim("afm").ActualSize)
         
         R("DOY") = Left(CNull(Rtim("DOY")), Rtim("doy").ActualSize)
         R("KOD") = Left(CNull(Rtim("KOD")), Rtim("kod").ActualSize)
         R("THL") = Left(CNull(Rtim("THL")), Rtim("thl").ActualSize)
         R("POL") = Left(CNull(Rtim("POL")), Rtim("pol").ActualSize)
   
   
         R("AJIA") = Rtim("AJIA")
         R("AJI") = Rtim("AJI")
         R("AJ1") = Rtim("AJ1")
         R("AJ2") = Rtim("AJ2") '
         R("AJ3") = Rtim("AJ3")
         R("AJ4") = Rtim("AJ4")
         R("AJ5") = Rtim("AJ5")
         R("AJ6") = Rtim("AJ6")
         
         R("FPA1") = Rtim("FPA1")
         R("FPA2") = Rtim("FPA2") '
         R("FPA3") = Rtim("FPA3")
         R("FPA4") = Rtim("FPA4")
         R("FPA6") = Rtim("FPA6")
         R("ELGA") = Rtim("ELGA")
         
         R("ATIM") = Rtim("ATIM")
         
         
         R("DEH") = Left(CNull(Rtim("DEH")), 12)
         
'         'BLEPEI AN EINAI PISTOTIKO
'         If Left(r("ATIM"), 1) = "P" Then
'            ISPIST = True
'         Else
'            ISPIST = False
'         End If
         
         
On Error GoTo LATOS
   R.Update
On Error Resume Next
   
Rtim.Close

R.Close

Set R = db.OpenRecordset("tim")

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


'=============  KENES SEIRES NA GEMISEI EGGTIM  =========


'==========

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
     If Abs(SAJIA - (R("AJ1") + R("AJ2") + R("AJ3") + R("AJ4") + R("AJ5") + R("AJ6"))) > 0.009 Then
       MsgBox "Aσυμφωνία παραστατικού. Διαγράψτε το παραστατικό και ξαναπεράστε το"
       printCrystal = -1000
       ' End
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
fmem("n3") = mTimh
fmem("n4") = mPoso
fmem("BASTIM_GID") = F_METRHTAposo


fmem("MEMO") = parat.Text
'If Combo4.Text = Combo4.List(0) Then ' METRITOIS
   fmem("par4") = Format(PROHG, "#####.00")
'Else
 '  fmem("par4") = Format(PEL.Recordset("TYP") - Val(SynoloG.Caption), "#####.00")
'End If

'If ISPIST = True Then fmem("PAR4") = " "



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
Dim Rtim As New ADODB.Recordset
Dim R As New ADODB.Recordset
Dim k As Integer
Dim KN As Integer

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
KN = f_SeiresTimologioy - R(0)  'f_SeiresTimologioy - R.RecordCount
R.Close





'On Error GoTo 0

'R.Close
R.Open "SELECT * FROM EGGTIMREP", Gdb, adOpenDynamic, adLockOptimistic

'R.Open "SELECT * FROM EGGTIMREP", Gdb, adOpenForwardOnly, adLockReadOnly



' On Error GoTo 0
For k = 1 To KN
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
sql = "select AJI,AJ1,AJ2,AJ3,FPA1,FPA2,FPA3,PEL.EPO,ATIM,HME as shme,AJI as ajia,PEL.EPA,PEL.DIE,PEL.AFM,PEL.TYP,PEL.DOY,PEL.KOD,AJ4,AJ5,FPA4,PEL.THL,PEL.POL,TIM.EKPT5 AS PROHGYPOL "
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
'If Combo4.Text = Combo4.List(0) Then ' METRITOIS
   fmem("par4") = Format(R("PROHGYPOL"), "#####.00")
'Else
'   fmem("par4") = Format(PEL.Recordset("TYP") - Val(SynoloG.Caption), "#####.00")
'End If
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
 
 
Dim SA, sB

Dim DUM
 Dim eField(1 To 50), ePict(1 To 50), eProw(1 To 50), ePcol(1 To 50), eFonts(1 To 50)
 Dim mField(1 To 50), mPict(1 To 50), mProw(1 To 50), mPcol(1 To 50), mFonts(1 To 50)
 Dim sField(1 To 50), sPict(1 To 50), sProw(1 To 50), sPcol(1 To 50), sFonts(1 To 50)
 Dim m_No_of_seir, I, k, M, C
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
Dim e, MM, S
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
fSCR.AddObject "mactext", MACtEXT

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
    'c=ltrim(str(select(),2))  '&& ‘¨«ι «€ §¨ ¦®γ ©«γ€ ¦§¦ε ¨ ©‘ζ£¦¬€

    'load_forma (arxeio)

'Dim DUM

If Len(Dir("c:\PETREL.TXT")) > 1 Then
      DUM = load_forma(arxeio, 0)
 Else
     DUM = load_forma(arxeio, 1)
End If
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


Dim mfil As String

'για να σωζει σε διαφορετικα σημεια
If f_YPOK = 0 Then
    mfil = F_File_Timologioy
Else
    mfil = F_File2_Timologioy
End If


10 Open mfil For Output As #1
    
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
100      M = 0
110    Do While Len(Trim(gm_f(e + M + 2))) > 0
120         M = M + 1

         'MM = Trim(gm_f(M + e + 1))

         'm_m = ExecuteLine("gvar=" + MM)

         'mField(M) = gvar
130         mPict(M) = Trim(gpic(M + e + 1))
140         mProw(M) = gm_r(M + e + 1)
150         mPcol(M) = gm_c(M + e + 1)
         'mFonts(M) = "" '     && trim( Fonts )
         'MM = mPict(M + e + 1)
160      Loop


Dim fcheckonly As Boolean


      '****************  ΣΥΝΟΛΑ  ****************************************
      S = 0
170      Do While Len(Trim(gm_f(e + M + S + 3))) > 0
         S = S + 1
'         MM = Trim(gm_f(M + e + s + 2))
'         MM = "gvar=" + MM
'         m_m = ExecuteLine("gvar=" + MM)
'
'         sField(s) = gvar
180         sPict(S) = Trim(gpic(M + e + S + 2))
190         sProw(S) = gm_r(M + e + S + 2)
200         sPcol(S) = gm_c(M + e + S + 2)
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
208       For k = 1 To e
             
             MM = ePict(k)
             'For i = 1 To eProw(K): Printer.Print: Next
             For I = 1 To eProw(k)
                 If k = 1 And m_No_of_seir = 1000 Then
                    ' ΜΗΝ ΒΑΖΕΙΣ ΚΕΝΗ ΣΕΙΡΑ ΣΤΗΝ ΑΡΧΗ
                 Else
                     Print #1, " "
                 End If
                 
             Next
             syn_row = syn_row + eProw(k) 'συγκεντρώνω τις σειρές
             
220           MM = mac(Trim(Trim(gm_f(k)))) ': m_m = ExecuteLine("gvar=" + MM)
           'Printer.Print Tab(ePcol(K)); Format(gvar, ePict(K))
           
              If eProw(k + 1) > eProw(k) Then 'αν πρόκειται να τυπώσει παρακάτω μην τυπώνεις δίπλα μου
                '   Print #1, Tab(ePcol(K)); Format(mM, ePict(K))
                Print #1, Tab(ePcol(k)); Format(Left(MM, Len(ePict(k))), ePict(k));
              Else
                  Print #1, Tab(ePcol(k)); Format(Left(MM, Len(ePict(k))), ePict(k));
              End If
             
                m1 = eProw(k)
             
230        Next
      
      
      
    If m_seir_eid >= m_No_of_seir Then
        GoTo 400
    End If
        
      
      
      
      
240   For k = 1 To m_seir_eid - syn_row
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
         For k = 1 To M
           
280           MM = mac(Trim(Trim(gm_f(e + k + 1))))                 ': m_m = ExecuteLine("gvar=" + MM)
           '   Printer.Print Format(gvar, gpic(e + K + 1));
290           Print #1, Tab(mPcol(k)); GFORMAT(MM, gpic(e + k + 1));
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
  
330    For k = 1 To m_seir_synol - syn_row: Print #1,: Next  'βρισκω την σειρά  των synolon
       
       For k = 1 To S
340         MM = Trim(gm_f(M + e + k + 2))
350        ' m_m = ExecuteLine("gvar=" + MM)
         
                 MM = mac(Trim(gm_f(M + e + k + 2)))
         
         
360         For I = 1 To sProw(k): Print #1,: Next
         'Printer.Print Format(gvar, gpic(M + e + K + 2));
370         Print #1, Tab(sPcol(k)); GFORMAT(MM, gpic(M + e + k + 2));
       Next
       
'       If SELIDA > 1 Then
'                Print #1, "Σελίδα " + Format(SELIDA, "##") ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
'       End If
       
       
       
       
       
       
400 Close #1

380 cEGGTIM.Recordset.Close

Set fSCR = Nothing


If f_YPOK = 0 Then
 If F_Type_File_Timologioy >= 1 Then
    If F928 = 1 Then
      If F_Type_File_Timologioy = 2 Then
         MDIForm1.Caption = "PRINTER.PRINT"
         DUM = EKTYP_ME_PRINT(mfil)
      
      ElseIf F_Type_File_Timologioy = 3 Then
         MDIForm1.Caption = "TYPE"
         DUM = Shell("C:\MERCVB\TYP.BAT ", vbMinimizedFocus)
      
      ElseIf F_Type_File_Timologioy = 4 Then
         MDIForm1.Caption = "TYPE-EXE"
         DUM = Shell("C:\MERCVB\TYPE.EXE ", vbMinimizedFocus)
      
      ElseIf Len(Dir("c:\marietos.txt")) > 1 Then
         ' kanei "type c:\print"  meso dos
         DUM = Shell("c:\mercvb\notedos.exe /p " + mfil, vbMaximizedFocus)
       ElseIf Len(Dir("c:\AGIOS.txt")) > 1 Then
         ' kanei "type c:\print"  meso dos
         DUM = Shell("c:\mercvb\ff.exe", vbMaximizedFocus)
         
       ElseIf Len(Dir("c:\KAMPOUR.txt")) > 1 Then
         ' kanei "type c:\print"  meso dos
            DUM = EKTYP_ME_PRINT(mfil)
            ' DUM = Shell("c:\mercvb\ff.exe", vbMaximizedFocus)
       Else
            DUM = Shell("c:\mercvb\notepad.exe /p " + mfil, vbMaximizedFocus)
       End If
   Else  ' <>928
        
        If F_Type_File_Timologioy = 2 Then
           DUM = EKTYP_ME_PRINT(mfil)
        ElseIf F_Type_File_Timologioy = 3 Then
           MDIForm1.Caption = "TYPE"
           DUM = Shell("C:\MERCVB\TYP.BAT ", vbMinimizedFocus)
        
        ElseIf F_Type_File_Timologioy = 4 Then
           MDIForm1.Caption = "TYPE-EXE"
           DUM = Shell("C:\MERCVB\TYPE.EXE ", vbMinimizedFocus)
        
        
        
        ElseIf Len(Dir("c:\FAKKAV.TXT")) > 1 Then
           DUM = EKTYP_ME_PRINT(mfil)
        ElseIf Len(Dir("C:\OLYMPIAN.TXT")) > 1 Then
            DUM = Shell("C:\TYP.BAT ", vbMinimizedFocus)
        
        
        ElseIf Len(Dir("c:\PETREL.TXT")) > 1 Then
           Open "c:\PETREL.TXT" For Input As #1
                Input #1, SA
                sB = Val(SA)
           Close #1
           DUM = EKTYP_ME_serial(mfil, sB)
           
           ' DUM = Shell("C:\T.BAT ", vbMinimizedFocus)
        Else
           DUM = Shell("C:\MERCVB\PRINT.EXE " + Trim(mfil), vbMinimizedFocus)
        End If
   End If '928
   
   
 Else  'If F_Type_File_Timologioy >= 1 Then

      If Len(Dir("c:\DION2.EXE")) > 1 Then
        ' DUM = EKTYP_ME_PRINT(mfil)
         MILSEC 5000
         FileCopy "C:\KENO.TXT", "C:\FOROL\FOROL.TXT"
'      ElseIf Len(Dir("C:\SOKRATIS.TXT")) > 1 Then
'        ' DUM = EKTYP_ME_PRINT(mfil)
'
'         'FileCopy "C:\KENO.TXT", "C:\FOROL\FOROL.TXT"
'
'         Printer.Print "     "
'         Printer.Print "     "
'         Printer.Print "     "
'
'         Printer.Print "     "
'         Printer.Print "     "
'         Printer.Print "     "
'
'         Printer.Print "     "
'         Printer.Print "     "
'         Printer.Print "     "
'         Printer.EndDoc
         
      
      
      
      
      ElseIf Len(Dir("c:\KAMPOUR.txt")) > 1 Then
          'MILSEC 8000
          Do While Len(Dir("c:\FOROL\FOROL.txt")) > 1
            MILSEC 500
          Loop
         FileCopy "C:\KENO.TXT", "C:\FOROL\FOROL.TXT"
         ' kanei "type c:\print"  meso dos
'            DUM = EKTYP_ME_PRINT(mfil)

      ElseIf Len(Dir("c:\KENO1.txt")) > 1 Then
          'MILSEC 8000
          MILSEC 1000
          Do While Len(Dir(mfil)) > 1
            MILSEC 1000
          Loop
          FileCopy "C:\KENO1.TXT", mfil
          FileCopy "C:\KENO1.TXT", "C:\PP2"
      End If
 End If

Else
'-----------------------------------------------
 If F_Type2_File_Timologioy >= 1 Then
   If F_Type2_File_Timologioy = 2 Then
      DUM = EKTYP_ME_PRINT(mfil)
      
   ElseIf F_Type2_File_Timologioy = 3 Then
      MDIForm1.Caption = "TYPE"
      DUM = Shell("C:\MERCVB\TYP.BAT ", vbMinimizedFocus)
   ElseIf F_Type2_File_Timologioy = 4 Then
      MDIForm1.Caption = "TYPE"
      DUM = Shell("C:\MERCVB\TYP.EXE ", vbMinimizedFocus)
   
   
   
   ElseIf F928 = 1 Then
   Else  ' <>928
         DUM = Shell("C:\MERCVB\PRINT.EXE " + mfil, vbMinimizedFocus)
   End If
 Else
      If Len(Dir("c:\DION2.EXE")) > 1 Then
        ' DUM = EKTYP_ME_PRINT(mfil)
         MILSEC 5000
         FileCopy "C:\KENO.TXT", "C:\FOROL\FOROL.TXT"
      ElseIf Len(Dir("c:\KAMPOUR.txt")) > 1 Then
      End If
 End If
End If







Exit Function



create_index:

'db.Execute "CREATE INDEX eidko ON eid(kod);"
Resume Next



errorl:
    MsgBox Str(Erl) + "MM=" + MM 'RECOVER USING objErrorInfo
      'sfalma = 1
      'MsgBox "’¦ͺ ©«¦ §ε¦ " + Field + "  £.§ε¦¬ " + Str(i)
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
       For k = 1 To e
             MM = ePict(k)
             For I = 1 To eProw(k)
                 If k = 1 And m_No_of_seir = 1000 Then
                    ' ΜΗΝ ΒΑΖΕΙΣ ΚΕΝΗ ΣΕΙΡΑ ΣΤΗΝ ΑΡΧΗ
                 Else
                     Print #1, " "
                 End If
             Next
             syn_row = syn_row + eProw(k) 'συγκεντρώνω τις σειρές
             MM = mac(Trim(Trim(gm_f(k)))) ': m_m = ExecuteLine("gvar=" + MM)
      If eProw(k + 1) > eProw(k) Then 'αν πρόκειται να τυπώσει παρακάτω μην τυπώνεις δίπλα μου
           Print #1, Tab(ePcol(k)); Format(Left(MM, Len(ePict(k))), ePict(k));
      Else
           Print #1, Tab(ePcol(k)); Format(Left(MM, Len(ePict(k))), ePict(k));
      End If
      m1 = eProw(k)
    Next
                If SELIDA > 1 Then
                    Print #1, cSelida + Format(SELIDA, "##") ' ΑΛΛΑΓΗ ΣΕΛΙΔΑΣ
                End If
    
    For k = 1 To m_seir_eid - syn_row - 1
          Print #1, Space(80) 'Printer.Print:
    Next 'βρισκω την σειρά  των ειδών
    syn_row = m_seir_eid
Return















MHNYMA:
HandleError "Par1:ektypforma"
Resume Next
            



End Function
Function EKTYP_ME_PRINT(mfil)
Dim a$
  
Printer.FontName = "Arial"
Printer.FONTSIZE = 12
Printer.FontBold = True

  Open mfil For Input As #1
  Do While Not EOF(1)
     Input #1, a$
     Printer.Print a$
  Loop
  Close #1
  Printer.EndDoc
End Function

Function EKTYP_ME_serial(mfil, com)
Dim a$
'Dim com1

'MSComm1.CommPort = com

'MSComm1.PortOpen = True

 
 Open mfil For Input As #1
  Do While Not EOF(1)
     Input #1, a$
     If IsNull(a$) Then a$ = ""
 '    MSComm1.Output = a$ + Chr(10)
  '   MILSEC 100
  Loop
  Close #1
'  Printer.EndDoc
'MSComm1.PortOpen = False


End Function



Function mac(ByVal S As String) As String
'δινω ενα στρινγκ και μου επιστρέφει μία τιμή
Dim X
'On Error GoTo 0
MACtEXT.Text = "   "
On Error GoTo err2
X = "MACtEXT.Text = " + S
fSCR.ExecuteStatement X

If F928 = 1 Then
    mac = MACtEXT.Text
Else
    mac = to437(MACtEXT.Text)
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
   Dim k As Integer
   Dim genEkpt
   genEkpt = gVal(Text5.Text)
   Dim MEKPT, MDOR
   Dim ANS
   
On Error Resume Next


   If Left(f_pelMERC, 3) = "SOK" And genEkpt > 0 Then
       MEKPT = SynoloG.Caption * genEkpt / 100
       ANS = MsgBox("ΕΚΠΤΩΣΗ " + Format(MEKPT, "###.00") + Chr(13) + "ΔΩΡΟ " + Format(2 * MEKPT, "###.00") + Chr(13) + " ΝΑ ΠΡΟΧΩΡΓΗΣΩ ΣΤΗΝ ΕΚΠΤΩΣΗ; Ν/Ο", vbYesNo)
       If ANS = vbNo Then
          Exit Sub
       End If
   End If
   
   
   
   
   If genEkpt > 0 And genEkpt < 100 Then
      
      Text5.Text = genEkpt
      For k = 1 To Grid1.ROWS - 1
         If Val(Grid1.TextMatrix(k, f_p)) <> 0 Then
            Grid1.TextMatrix(k, f_e) = (genEkpt + Val(Grid1.TextMatrix(k, f_e))) - (genEkpt * Val(Grid1.TextMatrix(k, f_e))) / 100
            Grid1.row = k: FindSynolo
         End If
      Next
   End If
   
   
   Dim POS2
   Dim telPOSO As Single, log As Double
   Dim LOG2, EKPT2
   
   If Left(LTrim(Text5.Text), 1) = "*" Then
      telPOSO = gVal(mID(LTrim(Text5.Text), 2, 10))
      
      log = telPOSO / SynoloG  'συντελεστης σμίκρυνσης
      If telPOSO < SynoloG Then
           LOG2 = SynoloG - telPOSO 'συντελεστης σμίκρυνσης
           EKPT2 = LOG2 / SynoloG * 100
      End If
      
      
      Dim arxSynoloG As Single
      arxSynoloG = SynoloG
      
      
      For k = 1 To Grid1.ROWS - 1
         If Val(Grid1.TextMatrix(k, F_T)) <> 0 Then
                     
             Grid1.TextMatrix(k, f_e) = Round(EKPT2, 2) ' + Val(Grid1.TextMatrix(k, f_e))) - (EKPT2 * Val(Grid1.TextMatrix(k, f_e))) / 100
            ' Grid1.TextMatrix(k, F_T) = Round(log * Val(Grid1.TextMatrix(k, F_T)), 2)
            Grid1.row = k: FindSynolo
         End If
      Next
      
      
      Dim maxT As Single
      maxT = 0 'βρισκω την μεγαλύτερη τιμή για να προσθέσω +0.01
      Dim colmaxt As Long
      colmaxt = 0 ' σε ποια σειρά βρέθηκε
    If Abs(telPOSO - SynoloG) >= 0.005 Then
      log = telPOSO / SynoloG  'epanypologismos συντελεστη σμίκρυνσης
      For k = 1 To Grid1.ROWS - 1
         If Val(Grid1.TextMatrix(k, F_T)) <> 0 Then
            Grid1.TextMatrix(k, F_T) = Round(log * Val(Grid1.TextMatrix(k, F_T)), 3)
            If Grid1.TextMatrix(k, F_T) > 0 Then
               maxT = Grid1.TextMatrix(k, F_T)
               colmaxt = k
            End If
            Grid1.row = k: FindSynolo
         End If
      Next
    End If
    
      
      
      
      Text5.Text = "    "
      
      Dim diaf
      diaf = telPOSO - SynoloG
      If Abs(diaf) >= 0.01 Then
         Grid1.TextMatrix(colmaxt, F_T) = Grid1.TextMatrix(colmaxt, F_T) + Round(Round(diaf, 2) / 1.23, 3)
         FindSynolo
      End If
      
      diaf = telPOSO - SynoloG
      If Abs(diaf) >= 0.005 Then
         Grid1.TextMatrix(colmaxt, F_T) = Grid1.TextMatrix(colmaxt, F_T) + Round(Round(diaf, 2) / 1.23, 3)
        FindSynolo
      End If

      diaf = telPOSO - SynoloG
      If Abs(diaf) >= 0.005 Then
         Grid1.TextMatrix(colmaxt, F_T) = Grid1.TextMatrix(colmaxt, F_T) + diaf
         f_psifiaAjias = 3
        FindSynolo
        f_psifiaAjias = 2
      End If
      
      
      
      
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
 Dim m_No_of_seir, I, k, M, C
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
Dim e, MM, S
Dim ektypoths
Dim m_syn_row, syn_row
Dim CDOK

Set fSCR = CreateObject("MSScriptControl.ScriptControl")

DOK.ConnectionString = gConnect
DOK.RecordSource = "select *from EID  WHERE MON IN ('ΚΙΛ','KIΛ') OR LEFT(ERG,2)='" + F_2CHFIA_ZYGIZOMENA + " ' ORDER BY ONO " '    LEFT(ERG,3)='210' "
DOK.Refresh


CTIM.Connect = "Access"
CTIM.DatabaseName = "C:\MERCVB\REPORTS\REPORTS.MDB"
CTIM.RecordSource = "select *from tim"
CTIM.Refresh


fSCR.language = "vbscript"
fSCR.AddObject "cTIM", CTIM
fSCR.AddObject "DOK", DOK 'cEGGTIM
fSCR.AddObject "cMEM", CMEM
fSCR.AddObject "mactext", MACtEXT
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
         For k = 1 To 4
             MM = mac(Trim(Trim(gm_f(k))))                   ': m_m = ExecuteLine("gvar=" + MM)
             S2 = S2 + ALIGNFORMAT(MM, gpic(k)) '
         Next
         Print #1, S2
         DOK.Recordset.MoveNext
       Loop
 Close #1
 DOK.Recordset.Close

Set fSCR = Nothing

End Function

Public Function ALIGNFORMAT(MM, pic)
' RIGHT ALIGNMENT
If Left(pic, 1) = "#" Then
      ALIGNFORMAT = Right(Space(Len(pic)) + MM, Len(pic))
Else
' LEFT ALIGNMENT
      ALIGNFORMAT = Left(MM + Space(Len(pic)), Len(pic))
End If


End Function



Sub AdjustDataGridColumns(intRecord As Integer, _
           intField As Integer, _
           Optional AccForHeaders As Boolean)


GridEidon.SetFocus

'Exit Sub


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
    Dim MM
    If intField - 1 > 3 Then MM = 3 Else MM = intField - 1
    
    Dim FF
    If intRecord - 1 > 20 Then FF = 20 Else FF = intRecord - 1
    
    
    
    For Col = 0 To MM 'intField - 1
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
       
       f_AkyrParas = Left(LTrim(R("akyr")), 1)
       
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
Dim k As Integer

For k = 1 To Len(X)
  pol = pol + "'" + mID$(X, k, 1) + "',"
Next

pol = Left(pol, Len(pol) - 2)
pol = mID$(pol, 2, Len(pol) - 1)

Get_meKomma = pol

End Function




'Private Function FIND_PARAMETROI()
'Dim K
'  For K = 1 To 200
'     If FORMA = F_AR_parametroi(K, 0) And Var = F_AR_parametroi(K, 1) Then
'          FIND_PARAMETROI = F_AR_parametroi(K, 2)
'          Exit For
'     End If
'  Next
'
'
'
'End Function

Function get_kerdos()
'---------------------------------------------------------------------------------------------------------
Dim kerdos As Single
Dim R As New ADODB.Recordset
Dim k, C, P, t, e, F, TIMH
On Error Resume Next

kerdos = 0
'ypologizei to kerdos toy timologioy
 For k = 1 To Grid1.ROWS - 1
     If Len(Grid1.TextMatrix(k, f_k)) > 0 Then
            Grid1.row = k
            C = Grid1.Col  ' απομνημονεύω την τρέχουσα στήλη kai σειρά
            Grid1.Col = f_p: P = Grid1.Text
            Grid1.Col = F_T: t = Grid1.Text
            Grid1.Col = f_e: e = Grid1.Text
            'timh polisis
            TIMH = Round(Val(t) * (100 - Val(e)) / 100, f_psifiaAjias)
            R.Open "select XTI FROM EID WHERE KOD='" + Grid1.TextMatrix(k, f_k) + "'", Gdb, adOpenDynamic, adLockOptimistic
            If IsNull(R(0)) Then
               kerdos = kerdos + Val(P) * TIMH
            Else
               kerdos = kerdos + Val(P) * (TIMH - R(0))
            End If
     End If
               
 Next
 get_kerdos = kerdos
            
End Function
