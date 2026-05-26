VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form pelat3 
   BackColor       =   &H00FF0000&
   Caption         =   "Παρουσίαση"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9045
   ScaleWidth      =   15150
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Bindings        =   "pelat3.frx":0000
      Height          =   4692
      Left            =   120
      TabIndex        =   25
      Top             =   4080
      Width           =   12612
      _ExtentX        =   22251
      _ExtentY        =   8281
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
      Splits(0).RecordSelectorWidth=   1005
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3043"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2963"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=3043"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2963"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=232,.bold=0,.fontsize=780,.italic=0"
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
   Begin MSComCtl2.DTPicker eos 
      Height          =   375
      Left            =   5610
      TabIndex        =   11
      Top             =   1320
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   309002241
      CurrentDate     =   38358
   End
   Begin VB.CommandButton EXODOS 
      BackColor       =   &H80000010&
      Caption         =   "Εξοδος"
      Height          =   375
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3600
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   8520
      TabIndex        =   18
      Top             =   360
      Width           =   4065
   End
   Begin VB.CheckBox rCheck 
      Caption         =   "Χρησιμοποιώ το επιλεγμενο report"
      Height          =   315
      Left            =   8520
      TabIndex        =   17
      Top             =   3120
      Width           =   2085
   End
   Begin VB.CommandButton ENHLIKIOSI_YPOLOIPON 
      Caption         =   "Υπόλοιπα σε καθυστέρηση"
      Height          =   375
      Left            =   8400
      TabIndex        =   16
      ToolTipText     =   "Υπολογίζει τα υπόλοιπα σε καθυστερήσεις 30,60,90,120 ημερών"
      Top             =   3600
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6345
      Top             =   8820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   8520
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   2760
      Width           =   4050
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Αποθήκευση Report"
      Height          =   375
      Left            =   10830
      TabIndex        =   14
      Top             =   3090
      Width           =   1740
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   4380
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.ListBox List1 
      Height          =   2205
      ItemData        =   "pelat3.frx":0014
      Left            =   240
      List            =   "pelat3.frx":0016
      MultiSelect     =   1  'Simple
      TabIndex        =   8
      Top             =   1080
      Width           =   3915
   End
   Begin VB.CommandButton cmdKRITHRIA 
      Caption         =   "Φίλτρα εγγραφών που θα τυπωθούν"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   1740
      Width           =   3345
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2040
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Mercvb\reports\par3-3.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      SelectionFormula=   $"pelat3.frx":0018
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GR1 
      Bindings        =   "pelat3.frx":009D
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   9480
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   2143
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc EGG 
      Height          =   375
      Left            =   3165
      Top             =   8700
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "EGG"
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   120
      Top             =   8160
      Visible         =   0   'False
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
      Caption         =   "DATA1"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GR2 
      Bindings        =   "pelat3.frx":00AF
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   65535
      Cols            =   10
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Εκτύπωση"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdYPOLOGISMOS 
      Caption         =   "Υπολογισμός"
      Height          =   375
      Left            =   4350
      TabIndex        =   2
      Top             =   3600
      Width           =   3975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "pelat3.frx":00C3
      Left            =   360
      List            =   "pelat3.frx":00D3
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   375
      Width           =   3495
   End
   Begin VB.ComboBox COMBO_PELPROM 
      Height          =   315
      ItemData        =   "pelat3.frx":0118
      Left            =   120
      List            =   "pelat3.frx":0122
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3600
      Width           =   4095
   End
   Begin MSComCtl2.DTPicker apo 
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   720
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   309395457
      CurrentDate     =   38358
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ετοιμες Εκτυπώσεις"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   8400
      TabIndex        =   21
      Top             =   120
      Width           =   4335
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ταξινόμηση & Στήλες"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   4095
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Επιλογή Στηλών που θα τυπωθούν"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   26
         Top             =   720
         Width           =   3510
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Φίλτρα"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4320
      TabIndex        =   22
      Top             =   120
      Width           =   3975
      Begin VB.CheckBox chkenergoi 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ενεργοί μόνο"
         Height          =   255
         Left            =   1800
         TabIndex        =   28
         Top             =   3120
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Κείμενο επικεφαλίδας"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   27
         Top             =   2280
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Από"
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
         Left            =   360
         TabIndex        =   24
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Εως"
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
         Left            =   360
         TabIndex        =   23
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "έως"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3555
      TabIndex        =   12
      Top             =   960
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kίνηση από"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3510
      TabIndex        =   10
      Top             =   585
      Width           =   930
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Κινήθηκαν πριν τις"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2940
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "pelat3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Dim F_RUN           As Integer

Dim Fsql            As String, FIsHidden As Boolean

Dim F_PLATOS_FORMAS As Long

Dim F_LHM1          As String

Dim F_LHM2          As String


Dim f_lHM3          As String

Dim f_lHM4          As String

Dim f_lHM5          As String

Dim f_lHM6          As String

Dim f_lab1          As String

Dim f_lab2          As String

Dim f_lab3          As String

Dim f_lab4          As String

Dim f_lab5          As String

Dim f_lab6          As String

Dim f_num1          As String

Dim f_num2          As String

Dim f_num3          As String

Dim f_num4          As String


Dim f_CO1          As String

Dim f_CO2          As String

Dim f_CO3          As String

Dim f_CO4          As String

Dim f_CO5          As String

Dim F_MONOENERGOS As Integer



Dim f_PALIAXRONIA    ' = Trim(FindParametroi(1,"PAR1", "F_PALIAXRONIA", "2005", "Γ.18 DSN ΠΑΛΙΑΣ ΧΡΟΝΙΑΣ"))
Attribute f_PALIAXRONIA.VB_VarUserMemId = 1073938455

Dim fList1ARX    ' arxikh epilogh toy List1
Attribute fList1ARX.VB_VarUserMemId = 1073938456

Dim F_ALIG(10), F_mWidth(10), F_sum(10)
Attribute F_ALIG.VB_VarUserMemId = 1073938457

Dim f_list1(30, 4)
Attribute f_list1.VB_VarUserMemId = 1073938460

'

Private Sub cmdYPOLOGISMOS_Click()

        'ypologismos ypoloipon
        '<EhHeader>
        On Error GoTo cmdYPOLOGISMOS_Click_Err

        '</EhHeader>
        Dim x As String, order As String

        Dim k As Integer, R As Recordset    ', sum As Single

        Dim mBUFF

        Dim execX, file

100     If COMBO_PELPROM.Text = COMBO_PELPROM.List(0) Then mBUFF = "e" Else mBUFF = "r"

110     F_RUN = 1

        'Data1.DatabaseName = gDir
120     data1.ConnectionString = gConnect

130     GR2.Clear

        On Error Resume Next

        'On Error GoTo Command1_Click_Err
        ' ΥΠΟΛΟΓΙΖΩ ΤΑ ΤΕΛΕΥΤΑΙA ΥΠΟΛΟΙΠΑ ΑΠΟ ΤΟ EGG
140     Gdb.Execute "DROP TABLE EGG1"

        Dim L As Long

150     Gdb.Execute "UPDATE PEL  SET TYP=0,XRESYN=0,PISSYN=0 FROM" & " PEL WHERE EIDOS='" + mBUFF + "'", L
        Gdb.Execute "UPDATE PEL  SET AYP=0 WHERE (AYP IS NULL) AND  EIDOS='" + mBUFF + "'", L
        'If gXEIRISTHS >= 8 Then

        On Error GoTo cmdYPOLOGISMOS_Click_Err

160     Gdb.Execute "SELECT EIDOS as MEIDOS,KOD,SUM(CASE WHEN XREOSI IS NULL  THEN 0 ELSE XREOSI  END ) AS SX," & " SUM(CASE WHEN PISTOSI IS NULL  THEN 0 ELSE PISTOSI  END ) AS SP INTO EGG1" & " FROM EGG where EIDOS='" + mBUFF + "' AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<='" + Format(eos, "MM/DD/YYYY") + "'  GROUP BY EIDOS,KOD ", L

        'Else

        'Gdb.Execute "SELECT EIDOS AS MEIDOS,KOD,SUM(CASE WHEN XREOSI IS NULL  THEN 0 ELSE XREOSI  END ) AS SX," _
        '& " SUM(CASE WHEN PISTOSI IS NULL  THEN 0 ELSE PISTOSI  END ) AS SP INTO EGG1" _
        '& " FROM EGG where HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<='" + Format(eos, "MM/DD/YYYY") + "' GROUP BY EIDOS,KOD "

        'End If

        'Gdb.Execute "CREATE VIEW dbo.[UPD_YPOL_PEL] as " _
        '& " SELECT    dbo.PEL.XRESYN,dbo.PEL.PISSYN,dbo.PEL.TYP, dbo.PEL.KOD, dbo.PEL.EIDOS, dbo.EGG1.SP, dbo.EGG1.SX, dbo.PEL.AYP" _
        '& "  FROM         dbo.PEL INNER JOIN" _
        '& "  dbo.EGG1 ON dbo.PEL.EIDOS = dbo.EGG1.EIDOS AND dbo.PEL.KOD = dbo.EGG1.KOD"

170     Gdb.Execute "UPDATE PEL  SET TYP=AYP WHERE EIDOS='" + mBUFF + "'"

        'pelates
180     If COMBO_PELPROM.Text = COMBO_PELPROM.List(0) Then    '  "e",
            'Gdb.Execute "UPDATE UPD_YPOL_PEL  SET TYP=AYP+SX-SP,XRESYN=SX,PISSYN=SP"

            ' If gXEIRISTHS >= 8 Then
190         Gdb.Execute "UPDATE PEL  SET TYP=AYP+(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )" & " - (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP " & " from PEL left JOIN EGG1 ON PEL.KOD=EGG1.KOD AND PEL.EIDOS=EGG1.MEIDOS WHERE PEL.EIDOS='e' ", L
            ' Else
            '   Gdb.Execute "UPDATE PEL  SET TYP=AYP+(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )" _
            '   & " - (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" _
            '   & " PEL left JOIN EGG1 ON PEL.KOD=EGG1.KOD AND PEL.EIDOS=EGG1.MEIDOS", L
            ' End If

        Else

            On Error GoTo cmdYPOLOGISMOS_Click_Err

            'Gdb.Execute "UPDATE UPD_YPOL_PEL  SET TYP=AYP+SP-SX,XRESYN=SX,PISSYN=SP"
            ' If gXEIRISTHS = 8 Or gXEIRISTHS = 2 Then

200         Gdb.Execute "UPDATE PEL  SET TYP= AYP-(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )+" & " (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" & " PEL INNER JOIN EGG1 ON PEL.KOD=EGG1.KOD AND PEL.EIDOS=EGG1.MEIDOS "
            '   Else

            '            Gdb.Execute "UPDATE PEL  SET TYP=-AYP-(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )+" _
            '            & " (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" _
            '            & " PEL INNER JOIN EGG1 ON PEL.KOD=EGG1.KOD AND PEL.EIDOS=EGG1.MEIDOS"
            '  '   End If

        End If

210     x = "select '@@@' AS [Α_Α],left(KOD,12) AS [Κωδ],LEFT(EPO,25) as [Επωνυμία]"
220     execX = "select LEFT(KOD,5) AS Kωδικός,LEFT(EPO,25) as Eπωνυμία"

        '& " STR(TYP,10,2) AS [Υπόλοιπο] ,THL as [Τηλέφωνο] ,DIE as [Διεύθυνση]," _
        '& " STR(AYP,10,2)as [Αρχ.Υπόλ.] ,EIDOS,KODGAL as [Κατηγορία]," _
        '& " LASTUPDT as [Τελ.Ενημέρωση],XRVMA,CONVERT(decimal(10,2),XRESYN) ,CONVERT(decimal(10,2),PISSYN) from PEL where" _
        '& " EIDOS='" + IIf(Combo1.Text = Combo1.List(0), "e", "r") + "'"

        'If Format(DTPicker1.Value, "DD/MM/YYYY") <> Format(Now, "DD/MM/YYYY") Then
        '   x = x + " and LASTUPDT <= '" + Format(DTPicker1.Value, "DD/MM/yyyy") + "' "
        'End If
        '
        'If IsNumeric(apo.Text) And IsNumeric(eos) Then
        '   x = x + " and  TYP >=" + apo.Text + " and TYP<=" + eos.Text
        'End If

        '0ΔΙΕΥΘΥΝΣΗ
        '1ΧΡΕΩΣΗ ΕΤΟΥΣ
        '2ΠΙΣΤΩΣΗ ΕΤΟΥΣ
        '3ΕΠΑΓΓΕΛΜΑ
        '4ΑΦΜ
        '5ΥΠΟΛΟΙΠΟ
        '6ΤΗΛΕΦΩΝΟ

        Dim mSYN As String

230     mSYN = ""

        Dim mc

240     mc = 2

        Dim EPIK(10)

        Dim temp

250     For k = 0 To List1.ListCount - 1

260         If List1.Selected(k) Then

                'ο πραγματικός δείκτης (γιατί μπορεί να μετακινήθηκε)
270             temp = Val(Right(List1.List(k), 2))
280             x = x + f_list1(temp, 0)
290             mc = mc + 1
                'x = x + f_list1(Val(Right(List1.List(KK), 2)), 0)
300             F_sum(mc) = f_list1(temp, 1)
310             F_ALIG(mc) = f_list1(temp, 2)
320             F_mWidth(mc) = f_list1(temp, 3)

                '  TDBGrid.Columns(mc).ValueItems.Presentation = dbgCheckBox

            End If

        Next

        'TDBGrid11.Columns(1).ValueItems.Presentation = dbgCheckBox

330     file = "PEL"

        'execX = execX + " into " + file + " IN 'c:\MERCVB\EXCEL\REPORTS.XLS' 'Excel 5.0;'  from PEL where EIDOS='" + IIf(Combo1.Text = Combo1.List(0), "e", "r") + "'"

340    ' If gUserId = 2 Then
350    '     x = x + " from PEL where  KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%'  AND EIDOS='" + IIf(COMBO_PELPROM.Text = COMBO_PELPROM.List(0), "e", "r") + "'"
       ' Else
       
       If F_MONOENERGOS = 0 Then  'PEL.KOD NOT IN (SELECT LTRIM(STR(TIMH)) FROM PINAKES WHERE (NOT TIMH IS NULL) AND  TYPOS=12) ANd
           x = x + " ,'' as [.] from PEL where  EIDOS='" + IIf(COMBO_PELPROM.Text = COMBO_PELPROM.List(0), "e", "r") + "'"
       
       Else
       
          If chkenergoi.Value = vbChecked Then  'AND PEL.KOD NOT IN (SELECT LTRIM(STR(TIMH)) FROM PINAKES WHERE (NOT TIMH IS NULL) AND  TYPOS=12)
              x = x + " ,'' as [.] from PEL where ENERGOS=1  AND EIDOS='" + IIf(COMBO_PELPROM.Text = COMBO_PELPROM.List(0), "e", "r") + "'"
          Else
            'PEL.KOD NOT IN (SELECT LTRIM(STR(TIMH)) FROM PINAKES WHERE (NOT TIMH IS NULL) AND  TYPOS=12) AND
360           x = x + " ,'' as [.] from PEL where  EIDOS='" + IIf(COMBO_PELPROM.Text = COMBO_PELPROM.List(0), "e", "r") + "'"
          End If

       End If


        'αν θα κανει σουμες στην στήλη
370     For k = 0 To mc

380         If F_sum(k) = 0 Then mSYN = mSYN + "0" Else mSYN = mSYN + "1"
        Next

390     If Len(Fsql) > 2 Then
400         Fsql = Text1.Text
410         x = x + " AND " + Fsql
420         execX = execX + " AND " + Fsql
        End If

430     order = "KOD"

440     For k = 0 To 2

450         If Combo2.Text = Combo2.List(0) Then
460             order = "KOD"
470         ElseIf Combo2.Text = Combo2.List(1) Then
480             order = "EPO"
490         ElseIf Combo2.Text = Combo2.List(2) Then
500             order = "TYP"
510         ElseIf Combo2.Text = Combo2.List(3) Then
520             order = "LASTUPDT"

            End If

        Next

530     x = x + " order by " + order
540     execX = execX + " order by " + order

550     If rCheck.Value = vbChecked Then
560         x = Text2
        Else
570         Text2.Text = x
        End If

580     data1.RecordSource = x

        'DBGrid1.Columns(2).NumberFormat = "###,###,##0.00"
        'On Error Resume Next
590     data1.Refresh

600     TDBGrid.Splits(0).columns(0).Locked = True
610     TDBGrid.Splits(0).columns(1).Locked = True

        Dim MCC

620     MCC = 2

630     For k = 0 To List1.ListCount - 1

640         If List1.Selected(k) Then

                'ο πραγματικός δείκτης (γιατί μπορεί να μετακινήθηκε)
650             temp = Val(Right(List1.List(k), 2))
                'X = X + f_list1(temp, 0)
660             MCC = MCC + 1

                ' F_sum(mc) = f_list1(temp, 1)
670             If InStr(f_list1(temp, 0), "NUM") > 0 Then
680                 TDBGrid.columns(MCC).ValueItems.Presentation = dbgCheckBox
                End If

            End If

        Next

        ' TDBGrid.Columns(0).ValueItems.Presentation = dbgCheckBox

690     GR2.ColWidth(1) = 200

        'Set R = DATA1.Recordset
        'On Error Resume Next
        'Data1.Recordset.MoveFirst
        Dim msum

700     msum = 0
        'Dim mc

710     GR2.ColWidth(1) = 600
720     GR2.ColWidth(2) = 2600

730     If rCheck.Value = vbChecked Then
           For k = 2 To List1.ListCount - 1
              F_sum(k) = 1
           Next
            ALIGNGRID 1
        Else
740         ALIGNGRID 0
        End If

750     If Check1 Then

            Dim PELA

760         If mBUFF = "e" Then PELA = "ΠΕΛΑΤΕΣ" Else PELA = "ΠΡΟΜΗΘΕΥΤΕΣ"

770         PELA = PELA + "     AΠΟ " + Format(apo, "DD/MM/YYYY") + " ΕΩΣ " + Format(eos, "DD/MM/YYYY")

            On Error GoTo cmdYPOLOGISMOS_Click_Err    ' Resume Next

780         mSYN = mSYN + "11111111"

790         print3_xar x, mSYN, PELA, 0    'typos2
        Else

        End If

        Exit Sub

MHNYMA:
800     MsgBox Err.Description

810     Resume Next

        '<EhFooter>
        Exit Sub

cmdYPOLOGISMOS_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub ALIGNGRID(ByVal monosumes As Integer)

        '<EhHeader>
        On Error GoTo ALIGNGRID_Err

        '</EhHeader>
        Dim k As Integer

        Dim L As Integer
      If monosumes = 0 Then
100     For k = 3 To TDBGrid.columns.Count - 1    '   GR2.ColS - 1
            'If Len(epik(K)) > 1 Then
            'Gr2.TextMatrix(0, K) = epik(K) ' A/A
            ' GR2.ColAlignment(K) = F_ALIG(K)
110         GR2.ColWidth(k) = F_mWidth(k)
120         TDBGrid.columns(k).Width = F_mWidth(k)
            '       If rCheck.value = vbChecked Then
            '          GR2.ColWidth(K) = Data1.Recordset(K).DefinedSize * 30
            '       End If
            '

            ' End If
        Next
      End If

        TDBGrid.RowHeight = 400

        TDBGrid.RowHeight(3) = 600

130     TDBGrid.columns(0).Width = 0
140     TDBGrid.columns(1).Width = 1000
150     TDBGrid.columns(2).Width = 2600
160     GR2.AddItem " "

        Dim synola(20)

        ' BRISKO TIS SUMES
170     For k = 3 To GR2.ColS - 1
180         synola(k) = 0

190         If F_sum(k) = 1 Then

200             For L = 1 To GR2.rows - 1
210                 synola(k) = synola(k) + Val(GR2.TextMatrix(L, k))
220                 GR2.TextMatrix(L, k) = Format(GR2.TextMatrix(L, k), "####,##0.00")
230             Next L

                ' GR2.TextMatrix(GR2.rows - 1, K) = Format(SYNOLA(K), "####,##0.00")
240             TDBGrid.Splits(0).columns(k).FooterText = Format(synola(k), "######0.00")
                TDBGrid.Splits(0).columns(k).FooterFont.Size = 12 ' = Format(SYNOLA(k), "######0.00")
            End If

250     Next k

        TDBGrid.FooterFont.Size = 12
        
        On Error GoTo ALIGNGRID_Err

        ' BRISKO TIS SUMES
        '           For k = 3 To TDBGrid.Columns.Count - 1
        '               SYNOLA(k) = 0
        '               TDBGrid.Col = k
        '              If F_sum(k) = 1 Then
        '                 For l = 1 To DATA1.Recordset.RecordCount
        '                   TDBGrid.row = l
        '                   If Not IsNull(TDBGrid.Text) Then
        '                      SYNOLA(k) = SYNOLA(k) + Val(TDBGrid.Text)
        '                   End If
        '
        '                   On Error GoTo FF
        '                   ' GR2.TextMatrix(L, K) = Format(GR2.TextMatrix(L, K), "####,##0.00")
        '                 Next l
        '                 ' GR2.TextMatrix(GR2.rows - 1, K) = Format(SYNOLA(K), "####,##0.00")
        '                 TDBGrid.Splits(0).Columns(k).FooterText = Format(SYNOLA(k), "######0.00")
        '              End If
        '           Next k

260     GR2.TextMatrix(GR2.rows - 1, 2) = "Σύνολα"
        GR2.RowHeight(GR2.rows - 1) = 1600

        On Error Resume Next

270     GR2.ColAlignment(3) = 6    'RIGHT   1=LEFT
280     GR2.ColAlignment(0) = 6    'RIGHT   1=LEFT
290     GR2.ColAlignment(11) = 6    'RIGHT   1=LEFT
300     GR2.ColAlignment(12) = 6    'RIGHT   1=LEFT

310     For k = 1 To GR2.rows - 2
320         GR2.TextMatrix(k, 0) = k
        Next

330     If GR2.rows > 10 Then
340         GR2.TopRow = GR2.rows - 10
        End If

        Exit Sub

FF:

        'MsgBox Err.Description
350     Resume Next

        '<EhFooter>
        Exit Sub

ALIGNGRID_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.ALIGNGRID " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.ALIGNGRID " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub Typose()
'    Dim R, sum, N
'    'Set R = DATA1.Recordset
'
'
'
'    Printer.FONTSIZE = 8
'    Printer.FontName = "Courier New"
'    Printer.Font.Charset = 161
'    Printer.FontBold = True
'
'
'    Printer.Print Tab(4); "Κωδικός"; Tab(10); "Επωνυμία"; Tab(36); "Υπόλοιπο"; Tab(48); "Τηλέφωνο"
'    data1.Recordset.MoveFirst
'    sum = 0
'    N = 0
'    Do While Not R.EOF
'        N = N + 1
'        Printer.Print Format(N, "###"); Tab(4); data1.Recordset("KOD"); Tab(10); Left(data1.Recordset("EPO"), 25); Tab(36); Format(data1.Recordset("TYP"), "###,###.00"); Tab(48); data1.Recordset("thl")
'        sum = sum + data1.Recordset("typ")
'        data1.Recordset.MoveNext
'    Loop
'    Printer.Print
'    Printer.Print Tab(36); Format(sum, "###,###.00")
'End Sub
'Sub typos2()
'Dim R, sum, n, db As Database, db2 As Database
'
'Dim x As String, order As String
'Dim K As Integer
'
'On Error Resume Next
'
'
'If gConnect = "dBase III;" Then
'   Set db = OpenDatabase(gDir, False, False, gConnect)
'Else
'   Set db = OpenDatabase(gDir, False, False)
'End If
'
'
'
''If Len(Dir("c:\mercvb\reports\dok1.dbf")) > 5 Then
''  On Error Resume Next
''    Kill "c:\mercvb\reports\dok1.dbf"
''    If Len(Dir("c:\mercvb\reports\dok1.dbf")) > 5 Then
''       MsgBox "Ξαναδοκιμάστε σε λίγο. Το αρχείο δεν είναι διαθέσιμο"
''       Exit Sub
''    End If
''End If
'
'
'
'
'F_RUN = 1
'
''DATA1.DatabaseName = gDir
'Data1.ConnectionString = gConnect
'
'Set db2 = OpenDatabase("c:\mercvb\reports\reports.mdb") ', False, False, gConnect)
' db2.Execute "drop table dok1 "
'' db2.Close
'
'
'x = " select lastupdt,kod,epo,typ,thl,die,AYP,xresyn,EIDOS,KODGAL,XRVMA into dok1 in 'c:\mercvb\reports\reports.mdb'  from pel where eidos='" + IIf(Combo1.Text = Combo1.List(0), "e", "r") + "'"
'
''x = "select kod,epo,typ,thl,die,AYP,EIDOS,KODGAL  into dok1 in 'c:\mercvb\reports' 'dBase III;'  from pel where eidos='" + IIf(Combo1.Text = Combo1.List(0), "e", "r") + "'"
'If DTPicker1.Value <> Now Then
'   x = x + " and lastupdt <= #" + Format(DTPicker1.Value, "mm/dd/yyyy") + "# "
'End If
''If IsNumeric(apo.Text) And IsNumeric(eos) Then
''   x = x + " and  typ>=" + apo.Text + " and typ<=" + eos.Text
''End If
'
'order = "kod"
'For K = 0 To 2
'  If Combo2.Text = Combo2.List(0) Then
'     order = "kod"
'  ElseIf Combo2.Text = Combo2.List(1) Then
'     order = "epo"
'  ElseIf Combo2.Text = Combo2.List(2) Then
'     order = "typ"
'  ElseIf Combo2.Text = Combo2.List(3) Then
'     order = "lastupdt"
'
'  End If
'Next
'
'
'
'
'x = x + " order by " + order
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
'
'
'
'
'
'
'   db.Execute x
'db.Close
'
'' CrystalReport1.Action = 1
'
'
'End Sub

Private Sub Command11_Click()

        '<EhHeader>
        On Error GoTo Command11_Click_Err

        '</EhHeader>

100     If Len(Dir("c:\mercvb\queries\pel", vbDirectory)) = 0 Then
110         MkDir "c:\mercvb\queries\pel"
        End If

120     cd1.InitDir = "c:\mercvb\queries\pel"
130     cd1.ShowSave

        Dim F

140     F = cd1.FILENAME

150     If Len(F) < 1 Then Exit Sub
155     Open F For Output As #5
160     Print #5, Text2.Text
170     Close #5

        '<EhFooter>
        Exit Sub

Command11_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.Command11_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.Command11_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub EXODOS_Click()

        '<EhHeader>
        On Error GoTo EXODOS_Click_Err

        '</EhHeader>

100     Unload Me

110     If gApoMenu = True Then
120         SendKeys "%"
130         SendKeys "~"    '..--ENTER
140         SendKeys "{DOWN}"
150         SendKeys "{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

EXODOS_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub


Private Sub ENHLIKIOSI_YPOLOIPON_Click()

        Dim FF As New UDialog

100     FF.SHOW
110     FF.OKButton.Visible = False
120     FF.List1.Visible = False

130     FF.CancelButton.Caption = "ΔΙΑΚΟΠΗ."
140     FF.CancelButton.Top = 120
150     FF.CancelButton.Left = 120
160     FF.CancelButton.Width = 2895
170     FF.CancelButton.Height = 495

        'FF.Top = 3000
        'FF.Left = 3000

180     FF.Width = 3210
190     FF.Height = 810

200     FF.Left = Screen.Width / 2 - FF.Width / 2
210     FF.Top = Screen.Height / 2 - FF.Height / 2

        'ypologismos ypoloipon
        Dim x As String, order As String

        Dim k As Integer, R As New ADODB.Recordset    ', sum As Single

        Dim mBUFF

        Dim execX, file

        Dim hmeres As Long

        'hmeres = Val(InputBox("Ημέρες πίστωσης πελατών", "Πέρα από τις ημέρες αυτές θεωρείται καθ.υπόλοιπο", 60))
       Dim MXR As String
        Dim mPis
220     If COMBO_PELPROM.Text = COMBO_PELPROM.List(0) Then
             mBUFF = "e"
             MXR = "XREOSI"
             mPis = "PISTOSI"
        Else
             mBUFF = "r"
             MXR = "PISTOSI"
             mPis = "XREOSI"
        End If
        
        
        
        
230     F_RUN = 1
240     data1.ConnectionString = gConnect

250     GR2.Clear

        Dim dum2
        dum2 = ADD_FIELD("PEL", "NUMER3", "DECIMAL(18,2)")
256     dum2 = ADD_FIELD("PEL", "NUMER4", "DECIMAL(18,2)")
260     dum2 = ADD_FIELD("PEL", "NUMER5", "DECIMAL(18,2)")
270     dum2 = ADD_FIELD("PEL", "NUMER6", "DECIMAL(18,2)")
280     dum2 = ADD_FIELD("PEL", "NUMER7", "DECIMAL(18,2)")
290     dum2 = ADD_FIELD("PEL", "NUMER8", "DECIMAL(18,2)")




        On Error Resume Next

        '--arxh ---------------------------- ΥΠΟΛΟΓΙΖΩ ΤΑ ΤΕΛΕΥΤΑΙA ΥΠΟΛΟΙΠΑ ΑΠΟ ΤΟ EGG
300     Gdb.Execute "DROP TABLE EGG1"

        Dim L As Long

310     Gdb.Execute "UPDATE PEL  SET TYP=0,XRESYN=0,PISSYN=0 FROM" & " PEL WHERE EIDOS='" + mBUFF + "'", L

5320    Gdb.Execute "SELECT UPPER(EIDOS) as MEIDOS,KOD,SUM(CASE WHEN " + MXR + " IS NULL  THEN 0 ELSE " + MXR + "  END ) AS SX," & " SUM(CASE WHEN " + mPis + " IS NULL  THEN 0 ELSE " + mPis + "  END ) AS SP INTO EGG1" & " FROM EGG where HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<='" + Format(eos, "MM/DD/YYYY") + "' GROUP BY UPPER(EIDOS),KOD "

330     Gdb.Execute "UPDATE PEL  SET TYP=AYP,NUMER4=0"
340     Gdb.Execute "UPDATE PEL  SET TYP=AYP+(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )" & " - (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" & " PEL left JOIN EGG1 ON PEL.KOD=EGG1.KOD AND UPPER(PEL.EIDOS)=EGG1.MEIDOS", L

        '-telos ----------------------------- ΥΠΟΛΟΓΙΖΩ ΤΑ ΤΕΛΕΥΤΑΙA ΥΠΟΛΟΙΠΑ ΑΠΟ ΤΟ EGG

350     x = "select '@@@' AS [Α_Α],KOD AS [Κωδ],LEFT(EPO,25) as [Επωνυμία],"
360     x = x + "ROUND(NUMER8,2) AS [HM(0-30)],ROUND(NUMER4,2) AS [HM(30-60)],ROUND(NUMER5,2) AS [HM(60-90)],ROUND(NUMER6,2) AS [90-120],ROUND(NUMER7,2) AS [120-],ROUND(TYP,2) AS [ΥΠΟΛΟΙΠΟ],PVLHTHS "
        ' execX = "select LEFT(KOD,5) AS Kωδικός,LEFT(EPO,25) as Eπωνυμία,TYP-NUM4 AS [KAΘ.ΥΠΟΛΟΙΠΟ] "

        Dim mSYN As String

370     mSYN = ""

        Dim mc

380     mc = 2

        Dim EPIK(10)

        Dim temp

390     For k = 0 To List1.ListCount - 1

400         If List1.Selected(k) Then

                'ο πραγματικός δείκτης (γιατί μπορεί να μετακινήθηκε)
410             temp = Val(Right(List1.List(k), 2))
420             x = x + f_list1(temp, 0)
430             mc = mc + 1
                'x = x + f_list1(Val(Right(List1.List(KK), 2)), 0)
440             F_sum(mc) = f_list1(temp, 1)
450             F_ALIG(mc) = f_list1(temp, 2)
460             F_mWidth(mc) = f_list1(temp, 3)
            End If

        Next

470     file = "PEL"
        'execX = execX + " into " + file + " IN 'c:\MERCVB\EXCEL\REPORTS.XLS' 'Excel 5.0;'  from PEL where EIDOS='" + IIf(Combo1.Text = Combo1.List(0), "e", "r") + "'"

480     If gUserId = 2 Then
490         x = x + " from PEL where TYP>0 and KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%'  AND EIDOS='" + IIf(COMBO_PELPROM.Text = COMBO_PELPROM.List(0), "e", "r") + "'"
        Else
500         x = x + " from PEL where TYP>0 AND  EIDOS='" + IIf(COMBO_PELPROM.Text = COMBO_PELPROM.List(0), "e", "r") + "'"
        End If

        'αν θα κανει σουμες στην στήλη
510     For k = 0 To mc

520         If F_sum(k) = 0 Then mSYN = mSYN + "0" Else mSYN = mSYN + "1"
        Next

530     If Len(Fsql) > 2 Then
540         Fsql = Text1.Text
550         x = x + " AND " + Fsql
            ' execX = execX + " AND " + Fsql
        End If

560     order = "KOD"

570     For k = 0 To 2

580         If Combo2.Text = Combo2.List(0) Then
590             order = "KOD"
600         ElseIf Combo2.Text = Combo2.List(1) Then
610             order = "EPO"
620         ElseIf Combo2.Text = Combo2.List(2) Then
630             order = "TYP"
640         ElseIf Combo2.Text = Combo2.List(3) Then
650             order = "LASTUPDT"

            End If

        Next

660     x = x + " order by PVLHTHS," + order
670     execX = execX + " order by " + order

680     If rCheck.Value = vbChecked Then
690         x = Text2
        Else
700         Text2.Text = x
        End If

710     data1.RecordSource = x
720     data1.Refresh

730     data1.Recordset.MoveFirst

        ' ελεγχος των εγγραφων EGGTIM
        Dim SUMA As Single    'ΣΥΝΟΛΟ ΕΝΤΟΣ ΤΩΝ ΗΜΕΡΩΝ ΠΛΗΡΩΜΗΣ

        Dim HME2 As Date

        Dim hme3 As Date

        Dim hme4 As Date

        Dim hme5 As Date

        Dim TYP  As Single

        Dim sql  As String, nc As Long

740     HME2 = DateAdd("D", -30, Now)
750     hme3 = DateAdd("D", -60, Now)
760     hme4 = DateAdd("D", -90, Now)
770     hme5 = DateAdd("D", -120, Now)

        Dim ANS2 As Integer

780     ANS2 = MsgBox("Να υπολογίζει και την παλιά χρονιά;", vbYesNo)

        Dim SUMA4 As Single, SUMA5 As Single, SUMA6 As Single, SUMA7 As Single, SUMA8 As Single

        Dim SUMA3 As Single

        Dim SQL2  As String

        'Dim mSums(10) As Single, msK As Integer
        'For msK = 1 To 10: mSums(msK) = 0: Next
        '=============================================================================================================
790     Do While Not data1.Recordset.EOF    ' LOOP PELATON

            '  HME2 = DateAdd("D", -hmeres, Now)
            On Error GoTo ENILIKIOSI_YPOLOIPON_Click_Err

            '  Exit Sub   f_PALIAXRONIA
            '  sql2 = "SELECT * FROM EGG WHERE HME<'" + Format(Now, "MM/DD/YYYY") + "'"
            ' sql2 = sql2 + " AND KOD='" + DATA1.Recordset(1) + "' AND EIDOS='e' AND XREOSI>0 order by HME DESC "

            'sql2 = sql2 + " union SELECT * FROM " + f_PALIAXRONIA + ".dbo.EGG  WHERE HME<'" + Format(Now, "MM/DD/YYYY") + "'"
            ' sql2 = sql2 + " AND KOD='" + DATA1.Recordset(1) + "' AND EIDOS='e' AND XREOSI>0 order by HME DESC "

            Dim ll

            ll = Len(data1.Recordset(1))

800         SQL2 = "SELECT * FROM EGG WHERE HME<='" + Format(Now, "MM/DD/YYYY") + "'"
810         SQL2 = SQL2 + " AND LEFT(KOD," + str(ll) + ")='" + Left(data1.Recordset(1), ll) + "' AND EIDOS='" + mBUFF + "' AND " + MXR + ">0  "

            'υπολογιζει και την παλια χρονία
820         If ANS2 = vbYes Then
830             SQL2 = SQL2 + " union SELECT * FROM " + f_PALIAXRONIA + ".dbo.EGG  WHERE HME<'" + Format(Now, "MM/DD/YYYY") + "'"
840             SQL2 = SQL2 + " AND LEFT(KOD," + str(ll) + ")='" + Left(data1.Recordset(1), ll) + "' AND EIDOS='" + mBUFF + "' AND " + MXR + ">0 "
            End If

850         SQL2 = SQL2 + " order by HME DESC "

            On Error GoTo OUT

860         R.Open SQL2, Gdb, adOpenDynamic, adLockOptimistic

            '  R.MoveFirst
870         SUMA = 0
880         SUMA4 = 0: SUMA5 = 0: SUMA6 = 0: SUMA7 = 0
890         SUMA3 = 0

900         Do While Not R.EOF    ' LOOP EGGRAFON ENOS PELATH

910             If R("HME") >= HME2 Then
920                 SUMA3 = SUMA3 + R(MXR)  '<30 hmeres
                End If

930             If R("HME") >= hme3 And R("HME") < HME2 Then
940                 SUMA4 = SUMA4 + R(MXR)   '30-60 days
                End If

950             If R("HME") >= hme4 And R("HME") < hme3 Then
960                 SUMA5 = SUMA5 + R(MXR)  ' 60-90 hmeres
                End If

970             If R("HME") >= hme5 And R("HME") < hme4 Then
980                 SUMA6 = SUMA6 + R(MXR)
                End If

990             If R("HME") < hme5 Then
1000                SUMA7 = SUMA7 + R(MXR)
                End If

1010            R.MoveNext

            Loop

1020        TYP = data1.Recordset("ΥΠΟΛΟΙΠΟ")

1030        If SUMA3 > TYP Then
1040            SUMA3 = TYP: SUMA4 = 0: SUMA5 = 0: SUMA6 = 0: SUMA7 = 0
            End If

1050        If SUMA3 + SUMA4 > TYP Then
1060            SUMA4 = TYP - SUMA3: SUMA5 = 0: SUMA6 = 0: SUMA7 = 0
            End If

1070        If SUMA3 + SUMA4 + SUMA5 > TYP Then
1080            SUMA5 = TYP - SUMA3 - SUMA4: SUMA6 = 0: SUMA7 = 0
            End If

1090        If SUMA3 + SUMA4 + SUMA5 + SUMA6 > TYP Then
1100            SUMA6 = TYP - SUMA3 - SUMA4 - SUMA5: SUMA7 = 0
            End If

1110        If SUMA3 + SUMA4 + SUMA5 + SUMA6 + SUMA7 > TYP Then
1120            SUMA7 = TYP - SUMA3 - SUMA4 - SUMA5 - SUMA6
            End If

            'AN ERXETAI APO PERYSI TO YPOLOIPO
1130        If SUMA3 + SUMA4 + SUMA5 + SUMA6 + SUMA7 < TYP Then
1140            SUMA7 = TYP - SUMA3 - SUMA4 - SUMA5 - SUMA6
            End If

1150        sql = "UPDATE PEL  SET NUMER4=" + Replace(Format(SUMA4, "#######.00"), ",", ".") + ", NUMER5=" + Replace(Format(SUMA5, "#######.00"), ",", ".") + ", NUMER6=" + Replace(Format(SUMA6, "#######.00"), ",", ".") + ", NUMER7=" + Replace(Format(SUMA7, "#######.00"), ",", ".") + ", NUMER8=" + Replace(Format(SUMA3, "#######.00"), ",", ".") + " WHERE LEFT(KOD," + str(ll) + ")='" + Left(data1.Recordset(1), ll) + "' AND EIDOS='" + mBUFF + "'"

1160        Gdb.Execute sql, nc

1170        R.Close
'            mSums(4) = mSums(4) + SUMA4
'            mSums(5) = mSums(5) + SUMA5
'            mSums(6) = mSums(6) + SUMA6
'            mSums(7) = mSums(7) + SUMA7
'            mSums(8) = mSums(8) + SUMA3
            
1180        DoEvents
1190        Me.Caption = data1.Recordset(1)
1200        data1.Recordset.MoveNext

1210        If FF.CancelButton.Enabled = False Then

                Exit Do

            End If

            '  Exit Do
        Loop
        
        
        
        
        
        
        
        

1220    data1.Refresh

1230    TDBGrid.Splits(0).columns(0).Locked = True
1240    TDBGrid.Splits(0).columns(1).Locked = True
1250    GR2.ColWidth(1) = 200


TDBGrid.Splits(0).columns(1).Width = 600
TDBGrid.Splits(0).columns(2).Width = 2600
TDBGrid.GroupByCaption = "συρετε την επικεφαλίδα για ομαδοποίηση"


Dim RR As New ADODB.Recordset
RR.Open "select SUM(TYP) AS STYP,SUM(NUMER3) AS S3,sum(NUMER4) AS S4,SUM(NUMER5) AS S5,SUM(NUMER6) AS S6,SUM(NUMER7) AS S7,SUM(NUMER8) AS S8 FROM PEL WHERE TYP>0 AND EIDOS='" + mBUFF + "'", Gdb, adOpenDynamic, adLockOptimistic

'ROUND(NUMER8,2) AS [HM(0-30)],ROUND(NUMER4,2) AS [HM(30-60)],ROUND(NUMER5,2) AS [HM(60-90)],ROUND(NUMER6,2) AS [90-120],ROUND(NUMER7,2) AS [120-],ROUND(TYP,2) AS [ΥΠΟΛΟΙΠΟ] "

TDBGrid.Splits(0).columns(3).FooterText = nNull(RR("S8"))
TDBGrid.Splits(0).columns(4).FooterText = nNull(RR("S4"))
TDBGrid.Splits(0).columns(5).FooterText = nNull(RR("S5"))
TDBGrid.Splits(0).columns(6).FooterText = nNull(RR("S6"))
TDBGrid.Splits(0).columns(7).FooterText = nNull(RR("S7"))
TDBGrid.Splits(0).columns(8).FooterText = nNull(RR("STYP"))
RR.Close













'GroupByCaption Property
'Syntax
'TDBGrid.GroupByCaption = string
'--------------------------------------------------------------
'Read/Write at run time and design time.
'Property applies to GroupByCaption control (OLE DB only).
'Description
'This property sets or returns the text displayed in the grouping area when no group is defined.
'When the Outlook-Style Grouping property is set to 2 - Group, a grouping area is created above the grid. Until column headers have been moved there, the TDBGrid text is displayed in the grouping area. The default caption is: "Drag a column header here to group by that column" The maximum length of the caption string is 255 characters.
'Note
'This property is only supported by the OLE DB version of True DBGrid Pro.
'See Column for more information.

 











        Dim msum

1260    msum = 0
'1270    GR2.ColWidth(2) = 600
'1280    GR2.ColWidth(3) = 2600

1290    If rCheck.Value = vbChecked Then
        Else
1300       ' ALIGNGRID
        End If

1310    If Check1 Then

            Dim PELA

1320        If mBUFF = "e" Then PELA = "ΠΕΛΑΤΕΣ" Else PELA = "ΠΡΟΜΗΘΕΥΤΕΣ"

1330        PELA = PELA + "     AΠΟ " + Format(apo, "DD/MM/YYYY") + " ΕΩΣ " + Format(eos, "DD/MM/YYYY")

            On Error GoTo ENILIKIOSI_YPOLOIPON_Click_Err    ' Resume Next

1340        mSYN = mSYN + "11111111"
1350        print3_xar x, mSYN, PELA, 0    'typos2
        Else
        End If

1360    Unload FF

        Exit Sub

MHNYMA:
1370    MsgBox Err.Description

1380    Resume Next

OUT:
1390    MsgBox "λάθος " + Err.Description
        Unload FF
        '<EhFooter>
       'Resume Next
       Exit Sub

ENILIKIOSI_YPOLOIPON_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.Command3_Click " & "at line " & Erl

        Resume Next


End Sub



'Private Sub ENHLI00000000KIOSI_YPOLOIPON_Click()
'
'        '<EhHeader>
'        On Error GoTo ENHLIKIOSI_YPOLOIPON_Click_Err
'
'        '</EhHeader>
'
'
'     '  If COMBO_PELPROM.List(0) = COMBO_PELPROM.Text Then
'          ' PELATES OK
'     '  Else
'          ENILIKIOSI_PROM
'          Exit Sub
'     '  End If
'
'
'
'
'        Dim FF As New UDialog
'
'100     FF.SHOW
'110     FF.OKButton.Visible = False
'120     FF.List1.Visible = False
'
'130     FF.CancelButton.Caption = "ΔΙΑΚΟΠΗ"
'140     FF.CancelButton.Top = 120
'150     FF.CancelButton.Left = 120
'160     FF.CancelButton.Width = 2895
'170     FF.CancelButton.Height = 495
'
'        'FF.Top = 3000
'        'FF.Left = 3000
'
'180     FF.Width = 3210
'190     FF.Height = 810
'
'200     FF.Left = Screen.Width / 2 - FF.Width / 2
'210     FF.Top = Screen.Height / 2 - FF.Height / 2
'
'        'ypologismos ypoloipon
'        Dim X As String, order As String
'
'        Dim k As Integer, R As New ADODB.Recordset    ', sum As Single
'
'        Dim mBUFF
'
'        Dim execX, file
'
'        Dim hmeres As Long
'
'        'hmeres = Val(InputBox("Ημέρες πίστωσης πελατών", "Πέρα από τις ημέρες αυτές θεωρείται καθ.υπόλοιπο", 60))
'
'220     If COMBO_PELPROM.Text = COMBO_PELPROM.List(0) Then mBUFF = "e" Else mBUFF = "r"
'230     F_RUN = 1
'240     data1.ConnectionString = gConnect
'
'250     GR2.Clear
'
'        Dim dum2
'
'260     dum2 = ADD_FIELD("PEL", "NUM5", "SINGLE")
'270     dum2 = ADD_FIELD("PEL", "NUM6", "SINGLE")
'280     dum2 = ADD_FIELD("PEL", "NUM7", "SINGLE")
'290     dum2 = ADD_FIELD("PEL", "NUM8", "SINGLE")
'
'        On Error Resume Next
'
'        '--arxh ---------------------------- ΥΠΟΛΟΓΙΖΩ ΤΑ ΤΕΛΕΥΤΑΙA ΥΠΟΛΟΙΠΑ ΑΠΟ ΤΟ EGG
'300     Gdb.Execute "DROP TABLE EGG1"
'
'        Dim l As Long
'
'310     Gdb.Execute "UPDATE PEL  SET TYP=0,XRESYN=0,PISSYN=0 FROM" & " PEL WHERE EIDOS='" + mBUFF + "'", l
'
'5320    Gdb.Execute "SELECT UPPER(EIDOS) as MEIDOS,KOD,SUM(CASE WHEN XREOSI IS NULL  THEN 0 ELSE XREOSI  END ) AS SX," & " SUM(CASE WHEN PISTOSI IS NULL  THEN 0 ELSE PISTOSI  END ) AS SP INTO EGG1" & " FROM EGG where HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<='" + Format(eos, "MM/DD/YYYY") + "' GROUP BY UPPER(EIDOS),KOD "
'
'330     Gdb.Execute "UPDATE PEL  SET TYP=AYP,NUM4=0"
'340     Gdb.Execute "UPDATE PEL  SET TYP=AYP+(CASE WHEN  EGG1.SX IS NULL THEN 0 ELSE EGG1.SX END )" & " - (CASE WHEN  EGG1.SP IS NULL THEN 0 ELSE EGG1.SP END ),XRESYN=EGG1.SX,PISSYN=EGG1.SP FROM" & " PEL left JOIN EGG1 ON PEL.KOD=EGG1.KOD AND UPPER(PEL.EIDOS)=EGG1.MEIDOS", l
'
'        '-telos ----------------------------- ΥΠΟΛΟΓΙΖΩ ΤΑ ΤΕΛΕΥΤΑΙA ΥΠΟΛΟΙΠΑ ΑΠΟ ΤΟ EGG
'
'350     X = "select '@@@' AS [Α_Α],KOD AS [Κωδ],LEFT(EPO,25) as [Επωνυμία],"
'360     X = X + "ROUND(NUMER8,2) AS [0-30],ROUND(NUMER4,2) AS [30-60],ROUND(NUMER5,2) AS [60-90],ROUND(NUMER6,2) AS [90-120],ROUND(NUMER7,2) AS [120-],ROUND(TYP,2) AS [ΥΠΟΛΟΙΠΟ] "
'        ' execX = "select LEFT(KOD,5) AS Kωδικός,LEFT(EPO,25) as Eπωνυμία,TYP-NUM4 AS [KAΘ.ΥΠΟΛΟΙΠΟ] "
'
'        Dim mSYN As String
'
'370     mSYN = ""
'
'        Dim mc
'
'380     mc = 2
'
'        Dim EPIK(10)
'
'        Dim temp
'
'390     For k = 0 To List1.ListCount - 1
'
'400         If List1.Selected(k) Then
'
'                'ο πραγματικός δείκτης (γιατί μπορεί να μετακινήθηκε)
'410             temp = Val(Right(List1.List(k), 2))
'420             X = X + f_list1(temp, 0)
'430             mc = mc + 1
'                'x = x + f_list1(Val(Right(List1.List(KK), 2)), 0)
'440             F_sum(mc) = f_list1(temp, 1)
'450             F_ALIG(mc) = f_list1(temp, 2)
'460             F_mWidth(mc) = f_list1(temp, 3)
'            End If
'
'        Next
'
'470     file = "PEL"
'        'execX = execX + " into " + file + " IN 'c:\MERCVB\EXCEL\REPORTS.XLS' 'Excel 5.0;'  from PEL where EIDOS='" + IIf(Combo1.Text = Combo1.List(0), "e", "r") + "'"
'
'480     If gUserId = 2 Then
'490         X = X + " from PEL where TYP>0 and KOD NOT LIKE '997%'  AND KOD NOT LIKE '998%'  AND EIDOS='" + IIf(COMBO_PELPROM.Text = COMBO_PELPROM.List(0), "e", "r") + "'"
'        Else
'500         X = X + " from PEL where TYP>0 AND  EIDOS='" + IIf(COMBO_PELPROM.Text = COMBO_PELPROM.List(0), "e", "r") + "'"
'        End If
'
'        'αν θα κανει σουμες στην στήλη
'510     For k = 0 To mc
'
'520         If F_sum(k) = 0 Then mSYN = mSYN + "0" Else mSYN = mSYN + "1"
'        Next
'
'530     If Len(Fsql) > 2 Then
'540         Fsql = Text1.Text
'550         X = X + " AND " + Fsql
'            ' execX = execX + " AND " + Fsql
'        End If
'
'560     order = "KOD"
'
'570     For k = 0 To 2
'
'580         If Combo2.Text = Combo2.List(0) Then
'590             order = "KOD"
'600         ElseIf Combo2.Text = Combo2.List(1) Then
'610             order = "EPO"
'620         ElseIf Combo2.Text = Combo2.List(2) Then
'630             order = "TYP"
'640         ElseIf Combo2.Text = Combo2.List(3) Then
'650             order = "LASTUPDT"
'
'            End If
'
'        Next
'
'660     X = X + " order by " + order
'670     execX = execX + " order by " + order
'
'680     If rCheck.Value = vbChecked Then
'690         X = Text2
'        Else
'700         Text2.Text = X
'        End If
'
'710     data1.RecordSource = X
'720     data1.Refresh
'
'730     data1.Recordset.MoveFirst
'
'        ' ελεγχος των εγγραφων EGGTIM
'        Dim SUMA As Single    'ΣΥΝΟΛΟ ΕΝΤΟΣ ΤΩΝ ΗΜΕΡΩΝ ΠΛΗΡΩΜΗΣ
'
'        Dim HME2 As Date
'
'        Dim hme3 As Date
'
'        Dim hme4 As Date
'
'        Dim hme5 As Date
'
'        Dim TYP  As Single
'
'        Dim sql  As String, nc As Long
'
'740     HME2 = DateAdd("D", -30, Now)
'750     hme3 = DateAdd("D", -60, Now)
'760     hme4 = DateAdd("D", -90, Now)
'770     hme5 = DateAdd("D", -120, Now)
'
'        Dim ANS2 As Integer
'
'780     ANS2 = MsgBox("Να υπολογίζει και την παλιά χρονιά;", vbYesNo)
'
'        Dim SUMA4 As Single, SUMA5 As Single, SUMA6 As Single, SUMA7 As Single, SUMA8 As Single
'
'        Dim SUMA3 As Single
'
'        Dim Sql2  As String
'
'        '=============================================================================================================
'790     Do While Not data1.Recordset.EOF    ' LOOP PELATON
'
'            '  HME2 = DateAdd("D", -hmeres, Now)
'            On Error GoTo ENHLIKIOSI_YPOLOIPON_Click_Err
'
'            '  Exit Sub   f_PALIAXRONIA
'            '  sql2 = "SELECT * FROM EGG WHERE HME<'" + Format(Now, "MM/DD/YYYY") + "'"
'            ' sql2 = sql2 + " AND KOD='" + DATA1.Recordset(1) + "' AND EIDOS='e' AND XREOSI>0 order by HME DESC "
'
'            'sql2 = sql2 + " union SELECT * FROM " + f_PALIAXRONIA + ".dbo.EGG  WHERE HME<'" + Format(Now, "MM/DD/YYYY") + "'"
'            ' sql2 = sql2 + " AND KOD='" + DATA1.Recordset(1) + "' AND EIDOS='e' AND XREOSI>0 order by HME DESC "
'
'            Dim ll
'
'            ll = Len(data1.Recordset(1))
'
'800         Sql2 = "SELECT * FROM EGG WHERE HME<'" + Format(Now, "MM/DD/YYYY") + "'"
'810         Sql2 = Sql2 + " AND LEFT(KOD," + str(ll) + ")='" + Left(data1.Recordset(1), ll) + "' AND EIDOS='e' AND XREOSI>0  "
'
'            'υπολογιζει και την παλια χρονία
'820         If ANS2 = vbYes Then
'830             Sql2 = Sql2 + " union SELECT * FROM " + f_PALIAXRONIA + ".dbo.EGG  WHERE HME<'" + Format(Now, "MM/DD/YYYY") + "'"
'840             Sql2 = Sql2 + " AND LEFT(KOD," + str(ll) + ")='" + Left(data1.Recordset(1), ll) + "' AND EIDOS='e' AND XREOSI>0 "
'            End If
'
'850         Sql2 = Sql2 + " order by HME DESC "
'
'            On Error GoTo out
'
'860         R.Open Sql2, Gdb, adOpenDynamic, adLockOptimistic
'
'            '  R.MoveFirst
'870         SUMA = 0
'880         SUMA4 = 0: SUMA5 = 0: SUMA6 = 0: SUMA7 = 0
'890         SUMA3 = 0
'
'900         Do While Not R.EOF    ' LOOP EGGRAFON ENOS PELATH
'
'910             If R("HME") >= HME2 Then
'920                 SUMA3 = SUMA3 + R("XREOSI")  '<30 hmeres
'                End If
'
'930             If R("HME") >= hme3 And R("HME") < HME2 Then
'940                 SUMA4 = SUMA4 + R("XREOSI")   '30-60 days
'                End If
'
'950             If R("HME") >= hme4 And R("HME") < hme3 Then
'960                 SUMA5 = SUMA5 + R("XREOSI")  ' 60-90 hmeres
'                End If
'
'970             If R("HME") >= hme5 And R("HME") < hme4 Then
'980                 SUMA6 = SUMA6 + R("XREOSI")
'                End If
'
'990             If R("HME") < hme5 Then
'1000                SUMA7 = SUMA7 + R("XREOSI")
'                End If
'
'1010            R.MoveNext
'
'            Loop
'
'1020        TYP = data1.Recordset("ΥΠΟΛΟΙΠΟ")
'
'1030        If SUMA3 > TYP Then
'1040            SUMA3 = TYP: SUMA4 = 0: SUMA5 = 0: SUMA6 = 0: SUMA7 = 0
'            End If
'
'1050        If SUMA3 + SUMA4 > TYP Then
'1060            SUMA4 = TYP - SUMA3: SUMA5 = 0: SUMA6 = 0: SUMA7 = 0
'            End If
'
'1070        If SUMA3 + SUMA4 + SUMA5 > TYP Then
'1080            SUMA5 = TYP - SUMA3 - SUMA4: SUMA6 = 0: SUMA7 = 0
'            End If
'
'1090        If SUMA3 + SUMA4 + SUMA5 + SUMA6 > TYP Then
'1100            SUMA6 = TYP - SUMA3 - SUMA4 - SUMA5: SUMA7 = 0
'            End If
'
'1110        If SUMA3 + SUMA4 + SUMA5 + SUMA6 + SUMA7 > TYP Then
'1120            SUMA7 = TYP - SUMA3 - SUMA4 - SUMA5 - SUMA6
'            End If
'
'            'AN ERXETAI APO PERYSI TO YPOLOIPO
'1130        If SUMA3 + SUMA4 + SUMA5 + SUMA6 + SUMA7 < TYP Then
'1140            SUMA7 = TYP - SUMA3 - SUMA4 - SUMA5 - SUMA6
'            End If
'
'1150        sql = "UPDATE PEL  SET NUM4=" + Replace(Format(SUMA4, "#######.00"), ",", ".") + ", NUM5=" + Replace(Format(SUMA5, "#######.00"), ",", ".") + ", NUM6=" + Replace(Format(SUMA6, "#######.00"), ",", ".") + ", NUM7=" + Replace(Format(SUMA7, "#######.00"), ",", ".") + ", NUM8=" + Replace(Format(SUMA3, "#######.00"), ",", ".") + " WHERE LEFT(KOD," + str(ll) + ")='" + Left(data1.Recordset(1), ll) + "' AND EIDOS='e'"
'
'1160        Gdb.Execute sql, nc
'
'1170        R.Close
'
'1180        DoEvents
'1190        Me.Caption = data1.Recordset(1)
'1200        data1.Recordset.MoveNext
'
'1210        If FF.CancelButton.Enabled = False Then
'
'                Exit Do
'
'            End If
'
'            '  Exit Do
'        Loop
'
'1220    data1.Refresh
'
'1230    TDBGrid.Splits(0).columns(0).Locked = True
'1240    TDBGrid.Splits(0).columns(1).Locked = True
'1250    GR2.ColWidth(1) = 200
'
'        Dim msum
'
'1260    msum = 0
'1270    GR2.ColWidth(1) = 600
'1280    GR2.ColWidth(2) = 2600
'
'1290    If rCheck.Value = vbChecked Then
'        Else
'1300        ALIGNGRID
'        End If
'
'1310    If Check1 Then
'
'            Dim PELA
'
'1320        If mBUFF = "e" Then PELA = "ΠΕΛΑΤΕΣ" Else PELA = "ΠΡΟΜΗΘΕΥΤΕΣ"
'
'1330        PELA = PELA + "     AΠΟ " + Format(apo, "DD/MM/YYYY") + " ΕΩΣ " + Format(eos, "DD/MM/YYYY")
'
'            On Error GoTo ENHLIKIOSI_YPOLOIPON_Click_Err    ' Resume Next
'
'1340        mSYN = mSYN + "11111111"
'1350        print3_xar X, mSYN, PELA, 0    'typos2
'        Else
'        End If
'
'1360    Unload FF
'
'        Exit Sub
'
'MHNYMA:
'1370    MsgBox Err.Description
'
'1380    Resume Next
'
'out:
'1390    MsgBox "λάθος " + Err.Description
'
'        '<EhFooter>
'        Exit Sub
'
'ENHLIKIOSI_YPOLOIPON_Click_Err:
'        'MsgBox Err.Description & vbCrLf & _
'         "in ADOMERCNEW.pelat3.Command3_Click " & _
'         "at line " & Erl, _
'         vbExclamation + vbOKOnly, "Application Error"
'        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.Command3_Click " & "at line " & Erl
'
'        Resume Next
'
'        '</EhFooter>
'
'End Sub

Private Sub show_kartella()

        '================
        '--  ΚΑΡΤΕΛΛΑ ---
        '----------------
        '<EhHeader>
        On Error GoTo show_kartella_Err

        '</EhHeader>
        Dim X2 As String, s As String

        Dim yp As Single

100     If F_RUN = 0 Then Exit Sub
        '

110     GR2.Col = 1    ' gia na bro ton kodiko

120     If IsEmpty(GR2.Text) Then Exit Sub

        ' EGG.DatabaseName = gDir
130     egg.ConnectionString = gConnect

        '  On Error GoTo akiro
140     s = GR2.Text

150     If COMBO_PELPROM.Text = COMBO_PELPROM.List(0) Then X2 = "e" Else X2 = "r"

160     If gXEIRISTHS = 8 Then
170         If X2 = "e" Then
180             X2 = "e','E"
            Else
190             X2 = "r','R"
            End If
        End If

        'If gXEIRISTHS = 8 Then
200     egg.RecordSource = "select ATIM as [Αρ.Παρ], CONVERT(DATETIME, HME, 5)  as [Hμερ/νία],EID as [Ειδ],APA,AIT as [Αιτιολογία], XREOSI, PISTOSI ,str(PROOD,10,2) from EGG where EIDOS  IN ('" + X2 + "') and KOD='" + s + "' order by HME,ATIM;"
        'Else
        '   EGG.RecordSource = "select ATIM as [Αρ.Παρ], CONVERT(DATETIME, HME, 5)  as [Hμερ/νία],EID as [Ειδ],APA,AIT as [Αιτιολογία], XREOSI, PISTOSI ,str(PROOD,10,2) from EGG where EIDOS='" + x2 + "' and KOD='" + s + "' order by HME,ATIM;"
        'End If
210     egg.Refresh

        Dim R As New ADODB.Recordset, NX, NP, NY

220     NX = 5: NP = 6: NY = 7

230     R.Open "SELECT *FROM PEL WHERE EIDOS IN ('" + X2 + "') AND KOD='" + GR2.Text + "';", Gdb, adOpenDynamic, adLockOptimistic

240     If R.RecordCount = 0 Then R.Close: Exit Sub
250     yp = R("AYP")
        'On Error Resume Next
        ' R.Close

        Dim k, L, SX, SP, SU

260     SU = yp: SX = 0: SP = 0

270     For k = 1 To GR1.rows - 1
280         GR1.TextMatrix(k, 1) = Format(GR1.TextMatrix(k, 1), "dd/mm/yyyy")

290         If Val(GR1.TextMatrix(k, NX)) <> 0 Then
300             GR1.TextMatrix(k, NP) = ""
310             GR1.TextMatrix(k, NX) = Format(Val(GR1.TextMatrix(k, NX)), "#####.00")
            End If

320         If Val(GR1.TextMatrix(k, NP)) <> 0 Then
330             GR1.TextMatrix(k, NP) = Format(Val(GR1.TextMatrix(k, NP)), "#####.00")
340             GR1.TextMatrix(k, NX) = " "
            End If

350         SX = SX + Val(GR1.TextMatrix(k, NX))
360         SP = SP + Val(GR1.TextMatrix(k, NP))
            'SU = sx - sp

370         If R("eidos") = "e" Then
380             GR1.TextMatrix(k, NY) = Format(SU + SX - SP, "#####.00")
            Else
390             GR1.TextMatrix(k, NY) = Format(SU - SX + SP, "#####.00")
            End If

            'GR1.TextMatrix(K, nT) = Format(Val(GR1.TextMatrix(K, nT)), "######.00")

400         If k Mod 2 = 0 Then

410             For L = 0 To GR1.ColS - 1
420                 GR1.row = k: GR1.Col = L: GR1.CellBackColor = &HC0E0FF
                Next

            End If

        Next

430     R.Close
440     GR1.ColAlignment(NY) = 6    'RIGHT

450     GR1.Refresh

        '  EGG.Recordset.MoveFirst
        'Do While Not EGG.Recordset.EOF
        '  yp = yp + EGG.Recordset("xreosi")
        '  yp = yp - EGG.Recordset("pistosi")
        '  'EGG.Recordset.Edit
        '  EGG.Recordset("prood") = yp
        '  EGG.Recordset.Update
        '  DoEvents
        '  EGG.Recordset.MoveNext
        'Loop

OLD100:

        'MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
        '  EGG.Refresh
460     If GR1.rows > 9 Then
470         GR1.TopRow = GR1.rows - 7
        End If

        'MSFlexGrid1.SetFocus
        'MSFlexGrid1.SelectionMode = True
        Exit Sub

akiro:
480     GoTo OLD100

        '<EhFooter>
        Exit Sub

show_kartella_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.show_kartella " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.show_kartella " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdKRITHRIA_Click()

        '<EhHeader>
        On Error GoTo cmdKRITHRIA_Click_Err

        '</EhHeader>
        Dim N As Integer

100     N = 0

110     If Not FIsHidden Then

120         With KRITHRIA.Grid1
130             N = N + 1: .TextMatrix(N, 0) = "Κωδικός": .TextMatrix(N, 3) = "KOD": .TextMatrix(N, 4) = "c"
140             N = N + 1: .TextMatrix(N, 0) = "Επωνυμία": .TextMatrix(N, 3) = "EPO": .TextMatrix(N, 4) = "c"
150             N = N + 1: .TextMatrix(N, 0) = "Διεύθυνση": .TextMatrix(N, 3) = "DIE": .TextMatrix(N, 4) = "c"
160             N = N + 1: .TextMatrix(N, 0) = "Επάγγελμα": .TextMatrix(N, 3) = "EPA": .TextMatrix(N, 4) = "c"
170             N = N + 1: .TextMatrix(N, 0) = "Υπόλοιπο": .TextMatrix(N, 3) = "TYP": .TextMatrix(N, 4) = "n"
180             N = N + 1: .TextMatrix(N, 0) = "ΧΡΕΩΣΗ ΕΤΟΥΣ": .TextMatrix(N, 3) = "XRESYN": .TextMatrix(N, 4) = "n"
190             N = N + 1: .TextMatrix(N, 0) = "ΠΙΣΤΩΣΗ ΕΤΟΥΣ": .TextMatrix(N, 3) = "PISSYN": .TextMatrix(N, 4) = "n"
200             N = N + 1: .TextMatrix(N, 0) = "Πόλη": .TextMatrix(N, 3) = "POL": .TextMatrix(N, 4) = "c"
210             N = N + 1: .TextMatrix(N, 0) = "Hμ.Τελ.Ενημέρωσης": .TextMatrix(N, 3) = "LASTUPDT": .TextMatrix(N, 4) = "d"
220             N = N + 1: .TextMatrix(N, 0) = "Κατηγορία-1": .TextMatrix(N, 3) = "KODGAL": .TextMatrix(N, 4) = "c"
230             N = N + 1: .TextMatrix(N, 0) = "Τρόπος Πληρωμής": .TextMatrix(N, 3) = "ARPARAG": .TextMatrix(N, 4) = "c"
240             N = N + 1: .TextMatrix(N, 0) = "Τηλέφωνο": .TextMatrix(N, 3) = "THL": .TextMatrix(N, 4) = "c"

250             N = N + 1: .TextMatrix(N, 0) = "Εκπτωση": .TextMatrix(N, 3) = "PEK": .TextMatrix(N, 4) = "n"
260             N = N + 1: .TextMatrix(N, 0) = "Πλαφόν": .TextMatrix(N, 3) = "PLAFON": .TextMatrix(N, 4) = "n"

270             N = N + 1: .TextMatrix(N, 0) = "Καθ.ΦΠΑ": .TextMatrix(N, 3) = "TYPOS": .TextMatrix(N, 4) = "c"
280             N = N + 1: .TextMatrix(N, 0) = "Μεταφορική": .TextMatrix(N, 3) = "PLAISIO": .TextMatrix(N, 4) = "c"
290             N = N + 1: .TextMatrix(N, 0) = "Τρόπος Πληρωμής": .TextMatrix(N, 3) = "ARPARAG": .TextMatrix(N, 4) = "c"
300             N = N + 1: .TextMatrix(N, 0) = "Κατηγορία": .TextMatrix(N, 3) = "KODGAL": .TextMatrix(N, 4) = "c"
310             N = N + 1: .TextMatrix(N, 0) = "Πωλητής": .TextMatrix(N, 3) = "PVLHTHS": .TextMatrix(N, 4) = "c"

                'TYPOS   KAUESTVS  FPA   CHAR
                'PLAISIO  =METAFORIKH  CHAR
                'ARPARAG  TROPOS PLHRVMHS   CHAR
                'KODGAL  KATHGORIA   CHAR

320             N = N + 1: .TextMatrix(N, 0) = F_LHM1: .TextMatrix(N, 3) = "HM1": .TextMatrix(N, 4) = "d"
330             N = N + 1: .TextMatrix(N, 0) = F_LHM2: .TextMatrix(N, 3) = "HM2": .TextMatrix(N, 4) = "d"
340             N = N + 1: .TextMatrix(N, 0) = f_lHM3: .TextMatrix(N, 3) = "HM3": .TextMatrix(N, 4) = "d"
350             N = N + 1: .TextMatrix(N, 0) = f_lHM4: .TextMatrix(N, 3) = "HM4": .TextMatrix(N, 4) = "d"

360             N = N + 1: .TextMatrix(N, 0) = f_lab1: .TextMatrix(N, 3) = "CH1": .TextMatrix(N, 4) = "c"
370             N = N + 1: .TextMatrix(N, 0) = f_lab2: .TextMatrix(N, 3) = "CH2": .TextMatrix(N, 4) = "c"
380             N = N + 1: .TextMatrix(N, 0) = f_lab3: .TextMatrix(N, 3) = "CH3": .TextMatrix(N, 4) = "c"
390             N = N + 1: .TextMatrix(N, 0) = f_lab4: .TextMatrix(N, 3) = "CH4": .TextMatrix(N, 4) = "c"

400             N = N + 1: .TextMatrix(N, 0) = f_num1: .TextMatrix(N, 3) = "NUM1": .TextMatrix(N, 4) = "N"
410             N = N + 1: .TextMatrix(N, 0) = f_num2: .TextMatrix(N, 3) = "NUM2": .TextMatrix(N, 4) = "N"
420             N = N + 1: .TextMatrix(N, 0) = f_num3: .TextMatrix(N, 3) = "NUM3": .TextMatrix(N, 4) = "N"
430             N = N + 1: .TextMatrix(N, 0) = f_num4: .TextMatrix(N, 3) = "NUM4": .TextMatrix(N, 4) = "N"


435             N = N + 1: .TextMatrix(N, 0) = f_CO1: .TextMatrix(N, 3) = "COMB1": .TextMatrix(N, 4) = "c"
436             N = N + 1: .TextMatrix(N, 0) = f_CO2: .TextMatrix(N, 3) = "COMB2": .TextMatrix(N, 4) = "c"
437             N = N + 1: .TextMatrix(N, 0) = f_CO3: .TextMatrix(N, 3) = "COMB3": .TextMatrix(N, 4) = "c"
439             N = N + 1: .TextMatrix(N, 0) = f_CO4: .TextMatrix(N, 3) = "COMB4": .TextMatrix(N, 4) = "c"
                N = N + 1: .TextMatrix(N, 0) = f_CO5: .TextMatrix(N, 3) = "COMB5": .TextMatrix(N, 4) = "c"




            End With

440         FIsHidden = True

        End If

450     KRITHRIA.SHOW 1
460     Fsql = MDIForm1.StatusBar1.Panels(6).Text
470     MDIForm1.StatusBar1.Panels(6).Text = ""
480     Text1.Text = Fsql

        '<EhFooter>
        Exit Sub

cmdKRITHRIA_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub File1_Click()

        '<EhHeader>
        On Error GoTo File1_Click_Err

        '</EhHeader>
        Dim a

100     a = 0

        Dim F As String

110     F = File1.FILENAME

        Dim SS As String

        Dim B  As String

120     SS = ""
130     Open File1.Path + "\" + F For Input As #1

140     Do While Not EOF(1)

150         Line Input #1, B
160         SS = SS + B + " "
        Loop

170     Close #1

180     rCheck.Value = vbChecked

190     Text2.Text = SS

        '<EhFooter>
        Exit Sub

File1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.File1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.File1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_DblClick()
    PARAMETROI.PARAM.Caption = "PELAT3"
    PARAMETROI.SHOW 1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>

100     If g_Stop = 1 Then  'entos loop
110         g_Stop = 2  'ektos loop

            Exit Sub

        End If

120     If KeyCode = 27 Then EXODOS_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     mForm_Load Me, fh, fw, ft, fl
110     Fsql = ""
120     Me.Picture = LoadPicture(gPicture)
130     COMBO_PELPROM.Text = COMBO_PELPROM.List(0)
140     Combo2.Text = Combo2.List(1)
150     F_RUN = 0
160     FIsHidden = False
170     Me.KeyPreview = True


F_MONOENERGOS = Val(FINDPARAMETROI(1, "PAR1", "F_MONOENERGOS", "0", "ΔΕΙΧΝΕΙ ΜΟΝΟ ΤΟΥΣ ΕΝΕΡΓΟΥΣ=1 ΟΛΟΥΣ=0")) '

180     f_PALIAXRONIA = Trim(FINDPARAMETROI(1, "PAR1", "F_PALIAXRONIA", "2005", "Γ.18 DSN ΠΑΛΙΑΣ ΧΡΟΝΙΑΣ"))

        'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

        F_PLATOS_FORMAS = FINDPARAMETROI(1, "PELAT3", "F_PLATOS_FORMAS", "15625", "Πλάτος φόρμας")

        'ειανι η καλυτερη διόρθωση με το δεξιό άκρο του δεξιότερου control (εξοδος =command2)
        If F_PLATOS_FORMAS > 0 Then F_PLATOS_FORMAS = EXODOS.Left + EXODOS.Width + 100

190     F_LHM1 = FINDPARAMETROI(1, "PELAT2", "F_lhm1", "Ημερ1", "Ετικέτα ημερ.1")
200     F_LHM2 = FINDPARAMETROI(1, "PELAT2", "F_lhm2", "Ημερ2", "Ετικέτα ημερ.2")
210     f_lHM3 = FINDPARAMETROI(1, "PELAT2", "F_lhm3", "Ημερ3", "Ετικέτα ημερ.3")
220     f_lHM4 = FINDPARAMETROI(1, "PELAT2", "F_lhm4", "Ημερ4", "Ετικέτα ημερ.4")
230     f_lHM5 = FINDPARAMETROI(1, "PELAT2", "F_lhm5", "Ημερ5", "Ετικέτα ημερ.5")
240     f_lHM6 = FINDPARAMETROI(1, "PELAT2", "F_lhm6", "Ημερ6", "Ετικέτα ημερ.6")

250     f_lab1 = FINDPARAMETROI(1, "PELAT2", "F_lab1", "Κειμ1", "Ετικέτα κειμ.1")
260     f_lab2 = FINDPARAMETROI(1, "PELAT2", "F_lab2", "Κειμ2", "Ετικέτα κειμ.2")
270     f_lab3 = FINDPARAMETROI(1, "PELAT2", "F_lab3", "Κειμ3", "Ετικέτα κειμ.3")
280     f_lab4 = FINDPARAMETROI(1, "PELAT2", "F_lab4", "Κειμ4", "Ετικέτα κειμ.4")
290     f_lab5 = FINDPARAMETROI(1, "PELAT2", "F_lab5", "Κειμ5", "Ετικέτα κειμ.5")
300     f_lab6 = FINDPARAMETROI(1, "PELAT2", "F_lab6", "Κειμ6", "Ετικέτα κειμ.6")

310     f_num1 = FINDPARAMETROI(1, "PELAT2", "F_num1", "Αριθ1", "Ετικέτα κειμ.1")
320     f_num2 = FINDPARAMETROI(1, "PELAT2", "F_num2", "Αριθ2", "Ετικέτα κειμ.2")
330     f_num3 = FINDPARAMETROI(1, "PELAT2", "F_num3", "Αριθ3", "Ετικέτα κειμ.3")
340     f_num4 = FINDPARAMETROI(1, "PELAT2", "F_num4", "Αριθ4", "Ετικέτα κειμ.4")


        f_CO1 = FINDPARAMETROI(52, "PELAT2", "f_lbcb1", "", "Επιλογές π.χ. Εμπορος;Ιδιώτης;Υπάλληλος")
        f_CO2 = FINDPARAMETROI(54, "PELAT2", "f_lbcb2", "", "Επιλογές π.χ. Εμπορος;Ιδιώτης;Υπάλληλος")
        f_CO3 = FINDPARAMETROI(56, "PELAT2", "f_lbcb3", "", "Επιλογές π.χ. Εμπορος;Ιδιώτης;Υπάλληλος")
        f_CO4 = FINDPARAMETROI(58, "PELAT2", "f_lbcb4", "", "Επιλογές π.χ. Εμπορος;Ιδιώτης;Υπάλληλος")
        f_CO5 = FINDPARAMETROI(60, "PELAT2", "f_lbcb5", "", "Επιλογές π.χ. Εμπορος;Ιδιώτης;Υπάλληλος")






        On Error Resume Next

350     If Len(Dir("c:\mercvb\queries\pel", vbDirectory)) = 0 Then
360         MkDir "c:\mercvb\queries\pel"
        End If

370     File1.Path = "c:\mercvb\queries\pel"

380     cd1.InitDir = "c:\mercvb\queries\pel"

390     apo.Value = gEnarjh
400     eos.Value = gLhjh

        'ΣΤΗΛΕΣ ΠΙΝΑΚΑ
        ' 0=sql string  1=sum  2=alignment  3=width  4=listbox

        Dim R As Integer

410     R = 0: f_list1(R, 0) = ",DIE AS [ΔΙΕΥΘΥΝΣΗ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "ΔΙΕΥΘΥΝΣΗ"

         R = R + 1: f_list1(R, 0) = ",POL AS [ΠΟΛΗ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "ΠΟΛΗ"
        
        


420     R = R + 1: f_list1(R, 0) = ",ROUND(XRESYN,2) AS [ΧΡ_ΕΤΟΥΣ]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "XΡ.ΕΤΟΥΣ"
430     R = R + 1: f_list1(R, 0) = ",ROUND(PISSYN,2) AS [ΠΙΣ_ΕΤΟΥΣ]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "ΠΙΣ.ΕΤΟΥΣ"

440     R = R + 1: f_list1(R, 0) = ",EPA AS [ΕΠΑΓΓΕΛΜΑ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "ΕΠΑΓΓΕΛΜΑ"

450     R = R + 1: f_list1(R, 0) = ",AFM AS [Α.Φ.Μ.]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1300: f_list1(R, 4) = "Α.Φ.Μ."

460     R = R + 1: f_list1(R, 0) = ",ROUND(TYP,2) AS [YΠΟΛΟΙΠΟ]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "ΥΠΟΛΟΙΠΟ"

470     R = R + 1: f_list1(R, 0) = ",THL AS [ΤΗΛΕΦΩΝΟ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "THΛΕΦΩΝΟ"
        
         R = R + 1: f_list1(R, 0) = ",KINHTO AS [KINHTO]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "KINHTO"
         R = R + 1: f_list1(R, 0) = ",EMAIL AS [EMAIL]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 3000: f_list1(R, 4) = "EMAIL"

480     R = R + 1: f_list1(R, 0) = ",PEK AS [EKΠΤΩΣΗ]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "ΕΚΠΤΩΣΗ"
490     R = R + 1: f_list1(R, 0) = ",PLAFON AS [ΠΛΑΦΟΝ]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "ΠΛΑΦΟΝ"

500     R = R + 1: f_list1(R, 0) = ",TYPOS AS [KAΘ.ΦΠΑ]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "KAΘ.ΦΠΑ"
510     R = R + 1: f_list1(R, 0) = ",PLAISIO AS [Μεταφορική]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Μεταφορική"
520     R = R + 1: f_list1(R, 0) = ",ARPARAG AS [Τρόπος Πληρωμής]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Τρόπος Πληρωμής"
530     R = R + 1: f_list1(R, 0) = ",KODGAL AS [Κατηγορία]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Κατηγορία"

540     R = R + 1: f_list1(R, 0) = ",PVLHTHS AS [Πωλητής]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = "Πωλητής"

550     R = R + 1: f_list1(R, 0) = ",ROUND(AYP,2) AS [Υπόλ.Προηγ.Χρον]": f_list1(R, 1) = 1: f_list1(R, 2) = 6: f_list1(R, 3) = 1300: f_list1(R, 4) = "ΥΠΟΛ.ΠΡΟΗΓ.ΧΡΟΝ"

560     R = R + 1: f_list1(R, 0) = ", CONVERT(CHAR(12),HM1,3) AS [" + F_LHM1 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = F_LHM1

570     R = R + 1: f_list1(R, 0) = ", CONVERT(CHAR(12),HM2,3) AS [" + F_LHM2 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = F_LHM2

580     R = R + 1: f_list1(R, 0) = ", CONVERT(CHAR(12),HM3,3) AS [" + f_lHM3 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lHM3

590     R = R + 1: f_list1(R, 0) = ", CONVERT(CHAR(12),HM4,3) AS [" + f_lHM4 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lHM4

600     R = R + 1: f_list1(R, 0) = ", CH1 AS [" + f_lab1 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab1
610     R = R + 1: f_list1(R, 0) = ", CH2 AS [" + f_lab2 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab2
620     R = R + 1: f_list1(R, 0) = ", CH3 AS [" + f_lab3 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab3
630     R = R + 1: f_list1(R, 0) = ", CH4 AS [" + f_lab4 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_lab4

640     R = R + 1: f_list1(R, 0) = ", NUM1 AS [" + f_num1 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_num1
650     R = R + 1: f_list1(R, 0) = ", NUM2 AS [" + f_num2 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_num2
660     R = R + 1: f_list1(R, 0) = ", NUM3 AS [" + f_num3 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_num3
670     R = R + 1: f_list1(R, 0) = ", NUM4 AS [" + f_num4 + "]": f_list1(R, 1) = 0: f_list1(R, 2) = 1: f_list1(R, 3) = 1600: f_list1(R, 4) = f_num4

        Dim k As Integer

680     For k = 0 To UBound(f_list1)

690         If Len(f_list1(k, 4)) > 0 Then
700             List1.AddItem f_list1(k, 4) + Space(180) + Format(k, "00")
            End If

        Next

710     List1.ToolTipText = "Με click επιλέγετε το πεδία" + " Με σύρσιμο μπορείτε να αλλάξετε την σειρά"

720     ' DTPicker1.Value = Now

730     TDBGrid.Splits(0).ExtendRightColumn = True
740     TDBGrid.ColumnFooters = True
750     TDBGrid.Splits(0).AlternatingRowStyle = True
760     TDBGrid.EmptyRows = True
770     TDBGrid.FilterBar = True

780     Me.KeyPreview = True

        On Error Resume Next

        'File1.Path = "C:\mercvb\queries"

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>
    Dim m As Long

    m = F_PLATOS_FORMAS
    
    If F_PLATOS_FORMAS < 0 Then
        ResizeFormFor Me
    Else
        '       If F_PLATOS_FORMAS > Screen.Width Then
        '          m = Screen.Width
        '       Else
        '          m = F_PLATOS_FORMAS
        '       End If
        '       'εαν το κουμπι εξόδου βγαίνει εξω να το μαζευει
        '       If Command2.Left + Command2.Width + 100 > m Then
        '          m = Command2.Left + Command2.Width + 100
        '       End If
        '
       
        mForm_Resize Me, 8, m, 10455, fh, fw, ft, fl
    End If
    
    'If F_PLATOS_FORMAS < 0 Then
    '   ResizeFormFor Me
    'Else
    '   mForm_Resize Me, 8, F_PLATOS_FORMAS, 10455, fh, fw, ft, fl
    'End If
End Sub

Private Sub TDBGrid_DblClick()

        'pelat4.Option1(1).value = IIf(Left(Combo1.Text, 1) = "e", -1, 0) ' Option1(1).value
        '<EhHeader>
        On Error GoTo TDBGrid_DblClick_Err

        '</EhHeader>

100     pelat4.PELPROM.ListIndex = IIf(Left(COMBO_PELPROM.Text, 1) = "e", 0, 1)    ' Option1(1).value
110     pelat4.SHOW
120     pelat4.APOALLO = "1"
130     pelat4.data1.ConnectionString = gConnect
140     TDBGrid.Col = 1
150     pelat4.Text2(0).Text = TDBGrid.Text
160     SendKeys "~"    'ENTER
170     SendKeys "~"    'enter  {DOWN}"
        ' SendKeys "~" 'ENTER

        '<EhFooter>
        Exit Sub

TDBGrid_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.TDBGrid_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.TDBGrid_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub TDBGrid_FootClick(ByVal ColIndex As Integer)

        '<EhHeader>
        On Error GoTo TDBGrid_FootClick_Err

        '</EhHeader>

        Dim fff As New pelat2

        Dim DD

100     DD = IIf(Left(COMBO_PELPROM.Text, 1) = "e", 0, 1)

110     fff.PELPROM.ListIndex = DD

120     TDBGrid.Col = 1
130     fff.ffbuff = Left(COMBO_PELPROM.Text, 1) + TDBGrid.Text  ' Data1.Recordset("kod") ' Text1(0).Text

140   '  fff.Command5.Enabled = False
150     gApoMenu = False

160     pelat2.SHOW

        Exit Sub

        '<EhFooter>
        Exit Sub

TDBGrid_FootClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.TDBGrid_FootClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.TDBGrid_FootClick " & "at line " & Erl

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

160     data1.Recordset.Sort = strSort

        '<EhFooter>
        Exit Sub

TDBGrid_GroupColMove_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.TDBGrid_GroupColMove " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.TDBGrid_GroupColMove " & "at line " & Erl

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

        Dim k          As Long

        On Error Resume Next

        'edo parakato exei latos 15/8/2009

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

        '    Adodc2.Recordset.Sort = "[" & TDBGrid.Columns(ColIndex).DataField & "] desc" ' strSort

        '<EhFooter>
        Exit Sub

TDBGrid_HeadClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.TDBGrid_HeadClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.TDBGrid_HeadClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub GR2_Click()

        '<EhHeader>
        On Error GoTo GR2_Click_Err

        '</EhHeader>
        Dim a As Long

100     a = GR2.row
110     show_kartella

        '<EhFooter>
        Exit Sub

GR2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.GR2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.GR2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub List1_DblClick()

    '<EhHeader>
    On Error GoTo List1_DblClick_Err

    '</EhHeader>
    Dim k  As Integer

    Dim nc As Integer

    nc = 0

80  For k = 0 To List1.ListCount - 1

90      If List1.Selected(k) = True Then
            nc = nc + 1
        End If

    Next

    If nc >= List1.ListCount - 1 Then

        For k = 0 To List1.ListCount - 1
            List1.Selected(k) = False
        Next

    Else

        For k = 0 To List1.ListCount - 1
            List1.Selected(k) = True
        Next

    End If
        
    '<EhFooter>
    Exit Sub

List1_DblClick_Err:
    'MsgBox Err.Description & vbCrLf & _
     "in ADOMERCNEW.pelat3.List1_DblClick " & _
     "at line " & Erl, _
     vbExclamation + vbOKOnly, "Application Error"
    SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.List1_DblClick " & "at line " & Erl

    Resume Next

    '</EhFooter>

End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

        '<EhHeader>
        On Error GoTo List1_MouseDown_Err

        '</EhHeader>

100     fList1ARX = List1.ListIndex

        '<EhFooter>
        Exit Sub

List1_MouseDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.List1_MouseDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.List1_MouseDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

        '<EhHeader>
        On Error GoTo List1_MouseUp_Err

        '</EhHeader>
        Dim TEL As Integer, ARX As Integer

        Dim s   As String

100     TEL = List1.ListIndex
110     ARX = fList1ARX

        ' Me.Caption = List1.ListIndex
120     If ARX = TEL Then  'ΠΑΡΕΜΕΙΝΕ ΣΤΟ ΙΔΙΟ ΚΕΛΙ
            'ΑΣΤΟ
        Else

130         If TEL > ARX Then
140             s = List1.List(ARX)
150             List1.RemoveItem ARX

160             List1.AddItem s, TEL
170             List1.Selected(TEL) = True

            Else
180             List1.AddItem List1.List(ARX), TEL
190             List1.Selected(TEL) = True
200             List1.RemoveItem ARX + 1
            End If

        End If

        '<EhFooter>
        Exit Sub

List1_MouseUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.pelat3.List1_MouseUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.pelat3.List1_MouseUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
