VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form DataEntry2 
   BackColor       =   &H00FF0000&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   6435
   ClientLeft      =   3090
   ClientTop       =   5355
   ClientWidth     =   11100
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   11100
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10500
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":0CF6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox mactext 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10470
      TabIndex        =   3
      Top             =   585
      Visible         =   0   'False
      Width           =   720
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   165
      Top             =   5640
      Visible         =   0   'False
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   635
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid ParamGrid 
      Height          =   1905
      Left            =   6420
      TabIndex        =   2
      Top             =   6375
      Visible         =   0   'False
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   3360
      _Version        =   393216
      Rows            =   20
      Cols            =   20
      _NumberOfBands  =   1
      _Band(0).Cols   =   20
   End
   Begin VB.TextBox Textboxes 
      BackColor       =   &H00C0C0FF&
      Height          =   285
      Index           =   0
      Left            =   2310
      TabIndex        =   1
      Top             =   435
      Visible         =   0   'False
      Width           =   2235
   End
   Begin TrueOleDBGrid80.TDBGrid Grid 
      Bindings        =   "dataentry3.frx":1148
      Height          =   3225
      Left            =   -75
      TabIndex        =   4
      Top             =   795
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   5689
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
      Splits(0).ShowCollapseExpandIcons=   0   'False
      Splits(0).MarqueeStyle=   4
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=260"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=260"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
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
      GroupByCaption  =   ""
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   9329.953
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
      _StyleDefs(24)  =   "Splits(0).Style:id=55,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=56,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=57,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=58,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=60,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=59,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=61,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=62,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=75,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=82,.parent=55"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=56,.alignment=0"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=57"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=59"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=86,.parent=55"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=56,.alignment=0"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=57"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=59"
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
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(55)  =   ":id=37,.strikethrough=0,.charset=161"
      _StyleDefs(56)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(57)  =   "Named:id=38:HighlightRow"
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=39:EvenRow"
      _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(61)  =   "Named:id=40:OddRow"
      _StyleDefs(62)  =   ":id=40,.parent=33"
      _StyleDefs(63)  =   "Named:id=41:RecordSelector"
      _StyleDefs(64)  =   ":id=41,.parent=34"
      _StyleDefs(65)  =   "Named:id=42:FilterBar"
      _StyleDefs(66)  =   ":id=42,.parent=33"
   End
   Begin MSComctlLib.ImageList imlToolbarPics 
      Left            =   10455
      Top             =   4470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483634
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   128
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":115D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":126F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":16C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":1B13
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":1F65
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":23B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":2809
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":2C5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":30AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":31BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":32D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":33E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":34F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":3607
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":3719
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":382B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "dataentry3.frx":393D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      HelpContextID   =   65278
      Style           =   1
      ImageList       =   "imlToolbarPics"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Table"
            Object.ToolTipText     =   "Νέα Εγγραφή"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Dynaset"
            Object.ToolTipText     =   "Αποθήκευση"
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Schedule"
            Object.ToolTipText     =   "Διαγραφή"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BIBLIO"
            Object.ToolTipText     =   "Εξοδος"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Label Labels 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   435
      Visible         =   0   'False
      Width           =   2235
   End
End
Attribute VB_Name = "DataEntry2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
 
 Const VK_TAB = &H9
 Dim I As Integer
 Dim Captions(50) As String
 Dim mfields(50) As String
 Dim Sizes(50) As Integer
 Dim sql As String
 Dim Sql2 As String
 Dim delSql As String
 Dim UpdSQL As String
 
 Dim F_NEW As Integer
Dim f_AddNew2

 Dim defSQL As String 'default string που προσθέτω
Dim n As Integer ' αριθμός πεδίων
 Dim fSCR As Object ' SCRIPT CONTROL
 Dim f_Typos1Field As Long  ' τι τυπος είναι το 1ο πεδίο
Dim fAddNewScript
 
Dim F_FIELD_EPIPLEON As String, F_TIMH
Dim F_2FIELD_EPIPLEON As String, F_2TIMH
 
 
 
 
 Dim flagProthFora
 
'If Index = 2 Then 'ΜΟΝΑΔΕΣ ΜΕΤΡΗΣΗΣ
'  'sql οπου θα γίνεται το update ή το addnew
' '  f.SHOW
'  Load f
'
'  Gdb.Execute "update PINAKES SET AYJON =0 WHERE AYJON IS NULL AND TYPOS=2"
'  f.ParamGrid.TextMatrix(0, 1) = "SELECT * FROM PINAKES WHERE TYPOS=2 " '    sql
''τα πεδία που θα μπαίνουν sto ParamGrid και στα textboxes
'f.ParamGrid.TextMatrix(0, 2) = "SELECT AYJON,PERIGRAFH " _
'& " FROM PINAKES WHERE TYPOS=2 ORDER BY AYJON " 'ParamGrid   sql2
'
''με ποιά εντολή θα κάνει το delete (το where ειναι το kod=...
'f.ParamGrid.TextMatrix(0, 3) = "delete from PINAKES where TYPOS=2 AND " 'ParamGrid              DelSQL
'
''ποιό update query θα δημιουργείται (προαιρετικά )
''αν δεν θέλω update τότε βάζω f.ParamGrid.TextMatrix(0, 4) = ""
''                    f.ParamGrid.TextMatrix(0, 4) = "  ""UPDATE ARITMISI2 SET ARITMISI= "" + t12 +"" where ID=""+t11 "   'UpdSQL
' DD = " UPDATE PINAKES SET TYPOS=2 where TYPOS IS NULL AND AYJON="" + t1   "
'' AddNewScript
'  DD = "Sub Main()" & vbCrLf & _
'         " DIM mCon  " & vbCrLf & _
'         " Set mCon=CreateObject(""ADODB.Connection"")" & vbCrLf & _
'         " mCon.Open """ + gConnect + " "" " & vbCrLf & _
'         " mCon.execute """ + DD + " " & vbCrLf & _
'        "Set r=Nothing" & vbCrLf & _
'        "Set mCon=Nothing" & vbCrLf & _
'        "End Sub"
'f.ParamGrid.TextMatrix(0, 4) = DD
'f.ParamGrid.TextMatrix(1, 1) = "A/A"
'f.ParamGrid.TextMatrix(2, 1) = "ΤΙΤΛΟΣ "
''ΜΕΓΑΛΩΝΩ ΛΙΓΟ ΤΑ LABELS
'f.Labels(0).width = f.Labels(0).width + 1000 '* TextWidth("A")
'f.Textboxes(0).Left = f.Textboxes(0).Left + 1000 ' * TextWidth("A")
'f.Left = Label1(Index).Left
'f.Top = Label1(Index).Top + MDIForm1.Top
'
'
'
'f.SHOW 1
'Combo1.Clear
' R.Open "SELECT *FROM PINAKES WHERE TYPOS=2 ORDER BY AYJON", Gdb, adOpenDynamic, _
'     adLockOptimistic
'Do While Not R.EOF
'  If R("typos") = 2 Then
'     Combo1.AddItem R("PERIGRAFH")
'  Else
'     Exit Do
'  End If
'  R.MoveNext
'Loop
'R.Close
'
'
'
'
'
'
'End If
'


Sub PaintForm()

Dim I As Integer


' n = Val(ParamGrid.TextMatrix(0, 0)) 'ο αριθμός των σειρών

sql = ParamGrid.TextMatrix(0, 1)
Sql2 = ParamGrid.TextMatrix(0, 2) ' string adodc1
delSql = ParamGrid.TextMatrix(0, 3) ' string adodc1

UpdSQL = ParamGrid.TextMatrix(0, 4) ' string adodc1

fAddNewScript = ParamGrid.TextMatrix(0, 5) ' string meta to addnew button

F_FIELD_EPIPLEON = ParamGrid.TextMatrix(0, 6) 'EPIPLEON FIELD
F_TIMH = ParamGrid.TextMatrix(0, 7)  ' TIMH TOY EPIPLEON FIELD


F_2FIELD_EPIPLEON = ParamGrid.TextMatrix(0, 8) 'EPIPLEON FIELD
F_2TIMH = ParamGrid.TextMatrix(0, 9)  ' TIMH TOY EPIPLEON FIELD



f_AddNew2 = ParamGrid.TextMatrix(0, 10)  ' addnew ap'ευθείας


'f.ParamGrid.TextMatrix(0, 8) = "TIMH" '2 EPIPLEON FIELD
'f.ParamGrid.TextMatrix(0, 9) = Left(Combo4.Text, 2)  ' TIMH TOY 2 EPIPLEON FIELD




Dim R As New ADODB.Recordset
  
   R.Open Sql2, Gdb, adOpenKeyset, adLockOptimistic

n = R.FIELDS.Count


For I = 1 To n
   ' Captions(i + 1) = r.Fields(i).Name 'ονομα ετικέτας
   Captions(I) = ParamGrid.TextMatrix(I, 1)  'ονομα ετικέτας
   If R.FIELDS(I - 1).DefinedSize > 60 Then
        Sizes(I) = 30 '  R.FIELDS(I - 1).ActualSize 'definedsize  ' φάρδος textbox
   Else
        Sizes(I) = R.FIELDS(I - 1).DefinedSize  ' ActualSize 'definedsize  ' φάρδος textbox
   End If
 

   mfields(I) = R.FIELDS(I - 1).Name ' field textbox
   
Next

'το χρειαζομαι για να συντάξω το sql αναζητησης
f_Typos1Field = R.FIELDS(0).Type  '3 integer  135=date  202=string



 For I = 1 To n
      Load Labels(I)
      Load TextBoxes(I)
      Labels(I).Top = Labels(I - 1).Top + 1.2 * Labels(0).height
      Labels(I).Left = Labels(0).Left
      TextBoxes(I).Top = TextBoxes(I - 1).Top + 1.2 * TextBoxes(I).height
      TextBoxes(I).Left = TextBoxes(0).Left
      TextBoxes(I).width = (Sizes(I) + 1) * TextWidth("A")
      TextBoxes(I).MaxLength = Sizes(I)
      Labels(I).Caption = Captions(I)
     ' Labels(I).Visible = True
     ' Textboxes(I).Visible = True
   Next


tlbToolBar.Buttons(3).Enabled = False


' ParamGrid.HEIGHT =7905  MAXIMUM


 End Sub

Sub paintgrid()

Dim I As Integer
Dim L As Long
L = 0

  Adodc1.ConnectionString = gConnect
  Adodc1.RecordSource = Sql2
  Adodc1.Refresh
For I = 1 To Adodc1.Recordset.FIELDS.Count
   
   'grid.ColWidth(i - 1) = (Sizes(i) + 1) * TextWidth("A")    'flexgrid
   'grid.TextMatrix(0, i - 1) = ParamGrid.TextMatrix(i, 1)    'flexgrid
   ' L = L + grid.ColWidth(i - 1)                             'flexgrid
     ' ΓΙΑ ΝΑ ΧΩΡΑΕΙ Ο ΤΙΤΛΟΣ
   If Sizes(I) < Len(Captions(I)) Then
      Grid.Columns(I - 1).width = Len(Captions(I)) * TextWidth("A")
   Else
      Grid.Columns(I - 1).width = (Sizes(I) + 1) * TextWidth("A")
   End If
   
   Grid.Columns(I - 1).Caption = ParamGrid.TextMatrix(I, 1)
   L = L + Grid.Columns(I - 1).width
Next
 
  Grid.width = Me.width
  Grid.height = Me.height - Grid.Top
  
 
 'If L + 5 * TextWidth("A") > Me.Width Then
 '   Grid.Width = Me.Width ' - 3 * TextWidth("A")
 'Else
   ' Grid.Width = L + 5 * TextWidth("A")
 '    Grid.Width = Me.Width
 
' End If

End Sub






Private Sub BUT_Click(Index As Integer)
  Dim ANS As Integer
  Dim R As New ADODB.Recordset
  Dim I As Integer
   Dim CurRow As Long
  Dim mkod As String
'-----------------------------------------------------------------------------------------------------------------------
If Index = 0 Then ' προσθετω εγγραφη
   F_NEW = 1
   For I = 1 To n
      TextBoxes(I).Text = ""
   Next
   
 '  BUT(0).Enabled = False
 '  BUT(1).Enabled = True
   
   
   If Len(fAddNewScript) > 0 Then
      trexe fAddNewScript
   End If

   For I = 1 To n
      
      ' Textboxes(I).Text = "              "
      Labels(I).Visible = True
      TextBoxes(I).Visible = True
   Next
   Grid.Top = Labels(n).Top + 1.2 * Labels(n).height
    TextBoxes(1).SetFocus
End If
'----------------------------------------------------------------------------------------------------------------
If Index = 1 Then 'ΑΠΟΘΗΚΕΥΣΗ
  mkod = TextBoxes(1).Text  ' grid.Columns(0).Text
'If f_Typos1Field = adInteger Or f_Typos1Field = adSingle Or f_Typos1Field = adDouble Then   ' integer
'      R.Open sql + " and " + mfields(1) + "= " + mkod + " ", gdb, adOpenKeyset, adLockOptimistic
'Else
'      R.Open sql + " and " + mfields(1) + "='" + mkod + "'", gdb, adOpenKeyset, adLockOptimistic
'End If
  If F_NEW = 1 Then
    ' R.AddNew
     trexe f_AddNew2
     F_NEW = 0
  Else
     trexe UpdSQL
  End If
 
 On Error GoTo LatosTypoi
   For I = 1 To n
      Labels(I).Visible = False
      TextBoxes(I).Visible = False
   Next
   
  Grid.Top = tlbToolBar.height + TextHeight("A")
  Grid.Visible = True
  Adodc1.Refresh
  paintgrid
  'BUT(0).Enabled = True ' add
  'BUT(1).Enabled = False 'save
  'BUT(2).Enabled = False 'DELETE
  'BUT(0).SetFocus
  
  Exit Sub
End If
  
'-------------------------------------------------------------------------------------------------------------------
If Index = 2 Then 'διαγραφη
  ANS = MsgBox("Είστε σίγουροι για την διαγραφή;", vbYesNo)
  If ANS = vbYes Then
     CurRow = Grid.row '+ 1
     ' MKOD = grid.TextMatrix(CURROW, 0) ' kodikos   'flexgrid
     mkod = Grid.Columns(0).Text
     Gdb.Execute delSql + "  " + mfields(1) + "='" + mkod + "'"
     Grid.Visible = True
     Adodc1.Refresh
    ' BUT(0).Enabled = True ' add
    ' BUT(1).Enabled = False 'save
   End If

   For I = 1 To n
      Labels(I).Visible = False
      TextBoxes(I).Visible = False
   Next




End If
If Index = 3 Then ' ΕΞΟΔΟσ
   Unload Me

End If



Exit Sub

LatosTypoi:
 MsgBox "λάθος στο πεδίο " + Labels(I)
 TextBoxes(I).SetFocus




NOSAVE:








End Sub

Private Sub Command1_Click()
   trexe "AA"
End Sub

Private Sub Form_Load()
  flagProthFora = 1
  F_NEW = 0
  
  
  
  Me.KeyPreview = True
 '  Me.Picture = LoadPicture(gPicture)
    
'    'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
'    'SkinFramework1.ApplyWindow Me.hWnd
'    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

 
  
  
  
  
  
  
  
  
  
  
  
  
  
End Sub

Private Sub Form_Paint()
Dim k As Integer

   If flagProthFora = 1 Then
      
   
   
   
      PaintForm
      paintgrid
      
      Set fSCR = CreateObject("MSScriptControl.ScriptControl")
      fSCR.language = "vbscript"
      fSCR.AddObject "mactext", MACText
      For k = 1 To n
         fSCR.AddObject "t" + Format(k, "##"), TextBoxes(k)
      Next
      
      flagProthFora = 0
   End If
   
End Sub

Private Sub Form_Resize()
  Grid.width = Me.width
  Grid.Top = tlbToolBar.height
  Grid.height = Me.height - Grid.Top - 390
  
End Sub

Private Sub grid_Click()
 '   grid_DblClick
End Sub

Private Sub grid_DblClick()
Dim CurRow As Long
Dim mkod As String

CurRow = Grid.row '+ 1
'MKOD = grid.TextMatrix(CURROW, 0) ' kodikos  'flexgrid

'ΑΝ ΔΕΝ ΕΧΕΙ ΕΓΓΡΑΦΕΣ ΝΑ ΜΗΝ ΠΡΟΧΩΡΑΕΙ
If Adodc1.Recordset.EOF Then Exit Sub

mkod = Grid.Columns(0).Text



  Dim R As New ADODB.Recordset
  Dim I As Integer
  ' "SELECT KOD,EPO,AYP FROM PEL0" 'ParamGrid

'  ' r.Open Sql + " and " + mFields(1) + "='" + MKOD + "'", gdb, adOpenKeyset, adLockOptimistic
'
'If f_Typos1Field = 3 Then ' integer
'    If mkod = "" Then
'      R.Open sql + " and " + mfields(1) + " is null ", gdb, adOpenKeyset, adLockOptimistic
'    Else
'      R.Open sql + " and " + mfields(1) + "= " + mkod + " ", gdb, adOpenKeyset, adLockOptimistic
'    End If
'Else
'   If mkod = "" Then
'      R.Open sql + " and " + mfields(1) + " is null ", gdb, adOpenKeyset, adLockOptimistic
'    Else
''      R.Open sql + " and " + mfields(1) + "='" + mkod + "'", gdb, adOpenKeyset, adLockOptimistic
'    End If
'     ' R.Open sql + " and " + mfields(1) + "='" + mkod + "'", Gdb, adOpenKeyset, adLockOptimistic
'End If


   For I = 1 To n
      Grid.Col = I - 1: TextBoxes(I).Text = Grid.Text ' R(mfields(I))
      Labels(I).Visible = True
      TextBoxes(I).Visible = True
   Next

Grid.Visible = False

tlbToolBar.Buttons(1).Enabled = False 'add
tlbToolBar.Buttons(2).Enabled = True ' update


TextBoxes(1).SetFocus






End Sub

Private Sub grid_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    grid_DblClick
  End If
  
    
End Sub

Private Sub grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  ' grid_DblClick
End Sub

Private Sub Textboxes_GotFocus(Index As Integer)
   TextBoxes(Index).SelStart = 0
   TextBoxes(Index).SelLength = Len(TextBoxes(Index).Text)
   
   TextBoxes(Index).BackColor = vbYellow
   
End Sub

Private Sub Textboxes_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       KeyAscii = 0 'suppress the beep
       keybd_event VK_TAB, 0, 0, 0 'send a tab
    End If
End Sub



Private Sub Textboxes_LostFocus(Index As Integer)
Dim X
    X = ParamGrid.TextMatrix(Index, 2)
   
   If Len(X) > 0 Then
       trexe X
   End If
   TextBoxes(Index).BackColor = TextBoxes(0).BackColor
   
   
End Sub

Function mac(ByVal S As String) As String
'δινω ενα στρινγκ και μου επιστρέφει μία τιμή
Dim X
On Error GoTo 0
MACText.Text = "   "
On Error GoTo err2
X = "MACtEXT.Text = " + S
fSCR.ExecuteStatement X

'If F928 = 1 Then
    mac = MACText.Text
'Else
'    mac = to437(mactext.Text)
'End If

Exit Function

err2:
Me.Caption = X
mac = "  "
Resume Next
End Function




Private Sub trexe(ByVal Scode As String)
' gconnect = "DSN=MERC2"
  'scode = "Sub Main()" & vbCrLf & _
         "Msgbox 15 " & vbCrLf & _
         " DIM mCon  " & vbCrLf & _
         " Set mCon=CreateObject(""ADODB.Connection"")" & vbCrLf & _
         " mCon.Open """ + gconnect + """ " & vbCrLf & _
         " mCon.execute ""UPDATE ARITMISI2 SET ARITMISI=457  WHERE ID= 7 "" " & vbCrLf & _
        "DIM R " & vbCrLf & _
        " Set R=CreateObject(""ADODB.Recordset"")" & vbCrLf & _
        " r.Open "" SELECT ARITMISI FROM ARITMISI2 WHERE ID=7 "",mCon " & vbCrLf & _
         "Msgbox r(0) " & vbCrLf & _
         "Set r=Nothing" & vbCrLf & _
        "Set mCon=Nothing" & vbCrLf & _
        "End Sub"
    Call fSCR.AddCode(Scode)
    Call fSCR.Run("Main")
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub tlbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)

  Dim ANS As Integer
  Dim R As New ADODB.Recordset
  Dim I As Integer
   Dim CurRow As Long
  Dim mkod As String
'-----------------------------------------------------------------------------------------------------------------------
If Button.Index = 1 Then ' προσθετω εγγραφη
   F_NEW = 1
   For I = 1 To n
      TextBoxes(I).Text = ""
   Next
   
'   BUT(0).Enabled = False
'   BUT(1).Enabled = True
   
   
   If Len(fAddNewScript) > 0 Then
      trexe fAddNewScript
   End If

   For I = 1 To n
      
      ' Textboxes(I).Text = "              "
      Labels(I).Visible = True
      TextBoxes(I).Visible = True
   Next
   Grid.Visible = False
   Button.value = tbrPressed
   Button.Enabled = False
   tlbToolBar.Buttons(2).Enabled = True
   
   'Grid.Top = Labels(n).Top + 1.2 * Labels(n).Height
    TextBoxes(1).SetFocus
End If
'----------------------------------------------------------------------------------------------------------------
If Button.Index = 2 Then 'ΑΠΟΘΗΚΕΥΣΗ
  mkod = TextBoxes(1).Text  ' grid.Columns(0).Text
'If f_Typos1Field = adInteger Or f_Typos1Field = adSingle Or f_Typos1Field = adDouble Then   ' integer
'      R.Open sql + " and " + mfields(1) + "= " + mkod + " ", gdb, adOpenKeyset, adLockOptimistic
'Else
'      R.Open sql + " and " + mfields(1) + "='" + mkod + "'", gdb, adOpenKeyset, adLockOptimistic
'End If
  If F_NEW = 1 Then
    ' R.AddNew
     trexe f_AddNew2
     F_NEW = 0
     Button.value = tbrUnpressed
  Else
     trexe UpdSQL
     Button.value = tbrUnpressed
     
     
  End If
 tlbToolBar.Buttons(1).Enabled = True  'add
 tlbToolBar.Buttons(2).Enabled = False  'update
 On Error GoTo LatosTypoi
   For I = 1 To n
      Labels(I).Visible = False
      TextBoxes(I).Visible = False
   Next
   
  'Grid.Top = TextHeight("A")
  Grid.Visible = True
  Adodc1.Refresh
  paintgrid
'  BUT(0).Enabled = True ' add
'  BUT(1).Enabled = False 'save
'  BUT(2).Enabled = False 'DELETE
'  BUT(0).SetFocus
  
  Exit Sub
End If
  
'-------------------------------------------------------------------------------------------------------------------
If Button.Index = 3 Then 'διαγραφη
  ANS = MsgBox("Είστε σίγουροι για την διαγραφή;", vbYesNo)
  If ANS = vbYes Then
     CurRow = Grid.row '+ 1
     ' MKOD = grid.TextMatrix(CURROW, 0) ' kodikos   'flexgrid
     mkod = Grid.Columns(0).Text
     'gdb.Execute delSql + "  " + mfields(1) + "='" + mkod + "'"
     trexe delSql
     
     Grid.Visible = True
     Adodc1.Refresh
'     BUT(0).Enabled = True ' add
'     BUT(1).Enabled = False 'save
   End If

   For I = 1 To n
      Labels(I).Visible = False
      TextBoxes(I).Visible = False
   Next
   
   tlbToolBar.Buttons(1).Enabled = True  'add


End If
If Button.Index = 4 Then ' ΕΞΟΔΟσ
   Unload Me
End If



Exit Sub

LatosTypoi:
 MsgBox "λάθος στο πεδίο " + Labels(I)
 TextBoxes(I).SetFocus




NOSAVE:





























End Sub
