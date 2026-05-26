VERSION 5.00
Object = "{CCA214C0-DFEB-4C91-9F0D-2665F77F6E23}#1.2#0"; "IDAutomationLinear.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form paralabh 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   12300
   ForeColor       =   &H80000014&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   12300
   WindowState     =   2  'Maximized
   Begin VB.TextBox da 
      Height          =   285
      Left            =   8520
      TabIndex        =   26
      Top             =   7425
      Width           =   1320
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   8550
      TabIndex        =   25
      Top             =   6975
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      Format          =   69599233
      CurrentDate     =   39310
   End
   Begin VB.CommandButton daPelath 
      Caption         =   "да пекатг"
      Height          =   555
      Left            =   8115
      TabIndex        =   24
      Top             =   6270
      Width           =   1515
   End
   Begin VB.CommandButton Command3 
      Caption         =   "йАХАЯИСЛЭР"
      Height          =   375
      Left            =   4665
      TabIndex        =   23
      Top             =   2265
      Width           =   1755
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HEADER"
      Height          =   495
      Left            =   8640
      TabIndex        =   22
      Top             =   3600
      Width           =   1095
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Bindings        =   "paralabh.frx":0000
      Height          =   1545
      Left            =   7485
      TabIndex        =   20
      Top             =   915
      Width           =   4515
      _ExtentX        =   7964
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
      Splits(0).RecordSelectorWidth=   794
      Splits(0)._SavedRecordSelectors=   0   'False
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   7455
      Top             =   2520
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
   Begin VB.TextBox Text8 
      Height          =   420
      Left            =   7470
      TabIndex        =   18
      Top             =   465
      Width           =   1980
   End
   Begin VB.ComboBox SKOPOS 
      Height          =   315
      ItemData        =   "paralabh.frx":0015
      Left            =   5715
      List            =   "paralabh.frx":0028
      TabIndex        =   17
      Text            =   "еписйеуг"
      Top             =   4215
      Width           =   2745
   End
   Begin VB.ComboBox MON 
      Height          =   315
      ItemData        =   "paralabh.frx":0065
      Left            =   2865
      List            =   "paralabh.frx":0072
      TabIndex        =   16
      Text            =   "тел"
      Top             =   4170
      Width           =   1695
   End
   Begin VB.TextBox POSO 
      Height          =   300
      Left            =   1785
      TabIndex        =   14
      Text            =   "1"
      Top             =   4140
      Width           =   660
   End
   Begin VB.CommandButton Command1 
      Caption         =   "йАТАВЧЯГСГ"
      Height          =   555
      Left            =   1845
      TabIndex        =   13
      Top             =   6300
      Width           =   3750
   End
   Begin VB.TextBox EPA 
      Height          =   435
      Left            =   1800
      TabIndex        =   2
      Top             =   1680
      Width           =   5175
   End
   Begin VB.TextBox DIE 
      Height          =   435
      Left            =   1800
      TabIndex        =   1
      Top             =   1050
      Width           =   5175
   End
   Begin VB.CheckBox DATA 
      BackColor       =   &H00FF8080&
      Caption         =   "тА СТОИВЕъА ВЯЕИэФОМТАИ- мА ЛГМ СБГСТОУМ"
      ForeColor       =   &H80000014&
      Height          =   345
      Left            =   1800
      TabIndex        =   10
      Top             =   5730
      Width           =   5790
   End
   Begin VB.TextBox BLABH 
      Height          =   1020
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   4590
      Width           =   6735
   End
   Begin VB.TextBox ONO 
      Height          =   465
      Left            =   1800
      TabIndex        =   4
      Top             =   3645
      Width           =   6660
   End
   Begin VB.TextBox THL 
      Height          =   390
      Left            =   1800
      TabIndex        =   3
      Top             =   2250
      Width           =   2730
   End
   Begin VB.TextBox EPO 
      Height          =   390
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   5130
   End
   Begin VB.Label Label10 
      Caption         =   "гЛ.да"
      Height          =   330
      Left            =   6855
      TabIndex        =   28
      Top             =   6960
      Width           =   1635
   End
   Begin VB.Label Label9 
      Caption         =   "аЯИХЛЭР да"
      Height          =   315
      Left            =   6855
      TabIndex        =   27
      Top             =   7410
      Width           =   1575
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode1 
      Height          =   990
      Left            =   1365
      TabIndex        =   21
      Top             =   7155
      Width           =   1515
      _cx             =   2672
      _cy             =   1746
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
      SymbologyId     =   0
      DataToEncode    =   "123"
      Orientation     =   0
      BarHeight       =   1
      NarrowBarWidth  =   0.03
      Wide2NarrowRatio=   2
      AddCheckDigit   =   1
      AddCheckDigitToText=   1
      Code128CharSet  =   2
      UPCESystem      =   0
      EANUPCSupplement=   0
      ShowText        =   1
      CodabarStartCharacter=   "A"
      CodabarStopCharacter=   "B"
      LeftMarginCM    =   0.2
      TopMarginCM     =   0.2
      SupplementToEncode=   ""
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "паяакабг амтийеилемым циа еписйеуг"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   330
      Left            =   195
      TabIndex        =   19
      Top             =   45
      Width           =   7035
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "пос╪тгта"
      ForeColor       =   &H80000014&
      Height          =   330
      Left            =   60
      TabIndex        =   15
      Top             =   4140
      Width           =   1485
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "еПэЦЦЕКЛА"
      ForeColor       =   &H80000014&
      Height          =   330
      Left            =   60
      TabIndex        =   12
      Top             =   1725
      Width           =   1710
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "дИЕУХУМСГ"
      ForeColor       =   &H80000014&
      Height          =   330
      Left            =   60
      TabIndex        =   11
      Top             =   1095
      Width           =   1710
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "бКэБГ"
      ForeColor       =   &H80000014&
      Height          =   345
      Left            =   60
      TabIndex        =   9
      Top             =   4590
      Width           =   1725
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "пЕЯИЦЯАЖч сУСЙЕУчР"
      ForeColor       =   &H80000014&
      Height          =   405
      Left            =   60
      TabIndex        =   8
      Top             =   3660
      Width           =   1710
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "тГКщЖЫМО"
      ForeColor       =   &H80000014&
      Height          =   390
      Left            =   60
      TabIndex        =   7
      Top             =   2310
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "оМОЛАТЕПЧМУЛО"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   480
      Width           =   1680
   End
End
Attribute VB_Name = "paralabh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'дглиоуяцоуле ема RECORD сто EGGTIM KAI STO EID
' KAI тупымоуле

Dim X As String, sql As String
Dim n As Integer


Dim PR, FOUND

If Len(da.Text) > 0 Then
   MsgBox "дгкысате да пекатг. патгсте то акко пкгйтяо"
   Exit Sub
End If





FOUND = False
For Each PR In Printers
    If InStr(PR.DeviceName, "POSPARALABH") > 0 Then
        FOUND = True
        Exit For
    End If
Next

If Not FOUND Then
   MsgBox "дем упаявеи се аутгм тгм хесг о ейтупытгс циа тис паяакабес"
   Unload Me
   Exit Sub
End If


Dim PELATHS
If EPO.Text = Adodc1.Recordset("EPO") Then
  PELATHS = Adodc1.Recordset("KOD")
Else
   PELATHS = "1112"
End If

If Len(Trim(EPO.Text)) = 0 Then
  MsgBox "дЕМ СУЛПКГЯЧСАТЕ ЭМОЛА"
  Exit Sub
End If


If Len(Trim(ono.Text)) = 0 Then
  MsgBox "дЕМ СУЛПКГЯЧСАТЕ ЕъДОР"
  Exit Sub
End If





'дглиоуяциа тоу йыдийоу апохгйгс

X = FIND_NUM(0, 0) ' бяисйы то текеутаио моулеяо KAI
    'диабафы тгм пяогцоулемг глеяолгмиа циа ма лгм евы глеяолгмиа лийяотеяг тгс текеутаиас йатавыягсгс

sql = "INSERT INTO EID (KOD,ONO,FPA,MEMO,MON,CH1,NUM1,CH2,CH3,CH4,CH5,CH6,HM2) VALUES (" _
& "'" + X + " ','" + Replace(ono.Text, "'", "-") + "',2,'" + Replace(BLABH.Text, "'", "-") + "'," _
& "'" + MON.Text + "','" + SKOPOS.Text + "'," + POSO.Text + "," _
& "'" + Left(EPO.Text, 20) + "','" + Left(DIE.Text, 20) + "','" + EPA.Text + "','" + THL.Text + "','" + PELATHS + "'," _
& "GETDATE())"

Gdb.Execute sql, n







If n = 0 Then
  MsgBox " ДЕМ ЙАТАВЫЯчХГЙЕ "
  Gdb.Execute "update MEM SET EPEL=EPEL-1 WHERE RPEL=" + rPEL, k
  
  Exit Sub
End If


Dim DUM As String

'DUM = FIND_NUM(1, 0) ' ЙАТАВЫЯЧ ТО ЕПЭЛЕМО МОЩЛЕЯО


'дглиоуяциа тгс йимгсгс апохгйгс (еGGTIM)


sql = "INSERT INTO EGGTIM (HME,KODE,ONOMA,FPA,EIDOS,APOT,POSO,XRE,PELKOD,ATIM) VALUES (" _
& "'" + Format(Now, "MM/DD/YYYY") + "','" + X + " ','" + ono.Text + "',2,'e',1," + POSO.Text + "," + POSO.Text + ",'" + PELATHS + "','S" + mID(X, 2, 6) + "')"
Gdb.Execute sql, n


On Error Resume Next

Dim STR_EKT As String
STR_EKT = X + " " + Format(Now, "DD/MM/YYYY") _
+ " " + Left(EPO.Text + Space(25), 25) _
+ " " + Left(EPA.Text + Space(20), 20) + _
" " + Left(DIE.Text + Space(22), 22) + _
" " + Left(ono.Text + Space(25), 25) + _
" " + Left(MON.Text + Space(4), 3) + _
" " + Format(POSO, "#0") + _
" " + Left(SKOPOS.Text + Space(15), 15)
'Printer.Print STR_EKT
'Printer.EndDoc








'Printer.FontName = "Courier New"
'Printer.FontSize = 8

Open "LPT1" For Output As 1
 Print #1, Chr(15) + to437(STR_EKT)
Close 1



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


For Each PR In Printers
    If InStr(PR.DeviceName, "SATO") > 0 Then
        Set Printer = PR
        Exit For
    End If
Next
      
      
      
      
      
'   BarCode1.ShowText = NO
'     BarCode1.NarrowBarWidth = 0.025  '0.03
'    BarCode1.DataToEncode = DataGrid1.Text
'    Printer.PaintPicture BarCode1.Picture, -18, 124   '-88,124


MsgBox "йатавыягхгйе ле том йыдийо " + X



10
Dim C As Integer
   C = 7
      
Dim W As Integer
For W = 1 To 2
      Printer.FontName = "128"
      Printer.FONTSIZE = 6
      Printer.CurrentY = 10
      Printer.Print Tab(C + 3); X
      
      Printer.CurrentY = 50
      
      Printer.FontBold = False '  True
      Printer.FONTSIZE = 6
      Printer.FontUnderline = False 'True
      Printer.FontName = "Arial Greek"
      
Printer.FONTSIZE = 7
      Printer.Print Tab(C); "кАЦэЙГР Computers"
Printer.FONTSIZE = 7
Printer.Print Tab(C); " ТГК.25210-22225"   'CommonDialog1'.FontName
      
      Printer.FontUnderline = False
      
      
      
      
      Printer.CurrentY = 360
      
      
      Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
      Printer.FontBold = True
      
      Printer.FONTSIZE = 10
      Printer.Print Tab(C); X;
      
      
      
      
      Printer.CurrentY = 390
      Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
      Printer.FontBold = False
      
      Printer.FONTSIZE = 7
      Printer.Print Tab(C + 30); Now
      
      
      
      
      
      
      
      
      
      
      
      Printer.CurrentY = 580
      
      Printer.FONTSIZE = 8
      Printer.FontBold = True
      Printer.Print Tab(C); Left(EPO.Text, 22) + " " + THL.Text
      
      
      Printer.FontName = "Courier New"
      Printer.FONTSIZE = 7
      Printer.Print Tab(C); Left(ono, 40)
      Printer.Print Tab(C); Left(BLABH.Text, 35)
      Printer.Print Tab(C); mID(BLABH.Text, 36, 40)
      
     ' Printer.Print
      ' Printer.Print Tab(c); Now
      Printer.EndDoc

Next W

'GoTo 10











Printer.EndDoc


Unload Me



End Sub
Function FIND_NUM(DUM As Integer, DELTIO As Integer) As String
'DELTIO=1 да пекатг   DELTIO=0 дектио пос.паяакабгс
'************************
Dim rec As ADODB.Recordset

'rec.Open "select * FROM MEM", Gdb, adOpenKeyset, adLockOptimistic
Dim REC2 As New ADODB.Recordset
'REC2.Open "SELECT * FROM MEM", Gdb, adOpenDynamic, adLockOptimistic
'REC2.MoveFirst

rPEL = ""

'  Exit Function

If DELTIO = 1 Then
'  REC2.Move 8  ' ДА ПЕКАТЫМ
   rPEL = "19591959"
Else
   rPEL = "19601960"
'  REC2.Move 9 ' ДЕК.ПОС.ПАЯАКАБчР
End If
  
  Dim k
  
  'rec.Edit
  REC2.Open "SELECT * FROM MEM WHERE RPEL=" + rPEL, Gdb, adOpenDynamic, adLockOptimistic
  
  If IsNull(REC2("epel")) Then
      Gdb.Execute "update MEM SET EPEL=1 WHERE RPEL=" + rPEL, k
      REC2("epel") = 1
  Else
      Gdb.Execute "update MEM SET EPEL=EPEL+1 WHERE RPEL=" + rPEL, k
      
  End If
If DUM > 0 Then
    'REC2.Update
End If

  
REC2.Close
  REC2.Open "SELECT * FROM MEM WHERE RPEL=" + rPEL, Gdb, adOpenDynamic, adLockOptimistic

  'диабафы тгм пяогцоулемг глеяолгмиа циа ма лгм евы глеяолгмиа лийяотеяг тгс текеутаиас йатавыягсгс
   
'REC2.Close





Dim REC3 As New ADODB.Recordset


 REC3.Open "SELECT HM2 FROM EID WHERE KOD='" + "!" + Right("00000" + LTrim(Str(REC2("EPEL") - 1)), 5) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
  If REC3.EOF Then
     Dim ANS As Integer
     
     ANS = MsgBox("пяосовг дем упаявеи г пяогцоулемг йатавыягсг TOY бибкиоу. NA сумевисы;", vbYesNo)
     If ANS = vbNo Then
           Gdb.Execute "update MEM SET EPEL=EPEL-1 WHERE RPEL=" + rPEL, k
           End
        End
     End If
     
  Else
     
     
     If DateDiff("d", REC3(0), Now) < 0 Then
         MsgBox "г сглеяимг глеяолгмиа еимаи лийяотеяг апо тгм текеутаиа йатавыягсг"
         Gdb.Execute "update MEM SET EPEL=EPEL-1 WHERE RPEL=" + rPEL, k
         End
     End If

     
     
     
  End If
  
  
  
  FIND_NUM = "!" + Right("00000" + LTrim(Str(REC2("EPEL"))), 5)


End Function

Private Sub Command2_Click()
'HEADER
Dim STR_EKT As String
Dim ANS
ANS = MsgBox("еисаи сицоуяос;", vbYesNo, "ейтупысг епийежакидас")

If ANS = vbNo Then Exit Sub





STR_EKT = "а/а  HMEяолгмиа " _
+ " " + Left("Eпымулиа пекатг" + Space(25), 25) _
+ " " + Left("епаццекла" + Space(20), 20) + _
" " + Left("диеухумсг" + Space(22), 22) + _
" " + Left("пеяицяажг лгв/тос" + Space(25), 25) + _
" " + Left("MON" + Space(4), 3) + _
" " + Format("пос", "#0") + _
" " + "сйоп.паяакабгс"


Printer.FontName = "Courier New"
Printer.FONTSIZE = 8

Open "LPT1" For Output As 1
 Print #1, Chr(15) + to437(STR_EKT)
Close 1












End Sub

Private Sub Command3_Click()
    EPO.Text = ""
    EPA.Text = ""
    DIE.Text = " "
    THL.Text = " "
    
End Sub

Private Sub Command4_Click()



End Sub

Private Sub daPelath_Click()
'---- да пекатг -------------------------------------------
'дглиоуяцоуле ема RECORD сто EGGTIM KAI STO EID
' KAI тупымоуле

Dim X As String, sql As String
Dim n As Integer


Dim PR, FOUND

FOUND = False
For Each PR In Printers
    If InStr(PR.DeviceName, "POSPARALABH") > 0 Then
        FOUND = True
        Exit For
    End If
Next

If Not FOUND Then
   MsgBox "дем упаявеи се аутгм тгм хесг о ейтупытгс циа тис паяакабес"
'   Unload Me
 '  Exit Sub
End If


Dim PELATHS
If EPO.Text = Adodc1.Recordset("EPO") Then
  PELATHS = Adodc1.Recordset("KOD")
Else
  MsgBox "пяепеи ма упаявеи сто аявеио о пекатгс"
  Exit Sub
End If

If Len(Trim(PELATHS)) < 4 Then
   MsgBox "дем бяисйетаи о йыдийос тоу пекатг"
   Exit Sub
   
End If




If Len(Trim(EPO.Text)) = 0 Then
  MsgBox "дЕМ СУЛПКГЯЧСАТЕ ЭМОЛА"
  Exit Sub
End If


If Len(Trim(ono.Text)) = 0 Then
  MsgBox "дЕМ СУЛПКГЯЧСАТЕ ЕъДОР"
  Exit Sub
End If


If Len(Trim(da.Text)) = 0 Then
  MsgBox "дЕМ СУЛПКГЯЧСАТЕ АЯИХ.ДЕКТъОУ"
  Exit Sub
End If



'дглиоуяциа тоу йыдийоу апохгйгс

X = FIND_NUM(0, 1) ' бяисйы то текеутаио моулеяо KAI
    'диабафы тгм пяогцоулемг глеяолгмиа циа ма лгм евы глеяолгмиа лийяотеяг тгс текеутаиас йатавыягсгс






Dim R As New ADODB.Recordset
R.Open "SELECT KOD FROM EID WHERE KOD LIKE '" + Left(X, 2) + "%' ORDER BY KOD", Gdb, adOpenDynamic, adLockOptimistic
R.MoveLast
X = Left(R(0), 2) + Format(Val(mID(R(0), 3, 5)) + 1, "0000")



sql = "INSERT INTO EID (KOD,ONO,FPA,MEMO,MON,CH1,NUM1,CH2,CH3,CH4,CH5,CH6,HM2) VALUES (" _
& "'" + X + " ','" + ono.Text + "',2,'" + BLABH.Text + "'," _
& "'" + MON.Text + "','" + SKOPOS.Text + "'," + POSO.Text + "," _
& "'" + Left(EPO.Text, 20) + "','" + Left(DIE.Text, 20) + "','" + EPA.Text + "','" + THL.Text + "','" + PELATHS + "'," _
& "'" + Format(DTPicker1, "mm/dd/yyyy") + "')"

Gdb.Execute sql, n




If n = 0 Then
  MsgBox " ДЕМ ЙАТАВЫЯчХГЙЕ "
  Exit Sub
End If


Dim DUM As String

DUM = FIND_NUM(1, 1) ' ЙАТАВЫЯЧ ТО ЕПЭЛЕМО МОЩЛЕЯО


'дглиоуяциа тгс йимгсгс апохгйгс (еGGTIM)


sql = "INSERT INTO EGGTIM (HME,KODE,ONOMA,FPA,EIDOS,APOT,POSO,XRE,PELKOD,ATIM) VALUES (" _
& "'" + Format(DTPicker1.value, "MM/DD/YYYY") + "','" + X + " ','" + ono.Text + "',2,'e',1," + POSO.Text + "," + POSO.Text + ",'" + PELATHS + "','А" + Format(da, "000000") + "')"
Gdb.Execute sql, n


sql = "INSERT INTO TIM (HME,EIDOS,ATIM,KPE,B_N1,AJ1,AJ2,AJ3,AJ4,AJ5) VALUES (" _
& "'" + Format(DTPicker1.value, "MM/DD/YYYY") + "','e','А" + Format(da, "000000") + "','" + PELATHS + "',1,0,0,0,0,0)"
Gdb.Execute sql, n


sql = "INSERT INTO EGG (HME,EIDOS,AIT,KOD,ATIM,XRE) VALUES (" _
& "'" + Format(DTPicker1.value, "MM/DD/YYYY") + "','e','да пекатг','" + PELATHS + "','А" + Format(da, "000000") + "',0)"
Gdb.Execute sql, n


'rt.AddNew
'  rt("hme") = DTPicker1.Value
'  rt("atim") = "н" + m_tim
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


For Each PR In Printers
    If InStr(PR.DeviceName, "SATO") > 0 Then
        Set Printer = PR
        Exit For
    End If
Next
      
      
      
      
      
'   BarCode1.ShowText = NO
'     BarCode1.NarrowBarWidth = 0.025  '0.03
'    BarCode1.DataToEncode = DataGrid1.Text
'    Printer.PaintPicture BarCode1.Picture, -18, 124   '-88,124


MsgBox "йатавыягхгйе ле том йыдийо " + X



10
Dim C As Integer
   C = 7
      
Dim W As Integer
SSS:

For W = 1 To 1
      
      Printer.FontName = "128"
      Printer.FONTSIZE = 6
      Printer.Print Tab(8 + C); X
      
      Printer.CurrentY = 100
      
      
      Printer.FontName = "Arial Greek"   'CommonDialog1'.FontName
      Printer.FontBold = True
      Printer.FONTSIZE = 10
      Printer.Print Tab(C); X
     
      
      Printer.CurrentY = 600
      
      Printer.FONTSIZE = 8
      Printer.FontBold = True
      Printer.Print Tab(C); EPO.Text + " " + THL.Text
      'Printer.Print
      
      Printer.FontName = "Courier New"
      Printer.FONTSIZE = 8
      Printer.Print Tab(C); Left(BLABH.Text, 40)
     ' Printer.Print Tab(c); Mid(BLABH.Text, 41, 40)
      
      'Printer.Print
       Printer.Print Tab(C); "ДА " + da.Text + " " + Format(DTPicker1, "dd/mm/yyyy")
      Printer.EndDoc

Next

'GoTo SSS












Printer.EndDoc


Unload Me



















End Sub

Private Sub Form_Load()
DTPicker1.value = Now

 
 paralabh.Picture = LoadPicture(gPicture)

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



Adodc1.ConnectionString = gConnect
Adodc1.RecordSource = "select top 10 * from PEL"
Adodc1.Refresh



















End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

  
  If Adodc1.Recordset.EOF Then
    Exit Sub
  End If
  
  
  EPO = CNull(Adodc1.Recordset("EPO"))
   
 'If IsNull(Adodc1.Recordset("DIE")) Then
  '   DIE = ""
   '  Else
     DIE = CNull(Adodc1.Recordset("DIE"))
     
   
   
   EPA = CNull(Adodc1.Recordset("EPA"))
   THL = CNull(Adodc1.Recordset("THL"))
   
   



End Sub

Private Sub Text8_LostFocus()
   Adodc1.RecordSource = "SELECT EPO,DIE,* FROM PEL WHERE EPO LIKE '" + UCase(SameLetters(Text8.Text)) + "%'"
   Adodc1.Refresh
   TDBGrid1.Columns(0).width = 3000
   TDBGrid1.Columns(1).width = 3000
   TDBGrid1.Splits(0).AlternatingRowStyle = True
   TDBGrid1.Splits(0).AnchorRightColumn = True
  If Adodc1.Recordset.EOF Then
    Exit Sub
  End If
  
  
   EPO = Adodc1.Recordset("EPO")
   DIE = CNull(Adodc1.Recordset("DIE"))
   EPA = CNull(Adodc1.Recordset("EPA"))
   THL = CNull(Adodc1.Recordset("THL"))
   
   
   
   
   
   
   
   
   
   
   
End Sub



Sub LAG_ETIK()

End Sub

