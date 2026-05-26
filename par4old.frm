VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form par4 
   Caption         =   "Form1"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9510
   ScaleWidth      =   11145
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
      Height          =   3855
      Left            =   240
      TabIndex        =   15
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6800
      _Version        =   393216
      Rows            =   14
      Cols            =   14
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   14
   End
   Begin MSDataGridLib.DataGrid DbGrid1 
      Bindings        =   "par4.frx":0000
      Height          =   2655
      Left            =   1200
      TabIndex        =   14
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2040
      Top             =   8400
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "data1"
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
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "par4.frx":0014
      Left            =   6840
      List            =   "par4.frx":001E
      TabIndex        =   13
      Text            =   "Combo2"
      Top             =   3960
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   3960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   50528257
      CurrentDate     =   37735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   10200
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ÓõãêñéôéêÜ"
      Height          =   375
      Left            =   7650
      TabIndex        =   11
      Top             =   3015
      Width           =   1935
   End
   Begin MSChart20Lib.MSChart G 
      Height          =   3885
      Left            =   120
      OleObjectBlob   =   "par4.frx":004B
      TabIndex        =   10
      Top             =   4320
      Width           =   10815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Óýíïëá êáôá ìÞíá"
      Height          =   495
      Left            =   8265
      TabIndex        =   9
      Top             =   60
      Width           =   1335
   End
   Begin VB.CommandButton SumHmeras 
      Caption         =   "Óýíïëá êáôá çìÝñá"
      Height          =   495
      Left            =   8265
      TabIndex        =   8
      Top             =   630
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   9960
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\MERCVB\par4-1.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin VB.CommandButton exit 
      Caption         =   "Åîïäïò"
      Height          =   495
      Left            =   8280
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton ShowPar 
      Caption         =   "ÐñïâïëÞ"
      Height          =   495
      Left            =   8280
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8040
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   3360
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   3960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   50528257
      CurrentDate     =   37735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Åùò"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Áðü"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   3960
      Width           =   615
   End
End
Attribute VB_Name = "par4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Combo2_Click()


'---
Dim WHERE, FROM, FIELDS, titfields, taxinomisi
Dim dum, db As Database, A
Dim R As Recordset


dum = OpenD(db)



If Combo2.ListIndex = 0 Then 'ÄÅËÔÉÁ


   A = InputBox("0=ÓÅ ÅÊÊÑÅÌÏÔÇÔÁ 1=ÔÉÌÏËÏÃÇÈÅÍÔÁ", "ÅÐÉËÏÃÇ ÐÁÑÁÓÔÁÔÉÊÙÍ", 0)
   WHERE = " left(atim,1)='A' and hme>=#" + Format(DTPicker1.Value, "mm/dd/yyyy") + "# and hme<=#" + Format(DTPicker2.Value, "mm/dd/yyyy") + "# "
   If A = "0" Then
      WHERE = WHERE + " AND ( ART IS NULL OR LEFT(ART,1)=' ' ) "
   Else
      WHERE = WHERE + " AND LEFT(ART,1)='" + A + "'"
   End If
   FROM = " tim inner join pel on tim.eidos=pel.eidos and tim.kpe=pel.kod"
   FIELDS = "HME,atim,pel.epo,SXETIKO,LEFT(skopos,25) "
   titfields = " hme,atim "
   taxinomisi = " hme"
End If

If Combo2.ListIndex = 1 Then  ' ÔÉÌÏËÏÃÉÁ
   'a = InputBox("0=ÓÅ ÅÊÊÑÅÌÏÔÇÔÁ 1=ÔÉÌÏËÏÃÇÈÅÍÔÁ", "ÅÐÉËÏÃÇ ÐÁÑÁÓÔÁÔÉÊÙÍ", 0)
   WHERE = "LEFT(ATIM,1)='" + Right(Combo1.Text, 1) + "' AND  hme>=#" + Format(DTPicker1.Value, "mm/dd/yyyy") + "# and hme<=#" + Format(DTPicker2.Value, "mm/dd/yyyy") + "# "
   FROM = " tim inner join pel on tim.eidos=pel.eidos and tim.kpe=pel.kod"
   FIELDS = "HME,atim,PEL.EPO,AJI,AJ1,AJ2,AJ3,AJ4,FPA1+FPA2+FPA3+FPA4 "
   titfields = " hme,atim "
   taxinomisi = " hme"
End If











Set R = db.OpenRecordset("select *from vbrepeid where aa=99")
R.Edit
R!WHERE = WHERE
R!Table = FROM
R!FIELDS = FIELDS
R!titfields = titfields
R!taxinomisi = taxinomisi
R.Update

'db.Execute "update vbrepeid set vrepeid.from='" + FROM + "' ,fields='" + FIELDS + "', titfields='" + titfields + "', taxinomisi='" + taxinomisi + "' where aa=99"
g_ektyp 99









End Sub

Private Sub Command1_Click()
Dim db
Set db = OpenDatabase(gDir, False, False, gConnect)
db.Execute "DROP TABLE TEMP"
db.Execute "select atim,format(hme,'dd/mm/yy') as shme,pel.epo,format(aji,'####0.##') as saji into temp from tim inner join pel on 'e'=pel.eidos and tim.kpe=pel.kod where instr('" + Right(Combo1.Text, 1) + "',left(atim,1))>0 and hme>=#" + Format(DTPicker1.Value, "mm/dd/yyyy") + "# and hme<#" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "# order by hme;"
'CrystalReport1.Action = 1
  
End Sub



Private Sub Command2_Click()
'===========================================
' êáôá ìÞíá ðùëÞóåéò áãïñÝò
'===========================================
Dim d As String

Dim pol As String
Dim POLEPIS As String
Dim ago As String
Dim AGOEPIS As String
Dim A As String
Dim Var As String
Dim Stat(1 To 30, 1 To 12) As Single
Dim syn(1 To 12)
Dim s As Integer
Dim R As Recordset
Dim k As Integer, l As Integer, addit As Integer, addit2 As Integer
Dim d1 As Date, d2 As Date
Dim sy

MSFlexGrid1.Cols = 14
'MSFlexGrid1.DataSource = Null

If Check1 Then ' palia xronia
    ' Data1.DatabaseName = gDir + "\" + Right(Format(Year(Now) - 1, "0000"), 2)
    Data1.Refresh
    d1 = DTPicker1.Value - 365
    d2 = DTPicker2.Value - 365
    addit = 1 'gia to Chart
    addit2 = 6 ' ãéá ôï grid

Else  ' ôñÝ÷ïõóá ÷ñïíéÜ
    'Data1.DatabaseName = gDir
    
    Data1.RecordSource = "select top 1 *from TIM"
    Data1.Refresh
    d1 = DTPicker1.Value
    d2 = DTPicker2.Value
    addit = 0 'gia to Chart
    addit2 = 0 ' ãéá ôï grid
End If


DBGrid1.Visible = False
MSFlexGrid1.Visible = True

  MSFlexGrid1.row = 0:  MSFlexGrid1.Col = 1:  MSFlexGrid1.Text = "ÐùëÞóåéò"
  MSFlexGrid1.row = 0:  MSFlexGrid1.Col = 2:  MSFlexGrid1.Text = "Åðéóôñ.Ðùë"
  
  MSFlexGrid1.row = 0:  MSFlexGrid1.Col = 3:  MSFlexGrid1.Text = "Êáè.Ðùë."
  
  MSFlexGrid1.row = 0:  MSFlexGrid1.Col = 4:  MSFlexGrid1.Text = "ÁãïñÝò"
  MSFlexGrid1.row = 0:  MSFlexGrid1.Col = 5:  MSFlexGrid1.Text = "Åðéóô.Áã."
  MSFlexGrid1.row = 0:  MSFlexGrid1.Col = 6:  MSFlexGrid1.Text = "Êáè.Áãïñ."


  MSFlexGrid1.row = 0:  MSFlexGrid1.Col = 7:  MSFlexGrid1.Text = "Ð.ÐùëÞóåéò"
  MSFlexGrid1.row = 0:  MSFlexGrid1.Col = 8:  MSFlexGrid1.Text = "Ð.Åðéó.Ðùë"
  MSFlexGrid1.row = 0:  MSFlexGrid1.Col = 9:  MSFlexGrid1.Text = "Ð.Êáè.Ðùë."
  
  MSFlexGrid1.row = 0:  MSFlexGrid1.Col = 10:  MSFlexGrid1.Text = "Ð.ÁãïñÝò"
  MSFlexGrid1.row = 0:  MSFlexGrid1.Col = 11:  MSFlexGrid1.Text = "Ð.Åð.Áã."
  MSFlexGrid1.row = 0:  MSFlexGrid1.Col = 13:  MSFlexGrid1.Text = "Êáè.Ðùë."


  MSFlexGrid1.row = 13:  MSFlexGrid1.Col = 0:  MSFlexGrid1.Text = "Óýíïëá"


A = Get_AJ(pol, POLEPIS, ago, AGOEPIS)




'-- ÐùëÞóåéò
Var = pol: s = 1 + addit2: GoSub findD

'-- ÅðéóôñïöÝò ÐùëÞóåùí
Var = POLEPIS: s = 2 + addit2: GoSub findD
On Error GoTo 0


sy = 0
For k = 1 To 12
  MSFlexGrid1.TextMatrix(k, addit2 + 3) = Format(Stat(k, addit2 + 1) - Stat(k, addit2 + 2), "###,###.00")
  sy = sy + (Stat(k, addit2 + 1) - Stat(k, addit2 + 2))
  MSFlexGrid1.row = k: MSFlexGrid1.Col = addit2 + 3
  MSFlexGrid1.CellBackColor = vbYellow
Next
  MSFlexGrid1.TextMatrix(13, addit2 + 3) = Format(sy, "###,###.00")
'-- AãïñÝò
Var = ago: s = 4 + addit2: GoSub findD

'ÅðéóôñïöÝò áãïñþí
Var = AGOEPIS: s = 5 + addit2: GoSub findD
  
sy = 0
For k = 1 To 12
  MSFlexGrid1.TextMatrix(k, addit2 + 6) = Format(Stat(k, addit2 + 4) - Stat(k, addit2 + 5), "###,###.00")
  sy = sy + (Stat(k, addit2 + 4) - Stat(k, addit2 + 5))
  MSFlexGrid1.row = k: MSFlexGrid1.Col = addit2 + 6
  MSFlexGrid1.CellBackColor = vbYellow
Next
  MSFlexGrid1.TextMatrix(13, addit2 + 6) = Format(sy, "###,###.00")


For k = 1 To 12
  MSFlexGrid1.row = k
  MSFlexGrid1.Col = 0
  MSFlexGrid1.Text = k
Next
  

' ÃñáöéêÞ ðáñÜóôáóç
s = 0

For l = 1 To 12
   ' Exit Sub
   G.row = 1 + addit: G.Column = l: G.Data = Stat(l, addit2 + 1) - Stat(l, addit2 + 2)
   G.row = 4 + addit: G.Column = l: G.Data = Stat(l, addit2 + 4) - Stat(l, addit2 + 5)
Next
  
  
  
  
  
  
 Me.MousePointer = vbNormal
Exit Sub



findD:
  
  d = "select MONTH(HME) as shme,sum(AJ1+AJ2+AJ3+AJ4+AJ5) as saji "
  d = d + "from TIM  where  left(ATIM,1) IN (" + Var + ")  "
  d = d + "and HME>='" + Format(d1, "mm/dd/yyyy") + "' and HME<'" + Format(d2 + 1, "mm/dd/yyyy") + "' "
  d = d + "group by MONTH(HME) ;"
  Data1.RecordSource = d
  Data1.Refresh
  
On Error Resume Next
Data1.Recordset.MoveFirst
  Do While Not Data1.Recordset.EOF
     Stat(Val(Data1.Recordset("shme")), s) = Data1.Recordset("saji")
     Data1.Recordset.MoveNext
  Loop
  syn(s) = 0
For k = 1 To 12
  MSFlexGrid1.row = k
  MSFlexGrid1.Col = s
  MSFlexGrid1.Text = Format(Stat(k, s), "###,###.00")
  syn(s) = syn(s) + Stat(k, s)
Next
  
  
  
  MSFlexGrid1.row = 13
  MSFlexGrid1.Col = s
  MSFlexGrid1.Text = Format(syn(s), "###,###.00")



Return
  
  
  
  

  
End Sub

Private Sub Command3_Click()
   G.top = 4680  'arxika
   G.height = 4485
   
   G.top = 46
   G.height = 8485
   
   
   
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim db As Database
Dim R As New ADODB.Recordset

'Set db = OpenDatabase(gDir, False, False, gConnect)

R.Open UCase("select pol,titlos,eidos from parastat;"), Gdb, adOpenDynamic, adLockOptimistic

Combo1.Clear

Combo1.AddItem "ÐùëÞóåéò"
Combo1.AddItem "ÁãïñÝò"
  
  R.MoveFirst
  Do While Not R.EOF
     If Not IsNull(R("TITLOS")) And R("pol") Then
        Combo1.AddItem R("TITLOS") + Space(30) + R("eidos")
     End If
     R.MoveNext
  Loop

Combo1.Enabled = True
Combo1.Text = Combo1.List(geidos_paras)

'Data1.DatabaseName = gDir
Data1.ConnectionString = gConnect
DTPicker1.Value = CDate("01/01/" + Str(Year(Now)))
DTPicker2.Value = Now


End Sub

Private Sub ShowPar_Click()

Dim d As String, db As Database
Dim R As New ADODB.Recordset, pol As String, ago As String
Dim mChoose, mSaji2
Dim A, POLEPIS As String, AGOEPIS As String

A = Get_AJ(pol, POLEPIS, ago, AGOEPIS)

DBGrid1.Visible = True
MSFlexGrid1.Visible = False

'Data1.RecordSource = "SELECT *FROM MEM"
'
'
'Data1.Refresh



'pol = ""
'ago = ""
''Set db = OpenDatabase(gDir, False, False, gConnect)
'R.Open "select POL,TITLOS,EIDOS from PARASTAT", Gdb, adOpenDynamic, adLockOptimistic
'
'  R.MoveFirst
'  Do While Not R.EOF
'     If Not IsNull(R("TITLOS")) Then
'       If R("pol") = "1" Then
'          pol = pol + R("eidos")
'       ElseIf R("pol") = 2 Then
'          ago = ago + R("eidos")
'       End If
'     End If
'     R.MoveNext
'  Loop
'R.Close

If Combo1.ListIndex = -1 Then Combo1.ListIndex = 0



If Combo1.ListIndex = 0 Then
   mChoose = pol
ElseIf Combo1.ListIndex = 1 Then
   mChoose = ago
Else
   mChoose = "'" + Right(Combo1.Text, 1) + "'"
End If



' d = "select atim,format(hme,'dd/mm/yy') as shme,kpe,aji,eidos into temp" + LTrim(Str(gXEIRISTHS))
'  d = d + " from tim  "
'  d = d + "where instr('" + mChoose + "',left(atim,1))>0 and hme>=#" + Format(DTPicker1.Value, "mm/dd/yyyy") + "# and hme<#" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "# order by hme;"
'
'db.Execute "drop table temp1"
'db.Execute d
'
'
'  d = "select temp1.*,pel.epo   from temp" + LTrim(Str(gXEIRISTHS)) + "  inner join pel on temp1.eidos+temp1.kpe=pel.eidos+pel.kod "
'
'  Data1.RecordSource = d '"select *from temp1"
' Data1.Refresh
'
'  'MSFlexGrid1.DataSource = "Data1"
'  Exit Sub










Me.MousePointer = vbHourglass
  d = "select ATIM,HME as shme,PEL.EPO,AJI " 'into temp" + LTrim(Str(gXEIRISTHS))
  d = d + " from TIM inner join PEL on TIM.EIDOS=PEL.EIDOS and TIM.KPE=PEL.KOD "
  d = d + "where left(ATIM,1) IN (" + mChoose + ") and HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "' order by HME;"
  
  'db.Execute d

On Error Resume Next
'Gdb.Execute "drop table TEMP" + LTrim(Str(gXEIRISTHS))
'Gdb.Execute d

  
  On Error GoTo 0
  Data1.RecordSource = d ' "select sum(AJI) as saji2 from TEMP" + LTrim(Str(gXEIRISTHS))
  
  Data1.Refresh
  
  'mSaji2 = data1.Recordset("saji2")
  
  'data1.RecordSource = "select *from TEMP" + LTrim(Str(gXEIRISTHS))
  
  
  'data1.Refresh
  'data1.Recordset.AddNew
  'data1.Recordset("aji") = mSaji2
  'data1.Recordset("epo") = "Ó Õ Í Ï Ë Á "
  'data1.Recordset.Update
  

  
  'data1.Refresh
  DBGrid1.Columns(3).Alignment = dbgRight
  
  DBGrid1.Columns(0).Caption = "Ðáñáóôáôéêü"
  DBGrid1.Columns(1).Caption = "Çìåñïìçíßá"
  DBGrid1.Columns(2).Caption = "ÐåëÜôçò /ÐñïìçèåõôÞò"
  DBGrid1.Columns(3).Caption = "Áîßá ìå ÖÐÁ"
  
  DBGrid1.Columns(0).width = 1000
  DBGrid1.Columns(1).width = 1000
  DBGrid1.Columns(2).width = 3800
  DBGrid1.Columns(3).width = 800
  DBGrid1.Columns(3).NumberFormat = "##,###,##0.00"
  Me.MousePointer = vbNormal
End Sub

Private Sub SumHmeras_Click()


Dim d As String

DBGrid1.Visible = True
MSFlexGrid1.Visible = False



Me.MousePointer = vbHourglass
  d = "select HME as shme,sum(AJI) as saji "
  d = d + "from TIM  where '" + Right(Combo1.Text, 1) + "'=left(ATIM,1) "
  d = d + "and HME>='" + Format(DTPicker1.Value, "mm/dd/yyyy") + "' and HME<'" + Format(DTPicker2.Value + 1, "mm/dd/yyyy") + "' "
  d = d + "group by HME;"
  Data1.RecordSource = d
  
  Data1.Refresh
  
  
  DBGrid1.Columns(1).Caption = "Óýíïëï ìå ÖÐÁ"
  DBGrid1.Columns(0).Caption = "Çìåñïìçíßá"
  DBGrid1.Columns(1).NumberFormat = "##,###,##0.00"
  
 
 Me.MousePointer = vbNormal

End Sub
