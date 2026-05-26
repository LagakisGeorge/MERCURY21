VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form par8 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   11115
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   23760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   26
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ейтупысг"
      Height          =   615
      Left            =   22200
      TabIndex        =   25
      Top             =   3360
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "par8.frx":0000
      Height          =   7215
      Left            =   11280
      TabIndex        =   24
      Top             =   3360
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   12726
      _Version        =   393216
      BackColor       =   12632256
      Cols            =   5
      FixedCols       =   0
      BackColorBkg    =   8421440
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   19920
      Top             =   5760
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "лгдемифы то  пистокайи"
      Height          =   495
      Left            =   20400
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "кгьг апо пистокайи"
      Height          =   495
      Left            =   20400
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "енодос"
      Height          =   495
      Left            =   20400
      TabIndex        =   19
      Top             =   2640
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc xalia 
      Height          =   330
      Left            =   13080
      Top             =   10800
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.CommandButton Command1 
      Caption         =   "йатавыягсг дектиоу"
      Height          =   495
      Left            =   7560
      TabIndex        =   17
      Top             =   10680
      Width           =   3255
   End
   Begin VB.CommandButton PELKATAX 
      BackColor       =   &H00C0C000&
      Caption         =   "йатавыягсг меоу пекатг"
      Enabled         =   0   'False
      Height          =   495
      Left            =   16200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox mTHL 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      TabIndex        =   14
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox mDie 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      TabIndex        =   12
      Top             =   1080
      Width           =   5055
   End
   Begin VB.TextBox mEpo 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      TabIndex        =   9
      Top             =   480
      Width           =   5055
   End
   Begin VB.TextBox mkod 
      BackColor       =   &H00FFFFC0&
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
      Left            =   13800
      TabIndex        =   8
      Top             =   20
      Width           =   975
   End
   Begin VB.CommandButton PELNEW 
      BackColor       =   &H00FFFFC0&
      Caption         =   "меос пекатгс"
      Height          =   495
      Left            =   20400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8640
      Top             =   240
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
      Caption         =   ""
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "par8.frx":0015
      Height          =   1695
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   24
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
         Size            =   12
         Charset         =   161
         Weight          =   700
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
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox Atim 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox timtext3 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   10800
      Width           =   1572
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   7215
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   12726
      _Version        =   393216
      Rows            =   25
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   200
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc XALIA2 
      Height          =   330
      Left            =   10680
      Top             =   10800
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "йЫДИЙЭР ПЕКэТГ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   27
      Top             =   720
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   11280
      Top             =   0
      Width           =   7815
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "текеутаиос         пекатгс поу тгкежымгсе"
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
      Left            =   11400
      TabIndex        =   23
      Top             =   2640
      Width           =   6135
   End
   Begin VB.Label Label9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   22
      Top             =   3000
      Width           =   7095
   End
   Begin VB.Label Label8 
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   3000
      Width           =   7095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "тгкежымо"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   15
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "диеухумсг-ояожос"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   13
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Eпымуло-ONOMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   11
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "йыдийос"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "еПЫМУЛъА ПЕКэТГ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "аЯИХЛЭР дЕК.пОС.пАЯАКАБчР"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "par8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F_MDI As String

Dim F_PROTH As Integer

Private Sub Atim_GotFocus()
   Atim.BackColor = vbYellow
   
End Sub

Private Sub Atim_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   KeyCode = 0
   Text.SetFocus
End If
End Sub

Private Sub Atim_LostFocus()
   Atim.BackColor = vbWhite
   
   Dim R As New ADODB.Recordset
   
   R.Open "select COUNT(*)  FROM TIM  WHERE ATIM='a" + Format(Atim, "00000") + "' and YEAR(HME)=YEAR(GETDATE())", Gdb, adOpenForwardOnly, adLockReadOnly
  If R(0) >= 1 Then
     MsgBox "пяосовг упаявеи паки о аяихлос " + Atim
     Atim.SetFocus
  End If
   
   
   
   
   
End Sub

Private Sub Command1_Click()
' йатавыягс дектиоу посотийгс паяакабгс


Dim k As Long

Dim c1, c2, c3, C4, C5, C6, C7



    
    
    sql = "INSERT INTO TIM (ATIM,HME,KPE,EIDOS,KLEIDI ) values ("
    sql = sql + "'a" + Format(Val(Atim), "00000") + "',"                      ' ATIM
    sql = sql + "'" + Format(Now, "MM/DD/YYYY") + "',"  ' HME
    sql = sql + "'" + Adodc1.Recordset("KOD") + "','e'," + "'a" + Format(Val(Atim), "00000") + "')" 'pelkod    eidos
    Gdb.Execute sql
    
    
    
    
    
    
    
    
    
    
    
    
'sql = "INSERT INTO TIM (ATIM,HME,KPE,EIDOS ) values ("
 '   sql = sql + "'a" + Format(Val(Atim), "00000") + "',"                      ' ATIM
  '  sql = sql + "'" + Format(Now, "MM/DD/YYYY") + "')"  ' HME
   ' sql = sql + "'" + Adodc1.Recordset("KOD") + "','e')"  'pelkod    eidos
    'Gdb.Execute sql


For k = 1 To Grid1.ROWS - 1
    c1 = Grid1.TextMatrix(k, 1)
    c2 = Grid1.TextMatrix(k, 2)
    c3 = Grid1.TextMatrix(k, 3)
    C4 = Grid1.TextMatrix(k, 4)
    C5 = Grid1.TextMatrix(k, 5)
    C6 = Grid1.TextMatrix(k, 6)
    
' дглиоуяцы то RECORD TOY EID
    C0 = Grid1.TextMatrix(k, 0)
    
    If Len(Trim(c1)) = 0 Then
       Exit For
    End If
    
    sql = "INSERT INTO EID (KOD,ONO,LTI,NUM3,NUM1,NUM2,FPA,CH1,CH3,HM1) values ("
    sql = sql + "'" + Format(Val(C0), "00000") + "',"  ' kod
    sql = sql + "'" + c2 + "',"                        ' ono
    sql = sql + "'" + Format(Val(C6), "000.00") + "'," ' LTI
    sql = sql + "'" + Format(Val(Grid1.TextMatrix(k, 7)), "000.00") + "'," ' NUM3
    sql = sql + "'" + Format(Val(c3), "00000") + "',"  ' NUM1
    sql = sql + "'" + Format(Val(C4), "00000") + "',2,"  ' NUM2 , FPA
    sql = sql + "'" + Adodc1.Recordset("KOD") + "',"
    sql = sql + "'" + Format(Val(Atim), "00000") + "'," ' ая.дектиоу посотийгс
    sql = sql + "'" + Format(Now, "MM/DD/YYYY") + "')" ' HME
    Gdb.Execute sql


    
    
    
' дглиоуяцы то RECORD TOY EGGTIM
    sql = "INSERT INTO EGGTIM (KODE,ONOMA,POSO,TIMM,ATIM,HME,PELKOD,EIDOS,APOT ) values ("
    sql = sql + "'" + Format(Val(C0), "00000") + "',"  ' kod
    sql = sql + "'" + c2 + "',"                        ' onoMA
    sql = sql + "'" + Format(Round(Val(C5), 2), "000.00") + "',"  ' POSO
    sql = sql + "'" + Format(Val(Grid1.TextMatrix(k, 7)) + Val(C6), "000.00") + "'," ' TIMM
    sql = sql + "'a" + Format(Val(Atim), "00000") + "'," ' ATIM
    sql = sql + "'" + Format(Now, "MM/DD/YYYY") + "',"  'HME
    sql = sql + "'" + Adodc1.Recordset("KOD") + "','e',1)"  'pelkod    eidos  APOT
    Gdb.Execute sql
    

Next



Command1.Enabled = False

Atim.Text = ""
Text.Text = ""

Grid1.Clear

Atim.SetFocus


End Sub

Function get_next_pel(meidos As String) As Integer
Dim db As Database
Dim R As New ADODB.Recordset
Dim R0 As New ADODB.Recordset

Dim k As Long

On Error Resume Next


'Set db = OpenDatabase(gDir)
'Set db = OpenDatabase(gDir, False, False)
'Set db = OpenDatabase(gDir, False, False, gConnect)

If kenoi.value = vbChecked Then
  R.Open "select KOD FROM PEL WHERE LEN(KOD)=5 AND EIDOS='" + meidos + "' order by KOD", Gdb, adOpenForwardOnly, adLockReadOnly
  If R.EOF Then
     get_next_pel = "00000"
    R.Close: Exit Function
  End If
  
  Me.MousePointer = vbHourglass
  

  R.MoveFirst
  k = 0
  Do While Not R.EOF
   If Val(R(0)) = 0 Then
      R.MoveNext
   Else
     k = k + 1
     If Left(R(0), 4) = Format(k, "0000") Then
        R.MoveNext
     Else
        On Error GoTo 0
       R0.Open "select COUNT(*) FROM PEL WHERE KOD='" + Format(k, "00000") + "' AND EIDOS='" + meidos + "' ", Gdb, adOpenForwardOnly, adLockReadOnly
       If R0(0) > 0 Then 'гдг упаявеи
          R0.Close
          Exit Do
       Else
         get_next_pel = Format(k, "00000")
         R.Close
         Me.MousePointer = vbNormal
         Exit Function
       End If
     
     
     End If
   End If
  Loop
  R.Close

Dim ll, L
'

For L = k + 1 To 9999
      R.Open "select COUNT(*) FROM PEL WHERE KOD='" + Format(L, "00000") + "' AND EIDOS='" + meidos + "' ", Gdb, adOpenForwardOnly, adLockReadOnly
      ll = R(0)
      R.Close
      DoEvents
      Me.Caption = k
      If ll = 0 Then
        Exit For
      End If
Next
get_next_pel = Format(L, "00000")
Me.MousePointer = vbNormal
Exit Function
  
  
  
  
  
  
  
End If





R.Open "SELECT *FROM MEM", Gdb, adOpenDynamic, adLockOptimistic

If meidos = "e" Then
   get_next_pel = R("epel") + 1
Else
   get_next_pel = R("rpel") + 1
End If
End Function



Private Sub Command2_Click()
    Dim ANS
    If Val(Atim.Text) > 0 Then
       
       ANS = MsgBox("ма вахоум ои аккацес;", vbYesNo)
       If ANS = vbNo Then
         Exit Sub
       End If
       
    End If
    Unload Me

End Sub

Private Sub Command3_Click()
Dim DUM
Dim mEggtim As Long
Dim mRafia As Long



mEggtim = 0
mRafia = 0



On Error Resume Next


Kill "C:\EGGTIM.TXT"
MILSEC 1000




Kill "C:\RAFIA.TXT"

MILSEC 1000

On Error GoTo 0



DUM = Shell("C:\SENDTOPC\FROMCE.BAT", vbNormalFocus)


MILSEC 2000
If Len(Dir("C:\EGGTIM.TXT", vbNormal)) Then
   MsgBox "паяекгжхг то аявеио летягсеым"
   mEggtim = 1
End If

If Len(Dir("C:\RAFIA.TXT", vbNormal)) Then
   MsgBox "паяекгжхг то аявеио топохетгсеым се яажиа"
   mRafia = 1
End If

Dim a As String
Dim M, P, X
Dim POSO As Single
Dim N As Long
Printer.FONTSIZE = 16
If mEggtim = 1 Then
  '
  '001112 0200 0100
  '001400 0242 0124
  '003214 0255 0144

    Open "c:\eggtim.txt" For Input As #1
          Do While Not EOF(1)
              Input #1, a
              If Len(a) > 4 Then
                 M = (mID(a, 8, 4))
                 P = (mID(a, 13, 4))
                 X = (mID(a, 2, 5))
                 POSO = Val(M) * Val(P) / 100 / 100
                 
                 If POSO = 0 Or POSO > 50 Then
                    Printer.Print Tab(5); "кахос диастасеис  о йыдийос " + X + "  дем емглеяыхгйе "
                 Else
                    Gdb.Execute "UPDATE EID SET NUM1=" + M + ",NUM2=" + P + " where KOD=" + X, N
                    If N = 0 Then
                       Printer.Print Tab(5); "дем бяехгйе о йыдийос     " + X + "      дем емглеяыхгйе "
                    Else
                       Gdb.Execute "UPDATE EGGTIM SET ONOMA=ONOMA+' " + M + " X " + P + "',POSO=" + Format(POSO, "####.##") + " where LEFT(ATIM,1)='a' and  KODE='" + X + "'", N
                    End If
                 End If
                 
                 
              End If
          
          
          Loop
          
     
     
     Close #1
         
     
End If


If mRafia = 1 Then
    
'000001 A2
'000005 A2
'000006 A2
'000007 A2

    Open "c:\rafia.txt" For Input As #1
          Do While Not EOF(1)
              Input #1, a
              If Len(a) > 4 Then
                 M = (mID(a, 8, 4))
                 X = (mID(a, 1, 6))
                 Gdb.Execute "UPDATE EID SET CH2='" + M + "' where CAST(KOD AS NUMERIC)=" + Trim(X), N
              End If
          Loop
     Close #1
End If






































Printer.EndDoc






End Sub

Private Sub Command4_Click()

Dim DUM
Dim ANS
ANS = MsgBox("мА ЛГДЕМИСТЕъ ТО ПИСТОКэЙИ;", vbYesNo)
If ANS = vbYes Then
    DUM = Shell("C:\SENDTOPC\DELRAFIA.BAT", vbNormalFocus)
    MsgBox "ой лГДЕМъСТГЙЕ"
End If


End Sub

Private Sub Command5_Click()
' ейтупысг

Dim sql As String

sql = XALIA2.RecordSource

        
'If ekt Then
  'typos2
  print3_xar sql, "001111111111111111", Label9.Caption, 0
  
'End If








End Sub

Private Sub DataGrid1_Click()
   
 'On Error Resume Next
'   Label8.Caption = Adodc1.Recordset("EPO") + Adodc1.Recordset("DIE")
 '  xalia.RecordSource = "select KOD,ONO,CH1 AS [яажи] ,NUM1 AS [лгйос] ,NUM2 AS [пкатос],LTI AS [TIMH] ,NUM3 AS [сумтгягсг] FROM EID WHERE CH1='" + Adodc1.Recordset("KOD") + "'"

  
 ' xalia.ConnectionString = gConnect
 ' xalia.Refresh
  
  
          '    DataGrid3.Columns(0).width = 500
'          DataGrid3.Rebind
    '      DataGrid3.Refresh
      Grid1.Enabled = True
          
          XALIA2.RecordSource = "select KOD,ONO ,CH2 AS [яажи] ,NUM1 AS [лгйос] ,NUM2 AS [пкатос],CAST(LTI AS NUMERIC(10, 2)) AS [TIMH] ,NUM3 AS [сумтгягсг],CAST(NUM1*NUM2/100/100*(LTI+NUM3) AS NUMERIC(10,2)) AS [аниа],CH3 AS [ая.дект],HM1 AS [глея.дек]  FROM EID WHERE CH1='" + Adodc1.Recordset("KOD") + "'"
          XALIA2.ConnectionString = gConnect
          XALIA2.Refresh
        
        
        
        
        MSHFlexGrid1.ColWidth(0) = 1000
        MSHFlexGrid1.ColWidth(1) = 3500
        MSHFlexGrid1.ColWidth(2) = 800
        MSHFlexGrid1.ColWidth(3) = 800
        MSHFlexGrid1.ColWidth(4) = 800
        
         MSHFlexGrid1.ColWidth(5) = 800
         
         
         MSHFlexGrid1.AddItem ""
         MSHFlexGrid1.AddItem ""
         MSHFlexGrid1.AddItem ""
         
         Dim R As New ADODB.Recordset
         R.Open "SELECT SUM(NUM1*NUM2/100/100*(LTI+NUM3)) AS SYN FROM EID WHERE CH1='" + Adodc1.Recordset("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic
         
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.ROWS - 3, 6) = "йахаяо"
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.ROWS - 3, 7) = Format(R(0), "####.00")
         
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.ROWS - 2, 6) = "жпа"
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.ROWS - 2, 7) = Format(R(0) * 0.23, "####.00") 'R(0) * 0.23
         
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.ROWS - 1, 6) = "сумоко"
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.ROWS - 1, 7) = Format(R(0) * 1.23, "####.00") 'R(0) * 1.23
         
         MSHFlexGrid1.ColAlignment(MSHFlexGrid1.ColS - 1) = 6
         
         
        
        'MSHFlexGrid1.AddItem ""
        'MSHFlexGrid1.AddItem ""
        'MSHFlexGrid1.AddItem ""
        
        
        
        
        
        
          'Dim R As New ADODB.Recordset
         R.Close
         R.Open "SELECT EPO,KOD,EPA,* FROM PEL WHERE KOD='" + Adodc1.Recordset("KOD") + "'", Gdb, adOpenDynamic, adLockOptimistic
          
          
          F_MDI = MDIForm1.Caption
          Label9.Caption = R("EPO") + " " + CNull(R("DIE")) + " " + CNull(R("THL")) ' Adodc1.Recordset("EPO") + Adodc1.Recordset("DIE")
'      End If

  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  timtext3.SetFocus
  
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   KeyCode = 0
  DataGrid1_Click
End If

End Sub

Private Sub Form_Load()
Dim k As Long

F_PROTH = 1

For k = 0 To Grid1.ROWS - 1
      Grid1.RowHeight(k) = 350
Next
Grid1.ColS = 8

'Grid1.height = Grid1.RowHeight(0) * Grid1.ROWS
Grid1.ColWidth(2) = 2600
Grid1.ColWidth(0) = 800


Dim w As Long
w = 0
For k = 0 To Grid1.ColS - 1
     w = w + Grid1.ColWidth(k)
Next

Grid1.width = DataGrid1.width

' Grid1.width = w + 350
      
Grid1.TextMatrix(0, 0) = "а/а"
Grid1.TextMatrix(0, 1) = "йыдийос"
Grid1.TextMatrix(0, 2) = "пеяицяажг"
Grid1.TextMatrix(0, 3) = "лгйос"
Grid1.TextMatrix(0, 4) = "пкатос"
Grid1.TextMatrix(0, 5) = "посотгта"

Grid1.TextMatrix(0, 6) = "тилг лом"
Grid1.TextMatrix(0, 7) = "сумтгягсг"


'          DataGrid3.Columns(0).width = 500
'          DataGrid3.Columns(2).width = 500
 '        DataGrid3.Columns(3).width = 500
      
      
      
      
      Grid1.Text = timtext3.Text ' ЙЯАТэЕИ ТО grid ТО ЙЕъЛЕМО
      Grid1_EnterCell
      
     ' Atim.SetFocus
      
End Sub

Private Sub Form_Paint()
 If F_PROTH = 1 Then
    F_PROTH = 0
    Atim.SetFocus
End If

End Sub

Private Sub mDie_GotFocus()
   mDie.BackColor = vbYellow
End Sub

Private Sub mDie_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
     KeyCode = 0
     mTHL.SetFocus
   End If

End Sub

Private Sub mDie_LostFocus()
   mDie.BackColor = vbYellow
End Sub

Private Sub mEpo_GotFocus()
   mEpo.BackColor = vbYellow
   
End Sub

Private Sub mEpo_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
     KeyCode = 0
     mDie.SetFocus
   End If

End Sub

Private Sub mEpo_LostFocus()
   mEpo.BackColor = vbWhite
End Sub

Private Sub mTHL_GotFocus()
  mTHL.BackColor = vbYellow
End Sub

Private Sub mTHL_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = 13 Then
      KeyCode = 0
      PELKATAX.SetFocus
   End If

End Sub

Private Sub mTHL_LostFocus()
  mEpo.BackColor = vbWhite
End Sub

Private Sub PELKATAX_Click()
Dim sql As String


sql = "INSERT INTO PEL (EIDOS,KOD,EPO,DIE,THL,AFM) VALUES ("
sql = sql + "'e',"
sql = sql + "'" + mkod + "',"
sql = sql + "'" + mEpo + "',"
sql = sql + "'" + mDie + "',"
sql = sql + "'" + mTHL + "','" + Atim.Text + "')"

Dim N As Long

Gdb.Execute sql, N

''Dim r As New ADODB.Recordset
   Adodc1.ConnectionString = gConnect
   
   Adodc1.RecordSource = "select EPO,DIE,THL,EPA,* from PEL WHERE EIDOS='e' and KOD='" + mkod + "'"
   Adodc1.Refresh
 
   Grid1.Enabled = True
   
   
   
   mkod = ""
   mEpo = ""
   mDie = ""
   mTHL = ""
   

   PELKATAX.Enabled = False
   SendKeys ("")
   timtext3.SetFocus
   Grid1.Col = 0
   Grid1_EnterCell
End Sub

Private Sub PELNEW_Click()
      mkod = Format(get_next_pel("e"), "00000")
      PELKATAX.Enabled = True
      
      mEpo.SetFocus
      
End Sub

Private Sub Text_GotFocus()
   Text.BackColor = vbYellow
End Sub

Private Sub Text_KeyUp(KeyCode As Integer, Shift As Integer)
'
 
   If KeyCode = 13 Then
     If Len(Text.Text) = 0 Then
         KeyCode = 0
          Text.BackColor = vbWhite
         Text1.SetFocus
     Else
        Dim R As New ADODB.Recordset
         Adodc1.ConnectionString = gConnect
   
         Adodc1.RecordSource = "select EPO,DIE,THL,KOD,* from PEL WHERE EIDOS='e' and EPO LIKE '%" + Text.Text + "%'"
         Adodc1.Refresh
         Text.BackColor = vbWhite
         If Adodc1.Recordset.RecordCount = 0 Then
            mEpo = Text.Text
            PELNEW_Click
         ElseIf Adodc1.Recordset.RecordCount = 1 Then
            DataGrid1.SetFocus
         ElseIf Adodc1.Recordset.RecordCount > 1 Then
            DataGrid1.SetFocus
         End If
         Command1.Enabled = True
         KeyCode = 0
         timtext3.SetFocus
     End If
   End If
End Sub

Private Sub Text1_GotFocus()
 Text1.BackColor = vbYellow
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'   KeyCode = 0
'   timtext3.SetFocus
   
'End If


   If KeyCode = 13 Then
         Dim R As New ADODB.Recordset
         Adodc1.ConnectionString = gConnect
   
         Adodc1.RecordSource = "select EPO,DIE,THL,KOD,* from PEL WHERE EIDOS='e' and KOD ='" + Text1.Text + "'"
         Adodc1.Refresh
         Text1.BackColor = vbWhite
         If Adodc1.Recordset.RecordCount = 0 Then
            Text.SetFocus  'дем бяехгйе
         ElseIf Adodc1.Recordset.RecordCount = 1 Then
            DataGrid1.SetFocus
         ElseIf Adodc1.Recordset.RecordCount > 1 Then
            DataGrid1.SetFocus
         End If
         Command1.Enabled = True
         KeyCode = 0
         timtext3.SetFocus
   End If


End Sub

Private Sub Timer1_Timer()
   On Error Resume Next
   
   If Left(F_MDI, 10) = Left(MDIForm1.Caption, 10) Then
   
   Else
      If Val(Trim(mID(MDIForm1.Caption, 17, 5))) > 0 Then
         
            DataGrid3.Columns(0).width = 500
'          DataGrid3.Rebind
          DataGrid3.Refresh
          
          
          XALIA2.RecordSource = "select KOD,ONO ,CH2 AS [яажи] ,NUM1 AS [лгйос] ,NUM2 AS [пкатос],CAST(LTI AS NUMERIC(10, 2)) AS [TIMH] ,NUM3 AS [сумтгягсг],CAST(NUM1*NUM2/100/100*(LTI+NUM3) AS NUMERIC(10,2)) AS [аниа],CH3 AS [ая.дект],HM1 AS [глея.дек]  FROM EID WHERE CH1='" + Trim(mID(MDIForm1.Caption, 17, 5)) + "'"
          XALIA2.ConnectionString = gConnect
          XALIA2.Refresh
        
        
        
        
        MSHFlexGrid1.ColWidth(0) = 1000
        MSHFlexGrid1.ColWidth(1) = 3500
        MSHFlexGrid1.ColWidth(2) = 800
        MSHFlexGrid1.ColWidth(3) = 800
        MSHFlexGrid1.ColWidth(4) = 800
        
         MSHFlexGrid1.ColWidth(5) = 800
         
         
         MSHFlexGrid1.AddItem ""
         MSHFlexGrid1.AddItem ""
         MSHFlexGrid1.AddItem ""
         
         Dim R As New ADODB.Recordset
         R.Open "SELECT SUM(NUM1*NUM2/100/100*(LTI+NUM3)) AS SYN FROM EID WHERE CH1='" + Trim(mID(MDIForm1.Caption, 17, 5)) + "'", Gdb, adOpenDynamic, adLockOptimistic
         
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.ROWS - 3, 6) = "йахаяо"
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.ROWS - 3, 7) = Format(R(0), "####.00")
         
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.ROWS - 2, 6) = "жпа"
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.ROWS - 2, 7) = Format(R(0) * 0.23, "####.00") 'R(0) * 0.23
         
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.ROWS - 1, 6) = "сумоко"
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.ROWS - 1, 7) = Format(R(0) * 1.23, "####.00") 'R(0) * 1.23
         
         MSHFlexGrid1.ColAlignment(MSHFlexGrid1.ColS - 1) = 6
         
         
        
        'MSHFlexGrid1.AddItem ""
        'MSHFlexGrid1.AddItem ""
        'MSHFlexGrid1.AddItem ""
        
        
        
        
        
        
          'Dim R As New ADODB.Recordset
          R.Open "SELECT EPO,KOD,EPA,* FROM PEL WHERE KOD='" + Trim(mID(MDIForm1.Caption, 17, 5)) + "'", Gdb, adOpenDynamic, adLockOptimistic
          
          
          F_MDI = MDIForm1.Caption
          Label9.Caption = R("EPO") + " " + CNull(R("DIE")) + " " + CNull(R("THL")) ' Adodc1.Recordset("EPO") + Adodc1.Recordset("DIE")
      End If
   End If
   
  
End Sub

Private Sub TIMText3_GotFocus()
     
     timtext3.BackColor = vbYellow
     timtext3.SelStart = 0
     timtext3.SelLength = Len(timtext3.Text)


End Sub

Private Sub timText3_KeyDown(KeyCode As Integer, Shift As Integer)
'емтяу тилокоциоу
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
'емтяу тилокоциоу
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

Dim R As New ADODB.Recordset


  If KeyCode = 38 Then   ' памы бекос
      
      Grid1.Text = timtext3.Text ' ЙЯАТэЕИ ТО grid ТО ЙЕъЛЕМО
      
      ' ПэЕИ ЛъА ЙОКЭМА АЯИСТЕЯэ
      If Grid1.Col >= 1 Then
         Grid1.Col = Grid1.Col - 1
      Else
        'П'АЕИ ЛъА СЕИЯэ ЕПэМЫ
        If Grid1.row > 1 Then
           Grid1.row = Grid1.row - 1
           Grid1.Col = 1
        End If
      End If
      Grid1_EnterCell
   End If

  If KeyCode = 40 Then   ' KATы бекос
      Grid1.Text = timtext3.Text ' ЙЯАТэЕИ ТО grid ТО ЙЕъЛЕМО
      ' ЛИэ СЕИЯэ ПАЯАЙэТЫ
      If Grid1.row < Grid1.ROWS - 1 Then
         'FindSynolo
         Grid1.row = Grid1.row + 1
               Grid1_EnterCell
      End If

   End If
   

   If KeyCode = 13 Then   ' enter
      Grid1.Text = timtext3.Text ' ЙЯАТэЕИ ТО grid ТО ЙЕъЛЕМО
      ' ЛИэ СТГКГ ПАЯАЙэТЫ
      If Grid1.Col = Grid1.ColS - 1 Then
         
         If Grid1.row = Grid1.ROWS - 1 Then
         Else
            Grid1.row = Grid1.row + 1
            Grid1.Col = 0
            timtext3.Text = Grid1.TextMatrix(Grid1.row - 1, 0) + 1
         End If
         
      Else
         If Grid1.Col = 1 Then
            
            If Len(Grid1.Text) = 0 Then
              Command1.SetFocus
              Exit Sub
            End If
            
            
            
            R.Open "select * from EID WHERE KOD='" + Grid1.Text + "'", Gdb, adOpenDynamic, adLockOptimistic
            
            Grid1.Col = 2
            If R.EOF Then
               Me.Caption = "дем бяехгйе"
               Exit Sub
            Else
               Grid1.Text = R("ONO")
               If IsNull(R("LTI")) Then
                 Grid1.TextMatrix(Grid1.row, 6) = 0
               Else
                 Grid1.TextMatrix(Grid1.row, 6) = R("LTI")
               End If
            End If
            
            Grid1.Col = Grid1.Col + 1
         ElseIf Grid1.Col = 4 Then
         
             Grid1.TextMatrix(Grid1.row, 5) = Val(Grid1.TextMatrix(Grid1.row, 3)) * Val(Grid1.TextMatrix(Grid1.row, 4)) / 10000
             Grid1.Col = Grid1.Col + 1
         
         ElseIf Grid1.Col = 0 Then 'ELEGXV MHPVS YPARXEI JANA O KODIKOS
             R.Open "select COUNT(*) from EID WHERE KOD='" + Grid1.Text + "'", Gdb, adOpenDynamic, adLockOptimistic
             
             If R(0) > 0 Then
                MsgBox "упаявеи о йыдийос"
                Grid1.Text = ""
             Else
                Grid1.Col = Grid1.Col + 1
             End If
         Else
             Grid1.Col = Grid1.Col + 1
         End If
         
         Grid1_EnterCell
      End If

   End If
  









  If KeyCode = 27 Then 'esc ТЕКОР ЕИСАЦЫЦчР
       KeyCode = 0
       Combo4.SetFocus
  End If


End Sub
Private Sub Grid1_EnterCell()
'  If Grid1.MouseRow = 0 Then
'    timtext3.Visible = False
'    Exit Sub
'   End If

On Error GoTo MHNYMA

   If Grid1.row > Grid1.ROWS - 1 Then
    Grid1.row = Grid1.ROWS - 1
   End If
   
   timtext3.Text = Grid1.Text
   timtext3.Visible = False
   timtext3.Top = Grid1.Top + Grid1.CellTop
   timtext3.Left = Grid1.Left + Grid1.CellLeft
   timtext3.width = Grid1.CellWidth - 15
   timtext3.height = Grid1.CellHeight - 25
   
   timtext3.SelStart = 0
   timtext3.SelLength = Len(timtext3.Text)
   

   
   Label9.Caption = Grid1.Text
   
   
   timtext3.Visible = True
   timtext3.SetFocus
   
   
   Exit Sub

MHNYMA:
'' HandleError "Par1:EnterCell"
Resume Next
            
   
   
   
   
End Sub

Private Sub Grid1_LeaveCell()
'Grid1.Text = timtext3.Text
End Sub

