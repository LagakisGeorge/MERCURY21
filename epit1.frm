VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form EPIT1 
   BackColor       =   &H00FF0000&
   Caption         =   "Εισαγωγή Επιταγής"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15435
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   15435
   WindowState     =   2  'Maximized
   Begin VB.CheckBox ISXEIROGRAFO 
      BackColor       =   &H00FF0000&
      Caption         =   "Χειρόγραφο"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6000
      TabIndex        =   33
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Αρίθμηση"
      Height          =   330
      Left            =   3840
      TabIndex        =   32
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   2760
      TabIndex        =   30
      Top             =   7800
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "epit1.frx":0000
      Left            =   240
      List            =   "epit1.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   7320
      Width           =   3495
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   135
      Left            =   120
      TabIndex        =   27
      Top             =   8520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   238
      _Version        =   393217
      TextRTF         =   $"epit1.frx":0004
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6600
      Top             =   9120
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8280
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Mercvb\reports\paralepit.rpt"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Εκτύπωση Παραλαβής"
      Height          =   330
      Left            =   3840
      TabIndex        =   26
      Top             =   7320
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GR2 
      Bindings        =   "epit1.frx":0534
      Height          =   1815
      Left            =   4560
      TabIndex        =   6
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3201
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GR1 
      Bindings        =   "epit1.frx":0548
      Height          =   3615
      Left            =   0
      TabIndex        =   24
      Top             =   1800
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   10
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc EGG 
      Height          =   375
      Left            =   4110
      Top             =   8340
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   375
      Left            =   4080
      Top             =   7920
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "DATA2"
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
      Height          =   375
      Left            =   4800
      Top             =   8820
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
   Begin VB.TextBox parat 
      Height          =   330
      Left            =   225
      TabIndex        =   12
      Top             =   6840
      Width           =   3495
   End
   Begin VB.CommandButton Update 
      Caption         =   "Καταχώρηση"
      Enabled         =   0   'False
      Height          =   330
      Left            =   3840
      TabIndex        =   13
      Top             =   6840
      Width           =   1935
   End
   Begin VB.TextBox arepit 
      Height          =   360
      Left            =   7470
      TabIndex        =   11
      Top             =   5850
      Width           =   1695
   End
   Begin VB.TextBox poso 
      Height          =   360
      Left            =   5640
      TabIndex        =   10
      Top             =   5850
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "epit1.frx":055A
      Left            =   3840
      List            =   "epit1.frx":055C
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5850
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Εξοδος"
      Height          =   330
      Left            =   7470
      TabIndex        =   17
      Top             =   6840
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Προμηθευτής"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   100
         TabIndex        =   2
         Top             =   720
         Width           =   1600
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Πελάτης"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   100
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1500
      End
   End
   Begin MSComCtl2.DTPicker hmee 
      Height          =   360
      Left            =   225
      TabIndex        =   7
      Top             =   5820
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   635
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   445186049
      CurrentDate     =   38814
   End
   Begin MSComCtl2.DTPicker hmel 
      Height          =   360
      Left            =   1965
      TabIndex        =   8
      Top             =   5850
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   635
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      Format          =   445186049
      CurrentDate     =   38814
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Αρίθμηση"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   31
      Top             =   7850
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Παράμετροι εκτύπωσης"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   2655
      Left            =   120
      Top             =   5520
      Width           =   10095
   End
   Begin VB.Label ONOMA 
      BackColor       =   &H00FF0000&
      Caption         =   "Label8"
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
      Left            =   0
      TabIndex        =   25
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Παρατηρήσεις"
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
      Left            =   270
      TabIndex        =   23
      Top             =   6495
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Αριθμός Επιταγής"
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
      Left            =   7440
      TabIndex        =   22
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ποσό επιταγής"
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
      Left            =   5640
      TabIndex        =   21
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Τράπεζα"
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
      Left            =   3840
      TabIndex        =   20
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ημερ.Λήξεως"
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
      Left            =   2040
      TabIndex        =   19
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ημερ.Καταχώρησ"
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
      Left            =   360
      TabIndex        =   18
      Top             =   5520
      Width           =   1215
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
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   16
      Top             =   1080
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
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   15
      Top             =   600
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
      Left            =   1920
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "EPIT1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PROTH, fflag
Attribute fflag.VB_VarUserMemId = 1073938432

Dim F_KODPEL As String

Dim f_buff   As String

Dim f_epo
Attribute f_epo.VB_VarUserMemId = 1073938436
Dim F_EKTYPOTHS As String ' '= FINDPARAMETROI(1, "PAR1", "F_EKTYPOTHS", "..", "Εκτυπωτής παραστατικών(με .. προεπιλεγμένος)") '
Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)
Attribute fh.VB_VarUserMemId = 1073938437

'Function TOST(ByVal S As Single, ByVal N As Integer) As String
'
'    TOST = Right(Space(30) + Format(S, "###,###.00"), N)
'
'End Function

'Function tostC(ByVal S, ByVal N As Integer) As String
'
'    If IsNull(S) Then
'        tostC = Space(N)
'    Else
'        tostC = Left(S + Space(N), N)
'    End If
'
'End Function

'Private Sub Command3_Click()
'    Dim k As Integer
'    'If fflag = 0 Then Exit Do
'    PELAT42.SHOW
'End Sub

'Private Sub Command4_Click()
'    pelat43.SHOW 1
'
'    EGG.Refresh
'    EGG.Recordset.MoveLast
'End Sub




Dim F_ISXEIROGRAFO As Integer     ' = Trim()



Private Sub Combo2_Change()

        '<EhHeader>
        On Error GoTo Combo2_Change_Err

        '</EhHeader>

        Dim R As New ADODB.Recordset

100     R.Open "select * from XREOPIS where EPIT=1 AND EIDOS=" + Right(Combo2.Text, 3), Gdb, adOpenDynamic, adLockOptimistic
110     Text1.Text = R("ARITMISI")
120     R.Close
  
        '<EhFooter>
        Exit Sub

Combo2_Change_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.Combo2_Change " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.Combo2_Change " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo2_Click()

        '<EhHeader>
        On Error GoTo Combo2_Click_Err

        '</EhHeader>

100     Combo2_Change

        '<EhFooter>
        Exit Sub

Combo2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.Combo2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.Combo2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>
        
       '  Gdb.Execute "UPDATE XREOPIS SET ARITMISI=" + str(Val(Text1.Text)) + " 1  WHERE EIDOS=" + Right(Combo2.Text, 3)

'         Dim R As New ADODB.Recordset
'
'          R.Open "SELECT ISNULL(ARITMISI,0)  FROM XREOPIS  WHERE EIDOS=" + Right(Combo2.Text, 3), Gdb, adOpenForwardOnly, adLockReadOnly
'          If Not R.EOF Then
'              Text1.Text = R(0)
'
'          End If
'          R.Close
          
          
          
      '

         On Error Resume Next

        
        ' Gdb.Execute "UPDATE XREOPIS SET ARITMISI=ARITMISI+1 WHERE EIDOS=" + Right(Combo2.Text, 3)

        ' R.Close
       '  R.Open "SELECT ARITMISI FROM XREOPIS WHERE EIDOS=7", Gdb, adOpenForwardOnly, adLockReadOnly





        Dim DDD

        ' R.Close
        'R.Open "SELECT ARITMISI FROM XREOPIS WHERE EIDOS=7", Gdb, adOpenForwardOnly, adLockReadOnly
        
        
        
'          Dim R9 As New ADODB.Recordset
'
'          R9.Open "select * from XREOPIS where EIDOS=7", Gdb, adOpenDynamic, adLockOptimistic
'          where EIDOS = " + Right(Combo2.Text, 3)"
'          R9.Close
        
        
        
        
        
        

100     DDD = Text1.Text  ' R("ARITMISI")

        ' R.Close

        On Error Resume Next

        Dim DB

        Dim RR As Recordset

110     Set DB = OpenDatabase("C:\MERCVB\REPORTS\REPORTS.MDB")
120     Set RR = DB.OpenRecordset("mem")

        ' On Error GoTo Command2_Click_Err

130     RR.Edit
140     RR!n1 = Val(POSO.Text)    ' poso
150     RR!PAR1 = DDD

160     RR!hme = hmee.Value

170     RR!PAR2 = Data1.Recordset("kod")
180     RR!C1 = Left(f_epo, 20)
190     RR!c2 = Data1.Recordset("epa")
200     RR!c3 = Data1.Recordset("die")
210     RR!c4 = Data1.Recordset("afm")

220     RR!par12 = parat.Text 'Left(Combo2.Text, Len(Combo2.Text) - 3) ' "ΠΑΡΑΛΑΒΗ ΕΠΙΤΑΓΗΣ"  '
    
230     RR!par17 = Combo1.Text ' TRAPEZA
240     RR!par18 = Format(hmel.Value, "DD/MM/YYYY") 'HMER.LHJEVS
250     RR!par19 = arepit ' ΑΡΙΘΜΟς ΕΠΙΤΑΓΗς

        Dim sql

        '    sql = "select sum(XREOSI) AS XX0,sum(PISTOSI) AS PP0 from EGG WHERE  HME>='" + "01/01/" + LTrim(Str(Year(Now))) + "' AND HME<='" + Format(Now, "MM/DD/YYYY") + "' AND EIDOS ='" + DATA1.Recordset("eidos") + "' and KOD='" + DATA1.Recordset("KOD") + "'"

260     RR!par10 = "Αρ.Επ." + arepit + " Λήξη " + Format(hmel, "dd/mm/yyyy")    'Format(R(0) + Data1.Recordset("ayp") - R(1), "###,000.00")
        'R.Close

270     RR!n4 = BRES_YPOLPEL(f_buff, Data1.Recordset("KOD"))

        'RR!PAR1 = Data1.Recordset("doy")
280     RR!PAR11 = Data1.Recordset("doy")
290     RR!memo = Olografos(Val(POSO.Text))
    
        On Error GoTo Command2_Click_Err

300     RR!n4 = BRES_YPOLPEL(f_buff, F_KODPEL)
    
310     RR.Update
    
320     If gBuff = "e" Then
330         CrystalReport1.ReportFileName = "C:\Mercvb\reports\paralepit.rpt"
        Else
340         CrystalReport1.ReportFileName = "C:\Mercvb\reports\epitpromi.rpt"
        End If
    
350     CrystalReport1.Destination = crptToPrinter
360   '  CrystalReport1.Action = 1
370    ' CrystalReport1.Action = 1

         Ucr9print.printingCR9 F_EKTYPOTHS, (CrystalReport1.ReportFileName)
        Gdb.Execute "UPDATE XREOPIS SET ARITMISI=" + str(Val(Text1.Text)) + "   WHERE EIDOS=" + Right(Combo2.Text, 3)

        
        Command2.Enabled = False
        

'        Dim R9 As New ADODB.Recordset

'
'     R9.Open "select * from XREOPIS where EPIT=1 AND EIDOS=" + Right(Combo2.Text, 3), Gdb, adOpenDynamic, adLockOptimistic
'     Text1.Text = R9("ARITMISI")
'     R9.Close
        
        
        
        
        
        
        
        
        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub arepit_GotFocus()

        '<EhHeader>
        On Error GoTo arepit_GotFocus_Err

        '</EhHeader>

100     arepit.BackColor = vbYellow

        '<EhFooter>
        Exit Sub

arepit_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.arepit_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.arepit_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub arepit_KeyPress(KeyAscii As Integer)
    '    If KeyAscii = vbKeyReturn Then
    '       KeyAscii = 0 'suppress the beep
    '       keybd_event VK_TAB, 0, 0, 0 'send a tab
    '    End If
End Sub

Private Sub arepit_LostFocus()

        '<EhHeader>
        On Error GoTo arepit_LostFocus_Err

        '</EhHeader>

100     arepit.BackColor = vbWhite

        '<EhFooter>
        Exit Sub

arepit_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.arepit_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.arepit_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    '    If KeyAscii = vbKeyReturn Then
    '       KeyAscii = 0 'suppress the beep
    '       keybd_event VK_TAB, 0, 0, 0 'send a tab
    '    End If

End Sub

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

100     Unload Me

110     If gApoMenu = True Then
120         SendKeys "%"
130         SendKeys "{RIGHT}+{RIGHT}+{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

'Private Sub DBGrid1_Click()
'    fflag = 1
'End Sub

'Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    If fflag = 0 Then Exit Sub
'    ' PELAT41.Show
'End Sub

Function KARTELLA(mgdir)

        '===============
        'Καρτελλα πελάτη
        '===============
        '<EhHeader>
        On Error GoTo KARTELLA_Err

        '</EhHeader>

        Dim X2 As String, s As String

        Dim yp As Single

        Dim SX As Single

        Dim SP As Single

        Dim DB As Database

        'Dim R As Recordset
        Dim XREOSI, PISTOSI

        Dim ayp  As Single

        Dim k    As Long

        Dim l    As Long, R As New ADODB.Recordset

        Dim MKOD As String

100     If Option1(1) Then X2 = "e" Else X2 = "r"
110     f_buff = X2
        '  If DATA1.Recordset.RecordCount = 0 Then Exit Function
        '  If IsNull(DATA1.Recordset("kod")) Then Exit Function

        On Error GoTo KARTELLA_Err

120     GR2.Col = 0    ' gia na bro ton kodiko

130     If IsEmpty(GR2.Text) Then Exit Function
140     EGG.ConnectionString = gConnect
150     s = GR2.Text
160     MKOD = s

170     GR2.Col = 3    ' gia na bro to  ayp
180     yp = Val(GR2.Text)

        'yp = IIf(IsNull(DATA1.Recordset("AYP")), 0, DATA1.Recordset("AYP"))

190     GR1.Visible = False

200     ayp = yp
210     EGG.RecordSource = "select *from EGG where EIDOS='" + X2 + "' and KOD='" + s + "' order by HME,ATIM; "

220     EGG.Refresh

        'Set R = EGG.Recordset
230     ' GoTo 200

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

240     s = MKOD  ' DATA1.Recordset("kod")
        'If x2 = "e" Then
250     EGG.RecordSource = "select ATIM as [Αρ_Παρ],CONVERT(smalldatetime,HME,5) as [Ημερoμηνία], EID as [Είδ], APA as [Αριθ_Παρ],AIT as [Αιτιολογία],XREOSI as [Χρέωση],PISTOSI as [Πίστωση],PROOD as [Υπόλοιπο] from EGG where EIDOS='" + X2 + "' and KOD='" + s + "' order by HME,ATIM;"
        'Else
        '  EGG.RecordSource = "select ATIM as [Αρ_Παρ],HME as [Ημερ/νία],EID as [Είδ],APA as [Αριθ_Παρ],AIT as [Αιτιολογία],XRE as [Χρέωση],XRE as [Πίστωση],PROOD as [Υπόλοιπο] from EGG where EIDOS='" + x2 + "' and KOD='" + s + "' order by HME,ATIM;"
        'End If
        'EGG.DatabaseName = mgdir

        '  EGG.Recordset.Close
260     EGG.Refresh

        '  yp = 0
270     GR1.Refresh

        On Error Resume Next

280     GR1.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "ΑΡΧΙΚΟ ΥΠΟΛΟΙΠΟ" & Chr(9) & Format(ayp, "###,###.00"), 1

290     GR1.ColWidth(4) = 2000
300     GR1.ColWidth(3) = 600    'ΑΠΑ
310     GR1.ColWidth(2) = 600    'ΕΙΔ

        Dim SU, NX, NP, NY

320     NX = 5: NP = 6: NY = 7
330     SU = ayp: SX = 0: SP = 0

340     For k = 2 To GR1.ROWS - 1
350         GR1.TextMatrix(k, 1) = Format(GR1.TextMatrix(k, 1), "dd/mm/yyyy")

360         If Val(GR1.TextMatrix(k, NX)) <> 0 Then
370             GR1.TextMatrix(k, NP) = ""
380             GR1.TextMatrix(k, NX) = Format(Val(GR1.TextMatrix(k, NX)), "#####.00")
            End If

390         If Val(GR1.TextMatrix(k, NP)) <> 0 Then
400             GR1.TextMatrix(k, NP) = Format(Val(GR1.TextMatrix(k, NP)), "#####.00")
410             GR1.TextMatrix(k, NX) = " "
            End If

420         GR1.TextMatrix(k, NY) = Format(Val(GR1.TextMatrix(k, NY)), "#####.00")
430         SX = SX + Val(DOT(GR1.TextMatrix(k, NX)))
440         SP = SP + Val(DOT(GR1.TextMatrix(k, NP)))
            ' SU = SU + SX - SP

450         If X2 = "e" Then
460             GR1.TextMatrix(k, NY) = Format(SU + SX - SP, "#####.00")
            Else
470             GR1.TextMatrix(k, NY) = Format(SU - SX + SP, "#####.00")
            End If

            'GR1.TextMatrix(K, nT) = Format(Val(GR1.TextMatrix(K, nT)), "######.00")

480         If k Mod 2 = 0 Then

490             For l = 0 To GR1.ColS - 1
500                 GR1.row = k: GR1.Col = l: GR1.CellBackColor = &HC0E0FF
                Next

            End If

        Next

510     GR1.AddItem ("" & Chr(9) & "Σύνολα" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(SX, "###,###.00") & Chr(9) & Format(SP, "###,###.00") & Chr(9))

520     GR1.TopRow = GR1.ROWS - 10
530     GR1.ColAlignment(5) = 6    ' RIGHT
540     GR1.ColAlignment(6) = 6    ' RIGHT
550     GR1.ColAlignment(7) = 6    ' RIGHT
560     GR1.Visible = True
        'End If

        '<EhFooter>
        Exit Function

KARTELLA_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.KARTELLA " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.KARTELLA " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub Command3_Click()
Dim FF As New bohu4
'bohu4.MDIChild = False
bohu4.SHOW 'vbModal




End Sub

Private Sub Form_DblClick()


    PARAMETROI.PARAM.Caption = "EPIT1"
    PARAMETROI.SHOW 1




End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then
110         Command1_Click

        End If

        'If KeyCode = 38 Then 'PANO BELOS KANEI SHIFT+TAB
        '   SendKeys "+{TAB}"
        'End If

120     If KeyCode = 13 Then    'PANO BELOS KANEI SHIFT+TAB
130         SendKeys "{TAB}"
        End If

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 11790, 9885, fh, fw, ft, fl
    
    
    Dim h As Long
    h = Combo1.Height
    
    hmee.Height = h
    hmel.Height = h
    POSO.Height = h
    arepit.Height = h
    parat.Height = h
    Update.Height = h
    Command1.Height = h
    Command2.Height = h
    
    Text1.Height = h
    
     Text1.Top = Label9.Top
    
    
    
    
    
    
    
End Sub

Private Sub GR2_Click()

        '<EhHeader>
        On Error GoTo GR2_Click_Err

        '</EhHeader>

        Dim DUM

100     GR2.Col = 1
110     ONOMA.Caption = GR2.Text
120     f_epo = GR2.Text
130     GR2.Col = 0: F_KODPEL = GR2.Text

        Dim l2 As Long

        Dim K2 As Long

        Dim MR As Long

140     MR = GR2.row

150     For K2 = 1 To GR2.ROWS - 1
160         GR2.row = K2

170         If K2 = MR Then

180             For l2 = 0 To GR2.ColS - 1
190                 GR2.Col = l2: GR2.CellBackColor = &HC0E0FF
                Next

            Else

200             For l2 = 0 To GR2.ColS - 1
210                 GR2.Col = l2: GR2.CellBackColor = vbWhite
                Next

            End If

        Next

220     GR2.row = MR

230     DUM = KARTELLA(gConnect)

240     Update.Enabled = True

           hmee.SetFocus ' den xreiazetai giati exo to enter

        Exit Sub

        '===============
        'Καρτελλα πελάτη
        '===============

        '================
        '--  ΚΑΡΤΕΛΛΑ ---
        '----------------
        Dim X2 As String, s As String

        Dim yp As Single

        'If F_RUN = 0 Then Exit Sub
        '

250     GR2.Col = 0    ' gia na bro ton kodiko

260     If IsEmpty(GR2.Text) Then Exit Sub
        ' EGG.DatabaseName = gDir
270     EGG.ConnectionString = gConnect

        '  On Error GoTo akiro
280     s = GR2.Text

290     If Option1(1) Then X2 = "e" Else X2 = "r"
300     EGG.RecordSource = "select ATIM as [Αρ.Παρ], CONVERT(DATETIME, HME, 5)  as [Hμερ/νία]," & " EID as [Ειδ],APA,AIT as [Αιτιολογία],XREOSI,PISTOSI,str(PROOD,10,2) from EGG where EIDOS='" + X2 + "' and KOD='" + s + "' order by HME,ATIM;"

        'select CONVERT(CHAR(10),HME,3) as [Ημερoμηνία], EID as [Είδ], APA as [Αριθ_Παρ]
        ',ATIM as [Πελατ],AIT as [Αιτιολογία],XREOSI as [Χρέωση],
        'PISTOSI as [Πίστωση],PROOD as [Υπόλοιπο],
        'USERID as [Χρήστης] from EGG
        'where ATIM LIKE '%' AND
        'HME>='01/01/2006' AND HME<='02/01/2007'
        'AND EIDOS in ('e','E') and KOD='0006' order by HME,ID;

310     EGG.Refresh

        Dim R As New ADODB.Recordset, NX, NP, NY

320     NX = 5: NP = 6: NY = 7

330     R.Open "SELECT *FROM PEL WHERE EIDOS='" + X2 + "' AND KOD='" + GR2.Text + "';", Gdb, adOpenDynamic, adLockOptimistic

340     If R.EOF Then
350         MsgBox "ΔΕΝ ΕΥΡΕΘΗ"
360         R.Close
370         Text2(0).SetFocus

            Exit Sub

        End If

380     F_KODPEL = R("kod")
390     f_buff = R("eidos")
400     f_epo = R("epo")

410     Update.Enabled = True

420     GR1.Visible = False
430     yp = R("AYP")
        'On Error Resume Next
440     R.Close
450     GR1.ColS = NY + 1

        Dim k, l, SX, SP, SU

460     SU = yp: SX = 0: SP = 0

470     For k = 1 To GR1.ROWS - 1
480         GR1.TextMatrix(k, 1) = Format(GR1.TextMatrix(k, 1), "dd/mm/yyyy")

490         If Val(GR1.TextMatrix(k, NX)) <> 0 Then
                ' GR1.TextMatrix(K, NP) = ""
500             GR1.TextMatrix(k, NX) = Format(Val(GR1.TextMatrix(k, NX)), "#####.00")
            End If

510         If Val(GR1.TextMatrix(k, NP)) <> 0 Then
520             GR1.TextMatrix(k, NP) = Format(Val(GR1.TextMatrix(k, NP)), "#####.00")
                ' GR1.TextMatrix(K, NX) = " "
            End If

530         SX = SX + Val(GR1.TextMatrix(k, NX))
540         SP = SP + Val(GR1.TextMatrix(k, NP))
            'SU = sx - sp
550         GR1.TextMatrix(k, NY) = Format(SU + SX - SP, "#####.00")
            'GR1.TextMatrix(K, nT) = Format(Val(GR1.TextMatrix(K, nT)), "######.00")

560         If k Mod 2 = 0 Then

570             For l = 0 To GR1.ColS - 1
580                 GR1.row = k: GR1.Col = l: GR1.CellBackColor = &HC0E0FF
                Next

            End If

        Next

590     GR1.Visible = True

600     GR1.ColAlignment(NY) = 6    'RIGHT

610     GR1.Refresh

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
620     If GR1.ROWS > 9 Then
630         GR1.TopRow = GR1.ROWS - 7
        End If

        'MSFlexGrid1.SetFocus
        'MSFlexGrid1.SelectionMode = True

640     hmee.SetFocus

        Exit Sub

akiro:
650     GoTo OLD100

        '<EhFooter>
        Exit Sub

GR2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.GR2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.GR2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()


F_ISXEIROGRAFO = Val(FINDPARAMETROI(1, "EPIT1", "F_ISXEIROGRAFO", "0", "Τσεκαρισμένο το χειρόγραφο=1 Οχι=0"))
If F_ISXEIROGRAFO = 1 Then
    ISXEIROGRAFO.Value = vbChecked
Else
    ISXEIROGRAFO.Value = vbUnchecked
End If

 F_EKTYPOTHS = FINDPARAMETROI(1, "PAR1", "F_EKTYPOTHS", "..", "Εκτυπωτής παραστατικών(με .. προεπιλεγμένος)") '
        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     mForm_Load Me, fh, fw, ft, fl

110     gBuff = "e"

120     Me.Picture = LoadPicture(gPicture)
130     hmee = Now
140     hmel = Now

150     PROTH = 1
160     fflag = 0
170     Me.KeyPreview = True

        Dim R As New ADODB.Recordset

180     R.Open "SELECT *FROM PINAKES WHERE TYPOS=6 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
190     Combo1.Clear

200     Do While Not R.EOF

210         If R("typos") = 6 Then
220             Combo1.AddItem Format(R("AYJON"), "#0") + " " + R("PERIGRAFH")  ' + Right(Space(30) + Format(R("TIMH"), "0000"), 30)
            End If

230         R.MoveNext
        Loop

240     R.Close
250     Combo1.Text = Combo1.List(0)

        'DATA2.DatabaseName = gDir
260     DATA2.ConnectionString = gConnect
270     DATA2.RecordSource = "select TOP 10 * from GRA"

        Dim DDD

280     Adodc1.ConnectionString = gConnect
290     Adodc1.RecordSource = "select * from XREOPIS where PEL ='" + gBuff + "'"
300     Adodc1.Refresh

        'If gXEIRISTHS = 8 Then
        '     Combo1.AddItem "ΧΡΕΩΣΗ ΚΑΡΤΕΛΑΣ.."
        '     Combo1.AddItem "ΠΙΣΤΩΣΗ ΚΑΡΤΕΛΑΣ.."
        'End If

        'Combo1.Text = Combo1.List(2) 'metrhta

        ' Text1.Text = DDD

        ' R.Close

        '<EhFooter>
        
        
        
          ' Dim R As New ADODB.Recordset

          R.Open "SELECT ARITMISI FROM XREOPIS WHERE EIDOS=71", Gdb, adOpenForwardOnly, adLockReadOnly
          If R.EOF Then
        
             Gdb.Execute "INSERT INTO XREOPIS (EIDOS,TITLOS,ARITMISI,EPIT,PEL) VALUES (71,'ΠΑΡΑΛΑΒΗ ΕΠΙΤΑΓΗΣ(MHX)',0,1,'e')"
          End If
          R.Close
        
         R.Open "SELECT ARITMISI FROM XREOPIS WHERE EIDOS=72", Gdb, adOpenForwardOnly, adLockReadOnly
          If R.EOF Then
        
             Gdb.Execute "INSERT INTO XREOPIS (EIDOS,TITLOS,ARITMISI,EPIT,PEL) VALUES (72,'ΠΑΡΑΔΟΣΗ ΕΠΙΤΑΓΗΣ(MHX)',0,1,'r')"
          End If
          R.Close
        
        
        
        
        
        
        
        
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Paint()

        '<EhHeader>
        On Error GoTo Form_Paint_Err

        '</EhHeader>
100     If PROTH = 1 Then
110         PROTH = 2
            'DATA1.DatabaseName = gDir
            'EGG.DatabaseName = gDir
120         Data1.ConnectionString = gConnect
130         Data1.RecordSource = "select KOD,EPO,STR(TYP,10,2),AYP,DOY,EPA,AFM,DIE from PEL WHERE EIDOS='e' and TYP>0"
140         Data1.Refresh
150         EGG.ConnectionString = gConnect
            'EGG.Refresh

        End If

        '<EhFooter>
        Exit Sub

Form_Paint_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.Form_Paint " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.Form_Paint " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub GR2_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo GR2_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
            'KeyCode = 0 'suppress the beep
110         GR2_Click
        End If

        '<EhFooter>
        Exit Sub

GR2_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.GR2_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.GR2_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub hmee_GotFocus()

        '<EhHeader>
        On Error GoTo hmee_GotFocus_Err

        '</EhHeader>

100     hmee.ToolTipText = "."

        '<EhFooter>
        Exit Sub

hmee_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.hmee_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.hmee_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub hmee_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo hmee_KeyDown_Err

        '</EhHeader>
100     If KeyCode = vbKeyReturn Then

        End If

        '<EhFooter>
        Exit Sub

hmee_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.hmee_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.hmee_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub hmee_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo hmee_KeyUp_Err

        '</EhHeader>
100     If KeyCode = vbKeyReturn Then
110         KeyCode = 0    'suppress the beep

120         If hmee.ToolTipText = "." Then
130             hmee.ToolTipText = ""
            Else
                'hmel.SetFocus
140             keybd_event VK_TAB, 0, 0, 0    'send a tab
            End If
        End If

        '<EhFooter>
        Exit Sub

hmee_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.hmee_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.hmee_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub hmel_GotFocus()
    '   hmel.ToolTipText = "."
End Sub

Private Sub hmel_KeyUp(KeyCode As Integer, Shift As Integer)
    ' If KeyCode = vbKeyReturn Then
    '      KeyCode = 0 'suppress the beep

    'If hmel.ToolTipText = "." Then
    '   hmel.ToolTipText = ""
    'Else
    '        keybd_event VK_TAB, 0, 0, 0 'send a tab
    'End If
    '  End If
End Sub

Private Sub Label4_Click()

        '<EhHeader>
        On Error GoTo Label4_Click_Err

        '</EhHeader>

        'If Index = 2 Then 'monades metrhshs
100     UPDATE_PINAKES 6, Label4, Combo1
        'End If

        '
        '
        '
        ''----------- OIKOGENEIES ---------------------------
        '  Dim f As New DataEntry2, DD
        '  Dim R As New ADODB.Recordset
        '
        '
        '
        '
        ''-----------------------------------------------------------------------------------------
        ''ΜΟΝΑΔΕΣ ΜΕΤΡΗΣΗΣ
        '  'sql οπου θα γίνεται το update ή το addnew
        ' '  f.SHOW
        '  Load f
        '
        '  Gdb.Execute "update PINAKES SET AYJON =0 WHERE AYJON IS NULL AND TYPOS=6"
        '  f.ParamGrid.TextMatrix(0, 1) = "SELECT * FROM PINAKES WHERE TYPOS=6 " '    sql
        ''τα πεδία που θα μπαίνουν sto ParamGrid και στα textboxes
        'f.ParamGrid.TextMatrix(0, 2) = "SELECT AYJON,PERIGRAFH " _
        '& " FROM PINAKES WHERE TYPOS=6 ORDER BY AYJON " 'ParamGrid   sql2
        '
        ''με ποιά εντολή θα κάνει το delete (το where ειναι το kod=...
        'f.ParamGrid.TextMatrix(0, 3) = "delete from PINAKES where TYPOS=6 AND " 'ParamGrid              DelSQL
        '
        ''ποιό update query θα δημιουργείται (προαιρετικά )
        ''αν δεν θέλω update τότε βάζω f.ParamGrid.TextMatrix(0, 4) = ""
        ''                    f.ParamGrid.TextMatrix(0, 4) = "  ""UPDATE ARITMISI2 SET ARITMISI= "" + t12 +"" where ID=""+t11 "   'UpdSQL
        ' DD = " UPDATE PINAKES SET TYPOS=6 where TYPOS IS NULL AND AYJON="" + t1   "
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
        'f.Left = Label4.Left
        'f.Top = Label4.Top + MDIForm1.Top
        '
        '
        '
        'f.SHOW 1
        'Combo1.Clear
        ' R.Open "SELECT *FROM PINAKES WHERE TYPOS=6 ORDER BY AYJON", Gdb, adOpenDynamic, _
        '     adLockOptimistic
        'Do While Not R.EOF
        '  Combo1.AddItem R("PERIGRAFH")
        '  R.MoveNext
        'Loop
        'R.Close
        '
        'Combo1.Refresh
        '

        '<EhFooter>
        Exit Sub

Label4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.Label4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.Label4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Label8_Click()

        '<EhHeader>
        On Error GoTo Label8_Click_Err

        '</EhHeader>

100     RichTextBox1.Height = 1455
110     RichTextBox1.Width = 3015

        '<EhFooter>
        Exit Sub

Label8_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.Label8_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.Label8_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Option1_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo Option1_Click_Err

        '</EhHeader>
        Dim x

100     If Option1(1) Then x = "e" Else x = "r"
110     gBuff = x
120     Option1_LostFocus Index

        '<EhFooter>
        Exit Sub

Option1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.Option1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.Option1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Option1_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = vbKeyReturn Then
110         KeyAscii = 0    'suppress the beep
120         keybd_event VK_TAB, 0, 0, 0    'send a tab
130         Text2_LostFocus 1
        End If

        '<EhFooter>
        Exit Sub

Option1_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.Option1_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.Option1_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Option1_LostFocus(Index As Integer)

        '<EhHeader>
        On Error GoTo Option1_LostFocus_Err

        '</EhHeader>
    
        Dim R As New ADODB.Recordset

        On Error Resume Next

100     R.Open "select * from XREOPIS where EPIT=1 AND PEL ='" + gBuff + "'", Gdb, adOpenDynamic, adLockOptimistic
    
110     Combo2.Clear

        '   R.MoveFirst
120     Do While Not R.EOF
130         Combo2.AddItem R("titlos") + Space(70) + str(R("EIDOS"))
140         R.MoveNext
        Loop

150    ' R.MoveFirst
        ' DDD = r("ARITMISI")
        Combo2.Text = Combo2.List(0)
        

        '<EhFooter>
        Exit Sub

Option1_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.Option1_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.Option1_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub parat_KeyPress(KeyAscii As Integer)
    '    If KeyAscii = vbKeyReturn Then
    '       KeyAscii = 0 'suppress the beep
    '       keybd_event VK_TAB, 0, 0, 0 'send a tab
    '    End If
End Sub

Private Sub poso_GotFocus()

        '<EhHeader>
        On Error GoTo poso_GotFocus_Err

        '</EhHeader>

100     POSO.BackColor = vbYellow

110     POSO.SelStart = 0
120     POSO.SelLength = Len(POSO.Text)

        '<EhFooter>
        Exit Sub

poso_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.poso_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.poso_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub poso_KeyPress(KeyAscii As Integer)
    '    If KeyAscii = vbKeyReturn Then
    '        KeyAscii = 0 'suppress the beep
    '        keybd_event VK_TAB, 0, 0, 0 'send a tab
    '    End If
End Sub

Private Sub poso_LostFocus()

        '<EhHeader>
        On Error GoTo poso_LostFocus_Err

        '</EhHeader>

100     POSO.Text = Replace(POSO.Text, ",", ".")
110     POSO.BackColor = vbWhite

        '<EhFooter>
        Exit Sub

poso_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.poso_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.poso_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub RichTextBox1_GotFocus()

        '<EhHeader>
        On Error GoTo RichTextBox1_GotFocus_Err

        '</EhHeader>

100     RichTextBox1.Height = 1455
110     RichTextBox1.Width = 3015

        '<EhFooter>
        Exit Sub

RichTextBox1_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.RichTextBox1_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.RichTextBox1_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text2_GotFocus(Index As Integer)

        '<EhHeader>
        On Error GoTo Text2_GotFocus_Err

        '</EhHeader>

100     Text2(Index).BackColor = vbYellow

        '<EhFooter>
        Exit Sub

Text2_GotFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.Text2_GotFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.Text2_GotFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text2_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
         KeyCode = 0    'suppress the beep
         keybd_event VK_TAB, 0, 0, 0    'send a tab
         
     End If

End Sub

Private Sub Text2_LostFocus(Index As Integer)

        '<EhHeader>
        On Error GoTo Text2_LostFocus_Err

        '</EhHeader>
        Dim x As String

100     Text2(Index).BackColor = vbWhite

110     If Len(Trim(Text2(Index))) = 0 Then Exit Sub

120     Update.Enabled = False

130     If Option1(1) Then x = "e" Else x = "r"
140     gBuff = x

150     If gUserId = 2 Then

160         If Index = 1 Then  'eponumo
170             Data1.RecordSource = "select KOD,EPO,STR(TYP,10,2),AYP,DOY,EPA,AFM,DIE from PEL where KOD NOT LIKE '997%' AND KOD NOT LIKE '998%' AND EIDOS='" + x + "' and EPO like '" + SameLetters(Text2(1).Text) + "%';"
            End If

180         If Index = 0 Then    'kodikos
190             Data1.RecordSource = "select KOD,EPO,STR(TYP,10,2),AYP,DOY,EPA,AFM,DIE from PEL where KOD NOT LIKE '997%' AND KOD NOT LIKE '998%' AND EIDOS='" + x + "' and KOD like '" + SameLetters(Text2(0).Text) + "%';"
            End If

200         If Index = 2 Then    'afm
210             Data1.RecordSource = "select KOD,EPO,STR(TYP,10,2),AYP,DOY,EPA,AFM,DIE from PEL where KOD NOT LIKE '997%' AND KOD NOT LIKE '998%' AND EIDOS='" + x + "' and AFM like '" + SameLetters(Text2(2).Text) + "%';"
            End If

        Else

220         If Index = 1 Then  'eponumo
230             Data1.RecordSource = "select KOD,EPO,STR(TYP,10,2),AYP,DOY,EPA,AFM,DIE from PEL where EIDOS='" + x + "' and EPO like '" + (SameLetters(Text2(1).Text)) + "%';"
            End If

240         If Index = 0 Then    'kodikos
250             Data1.RecordSource = "select KOD,EPO,STR(TYP,10,2),AYP,DOY,EPA,AFM,DIE from PEL where EIDOS='" + x + "' and KOD like '" + (SameLetters(Text2(0).Text)) + "%';"
            End If

260         If Index = 2 Then    'afm
270             Data1.RecordSource = "select KOD,EPO,STR(TYP,10,2),AYP,DOY,EPA,AFM,DIE from PEL where EIDOS='" + x + "' and AFM like '" + (SameLetters(Text2(2).Text)) + "%';"
            End If
        End If

280     Data1.Refresh
290     GR2.ColWidth(1) = 3800
        GR2.ColWidth(0) = 500
        'Data1.Recordset.MoveLast
        'Data1.Refresh

300     Data1.ConnectionString = gConnect

        '
        'x = "select KOD AS [Κωδ.],EPO as [Επωνυμία],STR(TYP,10,2) AS [Υπόλοιπο] ,THL as [Τηλέφωνο] ,DIE as [Διεύθυνση],STR(AYP,10,2)as [Αρχ.Υπόλ.] ,EIDOS,KODGAL as [Κατηγορία],LASTUPDT as [Τελ.Ενημέρωση],XRVMA from PEL where EIDOS='" + IIf(Combo1.Text = Combo1.List(0), "e", "r") + "'"
        'If Format(DTPicker1.Value, "DD/MM/YYYY") <> Format(Now, "DD/MM/YYYY") Then
        '   x = x + " and LASTUPDT <= '" + Format(DTPicker1.Value, "DD/MM/yyyy") + "' "
        'End If
        '
        'If IsNumeric(apo.Text) And IsNumeric(eos) Then
        '   x = x + " and  TYP >=" + apo.Text + " and TYP<=" + eos.Text
        'End If
        '
        'order = "KOD"
        'For k = 0 To 2
        '  If Combo2.Text = Combo2.List(0) Then
        '     order = "KOD"
        '  ElseIf Combo2.Text = Combo2.List(1) Then
        '     order = "EPO"
        '  ElseIf Combo2.Text = Combo2.List(2) Then
        '     order = "TYP"
        '  ElseIf Combo2.Text = Combo2.List(3) Then
        '     order = "LASTUPDT"
        '
        '  End If
        'Next

        '<EhFooter>
        Exit Sub

Text2_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.Text2_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.Text2_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Update_Click()

        '<EhHeader>
        On Error GoTo Update_Click_Err

        '</EhHeader>
        Dim R As New ADODB.Recordset

100     If hmee.Value > Now Then

            Dim ANS

110         ANS = MsgBox("Ημερομηνία καταχώρησης μεγαλύτερη της σημερινής. Συνεχίζω ;", vbYesNo)

120         If ANS = vbNo Then

                Exit Sub

            End If
        End If

130     Update.Enabled = False

140     DATA2.RecordSource = "select *from GRA where left(PAR,10)='" + Left(arepit, 10) + "'"

        On Error GoTo Update_Click_Err ' Resume Next

150     DATA2.Refresh

160     If DATA2.Recordset.RecordCount > 0 Then
170         MsgBox " Υπάρχει ξανά η επιταγή με ημερ.ληξ." + Format(DATA2.Recordset("HMEL"), "DD/MM/YYYY") + "-" + DATA2.Recordset("EPO")

180         Text2(0).SetFocus

            Exit Sub

        End If
    
'190     Combo2.Enabled = False
'200     Text1.Enabled = False

         ' Dim R As New ADODB.Recordset
If ISXEIROGRAFO.Value = vbChecked Then

Else
          R.Open "SELECT ISNULL(ARITMISI,0)+1  FROM XREOPIS  WHERE EIDOS=" + Right(Combo2.Text, 3), Gdb, adOpenForwardOnly, adLockReadOnly
          If Not R.EOF Then
              Text1.Text = R(0)
             
          End If
          R.Close
End If






    
'        Dim r3 As New ADODB.Recordset
'
'210     r3.Open "SELECT ARITMISI FROM XREOPIS WHERE EPIT=1", Gdb, adOpenForwardOnly, adLockReadOnly
'
'220     If r3.EOF Then
'
'230         Gdb.Execute "INSERT INTO XREOPIS (EIDOS,TITLOS,ARITMISI) VALUES (7,'ΠΑΡΑΛΑΒΗ ΕΠΙΤΑΓΗΣ',0)"
'            ' MsgBox "ΔΕΝ ΒΡΙΣΚΩ ΑΡΙΘΜΗΣΗ"
'            ' Exit Sub
'        End If
'
'        On Error Resume Next
'
'240     Gdb.Execute "UPDATE XREOPIS SET ARITMISI=1+" + str(Val(Text1.Text)) + "  WHERE EIDOS=" + Right(Combo2.Text, 3)

        ' DATA2.RecordSource = "gra"
        '
        '
        ' DATA2.Recordset.AddNew
        ' DATA2.Recordset("hmeE") = Format(hmee.Value, "DD/MM/YYYY")
        'DATA2.Recordset("hmel") = Format(hmel.Value, "DD/MM/YYYY")
        '
        ' DATA2.Recordset("pos") = Val(poso.Text)
        ' DATA2.Recordset("par") = arepit.Text
        '
        '
        ' DATA2.Recordset("kod") = DATA1.Recordset("kod")
        ' DATA2.Recordset("arx") = IIf(Option1(1), 1, 2)
        '
        'DATA2.Recordset("EPO") = left(DATA1.Recordset("EPO"), DATA2.Recordset("EPO").DefinedSize)
        '
        ' 'Data2.Recordset.Update
        ' DATA2.Recordset("EJO") = False

        ' Data2.Recordset.Update
        Dim X2

250     If Option1(1) Then X2 = "e" Else X2 = "r"

        Dim sql As String

260     sql = "INSERT INTO GRA (EJO,HMEL,HMEE,POS,PAR,EPO,KOD,ARX,KOD_BANK) "
270     sql = sql + " VALUES (0,'" + Format(hmel.Value, "MM/dd/YYYY") + "','" + Format(hmee.Value, "MM/dd/YYYY") + "',"

280     sql = sql + POSO.Text + ",'" + Left(arepit.Text + "  " + parat.Text, 30) + "','" + Left(f_epo, 15) + "','" + F_KODPEL + "'," + IIf(Option1(1), "1", "2") + ",'" + LTrim(Left(Combo1.Text, 2)) + "')"

        ' WHERE EIDOS='" + X2 + "' AND KOD='" + GR2.Text + "';"

        On Error Resume Next

290     Gdb.Execute sql, X2   '530.846 FAX  ELEYFTHEROS EPAGGELMATIAS

300     If X2 = 0 Then
310         MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ " + Chr(13) + Err.Description


            Exit Sub

        End If

320     R.Open "SELECT ID FROM GRA WHERE PAR= '" + Left(arepit.Text + "  " + parat.Text, 30) + "' AND HMEL='" + Format(hmel.Value, "MM/dd/YYYY") + "'", Gdb, adOpenDynamic, adLockOptimistic
        ' R.Open "SELECT ID FROM GRA WHERE PAR= '" + Left(arepit.Text + "  " + parat.Text, 30) + "' AND HMEL='" + Format(hmel.value, "MM/dd/YYYY") + "'", Gdb, adOpenDynamic, adLockOptimistic

330     sql = ""
340     sql = "INSERT INTO EGG (HME,XRE,AIT,KOD,EIDOS,APA,PISTOSI,XREOSI,EID,USERID,AAXREOPIS,IDGRA ) VALUES "
350     sql = sql + "('" + Format(hmee.Value, "MM/dd/YYYY")
360     sql = sql + "',-" + DOT(LTrim(POSO.Text)) + ",'"
370     sql = sql + Left(arepit.Text + Space(10), 10) + " ΛΗΞ" + Format(hmel.Value, "DD/MM/YY") + "','"
380     sql = sql + F_KODPEL + "','"
390     sql = sql + f_buff + "','"
400     sql = sql + str(Val(Text1.Text)) + "',"  ' APA ΑΡΙΘΜΟΣ ΑΠΟΔΕΙΞΗΣ
410     sql = sql + IIf(f_buff = "e", DOT(LTrim(POSO.Text)), "0") + ","
420     sql = sql + IIf(f_buff = "e", "0", DOT(LTrim(POSO.Text))) + ",'ΕΠ',"
430     sql = sql + str(gUserId) + ","
    
440     sql = sql + IIf(f_buff = "e", "16", "7") + ","   ' ΚΩΔΙΚΟς ΚΙΝΗΣΗς
    
450     If IsNull(R(0)) Then
460         sql = sql + str(0) + ")"
        Else
470         sql = sql + str(R(0)) + ")"
        End If

480     Gdb.Execute sql, X2

490     If X2 = 0 Then
500         MsgBox "ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ Η ΚΑΡΤΕΛΛΑ ΤΟΥ ΛΟΓΑΡΙΑΣΜΟΥ "

            Exit Sub

        End If

510     MsgBox "ΕΝΗΜΕΡΩΘΗΚΕ Η ΕΠΙΤΑΓΗ " + Chr(13) + "ΚΑΙ Η ΚΑΡΤΕΛΛΑ ΤΟΥ ΛΟΓΑΡΙΑΣΜΟΥ "
        '  Data2.RecordSource = "select *from GRA where left(PAR,8)='" + arepit + "'"
        ' Data2.Refresh

520     GR2_Click

        Dim k As Long, l As Long

        'For k = 1 To GR2.rows
        '        GR1.row = k
        '        For l = 0 To GR1.cols - 1
        '             GR1.col = l: GR1.CellBackColor = vbWhite
        '         Next
        'Next k

530     ' Text2(0).SetFocus
        ' ΔΕΙΧΝΩ ΤΗΝ ΚΑΡΤΕΛΑ ΤΟΥ ΠΕΛΑΤΗ ΑΛΛΑ ΜΟΥ ΚΑΝΕΙ Update.Enabled = True
        hmee.SetFocus
         'ΤΟ ΞΑΝΑΚΑΝΩ FALSE ΓΙΑ ΝΑ ΜΗΝ ΤΟ ΞΑΝΑΠΑΤΗΣΕΙ
         Update.Enabled = False
         
        Command2.Enabled = True
        Combo2.Enabled = True
        
       
        
        
        
        '<EhFooter>
        Exit Sub

Update_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.EPIT1.Update_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.EPIT1.Update_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
