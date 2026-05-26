VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form bohu7 
   BackColor       =   &H00FF0000&
   Caption         =   "Απογραφές-Πόντοι"
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14355
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10440
   ScaleWidth      =   14355
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   "ΛΗΨΗ ΑΠΟΓ ΑΠΟ EXCEL"
      Height          =   495
      Left            =   9360
      TabIndex        =   39
      ToolTipText     =   "3 ΣΤΗΛΕΣ ΜΕ : ΚΩΔΙΚΟΣ,ΠΟΣΟΤΗΤΑ,ΤΙΜΜ,ΠΕΡΙΓΡΑΦΗ"
      Top             =   2280
      Width           =   2340
   End
   Begin VB.CommandButton Command7 
      Caption         =   "ΑΝΑΓΩΓΗ ΑΠΟΓΡΑΦΗΣ ΣΕ 1/1"
      Height          =   735
      Left            =   9240
      TabIndex        =   38
      Top             =   880
      Width           =   2340
   End
   Begin VB.CommandButton Command10 
      Height          =   285
      Left            =   11160
      Picture         =   "bohu7.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1965
      Width           =   540
   End
   Begin VB.CommandButton CALC 
      Caption         =   "ΛΗΨΗ ΑΠΟΓ ΑΠΟ CALC OPENOFFICE"
      Height          =   495
      Left            =   9384
      TabIndex        =   36
      ToolTipText     =   "3 ΣΤΗΛΕΣ ΜΕ : ΚΩΔΙΚΟΣ,ΠΟΣΟΤΗΤΑ,ΤΙΜΜ"
      Top             =   2880
      Width           =   2316
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5460
      Top             =   3276
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command9 
      Caption         =   "ΑΠΟ SYMBOL"
      Height          =   255
      Left            =   6720
      TabIndex        =   35
      ToolTipText     =   "APOG.TXT"
      Top             =   2880
      Width           =   2316
   End
   Begin VB.ComboBox APOT2 
      BackColor       =   &H0080FF80&
      Height          =   288
      Left            =   5541
      TabIndex        =   34
      Top             =   6084
      Width           =   1335
   End
   Begin VB.CheckBox AutoDelete 
      BackColor       =   &H00FF0000&
      Caption         =   "Αυτόματο σβήσιμο παλιών"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5460
      TabIndex        =   33
      Top             =   4440
      Width           =   2610
   End
   Begin VB.CheckBox prosarmogh 
      Caption         =   "Προσαρμογή τρέχοντος υπολοίπου σύμφωνα με πρόχειρη απογραφή"
      Height          =   1080
      Left            =   11880
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   31
      Top             =   2280
      Width           =   2340
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "bohu7.frx":0102
      Height          =   30
      Left            =   60
      TabIndex        =   30
      Top             =   2805
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   53
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1560
      Top             =   9840
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin VB.CommandButton Command6 
      Caption         =   "Ειδη Πελάτου"
      Height          =   570
      Left            =   2850
      TabIndex        =   29
      Top             =   6510
      Width           =   2304
   End
   Begin VB.ComboBox apot 
      Height          =   288
      Left            =   7488
      TabIndex        =   28
      Top             =   1320
      Width           =   1572
   End
   Begin VB.CommandButton mhdenismos 
      Caption         =   "Μηδενισμός Αποθήκης με κινήσεις πρόχειρης απογραφής"
      Height          =   735
      Left            =   9240
      TabIndex        =   27
      ToolTipText     =   "Θα μηδενιστεί η παρακάτω αποθήκη με την παρακάτω ημερομηνία Απογραφής "
      Top             =   120
      Width           =   2340
   End
   Begin VB.CommandButton NEABARCODES 
      BackColor       =   &H00FFFF80&
      Caption         =   "ΝΕΑ BARCODES"
      Height          =   612
      Left            =   2850
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4230
      Width           =   2304
   End
   Begin VB.ListBox List2 
      Height          =   2205
      Left            =   228
      TabIndex        =   25
      Top             =   7305
      Width           =   6648
   End
   Begin MSComCtl2.DTPicker apog 
      Height          =   276
      Left            =   7824
      TabIndex        =   23
      Top             =   936
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      Format          =   435552257
      CurrentDate     =   38727
   End
   Begin VB.TextBox FILENAME 
      Height          =   285
      Left            =   9360
      TabIndex        =   21
      Top             =   1965
      Width           =   1980
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ΛΗΨΗ ΑΠΟΓΡΑΦΗΣ (APOG.DBF)"
      Height          =   735
      Left            =   6720
      TabIndex        =   20
      Top             =   120
      Width           =   2316
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2850
      TabIndex        =   19
      Text            =   "\\server\c\tameia\diatakt.mdb"
      Top             =   5304
      Width           =   2304
   End
   Begin VB.CommandButton Command4 
      Caption         =   "E ξ ο δ ο ς"
      Height          =   615
      Left            =   9360
      TabIndex        =   17
      Top             =   6864
      Width           =   2340
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "Λήψη κινήσεων από υποκατάστημα"
      Height          =   615
      Left            =   2850
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "C:\LAGEURO\SEND"
      Top             =   5790
      Width           =   2304
   End
   Begin VB.CommandButton SendKiniseis 
      BackColor       =   &H0080FF80&
      Caption         =   "Αποστολή Κινήσεων"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Αρχείο  c:\lageuro\send\eggtim.dbf"
      Top             =   5790
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Αποστολή πόντων"
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   4950
      Width           =   2295
   End
   Begin VB.CommandButton neesTimes 
      BackColor       =   &H00FFFF80&
      Caption         =   "Ενημέρωση με νέες τιμές"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Το αρχείο που θα φορτωθεί είναι το C:\LAGEURO\EID.DBF"
      Top             =   4230
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ενημέρωση από ΔΙΑΤΑΚΤΙΚΕΣ"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   3270
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   480
      TabIndex        =   6
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   1248
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker APO 
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   6510
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   435552257
      CurrentDate     =   38604
   End
   Begin MSComCtl2.DTPicker EOS 
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   6870
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   435552257
      CurrentDate     =   38604
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Λήψη απογραφής από εξωτ.αρχείο"
      Height          =   1935
      Left            =   9240
      TabIndex        =   40
      Top             =   1680
      Width           =   5055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Κωδ.Υποκατασ"
      ForeColor       =   &H8000000E&
      Height          =   336
      Left            =   5460
      TabIndex        =   32
      Top             =   5772
      Width           =   1128
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Αποθήκη"
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
      Height          =   348
      Left            =   6708
      TabIndex        =   24
      Top             =   1320
      Width           =   852
   End
   Begin VB.Label Label9 
      Caption         =   "Ημ.Απογραφής"
      Height          =   276
      Left            =   6732
      TabIndex        =   22
      Top             =   936
      Width           =   1044
   End
   Begin VB.Label Label8 
      Caption         =   "Διαδρομή Πόντων"
      Height          =   252
      Left            =   2850
      TabIndex        =   18
      Top             =   4992
      Width           =   2304
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ΑΠΟ"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   6510
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ΕΩΣ"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   6870
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Χειροκίνητο πέρασμα πόντων"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Επιπλέον πόντοι"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Πόντοι"
      Height          =   372
      Left            =   2724
      TabIndex        =   4
      Top             =   720
      Width           =   1368
   End
   Begin VB.Label Label2 
      Caption         =   "Βarcode κάρτας"
      Height          =   372
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Height          =   372
      Left            =   2652
      TabIndex        =   1
      Top             =   1248
      Width           =   1440
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   2676
      Left            =   60
      Top             =   99
      Width           =   6552
   End
End
Attribute VB_Name = "bohu7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim F_TAKENEESTIMES As Integer

Dim F_EXO_XROMATA   As Integer  '= Val(FindParametroi(1,"PAR1", "F_EXO_XROMATA", "0", "Γ.1 ΧΡΗΣΙΜΟΠΟΙΩ ΜΕΓΕΘΗ-ΧΡΩΜΑΤΑ=1")) 'posa psifia tha exei h kathe seira

Dim fr              As New ADODB.Recordset

Dim F_APO           As String

Dim F_EOS           As String

Dim gconn2 As New ADODB.Connection

Dim F_GCONN2 As String
Dim F_GCONN3 As String
Dim F_GCONN4 As String




Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Private Sub Command1_Click()

        ' ENHMERCSH PONTON KYRIOY ARXEIOY
        'Dim r As New ADODB.Recordset
        'Dim r2 As New ADODB.Recordset
        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>
        Dim dB   As DAO.Database

        Dim R    As DAO.Recordset

        Dim RADO As New ADODB.Recordset

        On Error GoTo CREATEDIATAKT

100     RADO.Open "SELECT TOP 10 * FROM DIATAKT WHERE  HME >='" + Format(apo, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' ", Gdb, adOpenDynamic, adLockOptimistic

        On Error Resume Next

        Dim sql As String

110     If RADO.EOF Then
            ' OK
        Else
120         MsgBox "ΥΠΑΡΧΟΥΝ ΗΔΗ ΟΙ ΕΓΓΡΑΦΕΣ ΤΩΝ ΠΟΝΤΩΝ"

            Exit Sub

        End If

        'Set db = OpenDatabase("C:\LAGEURO\SEND", False, False, "dBase III;")
130     Set dB = OpenDatabase("C:\LAGEURO\SEND", False, False, "dBase III;")

140     XX = "SELECT * FROM DIATAKT WHERE  HME >=#" + Format(apo, "MM/DD/YYYY") + "# AND HME<#" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "# ORDER BY ΗΜΕ "
150     Set R = dB.OpenRecordset("DIATAKT")
160     R.MoveFirst

170     Do While Not R.EOF

180         If IsNull(R("PONTOI")) Or IsNull(R("PELATHS")) Then
                'ASE
            Else
                '
190             sql = "INSERT INTO DIATAKT (PELATHS,PONTOI,HME) " & "VALUES ('" + R("PELATHS") + "'," + str(R("PONTOI")) + ",'" + Format(R("HME"), "MM/DD/YYYY") + "')"

                On Error GoTo Command1_Click_Err

200             Gdb.Execute sql
210             Gdb.Execute "UPDATE EID SET PONTOI=PONTOI+" + str(R("PONTOI")) + " WHERE LEFT(KOD,13)='" + R("PELATHS") + "'  "
            End If

220         Me.Caption = R("hme")
230         R.MoveNext

240         DoEvents

        Loop

250     R.Close
260     dB.Close

        Exit Sub

CREATEDIATAKT:

270     Gdb.Execute "CREATE TABLE DIATAKT (PELATHS CHAR(13),PONTOI DECIMAL(10,2),HME DATETIME  )"

        Exit Sub

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub update_DBF_from_SQL(arxeio As String, ARXEIOdbf As String, ByVal gg As ADODB.Connection)

        '<EhHeader>
        On Error GoTo update_DBF_from_SQL_Err

        '</EhHeader>

        Dim DBF    As Database

        Dim sql    As New ADODB.Connection

        Dim rDBF   As Recordset

        Dim rSQL   As New ADODB.Recordset

        Dim conDBF As String

        Dim conSQL As String

        Dim k      As Long

        Dim Fname  As String

        Dim dB     As DAO.Database

100     Set dB = OpenDatabase("C:\LAGEURO\SEND", False, False, "dBase III;")

110     Set rDBF = dB.OpenRecordset(ARXEIOdbf)

        Dim N As Long
        rSQL.Open "select count(*) from " + arxeio, gg, adOpenDynamic, adLockOptimistic

        If rSQL(0) = 0 Then
            Me.Caption = "pos1 0 ΕΓΓΡΑΦΕΣ"
            GoTo mexit
        End If
        rSQL.Close
        Me.Caption = "pos1 συνδεθηκε"


120     rSQL.Open "SELECT *FROM " + arxeio, gg, adOpenDynamic, adLockOptimistic
        ' rDBF.Open "SELECT *FROM " + arxeio, DBF, adOpenDynamic, adLockOptimistic

       ' On Error Resume Next

130     rSQL.MoveFirst

        Dim Z

140     Z = 0

        On Error GoTo WRITEERROR    ' Resume Next

150     Do While Not rSQL.EOF

160         rDBF.AddNew

170         For k = 0 To rDBF.FIELDS.Count - 1
180             Fname = rDBF.FIELDS(k).Name    ' p.x. FNAME=epo    R(0).NAME

                '      If IsNull(rDBF(k)) Then
                '      Else

                '      End If

190             If Fname = "ATIM" Then
200                 rDBF(Fname) = Left(rSQL(Fname), 7)    'ΓΙΑΤΙ ΣΕ ΜΕΓΑΛΟ ΑΤΙΜ ΔΕΝ ΓΡΑΦΕΙ ΤΙΠΟΤΑ
                Else
210                 rDBF(Fname) = rSQL(Fname)
                End If

            Next

220         rDBF.Update
230         Z = Z + 1

240         If Z Mod 10 = 0 Then
250             Me.Caption = Z
            End If

260         rSQL.MoveNext

270         DoEvents

        Loop

280     rSQL.Close
290     rDBF.Close
         
         List2.AddItem str(N) + "  " + gg
         
        dB.Execute "update EGGTIM SET FPA=2 WHERE FPA=24"
300     dB.Execute "update EGGTIM SET FPA=2 WHERE FPA=23"    ' 19%
310     dB.Execute "update EGGTIM SET FPA=4 WHERE FPA=21"    ' 4.5%
320     dB.Execute "update EGGTIM SET FPA=2 WHERE FPA=22"    ' 9%

330     Set rDBF = dB.OpenRecordset("SELECT SUM(MIK_AJIA) FROM EGGTIM")
340     MsgBox "Σύνολο αξίας " + Format(rDBF(0), "####,##0.00")


mexit:


        Exit Sub

WRITEERROR:
        'List2.Visible = True

        'List2.AddItem Err.Description
350     Resume Next

        '<EhFooter>
        Exit Sub

update_DBF_from_SQL_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.update_DBF_from_SQL " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
         List2.AddItem Err.Description + "at line " & Erl
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.update_DBF_from_SQL " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command10_Click()

        '<EhHeader>
        On Error GoTo Command10_Click_Err

        '</EhHeader>
100     If Len(Trim(Text1.Text)) = 0 Then
110         cd1.ShowOpen
120         FILENAME.Text = cd1.FILENAME
        Else

130         If Len(Dir(LTrim(FILENAME.Text), vbNormal)) < 2 Then
140             MsgBox "δεν υπάρχει το αρχείο " + Text1.Text

                Exit Sub

            End If
        End If

        'On Error GoTo open_error
150     Me.MousePointer = vbHourglass
        'Set xlwbook = xl.Workbooks.Open(Text1.Text)
        'Set xlsheet = xlwbook.Sheets.Item(1)

        On Error GoTo Command10_Click_Err

160     Me.MousePointer = vbNormal

        Exit Sub

        '<EhFooter>
        Exit Sub

Command10_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Command10_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Command10_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command11_Click()

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>

        ' etsi htan 10/8/2012 Set db = OpenDatabase(Text3.Text)
    
100     Set dB = OpenDatabase(Text3.Text, , ReadOnly)

        On Error Resume Next

110     Kill "C:\LAGEURO\SEND\DIATAKT.DBF"

120     dB.Execute "select * into DIATAKT in 'C:\LAGEURO\SEND' 'dBase III;' from diataktikes where  HME>=#" + Format(apo, "MM/DD/YYYY") + "# AND HME<#" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "#"

130     Set dB = OpenDatabase("C:\LAGEURO\SEND", False, False, "dBase III;")

140     Set R = dB.OpenRecordset("SELECT SUM(pontoi) FROM diatakt")
150     MsgBox "Σύνολο Πόντων  " + Chr(13) + Format(R(0), "####,##0")

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>
        Dim mYpok

100     mYpok = LTrim(str(Val(Left(APOT2.Text, 2))))    ' LTrim(Trim(Text4.Text))

        Dim DBF    As Database

        Dim sql    As New ADODB.Connection

        Dim rDBF   As Recordset

        Dim rSQL   As New ADODB.Recordset

        Dim conDBF As String

        Dim conSQL As String

        Dim k      As Long

        Dim Fname  As String

        Dim dB     As DAO.Database

        Dim arxeio As String, arxeio2 As String

        Dim ANS    As Integer

110     arxeio = "EGGTIM WHERE KOLA=" + mYpok + " and HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'"
120     ARXEIOdbf = "EGGTIM"

     
  '   If Dir("C:\LAGEURO\SEND\tim.mdb") Then
   '          Set DB = OpenDatabase("C:\LAGEURO\SEND\" + mYpok + "\tim.mdb")
   '  Else
             Set dB = OpenDatabase("C:\LAGEURO\SEND\" + mYpok, False, False, "dBase III;")
    ' End If
     
     
         
        Set dB = OpenDatabase("C:\LAGEURO\SEND\tim.mdb")
        'Set rDBF = db.OpenRecordset("SELECT SUM(MIK_AJIA) FROM EGGTIM where apot=" + mYpok)

       ' On Error Resume Next

140     Set rDBF = dB.OpenRecordset("SELECT COUNT(*) FROM EGGTIM where apot=" + mYpok + " AND HME>=#" + Format(apo, "MM/DD/YYYY") + "# AND HME<#" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "#")

150     If rDBF(0) = 0 Then
160         MsgBox "ΔΕΝ ΥΠΑΡΧΟΥΝ ΚΙΝΗΣΕΙΣ - ΔΙΑΚΟΠΗ ΔΙΑΔΙΚΑΣΙΑΣ"

            Exit Sub

        End If

170     rDBF.Close

180     Set rDBF = dB.OpenRecordset("SELECT * FROM EGGTIM WHERE HME>=#" + Format(apo, "MM/DD/YYYY") + "# AND HME<#" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "#")
        ' ελεγχος αν υπάρχουν ήδη τα παραστατικά

        ' arxeio = "EGGTIM WHERE APOT=" + mYpok + " AND LEFT(ATIM,1)='L' and HME>='" + Format(APO, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY") + "'"
190     rSQL.Open "SELECT count(*) FROM " + arxeio, Gdb, adOpenDynamic, adLockOptimistic

200     If rSQL(0) > 0 Then

            'MsgBox "υπάρχουν ήδη  " + Str(rSQL(0)) + "κινήσεις.   Δεν θα ενημερωθεί το αρχείο"
            'rSQL.Close
            ' αν πρόκειται για μία μέρα μόνο να έχει δυνατότητα διαγραφής
210         If apo <= eos Then
220             ANS = MsgBox("Να διαγράψω τις ήδη υπάρχουσες κινήσεις του υποκαταστήματος ΑΠΟ " + Format(apo, "dd/mm/yyyy") + " ΕΩΣ " + Format(eos, "dd/mm/yyyy"), vbYesNo)

230             If ANS = vbNo Then

                    Exit Sub

                End If

                Dim d

                ' antikaθιστα τις 3 παλιότερες σειρές 5-5-2007
                'ΣΒΗΝΕΙ ΟΛΑ ΑΥΤΑ ΠΟΥ ΠΡΟΗΡΘΑΝ ΑΠΟ ΤΟ ΥΠΟΚΑΤΑΣΤΗΜΑ Νο mYPOK
240             Gdb.Execute "DELETE FROM EGGTIM WHERE KOLA=" + mYpok + " AND HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'", d

                'PALIOTERA APO 5-5-2007
                '      Gdb.Execute "DELETE FROM EGGTIM WHERE (NOT LEFT(ATIM,1)='σ')  AND APOT=" + mYpok + " AND HME>='" + Format(APO, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY") + "'", d
                '
                '      Gdb.Execute "DELETE FROM EGGTIM WHERE  LEFT(ATIM,1)='σ' AND APOT=" + mYpok + " and XRE<0 AND HME>='" + Format(APO, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY") + "'", d
                '
                '      Gdb.Execute "DELETE FROM EGGTIM WHERE  LEFT(ATIM,1)='σ' AND APOT=1 and XRE>0 AND HME>='" + Format(APO, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY") + "'", d
                '

250             MsgBox "Διεγράφησαν " + Format(d, "####") + " εγγραφές "
            Else
260             MsgBox "υπάρχουν ήδη  " + str(rSQL(0)) + "κινήσεις.   Δεν θα ενημερωθεί το αρχείο"
270             rSQL.Close

                Exit Sub

            End If

        End If

280     rSQL.Close
290     rSQL.Open "SELECT * FROM EGGTIM ", Gdb, adOpenDynamic, adLockOptimistic

300     rDBF.MoveFirst

        Dim Z

310     Z = 0

        On Error GoTo WRITEERROR    ' Resume Next

320     Do While Not rDBF.EOF

330         rSQL.AddNew

340         For k = 0 To rDBF.FIELDS.Count - 1
350             Fname = rDBF.FIELDS(k).Name    ' p.x. FNAME=epo    R(0).NAME

360             If IsNull(rDBF(k)) Then
                Else
370                 rSQL(Fname) = rDBF(Fname)    ' rsqk("epo")=r(0)
                End If

            Next

380         If Left(rSQL("ATIM"), 1) = "L" Then
                'επειδη των ταμειακων δεν εχουν "xre" kai "pis"
390             rSQL("PIS") = 0: rSQL("XRE") = 0
400             rSQL("PIS") = rSQL("POSO")

            End If

410         rSQL("kola") = Val(mYpok)    ' σημαινει ότι ηρθαν απο το υποκαταστημα

            ' ΟΤΑΝ ΓΙΝΕΤΑΙ ΕΠΙΣΤΡΟΦΗ ΑΠΟ ΤΗΝ ΤΑΜΕΙΑΚΗ ΕΧΕΙ ΑΡΝΗΤΙΚΗ ΤΙΜΗ "ΤΙΜΜ"
            ' ΟΠΟΤΕ ΤΗΝ ΚΑΝΩ ΘΕΤΙΚΗ ΚΑΙ ΚΑΝΩ ΑΡΝΗΤΙΚΗ ΤΗΝ ΠΟΣΟΤΗΤΑ
420         If rSQL("TIMM") < 0 Then
430             rSQL("TIMM") = -rSQL("TIMM")
440             rSQL("POSO") = -Abs(rSQL("POSO"))
450             rSQL("XRE") = 0
460             rSQL("PIS") = rSQL("POSO")
            End If

470         rSQL.Update
480         Z = Z + 1

490         If Z Mod 10 = 0 Then
500             Me.Caption = Z
            End If

510         rDBF.MoveNext

520         DoEvents

        Loop

530     rSQL.Close
540     rDBF.Close

        '--------------------μεταφορα ΤΙΜ ------------------------------------------------------------
        '--------------- συντομος τρόπος με προβλήματα στο KPE ΔΕΝ ΤΟ ΦΈΡΝΕΙ ------------------------
        'Dim Sql2 As String
        'Sql2 = COLLECT_FIELDS()
        'Dim db2
        'Set db2 = OpenDatabase("", False, False, gConnect)
        'Sql2 = "INSERT INTO TIM (" + Sql2 + ") SELECT " + Sql2 + " FROM  TIM IN 'C:\LAGEURO\SEND\TIM.MDB'"
        'db2.Execute Sql2
        '--------------- συντομος τρόπος με προβλήματα στο KPE ΔΕΝ ΤΟ ΦΈΡΝΕΙ ------------------------

        '--------------- κανονικος τρόπος με προβλήματα στο KPE ΔΕΝ ΤΟ ΦΈΡΝΕΙ ------------------------
        'Set DB = OpenDatabase("C:\LAGEURO\SEND\tim.mdb")
        'Set rDBF = DB.OpenRecordset("SELECT * FROM TIM WHERE HME>=#" + Format(APO, "MM/DD/YYYY") + "# AND HME<#" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY") + "#")
        'rSQL.Open "SELECT *FROM TIM WHERE EIDOS='@' ", Gdb, adOpenDynamic, adLockOptimistic
        'Do While Not rDBF.EOF
        '  rSQL.AddNew
        '  For k = 0 To rDBF.FIELDS.Count - 1
        '      Fname = rDBF.FIELDS(k).Name ' p.x. FNAME=epo    R(0).NAME
        '      If IsNull(rDBF(k)) Then
        '      Else
        '         rSQL(Fname) = rDBF(Fname) ' rsqk("epo")=r(0)
        '      End If
        '  Next
        '  rSQL.Update
        '  Z = Z + 1
        '  If Z Mod 10 = 0 Then
        '      Me.Caption = Z
        '  End If
        '  rDBF.MoveNext
        '  DoEvents
        'Loop
        'rSQL.Close
        'rDBF.Close
        '--------------- κανονικος τρόπος με προβλήματα στο KPE ΔΕΝ ΤΟ ΦΈΡΝΕΙ ------------------------

550     MsgBox "Κατεχωρήθησαν " + str(Z) + " εγγραφές"
        'ενημερωνει τις ημερομηνίες του τελευταίου update
560     Gdb.Execute "UPDATE PARAMETROI SET TIMH='" + Format(apo, "DD/MM/YYYY") + "' WHERE FORMA='BOHU7' AND VAR='F_APO'"
570     Gdb.Execute "UPDATE PARAMETROI SET TIMH='" + Format(eos, "DD/MM/YYYY") + "' WHERE FORMA='BOHU7' AND VAR='F_EOS'"

        Exit Sub

WRITEERROR:

580     Resume Next

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command4_Click()

        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>

100     Unload Me

110     If gApoMenu = True Then
120         SendKeys "%"
130         SendKeys "{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}"
        End If

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command5_Click()

        ' ΛΗΨΗ ΑΠΟΓΡΑΦΗΣ ============================
        '<EhHeader>
        On Error GoTo Command5_Click_Err

        '</EhHeader>

        Dim DBF    As Database

        Dim sql    As New ADODB.Connection

        Dim rDBF   As Recordset

        Dim rSQL   As New ADODB.Recordset

        Dim rSQL2  As New ADODB.Recordset

        Dim rYPOL  As New ADODB.Recordset

        Dim mYPOL  As Single

        Dim conDBF As String

        Dim conSQL As String

        Dim k      As Long

        Dim Fname  As String

        Dim dB     As DAO.Database

        Dim arxeio As String, arxeio2 As String

        Dim ANS    As Integer

        Dim M_ATIM

100     M_ATIM = "λ" + Format(time(), "hhmmss")

110     If Len(Dir(FILENAME.Text, vbNormal)) > 0 Then
            ' OK
        Else
120         MsgBox " ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ ΑΡΧΕΙΟ " + FILENAME.Text

            Exit Sub

        End If

130     If Val(apot) = 0 Then
140         MsgBox " ΔΕΝ ΥΠΑΡΧΕΙ ΑΠΟΘΗΚΗ "

            Exit Sub

        End If

150     Set dB = OpenDatabase(Dir1.Path, False, False, "dBase III;")

160     Set rDBF = dB.OpenRecordset("SELECT * FROM " + File1.FILENAME)

170     rSQL.Open "SELECT * FROM EGGTIM ", Gdb, adOpenDynamic, adLockOptimistic

180     rDBF.MoveFirst

        Dim Z

190     Z = 0
200     k = 0

        On Error GoTo WRITEERROR    ' Resume Next

        Dim OK As Boolean

        Dim M_CODE

210     Do While Not rDBF.EOF
220         OK = True

230         M_CODE = rDBF("barcode")

240         If kodisbarcode.Value Then

            Else

250             rSQL2.Open "SELECT * FROM BARCODES WHERE LEFT(ERG,13)='" + Left(rDBF("BARCODE"), 13) + "'", Gdb, adOpenForwardOnly, adLockReadOnly

260             If rSQL2.EOF Then
270                 OK = False
280                 List2.AddItem rDBF("BARCODE") + ": ΔΕΝ ΥΠΑΡΧΕΙ ΣΤΟ ΑΡΧΕΙΟ ΕΙΔΩΝ"
                Else
290                 M_CODE = rSQL2("KOD")
                End If

300             rSQL2.Close
            End If

310         If OK Then    'And (Not IsNull(rDBF("YPOK")))'==========================================================================
320             If Val(rDBF("YPOK")) <> 33 Then
330                 mYPOL = 0

                    '-----------------------------------------------
340                 If prosarmogh.Value = vbChecked Then
                        'ΒΡΙΣΚΩ ΤΟ ΥΠΟΛΟΙΠΟ ΠΟΥ ΕΙΧΕ ΜΕΧΡΙ ΧΘΕΣ
350                     rYPOL.Open "SELECT KODE,SUM( CASE APOT  WHEN " + Left(apot.Text, 1) + "  THEN XRE ELSE 0  END ) AS S1X," & " SUM(CASE APOT  WHEN " + Left(apot.Text, 1) + " THEN PIS ELSE 0  END ) AS S1P FROM EGGTIM" & " WHERE  HME<='" + Format(apog.Value, "MM/DD/YYYY") + "' AND KODE='" + M_CODE + "' AND APOT=" + Left(apot.Text, 1) + "GROUP BY KODE", Gdb, adOpenForwardOnly, adLockReadOnly

360                     If rYPOL.EOF Then
                        Else
370                         mYPOL = IIf(IsNull(rYPOL(1)), 0, rYPOL(1)) - IIf(IsNull(rYPOL(2)), 0, rYPOL(2))
                        End If

380                     rYPOL.Close
                        'ΕΠΕΙΔΗ υπάρχει το ενδεχόμενο
                        ' το ίδιο είδος να υπήρχε σε 2 σημεία θα το διορθώσει με το mYPOL μόνο την
                        ' πρώτη φορά
                        'π.χ. απογραφω το είδος Χ μια φορά 1τεμ και 2η φορά 3τεμ
                        'με την πρώτη φορά μηδενίζω το υπόλοιπο που είχε -mypol και κάνω το πραγματικό υπόλοιπο =1
                        'με την 2η φορά
                        ' ΔΕΝ ΠΡΕΠΕΙ ΝΑ μηδενίσω το υπόλοιπο αλλα να προσθέσω +3
390                     rYPOL.Open "SELECT COUNT(*) FROM EGGTIM WHERE LEFT(ATIM,1)='λ' AND HME >='" + Format(apog.Value, "MM/DD/YYYY") + "' AND KODE='" + M_CODE + "' AND APOT=" + Left(apot.Text, 1), Gdb, adOpenForwardOnly, adLockReadOnly

400                     If rYPOL(0) > 0 Then
410                         mYPOL = 0
                        End If

420                     rYPOL.Close
                    End If

                    '-----------------------------------------------
                    On Error Resume Next

430                 rSQL.AddNew
440                 rSQL("hme") = apog.Value
450                 rSQL("ATIM") = M_ATIM    '"λ00002"
460                 rSQL("KODE") = M_CODE
470                 rSQL("timm") = rDBF("xti")
480                 rSQL("poso") = rDBF("POS") - mYPOL
490                 rSQL("xre") = rDBF("POS") - mYPOL
500                 rSQL("APOT") = Val(apot.Text)
510                 rSQL.Update
520                 k = k + 1
                End If   '     If Val(rDBF("YPOK")) <> 33 Then
            End If  '==========================================================================

            ' rSQL.Update
530         Z = Z + 1

540         If Z Mod 10 = 0 Then
550             Me.Caption = Z
            End If

560         rDBF.MoveNext

570         DoEvents

        Loop

580     rSQL.Close
590     rDBF.Close

        On Error Resume Next

600     rSQL.Open "SELECT * FROM TIM ", Gdb, adOpenDynamic, adLockOptimistic

610     rSQL.AddNew
620     rSQL("hme") = apog.Value
630     rSQL("ATIM") = M_ATIM    ' "λ00001"
        'rSQL("KODE") = M_CODE
        'rSQL("timm") = rDBF("xti")
        ' rSQL("poso") = rDBF("POS") - mYPOL
        ' rSQL("xre") = rDBF("POS") - mYPOL
640     rSQL("b_n1") = Val(apot.Text)
650     rSQL.Update

660     MsgBox "Ενημερώθηκαν " + Format(k, "###0") + " εγγραφές"

        Exit Sub

WRITEERROR:

670     Resume Next

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command6_Click()

        '<EhHeader>
        On Error GoTo Command6_Click_Err

        '</EhHeader>

100     MSHFlexGrid1.Height = 1700

110     Gdb.Execute "drop table DOK13"

120     Gdb.Execute "select  ATIM+CONVERT(CHAR(10),HME,3)  AS DDD INTO DOK13" & " FROM EGGTIM WHERE HME>='" + Format(apo, "MM/DD/YYYY") + "'" + "AND HME<'" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "' AND KODE='" + Text1.Text + "'"

        'Gdb.Execute "drop table DOK14"
        'Gdb.Execute "select  ATIM,HME  INTO DOK14" _
        '& " FROM EGGTIM WHERE HME>='" + Format(APO, "MM/DD/YYYY") + "'" + _
        '"AND HME<'" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY") + "' AND KODE='" + Text1.Text + "'"

130     Adodc1.ConnectionString = gConnect

140     Adodc1.RecordSource = "SELECT KODE,EID.ONO,SUM(POSO) AS SYNOLOPOS,CONVERT(DECIMAL(10,2),SUM(POSO*TIMM)) AS AJIA2  FROM EGGTIM  LEFT JOIN EID ON EGGTIM.KODE=EID.KOD WHERE  ATIM+CONVERT(CHAR(10),HME,3) IN (SELECT DDD FROM  DOK13) GROUP BY KODE,EID.ONO"
150     Adodc1.Refresh
160     MSHFlexGrid1.ColWidth(0) = 1500
170     MSHFlexGrid1.ColWidth(1) = 3500

180     MSHFlexGrid1.ColAlignment = flexAlignRightCenter

        Dim SX

190     SX = 0

200     For k = 1 To MSHFlexGrid1.ROWS - 1
210         SX = SX + MSHFlexGrid1.TextMatrix(k, 3)
        Next

220     MSHFlexGrid1.AddItem ("" & Chr(9) & "Σύνολα" & Chr(9) & Chr(9) & Format(SX, "###,###.00"))

        '<EhFooter>
        Exit Sub

Command6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Command6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Command6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command7_Click()

        '<EhHeader>
        On Error GoTo Command7_Click_Err

        '</EhHeader>

        Dim ANS

100     If Val(apot.Text) = 0 Then
110         MsgBox " Δεν δώσατε Αποθήκη"

            Exit Sub

        End If

120     MsgBox "Θα μας ρωτήσει ποιές οικογένειες και υποοικογένειες (αριθμητικώς) θέλουμε να σβήσουμε και θα μας ζητηθεί επιβεβαίωση στο τέλος"

        Dim ypooik, oik

130     ypooik = "0"
140     oik = "0"

        'ypooik = InputBox("Να μηδενιστεί η υπο οικογένεια ", "με 0 όλες οι υποοικογένειες")

        'oik = InputBox("Να μηδενιστεί η οικογένεια ", "με 0 όλες οι οικογένειες")

        Dim synt

150     If Val(oik) = 0 And Val(ypooik) = 0 Then
160         synt = "  "
        Else
170         synt = " EID.KODLOG=" + ypooik + " AND EID.AEG=" + oik + " AND "
        End If

        ' ANS = MsgBox("Θα μηδενιστεί η αποθήκη " + apot.Text + " με  ημερομηνία Απογραφής " + Format(apog.value, "dd/mm/yyyy"), vbYesNo, "ΕΙΣΘΕ ΣΙΓΟΥΡΟΙ;")
        'If Not ANS = vbYes Then
        '   Exit Sub
        'End If

        On Error Resume Next

180     Gdb.Execute "DROP TABLE DOKEGGT31"

        On Error GoTo Command7_Click_Err

190     Gdb.Execute "SELECT * INTO  DOKEGGT31 FROM EGGTIM WHERE HME<'" + Format(DateAdd("D", 1, apog.Value), "MM/DD/YYYY") + "'"

200     Gdb.Execute "UPDATE DOKEGGT31 SET XRE=-XRE WHERE LEFT(ATIM,1)='λ'"

210     a = "SELECT KODE,APOT,SUM( CASE APOT  WHEN 1  THEN XRE ELSE 0  END ) AS S1X," _
           & " SUM(CASE APOT  WHEN 1  THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS1X," _
           & " SUM(CASE APOT  WHEN 1  THEN PIS ELSE 0  END ) AS S1P," _
           & " SUM(CASE APOT  WHEN 1  THEN PIS*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS1P," _
           & " SUM( CASE APOT  WHEN 2 THEN XRE ELSE 0  END ) AS S2X," _
           & " SUM( CASE APOT  WHEN 2  THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS2X," _
           & " SUM(CASE APOT  WHEN 2  THEN PIS*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS2P," _
           & " SUM( CASE APOT WHEN 2 THEN PIS ELSE 0 END ) AS S2P," _
           & " SUM( CASE APOT  WHEN 3 THEN XRE ELSE 0  END ) AS S3X," _
           & " SUM( CASE APOT  WHEN 3  THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS3X," _
           & " SUM(CASE APOT  WHEN 3  THEN PIS*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS3P," _
           & " SUM( CASE APOT WHEN 3 THEN PIS ELSE 0 END ) AS S3P," _
           & " SUM( CASE APOT  WHEN 4 THEN XRE ELSE 0  END ) AS S4X," _
           & " SUM( CASE APOT  WHEN 4  THEN XRE*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS4X," _
           & " SUM(CASE APOT  WHEN 4  THEN PIS*TIMM*(1-EKPT/100) ELSE 0  END ) AS AS4P," _
           & " SUM( CASE APOT WHEN 4 THEN PIS ELSE 0 END ) AS S4P" _
           & " INTO DOKEGGT1 "
220     a = a + " FROM DOKEGGT31 " + "  GROUP BY KODE,APOT"

        ' WHERE (NOT APOT IS NULL) AND (NOT EKPT IS NULL) AND (NOT XRE IS NULL) AND (NOT PIS IS NULL) AND (NOT TIMM IS NULL)
        '   On Error GoTo Command7_Click_Err
        On Error Resume Next

230     Gdb.Execute "DROP TABLE DOKEGGT1"
240     Gdb.Execute a, k

        Dim R       As New ADODB.Recordset

        Dim REGGTIM As New ADODB.Recordset

        Dim yp      As Single

250     R.Open "SELECT * FROM DOKEGGT1 ", Gdb, adOpenForwardOnly, adLockReadOnly
260     REGGTIM.Open "SELECT * FROM EGGTIM ", Gdb, adOpenDynamic, adLockOptimistic

        Dim YPOL(4)

270     R.MoveFirst

        Dim F, FF

280     Do While Not R.EOF

            'YPOLOGIZV TO YPOLOIPO APO TO ARXEIO DOKEGGT1 (YP)
290         For k = 1 To 4
300             F = "S" + Format(k, "0") + "X"
310             FF = "S" + Format(k, "0") + "P"

320             If IsNull(R(F)) Then
330                 R(F) = 0
                End If

340             yp = -(R(F) - IIf(IsNull(R(FF)), 0, R(FF)))

                'AN DEN EINAI 0 TOTE PROSTHETO ή αφαιρώ για να γίνει 0 το υπολοιπό του
350             If yp <> 0 Then
360                 REGGTIM.AddNew
370                 REGGTIM("hme") = CDate("01/01/" + Format(Year(apog.Value), "0000"))

380                 REGGTIM("ATIM") = "λ00003"
390                 REGGTIM("KODE") = R("KODE")
400                 REGGTIM("poso") = yp
410                 REGGTIM("xre") = yp
420                 REGGTIM("APOT") = k
430                 REGGTIM.Update
                End If

            Next

440         Me.Caption = R("KODE")
450         R.MoveNext

460         DoEvents
        Loop

470     R.Close

        '<EhFooter>
        Exit Sub

Command7_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Command7_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Command7_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command8_Click()

        'ΛΗΨΗ ΑΠΟ EXCEL
        ' ============================
        '<EhHeader>
        On Error GoTo Command8_Click_Err

        '</EhHeader>
Dim Anoigma_AnDenYparxei As Boolean
Dim answer As Integer

 answer = MsgBox("Να ανοιχθεί το είδος αν δεν υπάρχει;" + Chr(13) + "Θα πάρει την περιγραφή από την 4η στήλη με κατηγορία φπα το 2", vbYesNoCancel)
 If answer = vbYes Then
    Anoigma_AnDenYparxei = True
 Else
     Anoigma_AnDenYparxei = False
 
 End If
        
        
        
        
        
        
        Dim xl             As New excel.Application

        Dim xlsheet        As excel.Worksheet

        Dim xlsheet3       As excel.Worksheet

        Dim xlwbook        As excel.workbook

        Dim DBF            As Database

        Dim sql            As New ADODB.Connection

        Dim rDBF           As Recordset

        Dim rSQL           As New ADODB.Recordset

        Dim rSQL2          As New ADODB.Recordset

        Dim rYPOL          As New ADODB.Recordset

        Dim mYPOL          As Single

        Dim conDBF         As String

        Dim conSQL         As String

        Dim k              As Long

        Dim Fname          As String

        Dim dB             As DAO.Database

        Dim arxeio         As String, arxeio2 As String

        Dim SynStoYpoloipo As Integer

        Dim m_Proeleysh
           m_Proeleysh = "           "
        Dim M_ATIM

100     M_ATIM = "λ" + Format(time(), "hhmmss")

110     SynStoYpoloipo = MsgBox("Να προστεθεί στο ήδη υπάρχον υπόλοιπο;" + Chr(13) + "Σε περίπτωση που απαντήσετε όχι η ποσότητα που απογράψατε θα είναι και το υπόλοιπο", vbYesNoCancel)

120     If SynStoYpoloipo = vbCancel Then
130         MsgBox "Η εργασία ακυρώθηκε"

            Exit Sub

        End If

        Dim BARC1STILI As Integer

140     BARC1STILI = MsgBox("Η πρώτη στήλη είναι το BARCODE; " + Chr(13) + " Oxι σημαίνει ότι είναι ο κωδικός", vbYesNo)

        '====================================================

        'On Error GoTo open_error
150     Me.MousePointer = vbHourglass

160     If Len(Dir(FILENAME.Text, vbNormal)) > 0 Then
            ' OK
        Else
170         MsgBox " ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ ΑΡΧΕΙΟ " + FILENAME.Text

            Exit Sub

        End If

180     If Val(apot) = 0 Then
190         MsgBox " ΔΕΝ ΥΠΑΡΧΕΙ ΑΠΟΘΗΚΗ "

            Exit Sub

        End If

200     If SynStoYpoloipo = vbNo Then  ' αν προσαρμοσει τα υπολοιπα
            '   UPDATE_YPOLOIPA2 ""
        End If

210     Set xlwbook = xl.Workbooks.Open(FILENAME.Text)
220     Set xlsheet = xlwbook.Sheets.Item(1)

        'Set db = OpenDatabase(Dir1.Path, False, False, "dBase III;")

        'Set rDBF = db.OpenRecordset("SELECT * FROM " + File1.FILENAME)

230     rSQL.Open "SELECT top 1 * FROM EGGTIM ", Gdb, adOpenDynamic, adLockOptimistic

        'rDBF.MoveFirst

        Dim Z

240     Z = 0
250     k = 0

        On Error GoTo WRITEERROR    ' Resume Next

        Dim OK As Boolean

        Dim M_CODE

        Dim RR      As New ADODB.Recordset

        Dim RR2     As New ADODB.Recordset

        Dim TREX_YP As Single

        Dim nn      As Integer

        ' If IsNull(xlsheet.cells(mRow, KOD_COLUMN)) Then

        Dim FOUND   As Boolean

        Dim CC      As String

260     mRow = 1

270     Do While True  '   Not IsNull(xlsheet.cells(mRow, 1))

280         If IsNull(xlsheet.cells(mRow, 1)) Then

                Exit Do

            End If

290         If IsEmpty(xlsheet.cells(mRow, 1)) Then

                Exit Do

            End If

300         OK = True

310         M_CODE = xlsheet.cells(mRow, 1)
            '  On Error Resume Next
            
            On Error GoTo diortsestring
                M_CODE = "" + M_CODE
            On Error Resume Next
            

320         FOUND = True

330         If BARC1STILI = vbYes Then    'barcode 1η στήλη
340             RR.Open "select * FROM BARCODES WHERE ERG='" + M_CODE + "'", Gdb, adOpenDynamic, adLockOptimistic

350             If RR.EOF Then

360                 FOUND = False
370                 MsgBox "ΔΕΝ ΒΡΕΘΗΚΕ ΤΟ BARCODE " + M_CODE
380                 List2.AddItem "ΔΕΝ ΒΡΕΘΗΚΕ : " + M_CODE
390                 xlsheet.cells(mRow, 4) = "ΔΕΝ ΒΡΕΘΗΚΕ"

                End If

400             M_CODE = RR("KOD")
410             m_Proeleysh = RR("Proeleysh")
420             RR.Close

            End If

430         If SynStoYpoloipo = vbNo Then

440             sql = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS YPOL FROM EGGTIM WHERE KODE='" + M_CODE + "'  and APOT=" + Left(apot, 1)
                'RR2.Open sql, Gdb, adOpenDynamic, adLockOptimistic

450             RR2.Open "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS YPOL FROM EGGTIM WHERE KODE='" + M_CODE + "'  and APOT=" + Left(apot, 1), Gdb, adOpenDynamic, adLockOptimistic

460             TREX_YP = RR2("YPOL")
470             RR2.Close
            Else
                ' rSQL("poso") = xlsheet.cells(mRow, 2)
480             TREX_YP = 0
            End If

490         If FOUND Then

                '          rSQL.AddNew
                '          rSQL("hme") = apog.value
                '          rSQL("ATIM") = M_ATIM '"λ00002"
                '          rSQL("KODE") = M_CODE
                '          rSQL("timm") = xlsheet.cells(mRow, 3)
                '          rSQL("Proeleysh") = Left(m_Proeleysh, 10)
                '          rSQL("poso") = -TREX_YP + xlsheet.cells(mRow, 2)
                '          rSQL("XRE") = -TREX_YP + xlsheet.cells(mRow, 2)
                '          rSQL("APOT") = Val(APOT.Text)
                '         rSQL.Update
                '

500             CC = "INSERT INTO EGGTIM (HME,ATIM,KODE,TIMM,PROELEYSH,POSO,XRE,APOT) VALUES"
510             CC = CC + "('" + Format(apog.Value, "MM/DD/YYYY") + "',"
520             CC = CC + "'" + M_ATIM + "',"
530             CC = CC + "'" + M_CODE + "',"
540             CC = CC + "" + Replace(str(xlsheet.cells(mRow, 3)), ",", ".") + ","
550             CC = CC + "'" + Left(m_Proeleysh, 10) + "',"
560             CC = CC + "" + str(-TREX_YP + xlsheet.cells(mRow, 2)) + ","
570             CC = CC + "" + Replace(str(-TREX_YP + xlsheet.cells(mRow, 2)), ",", ".") + ","
580             CC = CC + "" + str(Val(apot.Text)) + ")"
590             Gdb.Execute CC, nn
                If nn = 0 Then
                    xlsheet.cells(mRow, 6) = " ΔΕΝ ΚΑΤΑΩΡΗΘΗΚΕ  " + CC
                End If
600             Gdb.Execute "UPDATE EID SET HM1=GETDATE() WHERE KOD='" + M_CODE + "'", nn

610             If nn = 0 Then
                    ' MsgBox "προσοχη δεν βρέθηκε ο κωδικός " + M_CODE
                    If Anoigma_AnDenYparxei Then
                         ' Gdb.Execute "insert into EID (KOD,ONO,FPA,MON) VALUES ('" + M_CODE + "','" + Replace(xlsheet.cells(mRow, 4), "'", "`") + "',2,'TEM')"
                         ' Gdb.Execute "UPDATE EID SET HM1=GETDATE() WHERE KOD='" + M_CODE + "'", nn
                          
                          Gdb.Execute "insert into EID (KOD,ONO,FPA,MON) VALUES ('" + M_CODE + "','" + Replace(xlsheet.cells(mRow, 4), "'", "`") + "',2,'TEM')"
                          Gdb.Execute "UPDATE EID SET HM1=GETDATE() WHERE KOD='" + M_CODE + "'", nn
                          
                    Else
620                    List2.AddItem "ΔΕΝ ΒΡΕΘΗΚΕ : " + M_CODE
630                    xlsheet.cells(mRow, 5) = "ΔΕΝ ΒΡΕΘΗΚΕ"
                    End If
                    
                End If
            End If

640         mRow = mRow + 1

650         If mRow Mod 10 = 0 Then
660             Me.Caption = mRow
            End If

670         DoEvents
        Loop

680     rSQL.Close

690     rSQL.Open "SELECT * FROM TIM ", Gdb, adOpenDynamic, adLockOptimistic

700     rSQL.AddNew
710     rSQL("hme") = apog.Value
720     rSQL("ATIM") = M_ATIM    ' "λ00001"
        'rSQL("KODE") = M_CODE
        'rSQL("timm") = rDBF("xti")
        ' rSQL("poso") = rDBF("POS") - mYPOL
        ' rSQL("xre") = rDBF("POS") - mYPOL
730     rSQL("b_n1") = Val(apot.Text)
740     rSQL.Update

        On Error GoTo Command8_Click_Err

750     If List2.ListCount > 0 Then
760         MsgBox "ΤΑ ΛΑΘΗ ΑΠΟΘΗΚΕΥΤΗΚΑΝ ΣΤΟ EXCEL  "
770         xlwbook.save  '  "c:\ektyp2.xls"
        End If

780     Me.MousePointer = vbNormal

790     Set xlsheet = Nothing
800     Set xlwbook = Nothing
        ' excel.Quit

810     xl.Quit

820     MsgBox "Ενημερώθηκαν " + Format(mRow, "###0") + " εγγραφές"

        Exit Sub

WRITEERROR:

830     Resume Next

        '<EhFooter>
        Exit Sub
        
        
diortsestring:
        M_CODE = str(M_CODE)
        M_CODE = LTrim$(M_CODE)
        M_CODE = Trim$(M_CODE)
        Resume Next
        

Command8_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Command8_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Command8_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command9_Click()

        'ΑΠΟ CSV  / SYMBOL
        ' ============================
        '<EhHeader>
        On Error GoTo Command9_Click_Err

        '</EhHeader>

        Dim xl             As New excel.Application

        Dim xlsheet        As excel.Worksheet

        Dim xlsheet3       As excel.Worksheet

        Dim xlwbook        As excel.workbook

        Dim DBF            As Database

        Dim sql            As New ADODB.Connection

        Dim rDBF           As Recordset

        Dim rSQL           As New ADODB.Recordset

        Dim rSQL2          As New ADODB.Recordset

        Dim rYPOL          As New ADODB.Recordset

        Dim mYPOL          As Single

        Dim conDBF         As String

        Dim conSQL         As String

        Dim k              As Long

        Dim Fname          As String

        Dim dB             As DAO.Database

        Dim arxeio         As String, arxeio2 As String

        Dim SynStoYpoloipo As Integer

        Dim m_Proeleysh

        Dim M_ATIM

100     M_ATIM = "λ" + Format(time(), "hhmmss")

110     SynStoYpoloipo = MsgBox("Να προστεθεί στο ήδη υπάρχον υπόλοιπο;" + Chr(13) + "Σε περίπτωση που απαντήσετε όχι η ποσότητα που απογράψατε θα είναι και το υπόλοιπο", vbYesNoCancel)

120     If SynStoYpoloipo = vbCancel Then
130         MsgBox "Η εργασία ακυρώθηκε"

            Exit Sub

        End If

        Dim BARC1STILI As Integer

140     BARC1STILI = MsgBox("Η πρώτη στήλη είναι το BARCODE; " + Chr(13) + " Oxι σημαίνει ότι είναι ο κωδικός", vbYesNo)

        '====================================================

        'On Error GoTo open_error
150     Me.MousePointer = vbHourglass

160     If Len(Dir(FILENAME.Text, vbNormal)) > 0 Then
            ' OK
        Else
170         MsgBox " ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ ΑΡΧΕΙΟ " + FILENAME.Text

            Exit Sub

        End If

180     If Val(apot) = 0 Then
190         MsgBox " ΔΕΝ ΥΠΑΡΧΕΙ ΑΠΟΘΗΚΗ "

            Exit Sub

        End If

200     If SynStoYpoloipo = vbNo Then  ' αν προσαρμοσει τα υπολοιπα
            '   UPDATE_YPOLOIPA2 ""
        End If

210     Open FILENAME.Text For Input As #1
        'Set xlwbook = xl.Workbooks.Open(FILENAME.Text)
        'Set xlsheet = xlwbook.Sheets.Item(1)

        'Set db = OpenDatabase(Dir1.Path, False, False, "dBase III;")

        'Set rDBF = db.OpenRecordset("SELECT * FROM " + File1.FILENAME)

220     rSQL.Open "SELECT top 1 * FROM EGGTIM ", Gdb, adOpenDynamic, adLockOptimistic

        'rDBF.MoveFirst

        Dim Z

230     Z = 0
240     k = 0

        On Error GoTo WRITEERROR    ' Resume Next

        Dim OK As Boolean

        Dim M_CODE

        Dim RR      As New ADODB.Recordset

        Dim RR2     As New ADODB.Recordset

        Dim TREX_YP As Single

        Dim nn      As Integer

        ' If IsNull(xlsheet.cells(mRow, KOD_COLUMN)) Then

        Dim FOUND   As Boolean

        'KENES SEIRES
250     mRow = 1   '  data1.Recordset.Move

260     For k = 1 To mRow
270         Input #1, AA
        Next

        Dim ELEM(1 To 60)

280     mRow = 2

290     Do While True  '   Not IsNull(xlsheet.cells(mRow, 1))

300         Line Input #1, AA

            ' DUM = to437(AA)
310         If EOF(1) Then

                Exit Do

            End If

320         AA = to928(AA)
330         DUM = bohu5.FETES_DELIM(AA, ELEM)

340         OK = True

350         M_CODE = ELEM(1)
            '  On Error Resume Next

360         FOUND = True

370         If BARC1STILI = vbYes Then    'barcode 1η στήλη
380             RR.Open "select * FROM BARCODES WHERE ERG='" + M_CODE + "'", Gdb, adOpenDynamic, adLockOptimistic

390             If RR.EOF Then

400                 FOUND = False
410                 MsgBox "ΔΕΝ ΒΡΕΘΗΚΕ ΤΟ BARCODE " + M_CODE
420                 List2.AddItem "ΔΕΝ ΒΡΕΘΗΚΕ : " + M_CODE
                    ' xlsheet.cells(mRow, 4) = "ΔΕΝ ΒΡΕΘΗΚΕ"

                End If

430             M_CODE = RR("KOD")
440             m_Proeleysh = RR("Proeleysh")
450             RR.Close

            End If

460         If SynStoYpoloipo = vbNo Then

470             sql = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS YPOL FROM EGGTIM WHERE KODE='" + M_CODE + "'  and APOT=" + Left(apot, 1)
                'RR2.Open sql, Gdb, adOpenDynamic, adLockOptimistic

480             RR2.Open "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS YPOL FROM EGGTIM WHERE KODE='" + M_CODE + "'  and APOT=" + Left(apot, 1), Gdb, adOpenDynamic, adLockOptimistic

490             TREX_YP = RR2("YPOL")
500             RR2.Close
            Else
                ' rSQL("poso") = xlsheet.cells(mRow, 2)
510             TREX_YP = 0
            End If

520         If FOUND Then

530             rSQL.AddNew
540             rSQL("hme") = apog.Value
550             rSQL("ATIM") = M_ATIM    '"λ00002"
560             rSQL("KODE") = M_CODE
570             rSQL("timm") = ELEM(3)
580             rSQL("Proeleysh") = Left(m_Proeleysh, 10)
590             rSQL("poso") = -TREX_YP + ELEM(2)
600             rSQL("XRE") = -TREX_YP + ELEM(2)
610             rSQL("APOT") = Val(apot.Text)
620             rSQL.Update

630             Gdb.Execute "UPDATE EID SET HM1=GETDATE() WHERE KOD='" + M_CODE + "'", nn

640             If nn = 0 Then
                    ' MsgBox "προσοχη δεν βρέθηκε ο κωδικός " + M_CODE

650                 List2.AddItem "ΔΕΝ ΒΡΕΘΗΚΕ : " + M_CODE
660                 xlsheet.cells(mRow, 4) = "ΔΕΝ ΒΡΕΘΗΚΕ"
                End If
            End If

670         mRow = mRow + 1

680         If mRow Mod 10 = 0 Then
690             Me.Caption = mRow
            End If

700         DoEvents
        Loop

710     rSQL.Close

720     rSQL.Open "SELECT * FROM TIM ", Gdb, adOpenDynamic, adLockOptimistic

730     rSQL.AddNew
740     rSQL("hme") = apog.Value
750     rSQL("ATIM") = M_ATIM    ' "λ00001"
        'rSQL("KODE") = M_CODE
        'rSQL("timm") = rDBF("xti")
        ' rSQL("poso") = rDBF("POS") - mYPOL
        ' rSQL("xre") = rDBF("POS") - mYPOL
760     rSQL("b_n1") = Val(apot.Text)
770     rSQL.Update

        On Error GoTo Command9_Click_Err

780     If List2.ListCount > 0 Then
790         MsgBox "ΤΑ ΛΑΘΗ ΑΠΟΘΗΚΕΥΤΗΚΑΝ ΣΤΟ EXCEL  "
800         xlwbook.save  '  "c:\ektyp2.xls"
        End If

810     Me.MousePointer = vbNormal

        'Set xlsheet = Nothing
        'Set xlwbook = Nothing
        ' excel.Quit

        'xl.Quit

820     MsgBox "Ενημερώθηκαν " + Format(mRow, "###0") + " εγγραφές"

        Exit Sub

WRITEERROR:

830     Resume Next

        '<EhFooter>
        Exit Sub

Command9_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Command9_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Command9_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_DblClick()

        '<EhHeader>
        On Error GoTo Form_DblClick_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "BOHU7"
110     PARAMETROI.SHOW 1

        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Form_DblClick " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Form_DblClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 27 Then Command4_Click

        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Form_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Form_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
        On Error Resume Next

100     mForm_Load Me, fh, fw, ft, fl

110     Me.Picture = LoadPicture(gPicture)
120     F_EXO_XROMATA = Val(FINDPARAMETROI(1, "PAR1", "F_EXO_XROMATA", "0", "Γ.1 ΧΡΗΣΙΜΟΠΟΙΩ ΜΕΓΕΘΗ-ΧΡΩΜΑΤΑ=1"))   'posa psifia tha exei h kathe seira


        F_GCONN2 = FINDPARAMETROI(1, "BOHU7", "F_GCONN2", "", "CONNECTION STRING1")   'posa psifia tha exei h kathe seira
        
        F_GCONN3 = FINDPARAMETROI(1, "BOHU7", "F_GCONN3", "", "CONNECTION STRING2")   'posa psifia tha exei h kathe seira
        
        F_GCONN4 = FINDPARAMETROI(1, "BOHU7", "F_GCONN4", "", "CONNECTION STRING3")   'posa psifia tha exei h kathe seira
        
        
        
        
        
        
130     If Left(Command, 3) = "DEB" Then MsgBox "1"

140     Me.KeyPreview = True
150     apo = Now
160     eos = Now
170     apog = Now
180     Dir1.Path = "C:\LAGEURO"

        Dim R As New ADODB.Recordset

190     F_TAKENEESTIMES = Val(FINDPARAMETROI(1, "BOHU7", "F_TAKENEESTIMES", 0, "ΕΝΗΜΕΡΩΣΗ ΤΙΜΩΝ ΑΠΟ ΙΝΤΕΡΝΕΤ 0=OXI 1=NAI 2=ΟΛΑ ΕΛΕΥΘΕΡΑ"))

        'If F_TAKENEESTIMES = 0 Then
        '  neesTimes.Enabled = False
        '  NEABARCODES.Enabled = False
        'End If

200     If F_TAKENEESTIMES = 0 Then
210         neesTimes.Enabled = False
220         NEABARCODES.Enabled = False
230         SendKiniseis.Enabled = False
240     ElseIf F_TAKENEESTIMES = 2 Then
            '
        Else
250         Command3.Enabled = False
        End If

        '   'SkinFramework1.LoadSkin "c:\mercvb\Styles\Office2007.cjstyles", ""
        '   ' 'SkinFramework1.LoadSkin "c:\mercvb\Styles\WinXP.Luna.cjstyles", "NormalBlue.ini"
        '    'SkinFramework1.ApplyWindow Me.hWnd
        '    'SkinFramework1.ApplyOptions = 'SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

260     If Left(Command, 3) = "DEB" Then MsgBox "2"

        'αποθηκη
270     R.Open "SELECT *FROM PINAKES WHERE TYPOS=4 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

280     Do While Not R.EOF

290         If R("typos") = 4 Then
300             apot.AddItem Format(R("ayjon"), "##") + "." + R("PERIGRAFH")

310             If R("ayjon") > 1 And R("ayjon") < 99 Then
320                 APOT2.AddItem Format(R("ayjon"), "##") + "." + R("PERIGRAFH")
                End If

            Else

                Exit Do

            End If

330         R.MoveNext
        Loop

340     apot.Text = apot.List(0)

350     If Left(Command, 3) = "DEB" Then MsgBox "3"

360     If Len(Dir("C:\kentriko.TXT")) > 0 Then
370         APOT2.Text = APOT2.List(0)
        Else
380         APOT2.Text = APOT2.List(1)
        End If

390     If Len(Dir("C:\YPOK2.TXT")) > 0 Then
400         APOTH2.Text = APOTH2.List(2)
        End If

410     If Len(Dir("C:\YPOK3.TXT")) > 0 Then
420         APOTH2.Text = APOTH2.List(3)
        End If

430     If Left(Command, 3) = "DEB" Then MsgBox "4"

440     R.Close
450     F_APO = FINDPARAMETROI(1, "BOHU7", "F_APO", Format(Now, "DD/MM/YYYY"), "ΗΜΕΡ.ΕΝΗΜΕΡΩΣΗΣ ΑΠΟ ")
460     F_EOS = FINDPARAMETROI(1, "BOHU7", "F_EOS", Format(Now, "DD/MM/YYYY"), "ΗΜΕΡ.ΕΝΗΜΕΡΩΣΗΣ ΕΩΣ ")

470     apo = CDate(F_APO)
480     eos = CDate(F_EOS)

490     If Left(Command, 3) = "DEB" Then MsgBox "9"

500     If gXEIRISTHS < 8 Then
            'Dim K
510         kodisbarcode.Visible = False
520         FILENAME.Visible = False
530         File1.Visible = False
540         Dir1.Visible = False
550         apog.Visible = False
560         Label9.Visible = False

570         Label1.Visible = False
580         Label2.Visible = False
590         Label3.Visible = False
600         Label4.Visible = False
610         Label5.Visible = False
620         Label10.Visible = False
630         Label8.Visible = False
640         Command5.Visible = False
            'Command8.Visible = False
650         apot.Visible = False
660         Text3.Visible = False
670         Text1.Visible = False
680         Text2.Visible = False

690         List1.Visible = False

700         mhdenismos.Visible = False
710         Command5.Visible = False
720         Command2.Visible = False
730         Command1.Visible = False
        End If

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub UPDATE_YPOLOIPA2(ByVal synt)    ' ΣΥΝΘΗΚΗ  ""=ΟΛΑ   "KODLOG=1 AND AEG=11"  ΜΟΝΟ ΤΗΝ ΚΑΤΗΓΟΡΙΑ 1 ΚΑΙ 11 ΥΠΟΚΑΤΗΓΟΡΙΑ

        'υπολογιζει τα υπόλοιπα της αποθήκης με ημερομηνία  απογραφής
        '<EhHeader>
        On Error GoTo UPDATE_YPOLOIPA2_Err

        '</EhHeader>
        Dim a As String

        Dim B As String

        Dim d As Long

        'Dim krit As String

        ' ΒΡΙΣΚΩ ΜΙΑ ΦΟΡΑ ΤΑ ΥΠΟΛΟΙΠΑ ΑΝΑ ΑΠΟΘΗΚΗ

        On Error Resume Next

100     Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1

        On Error GoTo UPDATE_YPOLOIPA2_Err

110     a = "SELECT KODE,SUM( CASE APOT  WHEN 1  THEN XRE ELSE 0  END ) AS S1X," & " SUM(CASE APOT  WHEN 1  THEN PIS ELSE 0  END ) AS S1P," & "SUM( CASE APOT  WHEN 2 THEN XRE ELSE 0  END ) AS S2X," & "SUM( CASE APOT  WHEN 4 THEN XRE ELSE 0  END ) AS S4X," & " SUM(CASE APOT  WHEN 4  THEN PIS ELSE 0  END ) AS S4P," & "SUM( CASE APOT  WHEN 3 THEN XRE ELSE 0  END ) AS S3X," & " SUM(CASE APOT  WHEN 3  THEN PIS ELSE 0  END ) AS S3P," & "SUM( CASE APOT WHEN 2 THEN PIS ELSE 0 END ) AS S2P INTO DOKEGGT1 "
120     a = a + " FROM EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD WHERE " + synt + " APOT=" + Left(apot.Text, 1) + " AND  HME<'" + Format(DateAdd("d", 1, apog.Value), "MM/DD/YYYY") + "' GROUP BY KODE"
130     Gdb.Execute a, d

140     If d = 0 Then
150         MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"

            Exit Sub

        End If

        '<EhFooter>
        Exit Sub

UPDATE_YPOLOIPA2_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.UPDATE_YPOLOIPA2 " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.UPDATE_YPOLOIPA2 " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 12390, 8775, fh, fw, ft, fl
End Sub

Private Sub kodisbarcode_Click()

End Sub

Private Sub mhdenismos_Click()

        '<EhHeader>
        On Error GoTo mhdenismos_Click_Err

        '</EhHeader>

        Dim ANS

100     If Val(apot.Text) = 0 Then
110         MsgBox " Δεν δώσατε Αποθήκη"

            Exit Sub

        End If

120     MsgBox "Θα μας ρωτήσει ποιές οικογένειες και υποοικογένειες (αριθμητικώς) θέλουμε να σβήσουμε και θα μας ζητηθεί επιβεβαίωση στο τέλος"

        Dim ypooik, oik

130     ypooik = InputBox("Να μηδενιστεί η υπο οικογένεια ", "με 0 όλες οι υποοικογένειες ")

140     oik = InputBox("Να μηδενιστεί η οικογένεια ", "με 0 όλες οι οικογένειες")

        Dim synt

150     If Val(oik) = 0 And Val(ypooik) = 0 Then
160         synt = "  "
        Else

170         If Val(ypooik) = 0 Then
180             synt = " EID.AEG=" + oik + " AND "
            Else
190             synt = " EID.KODLOG=" + ypooik + " AND EID.AEG=" + oik + " AND "
            End If

        End If

200     ANS = MsgBox("Θα μηδενιστεί η αποθήκη " + apot.Text + " με  ημερομηνία Απογραφής " + Format(apog.Value, "dd/mm/yyyy"), vbYesNo, "ΕΙΣΘΕ ΣΙΓΟΥΡΟΙ;")

210     If Not ANS = vbYes Then

            Exit Sub

        End If

220     UPDATE_YPOLOIPA2 synt    'YPOLOGISE YPOLOIPA KAI ΦΕΡΤΑ ΣΤΟ ΑΡΧΕΙΟ DOKEGGT1

        Dim R       As New ADODB.Recordset

        Dim REGGTIM As New ADODB.Recordset

        Dim yp      As Single

230     R.Open "SELECT * FROM DOKEGGT1 ", Gdb, adOpenForwardOnly, adLockReadOnly
240     REGGTIM.Open "SELECT * FROM EGGTIM ", Gdb, adOpenDynamic, adLockOptimistic

250     R.MoveFirst

260     Do While Not R.EOF

            On Error GoTo mhdenismos_Click_Err

            'YPOLOGIZV TO YPOLOIPO APO TO ARXEIO DOKEGGT1 (YP)
270         If Val(apot.Text) = 1 Then
280             yp = R("S1X") - IIf(IsNull(R("S1P")), 0, R("S1P"))
290         ElseIf Val(apot.Text) = 3 Then
300             yp = R("S3X") - IIf(IsNull(R("S3P")), 0, R("S3P"))

310         ElseIf Val(apot.Text) = 4 Then
320             yp = R("S4X") - IIf(IsNull(R("S4P")), 0, R("S4P"))

330         ElseIf Val(apot.Text) = 2 Then
340             yp = R("S2X") - IIf(IsNull(R("S2P")), 0, R("S2P"))    'R("S2P")
            End If

            'AN DEN EINAI 0 TOTE PROSTHETO ή αφαιρώ για να γίνει 0 το υπολοιπό του
350         If yp <> 0 Then
360             REGGTIM.AddNew
370             REGGTIM("hme") = apog.Value
380             REGGTIM("ATIM") = "λ00003"
390             REGGTIM("KODE") = R("KODE")
400             REGGTIM("poso") = -yp
410             REGGTIM("xre") = -yp
420             REGGTIM("APOT") = Val(apot.Text)
430             REGGTIM.Update
            End If

440         Me.Caption = R("KODE")
450         R.MoveNext

460         DoEvents
        Loop

470     R.Close

        '<EhFooter>
        Exit Sub

mhdenismos_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.mhdenismos_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.mhdenismos_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub NEABARCODES_Click()
Dim a As Integer
Dim N As Integer

If AutoDelete Then N = 1 Else N = 0

    a = UPD_NEA_BARCODE(N)

End Sub

Public Function UPD_NEA_BARCODE(DELETEOLD As Integer) As Integer


        '<EhHeader>
        On Error GoTo NEABARCODES_Click_Err

        '</EhHeader>
        Dim sql

        On Error Resume Next

100     Gdb.Execute "DROP TABLE BARCODES0"

        On Error GoTo BGES

        Dim dB As DAO.Database

110     Set dB = OpenDatabase("", False, False, gConnect)

120     dB.Execute "SELECT * INTO BARCODES0 FROM BARCODES IN 'C:\LAGEURO' 'dBASE III;' "

        On Error Resume Next


130     If DELETEOLD = 1 Then
140         'List2.AddItem "DELETE FROM BARCODES"
150         Gdb.Execute "DELETE FROM BARCODES"
        End If

        Dim k

        On Error GoTo NEABARCODES_Click_Err  ' Resume Next

160     k = 0
         Gdb.Execute "DELETE FROM BARCODES0 WHERE ERG IS NULL", k
170     If DELETEOLD = 1 Then
180         sql = "INSERT INTO BARCODES(KOD,ERG) SELECT KOD,ERG FROM BARCODES0"
        Else
190         sql = "INSERT INTO BARCODES(KOD,ERG) SELECT KOD,ERG FROM BARCODES0 WHERE ERG NOT IN ( SELECT ERG FROM BARCODES)"
        End If
    
200    ' List2.AddItem sql
210     Gdb.Execute sql, k

220     DoEvents
    
230     sql = "DELETE FROM BARCODES WHERE ERG NOT IN ( SELECT ERG FROM BARCODES0 )"
240     ' List2.AddItem sql
250     Gdb.Execute sql

260     DoEvents

270     If k > 0 Then
280         Me.Caption = "ΕΝΗΜΕΡΩΘΗΚΕ ME NEA BARCODE"
            UPD_NEA_BARCODE = 1
        End If

        Exit Function

BGES:
290     Me.Caption = "ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ " + Err.Description
        UPD_NEA_BARCODE = 0
        '<EhFooter>
        Exit Function

NEABARCODES_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.NEABARCODES_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.NEABARCODES_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub neesTimes_Click()
     get_times_eid
        
End Sub
        
Sub get_times_eid()
        ' METAFORA apo ARXEIO DBF
        '<EhHeader>
        On Error GoTo neesTimes_Click_Err

        '</EhHeader>
        Dim arxeio

        Dim a As Long

100     a = GetTickCount()

110     arxeio = "EID"

        Dim s         As String

        Dim DBF       As Database

        Dim sql       As New ADODB.Connection

        Dim rDBF      As Recordset

        Dim rSQL      As New ADODB.Recordset

        Dim R         As New ADODB.Recordset

        Dim conSQL    As String

        Dim k         As Long

        Dim Fname     As String

        Dim RECS      As Integer

        Dim lti       As Single

        Dim XTI       As Single

        Dim PROM      As String

        Dim KODLOG    As String

        Dim memo      As String

        Dim pos_kerd  As Single

        Dim pos_kerd2 As Single

        Dim MON       As String

        Dim SPA       As Single

        Dim POS01     As Single

        Dim POS02     As Single

        Dim maxWait   As Integer

120     maxWait = 1

130     If Len(Dir("C:\LAGEURO\EIDFTPg.bat")) > 1 Then
140         If Len(Dir("c:\lageuro\eid.dbf")) > 2 Then
150             Kill "c:\lageuro\eid.dbf"
            End If

160         DUM = Shell("C:\LAGEURO\EIDFTPg.bat", vbMaximizedFocus)
170         MILSEC 1000

180         Do While Len(Dir("c:\lageuro\eid.dbf")) < 2
190             MILSEC 1000
200             maxWait = maxWait + 1

210             If maxWait > 200 Then
220                 MsgBox "ΑΔΥΝΑΤΗ Η ΠΑΡΑΛΑΒΗ ΤΟΥ ΑΡΧΕΙΟΥ ΕΙΔΩΝ"

                    Exit Sub

                End If

            Loop

        End If

230     sql.Open gConnect

        Dim dB As DAO.Database

240     Set dB = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
250     Set rDBF = dB.OpenRecordset("EID")

260     rSQL.Open "SELECT * FROM EID ", Gdb, adOpenDynamic, adLockOptimistic

270     rDBF.MoveFirst

        Dim Z

280     Z = 0

        On Error GoTo WRITEERROR    ' Resume Next

        Dim F_EXO_XROMATA, TEX As String

290     F_EXO_XROMATA = Val(FINDPARAMETROI(1, "PAR1", "F_EXO_XROMATA", "0", "Γ.1 ΧΡΗΣΙΜΟΠΟΙΩ ΜΕΓΕΘΗ-ΧΡΩΜΑΤΑ=1"))   'posa psifia tha exei h kathe seira

        Dim erg As String, PONT As Long, ono As String

300     Do While Not rDBF.EOF
            '  rSQL.Open "SELECT COUNT(*) FROM EID where KOD='" + rDBF("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly
            '  S = "DSFSFDSFS" _
            '  & "FDSFSFDSFFDSF"

310         If IsNull(rDBF("PONTOI")) Then PONT = 0 Else PONT = rDBF("PONTOI")
320         If IsNull(rDBF("ERG")) Then erg = "." Else erg = Left(rDBF("ERG"), 14)
330         If IsNull(rDBF("ONO")) Then ono = "." Else ono = Replace(rDBF("ONO"), "'", "^")
340         If IsNull(rDBF("FPA")) Then GoTo OLD100
350         If IsNull(rDBF("LTI5")) Then GoTo OLD100
360         If IsNull(rDBF("LTI")) Then lti = 0 Else lti = rDBF("LTI")

370         If IsNull(rDBF("XTI")) Then XTI = 0 Else XTI = rDBF("XTI")
380         If IsNull(rDBF("PROM")) Then PROM = "." Else PROM = rDBF("PROM")
390         If IsNull(rDBF("KODLOG")) Then KODLOG = "." Else KODLOG = rDBF("KODLOG")

400         If IsNull(rDBF("POS_KERD")) Then pos_kerd = 0 Else pos_kerd = rDBF("POS_KERD")
410         If IsNull(rDBF("POS_KERD2")) Then pos_kerd2 = 0 Else pos_kerd2 = rDBF("POS_KERD2")

420         If IsNull(rDBF("POS01")) Then POS01 = 0 Else POS01 = rDBF("POS01")
430         If IsNull(rDBF("POS02")) Then POS02 = 0 Else POS02 = rDBF("POS02")

440         If IsNull(rDBF("POS03")) Then POS03 = 0 Else POS03 = rDBF("POS03")
450         If IsNull(rDBF("POS04")) Then POS04 = 0 Else POS04 = rDBF("POS04")

460         If IsNull(rDBF("KATHGORIA")) Then KATHGORIA = 0 Else KATHGORIA = rDBF("KATHGORIA")  '

470         If IsNull(rDBF("SPA")) Then SPA = 0 Else SPA = rDBF("SPA")

480         If IsNull(rDBF("mon")) Then MON = "TEM" Else MON = rDBF("MON")

490         Text = ""

500         If F_EXO_XROMATA = 1 Then
510             If IsNull(rDBF("memo")) Then memo = "" Else memo = rDBF("memo")
520             TEX = "MEMO='" + rDBF("MEMO") + "',"
            End If

            '  Dim R As New ADODB.Recordset

530         s = "update EID SET " + TEX + "" & "KOD='" + rDBF("KOD") + "'," & "ONO='" + ono + "'," & "ERG='" + erg + "'," & "LTI=" + str(lti) + "," & "LTI5=" + str(rDBF("LTI5")) + "," & "PONTOI=" + str(PONT) + "," & "XTI=" + str(XTI) + "," & "POS_KERD=" + str(pos_kerd) + "," & "POS_KERD2=" + str(pos_kerd2) + "," & "POS01=" + str(POS01) + "," & "POS02=" + str(POS02) + "," & "POS03=" + str(POS03) + "," & "POS04=" + str(POS04) + "," & "SPA=" + str(SPA) + "," & "PROM='" + PROM + "'," & "MON='" + MON + "'," & "KODLOG='" + KODLOG + "'," & "FPA=" + str(rDBF("FPA")) + " where KOD='" + rDBF("KOD") + "'"

540         Gdb.Execute s, RECS
OLD100:

550         If RECS = 0 Then
560             rSQL.AddNew
570             rSQL("KOD") = rDBF("KOD")
580             rSQL("ONO") = ono

590             rSQL("ERG") = erg
600             rSQL("LTI5") = rDBF("LTI5")
610             rSQL("LTI") = lti
620             rSQL("pontoi") = PONT
630             rSQL("FPA") = rDBF("FPA")
640             rSQL("XTI") = XTI
650             rSQL("PROM") = PROM
660             rSQL("KODLOG") = KODLOG

670             rSQL("pOS01") = POS01
680             rSQL("POS02") = POS02
690             rSQL("pOS03") = POS03
700             rSQL("POS04") = POS04

710             rSQL("pOS_KERD") = pos_kerd
720             rSQL("POS_KERD2") = pos_kerd2
730             rSQL("SPA") = SPA
740             rSQL("MON") = MON

750             rSQL.Update
            End If

            'R.Open "SELECT SUM(XRE),SUM(PIS) FROM EGGTIM WHERE KODE='" + rDBF("kod") + "'", Gdb, adOpenForwardOnly, adLockReadOnly

            'rSQL("synago0") = POS02 + (R(0) - R(1))
            'rSQL.Update

760         Z = Z + 1

770         If Z Mod 10 = 0 Then
780             Me.Caption = Z
            End If

790         rDBF.MoveNext

800         DoEvents

        Loop

810     rSQL.Close
820     rDBF.Close

830     Me.Caption = (GetTickCount() - a) / 1000

        Exit Sub

WRITEERROR:

        ' MHNYMA:
840     HandleError "bohu7:neestimes"

850     Resume Next

        'List2.Visible = True

        'List2.AddItem Err.Description
860     Resume Next

        '<EhFooter>
        Exit Sub

neesTimes_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.neesTimes_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.neesTimes_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub SendKiniseis_Click()

        '<EhHeader>
        On Error GoTo SendKiniseis_Click_Err

        '</EhHeader>

100     FileCopy "c:\lageuro\send\eggtim0.dbf", "c:\lageuro\send\eggtim.dbf"

        Dim DUM

110     If F_EXO_XROMATA = 1 Then
120         DUM = ADD_FIELD("EGGTIM", "TRP", "Char(4)")
130         Gdb.Execute "UPDATE EGGTIM  SET EGGTIM.TRP=TIM.TRP FROM EGGTIM INNER JOIN TIM ON TIM.HME=EGGTIM.HME AND EGGTIM.ATIM=TIM.ATIM   WHERE EGGTIM.HME>='" + Format(apo, "MM/DD/YYYY") + "' AND EGGTIM.HME<'" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'"
        End If

        ' apo SQL SERVER ΣΕ DBF ΑΡΧΕΙΑ ΠΟΥ ΗΔΗ ΥΠΑΡΧΟΥΝ
        Dim k    As Integer

        Dim R    As New ADODB.Recordset

        Dim MKOD As Integer


   List2.AddItem "ΤΟΠΙΚΟΣ"
DoEvents
          Gdb.Execute "update EGGTIM SET ATIM2='0', APOT=" + LTrim(str(Val(Left(APOT2.Text, 2)))) + " WHERE APOT IS NULL"
DoEvents

140     update_DBF_from_SQL "EGGTIM WHERE HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'", "EGGTIM", Gdb


        If Len(F_GCONN2) > 2 Then  'DSN=POS1;uid=sa;pwd=12345678
               List2.AddItem "ΤΟΠΙΚΟΣ 2"
             'On Error GoTo noconn2
            On Error GoTo 0
            gconn2.Open F_GCONN2
            
             gconn2.Execute "update EGGTIM SET PIS=POSO,ATIM2='1', APOT=" + LTrim(str(Val(Left(APOT2.Text, 2)))) + " WHERE APOT IS NULL"
             'On Error Resume Next
             update_DBF_from_SQL "EGGTIM WHERE HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'", "EGGTIM", gconn2
             gconn2.Close
             
        End If

        If Len(F_GCONN3) > 2 Then  'DSN=POS1;uid=sa;pwd=12345678
             List2.AddItem "ΤΟΠΙΚΟΣ 3"
             'On Error GoTo noconn2
            On Error GoTo 0
            gconn2.Open F_GCONN3
            
            gconn2.Execute "update EGGTIM SET PIS=POSO,ATIM2='2' , APOT=" + LTrim(str(Val(Left(APOT2.Text, 2)))) + " WHERE APOT IS NULL"
             'On Error Resume Next
             update_DBF_from_SQL "EGGTIM WHERE HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'", "EGGTIM", gconn2
             gconn2.Close
             
        End If
        
       If Len(F_GCONN4) > 2 Then  'DSN=POS1;uid=sa;pwd=12345678
             List2.AddItem "ΤΟΠΙΚΟΣ 4"
             'On Error GoTo noconn2
            On Error GoTo 0
            gconn2.Open F_GCONN4
            
            gconn2.Execute "update EGGTIM SET PIS=POSO,ATIM2='3', APOT=" + LTrim(str(Val(Left(APOT2.Text, 2)))) + " WHERE APOT IS NULL"
             'On Error Resume Next
             update_DBF_from_SQL "EGGTIM WHERE HME>='" + Format(apo, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "'", "EGGTIM", gconn2
             gconn2.Close
             
        End If
        
        





        '--------------- ΑΝΤΙΓΡΑΦΗ ΑΡΧΕΙΟΥ ΤΙΜ -------------------------------------------
        Dim dB As Database

150     Set dB = OpenDatabase("", False, False, gConnect)

        Dim DD As String

160     DD = "C:\LAGEURO\SEND\TIM.MDB"

        On Error Resume Next

170     Kill DD

180     DUM = CREATE_DATABASE(DD)

        Dim sql As String, SQL2 As String

        'SQL2 = COLLECT_FIELDS()

190     sql = "SELECT * INTO TIM  IN '" + DD + "'  FROM  TIM WHERE HME>=#" + Format(apo, "MM/DD/YYYY") + "# AND HME<#" + Format(DateAdd("D", 1, eos), "MM/DD/YYYY") + "#"

        On Error GoTo SendKiniseis_Click_Err

200     dB.Execute sql
        '------------------------------------------------------------

210     Gdb.Execute "UPDATE PARAMETROI SET TIMH='" + Format(apo, "DD/MM/YYYY") + "' WHERE FORMA='BOHU7' AND VAR='F_APO'"
220     Gdb.Execute "UPDATE PARAMETROI SET TIMH='" + Format(eos, "DD/MM/YYYY") + "' WHERE FORMA='BOHU7' AND VAR='F_EOS'"

230     If Len(Dir("C:\LAGEURO\EGGTIFTP.bat")) > 1 Then
240         MILSEC 1000
250         DUM = Shell("C:\LAGEURO\EGGTIFTP.bat", vbMaximizedFocus)
        End If

        '<EhFooter>
        Exit Sub

SendKiniseis_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.SendKiniseis_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.SendKiniseis_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
'Function COLLECT_FIELDS()
'    Dim k As Integer, Sql2 As String
'    Dim DBLAG As Database, R As Recordset
'    'On Error Resume Next
'    'Set DBLAG = OpenDatabase(DD)
'    'Set R = DB.OpenRecordset("TIM")
'
'
'    Set DBLAG = OpenDatabase("C:\LAGEURO\SEND\TIM.MDB")
'    Set R = DBLAG.OpenRecordset("TIM")
'
'    Sql2 = ""
'    For k = 0 To R.FIELDS.Count - 1
'        Sql2 = Sql2 + "," + R.FIELDS(k).Name
'    Next
'    COLLECT_FIELDS = mID(Sql2, 2, Len(Sql2) - 1)
'End Function

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Text1_KeyUp_Err

        '</EhHeader>
100     If KeyCode = 13 Then
110         List1.SetFocus
        End If

        '<EhFooter>
        Exit Sub

Text1_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Text1_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Text1_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text1_LostFocus()

        '<EhHeader>
        On Error GoTo Text1_LostFocus_Err

        '</EhHeader>
100     If Len(Text1.Text) < 13 Then Exit Sub

110     fr.Open "SELECT KOD,PONTOI FROM EID WHERE KOD='" + Text1.Text + "'", Gdb, adOpenDynamic, adLockOptimistic

120     If fr.EOF Then
130         MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ Η ΚΑΡΤΑ"
140         fr.Close

            Exit Sub

        End If

        On Error Resume Next

150     Label1.Caption = fr(1)
160     fr.Close
170     List1.Clear
    
180     fr.Open "SELECT top 100 HME,PONTOI FROM DIATAKT WHERE PELATHS='" + Text1.Text + "' ORDER BY HME DESC", Gdb, adOpenDynamic, adLockOptimistic

190     Do While Not fr.EOF
200         List1.AddItem Format(fr("HME"), "DD/MM/YYYY") + " ΠΟΝΤΟΙ:" + Format(fr("PONTOI"), "###0") '+ "  ΠΑΡ/ΚΟ:" + CNull(fR("ATIM"))
210         fr.MoveNext
        Loop
    
220     fr.Close

        '<EhFooter>
        Exit Sub

Text1_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Text1_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Text1_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text2_LostFocus()

        '<EhHeader>
        On Error GoTo Text2_LostFocus_Err

        '</EhHeader>
        Dim a As String

100     If Len(Text1.Text) <> 13 Then
110         MsgBox "λάθος κωδικός"

            Exit Sub

        End If

120     If Val(Text2.Text) = 0 Then

            Exit Sub

        End If

130     a = str(Val(Text2.Text))

        Dim sql12 As String

        SQL2 = "INSERT INTO DIATAKT ( ARIUMOS,AJIA,SEIRA,PELATHS,HME,ATIM,PONTOI) values ("
        SQL2 = SQL2 + "99,"
        SQL2 = SQL2 + "0,"
        SQL2 = SQL2 + "'XEIP',"
        SQL2 = SQL2 + Text1.Text + ","
        SQL2 = SQL2 + "'" + Format(Now, "MM/DD/YYYY") + "',"
        SQL2 = SQL2 + "'XEIP',"
        SQL2 = SQL2 + a + ")"

        On Error GoTo 0

        Gdb.Execute SQL2

        On Error Resume Next

140     Gdb.Execute "update EID set PONTOI=(CASE WHEN PONTOI IS NULL THEN 0 ELSE PONTOI END) + " + a + " where left(KOD,13)='" + Text1.Text + "'", FF

150     If FF < 1 Then
160         MsgBox "δεν ενημερώθηκε"

            Exit Sub

        End If

170     List1.AddItem Text1.Text + "--" + Text2.Text

180     Text2.Text = ""

        '<EhFooter>
        Exit Sub

Text2_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Text2_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Text2_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
'Private Sub update_sql_from_dbf(arxeio As String)
'    Dim DBF As Database
'    Dim sql As New ADODB.Connection
'    Dim rDBF As Recordset
'    Dim rSQL As New ADODB.Recordset
'    Dim conDBF As String
'    Dim conSQL As String
'    Dim k As Long
'    Dim Fname As String
'
'    G2 = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP4;Data Source=PLATO"
'    conDBF = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dBASE Files;Initial Catalog=C:\LAGEURO"
'
'    sql.Open gConnect
'
'    Data1.DatabaseName = "c:\lageuro\send"
'    Data1.RecordSource = arxeio
'    Data1.Refresh
'
'    Set rDBF = Data1.Recordset
'
'    Label1.Caption = arxeio + " " + Format(rDBF.RecordCount, "######")
'
'
'    rSQL.Open "SELECT *FROM " + arxeio, sql, adOpenDynamic, adLockOptimistic
'    rDBF.MoveFirst
'
'    Dim Z
'    Z = 0
'    On Error GoTo WRITEERROR    ' Resume Next
'
'    Do While Not rDBF.EOF
'
'        rSQL.AddNew
'
'        For k = 0 To rDBF.FIELDS.Count - 1
'            Fname = rDBF.FIELDS(k).Name    ' p.x. FNAME=epo    R(0).NAME
'
'            If IsNull(rDBF(k)) Then
'            Else
'                rSQL(Fname) = rDBF(k)    ' rsqk("epo")=r(0)
'            End If
'            ' ΑΥΤΑ ΠΟΥ ΕΙΝΑΙ ΝULL ΤΑ ΓΕΜΙΖΕΙ
'            If IsNull(rDBF(k)) Then
'                If rDBF(k).Type = dbText Then    '10=STRING
'                    rSQL(Fname) = " "
'                ElseIf rDBF(k).Type = dbDate Then
'                    rSQL(Fname) = CDate("01/01/1990")
'                ElseIf rDBF(k).Type = dbBoolean Then
'                    rSQL(Fname) = False
'                ElseIf rDBF(k).Type = dbDouble Then
'                    rSQL(Fname) = 0
'                Else
'                    conSQL = ""
'                End If
'            End If
'        Next
'
'        ' METΑΤΡΕΠΕΙ ΤΟ 'POSO' ΣΕ 'XRE' KAI 'PIS'
'        If arxeio = "EGGTIM" Then
'            rSQL("PIS") = 0: rSQL("XRE") = 0
'            If Not IsNull(rSQL("POSO")) Then
'                If InStr(F_xre_APOU + F_axre_APOU, Left(rSQL("ATIM"), 1)) > 0 Then
'                    rSQL("XRE") = rSQL("POSO")
'                End If
'                If InStr(F_pis_APOU + F_apis_APOU, Left(rSQL("ATIM"), 1)) > 0 Then
'                    rSQL("PIS") = rSQL("POSO")
'                End If
'            End If
'        End If
'
'        rSQL.Update
'        Z = Z + 1
'        If Z Mod 10 = 0 Then
'            Me.Caption = Z
'        End If
'        rDBF.MoveNext
'        DoEvents
'
'    Loop
'
'    rSQL.Close
'    rDBF.Close
'    Exit Sub
'
'
'WRITEERROR:
'    Resume Next
'End Sub
Private Sub calc_Click()

        'ΛΗΨΗ ΑΠΟ EXCEL
        ' ============================
        '<EhHeader>
        On Error GoTo Command8_Click_Err

        '</EhHeader>

'        Dim xl             As New excel.Application
'
'        Dim xlsheet        As excel.Worksheet
'
'        Dim xlsheet3       As excel.Worksheet
'
'        Dim xlwbook        As excel.workbook

'-------------------------------------------------  calc ----------------------------------------
Dim oSM
Dim oDesk, oDoc As Object
Dim arg()
Dim ocell As Object
Dim oSheet As Object
Dim importDesc(0 To 2) As Object
Dim oImportrange As Object
Dim oSQL As Object
Dim oContext As Object
Set oSM = CreateObject("com.sun.star.ServiceManager")
Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
Set oDoc = oDesk.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg())
' Set oSheet = oDoc.Sheets.getByName("Φύλλο1")
Set oSheet = oDoc.getSheets().getByIndex(0)
oSM.createInstance ("com.sun.star.sheet.DatabaseImportDescriptor ")
oSM.createInstance ("com.sun.star.util.Ximportable")












        Dim DBF            As Database

        Dim sql            As New ADODB.Connection

        Dim rDBF           As Recordset

        Dim rSQL           As New ADODB.Recordset

        Dim rSQL2          As New ADODB.Recordset

        Dim rYPOL          As New ADODB.Recordset

        Dim mYPOL          As Single

        Dim conDBF         As String

        Dim conSQL         As String

        Dim k              As Long

        Dim Fname          As String

        Dim dB             As DAO.Database

        Dim arxeio         As String, arxeio2 As String

        Dim SynStoYpoloipo As Integer

        Dim m_Proeleysh

        Dim M_ATIM

100     M_ATIM = "λ" + Format(time(), "hhmmss")

110     SynStoYpoloipo = MsgBox("Να προστεθεί στο ήδη υπάρχον υπόλοιπο;" + Chr(13) + "Σε περίπτωση που απαντήσετε όχι η ποσότητα που απογράψατε θα είναι και το υπόλοιπο", vbYesNoCancel)

120     If SynStoYpoloipo = vbCancel Then
130         MsgBox "Η εργασία ακυρώθηκε"

            Exit Sub

        End If

        Dim BARC1STILI As Integer

140     BARC1STILI = MsgBox("Η πρώτη στήλη είναι το BARCODE; " + Chr(13) + " Oxι σημαίνει ότι είναι ο κωδικός", vbYesNo)

        '====================================================

        'On Error GoTo open_error
150     Me.MousePointer = vbHourglass

160     If Len(Dir(FILENAME.Text, vbNormal)) > 0 Then
            ' OK
        Else
170         MsgBox " ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ ΑΡΧΕΙΟ " + FILENAME.Text

            Exit Sub

        End If

180     If Val(apot) = 0 Then
190         MsgBox " ΔΕΝ ΥΠΑΡΧΕΙ ΑΠΟΘΗΚΗ "

            Exit Sub

        End If

200     If SynStoYpoloipo = vbNo Then  ' αν προσαρμοσει τα υπολοιπα
            '   UPDATE_YPOLOIPA2 ""
        End If

210     Set xlwbook = xl.Workbooks.Open(FILENAME.Text)
220     Set xlsheet = xlwbook.Sheets.Item(1)

        'Set db = OpenDatabase(Dir1.Path, False, False, "dBase III;")

        'Set rDBF = db.OpenRecordset("SELECT * FROM " + File1.FILENAME)

230     rSQL.Open "SELECT top 1 * FROM EGGTIM ", Gdb, adOpenDynamic, adLockOptimistic

        'rDBF.MoveFirst

        Dim Z

240     Z = 0
250     k = 0

        On Error GoTo WRITEERROR    ' Resume Next

        Dim OK As Boolean

        Dim M_CODE

        Dim RR      As New ADODB.Recordset

        Dim RR2     As New ADODB.Recordset

        Dim TREX_YP As Single

        Dim nn      As Integer

        ' If IsNull(xlsheet.cells(mRow, KOD_COLUMN)) Then

        Dim FOUND   As Boolean

        Dim CC      As String

260     mRow = 1

270     Do While True  '   Not IsNull(xlsheet.cells(mRow, 1))

280         If IsNull(xlsheet.cells(mRow, 1)) Then

                Exit Do

            End If

290         If IsEmpty(xlsheet.cells(mRow, 1)) Then

                Exit Do

            End If

300         OK = True

310         M_CODE = xlsheet.cells(mRow, 1)
            '  On Error Resume Next

320         FOUND = True

330         If BARC1STILI = vbYes Then    'barcode 1η στήλη
340             RR.Open "select * FROM BARCODES WHERE ERG='" + M_CODE + "'", Gdb, adOpenDynamic, adLockOptimistic

350             If RR.EOF Then

360                 FOUND = False
370                 MsgBox "ΔΕΝ ΒΡΕΘΗΚΕ ΤΟ BARCODE " + M_CODE
380                 List2.AddItem "ΔΕΝ ΒΡΕΘΗΚΕ : " + M_CODE
390                 xlsheet.cells(mRow, 4) = "ΔΕΝ ΒΡΕΘΗΚΕ"

                End If

400             M_CODE = RR("KOD")
410             m_Proeleysh = RR("Proeleysh")
420             RR.Close

            End If

430         If SynStoYpoloipo = vbNo Then

440             sql = "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS YPOL FROM EGGTIM WHERE KODE='" + M_CODE + "'  and APOT=" + Left(apot, 1)
                'RR2.Open sql, Gdb, adOpenDynamic, adLockOptimistic

450             RR2.Open "select SUM(CASE WHEN XRE IS NULL THEN 0 ELSE XRE END)-SUM(CASE WHEN PIS IS NULL THEN 0 ELSE PIS END) AS YPOL FROM EGGTIM WHERE KODE='" + M_CODE + "'  and APOT=" + Left(apot, 1), Gdb, adOpenDynamic, adLockOptimistic

460             TREX_YP = RR2("YPOL")
470             RR2.Close
            Else
                ' rSQL("poso") = xlsheet.cells(mRow, 2)
480             TREX_YP = 0
            End If

490         If FOUND Then

                '          rSQL.AddNew
                '          rSQL("hme") = apog.value
                '          rSQL("ATIM") = M_ATIM '"λ00002"
                '          rSQL("KODE") = M_CODE
                '          rSQL("timm") = xlsheet.cells(mRow, 3)
                '          rSQL("Proeleysh") = Left(m_Proeleysh, 10)
                '          rSQL("poso") = -TREX_YP + xlsheet.cells(mRow, 2)
                '          rSQL("XRE") = -TREX_YP + xlsheet.cells(mRow, 2)
                '          rSQL("APOT") = Val(APOT.Text)
                '         rSQL.Update
                '

500             CC = "INSERT INTO EGGTIM (HME,ATIM,KODE,TIMM,PROELEYSH,POSO,XRE,APOT) VALUES"
510             CC = CC + "('" + Format(apog.Value, "MM/DD/YYYY") + "',"
520             CC = CC + "'" + M_ATIM + "',"
530             CC = CC + "'" + M_CODE + "',"
540             CC = CC + "" + str(xlsheet.cells(mRow, 3)) + ","
550             CC = CC + "'" + Left(m_Proeleysh, 10) + "',"
560             CC = CC + "" + str(-TREX_YP + xlsheet.cells(mRow, 2)) + ","
570             CC = CC + "" + str(-TREX_YP + xlsheet.cells(mRow, 2)) + ","
580             CC = CC + "" + str(Val(apot.Text)) + ")"
590             Gdb.Execute CC, nn

600             Gdb.Execute "UPDATE EID SET HM1=GETDATE() WHERE KOD='" + M_CODE + "'", nn

610             If nn = 0 Then
                    ' MsgBox "προσοχη δεν βρέθηκε ο κωδικός " + M_CODE

620                 List2.AddItem "ΔΕΝ ΒΡΕΘΗΚΕ : " + M_CODE
630                 xlsheet.cells(mRow, 4) = "ΔΕΝ ΒΡΕΘΗΚΕ"
                End If
            End If

640         mRow = mRow + 1

650         If mRow Mod 10 = 0 Then
660             Me.Caption = mRow
            End If

670         DoEvents
        Loop

680     rSQL.Close

690     rSQL.Open "SELECT * FROM TIM ", Gdb, adOpenDynamic, adLockOptimistic

700     rSQL.AddNew
710     rSQL("hme") = apog.Value
720     rSQL("ATIM") = M_ATIM    ' "λ00001"
        'rSQL("KODE") = M_CODE
        'rSQL("timm") = rDBF("xti")
        ' rSQL("poso") = rDBF("POS") - mYPOL
        ' rSQL("xre") = rDBF("POS") - mYPOL
730     rSQL("b_n1") = Val(apot.Text)
740     rSQL.Update

        On Error GoTo Command8_Click_Err

750     If List2.ListCount > 0 Then
760         MsgBox "ΤΑ ΛΑΘΗ ΑΠΟΘΗΚΕΥΤΗΚΑΝ ΣΤΟ EXCEL  "
770         xlwbook.save  '  "c:\ektyp2.xls"
        End If

780     Me.MousePointer = vbNormal

790     Set xlsheet = Nothing
800     Set xlwbook = Nothing
        ' excel.Quit

810     xl.Quit

820     MsgBox "Ενημερώθηκαν " + Format(mRow, "###0") + " εγγραφές"

        Exit Sub

WRITEERROR:

830     Resume Next

        '<EhFooter>
        Exit Sub

Command8_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu7.Command8_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu7.Command8_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub


