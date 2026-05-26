VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form bohu7 
   BackColor       =   &H00FF0000&
   Caption         =   "Form2"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12150
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   12150
   WindowState     =   2  'Maximized
   Begin VB.ComboBox apot 
      Height          =   315
      Left            =   7920
      TabIndex        =   31
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton mhdenismos 
      Caption         =   "Μηδενισμός Αποθήκης με κινήσεις πρόχειρης απογραφής"
      Height          =   735
      Left            =   8760
      TabIndex        =   30
      ToolTipText     =   "Θα μηδενιστεί η παρακάτω αποθήκη με την παρακάτω ημερομηνία Απογραφής "
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton NEABARCODES 
      Caption         =   "ΝΕΑ BARCODES"
      Height          =   480
      Left            =   2715
      TabIndex        =   29
      Top             =   4230
      Width           =   1050
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   75
      TabIndex        =   28
      Top             =   7305
      Width           =   9255
   End
   Begin VB.CheckBox kodisbarcode 
      Caption         =   "Kωδικός είναι το barcode"
      Height          =   375
      Left            =   8475
      TabIndex        =   27
      Top             =   4785
      Width           =   2340
   End
   Begin MSComCtl2.DTPicker apog 
      Height          =   270
      Left            =   7980
      TabIndex        =   25
      Top             =   930
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   476
      _Version        =   393216
      Format          =   67698689
      CurrentDate     =   38727
   End
   Begin VB.TextBox FILENAME 
      Height          =   285
      Left            =   8475
      TabIndex        =   23
      Top             =   4455
      Width           =   2340
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   8445
      TabIndex        =   22
      Top             =   3330
      Width           =   2370
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   8430
      TabIndex        =   21
      Top             =   1815
      Width           =   2385
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ΛΗΨΗ ΑΠΟΓΡΑΦΗΣ (APOG.DBF)"
      Height          =   735
      Left            =   6720
      TabIndex        =   20
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2640
      TabIndex        =   19
      Text            =   "\\server\c\tameia\diatakt.mdb"
      Top             =   5190
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "E ξ ο δ ο ς"
      Height          =   615
      Left            =   3360
      TabIndex        =   17
      Top             =   6510
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Λήψη κινήσεων από υποκατάστημα"
      Height          =   495
      Left            =   3360
      TabIndex        =   16
      Top             =   5790
      Width           =   1935
   End
   Begin VB.CommandButton SendKiniseis 
      Caption         =   "Αποστολή Κινήσεων"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   5790
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Αποστολή πόντων"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   4950
      Width           =   2295
   End
   Begin VB.CommandButton neesTimes 
      Caption         =   "Ενημέρωση με νέες τιμές"
      Height          =   495
      Left            =   240
      TabIndex        =   9
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
      Height          =   1035
      Left            =   480
      TabIndex        =   6
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   1200
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
      Format          =   67698689
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
      Format          =   67698689
      CurrentDate     =   38604
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
      Height          =   345
      Left            =   7080
      TabIndex        =   26
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Ημ.Απογραφής"
      Height          =   270
      Left            =   6735
      TabIndex        =   24
      Top             =   930
      Width           =   1200
   End
   Begin VB.Label Label8 
      Caption         =   "Διαδρομή Πόντων"
      Height          =   255
      Left            =   2640
      TabIndex        =   18
      Top             =   4830
      Width           =   2295
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
   Begin VB.Shape Shape1 
      Height          =   2895
      Left            =   240
      Top             =   240
      Width           =   6375
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
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Βarcode κάρτας"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "bohu7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fR As New ADODB.Recordset
Dim F_APO As String
Dim F_EOS As String


Private Sub Command1_Click()
' ENHMERCSH PONTON KYRIOY ARXEIOY
'Dim r As New ADODB.Recordset
'Dim r2 As New ADODB.Recordset
Dim db As DAO.Database
Dim R As DAO.Recordset

'Set db = OpenDatabase("C:\LAGEURO\SEND", False, False, "dBase III;")
Set db = OpenDatabase("C:\LAGEURO\SEND", False, False, "dBase III;")

xx = "SELECT * FROM DIATAKT WHERE  HME >=#" + Format(APO, "MM/DD/YYYY") + "# AND HME<#" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY") + "# ORDER BY ΗΜΕ "
Set R = db.OpenRecordset("DIATAKT")
R.MoveFirst
Do While Not R.EOF
   If IsNull(R("PONTOI")) Or IsNull(R("PELATHS")) Then
     'ASE
   Else
      Gdb.Execute "UPDATE EID SET PONTOI=PONTOI+" + Str(R("PONTOI")) + " WHERE LEFT(KOD,13)='" + R("PELATHS") + "'  "
   End If
   Me.Caption = R("hme")
   R.MoveNext
   DoEvents
   
   
Loop

R.Close
db.Close
Exit Sub






End Sub


Private Sub update_DBF_from_SQL(arxeio As String, ARXEIOdbf As String)

Dim DBF As Database
Dim sql As New ADODB.Connection
Dim rDBF As Recordset
Dim rSQL As New ADODB.Recordset
Dim conDBF As String
Dim conSQL As String
Dim K As Long
Dim Fname As String

Dim db As DAO.Database
Set db = OpenDatabase("C:\LAGEURO\SEND", False, False, "dBase III;")

Set rDBF = db.OpenRecordset(ARXEIOdbf)





rSQL.Open "SELECT *FROM " + arxeio, Gdb, adOpenDynamic, adLockOptimistic
' rDBF.Open "SELECT *FROM " + arxeio, DBF, adOpenDynamic, adLockOptimistic

rSQL.MoveFirst

Dim Z
Z = 0
On Error GoTo WRITEERROR ' Resume Next

Do While Not rSQL.EOF
  
  rDBF.AddNew
  
  For K = 0 To rDBF.FIELDS.Count - 1
      Fname = rDBF.FIELDS(K).Name ' p.x. FNAME=epo    R(0).NAME
      
'      If IsNull(rDBF(k)) Then
'      Else
         rDBF(Fname) = rSQL(Fname) ' rsqk("epo")=r(0)
'      End If
  Next
  rDBF.Update
    Z = Z + 1
    If Z Mod 10 = 0 Then
      Me.Caption = Z
    End If
  rSQL.MoveNext
  DoEvents
  
Loop

rSQL.Close
rDBF.Close



db.Execute "update EGGTIM SET FPA=2 WHERE FPA=23" ' 19%
db.Execute "update EGGTIM SET FPA=4 WHERE FPA=21" ' 4.5%
db.Execute "update EGGTIM SET FPA=1 WHERE FPA=22" ' 9%

Set rDBF = db.OpenRecordset("SELECT SUM(MIK_AJIA) FROM EGGTIM")
MsgBox "Σύνολο αξίας " + Format(rDBF(0), "####,##0.00")
Exit Sub


WRITEERROR:
'List2.Visible = True

'List2.AddItem Err.Description
Resume Next




End Sub

Private Sub Command2_Click()


Set db = OpenDatabase(Text3.Text)

On Error Resume Next

Kill "C:\LAGEURO\SEND\DIATAKT.DBF"

db.Execute "select * into DIATAKT in 'C:\LAGEURO\SEND' 'dBase III;' from diataktikes where  HME>=#" + Format(APO, "MM/DD/YYYY") + "# AND HME<#" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY") + "#"


Set db = OpenDatabase("C:\LAGEURO\SEND", False, False, "dBase III;")



Set R = db.OpenRecordset("SELECT SUM(pontoi) FROM diatakt")
MsgBox "Σύνολο Πόντων  " + Chr(13) + Format(R(0), "####,##0")




End Sub

Private Sub Command3_Click()




Dim DBF As Database
Dim sql As New ADODB.Connection
Dim rDBF As Recordset
Dim rSQL As New ADODB.Recordset
Dim conDBF As String
Dim conSQL As String
Dim K As Long
Dim Fname As String

Dim db As DAO.Database

Dim arxeio As String, arxeio2 As String
Dim ans As Integer



 arxeio = "EGGTIM WHERE APOT=2 AND LEFT(ATIM,1)='L' and HME>='" + Format(APO, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY") + "'"
 ARXEIOdbf = "EGGTIM"





Set db = OpenDatabase("C:\LAGEURO\SEND", False, False, "dBase III;")

Set rDBF = db.OpenRecordset("SELECT SUM(MIK_AJIA) FROM EGGTIM where apot=2")
ans = MsgBox("Σύνολο αξίας " + Format(rDBF(0), "####,##0.00") + Chr(13) + " Συνεχίζω ;", vbYesNo)
If ans = vbNo Then
   rDBF.Close
   db.Close
   
   Exit Sub
End If




rDBF.Close









Set rDBF = db.OpenRecordset("SELECT * FROM EGGTIM WHERE HME>=#" + Format(APO, "MM/DD/YYYY") + "# AND HME<#" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY") + "#")
' ελεγχος αν υπάρχουν ήδη τα παραστατικά
rSQL.Open "SELECT count(*) FROM " + arxeio, Gdb, adOpenDynamic, adLockOptimistic


If rSQL(0) > 0 Then
  
   'MsgBox "υπάρχουν ήδη  " + Str(rSQL(0)) + "κινήσεις.   Δεν θα ενημερωθεί το αρχείο"
   'rSQL.Close
   ' αν πρόκειται για μία μέρα μόνο να έχει δυνατότητα διαγραφής
   If APO = EOS Then
      ans = MsgBox("Να διαγράψω τις ήδη υπάρχουσες κινήσεις του υποκαταστήματος για την ημέρα " + Format(APO, "dd/mm/yyyy"), vbYesNo)
      If ans = vbNo Then
         Exit Sub
      End If
      Dim d
      Gdb.Execute "DELETE FROM EGGTIM WHERE (NOT LEFT(ATIM,1)='σ')  AND APOT=2 AND HME>='" + Format(APO, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY") + "'", d
      
      Gdb.Execute "DELETE FROM EGGTIM WHERE  LEFT(ATIM,1)='σ' AND APOT=2 and XRE<0 AND HME>='" + Format(APO, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY") + "'", d
      
      Gdb.Execute "DELETE FROM EGGTIM WHERE  LEFT(ATIM,1)='σ' AND APOT=1 and XRE>0 AND HME>='" + Format(APO, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY") + "'", d
      
      
      MsgBox "Διεγράφησαν " + Format(d, "####") + " εγγραφές "
    Else
       MsgBox "υπάρχουν ήδη  " + Str(rSQL(0)) + "κινήσεις.   Δεν θα ενημερωθεί το αρχείο"
       rSQL.Close
       Exit Sub
    End If
    
End If
rSQL.Close
rSQL.Open "SELECT * FROM EGGTIM ", Gdb, adOpenDynamic, adLockOptimistic

rDBF.MoveFirst

Dim Z
Z = 0
On Error GoTo WRITEERROR ' Resume Next

Do While Not rDBF.EOF
  
  rSQL.AddNew
  
  For K = 0 To rDBF.FIELDS.Count - 1
      Fname = rDBF.FIELDS(K).Name ' p.x. FNAME=epo    R(0).NAME
      
      If IsNull(rDBF(K)) Then
      Else
         rSQL(Fname) = rDBF(Fname) ' rsqk("epo")=r(0)
      End If
  Next
  
  If left(rSQL("ATIM"), 1) = "L" Then
     'επειδη των ταμειακων δεν εχουν "xre" kai "pis"
     rSQL("PIS") = 0: rSQL("XRE") = 0
     rSQL("PIS") = rSQL("POSO")
    
     
  End If
  
   rSQL("kola") = 2 ' σημαινει ότι ηρθαν απο το υποκαταστημα
  
   ' ΟΤΑΝ ΓΙΝΕΤΑΙ ΕΠΙΣΤΡΟΦΗ ΑΠΟ ΤΗΝ ΤΑΜΕΙΑΚΗ ΕΧΕΙ ΑΡΝΗΤΙΚΗ ΤΙΜΗ "ΤΙΜΜ"
   ' ΟΠΟΤΕ ΤΗΝ ΚΑΝΩ ΘΕΤΙΚΗ ΚΑΙ ΚΑΝΩ ΑΡΝΗΤΙΚΗ ΤΗΝ ΠΟΣΟΤΗΤΑ
   If rSQL("TIMM") < 0 Then
      rSQL("TIMM") = -rSQL("TIMM")
      rSQL("POSO") = -Abs(rSQL("POSO"))
      rSQL("XRE") = 0
      rSQL("PIS") = rSQL("POSO")
   End If
   
  rSQL.Update
    Z = Z + 1
    If Z Mod 10 = 0 Then
      Me.Caption = Z
    End If
  rDBF.MoveNext
  DoEvents
  
Loop

rSQL.Close
rDBF.Close


Gdb.Execute "UPDATE PARAMETROI SET TIMH='" + Format(APO, "DD/MM/YYYY") + "' WHERE FORMA='BOHU7' AND VAR='F_APO'"
Gdb.Execute "UPDATE PARAMETROI SET TIMH='" + Format(EOS, "DD/MM/YYYY") + "' WHERE FORMA='BOHU7' AND VAR='F_EOS'"




Exit Sub


WRITEERROR:
'List2.Visible = True

'List2.AddItem Err.Description
Resume Next















End Sub

Private Sub Command4_Click()
   Unload Me
   SendKeys "%"
      SendKeys "{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}+{DOWN}"

   
   
End Sub

Private Sub Command5_Click()
' ΛΗΨΗ ΑΠΟΓΡΑΦΗΣ ============================

Dim DBF As Database
Dim sql As New ADODB.Connection
Dim rDBF As Recordset
Dim rSQL As New ADODB.Recordset
Dim rSQL2 As New ADODB.Recordset
Dim conDBF As String
Dim conSQL As String
Dim K As Long
Dim Fname As String

Dim db As DAO.Database

Dim arxeio As String, arxeio2 As String
Dim ans As Integer




If Len(Dir(FILENAME.Text, vbNormal)) > 0 Then
  ' OK
Else
   MsgBox " ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ ΑΡΧΕΙΟ " + FILENAME.Text
   Exit Sub
End If

If Val(apot) = 0 Then
   MsgBox " ΔΕΝ ΥΠΑΡΧΕΙ ΑΠΟΘΗΚΗ "
   Exit Sub
End If


Set db = OpenDatabase(Dir1.Path, False, False, "dBase III;")

Set rDBF = db.OpenRecordset("SELECT * FROM " + File1.FILENAME)





rSQL.Open "SELECT * FROM EGGTIM ", Gdb, adOpenDynamic, adLockOptimistic

rDBF.MoveFirst

Dim Z
Z = 0
On Error GoTo WRITEERROR ' Resume Next

Dim OK As Boolean
Dim M_CODE

Do While Not rDBF.EOF
  OK = True
  
  M_CODE = rDBF("barcode")
  
  If kodisbarcode.Value Then
  Else
    ' rSQL2.Open "SELECT * FROM EID WHERE LEFT(ERG,13)='" + left(rDBF("BARCODE"), 13) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
    rSQL2.Open "SELECT * FROM BARCODES WHERE LEFT(ERG,13)='" + left(rDBF("BARCODE"), 13) + "'", Gdb, adOpenForwardOnly, adLockReadOnly
    If rSQL2.EOF Then
       OK = False
       List2.AddItem rDBF("BARCODE") + ": ΔΕΝ ΥΠΑΡΧΕΙ ΣΤΟ ΑΡΧΕΙΟ ΕΙΔΩΝ"
    Else
       M_CODE = rSQL2("KOD")
    End If
    rSQL2.Close
  End If
  
    If OK And (Not IsNull(rDBF("YPOK"))) Then
       If Val(rDBF("YPOK")) <> 33 Then
          rSQL.AddNew
          rSQL("hme") = apog.Value
          rSQL("ATIM") = "λ00001"
          rSQL("KODE") = M_CODE
          rSQL("poso") = rDBF("POS")
          rSQL("xre") = rDBF("POS")
          rSQL("APOT") = Val(apot.Text)
         rSQL.Update
    End If
  End If
  ' rSQL.Update
    Z = Z + 1
    If Z Mod 10 = 0 Then
      Me.Caption = Z
    End If
  rDBF.MoveNext
  DoEvents
  
Loop

rSQL.Close
rDBF.Close





Exit Sub


WRITEERROR:
'List2.Visible = True

'List2.AddItem Err.Description
Resume Next

End Sub


Private Sub Dir1_Change()
   File1.Path = Dir1.Path
   File1.Refresh
End Sub

Private Sub File1_Click()
   FILENAME.Text = File1.Path + "\" + File1.FILENAME
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then Command4_Click
  
End Sub

Private Sub Form_Load()
Me.KeyPreview = True
APO = Now
EOS = Now
apog = Now
Dir1.Path = "C:\LAGEURO"
Dim R As New ADODB.Recordset

'αποθηκη
 R.Open "SELECT *FROM PINAKES WHERE TYPOS=4 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
Do While Not R.EOF
  If R("typos") = 4 Then
     apot.AddItem Format(R("ayjon"), "##") + "." + R("PERIGRAFH")
  Else
     Exit Do
  End If
  R.MoveNext
Loop
apot.Text = apot.List(0)

R.Close
F_APO = FindParametroi("BOHU7", "F_APO", Format(Now, "DD/MM/YYYY"), "ΗΜΕΡ.ΕΝΗΜΕΡΩΣΗΣ ΑΠΟ ")
F_EOS = FindParametroi("BOHU7", "F_EOS", Format(Now, "DD/MM/YYYY"), "ΗΜΕΡ.ΕΝΗΜΕΡΩΣΗΣ ΕΩΣ ")


APO = CDate(F_APO)
EOS = CDate(F_EOS)


If gXEIRISTHS < 9 Then
  'Dim K
    kodisbarcode.Visible = False
    FILENAME.Visible = False
    File1.Visible = False
    Dir1.Visible = False
    apog.Visible = False
    Label9.Visible = False
    
    Label1.Visible = False
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label10.Visible = False
    Label8.Visible = False
    Command5.Visible = False
    'Command8.Visible = False
    apot.Visible = False
    Text3.Visible = False
    Text1.Visible = False
    Text2.Visible = False
    
    
    List1.Visible = False
    

     mhdenismos.Visible = False
     Command5.Visible = False
     Command2.Visible = False
     Command1.Visible = False
End If



End Sub


Sub UPDATE_YPOLOIPA2(ByVal synt) ' ΣΥΝΘΗΚΗ  ""=ΟΛΑ   "KODLOG=1 AND AEG=11"  ΜΟΝΟ ΤΗΝ ΚΑΤΗΓΟΡΙΑ 1 ΚΑΙ 11 ΥΠΟΚΑΤΗΓΟΡΙΑ
'υπολογιζει τα υπόλοιπα της αποθήκης με ημερομηνία  απογραφής
Dim A As String
Dim B As String
Dim d As Long
'Dim krit As String

' ΒΡΙΣΚΩ ΜΙΑ ΦΟΡΑ ΤΑ ΥΠΟΛΟΙΠΑ ΑΝΑ ΑΠΟΘΗΚΗ

On Error Resume Next

Gdb.Execute "DROP TABLE DOKEGGT1"  'SOUMES ΑΠΟΘΗΚΗΣ 1

On Error GoTo 0

A = "SELECT KODE,SUM( CASE APOT  WHEN 1  THEN XRE ELSE 0  END ) AS S1X," _
& " SUM(CASE APOT  WHEN 1  THEN PIS ELSE 0  END ) AS S1P," _
& "SUM( CASE APOT  WHEN 2 THEN XRE ELSE 0  END ) AS S2X," _
& "SUM( CASE APOT WHEN 2 THEN PIS ELSE 0 END ) AS S2P INTO DOKEGGT1 "
A = A + " FROM EGGTIM INNER JOIN EID ON EGGTIM.KODE=EID.KOD WHERE " + synt + " APOT=" + left(apot.Text, 1) + " AND  HME<='" + Format(apog.Value, "MM/DD/YYYY") + "' GROUP BY KODE"
Gdb.Execute A, d

If d = 0 Then
   MsgBox " ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ DOKEGGT1"
   Exit Sub
End If


End Sub





Private Sub mhdenismos_Click()
   
  Dim ans
  
  If Val(apot.Text) = 0 Then
     MsgBox " Δεν δώσατε Αποθήκη"
     Exit Sub
  End If
  
  
  MsgBox "Θα μας ρωτήσει ποιές οικογένειες και υποοικογένειες (αριθμητικώς) θέλουμε να σβήσουμε και θα μας ζητηθεί επιβεβαίωση στο τέλος"
  
  
  Dim ypooik, oik
  
  ypooik = InputBox("Να μηδενιστεί η υπο οικογένεια ", "με 0 όλες οι υποοικογένειες")
  
  oik = InputBox("Να μηδενιστεί η οικογένεια ", "με 0 όλες οι οικογένειες")
  
  
  Dim synt
  synt = " EID.KODLOG=" + ypooik + " AND EID.AEG=" + oik + " AND "
  
  
  
  
  
  ans = MsgBox("Θα μηδενιστεί η αποθήκη " + apot.Text + " με  ημερομηνία Απογραφής " + Format(apog.Value, "dd/mm/yyyy"), vbYesNo, "ΕΙΣΘΕ ΣΙΓΟΥΡΟΙ;")
  If Not ans = vbYes Then
     Exit Sub
  End If
  
  
  UPDATE_YPOLOIPA2 synt 'YPOLOGISE YPOLOIPA KAI ΦΕΡΤΑ ΣΤΟΠ ΑΡΧΕΙΟ DOKEGGT1
  
  
  Dim R As New ADODB.Recordset
  Dim REGGTIM As New ADODB.Recordset
  
  Dim yp As Single
  
  R.Open "SELECT * FROM DOKEGGT1 ", Gdb, adOpenForwardOnly, adLockReadOnly
  REGGTIM.Open "SELECT * FROM EGGTIM ", Gdb, adOpenDynamic, adLockOptimistic
  
  R.MoveFirst
  
  Do While Not R.EOF
     
     'YPOLOGIZV TO YPOLOIPO APO TO ARXEIO DOKEGGT1 (YP)
     If Val(apot.Text) = 1 Then
        yp = R("S1X") - IIf(IsNull(R("S1P")), 0, R("S1P"))
     Else
        yp = R("S2X") - IIf(IsNull(R("S2P")), 0, R("S2P")) 'R("S2P")
     End If
      
      
     'AN DEN EINAI 0 TOTE PROSTHETO ή αφαιρώ για να γίνει 0 το υπολοιπό του
      If yp <> 0 Then
          REGGTIM.AddNew
          REGGTIM("hme") = apog.Value
          REGGTIM("ATIM") = "λ00003"
          REGGTIM("KODE") = R("KODE")
          REGGTIM("poso") = -yp
          REGGTIM("xre") = -yp
          REGGTIM("APOT") = Val(apot.Text)
          REGGTIM.Update
      End If
      Me.Caption = R("KODE")
      R.MoveNext
      DoEvents
  Loop
  R.Close
  

End Sub

Private Sub NEABARCODES_Click()
On Error Resume Next

Gdb.Execute "DROP TABLE BARCODES0"

On Error GoTo BGES

Dim db As DAO.Database
Set db = OpenDatabase("", False, False, gConnect)

db.Execute "SELECT * INTO BARCODES0 FROM BARCODES IN 'C:\LAGEURO' 'dBASE III;' "
On Error Resume Next
On Error GoTo 0
Gdb.Execute "INSERT INTO BARCODES SELECT * FROM BARCODES0 WHERE ERG NOT IN ( SELECT ERG FROM BARCODES)"
Gdb.Execute "DELETE FROM BARCODES WHERE ERG NOT IN ( SELECT ERG FROM BARCODES0 )"

Exit Sub

BGES:
MsgBox "ΔΕΝ ΕΝΗΜΕΡΩΘΗΚΕ"



End Sub

Private Sub neesTimes_Click()
' METAFORA apo ARXEIO DBF
Dim arxeio
Dim A As Long
A = GetTickCount()


arxeio = "EID"


Dim s As String
Dim DBF As Database
Dim sql As New ADODB.Connection
Dim rDBF As Recordset
Dim rSQL As New ADODB.Recordset
Dim R As New ADODB.Recordset
Dim conSQL As String
Dim K As Long
Dim Fname As String
Dim RECS As Integer
Dim LTI As Single
Dim xti As Single
Dim PROM As String
Dim KODLOG As String


Dim pOS_KERD As Single
Dim pOS_KERD2 As Single
Dim MON As String
Dim SPA As Single

Dim POS01 As Single

Dim POS02 As Single



sql.Open gConnect


Dim db As DAO.Database
Set db = OpenDatabase("C:\LAGEURO", False, False, "dBase III;")
Set rDBF = db.OpenRecordset("EID")

10
rSQL.Open "SELECT * FROM EID ", Gdb, adOpenDynamic, adLockOptimistic

rDBF.MoveFirst

Dim Z
Z = 0
On Error GoTo WRITEERROR ' Resume Next

20
Dim erg As String, PONT As Long, ono As String

Do While Not rDBF.EOF
'  rSQL.Open "SELECT COUNT(*) FROM EID where KOD='" + rDBF("KOD") + "'", Gdb, adOpenForwardOnly, adLockReadOnly
'  S = "DSFSFDSFS" _
'  & "FDSFSFDSFFDSF"
 
If IsNull(rDBF("PONTOI")) Then PONT = 0 Else PONT = rDBF("PONTOI")
If IsNull(rDBF("ERG")) Then erg = "." Else erg = left(rDBF("ERG"), 14)
If IsNull(rDBF("ONO")) Then ono = "." Else ono = Replace(rDBF("ONO"), "'", "^")
If IsNull(rDBF("FPA")) Then GoTo 100
If IsNull(rDBF("LTI5")) Then GoTo 100
If IsNull(rDBF("LTI")) Then LTI = 0 Else LTI = rDBF("LTI")
30
If IsNull(rDBF("XTI")) Then xti = 0 Else xti = rDBF("XTI")
If IsNull(rDBF("PROM")) Then PROM = "." Else PROM = rDBF("PROM")
If IsNull(rDBF("KODLOG")) Then KODLOG = "." Else KODLOG = rDBF("KODLOG")
40
If IsNull(rDBF("POS_KERD")) Then pOS_KERD = 0 Else pOS_KERD = rDBF("POS_KERD")
If IsNull(rDBF("POS_KERD2")) Then pOS_KERD2 = 0 Else pOS_KERD2 = rDBF("POS_KERD2")
50
If IsNull(rDBF("POS01")) Then POS01 = 0 Else POS01 = rDBF("POS01")
If IsNull(rDBF("POS02")) Then POS02 = 0 Else POS02 = rDBF("POS02")
60
If IsNull(rDBF("SPA")) Then SPA = 0 Else SPA = rDBF("SPA")
70
If IsNull(rDBF("mon")) Then MON = "TEM" Else MON = rDBF("MON")
80
    
'  Dim R As New ADODB.Recordset



 
 
 
 s = "update EID SET " _
  & "KOD='" + rDBF("KOD") + "'," _
  & "ONO='" + ono + "'," _
  & "ERG='" + erg + "'," _
  & "LTI=" + Str(LTI) + "," _
  & "LTI5=" + Str(rDBF("LTI5")) + "," _
  & "PONTOI=" + Str(PONT) + "," _
  & "XTI=" + Str(xti) + "," _
  & "POS_KERD=" + Str(pOS_KERD) + "," _
  & "POS_KERD2=" + Str(pOS_KERD2) + "," _
  & "POS01=" + Str(POS01) + "," _
  & "POS02=" + Str(POS02) + "," _
  & "SPA=" + Str(SPA) + "," _
  & "PROM='" + PROM + "'," _
  & "MON='" + MON + "'," _
  & "KODLOG='" + KODLOG + "'," _
  & "FPA=" + Str(rDBF("FPA")) + " where KOD='" + rDBF("KOD") + "'"
90
  Gdb.Execute s, RECS
100
  
  
If RECS = 0 Then
  rSQL.AddNew
     rSQL("KOD") = rDBF("KOD")
     rSQL("ONO") = ono
110
     rSQL("ERG") = erg
     rSQL("LTI5") = rDBF("LTI5")
     rSQL("LTI") = LTI
     rSQL("pontoi") = PONT
     rSQL("FPA") = rDBF("FPA")
     rSQL("XTI") = xti
     rSQL("PROM") = PROM
     rSQL("KODLOG") = KODLOG
120
     rSQL("pOS01") = POS01
     rSQL("POS02") = POS02

130

     rSQL("pOS_KERD") = pOS_KERD
     rSQL("POS_KERD2") = pOS_KERD2
     rSQL("SPA") = SPA
     rSQL("MON") = MON


140


  rSQL.Update
End If
    
'R.Open "SELECT SUM(XRE),SUM(PIS) FROM EGGTIM WHERE KODE='" + rDBF("kod") + "'", Gdb, adOpenForwardOnly, adLockReadOnly
    

'rSQL("synago0") = POS02 + (R(0) - R(1))
'rSQL.Update

    
    
145
    Z = Z + 1
    If Z Mod 10 = 0 Then
      Me.Caption = Z
    End If
150
  
  
  rDBF.MoveNext
  DoEvents
  
Loop

rSQL.Close
rDBF.Close

Me.Caption = (GetTickCount() - A) / 1000




Exit Sub


WRITEERROR:


' MHNYMA:
HandleError "bohu7:neestimes"
Resume Next


'List2.Visible = True

'List2.AddItem Err.Description
Resume Next


End Sub

Private Sub SendKiniseis_Click()

FileCopy "c:\lageuro\send\eggtim0.dbf", "c:\lageuro\send\eggtim.dbf"




' apo SQL SERVER ΣΕ DBF ΑΡΧΕΙΑ ΠΟΥ ΗΔΗ ΥΠΑΡΧΟΥΝ
Dim K As Integer
Dim R As New ADODB.Recordset
Dim mkod As Integer
update_DBF_from_SQL "EGGTIM WHERE HME>='" + Format(APO, "MM/DD/YYYY") + "' AND HME<'" + Format(DateAdd("D", 1, EOS), "MM/DD/YYYY") + "'", "EGGTIM"

Gdb.Execute "UPDATE PARAMETROI SET TIMH='" + Format(APO, "DD/MM/YYYY") + "' WHERE FORMA='BOHU7' AND VAR='F_APO'"
Gdb.Execute "UPDATE PARAMETROI SET TIMH='" + Format(EOS, "DD/MM/YYYY") + "' WHERE FORMA='BOHU7' AND VAR='F_EOS'"




End Sub

Private Sub Text1_LostFocus()
   If Len(Text1.Text) < 13 Then Exit Sub
   
   
   fR.Open "SELECT KOD,PONTOI FROM EID WHERE KOD='" + Text1.Text + "'", Gdb, adOpenDynamic, adLockOptimistic
   If fR.EOF Then
       MsgBox "ΔΕΝ ΥΠΑΡΧΕΙ Η ΚΑΡΤΑ"
       fR.Close
       
       Exit Sub
   End If
   
On Error Resume Next
Label1.Caption = fR(1)
   fR.Close
   
End Sub


Private Sub Text2_LostFocus()
Dim A As String
If Len(Text1.Text) <> 13 Then
    MsgBox "λάθος κωδικός"
    Exit Sub
End If


If Val(Text2.Text) = 0 Then
    Exit Sub
End If



A = Str(Val(Text2.Text))


Gdb.Execute "update EID set PONTOI=PONTOI+" + A + " where left(KOD,13)='" + Text1.Text + "'", ff
If ff < 1 Then
   MsgBox "δεν ενημερώθηκε"
   Exit Sub
   
End If

List1.AddItem Text1.Text + "--" + Text2.Text

Text2.Text = ""


End Sub
Private Sub update_sql_from_dbf(arxeio As String)
Dim DBF As Database
Dim sql As New ADODB.Connection
Dim rDBF As Recordset
Dim rSQL As New ADODB.Recordset
Dim conDBF As String
Dim conSQL As String
Dim K As Long
Dim Fname As String

G2 = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=EMP4;Data Source=PLATO"
conDBF = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dBASE Files;Initial Catalog=C:\LAGEURO"

sql.Open gConnect

DATA1.DatabaseName = "c:\lageuro\send"
DATA1.RecordSource = arxeio
DATA1.Refresh

Set rDBF = DATA1.Recordset

Label1.Caption = arxeio + " " + Format(rDBF.RecordCount, "######")


rSQL.Open "SELECT *FROM " + arxeio, sql, adOpenDynamic, adLockOptimistic
rDBF.MoveFirst

Dim Z
Z = 0
On Error GoTo WRITEERROR ' Resume Next

Do While Not rDBF.EOF
  
  rSQL.AddNew
  
  For K = 0 To rDBF.FIELDS.Count - 1
      Fname = rDBF.FIELDS(K).Name ' p.x. FNAME=epo    R(0).NAME
      
      If IsNull(rDBF(K)) Then
      Else
         rSQL(Fname) = rDBF(K) ' rsqk("epo")=r(0)
      End If
     ' ΑΥΤΑ ΠΟΥ ΕΙΝΑΙ ΝULL ΤΑ ΓΕΜΙΖΕΙ
     If IsNull(rDBF(K)) Then
        If rDBF(K).Type = dbText Then '10=STRING
            rSQL(Fname) = " "
        ElseIf rDBF(K).Type = dbDate Then
            rSQL(Fname) = CDate("01/01/1990")
        ElseIf rDBF(K).Type = dbBoolean Then
            rSQL(Fname) = False
        ElseIf rDBF(K).Type = dbDouble Then
            rSQL(Fname) = 0
        Else
             conSQL = ""
        End If
     End If
  Next
  
  ' METΑΤΡΕΠΕΙ ΤΟ 'POSO' ΣΕ 'XRE' KAI 'PIS'
  If arxeio = "EGGTIM" Then
     rSQL("PIS") = 0: rSQL("XRE") = 0
     If Not IsNull(rSQL("POSO")) Then
        If InStr(F_xre_APOU + F_axre_APOU, left(rSQL("ATIM"), 1)) > 0 Then
            rSQL("XRE") = rSQL("POSO")
        End If
        If InStr(F_pis_APOU + F_apis_APOU, left(rSQL("ATIM"), 1)) > 0 Then
            rSQL("PIS") = rSQL("POSO")
        End If
     End If
  End If
  
  rSQL.Update
    Z = Z + 1
    If Z Mod 10 = 0 Then
      Me.Caption = Z
    End If
  rDBF.MoveNext
  DoEvents
  
Loop

rSQL.Close
rDBF.Close
Exit Sub


WRITEERROR:
Resume Next
End Sub
