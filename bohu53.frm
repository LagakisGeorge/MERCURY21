VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form bohu53 
   BackColor       =   &H00FF0000&
   Caption         =   "IMPORT ΠΕΛΑΤΩΝ / ΠΡΟΜΗΘΕΥΤΩΝ"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   13635
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   7725
      TabIndex        =   16
      Top             =   4335
      Width           =   4680
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Δημιουργία αρχείου με τα πεδία"
      Height          =   480
      Left            =   8520
      TabIndex        =   12
      Top             =   3675
      Width           =   1845
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Εξοδος"
      Height          =   555
      Left            =   7680
      TabIndex        =   11
      Top             =   6840
      Width           =   1470
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4320
      TabIndex        =   10
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "Πρόσθεση νέων εγγραφών σε Κύριο αρχείο"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   150
      TabIndex        =   8
      Top             =   1005
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "Ενημέρωση τιμών,περιγραφών από Excel"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   150
      TabIndex        =   7
      Top             =   645
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3780
      TabIndex        =   5
      Text            =   "3"
      Top             =   210
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   9960
      TabIndex        =   4
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Text            =   " "
      Top             =   480
      Width           =   4335
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   9600
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Καταχώρηση"
      Enabled         =   0   'False
      Height          =   375
      Left            =   135
      TabIndex        =   1
      Top             =   2550
      Width           =   5925
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4590
      Left            =   150
      TabIndex        =   0
      Top             =   3900
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8096
      _Version        =   393217
      TextRTF         =   $"bohu53.frx":0000
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Height          =   450
      Left            =   150
      TabIndex        =   15
      Top             =   3105
      Width           =   2835
   End
   Begin VB.Label Label5 
      BackColor       =   &H000080FF&
      Height          =   450
      Left            =   3225
      TabIndex        =   14
      Top             =   3105
      Width           =   2835
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Πίνακας με πεδία"
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
      Height          =   315
      Left            =   150
      TabIndex        =   13
      Top             =   3645
      Width           =   1965
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "Αν δεν αναφέρεται άλλο βάζω προεπιλεγμένη κατηγορία ΦΠΑ  την :"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   150
      TabIndex        =   9
      Top             =   1365
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   "Εκκίνηση από την σειρά του Excel :"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   150
      TabIndex        =   6
      Top             =   165
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Αρχείο Excel που περιέχει τον τιμοκατάλογο"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5640
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "bohu53"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xl As New excel.Application
Dim xlsheet As excel.Worksheet
Dim xlsheet3 As excel.Worksheet
Dim xlwbook As excel.workbook






Private Sub Command1_Click()

Dim R As New ADODB.Recordset
Dim R0 As New ADODB.Recordset
Dim mok As Integer

Dim PIN(30, 2)

Dim COUNTER As Integer

Dim KOD_OK  As Boolean ' ΕΧΕΙ ΔΗΛΩΣΕΙ ΤΟΝ ΚΩΔΙΚΟ
Dim BARCODE_OK  As Boolean ' ΕΧΕΙ ΔΗΛΩΣΕΙ TA BARCODES
Dim sql As String

KOD_OK = False
BARCODE_OK = False
Dim k As Long

COUNTER = 1  ' ARIUMOS PEDION POY THA METAFERTHOYN
Dim KOD_COLUMN  As Integer
Dim BARCODE_COLUMN  As Integer

R0.Open "SELECT TOP 1 * FROM PEL ", Gdb, adOpenDynamic, adLockOptimistic











' FORTONO TON PINAKA 20 ME TIS ANTISTOIXISEIS TON PEDION
' ΚΑΙ ΕΛΕΓΧΩ ΑΝ ΕΧΕΙ ΔΗΛΩΘΕΙ Ο ΚΩΔΙΚΟΣ
R.Open "SELECT *FROM PINAKES WHERE TYPOS=21 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
Do While Not R.EOF
  If R("typos") = 21 Then
     PIN(COUNTER, 1) = R("PERIGRAFH") 'ΠΕΔΙΟ ΤΟΥ ΚΥΡΙΟΥ ΑΡΧΕΙΟΥ (SQLSERVER)
     PIN(COUNTER, 2) = R("TIMH") ' ΑΡΙΘΜΟΣ ΣΤΗΛΗΣ ΤΟΥ EXCEL
     
     If UCase(R("PERIGRAFH")) = "KOD" Then
        KOD_OK = True
        KOD_COLUMN = COUNTER
     End If
     
     
     ' On Error GoTo LATOSPEDIO
     mok = 0
     For k = 0 To R0.FIELDS.Count - 1
          If R0(k).Name = R("PERIGRAFH") Then
            mok = 1
            Exit For
          End If
     Next
          
     If mok = 0 Then
         MsgBox "Λάθος στην δήλωση των πεδίων." + Chr(13) + "Δεν υπάρχει στο αρχείο ειδών το πεδίο " + R("PERIGRAFH")
         Exit Sub
     End If
     
     On Error GoTo 0
     
     COUNTER = COUNTER + 1
  
  Else
     Exit Do
  End If
  R.MoveNext
Loop
' kathgories

R.Close

If KOD_OK = False Then
   MsgBox "δεν εχετε δηλώσει τον κωδικό του είδους" + Chr(13) + " Σταματά η ενημέρωση"
   Exit Sub
End If
















Dim ko As String
Dim mNew As Long, mUpd As Long
mNew = 0
mUpd = 0


        Label4.Caption = "Nέες εγγραφές 0"
        Label5.Caption = "Ενημέρωση εγγραφές 0"




Dim mRow As Long
' data1.Recordset.MoveFirst
mRow = Val(Text2.Text)    '  data1.Recordset.Move
 On Error GoTo error_name
        
Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF
     If IsNull(xlsheet.cells(mRow, KOD_COLUMN)) Then
         Exit Do
     End If
     
     If IsEmpty(xlsheet.cells(mRow, KOD_COLUMN)) Then
         Exit Do
     End If
     
     
     ko = Trim(xlsheet.cells(mRow, KOD_COLUMN)) '  ' ΚΩΔΙΚΟς ΤΟΥ ΕΙΔΟΥΣ  EXCEL
          
     
     R.Open "SELECT * FROM PEL WHERE EIDOS='e' and KOD='" + ko + "'", Gdb, adOpenDynamic, adLockOptimistic
     If R.EOF Then 'DEN BRHKA TON KODIKO
        R.AddNew
        R("KOD") = ko
        
        R("eidos") = "e"
        mNew = mNew + 1
        Label4.Caption = "Nέες εγγραφές " + Format(mNew, "######")
     Else  'BRHKA TON KODIKO
        mUpd = mUpd + 1
        Label5.Caption = "Ενημέρωση εγγραφές " + Format(mUpd, "######")
     End If
        
       
        
        'ΕΝΗΜΕΡΩΝΩ ΤΑ ΥΠΟΛΟΙΠΑ ΠΕΔΙΑ ΑΠΟ ΤΟ EXCEL
        For k = 2 To COUNTER - 1
          If R(PIN(k, 1)).Type = 202 Then ' KEIMENO
              R(PIN(k, 1)) = Left(xlsheet.cells(mRow, 2), R("ONO").DefinedSize)
          Else 'ARITMOS
              R(PIN(k, 1)) = xlsheet.cells(mRow, k)
          End If
         
        Next
        R.Update
        
        
        R.Close
        
        
        
        
        DoEvents
        Me.Caption = ko
        mRow = mRow + 1 'data1.Recordset.MoveNext
Loop

xl.Quit
Set xlwbook = Nothing
Set xl = Nothing




MsgBox "τέλος ενημέρωσης"


Exit Sub

error_name:
MsgBox "λάθος στην σειρά " + Format(mRow, "#######")
Resume Next
End Sub

Private Sub Command2_Click()
  
If Len(Trim(Text1.Text)) = 0 Then
  Cd1.ShowOpen
  Text1.Text = Cd1.FileName
Else
  If Len(Dir(LTrim(Text1.Text), vbNormal)) < 2 Then
     MsgBox "δεν υπάρχει το αρχείο " + Text1.Text
     Exit Sub
  End If
End If

On Error GoTo open_error
Me.MousePointer = vbHourglass
Set xlwbook = xl.Workbooks.Open(Text1.Text)
Set xlsheet = xlwbook.Sheets.Item(1)

On Error GoTo 0
Me.MousePointer = vbNormal
Command1.Enabled = True

Exit Sub


open_error:
MsgBox "λάθος στo  αρχείο " + Text1.Text
Me.MousePointer = vbNormal

Exit Sub

  
  
  
  
  
  
  
  
  
  
  
  
 ' data1.RecordSource = xlwbook.Sheets.Item(1).Name '   "Φύλλο1$"
  
' On Error Resume Next
'data1.Refresh

  
  
 ' data1.Recordset.Move 6
 ' Me.Caption = data1.Recordset(2)
End Sub

Private Sub Command3_Click()
  Unload Me
  
End Sub

Private Sub Command4_Click()

Dim R As New ADODB.Recordset
' On Error GoTo CREATEFIELDS
R.Open "select count(*) FROM MFIELDS", Gdb, adOpenForwardOnly, adLockReadOnly
If R(0) > 10 Then
   R.Close
   R.Open "select RTRIM(TITLOS)+'  : '+ONOMA FROM MFIELDS WHERE MFILE='EID'", Gdb, adOpenForwardOnly, adLockReadOnly
   Do While Not R.EOF
      List1.AddItem R(0)
      R.MoveNext
   Loop
   Exit Sub
End If




Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','KOD','ΚΩΔΙΚΟΣ ΕΙΔΟΥΣ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','ONO','ΠΕΡΙΓΡΑΦΗ ΕΙΔΟΥΣ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','ERG','ΒΟΗΘ.ΚΩΔΙΚΟΣ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','BARCODE','BARCODE')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','XTI','ΤΙΜΗ ΑΓΟΡΑΣ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','LTI','ΧΟΝΔΡΙΚΗ ΤΙΜΗ ΕΙΔΟΥΣ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','LTI5','ΛΙΑΝΙΚΗ ΤΙΜΗ ΕΙΔΟΥΣ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','AEG','OIKOΓΕΝΕΙΑ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','KODLOG','ΥΠΟΟΙΚΟΓΕΝΕΙΑ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','FPA','ΦΠΑ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','MON','MONAΔΑ ΜΕΤΡΗΣΗΣ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','KATHGORIA','ΚΑΤΗΓΟΡΙΑ ΕΙΔΟΥΣ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','KODLOG','ΥΠΟΟΙΚΟΓΕΝΕΙΑ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','KODSYNOD','ΚΩΔ.ΛΟΓΙΣΤΙΚΗΣ ΠΩΛΗΣΕΩΝ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','KODLOGAG','ΚΩΔ.ΛΟΓΙΣΤΙΚΗΣ ΑΓΟΡΩΝ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','PROM','ΠΡΟΜΗΘΕΥΤΗΣ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','SPA','ΟΡΙΟ ΑΝΑΠΑΡΑΓΓΕΛΙΑΣ(ΣΤΑΘΜΗ)')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','PROMHU','BONUS ΠΟΝΤΟΙ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','EPIUYP','ΕΠΙΘ.ΥΠΟΛΟΙΠΟ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','KODERG','ΚΩΔ.ΠΡΟΜΗΘΕΥΤΗ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','SYSKMAX','ΚΑΝΟΝΙΚΗ ΣΥΣΚΕΥΑΣΙΑ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','SYSKMIN','ΥΠΟΣΥΣΚΕΥΑΣΙΑ(ΜΙΚΡΟΤΕΡΗ ΣΥΣΚ.)')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','POS_KERD','MARKUP ΧΟΝΔΡΙΚΗΣ')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','POS_KERD2','MARKUP ΛΙΑΝΙΚΗΣ')"

Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','MEMO','MEMO(ΣΗΜΕΙΩΣΗ)')"

Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','CHAR1','" + FindParametroi("APOT2", "F_lab1", "Κειμενο 1", "Ετικέτα κειμ.1") + "')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','CHAR2','" + FindParametroi("APOT2", "F_lab2", "Κειμενο 2", "Ετικέτα κειμ.2") + "')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','CHAR3','" + FindParametroi("APOT2", "F_lab3", "Κειμενο 3", "Ετικέτα κειμ.3") + "')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','CHAR4','" + FindParametroi("APOT2", "F_lab4", "Κειμενο 4", "Ετικέτα κειμ.4") + "')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','CHAR5','" + FindParametroi("APOT2", "F_lab5", "Κειμενο 5", "Ετικέτα κειμ.5") + "')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','CHAR6','" + FindParametroi("APOT2", "F_lab6", "Κειμενο 6", "Ετικέτα κειμ.6") + "')"

Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','HM1','" + FindParametroi("APOT2", "F_lhm1", "Ημερ1", "Ετικέτα ημερ.1") + "')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','HM2','" + FindParametroi("APOT2", "F_lhm2", "Ημερ1", "Ετικέτα ημερ.1") + "')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','HM3','" + FindParametroi("APOT2", "F_lhm3", "Ημερ1", "Ετικέτα ημερ.1") + "')"

Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','NUM1','" + FindParametroi("APOT2", "F_num1", "Αριθμ1", "Ετικέτα αριθμ.1") + "')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','NUM2','" + FindParametroi("APOT2", "F_num2", "Aριθμ2", "Ετικέτα αριθμ.2") + "')"
Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','NUM3','" + FindParametroi("APOT2", "F_num3", "Aριθμ3", "Ετικέτα αριθμ.3") + "')"

       

Exit Sub
CREATEFIELDS:
'---------------------------------------------------------------
'ΔΗΜΙΟΥΡΓΙΑ ΑΡΧΕΙΟΥ ΠΕΔΙΩΝ
Gdb.Execute "create table MFIELDS (MFILE   char(20)," _
& "TITLOS CHAR(20)," _
& "ONOMA  char(30)," _
& "TYPOS  char(7)  )"


End Sub

Private Sub Form_Load()
Dim R As New ADODB.Recordset

   R.Open "SELECT *FROM PINAKES WHERE TYPOS=1 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
   

'FPA
Do While Not R.EOF
   If R("typos") = 1 Then
      Combo2.AddItem Str(R("AYJON")) + " -> " + Str(R("TIMH"))
   End If
   R.MoveNext
Loop
 R.Close

Combo2.Text = Combo2.List(1)



End Sub

Private Sub Label6_Click()
  Dim f As New DataEntry2, DD
  Dim R As New ADODB.Recordset
'pinakas pedion
    UPDATE_PINAKES3f 20, Label6 ', Combo1
End Sub
Sub UPDATE_PINAKES3f(MTYPOS As Integer, OO As Object)
'mtypos  αριθμος πίνακα
'οο: to object που θέλω για το κεντράρισμα της οθόνης DataEntry
'mcombo : το combo που θα γεμίσει με τις τιμές
  Dim f As New DataEntry2, DD
  Dim R As New ADODB.Recordset
'ΜΟΝΑΔΕΣ ΜΕΤΡΗΣΗΣ
  'sql οπου θα γίνεται το update ή το addnew
 '  f.SHOW
  Load f
  
  Gdb.Execute "update PINAKES SET AYJON =0 WHERE AYJON IS NULL AND TYPOS=" + Str(MTYPOS)
  f.ParamGrid.TextMatrix(0, 1) = "SELECT * FROM PINAKES WHERE TYPOS=" + Str(MTYPOS) '    sql
'τα πεδία που θα μπαίνουν sto ParamGrid και στα textboxes
f.ParamGrid.TextMatrix(0, 2) = "SELECT AYJON,PERIGRAFH,TIMH " _
& " FROM PINAKES WHERE TYPOS=" + Str(MTYPOS) + " ORDER BY AYJON " 'ParamGrid   sql2

'με ποιά εντολή θα κάνει το delete (το where ειναι το kod=...
f.ParamGrid.TextMatrix(0, 3) = "delete from PINAKES where TYPOS=" + Str(MTYPOS) + " AND " 'ParamGrid              DelSQL

'ποιό update query θα δημιουργείται (προαιρετικά )
'αν δεν θέλω update τότε βάζω f.ParamGrid.TextMatrix(0, 4) = ""
'                    f.ParamGrid.TextMatrix(0, 4) = "  ""UPDATE ARITMISI2 SET ARITMISI= "" + t12 +"" where ID=""+t11 "   'UpdSQL
 DD = " UPDATE PINAKES SET TYPOS=" + Str(MTYPOS) + " where TYPOS IS NULL AND AYJON="" + t1   "
' AddNewScript
  DD = "Sub Main()" & vbCrLf & _
         " DIM mCon  " & vbCrLf & _
         " Set mCon=CreateObject(""ADODB.Connection"")" & vbCrLf & _
         " mCon.Open """ + gConnect + " "" " & vbCrLf & _
         " mCon.execute """ + DD + " " & vbCrLf & _
        "Set r=Nothing" & vbCrLf & _
        "Set mCon=Nothing" & vbCrLf & _
        "End Sub"
f.ParamGrid.TextMatrix(0, 4) = DD

f.ParamGrid.TextMatrix(1, 1) = "A/A"
f.ParamGrid.TextMatrix(2, 1) = "ΤΙΤΛΟΣ "
f.ParamGrid.TextMatrix(3, 1) = "ΣΤΗΛΗ "
'ΜΕΓΑΛΩΝΩ ΛΙΓΟ ΤΑ LABELS
f.Labels(0).width = f.Labels(0).width + 1000 '* TextWidth("A")
f.TextBoxes(0).Left = f.TextBoxes(0).Left + 1000 ' * TextWidth("A")

'αν παει η φόρμα να φύγει από την οθόνη τότε να κεντραρεται δεξια με την mdiform1
If OO.Left + f.width > MDIForm1.Left + MDIForm1.width Then
  f.Left = MDIForm1.Left + MDIForm1.width - f.width
Else
  f.Left = OO.Left
End If

f.Top = OO.Top + MDIForm1.Top






f.SHOW 1
'MCOMBO.Clear
' R.Open "SELECT *FROM PINAKES WHERE TYPOS=" + Str(MTYPOS) + " ORDER BY AYJON", Gdb, adOpenDynamic, _
'     adLockOptimistic
'Do While Not R.EOF
'  MCOMBO.AddItem R("PERIGRAFH")
'  R.MoveNext
'Loop
'R.Close
'
'MCOMBO.Refresh

End Sub


