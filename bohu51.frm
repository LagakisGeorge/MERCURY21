VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form bohu51old 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   10560
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Εξοδος"
      Height          =   555
      Left            =   8550
      TabIndex        =   10
      Top             =   6840
      Width           =   1470
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "bohu51.frx":0000
      Left            =   120
      List            =   "bohu51.frx":000A
      TabIndex        =   9
      Top             =   1380
      Width           =   2565
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "Πρόσθεση νέων εγγραφών σε Κύριο αρχείο"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "Ενημέρωση εγγραφών από Excel"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Text            =   "3"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   8790
      Visible         =   0   'False
      Width           =   4215
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
      Height          =   375
      Left            =   135
      TabIndex        =   1
      Top             =   1800
      Width           =   4830
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3660
      Left            =   120
      TabIndex        =   0
      Top             =   3810
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6456
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"bohu51.frx":002C
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Πίνακας 22 (πεδία πελατών)"
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
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Εκκίνηση από την σειρά του Excel :"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Αρχείο Excel που περιέχει τον τιμοκατάλογο"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5640
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "bohu51old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim R As New ADODB.Recordset


Dim PIN(30, 2)

Dim COUNTER As Integer

 
COUNTER = 1  ' ARIUMOS PEDION POY THA METAFERTHOYN
 
' FORTONO TON PINAKA 21 ME TIS ANTISTOIXISEIS TON PEDION
R.Open "SELECT *FROM PINAKES WHERE TYPOS=21 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
Do While Not R.EOF
  If R("typos") = 21 Then
     PIN(COUNTER, 1) = R("PERIGRAFH") 'ΠΕΔΙΟ ΤΟΥ ΚΥΡΙΟΥ ΑΡΧΕΙΟΥ (SQLSERVER)
     PIN(COUNTER, 2) = R("TIMH") ' ΑΡΙΘΜΟΣ ΣΤΗΛΗΣ ΤΟΥ EXCEL
     COUNTER = COUNTER + 1
  Else
     Exit Do
  End If
  R.MoveNext
Loop
' kathgories

R.Close
' Exit Sub

Dim ko As String
Dim BUFF As String

BUFF = Left(Combo1.Text, 1)

Data1.Recordset.MoveFirst
Data1.Recordset.Move Val(Text2.Text) - 2


Dim KNEW
KNEW = 0
Do While Not Data1.Recordset.EOF
     If IsNull(Data1.Recordset(PIN(1, 2) - 1)) Then
         Exit Do
     End If
     On Error GoTo LATOS
     ko = Data1.Recordset(PIN(1, 2) - 1) '  ' ΚΩΔΙΚΟς ΤΟΥ ΕΙΔΟΥΣ  EXCEL
          
     
     R.Open "SELECT * FROM PEL WHERE KOD='" + ko + "' AND EIDOS='" + BUFF + "'", Gdb, adOpenDynamic, adLockOptimistic
     If R.EOF Then 'DEN BRHKA TON KODIKO
        
        KNEW = KNEW + 1
        R.AddNew
        R("KOD") = ko
        R("EIDOS") = BUFF
        
     Else  'BRHKA TON KODIKO
     End If
        
        'ΕΝΗΜΕΡΩΝΩ ΤΑ ΥΠΟΛΟΙΠΑ ΠΕΔΙΑ ΑΠΟ ΤΟ EXCEL
        For K = 2 To COUNTER - 1
          
          If R(PIN(K, 1)).Type = 202 Then ' KEIMENO
              R(PIN(K, 1)) = Left(Data1.Recordset(PIN(K, 2)), R(PIN(K, 1)).DefinedSize)
              
          Else 'ARITMOS
              R(PIN(K, 1)) = Data1.Recordset(PIN(K, 2) - 1)
          End If
          
        Next
        R.Update
        R.Close
        DoEvents
        Me.Caption = ko
        Data1.Recordset.MoveNext
Loop

MsgBox "τέλος ενημέρωσης" + Chr(13) + "Νέες εγγραφές " + Str(KNEW)


Exit Sub


LATOS:
MsgBox "ΛΑΘΟΣ ΣΤΗΝ ΣΕΙΡΑ " + Str(K)





















End Sub

Private Sub Command2_Click()
  Cd1.ShowOpen
  Data1.Connect = "Excel 8.0;"
  Data1.DatabaseName = Cd1.FILENAME
  Data1.RecordSource = "Φύλλο1$"
  
  Data1.Refresh
  Text1.Text = Cd1.FILENAME
  
  
  'Data1.Recordset.Move 6
  'Me.Caption = Data1.Recordset(2)
End Sub

Private Sub Command3_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Dim R As New ADODB.Recordset

  ' R.Open "SELECT *FROM PINAKES WHERE TYPOS=1 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
   



'Combo2.Text = Combo2.List(1)



End Sub

Private Sub Label3_Click()
  Dim f As New DataEntry2, DD
  Dim R As New ADODB.Recordset
'pinakas pedion
    UPDATE_PINAKES3f 21, Label3 ', Combo1
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
