VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form bohu53 
   BackColor       =   &H00FF0000&
   Caption         =   "IMPORT ΕΙΔΩΝ"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   13635
   WindowState     =   2  'Maximized
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   120
      TabIndex        =   26
      Top             =   3600
      Width           =   2895
   End
   Begin VB.CheckBox APESTALMENA2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "Σε πίνακα APESTALMENA2"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6240
      TabIndex        =   25
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox colSxolia 
      Height          =   285
      Left            =   5760
      TabIndex        =   22
      Text            =   "30"
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Εισαγωγή κινήσεων απογραφής από EXCEL"
      Height          =   252
      Left            =   8160
      TabIndex        =   20
      ToolTipText     =   "ΚΩΔΙΚΟΣ,ΠΟΣΟ ΧΡΕΩΣΗΣ(ΜΕ - ΠΙΣΤΩΣΗ),ΑΙΤΙΟΛΟΓΙΑ"
      Top             =   1536
      Width           =   3570
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "bohu52.frx":0000
      Left            =   8160
      List            =   "bohu52.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1824
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.CheckBox eiddok 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "Σε πίνακα EIDDOK(δοκιμή)"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   3552
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ΑΠΟ CSV"
      Height          =   375
      Left            =   3516
      TabIndex        =   17
      Top             =   2496
      Width           =   2895
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3840
      TabIndex        =   16
      Text            =   "Combo3"
      Top             =   8640
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4470
      Left            =   7725
      TabIndex        =   15
      Top             =   3900
      Width           =   4680
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Εξοδος"
      Height          =   435
      Left            =   7725
      TabIndex        =   11
      Top             =   8520
      Width           =   4695
   End
   Begin VB.ComboBox Combo2 
      Height          =   288
      Left            =   4956
      TabIndex        =   10
      Top             =   1368
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "Πρόσθεση νέων εγγραφών σε Κύριο αρχείο"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   159
      TabIndex        =   8
      Top             =   1056
      Value           =   1  'Checked
      Width           =   6252
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "Ενημέρωση τιμών,περιγραφών από Excel"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   207
      TabIndex        =   7
      Top             =   696
      Value           =   1  'Checked
      Width           =   6204
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5316
      TabIndex        =   5
      Text            =   "3"
      Top             =   264
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   6156
      TabIndex        =   4
      Top             =   1788
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3744
      TabIndex        =   2
      Text            =   " "
      Top             =   1776
      Width           =   2460
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
      Top             =   2520
      Width           =   2925
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3750
      Left            =   150
      TabIndex        =   0
      Top             =   4740
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6615
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"bohu52.frx":0024
   End
   Begin MSComCtl2.DTPicker APOT 
      Height          =   330
      Left            =   8160
      TabIndex        =   21
      Top             =   2205
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   582
      _Version        =   393216
      Format          =   171048961
      CurrentDate     =   39565
   End
   Begin VB.Shape Shape1 
      Height          =   2895
      Left            =   7680
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label eggtim 
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
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   8160
      TabIndex        =   24
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF0000&
      Caption         =   "Στήλη με αποτέλεσμα εισαγωγή 0=όχι σχόλια"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   2160
      Width           =   4140
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Height          =   450
      Left            =   150
      TabIndex        =   14
      Top             =   3105
      Width           =   2835
   End
   Begin VB.Label Label5 
      BackColor       =   &H000080FF&
      Height          =   456
      Left            =   3579
      TabIndex        =   13
      Top             =   3072
      Width           =   2832
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Κατασκευή Πίνακα με πεδία"
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
      Left            =   3120
      TabIndex        =   12
      Top             =   3960
      Width           =   2325
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
      Height          =   372
      Left            =   192
      TabIndex        =   6
      Top             =   288
      Width           =   4956
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Αρχείο Excel που περιέχει τον τιμοκατάλογο"
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   96
      TabIndex        =   3
      Top             =   1824
      Width           =   3420
   End
End
Attribute VB_Name = "bohu53"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim xl       As New excel.Application

Dim xlsheet  As excel.Worksheet

Dim xlsheet3 As excel.Worksheet

Dim xlwbook  As excel.workbook

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

Private Sub APESTALMENA2_Click()
  List1.Clear
 FILL_FIELDS "APESTALMENA2"
End Sub

'Dim xl As New excel.Application
'Dim xlsheet3 As excel.Worksheet
'Dim xlwbook As excel.workbook
'

'Set xlwbook = xl.Workbooks.Open(Text1.Text)
'Set xlsheet = xlwbook.Sheets.Item(1)
'ko = Trim(xlsheet.cells(mRow, KOD_COLUMN)) '  ' ΚΩΔΙΚΟς ΤΟΥ ΕΙΔΟΥΣ  EXCEL

'xl.Quit
'Set xlwbook = Nothing
'Set xl = Nothing

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

        Dim R   As New ADODB.Recordset

        Dim r0  As New ADODB.Recordset

        Dim mok As Integer

        Dim PIN(30, 2)

        Dim COUNTER    As Integer

        Dim KOD_OK     As Boolean    ' ΕΧΕΙ ΔΗΛΩΣΕΙ ΤΟΝ ΚΩΔΙΚΟ

        Dim BARCODE_OK As Boolean    ' ΕΧΕΙ ΔΗΛΩΣΕΙ TA BARCODES

        Dim sql        As String

        Dim NON_STOP   As Boolean

100     NON_STOP = False
110     KOD_OK = False
120     BARCODE_OK = False

        Dim k As Long

130     COUNTER = 1  ' ARIUMOS PEDION POY THA METAFERTHOYN

        Dim KOD_COLUMN     As Integer

        Dim BARCODE_COLUMN As Integer

        Dim mEID           As String

        mEID = "EID"

        If eiddok.Value = vbChecked Then
            mEID = "EIDDOK"
        End If
        
        If APESTALMENA2.Value = vbChecked Then
            mEID = "APESTALMENA2"
        End If
        
        
        
        Dim c As String
   c = List2.List(List2.ListIndex)   'ΛΙΣΤΑ  1; 2001
   Dim cnn As String
   cnn = Split(c, ";")(1) '2001
        
        
        
        
        
        
        

140     r0.Open "SELECT TOP 1 * FROM " + mEID, Gdb, adOpenDynamic, adLockOptimistic

        ' FORTONO TON PINAKA 20 ME TIS ANTISTOIXISEIS TON PEDION
        ' ΚΑΙ ΕΛΕΓΧΩ ΑΝ ΕΧΕΙ ΔΗΛΩΘΕΙ Ο ΚΩΔΙΚΟΣ
150     R.Open "SELECT *FROM PINAKES WHERE TYPOS=" + cnn + " ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

160     Do While Not R.EOF

170         If R("typos") = Val(cnn) Then
180             PIN(COUNTER, 1) = R("PERIGRAFH")    'ΠΕΔΙΟ ΤΟΥ ΚΥΡΙΟΥ ΑΡΧΕΙΟΥ (SQLSERVER)
190             PIN(COUNTER, 2) = R("TIMH")    ' ΑΡΙΘΜΟΣ ΣΤΗΛΗΣ ΤΟΥ EXCEL

200             If UCase(R("PERIGRAFH")) = "KOD" Then
210                 KOD_OK = True
220                 KOD_COLUMN = R("TIMH") '' COUNTER
                End If

230             If UCase(R("PERIGRAFH")) = "BARCODE" Then
240                 BARCODE_OK = True
250                 BARCODE_COLUMN = PIN(COUNTER, 2)    ' COUNTER
                End If

                ' On Error GoTo LATOSPEDIO
260             mok = 0

270             If UCase(R("PERIGRAFH")) = "BARCODE" Then
280                 mok = 1
                Else

290                 For k = 0 To r0.FIELDS.Count - 1     'ΕΛΕΓΧΟΣ ΑΝ ΕΙΝΑΙ ΥΠΑΡΚΤΟ ΤΟ ΠΕΔΙΟ

300                     If r0(k).Name = R("PERIGRAFH") Then
310                         mok = 1

                            Exit For

                        End If

                    Next

                End If

320             If mok = 0 Then
330                 MsgBox "Λάθος στην δήλωση των πεδίων." + Chr(13) + "Δεν υπάρχει στο αρχείο ειδών το πεδίο " + R("PERIGRAFH")

                    Exit Sub

                End If

                On Error GoTo Command1_Click_Err

340             COUNTER = COUNTER + 1
            Else

                Exit Do

            End If

350         R.MoveNext
        Loop

        ' kathgories

360     R.Close
 
 
 
 If APESTALMENA2.Value = vbChecked Then
            KOD_OK = True
            KOD_COLUMN = 2
 Else
 
 
370     If KOD_OK = False Then
380         MsgBox "δεν εχετε δηλώσει τον κωδικό του είδους" + Chr(13) + " Σταματά η ενημέρωση"

            Exit Sub

        End If

End If




390     Set xlwbook = xl.Workbooks.Open(Text1.Text)
400     Set xlsheet = xlwbook.Sheets.Item(1)

        Dim ko   As String

        Dim mNew As Long, mUpd As Long

410     mNew = 0
420     mUpd = 0

430     Label4.Caption = "Nέες εγγραφές 0"
440     Label5.Caption = "Ενημέρωση εγγραφές 0"

        Dim mRow As Long

        ' data1.Recordset.MoveFirst
450     mRow = Val(Text2.Text)    '  data1.Recordset.Move

        On Error GoTo error_name

460     Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF

470         If IsNull(xlsheet.cells(mRow, KOD_COLUMN)) Then

                Exit Do

            End If

480         If IsEmpty(xlsheet.cells(mRow, KOD_COLUMN)) Then

                Exit Do

            End If

490         ko = Trim(xlsheet.cells(mRow, KOD_COLUMN))    '  ' ΚΩΔΙΚΟς ΤΟΥ ΕΙΔΟΥΣ  EXCEL
If APESTALMENA2.Value = vbChecked Then
   R.Open "SELECT * FROM " + mEID + " WHERE MARK='" + ko + "'", Gdb, adOpenDynamic, adLockOptimistic

Else
500         R.Open "SELECT * FROM " + mEID + " WHERE KOD='" + ko + "'", Gdb, adOpenDynamic, adLockOptimistic
End If



510         If R.EOF Then    'DEN BRHKA TON KODIKO
520             R.AddNew

                 If APESTALMENA2.Value = vbChecked Then
                 Else

530                 R("KOD") = Left(ko, 14)
540                 R("fpa") = Val(Left(Combo2.Text, 2))
                 End If
                 
550             mNew = mNew + 1
560             Label4.Caption = "Nέες εγγραφές " + Format(mNew, "######")
                If Val(colSxolia.Text) > 0 Then xlsheet.cells(mRow, Val(colSxolia.Text)) = "Νέα"
            Else  'BRHKA TON KODIKO
570             mUpd = mUpd + 1
580             Label5.Caption = "Ενημέρωση εγγραφές " + Format(mUpd, "######")
                If Val(colSxolia.Text) > 0 Then xlsheet.cells(mRow, Val(colSxolia.Text)) = "Ενημέρωση"
                            
            End If

            'ΕΝΗΜΕΡΩΝΩ ΤΑ ΥΠΟΛΟΙΠΑ ΠΕΔΙΑ ΑΠΟ ΤΟ EXCEL
590         For k = 1 To COUNTER - 1

600             If PIN(k, 1) = "BARCODE" Then
                Else

610                 If R(PIN(k, 1)).Type = 202 Or R(PIN(k, 1)).Type = 129 Then    ' KEIMENO
620                     R(PIN(k, 1)) = Left(xlsheet.cells(mRow, PIN(k, 2)), R(PIN(k, 1)).DefinedSize)
                    Else    'ARITMOS
630                     R(PIN(k, 1)) = xlsheet.cells(mRow, PIN(k, 2))
                    End If
                End If

            Next

640         R.Update

650         R.Close

660         If KOD_OK And BARCODE_OK And mEID = "EID" Then

670             SQL2 = xlsheet.cells(mRow, BARCODE_COLUMN)
680             sql = "INSERT INTO BARCODES (KOD,ERG) VALUES ('" + ko + "','" + LTrim(SQL2) + "')"

690             If Len(SQL2) > 0 Then
700                 Gdb.Execute sql
                End If

            End If

710         DoEvents
720         Me.Caption = ko
730         mRow = mRow + 1    'data1.Recordset.MoveNext
        Loop

740     xl.Quit
750     Set xlwbook = Nothing
760     Set xl = Nothing

770     MsgBox "τέλος ενημέρωσης"

        Exit Sub

error_name:

780     If NON_STOP = False Then
790         MsgBox "λάθος στην σειρά " + Format(mRow, "#######")

800         If MsgBox("ΤΕΡΜΑΤΙΣΜΟΣ ;", vbYesNo) = vbYes Then

                Exit Sub

            End If

810         If MsgBox("ΣΥΝΕΧΕΙΑ ΧΩΡΙΣ ΕΡΩΤΗΣΗ ;", vbYesNo) = vbYes Then
820             NON_STOP = True
            End If
        End If

830     Resume Next

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu53.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu53.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>

100     If Len(Trim(Text1.Text)) = 0 Then
110         cd1.ShowOpen
120         Text1.Text = cd1.FILENAME
        Else

130         If Len(Dir(LTrim(Text1.Text), vbNormal)) < 2 Then
140             MsgBox "δεν υπάρχει το αρχείο " + Text1.Text

                Exit Sub

            End If
        End If

        On Error GoTo open_error

150     Me.MousePointer = vbHourglass
        'Set xlwbook = xl.Workbooks.Open(Text1.Text)
        'Set xlsheet = xlwbook.Sheets.Item(1)

        On Error GoTo Command2_Click_Err

160     Me.MousePointer = vbNormal
170     Command1.Enabled = True

        Exit Sub

open_error:
180     MsgBox "λάθος στo  αρχείο " + Text1.Text
190     Me.MousePointer = vbNormal

        Exit Sub

        ' data1.RecordSource = xlwbook.Sheets.Item(1).Name '   "Φύλλο1$"

        ' On Error Resume Next
        'data1.Refresh

        ' data1.Recordset.Move 6
        ' Me.Caption = data1.Recordset(2)
        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu53.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu53.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>

100     Unload Me

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu53.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu53.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub FILL_FIELDS(TABL As String) 'Command4_Click()

        '<EhHeader>
        On Error GoTo Command4_Click_Err
        
        
        
        
         Dim R As New ADODB.Recordset
        
        
       R.Open "SELECT COLUMN_NAME,DATA_TYPE,CHARACTER_MAXIMUM_LENGTH FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='" + TABL + "' ", Gdb, adOpenForwardOnly, adLockReadOnly

10       If True Then 'R(0) > 10 Then
20        ' R.Close
30         'R.Open "select RTRIM(TITLOS)+'  : '+ONOMA FROM MFIELDS WHERE MFILE='EID'", Gdb, adOpenForwardOnly, adLockReadOnly

40         Do While Not R.EOF
               List1.AddItem Left(R(0) + Space$(30), 20) + ";" + Left(R(1) + Space$(30), 6) + ";" + Right(Space$(5) + str(nNull(R(2))), 3) + ";" '+ R!PEDIO
50             'List1.AddItem R(0) + ";" + R(1) + ";" + str(R(2))
60             R.MoveNext
            Loop

            Exit Sub

          End If

        R.Close
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        

        '</EhHeader>

       ' Dim R As New ADODB.Recordset

        On Error GoTo CREATEFIELDS

100     R.Open "select count(*) FROM MFIELDS WHERE MFILE='EID' ", Gdb, adOpenForwardOnly, adLockReadOnly

110    If R(0) > 10 Then
120         R.Close
'130         R.Open "select RTRIM(TITLOS)+'  : '+ONOMA FROM MFIELDS WHERE MFILE='EID'", Gdb, adOpenForwardOnly, adLockReadOnly
'
'140         Do While Not R.EOF
'150             List1.AddItem R(0)
'160             R.MoveNext
'            Loop

            'Exit Sub
       Else

     

170     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','KOD','ΚΩΔΙΚΟΣ ΕΙΔΟΥΣ')"
180     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','ONO','ΠΕΡΙΓΡΑΦΗ ΕΙΔΟΥΣ')"
190     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','ERG','ΒΟΗΘ.ΚΩΔΙΚΟΣ')"
200     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','BARCODE','BARCODE')"
210     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','XTI','ΤΙΜΗ ΑΓΟΡΑΣ')"
220     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','LTI','ΧΟΝΔΡΙΚΗ ΤΙΜΗ ΕΙΔΟΥΣ')"
230     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','LTI5','ΛΙΑΝΙΚΗ ΤΙΜΗ ΕΙΔΟΥΣ')"
240     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','AEG','OIKOΓΕΝΕΙΑ')"
250     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','KODLOG','ΥΠΟΟΙΚΟΓΕΝΕΙΑ')"
260     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','FPA','ΦΠΑ')"
270     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','MON','MONAΔΑ ΜΕΤΡΗΣΗΣ')"
280     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','KATHGORIA','ΚΑΤΗΓΟΡΙΑ ΕΙΔΟΥΣ')"
290     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','KODLOG','ΥΠΟΟΙΚΟΓΕΝΕΙΑ')"
300     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','KODSYNOD','ΚΩΔ.ΛΟΓΙΣΤΙΚΗΣ ΠΩΛΗΣΕΩΝ')"
310     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','KODLOGAG','ΚΩΔ.ΛΟΓΙΣΤΙΚΗΣ ΑΓΟΡΩΝ')"
320     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','PROM','ΠΡΟΜΗΘΕΥΤΗΣ')"
330     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','SPA','ΟΡΙΟ ΑΝΑΠΑΡΑΓΓΕΛΙΑΣ(ΣΤΑΘΜΗ)')"
340     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','PROMHU','BONUS ΠΟΝΤΟΙ')"
350     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','EPIUYP','ΕΠΙΘ.ΥΠΟΛΟΙΠΟ')"
360     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','KODERG','ΚΩΔ.ΠΡΟΜΗΘΕΥΤΗ')"
370     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','SYSKMAX','ΚΑΝΟΝΙΚΗ ΣΥΣΚΕΥΑΣΙΑ')"
380     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','SYSKMIN','ΥΠΟΣΥΣΚΕΥΑΣΙΑ(ΜΙΚΡΟΤΕΡΗ ΣΥΣΚ.)')"
390     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','POS_KERD','MARKUP ΧΟΝΔΡΙΚΗΣ')"
400     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','POS_KERD2','MARKUP ΛΙΑΝΙΚΗΣ')"

410     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','MEMO','MEMO(ΣΗΜΕΙΩΣΗ)')"

420     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','CHAR1','" + FINDPARAMETROI(1, "APOT2", "F_lab1", "Κειμενο 1", "Ετικέτα κειμ.1") + "')"
430     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','CHAR2','" + FINDPARAMETROI(1, "APOT2", "F_lab2", "Κειμενο 2", "Ετικέτα κειμ.2") + "')"
440     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','CHAR3','" + FINDPARAMETROI(1, "APOT2", "F_lab3", "Κειμενο 3", "Ετικέτα κειμ.3") + "')"
450     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','CHAR4','" + FINDPARAMETROI(1, "APOT2", "F_lab4", "Κειμενο 4", "Ετικέτα κειμ.4") + "')"
460     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','CHAR5','" + FINDPARAMETROI(1, "APOT2", "F_lab5", "Κειμενο 5", "Ετικέτα κειμ.5") + "')"
470     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','CHAR6','" + FINDPARAMETROI(1, "APOT2", "F_lab6", "Κειμενο 6", "Ετικέτα κειμ.6") + "')"

480     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','HM1','" + FINDPARAMETROI(1, "APOT2", "F_lhm1", "Ημερ1", "Ετικέτα ημερ.1") + "')"
490     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','HM2','" + FINDPARAMETROI(1, "APOT2", "F_lhm2", "Ημερ1", "Ετικέτα ημερ.1") + "')"
500     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','HM3','" + FINDPARAMETROI(1, "APOT2", "F_lhm3", "Ημερ1", "Ετικέτα ημερ.1") + "')"

510     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','NUM1','" + FINDPARAMETROI(1, "APOT2", "F_num1", "Αριθμ1", "Ετικέτα αριθμ.1") + "')"
520     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','NUM2','" + FINDPARAMETROI(1, "APOT2", "F_num2", "Aριθμ2", "Ετικέτα αριθμ.2") + "')"
530     Gdb.Execute "INSERT INTO MFIELDS (MFILE,TITLOS,ONOMA) VALUES('EID','NUM3','" + FINDPARAMETROI(1, "APOT2", "F_num3", "Aριθμ3", "Ετικέτα αριθμ.3") + "')"


      End If


Dim Q As String
Q = "SELECT COLUMN_NAME,DATA_TYPE,ISNULL(CHARACTER_MAXIMUM_LENGTH,0) AS LL,ISNULL(M.ONOMA,'') AS PEDIO FROM INFORMATION_SCHEMA.COLUMNS F LEFT JOIN MFIELDS M ON LEFT(M.MFILE,3)='EID' AND RTRIM(F.COLUMN_NAME)=RTRIM(M.TITLOS)   WHERE TABLE_NAME='EID' "
On Error Resume Next


      R.Open Q, Gdb, adOpenForwardOnly, adLockReadOnly
        
       If True Then 'R(0) > 10 Then
        ' R.Close
         'R.Open "select RTRIM(TITLOS)+'  : '+ONOMA FROM MFIELDS WHERE MFILE='EID'", Gdb, adOpenForwardOnly, adLockReadOnly

         Do While Not R.EOF
             List1.AddItem Left(R(0) + Space$(30), 20) + ";" + Left(R(1) + Space$(30), 6) + ";" + Right(Space$(5) + str(R(2)), 3) + ";" + R!PEDIO
             R.MoveNext
         Loop

            Exit Sub

          End If










        Exit Sub

CREATEFIELDS:
        '---------------------------------------------------------------
        'ΔΗΜΙΟΥΡΓΙΑ ΑΡΧΕΙΟΥ ΠΕΔΙΩΝ
540    ' Gdb.Execute "create table MFIELDS (MFILE   char(20)," & "TITLOS CHAR(20)," & "ONOMA  char(30)," & "TYPOS  char(7)  )"
550     MsgBox "Ξαναπατήστε το πλήκτρο 'Δημιουργία εγγραφών' "

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu53.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu53.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command5_Click()

        'ΛΗΨΗ ΑΠΟ EXCEL ΚΙΝΗΣΕΩΝ ΑΠΟΘΗΚΗΣ
        ' ============================
        '<EhHeader>
        On Error GoTo Command5_Click_Err

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

        Dim DB             As DAO.Database

        Dim arxeio         As String, arxeio2 As String

        Dim SynStoYpoloipo As Integer

        Dim M_ATIM

100     M_ATIM = "λ" + Format(time(), "hhmmss")

        'SynStoYpoloipo = MsgBox("Να προστεθεί στο ήδη υπάρχον υπόλοιπο;" + Chr(13) + "Σε περίπτωση που απαντήσετε όχι η ποσότητα που απογράψατε θα είναι και το υπόλοιπο", vbYesNoCancel)
        'If SynStoYpoloipo = vbCancel Then
        '    MsgBox "Η εργασία ακυρώθηκε"
        '    Exit Sub
        'End If

        Dim BARC1STILI As Integer

        ' BARC1STILI = MsgBox("Η πρώτη στήλη είναι το BARCODE; " + Chr(13) + " Oxι σημαίνει ότι είναι ο κωδικός", vbYesNo)

        '====================================================

        'On Error GoTo open_error
110     Me.MousePointer = vbHourglass

120     If Len(Dir(Text1.Text, vbNormal)) > 0 Then
            ' OK
        Else
130         MsgBox " ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ ΑΡΧΕΙΟ " + Text1.Text

            Exit Sub

        End If

140     Set xlwbook = xl.Workbooks.Open(Text1.Text)
150     Set xlsheet = xlwbook.Sheets.Item(1)

        'Set db = OpenDatabase(Dir1.Path, False, False, "dBase III;")

        'Set rDBF = db.OpenRecordset("SELECT * FROM " + File1.TEXT1)

Dim n21 As Long




Gdb.Execute "insert into TIM (EIDOS,ATIM,KLEIDI,KPE,HME) VALUES ('r','λ000001','λ000001','0000','" + Format(apot.Value, "mm/dd/yyyy") + "')", n21
If n21 = 0 Then
      MsgBox "αποτυχια δημιουργιας ΤΙΜ"
      Exit Sub
End If

Dim r21 As New ADODB.Recordset
r21.Open "select max(ID_NUM) FROM TIM", Gdb, adOpenDynamic, adLockOptimistic
n21 = r21(0)
r21.Close


Gdb.Execute "UPDATE TIM SET B_N1=1,AJI=0,AJ1=0,AJ2=0,AJ3=0,AJ4=0,AJ5=0,AJ6=0,AJ7=0,FPA1=0,FPA2=0 WHERE ID_NUM=" + str(n21)

Dim PIN(20, 2) As String
Dim COUNTER As Integer
Dim R As New ADODB.Recordset

        ' FORTONO TON PINAKA 100 ME TIS ANTISTOIXISEIS TON PEDION
        ' ΚΑΙ ΕΛΕΓΧΩ ΑΝ ΕΧΕΙ ΔΗΛΩΘΕΙ Ο ΚΩΔΙΚΟΣ
     R.Open "SELECT *FROM PINAKES WHERE TYPOS=100 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
    
    Do While Not R.EOF
            COUNTER = COUNTER + 1
         If R("typos") = 100 Then
             PIN(COUNTER, 1) = R("PERIGRAFH")    'ΠΕΔΙΟ ΤΟΥ ΚΥΡΙΟΥ ΑΡΧΕΙΟΥ (SQLSERVER)
             PIN(COUNTER, 2) = R("TIMH")    ' ΑΡΙΘΜΟΣ ΣΤΗΛΗΣ ΤΟΥ EXCEL
         End If
         R.MoveNext
     Loop
        





     rSQL.Open "SELECT * FROM EGGTIM WHERE YEAR(HME)=1999", Gdb, adOpenDynamic, adLockOptimistic

        Dim Z

170     Z = 0
180     k = 0

        On Error GoTo 0  ' WRITEERROR    ' Resume Next

        Dim OK As Boolean

        Dim M_CODE

        Dim RR      As New ADODB.Recordset

        Dim RR2     As New ADODB.Recordset

        Dim TREX_YP As Single

        Dim mField As String
         Dim mSthlh As Integer
        
        
        Dim nn      As Integer

        ' If IsNull(xlsheet.cells(mRow, KOD_COLUMN)) Then

        Dim FOUND   As Boolean

        ' mRow = 1
190     mRow = Val(Text2.Text)

200     Do While True  '   Not IsNull(xlsheet.cells(mRow, 1))


            DoEvents
210         If IsNull(xlsheet.cells(mRow, 1)) Then

                Exit Do

            End If

220         If IsEmpty(xlsheet.cells(mRow, 1)) Then

                Exit Do

            End If

230         OK = True

240        ' M_CODE = xlsheet.cells(mRow, 1)
            '  On Error Resume Next

250         FOUND = True

'260         RR.Open "select * FROM PEL WHERE EIDOS='" + Left(Combo1.Text, 1) + "' and KOD='" + M_CODE + "'", Gdb, adOpenDynamic, adLockOptimistic
'
'270         If RR.EOF Then
'280             FOUND = False
'290             MsgBox "ΔΕΝ ΒΡΕΘΗΚΕ O ΠΕΛΑΤΗΣ ΜΕ ΚΩΔΙΚΟ  " + M_CODE
'300             List2.AddItem "ΔΕΝ ΒΡΕΘΗΚΕ : " + M_CODE
'310             xlsheet.cells(mRow, 4) = "ΔΕΝ ΒΡΕΘΗΚΕ"
'            End If

320       '/  M_CODE = RR("KOD")
330       '  RR.Close

340         If Len(xlsheet.cells(mRow, 2)) > 0 Then





350             rSQL.AddNew


                rSQL("apot") = 1
                rSQL("ID_NUM") = n21
360             rSQL("eidos") = "r"  ' Left(Combo1.Text, 1)

370             rSQL("ATIM") = "λ000001"    ' M_ATIM '"λ00002"
380             rSQL("hme") = apot.Value





               '' For i = 1 To COUNTER
                 '  mField = PIN(i, 1)
                  ' mSthlh = PIN(i, 2)
                   'rSQL(mField) = xlsheet.cells(mRow, mSthlh)
                'Next
                
                For i = 1 To COUNTER
                   mField = PIN(i, 1)
                   mSthlh = PIN(i, 2)
                   rSQL(mField) = xlsheet.cells(mRow, mSthlh)
                Next
                
                
                
                
450             rSQL.Update
            End If

460         mRow = mRow + 1

470         If mRow Mod 10 = 0 Then
480             Me.Caption = mRow
            End If

490         DoEvents
        Loop

500     rSQL.Close

510     'If List2.ListCount > 0 Then
520         MsgBox "ΤΑ ΛΑΘΗ ΑΠΟΘΗΚΕΥΤΗΚΑΝ ΣΤΟ EXCEL   "
             On Error Resume Next
530         xlwbook.save  '  "c:\ektyp2.xls"
       ' End If

540     Set xlsheet = Nothing
550     Set xlwbook = Nothing
        ' excel.Quit

560     xl.Quit

570     MsgBox "Ενημερώθηκαν " + Format(mRow, "###0") + " εγγραφές"

        
        Exit Sub

WRITEERROR:

580     Resume Next

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu53.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu53.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command6_Click()

        'apo  CSV  ARXEIO
        '<EhHeader>
        On Error GoTo Command6_Click_Err

        '</EhHeader>

        Dim R   As New ADODB.Recordset

        Dim r0  As New ADODB.Recordset

        Dim mok As Integer

        Dim PIN(30, 2)

        Dim COUNTER    As Integer

        Dim KOD_OK     As Boolean    ' ΕΧΕΙ ΔΗΛΩΣΕΙ ΤΟΝ ΚΩΔΙΚΟ

        Dim BARCODE_OK As Boolean    ' ΕΧΕΙ ΔΗΛΩΣΕΙ TA BARCODES

        Dim sql        As String

        Dim NON_STOP   As Boolean

100     NON_STOP = False
110     KOD_OK = False
120     BARCODE_OK = False

        Dim k As Long

130     COUNTER = 1  ' ARIUMOS PEDION POY THA METAFERTHOYN

        Dim KOD_COLUMN     As Integer

        Dim BARCODE_COLUMN As Integer

140     r0.Open "SELECT TOP 1 * FROM EID ", Gdb, adOpenDynamic, adLockOptimistic

        ' FORTONO TON PINAKA 20 ME TIS ANTISTOIXISEIS TON PEDION
        ' ΚΑΙ ΕΛΕΓΧΩ ΑΝ ΕΧΕΙ ΔΗΛΩΘΕΙ Ο ΚΩΔΙΚΟΣ
150     R.Open "SELECT *FROM PINAKES WHERE TYPOS=20 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

160     Do While Not R.EOF

170         If R("typos") = 20 Then
180             PIN(COUNTER, 1) = R("PERIGRAFH")    'ΠΕΔΙΟ ΤΟΥ ΚΥΡΙΟΥ ΑΡΧΕΙΟΥ (SQLSERVER)
190             PIN(COUNTER, 2) = R("TIMH")    ' ΑΡΙΘΜΟΣ ΣΤΗΛΗΣ ΤΟΥ EXCEL

200             If UCase(R("PERIGRAFH")) = "KOD" Then
210                 KOD_OK = True
220                 KOD_COLUMN = COUNTER
                End If

230             If UCase(R("PERIGRAFH")) = "BARCODE" Then
240                 BARCODE_OK = True
250                 BARCODE_COLUMN = PIN(COUNTER, 2)    ' COUNTER
                End If

                ' On Error GoTo LATOSPEDIO
260             mok = 0

270             If UCase(R("PERIGRAFH")) = "BARCODE" Then
280                 mok = 1
                Else

290                 For k = 0 To r0.FIELDS.Count - 1     'ΕΛΕΓΧΟΣ ΑΝ ΕΙΝΑΙ ΥΠΑΡΚΤΟ ΤΟ ΠΕΔΙΟ

300                     If r0(k).Name = R("PERIGRAFH") Then
310                         mok = 1

                            Exit For

                        End If

                    Next

                End If

320             If mok = 0 Then
330                 MsgBox "Λάθος στην δήλωση των πεδίων." + Chr(13) + "Δεν υπάρχει στο αρχείο ειδών το πεδίο " + R("PERIGRAFH")

                    Exit Sub

                End If

                On Error GoTo Command6_Click_Err

340             COUNTER = COUNTER + 1
            Else

                Exit Do

            End If

350         R.MoveNext
        Loop

        ' kathgories

360     R.Close

370     If KOD_OK = False Then
380         MsgBox "δεν εχετε δηλώσει τον κωδικό του είδους" + Chr(13) + " Σταματά η ενημέρωση"

            Exit Sub

        End If

        '-----------------------------------------------------------------------
        'Set xlwbook = xl.Workbooks.Open(Text1.Text)
        'Set xlsheet = xlwbook.Sheets.Item(1)

390     Open Text1.Text For Input As #1

        Dim ko   As String

        Dim mNew As Long, mUpd As Long

400     mNew = 0
410     mUpd = 0

420     Label4.Caption = "Nέες εγγραφές 0"
430     Label5.Caption = "Ενημέρωση εγγραφές 0"

        Dim AA, DUM

        Dim mRow As Long

        ' data1.Recordset.MoveFirst

        'KENES SEIRES
440     mRow = Val(Text2.Text)    '  data1.Recordset.Move

450     For k = 1 To mRow
460         Input #1, AA
        Next

        Dim ELEM(1 To 60)

        On Error GoTo error_name

470     Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF

480         Line Input #1, AA

            ' DUM = to437(AA)
490         If EOF(1) Then

                Exit Do

            End If

500         AA = to928(AA)
510         DUM = bohu5.FETES_DELIM(AA, ELEM)

            '     If IsEmpty(xlsheet.cells(mRow, KOD_COLUMN)) Then
            '        Exit Do
            '   End If

520         ko = Trim(ELEM(KOD_COLUMN))    '  ' ΚΩΔΙΚΟς ΤΟΥ ΕΙΔΟΥΣ  EXCEL

530         R.Open "SELECT * FROM EID WHERE KOD='" + ko + "'", Gdb, adOpenDynamic, adLockOptimistic

540         If R.EOF Then    'DEN BRHKA TON KODIKO
550             R.AddNew
560             R("KOD") = Left(ko, 14)
570             R("fpa") = Val(Left(Combo2.Text, 2))
580             mNew = mNew + 1
590             Label4.Caption = "Nέες εγγραφές " + Format(mNew, "######")
            Else  'BRHKA TON KODIKO
600             mUpd = mUpd + 1
610             Label5.Caption = "Ενημέρωση εγγραφές " + Format(mUpd, "######")
            End If

            'ΕΝΗΜΕΡΩΝΩ ΤΑ ΥΠΟΛΟΙΠΑ ΠΕΔΙΑ ΑΠΟ ΤΟ EXCEL
620         For k = 1 To COUNTER - 1

630             If PIN(k, 1) = "BARCODE" Then
                Else

640                 If R(PIN(k, 1)).Type = 202 Or R(PIN(k, 1)).Type = 129 Then    ' KEIMENO
650                     R(PIN(k, 1)) = Left(ELEM(PIN(k, 2)), R(PIN(k, 1)).DefinedSize)
                    Else    'ARITMOS
660                     R(PIN(k, 1)) = ELEM(PIN(k, 2))
                    End If
                End If

            Next

670         R.Update

680         R.Close

690         If KOD_OK And BARCODE_OK Then

700             SQL2 = ELEM(BARCODE_COLUMN)
710             sql = "INSERT INTO BARCODES (KOD,ERG) VALUES ('" + ko + "','" + LTrim(SQL2) + "')"

720             If Len(SQL2) > 0 Then
730                 Gdb.Execute sql
                End If

            End If

740         DoEvents
750         Me.Caption = ko
760         mRow = mRow + 1    'data1.Recordset.MoveNext
        Loop

        'xl.Quit
        'Set xlwbook = Nothing
        'Set xl = Nothing

770     Close #1

780     MsgBox "τέλος ενημέρωσης"

        Exit Sub

error_name:

790     If NON_STOP = False Then
800         MsgBox "λάθος στην σειρά " + Format(mRow, "#######")

810         If MsgBox("ΤΕΡΜΑΤΙΣΜΟΣ ;", vbYesNo) = vbYes Then

                Exit Sub

            End If

820         If MsgBox("ΣΥΝΕΧΕΙΑ ΧΩΡΙΣ ΕΡΩΤΗΣΗ ;", vbYesNo) = vbYes Then
830             NON_STOP = True
            End If
        End If

840     Resume Next

        '<EhFooter>
        Exit Sub

Command6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu53.Command6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu53.Command6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
        Dim R As New ADODB.Recordset

100     mForm_Load Me, fh, fw, ft, fl
110     R.Open "SELECT *FROM PINAKES WHERE TYPOS=1 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

        'FPA
120     Do While Not R.EOF

130         If R("typos") = 1 Then
140             Combo2.AddItem str(R("AYJON")) + " -> " + str(R("TIMH"))
            End If

150         R.MoveNext
        Loop

160     R.Close

170     Combo2.Text = Combo2.List(1)

180     Combo1.Text = Combo1.List(0)



 Dim N0 As Integer
   N0 = GGET_NVALUE("select count(*) from PINAKES WHERE TYPOS=20")
   If N0 < 10 Then
       Gdb.Execute "DELETE FROM PINAKES WHERE TYPOS=20"
       Dim nc As Integer
       For nc = 1 To 10
          Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,TIMH,PERIGRAFH) VALUES (20," + str(nc) + "," + str(2000 + nc) + ",'ΛΙΣΤΑ " + str(nc) + "')"
       Next
   End If
   Dim R20 As New ADODB.Recordset
   R20.Open "select isnull(PERIGRAFH,' ') AS PERIGRAFH,ISNULL(TIMH,0) AS TIMH FROM PINAKES WHERE TYPOS=20", Gdb, adOpenDynamic, adLockOptimistic
   Do While Not R20.EOF
      List2.AddItem R20!PERIGRAFH + ";" + str(R20!TIMH)
      R20.MoveNext
   Loop
   R20.Close






FILL_FIELDS "EID"

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu53.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu53.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 13875, 9480, fh, fw, ft, fl

End Sub

Private Sub EGGTIM_Click()

        '<EhHeader>
        On Error GoTo Label6_Click_Err

        '</EhHeader>
        Dim F As New DataEntry2, DD

        Dim R As New ADODB.Recordset

        'pinakas pedion
100     UPDATE2_PINAKES 100, Label6   ', Combo3

        '<EhFooter>
        Exit Sub

Label6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu53.Label6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu53.Label6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub







Private Sub Label6_Click()
' LISTA ME IMPORTS EIDON
        '<EhHeader>
        On Error GoTo Label6_Click_Err

'   Dim n0 As Integer
'   n0 = GGET_NVALUE("select count(*) from PINAKES WHERE TYPOS=20")
'   If n0 < 10 Then
'       Gdb.Execute "DELETE FROM PINAKES WHERE TYPOS=20"
'       Dim NC As Integer
'       For NC = 1 To 10
'          Gdb.Execute "INSERT INTO PINAKES (TYPOS,AYJON,TIMH,PERIGRAFH) VALUES (20," + str(NC) + "," + Comp(2000 + NC) + ",'ΛΙΣΤΑ " + Comp(NC) + "')"
'       Next
'   End If
'   Dim R20 As New ADODB.Recordset
'   Do While Not R20.EOF
'      List2.AddItem R20!PERIGRAFH + ";" + str(R20!TIMH)
'      R20.MoveNext
'   ENDDO
'   R20.Close
   Dim c As String
   c = List2.List(List2.ListIndex)   'ΛΙΣΤΑ  1; 2001
   Dim nn As Integer
   nn = Split(c, ";")(1) '2001
   

        '</EhHeader>
        Dim F As New DataEntry2, DD

        Dim R As New ADODB.Recordset

        'pinakas pedion
100     UPDATE2_PINAKES nn, Label6    ', Combo3

        '<EhFooter>
        Exit Sub

Label6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu53.Label6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu53.Label6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

