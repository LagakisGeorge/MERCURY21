VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form bohu51 
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
   Begin VB.CommandButton HELPTIM 
      Caption         =   "?"
      Height          =   360
      Left            =   10080
      TabIndex        =   26
      Top             =   3000
      Width           =   270
   End
   Begin VB.CommandButton ImportTim 
      Caption         =   "Εισαγωγή Τιμολογίων"
      Height          =   360
      Left            =   6480
      TabIndex        =   25
      ToolTipText     =   $"bohu51ok.frx":0000
      Top             =   3000
      Width           =   3615
   End
   Begin VB.CommandButton TIMOKAT 
      Caption         =   "ΕΙΣΑΓΩΓΗ ΤΙΜΟΚΑΤΑΛΟΓΩΝ"
      Height          =   375
      Left            =   10512
      TabIndex        =   21
      Top             =   1776
      Width           =   2415
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4440
      TabIndex        =   20
      Text            =   "Combo3"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      ItemData        =   "bohu51ok.frx":0076
      Left            =   6768
      List            =   "bohu51ok.frx":0080
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   2160
      Width           =   3228
   End
   Begin MSComCtl2.DTPicker APOT 
      Height          =   276
      Left            =   6768
      TabIndex        =   18
      Top             =   1776
      Width           =   3204
      _ExtentX        =   5662
      _ExtentY        =   503
      _Version        =   393216
      Format          =   156106753
      CurrentDate     =   39565
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Εισαγωγή κινήσεων πελατών από EXCEL"
      Height          =   255
      Left            =   6768
      TabIndex        =   17
      ToolTipText     =   "ΚΩΔΙΚΟΣ,ΠΟΣΟ ΧΡΕΩΣΗΣ(ΜΕ - ΠΙΣΤΩΣΗ),ΑΙΤΙΟΛΟΓΙΑ"
      Top             =   1488
      Width           =   3204
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
      Height          =   3840
      Left            =   7860
      TabIndex        =   16
      Top             =   4650
      Width           =   5640
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Δημιουργία αρχείου με τα πεδία"
      Height          =   480
      Left            =   7860
      TabIndex        =   12
      Top             =   3888
      Width           =   5640
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Εξοδος"
      Height          =   555
      Left            =   12030
      TabIndex        =   11
      Top             =   3000
      Width           =   1470
   End
   Begin VB.ComboBox Combo2 
      Height          =   288
      ItemData        =   "bohu51ok.frx":00B2
      Left            =   4524
      List            =   "bohu51ok.frx":00BC
      TabIndex        =   10
      Top             =   1224
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "Πρόσθεση νέων εγγραφών σε Κύριο αρχείο"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1164
      TabIndex        =   8
      Top             =   912
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "Ενημέρωση τιμών,περιγραφών από Excel"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1164
      TabIndex        =   7
      Top             =   552
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4884
      TabIndex        =   5
      Text            =   "3"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Height          =   377
      Left            =   5724
      TabIndex        =   4
      Top             =   1735
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3648
      TabIndex        =   2
      Text            =   " "
      Top             =   1728
      Width           =   2076
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
      Left            =   204
      TabIndex        =   1
      Top             =   2520
      Width           =   5784
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
      Enabled         =   -1  'True
      TextRTF         =   $"bohu51ok.frx":00DA
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Εισαγωγή Κινήσεων"
      Height          =   1692
      Left            =   6432
      TabIndex        =   22
      Top             =   1008
      Width           =   3708
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Caption         =   "Εισαγωγή Τιμοκαταλόγων"
      Height          =   1692
      Left            =   10224
      TabIndex        =   23
      Top             =   1008
      Width           =   3276
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Πίνακας με πεδία KANTE ΚΛΙΚ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   390
         TabIndex        =   24
         Top             =   390
         Width           =   2250
      End
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
      Left            =   3144
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
      Caption         =   "ΠΕΛΑΤΕΣ / ΠΡΟΜΗΘΕΥΤΕΣ"
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
      Height          =   252
      Left            =   96
      TabIndex        =   3
      Top             =   1824
      Width           =   3228
   End
End
Attribute VB_Name = "bohu51"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim xl       As New excel.Application

Dim xlsheet  As excel.Worksheet

Dim xlsheet3 As excel.Worksheet

Dim xlwbook  As excel.workbook

Dim fh(0 To 100), fw(0 To 100), ft(0 To 100), fl(0 To 100)

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

100     KOD_OK = False
110     BARCODE_OK = False

        Dim k As Long

120     COUNTER = 1  ' ARIUMOS PEDION POY THA METAFERTHOYN

        Dim KOD_COLUMN     As Integer

        Dim BARCODE_COLUMN As Integer

130     r0.Open "SELECT TOP 1 * FROM PEL ", Gdb, adOpenDynamic, adLockOptimistic

        ' FORTONO TON PINAKA 20 ME TIS ANTISTOIXISEIS TON PEDION
        ' ΚΑΙ ΕΛΕΓΧΩ ΑΝ ΕΧΕΙ ΔΗΛΩΘΕΙ Ο ΚΩΔΙΚΟΣ
140     R.Open "SELECT *FROM PINAKES WHERE TYPOS=21 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

150     Do While Not R.EOF

160         If R("typos") = 21 Then
170             PIN(COUNTER, 1) = R("PERIGRAFH")    'ΠΕΔΙΟ ΤΟΥ ΚΥΡΙΟΥ ΑΡΧΕΙΟΥ (SQLSERVER)
180             PIN(COUNTER, 2) = R("TIMH")    ' ΑΡΙΘΜΟΣ ΣΤΗΛΗΣ ΤΟΥ EXCEL

190             If UCase(R("PERIGRAFH")) = "KOD" Then
200                 KOD_OK = True
210                 KOD_COLUMN = COUNTER
                End If

                ' On Error GoTo LATOSPEDIO
220             mok = 0    'ΕΛΕΓΧΩ ΑΝ ΤΑ ΠΕΔΙΑ ΠΟΥ ΔΗΛΩΣΑ ΣΤΟ <<PINAKES>> ΥΠΑΡΧΟΥΝ ΣΤΟΝ ΠΙΝΑΚΑ PEL

230             For k = 0 To r0.FIELDS.Count - 1

240                 If r0(k).Name = R("PERIGRAFH") Then
250                     mok = 1

                        Exit For

                    End If

                Next

260             If mok = 0 Then
270                 MsgBox "Λάθος στην δήλωση των πεδίων." + Chr(13) + "Δεν υπάρχει στο αρχείο ειδών το πεδίο " + R("PERIGRAFH")

                    Exit Sub

                End If

                On Error GoTo Command1_Click_Err

280             COUNTER = COUNTER + 1
            Else

                Exit Do

            End If

290         R.MoveNext
        Loop

        ' kathgories

300     R.Close

310     If KOD_OK = False Then
320         MsgBox "δεν εχετε δηλώσει τον κωδικό του πελατη/προμηθευτή" + Chr(13) + " Σταματά η ενημέρωση"

            Exit Sub

        End If

330     Set xlwbook = xl.Workbooks.Open(Text1.Text)
340     Set xlsheet = xlwbook.Sheets.Item(1)

        Dim ko   As String

        Dim mNew As Long, mUpd As Long

350     mNew = 0
360     mUpd = 0

370     Label4.Caption = "Nέες εγγραφές 0"
380     Label5.Caption = "Ενημέρωση εγγραφές 0"

        Dim mRow As Long

        ' data1.Recordset.MoveFirst
390     mRow = Val(Text2.Text)    '  data1.Recordset.Move

        On Error GoTo error_name

400     Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF

410         If IsNull(xlsheet.cells(mRow, KOD_COLUMN)) Then

                Exit Do

            End If

420         If IsEmpty(xlsheet.cells(mRow, KOD_COLUMN)) Then

                Exit Do

            End If

430         ko = Trim(xlsheet.cells(mRow, KOD_COLUMN))    '  ' ΚΩΔΙΚΟς ΤΟΥ ΕΙΔΟΥΣ  EXCEL

440         R.Open "SELECT * FROM PEL WHERE EIDOS='" + Left(Combo2.Text, 1) + "' AND KOD='" + ko + "'", Gdb, adOpenDynamic, adLockOptimistic

450         If R.EOF Then    'DEN BRHKA TON KODIKO
460             R.AddNew
470             R("KOD") = Left(ko, 14)
                '        R("fpa") = Val(Left(Combo2.Text, 2))
480             mNew = mNew + 1
490             Label4.Caption = "Nέες εγγραφές " + Format(mNew, "######")
            Else  'BRHKA TON KODIKO
500             mUpd = mUpd + 1
510             Label5.Caption = "Ενημέρωση εγγραφές " + Format(mUpd, "######")
            End If

            'ΕΝΗΜΕΡΩΝΩ ΤΑ ΥΠΟΛΟΙΠΑ ΠΕΔΙΑ ΑΠΟ ΤΟ EXCEL
520         For k = 1 To COUNTER - 1

530             If PIN(k, 1) = "BARCODE" Then
                Else

540                 If R(PIN(k, 1)).Type = 202 Then    ' KEIMENO
550                     R(PIN(k, 1)) = Left(xlsheet.cells(mRow, PIN(k, 2)), R(PIN(k, 1)).DefinedSize)
                    Else    'ARITMOS
560                     R(PIN(k, 1)) = xlsheet.cells(mRow, PIN(k, 2))
                    End If
                End If

            Next

570         R("EIDOS") = Left(Combo2.Text, 1)
580         R.Update

590         R.Close

600         DoEvents
610         Me.Caption = ko
620         mRow = mRow + 1    'data1.Recordset.MoveNext
        Loop

630     xl.Quit
640     Set xlwbook = Nothing
650     Set xl = Nothing

660     MsgBox "τέλος ενημέρωσης"

        Exit Sub

error_name:
670     mok = MsgBox("λάθος στην σειρά " + Format(mRow, "#######"), vbYesNo, "Διακοπή ;")

680     If mok = vbYes Then Exit Sub

690     Resume Next

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu51.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu51.Command1_Click " & "at line " & Erl

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
         "in ADOMERCNEW.bohu51.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu51.Command2_Click " & "at line " & Erl

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
         "in ADOMERCNEW.bohu51.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu51.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub PEDIA() ' Command4_Click()

        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>

        Dim R As New ADODB.Recordset
        
        
       R.Open "SELECT COLUMN_NAME,DATA_TYPE,CHARACTER_MAXIMUM_LENGTH FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='PEL' ", Gdb, adOpenForwardOnly, adLockReadOnly

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
        Exit Sub
        
        
        

        On Error GoTo CREATEFIELDS

100     R.Open "select count(*) FROM MFIELDS", Gdb, adOpenForwardOnly, adLockReadOnly

110     If R(0) > 10 Then
120         R.Close
130        ' R.Open "select RTRIM(TITLOS)+'  : '+ONOMA FROM MFIELDS WHERE MFILE='EID'", Gdb, adOpenForwardOnly, adLockReadOnly

140       '  Do While Not R.EOF
150       '      List1.AddItem R(0)
160       '      R.MoveNext
          '  Loop

          '  Exit Sub

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










Dim q As String
q = "SELECT COLUMN_NAME,DATA_TYPE,ISNULL(CHARACTER_MAXIMUM_LENGTH,0) AS LL,ISNULL(M.ONOMA,'') AS PEDIO FROM INFORMATION_SCHEMA.COLUMNS F LEFT JOIN MFIELDS M ON LEFT(M.MFILE,3)='PEL' AND RTRIM(F.COLUMN_NAME)=RTRIM(M.TITLOS)   WHERE TABLE_NAME='PEL' "
On Error Resume Next


      R.Open q, Gdb, adOpenForwardOnly, adLockReadOnly
        
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
540     Gdb.Execute "create table MFIELDS (MFILE   char(20)," & "TITLOS CHAR(20)," & "ONOMA  char(30)," & "TYPOS  char(7)  )"
550     MsgBox "Ξαναπατήστε το πλήκτρο 'Δημιουργία εγγραφών' "

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu51.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu51.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command5_Click()

        'ΛΗΨΗ ΑΠΟ EXCEL
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

100     M_ATIM = "Σ" + Format(Time(), "hhmmss")

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

160     rSQL.Open "SELECT * FROM EGG WHERE YEAR(HME)=1999", Gdb, adOpenDynamic, adLockOptimistic

        Dim Z

170     Z = 0
180     k = 0

        On Error GoTo WRITEERROR    ' Resume Next

        Dim OK As Boolean

        Dim M_CODE

        Dim RR      As New ADODB.Recordset

        Dim RR2     As New ADODB.Recordset

        Dim TREX_YP As Single

        Dim nn      As Integer

        ' If IsNull(xlsheet.cells(mRow, KOD_COLUMN)) Then

        Dim FOUND   As Boolean

        ' mRow = 1
190     mRow = Val(Text2.Text)

200     Do While True  '   Not IsNull(xlsheet.cells(mRow, 1))

210         If IsNull(xlsheet.cells(mRow, 1)) Then

                Exit Do

            End If

220         If IsEmpty(xlsheet.cells(mRow, 1)) Then

                Exit Do

            End If

230         OK = True

240         M_CODE = xlsheet.cells(mRow, 1)
            '  On Error Resume Next

250         FOUND = True

260         RR.Open "select * FROM PEL WHERE EIDOS='" + Left(Combo1.Text, 1) + "' and KOD='" + M_CODE + "'", Gdb, adOpenDynamic, adLockOptimistic

270         If RR.EOF Then
280             FOUND = False
290             MsgBox "ΔΕΝ ΒΡΕΘΗΚΕ O ΠΕΛΑΤΗΣ ΜΕ ΚΩΔΙΚΟ  " + M_CODE
300             List2.AddItem "ΔΕΝ ΒΡΕΘΗΚΕ : " + M_CODE
310             xlsheet.cells(mRow, 4) = "ΔΕΝ ΒΡΕΘΗΚΕ"
            End If

320         M_CODE = RR("KOD")
330         RR.Close

340         If FOUND And Abs(xlsheet.cells(mRow, 2)) > 0 Then

350             rSQL.AddNew
360             rSQL("eidos") = Left(Combo1.Text, 1)

370             rSQL("ATIM") = "Σ00001"    ' M_ATIM '"λ00002"
380             rSQL("hme") = apot.Value

390             rSQL("ait") = xlsheet.cells(mRow, 3)
400             rSQL("KOD") = M_CODE

410             rSQL("XRE") = xlsheet.cells(mRow, 2)

420             If xlsheet.cells(mRow, 2) > 0 Then
430                 rSQL("XREOSI") = xlsheet.cells(mRow, 2)
                    rSQL("PISTOSI") = 0
                Else
440                 rSQL("PISTOSI") = -xlsheet.cells(mRow, 2)
                    rSQL("XREOSI") = 0
                End If

450             rSQL.Update
            End If

460         mRow = mRow + 1

470         If mRow Mod 10 = 0 Then
480             Me.Caption = mRow
            End If

490         DoEvents
        Loop
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        

500     rSQL.Close

510     If List2.ListCount > 0 Then
520         MsgBox "ΤΑ ΛΑΘΗ ΑΠΟΘΗΚΕΥΤΗΚΑΝ ΣΤΟ EXCEL  "
530         xlwbook.save  '  "c:\ektyp2.xls"
        End If

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
         "in ADOMERCNEW.bohu51.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu51.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub HELPTIM_Click()
  MsgBox " A            B          C         D        E        F      G" + Chr(13) + "HME        ATIM        KODPEL     AJ2       FPA2     AJ1    FPA1" + Chr(13) + "ΑΠΟ ΤΟ COMBO ΠΕΛΑΤΗΣ/ΠΡΟΜΗΘΕΥΤΗΣ ΓΙΝΕΤΑΙ ΑΓΟΡΑΣ ή ΠΩΛΗΣΗΣ" + Chr(13) + "ΤΟ ΠΑΡΑΣΤΑΤΙΚΟ ΠΟΥ ΔΗΜΙΟΥΡΓΕΙΤΑΙ ΕΙΝΑΙ ΤΟ Σ??????"
End Sub

Private Sub ImportTim_Click()

        'ΛΗΨΗ ΑΠΟ EXCEL
        '  A            B          C         D        E        F      G
         'HME        ATIM        KODPEL     AJ2       FPA2     AJ1    FPA1
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

100     M_ATIM = "Σ" + Format(Time(), "hhmmss")
List1.Clear
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

160     rSQL.Open "SELECT top 1 * FROM TIM ", Gdb, adOpenDynamic, adLockOptimistic

        Dim Z

170     Z = 0
180     k = 0

        On Error GoTo WRITEERROR    ' Resume Next

        Dim OK As Boolean

        Dim M_CODE

        Dim RR      As New ADODB.Recordset

        Dim RR2     As New ADODB.Recordset

        Dim TREX_YP As Single

        Dim nn      As Integer

        ' If IsNull(xlsheet.cells(mRow, KOD_COLUMN)) Then

        Dim FOUND   As Boolean
        Dim mHME As String
        ' mRow = 2
190     mRow = Val(Text2.Text)







200     Do While True  '   Not IsNull(xlsheet.cells(mRow, 1))

210         If IsNull(xlsheet.cells(mRow, 1)) Then

                Exit Do

            End If

220         If IsEmpty(xlsheet.cells(mRow, 1)) Then

                Exit Do

            End If

230         OK = True

240
            
            
            
                       
            '  On Error Resume Next

250         FOUND = True
               M_CODE = xlsheet.cells(mRow, xlsheet.range("F" & 1).Column)
               M_CODE = LTrim(str(M_CODE))
               
               
               Dim M_AFM As String
                M_AFM = xlsheet.cells(mRow, xlsheet.range("B" & 1).Column)
              ' M_CODE = LTrim(str(M_CODE))
               
260         RR.Open "select * FROM PEL WHERE EIDOS='" + Left(Combo1.Text, 1) + "' and AFM='" + M_AFM + "'", Gdb, adOpenDynamic, adLockOptimistic

270         If RR.EOF Then
                'Gdb.Execute "iNSERT into PEL (EIDOS,KOD,AFM"
280             FOUND = False
290             Dim anss As Integer: anss = MsgBox("ΔΕΝ ΒΡΕΘΗΚΕ O ΠΕΛΑΤΗΣ ΜΕ ΚΩΔΙΚΟ  " + M_CODE + " τερματισμός?", vbYesNo)
                If anss = vbNo Then
                    Exit Sub
                End If
300             List1.AddItem "ΔΕΝ ΒΡΕΘΗΚΕ :  LINE: " + str(mRow) + "--" + M_CODE
310             xlsheet.cells(mRow, 10) = "ΔΕΝ ΒΡΕΘΗΚΕ"
            End If

320         M_CODE = RR("KOD")
330         RR.Close
             Dim MATIM As String
340         If FOUND Then '    And Abs(xlsheet.cells(mRow, xlsheet.range("AH" & 1).Column)) + Abs(xlsheet.cells(mRow, xlsheet.range("AL" & 1).Column)) + Abs(xlsheet.cells(mRow, xlsheet.range("AM" & 1).Column)) + Abs(xlsheet.cells(mRow, xlsheet.range("AO" & 1).Column)) > 0 Then

350           ' rSQL.AddNew
360            ' rSQL("eidos") = Left(Combo1.Text, 1)
                If "ΑΛΠ" = Left(xlsheet.cells(mRow, 2), 3) Then
                     MATIM = "L" + Right(xlsheet.cells(mRow, 2), xlsheet.range("G" & 1).Column) ' M_ATIM '"λ00002"
                Else
370                  MATIM = "T" + Right(xlsheet.cells(mRow, xlsheet.range("G" & 1).Column), 6) ' M_ATIM '"λ00002"
                End If
                Dim hme As String: hme = xlsheet.cells(mRow, xlsheet.range("E" & 1).Column)
                mHME = Split(hme, "/")(1) + "/" + Split(hme, "/")(0) + "/" + Split(hme, "/")(2)
               '  rSQL!ATIM = MATIM
380

390
              Gdb.BeginTrans
              Dim n1 As Integer, problem As Integer
              problem = 0
              'On Error GoTo trans
              Gdb.Execute "INSERT INTO TIM (EIDOS,B_N1,ATIM,KLEIDI,KPE,HME) VALUES ('" + Left(Combo1.Text, 1) + "',1,'" + MATIM + "','" + MATIM + "','" + M_CODE + "','" + mHME + "')", nn
           If nn > 0 Then
                
                          Dim nc As Long: nc = GET_NVALUE("SELECT MAX(ID_NUM) FROM TIM")
 
                 Gdb.Execute "update TIM SET AJ1=" + Replace(str(xlsheet.cells(mRow, xlsheet.range("BU" & 1).Column)), ",", ".") + " where ID_NUM=" + str(nc)
                 Gdb.Execute "update TIM SET AJ2=" + Replace(str(xlsheet.cells(mRow, xlsheet.range("BZ" & 1).Column)), ",", ".") + " where ID_NUM=" + str(nc)
                 Gdb.Execute "update TIM SET AJ6=" + Replace(str(xlsheet.cells(mRow, xlsheet.range("BY" & 1).Column)), ",", ".") + " where ID_NUM=" + str(nc)
                 Gdb.Execute "update TIM SET AJ4=" + Replace(str(xlsheet.cells(mRow, xlsheet.range("CB" & 1).Column)), ",", ".") + " where ID_NUM=" + str(nc)
                 
                 Gdb.Execute "update TIM SET FPA1=" + Replace(str(xlsheet.cells(mRow, xlsheet.range("AG" & 1).Column)), ",", ".") + " where ID_NUM=" + str(nc)
                 Gdb.Execute "update TIM SET FPA2=" + Replace(str(xlsheet.cells(mRow, xlsheet.range("AL" & 1).Column)), ",", ".") + " where ID_NUM=" + str(nc)
                 Gdb.Execute "update TIM SET FPA6=" + Replace(str(xlsheet.cells(mRow, xlsheet.range("AK" & 1).Column)), ",", ".") + " where ID_NUM=" + str(nc)
                 Gdb.Execute "update TIM SET FPA4=" + Replace(str(xlsheet.cells(mRow, xlsheet.range("AM" & 1).Column)), ",", ".") + " where ID_NUM=" + str(nc)


                 Gdb.Execute "update TIM SET AJ3=0,FPA3=0,AJ5=0,AJ7=0,FPA7=0 where ID_NUM=" + str(nc)
                 Gdb.Execute "update TIM SET AJI=AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7+FPA1+FPA2+FPA3+FPA4+FPA6+FPA7 where ID_NUM=" + str(nc)
400            ' rSQL("KPE") = M_CODE
                'rSQL("hme") = xlsheet.cells(mRow, xlsheet.range("E" & 1).Column)
'                rSQL("AJ1") = xlsheet.cells(mRow, xlsheet.range("BV" & 1).Column) '13
'                rSQL("AJ2") = xlsheet.cells(mRow, xlsheet.range("CA" & 1).Column) '24
'
'                rSQL("AJ6") = xlsheet.cells(mRow, xlsheet.range("BZ" & 1).Column) '9
'                rSQL("AJ4") = xlsheet.cells(mRow, xlsheet.range("CC" & 1).Column) '17
'
'
'                 ''''rSQL("hme") = xlsheet.cells(mRow, xlsheet.range(XLSHEET,"E"))
'                rSQL("FPA1") = xlsheet.cells(mRow, xlsheet.range("AH" & 1).Column) '13
'                rSQL("FPA2") = xlsheet.cells(mRow, xlsheet.range("AM" & 1).Column) '24
'
'                rSQL("FPA6") = xlsheet.cells(mRow, xlsheet.range("AL" & 1).Column) '9
'                rSQL("FPA4") = xlsheet.cells(mRow, xlsheet.range("AO" & 1).Column) '17
                
                 ' rSQL("AJ3") = 0: rSQL("FPA3") = 0
                 'rSQL("AJ5") = 0
                 ' rSQL("AJ7") = 0: rSQL("FPA7") = 0
                
                
                
                
410             'rSQL("FPA2") = xlsheet.cells(mRow, 9)
                'rSQL("AJ1") = 0 ' xlsheet.cells(mRow, 66)
                'rSQL("FPA1") = 0 ' xlsheet.cells(mRow, 67)
                'rSQL("AJ3") = 0: rSQL("AJ4") = 0: rSQL("AJ5") = 0: rSQL("AJ6") = 0: rSQL("AJ7") = 0
                'rSQL("FPA3") = 0: rSQL("FPA4") = 0: rSQL("FPA7") = 0: rSQL("FPA6") = 0
               ' rSQL("AJI") = rSQL("AJ1") + rSQL("AJ2") + rSQL("FPA1") + rSQL("FPA2") + rSQL("AJ4") + rSQL("AJ6") + rSQL("FPA4") + rSQL("FPA6")
               ' rSQL("KLEIDI") = rSQL("ATIM")
                'rSQL("B_N1") = 1
450            ' rSQL.Update
            End If

           
'           If xlsheet.cells(mRow, 4) > 0 Then
'               sql = "INSERT INTO EGGTIM (ATIM,ID_NUM,EIDOS,APOT,FPA,HME,PELKOD,POSO,KODE,TIMM) VALUES ('" + MATIM + "'," + str(nc) + ",'"
'               sql = sql + Left(Combo1.Text, 1) + "',1,2,'" + mHME + "','" + M_CODE + "',1,'K24'," + Replace(str(xlsheet.cells(mRow, 4)), ",", ".") + ")"
'               Gdb.Execute sql
'           End If
           
           If xlsheet.cells(mRow, xlsheet.range("BZ" & 1).Column) > 0 Then '24%
               sql = "INSERT INTO EGGTIM (ATIM,ID_NUM,EIDOS,APOT,FPA,HME,PELKOD,POSO,KODE,TIMM) VALUES ('" + MATIM + "'," + str(nc) + ",'"
               sql = sql + Left(Combo1.Text, 1) + "',1,2,'" + mHME + "','" + M_CODE + "',1,'K24'," + Replace(str(xlsheet.cells(mRow, xlsheet.range("BZ" & 1).Column)), ",", ".") + ")"
               Gdb.Execute sql, n1: If n1 = 0 Then problem = 1
           End If
           
           If xlsheet.cells(mRow, xlsheet.range("BU" & 1).Column) > 0 Then '13%
               sql = "INSERT INTO EGGTIM (ATIM,ID_NUM,EIDOS,APOT,FPA,HME,PELKOD,POSO,KODE,TIMM) VALUES ('" + MATIM + "'," + str(nc) + ",'"
               sql = sql + Left(Combo1.Text, 1) + "',1,1,'" + mHME + "','" + M_CODE + "',1,'K13'," + Replace(str(xlsheet.cells(mRow, xlsheet.range("BU" & 1).Column)), ",", ".") + ")"
                Gdb.Execute sql, n1: If n1 = 0 Then problem = 1
           End If
           If xlsheet.cells(mRow, xlsheet.range("BY" & 1).Column) > 0 Then ' 9%
               sql = "INSERT INTO EGGTIM (ATIM,ID_NUM,EIDOS,APOT,FPA,HME,PELKOD,POSO,KODE,TIMM) VALUES ('" + MATIM + "'," + str(nc) + ",'"
               sql = sql + Left(Combo1.Text, 1) + "',1,6,'" + mHME + "','" + M_CODE + "',1,'K9'," + Replace(str(xlsheet.cells(mRow, xlsheet.range("BY" & 1).Column)), ",", ".") + ")"
                Gdb.Execute sql, n1: If n1 = 0 Then problem = 1
           End If
           
           
            Dim Axia17 As Single: Axia17 = xlsheet.cells(mRow, xlsheet.range("CB" & 1).Column)
           If Axia17 > 0 Then '17%
              
               sql = "INSERT INTO EGGTIM (ATIM,ID_NUM,EIDOS,APOT,FPA,HME,PELKOD,POSO,KODE,TIMM) VALUES ('" + MATIM + "'," + str(nc) + ",'"
               sql = sql + Left(Combo1.Text, 1) + "',1,4,'" + mHME + "','" + M_CODE + "',1,'K17'," + Replace(str(Axia17), ",", ".") + ")"
                Gdb.Execute sql, n1: If n1 = 0 Then problem = 1
           End If
           
           If problem = 0 Then
               Gdb.CommitTrans
           Else
               Gdb.RollbackTrans
                List1.AddItem "line " + str(mRow) + " error"
           End If
           
           
     Else  '//  nn=0 opote grafei oti den phge kala
          List1.AddItem "line " + str(mRow) + " error"
          Gdb.RollbackTrans
     End If



460         mRow = mRow + 1

470         If mRow Mod 10 = 0 Then
480             Me.Caption = mRow
            End If

490         DoEvents
        Loop
        
        
  
        
        

500     rSQL.Close
 Gdb.Execute " update EGGTIM SET KAU_AJIA=POSO*TIMM;update EGGTIM SET MIK_AJIA=POSO*TIMM*1.24 WHERE FPA=2;update EGGTIM SET MIK_AJIA=POSO*TIMM*1.13 WHERE FPA=1;update EGGTIM SET MIK_AJIA=POSO*TIMM*1.09 WHERE FPA=6;update EGGTIM SET MIK_AJIA=POSO*TIMM*1.17 WHERE FPA=4;"
 Gdb.Execute " UPDATE TIM SET TRP='1.METΡ'"
 
 
510     If List2.ListCount > 0 Then
520         MsgBox "ΤΑ ΛΑΘΗ ΑΠΟΘΗΚΕΥΤΗΚΑΝ ΣΤΟ EXCEL  "
530         xlwbook.save  '  "c:\ektyp2.xls"
        End If

540     Set xlsheet = Nothing
550     Set xlwbook = Nothing
        ' excel.Quit

560     xl.Quit

570     MsgBox "Ενημερώθηκαν " + Format(mRow, "###0") + " εγγραφές"

        Exit Sub

WRITEERROR:
List1.AddItem "ΛΑΘΟΣ ΣΕ " + str(mRow)

580     Resume Next

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu51.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu51.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub



Function Letter2Number(xlsheet As excel.Worksheet, cColumnLetter As String) As Long
'PURPOSE: Convert a given letter into it's corresponding Numeric Reference
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

'Dim ColumnNumber As Long
'Dim ColumnLetter As String

'Input Column Letter
 ' ColumnLetter = "AG"
  
'Convert To Column Number
   Letter2Number = xlsheet.range(ColumnLetter & 1).Column
  
'Display Result
'  MsgBox "Column " & ColumnLetter & " = Column " & ColumnNumber
    
End Function







Function GET_CVALUE(query As String) As String
   Dim R As New ADODB.Recordset
   
   On Error Resume Next
   R.Open query, Gdb, adOpenDynamic, adLockOptimistic
   
    GET_CVALUE = ""
   If Not R.EOF Then
      GET_CVALUE = R(0)
   End If
   R.Close
   
End Function

Function GET_NVALUE(query As String) As Single
   Dim R As New ADODB.Recordset
   
   On Error Resume Next
   R.Open query, Gdb, adOpenDynamic, adLockOptimistic
   
    GET_NVALUE = 0
   If Not R.EOF Then
      GET_NVALUE = R(0)
   End If
   
   
   
   R.Close
   
End Function







Private Sub Label7_Click()

        On Error GoTo Label7_Click_Err

        Dim F As New DataEntry2, DD

        Dim R As New ADODB.Recordset

        'pinakas pedion
100     UPDATE2_PINAKES 103, Label7  ', Combo3

        Exit Sub

Label7_Click_Err:
       
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu51.Label7_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub TIMOKAT_Click()

        'TIMOKATALOGOI
        On Error GoTo timokat_Click_Err

        Dim R   As New ADODB.Recordset

        Dim r0  As New ADODB.Recordset

        Dim mok As Integer

        Dim PIN(30, 2)

        Dim COUNTER    As Integer

        Dim KOD_OK     As Boolean    ' ΕΧΕΙ ΔΗΛΩΣΕΙ ΤΟΝ ΚΩΔΙΚΟ

        Dim BARCODE_OK As Boolean    ' ΕΧΕΙ ΔΗΛΩΣΕΙ TA BARCODES

        Dim sql        As String

        Dim k          As Long

Dim PEDIA As String

PEDIA = "": sql = ""


120     COUNTER = 0  ' ARIUMOS PEDION POY THA METAFERTHOYN

        Dim KOD_COLUMN     As Integer: KOD_COLUMN = 1

        Dim BARCODE_COLUMN As Integer

130     r0.Open "SELECT TOP 1 * FROM TIMOKAT", Gdb, adOpenDynamic, adLockOptimistic

        ' FORTONO TON PINAKA 20 ME TIS ANTISTOIXISEIS TON PEDION
        ' ΚΑΙ ΕΛΕΓΧΩ ΑΝ ΕΧΕΙ ΔΗΛΩΘΕΙ Ο ΚΩΔΙΚΟΣ
140     R.Open "SELECT *FROM PINAKES WHERE TYPOS=103 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

150     Do While Not R.EOF
            COUNTER = COUNTER + 1
170         PIN(COUNTER, 1) = R("PERIGRAFH")    'ΠΕΔΙΟ ΤΟΥ ΚΥΡΙΟΥ ΑΡΧΕΙΟΥ (SQLSERVER)
180         PIN(COUNTER, 2) = R("TIMH")    ' ΑΡΙΘΜΟΣ ΣΤΗΛΗΣ ΤΟΥ EXCEL
            ' On Error GoTo LATOSPEDIO
220         mok = 0    'ΕΛΕΓΧΩ ΑΝ ΤΑ ΠΕΔΙΑ ΠΟΥ ΔΗΛΩΣΑ ΣΤΟ <<PINAKES>> ΥΠΑΡΧΟΥΝ ΣΤΟΝ ΠΙΝΑΚΑ PEL

230         For k = 0 To r0.FIELDS.Count - 1

240             If r0(k).Name = R("PERIGRAFH") Then
250                 mok = 1

                    Exit For

                End If

            Next

260         If mok = 0 Then
270             MsgBox "Λάθος στην δήλωση των πεδίων." + Chr(13) + "Δεν υπάρχει στο αρχείο ειδών το πεδίο " + R("PERIGRAFH")

                Exit Sub

            End If

            On Error GoTo timokat_Click_Err

290         R.MoveNext
        Loop

        ' kathgories

300     R.Close

330     Set xlwbook = xl.Workbooks.Open(Text1.Text)
340     Set xlsheet = xlwbook.Sheets.Item(1)

        Dim ko   As String

        Dim mNew As Long, mUpd As Long

350     mNew = 0
360     mUpd = 0

370     Label4.Caption = "Nέες εγγραφές 0"
380     Label5.Caption = "Ενημέρωση εγγραφές 0"

        Dim mRow As Long

        ' data1.Recordset.MoveFirst
390     mRow = Val(Text2.Text)    '  data1.Recordset.Move

        On Error GoTo error_name

400     Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF

410         If IsNull(xlsheet.cells(mRow, KOD_COLUMN)) Then

                Exit Do

            End If

420         If IsEmpty(xlsheet.cells(mRow, KOD_COLUMN)) Then

                Exit Do

            End If

430         ko = Trim(xlsheet.cells(mRow, KOD_COLUMN))    '  ' ΚΩΔΙΚΟς ΤΟΥ ΕΙΔΟΥΣ  EXCEL
440         R.Open "SELECT * FROM TIMOKAT ", Gdb, adOpenDynamic, adLockOptimistic
450         ' If r.EOF Then    'DEN BRHKA TON KODIKO


460       '  R.AddNew
            '470             r("KOD") = Left(ko, 14)
            '        R("fpa") = Val(Left(Combo2.Text, 2))
480         mNew = mNew + 1
490         Label4.Caption = "Nέες εγγραφές " + Format(mNew, "######")
            ' Else  'BRHKA TON KODIKO
500         '     mUpd = mUpd + 1
510         Label5.Caption = "Ενημέρωση εγγραφές " + Format(mUpd, "######")

            ' End If
            'ΕΝΗΜΕΡΩΝΩ ΤΑ ΥΠΟΛΟΙΠΑ ΠΕΔΙΑ ΑΠΟ ΤΟ EXCEL
520         For k = 1 To COUNTER

540             If R(PIN(k, 1)).Type = 202 Then    ' KEIMENO
550                ' R(PIN(k, 1)) = Left(xlsheet.cells(mRow, PIN(k, 2)), R(PIN(k, 1)).DefinedSize)
                   sql = sql + "'" + Left(xlsheet.cells(mRow, PIN(k, 2)), R(PIN(k, 1)).DefinedSize) + "'" + IIf(k = COUNTER, "", ",")
                Else    'ARITMOS
560                ' R(PIN(k, 1)) = xlsheet.cells(mRow, PIN(k, 2))
                   sql = sql + Replace(xlsheet.cells(mRow, PIN(k, 2)), ",", ".") + IIf(k = COUNTER, "", ",")
                End If
                
                PEDIA = PEDIA + Replace(PIN(k, 1), ",", ".") + IIf(k = COUNTER, "", ",")
               
            Next

            Gdb.Execute "INSERT INTO TIMOKAT (" + PEDIA + ") VALUES (" + sql + ")"
            PEDIA = "": sql = ""

580        ' R.Update

590         R.Close

600         DoEvents
610         Me.Caption = ko
620         mRow = mRow + 1    'data1.Recordset.MoveNext
        Loop

630     xl.Quit
640     Set xlwbook = Nothing
650     Set xl = Nothing
660     MsgBox "τέλος ενημέρωσης"

        Exit Sub

error_name:
670     mok = MsgBox("λάθος στην σειρά " + Format(mRow, "#######"), vbYesNo, "Διακοπή ;")

680     If mok = vbYes Then Exit Sub

690     Resume Next

        '<EhFooter>
        Exit Sub

timokat_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu51.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu51.Command1_Click " & "at line " & Erl

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

120     Combo2.Text = Combo2.List(1)

130     Combo1.Text = Combo1.List(0)


PEDIA


        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu51.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu51.Form_Load " & "at line " & Erl

        Resume Next
        

        '</EhFooter>

End Sub

Private Sub Form_Resize()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    mForm_Resize Me, 8, 13875, 9480, fh, fw, ft, fl

End Sub

Private Sub Label6_Click()

        '<EhHeader>
        On Error GoTo Label6_Click_Err

        '</EhHeader>
        Dim F As New DataEntry2, DD

        Dim R As New ADODB.Recordset

        'pinakas pedion
100     UPDATE2_PINAKES 21, Label6  ', Combo3

        ' UPDATE_PINAKES3f 21, Label6  ', Combo1
        '<EhFooter>
        Exit Sub

Label6_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.bohu51.Label6_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.bohu51.Label6_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

