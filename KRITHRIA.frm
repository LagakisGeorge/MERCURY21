VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form KRITHRIA 
   BackColor       =   &H00FF0000&
   Caption         =   "Form2"
   ClientHeight    =   7065
   ClientLeft      =   1440
   ClientTop       =   3720
   ClientWidth     =   7740
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   7740
   Begin VB.ComboBox kathgpel 
      Height          =   315
      Left            =   4500
      TabIndex        =   8
      Top             =   4950
      Width           =   1680
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2025
      TabIndex        =   5
      Top             =   4935
      Width           =   1935
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   30
      TabIndex        =   4
      Top             =   4935
      Width           =   1785
   End
   Begin VB.CommandButton ΕΧΙΤ 
      Caption         =   "ΕΞΟΔΟΣ"
      Height          =   630
      Left            =   6270
      TabIndex        =   2
      Top             =   3480
      Width           =   780
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   -360
      TabIndex        =   0
      Top             =   4320
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   4020
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7091
      _Version        =   393216
      Rows            =   60
      Cols            =   5
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Κατηγ.Πελάτη"
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
      Height          =   210
      Index           =   1
      Left            =   4485
      TabIndex        =   9
      Top             =   4710
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Υποοικογένεια"
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
      Height          =   225
      Index           =   4
      Left            =   2010
      TabIndex        =   7
      Top             =   4710
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Οικογένεια"
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
      Height          =   330
      Index           =   25
      Left            =   60
      TabIndex        =   6
      Top             =   4710
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "$ ασυμπλήρωτο (null)   #συμπληρωμένο"
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   4320
      Width           =   4935
   End
End
Attribute VB_Name = "KRITHRIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Fty(50)    As String, fField(50) As String, PROTH_FORA As Boolean

Dim f_FreeToGo As Boolean

Private Sub Combo3_Change()

        '<EhHeader>
        On Error GoTo Combo3_Change_Err

        '</EhHeader>

100     Text3.Text = Left(Combo3.Text, 2)
110     Text3.SetFocus

        '<EhFooter>
        Exit Sub

Combo3_Change_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.KRITHRIA.Combo3_Change " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
     '  'SAVE_ERROR Err.Description & " in ADOMERCNEW.KRITHRIA.Combo3_Change " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Combo4_LostFocus()

        '<EhHeader>
        On Error GoTo Combo4_LostFocus_Err

        '</EhHeader>
        On Error Resume Next

        ' YPOOIKOGENEIES
        Dim R As New ADODB.Recordset

100     If Len(Combo4.Text) > 0 Then
110         R.Open "SELECT *FROM PINAKES WHERE TYPOS=3 AND TIMH=" + Left(Combo4.Text, 2) + "ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic
120         Combo3.Clear

130         Do While Not R.EOF

140             If R("typos") = 3 Then
150                 Combo3.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
                Else

                    Exit Do

                End If

160             R.MoveNext
            Loop

170         R.Close

        End If

180     Text3.Text = Left(Combo4.Text, 2)
190     Text3.SetFocus

        '<EhFooter>
        Exit Sub

Combo4_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.KRITHRIA.Combo4_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
       'SAVE_ERROR Err.Description & " in ADOMERCNEW.KRITHRIA.Combo4_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     Grid1.ColWidth(0) = 2500
110     f_FreeToGo = True

120     PROTH_FORA = False

        ' OIKOGENEIES

        Dim R As New ADODB.Recordset

130     R.Open "SELECT *FROM PINAKES WHERE TYPOS=11 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

        'R.MoveFirst
140     Do While Not R.EOF

150         If R("typos") = 11 Then
160             Combo4.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
            Else

                Exit Do

            End If

170         R.MoveNext
        Loop

180     R.Close

        ' kathgpel
190     R.Open "SELECT *FROM PINAKES WHERE TYPOS=13 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

        'R.MoveFirst
200     Do While Not R.EOF

210         If R("typos") = 13 Then
220             kathgpel.AddItem Format(R("ayjon"), "00") + " " + Left(R("PERIGRAFH") + Space(20), 20)
            Else

                Exit Do

            End If

230         R.MoveNext
        Loop

240     R.Close

        ' Grid1.ColIsVisible(4) = False
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.KRITHRIA.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
       'SAVE_ERROR Err.Description & " in ADOMERCNEW.KRITHRIA.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Paint()

        '<EhHeader>
        On Error GoTo Form_Paint_Err

        '</EhHeader>

100     Dim ctyp As Integer: ctyp = 4

110     Dim cf   As Integer: cf = 3

        Dim k    As Integer

120     If PROTH_FORA = True Then Exit Sub

130     PROTH_FORA = True

        Dim SEIRES

140     SEIRES = 0
150     Grid1.TextMatrix(0, 1) = "Από"
160     Grid1.TextMatrix(0, 2) = "Εως"

        '--------------------------------------------------------------------------------------
        'ΚΟΒΕΙ ΤΙΣ ΑΔΕΙΕΣ ΣΕΙΡΕΣ
170     For k = 1 To Grid1.rows - 1

180         With Grid1

190             If Len(.TextMatrix(k, cf)) > 0 Then
200                 Fty(k) = .TextMatrix(k, ctyp)
210                 fField(k) = .TextMatrix(k, cf)
220                 SEIRES = SEIRES + 1
                End If

            End With

        Next

230     Grid1.ColS = 3
240     Grid1.rows = SEIRES + 1
        
250     Grid1_EnterCell

        '<EhFooter>
        Exit Sub

Form_Paint_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.KRITHRIA.Form_Paint " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
       'SAVE_ERROR Err.Description & " in ADOMERCNEW.KRITHRIA.Form_Paint " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub load_arrays()


End Sub



Private Sub Form_Unload(Cancel As Integer)

        '<EhHeader>
        On Error GoTo Form_Unload_Err

        '</EhHeader>
        Dim k    As Integer

        Dim sql  As String

100     Dim ctyp As Integer: ctyp = 4

110     Dim cf   As Integer: cf = 3

120     Dim ct1  As Integer: ct1 = 1

130     Dim ct2  As Integer: ct2 = 2

        On Error Resume Next

140     sql = ""

        Dim sql1 As String

        Dim MCC1, MCC2

       ' load_arrays


150     f_FreeToGo = True

160     Grid1.Text = Text3.Text  ' κρατάει το grid το κείμενο

170     For k = 1 To Grid1.rows - 1
180         sql1 = ""

190         With Grid1

200             If Len(.TextMatrix(k, ct1)) > 0 Then

210                 MCC1 = .TextMatrix(k, ct1)
220                 MCC2 = .TextMatrix(k, ct2)

230                 MCC2 = Replace(MCC2, ",", ".")
240                 MCC1 = Replace(MCC1, ",", ".")

250                 If UCase(Fty(k)) = "C" Then
260                     If MCC1 = "$" Or MCC2 = "$" Then
270                         sql1 = fField(k) + " IS NULL "
280                     ElseIf MCC1 = "#" Or MCC2 = "#" Then
290                         sql1 = fField(k) + " IS NOT NULL "
                        Else

300                         If Len(.TextMatrix(k, ct2)) > 0 Then
310                             sql1 = fField(k) + ">='" + LTrim(MCC1) + "' and " + fField(k) + "<='" + LTrim(MCC2) + "'"
                            Else
320                             sql1 = fField(k) + " LIKE '" + SameLetters(LTrim(MCC1)) + "%'"
                            End If
                        End If
                    End If

330                 If UCase(Fty(k)) = "N" Then
340                     If MCC1 = "$" Or MCC2 = "$" Then
350                         sql1 = fField(k) + " IS NULL "
360                     ElseIf MCC1 = "#" Or MCC2 = "#" Then
370                         sql1 = fField(k) + " IS NOT NULL "
                        Else

380                         If Val(MCC2) < Val(MCC1) Then
390                             MsgBox "Ο αριθμός στη δεξιά στήλη " + Chr(13) + "πρέπει να είναι μεγαλύτερος από τον αριθμό " + Chr(13) + "της αριστερής στήλης"
400                             f_FreeToGo = False
410                             Grid1.row = k
420                             Grid1.Col = 2
430                             Grid1_EnterCell

                                Exit Sub

                            Else
440                             sql1 = fField(k) + ">=" + MCC1 + " and " + fField(k) + "<=" + MCC2
                            End If

                        End If
                    End If

450                 If UCase(Fty(k)) = "D" Then

460                     If MCC1 = "$" Or MCC2 = "$" Then
470                         sql1 = fField(k) + " IS NULL "
480                     ElseIf MCC1 = "#" Or MCC2 = "#" Then
490                         sql1 = fField(k) + " IS NOT NULL "
                        Else

500                         sql1 = fField(k) + ">='" + Format(MCC1, "mm/dd/yyyy") + "' and " + fField(k) + "<'" + Format(DateAdd("d", 1, MCC2), "mm/dd/yyyy") + "'"
                        End If

                    End If

                End If

            End With

            ' προσθετω το and
510         If Len(sql1) > 0 Then
520             If Len(sql) > 0 Then
530                 sql = sql + " and " + sql1
                Else
540                 sql = sql1
                End If
            End If

        Next

        'Dim xx
        'xx = InputBox("επιπλέον κριτήριο ")
        'If Len(xx) > 2 Then
        '   SQL = SQL + " and " + xx
        'End If

550     MDIForm1.StatusBar1.Panels(6).Text = sql

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.KRITHRIA.Form_Unload " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
       'SAVE_ERROR Err.Description & " in ADOMERCNEW.KRITHRIA.Form_Unload " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Grid1_Click()

        '<EhHeader>
        On Error GoTo Grid1_Click_Err

        '</EhHeader>

100     Grid1_EnterCell

        '<EhFooter>
        Exit Sub

Grid1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.KRITHRIA.Grid1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
       'SAVE_ERROR Err.Description & " in ADOMERCNEW.KRITHRIA.Grid1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Grid1_EnterCell()

        '<EhHeader>
        On Error GoTo Grid1_EnterCell_Err

        '</EhHeader>

100     Text3.Top = Grid1.Top + Grid1.CellTop    '- 10
110     Text3.Left = Grid1.Left + Grid1.CellLeft    ' - 10
120     Text3.Width = Grid1.CellWidth - 35
130     Text3.Height = Grid1.CellHeight - 80
140     Text3.Text = Grid1.Text

150     Text3.SetFocus

        '<EhFooter>
        Exit Sub

Grid1_EnterCell_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.KRITHRIA.Grid1_EnterCell " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
       'SAVE_ERROR Err.Description & " in ADOMERCNEW.KRITHRIA.Grid1_EnterCell " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Grid1_LeaveCell()

        '<EhHeader>
        On Error GoTo Grid1_LeaveCell_Err

        '</EhHeader>

100     Grid1.Text = Text3.Text  ' κρατάει το grid το κείμενο

        '<EhFooter>
        Exit Sub

Grid1_LeaveCell_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.KRITHRIA.Grid1_LeaveCell " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
       'SAVE_ERROR Err.Description & " in ADOMERCNEW.KRITHRIA.Grid1_LeaveCell " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub kathgpel_Click()

        '<EhHeader>
        On Error GoTo kathgpel_Click_Err

        '</EhHeader>

100     Text3.Text = Left(kathgpel.Text, 2)
        ' Text3.SetFocus
110     Grid1.Text = Text3.Text  ' κρατάει το grid το κείμενο

        '<EhFooter>
        Exit Sub

kathgpel_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.KRITHRIA.kathgpel_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
       'SAVE_ERROR Err.Description & " in ADOMERCNEW.KRITHRIA.kathgpel_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub kathgpel_LostFocus()
    ' Text3.Text = Left(kathgpel.Text, 2)
    '  ' Text3.SetFocus
    '   Grid1.Text = Text3.Text  ' κρατάει το grid το κείμενο

End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo Text3_KeyUp_Err

        '</EhHeader>

100     If KeyCode = 13 Then
110         Grid1.Text = Text3.Text  ' κρατάει το grid το κείμενο

120         If Grid1.Col < Grid1.ColS - 1 Then  ' πάει στην επόμενη κολόνα
130             Grid1.Col = Grid1.Col + 1
140             Me.Caption = "125"
            Else

                'πάει στην παρακάτω σειρά
150             If Grid1.row < Grid1.rows - 1 And Grid1.row < Grid1.rows - 1 Then
160                 Grid1.row = Grid1.row + 1
170                 Grid1.Col = 1
                Else    ' τερματισε κατω δεξιά

180                 Form_Unload 0

190                 If f_FreeToGo = True Then
200                     KRITHRIA.Hide
                    End If

                    'Unload Me
                    Exit Sub

                End If
            End If

210         Grid1_EnterCell
        End If

220     If KeyCode = 38 Then   ' ΠΑΝΩ ΒΕΛΟΣ

230         Grid1.Text = Text3.Text    ' κρατάει το grid το κείμενο

            ' πάει μία κολόνα αριστερά
240         If Grid1.Col > 1 Then
250             Grid1.Col = Grid1.Col - 1
            Else

                'π'αει μία σειρά επάνω
260             If Grid1.row > 1 Then
270                 Grid1.row = Grid1.row - 1
280                 Grid1.Col = 1
                End If
            End If

290         Grid1_EnterCell
        End If

300     If KeyCode = 40 Then   ' KATΩ ΒΕΛΟΣ
310         Grid1.Text = Text3.Text    ' κρατάει το grid το κείμενο

            ' μιά σειρά παρακάτω
320         If Grid1.row < Grid1.rows - 1 Then
330             Grid1.row = Grid1.row + 1
340             Grid1_EnterCell
            End If

        End If

350     If KeyCode = 27 Then    'esc τελος εισαγωγής
360         KeyCode = 0

370         Form_Unload 0
380         KRITHRIA.Hide
            'Unload Me
        End If

        '<EhFooter>
        Exit Sub

Text3_KeyUp_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.KRITHRIA.Text3_KeyUp " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
       'SAVE_ERROR Err.Description & " in ADOMERCNEW.KRITHRIA.Text3_KeyUp " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Text3_KeyPress_Err

        '</EhHeader>
        Dim DUM

100     If KeyAscii = 13 Then
110         KeyAscii = 0
120         DUM = 0
        End If

130     If KeyAscii = 27 Then
140         KeyAscii = 0
150         DUM = 0
        End If

        '<EhFooter>
        Exit Sub

Text3_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.KRITHRIA.Text3_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
       'SAVE_ERROR Err.Description & " in ADOMERCNEW.KRITHRIA.Text3_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text3_LostFocus()

        '<EhHeader>
        On Error GoTo Text3_LostFocus_Err

        '</EhHeader>

100     Grid1.Text = Text3.Text  ' κρατάει το grid το κείμενο

        '<EhFooter>
        Exit Sub

Text3_LostFocus_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.KRITHRIA.Text3_LostFocus " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
       'SAVE_ERROR Err.Description & " in ADOMERCNEW.KRITHRIA.Text3_LostFocus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub ΕΧΙΤ_Click()

        '<EhHeader>
        On Error GoTo ΕΧΙΤ_Click_Err

        '</EhHeader>

100     Form_Unload 0

110     If f_FreeToGo Then
120         Me.Hide
        End If

        '<EhFooter>
        Exit Sub

ΕΧΙΤ_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.KRITHRIA.ΕΧΙΤ_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
       'SAVE_ERROR Err.Description & " in ADOMERCNEW.KRITHRIA.ΕΧΙΤ_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub


