VERSION 5.00
Begin VB.Form utDialog 
   BackColor       =   &H000080FF&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   4425
   ClientLeft      =   2835
   ClientTop       =   3375
   ClientWidth     =   3270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   4  'Dash-Dot-Dot
   FillStyle       =   3  'Vertical Line
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton toPDF 
      Caption         =   "Σε εγγραφο PDF"
      Height          =   372
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "https://www.biopdf.com/download.php"
      Top             =   3120
      Width           =   3132
   End
   Begin VB.CommandButton cmdApacheOffice 
      Caption         =   "Εγγραφο  ApacheOffice"
      Height          =   372
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CheckBox ALLAGHSEIRAS 
      Caption         =   "Αλλαγή σειράς σε νέο κωδικό"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   3135
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Εξοδος"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Height          =   1620
      ItemData        =   "ektypotes.frx":0000
      Left            =   3255
      List            =   "ektypotes.frx":0002
      TabIndex        =   4
      Top             =   195
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Εγγραφο  Word"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Excel  Φύλλο Εργασίας"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "Εκτυπωτής"
      DragIcon        =   "ektypotes.frx":0004
      Height          =   375
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Σημειωματάριο (ASCII)"
      Height          =   375
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   690
      Width           =   3135
   End
   Begin VB.CommandButton cmdCSV 
      Caption         =   "Εγγραφο CSV"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   3135
   End
End
Attribute VB_Name = "utDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCSV_Click()


     MDIForm1.StatusBar1.Panels(6).Text = 5
     Unload Me



End Sub

Private Sub cmdApacheOffice_Click()
   ' On Error GoTo Command1_Click_Err

        '</EhHeader>

     MDIForm1.StatusBar1.Panels(6).Text = 8
     Unload Me



End Sub

'Private Sub OKButton_Click()
'
'End Sub

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

100     MDIForm1.StatusBar1.Panels(6).Text = 2
110     Unload Me

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.utDialog.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.utDialog.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>
        Dim x    As Printer

        Dim DEFA As String

100     DEFA = Printer.DeviceName

        ' List1.Enabled = False
        Dim L As Integer

        Dim m As Integer

110     L = 0

120     For Each x In Printers

130         List1.AddItem x.DeviceName

140         If x.DeviceName = DEFA Then
150             m = L

            End If

160         L = L + 1
        Next

170     List1.Visible = True
180     List1.SetFocus
190     List1.Left = Command2.Left
200     List1.Selected(m) = True

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.utDialog.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.utDialog.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        '<EhHeader>
        On Error GoTo Command3_Click_Err

        '</EhHeader>
100     If ALLAGHSEIRAS.Value = vbChecked Then
110         MDIForm1.StatusBar1.Panels(6).Text = 13
        Else
120         MDIForm1.StatusBar1.Panels(6).Text = 3
        End If

130     Unload Me

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.utDialog.Command3_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.utDialog.Command3_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command4_Click()

        '<EhHeader>
        On Error GoTo Command4_Click_Err

        '</EhHeader>

100     MDIForm1.StatusBar1.Panels(6).Text = 4
110     Unload Me

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.utDialog.Command4_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.utDialog.Command4_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command5_Click()

        '<EhHeader>
        On Error GoTo Command5_Click_Err

        '</EhHeader>

100     MDIForm1.StatusBar1.Panels(6).Text = 0
110     Unload Me

        '<EhFooter>
        Exit Sub

Command5_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.utDialog.Command5_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.utDialog.Command5_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub



Private Sub Form_Click()

        ' Dim seir_selid1
        ' seires_selidas1 = Val(FindParametroi(1,"ektypotes", "SEIR_SELID1", "70", "Σειρές ανά σελίδα(κάθετη)"))
        '<EhHeader>
        On Error GoTo Form_Click_Err

        '</EhHeader>

100     PARAMETROI.PARAM.Caption = "EKTYPOTES"
110     PARAMETROI.SHOW 1

        '<EhFooter>
        Exit Sub

Form_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.utDialog.Form_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.utDialog.Form_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo List1_KeyDown_Err

        '</EhHeader>
100     If KeyCode = 13 Then SELECT_PRINTER

        '<EhFooter>
        Exit Sub

List1_KeyDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.utDialog.List1_KeyDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.utDialog.List1_KeyDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

        '<EhHeader>
        On Error GoTo List1_MouseDown_Err

        '</EhHeader>

100     SELECT_PRINTER

        '<EhFooter>
        Exit Sub

List1_MouseDown_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.utDialog.List1_MouseDown " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.utDialog.List1_MouseDown " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Sub SELECT_PRINTER()

        '<EhHeader>
        On Error GoTo SELECT_PRINTER_Err

        '</EhHeader>

        Dim PX As Printer

100     For Each PX In Printers

            ' List1.AddItem x.DeviceName

110         If InStr(PX.DeviceName, List1.Text) > 0 Then
                ' Set printer as system default.
120             Set Printer = PX
                ' Stop looking for a printer  5200016000239

                Exit For

            End If

        Next

130     If ALLAGHSEIRAS.Value = vbChecked Then
140         MDIForm1.StatusBar1.Panels(6).Text = 11
        Else
150         MDIForm1.StatusBar1.Panels(6).Text = 1
        End If

160     Unload Me

        '<EhFooter>
        Exit Sub

SELECT_PRINTER_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.utDialog.SELECT_PRINTER " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.utDialog.SELECT_PRINTER " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub toPDF_Click()


     MDIForm1.StatusBar1.Panels(6).Text = 12
     Unload Me

        '<EhFooter>
        Exit Sub



End Sub
