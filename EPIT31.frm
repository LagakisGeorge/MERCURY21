VERSION 5.00
Begin VB.Form epit3100 
   Caption         =   "Form1"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   3480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1770
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Καταχώρηση"
      Height          =   435
      Left            =   300
      TabIndex        =   3
      Top             =   960
      Width           =   1350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Εξοδος"
      Height          =   435
      Left            =   1890
      TabIndex        =   2
      Top             =   960
      Width           =   1350
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "EPIT31.frx":0000
      Left            =   300
      List            =   "EPIT31.frx":0002
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   540
      Width           =   2955
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Τράπεζα"
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
      Left            =   300
      TabIndex        =   1
      Top             =   180
      Width           =   1575
   End
End
Attribute VB_Name = "epit3100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

100     Unload Me

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.epit3100.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.epit3100.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        '</EhHeader>
        Dim m As Integer

100     m = Val(Left(Combo1.Text, 2))

110     epit3.Caption = Combo1.Text    ' m
120     Unload Me

        ' epit3.f = 1

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.epit3100.Command2_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.epit3100.Command2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
        Dim R As New ADODB.Recordset

        'ΠΙΝΑΚΙΑ
100     R.Open "SELECT *FROM PINAKES WHERE TYPOS=18 ORDER BY AYJON", Gdb, adOpenDynamic, adLockOptimistic

110     Do While Not R.EOF

120         If R("typos") = 18 Then
130             Combo1.AddItem Format(R("AYJON"), "#0") + " " + R("PERIGRAFH") + Right(Space(30) + Format(R("TIMH"), "0000"), 30)
            End If

140         R.MoveNext
        Loop

150     Combo1.Text = Combo1.List(0)

160     R.Close

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.epit3100.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.epit3100.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
