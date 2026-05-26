VERSION 5.00
Begin VB.Form MHNYMA2 
   BackColor       =   &H00FFFF80&
   ClientHeight    =   984
   ClientLeft      =   9756
   ClientTop       =   9432
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   984
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "ΕΞΟΔΟΣ"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   4380
      Top             =   -30
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Η εγγραφή ΡΟΥΛΜΑΝ 6204 αποθηκεύτηκε με επιτυχία"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.4
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   4545
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Αποθήκευση εγγραφής"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   3135
   End
End
Attribute VB_Name = "MHNYMA2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

100     Unload Me

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.MHNYMA2.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.MHNYMA2.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()
    ' Timer1.Interval = 1000
End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    Unload Me
End Sub
