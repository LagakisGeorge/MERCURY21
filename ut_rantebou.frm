VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ut_rantebou 
   Caption         =   "Form1"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Καταχώρηση"
      Height          =   450
      Left            =   1650
      TabIndex        =   1
      Top             =   2175
      Width           =   2160
   End
   Begin VB.TextBox Text1 
      Height          =   885
      Left            =   405
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   645
      Width           =   5055
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   1560
      TabIndex        =   3
      Top             =   135
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16711680
      CustomFormat    =   "dd-M-yyyy  HH:mm"
      Format          =   225378307
      CurrentDate     =   37771
   End
   Begin VB.Label Label1 
      Height          =   360
      Left            =   165
      TabIndex        =   2
      Top             =   1575
      Width           =   5475
   End
End
Attribute VB_Name = "ut_rantebou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>
        Dim DUM, date1 As Date, Aitiologia As String, Topos As String

100     Topos = Label1.Caption

110     date1 = DTPicker1.Value
120     Aitiologia = Text1.Text

130     DUM = add_schedule(date1, Aitiologia, Topos)
140     Unload Me

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.ut_rantebou.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.ut_rantebou.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     DTPicker1.Value = Now

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.ut_rantebou.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.ut_rantebou.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
