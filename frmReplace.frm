VERSION 5.00
Begin VB.Form frmReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Replace"
   ClientHeight    =   1080
   ClientLeft      =   5136
   ClientTop       =   7812
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4710
      TabIndex        =   5
      Top             =   510
      Width           =   1635
   End
   Begin VB.CommandButton btnReplace 
      Caption         =   "Replace"
      Default         =   -1  'True
      Height          =   315
      Left            =   4710
      TabIndex        =   4
      Top             =   210
      Width           =   1635
   End
   Begin VB.TextBox txtReplace 
      Height          =   285
      Left            =   1710
      TabIndex        =   3
      Top             =   540
      Width           =   2955
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   1710
      TabIndex        =   2
      Top             =   240
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "Replace With:"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   630
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "Find What:"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   1305
   End
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event MyCallBack(ByVal fText As String, ByVal rText As String, bCancel As Boolean)

Friend Sub btnReplace_Click()
    RaiseEvent MyCallBack(txtFind.Text, txtReplace.Text, False)
    Unload Me
End Sub

Private Sub Command1_Click()
    RaiseEvent MyCallBack(txtFind.Text, txtReplace.Text, True)
    Unload Me
End Sub

