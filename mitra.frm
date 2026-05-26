VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   8880
   Begin VB.CommandButton m_Exit 
      Caption         =   "Εξοδος"
      Height          =   675
      Left            =   6915
      TabIndex        =   0
      Top             =   6750
      Width           =   1605
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then m_Exit_Click
  
End Sub

Private Sub Form_Load()
Me.KeyPreview = True

 Me.Picture = LoadPicture(gPicture)
End Sub

Private Sub m_Exit_Click()
   Unload Me
End Sub
