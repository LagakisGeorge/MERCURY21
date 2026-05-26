VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   2040
      TabIndex        =   0
      Top             =   2160
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim GDB As New ADODB.Connection
GDB.Open "MERCSQL"


Dim SUMAEID As New ADODB.Recordset





SUMAEID.Open "SELECT * FROM EIDFOTOS", GDB, adOpenDynamic, adLockBatchOptimistic

Do While Not SUMAEID.EOF
If Len(Dir("C:\MERCVB\IMAGES\" + SUMAEID("KOD") + ".JPG")) > 0 Then
   FileCopy "C:\MERCVB\IMAGES\" + SUMAEID("KOD") + ".JPG", "C:\DEL\" + SUMAEID("KOD") + ".JPG"
   If Len(Dir("C:\DEL\" + SUMAEID("KOD") + ".JPG")) > 0 Then
      Kill "C:\MERCVB\IMAGES\" + SUMAEID("KOD") + ".JPG"
      DoEvents
      Me.Caption = SUMAEID("KOD")
      
   End If
End If
SUMAEID.MoveNext

Loop

SUMAEID.Close





End Sub
