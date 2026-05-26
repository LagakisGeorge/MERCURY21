VERSION 5.00
Begin VB.Form bohu61 
   Caption         =   "Form2"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6750
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      TabIndex        =   10
      Top             =   4260
      Width           =   9345
   End
   Begin VB.CommandButton save 
      Caption         =   " ¡‘¡◊Ÿ—«”«"
      Height          =   360
      Left            =   2550
      TabIndex        =   8
      Top             =   3480
      Width           =   2880
   End
   Begin VB.TextBox lti5 
      Height          =   405
      Left            =   2430
      TabIndex        =   3
      Top             =   2220
      Width           =   3015
   End
   Begin VB.TextBox FPA 
      Height          =   405
      Left            =   2400
      TabIndex        =   2
      Text            =   "2"
      Top             =   1500
      Width           =   3015
   End
   Begin VB.TextBox ono 
      Height          =   405
      Left            =   2400
      TabIndex        =   1
      Top             =   840
      Width           =   5145
   End
   Begin VB.TextBox kod 
      Height          =   405
      Left            =   2400
      TabIndex        =   0
      Top             =   180
      Width           =   3015
   End
   Begin VB.Label NEA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      TabIndex        =   9
      Top             =   210
      Width           =   1455
   End
   Begin VB.Label lbl‘…Ã« 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘…Ã«"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   1425
   End
   Begin VB.Label lbl÷–¡1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "÷–¡ 1=13% 2=24% 4=6%"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   1620
      Width           =   1425
   End
   Begin VB.Label lbl–≈—…√—¡÷« 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "–≈—…√—¡÷«"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   990
      Width           =   1425
   End
   Begin VB.Label lblBarcode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "barcode"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   300
      Width           =   1425
   End
End
Attribute VB_Name = "bohu61"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fnew As Integer
Dim FID As Long


Private Sub save_Click()
Dim K As Integer

If fnew = 0 Then
' ƒ…œ—»Ÿ”« ‘…Ã«Ú
  Gdb.Execute "UPDATE EID SET LTI5=" + Replace(lti5.Text, ",", ".") + " WHERE ID=" + str(FID), K
     If K > 0 Then
       kod.Text = ""
       kod.SetFocus
     End If

     List1.AddItem lti5.Text + "--" + kod.Text, 0
     
     



Else
   
   Dim R As New ADODB.Recordset
    
    R.Open "SELECT * FROM EID WHERE KOD='" + kod.Text + "'", Gdb, adOpenDynamic, adLockOptimistic
   If R.EOF Then
   
   
         Gdb.Execute "INSERT INTO EID (KOD,ONO,FPA,LTI5) VALUES ('" + kod.Text + "','" + ono.Text + "'," + str(FPA.Text) + "," + Replace(lti5.Text, ",", ".") + ")", K
         Gdb.Execute "INSERT INTO BARCODES (KOD,ERG) VALUES ('" + kod.Text + "','" + kod.Text + "')"
        
   End If
        
        
        
    If K > 0 Then
       List1.AddItem kod.Text + "+++" + lti5.Text + "++" + ono.Text, 0
       kod.Text = ""
       kod.SetFocus
     End If


End If








End Sub

Private Sub FPA_GotFocus()
   FPA.BackColor = vbYellow
End Sub

Private Sub FPA_LostFocus()
    FPA.BackColor = vbWhite
End Sub

Private Sub kod_GotFocus()
    kod.BackColor = vbYellow
    
End Sub

Private Sub kod_KeyUp(KeyCode As Integer, Shift As Integer)
Dim R As New ADODB.Recordset
Dim MKOD As String
   If KeyCode = 13 Then
       If Len(kod.Text) = 0 Then
          'kod.SetFocus
          Exit Sub
       End If
       
       
       R.Open "SELECT * FROM BARCODES WHERE ERG='" + kod.Text + "'", Gdb, adOpenDynamic, adLockOptimistic
       If R.EOF Then
          fnew = 1
          NEA.Caption = "NEO EIƒœ”"
          ono.SetFocus
       Else
         MKOD = R("kod")
         R.Close
         
         R.Open "SELECT * FROM EID WHERE KOD='" + MKOD + "'", Gdb, adOpenDynamic, adLockOptimistic
         FID = R("ID")
         ono.Text = R("ONO")
         FPA.Text = R("FPA")
         lti5.Text = Format(R("lti5"), "###0.00")
         NEA.Caption = "ƒ…œ—»Ÿ”« ‘…Ã«”"
         fnew = 0
         lti5.SetFocus
       End If
       
       
   End If
   
   
   
   




End Sub

Private Sub kod_LostFocus()
    kod.BackColor = vbWhite
End Sub

Private Sub lti5_GotFocus()
  lti5.BackColor = vbYellow
End Sub

Private Sub lti5_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
   save.SetFocus

 End If
End Sub

Private Sub lti5_LostFocus()
   lti5.BackColor = vbWhite
End Sub

Private Sub ono_GotFocus()
   ono.BackColor = vbYellow
End Sub

Private Sub ono_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
   FPA.SetFocus

 End If

End Sub

Private Sub FPA_KeyUp(KeyCode As Integer, Shift As Integer)

 If KeyCode = 13 Then
   lti5.SetFocus
 End If


End Sub

Private Sub ono_LostFocus()
   ono.BackColor = vbWhite
End Sub
