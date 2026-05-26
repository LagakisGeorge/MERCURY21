VERSION 5.00
Begin VB.Form bohu62 
   Caption         =   "Form1"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   255
      TabIndex        =   9
      Top             =   3615
      Width           =   2325
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Καταχώρηση Φωτογραφίας"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   0
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   5640
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   2760
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   2040
      Left            =   5640
      ScaleHeight     =   1980
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   840
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "BARCODE"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "bohu62"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Kill "C:\MERCVB\IMAGES\" + Label2.Caption
On Error GoTo 0


         Name File1.Path + "\" + File1.FILENAME As "C:\MERCVB\IMAGES\" + Label2.Caption
Me.Caption = "Κατεχωρήθη το " + Label3.Caption
Text1.Text = ""
Label2.Caption = ""
Label3.Caption = ""


End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
   Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
   Picture2.Picture = LoadPicture(File1.Path + "\" + File1.FILENAME)
   
   Picture1.PaintPicture Picture2.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0  '  , Picture1.ScaleWidth, Picture1.ScaleHeight ', vbSrcCopy
   Text1.SetFocus
   
End Sub

Private Sub Text1_LostFocus()
   Dim R As New ADODB.Recordset
   On Error Resume Next
   
   R.Open "select KOD FROM BARCODES WHERE ERG='" + Text1.Text + "'", Gdb, adOpenForwardOnly, adLockReadOnly
   Label2.Caption = R(0)
   R.Close
   
  R.Open "select KOD,ONO FROM EID WHERE KOD='" + Label2.Caption + "'", Gdb, adOpenForwardOnly, adLockReadOnly
  Label3.Caption = R(1)
   R.Close
   
End Sub
