VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   165
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   8910
   WindowState     =   2  'Maximized
   Begin VB.Menu pelat 
      Caption         =   "Πελάτες/Προμηθευτές"
      Begin VB.Menu mpelat1 
         Caption         =   "Εισαγωγή νέου"
      End
      Begin VB.Menu mpelat2 
         Caption         =   "Διόρθωση/Διαγραφή"
      End
      Begin VB.Menu mpelat3 
         Caption         =   "Παρουσίαση"
      End
      Begin VB.Menu mpelat4 
         Caption         =   "Καρτέλλα"
      End
      Begin VB.Menu exit 
         Caption         =   "Εξοδος"
      End
   End
   Begin VB.Menu apot 
      Caption         =   "Αποθήκη"
   End
   Begin VB.Menu epit 
      Caption         =   "Επιταγές"
   End
   Begin VB.Menu par 
      Caption         =   "Παραστατικά"
   End
   Begin VB.Menu bohu 
      Caption         =   "Βοηθητικά"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub bohu_Click()

End Sub

Private Sub exit_Click()
  Unload Me
End Sub

Private Sub Form_Load()
   gDir = "d:\lageuro"
End Sub

Private Sub mpelat1_Click()
pelat1.Show
End Sub

Private Sub mpelat2_Click()
pelat2.Show
End Sub

Private Sub mpelat4_Click()
pelat4.Show
End Sub

Private Sub par_Click()

End Sub
