VERSION 5.00
Begin VB.Form pelat21 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3240
   ScaleWidth      =   9345
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "pelat21b.frx":0000
      Top             =   2640
      Width           =   9135
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "pelat21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VAR1

Private Sub Command5_Click()
Dim R As Recordset
Set R = Data1.Recordset
R.AddNew
R(0) = Label1 + Space(10) + Format(R.RecordCount + 1, "00")
R.Update


End Sub

Private Sub DBGrid2_AfterInsert()
  'On Error Resume Next
  
End Sub


Private Sub DBGrid2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'Data1.Recordset.Edit
'  Data1.Recordset(0) = Label1 + Space(10) + Format(Data1.Recordset.RecordCount + 1, "00")
'  Data1.Recordset.Update
End Sub

Private Sub Form_Load()
VAR1 = 1
End Sub

Private Sub Form_Paint()
Dim s
If VAR1 = 0 Then Exit Sub
VAR1 = 1

' DbGrid2.width = 6500
' DbGrid2.Visible = True

Data1.DatabaseName = gDir
Data1.Connect = gConnect
s = left(Label1.Caption + Space(15), 15)
Data1.RecordSource = "select  *from memotext where left(memo_key,15)='" + s + "' order by memo_key;"




On Error Resume Next
Data1.Refresh




DbGrid2.Columns(0).width = 20
DbGrid2.Columns(1).width = 6400

End Sub

Private Sub Form_Unload(Cancel As Integer)


Dim K, s As String


If Data1.Recordset.RecordCount = 0 Then Exit Sub

Data1.Recordset.MoveFirst
For K = 1 To Data1.Recordset.RecordCount
  Data1.Recordset.Edit
  s = Label1 + Space(18)
  Data1.Recordset(0) = left(s, 15) + Format(K, "000")
  Data1.Recordset.Update
  Data1.Recordset.MoveNext
  
 Next

   
   
End Sub

