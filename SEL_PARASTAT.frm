VERSION 5.00
Begin VB.Form uSEL_PARASTAT 
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   5850
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "SEL_PARASTAT.frx":0000
      Left            =   120
      List            =   "SEL_PARASTAT.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "dBASE IV;"
      DatabaseName    =   "D:\LAGEURO"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "PARASTAT"
      Top             =   3240
      Width           =   5295
   End
End
Attribute VB_Name = "uSEL_PARASTAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
 List1.Clear
  
  Data1.Recordset.MoveFirst
  Do While Not Data1.Recordset.EOF
     If Not IsNull(Data1.Recordset("TITLOS")) Then
        List1.AddItem Data1.Recordset("TITLOS")
     End If
     Data1.Recordset.MoveNext
  Loop
  
  
End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(gPicture)
End Sub
